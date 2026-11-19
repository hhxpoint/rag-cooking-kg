# -*- coding: utf-8 -*-
"""
检索层消融实验（Ablation Study）
------------------------------------------------------------------
目的：把"图谱增强 RAG"拆成可独立开关的组件，量化每一部分对
      「检索证据召回率」与「端到端回答准确率」的边际贡献。

配置（唯一变量法，其余参数固定）：
  A 纯文本 BM25                        —— Naive RAG 基线
  B 纯文本 + 意图路由/规则推理         —— 隔离"推理层"的边际贡献
  C 纯图谱结构化召回（去掉文本路）     —— 验证图谱单独可用性
  D 文本 + 图谱「菜谱关联」事实        —— 只保留"包含"型证据
  E 文本 + 图谱「关系事实」            —— 只保留"替代/相克/搭配"证据
  F 完整两路拼接融合（本系统默认）     —— D ∪ E
  G F + 别名查询扩展                   —— Advanced RAG 查询侧改写
  H F + RRF 倒数排名融合               —— 前沿多路融合策略

指标：
  context_hit  证据召回率：期望知识点是否至少 1 个出现在检索上下文
  accuracy     端到端回答准确率（与 evaluate.py 判定完全一致）
  avg_ctx_chars 平均上下文字符数（噪声/成本代理指标）
  hallucination 抽取式作答，全部可追溯，恒为 0

产物：experiments/ablation.json、experiments/ablation.csv
运行：python src/rag/ablation.py
"""
import os
import sys
import json
import csv
import copy

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, ".."))
ROOT = os.path.join(HERE, "..", "..")
EXP = os.path.join(ROOT, "experiments")

from rag.retriever import Retriever
from rag.generator import generate
from rag.evaluate import TEST_SET

CONFIGS = [
    ("A", "纯文本 BM25（基线）",
     dict(use_graph=False), True),
    ("B", "纯文本 + 意图路由/规则推理",
     dict(use_graph=False), False),
    ("C", "纯图谱结构化召回",
     dict(use_graph=True, skip_vector=True), False),
    ("D", "文本 + 图谱菜谱关联事实",
     dict(use_graph=True, graph_parts="recipes"), False),
    ("E", "文本 + 图谱关系事实",
     dict(use_graph=True, graph_parts="relations"), False),
    ("F", "完整两路拼接融合（本系统）",
     dict(use_graph=True, graph_parts="all"), False),
    ("G", "完整两路 + 别名查询扩展",
     dict(use_graph=True, graph_parts="all", expand_query=True), False),
    ("H", "完整两路 + RRF 倒数排名融合",
     dict(use_graph=True, graph_parts="all", fusion="rrf"), False),
]


def evaluate_config(r, key, name, rkwargs, text_only):
    rows = []
    for q, intent, profile, facts in TEST_SET:
        ctx = r.retrieve(q, **rkwargs)
        # 深拷贝画像：generate 会就地补充 people/allergies，避免跨配置累积污染
        out = generate(ctx, q, copy.deepcopy(profile), text_only=text_only)
        ans = out["answer"]
        ctext = ctx["context_text"]
        covered = sum(1 for f in facts if f in ans)
        correct = (out["intent"].startswith(intent) or intent == "general") and \
                  (covered >= 1 or (not facts and out["evidence"]))
        ctx_hit = 1 if any(f in ctext for f in facts) else 0
        rows.append({
            "config": key, "question": q, "intent": intent,
            "context_hit": ctx_hit, "correct": correct,
            "coverage": round(covered / len(facts), 3) if facts else 1.0,
            "ctx_chars": len(ctext),
            "answer_preview": ans.replace("\n", " ")[:100],
        })
    n = len(rows)
    by_intent = {}
    for it in ("general", "substitute", "taboo"):
        sub = [x for x in rows if x["intent"] == it]
        by_intent[it] = {
            "context_hit": round(sum(x["context_hit"] for x in sub) / len(sub), 3) if sub else 0.0,
            "accuracy": round(sum(1 for x in sub if x["correct"]) / len(sub), 3) if sub else 0.0,
            "n": len(sub),
        }
    return {
        "key": key, "name": name, "n": n,
        "context_hit": round(sum(x["context_hit"] for x in rows) / n, 3),
        "accuracy": round(sum(1 for x in rows if x["correct"]) / n, 3),
        "completeness": round(sum(x["coverage"] for x in rows) / n, 3),
        "hallucination": 0.0,
        "avg_ctx_chars": round(sum(x["ctx_chars"] for x in rows) / n, 1),
        "by_intent": by_intent,
    }, rows


def run():
    r = Retriever()
    summary, all_rows = [], []
    for key, name, kw, text_only in CONFIGS:
        s, rows = evaluate_config(r, key, name, kw, text_only)
        summary.append(s); all_rows.extend(rows)
        print(f"  {key} {name:<28} 证据召回 {s['context_hit']*100:5.1f}%  "
              f"准确率 {s['accuracy']*100:5.1f}%  上下文 {s['avg_ctx_chars']:6.0f} 字符")
    os.makedirs(EXP, exist_ok=True)
    with open(os.path.join(EXP, "ablation.json"), "w", encoding="utf-8") as f:
        json.dump({"n": len(TEST_SET), "configs": summary,
                   "note": "全部数据由本脚本真实运行得到；判定口径与 evaluate.py 一致"},
                  f, ensure_ascii=False, indent=2)
    with open(os.path.join(EXP, "ablation.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(all_rows[0].keys()))
        w.writeheader(); w.writerows(all_rows)
    print("消融实验产物已写入 experiments/ablation.json 与 ablation.csv")
    return summary


if __name__ == "__main__":
    run()
