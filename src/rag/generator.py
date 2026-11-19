# -*- coding: utf-8 -*-
"""
生成层：图谱增强 RAG 的答案生成。
  - 意图识别：替代 / 禁忌 / 通用菜谱问答
  - 调用推理模块（食材替代、饮食禁忌）与图谱/文本上下文
  - 生成方式：优先调用可插拔大模型；未配置时采用"抽取式"作答
    （从检索上下文抽取事实组织成自然语言，保证可离线运行且幻觉率极低）
"""
import os
import sys
import json

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, ".."))

from reasoning.substitute import recommend, substitute_in_recipe
from reasoning.taboo import check_ingredients
from extraction.rule_extractor import INGREDIENT_MASTER, disambiguate

# 可插拔大模型（与抽取层共用环境变量）
LLM_BASE_URL = os.environ.get("RAG_LLM_BASE_URL", "")
LLM_API_KEY = os.environ.get("RAG_LLM_API_KEY", "")
LLM_MODEL = os.environ.get("RAG_LLM_MODEL", "gpt-3.5-turbo")


# 体质/人群健康触发词（命中即走禁忌校验）
HEALTH_TERMS = ["脾胃虚寒", "虚寒", "湿热", "热性体质", "寒性体质", "结石体质",
                "糖尿病", "高血压", "痛风", "海鲜过敏", "牛奶过敏", "孕妇", "经期",
                "哺乳期", "贫血", "体虚", "上火"]


def _detect_intent(query):
    q = query
    if any(k in q for k in ["替代", "没有", "缺", "换", "代替", "用啥", "用什么", "换成", "顶替"]):
        return "substitute"
    if any(k in q for k in ["禁忌", "忌口", "过敏", "能不能吃", "能吃吗", "体质", "相克",
                            "慎用", "忌食", "一起吃", "同食", "冲突", "能不能", "适合吃",
                            "能吃", "不能吃", "可以吃", "慎吃", "不宜"]):
        return "taboo"
    if any(t in q for t in HEALTH_TERMS):
        return "taboo"
    return "general"


def _extract_ingredients(query):
    found = []
    for name in INGREDIENT_MASTER:
        if name in query:
            found.append(name)
    from extraction.rule_extractor import ALIAS_MAP
    for alias, std in ALIAS_MAP.items():
        if alias in query and std not in found:
            found.append(std)
    return found


def generate(ctx, query, profile=None, text_only=False):
    """
    :param ctx: Retriever.retrieve 的返回（含 vector_results / graph_facts）
    :param query: 用户问题
    :param profile: 用户画像（过敏/体质/人群），用于禁忌与替代
    :param text_only: True 时仅基于召回文本作答（纯文本 RAG，不调用图谱推理）
    :return: {answer, intent, evidence}
    """
    intent = _detect_intent(query)
    evidence = []

    if text_only:
        # 纯文本 RAG：仅汇总向量召回的菜谱文本，不调用图谱/推理模块
        vec = ctx.get("vector_results", [])
        lines = ["【以下是相关菜谱文本片段】"]
        for v in vec[:5]:
            lines.append(f"■ {v['name']}：{v['text'][:140]}")
        return {"answer": "\n".join(lines), "intent": "general(text-only)",
                "evidence": [v["name"] for v in vec[:5]], "grounded": True}

    if intent == "substitute":
        ings = _extract_ingredients(query)
        # 若查询明确点名某食材，则针对它；否则取向量召回命中的食材
        if not ings:
            for v in ctx.get("vector_results", []):
                for i in v.get("ingredients", []):
                    si = disambiguate(i)
                    if si in INGREDIENT_MASTER:
                        ings.append(si)
                if len(ings) >= 2:
                    break
        profile = profile or {}
        lines = []
        for ing in ings[:3]:
            alts = recommend(ing, user_allergies=profile.get("allergies", []),
                             dish_ingredients=ings, top_k=3)
            if alts:
                s = f"『{ing}』可替代为：" + "；".join(
                    f"{a['name']}（{a['reason']}）" for a in alts)
            else:
                s = f"『{ing}』暂无可推荐替代。"
            lines.append(s)
            evidence.append(ing)
        answer = "【食材替代建议】\n" + "\n".join(lines)

    elif intent == "taboo":
        ings = _extract_ingredients(query)
        if not ings:
            for v in ctx.get("vector_results", []):
                ings.extend(v.get("ingredients", []))
        profile = profile or {}
        # 若查询含人群/体质关键词，并入 profile
        if "孕妇" in query: profile.setdefault("people", []).append("孕妇")
        if "高血压" in query: profile.setdefault("people", []).append("高血压")
        if "糖尿病" in query: profile.setdefault("people", []).append("糖尿病")
        if "痛风" in query: profile.setdefault("people", []).append("痛风")
        if "海鲜过敏" in query: profile.setdefault("allergies", []).append("海鲜")
        if "牛奶过敏" in query or "奶过敏" in query: profile.setdefault("allergies", []).append("奶")
        res = check_ingredients(ings, profile)
        lines = [f"风险等级：{res['risk_level']}"]
        if res["warnings"]:
            lines.append("风险提示：")
            for w in res["warnings"]:
                lines.append(f"  - [{w['level']}] {w['detail']}")
        else:
            lines.append("未发现明显禁忌，可放心食用。")
        if res["suggestions"]:
            lines.append("调整建议：")
            for s in res["suggestions"]:
                lines.append(f"  - {s}")
        answer = "【饮食禁忌校验】\n" + "\n".join(lines)
        evidence = ings

    else:  # general
        vec = ctx.get("vector_results", [])
        lines = ["【为你找到以下相关菜谱】"]
        for v in vec[:5]:
            lines.append(f"■ {v['name']}（口味：{','.join(v.get('flavors', []))}；"
                         f"功效：{','.join(v.get('effects', []))}）")
            lines.append(f"   食材：{', '.join(v.get('ingredients', []))}")
            lines.append(f"   简介：{v['text'][:120]}")
        # 图谱增强：补充结构化事实
        gf = ctx.get("graph_facts", [])
        if gf:
            lines.append("【图谱补充知识】")
            for f in gf[:6]:
                lines.append(f"  - {f}")
        answer = "\n".join(lines)
        evidence = [v["name"] for v in vec[:5]]

    return {"answer": answer, "intent": intent, "evidence": evidence,
            "grounded": True}  # 抽取式作答，全部基于检索上下文，幻觉率低


def generate_with_llm(ctx, query, profile=None):
    """若配置大模型，则把融合上下文交给 LLM 生成（生成式 RAG）。"""
    if not (LLM_BASE_URL and LLM_API_KEY):
        # 退化为抽取式
        return generate(ctx, query, profile)
    try:
        import requests
        context = ctx.get("context_text", "")
        prompt = (f"你是烹饪知识助手。请基于以下检索上下文回答用户问题，"
                  f"不要编造上下文之外的信息。\n\n上下文：\n{context}\n\n问题：{query}\n回答：")
        resp = requests.post(f"{LLM_BASE_URL}/chat/completions",
                             headers={"Authorization": f"Bearer {LLM_API_KEY}",
                                      "Content-Type": "application/json"},
                             json={"model": LLM_MODEL, "messages": [
                                 {"role": "system", "content": "你是烹饪知识图谱问答助手。"},
                                 {"role": "user", "content": prompt}]}, timeout=30)
        ans = resp.json()["choices"][0]["message"]["content"]
        return {"answer": ans, "intent": _detect_intent(query),
                "evidence": [v["name"] for v in ctx.get("vector_results", [])[:5]],
                "grounded": False}
    except Exception:
        return generate(ctx, query, profile)


if __name__ == "__main__":
    from rag.retriever import Retriever
    r = Retriever()
    for q in ["牛奶过敏能用什么替代？", "孕妇能吃螃蟹吗？", "补气血的汤有哪些？"]:
        ctx = r.retrieve(q, use_graph=True)
        out = generate(ctx, q)
        print("Q:", q, "| intent:", out["intent"])
        print(out["answer"])
        print("-" * 40)
