# -*- coding: utf-8 -*-
"""
向量/文本召回层（纯 Python，零外部依赖，可离线运行）：
  - 中文分词：优先 jieba，否则退化为"字 + 二元组"切分（检索足够用）
  - 索引：BM25 倒排索引（轻量、无需训练、无需 GPU/联网）
  - 检索：给定查询返回 Top-K 相关菜谱文本片段
对标方案"向量文本召回"一路；如需更强语义向量，可在 config 中切换 sentence-transformers。
"""
import os
import re
import json
import math
import pickle

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..", "..")
sys_path = os.path.join(HERE, "..")
import sys
sys.path.insert(0, sys_path)
DATA_DIR = os.path.join(ROOT, "data")

from extraction.rule_extractor import INGREDIENT_MASTER

# 可选：jieba 增强分词
try:
    import jieba
    jieba.setLogLevel(20)
    _HAS_JIEBA = True
except Exception:
    _HAS_JIEBA = False


def tokenize(text: str):
    """中文分词：jieba 词 + 词二元组；ASCII 词保留。"""
    text = (text or "").lower()
    tokens = []
    for m in re.findall(r"[a-z0-9]+", text):
        tokens.append(m)
    han = re.findall(r"[一-龥]", text)
    if _HAS_JIEBA:
        words = [w for w in jieba.cut(text) if re.search(r"[一-龥a-z0-9]", w)]
        for w in words:
            tokens.append(w)
            if len(w) >= 2:
                for i in range(len(w) - 1):
                    tokens.append(w[i:i + 2])
    else:
        tokens.extend(han)
        for i in range(len(han) - 1):
            tokens.append(han[i] + han[i + 1])
    return [t for t in tokens if t.strip()]


class BM25:
    def __init__(self, k1=1.5, b=0.75):
        self.k1 = k1
        self.b = b
        self.docs = []          # list of token lists
        self.doc_meta = []      # 对应菜谱元信息
        self.df = {}
        self.idf = {}
        self.avgdl = 0.0
        self.N = 0

    def fit(self, documents):
        """documents: list of (meta, text)"""
        self.docs = []
        self.doc_meta = []
        for meta, text in documents:
            toks = tokenize(text)
            self.docs.append(toks)
            self.doc_meta.append(meta)
        self.N = len(self.docs)
        self.avgdl = sum(len(d) for d in self.docs) / max(self.N, 1)
        for toks in self.docs:
            seen = set(toks)
            for t in seen:
                self.df[t] = self.df.get(t, 0) + 1
        for t, d in self.df.items():
            self.idf[t] = math.log(1 + (self.N - d + 0.5) / (d + 0.5))
        return self

    def search(self, query: str, top_k=5):
        q_tokens = tokenize(query)
        scores = {}
        for q in q_tokens:
            if q not in self.idf:
                continue
            idf = self.idf[q]
            for i, doc in enumerate(self.docs):
                # 词频
                f = doc.count(q)
                if f == 0:
                    continue
                dl = len(doc)
                denom = f + self.k1 * (1 - self.b + self.b * dl / max(self.avgdl, 1e-9))
                scores[i] = scores.get(i, 0.0) + idf * (f * (self.k1 + 1) / denom)
        ranked = sorted(scores.items(), key=lambda x: x[1], reverse=True)[:top_k]
        results = []
        for i, s in ranked:
            meta = dict(self.doc_meta[i])
            meta["score"] = round(s, 4)
            results.append(meta)
        return results

    def to_dict(self):
        return {"k1": self.k1, "b": self.b, "docs": self.docs, "doc_meta": self.doc_meta,
                "df": self.df, "idf": self.idf, "avgdl": self.avgdl, "N": self.N}

    @classmethod
    def from_dict(cls, d):
        bm = cls(k1=d["k1"], b=d["b"])
        bm.docs = d["docs"]; bm.doc_meta = d["doc_meta"]; bm.df = d["df"]
        bm.idf = d["idf"]; bm.avgdl = d["avgdl"]; bm.N = d["N"]
        return bm


def build_corpus(recipes):
    """由结构化菜谱构建检索文档。"""
    documents = []
    for r in recipes:
        ings = "、".join(i["name"] for i in r.get("ingredients", []))
        steps = " ".join(r.get("steps", []))
        # 聚合营养标签（来自食材主表）
        nutri = set()
        for i in r.get("ingredients", []):
            a = INGREDIENT_MASTER.get(i["name"])
            if a:
                nutri.update(a.get("nutrition", []))
        text = (f"{r['name']}。菜系：{r.get('cuisine','')}。{r.get('description','')} "
                f"食材：{ings}。口味：{'、'.join(r.get('flavors',[]))}。"
                f"功效：{'、'.join(r.get('effects',[]))}。"
                f"营养：{'、'.join(sorted(nutri))}。"
                f"适用：{'、'.join(r.get('suitable_for',[]))}。做法：{steps}")
        documents.append(({"id": r["id"], "name": r["name"], "text": text,
                           "flavors": r.get("flavors", []), "effects": r.get("effects", []),
                           "ingredients": [i["name"] for i in r.get("ingredients", [])]},
                          text))
    return documents


def build_index(recipes_path=None, out_pkl=None):
    if recipes_path is None:
        recipes_path = os.path.join(DATA_DIR, "raw_recipes.json")
    if out_pkl is None:
        out_pkl = os.path.join(DATA_DIR, "bm25.pkl")
    with open(recipes_path, encoding="utf-8") as f:
        recipes = json.load(f)
    docs = build_corpus(recipes)
    bm = BM25().fit(docs)
    with open(out_pkl, "wb") as f:
        pickle.dump(bm.to_dict(), f)
    print(f"BM25 索引构建完成：{bm.N} 篇文档，词表 {len(bm.idf)}")
    return bm


if __name__ == "__main__":
    bm = build_index()
    for r in bm.search("产妇 补气血 汤", top_k=3):
        print(round(r["score"], 3), r["name"])
