# -*- coding: utf-8 -*-
"""
检索层：两路召回 + 简单融合（对标方案"向量文本召回 + 图谱结构化召回"）。
  路径 A（向量文本召回）：BM25 召回 Top-K 相关菜谱文本。
  路径 B（图谱结构化召回）：从查询抽取实体(食材/菜名)，查图谱得到
        - 包含该食材的菜谱
        - 食材的"替代/相克/搭配"关系
        - 相关功效/禁忌事实
  融合：两路结果统一拼接为结构化上下文，送入生成器。
"""
import os
import sys
import json
import pickle

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, ".."))
ROOT = os.path.join(HERE, "..", "..")
DATA_DIR = os.path.join(ROOT, "data")

from rag.vector_store import BM25, build_corpus
from kg.graph_store import GraphStore
from extraction.rule_extractor import disambiguate, INGREDIENT_MASTER, ALIAS_MAP


class Retriever:
    def __init__(self):
        self.bm = self._load_bm25()
        self.gs = GraphStore()
        # 构建 食材 -> 菜谱 反查（用于图谱结构化召回）
        self.ing2recipes = {}
        with open(os.path.join(DATA_DIR, "raw_recipes.json"), encoding="utf-8") as f:
            self.recipes = json.load(f)
        for r in self.recipes:
            for i in r.get("ingredients", []):
                std = disambiguate(i["name"])
                self.ing2recipes.setdefault(std, []).append(r["name"])

    def _load_bm25(self):
        pkl = os.path.join(DATA_DIR, "bm25.pkl")
        if not os.path.exists(pkl):
            from rag.vector_store import build_index
            return build_index()
        with open(pkl, "rb") as f:
            d = pickle.load(f)
        from rag.vector_store import BM25
        if isinstance(d, dict):
            return BM25.from_dict(d)
        return d

    def _extract_entities(self, query):
        """从查询中识别已知食材/菜名实体。"""
        ings, recipes_hit = [], []
        for name in INGREDIENT_MASTER:
            if name in query:
                ings.append(name)
        for alias, std in ALIAS_MAP.items():
            if alias in query and std not in ings:
                ings.append(std)
        for r in self.recipes:
            if r["name"] in query:
                recipes_hit.append(r["name"])
        return ings, recipes_hit

    def _expand_query(self, query):
        """Advanced RAG 查询侧改写（轻量版）：把识别到实体的别名并入 BM25 查询。"""
        ings, _ = self._extract_entities(query)
        extra = []
        for ing in ings:
            for alias, std in ALIAS_MAP.items():
                if std == ing and alias not in query:
                    extra.append(alias)
        seen, terms = set(), []
        for x in extra:
            if x not in seen:
                seen.add(x); terms.append(x)
        return (query + " " + " ".join(terms)).strip()

    def vector_recall(self, query, top_k=5):
        return self.bm.search(query, top_k=top_k)

    def graph_recall(self, query, parts="all"):
        """图谱结构化召回。
        parts: all=菜谱关联+关系事实（默认）；recipes=仅菜谱关联；relations=仅关系事实。
        该分片开关仅用于消融实验，不改变默认行为。
        """
        ings, recipes_hit = self._extract_entities(query)
        facts = []
        related_recipes = set()
        for ing in ings:
            for rn in self.ing2recipes.get(ing, []):
                related_recipes.add(rn)
        if recipes_hit:
            related_recipes.update(recipes_hit)
        # 1) 包含这些食材的菜谱（sorted 保证可复现）
        if parts in ("all", "recipes"):
            for rn in sorted(related_recipes)[:8]:
                facts.append(f"[图谱] 菜谱『{rn}』包含食材：{', '.join(self._recipe_ings(rn))}")
        # 2) 食材关系：替代/相克/搭配
        if parts in ("all", "relations"):
            for ing in ings:
                node = self.gs.node_detail(ing)
                if not node:
                    continue
                for rel in node["out"]:
                    if rel["relation"] in ("替代", "相克", "搭配"):
                        facts.append(f"[图谱] 食材『{ing}』{rel['relation']}『{rel['target']}』")
        return {"entities": ings, "facts": facts, "related_recipes": sorted(related_recipes)}

    def _recipe_ings(self, name):
        for r in self.recipes:
            if r["name"] == name:
                return [i["name"] for i in r.get("ingredients", []) if i["category"] == "食材"]
        return []

    @staticmethod
    def _rrf_fuse(vec, facts, k=60):
        """简化 RRF 倒数排名融合：异构证据在两路各自排名上按 1/(k+r) 叠加后重排。
        标准 RRF 面向同一文档集合的多路排名；此处用于文本证据与图谱证据两类异构单元的融合。
        """
        ev = []
        for i, v in enumerate(vec):
            ev.append({"text": f"【相关菜谱】{v['name']}：{v['text'][:160]}",
                       "name": v["name"], "ranks": [i + 1]})
        prio = {"相克": 0, "替代": 1, "搭配": 2, "包含": 3}
        def gkey(f):
            for rel, p in prio.items():
                if rel in f:
                    return p
            return 9
        gf = sorted(facts, key=gkey)
        for j, f in enumerate(gf):
            hit = None
            for e in ev:
                if e["name"] and e["name"] in f:
                    hit = e; break
            if hit is not None:
                hit["ranks"].append(j + 1)
            else:
                ev.append({"text": f, "name": None, "ranks": [j + 1]})
        for e in ev:
            e["score"] = sum(1.0 / (k + r) for r in e["ranks"])
        ev.sort(key=lambda x: -x["score"])
        return "\n".join(e["text"] for e in ev)

    def retrieve(self, query, top_k=5, use_graph=True, graph_parts="all",
                 expand_query=False, skip_vector=False, fusion="concat"):
        """两路召回 + 融合。默认参数与原实现完全等价，保证主实验可复现。
        扩展参数供消融实验使用：graph_parts / expand_query / skip_vector / fusion。
        """
        q = self._expand_query(query) if expand_query else query
        vec = [] if skip_vector else self.vector_recall(q, top_k=top_k)
        ctx = {"query": query, "vector_results": vec, "graph_results": None,
               "context_text": "", "graph_facts": [], "fusion": fusion}
        if use_graph:
            g = self.graph_recall(query, parts=graph_parts)
            ctx["graph_results"] = g
            ctx["graph_facts"] = g["facts"]
        if fusion == "rrf":
            ctx["context_text"] = self._rrf_fuse(vec, ctx["graph_facts"])
        else:
            parts = [f"【相关菜谱】{v['name']}：{v['text'][:160]}" for v in vec]
            parts.extend(ctx["graph_facts"])
            ctx["context_text"] = "\n".join(parts)
        return ctx


if __name__ == "__main__":
    r = Retriever()
    print("=== 向量召回：牛奶过敏 替代 ===")
    for v in r.vector_recall("牛奶过敏 用什么替代", 3):
        print("  ", v["name"], v["score"])
    print("\n=== 图谱结构化召回 ===")
    g = r.graph_recall("牛奶过敏 用什么替代")
    for f in g["facts"][:6]:
        print("  ", f)
