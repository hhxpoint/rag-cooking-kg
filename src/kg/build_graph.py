# -*- coding: utf-8 -*-
"""
知识图谱构建：将菜谱数据抽取为三元组并构建 NetworkX 有向图。
节点类型：菜谱/食材/调料/工序/口味/禁忌/功效/人群
边关系：包含/属于/搭配/相克/适用/具有功效/采用工序/替代
产物：
  data/graph.json        轻量图（前端可视化直接消费：nodes/edges）
  data/graph.pkl         NetworkX 对象（后端查询用）
"""
import os
import sys
import json
import pickle
import networkx as nx

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..", "..")
sys.path.insert(0, os.path.join(HERE, ".."))   # 将 src 加入搜索路径
DATA_DIR = os.path.join(ROOT, "data")

from extraction.rule_extractor import from_structured, disambiguate
from extraction.schema import EntityType


def node_type_of(name, category_hint=None):
    """依据尾实体类别或上下文推断节点类型。"""
    if category_hint == "食材":
        return EntityType.INGREDIENT.value
    if category_hint == "调料":
        return EntityType.SEASONING.value
    if category_hint in ("工序", "口味", "功效", "人群"):
        return category_hint
    if category_hint == "禁忌":
        return EntityType.TABOO.value
    return EntityType.RECIPE.value


def build(recipes_path=None, out_json=None, out_pkl=None):
    if recipes_path is None:
        recipes_path = os.path.join(DATA_DIR, "raw_recipes.json")
    if out_json is None:
        out_json = os.path.join(DATA_DIR, "graph.json")
    if out_pkl is None:
        out_pkl = os.path.join(DATA_DIR, "graph.pkl")

    with open(recipes_path, encoding="utf-8") as f:
        recipes = json.load(f)

    # 载入食材主表，用于补充"替代"边与节点属性
    with open(os.path.join(DATA_DIR, "ingredient_master.json"), encoding="utf-8") as f:
        master = json.load(f)
    # 载入相克对，并入图谱（使禁忌推理可走图查询）
    with open(os.path.join(DATA_DIR, "incompatible_pairs.json"), encoding="utf-8") as f:
        incompatible_pairs = json.load(f)

    G = nx.DiGraph()
    node_types = {}        # name -> type
    node_attrs = {}        # name -> 附加属性

    def ensure_node(name, ntype, **attrs):
        if name not in G:
            G.add_node(name)
            node_types[name] = ntype
            node_attrs[name] = dict(attrs)
        else:
            # 类型以更具体的优先（食材/调料 > 菜谱）
            if node_types[name] == EntityType.RECIPE.value and ntype != EntityType.RECIPE.value:
                node_types[name] = ntype
            node_attrs[name].update(attrs)

    edge_count = 0
    for r in recipes:
        triples = from_structured(r)
        for head, rel, tail, tail_cat in triples:
            h_type = EntityType.RECIPE.value
            ensure_node(head, h_type, source="recipe")
            t_type = node_type_of(tail, tail_cat)
            # 尾实体若为标准食材/调料，补充主表属性
            extra = {}
            if tail in master:
                extra = {"cat": master[tail]["cat"], "nature": master[tail].get("nature"),
                         "allergen": master[tail].get("allergen")}
            ensure_node(tail, t_type, **extra)
            if not G.has_edge(head, tail):
                G.add_edge(head, tail, relation=rel)
                edge_count += 1

    # 补充"替代"边（来自食材主表的 substitutes，作为推理层派生的结构化关系）
    sub_edges = 0
    for name, attr in master.items():
        ensure_node(name, EntityType.INGREDIENT.value, cat="食材", nature=attr.get("nature"),
                    allergen=attr.get("allergen"))
        for sub in attr.get("substitutes", []):
            sub_std = disambiguate(sub)
            if sub_std in master and sub_std != name:
                ensure_node(sub_std, EntityType.INGREDIENT.value, cat="食材",
                            nature=master[sub_std].get("nature"),
                            allergen=master[sub_std].get("allergen"))
                if not G.has_edge(name, sub_std):
                    G.add_edge(name, sub_std, relation="替代")
                    sub_edges += 1

    # 补充"相克"边（来自食材主表定义的相克关系对，方向无关）
    incom_edges = 0
    for a, b in incompatible_pairs:
        a_s, b_s = disambiguate(a), disambiguate(b)
        if a_s in master and b_s in master and a_s != b_s:
            ensure_node(a_s, EntityType.INGREDIENT.value, cat="食材",
                        nature=master[a_s].get("nature"), allergen=master[a_s].get("allergen"))
            ensure_node(b_s, EntityType.INGREDIENT.value, cat="食材",
                        nature=master[b_s].get("nature"), allergen=master[b_s].get("allergen"))
            if not G.has_edge(a_s, b_s):
                G.add_edge(a_s, b_s, relation="相克")
                incom_edges += 1
    sub_edges += 0  # 占位，避免未使用变量

    # 写回节点类型与属性
    for n in G.nodes():
        G.nodes[n]["type"] = node_types.get(n, EntityType.RECIPE.value)
        G.nodes[n].update(node_attrs.get(n, {}))

    # 统计
    type_counter = {}
    for n in G.nodes():
        t = G.nodes[n]["type"]
        type_counter[t] = type_counter.get(t, 0) + 1
    rel_counter = {}
    for _, _, d in G.edges(data=True):
        rel_counter[d["relation"]] = rel_counter.get(d["relation"], 0) + 1

    # 保存 NetworkX 对象（pickle）
    with open(out_pkl, "wb") as f:
        pickle.dump(G, f)

    # 保存轻量图（前端可视化）
    nodes = [{"id": n, "type": G.nodes[n]["type"],
              "cat": G.nodes[n].get("cat", ""), "nature": G.nodes[n].get("nature", ""),
              "allergen": G.nodes[n].get("allergen", "")} for n in G.nodes()]
    edges = [{"source": u, "target": v, "relation": d["relation"]}
             for u, v, d in G.edges(data=True)]
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump({"nodes": nodes, "edges": edges,
                   "stats": {"nodes": len(nodes), "edges": len(edges),
                             "type_counter": type_counter, "rel_counter": rel_counter}},
                  f, ensure_ascii=False, indent=2)

    print(f"图谱节点数: {len(nodes)}  边数: {len(edges)}  (替代边 {sub_edges}, 相克边 {incom_edges})")
    print("节点类型分布:", type_counter)
    print("关系分布:", rel_counter)
    print(f"已保存: {out_json}\n        {out_pkl}")
    return G


if __name__ == "__main__":
    build()
