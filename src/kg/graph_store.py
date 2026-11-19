# -*- coding: utf-8 -*-
"""
图谱存储与查询封装：后端 API 调用此模块完成图谱检索、子图抽取、实体搜索等。
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
GRAPH_PKL = os.path.join(DATA_DIR, "graph.pkl")
GRAPH_JSON = os.path.join(DATA_DIR, "graph.json")


class GraphStore:
    def __init__(self, pkl_path=GRAPH_PKL):
        self.pkl_path = pkl_path
        self.G = self._load()

    def _load(self):
        if not os.path.exists(self.pkl_path):
            from kg.build_graph import build
            build()
        with open(self.pkl_path, "rb") as f:
            return pickle.load(f)

    def reload(self):
        self.G = self._load()

    def stats(self):
        types = {}
        for n in self.G.nodes():
            t = self.G.nodes[n].get("type")
            types[t] = types.get(t, 0) + 1
        rels = {}
        for _, _, d in self.G.edges(data=True):
            rels[d["relation"]] = rels.get(d["relation"], 0) + 1
        return {"nodes": self.G.number_of_nodes(), "edges": self.G.number_of_edges(),
                "type_counter": types, "rel_counter": rels}

    def search_node(self, keyword, limit=20):
        kw = (keyword or "").strip()
        res = []
        for n in self.G.nodes():
            if kw and kw in n:
                res.append({"id": n, "type": self.G.nodes[n].get("type"),
                            "cat": self.G.nodes[n].get("cat", ""),
                            "nature": self.G.nodes[n].get("nature", ""),
                            "allergen": self.G.nodes[n].get("allergen", "")})
            if len(res) >= limit:
                break
        return res

    def node_detail(self, name):
        if name not in self.G:
            return None
        attrs = dict(self.G.nodes[name])
        out_rels = []
        for _, v, d in self.G.out_edges(name, data=True):
            out_rels.append({"relation": d["relation"], "target": v,
                             "target_type": self.G.nodes[v].get("type")})
        in_rels = []
        for u, _, d in self.G.in_edges(name, data=True):
            in_rels.append({"relation": d["relation"], "source": u,
                            "source_type": self.G.nodes[u].get("type")})
        return {"id": name, "attrs": attrs, "out": out_rels, "in": in_rels}

    def subgraph(self, names, radius=1):
        """抽取以 names 为中心、半径 radius 的子图（用于可视化与图谱召回）。"""
        nodes = set(names)
        for name in list(names):
            if name not in self.G:
                continue
            for _ in range(radius):
                new = set()
                for u, v in self.G.out_edges(name):
                    if v not in nodes:
                        new.add(v)
                for u, v in self.G.in_edges(name):
                    if u not in nodes:
                        new.add(u)
                nodes.update(new)
                name = None
        # 重新以集合内节点做 BFS 半径
        result = set(names)
        frontier = set(names)
        for _ in range(radius):
            nxt = set()
            for n in frontier:
                if n not in self.G:
                    continue
                nxt.update(self.G.successors(n))
                nxt.update(self.G.predecessors(n))
            nxt -= result
            result |= nxt
            frontier = nxt
        sg = self.G.subgraph(result)
        nodes_out = [{"id": n, "type": sg.nodes[n].get("type"),
                      "cat": sg.nodes[n].get("cat", ""),
                      "nature": sg.nodes[n].get("nature", ""),
                      "allergen": sg.nodes[n].get("allergen", "")} for n in sg.nodes()]
        edges_out = [{"source": u, "target": v, "relation": d["relation"]}
                     for u, v, d in sg.edges(data=True)]
        return {"nodes": nodes_out, "edges": edges_out,
                "stats": {"nodes": len(nodes_out), "edges": len(edges_out)}}


if __name__ == "__main__":
    gs = GraphStore()
    print(gs.stats())
    print("\n搜索'番茄':", gs.search_node("番茄", 5))
    print("\n'番茄炒蛋'详情出边:", [ (r['relation'], r['target']) for r in gs.node_detail('番茄炒蛋')['out'] ])
    sub = gs.subgraph(["牛肉"], radius=1)
    print("\n'牛肉'一跳子图节点数:", sub["stats"]["nodes"])
