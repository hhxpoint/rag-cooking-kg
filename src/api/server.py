# -*- coding: utf-8 -*-
"""
后端服务（零依赖：Python 标准库 http.server 实现轻量 REST API）。
提供：图谱统计/搜索/详情/子图、图谱增强 RAG 问答、食材替代、饮食禁忌校验、
后台菜谱 CRUD，并托管前端静态页面（D3.js 可视化 + 后台管理）。
运行：python3 src/api/server.py  （默认端口 8000）
"""
import os
import sys
import json
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, ".."))
ROOT = os.path.join(HERE, "..", "..")
WEB_DIR = os.path.join(HERE, "..", "web")
DATA_DIR = os.path.join(ROOT, "data")
RECIPES_PATH = os.path.join(DATA_DIR, "raw_recipes.json")

from kg.graph_store import GraphStore
from rag.retriever import Retriever
from rag.generator import generate, generate_with_llm
from reasoning.substitute import recommend, substitute_in_recipe
from reasoning.taboo import check_ingredients
from extraction.rule_extractor import disambiguate

_lock = threading.Lock()


class Backend:
    def __init__(self):
        self.gs = GraphStore()
        self.retriever = Retriever()
        with open(RECIPES_PATH, encoding="utf-8") as f:
            self.recipes = json.load(f)
        self.by_id = {r["id"]: r for r in self.recipes}

    # ---------- 数据持久化与重建 ----------
    def save_recipes(self):
        with open(RECIPES_PATH, "w", encoding="utf-8") as f:
            json.dump(self.recipes, f, ensure_ascii=False, indent=2)
        self.by_id = {r["id"]: r for r in self.recipes}

    def rebuild(self):
        from kg.build_graph import build
        from rag.vector_store import build_index
        build()
        build_index()
        self.gs.reload()
        self.retriever = Retriever()

    # ---------- 接口实现 ----------
    def stats(self):
        return self.gs.stats()

    def search(self, q, limit=20):
        return self.gs.search_node(q, limit)

    def node(self, name):
        return self.gs.node_detail(name)

    def subgraph(self, name, radius=1):
        return self.gs.subgraph([name], radius=radius)

    def sample_graph(self, n_recipes=12):
        """取若干种子菜谱及其食材构成可可视化的小图。"""
        names = [r["name"] for r in self.recipes[:n_recipes]]
        sg = self.gs.subgraph(names, radius=1)
        # 限制规模，保证浏览器流畅
        if sg["stats"]["nodes"] > 200:
            sg = {"nodes": sg["nodes"][:200], "edges": sg["edges"][:400],
                  "stats": {"nodes": 200, "edges": 400, "truncated": True}}
        return sg

    def qa(self, question, profile=None, use_llm=False):
        profile = profile or {}
        ctx = self.retriever.retrieve(question, use_graph=True)
        if use_llm:
            out = generate_with_llm(ctx, question, profile)
        else:
            out = generate(ctx, question, profile)
        return out

    def substitute(self, ingredient, allergies=None, dish=None):
        dish = dish or []
        alts = recommend(ingredient, user_allergies=allergies, dish_ingredients=dish, top_k=5)
        return {"ingredient": disambiguate(ingredient), "alternatives": alts}

    def taboo(self, ingredients, profile=None):
        return check_ingredients(ingredients, profile or {})

    def list_recipes(self, q="", limit=50):
        if q:
            res = [r for r in self.recipes if q in r["name"]]
        else:
            res = self.recipes
        return res[:limit]

    def get_recipe(self, rid):
        return self.by_id.get(rid)

    def create_recipe(self, data):
        rid = "R%04d" % (len(self.recipes) + 1)
        data["id"] = rid
        self.recipes.append(data)
        self.save_recipes()
        return data

    def update_recipe(self, rid, data):
        if rid not in self.by_id:
            return None
        data["id"] = rid
        idx = self.recipes.index(self.by_id[rid])
        self.recipes[idx] = data
        self.save_recipes()
        return data

    def delete_recipe(self, rid):
        if rid not in self.by_id:
            return False
        self.recipes.remove(self.by_id[rid])
        self.save_recipes()
        return True


BACKEND = Backend()


class Handler(BaseHTTPRequestHandler):
    def _send(self, code, obj=None, ctype="application/json; charset=utf-8"):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8") if obj is not None else b""
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        if body:
            self.wfile.write(body)

    def _send_file(self, path):
        if not os.path.exists(path) or not os.path.isfile(path):
            self._send(404, {"error": "not found"})
            return
        ext = os.path.splitext(path)[1]
        ctype = {"html": "text/html; charset=utf-8", "js": "application/javascript; charset=utf-8",
                 "css": "text/css; charset=utf-8", "json": "application/json; charset=utf-8",
                 "svg": "image/svg+xml"}.get(ext[1:], "application/octet-stream")
        with open(path, "rb") as f:
            data = f.read()
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _body(self):
        length = int(self.headers.get("Content-Length", 0))
        if length == 0:
            return {}
        try:
            return json.loads(self.rfile.read(length).decode("utf-8"))
        except Exception:
            return {}

    def do_GET(self):
        u = urlparse(self.path)
        path = u.path
        qs = parse_qs(u.query)
        if path == "/" or path == "/index.html":
            self._send_file(os.path.join(WEB_DIR, "index.html")); return
        if path.startswith("/static/"):
            self._send_file(os.path.join(WEB_DIR, path[len("/static/"):])); return
        if path == "/api/stats":
            self._send(200, BACKEND.stats()); return
        if path == "/api/search":
            self._send(200, BACKEND.search(qs.get("q", [""])[0], int(qs.get("limit", [20])[0]))); return
        if path == "/api/node":
            self._send(200, BACKEND.node(qs.get("name", [""])[0])); return
        if path == "/api/subgraph":
            self._send(200, BACKEND.subgraph(qs.get("name", [""])[0], int(qs.get("radius", [1])[0]))); return
        if path == "/api/graph/sample":
            self._send(200, BACKEND.sample_graph(int(qs.get("n", [12])[0]))); return
        if path == "/api/recipes":
            self._send(200, BACKEND.list_recipes(qs.get("q", [""])[0], int(qs.get("limit", [200])[0]))); return
        if path.startswith("/api/recipes/"):
            rid = path.split("/")[-1]
            r = BACKEND.get_recipe(rid)
            self._send(200 if r else 404, r or {"error": "not found"}); return
        self._send(404, {"error": "unknown endpoint"})

    def do_POST(self):
        u = urlparse(self.path)
        path = u.path
        data = self._body()
        if path == "/api/qa":
            out = BACKEND.qa(data.get("question", ""), data.get("profile"),
                             data.get("use_llm", False))
            self._send(200, out); return
        if path == "/api/substitute":
            out = BACKEND.substitute(data.get("ingredient", ""), data.get("allergies"),
                                     data.get("dish"))
            self._send(200, out); return
        if path == "/api/taboo":
            out = BACKEND.taboo(data.get("ingredients", []), data.get("profile"))
            self._send(200, out); return
        if path == "/api/recipes":
            out = BACKEND.create_recipe(data)
            with _lock:
                BACKEND.rebuild()
            self._send(200, out); return
        self._send(404, {"error": "unknown endpoint"})

    def do_PUT(self):
        u = urlparse(self.path)
        path = u.path
        if path.startswith("/api/recipes/"):
            rid = path.split("/")[-1]
            out = BACKEND.update_recipe(rid, self._body())
            if out:
                with _lock:
                    BACKEND.rebuild()
            self._send(200 if out else 404, out or {"error": "not found"}); return
        self._send(404, {"error": "unknown endpoint"})

    def do_DELETE(self):
        u = urlparse(self.path)
        path = u.path
        if path.startswith("/api/recipes/"):
            rid = path.split("/")[-1]
            ok = BACKEND.delete_recipe(rid)
            if ok:
                with _lock:
                    BACKEND.rebuild()
            self._send(200, {"ok": ok}); return
        self._send(404, {"error": "unknown endpoint"})

    def log_message(self, *args):
        pass  # 静默日志


def main():
    port = int(os.environ.get("RAG_PORT", "8000"))
    server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    print(f"烹饪知识图谱 RAG 系统已启动： http://localhost:{port}")
    print("按 Ctrl+C 停止。")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        server.shutdown()


if __name__ == "__main__":
    main()
