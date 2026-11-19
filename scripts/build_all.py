# -*- coding: utf-8 -*-
"""
一键构建：数据生成 -> 知识图谱构建 -> 向量索引构建。
运行：python3 scripts/build_all.py  （在 D:\rag 目录下）
"""
import os
import sys
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..")
sys.path.insert(0, os.path.join(ROOT, "src"))
SRC = os.path.join(ROOT, "src")


def run(py_file, label):
    print(f"\n>>> [{label}] {py_file}")
    r = subprocess.run([sys.executable, os.path.join(SRC, py_file)],
                       cwd=os.path.join(SRC, os.path.dirname(py_file)))
    if r.returncode != 0:
        print(f"!!! {label} 失败，请检查上方报错")
        sys.exit(1)


if __name__ == "__main__":
    print("===== 开始构建烹饪知识图谱 RAG 系统数据 =====")
    run("data_generator.py", "1. 生成菜谱数据集")
    run("kg/build_graph.py", "2. 构建知识图谱")
    run("rag/vector_store.py", "3. 构建 BM25 向量索引")
    print("\n===== 构建完成 =====")
    print("启动系统： python3 src/api/server.py   然后浏览器打开 http://localhost:8000")
