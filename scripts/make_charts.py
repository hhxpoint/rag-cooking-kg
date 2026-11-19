# -*- coding: utf-8 -*-
"""
生成报告用图表（学术黑白印刷风格）。
严格约束：全图只用黑 / 灰阶 / 白色，无任何彩色；系列用 hatch 纹理区分；
图内不放大标题（图题由 Word 文档以"图 N ..."题注统一给出，符合学术规范）。
输出：report/assets/*.png
运行：python scripts/make_charts.py
"""
import os
import json
import matplotlib
matplotlib.use("Agg")            # 无界面后端
matplotlib.rcParams["axes.unicode_minus"] = False
from matplotlib import font_manager as fm
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
import numpy as np
import networkx as nx

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
DATA = os.path.join(ROOT, "data")
ASSETS = os.path.join(ROOT, "report", "assets")
os.makedirs(ASSETS, exist_ok=True)

FONT_PATH = r"C:\Windows\Fonts\simhei.ttf"
if os.path.exists(FONT_PATH):
    fm.fontManager.addfont(FONT_PATH)
    plt.rcParams["font.sans-serif"] = ["SimHei"]
else:
    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimSun", "DejaVu Sans"]
plt.rcParams["font.size"] = 10
plt.rcParams["axes.edgecolor"] = "black"
plt.rcParams["text.color"] = "black"
plt.rcParams["axes.labelcolor"] = "black"
plt.rcParams["xtick.color"] = "black"
plt.rcParams["ytick.color"] = "black"

# 仅黑 / 灰阶 / 白
INK = "#000000"
G_DARK = "#4d4d4d"
G_MID = "#8c8c8c"
G_LIGHT = "#c8c8c8"
HATCH_A = "///"
HATCH_B = "\\\\\\"


def style_axis(ax, grid=True):
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color(INK); ax.spines["left"].set_linewidth(0.9)
    ax.spines["bottom"].set_color(INK); ax.spines["bottom"].set_linewidth(0.9)
    ax.tick_params(direction="in", length=4, width=0.8, colors=INK)
    if grid:
        ax.yaxis.grid(True, linestyle="--", linewidth=0.5, color=G_LIGHT, alpha=0.9)
        ax.set_axisbelow(True)


def save(fig, name):
    fig.savefig(os.path.join(ASSETS, name), dpi=220, bbox_inches="tight",
                facecolor="white", edgecolor="white")
    plt.close(fig)


def load_stats():
    with open(os.path.join(DATA, "graph.json"), encoding="utf-8") as f:
        return json.load(f)["stats"]


def chart_entity_rel(stats):
    for key, fname in (("type_counter", "entity_dist.png"),
                       ("rel_counter", "relation_dist.png")):
        d = stats[key]
        ks, vs = list(d.keys()), [d[k] for k in d]
        fig, ax = plt.subplots(figsize=(7.0, 3.3))
        ax.bar(range(len(ks)), vs, facecolor="white", edgecolor=INK,
               linewidth=0.9, hatch=HATCH_A, width=0.62)
        ax.set_xticks(range(len(ks)))
        ax.set_xticklabels(ks, fontsize=9, rotation=18, ha="right")
        ax.set_ylabel("节点数" if key == "type_counter" else "边数")
        ax.set_ylim(0, max(vs) * 1.16)
        for i, v in enumerate(vs):
            ax.text(i, v + max(vs) * 0.02, str(v), ha="center", va="bottom",
                    fontsize=8, color=INK)
        style_axis(ax)
        fig.tight_layout()
        save(fig, fname)


def chart_experiment():
    with open(os.path.join(ROOT, "experiments", "summary.json"), encoding="utf-8") as f:
        s = json.load(f)
    metrics = ["accuracy", "completeness", "hallucination"]
    labels = ["回答准确率", "知识点完整度", "幻觉率"]
    pure = [s["pure"][m] * 100 for m in metrics]
    graph = [s["graph"][m] * 100 for m in metrics]
    x = np.arange(len(metrics)); w = 0.34
    fig, ax = plt.subplots(figsize=(7.0, 3.5))
    b1 = ax.bar(x - w / 2, pure, w, facecolor="white", edgecolor=INK,
                linewidth=0.9, hatch=HATCH_A, label="纯文本 RAG")
    b2 = ax.bar(x + w / 2, graph, w, facecolor=G_MID, edgecolor=INK,
                linewidth=0.9, hatch=HATCH_B, label="图谱增强 RAG")
    for bars in (b1, b2):
        for r in bars:
            ax.text(r.get_x() + r.get_width() / 2, r.get_height() + 1.2,
                    f"{r.get_height():.1f}", ha="center", fontsize=8, color=INK)
    ax.set_xticks(x); ax.set_xticklabels(labels, fontsize=10)
    ax.set_ylabel("百分比 (%)"); ax.set_ylim(0, 108)
    ax.legend(frameon=False, fontsize=9, loc="upper right")
    style_axis(ax)
    fig.tight_layout()
    save(fig, "experiment.png")


def chart_ablation():
    path = os.path.join(ROOT, "experiments", "ablation.json")
    if not os.path.exists(path):
        return
    with open(path, encoding="utf-8") as f:
        d = json.load(f)
    cfgs = d["configs"]
    keys = [c["key"] for c in cfgs]
    short = ["A 纯文本", "B +推理", "C 纯图谱", "D +菜谱关联",
             "E +关系事实", "F 完整两路", "G +查询扩展", "H +RRF"]
    labels = short[:len(keys)]
    hit = [c["context_hit"] * 100 for c in cfgs]
    acc = [c["accuracy"] * 100 for c in cfgs]
    x = np.arange(len(keys)); w = 0.34
    fig, ax = plt.subplots(figsize=(7.4, 3.8))
    b1 = ax.bar(x - w / 2, hit, w, facecolor="white", edgecolor=INK,
                linewidth=0.9, hatch=HATCH_A, label="证据召回率")
    b2 = ax.bar(x + w / 2, acc, w, facecolor=G_MID, edgecolor=INK,
                linewidth=0.9, hatch=HATCH_B, label="回答准确率")
    for bars in (b1, b2):
        for r in bars:
            ax.text(r.get_x() + r.get_width() / 2, r.get_height() + 1.2,
                    f"{r.get_height():.1f}", ha="center", fontsize=7, color=INK)
    ax.set_xticks(x); ax.set_xticklabels(labels, fontsize=8, rotation=28, ha="right")
    ax.set_ylabel("百分比 (%)"); ax.set_ylim(0, 108)
    ax.legend(frameon=False, fontsize=9, loc="lower right")
    style_axis(ax)
    fig.tight_layout()
    save(fig, "ablation.png")


def chart_graph_viz():
    with open(os.path.join(DATA, "graph.json"), encoding="utf-8") as f:
        g = json.load(f)
    ntype = {n["id"]: n["type"] for n in g["nodes"]}
    seed = [n["id"] for n in g["nodes"] if n["type"] == "菜谱"][:10]
    nodeset = set(seed)
    for e in g["edges"]:
        if e["source"] in seed:
            nodeset.add(e["target"])
    G = nx.DiGraph()
    for e in g["edges"]:
        if e["source"] in nodeset and e["target"] in nodeset:
            G.add_edge(e["source"], e["target"])
    # 只用灰阶区分实体类型
    gray_map = {"菜谱": "#000000", "食材": "#333333", "调料": "#595959",
                "工序": "#7f7f7f", "口味": "#a6a6a6", "功效": "#c4c4c4", "人群": "#d9d9d9"}
    fig, ax = plt.subplots(figsize=(7.0, 5.0))
    pos = nx.spring_layout(G, seed=42, k=0.55)
    nx.draw_networkx_edges(G, pos, edge_color=G_LIGHT, alpha=0.55,
                           width=0.5, arrows=False, ax=ax)
    for typ, gray in gray_map.items():
        ns = [n for n in G.nodes() if ntype.get(n) == typ]
        if not ns:
            continue
        size = 150 if typ == "菜谱" else 42
        nx.draw_networkx_nodes(G, pos, nodelist=ns, node_color=gray,
                               edgecolors=INK, linewidths=0.5,
                               node_size=size, ax=ax)
    labels = {n: n for n in G.nodes() if ntype.get(n) == "菜谱"}
    nx.draw_networkx_labels(G, pos, labels=labels, font_size=6.5,
                            font_color=INK, ax=ax)
    handles = [Patch(facecolor=v, edgecolor=INK, label=k) for k, v in gray_map.items()
               if any(ntype.get(n) == k for n in G.nodes())]
    ax.legend(handles=handles, frameon=False, fontsize=8,
              loc="upper right", ncol=2, handlelength=1.1)
    ax.axis("off")
    fig.tight_layout()
    save(fig, "graph_viz.png")


if __name__ == "__main__":
    stats = load_stats()
    chart_entity_rel(stats)
    chart_experiment()
    chart_ablation()
    chart_graph_viz()
    print("学术灰度图表已生成：", sorted(os.listdir(ASSETS)))
