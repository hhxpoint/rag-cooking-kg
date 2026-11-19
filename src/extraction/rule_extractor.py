# -*- coding: utf-8 -*-
"""
规则 / 结构化抽取器（离线、零依赖）：
  1. from_structured(recipe)  —— 从结构化菜谱直接产出三元组（高可靠，生产图谱默认路径）
  2. from_text(text)          —— 对自由文本做词典匹配 + 别名消歧（模拟 LLM 抽取输出，
                                  保证无 API/无网络时整套流程仍可运行与演示）
  3. disambiguate(name)       —— 实体消歧：别名 -> 标准名，避免图谱节点重复
"""
import os
import json
import re

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HERE, "..", "..", "data")

# 载入食材主表与别名映射（用于在文本抽取时判定实体类别）
with open(os.path.join(DATA_DIR, "ingredient_master.json"), encoding="utf-8") as f:
    INGREDIENT_MASTER = json.load(f)
with open(os.path.join(DATA_DIR, "alias_map.json"), encoding="utf-8") as f:
    ALIAS_MAP = json.load(f)
with open(os.path.join(DATA_DIR, "incompatible_pairs.json"), encoding="utf-8") as f:
    INCOMPATIBLE_PAIRS = json.load(f)

# 工序词表
PROCESS_WORDS = ["炒", "炖", "蒸", "煮", "烧", "拌", "煎", "焖", "煲汤", "烤"]
# 口味词表
FLAVOR_WORDS = ["咸鲜", "酸甜", "麻辣", "清淡", "鲜香", "酱香", "酸辣", "甜", "香", "家常", "浓郁", "清爽"]
# 人群词表
GROUP_WORDS = ["儿童", "老人", "孕妇", "上班族", "体虚者", "一般人群", "婴幼儿", "产妇"]
# 功效词表（从主表收集）
EFFECT_WORDS = set()
for _a in INGREDIENT_MASTER.values():
    EFFECT_WORDS.update(_a.get("effects", []))
EFFECT_WORDS = list(EFFECT_WORDS)


def disambiguate(name: str) -> str:
    """实体消歧：别名 -> 标准名。"""
    name = (name or "").strip()
    return ALIAS_MAP.get(name, name)


def entity_category(name: str):
    """返回实体类别（食材/调料/其它）。"""
    std = disambiguate(name)
    a = INGREDIENT_MASTER.get(std)
    if a:
        return a["cat"], std
    return None, std


def from_structured(recipe: dict) -> list:
    """从结构化菜谱产出三元组列表。每条三元组为 [头实体, 关系, 尾实体, 尾实体类别]。"""
    triples = []
    rid = recipe["name"]

    # 1) 包含：食材/调料
    for ing in recipe.get("ingredients", []):
        std = disambiguate(ing["name"])
        triples.append([rid, "包含", std, ing["category"]])

    # 2) 采用工序：从菜名/步骤推断
    method = None
    for m in PROCESS_WORDS:
        if m in recipe.get("name", "") or m in " ".join(recipe.get("steps", [])):
            method = m
            break
    if method:
        triples.append([rid, "采用工序", method, "工序"])

    # 3) 属于：口味
    for fl in recipe.get("flavors", []):
        triples.append([rid, "属于", fl, "口味"])

    # 4) 具有功效
    for ef in recipe.get("effects", []):
        triples.append([rid, "具有功效", ef, "功效"])

    # 5) 适用人群
    for g in recipe.get("suitable_for", []):
        triples.append([rid, "适用", g, "人群"])

    # 6) 相克（来自 recipe.taboos 中 type==相克）
    for t in recipe.get("taboos", []):
        if t.get("type") == "相克":
            detail = t.get("detail", "")
            m = re.match(r"(.+?)与(.+?)相克", detail)
            if m:
                a, b = disambiguate(m.group(1)), disambiguate(m.group(2))
                if a != b:
                    triples.append([a, "相克", b, "食材"])

    # 7) 食材之间的"搭配"关系：同一菜谱共现的食材两两搭配（同组不重复）
    ings = [disambiguate(i["name"]) for i in recipe.get("ingredients", [])
            if INGREDIENT_MASTER.get(disambiguate(i["name"]), {}).get("cat") == "食材"]
    for i in range(len(ings)):
        for j in range(i + 1, len(ings)):
            triples.append([ings[i], "搭配", ings[j], "食材"])

    return triples


def from_text(text: str) -> list:
    """对自由文本做词典匹配抽取（离线模拟 LLM 输出）。返回三元组列表。"""
    triples = []
    text = text or ""

    # 抽取所有命中的食材/调料实体
    found = {}
    for std in INGREDIENT_MASTER:
        if std in text:
            found[std] = INGREDIENT_MASTER[std]["cat"]
    # 别名兜底
    for alias, std in ALIAS_MAP.items():
        if alias in text and std not in found:
            found[std] = INGREDIENT_MASTER.get(std, {}).get("cat", "食材")

    entities = list(found.keys())

    # 工序
    for m in PROCESS_WORDS:
        if m in text:
            triples.append(["文本", "采用工序", m, "工序"])
    # 口味
    for fl in FLAVOR_WORDS:
        if fl in text:
            triples.append(["文本", "属于", fl, "口味"])
    # 功效
    for ef in EFFECT_WORDS:
        if ef in text:
            triples.append(["文本", "具有功效", ef, "功效"])
    # 人群
    for g in GROUP_WORDS:
        if g in text:
            triples.append(["文本", "适用", g, "人群"])
    # 相克模式：A与B相克
    for m in re.finditer(r"([一-龥]{1,4})与([一-龥]{1,4})相克", text):
        a, b = disambiguate(m.group(1)), disambiguate(m.group(2))
        if a in found or b in found:
            triples.append([a, "相克", b, "食材"])
    # 搭配：共现食材两两搭配
    ings = [e for e, c in found.items() if c == "食材"]
    for i in range(len(ings)):
        for j in range(i + 1, len(ings)):
            triples.append([ings[i], "搭配", ings[j], "食材"])
    # 包含（若有菜名语境则省略，文本级仅记录实体）
    return triples


if __name__ == "__main__":
    sample = "番茄炒蛋：番茄2个、鸡蛋3个，加食盐炒制，口味酸甜，适合儿童。螃蟹性寒，与柿子相克，孕妇忌食。"
    print("自由文本抽取示例：")
    for t in from_text(sample):
        print("  ", t)
