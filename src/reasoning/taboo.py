# -*- coding: utf-8 -*-
"""
推理层（二）：饮食禁忌校验推理。
方法：基于领域规则遍历菜谱食材，匹配三类禁忌，给出风险提示与调整建议。
  1) 过敏源校验：菜谱含用户过敏的食材 -> 高风险
  2) 体质/人群校验：菜谱含不适合用户体质/人群的食材 -> 中风险
  3) 相克校验：菜谱内存在相克食材对 -> 高风险
答辩表述："基于领域规则的知识推理（饮食禁忌校验）"。
"""
import os
import sys
import json

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, ".."))   # 将 src 加入搜索路径
DATA_DIR = os.path.join(HERE, "..", "..", "data")

with open(os.path.join(DATA_DIR, "ingredient_master.json"), encoding="utf-8") as f:
    MASTER = json.load(f)
with open(os.path.join(DATA_DIR, "incompatible_pairs.json"), encoding="utf-8") as f:
    INCOMPATIBLE = set(tuple(sorted(p)) for p in json.load(f))
with open(os.path.join(DATA_DIR, "alias_map.json"), encoding="utf-8") as f:
    ALIAS = json.load(f)


def std(name):
    return ALIAS.get((name or "").strip(), (name or "").strip())


# 体质/人群关键词 -> 用于匹配食材 contraindication
CONSTITUTION_KEYS = ["寒性体质", "虚寒体质", "热性体质", "上火", "湿热", "脾胃虚寒", "结石体质"]
PEOPLE_KEYS = ["孕妇", "老人", "儿童", "婴幼儿", "经期女性", "哺乳期", "高血压", "高血脂",
               "糖尿病", "痛风", "海鲜过敏", "牛奶过敏", "大豆过敏", "乳糖不耐受",
               "胆囊炎", "胃炎", "胃溃疡", "痔疮", "咽炎", "哮喘", "肾结石", "贫血", "体虚"]


def check_ingredients(ingredients, profile=None):
    """
    对一组食材做禁忌校验。
    :param ingredients: 食材名列表（可含别名）
    :param profile: {"allergies":[...], "constitution":[...], "people":[...], "avoid":[食材名]}
    :return: {risk_level, warnings:[...], suggestions:[...]}
    """
    profile = profile or {}
    allergies = set(profile.get("allergies", []))
    constitutions = set(profile.get("constitution", []))
    people = set(profile.get("people", []))
    avoid = set(std(x) for x in profile.get("avoid", []))

    ings = [std(x) for x in ingredients if std(x) in MASTER]
    warnings = []
    suggestions = []

    # 1) 过敏源
    for ing in ings:
        attr = MASTER[ing]
        if attr["allergen"] in allergies:
            warnings.append({"level": "高", "type": "过敏源",
                             "detail": f"含『{ing}』，属于用户过敏源「{attr['allergen']}」，应避免食用"})
            # 调整建议：找同组替代
            from reasoning.substitute import recommend
            alts = recommend(ing, user_allergies=allergies, dish_ingredients=ings, top_k=2)
            if alts:
                suggestions.append(f"『{ing}』可用 {alts[0]['name']} 替代（{alts[0]['reason']}）")
            else:
                suggestions.append(f"『{ing}』建议从菜谱中去除")

    # 2) 用户明确忌口食材
    for ing in ings:
        if ing in avoid:
            warnings.append({"level": "高", "type": "忌口食材",
                             "detail": f"含用户忌口食材『{ing}』"})
            from reasoning.substitute import recommend
            alts = recommend(ing, user_allergies=allergies, dish_ingredients=ings, top_k=2)
            if alts:
                suggestions.append(f"『{ing}』可用 {alts[0]['name']} 替代")

    # 3) 体质/人群禁忌（子串双向匹配，鲁棒）
    for ing in ings:
        attr = MASTER[ing]
        ci = attr.get("contraindication", {})
        for c in ci.get("constitution", []):
            if any(p in c or c in p for p in constitutions):
                warnings.append({"level": "中", "type": "体质禁忌",
                                 "detail": f"『{ing}』不适合{c}者，建议少食或替换"})
        for p in ci.get("people", []):
            if any(u in p or p in u for u in people):
                warnings.append({"level": "中", "type": "人群禁忌",
                                 "detail": f"『{ing}』对「{p}」需慎食/忌食"})

    # 4) 相克（食材两两）
    for i in range(len(ings)):
        for j in range(i + 1, len(ings)):
            if tuple(sorted((ings[i], ings[j]))) in INCOMPATIBLE:
                warnings.append({"level": "高", "type": "相克",
                                 "detail": f"『{ings[i]}』与『{ings[j]}』相克，不宜同食"})
                suggestions.append(f"『{ings[i]}』与『{ings[j]}』相克，建议去掉其一或分时食用")

    # 风险定级
    levels = [w["level"] for w in warnings]
    if "高" in levels:
        risk = "高"
    elif "中" in levels:
        risk = "中"
    elif warnings:
        risk = "低"
    else:
        risk = "无"

    return {"risk_level": risk, "warnings": warnings, "suggestions": suggestions,
            "checked_ingredients": ings}


def check_recipe(recipe, profile=None):
    """对结构化菜谱(含 ingredients 字段)做禁忌校验。"""
    ings = [i["name"] for i in recipe.get("ingredients", [])]
    # 仅取主食材/调料中的真实食材
    real = [n for n in ings if std(n) in MASTER]
    return check_ingredients(real, profile)


if __name__ == "__main__":
    # 示例：孕妇 + 海鲜过敏，吃"螃蟹炒年糕"
    r = check_ingredients(["螃蟹", "年糕", "葱", "姜"],
                          profile={"allergies": ["海鲜"], "people": ["孕妇"]})
    print("风险等级:", r["risk_level"])
    for w in r["warnings"]:
        print("  [{}] {}".format(w["level"], w["detail"]))
    for s in r["suggestions"]:
        print("  -> 建议:", s)
