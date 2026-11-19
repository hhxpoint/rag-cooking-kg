# -*- coding: utf-8 -*-
"""
推理层（一）：食材替代推理。
方法：基于领域规则 + 图谱属性相似度匹配，不依赖图神经网络/复杂算法。
  - 优先使用主表中"显式替代"关系（可解释、可靠）；
  - 否则在同替代分组内，按 口味/性味/功效/营养 的相似度打分推荐；
  - 结合用户过敏源与当前菜品其它食材，过滤不安全/相克的候选。
答辩表述："基于领域规则与属性相似度的知识推理"。
"""
import os
import json

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HERE, "..", "..", "data")

with open(os.path.join(DATA_DIR, "ingredient_master.json"), encoding="utf-8") as f:
    MASTER = json.load(f)
with open(os.path.join(DATA_DIR, "incompatible_pairs.json"), encoding="utf-8") as f:
    INCOMPATIBLE = set(tuple(sorted(p)) for p in json.load(f))
with open(os.path.join(DATA_DIR, "alias_map.json"), encoding="utf-8") as f:
    ALIAS = json.load(f)


def std(name):
    return ALIAS.get((name or "").strip(), (name or "").strip())


def _nature_distance(a, b):
    order = {"寒": 0, "凉": 1, "平": 2, "温": 3, "热": 4}
    if a not in order or b not in order:
        return 2
    return abs(order[a] - order[b])


def recommend(ingredient, user_allergies=None, dish_ingredients=None, top_k=5):
    """
    为 ingredient 推荐替代食材。
    :param ingredient: 待替代食材名（支持别名）
    :param user_allergies: 用户过敏源列表，如 ["蛋","奶","海鲜"]
    :param dish_ingredients: 当前菜品中的其它食材，用于排除相克候选
    :param top_k: 返回数量
    :return: [{name, score, reason}]
    """
    user_allergies = set(user_allergies or [])
    dish_ingredients = [std(x) for x in (dish_ingredients or [])]
    src = std(ingredient)
    if src not in MASTER:
        return []
    a = MASTER[src]
    candidates = []

    # 1) 显式替代（最高优先）
    for sub in a.get("substitutes", []):
        s = std(sub)
        if s not in MASTER or s == src:
            continue
        sa = MASTER[s]
        reason = f"与{src}同属「{a['group']}」，主表显式推荐替代"
        score = 100
        if sa["allergen"] in user_allergies:
            reason += "（但含用户过敏源，不推荐）"
            score -= 80
        candidates.append({"name": s, "score": score, "reason": reason, "explicit": True})

    # 2) 同组相似度候选
    for name, attr in MASTER.items():
        if name == src or attr["cat"] != "食材":
            continue
        if any(c["name"] == name for c in candidates):
            continue
        score = 0
        reasons = []
        if attr["group"] == a["group"]:
            score += 30
            reasons.append("同替代分组")
        f_overlap = set(attr["flavors"]) & set(a["flavors"])
        if f_overlap:
            score += 10 + 5 * len(f_overlap)
            reasons.append(f"口味相近({'/'.join(f_overlap)})")
        e_overlap = set(attr["effects"]) & set(a["effects"])
        if e_overlap:
            score += 8 * len(e_overlap)
            reasons.append(f"功效相近({'/'.join(e_overlap)})")
        n_overlap = set(attr["nutrition"]) & set(a["nutrition"])
        if n_overlap:
            score += 4 * len(n_overlap)
            reasons.append("营养构成相近")
        nd = _nature_distance(attr["nature"], a["nature"])
        if nd <= 1:
            score += 10 - 2 * nd
            reasons.append("性味相近")
        # 过敏过滤
        if attr["allergen"] in user_allergies:
            score -= 70
            reasons.append("含用户过敏源")
        # 相克过滤（与菜品内其它食材）
        clash = False
        for d in dish_ingredients:
            if d != name and tuple(sorted((name, d))) in INCOMPATIBLE:
                clash = True
                break
        if clash:
            score -= 60
            reasons.append("与菜品中其它食材相克")
        if score > 0:
            candidates.append({"name": name, "score": score,
                               "reason": "；".join(reasons) or "属性较接近",
                               "explicit": False})

    candidates.sort(key=lambda x: x["score"], reverse=True)
    return candidates[:top_k]


def substitute_in_recipe(recipe_ingredients, missing=None, user_allergies=None):
    """
    给定一份菜谱食材列表，针对其中缺失/忌口的食材给出替代方案。
    :param recipe_ingredients: 菜谱原食材列表（名）
    :param missing: 用户缺失/忌口的食材（需替代）
    :param user_allergies: 用户过敏源
    :return: [{original, alternatives:[...]}]
    """
    missing = [std(x) for x in (missing or recipe_ingredients)]
    out = []
    for m in missing:
        alts = recommend(m, user_allergies=user_allergies,
                         dish_ingredients=[x for x in recipe_ingredients if std(x) != m], top_k=3)
        out.append({"original": m, "alternatives": alts})
    return out


if __name__ == "__main__":
    print("牛奶 的替代（无过敏）:")
    for c in recommend("牛奶"):
        print("  ", c)
    print("\n牛奶 的替代（用户对奶过敏）:")
    for c in recommend("牛奶", user_allergies=["奶"]):
        print("  ", c)
    print("\n五花肉 的替代（无过敏）:")
    for c in recommend("五花肉"):
        print("  ", c)
