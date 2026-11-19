# -*- coding: utf-8 -*-
"""
抽取层 Schema：实体类型与关系类型定义。
依据课程设计方案写死，供 LLM Few-shot 抽取与规则抽取共用。
实体类型：菜谱 / 食材 / 调料 / 工序 / 口味 / 禁忌 / 功效 / 人群
关系类型：包含 / 属于 / 搭配 / 相克 / 适用 / 具有功效 / 采用工序
"""
from enum import Enum

# 实体类型
class EntityType(str, Enum):
    RECIPE = "菜谱"
    INGREDIENT = "食材"
    SEASONING = "调料"
    PROCESS = "工序"
    FLAVOR = "口味"
    TABOO = "禁忌"
    EFFECT = "功效"
    GROUP = "人群"

# 关系类型
class RelationType(str, Enum):
    CONTAINS = "包含"        # 菜谱 包含 食材/调料
    BELONGS_TO = "属于"      # 菜谱 属于 菜系/口味
    PAIRS_WITH = "搭配"      # 食材 搭配 食材
    INCOMPATIBLE = "相克"    # 食材 相克 食材
    SUITABLE_FOR = "适用"    # 菜谱/食材 适用 人群
    HAS_EFFECT = "具有功效"  # 菜谱/食材 具有功效 功效
    USES_PROCESS = "采用工序" # 菜谱 采用工序 工序
    SUBSTITUTE = "替代"      # 食材 替代 食材（推理层派生，亦可入库）

# 人类可读说明（用于报告与前端展示）
ENTITY_DESC = {
    EntityType.RECIPE.value: "一道具体的菜谱，如「番茄炒蛋」",
    EntityType.INGREDIENT.value: "主食材，如「番茄」「牛肉」",
    EntityType.SEASONING.value: "调味料，如「酱油」「辣椒」",
    EntityType.PROCESS.value: "烹饪工序，如「炒」「炖」「蒸」",
    EntityType.FLAVOR.value: "口味标签，如「咸鲜」「麻辣」",
    EntityType.TABOO.value: "饮食禁忌，如「海鲜过敏者忌」",
    EntityType.EFFECT.value: "食疗功效，如「补气血」「清热」",
    EntityType.GROUP.value: "适用人群，如「儿童」「老人」",
}

RELATION_DESC = {
    RelationType.CONTAINS.value: "菜谱由其食材/调料构成",
    RelationType.BELONGS_TO.value: "菜谱归属某一菜系/口味",
    RelationType.PAIRS_WITH.value: "两种食材常搭配使用",
    RelationType.INCOMPATIBLE.value: "两种食材相克，不宜同食",
    RelationType.SUITABLE_FOR.value: "菜谱/食材适合某类人群",
    RelationType.HAS_EFFECT.value: "菜谱/食材具有某食疗功效",
    RelationType.USES_PROCESS.value: "菜谱采用的烹饪工序",
    RelationType.SUBSTITUTE.value: "某食材可被另一食材替代",
}

# 抽取示例（Few-shot）：给大模型看的 2~3 个样例
FEW_SHOT_EXAMPLES = [
    {
        "text": "番茄炒蛋：番茄2个、鸡蛋3个，加食盐、食用油炒制，口味酸甜咸鲜，适合儿童。",
        "triples": [
            ["番茄炒蛋", "包含", "番茄"],
            ["番茄炒蛋", "包含", "鸡蛋"],
            ["番茄炒蛋", "包含", "食盐"],
            ["番茄炒蛋", "采用工序", "炒"],
            ["番茄炒蛋", "属于", "咸鲜"],
            ["番茄炒蛋", "适用", "儿童"],
        ],
    },
    {
        "text": "山药排骨汤：猪肉、山药炖煮，滋补健脾，适合老人与体虚者，脾胃虚寒者少食。",
        "triples": [
            ["山药排骨汤", "包含", "猪肉"],
            ["山药排骨汤", "包含", "山药"],
            ["山药排骨汤", "采用工序", "炖"],
            ["山药排骨汤", "具有功效", "健脾"],
            ["山药排骨汤", "适用", "老人"],
            ["山药排骨汤", "适用", "体虚者"],
        ],
    },
    {
        "text": "螃蟹性寒，与柿子相克，孕妇忌食螃蟹。",
        "triples": [
            ["螃蟹", "相克", "柿子"],
            ["螃蟹", "适用", "孕妇忌食"],
        ],
    },
]
