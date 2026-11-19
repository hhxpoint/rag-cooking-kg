# -*- coding: utf-8 -*-
"""
数据生成器：构建"小而精"的中式家常菜谱数据集。
设计原则（对标课程设计方案）：
  1. 不爬取海量数据，使用"精选种子菜谱 + 规则化组合扩展"的方式，
     产出 ~600 条字段完整、规范统一的菜谱，覆盖常见家常菜。
  2. 食材名统一（主表 + 别名映射），步骤完整，口味/功效/禁忌标签清晰，
     从源头降低清洗与实体消歧成本。
  3. 纯 Python 实现，零第三方依赖，可复现（固定随机种子）。
输出：
  data/raw_recipes.json   原始菜谱（结构化）
  data/recipes.csv        扁平化便于查看/导入
  data/alias_map.json     实体别名映射表（供抽取层消歧）
  data/ingredient_master.json  食材主表（含属性，供推理层使用）
"""
import json
import csv
import random
import os

SEED = 20261104
random.seed(SEED)
HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HERE, "..", "data")
os.makedirs(DATA_DIR, exist_ok=True)

# ---------------------------------------------------------------------------
# 1. 食材主表（含属性，供图谱与推理层使用）
#    属性说明：
#      cat       : 食材 / 调料
#      flavors   : 口味特征（咸/甜/酸/辣/鲜/香/清淡...）
#      nature    : 中医性味（寒/凉/平/温/热）
#      effects   : 功效标签
#      nutrition : 营养/成分标签（高蛋白/低脂肪/高纤维/富钙/维C/碳水...）
#      allergen  : 过敏源（蛋/奶/海鲜/坚果/大豆/麸质/无）
#      group     : 替代分组（同类可互替）
#      substitutes: 显式推荐替代（优先使用，保证可解释）
#      contraindication: 禁忌（constitution=体质, people=人群）
# ---------------------------------------------------------------------------
INGREDIENTS = {
    # ---------------- 蛋白质类 ----------------
    "鸡蛋": {"cat": "食材", "flavors": ["鲜", "香"], "nature": "平", "effects": ["补气血", "健脑"], "nutrition": ["高蛋白", "富含卵磷脂"], "allergen": "蛋", "group": "蛋奶", "substitutes": ["鸭蛋", "鹌鹑蛋"], "contraindication": {"constitution": [], "people": ["胆囊炎患者(少油煎)"]}},
    "鸭蛋": {"cat": "食材", "flavors": ["鲜", "香"], "nature": "凉", "effects": ["滋阴"], "nutrition": ["高蛋白"], "allergen": "蛋", "group": "蛋奶", "substitutes": ["鸡蛋", "鹌鹑蛋"], "contraindication": {"constitution": [], "people": []}},
    "鹌鹑蛋": {"cat": "食材", "flavors": ["鲜", "香"], "nature": "平", "effects": ["补气血"], "nutrition": ["高蛋白"], "allergen": "蛋", "group": "蛋奶", "substitutes": ["鸡蛋", "鸭蛋"], "contraindication": {"constitution": [], "people": []}},
    "猪肉": {"cat": "食材", "flavors": ["鲜", "香"], "nature": "平", "effects": ["补气血", "滋阴"], "nutrition": ["高蛋白", "高脂肪"], "allergen": "无", "group": "红肉", "substitutes": ["牛肉", "鸡肉"], "contraindication": {"constitution": [], "people": ["高血脂者(少)"]}},
    "五花肉": {"cat": "食材", "flavors": ["香", "鲜"], "nature": "平", "effects": ["补气血"], "nutrition": ["高脂肪", "高蛋白"], "allergen": "无", "group": "红肉", "substitutes": ["瘦肉", "鸡腿肉"], "contraindication": {"constitution": [], "people": ["肥胖/高血脂者(少)"]}},
    "瘦肉": {"cat": "食材", "flavors": ["鲜"], "nature": "平", "effects": ["补气血"], "nutrition": ["高蛋白", "低脂肪"], "allergen": "无", "group": "红肉", "substitutes": ["鸡胸肉", "牛肉"], "contraindication": {"constitution": [], "people": []}},
    "牛肉": {"cat": "食材", "flavors": ["鲜", "香"], "nature": "温", "effects": ["补气血", "强筋骨"], "nutrition": ["高蛋白", "富铁"], "allergen": "无", "group": "红肉", "substitutes": ["羊肉", "瘦肉"], "contraindication": {"constitution": ["热性体质"], "people": ["上火者(少)"]}},
    "羊肉": {"cat": "食材", "flavors": ["鲜", "膻"], "nature": "温", "effects": ["温补", "暖身"], "nutrition": ["高蛋白", "富铁"], "allergen": "无", "group": "红肉", "substitutes": ["牛肉", "鸡肉"], "contraindication": {"constitution": ["热性体质", "上火"], "people": ["发热者忌"]}},
    "鸡肉": {"cat": "食材", "flavors": ["鲜", "香"], "nature": "温", "effects": ["补气血", "健脾"], "nutrition": ["高蛋白", "低脂肪"], "allergen": "无", "group": "禽肉", "substitutes": ["鸭肉", "瘦肉"], "contraindication": {"constitution": [], "people": []}},
    "鸭肉": {"cat": "食材", "flavors": ["鲜"], "nature": "凉", "effects": ["滋阴", "清热"], "nutrition": ["高蛋白", "低脂肪"], "allergen": "无", "group": "禽肉", "substitutes": ["鸡肉", "瘦肉"], "contraindication": {"constitution": ["虚寒体质"], "people": []}},
    "鸡胸肉": {"cat": "食材", "flavors": ["鲜"], "nature": "平", "effects": ["补蛋白", "低脂"], "nutrition": ["高蛋白", "低脂肪"], "allergen": "无", "group": "禽肉", "substitutes": ["瘦肉", "鸡腿肉"], "contraindication": {"constitution": [], "people": []}},
    "鱼": {"cat": "食材", "flavors": ["鲜", "嫩"], "nature": "平", "effects": ["补蛋白", "益智"], "nutrition": ["高蛋白", "富DHA", "低脂肪"], "allergen": "海鲜", "group": "水产", "substitutes": ["虾", "豆腐"], "contraindication": {"constitution": [], "people": ["海鲜过敏者忌"]}},
    "虾": {"cat": "食材", "flavors": ["鲜", "甜"], "nature": "温", "effects": ["补钙", "壮阳"], "nutrition": ["高蛋白", "富钙"], "allergen": "海鲜", "group": "水产", "substitutes": ["鱼", "蟹"], "contraindication": {"constitution": [], "people": ["海鲜/哮喘过敏者忌"]}},
    "螃蟹": {"cat": "食材", "flavors": ["鲜", "甜"], "nature": "寒", "effects": ["清热", "补钙"], "nutrition": ["高蛋白", "富钙"], "allergen": "海鲜", "group": "水产", "substitutes": ["虾", "鱼"], "contraindication": {"constitution": ["脾胃虚寒", "寒性体质"], "people": ["孕妇忌", "经期女性忌", "海鲜过敏者忌"]}},
    "豆腐": {"cat": "食材", "flavors": ["淡", "鲜"], "nature": "凉", "effects": ["清热", "补钙", "健脾"], "nutrition": ["高蛋白", "富钙", "植物蛋白"], "allergen": "大豆", "group": "豆制品", "substitutes": ["豆干", "鸡蛋"], "contraindication": {"constitution": ["虚寒体质"], "people": ["痛风者(适量)"]}},
    "豆干": {"cat": "食材", "flavors": ["香", "咸"], "nature": "平", "effects": ["补钙", "补蛋白"], "nutrition": ["高蛋白", "富钙"], "allergen": "大豆", "group": "豆制品", "substitutes": ["豆腐", "千张"], "contraindication": {"constitution": [], "people": []}},
    "千张": {"cat": "食材", "flavors": ["淡", "香"], "nature": "平", "effects": ["补钙"], "nutrition": ["高蛋白", "富钙"], "allergen": "大豆", "group": "豆制品", "substitutes": ["豆干", "豆腐"], "contraindication": {"constitution": [], "people": []}},
    "牛奶": {"cat": "食材", "flavors": ["甜", "香"], "nature": "平", "effects": ["补钙", "安神"], "nutrition": ["富钙", "高蛋白"], "allergen": "奶", "group": "蛋奶", "substitutes": ["燕麦奶", "豆浆", "酸奶"], "contraindication": {"constitution": [], "people": ["乳糖不耐受者忌", "牛奶过敏者忌"]}},
    "酸奶": {"cat": "食材", "flavors": ["酸", "甜"], "nature": "平", "effects": ["助消化", "补钙"], "nutrition": ["富钙", "益生菌"], "allergen": "奶", "group": "蛋奶", "substitutes": ["牛奶", "燕麦奶"], "contraindication": {"constitution": [], "people": ["牛奶过敏者忌"]}},
    "燕麦奶": {"cat": "食材", "flavors": ["淡", "香"], "nature": "平", "effects": ["降血脂", "饱腹"], "nutrition": ["高纤维", "低敏"], "allergen": "无", "group": "蛋奶", "substitutes": ["牛奶", "豆浆"], "contraindication": {"constitution": [], "people": ["麸质过敏者(选无麸质)"]}},
    "豆浆": {"cat": "食材", "flavors": ["淡", "香"], "nature": "平", "effects": ["补蛋白", "健脾"], "nutrition": ["植物蛋白", "富钙(强化)"], "allergen": "大豆", "group": "豆制品", "substitutes": ["牛奶", "燕麦奶"], "contraindication": {"constitution": [], "people": ["大豆过敏者忌"]}},

    # ---------------- 蔬菜/菌菇/主食类 ----------------
    "番茄": {"cat": "食材", "flavors": ["酸甜", "鲜"], "nature": "凉", "effects": ["开胃", "富含维C"], "nutrition": ["富含维C", "低热量"], "allergen": "无", "group": "蔬菜", "substitutes": ["圣女果", "番茄酱"], "contraindication": {"constitution": ["虚寒体质(生食少)"], "people": []}},
    "土豆": {"cat": "食材", "flavors": ["淡", "面"], "nature": "平", "effects": ["饱腹", "健脾"], "nutrition": ["碳水", "高纤维"], "allergen": "无", "group": "根茎", "substitutes": ["山药", "红薯"], "contraindication": {"constitution": [], "people": ["糖尿病者(控量)"]}},
    "青椒": {"cat": "食材", "flavors": ["微辣", "清"], "nature": "温", "effects": ["开胃", "富含维C"], "nutrition": ["富含维C", "低热量"], "allergen": "无", "group": "蔬菜", "substitutes": ["红椒", "彩椒"], "contraindication": {"constitution": ["上火"], "people": []}},
    "红椒": {"cat": "食材", "flavors": ["甜", "微辣"], "nature": "温", "effects": ["开胃"], "nutrition": ["富含维C"], "allergen": "无", "group": "蔬菜", "substitutes": ["青椒", "彩椒"], "contraindication": {"constitution": ["上火"], "people": []}},
    "黄瓜": {"cat": "食材", "flavors": ["清淡", "脆"], "nature": "凉", "effects": ["清热", "利水"], "nutrition": ["低热量", "高水分"], "allergen": "无", "group": "蔬菜", "substitutes": ["冬瓜", "西葫芦"], "contraindication": {"constitution": ["脾胃虚寒"], "people": []}},
    "冬瓜": {"cat": "食材", "flavors": ["清淡"], "nature": "凉", "effects": ["清热", "利水消肿"], "nutrition": ["低热量", "高水分"], "allergen": "无", "group": "蔬菜", "substitutes": ["黄瓜", "丝瓜"], "contraindication": {"constitution": ["虚寒体质"], "people": []}},
    "茄子": {"cat": "食材", "flavors": ["吸味", "软"], "nature": "凉", "effects": ["清热", "活血"], "nutrition": ["高纤维", "低热量"], "allergen": "无", "group": "蔬菜", "substitutes": ["西葫芦", "青椒"], "contraindication": {"constitution": ["虚寒体质"], "people": ["手术前后(少)"]}},
    "白菜": {"cat": "食材", "flavors": ["清淡", "甜"], "nature": "凉", "effects": ["清热", "润肠"], "nutrition": ["高纤维", "低热量"], "allergen": "无", "group": "蔬菜", "substitutes": ["娃娃菜", "油菜"], "contraindication": {"constitution": ["虚寒体质"], "people": []}},
    "油菜": {"cat": "食材", "flavors": ["清淡"], "nature": "凉", "effects": ["补钙", "清热"], "nutrition": ["富钙", "高纤维"], "allergen": "无", "group": "蔬菜", "substitutes": ["小白菜", "菠菜"], "contraindication": {"constitution": ["虚寒体质"], "people": []}},
    "菠菜": {"cat": "食材", "flavors": ["清淡"], "nature": "凉", "effects": ["补铁", "润肠"], "nutrition": ["富铁", "高纤维", "富钙"], "allergen": "无", "group": "蔬菜", "substitutes": ["油麦菜", "油菜"], "contraindication": {"constitution": ["虚寒体质", "结石体质"], "people": ["肾结石者(焯水)"]}},
    "西兰花": {"cat": "食材", "flavors": ["清", "微苦"], "nature": "平", "effects": ["抗癌", "富含维C"], "nutrition": ["高纤维", "富含维C"], "allergen": "无", "group": "蔬菜", "substitutes": ["菜花", "芦笋"], "contraindication": {"constitution": [], "people": []}},
    "胡萝卜": {"cat": "食材", "flavors": ["甜", "脆"], "nature": "平", "effects": ["明目", "富β胡萝素"], "nutrition": ["富β胡萝素", "高纤维"], "allergen": "无", "group": "根茎", "substitutes": ["南瓜", "红薯"], "contraindication": {"constitution": [], "people": []}},
    "洋葱": {"cat": "食材", "flavors": ["辛", "甜"], "nature": "温", "effects": ["开胃", "降脂"], "nutrition": ["高纤维", "富硒"], "allergen": "无", "group": "蔬菜", "substitutes": ["大葱", "蒜"], "contraindication": {"constitution": ["上火", "眼疾"], "people": ["胃炎者(少)"]}},
    "香菇": {"cat": "食材", "flavors": ["鲜", "香"], "nature": "平", "effects": ["增强免疫", "降脂"], "nutrition": ["高纤维", "富多糖"], "allergen": "无", "group": "菌菇", "substitutes": ["平菇", "杏鲍菇"], "contraindication": {"constitution": [], "people": ["痛风者(适量)"]}},
    "金针菇": {"cat": "食材", "flavors": ["鲜", "脆"], "nature": "平", "effects": ["益智", "降脂"], "nutrition": ["高纤维", "高蛋白"], "allergen": "无", "group": "菌菇", "substitutes": ["香菇", "平菇"], "contraindication": {"constitution": [], "people": []}},
    "木耳": {"cat": "食材", "flavors": ["脆", "淡"], "nature": "平", "effects": ["清血管", "补血"], "nutrition": ["高纤维", "富铁"], "allergen": "无", "group": "菌菇", "substitutes": ["银耳", "香菇"], "contraindication": {"constitution": [], "people": ["出血性疾病者(少)"]}},
    "南瓜": {"cat": "食材", "flavors": ["甜", "面"], "nature": "温", "effects": ["养胃", "明目"], "nutrition": ["碳水", "富β胡萝素"], "allergen": "无", "group": "根茎", "substitutes": ["红薯", "胡萝卜"], "contraindication": {"constitution": [], "people": ["糖尿病者(控量)"]}},
    "山药": {"cat": "食材", "flavors": ["淡", "糯"], "nature": "平", "effects": ["健脾", "益肺"], "nutrition": ["碳水", "黏蛋白"], "allergen": "无", "group": "根茎", "substitutes": ["土豆", "芋头"], "contraindication": {"constitution": [], "people": []}},
    "莲藕": {"cat": "食材", "flavors": ["脆", "甜"], "nature": "凉", "effects": ["清热", "健脾"], "nutrition": ["碳水", "高纤维"], "allergen": "无", "group": "根茎", "substitutes": ["山药", "土豆"], "contraindication": {"constitution": ["虚寒体质"], "people": []}},
    "生菜": {"cat": "食材", "flavors": ["清淡", "脆"], "nature": "凉", "effects": ["清热", "低卡"], "nutrition": ["低热量", "高纤维"], "allergen": "无", "group": "蔬菜", "substitutes": ["油麦菜", "白菜"], "contraindication": {"constitution": ["虚寒体质"], "people": []}},
    "豆腐皮": {"cat": "食材", "flavors": ["香", "淡"], "nature": "平", "effects": ["补蛋白"], "nutrition": ["高蛋白", "富钙"], "allergen": "大豆", "group": "豆制品", "substitutes": ["千张", "豆干"], "contraindication": {"constitution": [], "people": []}},
    "年糕": {"cat": "食材", "flavors": ["淡", "糯"], "nature": "温", "effects": ["饱腹"], "nutrition": ["碳水"], "allergen": "麸质", "group": "主食", "substitutes": ["米饭", "面条"], "contraindication": {"constitution": [], "people": ["糖尿病者(控量)"]}},
    "米饭": {"cat": "食材", "flavors": ["淡"], "nature": "平", "effects": ["补气"], "nutrition": ["碳水"], "allergen": "无", "group": "主食", "substitutes": ["面条", "馒头"], "contraindication": {"constitution": [], "people": ["糖尿病者(控量)"]}},
    "面条": {"cat": "食材", "flavors": ["淡"], "nature": "温", "effects": ["饱腹"], "nutrition": ["碳水"], "allergen": "麸质", "group": "主食", "substitutes": ["米饭", "馒头"], "contraindication": {"constitution": [], "people": ["乳糜泻者忌", "糖尿病者(控量)"]}},

    # ---------------- 调料类 ----------------
    "食盐": {"cat": "调料", "flavors": ["咸"], "nature": "平", "effects": ["调味"], "nutrition": [], "allergen": "无", "group": "调料", "substitutes": ["酱油", "低钠盐"], "contraindication": {"constitution": [], "people": ["高血压者(少)"]}},
    "酱油": {"cat": "调料", "flavors": ["咸", "鲜", "酱香"], "nature": "平", "effects": ["调味", "上色"], "nutrition": [], "allergen": "大豆", "group": "调料", "substitutes": ["生抽", "老抽", "盐"], "contraindication": {"constitution": [], "people": ["高血压者(少)"]}},
    "生抽": {"cat": "调料", "flavors": ["咸", "鲜"], "nature": "平", "effects": ["调味"], "nutrition": [], "allergen": "大豆", "group": "调料", "substitutes": ["酱油", "味极鲜"], "contraindication": {"constitution": [], "people": ["高血压者(少)"]}},
    "老抽": {"cat": "调料", "flavors": ["咸", "酱香"], "nature": "平", "effects": ["上色"], "nutrition": [], "allergen": "大豆", "group": "调料", "substitutes": ["酱油", "生抽"], "contraindication": {"constitution": [], "people": []}},
    "糖": {"cat": "调料", "flavors": ["甜"], "nature": "平", "effects": ["调味", "提鲜"], "nutrition": ["碳水"], "allergen": "无", "group": "调料", "substitutes": ["冰糖", "蜂蜜"], "contraindication": {"constitution": ["湿热"], "people": ["糖尿病者忌"]}},
    "冰糖": {"cat": "调料", "flavors": ["甜"], "nature": "平", "effects": ["润肺", "调味"], "nutrition": ["碳水"], "allergen": "无", "group": "调料", "substitutes": ["糖", "蜂蜜"], "contraindication": {"constitution": ["湿热"], "people": ["糖尿病者忌"]}},
    "醋": {"cat": "调料", "flavors": ["酸"], "nature": "温", "effects": ["开胃", "解腻"], "nutrition": [], "allergen": "无", "group": "调料", "substitutes": ["米醋", "陈醋"], "contraindication": {"constitution": ["胃酸过多"], "people": ["胃溃疡者(少)"]}},
    "料酒": {"cat": "调料", "flavors": ["酒香"], "nature": "温", "effects": ["去腥"], "nutrition": [], "allergen": "无", "group": "调料", "substitutes": ["黄酒", "白酒"], "contraindication": {"constitution": [], "people": ["酒精过敏者忌", "孕妇忌"]}},
    "姜": {"cat": "调料", "flavors": ["辛", "辣"], "nature": "温", "effects": ["驱寒", "去腥"], "nutrition": [], "allergen": "无", "group": "调料", "substitutes": ["干姜", "蒜"], "contraindication": {"constitution": ["上火", "热性体质"], "people": ["痔疮者(少)"]}},
    "蒜": {"cat": "调料", "flavors": ["辛", "辣"], "nature": "温", "effects": ["杀菌", "开胃"], "nutrition": [], "allergen": "无", "group": "调料", "substitutes": ["蒜苗", "洋葱"], "contraindication": {"constitution": ["上火", "眼疾"], "people": ["胃炎者(少)"]}},
    "葱": {"cat": "调料", "flavors": ["辛", "香"], "nature": "温", "effects": ["驱寒", "调味"], "nutrition": [], "allergen": "无", "group": "调料", "substitutes": ["大葱", "洋葱"], "contraindication": {"constitution": ["上火"], "people": ["蜂蜜同食忌"]}},
    "辣椒": {"cat": "调料", "flavors": ["辣"], "nature": "热", "effects": ["开胃", "驱寒"], "nutrition": ["富含维C"], "allergen": "无", "group": "调料", "substitutes": ["干辣椒", "辣椒粉"], "contraindication": {"constitution": ["上火", "热性体质", "湿热"], "people": ["胃溃疡者忌", "痔疮者忌", "咽炎者忌"]}},
    "花椒": {"cat": "调料", "flavors": ["麻", "辣"], "nature": "温", "effects": ["驱寒", "调味"], "nutrition": [], "allergen": "无", "group": "调料", "substitutes": ["麻椒", "辣椒"], "contraindication": {"constitution": ["上火"], "people": ["胃炎者(少)"]}},
    "八角": {"cat": "调料", "flavors": ["香", "甘"], "nature": "温", "effects": ["调味", "理气"], "nutrition": [], "allergen": "无", "group": "调料", "substitutes": ["桂皮", "香叶"], "contraindication": {"constitution": [], "people": []}},
    "豆瓣酱": {"cat": "调料", "flavors": ["咸", "辣", "酱香"], "nature": "温", "effects": ["调味"], "nutrition": [], "allergen": "大豆", "group": "调料", "substitutes": ["黄豆酱", "辣酱"], "contraindication": {"constitution": ["上火"], "people": ["高血压者(少)"]}},
    "食用油": {"cat": "调料", "flavors": ["香"], "nature": "平", "effects": ["传热"], "nutrition": ["脂肪"], "allergen": "无", "group": "调料", "substitutes": ["花生油", "菜籽油"], "contraindication": {"constitution": [], "people": ["高血脂者(少)"]}},
    "蚝油": {"cat": "调料", "flavors": ["鲜", "咸"], "nature": "平", "effects": ["提鲜"], "nutrition": [], "allergen": "海鲜", "group": "调料", "substitutes": ["鲜酱油", "鸡精"], "contraindication": {"constitution": [], "people": ["海鲜过敏者忌"]}},
    "蜂蜜": {"cat": "调料", "flavors": ["甜"], "nature": "平", "effects": ["润肺", "通便"], "nutrition": ["碳水"], "allergen": "无", "group": "调料", "substitutes": ["冰糖", "糖"], "contraindication": {"constitution": ["湿热", "婴儿"], "people": ["1岁以下婴儿忌", "糖尿病者忌"]}},
}

# 别名映射表（供抽取层实体消歧）： 别名 -> 标准名
ALIAS_MAP = {
    "西红柿": "番茄", "洋柿子": "番茄",
    "洋芋": "土豆", "马铃薯": "土豆",
    "圆白菜": "白菜", "大白菜": "白菜",
    "小葱": "葱", "大葱": "葱", "香葱": "葱",
    "生姜": "姜", "老姜": "姜",
    "大蒜": "蒜", "蒜头": "蒜",
    "番薯": "红薯", "地瓜": "红薯",
    "胡罗卜": "胡萝卜", "红萝卜": "胡萝卜",
    "金针菜": "金针菇",
    "木须肉用木耳": "木耳",
    "生猪肉": "猪肉", "胛心肉": "瘦肉",
    "牛里脊": "牛肉", "牛腩": "牛肉",
    "鸡腿": "鸡肉", "鸡翅": "鸡肉",
    "河虾": "虾", "基围虾": "虾",
    "黄豆芽": "豆芽", "绿豆芽": "豆芽",
    "嫩豆腐": "豆腐", "老豆腐": "豆腐",
    "味极鲜": "生抽", "豉油": "生抽",
    "花生油": "食用油", "菜籽油": "食用油", "色拉油": "食用油",
    "白砂糖": "糖", "砂糖": "糖",
    "粟米": "玉米", "苞米": "玉米",
}

# ---------------------------------------------------------------------------
# 2. 禁忌相克关系（食材两两相克，供禁忌推理使用）
#    采用民间与营养学常见配伍禁忌（课程设计层面合理即可）
# ---------------------------------------------------------------------------
INCOMPATIBLE_PAIRS = [
    ("蜂蜜", "葱"), ("蜂蜜", "豆腐"), ("螃蟹", "柿子"), ("螃蟹", "梨"),
    ("牛奶", "橘子"), ("牛奶", "柿子"), ("鸡蛋", "豆浆"), ("菠菜", "豆腐"),
    ("羊肉", "西瓜"), ("牛肉", "栗子"), ("鲤鱼", "甘草"), ("兔肉", "芹菜"),
    ("虾", "维生素C片"), ("狗肉", "绿豆"), ("柿子", "螃蟹"), ("豆腐", "蜂蜜"),
]
INCOMPATIBLE_SET = set(tuple(sorted(p)) for p in INCOMPATIBLE_PAIRS)

# ---------------------------------------------------------------------------
# 3. 烹饪方法模板
# ---------------------------------------------------------------------------
METHODS = {
    "炒": {"nature": "快手", "base_steps": [
        "将{main}洗净切好，用少许{season1}腌制入味。",
        "热锅冷油，下入{aromatics}爆香。",
        "大火下{main}快速翻炒至变色。",
        "加入{aux}继续翻炒，调入{seasonings}。",
        "出锅前淋少许{season2}，装盘即可。",
    ]},
    "炖": {"nature": "慢火", "base_steps": [
        "将{main}焯水去血沫，捞出备用。",
        "砂锅中加足量清水，放入{main}与{aromatics}。",
        "大火烧开转小火炖煮约{cook}分钟。",
        "加入{aux}与{seasonings}继续炖至软烂。",
        "尝味后出锅，汤鲜肉烂。",
    ]},
    "蒸": {"nature": "清淡", "base_steps": [
        "将{main}处理干净，铺于盘中。",
        "{aux}切配好铺在{main}上。",
        "水开后上锅蒸约{cook}分钟。",
        "出锅淋热油，浇上{seasonings}调成的料汁。",
    ]},
    "煮": {"nature": "汤水", "base_steps": [
        "锅中加水烧开，放入{main}与{aromatics}。",
        "煮约{cook}分钟后下{aux}。",
        "调入{seasonings}，煮至入味。",
        "盛出撒少许葱花即可。",
    ]},
    "烧": {"nature": "入味", "base_steps": [
        "将{main}煎/炸至表面金黄。",
        "加入{aromatics}与{seasonings}炒出香味。",
        "加少许热水，小火烧约{cook}分钟收汁。",
        "下{aux}翻匀，大火收汁出锅。",
    ]},
    "拌": {"nature": "冷食", "base_steps": [
        "将{main}与{aux}焯熟或生切处理。",
        "过凉水后沥干。",
        "加入{aromatics}与{seasonings}拌匀。",
        "静置入味后装盘。",
    ]},
    "煎": {"nature": "香脆", "base_steps": [
        "将{main}用{season1}稍腌。",
        "平底锅少油，中小火将{main}煎至两面金黄。",
        "加入{aux}与{seasonings}略煎入味。",
        "出锅装盘。",
    ]},
    "焖": {"nature": "软烂", "base_steps": [
        "将{main}炒香上色。",
        "加入{aromatics}、{seasonings}与少量水。",
        "盖盖中小火焖约{cook}分钟。",
        "下{aux}收汁，软烂入味后出锅。",
    ]},
    "煲汤": {"nature": "滋补", "base_steps": [
        "将{main}焯水，与{aromatics}一同入砂锅。",
        "加足水大火煮沸转小火煲{cook}分钟。",
        "下{aux}再煲30分钟。",
        "最后以{seasonings}调味即可。",
    ]},
}

# 方法适用的主材类型权重（提高组合自然度）
METHOD_MAIN_PREF = {
    "炒": ["蔬菜", "禽肉", "蛋奶", "豆制品", "红肉"],
    "炖": ["红肉", "水产", "根茎"],
    "蒸": ["水产", "蛋奶", "禽肉", "豆制品"],
    "煮": ["水产", "主食", "蔬菜", "豆制品"],
    "烧": ["红肉", "水产", "根茎"],
    "拌": ["蔬菜", "豆制品", "蛋奶"],
    "煎": ["禽肉", "水产", "蛋奶", "红肉"],
    "焖": ["红肉", "根茎", "禽肉"],
    "煲汤": ["水产", "禽肉", "红肉", "根茎"],
}

# 分组 -> 食材列表（用于组合生成）
GROUP_MEMBERS = {}
for _name, _attr in INGREDIENTS.items():
    if _attr["cat"] == "食材":
        GROUP_MEMBERS.setdefault(_attr["group"], []).append(_name)

SEASONINGS_COMMON = ["食用油", "食盐", "生抽", "料酒", "姜", "蒜", "葱"]
AROMATICS = ["姜", "蒜", "葱"]
FLAVOR_BY_METHOD = {
    "炒": ["咸鲜", "家常"], "炖": ["鲜香", "浓郁"], "蒸": ["清淡", "鲜"],
    "煮": ["清淡", "鲜"], "烧": ["酱香", "咸鲜"], "拌": ["清爽", "酸辣"],
    "煎": ["香", "咸鲜"], "焖": ["浓郁", "咸鲜"], "煲汤": ["鲜香", "滋补"],
}

# 精选种子菜谱（真实常见菜，保证辨识度）
SEED_DISHES = [
    ("番茄炒蛋", "炒", ["番茄", "鸡蛋"], ["青椒"], ["咸鲜", "酸甜"], ["开胃", "补钙"], ["儿童", "一般人群"]),
    ("红烧肉", "烧", ["五花肉"], ["土豆"], ["酱香", "咸鲜"], ["补气血"], ["一般人群"]),
    ("宫保鸡丁", "炒", ["鸡胸肉"], ["黄瓜", "胡萝卜"], ["麻辣", "咸鲜"], ["开胃"], ["一般人群"]),
    ("清蒸鱼", "蒸", ["鱼"], ["葱"], ["清淡", "鲜"], ["补蛋白", "益智"], ["儿童", "老人"]),
    ("麻婆豆腐", "烧", ["豆腐"], ["牛肉"], ["麻辣", "咸鲜"], ["补钙", "补蛋白"], ["一般人群"]),
    ("土豆炖牛肉", "炖", ["牛肉"], ["土豆", "胡萝卜"], ["鲜香", "浓郁"], ["补气血", "强筋骨"], ["一般人群"]),
    ("凉拌黄瓜", "拌", ["黄瓜"], [], ["清爽", "酸辣"], ["清热", "低卡"], ["一般人群"]),
    ("西红柿鸡蛋汤", "煮", ["番茄", "鸡蛋"], ["葱"], ["清淡", "鲜"], ["开胃", "补钙"], ["儿童", "一般人群"]),
    ("青椒炒肉丝", "炒", ["瘦肉"], ["青椒"], ["咸鲜", "家常"], ["补气血"], ["一般人群"]),
    ("山药排骨汤", "煲汤", ["猪肉"], ["山药"], ["鲜香", "滋补"], ["健脾", "补气血"], ["老人", "体虚者"]),
    ("蒜蓉西兰花", "炒", ["西兰花"], [], ["清淡", "咸鲜"], ["抗癌", "富含维C"], ["一般人群"]),
    ("香菇炒青菜", "炒", ["香菇"], ["油菜"], ["鲜", "清淡"], ["增强免疫"], ["一般人群"]),
    ("冬瓜排骨汤", "煲汤", ["猪肉"], ["冬瓜"], ["清淡", "滋补"], ["清热", "利水消肿"], ["一般人群"]),
    ("辣子鸡", "炒", ["鸡肉"], ["辣椒"], ["麻辣", "香"], ["开胃", "温补"], ["一般人群"]),
    ("糖醋里脊", "煎", ["瘦肉"], [], ["酸甜", "香"], ["补气血"], ["儿童", "一般人群"]),
    ("清炒菠菜", "炒", ["菠菜"], [], ["清淡"], ["补铁", "润肠"], ["一般人群"]),
    ("南瓜粥", "煮", ["南瓜"], ["米饭"], ["甜", "清淡"], ["养胃", "明目"], ["老人", "儿童"]),
    ("葱爆羊肉", "炒", ["羊肉"], ["葱"], ["咸鲜", "香"], ["温补", "暖身"], ["一般人群"]),
    ("蟹炒年糕", "炒", ["螃蟹"], ["年糕"], ["鲜", "咸鲜"], ["补钙", "清热"], ["一般人群"]),
    ("木耳炒山药", "炒", ["木耳"], ["山药"], ["清淡", "脆"], ["清血管", "健脾"], ["老人", "一般人群"]),
    ("酸辣土豆丝", "炒", ["土豆"], ["辣椒", "醋"], ["酸辣", "脆"], ["开胃"], ["一般人群"]),
    ("鲫鱼豆腐汤", "煲汤", ["鱼"], ["豆腐"], ["鲜香", "滋补"], ["补钙", "补蛋白"], ["老人", "儿童"]),
    ("蚝油生菜", "炒", ["生菜"], [], ["咸鲜", "鲜"], ["清热", "低卡"], ["一般人群"]),
    ("红烧茄子", "烧", ["茄子"], [], ["酱香", "咸鲜"], ["清热", "活血"], ["一般人群"]),
    ("白灼虾", "煮", ["虾"], ["葱"], ["鲜", "甜"], ["补钙", "壮阳"], ["一般人群"]),
    ("鸡蛋羹", "蒸", ["鸡蛋"], [], ["清淡", "鲜"], ["补气血", "健脑"], ["婴幼儿", "老人"]),
    ("牛肉面", "煮", ["牛肉"], ["面条"], ["咸鲜", "香"], ["补气血", "饱腹"], ["一般人群"]),
    ("皮蛋瘦肉粥", "煮", ["瘦肉"], ["米饭"], ["咸鲜", "香"], ["补气血"], ["一般人群"]),
    ("凉拌木耳", "拌", ["木耳"], ["黄瓜"], ["清爽", "酸辣"], ["清血管", "补血"], ["一般人群"]),
    ("烤鸡翅", "煎", ["鸡肉"], [], ["香", "咸鲜"], ["补蛋白"], ["儿童", "一般人群"]),
]

CUISINES = ["家常菜", "粤菜", "川菜", "鲁菜", "苏菜", "浙菜", "湘菜", "闽菜", "徽菜", "汤羹", "主食", "凉菜"]
DIFFICULTIES = ["简单", "中等", "稍难"]
EFFECT_GLOBAL = ["补气血", "健脾", "清热", "开胃", "补钙", "润肺", "养颜", "助眠", "增强免疫", "降脂", "明目", "温补"]


def pick(seq, n=None):
    if n is None:
        return random.choice(seq)
    return random.sample(seq, min(n, len(seq)))


def derive_flavor_effect(ingredients):
    flavors = set()
    effects = set()
    for ing in ingredients:
        a = INGREDIENTS.get(ing)
        if not a:
            continue
        for f in a["flavors"]:
            flavors.add(f)
        for e in a["effects"]:
            effects.add(e)
    return list(flavors), list(effects)


def build_taboo(ingredients):
    taboos = []
    names = set(ingredients)
    # 相克
    for a in ingredients:
        for b in ingredients:
            if a >= b:
                continue
            if tuple(sorted((a, b))) in INCOMPATIBLE_SET:
                taboos.append({"type": "相克", "detail": f"{a}与{b}相克，不宜同食"})
    # 忌口人群 / 体质
    people_set, constitution_set = set(), set()
    for ing in ingredients:
        a = INGREDIENTS.get(ing)
        if not a:
            continue
        ci = a["contraindication"]
        for p in ci.get("people", []):
            people_set.add(p)
        for c in ci.get("constitution", []):
            constitution_set.add(c)
    for p in sorted(people_set):
        taboos.append({"type": "忌口人群", "detail": f"{p}慎食/忌食"})
    for c in sorted(constitution_set):
        taboos.append({"type": "体质禁忌", "detail": f"{c}者少食"})
    return taboos


def gen_steps(method, main, aux, seasonings, cook_time):
    tpl = METHODS[method]["base_steps"]
    arom = pick(AROMATICS)
    season1 = pick(["食盐", "生抽", "料酒"])
    season2 = pick(["香油", "生抽", "醋"])
    seas = "、".join(seasonings[:3]) if seasonings else "食盐"
    text = []
    for step in tpl:
        s = step.format(main=main, aux="、".join(aux) if aux else "辅料",
                        aromatics=arom, seasonings=seas, season1=season1,
                        season2=season2, cook=cook_time)
        text.append(s)
    return text


def make_recipe(rid, name, method, mains, auxs, flavors, effects, suitable):
    all_ings = list(mains) + list(auxs)
    # 统一食材名（去别名）
    norm = [ALIAS_MAP.get(i, i) for i in all_ings]
    norm = [i for i in norm if i in INGREDIENTS]
    # 调料
    seasonings = list(SEASONINGS_COMMON)
    if method in ("拌",):
        seasonings += ["醋", "辣椒"]
    if method in ("烧", "炖", "焖"):
        seasonings += ["糖", "八角"]
    if "海鲜" in [INGREDIENTS.get(i, {}).get("allergen") for i in norm]:
        seasonings += ["蚝油"]
    # 主食材定量
    ingredients = []
    amounts = {"食材": ["适量", "200克", "300克", "2个", "1根", "1块", "半斤"],
               "调料": ["适量", "1勺", "少许", "2勺"]}
    for i in norm:
        a = INGREDIENTS[i]
        ingredients.append({"name": i, "amount": pick(amounts[a["cat"]]), "category": a["cat"]})
    for s in seasonings:
        if s in INGREDIENTS:
            ingredients.append({"name": s, "amount": pick(amounts["调料"]), "category": "调料"})
    # 口味/功效（合并方法默认 + 食材推导）
    f_ing, e_ing = derive_flavor_effect(norm)
    flavors_final = list(dict.fromkeys(list(flavors) + FLAVOR_BY_METHOD.get(method, []) + f_ing))
    effects_final = list(dict.fromkeys(list(effects) + e_ing))
    cook_time = pick([10, 12, 15, 18, 20, 25, 30, 40, 45, 60]) if method not in ("蒸", "煮", "拌") else pick([8, 10, 12, 15])
    steps = gen_steps(method, "、".join(mains), auxs, seasonings, cook_time)
    desc = f"这是一道{flavors_final[0] if flavors_final else '家常'}口味的{name}，采用{method}的烹饪方式，" \
           f"主要食材为{('、'.join(mains))}，具有{('、'.join(effects_final[:3]) if effects_final else '家常')}的特点，" \
           f"适合{('、'.join(suitable))}食用。"
    return {
        "id": rid,
        "name": name,
        "cuisine": pick(CUISINES),
        "difficulty": pick(DIFFICULTIES),
        "cooking_time": cook_time,
        "servings": pick([1, 2, 2, 3, 4]),
        "ingredients": ingredients,
        "steps": steps,
        "flavors": flavors_final[:4],
        "effects": effects_final[:5],
        "suitable_for": suitable,
        "taboos": build_taboo(norm),
        "description": desc,
    }


def combo_name(method, mains, auxs):
    mains_s = "、".join(mains)
    aux_s = "、".join(auxs)
    if auxs:
        return f"{aux_s}{method}{mains_s}"
    return f"{mains_s}{method}"


def generate():
    recipes = []
    # 1) 种子菜谱
    for i, (name, method, mains, auxs, flavors, effects, suitable) in enumerate(SEED_DISHES, 1):
        recipes.append(make_recipe(f"R{ i:04d}", name, method, mains, auxs, flavors, effects, suitable))

    # 2) 组合扩展至 ~600
    target = 600
    attempts = 0
    while len(recipes) < target and attempts < 20000:
        attempts += 1
        method = pick(list(METHODS.keys()))
        pref = METHOD_MAIN_PREF[method]
        main_group = pick(pref)
        if main_group not in GROUP_MEMBERS or not GROUP_MEMBERS[main_group]:
            continue
        main = pick(GROUP_MEMBERS[main_group])
        # 辅料：蔬菜/根茎/豆制品
        aux_pool = GROUP_MEMBERS.get("蔬菜", []) + GROUP_MEMBERS.get("根茎", []) + GROUP_MEMBERS.get("豆制品", [])
        auxn = pick([0, 1, 1, 2])
        auxs = pick(aux_pool, auxn)
        # 去重主辅
        if main in auxs:
            continue
        # 避免主材本身是调料组
        flavor_def, effect_def = derive_flavor_effect([main] + auxs)
        flavors = list(dict.fromkeys(FLAVOR_BY_METHOD.get(method, []) + flavor_def))[:3]
        effects = list(dict.fromkeys(effect_def))[:3] or ["家常"]
        suitable = pick(["一般人群", "儿童", "老人", "上班族", "体虚者", "孕妇(适量)"], 2)
        name = combo_name(method, [main], auxs)
        # 去重菜名
        if any(r["name"] == name for r in recipes):
            continue
        rid = f"R{len(recipes)+1:04d}"
        recipes.append(make_recipe(rid, name, method, [main], auxs, flavors, effects, suitable))
        if len(recipes) % 50 == 0:
            pass

    return recipes


def main():
    recipes = generate()
    # 写 JSON
    out_json = os.path.join(DATA_DIR, "raw_recipes.json")
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(recipes, f, ensure_ascii=False, indent=2)
    # 写 CSV（扁平）
    out_csv = os.path.join(DATA_DIR, "recipes.csv")
    with open(out_csv, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["id", "name", "cuisine", "difficulty", "cooking_time", "servings",
                    "ingredients", "steps", "flavors", "effects", "suitable_for", "taboos", "description"])
        for r in recipes:
            w.writerow([
                r["id"], r["name"], r["cuisine"], r["difficulty"], r["cooking_time"], r["servings"],
                "；".join(f"{i['name']}{i['amount']}" for i in r["ingredients"]),
                " | ".join(r["steps"]),
                "、".join(r["flavors"]), "、".join(r["effects"]), "、".join(r["suitable_for"]),
                "；".join(t["detail"] for t in r["taboos"]), r["description"],
            ])
    # 写食材主表
    out_ing = os.path.join(DATA_DIR, "ingredient_master.json")
    with open(out_ing, "w", encoding="utf-8") as f:
        json.dump(INGREDIENTS, f, ensure_ascii=False, indent=2)
    # 写别名映射
    out_alias = os.path.join(DATA_DIR, "alias_map.json")
    with open(out_alias, "w", encoding="utf-8") as f:
        json.dump(ALIAS_MAP, f, ensure_ascii=False, indent=2)
    # 写禁忌相克对
    out_incompat = os.path.join(DATA_DIR, "incompatible_pairs.json")
    with open(out_incompat, "w", encoding="utf-8") as f:
        json.dump(INCOMPATIBLE_PAIRS, f, ensure_ascii=False, indent=2)

    print(f"生成菜谱数量: {len(recipes)}")
    print(f"食材主表条目: {len(INGREDIENTS)}")
    print(f"别名映射条目: {len(ALIAS_MAP)}")
    print(f"相克关系对数: {len(INCOMPATIBLE_PAIRS)}")
    print(f"输出目录: {os.path.abspath(DATA_DIR)}")


if __name__ == "__main__":
    main()
