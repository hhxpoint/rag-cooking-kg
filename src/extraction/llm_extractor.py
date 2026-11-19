# -*- coding: utf-8 -*-
"""
LLM Few-shot 抽取器（可插拔）：
  - build_prompt(text)        ：拼装"固定 prompt 模板 + 2~3 个抽取示例"的抽取提示词
  - extract_with_llm(text)    ：调用可插拔的大模型（OpenAI 兼容 / 本地 Ollama）；
                                 未配置 API 时自动回退到规则抽取，保证可运行
  - extract(text)             ：统一入口，优先 LLM，失败/未配置则离线规则抽取

设计目标（对标方案）：不训练 NER，直接调用大模型做 Few-shot 抽取；
后端可接开源本地模型或 API，开发与部署成本极低。
"""
import os
import json
from extraction.schema import FEW_SHOT_EXAMPLES, EntityType, RelationType, ENTITY_DESC, RELATION_DESC
from extraction.rule_extractor import from_text

# ---- 可插拔大模型配置（通过环境变量或 config.py 注入） ----
LLM_BASE_URL = os.environ.get("RAG_LLM_BASE_URL", "")      # 如 https://api.openai.com/v1
LLM_API_KEY = os.environ.get("RAG_LLM_API_KEY", "")
LLM_MODEL = os.environ.get("RAG_LLM_MODEL", "gpt-3.5-turbo")

SYSTEM_PROMPT = (
    "你是一个中文菜谱知识抽取引擎。请从给定文本中抽取(头实体, 关系, 尾实体)三元组。\n"
    f"允许的实体类型：{', '.join(e.value for e in EntityType)}。\n"
    f"允许的关系类型：{', '.join(r.value for r in RelationType)}。\n"
    "只输出 JSON 数组，元素形如 [\"头实体\",\"关系\",\"尾实体\"]，不要输出多余说明。\n"
    "注意：食材名请统一为标准名（如'西红柿'统一为'番茄'）。"
)


def build_prompt(text: str) -> str:
    """拼装 Few-shot 抽取提示词。"""
    shots = ""
    for ex in FEW_SHOT_EXAMPLES:
        shots += f"文本：{ex['text']}\n输出：{json.dumps(ex['triples'], ensure_ascii=False)}\n"
    return f"{SYSTEM_PROMPT}\n\n{shots}\n文本：{text}\n输出："


def _call_openai_compatible(prompt: str) -> str:
    """调用 OpenAI 兼容接口（需要 requests）。"""
    import requests
    resp = requests.post(
        f"{LLM_BASE_URL}/chat/completions",
        headers={"Authorization": f"Bearer {LLM_API_KEY}", "Content-Type": "application/json"},
        json={"model": LLM_MODEL, "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": prompt},
        ]}, timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"]


def _parse_llm_output(out: str):
    """解析 LLM 输出为三元组列表。"""
    try:
        start = out.find("[")
        end = out.rfind("]") + 1
        arr = json.loads(out[start:end])
        return [t for t in arr if isinstance(t, list) and len(t) == 3]
    except Exception:
        return []


def extract_with_llm(text: str) -> list:
    """真正调用大模型做抽取；未配置则抛异常交由回退处理。"""
    if not (LLM_BASE_URL and LLM_API_KEY):
        raise RuntimeError("未配置 LLM（RAG_LLM_BASE_URL / RAG_LLM_API_KEY），回退规则抽取")
    prompt = build_prompt(text)
    out = _call_openai_compatible(prompt)
    return _parse_llm_output(out)


def extract(text: str, prefer_llm: bool = True) -> list:
    """统一抽取入口。"""
    if prefer_llm:
        try:
            return extract_with_llm(text)
        except Exception:
            pass
    # 离线回退：规则/词典抽取（产出与 LLM 同构的三元组）
    return from_text(text)


if __name__ == "__main__":
    txt = "麻婆豆腐用豆腐和牛肉烧制，麻辣鲜香，适合一般人群；豆腐性凉，痛风者适量。"
    print("Prompt 预览（前300字）：\n", build_prompt(txt)[:300])
    print("\n实际抽取结果（离线回退）：")
    for t in extract(txt):
        print("  ", t)
