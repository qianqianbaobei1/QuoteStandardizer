# -*- coding: utf-8 -*-
"""
DeepSeek 报价单表格理解与物料清洗
1. 表头/列角色推导 (任意行列、多级表头、左右并排子表)
2. 整行字段交给模型清洗为：标准名称、结构化参数、单位
3. 分类编码只从标准编码库 (CATEGORY_CODE_FILE) 中选；库中没有时模型可给出，但会标记为未经核验
4. 不调用模型时 (未配置或调用失败) 走规则降级，结果标记待复核
"""
import re
import json
import time
import requests
from concurrent.futures import ThreadPoolExecutor

from config import (DEEPSEEK_API_KEY, DEEPSEEK_BASE_URL, DEEPSEEK_MODEL, LLM_MAX_RETRIES, LLM_TEMPERATURE,
                    CATEGORY_CODE_MAP, CODE_LIBRARY_MAX_LINES, JEV_ROW_BATCH,
                    TYPESAFE_API_KEY, TYPESAFE_BASE_URL, TYPESAFE_MODEL,
                    GEMINI_API_KEY, GEMINI_MODEL)
from gemini_client import GeminiClient

# 表头关键词：单价类优先于合价/金额类，避免把“合价=数量×单价”当成单价
UNIT_PRICE_KEYS = ['单价', '面价', '出厂价', '市场价', '预算价', '含税价', '报价', '价格']
TOTAL_PRICE_KEYS = ['合价', '总价', '金额', '小计', '合计']
UNIT_KEYS = ['计量单位', '单位']
NAME_KEYS = ['名称', '品名', '产品', '材料', '设备']
# 只注入前 N 条：整个库逐批重发会成倍放大 token 成本，且模型实际无法逐条比对上千项
_code_items = list(CATEGORY_CODE_MAP.items())[:max(0, CODE_LIBRARY_MAX_LINES)]
CODE_LIBRARY_TEXT = "；".join(f"{name}={code}" for name, code in _code_items)
CODE_LIBRARY_TRUNCATED = len(CATEGORY_CODE_MAP) > len(_code_items)
JUNK_SPEC_KEYS = {'其他关键属性', '供应商', '品牌', '价格', '单价', '单位', '铝锭基价', '基价', '铜基价', '基价/吨', '铝锭基价/吨', '表标题', '指导价'}
IGNORE_SPEC_KEYS = ['单价', '价格', '序号', '图片', '金额', '合价', '总价', '小计', '数量', '备注', '单位']

SYSTEM_PROMPT_LAYOUT_DYNAMIC = """你是报价单表格结构分析员。
请分析用户提供的报价表前几行样本，判断表格的布局与表头结构。

【核心任务】：
1. 识别真实表头所在的行号 header_row (1-based) 与数据开始行号 data_start_row (1-based)；
2. 识别是否包含左右并排的双子表（如左边为某产品，右边为另一产品），若有请拆分为多个 sub_tables；
3. 动态标定该表中“单价列”、“单位列”，其余所有列（无论是品名、型号、规格、材质、工艺、颜色、备注等）均视为属性源；
4. 列名因厂商而异，请按含义判断，不要局限于固定的列名。

【返回格式要求】：严格返回 JSON 对象，不得输出任何 Markdown 以外的冗余字符：
{
  "header_row": 2,
  "data_start_row": 3,
  "is_side_by_side": false,
  "sub_tables": [
    {
      "table_name": "主表",
      "col_start": 1,
      "col_end": 15,
      "price_col": 6,       // 核心单价列号 (1-based)
      "unit_col": 7,        // 计量单位列号 (1-based, 若无填 null)
      "key_columns_desc": "第2列为品名，第3列为型号，第4列为截面，第5列为材质工艺..."
    }
  ]
}
"""

SYSTEM_PROMPT_BATCH_DYNAMIC = """你是一个国家级建设工程招投标与建材物料领域的专家。请深度理解每一行原始报价数据的真实业务含义，判断每个单元格应该填什么，杜绝片面硬拆！

【核心行业业务规则】：
1. standard_name (名称规格): 只能是纯净的物料品类名称（如：铝板、暗杆闸阀、明杆闸阀、涡轮蝶阀、真石漆、木门等）。严禁将一堆规格型号、尺寸、单价混堆在品名中！
2. structured_specs (规格参数): 必须且只能包含物料本身的【物理技术参数】（如材质/牌号、厚度、公称通径DN/De、系列型号、表面喷涂工艺等）。
   - 【严禁污染规格】：严禁将“基价”（如铝锭基价18000/吨、铜基价）、“指导价”、“调价依据”、“单价/总价/金额”、“品牌/厂家”、“税率/运费”、“表标题”等混入 structured_specs！
   - 若原始数据含有“铝锭基价/吨: 18000”或“基价”，这是价格联动核算依据，绝不是规格型号，必须从规格中剥离！
3. unit (计量单位): 必须是真实的工程物理计量单位（如：㎡、m、台、个、吨、套、kg、桶），去除“元/”、“单价：”等价格前缀（如“单价：元/㎡”应提炼出“㎡”）；
4. price (单价): 纯数字价格；
5. quote_desc_extra: 若从原始参数中剥离出了基价（如“铝锭基价18000元/吨”）、调价规则或计价依据，请提炼记录在此，无则填 null；
6. category_code: 从标准分类编码库中选择，没有填 null。

【返回格式】：严格返回合法 JSON 对象 {"items": [...]}，每个输入 id 对应一个对象：
{"items": [
  {
    "id": "行ID",
    "standard_name": "品类名称",
    "category_code": null,
    "unit": "纯净计量单位",
    "price": 0.0,
    "structured_specs": {"技术参数名": "参数值"},
    "raw_spec": "纯净技术规格原文(过滤基价/表标题/单价)",
    "quote_desc_extra": "基价或调价说明 或 null"
  }
]}
"""

class DeepSeekStandardizer:
    def __init__(self, api_key=None, base_url=None, model=None):
        self.api_key = DEEPSEEK_API_KEY if api_key is None else api_key
        self.base_url = (base_url or DEEPSEEK_BASE_URL).rstrip('/')
        self.model = model or DEEPSEEK_MODEL
        self.gemini_client = GeminiClient() if GEMINI_API_KEY else None

    def _post_chat(self, payload: dict, timeout: int, tag: str):
        """调用 chat/completions，带重试；失败返回 None 并打印原因 (不再静默)"""
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        last_err = ""
        for attempt in range(LLM_MAX_RETRIES + 1):
            try:
                resp = requests.post(f"{self.base_url}/chat/completions", headers=headers, json=payload, timeout=timeout)
                if resp.status_code == 200:
                    return resp.json()["choices"][0]["message"]["content"]
                last_err = f"HTTP {resp.status_code}: {resp.text[:200]}"
                # 4xx (除 429 限流) 重试无意义
                if 400 <= resp.status_code < 500 and resp.status_code != 429:
                    break
            except Exception as e:
                last_err = str(e)
            if attempt < LLM_MAX_RETRIES:
                time.sleep(2 ** attempt)
        print(f"[DeepSeek {tag} Warning] 调用失败: {last_err}")
        return None

    def analyze_sheet_layout(self, sheet_name: str, sample_rows: list) -> dict:
        """优先调用 Gemini (Gemini 3.8 / 3.6 Flash) 动态推导任意工作表的空间拓扑与列角色"""
        if not sample_rows:
            return self._dynamic_layout_fallback(sheet_name, sample_rows)

        user_content = f"工作表: {sheet_name}\n样本数据如下：\n" + json.dumps(sample_rows, ensure_ascii=False, indent=2)

        # 优先 Gemini 引擎
        if self.gemini_client:
            res_json = self.gemini_client.generate_json(
                prompt=user_content,
                system_prompt=SYSTEM_PROMPT_LAYOUT_DYNAMIC,
                timeout=25
            )
            if res_json:
                parsed = self._validate_layout(res_json, sample_rows)
                if parsed:
                    parsed["_source"] = "gemini"
                    return parsed

        if not self.api_key:
            return self._dynamic_layout_fallback(sheet_name, sample_rows)

        payload = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT_LAYOUT_DYNAMIC},
                {"role": "user", "content": user_content}
            ],
            "temperature": LLM_TEMPERATURE,
            "response_format": {"type": "json_object"}
        }

        content = self._post_chat(payload, timeout=25, tag="Layout")
        if content:
            try:
                parsed = self._validate_layout(json.loads(self._clean_json_str(content)), sample_rows)
                if parsed:
                    parsed["_source"] = "llm"
                    return parsed
                print("[DeepSeek Layout Warning] 返回的布局字段不合法，回退至启发式推导。")
            except Exception as e:
                print(f"[DeepSeek Layout Warning] JSON 解析失败: {e}，回退至启发式推导。")

        return self._dynamic_layout_fallback(sheet_name, sample_rows)

    @staticmethod
    def _to_int(v):
        try:
            return int(v) if v is not None and str(v).strip() != '' else None
        except (TypeError, ValueError):
            return None

    def _validate_layout(self, parsed: dict, sample_rows: list):
        """校验并规范 LLM 返回的布局：行列号必须为正整数，列区间合法"""
        if not isinstance(parsed, dict) or not parsed.get("sub_tables"):
            return None
        header_row = self._to_int(parsed.get("header_row")) or 1
        data_start = self._to_int(parsed.get("data_start_row")) or header_row + 1
        if data_start <= header_row:
            data_start = header_row + 1
        tables = []
        for t in parsed["sub_tables"]:
            if not isinstance(t, dict):
                continue
            cs = self._to_int(t.get("col_start")) or 1
            ce = self._to_int(t.get("col_end")) or cs
            if ce < cs:
                continue
            pc = self._to_int(t.get("price_col"))
            uc = self._to_int(t.get("unit_col"))
            tables.append({
                "table_name": t.get("table_name") or f"子表{len(tables) + 1}",
                "col_start": cs,
                "col_end": ce,
                "price_col": pc if pc and cs <= pc <= ce else None,
                "unit_col": uc if uc and cs <= uc <= ce else None,
                "key_columns_desc": t.get("key_columns_desc", "")
            })
        if not tables:
            return None
        return {
            "header_row": header_row,
            "data_start_row": data_start,
            "is_side_by_side": len(tables) > 1,
            "sub_tables": tables
        }

    def standardize_batch(self, items: list, supplier_hint="", brand_hint="") -> list:
        """
        调用 Gemini 3.8 / 3.6 Flash (或 DeepSeek 备用) 批量实体槽位注入与工艺无损提炼。
        每条结果带 _source 字段：'gemini'/'llm' 为模型清洗，'fallback' 为规则降级 (需人工复核)
        """
        if not items:
            return []

        if CODE_LIBRARY_TRUNCATED and CODE_LIBRARY_TEXT:
            print(f"[*] 编码库 {len(CATEGORY_CODE_MAP)} 条，按 CODE_LIBRARY_MAX_LINES="
                  f"{CODE_LIBRARY_MAX_LINES} 注入前 {len(_code_items)} 条 (可用 .env 调大)")
        user_content = (f"供应商: {supplier_hint or '未知'}，品牌: {brand_hint or '未知'}。\n"
                        f"标准分类编码库：\n{CODE_LIBRARY_TEXT or '(空)'}\n"
                        "待清洗物料数据清单：\n" + json.dumps(items, ensure_ascii=False))

        results = None

        # 优先调用 Gemini 引擎
        if self.gemini_client:
            res_json = self.gemini_client.generate_json(
                prompt=user_content,
                system_prompt=SYSTEM_PROMPT_BATCH_DYNAMIC,
                timeout=45
            )
            if res_json:
                results = self._extract_list(res_json)
                if results:
                    for r in results:
                        if isinstance(r, dict):
                            r["_source"] = "gemini"

        # 备选调用 DeepSeek
        if results is None and self.api_key:
            payload = {
                "model": self.model,
                "messages": [
                    {"role": "system", "content": SYSTEM_PROMPT_BATCH_DYNAMIC},
                    {"role": "user", "content": user_content}
                ],
                "temperature": LLM_TEMPERATURE,
                "response_format": {"type": "json_object"}
            }

            content = self._post_chat(payload, timeout=60, tag="Batch")
            if content:
                try:
                    results = self._extract_list(json.loads(self._clean_json_str(content)))
                except Exception as e:
                    print(f"[DeepSeek Batch Warning] JSON 解析失败: {e}")

        if results is None:
            print("[LLM Batch Warning] 模型未返回有效清洗结果，切入规则降级。")
            return self._dynamic_batch_fallback(items)

        for r in results:
            if isinstance(r, dict):
                r["_source"] = "llm"
                specs = r.get("structured_specs")
                if isinstance(specs, dict):
                    r["structured_specs"] = {k: v for k, v in specs.items()
                                             if k not in JUNK_SPEC_KEYS and v not in (None, "", "...")}
        # 模型漏掉的行单独走降级，保证不丢行
        returned_ids = {str(r.get("id")) for r in results if isinstance(r, dict)}
        missing = [it for it in items if str(it.get("id")) not in returned_ids]
        if missing:
            print(f"[DeepSeek Batch Warning] 模型漏返回 {len(missing)} 行，已对这些行走规则降级。")
            results.extend(self._dynamic_batch_fallback(missing))
        return [r for r in results if isinstance(r, dict)]

    @staticmethod
    def _extract_list(parsed):
        """兼容 [...] 与 {"items": [...]} / {任意键: [...]} 两种返回"""
        if isinstance(parsed, list):
            return parsed
        if isinstance(parsed, dict):
            if isinstance(parsed.get("items"), list):
                return parsed["items"]
            for v in parsed.values():
                if isinstance(v, list):
                    return v
        return None

    def _clean_json_str(self, text: str) -> str:
        text = text.strip()
        if text.startswith("```json"): text = text[7:]
        elif text.startswith("```"): text = text[3:]
        if text.endswith("```"): text = text[:-3]
        return text.strip()

    def _dynamic_layout_fallback(self, sheet_name: str, sample_rows: list) -> dict:
        """通用自适应表头探测 (零写死任何固定品类)"""
        header_row = 1
        data_start_row = 2
        price_col = None
        unit_col = None
        max_col = max((len(r["cells"]) for r in sample_rows), default=30) or 30

        header_keys = UNIT_PRICE_KEYS + TOTAL_PRICE_KEYS + UNIT_KEYS + NAME_KEYS + ['规格', '型号']
        for row in sample_rows:
            r_idx = row["row_idx"]
            cells = [str(c).strip() for c in row["cells"]]
            # 合并单元格展开后的标题行 (所有非空格内容相同) 不可能是表头
            if len({c for c in cells if c}) <= 1:
                continue
            # 统计包含表头特征的单元格数
            score = sum(1 for c in cells if c and any(k in c for k in header_keys))
            if score >= 2:
                header_row = r_idx
                data_start_row = r_idx + 1
                total_col = None
                for c_idx, val in enumerate(cells, 1):
                    if not val:
                        continue
                    if any(k in val for k in TOTAL_PRICE_KEYS):
                        total_col = total_col or c_idx
                    elif price_col is None and any(k in val for k in UNIT_PRICE_KEYS):
                        price_col = c_idx
                    elif unit_col is None and any(k in val for k in UNIT_KEYS):
                        unit_col = c_idx
                # 只有合价列、没有单价列时才退而求其次
                if price_col is None:
                    price_col = total_col
                break

        return {
            "header_row": header_row,
            "data_start_row": data_start_row,
            "is_side_by_side": False,
            "_source": "fallback",
            "sub_tables": [
                {
                    "table_name": sheet_name or "主表",
                    "col_start": 1,
                    "col_end": max_col,
                    "price_col": price_col,
                    "unit_col": unit_col,
                    "key_columns_desc": "通用全列自适应读取"
                }
            ]
        }

    def _dynamic_batch_fallback(self, items: list) -> list:
        """通用自适应降级 (零写死品类分支)，字段与 main.py 传入的 llm_input 保持一致"""
        results = []
        for it in items:
            raw_cells = it.get('raw_cells') or {}
            raw_price = it.get('raw_price', 0.0)
            raw_unit = it.get('raw_unit', '')

            specs = {}
            name = ''
            for k, v in raw_cells.items():
                k, v = str(k).strip(), str(v).strip()
                if not k or not v:
                    continue
                if not name and any(nk in k for nk in NAME_KEYS):
                    name = v
                    continue
                if not any(ign in k for ign in IGNORE_SPEC_KEYS):
                    specs[k] = v

            results.append({
                "id": it.get('id', ''),
                "standard_name": name,
                "category_code": "",
                "unit": raw_unit,
                "price": float(raw_price) if raw_price else 0.0,
                "structured_specs": specs,
                "_source": "fallback"
            })
        return results


# ====================================================================== TypeSafe Jev (试点)
ROW_TYPES = {
    "item": "物料报价行：描述一个具体的物料/规格型号，通常带单价",
    "title": "分组或章节标题行，例如“一、阀门类”",
    "header": "表头行，或数据区中重复出现的表头",
    "summary": "合计、小计、总计等汇总行",
    "note": "备注、说明、签章、落款或其他不属于物料的行",
}
TABLE_TYPES = {
    "list": "清单式：每一行是一个物料，单价在某一列",
    "matrix": "矩阵式：行标题(如规格档位) × 列标题(如型号)，交叉格是价格",
}
# JEV_ROW_BATCH 已移到 config


class JevJudge:
    """
    TypeSafe Jev 选择题判定。所有答案都必须落在给定选项内、概率合法，否则整次请求作废 (返回 None)。
    """

    def __init__(self):
        self.api_key = TYPESAFE_API_KEY
        self.url = TYPESAFE_BASE_URL
        self.model = TYPESAFE_MODEL
        self.usage = {"requests": 0, "input_tokens": 0, "output_tokens": 0, "failed": 0}

    @property
    def available(self):
        return bool(self.api_key)

    def ask(self, state, questions: dict, timeout=60):
        if not self.available:
            return None
        body = {"model": self.model, "state": state, "questions": questions}
        headers = {"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"}
        last = ""
        for attempt in range(LLM_MAX_RETRIES + 1):
            try:
                resp = requests.post(self.url, headers=headers, json=body, timeout=timeout)
                if resp.status_code == 200:
                    data = resp.json()
                    u = data.get("usage") or {}
                    self.usage["requests"] += 1
                    self.usage["input_tokens"] += int(u.get("input_tokens") or 0)
                    self.usage["output_tokens"] += int(u.get("output_tokens") or 0)
                    answers = data.get("answers") or {}
                    for qid, q in questions.items():
                        if q["type"] == "choice" and not self._valid_choice(answers.get(qid), q["criteria"]):
                            print(f"[Jev Warning] 问题 {qid} 的答案不合法，本次请求作废")
                            self.usage["failed"] += 1
                            return None
                    return answers
                last = f"HTTP {resp.status_code}: {resp.text[:200]}"
                if resp.status_code not in (429, 500, 502, 503, 529):
                    break
            except Exception as e:
                last = str(e)
            if attempt < LLM_MAX_RETRIES:
                time.sleep(2 ** attempt)
        self.usage["failed"] += 1
        print(f"[Jev Warning] 调用失败: {last}")
        return None

    @staticmethod
    def _valid_choice(answer, criteria):
        try:
            probs = answer["probabilities"]
            nums = [*probs.values(), answer["confidence"]]
            return (answer["choice"] in criteria and set(probs) == set(criteria)
                    and all(isinstance(n, (int, float)) and 0 <= n <= 1 for n in nums)
                    and abs(sum(probs.values()) - 1) < 0.02)
        except (KeyError, TypeError, AttributeError):
            return False

    # ------------------------------------------------------------------ 表格结构
    def judge_layout(self, sample_rows: list):
        """
        一次请求判定：表格类型、表头行、数据起始行、单价列、单位列。
        返回 {字段: (值, 置信度)}；失败返回 None
        """
        if not sample_rows:
            return None
        width = max(len(r["cells"]) for r in sample_rows)
        row_opts = {str(r["row_idx"]): " | ".join(c for c in r["cells"] if c)[:200] or "(空行)" for r in sample_rows}
        col_opts = {}
        for c in range(width):
            vals = [r["cells"][c] for r in sample_rows if c < len(r["cells"]) and r["cells"][c]]
            if vals:
                col_opts[str(c + 1)] = f"第{c + 1}列: " + " | ".join(vals[:6])[:200]
        col_opts_none = {**col_opts, "none": "表中没有这一列"}
        state = {"任务": "判断一张报价表的结构", "样本行": [{"行号": r["row_idx"], "单元格": r["cells"]} for r in sample_rows]}
        questions = {
            "table_type": {"type": "choice", "instructions": "这张表属于哪种结构", "criteria": TABLE_TYPES},
            "header_row": {"type": "choice",
                           "instructions": "表格第一次出现的表头是哪一行 (多级表头时选最下面一行；数据区中重复出现的表头不算)",
                           "criteria": row_opts},
            "data_start_row": {"type": "choice",
                               "instructions": "表头下面数据区的第一行是哪一行 (分组标题行也属于数据区)",
                               "criteria": row_opts},
            "price_col": {"type": "choice", "instructions": "清单式表格中，哪一列是单价 (不是合价/金额)", "criteria": col_opts_none},
            "unit_col": {"type": "choice", "instructions": "哪一列是计量单位", "criteria": col_opts_none},
        }
        answers = self.ask(state, questions)
        if not answers:
            return None
        out = {}
        for k in questions:
            a = answers[k]
            v = a["choice"]
            if k in ("header_row", "data_start_row"):
                v = int(v)
            elif k in ("price_col", "unit_col"):
                v = None if v == "none" else int(v)
            out[k] = (v, float(a["confidence"]))
        return out

    # ------------------------------------------------------------------ 行类型
    def judge_rows(self, headers: list, rows: list):
        """
        rows: [(ref, [单元格值...]), ...]；按 JEV_ROW_BATCH 分批并发判定。
        返回 {ref: (类型, 置信度)}；某批失败时该批的行不出现在结果中
        """
        results = {}

        def run(batch):
            state = {"表头": headers, "行": {ref: [v for v in vals] for ref, vals in batch}}
            qs = {f"r{i}": {"type": "choice", "instructions": f"“行”中键为 {ref} 的这一行属于哪一类",
                            "criteria": ROW_TYPES} for i, (ref, _) in enumerate(batch)}
            ans = self.ask(state, qs)
            if not ans:
                return {}
            return {ref: (ans[f"r{i}"]["choice"], float(ans[f"r{i}"]["confidence"])) for i, (ref, _) in enumerate(batch)}

        batches = [rows[i:i + JEV_ROW_BATCH] for i in range(0, len(rows), JEV_ROW_BATCH)]
        with ThreadPoolExecutor(max_workers=4) as ex:
            for part in ex.map(run, batches):
                results.update(part)
        return results
