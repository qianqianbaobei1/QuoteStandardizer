# -*- coding: utf-8 -*-
"""
llm_auditor.py - 大语言模型深度业务逻辑终审与强制闭环自愈回填引擎
"""
import json
import re
import requests
from config import DEEPSEEK_API_KEY, DEEPSEEK_BASE_URL, DEEPSEEK_MODEL, LLM_TEMPERATURE

AUDIT_FIX_PROMPT = """你是一个国家级建设工程招投标与建材物料领域的终审专家。
现在需要对即将交付客户的建材报价单进行【强制终审检查与自动自愈回填】。

【必须检查和修复的致命错误】：
1. 【规格参数中混入非技术参数】：
   - 检查“规格参数”与“厂家原始规格参数”中，是否包含了“铝锭基价”、“铜基价”、“基价”、“指导价”、“表标题”、“单价”、“金额”、“运费”、“税率”等内容。
   - 若有，必须彻底从规格参数中剔除！
   - 将剥离出来的基价或调价依据（如“铝锭基价18000元/吨”），自动追加写入到“报价说明”中！
2. 【品名与规格参数混淆】：
   - “名称规格”必须是纯净品类名，不得堆砌一堆参数。
3. 【计量单位量纲合理性】：
   - 铝板/门窗/幕墙等板材应为 ㎡；阀门/设备应为 台/个；涂料应为 kg/桶；型材应为 m/吨。不得填“元/㎡”或“单价：元”。
4. 【价格纯数字】：
   - 市场价必须是干净数字。

待核审的数据行（JSON格式）：
{data_json}

请以严格的 JSON 格式输出修正建议（只返回需要修正的行和字段，不需要修正的行不用返回）：
{{
  "need_healing": true,
  "summary": "审核简述",
  "corrections": [
    {{
      "row_id": 1,
      "fixes": {{
        "名称规格": "修正值 (若无问题则不填)",
        "规格参数": "修正后的纯净技术参数 (若无问题则不填)",
        "厂家原始规格参数": "修正后的纯净原始技术参数 (若无问题则不填)",
        "单位": "修正后的规范单位 (若无问题则不填)",
        "报价说明": "修正后的报价说明 (若无问题则不填)"
      }}
    }}
  ]
}}
请直接输出合法 JSON，不要包含任何前缀后缀废话。
"""

class LLMAuditor:
    # 价格只取原表、原文留痕不可改写：模型无权动这些列
    IMMUTABLE_FIELDS = {'报价ID', '市场价', '工程价', '厂家原始规格参数', '厂家报价文件页码',
                        '顺序号', '供应商名称', '类别编码'}
    # 只允许自愈这四项；其余列即使模型返回也一律忽略
    HEALABLE_FIELDS = ('名称规格', '规格参数', '单位', '报价说明')

    def __init__(self):
        self.api_key = DEEPSEEK_API_KEY
        self.base_url = DEEPSEEK_BASE_URL.rstrip('/')
        self.model = DEEPSEEK_MODEL

    def verify_and_auto_heal(self, ws, headers):
        """
        强制两道关卡：就地对 openpyxl Worksheet 数据进行模型审校，若有明确修正建议，强制执行就地自愈回填
        """
        rows_data = []
        for r in range(2, ws.max_row + 1):
            row_dict = {}
            for c_idx, h in enumerate(headers, 1):
                if h:
                    row_dict[h] = ws.cell(r, c_idx).value
            row_dict['_row_idx'] = r
            row_dict['seq'] = r - 1
            rows_data.append(row_dict)

        if not rows_data:
            return

        print(f"[*] [AI 自我核对中枢] 正在对当前 Sheet ({ws.title}) 的 {len(rows_data)} 行物料进行深度业务逻辑审计与自愈核验...")
        
        # 分批送审（每批最多 30 行）
        batch_size = 30
        healed_count = 0
        for b_idx in range(0, len(rows_data), batch_size):
            chunk = rows_data[b_idx:b_idx + batch_size]
            prompt = AUDIT_FIX_PROMPT.format(data_json=json.dumps(chunk, ensure_ascii=False, indent=2))
            payload = {
                'model': self.model,
                'messages': [
                    {'role': 'system', 'content': '你是一个资深严谨的造价与数据治理专家。'},
                    {'role': 'user', 'content': prompt}
                ],
                'temperature': LLM_TEMPERATURE,
                'response_format': {'type': 'json_object'}
            }
            try:
                resp = requests.post(f"{self.base_url}/chat/completions", headers={
                    'Authorization': f'Bearer {self.api_key}',
                    'Content-Type': 'application/json'
                }, json=payload, timeout=60)
                if resp.status_code == 200:
                    res_json = json.loads(resp.json()['choices'][0]['message']['content'])
                    corrections = res_json.get('corrections', [])
                    for corr in corrections:
                        seq = corr.get('row_id')
                        fixes = corr.get('fixes', {})
                        # 定位对应的行
                        target_row = None
                        for item in chunk:
                            if str(item.get('seq')) == str(seq) or str(item.get('_row_idx')) == str(seq):
                                target_row = item['_row_idx']
                                break
                        if not target_row:
                            try:
                                s_int = int(seq)
                                for item in chunk:
                                    if item.get('seq') == s_int:
                                        target_row = item['_row_idx']
                                        break
                                if not target_row and s_int + 1 <= ws.max_row:
                                    target_row = s_int + 1
                            except (ValueError, TypeError):
                                pass
                            
                        if target_row:
                            heal_log = []
                            for f_name, f_val in fixes.items():
                                if f_val is None or f_name not in headers:
                                    continue
                                if f_name in self.IMMUTABLE_FIELDS or f_name not in self.HEALABLE_FIELDS:
                                    print(f"    [自愈拦截] 行 {target_row} 字段 [{f_name}] 不允许模型改写 (价格/原文留痕/标识列)")
                                    continue
                                col_idx = headers.index(f_name) + 1
                                orig_val = ws.cell(target_row, col_idx).value
                                if str(orig_val).strip() != str(f_val).strip():
                                    ws.cell(target_row, col_idx).value = f_val
                                    healed_count += 1
                                    heal_log.append(f"{f_name}:{str(orig_val)[:20]}→{str(f_val)[:20]}")
                                    print(f"    [自愈修复成功] 行 {target_row} 字段 [{f_name}]: 原始值='{str(orig_val)[:30]}...' -> 自愈值='{str(f_val)[:30]}...'")
                            # 自愈必须留痕：改写内容回填到报价说明，否则交付表无法溯源
                            if heal_log and '报价说明' in headers:
                                dc = headers.index('报价说明') + 1
                                old = str(ws.cell(target_row, dc).value or '').strip()
                                stamp = '[模型自愈] ' + '；'.join(heal_log)
                                ws.cell(target_row, dc).value = f"{old}；{stamp}".strip('；') if old else stamp
            except Exception as e:
                print(f"    [LLM Auditor Chunk Error] 审计分批自愈调用异常: {e}")

        print(f"[*] [AI 自我核对中枢] 审核完成！共执行自动就地修复: {healed_count} 处单元格。")
