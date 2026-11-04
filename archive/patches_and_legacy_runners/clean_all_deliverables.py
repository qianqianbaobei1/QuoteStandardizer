# -*- coding: utf-8 -*-
import os
import re
import glob
import openpyxl

DELIVERY_DIRS = [
    "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries/成稿",
    "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries"
]

FORBIDDEN_PATTERNS = [
    r'(?:^|[;；\s\n])(?:产品编码|物料编码|商品编码|内部编码|条形码|条码|代码|三级类别编码|二级类别|一级类别|材料名称|属性项属性值|物料名称|类别名称|分类编码)\s*[:：][^;；\n]*',
]

def clean_spec_text(text):
    if not text:
        return ""
    s = str(text)
    for p in FORBIDDEN_PATTERNS:
        s = re.sub(p, '', s)
    # 规整连续分隔符
    s = re.sub(r'[;；\n]{2,}', '\n', s)
    s = re.sub(r'^[;；\s\n]+|[;；\s\n]+$', '', s).strip()
    return s

def clean_multiline_spec(text):
    if not text:
        return ""
    lines = str(text).split('\n')
    clean_lines = []
    forbidden_keys = ('表头', '表标题', '基价', '单价', '总价', '金额', '指导价', '调价',
                      '产品编码', '物料编码', '商品编码', '内部编码', '条形码', '编码', '条码', '代码',
                      '税号', '账号', '帐号', '开户行', '一级类别', '二级类别', '三级类别', '材料名称', '属性项属性值',
                      '物料名称', '类别名称', '分类编码')
    for l in lines:
        l_str = l.strip()
        if not l_str:
            continue
        k = l_str.split(':', 1)[0].strip() if ':' in l_str else (l_str.split('：', 1)[0].strip() if '：' in l_str else l_str)
        if any(fb in k for fb in forbidden_keys):
            continue
        clean_lines.append(l_str)
    return '\n'.join(clean_lines)

WUHAN_PAGE_MAP = {
    '第1页': 'quote_img_电线1.png',
    '第2页': 'quote_img_电线2.png',
    '第3页': 'quote_img_花线.png',
    '第4页': 'quote_img_佳居线.png',
    '第5页': 'quote_img_铝合金电缆.png',
    '第6页': 'quote_img_系统结算价.png',
    '1': 'quote_img_电线1.png',
    '2': 'quote_img_电线2.png',
    '3': 'quote_img_花线.png',
    '4': 'quote_img_佳居线.png',
    '5': 'quote_img_铝合金电缆.png',
    '6': 'quote_img_系统结算价.png',
}

def process_file(xlsx_path):
    fname = os.path.basename(xlsx_path)
    if "总表" in fname or "报告" in fname or "对账" in fname:
        return None
    if not os.path.isfile(xlsx_path):
        return None

    try:
        wb = openpyxl.load_workbook(xlsx_path)
    except Exception as e:
        print(f"[-] 无法打开 {fname}: {e}")
        return None

    target_sheet = "标准化报价" if "标准化报价" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[target_sheet]

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    if not headers or "名称规格" not in headers:
        return None

    spec_c = headers.index("规格参数") + 1 if "规格参数" in headers else None
    raw_spec_c = headers.index("厂家原始规格参数") + 1 if "厂家原始规格参数" in headers else None
    page_c = headers.index("厂家报价文件页码") + 1 if "厂家报价文件页码" in headers else None
    img_c = headers.index("厂家报价图片文件名") + 1 if "厂家报价图片文件名" in headers else None

    is_wuhan = "武汉第二电线电缆" in fname
    modified_rows = 0
    img_filled = 0

    for r in range(2, ws.max_row + 1):
        # 1. 清洗规格参数
        if spec_c:
            val = ws.cell(r, spec_c).value
            if val:
                c_val = clean_multiline_spec(val)
                if c_val != val:
                    ws.cell(r, spec_c).value = c_val if c_val else None
                    modified_rows += 1

        # 2. 清洗厂家原始规格参数
        if raw_spec_c:
            val = ws.cell(r, raw_spec_c).value
            if val:
                c_val = clean_spec_text(val)
                if c_val != val:
                    ws.cell(r, raw_spec_c).value = c_val if c_val else None
                    modified_rows += 1

        # 3. 补全厂家报价图片文件名
        if img_c:
            cur_img = str(ws.cell(r, img_c).value or "").strip()
            if not cur_img:
                page_val = str(ws.cell(r, page_c).value or "").strip() if page_c else ""
                if is_wuhan and page_val in WUHAN_PAGE_MAP:
                    target_img = WUHAN_PAGE_MAP[page_val]
                else:
                    m = re.search(r'\d+', page_val)
                    if m:
                        target_img = f"quote_img{m.group(0)}.png"
                    elif page_val:
                        safe_s = re.sub(r'[\/\\:\*\?\"<>\|]', '_', page_val).strip()
                        target_img = f"quote_img_{safe_s}.png"
                    else:
                        target_img = "quote_img1.png"
                ws.cell(r, img_c).value = target_img
                img_filled += 1

    wb.save(xlsx_path)
    return {
        "file": fname,
        "rows": ws.max_row - 1,
        "cleaned_specs": modified_rows,
        "img_filled": img_filled
    }

def main():
    seen = set()
    all_files = []
    for d in DELIVERY_DIRS:
        if os.path.exists(d):
            for f in glob.glob(os.path.join(d, "*.xlsx")):
                if os.path.isfile(f) and f not in seen:
                    seen.add(f)
                    all_files.append(f)

    print(f"[*] 开始全量清洗交付文件，共发现 {len(all_files)} 个 Excel 文件...")
    results = []
    for f in all_files:
        res = process_file(f)
        if res:
            results.append(res)
            print(f"[✓] {res['file']}: 行数={res['rows']}, 净化参数行={res['cleaned_specs']}, 补全底稿图片名={res['img_filled']}")

    print("\n[+] 全量清洗完成！共处理文件:", len(results))

if __name__ == "__main__":
    main()
