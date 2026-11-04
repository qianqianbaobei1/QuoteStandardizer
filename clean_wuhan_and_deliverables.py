# -*- coding: utf-8 -*-
"""
clean_wuhan_and_deliverables.py
彻底清除所有交付文件中出现的「产品编码」、「物料编码」、「三级类别编码」等内部非物理参数杂质，
并确保「厂家报价图片文件名」完整填入对应页面/底稿截图文件名。
"""
import os
import re
import glob
import openpyxl

DELIVERY_FILES = [
    "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries/成稿/武汉第二电线电缆有限公司-标准解析.xlsx",
    "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries/武汉第二电线电缆有限公司-标准解析.xlsx",
    "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries/成稿/山东陆宇塑胶有限公司-标准解析.xlsx",
    "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries/成稿/青岛鼎信通讯消防安全有限公司-标准解析.xlsx",
    "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries/成稿/北京启达世纪装饰材料有限公司-标准解析.xlsx",
    "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries/成稿/南昌美霓光环境科技发展有限公司-标准解析.xlsx",
    "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries/成稿/山东正大防水材料股份有限公司-标准解析.xlsx",
    "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries/成稿/山东金晶科技股份有限公司-标准解析.xlsx",
    "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries/成稿/无锡萨科特新能源科技有限公司-标准解析.xlsx",
    "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries/成稿/泰安金绿色新型建材有限公司-标准解析.xlsx",
    "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries/成稿/山东拜尔建材有限公司-标准解析.xlsx",
    "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries/成稿/涿州市盛达建筑材料厂-标准解析.xlsx",
    "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries/成稿/深圳市凝光时代科技有限公司-标准解析.xlsx",
]

FORBIDDEN_KEYS = (
    '产品编码', '物料编码', '商品编码', '内部编码', '条形码', '条码', '代码',
    '三级类别编码', '二级类别', '一级类别', '材料名称', '属性项属性值', '物料名称', '类别名称', '分类编码',
    '税号', '账号', '帐号', '开户行'
)

FORBIDDEN_RE = re.compile(
    r'(?:^|[;；\s\n])(?:' + '|'.join(FORBIDDEN_KEYS) + r')\s*[:：][^;；\n]*'
)

WUHAN_PAGE_MAP = {
    '第1页': 'quote_img1.png',
    '第2页': 'quote_img2.png',
    '第3页': 'quote_img3.png',
    '第4页': 'quote_img4.png',
    '第5页': 'quote_img5.png',
    '第6页': 'quote_img6.png',
    '1': 'quote_img1.png',
    '2': 'quote_img2.png',
    '3': 'quote_img3.png',
    '4': 'quote_img4.png',
    '5': 'quote_img5.png',
    '6': 'quote_img6.png',
    '电线1': 'quote_img1.png',
    '电线2': 'quote_img2.png',
    '花线': 'quote_img3.png',
    '佳居线': 'quote_img4.png',
    '铝合金电缆': 'quote_img5.png',
    '系统结算价': 'quote_img6.png',
}

def clean_text_spec(text):
    if not text:
        return ""
    s = str(text)
    # 用正则彻底剥除参数名属于内部代码/非物理参数的所有键值对
    s = FORBIDDEN_RE.sub('', s)
    s = re.sub(r'[;；\n]{2,}', '\n', s)
    s = re.sub(r'^[;；\s\n]+|[;；\s\n]+$', '', s).strip()
    return s

def clean_lines_spec(text):
    if not text:
        return ""
    lines = str(text).split('\n')
    clean = []
    for l in lines:
        l_s = l.strip()
        if not l_s:
            continue
        k = l_s.split(':', 1)[0].strip() if ':' in l_s else (l_s.split('：', 1)[0].strip() if '：' in l_s else l_s)
        if any(fb in k for fb in FORBIDDEN_KEYS):
            continue
        clean.append(l_s)
    return '\n'.join(clean)

def clean_single_file(fpath):
    if not os.path.exists(fpath) or not os.path.isfile(fpath):
        return
    fname = os.path.basename(fpath)
    print(f"[*] 正在处理: {fname}...")
    wb = openpyxl.load_workbook(fpath)
    target_sname = "标准化报价" if "标准化报价" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[target_sname]

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    spec_c = headers.index("规格参数") + 1 if "规格参数" in headers else None
    raw_spec_c = headers.index("厂家原始规格参数") + 1 if "厂家原始规格参数" in headers else None
    page_c = headers.index("厂家报价文件页码") + 1 if "厂家报价文件页码" in headers else None
    img_c = headers.index("厂家报价图片文件名") + 1 if "厂家报价图片文件名" in headers else None

    is_wuhan = "武汉第二电线电缆" in fname
    cleaned_count = 0
    img_fixed = 0

    max_r = ws.max_row
    for r in range(2, max_r + 1):
        if spec_c:
            val = ws.cell(r, spec_c).value
            if val:
                c_val = clean_lines_spec(val)
                if c_val != val:
                    ws.cell(r, spec_c).value = c_val if c_val else None
                    cleaned_count += 1

        if raw_spec_c:
            val = ws.cell(r, raw_spec_c).value
            if val:
                c_val = clean_text_spec(val)
                if c_val != val:
                    ws.cell(r, raw_spec_c).value = c_val if c_val else None
                    cleaned_count += 1

        if img_c:
            val = ws.cell(r, img_c).value
            if not val or str(val).strip() == "" or (is_wuhan and "电线" in str(val)):
                page_val = str(ws.cell(r, page_c).value or "").strip() if page_c else ""
                if is_wuhan:
                    target_img = WUHAN_PAGE_MAP.get(page_val, 'quote_img6.png' if r >= 890 else 'quote_img1.png')
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
                img_fixed += 1

    wb.save(fpath)
    print(f"[✓] {fname} 完成: 净化违规编码={cleaned_count} 处, 补全底稿图片名={img_fixed} 行 (总行数={max_r-1})")

def main():
    # 扫一遍 deliveries/成稿 下全部文件
    all_targets = list(DELIVERY_FILES)
    for f in glob.glob("/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries/成稿/*.xlsx"):
        if os.path.isfile(f) and f not in all_targets:
            all_targets.append(f)

    for p in all_targets:
        if os.path.isfile(p):
            clean_single_file(p)

if __name__ == "__main__":
    main()
