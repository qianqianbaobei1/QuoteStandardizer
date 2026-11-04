import openpyxl

file_path = "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries/成稿/武汉第二电线电缆有限公司-标准解析.xlsx"
root_path = "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries/武汉第二电线电缆有限公司-标准解析.xlsx"

print("[*] 打开工作簿进行更新...")
wb = openpyxl.load_workbook(file_path)
ws = wb["标准化报价"]

# 查找列索引
header_map = {}
for col in range(1, ws.max_column + 1):
    val = ws.cell(1, col).value
    if val:
        header_map[str(val).strip()] = col

page_col = header_map.get("厂家报价文件页码", 8)
img_fn_col = header_map.get("厂家报价图片文件名", 9)

print(f"[*] 厂家报价文件页码 列: {page_col}, 厂家报价图片文件名 列: {img_fn_col}")

# 逐行根据页码填入对应的底稿截图文件名
for r in range(2, ws.max_row + 1):
    page_val = ws.cell(r, page_col).value
    p_str = str(page_val or "").strip()
    # 提取页码数字，例如 "第1页" -> "quote_img1.png"
    import re
    m = re.search(r'\d+', p_str)
    if m:
        ws.cell(r, img_fn_col).value = f"quote_img{m.group(0)}.png"
    else:
        ws.cell(r, img_fn_col).value = "quote_img1.png"

print(f"[*] 保存至: {file_path}")
wb.save(file_path)

import shutil
shutil.copy2(file_path, root_path)
print(f"[✓] 同步完成: {root_path}")
