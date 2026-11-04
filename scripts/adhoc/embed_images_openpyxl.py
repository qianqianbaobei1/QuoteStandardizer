from _adhoc_paths import DELIVERIES_DIR
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
import os
import sys

file_path = os.path.join(DELIVERIES_DIR, "成稿", "武汉第二电线电缆有限公司-标准解析.xlsx")
root_path = os.path.join(DELIVERIES_DIR, "武汉第二电线电缆有限公司-标准解析.xlsx")
img_dir = os.path.join(DELIVERIES_DIR, "成稿", "武汉第二电线电缆有限公司_图片库")

print(f"[*] 加载工作簿: {file_path}")
wb = openpyxl.load_workbook(file_path)
ws = wb["标准化报价"]

# 准备缩略图映射
page_to_thumb = {
    "第1页": os.path.join(img_dir, "thumb_电线1.jpg"),
    "第2页": os.path.join(img_dir, "thumb_电线2.jpg"),
    "第3页": os.path.join(img_dir, "thumb_花线.jpg"),
    "第4页": os.path.join(img_dir, "thumb_佳居线.jpg"),
    "第5页": os.path.join(img_dir, "thumb_铝合金电缆.jpg"),
    "第6页": os.path.join(img_dir, "thumb_系统结算价.jpg"),
}

# 厂家报价图片文件名 列是 I 列 (9)
# 厂家报价文件页码 列是 H 列 (8)

# 清空 I 列文本，并在每个单元格中嵌入 OpenpyxlImage
print(f"[*] 遍历行并嵌入底稿图片 (共 {ws.max_row} 行)...")
embedded_count = 0

# 为避免 59,433 张图片导致 Excel 文件过大崩溃，
# 前 5 个 Sheet（共约 548 行）100% 逐行嵌入！
# 系统结算价（第6页）前 100 行嵌入！
for r in range(2, ws.max_row + 1):
    page_val = ws.cell(r, 8).value
    p_str = str(page_val or "").strip()
    
    # 单元格文字置空，放图不放字！
    ws.cell(r, 9).value = None
    
    thumb_path = page_to_thumb.get(p_str)
    if thumb_path and os.path.exists(thumb_path):
        # 限制最大嵌入数，保证 Excel 能够秒开且不崩
        if p_str != "第6页" or r < 1000:
            img = OpenpyxlImage(thumb_path)
            img.width = 100
            img.height = 55
            ws.add_image(img, f"I{r}")
            ws.row_dimensions[r].height = 58
            embedded_count += 1

print(f"[*] 嵌入完成，共嵌入图片: {embedded_count} 张。正在保存...")
wb.save(file_path)

import shutil
shutil.copy2(file_path, root_path)
print(f"[✓] 成功保存并同步交付文件！")
