import openpyxl
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
import os

input_file = "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries/成稿/武汉第二电线电缆有限公司-标准解析.xlsx"
img_dir = "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries/成稿/武汉第二电线电缆有限公司_图片库"

wb = openpyxl.load_workbook(input_file)
ws = wb["标准化报价"]

# 创建轻量缩略图
sheet_thumbs = {}
for s_name in ["电线1", "电线2", "花线", "佳居线", "铝合金电缆", "系统结算价"]:
    src_img = os.path.join(img_dir, f"quote_img_{s_name}.png")
    if os.path.exists(src_img):
        thumb_path = os.path.join(img_dir, f"thumb_{s_name}.jpg")
        with PILImage.open(src_img) as pim:
            pim.convert("RGB").save(thumb_path, "JPEG", quality=70)
        sheet_thumbs[s_name] = thumb_path

print("Sheet thumbs created:", sheet_thumbs)
