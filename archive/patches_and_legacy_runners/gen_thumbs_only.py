import os
from PIL import Image as PILImage

img_dir = "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries/成稿/武汉第二电线电缆有限公司_图片库"
sheet_thumbs = {}

for s_name in ["电线1", "电线2", "花线", "佳居线", "铝合金电缆", "系统结算价"]:
    src_img = os.path.join(img_dir, f"quote_img_{s_name}.png")
    if os.path.exists(src_img):
        thumb_path = os.path.join(img_dir, f"thumb_{s_name}.jpg")
        with PILImage.open(src_img) as pim:
            pim.thumbnail((200, 120))
            pim.convert("RGB").save(thumb_path, "JPEG", quality=75)
        sheet_thumbs[s_name] = thumb_path
        print(f"Created: {thumb_path} (size: {os.path.getsize(thumb_path)} bytes)")

print("All thumbnails ready.")
