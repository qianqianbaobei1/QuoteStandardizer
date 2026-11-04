import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))

import os

if __name__ == '__main__':
    import openpyxl
    from openpyxl.drawing.image import Image
    from PIL import Image as PILImage

    base_dir = _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__)))
    input_file = os.path.join(base_dir, "deliveries", "成稿", "武汉第二电线电缆有限公司-标准解析.xlsx")
    img_dir = os.path.join(base_dir, "deliveries", "成稿", "武汉第二电线电缆有限公司_图片库")

    if os.path.exists(input_file):
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
    else:
        print("Sample file not found, skipping embed test.")
