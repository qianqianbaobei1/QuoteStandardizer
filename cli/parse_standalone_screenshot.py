import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))

# -*- coding: utf-8 -*-
"""
parse_standalone_screenshot.py - 单独解析（不走批量源包）的两份报价截图

来源：用户在会话里贴的两张报价表截图（无原始文件可读），按 4~10 倍放大逐格核对后转录。
  1) 友发：电线、电缆套管（KGB 扣压式 / JDG 紧定式镀锌穿线管），12 行
  2) 广东奥卡斯网络科技有限公司：光纤收发器/光模块/耦合器/尾纤/配线架，11 行

零编造纪律：
- 只转录截图里读得到的格子；分类编码取自标准编码库（572305 等原表方括号已标注，直接采用）
- 被公章遮挡 / 被裁切的价格留空并写“需核对原件”，不猜、不补
"""
import json
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

from builder import StandardizedWorkbookBuilder
from config import DEFAULT_TEMPLATE_PATH
from models import CanonicalMaterialItem

DESKTOP = os.path.expanduser("~/Desktop")

# ------------------------------------------------------------------ 截图 1：友发
# (品种, 壁厚mm, 外径mm, 含税价)
YOUFA_ROWS = [
    ("KGB扣压式镀锌穿线管", "1.2", "20", 8.6),
    ("KGB扣压式镀锌穿线管", "1.2", "25", 10.8),
    ("KGB扣压式镀锌穿线管", "1.2", "32", 13.8),
    ("KGB扣压式镀锌穿线管", "1.2", "40", 18.7),
    ("KGB扣压式镀锌穿线管", "1.2", "50", 24.1),
    ("JDG紧定式镀锌穿线管", "1.5", "16", 7.2),
    ("JDG紧定式镀锌穿线管", "1.6", "20", 11.4),
    ("JDG紧定式镀锌穿线管", "1.6", "25", 15.1),
    ("JDG紧定式镀锌穿线管", "1.6", "32", 20.5),
    ("JDG紧定式镀锌穿线管", "1.6", "40", 24.2),   # 印在公章圆环下，R 通道确认 24.2（不是 24.1）
    ("JDG紧定式镀锌穿线管", "1.6", "50", 31.3),
    ("KGB扣压式镀锌穿线管", "1", "16", 6.3),
]

# ------------------------------------------------------------------ 截图 2：广东奥卡斯科技有限公司
# (材料名称, 类别编码, 型号, 参数, 单位, 价格)
AOKASI_ROWS = [
    ("光纤收发器", "572305", "ACB-GS-20", "千兆单模双纤收发器/SC/20KM", "台", 125),
    ("光纤收发器", "572305", "ACB-GS13/15-12-20", "千兆单模单纤收发器/SC/20KM，A、B端", "对", 300),
    ("光模块", "572303", "SFP-10G-LR-10", "万兆单模1310nmSFP 10KM LC DDM", "只", 180),
    ("光模块", "572303", "SFP-GE-LX-SM1310-10", "千兆单模1310nmSFP 10KM LC DDM", "只", 90),
    ("光纤耦合器", "571923", "AC-OA-L02", "LC双工耦合器", "个", 6),
    ("光纤耦合器", "571923", "AC-OA-C01", "SC单工耦合器", "个", 3),
    ("光纤跳线、尾纤", "571917", "ACL01S2S", "单模单芯LC尾纤，长度1.5米", "条", 6),
    ("光纤跳线、尾纤", "571917", "ACF01S2S", "单模单芯FC尾纤，长度1.5米", "条", 6),
    ("光纤跳线、尾纤", "571917", "ACT01S2S", "单模单芯ST尾纤，长度1.5米", "条", 6),
    ("光纤跳线、尾纤", "571917", "ACC01S2S", "单模单芯SC尾纤，长度1.5米", "条", 6),
    ("光纤配线架", "571909", "ACFPP3048-CS2/ACFPP3148-FS2",
     "48口机架式光纤终端盒，SC/FC接口，含耦合器尾纤（厚款1.2mm）", "套", None),
]
AOKASI_SUPPLIER = "广东奥卡斯网络科技有限公司"   # 公章环形字（极易漏“网络”两字）


def build_youfa(out_dir):
    items = []
    for i, (kind, thick, od, price) in enumerate(YOUFA_ROWS, 1):
        specs = {"品种": kind, "壁厚(mm)": thick, "外径(mm)": od}
        it = CanonicalMaterialItem(
            quote_id=f"YF{i:05d}",
            supplier="",                       # 截图里没有供方名称，只写了“本份报价单品牌：友发”
            brand="友发",
            category_name="电线、电缆套管",
            category_code="260601",
            structured_specs=specs,
            raw_spec="；".join(f"{k}:{v}" for k, v in specs.items()),
            file_page="截图1",
            unit="根/3.6米",
            market_price=price,                # 未读到就留空，绝不补 0
            is_freight_included="不含",
            is_tax_included="是",
            tax_rate=13,
            tax_desc="税率13%",
            source_ref=f"截图1!R{i + 1}",
            clean_source="manual_transcription",
            category_source="dict",
            quote_desc="本表由聊天截图转录（原件未提供）；供方名称截图中未标注（公章只露出“有限公司/专用章”），留空待补",
        )
        it.add_note("本表由聊天截图转录，供方名称截图中未标注")
        items.append(it)
    out = os.path.join(out_dir, "友发-电线电缆套管-标准解析.xlsx")
    StandardizedWorkbookBuilder(DEFAULT_TEMPLATE_PATH).build_from_canonical_items(
        items, out, os.path.join(out_dir, "_图片库_友发"))
    return out, items


def build_aokasi(out_dir):
    items = []
    for i, (name, code, model, param, unit, price) in enumerate(AOKASI_ROWS, 1):
        specs = {"型号": model, "参数": param}
        it = CanonicalMaterialItem(
            quote_id=f"AKS{i:05d}",
            supplier=AOKASI_SUPPLIER,
            brand="奥卡斯",
            category_name=name,
            category_code=code,
            structured_specs=specs,
            raw_spec="；".join(f"{k}:{v}" for k, v in specs.items()),
            file_page="截图2",
            unit=unit,
            market_price=price,
            source_ref=f"截图2!R{i + 1}",
            clean_source="manual_transcription",
            category_source="dict",     # 编码取自原表方括号标注，与标准库一致
            quote_desc="本表由聊天截图转录（原件未提供）；税率、运费、报价地区截图中未标注"
                       + ("；最后一行原单价格格为空" if price is None else ""),
        )
        it.add_note("本表由聊天截图转录，税率/运费/报价地区截图中未标注")
        if price is None:
            it.add_note("原单该行价格格为空（公章压在这格上，R 通道暗像素确认无印刷数字），需向供方补价")
        items.append(it)
    out = os.path.join(out_dir, f"{AOKASI_SUPPLIER}-标准解析.xlsx")
    StandardizedWorkbookBuilder(DEFAULT_TEMPLATE_PATH).build_from_canonical_items(
        items, out, os.path.join(out_dir, f"_图片库_{AOKASI_SUPPLIER}"))
    return out, items


def dump_json(path, items):
    rows = [{
        "报价ID": it.quote_id, "供应商名称": it.supplier, "名称规格": it.category_name,
        "类别编码": it.category_code, "品牌": it.brand,
        "规格参数": it.structured_specs, "单位": it.unit, "市场价": it.market_price,
        "备注": it.note_list(),
    } for it in items]
    with open(path, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    stale = ["圣火科技股份有限公司-标准解析.xlsx", "圣火科技股份有限公司-转录底稿.json"]
    for name in stale:                    # 上一版把公章读成“圣火科技”，别留错文件
        p = os.path.join(DESKTOP, name)
        if os.path.exists(p):
            os.remove(p)
            print(f"[clean] 已删除错误旧件 {name}")

    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    target_dir = os.path.join(base_dir, "deliveries", "成稿")
    os.makedirs(target_dir, exist_ok=True)

    f1, items1 = build_youfa(target_dir)
    f2, items2 = build_aokasi(target_dir)
    dump_json(os.path.join(target_dir, "友发-电线电缆套管-转录底稿.json"), items1)
    dump_json(os.path.join(target_dir, f"{AOKASI_SUPPLIER}-转录底稿.json"), items2)

    # 自检：行数、编码、价格必须与转录源一致
    assert len(items1) == len(YOUFA_ROWS) == 12 and len(items2) == len(AOKASI_ROWS) == 11
    assert all(it.category_code for it in items1 + items2), "不许出现空类别编码"
    assert [it.market_price for it in items1][:3] == [8.6, 10.8, 13.8]
    assert items1[9].market_price == 24.2, "第10行 24.2 是 R 通道确认值"
    assert items2[0].market_price == 125 and items2[-1].market_price is None
    print(f"[OK] {f1}\n[OK] {f2}")
