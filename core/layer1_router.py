# -*- coding: utf-8 -*-
"""
第 1 层：资产分流器 (Asset Router)
职责：
1. 接收输入（单文件、文件夹、解压后的数据包）；
2. 自动识别并分类文件资产：
   - 【报价主表 (QUOTE_MAIN)】：Excel 或多页 PDF / 图片单据；
   - 【报价说明/工程折扣单 (TERMS_DISCOUNT)】：折扣记录单、商务条款说明图、填空单；
   - 【无关附件/废表 (ATTACHMENT_IGNORE)】：检测报告、授权委托书、资质证书、图纸、照片；
3. 提取供应商名称先验与品牌推断。
"""
import os
import re
from typing import List, Tuple, Optional
from core.ir_models import AssetType, DocumentAsset, QuoteDocumentIR

# 常用公司与关键词匹配正则
COMPANY_RE = re.compile(r'([\u4e00-\u9fa5（）\(\)]{4,30}?(?:股份有限公司|有限责任公司|有限公司|股份公司))')
FILENAME_NOISE_RE = re.compile(r'标准化输出|标准解析|报价单|报价表|价格表|产品资料|报价数据|明细|总表|第\d+页|[\d_\-（）\(\)\s.]')

# 附件排除关键词（绝不作为报价主表误入解析管线）
ATTACHMENT_KEYWORDS = {
    "营业执照", "授权书", "授权委托书", "检测报告", "检验报告", "型式检验",
    "CCC认证", "质量认证", "发票", "收据", "合同", "中标通知书", "协议书", "安装图"
}

# 商务条款/折扣记录单关键词
TERMS_KEYWORDS = {
    "商务条款", "商务说明", "报价说明", "折扣单", "折扣表", "工程折扣", "折扣说明",
    "价格说明", "优惠政策", "付款方式", "运费说明", "开票说明"
}


class AssetRouter:
    """工业级资产分流器"""

    def __init__(self, input_path: str):
        self.input_path = os.path.abspath(input_path)

    def route(self) -> QuoteDocumentIR:
        """主入口：扫描并结构化分流资产"""
        doc = QuoteDocumentIR()
        all_files: List[str] = []

        if os.path.isdir(self.input_path):
            for root, _, files in os.walk(self.input_path):
                for f in files:
                    if not f.startswith('.') and not f.startswith('~$'):
                        all_files.append(os.path.join(root, f))
        elif os.path.isfile(self.input_path):
            all_files.append(self.input_path)
        else:
            raise FileNotFoundError(f"输入路径不存在: {self.input_path}")

        # 尝试推导供应商名称
        doc.supplier_name = self._detect_supplier_name(self.input_path, all_files)
        if doc.supplier_name:
            doc.brand = self._infer_brand(doc.supplier_name)

        # 逐个文件打标分流
        for fpath in all_files:
            asset = self._classify_asset(fpath)
            doc.assets.append(asset)

        return doc

    def _classify_asset(self, fpath: str) -> DocumentAsset:
        fname = os.path.basename(fpath)
        ext = os.path.splitext(fname)[1].lower()
        fname_no_ext = os.path.splitext(fname)[0]

        # 1. 检查是否为无关附件
        for kw in ATTACHMENT_KEYWORDS:
            if kw in fname_no_ext:
                return DocumentAsset(
                    file_path=fpath,
                    asset_type=AssetType.ATTACHMENT_IGNORE,
                    confidence=0.98,
                    ext=ext,
                    summary=f"匹配无关附件关键词 [{kw}]，自动过滤"
                )

        # 2. 检查是否为商务条款 / 工程折扣单
        for kw in TERMS_KEYWORDS:
            if kw in fname_no_ext:
                return DocumentAsset(
                    file_path=fpath,
                    asset_type=AssetType.TERMS_DISCOUNT,
                    confidence=0.95,
                    ext=ext,
                    summary=f"匹配商务条款关键词 [{kw}]"
                )

        # 3. 检查是否为报价主表
        if ext in ('.xlsx', '.xlsm', '.xls'):
            return DocumentAsset(
                file_path=fpath,
                asset_type=AssetType.QUOTE_MAIN,
                confidence=0.95,
                ext=ext,
                summary="Excel 报价主表"
            )
        elif ext == '.pdf':
            # PDF 需结合文件名进一步评估
            return DocumentAsset(
                file_path=fpath,
                asset_type=AssetType.QUOTE_MAIN,
                confidence=0.90,
                ext=ext,
                summary="PDF 报价主表"
            )
        elif ext in ('.png', '.jpg', '.jpeg', '.webp', '.bmp'):
            # 图片若是独立商务说明，已在前序规则命中；否则作为主表图片
            return DocumentAsset(
                file_path=fpath,
                asset_type=AssetType.QUOTE_MAIN,
                confidence=0.80,
                ext=ext,
                summary="图像单据报价主表"
            )

        return DocumentAsset(
            file_path=fpath,
            asset_type=AssetType.ATTACHMENT_IGNORE,
            confidence=0.99,
            ext=ext,
            summary=f"不支持或无关格式: {ext}"
        )

    def _detect_supplier_name(self, root_path: str, files: List[str]) -> str:
        """从路径名与文件名中探查供应商企业名称"""
        candidates = [os.path.basename(root_path)]
        d = os.path.dirname(os.path.abspath(root_path))
        for _ in range(2):
            candidates.append(os.path.basename(d))
            d = os.path.dirname(d)
        for f in files:
            candidates.append(os.path.basename(f))

        for c in candidates:
            m = COMPANY_RE.search(c)
            if m:
                return m.group(1)
        return ""

    def _infer_brand(self, supplier: str) -> str:
        """从供应商名称剔除组织形式与地域前缀，提取品牌首选"""
        core = re.sub(r'^(山东|河北|上海|浙江|江苏|广东|北京|天津|重庆|四川|湖北|湖南|河南|福建|江西|安徽|辽宁|吉林|黑龙江|陕西|山西|广西|贵州|云南|内蒙古|新疆|甘肃|青海|宁夏|海南)', '', supplier)
        core = re.sub(r'^(潍坊|临沂|苏州|杭州|温州|无锡|常州|宁波|嘉兴|绍兴|金华|台州|佛山|东莞|中山|保定|廊坊|唐山|沧州|邢台|邯郸)市?', '', core)
        core = re.sub(r'(股份有限公司|有限责任公司|有限公司|股份公司|集团|科技有限公司|新材料|消防器材|电气|管业|实业).*$', '', core)
        return core.strip() or supplier[:4]
