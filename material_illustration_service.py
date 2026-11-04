# -*- coding: utf-8 -*-
"""
材料配图服务 (Material Illustration Service)
- 对应 工料机类别.xlsx Sheet3: 材料配图接口
  1. 创建材料配图任务: POST https://ugw.m2.com.cn/agents-toolkit/api/material/illustrations
  2. 获取材料配图任务结果: GET https://ugw.m2.com.cn/agents-toolkit/api/material/illustration-tasks?taskId={taskId}
- 核心职责：
  1. 通过材料名称 (materialName) 及规格/品牌在配图接口中批量检索配图
  2. 轮询获取任务结果，获得对应图片链接 (pictureUrl)、图片文件名 (pictureName)、目录名 (folderName)
  3. 持久化缓存匹配结果，避免重复请求
  4. 将匹配到的图片链接及元数据回填至 CanonicalMaterialItem 或 Excel 单元格
"""
import os
import re
import json
import time
import threading
import urllib.request
import urllib.error
from typing import List, Dict, Any, Optional

from config import (
    ILLUSTRATION_API_BASE_URL,
    ILLUSTRATION_TASK_CREATE_URL,
    ILLUSTRATION_TASK_QUERY_URL,
    ENABLE_MATERIAL_ILLUSTRATION,
    DOWNLOAD_ILLUSTRATION_IMAGES
)

CACHE_FILE = os.path.join(os.path.dirname(__file__), "material_illustration_cache.json")
_CACHE_LOCK = threading.Lock()
_ILLUSTRATION_CACHE: Dict[str, Dict[str, Any]] = {}
_DIRTY = 0
SAVE_EVERY = 10


def load_cache():
    global _ILLUSTRATION_CACHE
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                _ILLUSTRATION_CACHE = json.load(f)
        except Exception:
            _ILLUSTRATION_CACHE = {}


def save_cache(force: bool = False):
    """原子落盘保存缓存，多线程安全"""
    global _DIRTY
    with _CACHE_LOCK:
        _DIRTY += 1
        if not force and SAVE_EVERY > 0 and _DIRTY % SAVE_EVERY != 0:
            return
        try:
            tmp = f"{CACHE_FILE}.tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(_ILLUSTRATION_CACHE, f, ensure_ascii=False, indent=2)
            os.replace(tmp, CACHE_FILE)
        except Exception:
            pass


load_cache()


def clean_material_name(name: str) -> str:
    """归一化材料名称作为检索主键与缓存键"""
    if not name:
        return ""
    s = str(name).strip()
    # 去除多余括号及前后空白
    s = re.sub(r'[\r\n\t]+', ' ', s)
    return s.strip()


def _download_image_bytes(url: str, timeout: int = 5) -> Optional[bytes]:
    """安全下载远程图片二进制内容，带超时与容错保护"""
    if not url or not (url.startswith("http://") or url.startswith("https://")):
        return None
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "QuoteStandardizer/1.0"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            if resp.status == 200:
                return resp.read()
    except Exception:
        pass
    return None


class MaterialIllustrationService:
    def __init__(self,
                 base_url: str = ILLUSTRATION_API_BASE_URL,
                 create_url: str = ILLUSTRATION_TASK_CREATE_URL,
                 query_url: str = ILLUSTRATION_TASK_QUERY_URL,
                 enable: bool = ENABLE_MATERIAL_ILLUSTRATION,
                 download_images: bool = DOWNLOAD_ILLUSTRATION_IMAGES):
        self.base_url = base_url
        self.create_url = create_url
        self.query_url = query_url
        self.enable = enable
        self.download_images = download_images

    def fetch_illustrations_batch(self,
                                  materials: List[Dict[str, Any]],
                                  company_name: str = "标准报价系统",
                                  timeout: int = 15,
                                  max_poll_seconds: int = 15,
                                  download_bytes: bool = False) -> Dict[str, Dict[str, Any]]:
        """
        批量为材料检索配图并返回字典映射：
        返回结果格式: {uid: {"pictureUrl": str, "pictureName": str, "folderName": str, "image_bytes": bytes or None}}
        """
        if not self.enable or not materials or not self.create_url:
            return {}

        # 兼容处理：若传入单纯字符串列表，自动标准化为带 uid 字典
        normalized_materials: List[Dict[str, Any]] = []
        for idx, m in enumerate(materials):
            if isinstance(m, str):
                normalized_materials.append({"uid": str(idx), "materialName": m})
            elif isinstance(m, dict):
                normalized_materials.append(m)
        materials = normalized_materials

        company_name = str(company_name or "标准报价系统").strip() or "标准报价系统"
        results: Dict[str, Dict[str, Any]] = {}
        uncached_materials: List[Dict[str, Any]] = []

        # 1. 优先查本地持久化缓存
        for mat in materials:
            uid = str(mat.get("uid") or "").strip()
            name = clean_material_name(mat.get("materialName") or "")
            if not uid or not name:
                continue

            cached = _ILLUSTRATION_CACHE.get(name)
            if cached is not None:
                if cached.get("matched") and cached.get("pictureUrl"):
                    hit = dict(cached)
                    if download_bytes or self.download_images:
                        hit["image_bytes"] = _download_image_bytes(hit["pictureUrl"])
                    results[uid] = hit
                # 若缓存标记为未命中(matched=False)，则直接跳过，不再重复请求网络
            else:
                uncached_materials.append(mat)

        if not uncached_materials:
            return results

        # 2. 对未缓存的材料按单次最多 100 条分批请求
        chunk_size = 100
        for i in range(0, len(uncached_materials), chunk_size):
            chunk = uncached_materials[i:i + chunk_size]
            api_materials = []
            for m in chunk:
                item_payload = {
                    "uid": str(m["uid"]),
                    "materialName": clean_material_name(m.get("materialName") or ""),
                }
                if m.get("categoryCode"):
                    item_payload["categoryCode"] = str(m["categoryCode"])
                if m.get("brandName"):
                    item_payload["brandName"] = str(m["brandName"])
                if m.get("specification"):
                    item_payload["specification"] = str(m["specification"])
                if m.get("unit"):
                    item_payload["unit"] = str(m["unit"])
                api_materials.append(item_payload)

            payload = {
                "companyName": company_name,
                "materials": api_materials
            }

            try:
                # 【1】创建配图任务
                req = urllib.request.Request(
                    self.create_url,
                    data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
                    headers={"Content-Type": "application/json", "User-Agent": "QuoteStandardizer"}
                )
                with urllib.request.urlopen(req, timeout=timeout) as resp:
                    resp_json = json.loads(resp.read().decode("utf-8"))

                resp_data = resp_json.get("data") or {}
                task_id = resp_data.get("taskId")
                if not task_id:
                    continue

                # 【2】轮询任务结果
                start_time = time.time()
                matched_pictures = []
                while (time.time() - start_time) < max_poll_seconds:
                    time.sleep(1.0)
                    query_url = f"{self.query_url}?taskId={task_id}"
                    q_req = urllib.request.Request(query_url, headers={"User-Agent": "QuoteStandardizer"})
                    try:
                        with urllib.request.urlopen(q_req, timeout=timeout) as q_resp:
                            q_json = json.loads(q_resp.read().decode("utf-8"))
                        q_data = q_json.get("data") or {}
                        status = q_data.get("status")
                        if status == "SUCCEEDED":
                            matched_pictures = q_data.get("pictures", [])
                            break
                        elif status == "FAILED":
                            break
                    except Exception:
                        pass

                # 解析命中结果并更新缓存
                matched_uids = set()
                uid_to_name = {str(m["uid"]): clean_material_name(m.get("materialName") or "") for m in chunk}

                for pic in matched_pictures:
                    p_uid = str(pic.get("uid") or "").strip()
                    matched_uids.add(p_uid)
                    m_name = uid_to_name.get(p_uid, "")
                    hit_info = {
                        "matched": True,
                        "pictureUrl": pic.get("pictureUrl") or "",
                        "pictureName": pic.get("pictureName") or "",
                        "folderName": pic.get("folderName") or "",
                        "templateName": pic.get("templateName") or "",
                        "matchType": pic.get("matchType") or ""
                    }
                    if m_name:
                        _ILLUSTRATION_CACHE[m_name] = hit_info

                    if download_bytes or self.download_images:
                        hit_info["image_bytes"] = _download_image_bytes(hit_info["pictureUrl"])
                    results[p_uid] = hit_info

                # 未命中的材料记录为 matched=False，防止重复网络空查
                for m in chunk:
                    m_uid = str(m["uid"])
                    if m_uid not in matched_uids:
                        m_name = uid_to_name.get(m_uid, "")
                        if m_name and m_name not in _ILLUSTRATION_CACHE:
                            _ILLUSTRATION_CACHE[m_name] = {
                                "matched": False,
                                "pictureUrl": "",
                                "pictureName": "",
                                "folderName": ""
                            }

                save_cache()
            except Exception as e:
                # 异常隔离：接口网络超时或故障时不阻断主处理流程
                print(f"[Warning] 材料配图接口调用异常: {e}")

        save_cache(force=True)
        return results

    def fetch_illustration_by_name(self,
                                   material_name: str,
                                   company_name: str = "标准报价系统",
                                   **kwargs) -> Optional[Dict[str, Any]]:
        """单物料名称快捷检索配图"""
        name = clean_material_name(material_name)
        if not name:
            return None
        mat = {"uid": "single_01", "materialName": name}
        mat.update(kwargs)
        res = self.fetch_illustrations_batch([mat], company_name=company_name)
        return res.get("single_01")


# 全局默认单例
default_illustration_service = MaterialIllustrationService()


def match_illustrations_for_canonical_items(items: list,
                                            company_name: str = None,
                                            download_bytes: bool = False) -> int:
    """
    为 CanonicalMaterialItem 列表批量检索并匹配材料图片链接与元数据。
    返回成功命中配图的物料数量。
    """
    if not items or not ENABLE_MATERIAL_ILLUSTRATION:
        return 0

    company = company_name or (items[0].supplier if items else "标准报价系统") or "标准报价系统"
    
    # 按照材料标准名称去重查询，避免全国展开时数千行重复请求拖慢数十倍
    unique_materials = {}
    item_key_map = {}
    for idx, it in enumerate(items):
        uid = it.quote_id or f"item_{idx}"
        mat_name = clean_material_name(it.category_name or it.raw_spec or "")
        key = (mat_name, it.category_code or "", it.brand or "")
        item_key_map[uid] = key
        if key not in unique_materials and mat_name:
            unique_materials[key] = {
                "uid": f"uniq_{len(unique_materials)}",
                "materialName": mat_name,
                "categoryCode": it.category_code or "",
                "brandName": it.brand or "",
                "specification": it.raw_spec or "",
                "unit": it.unit or ""
            }

    uniq_list = list(unique_materials.values())
    matched_uniq = default_illustration_service.fetch_illustrations_batch(
        uniq_list,
        company_name=company,
        download_bytes=download_bytes
    )

    # 映射回 key -> hit
    key_hit_map = {}
    for key, udata in unique_materials.items():
        hit = matched_uniq.get(udata["uid"])
        if hit:
            key_hit_map[key] = hit

    matched_count = 0
    for idx, it in enumerate(items):
        uid = it.quote_id or f"item_{idx}"
        key = item_key_map.get(uid)
        hit = key_hit_map.get(key)
        if hit and hit.get("pictureUrl"):
            it.picture_url = hit["pictureUrl"]
            if not getattr(it, "picture_urls", None):
                it.picture_urls = []
            if hit["pictureUrl"] not in it.picture_urls:
                it.picture_urls.append(hit["pictureUrl"])

            if hit.get("pictureName"):
                it.illustration_picture_name = hit["pictureName"]

            if hit.get("folderName"):
                it.illustration_folder = hit["folderName"]

            if hit.get("image_bytes") and not it.image_bytes_list:
                it.image_bytes_list.append(hit["image_bytes"])

            matched_count += 1

    return matched_count


def match_illustrations_for_openpyxl_sheet(ws, company_name: str = "标准报价系统") -> int:
    """
    为 openpyxl 工作表（如 zip 解包后的 报价数据.xlsx）批量按材料名称检索配图，
    并把对应的图片链接写回单元格与超链接中。
    """
    if not ws or not ENABLE_MATERIAL_ILLUSTRATION:
        return 0

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    name_col = headers.index('名称规格') + 1 if '名称规格' in headers else None
    if not name_col:
        return 0

    img_col = headers.index('图片') + 1 if '图片' in headers else None
    img_file_col = headers.index('厂家报价图片文件名') + 1 if '厂家报价图片文件名' in headers else None
    folder_col = headers.index('图片文件夹') + 1 if '图片文件夹' in headers else None
    cat_col = headers.index('类别编码') + 1 if '类别编码' in headers else None
    brand_col = headers.index('品牌') + 1 if '品牌' in headers else None
    spec_col = headers.index('厂家原始规格参数') + 1 if '厂家原始规格参数' in headers else None
    unit_col = headers.index('单位') + 1 if '单位' in headers else None

    materials_payload = []
    row_mapping = {}

    for r in range(2, ws.max_row + 1):
        name_val = str(ws.cell(r, name_col).value or "").strip()
        if not name_val:
            continue
        uid = f"r_{r}"
        row_mapping[uid] = r
        materials_payload.append({
            "uid": uid,
            "materialName": name_val,
            "categoryCode": str(ws.cell(r, cat_col).value or "").strip() if cat_col else "",
            "brandName": str(ws.cell(r, brand_col).value or "").strip() if brand_col else "",
            "specification": str(ws.cell(r, spec_col).value or "").strip() if spec_col else "",
            "unit": str(ws.cell(r, unit_col).value or "").strip() if unit_col else ""
        })

    if not materials_payload:
        return 0

    matched_dict = default_illustration_service.fetch_illustrations_batch(
        materials_payload,
        company_name=company_name
    )

    matched_count = 0
    from openpyxl.styles import Font
    link_font = Font(name='宋体', size=10, color='0563C1', underline='single')

    for uid, hit in matched_dict.items():
        r = row_mapping.get(uid)
        if not r or not hit or not hit.get("pictureUrl"):
            continue

        pic_url = hit["pictureUrl"]
        pic_name = hit.get("pictureName") or ""
        folder_name = hit.get("folderName") or ""

        # 写入图片列链接
        if img_col:
            cell = ws.cell(row=r, column=img_col, value=pic_url)
            cell.hyperlink = pic_url
            cell.font = link_font

        # 厂家报价图片文件名列：放底稿图片不放文字，单元格文字保持为空
        if img_file_col:
            ws.cell(row=r, column=img_file_col).value = None

        # 补全图片文件夹
        if folder_col and folder_name:
            cur_folder_val = str(ws.cell(r, folder_col).value or "").strip()
            if not cur_folder_val:
                ws.cell(row=r, column=folder_col, value=folder_name)

        matched_count += 1

    return matched_count
