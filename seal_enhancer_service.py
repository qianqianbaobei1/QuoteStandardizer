# -*- coding: utf-8 -*-
"""
印章智能检测、旋转扶正、极坐标展开与遮挡穿透增强服务
(Seal Enhancement & Under-Seal Text Penetration Service)

核心技术能力：
1. 【公章定位与分割】：在 HSV 空间提取红色印泥，通过轮廓面积与最小外接圆精确锁定印章圆心与外圈；
2. 【中枢五角星检测与倾角标定】：提取公章中心五角星轮廓，根据五角星尖角指向计算旋转偏角 theta；
3. 【自适应旋转扶正与极坐标展开】：
   - 仿射旋转变换将公章逆时针扶正至标准正立姿态；
   - 线性极坐标展开 (Warp Polar) 将环形弧字展开为水平文字条，极大提升通用 OCR/视觉模型识别率；
4. 【红色印章脱墨与底层文字穿透 (De-inking)】：利用光学通道物理特性 (黑色全吸收 vs 红色高反射)，提取 R 通道并进行背景自适应除法，消除红章遮挡；
5. 【凸线强化滤波 (Ridge & Edge Enhancement)】：通过 CLAHE 自适应对比度均衡与反向非锐化掩模 (Unsharp Masking)，强化被遮挡文字的微弱边缘与骨架。
"""

import os
import cv2
import numpy as np
from typing import Dict, Any, Optional, Tuple, List

try:
    from universal_pdf_grid_engine import auto_orient_page_upright
except ImportError:
    def auto_orient_page_upright(img):
        return img


class SealEnhancerService:
    def __init__(self, output_debug_dir: Optional[str] = None):
        self.output_debug_dir = output_debug_dir
        if output_debug_dir and not os.path.exists(output_debug_dir):
            os.makedirs(output_debug_dir, exist_ok=True)

    def extract_red_chroma(self, bgr_img: np.ndarray) -> np.ndarray:
        """纯红度色差场 R - (G+B)/2：黑字/黑签名/灰底三通道相等故恒为 0，只有红印泥 > 0"""
        b, g, r = cv2.split(bgr_img.astype(np.float32))
        return np.clip(r - 0.5 * (g + b), 0, 255)

    def extract_red_mask(self, bgr_img: np.ndarray) -> np.ndarray:
        """在 HSV 颜色空间精确提取红色公章印泥掩膜；HSV 无命中时回退到红度色差通道（浅淡/干涸印泥）"""
        hsv = cv2.cvtColor(bgr_img, cv2.COLOR_BGR2HSV)
        mask1 = cv2.inRange(hsv, np.array([0, 50, 40]), np.array([12, 255, 255]))
        mask2 = cv2.inRange(hsv, np.array([155, 50, 40]), np.array([180, 255, 255]))
        mask = cv2.bitwise_or(mask1, mask2)
        if cv2.countNonZero(mask) >= 50:
            return mask
        return self._faint_red_mask(bgr_img)

    def _faint_red_mask(self, bgr_img: np.ndarray) -> np.ndarray:
        """浅淡印泥掩膜：色度差 >= 12（人眼几乎看不见的粉红残墨也能进候选）"""
        mask = (self.extract_red_chroma(bgr_img) >= 12).astype(np.uint8) * 255
        return cv2.morphologyEx(mask, cv2.MORPH_OPEN, np.ones((2, 2), np.uint8))

    def locate_seal(self, bgr_img: np.ndarray) -> Optional[Dict[str, Any]]:
        """定位公章位置、圆心、半径及外接裁剪框（浓章走 HSV，浅章回退色度差）"""
        info = self._locate_from_mask(bgr_img, self.extract_red_mask(bgr_img))
        if info:
            return info
        # 回退：浅淡/残缺印泥。整页泛黄纸底色度差也 >0，故额外要求候选不得铺满外接圆（公章是圈+字，填充率低）
        info = self._locate_from_mask(bgr_img, self._faint_red_mask(bgr_img))
        if not info:
            return None
        circle_area = np.pi * info["radius"] ** 2
        fill = cv2.countNonZero(info["red_mask"]) / circle_area if circle_area > 0 else 1.0
        if info["radius"] > 0.35 * min(bgr_img.shape[:2]) or fill > 0.55:
            return None
        return info

    def _locate_from_mask(self, bgr_img: np.ndarray, red_mask: np.ndarray) -> Optional[Dict[str, Any]]:
        """由红色掩膜求圆心/半径/裁剪框"""
        contours, _ = cv2.findContours(red_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        valid_cnts = [c for c in contours if cv2.contourArea(c) > 15]
        if not valid_cnts:
            return None

        # 合并所有红色连通域，找最小外接圆
        all_pts = np.vstack(valid_cnts)
        (cx, cy), radius = cv2.minEnclosingCircle(all_pts)
        cx, cy, radius = int(cx), int(cy), int(radius)

        if radius < 20:  # 过小并非公章
            return None
        h, w = bgr_img.shape[:2]
        if radius > 0.35 * min(h, w):  # 报价单整页上公章只占 10% 左右；过大多为红色表头/logo/红框被合并成大圆
            return None

        h, w = bgr_img.shape[:2]
        pad = int(radius * 0.1) + 4
        x1 = max(0, cx - radius - pad)
        y1 = max(0, cy - radius - pad)
        x2 = min(w, cx + radius + pad)
        y2 = min(h, cy + radius + pad)

        crop = bgr_img[y1:y2, x1:x2]
        center_in_crop = (cx - x1, cy - y1)

        return {
            "center": (cx, cy),
            "radius": radius,
            "crop_bbox": (x1, y1, x2, y2),
            "seal_crop": crop,
            "center_in_crop": center_in_crop,
            "red_mask": red_mask[y1:y2, x1:x2]
        }

    def estimate_star_orientation(self, seal_crop: np.ndarray, center_in_crop: Tuple[int, int], radius: int) -> float:
        """
        利用印章中心五角星判定公章朝向。
        返回需要扶正的旋转角度 (角度制，逆时针为负，顺时针为正)。
        """
        h, w = seal_crop.shape[:2]
        cx, cy = center_in_crop

        red_mask = self.extract_red_mask(seal_crop)

        # 仅取内环五角星区域 (半径 0.35 * R 以内)
        star_search_mask = np.zeros((h, w), dtype=np.uint8)
        inner_r = max(10, int(radius * 0.35))
        cv2.circle(star_search_mask, (cx, cy), inner_r, 255, -1)

        inner_star_region = cv2.bitwise_and(red_mask, star_search_mask)
        contours, _ = cv2.findContours(inner_star_region, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        if not contours:
            return 0.0

        star_cnt = max(contours, key=cv2.contourArea)
        M = cv2.moments(star_cnt)
        if M["m00"] <= 0:
            return 0.0

        scx = int(M["m10"] / M["m00"])
        scy = int(M["m01"] / M["m00"])

        # 找五角星凸包极值尖角
        hull = cv2.convexHull(star_cnt)
        pts = hull.reshape(-1, 2)
        dists = np.hypot(pts[:, 0] - scx, pts[:, 1] - scy)
        angles = np.arctan2(pts[:, 1] - scy, pts[:, 0] - scx) * 180.0 / np.pi

        top_peak_idx = np.argmax(dists)
        angle_deg = angles[top_peak_idx]

        # 标准正向五角星顶尖指向 12 点钟方向 (-90 度)
        rot_angle = angle_deg + 90.0
        return rot_angle

    def rotate_seal_upright(self, seal_crop: np.ndarray, center_in_crop: Tuple[int, int], rot_angle: float) -> np.ndarray:
        """将倾斜/倒置公章以中心为轴仿射旋转扶正"""
        h, w = seal_crop.shape[:2]
        cx, cy = center_in_crop
        M = cv2.getRotationMatrix2D((cx, cy), -rot_angle, 1.0)
        upright = cv2.warpAffine(seal_crop, M, (w, h), borderMode=cv2.BORDER_CONSTANT, borderValue=(255, 255, 255))
        return upright

    def unwrap_seal_polar(self, seal_crop: np.ndarray, center_in_crop: Tuple[int, int], radius: int) -> np.ndarray:
        """
        极坐标环形展开 (Linear Polar Unwrapping)：
        将公章外圈的环形公司名称文字沿圆周展开为长条矩形文本带。
        """
        cx, cy = center_in_crop
        max_r = int(radius * 1.15)
        # dsize=(宽, 高) = (半径轴, 角度轴)：半径轴给 max_r 个采样即可，角度轴必须给足 2πR，否则环形字被欠采样糊成一团
        polar = cv2.warpPolar(seal_crop, (max_r, int(2 * np.pi * max_r)), (cx, cy), max_r, cv2.WARP_POLAR_LINEAR)
        # 逆时针旋转 90° 让角度轴水平、字头正立；旋转后行号 0 = 最大半径，故按半径反推行区间
        band = cv2.rotate(polar, cv2.ROTATE_90_COUNTERCLOCKWISE)
        h_band = band.shape[0]
        r_in, r_out = 0.52 * radius / max_r, 0.99 * radius / max_r
        y0 = max(0, int(h_band * (1 - r_out)))
        y1 = min(h_band, int(h_band * (1 - r_in)))
        return band[y0:y1, :]

    def penetrate_and_enhance_text(self, bgr_img: np.ndarray) -> np.ndarray:
        """
        【公章遮挡穿透与凸线强化】：
        1. 利用光学多通道解耦，提取 R 通道（红章高反射淡化，黑字强吸收保留）；
        2. 背景归一化消除残余印油与纸面不均；
        3. CLAHE 自适应直方图均衡化强化文字边缘；
        4. 反向非锐化掩模 (Unsharp Masking) 实现“凸线强化”，使被印章覆盖的文字边缘鲜明坚挺。
        """
        b, g, r = cv2.split(bgr_img)

        # 1. 背景除法 / 阴影归一化
        bg = cv2.morphologyEx(r, cv2.MORPH_DILATE, cv2.getStructuringElement(cv2.MORPH_RECT, (15, 15)))
        diff = cv2.absdiff(r, bg)
        norm = cv2.normalize(diff, None, alpha=0, beta=255, norm_type=cv2.NORM_MINMAX, dtype=cv2.CV_8UC1)
        inv = 255 - norm

        # 2. 对比度受限自适应直方图均衡化 (CLAHE)
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
        enhanced = clahe.apply(inv)

        # 3. 凸线强化 (Unsharp Masking 锐化)
        gaussian = cv2.GaussianBlur(enhanced, (0, 0), 2.0)
        unsharp = cv2.addWeighted(enhanced, 1.8, gaussian, -0.8, 0)

        # 4. 伽马增强突出深色墨迹
        inv_gamma = 1.0 / 1.2
        table = np.array([((i / 255.0) ** inv_gamma) * 255 for i in np.arange(0, 256)]).astype("uint8")
        out = cv2.LUT(unsharp, table)

        return out

    def process_image(self, bgr_img: np.ndarray) -> Dict[str, Any]:
        """
        端到端全量处理流水线：
        返回包含扶正公章、极坐标文本带、去印章穿透底稿的全套产物。
        """
        seal_info = self.locate_seal(bgr_img)
        penetrated_img = self.penetrate_and_enhance_text(bgr_img)

        result: Dict[str, Any] = {
            "has_seal": seal_info is not None,
            "penetrated_text_img": penetrated_img,
            "upright_seal_img": None,
            "polar_strip_img": None,
            "rotation_angle": 0.0,
            "seal_bbox": None
        }

        if seal_info:
            seal_crop = seal_info["seal_crop"]
            center = seal_info["center_in_crop"]
            radius = seal_info["radius"]

            rot_angle = self.estimate_star_orientation(seal_crop, center, radius)
            upright_seal = self.rotate_seal_upright(seal_crop, center, rot_angle)
            # 必须对已经扶正后的公章做极坐标展开，防止倒置印章在未扶正时展开将企业名称拦腰切断
            polar_strip = self.unwrap_seal_polar(upright_seal, center, radius)

            result.update({
                "upright_seal_img": upright_seal,
                "polar_strip_img": polar_strip,
                "rotation_angle": rot_angle,
                "seal_bbox": seal_info["crop_bbox"],
                "center_in_crop": center,
                "radius": radius
            })

            if self.output_debug_dir:
                cv2.imwrite(os.path.join(self.output_debug_dir, "debug_upright_seal.png"), upright_seal)
                cv2.imwrite(os.path.join(self.output_debug_dir, "debug_polar_strip.png"), polar_strip)
                cv2.imwrite(os.path.join(self.output_debug_dir, "debug_penetrated_text.png"), penetrated_img)

        return result

    # 别名兼容
    process_document_image = process_image

    def amplify_faint_red_seal_and_remove_black(self, bgr_crop: np.ndarray) -> np.ndarray:
        """
        【去黑显红光学通道】：
        红度色差 R-0.5(G+B) 上，盖在公章上的黑色印刷字（“公司盖章：”）、黑色表格线、粗黑手写签名
        与白纸背景的色差都恒为 0 → 全部物理剥离；再对极浅残缺印泥做分位数拉伸 + Gamma(0.45) 增益，
        输出白底黑字的高对比度显影图，便于 VLM/OCR 读出公司全称。
        """
        if bgr_crop is None or bgr_crop.size == 0:
            return bgr_crop
        chroma = self.extract_red_chroma(bgr_crop)
        chroma[chroma < 6.0] = 0.0
        peak = max(float(np.percentile(chroma, 99.5)), 10.0)
        norm = np.clip(chroma / peak, 0.0, 1.0)
        # 抬底：纸张泛黄/印泥晕开的低幅底噪统一归零，只留笔画
        norm = np.clip((norm - 0.18) / 0.82, 0.0, 1.0)
        ink_dark_on_white = (255.0 - (norm ** 0.75) * 255.0).astype(np.uint8)
        return cv2.cvtColor(cv2.medianBlur(ink_dark_on_white, 3), cv2.COLOR_GRAY2BGR)

    def build_seal_enhancement_banner(self, full_bgr: np.ndarray) -> Optional[np.ndarray]:
        """
        生成认章拼图条：[左 扶正原色章] + [右 去黑签字/浅印泥显影章]，下方再拼极坐标拉直的公司名条带。
        页面无公章时返回 None。
        """
        if full_bgr is None or full_bgr.size == 0:
            return None
        res = self.process_document_image(full_bgr)
        upright = res.get("upright_seal_img")
        if not res.get("has_seal") or upright is None:
            return None

        amplified = self.amplify_faint_red_seal_and_remove_black(upright)
        h_s, w_s = upright.shape[:2]
        if h_s > 360:  # 控制拼图体积，避免主图被压缩到看不清
            scale = 360.0 / h_s
            size = (max(1, int(w_s * scale)), 360)
            upright = cv2.resize(upright, size, interpolation=cv2.INTER_AREA)
            amplified = cv2.resize(amplified, size, interpolation=cv2.INTER_AREA)
        side_by_side = np.hstack([upright, amplified])

        center = res.get("center_in_crop") or (w_s // 2, h_s // 2)
        radius = int(res.get("radius") or min(w_s, h_s) // 2)
        if int(res.get("radius") or 0) < 100:
            return side_by_side  # 小章极坐标采样不足，拉直条只会是一堆噪点，不如不给
        if h_s != upright.shape[0]:
            k = upright.shape[0] / float(h_s)
            center = (int(center[0] * k), int(center[1] * k))
            radius = max(8, int(radius * k))
        clean_polar = self.unwrap_seal_polar(amplified, center, radius)
        if clean_polar is None or clean_polar.size == 0:
            return side_by_side
        target_w = max(side_by_side.shape[1], clean_polar.shape[1])
        panels = cv2.resize(side_by_side, (target_w, max(120, int(side_by_side.shape[0] * target_w / side_by_side.shape[1]))), interpolation=cv2.INTER_CUBIC)
        target_h = max(40, int(clean_polar.shape[0] * target_w / max(clean_polar.shape[1], 1)))
        polar_resized = cv2.resize(clean_polar, (target_w, target_h), interpolation=cv2.INTER_CUBIC)
        return np.vstack([panels, polar_resized])

    def balance_illumination_and_orient(self, bgr_img: np.ndarray, do_orient: bool = True) -> np.ndarray:
        """
        手机翻拍图与扫描件预处理（自动摆正 + LAB 亮度通道阴影均衡）：
        重要：必须在 LAB 色彩空间仅对 L (亮度) 通道做背景除法，100% 保留 A/B 色彩通道，
        绝不能转成 GRAY 黑白图，否则会抹掉红色公章的色彩信息导致后续 R 通道红章穿透失效！
        """
        if bgr_img is None or bgr_img.size == 0:
            return bgr_img
        try:
            oriented = auto_orient_page_upright(bgr_img) if do_orient else bgr_img
            h, w = oriented.shape[:2]
            lab = cv2.cvtColor(oriented, cv2.COLOR_BGR2LAB)
            l_ch, a_ch, b_ch = cv2.split(lab)
            ksize = 51 if min(h, w) > 300 else 21
            bg = cv2.GaussianBlur(l_ch, (ksize, ksize), 0)
            norm_l = cv2.divide(l_ch, bg, scale=255)
            return cv2.cvtColor(cv2.merge([norm_l, a_ch, b_ch]), cv2.COLOR_LAB2BGR)
        except Exception:
            return bgr_img

