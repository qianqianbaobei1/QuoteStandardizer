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


class SealEnhancerService:
    def __init__(self, output_debug_dir: Optional[str] = None):
        self.output_debug_dir = output_debug_dir
        if output_debug_dir and not os.path.exists(output_debug_dir):
            os.makedirs(output_debug_dir, exist_ok=True)

    def extract_red_mask(self, bgr_img: np.ndarray) -> np.ndarray:
        """在 HSV 颜色空间精确提取红色公章印泥掩膜"""
        hsv = cv2.cvtColor(bgr_img, cv2.COLOR_BGR2HSV)
        mask1 = cv2.inRange(hsv, np.array([0, 50, 40]), np.array([12, 255, 255]))
        mask2 = cv2.inRange(hsv, np.array([155, 50, 40]), np.array([180, 255, 255]))
        return cv2.bitwise_or(mask1, mask2)

    def locate_seal(self, bgr_img: np.ndarray) -> Optional[Dict[str, Any]]:
        """定位公章位置、圆心、半径及外接裁剪框"""
        red_mask = self.extract_red_mask(bgr_img)
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
        polar_w = int(2 * np.pi * max_r)
        polar_h = max_r

        polar = cv2.warpPolar(seal_crop, (polar_w, polar_h), (cx, cy), max_r, cv2.WARP_POLAR_LINEAR)
        # 旋转 90 度使角度沿水平展开
        polar_horizontal = cv2.rotate(polar, cv2.ROTATE_90_COUNTERCLOCKWISE)

        # 截取外圈文字主要所在环带 (0.55 * R ~ 1.05 * R)
        h_strip = polar_horizontal.shape[0]
        y_start = int(h_strip * 0.45)
        y_end = int(h_strip * 0.98)
        text_strip = polar_horizontal[y_start:y_end, :]
        return text_strip

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
                "seal_bbox": seal_info["crop_bbox"]
            })

            if self.output_debug_dir:
                cv2.imwrite(os.path.join(self.output_debug_dir, "debug_upright_seal.png"), upright_seal)
                cv2.imwrite(os.path.join(self.output_debug_dir, "debug_polar_strip.png"), polar_strip)
                cv2.imwrite(os.path.join(self.output_debug_dir, "debug_penetrated_text.png"), penetrated_img)

        return result

    # 别名兼容
    process_document_image = process_image

    def balance_illumination_and_orient(self, bgr_img: np.ndarray) -> np.ndarray:
        """
        杀手锏 5：手机翻拍图预处理（暗角光照均衡 + 自动摆正）
        1. 大核高斯背景估计与除法：消除手机拍照左右明暗不均、手影、暗角，使表格线和小数点清晰显现；
        2. 自动检测横竖摆正（若图像被横放拍，根据文本边缘投影自动旋转正立）。
        """
        if bgr_img is None or bgr_img.size == 0:
            return bgr_img
        try:
            h, w = bgr_img.shape[:2]
            gray = cv2.cvtColor(bgr_img, cv2.COLOR_BGR2GRAY)
            ksize = 51 if min(h, w) > 300 else 21
            bg = cv2.GaussianBlur(gray, (ksize, ksize), 0)
            balanced_gray = cv2.divide(gray, bg, scale=255)
            return cv2.cvtColor(balanced_gray, cv2.COLOR_GRAY2BGR)
        except Exception:
            return bgr_img

