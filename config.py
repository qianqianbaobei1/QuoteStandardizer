# -*- coding: utf-8 -*-
"""
全局配置
- 所有路径、密钥、模型均来自环境变量 / .env，代码中不写任何个人路径或默认业务数据
- 分类编码只来自用户提供的标准编码文件 (CATEGORY_CODE_FILE)，不内置任何未经核实的编码
- 省级行政区简称映射为公开的行政区划常识
"""
import csv
import os
import re
from dotenv import load_dotenv

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
# 固定从项目目录加载 .env，避免因启动目录不同而读不到 Key
load_dotenv(os.path.join(CURRENT_DIR, ".env"))

# 密钥只允许来自环境变量 / .env，严禁在代码中写默认值
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
DEEPSEEK_BASE_URL = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com/v1")
DEEPSEEK_MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-chat")
# 扫描件 / 商务单图片识别用的多模态模型；未单独配置时使用 DeepSeek (deepseek-chat 支持图片输入)
VISION_API_KEY = os.getenv("VISION_API_KEY") or DEEPSEEK_API_KEY
VISION_BASE_URL = os.getenv("VISION_BASE_URL") or DEEPSEEK_BASE_URL
VISION_MODEL = os.getenv("VISION_MODEL") or DEEPSEEK_MODEL
# 扫描件是否让视觉模型独立读两遍做交叉校验 (默认开启；关闭可省一半调用量)
VISION_DOUBLE_CHECK = os.getenv("VISION_DOUBLE_CHECK", "1") not in ("0", "false", "False", "")
# 扫描件识别增强的 gamma (>1 压暗浅色笔画)
VISION_GAMMA = float(os.getenv("VISION_GAMMA", "1.8"))
# 大页面切块识别配置
VISION_MAX_ROWS_PER_TILE = int(os.getenv("VISION_MAX_ROWS_PER_TILE", "40"))
VISION_TILE_OVERLAP_ROWS = int(os.getenv("VISION_TILE_OVERLAP_ROWS", "2"))

# 表格结构 / 行类型判定方式 (试点)：
#   llm     = 现有做法 (DeepSeek 推导表头 + 启发式行过滤)，默认
#   jev     = TypeSafe Jev 判定表头、单价列、单位列、表格类型和每行类型
#   compare = 输出仍按现有做法，同时调用 Jev，并把两者的判定差异写入「结构判定对比」Sheet
STRUCTURE_ENGINE = os.getenv("STRUCTURE_ENGINE", "llm").strip().lower()
TYPESAFE_API_KEY = os.getenv("TYPESAFE_API_KEY", "")
TYPESAFE_BASE_URL = os.getenv("TYPESAFE_BASE_URL", "https://api.typesafe.ai/v1/systemone")
TYPESAFE_MODEL = os.getenv("TYPESAFE_MODEL", "jev-latest")
# Jev 判定置信度低于此值时写入待复核说明 (需用真实样例校准)
JEV_MIN_CONFIDENCE = float(os.getenv("JEV_MIN_CONFIDENCE", "0.5"))

# LLM 调用重试次数
LLM_MAX_RETRIES = int(os.getenv("LLM_MAX_RETRIES", "2"))

DEFAULT_OUTPUT_DIR = os.getenv("OUTPUT_DIR") or os.path.join(CURRENT_DIR, "deliveries")
# 输出模板 (取第 1 行作为表头)；未配置时使用 builder.STANDARD_HEADERS
DEFAULT_TEMPLATE_PATH = os.getenv("TEMPLATE_PATH", "")
# 标准分类编码文件 (.csv / .xlsx，前两列为 名称、编码)；未配置时不做编码库匹配
CATEGORY_CODE_FILE = os.getenv("CATEGORY_CODE_FILE", "")

# 省级行政区全称、简称与常用别名
PROVINCE_ABBR_MAP = {
    '京': '北京市', '北京': '北京市', '北京市': '北京市',
    '津': '天津市', '天津': '天津市', '天津市': '天津市',
    '冀': '河北省', '河北': '河北省', '河北省': '河北省',
    '晋': '山西省', '山西': '山西省', '山西省': '山西省',
    '蒙': '内蒙古自治区', '内蒙': '内蒙古自治区', '内蒙古': '内蒙古自治区',
    '辽': '辽宁省', '辽宁': '辽宁省', '辽宁省': '辽宁省',
    '吉': '吉林省', '吉林': '吉林省', '吉林省': '吉林省',
    '黑': '黑龙江省', '黑龙江': '黑龙江省',
    '沪': '上海市', '上海': '上海市', '上海市': '上海市',
    '苏': '江苏省', '江苏': '江苏省', '江苏省': '江苏省',
    '浙': '浙江省', '浙江': '浙江省', '浙江省': '浙江省',
    '皖': '安徽省', '安徽': '安徽省', '安徽省': '安徽省',
    '闽': '福建省', '福建': '福建省', '福建省': '福建省',
    '赣': '江西省', '江西': '江西省', '江西省': '江西省',
    '鲁': '山东省', '山东': '山东省', '山东省': '山东省',
    '豫': '河南省', '河南': '河南省', '河南省': '河南省',
    '鄂': '湖北省', '湖北': '湖北省', '湖北省': '湖北省',
    '湘': '湖南省', '湖南': '湖南省', '湖南省': '湖南省',
    '粤': '广东省', '广东': '广东省', '广东省': '广东省',
    '桂': '广西壮族自治区', '广西': '广西壮族自治区',
    '琼': '海南省', '海南': '海南省', '海南省': '海南省',
    '渝': '重庆市', '重庆': '重庆市', '重庆市': '重庆市',
    '川': '四川省', '蜀': '四川省', '四川': '四川省', '四川省': '四川省',
    '黔': '贵州省', '贵': '贵州省', '贵州': '贵州省', '贵州省': '贵州省',
    '滇': '云南省', '云': '云南省', '云南': '云南省', '云南省': '云南省',
    '藏': '西藏自治区', '西藏': '西藏自治区',
    '陕': '陕西省', '秦': '陕西省', '陕西': '陕西省', '陕西省': '陕西省',
    '甘': '甘肃省', '陇': '甘肃省', '甘肃': '甘肃省', '甘肃省': '甘肃省',
    '青': '青海省', '青海': '青海省', '青海省': '青海省',
    '宁': '宁夏回族自治区', '宁夏': '宁夏回族自治区',
    '新': '新疆维吾尔自治区', '新疆': '新疆维吾尔自治区',
    '港': '香港特别行政区', '澳': '澳门特别行政区', '台': '台湾省'
}

# 视为“不限地区”的写法
NATIONWIDE_TOKENS = {'全国', '全国范围', '不限', '全国通用'}


def _load_category_codes(path: str) -> dict:
    """读取标准分类编码文件：前两列为 名称、编码；首行若为表头自动跳过"""
    if not path:
        return {}
    if not os.path.exists(path):
        raise FileNotFoundError(f"CATEGORY_CODE_FILE 不存在: {path}")
    rows = []
    if path.lower().endswith(('.xlsx', '.xlsm')):
        import openpyxl
        ws = openpyxl.load_workbook(path, read_only=True, data_only=True).active
        rows = [r[:2] for r in ws.iter_rows(values_only=True)]
    else:
        with open(path, encoding='utf-8-sig', newline='') as f:
            rows = [r[:2] for r in csv.reader(f)]
    codes = {}
    for r in rows:
        if len(r) < 2 or r[0] in (None, '') or r[1] in (None, ''):
            continue
        name, code = str(r[0]).strip(), str(r[1]).strip()
        if re.fullmatch(r'\d+', code):
            codes[name] = code
    return codes


CATEGORY_CODE_MAP = _load_category_codes(CATEGORY_CODE_FILE)


def normalize_province(text) -> str:
    """把 '河北' / '冀' / '河北省>石家庄市' 统一为标准省级全称；无法识别时原样返回"""
    if text is None:
        return ''
    s = str(text).strip()
    if '>' in s:
        s = s.split('>')[0].strip()
    if not s:
        return ''
    if s in NATIONWIDE_TOKENS:
        return '全国'
    if s in PROVINCE_ABBR_MAP:
        return PROVINCE_ABBR_MAP[s]
    # 前缀匹配 (如 '河北省石家庄市')，优先最长的别名
    for alias in sorted(PROVINCE_ABBR_MAP, key=len, reverse=True):
        if len(alias) >= 2 and s.startswith(alias):
            return PROVINCE_ABBR_MAP[alias]
    return s


def resolve_category_code(name: str, llm_code=None):
    """
    分类编码：只认标准编码文件中的编码。
    返回 (code, source)：'dict' = 编码库命中；'llm' = 模型给出、未经编码库核验；'' = 无法确定 (code 为空)
    """
    name = (name or '').strip()
    if name in CATEGORY_CODE_MAP:
        return CATEGORY_CODE_MAP[name], 'dict'
    # 中文物料名的中心词在末尾：只有以编码库词条结尾才算命中，多个命中取最长
    hits = [k for k in CATEGORY_CODE_MAP if name.endswith(k)]
    if hits:
        return CATEGORY_CODE_MAP[max(hits, key=len)], 'dict'
    code = str(llm_code or '').strip()
    if re.fullmatch(r'\d+', code):
        return code, 'llm'
    return '', ''


def expand_region(text) -> list:
    """'京津冀+山东' → ['北京市','天津市','河北省','山东省']；'全国' → ['全国']；无法识别返回 []"""
    s = re.sub(r'\s', '', str(text or ''))
    if not s:
        return []
    if any(t in s for t in NATIONWIDE_TOKENS):
        return ['全国']
    found = []
    for part in re.split(r'[+、,，/;；和及]', s):
        if not part:
            continue
        if normalize_province(part) in PROVINCE_ABBR_MAP.values():
            found.append(normalize_province(part))
            continue
        # 逐字/逐词拆解连写简称，如 “京津冀”
        i = 0
        while i < len(part):
            for L in (3, 2, 1):
                seg = part[i:i + L]
                if seg in PROVINCE_ABBR_MAP:
                    found.append(PROVINCE_ABBR_MAP[seg])
                    i += L
                    break
            else:
                i += 1
    return list(dict.fromkeys(found))
