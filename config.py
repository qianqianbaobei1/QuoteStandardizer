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
# 扫描件是否让视觉模型独立读两遍做交叉校验 (默认关闭以提高效率；可通过环境变量 VISION_DOUBLE_CHECK=1 开启)
VISION_DOUBLE_CHECK = os.getenv("VISION_DOUBLE_CHECK", "0").strip().lower() in ("1", "true")
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
# Jev 行类型每次请求带多少行
JEV_ROW_BATCH = int(os.getenv("JEV_ROW_BATCH", "40"))
# 注入清洗提示词的标准编码库条数上限 (库很大时逐批全量注入会显著推高 token 成本并引起延迟)
# 每批注入多少条“相关资料库”条目 (按本批物料文字筛，不是库的前 N 条)
CODE_LIBRARY_MAX_LINES = int(os.getenv("CODE_LIBRARY_MAX_LINES", "300"))

# 线路引擎选择: 'gemini' (Gemini原生全模态端到端直传，首选) / 'traditional' (原切图流水线) / 'compare' (双轨对比)
EXTRACTOR_ENGINE = os.getenv("EXTRACTOR_ENGINE", "traditional").strip().lower()

# Gemini 原生引擎配置
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.8-flash")
GEMINI_FALLBACK_MODEL = os.getenv("GEMINI_FALLBACK_MODEL", "gemini-3.6-flash")
# Gemini 思考预算 (0 为禁用思考/最低开销，-1 为自动，默认设为 0 以极速返回)
GEMINI_THINKING_BUDGET = int(os.getenv("GEMINI_THINKING_BUDGET", "0"))

# 防欺瞒工程参数配置
# 隐藏内容口径：Excel 里被隐藏的工作表 / 行 / 列一律不解析（厂商常把后台结算价、停用档位藏在隐藏表里）
SKIP_HIDDEN_CONTENT = os.getenv("SKIP_HIDDEN_CONTENT", "1").strip().lower() in ("1", "true")

ANTI_HALLUCINATION_MATH_AUDIT = os.getenv("ANTI_HALLUCINATION_MATH_AUDIT", "1").strip().lower() in ("1", "true")
ANTI_HALLUCINATION_STRICT_ROW_CONSERVATION = os.getenv("ANTI_HALLUCINATION_STRICT_ROW_CONSERVATION", "1").strip().lower() in ("1", "true")
EXPORT_INTEGRITY_CERTIFICATE = os.getenv("EXPORT_INTEGRITY_CERTIFICATE", "0").strip().lower() in ("1", "true")
# 严格引擎模式：1 时禁止任何静默回退，若请求引擎不可用或失败则直接硬失败退出；0 为兼容平滑模式 (但仍写入 degraded=true)
STRICT_ENGINE = os.getenv("STRICT_ENGINE", "0").strip().lower() in ("1", "true")

# LLM 调用重试次数
LLM_MAX_RETRIES = int(os.getenv("LLM_MAX_RETRIES", "2"))
# 清洗/审计温度：0 最确定但可能牺牲归纳能力；默认 0.1
LLM_TEMPERATURE = float(os.getenv("LLM_TEMPERATURE", "0.1"))

DEFAULT_OUTPUT_DIR = os.getenv("OUTPUT_DIR") or os.path.join(CURRENT_DIR, "deliveries")
# 输出模板 (取第 1 行作为表头)；未配置时使用 builder.STANDARD_HEADERS
DEFAULT_TEMPLATE_PATH = os.getenv("TEMPLATE_PATH", "")
# 标准分类编码文件 (.csv / .xlsx，前两列为 名称、编码)；未配置时不做编码库匹配
CATEGORY_CODE_FILE = os.getenv("CATEGORY_CODE_FILE", "")
# 类别属性/填写要求接口；留空则只按标准编码库的本地规则，不调外部接口
CATEGORY_ATTRS_API_URL = os.getenv("CATEGORY_ATTRS_API_URL", "")

# 材料配图接口配置 (来自 工料机类别.xlsx Sheet3)
ILLUSTRATION_API_BASE_URL = os.getenv("ILLUSTRATION_API_BASE_URL", "https://ugw.m2.com.cn/agents-toolkit")
ILLUSTRATION_TASK_CREATE_URL = os.getenv("ILLUSTRATION_TASK_CREATE_URL") or f"{ILLUSTRATION_API_BASE_URL.rstrip('/')}/api/material/illustrations"
ILLUSTRATION_TASK_QUERY_URL = os.getenv("ILLUSTRATION_TASK_QUERY_URL") or f"{ILLUSTRATION_API_BASE_URL.rstrip('/')}/api/material/illustration-tasks"
ENABLE_MATERIAL_ILLUSTRATION = os.getenv("ENABLE_MATERIAL_ILLUSTRATION", "0").strip().lower() in ("1", "true")
DOWNLOAD_ILLUSTRATION_IMAGES = os.getenv("DOWNLOAD_ILLUSTRATION_IMAGES", "0").strip().lower() in ("1", "true")

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

# 全国/无地区报价是否按省市铺开成 31 行（人工标杆口径）；关闭则原样输出
EXPAND_NATIONWIDE = os.getenv("EXPAND_NATIONWIDE", "1").strip().lower() in ("1", "true")
# 全国报价按省市铺开时使用的标准省市对（人工标杆口径）
NATIONWIDE_REGIONS = [
    ('上海市', '上海市'),
    ('北京市', '北京市'),
    ('天津市', '天津市'),
    ('河北省', '石家庄市'),
    ('山西省', '太原市'),
    ('内蒙古自治区', '呼和浩特市'),
    ('辽宁省', '沈阳市'),
    ('吉林省', '长春市'),
    ('黑龙江省', '哈尔滨市'),
    ('江苏省', '南京市'),
    ('浙江省', '杭州市'),
    ('安徽省', '合肥市'),
    ('福建省', '福州市'),
    ('江西省', '南昌市'),
    ('山东省', '济南市'),
    ('河南省', '郑州市'),
    ('湖北省', '武汉市'),
    ('湖南省', '长沙市'),
    ('广东省', '广州市'),
    ('广西壮族自治区', '南宁市'),
    ('海南省', '海口市'),
    ('重庆市', '重庆市'),
    ('四川省', '成都市'),
    ('贵州省', '贵阳市'),
    ('云南省', '昆明市'),
    ('西藏自治区', '拉萨市'),
    ('陕西省', '西安市'),
    ('甘肃省', '兰州市'),
    ('青海省', '西宁市'),
    ('宁夏回族自治区', '银川市'),
    ('新疆维吾尔自治区', '乌鲁木齐市')
]


def _read_category_rows(path: str) -> list:
    """把标准编码库文件读成原始行 (xlsx/csv 两路都认)"""
    if not path or not os.path.exists(path):
        return []
    if path.lower().endswith(('.xlsx', '.xlsm')):
        import openpyxl
        ws = openpyxl.load_workbook(path, read_only=True, data_only=True).active
        return [list(r) for r in ws.iter_rows(values_only=True)]
    with open(path, encoding='utf-8-sig', newline='') as f:
        return [list(r) for r in csv.reader(f)]


def _load_category_tree(path: str) -> list:
    """(一级名称, 二级名称, 三级编码, 三级名称)。按二级类别扩展候选时要用整棵树"""
    rows = _read_category_rows(path)
    if not rows:
        return []
    headers = [str(c or '').strip() for c in rows[0]]
    idx = {k: next((i for i, h in enumerate(headers) if k in h), -1)
           for k in ('一级类别名称', '二级类别名称', '三级类别编码', '三级类别名称')}
    if any(i == -1 for i in idx.values()):
        return []
    tree = []
    for r in rows[1:]:
        if len(r) <= max(idx.values()):
            continue
        l3c = str(r[idx['三级类别编码']] or '').strip()
        l3n = str(r[idx['三级类别名称']] or '').strip()
        if not (l3c and l3n):
            continue
        tree.append((str(r[idx['一级类别名称']] or '').strip(),
                     str(r[idx['二级类别名称']] or '').strip(), l3c, l3n))
    return tree


def _load_category_codes(path: str) -> dict:
    """读取标准分类编码文件：支持前两列为(名称,编码)或(编码,名称)，以及标准6列库(三级类别编码/名称)"""
    if not path:
        return {}
    if not os.path.exists(path):
        raise FileNotFoundError(f"CATEGORY_CODE_FILE 不存在: {path}")
    rows = _read_category_rows(path)
    if not rows:
        return {}
    headers = [str(c or '').strip() for c in rows[0]]
    codes = {}
    # 场景1：带有标准表头
    if any('三级' in h for h in headers):
        c_idx = next((i for i, h in enumerate(headers) if '三级类别编码' in h), -1)
        n_idx = next((i for i, h in enumerate(headers) if '三级类别名称' in h), -1)
        if c_idx != -1 and n_idx != -1:
            for r in rows[1:]:
                if len(r) > max(c_idx, n_idx):
                    code = str(r[c_idx] or '').strip()
                    name = str(r[n_idx] or '').strip()
                    if code and name:
                        codes[name] = code
            return codes
    # 场景2：常规两列或首两列 (名称, 编码) 或 (编码, 名称)
    for r in rows:
        if len(r) < 2 or r[0] in (None, '') or r[1] in (None, ''):
            continue
        v0, v1 = str(r[0]).strip(), str(r[1]).strip()
        if re.fullmatch(r'[A-Za-z0-9]+', v1) and not re.fullmatch(r'[A-Za-z0-9]+', v0):
            codes[v0] = v1
        elif re.fullmatch(r'[A-Za-z0-9]+', v0) and not re.fullmatch(r'[A-Za-z0-9]+', v1):
            codes[v1] = v0
    return codes


CATEGORY_CODE_MAP = _load_category_codes(CATEGORY_CODE_FILE)

# 整棵三级类别树：名称 → 编码命中不了时，按二级/一级类别扩展候选（灯管 → 灯具、光源 下的节点）
CATEGORY_TREE = _load_category_tree(CATEGORY_CODE_FILE)


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
