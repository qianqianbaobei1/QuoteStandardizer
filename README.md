# QuoteStandardizer

把厂家的非标准报价单（Excel / PDF 电子版 / PDF 扫描件）解析为统一格式的物料报价表，并逐行和原表对账。

## 原则

- **价格只取原表**：模型返回的价格只用来比对，不会覆盖原表的价格。
- **不知道就留空**：商务条款、单位、编码等，原件中没有或识别不出时一律留空，并写入待复核说明，不填任何默认值。
- **推断要标明**：由文件名推断的供应商、由供应商名推断的品牌、未经编码库核验的编码、识别存疑的单元格，全部写入待复核说明。

## 流程

1. **还原表格**
   - Excel：展开合并单元格；公式取缓存计算值；图片按单元格位置定位。
   - PDF 电子版：用 PyMuPDF 直接提取表格。
   - PDF 扫描件：交给视觉模型逐格转录。第一遍读增强图（去除红色公章、压暗褪色笔画），第二遍读原图，两遍不一致的格子标记待复核。视觉模型不可用时直接报错。
2. **识别表格类型**
   - 清单式：一行一个物料。由模型或启发式规则判断表头行、单价列、单位列。
   - 矩阵式：行标题 × 列标题，交叉格是价格。每个交叉格展开为一条物料，同一列的名称统一；同一列中档位越高价格反而越低时标记待复核。
3. **清洗**：DeepSeek 整理物料名称和结构化参数。原文已经是 `键:值;` 格式的参数原样保留。
4. **商务条款**：从同目录的报价说明/折扣单图片中识别，或用 `--terms-json` 直接提供。公章上识别出的供应商要和文件名、表格内容相互印证，对不上的不采用。
5. **输出与对账**：主交付**只有一张 `标准化报价` Sheet**（不含审计页/对账页，避免污染交付表）。
   - 待复核说明分 **致命 / 严重 / 提示** 三级：致命/严重的那一行在主表**整行红字黄底**，市场价写批注；提示级（推断来源、编码未配置等）只留痕不打扰。
   - 控制台打印**行级与字段级免审率**和「同类问题组」汇总，同组可一次核对/一次修复。
   - 对账与四维体检明细落在**独立文件** `<供应商>_对账与防欺瞒体检报告.xlsx`（`.env` 设 `EXPORT_INTEGRITY_CERTIFICATE=0` 可关闭）。
   - 模板第 1 行里超出内置 33 列的表头（如 `品种`、`厚度(mm)`、`胎基材料`、`是否曝露使用`）按**表头名**从结构化参数取值，不看列位置 —— 列序不同的模板（同一套物料走 112101 与 112115 两种列序）都能对得上。

## 配置（`.env`）

```bash
DEEPSEEK_API_KEY=sk-...
# 可选
DEEPSEEK_BASE_URL=https://api.deepseek.com/v1
DEEPSEEK_MODEL=deepseek-chat          # 需支持图片输入 (扫描件、商务单图片)
VISION_API_KEY= / VISION_BASE_URL= / VISION_MODEL=   # 扫描件改用其它 OpenAI 兼容多模态模型
VISION_DOUBLE_CHECK=1                 # 扫描件两遍识别交叉校验
VISION_GAMMA=1.8                      # 增强图压暗程度
TEMPLATE_PATH=                        # 输出模板 .xlsx (取第 1 行表头)；不填使用内置 33 列表头
CATEGORY_CODE_FILE=                   # 标准分类编码 .csv/.xlsx (前两列：名称、编码)；不填则编码留空
CATEGORY_ATTRS_API_URL=               # 类别属性/填写要求接口；不填则不调用（需配标准编码库）
OUTPUT_DIR=                           # 默认 ./deliveries
# 试点：表格结构 / 行类型判定
STRUCTURE_ENGINE=llm                  # llm (默认) / jev / compare
TYPESAFE_API_KEY=                     # jev、compare 模式需要
JEV_MIN_CONFIDENCE=0.5                # Jev 置信度低于此值时写入待复核说明
LLM_TEMPERATURE=0.1                   # 清洗/审计温度；0 最确定但可能牺牲归纳能力
```

## 试点：用 TypeSafe Jev 判定表格结构

Jev 只做选择题，答案只能是给定选项之一，并附带概率。试点用它判定以下内容：
- **表格结构**（每个 Sheet 发一次请求）：表格类型（清单式 / 矩阵式）、表头行、数据起始行、单价列、单位列；
- **行类型**（每 40 行发一次请求，并发执行）：物料行 / 分组标题 / 表头 / 合计小计 / 备注其他。

三种模式：

| `STRUCTURE_ENGINE` | 输出采用 | 额外内容 |
|---|---|---|
| `llm` | 现有做法 | 无 |
| `jev` | Jev 的判定；置信度低的写入待复核说明 | 「结构判定对比」Sheet |
| `compare` | 现有做法 | 同时调用 Jev，两者判定差异写入「结构判定对比」Sheet，供对照原件判定谁对 |

限制：
- 试点只覆盖清单式表格的结构和行类型；矩阵表只比较表格类型这一项。
- `jev` 模式下，左右并排的子表按整张表处理。
- Jev 返回的答案如果不在选项内、或概率不合法，整次请求作废，退回现有做法。

## 运行

```bash
pip install -r requirements.txt
python main.py -i "报价单.pdf"                       # 自动查找同目录的 报价说明/折扣单 图片
python main.py -i "报价单.xlsx" -t "报价说明.png"
python main.py -i "报价单.pdf" --terms-json terms.json  # 人工确认的商务条款
python3 run_batch_standardize.py "<批次源目录>" "<输出目录>" --label 批次A  # 整批逐家解析；任一家完整性校验不通过则该家 FAIL、退出码 2
python3 zip_standardizer_pipeline.py "报价单.zip"         # 线上导出的压缩包（报价数据.xlsx + quote_img*）→ 单 Sheet 交付
python3 reconcile_vs_source.py "<zip 目录>"               # 交付件 vs 压缩包内原表逐格对账，打印保真度，明细落 deliveries/对账_vs压缩包原表.xlsx
```

`terms.json` 的字段：`supplier`、`brand`、`province`、`city`、`is_tax_included`、`tax_rate`、`tax_desc`、`is_freight_included`、`freight_desc`、`freight_settle`、`discount_ratio`、`discount_series`、`quote_desc`、`quote_type`。没有的字段可以不写。

## 文件

```text
main.py             入口与解析主流程（清单式 / 矩阵式、清洗、对账）
extractor.py        Excel 合并单元格展开、精度保护、图片定位
pdf_extractor.py    PDF → 表格（文字层 / 视觉模型）
terms_extractor.py  商务条款识别与规范化
llm_agent.py        DeepSeek 表头推导与物料清洗
models.py           物料记录与原表对账
builder.py          输出工作簿
config.py           环境配置、省份映射、编码库加载
evaluate.py         评差引擎（与线上系统逐格比对、错误归纳、回归）
run_batch_standardize.py  整批供应商批量解析
zip_standardizer_pipeline.py  线上导出压缩包 → 单 Sheet 交付（多类别 Sheet 收口、列序对齐 33 列标准）
reconcile_vs_source.py       交付件 vs 压缩包内原表逐格对账（保真度 / 补全量 / 致命项）
health_check.py     模块自检（含评差引擎自检）
ERROR_CODES.md      错误码字典 + 每日流程 + 修复账本
```

## 评差与验收口径

线上系统对同一份报价单也出一份结果，本地结果要与它逐格比对评差：

```bash
python evaluate.py scan --ours-dir ours/ --theirs-dir theirs/ -o cases.csv   # 自动配对
python evaluate.py batch --cases cases.csv -o eval_out/     # 评差出报告
python evaluate.py regression                              # 回归，任一用例变差即退出码 1
```

**验收口径：单份文件零差异才算 PASS。** 10 份文件 9 份各错 1 处 = 1/10 PASS（1 份合格），不是 99% 准确率。
错误按 16 个固定错误码归纳（漏行/价格/单位/错字/规格/条款/页码/结构判错…），每个码绑定“该去查哪个函数”。
完整错误码表、每日进化流程和修复账本见 `ERROR_CODES.md`。
