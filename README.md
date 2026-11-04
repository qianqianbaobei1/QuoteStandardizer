# QuoteStandardizer

把厂家的非标准报价单（Excel / PDF 电子版 / PDF 扫描件）解析为统一格式的物料报价表，并逐行和原表对账。

## 原则

- **价格只取原表**：模型返回的价格只用来比对，不会覆盖原表的价格。
- **不知道就留空**：商务条款、单位、编码等，原件中没有或识别不出时一律留空，并写入待复核说明，不填任何默认值。
- **推断要标明**：由文件名推断的供应商、由供应商名推断的品牌、未经编码库核验的编码、识别存疑的单元格，全部写入待复核说明。
- **隐藏内容不解析**：Excel 里被隐藏的工作表、隐藏行、隐藏列一律不看在眼里（厂商常把后台结算价、停用档位藏在隐藏表里，`武汉第二电线电缆` 的「系统结算价」整表隐藏了 58,884 行）。开关是 `SKIP_HIDDEN_CONTENT`，默认开；关掉才回到“看得见、看不见的都解析”的老口径。

## 流程

1. **还原表格**
   - Excel：展开合并单元格；公式取缓存计算值；图片按单元格位置定位。
   - 隐藏口径：`sheet_state != visible` 的工作表整个跳过，`hidden="1"` 的行/列不参与识别（矩阵检测、清单收集、样本抽样、底稿截图都不取），跳过的表写进控制台与 `<输出>.run.json` 的 `skipped_hidden_sheets`。
   - PDF 电子版：用 PyMuPDF 直接提取表格。
   - PDF 扫描件：交给视觉模型逐格转录。第一遍读增强图（去除红色公章、压暗褪色笔画），第二遍读原图，两遍不一致的格子标记待复核。视觉模型不可用时直接报错。
2. **识别表格类型**
   - 清单式：一行一个物料。由模型或启发式规则判断表头行、单价列、单位列。
   - 矩阵式：行标题 × 列标题，交叉格是价格。每个交叉格展开为一条物料，同一列的名称统一；同一列中档位越高价格反而越低时标记待复核。
3. **清洗**：DeepSeek 整理物料名称和结构化参数。原文已经是 `键:值;` 格式的参数原样保留。
4. **商务条款**：从同目录的报价说明/折扣单图片中识别，或用 `--terms-json` 直接提供。公章上识别出的供应商要和文件名、表格内容相互印证，对不上的不采用。
5. **输出与交付规范**：严格遵循**双 Sheet 标准交付架构**：
   - **Sheet 1【客户原始解析】**：1:1 原汁原味保留客户原始解析字段与表头，不掺杂任何模型推断，方便核对溯源；
   - **Sheet 2【标准化报价】**：承载标准 33 列表头，包含标准分类编码、属性值校验回填、物料图示相对超链与高保真内嵌缩略图；
   - **印章穿透与扶正**：集成 `seal_enhancer_service`，自动检测中枢五角星角度偏转并旋转扶正提取盖章主体，利用光学 R 通道多光谱解耦与非锐化凸线滤波（Unsharp Masking），彻底穿透还原被红色印泥遮挡的真实单价、税率与运费；
   - **待复核说明分级**：分 **致命 / 严重 / 提示** 三级，致命/严重项整行红字黄底高亮标注，市场价写入自愈建议批注；
   - **零污染 Sidecar 审计**：运行元数据与对账勾稽明细统一落盘为 `<输出>.run.json` sidecar 文件，保证 Excel 本身清爽整洁。

## 工程目录架构

```text
QuoteStandardizer/
├── main.py                         # 单文件标准化主入口
├── config.py                       # 环境变量与模型配置
├── models.py                       # 33列标准物料模型与严密对账
├── extractor.py                    # Excel 物理展开与合并单元格提取
├── pdf_extractor.py                # PDF 电子版表格提取
├── gemini_pdf_extractor.py         # PDF 视觉模型转录引擎
├── gemini_client.py                # Gemini 全模态客户端
├── llm_agent.py                    # 大模型语义清洗与双遍校验
├── builder.py                      # 双 Sheet 标准工作簿构建器
├── category_attribute_service.py   # 材料分类匹配与特定参数要求服务
├── material_illustration_service.py# 材料配图检索与缩略图渲染服务
├── seal_enhancer_service.py        # 公章智能检测、旋转扶正与遮挡穿透服务
├── 工料机类别.xlsx                   # 标准材料三级分类树字典
│
├── cli/                            # 生产运维与批处理命令行工具集
│   ├── run_all_dual_sheets_rebuild.py # 全量供应商双 Sheet 标准交付构建入口
│   ├── run_batch_standardize.py    # 批量标准化主入口
│   ├── zip_standardizer_pipeline.py# ZIP 压缩包端到端流水线
│   ├── evaluate.py                 # 评差与工业级缺陷评分工具
│   ├── health_check.py             # 全生命周期健康检查与门禁测试
│   ├── audit_all_deliveries.py     # 全量交付件对账与健康巡检
│   ├── fast_audit_deliveries.py    # 极速交付件抽检工具
│   ├── backfill_delivery_images.py # 图片批量回填与相对超链嵌入运维工具
│   └── parse_standalone_screenshot.py # 独立非标单据截图多模态转录工具
│
├── tests/                          # 自动化测试用例集
└── archive/                        # 历史补丁与特定批次归档脚本 (已忽略 Git 追踪)
```

## 系统实际能力清单机读查询

系统支持通过命令行查看当前环境真实接线能力与特性开关，杜绝文档与代码脱节：

```bash
python main.py --capabilities
```

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
python3 cli/run_batch_standardize.py "<批次源目录>" "<输出目录>" --label 批次A  # 整批逐家解析；任一家完整性校验不通过则该家 FAIL、退出码 2
python3 cli/zip_standardizer_pipeline.py "报价单.zip"         # 线上导出的压缩包（报价数据.xlsx + quote_img*）→ 双 Sheet 交付
python3 cli/reconcile_vs_source.py "<zip 目录>"               # 交付件 vs 压缩包内原表逐格对账，打印保真度，明细落 deliveries/对账_vs压缩包原表.xlsx
```

`terms.json` 的字段：`supplier`、`brand`、`province`、`city`、`is_tax_included`、`tax_rate`、`tax_desc`、`is_freight_included`、`freight_desc`、`freight_settle`、`discount_ratio`、`discount_series`、`quote_desc`、`quote_type`。没有的字段可以不写。

## 质量保障与工业级缺陷门禁

系统建立并固化了 16+ 类缺陷错误分类体系（详见 [`ERROR_CODES.md`](ERROR_CODES.md)），支持端到端自动化评差、红线规则拦截与回归体检：

```bash
python3 cli/health_check.py     # 全生命周期模块自检与 CI 准入门禁
python3 cli/evaluate.py selftest # 评差引擎自检（0 差异基线 / 13 类错误全命中 / 等价写法不误报）
```

## 评差与验收口径

线上系统对同一份报价单也出一份结果，本地结果要与它逐格比对评差：

```bash
python3 cli/evaluate.py scan --ours-dir ours/ --theirs-dir theirs/ -o cases.csv   # 自动配对
python3 cli/evaluate.py batch --cases cases.csv -o eval_out/     # 评差出报告
python3 cli/evaluate.py regression                              # 回归，任一用例变差即退出码 1
```

**验收口径：单份文件零差异才算 PASS。** 10 份文件 9 份各错 1 处 = 1/10 PASS（1 份合格），不是 99% 准确率。
错误按 16 个固定错误码归纳（漏行/价格/单位/错字/规格/条款/页码/结构判错…），每个码绑定“该去查哪个函数”。
完整错误码表、每日进化流程和修复账本见 `ERROR_CODES.md`。
