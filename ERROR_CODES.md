# 评差与进化：错误码字典 + 每日流程

线上系统（下文简称「他们」）与本地生产线（「我们」）对同一份报价单出结果，评差引擎把两边逐格比对。
**验收口径只有一条：单份文件零差异才算 PASS。** 10 份文件 9 份各错 1 处 = 1/10 PASS，不是 99% 准确率。

引擎会同时打印两个数字，用来让甲方看清区别：

```
文件级：1/10 PASS（10.0%）   ← 验收口径：任一处对不上即 FAIL
字段级：99.97%（仅参考，不作为验收；会把个别文件的硬伤平均掉）
```

## 一、每日流程（5 条命令）

```bash
# 0. 新文件进来：ours/ 与 theirs/ 按文件名主干自动配对（只有一边的会告警，不许静默跳过）
python evaluate.py scan --ours-dir ours/ --theirs-dir theirs/ -o cases.csv

# 1. 评差
python evaluate.py batch --cases cases.csv -o eval_out/

# 2. 人工裁决：把 eval_out/评差报告.xlsx 的「差异明细」过一遍，对照原件填 truth
#    verdicts.csv: file_id,field,ours,theirs,truth
#    truth = ours(我们对) / theirs(他们对) / both_wrong / equiv

# 3. 带裁决重跑，看"谁错在哪一类"
python evaluate.py batch --cases cases.csv -o eval_out/ --verdicts verdicts.csv

# 4. 归档成回归用例 + 回归（只增不减）
python evaluate.py archive --cases cases.csv --dest fixtures/
python evaluate.py regression            # 任一用例比基线差 → 退出码 1
```

日常改完代码先跑：

```bash
python -m unittest health_check      # 模块级自检（含评差引擎自检）
python evaluate.py selftest          # 评差引擎：0 差异基线 / 13 类错误全命中 / 等价写法不误报
```

约定目录：`ours/` 放我们每份文件的输出（`公司名-标准解析.xlsx`），`theirs/` 放线上系统对同一批文件输出的 xlsx，文件名主干一致才能配对。

## 二、错误码字典

`evaluate.py` 里的 `CODE_META` 是唯一权威，**加字段/加错误码必须两边同步改**。

| 错误码 | 含义 | 级别 | 下一步查哪里 |
|---|---|---|---|
| `E00_EQUIV` | 写法差异（已归一，不判错） | INFO | 无：数值/全半角/单位同义词等价，明细里可见，防止引擎偷偷放水 |
| `E01_ROW_OURS_MISSING` | 我们少一条 | FATAL | `main._collect_list_rows` 行过滤 / PDF 切块去重 / 页丢失 |
| `E02_ROW_OURS_EXTRA` | 我们多一条 | FATAL | `main._detect_matrix` 矩阵误判 / 表头行判错 / 合计行被当物料 |
| `E03_PRICE` | 价格不一致 | FATAL | `extractor.safe_clean_price` / 单价列判定 `price_col` / 矩阵交叉格 |
| `E04_UNIT` | 单位不一致 | FATAL | `main.normalize_unit` / `unit_col` 判定 |
| `E05_NAME_TYPO` | 名称错字（编辑距离小） | FATAL | `llm_agent` 清洗 / 视觉模型转录 / 合并单元格展开 |
| `E06_NAME_DIFF` | 名称语义不同 | FATAL | `main._unify_group_names` 多数票统一 / 表头路径拼接 |
| `E07_SPEC_MISSING` | 规格参数缺项/多项 | FATAL | `main._verbatim_specs` / `_genoffice_raw_spec` / 单元格混写拆分 |
| `E08_SPEC_VALUE` | 规格参数值不一致 | FATAL | `llm_agent` 清洗（参数值必须照抄原文）/ 视觉转录 |
| `E09_CODE` | 类别编码不一致 | ERROR | `config.resolve_category_code` / `CATEGORY_CODE_FILE` 未配置 |
| `E10_BRAND` | 品牌不一致 | ERROR | `main.derive_supplier_and_brand_dynamic` / 表内品牌列 |
| `E11_TERMS` | 商务条款不一致（折扣/税/运费/结算） | ERROR | `terms_extractor` / 商务单图片识别 / 折扣解析 |
| `E12_REGION` | 报价地区不一致 | ERROR | `config.expand_region` / 行内地区列 |
| `E15_PAGE` | 原表页码不一致 | ERROR | `pdf_extractor` 页序 / Gemini `page_number` / 切块合并 |
| `E14_STRUCTURE` | **疑似整表结构判错（已折叠逐行差异）** | FATAL | `main._detect_matrix` / `_collect_list_rows` 表头行 |
| `E16_HEADER` | 输出表结构/列名对不上 | FATAL | 对方列名与标准 33 列的映射（`ALIASES`） |

`E14_STRUCTURE` 是**归纳的核心**：一份文件里 ≥20% 的行对不上时，逐行差异是症状、结构判错才是病根。
引擎会把它们折叠成一条 E14 并优先作为「主错误码」，明细仍完整保留在报告里。

## 三、进化闭环

每天的新文件进来后，只有一条路径算「进化」，其他都是打补丁：

```
新文件 FAIL
  → 错误码归到某一类（上面 16 个码之一，不允许出现"其他"）
  → 按「下一步查哪里」定位到具体函数
  → 改一条**通用规则**（不是 if 文件名/if 这家供应商）
  → 该文件归档进 fixtures/，baseline 只允许变好
  → 记一行账本（见下）
```

三条硬约束：

1. **禁止按文件名/供应商名写分支。** 一家的特例就是下一批文件的坑。只能改归一规则、结构判据、提示词。
2. **回归用例只增不减。** 修过的坑必须变成 `fixtures/` 里的一个用例，否则下周同一个坑会回来。
3. **`evaluate.py selftest` 必须一直全绿。** 它保证「引擎自己不出错」：0 差异基线、13 类错误各命中、等价写法不误报、结构错位可识别。

## 四、引擎不比对什么（避免误以为漏检）

| 字段 | 原因 |
|---|---|
| `报价ID` / `顺序号` | 两边编号规则不同，不是业务数据 |
| `厂家原始规格参数` | 我们的拼接口径自带本系统格式 |
| `报价说明` | 内含我们自己的「⚠️待人工复核」标记 |
| `图片文件夹` / 5 个 `图片` 列 | 附件列，比文件名无意义（图片缺失另记为 E15 页码错配） |

## 五、已知盲区（必须靠人补）

- **两边同错 = PASS。** 评差只能发现「不一致」，发现不了「一致但都错」。所以每天必须对原件**抽检**：`fixtures/` 里随机抽 2 份，对着 PDF 原件人工核 3 行。裁决里出现 `both_wrong` 就是在提醒这件事。
- **引擎不判定真相。** 差异只说明两套系统不一致，`truth` 必须人填。引擎替人裁决就是制造第二个幻觉源。
- **PDF 视觉模型两遍交叉校验默认关闭**（`VISION_DOUBLE_CHECK=1` 开启，代价是调用量翻倍）。扫描件批次建议开启。

## 六、账本

每修一个坑记一行：`日期 | 错误码 | 根因 | 改了什么 | 回归用例`

| 日期 | 错误码 | 根因 | 改了什么 | 回归用例 |
|---|---|---|---|---|
| 2026-10-30 | `E02_ROW_OURS_EXTRA` | `_detect_matrix` 把「名称\|长度\|宽度」清单表判成矩阵，一个数值属性格造一条假物料，对账还全绿 | 行标题末级必须是规格/档位维度（`MATRIX_LABEL_RE`），否则退回清单式并大声报错 | `health_check.test_09` |
| 2026-10-30 | `E01_ROW_OURS_MISSING` | PDF 文字层一页多表时 `max(area)` 只取最大一张，其余整张静默丢弃 | `_grids_from_text_layer` 返回全部表格，仅丢弃面积<20% 的碎片并打印 | 待补（需多表 PDF 原件） |
| 2026-10-30 | `E01_ROW_OURS_MISSING` | Gemini 只返回部分页时静默落盘，Sheet 名按序号编号导致页码全错、截图张冠李戴 | 页数不符显式上报「整篇」失败；Sheet 名改用模型自报 `page_number` | 待补（需多页 PDF 原件） |
| 2026-10-30 | `E01/E02` | 被判为「跳过」的行永不复核，规则误伤即静默丢失 | `reconcile` 对「有单价却被跳过」判异常；结构性跳过（标题/表头/合计）豁免 | `health_check.test_07` |
| 2026-10-30 | 全体 | 对账异常只打印、退出码仍 0，交付物「看起来成功」 | `[VERIFY]` 行 + `exit 2`；写盘后读回逐行校验（行数/报价ID/市场价） | `health_check.test_08` |
