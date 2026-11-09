# 工料机标准分类与属性拆分标准化独立部署包 (QuoteStandardizer Bundle v1.0)

本部署包提炼自 QuoteStandardizer 核心引擎，解耦了对主项目繁杂提取逻辑的依赖，完整包含了：
1. **核心计算引擎** (`core/`)：JEV 多维打分引擎、广材双大库反查、物理脱壳与名称归一化；
2. **底表元数据与知识库** (`data/`)：国标三级树、3,657万广材大库属性Schema库、39.8万品名企业分布词表；
3. **极简门面 SDK** (`standardizer_engine.py`)：3 行代码快速集成；
4. **线上微服务** (`api_server.py`)：FastAPI 高性能 HTTP 接口；
5. **规则白皮书** (`rules/RULEBOOK.md`)：全套打分机制与物理脱壳规则。

---

## 一、 快速开始

### 1. 安装依赖
```bash
pip install -r requirements.txt
```

### 2. Python SDK 方式直接调用
```python
from standardizer_engine import StandardizerEngine

engine = StandardizerEngine(use_jev=True)

# 单条标准化
res = engine.standardize_item(name="热轧圆钢", spec="Φ18 Q235B", raw_cat="金属材料")
print(res["category_code"])        # 01-01-01
print(res["level3"])               # 热轧圆钢
print(res["extracted_attributes"]) # {'品种': '热轧圆钢', '直径(mm)': '18', '牌号': 'Q235B'}
```

### 3. 启动 HTTP 微服务
```bash
python api_server.py
# 访问 Swagger UI: http://localhost:8080/docs
```

---

## 二、 环境变量配置

| 变量名 | 说明 | 默认值 |
| :--- | :--- | :--- |
| `STANDARDIZER_DATA_DIR` | 底表数据目录绝对路径 | 当前部署包下的 `data/` |
| `TYPESAFE_API_KEY` | JEV 模型鉴权 Key | 预置生产 Key |
| `TYPESAFE_BASE_URL` | JEV 线上推理服务地址 | 预置网关地址 |
| `PORT` | 微服务启动端口 | 8080 |
