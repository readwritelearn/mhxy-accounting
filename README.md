# 梦幻西游记账软件

梦幻西游点卡成本计账 & 代售/纯产出利润管理系统。

## 功能

- **在线时间 → 自动换算点卡成本**（默认 6 点/小时 × ¥0.1/点）
- **代售产品**：收货成本 & 售价 → 差价 = 利润
- **纯产出产品**（打造/活力服务）：仅记收入，成本已计入点卡
- **金币灵活收入**：当日金币 − 前日金币 = 差值收入
- **月度总览**：本月收入/利润趋势图表
- **Excel 数据库**：所有数据存于 `data/mhxy_data.xlsx`

## 快速启动

```bash
# 安装依赖
pip install flask openpyxl

# 启动
python app.py
```

浏览器打开 **http://localhost:5000**

## 项目结构

```
mhxy-accounting/
├── app.py              # Flask 后端 + Excel 数据层
├── requirements.txt    # Python 依赖
├── templates/
│   └── index.html      # 前端 SPA（6 个功能 Tab）
└── data/
    └── mhxy_data.xlsx  # Excel 数据库（自动创建）
```

## 点卡成本公式

```
点卡成本 = 在线小时 × 6 点/时 × 点卡单价（默认 ¥0.1/点）
         = 在线小时 × ¥0.6
```

可在 `data/mhxy_data.xlsx` 的 `config` sheet 中修改 `point_cost_rate`。
