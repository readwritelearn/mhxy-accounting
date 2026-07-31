# -*- coding: utf-8 -*-
"""
梦幻西游记账软件 - Flask Backend
==============================
数据存储在 Excel 文件中，提供 REST API 供前端调用。

核心概念:
  - 每小时消耗 6 点点卡，点卡成本可配置（默认 0.1 元/点）
  - 代售产品: 有收货成本 + 出售价格，利润 = 差价
  - 纯产出产品: 成本已计入点卡，仅记录收入
  - 灵活收入: 当日金币 - 前一日金币 = 差值收入
"""

import os
from datetime import datetime, date, timedelta
from flask import Flask, render_template, request, jsonify
from openpyxl import Workbook, load_workbook

# ==================== 配置 ====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
DATA_FILE = os.path.join(DATA_DIR, "mhxy_data.xlsx")

DEFAULT_POINT_COST_RATE = 0.1   # 每点点卡成本（元）
DEFAULT_POINTS_PER_HOUR = 6     # 每小时消耗点数

app = Flask(__name__)


# ==================== Excel 数据库 ====================
class ExcelDB:
    """Excel 文件作为数据库，每个 Sheet 对应一张表"""

    def __init__(self, filepath):
        self.filepath = filepath
        self._ensure_file()

    # ---------- 初始化 ----------
    def _ensure_file(self):
        os.makedirs(os.path.dirname(self.filepath), exist_ok=True)
        if not os.path.exists(self.filepath):
            wb = Workbook()
            # 默认 Sheet 改名
            ws = wb.active
            ws.title = "config"
            ws.append(["key", "value"])
            ws.append(["point_cost_rate", DEFAULT_POINT_COST_RATE])
            ws.append(["points_per_hour", DEFAULT_POINTS_PER_HOUR])
            # 其他 Sheet
            for name, headers in [
                ("products",    ["id", "name", "type", "cost_price", "selling_price", "created_at"]),
                ("sales",       ["id", "date", "product_id", "product_name", "type",
                                 "quantity", "unit_price", "revenue", "unit_cost",
                                 "total_cost", "profit"]),
                ("daily_costs", ["id", "date", "online_hours", "point_cost_rate",
                                 "points_per_hour", "points_consumed", "total_cost", "note"]),
                ("gold_balance",["id", "date", "gold_amount", "flexible_income", "note"]),
            ]:
                ws = wb.create_sheet(title=name)
                ws.append(headers)
            wb.save(self.filepath)

    # ---------- 通用方法 ----------
    def _read_all(self, sheet_name):
        """读取整个 Sheet 返回 list[dict]"""
        wb = load_workbook(self.filepath)
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        headers = [cell.value for cell in ws[1]]
        wb.close()
        return [dict(zip(headers, row)) for row in rows if any(v is not None for v in row)]

    def _append_row(self, sheet_name, values):
        """追加一行数据"""
        wb = load_workbook(self.filepath)
        ws = wb[sheet_name]
        ws.append(values)
        wb.save(self.filepath)
        wb.close()

    def _next_id(self, sheet_name):
        """获取下一个自增 ID"""
        wb = load_workbook(self.filepath)
        ws = wb[sheet_name]
        max_id = 0
        for row in ws.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True):
            if row[0] is not None:
                try:
                    max_id = max(max_id, int(row[0]))
                except (ValueError, TypeError):
                    pass
        wb.close()
        return max_id + 1

    # ---------- 配置 ----------
    def get_config(self):
        rows = self._read_all("config")
        return {r["key"]: r["value"] for r in rows}

    # ---------- 产品 ----------
    def get_products(self, product_type=None):
        products = self._read_all("products")
        if product_type:
            products = [p for p in products if p["type"] == product_type]
        # 转换数值类型
        for p in products:
            p["cost_price"] = float(p["cost_price"] or 0)
            p["selling_price"] = float(p["selling_price"] or 0)
            p["id"] = int(p["id"]) if p["id"] else 0
        return products

    def add_product(self, name, ptype, cost_price, selling_price):
        pid = self._next_id("products")
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._append_row("products", [pid, name, ptype, cost_price, selling_price, now])
        return pid

    def delete_product(self, pid):
        wb = load_workbook(self.filepath)
        ws = wb["products"]
        for row in ws.iter_rows(min_row=2):
            if row[0].value == pid:
                ws.delete_rows(row[0].row, 1)
                break
        wb.save(self.filepath)
        wb.close()

    # ---------- 销售 ----------
    def get_sales(self, start_date=None, end_date=None):
        sales = self._read_all("sales")
        result = []
        for s in sales:
            s_date = str(s["date"])[:10] if s["date"] else ""
            if start_date and s_date < start_date:
                continue
            if end_date and s_date > end_date:
                continue
            s["id"] = int(s["id"]) if s["id"] else 0
            s["quantity"] = int(s["quantity"] or 0)
            s["unit_price"] = float(s["unit_price"] or 0)
            s["revenue"] = float(s["revenue"] or 0)
            s["unit_cost"] = float(s["unit_cost"] or 0)
            s["total_cost"] = float(s["total_cost"] or 0)
            s["profit"] = float(s["profit"] or 0)
            result.append(s)
        return result

    def add_sale(self, sale_date, product_id, product_name, ptype,
                 quantity, unit_price):
        sid = self._next_id("sales")
        revenue = quantity * unit_price
        unit_cost = 0
        total_cost = 0

        if ptype == "consignment":
            # 代售产品: 从产品表获取成本
            products = self.get_products("consignment")
            prod = next((p for p in products if p["id"] == product_id), None)
            if prod:
                unit_cost = prod["cost_price"]
                total_cost = unit_cost * quantity

        profit = revenue - total_cost  # 代售利润=差价, 纯产出利润=收入-0=收入
        self._append_row("sales", [
            sid, sale_date, product_id, product_name, ptype,
            quantity, unit_price, revenue, unit_cost, total_cost, profit
        ])
        return sid

    # ---------- 每日成本 ----------
    def get_daily_costs(self, start_date=None, end_date=None):
        costs = self._read_all("daily_costs")
        result = []
        for c in costs:
            c_date = str(c["date"])[:10] if c["date"] else ""
            if start_date and c_date < start_date:
                continue
            if end_date and c_date > end_date:
                continue
            c["id"] = int(c["id"]) if c["id"] else 0
            c["online_hours"] = float(c["online_hours"] or 0)
            c["point_cost_rate"] = float(c["point_cost_rate"] or 0)
            c["points_per_hour"] = float(c["points_per_hour"] or 6)
            c["points_consumed"] = float(c["points_consumed"] or 0)
            c["total_cost"] = float(c["total_cost"] or 0)
            result.append(c)
        return result

    def add_daily_cost(self, cost_date, online_hours, note=""):
        cid = self._next_id("daily_costs")
        config = self.get_config()
        point_rate = float(config.get("point_cost_rate", DEFAULT_POINT_COST_RATE))
        points_per_hour = float(config.get("points_per_hour", DEFAULT_POINTS_PER_HOUR))
        points_consumed = round(online_hours * points_per_hour, 2)
        total_cost = round(points_consumed * point_rate, 2)
        self._append_row("daily_costs", [
            cid, cost_date, online_hours, point_rate,
            points_per_hour, points_consumed, total_cost, note
        ])
        return cid

    # ---------- 金币余额 ----------
    def get_gold_balances(self):
        balances = self._read_all("gold_balance")
        result = []
        for b in balances:
            b["id"] = int(b["id"]) if b["id"] else 0
            b["gold_amount"] = float(b["gold_amount"] or 0)
            b["flexible_income"] = float(b["flexible_income"] or 0)
            result.append(b)
        # 按日期排序
        result.sort(key=lambda x: str(x.get("date", "")))
        return result

    def add_gold_balance(self, bal_date, gold_amount, note=""):
        bid = self._next_id("gold_balance")
        # 自动计算灵活收入 = 今日金币 - 前一日金币
        all_balances = self.get_gold_balances()
        yesterday_amount = 0
        if all_balances:
            yesterday_amount = all_balances[-1]["gold_amount"]
        flexible_income = round(gold_amount - yesterday_amount, 2)
        self._append_row("gold_balance", [
            bid, bal_date, gold_amount, flexible_income, note
        ])
        return bid, flexible_income


# ==================== 全局数据库实例 ====================
db = ExcelDB(DATA_FILE)


# ==================== API 路由 ====================

# ---------- 页面 ----------
@app.route("/")
def index():
    return render_template("index.html")


# ---------- 汇总 ----------
@app.route("/api/summary")
def get_summary():
    """获取当日和当月的收入、利润汇总。
       可选参数 ?month=YYYY-MM，不传则默认当前月份。
       如果当前月份无数据，自动回退到最近有数据的月份。
    """
    req_month = request.args.get("month", "").strip()

    # 所有数据
    all_sales = db.get_sales()
    all_costs = db.get_daily_costs()
    all_balances = db.get_gold_balances()

    # 收集所有有数据的月份
    def month_of(d): return str(d)[:7] if d else ""
    data_months = set()
    for s in all_sales: data_months.add(month_of(s.get("date")))
    for c in all_costs: data_months.add(month_of(c.get("date")))
    for b in all_balances: data_months.add(month_of(b.get("date")))
    data_months.discard("")
    data_months = sorted(data_months, reverse=True)

    current_month = date.today().strftime("%Y-%m")
    today = date.today().isoformat()

    # 确定使用的月份
    if req_month:
        use_month = req_month
    else:
        # 默认：当前月有数据用当前月，否则用最近有数据的月
        if data_months and current_month not in data_months:
            use_month = data_months[0]
        else:
            use_month = current_month

    # 计算该月的日期范围
    month_start = use_month + "-01"
    # month_end: 如果是当前月，截至今天；否则截至该月最后一天
    if use_month == current_month:
        month_end = today
    else:
        y, m = int(use_month[:4]), int(use_month[5:7])
        if m == 12:
            month_end = f"{y+1}-01-01"
        else:
            month_end = f"{y}-{m+1:02d}-01"
        month_end_dt = date(int(month_end[:4]), int(month_end[5:7]), 1) - timedelta(days=1)
        month_end = month_end_dt.isoformat()

    # --- 当日 ---
    today_sales = [s for s in all_sales if str(s.get("date", ""))[:10] == today]
    today_sales_revenue = sum(s["revenue"] for s in today_sales)
    today_sales_profit = sum(s["profit"] for s in today_sales)
    today_costs = [c for c in all_costs if str(c.get("date", ""))[:10] == today]
    today_time_cost = sum(c["total_cost"] for c in today_costs)
    today_flexible = [b for b in all_balances if str(b.get("date", ""))[:10] == today]
    today_flexible_income = sum(b["flexible_income"] for b in today_flexible)
    today_revenue = round(today_sales_revenue + today_flexible_income, 2)
    today_profit = round(today_sales_profit + today_flexible_income - today_time_cost, 2)

    # --- 选定月份 ---
    month_sales = [s for s in all_sales
                   if str(s.get("date", ""))[:10] >= month_start
                   and str(s.get("date", ""))[:10] <= month_end]
    month_sales_revenue = sum(s["revenue"] for s in month_sales)
    month_sales_profit = sum(s["profit"] for s in month_sales)

    month_costs = [c for c in all_costs
                   if str(c.get("date", ""))[:10] >= month_start
                   and str(c.get("date", ""))[:10] <= month_end]
    month_time_cost = sum(c["total_cost"] for c in month_costs)

    month_flexible = [b for b in all_balances
                      if str(b.get("date", ""))[:10] >= month_start
                      and str(b.get("date", ""))[:10] <= month_end]
    month_flexible_income = sum(b["flexible_income"] for b in month_flexible)

    month_revenue = round(month_sales_revenue + month_flexible_income, 2)
    month_profit = round(month_sales_profit + month_flexible_income - month_time_cost, 2)

    # --- 当月每日趋势（始终显示选定月份的每天） ---
    days_in_month = []
    y, m = int(use_month[:4]), int(use_month[5:7])
    d = date(y, m, 1)
    if m == 12:
        next_month = date(y+1, 1, 1)
    else:
        next_month = date(y, m+1, 1)
    end_d = min(date.today(), next_month - timedelta(days=1))
    while d <= end_d:
        d_str = d.isoformat()
        d_sales = [s for s in all_sales if str(s.get("date", ""))[:10] == d_str]
        d_costs = [c for c in all_costs if str(c.get("date", ""))[:10] == d_str]
        d_flex = [b for b in all_balances if str(b.get("date", ""))[:10] == d_str]
        days_in_month.append({
            "date": d_str,
            "revenue": round(sum(s["revenue"] for s in d_sales) +
                             sum(b["flexible_income"] for b in d_flex), 2),
            "profit": round(sum(s["profit"] for s in d_sales) +
                            sum(b["flexible_income"] for b in d_flex) -
                            sum(c["total_cost"] for c in d_costs), 2),
        })
        d += timedelta(days=1)

    return jsonify({
        "today": {
            "date": today,
            "revenue": today_revenue,
            "profit": today_profit,
            "sales_revenue": today_sales_revenue,
            "sales_profit": today_sales_profit,
            "time_cost": today_time_cost,
            "flexible_income": today_flexible_income,
        },
        "month": {
            "month": use_month,
            "revenue": month_revenue,
            "profit": month_profit,
            "sales_revenue": month_sales_revenue,
            "sales_profit": month_sales_profit,
            "time_cost": month_time_cost,
            "flexible_income": month_flexible_income,
        },
        "daily_trend": days_in_month,
        "available_months": data_months,
    })


# ---------- 产品 ----------
@app.route("/api/products")
def api_get_products():
    ptype = request.args.get("type")
    return jsonify(db.get_products(ptype))


@app.route("/api/products", methods=["POST"])
def api_add_product():
    data = request.get_json()
    if not data or not data.get("name") or not data.get("type"):
        return jsonify({"error": "缺少必填字段: name, type"}), 400
    ptype = data["type"]
    if ptype not in ("consignment", "pure_output"):
        return jsonify({"error": "type 必须为 consignment 或 pure_output"}), 400
    cost_price = float(data.get("cost_price", 0))
    selling_price = float(data.get("selling_price", 0))
    pid = db.add_product(data["name"], ptype, cost_price, selling_price)
    return jsonify({"id": pid, "message": "产品添加成功"}), 201


@app.route("/api/products/<int:pid>", methods=["DELETE"])
def api_delete_product(pid):
    db.delete_product(pid)
    return jsonify({"message": "产品已删除"})


# ---------- 销售 ----------
@app.route("/api/sales")
def api_get_sales():
    start = request.args.get("start")
    end = request.args.get("end")
    return jsonify(db.get_sales(start, end))


@app.route("/api/sales", methods=["POST"])
def api_add_sale():
    data = request.get_json()
    required = ["date", "product_id", "product_name", "type", "quantity", "unit_price"]
    for field in required:
        if field not in data:
            return jsonify({"error": f"缺少必填字段: {field}"}), 400

    sid = db.add_sale(
        data["date"],
        int(data["product_id"]),
        data["product_name"],
        data["type"],
        int(data["quantity"]),
        float(data["unit_price"]),
    )
    return jsonify({"id": sid, "message": "销售记录添加成功"}), 201


# ---------- 每日点卡成本 ----------
@app.route("/api/daily-costs")
def api_get_daily_costs():
    start = request.args.get("start")
    end = request.args.get("end")
    return jsonify(db.get_daily_costs(start, end))


@app.route("/api/daily-costs", methods=["POST"])
def api_add_daily_cost():
    data = request.get_json()
    if not data or "date" not in data or "online_hours" not in data:
        return jsonify({"error": "缺少必填字段: date, online_hours"}), 400

    online_hours = float(data["online_hours"])
    if online_hours <= 0:
        return jsonify({"error": "在线时间必须大于0"}), 400

    note = data.get("note", "")
    cid = db.add_daily_cost(data["date"], online_hours, note)
    return jsonify({"id": cid, "message": "成本记录添加成功"}), 201


# ---------- 金币余额 ----------
@app.route("/api/gold-balance")
def api_get_gold_balance():
    return jsonify(db.get_gold_balances())


@app.route("/api/gold-balance", methods=["POST"])
def api_add_gold_balance():
    data = request.get_json()
    if not data or "date" not in data or "gold_amount" not in data:
        return jsonify({"error": "缺少必填字段: date, gold_amount"}), 400

    gold_amount = float(data["gold_amount"])
    note = data.get("note", "")
    bid, flexible_income = db.add_gold_balance(data["date"], gold_amount, note)
    return jsonify({
        "id": bid,
        "flexible_income": flexible_income,
        "message": f"金币记录添加成功，灵活收入: {flexible_income} 元"
    }), 201


# ---------- 配置 ----------
@app.route("/api/config")
def api_get_config():
    return jsonify(db.get_config())


@app.route("/api/config", methods=["PUT"])
def api_update_config():
    data = request.get_json()
    wb = load_workbook(DATA_FILE)
    ws = wb["config"]
    for key, value in data.items():
        found = False
        for row in ws.iter_rows(min_row=2):
            if row[0].value == key:
                row[1].value = value
                found = True
                break
        if not found:
            ws.append([key, value])
    wb.save(DATA_FILE)
    wb.close()
    return jsonify({"message": "配置已更新"})


# ==================== 启动 ====================
if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding='utf-8')
    print("🎮 梦幻西游记账软件 启动中...")
    print(f"📊 数据文件: {DATA_FILE}")
    print(f"🌐 访问地址: http://localhost:5000")
    app.run(debug=True, host="127.0.0.1", port=5000)
