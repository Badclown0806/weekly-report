#!/usr/bin/env python3
"""
build_data.py - 从源Excel文件生成 data.json
源文件:
  - D:/周汇报文件/运营日数据.xlsx
  - D:/周汇报文件/产品列表.xlsx
  - D:/周汇报文件/LX利润表.xlsx
  - D:/周汇报文件/2026WB年规进度.xlsx
输出: output/data.json
"""

import json
import math
import os
import re
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta

import openpyxl

# ── 配置 ──────────────────────────────────────────────
SRC_DIR = r"D:\周汇报文件"
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_PATH = os.path.join(OUTPUT_DIR, "data.js")

# ── 工具函数 ──────────────────────────────────────────

def sanitize_value(v):
    """确保值是合法JSON：NaN/Inf/-Inf → null，字符串数字 → float"""
    if v is None:
        return None
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            return None
        # Python float → JSON number, round to 4 decimals for rates
        return round(v, 6) if abs(v) < 10 else round(v, 2)
    if isinstance(v, bool):
        return v
    if isinstance(v, int):
        return v
    if isinstance(v, str):
        # Excel 可能将数值读为字符串，尝试转换
        v_stripped = v.strip()
        if not v_stripped:
            return None
        try:
            f = float(v_stripped)
            if math.isnan(f) or math.isinf(f):
                return None
            return round(f, 6) if abs(f) < 10 else round(f, 2)
        except (ValueError, TypeError):
            return v
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    return str(v)


def iso_week_to_date_range(iso_week_str):
    """'2025-W30' → (monday, sunday)"""
    year_s, week_s = iso_week_str.split("-W")
    year, week = int(year_s), int(week_s)
    jan4 = date(year, 1, 4)
    monday = jan4 - timedelta(days=jan4.weekday()) + timedelta(weeks=week - 1)
    return monday, monday + timedelta(days=6)


def date_to_iso_week(d):
    """date → '2025-W30'"""
    if isinstance(d, datetime):
        d = d.date()
    iso = d.isocalendar()
    return f"{iso[0]}-W{iso[1]:02d}"


def load_workbook_safe(path):
    """安全加载workbook"""
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        return wb
    except Exception as e:
        print(f"ERROR: 无法加载 {path}: {e}")
        return None


# ── 阶段 1: 生成WEEKS等基础数据 ───────────────────────

def generate_weeks():
    """生成52周的数组"""
    weeks = []
    weeks_iso = []
    week_labels = []
    year, wn = 2025, 30
    for i in range(52):
        iso = f"{year}-W{wn:02d}"
        start, end = iso_week_to_date_range(iso)
        label = f"{iso} (W{i+1}·{start.month:02d}.{start.day:02d}-{end.month:02d}.{end.day:02d})"
        weeks.append(f"W{i+1}")
        weeks_iso.append(iso)
        week_labels.append(label)
        wn += 1
        if wn > 52:
            wn = 1
            year += 1
    return weeks, weeks_iso, week_labels


# ── 阶段 2: 读取产品列表 ──────────────────────────────

def read_product_list():
    """从产品列表.xlsx 提取 SKU映射"""
    path = os.path.join(SRC_DIR, "产品列表.xlsx")
    wb = load_workbook_safe(path)
    if wb is None:
        return {}, {}, {}, {}, {}

    ws = wb[wb.sheetnames[0]]  # '产品列表'
    sku_img = {}
    sku_owner = {}
    sku_first_date = {}
    sku_wb_id = {}
    shop_owner_set = defaultdict(set)

    # 列: 0=WB商品ID, 1=卖家SKU, ..., 9=主图, 10=店铺名称, 13=负责人, 14=创建时间
    for i, row in enumerate(ws.iter_rows(min_row=2)):
        if i > 2000:
            break  # safety limit
        vals = [cell.value for cell in row[:15]]
        wb_id = vals[0]
        sku = vals[1]
        img = vals[9]
        shop = vals[10] if len(vals) > 10 else None
        owner = vals[13] if len(vals) > 13 else None
        create_time = vals[14] if len(vals) > 14 else None

        if not sku or not wb_id:
            continue

        wb_id_str = str(int(wb_id)) if isinstance(wb_id, float) else str(wb_id)
        
        # 关键修改：使用 SKU+WB商品ID 组合作为唯一标识
        sku_wb_key = f"{sku}|{wb_id_str}"
        if shop:
            sku_wb_key_shop = f"{sku}|{shop}|{wb_id_str}"
        
        if img:
            sku_img[sku_wb_key] = str(img)
            if shop:
                sku_img[sku_wb_key_shop] = str(img)
        if owner:
            sku_owner[sku_wb_key] = str(owner)
            if shop:
                shop_owner_set[str(shop)].add(str(owner))
                sku_owner[sku_wb_key_shop] = str(owner)
        if create_time:
            fd_str = None
            if isinstance(create_time, (datetime, date)):
                fd_str = create_time.strftime("%Y-%m-%d") if isinstance(create_time, datetime) else create_time.isoformat()
            elif isinstance(create_time, str):
                fd_str = create_time[:10]
            if fd_str:
                sku_first_date[sku_wb_key] = fd_str
                if shop:
                    sku_first_date[sku_wb_key_shop] = fd_str
        
        # 存储 WB商品ID 映射
        sku_wb_id[sku_wb_key] = wb_id_str
        if shop:
            sku_wb_id[sku_wb_key_shop] = wb_id_str

    # 转换 shop_owner_set → dict
    shop_owners = {s: {o: True for o in owners} for s, owners in shop_owner_set.items()}

    wb.close()
    print(f"  产品列表: {len(sku_img)} SKU图片, {len(sku_owner)} SKU负责人, "
          f"{len(shop_owners)} 店铺负责人, {len(sku_wb_id)} SKU-WB映射, {len(sku_first_date)} 首次日期")
    return sku_img, sku_owner, sku_first_date, sku_wb_id, shop_owners


# ── 阶段 3: 读取LX利润表 ──────────────────────────────

def read_lx_profit(weeks_iso):
    """从LX利润表.xlsx 生成 WEEK_DATA 和 SHOP_WEEKLY"""
    path = os.path.join(SRC_DIR, "LX利润表.xlsx")
    wb = load_workbook_safe(path)
    if wb is None:
        return {}, {}

    iso_to_w = {iso: f"W{i+1}" for i, iso in enumerate(weeks_iso)}

    # ─── 店铺分周利润表 → SHOP_WEEKLY ───
    ws_shop = wb["店铺分周利润表"]
    shop_weekly = defaultdict(dict)

    for i, row in enumerate(ws_shop.iter_rows(min_row=3)):
        if i > 5000:
            break
        vals = [cell.value for cell in row[:10]]
        week_end = vals[1]  # 星期结束值
        shop_name = vals[2]
        margin_val = vals[4]  # 毛利率CNY
        gsv_val = vals[6]    # 后台价GSV.CNY

        if not week_end or not shop_name:
            continue

        iso = date_to_iso_week(week_end) if isinstance(week_end, (datetime, date)) else str(week_end)
        w_key = iso_to_w.get(iso)
        if w_key is None:
            continue

        margin = sanitize_value(margin_val)
        gsv = sanitize_value(gsv_val)

        # 只保留 margin >= 0 的有效周数据
        if margin is not None and gsv is not None:
            # margin 在SHOP_WEEKLY中是百分比形式 (如 6.05 表示 6.05%)
            shop_weekly[str(shop_name)][w_key] = {
                "gsv": gsv,
                "margin": round(margin * 100, 4) if isinstance(margin, float) and margin < 1 else margin
            }

    # 转换为普通dict
    shop_weekly = {k: dict(v) for k, v in shop_weekly.items()}

    # ─── 分周SKU → WEEK_DATA ───
    ws_sku = wb["分周SKU"]
    # 列: 0=数据范围, 1=星期结束值, 2=店铺名称, 3=负责人, 4=类目,
    #     5=WB商品ID, 6=卖家SKU, 7=主图, 8=毛利量CNY, 9=毛利率CNY,
    #     10=GSV(后台价), 11=周订单量售完天数, 12=每周日库存量, 13=货值CNY,
    #     14=销售数量, 15=退款数量, 16=财报净销量, ..., 21=送达退货率,
    #     36=AK列(广告花费)

    week_data_raw = defaultdict(list)

    sku_count = 0
    for i, row in enumerate(ws_sku.iter_rows(min_row=3)):
        if i > 40000:
            break
        vals = [cell.value for cell in row[:37]]
        week_end = vals[1]
        shop = vals[2]
        cat = vals[4]
        sku = vals[6]
        profit = vals[8]     # 毛利量CNY
        margin_rate = vals[9]  # 毛利率CNY (as decimal e.g., 0.2507)
        gsv = vals[10]       # GSV(后台价)
        qty = vals[14]       # 销售数量
        return_rate = vals[21] if len(vals) > 21 else None  # 送达退货率
        ad_spend = vals[36] if len(vals) > 36 else None  # AK列 广告花费

        if not week_end or not sku:
            continue

        iso = date_to_iso_week(week_end) if isinstance(week_end, (datetime, date)) else str(week_end)
        w_key = iso_to_w.get(iso)
        if w_key is None:
            continue

        sku_str = str(sku).strip() if sku else ""

        product = {
            "sku": sku_str or "",
            "shop": str(shop) if shop else "",
            "cat": str(cat) if cat else "",
            "profit": sanitize_value(profit) or 0,
            "margin": sanitize_value(margin_rate),
            "gsv": sanitize_value(gsv) or 0,
            "qty": sanitize_value(qty) or 0,
            "return_rate": sanitize_value(return_rate),
            "ad_spend": sanitize_value(ad_spend) or 0
        }

        # margin是小数, 转为百分比
        if isinstance(product["margin"], float) and product["margin"] < 1:
            product["margin"] = round(product["margin"] * 100, 2)
        # return_rate 也是小数, 同样转为百分比
        if isinstance(product["return_rate"], float) and product["return_rate"] < 1:
            product["return_rate"] = round(product["return_rate"] * 100, 2)

        week_data_raw[w_key].append(product)
        sku_count += 1

    # 构建最终的 WEEK_DATA
    week_data = {}
    for w_key in sorted(week_data_raw.keys(), key=lambda x: int(x[1:])):
        products = week_data_raw[w_key]

        # 按 profit 降序排列
        products.sort(key=lambda p: p["profit"], reverse=True)

        # 计算店铺汇总
        shop_summary = defaultdict(lambda: {"gsv": 0, "profit": 0, "margin": 0, "products": 0, "ad_spend": 0})
        for p in products:
            s = shop_summary[p["shop"]]
            s["gsv"] += p["gsv"]
            s["profit"] += p["profit"]
            s["products"] += 1
            s["ad_spend"] += p.get("ad_spend", 0) or 0

        # 计算 weighted margin
        for s in shop_summary.values():
            if s["gsv"] > 0:
                s["margin"] = round(s["profit"] / s["gsv"] * 100, 2)
            else:
                s["margin"] = 0

        # top10 by profit
        top10 = [p for p in products if p["sku"] != "无匹配ID费用"][:10]

        week_data[w_key] = {
            "shops": {k: dict(v) for k, v in shop_summary.items()},
            "top10Profit": top10,
            "allProducts": products
        }

    wb.close()
    print(f"  SHOP_WEEKLY: {len(shop_weekly)} shops")
    print(f"  WEEK_DATA: {len(week_data)} weeks, {sku_count} total product-weeks")
    return shop_weekly, week_data


# ── 阶段 4: 读取运营日数据 → TRAFFIC_WEEKLY ──────────

def read_traffic_weekly(weeks_iso):
    """从运营日数据.xlsx 生成 TRAFFIC_WEEKLY"""
    path = os.path.join(SRC_DIR, "运营日数据.xlsx")
    wb = load_workbook_safe(path)
    if wb is None:
        return {}

    iso_to_w = {iso: f"W{i+1}" for i, iso in enumerate(weeks_iso)}
    ws = wb[wb.sheetnames[0]]  # Export

    # 列: 0=日期, 2=店铺名称, 3=卖家SKU, 7=访客, 10=加购数, 11=加购转化, 12=销量, 19=转化率,
    #     22=财报退货率, 26=广告点击率
    # 按 ISO周 + SKU 汇总: visitors, add_to_cart_count, sales_qty, click_cnt, return_qty, total_qty
    weekly_agg = defaultdict(lambda: defaultdict(lambda: {
        "visitors": 0, "atc": 0, "qty": 0,
        "click_cnt": 0, "click_impressions": 0,
        "return_qty": 0, "total_qty_ref": 0
    }))

    # 同时追踪每个SKU首次出现库存的日期（用于上架天数计算）
    sku_first_inventory_date = {}
    # 追踪每个SKU最新日期的可售数量（E列）
    sku_latest_inventory = {}  # sku -> {'date': date, 'value': float}

    for i, row in enumerate(ws.iter_rows(min_row=2)):
        if i > 200000:
            break
        if not row:
            continue

        d_val = row[0].value if len(row) > 0 else None  # 日期
        sku = row[3].value if len(row) > 3 else None     # 卖家SKU
        inventory = row[4].value if len(row) > 4 else None  # E列 可售数量
        visitors = row[7].value if len(row) > 7 else None  # 访客
        atc = row[10].value if len(row) > 10 else None     # 加购数
        qty = row[12].value if len(row) > 12 else None     # 销量
        click_rate_val = row[26].value if len(row) > 26 else None  # 广告点击率 (decimal)
        return_rate_raw = row[22].value if len(row) > 22 else None  # 财报退货率 (decimal)

        if not d_val or not sku:
            continue

        if isinstance(d_val, datetime):
            d_date = d_val.date()
        elif isinstance(d_val, date):
            d_date = d_val
        elif isinstance(d_val, str):
            try:
                d_date = datetime.strptime(d_val[:10], "%Y-%m-%d").date()
            except:
                continue
        else:
            continue

        iso = date_to_iso_week(d_date)
        w_key = iso_to_w.get(iso)
        if w_key is None:
            continue

        sku_str = str(sku).strip()
        agg = weekly_agg[w_key][sku_str]
        agg["visitors"] += float(visitors) if visitors else 0
        agg["atc"] += float(atc) if atc else 0
        agg["qty"] += float(qty) if qty else 0

        # 追踪每个SKU首次出现库存>0的日期（E列可售数量>0）
        if inventory is not None:
            try:
                inv_val = float(inventory)
                if inv_val > 0:
                    if sku_str not in sku_first_inventory_date or d_date < sku_first_inventory_date[sku_str]:
                        sku_first_inventory_date[sku_str] = d_date
                # 追踪最新日期的可售数量
                if sku_str not in sku_latest_inventory or d_date > sku_latest_inventory[sku_str]['date']:
                    sku_latest_inventory[sku_str] = {'date': d_date, 'value': inv_val}
            except (ValueError, TypeError):
                pass

        # 广告点击率: 累积每天的点击率值, 用于后续计算加权平均
        if click_rate_val is not None and visitors:
            click_rate_f = float(click_rate_val)
            if not (math.isnan(click_rate_f) or math.isinf(click_rate_f)):
                agg["click_cnt"] += click_rate_f * float(visitors)
                agg["click_impressions"] += float(visitors)

    # 转换为 TRAFFIC_WEEKLY 格式
    # [click_rate, add_to_cart_rate, conversion_rate, return_rate, sales_qty]
    traffic_weekly = {}
    for w_key, sku_data in sorted(weekly_agg.items(), key=lambda x: int(x[0][1:])):
        traffic_weekly[w_key] = {}
        for sku, agg in sku_data.items():
            v = agg["visitors"]
            # click_rate: 加权平均 (sum(click_rate * visitors) / sum(visitors))
            if agg["click_impressions"] > 0:
                click_rate = round(agg["click_cnt"] / agg["click_impressions"], 6)
            else:
                click_rate = None
            atc_rate = round(agg["atc"] / v, 6) if v > 0 else 0.0
            conv_rate = round(agg["qty"] / v, 6) if v > 0 else 0.0
            # 财报退货率暂不在 TRAFFIC_WEEKLY 中, 用 null 占位
            return_rate = None
            sales_qty = round(agg["qty"], 0)
            traffic_weekly[w_key][sku] = [click_rate, atc_rate, conv_rate, return_rate, sales_qty]

    wb.close()
    total_entries = sum(len(v) for v in traffic_weekly.values())
    print(f"  TRAFFIC_WEEKLY: {len(traffic_weekly)} weeks, {total_entries} total SKU-week entries")
    print(f"  SKU首次库存日期 (运营日数据): {len(sku_first_inventory_date)} 个SKU")
    print(f"  SKU最新可售数量 (运营日数据): {len(sku_latest_inventory)} 个SKU")
    return traffic_weekly, sku_first_inventory_date, sku_latest_inventory


# ── 阶段 5: 读取年规进度 → PERSON_TARGETS ─────────────

def read_person_targets():
    """从2026WB年规进度.xlsx 生成 PERSON_TARGETS"""
    path = os.path.join(SRC_DIR, "2026WB年规进度.xlsx")
    wb = load_workbook_safe(path)
    if wb is None:
        return {}

    # 列映射: Excel列 → 月份
    # Col 2=2月, 3=3月, 4=4月, 5=5月, 6=6月, 7=7月, 8=8月,
    # 9=9月, 10=10月, 11=11月, 12=12月, 13=1月
    col_to_month = {2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 7, 8: 8,
                    9: 9, 10: 10, 11: 11, 12: 12, 13: 1}

    # 行到字段的映射（不同sheet可能有不同结构）
    # 我们需要从列B识别
    all_targets = {}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        person_name = sheet_name.strip()

        # 读取所有行
        rows_data = {}
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=100)):
            vals = [cell.value for cell in row[:16]]
            rows_data[i+1] = vals

        # 从B列识别指标
        targets_by_month = defaultdict(dict)

        shop_section_started = False
        for row_num, vals in rows_data.items():
            # 遇到第一个店铺名行后，停止处理（后续都是逐店明细）
            a_val = vals[0] if len(vals) > 0 else None
            if a_val and isinstance(a_val, str) and "店" in a_val:
                shop_section_started = True
                continue
            if shop_section_started:
                continue  # 已进入逐店排分区，跳过所有后续行
            label = str(vals[1]).strip() if vals[1] is not None else ""
            if not label:
                continue

            # 月度目标 利润 → profit_target
            if "月度目标" in label and "利润" in label:
                for col, month in col_to_month.items():
                    v = vals[col]
                    if v is not None and month >= 1 and month <= 12:
                        targets_by_month[month]["profit_target"] = sanitize_value(v) or 0

            # 实际利润 → profit_done（排除 "利润率" 行）
            elif "实际利润" in label and "率" not in label:
                for col, month in col_to_month.items():
                    v = vals[col]
                    if v is not None and month >= 1 and month <= 12:
                        targets_by_month[month]["profit_done"] = sanitize_value(v) or 0

            # 销量目标 → sales_target（排除 "销量完成进度"）
            elif label == "销量目标":
                for col, month in col_to_month.items():
                    v = vals[col]
                    if v is not None and month >= 1 and month <= 12:
                        targets_by_month[month]["sales_target"] = sanitize_value(v) or 0

            # 销量完成 → sales_done（排除 "销量完成进度"）
            elif label == "销量完成":
                for col, month in col_to_month.items():
                    v = vals[col]
                    if v is not None and month >= 1 and month <= 12:
                        targets_by_month[month]["sales_done"] = sanitize_value(v) or 0

            # GMV目标
            elif label == "GMV目标":
                for col, month in col_to_month.items():
                    v = vals[col]
                    if v is not None and month >= 1 and month <= 12:
                        targets_by_month[month]["gmv_target"] = sanitize_value(v) or 0

            # GMV完成
            elif label == "GMV完成":
                for col, month in col_to_month.items():
                    v = vals[col]
                    if v is not None and month >= 1 and month <= 12:
                        targets_by_month[month]["gmv_done"] = sanitize_value(v) or 0

            # GSV目标
            elif label == "GSV目标":
                for col, month in col_to_month.items():
                    v = vals[col]
                    if v is not None and month >= 1 and month <= 12:
                        targets_by_month[month]["gsv_target"] = sanitize_value(v) or 0

            # GSV完成
            elif label == "GSV完成":
                for col, month in col_to_month.items():
                    v = vals[col]
                    if v is not None and month >= 1 and month <= 12:
                        targets_by_month[month]["gsv_done"] = sanitize_value(v) or 0

        # 计算 profit_rate
        for month in targets_by_month:
            t = targets_by_month[month]
            gsv_done = t.get("gsv_done", 0)
            profit_done = t.get("profit_done", 0)
            if gsv_done and gsv_done > 0 and profit_done:
                t["profit_rate"] = round(profit_done / gsv_done * 100, 2)
            else:
                t["profit_rate"] = 0

        # 确保所有12个月都有数据
        for m in range(1, 13):
            if m not in targets_by_month:
                targets_by_month[m] = {
                    "gsv_target": 0, "gsv_done": 0,
                    "gmv_target": 0, "gmv_done": 0,
                    "sales_target": 0, "sales_done": 0,
                    "profit_target": 0, "profit_done": 0,
                    "profit_rate": 0
                }

        # 转换 key 为字符串
        person_data = {str(m): dict(targets_by_month[m]) for m in range(1, 13)}
        all_targets[person_name] = person_data

    wb.close()
    print(f"  PERSON_TARGETS: {len(all_targets)} people")
    return all_targets


# ── 主函数 ────────────────────────────────────────────

def main():
    print("=" * 60)
    print("build_data.py - 生成 data.json")
    print("=" * 60)

    # 阶段 1: 基础周数组
    print("\n[1/5] 生成周数组...")
    weeks, weeks_iso, week_labels = generate_weeks()
    print(f"  共 {len(weeks)} 周: {weeks_iso[0]} → {weeks_iso[-1]}")

    # 阶段 2: 产品列表
    print("\n[2/5] 读取产品列表...")
    sku_img, sku_owner, sku_first_date, sku_wb_id, shop_owners = read_product_list()

    # 阶段 3: 利润表
    print("\n[3/5] 读取LX利润表...")
    shop_weekly, week_data = read_lx_profit(weeks_iso)

    # 阶段 4: 运营日数据
    print("\n[4/5] 读取运营日数据...")
    traffic_weekly, sku_first_inventory_date, sku_latest_inventory = read_traffic_weekly(weeks_iso)

    # 阶段 4.5: 将运营日数据 M列(销量) 合并到 WEEK_DATA.qty
    print("  合并运营日数据 M列(销量) 到 WEEK_DATA...")
    qty_merged = 0
    for w_key in week_data:
        if w_key not in traffic_weekly:
            continue
        sku_qty = traffic_weekly[w_key]  # dict: sku -> [click_rate, atc_rate, conv_rate, return_rate, sales_qty]
        for p in week_data[w_key]["allProducts"]:
            sku = p["sku"]
            if sku in sku_qty and len(sku_qty[sku]) > 4:
                new_qty = sku_qty[sku][4]  # index 4 = sales_qty from 运营日数据 M列
                if new_qty and new_qty > 0:
                    p["qty"] = new_qty
                    qty_merged += 1
    print(f"  已合并 {qty_merged} 条销量数据 (来源: 运营日数据 M列)")

    # 构建 SKU 到 WB_IDs 的映射
    sku_to_wb_ids = {}
    for key in sku_wb_id:
        parts = key.split('|')
        if len(parts) >= 2:
            sku = parts[0]
            wb_id = sku_wb_id[key]
            if sku not in sku_to_wb_ids:
                sku_to_wb_ids[sku] = set()
            sku_to_wb_ids[sku].add(wb_id)
    
    # SKU_FIRST_DATE: 合并两个来源
    # 1. 产品列表中的创建时间（已有 SKU+WB_ID 组合 key）
    # 2. 运营日数据首次库存>0的日期（仅 SKU key，作为补充/优先数据）
    merged_sku_first_date = {}
    
    # 先从运营日数据获取首次库存>0日期（优先级更高，因为反映真实上架）
    for sku, d in sku_first_inventory_date.items():
        date_str = d.strftime("%Y-%m-%d") if isinstance(d, date) else str(d)[:10]
        # 仅 SKU 级别 key
        merged_sku_first_date[sku] = date_str
        # 同时为每个 WB_ID 变体设置相同日期（因为运营日数据没有 WB_ID 区分）
        if sku in sku_to_wb_ids:
            for wb_id in sku_to_wb_ids[sku]:
                merged_sku_first_date[f"{sku}|{wb_id}"] = date_str
    
    # 再从产品列表补充（仅补充运营日数据中没有的）
    for key, date_str in sku_first_date.items():
        if key not in merged_sku_first_date:
            merged_sku_first_date[key] = date_str
    
    print(f"  SKU_FIRST_DATE: {len(merged_sku_first_date)} 个 (合并产品列表+运营日数据)")
    print(f"    无库存记录的SKU将不显示上架天数（显示为 '-'）")

    # SKU_INVENTORY: 每个SKU最新日期的可售数量
    sku_inventory = {}
    for sku, info in sku_latest_inventory.items():
        sku_inventory[sku] = int(info['value'])
    print(f"  SKU_INVENTORY: {len(sku_inventory)} 个SKU (最新日期可售数量)")

    # 阶段 5: 年规进度
    print("\n[5/5] 读取年规进度...")
    person_targets = read_person_targets()

    # ── 组装 data.json ──
    print("\n" + "=" * 60)
    print("组装 data.json...")

    data = {
        "WEEKS": weeks,
        "WEEKS_ISO": weeks_iso,
        "WEEK_LABELS": week_labels,
        "SHOP_WEEKLY": shop_weekly,
        "WEEK_DATA": week_data,
        "TRAFFIC_WEEKLY": traffic_weekly,
        "PERSON_TARGETS": person_targets,
        "SKU_IMG": sku_img,
        "SKU_FIRST_DATE": merged_sku_first_date,
        "SKU_OWNER": sku_owner,
        "SHOP_OWNERS": shop_owners,
        "SKU_WB_ID": sku_wb_id,
        "PRODUCT_NOTES": {},
        "NEW_PRODUCT_CREATED": [],
        "SKU_INVENTORY": sku_inventory,
    }

    # ── 写入 data.js ──
    # 确保 output 目录存在
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)

    # 自定义 JSON encoder 处理特殊值
    class SanitizedEncoder(json.JSONEncoder):
        def default(self, obj):
            if isinstance(obj, float):
                if math.isnan(obj) or math.isinf(obj):
                    return None
            return super().default(obj)

    json_str = json.dumps(data, ensure_ascii=False, separators=(',', ':'), cls=SanitizedEncoder)
    # 二次检查：确保没有任何 NaN/Infinity 出现在 JSON 中
    json_str = re.sub(r':NaN', ':null', json_str)
    json_str = re.sub(r':-Infinity', ':null', json_str)
    json_str = re.sub(r':Infinity', ':null', json_str)

    # 包装为 JS 变量声明（通过 <script src="data.js"> 加载，无需 fetch/XHR）
    js_content = f'var DATA = {json_str};'

    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        f.write(js_content)

    file_size = os.path.getsize(OUTPUT_PATH)
    print(f"  输出: {OUTPUT_PATH}")
    print(f"  大小: {file_size / 1024 / 1024:.2f} MB")

    # ── 验证 data.js 有效性 ──
    print("\n验证 data.js 有效性...")
    try:
        # 提取 JSON 部分（去除 var DATA =  和尾部 ;）
        if js_content.startswith('var DATA = ') and js_content.endswith(';'):
            inner = js_content[11:-1]
        else:
            inner = js_content
        verified = json.loads(inner)
        print(f"  ✓ DATA 有效: {len(verified)} 个顶级字段")
        for key in verified:
            v = verified[key]
            if isinstance(v, dict):
                print(f"    {key}: {len(v)} entries (dict)")
            elif isinstance(v, list):
                print(f"    {key}: {len(v)} entries (list)")
            else:
                print(f"    {key}: {type(v).__name__}")
    except json.JSONDecodeError as e:
        print(f"  ❌ DATA 无效: {e}")
        return 1

    # ── 自动更新 HTML 中 data.js 的版本号（强制浏览器刷新缓存）──
    html_path = os.path.join(OUTPUT_DIR, "product-weekly-report.html")
    if os.path.exists(html_path):
        new_ver = datetime.now().strftime("%Y%m%d%H%M")
        with open(html_path, 'r', encoding='utf-8') as f:
            html = f.read()
        updated_html = re.sub(r'data\.js\?v=\d+', f'data.js?v={new_ver}', html)
        if updated_html != html:
            with open(html_path, 'w', encoding='utf-8') as f:
                f.write(updated_html)
            print(f"\n  已更新 HTML 缓存版本号: ?v={new_ver}")
        else:
            print(f"\n  ⚠ HTML 中未找到 data.js?v= 模式，请手动检查")

    # ── 数据质量诊断 ──
    print("\n数据质量诊断:")
    # 检查 return_rate 是否为百分比
    rr_check = 0
    rr_small = 0
    qty_from_m = 0
    for wk in verified["WEEK_DATA"]:
        for p in verified["WEEK_DATA"][wk].get("allProducts", []):
            rr = p.get("return_rate", 0)
            if rr and rr > 0:
                rr_check += 1
                if rr < 1:
                    rr_small += 1
            if p.get("qty", 0) > 0:
                qty_from_m += 1
    print(f"  return_rate 百分比格式: {rr_check - rr_small}/{rr_check} (小数值 {rr_small})")
    print(f"  qty (来源 M列): {qty_from_m} 个产品")
    print(f"  TRAFFIC_WEEKLY: {len(verified['TRAFFIC_WEEKLY'])} 周")
    print(f"  SKU_OWNER: {len(verified['SKU_OWNER'])} 个 SKU")

    print("\n" + "=" * 60)
    print("完成!")
    return 0


if __name__ == "__main__":
    sys.exit(main())