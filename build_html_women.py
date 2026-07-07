"""
build_html_women.py - 女装版：一键从 Excel 生成带内嵌数据的 product-weekly-report-women.html

用法:
    python build_html_women.py

流程:
    1. 运行 build_data_women.py 生成 data_women.js
    2. 读取 data_women.js 并解析 JSON
    3. 将数据拆分为 CORE_DATA（轻量元数据）和 DETAIL_DATA（周数据+流量数据）
    3.5 从 WEEK_DATA 生成 RAW_SALES_DATA / RAW_PROFIT_DATA
    4. 替换 HTML 中三个标记块（CORE_DATA / CORE_DETAIL / SABC）
    5. 输出最终 HTML
"""

import json
import os
import subprocess
import sys

SRC_DIR = r"D:\周汇报文件"
DATA_JS = os.path.join(SRC_DIR, "data_women.js")
HTML_PATH = os.path.join(SRC_DIR, "product-weekly-report-women.html")

DETAIL_FIELDS = {"WEEK_DATA", "TRAFFIC_WEEKLY"}


def main():
    print("=" * 60)
    print("build_html_women.py - 生成女装版内嵌数据 HTML")
    print("=" * 60)

    # ── Step 1: 运行 build_data_women.py ──
    print("\n[1/4] 运行 build_data_women.py ...")
    result = subprocess.run(
        [sys.executable, os.path.join(SRC_DIR, "build_data_women.py")],
        capture_output=True, text=True, cwd=SRC_DIR
    )
    print(result.stdout)
    if result.returncode != 0:
        print(f"[ERROR] build_data_women.py 失败:\n{result.stderr}")
        return 1

    # ── Step 2: 读取 data_women.js ──
    print("\n[2/4] 读取 data_women.js ...")
    with open(DATA_JS, 'r', encoding='utf-8') as f:
        js_content = f.read()

    prefix = 'var DATA = '
    if not js_content.startswith(prefix) or not js_content.rstrip().endswith(';'):
        print("[ERROR] data_women.js 格式非预期，期望 'var DATA = {...};'")
        return 1

    json_str = js_content[len(prefix):-1]
    data = json.loads(json_str)
    print(f"  解析成功: {len(data)} 个顶级字段")

    # ── Step 3: 拆分 CORE_DATA / DETAIL_DATA ──
    print("\n[3/4] 拆分数据并生成 JS 声明 ...")
    core_data = {}
    detail_data = {}
    for key, value in data.items():
        if key in DETAIL_FIELDS:
            detail_data[key] = value
        else:
            core_data[key] = value

    core_json = json.dumps(core_data, ensure_ascii=False, separators=(',', ':'))
    detail_json = json.dumps(detail_data, ensure_ascii=False, separators=(',', ':'))

    print(f"  CORE_DATA: {len(core_data)} 字段, {len(core_json)} 字符")
    print(f"  DETAIL_DATA: {len(detail_data)} 字段, {len(detail_json)} 字符")

    # ── Step 3.5: 从 WEEK_DATA 生成 RAW_SALES_DATA / RAW_PROFIT_DATA ──
    print("\n[3.5/4] 从 WEEK_DATA 生成 SABC 数据 ...")
    week_data = data.get('WEEK_DATA', {})
    raw_sales = {}
    raw_profit = {}
    for week, wd in week_data.items():
        # allProducts 和 top10Profit 合并，确保 TOP10 产品也参与 SABC 标色计算
        seen_keys = set()
        raw_sales[week] = {}
        raw_profit[week] = {}
        for p in wd.get('allProducts', []) + wd.get('top10Profit', []):
            shop = p.get('shop', '')
            sku = p.get('sku', '')
            if not shop or not sku:
                continue
            key = f"{shop}|{sku}"
            if key in seen_keys:
                continue
            seen_keys.add(key)
            raw_sales[week][key] = int(p.get('qty', 0) or 0)
            raw_profit[week][key] = [
                round(float(p.get('profit', 0) or 0), 2),
                round(float(p.get('gsv', 0) or 0), 2),
                round(float(p.get('margin', 0) or 0), 4)
            ]
    raw_sales_json = json.dumps(raw_sales, ensure_ascii=False, separators=(',', ':'))
    raw_profit_json = json.dumps(raw_profit, ensure_ascii=False, separators=(',', ':'))
    print(f"  RAW_SALES_DATA: {len(raw_sales)} 周, {len(raw_sales_json)} 字符")
    print(f"  RAW_PROFIT_DATA: {len(raw_profit)} 周, {len(raw_profit_json)} 字符")

    def week_sort_key(w):
        try:
            return int(w[1:])
        except (ValueError, IndexError):
            return 0

    source_weeks = sorted(week_data.keys(), key=week_sort_key)
    source_count = len(source_weeks)
    if len(raw_sales) != source_count:
        print(f"[FATAL] RAW_SALES_DATA 周数 ({len(raw_sales)}) != data_women.js WEEK_DATA ({source_count})")
        return 1
    if len(raw_profit) != source_count:
        print(f"[FATAL] RAW_PROFIT_DATA 周数 ({len(raw_profit)}) != data_women.js WEEK_DATA ({source_count})")
        return 1
    print(f"  生成期验证通过: {source_count} 周一致")

    # ── Step 4: 替换 HTML ──
    print("\n[4/4] 替换 HTML 内嵌数据 ...")
    with open(HTML_PATH, 'r', encoding='utf-8') as f:
        html = f.read()

    start_marker = '/* ===DATA_START=== */'
    end_marker = '/* ===DATA_END=== */'

    idx_start = html.find(start_marker)
    idx_end = html.find(end_marker)

    if idx_start < 0 or idx_end < 0:
        print("[ERROR] HTML 中未找到数据标记!")
        return 1

    new_block = (
        f"{start_marker}\n"
        f"var CORE_DATA = {core_json};\n"
        f"{end_marker}"
    )

    new_html = html[:idx_start] + new_block + html[idx_end + len(end_marker):]

    # ── 替换 CORE_DETAIL 块 ──
    detail_start_marker = '/* ===DATA_DETAIL_START=== */'
    detail_end_marker = '/* ===DATA_DETAIL_END=== */'

    idx_ds = new_html.find(detail_start_marker)
    idx_de = new_html.find(detail_end_marker)

    if idx_ds >= 0 and idx_de >= 0:
        detail_block = (
            f"{detail_start_marker}\n"
            f"var CORE_DETAIL = {{}}; /* loaded async from data-detail-women.js */\n"
            f"{detail_end_marker}"
        )
        new_html = new_html[:idx_ds] + detail_block + new_html[idx_de + len(detail_end_marker):]
        print(f"  CORE_DETAIL 块已替换为异步加载占位")
    else:
        print(f"  [WARNING] 未找到 CORE_DETAIL 标记，跳过内嵌更新")

    # ── 替换 SABC 块 ──
    sabc_start = '// ========== SABC AUTO-GRADING DATA =========='
    sabc_end = '// ========== END SABC DATA =========='

    idx_ss = new_html.find(sabc_start)
    idx_se = new_html.find(sabc_end, idx_ss)

    if idx_ss >= 0 and idx_se >= 0:
        idx_se += len(sabc_end)
        sabc_block = (
            f"{sabc_start}\n"
            f"var RAW_SALES_DATA = {raw_sales_json};\n"
            f"var RAW_PROFIT_DATA = {raw_profit_json};\n"
            f"{sabc_end}"
        )
        new_html = new_html[:idx_ss] + sabc_block + new_html[idx_se:]
        print(f"  SABC 块已更新")
    else:
        print(f"  [WARNING] 未找到 SABC 标记，跳过 RAW 数据更新")

    # ── Step 5: 输出前验证 ──
    print("\n[5/4] 输出前验证 ...")
    errors = []

    def extract_json_from_html(html_str, var_name):
        marker = f'var {var_name} = '
        idx = html_str.find(marker)
        if idx < 0:
            return None, f"未找到 {var_name}"
        json_start = idx + len(marker)
        brace_count = 0
        in_str = False
        escape = False
        for i in range(json_start, len(html_str)):
            ch = html_str[i]
            if escape:
                escape = False
                continue
            if ch == '\\':
                escape = True
                continue
            if ch == '"' and not escape:
                in_str = not in_str
                continue
            if in_str:
                continue
            if ch == '{':
                brace_count += 1
            elif ch == '}':
                brace_count -= 1
                if brace_count == 0:
                    try:
                        return json.loads(html_str[json_start:i + 1]), None
                    except json.JSONDecodeError as e:
                        return None, f"JSON 解析失败: {e}"
        return None, "未找到闭合的大括号"

    # Verification
    for name, json_part in [("CORE_DATA", core_json), ("RAW_SALES_DATA", raw_sales_json), ("RAW_PROFIT_DATA", raw_profit_json)]:
        braces = 0
        brackets = 0
        in_str = False
        escape = False
        for ch in json_part:
            if escape:
                escape = False
                continue
            if ch == '\\':
                escape = True
                continue
            if ch == '"' and not escape:
                in_str = not in_str
                continue
            if in_str:
                continue
            if ch == '{':
                braces += 1
            elif ch == '}':
                braces -= 1
            elif ch == '[':
                brackets += 1
            elif ch == ']':
                brackets -= 1
        if braces != 0 or brackets != 0:
            errors.append(f"JS语法 {name}: braces={braces}, brackets={brackets}")
        else:
            print(f"  OK JS语法 {name}: braces=0, brackets=0")

    if errors:
        print(f"\n验证失败 ({len(errors)} 项)，拒绝输出:")
        for e in errors:
            print(f"  FAIL {e}")
        return 1

    print("\n所有验证通过，写入输出文件...")

    with open(HTML_PATH, 'w', encoding='utf-8') as f:
        f.write(new_html)

    file_size = os.path.getsize(HTML_PATH)
    print(f"  输出: {HTML_PATH}")
    print(f"  大小: {file_size / 1024:.0f} KB")

    # ── 生成 data-detail-women.js ──
    detail_js_path = os.path.join(SRC_DIR, "data-detail-women.js")
    detail_js_content = f"var DETAIL_DATA = {detail_json};"
    with open(detail_js_path, 'w', encoding='utf-8') as f:
        f.write(detail_js_content)
    detail_size = os.path.getsize(detail_js_path)
    print(f"  生成: {detail_js_path}")
    print(f"  大小: {detail_size / 1024:.0f} KB")
    
    print("\n完成!")
    return 0


if __name__ == "__main__":
    sys.exit(main())
