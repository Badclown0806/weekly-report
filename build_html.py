"""
build_html.py - 一键从 Excel 生成带内嵌数据的 product-weekly-report.html

用法:
    python build_html.py

流程:
    1. 运行 build_data.py 生成 data.js
    2. 读取 data.js 并解析 JSON
    3. 将数据拆分为 CORE_DATA（轻量元数据）和 DETAIL_DATA（周数据+流量数据）
    4. 替换 HTML 中标记块 /* ===DATA_START=== */ ~ /* ===DATA_END=== */
    5. 输出最终 HTML
"""

import json
import os
import subprocess
import sys

SRC_DIR = r"D:\周汇报文件"
DATA_JS = os.path.join(SRC_DIR, "data.js")
HTML_PATH = os.path.join(SRC_DIR, "product-weekly-report.html")

# 哪些字段属于 DETAIL_DATA（体积大、非首屏必要）
DETAIL_FIELDS = {"WEEK_DATA", "TRAFFIC_WEEKLY"}


def main():
    print("=" * 60)
    print("build_html.py - 生成内嵌数据版 HTML")
    print("=" * 60)

    # ── Step 1: 运行 build_data.py ──
    print("\n[1/4] 运行 build_data.py ...")
    result = subprocess.run(
        [sys.executable, os.path.join(SRC_DIR, "build_data.py")],
        capture_output=True, text=True, cwd=SRC_DIR
    )
    print(result.stdout)
    if result.returncode != 0:
        print(f"[ERROR] build_data.py 失败:\n{result.stderr}")
        return 1

    # ── Step 2: 读取 data.js ──
    print("\n[2/4] 读取 data.js ...")
    with open(DATA_JS, 'r', encoding='utf-8') as f:
        js_content = f.read()

    # 提取 JSON：var DATA = {...};
    prefix = 'var DATA = '
    if not js_content.startswith(prefix) or not js_content.rstrip().endswith(';'):
        print("[ERROR] data.js 格式非预期，期望 'var DATA = {...};'")
        return 1

    json_str = js_content[len(prefix):-1]  # 去掉前缀和末尾分号
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
        print("  请确保 HTML 包含 /* ===DATA_START=== */ 和 /* ===DATA_END=== */")
        return 1

    # 替换标记之间的内容（只内嵌 CORE_DATA）
    new_block = (
        f"{start_marker}\n"
        f"var CORE_DATA = {core_json};\n"
        f"{end_marker}\n"
        f"var DETAIL_DATA = null;"
    )

    new_html = html[:idx_start] + new_block + html[idx_end + len(end_marker):]

    with open(HTML_PATH, 'w', encoding='utf-8') as f:
        f.write(new_html)

    file_size = os.path.getsize(HTML_PATH)
    print(f"  输出: {HTML_PATH}")
    print(f"  大小: {file_size / 1024:.0f} KB")

    # ── 生成 data-detail.js（外置异步加载） ──
    detail_js_path = os.path.join(SRC_DIR, "data-detail.js")
    detail_js_content = f"var DETAIL_DATA = {detail_json};"
    with open(detail_js_path, 'w', encoding='utf-8') as f:
        f.write(detail_js_content)
    detail_size = os.path.getsize(detail_js_path)
    print(f"  生成: {detail_js_path}")
    print(f"  大小: {detail_size / 1024:.0f} KB")

    # ── 验证：JS 语法自检（括号平衡） ──
    print("\n验证 JS 语法...")
    for name, json_part in [("CORE_DATA", core_json), ("DETAIL_DATA", detail_json)]:
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
            print(f"  FAIL {name}: braces={braces}, brackets={brackets}")
            return 1
        print(f"  OK   {name}: braces={braces}, brackets={brackets}")
    
    print("\n完成!")
    return 0


if __name__ == "__main__":
    sys.exit(main())