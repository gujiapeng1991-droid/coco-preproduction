#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_scene_breakdown.py — 把结构化分场数据渲染成「Coco影视顺场表」Excel（v1.0）。

与分镜表的根本差别：**一场一行**，场号是顺序纯数字（1 / 2 / 3），不带字母后缀。
单场内换了位置、氛围、时间档或内外，直接另起一场。

脚本自动补一列：
  - 主场景：从「场景」列的大景解析出来，排在场景列之前，用于归组统计
角色每人一列打 ✓：顶层 roles 声明全部角色，替换掉 columns 里的 cast 位置；
不写 cast 时，角色列组自动插在「页数」列**之后**。

列顺序（v1.2 起）：集·场号·日/夜·内/外·主场景·场景·氛围·内容梗概·页数
                ·角色列组(每人一列 ✓)·特约·群众·道具·视效·备注
不再输出「时长(秒)」「累计页数」两列——时长由页数估算（1 页 ≈ 1 分钟），
累计页数对统筹没用、还容易和页数看混。旧数据里若仍带这两列会被自动剔除。

输入：
  - JSON 文件（推荐）：title / meta / columns(字段key) / rows(对象数组)
  - CSV 文件：首行为中文表头
  - stdin（用 - 作输入路径）

输出：
  - Sheet「顺场表」：一场一行，日/夜与内/外配色，末行合计
  - Sheet「场景统计」：总览 / 日·夜分布 / 内·外分布 / 按主场景归组（页数降序）
  - 可选 --html：同时导出可打印 HTML（仅主表）

依赖：openpyxl（缺失时自动 pip 安装到当前解释器）。

用法：
  python make_scene_breakdown.py data.json -o 顺场表.xlsx
  python make_scene_breakdown.py data.csv --html
  cat data.json | python make_scene_breakdown.py - -o out.xlsx
"""

import argparse
import csv
import io
import json
import re
import sys
from collections import OrderedDict
from pathlib import Path

# ---- 依赖自检 ---------------------------------------------------------------
try:
    import openpyxl
except ImportError:  # pragma: no cover
    import subprocess
    print("[info] 未检测到 openpyxl，正在安装…", file=sys.stderr)
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "openpyxl"])
    import openpyxl

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# 字段 key -> 中文表头
FIELD_CN = {
    "ep": "集",
    "scene": "场号",
    "slug": "场次名",
    "inout": "内/外",
    "daynight": "日/夜",
    "set": "场景",
    "mainset": "主场景",
    "mood": "氛围",
    "synopsis": "内容梗概",
    "cast": "出场人物",
    "pages": "页数",
    "extra": "特约",
    "mass": "群众",
    "props": "道具",
    "vfx": "视效",
    "note": "备注",
}
# v1.2 起移除的两列（旧数据里若还写着，会被自动剔除而不是照原样渲染）
DROPPED_FIELDS = {"sec": "时长(秒)", "cum": "累计页数"}

# 列宽
WIDTHS = {
    "集": 4.4,
    "场号": 6.5,
    "场次名": 22,
    "日/夜": 5,
    "内/外": 5,
    "场景": 16,
    "主场景": 13,
    "氛围": 8,
    "内容梗概": 46,
    "出场人物": 14,
    "页数": 5.2,
    "特约": 11,
    "群众": 11,
    "道具": 18,
    "视效": 14,
    "备注": 20,
}

# 配色
NAVY = "1F3864"          # 表头深蓝
TITLE_FILL = "D9E1F2"    # 标题浅蓝
MAINSET_FILL = "E8EEF8"  # 主场景列底
ZEBRA = "F7F9FC"         # 隔行底色
TOTAL_FILL = "FFF2CC"    # 合计行暖黄
MOOD_FILL = "EDE7F6"     # 氛围列浅紫
DAY_FILLS = {"日": "FFF3CD", "夜": "D6E0F5", "晨": "FFF9E6", "黄昏": "FDE7D0"}
IO_FILLS = {"内": "E8F1E8", "外": "E7F0F7", "内外": "EFEAF7"}
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# 角色独占列——与分镜表 skill 同一套规范
# v1.3：表头写角色全名，格子里填单字缩写（比 ✓ 更好扫，一眼看出谁在场）
CHECK = "✓"
ROLE_COLS = set()   # 表头上显示的名字（优先全名），用于窄列宽与统计
ROLE_ABBR = {}      # 表头显示名 -> 格子里填的单字缩写

# 场号解析（兼容残留的 1A 写法，仅用于统计去重）
SCENE_RE = re.compile(r"^([^0-9A-Za-z]*)(\d+)([A-Za-z]*)$")


# ---------------------------------------------------------------------------
def split_scene(v):
    """'3' -> ('3', '')；'1A' -> ('1', 'A')（兼容旧数据）。"""
    s = str(v).strip()
    m = SCENE_RE.match(s)
    if not m:
        return s, ""
    prefix, num, suffix = m.groups()
    return f"{prefix}{num}", suffix.upper()


def big_set(v):
    """'陈默家-客厅' -> '陈默家'（主场景）。全角/半角短横线都认。"""
    s = str(v).strip().replace("－", "-").replace("–", "-")
    if "-" in s:
        return s.split("-", 1)[0].strip() or s
    return s.strip()


def load_data(path):
    """读取 JSON / CSV，返回 (title, meta_text, headers_cn, rows, roles)。"""
    raw = sys.stdin.read() if path == "-" else Path(path).read_text(encoding="utf-8")

    if path == "-" or path.lower().endswith(".json"):
        data = json.loads(raw)
        title = data.get("title", "顺场表")
        meta = data.get("meta", "") or ""
        if isinstance(meta, dict):
            meta_text = "　".join(f"{k}：{v}" for k, v in meta.items())
        elif isinstance(meta, (list, tuple)):
            meta_text = "　".join(str(m) for m in meta)
        else:
            meta_text = str(meta)
        if "columns" in data and "rows" in data:
            keys = list(data["columns"])
            # v1.2 起不再输出时长(秒)/累计页数：key 阶段直接剔除
            # （CSV/中文表头的情况由 drop_legacy_cols 兜底）
            gone = [k for k in keys if k in DROPPED_FIELDS]
            if gone:
                print("[info] 已移除不再输出的列："
                      + "、".join(DROPPED_FIELDS[k] for k in gone), file=sys.stderr)
                keys = [k for k in keys if k not in DROPPED_FIELDS]
            headers_cn = [FIELD_CN.get(c, str(c)) for c in keys]
            # 角色解析成 (缩写, 表头显示名)；roles 没给全名时自动读 meta.角色缩写
            roles = parse_role_map(data.get("roles", []) or [],
                                   meta if isinstance(meta, dict) else {})
            rows = []
            for r in data["rows"]:
                row = [r.get(c, "") for c in keys]
                if roles:
                    v = r.get("roles", [])
                    if isinstance(v, str):
                        v = [p.strip() for p in v.replace("／", "/").split("/") if p.strip()]
                    row.append(v)   # 尾部携带出场角色，expand_roles 消费
                rows.append(row)
        elif "headers" in data and "rows" in data:
            headers_cn = list(data["headers"])
            rows = [list(r) for r in data["rows"]]
            roles = []
        else:
            raise ValueError("JSON 需含 columns+rows 或 headers+rows")
        return title, meta_text, headers_cn, rows, roles

    reader = csv.reader(io.StringIO(raw))
    lines = list(reader)
    if not lines:
        raise ValueError("CSV 为空")
    title = Path(path).stem if path != "-" else "顺场表"
    return title, "", lines[0], lines[1:], []


def parse_role_map(roles, meta=None):
    """把 roles 解析成 [(缩写, 表头显示名)]。

    支持四种写法（越靠前优先级越高）：
      1. ["四=阿萤", "枫=裴决"]            缩写=全名（推荐，与 meta.角色缩写 同款写法）
      2. [{"abbr":"四","name":"阿萤"}, …]  显式对象
      3. {"四": "阿萤", …}                 映射对象
      4. ["四", "枫"]                      只有缩写 —— 表头就显示缩写（向后兼容）

    roles 没给全名时，自动去 meta["角色缩写"] 里找（形如
    `四=阿萤(方四/沈萤，女主) / 枫=裴决(大理寺少卿，男主)`），取括号前那段作全名。
    """
    pairs = []
    if isinstance(roles, dict):        # {"陈": "陈默", …}
        roles = [[k, v] for k, v in roles.items()]
    for r in roles or []:
        if isinstance(r, (tuple, list)):
            a = str(r[0]).strip() if r else ""
            n = str(r[1]).strip() if len(r) > 1 else a
        elif isinstance(r, dict):
            a = str(r.get("abbr", r.get("name", ""))).strip()
            n = str(r.get("name", a)).strip()
        elif isinstance(r, str) and "=" in r:
            a, n = (p.strip() for p in r.split("=", 1))
        else:
            a = n = str(r).strip()
        if a:
            pairs.append((a, n or a))

    # roles 里没带全名的，回退到 meta.角色缩写 对照表
    if meta and isinstance(meta, dict):
        table = str(meta.get("角色缩写", "") or "")
        lookup = {}
        for seg in table.replace("／", "/").split("/"):
            seg = seg.strip()
            if "=" not in seg:
                continue
            a, rest = (p.strip() for p in seg.split("=", 1))
            # 「阿萤(方四/沈萤，女主)」→ 取括号前的「阿萤」
            n = re.split(r"[（(]", rest, maxsplit=1)[0].strip()
            lookup[a] = n or a
        pairs = [(a, n if n != a else lookup.get(a, a)) for a, n in pairs]
    return pairs


def expand_roles(headers, rows, roles):
    """每个角色独占一列，出场打 ✓（与分镜表 skill 同一套规范）。

    roles: 角色名列表（如 ["陈","林"]）
    v1.2 起角色列组**固定落在「页数」列之后**（…内容梗概·页数·角色列组…）。
    columns 里若写了 cast，它只作为占位符被吞掉，不再当位置锚点——
    否则用户把 cast 写在 pages 前面，列序就又退回去了。
    行内出场角色写在 row["roles"]（列表或 "陈/林" 字符串）；无则全空。
    """
    if not roles:
        return headers, rows
    ROLE_COLS.clear()
    ROLE_ABBR.clear()
    displays = []
    for item in roles:
        a, d = item if isinstance(item, (tuple, list)) else (item, item)
        ROLE_COLS.add(d)
        ROLE_ABBR[d] = a
        displays.append(d)

    # 1) 吞掉 cast 占位列（若有）
    if "出场人物" in headers:
        c = headers.index("出场人物")
        base, drop_at = headers[:c] + headers[c + 1:], c
    else:
        base, drop_at = list(headers), None

    # 2) 角色列组插在「页数」之后
    p_idx = next((i for i, h in enumerate(base) if h == "页数"), len(base) - 1)
    insert_at = min(p_idx + 1, len(base))
    new_headers = base[:insert_at] + displays + base[insert_at:]

    out = []
    for row in rows:
        raw = row.pop() if row and isinstance(row[-1], list) else []
        present = set(raw or [])
        row = list(row)
        if drop_at is not None:
            row = row[:drop_at] + row[drop_at + 1:]
        # 格子里填单字缩写（v1.3）；行内写缩写或全名都能对上
        vals = [ROLE_ABBR[d] if (ROLE_ABBR[d] in present or d in present) else ""
                for d in displays]
        out.append(row[:insert_at] + vals + row[insert_at:])
    return new_headers, out


def col_width(name):
    """角色列窄一些，但要放得下全名表头；其余按 WIDTHS。"""
    if name in ROLE_COLS:
        return max(4.6, len(name) * 1.9)
    return WIDTHS.get(name, 14)


def parse_page(v):
    """解析页数值：支持 '2 5/8' / '3/8' / '1.375' / 2，返回 float(页)。"""
    if isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("－", "-")
    if not s:
        return 0.0
    try:
        total = 0.0
        for p in s.split():
            if "/" in p:
                num, den = p.split("/", 1)
                total += float(num) / float(den)
            else:
                total += float(p)
        return total
    except (ValueError, ZeroDivisionError):
        return 0.0


def parse_num(v):
    if isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip())
    except ValueError:
        return 0.0


def is_num_col(name):
    """参与合计行汇总的列。"""
    return name == "页数"


def add_mainset_column(headers, rows):
    """主场景列，排在「场景」**之前**。

    - columns 里显式写了 `mainset`：按写的顺序渲染，值仍自动从场景列解析
    - 没写：在「场景」列之前自动插入
    """
    if "场景" not in headers:
        return headers, rows
    s_idx = headers.index("场景")  # 旧索引，等于 row 里场景值所在位置
    if "主场景" in headers:
        m_idx = headers.index("主场景")
        out = []
        for row in rows:
            row = list(row) + [""] * (len(headers) - len(row))
            row[m_idx] = big_set(row[s_idx])
            out.append(row)
        return headers, out
    # 自动插入：先按旧索引取场景值，再补齐长度，最后插入到场景之前
    m_idx = s_idx
    headers = headers[:m_idx] + ["主场景"] + headers[m_idx:]
    out = []
    for row in rows:
        row = list(row)
        val = big_set(row[s_idx]) if s_idx < len(row) else ""
        row = row + [""] * (len(headers) - len(row) - 1)  # 新表头长度 - 1
        out.append(row[:m_idx] + [val] + row[m_idx:])
    return headers, out


def drop_legacy_cols(headers, rows):
    """剔除 v1.2 起不再输出的两列：时长(秒)、累计页数。

    旧 JSON/CSV 里若还写着这两列，直接丢掉并提示，不要照原样渲染——
    否则会出现空列 / 与页数并排的重复口径，看着像没改干净。
    """
    keep = [i for i, h in enumerate(headers)
            if h not in set(DROPPED_FIELDS.values())]
    if len(keep) == len(headers):
        return headers, rows
    dropped = [h for h in headers if h in set(DROPPED_FIELDS.values())]
    print(f"[info] 已移除不再输出的列：{'、'.join(dropped)}", file=sys.stderr)
    return ([headers[i] for i in keep],
            [[row[i] for i in keep] for row in rows])


def compute_stats(headers, rows):
    """按场次统计。返回 dict。"""
    idx = {h: i for i, h in enumerate(headers)}

    def get(row, key, default=""):
        i = idx.get(key)
        return row[i] if i is not None and i < len(row) else default

    mains, subs, pages_total = set(), 0, 0.0
    daynight = OrderedDict()
    inout = OrderedDict()
    groups = OrderedDict()

    # 角色出场统计（演员档期用）
    role_idx = [(h, idx[h]) for h in headers if h in ROLE_COLS]
    role_stat = OrderedDict((name, {"场数": 0, "页数": 0.0}) for name, _ in role_idx)

    for row in rows:
        scene_val = str(get(row, "场号")).strip()
        main_val, suffix = split_scene(scene_val)
        # 多集项目各集场号都从 1 重新起编，统计场次必须带上「集」一起去重
        ep_val = str(get(row, "集")).strip()
        scene_key = (ep_val, main_val) if ep_val else main_val
        if scene_val:
            mains.add(scene_key)
            if suffix:
                subs += 1
        p = parse_page(get(row, "页数"))
        pages_total += p

        dn = str(get(row, "日/夜")).strip() or "未标"
        d = daynight.setdefault(dn, {"行数": 0, "页数": 0.0})
        d["行数"] += 1
        d["页数"] += p

        io_ = str(get(row, "内/外")).strip() or "未标"
        o = inout.setdefault(io_, {"行数": 0, "页数": 0.0})
        o["行数"] += 1
        o["页数"] += p

        # 主场景列优先；没有则回退从场景列解析
        g = str(get(row, "主场景")).strip() or big_set(get(row, "场景"))
        g = g or "未标"
        item = groups.setdefault(g, {"场次": set(), "行数": 0, "页数": 0.0})
        item["场次"].add(scene_key)
        item["行数"] += 1
        item["页数"] += p

        for name, i in role_idx:
            mark = ROLE_ABBR.get(name, CHECK)
            if i < len(row) and str(row[i]).strip() in (mark, CHECK):
                st = role_stat[name]
                st["场数"] += 1
                st["页数"] += p

    return {
        "总场数": len(mains),
        "子场数": subs,
        "总行数": len(rows),
        "总页数": pages_total,
        "日夜": daynight,
        "内外": inout,
        "主场景": groups,
        "角色": role_stat,
    }


def fmt_est_duration(pages):
    """按 1 页 ≈ 1 分钟的经验值，把总页数换算成预估时长。"""
    sec = int(round(pages * 60))
    if sec <= 0:
        return "—"
    h, m = sec // 3600, (sec % 3600) // 60
    return f"{h} 小时 {m} 分" if h else f"{m} 分"


# ---------------------------------------------------------------------------
def build_xlsx(title, meta_text, headers, rows, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "顺场表"

    ncols = len(headers)
    last_col = get_column_letter(ncols)

    # 标题行
    ws.merge_cells(f"A1:{last_col}1")
    tcell = ws["A1"]
    tcell.value = title + (f"\n{meta_text}" if meta_text else "")
    tcell.font = Font(name="微软雅黑", size=16, bold=True, color="1F3864")
    tcell.fill = PatternFill("solid", fgColor=TITLE_FILL)
    tcell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    tcell.border = BORDER
    ws.row_dimensions[1].height = 34 if not meta_text else 46

    # 表头行
    for j, name in enumerate(headers, start=1):
        c = ws.cell(row=2, column=j, value=name)
        c.font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[2].height = 28.5

    num_idxs = [i for i, h in enumerate(headers) if is_num_col(h)]
    LEFT_COLS = ("内容梗概", "场次名", "道具", "备注", "出场人物", "特约", "群众", "视效", "场景", "主场景")
    totals = {i: 0.0 for i in num_idxs}

    for i, row in enumerate(rows):
        r = i + 3
        zebra = (i % 2 == 1)
        for j in range(ncols):
            val = row[j] if j < len(row) else ""
            name = headers[j]
            c = ws.cell(row=r, column=j + 1, value=val)
            c.border = BORDER
            if j in totals:
                f = parse_page(val) if name == "页数" else parse_num(val)
                if str(val).strip():
                    c.value = round(f, 3)
                totals[j] += f if str(val).strip() else 0.0
            if name in LEFT_COLS:
                c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            if zebra:
                c.fill = PatternFill("solid", fgColor=ZEBRA)
            # 配色
            if name == "氛围":
                c.fill = PatternFill("solid", fgColor=MOOD_FILL)
                c.font = Font(name="微软雅黑", size=10, bold=True)
            elif name == "主场景":
                c.fill = PatternFill("solid", fgColor=MAINSET_FILL)
                c.font = Font(name="微软雅黑", size=10, bold=True, color="1F3864")
            elif name == "日/夜":
                fill = DAY_FILLS.get(str(val).strip())
                if fill:
                    c.fill = PatternFill("solid", fgColor=fill)
                    c.font = Font(name="微软雅黑", size=10, bold=True)
                else:
                    c.font = Font(name="微软雅黑", size=10)
            elif name == "内/外":
                fill = IO_FILLS.get(str(val).strip())
                if fill:
                    c.fill = PatternFill("solid", fgColor=fill)
                    c.font = Font(name="微软雅黑", size=10, bold=True)
                else:
                    c.font = Font(name="微软雅黑", size=10)
            elif name in ROLE_COLS:
                # 出场标记：v1.3 起填角色单字缩写（旧的 ✓ 也认）
                if str(val).strip() in (ROLE_ABBR.get(name), CHECK):
                    c.font = Font(name="微软雅黑", size=11, bold=True, color="1F3864")
                else:
                    c.font = Font(name="微软雅黑", size=10)
            elif name in ("场号", "集"):
                c.font = Font(name="微软雅黑", size=10, bold=True, color="1F3864")
            else:
                c.font = Font(name="微软雅黑", size=10)
        ws.row_dimensions[r].height = 34

    # 合计行
    stats = compute_stats(headers, rows)
    if totals:
        r = len(rows) + 3
        for j in range(ncols):
            c = ws.cell(row=r, column=j + 1, value="")
            c.border = BORDER
            c.fill = PatternFill("solid", fgColor=TOTAL_FILL)
            c.alignment = Alignment(horizontal="center", vertical="center")
        label = ws.cell(row=r, column=1, value=f"合计 · {stats['总场数']} 场")
        label.font = Font(name="微软雅黑", size=10, bold=True)
        label.alignment = Alignment(horizontal="center", vertical="center")
        for j in num_idxs:
            dcell = ws.cell(row=r, column=j + 1, value=round(totals[j], 3))
            dcell.font = Font(name="微软雅黑", size=10, bold=True, color="C00000")
            dcell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[r].height = 24

    for j, name in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(j)].width = col_width(name)

    if rows:
        ws.auto_filter.ref = f"A2:{last_col}{len(rows) + 2}"
    ws.freeze_panes = "A3"
    ws.print_title_rows = "1:2"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_view.showGridLines = False

    build_stats_sheet(wb, stats)
    wb.save(out_path)
    return out_path, stats


def build_stats_sheet(wb, stats):
    """第二个 sheet：按场次统计。"""
    ws = wb.create_sheet("场景统计")
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDE", (16, 10, 10, 12, 12)):
        ws.column_dimensions[col].width = w

    r = 1

    def section(title):
        nonlocal r
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        c = ws.cell(row=r, column=1, value=title)
        c.font = Font(name="微软雅黑", size=12, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[r].height = 24
        r += 1

    def head(*names):
        nonlocal r
        for j, n in enumerate(names, start=1):
            c = ws.cell(row=r, column=j, value=n)
            c.font = Font(name="微软雅黑", size=10, bold=True)
            c.fill = PatternFill("solid", fgColor=TITLE_FILL)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = BORDER
        r += 1

    def line(*vals, bold=False):
        nonlocal r
        for j, v in enumerate(vals, start=1):
            c = ws.cell(row=r, column=j, value=v)
            c.font = Font(name="微软雅黑", size=10, bold=bold)
            c.alignment = Alignment(horizontal="left" if j == 1 else "center",
                                    vertical="center")
            c.border = BORDER
        r += 1

    # 总览
    tot_rows = stats["总行数"] or 1
    tot_pages = stats["总页数"] or 1.0
    pct = lambda n: f"{n / tot_pages * 100:.1f}%"

    section("一、总览")
    head("指标", "数值", "", "指标", "数值")
    line("总场数", stats["总场数"], "", "总页数", round(stats["总页数"], 3), bold=True)
    line("数据行数", stats["总行数"], "",
         "预估总时长", fmt_est_duration(stats["总页数"]), bold=True)
    line("平均单场页数",
         round(stats["总页数"] / stats["总场数"], 3) if stats["总场数"] else 0,
         "", "时长口径", "1 页 ≈ 1 分钟（估算）")
    r += 1

    # 日/夜
    section("二、日 / 夜 分布（夜戏占比 = 周期与预算关键）")
    head("时间档", "场数", "场数占比", "页数", "页数占比")
    for k, v in stats["日夜"].items():
        line(k, v["行数"], f"{v['行数'] / tot_rows * 100:.1f}%",
             round(v["页数"], 3), pct(v["页数"]))
    r += 1

    # 内/外
    section("三、内 / 外 分布（外景占比 = 天气风险）")
    head("空间", "场数", "场数占比", "页数", "页数占比")
    for k, v in stats["内外"].items():
        line(k, v["行数"], f"{v['行数'] / tot_rows * 100:.1f}%",
             round(v["页数"], 3), pct(v["页数"]))
    r += 1

    # 主场景归组
    section("四、按主场景归组（页数降序 · 统筹排期用）")
    head("主场景", "场数", "场次", "页数", "页数占比")
    groups = sorted(stats["主场景"].items(), key=lambda kv: -kv[1]["页数"])
    for k, v in groups:
        line(k, v["行数"], len(v["场次"]),
             round(v["页数"], 3), pct(v["页数"]), bold=(v["页数"] >= 1.0))
    r += 1

    # 角色出场
    if stats["角色"]:
        section("五、角色出场（场数降序 · 演员档期用）")
        head("角色", "场数", "场数占比", "页数", "页数占比")
        for k, v in sorted(stats["角色"].items(), key=lambda kv: -kv[1]["场数"]):
            line(k, v["场数"], f"{v['场数'] / tot_rows * 100:.1f}%",
                 round(v["页数"], 3), pct(v["页数"]))
    return ws


def build_html(title, meta_text, headers, rows, out_path):
    head = "".join(f"<th>{h}</th>" for h in headers)
    body = ""
    for i, row in enumerate(rows):
        cls = ' class="zebra"' if i % 2 == 1 else ""
        tds = "".join(f"<td>{str(v).replace(chr(10), '<br>') if v != '' else ''}</td>" for v in row)
        body += f"<tr{cls}>{tds}</tr>\n"
    html = f"""<!DOCTYPE html><html lang="zh"><head><meta charset="utf-8">
<title>{title}</title>
<style>
 * {{ font-family: "Microsoft YaHei", sans-serif; }}
 body {{ margin: 24px; color:#222; }}
 h1 {{ text-align:center; color:#1F3864; font-size:20px; margin:0 0 4px; }}
 .meta {{ text-align:center; color:#666; font-size:12px; margin-bottom:14px; }}
 table {{ border-collapse:collapse; width:100%; font-size:13px; }}
 th {{ background:#1F3864; color:#fff; padding:8px 6px; border:1px solid #BFBFBF; }}
 td {{ border:1px solid #BFBFBF; padding:6px; vertical-align:top; }}
 tr.zebra td {{ background:#F7F9FC; }}
</style></head><body>
<h1>{title}</h1>
<div class="meta">{meta_text}</div>
<table><thead><tr>{head}</tr></thead><tbody>
{body}</tbody></table>
</body></html>"""
    Path(out_path).write_text(html, encoding="utf-8")
    return out_path


def main():
    ap = argparse.ArgumentParser(description="生成 Coco影视顺场表 Excel（一场一行，含场景统计）")
    ap.add_argument("input", help="JSON/CSV 路径，或 - 表示 stdin")
    ap.add_argument("-o", "--output", help="输出 .xlsx 路径（默认与输入同名）")
    ap.add_argument("--html", action="store_true", help="同时导出可打印 HTML（仅主表）")
    args = ap.parse_args()

    title, meta_text, headers, rows, roles = load_data(args.input)
    # 先剔除旧列（时长/累计页数），再做列变换，避免索引算错
    headers, rows = drop_legacy_cols(headers, rows)
    if roles:  # 每个角色独占一列（✓ 出场标记），落在「页数」列之后
        headers, rows = expand_roles(headers, rows, roles)
    headers, rows = add_mainset_column(headers, rows)

    out = args.output or (("顺场表" if args.input == "-" else Path(args.input).stem) + ".xlsx")
    if not out.lower().endswith(".xlsx"):
        out += ".xlsx"

    _, stats = build_xlsx(title, meta_text, headers, rows, out)
    print(f"[ok] 已生成 Excel：{out}")
    print(f"     {stats['总场数']} 场 / {stats['总页数']:.3g} 页 / "
          f"预估 {fmt_est_duration(stats['总页数'])} / 主场景 {len(stats['主场景'])} 组"
          + (f" / 角色 {len(stats['角色'])} 人" if stats["角色"] else ""))
    if stats["子场数"]:
        print(f"     [提示] 检测到 {stats['子场数']} 个带字母后缀的场号（1A 之类），"
              f"本表已按纯数字场次口径统计，建议改成顺序数字。")

    if args.html:
        hpath = Path(out).with_suffix(".html")
        build_html(title, meta_text, headers, rows, str(hpath))
        print(f"[ok] 已生成 HTML：{hpath}")


if __name__ == "__main__":
    main()
