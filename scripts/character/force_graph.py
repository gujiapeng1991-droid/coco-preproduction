#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
force_graph.py —— 把「人物 + 关系」JSON 渲染成一张可交互的力导向关系网 HTML。

零外部依赖（只用标准库），产出的 HTML 也是零依赖单文件，双击即可打开。

用法:
    python3 force_graph.py characters.json -o 人物关系网.html
    python3 force_graph.py characters.json -t "《XX》人物关系网" -s "副标题" -o out.html

JSON 结构见 examples/character_sample.json，最小必需字段:
    {
      "factions": [{"id":"hou","name":"侯府"}],
      "nodes":    [{"id":"a","name":"张三","sub":"身份","ring":0,"fac":"hou","desc":"小传"}],
      "edges":    [{"a":"a","b":"b","t":"blood","label":"父子","note":"备注"}]
    }
可选:
    "types": 自定义关系类型（颜色/线型/弹簧参数），不写则用默认 5 类
    "title"/"subtitle"/"notes"/"palette": 覆盖默认文案与阵营配色
"""
import argparse, json, re, sys, os
from pathlib import Path

# ---------------------------------------------------------------- 默认关系类型
# L=弹簧自然长度(px)  k=弹性系数   dash=虚线  strong=加粗(主线)  soft=半透明
DEFAULT_TYPES = [
    {"id": "blood", "name": "血亲 / 姻亲",   "L": 118, "k": .055, "dash": False, "strong": False, "soft": False},
    {"id": "love",  "name": "情感 / 爱慕",   "L": 104, "k": .075, "dash": False, "strong": True,  "soft": False},
    {"id": "ally",  "name": "同盟 / 守护",   "L": 126, "k": .050, "dash": False, "strong": False, "soft": False},
    {"id": "foe",   "name": "仇杀 / 对抗",   "L": 158, "k": .038, "dash": True,  "strong": False, "soft": False},
    {"id": "ctrl",  "name": "控制 / 胁迫",   "L": 172, "k": .026, "dash": True,  "strong": False, "soft": True},
]
# 关系类型配色 [深色, 浅色]
DEFAULT_TYPE_COLOR = {
    "blood": ["#C89B6A", "#8B5E3C"],
    "love":  ["#FF6B8A", "#D6455D"],
    "ally":  ["#4FD1A5", "#2E9E6B"],
    "foe":   ["#FF5A5A", "#C0392B"],
    "ctrl":  ["#7C8698", "#9aa0ab"],
}
# 阵营调色板：8 组，每组 [深{f,s,t}, 浅{f,s,t}]，按 factions 顺序自动分配
DEFAULT_PALETTE = [
    [["#3a2e17", "#D9A441", "#F2DCA6"], ["#FDF3DC", "#D9A441", "#7A5A16"]],  # 金
    [["#14263c", "#4A8FD4", "#A9CEF2"], ["#E3EEFA", "#4A8FD4", "#1E4E80"]],  # 蓝
    [["#3a2417", "#DD8B5C", "#F4C6A5"], ["#FCE9DC", "#DD8B5C", "#8A4212"]],  # 橙
    [["#241c3c", "#9B7FE0", "#CDBBF7"], ["#EDE7FB", "#8B6FD1", "#4C2E8A"]],  # 紫
    [["#3f1418", "#FF5C5C", "#FFB7B7"], ["#FDE8E8", "#D95C5C", "#8E2020"]],  # 红
    [["#0f332e", "#2FB3A6", "#8FE3D8"], ["#DFF5F2", "#2FB3A6", "#0B6B60"]],  # 青
    [["#15301c", "#5CB85C", "#B6EBB6"], ["#E4F5E4", "#4C9A4C", "#1F5C1F"]],  # 绿
    [["#3a1730", "#E06BA8", "#F7BEDC"], ["#FDE8F2", "#D4559B", "#8E2058"]],  # 粉
]
RS = [34, 27, 22, 18]   # ring -> 半径
FS = [15, 13, 12, 11]   # ring -> 字号


def js(obj):
    """Python 对象 -> JS 字面量（ensure_ascii=False 保住中文）"""
    return json.dumps(obj, ensure_ascii=False)


def build_palette(factions, override):
    pal = {}
    for i, f in enumerate(factions):
        fid = f["id"]
        if override and fid in override:
            o = override[fid]
            pal[fid] = {"d": o["dark"], "l": o["light"]}
        else:
            d, l = DEFAULT_PALETTE[i % len(DEFAULT_PALETTE)]
            pal[fid] = {"d": {"f": d[0], "s": d[1], "t": d[2]},
                        "l": {"f": l[0], "s": l[1], "t": l[2]}}
    return pal


def check(data, types, factions):
    """自检：把会让网页白屏或画错的地方提前拦下来"""
    err, warn = [], []
    ids = [n["id"] for n in data["nodes"]]
    if len(ids) != len(set(ids)):
        dup = [i for i in set(ids) if ids.count(i) > 1]
        err.append(f"人物 id 重复: {dup}")
    fid_set = {f["id"] for f in factions}
    tid_set = {t["id"] for t in types}
    id_set = set(ids)

    for n in data["nodes"]:
        if not n.get("name"):
            err.append(f"人物 {n.get('id')} 缺 name")
        if n.get("fac") not in fid_set:
            err.append(f"人物 {n.get('name')} 的阵营 '{n.get('fac')}' 不在 factions 里")

    linked = set()
    for i, e in enumerate(data["edges"]):
        if e.get("a") not in id_set or e.get("b") not in id_set:
            err.append(f"第 {i+1} 条关系指向不存在的人物: {e.get('a')}->{e.get('b')}")
        if e.get("a") == e.get("b"):
            err.append(f"第 {i+1} 条关系是自环: {e.get('a')}")
        if e.get("t") not in tid_set:
            err.append(f"第 {i+1} 条关系的类型 '{e.get('t')}' 不在 types 里")
        linked.add(e.get("a")); linked.add(e.get("b"))

    for i in id_set - linked:
        nm = next((n["name"] for n in data["nodes"] if n["id"] == i), i)
        warn.append(f"孤立节点（没有任何关系）: {nm}")

    if len(factions) > len(DEFAULT_PALETTE) and not data.get("palette"):
        warn.append(f"阵营数 {len(factions)} 超过内置调色板 {len(DEFAULT_PALETTE)} 组，颜色会开始循环复用")
    return err, warn


def render(data, out_path, title=None, subtitle=None, tpl=None):
    types = data.get("types") or DEFAULT_TYPES
    # 用户自定义类型时补齐缺省的弹簧/线型参数
    base = {t["id"]: t for t in DEFAULT_TYPES}
    for t in types:
        b = base.get(t["id"], {})
        t.setdefault("L", b.get("L", 130))
        t.setdefault("k", b.get("k", .045))
        t.setdefault("dash", b.get("dash", False))
        t.setdefault("strong", b.get("strong", False))
        t.setdefault("soft", b.get("soft", False))
    factions = data["factions"]

    err, warn = check(data, types, factions)
    for w in warn:
        print(f"  [警告] {w}", file=sys.stderr)
    if err:
        for e in err:
            print(f"  [错误] {e}", file=sys.stderr)
        sys.exit(1)

    # 补默认半径/字号
    for n in data["nodes"]:
        r = int(n.get("ring", 2))
        n.setdefault("r", RS[min(r, 3)])
        n.setdefault("fs", FS[min(r, 3)])
        n.setdefault("sub", "")
        n.setdefault("desc", "")

    tcolor = dict(DEFAULT_TYPE_COLOR)
    for k, v in (data.get("type_colors") or {}).items():
        tcolor[k] = v
    pal = build_palette(factions, data.get("palette"))

    tpl_path = Path(tpl) if tpl else Path(__file__).with_name("template.html")
    html = tpl_path.read_text(encoding="utf-8")

    notes = data.get("notes") or [
        "<b>拖拽</b>节点可重新摆位并自动钉住（出现 📌），双击节点解除。",
        "<b>双击</b>节点进入隔离模式，只看他的一阶邻域；双击空白处退出。",
        "关掉某类关系，剩下的网会当场重新排布。",
    ]

    rep = {
        "__TITLE__":     title or data.get("title") or "人物关系网",
        "__SUBTITLE__":  subtitle or data.get("subtitle") or "力导向布局 —— 距离越近，关系越紧。",
        "__NOTES__":     "<br>".join(notes),
        "__FACTIONS_JS__": js(factions),
        "__TYPES_JS__":    js(types),
        "__NODES_JS__":    js(data["nodes"]),
        "__EDGES_JS__":    js(data["edges"]),
        "__PAL_JS__":      js(pal),
        "__TYPECOLOR_JS__": js(tcolor),
    }
    for k, v in rep.items():
        html = html.replace(k, v)

    left = re.findall(r"__[A-Z_]+__", html)
    if left:
        print(f"  [错误] 模板仍有未替换占位符: {set(left)}", file=sys.stderr)
        sys.exit(1)

    Path(out_path).write_text(html, encoding="utf-8")
    print(f"  已生成 {out_path}")
    print(f"  {len(data['nodes'])} 人物 / {len(data['edges'])} 关系 / {len(factions)} 阵营 / {len(types)} 关系类型")
    return out_path


def main():
    ap = argparse.ArgumentParser(description="把人物关系 JSON 渲染成力导向关系网 HTML")
    ap.add_argument("json", help="人物数据 JSON 文件")
    ap.add_argument("-o", "--out", help="输出 HTML 路径（默认与 JSON 同名）")
    ap.add_argument("-t", "--title", help="覆盖标题")
    ap.add_argument("-s", "--subtitle", help="覆盖副标题")
    ap.add_argument("--tpl", help="自定义模板路径")
    a = ap.parse_args()

    data = json.loads(Path(a.json).read_text(encoding="utf-8"))
    out = a.out or str(Path(a.json).with_suffix(".html"))
    render(data, out, a.title, a.subtitle, a.tpl)


if __name__ == "__main__":
    main()
