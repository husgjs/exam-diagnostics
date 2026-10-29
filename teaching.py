#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""教学产出工具（纯标准库，0 次大模型调用）

子命令
  tree    知识点树 + 前置依赖（先覆盖「有理数」一章）+ 前置链查询
  layer   分层作业包（基础/中等/拔高，只从已有题里挑，不生成新题）
  review  讲评清单（按错误率降序取前 N，分配讲评时长，合计 <= 40 分钟）
  print   打印中心（错题卷 / 复习单 / 口算卡，纯 HTML+CSS，浏览器 Ctrl+P）
  import  答题卡导入整合（读卡机导出表 -> 讲评清单 + 分层作业）
  all     依次跑 tree/layer/review/print

用法示例
  python teaching.py tree   --query 一元一次方程移项
  python teaching.py layer
  python teaching.py review
  python teaching.py print
  python teaching.py import --csv "..\_tmp\card_export.csv" --paper "..\_tmp\paper_structure.json"

数据默认位置：本文件上一级目录的 _tmp\
所有输出只写到 教师工具\teaching_out\，不碰任何别人的文件，不往 C 盘写。
"""
import argparse
import csv
import html
import json
import os
import random
import re
import sqlite3
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(SCRIPT_DIR)
TMP_DIR = os.path.join(ROOT, "_tmp")
OUT_DIR = os.path.join(SCRIPT_DIR, "teaching_out")

ORDER_SRC = "待核对人教版目录"
RE_QNO = re.compile(r"第\s*(\d+)\s*题")


def out_dir():
    if not os.path.isdir(OUT_DIR):
        os.makedirs(OUT_DIR)
    return OUT_DIR


def jdump(obj, path):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)
    return path


# ============================================================ 知识点树
# 说明：前置依赖按「数学依赖关系」给出，这部分可确认；
#       章序号与小节编号因教材目录未拿到，统一标 待核对人教版目录。
CHAPTERS = [
    {"id": "rat", "name": "有理数", "seq": None, "source": ORDER_SRC,
     "coverage": "已建模", "note": "本章知识点已建模；章序号/小节编号未与教材核对"},
    {"id": "expr", "name": "整式的加减", "seq": None, "source": ORDER_SRC,
     "coverage": "未覆盖", "note": "仅登记章名，未建知识点"},
    {"id": "eq", "name": "一元一次方程", "seq": None, "source": ORDER_SRC,
     "coverage": "接口节点", "note": "仅建「移项」前置链所需的接口节点，未覆盖本章"},
]

NODES = [
    # ---------------- 有理数（本章完整建模） ----------------
    {"id": "rat-concept", "name": "有理数的概念与分类", "chapter": "rat",
     "prereq": [], "seq": None, "seq_source": ORDER_SRC,
     "errs": ["概念不清", "分类时漏掉 0 或负数", "把有限小数排除在分数之外"],
     "aliases": ["负数概念", "有理数概念", "有理数分类"]},
    {"id": "rat-numberline", "name": "数轴", "chapter": "rat",
     "prereq": ["rat-concept"], "seq": None, "seq_source": ORDER_SRC,
     "errs": ["概念不清", "原点/正方向/单位长度三要素缺项", "描点时方向看反"],
     "aliases": ["数轴", "数轴与符号"]},
    {"id": "rat-opposite", "name": "相反数", "chapter": "rat",
     "prereq": ["rat-concept"], "seq": None, "seq_source": ORDER_SRC,
     "errs": ["概念不清", "与倒数定义混淆", "把 -a 直接当成负数"],
     "aliases": ["相反数"]},
    {"id": "rat-reciprocal", "name": "倒数", "chapter": "rat",
     "prereq": ["rat-concept"], "seq": None, "seq_source": ORDER_SRC,
     "errs": ["概念不清", "与相反数混淆", "忘记 0 没有倒数"],
     "aliases": ["倒数"]},
    {"id": "rat-abs", "name": "绝对值", "chapter": "rat",
     "prereq": ["rat-numberline", "rat-opposite"], "seq": None, "seq_source": ORDER_SRC,
     "errs": ["概念不清", "去绝对值时漏讨论符号", "把 |a| 写成 a"],
     "aliases": ["绝对值", "绝对值化简"]},
    {"id": "rat-compare", "name": "有理数的大小比较", "chapter": "rat",
     "prereq": ["rat-numberline", "rat-abs"], "seq": None, "seq_source": ORDER_SRC,
     "errs": ["审题错误", "两个负数比大小把结论写反"],
     "aliases": ["有理数比较", "有理数大小比较"]},
    {"id": "rat-add", "name": "有理数加法", "chapter": "rat",
     "prereq": ["rat-abs", "rat-numberline"], "seq": None, "seq_source": ORDER_SRC,
     "errs": ["计算错误", "异号相加时符号判错", "绝对值相减顺序反了"],
     "aliases": ["有理数加法"]},
    {"id": "rat-sub", "name": "有理数减法", "chapter": "rat",
     "prereq": ["rat-add", "rat-opposite"], "seq": None, "seq_source": ORDER_SRC,
     "errs": ["计算错误", "减号没有转成加相反数", "减去负数直接写成了减"],
     "aliases": ["有理数减法"]},
    {"id": "rat-muldiv", "name": "有理数乘除", "chapter": "rat",
     "prereq": ["rat-abs", "rat-reciprocal"], "seq": None, "seq_source": ORDER_SRC,
     "errs": ["计算错误", "负负得正/除法符号法则用错", "除以一个数没换成乘倒数"],
     "aliases": ["有理数乘除"]},
    {"id": "rat-power", "name": "有理数乘方", "chapter": "rat",
     "prereq": ["rat-muldiv"], "seq": None, "seq_source": ORDER_SRC,
     "errs": ["计算错误", "底数是负数时括号漏写", "-2^2 与 (-2)^2 不分"],
     "aliases": ["乘方", "有理数乘方"]},
    {"id": "rat-mixed", "name": "有理数混合运算", "chapter": "rat",
     "prereq": ["rat-add", "rat-sub", "rat-muldiv", "rat-power"], "seq": None, "seq_source": ORDER_SRC,
     "errs": ["步骤遗漏", "运算顺序错", "跳步导致符号丢失"],
     "aliases": ["有理数混合运算", "运算顺序", "有理数运算", "新定义运算", "裂项求和"]},
    {"id": "rat-sci", "name": "科学记数法与近似数", "chapter": "rat",
     "prereq": ["rat-power"], "seq": None, "seq_source": ORDER_SRC,
     "errs": ["概念不清", "有效数字位数数错", "10 的指数与小数点位数对不上"],
     "aliases": ["科学记数法", "近似数"]},
    {"id": "rat-app", "name": "有理数的实际应用", "chapter": "rat",
     "prereq": ["rat-add", "rat-sub"], "seq": None, "seq_source": ORDER_SRC,
     "errs": ["审题错误", "把「增加/减少」的方向搞反", "答案漏单位"],
     "aliases": ["有理数应用", "正负数应用", "数轴应用"]},
    # ---------------- 一元一次方程（仅接口节点） ----------------
    {"id": "eq-property", "name": "等式性质", "chapter": "eq",
     "prereq": ["rat-sub"], "seq": None, "seq_source": ORDER_SRC,
     "errs": [], "stub": True, "aliases": ["等式性质"]},
    {"id": "eq-transpose", "name": "一元一次方程移项", "chapter": "eq",
     "prereq": ["eq-property"], "seq": None, "seq_source": ORDER_SRC,
     "errs": [], "stub": True, "aliases": ["移项", "一元一次方程移项"]},
    {"id": "eq-solve", "name": "一元一次方程的解法", "chapter": "eq",
     "prereq": ["eq-transpose"], "seq": None, "seq_source": ORDER_SRC,
     "errs": [], "stub": True, "aliases": ["一元一次方程"]},
]

NODE_BY_ID = {n["id"]: n for n in NODES}
CHAPTER_BY_ID = {c["id"]: c for c in CHAPTERS}


def _build_alias():
    idx = {}
    for n in NODES:
        idx[n["name"]] = n["id"]
        idx[n["id"]] = n["id"]
        for a in n.get("aliases", []):
            idx[a] = n["id"]
    return idx


ALIAS = _build_alias()


def find_node(query):
    """按 id / 名称 / 别名 找知识点；找不到返回 None。"""
    q = (query or "").strip()
    if q in ALIAS:
        return NODE_BY_ID[ALIAS[q]]
    for name, nid in ALIAS.items():          # 退一步做包含匹配
        if q and (q in name or name in q):
            return NODE_BY_ID[nid]
    return None


def node_label(nid):
    n = NODE_BY_ID[nid]
    return "%s（%s）" % (n["name"], nid)


def main_chain(nid, seen=()):
    """最长前置链，由近及远（元素含 nid 自身）。"""
    seen = seen + (nid,)
    best = []
    for p in NODE_BY_ID[nid]["prereq"]:
        if p in NODE_BY_ID and p not in seen:
            c = main_chain(p, seen)
            if len(c) > len(best):
                best = c
    return [nid] + best


def tree_lines(nid, prefix="", last=True, trail=()):
    n = NODE_BY_ID[nid]
    branch = "" if prefix == "" and trail == () else ("└─ " if last else "├─ ")
    head = prefix + branch
    lines = []
    if prefix == "" and trail == ():
        lines.append(node_label(nid))
    else:
        lines.append(head + node_label(nid))
    kids = [p for p in n["prereq"] if p in NODE_BY_ID and p not in trail + (nid,)]
    for i, p in enumerate(kids):
        is_last = (i == len(kids) - 1)
        if prefix == "" and trail == ():
            sub_prefix = ""
        else:
            sub_prefix = prefix + ("   " if last else "│  ")
        lines.extend(tree_lines(p, sub_prefix, is_last, trail + (nid,)))
    return lines


def cmd_tree(args):
    out_dir()
    doc = {
        "meta": {
            "name": "知识点树（先覆盖：有理数）",
            "version": 1,
            "order_verified": False,
            "order_source": ORDER_SRC,
            "prereq_basis": "前置依赖按数学依赖关系给出（可确认），与教材章节顺序无关",
            "coverage": {c["id"]: c["coverage"] for c in CHAPTERS},
            "note": "凡 order_verified=false 的章节，其「第几章/第几节」一律未填，需拿到人教版目录后核对",
        },
        "chapters": CHAPTERS,
        "nodes": NODES,
    }
    path = jdump(doc, os.path.join(out_dir(), "知识点树.json"))

    print("=" * 78)
    print("知识点树  |  章节 %d 个，知识点 %d 个" % (len(CHAPTERS), len(NODES)))
    print("=" * 78)
    for c in CHAPTERS:
        ns = [n for n in NODES if n["chapter"] == c["id"]]
        if c["id"] == "rat":
            print("[%s] %s  知识点 %d 个  (%s)" % (c["id"], c["name"], len(ns), c["coverage"]))
            for n in ns:
                pre = "、".join(NODE_BY_ID[p]["name"] for p in n["prereq"]) or "（本章起点，无前置）"
                print("   - %-16s %-22s 前置: %s" % (n["id"], n["name"], pre))
        else:
            print("[%s] %s  知识点 %d 个  (%s：%s)" % (c["id"], c["name"], len(ns), c["coverage"], c["note"]))
    print()
    print("『待核对人教版目录』标在：chapters[].source（共 %d 处）、nodes[].seq_source（共 %d 处）；"
          % (len(CHAPTERS), len([n for n in NODES if n.get("seq_source")])))
    print("nodes[].seq 一律为 null —— 章序号与小节编号未编，拿到人教版目录后再填。")
    print("树文件: %s (%d 字节)" % (path, os.path.getsize(path)))

    if args.query:
        n = find_node(args.query)
        print()
        print("-" * 78)
        if not n:
            print("查询失败：树里没有「%s」这个知识点（也未匹配到别名）。" % args.query)
            print("可用 id: %s" % "、".join(x["id"] for x in NODES))
            return 2
        nid = n["id"]
        ch = CHAPTER_BY_ID[n["chapter"]]
        print("知识点: %s" % n["name"])
        print("  id      : %s" % nid)
        print("  章节    : %s (%s)  覆盖情况: %s%s" % (
            ch["name"], ch["id"], ch["coverage"], "  ← 接口节点，本章未建模" if n.get("stub") else ""))
        print("  前置    : %s" % ("、".join(node_label(p) for p in n["prereq"]) or "（无）"))
        print("  常见错因: %s" % ("、".join(n["errs"]) or "（接口节点，未标注）"))
        print("  章节顺序: %s（seq=null，未编）" % n["seq_source"])
        chain = main_chain(nid)
        print()
        print("前置链（主链，由近及远）:")
        print("  " + " → ".join(NODE_BY_ID[x]["name"] for x in chain))
        print()
        print("前置树（含全部分支，不做去重）:")
        for line in tree_lines(nid):
            print("  " + line)
        print()
        print("注: 主链按「最长依赖路径」选出；括号里是知识点 id。箭头方向 = 依赖方向，不是课时先后。")
    return 0


# ============================================================ 数据加载
def _half_up(x):
    return int(x + 0.5)


def load_rows_paper(path):
    """class_rows.json：paper=[题号,知识点,满分,错因,预设错误率]"""
    doc = json.load(open(path, encoding="utf-8"))
    paper, n_students = {}, len(doc.get("rows", []))
    for row in doc.get("paper", []):
        no, kp, full, err, rate = row
        paper[str(no)] = {"kp": kp, "full": float(full), "err": err, "preset_rate": float(rate)}
    return paper, n_students


def load_paper_struct(path):
    """paper_structure.json：{题号: {kp, full, err, answer}}"""
    doc = json.load(open(path, encoding="utf-8"))
    out = {}
    for no, v in doc.items():
        out[str(no)] = {"kp": v.get("kp", "未标注"), "full": float(v.get("full", 1)),
                        "err": v.get("err", ""), "answer": v.get("answer")}
    return out


def load_questions(src="db", db_path=None, rows_path=None):
    """返回 (questions, meta)。questions 每项: no/kp/full/err/n/n_right/rate/P/node/node_name"""
    db_path = db_path or os.path.join(TMP_DIR, "class_demo.db")
    rows_path = rows_path or os.path.join(TMP_DIR, "class_rows.json")
    paper = {}
    if rows_path and os.path.exists(rows_path):
        paper, _ = load_rows_paper(rows_path)

    qs = []
    if src == "db":
        con = sqlite3.connect(db_path)
        acc = {}
        for ask, fs, gs in con.execute("select ask, full_score, got_score from questions"):
            m = RE_QNO.search(ask or "")
            no = m.group(1) if m else "?"
            a = acc.setdefault(no, {"n": 0, "got": 0.0, "tot": 0.0, "right": 0, "full": fs})
            g = float(gs) if gs is not None else 0.0
            a["n"] += 1
            a["got"] += g
            a["tot"] += float(fs or 1)
            if g >= float(fs or 1):
                a["right"] += 1
        con.close()
        for no, a in acc.items():
            info = paper.get(no, {})
            P = (a["got"] / a["tot"]) if a["tot"] else 0.0
            qs.append({"no": no, "kp": info.get("kp", "未标注"), "full": a["full"],
                       "err": info.get("err", ""), "n": a["n"], "n_right": a["right"],
                       "P": P, "rate": 1.0 - P})
        meta = {"src": "db", "path": db_path, "label": "class_demo.db 逐题得分",
                "n_students": max(a["n"] for a in acc.values()) if acc else 0}
    else:
        n_students = 0
        for no, info in paper.items():
            P = 1.0 - info["preset_rate"]
            qs.append({"no": no, "kp": info["kp"], "full": info["full"], "err": info["err"],
                       "n": 0, "n_right": 0, "P": P, "rate": info["preset_rate"]})
        meta = {"src": "rows", "path": rows_path,
                "label": "class_rows.json 预设错误率（无逐题明细）", "n_students": n_students}
        if rows_path and os.path.exists(rows_path):
            meta["n_students"] = len(json.load(open(rows_path, encoding="utf-8")).get("rows", []))

    for q in qs:
        nid = ALIAS.get(q["kp"])
        q["node"] = nid or ""
        q["node_name"] = NODE_BY_ID[nid]["name"] if nid else "未映射到知识点树"
        if not q["n"]:
            q["n"] = meta["n_students"]
            q["n_right"] = _half_up(q["P"] * q["n"])
    qs.sort(key=lambda q: (len(q["no"]), q["no"]))
    return qs, meta


# ============================================================ 分层作业
LAYERS = [
    ("基础", 0.65, 1.01, 0.6, "得分率 ≥ 65%，面向全体，必做"),
    ("中等", 0.40, 0.65, 0.7, "得分率 40%–65%，面向大多数，必做为主"),
    ("拔高", -1.0, 0.40, 0.9, "得分率 < 40%，面向学有余力，选做"),
]


def layer_of(P):
    for name, lo, hi, factor, _ in LAYERS:
        if lo <= P < hi:
            return name, factor
    return "中等", 0.7


def est_minutes(full, factor):
    return max(1, _half_up(float(full) * factor))


def build_layer(qs, meta):
    boxes = {name: [] for name, *_ in LAYERS}
    for q in qs:
        name, factor = layer_of(q["P"])
        boxes[name].append({
            "no": q["no"], "kp": q["kp"], "node": q["node_name"], "full": q["full"],
            "err": q["err"], "P": round(q["P"], 3), "rate": round(q["rate"], 4),
            "minutes": est_minutes(q["full"], factor),
        })
    doc = {"meta": {"source": meta["label"], "path": meta.get("path"),
                    "n_students": meta["n_students"],
                    "rule": "按逐题得分率 P 分层：≥0.65 基础 / 0.40–0.65 中等 / <0.40 拔高；"
                            "预计用时 = 满分 × 系数（基础0.6 / 中等0.7 / 拔高0.9），四舍五入，最小 1 分钟",
                    "no_new_questions": "只从已有题里分层挑选，不生成任何新题"},
           "layers": boxes}
    doc["layers"] = {k: sorted(v, key=lambda x: (len(x["no"]), x["no"])) for k, v in boxes.items()}
    return doc


def cmd_layer(args):
    qs, meta = load_questions(args.src, args.db, args.rows)
    doc = build_layer(qs, meta)
    path = jdump(doc, os.path.join(out_dir(), "分层作业.json"))
    print("=" * 78)
    print("分层作业包  |  数据源: %s" % meta["label"])
    print("  %s" % doc["meta"]["rule"])
    print("  说明: 只从已有 %d 道题里挑，不生成新题。" % len(qs))
    print("=" * 78)
    total_all = 0
    for name, lo, hi, factor, note in LAYERS:
        items = doc["layers"][name]
        sub = sum(i["minutes"] for i in items)
        total_all += sub
        print()
        print("[%s层] %d 题  预计用时合计 %d 分钟    （%s）" % (name, len(items), sub, note))
        if not items:
            print("   本卷无该层题目。")
            continue
        print("   %-4s %-14s %-18s %-5s %-7s %-4s" % ("题号", "知识点", "对应知识点树", "满分", "得分率", "用时"))
        for i in items:
            print("   %-4s %-14s %-18s %-5g %-7s %d 分" % (
                i["no"], i["kp"], i["node"], i["full"], "%.0f%%" % (i["P"] * 100), i["minutes"]))
        print("   题号: %s" % "、".join(i["no"] for i in items))
    print()
    print("-" * 78)
    print("三层合计: %d 题, %d 分钟（若按 40 分钟课后作业，建议：基础全做 + 中等选做 + 拔高选 1–2 题）" % (len(qs), total_all))
    print("清单文件: %s (%d 字节)" % (path, os.path.getsize(path)))
    return 0


# ============================================================ 讲评清单
WHY = {
    "概念不清": "「{kp}」的定义/法则没立住：正例、反例各给一个对照，让学生用自己的话复述，再用 1 道同类题即时验证",
    "计算错误": "不是不会，是过程崩：现场把这道题拆成单步，标出第一个出错的那一步，只练这一步的两三遍",
    "审题错误": "读题环节失分：圈出题干关键词，把「已知 → 所求」写成两列对照，再重读一遍题",
    "步骤遗漏": "过程写不完整：发一张步骤清单（每步打勾），让学生按清单把本题重新写一遍",
}
WHY_DEFAULT = "先请一名做错的学生说思路，再顺着他的思路讲「为什么不能这么想」，最后给同类题 1 道"


def alloc_minutes(items, budget=40, lo=3, hi=8, key="rate"):
    """按权重把总时长分配到每题，单题 [lo, hi]，合计不超过 budget。"""
    if not items:
        return []
    ws = [float(it.get(key) or 1e-9) for it in items]
    wmin, wmax = min(ws), max(ws)
    raw = [lo if wmax <= wmin else lo + (hi - lo) * (w - wmin) / (wmax - wmin) for w in ws]
    mins = [_half_up(x) for x in raw]
    mins = [max(lo, min(hi, m)) for m in mins]
    if lo * len(mins) > budget:          # 题量太大，连下限都放不下
        mins = [lo] * len(mins)
        while sum(mins) > budget and len(mins) > 1:
            mins.pop()
        return mins
    while sum(mins) > budget:            # 只削最小值里最大的那些，保证都 >= lo
        cand = [(m, i) for i, m in enumerate(mins) if m > lo]
        if not cand:
            break
        cand.sort(key=lambda t: (-t[0], ws[t[1]]))
        mins[cand[0][1]] -= 1
    return mins


def build_review(qs, meta, top=5, budget=40, time_by="rate"):
    ranked = sorted(qs, key=lambda q: (-q["rate"], -q["full"], int(q["no"])))
    picked = ranked[:top]
    key = "rate" if time_by == "rate" else "loss"
    for q in picked:
        q["loss"] = q["rate"] * q["full"]
    mins = alloc_minutes(picked, budget=budget, key=key)
    rows = []
    for q, m in zip(picked, mins):
        why = WHY.get(q["err"], WHY_DEFAULT).format(kp=q["kp"])
        rows.append({
            "no": q["no"], "kp": q["kp"], "node": q["node_name"], "full": q["full"],
            "err": q["err"], "rate": round(q["rate"], 4), "n_wrong": q["n"] - q["n_right"], "n": q["n"],
            "minutes": m, "why": why,
        })
    total = sum(r["minutes"] for r in rows)
    return {
        "meta": {"source": meta["label"], "path": meta.get("path"), "n_students": meta["n_students"],
                 "order": "按错误率降序" if time_by == "rate" else "按错误率降序，时长按失分总量分配",
                 "time_rule": "单题 3–8 分钟，按%s在选中题之间线性分配，合计 ≤ %d 分钟" % (
                     "错误率" if time_by == "rate" else "失分总量(=错误率×满分)", budget),
                 "total_minutes": total, "budget": budget, "slack": budget - total},
        "items": rows,
    }


def cmd_review(args):
    if getattr(args, "student", None):
        return cmd_review_student(args)
    if getattr(args, "scores", None):
        print("提示: --scores 只作用于 --student 个体版；要看全班的逐题得分清单，请用 review2 --scores。")
    qs, meta = load_questions(args.src, args.db, args.rows)
    doc = build_review(qs, meta, args.top, args.budget, args.time_by)
    path = jdump(doc, os.path.join(out_dir(), "讲评清单.json"))
    print("=" * 78)
    print("讲评清单（一节课 %d 分钟上限）" % doc["meta"]["budget"])
    print("  排序: %s ｜ 数据源: %s ｜ 学生数: %s" % (doc["meta"]["order"], meta["label"], meta["n_students"]))
    print("  时长规则: %s" % doc["meta"]["time_rule"])
    print("=" * 78)
    print("%-3s %-4s %-16s %-6s %-9s %-6s %s" % ("#", "题号", "知识点(满分)", "错误率", "错/总", "时长", "讲什么"))
    print("-" * 78)
    for i, r in enumerate(doc["items"], 1):
        print("%-3d %-4s %-16s %-6s %-9s %-6s %s" % (
            i, r["no"], "%s(%g分)" % (r["kp"], r["full"]), "%.1f%%" % (r["rate"] * 100),
            "%d/%d" % (r["n_wrong"], r["n"]), "%d 分钟" % r["minutes"], r["why"]))
    print("-" * 78)
    print("合计: %d 分钟 ｜ 上限 %d 分钟 ｜ 预留机动 %d 分钟" % (
        doc["meta"]["total_minutes"], doc["meta"]["budget"], doc["meta"]["slack"]))
    assert doc["meta"]["total_minutes"] <= doc["meta"]["budget"], "讲评时长超出预算"
    print("清单文件: %s (%d 字节)" % (path, os.path.getsize(path)))
    return 0


# ============================================================ 个体讲评 / 集体补漏（读逐题得分）
UNAVAIL = "不可用（缺逐题得分）"
SCORES_NEED = "要这两个视角，必须先有一份「逐题得分」表，表头：学生,题号,知识点,满分,得分（知识点、满分可省）。"


def _no_sort_key(no):
    m = re.search(r"\d+", str(no))
    return (int(m.group(0)) if m else 99999, str(no))


def _pick_student(names, want):
    """先精确匹配，再包含匹配。返回 (命中的名字 或 None, 候选列表)。"""
    w = str(want or "").strip()
    if w in names:
        return w, []
    hit = [n for n in names if w and w in n]
    if len(hit) == 1:
        return hit[0], []
    return None, hit


def _xlsx_probe(path):
    """只读探一下这张成绩单：多少人、几班、认不认得出「每题一列」。
    用 ingest.py 的公开 API，不改它。读不动就返回 err。"""
    out = {"err": "", "people": 0, "classes": 0, "sheets": 0, "qcols": 0, "names": []}
    if not path or not os.path.exists(path):
        out["err"] = "找不到这个文件：%s" % (path or "（没给路径）")
        return out
    try:
        import ingest
    except Exception as e:                                    # noqa: BLE001
        out["err"] = "调不到同目录的 ingest.py（%s）" % e
        return out
    try:
        recs, rep = ingest.read_all_sheets(path)
    except Exception as e:                                    # noqa: BLE001
        out["err"] = "这张成绩单读不动：%s" % e
        return out
    out["sheets"] = len(rep.get("sheets") or [])
    out["people"] = len(recs)
    out["classes"] = len({str(r.get("cls") or "").strip() for r in recs if str(r.get("cls") or "").strip()})
    out["names"] = [str(r.get("name") or "").strip() for r in recs if str(r.get("name") or "").strip()]
    seen = set()
    for r in recs:
        sh = r.get("__sheet")
        if sh in seen:
            continue
        seen.add(sh)
        try:
            out["qcols"] += len(ingest.detect_question_cols(list((r.get("raw") or {}).keys())))
        except Exception:                                     # noqa: BLE001
            pass
    return out


def _say_missing_scores(xlsx, tail_lines, who=None):
    """统一输出「不可用（缺逐题得分）」，并说清楚缺什么、怎么补。退出码由调用方给 0。"""
    print()
    if xlsx:
        info = _xlsx_probe(xlsx)
        if info["err"]:
            print("  成绩单先没读成：%s" % info["err"])
        else:
            print("  这张成绩单读到了：%d 人 / %d 班 / %d 个表 ｜ %s"
                  % (info["people"], info["classes"], info["sheets"], os.path.basename(xlsx)))
            if info["qcols"]:
                print("  表里确实有 %d 个「每题一列」的题号列，但本命令只吃标准逐题得分表"
                      "（学生,题号,知识点,满分,得分），请先另存成那种表。" % info["qcols"])
            else:
                print("  但它只有科目总分，一题都没有 → 算不出「哪几题」。")
            if who and info["names"]:
                hit = [n for n in info["names"] if n == who] or [n for n in info["names"] if who in n]
                if hit:
                    print("  这个人表里有：%s" % "、".join(hit[:3]))
                else:
                    print("  另外：表里没找到「%s」这个名字，先核对一下写法。" % who)
    print(UNAVAIL)
    for line in tail_lines:
        print("  " + line)
    print("  不按总分估算、不猜哪题错——这是口径，不放宽。")


def _student_unavailable(args, name):
    print()
    print("学生：%s" % name)
    xlsx = getattr(args, "xlsx", None)
    _say_missing_scores(xlsx, [
        "没有逐题得分，就看不出「%s 哪几题不会」，个体讲评不出。" % name,
        SCORES_NEED,
        "拿到之后这样跑：python teaching.py review --student %s --scores <逐题得分.csv>" % name,
    ], who=name)
    return 0


def cmd_review_student(args):
    name = str(args.student).strip()
    print("=" * 78)
    print("讲评清单 · 个体版（学生：%s）" % name)
    print("=" * 78)
    if not getattr(args, "scores", None):
        return _student_unavailable(args, name)

    bank = load_bank(getattr(args, "bank", None))
    if bank["err"]:
        print(bank["err"])
        print("题库读不到，「讲什么」补不上，个体版这次不出。")
        return 2
    try:
        rows, wide, warn = _parse_scores(args.scores, bank)
    except ScoresError as e:
        print("--scores 这份表用不了：")
        for line in str(e).splitlines():
            print("  " + line)
        return 2
    print("数据来源: 逐题得分 %s（%s ｜ %d 行 / %d 人）"
          % (os.path.basename(args.scores), "宽表" if wide else "长表",
             len(rows), len({r["stu"] for r in rows})))
    for w in warn[:4]:
        print("  ! %s" % w)

    names = sorted({r["stu"] for r in rows})
    who, cand = _pick_student(names, name)
    if not who:
        print()
        print("这份逐题得分里（共 %d 人）找不到「%s」。" % (len(names), name))
        if cand:
            print("  像是这几个：%s" % "、".join(cand[:10]))
        else:
            print("  表里的学生：%s%s" % ("、".join(names[:20]), "…" if len(names) > 20 else ""))
        return 2
    if who != name:
        print("  按「%s」匹配到表里的「%s」。" % (name, who))

    mine = [r for r in rows if r["stu"] == who]
    mine.sort(key=lambda r: _no_sort_key(r["no"]))
    tot_full = sum(r["full"] for r in mine)
    tot_got = sum(r["got"] for r in mine)
    print()
    print("这次考试: %d 题 ｜ 满分 %g ｜ 实得 %g ｜ 失分 %g"
          % (len(mine), tot_full, tot_got, tot_full - tot_got))

    lost = [r for r in mine if r["full"] and r["got"] < r["full"] - 1e-9]
    if not lost:
        print()
        print("这名学生这次没有失分题（%d 题全对），个体讲评不用出。" % len(mine))
        return 0

    item_by_no = {}
    for it in bank["items"]:
        item_by_no[str(it.get("no"))] = it
    for r in lost:
        it = item_by_no.get(r["no"], {})
        r["loss"] = r["full"] - r["got"]
        r["err"] = str(it.get("err") or "")
        r["kp"] = str(it.get("kp") or r.get("kp") or "未标注")
        r["stem"] = str(it.get("stem") or "")
    lost.sort(key=lambda r: (-r["loss"], -r["full"], _no_sort_key(r["no"])))

    sched, rest = plan_schedule(lost, args.budget, 3, 8, args.top, key="loss")
    total_min = sum(m for _, m in sched)
    print()
    print("%-3s %-5s %-16s %-8s %-5s %-14s %s"
          % ("#", "题号", "知识点(满分)", "得分", "失分", "讲多久", "讲什么"))
    print("-" * 78)
    for i, (r, m) in enumerate(sched, 1):
        print("%-3d %-5s %-16s %-8s %-5s %-14s %s"
              % (i, r["no"], "%s(%g分)" % (r["kp"], r["full"]),
                 "%g/%g" % (r["got"], r["full"]), "%g" % r["loss"],
                 "%d 分 / %d 秒" % (m, m * 60),
                 WHY.get(r["err"], WHY_DEFAULT).format(kp=r["kp"])))
    print("-" * 78)
    print("合计: %d 分钟 = %d 秒 ｜ 上限 %d 分钟 ｜ 富余 %d 分钟"
          % (total_min, total_min * 60, args.budget, args.budget - total_min))
    print("口径: 只含 %s 本人失分的题（共 %d 道），按失分从大到小排；单题 3–8 分钟，按失分分配。"
          % (who, len(lost)))
    if rest:
        print()
        print("这次排不下，下次继续（%d 道）:" % len(rest))
        for r in rest:
            print("  第 %s 题 · %s · 失分 %g" % (r["no"], r["kp"], r["loss"]))

    shown = [r for r, _ in sched]
    if any(r["stem"] for r in shown):
        print()
        print("每题题干（给老师对卷用）:")
        for r in shown:
            print("  第 %s 题：%s" % (r["no"], (r["stem"] or "（题库里这题还没录题面）")))
    nostem = [r for r in shown if not r["stem"]]
    if nostem:
        print("  注: 第 %s 题题库里还没录题面，讲之前先看一眼原卷。"
              % "、".join(r["no"] for r in nostem))

    # ---- 自检：个体题号集合 ⊆ 全班失分题集合 ----
    class_lost, agg = set(), {}
    for r in rows:
        a = agg.setdefault(r["no"], {"got": 0.0, "tot": 0.0})
        a["got"] += r["got"]
        a["tot"] += r["full"]
        if r["full"] and r["got"] < r["full"] - 1e-9:
            class_lost.add(r["no"])
    mine_nos = {r["no"] for r in lost}
    ok = mine_nos <= class_lost
    print()
    print("自检 1: 个体题号集合 ⊆ 全班失分题集合 → %s（个体 %d 题 / 全班失分题 %d 题）"
          % ("通过" if ok else "不通过", len(mine_nos), len(class_lost)))
    ranked_all = sorted(agg.items(), key=lambda kv: (-(1 - kv[1]["got"] / kv[1]["tot"]) if kv[1]["tot"] else 0,
                                                     _no_sort_key(kv[0])))
    top_class = {no for no, _ in ranked_all[:args.top]}
    inter = sorted(mine_nos & top_class, key=_no_sort_key)
    print("自检 2（参考，不做断言）: 全班按错误率只讲前 %d 题时，与个体清单重合 %d 题：%s"
          % (args.top, len(inter), "、".join(inter) or "无"))
    if not ok:
        print("★ 自检 1 没过：个体里出现了全班失分题之外题号，先别用这页。")
        return 3

    doc = {"student": who, "source": os.path.basename(args.scores),
           "items": [{"no": r["no"], "kp": r["kp"], "full": r["full"], "got": r["got"],
                      "loss": r["loss"], "minutes": m,
                      "why": WHY.get(r["err"], WHY_DEFAULT).format(kp=r["kp"])}
                     for r, m in sched],
           "rest": [{"no": r["no"], "kp": r["kp"], "loss": r["loss"]} for r in rest],
           "total_minutes": total_min, "budget": args.budget,
           "subset_check": ok}
    fname = "个体讲评_%s.json" % re.sub(r'[\\/:*?"<>|]+', "_", who)
    path = jdump(doc, os.path.join(out_dir(), fname))
    print()
    print("清单元数据: %s (%d 字节)" % (path, os.path.getsize(path)))
    return 0


# ============================================================ 集体补漏（按知识点聚合失分）
GAP_RULE = "集体讲的门槛：该知识点错误率 ≥ 40% 且失分的同学 ≥ 3 人（规则，不是因果结论）"


def _gaps_unavailable(args):
    print("=" * 78)
    print("集体补漏清单（按知识点聚合失分）")
    print("=" * 78)
    xlsx = getattr(args, "xlsx", None)
    _say_missing_scores(xlsx, [
        "没有逐题得分，就不知道哪个知识点集体塌了，补漏清单不出。",
        SCORES_NEED,
        "拿到之后这样跑：python teaching.py gaps --scores 第1次月考_逐题得分.csv",
    ])
    return 0


def cmd_gaps(args):
    paths = getattr(args, "scores", None) or []
    if isinstance(paths, str):
        paths = [paths]
    paths = [p for p in paths if str(p).strip()]
    if not paths:
        return _gaps_unavailable(args)

    print("=" * 78)
    print("集体补漏清单（按知识点聚合失分）")
    print("=" * 78)
    bank = load_bank(getattr(args, "bank", None))
    if bank["err"]:
        print(bank["err"])
        print("题库读不到，知识点名补不上；仍按逐题得分表里的「知识点」列聚合。")

    rows, labels, warns = [], [], []
    for p in paths:
        try:
            part, wide, warn = _parse_scores(p, bank)
        except ScoresError as e:
            print("--scores %s 这份表用不了：" % os.path.basename(p))
            for line in str(e).splitlines():
                print("  " + line)
            return 2
        rows.extend(part)
        labels.append("%s（%s ｜ %d 行 / %d 人）" % (os.path.basename(p), "宽表" if wide else "长表",
                                                   len(part), len({r["stu"] for r in part})))
        warns.extend(warn)
    for lab in labels:
        print("  读入: %s" % lab)
    for w in warns[:6]:
        print("  ! %s" % w)

    bank_nos = {str(it.get("no")) for it in (bank.get("items") or [])}
    unmatched = sorted({r["no"] for r in rows} - bank_nos, key=_no_sort_key) if bank_nos else []
    students = {r["stu"] for r in rows}

    g = {}
    for r in rows:
        kp = str(r.get("kp") or "").strip() or "未标注"
        a = g.setdefault(kp, {"kp": kp, "full": 0.0, "got": 0.0, "n": 0, "nos": set(),
                              "studs": set(), "lost_studs": set(), "loss": 0.0})
        a["full"] += r["full"]
        a["got"] += r["got"]
        a["n"] += 1
        a["loss"] += r["full"] - r["got"]
        a["nos"].add(r["no"])
        a["studs"].add(r["stu"])
        if r["full"] and r["got"] < r["full"] - 1e-9:
            a["lost_studs"].add(r["stu"])

    if not g:
        print()
        print("逐题得分读进来了，但按知识点聚合出来是空的（表里的知识点列全空？）。补漏清单不出。")
        return 0

    def rate_of(a):
        return (1.0 - a["got"] / a["full"]) if a["full"] else 0.0

    ranked = sorted(g.values(), key=lambda a: (-a["loss"], -rate_of(a), a["kp"]))
    print()
    print("共 %d 人 ｜ %d 道题 ｜ %d 个知识点 ｜ 失分合计 %g"
          % (len(students), len({r["no"] for r in rows}), len(g), sum(a["loss"] for a in g.values())))
    print(GAP_RULE)
    if unmatched:
        print("题库里对不上：%s" % "、".join(unmatched))
    print()
    print("%-3s %-20s %-5s %-13s %-10s %-8s %s"
          % ("#", "知识点", "题数", "涉及人数", "失分总分", "错误率", "判定"))
    print("-" * 78)
    for i, a in enumerate(ranked, 1):
        hit = rate_of(a) >= 0.40 and len(a["lost_studs"]) >= 3
        print("%-3d %-20s %-5d %-13s %-10g %-8s %s"
              % (i, a["kp"], len(a["nos"]),
                 "%d / %d" % (len(a["lost_studs"]), len(a["studs"])),
                 a["loss"], "%.1f%%" % (rate_of(a) * 100),
                 "建议集体讲" if hit else "个别人补"))
    print("-" * 78)
    need = [a for a in ranked if rate_of(a) >= 0.40 and len(a["lost_studs"]) >= 3]
    if need:
        print("先讲这三个（按失分从多到少）: %s"
              % " / ".join("%s（失分 %g，%d 人）" % (a["kp"], a["loss"], len(a["lost_studs"]))
                           for a in need[:3]))
    else:
        print("没有达到「集体讲」门槛的知识点：失分是散的，按人补更划算。")
    print("口径: 错误率 = 1 − 该知识点得分合计/满分合计；涉及人数 = 该知识点至少错一题的人数；"
          "只用逐题得分，不做任何总分估算。")

    doc = {"source": labels, "n_students": len(students),
           "rule": GAP_RULE, "unmatched": unmatched,
           "items": [{"kp": a["kp"], "n_items": len(a["nos"]), "loss": round(a["loss"], 3),
                      "rate": round(rate_of(a), 4), "lost_students": len(a["lost_studs"]),
                      "students": len(a["studs"]),
                      "nos": sorted(a["nos"], key=_no_sort_key)}
                     for a in ranked]}
    path = jdump(doc, os.path.join(out_dir(), "集体补漏清单.json"))
    print()
    print("清单文件: %s (%d 字节)" % (path, os.path.getsize(path)))
    return 0


# ============================================================ 打印中心
CSS = """
@page { size: A4; margin: 14mm 12mm; }
* { box-sizing: border-box; }
body { font-family: "Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", sans-serif;
       color: #000; margin: 0; font-size: 12pt; line-height: 1.55; }
h1 { font-size: 19pt; text-align: center; margin: 0 0 2mm; letter-spacing: 3px; }
.sub { text-align: center; color: #444; font-size: 9.5pt; margin-bottom: 5mm; }
.meta { display: flex; justify-content: space-between; border-bottom: 1.5px solid #000;
        padding-bottom: 2mm; margin-bottom: 4mm; font-size: 10.5pt; }
table { width: 100%; border-collapse: collapse; }
th, td { border: 1px solid #000; padding: 1.6mm 2mm; font-size: 10.5pt; vertical-align: top; }
th { background: #ececec; text-align: left; }
.item { border: 1px solid #000; padding: 3mm; margin-bottom: 4mm; page-break-inside: avoid; }
.item .h { display: flex; justify-content: space-between; font-weight: bold; margin-bottom: 2mm; font-size: 11pt; }
.q { margin: 0 0 2mm; }
.stem { border-left: 3px solid #999; padding-left: 3mm; color: #333; margin-bottom: 2mm; }
.ln { border-bottom: 1px solid #b0b0b0; height: 9mm; }
.note { font-size: 9pt; color: #555; }
.cell-n { font-weight: bold; margin-bottom: 1mm; }
.cell-blank { height: 12mm; border-bottom: 1px dashed #bbb; }
.answers { page-break-before: always; }
.answers td { text-align: center; }
.foot { margin-top: 6mm; font-size: 9pt; color: #666; border-top: 1px dashed #999; padding-top: 2mm; }
ul { margin: 0 0 2mm 5mm; padding: 0; }
li { margin-bottom: 1mm; }
@media print { .noprint { display: none !important; } }
"""

PAGE_HEAD = ("<!DOCTYPE html>\n<html lang=\"zh-CN\">\n<head>\n<meta charset=\"utf-8\">\n"
             "<title>%s</title>\n<style>%s</style>\n</head>\n<body>\n")


def _head(title):
    return PAGE_HEAD % (html.escape(title), CSS)


def _meta_line(left, right):
    return '<div class="meta"><div>%s</div><div>%s</div></div>\n' % (left, right)


def _write_html(name, text):
    path = os.path.join(out_dir(), name)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    return path, os.path.getsize(path)


def html_wrong_quiz(qs, meta, top=8, min_rate=0.5):
    picked = [q for q in sorted(qs, key=lambda q: (-q["rate"], int(q["no"]))) if q["rate"] >= min_rate][:top]
    if not picked:
        picked = sorted(qs, key=lambda q: (-q["rate"], int(q["no"])))[:top]
    b = [_head("错题卷")]
    b.append("<h1>错 题 卷</h1>\n")
    b.append('<div class="sub">来源：%s ｜ 取错误率最高的 %d 题（阈值 %.0f%%）｜ 纯 HTML+CSS，Ctrl+P 打印</div>\n'
             % (html.escape(meta["label"]), len(picked), min_rate * 100))
    b.append(_meta_line("班级：____________　姓名：____________", "学号：__________　得分：________"))
    for i, q in enumerate(picked, 1):
        b.append('<div class="item">\n')
        b.append('<div class="h"><span>第 %d 题　%s（满分 %g 分）</span>'
                 '<span>错误率 %.1f%%</span></div>\n'
                 % (i, html.escape(q["kp"]), q["full"], q["rate"] * 100))
        b.append('<div class="q">知识点：%s ｜ 本卷题号：第 %s 题 ｜ 主要错因：%s</div>\n'
                 % (html.escape(q["node_name"]), html.escape(q["no"]), html.escape(q["err"] or "未标注")))
        b.append('<div class="stem">题干：<span class="note">该题题面未录入题库，请老师按原卷粘贴</span></div>\n')
        b.append('<div class="ln"></div>\n<div class="ln"></div>\n<div class="ln"></div>\n')
        b.append('<div class="q note">解：</div>\n')
        b.append("</div>\n")
    b.append('<div class="foot">共 %d 题。答题区行数固定 3 行，解答题请另附草稿纸。</div>\n' % len(picked))
    b.append("</body>\n</html>\n")
    return "".join(b), picked


def _same_type(node_id, qmap, want=3):
    """返回 (picked, n_same)。picked=[(题号, 关系)]，关系: 同点/同族/前置。"""
    picked = []

    def add(nid, rel, pool):
        for q in pool.get(nid, []):
            if len(picked) >= want:
                return
            if q["no"] in [p[0] for p in picked]:
                continue
            picked.append((q["no"], rel))

    add(node_id, "同点", qmap)
    n_same = len(picked)
    me = NODE_BY_ID[node_id]
    if len(picked) < want:                      # 同族：同章且前置有交集
        for n in NODES:
            if n["chapter"] == me["chapter"] and n["id"] != node_id and set(n["prereq"]) & set(me["prereq"]):
                add(n["id"], "同族", qmap)
    if len(picked) < want:                      # 前置：补基础
        for p in me["prereq"]:
            if p in NODE_BY_ID:
                add(p, "前置", qmap)
    picked.sort(key=lambda t: (len(t[0]), t[0]))
    return picked, n_same


def html_review_sheet(qs):
    qmap = {}
    for q in qs:
        if q["node"]:
            qmap.setdefault(q["node"], []).append(q)
    rows = []
    for n in NODES:
        if n["chapter"] != "rat":
            continue
        same, n_same = _same_type(n["id"], qmap, 3)
        rows.append((n, same, n_same))
    b = [_head("复习单")]
    b.append("<h1>有 理 数 · 复 习 单</h1>\n")
    b.append('<div class="sub">知识点 / 前置 / 同类题号　｜　纯 HTML+CSS，Ctrl+P 打印</div>\n')
    b.append(_meta_line("班级：____________　姓名：____________", "日期：____ 月 ____ 日"))
    b.append("<table>\n<thead><tr><th style=\"width:15%\">知识点</th><th style=\"width:22%\">前置知识</th>"
             "<th style=\"width:9%\">本卷题号</th><th style=\"width:31%\">3 道同类题号</th><th>自查</th></tr></thead>\n<tbody>\n")
    for n, same, n_same in rows:
        inpaper = "、".join(q["no"] for q in qmap.get(n["id"], [])) or "—"
        cell = "、".join("%s(%s)" % (no, rel) for no, rel in same)
        if not same:
            cell = '<span class="note">本卷无同类题，需补题</span>'
        elif n_same < 3:
            cell = ('<span class="note">同点题不足 3 道（本卷 %d 道），已用邻近知识点补足：</span>%s'
                    % (n_same, cell))
        pre = "、".join(NODE_BY_ID[p]["name"] for p in n["prereq"]) or "—（本章起点）"
        b.append("<tr><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td>□□□</td></tr>\n"
                 % (html.escape(n["name"]), html.escape(pre), html.escape(inpaper), cell))
    b.append("</tbody>\n</table>\n")
    b.append('<div class="foot">标「同点」= 本卷同知识点；「同族」= 同章且前置相同的邻近知识点；'
             '「前置」= 该知识点的前置知识（补基础用）。自查栏每学完一项打一个勾。</div>\n')
    b.append("</body>\n</html>\n")
    return "".join(b), rows


def gen_oral(seed=20261028, n=20):
    """口算卡：确定性规则生成（固定种子），不调用任何模型。"""
    rnd = random.Random(seed)
    items = []
    while len(items) < n:
        kind = rnd.choice(["add", "add", "sub", "mul", "mul", "div", "pow", "abs", "mix"])
        if kind == "add":
            a, b = rnd.randint(-20, 20), rnd.randint(-20, 20)
            items.append(("%d + (%d)" % (a, b) if b < 0 else "%d + %d" % (a, b), a + b))
        elif kind == "sub":
            a, b = rnd.randint(-20, 20), rnd.randint(-20, 20)
            items.append(("%d - (%d)" % (a, b) if b < 0 else "%d - %d" % (a, b), a - b))
        elif kind == "mul":
            a, b = rnd.randint(-9, 9), rnd.randint(-9, 9)
            items.append(("%d × (%d)" % (a, b) if b < 0 else "%d × %d" % (a, b), a * b))
        elif kind == "div":
            b = rnd.choice([-9, -8, -6, -5, -4, -3, -2, 2, 3, 4, 5, 6, 8, 9])
            c = rnd.randint(-9, 9)
            a = b * c
            items.append(("%d ÷ (%d)" % (a, b) if b < 0 else "%d ÷ %d" % (a, b), c))
        elif kind == "pow":
            a = rnd.randint(2, 6)
            k = rnd.choice([2, 3])
            items.append(("(-%d)^%d" % (a, k), (-a) ** k))
        elif kind == "abs":
            a = rnd.randint(-15, 15)
            items.append(("|%d|" % a, abs(a)))
        else:
            a, b, c = rnd.randint(-6, 6), rnd.randint(-6, 6), rnd.randint(-4, 4)
            items.append(("%d + %d × (%d)" % (a, b, c) if c < 0 else "%d + %d × %d" % (a, b, c),
                          a + b * c))
    return items


def html_oral_card(items):
    b = [_head("口算卡")]
    b.append("<h1>口 算 卡</h1>\n")
    b.append('<div class="sub">20 题　限时 10 分钟　｜　算式由固定种子确定性生成（非题库题，0 次模型调用）</div>\n')
    b.append(_meta_line("班级：____________　姓名：____________", "用时：______ 分　正确：____ 题"))
    b.append("<table>\n")
    for r in range(5):
        b.append("<tr>\n")
        for c in range(4):
            i = r * 4 + c
            q = items[i][0]
            b.append('<td><div class="cell-n">%d. %s =</div><div class="cell-blank"></div></td>\n' % (i + 1, html.escape(q)))
        b.append("</tr>\n")
    b.append("</table>\n")
    b.append('<div class="foot">批改标准：一题 1 分，满分 20 分。第 13–16 题是乘方与绝对值，符号写错按错处理。</div>\n')
    b.append('<div class="answers">\n<h1>答 案 页</h1>\n')
    b.append('<div class="sub">发给学生前请沿虚线撕下</div>\n')
    b.append("<table>\n")
    for r in range(5):
        b.append("<tr>\n")
        for c in range(4):
            i = r * 4 + c
            b.append('<td><div class="cell-n">%d.</div>%s</td>\n' % (i + 1, html.escape(str(items[i][1]))))
        b.append("</tr>\n")
    b.append("</table>\n</div>\n")
    b.append("</body>\n</html>\n")
    return "".join(b)


def cmd_print(args):
    qs, meta = load_questions(args.src, args.db, args.rows)
    txt1, picked = html_wrong_quiz(qs, meta, args.quiz_top, args.min_rate)
    p1, s1 = _write_html("错题卷.html", txt1)
    txt2, rows = html_review_sheet(qs)
    p2, s2 = _write_html("复习单.html", txt2)
    items = gen_oral(args.seed, 20)
    txt3 = html_oral_card(items)
    p3, s3 = _write_html("口算卡.html", txt3)
    print("=" * 78)
    print("打印中心（纯 HTML+CSS，浏览器 Ctrl+P，A4）")
    print("=" * 78)
    for p, s in ((p1, s1), (p2, s2), (p3, s3)):
        print("%-52s %7d 字节" % (p, s))
    print()
    print("错题卷: 收录 %d 题（错误率阈值 %.0f%%）→ 题号 %s" % (
        len(picked), args.min_rate * 100, "、".join(q["no"] for q in picked)))
    n3 = sum(1 for _, _, n_same in rows if n_same >= 3)
    print("复习单: 覆盖有理数 %d 个知识点，其中 %d 个本就有 3 道同点题，%d 个同点题不足（已用邻近知识点补足并标注）" % (
        len(rows), n3, len(rows) - n3))
    print("口算卡: 20 题（种子 %d）+ 独立答案页" % args.seed)
    print()
    print("关键结构抽样（错题卷，前 10 行）:")
    for line in txt1.splitlines()[:10]:
        print("  | " + line[:110])
    print("关键结构抽样（复习单，表格行前 6 行）:")
    for line in [l for l in txt2.splitlines() if l.startswith("<tr>")][:6]:
        print("  | " + line[:110])
    print("关键结构抽样（口算卡，前 6 行）:")
    for line in txt3.splitlines()[:6]:
        print("  | " + line[:110])
    return 0


# ============================================================ 答题卡导入
def _read_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = [r for r in csv.reader(f) if any(c.strip() for c in r)]
    return rows[0], rows[1:]


def detect_mode(sample_rows):
    cells = [c.strip().upper() for row in sample_rows for c in row[2:] if c.strip()]
    opt = {"A", "B", "C", "D", "AB", "AC", "AD", "BC", "BD", "CD", "ABC", "ABD", "ACD", "BCD", "ABCD"}
    return "option" if cells and all(c in opt for c in cells) else "score"


def import_card(csv_path, paper_path):
    """答题卡导出表 -> 逐题统计（选项模式/得分模式自动识别）。"""
    paper = load_paper_struct(paper_path)
    head, body = _read_csv(csv_path)
    mode = detect_mode(body[:5])
    qnos = [h.strip() for h in head[2:]]
    stats, col_detail = [], {}
    for idx, no in enumerate(qnos, start=2):
        info = paper.get(no, {})
        full = float(info.get("full", 1))
        answer = info.get("answer")
        got, wrong_pick, blank, noscore = [], {}, 0, 0
        for r in body:
            v = r[idx].strip() if idx < len(r) else ""
            if not v:
                blank += 1
                got.append(0.0)
                continue
            if mode == "option":
                if not answer:
                    noscore += 1
                    got.append(0.0)
                    continue
                ok = v.upper() == str(answer).upper()
                got.append(full if ok else 0.0)
                if not ok:
                    wrong_pick[v.upper()] = wrong_pick.get(v.upper(), 0) + 1
            else:
                try:
                    got.append(float(v))
                except ValueError:
                    got.append(0.0)
        P = sum(got) / (full * len(body)) if body else 0.0
        n_right = sum(1 for g in got if g >= full)
        top_wrong = sorted(wrong_pick.items(), key=lambda kv: (-kv[1], kv[0]))[:2]
        stats.append({
            "no": no, "kp": info.get("kp", "未标注"), "full": full, "err": info.get("err", ""),
            "answer": answer, "P": round(P, 3), "rate": round(1 - P, 3), "n": len(body),
            "n_right": n_right, "n_blank": blank, "n_unscored": noscore,
            "judge": "白给分" if P >= 0.9 else ("太难" if P <= 0.2 else "正常"),
            "top_wrong": top_wrong,
        })
        col_detail[no] = wrong_pick
    return {"mode": mode, "head": head, "n_students": len(body),
            "qnos": qnos, "stats": stats}, col_detail, paper


def cmd_import(args):
    doc, detail, paper = import_card(args.csv, args.paper)
    od = out_dir()
    print("=" * 78)
    print("答题卡导入（并入 import_scores.py 的能力）")
    print("  导出表: %s" % os.path.abspath(args.csv))
    print("  卷子结构: %s" % os.path.abspath(args.paper))
    print("=" * 78)
    print("识别为【%s模式】｜ 表头: %s" % ("选项" if doc["mode"] == "option" else "得分", doc["head"][:6]))
    print("学生数: %d ｜ 题目数: %d" % (doc["n_students"], len(doc["qnos"])))
    miss = [n for n in paper if n not in doc["qnos"]]
    extra = [n for n in doc["qnos"] if n not in paper]
    if miss or extra:
        print("！结构与卡不匹配：结构中有而卡中没有的题号 %s ｜ 卡中有而结构没有的 %s"
              % (miss or "无", extra or "无"))
    print()
    print("%-5s %-13s %-4s %-6s %-8s %-8s %-6s %s" % ("题号", "知识点", "分值", "得分率", "答对人数", "判定", "空答", "主要错选"))
    print("-" * 78)
    for s in doc["stats"]:
        tw = "、".join("%s×%d" % (k, v) for k, v in s["top_wrong"]) or "—"
        print("%-5s %-13s %-4g %-6.2f %-8d %-8s %-6d %s" % (
            s["no"], s["kp"], s["full"], s["P"], s["n_right"], s["judge"], s["n_blank"], tw))
    tot_full = sum(s["full"] for s in doc["stats"])
    tot_got = sum(s["P"] * s["full"] for s in doc["stats"])
    print("-" * 78)
    print("全班平均分：%.1f 分（满分 %g）　平均得分率 %.1f%%" % (
        tot_got, tot_full, 100 * tot_got / tot_full if tot_full else 0))
    js = jdump(doc["stats"], os.path.join(od, "card_export_分析.json"))
    print("逐题明细: %s (%d 字节)  ← 注意：新增文件，未覆盖 _tmp 下原 import_scores.py 产出的同名文件"
          % (js, os.path.getsize(js)))

    # 把同一批数据直接喂给讲评清单 + 分层作业
    qs = []
    for s in doc["stats"]:
        nid = ALIAS.get(s["kp"])
        qs.append({"no": s["no"], "kp": s["kp"], "node": nid or "", "node_name": NODE_BY_ID[nid]["name"] if nid else "未映射",
                   "full": s["full"], "err": s["err"], "n": s["n"], "n_right": s["n_right"],
                   "P": s["P"], "rate": s["rate"]})
    meta = {"src": "csv", "path": args.csv, "label": "读卡机导出表 %s" % os.path.basename(args.csv),
            "n_students": doc["n_students"]}
    rv = build_review(qs, meta, args.top, args.budget, args.time_by)
    print()
    print("—— 链路下一步 1：讲评清单 ——")
    print("%-3s %-4s %-16s %-6s %-8s %-6s %s" % ("#", "题号", "知识点(满分)", "错误率", "错/总", "时长", "讲什么"))
    print("-" * 78)
    for i, r in enumerate(rv["items"], 1):
        print("%-3d %-4s %-16s %-6s %-8s %-6s %s" % (
            i, r["no"], "%s(%g分)" % (r["kp"], r["full"]), "%.1f%%" % (r["rate"] * 100),
            "%d/%d" % (r["n_wrong"], r["n"]), "%d 分钟" % r["minutes"], r["why"]))
    print("合计 %d 分钟 ｜ 上限 %d 分钟 ｜ 预留机动 %d 分钟" % (
        rv["meta"]["total_minutes"], rv["meta"]["budget"], rv["meta"]["slack"]))

    ly = build_layer(qs, meta)
    print()
    print("—— 链路下一步 2：分层作业 ——")
    for name, *_ in LAYERS:
        items = ly["layers"][name]
        if not items:
            print("[%s层] 本卷无该层题目。" % name)
            continue
        print("[%s层] %d 题，合计 %d 分钟：%s" % (
            name, len(items), sum(i["minutes"] for i in items),
            "、".join("%s(%s,%d分)" % (i["no"], i["kp"], i["minutes"]) for i in items)))
    jl = jdump({"review": rv, "layer": ly},
               os.path.join(od, "答题卡_讲评与作业.json"))
    print()
    print("合并结果: %s (%d 字节)" % (jl, os.path.getsize(jl)))
    return 0


# ============================================================ 题库（#14 / #27 / #28）
BANK_PATH = os.path.join(SCRIPT_DIR, "题库.json")
OUT_HTML = os.path.join(SCRIPT_DIR, "输出")
# 题库里的知识点名 -> 知识点树节点。只给新命令用，老命令那张映射表一个字没动。
BANK_KP_TO_NODE = {"正负数概念": "rat-concept", "负数概念": "rat-concept"}
LAYER_PAGE = [("easy", "A", "基础卷", 0.6), ("mid", "B", "提高卷", 0.7), ("hard", "C", "挑战卷", 0.9)]
DIFF_ORDER = {"easy": 1, "mid": 2, "hard": 3}

CSS_EXTRA = """
.qbox{border:1px solid #000;padding:3mm 3.5mm;margin-bottom:3.5mm;page-break-inside:avoid}
.qhead{display:flex;justify-content:space-between;font-size:10pt;color:#333;margin-bottom:1.5mm}
.qno2{font-weight:700;color:#000}
.stem2{font-size:11.5pt;line-height:1.7}
.opts{display:flex;flex-wrap:wrap;gap:1.5mm 8mm;margin-top:1.5mm;font-size:11pt}
.ln2{border-bottom:1px solid #999;height:8.5mm}
.ansline{display:inline-block;border-bottom:1px solid #999;min-width:42mm;margin-left:2mm}
.keybox{border:1.5px dashed #666;padding:2.5mm 3mm;margin-top:5mm;font-size:10pt;page-break-inside:avoid}
.keytable{width:100%;border-collapse:collapse;margin-top:2mm}
.keytable td{border:1px solid #999;padding:1.2mm 2mm;font-size:10pt}
.tag{display:inline-block;border:1px solid #000;border-radius:2mm;padding:0.4mm 2mm;font-size:9.5pt;margin-left:2mm}
.warn{border:1px solid #c0392b;color:#a02c1e;background:#fdf2f1;padding:2.5mm 3mm;margin:3mm 0;font-size:10pt}
.warn ul{margin:1mm 0 0 5mm;padding:0}
.pgbreak{page-break-before:always}
.plan th,.plan td{font-size:10pt}
.mini{font-size:9pt;color:#444}
.sectitle{font-size:12pt;font-weight:700;margin:5mm 0 2.5mm;border-left:4px solid #000;padding-left:2.5mm}
.asklist{margin:0;padding-left:6mm;font-size:10.5pt}
.asklist li{margin-bottom:2mm}
@media print{ .warn{background:#fff} }
"""


def _head2(title):
    """和 _head 同一套版式，只多一层新页面样式；老的 _head 没动过。"""
    return PAGE_HEAD % (html.escape(title), CSS + CSS_EXTRA)


def safe_filename(name):
    base = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", str(name or "")).strip() or "未命名"
    if not base.lower().endswith(".html"):
        base += ".html"
    return base[:120]


def _write_html_out(name, text):
    if not os.path.isdir(OUT_HTML):
        os.makedirs(OUT_HTML)
    path = os.path.join(OUT_HTML, safe_filename(name))
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    return path, os.path.getsize(path)


def load_bank(path=None):
    """读题库。出错不抛异常，err 里放中文说明。"""
    p = path or BANK_PATH
    if not os.path.exists(p):
        return {"items": [], "meta": {}, "err": "找不到题库文件：%s" % p,
                "hint": "题库.json 要和 teaching.py 放在同一个目录里（教师工具\\题库.json）。", "path": p}
    try:
        with open(p, encoding="utf-8") as f:
            doc = json.load(f)
    except Exception as e:                                    # noqa: BLE001
        return {"items": [], "meta": {}, "err": "题库.json 打不开或不是合法 JSON：%s" % e, "hint": "", "path": p}
    items = doc.get("items")
    if not isinstance(items, list):
        return {"items": [], "meta": doc.get("meta", {}) or {}, "err": "题库.json 里没有 items 列表。",
                "hint": "", "path": p}
    return {"items": items, "meta": doc.get("meta", {}) or {}, "err": "", "hint": "", "path": p}


def has_stem(it):
    return bool(str(it.get("stem") or "").strip())


def split_bank(items):
    """拆成（有题干的, 没题干的）。没题干的一律不进选题池。"""
    ok = [i for i in items if has_stem(i)]
    bad = [i for i in items if not has_stem(i)]
    return ok, bad


def _no_key(it):
    m = re.search(r"\d+", str(it.get("no") or ""))
    return int(m.group(0)) if m else 9999


def bank_node(it):
    kp = str(it.get("kp") or "").strip()
    if kp in BANK_KP_TO_NODE:
        return BANK_KP_TO_NODE[kp]
    return ALIAS.get(kp, "")


def bank_index(items):
    """知识点节点 -> 该知识点的题（按题号）"""
    out = {}
    for it in items:
        nid = bank_node(it)
        if nid:
            out.setdefault(nid, []).append(it)
    for v in out.values():
        v.sort(key=_no_key)
    return out


def diff_of(it):
    d = str(it.get("difficulty") or "").strip().lower()
    return d if d in DIFF_ORDER else "mid"


def prereq_layers(nid):
    """前置知识点按距离由近及远分层：[[第1层], [第2层], ...]，层内按树的声明顺序。"""
    order = [n["id"] for n in NODES]
    seen = {nid}
    frontier = [nid]
    layers = []
    while frontier:
        nxt = []
        for cur in frontier:
            for p in NODE_BY_ID[cur]["prereq"]:
                if p in NODE_BY_ID and p not in seen:
                    seen.add(p)
                    nxt.append(p)
        if not nxt:
            break
        nxt.sort(key=lambda x: order.index(x) if x in order else 99)
        layers.append(nxt)
        frontier = nxt
    return layers


def stem_head(it, n=20):
    s = re.sub(r"\s+", " ", str(it.get("stem") or "")).strip()
    return s if len(s) <= n else s[:n] + "…"


def blank_lines(it):
    full = float(it.get("full") or 0)
    if it.get("options"):
        return 0
    return 2 if full <= 3 else (3 if full <= 6 else 5)


def q_html(it, idx, show_meta=True):
    b = ['<div class="qbox">\n']
    meta = "第 %s 题%s%s" % (html.escape(str(it.get("no"))),
                            (" · " + html.escape(str(it.get("kp")))) if show_meta else "",
                            " · %g 分" % float(it.get("full") or 0))
    b.append('<div class="qhead"><span class="qno2">%d.</span><span>%s</span></div>\n' % (idx, meta))
    b.append('<div class="stem2">%s</div>\n' % html.escape(str(it.get("stem") or "").strip()))
    opts = it.get("options") or []
    if opts:
        b.append('<div class="opts">%s</div>\n' % "".join("<span>%s</span>" % html.escape(str(o)) for o in opts))
        b.append('<div class="qhead" style="margin-top:2mm">答：<span class="ansline"></span></div>\n')
    else:
        for _ in range(blank_lines(it)):
            b.append('<div class="ln2"></div>\n')
    b.append("</div>\n")
    return "".join(b)


def shortfall_html(lines):
    if not lines:
        return ""
    b = ['<div class="warn"><b>说明（不是出错，是题不够）</b><ul>']
    for s in lines:
        b.append("<li>%s</li>" % html.escape(s))
    b.append("</ul></div>\n")
    return "".join(b)


# ============================================================ 标准逐题得分表（--scores）
# 和 insight.py 同一套口径（长表三种写法都认），但这里是独立实现，不 import 它、不改它。
SCORES_STU = ("学生", "姓名", "名字")
SCORES_NO = ("题号", "题目", "编号")
SCORES_GOT = ("得分", "分数", "成绩")
SCORES_FULL = ("满分", "分值")
SCORES_KP = ("知识点", "考点")
SCORES_SKIP = ("班级", "班")
SCORES_MAX_REPORT = 8


class ScoresError(Exception):
    """逐题得分表读不动。消息本身就是给老师看的中文，不要往上抛栈。"""


def _col_of(head, names):
    for i, h in enumerate(head):
        if str(h).strip() in names:
            return i
    return -1


def _num_or_none(v):
    s = str(v).strip() if v is not None else ""
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _no_label(no):
    s = str(no or "").strip()
    if s.startswith("第"):
        s = s[1:]
    if s.endswith("题"):
        s = s[:-1]
    return s.strip()


def _parse_scores(path, bank=None):
    """标准逐题得分表 -> 逐行原始记录 parsed=[{stu,no,kp,full,got}] + wide + warn。

    认三种（与 insight.py 一致）：
      1) 学生,题号,知识点,满分,得分   2) 学生,题号,得分（知识点/满分从题库补）
      3) 宽表：学生 + 每列一个题号
    读不动就抛 ScoresError（中文一句话），由调用方变成提示。
    """
    if not path or not str(path).strip():
        raise ScoresError("没给逐题得分表的路径（--scores 后面要跟一个 .csv/.xlsx 转出的 csv）")
    if not os.path.exists(path):
        raise ScoresError("找不到逐题得分表：%s" % path)
    try:
        with open(path, encoding="utf-8-sig", newline="") as f:
            raw_rows = [r for r in csv.reader(f) if any(str(c).strip() for c in r)]
    except UnicodeDecodeError:
        raise ScoresError("逐题得分表不是 UTF-8 编码，另存为 CSV（UTF-8）再试：%s" % path)
    except OSError as e:
        raise ScoresError("逐题得分表打不开（可能正被 Excel 占着）：%s（%s）" % (path, e))
    if not raw_rows:
        raise ScoresError("逐题得分表是空文件，一行都没有：%s" % path)
    head = [str(h).strip() for h in raw_rows[0]]
    body = raw_rows[1:]
    if not body:
        raise ScoresError("逐题得分表只有表头、没有数据行：%s\n表头是：%s" % (path, "、".join(head)))

    bank = bank if bank is not None else load_bank(None)
    items = bank.get("items") or []
    full_of = {}
    kp_of = {}
    for it in items:
        no = str(it.get("no"))
        full_of[no] = float(it.get("full") or 0)
        kp_of[no] = str(it.get("kp") or "")
    bank_nos = set(full_of)

    c_stu = _col_of(head, SCORES_STU)
    c_no = _col_of(head, SCORES_NO)
    c_got = _col_of(head, SCORES_GOT)
    c_full = _col_of(head, SCORES_FULL)
    c_kp = _col_of(head, SCORES_KP)
    if c_stu < 0:
        raise ScoresError("逐题得分表里找不到「学生」或「姓名」列。\n表头是：%s" % "、".join(head))
    wide = (c_no < 0 and c_got < 0 and c_full < 0)
    if not wide and (c_no < 0 or c_got < 0):
        raise ScoresError("逐题得分表里找不到「%s」列。\n表头是：%s"
                          % ("题号" if c_no < 0 else "得分", "、".join(head)))

    parsed, warn = [], []
    if wide:
        skip = set(SCORES_STU) | set(SCORES_SKIP)
        qcols = [(i, head[i]) for i in range(len(head)) if head[i] and head[i] not in skip]
        if not qcols:
            raise ScoresError("这份表看不出题号列（宽表要「姓名 + 每列一个题号」）。\n表头是：%s" % "、".join(head))
        for r in body:
            stu = str(r[c_stu]).strip() if c_stu < len(r) else ""
            if not stu:
                continue
            for i, col in qcols:
                no = _no_label(col)
                if not no:
                    continue
                got = _num_or_none(r[i]) if i < len(r) else None
                if got is None:
                    if str(r[i]).strip() if i < len(r) else "":
                        warn.append("题 %s 的单元格不是数字，按未作答记 0 分" % no)
                    got = 0.0
                full = full_of.get(no, 0.0) or 1.0
                parsed.append({"stu": stu, "no": no, "kp": kp_of.get(no, ""), "full": full, "got": got})
    else:
        for r in body:
            stu = str(r[c_stu]).strip() if c_stu < len(r) else ""
            no = _no_label(r[c_no]) if 0 <= c_no < len(r) else ""
            if not stu or not no:
                continue
            got = _num_or_none(r[c_got]) if c_got < len(r) else None
            if got is None:
                got = 0.0
            full = _num_or_none(r[c_full]) if 0 <= c_full < len(r) else None
            if not full:
                full = full_of.get(no, 0.0)
            if not full:
                warn.append("题 %s 在题库里查不到满分，按 1 分算" % no)
                full = 1.0
            kp = ""
            if 0 <= c_kp < len(r) and r[c_kp] is not None:
                kp = str(r[c_kp]).strip()
            parsed.append({"stu": stu, "no": no, "kp": kp or kp_of.get(no, ""), "full": full, "got": got})

    if not parsed:
        raise ScoresError("逐题得分表里没有一行能用的数据（学生/题号都是空的？）：%s" % path)

    over = [p for p in parsed if p["got"] > p["full"] + 1e-9]
    if over:
        lines = ["有 %d 行「得分」超过「满分」，先把表改对再来（review2 不替你改数）：" % len(over)]
        for p in over[:SCORES_MAX_REPORT]:
            lines.append("  %s 第 %s 题：得分 %g > 满分 %g" % (p["stu"], p["no"], p["got"], p["full"]))
        if len(over) > SCORES_MAX_REPORT:
            lines.append("  …… 还有 %d 行没列出来" % (len(over) - SCORES_MAX_REPORT))
        lines.append("（insight.py 的做法是按满分封顶继续算；review2 这版按任务书要求直接拦下。也可以先用 --rates 手给错误率。）")
        raise ScoresError("\n".join(lines))

    return parsed, wide, warn


def load_scores_table(path, bank=None):
    """标准逐题得分表 -> 每题错误率。

    口径：全班得分合计 / 满分合计（和 --db 同一条算式，同一份数据得同一串数）。
    读不动抛 ScoresError（中文一句话），调用方负责变成提示，不是栈。
    """
    parsed, wide, warn = _parse_scores(path, bank)
    bank = bank if bank is not None else load_bank(None)
    bank_nos = {str(it.get("no")) for it in (bank.get("items") or [])}
    by_no = {}
    for p in parsed:
        b = by_no.setdefault(p["no"], {"n": 0, "got": 0.0, "tot": 0.0, "full": 0.0, "n_right": 0})
        b["n"] += 1
        b["got"] += p["got"]
        b["tot"] += p["full"]
        b["full"] = max(b["full"], p["full"])
        if p["full"] and p["got"] >= p["full"] - 1e-9:
            b["n_right"] += 1

    rates, detail = {}, {}
    for no, b in by_no.items():
        # 与 load_questions(--db) 同一算式：全班得分合计 / 满分合计
        P = (b["got"] / b["tot"]) if b["tot"] else 0.0
        rates[no] = max(0.0, min(1.0, 1.0 - P))
        detail[no] = {"n": b["n"], "n_right": b["n_right"], "P": P}

    unmatched = sorted([n for n in rates if n not in bank_nos], key=lambda x: (len(x), x))
    qc = {"rows": len(parsed), "students": len({p["stu"] for p in parsed}),
          "items": len(rates), "mode": "宽表" if wide else "长表", "warn": warn[:6]}
    return {"rates": rates, "detail": detail, "unmatched": unmatched, "qc": qc,
            "label": "逐题得分 %s（%s ｜ %d 行 / %d 人 / %d 题）"
                     % (os.path.basename(path), qc["mode"], qc["rows"], qc["students"], qc["items"])}


# ============================================================ #14 讲评清单
REVIEW_WHY = {
    "概念不清": "把它和最容易混的那个概念摆一起对比（正例、反例各一个），再让学生用自己的话复述一遍",
    "计算错误": "别从头讲：拆成单步，让学生指出第一步就错在哪，只练这一步，连做三遍",
    "审题错误": "带学生圈题干里的条件词和问题词，把「已知 → 要求」写成两行再动笔",
    "步骤遗漏": "给一张步骤清单（每步打一个勾），按清单把这道题的过程重写一遍",
}
DIFF_CN = {"easy": "基础", "mid": "中等", "hard": "挑战"}
REVIEW_ASK = {
    "概念不清": "「{kp}」的定义是什么？不看笔记，用一句话说清楚，再说这道题为什么不能落到别的选项上。",
    "计算错误": "这道题第一步该先算什么？只写第一步给我看，说说为什么是它而不是别的。",
    "审题错误": "题干里哪些是条件、哪些是要求？把这两类各圈一遍给我看。",
    "步骤遗漏": "做这类题一共分几步？能说出五步才算过关，说漏了说明还没成套路。",
}
WHY_DEFAULT2 = "先让一个做错的学生讲他的思路，顺着他的路走一遍，再指出在哪一步拐错了弯"
ASK_DEFAULT2 = "让学生把这道题的关键条件复述一遍，再说第一步做什么。"


def why_text(it):
    return REVIEW_WHY.get(str(it.get("err") or "").strip(), WHY_DEFAULT2)


def ask_text(it):
    s = REVIEW_ASK.get(str(it.get("err") or "").strip(), ASK_DEFAULT2).format(kp=it.get("kp") or "")
    if it.get("options"):
        s += "再追一句：这道题的错误选项是怎么骗人的？"
    return s


def plan_schedule(ranked, budget=40, lo=3, hi=8, max_n=5, key="rate"):
    """从排头往下装，装到这节课装不下为止。返回 (能讲的, 剩下的)。
    key: 时长按哪个字段分配（全班版按 rate，个体版按 loss）。"""
    n = 0
    mins = []
    limit = max(1, min(int(max_n), len(ranked)))
    for k in range(1, limit + 1):
        cand = alloc_minutes(ranked[:k], budget=budget, lo=lo, hi=hi, key=key)
        if sum(cand) <= budget:
            n, mins = k, cand
        else:
            break
    if n == 0 and ranked:
        n, mins = 1, [min(hi, max(lo, budget))]
    return list(zip(ranked[:n], mins)), ranked[n:]


def _rates_from_args(args):
    """返回 (题号->错误率, 说明, 附加件)。取不到就 (None, 原因, None)。"""
    if getattr(args, "scores", None):
        try:
            res = load_scores_table(args.scores, load_bank(getattr(args, "bank", None)))
        except ScoresError as e:
            return None, str(e), None
        except Exception as e:                                # noqa: BLE001
            return None, "这份逐题得分表读不动（%s）" % e, None
        if not res["rates"]:
            return None, "逐题得分表里没有算出任何一道题的错误率", None
        return res["rates"], res["label"], res
    if getattr(args, "rates", None):
        out = {}
        for part in str(args.rates).replace("，", ",").split(","):
            part = part.strip()
            if not part:
                continue
            if ":" not in part and "：" not in part:
                return None, "命令行 --rates 要写成「题号:错误率」，用逗号隔开，例如 1:0.15,2:0.3"
            k, v = re.split(r"[:：]", part, 1)
            try:
                val = float(v)
            except ValueError:
                return None, "「%s」里的错误率不是数字" % part
            if val > 1:
                val = val / 100.0
            out[k.strip()] = max(0.0, min(1.0, val))
        if not out:
            return None, "命令行 --rates 里没解析出任何一道题", None
        return out, "命令行给的错误率表", None
    try:
        qs, meta = load_questions(args.src, args.db, args.rows)
    except Exception as e:                                    # noqa: BLE001
        return None, "读这次考试的得分失败（%s）" % e, None
    if not qs:
        return None, "这次考试没读到任何一道题的得分", None
    return {q["no"]: q["rate"] for q in qs}, meta["label"], None


def review2_html(scheduled, leftover, total_min, budget, usable, skipped, src_label):
    b = [_head2("这节课讲什么")]
    b.append("<h1>这节课讲什么</h1>\n")
    b.append('<div class="sub">%s ｜ 这节课 %d 题 · 合计 %d 分钟（%d 秒）｜ 上限 %d 分钟</div>\n'
             % (html.escape(src_label), len(scheduled), total_min, total_min * 60, budget))
    b.append('<div class="meta"><div>科目 / 班级：____________________</div><div>日期：______ 月 ______ 日</div></div>\n')
    b.append('<table class="plan">\n<thead><tr><th style="width:6%">#</th><th style="width:10%">题号</th>'
             '<th style="width:24%">知识点</th><th style="width:11%">错误率</th>'
             '<th style="width:14%">讲多久</th><th>怎么讲</th></tr></thead>\n<tbody>\n')
    for i, (it, m) in enumerate(scheduled, 1):
        b.append("<tr><td>%d</td><td>第 %s 题</td><td>%s</td><td>%.0f%%</td><td>%d 分钟（%d 秒）</td><td>%s</td></tr>\n"
                 % (i, html.escape(str(it.get("no"))), html.escape(str(it.get("kp"))),
                    it["rate"] * 100, m, m * 60, html.escape(why_text(it))))
    b.append("</tbody>\n</table>\n")

    b.append('<div class="sectitle">每题讲完，立刻问这一句</div>\n<ol class="asklist">\n')
    for it, _ in scheduled:
        b.append("<li><b>第 %s 题</b>（%s）「%s」<br>问：%s</li>\n"
                 % (html.escape(str(it.get("no"))), html.escape(str(it.get("kp"))),
                    html.escape(stem_head(it, 34)), html.escape(ask_text(it))))
    b.append("</ol>\n")

    b.append('<div class="sectitle">下次再讲（这节课装不下）</div>\n')
    if leftover:
        b.append('<ul class="asklist">\n')
        for it in leftover:
            b.append("<li>第 %s 题 · %s · 错误率 %.0f%% ｜「%s」</li>\n"
                     % (html.escape(str(it.get("no"))), html.escape(str(it.get("kp"))),
                        it["rate"] * 100, html.escape(stem_head(it, 26))))
        b.append("</ul>\n")
    else:
        b.append('<p class="mini">没有剩下的题，这节课能把该讲的讲完。</p>\n')

    b.append('<div class="foot">选题口径：只从题面已录入的 %d 道题里选（题库共 %d 道，%d 道题面待补，补齐后自动纳入）。'
             '时长按错误率分配，单题 3–8 分钟。</div>\n'
             % (len(usable), len(usable) + len(skipped), len(skipped)))
    b.append("</body>\n</html>\n")
    return "".join(b)


def cmd_review2(args):
    print("=" * 78)
    print("#14 讲评清单（这节课讲什么）")
    print("=" * 78)
    bank = load_bank(getattr(args, "bank", None))
    if bank["err"]:
        print(bank["err"])
        if bank.get("hint"):
            print(bank["hint"])
        print("没有题库就选不了题，这次不出清单。补好题库再跑一次即可。")
        return 2
    usable, skipped = split_bank(bank["items"])
    print("题库: %s" % bank.get("path"))
    print("  共 %d 道，有题干 %d 道，题干为空 %d 道（空题一律不进选题池）"
          % (len(bank["items"]), len(usable), len(skipped)))
    if skipped:
        print("  被排除的题号: %s" % "、".join(str(i.get("no")) for i in skipped))
    if not usable:
        print()
        print("题库里还没有一道题录了题干，讲评清单出不来（不会印空题占位）。")
        print("下一步：把题面填进 题库.json 的 stem 字段，再跑一次 review2。")
        return 2

    rates, why, extra = _rates_from_args(args)
    if not rates:
        print()
        if getattr(args, "scores", None):
            print("--scores 这份表用不了：")
            for line in str(why).splitlines():
                print("  " + line)
        else:
            print("拿不到这次考试各题的错误率：%s" % why)
        print("四种给法任选一种，然后重跑：")
        print("  python teaching.py review2 --scores <逐题得分.csv>")
        print("  python teaching.py review2 --db <班级库.db>")
        print("  python teaching.py review2 --csv <卡导出.csv> --paper <卷子结构.json>")
        print('  python teaching.py review2 --rates "1:0.15,2:0.3"')
        return 2

    got = []
    for it in usable:
        r = rates.get(str(it.get("no")))
        if r is None:
            continue
        d = dict(it)
        d["rate"] = float(r)
        got.append(d)
    print("错误率来源: %s" % why)
    if extra:
        for w in extra["qc"]["warn"]:
            print("  ! %s" % w)
        print("  题库里对不上：%s" % ("、".join(extra["unmatched"]) if extra["unmatched"] else "无"))
        miss = [str(i.get("no")) for i in usable if str(i.get("no")) not in rates]
        if miss:
            if len(miss) <= 8:
                print("  题库里有题干、这次没有得分（不进清单）：%s" % "、".join(miss))
            else:
                print("  题库里有题干、这次没有得分（不进清单）：共 %d 道，不一一列出" % len(miss))
    print("  题库里能对上号的题: %d 道" % len(got))
    if not got:
        print()
        print("题库的题号和这次考试的题号对不上，一道都算不出来，没有清单可出。")
        print("  这次考试的题号: %s" % "、".join(sorted(rates, key=lambda x: (len(x), x))[:30]))
        print("  题库里的题号  : %s" % "、".join(str(i.get("no")) for i in usable))
        return 2

    ranked = sorted(got, key=lambda x: (-x["rate"], -float(x.get("full") or 0), _no_key(x)))
    scheduled, leftover = plan_schedule(ranked, args.budget, 3, 8, args.top)
    total_min = sum(m for _, m in scheduled)

    print()
    print("%-3s %-5s %-18s %-8s %-14s %s" % ("#", "题号", "知识点", "错误率", "讲多久", "讲什么"))
    print("-" * 78)
    for i, (it, m) in enumerate(scheduled, 1):
        print("%-3d %-5s %-18s %-8s %-14s %s"
              % (i, it.get("no"), it.get("kp"), "%.1f%%" % (it["rate"] * 100),
                 "%d 分 / %d 秒" % (m, m * 60), why_text(it)))
    print("-" * 78)
    print("合计: %d 分钟 = %d 秒 ｜ 上限 %d 分钟 ｜ 富余 %d 分钟"
          % (total_min, total_min * 60, args.budget, args.budget - total_min))
    print()
    print("讲完立刻问什么:")
    for it, _ in scheduled:
        print("  第 %s 题「%s」" % (it.get("no"), stem_head(it, 20)))
        print("        问: %s" % ask_text(it))
    print()
    print("下次再讲（这节课装不下，按错误率排序）:")
    if leftover:
        for it in leftover:
            print("  第 %-3s 题 · %-16s · 错误率 %.0f%%" % (it.get("no"), it.get("kp"), it["rate"] * 100))
    else:
        print("  没有剩下的题，这节课能把该讲的讲完。")

    path, size = _write_html_out("这节课讲什么.html",
                                review2_html(scheduled, leftover, total_min, args.budget, usable, skipped, why))
    print()
    print("打印页: %s (%d 字节)" % (path, size))
    return 0


# ============================================================ #27 课前前置检测
def cmd_prelesson(args):
    print("=" * 78)
    print("#27 课前 3 分钟前置检测")
    print("=" * 78)
    bank = load_bank(getattr(args, "bank", None))
    if bank["err"]:
        print(bank["err"])
        if bank.get("hint"):
            print(bank["hint"])
        print("没有题库就出不了检测卷，这次不生成。")
        return 2
    usable, skipped = split_bank(bank["items"])
    print("题库: %s ｜ 共 %d 道，有题干 %d 道，题干为空 %d 道（空题一律不用）"
          % (bank.get("path"), len(bank["items"]), len(usable), len(skipped)))
    if not usable:
        print()
        print("题库里还没有一道题录了题干，检测卷出不来（不会印空题占位）。")
        print("下一步：把题面填进 题库.json 的 stem 字段，再跑一次。")
        return 2

    node = find_node(args.kp)
    if not node:
        print()
        print("知识点树里没有「%s」这个知识点，也没匹配到别名。" % args.kp)
        print("可以这样查（只列本节建模过的）：")
        print("  " + "、".join(n["name"] for n in NODES if n["chapter"] == "rat"))
        return 2

    nid = node["id"]
    print()
    print("本课知识点: %s（%s）" % (node["name"], nid))
    layers = prereq_layers(nid)
    if not layers:
        print("前置链: 无 —— 它是本章起点，没有前置知识点。")
        print()
        print("所以这节课没有「前置」可测，检测卷不生成。两个办法：")
        print("  1. 换一个更靠后的知识点再跑，例如：python teaching.py prelesson --kp 有理数混合运算")
        print("  2. 若确实要测本课，直接在题库里挑这个知识点的题手动出卷")
        return 2

    by_node = bank_index(usable)
    print("前置链（由近及远，共 %d 层）:" % len(layers))
    for i, lay in enumerate(layers, 1):
        parts = []
        for x in lay:
            k = len(by_node.get(x, []))
            parts.append("%s%s" % (NODE_BY_ID[x]["name"], "（可出 %d 道）" % k if k else "（题库里没有可用题）"))
        print("  第 %d 层: %s" % (i, "、".join(parts)))

    order = [x for lay in layers for x in lay]
    with_q = [x for x in order if by_node.get(x)]
    if not with_q:
        print()
        print("这个知识点往前查了一整条链，题库里一道可用题都没有，检测卷出不来。")
        print("建议去题库补这几条（按由近及远）：")
        plan = []
        for x in order[:4]:
            same = [i for i in skipped if bank_node(i) == x]
            plan.append("  · %s%s" % (NODE_BY_ID[x]["name"],
                                      "：题库第 %s 题有记录但题干为空，录上即可" % "、".join(str(i.get("no")) for i in same)
                                      if same else "：题库里还没有这个知识点的题"))
        print("\n".join(plan))
        return 2

    primary = with_q[:2]
    picked = []
    used = set()
    notes = []
    for x in primary:
        cnt = 0
        for it in by_node[x]:
            if cnt >= 2:
                break
            if str(it.get("no")) in used:
                continue
            used.add(str(it.get("no")))
            picked.append((x, it))
            cnt += 1
        if cnt < 2:
            same = [i for i in skipped if bank_node(i) == x]
            notes.append("「%s」只凑到 %d 道（题库里可用题就这么多）%s"
                         % (NODE_BY_ID[x]["name"], cnt,
                            ("；第 %s 题题干为空，录上就能补" % "、".join(str(i.get("no")) for i in same)) if same else ""))
    from_primary = len(picked)
    if len(picked) < 4:
        topped = 0
        for x in with_q[2:]:
            if len(picked) >= 4:
                break
            for it in by_node[x]:
                if len(picked) >= 4:
                    break
                if str(it.get("no")) in used:
                    continue
                used.add(str(it.get("no")))
                picked.append((x, it))
                topped += 1
        if topped:
            notes.append("最近两个有题的前置知识点只出得了 %d 道，按由近及远往后又补了 %d 道。" % (from_primary, topped))
        if len(picked) < 4:
            notes.append("整条前置链只凑到 %d 道有题干的题（目标 4 道），能出几道出几道。" % len(picked))

    print()
    print("选中 %d 道（最近两个有题的前置知识点各 2 道）:" % len(picked))
    for i, (x, it) in enumerate(picked, 1):
        print("  %d. 第 %s 题 ｜ %-16s ｜ %s ｜「%s」" % (
            i, it.get("no"), NODE_BY_ID[x]["name"], DIFF_CN.get(diff_of(it), diff_of(it)), stem_head(it, 20)))
    total_full = sum(float(it.get("full") or 0) for _, it in picked)
    print("满分 %g 分 ｜ 建议限时 3 分钟" % total_full)

    # 答案
    print()
    print("答案（老师版）:")
    for i, (x, it) in enumerate(picked, 1):
        print("  %d. 第 %s 题 → %s" % (i, it.get("no"), it.get("answer") or "（题库里没填答案）"))

    path, size = _write_html_out("前置检测_%s.html" % node["name"],
                                prelesson_html(node, picked, notes, total_full))
    print()
    if notes:
        print("降级说明（不是出错）:")
        for n in notes:
            print("  · " + n)
    print("打印页: %s (%d 字节)" % (path, size))
    return 0


def prelesson_html(node, picked, notes, total_full):
    b = [_head2("课前 3 分钟前置检测")]
    b.append("<h1>课前 3 分钟 · 前置检测</h1>\n")
    b.append('<div class="sub">本课要讲：%s ｜ %d 道题 ｜ 限时 3 分钟 ｜ 满分 %g 分</div>\n'
             % (html.escape(node["name"]), len(picked), total_full))
    b.append('<div class="meta"><div>班级：____________　姓名：____________</div><div>得分：______ / %g</div></div>\n'
             % total_full)
    b.append(shortfall_html(notes))
    for i, (x, it) in enumerate(picked, 1):
        b.append(q_html(it, i))
        b.append('<div class="mini">来源前置知识点：%s</div>\n' % html.escape(NODE_BY_ID[x]["name"]))
    b.append('<div class="keybox">\n<b>答案（老师版 · 发给学生前请沿虚线剪下）</b>\n')
    b.append('<table class="keytable">\n<tr><th style="width:10%">题号</th><th style="width:18%">原题号</th>'
             '<th style="width:22%">前置知识点</th><th>答案</th></tr>\n')
    for i, (x, it) in enumerate(picked, 1):
        b.append("<tr><td>%d</td><td>第 %s 题</td><td>%s</td><td>%s</td></tr>\n"
                 % (i, html.escape(str(it.get("no"))), html.escape(NODE_BY_ID[x]["name"]),
                    html.escape(str(it.get("answer") or "题库未填答案"))))
    b.append("</table>\n</div>\n")
    b.append('<div class="foot">选题口径：只从题面已录入的题里取，按前置链由近及远挑；'
             '拿不出 4 道就在本页顶部说明还差几道。</div>\n')
    b.append("</body>\n</html>\n")
    return "".join(b)


# ============================================================ #28 分层作业三套
def cmd_layers3(args):
    print("=" * 78)
    print("#28 分层作业三套（A 基础 / B 提高 / C 挑战）")
    print("=" * 78)
    bank = load_bank(getattr(args, "bank", None))
    if bank["err"]:
        print(bank["err"])
        if bank.get("hint"):
            print(bank["hint"])
        print("没有题库就出不了作业卷，这次不生成。")
        return 2
    usable, skipped = split_bank(bank["items"])
    print("题库: %s ｜ 共 %d 道，有题干 %d 道，题干为空 %d 道（空题一律不进卷）"
          % (bank.get("path"), len(bank["items"]), len(usable), len(skipped)))
    if not usable:
        print()
        print("题库里还没有一道题录了题干，作业卷出不来（不会印空题占位）。")
        print("下一步：把题面填进 题库.json 的 stem 字段，再跑一次 layers3。")
        return 2

    per = max(1, int(args.per))
    kp_label = "全题库"
    if args.kp:
        node = find_node(args.kp)
        if not node:
            print()
            print("知识点树里没有「%s」这个知识点，也没匹配到别名。" % args.kp)
            print("可以这样查：" + "、".join(n["name"] for n in NODES if n["chapter"] == "rat"))
            return 2
        pool = [it for it in usable if bank_node(it) == node["id"]]
        kp_label = node["name"]
        if not pool:
            print()
            print("题库里「%s」暂时没有可用题（有题干的题一道都没有）。" % node["name"])
            same = [i for i in skipped if bank_node(i) == node["id"]]
            if same:
                print("  题库第 %s 题属于这个知识点，但题干为空，录上就能用。"
                      % "、".join(str(i.get("no")) for i in same))
            else:
                print("  题库里还没有这个知识点的题，先去题库补录。")
            print("  想看整本题库的分层：python teaching.py layers3 --per %d" % per)
            return 2
    else:
        pool = list(usable)

    print("选题范围: %s（可用题 %d 道）｜ 每层目标 %d 道" % (kp_label, len(pool), per))
    print()
    chosen = {}
    notes = []
    odd = [it for it in pool if str(it.get("difficulty") or "").strip().lower() not in DIFF_ORDER]
    if odd:
        notes.append("有 %d 道题的难度标得不是 easy/mid/hard，暂按「中等」处理：第 %s 题。"
                     % (len(odd), "、".join(str(i.get("no")) for i in odd)))
    for d, letter, name, factor in LAYER_PAGE:
        cand = sorted([it for it in pool if diff_of(it) == d], key=lambda x: (_no_key(x), str(x.get("no"))))
        take = cand[:per]
        chosen[d] = take
        short = per - len(take)
        if short > 0:
            notes.append("%s 卷（%s）差 %d 道：这个难度只有 %d 道可用题，要 %d 道。"
                         % (letter, name, short, len(take), per))
        print("%s 卷（%s）: %d 道%s" % (letter, name, len(take),
                                      "" if short <= 0 else "  ← 差 %d 道" % short))
        if take:
            for i, it in enumerate(take, 1):
                print("   %d. 第 %s 题 ｜ %-16s ｜ %g 分 ｜「%s」" % (
                    i, it.get("no"), it.get("kp"), float(it.get("full") or 0), stem_head(it, 22)))
        else:
            print("   （这一层一道都没有）")

    if any(n[:2] in ("A ", "B ", "C ") for n in notes):
        if args.kp:
            notes.append("想看更宽的卷子：python teaching.py layers3 --per %d（不限知识点）" % per)
        else:
            notes.append("想只看某个知识点：python teaching.py layers3 --kp 绝对值 --per %d" % per)
    total = sum(len(v) for v in chosen.values())
    if total == 0:
        print()
        print("三层都没选出题来，作业卷不生成。")
        return 2

    path, size = _write_html_out("分层作业_%s.html" % kp_label, layers3_html(kp_label, chosen, notes))
    print()
    if notes:
        print("降级说明（不是出错）:")
        for n in notes:
            print("  · " + n)
    print("合计 %d 道题：三个 A4 页（A/B/C 各自分页）+ 第 4 页答案与出处。" % total)
    print("打印页: %s (%d 字节)" % (path, size))
    return 0


def layers3_html(kp_label, chosen, notes):
    b = [_head2("分层作业 · %s" % kp_label)]
    first = True
    for d, letter, name, factor in LAYER_PAGE:
        items = chosen.get(d) or []
        if not first:
            b.append('<div class="pgbreak"></div>\n')
        first = False
        b.append("<h1>%s %s卷</h1>\n" % (letter, "基础" if d == "easy" else ("提高" if d == "mid" else "挑战")))
        mins = sum(est_minutes(float(it.get("full") or 0), factor) for it in items)
        full = sum(float(it.get("full") or 0) for it in items)
        b.append('<div class="sub">%s ｜ %d 道题 ｜ 建议 %d 分钟 ｜ 满分 %g 分</div>\n'
                 % (html.escape(kp_label), len(items), mins, full))
        b.append('<div class="meta"><div>班级：____________　姓名：____________</div>'
                 '<div>用时：______ 分　得分：______ / %g</div></div>\n' % full)
        if not items:
            b.append('<div class="warn">这一层这次一道题都没选出来 —— 题库里这个难度的可用题是 0 道。'
                     '老师可以把本层留空，或先补题库。</div>\n')
        else:
            for i, it in enumerate(items, 1):
                b.append(q_html(it, i))
        b.append('<div class="foot">%s卷 · 共 %d 题。只从题面已录入的题里选；题不够时本页会写明差几道，不会凑空题。</div>\n'
                 % (letter, len(items)))
    b.append('<div class="pgbreak"></div>\n')
    b.append("<h1>答案与出处</h1>\n")
    b.append('<div class="sub">老师版 · 发给学生前请撕下本页</div>\n')
    b.append('<table class="keytable">\n<tr><th style="width:8%">卷</th><th style="width:10%">题号</th>'
             '<th style="width:22%">知识点</th><th style="width:11%">原题号</th><th>答案</th></tr>\n')
    for d, letter, name, factor in LAYER_PAGE:
        for i, it in enumerate(chosen.get(d) or [], 1):
            b.append("<tr><td>%s-%d</td><td>第 %s 题</td><td>%s</td><td>%s</td><td>%s</td></tr>\n"
                     % (letter, i, i, html.escape(str(it.get("kp"))), html.escape(str(it.get("no"))),
                        html.escape(str(it.get("answer") or "题库未填答案"))))
    b.append("</table>\n")
    b.append(shortfall_html(notes))
    b.append("</body>\n</html>\n")
    return "".join(b)


# ============================================================ 入口
def cmd_all(args):
    cmd_tree(argparse.Namespace(query=None))
    print()
    cmd_layer(args)
    print()
    cmd_review(args)
    print()
    return cmd_print(args)


def _common(p, suppress=False):
    """公共参数。suppress=True 时用 SUPPRESS 当默认值，
    这样「review --top 8」与「--top 8 review」两种写法都能用，且互不覆盖。"""
    def d(v):
        return argparse.SUPPRESS if suppress else v
    p.add_argument("--src", default=d("db"), choices=["db", "rows"], help="逐题数据来源（默认 db）")
    p.add_argument("--db", default=d(os.path.join(TMP_DIR, "class_demo.db")))
    p.add_argument("--rows", default=d(os.path.join(TMP_DIR, "class_rows.json")))
    p.add_argument("--top", type=int, default=d(5), help="讲评题数（默认 5）")
    p.add_argument("--budget", type=int, default=d(40), help="讲评总时长上限（默认 40 分钟）")
    p.add_argument("--time-by", default=d("rate"), choices=["rate", "loss"], help="时长分配依据")
    p.add_argument("--min-rate", type=float, default=d(0.5), help="错题卷错误率阈值")
    p.add_argument("--quiz-top", type=int, default=d(8), help="错题卷最多收录题数（默认 8）")
    p.add_argument("--seed", type=int, default=d(20261028), help="口算卡随机种子")
    return p


def build_parser():
    p = argparse.ArgumentParser(description="教学产出工具（纯标准库，0 次大模型调用）")
    _common(p)
    sub = p.add_subparsers(dest="cmd")
    t = sub.add_parser("tree", help="知识点树 + 前置链")
    t.add_argument("--query", default=None, help="查某个知识点的前置链")
    for name, h in (("layer", "分层作业包"), ("review", "讲评清单"),
                    ("print", "打印中心（3 个 HTML）"), ("all", "依次跑 tree/layer/review/print")):
        sp = sub.add_parser(name, help=h)
        _common(sp, suppress=True)
        if name == "review":
            sp.add_argument("--student", default=None, help="切出个体版：只看这名学生该讲哪几题")
            sp.add_argument("--scores", default=None,
                            help="逐题得分表（学生,题号,知识点,满分,得分）；个体版必需")
            sp.add_argument("--xlsx", default=None,
                            help="真实成绩单；没有逐题得分时用它给出明确说明")
            sp.add_argument("--bank", default=None)
    i = sub.add_parser("import", help="答题卡导入 + 讲评 + 分层")
    _common(i, suppress=True)
    i.add_argument("--csv", required=True)
    i.add_argument("--paper", required=True)
    # ---- 读题库的三条新命令（#14 / #27 / #28）----
    r2 = sub.add_parser("review2", help="#14 讲评清单（按错误率，含「讲完立刻问什么」）")
    _common(r2, suppress=True)
    r2.add_argument("--bank", default=None, help="题库路径（默认 教师工具\题库.json）")
    r2.add_argument("--rates", default=None, help='错误率表，例如 "1:0.15,2:0.3"')
    r2.add_argument("--scores", default=None,
                    help="标准逐题得分表（学生,题号,知识点,满分,得分），给了就用它算错误率")
    pl = sub.add_parser("prelesson", help="#27 课前 3 分钟前置检测")
    _common(pl, suppress=True)
    pl.add_argument("--kp", required=True, help="本课要讲的知识点")
    pl.add_argument("--bank", default=None)
    gp = sub.add_parser("gaps", help="集体补漏清单：按知识点聚合失分（吃逐题得分）")
    _common(gp, suppress=True)
    gp.add_argument("--scores", nargs="+", default=None, help="一份或多份逐题得分表")
    gp.add_argument("--xlsx", default=None, help="真实成绩单；没有逐题得分时用它给出明确说明")
    gp.add_argument("--bank", default=None)
    l3 = sub.add_parser("layers3", help="#28 分层作业三套（A/B/C）")
    _common(l3, suppress=True)
    l3.add_argument("--kp", default=None, help="限定知识点；不给就覆盖全题库")
    l3.add_argument("--per", type=int, default=6, help="每层取几道（默认 6）")
    l3.add_argument("--bank", default=None)
    return p


def main(argv=None):
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8")
        except Exception:
            pass
    args = build_parser().parse_args(argv)
    if not args.cmd:
        build_parser().print_help()
        return 1
    if args.cmd == "import":
        return cmd_import(args)
    return {"tree": cmd_tree, "layer": cmd_layer, "review": cmd_review,
            "print": cmd_print, "all": cmd_all,
            "review2": cmd_review2, "prelesson": cmd_prelesson,
            "layers3": cmd_layers3, "gaps": cmd_gaps}[args.cmd](args)


if __name__ == "__main__":
    sys.exit(main())
