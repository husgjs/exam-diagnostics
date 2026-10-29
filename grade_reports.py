# -*- coding: utf-8 -*-
r"""grade_reports.py — 教师侧报告工具（功能 3 / 5 / 6）

纯 Python + openpyxl，无第三方依赖，0 次大模型调用，只在本工作区内读写。

子命令
------
  reports        成绩单批量生成：明细表（一人一行、可筛选）+ 打印版（每页 4 人，A4 竖版）
  trend          多次考试总分趋势：折线数据 / 波动幅度 / 大起大落标记
  items archive  命题质量归档：算每题 P(难度)、D(区分度)、错误率，追加进 _命题档案.json
  items compare  命题质量历次对比：同一题号/知识点 P 值变化，标出「连续三次都太容易」
  items seed-demo 用模拟库造三次模拟考试归档（只为验证「连续三次」判定逻辑）

功能 3 用法
----------
  python grade_reports.py reports
  python grade_reports.py reports --src "D:\新建文件夹 (2)\九年级月考成绩单.xlsx" --out-dir "教师工具\输出"

功能 5 用法（多文件写法：把每次考试的 xlsx 依次列在后面，先旧后新）
-----------
  # 只有一次考试 —— 照常能跑，只出单点，不做趋势判断
  python grade_reports.py trend --files "D:\新建文件夹 (2)\九年级月考成绩单.xlsx"

  # 有多次考试 —— 文件顺序即时间顺序，人数/班级可以不同，按「姓名」对齐
  python grade_reports.py trend --files "10月月考.xlsx" "11月月考.xlsx" "12月月考.xlsx"
  python grade_reports.py trend --files "10月.xlsx" "11月.xlsx" --labels 十月 十一月

  表结构要求：每个 xlsx 的第一张含「姓名」的 sheet，表头必须有「姓名」和「总分」
  （必须是精确的「总分」，不是「小总分」；本工具用精确匹配，见 col_index）。

功能 6 真实数据接入方式
----------------------
  归档需要「每题得分表」，两种格式都支持，任选其一：

  1) CSV / TXT（推荐，从阅卷系统直接导出）—— 表头需含这 5 列：
         学生,题号,知识点,满分,得分
         张三,第1题,相反数,3,3
         张三,第2题,负数概念,3,0
         李四,第1题,相反数,3,0
     python grade_reports.py items archive --csv 每题得分.csv --exam "2026春·第一次月考" --subject 数学

  2) SQLite 库（本仓库 _tmp/class_demo.db 就是这种结构，questions 表）：
         student_id, ask, kp, full_score, got_score
     python grade_reports.py items archive --db "_tmp\class_demo.db" --exam "2026春·第一次月考" --subject 数学

  P / D 定义：P = 该题平均得分率；D = 高分组(前 27%)平均得分率 − 低分组(后 27%)平均得分率。
  分组依据是该生「本次这张卷」的总得分率。
"""

import argparse
import csv
import hashlib
import json
import os
import sqlite3
import sys
import time
from datetime import datetime

try:
    if (sys.stdout.encoding or "").lower().replace("-", "") not in ("utf8", "utf_8"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

import openpyxl
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.properties import PageSetupProperties

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
DEFAULT_SRC = os.path.join(ROOT, "九年级月考成绩单.xlsx")
DEFAULT_OUT = os.path.join(HERE, "输出")
ARCHIVE = os.path.join(HERE, "_命题档案.json")

# (学科, 表头名, 满分)
SUBJECTS = [
    ("语文", "语文分数", 100.0),
    ("数学", "数学分数", 150.0),
    ("英语", "英语分数", 150.0),
    ("物理", "物理分数", 100.0),
    ("化学", "化学分数", 100.0),
]
TOTAL_FULL = sum(s[2] for s in SUBJECTS)          # 600
PER_PAGE = 4                                       # 打印版每页人数
BLOCK_ROWS = 8                                     # 打印版每人占 8 行

THIN = Side(style="thin", color="999999")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="DDE7F3")
BAND_FILL = PatternFill("solid", fgColor="F2F6FB")


# --------------------------------------------------------------------------- #
# 读表：精确定位列名
# --------------------------------------------------------------------------- #
def col_index(header, name):
    """精确匹配表头，返回 0 基列号。

    坑：本成绩表同时有「小总分」（语数英小计）和「总分」（五科合计）。
    任何 in / startswith / 模糊匹配只要查「总分」都会先命中「小总分」，
    导致总分列全体错位。所以这里只认 strip 后完全相等。
    """
    hdr = [(str(c).strip() if c is not None else "") for c in header]
    if name not in hdr:
        near = [h for h in hdr if name and name in h]
        raise KeyError("表头里找不到精确列「%s」；包含它的列有 %s" % (name, near))
    if hdr.count(name) > 1:
        raise KeyError("表头里「%s」出现了 %d 次，无法确定用哪一列" % (name, hdr.count(name)))
    return hdr.index(name)


def load_scores(path):
    """读一份成绩表，返回 (records, meta)。records 每项含 班级/姓名/五科/总分。"""
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    wb = openpyxl.load_workbook(path, data_only=True)
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        hdr_i = None
        for i, row in enumerate(rows[:10]):
            vals = [(str(c).strip() if c is not None else "") for c in row]
            if "姓名" in vals and "总分" in vals:
                hdr_i = i
                break
        if hdr_i is None:
            continue
        hdr = rows[hdr_i]
        i_name = col_index(hdr, "姓名")
        i_total = col_index(hdr, "总分")            # 精确匹配，避开「小总分」
        i_cls = None
        try:
            i_cls = col_index(hdr, "班级")
        except KeyError:
            pass
        subj_idx = []
        for sname, hname, _full in SUBJECTS:
            try:
                subj_idx.append((sname, col_index(hdr, hname)))
            except KeyError:
                subj_idx.append((sname, None))
        i_rank = None
        for cand in ("总名次", "年级名次"):
            try:
                i_rank = col_index(hdr, cand)
                break
            except KeyError:
                continue

        recs = []
        for row in rows[hdr_i + 1:]:
            name = row[i_name] if i_name < len(row) else None
            if name is None or str(name).strip() == "":
                continue
            rec = {
                "班级": (str(row[i_cls]).strip() if i_cls is not None and row[i_cls] is not None else "未分班"),
                "姓名": str(name).strip(),
                "总分": float(row[i_total] or 0),
            }
            for sname, idx in subj_idx:
                rec[sname] = float(row[idx] or 0) if idx is not None else 0.0
            if i_rank is not None and i_rank < len(row) and row[i_rank] is not None:
                try:
                    rec["源名次"] = int(row[i_rank])
                except (TypeError, ValueError):
                    pass
            recs.append(rec)
        meta = {"file": os.path.abspath(path), "sheet": ws.title, "rows": len(recs), "header": [str(c) for c in hdr if c is not None]}
        return recs, meta
    raise KeyError("在 %s 里找不到同时含「姓名」「总分」表头的 sheet" % path)


# --------------------------------------------------------------------------- #
# 统计
# --------------------------------------------------------------------------- #
def rank_positions(recs, idxs):
    """把 idxs 这群人按「总分降序 + 原表行序」排出 1..k 的位次。

    用位次而不是并列名次，是为了和学校公布的总名次列完全对齐：
    本表 367 行里有 89 行总分与他人并列，学校是按行序打断并列的，
    照它的排法排，打印出去的名次才和学校发的那张表一模一样。
    """
    order = sorted(idxs, key=lambda i: (-round(recs[i]["总分"], 2), i))
    return {i: pos for pos, i in enumerate(order, 1)}


def enrich(recs, min_group=3):
    """补 年级名次 / 班级名次 / 各科得分率 / 最强最弱 / 结论。"""
    totals = [r["总分"] for r in recs]
    grade_rank = rank_positions(recs, range(len(recs)))
    recs_grade = grade_rank
    by_class = {}
    for i, r in enumerate(recs):
        by_class.setdefault(r["班级"], []).append(i)
    class_rank = [0] * len(recs)
    for cls, idxs in by_class.items():
        for i, pos in rank_positions(recs, idxs).items():
            class_rank[i] = pos

    n = len(recs)
    # 名次是群体口径：全表有效人数 < min_group 时，结论里也不排名次（与明细两列、打印版同一条规则）
    n_valid = sum(1 for r in recs if r["总分"] > 0)
    rank_note = None
    if n_valid < min_group:
        rank_note = ("样本 %d 人，少于群组下限 %d 人：不排名次，只列分数与强弱科。"
                     % (n_valid, min_group))
    # 总分 0 = 缺考/未录入。算任何均分之前先剔除，否则几个缺考就能把年级均分拉下去。
    # 口径要和 grade_deep / grade_analysis / grade_tools / teacher_app 一致。
    counted = [r for r in recs if r["总分"] > 0] or recs
    n_count = len(counted)
    avg = {"__年级__": sum(r["总分"] for r in counted) / n_count if n_count else 0.0,
           "__计数__": n_count, "__剔除__": n - n_count}
    cls_sizes = {c: len(v) for c, v in by_class.items()}
    for sname, _h, _f in SUBJECTS:
        avg[sname] = sum(r[sname] for r in counted) / n_count if n_count else 0.0

    for i, r in enumerate(recs):
        r["算出名次"] = recs_grade[i]
        r["年级名次"] = r.get("源名次", recs_grade[i])   # 学校公布的总名次优先
        r["班级名次"] = class_rank[i]
        r["年级人数"] = n
        r["班级人数"] = cls_sizes[r["班级"]]
        for sname, _h, full in SUBJECTS:
            r[sname + "_均差"] = round(r[sname] - avg[sname], 1)
            r[sname + "_得分率"] = (r[sname] / full) if full else 0.0
        r["结论"] = conclusion(r, avg, n, rank_note)
        if r["总分"] <= 0:
            r["最强科"], r["最弱科"] = "—", "—"      # 缺考的人没有强弱科可言
        else:
            rates = {s[0]: r[s[0] + "_得分率"] for s in SUBJECTS}
            best = max(rates, key=lambda k: (rates[k], -[s[0] for s in SUBJECTS].index(k)))
            worst = min(rates, key=lambda k: (rates[k], [s[0] for s in SUBJECTS].index(k)))
            r["最强科"], r["最弱科"] = best, worst
    return avg


def conclusion(r, avg, n, rank_note=None):
    """按最强/最弱科自动生成一句结论。"""
    if r["总分"] <= 0:
        return "本次五科均为 0 分，疑似缺考，请核实分数是否已录入；本条按缺考保留，不计入年级均分。"
    rates = {s[0]: r[s[0] + "_得分率"] for s in SUBJECTS}
    best = max(rates, key=lambda k: rates[k])
    worst = min(rates, key=lambda k: rates[k])
    gap = rates[best] - rates[worst]
    head = rank_note or ("年级第%d名/%d人，班级第%d名/%d人。"
                         % (r["年级名次"], n, r["班级名次"], r["班级人数"]))
    if gap >= 0.15:
        body = "优势在{0}（得分率{1:.0%}，比年级均分{2:+.1f}分），短板在{3}（{4:.0%}，{5:+.1f}分）。提分优先补{3}，{0}保持。".format(
            best, rates[best], r[best + "_均差"], worst, rates[worst], r[worst + "_均差"])
    elif gap >= 0.06:
        body = "最强{0}（{1:.0%}）与最弱{2}（{3:.0%}）差{4:.0f}个百分点，整体偏稳、短板可补。".format(
            best, rates[best], worst, rates[worst], gap * 100)
    else:
        body = "五科发展均衡：最强{0}（{1:.0%}）与最弱{2}（{3:.0%}）只差{4:.0f}个百分点，主要差距在整体水平而非偏科。".format(
            best, rates[best], worst, rates[worst], gap * 100)
    tail = "总分%s，%s年级均分。" % (
        ("%g" % r["总分"]), ("高于" if r["总分"] >= avg["__年级__"] else "低于"))
    return head + body + tail


# --------------------------------------------------------------------------- #
# 功能 3：成绩单批量生成
# --------------------------------------------------------------------------- #
DETAIL_COLS = (
    ["班级", "姓名"]
    + [s[0] for s in SUBJECTS]
    + ["总分", "年级名次", "班级名次"]
    + ["%s-年级均差" % s[0] for s in SUBJECTS]
    + ["最强科", "最弱科", "结论"]
)


def build_detail(recs, path, rank_note=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "成绩单明细"
    ws.append(DETAIL_COLS)
    for c in range(1, len(DETAIL_COLS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True)
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX
    for r in recs:
        row = [r["班级"], r["姓名"]]
        row += [r[s[0]] for s in SUBJECTS]
        # 小样本时名次列写说明：n=2 的「年级第 1 名/共 2 人」字面为真但会误导
        row += [r["总分"], rank_note if rank_note else r["年级名次"],
                rank_note if rank_note else r["班级名次"]]
        row += [r[s[0] + "_均差"] for s in SUBJECTS]
        row += [r["最强科"], r["最弱科"], r["结论"]]
        ws.append(row)

    last = ws.max_row
    widths = [10, 9] + [7] * 5 + [8, 8, 8] + [11] * 5 + [8, 8, 62]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(DETAIL_COLS)), last)
    for row in ws.iter_rows(min_row=2, max_row=last, max_col=len(DETAIL_COLS)):
        for cell in row:
            cell.border = BOX
            if cell.column_letter in ("R",):
                cell.alignment = Alignment(vertical="center", wrap_text=True)

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:1"
    wb.save(path)
    return path, ws.title, len(DETAIL_COLS)


def build_print(recs, path, rank_note=None):
    """打印版：每人一个 8 行分块、每页 4 人、A4 竖版、块间显式分页。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "打印版_每页4人"
    for i in range(1, 8):
        ws.column_dimensions[get_column_letter(i)].width = 9.5
    ws.column_dimensions["A"].width = 12

    row = 1
    for r in recs:
        base = row
        # 0 抬头
        ws.merge_cells(start_row=base, start_column=1, end_row=base, end_column=7)
        t = ws.cell(row=base, column=1)
        if rank_note:
            # 小样本：一人一页的结构不动，只把名次换成一句说明
            t.value = "%s  %s（%s）   总分 %g   %s" % (
                "九年级成绩单", r["姓名"], r["班级"], r["总分"], rank_note)
        else:
            t.value = "%s  %s（%s）   总分 %g   年级第 %d 名/%d 人   班级第 %d 名/%d 人" % (
                "九年级成绩单", r["姓名"], r["班级"], r["总分"],
                r["年级名次"], r["年级人数"], r["班级名次"], r["班级人数"])
        t.font = Font(bold=True, size=11)
        t.alignment = Alignment(horizontal="center", vertical="center")
        t.fill = HEAD_FILL
        ws.row_dimensions[base].height = 20

        labels = ["项目"] + [s[0] for s in SUBJECTS] + ["总分"]
        line1 = [r[s[0]] for s in SUBJECTS] + [r["总分"]]
        line2 = [s[2] for s in SUBJECTS] + [TOTAL_FULL]
        line3 = [round(r.get("__年级均分_" + s[0], 0), 1) for s in SUBJECTS] + [round(r.get("__年级均分_总分", 0), 1)]
        line4 = [r[s[0] + "_均差"] for s in SUBJECTS] + [round(r["总分"] - r.get("__年级均分_总分", 0), 1)]
        grid = [
            ("科目", labels[1:], True),
            ("得分", line1, False),
            ("满分", line2, False),
            ("年级均分", line3, False),
            ("与均分差", line4, False),
        ]
        for k, (lab, vals, _) in enumerate(grid, start=1):
            rr = base + k
            c0 = ws.cell(row=rr, column=1, value=lab)
            c0.font = Font(bold=(k == 1), size=9)
            c0.fill = BAND_FILL
            c0.alignment = Alignment(horizontal="center", vertical="center")
            for j, v in enumerate(vals, start=2):
                cell = ws.cell(row=rr, column=j, value=v)
                cell.font = Font(size=9)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(v, float):
                    cell.number_format = "+0.0;-0.0;0.0" if lab == "与均分差" else "0.0"
            for j in range(1, 8):
                ws.cell(row=rr, column=j).border = BOX
            ws.row_dimensions[rr].height = 17

        rr = base + 6
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=7)
        c = ws.cell(row=rr, column=1, value="结论：" + r["结论"])
        c.font = Font(size=9)
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c.border = BOX
        ws.row_dimensions[rr].height = 30

        ws.row_dimensions[base + 7].height = 6
        row += BLOCK_ROWS

    total_rows = row - 1
    for b in range(BLOCK_ROWS * PER_PAGE, total_rows, BLOCK_ROWS * PER_PAGE):
        ws.row_breaks.append(Break(id=b))

    ws.print_options.horizontalCentered = True
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5, header=0.2, footer=0.2)
    ws.print_area = "A1:G%d" % total_rows
    wb.save(path)
    return path, ws.title, total_rows


def cmd_reports(args):
    src = args.src or DEFAULT_SRC
    out = args.out_dir or DEFAULT_OUT
    os.makedirs(out, exist_ok=True)
    recs, meta = load_scores(src)
    # 名次是群体口径：明细两列 / 结论 / 打印版首行 统一按这一条降级
    min_group = getattr(args, "min_group", 3)
    n_valid = sum(1 for r in recs if (r.get("总分") or 0) > 0)
    avg = enrich(recs, min_group)
    for r in recs:
        r["__年级均分_总分"] = avg["__年级__"]
        for sname, _h, _f in SUBJECTS:
            r["__年级均分_" + sname] = avg[sname]

    recs.sort(key=lambda r: (r["年级名次"], r["姓名"]))

    rank_note = print_note = None
    if n_valid < min_group:
        rank_note = "—（样本 %d 人，少于群组下限 %d 人，不排名次）" % (n_valid, min_group)
        print_note = "样本 %d 人，少于群组下限 %d 人，不排名次" % (n_valid, min_group)
        print("小样本：全表有效人数 %d < 群组下限 %d —— 明细两列名次、结论、打印版首行都不出名次"
              % (n_valid, min_group))

    detail_path = os.path.join(out, "成绩单_批量明细.xlsx")
    print_path = os.path.join(out, "成绩单_打印版_每页4人.xlsx")
    build_detail(recs, detail_path, rank_note)
    p, psheet, prows = build_print(recs, print_path, print_note)

    print("=" * 78)
    print("【功能 3】成绩单批量生成")
    print("=" * 78)
    print("源文件 : %s" % meta["file"])
    print("工作表 : %s   学生 %d 人   班级 %d 个   五科满分合计 %g" % (
        meta["sheet"], meta["rows"], len({r["班级"] for r in recs}), TOTAL_FULL))
    zeros = [r for r in recs if r["总分"] <= 0]
    if zeros:
        print("数据质检：剔除 %d 人总分 0（缺考/未录入）：%s"
              % (len(zeros), "、".join("%s(%s)" % (r["姓名"], r["班级"]) for r in zeros[:10])
                 + ("…" if len(zeros) > 10 else "")))
    print("年级均分（按 %d 人算）: " % avg["__计数__"]
          + "  ".join("%s %.2f" % (s[0], avg[s[0]]) for s in SUBJECTS)
          + "  总分 %.2f" % avg["__年级__"])
    print("              成绩单仍然 %d 人一人一份，缺考的 %d 人保留在里面，只标注「疑似缺考，请核实」"
          % (len(recs), len(zeros)))

    mism = sum(1 for r in recs if "源名次" in r and r["源名次"] != r["算出名次"])
    ties = len(recs) - len({round(r["总分"], 2) for r in recs})
    print("名次自洽校验: 本工具按「总分」降序+原表行序重算的年级名次，与源表「总名次」列 → %d/%d 行完全一致%s" % (
        len(recs) - mism, len(recs),
        "" if mism == 0 else "（%d 行不一致，请查源表名次列）" % mism))
    print("              并列总分 %d 人（与别人的分数完全相同），名次按源表行序打断，与学校公布口径一致" % ties)

    # 读回生成的文件做验收，而不是只相信自己写过
    wb = openpyxl.load_workbook(detail_path, data_only=True)
    ws = wb["成绩单明细"]
    rows = list(ws.iter_rows(values_only=True))
    print()
    print("① 明细表路径: %s" % os.path.abspath(detail_path))
    print("   sheet 名  : %s" % ws.title)
    print("   总行数    : %d（表头 1 + 学生 %d）   列数 %d" % (len(rows), len(rows) - 1, len(rows[0])))
    print("   表头      : %s" % " | ".join(str(c) for c in rows[0]))
    print("   前 3 行   :")
    for r in rows[1:4]:
        print("     " + " | ".join("" if c is None else str(c) for c in r))

    wb2 = openpyxl.load_workbook(print_path)
    ws2 = wb2[psheet]
    breaks = [b.id for b in ws2.row_breaks.brk] if ws2.row_breaks else []
    print()
    print("② 打印版路径: %s" % os.path.abspath(print_path))
    print("   sheet 名  : %s" % psheet)
    print("   分页逻辑  : 每人 1 个 %d 行的分块（抬头/科目/得分/满分/年级均分/与均分差/结论/间隔），" % BLOCK_ROWS)
    print("               每页 %d 人 = %d 行，在第 %s 行之后强制分页（共 %d 处分页符）；" % (
        PER_PAGE, BLOCK_ROWS * PER_PAGE, ",".join(str(b) for b in breaks[:6]) + (",…" if len(breaks) > 6 else ""), len(breaks)))
    print("               纸张 A4 竖版、1 页宽自适应、水平居中、打印区域 A1:G%d；总行数 %d → 约 %d 页" % (
        prows, prows, (prows + BLOCK_ROWS * PER_PAGE - 1) // (BLOCK_ROWS * PER_PAGE)))
    print("   实测分页符: %s" % ("无" if not breaks else "%d 个" % len(breaks)))
    wb.close(); wb2.close()
    return 0


# --------------------------------------------------------------------------- #
# 功能 5：多次考试趋势
# --------------------------------------------------------------------------- #
BIG_JUMP = 50.0


def cmd_trend(args):
    files = args.files or [DEFAULT_SRC]
    labels = args.labels or []
    rounds = []
    print("=" * 78)
    print("【功能 5】多次考试总分趋势")
    print("=" * 78)
    print("输入 %d 份文件（顺序 = 时间顺序，先旧后新）" % len(files))
    for i, f in enumerate(files):
        recs, meta = load_scores(f)
        dup = len(recs) - len({r["姓名"] for r in recs})
        if dup:
            print("  警告：%s 内有 %d 个重名，按姓名对齐会串行，请改用 (班级,姓名)" % (meta["file"], dup))
        zeros = [r for r in recs if r["总分"] <= 0]
        if zeros:
            print("  数据质检：剔除 %d 人总分 0（缺考/未录入）：%s"
                  % (len(zeros), "、".join("%s(%s)" % (r["姓名"], r["班级"]) for r in zeros[:10])
                     + ("…" if len(zeros) > 10 else "")))
        # 0 分不是「考了 0 分」而是缺考。把它当数据点会造出假的「大起大落」，
        # 所以只让有分的人进序列；缺考那次就是缺口，不判趋势。
        scores = {r["姓名"]: r["总分"] for r in recs if r["总分"] > 0}
        cls = {r["姓名"]: r["班级"] for r in recs}
        name = labels[i] if i < len(labels) else os.path.splitext(os.path.basename(f))[0]
        rounds.append({"name": name, "file": meta["file"], "scores": scores, "class": cls})
        avg = sum(scores.values()) / len(scores) if scores else 0
        print("  %s %-28s %3d 人有分   总分均分 %.2f" % (
            "①②③④⑤⑥⑦⑧⑨"[i] if i < 9 else str(i + 1), name, len(scores), avg))
    print()

    n = len(rounds)
    all_names = []
    seen = set()
    for rd in rounds:
        for nm in rd["scores"]:
            if nm not in seen:
                seen.add(nm); all_names.append(nm)
    common = [nm for nm in all_names if all(nm in rd["scores"] for rd in rounds)]
    partial = len(all_names) - len(common)
    print("学生去重合计 %d 人；%d 次考试都在的 %d 人%s" % (
        len(all_names), n, len(common),
        "，只出现在部分考试的 %d 人（这些人不参与趋势判断）" % partial if partial else ""))

    if n == 1:
        print()
        print("── 只有 1 次考试：[事实] 单点分数可出，趋势/波动幅度/大起大落都算不出来，本工具不编造。")
        print("   补第 2、3 次考试后再跑，就会自动出折线、波动幅度和 ±%g 分的大起大落标记。" % BIG_JUMP)
        print()
        print("   %-8s %-8s %s" % ("姓名", "班级", "总分（单点）"))
        shown = 0
        for nm in common[:3]:
            print("   %-8s %-8s %g" % (nm, rounds[0]["class"].get(nm, ""), rounds[0]["scores"][nm]))
            shown += 1
        print("   …（共 %d 人，完整数据见输出 xlsx）" % len(all_names))

    table = []
    flags = []
    for nm in all_names:
        vals, present = [], []
        for rd in rounds:
            if nm in rd["scores"]:
                vals.append(rd["scores"][nm]); present.append(rd["name"])
        rec = {"姓名": nm, "班级": rounds[0]["class"].get(nm) or [rd["class"].get(nm) for rd in rounds if nm in rd["class"]][0],
               "序列": vals, "考试": present, "次数": len(vals)}
        if len(vals) >= 2:
            diffs = [round(vals[i + 1] - vals[i], 1) for i in range(len(vals) - 1)]
            rec["变化"] = diffs
            rec["波动幅度"] = round(max(vals) - min(vals), 1)
            rec["最大单次变化"] = max(diffs, key=abs)
            marks = []
            for i, d in enumerate(diffs):
                if abs(d) >= BIG_JUMP:
                    marks.append("%s→%s %s%g（%s）" % (
                        present[i], present[i + 1], "↑" if d > 0 else "↓", abs(d),
                        "大起" if d > 0 else "大落"))
            rec["标记"] = "；".join(marks)
            if marks:
                flags.append(rec)
        else:
            rec["变化"] = []
            rec["波动幅度"] = None
            rec["最大单次变化"] = None
            rec["标记"] = ""
        table.append(rec)

    if n >= 2:
        print()
        print("── 趋势明细（波动幅度 = 最高分 − 最低分）")
        print("   %-8s %-8s %-22s %8s %8s   %s" % ("姓名", "班级", "总分序列", "波动幅度", "最大单次", "大起大落标记"))
        for rec in sorted(table, key=lambda x: -(x["波动幅度"] or 0))[:args.top]:
            print("   %-8s %-8s %-22s %8.1f %+8.1f   %s" % (
                rec["姓名"], rec["班级"], "→".join("%g" % v for v in rec["序列"]),
                rec["波动幅度"], rec["最大单次变化"], rec["标记"] or "—"))
        print()
        print("── 大起大落（相邻两次 |Δ|≥%g 分）：命中 %d 人 / 共同学生 %d 人" % (BIG_JUMP, len(flags), len(common)))
        if flags:
            for rec in flags[:20]:
                print("   ⚑ %s（%s） %s   %s" % (rec["姓名"], rec["班级"], "→".join("%g" % v for v in rec["序列"]), rec["标记"]))
        else:
            print("   无（这份数据里没有相邻两次变化达到 ±%g 分的学生）" % BIG_JUMP)

    out_dir = args.out_dir or DEFAULT_OUT
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, "趋势_总分.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "总分趋势"
    head = ["姓名", "班级"] + [rd["name"] for rd in rounds] + ["次数", "平均", "最高", "最低", "波动幅度", "最大单次变化", "大起大落标记"]
    ws.append(head)
    for c in range(1, len(head) + 1):
        ws.cell(row=1, column=c).font = Font(bold=True)
        ws.cell(row=1, column=c).fill = HEAD_FILL
        ws.cell(row=1, column=c).alignment = Alignment(horizontal="center", wrap_text=True)
    for rec in sorted(table, key=lambda x: x["姓名"]):
        pad = [None] * (n - len(rec["序列"]))
        row = [rec["姓名"], rec["班级"]] + rec["序列"]
        row += [None] * (n - len(rec["序列"]))
        row += [rec["次数"], round(sum(rec["序列"]) / len(rec["序列"]), 1) if rec["序列"] else None,
                max(rec["序列"]) if rec["序列"] else None, min(rec["序列"]) if rec["序列"] else None,
                rec["波动幅度"], rec["最大单次变化"], rec["标记"]]
        ws.append(row)
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 12
    for i in range(3, 3 + n):
        ws.column_dimensions[get_column_letter(i)].width = 12
    for i in range(3 + n, len(head) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 13
    ws.freeze_panes = "C2"
    if n >= 2:
        ch = LineChart()
        ch.title = "每人总分趋势"
        ch.y_axis.title = "总分"
        ch.x_axis.title = "考试次序"
        data = Reference(ws, min_col=3, max_col=2 + n, min_row=1, max_row=ws.max_row)
        ch.add_data(data, from_rows=True, titles_from_data=True)
        ch.height, ch.width = 9, 20
        ws.add_chart(ch, "B%d" % (ws.max_row + 3))
    wb.save(out)
    print()
    print("输出: %s  （sheet 名「总分趋势」%s）" % (
        os.path.abspath(out), "，含折线图" if n >= 2 else "，单点不出折线图"))
    return 0


# --------------------------------------------------------------------------- #
# 功能 6：命题质量档案
# --------------------------------------------------------------------------- #
def split_no(ask):
    """'第1题 相反数' → ('第1题', '相反数')；无题号时整串当题号。"""
    s = (ask or "").strip()
    parts = s.replace("\u3000", " ").split(" ", 1)
    head = parts[0]
    tail = parts[1].strip() if len(parts) > 1 else ""
    if head.startswith("第") or head.endswith("题"):
        return head, tail
    return s, ""


def load_items_db(path):
    con = sqlite3.connect(path)
    try:
        rows = con.execute(
            "select student_id, ask, kp, full_score, got_score from questions "
            "where full_score is not null and full_score > 0 and got_score is not null").fetchall()
    finally:
        con.close()
    out = []
    for sid, ask, kp, full, got in rows:
        no, tail = split_no(ask)
        out.append({"stu": str(sid), "no": no, "kp": (kp or tail or "未标注"), "full": float(full), "got": float(got)})
    return out


def load_items_csv(path):
    raw = None
    for enc in ("utf-8-sig", "gbk", "utf-8"):
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                raw = list(csv.DictReader(f))
            break
        except UnicodeDecodeError:
            continue
    if raw is None:
        raise ValueError("读不出 %s 的编码（试过 utf-8-sig/gbk）" % path)
    alias = [("stu", ("学生", "姓名", "学号")), ("no", ("题号", "题目")),
             ("kp", ("知识点", "知识点名称")), ("full", ("满分", "分值")), ("got", ("得分", "实际得分"))]
    out = []
    for rec in raw:
        got = {}
        for key, names in alias:
            for nm in names:
                if nm in rec and rec[nm] not in (None, ""):
                    got[key] = rec[nm]
                    break
        if len(got) < 5:
            raise ValueError("CSV 表头需含 学生/题号/知识点/满分/得分 五列，缺: %s" %
                             [k for k, _v in alias if k not in got])
        no, tail = split_no(got["no"])
        out.append({"stu": str(got["stu"]).strip(), "no": no, "kp": (got["kp"] or tail or "未标注").strip(),
                    "full": float(got["full"]), "got": float(got["got"])})
    return out


RECOVER = (0.0, 0.45, 0.75, 0.90)


def jitter_rows(rows, seed):
    """只给 seed-demo 用：模拟「同一批题隔一段时间再考，上次错的这次有一部分做对了」。

    seed k 表示原来没拿满的题里，有 RECOVER[k] 比例的（按学生+题号确定性地挑选）这一次拿满。
    k 越大越像「反复练过的题」→ 难度 P 越高。真实归档不走这条路径。
    """
    rate = RECOVER[min(seed, len(RECOVER) - 1)]
    out = []
    for r in rows:
        key = (r["stu"] + "|" + r["no"] + "|" + str(seed)).encode("utf-8")
        f = int.from_bytes(hashlib.md5(key).digest()[:4], "big") / 2.0 ** 32
        got = r["got"]
        if seed and got < r["full"] and f < rate:
            got = r["full"]
        out.append({"stu": r["stu"], "no": r["no"], "kp": r["kp"], "full": r["full"], "got": got})
    return out


def compute_items(rows):
    """P = 平均得分率；D = 高分组(前27%) − 低分组(后27%) 的平均得分率。"""
    per_stu = {}
    for r in rows:
        d = per_stu.setdefault(r["stu"], {"got": 0.0, "full": 0.0, "items": {}})
        d["got"] += r["got"]; d["full"] += r["full"]
        d["items"][r["no"]] = (r["full"], r["got"])
    stus = sorted(per_stu.items(), key=lambda kv: -(kv[1]["got"] / kv[1]["full"] if kv[1]["full"] else 0))
    n = len(stus)
    k = max(1, int(round(n * 0.27)))
    upper = {s for s, _ in stus[:k]}
    lower = {s for s, _ in stus[-k:]}

    groups = {}
    for r in rows:
        g = groups.setdefault(r["no"], {"kp": r["kp"], "full": [], "got": [], "up": [], "low": []})
        rate = r["got"] / r["full"] if r["full"] else 0.0
        g["full"].append(r["full"]); g["got"].append(r["got"])
        if r["stu"] in upper:
            g["up"].append(rate)
        elif r["stu"] in lower:
            g["low"].append(rate)
    items = []
    for no, g in groups.items():
        full = max(g["full"])
        p = (sum(g["got"]) / sum(g["full"])) if sum(g["full"]) else 0.0
        pu = sum(g["up"]) / len(g["up"]) if g["up"] else None
        pl = sum(g["low"]) / len(g["low"]) if g["low"] else None
        d = (pu - pl) if (pu is not None and pl is not None) else None
        items.append({"题号": no, "知识点": g["kp"], "满分": full,
                      "难度P": round(p, 3), "区分度D": (round(d, 3) if d is not None else None),
                      "错误率": round(1 - p, 3), "作答人数": len(g["got"])})
    items.sort(key=lambda x: (len(x["题号"]), x["题号"]))
    return items, n, k


def read_archive():
    if not os.path.exists(ARCHIVE):
        return {"version": 1, "exams": []}
    with open(ARCHIVE, "r", encoding="utf-8") as f:
        return json.load(f)


def write_archive(arc):
    arc["updated"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    tmp = ARCHIVE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(arc, f, ensure_ascii=False, indent=2)
    os.replace(tmp, ARCHIVE)


def cmd_items_archive(args):
    if args.csv:
        rows = load_items_csv(args.csv)
        src = os.path.abspath(args.csv)
    elif args.db:
        rows = load_items_db(args.db)
        src = os.path.abspath(args.db)
    else:
        raise SystemExit("必须给 --db 或 --csv（真实数据接入方式见文件头 docstring）")
    if args.jitter:
        rows = jitter_rows(rows, args.jitter)
    items, nstu, k = compute_items(rows)
    arc = read_archive()
    rec = {"exam": args.exam, "subject": args.subject, "archived_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
           "source": src, "demo": bool(args.demo or args.jitter), "students": nstu,
           "group_size": k, "items": items}
    replaced = False
    for i, old in enumerate(arc["exams"]):
        if old["exam"] == args.exam and old.get("subject") == args.subject:
            arc["exams"][i] = rec; replaced = True
            break
    if not replaced:
        arc["exams"].append(rec)
    write_archive(arc)

    print("[归档] %s / %s ← %s" % (rec["exam"], rec["subject"], src))
    print("       学生 %d 人，高/低分组各 %d 人（27%%），科目题数 %d" % (nstu, k, len(items)))
    print("       %s" % ("覆盖了同名的旧归档记录" if replaced else "新增一条归档记录"))
    print("       %-8s %-14s %6s %8s %9s %9s" % ("题号", "知识点", "满分", "难度P", "区分度D", "错误率"))
    for it in items[:8]:
        print("       %-8s %-14s %6g %8.3f %9s %9.3f" % (
            it["题号"], it["知识点"], it["满分"], it["难度P"],
            "%.3f" % it["区分度D"] if it["区分度D"] is not None else "—", it["错误率"]))
    if len(items) > 8:
        print("       …共 %d 题" % len(items))
    print("       档案: %s（现有 %d 次归档：%s）" % (
        ARCHIVE, len(arc["exams"]), "、".join(e["exam"] for e in arc["exams"])))
    return 0


def cmd_items_seed_demo(args):
    db = args.db or os.path.join(ROOT, "_tmp", "class_demo.db")
    base = load_items_db(db)
    print("[seed-demo] 用模拟库造 3 次模拟考试，只为验证「连续三次都太容易」的判定逻辑")
    print("            数据来源: %s（模拟数据，不是真实历次考试！）" % os.path.abspath(db))
    plan = [("模拟考试A", 1), ("模拟考试B", 2), ("模拟考试C", 3)]
    for name, j in plan:
        rows = jitter_rows(base, j) if j else base
        items, nstu, k = compute_items(rows)
        arc = read_archive()
        rec = {"exam": name, "subject": args.subject, "archived_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
               "source": os.path.abspath(db), "demo": True, "students": nstu, "group_size": k, "items": items}
        arc["exams"] = [e for e in arc["exams"] if not (e["exam"] == name and e.get("subject") == args.subject)]
        arc["exams"].append(rec)
        write_archive(arc)
        avgp = sum(i["难度P"] for i in items) / len(items)
        print("           归档 %-18s 平均 P=%.3f  学生 %d 人  题 %d 道" % (name, avgp, nstu, len(items)))
    return 0


def cmd_items_compare(args):
    arc = read_archive()
    exams = [e for e in arc.get("exams", []) if e.get("subject") == args.subject]
    if not exams:
        print("档案里还没有「%s」的归档记录，先跑 items archive / items seed-demo" % args.subject)
        return 1
    demo = any(e.get("demo") for e in exams)
    print("=" * 78)
    print("【功能 6】命题质量档案 · 历次对比（科目 %s）" % args.subject)
    print("=" * 78)
    print("归档次数: %d  →  %s" % (len(exams), " → ".join(e["exam"] for e in exams)))
    print("档案文件: %s" % ARCHIVE)
    if demo:
        print("⚠ 其中 %d 条标记为 demo（模拟数据），下列 P 值变化只用于验证判定逻辑，"
              "不代表任何真实考试。" % sum(1 for e in exams if e.get("demo")))
        print("  真实使用时每次考完 archive 一次真实每题数据即可。")
    print()

    # 题号 → [(exam_idx, P, D)]
    seq = {}
    for i, e in enumerate(exams):
        for it in e["items"]:
            seq.setdefault(it["题号"], {"kp": it["知识点"], "pts": {}})
            seq[it["题号"]]["kp"] = it["知识点"] or seq[it["题号"]]["kp"]
            seq[it["题号"]]["pts"][i] = it

    print("── 同一题号历次 P 值变化（P 越高越容易）")
    print("   %-8s %-14s %s %10s %10s" % ("题号", "知识点", "P 值序列", "均值", "极差"))
    rows = sorted(seq.items(), key=lambda kv: (len(kv[0]), kv[0]))
    for no, info in rows[:args.top]:
        vals = [info["pts"][i]["难度P"] if i in info["pts"] else None for i in range(len(exams))]
        shown = " → ".join("—" if v is None else "%.3f" % v for v in vals)
        real = [v for v in vals if v is not None]
        print("   %-8s %-14s %s %10.3f %10.3f" % (
            no, info["kp"], shown, sum(real) / len(real), max(real) - min(real)))
    if len(rows) > args.top:
        print("   …共 %d 个题号" % len(rows))

    print()
    print("── 判定：【连续 %d 次 P ≥ %.2f 视为「都太容易」】" % (args.streak, args.easy_p))
    hits, near = [], []
    for no, info in rows:
        vals = [info["pts"][i]["难度P"] if i in info["pts"] else None for i in range(len(exams))]
        longest, cur = 0, 0
        for v in vals:
            cur = cur + 1 if (v is not None and v >= args.easy_p) else 0
            longest = max(longest, cur)
        tail = vals[-args.streak:] if len(vals) >= args.streak else vals
        tail_hit = len(tail) == args.streak and all(v is not None and v >= args.easy_p for v in tail)
        if longest >= args.streak:
            hits.append((no, info, vals, tail_hit))
        elif sum(1 for v in vals if v is not None and v >= args.easy_p) >= 2:
            near.append((no, info, vals))

    if hits:
        print("   命中 %d 个题号：" % len(hits))
        for no, info, vals, tail_hit in hits:
            print("   ✓ %-8s %-14s %s   %s" % (
                no, info["kp"], " → ".join("—" if v is None else "%.3f" % v for v in vals),
                "最近连续 %d 次都 ≥ %.2f，属长期过易" % (args.streak, args.easy_p)
                if tail_hit else "历史上曾连续 %d 次过易，本次已回落到 %.3f" % (args.streak, vals[-1])))
    else:
        print("   命中 0 个：没有任何题号达成「连续 %d 次 P ≥ %.2f」" % (args.streak, args.easy_p))
    if near:
        print("   另 %d 个题号 P ≥ %.2f 但不连续 %d 次，不计入：%s" % (
            len(near), args.easy_p, args.streak, "、".join("%s(%s)" % (n, i["kp"]) for n, i, _v in near[:8])))

    print()
    print("── 区分度 D 提示（D < 0.20 说明这题拉不开高低分）")
    low = [(no, info) for no, info in rows
           if all(info["pts"][i]["区分度D"] is not None and info["pts"][i]["区分度D"] < 0.20 for i in info["pts"])]
    print("   %s" % ("命中 %d 个：%s" % (len(low), "、".join("%s(%s)" % (n, i["kp"]) for n, i in low))
                    if low else "无（本题库的题区分度都在 0.20 以上）"))
    return 0


# --------------------------------------------------------------------------- #
# 学生趋势（教师侧）：一个学生跨几次考试的表现
#   读表走 ingest.py（一表多 sheet / 表头不在第一行 / 合并单元格都能吃）；
#   口径：总分 0 或缺 = 缺考，不进该次统计，趋势在该次断开，不猜不补。
#   预警只陈述事实（谁、哪一科、从多少到多少），不写因果、不评价老师。
# --------------------------------------------------------------------------- #
TREND_DROP_SUBJECT = 15.0        # 单科相邻两次波动 > 15 分算「波动大」
TREND_STREAK = 2                 # 连续下降几次算「连续下滑」
MISSING_TOKENS = ("", "缺考", "缺", "免考", "-", "—", "/", "无")


def _num(v):
    """宽严有度：认得出的数字就转，认不出（缺考/横杠/空）就 None，绝不默认 0。"""
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip().replace(",", "")
    if s in MISSING_TOKENS:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def read_exams(path):
    """用 ingest 读一份成绩表 -> (学生列表, 读取报告)。"""
    here = os.path.dirname(os.path.abspath(__file__))
    if here not in sys.path:
        sys.path.insert(0, here)
    import ingest                                     # 读表层归 Lead 维护，只调用不改

    raw_recs, report = ingest.read_all_sheets(path)
    out = []
    for r in raw_recs:
        name = str(r.get("name") or "").strip()
        if not name:
            continue
        raw = r.get("raw") or {}
        subj = {}
        for sname, hname, _full in SUBJECTS:
            subj[sname] = _num(raw.get(hname))
        total = _num(r.get("total"))
        if total is None:
            got = [v for v in subj.values() if v is not None]
            total = sum(got) if got else None
        out.append({"姓名": name,
                    "班级": str(r.get("cls") or "").strip() or "未分班",
                    "总分": total, "各科": subj})
    return out, report


def rank_of(students):
    """年级名次：只排在这次有分的学生里，按总分降序 + 原表行序打断并列。"""
    idx = [i for i, s in enumerate(students) if s["总分"] is not None and s["总分"] > 0]
    idx.sort(key=lambda i: (-students[i]["总分"], i))
    return {i: pos for pos, i in enumerate(idx, 1)}


def cmd_student_trend(args):
    files = [x for x in (args.files or []) if x]
    if not files:
        raise SystemExit("student-trend 至少要给 1 个成绩表：--files A.xlsx [B.xlsx ...]")
    labels = args.labels or []
    out_dir = args.out_dir or DEFAULT_OUT
    os.makedirs(out_dir, exist_ok=True)

    print("=" * 78)
    print("【学生趋势】一名学生跨几次考试的表现（教师侧 xlsx）")
    print("=" * 78)
    if len(files) < 2:
        print("注意：只给了 %d 份成绩表。单次考试只能出「点」，出不了趋势，" % len(files))
        print("      下面的增减列会全空，下滑预警也不会命中——这是预期，不是 bug。")
        print()

    rounds = []
    for i, path in enumerate(files):
        if not os.path.exists(path):
            raise SystemExit("找不到成绩表：%s" % path)
        students, report = read_exams(path)
        label = labels[i] if i < len(labels) else os.path.splitext(os.path.basename(path))[0]
        valid = [s for s in students if s["总分"] is not None and s["总分"] > 0]
        zeros = [s for s in students if not (s["总分"] is not None and s["总分"] > 0)]
        ranks = rank_of(students)
        for j, s in enumerate(students):
            s["名次"] = ranks.get(j)
        rounds.append({"name": label, "path": path, "students": students,
                       "valid": len(valid), "zeros": len(zeros), "report": report})
        print("  %s %-26s 读到 %d 人，有分 %d 人，缺考/未录 %d 人"
              % ("①②③④⑤⑥⑦⑧⑨"[i] if i < 9 else str(i + 1), label, len(students), len(valid), len(zeros)))
        if zeros:
            print("     数据质检：剔除 %d 人总分 0 或缺（缺考/未录入）：%s"
                  % (len(zeros), "、".join("%s(%s)" % (s["姓名"], s["班级"]) for s in zeros[:10])
                     + ("…" if len(zeros) > 10 else "")))
        skipped = (report or {}).get("skipped") or []
        if skipped:
            print("     跳过的 sheet：%s" % "；".join("%s（%s）" % (n, why) for n, why in skipped))

    # 重名检查：按姓名对不齐就明说，不猜
    dup_all = set()
    for rd in rounds:
        cnt = {}
        for s in rd["students"]:
            cnt[s["姓名"]] = cnt.get(s["姓名"], 0) + 1
        dup_all |= {k for k, v in cnt.items() if v > 1}
    if dup_all:
        print()
        print("注意：有重名（%s），已改用「姓名+班级」对齐。" % "、".join(sorted(dup_all)[:6]))
    keyfn = (lambda s: (s["姓名"], s["班级"])) if dup_all else (lambda s: s["姓名"])

    # 汇总每个学生在各次考试里的表现
    order, seen = [], set()
    for rd in rounds:
        for s in rd["students"]:
            k = keyfn(s)
            if k not in seen:
                seen.add(k)
                order.append(k)
    table = {}
    for k in order:
        table[k] = {"姓名": k[0] if isinstance(k, tuple) else k,
                    "班级": "", "points": {}}
    for i, rd in enumerate(rounds):
        for s in rd["students"]:
            k = keyfn(s)
            rec = table[k]
            if not rec["班级"]:
                rec["班级"] = s["班级"]
            if s["总分"] is not None and s["总分"] > 0:
                rec["points"][i] = s

    names = sorted({r["姓名"] for r in table.values()})
    if args.name:
        want = str(args.name).strip()
        hit = [k for k in order if table[k]["姓名"] == want]
        if not hit:
            raise SystemExit("这几次考试里都找不到学生「%s」。现有名单前 10 个：%s"
                             % (want, "、".join(names[:10])))
        order = hit

    # ---------- 写 xlsx ----------
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "学生趋势"
    head = ["姓名", "班级"]
    for rd in rounds:
        head += ["%s·总分" % rd["name"], "%s·名次（仅教师可见）" % rd["name"], "%s·较上次增减" % rd["name"]]
    head += ["有分次数", "最近一次较上一次增减", "备注"]
    ws.append(head)

    for k in order:
        rec = table[k]
        row = [rec["姓名"], rec["班级"]]
        prev = None
        deltas = []
        for i, rd in enumerate(rounds):
            s = rec["points"].get(i)
            if s is None:
                row += [None, None, None]
                prev = None                    # 缺考把相邻比较断开
                continue
            d = None if prev is None else round(s["总分"] - prev, 1)
            row += [s["总分"], s["名次"], d]
            if d is not None:
                deltas.append(d)
            prev = s["总分"]
        row += [len(rec["points"]), deltas[-1] if deltas else None]
        if not rec["points"]:
            row.append("%d 次都缺考/未录入，请核实" % len(rounds))
        elif len(rec["points"]) < len(rounds):
            row.append("有 %d 次缺考（缺口不补值）" % (len(rounds) - len(rec["points"])))
        else:
            row.append("")
        ws.append(row)

    for c in range(1, len(head) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True)
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(head)), ws.max_row)
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 12
    for c in range(3, len(head) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 15

    # 第二张：各科明细（长表，便于筛选）
    ws2 = wb.create_sheet("各科明细")
    head2 = ["姓名", "班级", "考试", "科目", "满分", "分数", "得分率", "较上次同科增减"]
    ws2.append(head2)
    for k in order:
        rec = table[k]
        for sname, _h, full in SUBJECTS:
            prev = None
            for i, rd in enumerate(rounds):
                s = rec["points"].get(i)
                if s is None:
                    prev = None
                    continue
                v = (s["各科"] or {}).get(sname)
                if v is None:
                    prev = None
                    continue
                d = None if prev is None else round(v - prev, 1)
                prev = v
                ws2.append([rec["姓名"], rec["班级"], rd["name"], sname, full, v,
                            (v / full) if full else None, d])
    for c in range(1, len(head2) + 1):
        cell = ws2.cell(row=1, column=c)
        cell.font = Font(bold=True)
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for r in range(2, ws2.max_row + 1):
        ws2.cell(row=r, column=7).number_format = "0.0%"
    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = "A1:H%d" % ws2.max_row
    for col, w in zip("ABCDEFGH", [10, 12, 16, 8, 8, 9, 9, 16]):
        ws2.column_dimensions[col].width = w

    # 第三张：下滑预警（只陈述事实）
    ws3 = wb.create_sheet("下滑预警")
    head3 = ["姓名", "班级", "预警", "事实依据"]
    ws3.append(head3)
    warn_rows = []
    for k in order:
        rec = table[k]
        got = [i for i in range(len(rounds)) if i in rec["points"]]
        if len(got) < 2:
            continue
        # 连续下滑：末尾连续下降的次数
        seq, run, facts = [], 0, []
        for idx, i in enumerate(got):
            seq.append((rounds[i]["name"], rec["points"][i]["总分"]))
            if idx >= 1:
                prev_i, this_i = got[idx - 1], i
                # 必须是相邻两次考试，中间缺考就断开
                if this_i - prev_i != 1:
                    run = 0
                    continue
                if rec["points"][this_i]["总分"] < rec["points"][prev_i]["总分"]:
                    run += 1
                else:
                    run = 0
        if run >= TREND_STREAK:
            tail = [x for x in seq][-run - 1:]
            warn_rows.append([rec["姓名"], rec["班级"], "连续下滑",
                              "%s：总分连续下降 %d 次（%s）"
                              % ("→".join(n for n, _v in tail),
                                 run, " → ".join("%g" % v for _n, v in tail))])
        # 单科波动大
        for sname, _h, _full in SUBJECTS:
            for idx in range(1, len(got)):
                a, b = got[idx - 1], got[idx]
                if b - a != 1:
                    continue
                va = (rec["points"][a]["各科"] or {}).get(sname)
                vb = (rec["points"][b]["各科"] or {}).get(sname)
                if va is None or vb is None:
                    continue
                d = vb - va
                if abs(d) > TREND_DROP_SUBJECT:
                    warn_rows.append([rec["姓名"], rec["班级"], "单科波动大",
                                      "%s：%s 从 %g 变到 %g（%+.1f 分，超过 %g 分）"
                                      % (sname, rounds[a]["name"] + "→" + rounds[b]["name"],
                                         va, vb, d, TREND_DROP_SUBJECT)])
    for row in warn_rows:
        ws3.append(row)
    for c in range(1, len(head3) + 1):
        cell = ws3.cell(row=1, column=c)
        cell.font = Font(bold=True)
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws3.freeze_panes = "A2"
    ws3.auto_filter.ref = "A1:D%d" % max(2, ws3.max_row)
    for col, w in zip("ABCD", [10, 12, 12, 72]):
        ws3.column_dimensions[col].width = w
    if not warn_rows:
        ws3.append(["—", "—", "未命中", "按上面的判定口径，这次没有触发任何预警"])

    # 第四张：口径与局限（把话说在前面）
    ws4 = wb.create_sheet("口径与局限")
    lines = [
        ("统计口径", ""),
        ("参加考试", "总分 > 0 才算参加了这次考试"),
        ("缺考/未录", "总分为 0 或空的，不进这次的名次与统计；趋势在该次断开，不补值、不插值"),
        ("名次", "只排在这次有分的学生里，按总分降序 + 原表行序打断并列（与成绩单口径一致）"),
        ("名次可见性", "名次只给老师看，学生的成绩单里不出名次以外的内容"),
        ("得分率", "分数 ÷ 满分；五科满分 语文100 数学150 英语150 物理100 化学100"),
        ("", ""),
        ("预警判定（只陈述事实，不做因果推断）", ""),
        ("连续下滑", "相邻两次考试总分连续下降 ≥ %d 次；中间缺考则该次断开" % TREND_STREAK),
        ("单科波动大", "同一科目相邻两次考试相差 > %g 分" % TREND_DROP_SUBJECT),
        ("", ""),
        ("局限", ""),
        ("需要几次考试", "趋势结论至少要 2 次真实考试；「连续下滑」至少要 3 次才能触发（2 次最多只有 1 次下降）"),
        ("本次数据", "本次输入 %d 份成绩表：%s" % (len(rounds), "、".join(r["name"] for r in rounds))),
        ("未验证的部分", "只有一份真实成绩单时，各次数据相同，趋势恒为 0 —— 这是预期，不代表功能验证过结论"),
    ]
    for a, b in lines:
        ws4.append([a, b])
    ws4.column_dimensions["A"].width = 26
    ws4.column_dimensions["B"].width = 88

    suffix = ("_" + str(args.name).strip()) if args.name else ""
    out = os.path.join(out_dir, "学生趋势%s_%d次_%s.xlsx" % (suffix, len(rounds), time.strftime("%Y%m%d")))
    wb.save(out)

    # ---------- 控制台 ----------
    print()
    print("学生 %d 人；有分次数分布：%s" % (
        len(order), "、".join("%d 次 %d 人" % (n, c) for n, c in
                              sorted(_count(len(table[k]["points"]) for k in order).items()))))
    print()
    print("── 下滑预警：命中 %d 条%s" % (len(warn_rows), "" if warn_rows else "（本次没触发）"))
    for row in warn_rows[:8]:
        print("   ⚑ %s（%s）%s：%s" % (row[0], row[1], row[2], row[3]))
    if len(warn_rows) > 8:
        print("   …共 %d 条，完整见 xlsx" % len(warn_rows))
    if len(rounds) < 3:
        print("   说明：「连续下滑」需要 %d 次以上考试才可能触发，本次 %d 次，即使下降也最多算 1 次。"
              % (TREND_STREAK + 1, len(rounds)))

    print()
    print("── 样例（前 3 人）")
    show = order[:3]
    for k in show:
        rec = table[k]
        parts = []
        for i, rd in enumerate(rounds):
            s = rec["points"].get(i)
            parts.append("%s %s" % (rd["name"], ("%g" % s["总分"]) if s else "缺考"))
        print("   %s（%s）：%s" % (rec["姓名"], rec["班级"], " ｜ ".join(parts)))

    print()
    print("输出: %s" % os.path.abspath(out))
    print("       sheet：学生趋势 / 各科明细 / 下滑预警 / 口径与局限")
    return 0


def _count(it):
    d = {}
    for x in it:
        d[x] = d.get(x, 0) + 1
    return d


# --------------------------------------------------------------------------- #
def main(argv=None):
    ap = argparse.ArgumentParser(description="教师侧报告工具：成绩单批量 / 多次考试趋势 / 命题质量档案")
    sub = ap.add_subparsers(dest="cmd")

    p1 = sub.add_parser("reports", help="功能3 成绩单批量生成")
    p1.add_argument("--src", default=None, help="成绩表 xlsx（默认 九年级月考成绩单.xlsx）")
    p1.add_argument("--out-dir", default=None, help="输出目录（默认 教师工具/输出）")
    p1.add_argument("--min-group", dest="min_group", type=int, default=3,
                    help="群组下限：全表有效人数低于它时「年级名次/班级名次」两列改为说明（默认 3）")
    p1.set_defaults(func=cmd_reports)

    p2 = sub.add_parser("trend", help="功能5 多次考试趋势")
    p2.add_argument("--files", nargs="*", default=None, help="按时间顺序列出的 xlsx（可只给一个）")
    p2.add_argument("--labels", nargs="*", default=None, help="每次考试的名称，缺省用文件名")
    p2.add_argument("--out-dir", default=None)
    p2.add_argument("--top", type=int, default=15, help="控制台最多列几行")
    p2.set_defaults(func=cmd_trend)

    p4 = sub.add_parser("student-trend", help="单个学生跨次考试趋势（教师侧）")
    p4.add_argument("--files", nargs="*", required=True, help="按时间顺序列出的成绩表 xlsx")
    p4.add_argument("--labels", nargs="*", default=None, help="每次考试的名称，缺省用文件名")
    p4.add_argument("--name", default=None, help="只看某一名学生")
    p4.add_argument("--out-dir", default=None, help="输出目录（默认 教师工具/输出）")
    p4.set_defaults(func=cmd_student_trend)

    p3 = sub.add_parser("items", help="功能6 命题质量档案")
    s3 = p3.add_subparsers(dest="sub")

    a1 = s3.add_parser("archive", help="算 P/D 并追加归档")
    a1.add_argument("--db", default=None)
    a1.add_argument("--csv", default=None)
    a1.add_argument("--exam", required=True, help="考试名称，如「2026春·第一次月考」")
    a1.add_argument("--subject", default="数学")
    a1.add_argument("--jitter", type=float, default=0.0, help=argparse.SUPPRESS)
    a1.add_argument("--demo", action="store_true", help="标记为模拟数据")
    a1.set_defaults(func=cmd_items_archive)

    a2 = s3.add_parser("compare", help="历次 P 值对比与「连续三次太容易」判定")
    a2.add_argument("--subject", default="数学")
    a2.add_argument("--easy-p", type=float, default=0.90, help="太容易阈值（默认 0.90）")
    a2.add_argument("--streak", type=int, default=3, help="连续几次（默认 3）")
    a2.add_argument("--top", type=int, default=15)
    a2.set_defaults(func=cmd_items_compare)

    a3 = s3.add_parser("seed-demo", help="用模拟库造三次模拟考试归档")
    a3.add_argument("--db", default=None)
    a3.add_argument("--subject", default="数学")
    a3.set_defaults(func=cmd_items_seed_demo)

    args = ap.parse_args(argv)
    if not getattr(args, "func", None):
        ap.print_help()
        return 2
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
