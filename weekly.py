# -*- coding: utf-8 -*-
r"""weekly.py — 周作业登记 + 周报（教师侧命令行，0 次大模型调用）

为什么单开一个工具
------------------
辅导机构的真实频率是「每周」：每周发作业、每周收、每周改。
现有工具里只有考试分析（月考），一个月才用一次，中间三周全班没有任何抓手。
周报要解决的就是这三周：谁交了、谁没交、错在哪、下周先抓谁。

子命令
------
  template  生成空白登记表（固定八列，是/否 带下拉）
  report    读登记表 -> 一人一页 HTML + 一份全班汇总 xlsx

用法
----
  python weekly.py template --out "周作业登记模板.xlsx"
  python weekly.py report --src "周表.xlsx" --out-dir "输出"
  python weekly.py report --src "周表.xlsx" --out-dir "_tmp\周报" --week 2026-W44

口径（写死在这儿，不许在别处放宽）
--------------------------------
1. 缺交 ≠ 未登记：
   · 实交 = 「是」   -> 已交
   · 实交 = 「否」   -> 缺交（进缺交明细、进缺交名单）
   · 实交 留空       -> 未登记（单列一栏，不算缺交、也不算已交）
   · 实交 写了是/否以外的字 -> 记为「无法判定」，原样列出，绝不替老师改判
2. 没有错题数就不算错题分布：错题数空着或整张表都没有这一列 -> 报「算不出来」，不默认 0。
3. 应交只认「是」：应交 = 是 的行才计入「应交次数」；完成率 = 已交 ÷ 应交。
   应交一次都没登记 -> 完成率写「—」，不拿登记行数顶替。
4. 日期解析不了的行不进任何统计，单独计数并列出，等老师修。
5. 一张表里多个 sheet 各自当一张登记表读（一个班一个 sheet），汇总时不合并学生身份。

读表层复用 ingest.py（Lead 维护）：sheet_rows / find_header / norm / read_all_sheets。
本文件只读 xlsx、只写 xlsx 与 html，不改任何既有文件。
"""

import argparse
import html
import io
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import date, datetime

try:
    if (sys.stdout.encoding or "").lower().replace("-", "") not in ("utf8", "utf_8"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import ingest

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_OUT = os.path.join(HERE, "输出")

# ---- 固定八列（模板与读取共用一份定义，改这里就够了）----
WEEK_COLS = ["日期", "学生", "科目", "作业项", "应交", "实交", "错题数", "备注"]
REQUIRED_COLS = ["日期", "学生", "科目", "作业项", "实交"]
OPTIONAL_COLS = ["应交", "错题数", "备注"]

# 列名同义词：只收「明显是同一个意思」的写法，不做模糊匹配
COL_ALIASES = {
    "日期": ["日期", "上课日期", "作业日期", "日期(周)"],
    "学生": ["学生", "姓名", "学生姓名", "学员姓名"],
    "科目": ["科目", "学科"],
    "作业项": ["作业项", "作业", "作业内容", "作业名称", "作业项名称"],
    "应交": ["应交", "应收"],
    "实交": ["实交", "实际交", "实收"],
    "错题数": ["错题数", "错题", "错题个数", "错题数量"],
    "备注": ["备注", "说明", "备注说明"],
}
_ALIAS_NORM = {ingest.norm(a): k for k, names in COL_ALIASES.items() for a in names}

DATE_FORMATS = ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d",
                "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%Y%m%d")

YES, NO, BLANK, WEIRD = "是", "否", "", "?"


class InputError(Exception):
    """输入表读不动（缺列、没有表头、文件打不开）。调用方按退出码 2 处理。"""


# ============================ 读取 ============================

def parse_date(v):
    """日期归一：date/datetime 直接用；字符串只认几种写死的格式；其它返回 None。"""
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip() if v is not None else ""
    if not s:
        return None
    for f in DATE_FORMATS:
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            continue
    return None


def week_label(d):
    """ISO 周：2026-W44。跨年那周按 ISO 的归属年算，不按自然年。"""
    y, w, _ = d.isocalendar()
    return "%d-W%02d" % (y, w)


def parse_yn(v):
    """是/否/空 三分；其它一律当「无法判定」，不替老师改判。"""
    if v is None:
        return BLANK
    s = str(v).strip()
    if s == "":
        return BLANK
    if s == "是":
        return YES
    if s == "否":
        return NO
    return WEIRD


def parse_count(v):
    """错题数：空 -> (None, None)；非负整数 -> (n, '')；其它 -> (None, 原因)。"""
    if v is None:
        return None, None
    if isinstance(v, bool):
        return None, "不是数字"
    if isinstance(v, (int, float)):
        f = float(v)
        if f != int(f):
            return None, "不是整数"
        if f < 0:
            return None, "负数"
        return int(f), None
    s = str(v).strip()
    if s == "":
        return None, None
    if re.fullmatch(r"\d+", s):
        return int(s), None
    return None, "不是非负整数"


def _headers_by_sheet(path):
    """每个 sheet 的表头（找不到表头的 sheet 不出现）。只用于缺列诊断。"""
    out = {}
    for name in ingest.list_sheets(path):
        try:
            rows = ingest.sheet_rows(path, name)
        except Exception as e:
            out[name] = None
            continue
        hi, hdr = ingest.find_header(rows)
        if hi is None:
            out[name] = None
            continue
        out[name] = [("" if h is None else str(h).strip()) for h in hdr]
    return out


def col_map_of(header_names):
    """表头 -> {标准列名: 下标}。认不出的列不进结果。"""
    m = {}
    for i, h in enumerate(header_names or []):
        k = _ALIAS_NORM.get(ingest.norm(h))
        if k and k not in m:
            m[k] = i
    return m


def load_week_table(path):
    """读一张周作业登记表。

    返回 (rows, meta)：
      rows  每条登记一行的 dict（已归一）
      meta  质检与溯源信息：源文件、用到的 sheet、跳过的 sheet 与原因、
            无法解析的行数、无法判定的实交数、缺列情况
    表里认不出表头或缺必填列 -> 抛 InputError（退出码 2），不静默当空表。
    """
    if not os.path.exists(path):
        raise InputError("找不到源文件：%s" % path)

    heads = _headers_by_sheet(path)
    recs, rep = ingest.read_all_sheets(path)
    meta = {"src": os.path.abspath(path), "sheets": rep.get("sheets") or [],
            "used": [name for name, _cnt in (rep.get("used") or [])],
            "skipped": [], "headers": heads,
            "bad_date": [], "bad_yn": [], "bad_count": [], "colmap": {}}

    # 找一张「像周作业表」的 sheet：表头认得出「学生」且认得出「实交」
    usable = []
    for name, hdr in heads.items():
        if hdr is None:
            meta["skipped"].append((name, "前 25 行里找不到表头"))
            continue
        cm = col_map_of(hdr)
        missing = [c for c in REQUIRED_COLS if c not in cm]
        if missing:
            meta["skipped"].append((name, "缺列：%s（表头是：%s）"
                                    % ("、".join(missing), " | ".join(h for h in hdr if h))))
            continue
        usable.append((name, cm))
    if not usable:
        raise InputError("这张表不是周作业登记表：没有任何一个 sheet 同时有 %s。\n"
                         "各 sheet 情况：\n%s\n"
                         "（如果是成绩单，请用 grade_reports.py；周作业登记表请先跑 "
                         "weekly.py template 生成。）"
                         % ("、".join(REQUIRED_COLS),
                            "\n".join("  · %s -> %s" % (n, r) for n, r in meta["skipped"]) or "  （无 sheet）"))
    meta["colmap"] = {n: {k: v for k, v in cm.items()} for n, cm in usable}

    rows = []
    for rec in recs:
        sheet = rec.get("__sheet")
        raw = rec.get("raw") or {}
        raw_norm = {ingest.norm(k): v for k, v in raw.items()}

        def pick(col):
            for a in COL_ALIASES[col]:
                v = raw_norm.get(ingest.norm(a))
                if v is not None and str(v).strip() != "":
                    return v
            return None

        student = str(pick("学生") or "").strip()
        if not student:
            meta.setdefault("no_student", 0)
            meta["no_student"] = int(meta.get("no_student") or 0) + 1
            continue
        row = {"sheet": sheet, "student": student,
               "subject": str(pick("科目") or "").strip(),
               "item": str(pick("作业项") or "").strip(),
               "note": str(pick("备注") or "").strip(),
               "raw_date": pick("日期"), "raw_got": pick("实交"),
               "raw_need": pick("应交"), "raw_errs": pick("错题数")}
        row["date"] = parse_date(row["raw_date"])
        row["week"] = week_label(row["date"]) if row["date"] else ""
        row["got"] = parse_yn(row["raw_got"])
        row["need"] = parse_yn(row["raw_need"])
        n, why = parse_count(row["raw_errs"])
        row["errs"] = n
        if why:
            meta["bad_count"].append((student, row["subject"], str(row["raw_errs"]).strip(), why))
        if row["date"] is None:
            meta["bad_date"].append((student, row["subject"], row["item"], str(row["raw_date"])))
        if row["got"] == WEIRD:
            meta["bad_yn"].append((student, row["subject"], row["item"], str(row["raw_got"]).strip()))
        rows.append(row)
    seen = set()
    uniq = []
    for name, why in meta["skipped"]:
        if name in seen:
            continue
        seen.add(name)
        uniq.append((name, why))
    meta["skipped"] = uniq
    return rows, meta


# ============================ 统计 ============================

SUGGEST_LIMIT = 3


def stats_by_student(rows):
    """按学生聚合。所有计数都写清楚来源，缺数据就是 None，不填 0 顶替。"""
    by = defaultdict(lambda: {"rows": 0, "need": 0, "done": 0, "miss": 0, "blank": 0,
                              "weird": 0, "miss_list": [], "blank_list": [], "weird_list": [],
                              "errs_total": None, "errs_by_subject": Counter(),
                              "errs_rows": 0, "subjects": set(), "dates": [],
                              "sheets": set()})
    for r in rows:
        s = by[r["student"]]
        s["rows"] += 1
        s["sheets"].add(r.get("sheet") or "")
        if r["subject"]:
            s["subjects"].add(r["subject"])
        if r["date"]:
            s["dates"].append(r["date"])
        if r["need"] == YES:
            s["need"] += 1
        if r["got"] == YES:
            s["done"] += 1
        elif r["got"] == NO:
            s["miss"] += 1
            s["miss_list"].append(r)
        elif r["got"] == BLANK:
            s["blank"] += 1
            s["blank_list"].append(r)
        else:
            s["weird"] += 1
            s["weird_list"].append(r)
        if r["errs"] is not None:
            s["errs_total"] = (s["errs_total"] or 0) + r["errs"]
            s["errs_by_subject"][r["subject"] or "未填科目"] += r["errs"]
            if r["errs"] > 0:
                s["errs_rows"] += 1
    for s in by.values():
        s["dates"] = sorted(d for d in s["dates"] if d)
    return dict(by)


def rate_of(s):
    """完成率 = 已交 / 应交。应交为 0 -> None（算不出来就写「—」，不拿行数顶替）。"""
    if not s["need"]:
        return None
    return s["done"] * 100.0 / s["need"]


def week_span(rows):
    """这张表覆盖哪几周：单周 -> 2026-W44；跨周 -> 2026-W43～W44。"""
    weeks = sorted({r["week"] for r in rows if r["week"]})
    if not weeks:
        return "未标周次"
    if len(weeks) == 1:
        return weeks[0]
    return "%s～%s" % (weeks[0], weeks[-1].split("-")[-1] if weeks[0][:4] == weeks[-1][:4] else weeks[-1])


def date_span(rows):
    ds = sorted({r["date"] for r in rows if r["date"]})
    if not ds:
        return "（没有可解析的日期）"
    if ds[0] == ds[-1]:
        return ds[0].isoformat()
    return "%s ~ %s" % (ds[0].isoformat(), ds[-1].isoformat())


def suggestions(name, s):
    """三条建议，每条都带数字；数据不够就直说不够，不编安慰话。"""
    out = []
    r = rate_of(s)
    # ① 完成率
    if s["need"] == 0:
        out.append("这周的「应交」一栏没有一条填了「是」，完成率算不出来——"
                   "先确认这周到底该交几次，再登记。登记里的 %d 条作业不参与完成率。" % s["rows"])
    else:
        if r >= 95:
            tail = "缺交 %d 次，抓一次就够。" % s["miss"] if s["miss"] else "一次没缺。"
            out.append("应交 %d 次、已交 %d 次，完成率 %.0f%%：%s" % (s["need"], s["done"], r, tail))
        elif r >= 80:
            out.append("应交 %d 次、已交 %d 次，完成率 %.0f%%：缺的 %d 次建议下周上课前当面补查，"
                       "别攒到周末。" % (s["need"], s["done"], r, s["miss"]))
        else:
            out.append("应交 %d 次、已交 %d 次，完成率 %.0f%% 明显偏低：先按下面的缺交明细"
                       "逐条对一遍，确认是没写还是没交。" % (s["need"], s["done"], r))
    # ② 缺交/未登记的结构
    if s["miss"]:
        subj = Counter((x["subject"] or "未填科目") for x in s["miss_list"]).most_common(1)[0]
        out.append("缺交 %d 次，最集中在「%s」（%d 次）；另有 %d 条实交空着没填。"
                   "先补交「%s」这条线上的作业，再回填空白登记。"
                   % (s["miss"], subj[0], subj[1], s["blank"], subj[0]))
    elif s["blank"]:
        out.append("这周没有「否」，但有 %d 条实交空着（未登记，不算缺交）："
                   "这些不填清楚，周报只能当没发生——先向学生确认再登记。"
                   % s["blank"])
    elif s["weird"]:
        out.append("%d 条的「实交」写的不是是/否，系统不替老师改判，已原样列在下面。"
                   % s["weird"])
    else:
        out.append("登记是全的：%d 条要么是、要么否，没有空白。" % s["rows"])
    # ③ 错题
    if s["errs_total"] is None:
        out.append("这张表没有填错题数，所以这周不做错题分布（不默认 0）。"
                   "想看点错在哪，下周登记时把错题数填上，一条就够。")
    else:
        top = s["errs_by_subject"].most_common(1)
        if not top or s["errs_total"] == 0:
            out.append("错题合计 0 道（%d 条登记里填了错题数的都记了 0）："
                       "这周不用补练，下周把难度提一档。" % s["errs_rows"])
        else:
            sub, n = top[0]
            out.append("错题合计 %d 道，最集中在「%s」（%d 道，占 %.0f%%）；"
                       "下周先拿这一类做 3 道同类题，再往下走。"
                       % (s["errs_total"], sub, n, n * 100.0 / s["errs_total"]))
    return out[:SUGGEST_LIMIT]


def class_stats(rows, by):
    """全班口径：用于汇总表与终端摘要。"""
    total = len(rows)
    miss = sum(1 for r in rows if r["got"] == NO)
    blank = sum(1 for r in rows if r["got"] == BLANK)
    weird = sum(1 for r in rows if r["got"] == WEIRD)
    need = sum(1 for r in rows if r["need"] == YES)
    done = sum(1 for r in rows if r["got"] == YES)
    return {"rows": total, "students": len(by), "need": need, "done": done,
            "miss": miss, "blank": blank, "weird": weird,
            "rate": (done * 100.0 / need) if need else None}

# ============================ 模板 ============================

TEMPLATE_NOTES = [
    ("怎么用", "一行 = 一次作业登记。同一天同一科可以有多行（两个作业项就写两行）。"),
    ("日期", "填 2026-10-27 这种格式；写成 2026/10/27 也认。解析不了的行不进统计，会在周报里单列。"),
    ("应交", "这次作业是不是必须交：是 / 否。留空表示「还没定」，不计入应交次数。"),
    ("实交", "是 = 交了；否 = 缺交；**留空 = 未登记**（没确认，不算缺交也不算已交）。"),
    ("错题数", "这次作业错了几个题，填整数 0/1/2…；不知道就留空——留空不算 0，不会出现在错题分布里。"),
    ("备注", "老师自己看的话，比如「订正了两遍」「家长说发烧」。"),
    ("示例", "2026-10-27 ｜ 学生A ｜ 数学 ｜ 第三单元练习册 P32 ｜ 是 ｜ 否 ｜ 3 ｜ 订正了两遍"),
    ("提醒", "本模板只有表头、没有示例数据行：直接照着示例往下面加行即可，别把示例那行留在里面。"),
    ("口径", "缺交只算「实交=否」；实交空着叫「未登记」，单独一栏，绝不合并进缺交。"),
]


def build_template(path):
    """固定八列的空白登记表：表头 + 是/否 下拉 + 一行示例 + 说明 sheet。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "周作业登记"
    thin = Side(style="thin", color="C8CDD6")
    head_fill = PatternFill("solid", fgColor="1F4E79")
    head_font = Font(name="微软雅黑", size=10.5, bold=True, color="FFFFFF")
    body_font = Font(name="微软雅黑", size=10.5)
    for i, name in enumerate(WEEK_COLS, start=1):
        c = ws.cell(row=1, column=i, value=name)
        c.fill, c.font = head_fill, head_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        ws.column_dimensions[get_column_letter(i)].width = {
            "日期": 13, "学生": 11, "科目": 8, "作业项": 22,
            "应交": 8, "实交": 8, "错题数": 9, "备注": 26}[name]
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:%s1" % get_column_letter(len(WEEK_COLS))

    # 是/否 下拉：只把这两个值当合法值，写别的会被周报标成「无法判定」
    dv = DataValidation(type="list", formula1='"是,否"', allow_blank=True,
                        showErrorMessage=True, error="只填「是」或「否」；没确认就留空。")
    ws.add_data_validation(dv)
    dv.add("E2:E500")   # 应交
    dv.add("F2:F500")   # 实交

    # 故意不写示例数据行：模板一旦带了示例，老师忘了删就会进统计。
    # 示例写法放在「填写说明」sheet 的「示例」一行里，只作参照。
    ws2 = wb.create_sheet("填写说明")
    ws2.column_dimensions["A"].width = 12
    ws2.column_dimensions["B"].width = 88
    ws2["A1"] = "这一列填什么"
    ws2["B1"] = "说明"
    for c in (ws2["A1"], ws2["B1"]):
        c.font = head_font
        c.fill = head_fill
    for r, (k, v) in enumerate(TEMPLATE_NOTES, start=2):
        ws2.cell(row=r, column=1, value=k).font = Font(name="微软雅黑", size=10.5, bold=True)
        cell = ws2.cell(row=r, column=2, value=v)
        cell.font = body_font
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    wb.save(path)
    return path


# ============================ 周报 HTML ============================

def esc(s):
    return html.escape("" if s is None else str(s), quote=True)


REPORT_CSS = """
@page{size:A4;margin:14mm 12mm}
*{box-sizing:border-box}
body{font-family:"Microsoft YaHei","微软雅黑",sans-serif;color:#1c2430;margin:0;background:#f3f5f8;font-size:12px}
.page{width:186mm;margin:0 auto;background:#fff;padding:8mm 9mm}
.head{display:flex;justify-content:space-between;align-items:flex-end;border-bottom:2px solid #1f4e79;padding-bottom:5px}
h1{font-size:18px;margin:0;color:#1f4e79}
.hmeta{font-size:11px;color:#667;text-align:right;line-height:1.6}
h2{font-size:13px;color:#1f4e79;margin:5mm 0 2mm}
.nums{display:flex;gap:4mm;margin-top:3mm;flex-wrap:wrap}
.num{border:1px solid #e2e8f0;border-radius:8px;padding:2mm 3mm;min-width:30mm}
.num b{display:block;font-size:19px;color:#1f4e79;line-height:1.3}
.num span{font-size:10.5px;color:#77809a}
table{border-collapse:collapse;width:100%;font-size:11px}
th,td{border-bottom:1px solid #eef1f5;padding:3px 4px;text-align:left;vertical-align:top}
th{background:#f5f8fc;color:#3b4a63;font-weight:700}
td.n{text-align:right;white-space:nowrap}
.tag{display:inline-block;border-radius:4px;padding:0 5px;font-size:10px;margin-left:4px}
.tag.miss{background:#fdecea;color:#b03a2e}
.tag.blank{background:#fff6e5;color:#8a5a1a}
.muted{color:#8b93a5}
ol.adv{margin:0;padding-left:16px;line-height:1.85}
ol.adv li{margin-bottom:1.2mm}
.fill{border:1px dashed #c8cdd6;border-radius:8px;padding:3mm;margin-top:3mm}
.fill .line{border-bottom:1px solid #dfe4ec;height:7mm}
.foot{margin-top:4mm;border-top:1px dashed #cbd3de;padding-top:2mm;font-size:10px;color:#889}
.bar{height:7px;background:#1f4e79;border-radius:4px}
.barwrap{background:#eef1f5;border-radius:4px;height:7px}
@media print{body{background:#fff}.page{width:auto;margin:0;padding:0}}
"""


def html_for_student(name, s, meta, week_text, dspan, sheets):
    r = rate_of(s)
    rate_txt = "%.0f%%" % r if r is not None else "—"
    rows = []
    rows.append('<div class="head"><div><h1>周作业周报 · %s</h1>'
                '<div class="hmeta" style="text-align:left;color:#8b93a5">%s</div></div>'
                '<div class="hmeta">周次：%s<br>日期范围：%s<br>生成：%s</div></div>'
                % (esc(name), esc("、".join(sorted(x for x in sheets if x)) or "周作业登记"),
                   esc(week_text), esc(dspan), esc(datetime.now().strftime("%Y-%m-%d %H:%M"))))

    rows.append('<div class="nums">'
                '<div class="num"><b>%d</b><span>登记条数</span></div>'
                '<div class="num"><b>%d</b><span>应交次数</span></div>'
                '<div class="num"><b>%d</b><span>已交次数</span></div>'
                '<div class="num"><b>%d</b><span>缺交（实交=否）</span></div>'
                '<div class="num"><b>%d</b><span>未登记（实交空）</span></div>'
                '<div class="num"><b>%s</b><span>完成率（已交/应交）</span></div>'
                '</div>'
                % (s["rows"], s["need"], s["done"], s["miss"], s["blank"], rate_txt))
    if s["need"] == 0:
        rows.append('<div class="muted" style="margin-top:1mm">应交一栏没有登记「是」，'
                    '所以完成率是「—」——不是 0%%。</div>')

    # 缺交明细
    rows.append("<h2>缺交明细（实交=否）</h2>")
    if s["miss_list"]:
        rows.append("<table><tr><th>日期</th><th>科目</th><th>作业项</th><th>备注</th></tr>")
        for x in sorted(s["miss_list"], key=lambda z: (z["date"] or date.min, z["subject"])):
            rows.append("<tr><td>%s</td><td>%s</td><td>%s</td><td class='muted'>%s</td></tr>"
                        % (esc(x["date"].isoformat() if x["date"] else x["raw_date"]),
                           esc(x["subject"] or "—"), esc(x["item"] or "—"), esc(x["note"] or "")))
        rows.append("</table>")
    else:
        rows.append('<div class="muted">这周没有「实交=否」的记录。</div>')

    # 未登记
    rows.append("<h2>未登记（实交留空，不算缺交）</h2>")
    if s["blank_list"]:
        rows.append("<table><tr><th>日期</th><th>科目</th><th>作业项</th></tr>")
        for x in sorted(s["blank_list"], key=lambda z: (z["date"] or date.min, z["subject"])):
            rows.append("<tr><td>%s</td><td>%s</td><td>%s</td></tr>"
                        % (esc(x["date"].isoformat() if x["date"] else x["raw_date"]),
                           esc(x["subject"] or "—"), esc(x["item"] or "—")))
        rows.append("</table>")
        rows.append('<div class="muted">这些是没确认的，不算缺交；确认后再改成「是」或「否」。</div>')
    else:
        rows.append('<div class="muted">没有空白登记。</div>')

    # 无法判定
    if s["weird"]:
        rows.append("<h2>无法判定（实交不是是/否）</h2><table><tr><th>日期</th><th>科目</th>"
                    "<th>作业项</th><th>实交原文</th></tr>")
        for x in s["weird_list"]:
            rows.append("<tr><td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>"
                        % (esc(x["date"].isoformat() if x["date"] else x["raw_date"]),
                           esc(x["subject"] or "—"), esc(x["item"] or "—"), esc(x["raw_got"])))
        rows.append("</table><div class='muted'>系统不替老师改判，请改成「是」或「否」。</div>")

    # 错题分布
    rows.append("<h2>错题分布（按科目）</h2>")
    if s["errs_total"] is None:
        rows.append('<div class="muted">这张表没有填错题数，所以不出错题分布（不默认 0）。</div>')
    else:
        tot = s["errs_total"]
        rows.append('<div class="muted" style="margin-bottom:1mm">错题合计 <b>%d</b> 道，'
                    '来自 %d 条登记；空白登记不参与。</div>' % (tot, s["errs_rows"]))
        rows.append("<table><tr><th>科目</th><th>错题数</th><th>占比</th><th style='width:40%'>图示</th></tr>")
        for sub, n in s["errs_by_subject"].most_common():
            pct = (n * 100.0 / tot) if tot else 0
            rows.append("<tr><td>%s</td><td class='n'>%d</td><td class='n'>%.0f%%</td>"
                        "<td><div class='barwrap'><div class='bar' style='width:%.0f%%'></div></div></td></tr>"
                        % (esc(sub), n, pct, (max(2, pct) if n else 0)))
        rows.append("</table>")

    # 三条建议
    rows.append("<h2>三条建议</h2><ol class='adv'>")
    for t in suggestions(name, s):
        rows.append("<li>%s</li>" % esc(t))
    rows.append("</ol>")

    # 老师手填
    rows.append('<h2>数据看不到的部分（请老师填）</h2><div class="fill">'
                '<div class="muted" style="margin-bottom:2mm">表里只有登记数据：孩子在家做了什么、'
                '为什么没交、情绪状态、家长配合度，这些系统看不到，请老师当面补。</div>'
                + "".join('<div class="line"></div>' for _ in range(5)) + "</div>")

    rows.append('<div class="foot">口径：缺交只算「实交=否」；实交留空记「未登记」，不并入缺交；'
                '没有错题数就不出错题分布；完成率 = 已交 ÷ 应交，应交为 0 时写「—」。'
                '本页由 weekly.py 生成，数据来自 %s。</div>' % esc(os.path.basename(meta["src"])))
    return ("<!doctype html><html lang=\"zh\"><meta charset=\"utf-8\">"
            "<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">"
            "<title>周报 · %s</title><style>%s</style><div class=\"page\">%s</div></html>"
            % (esc(name), REPORT_CSS, "".join(rows)))


def safe_filename(s):
    """文件名里不能出现的字符换掉，其余原样（中文保留）。"""
    return re.sub(r'[\\/:*?"<>|\s]+', "_", str(s)).strip("_") or "未命名"

# ============================ 全班汇总 xlsx ============================

def _style_header(ws, ncols, height=22):
    thin = Side(style="thin", color="C8CDD6")
    fill = PatternFill("solid", fgColor="1F4E79")
    font = Font(name="微软雅黑", size=10.5, bold=True, color="FFFFFF")
    for i in range(1, ncols + 1):
        c = ws.cell(row=1, column=i)
        c.fill, c.font = fill, font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.row_dimensions[1].height = height
    ws.freeze_panes = "A2"


def build_summary_xlsx(path, by, valid_rows, meta, week_text, dspan):
    """全班汇总：一张汇总表 + 缺交名单 + 未登记名单 + 口径与质检。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "汇总"
    head = ["学生", "来源 sheet", "登记条数", "应交次数", "已交次数",
            "缺交(实交=否)", "未登记(实交空)", "完成率", "错题合计", "错题最多科目"]
    ws.append(head)
    _style_header(ws, len(head))
    for name in sorted(by):
        s = by[name]
        r = rate_of(s)
        top = s["errs_by_subject"].most_common(1)
        ws.append([name,
                   "、".join(sorted(x for x in s["sheets"] if x)),
                   s["rows"], s["need"], s["done"], s["miss"], s["blank"],
                   (r / 100.0) if r is not None else "—",
                   (s["errs_total"] if s["errs_total"] is not None else "未填"),
                   (("%s(%d)" % (top[0][0], top[0][1])) if top and s["errs_total"] else "")])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=8, max_col=8):
        for c in row:
            if isinstance(c.value, float):
                c.number_format = "0.0%"
    widths = [12, 10, 10, 10, 10, 13, 15, 10, 10, 16]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(head)), max(ws.max_row, 2))

    ws2 = wb.create_sheet("缺交名单")
    ws2.append(["学生", "日期", "科目", "作业项", "备注", "来源 sheet"])
    _style_header(ws2, 6)
    for name in sorted(by):
        for x in sorted(by[name]["miss_list"], key=lambda z: (z["date"] or date.min, z["subject"])):
            ws2.append([name, (x["date"].isoformat() if x["date"] else str(x["raw_date"])),
                        x["subject"] or "—", x["item"] or "—", x["note"] or "", x["sheet"] or ""])
    for i, w in enumerate([12, 12, 8, 26, 22, 10], start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.auto_filter.ref = "A1:F%d" % max(ws2.max_row, 2)

    ws3 = wb.create_sheet("未登记名单")
    ws3.append(["学生", "日期", "科目", "作业项", "来源 sheet", "说明"])
    _style_header(ws3, 6)
    for name in sorted(by):
        for x in sorted(by[name]["blank_list"], key=lambda z: (z["date"] or date.min, z["subject"])):
            ws3.append([name, (x["date"].isoformat() if x["date"] else str(x["raw_date"])),
                        x["subject"] or "—", x["item"] or "—", x["sheet"] or "",
                        "实交留空＝未登记，不算缺交"])
    for i, w in enumerate([12, 12, 8, 26, 10, 30], start=1):
        ws3.column_dimensions[get_column_letter(i)].width = w
    ws3.auto_filter.ref = "A1:F%d" % max(ws3.max_row, 2)

    ws4 = wb.create_sheet("口径与质检")
    ws4.column_dimensions["A"].width = 22
    ws4.column_dimensions["B"].width = 96
    items = [
        ("源文件", meta["src"]),
        ("周次", week_text),
        ("日期范围", dspan),
        ("生成时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("处理行数", "%d 行进入统计；%d 行因日期解析不了被排除"
                   % (len(valid_rows), len(meta.get("bad_date") or []))),
        ("学生数", "%d 人" % len(by)),
        ("缺交口径", "只算「实交=否」；实交留空为「未登记」，单列不合并"),
        ("完成率口径", "已交 ÷ 应交（应交=是 的次数）；应交为 0 时写「—」"),
        ("错题口径", "只有填了错题数的行才参与；整张表没填就不出错题分布，不默认 0"),
        ("用到的 sheet", "、".join(n for n in meta.get("used") or []) or "（无）"),
        ("跳过的 sheet", "；".join("%s：%s" % (n, r) for n, r in meta.get("skipped") or []) or "（无）"),
    ]
    if meta.get("bad_date"):
        items.append(("日期解析不了", "；".join("%s/%s/%s（原值 %r）" % (a, b, c, d)
                                            for a, b, c, d in meta["bad_date"][:10])))
    if meta.get("bad_yn"):
        items.append(("实交无法判定", "；".join("%s/%s/%s（原值 %r）" % (a, b, c, d)
                                            for a, b, c, d in meta["bad_yn"][:10])))
    if meta.get("bad_count"):
        items.append(("错题数无效", "；".join("%s/%s（原值 %r，%s）" % (a, b, c, d)
                                          for a, b, c, d in meta["bad_count"][:10])))
    ws4.append(["项目", "内容"])
    _style_header(ws4, 2)
    for k, v in items:
        ws4.append([k, v])
        ws4.cell(row=ws4.max_row, column=2).alignment = Alignment(wrap_text=True, vertical="top")

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    try:
        from openpyxl.worksheet.properties import PageSetupProperties
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    except Exception:
        pass
    ws.print_title_rows = "1:1"
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    wb.save(path)
    return path


# ============================ 子命令 ============================

def cmd_template(a):
    out = os.path.abspath(a.out)
    build_template(out)
    print("=" * 78)
    print("【weekly template】周作业登记模板")
    print("=" * 78)
    print("输出文件: %s" % out)
    print("固定列  : %s" % " | ".join(WEEK_COLS))
    print("必填列  : %s" % " | ".join(REQUIRED_COLS))
    print("取值    : 应交/实交 只认「是」「否」，留空 = 未登记（不算缺交）")
    print("sheet   : %s" % " | ".join(openpyxl.load_workbook(out).sheetnames))
    print()
    print("下一步  : 老师把每天的作业登记进去，然后跑")
    print('          python weekly.py report --src "%s" --out-dir 输出' % os.path.basename(out))
    return 0


def cmd_report(a):
    try:
        rows, meta = load_week_table(a.src)
    except InputError as e:
        print("=" * 78)
        print("【weekly report】读表失败（退出码 2）")
        print("=" * 78)
        print(str(e))
        return 2

    week_filter = (a.week or "").strip()
    all_weeks = sorted({r["week"] for r in rows if r["week"]})
    if week_filter:
        rows = [r for r in rows if r["week"] == week_filter]
        if not rows:
            print("这张表里没有周次 %s 的行。表里出现的周次：%s"
                  % (week_filter, "、".join(all_weeks) or "（没有可解析的日期）"))
            return 2

    dated = [r for r in rows if r["date"]]
    undated = [r for r in rows if not r["date"]]
    by = stats_by_student(dated)
    cs = class_stats(dated, by)
    week_text = week_span(dated)
    dspan = date_span(dated)

    out_dir = os.path.abspath(a.out_dir or DEFAULT_OUT)
    os.makedirs(out_dir, exist_ok=True)

    written = []
    for name in sorted(by):
        s = by[name]
        p = os.path.join(out_dir, "周报_%s.html" % safe_filename(name))
        with open(p, "w", encoding="utf-8") as fh:
            fh.write(html_for_student(name, s, meta, week_text, dspan, s["sheets"]))
        written.append(p)

    xlsx_name = "周报_全班_%s.xlsx" % safe_filename(week_text)
    xpath = os.path.join(out_dir, xlsx_name)
    build_summary_xlsx(xpath, by, dated, meta, week_text, dspan)

    print("=" * 78)
    print("【weekly report】周作业周报")
    print("=" * 78)
    print("源文件  : %s" % meta["src"])
    print("周次    : %s    日期范围: %s" % (week_text, dspan))
    print("用到的 sheet: %s" % ("、".join(n for n in meta["used"]) or "（无）"))
    if meta["skipped"]:
        print("跳过的 sheet: %s" % "；".join("%s：%s" % (n, r) for n, r in meta["skipped"]))
    print("-" * 78)
    print("登记行  : %d 行进入统计（学生 %d 人），其中缺交 %d 行、未登记 %d 行、无法判定 %d 行"
          % (cs["rows"], cs["students"], cs["miss"], cs["blank"], cs["weird"]))
    if undated:
        print("排除行  : %d 行日期解析不了，未进任何统计：" % len(undated))
        for r in undated[:5]:
            print("          · %s / %s / %s（原值 %r）"
                  % (r["student"], r["subject"] or "—", r["item"] or "—", r["raw_date"]))
        if len(undated) > 5:
            print("          … 还有 %d 行（明细见汇总 xlsx 的「口径与质检」）" % (len(undated) - 5))
    if meta.get("bad_yn"):
        print("无法判定: %d 条实交写的不是是/否，原样列出，不替老师改判：" % len(meta["bad_yn"]))
        for x in meta["bad_yn"][:5]:
            print("          · %s / %s / %s（原值 %r）" % x)
    if meta.get("bad_count"):
        print("错题数无效: %d 条不是非负整数，已按「没填」处理：" % len(meta["bad_count"]))
        for x in meta["bad_count"][:5]:
            print("          · %s / %s（原值 %r，%s）" % x)
    if cs["rate"] is None:
        print("完成率  : —（应交一栏没有登记「是」，不拿登记行数顶替）")
    else:
        print("完成率  : %.0f%%（已交 %d ÷ 应交 %d）" % (cs["rate"], cs["done"], cs["need"]))
    errs = sum((s["errs_total"] or 0) for s in by.values())
    have_errs = any(s["errs_total"] is not None for s in by.values())
    print("错题    : %s" % (("合计 %d 道（只有填了错题数的行参与）" % errs) if have_errs
                            else "整张表没有错题数，本次不出错题分布（不默认 0）"))
    print("-" * 78)
    print("① 一人一页 HTML：%d 份" % len(written))
    for p in written[:6]:
        print("   %s" % p)
    if len(written) > 6:
        print("   … 还有 %d 份" % (len(written) - 6))
    print("② 全班汇总 xlsx：%s" % xpath)
    wb = openpyxl.load_workbook(xpath)
    for name in wb.sheetnames:
        ws = wb[name]
        print("   sheet %-10s %d 行 × %d 列" % (name, ws.max_row, ws.max_column))
    print()
    print("口径：缺交只算「实交=否」；实交留空＝未登记（不并入缺交）；"
          "没有错题数就不出错题分布；完成率＝已交÷应交，应交为 0 写「—」。")
    if not dated:
        print()
        print("注意：这张表还没有任何登记行（只有表头），所以没有生成个人周报。"
              "填完数据再跑一次即可。")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="周作业登记与周报（教师侧）：template 出表 / report 出周报")
    sub = ap.add_subparsers(dest="cmd")

    p1 = sub.add_parser("template", help="生成空白周作业登记表（固定八列）")
    p1.add_argument("--out", default=os.path.join(DEFAULT_OUT, "周作业登记模板.xlsx"),
                    help="输出 xlsx 路径（默认 教师工具/输出/周作业登记模板.xlsx）")

    p2 = sub.add_parser("report", help="读登记表 -> 一人一页 HTML + 全班汇总 xlsx")
    p2.add_argument("--src", required=True, help="周作业登记表 xlsx")
    p2.add_argument("--out-dir", default=None, help="输出目录（默认 教师工具/输出）")
    p2.add_argument("--week", default=None,
                    help="只出某一周，如 2026-W44（不写＝整张表；表里多周也可以）")

    a = ap.parse_args(argv)
    if a.cmd == "template":
        return cmd_template(a)
    if a.cmd == "report":
        return cmd_report(a)
    ap.print_help()
    return 1


if __name__ == "__main__":
    sys.exit(main())


