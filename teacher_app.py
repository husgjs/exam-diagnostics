# -*- coding: utf-8 -*-
r"""teacher_app.py — 帮老师算成绩的小工具（本地网页版），双击就能用。

老师把成绩表拖进浏览器页面，后台把 教师工具\ 下的分析脚本挨个跑一遍，
结果按卡片和文件列表吐回页面。

    python teacher_app.py            # 起服务并自动打开浏览器
    python teacher_app.py --port 8800 --no-open

设计上的几条硬规矩
------------------
1. 只监听 127.0.0.1（回环地址）。里面有真实学生姓名，绝不绑 0.0.0.0，
   同一局域网里的其它电脑连不上；另外校验 Host 头，防止 DNS rebinding
   把浏览器骗过来读数据。
2. 六个分析脚本只**调用**不修改，全部走 subprocess，不 import 任何一个是别人的脚本。
3. 产物只写 教师工具\输出\<job>\，不往 C 盘写任何东西。
4. 除 openpyxl（脚本自己用）外只用标准库，不装任何东西。

接口
----
    GET  /                        ui.html 原样吐出（不做变量替换）
    POST /api/upload              表单字段名 file，返回 {"job","file","kind"}
    POST /api/run                 {"job": "..."} → 后台跑，立刻返回 {"job","state":"running"}
    GET  /api/job/<id>            {"state","steps","cards","files","error"}
    GET  /api/file/<id>/<name>    下载单个产物
    GET  /api/zip/<id>            全部产物打成一个 zip（内存里拼，不落盘）
"""

import argparse
import csv
import email.parser
import email.policy
import hashlib
import http.client
import io
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import uuid
import webbrowser
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import quote, unquote

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))          # ...\教师工具
ROOT = os.path.dirname(HERE)                                # 工作区根
OUT_ROOT = os.path.join(HERE, "输出")
UI_FILE = os.path.join(HERE, "ui.html")
PYTHON = sys.executable or r"C:\Users\Administrator\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe"

# 只监听本机：0.0.0.0 = 整条局域网都能打开，表里有真实姓名，绝对不能那样。
BIND_HOST = "127.0.0.1"
DEFAULT_PORT = 8765

# 这个进程跑的到底是哪一版代码。验收前先打 /api/ping 对 build，
# 免得又拿"改之前的进程"测出新问题来（已经栽过两次了）。
APP_FEATURES = ["source-field", "rc=2-degraded", "port-probe", "absentee-excluded",
                "score-entry", "line-band"]
try:
    with open(os.path.abspath(__file__), "rb") as _fh:
        APP_BUILD = hashlib.md5(_fh.read()).hexdigest()[:10]
except Exception:                                        # noqa: BLE001
    APP_BUILD = "unknown"

MAX_UPLOAD = 60 * 1024 * 1024
STEP_TIMEOUT = 300
LINE_TOTAL = 500          # 重点线，和 grade_deep.py 的默认值一致
BAND = 40                 # 临界带宽度，和 grade_deep.py 的默认值一致

# 临界带上沿：已过线、但超出不超过 10 分的也算临界（刚过线的人最容易被忽视）。
# ⚠ 这个上沿必须与 grade_deep.py 里 crit 过滤器的 `< a.line + 10` 保持一致：
#    那处和这里任意一边改了，另一边必须跟着改，否则首页的临界生人数会和明细表的名单差出一批人
#    （统一前就是这样：500/40 时首页 70 人、明细 90 人，差的正是 [line, line+10) 这 20 人）。
CRIT_UP = 10

# /api/run 允许覆盖这两个值。非法值一律报错，绝不静默退回默认值 ——
# 否则老师以为按 450/30 算了，其实还是 500/40，名单对不上却查不出原因。
LINE_MIN, LINE_MAX = 0, 2000
BAND_MIN, BAND_MAX = 0, 300
JOBS_DIR = os.path.join(HERE, "jobs")


def parse_int_param(body, key, default, lo, hi, label):
    """从请求体取一个整数参数。没传 -> default；传了但不合法 -> (None, 错误说明)。"""
    raw = body.get(key, None)
    if raw is None or (isinstance(raw, str) and not raw.strip()):
        return default, None
    if isinstance(raw, bool):
        return None, "%s必须是 %d~%d 之间的整数，收到的是「%s」" % (label, lo, hi, raw)
    try:
        if isinstance(raw, float):
            if raw != int(raw):
                raise ValueError(raw)
            val = int(raw)
        elif isinstance(raw, int):
            val = raw
        else:
            val = int(str(raw).strip())
    except (TypeError, ValueError):
        return None, "%s必须是 %d~%d 之间的整数，收到的是「%s」" % (label, lo, hi, raw)
    if not (lo <= val <= hi):
        return None, "%s必须在 %d~%d 之间，收到的是 %d" % (label, lo, hi, val)
    return val, None


def write_job_json(job, line, band):
    """每次分析落一份任务单：jobs/<job_id>/job.json（口径可回溯，进程重启也不丢）。"""
    d = os.path.join(JOBS_DIR, job["id"])
    os.makedirs(d, exist_ok=True)
    path = os.path.join(d, "job.json")
    payload = {"line": line, "band": band, "src": job["src"], "kind": job["kind"]}
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=1)
    return path

SUBJECTS = [("语文", "语文分数", 100.0), ("数学", "数学分数", 150.0), ("英语", "英语分数", 150.0),
            ("物理", "物理分数", 100.0), ("化学", "化学分数", 100.0)]
FULL_TOTAL = sum(s[2] for s in SUBJECTS)

JOBS = {}
LOCK = threading.Lock()


# --------------------------------------------------------------------------- #
# 小工具
# --------------------------------------------------------------------------- #
def human_size(n):
    for unit in ("B", "KB", "MB"):
        if n < 1024 or unit == "MB":
            return ("%d %s" % (n, unit)) if unit == "B" else ("%.1f %s" % (n, unit))
        n /= 1024.0
    return str(n)


def safe_name(name):
    """去掉路径成分，只留文件名，避免 ../ 之类跑到别的目录去。"""
    name = (name or "").replace("\\", "/").split("/")[-1].strip()
    name = re.sub(r'[<>:"|?*\x00-\x1f]', "_", name)
    if not name or name in (".", ".."):
        name = "成绩表.xlsx"
    return name[:120]


def fix_mojibake(text):
    """浏览器发上来的中文文件名常被按 latin-1 解了一遍，这里试着还原。"""
    if not text:
        return text
    try:
        fixed = text.encode("latin-1").decode("utf-8")
        return fixed
    except (UnicodeEncodeError, UnicodeDecodeError):
        return text


def col_index(header, name):
    """精确匹配列名。表里同时有「小总分」和「总分」，用包含匹配会先命中「小总分」。"""
    hdr = [(str(c).strip() if c is not None else "") for c in header]
    return hdr.index(name) if name in hdr else None


# --------------------------------------------------------------------------- #
# 读表：判断类型 + 算首页卡片
# --------------------------------------------------------------------------- #
def _find_header(rows):
    for i, row in enumerate(rows[:8]):
        vals = [(str(c).strip() if c is not None else "") for c in row]
        if "姓名" in vals or "学号" in vals:
            return i, vals
    return None, None


def detect_kind(path):
    """grade（有总分）/ daily（一行一人多次考试）/ card（一行一题，答题卡）。"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        rows = list(ws.iter_rows(min_row=1, max_row=8, values_only=True))
    finally:
        wb.close()
    hi, hdr = _find_header(rows)
    if hdr is None:
        return "card"
    if "总分" in hdr:
        return "grade"
    score_cols = [h for h in hdr if h not in ("姓名", "学号", "班级")]
    if len(score_cols) >= 2:
        return "daily"
    return "card"


def compute_cards(path, line=None, band=None):
    """首页四张卡片：整体、各班、各科、需要盯的学生。

    line/band 跟着本次分析走：页面上的「临界生」数字必须和 grade_deep 那张表同一个口径，
    否则老师改了口径，明细表变了、首页数字没变，就成了两个数。
    """
    line = LINE_TOTAL if line is None else line
    band = BAND if band is None else band
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        ws = wb.worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()
    hi, hdr = _find_header(rows)
    if hdr is None:
        return [], "表格里找不到「姓名」或「学号」列，没办法按人统计。"

    i_name = col_index(hdr, "姓名")
    if i_name is None:
        i_name = col_index(hdr, "学号")
    i_total = col_index(hdr, "总分")
    i_cls = col_index(hdr, "班级")
    if i_total is None:
        return [], "表格里没有「总分」这一列，只能出文件，出不了首页卡片。"

    recs = []
    for row in rows[hi + 1:]:
        if i_name >= len(row) or row[i_name] is None or str(row[i_name]).strip() == "":
            continue
        total = row[i_total] if i_total < len(row) else None
        try:
            total = float(total or 0)
        except (TypeError, ValueError):
            continue
        rec = {"cls": (str(row[i_cls]).strip() if i_cls is not None and i_cls < len(row) and row[i_cls] else "未分班"),
               "name": str(row[i_name]).strip(), "total": total}
        for sname, hname, _full in SUBJECTS:
            j = col_index(hdr, hname)
            if j is not None and j < len(row):
                try:
                    rec[sname] = float(row[j] or 0)
                except (TypeError, ValueError):
                    rec[sname] = 0.0
        recs.append(rec)
    if not recs:
        return [], "表格里没有读到学生数据。"

    # 总分 0 = 缺考/未录入：算任何统计之前先剔除，
    # 口径要和 grade_deep / grade_analysis / grade_tools / grade_reports 一致。
    zeros = [r for r in recs if r["total"] <= 0]
    valid = [r for r in recs if r["total"] > 0] or recs
    if zeros:
        print("数据质检：剔除 %d 人总分 0（缺考/未录入）：%s"
              % (len(zeros), "、".join("%s(%s)" % (r["name"], r["cls"]) for r in zeros[:10])
                 + ("…" if len(zeros) > 10 else "")), flush=True)
    n = len(valid)
    totals = [r["total"] for r in valid]
    avg = sum(totals) / n
    pass_line, good_line = FULL_TOTAL * 0.6, FULL_TOTAL * 0.85
    pass_rate = sum(1 for t in totals if t >= pass_line) / n
    good_rate = sum(1 for t in totals if t >= good_line) / n

    by_cls = {}
    for r in valid:
        by_cls.setdefault(r["cls"], []).append(r["total"])
    cls_avg = {c: sum(v) / len(v) for c, v in by_cls.items()}
    best_cls = max(cls_avg, key=lambda c: cls_avg[c])
    worst_cls = min(cls_avg, key=lambda c: cls_avg[c])

    cards = []
    cards.append({
        "title": "这次考得怎么样",
        "kpis": [
            {"k": "实际参加考试", "v": "%d 人" % n},
            {"k": "总分平均", "v": "%.1f" % avg},
            {"k": "最高分", "v": "%g" % max(totals)},
            {"k": "及格率", "v": "%.1f%%" % (pass_rate * 100)},
            {"k": "优秀率", "v": "%.1f%%" % (good_rate * 100)},
        ] + ([{"k": "缺考/未录入", "v": "%d 人" % len(zeros)}] if zeros else []),
        "note": "共 %d 个班；满分 %g 分，及格线 %g 分（60%%），优秀线 %g 分（85%%）。%s" % (
            len(by_cls), FULL_TOTAL, pass_line, good_line,
            ("已剔除 %d 人总分 0（缺考/未录入），均分、及格率、优秀率都按 %d 人算。" % (len(zeros), n)) if zeros else ""),
    })

    cards.append({
        "title": "各班总分平均",
        "kpis": [{"k": c, "v": "%.1f" % cls_avg[c]} for c in sorted(cls_avg, key=lambda c: -cls_avg[c])],
        "note": "最高的是%s（%.1f 分），最低的是%s（%.1f 分），相差 %.1f 分。" % (
            best_cls, cls_avg[best_cls], worst_cls, cls_avg[worst_cls], cls_avg[best_cls] - cls_avg[worst_cls]),
    })

    subj_rate = {}
    for sname, _h, full in SUBJECTS:
        if all(sname in r for r in valid):
            subj_rate[sname] = sum(r[sname] for r in valid) / n / full
    if subj_rate:
        weak = min(subj_rate, key=lambda s: subj_rate[s])
        cards.append({
            "title": "五个学科的得分率",
            "kpis": [{"k": s, "v": "%.1f%%" % (subj_rate[s] * 100)} for s, _h, _f in SUBJECTS if s in subj_rate],
            "note": "得分率 = 拿到的分占满分的比例。五科满分不一样，比强弱要看这个，不能直接比分数。"
                    "这次最低的是%s（%.1f%%），是最该补的一科。" % (weak, subj_rate[weak] * 100),
        })

    edge = [t for t in totals if line - band <= t < line + CRIT_UP]
    # 注意：zero 必须从剔除前的原始行里数。totals 已经剔过缺考，
    # 拿它去数 0 分会永远数出 0 人，老师就看不到「有几个人缺考、请核实」这个提醒了。
    zero = len(zeros)
    note = ("临界生：总分从 %d 分到不满 %d 分的都算 —— 没到线但差得不多的，"
            "和刚过线、只超了不到 %d 分的，都是随时会滑下去的那一批（重点线 %d 分，临界带 %d 分）。"
            % (line - band, line + CRIT_UP, CRIT_UP, line, band))
    if zero:
        note += ("另有 %d 人总分 0（缺考/未录入），这 %d 人不计入均分与及格率；"
                 "若是漏登分，补录后需重跑。" % (zero, zero))
    else:
        note += "这次没有总分 0 的人，不用担心有人被漏掉。"
    cards.append({
        "title": "需要盯的学生",
        "kpis": [
            {"k": "临界生", "v": "%d 人" % len(edge)},
            {"k": "缺考/未录入", "v": "%d 人" % zero},
        ],
        "note": note,
    })
    return cards, ""


# --------------------------------------------------------------------------- #
# 跑分析
# --------------------------------------------------------------------------- #
class StepFailed(Exception):
    pass


def build_steps(src, outdir, line=None, band=None):
    """生成流水线。line/band 不传就用 LINE_TOTAL/BAND（老调用照常能跑）。"""
    py = PYTHON
    line = LINE_TOTAL if line is None else int(line)
    band = BAND if band is None else int(band)

    def cmd(*args):
        return [py] + list(args)
    return [
        {"name": "读取成绩表", "cmd": None, "side": False},
        {"name": "班级对比", "cmd": cmd("teacher_stats.py", src), "side": False},
        {"name": "学科强弱", "cmd": cmd("grade_analysis.py", src), "side": False},
        {"name": "需要盯的学生",
         "cmd": cmd("grade_deep.py", src, "--line", str(line), "--band", str(band)), "side": True},
        {"name": "分组与工具表", "cmd": cmd("grade_tools.py", src), "side": False},
        {"name": "每人成绩单", "cmd": cmd("grade_reports.py", "reports", "--src", src, "--out-dir", outdir), "side": False},
        {"name": "考试趋势", "cmd": cmd("grade_reports.py", "trend", "--files", src, "--out-dir", outdir), "side": False},
        {"name": "打包文件", "cmd": None, "side": False},
    ]


RC_OK = 0
RC_WARN = 2      # 别的脚本「降级但已产出」时的约定：活干了、页面也写了，只是要提个醒


def run_cmd(cmd):
    """返回 (stdout, 提示语)。rc=0 完全成功；rc=2 完成但有话要说；其它才算失败。

    注意 subprocess.run 对非零退出**不抛异常**，靠 returncode 判，不能只 try/except。
    """
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    p = subprocess.run(cmd, cwd=HERE, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", timeout=STEP_TIMEOUT, env=env)
    out = p.stdout or ""
    if p.returncode == RC_OK:
        return out, ""
    lines = [ln.strip() for ln in (p.stderr or "").strip().splitlines() if ln.strip()]
    if not lines:
        lines = [ln.strip() for ln in out.strip().splitlines() if ln.strip()]
    detail = "；".join(lines[-3:])
    if p.returncode == RC_WARN:
        return out, detail or "这一步已经生成了结果，但有情况要提示"
    raise StepFailed(detail or ("退出码 %d" % p.returncode))


def collect_files(outdir, src):
    skip = os.path.basename(src)
    files = []
    if not os.path.isdir(outdir):
        return files
    for name in sorted(os.listdir(outdir)):
        full = os.path.join(outdir, name)
        if not os.path.isfile(full) or name == skip or name.startswith("."):
            continue
        if name.lower().endswith(".json"):
            continue   # grade_deep.py 的中间统计，老师看不懂，留在目录里存档但不摆到页面上
        files.append({"name": name, "path": full, "size": human_size(os.path.getsize(full))})
    return files


def run_pipeline(job_id):
    with LOCK:
        job = JOBS[job_id]
        job["state"] = "running"
    src, outdir = job["src"], job["dir"]
    steps = build_steps(src, outdir, job.get("line"), job.get("band"))
    with LOCK:
        job["steps"] = [{"name": s["name"], "state": "pending", "ms": 0} for s in steps]
    t0 = time.time()
    i = 0
    warns = []
    try:
        for i, st in enumerate(steps):
            with LOCK:
                job["steps"][i]["state"] = "running"
            t1 = time.time()
            if st["name"] == "读取成绩表":
                cards, warn = compute_cards(src, job.get("line"), job.get("band"))
                if warn:
                    warns.append(("读取成绩表", warn))
                with LOCK:
                    job["cards"] = cards
                    job["warn"] = warn
            elif st["name"] == "打包文件":
                with LOCK:
                    job["files"] = collect_files(outdir, src)
            else:
                # grade_deep.py 会把统计结果另存到 _tmp/grade_stats.json（写死的全局路径）。
                # 那是别人的位置：先备份，跑完把这个文件复制进本次产物，再把原内容原样还回去。
                side_path = os.path.join(ROOT, "_tmp", "grade_stats.json")
                side_bak = None
                if st["side"] and os.path.exists(side_path):
                    side_bak = open(side_path, "rb").read()
                _out, note = run_cmd(st["cmd"])
                if note:
                    warns.append((st["name"], note))
                if side_bak is not None:
                    if os.path.exists(side_path):
                        shutil.copy2(side_path, os.path.join(outdir, "grade_stats.json"))
                    with open(side_path, "wb") as f:
                        f.write(side_bak)
            with LOCK:
                job["steps"][i]["ms"] = int((time.time() - t1) * 1000)
                job["steps"][i]["state"] = "done"
        with LOCK:
            job["state"] = "done"
            job["warn"] = "；".join("%s：%s" % (n, m) for n, m in warns)
            if warns:
                # 页面不认 warn 这个字段也要看得见，所以再挂一张卡片
                cards = list(job["cards"] or [])
                cards.append({
                    "title": "有几件事要你留意",
                    "kpis": [{"k": n, "v": "已生成"} for n, _m in warns],
                    "note": "；".join(m for _n, m in warns),
                })
                job["cards"] = cards
    except StepFailed as e:
        with LOCK:
            job["state"] = "failed"
            job["error"] = "「%s」这一步没跑通：%s" % (steps[min(i, len(steps) - 1)]["name"], e)
            job["steps"][min(i, len(steps) - 1)]["state"] = "failed"
    except subprocess.TimeoutExpired:
        with LOCK:
            job["state"] = "failed"
            job["error"] = "「%s」跑得太久（超过 %d 秒），已经停下。" % (steps[i]["name"], STEP_TIMEOUT)
            job["steps"][i]["state"] = "failed"
    except Exception as e:                                    # noqa: BLE001 兜底，不让线程静默死掉
        with LOCK:
            job["state"] = "failed"
            job["error"] = "分析中途出错：%s" % e
    finally:
        with LOCK:
            job["total_ms"] = int((time.time() - t0) * 1000)


def job_view(job):
    return {
        "job": job["id"],
        "file": job["file"],
        "kind": job["kind"],
        "state": job["state"],
        "steps": job["steps"],
        "cards": job["cards"],
        "files": job["files"],
        "error": job["error"],
        "warn": job.get("warn", ""),
        "total_ms": job.get("total_ms", 0),
        "line": job.get("line", LINE_TOTAL),
        "band": job.get("band", BAND),
        "job_json": job.get("job_json", ""),
    }


# --------------------------------------------------------------------------- #
# 题库（task-12）：题库.json 的读写内核
#   规矩：校验不过整批拒绝 -> 写 .tmp -> 把旧文件转成备份 -> os.replace 顶上去
# --------------------------------------------------------------------------- #
QB_FILE = os.path.join(HERE, "题库.json")
QB_HTML = os.path.join(HERE, "qb.html")

# --------------------------------------------------------------------------- #
# 逐题得分录入（task-16）
#   空格子 = 没做（导出时不出这一行）；0 分 = 做了得 0 分（出这一行，值是 0）。
#   两者在数据里必须能分开，所以"没做"只能靠"没有这一行"来表达。
# --------------------------------------------------------------------------- #
SCORE_HTML = os.path.join(HERE, "score.html")
SCORES_DIR = os.path.join(HERE, "_scores")
SCORE_HEADER = ["学生", "题号", "知识点", "满分", "得分"]
SCORE_LOCK = threading.Lock()


def score_key(exam):
    """考试名 -> 安全的文件名（去路径成分、去扩展名、限长）。"""
    s = os.path.splitext(safe_name(str(exam or "").strip()))[0]
    return (s[:60] or "未命名考试")


def num_str(v):
    """3.0 写成 3，2.5 写成 2.5，别让 Excel 看到 3.0 这种。"""
    f = float(v)
    return str(int(f)) if f.is_integer() else ("%g" % f)
QB_BAK = QB_FILE + ".bak"
QB_BAK_KEEP = 5
QB_LOCK = threading.Lock()
QB_PENDING = {}
QB_DIFF = ("easy", "mid", "hard")
QB_DIFF_CN = {"easy": "简单", "mid": "中等", "hard": "难题"}
QB_DIFF_FROM_CN = {"简单": "easy", "容易": "easy", "中等": "mid", "中": "mid", "较难": "hard",
                   "难": "hard", "难题": "hard"}
# 题目来源：可选字段，不填就是「真实卷」。校验上不设限，任意字符串都收。
QB_SOURCE_DEFAULT = "真实卷"
QB_SOURCES = ("真实卷", "自编", "真题改编")


def qb_source(it):
    """取一条题的来源；缺这个字段就当真实卷。"""
    if not isinstance(it, dict):
        return QB_SOURCE_DEFAULT
    return str(it.get("source") or "").strip() or QB_SOURCE_DEFAULT


def qb_load():
    if not os.path.exists(QB_FILE):
        return {"meta": {"version": 1, "book": "", "note": "", "updated": ""}, "items": []}
    with open(QB_FILE, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    if not isinstance(data, dict) or not isinstance(data.get("items"), list):
        raise ValueError("题库.json 的结构不对：应该有 meta 和 items 两段")
    return data


def qb_num(v):
    """满分：整数就存整数，保持 3 而不是 3.0，别让文件无谓地变样。"""
    return int(v) if float(v).is_integer() else float(v)


def qb_normalize(it, fallback_source=None):
    opts = it.get("options") or []
    if isinstance(opts, str):
        opts = re.split(r"[|\n]", opts)
    opts = [str(o).strip() for o in opts if str(o).strip()]
    full = float(str(it.get("full")).strip())
    return {
        "no": str(it.get("no")).strip(),
        "chapter": str(it.get("chapter") or "").strip(),
        "kp": str(it.get("kp") or "").strip(),
        "full": qb_num(full),
        "difficulty": str(it.get("difficulty") or "").strip(),
        "err": str(it.get("err") or "").strip(),
        "stem": str(it.get("stem") or "").strip(),
        "answer": str(it.get("answer") or "").strip(),
        "options": opts,
        # 来源留空就沿用原来的（改题不该顺手把来源改掉），没有原来的才按真实卷
        "source": str(it.get("source") or "").strip() or (fallback_source or QB_SOURCE_DEFAULT),
    }


def qb_errors(it):
    """一条题的毛病，返回 [{field, message}]。任何一条有错，整批都不写。"""
    if not isinstance(it, dict):
        return [{"field": "整条", "message": "这一条不是题目记录，看不懂"}]
    errs = []
    if not str(it.get("no") or "").strip():
        errs.append({"field": "题号", "message": "题号不能空"})
    raw = it.get("full")
    if raw is None or str(raw).strip() == "":
        errs.append({"field": "满分", "message": "满分不能空"})
    else:
        try:
            v = float(str(raw).strip())
        except (TypeError, ValueError):
            errs.append({"field": "满分", "message": "满分要填数字，收到「%s」" % raw})
        else:
            if not (v > 0):
                errs.append({"field": "满分", "message": "满分必须大于 0，收到 %s" % int(v)})
            elif v > 100:
                errs.append({"field": "满分", "message": "满分大得不像话（%s），请确认" % int(v)})
    d = str(it.get("difficulty") or "").strip()
    if d not in QB_DIFF:
        errs.append({"field": "难度", "message": "难度只认简单/中等/难题，收到「%s」" % d})
    if not str(it.get("kp") or "").strip():
        errs.append({"field": "知识点", "message": "知识点不能空"})
    opts = it.get("options")
    if opts is not None and not isinstance(opts, (list, str)):
        errs.append({"field": "选项", "message": "选项要一行一个地写"})
    return errs


def qb_rotate_backup():
    """当前题库 -> .bak.1，老的往后挪一位，最多留 QB_BAK_KEEP 份。"""
    if not os.path.exists(QB_FILE):
        return
    for i in range(QB_BAK_KEEP, 0, -1):
        src, dst = "%s.%d" % (QB_BAK, i), "%s.%d" % (QB_BAK, i + 1)
        if os.path.exists(src):
            os.replace(src, dst)
    shutil.copy2(QB_FILE, QB_BAK + ".1")
    for name in os.listdir(HERE):
        m = re.match(r"^题库\.json\.bak\.(\d+)$", name)
        if m and int(m.group(1)) > QB_BAK_KEEP:
            try:
                os.remove(os.path.join(HERE, name))
            except OSError:
                pass


def qb_write(obj):
    """原子写：先落 .tmp，备份旧文件，再 os.replace 一步顶上去。"""
    obj.setdefault("meta", {})
    obj["meta"]["updated"] = time.strftime("%Y-%m-%d")
    data = json.dumps(obj, ensure_ascii=False, indent=2)
    tmp = QB_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="") as fh:
        fh.write(data)
        fh.flush()
        os.fsync(fh.fileno())
    qb_rotate_backup()
    os.replace(tmp, QB_FILE)


def qb_same(a, b):
    """两条题内容是否完全一样（用来判断「覆盖」到底有没有真的改东西）。"""
    for k in ("no", "kp", "err", "stem", "answer", "difficulty"):
        if str(a.get(k) or "").strip() != str(b.get(k) or "").strip():
            return False
    ch = str(a.get("chapter") or "").strip()
    if ch and ch != str(b.get("chapter") or "").strip():
        return False                      # 空章节表示「沿用原来的」，不算改动
    try:
        if float(a.get("full")) != float(b.get("full")):
            return False
    except (TypeError, ValueError):
        return False
    oa = [str(x).strip() for x in (a.get("options") or [])]
    ob = [str(x).strip() for x in (b.get("options") or [])]
    if oa != ob:
        return False
    return qb_source(a) == qb_source(b)      # 比"生效值"：缺字段 == 真实卷


def qb_apply(base, incoming):
    """把 incoming 按 no 去重合进 base，返回 (items, 新增数, 覆盖数)。"""
    by_no = {it["no"]: i for i, it in enumerate(base)}
    added = overwritten = 0
    for it in incoming:
        if it["no"] in by_no:
            idx = by_no[it["no"]]
            merged = dict(it)
            if not merged.get("chapter"):
                merged["chapter"] = base[idx].get("chapter") or ""
            base[idx] = merged
            overwritten += 1
        else:
            by_no[it["no"]] = len(base)
            base.append(it)
            added += 1
    return base, added, overwritten


QB_IMPORT_ALIAS = {
    "no": ("题号", "编号", "序号"),
    "chapter": ("章节", "章"),
    "kp": ("知识点", "考点"),
    "full": ("满分", "分值"),
    "difficulty": ("难度",),
    "err": ("错因", "错因倾向"),
    "stem": ("题干", "题目"),
    "answer": ("答案",),
    "options": ("选项",),
    "source": ("来源", "出处", "题目来源"),
}


def qb_read_xlsx(data):
    """把老师导出的表格读回来 -> [(行号, 记录)]。表头认中文名。"""
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    try:
        ws = wb.worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()
    if not rows:
        raise ValueError("这个表格是空的，一行内容都没有")
    hdr = [(str(c).strip() if c is not None else "") for c in rows[0]]
    col = {}
    for key, names in QB_IMPORT_ALIAS.items():
        for nm in names:
            if nm in hdr:
                col[key] = hdr.index(nm)
                break
    for need in ("no", "kp", "full", "difficulty"):
        if need not in col:
            raise ValueError("表头里找不到「%s」这一列；导入至少要这几列：题号、知识点、满分、难度"
                             % QB_IMPORT_ALIAS[need][0])
    out = []
    for r_i, row in enumerate(rows[1:], start=2):
        def get(key):
            j = col.get(key)
            if j is None or j >= len(row) or row[j] is None:
                return ""
            return str(row[j]).strip()
        no, kp, stem = get("no"), get("kp"), get("stem")
        if not no and not kp and not stem:
            continue                                   # 整行空着，跳过，不当错误
        d = get("difficulty")
        d = QB_DIFF_FROM_CN.get(d, d)
        raw_opts = get("options")
        opts = []
        if raw_opts:
            # 选项按换行分；老师手打成一行的，退回用「 | 」（带空格）分。
            # 不能用裸竖线：选项里本来就有 |-2| 这种绝对值符号，会被切坏。
            parts = raw_opts.split("\n") if "\n" in raw_opts else raw_opts.split(" | ")
            opts = [x.strip() for x in parts if x.strip()]
        out.append((r_i, {"no": no, "chapter": get("chapter"), "kp": kp, "full": get("full"),
                          "difficulty": d, "err": get("err"), "stem": stem,
                          "answer": get("answer"), "options": opts,
                          # 整张表没有「来源」列 -> None，表示"别动原来的来源"，
                          # 否则拿旧表导一次就会把自编题全刷成真实卷
                          "source": get("source") if "source" in col else None}))
    return out


# --------------------------------------------------------------------------- #
# HTTP
# --------------------------------------------------------------------------- #
def content_disposition(name):
    stem, ext = os.path.splitext(name)
    ascii_stem = re.sub(r"[^A-Za-z0-9._-]", "_", stem)
    if len(re.sub(r"[^A-Za-z0-9]", "", ascii_stem)) < 3:   # 中文名占多数：给老客户端一个能用的兜底名
        ascii_stem = "download"
    ascii_name = ascii_stem + re.sub(r"[^A-Za-z0-9.]", "", ext)
    return "attachment; filename=\"%s\"; filename*=UTF-8''%s" % (ascii_name, quote(name))


class Handler(BaseHTTPRequestHandler):
    server_version = "TeacherWorkbench/1.0"
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):                        # 控制台只留有用的一行
        sys.stderr.write("[%s] %s\n" % (time.strftime("%H:%M:%S"), fmt % args))

    # ---------- 基础设施 ----------
    def _host_ok(self):
        host = (self.headers.get("Host") or "").strip().lower()
        host = host.split(":")[0] if not host.startswith("[") else host.rsplit("]:", 1)[0].strip("[]")
        return host in ("127.0.0.1", "localhost", "::1", "")

    def _send(self, code, body, ctype="application/json; charset=utf-8", extra=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False)
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != "HEAD":
            try:
                self.wfile.write(body)
            except (BrokenPipeError, ConnectionResetError):
                pass

    def _err(self, code, msg):
        self._send(code, {"error": msg})

    # ---------- GET ----------
    def do_GET(self):
        if not self._host_ok():
            return self._err(403, "只允许从本机访问")
        path = self.path.split("?")[0]
        if path in ("/", "/index.html"):
            return self._serve_ui()
        if path in ("/qb", "/qb.html"):
            return self._serve_qb()
        if path in ("/score", "/score.html"):
            return self._serve_score()
        if path.startswith("/api/scores/"):
            return self._api_scores_get(unquote(path[len("/api/scores/"):]))
        if path == "/api/qb/export":
            return self._api_qb_export()
        if path == "/api/qb":
            return self._api_qb_list()
        if path.startswith("/api/qb/"):
            return self._api_qb_one(unquote(path[len("/api/qb/"):]))
        if path.startswith("/api/job/"):
            return self._api_job(path[len("/api/job/"):])
        if path.startswith("/api/file/"):
            return self._api_file(path[len("/api/file/"):])
        if path.startswith("/api/zip/"):
            return self._api_zip(path[len("/api/zip/"):])
        if path == "/api/ping":
            return self._send(200, {"app": "teacher-workbench",
                                    "port": self.server.server_address[1],
                                    "build": APP_BUILD,
                                    "features": APP_FEATURES})
        if path == "/favicon.ico":
            self.send_response(204)
            self.send_header("Content-Length", "0")
            self.end_headers()
            return
        self._err(404, "没有这个地址")

    def _serve_ui(self):
        if not os.path.exists(UI_FILE):
            return self._send(503,
                              "<!doctype html><meta charset='utf-8'><h2>界面文件还没放好</h2>"
                              "<p>缺少 ui.html。把它放到 teacher_app.py 同一个目录下再刷新这一页。</p>",
                              ctype="text/html; charset=utf-8")
        with open(UI_FILE, "rb") as f:
            data = f.read()
        self._send(200, data, ctype="text/html; charset=utf-8")

    def _get_job(self, job_id):
        with LOCK:
            job = JOBS.get(job_id)
            return job

    def _api_job(self, job_id):
        job = self._get_job(job_id)
        if not job:
            return self._err(404, "没有这个任务")
        with LOCK:
            view = job_view(job)
        self._send(200, view)

    def _api_file(self, rest):
        job_id, _, name = rest.partition("/")
        name = unquote(name)
        job = self._get_job(job_id)
        if not job:
            return self._err(404, "没有这个任务")
        full = os.path.abspath(os.path.join(job["dir"], name))
        outdir = os.path.abspath(job["dir"])
        if os.path.dirname(full) != outdir or not os.path.isfile(full):
            return self._err(404, "没有这个文件")
        with open(full, "rb") as f:
            data = f.read()
        self._send(200, data, ctype="application/octet-stream",
                   extra={"Content-Disposition": content_disposition(name)})

    def _api_zip(self, job_id):
        job = self._get_job(job_id)
        if not job:
            return self._err(404, "没有这个任务")
        files = collect_files(job["dir"], job["src"])
        if not files:
            return self._err(404, "这个任务还没有产物")
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for f in files:
                z.write(f["path"], arcname=f["name"])
        data = buf.getvalue()
        name = "成绩分析_%s.zip" % job.get("stem", job["id"])
        self._send(200, data, ctype="application/zip",
                   extra={"Content-Disposition": content_disposition(name)})

    # ---------- POST ----------
    def do_POST(self):
        if not self._host_ok():
            return self._err(403, "只允许从本机访问")
        path = self.path.split("?")[0]
        if path == "/api/upload":
            return self._api_upload()
        if path == "/api/run":
            return self._api_run()
        if path == "/api/scores/export":
            return self._api_scores_export()
        if path == "/api/scores/save":
            return self._api_scores_save()
        if path == "/api/scores/students":
            return self._api_scores_students()
        if path == "/api/qb":
            return self._api_qb_save()
        if path == "/api/qb/import":
            return self._api_qb_import()
        if path == "/api/qb/import/commit":
            return self._api_qb_import_commit()
        self._err(404, "没有这个地址")

    def do_DELETE(self):
        if not self._host_ok():
            return self._err(403, "只允许从本机访问")
        path = self.path.split("?")[0]
        if path.startswith("/api/qb/"):
            return self._api_qb_delete(unquote(path[len("/api/qb/"):]))
        self._err(404, "没有这个地址")

    def _read_multipart(self):
        ctype = self.headers.get("Content-Type", "")
        try:
            length = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            raise ValueError("上传长度不对")
        if length <= 0:
            raise ValueError("没有收到内容")
        if length > MAX_UPLOAD:
            raise ValueError("文件太大（超过 %s）" % human_size(MAX_UPLOAD))
        if "multipart/form-data" not in ctype:
            raise ValueError("请用表单方式上传")
        body = self.rfile.read(length)
        raw = b"Content-Type: " + ctype.encode("latin-1", "replace") + \
              b"\r\nMIME-Version: 1.0\r\n\r\n" + body
        msg = email.parser.BytesParser(policy=email.policy.default).parsebytes(raw)
        if not msg.is_multipart():
            raise ValueError("表单格式不对")
        out = []
        for part in msg.iter_parts():
            field = part.get_param("name", header="content-disposition")
            fname = fix_mojibake(part.get_filename() or "")
            payload = part.get_payload(decode=True) or b""
            out.append((field, fname, payload))
        return out

    def _api_upload(self):
        try:
            parts = self._read_multipart()
        except Exception as e:                                # noqa: BLE001
            return self._err(400, str(e))

        picked = None
        for field, fname, payload in parts:
            if field == "file" or (picked is None and fname):
                picked = (fname or "成绩表.xlsx", payload)
                if field == "file":
                    break
        if not picked:
            return self._err(400, "没有收到文件（表单字段名要用 file）")
        fname, payload = picked
        if not fname.lower().endswith(".xlsx"):
            return self._err(400, "只认 .xlsx 表格，收到的是「%s」" % fname)
        if not payload.startswith(b"PK"):
            return self._err(400, "这个文件不像 Excel 表格（.xlsx 应该是个压缩包）")

        job_id = uuid.uuid4().hex[:12]
        stem = os.path.splitext(safe_name(fname))[0]
        outdir = os.path.join(OUT_ROOT, job_id)
        os.makedirs(outdir, exist_ok=True)
        src = os.path.join(outdir, safe_name(fname))
        with open(src, "wb") as f:
            f.write(payload)

        try:
            kind = detect_kind(src)
        except Exception:                                     # noqa: BLE001
            kind = "card"

        with LOCK:
            JOBS[job_id] = {
                "id": job_id, "file": safe_name(fname), "stem": stem, "kind": kind,
                "src": src, "dir": outdir, "state": "ready", "steps": [], "cards": [],
                "files": [], "error": "", "warn": "", "total_ms": 0, "created": time.time(),
                "line": LINE_TOTAL, "band": BAND, "job_json": "",
            }
        self._send(200, {"job": job_id, "file": safe_name(fname), "kind": kind})

    def _api_run(self):
        try:
            length = int(self.headers.get("Content-Length") or 0)
            body = json.loads(self.rfile.read(length).decode("utf-8") or "{}") if length else {}
        except (ValueError, json.JSONDecodeError):
            return self._err(400, "请求格式不对")
        if not isinstance(body, dict):
            return self._err(400, "请求格式不对")
        job_id = body.get("job")
        job = self._get_job(job_id)
        if not job:
            return self._err(404, "没有这个任务")
        line, err = parse_int_param(body, "line", LINE_TOTAL, LINE_MIN, LINE_MAX, "目标重点线")
        if err:
            return self._err(400, err)
        band, err = parse_int_param(body, "band", BAND, BAND_MIN, BAND_MAX, "临界带宽度")
        if err:
            return self._err(400, err)
        with LOCK:
            if job["state"] == "running":
                return self._send(200, {"job": job_id, "state": "running",
                                        "line": job.get("line"), "band": job.get("band")})
            job["state"] = "running"
            job["error"] = ""
            job["cards"] = []
            job["files"] = []
            job["total_ms"] = 0
            job["line"] = line
            job["band"] = band
        job_json = write_job_json(job, line, band)
        with LOCK:
            job["job_json"] = job_json
        threading.Thread(target=run_pipeline, args=(job_id,), daemon=True).start()
        self._send(200, {"job": job_id, "state": "running", "line": line, "band": band,
                         "job_json": job_json})

    # ---------- 题库 ----------
    def _serve_qb(self):
        if not os.path.exists(QB_HTML):
            return self._send(503,
                              "<!doctype html><meta charset='utf-8'><h2>录入界面还没放好</h2>"
                              "<p>缺少 qb.html。把它放到 teacher_app.py 同一个目录下再刷新这一页。</p>",
                              ctype="text/html; charset=utf-8")
        with open(QB_HTML, "rb") as fh:
            self._send(200, fh.read(), ctype="text/html; charset=utf-8")

    def _read_json(self):
        try:
            length = int(self.headers.get("Content-Length") or 0)
            raw = self.rfile.read(length) if length else b""
            return json.loads(raw.decode("utf-8") or "{}")
        except (ValueError, UnicodeDecodeError, json.JSONDecodeError):
            self._err(400, "请求内容看不懂，不是正常提交的数据")
            return None

    def _api_qb_list(self):
        try:
            data = qb_load()
        except Exception as e:                                # noqa: BLE001
            return self._err(500, "题库文件读不出来：%s" % e)
        items = data.get("items", [])
        self._send(200, {
            "ok": True,
            "meta": data.get("meta", {}),
            "book": (data.get("meta") or {}).get("book", ""),
            "total": len(items),
            "missing_stem": sum(1 for it in items if not str(it.get("stem") or "").strip()),
            "items": items,
        })

    def _api_qb_one(self, no):
        try:
            data = qb_load()
        except Exception as e:                                # noqa: BLE001
            return self._err(500, "题库文件读不出来：%s" % e)
        for it in data.get("items", []):
            if str(it.get("no")) == str(no):
                return self._send(200, {"ok": True, "item": it})
        self._err(404, "题库里没有题号「%s」" % no)

    def _api_qb_save(self):
        body = self._read_json()
        if body is None:
            return
        if isinstance(body, dict) and isinstance(body.get("item"), dict):
            body = body["item"]
        errs = qb_errors(body)
        if errs:
            return self._send(400, {
                "ok": False,
                "error": "有 %d 处写错了，整条都没有保存" % len(errs),
                "errors": errs,
            })
        with QB_LOCK:
            data = qb_load()
            no = str(body.get("no") or "").strip()
            old = next((x for x in data["items"] if str(x.get("no")) == no), None)
            try:
                item = qb_normalize(body, qb_source(old) if old else QB_SOURCE_DEFAULT)
            except Exception as e:                            # noqa: BLE001
                return self._send(400, {"ok": False, "error": "这条题存不下来：%s" % e, "errors": []})
            mode = "updated" if old else "added"
            data["items"], _added, _over = qb_apply(data["items"], [item])
            qb_write(data)
            total = len(data["items"])
        missing = not item["stem"]
        self._send(200, {
            "ok": True, "mode": mode, "no": item["no"], "total": total,
            "missing_stem": missing,
            "message": ("已存好。这题还缺题干，回头记得补。" if missing else "已存好。"),
        })

    def _api_qb_delete(self, no):
        with QB_LOCK:
            data = qb_load()
            before = len(data["items"])
            data["items"] = [x for x in data["items"] if str(x.get("no")) != str(no)]
            if len(data["items"]) == before:
                return self._err(404, "题库里没有题号「%s」" % no)
            qb_write(data)
            total = len(data["items"])
        self._send(200, {"ok": True, "removed": no, "total": total})

    def _api_qb_import(self):
        try:
            parts = self._read_multipart()
        except Exception as e:                                # noqa: BLE001
            return self._err(400, str(e))
        payload = None
        for field, fname, data in parts:
            if field == "file" or (payload is None and fname):
                payload = data
                if field == "file":
                    break
        if not payload:
            return self._err(400, "没有收到表格文件")
        try:
            rows = qb_read_xlsx(payload)
        except Exception as e:                                # noqa: BLE001
            return self._err(400, "这个表格读不进来：%s" % e)
        if not rows:
            return self._err(400, "表格里只有表头，没有题目")

        with QB_LOCK:
            existing = {str(x.get("no")): x for x in qb_load()["items"]}
        good, skipped, seen = [], [], set()
        add = over = 0
        for row_no, rec in rows:
            errs = qb_errors(rec)
            norm = None
            if not errs:
                _old = existing.get(str(rec.get("no") or "").strip())
                try:
                    norm = qb_normalize(rec, qb_source(_old) if _old else QB_SOURCE_DEFAULT)
                except Exception as e:                        # noqa: BLE001
                    errs = [{"field": "整条", "message": str(e)}]
            if errs:
                skipped.append({"row": row_no, "no": rec.get("no", ""),
                                "reason": "；".join("%s：%s" % (e["field"], e["message"]) for e in errs)})
                continue
            if norm["no"] in seen:
                skipped.append({"row": row_no, "no": norm["no"], "reason": "这张表里题号重复了，前面那条为准"})
                continue
            seen.add(norm["no"])
            old = existing.get(norm["no"])
            if old is not None:
                if qb_same(norm, old):
                    skipped.append({"row": row_no, "no": norm["no"], "reason": "跟库里一模一样，不用改"})
                    continue
                over += 1
            else:
                add += 1
            good.append(norm)
        token = uuid.uuid4().hex[:12]
        with QB_LOCK:
            if len(QB_PENDING) > 20:
                QB_PENDING.clear()
            QB_PENDING[token] = {"items": good, "at": time.time()}
        self._send(200, {
            "ok": True, "token": token,
            "add": add, "overwrite": over, "skip": len(skipped), "rows": len(rows),
            "missing_stem": sum(1 for x in good if not x["stem"]),
            "skipped": skipped,
            "preview": [{"no": x["no"], "kp": x["kp"], "full": x["full"],
                         "difficulty": QB_DIFF_CN.get(x["difficulty"], x["difficulty"]),
                         "stem": x["stem"][:20]} for x in good[:60]],
            "message": ("一共 %d 行：新增 %d 条、覆盖 %d 条、跳过 %d 条。" % (len(rows), add, over, len(skipped))),
        })

    def _api_qb_import_commit(self):
        body = self._read_json()
        if body is None:
            return
        token = str((body or {}).get("token") or "")
        with QB_LOCK:
            pending = QB_PENDING.pop(token, None)
            if not pending:
                return self._err(400, "这次预览已经作废了，请重新选一次表格导入")
            data = qb_load()
            if pending["items"]:
                data["items"], added, overwritten = qb_apply(data["items"], pending["items"])
                qb_write(data)
            else:
                added = overwritten = 0          # 没有需要改的，文件一个字都不动
            total = len(data["items"])
        self._send(200, {"ok": True, "add": added, "overwrite": overwritten, "total": total})

    def _api_qb_export(self):
        try:
            data = qb_load()
        except Exception as e:                                # noqa: BLE001
            return self._err(500, "题库文件读不出来：%s" % e)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "题库"
        head = ["题号", "章节", "知识点", "满分", "难度", "来源", "错因", "题干", "答案", "选项"]
        ws.append(head)
        for c in range(1, len(head) + 1):
            cell = ws.cell(row=1, column=c)
            cell.font = Font(bold=True)
        for it in data.get("items", []):
            opts = it.get("options") or []
            ws.append([it.get("no", ""), it.get("chapter") or "", it.get("kp") or "", it.get("full"),
                       QB_DIFF_CN.get(it.get("difficulty"), it.get("difficulty") or ""),
                       qb_source(it),
                       it.get("err") or "", it.get("stem") or "", it.get("answer") or "",
                       "\n".join(opts) if opts else ""])
        for i, w in enumerate([8, 8, 18, 8, 8, 11, 12, 52, 16, 40], 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
        buf = io.BytesIO()
        wb.save(buf)
        self._send(200, buf.getvalue(),
                   ctype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   extra={"Content-Disposition": content_disposition(
                       "题库_导出_%s.xlsx" % time.strftime("%Y%m%d"))})


    # ---------- 逐题得分录入 ----------
    def _serve_score(self):
        if not os.path.exists(SCORE_HTML):
            return self._send(503,
                              "<!doctype html><meta charset='utf-8'><h2>录入界面还没放好</h2>"
                              "<p>缺少 score.html。把它放到 teacher_app.py 同一个目录下再刷新。</p>",
                              ctype="text/html; charset=utf-8")
        with open(SCORE_HTML, "rb") as fh:
            self._send(200, fh.read(), ctype="text/html; charset=utf-8")

    def _api_scores_get(self, name):
        raw = str(name or "").strip()
        if raw in ("export", "save", "students"):
            return self._err(404, "这个地址只接受另一种请求方式")
        key = score_key(raw)
        path = os.path.abspath(os.path.join(SCORES_DIR, key + ".json"))
        if os.path.dirname(path) != os.path.abspath(SCORES_DIR) or not os.path.isfile(path):
            return self._err(404, "还没有存过「%s」这份草稿" % key)
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        if not isinstance(data, dict):
            return self._err(500, "草稿文件格式不对")
        data["ok"] = True
        self._send(200, data)

    def _api_scores_save(self):
        body = self._read_json()
        if body is None:
            return
        exam = str((body or {}).get("exam") or "").strip()
        if not exam:
            return self._err(400, "先给这次考试起个名字，再存草稿")
        key = score_key(exam)
        rows = body.get("rows") or []
        if not isinstance(rows, list):
            return self._err(400, "rows 应该是一串格子的得分")
        os.makedirs(SCORES_DIR, exist_ok=True)
        payload = {
            "exam": exam,
            "students": body.get("students") or [],
            "nos": body.get("nos") or [],
            "rows": rows,
            "saved_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        }
        path = os.path.abspath(os.path.join(SCORES_DIR, key + ".json"))
        if os.path.dirname(path) != os.path.abspath(SCORES_DIR):
            return self._err(400, "这个考试名不能当文件名，换一个")
        tmp = path + ".tmp"
        with SCORE_LOCK:
            with open(tmp, "w", encoding="utf-8", newline="") as fh:
                fh.write(json.dumps(payload, ensure_ascii=False, indent=2))
                fh.flush()
                os.fsync(fh.fileno())
            os.replace(tmp, path)
        self._send(200, {"ok": True, "exam": exam, "saved_at": payload["saved_at"],
                         "cells": len(rows), "file": path})

    def _api_scores_export(self):
        body = self._read_json()
        if body is None:
            return
        exam = str((body or {}).get("exam") or "").strip() or "未命名考试"
        rows = (body or {}).get("rows") or []
        if not isinstance(rows, list) or not rows:
            return self._err(400, "一个格子都还没填，没有可以导出的分数")
        nos = [str(x).strip() for x in ((body or {}).get("nos") or []) if str(x).strip()]
        students = [str(x).strip() for x in ((body or {}).get("students") or []) if str(x).strip()]
        if not nos:
            for r in rows:
                n = str((r or {}).get("no") or "").strip()
                if n and n not in nos:
                    nos.append(n)

        errors, clean = [], []
        for i, r in enumerate(rows, 1):
            r = r if isinstance(r, dict) else {}
            stu = str(r.get("student") or "").strip()
            no = str(r.get("no") or "").strip()
            kp = str(r.get("kp") or "").strip()
            r_idx = (students.index(stu) + 1) if stu in students else i
            c_idx = (nos.index(no) + 1) if no in nos else 0
            where = "第 %d 行「%s」" % (r_idx, stu or "（没填名字）")
            if c_idx:
                where += "第 %d 列（第 %s 题%s）" % (c_idx, no, (" " + kp) if kp else "")
            else:
                where += "「%s」" % (no or "没填题号")

            bad = []
            if not stu:
                bad.append("学生名字是空的")
            if not no:
                bad.append("题号是空的")
            full = score = None
            try:
                full = float(str(r.get("full")).strip())
                if not (full > 0):
                    bad.append("满分必须是大于 0 的数，收到 %s" % num_str(full))
                    full = None
            except (TypeError, ValueError):
                bad.append("满分不是数字（收到「%s」）" % (r.get("full"),))
            try:
                score = float(str(r.get("score")).strip())
            except (TypeError, ValueError):
                bad.append("得分不是数字（收到「%s」）" % (r.get("score"),))
            if score is not None and full is not None:
                if score < 0:
                    bad.append("得分不能是负数（%s）" % num_str(score))
                elif score > full:
                    bad.append("得分 %s 超过了满分 %s（只能是 0 到 %s）"
                               % (num_str(score), num_str(full), num_str(full)))
            if bad:
                errors.append({"row": r_idx, "col": c_idx, "student": stu, "no": no,
                               "where": where, "message": "；".join(bad)})
            else:
                clean.append({"student": stu, "no": no, "kp": kp,
                              "full": full, "score": score})

        if errors:
            head = "有 %d 处不合格，没导出：%s" % (len(errors), errors[0]["where"] + " " + errors[0]["message"])
            if len(errors) > 1:
                head += "（另外还有 %d 处）" % (len(errors) - 1)
            return self._send(400, {"ok": False, "error": head, "errors": errors})

        buf = io.StringIO()
        w = csv.writer(buf, lineterminator="\r\n")
        w.writerow(SCORE_HEADER)
        for r in clean:
            w.writerow([r["student"], r["no"], r["kp"], num_str(r["full"]), num_str(r["score"])])
        data = b"\xef\xbb\xbf" + buf.getvalue().encode("utf-8")   # UTF-8 BOM，Excel 才不乱码
        fname = "逐题得分_%s_%s.csv" % (score_key(exam), time.strftime("%Y%m%d"))
        self._send(200, data, ctype="text/csv; charset=utf-8",
                   extra={"Content-Disposition": content_disposition(fname)})

    def _api_scores_students(self):
        try:
            parts = self._read_multipart()
        except Exception as e:                                # noqa: BLE001
            return self._err(400, str(e))
        payload = None
        for field, fname, data in parts:
            if field == "file" or (payload is None and fname):
                payload = data
                if field == "file":
                    break
        if not payload:
            return self._err(400, "没有收到成绩表")
        try:
            wb = openpyxl.load_workbook(io.BytesIO(payload), data_only=True)
            try:
                names = []
                for ws in wb.worksheets:
                    rows = list(ws.iter_rows(values_only=True))
                    hi = None
                    for i, row in enumerate(rows[:8]):
                        vals = [(str(c).strip() if c is not None else "") for c in row]
                        if "姓名" in vals or "学号" in vals:
                            hi = i
                            break
                    if hi is None:
                        continue
                    hdr = [(str(c).strip() if c is not None else "") for c in rows[hi]]
                    j = hdr.index("姓名") if "姓名" in hdr else hdr.index("学号")
                    for row in rows[hi + 1:]:
                        if j < len(row) and row[j] is not None and str(row[j]).strip():
                            nm = str(row[j]).strip()
                            if nm not in names:
                                names.append(nm)
                    if names:
                        break
            finally:
                wb.close()
        except Exception as e:                                # noqa: BLE001
            return self._err(400, "这个表格读不进来：%s" % e)
        if not names:
            return self._err(400, "表格里没找到「姓名」或「学号」那一列")
        self._send(200, {"ok": True, "students": names, "count": len(names)})


def probe_port(port):
    """端口状态：free 没人用 / ours 已经是这个工作台 / foreign 被别的程序占着。

    只判断「有没有人在听」是不够的：Windows 的 SO_REUSEADDR 允许两个进程绑同一端口，
    而本机 8765 上就住着另一个程序（一个 QQ 机器人）。分不清 ours 和 foreign，
    老师双击两次会起两个服务，撞上别人的程序还会把浏览器指到不相干的页面上去。
    """
    import socket
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(0.6)
    try:
        s.connect(("127.0.0.1", port))
    except OSError:
        return "free"
    finally:
        s.close()
    try:
        conn = http.client.HTTPConnection("127.0.0.1", port, timeout=1.0)
        conn.request("GET", "/api/ping")
        resp = conn.getresponse()
        body = resp.read(300)
        conn.close()
        if resp.status == 200 and b"teacher-workbench" in body:
            return "ours"
    except Exception:                                     # noqa: BLE001
        pass
    return "foreign"


def main(argv=None):
    global UI_FILE
    ap = argparse.ArgumentParser(description="帮老师算成绩的小工具：本地网页版成绩分析（只在本机开放）")
    ap.add_argument("--port", type=int, default=DEFAULT_PORT)
    ap.add_argument("--no-open", action="store_true", help="启动后不自动打开浏览器")
    ap.add_argument("--ui", default=UI_FILE, help=argparse.SUPPRESS)
    args = ap.parse_args(argv)

    UI_FILE = os.path.abspath(args.ui)
    os.makedirs(OUT_ROOT, exist_ok=True)

    # 先看看端口状况再决定怎么起：本机 8765 上住着别的程序（一个 QQ 机器人），
    # 而 Windows 的 SO_REUSEADDR 又允许两个进程绑同一端口，不先探一下会绑串。
    first = args.port
    action, target = "fail", first
    for p in [first] + list(range(first + 1, first + 21)):
        st = probe_port(p)
        if st == "ours":
            action, target = "already", p
            break
        if st == "free":
            action, target = "run", p
            break
    if action == "already":
        url = "http://%s:%d/" % (BIND_HOST, target)
        print("=" * 68)
        print(" 这个工作台已经开着了：%s" % url)
        print(" 不用再开一个；直接把浏览器切过去用就行。")
        print("=" * 68)
        if not args.no_open:
            webbrowser.open(url)
        return 0
    if action == "fail":
        print(" 端口 %d 往后 20 个都被别的程序占着了。" % first)
        print(" 换一个端口再启动，例如：python teacher_app.py --port %d" % (first + 100))
        return 1
    if target != first:
        print(" [!] 端口 %d 被别的程序占着，自动改用 %d。" % (first, target))
    args.port = target

    url = "http://%s:%d/" % (BIND_HOST, args.port)
    print("=" * 68)
    print(" 成绩分析工作台已启动")
    print(" 打开这个地址使用：%s" % url)
    print(" 只在本机开放（绑 %s），同一局域网里的其它电脑打不开。" % BIND_HOST)
    print(" 数据只留在这台电脑的 %s" % OUT_ROOT)
    print(" 关掉这个黑窗口就停止。按 Ctrl+C 也可以。")
    print("=" * 68)
    if not os.path.exists(UI_FILE):
        print(" [!] 还没找到界面文件 %s" % UI_FILE)
    sys.stdout.flush()

    if not args.no_open:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()

    httpd = ThreadingHTTPServer((BIND_HOST, args.port), Handler)
    httpd.daemon_threads = True
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止。")
    finally:
        httpd.server_close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
