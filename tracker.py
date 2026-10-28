# -*- coding: utf-8 -*-
"""教师工具 · 第 3/4 层：知识追踪（BKT）/ 边际收益排序 / 等值化

    python tracker.py mastery --scores 第1次.csv 第2次.csv 第3次.csv [--student 学生07] [--kp 绝对值]
    python tracker.py gain    --scores 第1次.csv 第2次.csv 第3次.csv [--student 学生07]
    python tracker.py anchor  --files 第1次.csv 第2次.csv 第3次.csv --anchor 8,12,6

数学依据：教师工具\\深度分析口径.md（第 3/4 层）。本文件只实现口径，不发明口径；
改口径先改那份 md，再同步改这里的顶部常量。

三条规矩（沿用 insight.py，一个字不放宽）：
  1. 样本不足就明说：n=0 说「还没练过」，n=1~2 标「低置信」，次数不够就说「趋势做不了」。
  2. 阈值/参数全部写在顶部常量里，并写清依据；代码里不出现魔法数字。
  3. 所有数字由代码算出：0 次大模型调用，没有任何一句是模型生成的。

**不做 EM 拟合。** 理由（口径原文）：EM 需要每个知识点上百条观测，现在一个班一个知识点
只有几十条；用 EM 会拟合出极端参数（比如 transit=0.9），看起来准、实际在拟合噪声。
默认参数 + 可配置更稳。所以这里只用口径给的默认四参数，允许命令行覆盖。

实现约定（口径没写死、但必须一致的地方，全部写在这里）：
  · BKT 是二值观测：得分 == 满分才算「答对」；部分得分按「没答对」处理，条数会打在质检里。
  · 同一个学生同一个知识点在一场考试里有多道题时，按题号自然序逐条更新（题号是字符串，
    所以先按长度再按字典序，'2' 在 '19' 前面）。
  · 边际收益的成本口径：一道同点变式题按 ITEM_MINUTES 分钟计（含讲评与订正），
    宁可高估成本，也不虚报收益。
  · 等值化的 k 用「相对第一次考试」的累积口径，否则第三次的可比分数只跟第二次可比。
"""
import argparse
import datetime
import json
import math
import os
import sys

try:
    if (sys.stdout.encoding or "").lower().replace("-", "") not in ("utf8", "utf_8"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

try:
    # 复用 insight.py 的读表（csv/txt/xlsx + 编码兜底）与写表（表头样式统一），
    # 免得两套文件读出两套口径。两个文件都在 教师工具 目录下。
    from insight import num, load_bank, out_path, pct, read_table, write_xlsx
except ImportError as e:      # pragma: no cover
    raise SystemExit("tracker.py 复用同目录 insight.py 的读表/写表/题库加载；"
                     "请把两个文件放在同一个 教师工具 目录下。原始错误：%s" % e)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(HERE, "输出")
BANK = os.path.join(HERE, "题库.json")

# ══════════════════════════════════════════════════════════════════════════
# 参数与阈值 —— 全部来自 深度分析口径.md，改口径先改 md 再改这里
# ══════════════════════════════════════════════════════════════════════════

# BKT 四参数（口径表原文依据）
P_INIT = 0.30       # 第一次作答前就掌握了的概率。依据：保底假设
P_TRANSIT = 0.15    # 每练一次从未掌握变为掌握的概率。依据：约 7 次练习到 70% 掌握，
                    # 与一线「同点练 5–8 道」的经验一致
P_GUESS = 0.20      # 没掌握却蒙对的概率。依据：四选一概率 0.25，填空更低，取 0.20
P_SLIP = 0.10       # 掌握了却做错的概率。依据：粗心/口误，取 10%

# 置信度分档（口径）：n=0 不出概率；n=1~2 出概率但标低置信；n>=3 正常出
N_LOW_CONF = 3

# 边际收益（口径）：目标掌握度 P*、单题耗时
P_TARGET = 0.70     # 口径原文：过及格线所需的经验掌握度
ITEM_MINUTES = 5.0  # 一道同点变式题的成本（独立做 ~2 分钟 + 讲评与订正 ~3 分钟）。
                    # 取 5 而不是 2：宁可高估成本、低估收益，也不给老师虚高的「每小时收益」

# 成本过高提醒（口径没给线，这里为「说清楚」设一条，不改排序）
PRACTICE_WARN = 40  # 一个知识点练到 40 道以上已不现实：一学期同点练习量通常 20 道内；
                    # 超过它说明该生的坎不在「练得少」，得换讲法或退到更基础的题

# 数据前提（口径「数据前提」表）
TREND_MIN_EXAMS = 3   # 第 4 层要 ≥3 次考试才谈趋势；不到就只出单次边际收益
ANCHOR_MIN = 2        # 锚题少于 2 道不做等值化
EXAM_MIN_ANCHOR = 2   # 等值化至少要 2 次考试

# 口径补充 v2（2026-10-29）：前置链约束 + 遗忘衰减。依据：深度分析口径.md 末尾「补充 v2」
P_PREREQ_LOW = 0.50   # 口径原文：P(前置) < 0.5 且 P(本点) < 0.7 -> 本点补不动，先补前置
CLASS_PREREQ_MIN = 3  # 班级前置瓶颈的报出门槛：1~2 人的前置问题是个体辅导，不占用课堂时间
DECAY_TAU_DAYS = 30   # 口径原文：两次作答相隔 d 天，作答前先回落 exp(-d/τ)；τ 默认 30 天
TREE = os.path.join(HERE, "teaching_out", "知识点树.json")


class TrackerError(Exception):
    """可读的业务错误：命令行里直接打印，退出码 2。"""


def natkey(no):
    """题号自然序：'2' < '19' < '105'。题号是字符串，直接比会排出 '105' < '19'。"""
    s = str(no).strip()
    return (0, int(s), "") if s.isdigit() else (1, 0, s)


def fmt_hours(hours):
    if hours is None:
        return "—"
    return "%.1f h" % hours if hours >= 1 else "%.0f 分钟" % (hours * 60)


# ══════════════════════════════════════════════════════════════════════════
# 读考试数据 + 质检
# ══════════════════════════════════════════════════════════════════════════
def load_exam(path):
    """读一次考试的逐题得分（口径同 grade_reports/insight：学生,题号,知识点,满分,得分）。

    质检：满分缺失/为 0、得分超满分、得分是负、缺学生名——逐条记下来，不静默丢。
    """
    if not os.path.exists(path):
        raise TrackerError("找不到文件：%s" % path)
    head, rows = read_table(path)
    c_stu, c_no = 0, 1
    c_kp, c_full, c_got = 2, 3, 4
    # 按表头找列，找不到才退回固定位置（口径的五列顺序）
    for i, h in enumerate(head):
        if h in ("学生", "姓名"):
            c_stu = i
        elif h in ("题号", "题目"):
            c_no = i
        elif h in ("知识点", "考点"):
            c_kp = i
        elif h in ("满分", "分值"):
            c_full = i
        elif h in ("得分", "分数"):
            c_got = i
    out, qc = [], {"rows_in": len(rows), "rows_ok": 0, "partial": 0,
                   "bad_full": 0, "over_full": 0, "no_name": 0, "notes": []}
    seen = set()
    for r in rows:
        if len(r) <= max(c_stu, c_no, c_got):
            continue
        stu = str(r[c_stu] or "").strip()
        no = str(r[c_no] or "").strip()
        if not stu or not no:
            qc["no_name"] += 1
            continue
        full = num(r[c_full]) if c_full < len(r) else None
        got = num(r[c_got]) if c_got < len(r) else None
        if got is None:
            got = 0.0
        if not full or full <= 0:
            qc["bad_full"] += 1
            continue
        if got > full + 1e-9:
            qc["over_full"] += 1
            got = full
        if got < 0:
            qc["notes"].append("%s 题 %s 得分是负数（%g），按 0 记" % (stu, no, got))
            got = 0.0
        if 0 < got < full - 1e-9:
            qc["partial"] += 1
        kp = str(r[c_kp] or "").strip() if c_kp < len(r) else ""
        if not kp:
            kp = "未标注"
        key = (stu, no)
        if key in seen:
            qc["notes"].append("%s 题 %s 重复出现，只认第一条" % (stu, no))
            continue
        seen.add(key)
        out.append({"student": stu, "no": no, "kp": kp, "full": full, "got": got,
                    "ok": got >= full - 1e-9})
    if not out:
        raise TrackerError("这份逐题得分里没有能用的行（表头要有 学生,题号,知识点,满分,得分）：%s" % path)
    qc["rows_ok"] = len(out)
    qc["students"] = len({r["student"] for r in out})
    qc["items"] = len({r["no"] for r in out})
    qc["kps"] = len({r["kp"] for r in out})
    qc["full_total"] = sum(r["full"] for r in out) / qc["students"] if qc["students"] else 0.0
    return {"path": path, "label": os.path.splitext(os.path.basename(path))[0],
            "rows": out, "qc": qc}


def print_qc(exam, idx=None):
    qc = exam["qc"]
    print("  数据质检〔%s〕：%d 行 → 有效 %d 行｜%d 名学生｜%d 道题｜%d 个知识点｜每人满分 %.0f"
          % (exam["label"], qc["rows_in"], qc["rows_ok"], qc["students"], qc["items"],
             qc["kps"], qc["full_total"]))
    bits = []
    if qc["partial"]:
        bits.append("部分得分 %d 行（BKT 是二值观测：未满分按「没答对」记）" % qc["partial"])
    if qc["over_full"]:
        bits.append("得分超满分 %d 行（按满分封顶）" % qc["over_full"])
    if qc["bad_full"]:
        bits.append("满分缺失/为 0 的 %d 行（已丢弃）" % qc["bad_full"])
    if qc["no_name"]:
        bits.append("缺学生名/题号的 %d 行（已丢弃）" % qc["no_name"])
    if bits:
        print("    " + "；".join(bits))
    for n in qc["notes"][:4]:
        print("    提示：%s" % n)


def require_exams(paths, need=1, what="这次分析"):
    if len(paths) < need:
        raise TrackerError("%s至少要 %d 份逐题得分，现在只给了 %d 份。" % (what, need, len(paths)))
    exams = []
    for p in paths:
        exams.append(load_exam(p))
    return exams


def data_premise(exams, anchors=None, decay_line=None):
    """把口径「数据前提」表对照着讲一遍：够做什么、不够做什么。"""
    n = len(exams)
    print("  数据前提：%d 次考试%s｜口径要求：第 4 层趋势需 ≥%d 次；等值化需 ≥%d 次且每次 ≥%d 道锚题"
          % (n, ("｜锚题 %s" % "、".join(anchors)) if anchors else "",
             TREND_MIN_EXAMS, EXAM_MIN_ANCHOR, ANCHOR_MIN))
    if n == 1:
        print("    → 只有 1 次数据：掌握概率能算（按这一场），但**追踪曲线只有一段、趋势做不了**。")
    elif n < TREND_MIN_EXAMS:
        print("    → 只有 %d 次数据：掌握概率按 %d 条观测更新，但**趋势做不了**（口径要 ≥%d 次）。"
              % (n, n, TREND_MIN_EXAMS))
    else:
        print("    → %d 次考试，够做追踪；趋势结论仍以「本次观测」为准。" % n)
    if decay_line:
        print("    → %s" % decay_line)
    return n


# ══════════════════════════════════════════════════════════════════════════
# BKT
# ══════════════════════════════════════════════════════════════════════════
def parse_dates(dates, n_exams, tau):
    """--dates 与 --scores 一一对应；给了就必须等长，不猜、不补、不排序。"""
    if tau is None or tau <= 0:
        raise TrackerError("--tau 必须 > 0（遗忘衰减的时间常数，单位天），现在是 %s" % tau)
    if not dates:
        return None
    if len(dates) != n_exams:
        raise TrackerError("--dates 给了 %d 个日期，--scores 有 %d 份考试：必须一一对应，"
                           "少给/多给都当没说清。" % (len(dates), n_exams))
    days = []
    for s in dates:
        try:
            days.append(datetime.date.fromisoformat(str(s).strip()))
        except ValueError:
            raise TrackerError("--dates 里的「%s」不是 YYYY-MM-DD 格式，照 2026-09-01 这样写。" % s)
    for a, b in zip(days, days[1:]):
        if b < a:
            raise TrackerError("--dates 必须按时间顺序给：%s 在 %s 之后。" % (b, a))
    return days


def decay_line(days, tau):
    """口径补充 1 的「说清楚」：没给日期就明说没做衰减，不许偷偷当没有间隔。"""
    if not days:
        return ("未提供考试日期，本次未做遗忘衰减——跨月考的长间隔会高估掌握度"
                "（口径补充 v2 补充 1；补法：加 --dates 2026-09-01 2026-10-01 2026-11-01）")
    gaps = [(b - a).days for a, b in zip(days, days[1:])]
    span = (days[-1] - days[0]).days
    return ("遗忘衰减：τ=%.0f 天，按 %d 个考试日期在每次作答前回落，跨场间隔 %s 天（跨度 %d 天）"
            % (tau, len(days), "/".join("%d" % g for g in gaps) if gaps else "0", span))


def load_tree(path):
    """知识点树：返回 (名字->节点, 提示语或 None)。

    prereq 存的是节点 id（如 rat-numberline），这里立刻换成名字 —— 少了这一步，前置链等于没有。
    树缺了就明说，不假装做过前置判断。
    """
    if not path or not os.path.exists(path):
        return None, "没找到知识点树（%s），本次不做前置链判断（口径补充 v2 补充 2）。" % path
    try:
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)
    except Exception as e:
        return None, "知识点树读不出来（%s），本次不做前置链判断。原始错误：%s" % (path, e)
    nodes = data.get("nodes") or []
    id2name = {}
    for n in nodes:
        if n.get("id"):
            id2name[n["id"]] = n.get("name") or n["id"]
    by_name = {}
    for n in nodes:
        name = n.get("name")
        if not name:
            continue
        by_name[name] = {"id": n.get("id"), "name": name,
                         "prereq": [id2name.get(x, x) for x in (n.get("prereq") or [])]}
    if not by_name:
        return None, "知识点树里没有 nodes[].name，本次不做前置链判断（%s）。" % path
    return by_name, None


def ancestors_of(tree, kp):
    """沿 prereq 往上走，返回 [(名字, 距本点层数)]，层数越大越靠前（越基础）。"""
    if not tree or kp not in tree:
        return []
    seen, out, frontier = {kp}, [], [(x, 1) for x in tree[kp]["prereq"]]
    while frontier:
        name, depth = frontier.pop(0)
        if name in seen:
            continue
        seen.add(name)
        out.append((name, depth))
        if name in tree:
            frontier.extend([(x, depth + 1) for x in tree[name]["prereq"]])
    return out


def find_break(tree, kp, p_of):
    """口径补充 2：P(前置)<0.5 且 P(本点)<0.7 才算断点；多个断点只报最靠前的那个。

    p_of(name) 返回 None 表示这一点还没有观测：不判、不臆造，只记进 unknown 提示。
    """
    p_kp = p_of(kp)
    if not tree or p_kp is None or p_kp >= P_TARGET:
        return None
    broken, unknown = [], []
    for name, depth in ancestors_of(tree, kp):
        pv = p_of(name)
        if pv is None:
            unknown.append(name)
        elif pv < P_PREREQ_LOW:
            broken.append((depth, name, pv))
    if not broken:
        return {"break": None, "unknown": unknown}
    depth, name, pv = sorted(broken, key=lambda x: (-x[0], x[1]))[0]
    return {"break": {"kp": name, "P": pv, "depth": depth}, "unknown": unknown}


def inconsistency(tree, kp, p_of):
    """口径补充 3：前置 P<0.5 但本点 P>0.7 —— 不改概率、不压制，只返回最靠前的一条给人看。"""
    p_kp = p_of(kp)
    if not tree or p_kp is None or p_kp <= P_TARGET:
        return None
    hits = []
    for name, depth in ancestors_of(tree, kp):
        pv = p_of(name)
        if pv is not None and pv < P_PREREQ_LOW:
            hits.append({"prereq": name, "p_prereq": pv, "kp": kp, "p_kp": p_kp, "depth": depth})
    if not hits:
        return None
    return sorted(hits, key=lambda x: (-x["depth"], x["prereq"]))[0]


def p_lookup(res, student):
    """返回 name -> P 的查表函数；没观测过的点给 None（不臆造概率）。"""
    def f(name):
        cell = res.get((student, name))
        return cell["P"] if cell else None
    return f


def class_prereq_block(tree, kps, students, res):
    """口径补充 2 的班级用法：按「断点」聚合，一个断点只出一行。

    理由：课堂上老师只会做一次「过一遍数轴」，把同一个断点拆成两行（有理数减法一行、
    有理数乘方一行）等于把一件事拆成两条账单。

    人数 = 卡在这个断点上的**不同学生数**（同一个人因不同知识点卡在同一断点，只算一次）。
    residual = 过完这个断点之后，受影响知识点里还剩多少人卡在别的前置（<2 人属于课后单独
    辅导，不写进班级建议；这一层过滤放在打印处做）。只统计有观测的学生。
    """
    per_kp = {}          # 知识点 -> {学生: 该生在该点上的最早断点}
    for kp in kps:
        m = {}
        for s in students:
            cell = res.get((s, kp))
            if not cell or cell["n"] == 0:
                continue
            br = find_break(tree, kp, p_lookup(res, s))
            if br and br["break"]:
                m[s] = br["break"]["kp"]
        if m:
            per_kp[kp] = m
    nodes = {}           # 断点 -> {"students": set, "kps": set}
    for kp, m in per_kp.items():
        for s, node in m.items():
            d = nodes.setdefault(node, {"students": set(), "kps": set()})
            d["students"].add(s)
            d["kps"].add(kp)
    rows = [{"node": node, "n": len(d["students"]), "students": d["students"],
             "kps": sorted(d["kps"])} for node, d in nodes.items()]
    rows.sort(key=lambda r: (-r["n"], r["node"]))
    for r in rows:
        resid = {}
        for kp in r["kps"]:
            for s, node in per_kp.get(kp, {}).items():
                if node == r["node"]:
                    continue
                resid.setdefault(node, set()).add(s)
        r["residual"] = sorted(((len(v), k) for k, v in resid.items()), key=lambda x: (-x[0], x[1]))
    return rows


def bkt_step(p, ok, p_transit, p_guess, p_slip):
    """一步 BKT：先按观测更新后验 P(L|obs)，再做「学习」迁移。公式与口径一致。"""
    if ok:
        num_ = p * (1 - p_slip)
        den = num_ + (1 - p) * p_guess
    else:
        num_ = p * p_slip
        den = num_ + (1 - p) * (1 - p_guess)
    post = (num_ / den) if den > 1e-12 else p
    post = min(max(post, 0.0), 1.0)
    return post + (1 - post) * p_transit


def check_params(p_init, p_transit, p_guess, p_slip):
    for name, v in (("p-init", p_init), ("p-transit", p_transit),
                    ("p-guess", p_guess), ("p-slip", p_slip)):
        if not (0.0 <= v < 1.0):
            raise TrackerError("--%s 必须在 [0,1) 之间，现在是 %s" % (name, v))
    if p_transit <= 0:
        raise TrackerError("--p-transit 必须 > 0，否则「练了也不长」，边际收益公式无意义")
    return {"init": p_init, "transit": p_transit, "guess": p_guess, "slip": p_slip}


def build_observations(exams):
    """把多场考试拼成「学生×知识点」的有序观测序列。

    顺序：先按考试次序，同一次考试内按题号自然序。返回 {(学生, 知识点): [obs...]}。
    """
    obs = {}
    for ei, exam in enumerate(exams):
        for r in exam["rows"]:
            obs.setdefault((r["student"], r["kp"]), []).append({
                "exam": ei, "exam_no": ei + 1, "no": r["no"], "kp": r["kp"],
                "full": r["full"], "got": r["got"], "ok": r["ok"],
                "label": exam["label"],
            })
    for key, seq in obs.items():
        seq.sort(key=lambda o: (o["exam"], natkey(o["no"])))
    return obs


def run_bkt(obs, params, exam_days=None, tau=DECAY_TAU_DAYS):
    """对每条 (学生×知识点) 序列跑一遍 BKT。

    给了 exam_days（--dates）就先做口径补充 1 的遗忘衰减：
        P_before = p_init + (P_last - p_init) · exp(-d / τ)
    d 是这两次作答相隔的天数；同一天的多道题 d=0，等于不衰减。不给日期就完全不衰减。
    """
    out = {}
    for key, seq in obs.items():
        p = params["init"]
        curve = []
        prev_day = None
        for o in seq:
            day = exam_days[o["exam"]] if exam_days else None
            gap = (day - prev_day).days if (day is not None and prev_day is not None) else 0
            p_last = p
            if gap > 0:
                p = params["init"] + (p - params["init"]) * math.exp(-gap / float(tau))
            p = bkt_step(p, o["ok"], params["transit"], params["guess"], params["slip"])
            curve.append({"step": len(curve) + 1, "exam_no": o["exam_no"], "label": o["label"],
                          "no": o["no"], "got": o["got"], "full": o["full"],
                          "ok": o["ok"], "P": p, "gap": gap, "P_last": p_last})
            prev_day = day
        out[key] = {"student": key[0], "kp": key[1], "P": p, "n": len(seq), "curve": curve}
    return out


def decay_stats(res):
    """跨场回落了几次、累计隔了多少天 —— 写进质检，让老师看得见衰减到底生效没有。"""
    events = gaps = 0
    for cell in res.values():
        for c in cell["curve"]:
            if c.get("gap", 0) > 0:
                events += 1
                gaps += c["gap"]
    return events, gaps


def conf_label(n):
    """口径的置信度分档。"""
    if n <= 0:
        return "还没练过"
    if n < N_LOW_CONF:
        return "低置信"
    return "正常"


def practice_needed(p, p_target, p_transit):
    """口径公式：n = ceil( log((1-P*)/(1-P)) / log(1-transit) )，P>=P* 时为 0。"""
    if p >= p_target:
        return 0
    ratio = (1.0 - p_target) / (1.0 - p)
    return max(0, int(math.ceil(math.log(ratio) / math.log(1.0 - p_transit))))


def kp_weights(exams, bank_kps):
    """各知识点的卷面分值 W：最近一次考试该点满分之和；这一次没考就用题库，标注「按题库估」。"""
    last = exams[-1]
    w_last, w_bank, src = {}, dict(bank_kps), {}
    for r in last["rows"]:
        w_last[r["kp"]] = w_last.get(r["kp"], 0.0) + r["full"] / max(1, last["qc"]["students"])
    kps = sorted(set(list(w_last) + list(w_bank)))
    out = {}
    for kp in kps:
        if kp in w_last and w_last[kp] > 0:
            out[kp] = (w_last[kp], "最近一次卷面")
            src[kp] = "最近一次卷面"
        else:
            out[kp] = (w_bank.get(kp, 0.0), "按题库估")
            src[kp] = "按题库估"
    return {k: v[0] for k, v in out.items()}, src


def bank_kp_full(bank):
    acc = {}
    for it in bank["items"]:
        acc[it["kp"]] = acc.get(it["kp"], 0.0) + it["full"]
    return acc


# ══════════════════════════════════════════════════════════════════════════
# 子命令 1：mastery
# ══════════════════════════════════════════════════════════════════════════
def cmd_mastery(args):
    exams = require_exams(args.scores, 1, "知识追踪")
    params = check_params(args.p_init, args.p_transit, args.p_guess, args.p_slip)
    tau = getattr(args, "tau", DECAY_TAU_DAYS)
    exam_days = parse_dates(getattr(args, "dates", None), len(exams), tau)
    decay_note = decay_line(exam_days, tau)
    tree, tree_warn = load_tree(getattr(args, "tree", None) or TREE)
    print("=" * 78)
    print("第 3 层 · 知识追踪（BKT）：%d 次考试" % len(exams))
    print("=" * 78)
    print("  参数：p_init=%.2f  p_transit=%.2f  p_guess=%.2f  p_slip=%.2f（不做 EM 拟合）"
          % (params["init"], params["transit"], params["guess"], params["slip"]))
    for e in exams:
        print_qc(e)
    data_premise(exams, decay_line=decay_note)
    if tree_warn:
        print("    → %s" % tree_warn)

    bank = load_bank(args.bank or BANK)
    W_bank = bank_kp_full(bank)
    obs = build_observations(exams)
    res = run_bkt(obs, params, exam_days, tau)
    students = sorted({k[0] for k in res})
    kps = sorted({k[1] for k in res})

    if args.kp and args.kp not in kps:
        raise TrackerError("这份数据里没有知识点「%s」。有的：%s" % (args.kp, "、".join(kps)))
    if args.student and args.student not in students:
        raise TrackerError("这份数据里没有学生「%s」。有的：%s" % (args.student, "、".join(students)))

    # 全班视角：每个知识点的人均掌握度 + n 的分档人数
    kp_rows, sheets_curve = [], []
    for kp in kps:
        buckets = {"zero": 0, "low": 0, "ok": 0}
        ps = []
        for s in students:
            cell = res.get((s, kp))
            n = cell["n"] if cell else 0
            if n == 0:
                buckets["zero"] += 1
                continue
            ps.append(cell["P"])
            buckets["low" if n < N_LOW_CONF else "ok"] += 1
        avg = sum(ps) / len(ps) if ps else None
        kp_rows.append({"kp": kp, "avg": avg, "n_stu": len(ps), "zero": buckets["zero"],
                        "low": buckets["low"], "ok": buckets["ok"],
                        "W": W_bank.get(kp, 0.0)})
    kp_rows.sort(key=lambda x: (x["avg"] is None, x["avg"] if x["avg"] is not None else 9))

    print()
    print("  一、全班知识点掌握度（n=0 的人不出概率，所以人均只算练过的人）")
    print("     %-18s %-8s %-6s %-6s %-6s %-6s %s" % (
        "知识点", "人均掌握", "样本数", "n=0", "n=1-2", "n≥3", "题库满分"))
    for r in kp_rows:
        print("     %-18s %-8s %-6d %-6d %-6d %-6d %.0f" % (
            r["kp"], "—（还没练过）" if r["avg"] is None else pct(r["avg"], 1),
            r["n_stu"], r["zero"], r["low"], r["ok"], r["W"]))
    tot_zero = sum(r["zero"] for r in kp_rows)
    tot_low = sum(r["low"] for r in kp_rows)
    tot_ok = sum(r["ok"] for r in kp_rows)
    print("     n 分布合计：n=0 共 %d 个「学生×知识点」｜n=1-2 共 %d 个（低置信）｜n≥3 共 %d 个"
          % (tot_zero, tot_low, tot_ok))

    # 学生视角
    stu_rows, verdicts, curve_rows = [], [], []
    focus = [args.student] if args.student else []
    if focus:
        s = focus[0]
        print()
        print("  二、%s 的知识点掌握度（%d 个知识点）" % (s, len(kps)))
        print("     %-18s %-4s %-8s %-8s %-6s %s" % ("知识点", "n", "掌握概率", "置信", "分值W", "说明"))
        cells = []
        for kp in kps:
            cell = res.get((s, kp)) or {"P": None, "n": 0, "curve": []}
            n = cell["n"]
            gain = 0.0 if cell["P"] is None else max(0.0, (P_TARGET - cell["P"]) * W_bank.get(kp, 0.0))
            cells.append((kp, cell, gain))
            print("     %-18s %-4d %-8s %-8s %-6.0f %s" % (
                kp, n, "—" if cell["P"] is None else pct(cell["P"], 1), conf_label(n),
                W_bank.get(kp, 0.0),
                "还没练过，不出概率" if n == 0 else (
                    "观测 %d 条，样本薄" % n if n < N_LOW_CONF else "")))
            if n:
                stu_rows.append([s, kp, n, round(cell["P"], 3), conf_label(n),
                                 round(W_bank.get(kp, 0.0), 1)])
                for c in cell["curve"]:
                    curve_rows.append([s, kp, c["step"], c["exam_no"], c["label"], c["no"],
                                       c["got"], c["full"], "对" if c["ok"] else "错",
                                       round(c["P"], 4)])
        # 修 bug（不改进口径）：原来的第二排序键是 cell(dict)，两条「可提分」并列时会
        # 拿 dict 比大小 -> TypeError: '<' not supported。改成按知识点名兜底，排序结果不变。
        ranked = sorted([(kp, cell, g) for kp, cell, g in cells if cell["n"] > 0],
                        key=lambda x: (-x[2], x[0]))
        print()
        print("  三、结论：该生最该补的 2 个知识点（按 (P*−P)×分值W 从大到小，P*=%.2f）" % P_TARGET)
        for i, (kp, cell, g) in enumerate(ranked[:2], 1):
            print("     %d) %s：掌握 %s（观测 %d 条，%s），分值 %.0f → 可提 %.1f 分"
                  % (i, kp, pct(cell["P"], 1), cell["n"], conf_label(cell["n"]),
                     W_bank.get(kp, 0.0), g))
            verdicts.append(["%d" % i, kp, round(cell["P"], 3), cell["n"], conf_label(cell["n"]),
                             round(W_bank.get(kp, 0.0), 1), round(g, 1)])
        if len(ranked) < 2:
            print("     只有 %d 个知识点有观测，凑不出两个；另一个先按「还没练过」记，不编概率。"
                  % len(ranked))

        # ---- 口径补充 v2：前置链断点 + 不一致标记（以下全是新增行，上面一行没动）----
        if tree:
            def p_of(name):
                c2 = res.get((s, name))
                return c2["P"] if c2 else None
            print()
            print("  【前置链检查 · 口径补充 v2】判断线：前置 P<%.2f 且本点 P<%.2f 才算本点补不动"
                  % (P_PREREQ_LOW, P_TARGET))
            shown = 0
            for kp, cell, g in ranked:
                br = find_break(tree, kp, p_of)
                if br and br["break"]:
                    b = br["break"]
                    n_self = practice_needed(cell["P"], P_TARGET, params["transit"])
                    n_pre = practice_needed(b["P"], P_TARGET, params["transit"])
                    print("     · %s：先补前置『%s』（P=%.2f），否则这个点练了也站不住；"
                          "可行练习数 = 本点 %d + 前置 %d = %d 次"
                          % (kp, b["kp"], b["P"], n_self, n_pre, n_self + n_pre))
                    shown += 1
                elif br and br["unknown"]:
                    print("     · %s：前置『%s』还没有观测，判不了断点（按未断处理，不臆造概率）"
                          % (kp, "、".join(br["unknown"])))
                    shown += 1
            if not shown:
                print("     · 没有断点：前置都通（P ≥ %.2f）或本点已达标，直接练这个点即可。"
                      % P_PREREQ_LOW)
            hits = 0
            for kp, cell, g in ranked:
                w = inconsistency(tree, kp, p_of)
                if w:
                    print("     ⚠ 可疑：前置「%s」掌握 %.2f，但「%s」已 %.2f"
                          % (w["prereq"], w["p_prereq"], w["kp"], w["p_kp"]))
                    print("       可能原因：① 蒙对  ② 该点题目过易（低于门槛）  ③ 前置判断有误")
                    hits += 1
                    if hits >= 2:
                        break
            if hits:
                print("       注：按口径补充 3 只做提示，不改概率、不做硬性压制。")
    else:
        # 全班跑：每个知识点给一条「班级平均曲线」（写进同一个 sheet）
        weak = [r for r in kp_rows if r["avg"] is not None][:3]
        for r in weak:
            kp = r["kp"]
            seq_len = max((res[(s, kp)]["n"] for s in students if (s, kp) in res), default=0)
            for step in range(1, seq_len + 1):
                vals = []
                for s in students:
                    cell = res.get((s, kp))
                    if cell and cell["n"] >= step:
                        vals.append(cell["curve"][step - 1]["P"])
                if vals:
                    curve_rows.append(["（全班）", kp, step, "", "", "", "", "", "",
                                       round(sum(vals) / len(vals), 4)])
        print()
        print("  二、掌握曲线：已写进 xlsx 的「掌握曲线」sheet（全班跑取下限最大的 3 个知识点，"
              "按步数给班级平均 P；要某个人的逐条曲线请加 --student）")
        for r in weak:
            print("     %s：班级平均 %s（%d 人练过）" % (r["kp"], pct(r["avg"], 1), r["n_stu"]))

        # ---- 口径补充 v2：全班前置瓶颈（只报 ≥CLASS_PREREQ_MIN 人的）----
        if tree:
            pres = class_prereq_block(tree, kps, students, res)
            top = [r for r in pres if r["n"] >= CLASS_PREREQ_MIN]
            print()
            print("  【全班前置瓶颈 · 口径补充 v2】只报 ≥%d 人的（1~2 人属于个体辅导，不是班级决策）"
                  % CLASS_PREREQ_MIN)
            shown = top[:3]
            shown_nodes = {r["node"] for r in shown}
            for r in shown:
                kp_list = r["kps"]
                brief = "、".join(kp_list[:3]) + ("等" if len(kp_list) > 3 else "")
                line = ("     · %d 人卡在『%s』这个前置上，影响 %d 个知识点（%s）——这节课之前"
                        "先花 5 分钟集体过一遍%s"
                        % (r["n"], r["node"], len(kp_list), brief, r["node"]))
                # 次要断点：过完这一遍还剩多少人卡在别的前置；≥2 人才提，1 人是课后单独辅导
                resid = [x for x in r["residual"] if x[1] not in shown_nodes and x[0] >= 2]
                if resid:
                    line += "（另有 %d 人卡在『%s』）" % (resid[0][0], resid[0][1])
                print(line)
            if not top:
                if pres:
                    r = pres[0]
                    print("     · 没有断点达到 %d 人：最多的是『%s』%d 人（影响 %d 个知识点：%s），"
                          "按个体辅导处理，不占用课堂时间。"
                          % (CLASS_PREREQ_MIN, r["node"], r["n"], len(r["kps"]),
                             "、".join(r["kps"][:3])))
                else:
                    print("     · 没有学生因为前置断而补不动：前置都通，按上面的知识点顺序练即可。")

    note = ("BKT 四参数 p_init=%.2f p_transit=%.2f p_guess=%.2f p_slip=%.2f；不做 EM 拟合。"
            "观测：得分==满分才算答对。" % (params["init"], params["transit"],
                                     params["guess"], params["slip"]))
    path = out_path("知识追踪_%s" % (args.student or "全班"))
    sheets = [("知识点汇总", ["知识点", "人均掌握概率", "样本学生数", "n=0 人数", "n=1-2 人数",
                          "n≥3 人数", "题库满分"], [
        [r["kp"], "" if r["avg"] is None else round(r["avg"], 3), r["n_stu"], r["zero"],
         r["low"], r["ok"], round(r["W"], 1)] for r in kp_rows], note)]
    if stu_rows:
        sheets.append(("该生知识点", ["学生", "知识点", "观测条数", "掌握概率", "置信", "分值W"],
                       stu_rows, note))
    if curve_rows:
        sheets.append(("掌握曲线", ["学生", "知识点", "第几次作答", "第几次考试", "场次", "题号",
                                "得分", "满分", "对错", "更新后掌握概率"], curve_rows, note))
    if verdicts:
        sheets.append(("补强结论", ["优先级", "知识点", "掌握概率", "观测条数", "置信", "分值W",
                                "可提分数"], verdicts, "按 (P*−P)×W 排序，P*=%.2f" % P_TARGET))
    sheets.append(("数据质检", ["项", "值"], [
        ["考试次数", "%d（%s）" % (len(exams), " → ".join(e["label"] for e in exams))],
        ["各次有效行", "、".join("%s:%d" % (e["label"], e["qc"]["rows_ok"]) for e in exams)],
        ["部分得分", "、".join("%s:%d 行" % (e["label"], e["qc"]["partial"]) for e in exams)],
        ["n 分布合计", "n=0:%d｜n=1-2:%d｜n≥3:%d" % (tot_zero, tot_low, tot_ok)],
        ["口径", note],
        ["样本不足怎么办", "n=0 不出概率标「还没练过」；n=1~2 标「低置信」；只有 1 次考试明说趋势做不了"],
        ["遗忘衰减", decay_note],
        ["衰减统计", "跨场回落 %d 次，累计间隔 %d 天" % decay_stats(res)],
        ["前置链", tree_warn or ("已按知识点树 prereq 判断前置断点；多个断点只报最靠前的那个"
                                "（口径补充 v2 补充 2/3）")],
    ], None))
    write_xlsx(path, sheets)
    print()
    print("  → 已落表：%s" % path)
    return 0


# ══════════════════════════════════════════════════════════════════════════
# 子命令 2：gain
# ══════════════════════════════════════════════════════════════════════════
def cmd_gain(args):
    exams = require_exams(args.scores, 1, "边际收益排序")
    params = check_params(args.p_init, args.p_transit, args.p_guess, args.p_slip)
    tau = getattr(args, "tau", DECAY_TAU_DAYS)
    exam_days = parse_dates(getattr(args, "dates", None), len(exams), tau)
    decay_note = decay_line(exam_days, tau)
    tree, tree_warn = load_tree(getattr(args, "tree", None) or TREE)
    print("=" * 78)
    print("第 4 层 · 边际收益排序：可提分数 ÷ 练习成本")
    print("=" * 78)
    print("  参数：P*=%.2f  p_transit=%.2f  单题成本 %.0f 分钟（含讲评与订正）"
          % (P_TARGET, params["transit"], ITEM_MINUTES))
    for e in exams:
        print_qc(e)
    data_premise(exams, decay_line=decay_note)
    if tree_warn:
        print("    → %s" % tree_warn)
    if len(exams) < TREND_MIN_EXAMS:
        print("    → 按口径只出**单次边际收益**，不出趋势（趋势要 ≥%d 次考试）。" % TREND_MIN_EXAMS)

    bank = load_bank(args.bank or BANK)
    W, W_src = kp_weights(exams, bank_kp_full(bank))
    obs = build_observations(exams)
    res = run_bkt(obs, params, exam_days, tau)
    students = sorted({k[0] for k in res})
    kps = sorted({k[1] for k in res})
    if args.student and args.student not in students:
        raise TrackerError("这份数据里没有学生「%s」。有的：%s" % (args.student, "、".join(students)))

    def cell_of(s, kp):
        return res.get((s, kp)) or {"P": None, "n": 0}

    # ---- 个体视角 ----
    stu_out = []
    if args.student:
        s = args.student

        def p_of(name):
            c2 = res.get((s, name))
            return c2["P"] if c2 else None

        rows = []
        for kp in kps:
            cell = cell_of(s, kp)
            if cell["n"] == 0:
                continue
            p = cell["P"]
            n_self = practice_needed(p, P_TARGET, params["transit"])
            gain = max(0.0, (min(P_TARGET, 1.0) - p)) * W.get(kp, 0.0)
            # 口径补充 v2 补充 2 + 对第 4 层的影响：前置断了，成本按「前置补通之后」重算
            br = find_break(tree, kp, p_of) if tree else None
            fix = None
            n_need = n_self
            if br and br["break"]:
                b = br["break"]
                n_pre = practice_needed(b["P"], P_TARGET, params["transit"])
                fix = {"kp": b["kp"], "P": b["P"], "n_pre": n_pre, "n_self": n_self,
                       "n_total": n_self + n_pre}
                n_need = n_self + n_pre
            hours = n_need * ITEM_MINUTES / 60.0
            per_hour = (gain / hours) if hours > 0 else None
            rows.append({"kp": kp, "P": p, "n": cell["n"], "conf": conf_label(cell["n"]),
                         "W": W.get(kp, 0.0), "W_src": W_src.get(kp, ""),
                         "n_need": n_need, "n_self": n_self, "fix": fix,
                         "gain": gain, "hours": hours, "per_hour": per_hour})
        done = [r for r in rows if r["n_need"] == 0]
        todo = [r for r in rows if r["n_need"] > 0]
        todo.sort(key=lambda r: (-(r["per_hour"] or 0), r["kp"]))
        print()
        print("  一、%s 的边际收益（按每小时收益从高到低）" % s)
        print("     %-18s %-8s %-4s %-8s %-6s %-9s %-8s %s" % (
            "知识点", "掌握P", "n", "置信", "可提分", "要练几道", "成本", "每小时收益"))
        for r in todo:
            print("     %-18s %-8s %-4d %-8s %-6.1f %-9d %-8s %s" % (
                r["kp"], pct(r["P"], 1), r["n"], r["conf"], r["gain"], r["n_need"],
                fmt_hours(r["hours"]),
                "—" if r["per_hour"] is None else "%.1f 分/h" % r["per_hour"]))
            if r["fix"]:
                print("       ↑ 先补前置『%s』（P=%.2f），否则这个点练了也站不住；成本已按"
                      "「前置补通之后」重算：本点 %d 道 + 前置 %d 道 = %d 道"
                      % (r["fix"]["kp"], r["fix"]["P"], r["fix"]["n_self"],
                         r["fix"]["n_pre"], r["fix"]["n_total"]))
            if r["n_need"] > PRACTICE_WARN:
                print("       ↑ 要练 %d 道（>%d）已不现实：他不是练得少，得换讲法或退到更基础的题"
                      % (r["n_need"], PRACTICE_WARN))
        if done:
            print("     已达标（P ≥ %.2f，不用再投课时）：%s" % (
                P_TARGET, "、".join("%s(%s)" % (r["kp"], pct(r["P"], 1)) for r in done)))
        print()
        print("  二、结论：这个学生先补这 2 个（每条都同时给收益和成本）")
        for i, r in enumerate(todo[:2], 1):
            print("     %d) %s：可提 %.1f 分｜要练 %d 道（约 %s）｜每小时收益 %s｜"
                  "当前掌握 %s（观测 %d 条，%s，分值 %.0f 按%s）"
                  % (i, r["kp"], r["gain"], r["n_need"], fmt_hours(r["hours"]),
                     "—" if r["per_hour"] is None else "%.1f 分/h" % r["per_hour"],
                     pct(r["P"], 1), r["n"], r["conf"], r["W"], r["W_src"]))
            if r["fix"]:
                print("        → 先补前置『%s』（P=%.2f），否则这个点练了也站不住"
                      "（要练 %d 道 = 本点 %d + 前置 %d）"
                      % (r["fix"]["kp"], r["fix"]["P"], r["fix"]["n_total"],
                         r["fix"]["n_self"], r["fix"]["n_pre"]))
            elif r["n_self"] > 0:
                print("        → 直接练这个点，需 %d 次" % r["n_self"])
            stu_out.append(["%d" % i, s, r["kp"], round(r["P"], 3), r["n"], r["conf"],
                            round(r["W"], 1), r["W_src"], round(r["gain"], 1),
                            r["n_need"], round(r["hours"], 2),
                            "" if r["per_hour"] is None else round(r["per_hour"], 2)])
        if len(todo) < 2:
            print("     只有 %d 个知识点还需要练；不够两条就如实只给 %d 条，不硬凑。"
                  % (len(todo), len(todo)))

        # ---- 口径补充 v2：前置链汇总 + 不一致标记（以下全是新增行）----
        if tree:
            print()
            print("  【前置链检查 · 口径补充 v2】前置 P<%.2f 且本点 P<%.2f 才算断点，只报最靠前的"
                  % (P_PREREQ_LOW, P_TARGET))
            brs = [(r["kp"], r["fix"]) for r in todo if r.get("fix")]
            for kp, f in brs[:3]:
                print("     · %s：先补前置『%s』（P=%.2f），否则这个点练了也站不住"
                      % (kp, f["kp"], f["P"]))
            if not brs:
                print("     · 没有断点：前置都通（P ≥ %.2f），按上面的成本直接练即可。"
                      % P_PREREQ_LOW)
            hits = 0
            for kp in kps:
                w = inconsistency(tree, kp, p_of)
                if w:
                    print("     ⚠ 可疑：前置「%s」掌握 %.2f，但「%s」已 %.2f"
                          % (w["prereq"], w["p_prereq"], w["kp"], w["p_kp"]))
                    print("       可能原因：① 蒙对  ② 该点题目过易（低于门槛）  ③ 前置判断有误")
                    hits += 1
                    if hits >= 2:
                        break
            if hits:
                print("       注：按口径补充 3 只做提示，不改概率、不压制。")

    # ---- 全班视角 ----
    rows = []
    for kp in kps:
        gain_sum, n_sum, people = 0.0, 0, 0
        for s in students:
            cell = cell_of(s, kp)
            if cell["n"] == 0:
                continue
            n_need = practice_needed(cell["P"], P_TARGET, params["transit"])
            if n_need <= 0:
                continue
            gain_sum += max(0.0, (P_TARGET - cell["P"])) * W.get(kp, 0.0)
            n_sum += n_need
            people += 1
        hours = n_sum * ITEM_MINUTES / 60.0
        rows.append({"kp": kp, "gain": gain_sum, "n": n_sum, "people": people,
                     "hours": hours,
                     "per_hour": (gain_sum / hours) if hours > 0 else None,
                     "W": W.get(kp, 0.0), "W_src": W_src.get(kp, "")})
    rows.sort(key=lambda r: (-(r["per_hour"] or 0), r["kp"]))
    print()
    print("  三、全班视角：先补哪 3 个知识点（按每小时收益排序；成本＝全班要练的总题数）")
    print("     %-18s %-4s %-9s %-8s %-8s %-6s %s" % (
        "知识点", "人数", "可提总分", "总练习数", "总成本", "分值W", "每小时收益"))
    for r in rows[:3]:
        print("     %-18s %-4d %-9.1f %-8d %-8s %-6.0f %s" % (
            r["kp"], r["people"], r["gain"], r["n"], fmt_hours(r["hours"]), r["W"],
            "—" if r["per_hour"] is None else "%.1f 分/h" % r["per_hour"]))
    top3 = [r for r in rows[:3] if r["gain"] > 0]
    if top3:
        print("     合计：这 %d 个知识点可提 %.1f 分，全班要练 %d 道（约 %s）"
              % (len(top3), sum(r["gain"] for r in top3), sum(r["n"] for r in top3),
                 fmt_hours(sum(r["hours"] for r in top3))))
    else:
        print("     全班没有「还没达标」的知识点：不用再投课时，按现有讲法做复习即可。")
    print("     注意：可提总分是「每个学生各自补到 %.2f」的加总，不是一次考试的分差。" % P_TARGET)
    if len(exams) < TREND_MIN_EXAMS:
        print("     只有 %d 次数据：上面是单次口径，趋势做不了（口径要求 ≥%d 次）。"
              % (len(exams), TREND_MIN_EXAMS))

    note = ("当前掌握 P 来自 BKT（p_init=%.2f p_transit=%.2f p_guess=%.2f p_slip=%.2f，不做 EM）；"
            "n=ceil(log((1-P*)/(1-P))/log(1-transit))，P*=%.2f；可提分=(P*−P)×W；"
            "成本=%d 道/人 × %.0f 分钟" % (params["init"], params["transit"], params["guess"],
                                      params["slip"], P_TARGET, 1, ITEM_MINUTES))
    path = out_path("边际收益_%s" % (args.student or "全班"))
    detail = [[r["kp"], round(r["W"], 1), r["W_src"], r["people"], round(r["gain"], 1),
               r["n"], round(r["hours"], 2),
               "" if r["per_hour"] is None else round(r["per_hour"], 2)] for r in rows]
    sheets = [("全班TOP3", ["知识点", "人数", "可提总分", "总练习数", "总成本(h)", "分值W", "每小时收益"],
               [[r["kp"], r["people"], round(r["gain"], 1), r["n"], round(r["hours"], 2),
                 round(r["W"], 1),
                 "" if r["per_hour"] is None else round(r["per_hour"], 2)] for r in rows[:3]],
               note),
              ("知识点成本收益", ["知识点", "卷面分值W", "W来源", "涉及人数", "可提总分",
                            "总练习数", "总成本(h)", "每小时收益"], detail, note)]
    if stu_out:
        sheets.insert(0, ("个体建议", ["优先级", "学生", "知识点", "掌握P", "观测n", "置信",
                                  "分值W", "W来源", "可提分数", "要练几道", "成本(h)", "每小时收益"],
                         stu_out, note))
    sheets.append(("数据质检", ["项", "值"], [
        ["考试次数", "%d（%s）" % (len(exams), " → ".join(e["label"] for e in exams))],
        ["各次有效行", "、".join("%s:%d" % (e["label"], e["qc"]["rows_ok"]) for e in exams)],
        ["W 来源", "、".join("%s:%s" % (k, v) for k, v in sorted(W_src.items()))],
        ["成本口径", "一道同点变式题 %.0f 分钟（含讲评与订正）；宁可高估成本不虚报收益"
         % ITEM_MINUTES],
        ["样本不足怎么办", "n=0 的知识点不进排序（还没练过就没有 P）；只有 1~2 次考试明说趋势做不了"],
        ["遗忘衰减", decay_note],
        ["衰减统计", "跨场回落 %d 次，累计间隔 %d 天" % decay_stats(res)],
        ["前置链", tree_warn or ("已按知识点树 prereq 判断前置断点；前置断时成本 = 本点所需 + 前置所需"
                                "（口径补充 v2 补充 2 / 对第 4 层的影响）")],
    ], None))
    write_xlsx(path, sheets)
    print()
    print("  → 已落表：%s" % path)
    return 0


# ══════════════════════════════════════════════════════════════════════════
# 子命令 3：anchor
# ══════════════════════════════════════════════════════════════════════════
def cmd_anchor(args):
    files = list(args.files or [])
    exams = require_exams(files, EXAM_MIN_ANCHOR, "等值化")
    anchors = []
    if args.anchor:
        anchors = [a.strip() for a in str(args.anchor).replace("，", ",").split(",") if a.strip()]
    else:
        common = set(exams[0]["rows"][0] and {r["no"] for r in exams[0]["rows"]})
        for e in exams[1:]:
            common &= {r["no"] for r in e["rows"]}
        anchors = sorted(common, key=natkey)
        allnos = {r["no"] for r in exams[0]["rows"]}
        if len(anchors) == len(allnos):
            raise TrackerError(
                "没给 --anchor，而 %d 次考试的题号完全一样（%d 道）——那等于拿整张卷当锚题，"
                "算出来的 k 没有意义。请用 --anchor 指定真正跨考试稳定的题号，例如 --anchor 8,12,6。"
                % (len(exams), len(allnos)))
    if len(anchors) < ANCHOR_MIN:
        raise TrackerError("锚题只有 %d 道，少于下限 %d 道——**不做等值化**，直接说「可比性不足」。"
                           "跨考试对不齐，硬算出来的 k 是假的。" % (len(anchors), ANCHOR_MIN))

    print("=" * 78)
    print("第 4 层 · 等值化（跨考试可比）：锚题 %s" % "、".join(anchors))
    print("=" * 78)
    for e in exams:
        print_qc(e)

    # 每题在每次考试里的得分率（pooled），用于核对锚题是否每次都在
    stats = []
    missing = {}
    for e in exams:
        rows = e["rows"]
        a_rows = [r for r in rows if r["no"] in anchors]
        got = sum(r["got"] for r in a_rows)
        full = sum(r["full"] for r in a_rows)
        rate = got / full if full else 0.0
        miss = [a for a in anchors if a not in {r["no"] for r in rows}]
        if miss:
            missing[e["label"]] = miss
        tot_got = sum(r["got"] for r in rows)
        avg = tot_got / e["qc"]["students"] if e["qc"]["students"] else 0.0
        stats.append({"label": e["label"], "rate": rate, "avg": avg,
                      "full": e["qc"]["full_total"], "anchors": len(anchors) - len(miss)})
    if missing:
        raise TrackerError("锚题在以下场次里缺失：%s —— 对不齐就不能等值化。"
                           % "；".join("%s: %s" % (k, "、".join(v)) for k, v in missing.items()))

    base = stats[0]
    print()
    print("  锚题平均得分率：%s" % " ｜ ".join(
        "%s %.1f%%（%d 道）" % (s["label"], s["rate"] * 100, s["anchors"]) for s in stats))
    print()
    print("  %-22s %-8s %-9s %-9s %-10s %s" % (
        "场次", "原始均分", "锚题得分率", "k(相对第1次)", "环比 k", "可比均分(÷k)"))
    out_rows = []
    for i, s in enumerate(stats):
        k_cum = 1.0 if i == 0 else (s["rate"] / base["rate"] if base["rate"] else 0.0)
        k_prev = 1.0 if i == 0 else (s["rate"] / stats[i - 1]["rate"]
                                     if stats[i - 1]["rate"] else 0.0)
        comp = (s["avg"] / k_cum) if k_cum else s["avg"]
        s["k_cum"], s["k_prev"], s["comp"] = k_cum, k_prev, comp
        print("  %-22s %-8.1f %-9s %-9s %-10s %.1f" % (
            s["label"], s["avg"], "%.1f%%" % (s["rate"] * 100),
            "%.3f" % k_cum, "%.3f" % k_prev, comp))
        out_rows.append([s["label"], round(s["avg"], 1), round(s["rate"], 4), round(k_cum, 4),
                         round(k_prev, 4), round(comp, 1), s["anchors"]])
    print()
    print("  读法：k>1 说明这次锚题整体比第 1 次容易（分要除小一点才可比）；"
          "k<1 说明更难。可比均分 = 原始均分 ÷ k(相对第 1 次)。")
    if len(stats) >= 2:
        d_raw = stats[-1]["avg"] - stats[0]["avg"]
        d_comp = stats[-1]["comp"] - stats[0]["comp"]
        print("  第 1 次 → 最后一次：原始 %+.1f 分，可比 %+.1f 分（差 %.1f 分是卷面难度变化造成的）"
              % (d_raw, d_comp, d_raw - d_comp))
    print("  注意：口径要求 ≥%d 次考试且每次 ≥%d 道锚题；本次 %d 次 × %d 道锚题，达标。"
          % (EXAM_MIN_ANCHOR, ANCHOR_MIN, len(stats), len(anchors)))

    note = ("k(相对第1次)=本次锚题得分率/第1次锚题得分率；可比均分=原始均分÷k。"
            "锚题 %s；口径：锚题 <2 道不做等值化。" % "、".join(anchors))
    path = out_path("等值化_锚题%s" % "_".join(anchors))
    write_xlsx(path, [
        ("等值结果", ["场次", "原始均分", "锚题平均得分率", "k(相对第1次)", "环比k",
                  "可比均分", "锚题数"], out_rows, note),
        ("锚题明细", ["题号", "在场次数"], [[a, sum(1 for e in exams if a in {r["no"] for r in e["rows"]})]
                                     for a in anchors], note),
        ("数据质检", ["项", "值"], [
            ["考试次数", "%d（%s）" % (len(exams), " → ".join(e["label"] for e in exams))],
            ["各次有效行", "、".join("%s:%d" % (e["label"], e["qc"]["rows_ok"]) for e in exams)],
            ["锚题", "、".join(anchors) + "（%d 道，下限 %d）" % (len(anchors), ANCHOR_MIN)],
            ["口径", note],
        ], None),
    ])
    print()
    print("  → 已落表：%s" % path)
    return 0


# ══════════════════════════════════════════════════════════════════════════
def build_parser():
    p = argparse.ArgumentParser(
        description="教师工具 · 第 3/4 层：知识追踪（BKT）/ 边际收益排序 / 等值化（0 次模型调用）")
    sub = p.add_subparsers(dest="cmd")

    def add_bkt(sp):
        sp.add_argument("--p-init", dest="p_init", type=float, default=P_INIT,
                        help="p_init（默认 %.2f）" % P_INIT)
        sp.add_argument("--p-transit", dest="p_transit", type=float, default=P_TRANSIT,
                        help="p_transit（默认 %.2f）" % P_TRANSIT)
        sp.add_argument("--p-guess", dest="p_guess", type=float, default=P_GUESS,
                        help="p_guess（默认 %.2f）" % P_GUESS)
        sp.add_argument("--p-slip", dest="p_slip", type=float, default=P_SLIP,
                        help="p_slip（默认 %.2f）" % P_SLIP)
        sp.add_argument("--bank", default=None, help="题库.json（默认 教师工具/题库.json）")
        sp.add_argument("--dates", nargs="+", default=None, metavar="YYYY-MM-DD",
                        help="每次考试的日期，与 --scores 一一对应（如 2026-09-01 2026-10-01 "
                             "2026-11-01）。给了才做遗忘衰减；不给就明说「本次未做衰减」")
        sp.add_argument("--tau", type=float, default=DECAY_TAU_DAYS,
                        help="遗忘衰减时间常数 τ（天，默认 %d）" % DECAY_TAU_DAYS)
        sp.add_argument("--tree", default=None,
                        help="知识点树.json（默认 教师工具/teaching_out/知识点树.json）")

    a = sub.add_parser("mastery", help="知识追踪：每个学生×知识点的掌握概率")
    a.add_argument("--scores", nargs="+", required=True, help="逐题得分，按时间顺序")
    a.add_argument("--student", default=None, help="只看这个学生")
    a.add_argument("--kp", default=None, help="只看这个知识点")
    add_bkt(a)
    a.set_defaults(func=cmd_mastery)

    b = sub.add_parser("gain", help="边际收益：可提分数 + 练习成本，按每小时收益排序")
    b.add_argument("--scores", nargs="+", required=True, help="逐题得分，按时间顺序")
    b.add_argument("--student", default=None, help="个体视角")
    add_bkt(b)
    b.set_defaults(func=cmd_gain)

    c = sub.add_parser("anchor", help="等值化：用锚题把不同次考试的分数对齐")
    c.add_argument("--files", nargs="+", required=True, help="逐题得分，按时间顺序")
    c.add_argument("--anchor", default=None, help="锚题号，逗号分隔，例如 8,12,6")
    c.set_defaults(func=cmd_anchor)
    return p


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)
    if not getattr(args, "cmd", None):
        parser.print_help()
        return 2
    try:
        return args.func(args)
    except TrackerError as e:
        print("！%s" % e)
        return 2


if __name__ == "__main__":
    sys.exit(main())
