# -*- coding: utf-8 -*-
"""教师工具 · 第 5 层：单班 A/B 实验工具（只做证据整理，不做因果推断）

    python abtest.py plan   --name 移项用口诀还是数轴 --kp 一元一次方程移项 \
                            --group-a 讲法A:12人 --group-b 讲法B:11人 \
                            --metric 同类型题正确率 --assign 按上次月考名次蛇形分组
    python abtest.py result --name 移项用口诀还是数轴 --scores 各组得分.csv
    python abtest.py followup --name 移项用口诀还是数轴 --text "第二周又测了一次…"
    python abtest.py ledger

三层纪律（口径见 深度分析口径.md 第 5 节）：
  1. 事前登记，登记块冻结（sha256），事后只能追加 followup，改不动原来写的字
  2. 结果只报「本次观察到 A 组均值比 B 组高 X 分（波动范围 ±Y）」，不做任何因果判断
  3. 禁语闸门：证明/显著/有效/提升/因为/所以/必然 —— 命中就抛错，宁可不落表

0 次大模型调用：所有数字都是本地算的。
"""
import argparse
import csv
import datetime
import hashlib
import json
import math
import os
import random
import re
import statistics
import sys

try:                      # Windows 控制台默认 GBK，中文会乱码
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "_abtest")
QBANK = os.path.join(HERE, "题库.json")

# ── 禁语：第 5 层只描述「这次观察到什么」，任何评价效果/推断因果的词都不许出现。
# 词表和 insight.py 的 FORBID_WORDS（评价老师个人）不是一回事，各管一段，这里不 import 它。
FORBID_WORDS = ("证明", "显著", "有效", "提升", "因为", "所以", "必然")

BOOT_RESAMPLES = 10000
BOOT_SEED = 20261029      # 固定种子：同一个输入永远给同一组数字，便于老师复核
BOOT_ALPHA = 0.05
SMALL_SAMPLE = 10         # 每组少于这个人数就打醒目提示
FROZEN_FIELDS = ("kp", "groups", "metric", "assign", "expect", "registered_at")


class AbError(Exception):
    """给人看的错误：命令行里只打印一句话，不甩 traceback。"""


# ══════════════════════════════════════════════════════════════ 基础
def now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def canon(obj):
    """规范化序列化：算登记块指纹用，跨平台跨版本稳定。"""
    return json.dumps(obj, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def guard(*texts):
    """禁语闸门：命中就抛错。这一层的「诚实」靠它，不靠自觉。"""
    blob = " ".join(str(t) for t in texts if t)
    hits = [w for w in FORBID_WORDS if w in blob]
    if hits:
        raise AbError("输出文案里出现了这些词：%s —— 第 5 层只做证据整理，"
                      "写「本次观察到什么」就行；要评价效果请去别处，本工具不接。"
                      % "、".join(sorted(set(hits))))
    return len(blob)


def f2(x, nd=2):
    return "%.*f" % (nd, x)


def signed(x, nd=2):
    return ("+" if x >= 0 else "-") + ("%.*f" % (nd, abs(x)))


def slug_of(name):
    """实验名 → 目录名。中文原样保留，只把文件系统不认的字符换掉。"""
    s = re.sub(r'[\\/:*?"<>|\s]+', "_", str(name or "").strip()).strip("._")
    if not s:
        raise AbError("实验名不能是空的。")
    return s[:80]


def plan_path(slug):
    return os.path.join(ROOT, slug, "plan.json")


def result_path(slug):
    return os.path.join(ROOT, slug, "result.json")


def frozen_of(plan):
    return {k: plan.get(k) for k in FROZEN_FIELDS}


def frozen_hash(plan):
    return hashlib.sha256(canon(frozen_of(plan)).encode("utf-8")).hexdigest()


def load_plan(slug):
    p = plan_path(slug)
    if not os.path.exists(p):
        raise AbError("没找到这个实验：「%s」。先跑 plan 登记（当前目录：%s）。"
                      % (slug, os.path.join(ROOT, slug)))
    try:
        with open(p, encoding="utf-8") as f:
            plan = json.load(f)
    except Exception as e:
        raise AbError("plan.json 读不出来（%s）：%s" % (p, str(e)[:80]))
    if frozen_hash(plan) != plan.get("frozen_sha256"):
        raise AbError("「%s」的登记内容被手改过（指纹对不上）。%s\n"
                      "  登记块一旦写下就不许改 —— 事后改口径，整个实验就作废了。"
                      "\n  要补充观察请用 followup 子命令；要重做请换个新实验名。"
                      % (slug, p))
    return plan


def all_plans():
    if not os.path.isdir(ROOT):
        return []
    out = []
    for d in sorted(os.listdir(ROOT)):
        if d.startswith("_") or d.startswith("."):
            continue                      # _demo 之类的辅助目录不算实验
        p = os.path.join(ROOT, d, "plan.json")
        if os.path.exists(p):
            try:
                out.append(load_plan(d))
            except AbError as e:
                out.append({"_broken": str(e), "slug": d, "dir": d})
    return out


# ══════════════════════════════════════════════════════════════ plan
def parse_group(spec, key):
    """"讲法A:12人" → {"key": "A", "label": "讲法A", "n_declared": 12}
    人数可以不写；写了就登记进冻结块，事后实际人数不符会提示。"""
    s = str(spec or "").strip()
    if not s:
        raise AbError("--group-%s 不能为空。" % key.lower())
    label, n = s, None
    m = re.match(r"^(.*?)[:：]\s*([0-9]+)\s*人?\s*$", s)
    if m:
        label, n = m.group(1).strip(), int(m.group(2))
    if not label:
        raise AbError("--group-%s 的讲法名是空的：%r" % (key.lower(), spec))
    return {"key": key, "label": label, "n_declared": n}


def qbank_kps():
    """题库里的知识点（只读）。用来提醒 kp 是不是写错了，写错会导致台账合并错。"""
    try:
        with open(QBANK, encoding="utf-8") as f:
            data = json.load(f)
        return sorted({str(i.get("kp")).strip() for i in data.get("items", []) if i.get("kp")})
    except Exception:
        return []


def cmd_plan(a):
    slug = slug_of(a.name)
    incoming = {
        "kp": (a.kp or "").strip(),
        "groups": [parse_group(a.group_a, "A"), parse_group(a.group_b, "B")],
        "metric": (a.metric or "").strip(),
        "assign": (a.assign or "（未填写）").strip(),
        "expect": (a.expect or "（登记时没写预期）").strip(),
        "registered_at": now(),
    }
    if not incoming["kp"]:
        raise AbError("--kp 不能为空：知识点是台账合并的键，写错就没法横向比。")
    if not incoming["metric"]:
        raise AbError("--metric 不能为空：先说好量什么，再去做实验。")
    guard(incoming["kp"], incoming["metric"], incoming["assign"], incoming["expect"])

    p = plan_path(slug)
    if os.path.exists(p):
        old = load_plan(slug)
        diff = []
        for k in FROZEN_FIELDS:
            if k == "registered_at":
                continue
            if canon(old.get(k)) != canon(incoming.get(k)):
                diff.append("%s：已登记 %r → 这次传入 %r" % (k, old.get(k), incoming.get(k)))
        if not diff:
            print("实验「%s」已经登记过，内容完全一致，没有重复登记。" % slug)
            print("登记时间：%s ｜ 登记块指纹：%s" % (old["registered_at"], old["frozen_sha256"][:16]))
            print("plan.json：%s" % p)
            return 0
        raise AbError(
            "实验「%s」已经在 %s 登记过，**登记内容不可更改**（防止事后改口径找解释）。差异：\n  %s\n"
            "  要补充观察：abtest.py followup --name %s --text \"……\"\n"
            "  要重做：换一个新的实验名重新登记。"
            % (slug, old["registered_at"], "\n  ".join(diff), slug))

    plan = dict(incoming)
    plan["slug"] = slug
    plan["dir"] = slug
    plan["frozen_fields"] = list(FROZEN_FIELDS)
    plan["frozen_sha256"] = frozen_hash(plan)
    plan["followups"] = []
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(plan, f, ensure_ascii=False, indent=2)

    print("已登记（事前登记，写完即冻结）：%s" % plan["name"] if False else "")
    print("=" * 62)
    print("实验登记 ｜ %s" % slug)
    print("=" * 62)
    print("知识点　：%s" % plan["kp"])
    print("两组　　：%s（%s）vs %s（%s）"
          % (plan["groups"][0]["label"],
             ("登记 %d 人" % plan["groups"][0]["n_declared"]) if plan["groups"][0]["n_declared"] else "人数未写",
             plan["groups"][1]["label"],
             ("登记 %d 人" % plan["groups"][1]["n_declared"]) if plan["groups"][1]["n_declared"] else "人数未写"))
    print("评价指标：%s" % plan["metric"])
    print("分组方式：%s" % plan["assign"])
    print("预期效果：%s" % plan["expect"])
    print("登记时间：%s" % plan["registered_at"])
    print("登记块指纹：%s" % plan["frozen_sha256"])
    print("plan.json：%s" % p)
    print("")
    if plan["assign"] == "（未填写）":
        print("! 提醒：这次没写分组方式。分组怎么分的（随机？按名次？按座位？）"
              "决定了后面那点差异能不能读 —— 下次登记请填 --assign。")
    kps = qbank_kps()
    if kps and plan["kp"] not in kps:
        near = [k for k in kps if plan["kp"] in k or k in plan["kp"]][:3]
        print("! 提醒：题库里没有「%s」这个知识点。台账按知识点名合并，写错了横向比会散。" % plan["kp"]
              + (" 你是不是想写：%s" % "、".join(near) if near else ""))
    print("登记完成。下一步：做完实验后用 result 子命令把两组得分贴进来。")
    return 0


def cmd_followup(a):
    slug = slug_of(a.name)
    plan = load_plan(slug)
    text = (a.text or "").strip()
    if not text:
        raise AbError("--text 不能为空。")
    guard(text)
    plan.setdefault("followups", []).append({"at": now(), "text": text})
    with open(plan_path(slug), "w", encoding="utf-8") as f:
        json.dump(plan, f, ensure_ascii=False, indent=2)
    print("已追加一条 followup（登记块指纹没有变，说明原登记一个字没动）：%s" % plan["frozen_sha256"][:16])
    print("第 %d 条 followup ｜ %s ｜ %s" % (len(plan["followups"]), plan["followups"][-1]["at"], text))
    return 0


# ══════════════════════════════════════════════════════════════ result
def read_scores(path):
    """读 学生,组,得分,满分 的 CSV。返回 [{学生,组,得分,满分}]。"""
    if not os.path.exists(path):
        raise AbError("找不到成绩文件：%s" % path)
    rows, enc_used = None, None
    for enc in ("utf-8-sig", "utf-8", "gbk"):
        try:
            with open(path, encoding=enc, newline="") as f:
                rows = list(csv.DictReader(f))
            enc_used = enc
            break
        except UnicodeDecodeError:
            continue
    if rows is None:
        raise AbError("成绩文件编码读不出来（试过 utf-8 / gbk）：%s" % path)
    if not rows:
        raise AbError("成绩文件是空的：%s" % path)

    head = {(k or "").strip() for k in rows[0].keys()}
    alias = {}
    for want, cands in (("学生", ("学生", "姓名", "student")), ("组", ("组", "组别", "分组", "group")),
                        ("得分", ("得分", "分数", "score")), ("满分", ("满分", "总分", "full", "full_score"))):
        hit = next((c for c in cands if c in head), None)
        if hit is None:
            raise AbError("成绩文件缺少「%s」列。表头要写成：学生,组,得分,满分（当前表头：%s）"
                          % (want, "、".join(sorted(head))))
        alias[want] = hit

    out, seen = [], {}
    for i, r in enumerate(rows, 2):
        name = str(r.get(alias["学生"]) or "").strip()
        grp = str(r.get(alias["组"]) or "").strip()
        if not name and not grp:
            continue
        if not name or not grp:
            raise AbError("第 %d 行缺学生名或组别：%r" % (i, r))
        try:
            score = float(str(r.get(alias["得分"])).strip())
            full = float(str(r.get(alias["满分"])).strip())
        except (TypeError, ValueError):
            raise AbError("第 %d 行（%s）的得分/满分不是数字：得分=%r 满分=%r"
                          % (i, name, r.get(alias["得分"]), r.get(alias["满分"])))
        if full <= 0:
            raise AbError("第 %d 行（%s）满分是 %s：算不出这次实验的指标，先把数据核对一下。"
                          % (i, name, full))
        if score < 0 or score > full + 1e-9:
            raise AbError("第 %d 行（%s）得分 %s 不在 0～%s 之间，数据有误。" % (i, name, score, full))
        key = (name, grp)
        if key in seen:
            raise AbError("第 %d 行：%s 在「%s」组里出现了两次，会让均值失真，先去掉重复行。"
                          % (i, name, grp))
        seen[key] = True
        out.append({"学生": name, "组": grp, "得分": score, "满分": full})
    if not out:
        raise AbError("成绩文件里一行有效数据都没有：%s" % path)
    return out


def bootstrap_ci(a, b, resamples=BOOT_RESAMPLES, seed=BOOT_SEED, alpha=BOOT_ALPHA):
    """两组均值差的 95% 百分位 bootstrap 区间。

    为什么选 bootstrap 不选 t 区间（理由会原样打进报告）：
      1) 只有标准库，t 分布的分位数要自己写，容易写错；
      2) t 区间假设两组近似正态、且要算合并方差 —— 组内方差为 0 时直接除零，
         而这个工具里「全班都考了同一个分数」是真会发生的（满分就一道题）；
      3) 每组只有 10 来个人，正态假设本身站不住；bootstrap 只要「样本能代表那一组」；
      4) 固定种子 → 同一个输入永远给同一组数字，老师复核时能对得上。
    """
    if not a or not b:
        raise AbError("两组都得有人才能算差异。")
    rnd = random.Random(seed)
    na, nb = len(a), len(b)
    diffs = []
    for _ in range(resamples):
        ma = sum(a[rnd.randrange(na)] for _ in range(na)) / na
        mb = sum(b[rnd.randrange(nb)] for _ in range(nb)) / nb
        diffs.append(ma - mb)
    diffs.sort()
    lo = diffs[int(resamples * (alpha / 2.0))]
    hi = diffs[min(resamples - 1, int(resamples * (1 - alpha / 2.0)))]
    return lo, hi


def cmd_result(a):
    slug = slug_of(a.name)
    plan = load_plan(slug)
    rows = read_scores(a.scores)

    la, lb = plan["groups"][0]["label"], plan["groups"][1]["label"]
    labels = {r["组"] for r in rows}
    unknown = sorted(labels - {la, lb})
    if unknown:
        raise AbError("成绩文件里出现了登记时没有的组：%s。这次实验登记的是「%s」和「%s」，"
                      "组名必须对得上，否则比出来的不是同一个实验。"
                      % ("、".join(unknown), la, lb))
    miss = [x for x in (la, lb) if x not in labels]
    if miss:
        raise AbError("成绩文件里缺这一组：%s（登记时是「%s」vs「%s」）。"
                      % ("、".join(miss), la, lb))

    ga = [r["得分"] for r in rows if r["组"] == la]
    gb = [r["得分"] for r in rows if r["组"] == lb]
    ra = [r["得分"] / r["满分"] for r in rows if r["组"] == la]
    rb = [r["得分"] / r["满分"] for r in rows if r["组"] == lb]
    fulls = sorted({r["满分"] for r in rows})

    mean_a, mean_b = statistics.mean(ga), statistics.mean(gb)
    sd_a = statistics.pstdev(ga) if len(ga) > 1 else 0.0
    sd_b = statistics.pstdev(gb) if len(gb) > 1 else 0.0
    diff = mean_a - mean_b
    lo, hi = bootstrap_ci(ga, gb)
    half = max(abs(hi - diff), abs(diff - lo))
    rate_a, rate_b = statistics.mean(ra), statistics.mean(rb)
    small = (len(ga) < SMALL_SAMPLE) or (len(gb) < SMALL_SAMPLE)
    zero_var = (sd_a < 1e-12 and sd_b < 1e-12)

    who_hi = la if diff > 1e-12 else (lb if diff < -1e-12 else None)
    if who_hi is None:
        sentence = ("本次观察到 %s 组与 %s 组均值相同（%s 分，波动范围 ±%s），样本 %d vs %d，"
                    "不足以下因果结论。" % (la, lb, f2(mean_a), f2(half), len(ga), len(gb)))
        direction = "两组均值相同"
    else:
        sentence = ("本次观察到 %s 组均值比 %s 组%s %s 分（波动范围 ±%s），样本 %d vs %d，"
                    "不足以下因果结论。"
                    % (who_hi, (lb if who_hi == la else la), "高" if diff > 1e-12 else "低",
                       f2(abs(diff)), f2(half), len(ga), len(gb)))
        direction = "%s 高 %s 分" % (who_hi, f2(abs(diff)))

    # ── 报告正文（全程只说「观察到什么」）
    L = []
    if small:
        L.append("!" * 66)
        L.append("!! 样本过小（%d vs %d，每组不足 %d 人）：波动范围会很宽，请勿据此下结论。" % (len(ga), len(gb), SMALL_SAMPLE))
        L.append("!" * 66)
        L.append("")
    L.append("=" * 66)
    L.append("A/B 实验 · 结果对照 ｜ %s" % slug)
    L.append("=" * 66)
    L.append("知识点：%s ｜ 指标：%s" % (plan["kp"], plan["metric"]))
    L.append("登记：%s ｜ 分组方式：%s" % (plan["registered_at"], plan["assign"]))
    L.append("登记时写的预期：%s" % plan["expect"])
    L.append("")
    L.append("【两组】")
    L.append("  组  讲法          人数   均值     标准差   得分率")
    L.append("  A   %-12s %4d   %6s   %6s   %5s%%"
             % (la, len(ga), f2(mean_a), f2(sd_a), f2(rate_a * 100, 1)))
    L.append("  B   %-12s %4d   %6s   %6s   %5s%%"
             % (lb, len(gb), f2(mean_b), f2(sd_b), f2(rate_b * 100, 1)))
    L.append("")
    L.append("【差异】")
    L.append("  A − B = %s 分（%s 组%s）" % (signed(diff), who_hi or "两组", "更高" if who_hi else "相同"))
    L.append("  95%% bootstrap 区间 [%s, %s]，重抽 %d 次，种子 %d"
             % (f2(lo), f2(hi), BOOT_RESAMPLES, BOOT_SEED))
    L.append("  波动范围 ±%s（取区间两端离观测值较远的那一侧）" % f2(half))
    L.append("")
    L.append("【只能这么写的一句话】")
    L.append("  " + sentence)
    L.append("")
    L.append("【这次观察不能说明什么】")
    L.append("  · 分组不是随机抽签：这次是「%s」。两组本来就可能有差别。" % plan["assign"])
    L.append("  · bootstrap 区间只反映「同一批人重抽」的抽样波动，不反映分组偏差、前测差异、"
             "上课时间不同、做题时的状态 —— 这些才是真正会让结论翻车的地方。")
    if lo <= 0 <= hi:
        L.append("  · 区间跨过 0：本次这点差异，跟「两组其实差不多」并不矛盾。")
    else:
        L.append("  · 区间没有跨过 0，但样本只有 %d vs %d、分组也不是随机的 —— 这仍然是「本次观察到」，"
                 "不是「哪个讲法更好」。" % (len(ga), len(gb)))
    if len(fulls) > 1:
        L.append("  · 满分不一致（%s）：原始分均值直接比会失真，以上差异只能配合上面的得分率看。"
                 % "、".join("%g" % x for x in fulls))
    if zero_var:
        L.append("  · 两组组内方差都是 0（每个人得分完全一样）：这种数据通常是满分只有一道题、"
                 "或者分数抄错了。波动范围算出来是 ±0，不代表真的没有波动，"
                 "只说明这组数字里看不出波动。")
    elif sd_a < 1e-12 or sd_b < 1e-12:
        L.append("  · 有一组组内方差是 0（%s 组每个人得分完全一样）：这一组的波动被压成了 0，"
                 "区间会偏窄，看的时候留个心。"
                 % (la if sd_a < 1e-12 else lb))
    d_a, d_b = plan["groups"][0].get("n_declared"), plan["groups"][1].get("n_declared")
    if (d_a and d_a != len(ga)) or (d_b and d_b != len(gb)):
        L.append("  · 人数跟登记时不一致：登记时写的是 %s vs %s，实际拿到 %d vs %d。"
                 % (d_a if d_a else "?", d_b if d_b else "?", len(ga), len(gb)))
    if a.note:
        L.append("")
        L.append("【老师备注】%s" % a.note.strip())
    L.append("")
    L.append("【为什么用 bootstrap 而不是 t 区间】")
    L.append("  1) 只用标准库，不用装 scipy；2) t 区间要算合并方差，组内方差为 0 时会除零，"
             "而「全班同分」在这类小测里真会发生；3) 每组十来人，正态假设站不住；"
             "4) 固定随机种子，同一个输入永远给同一组数字，方便复核。")
    L.append("")
    L.append("【下一步】这次的结果已经写进台账：abtest.py ledger")
    text = "\n".join(L)

    guard(text, a.note or "")
    os.makedirs(os.path.dirname(result_path(slug)), exist_ok=True)
    res = {
        "slug": slug, "kp": plan["kp"], "metric": plan["metric"],
        "computed_at": now(), "group_a": plan["groups"][0], "group_b": plan["groups"][1],
        "n_a": len(ga), "n_b": len(gb), "mean_a": mean_a, "mean_b": mean_b,
        "sd_a": sd_a, "sd_b": sd_b, "rate_a": rate_a, "rate_b": rate_b,
        "diff": diff, "ci": [lo, hi], "half_width": half,
        "higher": who_hi, "direction": direction, "sentence": sentence,
        "method": "bootstrap", "resamples": BOOT_RESAMPLES, "seed": BOOT_SEED, "alpha": BOOT_ALPHA,
        "small_sample": small, "zero_variance": zero_var,
        "full_scores": fulls, "note": (a.note or "").strip(),
        "scores_file": os.path.basename(a.scores),
    }
    with open(result_path(slug), "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=2)
    print(text)
    print("=" * 66)
    print("result.json 已写入：%s" % result_path(slug))
    return 0


# ══════════════════════════════════════════════════════════════ ledger
def cmd_ledger(a):
    plans = all_plans()
    if not plans:
        print("还没有登记任何实验。先跑：abtest.py plan --name … --kp … --group-a … --group-b … --metric …")
        print("（目录：%s）" % ROOT)
        return 0

    L = []
    L.append("=" * 78)
    L.append("A/B 实验台账 · 只做证据整理，不做因果推断")
    L.append("=" * 78)
    L.append("")
    L.append("%-22s %-20s %-16s %-12s %s" % ("知识点", "两种讲法", "差值方向", "样本量", "指标"))
    L.append("-" * 78)

    done, pending, broken = [], [], []
    for p in plans:
        if p.get("_broken"):
            broken.append(p)
            continue
        rp = result_path(p["dir"])
        if not os.path.exists(rp):
            pending.append(p)
            continue
        try:
            with open(rp, encoding="utf-8") as f:
                r = json.load(f)
        except Exception:
            pending.append(p)
            continue
        done.append((p, r))

    def show(p, r):
        la = p["groups"][0]["label"]
        lb = p["groups"][1]["label"]
        pair = "%s vs %s" % (la, lb)
        if r["higher"] is None:
            dirn = "两组均值相同"
        else:
            dirn = "%s 高 %s" % (r["higher"], f2(abs(r["diff"])))
        L.append("%-22s %-20s %-16s %-12s %s"
                 % (p["kp"][:20], pair[:18], dirn[:14], "%d vs %d" % (r["n_a"], r["n_b"]),
                    p["metric"][:16]))

    for p, r in sorted(done, key=lambda x: x[0]["registered_at"]):
        show(p, r)
    for p in sorted(pending, key=lambda x: x["registered_at"]):
        L.append("%-22s %-20s %-16s %-12s %s"
                 % (p["kp"][:20], ("%s vs %s" % (p["groups"][0]["label"], p["groups"][1]["label"]))[:18],
                    "（还没有结果）",
                    "%s vs %s" % (p["groups"][0].get("n_declared") or "?", p["groups"][1].get("n_declared") or "?"),
                    p["metric"][:16]))
    for p in broken:
        L.append("%-22s %s" % (p["slug"][:20], "读不出来：" + p["_broken"][:40]))
    L.append("")

    # ── 同向差异计数
    L.append("【同向差异计数】（同一个讲法，在不同知识点上是不是都往同一个方向走）")
    counts = {}
    for p, r in done:
        for g in (p["groups"][0], p["groups"][1]):
            c = counts.setdefault(g["label"], {"n": 0, "same": 0, "kps": []})
            c["n"] += 1
            other = p["groups"][1]["label"] if g is p["groups"][0] else p["groups"][0]["label"]
            if r["higher"] == g["label"]:
                c["same"] += 1
                c["kps"].append(p["kp"])
            elif r["higher"] == other:
                pass
        L.append("") if False else None
    if not counts:
        L.append("  还没有任何实验结果，先跑 plan + result。")
    for label, c in sorted(counts.items(), key=lambda kv: -kv[1]["same"]):
        if c["same"]:
            L.append("  · %s：在 %d 个不同知识点上都是它更高（%d/%d）：%s"
                     % (label, c["same"], c["same"], c["n"], "、".join(c["kps"])))
        else:
            L.append("  · %s：出现 %d 次，一次都没有更高（0/%d）" % (label, c["n"], c["n"]))
    L.append("")
    L.append("这只说明值得再做一次更严谨的实验 —— 本工具不宣称因果，也不判断哪种讲法更好。")
    if done:
        L.append("方向一致不等于因果关系：每个实验的分组方式、学生底子、上课时间都不同，"
                 "凑出来的「同向」也可能只是巧合。要更严谨，就做下一个实验：")
        L.append("  · 分组用抽签或按名次蛇形，别按老师印象分")
        L.append("  · 实验前先测一次同类型题，把两组的起点记下来")
        L.append("  · 样本每组尽量 ≥ 20 人；不到 10 人时，本工具会在结果顶上打警示")
    if pending:
        L.append("")
        L.append("还有 %d 个实验登记了但没出结果（result 子命令）。" % len(pending))
    text = "\n".join(L)
    guard(text)
    print(text)
    return 0


# ══════════════════════════════════════════════════════════════ CLI
def main(argv=None):
    ap = argparse.ArgumentParser(
        description="单班 A/B 实验工具（只做证据整理，不做因果推断；0 次大模型调用）")
    sub = ap.add_subparsers(dest="cmd")

    p = sub.add_parser("plan", help="事前登记一个实验（登记块冻结，事后不可改）")
    p.add_argument("--name", required=True, help="实验名，同时作为目录名")
    p.add_argument("--kp", required=True, help="知识点（台账按它横向合并）")
    p.add_argument("--group-a", required=True, help="A 组讲法，可带人数：讲法A:12人")
    p.add_argument("--group-b", required=True, help="B 组讲法，可带人数：讲法B:11人")
    p.add_argument("--metric", required=True, help="这次量什么，例如 同类型题正确率")
    p.add_argument("--assign", default=None, help="分组方式：抽签 / 按名次蛇形 / 按座位 …")
    p.add_argument("--expect", default=None, help="预期效果（登记时写下来，事后不许改）")

    r = sub.add_parser("result", help="结果对照：两组均值、差、波动范围")
    r.add_argument("--name", required=True, help="实验名（plan 时用的那个）")
    r.add_argument("--scores", required=True, help="CSV：学生,组,得分,满分")
    r.add_argument("--note", default=None, help="可选备注（同样要过禁语闸门）")

    f = sub.add_parser("followup", help="给已登记的实验追加一条事后记录（不改登记块）")
    f.add_argument("--name", required=True)
    f.add_argument("--text", required=True)

    sub.add_parser("ledger", help="多次实验合并成一张台账 + 同向差异计数")

    a = ap.parse_args(argv)
    if not a.cmd:
        ap.print_help()
        return 0
    try:
        return {"plan": cmd_plan, "result": cmd_result,
                "followup": cmd_followup, "ledger": cmd_ledger}[a.cmd](a)
    except AbError as e:
        print("！%s" % e)
        return 2


if __name__ == "__main__":
    sys.exit(main())
