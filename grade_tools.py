
# -*- coding: utf-8 -*-
"""教师工具箱：数据质检 / 均衡分班 / 推进率 / 小组建议 / 学生学科雷达"""
import io, sys, os, json, statistics as st, argparse
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from collections import defaultdict

HF = PatternFill("solid", fgColor="1F4E79"); HFONT = Font(color="FFFFFF", bold=True, size=11)
WARN = PatternFill("solid", fgColor="FDE2E2"); GOOD = PatternFill("solid", fgColor="E4F2E4")
MID = PatternFill("solid", fgColor="FFF2CC"); TITLE = Font(bold=True, size=13, color="1F4E79")
SUBS = [("语文", 100), ("数学", 150), ("英语", 150), ("物理", 100), ("化学", 100)]

def hs(ws, n):
    for c in range(1, n + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HF; cell.font = HFONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"

ap = argparse.ArgumentParser()
ap.add_argument("file"); ap.add_argument("--classes", type=int, default=8)
ap.add_argument("--group", type=int, default=4, help="小组人数")
ap.add_argument("--min-group", dest="min_group", type=int, default=3,
                help="群组下限：有效人数低于它时不出分班/分组这类群体建议（默认 3）")
a = ap.parse_args()

ws = load_workbook(a.file, data_only=True).active
head = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
idx = {h: i for i, h in enumerate(head)}
for need in ("姓名", "总分"):
    if need not in idx:
        raise SystemExit("表头缺少「%s」列。当前表头：%s" % (need, head))

rows = []
for r in range(2, ws.max_row + 1):
    v = [c.value for c in ws[r]]
    if all(x is None or str(x).strip() == "" for x in v): continue
    nm = str(v[idx["姓名"]]).strip()
    if not nm or nm == "None": continue
    rec = {"row": r, "cls": str(v[idx["班级"]]).strip(), "name": nm}
    for s, F in SUBS:
        try: rec[s] = float(v[idx[s + "分数"]])
        except Exception: rec[s] = None
    try: rec["total"] = float(v[idx["总分"]])
    except Exception: rec["total"] = None
    if rec["total"] is None: continue
    rows.append(rec)

# 总分 0 = 缺考/未录入，算统计前必须先剔除，否则几个缺考就能造出假的"两极分化"
_zeros = [d for d in rows if (d.get("total") or 0) <= 0]
if _zeros:
    print("数据质检：剔除 %d 人总分 0（缺考/未录入）：%s"
          % (len(_zeros), "、".join("%s(%s)" % (d["name"], d["cls"]) for d in _zeros[:10])))
    rows = [d for d in rows if (d.get("total") or 0) > 0]

N = len(rows)
if N == 0:
    raise SystemExit("表里没有有效学生数据")
# 群体口径门槛：1~2 个人也能"分 8 个班""分若干小组"，但那是编出来的结论。
# 只拦「均衡分班建议」「小组搭配建议」两个群体 sheet；数据质检、学生学科雷达是个体口径，照常出。
SMALL = N < a.min_group
SMALL_MSG = ("样本 %d 人，少于群组下限 %d 人，不出分班/分组建议" % (N, a.min_group))
if SMALL:
    print("注意：有效人数 %d < 群组下限 %d：分班/分组建议不出（数据质检、学生学科雷达照常出）"
          % (N, a.min_group))
FULL = sum(F for _, F in SUBS)
ms, ss = {}, {}
for s, F in SUBS:
    vs = [d[s] for d in rows if d[s] is not None]
    if not vs:
        ms[s] = 0.0; ss[s] = 0.0
        print("  警告：%s 一科全空（文字或缺考），该科统计按 0 处理" % s)
    else:
        ms[s] = sum(vs) / len(vs)
        ss[s] = st.pstdev(vs) if len(vs) > 1 else 0.0
t_avg = sum(d["total"] for d in rows) / N
t_sd = st.pstdev([d["total"] for d in rows])
print("学生 %d ｜ 满分 %d ｜ 均分 %.1f ｜ 标准差 %.1f" % (N, FULL, t_avg, t_sd))

out = os.path.splitext(a.file)[0] + "_教师工具箱.xlsx"
wb = Workbook()

# ---------- 1 数据质检 ----------
ws1 = wb.active; ws1.title = "数据质检"
ws1.append(["班级", "姓名", "问题类型", "详情", "建议处理"])
hs(ws1, 5)
issues = 0
for d in rows:
    if d["total"] <= 0:
        ws1.append([d["cls"], d["name"], "总分异常", "总分 = %.0f" % d["total"], "确认是缺考还是零分，缺考应从统计剔除"])
        issues += 1
        continue
    ssum = sum(d[s] for s, _ in SUBS if d[s] is not None)
    if abs(ssum - d["total"]) > 1.0:
        ws1.append([d["cls"], d["name"], "总分与各科不符", "各科合计 %.1f，表里总分 %.1f" % (ssum, d["total"]), "核对录入"])
        issues += 1
    for s, F in SUBS:
        if d[s] is None:
            ws1.append([d["cls"], d["name"], "缺考", "%s 缺" % s, "成绩按缺考处理，不计入均分"])
            issues += 1
            continue
        z = (d[s] - ms[s]) / ss[s] if ss[s] else 0
        if z < -2.5 and d[s] > 0:
            ws1.append([d["cls"], d["name"], "单科异常偏低", "%s %.0f（年级均分 %.0f，z=%.1f）" % (s, d[s], ms[s], z), "确认是否涂错卡或录错"])
            issues += 1
        if d[s] > F:
            ws1.append([d["cls"], d["name"], "超出满分", "%s %.0f > 满分 %.0f" % (s, d[s], F), "核对满分设定"])
            issues += 1
print("数据质检: 发现 %d 处问题" % issues)

# ---------- 2 均衡分班 ----------
ws2 = wb.create_sheet("均衡分班建议")
if SMALL:
    ws2.append(["说明"]); hs(ws2, 1)
    ws2.append([SMALL_MSG])
    ws2.column_dimensions["A"].width = 60
    print()
    print("== 均衡分班 ==")
    print("   %s（按 S 型分 %d 个班要有全年级样本才有意义）" % (SMALL_MSG, a.classes))
else:
    srt = sorted(rows, key=lambda x: -x["total"])
    K = a.classes
    buckets = [[] for _ in range(K)]
    # S 型蛇形分配
    for i, d in enumerate(srt):
        rnd = i // K; pos = i % K
        buckets[pos if rnd % 2 == 0 else K - 1 - pos].append(d)
    ws2.append(["新班", "人数", "均分", "标准差", "最高", "最低", "男生?"])
    hs(ws2, 7)
    print()
    print("== 均衡分班（S 型，%d 个班）==" % K)
    sums = []
    for i, b in enumerate(buckets, 1):
        ts = [d["total"] for d in b]
        if not b:
            ws2.append(["新%d班" % i, 0, "", "", "", "", ""])
            continue
        avg = sum(ts) / len(ts); sd = st.pstdev(ts) if len(ts) > 1 else 0.0
        sums.append(avg)
        ws2.append(["新%d班" % i, len(b), round(avg, 1), round(sd, 1), max(ts), min(ts), ""])
        print("   新%d班  n=%-3d 均分 %.1f  标准差 %.1f  极差 %.0f" % (i, len(b), avg, sd, max(ts) - min(ts)))
    print("   班均分极差: %.1f 分（越小越均衡）" % (max(sums) - min(sums)))
    for i, w in enumerate([10, 7, 8, 9, 8, 8, 8], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

# ---------- 3 班级推进率 ----------
# 按班的均分/及格率/低分率同样是群体口径：n < 群组下限时一并拦掉（和均衡分班同一类误导）
PL, GL, LOW = FULL * .6, FULL * .85, FULL * .4
byc = defaultdict(list)
for d in rows: byc[d["cls"]].append(d)
ws3 = wb.create_sheet("班级推进率")
if SMALL:
    ws3.append(["说明"]); hs(ws3, 1)
    ws3.append([SMALL_MSG])
    ws3.column_dimensions["A"].width = 60
    print()
    print("== 班级推进率 ==")
    print("   %s（按班的均分/及格率要先有足够样本才算得准）" % SMALL_MSG)
else:
    ws3.append(["班级", "人数", "均分", "及格率", "优秀率", "低分率(<40%)", "临界人数(±20分)", "达标情况"])
    hs(ws3, 8)
    print()
    print("== 班级推进率 ==")
    for c in sorted(byc, key=lambda x: int("".join(ch for ch in x if ch.isdigit()) or 99)):
        g = byc[c]; ts = [x["total"] for x in g]
        pa = 100 * sum(1 for t in ts if t >= PL) / len(ts)
        gd = 100 * sum(1 for t in ts if t >= GL) / len(ts)
        low = 100 * sum(1 for t in ts if t < LOW) / len(ts)
        near = sum(1 for t in ts if abs(t - PL) <= 20)
        tag = "良好" if pa >= 85 else ("需关注" if pa < 75 else "正常")
        ws3.append([c, len(g), round(sum(ts) / len(ts), 1), "%.0f%%" % pa, "%.0f%%" % gd,
                    "%.0f%%" % low, near, tag])
        if low > 15: ws3.cell(row=ws3.max_row, column=6).fill = WARN
        print("   %-9s n=%-3d 均分 %.1f 及格 %.0f%% 低分率 %.0f%% 临界 %d 人  %s" %
              (c, len(g), sum(ts) / len(ts), pa, low, near, tag))
    for i, w in enumerate([11, 7, 8, 9, 9, 12, 13, 10], 1):
        ws3.column_dimensions[get_column_letter(i)].width = w

# ---------- 4 小组建议 ----------
ws4 = wb.create_sheet("小组搭配建议")
if SMALL:
    ws4.append(["说明"]); hs(ws4, 1)
    ws4.append([SMALL_MSG])
    ws4.column_dimensions["A"].width = 60
    print()
    print("== 小组搭配建议 ==")
    print("   %s（强弱混编分组要先有足够的强弱样本）" % SMALL_MSG)
else:
    ws4.append(["班级", "组号", "成员", "组内均分", "强弱差"])
    hs(ws4, 5)
    print()
    print("== 小组搭配建议（每组 %d 人，强弱混编）==" % a.group)
    for c in sorted(byc, key=lambda x: int("".join(ch for ch in x if ch.isdigit()) or 99)):
        g = sorted(byc[c], key=lambda x: -x["total"])
        groups = [[] for _ in range((len(g) + a.group - 1) // a.group)]
        for i, d in enumerate(g):
            rnd = i // len(groups); pos = i % len(groups)
            groups[pos if rnd % 2 == 0 else len(groups) - 1 - pos].append(d)
        for gi, mem in enumerate(groups, 1):
            if not mem: continue
            ts = [x["total"] for x in mem]
            ws4.append([c, "%s-第%d组" % (c[-2:], gi), "、".join(x["name"] for x in mem),
                        round(sum(ts) / len(ts), 1), round(max(ts) - min(ts))])
        print("   %-9s 分 %d 组，每组 %d–%d 人" % (c, len(groups), min(len(m) for m in groups), max(len(m) for m in groups)))
    for i, w in enumerate([11, 12, 40, 9, 8], 1):
        ws4.column_dimensions[get_column_letter(i)].width = w

# ---------- 5 学生学科雷达数据 ----------
ws5 = wb.create_sheet("学生学科雷达")
ws5.append(["班级", "姓名", "总分"] + [s for s, _ in SUBS] + [s + "(z)" for s, _ in SUBS] + ["最弱科", "不均衡度"])
hs(ws5, 4 + len(SUBS) * 2 + 2)
for d in sorted(rows, key=lambda x: -x["total"]):
    zs = {s: (d[s] - ms[s]) / ss[s] if (d[s] is not None and ss[s]) else 0 for s, _ in SUBS}
    weak = min(zs, key=lambda k: zs[k])
    spread = max(zs.values()) - min(zs.values())
    ws5.append([d["cls"], d["name"], d["total"]] + [d[s] for s, _ in SUBS] +
               [round(zs[s], 2) for s, _ in SUBS] + [weak, round(spread, 2)])
    if spread >= 2.5: ws5.cell(row=ws5.max_row, column=4 + len(SUBS) * 2 + 2).fill = WARN
for i, w in enumerate([11, 10, 8] + [7] * len(SUBS) + [9] * len(SUBS) + [9, 9], 1):
    ws5.column_dimensions[get_column_letter(i)].width = w
ws5.auto_filter.ref = "A1:%s%d" % (get_column_letter(4 + len(SUBS) * 2 + 2), ws5.max_row)

wb.save(out)
print()
print("已生成:", out)
print("工作表:", wb.sheetnames)
