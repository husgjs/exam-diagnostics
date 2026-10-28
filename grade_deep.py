
# -*- coding: utf-8 -*-
"""九年级学情分析：班级对比 / 学科诊断 / 临界生提分方案 / 偏科诊断"""
import io, sys, os, json, statistics as st, argparse
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from collections import defaultdict

HEAD_FILL = PatternFill("solid", fgColor="1F4E79")
HEAD_FONT = Font(color="FFFFFF", bold=True, size=11)
WARN = PatternFill("solid", fgColor="FDE2E2")
GOOD = PatternFill("solid", fgColor="E4F2E4")
MID = PatternFill("solid", fgColor="FFF2CC")
TITLE = Font(bold=True, size=13, color="1F4E79")

SUBS = [("语文", 100), ("数学", 150), ("英语", 150), ("物理", 100), ("化学", 100)]

def _avg(vs, nd=1):
    """空列表返回空串，不能返回 0 冒充真实均分"""
    return round(sum(vs) / len(vs), nd) if vs else ""

def hs(ws, n):
    for c in range(1, n + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEAD_FILL; cell.font = HEAD_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--line", type=float, default=500, help="重点线，默认 500")
    ap.add_argument("--band", type=float, default=40, help="临界带宽度，默认线下40分")
    a = ap.parse_args()

    ws = load_workbook(a.file, data_only=True).active
    head = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    idx = {h: i for i, h in enumerate(head)}

    rows = []
    for r in range(2, ws.max_row + 1):
        v = [c.value for c in ws[r]]
        if all(x is None or str(x).strip() == "" for x in v): continue
        nm = str(v[idx["姓名"]]).strip() if idx.get("姓名") is not None else ""
        if not nm: continue
        rec = {"cls": str(v[idx["班级"]]).strip(), "name": nm}
        ok = True
        for s, F in SUBS:
            col = idx.get(s + "分数")
            try: rec[s] = float(v[col])
            except Exception: rec[s] = None; ok = False
        try: rec["total"] = float(v[idx["总分"]])
        except Exception: rec["total"] = None
        if rec["total"] is None: continue
        rows.append(rec)

    # 总分 0 = 缺考/未录入，**必须先剔除再算任何统计**。
    # 否则一个班 6 个 0 分会把均分拉低、把标准差顶到年级的 1.5 倍，
    # 于是「两极分化」整个结论都是假的——2 班就是活例子（含 0：1.54 倍；剔掉后：0.96 倍）。
    zeros = [d for d in rows if (d.get("total") or 0) <= 0]
    if zeros:
        print("数据质检：剔除 %d 人总分 0（缺考/未录入）：%s"
              % (len(zeros), "、".join("%s(%s)" % (d["name"], d["cls"]) for d in zeros[:10])))
        rows = [d for d in rows if (d.get("total") or 0) > 0]

    N = len(rows)
    if N == 0:
        raise SystemExit("表里没有有效学生数据：请检查「姓名」和「总分」列")
    if N < 2:
        raise SystemExit("只有 1 名学生，做不了分布分析（至少需要 2 人）")
    FULL = sum(F for _, F in SUBS)
    print("学生 %d 人 ｜ 满分 %.0f ｜ 重点线 %.0f ｜ 临界带 %.0f–%.0f" % (N, FULL, a.line, a.line - a.band, a.line + a.band))

    # 均分/标准差
    ms, ss = {}, {}
    for s, F in SUBS:
        vs = [d[s] for d in rows if d[s] is not None]
        if not vs:
            ms[s] = 0.0
            ss[s] = 0.0
            print("  警告：%s 一科全空，该科统计跳过" % s)
        else:
            ms[s] = sum(vs) / len(vs)
            ss[s] = st.pstdev(vs) if len(vs) > 1 else 0.0
    totals = [d["total"] for d in rows]
    t_avg, t_sd = sum(totals) / N, st.pstdev(totals)

    # 班级
    by_cls = defaultdict(list)
    for d in rows: by_cls[d["cls"]].append(d)
    classes = sorted(by_cls, key=lambda x: int("".join(ch for ch in x if ch.isdigit()) or 99))

    # 临界生：lead - band <= total < lead + 10
    # 高分组的各科水平（前 30%），用来衡量「该生这科还能提多少」
    top30 = sorted(rows, key=lambda x: -x["total"])[:max(1, int(N * 0.3))]
    high_avg = {}
    for s, F in SUBS:
        vs = [d[s] for d in top30 if d[s] is not None]
        high_avg[s] = sum(vs) / len(vs) if vs else ms[s]

    crit = [d for d in rows if a.line - a.band <= d["total"] < a.line + 10]
    crit.sort(key=lambda x: -(x["total"]))
    for d in crit:
        zs = {s: (d[s] - ms[s]) / ss[s] for s, F in SUBS if d[s] is not None and ss[s] > 0}
        main_s = min(zs, key=lambda k: zs[k]) if zs else None      # 自己最拖后腿的一科
        room = round(max(0.0, high_avg[main_s] - d[main_s]), 1) if main_s else 0
        d["need"] = round(a.line - d["total"], 1)
        d["main"] = main_s
        d["pot"] = round(min(room, 30), 1)
        d["gap_vs_avg"] = round(ms[main_s] - d[main_s], 1) if main_s else 0
        d["group"] = "保线" if d["total"] >= a.line else "冲线"
        if d["pot"] <= 0:
            d["advice"] = "各科都已在高分组水平，主要靠稳定性，别动科目结构"
        else:
            d["advice"] = "主攻%s：距高分组均分还差 %.1f 分，可提 ≈%.0f 分" % (main_s, high_avg[main_s] - d[main_s], d["pot"])

    # 偏科：z 分数
    for d in rows:
        zs = {s: (d[s] - ms[s]) / ss[s] for s, F in SUBS if d[s] is not None and ss[s] > 0}
        if not zs: continue
        hi = max(zs, key=lambda k: zs[k]); lo = min(zs, key=lambda k: zs[k])
        d["z"] = zs
        d["best"], d["worst"] = hi, lo
        d["spread"] = round(zs[hi] - zs[lo], 2)
        d["lone"] = d["spread"] >= 1.5            # 一根独苗
        d["weak"] = [s for s in zs if zs[s] <= -1.0]
    lone = sorted([d for d in rows if d.get("lone")], key=lambda x: -x["spread"])

    print("临界生 %d 人 ｜ 明显偏科 %d 人" % (len(crit), len(lone)))

    out = os.path.splitext(a.file)[0] + "_学情分析.xlsx"
    wb = Workbook()

    # 1 年级概览
    ws1 = wb.active; ws1.title = "年级概览"
    ws1["A1"] = "九年级学情概览"; ws1["A1"].font = TITLE
    items = [("学生人数", N), ("满分", FULL), ("年级均分", round(t_avg, 1)),
             ("标准差", round(t_sd, 1)), ("最高", max(totals)), ("最低", min(totals)),
             ("重点线（假设）", a.line), ("临界生人数", len(crit)),
             ("明显偏科人数", len(lone)), ("班级数", len(classes))]
    r = 3
    for k, v in items:
        ws1.cell(row=r, column=1, value=k).font = Font(bold=True)
        ws1.cell(row=r, column=2, value=v); r += 1
    ws1.column_dimensions["A"].width = 18; ws1.column_dimensions["B"].width = 16

    # 2 班级对比
    ws2 = wb.create_sheet("班级对比")
    ws2.append(["班级", "人数", "均分", "标准差", "极差", "及格率(60%)", "优秀率(85%)"] + ["%s均分" % s for s, _ in SUBS] + ["均衡度"])
    hs(ws2, 8 + len(SUBS))
    for c in classes:
        g = by_cls[c]; ts = [x["total"] for x in g]
        sd = st.pstdev(ts)
        lvl = "两极分化" if sd > 1.4 * t_sd else ("较均衡" if sd < 0.85 * t_sd else "正常")
        ws2.append([c, len(g), round(sum(ts) / len(ts), 1), round(sd, 1), round(max(ts) - min(ts)),
                    "%.0f%%" % (100 * sum(1 for t in ts if t >= FULL * .6) / len(ts)),
                    "%.0f%%" % (100 * sum(1 for t in ts if t >= FULL * .85) / len(ts))] +
                   [_avg([x[s] for x in g if x[s] is not None]) for s, _ in SUBS] +
                   [lvl])
        if lvl == "两极分化": ws2.cell(row=ws2.max_row, column=8 + len(SUBS)).fill = WARN
    for i, w in enumerate([11, 7, 8, 9, 8, 12, 12] + [8] * len(SUBS) + [11], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    # 3 各科诊断
    ws3 = wb.create_sheet("各科诊断")
    ws3.append(["科目", "满分", "均分", "得分率", "标准差", "归一化波动", "及格率", "对总分差距的贡献"])
    hs(ws3, 8)
    norm = sorted([(s, ss[s] / F) for s, F in SUBS], key=lambda x: -x[1])
    rank_of = {s: i + 1 for i, (s, _) in enumerate(norm)}
    for s, F in SUBS:
        vs = [d[s] for d in rows if d[s] is not None]
        pass_rate = ("%.0f%%" % (100 * sum(1 for v in vs if v >= F * .6) / len(vs))) if vs else "—"
        ws3.append([s, F, round(ms[s], 1), "%.0f%%" % (100 * ms[s] / F), round(ss[s], 1),
                    "%.1f%%" % (100 * ss[s] / F),
                    pass_rate,
                    "第%d位" % rank_of[s]])
        if norm[0][0] == s: ws3.cell(row=ws3.max_row, column=6).fill = WARN
    for i, w in enumerate([8, 8, 8, 9, 9, 12, 9, 14], 1):
        ws3.column_dimensions[get_column_letter(i)].width = w

    # 4 临界生提分方案
    ws4 = wb.create_sheet("临界生提分方案")
    ws4.append(["组别", "距重点线", "班级", "姓名", "总分"] + [s for s, _ in SUBS] + ["主攻科目", "可提分", "建议"])
    hs(ws4, 9 + len(SUBS))
    for d in crit:
        ws4.append([d["group"], d["need"], d["cls"], d["name"], d["total"]] +
                   [d[s] for s, _ in SUBS] + [d["main"] or "—", d["pot"], d["advice"]])
        ws4.cell(row=ws4.max_row, column=2).fill = MID if abs(d["need"]) <= 20 else WARN
    for i, w in enumerate([9, 11, 10, 8] + [7] * len(SUBS) + [10, 12, 8, 40], 1):
        ws4.column_dimensions[get_column_letter(i)].width = w
    ws4.auto_filter.ref = "A1:%s%d" % (get_column_letter(9 + len(SUBS)), ws4.max_row)

    # 5 偏科诊断
    ws5 = wb.create_sheet("偏科诊断")
    ws5.append(["班级", "姓名", "总分", "最强科", "最弱科", "落差(z)", "诊断"] + [s + "(z)" for s, _ in SUBS])
    hs(ws5, 7 + len(SUBS))
    for d in lone:
        diag = "一根独苗：只有%s能打" % d["best"] if d["z"][d["best"]] > 0.8 else "严重偏科：%s拖后腿" % d["worst"]
        ws5.append([d["cls"], d["name"], d["total"], d["best"], d["worst"], d["spread"], diag] +
                   [round(d["z"].get(s, 0), 2) for s, _ in SUBS])
        ws5.cell(row=ws5.max_row, column=6).fill = WARN
    for i, w in enumerate([11, 10, 8, 9, 9, 9, 30] + [9] * len(SUBS), 1):
        ws5.column_dimensions[get_column_letter(i)].width = w
    ws5.auto_filter.ref = "A1:%s%d" % (get_column_letter(7 + len(SUBS)), ws5.max_row)

    # 6 学生明细
    ws6 = wb.create_sheet("学生明细")
    ws6.append(["总分", "班级", "姓名"] + [s for s, _ in SUBS] + ["年级名次"])
    hs(ws6, 5 + len(SUBS))
    for i, d in enumerate(sorted(rows, key=lambda x: -x["total"]), 1):
        ws6.append([d["total"], d["cls"], d["name"]] + [d[s] for s, _ in SUBS] + [i])
        if d["total"] < FULL * .6: ws6.cell(row=ws6.max_row, column=1).fill = WARN
        elif d["total"] >= FULL * .85: ws6.cell(row=ws6.max_row, column=1).fill = GOOD
    # 缺考的不进统计，但必须在表里打上「缺」字，不能凭空消失
    for d in zeros:
        ws6.append(["缺考", d["cls"], d["name"]] + ["缺"] * len(SUBS) + ["缺考"])
        for _c in range(1, 6 + len(SUBS)):
            ws6.cell(row=ws6.max_row, column=_c).fill = MID
    for i, w in enumerate([8, 11, 10] + [7] * len(SUBS) + [9], 1):
        ws6.column_dimensions[get_column_letter(i)].width = w
    ws6.auto_filter.ref = "A1:%s%d" % (get_column_letter(5 + len(SUBS)), ws6.max_row)

    wb.save(out)

    # 存中间结果给报告用
    js = {"N": N, "FULL": FULL, "t_avg": t_avg, "t_sd": t_sd, "line": a.line, "band": a.band,
          "subjects": [{"s": s, "F": F, "avg": ms[s], "sd": ss[s], "rate": ms[s] / F * 100,
                        "norm": ss[s] / F * 100,
                        "pass": (100 * sum(1 for d in rows if d[s] is not None and d[s] >= F * .6) /
                                 len([d for d in rows if d[s] is not None]))
                                if [d for d in rows if d[s] is not None] else 0.0}
                       for s, F in SUBS],
          "classes": [{"c": c, "n": len(by_cls[c]), "avg": sum(x["total"] for x in by_cls[c]) / len(by_cls[c]),
                       "sd": st.pstdev([x["total"] for x in by_cls[c]]),
                       "range": max(x["total"] for x in by_cls[c]) - min(x["total"] for x in by_cls[c]),
                       "max": max(x["total"] for x in by_cls[c]),
                       "min": min(x["total"] for x in by_cls[c]),
                       "pass": 100 * sum(1 for x in by_cls[c] if x["total"] >= FULL * .6) / len(by_cls[c]),
                       "good": 100 * sum(1 for x in by_cls[c] if x["total"] >= FULL * .85) / len(by_cls[c])}
                      for c in classes],
          "crit": [{"name": d["name"], "cls": d["cls"], "total": d["total"], "need": d["need"],
                    "main": d["main"], "pot": d["pot"], "advice": d["advice"]} for d in crit],
          "lone_n": len(lone),
          "lone_top": [{"name": d["name"], "cls": d["cls"], "total": d["total"],
                        "best": d["best"], "worst": d["worst"], "spread": d["spread"]} for d in lone[:10]]}
    json.dump(js, open(r"D:/新建文件夹 (2)/_tmp/grade_stats.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    print()
    print("== 最该补的科目（按波动）==")
    for s, v in norm: print("   %-4s 归一化波动 %.1f%%" % (s, v * 100))
    print()
    print("== 临界生（前 8 名，距线 %s）==" % ("%.0f" % a.line))
    for d in crit[:8]:
        print("   %-4s %-10s 总分 %-6.1f 差 %-5.1f 主攻 %-4s 可提 %.0f 分" % (
            d["cls"][-2:], d["name"], d["total"], d["need"], d["main"] or "—", d["pot"]))
    print()
    print("== 明显偏科（前 6）==")
    for d in lone[:6]:
        print("   %-4s %-10s 强项 %-4s 弱项 %-4s 落差 %.2f" % (d["cls"][-2:], d["name"], d["best"], d["worst"], d["spread"]))
    print()
    print("已生成:", out)
    return out

if __name__ == "__main__":
    main()
