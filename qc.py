# -*- coding: utf-8 -*-
"""数据质检：读表之前先看看这张表有没有毛病。

为什么单独一个文件：
    成绩单里最常见的错，不是算错，是数据本身错——满分 150 的科目出现 200 分、
    总分跟各科加起来对不上、某一列整列没录。这些错会一路传下去，
    最后算出一个老师一眼就知道不对的数。那一次之后，他就不会再用第二次了。

三条口径（和别处一致）：
    ① 只报告，不改数。到底是「满分写错了」还是「分数录错了」，只有老师知道。
    ② 不猜满分。满分要么由 --full 给出，要么从列名里读（如「数学(150)」）；
       拿不到就说「无法判定是否超满分」，绝不拿这一列的最大值当满分——那是循环论证，
       真有一个 200 的错值，它就成了满分。
    ③ 缺考是正常情况，不算错。"" 缺考 缺 - — / 一律跳过，不报。
"""
import argparse
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

import openpyxl
import ingest

KNOWN_ABSENT = ("", "缺考", "缺", "-", "—", "/", "／", "无", "未考", "缓考")
NOT_SCORE_COL = ("学号", "考号", "座号", "序号", "编号", "班级", "班次", "姓名", "名字",
                 "性别", "排名", "名次", "等级", "等第", "备注")
TOTAL_COL = ("总分", "合计", "总成绩", "总得分")
TOL = 0.01

SEV_ERROR = "必须处理"
SEV_WARN = "建议核对"


def is_absent(v):
    if v is None:
        return True
    return str(v).strip() in KNOWN_ABSENT


def as_num(v):
    """能当数字就返回 float，否则 None。"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "")
    if s in KNOWN_ABSENT:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def parse_full(spec):
    """"语文=120,数学=150" -> {"语文":120.0, "数学":150.0}"""
    out = {}
    if not spec:
        return out
    for part in re.split(r"[,，;；]", spec):
        part = part.strip()
        if not part:
            continue
        m = re.match(r"^(.+?)\s*[=:：]\s*([0-9.]+)$", part)
        if not m:
            raise SystemExit("--full 写法不对：%s\n应该像这样：语文=120,数学=150" % part)
        out[m.group(1).strip()] = float(m.group(2))
    return out


def full_from_header(name):
    """从列名里读满分：数学(150) / 数学（满分150） / 数学 150分"""
    if not name:
        return None
    for pat in (r"[（(]\s*满分\s*([0-9.]+)\s*[)）]",
                r"[（(]\s*([0-9.]+)\s*分?\s*[)）]",
                r"满分\s*([0-9.]+)",
                r"\s([0-9.]+)\s*分$"):
        m = re.search(pat, str(name))
        if m:
            try:
                return float(m.group(1))
            except ValueError:
                return None
    return None


def score_columns(records):
    """哪些列像分数列：多数格子能当数字或是有名的缺考标记，且列名不像学号/姓名/排名。"""
    cols = {}
    for rec in records:
        for k, v in (rec.get("raw") or {}).items():
            cols.setdefault(k, [])
            cols[k].append(v)
    picked = []
    for name, vals in cols.items():
        low = str(name).strip()
        if any(t in low for t in NOT_SCORE_COL):
            continue
        if any(t in low for t in TOTAL_COL):
            continue
        nums = sum(1 for v in vals if as_num(v) is not None)
        cells = sum(1 for v in vals if v is not None and str(v).strip() != "")
        if cells == 0:
            continue
        if nums >= max(1, int(cells * 0.6)):
            picked.append(name)
    return picked


def check(path, full_map=None):
    """返回 (问题列表, 报告)。问题每项：{sev, sheet, who, col, msg}"""
    full_map = dict(full_map or {})
    records, rep = ingest.read_all_sheets(path)
    issues = []

    def add(sev, sheet, who, col, msg):
        issues.append({"sev": sev, "sheet": sheet, "who": who, "col": col, "msg": msg})

    # 按 sheet 分组
    by_sheet = {}
    for rec in records:
        by_sheet.setdefault(rec.get("__sheet") or "(无表名)", []).append(rec)

    for sheet, recs in by_sheet.items():
        scols = score_columns(recs)

        # 满分来源：--full 优先，其次列名
        # 满分怎么定：①列名里写着（数学(150)）②--full 里给的（按最长前缀匹配，
        # 因为真实表里列名常是「语文分数」「数学成绩」，不会正好等于「语文」）
        # ③两样都没有 → 不填，下面如实说「不知道满分」。
        fulls = {}
        fulls_note = []
        for c in scols:
            f = full_from_header(c)
            if f:
                fulls[c] = f
                fulls_note.append((c, f, "列名里写着满分"))
                continue
            keys = [k for k in full_map if k and k in str(c)]
            if not keys:
                continue          # 两样都没有：不填，也不编一个满分出来
            k = max(keys, key=len)
            fulls[c] = full_map[k]
            fulls_note.append((c, fulls[c], "来自 --full「%s」" % k))

        for rec in recs:
            who = str(rec.get("name") or "").strip()
            sid = str(rec.get("sid") or "").strip()
            label = who or (("学号 " + sid) if sid else "(没名字的一行)")

            if not who:
                add(SEV_ERROR, sheet, label, "姓名", "这一行没有姓名，不知道是谁")

            for c in scols:
                v = (rec.get("raw") or {}).get(c)
                if is_absent(v):
                    continue
                num = as_num(v)
                if num is None:
                    add(SEV_ERROR, sheet, label, c, "「%s」不是数字（原始值：%s）" % (c, v))
                    continue
                if num < 0:
                    add(SEV_ERROR, sheet, label, c, "%s 是负分：%g" % (c, num))
                if c in fulls and num > fulls[c] + 1e-9:
                    add(SEV_ERROR, sheet, label, c,
                        "%s 得分 %g 超过满分 %g" % (c, num, fulls[c]))

            # 总分核对：只有在总分是数字、且各科都不缺考时才比
            tot_raw = None
            for k, v in (rec.get("raw") or {}).items():
                if any(t in str(k) for t in TOTAL_COL):
                    tot_raw = v
                    break
            if tot_raw is not None and scols:
                tot = as_num(tot_raw)
                if tot is not None:
                    parts = [(c, (rec.get("raw") or {}).get(c)) for c in scols]
                    # 只要有科目缺考、或某一科不是数字，总分就没法核——
                    # 这时候硬加起来会报出一条假警报（把「一百」当 0 之类），比不报更糟。
                    if all(not is_absent(v) for _c, v in parts) and \
                       all(as_num(v) is not None for _c, v in parts):
                        s = sum(as_num(v) for _c, v in parts)
                        if abs(s - tot) > TOL:
                            add(SEV_WARN, sheet, label, "总分",
                                "总分写 %g，但各科加起来是 %g，差 %g" % (tot, s, tot - s))

        # 必需列：后面的引擎要吃「班级」和「姓名」，缺了会在别处崩成 KeyError
        heads = set()
        for rec in recs:
            heads |= set((rec.get("raw") or {}).keys())
        for need in ("班级", "姓名"):
            if heads and not any(need in str(h) for h in heads):
                add(SEV_ERROR, sheet, "-", need,
                    "这张表没有「%s」列 —— 后面的分析要按它归类，缺了会直接报错" % need)

        # 重名
        seen = {}
        for rec in recs:
            n = str(rec.get("name") or "").strip()
            if n:
                seen[n] = seen.get(n, 0) + 1
        for n, k in seen.items():
            if k > 1:
                add(SEV_WARN, sheet, n, "姓名", "同名 %d 行，可能是重复录入，也可能是两个人" % k)

        # 整列可疑
        for c in scols:
            vals = [as_num((r.get("raw") or {}).get(c)) for r in recs]
            nums = [v for v in vals if v is not None]
            if len(nums) >= 5:
                if all(abs(v) < 1e-9 for v in nums):
                    add(SEV_WARN, sheet, "-", c, "%s 这一列全是 0，可能整列没录" % c)
                elif len(set(round(v, 4) for v in nums)) == 1:
                    add(SEV_WARN, sheet, "-", c, "%s 这一列全都是 %g，可能复制错了" % (c, nums[0]))

        # 有满分来源却查不到某列：明说，不猜
        unknown = [c for c in scols if c not in fulls]
        if unknown:
            add(SEV_WARN, sheet, "-", "、".join(unknown[:6]),
                "这几列不知道满分，无法判断是否超分（用 --full 语文=120,数学=150 指定即可）")

    return issues, {"report": rep, "fulls": fulls_note, "sheets": list(by_sheet)}


def write_xlsx(issues, path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "数据质检"
    ws.append(["级别", "工作表", "谁", "列", "问题"])
    for it in issues:
        ws.append([it["sev"], it["sheet"], it["who"], it["col"], it["msg"]])
    for i, w in enumerate([12, 14, 16, 18, 70], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    wb.save(path)


def main(argv=None):
    ap = argparse.ArgumentParser(description="数据质检：先看表有没有毛病，再谈分析")
    ap.add_argument("src", help="成绩单 .xlsx")
    ap.add_argument("--full", default=None, help='各科满分，例如 "语文=120,数学=150"')
    ap.add_argument("--out-dir", default=None, help="清单输出目录（默认 教师工具/输出）")
    ap.add_argument("--soft", action="store_true",
                    help="发现问题也返回 0（给流水线用：质检不该拦住后面的分析）")
    a = ap.parse_args(argv)

    if not os.path.exists(a.src):
        print("找不到这个文件：%s" % a.src)
        return 2

    issues, meta = check(a.src, parse_full(a.full))
    fulls_note = meta.get("fulls") or []
    errs = [i for i in issues if i["sev"] == SEV_ERROR]
    warns = [i for i in issues if i["sev"] == SEV_WARN]

    print("=" * 78)
    print("数据质检：%s" % os.path.basename(a.src))
    print("=" * 78)
    print("读入 %d 个工作表：%s" % (len(meta["sheets"]), "、".join(meta["sheets"]) or "(无)"))
    sk = meta["report"]["skipped"]
    if sk:
        for name, why in sk:
            print("  跳过「%s」：%s" % (name, why))
    print("问题合计 %d 条（必须处理 %d ／ 建议核对 %d）" % (len(issues), len(errs), len(warns)))
    if fulls_note:
        print()
        print("满分取值（判断超分的依据）：")
        for col, val, src in fulls_note:
            print("  %-16s = %-6g  %s" % (col, val, src))
    print()

    def show(title, items):
        if not items:
            return
        print("-" * 78)
        print(title)
        print("-" * 78)
        for it in items:
            print("  [%s] %s ｜ %s ｜ %s" % (it["sheet"], it["who"], it["col"], it["msg"]))

    show("必须处理（这些会让结果直接算错）", errs)
    show("建议核对", warns)
    if not issues:
        print("没发现明显问题。")

    out_dir = a.out_dir or os.path.join(HERE, "输出")
    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(a.src))[0]
    out = os.path.join(out_dir, "数据质检_%s.xlsx" % stem)
    write_xlsx(issues, out)
    print()
    print("清单已写出：%s" % out)
    print("说明：这里只报告，不改任何数。到底是满分写错还是分数录错，得你看着改。")
    if a.soft:
        return 0
    return 1 if errs else 0


if __name__ == "__main__":
    sys.exit(main())
