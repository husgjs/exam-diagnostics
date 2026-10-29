# -*- coding: utf-8 -*-
"""辅导机构常见表格的读取层。

辅导机构的表有三个特点，直接读会读不动：
  1. 一个班一个 sheet —— 只读第一个就漏掉其余班
  2. 表头不在第一行 —— 前面常有标题行、统计行
  3. 学号/姓名纵向合并 —— 只有第一行有值

另外，答题卡导出的通常是「一人一行 × 每题一列」的宽表，
而讲评清单要的是「一人一题一行」的长表，需要转换。

口径（与项目其余部分一致）：
  · 认不出表头就明说，不猜
  · 宽表转长表必须有每题满分；没有就报错，不默认满分
"""
import os
import re
import unicodedata

try:
    import openpyxl
except ImportError:          # 允许在没有 openpyxl 的环境里 import 本模块做单测
    openpyxl = None


# 列名同义词：同一个含义的不同写法。比较时先归一化（去空格、全角转半角、小写）。
ALIASES = {
    "name":  ["姓名", "学生姓名", "学员姓名", "学生", "名字", "学生名"],
    "sid":   ["学号", "学籍号", "考号", "编号", "学籍"],
    "cls":   ["班级", "班", "所在班级", "班型", "班别", "班级名称"],
    "total": ["总分", "合计", "总成绩", "合计分", "总计"],
}


def norm(v):
    """列名归一化：None -> ''，全角空格/NBSP/换行都去掉，转小写。"""
    if v is None:
        return ""
    s = unicodedata.normalize("NFKC", str(v))
    s = s.replace("\u3000", " ").replace("\xa0", " ")
    s = re.sub(r"\s+", "", s)
    return s.lower()


def _norm_alias_map():
    m = {}
    for key, names in ALIASES.items():
        for n in names:
            m[norm(n)] = key
    return m


_ALIAS = _norm_alias_map()


def canon(header):
    """把一行表头映射成 {标准名: 列下标}。认不出的列不进结果。"""
    out = {}
    for i, h in enumerate(header):
        k = _ALIAS.get(norm(h))
        if k and k not in out:          # 同名列取第一个
            out[k] = i
    return out


def find_header(rows, max_scan=25):
    """在行里找表头：第一个同时认得「姓名」或「学号」的行。

    返回 (行号, 归一化后的表头列表)；找不到返回 (None, None)。
    不猜：宁可返回 None，也不把第 0 行当表头硬算。
    """
    for i, row in enumerate(rows[:max_scan]):
        if row is None:
            continue
        vals = list(row)
        if canon(vals).get("name") is not None or canon(vals).get("sid") is not None:
            return i, vals
    return None, None


def fill_merged(rows, ranges):
    """在一份「行 -> 列表」的矩阵上填满合并区域，返回填了多少格。

    不能直接改 worksheet 里的 MergedCell —— openpyxl 把它的 value 设为只读，
    硬改会抛 AttributeError。所以一律在内存矩阵上处理。
    """
    filled = 0
    for rng in ranges:
        try:
            r0, c0 = rng.min_row - 1, rng.min_col - 1
            r1, c1 = rng.max_row - 1, rng.max_col - 1
        except Exception:
            continue
        if r0 >= len(rows) or c0 >= len(rows[r0]):
            continue
        v = rows[r0][c0]
        if v is None or v == "":
            continue
        for r in range(r0, min(r1 + 1, len(rows))):
            for c in range(c0, min(c1 + 1, len(rows[r]))):
                if rows[r][c] in (None, ""):
                    rows[r][c] = v
                    filled += 1
    return filled


def sheet_rows(path, sheet, fill_merged_cells=True):
    """读一个 sheet 的全部行（值列表）。合并单元格的值会被填满。"""
    if openpyxl is None:
        raise RuntimeError("需要 openpyxl")
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        ws = wb[sheet]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        if fill_merged_cells:
            fill_merged(rows, list(ws.merged_cells.ranges))
        return rows
    finally:
        wb.close()


def list_sheets(path):
    if openpyxl is None:
        raise RuntimeError("需要 openpyxl")
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def read_all_sheets(path, max_scan=25):
    """把所有 sheet 里**认得出表头**的表都读出来。

    返回 (records, report)：
      records  [{"__sheet": 表名, "name": ..., "cls": ..., "total": ..., "raw": {列名: 值}}]
      report   {"sheets": [...], "used": [...], "skipped": [(表名, 原因)], "merged": 次数}
    认不出表头的 sheet 记进 skipped 并写明原因，不静默丢弃。
    """
    if openpyxl is None:
        raise RuntimeError("需要 openpyxl")
    report = {"sheets": [], "used": [], "skipped": [], "merged": 0}
    records = []
    for name in list_sheets(path):
        report["sheets"].append(name)
        try:
            rows = sheet_rows(path, name)
        except Exception as e:
            report["skipped"].append((name, "打不开：%s" % e))
            continue
        hi, hdr = find_header(rows, max_scan=max_scan)
        if hi is None:
            report["skipped"].append((name, "前 %d 行里找不到「姓名」或「学号」" % max_scan))
            continue
        cmap = canon(hdr)
        used = 0
        for row in rows[hi + 1:]:
            if row is None:
                continue
            vals = list(row)
            rec = {"__sheet": name, "raw": {}}
            # 原始列名 -> 值（保留全部列，供逐题使用）
            for i, h in enumerate(hdr):
                hh = ("" if h is None else str(h)).strip()
                if hh and i < len(vals):
                    rec["raw"][hh] = vals[i]
            for key, idx in cmap.items():
                rec[key] = vals[idx] if idx < len(vals) else None
            if rec.get("name") in (None, "") and rec.get("sid") in (None, ""):
                continue
            records.append(rec)
            used += 1
        report["used"].append((name, used))
        report["merged"] += 0
    return records, report


QNUM_PATTERNS = [
    re.compile(r"^(\d{1,3})$"),                 # 1, 2, 3 …
    re.compile(r"^第?(\d{1,3})题$"),             # 第1题 / 1题
    re.compile(r"^[qQtT](\d{1,3})$"),           # Q1 / T1
    re.compile(r"^(\d{1,3})[.、]"),              # 1. / 1、
]


def detect_question_cols(header, exclude=None):
    """认出「每题一列」的题号列，返回 [(列下标, 题号字符串)]。

    只认形态明确的列名；认不出就返回空列表，不做模糊猜测。
    """
    exclude = set(exclude or [])
    out = []
    for i, h in enumerate(header):
        if i in exclude:
            continue
        s = str(h).strip() if h is not None else ""
        if not s:
            continue
        for pat in QNUM_PATTERNS:
            m = pat.match(s)
            if m:
                out.append((i, m.group(1)))
                break
    return out


def wide_to_long(header, rows, student_key, kp_map=None, full_map=None,
                 exclude=None):
    """宽表 -> 长表。

    输入  header 表头；rows 数据行；student_key 学生标识列下标
          kp_map {题号: 知识点}（可选）；full_map {题号: 满分}
    输出  (long_rows, missing)
          long_rows [{"student":.., "qno":.., "kp":.., "full":.., "got":..}]
          missing   缺满分的题号列表 —— 非空时调用方必须处理，不许默认为满分

    口径：满分未知的题一律不写估计值，只把题号列进 missing 让人补。
    """
    qcols = detect_question_cols(header, exclude=exclude)
    if not qcols:
        return [], []
    missing = [q for _, q in qcols if not full_map or q not in full_map]
    out = []
    for row in rows:
        if row is None:
            continue
        vals = list(row)
        stu = vals[student_key] if student_key is not None and student_key < len(vals) else None
        if stu in (None, ""):
            continue
        for idx, q in qcols:
            if idx >= len(vals):
                continue
            got = vals[idx]
            if got in (None, ""):
                continue
            out.append({
                "student": stu,
                "qno": q,
                "kp": (kp_map or {}).get(q),
                "full": (full_map or {}).get(q),
                "got": got,
            })
    return out, missing
