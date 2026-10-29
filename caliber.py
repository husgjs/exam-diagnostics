# -*- coding: utf-8 -*-
"""口径说明：把「这个数字怎么算的」跟着产物一起交出去。

为什么单独一份、而不是往每个 xlsx 里插一页：
    插入就要重新保存工作簿，打印设置、列宽、分页符都可能被改掉。
    打印版一乱，老师家长会现场就没法用了。
    所以宁可多交一个文件，也不动已经生成的表。

为什么只写「本次实际取值」，不写通用定义：
    老师在家长会上要念的是「30 人参加、剔除 2 人缺考」，
    不是「有效样本指……」。数字才有用。
"""
import os
import time

import openpyxl
from openpyxl.styles import Font, Alignment


def _rows_from_cards(cards, warn=""):
    """把首页那几张卡里的 KPI 翻译成口径行。"""
    out = []
    for card in cards or []:
        title = str(card.get("title") or "").strip()
        for kv in card.get("kpis") or []:
            k = str(kv.get("k") or "").strip()
            v = str(kv.get("v") or "").strip()
            if k and v:
                out.append((title, k, v))
        note = str(card.get("note") or "").strip()
        if note:
            out.append((title, "说明", note))
    if warn:
        out.append(("需要留意", "提示", warn))
    return out


def build(outdir, cards=None, src_name="", warn="", job_id=""):
    """写出一份 口径说明.xlsx，返回路径。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "口径说明"

    ws.append(["这一段念给家长听，就是这句话的依据"])
    ws.merge_cells("A1:D1")
    ws["A1"].font = Font(bold=True, size=12)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")

    ws.append([])
    ws.append(["来源文件", src_name or "(未登记)", "", ""])
    ws.append(["生成时间", time.strftime("%Y-%m-%d %H:%M"), "", ""])
    if job_id:
        ws.append(["任务编号", job_id, "", ""])

    ws.append([])
    ws.append(["卡片", "指标", "本次实际取值", "怎么算的"])
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)

    # 通用口径：写死的规则，逐条对应 README 里那五条
    RULES = {
        "实际参加考试": "剔除了总分 0（缺考/未录入）的人；名单里仍保留并标「缺」",
        "总分平均": "只算实际参加考试的人；缺考的不按 0 分计入",
        "及格率": "及格人数 ÷ 实际参加考试人数",
        "优秀率": "优秀人数 ÷ 实际参加考试人数",
        "得分率": "该科得分合计 ÷ 满分合计（不是原始分直接比）",
        "临界生": "落在「目标线 − 临界带」到「目标线 + 10 分」之间的人",
        "缺考/未录入": "总分 0 或该格写着「缺考/缺/-/—」；不进任何均分与名次",
        "错误率": "1 − 该知识点得分合计 ÷ 满分合计",
        "涉及人数": "该知识点至少错一题的人数",
        "区分度": "高分组得分率 − 低分组得分率，高低分组各取 27%",
        "讲评时长": "按每题失分占比分配总时长；单题下限 3 分钟",
        "合计": "各题讲评时长之和，与设定上限对比",
        "最高分": "实际参加考试的人里最高的总分（缺考不参与）",
        "最低分": "实际参加考试的人里最低的总分（缺考不参与）",
        "满分": "各科满分相加",
        "及格线": "满分的 60%",
        "优秀线": "满分的 85%",
        "相差": "两个班（或两个人）之间的差值，正数表示前者更高",
    }

    for title, k, v in _rows_from_cards(cards, warn):
        how = RULES.get(k, "")
        if not how:
            for key, rule in RULES.items():
                if key in k:
                    how = rule
                    break
        ws.append([title, k, v, how])

    ws.append([])
    ws.append(["固定规则（每次输出都适用）", "", "", ""])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
    for key in ("缺考/未录入", "及格率", "得分率", "临界生", "区分度"):
        ws.append(["", key, "", RULES[key]])

    ws.append([])
    ws.append(["数据不足时的处理", "", "", ""])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
    ws.append(["", "人数少于 3 人", "不出均分 / 及格率 / 名次",
               "改出个体数据并注明「样本 N 人，不排名次」，不硬算群体结论"])
    ws.append(["", "没有逐题得分", "不出版难度 / 区分度 / 讲评清单",
               "明确写「不可用（缺逐题得分）」，绝不按总分估算"])
    ws.append(["", "满分不知道", "不判断是否超分",
               "在读表层写明「不知道满分」，不用列最大值反推"])

    for i, w in enumerate([22, 20, 30, 62], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)

    path = os.path.join(outdir, "口径说明.xlsx")
    wb.save(path)
    return path
