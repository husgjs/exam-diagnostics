
# -*- coding: utf-8 -*-
"""一节课任务单 —— 分不了班，就分任务。

输入一个知识点 + 一节课的时长，输出可直接打印的 A4：
  第 1 页  教师用：本课设计（前置点 / 时间分配 / 每题讲评要点）
  第 2 页  学生 A 卷（保底）：课前检测 + 全班必做 + 还没过的人的再练
  第 3 页  学生 B 卷（拔高）：课前检测 + 全班必做 + 先完成的人的拔高

用法：
  python lesson_plan.py --kp 有理数减法
  python lesson_plan.py --kp 绝对值 --minutes 45 --paper 题库.json --tree 知识点树.json

题库格式（--paper）：[[题号, 知识点, 满分, 错因倾向, 错误率], ...]
   或 [{"no":1,"kp":"相反数","full":3,"err":"概念不清","rate":0.15,"stem":"题干（可选）"}, ...]
"""
import io, sys, os, json, argparse, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
# 优先用题库.json：它有真题干，印出来能直接发；旧的 class_rows.json 只有题号，作兜底
DEFAULT_PAPER = os.path.join(HERE, "题库.json")
FALLBACK_PAPER = os.path.join(os.path.dirname(HERE), "_tmp", "class_rows.json")
DIFF_RATE = {"easy": 0.2, "mid": 0.5, "hard": 0.8}   # 难度 -> 折合错误率，仅用于排序
DEFAULT_TREE = os.path.join(HERE, "teaching_out", "知识点树.json")

CSS = """
@page { size: A4; margin: 14mm 13mm; }
* { box-sizing: border-box; }
body { font-family: "Microsoft YaHei", "PingFang SC", sans-serif; color: #1a1a1a;
       font-size: 11pt; line-height: 1.55; margin: 0; }
.page { page-break-after: always; }
.page:last-child { page-break-after: auto; }
h1 { font-size: 17pt; color: #1F4E79; margin: 0 0 2mm; letter-spacing: .5px; }
h2 { font-size: 12pt; color: #1F4E79; margin: 5mm 0 2mm; border-left: 3px solid #1F4E79;
     padding-left: 2.5mm; }
.sub { color: #666; font-size: 9pt; margin-bottom: 4mm; }
table { width: 100%; border-collapse: collapse; margin: 2mm 0; font-size: 9.5pt; }
th, td { border: 1px solid #c9d3de; padding: 1.6mm 2mm; text-align: left; vertical-align: top; }
th { background: #E8EEF5; font-weight: 600; }
.time { display: flex; gap: 2mm; margin: 3mm 0; }
.time div { flex: 1; border: 1px solid #c9d3de; border-radius: 2mm; padding: 2mm; text-align: center; }
.time b { display: block; font-size: 13pt; color: #1F4E79; }
.time span { font-size: 8.5pt; color: #666; }
.q { margin: 0 0 4mm; page-break-inside: avoid; }
.qh { font-size: 9pt; color: #666; margin-bottom: 1mm; }
.qh b { color: #1F4E79; font-size: 10pt; }
.stem { font-size: 11.5pt; margin: 1mm 0; }
.stem .blank { color: #b00; background: #fdf2f2; border-bottom: 1px dashed #e0b4b4;
               padding: 0 1mm; font-size: 9pt; }
.ln { border-bottom: 1px solid #cfcfcf; height: 8mm; }
.keys { page-break-before: always; }
.kv { color: #1F4E79; font-weight: 600; }
.tip { background: #FFF9E6; border-left: 3px solid #E0A800; padding: 2mm 3mm;
       margin: 3mm 0; font-size: 9.5pt; }
.warn { background: #FDF2F2; border-left: 3px solid #C0392B; padding: 2mm 3mm;
        margin: 3mm 0; font-size: 9.5pt; }
.foot { color: #999; font-size: 8pt; margin-top: 4mm; }
.fill { font-size: 9.5pt; color: #333; margin: 2mm 0 3mm; }
.fill span { display: inline-block; border-bottom: 1px solid #999; min-width: 28mm; }
@media print { body { font-size: 10.5pt; } .noprint { display: none; } }
"""

def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

def load_paper(path):
    if not os.path.exists(path) and os.path.exists(FALLBACK_PAPER):
        path = FALLBACK_PAPER
    if not os.path.exists(path):
        raise SystemExit("找不到题库文件：%s" % path)
    d = json.load(open(path, encoding="utf-8"))
    # 题库.json：{"meta":..,"items":[{no,kp,full,difficulty,err,stem,answer}]}
    if isinstance(d, dict) and "items" in d:
        out, skipped = [], 0
        for x in d["items"]:
            if not str(x.get("stem") or "").strip():
                skipped += 1          # 没题干就不印占位，直接不进选题池
                continue
            out.append({"no": x.get("no"), "kp": str(x.get("kp") or ""),
                        "full": float(x.get("full") or 1), "err": str(x.get("err") or ""),
                        "rate": DIFF_RATE.get(str(x.get("difficulty") or "mid"), 0.5),
                        "stem": str(x.get("stem") or ""),
                        "answer": str(x.get("answer") or "")})
        if skipped:
            print("题库共 %d 条，其中 %d 条还没有题干，这次不印（补上就会自动进来）"
                  % (len(d["items"]), skipped))
        if not out:
            raise SystemExit("题库里还没有一条录了题干的题，任务单印不出来。\n"
                             "先把题面填进 题库.json 的 stem 字段，或跑 teacher_app 的题库录入器。")
        return out
    rows = d["paper"] if isinstance(d, dict) and "paper" in d else d
    out = []
    for i, x in enumerate(rows):
        if isinstance(x, (list, tuple)):
            no, kp, full, err, rate = (list(x) + [None] * 5)[:5]
            stem = None
        else:
            no, kp, full = x.get("no", i + 1), x.get("kp", ""), x.get("full", 1)
            err, rate, stem = x.get("err", ""), x.get("rate"), x.get("stem")
        out.append({"no": no, "kp": str(kp or ""), "full": float(full or 1),
                    "err": str(err or ""), "rate": float(rate) if rate is not None else 0.5,
                    "stem": stem})
    return out

def load_tree(path):
    """返回 (name->node, alias->name, id->name)

    注意：树里 prereq 存的是节点 id（如 rat-numberline），不是名字。
    不做 id->name 映射，出来的就是一串 id，前置链等于没用。
    """
    if not os.path.exists(path):
        return {}, {}, {}
    d = json.load(open(path, encoding="utf-8"))
    nodes = {n["name"]: n for n in (d.get("nodes") or [])}
    id2name = {n.get("id"): n["name"] for n in nodes.values() if n.get("id")}
    alias = {}
    for n in nodes.values():
        alias[n["name"]] = n["name"]
        for a in (n.get("aliases") or []):
            alias[a] = n["name"]
    return nodes, alias, id2name

def canon(kp, alias):
    return alias.get(kp, kp)

def prereq_chain(kp, nodes, id2name, limit=3):
    """由近及远的前置链（返回知识点名字，不是 id）"""
    chain, seen, cur = [], set(), kp
    while cur and cur not in seen and len(chain) < limit:
        seen.add(cur)
        n = nodes.get(cur) or {}
        pres = [p for p in (n.get("prereq") or []) if p]
        if not pres:
            break
        cur = id2name.get(pres[0], pres[0])
        if cur in chain:
            break
        chain.append(cur)
        if cur not in nodes:      # 撞到接口节点就停
            break
    return chain

def pick(items, kp, n, order="easy"):
    pool = [x for x in items if x["kp"] == kp]
    pool.sort(key=lambda x: x["rate"], reverse=(order == "hard"))
    return pool[:n]

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--kp", required=True, help="这节课讲哪个知识点")
    ap.add_argument("--minutes", type=int, default=40)
    ap.add_argument("--paper", default=DEFAULT_PAPER)
    ap.add_argument("--tree", default=DEFAULT_TREE)
    ap.add_argument("--out", default=os.path.join(HERE, "输出"))
    ap.add_argument("--stems", help="题号->题干的 json 文件（可选）")
    a = ap.parse_args()

    items = load_paper(a.paper)
    nodes, alias, id2name = load_tree(a.tree)
    stems = {}
    if a.stems and os.path.exists(a.stems):
        stems = {str(k): v for k, v in json.load(open(a.stems, encoding="utf-8")).items()}

    target = canon(a.kp, alias)
    if not any(canon(x["kp"], alias) == target for x in items):
        known = sorted({x["kp"] for x in items})
        msg = ("题库里没有「%s」这个知识点的题。\n"
               "题库里的原始知识点是：%s\n"
               "（如果名字对不上，在知识点树里给它加一条 aliases 即可）")
        raise SystemExit(msg % (a.kp, "、".join(known)))
    for x in items:                       # 统一成规范名，别名才对得上
        x["kp"] = canon(x["kp"], alias)

    chain = prereq_chain(target, nodes, id2name)
    pre_kp = chain[0] if chain else None

    # 时间分配
    M = a.minutes
    t_pre = 3
    t_main = max(6, int(M * 0.6))
    t_split = max(4, M - t_pre - t_main - 2)

    # 选题
    # 前置点没有题就沿链继续往下找，最多找两层；找到就用它做课前检测
    pre_q, pre_used = [], None
    for cand in (chain[:2] if chain else []):
        got = pick(items, cand, 2, "easy")
        if got:
            pre_q, pre_used = got, cand
            break
    pre_kp = pre_used or pre_kp
    pool = sorted([x for x in items if x["kp"] == target], key=lambda x: x["rate"])
    # 题库小时不能把题全塞进「必做」，否则拔高/再练两段必然空掉
    if len(pool) >= 8:
        n_main, n_each = 4, 2
    elif len(pool) >= 6:
        n_main, n_each = 3, 2
    elif len(pool) >= 4:
        n_main, n_each = 2, 1
    else:
        n_main, n_each = len(pool), 0
    main_q = pool[:n_main]
    rest = pool[n_main:]
    hard_q = sorted(rest, key=lambda x: -x["rate"])[:n_each]
    used = {x["no"] for x in hard_q}
    drill_q = [x for x in rest if x["no"] not in used][:n_each]

    today = datetime.date.today().isoformat()
    node = nodes.get(target) or {}
    errs = node.get("errs") or []
    mis = [x for x in items if x["kp"] == target and x["rate"] >= 0.6]

    def qblock(x, idx, show_kp=True):
        stem = stems.get(str(x["no"])) or x.get("stem")
        if stem:
            body = '<div class="stem">%s</div>' % esc(stem)
        else:
            body = ('<div class="stem"><span class="blank">题面未录入题库 · 请按原卷第 %s 题粘贴</span></div>'
                    % esc(x["no"]))
        kp_txt = ('<b>第 %d 题</b>（原卷第 %s 题 · %s · %g 分）' % (idx, x["no"], x["kp"], x["full"])
                  if show_kp else '<b>第 %d 题</b>' % idx)
        return ('<div class="q"><div class="qh">%s</div>%s'
                '<div class="ln"></div><div class="ln"></div></div>') % (kp_txt, body)

    # ---------- 第 1 页：教师用 ----------
    rows_html = []
    for n, title, qs in [("课前 %d 分钟" % t_pre, "前置检测：%s" % (pre_kp or "无"), pre_q),
                         ("%d 分钟" % t_main, "全班必做：%s" % target, main_q),
                         ("%d 分钟" % t_split, "先完成的人 → 拔高", hard_q),
                         ("%d 分钟" % t_split, "还没过的人 → 同类再练", drill_q)]:
        nos = "、".join(str(x["no"]) for x in qs) or "（题库里没有可用题，需自备）"
        rows_html.append("<tr><td>%s</td><td>%s</td><td>%s</td></tr>" % (n, esc(title), nos))

    mis_html = ""
    if mis:
        mis_html = ("<div class=\"tip\"><b>本知识点上次考试的失分点（错误率 ≥60%%）：</b><br>%s</div>"
                    % "；".join("第 %s 题（错 %.0f%%，主要错因：%s）" %
                               (m["no"], m["rate"] * 100, m["err"] or "未标注") for m in mis))
    err_html = ""
    if errs:
        err_html = ("<div class=\"warn\"><b>这个知识点学生最常犯的错：</b><br>%s</div>"
                    % "；".join(esc(e) for e in errs[:4]))
    chain_html = (" → ".join([target] + chain)) if chain else "%s（知识树里没有前置记录）" % target

    teacher = """
<div class="page">
  <h1>一节课任务单 · %s</h1>
  <div class="sub">%s ｜ %d 分钟 ｜ 班内分层（不拆班、不贴标签，只分任务）</div>

  <h2>时间怎么排</h2>
  <div class="time">
    <div><b>%d'</b><span>课前检测<br>%s</span></div>
    <div><b>%d'</b><span>全班必做<br>保底目标</span></div>
    <div><b>%d'</b><span>分岔<br>拔高 / 再练</span></div>
    <div><b>2'</b><span>收口<br>讲一道错得最多的</span></div>
  </div>

  <h2>这几段各做什么</h2>
  <table><tr><th style="width:16%%">时间</th><th style="width:42%%">做什么</th><th>用哪几题</th></tr>%s</table>

  <h2>为什么要先测前置</h2>
  <div class="tip"><b>前置链：</b>%s<br>
  跟不上的学生，往往不是这节课不会，是这条链上更靠前的一环断了。
  课前 3 分钟不讲新课，只测前置点——谁断了当场暴露，不用等到作业。</div>

  %s
  %s

  <h2>两个发卷原则</h2>
  <table>
    <tr><th style="width:26%%">原则</th><th>说明</th></tr>
    <tr><td>不贴标签</td><td>A 卷叫「保底卷」、B 卷叫「挑战卷」，不发名字、不喊组号；
        想要挑战的可以自己拿 B 卷，卡住的可以回头做 A 卷剩余题。</td></tr>
    <tr><td>保底优先</td><td>B 卷只在完成 A 卷的人手里出现。保底没过的学生，
        这一节课的任务就是把 A 卷做完做对——这比多刷十道题有用。</td></tr>
  </table>

  <div class="foot">生成于 %s ｜ 题库 %d 题 ｜ 任务单不含真实学生姓名</div>
</div>
""" % (esc(target), today, M, t_pre, esc(pre_kp or "无前置"), t_main, t_split,
       "".join(rows_html), esc(chain_html), mis_html, err_html, today, len(items))

    # ---------- 第 2 页：A 卷（保底）----------
    def stu_page(title, tagline, lists, note):
        blocks = []
        i = 0
        for head, qs in lists:
            if not qs: continue
            blocks.append('<h2>%s</h2>' % esc(head))
            for x in qs:
                i += 1
                blocks.append(qblock(x, i))
        return """
<div class="page">
  <h1>%s</h1>
  <div class="sub">%s</div>
  <div class="fill">姓名：<span></span>　班级：<span></span>　用时：<span></span> 分钟</div>
  %s
  <div class="tip">%s</div>
  <div class="foot">做完先自己检查一遍符号和单位，再交给老师。</div>
</div>
""" % (esc(title), esc(tagline), "".join(blocks), esc(note))

    page_a = stu_page("任务单 A · 保底卷",
                      "%s ｜ 先把这几道做对，就算完成今天的任务" % target,
                      [("课前热身（%d 分钟）" % t_pre, pre_q),
                       ("必做（%d 分钟）" % t_main, main_q),
                       ("还有时间就做这些（%d 分钟）" % t_split, drill_q)],
                      "卡住不要空着：写出你能想到的第一步，老师看过程就知道你卡在哪。")

    page_b = stu_page("任务单 B · 挑战卷",
                      "%s ｜ 保底题做完并检查过了，再拿这张" % target,
                      [("课前热身（%d 分钟）" % t_pre, pre_q),
                       ("必做（%d 分钟）" % t_main, main_q),
                       ("挑战（%d 分钟）" % t_split, hard_q)],
                      "挑战题做不出来是正常的——把想到的思路写下来，比空着有价值。")

    html = ("<!DOCTYPE html><html lang=\"zh-CN\"><head><meta charset=\"utf-8\">"
            "<title>一节课任务单 · %s</title><style>%s</style></head><body>%s%s%s</body></html>"
            % (esc(target), CSS, teacher, page_a, page_b))

    os.makedirs(a.out, exist_ok=True)
    fn = os.path.join(a.out, "任务单_%s.html" % target.replace("/", "_"))
    open(fn, "w", encoding="utf-8").write(html)

    print("知识点: %s" % target)
    print("前置链: %s" % chain_html)
    print("时间: 课前 %d' + 必做 %d' + 分岔 %d' + 收口 2' = %d'" % (t_pre, t_main, t_split, M))
    print()
    print("课前检测(%s): %s" % (pre_kp or "无", [x["no"] for x in pre_q]))
    print("全班必做:      %s" % [x["no"] for x in main_q])
    print("拔高(B卷):     %s" % [x["no"] for x in hard_q])
    print("同类再练(A卷): %s" % [x["no"] for x in drill_q])
    missing = [x["no"] for x in (pre_q + main_q + hard_q + drill_q) if not (stems.get(str(x["no"])) or x.get("stem"))]
    if missing:
        print()
        print("注意: 这 %d 道题没有题干，页面上会显示占位" % len(missing), missing)
        print("      给 --stems 一个 {题号: 题干} 的 json，就能直接印出来")
    if not hard_q or not drill_q:
        print()
        print("提示: 题库里「%s」的题只有 %d 道，已全部用于必做——" % (target, len(pool)))
        print("      拔高 / 同类再练这两段是空的。补几道同类题进题库，这两段会自动填上。")
    print()
    print("已生成:", fn)
    print("页数: 3 （教师用 / A 保底卷 / B 挑战卷）")

if __name__ == "__main__":
    main()
