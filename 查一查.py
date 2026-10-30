# -*- coding: utf-8 -*-
"""查一查：按科目查知识点与题库。零依赖，不联网。

三科的结构不一样，不能套同一个模型：
    数学  prereq     链式：严格前置依赖，先会 A 才学得动 B
    英语  strand     网状：词汇/语法/听说/阅读/写作几条线并行，线内可排，线间不套前置
    语文  dimension  维度：积累/现代文/文言文/写作，彼此几乎无依赖

    给语文算「前置链」会算出一个假的顺序，所以这里按 model 决定显示什么。
"""
import argparse
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

BANK = os.path.join(HERE, "题库.json")
SUBJECT_DIR = os.path.join(HERE, "科目")
DEFAULT_SUBJECT = "数学"


def load_bank():
    if not os.path.exists(BANK):
        return []
    with open(BANK, encoding="utf-8") as f:
        return json.load(f).get("items") or []


def list_subjects():
    if not os.path.isdir(SUBJECT_DIR):
        return []
    return sorted(n[:-5] for n in os.listdir(SUBJECT_DIR) if n.endswith(".json"))


def load_subject(name):
    p = os.path.join(SUBJECT_DIR, "%s.json" % name)
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            meta = json.load(f)
        nodes = meta.get("nodes") or []
        alias = {}
        for n in nodes:
            for k in [n.get("name"), n.get("id")] + list(n.get("aliases") or []):
                if k:
                    alias.setdefault(str(k), n)
        return meta, nodes, alias
    try:
        import teaching
        nodes = list(teaching.NODES)
        alias = {}
        for k, v in teaching.ALIAS.items():
            for n in nodes:
                if n.get("id") == v:
                    alias[k] = n
                    break
        return {"subject": "数学", "model": "prereq", "nodes": nodes}, nodes, alias
    except Exception:
        return {"subject": name, "model": "prereq", "nodes": []}, [], {}


def resolve_node(word, nodes, alias):
    w = (word or "").strip()
    if not w:
        return None
    if w in alias:
        return alias[w]
    for n in nodes:
        if w == n.get("name") or w == n.get("id"):
            return n
    for n in nodes:
        if w in (n.get("name") or "") or any(w in a for a in (n.get("aliases") or [])):
            return n
    return None


def by_id(nodes, nid):
    for n in nodes:
        if n.get("id") == nid:
            return n
    return None


def upstream(nodes, node, depth=0, seen=None):
    seen = seen or set()
    out = []
    for pid in node.get("prereq") or []:
        if pid in seen:
            continue
        seen.add(pid)
        n = by_id(nodes, pid)
        if not n:
            continue
        out.append((depth, n))
        out += upstream(nodes, n, depth + 1, seen)
    return out


def downstream(nodes, nid):
    return [n for n in nodes if nid in (n.get("prereq") or [])]


def items_of_kp(items, node):
    names = {node.get("name")} | set(node.get("aliases") or [])
    return [i for i in items if (i.get("kp") or "").strip() in names]


def group_of(meta, node):
    model = meta.get("model")
    if model == "strand":
        for s in meta.get("strands") or []:
            if s.get("id") == node.get("strand"):
                return ("所属线", s)
    if model == "dimension":
        for d in meta.get("dimensions") or []:
            if d.get("id") == node.get("dimension"):
                return ("所属维度", d)
    return (None, None)


def show_node(word, items, meta, nodes, alias):
    node = resolve_node(word, nodes, alias)
    if not node:
        print("「%s」里没找到「%s」。" % (meta.get("subject"), word))
        if nodes:
            print("可查的：" + "、".join(n.get("name") for n in nodes[:16]))
        return 2

    model = meta.get("model")
    print("=" * 70)
    print("%s ｜ %s" % (meta.get("subject"), node.get("name")))
    if node.get("aliases"):
        print("别名：%s" % "、".join(node["aliases"]))
    if node.get("stub"):
        print("★ " + node["stub"])
    print("=" * 70)

    label, g = group_of(meta, node)
    if label:
        print("\n%s：%s%s" % (label, g.get("name"), ("（%s）" % g["note"]) if g.get("note") else ""))

    if model == "prereq":
        up = upstream(nodes, node)
        if not up:
            print("\n前置：没有（这是起点）")
        else:
            print("\n前置链（要学会它，得先会这些）：")
            for d, n in up:
                print("  " + "  " * d + ("└ " if d else "") + n.get("name"))
        down = downstream(nodes, node.get("id"))
        if down:
            print("\n它卡住会影响：" + "、".join(n.get("name") for n in down))
    else:
        print("\n（这个科目不套前置链：%s）" % meta.get("model_note", "结构不同"))

    its = items_of_kp(items, node)
    print("\n题库里的题：%d 道" % len(its))
    if not its:
        print("  （这个知识点还没有题）")
    for i in its[:12]:
        stem = (i.get("stem") or "").strip() or "(题面未录)"
        print("  第%-4s %s分 [%s]  %s" % (i.get("no"), i.get("full"), i.get("difficulty"), stem[:46]))
        if i.get("answer"):
            print("        答案：%s" % str(i.get("answer"))[:60])
    errs = node.get("errs") or []
    if errs:
        print("\n常见错法：")
        for e in errs:
            print("  · " + e)
    return 0


def show_item(no, items, meta, nodes, alias):
    no = str(no).strip()
    hit = [i for i in items if str(i.get("no")) == no]
    if not hit:
        print("题库里没有第 %s 题。" % no)
        return 2
    i = hit[0]
    print("=" * 70)
    print("第 %s 题 ｜ %s ｜ %s 分 ｜ 难度 %s" % (i.get("no"), i.get("kp"), i.get("full"), i.get("difficulty")))
    print("=" * 70)
    print("\n题面：" + ((i.get("stem") or "").strip() or "(题面未录)"))
    for o in (i.get("options") or []):
        print("   " + str(o))
    print("\n答案：" + (i.get("answer") or "(无)"))
    if i.get("err"):
        print("典型错因：" + str(i["err"]))
    node = resolve_node(i.get("kp") or "", nodes, alias)
    if node:
        print("\n考的知识点：" + node.get("name"))
        if meta.get("model") == "prereq":
            up = upstream(nodes, node)
            if up:
                print("要先会：" + "、".join(n.get("name") for _d, n in up))
    return 0


def search(word, items, meta, nodes, alias):
    w = (word or "").strip()
    if not w:
        print("给个关键词")
        return 2
    hits = [i for i in items
            if w in " ".join(str(i.get(k) or "") for k in ("stem", "answer", "kp", "err", "no"))]
    print("=" * 70)
    print("搜「%s」：题库命中 %d 道题" % (w, len(hits)))
    print("=" * 70)
    node = resolve_node(w, nodes, alias)
    if node:
        print("它也是一个知识点：%s（用「查一查.py %s」看详情）" % (node.get("name"), node.get("name")))
        print()
    for i in hits[:15]:
        print("  第%-4s %-14s %s" % (i.get("no"), i.get("kp"), ((i.get("stem") or "").strip() or "(题面未录)")[:48]))
    if len(hits) > 15:
        print("  …还有 %d 道" % (len(hits) - 15))
    return 0


def show_errs(items):
    from collections import Counter
    c = Counter(str(i.get("err") or "(未标)") for i in items)
    print("=" * 70)
    print("错因分布（按题目标注统计，不是按人）")
    print("=" * 70)
    for err, n in c.most_common():
        kps = sorted({i.get("kp") for i in items if str(i.get("err") or "(未标)") == err and i.get("kp")})
        print("  %-12s %2d 道  ｜ %s" % (err, n, "、".join(kps[:6])))
    return 0


def show_list(items, meta, nodes):
    print("=" * 70)
    print("%s ｜ 知识点一览（题库覆盖）" % meta.get("subject"))
    print("=" * 70)
    model = meta.get("model")
    if model == "strand":
        for s in meta.get("strands") or []:
            print("  【%s】%s" % (s.get("name"), s.get("note") or ""))
            for n in nodes:
                if n.get("strand") == s.get("id"):
                    k = len(items_of_kp(items, n))
                    print("     %-22s %s" % (n.get("name"), ("有题 %d 道" % k) if k else "还没有题"))
    elif model == "dimension":
        for d in meta.get("dimensions") or []:
            print("  【%s】%s" % (d.get("name"), d.get("note") or ""))
            for n in nodes:
                if n.get("dimension") == d.get("id"):
                    k = len(items_of_kp(items, n))
                    print("     %-22s %s" % (n.get("name"), ("有题 %d 道" % k) if k else "还没有题"))
    else:
        for n in nodes:
            k = len(items_of_kp(items, n))
            print("  %-18s %-12s 前置 %d 个" % (n.get("name"),
                                            ("有题 %d 道" % k) if k else "★ 还没有题",
                                            len(n.get("prereq") or [])))
    return 0


def main(argv=None):
    subs = list_subjects()
    ap = argparse.ArgumentParser(description="查一查：按科目查知识点与题库，零依赖")
    ap.add_argument("word", nargs="?", help="知识点名称，例如 绝对值")
    ap.add_argument("--科", dest="subject", default=DEFAULT_SUBJECT,
                    help="科目（%s），默认 %s" % ("、".join(subs) or "数学", DEFAULT_SUBJECT))
    ap.add_argument("--题", dest="no", help="看某一题的详情")
    ap.add_argument("--找", dest="kw", help="在题干/答案/知识点/错因里搜关键词")
    ap.add_argument("--错因", dest="cuo", action="store_true", help="错因分布")
    ap.add_argument("--清单", dest="lst", action="store_true", help="知识点一览与题库覆盖")
    a = ap.parse_args(argv)

    items = load_bank()
    meta, nodes, alias = load_subject(a.subject)
    if not items:
        print("读不到题库.json")
        return 2

    if a.lst:
        return show_list(items, meta, nodes)
    if a.cuo:
        return show_errs(items)
    if a.kw:
        return search(a.kw, items, meta, nodes, alias)
    if a.no:
        return show_item(a.no, items, meta, nodes, alias)
    if a.word:
        return show_node(a.word, items, meta, nodes, alias)

    print("查一查：按科目查知识点与题库，零依赖")
    print()
    print("  科目：" + ("、".join(subs) if subs else "数学"))
    print()
    print("  python 查一查.py 绝对值                数学·绝对值：前置链、影响范围、题目、错法")
    print("  python 查一查.py --科 英语             英语的线/维度与覆盖")
    print("  python 查一查.py --科 语文 --清单      语文的维度与覆盖")
    print("  python 查一查.py --题 19               某一题")
    print("  python 查一查.py --找 乘方             搜关键词")
    print("  python 查一查.py --错因                错因分布")
    return 0


if __name__ == "__main__":
    sys.exit(main())
