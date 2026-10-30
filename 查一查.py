# -*- coding: utf-8 -*-
"""查一查：把题库和知识点树当成一个能查的知识库。

为什么不做向量检索（RAG）：
    题库是结构化的——每道题都写着题号、知识点、满分、难度、错因。
    这种东西用题号和一串关键词直接查，又快又准，还不用装任何东西。
    向量化是把「看不懂的数据」变成「能比距离的向量」；这批数据本来就看得懂。
    只用 Python 标准库 + 同目录的题库.json、teaching.py 里的知识点树。
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


def load_bank():
    if not os.path.exists(BANK):
        return []
    with open(BANK, encoding="utf-8") as f:
        return json.load(f).get("items") or []


def load_nodes():
    """知识点树的正本在 teaching.py 里（teaching_out 那个是产物）。"""
    try:
        import teaching
        return list(teaching.NODES), dict(teaching.ALIAS)
    except Exception:
        return [], {}


def resolve_node(word, nodes, alias):
    w = (word or "").strip()
    if not w:
        return None
    if w in alias:
        nid = alias[w]
        for n in nodes:
            if n.get("id") == nid:
                return n
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


def show_node(word, items, nodes, alias):
    node = resolve_node(word, nodes, alias)
    if not node:
        print("没找到「%s」这个知识点。" % word)
        print("可查的：" + "、".join(n.get("name") for n in nodes))
        return 2
    print("=" * 70)
    print("%s（%s）" % (node.get("name"), node.get("chapter")))
    if node.get("aliases"):
        print("别名：%s" % "、".join(node["aliases"]))
    print("=" * 70)

    up = upstream(nodes, node)
    if not up:
        print("\n前置：没有（这是起点）")
    else:
        print("\n前置链（要学会它，得先会这些）：")
        for d, n in up:
            print("  " + "  " * d + ("└ " if d else "") + n.get("name"))

    down = downstream(nodes, node.get("id"))
    if down:
        print("\n它卡住会影响（这些知识点都依赖它）：")
        print("  " + "、".join(n.get("name") for n in down))

    its = items_of_kp(items, node)
    print("\n题库里的题：%d 道" % len(its))
    if not its:
        print("  （这个知识点还没有题——复习时得自己找题）")
    for i in its[:12]:
        print("  第%-4s %s分 [%s]  %s" % (i.get("no"), i.get("full"), i.get("difficulty"),
                                        ((i.get("stem") or "").strip() or "(题面未录)")[:46]))
        if i.get("answer"):
            print("        答案：%s" % str(i.get("answer"))[:60])
    errs = node.get("errs") or []
    if errs:
        print("\n常见错法：")
        for e in errs:
            print("  · " + e)
    return 0


def show_item(no, items, nodes, alias):
    no = str(no).strip()
    hit = [i for i in items if str(i.get("no")) == no]
    if not hit:
        print("题库里没有第 %s 题。" % no)
        return 2
    i = hit[0]
    print("=" * 70)
    print("第 %s 题 ｜ %s ｜ %s 分 ｜ 难度 %s" % (i.get("no"), i.get("kp"), i.get("full"), i.get("difficulty")))
    print("=" * 70)
    print("\n题面：" + (i.get("stem") or "(无)"))
    for o in (i.get("options") or []):
        print("   " + str(o))
    print("\n答案：" + (i.get("answer") or "(无)"))
    if i.get("err"):
        print("典型错因：" + str(i["err"]))
    node = resolve_node(i.get("kp") or "", nodes, alias)
    if node:
        up = upstream(nodes, node)
        print("\n考的知识点：" + node.get("name"))
        if up:
            print("要先会：" + "、".join(n.get("name") for _d, n in up))
    return 0


def search(word, items, nodes, alias):
    w = (word or "").strip()
    if not w:
        print("给个关键词")
        return 2
    hits = []
    for i in items:
        blob = " ".join(str(i.get(k) or "") for k in ("stem", "answer", "kp", "err", "no"))
        if w in blob:
            hits.append(i)
    node = resolve_node(w, nodes, alias)
    print("=" * 70)
    print("搜「%s」：题库命中 %d 道题" % (w, len(hits)))
    print("=" * 70)
    if node:
        print("它也是一个知识点：%s（用「查一查.py %s」看详情）" % (node.get("name"), node.get("name")))
        print()
    for i in hits[:15]:
        print("  第%-4s %-14s %s" % (i.get("no"), i.get("kp"), (i.get("stem") or "")[:50]))
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


def show_list(items, nodes):
    print("=" * 70)
    print("知识点一览（题库覆盖情况）")
    print("=" * 70)
    for n in nodes:
        k = len(items_of_kp(items, n))
        mark = ("有题 %d 道" % k) if k else "★ 还没有题"
        print("  %-16s %-10s 前置 %d 个" % (n.get("name"), mark, len(n.get("prereq") or [])))
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="查一查：题库 + 知识点树，零依赖")
    ap.add_argument("word", nargs="?", help="知识点名称，例如 绝对值")
    ap.add_argument("--题", dest="no", help="看某一题的详情，例如 --题 19")
    ap.add_argument("--找", dest="kw", help="在题干/答案/知识点/错因里搜关键词")
    ap.add_argument("--错因", dest="cuo", action="store_true", help="错因分布")
    ap.add_argument("--清单", dest="list", action="store_true", help="知识点一览与题库覆盖")
    a = ap.parse_args(argv)

    items = load_bank()
    nodes, alias = load_nodes()
    if not items:
        print("读不到题库.json（应该和这个脚本放在同一目录）")
        return 2

    if a.list:
        return show_list(items, nodes)
    if a.cuo:
        return show_errs(items)
    if a.kw:
        return search(a.kw, items, nodes, alias)
    if a.no:
        return show_item(a.no, items, nodes, alias)
    if a.word:
        return show_node(a.word, items, nodes, alias)

    print("查一查：题库 + 知识点树，零依赖")
    print()
    print("  python 查一查.py 绝对值          一个知识点：前置、影响、题库里的题、常见错法")
    print("  python 查一查.py --题 19         第 19 题")
    print("  python 查一查.py --找 乘方        全库搜关键词")
    print("  python 查一查.py --错因          错因分布")
    print("  python 查一查.py --清单          16 个知识点与题库覆盖")
    return 0


if __name__ == "__main__":
    sys.exit(main())
