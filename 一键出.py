# -*- coding: utf-8 -*-
"""一条命令出「哪方面不行」。

背景：集体补漏、个体讲评、学情卡原本是三条命令、三套参数，
老师记不住，也不会用。这个文件把它们串成一次调用。

    python 一键出.py --scores 逐题得分.csv
    python 一键出.py --scores 逐题得分.csv --student 张三
    python 一键出.py --scores A.csv B.csv --student 张三

输入：逐题得分表（学生, 题号, 知识点, 满分, 得分）。
      这份表就是「答题卡对了、按知识点分好类」之后的产物。
输出：输出/薄弱项清单_<时间>.md —— 集体补漏排在前面，个体讲评跟在后头。

口径：一棵草都不重写。全部调用 teaching.py 里已有的函数，
      这样它出的数和单跑那些命令出的数是同一个数。
"""
import argparse
import contextlib
import datetime
import io
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

import teaching

OUT_DIR = os.path.join(HERE, "输出")


class _NS(object):
    """给 teaching 里那些命令用的参数容器。

    teaching 的命令会读 args 上的一堆字段（budget / top / xlsx …），
    这里缺哪个就返回 None，不去猜它该是什么值——猜了反而会把别人的默认值覆盖掉。
    """

    def __init__(self, **kw):
        # 值为 None 的键直接丢掉：留下它会让 __getattr__ 失效，
        # 下游 int(None) 就炸 —— 缺字段应该走下面的默认值。
        self.__dict__.update({k: v for k, v in kw.items() if v is not None})

    # 与 teaching.py 里 argparse 的默认值保持一致：
    # 缺字段时给命令行的默认值，而不是 None —— None 会让 int()/比较直接炸。
    _DEFAULTS = {
        "top": 5, "budget": 40, "min_rate": 0.5, "quiz_top": 8,
        "per": 6, "seed": 20261028, "time_by": "rate",
        "src": "db", "xlsx": None, "db": None, "rows": None,
    }

    def __getattr__(self, name):
        return self._DEFAULTS.get(name)


FOOTER = """
================================================================================
口径与局限（每次都打，不因结果好看就不打）
================================================================================
· 集体讲的判定是**规则**，不是因果：该知识点错误率 ≥ 40% 且失分同学 ≥ 3 人。
· 只有逐题得分才看得出「哪几题」，没有它一律写「不可用」，不按总分估算。
· 个体讲评只含这名学生本人失分的题，不塞别人的题。
· 这里不评价老师、不下「因为所以」的结论。
"""


def _capture(fn, ns):
    """跑一个 teaching 里的命令，把它的控制台输出抓下来。返回 (rc, 文本)。"""
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            rc = fn(ns)
    except SystemExit as e:
        rc = e.code if isinstance(e.code, int) else 2
    except Exception as e:                      # 任何意外都要留住已产出的部分
        buf.write("\n（这一步出错了：%s）\n" % e)
        rc = 2
    return (0 if rc is None else rc), buf.getvalue()


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="一条命令出薄弱项：集体补漏 +（可选）某个学生的讲评清单")
    ap.add_argument("--scores", nargs="+", required=True,
                    help="逐题得分表（学生,题号,知识点,满分,得分），可以给多份")
    ap.add_argument("--student", default=None,
                    help="只出这个学生的讲评清单；不给就只出集体补漏")
    ap.add_argument("--bank", default=None, help="题库.json（默认用同目录那份）")
    ap.add_argument("--out-dir", default=None, help="输出目录（默认 教师工具/输出）")
    a = ap.parse_args(argv)

    parts = []
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M")
    head = "薄弱项清单｜%s\n生成时间：%s\n数据来源：%s\n" % (
        stamp, datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "、".join(os.path.basename(p) for p in a.scores))
    parts.append(head)

    # 一、集体补漏
    g = _NS(scores=a.scores, bank=a.bank)
    rc1, out1 = _capture(teaching.cmd_gaps, g)
    parts.append(out1)

    # 二、个体讲评（可选）
    rc2, out2 = 0, ""
    if a.student:
        s = _NS(student=a.student, scores=a.scores[0], bank=a.bank)
        rc2, out2 = _capture(teaching.cmd_review_student, s)
        parts.append(out2)

    text = "\n".join(parts) + FOOTER

    out_dir = a.out_dir or OUT_DIR
    os.makedirs(out_dir, exist_ok=True)
    name = "薄弱项清单_%s%s.md" % (stamp, ("_" + a.student) if a.student else "")
    path = os.path.join(out_dir, name)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)

    print(text)
    print("已写入: %s" % path)

    if rc1 != 0:
        return rc1
    return rc2


if __name__ == "__main__":
    sys.exit(main())
