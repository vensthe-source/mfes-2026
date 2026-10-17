# -*- coding: utf-8 -*-
"""MFES-2026 四案例复算脚本：从 CSV 独立重算三轨得分、总分与等级，并与白皮书公布值核对。
用法：python recompute_cases.py
"""
import csv
import io
import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

CSV = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output", "cases_board_scores.csv")

EXPECTED = {
    "《霸王别姬》（1993）": (856, "优秀"),
    "《给阿嬷的情书》（2026）": (819, "优秀"),
    "《流浪地球2》（2023）": (809, "优秀"),
    "《功夫女足》（2026）": (560, "一般"),
}

def grade(s):
    if s >= 900:
        return "传世之作"
    if s >= 800:
        return "优秀"
    if s >= 700:
        return "良好"
    if s >= 600:
        return "合格"
    return "一般"

def main():
    data = {}
    with io.open(CSV, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            film = row["film"]
            data.setdefault(film, {"本体翼": 0, "价值翼": 0, "文明翼": 0,
                                   "max本体翼": 0, "max价值翼": 0, "max文明翼": 0})
            track = row["track"]
            mx, sc = int(row["max"]), int(row["score"])
            assert 0 <= sc <= mx, "分数越界: %s/%s %s %d>%d" % (film, track, row["board"], sc, mx)
            data[film][track] += sc
            data[film]["max" + track] += mx

    ok_all = True
    print("film | B/500 | V/300 | C/200 | S | 等级 | 与公布值核对")
    for film, exp in EXPECTED.items():
        d = data[film]
        for track, total in [("本体翼", 500), ("价值翼", 300), ("文明翼", 200)]:
            assert d["max" + track] == total, "%s %s 满分合计=%d ≠ %d（N/A 再分配应保证合计不变）" % (
                film, track, d["max" + track], total)
        s = d["本体翼"] + d["价值翼"] + d["文明翼"]
        g = grade(s)
        ok = (s, g) == exp
        ok_all &= ok
        print("%s | %d | %d | %d | %d | %s | %s" % (
            film, d["本体翼"], d["价值翼"], d["文明翼"], s, g, "一致 OK" if ok else "不一致 MISMATCH"))
    print("复算结论：%s" % ("全部一致，白皮书公布值可复现。", "存在不一致，请检查数据。")[not ok_all])
    return 0 if ok_all else 1

if __name__ == "__main__":
    raise SystemExit(main())
