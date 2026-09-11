"""Phase 0 探针 1：逐个附件的结构探测（工作表名、规模、首尾行），输出 UTF-8 报告。"""

import openpyxl
import os

WORK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILES = [
    "附件/附件1.xlsx",
    "附件/附件2.xlsx",
    "附件/附件3.xlsx",
    "附件/附件4.xlsx",
    "附件/附件5/result1.xlsx",
    "附件/附件5/result2.xlsx",
    "附件/附件5/result3.xlsx",
    "附件/附件5/result4-2.xlsx",
    "附件/附件5/result4-3.xlsx",
]

out = []
for rel in FILES:
    path = os.path.join(WORK, rel.replace("/", os.sep))
    out.append("=" * 78)
    out.append("文件: %s" % rel)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        nrow = len(rows)
        ncol = max((len(r) for r in rows), default=0)
        out.append("  工作表: %s | 行数=%d | 列数=%d" % (ws.title, nrow, ncol))
        for i in range(min(4, nrow)):
            out.append("    首[%d]: %s" % (i, rows[i]))
        if nrow > 6:
            out.append("     ...")
        for i in range(max(4, nrow - 2), nrow):
            out.append("     ���[%d]: %s" % (i, rows[i]))
    wb.close()

os.makedirs(os.path.join(WORK, "_phase0"), exist_ok=True)
with open(os.path.join(WORK, "_phase0", "报告01_文件结构.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(out))
print("done")
