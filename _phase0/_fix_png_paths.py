"""一次性维护脚本：把日志/明细中的图片绝对路径同步到 图片/问题N/（2026-09-11 图片目录迁移）。

匹配形如 `…数学建模大赛\问题2\紧急购电量分布图.png` 的 Windows 绝对路径，
替换为 `…数学建模大赛\图片\问题2\紧急购电量分布图.png`。
只处理 .txt/.csv/.md；run_*.py 的出图代码已另行改为 FIGDIR，不在此处理。
"""

import glob
import io
import re

PAT = re.compile(r"\\问题([1-4])\\([^\\\s/()]*\.png)")

total = 0
for path in sorted(glob.glob("问题*/*.txt") + glob.glob("问题*/*.csv") + glob.glob("问题*/*.md")):
    try:
        s = io.open(path, encoding="utf-8").read()
    except (FileNotFoundError, UnicodeDecodeError):
        continue
    s2, cnt = PAT.subn(lambda m: "\\图片\\问题" + m.group(1) + "\\" + m.group(2), s)
    if cnt:
        io.open(path, "w", encoding="utf-8", newline="").write(s2)
        print("%-40s %d 处" % (path, cnt))
        total += cnt
print("合计:", total)
