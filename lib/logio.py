"""控制台日志工具（全工作区共用）：把脚本的标准输出同时写入日志文件。

用途：`roles/编程手.md` §7 要求"代码能被重新运行并得到相同结果、关键数值在控制台打印
并落盘"。各问主脚本（run_q1/run_q2/run_q3…）在 `main()` 外层套一个 Tee 即可，
避免在每份脚本里重复定义一个日志类。

用法：
    from lib.logio import Tee
    tee = Tee("问题3/主模型运行日志.txt")   # 传入日志文件的绝对路径
    sys.stdout = tee
    try:
        main()
    finally:
        sys.stdout = tee.stdout
        tee.close()
"""

import sys


class Tee:
    """把控制台输出同时写入日志文件，保证"落盘"不依赖人工复制。

    输入：path，str，日志文件路径
    输出：无（作为上下文管理器使用；write/flush 同时作用于控制台与文件）
    """

    def __init__(self, path):
        # 控制台默认可能是 GBK 代码页，先切到 utf-8，避免中文与数学符号报编码错
        self.stdout = sys.stdout
        try:
            self.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
        # 以 utf-8 打开日志文件（python 文件本身不写编码注释，编码在 open 时指定）
        self.file = open(path, "w", encoding="utf-8")

    def write(self, text):
        # 同时写日志与回显，二者内容完全一致
        self.stdout.write(text)
        self.file.write(text)

    def flush(self):
        # 刷新两个通道，避免异常退出时日志残缺
        self.stdout.flush()
        self.file.flush()

    def close(self):
        # 关闭日志文件句柄
        self.file.close()
