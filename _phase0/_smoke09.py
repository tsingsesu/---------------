import os, sys
sys.argv=["x"]
# 冒烟测试：确认验收对账脚本在文件缺失时不崩
os.chdir(r"C:\Users\lhl87\Desktop\高教社杯全国大学生数学建模大赛")
exec(open("_phase0/probe09_verify_deliverable.py", encoding="utf-8").read())
print("--- report ---")
print(open("_phase0/报告14_交付物验收对账.txt", encoding="utf-8").read())
