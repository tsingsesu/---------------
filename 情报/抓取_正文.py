"""情报手：网页正文抓取与关键句摘录工具。

用法
----
1) 把要抓的 URL 写进 情报/urls.json（形如 {"name": "标题", "url": "https://...", "note": "用途"} 的列表）；
2) 运行 `python 情报/抓取_正文.py`；
3) 每个 URL 的**清洗后正文**落到 情报/原始摘录/<序号>_<安全名>.txt，
   摘要（前 1200 字 + 命中关键词的句子）同时打印到终端，便于人工核对。

设计要点
--------
- 本机访问境外站点必须显式走代理 http://127.0.0.1:7890（AGENTS.md §4.1）；
  境外站点失败时自动回退一次无代理重试（部分国内站点经代理反而更慢/被拒）。
- 抓取带浏览器 User-Agent；记录 HTTP 状态、抓取时间、最终 URL（含跳转后地址）。
- 只落盘正文，不落盘 HTML；关键句由 KEYWORDS 命中筛选，避免报告里粘贴整页。
"""
from __future__ import annotations

import json
import re
import sys
import time
from pathlib import Path

import requests
from bs4 import BeautifulSoup

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "原始摘录"
PROXY = {"http": "http://127.0.0.1:7890", "https": "http://127.0.0.1:7890"}
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"

# 用于从正文里挑"可能有用"的句子，覆盖四个待答问题
KEYWORDS = [
    "峰谷", "价差", "3:1", "4:1", "尖峰", "分时电价",
    "效率", "往返", "循环", "倍率", "C-rate", "放电深度", "SOC",
    "MILP", "混合整数", "线性规划", "随机", "鲁棒", "滚动", "模型预测", "MPC",
    "套利", "收益", "元/kWh", "回收期", "度电",
    "round-trip", "efficiency", "depth of discharge", "cycle life", "C-rate",
    "MILP", "stochastic", "robust", "rolling horizon", "model predictive",
    "arbitrage", "peak-valley", "time-of-use", "penalty", "deviation",
]


def clean_html(html: str) -> tuple[str, str]:
    soup = BeautifulSoup(html, "lxml")
    for tag in soup(["script", "style", "noscript", "nav", "footer", "header", "form", "svg"]):
        tag.decompose()
    title = (soup.title.get_text(strip=True) if soup.title else "") or ""
    main = soup.find("article") or soup.find("main") or soup.body or soup
    text = main.get_text("\n", strip=True)
    text = re.sub(r"\n{2,}", "\n", text)
    return title, text


def pick_sentences(text: str, limit: int = 40) -> list[str]:
    sents = re.split(r"(?<=[。；！？!?;])\s*|\n", text)
    hits: list[str] = []
    for s in sents:
        s = s.strip()
        if not (12 <= len(s) <= 400):
            continue
        if any(k in s for k in KEYWORDS):
            hits.append(s)
        if len(hits) >= limit:
            break
    return hits


def safe(name: str, n: int) -> str:
    name = re.sub(r"[^\w\u4e00-\u9fff-]+", "_", name)[:60] or "page"
    return f"{n:02d}_{name}"


def fetch(item: dict, idx: int) -> dict:
    url = item["url"]
    rec = {"index": idx, "name": item.get("name", ""), "url": url, "note": item.get("note", "")}
    last_err = ""
    for use_proxy in (True, False):
        try:
            r = requests.get(
                url,
                proxies=PROXY if use_proxy else None,
                timeout=40,
                headers={"User-Agent": UA, "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8"},
            )
            rec.update(
                status=r.status_code,
                final_url=r.url,
                via_proxy=use_proxy,
                encoding=r.encoding,
                fetched_at=time.strftime("%Y-%m-%d %H:%M:%S"),
                bytes=len(r.content),
            )
            ctype = r.headers.get("Content-Type", "")
            if "pdf" in ctype.lower() or url.lower().endswith(".pdf"):
                rec["kind"] = "pdf"
                rec["text"] = f"[PDF 二进制，{len(r.content)} bytes；请用 read_document 阅读]"
                body = rec["text"]
            else:
                r.encoding = r.apparent_encoding or r.encoding
                title, body = clean_html(r.text)
                rec["kind"] = "html"
                rec["page_title"] = title
                rec["text"] = body
            if len(body) < 200:
                last_err = f"正文过短({len(body)} 字符)"
                continue
            break
        except Exception as exc:  # noqa: BLE001
            last_err = f"{type(exc).__name__}: {exc}"
            rec["error"] = last_err
            continue
    rec.setdefault("error", last_err)
    return rec


def main() -> int:
    urls_file = ROOT / (sys.argv[1] if len(sys.argv) > 1 else "urls.json")
    if not urls_file.exists():
        print(f"未找到 {urls_file}，请先写入待抓清单。", file=sys.stderr)
        return 2
    items = json.loads(urls_file.read_text(encoding="utf-8"))
    OUT.mkdir(exist_ok=True)
    index = []
    for i, item in enumerate(items, 1):
        print(f"\n=== [{i}] {item.get('name','')} <{item['url']}>")
        rec = fetch(item, i)
        status = rec.get("status", "-")
        print(f"    HTTP {status}  proxy={rec.get('via_proxy')}  bytes={rec.get('bytes','-')}  err={rec.get('error','') or '-'}")
        body = rec.get("text", "")
        stem = safe(rec["name"] or rec.get("page_title", ""), i)
        fp = OUT / f"{stem}.txt"
        head = (
            f"# {rec['name']}\n"
            f"URL: {rec['url']}\n最终URL: {rec.get('final_url','')}\n"
            f"访问时间: {rec.get('fetched_at','')}\nHTTP: {status}  代理: {rec.get('via_proxy')}  字节: {rec.get('bytes','-')}\n"
            f"抓取方式: 正文抓取（requests+bs4+lxml）\n用途: {rec['note']}\n"
            f"{'-'*70}\n"
        )
        fp.write_text(head + body, encoding="utf-8")
        rec["saved_to"] = str(fp.relative_to(ROOT))
        hits = pick_sentences(body)
        rec["key_sentences"] = hits
        for s in hits[:18]:
            print("    · " + s[:220])
        index.append({k: v for k, v in rec.items() if k != "text"})
    (ROOT / "抓取清单.json").write_text(
        json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"\n共 {len(index)} 条，清单与摘录已写入 {ROOT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
