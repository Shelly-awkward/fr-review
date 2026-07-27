# -*- coding: utf-8 -*-
r"""
fetch_company_list.py — 從 MOPS 抓「公開發行公司」名單 → data/companies.json。

    python scripts/fetch_company_list.py [--typek pub] [--out data/companies.json]

typek：pub＝公開發行（未上市櫃，本工具的主要對象）、sii＝上市、otc＝上櫃、rotc＝興櫃。
名單一年變動不大，年度歸檔前跑一次即可。
"""
import argparse
import json
import re
import sys

_BASE = "https://mopsov.twse.com.tw"
_HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"),
    "Referer": _BASE + "/mops/web/t51sb01",
}


def decode(content: bytes) -> str:
    for enc in ("utf-8", "big5", "big5hkscs"):
        try:
            return content.decode(enc)
        except UnicodeDecodeError:
            continue
    return content.decode("utf-8", "replace")


def parse(html: str) -> list:
    """表格每列前四欄＝公司代號／公司名稱／簡稱／產業別。"""
    out, seen = [], set()
    for row in re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S):
        cells = [re.sub(r"<[^>]+>", "", c).replace("&nbsp;", " ").strip()
                 for c in re.findall(r"<td[^>]*>(.*?)</td>", row, re.S)]
        if len(cells) >= 4 and re.fullmatch(r"\d{4,6}", cells[0]) and cells[0] not in seen:
            seen.add(cells[0])
            out.append({"co": cells[0], "name": cells[1],
                        "abbr": cells[2], "industry": cells[3]})
    return sorted(out, key=lambda x: x["co"])


def main():
    import requests
    ap = argparse.ArgumentParser(description="抓 MOPS 公司名單")
    ap.add_argument("--typek", default="pub", choices=["pub", "sii", "otc", "rotc"])
    ap.add_argument("--out", default="data/companies.json")
    a = ap.parse_args()

    r = requests.post(_BASE + "/mops/web/ajax_t51sb01", headers=_HEADERS, timeout=60,
                      data={"encodeURIComponent": 1, "step": 1, "firstin": 1,
                            "TYPEK": a.typek, "code": ""})
    r.raise_for_status()
    cos = parse(decode(r.content))
    if not cos:
        print("解析不到任何公司——MOPS 版面可能改版，請檢查回應內容")
        sys.exit(1)

    with open(a.out, "w", encoding="utf-8") as f:
        json.dump({"typek": a.typek, "companies": cos}, f, ensure_ascii=False, indent=1)
    print(f"✔ {a.out}（{a.typek} 共 {len(cos)} 家）")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
