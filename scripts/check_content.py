# -*- coding: utf-8 -*-
r"""
check_content.py — 管區意見內容 JSON 的驗收閘門（AI 填完質性段落後、產 docx 前必跑）。

    python scripts/check_content.py out/8304_114_checklist_content.json
    python scripts/check_content.py out/8304_114_checklist_content.json --review out/8304_114_review_content.json

檢查：
  1. 結構完整：title／groups（18 項固定檢查表，每項 mark 已勾）／sections 含五段＋資料來源與限制
  2. 無佔位字：【AI待填、（候選風險）、待補、TODO、XXX、○○
  3. 給 --review 時，抽核內容 JSON 出現的金額數字是否存在於數字層 facts（防 AI 自編數字）
退出碼：0＝通過；1＝未過（逐條列出）。
"""
import argparse
import json
import re
import sys

PLACEHOLDERS = ["【AI待填", "（候選風險）", "待補", "TODO", "XXX", "○○"]
REQUIRED_SECTIONS = ["一、", "二、", "三、", "四、", "五、", "資料來源與限制"]


def walk_text(o):
    if isinstance(o, str):
        yield o
    elif isinstance(o, dict):
        for v in o.values():
            yield from walk_text(v)
    elif isinstance(o, list):
        for v in o:
            yield from walk_text(v)


def tables(sections):
    """把 sections 裡連續的「【表】a｜b｜c」段落還原成表格（list of rows）。"""
    for s in sections:
        for b in s.get("body", []):
            rows = []
            for p in b.get("paras", []):
                if isinstance(p, str) and p.startswith("【表】"):
                    rows.append(p[3:].split("｜"))
                elif rows:
                    yield rows
                    rows = []
            if rows:
                yield rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("content")
    ap.add_argument("--review", default="", help="review_content.json（抽核數字用，建議提供）")
    a = ap.parse_args()

    with open(a.content, encoding="utf-8") as f:
        c = json.load(f)
    errs = []

    # 1. 結構
    for key in ("title", "groups", "sections"):
        if not c.get(key):
            errs.append(f"缺 {key}")
    items = [it for g in c.get("groups", []) for it in g.get("items", [])]
    if len(items) < 17:
        errs.append(f"檢查表項目僅 {len(items)} 項（模板應為 17 項以上）")
    for it in items:
        if it.get("mark") not in ("yes", "no"):
            errs.append(f"檢查表 {it.get('id')} 未勾選（mark 須為 yes/no）")
        if it.get("mark") == "no" and not (it.get("note") or "").strip():
            errs.append(f"檢查表 {it.get('id')} 勾「否(異常)」但備註空白")
    titles = "".join(s.get("title", "") for s in c.get("sections", []))
    for req in REQUIRED_SECTIONS:
        if req not in titles:
            errs.append(f"sections 缺「{req}」段")

    # 2. 佔位字
    for t in walk_text(c):
        for p in PLACEHOLDERS:
            if p in t:
                errs.append(f"殘留佔位字「{p}」：{t[:60]}…")
                break

    # 3. 數字抽核（內容 JSON 的千分位金額須能在 facts 裡找到）
    if a.review:
        with open(a.review, encoding="utf-8") as f:
            facts = json.load(f).get("facts", {})
        pool = set()
        def collect(o):
            if isinstance(o, (int, float)) and o == o:
                v = abs(o)
                pool.add(round(v, 2))
                if v == int(v):
                    pool.add(int(v))
            elif isinstance(o, dict):
                for x in o.values():
                    collect(x)
            elif isinstance(o, list):
                for x in o:
                    collect(x)
        collect(facts)
        body = "\n".join(walk_text(c.get("sections", [])))
        def derivable(v):
            """v 是否為 pool 內任兩數之和或差（毛利減少額、貸與合計等一階衍生數）。"""
            for a in pool:
                if (v + a) in pool or (a - v) in pool or (v - a) in pool:
                    return True
            return False

        unknown = set()
        for m in re.finditer(r"(\d{1,3}(?:,\d{3})+)千元", body):
            v = int(m.group(1).replace(",", ""))
            if v not in pool and v >= 1000 and not derivable(v):
                unknown.add(m.group(1))
        if unknown:
            errs.append("下列金額在數字層 facts 找不到，請確認出處（衍生計算請於段內寫明算式，"
                        "查無出處＝不得引用）：" + "、".join(sorted(unknown)[:15]))

        # 同業平均是最容易被 AI 憑記憶編造的數字（且無法由財報推得）。
        # facts.peer_avg 為字串＝使用者未提供，此時內文不得出現任何同業平均數值。
        if isinstance(facts.get("peer_avg"), str):
            fake = re.findall(r"同業平均[為是]?\s*[｜|]?\s*(-?\d+(?:\.\d+)?\s*[%次])", body)
            for tbl in tables(c.get("sections", [])):
                cols = [i for i, h in enumerate(tbl[0]) if "同業平均" in h]
                for row in tbl[1:]:
                    for i in cols:
                        if i < len(row) and re.search(r"\d", row[i]):
                            fake.append(row[i].strip())
            if fake:
                errs.append("未提供同業平均資料，內文卻出現同業平均數值："
                            + "、".join(list(dict.fromkeys(fake))[:10])
                            + "——同業平均不可推估或憑記憶填入，請改標「行前請至公開資訊"
                              "觀測站財務業務資訊查填」。")

    if errs:
        print("✘ 驗收未過：")
        for e in errs:
            print("  -", e)
        sys.exit(1)
    print(f"✔ 驗收通過（檢查表 {len(items)} 項、sections {len(c.get('sections', []))} 段）")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
