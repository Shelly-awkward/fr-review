# -*- coding: utf-8 -*-
r"""
soften_content.py — B 版（fr-review2.html）用的「不擋關」驗收：修補＋標紅，而不是退件。

    python scripts/soften_content.py out/8304_114_checklist_content.json \
        --review out/8304_114_review_content.json --out out/8304_114_checklist_soft.json

與 check_content.py 的差別：check_content 不過就不給 Word；本檔一律產出可用的內容 JSON，
把驗收不過的地方就地加註「【未經驗證…】」，gen_checklist_docx 會把含此字樣的句子印成紅字，
並在文件最前面列出未經驗證清單（content["warnings"]）。

  1. 結構缺漏（cover 欄位、檢查表項目、五段＋資料來源）→ 用 draft 骨架補回
  2. 殘留佔位字 → 就地加註；【AI待填：…】改寫成「【未經驗證：AI 未填，請自行撰寫】」
  3. 金額在 facts 找不到 → 數字後加「（未經驗證）」
  4. 未提供同業平均卻出現數值 → 該值後加註
  5. 檢查表未勾選／勾異常卻無事由 → 備註加註
輸入的內容 JSON 可以是空物件 {}（＝完全不用 AI，直接以骨架產初稿）。
"""
import argparse
import json
import re
import sys

from check_content import PLACEHOLDERS, REQUIRED_SECTIONS, walk_text

MARK = "【未經驗證】"
AI_TODO = re.compile(r"【AI待填[：:][^】]*】")


def number_pool(facts):
    pool = set()

    def collect(o):
        if isinstance(o, (int, float)) and not isinstance(o, bool) and o == o:
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
    return pool


def derivable(v, pool):
    """與 check_content 相同：pool 內任兩數之和或差視為有出處。"""
    return any((v + a) in pool or (a - v) in pool or (v - a) in pool for a in pool)


def map_strings(o, fn):
    """遞迴把所有字串套 fn，回新物件。"""
    if isinstance(o, str):
        return fn(o)
    if isinstance(o, dict):
        return {k: map_strings(v, fn) for k, v in o.items()}
    if isinstance(o, list):
        return [map_strings(v, fn) for v in o]
    return o


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("content")
    ap.add_argument("--review", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()

    with open(a.content, encoding="utf-8") as f:
        c = json.load(f)
    if not isinstance(c, dict):
        c = {}
    with open(a.review, encoding="utf-8") as f:
        review = json.load(f)
    draft, facts = review["draft"], review.get("facts", {})
    warns = []

    # 1. 結構：缺的用骨架補
    for key in ("title", "cover", "groups", "footnotes", "sections"):
        if not c.get(key):
            c[key] = draft.get(key)
            if key != "footnotes":
                warns.append(f"AI 回覆缺「{key}」，已改用空白骨架")

    cover = c["cover"] if isinstance(c["cover"], dict) else {}
    for k, v in draft["cover"].items():
        if not str(cover.get(k) or "").strip():
            cover[k] = v
            if AI_TODO.search(str(v)):
                warns.append(f"複核表「{k}」AI 未填")
    c["cover"] = cover

    ai_items = {it.get("id"): it for g in c["groups"] for it in g.get("items", [])}
    groups, unmarked, no_reason = [], [], []
    for g in draft["groups"]:
        items = []
        for d_it in g["items"]:
            it = dict(ai_items.get(d_it["id"]) or d_it)
            it["text"] = d_it["text"]        # 檢查內容是法定文字，一律以骨架為準
            note = (it.get("note") or "").strip()
            if it.get("mark") not in ("yes", "no"):
                it["mark"] = None
                it["note"] = f"{MARK}AI 未判斷，請自行勾選。" + note
                unmarked.append(d_it["id"])
            elif it["mark"] == "no" and not note:
                it["note"] = f"{MARK}勾異常但未寫事由，請自行補充。"
                no_reason.append(d_it["id"])
            items.append(it)
        groups.append({**g, "items": items})
    c["groups"] = groups
    if unmarked:
        warns.append("檢查表未勾選：" + "、".join(unmarked))
    if no_reason:
        warns.append("檢查表勾異常但無事由：" + "、".join(no_reason))

    titles = [s.get("title", "") for s in c["sections"]]
    for req in REQUIRED_SECTIONS:
        if not any(req in t for t in titles):
            sec = next((s for s in draft["sections"] if req in s.get("title", "")), None)
            if sec:
                c["sections"].append(sec)
                warns.append(f"AI 回覆缺「{sec['title']}」段，已改用空白骨架")

    # 2. 佔位字：就地加註
    left = []

    def fix_placeholder(t):
        t, n = AI_TODO.subn("【未經驗證：AI 未填，請自行撰寫】", t)
        hit = n > 0
        for p in PLACEHOLDERS:
            if p != "【AI待填" and p in t:
                t = t.replace(p, p + MARK)
                hit = True
        if hit:
            left.append(t)
        return t
    for key in ("title", "cover", "groups", "footnotes", "sections"):
        c[key] = map_strings(c[key], fix_placeholder)
    if left:
        warns.append(f"有 {len(left)} 處 AI 未填完或含未完成字樣")

    # 3. 金額抽核（同 check_content：sections 內的千分位千元；另加 cover）
    pool = number_pool(facts)
    # 骨架本身的金額是數字層程式算的（如多科目合計），不是 AI 編的，一併視為有出處
    for t in walk_text(draft):
        pool.update(int(m.replace(",", "")) for m in re.findall(r"(\d{1,3}(?:,\d{3})+)千元", t))
    bad = []

    def fix_number(t):
        def sub(m):
            v = int(m.group(1).replace(",", ""))
            if v >= 1000 and v not in pool and not derivable(v, pool):
                bad.append(m.group(1))
                return m.group(0) + "（未經驗證）"
            return m.group(0)
        return re.sub(r"(\d{1,3}(?:,\d{3})+)千元(?!（未經驗證）)", sub, t)
    c["sections"] = map_strings(c["sections"], fix_number)
    c["cover"] = map_strings(c["cover"], fix_number)
    if bad:
        warns.append("金額在財報數字找不到出處：" + "、".join(dict.fromkeys(bad)))

    # 4. 未提供同業平均卻出現數值
    if isinstance(facts.get("peer_avg"), str):
        fake = []
        for s in c["sections"]:
            for b in s.get("body", []):
                paras = b.get("paras", [])
                cols = []
                for i, p in enumerate(paras):
                    if not (isinstance(p, str) and p.startswith("【表】")):
                        cols = []
                        continue
                    cells = p[3:].split("｜")
                    if not cols:
                        cols = [j for j, h in enumerate(cells) if "同業平均" in h] or [-1]
                        continue
                    for j in cols:
                        if 0 <= j < len(cells) and re.search(r"\d", cells[j]) \
                                and "未經驗證" not in cells[j]:
                            fake.append(cells[j].strip())
                            cells[j] = cells[j] + "（未經驗證，請自行查填）"
                    paras[i] = "【表】" + "｜".join(cells)

        def fix_peer(t):
            def sub(m):
                fake.append(m.group(1))
                return m.group(0) + "（未經驗證，請自行查填）"
            return re.sub(r"同業平均[為是]?\s*(-?\d+(?:\.\d+)?\s*[%次])(?!（未經驗證)", sub, t)
        c["sections"] = map_strings(c["sections"], fix_peer)
        if fake:
            warns.append("未提供同業平均，AI 卻填了數值：" + "、".join(dict.fromkeys(fake)))

    c["warnings"] = warns
    with open(a.out, "w", encoding="utf-8") as f:
        json.dump(c, f, ensure_ascii=False, indent=2)

    if warns:
        print(f"⚠ 已產出，但有 {len(warns)} 類未經驗證（Word 內以紅字標示）：")
        for w in warns:
            print("  -", w)
    else:
        print("✔ 驗收通過，無未經驗證項目")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
