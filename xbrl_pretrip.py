# -*- coding: utf-8 -*-
r"""
xbrl_pretrip.py — 解析 MOPS 財報頁的 inline-XBRL 標記，產出「勤前包」JSON（供勤前分析／行前報告產線）。

與 fetch_mops.py 的分工：
  fetch_mops.py   → pandas 掃「表格版面」：查單一科目、印全表（人眼核對用）。
  xbrl_pretrip.py → 直接讀 ix: 標記（taxonomy 概念名），一次抽出勤前分析要的全部結構化資料：
                    會計師報告（事務所／簽證會計師／日期／意見型態旗標／報告全文含 KAM）、
                    附註重大交易附表 tuple（資金貸與含利率、背書保證含佔淨值比、關係人交易、
                    集團架構持股%、大陸投資）、四大表全部數字、附註文字塊（期後事項、或有負債），
                    並自動判定紅旗（非無保留意見、無息／低利貸與、背書比率過高、虧損子公司…）。

為什麼讀標記而非版面：MOPS 財報頁（t_ifrs_fr1m1_cicr / t164sb01 step=3）的意見型態、KAM 全文、
附註附表逐列數值都以 XBRL 概念標記（tifrs-ar / tifrs-notes / tifrs-bsci-ci / tifrs-es / tifrs-scf /
ifrs-full），概念名跨公司一致，不受版面差異影響；意見型態的機制是「該型態的概念出現＝該型態成立」
（如 QualifiedOpinionAbstract 出現＝保留）。附註明細一列＝一個 <ix:tuple>（實體巢狀，無 tupleRef），
需按 tuple 分組還原，不能把同概念的值攤平。已用 2371 大同 2025Q4／2026Q1、8933 愛地雅 2025Q4 實檔驗證。

用法：
    # 離線（推薦）：完整勤前包 JSON
    python xbrl_pretrip.py 2371_2025Q4.html -o 2371_2025Q4_pretrip.json
    # 終端摘要（token 精簡：基本資料＋意見＋紅旗）
    python xbrl_pretrip.py 2371_2025Q4.html --summary
    # 只印某一節（audit / tuples / notes_text / statements / red_flags / meta）
    python xbrl_pretrip.py 2371_2025Q4.html --section audit
    # 線上抓（複用 fetch_mops.fetch_online；需能連 MOPS）
    python xbrl_pretrip.py --co 2371 --year 2025 --season 4 --summary

編碼：MOPS 線上回應與另存檔可能是 UTF-8 或 Big5，自動偵測。
依賴：純 Python 標準庫（離線模式零安裝）；線上模式才需要 requests。

⚠ 財報數字為公司依法公開申報之政府公開資訊，可查閱引用；線上抓請節制，勿高頻打站。
"""
import argparse
import html as _html
import json
import re
import sys

# ---------------------------------------------------------------- 讀檔 / 編碼

def read_html(path: str) -> str:
    """自動偵測編碼：先看 <meta charset>，再依序試 utf-8 / big5（含 big5hkscs 後備）。"""
    raw = open(path, "rb").read()
    head = raw[:2000].decode("ascii", "ignore").lower()
    m = re.search(r"charset=([a-z0-9-]+)", head)
    order = []
    if m:
        order.append(m.group(1))
    order += ["utf-8", "big5", "big5hkscs"]
    for enc in order:
        try:
            return raw.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode("utf-8", "replace")


# ---------------------------------------------------------------- ix 標記解析

_ATTR = re.compile(r'([\w:-]+)="([^"]*)"')
# 事件：tuple 開／關、fact（nonNumeric / nonFraction，含自閉合）
_EVENT = re.compile(
    r"<ix:tuple\b[^>]*>"
    r"|</ix:tuple>"
    r"|<ix:(?:nonNumeric|nonFraction)\b[^>]*/>"
    r"|<ix:(nonNumeric|nonFraction)\b[^>]*>(.*?)</ix:\1>",
    re.S | re.I,
)
_TAGS = re.compile(r"<[^>]+>")


def _attrs(tag_text: str) -> dict:
    return dict(_ATTR.findall(tag_text))


def _clean(text: str) -> str:
    """去內部標籤、解 HTML entity、收攏空白（保留換行供長文閱讀）。"""
    t = _TAGS.sub("", text)
    t = _html.unescape(t)
    t = re.sub(r"[ \t　]+", " ", t)
    t = re.sub(r"\n\s*\n+", "\n", t)
    return t.strip()


def _num(raw: str, attrs: dict):
    """nonFraction 顯示值 → float（維持頁面顯示單位：金額＝千元、百分比＝%）。sign='-' 取負。"""
    t = raw.replace(",", "").replace("　", "").strip()
    if t in ("", "-", "－"):
        return None
    neg = attrs.get("sign") == "-"
    if t.startswith("(") and t.endswith(")"):
        t, neg = t[1:-1], True
    try:
        v = float(t)
    except ValueError:
        return None
    return -v if neg else v


def parse_facts(doc: str):
    """單遍掃描全部 ix 事件。回傳 (facts, tuple_stats)。

    tuple 成員帶 (tuple 概念名, tuple 序號)；巢狀子 tuple（如貸與列內的擔保品、背書列內的
    被背書公司）一律歸到最外層 tuple ＝ 併回父列，附表一列的資訊才完整。
    tuple_stats 供偵測「瀏覽器另存檔」：ix:tuple 包在 <tr> 內非合法 HTML，瀏覽器 DOM
    序列化時會被 foster-parenting 搬成空殼（facts 仍留在 <td> 原位），此時巢狀還原失效。
    """
    facts = []
    stack = []          # [(tuple_name, tuple_seq)]
    tuple_seq = 0
    opened = populated = 0
    for m in _EVENT.finditer(doc):
        tag = m.group(0)
        low = tag[:12].lower()
        if low.startswith("<ix:tuple"):
            a = _attrs(tag.split(">", 1)[0])
            if tag.rstrip().endswith("/>"):
                continue
            tuple_seq += 1
            opened += 1
            stack.append([a.get("name", ""), tuple_seq, 0])   # [name, seq, fact_count]
        elif low.startswith("</ix:tuple"):
            if stack:
                if stack[-1][2]:
                    populated += 1
                stack.pop()
        else:
            head = tag.split(">", 1)[0]
            a = _attrs(head)
            name = a.get("name", "")
            if not name:
                continue
            body = m.group(2) or ""
            text = _clean(body)
            fact = {
                "name": name,
                "context": a.get("contextRef", ""),
                "text": text,
            }
            if "nonfraction" in head.lower():
                fact["value"] = _num(text, a)
                if "scale" in a:
                    fact["scale"] = a["scale"]
            if stack:
                fact["tuple"] = (stack[0][0], stack[0][1])    # 歸最外層＝子 tuple 併回父列
                stack[-1][2] += 1
            facts.append(fact)
    return facts, {"opened": opened, "populated": populated}


# ---------------------------------------------------------------- 分節組裝

# 意見型態：概念出現＝型態成立（named per-type in tifrs-ar）。順序＝比對優先序（長名先比）。
_OPINION_MAP = [
    ("UnqualifiedOpinionWithEmphasisOfMatterParagraphsOrOtherMatterParagraphs",
     "修正式無保留（含強調事項或其他事項段）"),
    ("QualifiedOpinionAbstract", "保留意見／結論"),
    ("UnqualifiedOpinionAbstract", "無保留意見／結論"),
    ("AdverseOpinion", "否定意見"),
    ("DisclaimerOfOpinion", "無法表示意見"),
]

# 勤前分析關注的附註文字塊（非 tuple 的 tifrs-notes 敘述）
_NOTE_TEXT_CONCEPTS = [
    "HistoryAndOrganization",
    "SignificantEventsAfterTheEndOfThePeriod",
    "SignificantContingentLiabilitiesAndUnrecognisedContractualCommitments",
    "SourcesOfUncertaintyFromSignificantAccountingJudgmentsAssumptionsAndEstimations",
    "SubsidiariesNotConsolidatedInTheFinancialStatements",
    "PrinciplesOfConsolidation",
    "OtherTransactionWithRelatedParties",
    # ↓ 實審「財務報告審閱說明」四項所需之會計政策原文——有了這些，管區意見的
    #   「認列及衡量」段才寫得出公司特定政策，而非一句「擬行前核閱財報附註確認」
    "Leasing",                          # IFRS16 租賃
    "FinancialInstruments",             # IFRS9 認列衡量＋IFRS7 揭露
    "RevenueRecognition",               # IFRS15 收入
    "InvestmentsInAssociates",          # IAS28／IFRS10 控制與重大影響力判斷
    "ImpairmentOfNonFinancialAsset",    # 減損評估（風險事項常引用）
    "PropertyPlantAndEquipment",        # 占比常最大的資產科目
    "ApplicationOfNewlyIssuedOrAmendedStandardsAndInterpretations",
    "GuaranteesAmongRelatedParties",    # 背書保證
    "ScheduleOfTradeAndOtherReceivables",
]

_META_CONCEPTS = ["CompanyID", "CompanyChineseName", "CompanyEnglishName",
                  "Year", "Quarter", "ReportType", "ReportCategory", "IndustrySector"]

_STMT_PREFIXES = ("ifrs-full:", "tifrs-bsci-ci:", "tifrs-es:", "tifrs-scf:", "tifrs-SCF:")


def _local(name: str) -> str:
    return name.split(":", 1)[-1]


def _by_concept(facts):
    d = {}
    for f in facts:
        d.setdefault(f["name"], []).append(f)
    return d


def _reconstruct_rows(facts):
    """降級重建（瀏覽器另存檔）：tuple 全成空殼時，facts 仍按列序留在 <td> 原位。
    以「概念重複出現＝換列」重切附註 facts。列歸屬的附表名無法還原，統一掛 _reconstructed；
    紅旗偵測按欄位特徵認列，不依賴附表名。"""
    rows, row = [], {}
    skip = set(_META_CONCEPTS) | set(_NOTE_TEXT_CONCEPTS)
    for f in facts:
        if "tuple" in f or not f["name"].startswith("tifrs-notes:"):
            continue
        local = _local(f["name"])
        if local in skip:
            continue
        val = f.get("value", f["text"])
        if local in row:
            rows.append(row)
            row = {}
        row[local] = val
    if row:
        rows.append(row)
    return [r for r in rows if len(r) > 1]


def build_pretrip(doc: str, source: str = "") -> dict:
    facts, tstats = parse_facts(doc)
    by = _by_concept(facts)

    def first_text(prefix, local):
        for f in by.get(prefix + local, []):
            if f["text"]:
                return f["text"]
        return None

    def all_texts(prefix, local):
        seen, out = set(), []
        for f in by.get(prefix + local, []):
            if f["text"] and f["text"] not in seen:
                seen.add(f["text"])
                out.append(f["text"])
        return out

    # --- meta
    meta = {"source": source}
    for c in _META_CONCEPTS:
        v = all_texts("tifrs-notes:", c)
        if v:
            meta[c] = v[0] if len(v) == 1 else v

    # --- audit（tifrs-ar 全收，再整理重點欄）
    ar = {k: v for k, v in by.items() if k.startswith("tifrs-ar:")}
    opinion = None
    for key, label in _OPINION_MAP:
        hit = [k for k in ar if key in k]
        if hit:
            opinion = {"concept": _local(hit[0]), "label": label}
            break
    bodies = all_texts("tifrs-ar:", "AccountantsReportBody")
    audit = {
        "firm": all_texts("tifrs-ar:", "AccountantName"),
        "cpa": (all_texts("tifrs-ar:", "AssuranceAccountantName1")
                + all_texts("tifrs-ar:", "AssuranceAccountantName2")),
        "report_date": first_text("tifrs-ar:", "ReviewAuditDate"),
        "report_kind": ("查核報告" if "tifrs-ar:AuditReport" in ar
                        else "核閱報告" if "tifrs-ar:ReviewReport" in ar else None),
        "opinion": opinion,
        "flags": sorted(_local(k) for k in ar),
        "report_body_zh": bodies[0] if bodies else None,
        "report_body_en": bodies[1] if len(bodies) > 1 else None,
        "emphasis_of_matter": first_text("tifrs-ar:", "DescriptionEmphasizeItemBody"),
        "other_matter": first_text("tifrs-ar:", "DescriptionOtherItemBody"),
    }
    # 未經查核（核閱）之子公司／權益法投資金額（保留結論常見基礎）
    unaudited = {}
    for k, fs in ar.items():
        lk = _local(k)
        if "Unaudited" in lk or "UnauditedOrUnreviewed" in lk or "OnUnauditedOrUnreviewed" in lk:
            vals = [f.get("value") for f in fs if f.get("value") is not None]
            if vals:
                unaudited[lk] = vals[0]
    audit["unaudited_amounts"] = unaudited or None

    # --- tuples（附註重大交易附表：一 tuple＝一列）
    tuples = {}
    rows = {}
    for f in facts:
        if "tuple" not in f:
            continue
        tname, tseq = f["tuple"]
        key = (tname, tseq)
        row = rows.setdefault(key, {})
        local = _local(f["name"])
        val = f.get("value", f["text"])
        if local in row:                      # 同概念多值（如中英兩行）→ 收成 list
            if not isinstance(row[local], list):
                row[local] = [row[local]]
            row[local].append(val)
        else:
            row[local] = val
    for (tname, tseq) in sorted(rows, key=lambda k: k[1]):
        tuples.setdefault(_local(tname), []).append(rows[(tname, tseq)])

    # 瀏覽器另存檔降級：tuple 開了很多但大多是空殼 → 逐列重建
    degraded = tstats["opened"] > 0 and tstats["populated"] < tstats["opened"] * 0.5
    if degraded:
        rec = _reconstruct_rows(facts)
        if rec:
            tuples["_reconstructed"] = rec
        meta["warning"] = (
            f"此檔疑為瀏覽器另存之 DOM 序列化版（ix:tuple {tstats['opened']} 個、"
            f"有內容者僅 {tstats['populated']} 個）：附表列以欄位重複啟發式重建（_reconstructed），"
            "列切分可能不精確；建議改用線上模式或 requests/curl 取得原始 HTML。")

    # --- 附註文字塊
    notes_text = {}
    for c in _NOTE_TEXT_CONCEPTS:
        v = first_text("tifrs-notes:", c)
        if v:
            notes_text[c] = v

    # --- 四大表數字（非 tuple 的報表概念；同概念不同期別以 context 區分）
    statements = {}
    for f in facts:
        if "tuple" in f or "value" not in f or f["value"] is None:
            continue
        if not f["name"].startswith(_STMT_PREFIXES):
            continue
        statements.setdefault(f["name"], {})[f["context"]] = f["value"]

    out = {
        "meta": meta,
        "audit": audit,
        "tuples": tuples,
        "notes_text": notes_text,
        "statements": statements,
    }
    out["red_flags"] = detect_red_flags(out)
    return out


# ---------------------------------------------------------------- 紅旗偵測

def _f(row, *keys):
    """取列中第一個存在的欄位；list 取首值；轉 float 失敗回 None。"""
    for k in keys:
        if k in row:
            v = row[k]
            if isinstance(v, list):
                v = v[0]
            if isinstance(v, (int, float)):
                return v
            try:
                return float(str(v).replace(",", "").replace("%", ""))
            except (ValueError, TypeError):
                return None
    return None


def _s(row, *keys):
    for k in keys:
        if k in row:
            v = row[k]
            return v[0] if isinstance(v, list) else v
    return None


def detect_red_flags(pkg: dict) -> list:
    flags = []
    audit = pkg["audit"]

    # 1. 意見型態非無保留
    op = audit.get("opinion")
    if op and "無保留" not in op["label"]:
        flags.append({"type": "意見非無保留", "detail": op["label"],
                      "hint": "讀 audit.report_body_zh 的『基礎』段確認原因"})
    elif op and "修正式" in op["label"]:
        flags.append({"type": "修正式無保留", "detail": op["label"],
                      "hint": "有強調事項或其他事項段，讀 emphasis_of_matter / other_matter"})
    for k in audit.get("flags", []):
        if "Uncertainty" in k:
            flags.append({"type": "不確定性事項", "detail": k,
                          "hint": "查核報告揭露重大不確定性（訴訟／監理／繼續經營）"})
    ua = audit.get("unaudited_amounts") or {}
    for k, v in ua.items():
        if v:
            flags.append({"type": "部分個體未經查核（核閱）", "detail": f"{k}={v:,.0f}",
                          "hint": "常為保留結論基礎；核對占合併資產／損益比重"})

    # 2~5. 附表逐列：按欄位特徵認列（不依附表名，降級重建之列同樣適用）
    all_rows = [r for rows_ in pkg["tuples"].values() for r in rows_]
    for r in all_rows:
        # 資金貸與列：有利率或動支欄
        if "RangeOfInterestRates" in r or "LoansToOthers" in r:
            rate = _f(r, "RangeOfInterestRates")
            amt = _f(r, "ActualAmountProvided", "EndingBalance1", "MaximumBalanceForThePeriod1")
            who = (f'{_s(r, "Lender1", "CompanyName") or "?"}→'
                   f'{_s(r, "Counterparty1", "Counterparty", "NameOfTheCompany") or "?"}')
            if amt and (rate is None or rate == 0):
                flags.append({"type": "資金貸與未計息", "detail": f"{who} 動支 {amt:,.0f} 千元",
                              "hint": "未計息或利率欄空白，確認決策程序與資貸背書準則規範"})
            elif amt and rate is not None and 0 < rate < 1:
                flags.append({"type": "資金貸與利率偏低", "detail": f"{who} 利率 {rate}%、動支 {amt:,.0f} 千元"})
            allow = _f(r, "AllowanceForBadDebts")
            if amt and allow and allow >= amt * 0.5:
                flags.append({"type": "資金貸與已提列高額備抵呆帳",
                              "detail": f"{who} 動支 {amt:,.0f}、備抵 {allow:,.0f} 千元",
                              "hint": "回收性存疑，確認貸與決策與後續評估"})
        # 背書保證列：佔淨值比偏高
        ratio = _f(r, "RatioOfAccumulatedEndorsementGuaranteeAmountToNetAssetOfTheCompanyPerLatestFinancialStatements")
        if ratio is not None and ratio >= 50:
            who = _s(r, "CompanyNameOfTheEndorserGuarantor") or "?"
            flags.append({"type": "背書保證佔淨值比偏高", "detail": f"{who} {ratio}%",
                          "hint": "核對公司背書保證作業程序之限額與決策程序"})
        # 轉投資列：被投資公司虧損
        ni = _f(r, "NetIncomeLossesOfTheInvestee")
        if ni is not None and ni < 0:
            who = _s(r, "NameOfInvestee", "CompanyNameOfTheInvestee",
                     "CompanyNameOfTheInvesteeInMainlandChina") or "?"
            flags.append({"type": "被投資公司虧損", "detail": f"{who} 本期損益 {ni:,.0f} 千元"})

    # 去重（同一公司同時列在合併個體表與權益法投資表等）
    seen, out = set(), []
    for fl in flags:
        key = (fl["type"], fl["detail"])
        if key not in seen:
            seen.add(key)
            out.append(fl)
    return out


# ---------------------------------------------------------------- 摘要輸出

def print_summary(pkg: dict):
    m, a = pkg["meta"], pkg["audit"]
    if m.get("warning"):
        print(f"⚠ {m['warning']}\n")
    print(f"公司：{m.get('CompanyChineseName', '?')}（{m.get('CompanyID', '?')}）"
          f"　期別：{m.get('Year', '?')} Q{m.get('Quarter', '?')}　{m.get('ReportCategory', '')}")
    firm = "、".join(a["firm"][:1]) if a["firm"] else "?"
    cpa = "、".join(x for x in a["cpa"] if re.search(r"[一-鿿]", x)) or "、".join(a["cpa"])
    print(f"{a.get('report_kind') or '會計師報告'}：{firm}　簽證會計師：{cpa}　日期：{a.get('report_date')}")
    op = a.get("opinion")
    print(f"意見型態：{op['label'] if op else '（未偵測到旗標，讀 report_body 判斷）'}")
    body = a.get("report_body_zh") or ""
    if "關鍵查核事項" in body:
        kams = re.findall(r"[一二三四五六七八九十]、([^\n]{2,40})", body.split("關鍵查核事項", 1)[1])
        print(f"KAM：{'；'.join(kams) if kams else '（有 KAM 段，讀 report_body_zh 摘要）'}")
    print(f"附表 tuple：{ {k: len(v) for k, v in pkg['tuples'].items()} }")
    print(f"\n紅旗（{len(pkg['red_flags'])}）：")
    for fl in pkg["red_flags"]:
        print(f"  ⚠ [{fl['type']}] {fl['detail']}" + (f"　→ {fl['hint']}" if fl.get("hint") else ""))
    if not pkg["red_flags"]:
        print("  （無自動偵測紅旗；質性研判仍需讀報告全文與附註）")


# ---------------------------------------------------------------- 主流程

def main():
    ap = argparse.ArgumentParser(description="MOPS inline-XBRL 勤前包擷取")
    ap.add_argument("html", nargs="?", help="本地 MOPS 財報 HTML 檔")
    ap.add_argument("--co", help="公司代號（線上抓）")
    ap.add_argument("--year", type=int, help="年度（西元或民國）")
    ap.add_argument("--season", type=int, choices=[1, 2, 3, 4])
    ap.add_argument("--typek", default="sii", choices=["sii", "otc", "rotc", "pub"])
    ap.add_argument("-o", "--out", help="輸出 JSON 路徑")
    ap.add_argument("--summary", action="store_true", help="終端印精簡摘要（含紅旗）")
    ap.add_argument("--section", choices=["meta", "audit", "tuples", "notes_text",
                                          "statements", "red_flags"],
                    help="只印指定節之 JSON")
    a = ap.parse_args()

    if a.html:
        doc = read_html(a.html)
        src = a.html
    elif a.co and a.year and a.season:
        from fetch_mops import fetch_online
        doc = fetch_online(a.co, a.year, a.season, a.typek, "t_ifrs_fr1m1_cicr")
        src = f"MOPS online {a.co} {a.year} S{a.season}"
    else:
        ap.error("請給本地 HTML 檔，或同時給 --co/--year/--season 線上抓")

    pkg = build_pretrip(doc, source=src)

    if a.out:
        with open(a.out, "w", encoding="utf-8") as f:
            json.dump(pkg, f, ensure_ascii=False, indent=1)
        print(f"已輸出：{a.out}（facts：statements {len(pkg['statements'])} 概念、"
              f"tuples {sum(len(v) for v in pkg['tuples'].values())} 列、紅旗 {len(pkg['red_flags'])}）")
    if a.section:
        json.dump(pkg[a.section], sys.stdout, ensure_ascii=False, indent=1)
        print()
    if a.summary or (not a.out and not a.section):
        print_summary(pkg)


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
