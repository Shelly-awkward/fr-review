# -*- coding: utf-8 -*-
r"""
build_review_content.py — 財報實審數字層：pretrip JSON → 兩份中介 JSON。

    python scripts/build_review_content.py --co 8304 --year 114 [--data-dir data]
        [--outdir out] [--peer-avg peers.json]

輸出（--outdir，預設 out/）：
  <co>_<民國年>_inquiry.json         → gen_inquiry_xlsx.py 產「財務報告說明」查詢函 Excel
  <co>_<民國年>_review_content.json  → AI 依 REVIEW_PROMPT.md 填質性段落後，
                                        交 gen_checklist_docx.js 產管區意見 Word

分工鐵律：本腳本只算數字（變動%、週轉率、限額核對……全部寫死在輸出裡）；
質性判斷一律留「【AI待填：…指引…】」佔位，由 AI 填、由 check_content.py 把關。
不得引用資料層沒有的數字；同業平均拿不到就留白標註，寧缺勿假。

--peer-avg 格式（使用者自公開資訊觀測站「財務業務資訊」查得後提供）：
  {"上市櫃同業平均": {"營收成長率": -0.19, "銷貨毛利率": 26.99, ...},
   "所有同業平均":   {"營收成長率": -4.29, ...}}
"""
import argparse
import glob
import json
import os
import re
import sys

# ---------- 期別與檔案 ----------

def to_ad(year: int) -> int:
    return year + 1911 if year < 1000 else year


def to_roc(year: int) -> int:
    return year - 1911 if year >= 1000 else year


def load_pretrips(data_dir: str, co: str, ad_year: int, span: int = 6) -> dict:
    """掃描 data/<co>_<西元年>Q4_pretrip.json，回 {西元年: pretrip}（近 span 年內有檔者）。"""
    out = {}
    for p in glob.glob(os.path.join(data_dir, f"{co}_*Q4_pretrip.json")):
        m = re.search(rf"{co}_(\d{{4}})Q4_pretrip\.json$", p.replace("\\", "/"))
        if m and ad_year - span < int(m.group(1)) <= ad_year:
            with open(p, encoding="utf-8") as f:
                out[int(m.group(1))] = json.load(f)
    return out


# ---------- 取數 ----------

def flow(st: dict, keys, y: int):
    """損益／現流科目：取 From<y>0101To<y>1231；keys 依序 fallback。"""
    for k in keys if isinstance(keys, list) else [keys]:
        v = st.get(k, {}).get(f"From{y}0101To{y}1231")
        if v is not None:
            return v
    return None


def stock(st: dict, keys, y: int):
    """資產負債科目：取 AsOf<y>1231。"""
    for k in keys if isinstance(keys, list) else [keys]:
        v = st.get(k, {}).get(f"AsOf{y}1231")
        if v is not None:
            return v
    return None


METRICS_FLOW = {
    "營業收入": ["ifrs-full:Revenue"],
    "營業成本": ["tifrs-bsci-ci:OperatingCosts", "ifrs-full:CostOfSales"],
    "營業毛利": ["tifrs-bsci-ci:GrossProfitLossFromOperations", "ifrs-full:GrossProfit"],
    "營業損益": ["ifrs-full:ProfitLossFromOperatingActivities"],
    "稅前損益": ["ifrs-full:ProfitLossBeforeTax"],
    "本期淨利": ["ifrs-full:ProfitLoss"],
    "其他綜合損益": ["ifrs-full:OtherComprehensiveIncome"],
    "綜合損益": ["ifrs-full:ComprehensiveIncome"],
    "每股盈餘(元)": ["ifrs-full:BasicEarningsLossPerShare"],
    "營業活動現金流量": ["ifrs-full:CashFlowsFromUsedInOperatingActivities"],
    "減損損失": ["ifrs-full:ImpairmentLossRecognisedInProfitOrLoss"],
    "採用權益法之投資損益份額": ["ifrs-full:ShareOfProfitLossOfAssociatesAndJointVentures"
                        "AccountedForUsingEquityMethod"],
}
METRICS_STOCK = {
    "資產總額": ["ifrs-full:Assets"],
    "負債總額": ["ifrs-full:Liabilities"],
    "淨值": ["ifrs-full:Equity"],
    "應收票據": ["tifrs-bsci-ci:NotesReceivableNet"],
    "應收帳款": ["tifrs-bsci-ci:AccountsReceivableNet"],
    "應收帳款-關係人": ["tifrs-bsci-ci:AccountsReceivableDuefromRelatedPartiesNet"],
    "其他應收款": ["ifrs-full:OtherReceivables", "tifrs-bsci-ci:OtherReceivablesNet"],
    "其他應收款-關係人": ["tifrs-bsci-ci:OtherReceivablesDueFromRelatedParties",
                     "ifrs-full:OtherReceivablesDueFromRelatedParties"],
    "預付款項": ["ifrs-full:CurrentPrepayments", "tifrs-bsci-ci:PrepaymentsNet"],
    "存貨": ["ifrs-full:Inventories"],
    "採用權益法之投資": ["ifrs-full:InvestmentAccountedForUsingEquityMethod"],
    "不動產廠房及設備": ["ifrs-full:PropertyPlantAndEquipment"],
    "使用權資產": ["ifrs-full:RightofuseAssets"],
    "租賃負債-流動": ["tifrs-bsci-ci:CurrentLeaseLiabilities"],
    "租賃負債-非流動": ["ifrs-full:NoncurrentFinanceLeaseLiabilities",
                    "tifrs-bsci-ci:NoncurrentLeaseLiabilities"],
    "合約負債-流動": ["ifrs-full:CurrentContractLiabilities"],
    "應付帳款及票據": ["ifrs-full:TradeAndOtherCurrentPayablesToTradeSuppliers"],
    "現金及約當現金": ["ifrs-full:CashAndCashEquivalents"],
    "流動資產": ["ifrs-full:CurrentAssets"],
    "流動負債": ["ifrs-full:CurrentLiabilities"],
    "普通股股本": ["tifrs-bsci-ci:OrdinaryShare", "ifrs-full:IssuedCapital"],
}


def year_values(st: dict, y: int) -> dict:
    out = {k: flow(st, keys, y) for k, keys in METRICS_FLOW.items()}
    out.update({k: stock(st, keys, y) for k, keys in METRICS_STOCK.items()})
    return out


def merge_series(pretrips: dict) -> dict:
    """各 pretrip 供兩個年度的數；新檔優先（同年度以較新申報為準）。回 {西元年: {科目: 值}}。"""
    series = {}
    for fy in sorted(pretrips):  # 由舊到新，新檔覆蓋舊檔同年數字
        st = pretrips[fy]["statements"]
        for y in (fy - 1, fy):
            vals = year_values(st, y)
            if any(v is not None for v in vals.values()):
                cur = series.setdefault(y, {})
                for k, v in vals.items():
                    if v is not None:
                        cur[k] = v
    return series


# ---------- 計算 ----------

def pct(cur, prev):
    """變動%（分母取絕對值；基期 0 或 None 回 None）。"""
    if cur is None or prev in (None, 0):
        return None
    return round((cur - prev) / abs(prev) * 100, 2)


def fmt(v, unit="千元"):
    if v is None:
        return "－"
    if isinstance(v, float) and v == int(v):
        v = int(v)
    return f"{v:,}{unit}" if unit else f"{v:,}"


def fmt_pct(v):
    return "－" if v is None else f"{v:+.2f}%"


def receivables(vals: dict):
    """應收款項＝應收票據＋應收帳款＋應收帳款-關係人（皆無值回 None）。"""
    parts = [vals.get(k) for k in ("應收票據", "應收帳款", "應收帳款-關係人")]
    if all(p is None for p in parts):
        return None
    return sum(p or 0 for p in parts)


def turnover(numer, end, beg):
    """週轉率＝分子÷平均餘額；期初或期末缺任一即回 None（不得只用單期充當平均）。"""
    if numer is None or end is None or beg is None or (end + beg) == 0:
        return None
    return round(numer / ((end + beg) / 2), 2)


ASSET_ITEMS = ["現金及約當現金", "應收票據", "應收帳款", "應收帳款-關係人", "其他應收款",
               "其他應收款-關係人", "預付款項", "存貨", "採用權益法之投資",
               "不動產廠房及設備", "使用權資產"]


def asset_mix(vals: dict) -> list:
    """資產科目佔資產總額比重，由大到小——風險事項優先從占比大的科目挑。"""
    total = vals.get("資產總額")
    if not total:
        return []
    out = []
    for it in ASSET_ITEMS:
        v = vals.get(it)
        if v:
            out.append({"item": it, "amount": v, "pct": round(v / total * 100, 1)})
    return sorted(out, key=lambda x: -x["pct"])


# 查詢函「二、與上市櫃同業比較」及檢查表個別資料庫(四)固定比較的六項比率
INQUIRY_RATIOS = ["營收成長率", "銷貨毛利率", "營業損益率", "稅前損益成長率",
                  "應收款項週轉率", "存貨週轉率"]


def company_ratios(series: dict, y: int) -> dict:
    """個別公司十一比率（順序＝個別資料庫查詢系統「財務比率」表）。缺基礎數＝None。

    前六項與公開資訊觀測站財務業務資訊同口徑（INQUIRY_RATIOS 供同業比較）；
    後五項為財務比率表之補充比率，依 XBRL 財報數字計算，口徑可能與資料庫略異。
    """
    cur, prev = series.get(y, {}), series.get(y - 1, {})
    rev, rev_p = cur.get("營業收入"), prev.get("營業收入")

    def ratio(numer, denom):
        return round(numer / denom * 100, 2) if (numer is not None and denom) else None

    ca, cl = cur.get("流動資產"), cur.get("流動負債")
    quick = (ca - (cur.get("存貨") or 0) - (cur.get("預付款項") or 0)) if ca is not None else None
    r = {
        "營收成長率": pct(rev, rev_p),
        "銷貨毛利率": ratio(cur.get("營業毛利"), rev),
        "營業損益率": ratio(cur.get("營業損益"), rev),
        "稅前損益成長率": pct(cur.get("稅前損益"), prev.get("稅前損益")),
        "流動比率": ratio(ca, cl),
        "負債比率": ratio(cur.get("負債總額"), cur.get("資產總額")),
        "應收款項週轉率": turnover(rev, receivables(cur), receivables(prev)),
        "存貨週轉率": turnover(cur.get("營業成本"), cur.get("存貨"), prev.get("存貨")),
        "現金流量比率": ratio(cur.get("營業活動現金流量"), cl),
        "股東權益報酬率": (round(cur["本期淨利"] / ((cur["淨值"] + prev["淨值"]) / 2) * 100, 2)
                       if all(x is not None for x in
                              (cur.get("本期淨利"), cur.get("淨值"), prev.get("淨值")))
                       and (cur["淨值"] + prev["淨值"]) else None),
        "速動比率": ratio(quick, cl),
    }
    return r


# ---------- 查詢函題目 ----------

PL_ITEMS = ["營業收入", "營業毛利", "營業損益", "稅前損益"]


def build_questions(series, y, ratios, ratios_prev, peer, tuples, meta, qmap=None):
    """回傳題目清單；qmap 為 dict 時另記錄「觸發項目 → 題號」，供分析表回溯完整性。"""
    roc = to_roc(y)
    cur, prev = series.get(y, {}), series.get(y - 1, {})
    qs = []

    def add(cat, text, prefill="", key=None):
        qid = f"Q{len(qs)+1:02d}"
        qs.append({"id": qid, "category": cat, "question": text, "prefill": prefill})
        if qmap is not None and key:
            qmap.setdefault(key, []).append(qid)

    # 1. 損益四項變動達 30%
    for it in PL_ITEMS:
        ch = pct(cur.get(it), prev.get(it))
        if ch is not None and abs(ch) >= 30:
            add("損益變動",
                f"{roc}年度{it}較{roc-1}年度變動達30%以上，請說明變動原因及合理性。",
                f"{roc}年度 {fmt(cur.get(it))}／{roc-1}年度 {fmt(prev.get(it))}／"
                f"變動 {fmt_pct(ch)}", key=f"變動:{it}")

    # 2. 變動成長率（本期成長率－去年同期成長率）差異達 10 個百分點
    prev2 = series.get(y - 2, {})
    for it in PL_ITEMS + ["綜合損益"]:
        g_cur = pct(cur.get(it), prev.get(it))
        g_prev = pct(prev.get(it), prev2.get(it))
        if g_cur is not None and g_prev is not None and abs(g_cur - g_prev) >= 10:
            add("成長率差異",
                f"{it}成長率本期與去年同期差異達10個百分點以上，請說明原因。",
                f"{roc}年度成長率 {fmt_pct(g_cur)}／{roc-1}年度成長率 {fmt_pct(g_prev)}／"
                f"差異 {abs(round(g_cur - g_prev, 2))} 個百分點", key=f"成長率差異:{it}")

    # 3. 週轉率變動達 10%
    for name in ("應收款項週轉率", "存貨週轉率"):
        t_cur, t_prev = ratios.get(name), (ratios_prev or {}).get(name)
        ch = pct(t_cur, t_prev)
        if ch is not None and abs(ch) >= 10:
            add("週轉率",
                f"{name}本期與去年同期變動差異達10%以上，請說明原因。",
                f"{roc}年度 {t_cur}次／{roc-1}年度 {t_prev}次／變動 {fmt_pct(ch)}",
                key=f"變動:{name}")

    # 4. 與同業平均比較（固定六項；無資料則出題但留待貼入）
    for rname in INQUIRY_RATIOS:
        rv = ratios.get(rname)
        pv = (peer.get("上市櫃同業平均", {}) or {}).get(rname) if peer else None
        if pv is not None and rv is not None:
            unit = "次" if "週轉" in rname else "%"
            diff = round(rv - pv, 2)
            trigger = (abs(pct(rv, pv) or 0) >= 10) if "週轉" in rname else (abs(diff) >= 10)
            if trigger:
                add("同業比較",
                    f"{rname}與上市櫃同業平均差異達10%（或10個百分點）以上，請說明原因。",
                    f"公司 {rv}{unit}／上市櫃同業平均 {pv}{unit}", key=f"同業:{rname}")
        elif rv is not None:
            add("同業比較",
                f"{rname}與上市櫃同業平均及所有同業平均之比較，如差異達10%（或10個百分點）"
                f"以上請說明原因。（同業平均請自公開資訊觀測站財務業務資訊查填）",
                f"公司 {rv}{'次' if '週轉' in rname else '%'}／同業平均：請自公開資訊觀測站貼入",
                key=f"同業:{rname}")

    # 5. 其他應收款／預付款項／轉投資大幅增加（增幅≥30% 且期末達資產總額 1%）
    assets = cur.get("資產總額")
    for it in ("其他應收款", "其他應收款-關係人", "預付款項", "採用權益法之投資"):
        ch = pct(cur.get(it), prev.get(it))
        material = assets and cur.get(it) and cur[it] >= assets * 0.01
        if ch is not None and ch >= 30 and material:
            add("資產負債項目",
                f"{it}本期大幅增加，請說明增加原因、性質及必要性。",
                f"{roc}年度 {fmt(cur.get(it))}／{roc-1}年度 {fmt(prev.get(it))}／"
                f"增加 {fmt_pct(ch)}", key=f"變動:{it}")

    # 5.5 占比最大的資產科目——就前三大（且達資產10%）請公司說明組成與評價
    for a in [x for x in asset_mix(cur) if x["pct"] >= 10][:3]:
        ch = pct(cur.get(a["item"]), prev.get(a["item"]))
        add("重大資產科目",
            f"{a['item']}佔資產總額{a['pct']}%，為重大資產項目，請說明其組成內容、"
            f"評價方法及重要假設，並說明有無減損跡象及評估情形。",
            f"{roc}年度 {fmt(a['amount'])}（佔資產{a['pct']}%）"
            + (f"／{roc-1}年度 {fmt(prev.get(a['item']))}／變動 {fmt_pct(ch)}"
               if ch is not None else ""), key=f"占比:{a['item']}")

    # 6. 減損
    if cur.get("減損損失"):
        add("減損",
            "本期認列資產減損損失，請說明減損標的、減損跡象、可回收金額之評估方法"
            "與重要假設（折現率、成長率等），並提供評價報告。",
            f"{roc}年度認列減損損失 {fmt(cur.get('減損損失'))}", key="變動:減損損失")

    # 7. OCI
    if cur.get("其他綜合損益") is not None:
        add("其他綜合損益",
            "請說明本期其他綜合損益組成項目之變動內容及原因。",
            f"{roc}年度 {fmt(cur.get('其他綜合損益'))}／{roc-1}年度 {fmt(prev.get('其他綜合損益'))}",
        key="變動:其他綜合損益")

    # 8. IFRS 專項（各一題，帶財報既有數字）
    add("IFRS16",
        "請說明公司租賃標的、租賃期間、使用權資產折舊方法及租賃負債衡量方式，"
        "是否依IFRS16認列及衡量；與關係人之租賃並請說明租金條件之常規性。",
        f"使用權資產 {fmt(cur.get('使用權資產'))}／租賃負債-流動 {fmt(cur.get('租賃負債-流動'))}"
        f"／租賃負債-非流動 {fmt(cur.get('租賃負債-非流動'))}")
    add("IFRS9",
        "請說明公司金融資產之分類（攤銷後成本／FVOCI／FVTPL）、預期信用損失之評估方法"
        "（含損失率訂定依據與帳齡分析），是否依IFRS9認列衡量並依IFRS7揭露。",
        f"現金及約當現金 {fmt(cur.get('現金及約當現金'))}／應收款項合計 "
        f"{fmt(receivables(cur))}／其他應收款-關係人 {fmt(cur.get('其他應收款-關係人'))}")
    add("IFRS15",
        "請說明公司各類收入之認列時點（某一時點／隨時間逐步）、合約負債之性質及沖轉情形，"
        "是否依IFRS15認列衡量及揭露。",
        f"營業收入 {fmt(cur.get('營業收入'))}／合約負債-流動 {fmt(cur.get('合約負債-流動'))}")
    inv_names = "、".join(
        f"{t.get('CompanyNameOfTheInvestee') or t.get('NameOfInvestee') or '（名稱見財報附表）'}"
        for t in (tuples.get("NamesLocationsAndRelatedInformationOfInvesteesOverWhich"
                             "TheCompanyExercisesSignificantInfluence") or [])[:6]) or "（詳財報附表）"
    add("IFRS10",
        "對被投資公司持股未逾50%惟採權益法者，請說明依IFRS10第B38~B50段評估控制／"
        "重大影響力之判斷過程（董事席次、綜合持股、參與決策情形）及相關佐證。",
        f"採用權益法之投資期末 {fmt(cur.get('採用權益法之投資'))}／被投資公司：{inv_names}")

    # 9. 資金貸與及背書保證題組（五連問）
    loans = tuples.get("LoansToOthers") or []
    endos = tuples.get("EndorsementGuaranteeProvidedToOthers") or []
    loan_sum = sum(t.get("EndingBalance1") or 0 for t in loans)
    endo_sum = sum(t.get("EndingBalance2") or 0 for t in endos)
    base = (f"資金貸與期末餘額 {fmt(loan_sum)}（{len(loans)}筆）／"
            f"背書保證期末餘額 {fmt(endo_sum)}（{len(endos)}筆）")
    add("資金貸與背書保證", "公司是否訂有「資金貸與及背書保證作業程序」？最近修訂日期為何？"
        "請提供全文。", base)
    add("資金貸與背書保證", "是否依規定設置資金貸與及背書保證備查簿，並逐筆詳實登載？"
        "請提供備查簿影本。", base)
    add("資金貸與背書保證", "各筆資金貸與及背書保證是否逐案經董事會決議（或依授權辦理）？"
        "必要性與合理性如何評估？請提供董事會議事錄及評估文件。", base)
    add("資金貸與背書保證", "是否依「公開發行公司資金貸與及背書保證處理準則」辦理公告申報"
        "（每月10日前及達限額標準時）？", base)
    add("資金貸與背書保證", "請說明期末各筆餘額之對象、原因、利率及計息情形，"
        "與內規限額（個別／總額）之遵循情形。",
        "；".join(f"{t.get('Counterparty1', '?')} {fmt(t.get('EndingBalance1'))} "
                  f"利率{t.get('RangeOfInterestRates', '?')}" for t in loans) or base)

    # 10. 會計主管資格題組（四連問）
    add("會計主管", "請提供會計主管之姓名、學經歷及到任日期。", "")
    add("會計主管", "會計主管是否符合「發行人證券商證券交易所會計主管資格條件及專業進修辦法」"
        "第3條所定資格條件？符合哪一款？請提供證明文件。", "")
    add("會計主管", "會計主管最近年度持續進修時數為何？是否已依同辦法申報進修情形？"
        "請提供進修時數證明及申報紀錄。", "")
    add("會計主管", "會計主管有無同辦法第4條所列消極資格情事？", "")
    return qs


def build_analysis(series, y, ratios, ratios_prev, peer, qmap):
    """全科目分析表——每個科目都列，看得出「都分析過」而非只列達標者。

    欄位：類別／項目／本期／上期／增減／變動%／分析門檻／是否達標／對應題號。
    這是回溯完整性的依據：未達標者也在表上並註明門檻，不是漏掉。
    """
    roc = to_roc(y)
    cur, prev, prev2 = series.get(y, {}), series.get(y - 1, {}), series.get(y - 2, {})
    total = cur.get("資產總額")
    rows = []

    def row(cat, item, c, p, thr, hit, key=None, extra=""):
        d = (c - p) if (c is not None and p is not None) else None
        # 前期為 0／無值而本期有數＝本期新增，屬應說明事項（變動率算不出來不代表沒事）
        new_item = bool(c) and not p
        rows.append({
            "類別": cat, "項目": item,
            f"{roc}年度": c, f"{roc - 1}年度": p,
            "增減": round(d, 4) if isinstance(d, float) else d,
            "變動%": pct(c, p),
            "分析門檻": thr,
            "是否達標": ("★達標（本期新增）" if new_item
                       else "－（缺資料）" if c is None
                       else "★達標" if hit else "未達標"),
            "對應題號": "、".join(qmap.get(key, [])) if key else "",
            "備註": extra,
        })

    # 四大損益項（金管會檢查表指定）與其他損益科目分開標示門檻依據
    main_qs = "、".join(dict.fromkeys(
        sum((qmap.get(f"變動:{it}", []) for it in PL_ITEMS), [])))
    for it in PL_ITEMS:
        ch = pct(cur.get(it), prev.get(it))
        row("損益（檢查表指定）", it, cur.get(it), prev.get(it), "變動達30%",
            ch is not None and abs(ch) >= 30, f"變動:{it}")
    for it in ["本期淨利", "其他綜合損益", "綜合損益", "減損損失",
               "採用權益法之投資損益份額", "每股盈餘(元)"]:
        ch = pct(cur.get(it), prev.get(it))
        hit = ch is not None and abs(ch) >= 30
        note = ""
        if hit and not qmap.get(f"變動:{it}"):
            note = (f"變動原因與營業損益／稅前損益同源，併同 {main_qs} 說明"
                    if main_qs else "請併同損益變動題項說明")
        row("損益（延伸覆核）", it, cur.get(it), prev.get(it), "變動達30%",
            hit, f"變動:{it}", note)

    for it in PL_ITEMS + ["綜合損益"]:
        g_c, g_p = pct(cur.get(it), prev.get(it)), pct(prev.get(it), prev2.get(it))
        d = abs(g_c - g_p) if (g_c is not None and g_p is not None) else None
        rows.append({
            "類別": "成長率差異", "項目": f"{it}成長率",
            f"{roc}年度": g_c, f"{roc - 1}年度": g_p,
            "增減": round(d, 2) if d is not None else None, "變動%": None,
            "分析門檻": "兩期成長率差異達10個百分點",
            "是否達標": ("－（缺資料）" if d is None else ("★達標" if d >= 10 else "未達標")),
            "對應題號": "、".join(qmap.get(f"成長率差異:{it}", [])),
            "備註": "基期為負，成長率不具比較意義" if (prev2.get(it) or 0) < 0 else "",
        })

    for it in ASSET_ITEMS + ["資產總額", "負債總額", "淨值", "合約負債-流動",
                             "租賃負債-流動", "租賃負債-非流動", "應付帳款及票據"]:
        c, p = cur.get(it), prev.get(it)
        ch = pct(c, p)
        share = round(c / total * 100, 1) if (c and total) else None
        big = share is not None and share >= 10 and it in ASSET_ITEMS
        jump = ch is not None and ch >= 30 and c and total and c >= total * 0.01
        row("資產負債", it, c, p, "佔資產達10% 或 增加達30%", bool(big or jump),
            f"占比:{it}" if big else f"變動:{it}",
            f"佔資產總額 {share}%" if share is not None else "")

    for name in INQUIRY_RATIOS:
        v, pv = ratios.get(name), ratios_prev.get(name)
        ch = pct(v, pv)
        peer_v = ((peer or {}).get("上市櫃同業平均") or {}).get(name)
        rows.append({
            "類別": "財務比率", "項目": name,
            f"{roc}年度": v, f"{roc - 1}年度": pv,
            "增減": round(v - pv, 2) if (v is not None and pv is not None) else None,
            "變動%": ch,
            "分析門檻": "兩期變動達10%；與同業平均差異達10%",
            "是否達標": ("－（缺資料）" if v is None or pv is None
                       else ("★達標" if (ch is not None and abs(ch) >= 10) else "未達標")),
            "對應題號": "、".join(dict.fromkeys(
                qmap.get(f"變動:{name}", []) + qmap.get(f"同業:{name}", []))),
            "備註": (f"上市櫃同業平均 {peer_v}" if peer_v is not None
                   else "同業平均未提供，請自公開資訊觀測站財務業務資訊查填後比較"),
        })
    return rows


def build_diff_sections(series, y, ratios, ratios_prev, peer, tuples, tuples_prev,
                        report_category=None):
    """「差異說明」工作表各節的數字列。

    版面依過去實審實際樣本（110／112／114 年度三份「財務比率差異分析說明」），
    節次採最新版式：一 損益與成長率、二 同業比較、三 資產負債變動、五 風險事項、
    七 資金貸與及背書保證；固定題文（四、六、八）由 gen_inquiry_xlsx.py 排版。
    """
    roc = to_roc(y)
    cur, prev, prev2 = series.get(y, {}), series.get(y - 1, {}), series.get(y - 2, {})

    # 一(A) 損益四項金額：首列「說明」欄放引導題（樣本慣例）；門檻不點名個別科目，
    # 統一請公司就變動達30%者加強說明（達標與否的內部判斷見「科目分析(內部)」表）
    lead_q = ("請說明貴公司營業收入主要來源(ex.何種商品或服務等)，並據以分析"
              f"{roc}年度營業收入淨額、營業毛利、營業損益及稅前損益之變動原因及合理性；"
              "變動達30%以上之項目請加強說明。")
    sec1_amounts = []
    for i, it in enumerate(PL_ITEMS, 1):
        c, p = cur.get(it), prev.get(it)
        sec1_amounts.append({
            "項次": i, "項目": "營業收入淨額" if it == "營業收入" else it,
            "本期": c, "前期": p,
            "增減": (c - p) if (c is not None and p is not None) else None,
            "變動%": pct(c, p), "說明": lead_q if i == 1 else "",
        })

    # 一(B) 成長率兩期比較（百分點）＋週轉率兩期變動（%）
    # 說明欄不逐項點名達標科目，統一放一句（首列）——數字差異照列，判斷交由公司
    growth_note = ("請就變動差異達10%（成長率達10個百分點）以上之項目，"
                   "說明兩期變動原因及合理性。")
    sec1_growth = []
    for i, it in enumerate(PL_ITEMS, 1):
        g_c, g_p = pct(cur.get(it), prev.get(it)), pct(prev.get(it), prev2.get(it))
        d = round(g_c - g_p, 2) if (g_c is not None and g_p is not None) else None
        sec1_growth.append({
            "項次": i, "項目": "營收成長率" if it == "營業收入" else f"{it}成長率",
            "本期": g_c, "前期": g_p, "增減": d, "增減單位": "百分點",
            "說明": growth_note if i == 1 else "",
        })
    for i, nm in enumerate(("應收款項週轉率", "存貨週轉率"), len(PL_ITEMS) + 1):
        t_c, t_p = ratios.get(nm), (ratios_prev or {}).get(nm)
        ch = pct(t_c, t_p)
        sec1_growth.append({
            "項次": i, "項目": nm, "本期": t_c, "前期": t_p, "增減": ch, "增減單位": "%",
            "說明": "",
        })

    # 二 與上市櫃同業比較（六項；率＝百分點差、週轉率＝相對%差）
    # 使用者未提供同業平均時（常態），改請公司自行選擇相近上市櫃同業比較：
    # 同業欄留白由公司填，說明欄僅首列放一句統一指引，不逐項出題
    sec2 = []
    for i, nm in enumerate(INQUIRY_RATIOS, 1):
        rv = ratios.get(nm)
        pv = (peer.get("上市櫃同業平均", {}) or {}).get(nm) if peer else None
        turn = "週轉" in nm
        diff, note = None, ""
        if rv is not None and pv is not None:
            diff = pct(rv, pv) if turn else round(rv - pv, 2)
            if diff is not None and abs(diff) >= 10:
                note = "差異超過10%，請分析貴公司與上市櫃同業差異原因"
        elif i == 1:
            note = "請填列所選上市櫃同業之公司名稱及各項數據，並就差異說明原因及合理性。"
        sec2.append({"項次": i, "項目": nm, "本期": rv, "同業": pv, "增減": diff,
                     "增減單位": "%" if turn else "百分點", "說明": note})

    # 三 資產負債重大變動：|變動|≥30%且達資產1%，或占資產≥10%之重大科目（現金除外）
    total = cur.get("資產總額")
    big = {a["item"] for a in asset_mix(cur) if a["pct"] >= 10}
    sec3 = []
    for it in ASSET_ITEMS + ["應付帳款及票據", "合約負債-流動"]:
        c, p = cur.get(it), prev.get(it)
        ch = pct(c, p)
        material = total and c and c >= total * 0.01
        if ((ch is not None and abs(ch) >= 30 and material)
                or (it in big and it != "現金及約當現金")):
            sec3.append({"項次": len(sec3) + 1, "項目": it, "本期": c, "前期": p,
                         "增減": (c - p) if (c is not None and p is not None) else None,
                         "變動%": ch, "說明": ""})
    sec3 = sec3[:8]

    # 五 風險事項（占資產≥10%前三大，現金除外；與管區意見候選風險同源）
    sec5 = []
    for a in [x for x in asset_mix(cur)
              if x["pct"] >= 10 and x["item"] != "現金及約當現金"][:3]:
        sec5.append({
            "項目": a["item"], "金額": a["amount"], "占比": a["pct"],
            "問題": (f"貴公司{roc}年度{a['item']}金額為{fmt(a['amount'], '')}千元，"
                    f"佔資產總額{a['pct']}%，對財務報表影響重大，請說明其組成內容、"
                    "評價方法及有無減損跡象，暨貴公司之因應措施(如何執行評價、資產保全等)。"),
        })

    # 七 資金貸與及背書保證（前期無該年度 pretrip 時留 None＝Excel 留白）
    def sums(tp):
        loans = (tp or {}).get("LoansToOthers") or []
        endos = (tp or {}).get("EndorsementGuaranteeProvidedToOthers") or []
        return (sum(t.get("EndingBalance1") or 0 for t in loans),
                sum(t.get("EndingBalance2") or 0 for t in endos))

    l_c, e_c = sums(tuples)
    l_p, e_p = sums(tuples_prev) if tuples_prev is not None else (None, None)
    sec7 = {"資金貸與他人金額": {"本期": l_c, "前期": l_p},
            "背書保證金額": {"本期": e_c, "前期": e_p}}

    # 一(C) 合併報表三列：申報之XBRL即為合併報表者，數字恆等於上表對應科目、直接帶入；
    # 僅申報個別報表者留白（公司如有編合併報表再自填）
    def recv_sum(d):
        vals = [d.get(k) for k in ("應收票據", "應收帳款", "應收帳款-關係人")
                if d.get(k) is not None]
        return sum(vals) if vals else None

    consolidated = (report_category or "").startswith("Consolidated")
    sec1_cons = {"consolidated": consolidated, "rows": []}
    if consolidated:
        for nm, c, p in [("合併報表營業收入淨額", cur.get("營業收入"), prev.get("營業收入")),
                         ("合併報表應收款項淨額\n(含應收票據)", recv_sum(cur), recv_sum(prev)),
                         ("合併報表存貨淨額", cur.get("存貨"), prev.get("存貨"))]:
            sec1_cons["rows"].append({
                "項目": nm, "本期": c, "前期": p,
                "增減": (c - p) if (c is not None and p is not None) else None,
                "變動%": pct(c, p),
            })

    return {"sec1_amounts": sec1_amounts, "sec1_growth": sec1_growth,
            "sec1_consolidated": sec1_cons,
            "sec2_peer": sec2, "sec3_bs": sec3, "sec5_risk": sec5, "sec7_loans": sec7}


# ---------- 管區意見骨架 ----------

AI = "【AI待填："  # 佔位字首；check_content.py 以此把關


def build_checklist_draft(co, roc, meta, audit, series, y, ratios, ratios_prev,
                          peer, tuples, red_flags):
    cur, prev = series.get(y, {}), series.get(y - 1, {})
    name = meta.get("CompanyChineseName") or f"（{co}）"
    equity = cur.get("淨值")

    def chg(it):
        return pct(cur.get(it), prev.get(it))

    # --- 檢查表 18 項（固定文字照 golden 模板；數字面 note 由此填、質性 note 留 AI） ---
    ar_up = (chg("應收帳款") or 0) >= 30 or (chg("應收票據") or 0) >= 30
    ap_down = (chg("應付帳款及票據") or 0) < 0
    imp = cur.get("減損損失")
    groups = [
      {"group": "個別資料庫", "items": [
        {"id": "db1", "text": "1.該公司本期營運情形與所屬產業成長（或衰退）之變動是否相同及合理。",
         "mark": None, "note": "詳個別資料庫說明"},
        {"id": "db2", "text": "2.損益表\n查詢公司財務業務資訊有關本期與去年同期之合併營業收入淨額、合併營業毛利、合併營業損益及合併稅前損益金額變動差異達30%，有無重大異常變動。\n若前開變動率已達30%，將加強查核合併個體內重要子公司之合併營業收入淨額、合併營業毛利、合併營業損益及合併稅前損益之變動情形，並就重要子公司差異達30%以上者再詳細說明有無重大異常變動。\n另合併營業收入淨額、合併營業毛利、合併營業損益及綜合損益變動成長率（或衰退率），本期與去年同期之變動率(本期－去年同期)差異若達10%以上，有無重大異常變動。",
         "mark": None, "note": "詳個別資料庫說明"},
        {"id": "db3", "text": "3.週轉率\n查詢公司財務業務資訊之合併應收帳款週轉率、及合併存貨週轉率，本期與去年同期之變動率(本期－去年同期)差異若達10%以上，有無重大異常變動。",
         "mark": None, "note": "詳個別資料庫說明"},
        {"id": "db4", "text": "4.成長率\n查詢公司合併營收成長率、合併營業毛利率、合併營業損益率、合併稅前損益成長率、合併應收帳款週轉率、合併存貨週轉率與上市櫃公司同業平均比較，本期與上市櫃同業平均之變動率(本期－上市櫃同業平均)差異若達10%，有無重大異常變動。",
         "mark": None, "note": "詳個別資料庫說明"},
      ]},
      {"group": "資產負債表及損益表", "items": [
        {"id": "bs1", "text": "1.本期較去年同期應收帳款及票據大幅增加，惟本期應付帳款及票據不增反減，有無重大異常。",
         "mark": None,
         "note": (AI + "應收增且應付減，請判斷有無異常】" if (ar_up and ap_down)
                  else f"無此情事（應收款項變動{fmt_pct(chg('應收帳款'))}）")},
        {"id": "bs2", "text": "2.本期較去年同期應收帳款、存貨、其它應收款、預付款及轉投資大幅增加，有無重大異常。本期與關係人大幅背書保證或資金貸與，有無重大異常。",
         "mark": None, "note": "詳個別資料庫及資金貸與說明"},
        {"id": "bs3", "text": "3.本期較去年同期銷貨增加，惟本期機器設備無增加，且推銷費用不增反降，有無重大異常。",
         "mark": None,
         "note": ("無此情事（本期銷貨未增加）" if (chg("營業收入") or 0) <= 0
                  else AI + "銷貨增加，請核對設備與推銷費用變動】")},
        {"id": "bs4", "text": "4.本期應收關係人款未收回，是否已提列備抵呆帳，且是否與關係人間銷貨仍持續將增加，有無重大異常。",
         "mark": None, "note": AI + "依帳齡與關係人交易判斷，一句話】"},
        {"id": "bs5", "text": "5.本期與關係人間背書保證或資金貸與，有無重大異常。",
         "mark": None, "note": "詳資金貸與及背書保證說明"},
        {"id": "bs6", "text": "6.本期有無重大關係人交易（註1）。",
         "mark": None, "note": AI + "依 tuples 關係人交易金額判斷，一句話】"},
        {"id": "bs7", "text": "7.本期有無重大資產減損提列情形。",
         "mark": None,
         "note": (f"本期認列減損損失{fmt(imp)}，詳說明" if imp else "無此情事")},
        {"id": "bs8", "text": "8.本期有無認列資產減損迴轉金額，是否有重大異常。",
         "mark": None, "note": AI + "查現流表減損迴轉科目，多為「無此情事」】"},
        {"id": "bs10", "text": "10.不動產、廠房及設備暨不動產之重大組成項目，及投資性不動產公允價值、方法及假設是否符合規定。",
         "mark": None, "note": AI + "無投資性不動產則寫「尚無重大異常（無投資性不動產）」】"},
        {"id": "bs11", "text": "11.自願改變會計政策或會計估計值變動中屬折舊(耗)性資產耐用年限、折舊（耗）方法與無形資產攤銷期間、攤銷方法之變動、殘值之變動及其公允價值之評價技術變動所致者，是否依證券發行人財務報告編製準則規定辦理、其他綜合損益組成項目變動是否合理。",
         "mark": None, "note": AI + "無變動則「無此情事」】"},
        {"id": "bs12", "text": "12.其它綜合損益組成項目變動是否合理。",
         "mark": None, "note": "尚無重大異常"},
      ]},
      {"group": "會計主管", "items": [
        {"id": "acc1", "text": "於財務報告簽名蓋章之會計主管是否符合發行人證券商證券交易所會計主管資格條件及專業進修辦法之規定。",
         "mark": None, "note": "詳會計主管審閱說明"},
      ]},
      {"group": "資金貸與及背書保證", "items": [
        {"id": "loan1", "text": "公司從事資金貸與及背書保證是否符合公開發行公司資金貸與及背書保證處理準則規定。",
         "mark": None, "note": "詳資金貸與及背書保證審閱說明"},
      ]},
    ]

    # --- 五段說明骨架 ---
    # 一、風險事項候選（數字訊號＋AI 判斷取捨）
    # 占比最大的資產科目優先列為風險事項——帳面重大者一旦評價有誤，對財報影響最大
    risk_paras = []
    mix = asset_mix(cur)
    for a in [x for x in mix if x["pct"] >= 10][:3]:
        it, amt, p = a["item"], a["amount"], a["pct"]
        ch = pct(amt, prev.get(it))
        base = (f"「{it}」期末帳面金額{fmt(amt)}，佔資產總額之{p}%"
                + (f"，較{roc - 1}年度變動{fmt_pct(ch)}" if ch is not None else "")
                + "，對財務報表影響重大，故將其列為風險事項。")
        if it == "採用權益法之投資":
            risk_paras.append({"h": f"（候選風險）{it}之減損評估：", "paras": [
                base + AI + "逐一列出被投資公司帳面金額／持股／本期損益"
                "（見 facts.investees 與 red_flags）；減損評估之因應措施屬公司才答得出者"
                "寫「擬行前查證」】"]})
        elif it == "現金及約當現金":
            continue          # 現金部位大非屬風險，除非有其他訊號，交由 AI 判斷
        else:
            risk_paras.append({"h": f"（候選風險）{it}之評價及組成：", "paras": [
                base + AI + f"依財報附註說明{it}之組成與評價方法；公司才答得出的部分"
                "（評價假設、個別客戶／標的狀況）寫「擬行前請公司說明」；"
                "若經判斷不成立則整段刪除】"]})
    ptx_chg = chg("稅前損益")
    if ptx_chg is not None and ptx_chg <= -30:
        risk_paras.append({"h": "（候選風險）獲利持續衰退：", "paras": [
            f"{roc}年度稅前損益{fmt(cur.get('稅前損益'))}，較{roc - 1}年度"
            f"{fmt(prev.get('稅前損益'))}變動{fmt_pct(ptx_chg)}；本期淨利"
            f"{fmt(cur.get('本期淨利'))}（EPS {cur.get('每股盈餘(元)', '－')}元）。"
            + AI + "衰退主因分解（毛利、減損、權益法損益），近六年營收趨勢見 six_year】"]})
    loans = tuples.get("LoansToOthers") or []
    endos = tuples.get("EndorsementGuaranteeProvidedToOthers") or []
    loan_sum = sum(t.get("EndingBalance1") or 0 for t in loans)
    endo_sum = sum(t.get("EndingBalance2") or 0 for t in endos)
    if equity and (loan_sum + endo_sum) / equity >= 0.10:
        risk_paras.append({"h": "（候選風險）集團資金相互支援：", "paras": [
            f"對關係人資金貸與期末餘額{fmt(loan_sum)}加計背書保證期末餘額{fmt(endo_sum)}，"
            f"合計佔淨值{round((loan_sum + endo_sum) / equity * 100, 1)}%。"
            + AI + "補集團關係與持股（如有母公司合併報告 pretrip 可查證），詳說明五】"]})
    risk_paras.append({"h": "（候選風險）其他關注事項：", "paras": [
        AI + "自 audit.flags、red_flags、權益變動（特別盈餘公積、累積虧損）、"
        "會計師異動等挑實質事項；無則整段刪除】"]})

    # 二、個別資料庫四項（數字句 Python 寫死，原因句留 AI）
    pl_lines = "、".join(
        f"{it}{fmt_pct(chg(it))}" for it in PL_ITEMS if chg(it) is not None)
    hits30 = [it for it in PL_ITEMS if chg(it) is not None and abs(chg(it)) >= 30]
    g_diff_lines = []
    prev2 = series.get(y - 2, {})
    for it in PL_ITEMS + ["綜合損益"]:
        g_cur_, g_prev_ = pct(cur.get(it), prev.get(it)), pct(prev.get(it), prev2.get(it))
        if g_cur_ is not None and g_prev_ is not None and abs(g_cur_ - g_prev_) >= 10:
            g_diff_lines.append(f"{it}成長率{fmt_pct(g_cur_)}對{fmt_pct(g_prev_)}"
                                f"（差異{abs(round(g_cur_ - g_prev_, 2))}個百分點）")
    tv_lines = []
    for nm in ("應收款項週轉率", "存貨週轉率"):
        t_cur, t_prev = ratios.get(nm), (ratios_prev or {}).get(nm)
        ch = pct(t_cur, t_prev)
        if ch is not None:
            tv_lines.append(f"{nm}較{roc - 1}年度變動{fmt_pct(ch)}（{t_cur}次對{t_prev}次）")
    peer_rows = ["【表】項目｜" + f"{roc}年度｜上市櫃同業平均｜比較增減｜說明"]
    for rname in INQUIRY_RATIOS:
        rv = ratios.get(rname)
        unit = "次" if "週轉" in rname else "%"
        pv = (peer.get("上市櫃同業平均", {}) or {}).get(rname) if peer else None
        peer_rows.append(
            f"【表】{rname}｜{('－' if rv is None else f'{rv}{unit}')}｜"
            f"{('請自公開資訊觀測站貼入' if pv is None else f'{pv}{unit}')}｜"
            + AI + "計算增減】｜" + AI + "差異原因一句話】")

    sections = [
      {"title": "一、公司之風險事項：", "body": risk_paras or
       [{"paras": [AI + "無數字面候選風險時仍應綜合判斷是否列風險事項】"]}]},
      {"title": "二、個別資料庫", "body": [
        {"h": "（一）該公司本期營運情形與所屬產業成長（或衰退）之變動是否相同及合理。",
         "paras": [
            f"該公司{roc}年度營業收入較{roc - 1}年度變動{fmt_pct(chg('營業收入'))}"
            f"（近年營收成長率見 facts.six_year）。"
            + AI + "所屬產業趨勢與同業平均（無資料則標「行前請至公開資訊觀測站查填」）、"
            "公司營運內容，收尾「其與所屬產業變動相同/不同，核尚無重大異常」】"]},
        {"h": "（二）本期與去年同期之營業收入淨額、營業毛利、營業損益及稅前損益金額變動差異達30%，有無重大異常變動。",
         "paras": [
            f"1.營業收入淨額、營業毛利、營業損益及稅前損益較{roc - 1}年度分別變動"
            f"{pl_lines}，其中{('、'.join(hits30) + '變動達30%') if hits30 else '均未達30%'}。"
            + AI + "變動原因（公司未回覆前屬推測者寫「擬行前請公司說明」）；"
            "無子公司者敘明「無合併個體內重要子公司加強查核之適用」】",
            ("2.變動成長率比較（本期變動率減去年同期變動率）：" + "、".join(g_diff_lines)
             + "，均達10%以上。" + AI + "原因；基期為負者敘明成長率不具比較意義】")
            if g_diff_lines else
            "2.變動成長率比較：本期與去年同期之變動率差異均未達10%，核尚無重大異常。"]},
        {"h": "（三）本期與去年同期應收帳款週轉率及存貨週轉率變動差異達10%，有無重大異常變動。",
         "paras": [(("；".join(tv_lines) + "。") if tv_lines else
                    "週轉率因缺前期平均餘額基礎數，無法由財報自動計算，行前請至公開資訊"
                    "觀測站財務業務資訊查填。")
                   + AI + "逐一說明變動原因；金額微小科目敘明比率變動不具分析意義】"]},
        {"h": "（四）與上市櫃公司同業平均比較，變動差異是否達10%，有無重大異常變動。",
         "paras": ["該公司與上市櫃同業平均（公開資訊觀測站財務業務資訊）各比率暨說明如下表：",
                   *peer_rows,
                   AI + "總結：差異達10%項目之結構性原因，收尾「核尚無重大異常」】"]},
      ]},
      {"title": "三、財務報告審閱說明", "body": [
        {"h": "（一）使用權資產及租賃負債是否依IFRS16規定認列及衡量並揭露攸關資訊。",
         "paras": [
            AI + "認列及衡量：依 facts.notes.Leasing 原文摘寫該公司租賃政策（租賃標的、"
            "豁免適用、使用權資產與租賃負債之衡量）；該節缺漏才寫「擬行前核閱財報附註確認」】",
            f"財報揭露：已認列使用權資產{fmt(cur.get('使用權資產'))}及租賃負債-流動"
            f"{fmt(cur.get('租賃負債-流動'))}、租賃負債-非流動{fmt(cur.get('租賃負債-非流動'))}。"
            + AI + "揭露內容評述；關係人租賃寫「擬行前抽閱租約」。"
            "句末加註頁碼標記（詳財務報告第＿頁）供承辦補填】"]},
        {"h": "（二）金融資產是否依IFRS9規定認列及衡量，並依IFRS7規定揭露。",
         "paras": [
            f"認列及衡量：主要金融資產部位——現金及約當現金{fmt(cur.get('現金及約當現金'))}、"
            f"應收款項合計{fmt(receivables(cur))}、其他應收款-關係人"
            f"{fmt(cur.get('其他應收款-關係人'))}。"
            + AI + "分類（攤銷後成本/FVOCI/FVTPL）依 facts.notes.FinancialInstruments 原文摘寫】",
            AI + "財報揭露：預期信用損失評估方法（facts.notes.FinancialInstruments 與 "
            "ScheduleOfTradeAndOtherReceivables）、帳齡分布（facts.age_distribution 有數字）、"
            "信用風險揭露評述。句末加註頁碼標記（詳財務報告第＿頁）供承辦補填】",
            AI + "公允價值衡量：facts.notes 內名稱含 FairValue 之節（如有）為公司公允價值"
            "揭露原文——請評述：各層級（第1/2/3等級）部位與金額、第3等級之評價技術與"
            "重要不可觀察輸入值是否揭露、層級間有無移轉；未上市（櫃）股票等無活絡市場"
            "報價者之評價假設合理性屬公司始能說明事項，標「擬行前請公司說明」。"
            "facts.notes 無該節時寫「公允價值層級及評價技術之揭露，擬行前核閱財報附註確認」，"
            "不得憑記憶補寫】"]},
        {"h": "（三）客戶合約之收入是否依IFRS15規定認列及衡量並為相關揭露。",
         "paras": [
            AI + "認列及衡量：依 facts.notes.RevenueRecognition 原文摘寫收入類別與認列時點"
            "（某一時點／隨時間逐步）】",
            f"財報揭露：{roc}年度營業收入{fmt(cur.get('營業收入'))}；期末合約負債"
            f"{fmt(cur.get('合約負債-流動'))}（{roc - 1}年度{fmt(prev.get('合約負債-流動'))}）。"
            + AI + "合約負債性質與變動方向是否與業務模式一致。"
            "句末加註頁碼標記（詳財務報告第＿頁）供承辦補填】"]},
        {"h": "（四）對被投資公司持股未逾50%且為單一最大股東者，是否依IFRS10第B38~B50段評估權力並揭露重大判斷。",
         "paras": [
            AI + "被投資公司清單見 facts.investees（名稱/持股/帳面/損益），權益法政策見 "
            "facts.notes.InvestmentsInAssociates。評估三要件、"
            "重大影響力判斷（董事席次、綜合持股）；持股低仍採權益法者寫「佐證文件擬行前調閱」，"
            "並評估集團有無構成控制而應納入他方合併個體。相關判斷之揭露處"
            "加註頁碼標記（詳財務報告第＿頁）供承辦補填】"]},
      ]},
      {"title": "四、會計主管審閱情形", "body": [
        {"paras": [
            AI + "財報查不到會計主管資格資訊時，寫「會計主管是否符合『發行人證券商證券交易所"
            "會計主管資格條件及專業進修辦法』之資格條件及持續進修規定，行前請至公開資訊觀測站"
            "查證並向公司調閱進修時數證明」；有公司回覆才寫符合情形】"]}]},
      {"title": "五、資金貸與及背書保證審閱情形", "body": [
        {"h": "1.執行面：", "paras": [
            _loans_para(loans, equity, roc),
            _endos_para(endos, equity, roc),
        ]},
        {"h": "2.訂定面：", "paras": [
            AI + "使用者有提供公司作業程序 PDF：核對內規限額（含持股90%以上公司間背書以淨值"
            "10%為限等特別條款）與執行面數字；未提供：寫「公司『資金貸與及背書保證作業程序』"
            "之限額規定與本期執行之相符性，擬行前調閱作業程序核對」】"]},
      ]},
      {"title": "附註：資料來源與限制", "body": [
        {"paras": [
            f"本檢查表依公開資訊觀測站{name}{roc}年度及{roc - 1}年度財務報告XBRL申報資料"
            "分析產出。" + AI + "逐一列實際引用之其他來源（同業平均、公司作業程序、母公司"
            "合併報告、前次實審……有用才列）】"
            "公司尚未回覆前，凡屬公司始能說明之事項均已標明「擬行前查證／請公司說明」。"
            "財務重點專區、財報重編紀錄、裁罰紀錄及董監持股質押未及查詢，行前請至公開資訊"
            "觀測站確認。本表僅依公開財報訊號分析，判斷責任在檢查員。"]}]},
    ]

    title = f"公開發行{name}股份有限公司{roc}年度財務報告公告檢查表—管區意見"
    if name.endswith("股份有限公司"):
        title = f"公開發行{name}{roc}年度財務報告公告檢查表—管區意見"
    # --- 複核表（財務報告實質審閱案件複核表；體例照過去實審樣本，另存一份 docx） ---
    cover = {
        "保存期限": "5年", "檔號": "",
        "公司編號": co, "公司名稱": name,
        "財務報告年度期別": f"{roc}年度",
        "公司背景介紹": AI + "一至三句：公司設立／公開發行時間與主要營業項目。僅可依 "
                       "facts 與財報附註既有資訊撰寫；查不到的（如設立日期）寫"
                       "「（設立及公開發行日期行前請至公開資訊觀測站基本資料查填）」，不得編造】",
        "風險事項": AI + "填入採認之風險事項科目名稱（與 sections 一、一致；無則寫「無」）】",
        "理由及因應措施": "詳檢查表一",
        "所屬產業趨勢": "詳檢查表二、個別資料庫",
        "加強查核重點": "公司及會計師填報之案件檢查表有無異常項目",
        "複核意見1": AI + "檢查表複核一句話；無異常時用「尚無重大異常」】",
        "檢視異常事項": ["檢視資料庫內容有無異常事項(管區意見檢查表)",
                     "檢視其它異常項目",
                     "針對異常事項之揭露、評價等會計處理進行查核，經電（函）請公司或會計師"
                     "說明後是否無異常"],
        "複核意見2": AI + "一句話；無異常時用「經核尚無發現重大異常」】",
        "結論及擬辦": AI + "無異常時用「尚無發現重大異常，文擬陳閱後存查，當否？謹請核示。」】",
    }

    return {
        "title": title,
        "cover": cover,
        "groups": groups,
        "footnotes": [
            "備註：填寫個別資料庫資料應逐項說明填寫「是」與「否」的合理性",
            "註1：加強審查關係人交易：關係人間重大資產之買賣或重大委託交易等，應瞭解其決策過程及相關帳務處理，以評估交易目的之合理性、會計處理及財務報告揭露之允當性，暨該交易是否符合公司內部控制制度及「公開發行公司取得或處分資產處理準則」相關規定。如查有涉及違反「公開發行公司取得或處分資產處理準則」相關規定情事，應移請證券發行組卓辦。",
        ],
        "sections": sections,
    }


def _loans_para(loans, equity, roc):
    if not loans:
        return "(1)資金貸與：期末無資金貸與他人餘額，無此情事。"
    rows = []
    for t in loans:
        rows.append(f"{t.get('Counterparty1', '?')}{fmt(t.get('EndingBalance1'))}"
                    f"（本期最高{fmt(t.get('MaximumBalanceForThePeriod1'))}、"
                    f"利率{t.get('RangeOfInterestRates', '?')}、"
                    f"性質{t.get('NatureOfLoans', '?')}）")
    lim = loans[0].get("LimitOfTotalLoanAmount1")
    total = sum(t.get("EndingBalance1") or 0 for t in loans)
    over = "超逾申報限額，應深入查明" if (lim and total > lim) else \
           f"總額{fmt(total)}在申報之總限額{fmt(lim)}內"
    ratio = f"，佔淨值{round(total / equity * 100, 1)}%" if equity else ""
    return ("(1)資金貸與（詳財報附表）：期末餘額——" + "、".join(rows)
            + f"。{over}{ratio}。" + AI + "貸與原因與集團資金規劃屬公司才答得出者寫"
            "「擬行前查證」；備抵呆帳與逾期情形依 tuples 數字評述】")


def _endos_para(endos, equity, roc):
    if not endos:
        return "(2)背書保證：期末無背書保證餘額，無此情事。"
    rows = []
    for t in endos:
        ratio = t.get("RatioOfAccumulatedEndorsementGuaranteeAmountToNetAssetOfTheCompany"
                      "PerLatestFinancialStatements", "?")
        lim_one = t.get("LimitOnEndorsementGuaranteeAmountProvidedToIndividualCounterparty")
        rows.append(f"為{t.get('NameOfTheCompany', '?')}（{t.get('Relationship1', '?')}）"
                    f"背書{fmt(t.get('EndingBalance2'))}，佔淨值{ratio}%，"
                    f"申報之單一對象限額{fmt(lim_one)}、"
                    f"總限額{fmt(t.get('LimitOfTotalGuaranteeEndorsementAmount'))}")
    anomalies = []
    for t in endos:
        eb, act = t.get("EndingBalance2"), t.get("ActualAmountProvided")
        if eb and act and act > eb * 3:
            anomalies.append(f"申報附表「實際動支金額」{fmt(act)}與期末背書餘額{fmt(eb)}"
                             "顯不相當，疑係申報欄位錯置（或屬聯貸案總額資訊），"
                             "擬行前向公司查明並要求更正申報")
    return ("(2)背書保證（詳財報附表）：" + "；".join(rows) + "。"
            + ("另查" + "；".join(anomalies) + "。" if anomalies else "")
            + AI + "背書原因屬公司才答得出者寫「擬行前查證」；如可取得母公司合併報告"
            "pretrip，核對集團持股（90%以上未達100%者注意內規淨值10%限額條款）】")


# ---------- 主流程 ----------

def main():
    ap = argparse.ArgumentParser(description="pretrip → 實審中介 JSON（inquiry + review_content）")
    ap.add_argument("--co", required=True)
    ap.add_argument("--year", required=True, type=int, help="年度（民國或西元皆可）")
    ap.add_argument("--data-dir", default="data")
    ap.add_argument("--outdir", default="out")
    ap.add_argument("--peer-avg", default="", help="同業平均 JSON（選填，格式見檔頭）")
    a = ap.parse_args()

    y = to_ad(a.year)
    roc = to_roc(y)
    pretrips = load_pretrips(a.data_dir, a.co, y)
    if y not in pretrips:
        print(f"找不到 {a.co}_{y}Q4_pretrip.json（--data-dir {a.data_dir}）。"
              f"請先把 {{\"co\": \"{a.co}\", \"year\": {y}, \"season\": 4}} 加進 "
              "data_requests.json 並 push（或本機跑 scripts/fetch_requests.py）。")
        sys.exit(2)

    main_pt = pretrips[y]
    meta, audit = main_pt["meta"], main_pt["audit"]
    tuples = main_pt.get("tuples", {})
    red_flags = main_pt.get("red_flags", [])
    series = merge_series(pretrips)

    peer = None
    if a.peer_avg:
        with open(a.peer_avg, encoding="utf-8") as f:
            peer = json.load(f)

    ratios = company_ratios(series, y)
    ratios_prev = company_ratios(series, y - 1)

    # 近 6 年表與成長率表
    years6 = list(range(y - 5, y + 1))
    six_year, growth = {}, {}
    for it in ["營業收入", "營業毛利", "營業損益", "稅前損益", "本期淨利", "每股盈餘(元)",
               "資產總額", "負債總額", "淨值"]:
        six_year[it] = {to_roc(yy): series.get(yy, {}).get(it) for yy in years6}
    for it in ["營業收入", "營業毛利", "營業損益", "稅前損益", "本期淨利"]:
        growth[it] = {to_roc(yy): pct(series.get(yy, {}).get(it),
                                      series.get(yy - 1, {}).get(it)) for yy in years6}

    investees = [
        {"name": t.get("CompanyNameOfTheInvestee") or t.get("NameOfInvestee"),
         "detail": t}
        for t in (tuples.get("NamesLocationsAndRelatedInformationOfInvesteesOverWhich"
                             "TheCompanyExercisesSignificantInfluence") or [])]

    facts = {
        "co": a.co, "roc_year": roc, "ad_year": y,
        "company": meta.get("CompanyChineseName"),
        "report_type": meta.get("ReportType"), "industry": meta.get("IndustrySector"),
        "audit": {k: audit.get(k) for k in
                  ("firm", "cpa", "report_date", "report_kind", "opinion", "flags")},
        "current": series.get(y, {}), "prior": series.get(y - 1, {}),
        "changes_pct": {it: pct(series.get(y, {}).get(it), series.get(y - 1, {}).get(it))
                        for it in list(METRICS_FLOW) + list(METRICS_STOCK)},
        "ratios": ratios, "ratios_prior": ratios_prev,
        "asset_mix": asset_mix(series.get(y, {})),
        "peer_avg": peer or "（未提供，Excel 相應欄留白待貼）",
        "six_year": six_year, "growth": growth,
        "loans": tuples.get("LoansToOthers") or [],
        "endorsements": tuples.get("EndorsementGuaranteeProvidedToOthers") or [],
        "age_distribution": tuples.get("AgeDistributionAndAmount") or [],
        "related_party_amounts": tuples.get(
            "FinancialStatementAccountAndCategoriesOfRelatedPartiesAndAmount") or [],
        "investees": investees,
        "red_flags": red_flags,
        # 財報附註原文（會計政策）——「認列及衡量」段須據此摘寫，不得憑印象編寫
        "notes": main_pt.get("notes_text", {}),
        "data_years_available": sorted(to_roc(k) for k in series),
    }

    # 差異說明各節數字列與五表版面資料（版面依過去實審實際樣本）
    tuples_prev = pretrips[y - 1].get("tuples") if (y - 1) in pretrips else None
    diff = build_diff_sections(series, y, ratios, ratios_prev, peer, tuples, tuples_prev,
                               report_category=meta.get("ReportCategory"))
    fin_items = {
        "營業收入淨額": lambda v: v.get("營業收入"),
        "營業毛利": lambda v: v.get("營業毛利"),
        "營業損益": lambda v: v.get("營業損益"),
        "稅前損益": lambda v: v.get("稅前損益"),
        "應收款項淨額(全部)": receivables,
        "存貨淨額": lambda v: v.get("存貨"),
        "營業活動現金流量": lambda v: v.get("營業活動現金流量"),
    }
    fin_data = {nm: {to_roc(yy): fn(series.get(yy, {})) for yy in years6}
                for nm, fn in fin_items.items()}
    ratios_by_year = {to_roc(yy): company_ratios(series, yy) for yy in years6}
    growth4 = {("營收成長率" if it == "營業收入" else f"{it}成長率"):
               {to_roc(yy): pct(series.get(yy, {}).get(it),
                                series.get(yy - 1, {}).get(it)) for yy in years6}
               for it in PL_ITEMS}

    qmap = {}
    questions = build_questions(series, y, ratios, ratios_prev, peer, tuples, meta, qmap)
    analysis = build_analysis(series, y, ratios, ratios_prev, peer, qmap)

    # 資料充足度：成長率「兩期比較」需 y、y-1、y-2 三年數字，
    # 亦即至少要有 y 與 y-2（或 y-1）兩份年報。缺了要明講，不能靜靜算不出來。
    have = sorted(series)
    need = [y, y - 1, y - 2]
    lack = [to_roc(n) for n in need if n not in have]
    coverage = {
        "資料年度": [to_roc(v) for v in have],
        "成長率兩期比較所需年度": [to_roc(n) for n in need],
        "缺少年度": lack,
        "足以比較兩期成長率": not lack,
        "說明": ("完整" if not lack else
               f"缺 {lack} 年度資料（民國），成長率兩期比較無法計算——請補抓："
               f"python scripts/fetch_archive.py --only {a.co} "
               f"--years {','.join(str(n) for n in lack)}"),
    }
    if lack:
        print(f"⚠ 資料不足：缺 {lack} 年度，成長率兩期比較將顯示「－（缺資料）」。{coverage['說明']}")
    inquiry = {
        "meta": {"co": a.co, "roc_year": roc, "company": meta.get("CompanyChineseName"),
                 "industry": meta.get("IndustrySector"),
                 "report_category": meta.get("ReportCategory"),
                 "audit_firm": "、".join(audit.get("firm") or []),
                 "cpa": "、".join(audit.get("cpa") or []),
                 "opinion": (audit.get("opinion") or {}).get("label"),
                 "report_date": audit.get("report_date")},
        "six_year": six_year, "growth": growth,
        "ratios": {"公司": ratios,
                   "上市櫃同業平均": (peer or {}).get("上市櫃同業平均"),
                   "所有同業平均": (peer or {}).get("所有同業平均")},
        "fin_data": fin_data,
        "ratios_by_year": ratios_by_year,
        "growth4": growth4,
        "diff": diff,
        "capital": series.get(y, {}).get("普通股股本"),
        "questions": questions,
        "analysis": analysis,
        "data_coverage": coverage,
    }

    draft = build_checklist_draft(a.co, roc, meta, audit, series, y, ratios,
                                  ratios_prev, peer, tuples, red_flags)
    review = {"meta": inquiry["meta"], "facts": facts, "draft": draft}

    os.makedirs(a.outdir, exist_ok=True)
    p1 = os.path.join(a.outdir, f"{a.co}_{roc}_inquiry.json")
    p2 = os.path.join(a.outdir, f"{a.co}_{roc}_review_content.json")
    with open(p1, "w", encoding="utf-8") as f:
        json.dump(inquiry, f, ensure_ascii=False, indent=1)
    with open(p2, "w", encoding="utf-8") as f:
        json.dump(review, f, ensure_ascii=False, indent=1)

    # 完整性不變式：每個達標科目都必須有對應題號或說明，否則就是分析漏網
    orphan = [r["項目"] for r in analysis
              if r["是否達標"].startswith("★") and not r["對應題號"] and not r["備註"]]
    if orphan:
        print(f"⚠ 下列科目達分析門檻卻未出題亦無說明，請檢查出題規則：{orphan}")
    hit = sum(1 for r in analysis if r["是否達標"].startswith("★"))
    print(f"✔ {p1}（題目 {len(questions)} 題；分析表 {len(analysis)} 個科目，達標 {hit} 項）")
    print(f"✔ {p2}（風險候選 {len(draft['sections'][0]['body'])}、"
          f"資料年度 {facts['data_years_available']}）")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
