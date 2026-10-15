#!/usr/bin/env python3
"""蝦皮台灣訂單試算：這單賺多少、每項費用扣多少、方案一和方案二哪個划算。

用法（單筆）：
  python3 <skill資料夾>/scripts/calc.py --price 500 --cost 200
  python3 <skill資料夾>/scripts/calc.py --price 50 --cost 20 --plan 方案二 --commission-rate 5%

用法（批次，給 Google 試算表或 Excel 匯出的 CSV）：
  python3 <skill資料夾>/scripts/calc.py --csv 訂單.csv --out 訂單_試算結果.csv
  算完會印出整批的合計，以及全部用方案一、全部用方案二各要多少，用來判斷該不該換方案。

只用 Python 標準函式庫，Claude 和 Codex 都能直接跑。
費率都在 references/fees.json，蝦皮改規則只改那個檔。
"""
import argparse
import csv
import json
import math
from decimal import InvalidOperation
import re
import sys
import unicodedata
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

FEES_PATH = Path(__file__).resolve().parent.parent / "references" / "fees.json"

# 欄位名稱 -> 內部參數名。中英文都接受，標題裡括號和「元」會先去掉再比對。
COLUMN_ALIASES = {
    "price": ["售價", "單價", "價格", "price"],
    "qty": ["數量", "qty", "quantity"],
    "cost": ["成本", "商品成本", "進價", "cost"],
    "buyer_shipping": ["買家運費", "買家付運費", "buyer_shipping"],
    "seller_shipping": ["賣家運費", "賣家負擔運費", "seller_shipping"],
    "seller_discount": ["賣家折扣", "折扣", "seller_discount"],
    "shopee_discount": ["蝦皮折抵", "蝦皮折扣", "shopee_discount"],
    "commission_rate": ["成交手續費率", "手續費率", "commission_rate"],
    "plan": ["方案", "免運方案", "免運專案", "plan"],
    "campaign": ["檔期", "促銷檔期", "campaign"],
    "mall": ["商城", "mall"],
    "preorder": ["預購", "較長備貨", "preorder"],
    "coin_cashback": ["蝦幣回饋", "coin_cashback"],
    "ad_cost": ["廣告費", "ad_cost"],
    "other_cost": ["其他成本", "包材", "other_cost"],
    "invoice": ["開發票", "invoice"],
}
ALIAS_TO_KEY = {a.lower(): k for k, names in COLUMN_ALIASES.items() for a in names}

TRUE_WORDS = {"1", "true", "yes", "y", "是", "有", "v", "要", "開", "o"}
FALSE_WORDS = {"0", "false", "no", "n", "否", "無", "沒有", "不要", "不開", "x", ""}

PLAN_WORDS = {
    "plan1": {"plan1", "p1", "1", "一", "方案1", "方案一", "6%", "6", "0.06"},
    "plan2": {"plan2", "p2", "2", "二", "方案2", "方案二", "60", "60元"},
}


PLAN_NAME = {"plan1": "方案一", "plan2": "方案二"}


class InputError(ValueError):
    """使用者輸入有問題，給白話訊息，不丟 Traceback。"""


def r(x):
    """四捨五入到元（不是銀行家捨入）。"""
    return int(Decimal(str(x)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def rm(base, rate):
    """金額 × 費率，用十進位精確計算後四捨五入到元，避免 12.5 被浮點誤差算成 12.4999。"""
    d = Decimal(str(round(base, 6))) * Decimal(str(round(rate, 8)))
    return int(d.quantize(Decimal("1"), rounding=ROUND_HALF_UP))


REQUIRED_FEES = [
    "fees_as_of", "commission_rate_default", "commission_rate_range",
    "shipping_program", "payment_fee_rate", "campaign_surcharge",
    "coin_cashback_rates", "invoice_fee_per_order", "invoice_default",
    "commission_cap_per_item", "commission_cap_per_order",
    "preorder_rate", "preorder_cap_per_item", "coin_program_discount",
]


def load_fees(path=FEES_PATH):
    """讀 fees.json 並檢查欄位，改壞了給白話訊息。"""
    try:
        with open(path, encoding="utf-8") as f:
            fees = json.load(f)
    except FileNotFoundError:
        raise InputError(f"找不到費率設定檔 {path}")
    except json.JSONDecodeError as e:
        raise InputError(f"費率設定檔 fees.json 格式打錯了（第 {e.lineno} 行附近），"
                         "常見原因是少了逗號或引號")
    missing = [k for k in REQUIRED_FEES if k not in fees]
    sp = fees.get("shipping_program", {})
    missing += [f"shipping_program.{k}" for k in ("plan1_rate", "plan2_fixed", "plan2_small_order_max",
                                                  "plan2_small_order_rate", "default_plan")
                if k not in sp]
    if missing:
        raise InputError("fees.json 少了這些欄位：" + "、".join(missing))
    cs = fees.get("campaign_surcharge")
    coins = fees.get("coin_cashback_rates")
    try:
        ok = (isinstance(cs, dict) and 0 <= float(cs["rate"]) < 1 and 0 <= float(cs["rate_mall"]) < 1
              and isinstance(coins, list) and all(0 < float(c) < 1 for c in coins)
              and len(fees["commission_rate_range"]) == 2
              and 0 < float(sp["plan1_rate"]) < 1 and float(sp["plan2_fixed"]) > 0
              and 0 < float(sp["plan2_small_order_rate"]) < 1 and float(sp["plan2_small_order_max"]) >= 0
              and 0 < float(fees["payment_fee_rate"]) < 1
              and sp["default_plan"] in ("plan1", "plan2"))
    except (TypeError, ValueError, KeyError):
        ok = False
    if not ok:
        raise InputError("fees.json 的數字不合理：費率要填小數（例如 6% 填 0.06），"
                         "方案二金額要大於 0，小額訂單門檻不能是負數，default_plan 要是 plan1 或 plan2，"
                         "檔期加收要有 rate 和 rate_mall，蝦幣回饋要是費率清單")
    return fees


def clean_text(v):
    """去空白、全形轉半形。"""
    if v is None:
        return ""
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(v)))


def parse_num(v, name, default=None):
    """把 'NT$1,200'、' 500 元'、'5.5%' 之類轉成數字。回傳 (值, 是否帶%)。"""
    if v is None or (isinstance(v, str) and v.strip() == ""):
        if default is None:
            raise InputError(f"{name}沒有填")
        return default, False
    if isinstance(v, bool):
        raise InputError(f"{name}要填數字")
    if isinstance(v, (int, float)):
        x, pct = float(v), False
    else:
        s = clean_text(v).replace(",", "").replace("，", "")
        s = re.sub(r"^(NT\$|NTD|TWD|\$|＄)", "", s, flags=re.I).replace("元", "")
        pct = s.endswith("%") or s.endswith("％")
        s = s.rstrip("%％")
        try:
            x = float(s)
        except ValueError:
            raise InputError(f"{name}「{v}」看不懂，要填數字")
    if math.isnan(x) or math.isinf(x):
        raise InputError(f"{name}「{v}」不是有效數字")
    return x, pct


def parse_rate(v, name, lo, hi):
    """費率可填 5.5%、5.5 或 0.055。超出合理範圍就報錯，不默默算。"""
    x, pct = parse_num(v, name)
    if pct or x >= 1:
        x = x / 100
    if not (lo <= x <= hi):
        raise InputError(
            f"{name}「{v}」換算成 {x:.2%}，不在合理範圍 {lo:.1%}–{hi:.1%}。"
            f"請填像 5.5% 或 0.055 這樣的值"
        )
    return x


def parse_bool(v, name, default=False):
    if v is None:
        return default
    if isinstance(v, bool):
        return v
    s = clean_text(v).lower()
    if s == "":
        return default
    if s in TRUE_WORDS:
        return True
    if s in FALSE_WORDS:
        return False
    raise InputError(f"{name}「{v}」看不懂，請填 是／否")


def _match_plan(s):
    for plan, words in PLAN_WORDS.items():
        if s in words:
            return plan
    return None


def parse_plan(v, default_plan):
    raw = clean_text(v).lower()
    if raw == "":
        return default_plan, True
    inner = re.findall(r"[\(（](.*?)[\)）]", raw)  # 括號內容，例如「方案一（6%）」的 6%
    s = re.sub(r"[\(（].*?[\)）]", "", raw)
    plan = _match_plan(s)
    if plan is None:
        raise InputError(f"方案「{v}」看不懂，請填 方案一 或 方案二")
    for i in inner:
        other = _match_plan(i)
        if other and other != plan:
            raise InputError(f"方案「{v}」前後矛盾，請確認是方案一（6%）還是方案二（60 元）")
    return plan, False


def plan_fees(goods, sp):
    """免運專案費用（官方 18733）。
    方案一：商品成交單價 × 數量 × 6%（不含運費）。
    方案二：訂單金額 99 元以下收 6%，100 元以上每單固定 60 元（2026/7/1 起）。"""
    plan1 = rm(goods, sp["plan1_rate"])
    if goods < sp["plan2_small_order_max"] + 1:  # 99 元（含）以下
        plan2 = rm(goods, sp["plan2_small_order_rate"])
    else:
        plan2 = sp["plan2_fixed"]
    return plan1, plan2


def calc(o, fees):
    """o: 輸入的 dict。回傳每項明細與結論。輸入有問題丟 InputError。"""
    notes = []
    sp = fees["shipping_program"]
    lo, hi = fees["commission_rate_range"]

    price, _ = parse_num(o.get("price"), "售價")
    if price <= 0:
        raise InputError("售價要大於 0")
    if price > 10_000_000:
        raise InputError("售價超過一千萬，請確認數字")
    qty_raw, _ = parse_num(o.get("qty"), "數量", default=1)
    if qty_raw < 1 or qty_raw != int(qty_raw):
        raise InputError(f"數量「{o.get('qty')}」要是 1 以上的整數")
    if qty_raw > 100_000:
        raise InputError("數量超過十萬，請確認數字")
    qty = int(qty_raw)

    def money(key, label):
        x, _ = parse_num(o.get(key), label, default=0)
        if x < 0:
            raise InputError(f"{label}不能是負數")
        return x

    cost = money("cost", "成本")
    if o.get("cost") in (None, ""):
        notes.append("沒填成本，淨利沒有扣商品成本，會比實際高")
    buyer_ship = money("buyer_shipping", "買家運費")
    seller_ship = money("seller_shipping", "賣家運費")
    discount = money("seller_discount", "賣家折扣")
    shopee_disc = money("shopee_discount", "蝦皮折抵")
    ad = money("ad_cost", "廣告費")
    other = money("other_cost", "其他成本")

    goods = price * qty - discount  # 商品金額（扣賣家折扣）
    if goods <= 0:
        raise InputError("賣家折扣大於或等於商品金額，請確認數字")

    if o.get("commission_rate") in (None, ""):
        rate = fees["commission_rate_default"]
        notes.append(f"沒填成交手續費率，先用 {rate:.1%} 假設，請改成你分類的費率")
    else:
        rate = parse_rate(o.get("commission_rate"), "成交手續費率", lo, hi)

    coin = 0.0
    if o.get("coin_cashback") not in (None, "", 0, 0.0):
        x, pct = parse_num(o.get("coin_cashback"), "蝦幣回饋")
        if x != 0:
            if pct or x >= 1:
                x = x / 100
            allowed = fees["coin_cashback_rates"]
            if not any(abs(x - a) < 1e-9 for a in allowed):
                raise InputError(
                    f"蝦幣回饋「{o.get('coin_cashback')}」只能是 "
                    + " 或 ".join(f"{a:.1%}" for a in allowed) + "，沒參加填 0"
                )
            coin = x

    campaign = parse_bool(o.get("campaign"), "檔期")
    mall = parse_bool(o.get("mall"), "商城")
    preorder = parse_bool(o.get("preorder"), "預購")
    invoice = parse_bool(o.get("invoice"), "開發票", default=fees["invoice_default"])

    plan, defaulted = parse_plan(o.get("plan"), sp["default_plan"])
    if defaulted:
        notes.append("沒說是哪個方案，照平台預設算方案二；請到賣家中心「蝦皮活動 > 已報名」確認自己在哪個方案")

    plan1_fee, plan2_fee = plan_fees(goods, sp)
    breakeven = sp["plan2_fixed"] / sp["plan1_rate"]

    # 成交手續費（官方）：單件售出金額 × 數量 ×（基本費率＋檔期費率）
    #   售出金額 = 原價 − 賣家折扣；運費不算；一般賣家單件、單筆訂單都以 35,000 元為上限
    unit_sold = goods / qty
    cap_item = fees["commission_cap_per_item"]
    cap_order = fees["commission_cap_per_order"]
    comm_base = min(unit_sold, cap_item) * qty if not mall else goods
    if not mall and comm_base > cap_order:
        comm_base = cap_order
    if comm_base < goods:
        notes.append(f"一般賣家成交手續費最多只算到 {cap_order:,.0f} 元，超過的部分不收")

    cs = fees["campaign_surcharge"]
    surcharge = 0.0
    if campaign and coin == 0:
        surcharge = cs["rate_mall"] if mall else cs["rate"]
    elif campaign and coin > 0:
        notes.append("有參加蝦幣回饋，檔期不另外加收成交手續費")
    if mall and not campaign:
        notes.append("商城只影響檔期加收的費率，這單不是檔期，所以沒差")
    comm_label = "成交手續費" + (f"（含檔期加收 {surcharge:.0%}）" if surcharge else "")

    # 金流費（官方）：平台交易金額 × 2.5%，平台交易金額 = 實際成交金額（含運費）+ 蝦皮出的折扣
    #   goods 是蝦皮折抵前的金額，所以蝦皮券、蝦幣不另外加；運費要算進去
    pay_base = goods + buyer_ship
    if shopee_disc:
        notes.append("蝦皮全額出錢的券和蝦幣，費用照折抵前的金額算，不影響你被扣多少；"
                     "賣家要分攤的券請填在賣家折扣")

    items = [
        (comm_label, rm(comm_base, rate + surcharge)),
        ("免運專案（方案一）" if plan == "plan1" else "免運專案（方案二）",
         plan1_fee if plan == "plan1" else plan2_fee),
        ("金流服務費", rm(pay_base, fees["payment_fee_rate"])),
    ]
    if coin > 0:
        coin_rate = coin
        disc = fees["coin_program_discount"]
        if abs(coin - disc["applies_to_rate"]) < 1e-9:
            coin_rate = coin - disc["rate"]
            notes.append(f"同時參加免運專案和 10% 蝦幣回饋，蝦幣回饋費官方減免 {disc['rate']:.1%}，"
                         f"以 {coin_rate:.1%} 計")
        items.append(("蝦幣回饋", rm(goods, coin_rate)))
    if preorder:
        pre_base = min(unit_sold, fees["preorder_cap_per_item"]) * qty
        items.append(("預購訂單服務費", rm(pre_base, fees["preorder_rate"])))
    if invoice:
        items.append(("發票服務費", fees["invoice_fee_per_order"]))
        notes.append(f"已算發票服務費 {fees['invoice_fee_per_order']:g} 元；沒用蝦皮代開發票可以關掉")

    platform_total = sum(v for _, v in items)
    payout = goods - platform_total  # 撥款（平台扣完的錢）
    product_cost = cost * qty
    profit = payout - product_cost - seller_ship - ad - other
    margin = profit / goods

    cheaper = "plan1" if plan1_fee < plan2_fee else ("plan2" if plan2_fee < plan1_fee else "same")
    if cheaper != "same" and cheaper != plan:
        diff = abs(plan1_fee - plan2_fee)
        notes.append(f"這單用{'方案一' if cheaper == 'plan1' else '方案二'}可以少付 {diff:g} 元"
                     "（方案是整個賣場一起選，要看全部訂單，不是這一單）")
    if profit < 0:
        notes.append("這單虧錢")
    elif margin < 0.3:
        notes.append("毛利率低於 30%，空間很小")
    if seller_ship:
        notes.append("賣家負擔運費當成本扣；實際是否從撥款扣，看你的運費設定")

    return {
        "商品金額": round(goods, 2),
        "平台扣款明細": items,
        "平台扣款合計": round(platform_total, 2),
        "撥款（實拿）": round(payout, 2),
        "商品成本": round(product_cost, 2),
        "賣家負擔運費": seller_ship,
        "廣告費": ad,
        "其他成本": other,
        "淨利": round(profit, 2),
        "毛利率": margin,
        "方案一費用": plan1_fee,
        "方案二費用": plan2_fee,
        "打平訂單金額": breakeven,
        "使用方案": plan,
        "提醒": notes,
        "條件": {
            "成交手續費率": f"{rate:.2%}",
            "方案": PLAN_NAME.get(plan, plan),
            "檔期": campaign, "商城": mall, "預購": preorder,
            "蝦幣回饋": f"{coin:.1%}" if coin else "沒參加",
            "發票服務費": invoice,
        },
        "費率版本": fees["fees_as_of"],
    }


def to_json(res):
    """給 --json 用：名稱對金額、百分比、中文方案名。"""
    out = dict(res)
    out["平台扣款明細"] = {n: v for n, v in res["平台扣款明細"]}
    out["毛利率"] = f"{res['毛利率']:.1%}"
    out["使用方案"] = PLAN_NAME.get(res["使用方案"], res["使用方案"])
    return out


def safe_calc(o, fees):
    try:
        return calc(o, fees)
    except (OverflowError, InvalidOperation):
        raise InputError("數字太大，算不出來，請確認輸入")


def fmt(x):
    """金額顯示：千分位、不用科學記號，有小數才顯示一位。"""
    x = float(x)
    return f"{x:,.0f}" if x == int(x) else f"{x:,.1f}"


def print_report(res):
    print(f"商品金額：{fmt(res['商品金額'])} 元")
    print("平台扣款：")
    for name, v in res["平台扣款明細"]:
        print(f"  - {name}：{fmt(v)} 元")
    print(f"  合計：{fmt(res['平台扣款合計'])} 元")
    print(f"撥款（實拿）：{fmt(res['撥款（實拿）'])} 元")
    print(f"扣掉商品成本 {fmt(res['商品成本'])}、賣家運費 {fmt(res['賣家負擔運費'])}、"
          f"廣告 {fmt(res['廣告費'])}、其他 {fmt(res['其他成本'])} 後")
    print(f"淨利：{fmt(res['淨利'])} 元（毛利率 {res['毛利率']:.1%}）")
    print(f"方案比較：方案一 {fmt(res['方案一費用'])} 元 / 方案二 {fmt(res['方案二費用'])} 元"
          f"（訂單商品金額 100 元以下兩個一樣，{res['打平訂單金額']:,.0f} 元打平）")
    for n in res["提醒"]:
        print(f"提醒：{n}")
    print(f"費率版本：{res['費率版本']}，請到蝦皮賣家中心核對")


def normalize_header(h):
    h = clean_text(h)
    h = re.sub(r"[\(（\[【].*?[\)）\]】]", "", h)  # 去掉括號內容，例如「售價(元)」
    h = h.replace("元", "").lower()
    return ALIAS_TO_KEY.get(h, h)


def read_csv_rows(path):
    """依序試 UTF-8（含 BOM）和 Big5（台灣 Excel 預設），回傳 (rows, 編碼)。"""
    last_err = None
    for enc in ("utf-8-sig", "cp950"):
        try:
            with open(path, encoding=enc, newline="") as f:
                reader = csv.DictReader(f)
                rows = list(reader)
                return rows, enc, [h for h in (reader.fieldnames or []) if h is not None]
        except UnicodeDecodeError as e:
            last_err = e
    raise InputError(f"讀不懂這個 CSV 的編碼（{last_err}），請另存成 UTF-8 的 CSV")


def run_csv(in_path, out_path, fees):
    rows, enc, headers = read_csv_rows(in_path)
    if not headers:
        raise InputError("CSV 是空的")
    if not rows:
        raise InputError("CSV 只有標題列，沒有訂單資料")
    keys, seen = {}, {}
    for h in headers:
        k = normalize_header(h)
        if k in ALIAS_TO_KEY.values():
            if k in seen:
                raise InputError(f"「{seen[k]}」和「{h}」對到同一個欄位，請只留一欄（或先合併成一欄）")
            seen[k] = h
            keys[h] = k
    if "price" not in keys.values():
        raise InputError("CSV 找不到「售價」欄，第一列要是標題，例如：售價,數量,成本")

    # 只提醒「看起來跟計算有關」但名稱對不上的欄位，訂單編號、客人姓名這類不吵
    related = re.compile(r"費|運|成本|率|價|折|方案|數量|檔期|蝦幣|發票|商城")
    unused = [h for h in headers if h not in keys and related.search(clean_text(h))]
    unused_note = ("沒用到的欄位：" + "、".join(unused) + "（名稱對不上，請改成說明裡的欄位名）") if unused else ""
    blank_extra = {"試算_狀態": "空白", "試算_平台扣款合計": "", "試算_撥款": "", "試算_淨利": "",
                   "試算_毛利率": "", "試算_方案一費用": "", "試算_方案二費用": "",
                   "試算_扣款明細": "", "試算_提醒": ""}
    out_rows, ok, bad = [], 0, 0
    results = []
    for i, row in enumerate(rows, start=2):  # 第 1 列是標題
        values = [v for v in row.values() if isinstance(v, str)]
        if not any(v.strip() for v in values):
            # 空白列照樣輸出，貼回原表時才不會錯位
            out_rows.append({**{k: v for k, v in row.items() if k is not None}, **blank_extra})
            continue
        o = {keys[h]: v for h, v in row.items() if h in keys}
        try:
            extra_cells = [c for c in (row.get(None) or []) if str(c).strip()]
            if extra_cells:
                raise InputError("這列的格數比標題多，常見原因是金額有逗號沒加引號（例如 1,200 要寫成 \"1,200\" 或 1200）")
            res = safe_calc(o, fees)
            results.append(res)
            extra = {
                "試算_狀態": "OK",
                "試算_平台扣款合計": res["平台扣款合計"],
                "試算_撥款": res["撥款（實拿）"],
                "試算_淨利": res["淨利"],
                "試算_毛利率": f"{res['毛利率']:.1%}",
                "試算_方案一費用": res["方案一費用"],
                "試算_方案二費用": res["方案二費用"],
                "試算_扣款明細": "；".join(f"{n} {fmt(v)}" for n, v in res["平台扣款明細"]),
                "試算_提醒": "；".join(res["提醒"] + ([unused_note] if unused_note else [])),
            }
            ok += 1
        except InputError as e:
            extra = {"試算_狀態": "錯誤", "試算_平台扣款合計": "", "試算_撥款": "", "試算_淨利": "",
                     "試算_毛利率": "", "試算_方案一費用": "", "試算_方案二費用": "",
                     "試算_扣款明細": "", "試算_提醒": f"第 {i} 列：{e}"}
            bad += 1
        clean_row = {k: v for k, v in row.items() if k is not None}
        out_rows.append({**clean_row, **extra})

    if not out_rows:
        raise InputError("CSV 裡沒有可計算的列")
    fieldnames = list(out_rows[0].keys())
    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        w.writeheader()
        w.writerows(out_rows)
    if unused_note:
        print(unused_note)
    print(f"完成：{ok} 筆算好，{bad} 筆有問題（看「試算_提醒」欄），輸出到 {out_path}")
    if results:
        print_summary(summarize(results))
    print(f"讀檔編碼 {enc}，費率版本 {fees['fees_as_of']}，請到蝦皮賣家中心核對")
    return 0 if ok else 1


def summarize(results):
    """整批訂單的合計，以及全部用方案一、全部用方案二各要付多少。"""
    plan1 = sum(x["方案一費用"] for x in results)
    plan2 = sum(x["方案二費用"] for x in results)
    current = sum(x["方案一費用"] if x["使用方案"] == "plan1" else x["方案二費用"] for x in results)
    plans_used = {x["使用方案"] for x in results}
    buckets = {"99 元以下": 0, "100–999 元": 0, "1,000 元以上": 0}
    for x in results:
        g = x["商品金額"]
        key = "99 元以下" if g < 100 else ("100–999 元" if g < x["打平訂單金額"] else "1,000 元以上")
        buckets[key] += 1
    return {
        "筆數": len(results),
        "商品金額合計": round(sum(x["商品金額"] for x in results), 2),
        "平台扣款合計": round(sum(x["平台扣款合計"] for x in results), 2),
        "撥款合計": round(sum(x["撥款（實拿）"] for x in results), 2),
        "淨利合計": round(sum(x["淨利"] for x in results), 2),
        "虧錢筆數": sum(1 for x in results if x["淨利"] < 0),
        "全用方案一": plan1,
        "全用方案二": plan2,
        "目前方案費用": current,
        "目前方案": PLAN_NAME[plans_used.pop()] if len(plans_used) == 1 else "混用",
        "訂單金額分布": buckets,
    }


def print_summary(s):
    print(f"整批 {s['筆數']} 筆：商品金額 {fmt(s['商品金額合計'])} 元，平台扣 {fmt(s['平台扣款合計'])} 元，"
          f"撥款 {fmt(s['撥款合計'])} 元，淨利 {fmt(s['淨利合計'])} 元，虧錢 {s['虧錢筆數']} 筆")
    b = s["訂單金額分布"]
    print("訂單金額分布：" + "、".join(f"{k} {v} 筆" for k, v in b.items()))
    p1, p2 = s["全用方案一"], s["全用方案二"]
    print(f"免運專案：全部用方案一共 {fmt(p1)} 元，全部用方案二共 {fmt(p2)} 元"
          f"（目前{s['目前方案']}，付 {fmt(s['目前方案費用'])} 元）")
    if p1 == p2:
        print("方案建議：這批訂單兩個方案一樣")
    else:
        better = "方案一" if p1 < p2 else "方案二"
        save = s["目前方案費用"] - min(p1, p2)
        if s["目前方案"] == better or save <= 0:
            print(f"方案建議：{better}比較省，你已經在用了")
        else:
            print(f"方案建議：這批訂單用{better}可以少付 {fmt(save)} 元。"
                  "方案 30 天內可以換一次，建議拿一整個月的訂單來看再決定")


def main():
    p = argparse.ArgumentParser(description="蝦皮台灣單筆訂單試算")
    p.add_argument("--price", help="單價")
    p.add_argument("--qty", default=1, help="數量（1 以上整數）")
    p.add_argument("--cost", default=None, help="每件商品成本")
    p.add_argument("--buyer-shipping", default=0, help="買家付的運費")
    p.add_argument("--seller-shipping", default=0, help="賣家負擔的運費")
    p.add_argument("--seller-discount", default=0, help="賣家自己給的折扣（整筆）")
    p.add_argument("--shopee-discount", default=0, help="蝦皮券、蝦幣、蝦皮補貼的折抵（整筆）")
    p.add_argument("--commission-rate", help="成交手續費率，例如 5.5%、5.5 或 0.055")
    p.add_argument("--plan", help="免運專案：方案一 / 方案二")
    p.add_argument("--campaign", action="store_true", help="促銷檔期")
    p.add_argument("--mall", action="store_true", help="商城賣家")
    p.add_argument("--preorder", action="store_true", help="預購／較長備貨（備貨 2 個工作日以上）")
    p.add_argument("--coin-cashback", default=0, help="蝦幣回饋 1.5% 或 3%，沒參加填 0")
    p.add_argument("--ad-cost", default=0)
    p.add_argument("--other-cost", default=0, help="包材等其他成本")
    p.add_argument("--no-invoice", action="store_true", help="沒用蝦皮代開發票")
    p.add_argument("--csv", help="批次輸入 CSV")
    p.add_argument("--out", help="批次輸出 CSV（預設放在輸入檔旁邊）")
    p.add_argument("--json", action="store_true", help="輸出 JSON")
    p.add_argument("--fees", default=str(FEES_PATH), help="費率設定檔")
    a = p.parse_args()

    try:
        fees = load_fees(a.fees)
        if a.csv:
            in_path = Path(a.csv)
            if in_path.is_dir():
                raise InputError(f"{in_path} 是資料夾，請給 CSV 檔案")
            out = Path(a.out) if a.out else in_path.with_name(in_path.stem + "_試算結果.csv")
            if out.is_dir():
                out = out / (in_path.stem + "_試算結果.csv")
            if out.resolve() == in_path.resolve():
                raise InputError("輸出檔不能跟輸入檔同一個，不然原始資料會被蓋掉")
            return run_csv(in_path, out, fees)
        if a.price is None:
            raise InputError("需要 --price，或用 --csv 批次計算")
        o = {
            "price": a.price, "qty": a.qty, "cost": a.cost,
            "buyer_shipping": a.buyer_shipping, "seller_shipping": a.seller_shipping,
            "seller_discount": a.seller_discount, "shopee_discount": a.shopee_discount,
            "commission_rate": a.commission_rate, "plan": a.plan,
            "campaign": a.campaign, "mall": a.mall, "preorder": a.preorder, "coin_cashback": a.coin_cashback,
            "ad_cost": a.ad_cost, "other_cost": a.other_cost,
            "invoice": False if a.no_invoice else None,
        }
        res = safe_calc(o, fees)
    except InputError as e:
        print(f"輸入有問題：{e}", file=sys.stderr)
        return 2
    except FileNotFoundError as e:
        print(f"找不到檔案：{e.filename}", file=sys.stderr)
        return 2
    except PermissionError as e:
        print(f"沒辦法寫入或讀取 {e.filename}：檔案可能正被 Excel 開著，或沒有權限。"
              "關掉檔案或換個輸出位置再試", file=sys.stderr)
        return 2
    except OSError as e:
        print(f"檔案讀寫失敗：{e}", file=sys.stderr)
        return 2

    if a.json:
        print(json.dumps(to_json(res), ensure_ascii=False, indent=2))
    else:
        print_report(res)
    return 0


if __name__ == "__main__":
    sys.exit(main())
