#!/usr/bin/env python3
"""用蝦皮官方公告裡的範例驗證計算結果。改 fees.json 或程式之後都要跑一次。

用法：python3 tests/test_official_examples.py
全部通過會印「全部通過」，有錯會列出哪一題、應該是多少、算出多少，並以代碼 1 結束。
"""
import sys
from pathlib import Path

SKILL = Path(__file__).resolve().parent.parent / "skills" / "shopee-fee-calc"
sys.path.insert(0, str(SKILL / "scripts"))
import calc  # noqa: E402

FEES = calc.load_fees(SKILL / "references" / "fees.json")

# (說明, 輸入, 費用項目關鍵字, 應該是多少, 出處)
CASES = [
    ("售出 400、運費 60、費率 5%：成交手續費",
     {"price": 400, "buyer_shipping": 60, "commission_rate": "5%"}, "成交手續費", 20, "26470"),
    ("售出 40,000、費率 5%：單件上限 35,000",
     {"price": 40000, "commission_rate": "5%"}, "成交手續費", 1750, "26470"),
    ("售出 400、費率 5%、檔期沒參加蝦幣回饋：5%+2%",
     {"price": 400, "commission_rate": "5%", "campaign": True}, "成交手續費", 28, "26470"),
    ("售出 40,000、費率 5%、檔期沒參加蝦幣回饋",
     {"price": 40000, "commission_rate": "5%", "campaign": True}, "成交手續費", 2450, "26470"),
    ("售出 400、費率 5.5%",
     {"price": 400, "commission_rate": "5.5%"}, "成交手續費", 22, "25681"),
    ("售出 40,000、費率 5.5%",
     {"price": 40000, "commission_rate": "5.5%"}, "成交手續費", 1925, "25681"),
    ("售出 400、費率 5.5%、檔期：7.5%",
     {"price": 400, "commission_rate": "5.5%", "campaign": True}, "成交手續費", 30, "25681"),
    ("售出 40,000、費率 5.5%、檔期",
     {"price": 40000, "commission_rate": "5.5%", "campaign": True}, "成交手續費", 2625, "25681"),
    ("售出 400、運費 60：金流費 11.5 記 12",
     {"price": 400, "buyer_shipping": 60}, "金流服務費", 12, "26470"),
    ("售出 21,915、運費 60：金流費 549.3 記 549",
     {"price": 21915, "buyer_shipping": 60}, "金流服務費", 549, "26470"),
    ("預購商品單價 100：預購訂單服務費 3",
     {"price": 100, "preorder": True}, "預購訂單服務費", 3, "25681"),
    ("沒選方案：預設方案二每單 60",
     {"price": 500}, "免運專案（方案二）", 60, "26430"),
    ("方案一 6%：500 元收 30",
     {"price": 500, "plan": "方案一"}, "免運專案（方案一）", 30, "26430"),
    ("方案一不含運費：500 元、運費 60 還是收 30",
     {"price": 500, "buyer_shipping": 60, "plan": "方案一"}, "免運專案（方案一）", 30, "18733"),
    ("方案一 1,000 元收 60",
     {"price": 1000, "plan": "方案一"}, "免運專案（方案一）", 60, "18733"),
    ("方案二 25 元訂單：25 × 6% = 1.5 記 2",
     {"price": 25, "plan": "方案二"}, "免運專案（方案二）", 2, "18733"),
    ("方案二 99 元訂單：99 × 6% = 5.94 記 6",
     {"price": 99, "plan": "方案二"}, "免運專案（方案二）", 6, "18733"),
    ("方案二 100 元訂單：固定 60",
     {"price": 100, "plan": "方案二"}, "免運專案（方案二）", 60, "18733"),
    ("方案二看整筆訂單：單價 50 × 2 件 = 100 元，收 60",
     {"price": 50, "qty": 2, "plan": "方案二"}, "免運專案（方案二）", 60, "18733"),
    ("10% 蝦幣回饋減免 0.5%：400 元 × 2.5% = 10",
     {"price": 400, "coin_cashback": "3%"}, "蝦幣回饋", 10, "18733"),
    ("發票服務費每張 2.5",
     {"price": 500}, "發票服務費", 2.5, "26430"),
]


def main():
    failed = 0
    for desc, o, key, expected, src in CASES:
        res = calc.calc(dict(o), FEES)
        got = next((v for n, v in res["平台扣款明細"] if n.startswith(key)), None)
        ok = got == expected
        failed += not ok
        print(f"{'OK ' if ok else '錯 '} {desc}（公告 {src}）：應 {expected}，算出 {got}")
    # 整批合計：方案比較
    rows = [calc.calc({"price": p, "plan": "方案二"}, FEES) for p in (50, 300, 1500)]
    summ = calc.summarize(rows)
    for desc, got, expected in [
        ("整批全用方案一：3 + 18 + 90", summ["全用方案一"], 111),
        ("整批全用方案二：3 + 60 + 60", summ["全用方案二"], 123),
        ("整批目前方案二", summ["目前方案費用"], 123),
    ]:
        ok = got == expected
        failed += not ok
        print(f"{'OK ' if ok else '錯 '} {desc}：應 {expected}，算出 {got}")
    # 欄位判斷：寫法不標準的方案、欄位名稱
    for raw, expected in [("型1", "plan1"), ("1型", "plan1"), ("免運方案(一)", "plan1"),
                          ("6%方案", "plan1"), ("型2", "plan2"), ("固定60元", "plan2"),
                          ("方案二（60元）", "plan2")]:
        got = calc.parse_plan(raw, "plan2")[0]
        ok = got == expected
        failed += not ok
        print(f"{'OK ' if ok else '錯 '} 方案寫成「{raw}」：應 {expected}，讀成 {got}")
    for raw in ("甲", "方案1/2"):
        try:
            calc.parse_plan(raw, "plan2")
            ok = False
        except calc.InputError:
            ok = True
        failed += not ok
        print(f"{'OK ' if ok else '錯 '} 方案寫成「{raw}」看不懂要報錯，不能亂猜")
    headers = ["訂單號", "收件人", "手機", "商品金額(元)", "件數", "進貨價", "方案型別", "運費", "賣家負擔運費"]
    keys, _ = calc.detect_columns(headers)
    for h, expected in [("商品金額(元)", "price"), ("件數", "qty"), ("進貨價", "cost"),
                        ("方案型別", "plan"), ("運費", "buyer_shipping"), ("賣家負擔運費", "seller_shipping"),
                        ("收件人", None), ("手機", None)]:
        ok = keys.get(h) == expected
        failed += not ok
        print(f"{'OK ' if ok else '錯 '} 欄位「{h}」：應 {expected}，判斷成 {keys.get(h)}")
    keys, _ = calc.detect_columns(headers, calc.parse_column_map("運費=忽略,商品金額=售價"))
    ok = "運費" not in keys and keys.get("商品金額(元)") == "price"
    failed += not ok
    print(f"{'OK ' if ok else '錯 '} --map 可以忽略欄位、可以省略標題的括號")
    print("全部通過" if not failed else f"{failed} 題沒過")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
