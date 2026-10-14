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
    print("全部通過" if not failed else f"{failed} 題沒過")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
