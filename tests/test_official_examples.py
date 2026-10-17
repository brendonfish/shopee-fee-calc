#!/usr/bin/env python3
"""用蝦皮官方公告裡的範例驗證計算結果。改 fees.json 或程式之後都要跑一次。

用法：python3 tests/test_official_examples.py
全部通過會印「全部通過」，有錯會列出哪一題、應該是多少、算出多少，並以代碼 1 結束。
"""
import re
import subprocess
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
    failed += check_user_fees()
    print("全部通過" if not failed else f"{failed} 題沒過")
    return 1 if failed else 0


def run_cli(cwd, *args):
    """在指定資料夾跑 calc.py，回傳輸出（含錯誤訊息）。預設賣 400、買家運費 60。"""
    import subprocess
    base = ["--price", "400", "--buyer-shipping", "60"]
    if "--price" in args or "--csv" in args:
        base = []
    r = subprocess.run([sys.executable, str(SKILL / "scripts" / "calc.py"), *base, *args],
                       cwd=cwd, capture_output=True, text=True)
    return r.stdout + r.stderr


def check_batch_defaults(check):
    """批次 CSV：命令列給的 --commission-rate、--no-invoice 等當整批預設，CSV 有填就照 CSV。"""
    import tempfile
    with tempfile.TemporaryDirectory() as d:
        src = Path(d) / "o.csv"
        out_csv = Path(d) / "out.csv"

        def deductions(*args):
            run_cli(d, "--csv", str(src), "--out", str(out_csv), *args)
            rows = out_csv.read_text(encoding="utf-8-sig").splitlines()
            col = rows[0].split(",").index("試算_平台扣款合計")
            return [r.split(",")[col] for r in rows[1:]]

        src.write_text("售價,成本\n500,250\n", encoding="utf-8")
        check("CSV 沒指定時照預設：28 + 60 + 13 + 發票 2.5", deductions() == ["103.5"])
        check("CSV 吃 --commission-rate 10% 和 --no-invoice：50 + 60 + 13", deductions("--commission-rate", "10%", "--no-invoice") == ["123"])
        out = run_cli(d, "--csv", str(src), "--out", str(out_csv), "--commission-rate", "10%", "--no-invoice")
        check("CSV 印出用了哪些整批預設", "整批預設" in out and "成交手續費率 10%" in out and "開發票 否" in out)
        src.write_text("售價,成交手續費率\n500,6%\n500,\n", encoding="utf-8")
        check("CSV 有填就照 CSV，空白才用預設：30+60+13、50+60+13", deductions("--commission-rate", "10%", "--no-invoice") == ["103", "123"])
        src.write_text("售價,成本,數量\n500,250,1\n", encoding="utf-8")
        run_cli(d, "--csv", str(src), "--out", str(out_csv))
        rows = out_csv.read_text(encoding="utf-8-sig").splitlines()
        col = rows[0].split(",").index("試算_撥款")
        run_cli(d, "--csv", str(src), "--out", str(out_csv), "--no-invoice")
        rows = out_csv.read_text(encoding="utf-8-sig").splitlines()
        check(f"CSV 撥款是整數就不帶 .0：{rows[1].split(',')[col]}", rows[1].split(",")[col] == "399")
        out = run_cli(d, "--csv", str(src), "--inspect", "--no-invoice")
        check("--inspect 講明成本是每件、會乘上數量", "每件成本" in out)
        check("--inspect 列出整批預設", "整批預設" in out and "開發票 否" in out)
        (Path(d) / "shopee-fees.json").write_text('{"payment_fee_rate": 0.03}', encoding="utf-8")
        out = run_cli(d, "--csv", str(src), "--inspect")
        check("--inspect 也講會套用使用者費率", "用了使用者費率（shopee-fees.json）：金流服務費率 2.5% → 3%" in out)
        out = run_cli(d, "--csv", str(src), "--out", str(out_csv), "--no-invoice")
        check("整批合計印出毛利率", "（毛利率 " in out.split("整批 1 筆")[1].split("\n")[0])
    r = subprocess.run([sys.executable, str(SKILL / "scripts" / "calc.py"), "--help"], capture_output=True, text=True)
    check("--help 不會當掉", r.returncode == 0 and "--set" in r.stdout)


def check_user_fees():
    """使用者自己的費率檔：只寫要改的欄位，疊在內建費率上。"""
    import json
    import subprocess
    import tempfile
    failed = 0

    def check(desc, ok):
        nonlocal failed
        failed += not ok
        print(f"{'OK ' if ok else '錯 '} {desc}")

    def raises(user):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "shopee-fees.json"
            p.write_text(json.dumps(user), encoding="utf-8")
            try:
                calc.load_fees(SKILL / "references" / "fees.json", p)
                return False
            except calc.InputError:
                return True

    with tempfile.TemporaryDirectory() as d:
        p = Path(d) / "shopee-fees.json"
        p.write_text('{"payment_fee_rate": 0.028}', encoding="utf-8")
        fees = calc.load_fees(SKILL / "references" / "fees.json", p)
        res = calc.calc({"price": 400, "buyer_shipping": 60}, fees)
        got = dict(res["平台扣款明細"])["金流服務費"]
        check(f"只改金流費 2.8%：460 × 2.8% = 12.88 記 13，算出 {got}", got == 13)
        check("其他費率照內建：成交手續費率預設還是 5.5%", fees["commission_rate_default"] == FEES["commission_rate_default"])
        check(f"結果寫出改了什麼：{res['使用者費率']['修改']}",
              res["使用者費率"]["修改"] == "金流服務費率 2.5% → 2.8%")
        check(f"費率版本照實寫內建的日期，不被使用者檔案蓋掉：{fees['fees_as_of']}",
              fees["fees_as_of"] == FEES["fees_as_of"])
        check("工作資料夾有 shopee-fees.json 會自動找到", calc.find_user_fees(cwd=d) == p)
        out = subprocess.run([sys.executable, str(SKILL / "scripts" / "calc.py"), "--price", "400",
                              "--buyer-shipping", "60"], cwd=d, capture_output=True, text=True).stdout
        check("命令列在那個資料夾跑，會用使用者的費率並講出來",
              "金流服務費：13 元" in out and "用了使用者費率（" in out)
        out = run_cli(d, "--builtin")
        check("--builtin 不套用使用者費率檔，照內建 2.5% 算 12 元，並講出沒套用",
              "金流服務費：12 元" in out and "沒有套用" in out)
        out = run_cli(d, "--set", "payment_fee_rate=3%")
        check("--set 蓋過使用者費率檔：460 × 3% = 13.8 記 14", "金流服務費：14 元" in out)

    with tempfile.TemporaryDirectory() as d:
        check("工作資料夾沒有就不用", calc.find_user_fees(cwd=d) is None)
        p = Path(d) / "my.json"
        p.write_text('{"shipping_program": {"plan2_fixed": 70}}', encoding="utf-8")
        fees = calc.load_fees(SKILL / "references" / "fees.json", p)
        got = dict(calc.calc({"price": 500, "plan": "方案二"}, fees)["平台扣款明細"])["免運專案（方案二）"]
        check(f"只改方案二每單 70：算出 {got}，方案一費率沒被蓋掉",
              got == 70 and fees["shipping_program"]["plan1_rate"] == FEES["shipping_program"]["plan1_rate"])

    with tempfile.TemporaryDirectory() as d:
        p = Path(d) / "shopee-fees.json"
        p.write_text('{"payment_fee_rate": 0.028, "fees_as_of": "2026-01-01"}', encoding="utf-8")
        u = calc.load_fees(SKILL / "references" / "fees.json", p)["使用者費率"]
        check("使用者檔案比內建費率舊，要提醒新版可能已經包含", any("比內建費率" in n for n in u.get("提醒", [])))
        p.write_text('{"payment_fee_rate": 0.03, "fees_as_of": "2026-11-01"}', encoding="utf-8")
        fees = calc.load_fees(SKILL / "references" / "fees.json", p)
        check("使用者檔案比內建新就不提醒，而且不管日期直接套用",
              "提醒" not in fees["使用者費率"] and fees["payment_fee_rate"] == 0.03)
        out = run_cli(d)
        check("費率版本寫內建日期，加上使用者費率", "費率版本：內建費率 " + FEES["fees_as_of"] + "，加上上面的使用者費率" in out)
        check("輸出不印完整路徑，只寫檔名", d not in out and "shopee-fees.json" in out)

    with tempfile.TemporaryDirectory() as d:
        out = run_cli(d, "--set", "payment_fee_rate=3%")
        check("沒有費率檔也能用 --set 臨時試算，並講出來源",
              "金流服務費：14 元" in out and "--set" in out and "2.5% → 3%" in out)
        out = run_cli(d, "--set", "invoice_default=false")
        check("--set 改發票設定，說明用中文", "沒指定時算不算發票服務費 算 → 不算" in out and "發票服務費" not in out.split("平台扣款")[1].split("合計")[0])
        out = run_cli(d, "--set", "shipping_program.plan2_fixed=70", "--plan", "方案二", "--price", "500")
        check("--set 可以改巢狀欄位：方案二每單 70", "免運專案（方案二）：70 元" in out)
        out = run_cli(d, "--set", "payment_fee_rate=0.03", "--set", "invoice_fee_per_order=3")
        check("--set 可以寫好幾個", "金流服務費：14 元" in out and "發票服務費：3 元" in out)
        check("--set 沒有等號要報錯", "--set 要寫成" in run_cli(d, "--set", "payment_fee_rate"))
        check("--set 欄位打錯要報錯", "不認得的欄位" in run_cli(d, "--set", "payment_rate=3%"))
        check("沒有費率檔時 --builtin 照內建算，不多講", "使用者費率" not in run_cli(d, "--builtin"))

    check_batch_defaults(check)
    check("欄位名稱打錯（payment_rate）要報錯，不能當作沒改", raises({"payment_rate": 0.028}))
    check("費率填成 2.8 而不是 0.028 要報錯", raises({"payment_fee_rate": 2.8}))
    return failed


if __name__ == "__main__":
    sys.exit(main())
