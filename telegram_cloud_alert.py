import os
import io
import json
import time
import datetime
import pyotp
import openpyxl
import requests
from collections import defaultdict
from SmartApi import SmartConnect

# ================= CONFIGURATION =================
TELEGRAM_BOT_TOKEN = "9166325450:AAH7Zh2oGZ-VBtwlkINngFaZo6PWre-doUM"
TELEGRAM_CHAT_ID = "1448796295"
STATE_FILE = "holdings_state.json"
GAIN_STEP_PCT = 5  # Intraday mein har +5% (+5%, +10%, +15%...) up hone par alert bhejega

ONEDRIVE_CLIENT_ID = os.environ.get("ONEDRIVE_CLIENT_ID", "")
ONEDRIVE_CLIENT_SECRET = os.environ.get("ONEDRIVE_CLIENT_SECRET", "")
ONEDRIVE_REFRESH_TOKEN = os.environ.get("ONEDRIVE_REFRESH_TOKEN", "")
ONEDRIVE_FILE_PATH = "Documents/InvestMart_Master_Wealth_Tracker.xlsx"

START_ROW = 7
DASH_HEADER_ROW = 15
DASH_START_ROW = 16

ACCOUNTS = [
    {"account_name": "Roshan lal", "client_code": "AACI043934", "pin": "0135", "api_key": "j9YmE0qu", "totp_secret": "BGTJAYGATBJT2O3GFWKI7JPFD4"},
    {"account_name": "Roshan lal huf", "client_code": "AACN450678", "pin": "0135", "api_key": "WQaOljLk", "totp_secret": "4M4ZN6ODYPVBFCL42WKQLK3TMA"},
    {"account_name": "Aniket", "client_code": "AACI043592", "pin": "0135", "api_key": "YDbeLBqW", "totp_secret": "GT36L6XXQ67HVUHNMZXNNAQZFE"},
    {"account_name": "Bhumika", "client_code": "AACI063602", "pin": "0135", "api_key": "H7m9OODA", "totp_secret": "W5PP36UTOFOP6PEXRXBBM2HH5A"},
    {"account_name": "Rakesh Bansal", "client_code": "R297279", "pin": "0135", "api_key": "ErMVysb3", "totp_secret": "OGNEXOGC5IOPGG26HTIOTETY6U"},
    {"account_name": "Rakesh bansal HUF", "client_code": "AACN450426", "pin": "0135", "api_key": "9UIEQ3XH", "totp_secret": "24ORPLKAQ57ZJ4XDOKUXWR2CFI"},
    {"account_name": "Kashish Bansal", "client_code": "AAAC015110", "pin": "0135", "api_key": "NrAfSfRm", "totp_secret": "IEF3A2CBWDEJ4R2WSXZIEWQR34"},
    {"account_name": "Reena devi", "client_code": "R282962", "pin": "0135", "api_key": "8E1oTBdv", "totp_secret": "3RUYUPX4OFYW6UVIK5XE44TEDE"},
    {"account_name": "Pawan Kumar HUF", "client_code": "AACN448804", "pin": "0135", "api_key": "ks378IS0", "totp_secret": "FQ73WH6HXGJUWRHSG3ZIQLKUX4"},
    {"account_name": "Mukesh Kumar Narang HUF", "client_code": "AACN784607", "pin": "0135", "api_key": "j9YmEqu", "totp_secret": "BGTJAYGATBJT2O3GFWKI7JPFD4"},
    {"account_name": "Rukmani devi", "client_code": "AAAF368129", "pin": "0135", "api_key": "dDmDd1bS", "totp_secret": "7FXGBYZSJ6MHGIPC5S33SRCQEY"},
    {"account_name": "RIYA", "client_code": "R597315", "pin": "0135", "api_key": "57EvxIi7", "totp_secret": "PIOKVKIJIEJGOQJSFXDDBVQK6E"},
    {"account_name": "Seema devi", "client_code": "S54748898", "pin": "0135", "api_key": "yc8yTOIg", "totp_secret": "S3UYIELEZYRIG3YULOP3U7VSTA"},
    {"account_name": "RAJESH KUMAR", "client_code": "AACI064276", "pin": "0135", "api_key": "1M69CtnL", "totp_secret": "2QQWJEVIVG3FCA7YESVY5ATGPY"},
    {"account_name": "Diksha Bansal", "client_code": "D54745460", "pin": "0135", "api_key": "5M78l5lP", "totp_secret": "L7CN5TVSQXX7KVDGJLINA5FRZA"},
    {"account_name": "SANJU DEVI", "client_code": "S1136788", "pin": "0135", "api_key": "lK6vgf14", "totp_secret": "D3L52DH4VFPH65HNY3FRXJOCTY"},
    {"account_name": "Rajender Kumar", "client_code": "R53517293", "pin": "0135", "api_key": "FUv8x8TD", "totp_secret": "IIHBBS4GIY7H2BKNEZJSALYPMI"},
    {"account_name": "Rajender Kumar HUF", "client_code": "AACN450610", "pin": "0135", "api_key": "09uMRRcy", "totp_secret": "D4BPDB6WCLEWG3R33VPF6LPTRI"},
    {"account_name": "Narender Garg", "client_code": "AACH009422", "pin": "0135", "api_key": "F5eInGid", "totp_secret": "FCL2C6LMSDQNNBCNNLPOCLSZ4I"},
    {"account_name": "Narender Garg HUF", "client_code": "AACN450253", "pin": "0135", "api_key": "GvAkAXUG", "totp_secret": "G2UNVVSWPRCM435KKDUH7XMLIM"},
    {"account_name": "Pawan Kumar", "client_code": "AACM885301", "pin": "0135", "api_key": "9XlBd97D", "totp_secret": "4I5JUW64T5TWYSAPFCIY66HIXA"},
    {"account_name": "Peeyush Sharma", "client_code": "P63048331", "pin": "0135", "api_key": "q9TYE8Nh", "totp_secret": "GQTCQGHA6TYK6U2CUGEDONIQE4"},
    {"account_name": "Saroj devi", "client_code": "AACM885810", "pin": "0135", "api_key": "gRvZQHPU", "totp_secret": "AOK4MS3UWTJWF7MGB2VER3SJ7E"},
    {"account_name": "GOURAV DHURIA", "client_code": "AACM872342", "pin": "0135", "api_key": "X5Rk6caa", "totp_secret": "DSIFI4B5A444WWYHQ5MY5PW64A"},
    {"account_name": "Rajni Bala", "client_code": "AACM873426", "pin": "0135", "api_key": "roARzAl4", "totp_secret": "TAOJRZNVZULFXURLTW5LD24XVQ"},
    {"account_name": "Vikram Suthar", "client_code": "AACM87426", "pin": "0135", "api_key": "roARzA4", "totp_secret": "TAOJRZNVZULFXURLTW5LD24XVQ"},
    {"account_name": "Raju Udaniya", "client_code": "AACM87426", "pin": "0135", "api_key": "roARzA4", "totp_secret": "TAOJRZNVZULFXURLTW5LD24XVQ"},
    {"account_name": "Baby", "client_code": "AACM848071", "pin": "0135", "api_key": "lfLFt7Zv", "totp_secret": "3YFQWFNHGDJEA2GUVIPVFMHGNU"},
    {"account_name": "Sandeep Singharia", "client_code": "AACM846943", "pin": "0135", "api_key": "QGKhu12E", "totp_secret": "LM67JG63XYRMBLX57BYBUMAKOM"},
    {"account_name": "Mukesh Kumar Narang", "client_code": "AACM853089", "pin": "0135", "api_key": "d5ISqdtR", "totp_secret": "HVKJZDRJUQID5VXY7DUTTIGZGY"},
    {"account_name": "Monika devi", "client_code": "AAAH788468", "pin": "0135", "api_key": "rLmVDkmm", "totp_secret": "5KAE2CYQZJQPXXZJO32EC4Z52A"},
    {"account_name": "ATMA RAM", "client_code": "AAAA932378", "pin": "0135", "api_key": "3HilUF2v", "totp_secret": "XTNPQ2BH43PZUJXL3D2OF2KERQ"},
]
# =================================================


def get_ist_now():
    return datetime.datetime.utcnow() + datetime.timedelta(hours=5, minutes=30)


def send_telegram_alert(message, chat_id=None):
    target_chat = str(chat_id or TELEGRAM_CHAT_ID)
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    try:
        r = requests.post(url, json={"chat_id": target_chat, "text": message, "parse_mode": "HTML"}, timeout=10)
        print("Telegram API Response:", r.status_code, r.text)
    except Exception as e:
        print(f"Telegram send error: {e}")


def fetch_account_data(acc):
    """Fetches Holdings (with Intraday % change from Prev Close) AND completed non-Intraday orders."""
    try:
        obj = SmartConnect(api_key=acc["api_key"], timeout=20)
        totp = pyotp.TOTP(acc["totp_secret"]).now()
        session = obj.generateSession(acc["client_code"], acc["pin"], totp)
        if not session.get("status"):
            return False, [], [], session.get("message", "Login failed")

        # 1. Fetch Holdings
        resp = obj.holding()
        if not resp.get("status"):
            return False, [], [], resp.get("message", "Fetch failed")

        items = []
        for item in (resp.get("data") or []):
            settled_qty = int(item.get("quantity", 0) or 0)
            t1_qty = int(item.get("t1quantity", 0) or item.get("t1Quantity", 0) or 0)
            qty = settled_qty if (settled_qty > 0 and settled_qty == t1_qty) else (settled_qty + t1_qty)
            symbol = str(item.get("tradingsymbol", "")).strip()

            if qty > 0 and symbol:
                avg = float(item.get("averageprice", 0.0) or 0.0)
                ltp = float(item.get("ltp", 0.0) or 0.0)
                prev_close = float(item.get("close", 0.0) or item.get("prevclose", 0.0) or 0.0)

                inv = qty * avg
                cur = qty * ltp

                if prev_close > 0:
                    intraday_pct = round(((ltp - prev_close) / prev_close) * 100, 2)
                else:
                    intraday_pct = 0.0

                curr_milestone = int(intraday_pct // GAIN_STEP_PCT) * GAIN_STEP_PCT if intraday_pct >= GAIN_STEP_PCT else 0

                items.append({
                    "stock": symbol,
                    "qty": qty,
                    "avg": round(avg, 2),
                    "ltp": round(ltp, 2),
                    "prev_close": round(prev_close, 2),
                    "invested": round(inv, 2),
                    "current": round(cur, 2),
                    "pnl": round(cur - inv, 2),
                    "intraday_pct": intraday_pct,
                    "milestone": curr_milestone,
                })

        # 2. Fetch Today's OrderBook — ONLY completed Delivery/MTF orders (Intraday ignored)
        delivery_orders = []
        try:
            ob_resp = obj.orderBook()
            if ob_resp and ob_resp.get("status") and ob_resp.get("data"):
                for o in ob_resp["data"]:
                    status = str(o.get("status", "") or o.get("orderstatus", "")).lower().strip()
                    prod = str(o.get("producttype", "")).upper().strip()
                    if status == "complete" and prod not in ("INTRADAY", "MIS"):
                        oid = str(o.get("orderid", "")).strip()
                        if oid:
                            delivery_orders.append({
                                "orderid": oid,
                                "account": acc["account_name"],
                                "client_code": acc["client_code"],
                                "stock": str(o.get("tradingsymbol", "")).strip(),
                                "side": str(o.get("transactiontype", "")).upper().strip(),
                                "qty": int(float(o.get("filledshares", 0) or o.get("quantity", 0) or 0)),
                                "price": float(o.get("averageprice", 0.0) or o.get("price", 0.0) or 0.0),
                            })
        except Exception:
            pass

        return True, items, delivery_orders, "OK"
    except Exception as e:
        return False, [], [], str(e)


# ================= TELEGRAM COMMAND HANDLERS =================
def build_portfolio_summary_msg(stocks_dict, title_label="Portfolio Summary"):
    tot_inv = 0.0
    tot_cur = 0.0
    stock_agg = {}

    for info in stocks_dict.values():
        qty = int(info.get("qty", 0))
        avg = float(info.get("avg", 0.0))
        ltp = float(info.get("ltp", 0.0))
        inv = qty * avg
        cur = qty * ltp
        tot_inv += inv
        tot_cur += cur

        sym = info["stock"]
        if sym not in stock_agg:
            stock_agg[sym] = {"qty": 0, "inv": 0.0, "cur": 0.0}
        stock_agg[sym]["qty"] += qty
        stock_agg[sym]["inv"] += inv
        stock_agg[sym]["cur"] += cur

    net_pnl = tot_cur - tot_inv
    net_pct = (net_pnl / tot_inv * 100) if tot_inv > 0 else 0.0
    icon = "🟢" if net_pnl >= 0 else "🔴"

    ranked = sorted(stock_agg.items(), key=lambda x: x[1]["cur"] - x[1]["inv"], reverse=True)
    lines = [
        f"📊 <b>{title_label}</b> ({get_ist_now().strftime('%d-%b %I:%M %p IST')})",
        f"💰 <b>Total Invested:</b> ₹{tot_inv:,.2f}",
        f"💼 <b>Current Value:</b> ₹{tot_cur:,.2f}",
        f"{icon} <b>Net P&L:</b> ₹{net_pnl:,.2f} ({net_pct:+.2f}%)\n",
        "🚀 <b>Top Gainers:</b>"
    ]
    for sym, d in ranked[:3]:
        lines.append(f"  • {sym}: ₹{d['cur'] - d['inv']:+,.0f} ({d['qty']} qty)")

    lines.append("\n🔻 <b>Bottom Movers:</b>")
    for sym, d in ranked[-3:]:
        lines.append(f"  • {sym}: ₹{d['cur'] - d['inv']:+,.0f} ({d['qty']} qty)")

    return "\n".join(lines)


def handle_stock_query(symbol_query, stocks_dict, chat_id):
    q = symbol_query.upper().replace("/STOCK", "").strip()
    matches = [v for v in stocks_dict.values() if q in v["stock"].upper()]
    if not matches:
        send_telegram_alert(f"❌ No active holdings found matching <b>{q}</b>.", chat_id)
        return

    total_qty = sum(int(m["qty"]) for m in matches)
    total_cur = sum(int(m["qty"]) * float(m["ltp"]) for m in matches)
    total_pnl = sum(int(m["qty"]) * (float(m["ltp"]) - float(m["avg"])) for m in matches)
    icon = "🟢" if total_pnl >= 0 else "🔴"

    lines = [
        f"🔎 <b>Stock Lookup: {q}</b>",
        f"📦 <b>Total Qty:</b> {total_qty} | <b>Current:</b> ₹{total_cur:,.0f} ({icon} ₹{total_pnl:+,.0f})\n"
    ]
    for m in sorted(matches, key=lambda x: int(x["qty"]), reverse=True):
        pnl = int(m["qty"]) * (float(m["ltp"]) - float(m["avg"]))
        lines.append(f"• <b>{m['account']}</b>: {m['qty']} qty @ ₹{m['avg']} (₹{pnl:+,.0f})")
    send_telegram_alert("\n".join(lines), chat_id)


def handle_client_query(query_text, chat_id):
    q = query_text.lower().replace("/holding", "").replace("/holdings", "").strip()
    if q in ("/clients", "clients", "/list", "list"):
        lines = ["📋 <b>Available Clients List:</b>\n"]
        for idx, acc in enumerate(ACCOUNTS, 1):
            lines.append(f"{idx}. <b>{acc['account_name']}</b> (<code>{acc['client_code']}</code>)")
        send_telegram_alert("\n".join(lines), chat_id)
        return

    matched = [acc for acc in ACCOUNTS if q in acc["account_name"].lower() or q == acc["client_code"].lower()]
    if not matched:
        send_telegram_alert(f"❌ Koi client nahi mila: <b>{query_text}</b>", chat_id)
        return

    for acc in matched:
        ok, items, _, msg = fetch_account_data(acc)
        if not ok:
            send_telegram_alert(f"⚠️ <b>{acc['account_name']}</b>: {msg}", chat_id)
            continue
        if not items:
            send_telegram_alert(f"📭 <b>{acc['account_name']}</b>: Koi active holding nahi hai.", chat_id)
            continue

        tot_inv = sum(x["invested"] for x in items)
        tot_cur = sum(x["current"] for x in items)
        tot_pnl = tot_cur - tot_inv
        pnl_icon = "🟢" if tot_pnl >= 0 else "🔴"

        lines = [
            f"📊 <b>{acc['account_name']}</b> (<code>{acc['client_code']}</code>)",
            f"💰 Invested: ₹{tot_inv:,.2f} | Current: ₹{tot_cur:,.2f}",
            f"{pnl_icon} Net P&L: <b>₹{tot_pnl:,.2f}</b>\n"
        ]
        for x in items:
            s_icon = "🟢" if x["pnl"] >= 0 else "🔴"
            lines.append(f"• <b>{x['stock']}</b>: {x['qty']} qty @ ₹{x['avg']} → LTP ₹{x['ltp']} ({s_icon} ₹{x['pnl']:,.0f} | Today: {x['intraday_pct']:+.2f}%)")
        send_telegram_alert("\n".join(lines), chat_id)
        time.sleep(0.5)


def check_telegram_commands(state_data, stocks_dict):
    last_update_id = state_data.get("_last_update_id", 0)
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/getUpdates"
    try:
        resp = requests.get(url, params={"offset": last_update_id + 1, "timeout": 3}, timeout=10).json()
        if not resp.get("ok"):
            return last_update_id

        for upd in resp.get("result", []):
            last_update_id = max(last_update_id, upd["update_id"])
            msg = upd.get("message") or {}
            sender_id = str(msg.get("chat", {}).get("id", ""))
            text = str(msg.get("text", "")).strip()
            if sender_id != TELEGRAM_CHAT_ID or not text:
                continue

            low = text.lower()
            if low in ("/start", "/help", "help"):
                send_telegram_alert(
                    "👋 <b>Commands Available:</b>\n"
                    "• <code>/summary</code> — Total Invested, Current Value & Net P&L\n"
                    "• <code>/stock SYMBOL</code> — Check who holds a stock\n"
                    "• <code>/clients</code> — List all 32 accounts\n"
                    "• Type any <b>Client Name</b> (e.g., <code>Aniket</code>) for live holdings",
                    sender_id
                )
            elif low in ("/summary", "summary"):
                send_telegram_alert(build_portfolio_summary_msg(stocks_dict, "Live Portfolio Summary"), sender_id)
            elif low.startswith("/stock"):
                handle_stock_query(text, stocks_dict, sender_id)
            else:
                handle_client_query(text, sender_id)
    except Exception as e:
        print(f"Command check error: {e}")
    return last_update_id


# ================= ONEDRIVE EXCEL SYNC =================
def get_onedrive_access_token():
    if not (ONEDRIVE_CLIENT_ID and ONEDRIVE_CLIENT_SECRET and ONEDRIVE_REFRESH_TOKEN):
        return None
    url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    data = {
        "client_id": ONEDRIVE_CLIENT_ID,
        "client_secret": ONEDRIVE_CLIENT_SECRET,
        "refresh_token": ONEDRIVE_REFRESH_TOKEN,
        "grant_type": "refresh_token",
        "scope": "offline_access Files.ReadWrite.All",
    }
    try:
        return requests.post(url, data=data, timeout=15).json().get("access_token")
    except Exception:
        return None


def sync_excel_to_onedrive(all_excel_rows):
    if not all_excel_rows:
        return
    token = get_onedrive_access_token()
    if not token:
        return

    headers = {"Authorization": f"Bearer {token}"}
    file_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{ONEDRIVE_FILE_PATH}:/content"
    dl_resp = requests.get(file_url, headers=headers, timeout=30)
    if not dl_resp.ok:
        return

    wb = openpyxl.load_workbook(io.BytesIO(dl_resp.content))
    total_invested = sum(float(r[6]) for r in all_excel_rows)
    total_current = sum(float(r[7]) for r in all_excel_rows)
    total_pnl = sum(float(r[8]) for r in all_excel_rows)
    ist_time = get_ist_now().strftime("%H:%M:%S")

    total_row = ["TOTAL", "", "", "", "", "", round(total_invested, 2), round(total_current, 2), round(total_pnl, 2), ist_time]
    sheet = wb["Demat_Holdings_(Equity_ETFs)"]
    for row in sheet[f"A{START_ROW}:J500"]:
        for cell in row:
            cell.value = None
    for r_idx, row_vals in enumerate(all_excel_rows + [total_row], start=START_ROW):
        for c_idx, val in enumerate(row_vals, start=1):
            sheet.cell(row=r_idx, column=c_idx, value=val)

    if "Dashboard_Summary" in wb.sheetnames:
        dash = wb["Dashboard_Summary"]
        dash["C12"].value = round(total_invested, 2)
        dash["D12"].value = round(total_current, 2)
        client_names = [acc["account_name"].strip() for acc in ACCOUNTS]
        stock_data = defaultdict(lambda: defaultdict(int))
        stock_totals = defaultdict(int)
        for r in all_excel_rows:
            stock_data[r[2]][r[0]] += int(r[3])
            stock_totals[r[2]] += int(r[3])

        for row in dash[f"A{DASH_HEADER_ROW}:AZ500"]:
            for cell in row:
                cell.value = None
        for c_idx, h in enumerate(["Stock/Name", "TOTAL QTY"] + client_names, start=1):
            dash.cell(row=DASH_HEADER_ROW, column=c_idx, value=h)
        for r_idx, stock in enumerate(sorted(stock_totals.keys()), start=DASH_START_ROW):
            dash.cell(row=r_idx, column=1, value=stock)
            dash.cell(row=r_idx, column=2, value=stock_totals[stock])
            for c_idx, client in enumerate(client_names, start=3):
                qty = stock_data[stock].get(client, 0)
                dash.cell(row=r_idx, column=c_idx, value=(qty if qty > 0 else None))

    out_stream = io.BytesIO()
    wb.save(out_stream)
    wb.close()
    requests.put(file_url, headers={"Authorization": f"Bearer {token}", "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}, data=out_stream.getvalue(), timeout=30)


# ================= MAIN EXECUTION =================
def main():
    print("=== START CLOUD ALERT SCRIPT ===")
    prev_state = {}
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            prev_state = json.load(f)
    print(f"Loaded prev_state keys: {len(prev_state)}")

    curr_stocks = {}
    ok_codes = set()
    all_excel_rows = []
    all_delivery_orders = []
    ist_now = get_ist_now()
    ist_time = ist_now.strftime("%H:%M:%S")
    today_str = ist_now.strftime("%Y-%m-%d")

    prev_date = prev_state.get("_intraday_date", "")
    is_new_day = (prev_date != today_str)

    for acc in ACCOUNTS:
        ok, items, del_orders, msg = fetch_account_data(acc)
        if ok:
            ok_codes.add(acc["client_code"])
            all_delivery_orders.extend(del_orders)
            for x in items:
                key = f"{acc['client_code']}|{x['stock']}"
                curr_stocks[key] = {
                    "account": acc["account_name"],
                    "stock": x["stock"],
                    "qty": x["qty"],
                    "avg": x["avg"],
                    "ltp": x["ltp"],
                    "prev_close": x["prev_close"],
                    "intraday_pct": x["intraday_pct"],
                    "milestone": x["milestone"],
                }
                all_excel_rows.append([
                    acc["account_name"].strip(), acc["client_code"].strip(),
                    x["stock"], x["qty"], x["avg"], x["ltp"],
                    x["invested"], x["current"], x["pnl"], ist_time
                ])
        else:
            print(f"FAILED {acc['account_name']}: {msg}")
        time.sleep(0.5)

    print(f"Fetched OK accounts: {len(ok_codes)} / {len(ACCOUNTS)} | Active stocks: {len(curr_stocks)} | Orders today: {len(all_delivery_orders)}")
    if not ok_codes:
        print("EXITING: No accounts fetched!")
        return

    new_update_id = check_telegram_commands(prev_state, curr_stocks)

    if ist_now.minute % 15 < 2:
        sync_excel_to_onedrive(all_excel_rows)

    seen_orders = {oid for oid in prev_state.get("_seen_delivery_orders", []) if oid.startswith(today_str)}
    order_alerts = []
    for o in all_delivery_orders:
        unique_oid = f"{today_str}_{o['client_code']}_{o['orderid']}"
        if unique_oid not in seen_orders:
            seen_orders.add(unique_oid)
            if prev_state:
                badge = "🟢 <b>DELIVERY BUY EXECUTED</b>" if o["side"] == "BUY" else "🔴 <b>DELIVERY SELL EXECUTED</b>"
                order_alerts.append(
                    f"{badge} | {o['account']} | <b>{o['stock']}</b>\n"
                    f"   Qty: <b>{o['qty']}</b> @ ₹{o['price']}"
                )

    prev_stocks = {k: v for k, v in prev_state.items() if not k.startswith("_")}
    qty_alerts = []
    pct_alerts = []
    alerted_intraday_symbols = set()

    if prev_stocks:
        for key, info in curr_stocks.items():
            if key not in prev_stocks:
                qty_alerts.append(f"🟢 <b>HOLDING ADDED</b> | {info['account']} | <b>{info['stock']}</b>: 0 → <b>{info['qty']}</b>")
            else:
                old_info = prev_stocks[key]
                old_qty = int(old_info["qty"])

                if info["qty"] != old_qty:
                    diff = info["qty"] - old_qty
                    sign = "+" if diff > 0 else ""
                    qty_alerts.append(f"🔄 <b>QTY CHANGED</b> | {info['account']} | <b>{info['stock']}</b>: {old_qty} → <b>{info['qty']}</b> ({sign}{diff})")

                prev_milestone = 0 if is_new_day else int(old_info.get("milestone", 0))
                curr_milestone = int(info["milestone"])

                if curr_milestone > prev_milestone and curr_milestone >= GAIN_STEP_PCT:
                    sym_ms_key = f"{info['stock']}_{curr_milestone}"
                    if sym_ms_key not in alerted_intraday_symbols:
                        alerted_intraday_symbols.add(sym_ms_key)
                        pct_alerts.append(
                            f"🚀 <b>INTRADAY +{curr_milestone}% UP</b> | <b>{info['stock']}</b>\n"
                            f"   Prev Close: ₹{info['prev_close']} → LTP: <b>₹{info['ltp']}</b> (<b>+{info['intraday_pct']}% Today</b>)"
                        )
                else:
                    info["milestone"] = max(prev_milestone, curr_milestone)

        for key, old_info in prev_stocks.items():
            code = key.split("|")[0]
            if code in ok_codes and key not in curr_stocks:
                qty_alerts.append(f"🔴 <b>SOLD OUT</b> | {old_info['account']} | <b>{old_info['stock']}</b>: {old_info['qty']} → <b>0</b>")
            elif code not in ok_codes:
                curr_stocks[key] = old_info

    all_alerts = order_alerts + qty_alerts + pct_alerts
    print(f"Alerts generated: {len(all_alerts)} (Orders: {len(order_alerts)}, Qty: {len(qty_alerts)}, Intraday%: {len(pct_alerts)})")
    if all_alerts:
        send_telegram_alert(f"🚨 <b>Portfolio Alert ({ist_now.strftime('%d-%b %I:%M %p')})</b>\n\n" + "\n\n".join(all_alerts))

    sent_slots = {s for s in prev_state.get("_sent_summary_slots", []) if s.startswith(today_str)}
    if ist_now.weekday() < 5:
        hhmm = ist_now.hour * 60 + ist_now.minute
        for slot_key, start_m, end_m, label in [
            ("09:30_OPEN", 9 * 60 + 30, 9 * 60 + 45, "Market Open Portfolio Snapshot (9:30 AM)"),
            ("13:00_MID", 13 * 60, 13 * 60 + 15, "Mid-Day Portfolio Snapshot (1:00 PM)"),
            ("15:35_CLOSE", 15 * 60 + 35, 15 * 60 + 50, "Market Close Portfolio Snapshot (3:35 PM)"),
        ]:
            full_slot = f"{today_str}_{slot_key}"
            if start_m <= hhmm <= end_m and full_slot not in sent_slots:
                send_telegram_alert(build_portfolio_summary_msg(curr_stocks, label))
                sent_slots.add(full_slot)

    final_state = {
        "_last_update_id": new_update_id,
        "_intraday_date": today_str,
        "_sent_summary_slots": sorted(list(sent_slots)),
        "_seen_delivery_orders": sorted(list(seen_orders)),
    }
    final_state.update(curr_stocks)

    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(final_state, f, indent=2)
    print("=== END CLOUD ALERT SCRIPT ===")


if __name__ == "__main__":
    main()
