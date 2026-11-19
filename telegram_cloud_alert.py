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
        requests.post(url, json={"chat_id": target_chat, "text": message, "parse_mode": "HTML"}, timeout=10)
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
    last_update_id = state_data
