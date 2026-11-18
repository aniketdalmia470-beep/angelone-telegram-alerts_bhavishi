import os
import json
import time
import pyotp
import requests
from SmartApi import SmartConnect

TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "9166325450:AAH7Zh2oGZ-VBtwlkINngFaZo6PWre-doUM")
TELEGRAM_CHAT_ID = str(os.environ.get("TELEGRAM_CHAT_ID", "1448796295"))
STATE_FILE = "holdings_state.json"
GAIN_STEP_PCT = 5  # Har 5% (5%, 10%, 15%...) gain par alert bhejega

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


def send_telegram_alert(message, chat_id=None):
    target_chat = str(chat_id or TELEGRAM_CHAT_ID)
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    try:
        requests.post(url, json={"chat_id": target_chat, "text": message, "parse_mode": "HTML"}, timeout=10)
    except Exception as e:
        print(f"Telegram send error: {e}")


def fetch_single_account(acc):
    """Fetches detailed holdings for a single account with T+1 overlap fix and gain %."""
    try:
        obj = SmartConnect(api_key=acc["api_key"], timeout=20)
        totp = pyotp.TOTP(acc["totp_secret"]).now()
        session = obj.generateSession(acc["client_code"], acc["pin"], totp)
        if not session.get("status"):
            return False, [], session.get("message", "Login failed")

        resp = obj.holding()
        if not resp.get("status"):
            return False, [], resp.get("message", "Fetch failed")

        items = []
        for item in (resp.get("data") or []):
            settled_qty = int(item.get("quantity", 0) or 0)
            t1_qty = int(item.get("t1quantity", 0) or item.get("t1Quantity", 0) or 0)
            qty = settled_qty if (settled_qty > 0 and settled_qty == t1_qty) else (settled_qty + t1_qty)
            symbol = str(item.get("tradingsymbol", "")).strip()

            if qty > 0 and symbol:
                avg = float(item.get("averageprice", 0.0) or 0.0)
                ltp = float(item.get("ltp", 0.0) or 0.0)
                inv = qty * avg
                cur = qty * ltp
                gain_pct = round(((ltp - avg) / avg) * 100, 2) if avg > 0 else 0.0
                curr_milestone = int(gain_pct // GAIN_STEP_PCT) * GAIN_STEP_PCT if gain_pct >= GAIN_STEP_PCT else 0

                items.append({
                    "stock": symbol,
                    "qty": qty,
                    "avg": round(avg, 2),
                    "ltp": round(ltp, 2),
                    "invested": round(inv, 2),
                    "current": round(cur, 2),
                    "pnl": round(cur - inv, 2),
                    "gain_pct": gain_pct,
                    "milestone": curr_milestone,
                })
        return True, items, "OK"
    except Exception as e:
        return False, [], str(e)


def handle_client_query(query_text, chat_id):
    """Finds matching client(s) by name or client code and sends their live holdings."""
    q = query_text.lower().replace("/holding", "").replace("/holdings", "").strip()

    if q in ("/clients", "clients", "/list", "list"):
        lines = ["📋 <b>Available Clients List:</b>\n"]
        for idx, acc in enumerate(ACCOUNTS, 1):
            lines.append(f"{idx}. <b>{acc['account_name']}</b> (<code>{acc['client_code']}</code>)")
        lines.append("\n💡 <i>Kisi bhi client ki holding dekhne ke liye uska naam ya code likh kar bhejein (e.g., Aniket)</i>")
        send_telegram_alert("\n".join(lines), chat_id)
        return

    matched = [
        acc for acc in ACCOUNTS
        if q in acc["account_name"].lower() or q == acc["client_code"].lower()
    ]

    if not matched:
        send_telegram_alert(
            f"❌ Koi client nahi mila: <b>{query_text}</b>\n"
            f"Saare clients dekhne ke liye <code>/clients</code> bhejein.",
            chat_id
        )
        return

    for acc in matched:
        ok, items, msg = fetch_single_account(acc)
        if not ok:
            send_telegram_alert(f"⚠️ <b>{acc['account_name']}</b> ({acc['client_code']}): Error — {msg}", chat_id)
            continue

        if not items:
            send_telegram_alert(f"📭 <b>{acc['account_name']}</b> ({acc['client_code']}): Koi active holding nahi hai.", chat_id)
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
            lines.append(
                f"• <b>{x['stock']}</b>: {x['qty']} qty @ ₹{x['avg']} → LTP ₹{x['ltp']} ({s_icon} ₹{x['pnl']:,.0f})"
            )

        send_telegram_alert("\n".join(lines), chat_id)
        time.sleep(0.5)


def check_telegram_commands(state_data):
    """Checks incoming Telegram messages from the authorized user and replies immediately."""
    last_update_id = state_data.get("_last_update_id", 0)
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/getUpdates"
    params = {"offset": last_update_id + 1, "timeout": 3}

    try:
        resp = requests.get(url, params=params, timeout=10).json()
        if not resp.get("ok"):
            return last_update_id

        for update in resp.get("result", []):
            last_update_id = max(last_update_id, update["update_id"])
            msg = update.get("message") or {}
            sender_id = str(msg.get("chat", {}).get("id", ""))
            text = str(msg.get("text", "")).strip()

            if sender_id != TELEGRAM_CHAT_ID or not text:
                continue

            if text.lower() in ("/start", "/help", "help"):
                send_telegram_alert(
                    "👋 <b>AngelOne Portfolio Bot Ready!</b>\n\n"
                    "• Kisi bhi client ki holding mangne ke liye uska <b>naam</b> ya <b>client code</b> likhein (e.g., <code>Aniket</code> ya <code>Rakesh</code>)\n"
                    "• Saare clients ki list ke liye <code>/clients</code> bhejein\n"
                    "• Overall portfolio summary ke liye <code>/summary</code> bhejein",
                    sender_id
                )
            elif text.lower() in ("/summary", "summary"):
                total_qty_items = [v for k, v in state_data.items() if not k.startswith("_")]
                send_telegram_alert(
                    f"📈 <b>Tracked Active Positions:</b> {len(total_qty_items)} across {len(ACCOUNTS)} accounts.\n"
                    f"Kisi specific client ka full P&L dekhne ke liye uska naam likhein (e.g., <code>Aniket</code>).",
                    sender_id
                )
            else:
                handle_client_query(text, sender_id)

    except Exception as e:
        print(f"Command check error: {e}")

    return last_update_id


def main():
    prev_state = {}
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            prev_state = json.load(f)

    # 1. Pehle aapke bheje hue Telegram commands check karein aur turant reply bhejein
    new_update_id = check_telegram_commands(prev_state)

    # 2. Phir saare accounts ka Buy/Sell + +5%/+10% Milestone check karein
    curr_state = {"_last_update_id": new_update_id}
    ok_codes = set()

    for acc in ACCOUNTS:
        ok, items, _ = fetch_single_account(acc)
        if ok:
            ok_codes.add(acc["client_code"])
            for x in items:
                key = f"{acc['client_code']}|{x['stock']}"
                curr_state[key] = {
                    "account": acc["account_name"],
                    "stock": x["stock"],
                    "qty": x["qty"],
                    "avg": x["avg"],
                    "ltp": x["ltp"],
                    "gain_pct": x["gain_pct"],
                    "milestone": x["milestone"],
                }
        time.sleep(0.5)

    if not ok_codes:
        return

    prev_stocks = {k: v for k, v in prev_state.items() if not k.startswith("_")}
    curr_stocks = {k: v for k, v in curr_state.items() if not k.startswith("_")}

    if not prev_stocks:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(curr_state, f, indent=2)
        return

    qty_alerts = []
    pct_alerts = []

    for key, info in curr_stocks.items():
        if key not in prev_stocks:
            qty_alerts.append(f"🟢 <b>NEW BUY</b> | {info['account']} | <b>{info['stock']}</b>: 0 → <b>{info['qty']}</b>")
        else:
            old_info = prev_stocks[key]
            old_qty = int(old_info["qty"])

            # Quantity Change Check
            if info["qty"] != old_qty:
                diff = info["qty"] - old_qty
                sign = "+" if diff > 0 else ""
                qty_alerts.append(f"🔄 <b>QTY CHANGED</b> | {info['account']} | <b>{info['stock']}</b>: {old_qty} → <b>{info['qty']}</b> ({sign}{diff})")

            # +5% / +10% / +15% Gain Milestone Check
            prev_milestone = int(old_info.get("milestone", info["milestone"]))
            curr_milestone = int(info["milestone"])

            if curr_milestone > prev_milestone and curr_milestone >= GAIN_STEP_PCT:
                pct_alerts.append(
                    f"🚀 <b>+{curr_milestone}% GAIN HIT</b> | {info['account']} | <b>{info['stock']}</b>\n"
                    f"   Avg: ₹{info['avg']} → LTP: <b>₹{info['ltp']}</b> (<b>+{info['gain_pct']}%</b>)"
                )
            else:
                info["milestone"] = max(prev_milestone, curr_milestone)

    # Sold Out Check
    for key, old_info in prev_stocks.items():
        code = key.split("|")[0]
        if code in ok_codes and key not in curr_state:
            qty_alerts.append(f"🔴 <b>SOLD OUT</b> | {old_info['account']} | <b>{old_info['stock']}</b>: {old_info['qty']} → <b>0</b>")
        elif code not in ok_codes:
            curr_state[key] = old_info

    all_alerts = qty_alerts + pct_alerts
    if all_alerts:
        send_telegram_alert(f"🚨 <b>Portfolio Alert ({time.strftime('%d-%b %I:%M %p')})</b>\n\n" + "\n\n".join(all_alerts))

    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(curr_state, f, indent=2)


if __name__ == "__main__":
    main()
