import os
import json
import time
import pyotp
import requests
from SmartApi import SmartConnect

TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "YOUR_BOT_TOKEN_HERE")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "YOUR_CHAT_ID_HERE")
STATE_FILE = "holdings_state.json"

# Paste your 24 ACCOUNTS list here
ACCOUNTS = [
    # ... your 24 accounts ...
]

def send_telegram_alert(message):
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    requests.post(url, json={"chat_id": TELEGRAM_CHAT_ID, "text": message, "parse_mode": "HTML"}, timeout=10)

def fetch_all():
    curr_state = {}
    successful_codes = set()

    for acc in ACCOUNTS:
        try:
            obj = SmartConnect(api_key=acc["api_key"])
            totp = pyotp.TOTP(acc["totp_secret"]).now()
            session = obj.generateSession(acc["client_code"], acc["pin"], totp)
            if not session.get("status"):
                continue

            resp = obj.holding()
            if not resp.get("status"):
                continue

            successful_codes.add(acc["client_code"])
            for item in (resp.get("data") or []):
                qty = int(item.get("quantity", 0) or 0) + int(item.get("t1quantity", 0) or item.get("t1Quantity", 0) or 0)
                symbol = str(item.get("tradingsymbol", "")).strip()
                if qty > 0 and symbol:
                    key = f"{acc['client_code']}|{symbol}"
                    curr_state[key] = {"account": acc["account_name"], "stock": symbol, "qty": qty}
        except Exception as e:
            print(f"Error {acc['account_name']}: {e}")
        time.sleep(0.5)

    return curr_state, successful_codes

def main():
    prev_state = {}
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            prev_state = json.load(f)

    curr_state, ok_codes = fetch_all()
    if not ok_codes:
        return

    if not prev_state:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(curr_state, f, indent=2)
        return

    alerts = []
    for key, info in curr_state.items():
        if key not in prev_state:
            alerts.append(f"🟢 <b>NEW BUY</b> | {info['account']} | <b>{info['stock']}</b>: 0 → <b>{info['qty']}</b>")
        elif info["qty"] != int(prev_state[key]["qty"]):
            diff = info["qty"] - int(prev_state[key]["qty"])
            sign = "+" if diff > 0 else ""
            alerts.append(f"🔄 <b>QTY CHANGED</b> | {info['account']} | <b>{info['stock']}</b>: {prev_state[key]['qty']} → <b>{info['qty']}</b> ({sign}{diff})")

    for key, old_info in prev_state.items():
        code = key.split("|")[0]
        if code in ok_codes and key not in curr_state:
            alerts.append(f"🔴 <b>SOLD OUT</b> | {old_info['account']} | <b>{old_info['stock']}</b>: {old_info['qty']} → <b>0</b>")
        elif code not in ok_codes:
            curr_state[key] = old_info

    if alerts:
        send_telegram_alert(f"🚨 <b>Portfolio Alert ({time.strftime('%d-%b %I:%M %p')})</b>\n\n" + "\n".join(alerts))

    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(curr_state, f, indent=2)

if __name__ == "__main__":
    main()
