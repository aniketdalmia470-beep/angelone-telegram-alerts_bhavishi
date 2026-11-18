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
    {
        "account_name": "Roshan lal",
        "client_code": "AACI043934",
        "pin": "0135",
        "api_key": "j9YmE0qu",
        "totp_secret": "BGTJAYGATBJT2O3GFWKI7JPFD4",
    },
    {
        "account_name": "Roshan lal huf",
        "client_code": "AACN450678",
        "pin": "0135",
        "api_key": "WQaOljLk",
        "totp_secret": "4M4ZN6ODYPVBFCL42WKQLK3TMA",
    },
    {
        "account_name": "Aniket",
        "client_code": "AACI043592",
        "pin": "0135",
        "api_key": "YDbeLBqW",
        "totp_secret": "GT36L6XXQ67HVUHNMZXNNAQZFE",
    },
    {
        "account_name": "Bhumika",
        "client_code": "AACI063602",
        "pin": "0135",
        "api_key": "H7m9OODA",
        "totp_secret": "W5PP36UTOFOP6PEXRXBBM2HH5A",
    },
    {
        "account_name": "Rakesh Bansal",
        "client_code": "R297279",
        "pin": "0135",
        "api_key": "ErMVysb3",
        "totp_secret": "OGNEXOGC5IOPGG26HTIOTETY6U",
    },
    {
        "account_name": "Rakesh bansal HUF",
        "client_code": "AACN450426",
        "pin": "0135",
        "api_key": "9UIEQ3XH",
        "totp_secret": "24ORPLKAQ57ZJ4XDOKUXWR2CFI",
    },
    {
        "account_name": "Kashish Bansal",
        "client_code": "AAAC015110",
        "pin": "0135",
        "api_key": "NrAfSfRm",
        "totp_secret": "IEF3A2CBWDEJ4R2WSXZIEWQR34",
    },
    {
        "account_name": "Reena devi",
        "client_code": "R282962",
        "pin": "0135",
        "api_key": "8E1oTBdv",
        "totp_secret": "3RUYUPX4OFYW6UVIK5XE44TEDE",
    },
    {
        "account_name": "Pawan Kumar HUF",
        "client_code": "AACN448804",
        "pin": "0135",
        "api_key": "ks378IS0",
        "totp_secret": "FQ73WH6HXGJUWRHSG3ZIQLKUX4",
    },
    {
        "account_name": "Mukesh Kumar Narang HUF",
        "client_code": "AACN784607",
        "pin": "0135",
        "api_key": "j9YmEqu",
        "totp_secret": "BGTJAYGATBJT2O3GFWKI7JPFD4",
    },
    {
        "account_name": "Rukmani devi",
        "client_code": "AAAF368129",
        "pin": "0135",
        "api_key": "dDmDd1bS",
        "totp_secret": "7FXGBYZSJ6MHGIPC5S33SRCQEY",
    },
    {
        "account_name": "RIYA",
        "client_code": "R597315",
        "pin": "0135",
        "api_key": "57EvxIi7",
        "totp_secret": "PIOKVKIJIEJGOQJSFXDDBVQK6E",
    },
    {
        "account_name": "Seema devi",
        "client_code": "S54748898",
        "pin": "0135",
        "api_key": "yc8yTOIg",
        "totp_secret": "S3UYIELEZYRIG3YULOP3U7VSTA",
    },
    {
        "account_name": "RAJESH KUMAR",
        "client_code": "AACI064276",
        "pin": "0135",
        "api_key": "1M69CtnL",
        "totp_secret": "2QQWJEVIVG3FCA7YESVY5ATGPY",
    },
    {
        "account_name": "Diksha Bansal",
        "client_code": "D54745460",
        "pin": "0135",
        "api_key": "5M78l5lP",
        "totp_secret": "L7CN5TVSQXX7KVDGJLINA5FRZA",
    },
    {
        "account_name": "SANJU DEVI",
        "client_code": "S1136788",
        "pin": "0135",
        "api_key": "lK6vgf14",
        "totp_secret": "D3L52DH4VFPH65HNY3FRXJOCTY",
    },
    {
        "account_name": "Rajender Kumar",
        "client_code": "R53517293",
        "pin": "0135",
        "api_key": "FUv8x8TD",
        "totp_secret": "IIHBBS4GIY7H2BKNEZJSALYPMI",
    },
    {
        "account_name": "Rajender Kumar HUF",
        "client_code": "AACN450610",
        "pin": "0135",
        "api_key": "09uMRRcy",
        "totp_secret": "D4BPDB6WCLEWG3R33VPF6LPTRI",
    },
    {
        "account_name": "Narender Garg",
        "client_code": "AACH009422",
        "pin": "0135",
        "api_key": "F5eInGid",
        "totp_secret": "FCL2C6LMSDQNNBCNNLPOCLSZ4I",
    },
    {
        "account_name": "Narender Garg HUF",
        "client_code": "AACN450253",
        "pin": "0135",
        "api_key": "GvAkAXUG",
        "totp_secret": "G2UNVVSWPRCM435KKDUH7XMLIM",
    },
    {
        "account_name": "Pawan Kumar",
        "client_code": "AACM885301",
        "pin": "0135",
        "api_key": "9XlBd97D",
        "totp_secret": "4I5JUW64T5TWYSAPFCIY66HIXA",
    },
    {
        "account_name": "Peeyush Sharma",
        "client_code": "P63048331",
        "pin": "0135",
        "api_key": "q9TYE8Nh",
        "totp_secret": "GQTCQGHA6TYK6U2CUGEDONIQE4",
    },
    {
        "account_name": "Saroj devi",
        "client_code": "AACM885810",
        "pin": "0135",
        "api_key": "gRvZQHPU",
        "totp_secret": "AOK4MS3UWTJWF7MGB2VER3SJ7E",
    },
    {
        "account_name": "GOURAV DHURIA",
        "client_code": "AACM872342",
        "pin": "0135",
        "api_key": "X5Rk6caa",
        "totp_secret": "DSIFI4B5A444WWYHQ5MY5PW64A",
    },
    {
        "account_name": "Rajni Bala",
        "client_code": "AACM873426",
        "pin": "0135",
        "api_key": "roARzAl4",
        "totp_secret": "TAOJRZNVZULFXURLTW5LD24XVQ",
    },
    {
        "account_name": "Vikram Suthar",
        "client_code": "AACM87426",
        "pin": "0135",
        "api_key": "roARzA4",
        "totp_secret": "TAOJRZNVZULFXURLTW5LD24XVQ",
    },
    {
        "account_name": "Raju Udaniya",
        "client_code": "AACM87426",
        "pin": "0135",
        "api_key": "roARzA4",
        "totp_secret": "TAOJRZNVZULFXURLTW5LD24XVQ",
    },
    {
        "account_name": "Baby",
        "client_code": "AACM848071",
        "pin": "0135",
        "api_key": "lfLFt7Zv",
        "totp_secret": "3YFQWFNHGDJEA2GUVIPVFMHGNU",
    },
    {
        "account_name": "Sandeep Singharia",
        "client_code": "AACM846943",
        "pin": "0135",
        "api_key": "QGKhu12E",
        "totp_secret": "LM67JG63XYRMBLX57BYBUMAKOM",
    },
    {
        "account_name": "Mukesh Kumar Narang",
        "client_code": "AACM853089",
        "pin": "0135",
        "api_key": "d5ISqdtR",
        "totp_secret": "HVKJZDRJUQID5VXY7DUTTIGZGY",
    },
    {
        "account_name": "Monika devi",
        "client_code": "AAAH788468",
        "pin": "0135",
        "api_key": "rLmVDkmm",
        "totp_secret": "5KAE2CYQZJQPXXZJO32EC4Z52A",
    },
    {
        "account_name": "ATMA RAM",
        "client_code": "AAAA932378",
        "pin": "0135",
        "api_key": "3HilUF2v",
        "totp_secret": "XTNPQ2BH43PZUJXL3D2OF2KERQ",
    },
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
                settled_qty = int(item.get("quantity", 0) or 0)
t1_qty = int(item.get("t1quantity", 0) or item.get("t1Quantity", 0) or 0)
qty = settled_qty if (settled_qty > 0 and settled_qty == t1_qty) else (settled_qty + t1_qty)
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
