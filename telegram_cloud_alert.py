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
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "9166325450:AAH7Zh2oGZ-VBtwlkINngFaZo6PWre-doUM")
TELEGRAM_CHAT_ID = str(os.environ.get("TELEGRAM_CHAT_ID", "1448796295"))
STATE_FILE = "holdings_state.json"
GAIN_STEP_PCT = 5  # Alerts at every +5% milestone (+5%, +10%, +15%...)

# OneDrive Cloud API Credentials (GitHub Secrets se aayenge)
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


def fetch_single_account(acc):
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


# ================= ONEDRIVE EXCEL SYNC FUNCTIONS =================
def get_onedrive_access_token():
    """Refresh token se naya Microsoft Graph access token leta hai."""
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
        resp = requests.post(url, data=data, timeout=15).json()
        return resp.get("access_token")
    except Exception as e:
        print(f"OneDrive token error: {e}")
        return None


def sync_excel_to_onedrive(all_excel_rows):
    """OneDrive se .xlsx download karke openpyxl se update karta hai aur wapas upload karta hai."""
    if not all_excel_rows:
        return

    token = get_onedrive_access_token()
    if not token:
        print("OneDrive secrets missing; skipping cloud Excel update.")
        return

    headers = {"Authorization": f"Bearer {token}"}
    file_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{ONEDRIVE_FILE_PATH}:/content"

    # 1. OneDrive se current Excel file download karein
    dl_resp = requests.get(file_url, headers=headers, timeout=30)
    if not dl_resp.ok:
        print(f"OneDrive download failed ({dl_resp.status_code}): {dl_resp.text}")
        return

    wb = openpyxl.load_workbook(io.BytesIO(dl_resp.content))

    # 2. Grand Totals aur Rows tayyar karein
    total_invested = sum(float(r[6]) for r in all_excel_rows)
    total_current = sum(float(r[7]) for r in all_excel_rows)
    total_pnl = sum(float(r[8]) for r in all_excel_rows)
    ist_time = get_ist_now().strftime("%H:%M:%S")

    total_row = [
        "TOTAL", "", "", "", "", "",
        round(total_invested, 2), round(total_current, 2), round(total_pnl, 2), ist_time
    ]

    # 3. Demat_Holdings_(Equity_ETFs) update karein (Row 7 se)
    sheet = wb["Demat_Holdings_(Equity_ETFs)"]
    for row in sheet[f"A{START_ROW}:J500"]:
        for cell in row:
            cell.value = None

    for r_idx, row_vals in enumerate(all_excel_rows + [total_row], start=START_ROW):
        for c_idx, val in enumerate(row_vals, start=1):
            sheet.cell(row=r_idx, column=c_idx, value=val)

    # 4. Dashboard_Summary update karein (C12, D12 & Row 15+ Matrix)
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

        matrix_headers = ["Stock/Name", "TOTAL QTY"] + client_names
        for c_idx, h in enumerate(matrix_headers, start=1):
            dash.cell(row=DASH_HEADER_ROW, column=c_idx, value=h)

        for r_idx, stock in enumerate(sorted(stock_totals.keys()), start=DASH_START_ROW):
            dash.cell(row=r_idx, column=1, value=stock)
            dash.cell(row=r_idx, column=2, value=stock_totals[stock])
            for c_idx, client in enumerate(client_names, start=3):
                qty = stock_data[stock].get(client, 0)
                dash.cell(row=r_idx, column=c_idx, value=(qty if qty > 0 else None))

        dash.column_dimensions["A"].width = 24
        dash.column_dimensions["B"].width = 16
        dash.column_dimensions["C"].width = 22
        dash.column_dimensions["D"].width = 22

    # 5. Updated Excel file wapas OneDrive cloud par upload karein
    out_stream = io.BytesIO()
    wb.save(out_stream)
    wb.close()

    up_headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }
    up_resp = requests.put(file_url, headers=up_headers, data=out_stream.getvalue(), timeout=30)
    if up_resp.ok:
        print(f"SUCCESS: OneDrive Excel updated at {ist_time} IST!")
    else:
        print(f"OneDrive upload failed ({up_resp.status_code}): {up_resp.text}")
# =================================================================


def main():
    prev_state = {}
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            prev_state = json.load(f)

    curr_state = {}
    ok_codes = set()
    all_excel_rows = []
    ist_time = get_ist_now().strftime("%H:%M:%S")

    # 1. Saare accounts se holdings fetch karein (Alerts + Excel dono ke liye)
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
                # Excel sheet ke liye 10 columns ki row tayyar karein
                all_excel_rows.append([
                    acc["account_name"].strip(),
                    acc["client_code"].strip(),
                    x["stock"],
                    x["qty"],
                    x["avg"],
                    x["ltp"],
                    x["invested"],
                    x["current"],
                    x["pnl"],
                    ist_time
                ])
        time.sleep(0.5)

    if not ok_codes:
        return

    # 2. Har 15 minute mein OneDrive Cloud Excel file update karein (00, 15, 30, 45)
    ist_now = get_ist_now()
    if ist_now.minute % 15 < 2:
        sync_excel_to_onedrive(all_excel_rows)

    # 3. Buy/Sell aur +5% / +10% / +15% Gain Alerts check karein (Har 2 minute)
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
        send_telegram_alert(f"🚨 <b>Portfolio Alert ({ist_now.strftime('%d-%b %I:%M %p')})</b>\n\n" + "\n\n".join(all_alerts))

    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(curr_state, f, indent=2)


if __name__ == "__main__":
    main()
