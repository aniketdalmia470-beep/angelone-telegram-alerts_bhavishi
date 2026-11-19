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

ONEDRIVE_CLIENT_ID = os.environ.get("ONEDRIVE_CLIENT_ID", "").strip().strip('"')
ONEDRIVE_CLIENT_SECRET = os.environ.get("ONEDRIVE_CLIENT_SECRET", "").strip().strip('"')
ONEDRIVE_REFRESH_TOKEN = os.environ.get("ONEDRIVE_REFRESH_TOKEN", "").strip().strip('"')
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
    {"account_name": "Peeyush Sharma", "client_code
