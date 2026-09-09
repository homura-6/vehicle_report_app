from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse, Response
import uvicorn
from datetime import datetime, timedelta
import pandas as pd
import os
import json
from io import BytesIO

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

app = FastAPI()

EXCEL_PATH = "Database_Model_X_Price.xlsx"
CASES_JSON_PATH = "Database_Saved_Cases.json"
CASES_DB_PATH = "Database_Saved_Cases.xlsx"


# 辅助函数：读取和写入 JSON 数据库
def load_json_db():
    if not os.path.exists(CASES_JSON_PATH):
        return {}
    try:
        with open(CASES_JSON_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_json_db(data_dict):
    with open(CASES_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(data_dict, f, ensure_ascii=False, indent=2)


def parse_and_format_date(date_str: str) -> str:
    if not date_str:
        return ""
    for fmt in ("%Y-%m-%d", "%d %b %Y"):
        try:
            dt = datetime.strptime(date_str.strip(), fmt)
            return dt.strftime("%d %b %Y")
        except ValueError:
            pass
    return date_str


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.join(BASE_DIR, "fonts")
LOGO_PATH = os.path.join(BASE_DIR, "assets", "tc.png")

# ---------------------------------------------------------------------------
# 字体：优先 Arimo，找不到就退回 Helvetica（视觉上非常接近 Arial）
# 若要用真正的 Arimo/Arial，把对应的 .ttf 放进 /fonts 文件夹：
#   fonts/Arimo-Regular.ttf , fonts/Arimo-Bold.ttf
#   fonts/Arial.ttf         , fonts/Arial-Bold.ttf
# ---------------------------------------------------------------------------
FONT_REG = "Helvetica"
FONT_BOLD = "Helvetica-Bold"


def _register_fonts():
    global FONT_REG, FONT_BOLD
    candidates = [
        ("Arimo", "Arimo-Regular.ttf", "Arimo-Bold.ttf"),
        ("Arial", "Arial.ttf", "Arial-Bold.ttf"),
    ]
    for name, reg_file, bold_file in candidates:
        reg_path = os.path.join(FONT_DIR, reg_file)
        bold_path = os.path.join(FONT_DIR, bold_file)
        if os.path.exists(reg_path) and os.path.exists(bold_path):
            try:
                pdfmetrics.registerFont(TTFont(name, reg_path))
                pdfmetrics.registerFont(TTFont(name + "-Bold", bold_path))
                FONT_REG = name
                FONT_BOLD = name + "-Bold"
                return
            except Exception:
                continue


_register_fonts()

# ---------------------------------------------------------------------------
# 公司固定信息
# ---------------------------------------------------------------------------
COMPANY_NAME = "T & C AUTOMOBILE ASSESSORS"
COMPANY_ADDR_LINE = "117 BUKIT MERAH VIEW #12-173 SINGAPORE 151117"
COMPANY_CONTACT_LINE = "HP : 97667875   Email : tcautomobileassessors@yahoo.com.sg"
COMPANY_REG_LINE = "Co.Reg.53256293C"

PAGE_W, PAGE_H = A4
MARGIN_L = 50
MARGIN_R = 50
MARGIN_TOP = 10
MARGIN_BOTTOM = 45
CONTENT_W = PAGE_W - MARGIN_L - MARGIN_R
PARTS_ROWS_PER_PAGE = 37


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------
def _num(v):
    if v is None:
        return 0.0
    try:
        return float(str(v).replace("S$", "").replace(",", "").strip())
    except Exception:
        return 0.0


def fmt_money(x):
    try:
        x = float(x)
    except Exception:
        x = 0.0
    return f"{x:,.2f}"


def parse_case_date(s):
    if not s:
        return None
    s = str(s).strip()
    for fmt in ("%d %b %Y", "%d %B %Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def full_date(s):
    d = parse_case_date(s)
    if not d:
        return s or ""
    return d.strftime("%d %B %Y")


def minus_one_day_full(s):
    d = parse_case_date(s)
    if not d:
        return ""
    return (d - timedelta(days=1)).strftime("%d %B %Y")

def compute_date_of_request(inspection_str, accident_str):
    """Date Of Request = Date Of Inspection - 1 天；
    如果算出来的日期跟 Date Of Accident 同一天，就改回显示 Date Of Inspection。"""
    insp = parse_case_date(inspection_str)
    if not insp:
        return ""
    candidate = insp - timedelta(days=1)
    acc = parse_case_date(accident_str)
    if acc and candidate.date() == acc.date():
        return insp.strftime("%d %B %Y")
    return candidate.strftime("%d %B %Y")


def job_no_compact(job_no):
    return (job_no or "").replace("/", "")


# ---------------------------------------------------------------------------
# 与 main2.py 前端 JS 完全对应的业务规则
# ---------------------------------------------------------------------------
REMARK_RULES = {
    "rp": {"wording": "REPAIR", "type": "zero"},
    "sv": {"wording": "SERVICE", "type": "zero"},
    "rj": {"wording": "REJECTED", "type": "zero"},
    "na": {"wording": "NA", "type": "zero"},
    "be": {"wording": "BENT", "type": "full"},
    "de": {"wording": "DENT", "type": "full"},
    "bu": {"wording": "BUCKLED", "type": "full"},
    "ne": {"wording": "NECESSARY", "type": "full"},
    "br": {"wording": "BROKEN", "type": "full"},
    "cr": {"wording": "CRACKED", "type": "full"},
    "ma": {"wording": "MALFUNCTION", "type": "full"},
    "df": {"wording": "DEFORMED", "type": "full"},
    "dp": {"wording": "DEPLOYED", "type": "full"},
    "ja": {"wording": "JAMMED", "type": "full"},
    "to": {"wording": "TORN", "type": "full"},
    "ru": {"wording": "REUSE", "type": "zero"},
    "gl": {"wording": "GLAZED", "type": "full"},
    "mi": {"wording": "MISSING", "type": "full"},
    "we": {"wording": "WEAK", "type": "full"},
    "cu": {"wording": "CUT", "type": "full"},
    "sc": {"wording": "SCRATCHED", "type": "full"},
    "1m": {"wording": "1PC MALFUNCTION", "type": "m_pcs", "pcs": 1},
    "2m": {"wording": "2PCS MALFUNCTION", "type": "m_pcs", "pcs": 2},
    "3m": {"wording": "3PCS MALFUNCTION", "type": "m_pcs", "pcs": 3},
    "4m": {"wording": "4PCS MALFUNCTION", "type": "m_pcs", "pcs": 4},
    "5m": {"wording": "5PCS MALFUNCTION", "type": "m_pcs", "pcs": 5},
    "c1": {"wording": "Cut/10%", "type": "cut", "percent": 0.90},
    "c2": {"wording": "Cut/20%", "type": "cut", "percent": 0.80},
    "c3": {"wording": "Cut/30%", "type": "cut", "percent": 0.70},
    "c4": {"wording": "Cut/40%", "type": "cut", "percent": 0.60},
    "c5": {"wording": "Cut/50%", "type": "cut", "percent": 0.50},
    "c6": {"wording": "Cut/60%", "type": "cut", "percent": 0.40},
    "c7": {"wording": "Cut/70%", "type": "cut", "percent": 0.30},
    "c8": {"wording": "Cut/80%", "type": "cut", "percent": 0.20},
    "c9": {"wording": "Cut/90%", "type": "cut", "percent": 0.10},
}

TYRE_POSITIONS = {
    "2": ["Front", "Rear"],
    "4": ["Front RHS", "Front LHS", "Rear RHS", "Rear LHS"],
    "6": ["Front RHS", "Front LHS", "Rear Outer RHS", "Rear Inner RHS", "Rear Inner LHS", "Rear Outer LHS"],
    "8": ["Front RHS", "Front LHS", "Mid RHS", "Mid LHS", "Rear Outer RHS", "Rear Inner RHS", "Rear Inner LHS", "Rear Outer LHS"],
    "10": ["Front RHS", "Front LHS", "Mid Outer RHS", "Mid Inner RHS", "Mid Inner LHS", "Mid Outer LHS", "Rear Outer RHS", "Rear Inner RHS", "Rear Inner LHS", "Rear Outer LHS"],
}

PAINT_CODE_LABELS = {
    "1": "Two-Coat Metallic Paint",
    "2": "Three-Stage Pearlescent Paint",
    "3": "Single-Stage Solid Paint",
    "4": "Two-Coat Solid Paint",
}

LABOUR_ITEMS = [
    "To remove, cut out damage portion, jack out, straighten, panel beating, welding, align and renew replaced parts.",
    "To check wirings and lightings.",
    "To remove & install reverse sensor.",
    "To remove & refix tailgate fittings to facilitate repair.",
    "To remove & refix rear upholstery, garnish and attachment.",
    "To supply sealant and to seal off all weld spot seam and gaps.",
    "To spray anti-rust coating.",
    "To remove and install rear windscreen glass.",
    "To remove and install front windscreen glass.",
    "To remove and install rear right quarter glass.",
    "To remove and install rear left quarter glass.",
    "To remove and install front right quarter glass.",
    "To remove and install front left quarter glass.",
    "To recalibrate right power gate sliding door and use diagnostic computer to reset the system.",
    "To recalibrate left power gate sliding door and use diagnostic computer to reset the system.",
    "To change radiator and related parts , and to perform coolant leakage test .",
    "To change a/c parts and to recharge a/c gas.",
    "To conduct wheel alignment.",
    "To transfer door glass, trimboard and other components to replacement door.",
    "To remove and install rear suspension unit.",
    "To remove and install front suspension unit.",
    "To remove and install front left suspension unit.",
    "To remove and install front right suspension unit.",
    "To remove and install rear left suspension unit.",
    "To remove and install rear right suspension unit.",
    "To remove and install exhaust pipes.",
]


def compute_part_row(raw):
    qty = _num(raw.get("qty"))
    nett = _num(raw.get("nett"))
    global_pct = _num(raw.get("_global_pct", 100))
    code_l = (raw.get("code") or "").strip().lower()
    est = nett * qty * (global_pct / 100.0)

    if code_l in REMARK_RULES:
        rule = REMARK_RULES[code_l]
        wording = rule["wording"]
        t = rule["type"]
        if t == "zero":
            mult = 0.0
        elif t == "full":
            mult = 1.0
        elif t == "m_pcs":
            mult = (1.0 / qty * rule["pcs"]) if qty > 0 else 0.0
        elif t == "cut":
            mult = 1.0 - rule["percent"]
        else:
            mult = 1.0
    elif code_l == "":
        wording = "-"
        mult = 1.0
    else:
        wording = "UNKNOWN CODE"
        mult = 0.0

    ass = est * mult
    return {
        "sn": None,
        "qty": raw.get("qty"),
        "name": raw.get("name") or "",
        "code": (raw.get("code") or "").upper(),
        "condition": wording,
        "est": est,
        "ass": ass,
    }


def compute_misc_row(raw):
    code_l = (raw.get("code") or "").strip().lower()
    if code_l in REMARK_RULES:
        wording = REMARK_RULES[code_l]["wording"]
    elif code_l == "":
        wording = "-"
    else:
        wording = "UNKNOWN CODE"
    return {
        "qty": raw.get("qty"),
        "desc": raw.get("desc") or "",
        "code": (raw.get("code") or "").upper(),
        "condition": wording,
        "est": _num(raw.get("est")),
        "ass": _num(raw.get("ass")),
    }


# ---------------------------------------------------------------------------
# Header / Footer / Subheader
# ---------------------------------------------------------------------------
def draw_header(c):
    top_y = PAGE_H - MARGIN_TOP
    logo_size = 80
    try:
        img = ImageReader(LOGO_PATH)
        c.drawImage(img, MARGIN_L, top_y - logo_size, width=logo_size, height=logo_size,
                    preserveAspectRatio=True, mask='auto')
    except Exception:
        pass
    c.setFont(FONT_BOLD, 25)
    c.drawString(MARGIN_L + logo_size + 18, top_y - logo_size / 2 - 8, COMPANY_NAME)
    return top_y - logo_size - 12


def draw_footer_report(c, page_num, total_pages):
    c.setFont(FONT_REG, 10)
    c.drawCentredString(PAGE_W / 2, MARGIN_BOTTOM - 18, f"Page {page_num} of {total_pages}")


def draw_footer_invoice(c):
    y = MARGIN_BOTTOM + 46
    c.setLineWidth(1)
    c.setFont(FONT_REG, 8)
    c.drawCentredString(PAGE_W / 2, y - 14, COMPANY_ADDR_LINE)
    c.drawCentredString(PAGE_W / 2, y - 24, COMPANY_CONTACT_LINE)
    c.drawCentredString(PAGE_W / 2, y - 34, COMPANY_REG_LINE)
    c.setFont(FONT_REG, 10)
    c.drawCentredString(PAGE_W / 2, MARGIN_BOTTOM - 18, "Page 1 of 1")


def draw_subheader(c, y, annex_label, vehicle_no, report_no):
    y -= 8
    c.setFont(FONT_REG, 8)
    c.drawRightString(PAGE_W - MARGIN_R, y, annex_label)
    y -= 18
    c.setFont(FONT_BOLD, 12)
    c.drawString(MARGIN_L, y, "ADJUSTMENT ON REPAIR COST & REPLACEMENT OF PARTS")
    c.line(MARGIN_L, y - 2, MARGIN_L + c.stringWidth(
        "ADJUSTMENT ON REPAIR COST & REPLACEMENT OF PARTS", FONT_BOLD, 12), y - 2)
    y -= 20
    c.setFont(FONT_REG, 10)
    line = f"Vehicle No. :  {vehicle_no}"
    c.drawCentredString(PAGE_W / 2 - 150, y, line)
    c.drawCentredString(PAGE_W / 2 + 110, y, f"Report No. :  {report_no}")
    y -= 12
    c.setLineWidth(1.8)
    c.line(MARGIN_L, y, PAGE_W - MARGIN_R, y)
    c.setLineWidth(1)
    y -= 20
    return y


def section_title(c, y, text, size=12):
    c.setFont(FONT_BOLD, size)
    c.drawString(MARGIN_L, y, text)
    c.line(MARGIN_L, y - 2, MARGIN_L + c.stringWidth(text, FONT_BOLD, size), y - 2)
    return y - (size + 10)


def label_bold(c, size=10):
    c.setFont(FONT_BOLD, size)


def label_reg(c, size=10):
    c.setFont(FONT_REG, size)


# ---------------------------------------------------------------------------
# INVOICE
# ---------------------------------------------------------------------------
def generate_invoice_pdf(case: dict) -> bytes:
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    s1 = case.get("step1Data", {}) or {}
    job_no = case.get("job_no", "") or ""
    inv_no = job_no_compact(job_no)

    y = draw_header(c)

    # [INVOICE] 14pt bold underline, centered
    y -= 8
    c.setFont(FONT_BOLD, 16)
    title = "INVOICE"
    tw = c.stringWidth(title, FONT_BOLD, 14)
    tx = (PAGE_W - tw) / 2
    c.drawString(tx, y, title)
    c.line(tx - 1, y - 3, tx + 8 + tw, y - 3)

    y -= 22
    box_top = y

    owner_name = s1.get("ownerName", "")
    owner_addr_lines = [s1.get("ownerAddr1", ""), s1.get("ownerAddr2", ""), s1.get("ownerAddr3", "")]
    owner_addr_lines = [ln for ln in owner_addr_lines if ln]

    box_w = 270
    line_h = 14
    box_h = 16 + line_h * len(owner_addr_lines) + 14
    c.roundRect(MARGIN_L, box_top - box_h, box_w, box_h, 10)
    c.setFont(FONT_REG, 10)
    ty = box_top - 18
    c.drawString(MARGIN_L + 12, ty, owner_name)
    for ln in owner_addr_lines:
        ty -= line_h
        c.drawString(MARGIN_L + 12, ty, ln)

    # right column: INV NO / DATE
    rx_label = MARGIN_L + box_w + 80
    ry = box_top - 18
    c.setFont(FONT_REG, 10)
    c.drawString(rx_label, ry, "INV NO.  :  " + inv_no)
    ry -= 16 * 3
    c.drawString(rx_label, ry, "DATE   :  " + full_date(case.get("finishDate", "")))

    y = box_top - box_h - 26
    y -= 22
    # VEHICLE NO / JOB REFERENCE NO / ACCIDENT DATE / SURVEY DATE  (centered block)
    c.setFont(FONT_REG, 10)
    rows = [
        ("VEHICLE NO.", case.get("vehicleNo", "")),
        ("JOB REFERENCE NO.", job_no),
        ("ACCIDENT DATE", full_date(case.get("accidentDate", ""))),
        ("SURVEY DATE", full_date(case.get("inspectionDate", ""))),
    ]
    label_right_x = PAGE_W / 2 + 10
    for label, val in rows:
        c.drawRightString(label_right_x, y, label + "  :")
        c.drawString(label_right_x + 8, y, str(val))
        y -= 14

    y -= 28
    c.line(MARGIN_L, y, PAGE_W - MARGIN_R, y)
    y -= 14

    # Table header
    c.setFont(FONT_BOLD, 12)
    c.drawString(MARGIN_L, y, "S/N")
    c.drawString(MARGIN_L + 50, y, "DESCRIPTION")
    c.drawRightString(PAGE_W - MARGIN_R, y, "AMOUNT")
    y -= 4
    c.line(MARGIN_L, y - 1, PAGE_W - MARGIN_R, y - 1)
    y -= 15

    photo_qty = case.get("photoQty", "") or "0"
    survey_fee = _num(case.get("surveyFees"))
    transport = _num(case.get("transportation"))
    photo_cost = _num(case.get("photographsCost"))
    total_survey_fee = _num(case.get("totalSurveyFee"))

    items = [
        ("1", "Survey Fees", survey_fee),
        ("2", "Transportation", transport),
        ("3", f"Photographs ($2) per copies :    {photo_qty}", photo_cost),
    ]

    c.setFont(FONT_REG, 10)
    for sn, desc, amt in items:
        c.drawString(MARGIN_L + 5, y, sn)
        c.drawString(MARGIN_L + 50, y, desc)
        c.drawString(PAGE_W - MARGIN_R - 80, y, "S$")
        c.drawRightString(PAGE_W - MARGIN_R, y, fmt_money(amt))
        y -= 14

    y -= 0
    c.line(PAGE_W - MARGIN_R - 95, y, PAGE_W - MARGIN_R, y)
    y -= 12
    c.setFont(FONT_BOLD, 10)
    c.drawRightString(PAGE_W - MARGIN_R - 100, y, "TOTAL AMOUNT :")
    c.drawString(PAGE_W - MARGIN_R - 80, y, "S$")
    c.drawRightString(PAGE_W - MARGIN_R, y, fmt_money(total_survey_fee))
    y -= 5
    c.line(PAGE_W - MARGIN_R - 95, y, PAGE_W - MARGIN_R, y)
    y -= 3
    c.line(PAGE_W - MARGIN_R - 95, y, PAGE_W - MARGIN_R, y)

    # Notes block
    y -= 90
    c.setFont(FONT_REG, 10)
    c.drawString(MARGIN_L, y, "Notes:")
    y -= 16
    c.drawString(MARGIN_L, y,
                 'All cheque payment should be "Crossed" and make payable to " T & C Automobile Assessors "')

    y -= 34
    c.drawString(MARGIN_L, y, "T&C AUTOMOBILE ASSESSORS")

    y -= 130
    c.setLineWidth(2.0)
    c.line(MARGIN_L, y, MARGIN_L + 160, y)
    c.setLineWidth(1.0)

    draw_footer_invoice(c)
    c.showPage()
    c.save()
    buf.seek(0)
    return buf.read()


# ---------------------------------------------------------------------------
# SURVEY REPORT
# ---------------------------------------------------------------------------
def _draw_wrapped_or_ellipsis(c, text, x, y, max_width, font, size):
    """单行显示；超出宽度则用省略号截断，保证行高固定。"""
    c.setFont(font, size)
    text = str(text) if text is not None else ""
    if c.stringWidth(text, font, size) <= max_width:
        c.drawString(x, y, text)
        return
    ell = "..."
    ell_w = c.stringWidth(ell, font, size)
    out = ""
    for ch in text:
        if c.stringWidth(out + ch, font, size) + ell_w > max_width:
            break
        out += ch
    c.drawString(x, y, out + ell)


def _page1(c, case, page_num, total_pages):
    s1 = case.get("step1Data", {}) or {}
    job_no = case.get("job_no", "") or ""

    y = draw_header(c)

    # top-left: owner + C/O workshop ; top-right: REF / Date
    c.setFont(FONT_REG, 10)
    left_x = MARGIN_L
    top_y = y
    c.drawString(left_x, top_y, s1.get("ownerName", ""))
    top_y -= 14
    c.drawString(left_x, top_y, "C/O  " + (s1.get("wsName", "") or ""))
    for ln in [s1.get("wsAddr1", ""), s1.get("wsAddr2", ""), s1.get("wsAddr3", "")]:
        top_y -= 14
        if ln:
            c.drawString(left_x, top_y, ln)

    ry = y
    c.drawRightString(PAGE_W - MARGIN_R, ry, "REF :  " + job_no)
    ry -= 52
    c.drawRightString(PAGE_W - MARGIN_R, ry, "Date :  " + full_date(case.get("finishDate", "")))

    y = min(top_y, ry) - 12

    # Title
    c.setFont(FONT_BOLD, 12)
    title = "VEHICLE SURVEY REPORT"
    tw = c.stringWidth(title, FONT_BOLD, 12)
    tx = (PAGE_W - tw) / 2
    c.drawString(tx, y, title)
    c.line(tx, y - 2, tx + tw, y - 2)
    y -= 12
    c.line(MARGIN_L, y, PAGE_W - MARGIN_R, y)
    y -= 18

    # REFERENCE
    y = section_title(c, y - -3, "REFERENCE", 10)
    ws_lines = [s1.get("wsName", ""), s1.get("wsAddr1", ""), s1.get("wsAddr2", ""), s1.get("wsAddr3", "")]
    ws_lines = [ln for ln in ws_lines if ln]
    ref_rows = [
        ("Request By", "Workshop , On Behalf Of Owner"),
        ("Date Of Request", compute_date_of_request(case.get("inspectionDate", ""), case.get("accidentDate", ""))),
        ("Date Of Accident", full_date(case.get("accidentDate", ""))),
        ("Date Of Inspection", full_date(case.get("inspectionDate", ""))),
    ]
    label_x_end = MARGIN_L + 130
    c.setFont(FONT_REG, 10)
    for label, val in ref_rows:
        c.drawRightString(label_x_end, y, label + "  :")
        c.drawString(label_x_end + 8, y, str(val))
        y -= 15
    c.drawRightString(label_x_end, y, "Inspected At  :")
    c.drawString(label_x_end + 8, y, ws_lines[0] if ws_lines else "")
    for ln in ws_lines[1:]:
        y -= 15
        c.drawString(label_x_end + 8, y, ln)
    y -= 8
    c.line(MARGIN_L, y, PAGE_W - MARGIN_R, y)
    y -= 18

    # PARTICULARS OF VEHICLE
    y = section_title(c, y - -3, "PARTICULARS OF VEHICLE", 10)
    left_label_end = MARGIN_L + 115
    right_label_start = PAGE_W / 2 + 20
    right_label_end = right_label_start + 95
    row_y = y
    left_rows = [
        ("Vehicle Reg. No.", case.get("vehicleNo", "")),
        ("Year Make", s1.get("yearMake", "")),
        ("Engine No.", s1.get("engineNo", "") or "-"),
        ("Engine Capacity", (str(s1.get("engineCap", "")) + "   cc") if s1.get("engineCap") else "-"),
        ("Power Rating", (str(s1.get("powerRating", "")) + "   kW") if s1.get("powerRating") else "-"),
    ]
    right_rows = [
        ("Make & Model", case.get("makeModel", "")),
        ("Colour", s1.get("colour", "") or "-"),
        ("Chassis No.", s1.get("chassisNo", "") or "-"),
        ("Motor No.", s1.get("motorNo", "") or "-"),
    ]
    c.setFont(FONT_REG, 10)
    yy = row_y
    for label, val in left_rows:
        c.drawRightString(left_label_end, yy, label + "  :")
        c.drawString(left_label_end + 8, yy, str(val))
        yy -= 15
    left_bottom = yy
    yy = row_y
    # Make & Model 可能较长，预留自动换行空间
    label, val = right_rows[0]
    c.drawRightString(right_label_end, yy, label + "  :")
    max_w = (PAGE_W - MARGIN_R) - (right_label_end + 8)
    val = str(val)
    if c.stringWidth(val, FONT_REG, 10) > max_w:
        words = val.split(" ")
        line1, line2 = "", ""
        for w in words:
            if c.stringWidth((line1 + " " + w).strip(), FONT_REG, 10) <= max_w:
                line1 = (line1 + " " + w).strip()
            else:
                line2 = (line2 + " " + w).strip()
        c.drawString(right_label_end + 8, yy, line1)
        yy -= 15
        c.drawString(right_label_end + 8, yy, line2)
    else:
        c.drawString(right_label_end + 8, yy, val)
    yy -= 15
    for label, val in right_rows[1:]:
        c.drawRightString(right_label_end, yy, label + "  :")
        c.drawString(right_label_end + 8, yy, str(val))
        yy -= 15
    right_bottom = yy

    y = min(left_bottom, right_bottom) - 1
    y -= -5
    c.line(MARGIN_L, y, PAGE_W - MARGIN_R, y)
    y -= 18

    # CONDITION OF VEHICLE AND TYRES
    y = section_title(c, y - - 3, "CONDITION OF VEHICLE AND TYRES", 10)
    mileage_val = s1.get("mileage", "")
    try:
        mileage_disp = f"{float(str(mileage_val).replace(',', '')):,.0f}" if mileage_val else "-"
    except Exception:
        mileage_disp = str(mileage_val)
    yy = y
    c.setFont(FONT_REG, 10)
    cond_left = [("General condition", "Good"), ("Brakes", "Serviceable"), ("Steering", "Serviceable")]
    cond_right = [("Mileage", mileage_disp + "   KM"), ("Modification", "None"), ("Handbrake", "Serviceable")]
    for label, val in cond_left:
        c.drawRightString(left_label_end, yy, label + "  :")
        c.drawString(left_label_end + 8, yy, val)
        yy -= 15
    left_bottom = yy
    yy = y
    for label, val in cond_right:
        c.drawRightString(right_label_end, yy, label + "  :")
        c.drawString(right_label_end + 8, yy, val)
        yy -= 15
    right_bottom = yy
    y = min(left_bottom, right_bottom) - 6

    wheel_count = str(s1.get("wheelCount", "") or "")
    positions = TYRE_POSITIONS.get(wheel_count, [])
    tyres = s1.get("tyres", []) or []
    if positions:
        col1_x, col2_x, col3_x, col4_x = MARGIN_L + 50, MARGIN_L + 150, MARGIN_L + 250, MARGIN_L + 350
        c.setFont(FONT_BOLD, 10)
        c.drawString(col1_x, y, "Tyres")
        c.drawString(col2_x, y, "Make")
        c.drawString(col3_x, y, "Size")
        c.drawString(col4_x, y, "Balance (mm)")
        c.line(col1_x, y - 3, col1_x + 365, y - 3)
        y -= 16
        c.setFont(FONT_REG, 10)
        for i, pos in enumerate(positions):
            t = tyres[i] if i < len(tyres) else {}
            c.drawString(col1_x, y - 1, pos)
            c.drawString(col2_x, y - 1, str(t.get("brand", "")))
            c.drawString(col3_x, y - 1, str(t.get("size", "")))
            c.drawString(col4_x + 30, y, str(t.get("mm", "")))
            y -= 14
    y -= -3
    c.line(MARGIN_L, y, PAGE_W - MARGIN_R, y)
    y -= 18

    # DESCRIPTION OF DAMAGES
    y = section_title(c, y - -3, "DESCRIPTION OF DAMAGES", 10)
    damaged = s1.get("damagedPos", "") or ""
    c.setFont(FONT_REG, 10)
    prefix = "The vehicle sustained damages at the "
    suffix = " portion."
    c.drawString(MARGIN_L, y, prefix)
    px = MARGIN_L + c.stringWidth(prefix, FONT_REG, 10)
    c.setFont(FONT_BOLD, 10)
    c.drawString(px, y, damaged)
    px2 = px + c.stringWidth(damaged, FONT_BOLD, 10)
    c.setFont(FONT_REG, 10)
    c.drawString(px2, y, suffix)
    y -= 15
    c.drawString(MARGIN_L, y, "(For information of damages , please refer to Parts / Labour / Photographs attached)")
    y -= 12
    c.line(MARGIN_L, y, PAGE_W - MARGIN_R, y)
    y -= 18

    # INSTRUCTION
    y = section_title(c, y, "INSTRUCTION", 10)
    part_a = 'This survey was conducted entirely on a "'
    part_b = "WITHOUT PREJUDICE"
    part_c = '" basis , and we have not authorised '
    line2 = "any repair ."

    x = MARGIN_L
    c.setFont(FONT_REG, 10)
    c.drawString(x, y, part_a)
    x += c.stringWidth(part_a, FONT_REG, 10)
    c.setFont(FONT_BOLD, 10)
    c.drawString(x, y, part_b)
    x += c.stringWidth(part_b, FONT_BOLD, 10)
    c.setFont(FONT_REG, 10)
    c.drawString(x, y, part_c)

    y -= 14
    c.drawString(MARGIN_L, y, line2)

    draw_footer_report(c, page_num, total_pages)
    c.showPage()


def _table_header(c, y, cols):
    """cols: list of (text, x, align, width)"""
    c.setFont(FONT_BOLD, 10)
    for text, x, align, width in cols:
        if align == "center":
            c.drawCentredString(x + width / 2, y, text)
        elif align == "right":
            c.drawRightString(x + width, y, text)
        else:
            c.drawString(x, y, text)
    c.line(MARGIN_L, y - 4, PAGE_W - MARGIN_R, y - 4)
    return y - 20


def _items_page(c, case, section_title_text, annex_label, rows, page_num, total_pages,
                 is_last_page, totals=None, show_discount=False):
    """PARTS / MISCELLANEOUS ITEMS 的表格页（每页最多 37 行）"""
    y = draw_header(c)
    y = draw_subheader(c, y, annex_label, case.get("vehicleNo", ""), case.get("job_no", ""))

    y = section_title(c, y, section_title_text, 12)

    col_sno = (MARGIN_L, 26)
    col_qty = (col_sno[0] + col_sno[1], 30)
    col_desc = (col_qty[0] + col_qty[1], 175)
    col_cond = (col_desc[0] + col_desc[1], 85)
    col_est = (col_cond[0] + col_cond[1], 105)
    col_ass = (col_est[0] + col_est[1], PAGE_W - MARGIN_R - (col_est[0] + col_est[1]))

    headers = [
        ("S/NO", col_sno[0], "center", col_sno[1]),
        ("QTY", col_qty[0], "center", col_qty[1]),
        ("DESCRIPTION", col_desc[0], "left", col_desc[1]),
        ("CONDITION /\nREMARKS", col_cond[0], "center", col_cond[1]),
        ("REPAIRER'S\nESTIMATES S$", col_est[0], "center", col_est[1]),
        ("OUR\nASSESSMENT S$", col_ass[0], "center", col_ass[1]),
    ]
    c.setFont(FONT_BOLD, 10)
    header_line_h = 10
    lines_per_header = [text.split("\n") for text, x, align, width in headers]
    max_lines = max(len(ls) for ls in lines_per_header)
    block_h = max_lines * header_line_h
    hy = y
    for (text, x, align, width), lines in zip(headers, lines_per_header):
        # 垂直居中：行数少的表头往下移，跟行数最多的表头对齐到同一个竖直中心
        top_offset = (max_lines - len(lines)) * header_line_h / 2
        yy = hy - top_offset
        for ln in lines:
            if align == "center":
                c.drawCentredString(x + width / 2, yy, ln)
            elif align == "right":
                c.drawRightString(x + width, yy, ln)
            else:
                c.drawString(x, yy, ln)
            yy -= header_line_h
    y = hy - block_h - 4
    c.line(MARGIN_L, y + 8, PAGE_W - MARGIN_R, y + 8)

    bottom_limit = MARGIN_BOTTOM + 45
    row_height = (y - bottom_limit) / PARTS_ROWS_PER_PAGE
    y -= 4
    c.setFont(FONT_REG, 10)
    ry = y
    for idx, row in enumerate(rows):
        c.setFont(FONT_REG, 10)
        c.drawCentredString(col_sno[0] + col_sno[1] / 2, ry, str(row["sn"]))
        c.drawCentredString(col_qty[0] + col_qty[1] / 2, ry, str(row.get("qty", "")))
        _draw_wrapped_or_ellipsis(c, (row.get("name") or row.get("desc", "")).title(), col_desc[0] + 3, ry,
                                   col_desc[1] - 6, FONT_REG, 10)
        _draw_wrapped_or_ellipsis(c, (row.get("condition", "")).title(), col_cond[0] + 16, ry,
                                   col_cond[1] - 6, FONT_REG, 10)
        c.drawRightString(col_est[0] + col_est[1] - 24, ry, fmt_money(row.get("est", 0)))
        c.drawRightString(col_ass[0] + col_ass[1] - 4, ry, fmt_money(row.get("ass", 0)))
        ry -= row_height

    if not is_last_page:
        c.setFont(FONT_REG, 8)
        c.drawRightString(PAGE_W - MARGIN_R, bottom_limit - 4, "Continue on next page...")
    elif totals:
        ty = ry - 8  # 紧贴最后一项之下，不再固定钉在页面底部
        c.setFont(FONT_REG, 10)
        label_x = col_cond[0] + col_cond[1] - 10
        if show_discount:
            c.line(col_est[0], ty + 12, col_ass[0] + col_ass[1], ty + 12)
            c.drawRightString(label_x, ty, "Sub Total:")
            c.drawRightString(col_est[0] + col_est[1] - 24, ty, fmt_money(totals["sub_est"]))
            c.drawRightString(col_ass[0] + col_ass[1] - 4, ty, fmt_money(totals["sub_ass"]))
            ty -= 14
            c.drawRightString(label_x, ty, "Less : Discount 10%")
            c.drawRightString(col_est[0] + col_est[1] - 24, ty, "-" + fmt_money(totals["disc_est"]))
            c.drawRightString(col_ass[0] + col_ass[1] - 4, ty, "-" + fmt_money(totals["disc_ass"]))
            ty -= 16
        c.setFont(FONT_BOLD, 10)
        c.line(col_est[0], ty + 12, col_ass[0] + col_ass[1], ty + 12)
        c.drawRightString(label_x, ty, "Total :")
        c.drawRightString(col_est[0] + col_est[1] - 24, ty, fmt_money(totals["total_est"]))
        c.drawRightString(col_ass[0] + col_ass[1] - 4, ty, fmt_money(totals["total_ass"]))
        c.line(col_est[0], ty - 3, col_ass[0] + col_ass[1], ty - 3)
        c.line(col_est[0], ty - 5, col_ass[0] + col_ass[1], ty - 5)

    draw_footer_report(c, page_num, total_pages)
    c.showPage()


def _labour_page(c, case, rows, page_num, total_pages, total_est, total_ass):
    y = draw_header(c)
    y = draw_subheader(c, y, "ANNEX B", case.get("vehicleNo", ""), case.get("job_no", ""))
    y = section_title(c, y, "PANEL & MECHANICAL LABOUR", 12)

    col_sno = (MARGIN_L, 33)
    col_est = (PAGE_W - MARGIN_R - 175, 85)
    col_ass = (col_est[0] + col_est[1], PAGE_W - MARGIN_R - (col_est[0] + col_est[1]))
    col_desc = (col_sno[0] + col_sno[1], col_est[0] - (col_sno[0] + col_sno[1]) - 10)

    c.setFont(FONT_BOLD, 10)
    header_line_h = 10
    headers = [
        ("S/NO", col_sno[0], col_sno[1], 10, "left"),
        ("DESCRIPTION", col_desc[0], col_desc[1], 10, "left"),
        ("REPAIRER'S\nESTIMATES S$", col_est[0], col_est[1], 10, "center"),
        ("OUR\nASSESSMENT S$", col_ass[0], col_ass[1], 10, "center"),
    ]
    lines_per_header = [text.split("\n") for text, x, width, size, align in headers]
    max_lines = max(len(ls) for ls in lines_per_header)
    block_h = max_lines * header_line_h
    hy = y
    for (text, x, width, size, align), lines in zip(headers, lines_per_header):
        c.setFont(FONT_BOLD, size)
        top_offset = (max_lines - len(lines)) * header_line_h / 2
        yy = hy - top_offset
        for ln in lines:
            if align == "center":
                c.drawCentredString(x + width / 2, yy, ln)
            else:
                c.drawString(x, yy, ln)
            yy -= header_line_h
    y = hy - block_h - -5
    c.line(MARGIN_L, y, PAGE_W - MARGIN_R, y)
    y -= 16

    for idx, row in enumerate(rows):
        c.setFont(FONT_REG, 10)
        desc_lines = _wrap_text(c, row["desc"], FONT_REG, 10, col_desc[1] - 8)
        n_lines = max(1, len(desc_lines))
        c.drawString(col_sno[0] + 8, y, str(idx + 1))  # S/NO
        ly = y
        for ln in desc_lines:
            c.drawString(col_desc[0], ly, ln)  # DESCRIPTION（靠左）
            ly -= 13
        c.drawRightString(col_est[0] + col_est[1] - 15, y, fmt_money(row["est"]))  # 估价（靠右）
        c.drawRightString(col_ass[0] + col_ass[1] - 15, y, fmt_money(row["ass"]))  # 核定价（靠右）
        y -= max(16, n_lines * 13 + 3)

    y -= 6
    c.setFont(FONT_BOLD, 10)
    c.line(col_est[0], y + 12, col_ass[0] + col_ass[1], y + 12)
    c.drawRightString(col_desc[0] + col_desc[1] - 4, y, "Total :")
    c.drawRightString(col_est[0] + col_est[1] - 15, y, fmt_money(total_est))
    c.drawRightString(col_ass[0] + col_ass[1] - 15, y, fmt_money(total_ass))
    c.line(col_est[0], y - 5, col_ass[0] + col_ass[1], y - 5)
    c.line(col_est[0], y - 7, col_ass[0] + col_ass[1], y - 7)

    draw_footer_report(c, page_num, total_pages)
    c.showPage()


def _wrap_text(c, text, font, size, max_width):
    words = str(text).split(" ")
    lines = []
    cur = ""
    for w in words:
        trial = (cur + " " + w).strip()
        if c.stringWidth(trial, font, size) <= max_width:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines or [""]


def _paint_final_page(c, case, parts_totals, misc_totals, labour_totals, page_num, total_pages):
    y = draw_header(c)
    y = draw_subheader(c, y, "ANNEX B", case.get("vehicleNo", ""), case.get("job_no", ""))
    y = section_title(c, y, "PAINT", 12)

    paint_code = (case.get("step1Data", {}) or {}).get("paintCode", "1")
    paint_code_label = PAINT_CODE_LABELS.get(str(paint_code), PAINT_CODE_LABELS["1"])
    c.setFont(FONT_REG, 10)
    c.drawString(MARGIN_L, y, "Code   :   " + paint_code_label)
    y -= 24

    col_est = (PAGE_W - MARGIN_R - 175, 85)
    col_ass = (col_est[0] + col_est[1], PAGE_W - MARGIN_R - (col_est[0] + col_est[1]))
    col_desc_x = MARGIN_L
    col_desc_w = col_est[0] - MARGIN_L - 10

    def draw_cost_header(y):
        """DESCRIPTION 靠左但垂直居中；REPAIRER'S/OUR 水平+垂直都居中。返回表头下方分割线之后的 y。"""
        header_line_h = 10
        headers = [
            ("DESCRIPTION", col_desc_x, col_desc_w, 10, "left"),
            ("REPAIRER'S\nESTIMATES S$", col_est[0], col_est[1], 9, "center"),
            ("OUR\nASSESSMENT S$", col_ass[0], col_ass[1], 9, "center"),
        ]
        lines_per_header = [t.split("\n") for t, x, w, size, align in headers]
        max_lines = max(len(ls) for ls in lines_per_header)
        block_h = max_lines * header_line_h
        hy = y
        for (t, x, w, size, align), lines in zip(headers, lines_per_header):
            c.setFont(FONT_BOLD, size)
            top_offset = (max_lines - len(lines)) * header_line_h / 2
            yy = hy - top_offset
            for ln in lines:
                if align == "center":
                    c.drawCentredString(x + w / 2, yy, ln)
                else:
                    c.drawString(x, yy, ln)
                yy -= header_line_h
        y = hy - block_h - 4
        c.line(MARGIN_L, y + 10, PAGE_W - MARGIN_R, y + 10)
        return y + 10

    def cost_block(title, desc_text, est_val, ass_val, y):
        c.setFont(FONT_BOLD, 10)
        c.drawString(MARGIN_L, y, title)
        title_w = c.stringWidth(title, FONT_BOLD, 10)
        c.line(MARGIN_L, y - 2, MARGIN_L + title_w, y - 2)
        y -= 16
        y = draw_cost_header(y)
        y -= 14
        lines = _wrap_text(c, desc_text, FONT_REG, 10, col_desc_w - 6)
        c.setFont(FONT_REG, 10)
        top_line_y = y
        for ln in lines:
            c.drawString(col_desc_x, y, ln)
            y -= 13
        c.drawRightString(col_est[0] + col_est[1] - 17, top_line_y, fmt_money(est_val))
        c.drawRightString(col_ass[0] + col_ass[1] - 17, top_line_y, fmt_money(ass_val))
        y -= 4
        c.setFont(FONT_BOLD, 10)
        c.line(col_est[0], y + 12, col_ass[0] + col_ass[1], y + 12)
        c.drawRightString(col_est[0] - 6, y, "Total :")
        c.drawRightString(col_est[0] + col_est[1] - 17, y, fmt_money(est_val))
        c.drawRightString(col_ass[0] + col_ass[1] - 17, y, fmt_money(ass_val))
        c.line(col_est[0], y - 3, col_ass[0] + col_ass[1], y - 3)
        c.line(col_est[0], y - 5, col_ass[0] + col_ass[1], y - 5)
        return y - 24

    paint_labour_est = _num(case.get("paintLabourEst"))
    paint_labour_ass = _num(case.get("paintLabourAss"))
    paint_mat_est = _num(case.get("paintMatEst"))
    paint_mat_ass = _num(case.get("paintMatAss"))

    y = cost_block(
        "LABOUR COST - PAINT",
        "To applies body filler , sanding , applies primer coat , spray "
        "painting and clearcoat. And polish the affected portions.",
        paint_labour_est, paint_labour_ass, y,
    )
    y -= 14
    y = cost_block(
        "MATERIAL COST - PAINT",
        "To supplies polyurethane paint, clear coat lacquer, primer, masking tape, cleaning solvent, "
        "1200- and 2000-grit sandpaper, filler, and paint thinner.",
        paint_mat_est, paint_mat_ass, y,
    )

    y -= 8
    c.setLineWidth(1.8)
    c.line(MARGIN_L, y, PAGE_W - MARGIN_R, y)
    c.setLineWidth(1)
    y -= 22

    y = section_title(c, y, "FINAL CALCULATION", 12)
    y = draw_cost_header(y)
    y -= 16

    final_rows = [
        ("Parts", parts_totals["total_est"], parts_totals["total_ass"]),
        ("Miscellaneous items", misc_totals["total_est"], misc_totals["total_ass"]),
        ("Panel & Mechanical Labour", labour_totals["total_est"], labour_totals["total_ass"]),
        ("Paint Labour", paint_labour_est, paint_labour_ass),
        ("Paint Material", paint_mat_est, paint_mat_ass),
    ]
    c.setFont(FONT_REG, 10)
    total_est_sum = 0.0
    total_ass_sum = 0.0
    for label, est_v, ass_v in final_rows:
        c.drawString(col_desc_x, y, label)
        c.drawRightString(col_est[0] + col_est[1] - 17, y, fmt_money(est_v))
        c.drawRightString(col_ass[0] + col_ass[1] - 17, y, fmt_money(ass_v))
        total_est_sum += est_v
        total_ass_sum += ass_v
        y -= 16

    c.setFont(FONT_BOLD, 10)
    c.drawRightString(col_est[0] - 6, y, "Total Repair Cost :")
    c.line(col_est[0], y + 12, col_ass[0] + col_ass[1], y + 12)
    c.drawRightString(col_est[0] + col_est[1] - 17, y, fmt_money(total_est_sum))
    c.drawRightString(col_ass[0] + col_ass[1] - 17, y, fmt_money(total_ass_sum))
    c.line(col_est[0], y - 3, col_ass[0] + col_ass[1], y - 3)
    c.line(col_est[0], y - 5, col_ass[0] + col_ass[1], y - 5)
    y -= 30

    y = section_title(c, y, "Remarks", 10)
    c.setFont(FONT_REG, 10)
    repair_days = case.get("repairDays", "") or "-"
    line1a = "The repairer has agreed to undertake the repair of  "
    line1b = f"S$ {fmt_money(total_ass_sum)} ."
    c.drawString(MARGIN_L, y, line1a)
    x2 = MARGIN_L + c.stringWidth(line1a, FONT_REG, 10)
    c.setFont(FONT_BOLD, 10)
    c.drawString(x2, y, line1b)
    y -= 15
    c.setFont(FONT_REG, 10)
    part2a = "And with a repair period of  "
    part2b = f"{repair_days}"
    part2c = "  working days ."
    c.drawString(MARGIN_L, y, part2a)
    x3 = MARGIN_L + c.stringWidth(part2a, FONT_REG, 10)
    c.setFont(FONT_BOLD, 10)
    c.drawString(x3, y, part2b)
    x4 = x3 + c.stringWidth(part2b, FONT_BOLD, 10)
    c.setFont(FONT_REG, 10)
    c.drawString(x4, y, part2c)

    y -= 50
    c.setFont(FONT_REG, 10)
    c.drawString(PAGE_W - MARGIN_R - 160, y, "Surveyed by:")
    y -= 55
    c.line(PAGE_W - MARGIN_R - 160, y, PAGE_W - MARGIN_R, y)
    y -= 12
    c.drawCentredString(PAGE_W - MARGIN_R - 80, y, "CHIN ZOEN LI")

    draw_footer_report(c, page_num, total_pages)
    c.showPage()


def generate_survey_report_pdf(case: dict) -> bytes:
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    global_pct = _num(case.get("globalPercentage", 100))
    parts_raw = case.get("parts", []) or []
    parts_computed = []
    for p in parts_raw:
        p2 = dict(p)
        p2["_global_pct"] = global_pct
        parts_computed.append(compute_part_row(p2))
    for i, row in enumerate(parts_computed):
        row["sn"] = i + 1

    misc_raw = case.get("misc", []) or []
    misc_computed = [compute_misc_row(m) for m in misc_raw]
    for i, row in enumerate(misc_computed):
        row["sn"] = i + 1

    labour_checks = case.get("labourChecks", []) or []
    custom_labour = case.get("customLabour", []) or []
    labour_rows = []
    for lc in labour_checks:
        idx = lc.get("index")
        desc = LABOUR_ITEMS[idx] if isinstance(idx, int) and 0 <= idx < len(LABOUR_ITEMS) else ""
        labour_rows.append({"desc": desc, "est": _num(lc.get("est")), "ass": _num(lc.get("ass"))})
    for cl in custom_labour:
        labour_rows.append({"desc": cl.get("desc", ""), "est": _num(cl.get("est")), "ass": _num(cl.get("ass"))})

    # ---- totals ----
    parts_sub_est = sum(r["est"] for r in parts_computed)
    parts_sub_ass = sum(r["ass"] for r in parts_computed)
    parts_disc_est = parts_sub_est * 0.10
    parts_disc_ass = parts_sub_ass * 0.10
    parts_totals = {
        "sub_est": parts_sub_est, "sub_ass": parts_sub_ass,
        "disc_est": parts_disc_est, "disc_ass": parts_disc_ass,
        "total_est": parts_sub_est - parts_disc_est,
        "total_ass": parts_sub_ass - parts_disc_ass,
    }
    misc_totals = {
        "total_est": sum(r["est"] for r in misc_computed),
        "total_ass": sum(r["ass"] for r in misc_computed),
    }
    labour_totals = {
        "total_est": sum(r["est"] for r in labour_rows),
        "total_ass": sum(r["ass"] for r in labour_rows),
    }

    # ---- pagination ----
    def chunk(lst, n):
        if not lst:
            return [[]]
        return [lst[i:i + n] for i in range(0, len(lst), n)]

    parts_chunks = chunk(parts_computed, PARTS_ROWS_PER_PAGE)
    misc_chunks = chunk(misc_computed, PARTS_ROWS_PER_PAGE)

    total_pages = 1 + len(parts_chunks) + len(misc_chunks) + 1 + 1
    page_num = 1

    _page1(c, case, page_num, total_pages)
    page_num += 1

    for i, ch in enumerate(parts_chunks):
        is_last = (i == len(parts_chunks) - 1)
        _items_page(c, case, "PARTS", "ANNEX A", ch, page_num, total_pages, is_last,
                    totals=parts_totals if is_last else None, show_discount=True)
        page_num += 1

    for i, ch in enumerate(misc_chunks):
        is_last = (i == len(misc_chunks) - 1)
        _items_page(c, case, "MISCELLANEOUS ITEMS", "ANNEX A", ch, page_num, total_pages, is_last,
                    totals=misc_totals if is_last else None, show_discount=False)
        page_num += 1

    _labour_page(c, case, labour_rows, page_num, total_pages,
                 labour_totals["total_est"], labour_totals["total_ass"])
    page_num += 1

    _paint_final_page(c, case, parts_totals, misc_totals, labour_totals, page_num, total_pages)

    c.save()
    buf.seek(0)
    return buf.read()


# --- API：生成 Job Number ---
@app.get("/api/generate_job_no")
async def generate_job_no(inspection_date: str):
    try:
        date_obj = None
        for fmt in ("%Y-%m-%d", "%d %b %Y"):
            try:
                date_obj = datetime.strptime(inspection_date.strip(), fmt)
                break
            except ValueError:
                pass

        if not date_obj:
            return {"job_no": "ERROR: Invalid Date"}

        year = date_obj.strftime("%Y")
        month = date_obj.strftime("%m")
        base_calc_val = date_obj.day * 12

        db = load_json_db()
        existing_jobs = list(db.keys())

        while True:
            day_calc_str = str(base_calc_val).zfill(3)
            job_no = f"TC/6/{year}/{month}{day_calc_str}"
            if job_no in existing_jobs:
                base_calc_val += 1
            else:
                break

        return {"job_no": job_no}
    except Exception:
        return {"job_no": "ERROR: Invalid Date"}


# --- API：获取所有 Cases 列表 ---
@app.get("/api/get_cases")
async def get_cases():
    db = load_json_db()
    cases_list = []
    for job_no, data in db.items():
        cases_list.append({
            "Job Number": job_no,
            "Batch No.": data.get("batchNo", ""),
            "Vehicle No.": data.get("vehicleNo", ""),
            "Make & Model": data.get("makeModel", ""),
            "Date Of Accident": data.get("accidentDate", ""),
            "Total Repair Cost": data.get("totalAss", 0)
        })
    return {"cases": cases_list}


# --- API：获取单个 Case 完整数据 ---
@app.get("/api/get_case_detail")
async def get_case_detail(job_no: str):
    db = load_json_db()
    if job_no not in db:
        raise HTTPException(status_code=404, detail="Case not found")
    return {"status": "success", "data": db[job_no]}


# --- API：删除指定 Case ---
@app.delete("/api/delete_case")
async def delete_case(job_no: str):
    db = load_json_db()
    if job_no not in db:
        return {"status": "error", "message": "Job Number not found"}

    del db[job_no]
    save_json_db(db)

    # 同步从 Excel 副本里删掉对应那一行，避免JSON跟Excel对不上
    if os.path.exists(CASES_DB_PATH):
        try:
            cases_df = pd.read_excel(CASES_DB_PATH)
            cases_df = cases_df[cases_df['Job Number'].astype(str) != job_no]
            cases_df.to_excel(CASES_DB_PATH, index=False)
        except Exception:
            pass  # Excel只是方便另作用途的副本，就算这里失败也不影响JSON主数据已删除

    return {"status": "success", "message": f"Case {job_no} DELETED！"}


# --- API：保存完整 Case 数据 ---
@app.post("/api/save_case")
async def save_case(payload: dict):
    try:
        job_no = payload.get("job_no")
        if not job_no:
            return {"status": "error", "message": "Job Number cannot be empty"}

        # 1. 全量保存所有表单数据到 JSON
        db = load_json_db()
        db[job_no] = payload
        save_json_db(db)

        # 2. 同步导出简易 Excel 供您另作用途
        record = {
            "Batch No.": payload.get("batchNo", ""),
            "Vehicle No.": payload.get("vehicleNo", ""),
            "Job Number": job_no,
            "Make & Model": payload.get("makeModel", ""),
            "Date Of Accident": payload.get("accidentDate", ""),
            "Date Of Inspection": payload.get("inspectionDate", ""),
            "Report Finish Date": payload.get("finishDate", ""),
            "Total Repair Cost": payload.get("totalAss", 0),
            "Repair Days": payload.get("repairDays", ""),
            "Rental Per Day": payload.get("rentalPerDay", ""),
            "Total Rental": payload.get("totalRental", ""),
            "Survey Fees": payload.get("surveyFees", ""),
            "Transportation": payload.get("transportation", ""),
            "Photographs Cost": payload.get("photographsCost", ""),
            "Total Survey Fee": payload.get("totalSurveyFee", "")
        }
        new_row_df = pd.DataFrame([record])
        if os.path.exists(CASES_DB_PATH):
            try:
                cases_df = pd.read_excel(CASES_DB_PATH)
                cases_df = cases_df[cases_df['Job Number'].astype(str) != job_no]
                cases_df = pd.concat([cases_df, new_row_df], ignore_index=True)
                cases_df.to_excel(CASES_DB_PATH, index=False)
            except Exception:
                new_row_df.to_excel(CASES_DB_PATH, index=False)
        else:
            new_row_df.to_excel(CASES_DB_PATH, index=False)

        return {"status": "success", "message": f"Case {job_no} saved！"}
    except Exception as e:
        return {"status": "error", "message": str(e)}


# --- API：生成 Invoice PDF ---
@app.get("/api/generate_invoice_pdf")
async def api_generate_invoice_pdf(job_no: str):
    db = load_json_db()
    if job_no not in db:
        raise HTTPException(status_code=404, detail="Case not found")
    try:
        pdf_bytes = generate_invoice_pdf(db[job_no])
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"生成 Invoice PDF 失败: {e}")
    filename = f"Invoice_{job_no.replace('/', '_')}.pdf"
    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={"Content-Disposition": f'inline; filename="{filename}"'},
    )


# --- API：生成 Survey Report PDF ---
@app.get("/api/generate_survey_report_pdf")
async def api_generate_survey_report_pdf(job_no: str):
    db = load_json_db()
    if job_no not in db:
        raise HTTPException(status_code=404, detail="Case not found")
    try:
        pdf_bytes = generate_survey_report_pdf(db[job_no])
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"生成 Survey Report PDF 失败: {e}")
    filename = f"SurveyReport_{job_no.replace('/', '_')}.pdf"
    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={"Content-Disposition": f'inline; filename="{filename}"'},
    )


# --- 其他查询 API ---
@app.get("/api/get_model_info")
async def get_model_info(model_code: str):
    if not os.path.exists(EXCEL_PATH):
        return {"description": "Excel not found", "default_percentage": 100}
    try:
        df = pd.read_excel(EXCEL_PATH)
        df['Model Code Clean'] = df['Model Code'].astype(str).str.strip().str.upper()
        matched = df[df['Model Code Clean'] == model_code.strip().upper()]
        if matched.empty:
            return {"description": "Model Not Found", "default_percentage": 100}
        desc = str(matched.iloc[0].get('Model Description', ''))
        pct = float(matched.iloc[0].get('Default Percentage', 150))
        return {"description": desc, "default_percentage": pct}
    except Exception as e:
        return {"description": str(e), "default_percentage": 100}


@app.get("/api/get_part_price")
async def get_part_price(model_code: str, part_name: str):
    if not os.path.exists(EXCEL_PATH):
        return {"price": 0.0}
    try:
        df = pd.read_excel(EXCEL_PATH)
        df['Model Code Clean'] = df['Model Code'].astype(str).str.strip().str.upper()
        df['Parts Name Clean'] = df['Parts Name'].astype(str).str.strip().str.upper()
        matched = df[(df['Model Code Clean'] == model_code.strip().upper()) & (
                df['Parts Name Clean'] == part_name.strip().upper())]
        if not matched.empty:
            return {"price": float(matched.iloc[0].get('Price', 0.0))}
        return {"price": 0.0}
    except Exception:
        return {"price": 0.0}


@app.get("/", response_class=HTMLResponse)
async def get_form():
    html_content = """
    <!DOCTYPE html>
    <html lang="zh-CN">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Vehicle Assessment System</title>
        <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-slate-100 p-4 md:p-8 font-sans">

        <!-- Top Navigation Bar -->
        <div id="topNavBar" class="max-w-5xl mx-auto bg-slate-800 text-white rounded-t-xl p-3 flex justify-between items-center shadow-md sticky top-0 z-50">
            <div class="flex space-x-2">
                <button onclick="goToLandingPage()" class="bg-slate-600 hover:bg-slate-500 px-3 py-1 rounded text-xs font-semibold">🏠 Home Menu</button>
                <button onclick="openCaseListModal()" class="bg-purple-600 hover:bg-purple-500 px-3 py-1 rounded text-xs font-bold">📂 Open Cases</button>
                <button id="backBtn" onclick="goBackToStep1()" class="hidden bg-slate-600 hover:bg-slate-500 px-3 py-1 rounded text-xs">&larr; Back</button>
            </div>
            <div class="text-xs md:text-sm font-semibold" id="topBarInfo">Dashboard</div>
            <div>
                <button id="topEditBtn" onclick="enableEditStep1()" class="hidden bg-amber-500 hover:bg-amber-400 px-3 py-1 rounded text-xs font-bold mr-2">Edit Basic Info</button>
                <button id="topSaveBtn" onclick="saveBasicInfo()" class="hidden bg-blue-500 hover:bg-blue-400 px-3 py-1 rounded text-xs font-bold shadow">Save Basic Info</button>
                <button id="topSaveStep2Btn" onclick="lockAndSaveCase()" class="hidden bg-green-500 hover:bg-green-400 px-3 py-1 rounded text-xs font-bold shadow">🔒 Save & Lock Case</button>
                <button id="topEditStep2Btn" onclick="unlockStep2()" class="hidden bg-amber-500 hover:bg-amber-400 px-3 py-1 rounded text-xs font-bold">✏️ Edit Details</button>
            </div>
        </div>

        <!-- LANDING PAGE / HOME MENU -->
        <div id="landingPage" class="max-w-5xl mx-auto bg-white rounded-b-xl shadow-md p-8 mb-10 border-t-4 border-blue-600">
            <div class="text-center py-6 border-b mb-8">
                <h1 class="text-3xl font-extrabold text-slate-800 tracking-tight">Vehicle Assessment Management System</h1>
                <p class="text-gray-500 text-sm mt-2">Powered by RichYield</p>
            </div>

            <div class="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
                <!-- Action 1: Create New Case -->
                <div onclick="createNewCase()" class="group cursor-pointer bg-slate-50 hover:bg-blue-50 border-2 border-slate-200 hover:border-blue-500 rounded-xl p-6 text-center transition duration-200 shadow-sm hover:shadow-md flex flex-col justify-between items-center">
                    <div class="p-4 bg-blue-100 group-hover:bg-blue-500 text-blue-600 group-hover:text-white rounded-full mb-4 transition">
                        <svg class="w-10 h-10" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4v16m8-8H4"></path></svg>
                    </div>
                    <div>
                        <h2 class="text-xl font-bold text-slate-800 group-hover:text-blue-700">➕ Create New Case</h2>
                    </div>
                    <span class="mt-6 inline-block bg-blue-600 text-white px-4 py-2 rounded-lg text-xs font-bold group-hover:bg-blue-700">Create &rarr;</span>
                </div>

                <!-- Action 2: Open Saved Cases -->
                <div onclick="openCaseListModal()" class="group cursor-pointer bg-slate-50 hover:bg-purple-50 border-2 border-slate-200 hover:border-purple-500 rounded-xl p-6 text-center transition duration-200 shadow-sm hover:shadow-md flex flex-col justify-between items-center">
                    <div class="p-4 bg-purple-100 group-hover:bg-purple-500 text-purple-600 group-hover:text-white rounded-full mb-4 transition">
                        <svg class="w-10 h-10" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 8h14M5 8a2 2 0 01-2-2V5a2 2 0 012-2h14a2 2 0 012 2v1a2 2 0 01-2 2M5 8v10a2 2 0 002 2h10a2 2 0 002-2V8m-9 4h4"></path></svg>
                    </div>
                    <div>
                        <h2 class="text-xl font-bold text-slate-800 group-hover:text-purple-700">📂 Search Saved Cases</h2>
                    </div>
                    <span class="mt-6 inline-block bg-purple-600 text-white px-4 py-2 rounded-lg text-xs font-bold group-hover:bg-purple-700">Search &rarr;</span>
                </div>
            </div>
        </div>

        <!-- Step 1 Page -->
        <div id="step1Page" class="hidden max-w-5xl mx-auto bg-white rounded-b-xl shadow-md p-6 md:p-8 mb-10">
            <h1 class="text-2xl font-bold text-slate-800 mb-6 border-b pb-3">🚗 Step 1: Basic Vehicle Info</h1>
            <form id="basicForm" class="space-y-6">
                <div class="bg-blue-50 p-4 rounded border border-blue-100 mb-6 flex flex-col md:flex-row gap-4">
                    <div class="flex-1">
                        <label class="block text-xs font-bold text-blue-800 uppercase mb-1">Batch Number</label>
                        <input type="text" id="batchNo" onblur="formatBatchNo(this)" class="w-full border rounded p-2 text-sm outline-none step1-input" placeholder="Enter Batch No...(e.g. 25)">
                    </div>
                    <div class="flex-1">
                        <label class="block text-xs font-bold text-blue-800 uppercase mb-1">Job Number</label>
                        <input type="text" id="jobNo" readonly class="w-full border border-blue-300 bg-blue-100 rounded p-2 text-sm font-bold text-blue-900 outline-none" placeholder="Click SAVE to generate">
                    </div>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">1. Owner Name</label>
                        <input type="text" id="ownerNameInput" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div class="md:row-span-2">
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">2. Owner Address</label>
                        <div class="space-y-2">
                            <input type="text" id="ownerAddr1" placeholder="Line 1" class="w-full border rounded p-2 text-sm step1-input">
                            <input type="text" id="ownerAddr2" placeholder="Line 2" class="w-full border rounded p-2 text-sm step1-input">
                            <input type="text" id="ownerAddr3" placeholder="Line 3" class="w-full border rounded p-2 text-sm step1-input">
                        </div>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">3. Vehicle No.</label>
                        <input type="text" id="vehicleNo" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">4. Make & Model</label>
                        <input type="text" id="makeModelInput" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">5. Year Make</label>
                        <input type="number" id="yearMakeInput" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">6. Chassis No.</label>
                        <input type="text" id="chassisNoInput" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">7. Engine No.</label>
                        <input type="text" id="engineNoInput" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">8. Engine Capacity (cc)</label>
                        <input type="text" id="engineCapInput" oninput="formatNumberInput(this)" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">9. Power Rating (kW)</label>
                        <input type="number" id="powerRatingInput" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">10. Motor No.</label>
                        <input type="text" id="motorNoInput" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">11. Colour</label>
                        <input type="text" id="colourInput" class="w-full border rounded p-2 text-sm step1-input">
                        <select id="paintCodeSelect" class="w-full border rounded p-2 text-sm step1-input mt-2 bg-gray-50">
                        <option value="1">1 : Two-Coat Metallic Paint</option>
                        <option value="2">2 : Three-Stage Pearlescent Paint</option>
                        <option value="3">3 : Single-Stage SOLID Paint (Truck / Commercial)</option>
                        <option value="4">4 : Two-Coat SOLID Paint (Passenger Car)</option>
                        </select>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">12. Mileage (km)</label>
                        <input type="number" id="mileageInput" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                </div>

                <div class="border-t pt-4 bg-slate-50 p-4 rounded-lg border">
                    <label class="block text-sm font-bold text-slate-700 mb-2">13. TYRES CONDITION</label>
                    <div class="mb-3">
                        <select id="wheelCount" onchange="renderTyreTable()" class="border rounded p-2 text-sm bg-white outline-none step1-input">
                            <option value="">-- Please select the number of tires --</option>
                            <option value="2">2 Wheel (Motorcycle)</option>
                            <option value="4">4 Wheel (Car / Van)</option>
                            <option value="6">6 Wheel (Truck)</option>
                            <option value="8">8 Wheel (Heavy Truck)</option>
                            <option value="10">10 Wheel (Trailer)</option>
                        </select>
                    </div>
                    <div id="tyreTableContainer">
                        <p class="text-xs text-gray-400 italic">The table will automatically expand after you select the number of tires</p>
                    </div>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-2 gap-4 border-t pt-4">
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">14. Damaged Position</label>
                        <textarea id="damagedPosInput" rows="2" class="w-full border rounded p-2 text-sm step1-input"></textarea>
                    </div>
                    <div class="md:row-span-3">
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">18. Workshop Name & Address</label>
                        <select id="wsPreset" onchange="applyWorkshopPreset()" class="w-full border rounded p-2 text-sm mb-2 bg-gray-50">
                            <option value="">-- 手动输入 / Manual Input --</option>
                            <option value="ming_hua">Ming Hua Auto Services</option>
                            <option value="kang_auto">KANG AUTO ENGINEERING PTE LTD</option>
                        </select>
                        <div class="space-y-2">
                            <input type="text" id="wsName" placeholder="Workshop Name" class="w-full border rounded p-2 text-sm step1-input">
                            <input type="text" id="wsAddr1" placeholder="Address Line 1" class="w-full border rounded p-2 text-sm step1-input">
                            <input type="text" id="wsAddr2" placeholder="Address Line 2" class="w-full border rounded p-2 text-sm step1-input">
                            <input type="text" id="wsAddr3" placeholder="Address Line 3" class="w-full border rounded p-2 text-sm step1-input">
                        </div>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">15. Date of Accident</label>
                        <input type="text" id="accidentDate" placeholder="e.g. 12 Oct 2026" onfocus="(this.type='date')" onblur="handleDateBlur(this)" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">16. Date of Inspection</label>
                        <input type="text" id="inspectionDate" placeholder="e.g. 12 Oct 2026" onfocus="(this.type='date')" onchange="autoCalculateFinishDate()" onblur="handleDateBlur(this)" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">17. Report Finish Date</label>
                        <input type="text" id="finishDate" readonly placeholder="Auto-calculated" class="w-full border rounded p-2 text-sm bg-gray-100 text-gray-600">
                    </div>
                </div>

                <div class="pt-6 flex justify-between items-center border-t">
                    <button type="button" id="mainSaveBtn" onclick="saveBasicInfo()" class="bg-blue-600 hover:bg-blue-700 text-white font-bold py-3 px-8 rounded-lg shadow-md">
                        💾 SAVE & GENERATE JOB NO.
                    </button>
                    <button type="button" id="nextBtn" onclick="goToStep2()" class="hidden bg-green-600 hover:bg-green-700 text-white font-bold py-3 px-8 rounded-lg shadow-md">
                        NEXT: Damage Details &rarr;
                    </button>
                </div>
            </form>
        </div>

        <!-- Step 2 Page -->
        <div id="step2Page" class="hidden max-w-5xl mx-auto bg-white rounded-b-xl shadow-md p-6 md:p-8 mb-10">
            <h1 class="text-2xl font-bold text-slate-800 mb-6 border-b pb-3">🛠️ Step 2: Assessment Details & Cost</h1>

            <!-- PARTS -->
            <div class="mb-10 bg-slate-50 p-4 rounded-xl border border-slate-200">
                <div class="flex flex-col mb-4 gap-3">
                    <h2 class="text-lg font-bold text-slate-800">PARTS ASSESSMENT</h2>
                    <div class="bg-white p-3 rounded-lg border shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
                        <div class="flex items-center space-x-3">
                            <div class="flex items-center space-x-1">
                                <label class="text-xs font-bold text-slate-600">MODEL CODE:</label>
                                <input type="text" id="modelCodeInput" placeholder="e.g. A5" onblur="fetchModelInfo()" class="border rounded px-2 py-1 text-xs uppercase focus:ring-2 focus:ring-blue-500 outline-none step2-input w-28 font-bold">
                            </div>
                            <div class="flex items-center space-x-1">
                                <label class="text-xs font-bold text-slate-600">[%]:</label>
                                <input type="number" id="globalPercentage" value="150" step="1" onchange="recalcAllPartsRows()" onkeyup="recalcAllPartsRows()" class="border rounded px-2 py-1 text-xs w-20 text-center font-bold text-blue-600 step2-input">
                                <span class="text-xs font-bold text-slate-600">%</span>
                            </div>
                        </div>
                        <div class="text-xs text-slate-600 font-medium bg-slate-50 p-2 rounded border flex-1 md:text-right" id="modelDescriptionText">
                            Description: <span class="italic text-gray-400">Please enter model code</span>
                        </div>
                    </div>
                </div>

                <div class="overflow-x-auto">
                    <table class="w-full text-left bg-white rounded border text-xs">
                        <thead>
                            <tr class="bg-slate-200 text-slate-700 font-bold uppercase">
                                <th class="p-2 border w-10 text-center">S/N</th>
                                <th class="p-2 border w-14">QTY</th>
                                <th class="p-2 border">PARTS NAME</th>
                                <th class="p-2 border w-24">CODE</th>
                                <th class="p-2 border w-28">CONDITION</th>
                                <th class="p-2 border w-28">NETT PRICE (S$)</th>
                                <th class="p-2 border w-28">ESTIMATE (S$)</th>
                                <th class="p-2 border w-28">ASSESSMENT (S$)</th>
                                <th class="p-2 border w-10 text-center">DEL</th>
                            </tr>
                        </thead>
                        <tbody id="partsTbody"></tbody>
                    </table>
                </div>
                <button onclick="addPartsRow()" class="step2-btn mt-3 bg-blue-600 hover:bg-blue-700 text-white font-bold px-3 py-1.5 rounded text-xs">
                    + Add Part Row
                </button>
                <div class="mt-4 border-t pt-3 text-xs flex justify-end">
                    <div class="w-80 space-y-1 bg-white p-3 rounded border shadow-sm">
                        <div class="flex justify-between"><span class="text-gray-600">Sub Total:</span><span class="font-bold">S$ <span id="partsSubEst">0.00</span> | S$ <span id="partsSubAss">0.00</span></span></div>
                        <div class="flex justify-between text-red-600"><span>Less: Discount 10%:</span><span class="font-bold">-S$ <span id="partsDiscEst">0.00</span> | -S$ <span id="partsDiscAss">0.00</span></span></div>
                        <div class="flex justify-between text-sm font-bold border-t pt-1 text-blue-900"><span>Total Parts:</span><span>S$ <span id="partsTotalEst">0.00</span> | S$ <span id="partsTotalAss">0.00</span></span></div>
                    </div>
                </div>
            </div>

            <!-- MISCELLANEOUS -->
            <div class="mb-10 bg-slate-50 p-4 rounded-xl border border-slate-200">
                <h2 class="text-lg font-bold text-slate-800 mb-3">MISCELLANEOUS ITEMS</h2>
                <div class="overflow-x-auto">
                    <table class="w-full text-left bg-white rounded border text-xs">
                        <thead>
                            <tr class="bg-slate-200 text-slate-700 font-bold uppercase">
                                <th class="p-2 border w-10 text-center">S/N</th>
                                <th class="p-2 border w-14">QTY</th>
                                <th class="p-2 border">DESCRIPTION</th>
                                <th class="p-2 border w-24">CODE</th>
                                <th class="p-2 border w-36">CONDITION</th>
                                <th class="p-2 border w-28">ESTIMATE (S$)</th>
                                <th class="p-2 border w-28">ASSESSMENT (S$)</th>
                                <th class="p-2 border w-10 text-center">DEL</th>
                            </tr>
                        </thead>
                        <tbody id="miscTbody"></tbody>
                    </table>
                </div>
                <div class="flex justify-between items-center mt-3">
                    <button onclick="addMiscRow()" class="step2-btn bg-blue-600 hover:bg-blue-700 text-white font-bold px-3 py-1.5 rounded text-xs">+ Add Misc Row</button>
                    <div class="text-xs bg-white p-2 rounded border font-bold text-slate-800">Misc Total: S$ <span id="miscTotalEst">0.00</span> | S$ <span id="miscTotalAss">0.00</span></div>
                </div>
            </div>

            <!-- LABOUR -->
            <div class="mb-10 bg-slate-50 p-4 rounded-xl border border-slate-200">
                <h2 class="text-lg font-bold text-slate-800 mb-3">PANEL & MECHANICAL LABOUR</h2>
                <div id="labourChecklist" class="space-y-2 mb-4"></div>
                <div class="border-t pt-3 mt-3">
                    <div class="text-xs font-bold text-slate-600 mb-2 uppercase">Custom Labour Items</div>
                    <div id="customLabourContainer" class="space-y-2"></div>
                    <button onclick="addCustomLabourRow()" class="step2-btn mt-2 bg-slate-700 hover:bg-slate-800 text-white font-bold px-3 py-1.5 rounded text-xs">+ Add Custom Labour Row</button>
                </div>
                <div class="flex justify-end mt-3 text-xs bg-white p-2 rounded border font-bold text-slate-800">Labour Total: S$ <span id="labourTotalEst">0.00</span> | S$ <span id="labourTotalAss">0.00</span></div>
            </div>

            <!-- PAINT -->
            <div class="mb-10 bg-slate-50 p-4 rounded-xl border border-slate-200">
                <h2 class="text-lg font-bold text-slate-800 mb-3">PAINT</h2>
                <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div>
                        <h3 class="font-bold text-xs text-slate-600 uppercase mb-2">Paint Labour Cost</h3>
                        <div class="flex gap-2">
                            <input type="text" value="Paint Labour" readonly class="flex-1 border rounded p-1.5 text-xs bg-gray-50">
                            <input type="text" placeholder="Est (S$)" id="paintLabourEst" oninput="formatNumberInput(this)" onchange="this.value = formatNumber(parseNumber(this.value)); calcGrandTotals();" onkeyup="calcGrandTotals()" class="w-28 border rounded p-1.5 text-xs font-mono step2-input">
                            <input type="text" placeholder="Ass (S$)" id="paintLabourAss" oninput="formatNumberInput(this)" onchange="this.value = formatNumber(parseNumber(this.value)); calcGrandTotals();" onkeyup="calcGrandTotals()" class="w-28 border rounded p-1.5 text-xs font-mono font-bold step2-input">
                        </div>
                    </div>
                    <div>
                        <h3 class="font-bold text-xs text-slate-600 uppercase mb-2">Paint Material Cost</h3>
                        <div class="flex gap-2">
                            <input type="text" value="Paint Material" readonly class="flex-1 border rounded p-1.5 text-xs bg-gray-50">
                            <input type="text" placeholder="Est (S$)" id="paintMatEst" oninput="formatNumberInput(this)" onblur="if(this.value.trim() !== '') { this.value = formatNumber(parseNumber(this.value)); } calcGrandTotals();" calcGrandTotals();" onkeyup="calcGrandTotals()" class="w-28 border rounded p-1.5 text-xs font-mono step2-input">
                            <input type="text" placeholder="Ass (S$)" id="paintMatAss" oninput="formatNumberInput(this)" onblur="if(this.value.trim() !== '') { this.value = formatNumber(parseNumber(this.value)); } calcGrandTotals();" class="w-28 border rounded p-1.5 text-xs font-mono font-bold step2-input">
                        </div>
                    </div>
                </div>
            </div>

            <!-- RENTAL & SURVEY DETAILS -->
            <div class="mb-10 bg-blue-50 p-4 rounded-xl border border-blue-100">
                <h2 class="text-base font-bold text-blue-900 mb-4 uppercase border-b border-blue-200 pb-2">Rental & Survey Fee Details</h2>
                <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div>
                        <label class="block text-xs font-bold text-blue-900 uppercase mb-1">19. Working Days (Repair Days)</label>
                        <input type="text" id="repairDaysInput" placeholder="e.g. 5" onkeyup="calcTotalRental()" onchange="calcTotalRental()" class="w-full border rounded p-2 text-sm step2-input">
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-blue-900 uppercase mb-1">20. Photo No. (Quantity)</label>
                        <input type="text" id="photoQtyInput" placeholder="e.g. 12" onkeyup="calcSurveyFees()" onchange="calcSurveyFees()" class="w-full border rounded p-2 text-sm step2-input">
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-blue-900 uppercase mb-1">21.1. Rental per day (S$)</label>
                        <input type="text" id="rentalPerDayInput" placeholder="0.00" oninput="formatNumberInput(this)" onkeyup="calcTotalRental()" onchange="calcTotalRental()" class="w-full border rounded p-2 text-sm font-mono step2-input">
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-blue-900 uppercase mb-1">21.2. Total Rental Cost (Calculated)</label>
                        <input type="text" id="totalRentalInput" readonly placeholder="S$ 0.00" class="w-full border border-blue-200 rounded p-2 text-sm font-mono font-bold bg-blue-100 text-blue-900 outline-none">
                    </div>

                    <div class="border-t border-blue-200 col-span-1 md:col-span-2 my-1"></div>

                    <div>
                        <label class="block text-xs font-bold text-blue-900 uppercase mb-1">22. Survey Fees S$ (Auto-calculated)</label>
                        <input type="text" id="surveyFeesInput" readonly placeholder="S$ 0.00" class="w-full border border-blue-200 rounded p-2 text-sm font-mono font-bold bg-blue-100 text-blue-900 outline-none">
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-blue-900 uppercase mb-1">23. Transportation S$</label>
                        <input type="text" id="transportationInput" value="200.00" oninput="formatNumberInput(this)" onkeyup="calcSurveyFees()" onchange="calcSurveyFees()" class="w-full border rounded p-2 text-sm font-mono step2-input">
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-blue-900 uppercase mb-1">24. Photographs ($2) per copies</label>
                        <input type="text" id="photographsCostInput" readonly placeholder="S$ 0.00" class="w-full border border-blue-200 rounded p-2 text-sm font-mono font-bold bg-blue-100 text-blue-900 outline-none">
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-blue-900 uppercase mb-1">25. Total Survey Fee S$</label>
                        <input type="text" id="totalSurveyFeeInput" readonly placeholder="S$ 0.00" class="w-full border border-blue-300 rounded p-2 text-sm font-mono font-bold bg-blue-200 text-blue-950 outline-none">
                    </div>
                </div>
            </div>

            <div class="flex justify-end gap-4 mb-8">
                <button id="bottomSaveStep2Btn" onclick="lockAndSaveCase()" class="bg-green-600 hover:bg-green-700 text-white font-bold py-3 px-8 rounded-lg shadow-md">🔒 LOCK & SAVE CASE TO DB</button>
                <button id="bottomEditStep2Btn" onclick="unlockStep2()" class="hidden bg-amber-500 hover:bg-amber-600 text-white font-bold py-3 px-8 rounded-lg shadow-md">✏️ UNLOCK TO EDIT DETAILS</button>
                <button id="bottomInvoiceBtn" onclick="downloadInvoicePDF()" class="hidden bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-3 px-6 rounded-lg shadow-md">📄 Invoice PDF</button>
                <button id="bottomReportBtn" onclick="downloadSurveyReportPDF()" class="hidden bg-teal-600 hover:bg-teal-700 text-white font-bold py-3 px-6 rounded-lg shadow-md">📄 Survey Report PDF</button>
            </div>

            <!-- FINAL CALCULATION -->
            <div class="bg-slate-800 text-white p-6 rounded-xl shadow-lg">
                <h2 class="text-lg font-bold text-yellow-400 mb-4 uppercase">Final Calculation</h2>
                <table class="w-full text-left text-sm">
                    <thead>
                        <tr class="border-b border-slate-700 text-slate-400 text-xs uppercase">
                            <th class="py-2">Description</th>
                            <th class="py-2 text-right">Repairer's Estimate (S$)</th>
                            <th class="py-2 text-right">Our Assessment (S$)</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-700">
                        <tr><td class="py-2">PARTS</td><td class="text-right" id="finalPartsEst">S$ 0.00</td><td class="text-right" id="finalPartsAss">S$ 0.00</td></tr>
                        <tr><td class="py-2">MISCELLANEOUS ITEMS</td><td class="text-right" id="finalMiscEst">S$ 0.00</td><td class="text-right" id="finalMiscAss">S$ 0.00</td></tr>
                        <tr><td class="py-2">PANEL & MECHANICAL LABOUR</td><td class="text-right" id="finalLabourEst">S$ 0.00</td><td class="text-right" id="finalLabourAss">S$ 0.00</td></tr>
                        <tr><td class="py-2">PAINT LABOUR</td><td class="text-right" id="finalPaintLabourEst">S$ 0.00</td><td class="text-right" id="finalPaintLabourAss">S$ 0.00</td></tr>
                        <tr><td class="py-2">PAINT MATERIAL</td><td class="text-right" id="finalPaintMatEst">S$ 0.00</td><td class="text-right" id="finalPaintMatAss">S$ 0.00</td></tr>
                        <tr class="font-bold text-base text-yellow-400 pt-2">
                            <td class="py-3">TOTAL REPAIR COST</td>
                            <td class="text-right py-3" id="finalTotalEst">S$ 0.00</td>
                            <td class="text-right py-3" id="finalTotalAss">S$ 0.00</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>

        <!-- Cases 列表弹窗 -->
        <div id="caseModal" class="hidden fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
            <div class="bg-white rounded-xl shadow-2xl w-full max-w-4xl max-h-[80vh] flex flex-col overflow-hidden">
                <div class="bg-slate-800 text-white p-4 flex justify-between items-center">
                    <h3 class="font-bold text-lg">📂 Saved Cases Database</h3>
                    <button onclick="closeCaseListModal()" class="text-gray-400 hover:text-white font-bold text-xl">✕</button>
                </div>
                <div class="p-4 bg-slate-50 border-b">
                    <input type="text" id="caseSearchInput" onkeyup="filterCasesTable()" placeholder="🔍 Search by Job No, Vehicle No, Batch No..." class="w-full border rounded-lg p-2 text-sm outline-none">
                </div>
                <div class="overflow-y-auto p-4 flex-1">
                    <table class="w-full text-left border text-xs">
                        <thead>
                            <tr class="bg-slate-200 text-slate-700 font-bold uppercase sticky top-0">
                                <th class="p-2 border">Job Number</th>
                                <th class="p-2 border">Batch No.</th>
                                <th class="p-2 border">Vehicle No.</th>
                                <th class="p-2 border">Make & Model</th>
                                <th class="p-2 border">Date Of Accident</th>
                                <th class="p-2 border text-right">Total Assessment</th>
                                <th class="p-2 border text-center">Actions</th>
                            </tr>
                        </thead>
                        <tbody id="caseListTbody"></tbody>
                    </table>
                </div>
            </div>
        </div>

        <script>
            // --- 页面导航控制器 ---
            function goToLandingPage() {
                document.getElementById('landingPage').classList.remove('hidden');
                document.getElementById('step1Page').classList.add('hidden');
                document.getElementById('step2Page').classList.add('hidden');
                document.getElementById('backBtn').classList.add('hidden');
                document.getElementById('topSaveBtn').classList.add('hidden');
                document.getElementById('topEditBtn').classList.add('hidden');
                document.getElementById('topSaveStep2Btn').classList.add('hidden');
                document.getElementById('topEditStep2Btn').classList.add('hidden');
                document.getElementById('topBarInfo').innerText = "Dashboard";
            }

            function createNewCase() {
                resetAllForms();
                document.getElementById('landingPage').classList.add('hidden');
                document.getElementById('step1Page').classList.remove('hidden');
                document.getElementById('step2Page').classList.add('hidden');
                enableEditStep1();
                document.getElementById('topBarInfo').innerText = "NEW CASE - Basic Info";
            }

            function resetAllForms() {
                document.getElementById('basicForm').reset();
                document.getElementById('jobNo').value = "";
                document.getElementById('tyreTableContainer').innerHTML = `<p class="text-xs text-gray-400 italic">The table will automatically expand after you select the number of tires</p>`;
                document.getElementById('partsTbody').innerHTML = "";
                document.getElementById('miscTbody').innerHTML = "";
                partsRowCounter = 0;
                miscRowCounter = 0;
                addPartsRow();
                addMiscRow();

                // 重置所有 Labour 选项
                labourItems.forEach((_, idx) => {
                    const cb = document.querySelector(`#labourRow_${idx} input[type="checkbox"]`);
                    if (cb && cb.checked) cb.click();
                });
                document.getElementById('customLabourContainer').innerHTML = "";
                
                // 强制重新计算 PARTS / MISC / LABOUR 的总计，避免残留上一个 Case 的数字
                calcPartsTotals();
                calcMiscTotals();
                calcLabourTotals();

                // 重置其他 Step 2 字段
                document.getElementById('modelCodeInput').value = "";
                document.getElementById('globalPercentage').value = "150";
                document.getElementById('modelDescriptionText').innerHTML = `Description: <span class="italic text-gray-400">Please enter model code</span>`;
                document.getElementById('repairDaysInput').value = "";
                document.getElementById('photoQtyInput').value = "";
                document.getElementById('rentalPerDayInput').value = "";
                document.getElementById('totalRentalInput').value = "S$ 0.00";
                document.getElementById('transportationInput').value = "200.00";
                document.getElementById('paintLabourEst').value = "";
                document.getElementById('paintLabourAss').value = "";
                document.getElementById('paintMatEst').value = "";
                document.getElementById('paintMatAss').value = "";
                document.getElementById('bottomInvoiceBtn').classList.add('hidden');
                document.getElementById('bottomReportBtn').classList.add('hidden');
                calcGrandTotals();
            }

            function formatDateToDDMMMYYYY(dateStr) {
                if (!dateStr) return "";
                const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
                if (dateStr.includes("-")) {
                    const parts = dateStr.split("-");
                    if (parts.length === 3) {
                        const year = parts[0];
                        const monthIdx = parseInt(parts[1], 10) - 1;
                        const day = String(parseInt(parts[2], 10)).padStart(2, '0');
                        if (!isNaN(monthIdx) && months[monthIdx]) {
                            return `${day} ${months[monthIdx]} ${year}`;
                        }
                    }
                }
                return dateStr;
            }

            function handleDateBlur(inputEl) {
                inputEl.type = 'text';
                if (inputEl.value) inputEl.value = formatDateToDDMMMYYYY(inputEl.value);
            }

            function formatNumber(num) {
                if (isNaN(num) || num === null) return "0.00";
                return Number(num).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            }

            function parseNumber(str) {
                if (!str) return 0;
                const cleanStr = String(str).replace(/,/g, '');
                return parseFloat(cleanStr) || 0;
            }
            
            function formatBatchNo(inputEl) {
                let val = inputEl.value.trim();
                if (!val) return;
                // 如果已经有 BATCH 开头(不分大小写)，先去掉，避免重复变成 BATCHBATCH25
                val = val.replace(/^batch/i, "").trim();
                if (val === "") { inputEl.value = ""; return; }
                inputEl.value = "BATCH" + val;
            }

            function formatNumberInput(inputEl) {
                let cursorPosition = inputEl.selectionStart;
                let originalLength = inputEl.value.length;
                let val = parseNumber(inputEl.value);
                if (inputEl.value === "" || inputEl.value === "-") return;
                if (inputEl.value.endsWith('.')) return;
                inputEl.value = val.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
                let newLength = inputEl.value.length;
                inputEl.setSelectionRange(cursorPosition + (newLength - originalLength), cursorPosition + (newLength - originalLength));
            }

            function calcTotalRental() {
                const repairDays = parseNumber(document.getElementById('repairDaysInput').value);
                const rentalPerDay = parseNumber(document.getElementById('rentalPerDayInput').value);
                const totalRental = repairDays * rentalPerDay;
                document.getElementById('totalRentalInput').value = "S$ " + formatNumber(totalRental);
            }

            function getSurveyFeeByAssessment(totalAssVal) {
                if (totalAssVal < 1.00) return 0.00;
                if (totalAssVal <= 6999.00) return 700.00;
                if (totalAssVal > 72999.00) return 3340.00;
                const steps = Math.floor((totalAssVal - 7000.00) / 1500.00);
                return 760.00 + (steps * 60.00);
            }

            function calcSurveyFees() {
                const totalAssVal = parseNumber(document.getElementById('finalTotalAss').innerText.replace("S$ ", ""));
                const surveyFee = getSurveyFeeByAssessment(totalAssVal);
                document.getElementById('surveyFeesInput').value = "S$ " + formatNumber(surveyFee);

                const transFee = parseNumber(document.getElementById('transportationInput').value);
                const photoQty = parseNumber(document.getElementById('photoQtyInput').value);
                const photoCost = photoQty * 2.00;
                document.getElementById('photographsCostInput').value = "S$ " + formatNumber(photoCost);

                const totalSurveyFee = surveyFee + transFee + photoCost;
                document.getElementById('totalSurveyFeeInput').value = "S$ " + formatNumber(totalSurveyFee);
            }

            const remarkRules = {
                "rp": { wording: "REPAIR", type: "zero" },
                "sv": { wording: "SERVICE", type: "zero" },
                "rj": { wording: "REJECTED", type: "zero" },
                "na": { wording: "NA", type: "zero" },
                "be": { wording: "BENT", type: "full" },
                "de": { wording: "DENT", type: "full" },
                "bu": { wording: "BUCKLED", type: "full" },
                "ne": { wording: "NECESSARY", type: "full" },
                "br": { wording: "BROKEN", type: "full" },
                "cr": { wording: "CRACKED", type: "full" },
                "ma": { wording: "MALFUNCTION", type: "full" },
                "df": { wording: "DEFORMED", type: "full" },
                "dp": { wording: "DEPLOYED", type: "full" },
                "ja": { wording: "JAMMED", type: "full" },
                "to": { wording: "TORN", type: "full" },
                "ru": { wording: "REUSE", type: "zero" },
                "gl": { wording: "GLAZED", type: "full" },
                "mi": { wording: "MISSING", type: "full" },
                "we": { wording: "WEAK", type: "full" },
                "cu": { wording: "CUT", type: "full" },
                "sc": { wording: "SCRATCHED", type: "full" },
                "1m": { wording: "1PC MALFUNCTION", type: "m_pcs", pcs: 1 },
                "2m": { wording: "2PCS MALFUNCTION", type: "m_pcs", pcs: 2 },
                "3m": { wording: "3PCS MALFUNCTION", type: "m_pcs", pcs: 3 },
                "4m": { wording: "4PCS MALFUNCTION", type: "m_pcs", pcs: 4 },
                "5m": { wording: "5PCS MALFUNCTION", type: "m_pcs", pcs: 5 },
                "c1": { wording: "Cut/10%", type: "cut", percent: 0.90 },
                "c2": { wording: "Cut/20%", type: "cut", percent: 0.80 },
                "c3": { wording: "Cut/30%", type: "cut", percent: 0.70 },
                "c4": { wording: "Cut/40%", type: "cut", percent: 0.60 },
                "c5": { wording: "Cut/50%", type: "cut", percent: 0.50 },
                "c6": { wording: "Cut/60%", type: "cut", percent: 0.40 },
                "c7": { wording: "Cut/70%", type: "cut", percent: 0.30 },
                "c8": { wording: "Cut/80%", type: "cut", percent: 0.20 },
                "c9": { wording: "Cut/90%", type: "cut", percent: 0.10 }
            };
            
            const PINK_HIGHLIGHT_CODES = ["ru", "sv", "na", "rp", "rj"];

            const tyrePositions = {
                "2": ["Front", "Rear"],
                "4": ["Front RHS", "Front LHS", "Rear RHS", "Rear LHS"],
                "6": ["Front RHS", "Front LHS", "Rear Outer RHS", "Rear Inner RHS", "Rear Inner LHS", "Rear Outer LHS"],
                "8": ["Front RHS", "Front LHS", "Mid RHS", "Mid LHS", "Rear Outer RHS", "Rear Inner RHS", "Rear Inner LHS", "Rear Outer LHS"],
                "10": ["Front RHS", "Front LHS", "Mid Outer RHS", "Mid Inner RHS", "Mid Inner LHS", "Mid Outer LHS", "Rear Outer RHS", "Rear Inner RHS", "Rear Inner LHS", "Rear Outer LHS"]
            };

            let isStep2Locked = false;
            let allCasesData = [];

            async function openCaseListModal() {
                try {
                    const res = await fetch('/api/get_cases');
                    const data = await res.json();
                    allCasesData = data.cases || [];
                    renderCasesTable(allCasesData);
                    document.getElementById('caseModal').classList.remove('hidden');
                } catch (e) {
                    alert("无法获取历史 Cases 数据！");
                }
            }

            function closeCaseListModal() {
                document.getElementById('caseModal').classList.add('hidden');
            }

            function renderCasesTable(cases) {
                const tbody = document.getElementById('caseListTbody');
                if (cases.length === 0) {
                    tbody.innerHTML = `<tr><td colspan="7" class="text-center p-4 text-gray-400">暂无任何保存的 Case 记录</td></tr>`;
                    return;
                }
                let html = "";
                cases.forEach(c => {
                    html += `
                        <tr class="border-b hover:bg-blue-50">
                            <td class="p-2 border font-bold text-blue-700">${c['Job Number']}</td>
                            <td class="p-2 border">${c['Batch No.']}</td>
                            <td class="p-2 border font-semibold">${c['Vehicle No.']}</td>
                            <td class="p-2 border">${c['Make & Model']}</td>
                            <td class="p-2 border">${c['Date Of Accident']}</td>
                            <td class="p-2 border text-right font-bold">S$ ${formatNumber(c['Total Repair Cost'])}</td>
                            <td class="p-2 border text-center space-x-2">
                                <button onclick="loadCaseForEdit('${c['Job Number']}')" class="bg-blue-600 hover:bg-blue-700 text-white font-bold px-2 py-1 rounded text-xs">✏️ 修改/查看</button>
                                <button onclick="confirmDeleteCase('${c['Job Number']}')" class="bg-red-600 hover:bg-red-700 text-white font-bold px-2 py-1 rounded text-xs">🗑️ 删除</button>
                            </td>
                        </tr>
                    `;
                });
                tbody.innerHTML = html;
            }

            function filterCasesTable() {
                const query = document.getElementById('caseSearchInput').value.toLowerCase();
                const filtered = allCasesData.filter(c => 
                    String(c['Job Number']).toLowerCase().includes(query) ||
                    String(c['Vehicle No.']).toLowerCase().includes(query) ||
                    String(c['Batch No.']).toLowerCase().includes(query) ||
                    String(c['Make & Model']).toLowerCase().includes(query)
                );
                renderCasesTable(filtered);
            }

            async function confirmDeleteCase(jobNo) {
                if (!confirm(`⚠️ Confirm permanent delete Case [ ${jobNo} ] ？`)) return;
                try {
                    const res = await fetch(`/api/delete_case?job_no=${encodeURIComponent(jobNo)}`, { method: 'DELETE' });
                    const result = await res.json();
                    if (result.status === "success") {
                        alert(result.message);
                        openCaseListModal();
                    } else { alert("Delete unsuccess: " + result.message); }
                } catch (e) { alert("Network request failed！"); }
            }

            function collectAllFormData() {
                const jobNo = document.getElementById('jobNo').value;
                const batchNo = document.getElementById('batchNo').value;
                const vehicleNo = document.getElementById('vehicleNo').value;
                const makeModel = document.getElementById('makeModelInput').value;
                const accidentDate = formatDateToDDMMMYYYY(document.getElementById('accidentDate').value);
                const inspectionDate = formatDateToDDMMMYYYY(document.getElementById('inspectionDate').value);
                const finishDate = document.getElementById('finishDate').value;

                const step1Data = {
                    ownerName: document.getElementById('ownerNameInput').value,
                    ownerAddr1: document.getElementById('ownerAddr1').value,
                    ownerAddr2: document.getElementById('ownerAddr2').value,
                    ownerAddr3: document.getElementById('ownerAddr3').value,
                    yearMake: document.getElementById('yearMakeInput').value,
                    chassisNo: document.getElementById('chassisNoInput').value,
                    engineNo: document.getElementById('engineNoInput').value,
                    engineCap: document.getElementById('engineCapInput').value,
                    powerRating: document.getElementById('powerRatingInput').value,
                    motorNo: document.getElementById('motorNoInput').value,
                    colour: document.getElementById('colourInput').value,
                    paintCode: document.getElementById('paintCodeSelect').value,
                    mileage: document.getElementById('mileageInput').value,
                    wheelCount: document.getElementById('wheelCount').value,
                    tyres: [],
                    damagedPos: document.getElementById('damagedPosInput').value,
                    wsName: document.getElementById('wsName').value,
                    wsAddr1: document.getElementById('wsAddr1').value,
                    wsAddr2: document.getElementById('wsAddr2').value,
                    wsAddr3: document.getElementById('wsAddr3').value
                };

                document.querySelectorAll('#tyreTableContainer tr').forEach((tr, idx) => {
                    if (idx === 0) return;
                    const inputs = tr.querySelectorAll('input');
                    if (inputs.length === 3) {
                        step1Data.tyres.push({
                            brand: inputs[0].value,
                            size: inputs[1].value,
                            mm: inputs[2].value
                        });
                    }
                });

                const parts = [];
                document.querySelectorAll('#partsTbody .parts-row').forEach(row => {
                    const rId = row.id.split('_')[1];
                    parts.push({
                        qty: document.getElementById(`partQty_${rId}`).value,
                        name: document.getElementById(`partNameInput_${rId}`).value,
                        code: document.getElementById(`partCode_${rId}`).value,
                        nett: document.getElementById(`partNett_${rId}`).value
                    });
                });

                const misc = [];
                document.querySelectorAll('#miscTbody .misc-row').forEach(row => {
                    const rId = row.id.split('_')[1];
                    misc.push({
                        qty: document.getElementById(`miscQty_${rId}`).value,
                        desc: row.querySelectorAll('input')[1].value,
                        code: document.getElementById(`miscCode_${rId}`).value,
                        est: document.getElementById(`miscEst_${rId}`).value,
                        ass: document.getElementById(`miscAss_${rId}`).value
                    });
                });

                const labourChecks = [];
                labourItems.forEach((_, idx) => {
                    const checkbox = document.querySelector(`#labourRow_${idx} input[type="checkbox"]`);
                    if (checkbox && checkbox.checked) {
                        labourChecks.push({
                            index: idx,
                            est: document.getElementById(`labourEst_${idx}`).value,
                            ass: document.getElementById(`labourAss_${idx}`).value
                        });
                    }
                });

                const customLabour = [];
                document.querySelectorAll('.custom-labour-row').forEach(row => {
                    const cId = row.id.split('_')[1];
                    customLabour.push({
                        desc: row.querySelector('input[type="text"]').value,
                        est: document.getElementById(`customLabourEst_${cId}`).value,
                        ass: document.getElementById(`customLabourAss_${cId}`).value
                    });
                });

                const totalAss = parseNumber(document.getElementById('finalTotalAss').innerText.replace("S$ ", ""));

                return {
                    job_no: jobNo,
                    batchNo: batchNo,
                    vehicleNo: vehicleNo,
                    makeModel: makeModel,
                    accidentDate: accidentDate,
                    inspectionDate: inspectionDate,
                    finishDate: finishDate,
                    totalAss: totalAss,
                    repairDays: document.getElementById('repairDaysInput').value,
                    photoQty: document.getElementById('photoQtyInput').value,
                    rentalPerDay: document.getElementById('rentalPerDayInput').value,
                    totalRental: document.getElementById('totalRentalInput').value,
                    surveyFees: document.getElementById('surveyFeesInput').value,
                    transportation: document.getElementById('transportationInput').value,
                    photographsCost: document.getElementById('photographsCostInput').value,
                    totalSurveyFee: document.getElementById('totalSurveyFeeInput').value,
                    modelCode: document.getElementById('modelCodeInput').value,
                    globalPercentage: document.getElementById('globalPercentage').value,
                    paintLabourEst: document.getElementById('paintLabourEst').value,
                    paintLabourAss: document.getElementById('paintLabourAss').value,
                    paintMatEst: document.getElementById('paintMatEst').value,
                    paintMatAss: document.getElementById('paintMatAss').value,
                    step1Data: step1Data,
                    parts: parts,
                    misc: misc,
                    labourChecks: labourChecks,
                    customLabour: customLabour
                };
            }

            async function loadCaseForEdit(jobNo) {
                try {
                    const res = await fetch(`/api/get_case_detail?job_no=${encodeURIComponent(jobNo)}`);
                    const result = await res.json();

                    if (result.status === "success") {
                        const d = result.data;

                        document.getElementById('jobNo').value = d.job_no || "";
                        document.getElementById('batchNo').value = d.batchNo || "";
                        document.getElementById('vehicleNo').value = d.vehicleNo || "";
                        document.getElementById('makeModelInput').value = d.makeModel || "";
                        document.getElementById('accidentDate').value = d.accidentDate || "";
                        document.getElementById('inspectionDate').value = d.inspectionDate || "";
                        document.getElementById('finishDate').value = d.finishDate || "";

                        const s1 = d.step1Data || {};
                        document.getElementById('ownerNameInput').value = s1.ownerName || "";
                        document.getElementById('ownerAddr1').value = s1.ownerAddr1 || "";
                        document.getElementById('ownerAddr2').value = s1.ownerAddr2 || "";
                        document.getElementById('ownerAddr3').value = s1.ownerAddr3 || "";
                        document.getElementById('yearMakeInput').value = s1.yearMake || "";
                        document.getElementById('chassisNoInput').value = s1.chassisNo || "";
                        document.getElementById('engineNoInput').value = s1.engineNo || "";
                        document.getElementById('engineCapInput').value = s1.engineCap || "";
                        document.getElementById('powerRatingInput').value = s1.powerRating || "";
                        document.getElementById('motorNoInput').value = s1.motorNo || "";
                        document.getElementById('colourInput').value = s1.colour || "";
                        document.getElementById('paintCodeSelect').value = s1.paintCode || "1";
                        document.getElementById('mileageInput').value = s1.mileage || "";
                        document.getElementById('damagedPosInput').value = s1.damagedPos || "";
                        document.getElementById('wsName').value = s1.wsName || "";
                        document.getElementById('wsAddr1').value = s1.wsAddr1 || "";
                        document.getElementById('wsAddr2').value = s1.wsAddr2 || "";
                        document.getElementById('wsAddr3').value = s1.wsAddr3 || "";

                        if (s1.wheelCount) {
                            document.getElementById('wheelCount').value = s1.wheelCount;
                            renderTyreTable();
                            if (s1.tyres) {
                                document.querySelectorAll('#tyreTableContainer tr').forEach((tr, idx) => {
                                    if (idx > 0 && s1.tyres[idx - 1]) {
                                        const inputs = tr.querySelectorAll('input');
                                        inputs[0].value = s1.tyres[idx - 1].brand || "";
                                        inputs[1].value = s1.tyres[idx - 1].size || "";
                                        inputs[2].value = s1.tyres[idx - 1].mm || "";
                                    }
                                });
                            }
                        }

                        document.getElementById('modelCodeInput').value = d.modelCode || "";
                        document.getElementById('globalPercentage').value = d.globalPercentage || 150;
                        document.getElementById('repairDaysInput').value = d.repairDays || "";
                        document.getElementById('photoQtyInput').value = d.photoQty || "";
                        document.getElementById('rentalPerDayInput').value = d.rentalPerDay || "";
                        document.getElementById('totalRentalInput').value = d.totalRental || "S$ 0.00";
                        document.getElementById('transportationInput').value = d.transportation || "200.00";
                        document.getElementById('paintLabourEst').value = d.paintLabourEst || "";
                        document.getElementById('paintLabourAss').value = d.paintLabourAss || "";
                        document.getElementById('paintMatEst').value = d.paintMatEst || "";
                        document.getElementById('paintMatAss').value = d.paintMatAss || "";

                        document.getElementById('partsTbody').innerHTML = "";
                        partsRowCounter = 0;
                        if (d.parts && d.parts.length > 0) {
                            d.parts.forEach(p => {
                                addPartsRow();
                                const rId = partsRowCounter;
                                document.getElementById(`partQty_${rId}`).value = p.qty;
                                document.getElementById(`partNameInput_${rId}`).value = p.name;
                                document.getElementById(`partCode_${rId}`).value = p.code;
                                document.getElementById(`partNett_${rId}`).value = p.nett;
                                calcPartsRow(rId);
                            });
                        }

                        document.getElementById('miscTbody').innerHTML = "";
                        miscRowCounter = 0;
                        if (d.misc && d.misc.length > 0) {
                            d.misc.forEach(m => {
                                addMiscRow();
                                const rId = miscRowCounter;
                                document.getElementById(`miscQty_${rId}`).value = m.qty;
                                document.querySelectorAll(`#miscRow_${rId} input`)[1].value = m.desc;
                                document.getElementById(`miscCode_${rId}`).value = m.code;
                                document.getElementById(`miscEst_${rId}`).value = m.est;
                                document.getElementById(`miscAss_${rId}`).value = m.ass;
                                updateMiscCondition(rId);
                            });
                        }

                        labourItems.forEach((_, idx) => {
                            const cb = document.querySelector(`#labourRow_${idx} input[type="checkbox"]`);
                            if (cb && cb.checked) cb.click();
                        });
                        if (d.labourChecks) {
                            d.labourChecks.forEach(lc => {
                                const cb = document.querySelector(`#labourRow_${lc.index} input[type="checkbox"]`);
                                if (cb && !cb.checked) {
                                    cb.click();
                                    document.getElementById(`labourEst_${lc.index}`).value = lc.est;
                                    document.getElementById(`labourAss_${lc.index}`).value = lc.ass;
                                }
                            });
                        }

                        document.getElementById('customLabourContainer').innerHTML = "";
                        customLabourCounter = 0;
                        if (d.customLabour) {
                            d.customLabour.forEach(cl => {
                                addCustomLabourRow();
                                const cId = customLabourCounter;
                                document.querySelector(`#customLabourRow_${cId} input[type="text"]`).value = cl.desc;
                                document.getElementById(`customLabourEst_${cId}`).value = cl.est;
                                document.getElementById(`customLabourAss_${cId}`).value = cl.ass;
                            });
                        }

                        fetchModelInfo();
                        calcPartsTotals();
                        calcMiscTotals();
                        calcLabourTotals();
                        calcTotalRental();

                        document.getElementById('topBarInfo').innerHTML = 
                            `<span class="text-gray-400">Batch:</span> ${d.batchNo || 'N/A'} | 
                             <span class="text-gray-400">Veh:</span> ${d.vehicleNo || 'N/A'} | 
                             <span class="text-yellow-400 font-bold">${d.job_no}</span>`;

                        // 切换页面显示
                        document.getElementById('landingPage').classList.add('hidden');
                        document.getElementById('step1Page').classList.remove('hidden');
                        document.getElementById('step2Page').classList.add('hidden');

                        enableEditStep1();
                        unlockStep2();
                        closeCaseListModal();

                        document.getElementById('bottomInvoiceBtn').classList.remove('hidden');
                        document.getElementById('bottomReportBtn').classList.remove('hidden');
                    }
                } catch (e) {
                    alert("载入 Case 失败！");
                }
            }

            async function lockAndSaveCase() {
                const jobNo = document.getElementById('jobNo').value;
                if (!jobNo || jobNo.startsWith("ERROR")) {
                    alert("Error: Invalid Job Number; unable to archive!");
                    return;
                }

                const payload = collectAllFormData();

                try {
                    const response = await fetch('/api/save_case', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify(payload)
                    });
                    const result = await response.json();

                    if (result.status === "success") {
                        isStep2Locked = true;
                        document.querySelectorAll('#step2Page input, #step2Page select, #step2Page textarea').forEach(el => el.setAttribute('disabled', 'true'));
                        document.querySelectorAll('.step2-btn, .del-btn').forEach(btn => btn.classList.add('hidden'));

                        document.getElementById('topSaveStep2Btn').classList.add('hidden');
                        document.getElementById('topEditStep2Btn').classList.remove('hidden');
                        document.getElementById('bottomSaveStep2Btn').classList.add('hidden');
                        document.getElementById('bottomEditStep2Btn').classList.remove('hidden');
                        document.getElementById('bottomInvoiceBtn').classList.remove('hidden');
                        document.getElementById('bottomReportBtn').classList.remove('hidden');

                        alert("🔒 " + result.message);
                    } else {
                        alert("存档失败: " + result.message);
                    }
                } catch (e) {
                    alert("网络请求失败，无法保存。");
                }
            }

            function downloadInvoicePDF() {
                const jobNo = document.getElementById('jobNo').value;
                if (!jobNo || jobNo.startsWith("ERROR")) {
                    alert("Please save the case before generating the Invoice PDF ! ");
                    return;
                }
                window.open(`/api/generate_invoice_pdf?job_no=${encodeURIComponent(jobNo)}`, '_blank');
            }

            function downloadSurveyReportPDF() {
                const jobNo = document.getElementById('jobNo').value;
                if (!jobNo || jobNo.startsWith("ERROR")) {
                    alert("Please save the case before generating the Survey Report PDF ! ");
                    return;
                }
                window.open(`/api/generate_survey_report_pdf?job_no=${encodeURIComponent(jobNo)}`, '_blank');
            }

            async function saveBasicInfo() {
                const inspectionDate = document.getElementById('inspectionDate').value;
                const batchNo = document.getElementById('batchNo').value;
                const vehicleNo = document.getElementById('vehicleNo').value;
                let currentJobNo = document.getElementById('jobNo').value;

                if (!currentJobNo || currentJobNo.startsWith("ERROR")) {
                    if (!inspectionDate) {
                        alert("Please select "16. Date of Inspection" first to automatically calculate and generate the Job Number ! ");
                        return;
                    }
                    const res = await fetch(`/api/generate_job_no?inspection_date=${encodeURIComponent(inspectionDate)}`);
                    const data = await res.json();
                    currentJobNo = data.job_no;
                    document.getElementById('jobNo').value = currentJobNo;
                }

                document.getElementById('topBarInfo').innerHTML = 
                    `<span class="text-gray-400">Batch:</span> ${batchNo || 'N/A'} | 
                     <span class="text-gray-400">Veh:</span> ${vehicleNo || 'N/A'} | 
                     <span class="text-yellow-400 font-bold">${currentJobNo}</span>`;

                document.querySelectorAll('.step1-input').forEach(i => i.setAttribute('disabled', true));
                document.getElementById('mainSaveBtn').classList.add('hidden');
                document.getElementById('topSaveBtn').classList.add('hidden');
                document.getElementById('topEditBtn').classList.remove('hidden');
                document.getElementById('nextBtn').classList.remove('hidden');

                await lockAndSaveCase();
            }

            async function fetchModelInfo() {
                const modelCode = document.getElementById('modelCodeInput').value.trim();
                if (!modelCode) return;
                try {
                    const res = await fetch(`/api/get_model_info?model_code=${encodeURIComponent(modelCode)}`);
                    const data = await res.json();
                    document.getElementById('modelDescriptionText').innerHTML = `Description: <span class="font-bold text-slate-800">${data.description}</span>`;
                    document.getElementById('globalPercentage').value = data.default_percentage;
                    recalcAllPartsRows();
                } catch (e) {}
            }

            async function fetchPartPriceByInput(rowId) {
                const modelCode = document.getElementById('modelCodeInput').value.trim();
                const partNameInput = document.getElementById(`partNameInput_${rowId}`);
                const partName = partNameInput ? partNameInput.value.trim() : "";
                const nettInput = document.getElementById(`partNett_${rowId}`);
                if (!modelCode || !partName) return;
                try {
                    const res = await fetch(`/api/get_part_price?model_code=${encodeURIComponent(modelCode)}&part_name=${encodeURIComponent(partName)}`);
                    const data = await res.json();
                    if (data.price > 0) {
                        nettInput.value = formatNumber(data.price);
                    } else {
                        nettInput.value = "";
                    }
                    calcPartsRow(rowId);
                } catch (e) {}
            }

            function renderTyreTable() {
                const count = document.getElementById('wheelCount').value;
                const container = document.getElementById('tyreTableContainer');
                if (!count || !tyrePositions[count]) {
                    container.innerHTML = `<p class="text-xs text-gray-400 italic">The table will automatically expand after you select the number of tires</p>`;
                    return;
                }
                let html = `<table class="w-full text-left border bg-white text-xs mt-2">
                    <tr class="bg-slate-200 font-bold">
                        <th class="p-2 border">Tyres Position</th><th class="p-2 border">Make</th><th class="p-2 border">Size</th><th class="p-2 border">Balance (mm)</th>
                    </tr>`;
                tyrePositions[count].forEach(pos => {
                    html += `<tr>
                        <td class="p-2 border font-semibold bg-slate-50">${pos}</td>
                        <td class="p-1 border"><input type="text" class="w-full border rounded p-1 step1-input" placeholder="Brand"></td>
                        <td class="p-1 border"><input type="text" class="w-full border rounded p-1 step1-input" placeholder="Size"></td>
                        <td class="p-1 border"><input type="number" step="0.1" class="w-full border rounded p-1 step1-input" placeholder="mm"></td>
                    </tr>`;
                });
                container.innerHTML = html + `</table>`;
            }

            function autoCalculateFinishDate() {
                const input = document.getElementById('inspectionDate').value;
                if (!input) return;
                let date = input.includes("-") ? new Date(input) : new Date(Date.parse(input));
                if (isNaN(date.getTime())) return;
                date.setDate(date.getDate() + 42); 
                if (date.getDay() === 6) date.setDate(date.getDate() + 2);
                if (date.getDay() === 0) date.setDate(date.getDate() + 1);
                const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
                document.getElementById('finishDate').value = `${String(date.getDate()).padStart(2, '0')} ${months[date.getMonth()]} ${date.getFullYear()}`;
            }

            function enableEditStep1() {
                document.querySelectorAll('.step1-input').forEach(i => i.removeAttribute('disabled'));
                document.getElementById('topSaveBtn').classList.remove('hidden');
                document.getElementById('mainSaveBtn').classList.remove('hidden');
                document.getElementById('topEditBtn').classList.add('hidden');
            }

            function goToStep2() {
                document.getElementById('landingPage').classList.add('hidden');
                document.getElementById('step1Page').classList.add('hidden');
                document.getElementById('step2Page').classList.remove('hidden');
                document.getElementById('backBtn').classList.remove('hidden');
                document.getElementById('topEditBtn').classList.add('hidden');
                document.getElementById('topSaveBtn').classList.add('hidden');
                if (isStep2Locked) document.getElementById('topEditStep2Btn').classList.remove('hidden');
                else document.getElementById('topSaveStep2Btn').classList.remove('hidden');
            }

            function goBackToStep1() {
                document.getElementById('step2Page').classList.add('hidden');
                document.getElementById('step1Page').classList.remove('hidden');
                document.getElementById('backBtn').classList.add('hidden');
                document.getElementById('topSaveStep2Btn').classList.add('hidden');
                document.getElementById('topEditStep2Btn').classList.add('hidden');
                document.getElementById('topEditBtn').classList.remove('hidden');
            }

            function unlockStep2() {
                isStep2Locked = false;
                document.querySelectorAll('#step2Page input, #step2Page select, #step2Page textarea').forEach(el => {
                    if (!el.readOnly) el.removeAttribute('disabled');
                });
                document.querySelectorAll('.step2-btn, .del-btn').forEach(btn => btn.classList.remove('hidden'));
                document.getElementById('topSaveStep2Btn').classList.remove('hidden');
                document.getElementById('topEditStep2Btn').classList.add('hidden');
                document.getElementById('bottomSaveStep2Btn').classList.remove('hidden');
                document.getElementById('bottomEditStep2Btn').classList.add('hidden');
            }

            let partsRowCounter = 0;
            function addPartsRow() {
                partsRowCounter++;
                const tbody = document.getElementById('partsTbody');
                const tr = document.createElement('tr');
                tr.className = "border-b hover:bg-slate-50 parts-row";
                tr.id = `partRow_${partsRowCounter}`;
                tr.innerHTML = `
                    <td class="p-2 border text-center font-bold text-gray-500 row-sn"></td>
                    <td class="p-1 border"><input type="number" value="1" min="1" onchange="calcPartsRow(${partsRowCounter})" onkeyup="calcPartsRow(${partsRowCounter})" id="partQty_${partsRowCounter}" class="w-full border rounded p-1 text-xs text-center font-semibold step2-input"></td>
                    <td class="p-1 border">
                        <input type="text" id="partNameInput_${partsRowCounter}" placeholder="Type part & press Enter..." onblur="fetchPartPriceByInput(${partsRowCounter})" onkeydown="if(event.key==='Enter'){this.blur();}" class="w-full border rounded p-1 text-xs bg-white step2-input font-medium">
                    </td>
                    <td class="p-1 border"><input type="text" placeholder="e.g. de / c2" onkeyup="calcPartsRow(${partsRowCounter})" id="partCode_${partsRowCounter}" class="w-full border rounded p-1 text-xs font-mono uppercase font-bold text-blue-600 step2-input"></td>
                    <td class="p-1 border bg-slate-50"><input type="text" readonly id="partCondition_${partsRowCounter}" class="w-full p-1 text-xs bg-transparent font-medium text-slate-700 outline-none cursor-default" placeholder="-"></td>
                    <td class="p-1 border"><input type="text" placeholder="0.00" oninput="formatNumberInput(this)" onchange="calcPartsRow(${partsRowCounter})" onkeyup="calcPartsRow(${partsRowCounter})" id="partNett_${partsRowCounter}" class="w-full border rounded p-1 text-xs font-mono text-right step2-input font-bold text-slate-700"></td>
                    <td class="p-1 border bg-slate-50"><input type="text" readonly id="partEst_${partsRowCounter}" class="w-full p-1 text-xs font-mono text-right bg-transparent text-slate-800 outline-none font-bold" placeholder="0.00"></td>
                    <td class="p-1 border bg-slate-50"><input type="text" readonly id="partAss_${partsRowCounter}" class="w-full p-1 text-xs font-mono text-right font-bold bg-transparent text-slate-800 outline-none" placeholder="0.00"></td>
                    <td class="p-1 border text-center"><button onclick="removePartsRow(${partsRowCounter})" class="del-btn text-red-500 hover:text-red-700 font-bold px-1">✕</button></td>
                `;
                tbody.appendChild(tr);
                updatePartsSN();
            }

            function removePartsRow(rowId) {
                const row = document.getElementById(`partRow_${rowId}`);
                if (row) { row.remove(); updatePartsSN(); calcPartsTotals(); }
            }

            function updatePartsSN() {
                document.querySelectorAll('#partsTbody .parts-row').forEach((row, idx) => { row.querySelector('.row-sn').innerText = idx + 1; });
            }

            function recalcAllPartsRows() {
                document.querySelectorAll('#partsTbody .parts-row').forEach(row => { calcPartsRow(row.id.split('_')[1]); });
            }

            function calcPartsRow(rowId) {
                const qty = parseNumber(document.getElementById(`partQty_${rowId}`).value);
                const nettInput = document.getElementById(`partNett_${rowId}`);
                const nettPrice = parseNumber(nettInput.value);
                const partName = document.getElementById(`partNameInput_${rowId}`).value.trim();
                const code = document.getElementById(`partCode_${rowId}`).value.trim().toLowerCase();

                if (partName !== "" && (nettInput.value === "" || nettPrice === 0)) {
                    nettInput.classList.add("bg-red-200", "text-red-900", "border-red-400");
                } else {
                    nettInput.classList.remove("bg-red-200", "text-red-900", "border-red-400");
                }

                const globalPctInput = parseNumber(document.getElementById('globalPercentage').value);
                const estimateTotal = nettPrice * qty * (globalPctInput / 100.0);
                document.getElementById(`partEst_${rowId}`).value = formatNumber(estimateTotal);

                let conditionWording = "-";
                let conditionAssMultiplier = 1.0; 

                if (code in remarkRules) {
                    const rule = remarkRules[code];
                    conditionWording = rule.wording;
                    if (rule.type === "zero") conditionAssMultiplier = 0;
                    else if (rule.type === "full") conditionAssMultiplier = 1.0;
                    else if (rule.type === "m_pcs") conditionAssMultiplier = qty > 0 ? (1.0 / qty) * rule.pcs : 0;
                    else if (rule.type === "cut") conditionAssMultiplier = 1.0 - rule.percent;
                } else if (code === "") {
                    conditionWording = "-"; 
                    conditionAssMultiplier = 1.0;
                } else {
                    conditionWording = "UNKNOWN CODE"; 
                    conditionAssMultiplier = 0;
                }

                document.getElementById(`partCondition_${rowId}`).value = conditionWording;
                document.getElementById(`partAss_${rowId}`).value = formatNumber(estimateTotal * conditionAssMultiplier);
                const partAssField = document.getElementById(`partAss_${rowId}`);
                if (PINK_HIGHLIGHT_CODES.includes(code)) {
                    partAssField.classList.remove("bg-transparent");
                    partAssField.classList.add("bg-pink-200");
                } else {
                    partAssField.classList.remove("bg-pink-200");
                    partAssField.classList.add("bg-transparent");
                }
                calcPartsTotals();
            }

            function calcPartsTotals() {
                let subEst = 0, subAss = 0;
                document.querySelectorAll('#partsTbody .parts-row').forEach(row => {
                    const rId = row.id.split('_')[1];
                    subEst += parseNumber(document.getElementById(`partEst_${rId}`).value);
                    subAss += parseNumber(document.getElementById(`partAss_${rId}`).value);
                });
                const discEst = subEst * 0.10, discAss = subAss * 0.10;
                const totalEst = subEst - discEst, totalAss = subAss - discAss;

                document.getElementById('partsSubEst').innerText = formatNumber(subEst);
                document.getElementById('partsSubAss').innerText = formatNumber(subAss);
                document.getElementById('partsDiscEst').innerText = formatNumber(discEst);
                document.getElementById('partsDiscAss').innerText = formatNumber(discAss);
                document.getElementById('partsTotalEst').innerText = formatNumber(totalEst);
                document.getElementById('partsTotalAss').innerText = formatNumber(totalAss);
                calcGrandTotals();
            }

            let miscRowCounter = 0;
            function addMiscRow() {
                miscRowCounter++;
                const tbody = document.getElementById('miscTbody');
                const tr = document.createElement('tr');
                tr.className = "border-b hover:bg-slate-50 misc-row";
                tr.id = `miscRow_${miscRowCounter}`;
                tr.innerHTML = `
                    <td class="p-2 border text-center font-bold text-gray-500 misc-sn"></td>
                    <td class="p-1 border"><input type="number" value="1" min="1" id="miscQty_${miscRowCounter}" class="w-full border rounded p-1 text-xs text-center font-semibold step2-input"></td>
                    <td class="p-1 border"><input type="text" placeholder="Misc description..." class="w-full border rounded p-1 text-xs step2-input"></td>
                    <td class="p-1 border"><input type="text" placeholder="Code..." onkeyup="updateMiscCondition(${miscRowCounter})" id="miscCode_${miscRowCounter}" class="w-full border rounded p-1 text-xs font-mono uppercase font-bold text-blue-600 step2-input"></td>
                    <td class="p-1 border bg-slate-50"><input type="text" readonly id="miscCondition_${miscRowCounter}" class="w-full p-1 text-xs bg-transparent font-medium text-slate-700 outline-none" placeholder="-"></td>
                    <td class="p-1 border"><input type="text" placeholder="0.00" oninput="formatNumberInput(this)" onkeyup="calcMiscTotals()" id="miscEst_${miscRowCounter}" class="w-full border rounded p-1 text-xs font-mono text-right step2-input"></td>
                    <td class="p-1 border"><input type="text" placeholder="0.00" oninput="formatNumberInput(this)" onkeyup="calcMiscTotals()" id="miscAss_${miscRowCounter}" class="w-full border rounded p-1 text-xs font-mono text-right font-bold text-slate-800 step2-input"></td>
                    <td class="p-1 border text-center"><button onclick="removeMiscRow(${miscRowCounter})" class="del-btn text-red-500 hover:text-red-700 font-bold px-1">✕</button></td>
                `;
                tbody.appendChild(tr);
                updateMiscSN();
            }

            function updateMiscCondition(rowId) {
                const code = document.getElementById(`miscCode_${rowId}`).value.trim().toLowerCase();
                let wording = "-";
                if (code in remarkRules) wording = remarkRules[code].wording;
                else if (code !== "") wording = "UNKNOWN CODE";
                document.getElementById(`miscCondition_${rowId}`).value = wording;
            }

            function removeMiscRow(rowId) {
                const row = document.getElementById(`miscRow_${rowId}`);
                if (row) { row.remove(); updateMiscSN(); calcMiscTotals(); }
            }

            function updateMiscSN() {
                document.querySelectorAll('#miscTbody .misc-row').forEach((row, idx) => { row.querySelector('.misc-sn').innerText = idx + 1; });
            }

            function calcMiscTotals() {
                let totalEst = 0, totalAss = 0;
                document.querySelectorAll('#miscTbody .misc-row').forEach(row => {
                    const rId = row.id.split('_')[1];
                    totalEst += parseNumber(document.getElementById(`miscEst_${rId}`).value);
                    totalAss += parseNumber(document.getElementById(`miscAss_${rId}`).value);
                });
                document.getElementById('miscTotalEst').innerText = formatNumber(totalEst);
                document.getElementById('miscTotalAss').innerText = formatNumber(totalAss);
                calcGrandTotals();
            }

            const labourItems = [
                "To remove, cut out damage portion, jack out, straighten, panel beating, welding, align and renew replaced parts.",
                "To check wirings and lightings.",
                "To remove & install reverse sensor.",
                "To remove & refix tailgate fittings to facilitate repair.",
                "To remove & refix rear upholstery, garnish and attachment.",
                "To supply sealant and to seal off all weld spot seam and gaps.",
                "To spray anti-rust coating.",
                "To remove and install rear windscreen glass.",
                "To remove and install front windscreen glass.",
                "To remove and install rear right quarter glass.",
                "To remove and install rear left quarter glass.",
                "To remove and install front right quarter glass.",
                "To remove and install front left quarter glass.",
                "To recalibrate right power gate sliding door and use diagnostic computer to reset the system.",
                "To recalibrate left power gate sliding door and use diagnostic computer to reset the system.",
                "To change radiator and related parts , and to perform coolant leakage test .",
                "To change a/c parts and to recharge a/c gas.",
                "To conduct wheel alignment.",
                "To transfer door glass, trimboard and other components to replacement door.",
                "To remove and install rear suspension unit.",
                "To remove and install front suspension unit.",
                "To remove and install front left suspension unit.",
                "To remove and install front right suspension unit.",
                "To remove and install rear left suspension unit.",
                "To remove and install rear right suspension unit.",
                "To remove and install exhaust pipes."
            ];

            function initLabourChecklist() {
                const container = document.getElementById('labourChecklist');
                labourItems.forEach((item, idx) => {
                    const div = document.createElement('div');
                    div.id = `labourRow_${idx}`;
                    div.className = "flex items-center justify-between p-2 rounded bg-slate-100 border text-xs text-slate-500";
                    div.innerHTML = `
                        <div class="flex items-center space-x-2 flex-1 mr-4">
                            <input type="checkbox" onchange="toggleLabourRow(${idx}, this.checked)" class="step2-input w-4 h-4 rounded text-blue-600 focus:ring-0">
                            <span class="font-medium">${item}</span>
                        </div>
                        <div class="flex space-x-2 hidden" id="labourInputs_${idx}">
                            <input type="text" placeholder="Est (S$)" oninput="formatNumberInput(this)" onkeyup="calcLabourTotals()" id="labourEst_${idx}" class="step2-input w-24 border rounded p-1 bg-white text-slate-800 font-mono">
                            <input type="text" placeholder="Ass (S$)" oninput="formatNumberInput(this)" onkeyup="calcLabourTotals()" id="labourAss_${idx}" class="step2-input w-24 border rounded p-1 bg-white text-slate-800 font-mono font-bold">
                        </div>
                    `;
                    container.appendChild(div);
                });
            }

            function toggleLabourRow(idx, checked) {
                const row = document.getElementById(`labourRow_${idx}`);
                const inputs = document.getElementById(`labourInputs_${idx}`);
                if (checked) {
                    row.className = "flex items-center justify-between p-2 rounded bg-blue-50 border-blue-300 text-xs text-blue-900 font-bold shadow-sm";
                    inputs.classList.remove('hidden');
                } else {
                    row.className = "flex items-center justify-between p-2 rounded bg-slate-100 border text-xs text-slate-500";
                    inputs.classList.add('hidden');
                    document.getElementById(`labourEst_${idx}`).value = "";
                    document.getElementById(`labourAss_${idx}`).value = "";
                }
                calcLabourTotals();
            }

            let customLabourCounter = 0;
            function addCustomLabourRow() {
                customLabourCounter++;
                const container = document.getElementById('customLabourContainer');
                const div = document.createElement('div');
                div.className = "flex items-center justify-between p-2 rounded bg-white border text-xs shadow-sm custom-labour-row";
                div.id = `customLabourRow_${customLabourCounter}`;
                div.innerHTML = `
                    <div class="flex-1 mr-4"><input type="text" placeholder="Custom labour description..." class="step2-input w-full border rounded p-1 text-xs"></div>
                    <div class="flex space-x-2 items-center">
                        <input type="text" placeholder="Est (S$)" oninput="formatNumberInput(this)" onkeyup="calcLabourTotals()" id="customLabourEst_${customLabourCounter}" class="step2-input w-24 border rounded p-1 bg-white font-mono">
                        <input type="text" placeholder="Ass (S$)" oninput="formatNumberInput(this)" onkeyup="calcLabourTotals()" id="customLabourAss_${customLabourCounter}" class="step2-input w-24 border rounded p-1 bg-white font-mono font-bold">
                        <button onclick="removeCustomLabourRow(${customLabourCounter})" class="del-btn text-red-500 hover:text-red-700 font-bold px-1 text-sm">✕</button>
                    </div>
                `;
                container.appendChild(div);
            }

            function removeCustomLabourRow(id) {
                const row = document.getElementById(`customLabourRow_${id}`);
                if(row) { row.remove(); calcLabourTotals(); }
            }

            function calcLabourTotals() {
                let totalEst = 0, totalAss = 0;
                labourItems.forEach((_, idx) => {
                    totalEst += parseNumber(document.getElementById(`labourEst_${idx}`)?.value);
                    totalAss += parseNumber(document.getElementById(`labourAss_${idx}`)?.value);
                });
                document.querySelectorAll('.custom-labour-row').forEach(row => {
                    const cId = row.id.split('_')[1];
                    totalEst += parseNumber(document.getElementById(`customLabourEst_${cId}`)?.value);
                    totalAss += parseNumber(document.getElementById(`customLabourAss_${cId}`)?.value);
                });
                document.getElementById('labourTotalEst').innerText = formatNumber(totalEst);
                document.getElementById('labourTotalAss').innerText = formatNumber(totalAss);
                calcGrandTotals();
            }

            function calcGrandTotals() {
                const partsEst = parseNumber(document.getElementById('partsTotalEst').innerText);
                const partsAss = parseNumber(document.getElementById('partsTotalAss').innerText);
                const miscEst = parseNumber(document.getElementById('miscTotalEst').innerText);
                const miscAss = parseNumber(document.getElementById('miscTotalAss').innerText);
                const labourEst = parseNumber(document.getElementById('labourTotalEst').innerText);
                const labourAss = parseNumber(document.getElementById('labourTotalAss').innerText);

                const paintLabourEst = parseNumber(document.getElementById('paintLabourEst').value);
                const paintLabourAss = parseNumber(document.getElementById('paintLabourAss').value);
                const paintMatEst = parseNumber(document.getElementById('paintMatEst').value);
                const paintMatAss = parseNumber(document.getElementById('paintMatAss').value);

                document.getElementById('finalPartsEst').innerText = "S$ " + formatNumber(partsEst);
                document.getElementById('finalPartsAss').innerText = "S$ " + formatNumber(partsAss);
                document.getElementById('finalMiscEst').innerText = "S$ " + formatNumber(miscEst);
                document.getElementById('finalMiscAss').innerText = "S$ " + formatNumber(miscAss);
                document.getElementById('finalLabourEst').innerText = "S$ " + formatNumber(labourEst);
                document.getElementById('finalLabourAss').innerText = "S$ " + formatNumber(labourAss);
                document.getElementById('finalPaintLabourEst').innerText = "S$ " + formatNumber(paintLabourEst);
                document.getElementById('finalPaintLabourAss').innerText = "S$ " + formatNumber(paintLabourAss);
                document.getElementById('finalPaintMatEst').innerText = "S$ " + formatNumber(paintMatEst);
                document.getElementById('finalPaintMatAss').innerText = "S$ " + formatNumber(paintMatAss);

                document.getElementById('finalTotalEst').innerText = "S$ " + formatNumber(partsEst + miscEst + labourEst + paintLabourEst + paintMatEst);
                document.getElementById('finalTotalAss').innerText = "S$ " + formatNumber(partsAss + miscAss + labourAss + paintLabourAss + paintMatAss);

                calcSurveyFees();
            }
 
            const WORKSHOP_PRESETS = {
                ming_hua: {
                    name: "Ming Hua Auto Services",
                    addr1: "Blk 160 Sin Ming Drive",
                    addr2: "#02-16 Sin Ming Autocity",
                    addr3: "Singapore 575722"
                },
                kang_auto: {
                    name: "KANG AUTO ENGINEERING PTE LTD",
                    addr1: "Blk 160 Sin Ming Drive",
                    addr2: "#02-16 Sin Ming Autocity",
                    addr3: "Singapore 575722"
                }
            };
 
            function applyWorkshopPreset() {
                const key = document.getElementById('wsPreset').value;
                if (!key) return; // 选了"手动输入"，不动现有内容
                const preset = WORKSHOP_PRESETS[key];
                if (!preset) return;
                document.getElementById('wsName').value = preset.name;
                document.getElementById('wsAddr1').value = preset.addr1;
                document.getElementById('wsAddr2').value = preset.addr2;
                document.getElementById('wsAddr3').value = preset.addr3;
            }
 
            window.onload = function() {
                addPartsRow();
                addMiscRow();
                initLabourChecklist();
            };
        </script>
    </body>
    </html>
    """
    return html_content


if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8001)