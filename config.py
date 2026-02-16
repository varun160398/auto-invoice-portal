import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(BASE_DIR, "data")

UPLOAD_DIR = os.path.join(BASE_DIR, "uploads", "signatures")
UPLOAD_EXCEL_DIR = os.path.join(BASE_DIR, "uploads", "excel")

OUT_INVOICE_DIR = os.path.join(BASE_DIR, "output", "invoices")
OUT_ZIP_DIR = os.path.join(BASE_DIR, "output", "zips")

# Not used now (excel is uploaded by user)
EXCEL_PATH = os.path.join(DATA_DIR, "experts.xlsx")
SHEET_NAME = "List"

COMPANY_NAME = "Nutritionalab Private Limited"
COMPANY_ADDR_LINES = [
    "A-2004, PHOENIX TOWER, SB MARG, DELISLE ROAD,",
    "LOWER PAREL WEST, MUMBAI - 400050",
    "PAN: AAFCN3553R",
]
