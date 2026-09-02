import win32com.client
import pythoncom
import json
import re
from datetime import datetime

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
wb = excel.Workbooks.Open(r'C:\Users\X\Desktop\TL_Schedule.xlsx')

# ใช้ sheet 9月2026
ws = wb.Sheets('9月2026')

records = []
# เริ่มอ่านจาก row 6 (row 1-5 เป็น header)
for row_idx in range(6, ws.UsedRange.Rows.Count + 1):
    b = ws.Cells(row_idx, 2).Value   # 記録日
    c = ws.Cells(row_idx, 3).Value   # Quotation No.
    d = ws.Cells(row_idx, 4).Value   # Pcs
    e = ws.Cells(row_idx, 5).Value   # KG
    f = ws.Cells(row_idx, 6).Value   # Customer
    g = ws.Cells(row_idx, 7).Value   # 入荷日
    h = ws.Cells(row_idx, 8).Value   # 出荷日
    i_ = ws.Cells(row_idx, 9).Value  # ラップ出荷日
    j = ws.Cells(row_idx, 10).Value  # ラップ入荷日
    k = ws.Cells(row_idx, 11).Value  # NDK出荷日
    l = ws.Cells(row_idx, 12).Value  # NDK入荷日
    m = ws.Cells(row_idx, 13).Value  # Memo

    if c is None and f is None:
        continue  # ข้ามแถวเปล่าๆ

    def to_dt(v):
        if v is None:
            return None
        if isinstance(v, datetime):
            return v
        try:
            return v  # pywintypes.datetime
        except:
            return None

    def to_float(v):
        if v is None:
            return None
        try:
            return float(v)
        except:
            return None

    def to_str(v):
        if v is None:
            return None
        return str(v).strip()

    rec = {
        'date': to_dt(b),
        'qNo': to_str(c),
        'customer': to_str(f),
        'pcs': to_float(d),
        'kg': to_float(e),
        'arrival': to_dt(g),
        'ship': to_dt(h),
        'wrapShip': to_dt(i_),
        'wrapArr': to_dt(j),
        'ndkShip': to_dt(k),
        'ndkArr': to_dt(l),
        'memo': to_str(m),
    }
    records.append(rec)

excel.Quit()

print(f'Total records: {len(records)}')
print(json.dumps(records, ensure_ascii=False, default=str, indent=2))
