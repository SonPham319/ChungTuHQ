import os
import re
from datetime import datetime, date, timedelta
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_CELL_VERTICAL_ALIGNMENT

from data import get_hs_code


THIN_BORDER = Border(
    left=Side(style='thin', color='000000'),
    right=Side(style='thin', color='000000'),
    top=Side(style='thin', color='000000'),
    bottom=Side(style='thin', color='000000')
)

FONT_TIMES_13 = Font(name='Times New Roman', size=13, bold=False)
FONT_TIMES_13_BOLD = Font(name='Times New Roman', size=13, bold=True)
FONT_TIMES_12 = Font(name='Times New Roman', size=12, bold=False)
FONT_TIMES_12_BOLD = Font(name='Times New Roman', size=12, bold=True)

ALIGN_CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
ALIGN_LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)

def to_date_obj(dt_input):
    if isinstance(dt_input, (datetime, date)):
        return dt_input
    if isinstance(dt_input, str):
        for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%d/%m/%y'):
            try:
                return datetime.strptime(dt_input.strip(), fmt).date()
            except ValueError:
                pass
    return datetime.now().date()


def find_last_data_row(ws):
    """Find the last row in worksheet that actually contains non-empty data in col 1 or col 5."""
    for r in range(ws.max_row, 6, -1):
        v1 = ws.cell(r, 1).value
        v5 = ws.cell(r, 5).value
        if v1 is not None or v5 is not None:
            return r
    return 7


def generate_bk_excel(delivery_info, booking_info, chosen_date, output_path, base_template_path=None):
    """
    Generates BK.xlsx by loading BK_template.xlsx and injecting dynamic data,
    following the same pattern as generate_mau25_excel.
    """
    dt = to_date_obj(chosen_date)

    if not base_template_path:
        base_template_path = os.path.join(os.path.dirname(__file__), '..', 'templates', 'BK_template.xlsx')

    if not os.path.exists(base_template_path):
        raise FileNotFoundError(f"BK template not found at {base_template_path}")

    wb = openpyxl.load_workbook(base_template_path)
    ws = wb.active

    # --- Fill dynamic header cells ---
    day_str = f"{dt.day:02d}"
    month_str = f"{dt.month:02d}"
    year_str = f"{dt.year}"

    # Row 2: Số DDMMYYYY
    ws.cell(2, 1).value = f"2. Số {dt.strftime('%d%m%Y')}"

    # Row 3: Date string (merged G3:I3 already in template)
    ws.cell(3, 7).value = f"Hải Phòng, ngày {day_str} tháng {month_str} năm {year_str}"

    # Row 5: TK vận chuyển date
    ws.cell(5, 1).value = (
        f"3. Kèm theo Tờ khai vận chuyển số: . Ngày {day_str}/{month_str}/{year_str}"
        f" đăng kí tại  Hải quan Cửa khẩu Cảng Đình Vũ"
    )

    # Row 6: Package count
    pkg_str = delivery_info.get('package_str', '1 PP')
    ws.cell(6, 1).value = f"4. Số lượng hàng hóa vận chuyển thuộc tờ khai: {pkg_str}"

    # --- Determine data layout ---
    # Template has: row 9 = placeholder data row, row 10 = spacer,
    # rows 11-12 = signature rows, rows 13-14 = trailing empties.
    items = delivery_info.get('grouped_items', [])
    num_items = len(items)
    DATA_START = 9      # first data row in template
    SPACER_ROW = 10     # spacer row in template (1-item layout)
    SIG_ROW1 = 11       # "Đại diện..." row in template (1-item layout)
    SIG_ROW2 = 12       # "(Ký tên...)" row in template (1-item layout)

    # If more than 1 item, insert extra rows between DATA_START and SPACER_ROW
    # so that spacer + signature rows shift down automatically.
    extra_rows = num_items - 1
    if extra_rows > 0:
        ws.insert_rows(SPACER_ROW, extra_rows)
        # insert_rows does NOT auto-shift existing merged ranges; fix them manually.
        # Remove old signature merges (now at wrong rows) and re-add at shifted rows.
        new_sig1 = SIG_ROW1 + extra_rows
        new_sig2 = SIG_ROW2 + extra_rows
        old_sig_merges = [
            m for m in list(ws.merged_cells.ranges)
            if m.min_row >= SIG_ROW1
        ]
        for m in old_sig_merges:
            ws.merged_cells.remove(m)
        # Re-add signature merges at shifted rows (col 8 for HQ nơi đến is unmerged)
        ws.merge_cells(start_row=new_sig1, start_column=2, end_row=new_sig1, end_column=3)
        ws.merge_cells(start_row=new_sig1, start_column=5, end_row=new_sig1, end_column=6)
        ws.merge_cells(start_row=new_sig2, start_column=2, end_row=new_sig2, end_column=3)
        ws.merge_cells(start_row=new_sig2, start_column=5, end_row=new_sig2, end_column=6)
    else:
        new_sig1 = SIG_ROW1
        new_sig2 = SIG_ROW2

    # --- Write data rows ---
    flight_code = booking_info.get('flight', '')
    end_data_row = DATA_START + num_items - 1

    for idx, item in enumerate(items):
        r = DATA_START + idx
        ws.row_dimensions[r].height = 27.0
        for c in range(1, 10):
            ws.cell(r, c).border = THIN_BORDER

        c1 = ws.cell(r, 1, idx + 1)
        c1.font = FONT_TIMES_13
        c1.alignment = ALIGN_CENTER

        c2 = ws.cell(r, 2, item.get('tk_nhap', ''))
        c2.font = FONT_TIMES_13
        c2.alignment = ALIGN_CENTER

        c3 = ws.cell(r, 3, item.get('short_name', ''))
        c3.font = FONT_TIMES_13
        c3.alignment = ALIGN_CENTER

    # Merge cols D, E, F, I across all data rows; fill values in first data row only.
    # D (Container) and F (Niêm phong) are merged but left empty; E (Số chì) = 'NA'.
    if items:
        ws.cell(DATA_START, 5).value = 'NA'
        ws.cell(DATA_START, 5).font = FONT_TIMES_13
        ws.cell(DATA_START, 5).alignment = ALIGN_CENTER

        ws.cell(DATA_START, 9).value = flight_code
        ws.cell(DATA_START, 9).font = FONT_TIMES_13
        ws.cell(DATA_START, 9).alignment = ALIGN_CENTER

        if end_data_row > DATA_START:
            # Remove any stale single-cell merges from template row 9 (D9, E9, F9, I9)
            stale = [m for m in list(ws.merged_cells.ranges)
                     if m.min_row == DATA_START and m.min_col in (4, 5, 6, 9)]
            for m in stale:
                ws.merged_cells.remove(m)
            ws.merge_cells(start_row=DATA_START, start_column=4, end_row=end_data_row, end_column=4)
            ws.merge_cells(start_row=DATA_START, start_column=5, end_row=end_data_row, end_column=5)
            ws.merge_cells(start_row=DATA_START, start_column=6, end_row=end_data_row, end_column=6)
            ws.merge_cells(start_row=DATA_START, start_column=9, end_row=end_data_row, end_column=9)

    wb.save(output_path)
    return output_path


def generate_mau25_excel(delivery_info, booking_info, chosen_date, output_path, base_template_path=None):
    """
    Appends shipment rows and sets header date and BK No in mẫu 25 NQ sheet '2026'.
    """
    dt = to_date_obj(chosen_date)

    if not base_template_path:
        base_template_path = os.path.join(os.path.dirname(__file__), '..', 'templates', 'mau_25_NQ_template.xlsx')

    if not os.path.exists(base_template_path):
        raise FileNotFoundError(f"Base template not found at {base_template_path}")

    wb = openpyxl.load_workbook(base_template_path)
    sheet_name = '2026' if '2026' in wb.sheetnames else wb.sheetnames[-1]
    ws = wb[sheet_name]

    # Update header date (Row 4 Col 10) & BK No (Row 5 Col 10)
    ws.cell(4, 10).value = datetime(dt.year, dt.month, dt.day)
    ws.cell(4, 10).number_format = 'mm-dd-yy'

    bk_num = booking_info.get('bk_no', '')
    ws.cell(5, 10).value = str(bk_num)
    ws.cell(5, 10).number_format = '@'

    # Determine start row for new data based on actual last data row
    last_row = find_last_data_row(ws)
    start_row = last_row + 1
    
    # Remove any lingering merged ranges at or after start_row
    lingering = [m for m in ws.merged_cells.ranges if m.min_row >= start_row or m.max_row >= start_row]
    for m in lingering:
        ws.merged_cells.remove(m)

    items = delivery_info.get('grouped_items', [])
    num_items = len(items)

    pkg_str = delivery_info.get('package_str', '1 PP')
    flight_date_str = booking_info.get('flight_date') or booking_info.get('flight', '')
    warehouse_str = booking_info.get('warehouse', '')

    for idx, item in enumerate(items):
        r = start_row + idx
        for c in range(1, 17):
            ws.cell(r, c).border = THIN_BORDER

        # Col 1: TT
        c1 = ws.cell(r, 1, idx + 1)
        c1.font = FONT_TIMES_12
        c1.alignment = ALIGN_CENTER

        # Col 2: Lượng hàng (kiện) - on first row
        if idx == 0:
            c2 = ws.cell(r, 2, pkg_str)
            c2.font = FONT_TIMES_12
            c2.alignment = ALIGN_CENTER

        # Col 3: Số lượng (quantity PCS)
        c3 = ws.cell(r, 3, item.get('quantity', 0))
        c3.font = FONT_TIMES_12
        c3.alignment = ALIGN_CENTER

        # Col 4: Tên hàng
        c4 = ws.cell(r, 4, item.get('short_name', ''))
        c4.font = FONT_TIMES_12
        c4.alignment = ALIGN_CENTER

        # Col 5: Số tờ khai nhập (TK nhập)
        c5 = ws.cell(r, 5, item.get('tk_nhap', ''))
        c5.font = FONT_TIMES_12
        c5.alignment = ALIGN_CENTER

        # First row metadata
        if idx == 0:
            # Col 6: Tên PTVC
            c6 = ws.cell(r, 6, flight_date_str)
            c6.font = FONT_TIMES_12
            c6.alignment = ALIGN_CENTER

            # Col 7: Cửa khẩu xuất
            c7 = ws.cell(r, 7, warehouse_str)
            c7.font = FONT_TIMES_12
            c7.alignment = ALIGN_CENTER

            # Col 9: NA
            c9 = ws.cell(r, 9, 'NA')
            c9.font = FONT_TIMES_12
            c9.alignment = ALIGN_CENTER

            # Col 8: NA
            c8 = ws.cell(r, 8, '')
            c8.font = FONT_TIMES_12
            c8.alignment = ALIGN_CENTER

            # Col 10: NA
            c10 = ws.cell(r, 10, '')
            c10.font = FONT_TIMES_12
            c10.alignment = ALIGN_CENTER

            # Col 15: NGÀY ĐÓNG (same as cell(4, 10))
            c15 = ws.cell(r, 15, ws.cell(4, 10).value)
            c15.font = FONT_TIMES_12
            c15.alignment = ALIGN_CENTER
            c15.number_format = '[$-409]d\\-mmm\\-yyyy;@'

            # Col 16: ĐẠI LÝ (always TAOGLAS)
            c16 = ws.cell(r, 16, 'TAOGLAS')
            c16.font = FONT_TIMES_12
            c16.alignment = ALIGN_CENTER

    end_row = start_row + num_items - 1

    # Merge columns across all item rows if more than 1 item
    if num_items > 1:
        ws.merge_cells(start_row=start_row, start_column=2, end_row=end_row, end_column=2)
        ws.merge_cells(start_row=start_row, start_column=6, end_row=end_row, end_column=6)
        ws.merge_cells(start_row=start_row, start_column=7, end_row=end_row, end_column=7)
        ws.merge_cells(start_row=start_row, start_column=8, end_row=end_row, end_column=8)
        ws.merge_cells(start_row=start_row, start_column=9, end_row=end_row, end_column=9)
        ws.merge_cells(start_row=start_row, start_column=10, end_row=end_row, end_column=10)

    # Summary Row
    sum_row = end_row + 1
    for c in range(1, 17):
        ws.cell(sum_row, c).border = THIN_BORDER

    c_sum_1 = ws.cell(sum_row, 1, 'TT')
    c_sum_1.font = FONT_TIMES_12_BOLD
    c_sum_1.alignment = ALIGN_CENTER

    total_pcs = delivery_info.get('total_pcs', sum(it['quantity'] for it in items))
    c_sum_2 = ws.cell(sum_row, 2, f"{pkg_str} = {total_pcs} PCS")
    c_sum_2.font = FONT_TIMES_12_BOLD
    c_sum_2.alignment = ALIGN_CENTER

    ws.merge_cells(start_row=sum_row, start_column=2, end_row=sum_row, end_column=3)

    wb.save(output_path)
    return output_path


def generate_cv_dong_ghep_docx(delivery_info, output_path, base_template_path=None):
    """
    Generates 'CV ĐÓNG GHÉP 2026BT-VHD.docx' if there are 2 or more distinct
    import declaration numbers (Số TK nhập).
    Returns output_path if generated, or None if skipped (only 1 TK nhập).
    """
    items = delivery_info.get('grouped_items', [])
    distinct_tk = set(it.get('tk_nhap') for it in items if it.get('tk_nhap'))
    if len(distinct_tk) < 2:
        return None

    if not base_template_path:
        base_template_path = os.path.join(os.path.dirname(__file__), '..', 'templates', 'CV_dong_ghep_template.docx')

    if not os.path.exists(base_template_path):
        raise FileNotFoundError(f"CV template not found at {base_template_path}")

    doc = docx.Document(base_template_path)
    t = doc.tables[1]

    # Remove template placeholder rows after header (keep row 0)
    for _ in range(len(t.rows) - 1):
        t._tbl.remove(t.rows[1]._tr)

    total_pcs = sum(it.get('quantity', 0) for it in items)

    # Populate items rows
    for it in items:
        row = t.add_row()
        row.height = Cm(0.9)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        qty_str = f"{it.get('quantity', 0):,}".replace(',', '.')
        vals = [
            str(it.get('tk_nhap', '')),
            str(it.get('short_name', '')),
            qty_str,
            str(it.get('unit', 'PCS'))
        ]
        for ci, v in enumerate(vals):
            cell = row.cells[ci]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run = p.add_run(v)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(13)

    # Add summary row
    sum_row = t.add_row()
    sum_row.height = Cm(0.9)
    sum_row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    tot_str = f"{total_pcs:,}".replace(',', '.')
    sum_vals = ['', 'Tổng', tot_str, 'PCS']
    for ci, v in enumerate(sum_vals):
        cell = sum_row.cells[ci]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        if v:
            run = p.add_run(v)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(13)
            run.bold = True

    # Ensure every row in the table strictly adheres to 0.9 cm height and vertical centering
    for row in t.rows:
        row.height = Cm(0.9)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    # Parse package count and unit from package_str (e.g. '1 PP' -> 1 Pallet, '2 CT' -> 2 Carton)
    pkg_str = delivery_info.get('package_str', '1 PP')
    m = re.search(r'(\d+)\s*([A-Za-z]+)', pkg_str)
    pkg_cnt = m.group(1) if m else '1'
    pkg_raw_unit = m.group(2).upper() if m else 'PP'
    if 'PP' in pkg_raw_unit or 'PL' in pkg_raw_unit:
        pkg_unit = 'Pallet'
    elif 'CT' in pkg_raw_unit:
        pkg_unit = 'Carton'
    else:
        pkg_unit = pkg_raw_unit

    # Exclude pallet item from product count if mixed with pallet
    non_pallet = [it for it in items if 'pallet' not in str(it.get('short_name', '')).lower()]
    prod_pcs = sum(it.get('quantity', 0) for it in non_pallet) if non_pallet else total_pcs
    prod_str = f"{prod_pcs:,}".replace(',', '.')

    # Update requirement paragraph
    for p in doc.paragraphs:
        if 'Yêu cầu đóng ghép' in p.text:
            p.text = f"- Yêu cầu đóng ghép: {prod_str} PCS đóng trong: {pkg_cnt} {pkg_unit}"
            for r in p.runs:
                r.font.name = 'Times New Roman'
                r.font.size = Pt(13)

    doc.save(output_path)
    return output_path


def _ahk_escape(val):
    """Escape double quotes for AutoHotkey v2 string literals."""
    return str(val or '').replace('"', '`"')


def generate_fill_script_ahk(delivery_info, booking_info, invoice_info, pl_info, chosen_date, output_path):
    """
    Generates fillScript.ahk (AutoHotkey v2.0) combining data from:
      - delivery_info (DO)
      - booking_info (Booking/AWB)
      - invoice_info (Invoice PDF: importer name/address, ship_to_iso)
      - pl_info (Packing List PDF: total_volume_m3)
      - chosen_date (document date)
    """
    dt_start = to_date_obj(chosen_date)
    carrier = str(booking_info.get('carrier', '')).upper()

    # Transport date/time rules: FEDEX has 12 -> 17 on same day; others have 08 -> 12 next day
    start_date_str = dt_start.strftime('%d%m%Y')
    if 'FEDEX' in carrier:
        start_hour = '12'
        end_date_str = start_date_str
        end_hour = '17'
    else:
        start_hour = '08'
        end_date_str = (dt_start + timedelta(days=1)).strftime('%d%m%Y')
        end_hour = '12'

    bk_no = str(booking_info.get('bk_no', '')).strip()

    # Main product name (highest total PCS excluding PALLET) & HS code
    qty_by_name = {}
    for it in delivery_info.get('grouped_items', []):
        sname = str(it.get('short_name', '')).strip()
        if not sname or 'pallet' in sname.lower():
            continue
        qty_by_name[sname] = qty_by_name.get(sname, 0) + (it.get('quantity', 0) or 0)

    if qty_by_name:
        ten_hang = max(qty_by_name.items(), key=lambda x: x[1])[0].upper()
    else:
        items = delivery_info.get('grouped_items', [])
        ten_hang = str(items[0].get('short_name', '')).upper() if items else ''

    ma_hs = get_hs_code(ten_hang, default='') or ''

    # ISO country code from SHIP TO ADDRESS
    iso_code = str((invoice_info or {}).get('ship_to_iso', '')).strip().upper()

    # PTVC / ngày & ngày PTVC đi
    flight_date = str(booking_info.get('flight_date', '')).strip()
    flight = str(booking_info.get('flight', '')).strip()
    if '/' in flight_date:
        ptvc_ngay = flight_date
    else:
        base_fl = flight_date or flight
        next_day_tag = (dt_start + timedelta(days=1)).strftime('%d%b').upper()
        ptvc_ngay = f"{base_fl}/{next_day_tag}" if base_fl else next_day_tag

    ngay_ptvc_di = (dt_start + timedelta(days=1)).strftime('%d%m%Y')
    if '/' in ptvc_ngay:
        day_part = ptvc_ngay.split('/')[-1].strip()
        m_dt = re.match(r'^(\d{1,2})([A-Za-z]{3})', day_part)
        if m_dt:
            try:
                parsed_dep = datetime.strptime(f"{m_dt.group(1).zfill(2)}{m_dt.group(2)}{dt_start.year}", "%d%b%Y")
                ngay_ptvc_di = parsed_dep.strftime('%d%m%Y')
            except ValueError:
                pass

    # Importer name & address from INVOICE TO
    importer_name = str((invoice_info or {}).get('importer_name', '')).strip()
    importer_address = str((invoice_info or {}).get('importer_address', '')).strip()

    # Package count & unit (e.g. '1 PP' -> '1', 'pp')
    pkg_str = str(delivery_info.get('package_str', '1 PP')).strip()
    m_pkg = re.search(r'(\d+)\s*([A-Za-z]+)', pkg_str)
    pkg_cnt = m_pkg.group(1) if m_pkg else '1'
    pkg_unit = m_pkg.group(2).lower() if m_pkg else 'pp'

    # Gross weight (kgs)
    gw = float(delivery_info.get('total_gw', 0) or 0)
    gw_str = str(int(gw)) if gw.is_integer() else str(round(gw, 2))

    # Volume (m3 / mtq)
    vol = float((pl_info or {}).get('total_volume_m3', 0) or 0)
    vol_str = str(round(vol, 2))

    script_content = f'''#Requires AutoHotkey v2.0
#SingleInstance Force
SendMode "Input"
SetTitleMatchMode 2

F4::{{
\t; ngày bắt đầu vận chuyển (ngày làm chứng từ)
\tsend "{_ahk_escape(start_date_str)}"

\tsend "{{tab}}^a{{Delete}}"

\t; giờ bắt đầu vận chuyển
\tsend "{_ahk_escape(start_hour)}"

\tsend "{{tab}}"

\t; ngày kết thúc vận chuyển
\tsend "{_ahk_escape(end_date_str)}"

\tsend "{{tab}}^a{{Delete}}"

\t; giờ kết thúc vận chuyển
\tsend "{_ahk_escape(end_hour)}"

\tsend "{{tab}}{{tab}}{{tab}}{{tab}}{{tab}}{{tab}}{{tab}}{{tab}}{{tab}}{{tab}}{{tab}}{{tab}}{{tab}}{{tab}}{{tab}}{{tab}}^a{{Delete}}"
\tsend "DINH KEM BANG KE SO . DONG CHUNG XE VOI TO KHAI"
\tsend "{{tab}}^a{{Delete}}"

\t; số bk
\tsend "{_ahk_escape(bk_no)}"

\tsend "{{tab}}{{tab}}^a{{Delete}}"

\t; tên hàng
\tsend "{_ahk_escape(ten_hang)}"

\tsend "{{tab}}^a{{Delete}}"

\t; mã hs
\tsend "{_ahk_escape(ma_hs)}"

\tsend "{{tab}}{{tab}}{{tab}}{{tab}}{{tab}}{{tab}}{{tab}}{{tab}}"

\t; mã iso của nước ship đến
\tsend "{_ahk_escape(iso_code)}"

\tsend "ZZZ"
\tsend "{{tab}}{{tab}}{{tab}}^a{{Delete}}"

\t; ptvc / ngày
\tsend "{_ahk_escape(ptvc_ngay)}"

\tsend "{{tab}}{{tab}}"

\t; ngày ptvc đi
\tsend "{_ahk_escape(ngay_ptvc_di)}"

\tsend "{{tab}}{{tab}}"

\t; tên người nhập khẩu
\tsend "{_ahk_escape(importer_name)}"

\tsend "{{tab}}"

\t; địa chỉ người nhập khẩu
\tsend "{_ahk_escape(importer_address)}"

\tsend "{{tab}}{{tab}}"
\tsend "TAOGLAS LIMITED (IRELAND)"
\tsend "{{tab}}"
\tsend "U5 Kilcannon Business Park,Old Dublin Road,Enniscorthy Co. Wexford Y21 XW56 IRELAND"
\tsend "{{tab}}{{tab}}{{tab}}{{tab}}{{tab}}{{tab}}{{tab}}{{tab}}{{tab}}"

\t; số lượng kiện
\tsend "{_ahk_escape(pkg_cnt)}"

\tsend "{{tab}}"

\t; đơn vị kiện ct / pp
\tsend "{_ahk_escape(pkg_unit)}"

\tsend "{{tab}}"

\t; gw (kgs)
\tsend "{_ahk_escape(gw_str)}"

\tsend "{{tab}}"
\tsend "kgm"
\tsend "{{tab}}"

\t; thể tích
\tsend "{_ahk_escape(vol_str)}"

\tsend "{{tab}}"
\tsend "mtq"
\tsend "{{tab}}"
}}
'''

    with open(output_path, 'w', encoding='utf-8-sig') as f:
        f.write(script_content)

    return output_path


def generate_cv_chia_tach_docx(tokhai_info, chosen_date, output_path, base_template_path=None):
    """
    Generates 'CV CHIA TÁCH 2026BT-VHD.docx' for Vidifi Duyen Hai bonded warehouse.
    """
    if not base_template_path:
        base_template_path = os.path.join(os.path.dirname(__file__), '..', 'templates', 'CV_chia_tach_template.docx')

    if not os.path.exists(base_template_path):
        raise FileNotFoundError(f"CV chia tách template not found at {base_template_path}")

    doc = docx.Document(base_template_path)
    dt = to_date_obj(chosen_date)
    month_str = f"{dt.month:02d}"
    year_str = f"{dt.year}"

    

    # P5: a) Số tờ khai hải quan: ...
    p5 = doc.paragraphs[3]
    p5.text = ""
    r = p5.add_run(
        f"a) Số tờ khai hải quan: {tokhai_info.get('so_tk', '')} "
        f"loại hình {tokhai_info.get('loai_hinh', '')} "
        f"ngày {tokhai_info.get('ngay_dk', '')} đăng ký tại Hải quan Đình Vũ"
    )
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)

    # P6: b) Vận đơn số: ...
    p6 = doc.paragraphs[4]
    p6.text = ""
    r = p6.add_run(f"b) Vận đơn số: {tokhai_info.get('so_van_don', '')}")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)

    # P7: c) Số lượng kiện: ...
    p7 = doc.paragraphs[5]
    p7.text = ""
    r = p7.add_run(f"c) Số lượng kiện: {tokhai_info.get('so_luong_kien', '')}")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)

    # P8: d)  Trọng lượng hàng hóa: ...
    p8 = doc.paragraphs[6]
    p8.text = ""
    r = p8.add_run(f"d)  Trọng lượng hàng hóa: {tokhai_info.get('trong_luong_str', '')} KG")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)

    # P9: e)  Tên hàng hóa: ...
    p9 = doc.paragraphs[7]
    p9.text = ""
    r = p9.add_run(f"e)  Tên hàng hóa: {tokhai_info.get('ten_hang', '')}")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)

    # P13: Bóc tách kiện hàng: ...
    total_pcs = tokhai_info.get('total_pcs', 0)
    total_pcs_str = f"{total_pcs:,}".replace(',', '.')
    p13 = doc.paragraphs[11]
    p13.text = ""
    r = p13.add_run(f"Bóc tách kiện hàng:  {tokhai_info.get('so_luong_kien', '')} = {total_pcs_str} PCE")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc.save(output_path)
    return output_path


def generate_nhapkho_excel(tokhai_info, plhd_num, import_date, output_path, base_template_path=None):
    """
    Generates 'NhapKho.xlsx' inventory file based on customs declaration items.
    """
    if not base_template_path:
        base_template_path = os.path.join(os.path.dirname(__file__), '..', 'templates', 'NhapKho_template.xlsx')

    if not os.path.exists(base_template_path):
        raise FileNotFoundError(f"NhapKho template not found at {base_template_path}")

    wb = openpyxl.load_workbook(base_template_path)
    ws = wb['Inventory']

    items = tokhai_info.get('items', [])
    if not items:
        raise ValueError("Không có danh sách hàng hóa trong tờ khai")

    start_row = 4
    end_row = start_row + len(items) - 1

    dt = to_date_obj(import_date)
    date_val = datetime(dt.year, dt.month, dt.day, 0, 0)

    try:
        plhd_val = int(plhd_num)
    except (ValueError, TypeError):
        plhd_val = plhd_num

    so_tk = str(tokhai_info.get('so_tk', ''))
    pkg_str = f"{tokhai_info.get('pkg_count', 0)}{tokhai_info.get('pkg_unit', 'PP')}".strip()

    weight_val = tokhai_info.get('trong_luong_kg', 0.0)
    if isinstance(weight_val, (int, float)) and weight_val == int(weight_val):
        weight_val = int(weight_val)

    font_11 = Font(name='Times New Roman', size=11, bold=False)
    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')

    for idx, item in enumerate(items):
        r = start_row + idx
        for c in range(1, 14):
            cell = ws.cell(r, c)
            cell.font = font_11
            cell.border = THIN_BORDER

        # Col 4: Số tờ khai
        ws.cell(r, 4).value = so_tk
        ws.cell(r, 4).number_format = '@'
        ws.cell(r, 4).alignment = align_center

        # Col 5: Mã hàng
        ws.cell(r, 5).value = item.get('item_code', '')
        ws.cell(r, 5).alignment = align_left

        # Col 6: Số lượng
        ws.cell(r, 6).value = item.get('quantity', 0)
        ws.cell(r, 6).number_format = r'_(* #,##0_);_(* \(#,##0\);_(* "-"??_);_(@_)'
        ws.cell(r, 6).alignment = align_right

        # Col 7: Empty date format
        ws.cell(r, 7).number_format = '[$-1010000]d/m/yyyy;@'
        ws.cell(r, 7).alignment = align_center

        # Col 8: Empty count
        ws.cell(r, 8).number_format = '0'
        ws.cell(r, 8).alignment = align_center

        # Col 10: Formula =F{r}-I{r}
        ws.cell(r, 10).value = f"=F{r}-I{r}"
        ws.cell(r, 10).number_format = r'_(* #,##0_);_(* \(#,##0\);_(* "-"??_);_(@_)'
        ws.cell(r, 10).alignment = align_right

    # Merged columns values and formatting
    ws.cell(start_row, 1).value = 1
    ws.cell(start_row, 1).alignment = align_center

    ws.cell(start_row, 2).value = date_val
    ws.cell(start_row, 2).number_format = '[$-1010000]d/m/yyyy;@'
    ws.cell(start_row, 2).alignment = align_center

    ws.cell(start_row, 3).value = plhd_val
    ws.cell(start_row, 3).alignment = align_center

    ws.cell(start_row, 12).value = pkg_str
    ws.cell(start_row, 12).alignment = align_center

    ws.cell(start_row, 13).value = weight_val
    if isinstance(weight_val, float):
        ws.cell(start_row, 13).number_format = '0.00'
    else:
        ws.cell(start_row, 13).number_format = '0'
    ws.cell(start_row, 13).alignment = align_center

    # Apply merges if multiple rows
    if end_row > start_row:
        ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
        ws.merge_cells(start_row=start_row, start_column=2, end_row=end_row, end_column=2)
        ws.merge_cells(start_row=start_row, start_column=3, end_row=end_row, end_column=3)
        ws.merge_cells(start_row=start_row, start_column=12, end_row=end_row, end_column=12)
        ws.merge_cells(start_row=start_row, start_column=13, end_row=end_row, end_column=13)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    return output_path


def generate_ma_xuat_excel(delivery_info, chosen_date, output_path, base_template_path=None):
    """
    Generates 'MaXuat.xlsx' from Delivery Order items using templates/MaXuat_template.xlsx.
    Items are merged ONLY when both TK nhập and mã liệu (item_code) match.
    Columns:
      1 (A): Số TK nhập (int if numeric)
      2 (B): Mã liệu (item_code)
      3 (C): Ngày làm chứng từ (datetime, mm-dd-yy)
      4 (D): Số DO (int extracted from do_code, e.g. EXP26-1645 -> 1645)
      5 (E): Số lượng (int, summed)
    """
    if not base_template_path:
        base_template_path = os.path.join(os.path.dirname(__file__), '..', 'templates', 'MaXuat_template.xlsx')

    if not os.path.exists(base_template_path):
        raise FileNotFoundError(f"MaXuat template not found at {base_template_path}")

    wb = openpyxl.load_workbook(base_template_path)
    ws = wb['Sheet1'] if 'Sheet1' in wb.sheetnames else wb.active

    dt = to_date_obj(chosen_date)
    date_val = datetime(dt.year, dt.month, dt.day, 0, 0)

    do_code = str(delivery_info.get('do_code', '') or '').strip()
    m_do = re.search(r'-(\d+)$', do_code)
    if not m_do:
        m_do = re.search(r'(\d+)$', do_code)
    do_num = int(m_do.group(1)) if m_do else do_code

    items = delivery_info.get('ma_xuat_items')
    if items is None:
        raw_items = delivery_info.get('raw_items') or delivery_info.get('grouped_items') or []
        grouped = {}
        for it in raw_items:
            key = (it.get('tk_nhap', ''), it.get('item_code', ''))
            if key in grouped:
                grouped[key]['quantity'] += it.get('quantity', 0)
            else:
                grouped[key] = {
                    'tk_nhap': it.get('tk_nhap', ''),
                    'item_code': it.get('item_code', ''),
                    'quantity': it.get('quantity', 0),
                }
        items = list(grouped.values())

    font_11 = Font(name='Times New Roman', size=11, bold=False)
    align_center = Alignment(horizontal='center', vertical='center')
    align_vcenter = Alignment(vertical='center')

    for r, item in enumerate(items, start=1):
        tk_str = str(item.get('tk_nhap', '')).strip()
        tk_val = int(tk_str) if tk_str.isdigit() else tk_str

        ws.cell(r, 1).value = tk_val
        ws.cell(r, 2).value = item.get('item_code', '')
        ws.cell(r, 3).value = date_val
        ws.cell(r, 4).value = do_num
        ws.cell(r, 5).value = item.get('quantity', 0)

        # Ensure styling if row exceeds pre-formatted template rows
        if r > 151:
            for c in range(1, 6):
                cell = ws.cell(r, c)
                cell.font = font_11
                cell.border = THIN_BORDER
            ws.cell(r, 1).number_format = '0'
            ws.cell(r, 1).alignment = align_center
            ws.cell(r, 2).number_format = 'General'
            ws.cell(r, 2).alignment = align_center
            ws.cell(r, 3).number_format = 'mm-dd-yy'
            ws.cell(r, 3).alignment = align_center
            ws.cell(r, 4).number_format = '0'
            ws.cell(r, 4).alignment = align_center
            ws.cell(r, 5).number_format = 'General'
            ws.cell(r, 5).alignment = align_vcenter

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    return output_path



