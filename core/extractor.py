import os
import re
from datetime import datetime
from collections import defaultdict
import openpyxl
import xlrd
import pdfplumber

from data import get_country_iso

try:
    import pypdfium2 as pdfium
    from pyzbar.pyzbar import decode as decode_barcode
    HAS_BARCODE = True
except ImportError:
    HAS_BARCODE = False


def clean_desc(desc):
    """Clean model prefix and country suffix."""
    d = re.sub(r'^[A-Za-z0-9\.\-]+\#\&', '', str(desc).strip())
    d = re.sub(r'\#\&[A-Za-z0-9]+$', '', d.strip())
    return d.strip()


def get_short_product_name(desc):
    """
    Map full product description to standard Vietnamese customs summary name.
    """
    d = clean_desc(desc).lower()
    if d.startswith('ăng-ten') or d.startswith('ăng ten') or d.startswith('antenna'):
        return 'Ăng ten'
    elif 'cáp đồng trục' in d or 'dây cáp đồng trục' in d:
        return 'Cáp đồng trục'
    elif 'pallet' in d:
        return 'Pallet gỗ'
    elif 'biến áp' in d:
        return 'Biến áp'
    elif 'đầu nối' in d:
        return 'Đầu nối'
    elif 'bộ lọc' in d:
        return 'Bộ lọc'
    elif 'ăng ten' in d or 'ăng-ten' in d:
        return 'Ăng ten'
    else:
        part = d.split(',')[0].strip()
        return part[:30] if part else 'Linh kiện điện tử'


def parse_delivery_order(filepath):
    """
    Reads information from existing Delivery Order Excel file.
    Extracts header info, items, checks duplicates, and groups by TK nhập.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"File not found: {filepath}")

    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb['Sheet1'] if 'Sheet1' in wb.sheetnames else wb.active

    do_code = ""
    # Search DO code in rows 1-6
    for r in range(1, 7):
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(r, c).value or '')
            m = re.search(r'EXP\d+-\d+', val)
            if m:
                do_code = m.group(0)
                break
        if do_code:
            break

    # Locate table header
    header_row = -1
    for r in range(1, 15):
        val = str(ws.cell(r, 2).value or '')
        if 'IMPORT' in val.upper() or 'TK NHẬP' in val.upper():
            header_row = r
            break

    if header_row == -1:
        raise ValueError("Could not find table header in Delivery Order (Row with IMPORT CD / TK nhập)")

    items = []
    consignee = ""
    total_packages = 0
    package_unit = ""
    total_gw = 0.0
    total_pcs = 0

    for r in range(header_row + 1, ws.max_row + 1):
        col1 = str(ws.cell(r, 1).value or '').strip()
        col2 = ws.cell(r, 2).value

        # Match TOTAL: row first (contains weight and package unit like 10.4KGM/3CT)
        if col1.upper().startswith('TOTAL:'):
            val2 = str(ws.cell(r, 2).value or '')
            # Example: 640KGM/1PAL or 10.4KGM/3CT
            m = re.search(r'(\d+(?:\.\d+)?)\s*KGM\s*/\s*(\d+)\s*([A-Za-z]+)', val2, re.IGNORECASE)
            if m:
                if not total_gw:
                    total_gw = float(m.group(1))
                if not total_packages:
                    total_packages = int(m.group(2))
                if not package_unit:
                    package_unit = m.group(3).upper()

            # Check other cells in this row for unit like 'CT' or 'PAL'
            if not package_unit:
                for c in range(3, ws.max_column + 1):
                    cval = str(ws.cell(r, c).value or '').strip().upper()
                    if cval in ('PAL', 'CT', 'PKG', 'BOX', 'PP'):
                        package_unit = cval
                        break
            continue

        # Match TỔNG row
        if col1.upper() == 'TỔNG' or col1.upper() == 'TOTAL':
            pcs = ws.cell(r, 5).value
            if pcs is not None:
                try:
                    total_pcs = int(float(pcs))
                except (ValueError, TypeError):
                    pass
            pkgs = ws.cell(r, 7).value
            if pkgs is not None:
                try:
                    total_packages = int(float(pkgs))
                except (ValueError, TypeError):
                    pass
            gw = ws.cell(r, 8).value
            if gw is not None:
                try:
                    total_gw = float(gw)
                except (ValueError, TypeError):
                    pass
            continue

        if 'Người nhận:' in col1 or 'Consignee:' in col1:
            consignee = str(ws.cell(r, 2).value or '').strip()
            continue

        # Item row
        if col2 is not None:
            tk_nhap = str(col2).strip()
            if tk_nhap and tk_nhap.isdigit() and len(tk_nhap) >= 10:
                item_code = str(ws.cell(r, 3).value or '').strip()
                desc = str(ws.cell(r, 4).value or '').strip()
                qty = ws.cell(r, 5).value or 0
                try:
                    qty = int(float(qty))
                except (ValueError, TypeError):
                    qty = 0
                unit = str(ws.cell(r, 6).value or 'PCS').strip()

                row_pkgs = ws.cell(r, 7).value
                row_gw = ws.cell(r, 8).value
                if row_pkgs and not total_packages:
                    try:
                        total_packages = int(float(row_pkgs))
                    except (ValueError, TypeError):
                        pass
                if row_gw and not total_gw:
                    try:
                        total_gw = float(row_gw)
                    except (ValueError, TypeError):
                        pass

                items.append({
                    'tk_nhap': tk_nhap,
                    'item_code': item_code,
                    'description': desc,
                    'quantity': qty,
                    'unit': unit,
                    'row': r
                })

    # If package unit wasn't found in TOTAL: line, check columns or default
    if not package_unit:
        for r in range(header_row + 1, min(header_row + 10, ws.max_row + 1)):
            for c in range(1, ws.max_column + 1):
                val = str(ws.cell(r, c).value or '')
                m = re.search(r'\d+\s*(PAL|CT|PKG|BOX)', val, re.IGNORECASE)
                if m:
                    package_unit = m.group(1).upper()
                    break
            if package_unit:
                break

    # Display unit mapping (PAL -> PP, CT -> CT)
    display_unit = 'PP' if package_unit == 'PAL' else ('CT' if package_unit == 'CT' else (package_unit or 'PP'))

    # Deduplicate and aggregate by TK nhập
    grouped = {}
    duplicates = []

    for it in items:
        tk = it['tk_nhap']
        if tk in grouped:
            duplicates.append(it)
            grouped[tk]['quantity'] += it['quantity']
            # If current short name is Bộ lọc and new one is Ăng ten, upgrade it
            existing_short = grouped[tk]['short_name']
            new_short = get_short_product_name(it['description'])
            if existing_short == 'Bộ lọc' and new_short != 'Bộ lọc':
                grouped[tk]['short_name'] = new_short
                grouped[tk]['description'] = it['description']
        else:
            short_name = get_short_product_name(it['description'])
            grouped[tk] = {
                'tk_nhap': tk,
                'item_code': it['item_code'],
                'description': it['description'],
                'short_name': short_name,
                'quantity': it['quantity'],
                'unit': it['unit']
            }

    calculated_pcs = sum(it['quantity'] for it in grouped.values())
    if not total_pcs:
        total_pcs = calculated_pcs

    return {
        'filepath': filepath,
        'filename': os.path.basename(filepath),
        'do_code': do_code,
        'consignee': consignee,
        'total_packages': total_packages,
        'package_unit': package_unit,
        'display_unit': display_unit,
        'package_str': f"{total_packages} {display_unit}",
        'total_gw': total_gw,
        'total_pcs': total_pcs,
        'raw_items_count': len(items),
        'grouped_items': list(grouped.values()),
        'duplicates_count': len(duplicates),
        'duplicates': duplicates
    }


def parse_booking_awb(filepath):
    """
    Reads tables and text from PDF or Excel Booking Confirmation / AWB files.
    Extracts carrier, SỐ BK (MAWB/Tracking/Waybill), flight, warehouse, GW, PKG.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"File not found: {filepath}")

    fl = filepath.lower()
    info = {
        'filepath': filepath,
        'filename': os.path.basename(filepath),
        'file_type': 'excel' if fl.endswith(('.xlsx', '.xls')) else 'pdf',
        'carrier': 'AIR',
        'bk_no': '',
        'flight': '',
        'flight_date': '',
        'warehouse': '',
        'pkg': None,
        'gw': None
    }

    if info['file_type'] == 'excel':
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active
        for r in range(1, min(ws.max_row + 1, 30)):
            for c in range(1, min(ws.max_column + 1, 25)):
                val = str(ws.cell(r, c).value or '').strip()
                val_u = val.upper()
                if 'MAWB' in val_u and not info['bk_no']:
                    below = str(ws.cell(r + 1, c).value or '').strip()
                    if below and (re.search(r'\d{3}-\d+', below) or len(below) >= 8):
                        info['bk_no'] = below
                if 'FLIGHT' in val_u and not info['flight']:
                    below = str(ws.cell(r + 1, c).value or '').strip()
                    if below:
                        info['flight_date'] = below
                        info['flight'] = below.split('/')[0].strip()
                if val_u == 'WAREHOUSE' and not info['warehouse']:
                    below = str(ws.cell(r + 1, c).value or '').strip()
                    if below.upper() == 'ACS':
                        info['warehouse'] = 'ACSV'
                    else:
                        info['warehouse'] = below
                if val_u == 'PKG' and info['pkg'] is None:
                    info['pkg'] = ws.cell(r + 1, c).value
                if (val_u == 'GW' or val_u == 'G.W') and info['gw'] is None:
                    info['gw'] = ws.cell(r + 1, c).value

        # Normalize flight date with month if missing
        if info['flight_date'] and not re.search(r'[A-Za-z]', info['flight_date'].split('/')[-1]):
            now_month = datetime.now().strftime('%b').upper()
            info['flight_date'] = f"{info['flight_date']}{now_month}"

        info['carrier'] = 'AIR'
        return info

    # PDF Parsing
    full_text = ""
    with pdfplumber.open(filepath) as pdf:
        for p in pdf.pages:
            t = p.extract_text()
            if t:
                full_text += "\n" + t

    ft_upper = full_text.upper()

    if 'FEDEX' in ft_upper or 'FEDERAL EXPRESS' in ft_upper:
        info['carrier'] = 'FEDEX'
        info['warehouse'] = 'FEDEX'
        info['flight'] = 'FX6002'
        info['flight_date'] = 'FX6002'
        m = re.search(r'TRK#?\s*([\d\s]{12,18})', full_text)
        if m:
            info['bk_no'] = m.group(1).replace(' ', '').strip()
        else:
            m = re.search(r'(\d{4}\s+\d{4}\s+\d{4})', full_text)
            if m:
                info['bk_no'] = m.group(1).replace(' ', '').strip()

        m_wgt = re.search(r'TOTWGT:\s*([\d\.]+)\s*KG', full_text, re.IGNORECASE)
        if m_wgt:
            info['gw'] = float(m_wgt.group(1))

    elif 'MYDHL+' in ft_upper or 'DHL' in ft_upper:
        info['carrier'] = 'DHL'
        info['warehouse'] = 'DHL'
        info['flight'] = '8K0525'
        info['flight_date'] = '8K0525'
        m = re.search(r'WAYBILL\s*([\d\s]{10,15})', full_text, re.IGNORECASE)
        if m:
            info['bk_no'] = m.group(1).replace(' ', '').strip()

        m_wgt = re.search(r'([\d\.]+)\s*kg\s+(\d+/\d+|\d+)', full_text)
        if m_wgt:
            info['gw'] = float(m_wgt.group(1))

    elif 'UPS' in ft_upper:
        info['carrier'] = 'UPS'
        if 'AIR SHIPMENT - BOOKING CONFIRMATION' in ft_upper or 'AIR EXPORT BOOKING' in ft_upper:
            info['carrier'] = 'AIR'
            m_mawb = re.search(r'MAWB:\s*([\d\s\-]+)', full_text)
            if m_mawb:
                info['bk_no'] = m_mawb.group(1).replace(' ', '').strip()
            m_fl = re.search(r'Flight\s*(?:\(estimated\))?:\s*([A-Za-z0-9]+)(?:\s*/\s*([0-9A-Za-z\s]+))?', full_text, re.IGNORECASE)
            if m_fl:
                info['flight'] = m_fl.group(1).strip()
                date_part = m_fl.group(2).strip().split('Connecting')[0].replace(' ', '') if m_fl.group(2) else ''
                info['flight_date'] = f"{info['flight']}/{date_part}" if date_part else info['flight']
            m_term = re.search(r'Loading terminal:\s*([A-Za-z0-9]+)', full_text, re.IGNORECASE)
            if m_term:
                info['warehouse'] = m_term.group(1).strip()
            m_qty = re.search(r'Quantity:\s*(\d+)\s*PACKAGES\s*([\d\.]+)\s*kgs', full_text, re.IGNORECASE)
            if m_qty:
                info['pkg'] = int(m_qty.group(1))
                info['gw'] = float(m_qty.group(2))
        else:
            info['warehouse'] = 'UPS'
            info['flight'] = '5X0074'
            info['flight_date'] = '5X0074'
            m_1z = re.search(r'\b(1Z[0-9A-Z]{16})\b', full_text)
            if m_1z:
                info['bk_no'] = m_1z.group(1)
            elif HAS_BARCODE:
                try:
                    pdf_doc = pdfium.PdfDocument(filepath)
                    for page_idx in range(len(pdf_doc)):
                        page = pdf_doc[page_idx]
                        img = page.render(scale=2).to_pil()
                        barcodes = decode_barcode(img)
                        for b in barcodes:
                            b_str = b.data.decode('utf-8', errors='ignore')
                            if b_str.startswith('1Z') and len(b_str) == 18:
                                info['bk_no'] = b_str
                                break
                        if info['bk_no']:
                            break
                except Exception:
                    pass

    return info


def _clean_address_lines(raw_text):
    """Split block text and strip trailing Contact/Tel/Fax/Attn/Email lines."""
    if not raw_text:
        return []
    lines = []
    for line in str(raw_text).splitlines():
        s = line.strip()
        if not s:
            continue
        sl = s.lower()
        if sl.startswith(('att:', 'attn:', 'tel:', 'fax:', 'mail:', 'email:', 'contact:')):
            continue
        if '@' in s and ' ' not in s:
            continue
        lines.append(s)
    return lines


def parse_invoice_pdf(filepath):
    """
    Parses an Invoice PDF file to extract:
      - importer_name (from INVOICE TO)
      - importer_address (from INVOICE TO)
      - ship_to_address (from SHIP TO ADDRESS)
      - ship_to_country
      - ship_to_iso (2-letter ISO code from SHIP TO ADDRESS)
    """
    info = {
        'importer_name': '',
        'importer_address': '',
        'ship_to_address': '',
        'ship_to_country': '',
        'ship_to_iso': ''
    }

    invoice_to_raw = ''
    ship_to_raw = ''

    with pdfplumber.open(filepath) as pdf:
        page = pdf.pages[0]
        tables = page.extract_tables() or []
        for idx, tbl in enumerate(tables):
            # Look for the header table containing INVOICE TO / SHIP TO ADDRESS
            flat = ' '.join(str(c or '') for row in tbl for c in row).upper()
            if 'INVOICE TO' in flat and 'SHIP TO' in flat and idx + 1 < len(tables):
                next_tbl = tables[idx + 1]
                if next_tbl and len(next_tbl[0]) >= 2:
                    invoice_to_raw = next_tbl[0][0] or ''
                    ship_to_raw = next_tbl[0][1] or ''
                    break

        # Fallback: first table with >= 2 multi-line address cells
        if not invoice_to_raw and len(tables) >= 2 and tables[1] and len(tables[1][0]) >= 2:
            invoice_to_raw = tables[1][0][0] or ''
            ship_to_raw = tables[1][0][1] or ''

    # Parse INVOICE TO
    inv_lines = _clean_address_lines(invoice_to_raw)
    if inv_lines:
        name = inv_lines[0]
        addr_start_idx = 1
        # Merge wrapped company name suffix (e.g. "(DG) Ltd")
        if len(inv_lines) >= 2:
            second = inv_lines[1]
            sec_l = second.lower()
            if (second.startswith('(') or
                (len(second) <= 25 and any(k in sec_l for k in ('ltd', 'limited', 'inc', 'gmbh', 'llc', 'sdn bhd', 'jsc', 'co.')))):
                name = f"{name} {second}"
                addr_start_idx = 2
        info['importer_name'] = name.strip()
        info['importer_address'] = ' '.join(inv_lines[addr_start_idx:]).strip()

    # Parse SHIP TO ADDRESS & ISO code
    ship_lines = _clean_address_lines(ship_to_raw)
    info['ship_to_address'] = ' '.join(ship_lines).strip()
    for line in reversed(ship_lines):
        iso = get_country_iso(line)
        if iso:
            info['ship_to_country'] = line
            info['ship_to_iso'] = iso
            break
        # Try last token if postal/city is on the same line
        tokens = [t.strip(',. ') for t in re.split(r'[,\s]+', line) if t.strip(',. ')]
        if tokens:
            iso_tok = get_country_iso(tokens[-1])
            if iso_tok:
                info['ship_to_country'] = tokens[-1]
                info['ship_to_iso'] = iso_tok
                break

    return info


def _reconstruct_page_lines(page):
    """Group page characters by vertical y coordinate to handle rotated/tightly-spaced tables."""
    rows = defaultdict(list)
    for c in page.chars:
        rows[round(c['top'] / 6.0) * 6].append(c)
    lines = []
    for y in sorted(rows.keys()):
        chars = sorted(rows[y], key=lambda c: c['x0'])
        lines.append(''.join(c['text'] for c in chars))
    return lines


def parse_packing_list_pdf(filepath):
    """
    Parses a Packing List PDF file to calculate total volume in m3 (mtq):
      volume = sum(cartons_count * L * W * H) / 1_000_000
    Supports both:
      - Pallet dimensions: "Dims (include pallet)" -> L*W*H cm
      - Carton dimensions: "N Cartons @ DIMs: Lcm*Wcm*Hcm"
    """
    info = {
        'total_volume_m3': 0.0,
        'dimensions_detail': []
    }

    all_lines = []
    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ''
            all_lines.extend(txt.splitlines())
            # Also reconstruct lines from raw chars for rotated/tightly-spaced tables
            all_lines.extend(_reconstruct_page_lines(page))

    # Deduplicate consecutive identical lines while preserving order
    seen_text = '\n'.join(all_lines)

    # 1. Check for Carton DIMs pattern: "1 Cartons @ DIMs: 24cm*18cm*17cm"
    carton_matches = re.findall(
        r'(\d+)\s*Cartons?\s*@\s*DIMs?\s*:\s*(\d+(?:\.\d+)?)\s*(?:cm)?\s*\*\s*(\d+(?:\.\d+)?)\s*(?:cm)?\s*\*\s*(\d+(?:\.\d+)?)\s*(?:cm)?',
        seen_text,
        re.IGNORECASE
    )
    if carton_matches:
        # Deduplicate identical matches from text vs char reconstruction
        # Use extract_text only first to avoid double-counting!
        with pdfplumber.open(filepath) as pdf:
            plain_text = '\n'.join((p.extract_text() or '') for p in pdf.pages)
        plain_matches = re.findall(
            r'(\d+)\s*Cartons?\s*@\s*DIMs?\s*:\s*(\d+(?:\.\d+)?)\s*(?:cm)?\s*\*\s*(\d+(?:\.\d+)?)\s*(?:cm)?\s*\*\s*(\d+(?:\.\d+)?)\s*(?:cm)?',
            plain_text,
            re.IGNORECASE
        )
        matches_to_use = plain_matches if plain_matches else carton_matches
        total_cm3 = 0.0
        for cnt_s, l_s, w_s, h_s in matches_to_use:
            cnt, l, w, h = int(cnt_s), float(l_s), float(w_s), float(h_s)
            total_cm3 += cnt * l * w * h
            info['dimensions_detail'].append(f"{cnt}x({l:g}*{w:g}*{h:g}cm)")
        info['total_volume_m3'] = round(total_cm3 / 1_000_000.0, 2)
        return info

    # 2. Check for Pallet dimensions ("Dims (include pallet)")
    pallet_vols = []
    # Scan char-reconstructed lines (most reliable for both upright and rotated tables)
    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            recon_lines = _reconstruct_page_lines(page)
            has_pallet_header = any('include pallet' in ln.lower() for ln in recon_lines)
            if not has_pallet_header:
                continue
            in_pallet_section = False
            for ln in recon_lines:
                if 'include pallet' in ln.lower():
                    in_pallet_section = True
                    continue
                if in_pallet_section:
                    if 'total plt' in ln.lower() or 'total carton' in ln.lower():
                        break
                    # Look for pallet dims after 'kgs' (e.g. "640.0kgs110*110*157 cm")
                    m_pal = re.search(r'kgs\s*(\d+(?:\.\d+)?)\s*\*\s*(\d+(?:\.\d+)?)\s*\*\s*(\d+(?:\.\d+)?)\s*cm', ln, re.IGNORECASE)
                    if m_pal:
                        l, w, h = float(m_pal.group(1)), float(m_pal.group(2)), float(m_pal.group(3))
                        pallet_vols.append(l * w * h)
                        info['dimensions_detail'].append(f"Pallet({l:g}*{w:g}*{h:g}cm)")

    if pallet_vols:
        info['total_volume_m3'] = round(sum(pallet_vols) / 1_000_000.0, 2)

    return info


def extract_tokhai_item_code(desc):
    """
    Extract item code / model from customs goods description:
      1. Model: ...
      2. P/N: ...
      3. Code before #& (e.g. MA284.LBIC.001#&...)
    """
    desc = str(desc or '').strip()
    m = re.search(r'Model:\s*([A-Za-z0-9\._\-]+)', desc, re.I)
    if m:
        return m.group(1).strip(' ,;')
    m = re.search(r'P/N\s*:\s*([A-Za-z0-9\._\-]+)', desc, re.I)
    if m:
        return m.group(1).strip(' ,;')
    m = re.match(r'^([A-Za-z0-9\._\-]+)#&', desc)
    if m:
        return m.group(1).strip()
    return desc.split(',')[0].strip()


def parse_tokhai_hq_xls(filepath):
    """
    Parses an export customs declaration Excel file (ToKhaiHQ...xls / .xlsx)
    for import to warehouse:
      - Declaration header (so_tk, loai_hinh, ngay_dk, hai_quan, so_van_don, so_luong_kien, trong_luong)
      - Line items (item_code, hs_code, quantity, unit, description)
    """
    wb = xlrd.open_workbook(filepath)
    sheet_name = 'TKX' if 'TKX' in wb.sheet_names() else wb.sheet_names()[0]
    s = wb.sheet_by_name(sheet_name)

    info = {
        'filepath': filepath,
        'so_tk': '',
        'loai_hinh': '',
        'ngay_dk': '',
        'ngay_dk_raw': '',
        'hai_quan': '',
        'nguoi_xk_ma': '',
        'nguoi_xk_ten': '',
        'so_van_don': '',
        'so_luong_kien': '',
        'pkg_count': 1,
        'pkg_unit': 'PP',
        'trong_luong_str': '',
        'trong_luong_kg': 0.0,
        'items': [],
        'total_pcs': 0,
        'ten_hang': 'Ăng ten'
    }

    # 1. Header scan
    for r in range(min(s.nrows, 60)):
        for c in range(s.ncols):
            val = str(s.cell_value(r, c)).strip()
            if val == 'Số tờ khai' and not info['so_tk']:
                for c2 in range(c + 1, min(c + 10, s.ncols)):
                    v = s.cell_value(r, c2)
                    if str(v).strip():
                        info['so_tk'] = str(int(v)) if isinstance(v, float) else str(v).strip()
                        break
            elif 'Mã loại hình' in val and not info['loai_hinh']:
                for c2 in range(c + 1, min(c + 10, s.ncols)):
                    v = str(s.cell_value(r, c2)).strip()
                    if v:
                        m = re.match(r'^([A-Za-z0-9]+)', v)
                        if m:
                            info['loai_hinh'] = m.group(1)
                        break
            elif val == 'Ngày đăng ký' and not info['ngay_dk']:
                for c2 in range(c + 1, min(c + 10, s.ncols)):
                    v = str(s.cell_value(r, c2)).strip()
                    if v:
                        info['ngay_dk_raw'] = v
                        m = re.search(r'(\d{2}/\d{2}/\d{4})', v)
                        if m:
                            info['ngay_dk'] = m.group(1)
                        break
            elif val == 'Tên cơ quan Hải quan tiếp nhận tờ khai' and not info['hai_quan']:
                for c2 in range(c + 1, min(c + 10, s.ncols)):
                    v = str(s.cell_value(r, c2)).strip()
                    if v:
                        info['hai_quan'] = v
                        break
            elif val == 'Số vận đơn' and not info['so_van_don']:
                for c2 in range(c + 1, min(c + 10, s.ncols)):
                    v = s.cell_value(r, c2)
                    if str(v).strip():
                        info['so_van_don'] = str(int(v)) if isinstance(v, float) else str(v).strip()
                        break
            elif val == 'Số lượng' and not info['so_luong_kien'] and r < 45:
                v_cnt = s.cell_value(r, c + 1) or s.cell_value(r, c + 5) or ''
                cnt_str = str(int(v_cnt)) if isinstance(v_cnt, (int, float)) else str(v_cnt).strip()
                v_u = ''
                for c_u in range(c + 2, min(c + 12, s.ncols)):
                    cand = str(s.cell_value(r, c_u)).strip()
                    if cand in ('PP', 'PK', 'CT', 'PLT', 'PCE', 'KGM'):
                        v_u = cand
                        break
                info['pkg_count'] = int(cnt_str) if cnt_str.isdigit() else 1
                info['pkg_unit'] = v_u or 'PP'
                info['so_luong_kien'] = f"{cnt_str} {v_u}".strip()
            elif 'Tổng trọng lượng hàng' in val and not info['trong_luong_str']:
                for c2 in range(c + 1, min(c + 10, s.ncols)):
                    v = str(s.cell_value(r, c2)).strip()
                    if v and not v.startswith('-'):
                        info['trong_luong_str'] = v
                        clean_w = v.replace('.', '').replace(',', '.')
                        try:
                            info['trong_luong_kg'] = float(clean_w)
                        except Exception:
                            info['trong_luong_kg'] = 0.0
                        break

    # 2. Line items scan
    r = 0
    stt = 1
    while r < s.nrows:
        for c in range(min(s.ncols, 5)):
            if str(s.cell_value(r, c)).strip() == 'Mã số hàng hóa':
                hs_code = ''
                for c_hs in range(c + 1, min(c + 10, s.ncols)):
                    cand = str(s.cell_value(r, c_hs)).strip()
                    if cand.isdigit() and len(cand) >= 4:
                        hs_code = cand
                        break
                desc = ''
                qty = 0
                unit = 'PCE'
                for r2 in range(r + 1, min(r + 15, s.nrows)):
                    for c_d in range(min(s.ncols, 5)):
                        if str(s.cell_value(r2, c_d)).strip() == 'Mô tả hàng hóa':
                            desc = str(s.cell_value(r2, c_d + 3) or s.cell_value(r2, c_d + 5) or '').strip()
                    for c_q in range(s.ncols):
                        if str(s.cell_value(r2, c_q)).strip() == 'Số lượng (1)':
                            q_val = s.cell_value(r2, c_q + 2)
                            u_val = str(s.cell_value(r2, c_q + 10) or s.cell_value(r2, c_q + 8) or 'PCE').strip()
                            if isinstance(q_val, (int, float)):
                                qty = float(q_val)
                            else:
                                clean_q = str(q_val).replace('.', '').replace(',', '.')
                                qty = float(clean_q) if clean_q else 0
                            unit = u_val
                item_code = extract_tokhai_item_code(desc)
                info['items'].append({
                    'stt': stt,
                    'item_code': item_code,
                    'hs_code': hs_code,
                    'description': desc,
                    'quantity': int(qty) if qty.is_integer() else qty,
                    'unit': unit
                })
                stt += 1
                r = r2
                break
        r += 1

    info['total_pcs'] = sum(it['quantity'] for it in info['items'])

    # Determine main product name
    hs_counts = {}
    for it in info['items']:
        prefix = it['hs_code'][:4]
        hs_counts[prefix] = hs_counts.get(prefix, 0) + it['quantity']
    if hs_counts:
        top_hs = max(hs_counts.items(), key=lambda x: x[1])[0]
        if top_hs == '8517':
            info['ten_hang'] = 'Ăng ten'
        elif top_hs == '8544':
            info['ten_hang'] = 'Cáp đồng trục'
        elif top_hs == '8536':
            info['ten_hang'] = 'Đầu nối'
        else:
            info['ten_hang'] = 'Ăng ten'

    return info

