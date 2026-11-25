import os, glob, sys, re
import openpyxl
import docx
from core.extractor import parse_delivery_order, parse_booking_awb, parse_invoice_pdf, parse_packing_list_pdf
from core.validator import validate_shipment_data
from core.generator import generate_bk_excel, generate_mau25_excel, generate_cv_dong_ghep_docx, generate_fill_script_ahk

sys.stdout.reconfigure(encoding='utf-8')

os.makedirs('test_outputs', exist_ok=True)

test_dates = {
    'EXP26-1645-AIR-SO #34803': '02/11/2026',
    'EXP26-1650-AIR-SO #36592+37616+37664+37701+38213+38290+38292+38314': '02/11/2026',
    'EXP26-1651-FDX-SO #37392': '04/11/2026',
    'EXP26-1652-DHL-SO #37238': '04/11/2026',
    'EXP26-1657-UPS-SO #36396+38371': '04/11/2026'
}

folders = sorted(glob.glob('ResultExample/*'))
for folder in folders:
    fname = os.path.basename(folder)
    print(f"\n=======================================================")
    print(f"TESTING FOLDER: {fname}")
    print(f"=======================================================")
    
    # Locate DO
    do_files = [f for f in os.listdir(folder) if 'DELIVERY' in f.upper() and f.endswith('.xlsx')]
    if not do_files:
        print("No delivery order found!")
        continue
    do_path = os.path.join(folder, do_files[0])
    
    # Locate Booking / AWB
    booking_files = [f for f in os.listdir(folder) if ('booking' in f.lower() or 'awb' in f.lower()) and not f.startswith('~')]
    if not booking_files:
        print("No booking/AWB found!")
        continue
    bk_awb_path = os.path.join(folder, booking_files[0])
    
    # Parse DO
    do_info = parse_delivery_order(do_path)
    print(f"DO: {do_info['do_code']} | Pkgs: {do_info['package_str']} | GW: {do_info['total_gw']} | Total PCS: {do_info['total_pcs']}")
    print(f"Items count: {len(do_info['grouped_items'])}, Duplicates merged: {do_info['duplicates_count']}")
    
    # Parse Booking/AWB
    booking_info = parse_booking_awb(bk_awb_path)
    print(f"Booking/AWB: Carrier: {booking_info['carrier']} | BK No: {booking_info['bk_no']} | Flight: {booking_info['flight']} | FlightDate: {booking_info['flight_date']} | Terminal: {booking_info['warehouse']}")
    
    # Validation
    val_report = validate_shipment_data(do_info, booking_info)
    print(f"Validation Status: {val_report['status']}")
    if val_report['warnings']:
        print("  Warnings:")
        for w in val_report['warnings']:
            print(f"    - {w}")
    if val_report['issues']:
        print("  Issues:")
        for iss in val_report['issues']:
            print(f"    - {iss}")
            
    # Locate Invoice & Packing List
    pdfs = sorted([f for f in os.listdir(folder) if f.lower().endswith('.pdf')])
    inv_candidates = [f for f in pdfs if not f.lower().startswith(('pl ', 'pl_', 'pl-')) and 'booking' not in f.lower() and 'awb' not in f.lower()]
    pl_candidates = [f for f in pdfs if f.lower().startswith(('pl ', 'pl_', 'pl-')) or 'packing' in f.lower()]
    inv_info = parse_invoice_pdf(os.path.join(folder, inv_candidates[0])) if inv_candidates else None
    pl_info = parse_packing_list_pdf(os.path.join(folder, pl_candidates[0])) if pl_candidates else None

    # Generate Output files
    c_date = test_dates.get(fname, '04/11/2026')
    out_dir = os.path.join('test_outputs', fname)
    os.makedirs(out_dir, exist_ok=True)
    
    gen_bk_path = os.path.join(out_dir, 'BK.xlsx')
    gen_m25_path = os.path.join(out_dir, 'mẫu 25 NQ.xlsx')
    gen_cv_path = os.path.join(out_dir, 'CV ĐÓNG GHÉP 2026BT-VHD.docx')
    gen_ahk_path = os.path.join(out_dir, 'fillScript.ahk')
    
    generate_bk_excel(do_info, booking_info, c_date, gen_bk_path)
    generate_mau25_excel(do_info, booking_info, c_date, gen_m25_path)
    cv_res = generate_cv_dong_ghep_docx(do_info, gen_cv_path)
    ahk_res = generate_fill_script_ahk(do_info, booking_info, inv_info, pl_info, c_date, gen_ahk_path)
    print(f"Generated BK -> {gen_bk_path}")
    print(f"Generated Mẫu 25 -> {gen_m25_path}")
    if cv_res:
        print(f"Generated CV Đóng Ghép -> {gen_cv_path}")
    if ahk_res:
        print(f"Generated fillScript.ahk -> {gen_ahk_path}")
    
    # Verify generated BK vs original BK
    orig_bk = os.path.join(folder, 'BK.xlsx')
    if os.path.exists(orig_bk):
        wb_orig = openpyxl.load_workbook(orig_bk, data_only=True)
        wb_gen = openpyxl.load_workbook(gen_bk_path, data_only=True)
        ws_o = wb_orig.active
        ws_g = wb_gen.active
        print(f"BK Verification: Orig max_row={ws_o.max_row}, Gen max_row={ws_g.max_row}")
        match = True
        for r in range(1, max(ws_o.max_row, ws_g.max_row)+1):
            for c in range(1, 10):
                vo = ws_o.cell(r, c).value
                vg = ws_g.cell(r, c).value
                if vo != vg and not (vo is None and vg is None):
                    # compare strings
                    if str(vo).strip() != str(vg).strip():
                        print(f"  BK Diff at ({r},{c}): Orig='{vo}' vs Gen='{vg}'")
                        match = False
        if match:
            print("  >>> BK MATCH 100%! <<<")
            
    # Verify generated Mẫu 25 vs original Mẫu 25
    orig_m25_candidates = [f for f in os.listdir(folder) if '25' in f and f.endswith('.xlsx')]
    if orig_m25_candidates:
        orig_m25 = os.path.join(folder, orig_m25_candidates[0])
        wb_m_orig = openpyxl.load_workbook(orig_m25, data_only=True)
        wb_m_gen = openpyxl.load_workbook(gen_m25_path, data_only=True)
        ws_mo = wb_m_orig['2026']
        ws_mg = wb_m_gen['2026']
        print(f"Mẫu 25 Verification: Orig max_row={ws_mo.max_row}, Gen max_row={ws_mg.max_row}")
        match_m = True
        # Compare header
        for r in [4, 5]:
            vo = ws_mo.cell(r, 10).value
            vg = ws_mg.cell(r, 10).value
            if str(vo).strip() != str(vg).strip():
                print(f"  Header Diff at ({r},10): Orig='{vo}' vs Gen='{vg}'")
        # Compare new data rows from 1190 to end
        for r in range(1190, max(ws_mo.max_row, ws_mg.max_row)+1):
            for c in range(1, 17):
                vo = ws_mo.cell(r, c).value
                vg = ws_mg.cell(r, c).value
                if str(vo).strip() != str(vg).strip() and not (vo is None and vg is None):
                    print(f"  Data Diff at ({r},{c}): Orig='{vo}' vs Gen='{vg}'")
                    match_m = False
        if match_m:
            print("  >>> MẪU 25 NQ MATCH 100%! <<<")

    # Verify generated CV Đóng Ghép vs original CV
    orig_cv = os.path.join(folder, 'CV ĐÓNG GHÉP 2026BT-VHD.docx')
    if os.path.exists(orig_cv):
        if not cv_res or not os.path.exists(gen_cv_path):
            print(f"  [LỖI] Original CV exists but gen_cv was not generated!")
        else:
            doc_orig = docx.Document(orig_cv)
            doc_gen = docx.Document(gen_cv_path)

            orig_texts = [p.text.strip() for p in doc_orig.paragraphs if p.text.strip()]
            gen_texts = [p.text.strip() for p in doc_gen.paragraphs if p.text.strip()]
            p_match = (orig_texts == gen_texts)

            t_orig = [[c.text.strip() for c in r.cells] for r in doc_orig.tables[1].rows]
            t_gen = [[c.text.strip() for c in r.cells] for r in doc_gen.tables[1].rows]
            t_match = (t_orig == t_gen)

            if p_match and t_match:
                print("  >>> CV ĐÓNG GHÉP MATCH 100%! <<<")
            else:
                if not p_match:
                    print("  [Diff] CV paragraphs do not match!")
                if not t_match:
                    print(f"  [Diff] CV table rows diff: Orig={len(t_orig)} vs Gen={len(t_gen)}")
    else:
        if cv_res:
            print(f"  [Chú ý] Generated CV Đóng Ghép for {fname} (folder has no orig CV)")
        else:
            print(f"  CV Đóng Ghép: Không cần tạo (chỉ có 1 số TK nhập) - Hợp lệ!")

    # Verify generated fillScript.ahk vs original fillScript.ahk
    orig_ahk = os.path.join(folder, 'fillScript.ahk')
    if os.path.exists(orig_ahk) and os.path.exists(gen_ahk_path):
        with open(orig_ahk, 'r', encoding='utf-8-sig') as f1:
            orig_sends = re.findall(r'send\s+"([^"]*)"', f1.read())
        with open(gen_ahk_path, 'r', encoding='utf-8-sig') as f2:
            gen_sends = re.findall(r'send\s+"([^"]*)"', f2.read())
        diffs = []
        for idx, (vo, vg) in enumerate(zip(orig_sends, gen_sends)):
            if vo != vg:
                diffs.append((idx, vo, vg))
        if not diffs and len(orig_sends) == len(gen_sends):
            print("  >>> fillScript.ahk MATCH 100%! <<<")
        else:
            print(f"  fillScript.ahk Diffs ({len(diffs)}):")
            for idx, vo, vg in diffs:
                print(f"    Send #{idx}: Orig='{vo}' vs Gen='{vg}'")



