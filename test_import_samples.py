import os
import glob
import re
import docx
import openpyxl
from core.extractor import parse_tokhai_hq_xls
from core.generator import generate_cv_chia_tach_docx, generate_nhapkho_excel

def run_tests():
    samples = sorted(glob.glob('ResultExampleImport/*'))
    print(f"Found {len(samples)} sample folders.")

    for s in samples:
        folder_name = os.path.basename(s)
        print(f"\n=================== Testing {folder_name} ===================")
        xls_files = glob.glob(os.path.join(s, '*ToKhai*.xls'))
        assert len(xls_files) > 0, "No ToKhai xls found"
        tokhai_info = parse_tokhai_hq_xls(xls_files[0])
        
        m = re.search(r'IMP26-(\d+)', folder_name)
        plhd = int(m.group(1)) if m else 100
        
        out_dir = os.path.join('test_output', folder_name)
        os.makedirs(out_dir, exist_ok=True)
        
        cv_out = os.path.join(out_dir, 'CV CHIA TÁCH 2026BT-VHD.docx')
        generate_cv_chia_tach_docx(tokhai_info, '12/11/2026', cv_out)
        
        nk_out = os.path.join(out_dir, 'NhapKho.xlsx')
        generate_nhapkho_excel(tokhai_info, plhd, '12/11/2026', nk_out)
        
        # Check CV docx
        orig_cv = os.path.join(s, 'CV CHIA TÁCH 2026BT-VHD.docx')
        if os.path.exists(orig_cv):
            d_orig = docx.Document(orig_cv)
            d_gen = docx.Document(cv_out)
            for p_idx in [0, 5, 6, 7, 8, 9, 13]:
                t_orig = d_orig.paragraphs[p_idx].text.strip()
                t_gen = d_gen.paragraphs[p_idx].text.strip()
                match = (t_orig == t_gen)
                print(f"  P{p_idx} match: {match}")
                if not match:
                    print(f"    ORIG: {repr(t_orig)}")
                    print(f"    GEN : {repr(t_gen)}")
        
        # Check NhapKho.xlsx
        orig_nk = os.path.join(s, 'NhapKho.xlsx')
        if os.path.exists(orig_nk):
            wb_orig = openpyxl.load_workbook(orig_nk)
            wb_gen = openpyxl.load_workbook(nk_out)
            ws_orig = wb_orig['Inventory']
            ws_gen = wb_gen['Inventory']
            item_count = len(tokhai_info['items'])
            print(f"  Checking {item_count} items in NhapKho.xlsx...")
            for r in range(4, 4 + item_count):
                for c in [4, 5, 6, 10]:
                    v_orig = ws_orig.cell(r, c).value
                    v_gen = ws_gen.cell(r, c).value
                    assert v_orig == v_gen, f"Mismatch at R{r}C{c}: orig={v_orig} gen={v_gen}"
            print("  NhapKho.xlsx items matched 100%!")

if __name__ == '__main__':
    run_tests()
