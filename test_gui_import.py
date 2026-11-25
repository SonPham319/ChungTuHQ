import os
from unittest.mock import patch
import gui

def test_gui_import():
    app = gui.AppGUI()
    with patch('tkinter.messagebox.showinfo') as mock_info:
        for idx, sample_name in enumerate(app.import_sample_combo['values']):
            app.import_sample_combo.current(idx)
            app.load_selected_import_sample()
            print(f"Sample {idx}: {sample_name}")
            tk_no = app.tokhai_info['so_tk']
            item_cnt = len(app.tokhai_info['items'])
            plhd = app.plhd_var.get()
            print(f"  TK: {tk_no}, items: {item_cnt}, PLHD: {plhd}")
            
            test_out = os.path.join('test_output', 'gui_test', sample_name)
            app.import_out_dir_var.set(os.path.abspath(test_out))
            app.generate_import_files()
            
            cv_file = os.path.join(test_out, 'CV CHIA TÁCH 2026BT-VHD.docx')
            nk_file = os.path.join(test_out, 'NhapKho.xlsx')
            assert os.path.exists(cv_file), f"Missing {cv_file}"
            assert os.path.exists(nk_file), f"Missing {nk_file}"
            print(f"  -> Generated files verified successfully!")
    app.destroy()
    print("All GUI import workflows tested successfully!")

if __name__ == '__main__':
    test_gui_import()
