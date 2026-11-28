import os
import sys
import re
import subprocess
import calendar
from datetime import datetime, date
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from core.extractor import (
    parse_delivery_order, parse_booking_awb, parse_invoice_pdf, parse_packing_list_pdf,
    parse_tokhai_hq_xls
)
from core.validator import validate_shipment_data
from core.generator import (
    generate_bk_excel, generate_mau25_excel, generate_cv_dong_ghep_docx, generate_fill_script_ahk,
    generate_ma_xuat_excel, generate_cv_chia_tach_docx, generate_nhapkho_excel, to_date_obj
)


class CalendarDialog(tk.Toplevel):
    """Mini calendar picker dialog for selecting dates."""
    def __init__(self, parent, initial_date=None, callback=None):
        super().__init__(parent)
        self.title("Chọn Ngày")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.callback = callback
        if initial_date:
            self.current_date = to_date_obj(initial_date)
        else:
            self.current_date = date.today()

        self.view_year = self.current_date.year
        self.view_month = self.current_date.month

        self.setup_ui()
        self.update_calendar()

        # Center on parent
        self.geometry("+%d+%d" % (parent.winfo_rootx() + 150, parent.winfo_rooty() + 150))

    def setup_ui(self):
        nav_frame = ttk.Frame(self, padding=5)
        nav_frame.pack(fill='x')

        prev_btn = ttk.Button(nav_frame, text="<", width=3, command=self.prev_month)
        prev_btn.pack(side='left', padx=2)

        self.header_label = ttk.Label(nav_frame, text="", font=('Segoe UI', 10, 'bold'))
        self.header_label.pack(side='left', expand=True)

        next_btn = ttk.Button(nav_frame, text=">", width=3, command=self.next_month)
        next_btn.pack(side='right', padx=2)

        self.cal_frame = ttk.Frame(self, padding=5)
        self.cal_frame.pack()

    def update_calendar(self):
        for widget in self.cal_frame.winfo_children():
            widget.destroy()

        self.header_label.config(text=f"Tháng {self.view_month:02d} / {self.view_year}")

        days = ["T2", "T3", "T4", "T5", "T6", "T7", "CN"]
        for c, day in enumerate(days):
            lbl = ttk.Label(self.cal_frame, text=day, font=('Segoe UI', 9, 'bold'), width=4, anchor='center')
            lbl.grid(row=0, column=c, padx=1, pady=2)

        cal = calendar.monthcalendar(self.view_year, self.view_month)
        for r, week in enumerate(cal, 1):
            for c, day in enumerate(week):
                if day != 0:
                    btn = tk.Button(
                        self.cal_frame,
                        text=str(day),
                        width=3,
                        relief='flat',
                        bg='#F1F5F9',
                        activebackground='#0284C7',
                        activeforeground='white',
                        font=('Segoe UI', 9),
                        command=lambda d=day: self.select_day(d)
                    )
                    if (self.view_year == self.current_date.year and
                        self.view_month == self.current_date.month and
                        day == self.current_date.day):
                        btn.config(bg='#0284C7', fg='white', font=('Segoe UI', 9, 'bold'))
                    btn.grid(row=r, column=c, padx=1, pady=1)

    def prev_month(self):
        if self.view_month == 1:
            self.view_month = 12
            self.view_year -= 1
        else:
            self.view_month -= 1
        self.update_calendar()

    def next_month(self):
        if self.view_month == 12:
            self.view_month = 1
            self.view_year += 1
        else:
            self.view_month += 1
        self.update_calendar()

    def select_day(self, day):
        chosen = date(self.view_year, self.view_month, day)
        if self.callback:
            self.callback(chosen)
        self.destroy()


class AppGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Hệ Thống Tự Động Xử Lý Chứng Từ Xuất / Nhập Khẩu - Antigravity")
        self.geometry("1120x860")
        self.minsize(980, 750)

        # Export workflow state
        self.delivery_info = None
        self.booking_info = None
        self.invoice_info = None
        self.pl_info = None
        self.val_report = None

        # Import workflow state
        self.tokhai_info = None
        self.last_import_cv_file = None
        self.last_import_nk_file = None

        self.init_styles()
        self.create_widgets()
        self.detect_default_samples()
        self.detect_default_import_samples()

    def init_styles(self):
        self.style = ttk.Style(self)
        try:
            self.style.theme_use('clam')
        except:
            pass

        # Color definitions
        self.PRIMARY = "#0F172A"       # Dark Navy
        self.ACCENT = "#0284C7"        # Sky Blue
        self.BG_LIGHT = "#F8FAFC"      # Off-white / light slate
        self.CARD_BG = "#FFFFFF"       # White
        self.SUCCESS = "#10B981"       # Green
        self.WARNING = "#F59E0B"       # Orange
        self.DANGER = "#EF4444"        # Red
        self.TEXT_MAIN = "#1E293B"

        self.configure(bg=self.BG_LIGHT)

        # Notebook tabs styling
        self.style.configure('TNotebook', background=self.BG_LIGHT, borderwidth=0)
        self.style.configure(
            'TNotebook.Tab',
            font=('Segoe UI', 10, 'bold'),
            padding=[20, 8],
            background='#E2E8F0',
            foreground='#475569'
        )
        self.style.map(
            'TNotebook.Tab',
            background=[('selected', self.CARD_BG), ('!selected', '#E2E8F0')],
            foreground=[('selected', self.ACCENT), ('!selected', '#64748B')]
        )

        self.style.configure('Title.TLabel', font=('Segoe UI', 16, 'bold'), foreground=self.PRIMARY, background=self.BG_LIGHT)
        self.style.configure('Subtitle.TLabel', font=('Segoe UI', 10), foreground="#64748B", background=self.BG_LIGHT)

        self.style.configure('SectionHeader.TLabel', font=('Segoe UI', 11, 'bold'), foreground=self.PRIMARY, background=self.CARD_BG)
        self.style.configure('Card.TFrame', background=self.CARD_BG, relief='flat')

        self.style.configure('Primary.TButton', font=('Segoe UI', 10, 'bold'), background=self.ACCENT, foreground='white')
        self.style.map('Primary.TButton', background=[('active', '#0369A1'), ('disabled', '#94A3B8')])

        self.style.configure('Success.TButton', font=('Segoe UI', 11, 'bold'), background=self.SUCCESS, foreground='white')
        self.style.map('Success.TButton', background=[('active', '#047857'), ('disabled', '#94A3B8')])

        self.style.configure('Outline.TButton', font=('Segoe UI', 9), background='#E2E8F0', foreground=self.TEXT_MAIN)
        self.style.map('Outline.TButton', background=[('active', '#CBD5E1')])

    def create_widgets(self):
        # Top Header Banner
        header_frame = tk.Frame(self, bg=self.PRIMARY, height=65)
        header_frame.pack(fill='x', side='top')
        header_frame.pack_propagate(False)

        title_lbl = tk.Label(
            header_frame,
            text="HỆ THỐNG XỬ LÝ CHỨNG TỪ XUẤT NHẬP KHẨU & TỰ ĐỘNG HÓA HẢI QUAN",
            font=('Segoe UI', 13, 'bold'),
            fg='white',
            bg=self.PRIMARY
        )
        title_lbl.pack(anchor='w', padx=20, pady=(8, 2))

        sub_lbl = tk.Label(
            header_frame,
            text="Xuất Khẩu: DO & Booking -> BK, Mẫu 25, CV Đóng Ghép, fillScript.ahk | Nhập Khẩu: Tờ Khai HQ -> CV Chia Tách & Nhập Kho",
            font=('Segoe UI', 9),
            fg='#94A3B8',
            bg=self.PRIMARY
        )
        sub_lbl.pack(anchor='w', padx=20)

        # Main scrollable container
        main_container = tk.Frame(self, bg=self.BG_LIGHT)
        main_container.pack(fill='both', expand=True, padx=15, pady=(8, 10))

        # Tabs container (ttk.Notebook)
        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill='both', expand=True)

        # Tab 1: Xuất Khẩu
        tab_export = tk.Frame(self.notebook, bg=self.BG_LIGHT)
        self.notebook.add(tab_export, text="   📦   XUẤT KHẨU (DO & BOOKING)   ")

        # Tab 2: Nhập Khẩu
        tab_import = tk.Frame(self.notebook, bg=self.BG_LIGHT)
        self.notebook.add(tab_import, text="   📥   NHẬP KHẨU (TỜ KHAI HQ)   ")

        # ============================================================
        # TAB 1: XUẤT KHẨU
        # ============================================================
        # Left / Top panel: Controls & File Pickers
        top_card = tk.LabelFrame(
            tab_export,
            text=" 1. Chọn Tệp Đầu Vào & Thư Mục Xuất ",
            font=('Segoe UI', 10, 'bold'),
            fg=self.PRIMARY,
            bg=self.CARD_BG,
            padx=12,
            pady=8,
            relief='groove'
        )
        top_card.pack(fill='x', pady=(0, 8))

        # Row 0: Quick Toolbar (Buttons to choose folder or batch files) + Sample presets
        r0 = tk.Frame(top_card, bg=self.CARD_BG)
        r0.pack(fill='x', pady=(0, 8))

        ttk.Button(
            r0,
            text="📁 Chọn Thư Mục Lô Hàng...",
            style='Primary.TButton',
            command=self.browse_shipment_folder
        ).pack(side='left', padx=(0, 6))

        

        ttk.Button(
            r0,
            text="🔄 Xóa Trắng / Nhập Mới",
            style='Outline.TButton',
            command=self.reset_inputs
        ).pack(side='left', padx=(0, 15))

        # Sample Presets on Right
        tk.Label(r0, text="Hoặc chọn mẫu ví dụ:", font=('Segoe UI', 9), bg=self.CARD_BG, fg='#64748B').pack(side='left', padx=(5, 5))
        self.sample_combo = ttk.Combobox(r0, state='readonly', font=('Segoe UI', 9), width=32)
        self.sample_combo.pack(side='left', padx=(0, 6))
        self.sample_combo.bind("<<ComboboxSelected>>", self.on_sample_selected)
        ttk.Button(r0, text="Nạp Mẫu", style='Outline.TButton', width=9, command=self.load_selected_sample).pack(side='left')

        # Row 1: Delivery Order File
        r1 = tk.Frame(top_card, bg=self.CARD_BG)
        r1.pack(fill='x', pady=3)
        tk.Label(r1, text="Phiếu xuất kho (DO):", font=('Segoe UI', 9, 'bold'), bg=self.CARD_BG, fg=self.TEXT_MAIN, width=18, anchor='w').pack(side='left')
        self.do_path_var = tk.StringVar()
        self.do_entry = ttk.Entry(r1, textvariable=self.do_path_var, font=('Segoe UI', 9))
        self.do_entry.pack(side='left', fill='x', expand=True, padx=(5, 8))
        ttk.Button(r1, text="📄 Chọn Tệp DO...", style='Outline.TButton', width=18, command=self.browse_do_file).pack(side='left')

        # Row 2: Booking Confirmation / AWB File
        r2 = tk.Frame(top_card, bg=self.CARD_BG)
        r2.pack(fill='x', pady=3)
        tk.Label(r2, text="Booking / AWB:", font=('Segoe UI', 9, 'bold'), bg=self.CARD_BG, fg=self.TEXT_MAIN, width=18, anchor='w').pack(side='left')
        self.bk_path_var = tk.StringVar()
        self.bk_entry = ttk.Entry(r2, textvariable=self.bk_path_var, font=('Segoe UI', 9))
        self.bk_entry.pack(side='left', fill='x', expand=True, padx=(5, 8))
        ttk.Button(r2, text="📄 Chọn Tệp Booking...", style='Outline.TButton', width=18, command=self.browse_bk_file).pack(side='left')

        # Row 2b: Invoice File (PDF)
        r_inv = tk.Frame(top_card, bg=self.CARD_BG)
        r_inv.pack(fill='x', pady=3)
        tk.Label(r_inv, text="Invoice (PDF):", font=('Segoe UI', 9, 'bold'), bg=self.CARD_BG, fg=self.TEXT_MAIN, width=18, anchor='w').pack(side='left')
        self.inv_path_var = tk.StringVar()
        self.inv_entry = ttk.Entry(r_inv, textvariable=self.inv_path_var, font=('Segoe UI', 9))
        self.inv_entry.pack(side='left', fill='x', expand=True, padx=(5, 8))
        ttk.Button(r_inv, text="📄 Chọn Invoice...", style='Outline.TButton', width=18, command=self.browse_inv_file).pack(side='left')

        # Row 2c: Packing List File (PDF)
        r_pl = tk.Frame(top_card, bg=self.CARD_BG)
        r_pl.pack(fill='x', pady=3)
        tk.Label(r_pl, text="Packing List (PDF):", font=('Segoe UI', 9, 'bold'), bg=self.CARD_BG, fg=self.TEXT_MAIN, width=18, anchor='w').pack(side='left')
        self.pl_path_var = tk.StringVar()
        self.pl_entry = ttk.Entry(r_pl, textvariable=self.pl_path_var, font=('Segoe UI', 9))
        self.pl_entry.pack(side='left', fill='x', expand=True, padx=(5, 8))
        ttk.Button(r_pl, text="📄 Chọn Packing List...", style='Outline.TButton', width=18, command=self.browse_pl_file).pack(side='left')

        # Row 3: Output Directory
        r3 = tk.Frame(top_card, bg=self.CARD_BG)
        r3.pack(fill='x', pady=3)
        tk.Label(r3, text="Thư mục xuất:", font=('Segoe UI', 9, 'bold'), bg=self.CARD_BG, fg=self.TEXT_MAIN, width=18, anchor='w').pack(side='left')
        self.out_dir_var = tk.StringVar()
        self.out_entry = ttk.Entry(r3, textvariable=self.out_dir_var, font=('Segoe UI', 9))
        self.out_entry.pack(side='left', fill='x', expand=True, padx=(5, 8))
        ttk.Button(r3, text="📁 Chọn Thư Mục Xuất...", style='Outline.TButton', width=18, command=self.browse_out_dir).pack(side='left')

        # Row 4: Date & Action Analyze Button
        r4 = tk.Frame(top_card, bg=self.CARD_BG)
        r4.pack(fill='x', pady=(6, 2))
        tk.Label(r4, text="Ngày làm chứng từ:", font=('Segoe UI', 9, 'bold'), bg=self.CARD_BG, fg=self.TEXT_MAIN, width=18, anchor='w').pack(side='left')
        self.date_var = tk.StringVar(value=datetime.now().strftime('%d/%m/%Y'))
        self.date_entry = ttk.Entry(r4, textvariable=self.date_var, font=('Segoe UI', 9, 'bold'), width=14)
        self.date_entry.pack(side='left', padx=(5, 5))
        ttk.Button(r4, text="📅 Lịch", style='Outline.TButton', width=6, command=self.open_calendar).pack(side='left', padx=2)
        ttk.Button(r4, text="Hôm nay", style='Outline.TButton', width=8, command=self.set_today).pack(side='left', padx=2)

        self.analyze_btn = ttk.Button(
            r4,
            text="  🔍  Phân Tích Dữ Liệu & Kiểm Tra Trùng Lặp  ",
            style='Primary.TButton',
            command=self.analyze_data
        )
        self.analyze_btn.pack(side='right')

        # Mid panel: Validation & Quality Control Banner
        self.qc_card = tk.LabelFrame(
            tab_export,
            text=" 2. Kết Quả Kiểm Tra Dữ Liệu & Thông Tin Trích Xuất ",
            font=('Segoe UI', 11, 'bold'),
            fg=self.PRIMARY,
            bg=self.CARD_BG,
            padx=15,
            pady=10,
            relief='groove'
        )
        self.qc_card.pack(fill='x', pady=(0, 10))

        # Summary details row
        self.summary_frame = tk.Frame(self.qc_card, bg=self.CARD_BG)
        self.summary_frame.pack(fill='x', pady=(0, 8))

        # DO info subframe
        self.do_box = tk.LabelFrame(self.summary_frame, text=" Phiếu Xuất Kho (DO) ", font=('Segoe UI', 9, 'bold'), bg='#F8FAFC', padx=10, pady=6)
        self.do_box.pack(side='left', fill='both', expand=True, padx=(0, 6))
        self.lbl_do_code = tk.Label(self.do_box, text="Mã DO: -", bg='#F8FAFC', font=('Segoe UI', 9))
        self.lbl_do_code.pack(anchor='w')
        self.lbl_consignee = tk.Label(self.do_box, text="Người nhận: -", bg='#F8FAFC', font=('Segoe UI', 9))
        self.lbl_consignee.pack(anchor='w')
        self.lbl_pkg_gw = tk.Label(self.do_box, text="Số kiện: - | GW: - | Tổng PCS: -", bg='#F8FAFC', font=('Segoe UI', 9))
        self.lbl_pkg_gw.pack(anchor='w')

        # Booking/AWB subframe
        self.bk_box = tk.LabelFrame(self.summary_frame, text=" Vận Đơn / Booking ", font=('Segoe UI', 9, 'bold'), bg='#F8FAFC', padx=10, pady=6)
        self.bk_box.pack(side='left', fill='both', expand=True, padx=6)
        self.lbl_carrier = tk.Label(self.bk_box, text="Hãng VC: - | Số BK/AWB: -", bg='#F8FAFC', font=('Segoe UI', 9))
        self.lbl_carrier.pack(anchor='w')
        self.lbl_flight = tk.Label(self.bk_box, text="Số hiệu PTVC: - | Chuyến bay/Ngày: -", bg='#F8FAFC', font=('Segoe UI', 9))
        self.lbl_flight.pack(anchor='w')
        self.lbl_wh = tk.Label(self.bk_box, text="Cửa khẩu xuất / Kho: -", bg='#F8FAFC', font=('Segoe UI', 9))
        self.lbl_wh.pack(anchor='w')

        # Quality check badge & alerts
        self.qc_box = tk.LabelFrame(self.summary_frame, text=" Tình Trạng Dữ Liệu ", font=('Segoe UI', 9, 'bold'), bg='#F8FAFC', padx=10, pady=6)
        self.qc_box.pack(side='left', fill='both', expand=True, padx=(6, 0))
        self.lbl_qc_status = tk.Label(self.qc_box, text="Chưa nạp tệp", font=('Segoe UI', 10, 'bold'), fg='#64748B', bg='#F8FAFC')
        self.lbl_qc_status.pack(anchor='w')
        self.lbl_qc_dup = tk.Label(self.qc_box, text="Trùng lặp: -", bg='#F8FAFC', font=('Segoe UI', 9))
        self.lbl_qc_dup.pack(anchor='w')
        self.lbl_qc_notes = tk.Label(self.qc_box, text="Ghi chú: Sẵn sàng phân tích", bg='#F8FAFC', font=('Segoe UI', 9), fg='#475569')
        self.lbl_qc_notes.pack(anchor='w')

        # Editable Parameters Bar
        edit_bar = tk.Frame(self.qc_card, bg=self.CARD_BG)
        edit_bar.pack(fill='x', pady=(4, 6))

        tk.Label(edit_bar, text="Hiệu chỉnh thông tin BK & Mẫu 25 trước khi xuất:", font=('Segoe UI', 9, 'bold'), fg=self.PRIMARY, bg=self.CARD_BG).pack(anchor='w', pady=(0, 4))

        eb_grid = tk.Frame(edit_bar, bg=self.CARD_BG)
        eb_grid.pack(fill='x')

        tk.Label(eb_grid, text="Số BK:", font=('Segoe UI', 9), bg=self.CARD_BG).grid(row=0, column=0, sticky='w')
        self.edit_bk_no = ttk.Entry(eb_grid, width=22, font=('Segoe UI', 9))
        self.edit_bk_no.grid(row=0, column=1, padx=(4, 15), sticky='w')

        tk.Label(eb_grid, text="Số PTVC (BK):", font=('Segoe UI', 9), bg=self.CARD_BG).grid(row=0, column=2, sticky='w')
        self.edit_flight = ttk.Entry(eb_grid, width=15, font=('Segoe UI', 9))
        self.edit_flight.grid(row=0, column=3, padx=(4, 15), sticky='w')

        tk.Label(eb_grid, text="PTVC / Ngày (M25):", font=('Segoe UI', 9), bg=self.CARD_BG).grid(row=0, column=4, sticky='w')
        self.edit_flight_date = ttk.Entry(eb_grid, width=18, font=('Segoe UI', 9))
        self.edit_flight_date.grid(row=0, column=5, padx=(4, 15), sticky='w')

        tk.Label(eb_grid, text="Cửa khẩu xuất:", font=('Segoe UI', 9), bg=self.CARD_BG).grid(row=0, column=6, sticky='w')
        self.edit_wh = ttk.Entry(eb_grid, width=14, font=('Segoe UI', 9))
        self.edit_wh.grid(row=0, column=7, padx=(4, 0), sticky='w')

        # Bottom panel: Items Table & Generate Button
        bottom_card = tk.LabelFrame(
            tab_export,
            text=" 3. Danh Sách Hàng Hóa Sau Khi Gộp Trùng Lặp & Tạo Tệp ",
            font=('Segoe UI', 10, 'bold'),
            fg=self.PRIMARY,
            bg=self.CARD_BG,
            padx=12,
            pady=6,
            relief='groove'
        )
        bottom_card.pack(fill='both', expand=True)

        # Bottom Action Bar (pack side='bottom' FIRST so it is NEVER squeezed or cut off)
        act_bar = tk.Frame(bottom_card, bg=self.CARD_BG)
        act_bar.pack(side='bottom', fill='x', pady=(6, 2))

        self.lbl_result_status = tk.Label(act_bar, text="", font=('Segoe UI', 9, 'bold'), bg=self.CARD_BG, fg=self.SUCCESS)
        self.lbl_result_status.pack(side='left')

        self.open_ahk_btn = ttk.Button(act_bar, text="Mở fillScript.ahk", style='Outline.TButton', command=self.open_generated_ahk, state='disabled')
        self.open_ahk_btn.pack(side='right', padx=(4, 0), pady=2)

        self.open_cv_btn = ttk.Button(act_bar, text="Mở CV Đóng Ghép", style='Outline.TButton', command=self.open_generated_cv, state='disabled')
        self.open_cv_btn.pack(side='right', padx=(4, 0), pady=2)

        self.open_m25_btn = ttk.Button(act_bar, text="Mở Mẫu 25 NQ", style='Outline.TButton', command=self.open_generated_m25, state='disabled')
        self.open_m25_btn.pack(side='right', padx=(4, 0), pady=2)

        self.open_bk_btn = ttk.Button(act_bar, text="Mở BK.xlsx", style='Outline.TButton', command=self.open_generated_bk, state='disabled')
        self.open_bk_btn.pack(side='right', padx=(4, 0), pady=2)

        self.open_dir_btn = ttk.Button(act_bar, text="Mở Thư Mục", style='Outline.TButton', command=self.open_output_dir, state='disabled')
        self.open_dir_btn.pack(side='right', padx=(4, 0), pady=2)

        self.gen_btn = ttk.Button(
            act_bar,
            text="  ⚡  TẠO TÀI LIỆU  ",
            style='Success.TButton',
            command=self.generate_files,
            state='disabled'
        )
        self.gen_btn.pack(side='right', padx=(8, 10), pady=2)

        # Table treeview (takes remaining vertical space)
        tree_frame = tk.Frame(bottom_card, bg=self.CARD_BG)
        tree_frame.pack(side='top', fill='both', expand=True, pady=(0, 4))

        cols = ("stt", "tk_nhap", "ten_hang", "so_luong", "don_vi", "mo_ta")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=6)
        self.tree.heading("stt", text="STT")
        self.tree.heading("tk_nhap", text="Số TK Nhập")
        self.tree.heading("ten_hang", text="Tên Hàng (Tiếng Việt)")
        self.tree.heading("so_luong", text="Số Lượng (PCS)")
        self.tree.heading("don_vi", text="Đơn Vị")
        self.tree.heading("mo_ta", text="Mô Tả Gốc / Mã Liệu")

        self.tree.column("stt", width=50, anchor='center')
        self.tree.column("tk_nhap", width=140, anchor='center')
        self.tree.column("ten_hang", width=160, anchor='center')
        self.tree.column("so_luong", width=120, anchor='center')
        self.tree.column("don_vi", width=80, anchor='center')
        self.tree.column("mo_ta", width=420, anchor='w')

        scroll_y = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)

        self.tree.pack(side='left', fill='both', expand=True)
        scroll_y.pack(side='right', fill='y')

        self.last_bk_file = None
        self.last_m25_file = None
        self.last_cv_file = None
        self.last_ahk_file = None

        # Setup Tab 2: Nhập Khẩu
        self.setup_import_tab(tab_import)

    def setup_import_tab(self, tab_import):
        """Builds all UI components for Tab 2: Nhập Khẩu (Tờ khai HQ -> CV Chia Tách & NhapKho)."""
        # 1. Top Card: Input files & destination
        top_import_card = tk.LabelFrame(
            tab_import,
            text=" 1. Chọn Tệp Tờ Khai Hải Quan & Thư Mục Xuất ",
            font=('Segoe UI', 10, 'bold'),
            fg=self.PRIMARY,
            bg=self.CARD_BG,
            padx=12,
            pady=8,
            relief='groove'
        )
        top_import_card.pack(fill='x', pady=(0, 8))

        # Row 0: Quick Toolbar (Folder picker, Reset, Sample presets)
        r0 = tk.Frame(top_import_card, bg=self.CARD_BG)
        r0.pack(fill='x', pady=(0, 8))

        ttk.Button(
            r0,
            text="📁 Chọn Thư Mục Lô Hàng...",
            style='Primary.TButton',
            command=self.browse_import_folder
        ).pack(side='left', padx=(0, 6))

        ttk.Button(
            r0,
            text="🔄 Xóa Trắng / Nhập Mới",
            style='Outline.TButton',
            command=self.reset_import_inputs
        ).pack(side='left', padx=(0, 15))

        tk.Label(r0, text="Hoặc chọn mẫu ví dụ:", font=('Segoe UI', 9), bg=self.CARD_BG, fg='#64748B').pack(side='left', padx=(5, 5))
        self.import_sample_combo = ttk.Combobox(r0, state='readonly', font=('Segoe UI', 9), width=36)
        self.import_sample_combo.pack(side='left', padx=(0, 6))
        self.import_sample_combo.bind("<<ComboboxSelected>>", self.on_import_sample_selected)
        ttk.Button(r0, text="Nạp Mẫu", style='Outline.TButton', width=9, command=self.load_selected_import_sample).pack(side='left')

        # Row 1: Tờ khai HQ Excel file
        r1 = tk.Frame(top_import_card, bg=self.CARD_BG)
        r1.pack(fill='x', pady=3)
        tk.Label(r1, text="Tờ khai HQ (Excel):", font=('Segoe UI', 9, 'bold'), bg=self.CARD_BG, fg=self.TEXT_MAIN, width=18, anchor='w').pack(side='left')
        self.tokhai_path_var = tk.StringVar()
        self.tokhai_entry = ttk.Entry(r1, textvariable=self.tokhai_path_var, font=('Segoe UI', 9))
        self.tokhai_entry.pack(side='left', fill='x', expand=True, padx=(5, 8))
        ttk.Button(r1, text="📄 Chọn Tệp Tờ Khai...", style='Outline.TButton', width=18, command=self.browse_tokhai_file).pack(side='left')

        # Row 2: Số PLHĐ & Ngày nhập kho Vidifi
        r2 = tk.Frame(top_import_card, bg=self.CARD_BG)
        r2.pack(fill='x', pady=3)
        tk.Label(r2, text="Số PLHĐ (Phụ lục):", font=('Segoe UI', 9, 'bold'), bg=self.CARD_BG, fg=self.TEXT_MAIN, width=18, anchor='w').pack(side='left')
        self.plhd_var = tk.StringVar(value="359")
        self.plhd_entry = ttk.Entry(r2, textvariable=self.plhd_var, font=('Segoe UI', 9, 'bold'), width=12)
        self.plhd_entry.pack(side='left', padx=(5, 15))

        tk.Label(r2, text="Ngày nhập kho Vidifi:", font=('Segoe UI', 9, 'bold'), bg=self.CARD_BG, fg=self.TEXT_MAIN).pack(side='left', padx=(5, 5))
        self.import_date_var = tk.StringVar(value=datetime.now().strftime('%d/%m/%Y'))
        self.import_date_entry = ttk.Entry(r2, textvariable=self.import_date_var, font=('Segoe UI', 9, 'bold'), width=12)
        self.import_date_entry.pack(side='left', padx=(0, 5))
        ttk.Button(r2, text="📅 Lịch", style='Outline.TButton', width=6, command=self.open_import_calendar).pack(side='left', padx=2)
        ttk.Button(r2, text="Hôm nay", style='Outline.TButton', width=8, command=self.set_import_today).pack(side='left', padx=2)

        # Row 3: Output Directory
        r3 = tk.Frame(top_import_card, bg=self.CARD_BG)
        r3.pack(fill='x', pady=3)
        tk.Label(r3, text="Thư mục xuất:", font=('Segoe UI', 9, 'bold'), bg=self.CARD_BG, fg=self.TEXT_MAIN, width=18, anchor='w').pack(side='left')
        self.import_out_dir_var = tk.StringVar()
        self.import_out_entry = ttk.Entry(r3, textvariable=self.import_out_dir_var, font=('Segoe UI', 9))
        self.import_out_entry.pack(side='left', fill='x', expand=True, padx=(5, 8))
        ttk.Button(r3, text="📁 Chọn Thư Mục Xuất...", style='Outline.TButton', width=18, command=self.browse_import_out_dir).pack(side='left')

        # Row 4: Action button
        r4 = tk.Frame(top_import_card, bg=self.CARD_BG)
        r4.pack(fill='x', pady=(6, 2))
        self.import_analyze_btn = ttk.Button(
            r4,
            text="  🔍  Phân Tích Tờ Khai Hải Quan & Trích Xuất Hàng Hóa  ",
            style='Primary.TButton',
            command=self.analyze_import_data
        )
        self.import_analyze_btn.pack(side='right')

        # 2. Mid Card: QC / Summary
        self.import_qc_card = tk.LabelFrame(
            tab_import,
            text=" 2. Kết Quả Trích Xuất Tờ Khai Hải Quan & Tùy Chỉnh ",
            font=('Segoe UI', 11, 'bold'),
            fg=self.PRIMARY,
            bg=self.CARD_BG,
            padx=15,
            pady=10,
            relief='groove'
        )
        self.import_qc_card.pack(fill='x', pady=(0, 10))

        summary_frame = tk.Frame(self.import_qc_card, bg=self.CARD_BG)
        summary_frame.pack(fill='x', pady=(0, 8))

        # Subframe 1: Tờ khai
        box1 = tk.LabelFrame(summary_frame, text=" Thông Tin Tờ Khai ", font=('Segoe UI', 9, 'bold'), bg='#F8FAFC', padx=10, pady=6)
        box1.pack(side='left', fill='both', expand=True, padx=(0, 6))
        self.lbl_import_tk = tk.Label(box1, text="Số tờ khai: -", bg='#F8FAFC', font=('Segoe UI', 9))
        self.lbl_import_tk.pack(anchor='w')
        self.lbl_import_lh_ngay = tk.Label(box1, text="Loại hình: - | Ngày ĐK: -", bg='#F8FAFC', font=('Segoe UI', 9))
        self.lbl_import_lh_ngay.pack(anchor='w')
        self.lbl_import_vandon = tk.Label(box1, text="Vận đơn số: -", bg='#F8FAFC', font=('Segoe UI', 9))
        self.lbl_import_vandon.pack(anchor='w')

        # Subframe 2: Kiện & Trọng lượng
        box2 = tk.LabelFrame(summary_frame, text=" Kiện Hàng & Trọng Lượng ", font=('Segoe UI', 9, 'bold'), bg='#F8FAFC', padx=10, pady=6)
        box2.pack(side='left', fill='both', expand=True, padx=6)
        self.lbl_import_kien = tk.Label(box2, text="Số lượng kiện: -", bg='#F8FAFC', font=('Segoe UI', 9))
        self.lbl_import_kien.pack(anchor='w')
        self.lbl_import_weight = tk.Label(box2, text="Trọng lượng: - KG", bg='#F8FAFC', font=('Segoe UI', 9))
        self.lbl_import_weight.pack(anchor='w')
        self.lbl_import_total_pcs = tk.Label(box2, text="Tổng số lượng: - PCE", bg='#F8FAFC', font=('Segoe UI', 9))
        self.lbl_import_total_pcs.pack(anchor='w')

        # Subframe 3: Trạng thái & Tên hàng
        box3 = tk.LabelFrame(summary_frame, text=" Tình Trạng Dữ Liệu ", font=('Segoe UI', 9, 'bold'), bg='#F8FAFC', padx=10, pady=6)
        box3.pack(side='left', fill='both', expand=True, padx=(6, 0))
        self.lbl_import_qc_status = tk.Label(box3, text="Chưa nạp tờ khai", font=('Segoe UI', 10, 'bold'), fg='#64748B', bg='#F8FAFC')
        self.lbl_import_qc_status.pack(anchor='w')
        self.lbl_import_ten_hang = tk.Label(box3, text="Tên hàng: -", bg='#F8FAFC', font=('Segoe UI', 9))
        self.lbl_import_ten_hang.pack(anchor='w')
        self.lbl_import_item_count = tk.Label(box3, text="Số dòng hàng: -", bg='#F8FAFC', font=('Segoe UI', 9), fg='#475569')
        self.lbl_import_item_count.pack(anchor='w')

        # Editable Parameters Bar
        edit_bar = tk.Frame(self.import_qc_card, bg=self.CARD_BG)
        edit_bar.pack(fill='x', pady=(4, 6))

        tk.Label(edit_bar, text="Hiệu chỉnh thông tin CV Chia Tách trước khi xuất:", font=('Segoe UI', 9, 'bold'), fg=self.PRIMARY, bg=self.CARD_BG).pack(anchor='w', pady=(0, 4))

        eb_grid = tk.Frame(edit_bar, bg=self.CARD_BG)
        eb_grid.pack(fill='x')

        tk.Label(eb_grid, text="Tên hàng hóa (CV):", font=('Segoe UI', 9), bg=self.CARD_BG).grid(row=0, column=0, sticky='w')
        self.edit_import_ten_hang = ttk.Entry(eb_grid, width=22, font=('Segoe UI', 9))
        self.edit_import_ten_hang.grid(row=0, column=1, padx=(4, 15), sticky='w')

        tk.Label(eb_grid, text="Số kiện (CV & NK):", font=('Segoe UI', 9), bg=self.CARD_BG).grid(row=0, column=2, sticky='w')
        self.edit_import_so_kien = ttk.Entry(eb_grid, width=15, font=('Segoe UI', 9))
        self.edit_import_so_kien.grid(row=0, column=3, padx=(4, 15), sticky='w')

        tk.Label(eb_grid, text="Trọng lượng KG (CV & NK):", font=('Segoe UI', 9), bg=self.CARD_BG).grid(row=0, column=4, sticky='w')
        self.edit_import_weight = ttk.Entry(eb_grid, width=18, font=('Segoe UI', 9))
        self.edit_import_weight.grid(row=0, column=5, padx=(4, 0), sticky='w')

        # 3. Bottom Card: Items Table & Actions
        bottom_import_card = tk.LabelFrame(
            tab_import,
            text=" 3. Danh Sách Hàng Hóa Chi Tiết & Tạo Tệp ",
            font=('Segoe UI', 10, 'bold'),
            fg=self.PRIMARY,
            bg=self.CARD_BG,
            padx=12,
            pady=6,
            relief='groove'
        )
        bottom_import_card.pack(fill='both', expand=True)

        # Bottom Action Bar (packed side='bottom' FIRST)
        act_bar = tk.Frame(bottom_import_card, bg=self.CARD_BG)
        act_bar.pack(side='bottom', fill='x', pady=(6, 2))

        self.lbl_import_result_status = tk.Label(act_bar, text="", font=('Segoe UI', 9, 'bold'), bg=self.CARD_BG, fg=self.SUCCESS)
        self.lbl_import_result_status.pack(side='left')

        self.open_import_cv_btn = ttk.Button(act_bar, text="Mở CV Chia Tách", style='Outline.TButton', command=self.open_import_cv, state='disabled')
        self.open_import_cv_btn.pack(side='right', padx=(4, 0), pady=2)

        self.open_import_nk_btn = ttk.Button(act_bar, text="Mở NhapKho.xlsx", style='Outline.TButton', command=self.open_import_nhapkho, state='disabled')
        self.open_import_nk_btn.pack(side='right', padx=(4, 0), pady=2)

        self.open_import_dir_btn = ttk.Button(act_bar, text="Mở Thư Mục", style='Outline.TButton', command=self.open_import_dir, state='disabled')
        self.open_import_dir_btn.pack(side='right', padx=(4, 0), pady=2)

        self.gen_import_btn = ttk.Button(
            act_bar,
            text="  ⚡  TẠO CV CHIA TÁCH & NHẬP KHO  ",
            style='Success.TButton',
            command=self.generate_import_files,
            state='disabled'
        )
        self.gen_import_btn.pack(side='right', padx=(8, 10), pady=2)

        # Table Treeview
        tree_frame = tk.Frame(bottom_import_card, bg=self.CARD_BG)
        tree_frame.pack(side='top', fill='both', expand=True, pady=(0, 4))

        cols = ("stt", "item_code", "so_luong", "don_vi", "ma_hs", "mo_ta")
        self.import_tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=6)
        self.import_tree.heading("stt", text="STT")
        self.import_tree.heading("item_code", text="Mã Hàng (Item Code)")
        self.import_tree.heading("so_luong", text="Số Lượng")
        self.import_tree.heading("don_vi", text="ĐVT")
        self.import_tree.heading("ma_hs", text="Mã HS")
        self.import_tree.heading("mo_ta", text="Mô Tả Hàng Hóa Trên Tờ Khai")

        self.import_tree.column("stt", width=50, anchor='center')
        self.import_tree.column("item_code", width=220, anchor='w')
        self.import_tree.column("so_luong", width=120, anchor='center')
        self.import_tree.column("don_vi", width=80, anchor='center')
        self.import_tree.column("ma_hs", width=110, anchor='center')
        self.import_tree.column("mo_ta", width=420, anchor='w')

        scroll_y = ttk.Scrollbar(tree_frame, orient='vertical', command=self.import_tree.yview)
        self.import_tree.configure(yscrollcommand=scroll_y.set)

        self.import_tree.pack(side='left', fill='both', expand=True)
        scroll_y.pack(side='right', fill='y')

    def detect_default_samples(self):
        samples = []
        result_dir = "ResultExample"
        if os.path.exists(result_dir):
            for entry in sorted(os.listdir(result_dir)):
                full_p = os.path.join(result_dir, entry)
                if os.path.isdir(full_p):
                    samples.append(entry)

        if samples:
            self.sample_combo['values'] = samples
            self.sample_combo.current(0)
            self.load_selected_sample()

    def on_sample_selected(self, event=None):
        self.load_selected_sample()

    def load_selected_sample(self):
        chosen = self.sample_combo.get()
        if not chosen:
            return

        sample_dir = os.path.join("ResultExample", chosen)
        if not os.path.exists(sample_dir):
            return

        do_file, bk_file, inv_file, pl_file = self.detect_files_in_dir(sample_dir)
        self.out_dir_var.set(os.path.abspath(sample_dir))
        self.do_path_var.set(os.path.abspath(do_file) if do_file else "")
        self.bk_path_var.set(os.path.abspath(bk_file) if bk_file else "")
        self.inv_path_var.set(os.path.abspath(inv_file) if inv_file else "")
        self.pl_path_var.set(os.path.abspath(pl_file) if pl_file else "")

        # Suggested date based on sample folder
        if '1645' in chosen or '1650' in chosen:
            self.date_var.set('02/11/2026')
        elif '1651' in chosen or '1652' in chosen or '1657' in chosen:
            self.date_var.set('04/11/2026')

        # Automatically trigger analyze
        self.analyze_data()

    def detect_files_in_dir(self, directory):
        """
        Detect DO, Booking/AWB, Invoice, and Packing List files in a directory.
        Returns (do_path, bk_path, inv_path, pl_path)
        """
        if not directory or not os.path.isdir(directory):
            return None, None, None, None
        try:
            files = sorted([f for f in os.listdir(directory) if not f.startswith('~') and not f.startswith('.')])
        except Exception:
            return None, None, None, None

        do_path = None
        bk_path = None
        inv_path = None
        pl_path = None

        # Look for DO excel
        for f in files:
            fl = f.lower()
            if fl.endswith(('.xlsx', '.xls')):
                if 'delivery' in fl or 'phiếu' in fl or 'phieu' in fl or 'do' in fl:
                    do_path = os.path.join(directory, f)
                    break
        if not do_path:
            for f in files:
                fl = f.lower()
                if fl.endswith(('.xlsx', '.xls')) and 'bk' not in fl and 'mẫu 25' not in fl and 'mau 25' not in fl and 'booking' not in fl and 'maxuat' not in fl:
                    do_path = os.path.join(directory, f)
                    break

        # Look for Booking / AWB
        for f in files:
            fl = f.lower()
            if fl.endswith(('.pdf', '.xlsx', '.xls')):
                if 'booking' in fl or 'awb' in fl or 'bkg' in fl:
                    bk_path = os.path.join(directory, f)
                    break

        # Look for Packing List (PL ...pdf)
        for f in files:
            fl = f.lower()
            if fl.endswith('.pdf') and (fl.startswith('pl ') or fl.startswith('pl_') or fl.startswith('pl-') or 'packing' in fl):
                pl_path = os.path.join(directory, f)
                break

        # Look for Invoice PDF (PDF that is not Booking/AWB and not Packing List)
        for f in files:
            fl = f.lower()
            if fl.endswith('.pdf'):
                if ('booking' in fl or 'awb' in fl or 'bkg' in fl or
                    fl.startswith('pl ') or fl.startswith('pl_') or fl.startswith('pl-') or 'packing' in fl):
                    continue
                inv_path = os.path.join(directory, f)
                break

        return do_path, bk_path, inv_path, pl_path

    

    def reset_inputs(self):
        """Clear all inputs, reset summary labels and treeview."""
        self.do_path_var.set("")
        self.bk_path_var.set("")
        self.inv_path_var.set("")
        self.pl_path_var.set("")
        self.out_dir_var.set("")
        self.sample_combo.set("")

        self.lbl_do_code.config(text="Mã DO: -")
        self.lbl_consignee.config(text="Người nhận: -")
        self.lbl_pkg_gw.config(text="Số kiện: - | GW: - | Tổng PCS: -")

        self.lbl_carrier.config(text="Hãng VC: - | Số BK/AWB: -")
        self.lbl_flight.config(text="Số hiệu PTVC: - | Chuyến bay/Ngày: -")
        self.lbl_wh.config(text="Cửa khẩu xuất / Kho: -")

        self.lbl_qc_status.config(text="Chưa nạp tệp", fg='#64748B')
        self.lbl_qc_dup.config(text="Trùng lặp: -", fg='#475569')
        self.lbl_qc_notes.config(text="Ghi chú: Vui lòng chọn tệp DO, Booking và Thư mục xuất", fg='#475569')

        self.edit_bk_no.delete(0, tk.END)
        self.edit_flight.delete(0, tk.END)
        self.edit_flight_date.delete(0, tk.END)
        self.edit_wh.delete(0, tk.END)

        for row in self.tree.get_children():
            self.tree.delete(row)

        self.gen_btn.config(state='disabled')
        self.open_bk_btn.config(state='disabled')
        self.open_m25_btn.config(state='disabled')
        self.open_cv_btn.config(state='disabled')
        self.open_ahk_btn.config(state='disabled')
        self.open_dir_btn.config(state='disabled')
        self.lbl_result_status.config(text="")

        self.delivery_info = None
        self.booking_info = None
        self.invoice_info = None
        self.pl_info = None
        self.val_report = None

    def browse_shipment_folder(self):
        """Prompt user to select a shipment directory; auto-detect DO, Booking, Invoice, PL, and set output dir."""
        path = filedialog.askdirectory(title="Chọn Thư Mục Lô Hàng (Chứa file DO, Booking, Invoice, PL)")
        if not path:
            return
        folder_path = os.path.abspath(path)
        self.out_dir_var.set(folder_path)
        self.sample_combo.set("")

        do_file, bk_file, inv_file, pl_file = self.detect_files_in_dir(folder_path)
        self.do_path_var.set(do_file or "")
        self.bk_path_var.set(bk_file or "")
        self.inv_path_var.set(inv_file or "")
        self.pl_path_var.set(pl_file or "")

        if do_file and bk_file:
            self.analyze_data()
        elif not do_file and not bk_file:
            messagebox.showinfo(
                "Thông Báo",
                f"Đã chọn thư mục xuất:\n{folder_path}\n\nKhông tìm thấy file DO (.xlsx) hoặc Booking (.pdf/.xlsx) trong thư mục này.\nVui lòng tự chọn file DO và Booking bằng các nút bên dưới."
            )
        elif not do_file:
            messagebox.showinfo(
                "Thông Báo",
                f"Đã tìm thấy Booking: {os.path.basename(bk_file)}\n\nChưa tìm thấy file Phiếu xuất kho (DO). Vui lòng chọn file DO."
            )
        elif not bk_file:
            messagebox.showinfo(
                "Thông Báo",
                f"Đã tìm thấy DO: {os.path.basename(do_file)}\n\nChưa tìm thấy file Booking/AWB. Vui lòng chọn file Booking."
            )

    def browse_multiple_files(self):
        """Prompt user to select multiple files at once."""
        paths = filedialog.askopenfilenames(
            title="Chọn các tệp chứng từ (DO, Booking, Invoice, Packing List)",
            filetypes=[
                ("Tất cả tệp chứng từ", "*.xlsx *.xls *.pdf"),
                ("Excel Files", "*.xlsx *.xls"),
                ("PDF Files", "*.pdf"),
                ("All Files", "*.*")
            ]
        )
        if not paths:
            return

        file_list = [os.path.abspath(p) for p in paths]
        common_dir = os.path.dirname(file_list[0])
        self.out_dir_var.set(common_dir)
        self.sample_combo.set("")

        for p in file_list:
            ext = os.path.splitext(p)[1].lower()
            base = os.path.basename(p).lower()
            if ext in ('.xlsx', '.xls'):
                if 'booking' in base or 'awb' in base:
                    self.bk_path_var.set(p)
                else:
                    self.do_path_var.set(p)
            elif ext == '.pdf':
                if 'booking' in base or 'awb' in base or 'bkg' in base:
                    self.bk_path_var.set(p)
                elif base.startswith(('pl ', 'pl_', 'pl-')) or 'packing' in base:
                    self.pl_path_var.set(p)
                else:
                    self.inv_path_var.set(p)

        do_auto, bk_auto, inv_auto, pl_auto = self.detect_files_in_dir(common_dir)
        if not self.do_path_var.get() and do_auto:
            self.do_path_var.set(do_auto)
        if not self.bk_path_var.get() and bk_auto:
            self.bk_path_var.set(bk_auto)
        if not self.inv_path_var.get() and inv_auto:
            self.inv_path_var.set(inv_auto)
        if not self.pl_path_var.get() and pl_auto:
            self.pl_path_var.set(pl_auto)

        if self.do_path_var.get() and self.bk_path_var.get():
            self.analyze_data()

    def browse_inv_file(self):
        path = filedialog.askopenfilename(
            title="Chọn Tệp Invoice (PDF)",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        if path:
            self.inv_path_var.set(os.path.abspath(path))
            self.sample_combo.set("")

    def browse_pl_file(self):
        path = filedialog.askopenfilename(
            title="Chọn Tệp Packing List (PDF)",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        if path:
            self.pl_path_var.set(os.path.abspath(path))
            self.sample_combo.set("")

    

    def browse_do_file(self):
        path = filedialog.askopenfilename(
            title="Chọn Phiếu Xuất Kho (Delivery Order)",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if path:
            abs_p = os.path.abspath(path)
            self.do_path_var.set(abs_p)
            self.sample_combo.set("")
            parent_dir = os.path.dirname(abs_p)

            # Auto-update output dir if not set or was pointing to ResultExample
            current_out = self.out_dir_var.get().strip()
            if not current_out or 'resultexample' in current_out.lower():
                self.out_dir_var.set(parent_dir)

            # If booking/inv/pl are not set or were from ResultExample, try to auto-detect in the same directory
            _, detected_bk, detected_inv, detected_pl = self.detect_files_in_dir(parent_dir)
            current_bk = self.bk_path_var.get().strip()
            if (not current_bk or 'resultexample' in current_bk.lower()) and detected_bk:
                self.bk_path_var.set(detected_bk)
            current_inv = self.inv_path_var.get().strip()
            if (not current_inv or 'resultexample' in current_inv.lower()) and detected_inv:
                self.inv_path_var.set(detected_inv)
            current_pl = self.pl_path_var.get().strip()
            if (not current_pl or 'resultexample' in current_pl.lower()) and detected_pl:
                self.pl_path_var.set(detected_pl)

            if self.do_path_var.get() and self.bk_path_var.get() and os.path.exists(self.bk_path_var.get()):
                self.analyze_data()

    def browse_bk_file(self):
        path = filedialog.askopenfilename(
            title="Chọn Booking Confirmation hoặc AWB",
            filetypes=[
                ("Tất cả tệp hợp lệ", "*.pdf *.xlsx *.xls"),
                ("PDF Files", "*.pdf"),
                ("Excel Files", "*.xlsx *.xls"),
                ("All Files", "*.*")
            ]
        )
        if path:
            abs_p = os.path.abspath(path)
            self.bk_path_var.set(abs_p)
            self.sample_combo.set("")
            parent_dir = os.path.dirname(abs_p)

            # Auto-update output dir if not set or was pointing to ResultExample
            current_out = self.out_dir_var.get().strip()
            if not current_out or 'resultexample' in current_out.lower():
                self.out_dir_var.set(parent_dir)

            # If DO/inv/pl are not set or were from ResultExample, try to auto-detect in the same directory
            detected_do, _, detected_inv, detected_pl = self.detect_files_in_dir(parent_dir)
            current_do = self.do_path_var.get().strip()
            if (not current_do or 'resultexample' in current_do.lower()) and detected_do:
                self.do_path_var.set(detected_do)
            current_inv = self.inv_path_var.get().strip()
            if (not current_inv or 'resultexample' in current_inv.lower()) and detected_inv:
                self.inv_path_var.set(detected_inv)
            current_pl = self.pl_path_var.get().strip()
            if (not current_pl or 'resultexample' in current_pl.lower()) and detected_pl:
                self.pl_path_var.set(detected_pl)

            if self.do_path_var.get() and os.path.exists(self.do_path_var.get()) and self.bk_path_var.get():
                self.analyze_data()

    def browse_out_dir(self):
        path = filedialog.askdirectory(title="Chọn Thư Mục Lưu Tệp Xuất (BK, Mẫu 25, CV)")
        if path:
            self.out_dir_var.set(os.path.abspath(path))

    def open_calendar(self):
        CalendarDialog(self, initial_date=self.date_var.get(), callback=self.on_date_selected)

    def on_date_selected(self, selected_date):
        self.date_var.set(selected_date.strftime('%d/%m/%Y'))

    def set_today(self):
        self.date_var.set(datetime.now().strftime('%d/%m/%Y'))

    def analyze_data(self):
        do_path = self.do_path_var.get().strip()
        bk_path = self.bk_path_var.get().strip()
        inv_path = self.inv_path_var.get().strip()
        pl_path = self.pl_path_var.get().strip()

        if not do_path or not os.path.exists(do_path):
            messagebox.showerror("Thiếu tệp", "Vui lòng chọn tệp Delivery Order (.xlsx) hợp lệ!")
            return

        if not bk_path or not os.path.exists(bk_path):
            messagebox.showerror("Thiếu tệp", "Vui lòng chọn tệp Booking Confirmation / AWB (.pdf hoặc .xlsx) hợp lệ!")
            return

        try:
            self.delivery_info = parse_delivery_order(do_path)
            self.booking_info = parse_booking_awb(bk_path)
            self.val_report = validate_shipment_data(self.delivery_info, self.booking_info)
            self.invoice_info = parse_invoice_pdf(inv_path) if (inv_path and os.path.exists(inv_path)) else None
            self.pl_info = parse_packing_list_pdf(pl_path) if (pl_path and os.path.exists(pl_path)) else None
        except Exception as e:
            messagebox.showerror("Lỗi đọc tệp", f"Đã xảy ra lỗi khi phân tích dữ liệu:\n{str(e)}")
            return

        # Update DO summary UI
        self.lbl_do_code.config(text=f"Mã DO: {self.delivery_info['do_code'] or 'Không rõ'}")
        self.lbl_consignee.config(text=f"Người nhận: {self.delivery_info['consignee'][:38]}")
        self.lbl_pkg_gw.config(
            text=f"Số kiện: {self.delivery_info['package_str']} | GW: {self.delivery_info['total_gw']} kg | PCS: {self.delivery_info['total_pcs']}"
        )

        # Update Booking summary UI
        self.lbl_carrier.config(text=f"Hãng VC: {self.booking_info['carrier']} | Số BK: {self.booking_info['bk_no']}")
        self.lbl_flight.config(text=f"Số PTVC: {self.booking_info['flight']} | PTVC/Ngày: {self.booking_info['flight_date']}")
        self.lbl_wh.config(text=f"Cửa khẩu xuất / Kho: {self.booking_info['warehouse']}")

        # Pre-fill editable inputs
        self.edit_bk_no.delete(0, tk.END)
        self.edit_bk_no.insert(0, self.booking_info.get('bk_no', ''))

        self.edit_flight.delete(0, tk.END)
        self.edit_flight.insert(0, self.booking_info.get('flight', ''))

        self.edit_flight_date.delete(0, tk.END)
        self.edit_flight_date.insert(0, self.booking_info.get('flight_date', ''))

        self.edit_wh.delete(0, tk.END)
        self.edit_wh.insert(0, self.booking_info.get('warehouse', ''))

        # Update Status Badge
        status = self.val_report['status']
        if status == 'OK':
            self.lbl_qc_status.config(text="✓ DỮ LIỆU ĐẦY ĐỦ & HỢP LỆ", fg=self.SUCCESS)
        elif status == 'WARNING':
            self.lbl_qc_status.config(text="⚠ CÓ CẢNH BÁO (HỢP LỆ ĐỂ TẠO)", fg=self.WARNING)
        else:
            self.lbl_qc_status.config(text="✕ THIẾU THÔNG TIN BẮT BUỘC", fg=self.DANGER)

        dup_count = self.delivery_info['duplicates_count']
        if dup_count > 0:
            self.lbl_qc_dup.config(
                text=f"Trùng lặp: Phát hiện {dup_count} dòng trùng TK nhập (Đã gộp & cộng dồn)",
                fg=self.WARNING
            )
        else:
            self.lbl_qc_dup.config(text="Trùng lặp: Không có dòng trùng lặp", fg='#475569')

        if self.val_report['warnings']:
            self.lbl_qc_notes.config(text="Ghi chú: " + self.val_report['warnings'][0][:55] + "...", fg='#64748B')
        elif self.val_report['issues']:
            self.lbl_qc_notes.config(text="Lỗi: " + self.val_report['issues'][0], fg=self.DANGER)
        else:
            self.lbl_qc_notes.config(text="Ghi chú: Tất cả trường dữ liệu sẵn sàng để xuất.", fg=self.SUCCESS)

        # Populate Items Table
        for row in self.tree.get_children():
            self.tree.delete(row)

        for idx, item in enumerate(self.delivery_info['grouped_items'], 1):
            self.tree.insert(
                "",
                "end",
                values=(
                    idx,
                    item['tk_nhap'],
                    item['short_name'],
                    item['quantity'],
                    item['unit'],
                    item['description'][:65]
                )
            )

        # Enable generation
        self.gen_btn.config(state='normal')
        self.lbl_result_status.config(text="Dữ liệu đã sẵn sàng! Nhấn 'Tạo BK.xlsx & Mẫu 25 NQ' để xuất.", fg='#0284C7')

    def generate_files(self):
        if not self.delivery_info or not self.booking_info:
            messagebox.showwarning("Chưa phân tích", "Vui lòng nhấn Phân Tích Dữ Liệu trước!")
            return

        out_dir = self.out_dir_var.get().strip()
        if not out_dir:
            out_dir = os.path.dirname(self.do_path_var.get())
        os.makedirs(out_dir, exist_ok=True)

        chosen_date = self.date_var.get().strip()
        if not chosen_date:
            chosen_date = datetime.now().strftime('%d/%m/%Y')

        # Read any user manual overrides from edit bar
        override_bk = self.edit_bk_no.get().strip()
        override_flight = self.edit_flight.get().strip()
        override_flight_date = self.edit_flight_date.get().strip()
        override_wh = self.edit_wh.get().strip()

        if override_bk:
            self.booking_info['bk_no'] = override_bk
        if override_flight:
            self.booking_info['flight'] = override_flight
        if override_flight_date:
            self.booking_info['flight_date'] = override_flight_date
        if override_wh:
            self.booking_info['warehouse'] = override_wh

        bk_out_path = os.path.join(out_dir, "BK.xlsx")
        m25_out_path = os.path.join(out_dir, "mẫu 25 NQ.xlsx")
        ma_xuat_out_path = os.path.join(out_dir, "MaXuat.xlsx")
        cv_out_path = os.path.join(out_dir, "CV ĐÓNG GHÉP 2026BT-VHD.docx")
        ahk_out_path = os.path.join(out_dir, "fillScript.ahk")

        try:
            generate_bk_excel(self.delivery_info, self.booking_info, chosen_date, bk_out_path)
            generate_mau25_excel(self.delivery_info, self.booking_info, chosen_date, m25_out_path)
            generate_ma_xuat_excel(self.delivery_info, chosen_date, ma_xuat_out_path)
            cv_res = generate_cv_dong_ghep_docx(self.delivery_info, cv_out_path)

            ahk_res = None
            if self.invoice_info or self.pl_info:
                ahk_res = generate_fill_script_ahk(
                    self.delivery_info,
                    self.booking_info,
                    self.invoice_info,
                    self.pl_info,
                    chosen_date,
                    ahk_out_path
                )

            self.last_bk_file = bk_out_path
            self.last_m25_file = m25_out_path
            self.last_cv_file = cv_res
            self.last_ahk_file = ahk_res

            self.open_bk_btn.config(state='normal')
            self.open_m25_btn.config(state='normal')
            self.open_dir_btn.config(state='normal')
            self.open_cv_btn.config(state='normal' if cv_res else 'disabled')
            self.open_ahk_btn.config(state='normal' if ahk_res else 'disabled')

            created_files = [
                os.path.basename(bk_out_path),
                os.path.basename(m25_out_path),
                os.path.basename(ma_xuat_out_path),
            ]
            if cv_res:
                created_files.append(os.path.basename(cv_out_path))
            if ahk_res:
                created_files.append(os.path.basename(ahk_out_path))

            files_list_str = "\n".join(f"{i}. {fn}" for i, fn in enumerate(created_files, 1))
            extra_note = ""
            if not cv_res:
                extra_note += "\n(Lô hàng chỉ có 1 số TK nhập, không cần tạo CV đóng ghép)"
            if not ahk_res:
                extra_note += "\n(Chưa chọn Invoice / Packing List nên bỏ qua fillScript.ahk)"

            msg = f"ĐÃ TẠO THÀNH CÔNG {len(created_files)} TỆP TẠI:\n{out_dir}\n\n{files_list_str}{extra_note}"
            self.lbl_result_status.config(
                text=f"✓ Đã tạo thành công {len(created_files)} tệp!",
                fg=self.SUCCESS
            )

            messagebox.showinfo("Thành Công", msg)
        except Exception as e:
            messagebox.showerror("Lỗi Tạo Tệp", f"Đã xảy ra lỗi khi tạo tệp:\n{str(e)}")

    def open_generated_ahk(self):
        if self.last_ahk_file and os.path.exists(self.last_ahk_file):
            os.startfile(self.last_ahk_file)

    def open_generated_cv(self):
        if self.last_cv_file and os.path.exists(self.last_cv_file):
            os.startfile(self.last_cv_file)

    def open_generated_bk(self):
        if self.last_bk_file and os.path.exists(self.last_bk_file):
            os.startfile(self.last_bk_file)

    def open_generated_m25(self):
        if self.last_m25_file and os.path.exists(self.last_m25_file):
            os.startfile(self.last_m25_file)

    def open_output_dir(self):
        out_dir = self.out_dir_var.get().strip()
        if out_dir and os.path.exists(out_dir):
            os.startfile(out_dir)

    # ============================================================
    # TAB 2: NHẬP KHẨU (CUSTOMS DECLARATION & INVENTORY)
    # ============================================================
    def detect_default_import_samples(self):
        samples = []
        result_dir = "ResultExampleImport"
        if os.path.exists(result_dir):
            for entry in sorted(os.listdir(result_dir)):
                full_p = os.path.join(result_dir, entry)
                if os.path.isdir(full_p):
                    samples.append(entry)

        if samples:
            self.import_sample_combo['values'] = samples
            self.import_sample_combo.current(0)
            self.load_selected_import_sample()

    def on_import_sample_selected(self, event=None):
        self.load_selected_import_sample()

    def load_selected_import_sample(self):
        chosen = self.import_sample_combo.get()
        if not chosen:
            return

        sample_dir = os.path.join("ResultExampleImport", chosen)
        if not os.path.exists(sample_dir):
            return

        tokhai_file = self.detect_import_files_in_dir(sample_dir)
        self.import_out_dir_var.set(os.path.abspath(sample_dir))
        self.tokhai_path_var.set(os.path.abspath(tokhai_file) if tokhai_file else "")

        # Auto-extract PLHĐ number from folder name
        m = re.search(r'IMP26-(\d+)', chosen)
        if m:
            self.plhd_var.set(m.group(1))

        # Default date
        self.import_date_var.set("12/11/2026")

        self.analyze_import_data()

    def detect_import_files_in_dir(self, directory):
        """Find customs declaration Excel file in directory."""
        if not directory or not os.path.isdir(directory):
            return None
        try:
            files = sorted([f for f in os.listdir(directory) if not f.startswith('~') and not f.startswith('.')])
        except Exception:
            return None

        # Priority 1: Contains 'tokhai' or 'to khai'
        for f in files:
            fl = f.lower()
            if fl.endswith(('.xls', '.xlsx')) and ('tokhai' in fl or 'tờ khai' in fl or 'to khai' in fl):
                return os.path.join(directory, f)

        # Priority 2: Any .xls file not NhapKho or BK
        for f in files:
            fl = f.lower()
            if fl.endswith(('.xls', '.xlsx')) and not fl.startswith('nhapkho') and not fl.startswith('bk'):
                return os.path.join(directory, f)

        return None

    def browse_import_folder(self):
        folder = filedialog.askdirectory(title="Chọn Thư Mục Lô Hàng Nhập Khẩu")
        if not folder:
            return
        folder_path = os.path.abspath(folder)
        self.import_out_dir_var.set(folder_path)
        self.import_sample_combo.set("")

        tokhai_file = self.detect_import_files_in_dir(folder_path)
        self.tokhai_path_var.set(tokhai_file or "")

        m = re.search(r'IMP26-(\d+)', os.path.basename(folder_path))
        if m:
            self.plhd_var.set(m.group(1))

        if tokhai_file:
            self.analyze_import_data()
        else:
            messagebox.showinfo(
                "Thông Báo",
                f"Đã chọn thư mục:\n{folder_path}\n\nKhông tìm thấy file Tờ khai HQ (.xls/.xlsx).\nVui lòng tự chọn file Tờ khai bằng nút 'Chọn Tệp Tờ Khai'."
            )

    def browse_tokhai_file(self):
        path = filedialog.askopenfilename(
            title="Chọn Tệp Tờ Khai Hải Quan (.xls / .xlsx)",
            filetypes=[("Excel Files", "*.xls *.xlsx"), ("All Files", "*.*")]
        )
        if path:
            abs_p = os.path.abspath(path)
            self.tokhai_path_var.set(abs_p)
            self.import_sample_combo.set("")
            parent_dir = os.path.dirname(abs_p)

            current_out = self.import_out_dir_var.get().strip()
            if not current_out or 'resultexampleimport' in current_out.lower():
                self.import_out_dir_var.set(parent_dir)

            m = re.search(r'IMP26-(\d+)', os.path.basename(parent_dir))
            if m:
                self.plhd_var.set(m.group(1))

            self.analyze_import_data()

    def browse_import_out_dir(self):
        path = filedialog.askdirectory(title="Chọn Thư Mục Lưu Tệp Xuất Nhập Khẩu")
        if path:
            self.import_out_dir_var.set(os.path.abspath(path))

    def open_import_calendar(self):
        CalendarDialog(self, initial_date=self.import_date_var.get(), callback=self.on_import_date_selected)

    def on_import_date_selected(self, selected_date):
        self.import_date_var.set(selected_date.strftime('%d/%m/%Y'))

    def set_import_today(self):
        self.import_date_var.set(datetime.now().strftime('%d/%m/%Y'))

    def reset_import_inputs(self):
        self.tokhai_path_var.set("")
        self.plhd_var.set("")
        self.import_out_dir_var.set("")
        self.import_sample_combo.set("")

        self.lbl_import_tk.config(text="Số tờ khai: -")
        self.lbl_import_lh_ngay.config(text="Loại hình: - | Ngày ĐK: -")
        self.lbl_import_vandon.config(text="Vận đơn số: -")
        self.lbl_import_kien.config(text="Số lượng kiện: -")
        self.lbl_import_weight.config(text="Trọng lượng: - KG")
        self.lbl_import_total_pcs.config(text="Tổng số lượng: - PCE")
        self.lbl_import_qc_status.config(text="Chưa nạp tờ khai", fg='#64748B')
        self.lbl_import_ten_hang.config(text="Tên hàng: -")
        self.lbl_import_item_count.config(text="Số dòng hàng: -", fg='#475569')

        self.edit_import_ten_hang.delete(0, tk.END)
        self.edit_import_so_kien.delete(0, tk.END)
        self.edit_import_weight.delete(0, tk.END)

        for row in self.import_tree.get_children():
            self.import_tree.delete(row)

        self.gen_import_btn.config(state='disabled')
        self.open_import_cv_btn.config(state='disabled')
        self.open_import_nk_btn.config(state='disabled')
        self.open_import_dir_btn.config(state='disabled')
        self.lbl_import_result_status.config(text="")

        self.tokhai_info = None

    def analyze_import_data(self):
        tokhai_path = self.tokhai_path_var.get().strip()
        if not tokhai_path or not os.path.exists(tokhai_path):
            messagebox.showerror("Thiếu tệp", "Vui lòng chọn tệp Tờ Khai Hải Quan (.xls) hợp lệ!")
            return

        try:
            self.tokhai_info = parse_tokhai_hq_xls(tokhai_path)
        except Exception as e:
            messagebox.showerror("Lỗi đọc tệp", f"Đã xảy ra lỗi khi phân tích tờ khai hải quan:\n{str(e)}")
            return

        # Update UI labels
        self.lbl_import_tk.config(text=f"Số tờ khai: {self.tokhai_info['so_tk']}")
        self.lbl_import_lh_ngay.config(text=f"Loại hình: {self.tokhai_info['loai_hinh']} | Ngày ĐK: {self.tokhai_info['ngay_dk']}")
        self.lbl_import_vandon.config(text=f"Vận đơn số: {self.tokhai_info['so_van_don']}")

        self.lbl_import_kien.config(text=f"Số lượng kiện: {self.tokhai_info['so_luong_kien']}")
        self.lbl_import_weight.config(text=f"Trọng lượng: {self.tokhai_info['trong_luong_str']} KG")
        self.lbl_import_total_pcs.config(text=f"Tổng số lượng: {self.tokhai_info['total_pcs']:,} PCE".replace(',', '.'))

        self.lbl_import_qc_status.config(text="✓ DỮ LIỆU TỜ KHAI HỢP LỆ", fg=self.SUCCESS)
        self.lbl_import_ten_hang.config(text=f"Tên hàng: {self.tokhai_info['ten_hang']}")
        self.lbl_import_item_count.config(text=f"Số dòng hàng: {len(self.tokhai_info['items'])} items", fg=self.SUCCESS)

        # Pre-fill editable inputs
        self.edit_import_ten_hang.delete(0, tk.END)
        self.edit_import_ten_hang.insert(0, self.tokhai_info.get('ten_hang', ''))

        self.edit_import_so_kien.delete(0, tk.END)
        self.edit_import_so_kien.insert(0, self.tokhai_info.get('so_luong_kien', ''))

        self.edit_import_weight.delete(0, tk.END)
        self.edit_import_weight.insert(0, str(self.tokhai_info.get('trong_luong_str', '')))

        # Populate Items Table
        for row in self.import_tree.get_children():
            self.import_tree.delete(row)

        for item in self.tokhai_info['items']:
            self.import_tree.insert(
                "",
                "end",
                values=(
                    item['stt'],
                    item['item_code'],
                    f"{item['quantity']:,}".replace(',', '.'),
                    item['unit'],
                    item['hs_code'],
                    item['description'][:65]
                )
            )

        self.gen_import_btn.config(state='normal')
        self.lbl_import_result_status.config(text="Dữ liệu đã sẵn sàng! Nhấn 'Tạo CV Chia Tách & Nhập Kho' để xuất.", fg='#0284C7')

    def generate_import_files(self):
        if not self.tokhai_info:
            messagebox.showwarning("Chưa phân tích", "Vui lòng nhấn Phân Tích Tờ Khai trước!")
            return

        out_dir = self.import_out_dir_var.get().strip()
        if not out_dir:
            out_dir = os.path.dirname(self.tokhai_path_var.get())
        os.makedirs(out_dir, exist_ok=True)

        chosen_date = self.import_date_var.get().strip()
        if not chosen_date:
            chosen_date = datetime.now().strftime('%d/%m/%Y')

        plhd_num = self.plhd_var.get().strip()
        if not plhd_num:
            plhd_num = "100"

        # Apply overrides
        override_ten_hang = self.edit_import_ten_hang.get().strip()
        override_so_kien = self.edit_import_so_kien.get().strip()
        override_weight = self.edit_import_weight.get().strip()

        if override_ten_hang:
            self.tokhai_info['ten_hang'] = override_ten_hang
        if override_so_kien:
            self.tokhai_info['so_luong_kien'] = override_so_kien
            m_pkg = re.match(r'(\d+)\s*([A-Za-z]+)?', override_so_kien)
            if m_pkg:
                self.tokhai_info['pkg_count'] = int(m_pkg.group(1))
                if m_pkg.group(2):
                    self.tokhai_info['pkg_unit'] = m_pkg.group(2).upper()
        if override_weight:
            self.tokhai_info['trong_luong_str'] = override_weight
            clean_w = override_weight.replace('.', '').replace(',', '.')
            try:
                self.tokhai_info['trong_luong_kg'] = float(clean_w)
            except ValueError:
                pass

        cv_out_path = os.path.join(out_dir, "CV CHIA TÁCH 2026BT-VHD.docx")
        nk_out_path = os.path.join(out_dir, "NhapKho.xlsx")

        try:
            generate_cv_chia_tach_docx(self.tokhai_info, chosen_date, cv_out_path)
            generate_nhapkho_excel(self.tokhai_info, plhd_num, chosen_date, nk_out_path)

            self.last_import_cv_file = cv_out_path
            self.last_import_nk_file = nk_out_path

            self.open_import_cv_btn.config(state='normal')
            self.open_import_nk_btn.config(state='normal')
            self.open_import_dir_btn.config(state='normal')

            created_files = [
                os.path.basename(cv_out_path),
                os.path.basename(nk_out_path)
            ]
            files_list_str = "\n".join(f"{i}. {fn}" for i, fn in enumerate(created_files, 1))
            msg = f"ĐÃ TẠO THÀNH CÔNG {len(created_files)} TỆP TẠI:\n{out_dir}\n\n{files_list_str}"
            self.lbl_import_result_status.config(
                text=f"✓ Đã tạo thành công {len(created_files)} tệp!",
                fg=self.SUCCESS
            )
            messagebox.showinfo("Thành Công", msg)
        except Exception as e:
            messagebox.showerror("Lỗi Tạo Tệp", f"Đã xảy ra lỗi khi tạo tệp:\n{str(e)}")

    def open_import_cv(self):
        if self.last_import_cv_file and os.path.exists(self.last_import_cv_file):
            os.startfile(self.last_import_cv_file)

    def open_import_nhapkho(self):
        if self.last_import_nk_file and os.path.exists(self.last_import_nk_file):
            os.startfile(self.last_import_nk_file)

    def open_import_dir(self):
        out_dir = self.import_out_dir_var.get().strip()
        if out_dir and os.path.exists(out_dir):
            os.startfile(out_dir)


if __name__ == '__main__':
    app = AppGUI()
    app.mainloop()

