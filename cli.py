import argparse
import os
import sys
from datetime import datetime

# Configure UTF-8 for console output
try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

from core.extractor import parse_delivery_order, parse_booking_awb
from core.validator import validate_shipment_data
from core.generator import generate_bk_excel, generate_mau25_excel, generate_cv_dong_ghep_docx


def process_shipment(do_path, bk_path, date_str, output_dir, template_path=None, bk_template_path=None):
    if not os.path.exists(do_path):
        raise FileNotFoundError(f"Delivery order file not found: {do_path}")
    if not os.path.exists(bk_path):
        raise FileNotFoundError(f"Booking/AWB file not found: {bk_path}")

    os.makedirs(output_dir, exist_ok=True)

    print(f"\n==========================================")
    print(f"Đang xử lý lô hàng:")
    print(f"  DO File     : {os.path.basename(do_path)}")
    print(f"  Booking/AWB : {os.path.basename(bk_path)}")
    print(f"  Ngày        : {date_str}")
    print(f"  Thư mục xuất: {output_dir}")
    print(f"==========================================")

    # 1. Parse DO
    do_info = parse_delivery_order(do_path)
    print(f"-> DO Code: {do_info['do_code']}, Consignee: {do_info['consignee']}")
    print(f"-> Số kiện: {do_info['package_str']}, GW: {do_info['total_gw']} kg, PCS: {do_info['total_pcs']}")
    print(f"-> Số dòng hàng sau khi gộp: {len(do_info['grouped_items'])} (Đã gộp {do_info['duplicates_count']} dòng trùng)")

    # 2. Parse Booking/AWB
    booking_info = parse_booking_awb(bk_path)
    print(f"-> Hãng vận chuyển: {booking_info['carrier']}, Số BK: {booking_info['bk_no']}")
    print(f"-> Số PTVC: {booking_info['flight']}, Chuyến bay/Ngày: {booking_info['flight_date']}, Kho/Terminal: {booking_info['warehouse']}")

    # 3. Validate
    val = validate_shipment_data(do_info, booking_info)
    print(f"-> Tình trạng kiểm tra chất lượng: {val['status']}")
    for w in val['warnings']:
        print(f"   [CẢNH BÁO] {w}")
    for iss in val['issues']:
        print(f"   [LỖI] {iss}")

    # 4. Generate
    bk_out = os.path.join(output_dir, "BK.xlsx")
    m25_out = os.path.join(output_dir, "mẫu 25 NQ.xlsx")
    cv_out = os.path.join(output_dir, "CV ĐÓNG GHÉP 2026BT-VHD.docx")

    generate_bk_excel(do_info, booking_info, date_str, bk_out, base_template_path=bk_template_path)
    generate_mau25_excel(do_info, booking_info, date_str, m25_out, base_template_path=template_path)
    cv_res = generate_cv_dong_ghep_docx(do_info, cv_out)

    print(f"-> Xuất thành công: {bk_out}")
    print(f"-> Xuất thành công: {m25_out}")
    if cv_res:
        print(f"-> Xuất thành công: {cv_out} (Lô hàng có từ 2 số TK nhập khác nhau)")
    return bk_out, m25_out, cv_res


def main():
    parser = argparse.ArgumentParser(description="Tự động trích xuất chứng từ & tạo BK.xlsx, Mẫu 25 NQ.xlsx")
    parser.add_argument("--do", required=True, help="Đường dẫn tệp Delivery Order (.xlsx)")
    parser.add_argument("--bk", required=True, help="Đường dẫn tệp Booking Confirmation hoặc AWB (.pdf / .xlsx)")
    parser.add_argument("--date", default=datetime.now().strftime("%d/%m/%Y"), help="Ngày kê khai (DD/MM/YYYY)")
    parser.add_argument("--out", default="output", help="Thư mục xuất kết quả")
    parser.add_argument("--template", default=None, help="Đường dẫn mẫu 25 NQ template tùy chỉnh")
    parser.add_argument("--bk-template", default=None, help="Đường dẫn BK template tùy chỉnh")

    args = parser.parse_args()
    process_shipment(args.do, args.bk, args.date, args.out, args.template, args.bk_template)


if __name__ == '__main__':
    main()
