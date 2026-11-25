# Hệ Thống Tự Động Trích Xuất Chứng Từ & Tạo BK / Mẫu 25 NQ

Ứng dụng Python chuyên dụng để tự động hóa quy trình xử lý chứng từ xuất nhập khẩu:
1. **Đọc thông tin từ tệp Excel hiện có**: Đọc và trích xuất dữ liệu từ các tệp **Phiếu xuất kho (Delivery Order - DO)**.
2. **Đọc bảng biểu và văn bản từ tệp PDF/Excel**: Đọc tệp **Booking Confirmation (Xác nhận đặt chỗ hàng không)** hoặc **AWB (Vận đơn chuyển phát nhanh)** định dạng PDF hoặc Excel.
3. **Trích xuất thông tin đặc thù**: Tự động nhận diện Hãng vận chuyển (Air Expeditors, UPS Air, FedEx, DHL, UPS Courier), Số BK/AWB/MAWB, Số hiệu phương tiện vận chuyển (Số PTVC / Chuyến bay), Cửa khẩu xuất / Kho hàng (ACSV, NCTS, FEDEX, DHL, UPS), Mã DO, Người nhận (Consignee), Trọng lượng (GW), Số kiện (Packages).
4. **Tổ chức và liên kết thông tin**: Tự động gộp (aggregate) các dòng hàng có cùng **Số tờ khai nhập (TK nhập / IMPORT CD)**, cộng dồn số lượng PCS, và tự động gán tên hàng tiếng Việt tiêu chuẩn hải quan (`Ăng ten`, `Cáp đồng trục`, `Pallet gỗ`, `Biến áp`, `Đầu nối`, `Bộ lọc`).
5. **Kiểm tra thông tin thiếu & Trùng lặp (Quality Control & Deduplication)**: Phát hiện và cảnh báo các dòng trùng lặp đã được gộp, kiểm tra trường dữ liệu bị thiếu hoặc lệch số kiện/trọng lượng giữa DO và Booking/AWB.
6. **Tự động tạo các tệp Excel mới**: Tạo tự động 2 tệp chứng từ chính xác 100% theo mẫu thực tế:
   - **`BK.xlsx`**: Bảng kê tờ khai vận chuyển độc lập (phông chữ Times New Roman 13pt, căn lề, định dạng ô gộp, chữ ký doanh nghiệp và hải quan).
   - **`mẫu 25 NQ.xlsx`**: Danh mục hàng hóa xuất khẩu sheet `2026` (cập nhật ngày kê khai, số BK trên phần tiêu đề, tự động thêm các dòng hàng hóa và dòng tổng cộng ở cuối bảng).

---

## 📁 Cấu Trúc Dự Án

```
PythonApp/
├── core/
│   ├── extractor.py        # Module đọc & trích xuất Delivery Order (Excel) và Booking/AWB (PDF/Excel)
│   ├── validator.py        # Module kiểm tra dữ liệu, phát hiện trùng lặp & thiếu thông tin
│   └── generator.py        # Module tạo tệp BK.xlsx & Mẫu 25 NQ.xlsx theo mẫu chuẩn
├── templates/
│   └── mau_25_NQ_template.xlsx  # Mẫu gốc chuẩn của Mẫu 25 NQ (chứa dữ liệu lịch sử đến dòng 1189)
├── ResultExample/          # Thư mục chứa 5 bộ mẫu thực tế kiểm thử
│   ├── EXP26-1645-AIR-SO #34803/
│   ├── EXP26-1650-AIR-SO #36592+37616+.../
│   ├── EXP26-1651-FDX-SO #37392/
│   ├── EXP26-1652-DHL-SO #37238/
│   └── EXP26-1657-UPS-SO #36396+38371/
├── gui.py                  # Giao diện đồ họa người dùng (Tkinter GUI)
├── main.py                 # Tệp khởi chạy chính của ứng dụng
├── cli.py                  # Công cụ dòng lệnh (CLI) chạy tự động
├── test_all_samples.py     # Bộ kiểm thử tự động so khớp 100% với cả 5 thư mục mẫu
└── requirements.txt        # Danh sách thư viện Python cần thiết
```

---

## 🚀 Hướng Dẫn Cài Đặt & Sử Dụng

### 1. Cài đặt thư viện
```bash
pip install -r requirements.txt
```

### 2. Khởi chạy Giao diện Người dùng (GUI)
Chạy lệnh sau trong terminal hoặc PowerShell:
```bash
python main.py
```
*(hoặc `python gui.py`)*

#### Tính năng trên giao diện:
- **Chọn tệp**: 
  - Chọn tệp Delivery Order (`.xlsx`)
  - Chọn tệp Booking Confirmation hoặc AWB (`.pdf` hoặc `.xlsx`)
  - Hỗ trợ menu xổ xuống **"Mẫu ví dụ có sẵn"** để thử nghiệm ngay lập tức các mẫu trong `ResultExample`!
- **Chọn ngày**:
  - Nhập trực tiếp ngày dạng `DD/MM/YYYY`
  - Nhấn nút **"📅 Lịch"** để mở lịch trực quan chọn ngày/tháng/năm
  - Nhấn nút **"Hôm nay"** để lấy ngày hiện tại.
- **Thư mục xuất**: Chọn thư mục lưu tệp kết quả (mặc định cùng thư mục với phiếu DO).
- **Phân tích & Kiểm tra dữ liệu**:
  - Nhấn **"Phân Tích Dữ Liệu & Kiểm Tra Trùng Lặp"** để ứng dụng tự động bóc tách dữ liệu.
  - Hiển thị bảng tổng hợp thông tin lô hàng, trạng thái hợp lệ, thông báo số dòng trùng lặp đã gộp.
  - Bảng danh sách hàng hóa cho phép xem trước STT, Số TK nhập, Tên hàng, Số lượng PCS, Đơn vị.
  - Cho phép người dùng chỉnh sửa nhanh Số BK, Số hiệu chuyến bay (PTVC), Cửa khẩu xuất ngay trên giao diện trước khi xuất tệp.
- **Tạo tệp**:
  - Nhấn **"TẠO BK.XLSX & MẪU 25 NQ.XLSX"** để tạo cả 2 tệp.
  - Sau khi tạo xong, có sẵn các nút bấm 1-click: **"Mở BK.xlsx"**, **"Mở Mẫu 25 NQ"**, và **"Mở Thư Mục"**.

---

### 3. Sử dụng qua Dòng Lệnh (CLI)
Dành cho tự động hóa hoặc tích hợp script:
```bash
python cli.py --do "duong_dan_den_delivery_order.xlsx" --bk "duong_dan_den_awb_hoac_booking.pdf" --date "04/11/2026" --out "thu_muc_xuat"
```

---

## 🧪 Kết Quả Kiểm Thử (So Sánh 100% Với Mẫu Thực Tế)

Hệ thống đã được kiểm thử tự động với toàn bộ 5 thư mục trong `ResultExample`:
```bash
python test_all_samples.py
```

| Thư mục mẫu | Loại hình / Hãng | Tệp Booking/AWB | Kiểm tra BK.xlsx | Kiểm tra Mẫu 25 NQ |
| :--- | :--- | :--- | :---: | :---: |
| **EXP26-1645** | AIR (Expeditors) | `Booking confirmation Noi Bai ...xlsx` | ✅ Khớp 100% | ✅ Khớp 100% |
| **EXP26-1650** | AIR (UPS Air) | `Air booking confirmation...pdf` | ✅ Khớp 100% | ✅ Khớp 100% |
| **EXP26-1651** | Chuyển phát nhanh FedEx | `AWB 37392-34238...pdf` | ✅ Khớp 100% | ✅ Khớp 100% |
| **EXP26-1652** | Chuyển phát nhanh DHL | `AWB 37238-34239...pdf` | ✅ Khớp 100% | ✅ Khớp 100% |
| **EXP26-1657** | Chuyển phát nhanh UPS | `AWB 36396-34246...pdf` (Mã vạch 1Z) | ✅ Khớp 100% | ✅ Khớp 100% |

---

## 📋 Quy Chuẩn Định Dạng & Quy Tắc Xử Lý

1. **Đơn vị kiện hàng**:
   - `PAL` (Pallet) trong DO -> Tự động chuyển đổi thành `PP` (Ví dụ: `1 PP`).
   - `CT` (Carton) trong DO -> Chuyển đổi thành `CT` (Ví dụ: `3 CT`).
2. **Quy tắc gộp & cộng dồn trùng lặp**:
   - Nhiều dòng trong DO có cùng `Số tờ khai nhập (IMPORT CD)` sẽ được gộp lại thành 1 dòng duy nhất trên BK và Mẫu 25 NQ, số lượng PCS được cộng dồn chính xác.
   - Dòng tổng kết cuối Mẫu 25 NQ tự động hiển thị: `<Số_kiện> <Đơn_vị> = <Tổng_PCS> PCS` (Ví dụ: `1 PP = 100001 PCS` hoặc `3 CT = 47500 PCS`).
3. **Phân loại tên hàng tiếng Việt**:
   - Tự động nhận diện từ mô tả hàng hóa theo thứ tự ưu tiên: `Ăng ten`, `Cáp đồng trục`, `Pallet gỗ`, `Biến áp`, `Đầu nối`, `Bộ lọc`.
4. **Phương tiện vận chuyển (Số PTVC) & Cửa khẩu xuất**:
   - **FedEx**: Số PTVC `FX6002`, Cửa khẩu `FEDEX`.
   - **DHL**: Số PTVC `8K0525`, Cửa khẩu `DHL`.
   - **UPS Courier**: Số PTVC `5X0074`, Cửa khẩu `UPS`.
   - **Đường hàng không (AIR)**: Trích xuất mã hiệu chuyến bay và ngày bay từ xác nhận đặt chỗ (ví dụ `PVV9646`, `BR6062`, `ACSV`, `NCTS`).

"# ChungTuHQ" 
