def validate_shipment_data(delivery_info, booking_info):
    """
    Checks for missing information, duplicate records, and discrepancies between files.
    Returns a dict with status, issues list, duplicates list, and warnings list.
    """
    issues = []
    warnings = []
    duplicates = []
    
    # 1. Missing information checks on Delivery Order
    if not delivery_info.get('do_code'):
        warnings.append("Mã phiếu xuất kho (DO Code) không tìm thấy.")
    if not delivery_info.get('consignee'):
        warnings.append("Tên người nhận (Consignee) để trống.")
    if not delivery_info.get('total_packages'):
        issues.append("Thiếu số lượng kiện hàng (Packages).")
    if not delivery_info.get('package_unit'):
        warnings.append("Thiếu đơn vị kiện (PAL/CT), mặc định sử dụng PP.")
    if not delivery_info.get('total_gw') or delivery_info.get('total_gw') <= 0:
        warnings.append("Tổng trọng lượng (GW) bằng 0 hoặc không tìm thấy.")
    if not delivery_info.get('grouped_items'):
        issues.append("Không có dòng hàng hóa hợp lệ nào được tìm thấy trong phiếu giao hàng.")

    # Check each item
    for idx, item in enumerate(delivery_info.get('grouped_items', []), 1):
        if not item.get('tk_nhap'):
            issues.append(f"Mục {idx}: Thiếu số tờ khai nhập (IMPORT CD).")
        if not item.get('quantity') or item.get('quantity') <= 0:
            warnings.append(f"Mục {idx} (TK {item.get('tk_nhap')}): Số lượng PCS bằng 0.")
        if not item.get('short_name'):
            warnings.append(f"Mục {idx} (TK {item.get('tk_nhap')}): Chưa xác định tên hàng tiếng Việt.")

    # 2. Duplicate records check
    if delivery_info.get('duplicates_count', 0) > 0:
        for dup in delivery_info.get('duplicates', []):
            duplicates.append({
                'tk_nhap': dup['tk_nhap'],
                'item_code': dup.get('item_code', ''),
                'quantity': dup.get('quantity', 0),
                'row': dup.get('row', '')
            })
        warnings.append(
            f"Phát hiện {delivery_info['duplicates_count']} dòng trùng lặp số TK nhập trong phiếu giao hàng. "
            f"Đã tự động gộp và cộng dồn số lượng chính xác."
        )

    # 3. Missing information checks on Booking / AWB
    if not booking_info.get('bk_no'):
        issues.append("Thiếu Số BK / Vận đơn (MAWB / Tracking No / Waybill).")
    if not booking_info.get('flight'):
        warnings.append("Thiếu số hiệu chuyến bay / Phương tiện vận chuyển (Số PTVC).")
    if not booking_info.get('warehouse'):
        warnings.append("Thiếu thông tin Cửa khẩu xuất / Kho hàng.")

    # 4. Discrepancy checks between DO and Booking/AWB
    if booking_info.get('pkg') is not None and delivery_info.get('total_packages'):
        b_pkg = int(booking_info['pkg'])
        d_pkg = int(delivery_info['total_packages'])
        if b_pkg != d_pkg:
            warnings.append(f"Lệch số kiện: Phiếu xuất kho ghi {d_pkg} kiện, Booking ghi {b_pkg} kiện.")

    if booking_info.get('gw') is not None and delivery_info.get('total_gw'):
        b_gw = float(booking_info['gw'])
        d_gw = float(delivery_info['total_gw'])
        if abs(b_gw - d_gw) > 1.0:
            warnings.append(f"Lệch trọng lượng GW: Phiếu xuất kho ghi {d_gw} kg, Booking ghi {b_gw} kg.")

    # Determine overall status
    if issues:
        status = 'ERROR'
    elif warnings:
        status = 'WARNING'
    else:
        status = 'OK'

    return {
        'status': status,
        'issues': issues,
        'warnings': warnings,
        'duplicates': duplicates,
        'duplicates_count': len(duplicates),
        'is_valid': len(issues) == 0
    }

