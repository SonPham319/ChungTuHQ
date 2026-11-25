#Requires AutoHotkey v2.0
#SingleInstance Force
SendMode "Input"
SetTitleMatchMode 2

F4::{
	; ngày bắt đầu vận chuyển (ngày làm chứng từ)
	send "02112026"

	send "{tab}^a{Delete}"

	; giờ bắt đầu vận chuyển
	send "08"

	send "{tab}"

	; ngày kết thúc vận chuyển
	send "03112026"

	send "{tab}^a{Delete}"

	; giờ kết thúc vận chuyển
	send "12"

	send "{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}^a{Delete}"
	send "DINH KEM BANG KE SO . DONG CHUNG XE VOI TO KHAI"
	send "{tab}^a{Delete}"

	; số bk
	send "69558911893"

	send "{tab}{tab}^a{Delete}"

	; tên hàng
	send "CÁP ĐỒNG TRỤC"

	send "{tab}^a{Delete}"

	; mã hs
	send "8544"

	send "{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}"

	; mã iso của nước ship đến
	send "US"

	send "ZZZ"
	send "{tab}{tab}{tab}^a{Delete}"

	; ptvc / ngày
	send "BR6062/06NOV"

	send "{tab}{tab}"

	; ngày ptvc đi
	send "06112026"

	send "{tab}{tab}"

	; tên người nhập khẩu
	send "Mouser Electronics, Inc."

	send "{tab}"

	; địa chỉ người nhập khẩu
	send "1000 N Main St Mansfield TX 76063-1514 USA"

	send "{tab}{tab}"
	send "TAOGLAS LIMITED (IRELAND)"
	send "{tab}"
	send "U5 Kilcannon Business Park,Old Dublin Road,Enniscorthy Co. Wexford Y21 XW56 IRELAND"
	send "{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}"

	; số lượng kiện
	send "1"

	send "{tab}"

	; đơn vị kiện ct / pp
	send "pp"

	send "{tab}"

	; gw (kgs)
	send "283"

	send "{tab}"
	send "kgm"
	send "{tab}"

	; thể tích
	send "1.09"

	send "{tab}"
	send "mtq"
	send "{tab}"
}
