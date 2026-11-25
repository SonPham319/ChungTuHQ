#Requires AutoHotkey v2.0
#SingleInstance Force
SendMode "Input"
SetTitleMatchMode 2

F4::{
	; ngày bắt đầu vận chuyển (ngày làm chứng từ)
	send "04112026"

	send "{tab}^a{Delete}"

	; giờ bắt đầu vận chuyển
	send "12"

	send "{tab}"

	; ngày kết thúc vận chuyển
	send "04112026"

	send "{tab}^a{Delete}"

	; giờ kết thúc vận chuyển
	send "17"

	send "{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}^a{Delete}"
	send "DINH KEM BANG KE SO . DONG CHUNG XE VOI TO KHAI"
	send "{tab}^a{Delete}"

	; số bk
	send "879546740471"

	send "{tab}{tab}^a{Delete}"

	; tên hàng
	send "ĐẦU NỐI"

	send "{tab}^a{Delete}"

	; mã hs
	send "8536"

	send "{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}"

	; mã iso của nước ship đến
	send "CN"

	send "ZZZ"
	send "{tab}{tab}{tab}^a{Delete}"

	; ptvc / ngày
	send "FX6002/05NOV"

	send "{tab}{tab}"

	; ngày ptvc đi
	send "05112026"

	send "{tab}{tab}"

	; tên người nhập khẩu
	send "Aztech Communication Device (DG) Ltd"

	send "{tab}"

	; địa chỉ người nhập khẩu
	send "No.6, Jiujiang Road Chang Ping Town, Dong Guan City,Guang Dong Province CHINA"

	send "{tab}{tab}"
	send "TAOGLAS LIMITED (IRELAND)"
	send "{tab}"
	send "U5 Kilcannon Business Park,Old Dublin Road,Enniscorthy Co. Wexford Y21 XW56 IRELAND"
	send "{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}"

	; số lượng kiện
	send "3"

	send "{tab}"

	; đơn vị kiện ct / pp
	send "ct"

	send "{tab}"

	; gw (kgs)
	send "10.4"

	send "{tab}"
	send "kgm"
	send "{tab}"

	; thể tích
	send "0.09"

	send "{tab}"
	send "mtq"
	send "{tab}"
}