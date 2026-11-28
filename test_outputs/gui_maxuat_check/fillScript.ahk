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
	send "315-05281905"

	send "{tab}{tab}^a{Delete}"

	; tên hàng
	send "ĂNG TEN"

	send "{tab}^a{Delete}"

	; mã hs
	send "8517"

	send "{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}"

	; mã iso của nước ship đến
	send "MY"

	send "ZZZ"
	send "{tab}{tab}{tab}^a{Delete}"

	; ptvc / ngày
	send "PVV9646/03NOV"

	send "{tab}{tab}"

	; ngày ptvc đi
	send "03112026"

	send "{tab}{tab}"

	; tên người nhập khẩu
	send "Jabil Circuit Sdn Bhd"

	send "{tab}"

	; địa chỉ người nhập khẩu
	send "56, Hilir Sungai Keluang 1, Phase 4, 11900 Bayan Lepas Industrial Park,Penang MALAYSIA"

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
	send "640"

	send "{tab}"
	send "kgm"
	send "{tab}"

	; thể tích
	send "1.9"

	send "{tab}"
	send "mtq"
	send "{tab}"
}
