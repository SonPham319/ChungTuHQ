#Requires AutoHotkey v2.0
#SingleInstance Force
SendMode "Input"
SetTitleMatchMode 2

F4::{
	; ngày bắt đầu vận chuyển (ngày làm chứng từ)
	send "04112026"

	send "{tab}^a{Delete}"

	; giờ bắt đầu vận chuyển
	send "08"

	send "{tab}"

	; ngày kết thúc vận chuyển
	send "05112026"

	send "{tab}^a{Delete}"

	; giờ kết thúc vận chuyển
	send "12"

	send "{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}^a{Delete}"
	send "DINH KEM BANG KE SO . DONG CHUNG XE VOI TO KHAI"
	send "{tab}^a{Delete}"

	; số bk
	send "9665270204"

	send "{tab}{tab}^a{Delete}"

	; tên hàng
	send "ĂNG TEN"

	send "{tab}^a{Delete}"

	; mã hs
	send "8517"

	send "{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}"

	; mã iso của nước ship đến
	send "GB"

	send "ZZZ"
	send "{tab}{tab}{tab}^a{Delete}"

	; ptvc / ngày
	send "8K0525/05NOV"

	send "{tab}{tab}"

	; ngày ptvc đi
	send "05112026"

	send "{tab}{tab}"

	; tên người nhập khẩu
	send "RS components Limited"

	send "{tab}"

	; địa chỉ người nhập khẩu
	send "Accounts Payable Department Birchington Road Corby Northamptonshire NN17 9RS UNITED KINGDOM"

	send "{tab}{tab}"
	send "TAOGLAS LIMITED (IRELAND)"
	send "{tab}"
	send "U5 Kilcannon Business Park,Old Dublin Road,Enniscorthy Co. Wexford Y21 XW56 IRELAND"
	send "{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}"

	; số lượng kiện
	send "1"

	send "{tab}"

	; đơn vị kiện ct / pp
	send "ct"

	send "{tab}"

	; gw (kgs)
	send "1.8"

	send "{tab}"
	send "kgm"
	send "{tab}"

	; thể tích
	send "0.02"

	send "{tab}"
	send "mtq"
	send "{tab}"
}