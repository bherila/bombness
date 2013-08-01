<%@LANGUAGE="VBScript"%>
<%
on error resume next

Dim Item_name, Item_number, Payment_status, Payment_amount
Dim Txn_id, Receiver_email, Payer_email
Dim objHttp, str

' read post from PayPal system and add 'cmd'
str = Request.Form & "&cmd=_notify-validate"

' post back to PayPal system to validate
set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
' set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
' set objHttp = Server.CreateObject("Microsoft.XMLHTTP")
objHttp.open "POST", "https://www.paypal.com/cgi-bin/webscr", false
'objHttp.open "POST", "https://www.eliteweaver.co.uk/cgi-bin/webscr", false
objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
objHttp.Send str

' assign posted variables to local variables
Item_name = Request.Form("item_name")
Item_number = Request.Form("item_number")
Payment_status = Request.Form("payment_status")
Payment_amount = Request.Form("mc_gross")
Payment_currency = Request.Form("mc_currency")
Txn_id = Request.Form("txn_id")
Receiver_email = Request.Form("receiver_email")
Payer_email = Request.Form("payer_email")

set conn = server.CreateObject("adodb.connection")
conn.open("dsn=bombness;uid=bombness;pwd=brianjef")
cols = ""
vals = ""

function add(what)
	cols = cols & ", [" & what & "]"
	if len(request(what)) < 1 then
		vals = vals & ", NULL"
	elseif len(request(what)) < 4 then
		vals = vals & ", '" & replace(request(what), "'", "''") & "'"
	else
		if ucase(right(request(what), 3) = "PST") or ucase(right(request(what), 3) = "PDT") then
			vals = vals & ", '" & Left(replace(request(what), "'", "''"), len(request(what)) - 3) & "'"		
		else
			vals = vals & ", '" & replace(request(what), "'", "''") & "'"
		end if
	end if
end function

function addVal(col, val)
	cols = cols & ", [" & col & "]"
	if len(val) < 1 then
		vals = vals & ", NULL"
	elseif len(val) < 4 then
		vals = vals & ", '" & replace(val, "'", "''") & "'"
	else
		if ucase(right(val, 3) = "PST") or ucase(right(val, 3) = "PDT") then
			vals = vals & ", '" & Left(replace(val, "'", "''"), len(val) - 3) & "'"		
		else
			vals = vals & ", '" & replace(val, "'", "''") & "'"
		end if
	end if
end function

if request("item_number") = "BN-HP" and request("business") = "bherila@bombness.com" then
	pts = split(cstr(request("mc_gross")), ".")
	amt = cint(pts(0))
	hp = amt * 1000
	user = replace(request("option_selection1"), "'", "''")
	if len(request("option_selection1")) < 1 then
		pt = request("parent_txn_id")
		set rs = conn.execute("select top 1 option_selection1 from ipn where txn_id = '" & replace(pt, "'", "''") & "'")
		user = replace(rs(0), "'", "''")
		rs.close
		set rs = nothing
	end if
	xsql = "update users set hp = hp + " & hp & " where username = '" & user & "'"
	conn.execute(xsql)
	conn.execute("insert into finance ([desc], amt) values ('PayPal Transaction (TXN_ID " & Replace(txn_id, "'", "''") & ")', " & Request("mc_gross") & ")")
	conn.execute("insert into finance ([desc], amt) values ('PayPal Transaction Fee (TXID " & Replace(txn_id, "'", "''") & ")', " & cstr(ccur(Request("mc_fee")) * -1) & ")")
	conn.execute("insert into errors (error) values ('" & replace(xsql, "'", "''") & "')")
end if

add("txn_id")
add("parent_txn_id")
add("item_name")
add("item_number")
add("quantity")
add("mc_gross")
add("mc_fee")
add("settle_amount")
add("exchange_rate")
add("custom")
add("memo")
add("note")
add("tax")
add("option_name1")

c1_name = request("option_name1")
c1_value = request("option_selection1")
c2_name = request("option_name2")
c2_value = request("option_selection2")
productID = request("item_number")
amount = request("mc_gross")

add("option_selection1")
add("option_name2")
add("option_selection2")
add("payment_status")
add("pending_reason")
add("reason_code")
add("payment_date")
add("first_name")
add("last_name")
add("payer_business_name")
add("address_name")
add("address_street")
add("address_city")
add("address_state")
add("address_zip")
add("address_country")
add("address_status")
add("payer_email")
add("payer_status")
add("txn_type")
add("subscr_date")
add("subscr_effective")
add("period1")
add("period2")
add("period3")
add("mc_amount1")
add("mc_amount2")
add("mc_amount3")
add("mc_currency")
add("username")
add("password")
add("subscr_id")

' Check notification validation
if (objHttp.status <> 200 ) then
' HTTP error handling
elseif (objHttp.responseText = "VERIFIED") then
' check that Payment_status=Completed
' check that Txn_id has not been previously processed
' check that Receiver_email is your Primary PayPal email
' check that Payment_amount/Payment_currency are correct
' process payment
addVal "status", "2"
elseif (objHttp.responseText = "INVALID") then
' log for manual investigation
addVal "status", "1"
else
' error
addVal "status", "0"
end if

vals = right(vals, len(vals) - 1)
cols = right(cols, len(cols) - 1)

sql = "insert into IPN (" & cols & ") values (" & vals & ")"
response.Write(sql)
conn.execute(sql)

if productID = "BH0123095782934" and amount >= 15.00 and request("business") = "bherila@bombness.com" then
	'MoneyPayment
	err.number = 0
	sql = "update money_users set paid = 1 where username = '" & Replace(c1_value, "'", "''") & "'"
	conn.execute(sql)
	if err.number <> 0 then
	sql = "insert into errors (error) values ('" & replace(sql, "'", "''") & "')"
	conn.execute(sql)
	end if
	err.number = 0
	
	set rs = conn.execute("select top 1 id from money_users where username = '" & Replace(c1_value, "'", "''") & "'")
	uid = rs(0)
	rs.close
	set rs = nothing
	conn.execute("insert into money_weight (uid) values (" & cstr(cint(uid)) & ")")
	r = split(c2_value, ":", 2)
	if lbound(r) = 0 and ubound(r) > 0 then
		conn.execute("update money_users set credit = credit + 5 where id = '" & Replace(r(0), "'", "''") & "'")
		conn.execute("update money_users set credit = credit + 5 where id = '" & Replace(r(1), "'", "''") & "'")
	end if
	
		
end if

if productID = "BH0123095782937" and request("business") = "bherila@bombness.com" then
	q = round(amount / 3)
	i = 0
	for i = 1 to q
		conn.execute("insert into money_weight (uid) values (" & cstr(cint(c1_value)) & ")")	
	next
end if

if productID = "BN-SVC" and request("business") = "bherila@bombness.com" then
	tid = Request("txn_id")
	xxx = c1_value
	if len(tid) < 1 and len(Request("parent_txn_id") > 0) then
		set rx = conn.execute("select top 1 txn_id, [option_selection1] from ipn where txn_id = '" & Request("parent_txn_id") & "'")
		tid = rx(0)
		xxx = rx(1)
		rx.close
		set rx = nothing
	end if
	s = "Transaction ID " & tid & " " & Request("payment_status")
	sql = "insert into bherila.finance (uid, name, amt, description) values ('" & Replace(xxx, "'", "''") & "', '" & Replace(s, "'", "''") & "', '" & (amount * -1) & "', '" & Request("memo") & "')"
	conn.execute(sql)
end if

if err.number <> 0 then
sql = "insert into errors (error) values ('" & replace(sql, "'", "''") & "')"
conn.execute(sql)
end if

conn.close
set conn = nothing
set objHttp = nothing
%>