<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%><%

set conn = server.CreateObject("adodb.connection")
conn.open("Driver={MySQL ODBC 3.51 Driver};Server=bherila.net;uid=mexicouf_bherila;pwd=eggbert;database=mexicouf_logs;")
set rs = Conn.Execute("select body from sent_mail where id = '" & Replace(Request("id"), "'", "''") & "' and `to` = '" & Replace(Request("address"), "'", "''") & "' limit 1")
if rs.eof and rs.bof then
	body = "E-Mail Not Found"
else
	body = rs(0)
end if
rs.close
set rs = nothing
conn.close
set conn = nothing

%><%= body %>