<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
dim conn
set conn = server.CreateObject("adodb.connection")
conn.open("dsn=bombness;uid=bombness;pwd=brianjef")
set rs1 = conn.execute("select sec, sum([actions.max] - [actions]) from ads where [actions.max] > 0 and [actions.max] - [actions] > 0 and type = 10 group by sec")
while not rs1.eof
	set rs2 = conn.execute("select top 1 username from ad_users where id = "  & rs1(0))
	response.Write(rs2(0))
	response.Write("--")
	rs2.close
	response.Write(rs1(1)  & " credits remaining<br>")
	rs1.movenext
wend
rs1.close
set rs1 = nothing
set rs2 = nothing
conn.close
set conn = nothing
%>