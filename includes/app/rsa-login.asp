<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%><%
Set cn = Server.CreateObject("ADODB.Connection")
cn.Open("Driver={MySQL ODBC 3.51 Driver};Server=orf-mysql1.brinkster.com;uid=bombness;pwd=brianjef;database=bombness;")
set rs = cn.execute("select start, end, leaveoff, target from rsa_users where hardware_key = '" & Replace(Request("key"), "'", "''") & "' limit 1")
response.Clear()
if not (rs.bof and rs.eof) then
response.Write(rs(0) & ":" & rs(1) & ":" & rs(2) & ":" & rs(3))
else
response.Write("ERR")
end if
rs.close
set rs = nothing
cn.close
set cn = nothing
response.End()
%>