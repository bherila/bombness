<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%><%
Set cn = Server.CreateObject("ADODB.Connection")
cn.Open("DSN=bombness; UID=bombness; PWD=brianjef;")
set rs = cn.execute("select top 1 id from users where username = '" & Replace(Request("username"), "'", "''") & "' and password = '" & Replace(Request("password"), "'", "''") & "'")
response.Clear()
if not (rs.bof and rs.eof) then
response.Write(rs(0))
else
response.Write("-1")
end if
rs.close
set rs = nothing
cn.close
set cn = nothing
response.End()
%>