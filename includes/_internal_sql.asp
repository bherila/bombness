<%
on error resume next
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open("DSN=bombness; UID=bombness; PWD=brianjef;")
sql = Split(Request("sql"), vbNewLine)
Response.Write(ubound(sql))
'Response.Write(Request("sql"))
dim i
i = 0
for i = lbound(sql) to (ubound(sql) - 1)
	Conn.Execute(sql(i))
	if err.number <> 0 then
		response.write(vbnewline)
		Response.Write(i)
		Response.Write("> ")
		response.write(sql(i))
		response.write(": ")
		response.write(err.description)
	else
	end if
	Response.Flush()
	err.number = 0
next
Conn.Close
Set Conn = Nothing
%>
Done!