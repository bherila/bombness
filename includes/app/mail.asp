<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%><%

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open("DSN=bombness; UID=bombness; PWD=brianjef;")
Set Rs = Conn.Execute("select id, [read], [from], [sent], [subj], message from messages where [to] = " & cstr(cint(request("id"))) & " order by id desc")
Response.Clear()
While Not Rs.EOF
	response.Write(Replace(Replace(rs(0), chr(2), ""), chr(1), "") & chr(2))  		'Message ID
	Response.Write(Replace(Replace(rs(1), chr(2), ""), chr(1), "") & chr(2))  		'Read?
	set rs2 = conn.execute("select top 1 username from users where id = " & rs(2))
		if rs2.bof and rs2.eof then
			username = "Anonymous"
		else
			username = rs2(0)
		end if
		Response.Write(Replace(Replace(username, chr(2), ""), chr(1), "") & chr(2)) 	'From
		rs2.close
	Response.Write(Replace(Replace(rs(3), chr(2), ""), chr(1), "") & chr(2))  		'Date Sent
	Response.Write(Replace(Replace(rs(4), chr(2), ""), chr(1), "") & chr(2))  		'Subject
	Response.Write(Replace(Replace(rs(5), chr(2), ""), chr(1), "") & chr(2))  		'Message
	response.Write(chr(1))
	rs.movenext
wend
set rs2 = nothing
rs.close
set rs = Conn.execute("select count(*) from messages where [to] = " & cstr(cint(request("id"))) & " and [read] = 0")
response.Write(rs(0))
rs.close
set rs = nothing
Conn.Close
Set Conn = Nothing
Response.End()

%>