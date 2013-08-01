<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%><%
		
function old_sendmail(mail_connection, log_connection, xID, xTo, xFrom, xSubj, xBody)
	dim yBody
	yBody = cStr(xBody)
	dim mail_rs
	set mail_rs = mail_connection.execute("select top 1 id, smtp_server, smtp_username, smtp_password from mail_bots where mails_sent < 1500")
	if not mail_rs.eof then
		mail_connection.execute("update mail_bots set mails_sent = mails_sent + 1 where id = " & mail_rs(0))
        Dim CDOSYSMsg
        Dim Config
        Dim Fields
        set CDOSYSMsg = Server.CreateObject("CDO.Message")
		set Config = Server.CreateObject("CDO.Configuration")
		set Fields = Config.Fields
		With Fields
			.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = mail_rs(1)
			sf = mail_rs(2)
			.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = sf
			.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = mail_rs(3)
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
			.Update()
		End With
        CDOSYSMsg.Configuration = Config
        CDOSYSMsg.To = xTo
        CDOSYSMsg.From = xFrom
        CDOSYSMsg.Subject = xSubj
        'CDOSYSMsg.TextBody = xBody
		if uCase(Left(yBody, 6)) = "<HTML>" then
			CDOSYSMsg.HTMLBody = cStr(yBody)
			CDOSYSMsg.TextBody = "This message contains rich formatting that this e-mail reader cannot display. Please click the following link to display the contents of this message." & vbnewline & vbnewline & "http://www.bombness.com/bombness/readmail.asp?id=" & xid & "&address=" & Server.URLEncode(xTo)
		elseif uCase(left(yBody, 6)) = "MHTML:" then
			pts = split(yBody, ":", 2)(1)
			Response.Write(pts)
			CDOSYSMsg.TextBody = "This message contains rich formatting that this e-mail reader cannot display. Please click the following link to display the contents of this message." & vbnewline & vbnewline & "http://www.bombness.com/bombness/readmail.asp?id=" & xid & "&address=" & Server.URLEncode(xTo)
			CDOSYSMsg.CreateMHTMLBody pts
		else
			CDOSYSMsg.TextBody = cStr(yBody)
		end if
		sql = "insert into sent_mail (`id`, `to`, `from`, `subject`, `body`, `user`) values (" & xID & ", '"
		sql = sql & Replace(xTo, "'", "''")
		sql = sql & "', '" & Replace(xFrom, "'", "''")
		sql = sql & "', '" & Replace(xSubj, "'", "''")
		sql = sql & "', '" & Replace(yBody, "'", "''") & "', '" & sf & "')"
        CDOSYSMsg.Send()
		log_connection.execute(sql)
        set CDOSYSMsg = Nothing
        set Config = Nothing
        set Fields = Nothing
		mail_rs.close
		sendmail = true
		set mail_rs = nothing
	else
		mail_rs.close
		set mail_rs = nothing
		sendmail = false
	end if
end function

server.ScriptTimeout = 10000

	set conn = server.CreateObject("adodb.connection")
	conn.open("dsn=bombness;uid=bombness;pwd=brianjef")
	set conn2 = server.CreateObject("adodb.connection")
	conn2.open("Driver={MySQL ODBC 3.51 Driver};Server=bherila.net;uid=mexicouf_bherila;pwd=eggbert;database=mexicouf_logs;")
	set rs = conn.execute("select top 5 [id], [to], [from], [subject], [body] from outbox where GetUTCDate() > [send_date] order by id asc")
	s = 0
	while not rs.eof
		on error resume next
		s = s + 1
		id = rs("id")
		old_sendmail conn, conn2, rs("id"), rs("to"), rs("from"), rs("subject"), rs("body")
		conn.execute("delete from outbox where id = " & id)
		on error goto 0
		rs.movenext
	wend
	rs.close
	set rs = conn.execute("select count(*) from outbox")
	rm = rs(0)
	Response.Write(s & " just sent. " & rm & " email(s) in queue")
	rs.close
	set rs = nothing
	conn.close()
	set conn = nothing
	conn2.close()
	set conn2 = nothing


%><script language="javascript">setTimeout("location.reload();", <%

if rm > 0 then
	Response.Write("1")
else
	Response.Write("10000")
end if

%>);</script>