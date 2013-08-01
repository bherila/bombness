<%

function sendmail(xTo, xFrom, xSubj, xBody)
	dim mail_connection
	set mail_connection = server.CreateObject("adodb.connection")
	mail_connection.open("dsn=bombness-mysql;uid=bombness;pwd=brianjef")
	mail_connection.execute("insert into outbox (`to`, `from`, subject, body) values ('" & Replace(xto, "'", "''") & "', '" & Replace(xfrom, "'", "''") & "', '" & Replace(xsubj, "'", "''") & "', '" & Replace(xbody, "'", "''") & "')")
	mail_connection.close
	set mail_connection = nothing
end function

function sendmail2(conn, xTo, xFrom, xSubj, xBody)
	conn.execute("insert into outbox ([to], [from], subject, body) values ('" & Replace(xto, "'", "''") & "', '" & Replace(xfrom, "'", "''") & "', '" & Replace(xsubj, "'", "''") & "', '" & Replace(xbody, "'", "''") & "')")
end function

function sendmail_immed(xTo, xFrom, xSubj, xBody)
	dim mail_connection
	set mail_connection = server.CreateObject("adodb.connection")
	mail_connection.open("dsn=bombness;uid=bombness;pwd=brianjef")
	set mail_rs = mail_connection.execute("select top 1 id, smtp_server, smtp_username, smtp_password from mail_bots where mails_sent < 1500")
	if not mail_rs.eof then
		mail_connection.execute("update mail_bots set mails_sent = mails_sent + 1 where id = " & mail_rs(0))
		Dim MyMail
		Set MyMail = Server.CreateObject("IPWorksASP.Smtp")
		MyMail.MailServer = mail_rs(1)
		MyMail.MailPort = 25
		MyMail.User = mail_rs(2)
		MyMail.Password = mail_rs(3)
		MyMail.From = xFrom
		MyMail.SendTo = xTo
		MyMail.Subject = xSubj
		MyMail.MessageText = xBody
		mail_rs.close
		set mail_rs=nothing
		If MyMail.Send() Then
			sendmail_immed = true
		Else
			sendmail_immed = false
		End If
		Set MyMail = Nothing
	else
		mail_rs.close
		set mail_rs = nothing
		sendmail_immed = false
	end if
	mail_connection.close
	set mail_connection = nothing
end function

%>