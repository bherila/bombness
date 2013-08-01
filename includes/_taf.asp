<%
function tellAFriend(FromName, FromEmail, ToName, ToEmail, ServerEmail, Message)
	On Error Resume Next
	Dim Cn
	Dim Rs
	Set Cn = Server.CreateObject("ADODB.Connection")
	Cn.Open("DSN=bombness; UID=bombness; PWD=brianjef;")
	Cn.Execute("insert into taf_history (name, email, ip) values ('" & Replace(FromName, "'", "''") & "', '" & Replace(FromEmail, "'", "''") & "', '" & Replace(Request.ServerVariables("REMOTE_ADDR"), "'", "''") & "')")
	Cn.Execute("insert into taf_history (name, email, ip) values ('" & Replace(ToName,   "'", "''") & "', '" & Replace(ToEmail,   "'", "''") & "', '" & Replace(Request.ServerVariables("REMOTE_ADDR"), "'", "''") & "')")
	Cn.Execute("update taf_history set instances = instances + 1 where email = '" & Replace(FromEmail, "'", "''") & "' or email = '" & Replace(ToEmail, "'", "''") & "'")
	Cn.Close
	Set Cn = Nothing
	
	Message = Message & vbNewLine & vbNewLine & "_______________________________________________________" & vbNewLine & "This E-Mail was sent from a Bombness Networks Tell-a-Friend form under our terms of use. If this message is unsolicited and you do not know the person who sent this to you, please forward this entire message to abuse@bombness.com, and we apologize for any inconvenience. To prevent your e-mail address from receiving further e-mail from Bombness Networks, please visit http://www.bombness.com/bombness/abuse ."
	
	Dim MyMail
	Set MyMail = Server.CreateObject("IPWorksASP.Smtp")
	MyMail.MailServer = "sendmail.brinkster.com"
	MyMail.MailPort = 25
	MyMail.User = "submissions@bombness.com"
	MyMail.Password = "eggbert"
	MyMail.From = FromEmail
	MyMail.SendTo = ToEmail
	MyMail.Subject = "It's " & FromName & " - Check this cool site out!"
	MyMail.MessageText = Message
	MyMail.Send
	Set MyMail = Nothing
	
end function

%>