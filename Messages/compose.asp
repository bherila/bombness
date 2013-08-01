<%
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1 
%><!--#include virtual="/bombness/includes/_id.asp"--><%
if Request("action") = "send" then
	Set Conn  = Server.CreateObject("ADODB.Connection")
	Set Conn2 = Server.CreateObject("ADODB.Connection")
	Conn.Open("DSN=bombness-mysql; UID=bombness; PWD=brianjef;")
	Conn2.Open("DSN=bombness-mysql; UID=bombness; PWD=brianjef;")

	sql = "select id, email, username from users where username = '" & Replace(Request("to"), "'", "''") & "' limit 1"
	set rs = conn.execute(sql)
	if rs.eof and rs.bof then
		ok = false
	else
		ok = true
		uid = rs(0)
		em=rs(1)
		uname = rs(2)
	end if
	rs.close
	set rs = nothing
	sent = false
	if ok then
		sql = "insert into messages (`to`, `from`, `subj`, `message`) values (" & uid & ", " & CurrentUserID & ", '" & Replace(Request("subject"), "'", "''") & "', '" & Replace(Request("message"), "'", "''") & "')"
		Conn2.Execute(sql)
		sent = true
		msg = "You have a new message! Log in to MeTeen.com or another supporting Bombness Networks site to view it." & vbnewline & vbnewline & "http://www.meteen.com/meteen/login.asp?username=" & server.URLEncode(uname)
		if sendmail(em, "support@meteen.com", "[Bombness Networks] - New Message", msg) = false then
		end if
	end if
	Conn.Close
	Conn2.Close
	set Conn2 = Nothing
	Set Conn = Nothing
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Compose Message</title>
<script language="javascript">
//document.domain = "bombness.com";
</script>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body,td,th {
	font-family: Tahoma, Verdana, Arial, Helvetica, sans-serif;
	font-size: 11px;
}
body {
	background-color: #FFFFFF;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	height: 100%;
	width: 100%;
}
html
{
	height: 100%;
	width: 100%;
}
.style1 {
	color: #FFFFFF;
	font-weight: bold;
}
-->
</style></head>
<body scroll="no">
<table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="0">
  <tr>
    <td align="center" valign="middle" bgcolor="#BBB4A8"><% if not sent then %><form action="compose.asp" method="get" enctype="multipart/form-data" name="frmMain" target="_self" id="frmMain">
      <table width="375"  border="0" cellpadding="5" cellspacing="0" bgcolor="#D4D0C8" style="border: 1px solid #333333;">
      <tr bgcolor="#595446">
        <td colspan="2"><span class="style1">Compose New Message </span></td>
        </tr>
      <tr>
        <td>To:</td>
        <td><input name="to" type="text" id="to" style="width: 300px; " value="<%= Request("to") %>"></td>
      </tr>
      <tr>
        <td>From:</td>
        <td><%= Server.HTMLEncode(CurrentUserName) %></td>
      </tr>
      <tr>
        <td>Subject:</td>
        <td><input name="subject" type="text" id="subject" style="width: 300px; " value="<%= Request("subject") %>" maxlength="50"></td>
      </tr>
      <tr>
        <td>Message:</td>
        <td><textarea name="message" wrap="VIRTUAL" id="message" style="width: 300px; height: 200px; "><%= Request("message") %></textarea></td>
      </tr>
      <tr>
        <td><input name="action" type="hidden" id="action" value="send"></td>
        <td><input type="submit" name="Submit" value="Send"></td>
      </tr>
    </table></form><% else %>
    <p>Message Sent! </p><% end if %></td>
  </tr>
</table>
<script>
parent.frames[0].location.reload();
</script>
</body>
</html>
