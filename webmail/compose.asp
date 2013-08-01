<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/BN-CompactLayout.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>Bombness Networks</title>
<!-- InstanceEndEditable --><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	background-color: #CCCCCC;
}
body,td,th {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 10pt;
	color: #000000;
}
a:link {
	color: #0066CC;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #0000CC;
}
a:hover {
	text-decoration: underline;
	color: #009966;
}
a:active {
	text-decoration: none;
	color: #FF0000;
}
.navlink {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 9pt;
	font-weight:bold;
	color: #374278;
}
h1,h2,h3,h4,h5,h6 {
	font-family: "Franklin Gothic Medium", Verdana, Arial;
	font-weight: normal;
}
h1 {
	font-size: 18pt;
}
-->
</style>
<!-- InstanceBeginEditable name="head" -->
<style type="text/css">
<!--
.style1 {font-size: 10px}
-->
</style>
<!-- InstanceEndEditable -->
</head>

<body>
<table width="775" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="../layout-2005/logo1.jpg" width="166" height="102"></td>
    <td><img src="../layout-2005/logo2.jpg" width="601" height="102"></td>
    <td><img src="../layout-2005/logo3.gif" width="8" height="102"></td>
  </tr>
  <tr>
    <td align="left" valign="top" background="../layout-2005/nav.jpg" style="padding-left: 18px;"><!-- InstanceBeginEditable name="Sidebar" -->
<%

	set cn = server.createobject("adodb.connection")
	cn.open("dsn=bombness-mysql;uid=root;pwd=eggbert;")
	set rs = cn.execute("select accountmaxsize from mail.hm_accounts where accountid = '" & Replace(session("uid"), "'", "''") & "'")
	maxsize = ccur(rs(0))
	rs.close
	set rs = cn.execute("select CAST(sum(messagesize) as CHAR) from mail.hm_messages where messageaccountid = '" & Replace(session("uid"), "'", "''") & "'")
if isNull(rs(0)) then 
	usedspace = 0
else
	usedspace = ccur(rs(0))
end if
	rs.close
	freespace = maxsize - usedspace
	set rs = nothing
	cn.close
	set cn = nothing

%>
    <p><a href="mailbox.asp">Read Messages</a><br>
      <a href="compose.asp">New Message</a><br>
      <a href="chpwd.asp">Change Password</a>    </p>
    <p><a href="logout.asp">Log Out</a> </p>
    <p><%= FormatNumber(freespace / 1000000, 1) %> MB remaining </p>
    <!-- InstanceEndEditable --></td>
    <td height="600" align="left" valign="top" background="../layout-2005/main.gif" bgcolor="#FFFFFF" style="background-repeat: repeat-x; padding-left: 20px; padding-right: 20px;"><!-- InstanceBeginEditable name="PageTitle" -->
      <h1>New Message</h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
<%

if len(Request("to")) > 0 then
	Dim ObjSendMail
	Set ObjSendMail = CreateObject("CDO.Message")
	
	'This section provides the configuration information for the remote SMTP server.
	
	ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'Send the message using the network (SMTP over the network).
	ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") ="bombness.com"
	ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = Session("username")
	ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Session("password")
	
	' If your server requires outgoing authentication uncomment the lines bleow and use a valid email address and password.
	'ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication
	'ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") ="somemail@yourserver.com"
	'ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") ="yourpassword"
	
	ObjSendMail.Configuration.Fields.Update
	
	'End remote SMTP server configuration section==
	
	ObjSendMail.To = Request("to")
	ObjSendMail.Subject = Request("subject")
	ObjSendMail.From = Session("username")
	
	' we are sending a text email.. simply switch the comments around to send an html email instead
	'ObjSendMail.HTMLBody = "this is the body"
	ObjSendMail.TextBody = Request("message")
	
	ObjSendMail.Send
	
	Set ObjSendMail = Nothing 
end if

%>
      <form name="form1" method="post" action="">
        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr align="left" valign="top">
            <td width="100" height="32">To:</td>
            <td height="32"><input name="to" type="text" id="to"></td>
          </tr>
          <tr align="left" valign="top">
            <td width="100" height="32">From:</td>
            <td height="32"><input name="from" type="text" id="from" value="<%= Session("username") %>" readOnly></td>
          </tr>
          <tr align="left" valign="top">
            <td width="100" height="32">Subject:</td>
            <td height="32"><input name="subject" type="text" id="subject"></td>
          </tr>
          <tr align="left" valign="top">
            <td width="100">Message:</td>
            <td><textarea name="message" cols="70" rows="15" wrap="VIRTUAL" id="message"></textarea></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td><input type="submit" name="Submit" value="Send"></td>
          </tr>
        </table>
      </form>
    <!-- InstanceEndEditable --></td>
    <td background="../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
