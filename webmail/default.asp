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
<!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
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
    <p>Enter your e-mail<br>
      address and password<br>
      to continue. </p>
    <!-- InstanceEndEditable --></td>
    <td height="600" align="left" valign="top" background="../layout-2005/main.gif" bgcolor="#FFFFFF" style="background-repeat: repeat-x; padding-left: 20px; padding-right: 20px;"><!-- InstanceBeginEditable name="PageTitle" -->
      <h1>Webmail Login </h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
      <p>Please enter your e-mail address and password to access your mailbox.</p>
<%

if len(Request("email")) > 0 then
	dest =""
	set cn = server.createobject("adodb.connection")
	cn.open("dsn=bombness-mysql;uid=root;pwd=eggbert;")
	set rs = cn.execute("select accountid from mail.hm_accounts where accountaddress = '" & Replace(Request("email"), "'", "''") & "' and accountpassword = MD5('" & Replace(Request("password"), "'", "''") & "')")
	if rs.eof and rs.bof then
		Response.Write("<p style=""color: red;"">Your user name and password were not acceptable. Please check your credentials, and try again.</p>")
	else
		Session("uid") = rs(0)
		session("username") = request("email")
		session("password") = request("password")
		dest = "mailbox.asp"
	end if
	rs.close
	cn.close
	set rs = nothing
	set cn = nothing
	if len(dest) > 2 then Response.Redirect(dest)
	Response.Write("<p style=""color: green;""><b>Message Sent!</b></p>")
end if

%>
      <form name="form1" method="post" action="">
        <table border="0" cellspacing="0" cellpadding="0">
          <tr align="left" valign="top">
            <td width="80" height="30" nowrap>E-Mail:</td>
            <td height="30"><input name="email" type="text" id="email" value="<%= Request("email") %>"></td>
          </tr>
          <tr align="left" valign="top">
            <td width="80" height="30" nowrap>Password:</td>
            <td height="30"><input name="password" type="password" id="password"></td>
          </tr>
          <tr align="left" valign="top">
            <td width="80" height="30">&nbsp;</td>
            <td height="30"><input type="submit" name="Submit" value="Login"></td>
          </tr>
        </table>
      </form>
      <p>&nbsp; </p>
      <!-- InstanceEndEditable --></td>
    <td background="../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
