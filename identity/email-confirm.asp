<%
incorrect = false
if len(Request("pwd")) > 0 then
	set conn = server.createobject("adodb.connection")
	conn.open("dsn=bombness;uid=bombness;pwd=brianjef")
	ws = " where username = '" & Replace(Request("username"), "'", "''") & "' and password = '" & Replace(Request("pwd"), "'", "''") & "' and email='" & Replace(Request("email"), "'", "''") & "'"
	set rs = conn.execute("select count(*) from users" & ws)
	if rs(0) > 0 then
		conn.execute("update users set confirmed = 1" & ws)
	else
		incorrect = true
	end if
	rs.close
	set rs = nothing
	conn.close
	set conn = nothing
	response.redirect("email-confirm-done.asp")
end if

%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="../Templates/BN-CompactLayout.dwt" codeOutsideHTMLIsLocked="false" -->
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
.style1 {
	color: #CC0000;
	font-weight: bold;
}
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
          <p>Confirming your e-mail address will ensure you can access more of Bombness Network's features. </p>
          <p><strong>Page 2 of 3 </strong></p>
        <!-- InstanceEndEditable --></td>
    <td height="600" align="left" valign="top" background="../layout-2005/main.gif" bgcolor="#FFFFFF" style="background-repeat: repeat-x; padding-left: 20px; padding-right: 20px;"><!-- InstanceBeginEditable name="PageTitle" -->
    <h1>Confirm E-Mail </h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
      <% if incorrect then %><p class="style1">Sorry, the password you entered is not correct. </p><% end if %>
      <p>Please enter your password to complete the confirmation process:</p>
      <form action="" method="post" name="form1">
        <strong>
Password: 
<input name="pwd" type="password" id="pwd" value="<%= Request("pwd") %>" maxlength="50">
<input type="submit" name="Submit" value="Finish &gt;">
<input name="username" type="hidden" id="username" value="<%= Request("username") %>">
<input name="email" type="hidden" id="email" value="<%= Request("email") %>">
</strong>
      </form>
      <p>&nbsp;</p>
      <!-- InstanceEndEditable --></td><td background="../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
