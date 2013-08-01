<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
denied = false
if len(Request("username")) > 0 then
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open("DSN=bombness; UID=bombness; PWD=brianjef;")
	sql = "select top 1 id from users where username = '" & Replace(Request("username"), "'", "''") & "' and password = '" & Replace(Request("password"), "'", "''") & "'"
	set rs = conn.execute(sql)
	if rs.eof and rs.bof then
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
		denied = true
	else
		userid = rs(0)
		userip = Request.ServerVariables("REMOTE_ADDR")
		sql = "delete from [sessions] where uid = " & userid
		conn.execute(sql)
		sql = "insert into [sessions] (uid, ip) values (" & userid & ", '" & Replace(userip, "'", "''") & "')"
		conn.execute(sql)
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
		Response.Redirect(Request("jump"))
	end if
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
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
<script>
<!--
function fc(){document.f.username.focus();}
function fd(){document.f.password.focus();}
// -->
</script>
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
          <p>You have reached this page by visiting a Bombness Networks site and requesting to log in.</p>
          <p>Please log in with your Bombness Network Identity details to continue. </p>
        <!-- InstanceEndEditable --></td>
    <td height="600" align="left" valign="top" background="../layout-2005/main.gif" bgcolor="#FFFFFF" style="background-repeat: repeat-x; padding-left: 20px; padding-right: 20px;"><!-- InstanceBeginEditable name="PageTitle" -->
    <h1>Sign In </h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
      <p>Please enter your user name and password to continue. Once verified, you will be returned to the site you came from.</p>
      <p>&nbsp;</p>
      <% if denied then %>
      <table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><span class="style1"><img src="error.gif" width="20" height="20" hspace="10" align="absmiddle"></span></td>
          <td><p class="style1"> The user name and password combination you entered could not be recognized. Please make sure both are correct, then try again. </p></td>
        </tr>
      </table>
      <% end if %>
      <form action="" method="post" name="f" id="f">
        <table border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="90">User Name: </td>
            <td><input name="username" type="text" id="username" style="border: 1px solid #CCCCCC; font-family: Verdana, Arial, Helvetica, sans-serif; font-size:10pt; width: 180px;" value="<%= Request("username") %>" maxlength="50"></td>
          </tr>
          <tr>
            <td width="90" style="padding-top: 5px; ">Password:</td>
            <td style="padding-top: 5px; "><input name="password" type="password" id="password" style="border: 1px solid #CCCCCC; font-family: Verdana, Arial, Helvetica, sans-serif; font-size:10pt; width: 180px;" maxlength="50"></td>
          </tr>
          <tr>
            <td width="90" style="padding-top: 15px; "><input name="jump" type="hidden" id="jump" value="<%= Request("jump") %>"></td>
            <td nowrap style="padding-top: 15px; "><input name="imageField" type="image" src="sign-in.gif" width="66" height="24" border="0"> <a href="<%= Request("jump") %>"><img src="cancel.gif" width="66" height="24" border="0"></a></td>
          </tr>
        </table>
      </form>
      <p>&nbsp; </p>
	  <script><% if denied then %>fd();<% else %>fc();<% end if %></script>
      <!-- InstanceEndEditable --></td><td background="../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
