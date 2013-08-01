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
		Response.Redirect(Request("dest"))
	end if
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/BN-IdentityPage.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>Untitled Document</title>
<!-- InstanceEndEditable --><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body {
	background-image:   url(background.jpg);
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	background-repeat: no-repeat;
	line-height: 17px;
}
body,td,th {
	font-family: Tahoma, Verdana, Arial, Helvetica, sans-serif;
	font-size: 11px;
}
-->
</style>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
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

<body scroll="no">

<div id="Layer1" style="position:absolute; left:29px; top:53px; width:249px; height:330px; z-index:1"><!-- InstanceBeginEditable name="Main" -->
  <p>Welcome to Bombness Identities. Please sign in with your user name and password, or <a href="signup.asp">create a free account</a> today! </p>
  <% if denied then %><p class="style1">The user name and password you entered are not correct. Please try again. </p><% end if %>
  <form name="form1" method="post" action="">
    <table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td style="padding-top: 5px; " width="80">User Name: </td>
        <td style="padding-top: 5px; "><input name="username" type="text" id="username" style="border: 1px solid #CCCCCC; font-family: Verdana, Arial, Helvetica, sans-serif; font-size:10pt; width: 130px;" value="<%= Request("username") %>" maxlength="50"></td>
      </tr>
      <tr>
        <td style="padding-top: 5px; " width="80">Password:</td>
        <td style="padding-top: 5px; "><input name="password" type="password" id="password" style="border: 1px solid #CCCCCC; font-family: Verdana, Arial, Helvetica, sans-serif; font-size:10pt; width: 130px;" maxlength="50"></td>
      </tr>
      <tr>
        <td style="padding-top: 10px; "><input name="dest" type="hidden" id="dest" value="<%= Request("dest") %>"></td>
        <td style="padding-top: 10px; "><input name="imageField" type="image" src="sign-in.gif" width="66" height="24" border="0"></td>
      </tr>
    </table>
  </form>
  <p><strong><a href="signup.asp">Create an Identity</a> </strong></p>
<!-- InstanceEndEditable --></div>
</body>
<!-- InstanceEnd --></html>
