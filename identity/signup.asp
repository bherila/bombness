<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
Session("attempts") = 0
UserTaken = False
EmailTaken = False

if len(Request("username")) > 0 and len(Request("email")) > 0 and len(Request("password")) > 0 then
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open("DSN=bombness; UID=bombness; PWD=brianjef;")
	sql = "select count(*) from users where username = '" & Replace(Request("username"), "'", "''") & "'"
	
	set rs = conn.execute(sql)
	usertaken = (rs(0) > 0)
	rs.close
	
	sql = "select count(*) from users where email = '" & Replace(Request("email"), "'", "''") & "'"
	set rs = conn.execute(sql)
	emailtaken = (rs(0) > 0)
	rs.close
	
	set rs = nothing
	conn.close
	set conn = nothing
	
	Session("username") = Request("username")
	Session("password") = Request("password")
	Session("email") = Request("email")
	
	if (usertaken = false) and (emailtaken = false) then Response.Redirect("turing.asp")
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="../Templates/BN-IdentityPage.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>Untitled Document</title>
<!-- InstanceEndEditable --><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body {
	background-image:    url(background.jpg);
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
.style1 {color: #CC0000}
-->
</style>
<!-- InstanceEndEditable -->
</head>

<body scroll="no">

<div id="Layer1" style="position:absolute; left:29px; top:53px; width:249px; height:330px; z-index:1"><!-- InstanceBeginEditable name="Main" -->
<% if Len(Request("username")) < 1 and len(Request("password")) < 1 and len(Request("email")) < 1 then %>
<p>Please select a user name and password, and enter your e-mail address. You will need to confirm your e-mail address later to use some features of Bombness Networks. </p>
<% elseif Len(Request("username")) < 1 or len(Request("password")) < 1 or len(Request("email")) < 1 then %>
<p><strong>All fields on this page are required. Please make sure they are all filled in. </strong></p>
<% end if %>
<% if UserTaken then %><p class="style1"><strong>The user name you selected is already in use by another member. </strong></p>
<% end if %>
<% if EmailTaken then %><p class="style1"><strong>The e-mail address you entered is already in use by another member.</strong></p>
<% end if %>
<form name="form1" method="post" action="">
  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td style="padding-top: 5px; " width="80">User Name: </td>
      <td style="padding-top: 5px; "><input name="username" type="text" id="username" style="border: 1px solid #CCCCCC; font-family: Verdana, Arial, Helvetica, sans-serif; font-size:10pt; width: 130px;" value="<%= Request("username") %>" maxlength="50">
<% if UserTaken or len(Request("username")) < 1 then %><span class="style1">*</span><% end if %></td>
    </tr>
    <tr>
      <td style="padding-top: 5px; " width="80">Password:</td>
      <td style="padding-top: 5px; "><input name="password" type="password" id="password" style="border: 1px solid #CCCCCC; font-family: Verdana, Arial, Helvetica, sans-serif; font-size:10pt; width: 130px;" value="<%= Request("password") %>" maxlength="50">
<% if len(Request("password")) < 1 then %><span class="style1">*</span><% end if %></td>
    </tr>
    <tr>
      <td style="padding-top: 5px; ">E-Mail Address:</td>
      <td style="padding-top: 5px; "><input name="email" type="text" id="email" style="border: 1px solid #CCCCCC; font-family: Verdana, Arial, Helvetica, sans-serif; font-size:10pt; width: 130px;" value="<%= Request("email") %>" maxlength="255">
<% if EmailTaken or len(Request("email")) < 1 then %><span class="style1">*</span><% end if %></td>
    </tr>
    <tr>
      <td style="padding-top: 10px; ">&nbsp;</td>
      <td style="padding-top: 10px; "><input name="imageField" type="image" src="next.gif" width="66" height="24" border="0"></td>
    </tr>
  </table>
</form>
<p>&nbsp; </p>
<!-- InstanceEndEditable --></div>
</body>
<!-- InstanceEnd --></html>
