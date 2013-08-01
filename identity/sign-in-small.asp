<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bombness/includes/_id.asp"-->
<%
'Response.Write(site & ":" & remoteip)
'SetSite(Request("site"))
eval_loginStatus()
if bh_loggedin then
%>
<script>
top.document.location.reload();
</script>
<%
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Log In</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body,td,th {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 10pt;
	color: #000000;
}
body {
	background-color: #FFFFFF;
	margin-left: 15px;
	margin-top: 15px;
	margin-right: 15px;
	margin-bottom: 15px;
}
.style1 {font-weight: bold}
.style2 {
	color: #CC0000;
	font-weight: bold;
}
-->
</style>
<script>
<!--
function fc(){document.f.username.focus();}
function fd(){document.f.password.focus();}
// -->
</script>
</head>
<body>
      <p>Please enter your user name and password to continue. If others use this computer, don't forget to sign out when you're finished.</p>
      <% if denied then %>
      <table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><span class="style1"><img src="error.gif" width="20" height="20" hspace="10" align="absmiddle"></span></td>
          <td><p class="style2"> The user name and password combination you entered could not be recognized. Please make sure both are correct, then try again. </p></td>
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
            <td width="90" style="padding-top: 15px; "><input name="action" type="hidden" id="action" value="bn-login">
            <input name="site" type="hidden" id="site" value="<%= Request("site") %>"></td>
            <td nowrap style="padding-top: 15px; "><input name="submit" type="submit" id="submit" style="width: 66px; height: 24px; background-image:url(sign-in.gif); border: none; background-color: white;" value="">
            <input name="reset" type="reset" id="reset" style="width: 66px; height: 24px; background-image:url(cancel.gif); border: none; background-color: white;" onClick="top.document.getElementById('loginA').style.visibility='hidden'; top.document.getElementById('loginB').style.visibility='hidden';" value="">              </td>
          </tr>
        </table>
      </form>
      <p>&nbsp; </p>
	  <script>function fx() {<% if denied then %>fd();<% else %>fc();<% end if %>}</script>
</body>
</html>
