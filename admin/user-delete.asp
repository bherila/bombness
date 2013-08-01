<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/BN-Layout.dwt" codeOutsideHTMLIsLocked="false" -->
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
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="20" align="center" valign="top" background="../bnBar/bg.gif"><img src="../bnBar/bombness_sel.gif" width="120" height="20"><img src="../bnBar/separator.gif" width="15" height="20"><a href="http://www.meteen.com" target="_top"><img src="../bnBar/meteen.gif" width="45" height="20" border="0"></a><img src="../bnBar/separator.gif" width="15" height="20"><a href="http://www.highschoolhumor.com" target="_top"><img src="../bnBar/hsh.gif" width="97" height="20" border="0"></a><img src="../bnBar/separator.gif" width="15" height="20"><a href="http://www.perfectjoke.com" target="_top"><img src="../bnBar/pj.gif" width="72" height="20" border="0"></a><img src="../bnBar/separator.gif" width="15" height="20"><a href="http://www.nobullgames.com" target="_top"><img src="../bnBar/nbg.gif" width="84" height="20" border="0"></a></td>
  </tr><tr>
    <td align="center" valign="top">
      <table width="775" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="../layout-2005/logo1.jpg" width="166" height="102"></td>
    <td><img src="../layout-2005/logo2.jpg" width="601" height="102"></td>
    <td><img src="../layout-2005/logo3.gif" width="8" height="102"></td>
  </tr>
  <tr>
    <td align="left" valign="top" background="../layout-2005/nav.jpg" style="padding-left: 18px;"><a href="/bombness/" class="navlink">Home</a><br>
      <a href="/bombness/advertising.asp" class="navlink">Advertising</a><br>
      <a href="/bombness/sites/" class="navlink">Our Sites</a><br>
      <a href="/bombness/identities.asp" class="navlink">BN Identities</a><br>
      <a href="/bombness/investors.asp" class="navlink">Investor Relations</a><br>
      <a href="/bombness/about.asp" class="navlink">Company Information </a></td><td height="600" align="left" valign="top" background="../layout-2005/main.gif" bgcolor="#FFFFFF" style="background-repeat: repeat-x; padding-left: 20px; padding-right: 20px;"><!-- InstanceBeginEditable name="PageTitle" -->
    <h1>Delete User </h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
      <form name="form1" method="post" action="user-delete.asp">
        <p>Enter User ID<br>
          <input name="userid" type="text" id="userid">
  </p>
        <p>Administrative Password<br>
          <input name="password" type="password" id="password"> 
          <input type="submit" name="Submit" value="Go">
</p>
<%
del = 0
if request("password") = "Benjamin11." then
	set sql = Server.CreateObject("adodb.connection")
	set mysql = Server.CreateObject("adodb.connection")
	sql.open("dsn=bombness;uid=bombness;pwd=brianjef")
	mysql.open("Driver={MySQL ODBC 3.51 Driver};Server=orf-mysql1.brinkster.com;uid=bombness;pwd=brianjef;database=bombness;")
	userid = Request("userid")
	sql.execute("delete from users where id = " & userid)
	sql.execute("delete from user_pics where uid = " & Userid)
	sql.execute("delete from friends where uid = " & userid & " or fid = " & userid)
	sql.execute("delete from forum where author = " & userid)
	sql.execute("delete from paid_email_actions where uid = " & userid)
	sql.execute("delete from contest_entries where uid = " & userid)
	mysql.execute("delete from journal_entries where uid = " & userid)
	mysql.execute("delete from thread_subs where uid = " & userid)
	set rs = mysql.execute("select id from threads where author = " & userid)
	while not rs.eof
		mysql.execute("delete from threads where parent = " & rs(0))
		rs.movenext
	wend
	rs.close
	mysql.execute("delete from threads where author = " & userid)
	sql.close
	mysql.close
	set sql = nothing
	set mysql = nothing
	set rs = nothing
	del = 1
end if
if del > 0 then
%>
<p><strong>User Removed</strong></p>
<% end if %>
      </form>
      <p>&nbsp;</p>
    <!-- InstanceEndEditable --></td><td background="../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table></td>
  </tr></table>
</body>
<!-- InstanceEnd --></html>
