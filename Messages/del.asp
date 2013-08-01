<%
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1 
%><!--#include virtual="/bombness/includes/_id.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<script language="javascript">
//document.domain = "bombness.com";
</script>
<title>Untitled Document</title>
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
-->
</style></head>
<%
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open("DSN=bombness-mysql; UID=bombness; PWD=brianjef;")
sql = "update messages set deleted = 1 where id = " & cint(Request("id") * 1) & " and `to` = " & CurrentUserID
Conn.Execute(sql)
Conn.Close
Set Conn = Nothing
%>
<body scroll="no">
<table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="0">
  <tr>
    <td align="center" valign="middle" bgcolor="#D4D0C8">Message deleted</td>
  </tr>
</table>
<script>
parent.frames[0].location.reload();
</script>
</body>
</html>
