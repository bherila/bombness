<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bombness/includes/_id.asp"-->
<%
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open("DSN=bombness; UID=bombness; PWD=brianjef;")

sql = "select top 1 sec from ads where id = '" & Replace(Request("id"), "'", "''") & "'"
Set RS = Conn.Execute(sql)

if rs(0) <> CurrentUserID then
	Conn.Close
	Set Conn = Nothing
	Response.Redirect("denied.htm")
else
	if Request("action") = "delete" then
		sql = "delete from ads where id = '" & Replace(Request("id"), "'", "''") & "'"
		Conn.Execute(sql)
		Conn.Close
		Set Conn = Nothing
		Response.Redirect("campaigns.asp")
	end if
end if
%>

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
          <!--#include file="nav.htm"-->
        <!-- InstanceEndEditable --></td>
    <td height="600" align="left" valign="top" background="../layout-2005/main.gif" bgcolor="#FFFFFF" style="background-repeat: repeat-x; padding-left: 20px; padding-right: 20px;"><!-- InstanceBeginEditable name="PageTitle" -->
    <h1><img src="../layout/i-green.gif" width="48" height="48" hspace="10" align="absmiddle">Delete Ad Campaign </h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
      <p>Are you really sure you want to delete this ad campaign? Doing so will immediately halt its progress. Please note that we do not refund partially completed campaigns. </p>
      <form name="form1" method="get" action="">
        <input name="imageField" type="image" src="yes.gif" alt="Yes" width="68" height="24" border="0">
        <a href="javascript:history.back(1);"><img src="no.gif" alt="No" width="68" height="24" border="0"></a> 
        <input name="action" type="hidden" id="action" value="delete"> 
        <input name="id" type="hidden" id="id" value="<%= Request("id") %>">            
      </form>
    <!-- InstanceEndEditable --></td><td background="../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
