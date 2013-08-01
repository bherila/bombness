<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="../Templates/BN-IdentityPage.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>Session Information</title>
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
<!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
</head>

<body scroll="no">

<div id="Layer1" style="position:absolute; left:29px; top:53px; width:249px; height:330px; z-index:1"><!-- InstanceBeginEditable name="Main" -->
<!--#include virtual="/bombness/includes/_id.asp"-->
<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open("DSN=bombness; UID=bombness; PWD=brianjef;")
	sql = "select uid, opened, expires, ip from sessions where ip = '" & Replace(Request.ServerVariables("REMOTE_ADDR"), "'", "''") & "'"
	set rs = conn.execute(sql)
	response.Write("User:" & rs(0) & "<br>")
	response.Write("Session Start:" & rs(1) & "<br>")
	response.Write("Session Expires:" & rs(2) & "<br>")
	response.Write("Valid IP:" & rs(3) & "<br>")
	rs.close
	conn.close
	set rs = nothing
	set conn = nothing
%>
<!-- InstanceEndEditable --></div>
</body>
<!-- InstanceEnd --></html>
