<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%><%

Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1 

if Request("action") = "clear" then
	Session("last0") = ""
	Response.End()
end if

%><!--#include virtual="/bombness/includes/_ads.asp"-->
<html>
<head>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<title>Advertisement</title><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body scroll="no">
<%

Response.Write(ShowAd(Request("type"), Request("site")))

%>
</body>
</html>