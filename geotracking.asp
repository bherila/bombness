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
    <td height="20" align="center" valign="top" background="bnBar/bg.gif"><img src="bnBar/bombness_sel.gif" width="120" height="20"><img src="bnBar/separator.gif" width="15" height="20"><a href="http://www.meteen.com" target="_top"><img src="bnBar/meteen.gif" width="45" height="20" border="0"></a><img src="bnBar/separator.gif" width="15" height="20"><a href="http://www.highschoolhumor.com" target="_top"><img src="bnBar/hsh.gif" width="97" height="20" border="0"></a><img src="bnBar/separator.gif" width="15" height="20"><a href="http://www.perfectjoke.com" target="_top"><img src="bnBar/pj.gif" width="72" height="20" border="0"></a><img src="bnBar/separator.gif" width="15" height="20"><a href="http://www.nobullgames.com" target="_top"><img src="bnBar/nbg.gif" width="84" height="20" border="0"></a></td>
  </tr><tr>
    <td align="center" valign="top">
      <table width="775" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="layout-2005/logo1.jpg" width="166" height="102"></td>
    <td><img src="layout-2005/logo2.jpg" width="601" height="102"></td>
    <td><img src="layout-2005/logo3.gif" width="8" height="102"></td>
  </tr>
  <tr>
    <td align="left" valign="top" background="layout-2005/nav.jpg" style="padding-left: 18px;"><a href="/bombness/" class="navlink">Home</a><br>
      <a href="/bombness/advertising.asp" class="navlink">Advertising</a><br>
      <a href="/bombness/sites/" class="navlink">Our Sites</a><br>
      <a href="/bombness/identities.asp" class="navlink">BN Identities</a><br>
      <a href="/bombness/investors.asp" class="navlink">Investor Relations</a><br>
      <a href="/bombness/about.asp" class="navlink">Company Information </a></td><td height="600" align="left" valign="top" background="layout-2005/main.gif" bgcolor="#FFFFFF" style="background-repeat: repeat-x; padding-left: 20px; padding-right: 20px;"><!-- InstanceBeginEditable name="PageTitle" -->
      <h1><a href="advertising.asp">Advertising</a> &raquo; Geotracking Sample </h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
      <p>This page shows you how we can detect where our visitors are located based on where their connections are coming from. Your current location is shown below. If you want to test us, try using a proxy server (you can find free anonymous proxies on your favorite search engine). Once you set up the proxy server, the location of your proxy server should be displayed instead.</p>
      <table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="100">IP Address: </td>
		  <% ip = Request.ServerVariables("REMOTE_ADDR") %>
          <td><%= ip %></td>
        </tr>
        <tr>
          <td width="100">Your Location: </td>
          <td><%
		  
		  '192.168.2.1
		  
		  pts = split(cstr(ip), ".", 4)
		  'dim pts(4)
		  'pts(0) = 192
		  'pts(1) = 168
		  'pts(2) = 2
		  'pts(3) = 1
		  A = cint(pts(0))
  		  B = cint(pts(1))
		  C = cint(pts(2))
		  D = cint(pts(3))
		  ipn = (A * 16777216) + (B * 65536) + (C * 256) + D
		  erase pts
		  
		  set conn = Server.CreateObject("Adodb.connection")
		  conn.open("dsn=bombness;uid=bombness;pwd=brianjef")
		  set rs = conn.execute("select top 1 c.country from ip2nationCountries c, ip2nation i where i.ip < " & ipn & " and c.code = i.country order by i.ip desc")
		  Response.Write(rs(0))
		  rs.close
		  set rs = nothing
		  conn.close
		  set conn = nothing		  
		  
		  %></td>
        </tr>
      </table>
      <p>&nbsp; </p>
      <!-- InstanceEndEditable --></td><td background="layout-2005/right.gif">&nbsp;</td>
  </tr>
</table></td>
  </tr></table>
</body>
<!-- InstanceEnd --></html>
