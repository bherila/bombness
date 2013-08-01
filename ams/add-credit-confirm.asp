<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bombness/includes/_id.asp"-->
<%
if Request("conf") = "yes" then

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open("DSN=bombness; UID=bombness; PWD=brianjef;")
'Get Price
	amt = cint(Request("amt"))
	cid = cint(Request("cid"))
	sql = "select top 1 price from ads where id = " & cstr(cid)
	set rs = conn.execute(sql)
	price = (rs(0) * amt)
	rs.close
	set rs = nothing
'Get Ad Format
	sql = "select top 1 [actions.max], [impressions.max] from ads where id = " & cstr(cid)
	set rs = conn.execute(sql)
	md = "inf"
	a = rs(0)
	b = rs(1)
	rs.close
	set rs = nothing
	if a > -1 then md = "actions"
	if b > -1 then md = "imps"
	if a > -1 and b > -1 then md = "both"
	if a < 0 and b < 0 then md = "inf"
'Get Balance
	sql = "select top 1 hp from users where id = '" & Replace(CurrentUserID(), "'", "''") & "'"
	set rs = conn.execute(sql)
	balance = rs(0)
	rs.close
	set rs = nothing
	if price > balance then
		conn.close
		set conn = nothing
		response.Redirect("add-credit.asp?insuf=yes&cid=" & Request("cid") & "&amt=" & Request("amt"))
	end if
'Add Credit to Campaign
	if md = "actions" then
		sql = "update ads set [actions.max] = [actions.max] + " & cstr(amt) & " where id = " & cstr(cid)
	elseif md = "imps" then
		sql = "update ads set [impressions.max] = [impressions.max] + " & cstr(amt) & " where id = " & cstr(cid)
	else
		conn.close
		set conn = nothing
		response.Redirect("add-credit.asp?unsup=yes&cid=" & Request("cid") & "&amt=" & Request("amt"))
		response.End()
	end if
	conn.execute(sql)
'Subtract Balance
	sql = "update users set hp = hp - " & cstr(price) & " where id = '" & Replace(CurrentUserID(), "'", "''") & "'"
	Conn.Execute(sql)

Conn.Close
Set Conn = Nothing
Randomize
Response.Redirect("add-credit.asp?cid=" & Request("cid") & "&updated=yes&rnd=" & rnd())

end if
%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
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
      <h1><img src="../layout/i-green.gif" width="48" height="48" hspace="10" align="absmiddle">My Campaigns &raquo; Add Value </h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
      <p>Are you sure you want to spend X HumorPoints? </p>
      <form name="form1" method="post" action="">
        <input name="imageField" type="image" src="yes.gif" width="68" height="24" border="0">
        <a href="javascript:history.back(1);"><img src="no.gif" width="68" height="24" border="0"></a> 
        <input name="conf" type="hidden" id="conf" value="yes">
        <input name="cid" type="hidden" id="cid" value="<%= Request("cid") %>">
        <input name="amt" type="hidden" id="amt" value="<%= Request("amt") %>">     
      </form>
      <p>&nbsp;</p>
      <!-- InstanceEndEditable --></td><td background="../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
