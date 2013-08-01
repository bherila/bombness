<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bombness/includes/_id.asp"-->
<%
on error resume next
err.number = 0
dim i
i = cdbl(Request("amt"))
en = err.number
on error goto 0
if en = 0 then
if Request("conf") = "yes" and (i = round(i)) and i > 0 then

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
	sql = "select top 1 cash from users where id = '" & Replace(CurrentUserID(), "'", "''") & "'"
	set rs = conn.execute(sql)
	balance = rs(0)
	rs.close
	set rs = nothing
	if price > balance then
		conn.close
		set conn = nothing
		response.Redirect("add-credit.asp?active=" & Request("active") & "&inactive=" & Request("inactive") & "&view=" & Request("view") & "&insuf=yes&cid=" & Request("cid") & "&amt=" & Request("amt"))
	end if
'Add Credit to Campaign
	if md = "actions" then
		sql = "update ads set [actions.max] = [actions.max] + " & cstr(amt) & " where id = " & cstr(cid)
	elseif md = "imps" then
		sql = "update ads set [impressions.max] = [impressions.max] + " & cstr(amt) & " where id = " & cstr(cid)
	else
		conn.close
		set conn = nothing
		response.Redirect("add-credit.asp?active=" & Request("active") & "&inactive=" & Request("inactive") & "&view=" & Request("view") & "&unsup=yes&cid=" & Request("cid") & "&amt=" & Request("amt"))
		response.End()
	end if
	conn.execute(sql)
'Subtract Balance
	sql = "update users set cash = cash - " & cstr(price) & " where id = '" & Replace(CurrentUserID(), "'", "''") & "'"
	Conn.Execute(sql)

Conn.Close
Set Conn = Nothing
Randomize
Response.Redirect("add-credit.asp?active=" & Request("active") & "&inactive=" & Request("inactive") & "&view=" & Request("view") & "&cid=" & Request("cid") & "&updated=yes&rnd=" & rnd())

end if
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
<!-- InstanceBeginEditable name="head" -->
<%
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open("DSN=bombness; UID=bombness; PWD=brianjef;")
sql = "select top 1 price, [impressions], [impressions.max], [actions], [actions.max] from ads where id = " & cstr(cint(Request("cid")))
set rs = Conn.execute(sql)
	unit_price = rs(0)
	imps = rs(1)
	impsmax = rs(2)
	acts = rs(3)
	actsmax = rs(4)
	rs.close
sql = "select top 1 cash from users where id = '" & Replace(CurrentUserID, "'", "''") & "'"
set rs = conn.execute(sql)
	hp_bal = rs(0)
	rs.close
	set rs = nothing
t = "inf"	
if impsmax > -1 then t = "imps"
if actsmax > -1 then t = "acts"
if impsmax > -1 and actsmax > -1 then t = "both"
if impsmax < 0 and actsmax < 0 then t = "inf"
if t = "imps" then
	ad_bal = impsmax
elseif t = "acts" then
	ad_bal = actsmax
end if
conn.close
set conn = nothing
if t <> "imps" and t <> "acts" then
	Response.Clear()
	Response.Write("<scr" + "ipt>alert('Please contact us to edit this campaign, as it is not supported by our online editor.'); history.back(1);</scr" & "ipt>")
	Response.End()
end if
%>
<script language="javascript">
var ad_bal = <%= ad_bal %>;
var hp_bal = <%= hp_bal %>;
var unit_price = <%= unit_price %>;
var total_price = 0;
var units = <%

	on error resume next
	err.number = 0
	amt = 0
	amt = cint(amt) 	
	if len(request("amt")) > 0 and err.number <> 0 then
		Response.Write(amt)
	else
		Response.Write("0")
	end if
	on error goto 0
	
%>;

function d()
{
	units = document.getElementById('amt').value;
	total_price = units * unit_price;
	document.getElementById('a2').innerHTML = (hp_bal * 1) - (total_price * 1);
	document.getElementById('a1').innerHTML = (ad_bal * 1) + (units * 1);
	document.getElementById('d1').innerHTML = (units * 1);
	document.getElementById('d2').innerHTML = (-1 * (total_price * 1));
}
</script>
<!-- InstanceEndEditable -->
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
      <p>Increase the life of your ad campaign by adding more value to it. The table below will update with your input. </p>
	  <form action="add-credit.asp" method="get">
      <table  border="0" cellpadding="10" cellspacing="0" bordercolor="#A7A6AA" style="border: 1px solid gray;">
<% if request("updated") = "yes" then %>        <tr bgcolor="#E8FFEA">
          <td colspan="4" style="border-bottom: 1px solid #A6B3CE;"><img src="green.gif" width="11" height="11" hspace="10">Your campaign has been updated.</td>
        </tr><% end if 
		
		if len(request("action")) > 0 then
		
		on error resume next
		err.number = 0
		i = cdbl(Request("amt"))
		if err.number <> 0 or (i <> round(i)) then %>
        <tr bgcolor="#FFE8E8">
          <td colspan="4" style="border-bottom: 1px solid #A6B3CE;"><img src="red.gif" width="11" height="11" hspace="10">The value you entered is not valid. </td>
        </tr><%
		else
			if i < 1 then %>
        <tr bgcolor="#FFE8E8">
          <td colspan="4" style="border-bottom: 1px solid #A6B3CE;"><img src="red.gif" width="11" height="11" hspace="10">You must enter a value greater than zero. </td>
        </tr><%
			end if
		
		end if
		err.number = 0
		on error goto 0
		
		end if
		
		if request("insuf") = "yes" then
		%>
        <tr bgcolor="#FFE8E8">
          <td colspan="4" style="border-bottom: 1px solid #A6B3CE;"><img src="red.gif" width="11" height="11" hspace="10">You do not have enough points to complete this transaction. </td>
        </tr><% end if
		
		if request("unsup") = "yes" then %>
		<tr bgcolor="#FFE8E8">
          <td colspan="4" style="border-bottom: 1px solid #A6B3CE;"><img src="red.gif" width="11" height="11" hspace="10">The transaction failed because this campaign does not support automatic updates. </td>
        </tr><% end if %>
        <tr bgcolor="#DEE3ED">
          <td width="190" style="border-right: 1px solid #A6B3CE; border-bottom: 1px solid #D6DCE9;"><strong>Description</strong></td>
          <td width="130" style="border-right: 1px solid #A6B3CE; border-bottom: 1px solid #D6DCE9;"><strong>Current Balance </strong></td>
          <td style="border-right: 1px solid #A6B3CE; border-bottom: 1px solid #D6DCE9;"><strong>Difference</strong></td>
          <td width="190" style="border-bottom: 1px solid #D6DCE9;"><strong>Balance After Transaction </strong></td>
        </tr>
        <tr>
          <td bgcolor="#F2F4F9" style="border-right: 1px solid #A6B3CE;">Campaign <%
		  
		  if t = "imps" then Response.Write("Impressions")
		  if t = "acts" then Response.Write("Actions")
		  
		  %></td>
          <td align="right" style="border-right: 1px solid #A6B3CE;"><script>document.write(ad_bal);</script></td>
          <td align="right" style="border-right: 1px solid #A6B3CE;"><div id="d1"></div></td><td align="right"><div id="a1"></div></td><!-- Impressions Remaining -->
        </tr>
        <tr>
          <td bgcolor="#F2F4F9" style="border-right: 1px solid #A6B3CE;">My Account</td>
          <td align="right" style="border-right: 1px solid #A6B3CE;"><script>document.write(hp_bal);</script></td>
          <td align="right" style="border-right: 1px solid #A6B3CE;"><div id="d2"></div></td><td align="right"><div id="a2"></div></td><!-- Amt after Transaction -->
        </tr>
        <tr bgcolor="#E9ECF3">
          <td bgcolor="#DEE3ED" style="border-top: 1px solid #A6B3CE;"><strong>Amount to Add</strong></td>
          <td colspan="2" bgcolor="#DEE3ED" style="border-top: 1px solid #A6B3CE;"><input onKeyUp="d();" onChange="d();" onBlur="d();" name="amt" type="text" id="amt" style="border: 1px solid #A7A6AA; width: 80px; font-family:Arial, Helvetica, sans-serif; font-size: 10pt; font-weight:bold; text-align:right;" value="<%= Request("amt") %>"> 
            <strong><%
		  
		  if t = "imps" then Response.Write("impressions")
		  if t = "acts" then Response.Write("actions")
		  
		  %></strong></td>
          <td bgcolor="#DEE3ED" style="border-top: 1px solid #A6B3CE;"><input type="submit" name="Submit" value="Finish">
            <input name="view" type="hidden" id="view" value="<%= Request("view") %>">            
            <input name="inactive" type="hidden" id="inactive" value="<%= Request("inactive") %>">            
            <input name="active" type="hidden" id="active" value="<%= Request("active") %>">            
            <input type="hidden" name="action" value="add-credit">
			<input type="hidden" name="conf" value="yes">
            <input name="cid" type="hidden" id="cid" value="<%= Request("cid") %>"></td>
        </tr>
      </table>
	  <p>
	    <input type="button" name="Submit2" value="Return to Campaign List" onClick="location.href='campaigns.asp?active=<%= Request("active") %>&inactive=<%= Request("inactive") %>&view=<%= Request("view") %>';">
	  </p>
	  </form>
	  <script>
	  d();
	  </script>
    <!-- InstanceEndEditable --></td><td background="../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
