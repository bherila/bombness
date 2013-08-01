<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bombness/includes/_id.asp"-->
<%

function contains(arr, obj)
	contains = false
	dim ijj
	for ijj = lbound(arr) to ubound(arr)
		if arr(ijj) = obj then contains = true
	next
end function

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open("DSN=bombness-mysql; UID=root; PWD=eggbert;")

if len(Request("act")) > 0 then
	conn.execute("update ads set active = 1 where sec = '" & Replace(CurrentUserID(), "'", "''") & "' and id = '" & Replace(Request("act"), "'", "''") & "'")
end if
if len(Request("deact")) > 0 then
	conn.execute("update ads set active = 0 where sec = '" & Replace(CurrentUserID(), "'", "''") & "' and id = '" & Replace(Request("deact"), "'", "''") & "'")
end if

RequireLogin("default.asp")
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
<!-- InstanceBeginEditable name="head" -->
<style type="text/css">
<!--
.style1 {
	color: #FFFFFF;
	font-weight: bold;
}
.style2 {font-size: 7pt}
-->
</style>
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
    <td align="left" valign="top" background="../layout-2005/nav.jpg" style="padding-left: 18px;"><!-- InstanceBeginEditable name="Sidebar" --><!--#include file="nav.htm"--><!-- InstanceEndEditable --></td>
    <td height="600" align="left" valign="top" background="../layout-2005/main.gif" bgcolor="#FFFFFF" style="background-repeat: repeat-x; padding-left: 20px; padding-right: 20px;"><!-- InstanceBeginEditable name="PageTitle" -->
    <h1><img src="../layout/i-green.gif" width="48" height="48" hspace="10" align="absmiddle">My Campaigns</h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
      <p><em>If you're not <%= CurrentUserName %>, <a href="logout.asp">click here</a>.</em></p>
      <table width="100%"  border="1" cellpadding="10" cellspacing="0" bordercolor="#646462" bgcolor="#DEE3ED">
        <tr>
          <td bgcolor="#EEECE8"><form name="form1" method="get" action="">
            <p>You have <%
			
			set rsx = conn.execute("select cash from users where id = '" & Replace(CurrentUserID, "'", "''") & "' limit 1")
			response.Write(formatcurrency(rsx(0), 2))
			rsx.close
			set rsx = nothing
			
			%> 
            in your account. <a href="/bombness/hp-buy.asp"></a></p>
			<%
			set rsx = conn.execute("select owner_of from users where id = '" & Replace(CurrentUserID, "'", "''") & "' limit 1")
			q = rsx(0)
			if len(q) > 0 then
			sites = split(q, ":")
			%>
            <p>You have administrative privilages for the following sites: <%
			
			qstr = "active=" & Request("active") & "&inactive=" & Request("inactive")
			str = ""
			for i = lbound(sites) to ubound(sites)
				if len(trim(sites(i))) > 0 then str = str & trim(sites(i)) & ", "
				qstr = qstr & "&site-" & sites(i) & "="
				if len(request("site-" & sites(i))) > 0 then qstr = qstr & "1"
			next
			Response.Write(left(str, len(str) - 2))
			
			%></p>
			<% end if
			rsx.close
			set rsx = nothing %>
            <p><strong>Show:</strong><br>
                <input <%If (CStr(Request("active")) = CStr("1")) Then Response.Write("checked") : Response.Write("")%> name="active" type="checkbox" id="active" value="1">
    Active Campaigns<br>
    <input <%If (CStr(Request("inactive")) = CStr("1")) Then Response.Write("checked") : Response.Write("")%> name="inactive" type="checkbox" id="inactive" value="1">
    Inactive Campaigns<%
	
	if len(q) > 0 then
	Response.Write("<br>")
		for i = lbound(sites) to ubound(sites)
			if len(trim(sites(i))) > 0 then
				Response.Write("<br><input type=""checkbox"" value=""qpb"" name=""site-" & sites(i) & """")
				if len(request("site-" & sites(i))) > 0 then Response.Write(" CHECKED")
				response.Write("> Campaigns for site " & sites(i))
			end if
		next
	end if
	
	%></p> <!--
            <p>
              <select name="view" id="view">
                <option value="1" <%If (Not isNull(Request("view"))) Then If ("1" = CStr(Request("view"))) Then Response.Write("SELECTED") : Response.Write("")%>>Summary View</option>
                <option value="2" <%If (Not isNull(Request("view"))) Then If ("2" = CStr(Request("view"))) Then Response.Write("SELECTED") : Response.Write("")%>>Campaign View</option>
              </select>
            </p> -->
            <p>      <input type="submit" name="Submit" value="Refresh" style="font-weight: bold;">
      <input type="button" name="Submit2" value="New Campaign" onClick="location.href='create';">
      <input type="button" name="Submit222" value="Banner Exchange Code" onClick="location.href='exchange.asp';">     
              </p>
          </form></td>
        </tr>
      </table>
      <p>
        <%
Function WriteType(t)
	Select Case t
		case 0 : Response.Write("Banner Ad")
		case 1 : Response.Write("Skyscraper Particle")
		case 2 : Response.Write("Skyscraper")
		case 3 : Response.Write("3")
		case 4 : Response.Write("4")
		case 5 : Response.Write("Static Affiliate Link")
	End Select
End Function

Function WriteInfoRow(rs, isActive)
	dim isExpired
	Response.Write("<a href=""javascript:void(0);"" onClick=""if (document.getElementById('_ad" & rs("id") & "').style.display != 'inline') { document.getElementById('_ad" & rs("id") & "').style.display = 'inline';} else {document.getElementById('_ad" & rs("id") & "').style.display='none';}"">Preview</a>")
	if isActive then     Response.Write(" - <a href=""deactivate.asp?id=" & rs("id") &""">Deactivate</a>")
	if not isActive and rs("active") = 0 then Response.Write(" - <a href=""?act=" & rs("id") & "&active=" & Request("active") & "&inactive=" & Request("inactive") & """>Activate</a>")
	Response.Write(" - <a href=""delete.asp?id=" & rs("id") & """>Delete</a>")
	
	Response.Write("<div id=""_ad" & rs("id") & """ name=""_ad" & rs("id") & """ style=""display: inline;""><br>")
	Response.Write(rs("html"))
	Response.Write("</div><script>document.getElementById('_ad" & rs("id") & "').style.display='none';</script>")
	
	Response.Write("</td></tr><tr><td><font color=""#888888"">")
	
	
	WriteType(rs("type"))
	Response.Write(" : ")
	Response.Write(rs("impressions"))
	if rs("impressions.max") >= 0 then
		Response.Write(" of ")
		Response.Write(rs("impressions.max"))
	end if
	Response.Write(" imps. : ")
	Response.Write(rs("actions"))
	if rs("actions.max") >= 0 then
		Response.Write(" of ")
		Response.Write(rs("actions.max"))
	end if
	Response.Write(" actions")
End Function

dim ActiveCondition
dim InactiveCondition
ActiveCondition   = "active = 1 and (((impressions < `impressions.max`) or (`impressions.max` < 0)) and ((actions < `actions.max`) or (`actions.max` < 0)))"
InactiveCondition = "active = 0 or (((impressions >= `impressions.max`) and (`impressions.max` >= 0)) or ((actions >= `actions.max`) and (`actions.max` >= 0))"


sql = "select active, impressions, `impressions.max`, actions, `actions.max`, `limit`, name, active, id, `html`, type, impressions, `impressions.max`, actions, `actions.max` from ads where (sec = " & currentuserid & ") and (" & InactiveCondition & ") order by type asc, name asc"






if cstr(request("view")) = "2" then
if len(request("inactive")) = 1 then
set rs = Conn.Execute(sql)
Do While Not Rs.EOF
	Response.Write("<table style=""border-top: 1px solid #cccccc;"" width=""90%"" border=""0"" cellpadding=""2"" cellspacing=""0""><tr><td width=""11"" rowspan=""3"" valign=""middle"" bgcolor=""#FFE8E8""><img src=""red.gif"" hspace=""5"" vspace=""5"" width=""11"" height=""11""></td><td><b>")
	Response.Write(rs("name"))
	Response.Write("</b></td></tr><tr><td>")
	WriteInfoRow rs, false
	Response.Write("</font></td></tr></table>")
	rs.MoveNext
Loop
rs.close
end if

if len(request("active")) = 1 then
sql = Replace(sql, Inactivecondition, Activecondition)
set rs = Conn.Execute(sql)
Do While Not Rs.EOF
	Response.Write("<table style=""border-top: 1px solid #cccccc;"" width=""90%"" border=""0"" cellpadding=""2"" cellspacing=""0""><tr><td width=""11"" rowspan=""3"" valign=""middle"" bgcolor=""#E1FFE1""><img src=""green.gif"" hspace=""5"" vspace=""5"" width=""11"" height=""11""></td><td><b>")
	Response.Write(rs("name"))
	Response.Write("</b></td></tr><tr><td>")
	WriteInfoRow rs, true
	Response.Write("</font></td></tr></table>")
	rs.MoveNext
Loop
%></p><%
end if
else
%>
      <table width="97%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td bgcolor="#E0DDE3" style="border: 1px solid #BBBBBB; padding: 0px;"><table width="100%"  border="0" cellspacing="0" cellpadding="5">
<%
function writeHeader(tipe, editable, editID)
%>
              <tr bgcolor="#CBC6D0" style="border-bottom: 1px solid #CCCCCC;">
                <td colspan="9" bgcolor="#68718C"><span class="style1"><%= tipe %><% if editable then %>| <small><a href="admin-create.asp?id=<%= editID %>&<%=qstr%>" class="style1"><font color="#FFFFFF">New Campaign</font></a></small><% end if %></span></td>
              </tr>
              <tr valign="middle" bgcolor="#CBC6D0" style="border-bottom: 1px solid #CCCCCC;">
                <td width="15" height="30" bgcolor="#68718C">&nbsp;</td>
                <td width="60" height="30" align="left" bgcolor="#BFC7D7"><strong>State</strong></td>
                <td width="35" height="30" align="center" bgcolor="#BFC7D7"><strong>Site</strong></td>
                <td height="30" bgcolor="#BFC7D7"><strong>Type</strong></td>
                <td height="30" bgcolor="#BFC7D7"><strong>Campaign Name </strong></td>
                <td width="60" height="30" bgcolor="#BFC7D7"><strong>Imps</strong></td>
                <td width="60" height="30" bgcolor="#BFC7D7"><strong>Clicks</strong></td>
                <td width="60" height="30" bgcolor="#BFC7D7"><strong>CTR</strong></td>
                <td width="60" height="30" align="center" bgcolor="#BFC7D7"><strong>Extend</strong></td>
              </tr>
<%
end function

function doStuff(tsql, ext)
	set rs = Conn.Execute(tsql)
	total_imps = 0
	total_actions = 0
	i = 0
	While not rs.eof
	i = i + 1 
					%>
				  <tr valign="middle"<%
					
					if i / 2 = round(i / 2) then Response.Write(" bgcolor=""#FFFFFF""")
					
					%>>
					<td width="15" height="30" bgcolor="#68718C">&nbsp;</td>
					<td width="60" height="30" align="left" bgcolor="<%
					st = 0
					imps = rs("impressions")
					imps_max = rs("impressions.max")
					actions = rs("actions")
					actions_max = rs("actions.max")
					
					if rs("active") = 0 then
						st = 2
					else
						if ((imps >= imps_max) and (imps_max >= 0)) or ((actions >= actions_max) and (actions_max >= 0)) then
							st = 3
						else
							st = 1
						end if
					end if
					
					if st = 1 then
						response.Write("#E6FFBF")
					elseif st = 2 then
						response.Write("#FFCC00")
					else
						response.Write("#990000")
					end if
					
					%>"><%
					
					session("qstr") = request.QueryString()
					id = rs("id")
					if st = 1 then
						response.Write("<a onClick=""return confirm('Are you sure you want to pause this ad campaign?');"" href=""?deact=" & rs("id") & "&" & qstr & """>Pause</a>")
					elseif st = 2 then
						response.Write("<a href=""?act=" & id & "&" & QStr & """>Resume</a>")
					else
						response.Write("<font color=""#FFFFFF"">Expired</font>")
					end if
					
					if ext then
					%><br>
					  <span class="style2"><a href="admin-edit.asp?id=<%= id %>&<%= qstr %>">Edit</a><br>
					  <a href="admin-delete.asp?id=<%= id %>"&<%= qstr %>>Delete</a></span>
					  
					  <% end if %>
					  
					</td>
					<td width="35" height="30" align="center"><%= Server.HTMLEncode(Rs("limit")) %></td>
					<td height="30"><%
					
					t = "INF"
					if imps_max > -1 then t = "CPM"
					if actions_max > -1 then t = "CPA"
					if imps_max > -1 and actions_max > -1 then t = "HYB"
					if impx_max < 0 and actions_max < 0 then t = "INF"
					Response.Write(t)
					
					%></td>
					<td height="30"><%= Server.HTMLEncode(Rs("name")) %></td>
					<td width="60" height="30"><%
					
					total_imps = total_imps + imps
					Response.Write FormatNumber(imps, 0)
					if imps_max > -1 then
					response.Write(" of " & formatnumber(imps_max, 0))
					end if
					
					%></td>
					<td width="60" height="30"><%
					
					total_actions = total_actions + actions
					Response.Write(FormatNumber(actions, 0))
					if actions_max > -1 then
					response.Write (" of " & formatnumber(actions_max, 0))
					end if
					
					%></td>
					<td width="60" height="30"><%
					
					if imps > 0 then
					response.Write formatpercent(actions/imps, 2)
					else
					response.Write("Undef")
					end if
					
					%></td>
				    <td width="60" height="30" align="center" style="padding:0px;"><% if (t = "CPA" or t = "CPC" or t = "CPM") then %><a href="add-credit.asp?active=<%= Request("active") %>&inactive=<%= Request("inactive") %>&view=<%= Request("view") %>&cid=<%= id %>"><% end if %><img src="<% if not (t = "CPA" or t = "CPC" or t = "CPM") then response.Write("no_") %>extend.gif" width="60" height="26" border="0"><% if (t = "CPA" or t = "CPC" or t = "CPM") then %></a><% end if %></td>
				  </tr>
				  <%
	rs.movenext
	wend
	rs.close
	%>
              <tr bgcolor="#BFC7D7">
                <td width="15" bgcolor="#68718C">&nbsp;</td>
                <td colspan="4"><strong>Total</strong></td>
                <td><b><%= FormatNumber(total_imps, 0) %></b></td>
                <td><b><%= FormatNumber(total_actions, 0) %></b></td>
                <td><b><%  if total_imps > 0 then 
						      Response.Write(FormatPercent(total_actions / total_imps, 2))
						   else
						      Response.Write("Undef")
						   end if
						%></b></td>
                <td align="center">&nbsp;</td>
              </tr>
	<%
end function

if len(request("inactive")) = 1 then
	writeHeader "Inactive Campaigns", false, ""
	doStuff sql, false
end if
Response.Flush()
if len(request("active")) = 1 then
	writeHeader "Active Campaigns", false, ""
	sql = Replace(sql, Inactivecondition, Activecondition)
	doStuff sql, false
end if
Response.Flush()
if len(q) > 0 then
Response.Write("<br>")
	max = ubound(sites)
	if max > 10 then max = 10
	for ixp = 0 to max
		if len(trim(sites(ixp))) > 0 then
		if len(request("site-" & trim(sites(ixp)))) > 0 then
			writeHeader "Campaigns for Site " & sites(ixp), true, sites(ixp)
			sql = "select active, impressions, `impressions.max`, actions, `actions.max`, `limit`, name, active, id, `html`, type, impressions, `impressions.max`, actions, `actions.max` from ads where `limit` = '" & replace(sites(ixp), "'", "''") & "'"			
			doStuff sql, true
		end if
		end if
		response.Flush()
	next
end if

%>
          </table></td>
        </tr>
      </table>
<%
end if 
%>
<br>
<!-- InstanceEndEditable --></td><td background="../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>