<%
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1 
%><!--#include virtual="/bombness/includes/_id.asp"-->
<!--#include virtual="/bombness/includes/_global.asp"-->
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
html
{
	width: 100%;
	height: 100%;
}
body {
	background-color: #FFFFFF;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	width: 100%;
	height: 100%;
}
.style1 {	color: #666666;
}
-->
</style></head>

<body scroll="no" text="black">
<%

		Set Conn  = Server.CreateObject("ADODB.Connection")
		Set Conn2 = Server.CreateObject("ADODB.Connection")
		Conn.Open("DSN=bombness-mysql; UID=bombness; PWD=brianjef;")
		Conn2.Open("DSN=bombness-mysql; UID=bombness; PWD=brianjef;")

		Conn2.Execute("update messages set`read` = 1 where `to` = " & CurrentUserID & " and id = " & (Request("id") * 1))
		Set Rs = Conn2.Execute("select message, subj, `from`, deleted, `to`, `sent` from messages where (`to` = " & CurrentUserID & " or `from` = " & CurrentUserID & ") and id = " & (Request("id") * 1))
		
	  %><table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
  <tr align="left" bgcolor="#D4D0C8">
    <td height="16" colspan="2" align="center" valign="middle" style="border-bottom: 2px solid #666666; padding: 7px;"><table width="100%"  border="0" cellspacing="0" cellpadding="5">
        <tr class="style1">
          <td width="70" align="left" valign="middle"><strong>From:</strong></td>
          <td align="left" valign="middle"><span class="style1">
            <%
			if rs.bof and rs.eof then
				response.write("-")
			else
				msg = Rs("message")
				subj = Rs("subj")
				from = Rs("from")
				deleted = rs("deleted")
				xto = rs("to")
				Set rx = Conn.Execute("select username from users where id = " & from)
				frm = "Anonymous"
				if not (rx.eof and rx.bof) then frm = rx(0)
				Response.Write("<a href=""http://www.meteen.com/meteen/profile.asp?id=" & from & """ target=""_blank"">" & Server.HTMLEncode(frm) & "</a>")
				rx.close
				set rx = nothing
			end if
			%>
          </span></td>
        </tr>
        <tr class="style1">
          <td width="70" align="left" valign="middle"><strong>Sent:</strong></td>
          <td align="left" valign="middle"><span class="style1">
            <%
			if rs.bof and rs.eof then
				response.write("-")
			else
				Response.Write(Rs("sent"))
			end if
			%>
          </span></td>
        </tr>
        <tr class="style1">
          <td width="70" align="left" valign="middle"><strong>Subject:</strong></td>
          <td align="left" valign="middle"><span class="style1">
            <%
	if RS.EOF and RS.BOF then
		Response.Write("No Message")
	else
		Response.Write(Server.HTMLEncode(subj))
	end if
	%>
          </span></td>
        </tr>
		<% if rs.eof and rs.bof then %>
        <tr class="style1">
          <td align="left" valign="middle">&nbsp;</td>
          <td align="left" valign="middle">Delete this Message - Reply - Forward</td>
        </tr>
		<% else %>
        <tr class="style1">
          <td align="left" valign="middle">&nbsp;</td>
          <td align="left" valign="middle"><% if (not deleted) or (cint(xto) <> cint(CurrentUserID)) then %><a onClick="return confirm('Are you sure you want to permanently delete this message?');" href="del.asp?id=<%= Request("id") %>">Delete this Message</a> - <% else %>Delete this Message - <%end if %><a href="compose.asp?to=<%= Server.URLEncode(frm) %>">Reply</a> - <a href="compose.asp">Forward</a></td>
        </tr>
		<% end if %>
    </table></td>
  </tr>
  <tr>
    <td width="20" align="center" valign="middle" bgcolor="#E4E1DC"><img src="../layout/spacer.gif" width="20" height="1"></td>
    <td width="100%" height="100%" align="left" valign="top" bgcolor="#E4E1DC" style="padding: 5px;"<% if RS.EOF and RS.BOF then %><% end if %>><div style="width: 100%; height: 100%; overflow: auto;">
        <% 	if RS.EOF and RS.BOF  then
		Response.Write("&nbsp;")
	else
		Response.Write(Replace(bhtml(msg), vbNewLine, "<br>"))
	end if %>
    </div></td>
  </tr>
</table>
<%
	RS.Close
	Set RS = Nothing
	Conn.Close
	Conn2.Close
	set Conn2 = Nothing
	Set Conn = Nothing
%>
</body>
</html>
