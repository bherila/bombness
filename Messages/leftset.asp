<%
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1 
%><!--#include virtual="/bombness/includes/_id.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Messages</title>
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
}
.style1 {
	color: #666666;
	font-weight: bold;
}
-->
</style>

<script language="javascript">
//document.domain = "bombness.com";

var SelectedMessage = null;
function SelectMessage(target, id)
{
if (SelectedMessage != null)
{
	SelectedMessage.style.backgroundColor = '#FFFFFF';
	SelectedMessage.style.color = '#AAAAAA';
}
SelectedMessage = target;
SelectedMessage.style.backgroundColor = '#0066CC';
SelectedMessage.style.color = '#FFFFFF';
var newurl = 'message.asp?id=' + id;
if (parent.frames[1].location.href != newurl)
{
	parent.frames[1].location.href = newurl;
}
//parent.getElementById('mainFrame').location.href=newurl;

}

</script>

</head>
<body text="#000000">
<table width="100%"  border="0" cellpadding="0" cellspacing="0">
<% function doFolder(folderName, sqlQuery) %>
  <tr align="left" bgcolor="#D4D0C8">
    <td height="16" align="center" valign="middle" style="border-bottom: 2px solid #666666; padding: 7px;"><strong><img src="minus.gif" width="9" height="9" align="absmiddle"></strong></td>
    <td height="16" valign="middle" style="border-bottom: 2px solid #666666; padding: 5px;"><span class="style1"><%= folderName %></span></td>
  </tr>
  <tr>
    <td width="20" align="center" valign="middle" bgcolor="#E4E1DC"></td>
    <td>
	  <%
	  	on error resume next
	    Response.Flush()
		Set Conn  = Server.CreateObject("ADODB.Connection")
		Set Conn2 = Server.CreateObject("ADODB.Connection")
		Conn.Open("DSN=bombness-mysql; UID=bombness; PWD=brianjef;")
		Conn2.Open("DSN=bombness-mysql; UID=bombness; PWD=brianjef;")

		Set Rs = Conn2.Execute(sqlQuery)
	    While Not Rs.EOF
		
	  %>	  <div id="msg0" style="color: <% if rs("read") = 0 then %>#000000<% else %>#AAAAAA<% end if %>;cursor: hand; background-color: #FFFFFF; border:none; width: 100%;" onClick="SelectMessage(this, <%= RS("id") %>);">
	  	<table style="border-bottom: 1px solid #CCCCCC;" width="100%"  border="0" cellspacing="0" cellpadding="5">
	      <tr>
	        <td><strong><%
			
			Set Rs2 = Conn.Execute("select username from users where id = " & RS("from"))
			if not (rs2.eof and rs2.bof) then
				Response.Write(Server.HTMLEncode(Rs2(0)))
			else
				Response.Write("Anonymous")
			end if
			Rs2.Close
			
			%></strong></td>
	        <td align="right"><%= Server.HTMLEncode(RS("sent")) %></td>
	      </tr>
	      <tr>
	        <td colspan="2"><%= Server.HTMLEncode(Rs("subj")) %></td>
	        </tr>
	    </table>
	  </div>
	  <%
	    Response.Flush()
	  	RS.MoveNext
	  Wend
	  RS.Close
	  Set RS = Nothing
	Conn.Close
	Conn2.Close
	set Conn2 = Nothing
	Set Conn = Nothing
	  
	  %>	</td>
  </tr>
  <% end function %>
  <% doFolder "Inbox", ("select id, `read`, `from`, `sent`, `subj` from messages where deleted = 0 and `to` = " & CurrentUserID & " order by id desc") %>
  <% 'doFolder "Sent Items", ("select id, `read`, `from`, `sent`, `subj` from messages where `from` = " & CurrentUserID & " order by id desc") %>
  <% 'doFolder "Deleted Items", ("select id, `read`, `from`, `sent`, `subj` from messages where deleted = 1 and `to` = " & CurrentUserID & " order by id desc") %>
</table>
</body>
</html>
