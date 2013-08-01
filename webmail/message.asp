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
.style1 {font-size: 10px}
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
    <td align="left" valign="top" background="../layout-2005/nav.jpg" style="padding-left: 18px;"><!-- InstanceBeginEditable name="Sidebar" -->
<%

	set cn = server.createobject("adodb.connection")
	cn.open("dsn=bombness-mysql;uid=root;pwd=eggbert;")
	set rs = cn.execute("select accountmaxsize from mail.hm_accounts where accountid = '" & Replace(session("uid"), "'", "''") & "'")
	maxsize = ccur(rs(0))
	rs.close
	set rs = cn.execute("select CAST(sum(messagesize) as CHAR) from mail.hm_messages where messageaccountid = '" & Replace(session("uid"), "'", "''") & "'")
if isNull(rs(0)) then 
	usedspace = 0
else
	usedspace = ccur(rs(0))
end if
	rs.close
	freespace = maxsize - usedspace
	

%>
    <p><a href="mailbox.asp">Read Messages</a><br>
      <a href="compose.asp">New Message</a><br>
      <a href="chpwd.asp">Change Password</a>    </p>
    <p><a href="logout.asp">Log Out</a> </p>
    <p><%= FormatNumber(freespace / 1000000, 1) %> MB remaining </p>
    <!-- InstanceEndEditable --></td>
    <td height="600" align="left" valign="top" background="../layout-2005/main.gif" bgcolor="#FFFFFF" style="background-repeat: repeat-x; padding-left: 20px; padding-right: 20px;"><!-- InstanceBeginEditable name="PageTitle" -->
      <h1>Message</h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
<%
Server.ScriptTimeout = 20
set filesys = Server.CreateObject("Scripting.FileSystemObject")
on error resume next
err.number = 0
set txt = filesys.openTextFile("C:\Program Files\hMailServer\Data\" & Request("fname"), 1, 0)
str = txt.ReadAll
pts = Split(str, vbNewLine)
txt.close
set txt = nothing
xfrom = ""
xsubject = ""
isHdr = true
msgText = ""
curType = ""
for i = 0 to uBound(pts)
	if isHdr then
		if pts(i) = "" then
			isHdr = false
		elseif lcase(left(pts(i), 24)) = "content-type: text/html;" then
			curType = "html"
		elseif lcase(left(pts(i), 25)) = "content-type: text/plain;" then
			curType = "plain"
		else
			ptt = Split(pts(i), ":", 2)
			if lcase(ptt(0)) = "subject" then
				xsubject = ptt(1)
			elseif lcase(ptt(0)) = "from" then
				xfrom = ptt(1)
			end if
			erase ptt
		end if
	else
		if lcase(left(pts(i), 16)) = "------=_nextpart" or lcase(left(pts(i), 16)) = "--90059837.11213" then
			msgText = msgText & "<hr>"
			isHdr = true
		elseif curType = "html" then
			msgText = msgText & pts(i)
		elseif curType = "plain" then
			msgText = msgText & "<br>" & pts(i)
		end if
	end if
next
if err.number <> 0 or len(xsubject) > 255 then
	xfrom = "Unknown"
	xsubject = "{An error occurred while processing this message}"
end if

  set filesys = nothing
  rs.close
  cn.close
  set rs = nothing
  set cn = nothing
  Response.Write(msgText)
  %>
    <!-- InstanceEndEditable --></td><td background="../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
