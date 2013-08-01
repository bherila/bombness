<!--#include virtual="/bombness/includes/_id.asp"-->
<!--#include virtual="/bombness/includes/_global.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/BN-CompactLayout.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>MeTeen.com - Free Profiles, Subprofiles, Blogs, and More! - Start Meeting</title>
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
<script language="javascript">
function p(x, y)
{
document.getElementById('prg').innerHTML = '<strong>' + x + ' of ' + y + ' sent (' + Math.round(x/y*100) + '%)</strong>';
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
          <p><strong>Working</strong></p>
          <p>Please wait while Bombness Networks performs the requested actions. This may take several minutes. </p>
        <!-- InstanceEndEditable --></td>
    <td height="600" align="left" valign="top" background="../layout-2005/main.gif" bgcolor="#FFFFFF" style="background-repeat: repeat-x; padding-left: 20px; padding-right: 20px;"><!-- InstanceBeginEditable name="PageTitle" -->
      <h1>Working</h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
      <p>Please wait while Bombness Networks sends your e-mails. Warning! Do not interrupt this page or some of your messages may not be sent.</p>
      <div id="prg"><strong>Preparing to send...</strong></div>
      <p></p>
      <!-- InstanceEndEditable --></td><td background="../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
<%

on error resume next

sql = request("sql")
ok = true
if inStr(sql, "truncate table") then ok = false
if inStr(sql, "delete from") then ok = false
if inStr(sql, "update table") then ok = false

dim vars(5)
dim data(5)

for z = 1 to 4
	vars(z) = Request("data" & cstr(z))
next

function format(str, rt)
	for z = 1 to 4
		if len(request("data" & z)) > 0 and uCase(Request("data" & z)) <> "%UNDEFINED%" then str = Replace(str, vars(z), data(z))
	next
	str = Replace(str, "%EMAIL%", rt)
	format = str
end function

if ok then
	Response.Flush
	set conn = server.CreateObject("adodb.connection")
	conn.open("dsn=bombness;uid=bombness;pwd=brianjef")
	
	i = 0
	total = 0
	
	set rs = conn.execute("select count(*) from (" & sql & ") derivedtbl")
	total = rs(0)
	rs.close
	
	set rs = conn.execute(sql)
	while not rs.eof

		Server.ScriptTimeout = Server.ScriptTimeout + 1 
		xxTo = rs(0)
		for z = 1 to 4
			if len(request("data" & z)) > 0 and uCase(Request("data" & z)) <> "%UNDEFINED%" then data(z) = rs(z)
		next
		xxFrom = format(Request("from"), xxTo)
		xxSubj = format(Request("subject"), xxTo)
		xxBody = format(Request("body") & vbnewline & vbnewline & "=========================" & vbnewline & "You are receiving this e-mail because you are a member of Bombness Networks and have agreed to receive such e-mail when you subscribed. If you would not like to receive these mailings any more, please contact bherila@bombness.com.", xxTo)
		
		sendmail2 conn, xxTo, xxFrom, xxSubj, xxBody
		
		i = i + 1
		if i / 100 = i \ 100 then
			Response.Write("<scr" & "ipt>p(" & cstr(i) & ", " & total & ");</scr" & "ipt>")
		end if
		Response.Flush()
		rs.movenext
	wend
	rs.close
	set rs = nothing
	conn.close
	set conn = nothing
end if

Response.Write("<scr" & "ipt>location.replace('done.asp');</scr" & "ipt>")
%>