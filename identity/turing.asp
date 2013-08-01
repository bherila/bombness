<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
Session("attempts") = Session("attempts") + 1
correct = true

if len(Request("response")) > 0 then
	'Determine Answer
	if lcase(Request("response")) = lcase(Session("answer")) then
		correct = true
		'Create the user
	else
		correct = false
	end if
end if

dim colors(5)
dim colornames(5)
dim answers(5)
colors(0) = "#B9FFFD" 'Teal
colornames(0) = "teal"
colors(1) = "#FFD1BB" 'Orange
colornames(1) = "orange"
colors(2) = "#FFFFD7" 'Yellow
colornames(2) = "yellow"
colors(3) = "#D6EFD1" 'Green
colornames(3) = "green"
colors(4) = "#CFCDFE" 'Blue
colornames(4) = "lavender"

Session("color1") = ""
Session("color2") = ""
Session("color3") = ""
Session("color4") = ""
Session("color5") = ""

Randomize()
first = round(rnd() * 4)
x = 1
while x < 6
	first = first + 1
	if first > 4 then first = 0
	Session("color" & x) = colors(first)
	answers(x) = colornames(first)
	x = x + 1
wend

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="../Templates/BN-IdentityPage.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>Untitled Document</title>
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
<!-- InstanceBeginEditable name="head" -->
<style type="text/css">
<!--
.style1 {
	color: #990000;
	font-weight: bold;
}
-->
</style>
<!-- InstanceEndEditable -->
</head>

<body scroll="no">

<div id="Layer1" style="position:absolute; left:29px; top:53px; width:249px; height:330px; z-index:1"><!-- InstanceBeginEditable name="Main" -->
<% if not correct then %>
<p class="style1">You entered an incorrect value. Please try again.</p>
<p>Hint: In Internet Explorer, you can select the correct value with the mouse and drag it into the box below to avoid typing errors.</p>
<% else %>
<p>Finally, to create your identity, please answer the following basic question. If you can't determine the answer, make your best guess; you will have the opportunity to answer another question if you answer this one incorrectly. </p><% end if %>
<table width="100%"  border="1" cellpadding="5" cellspacing="0" bordercolor="#333333" bgcolor="#E2E2E2">
  <tr align="center" valign="middle" bordercolor="#C7C7C7">
    <td bordercolor="#C7C7C7" bgcolor="<%= Session("color1") %>"><strong>
      <%
	
	Randomize()
	val = Replace(Hex(Round(rnd() * 60000)), "0", "P")
	Response.Write(val)
	Session("A1") = val
	
	%></strong></td>
    <td bordercolor="#C7C7C7" bgcolor="<%= Session("color2") %>"><strong>
      <%
	
	Randomize()
	val = Replace(Hex(Round(rnd() * 60000)), "0", "K")
	Response.Write(val)
	Session("A2") = val
	
	%></strong></td>
    <td bordercolor="#C7C7C7" bgcolor="<%= Session("color3") %>"><strong>
      <%
	
	Randomize()
	val = Replace(Hex(Round(rnd() * 60000)), "0", "R")
	Response.Write(val)
	Session("A3") = val
	
	%></strong></td>
    <td bordercolor="#C7C7C7" bgcolor="<%= Session("color4") %>"><strong>
      <%
	
	Randomize()
	val = Replace(Hex(Round(rnd() * 60000)), "0", "B")
	Response.Write(val)
	Session("A4") = val
	
	%></strong></td>
    <td bordercolor="#C7C7C7" bgcolor="<%= Session("color5") %>"><strong>
      <%
	
	Randomize()
	val = Replace(Hex(Round(rnd() * 60000)), "0", "J")
	Response.Write(val)
	Session("A5") = val
	
	%></strong></td>
  </tr>
</table>
<form name="form1" method="post" action="">
  <p>What value is shown in the box <%
  
  Randomize()
  if round(rnd()) = 0 then
		Randomize()
		pos = round(rnd() * 3) + 2
		select case pos
			case 1 : Response.Write("first")
			case 2 : Response.Write("second")
			case 3 : Response.Write("third")
			case 4 : Response.Write("fourth")
			case 5 : Response.Write("fifth")
		end select
  %> 
    from the <%
	
	Randomize()
	dir = round(rnd())
	if dir = 0 then
		response.Write("right")
		pos = 6 - pos
	else
		response.Write("left")
	end if
	
	%>
    ?
<%
  else
  		pos = round(rnd() * 4) + 1
		'pos = Session("color" & pos)
		Response.Write("that is tinted " & answers(pos) & "?")
  end if %>
</p>
  <p>
    <input name="response" type="text" id="response" style=" text-transform:uppercase; border: 1px solid #CCCCCC; font-family: Verdana, Arial, Helvetica, sans-serif; font-size:10pt; width: 130px;">
</p>
  <p>
    <input name="imageField" type="image" src="next.gif" width="66" height="24" border="0">
  </p>
</form>
<p>&nbsp;</p>
<!-- InstanceEndEditable --></div>
</body>
<!-- InstanceEnd --></html>
<%
Session("answer") = Session("A" & pos)
Session.Contents.Remove("A1")
Session.Contents.Remove("A2")
Session.Contents.Remove("A3")
Session.Contents.Remove("A4")
Session.Contents.Remove("A5")
%>