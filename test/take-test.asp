<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="../Templates/BN-CompactLayout.dwt" codeOutsideHTMLIsLocked="false" -->
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
          <p>Sidebar Info </p>
        <!-- InstanceEndEditable --></td>
    <td height="600" align="left" valign="top" background="../layout-2005/main.gif" bgcolor="#FFFFFF" style="background-repeat: repeat-x; padding-left: 20px; padding-right: 20px;"><!-- InstanceBeginEditable name="PageTitle" -->
    <h1>Take Test </h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
<%

if request("submit") = "Take Another Test" then
Response.Redirect("tests.asp")
end if

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open("DSN=bombness; UID=bombness; PWD=brianjef;")

dim nx
dim isCorrect
isCorrect = 0

if len(request("correct")) < 1 then
	correct = 0
else
	correct = cint(request("correct"))
end if
if len(request("incorrect")) < 1 then
	incorrect = 0
else
	incorrect = cint(request("incorrect"))
end if
if len(request("iq")) < 1 then
	iq = "0"
else
	iq = Request("iq")
end if
if len(request("iu")) < 1 then
	iu = "0"
else
	iu = Request("iu")
end if

question = Request("question")
if session("nx") = "done" then
	correct = 0
	incorrect = 0
	iq = "0"
	iu = "0"
	question = ""
end if

nx = ""
if len(session("nx")) > 0 and session("nx") <> "done" then nx = session("nx")
if len(question) > 0 then
	if instr(nx, "id <> '" & Replace(question, "'", "''") & "'") = false then
		nx = nx & " and id <> '" & Replace(question, "'", "''") & "'"
		Set Rs = Conn.Execute("select top 1 answer from test_questions where test = '" & Replace(Request("id"), "'", "''") & "' and id = '" & Replace(question, "'", "''") & "'")
		if lcase(request("answer")) = lcase(rs(0)) then
			correct = correct + 1
			isCorrect = 1
		else
			incorrect = incorrect + 1
			iq = iq & "." & question
			iu = iu & "." & Request("answer")
			isCorrect = 2
		end if
	end if
end if
session("nx") = nx

Set Rs = Conn.Execute("select count(*) from test_questions where test = '" & Replace(Request("id"), "'", "''") & "'")
tot = rs(0)
rs.close

Set Rs = Conn.Execute("select count(*) from test_questions where test = '" & Replace(Request("id"), "'", "''") & "'" & nx)
cnt = rs(0)
rs.close

if cnt > 0 then

Set Rs = Conn.Execute("select top 1 * from test_questions where test = '" & Replace(Request("id"), "'", "''") & "'" & nx & " order by newid()")
%>
<form method="post">
<p>Question <%= FormatNumber(cint(correct) + cint(incorrect) + 1, 0) %> of <%= tot %></p>
<table width="90%" border="0" cellpadding="5" cellspacing="0" style="border: 1px solid #CCCCCC;">
<% if isCorrect > 0 then %>
  <tr bgcolor="#EBE9ED">
    <td colspan="2" bgcolor="#E3EBF3">&raquo; Your previous question was answered<%
	
	if isCorrect = 1 then response.write(" correctly")
	if isCorrect = 2 then response.write(" incorrectly")
	
	%>.</td>
  </tr>
<% end if %>
  <tr bgcolor="#EBE9ED">
    <td colspan="2"><strong><input type="hidden" name="question" value="<%= Rs("id") %>"><%= Server.HTMLEncode(Rs("question")) %></strong></td>
  </tr><tr>
    <td width="20" align="center" valign="middle" bgcolor="#EBE9ED"><input name="answer" type="radio" value="A" onClick="document.getElementById('submit').disabled=false;" onBlur="document.getElementById('submit').disabled=false;" onChange="document.getElementById('submit').disabled=false;"></td>
  <td width="100%"><%= Server.HTMLEncode(Rs("a")) %></td>
  </tr>
  <tr>
    <td width="20" align="center" valign="middle" bgcolor="#EBE9ED"><input name="answer" type="radio" value="B" onClick="document.getElementById('submit').disabled=false;" onBlur="document.getElementById('submit').disabled=false;" onChange="document.getElementById('submit').disabled=false;"></td>
    <td width="100%" bgcolor="#EFEEF0"><%= Server.HTMLEncode(Rs("b")) %></td>
  </tr>
  <tr>
    <td width="20" align="center" valign="middle" bgcolor="#EBE9ED"><input name="answer" type="radio" value="C" onClick="document.getElementById('submit').disabled=false;" onBlur="document.getElementById('submit').disabled=false;" onChange="document.getElementById('submit').disabled=false;"></td>
    <td width="100%"><%= Server.HTMLEncode(Rs("c")) %></td>
  </tr>
  <tr>
    <td width="20" align="center" valign="middle" bgcolor="#EBE9ED"><input name="answer" type="radio" value="D" onClick="document.getElementById('submit').disabled=false;" onBlur="document.getElementById('submit').disabled=false;" onChange="document.getElementById('submit').disabled=false;"></td>
    <td width="100%" bgcolor="#EFEEF0"><%= Server.HTMLEncode(Rs("d")) %></td>
  </tr>
  <tr>
    <td width="20" align="center" valign="middle" bgcolor="#EBE9ED"><input name="answer" type="radio" value="E" onClick="document.getElementById('submit').disabled=false;" onBlur="document.getElementById('submit').disabled=false;" onChange="document.getElementById('submit').disabled=false;"></td>
    <td width="100%"><%= Server.HTMLEncode(Rs("e")) %></td>
  </tr>
  <tr align="left" valign="middle">
    <td colspan="2" bgcolor="#EBE9ED">
		<input type="hidden" name="iq" value="<%= iq %>">
		<input type="hidden" name="iu" value="<%= iu %>">
		<input type="hidden" name="correct" value="<%= correct %>">
		<input type="hidden" name="incorrect" value="<%= incorrect %>">
		<input type="submit" name="Submit" value="Next Question" disabled id="submit"><input type="hidden" name="id" value="<%= Request("id") %>"></td>
    </tr>
</table>
</form>
<%	
else
%>
<form method="post">
    <table width="90%" border="0" cellpadding="5" cellspacing="0" style="border: 1px solid #CCCCCC;">
      <tr bgcolor="#EBE9ED">
        <td colspan="2" bgcolor="#D3CFD8"><strong>Test Results </strong></td>
      </tr>
      <tr align="left" valign="middle">
        <td width="120" nowrap bgcolor="#ECEAEE">Questions Correct:</td>
        <td width="100%" bgcolor="#F0EEEE"><%= correct %></td>
      </tr>
      <tr align="left" valign="middle">
        <td width="120" nowrap bgcolor="#ECEAEE">Questions Wrong:</td>
        <td width="100%"><%= incorrect %></td>
      </tr>
      <tr align="left" valign="middle">
        <td width="120" nowrap bgcolor="#ECEAEE">Total Questions: </td>
        <td width="100%" bgcolor="#EFEEF0"><% 
		
		sql = "insert into test_results (tid, correct, incorrect) values ('" & Replace(Request("id"), "'", "''") & "', " & CINT(correct) & ", " & CINT(incorrect) & ")"
		Conn.Execute(sql)
		
								 total = (cint(correct) + cint(incorrect))
								 Response.Write(total)
								 
							  %></td>
      </tr>
      <tr align="left" valign="middle">
        <td width="120" nowrap bgcolor="#ECEAEE">Percent Correct: </td>
        <td width="100%"><%
		if total < 1 then total = 1 
		 pc = cdbl(cdbl(correct) / cdbl(total))
		 Response.Write(FormatPercent(pc, 1)) %></td>
      </tr>
      <tr align="left" valign="middle">
        <td nowrap bgcolor="#ECEAEE">Grade: </td>
        <td width="100%" bgcolor="#EFEEF0"><%
		
		pc = round(pc * 100)
		if pc >= 97 then
			Response.Write("A+")
		elseif pc >= 93 then
			Response.Write("A")
		elseif pc >= 90 then
			Response.Write("A-")
		elseif pc >= 87 then
			Response.Write("B+")
		elseif pc >= 83 then
			Response.Write("B")
		elseif pc >= 80 then
			Response.Write("B-")
		elseif pc >= 77 then
			Response.Write("C+")
		elseif pc >= 73 then
			Response.Write("C")
		elseif pc >= 70 then
			Response.Write("C-")
		elseif pc >= 67 then
			Response.Write("D+")
		elseif pc >= 63 then
			Response.Write("D")
		elseif pc >= 60 then
			Response.Write("D-")
		else
			Response.Write("F")
		end if
		%></td>
      </tr>
      <tr align="left" valign="middle">
        <td colspan="2" bgcolor="#D3CFD8"><strong>Incorrect Questions </strong></td>
      </tr>
      <tr align="left" valign="middle" bgcolor="#FFFFFF">
        <td colspan="2" style="padding: 0px;">
		<%

xiq = split(iq, ".")
xiu = split(iu, ".")

dim i
for i = lbound(xiq) + 1 to ubound(xiq)
Set Rs = Conn.Execute("select top 1 answer, question, a, b, c, d, e from test_questions where test = '" & Replace(Request("id"), "'", "''") & "' and id = '" & Replace(xiq(i), "'", "''") & "'")
ans = rs(0)
choice = xiu(i)

%>
        <table width="100%" border="0" cellpadding="5" cellspacing="0" style="border: 1px solid #CCCCCC;">
          <tr bgcolor="#EBE9ED">
            <td colspan="2"><strong>
              <%= Server.HTMLEncode(Rs("question")) %></strong></td>
          </tr>
          <tr>
            <td width="20" align="center" valign="middle" bgcolor="#EBE9ED"><%= a("A") %></td>
            <td width="100%"><%= Server.HTMLEncode(Rs("a")) %></td>
          </tr>
          <tr>
            <td width="20" align="center" valign="middle" bgcolor="#EBE9ED"><%= a("B") %></td>
            <td width="100%" bgcolor="#EFEEF0"><%= Server.HTMLEncode(Rs("b")) %></td>
          </tr>
          <tr>
            <td width="20" align="center" valign="middle" bgcolor="#EBE9ED"><%= a("C") %></td>
            <td width="100%"><%= Server.HTMLEncode(Rs("c")) %></td>
          </tr>
          <tr>
            <td width="20" align="center" valign="middle" bgcolor="#EBE9ED"><%= a("D") %></td>
            <td width="100%" bgcolor="#EFEEF0"><%= Server.HTMLEncode(Rs("d")) %></td>
          </tr>
          <tr>
            <td width="20" align="center" valign="middle" bgcolor="#EBE9ED"><%= a("E") %></td>
            <td width="100%"><%= Server.HTMLEncode(Rs("e")) %></td>
          </tr>
        </table>
<%
Rs.Close
next
%>
</td></tr>
      <tr align="left" valign="middle">
        <td colspan="2" bgcolor="#D3CFD8"><strong>Options</strong></td>
      </tr>
      <tr align="left" valign="middle" bgcolor="#EFEEF0">
        <td colspan="2"><input name="submit" type="submit" id="submit" value="Retake Test">
          <input name="submit" type="submit" id="submit" value="Take Another Test"></td>
      </tr>
    </table>
</form>
<%
session("nx") = "done"
end if
Conn.Close
Set Conn = Nothing

function a(answer)
	a = "&nbsp;"
	if ucase(ans) = ucase(answer) then
		a = "&raquo;"
	end if
	if ucase(answer) = ucase(choice) then
		a = "X"
	end if
end function

%>
      <!-- InstanceEndEditable --></td><td background="../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
