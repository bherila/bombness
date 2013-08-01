<%
added = false
if len(Request("question")) > 0 then
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open("DSN=bombness; UID=bombness; PWD=brianjef;")
	Conn.Execute("insert into test_questions (test, question, answer, a, b, c, d, e) values ('" & Replace(Request("id"), "'", "''") & "', '" & Replace(Request("question"), "'", "''") & "', '" & Replace(Request("answer"), "'", "''") & "', '" & Replace(Request("a"), "'", "''") & "', '" & Replace(Request("b"), "'", "''") & "', '" & Replace(Request("c"), "'", "''") & "', '" & Replace(Request("d"), "'", "''") & "', '" & Replace(Request("e"), "'", "''") & "')")
	Conn.Close
	Set Conn = Nothing
	added = true
end if
%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
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
<!-- InstanceBeginEditable name="head" -->
<style type="text/css">
<!--
.style1 {color: #009900}
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
          <p><a href="test-questions.asp?id=<%= Request("id") %>">            Return to Test Questions</a> </p>
        <!-- InstanceEndEditable --></td>
    <td height="600" align="left" valign="top" background="../layout-2005/main.gif" bgcolor="#FFFFFF" style="background-repeat: repeat-x; padding-left: 20px; padding-right: 20px;"><!-- InstanceBeginEditable name="PageTitle" -->
    <h1>Add Question </h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
	  <% if added then %>
      <p class="style1">&quot;<%= Server.HTMLEncode(Request("question")) %>&quot; was added successfully!</p><% end if %>
      <p>Fill out this form to add your question.</p>
      <form name="form1" method="post" action="">
        <table width="500" border="0" cellspacing="0" cellpadding="5" style="border: 1px solid #CCCCCC;">
          <tr bgcolor="#EBE9ED">
            <td width="128">Question:</td>
            <td><textarea name="question" cols="50" rows="4" wrap="VIRTUAL" id="question"></textarea></td>
          </tr>
          <tr bgcolor="#F2F1F3">
            <td width="128">Choice A: </td>
            <td><input name="answer" type="radio" value="A" checked>
            <input name="a" type="text" id="a" size="50"></td>
          </tr>
          <tr bgcolor="#EBE9ED">
            <td>Choice B: </td>
            <td><input name="answer" type="radio" value="B">
            <input name="b" type="text" id="b" size="50"></td>
          </tr>
          <tr bgcolor="#F2F1F3">
            <td>Choice C: </td>
            <td><input name="answer" type="radio" value="C">
            <input name="c" type="text" id="c" size="50"></td>
          </tr>
          <tr bgcolor="#EBE9ED">
            <td>Choice D: </td>
            <td><input name="answer" type="radio" value="D">
            <input name="d" type="text" id="d" size="50"></td>
          </tr>
          <tr bgcolor="#F2F1F3">
            <td>Choice E: </td>
            <td><input name="answer" type="radio" value="E">
            <input name="e" type="text" id="e" size="50"></td>
          </tr>
          <tr bgcolor="#EBE9ED">
            <td><input name="id" type="hidden" id="id" value="<%= Request("id") %>"></td>
            <td><input type="submit" name="Submit" value="Add"></td>
          </tr>
        </table></form>
      <p>&nbsp;</p>
      <!-- InstanceEndEditable --></td><td background="../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
