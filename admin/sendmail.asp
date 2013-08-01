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
          <p><strong>Send Mail</strong></p>
          <p>This script will send e-mail messages to members specified in the SQL query here. </p>
        <!-- InstanceEndEditable --></td>
    <td height="600" align="left" valign="top" background="../layout-2005/main.gif" bgcolor="#FFFFFF" style="background-repeat: repeat-x; padding-left: 20px; padding-right: 20px;"><!-- InstanceBeginEditable name="PageTitle" -->
    <h1>Send Mail </h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
      <form name="form1" method="post" action="send-2.asp">
        <p>Enter an SQL query to select e-mail addresses.</p>
        <p>
          <textarea name="sql" cols="60" rows="4" wrap="OFF" id="sql">select email, username, password from users</textarea>
        </p>
        <p>Now, select data mappings. <strong>Note: Be sure your dataset can support any defined mappings. The <em>to </em>e-mail address must always be the first column in the SQL result set. </strong></p>
        <table width="300" border="0" cellspacing="0" cellpadding="5">
          <tr bgcolor="#EBE9ED">
            <td><strong>Col 0</strong></td>
            <td nowrap>%EMAIL%</td>
          </tr>
          <tr>
            <td><strong>Col 1</strong></td>
            <td nowrap>            <input name="data1" type="text" id="data1" value="%USERNAME%"></td>
          </tr>
          <tr bgcolor="#EBE9ED">
            <td><strong>Col 2</strong></td>
            <td nowrap>              <input name="data2" type="text" id="data2" value="%PASSWORD%"></td>
          </tr>
          <tr>
            <td><strong>Col 3</strong></td>
            <td nowrap>              <input name="data3" type="text" id="data3" value="%UNDEFINED%"></td>
          </tr>
          <tr bgcolor="#EBE9ED">
            <td><strong>Col 4</strong></td>
            <td nowrap>              <input name="data4" type="text" id="data4" value="%UNDEFINED%"></td>
          </tr>
        </table>
        <p>Next, enter the subject, from address, and body of the e-mail you wish to send. You may include the variables you defined earlier to make these fields dynamic. </p>
        <p><strong>Subject<br>
          <input name="subject" type="text" id="subject">
</strong></p>
        <p><strong>From Address<br>
          <input name="from" type="text" id="from"> 
        </strong></p>
        <p><strong>Message Body<br>
          <textarea name="body" wrap="VIRTUAL" id="body" style="font-family: Arial; font-size: 10pt; width: 300px; height: 150px;"></textarea> 
        </strong></p>
        <p>Finally, click the button below to send your messages. Be sure to review your data before clicking &quot;Send,&quot; as your e-mails will be sent immediately. </p>
        <p>
          <input type="submit" name="Submit" value="Send">
        </p>
      </form>
      <p>&nbsp;</p>
    <!-- InstanceEndEditable --></td><td background="../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
