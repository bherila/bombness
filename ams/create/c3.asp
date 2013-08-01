<!--#include virtual="/bombness/includes/_id.asp"-->
<% RequireLogin("default.asp") %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="../../Templates/BN-CompactLayout.dwt" codeOutsideHTMLIsLocked="false" -->
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
    <td><img src="../../layout-2005/logo1.jpg" width="166" height="102"></td>
    <td><img src="../../layout-2005/logo2.jpg" width="601" height="102"></td>
    <td><img src="../../layout-2005/logo3.gif" width="8" height="102"></td>
  </tr>
  <tr>
    <td align="left" valign="top" background="../../layout-2005/nav.jpg" style="padding-left: 18px;"><!-- InstanceBeginEditable name="Sidebar" -->
          <p><em>Step 3 of 5 </em></p>
          <p>&raquo; Select Ad Type<br>
&raquo; Design Preference<br>
<strong>&raquo; Specify Content</strong><br>
&raquo; Payment<br>
&raquo; Done </p>
          <p>
            <!--#include file="../nav.htm"-->
          </p>
        <!-- InstanceEndEditable --></td>
    <td height="600" align="left" valign="top" background="../../layout-2005/main.gif" bgcolor="#FFFFFF" style="background-repeat: repeat-x; padding-left: 20px; padding-right: 20px;"><!-- InstanceBeginEditable name="PageTitle" -->
      <h1><img src="../../layout/i-green.gif" width="48" height="48" hspace="10" align="absmiddle">Create Advertising Campaign </h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
      <form action="" method="post" enctype="multipart/form-data" name="form1">
<% if request("design") = "yes" then %>
        <p>Please describe the ad you want us to create in as much detail as possible.<br>
        Be sure to include any important details such as a link to your corporate<br>
        logo and specific slogans and wording relevent to your ad.</p>
        <p><textarea name="textarea" cols="70" rows="9" wrap="VIRTUAL"></textarea></p>
<% end if %>
<% if request("design") = "no" then %>
        <p>Select a creative to upload. We will look at each one and approve<br>
        it. Your ad will not go live until the creative has been approved.</p>
        <p><em>Note: There is a maximum file size of 0 KB for this type of ad. </em></p>
        <p>File: <input name="file" type="file" size="50">
          <input name="maxsize" type="hidden" id="maxsize" value="0">
</p>
<% end if %>
        <p>
          <input type="submit" name="Submit" value="Submit">
          <input name="mxp" type="hidden" id="mxp" value="<%= Request("mxp") %>">
          <input name="design" type="hidden" id="design" value="<%= Request("design") %>">
        </p>
      </form>
      <p>&nbsp;</p>
    <!-- InstanceEndEditable --></td><td background="../../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
