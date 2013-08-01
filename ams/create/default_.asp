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
          <p><em>Step 1 of 5 </em></p>
          <p><strong>&raquo; Select Ad Type</strong><br>
&raquo; Design Preference<br>
            &raquo; Specify Content<br>
            &raquo; Payment<br>
            &raquo; Done
</p>
          <p>&nbsp;</p>
          <!--#include file="../nav.htm"-->
        <!-- InstanceEndEditable --></td>
    <td height="600" align="left" valign="top" background="../../layout-2005/main.gif" bgcolor="#FFFFFF" style="background-repeat: repeat-x; padding-left: 20px; padding-right: 20px;"><!-- InstanceBeginEditable name="PageTitle" -->
      <h1><img src="../../layout/i-green.gif" width="48" height="48" hspace="10" align="absmiddle">Create Advertising Campaign </h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
      <form name="form1" method="post" action="c3.asp">
        <p>Congratulations on your decision to use Bombness Networks for your advertising needs! For standard image ads, please choose a site to view available placement regions:</p>
        <ul>
          <li><a href="?site=mt">MeTeen.com</a></li>
          <li><a href="?site=hsh">HighSchoolHumor.com</a></li>
          <li><a href="?site=pj">PerfectJoke.com</a></li>
        </ul>
        <p>Click on the designated area you want your ad to appear to continue.</p>
		<% if request("site") = "hsh" then %>
        <p><img src="hsh.jpg" width="495" height="494" border="0" usemap="#Map" style="border: 1px solid black;"></p>
		<% elseif request("site") = "mt" then %>
        <p><img src="mt.jpg" width="495" height="371" border="0" usemap="#Map2" style="border: 1px solid black;"></p>
		<% elseif request("site") = "pj" then %>
        <p><img src="jjj.jpg" width="495" height="391" border="1" usemap="#Map3"></p>
		<% end if %>
          <map name="Map3">
            <area shape="rect" coords="404,87,480,373" href="?site=jjj&type=3">
            <area shape="rect" coords="248,6,484,39" href="?site=jjj&type=0">
            <area shape="rect" coords="77,322,313,355" href="?site=jjj&type=0">
          </map>          
          <map name="Map2">
            <area shape="rect" coords="199,10,487,47" href="?site=mt&type=0">
            <area shape="rect" coords="319,201,473,354" href="?site=mt&type=6">
          </map>          
          <map name="Map">
            <area shape="rect" coords="27,226,102,381" href="?site=hsh&type=2">
            <area shape="rect" coords="408,64,487,471" href="?site=hsh&type=3">
            <area shape="rect" coords="240,148,398,306" href="?site=hsh&type=6">
            <area shape="rect" coords="194,18,488,58" href="?site=hsh&type=0">
          </map>
      </form>
      <p>&nbsp;</p>
    <!-- InstanceEndEditable --></td><td background="../../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
