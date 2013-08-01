<!--#include virtual="/bombness/includes/_id.asp"-->
<% RequireLogin("default.asp") %>
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
          <!--#include file="nav.htm"-->
        <!-- InstanceEndEditable --></td>
    <td height="600" align="left" valign="top" background="../layout-2005/main.gif" bgcolor="#FFFFFF" style="background-repeat: repeat-x; padding-left: 20px; padding-right: 20px;"><!-- InstanceBeginEditable name="PageTitle" -->
      <h1><img src="../layout/i-green.gif" width="48" height="48" hspace="10" align="absmiddle">My Campaigns</h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
      <p>Please select one of the following available advertising formats to begin. </p>
      <form name="form1" method="post" action="create/c2.asp">
        <table width="480" border="0" cellspacing="0" cellpadding="5" style="border: 1px solid #666666;">
          <tr bgcolor="#D4D0C8">
            <td>&nbsp;</td>
            <td><strong>Type</strong></td>
            <td width="50" align="center" valign="middle"><strong>HP/CPM</strong></td>
            <td width="50" align="center" valign="middle"><strong>CPM</strong></td>
            <td width="50" align="center" valign="middle"><strong>HP/CPC</strong></td>
            <td width="50" align="center" valign="middle"><strong>CPC</strong></td>
          </tr>
          <tr>
            <td width="20" bgcolor="#D4D0C8"><input name="mxp" type="radio" value="1" checked></td>
            <td><strong>Banner Ad</strong><br>
              <em>468x60 GIF/JPEG Image </em></td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">2,500</td>
            <td align="center" valign="middle">$2.50</td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">100</td>
            <td align="center" valign="middle">$0.10</td>
          </tr>
          <tr>
            <td width="20" bgcolor="#D4D0C8"><input name="mxp" type="radio" value="2"></td>
            <td><strong>Flash Banner Ad </strong><br>
              <em>468x60 SWF </em></td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">3,000</td>
            <td align="center" valign="middle">$3.00</td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">-</td>
            <td align="center" valign="middle">-</td>
          </tr>
          <tr>
            <td width="20" bgcolor="#D4D0C8"><input name="mxp" type="radio" value="2"></td>
            <td><strong>Sponsored Link </strong><br>
              <em>20-Character Hyperlink </em></td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">900</td>
            <td align="center" valign="middle">$0.90</td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">50</td>
            <td align="center" valign="middle">$0.05</td>
          </tr>
          <tr>
            <td width="20" bgcolor="#D4D0C8"><input name="mxp" type="radio" value="3"></td>
            <td><strong>Skyscraper Particle</strong><br>
              <em>125x125 GIF/JPEG </em></td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">1,000</td>
            <td align="center" valign="middle">$1.00</td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">120</td>
            <td align="center" valign="middle">$0.12</td>
          </tr>
          <tr>
            <td bgcolor="#D4D0C8"><input name="mxp" type="radio" value="4"></td>
            <td><strong>Flash Skyscraper Particle</strong><br>
              <em>125x125 SWF </em></td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">1,500</td>
            <td align="center" valign="middle">$1.50</td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">-</td>
            <td align="center" valign="middle">-</td>
          </tr>
          <tr>
            <td bgcolor="#D4D0C8"><input name="mxp" type="radio" value="5"></td>
            <td><strong>E-Mail Campaign</strong><br>
              <em>All or Double-Opt-In Members</em> </td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">5,000</td>
            <td align="center" valign="middle">$5.00</td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">-</td>
            <td align="center" valign="middle">-</td>
          </tr>
          <tr>
            <td bgcolor="#D4D0C8"><input name="mxp" type="radio" value="6"></td>
            <td><strong>Block Ad</strong><br>
              <em>250x250 GIF/JPEG</em> </td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">4,000</td>
            <td align="center" valign="middle">$4.00</td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">100</td>
            <td align="center" valign="middle">$0.10</td>
          </tr>
          <tr>
            <td bgcolor="#D4D0C8"><input name="mxp" type="radio" value="7"></td>
            <td><strong>Flash Block Ad</strong> <br>
              <em>250x250 SWF </em></td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">4,500</td>
            <td align="center" valign="middle">$4.50</td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">-</td>
            <td align="center" valign="middle">-</td>
          </tr>
          <tr>
            <td bgcolor="#D4D0C8"><input name="mxp" type="radio" value="8"></td>
            <td><strong>IFRAME Block Ad</strong> <br>
              <em>250x250 HTML/ASP </em></td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">5,000</td>
            <td align="center" valign="middle">$5.00</td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">-</td>
            <td align="center" valign="middle">-</td>
          </tr>
          <tr>
            <td bgcolor="#D4D0C8"><input name="mxp" type="radio" value="9"></td>
            <td><strong>Layer Ad </strong><br>
              <em>350x300 GIF/JPEG </em></td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">5,000</td>
            <td align="center" valign="middle">$5.00</td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">100</td>
            <td align="center" valign="middle">$0.10</td>
          </tr>
          <tr>
            <td bgcolor="#D4D0C8"><input name="mxp" type="radio" value="10"></td>
            <td><strong>Flash Layer Ad</strong><br>
              <em>350x300 SWF </em> </td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">6,000</td>
            <td align="center" valign="middle">$6.00</td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">-</td>
            <td align="center" valign="middle">-</td>
          </tr>
          <tr>
            <td bgcolor="#D4D0C8"><input name="mxp" type="radio" value="12"></td>
            <td><strong>IFRAME Layer Ad</strong><br>
              <em>350x300 HTML/ASP </em> </td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">7,000</td>
            <td align="center" valign="middle">$7.00</td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">-</td>
            <td align="center" valign="middle">-</td>
          </tr>
          <tr>
            <td bgcolor="#D4D0C8"><input name="mxp" type="radio" value="13"></td>
            <td><strong>Movie Ad</strong></td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">10,000</td>
            <td align="center" valign="middle">$10.00</td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">-</td>
            <td align="center" valign="middle">-</td>
          </tr>
          <tr>
            <td bgcolor="#D4D0C8"><input name="mxp" type="radio" value="14"></td>
            <td><strong>Custom Solution </strong></td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">n/a</td>
            <td align="center" valign="middle">n/a</td>
            <td align="center" valign="middle" bgcolor="#ECEBE8">n/a</td>
            <td align="center" valign="middle">n/a</td>
          </tr>
          <tr>
            <td bgcolor="#D4D0C8">&nbsp;</td>
            <td colspan="5" bgcolor="#D4D0C8"><input type="submit" name="Submit" value="Continue"></td>
          </tr>
        </table>
      </form>
      <p>&nbsp; </p>
      <!-- InstanceEndEditable --></td><td background="../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
