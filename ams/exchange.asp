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
<!-- InstanceBeginEditable name="head" --><style>
.selTab
{
 background-color: #D2D1C6;
 border: 1px solid #666666;
 border-bottom: none;
 margin-top: 4px;
 margin-right: 0px;
 margin-left: 2px;
 padding: 3px;
 font-size:12px;
 font-family:Arial, Helvetica, sans-serif;
 height: 30px;
 vertical-align:middle;
 text-align:center;
}
.style1 {
	font-size: 15pt;
	font-weight: bold;
}
</style>
<script language="javascript">
var selTab = null;
function selectTab(obj, target)
	{
	var cnt = document.getElementById(target);
	var tab = obj;
	hideAll();
	cnt.style.display='block';
	}
function setSel(tab)
	{
	if (selTab != null)
		{
		selTab.style.marginBottom='1px';
		selTab.style.borderBottom='1px solid #666666';
		tab.style.marginBottom='0px';
		tab.style.borderBottom = '0px none';
		selTab = tab;
		}
	}
function hideAll()
	{
	var i;
	for (i=0; i<5; i++)
		{
		document.getElementById('content' + i).style.display='none';
		}
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
          <!--#include file="nav.htm"-->
        <!-- InstanceEndEditable --></td>
    <td height="600" align="left" valign="top" background="../layout-2005/main.gif" bgcolor="#FFFFFF" style="background-repeat: repeat-x; padding-left: 20px; padding-right: 20px;"><!-- InstanceBeginEditable name="PageTitle" -->
      <h1><img src="../layout/i-green.gif" width="48" height="48" hspace="10" align="absmiddle">Banner Exchange </h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
      <p>Display our banners on your web site and get credit that you can apply to your own ad campaigns! </p>
		<table width="100%" height="300" border="0" cellpadding="5" cellspacing="0">
		  <tr>
			<td style="padding: 0px; border-bottom: 1px solid #666666;" height="19" valign="top"><span id="tab4" class="selTab" onClick="selectTab(this, 'content4');">Text Links </span><span id="tab0" class="selTab" onClick="selectTab(this, 'content0');">468x60</span><span class="selTab" onClick="selectTab(this, 'content1');">125x125</span><span class="selTab" onClick="selectTab(this, 'content2');">125x600</span><span class="selTab" onClick="selectTab(this, 'content3');">Popup Ads</span></td>
		  </tr>
		  <tr>
			<td align="left" valign="top" bgcolor="#D2D1C6" style="border-left: 1px solid #666666; border-right: 1px solid #666666; border-bottom: 1px solid #666666; padding: 15px;">
				<div id="content4">
				  <p class="style1">Text Link  Ads </p>
				  <table width="100%"  border="0" cellspacing="0" cellpadding="5">
                    <tr>
                      <td style="border: 1px solid gray; background-color: white;"><a href="http://www.meteen.com">Meet Hot Teens</a> </td>
                    </tr>
                    <tr>
                      <td><ul>
                        <li>$0.02 per click </li>
                        <li>Only unique clicks (per IP address) within 12 hour period are valid </li>
                        <li>$0.50 bonus will be instantly credited if user signs up within 12 hours of clickthrough (from same IP address)</li>
                        </ul>
                      <p>Copy and Paste the following HTML code into your web page or ad rotation script. If you are an experienced HTML user, you may modify the link text as you wish, but remember, you may not be deceptive or lead to false clickthroughs. Also note that you may promote this URL in any other ways you see fit... like your AIM profile or e-mail signature. </p>
                      <p>
                        <textarea name="textarea2" rows="3" style="width: 100%; height: 80px;"><a href="http://www.meteen.com?aid=<%= CurrentUserID() %>" onMouseOver="window.status='http://www.meteen.com';" onMouseOut="window.status='';" target="_top">Meet Hot Teens</a></textarea>
                      </p></td>
                    </tr>
                  </table>
				  <p>&nbsp;</p>
				</div>
				<div id="content0">
				  <p class="style1">468x60 Ads </p>
				  <table width="100%"  border="0" cellspacing="0" cellpadding="5">
                    <tr>
                      <td><img src="../../MeTeen/ads/468x60x1.gif" width="468" height="60" border="1"></td>
                    </tr>
                    <tr>
                      <td><ul>
                        <li>$0.02 per click </li>
                        <li>Only unique clicks (per IP address) within 12 hour period are valid </li>
                        <li>$0.50 bonus will be instantly credited if user signs up within 12 hours of clickthrough (from same IP address)</li>
                        </ul>
                      <p>Copy and Paste the following HTML code into your web page or ad rotation script:</p>
                      <p>
                        <textarea name="textarea2" rows="3" style="width: 100%; height: 80px;"><a href="http://www.meteen.com?aid=<%= CurrentUserID() %>"><img src="http://www.meteen.com/meteen/ads/468x60x1.gif" border="0" alt="MeTeen.com - Start Meeting! Click Here!"></a></textarea>
                      </p></td>
                    </tr>
                  </table>
				  <p>&nbsp;</p>
				</div>
				<div id="content1">
				  <p class="style1">125x125 Ads </p>
				  <table width="100%"  border="1" cellpadding="5" cellspacing="0" bordercolor="#999999" bgcolor="#FFFFFF">
                    <tr bgcolor="#CCCCCC">
                      <td width="150" align="center"><strong>Ad Preview </strong></td>
                      <td><strong>Details</strong></td>
                    </tr>
                    <tr>
                      <td width="150" rowspan="2" align="center" valign="middle" bgcolor="#CCCCCC"><img src="../../MeTeen/ads/125x125x1.gif" width="125" height="125"></td>
                      <td><ul>
                        <li>$0.02 per click </li>
                        <li>Only unique clicks (per IP address) within 12 hour period are valid </li>
                        <li>$0.50 bonus will be instantly credited if user signs up within 12 hours of clickthrough (from same IP address)</li>
                        </ul>                      </td>
                    </tr>
                    <tr>
                      <td><p>Copy and Paste the following HTML code into your web page or ad rotation script:</p>
                      <p>
                        <textarea name="textarea" rows="3" style="width: 100%; height: 80px;"><a href="http://www.meteen.com?aid=<%= CurrentUserID() %>"><img src="http://www.meteen.com/meteen/ads/125x125x1.gif" border="0" alt="MeTeen.com - Start Meeting! Click Here!"></a></textarea>
                      </p></td>
                    </tr>
                  </table>
				  <p>&nbsp;</p>
				</div>
				<div id="content2">
				  <p class="style1">125x600 Ads </p>
				  <p>No ads in this format are currently available. </p>
				</div>
				<div id="content3">
				  <p class="style1">Popup Ads </p>
				  <p>No ads in this format are currently available. </p>
				</div>
			</td>
		  </tr>
	  </table>
		<script>hideAll(); selectTab(document.getElementById('tab0'), 'content4');</script>
	<!-- InstanceEndEditable --></td><td background="../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
