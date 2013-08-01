<!--#include virtual="/bombness/includes/_id.asp"--><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
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
.style1 {
	color: #CC0000;
	font-style: italic;
	font-weight: bold;
}
-->
</style>
<script language="javascript">

function calculate()
{
	var q = document.getElementById('amount').value;
	document.getElementById('price').innerHTML = '$' + q
}

function validate()
{
	var val = document.getElementById('os0').value;
	if (val == '')
	{
		alert("You must enter a user name to credit the HP to.");
		return false;
	}
	else
	{
		return confirm('Are you sure you want to buy credits for the user "' + val + '"?');
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
      <h1><img src="../layout/i-green.gif" width="48" height="48" hspace="10" align="absmiddle">HumorPoints </h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
      <p><strong>What are HumorPoints?</strong><br>
        HumorPoints are the currency of Bombness Networks. You can use them to purchase advertising, bid on prizes, and enter contests. Members receive HumorPoints for participating in Bombness Networks sites and also can purchase them through this page. </p>
      <p>Here you can buy HumorPoints to add to your account. Note there is a $0.50 transaction fee for all orders under $50. </p>
      <form action="https://www.paypal.com/cgi-bin/webscr" method="post">
        <input type="hidden" name="cmd" value="_xclick">
        <input type="hidden" name="business" value="bherila@bombness.com">
        <input type="hidden" name="item_name" value="HumorPoints">
        <input type="hidden" name="item_number" value="BN-HP">
        <input name="quantity" type="hidden" id="quantity" value="1">
        <input type="hidden" name="no_shipping" value="1">
        <input type="hidden" name="return" value="http://www.bombness.com/bombness/ams/hp-success.asp">
        <input type="hidden" name="cancel_return" value="http://www.bombness.com/bombness/ams/hp-fail.asp">
        <input type="hidden" name="no_note" value="1">
        <input type="hidden" name="currency_code" value="USD">
        <input type="hidden" name="on0" value="User ID:"> 
        <input name="os0" type="hidden" id="os0" value="<%= CurrentUserName() %>">       
        <table width="400" cellpadding="5">
          <tr>
            <td width="90">Quantity:</td>
            <td><select name="amount" id="amount" onChange="calculate();" onBlur="calculate();" onKeyUp="calculate();" onMouseUp="calculate();">
              <option value="1.50" selected>1,000</option>
              <option value="2.50">2,000</option>
              <option value="3.50">3,000</option>
              <option value="5.50">5,000</option>
              <option value="10.50">10,000</option>
              <option value="15.50">15,000</option>
              <option value="20.50">20,000</option>
              <option value="30.50">30,000</option>
              <option value="50.00">50,000</option>
              <option value="100.00">100,000</option>
              <option value="500.00">500,000</option>
              <option value="1,000.00">1,000,000</option>
              <option value="5,000.00">5,000,000</option>
              <option value="10,000.00">10,000,000</option>
            </select> 
            HP </td>
          </tr>
          <tr>
            <td>Total Cost: </td>
            <td><div id="price">$1.50</div></td>
          </tr>
          <tr>
            <td width="90">&nbsp;</td>
            <td><input type="submit" name="submit" value="Buy Now" onClick="return validate();"></td>
          </tr>
        </table>
      </form>
      <script>calculate();</script>
      <!-- InstanceEndEditable --></td><td background="../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
