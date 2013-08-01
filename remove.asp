<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/BN-Layout.dwt" codeOutsideHTMLIsLocked="false" -->
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
	color: #009900;
	font-weight: bold;
}
.style2 {color: #990000}
-->
</style>
<!-- InstanceEndEditable -->
</head>

<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="20" align="center" valign="top" background="bnBar/bg.gif"><img src="bnBar/bombness_sel.gif" width="120" height="20"><img src="bnBar/separator.gif" width="15" height="20"><a href="http://www.meteen.com" target="_top"><img src="bnBar/meteen.gif" width="45" height="20" border="0"></a><img src="bnBar/separator.gif" width="15" height="20"><a href="http://www.highschoolhumor.com" target="_top"><img src="bnBar/hsh.gif" width="97" height="20" border="0"></a><img src="bnBar/separator.gif" width="15" height="20"><a href="http://www.perfectjoke.com" target="_top"><img src="bnBar/pj.gif" width="72" height="20" border="0"></a><img src="bnBar/separator.gif" width="15" height="20"><a href="http://www.nobullgames.com" target="_top"><img src="bnBar/nbg.gif" width="84" height="20" border="0"></a></td>
  </tr><tr>
    <td align="center" valign="top">
      <table width="775" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="layout-2005/logo1.jpg" width="166" height="102"></td>
    <td><img src="layout-2005/logo2.jpg" width="601" height="102"></td>
    <td><img src="layout-2005/logo3.gif" width="8" height="102"></td>
  </tr>
  <tr>
    <td align="left" valign="top" background="layout-2005/nav.jpg" style="padding-left: 18px;"><a href="/bombness/" class="navlink">Home</a><br>
      <a href="/bombness/advertising.asp" class="navlink">Advertising</a><br>
      <a href="/bombness/sites/" class="navlink">Our Sites</a><br>
      <a href="/bombness/identities.asp" class="navlink">BN Identities</a><br>
      <a href="/bombness/investors.asp" class="navlink">Investor Relations</a><br>
      <a href="/bombness/about.asp" class="navlink">Company Information </a></td><td height="600" align="left" valign="top" background="layout-2005/main.gif" bgcolor="#FFFFFF" style="background-repeat: repeat-x; padding-left: 20px; padding-right: 20px;"><!-- InstanceBeginEditable name="PageTitle" -->
      <h1>Remove Yourself from our Mailing List </h1>
      <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
	  <%
	  
	  success = false
	  failure = false
	  if Request("b") = "Yes" then
	  	dim cn
		set cn = server.CreateObject("adodb.connection")
		cn.open("dsn=bombness-mysql;uid=root;pwd=eggbert;")
		cn.execute("delete from mail.`hm_distributionlistsrecipients` where distributionlistrecipientaddress = '" & Replace(Request("a"), "'", "''") & "'")
		cn.close
		set cn = nothing
		success = true
		failure = false
	  else
	  	failure = true
	  end if
	  
	  %>
      <p>When you signed up for Bombness Networks, you agreed to receive e-mail from time to time about our products and services. If you no longer wish to receive this e-mail, please enter your e-mail address in the box below and click Submit.</p>
      <% if success = true then %><p class="style1">Your e-mail address was removed successfully. </p><% end if %>
	  <% if failure = true then %>
	  <p class="style1 style2">You didn't say that you were certain you wanted your e-mail address removed, so we did not remove your e-mail address. Please complete the form again if you wish to have your e-mail address removed. </p>
	  <% end if %>
      <form action="" method="post" name="form1">
        <p><strong>E-Mail Address <br>
              <input name="a" type="text" id="a">
</strong></p>
        <p><strong>Are you sure you want to remove yourself from the Bombness Networks mailing list?<br>
          <select name="b" id="b">
            <option value="No" selected>No</option>
            <option value="Yes">Yes</option>
          </select> 
          </strong></p>
        <p><strong>      
          <input type="submit" name="Submit" value="Submit">
                    </strong>
            </p>
      </form>
      <p>&nbsp; </p>
      <!-- InstanceEndEditable --></td>
      <td background="layout-2005/right.gif">&nbsp;</td>
  </tr>
</table></td>
  </tr></table>
</body>
<!-- InstanceEnd --></html>
