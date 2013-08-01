<%@LANGUAGE="VBSCRIPT"%><!--#include virtual="/bombness/includes/_id.asp"--><!--#include file="../../Connections/sql_db.asp" --><%
RequireLogin("/bombness/ams/login.asp")
%>
<%
Dim ad_types
Dim ad_types_numRows

Set ad_types = Server.CreateObject("ADODB.Recordset")
ad_types.ActiveConnection = MM_sql_db_STRING
ad_types.Source = "SELECT intdata, data  FROM dbo.hshdata  WHERE name = 'Ad.Type'  ORDER BY name ASC"
ad_types.CursorType = 0
ad_types.CursorLocation = 2
ad_types.LockType = 1
ad_types.Open()

ad_types_numRows = 0
%><%

If Request("submit") = "Submit" then
	errtext = ""
	if len(request("name")) < 1 then errtext = errtext & "<li>Please enter a name for this campaign to display in your campaign manager.</li>"
	if len(request("name")) > 50 then errtext = errtext & "<li>Please limit the name of this campaign to 50 characters.</li>"
	if len(request("type")) < 1 then errtext = errtext & "<li>You must select an ad type.</li>"
	if len(request("pricing")) < 1 then errtext = errtext & "<li>Please indicate your preference pricing model.</li>"
	if len(request("ad_details")) < 20 then errtext = errtext & "<li>Please enter some more details for your advertisement.</li>"
	if len(request("ad_details")) > 20000 then errtext = errtext & "<li>Please limit your details to 20,000 characters.</li>"
	if len(request("budget")) < 1 then
		errtext = errtext & "<li>Please enter your budget for this campaign.</li>"
	else
		on error resume next
		err.number = 0
		tmp = ccur(Replace(Replace(request("budget"), "$", ""), ",", ""))
		if err.number <> 0 then errtext = errtext & "<li>Please enter a reasonable dollar amount in the budget field.</li>"
		err.number = 0
		on error goto 0
	end if
	if len(errtext) = 0 then
		dashes = "----------------------------------------"
		set cn = server.CreateObject("adodb.connection")
		cn.open(MM_sql_db_STRING)
		set rs = cn.execute("select top 1 firstname, lastname, email, username from users where id = " & CurrentUserID)
		et = ""
		if not (rs.eof and rs.bof) then
			et = et & "FNAME     = " & rs(0) & vbnewline
			et = et & "LNAME     = " & rs(1) & vbnewline & vbnewline
			et = et & "EMAIL     = " & rs(2) & vbnewline
			et = et & "USERNAME  = " & rs(3) & vbnewline
		else
			et = et & "INVALID USERINFO"
		end if
		rs.close
		set rs = nothing
		cn.close
		set cn = nothing
		et = et & "USER ID   = " & CurrentUserID() & vbnewline
		et = et & "TIMESTAMP = " & Now & vbnewline & dashes & vbnewline
		af et, "name"
		et = et & "AD TYPES"
		while not ad_types.eof
			et = et & vbnewline & "  " & ad_types(0) & vbTab & ad_types(1)
			ad_types.MoveNext
		wend
		et = et & vbnewline & vbnewline & dashes & vbnewline
		af et, "type"
		af et, "pricing"
		af et, "ad_details"
		af et, "budget"
		'Response.Write("<pre>" & et & "</pre>")
		sendmail_immed "official_mail@bombness.com", "official_mail@bombness.com", "Advertising Inquiry", et
		ad_types.Close()
		Set ad_types = Nothing
		Response.Redirect("done.asp")
	else
		errtext = "<p class=""style1"">Please correct the following errors:</p><ul class=""style1"">" & Replace(errtext, "<li>", "<li class=""style1"">") & "</ul>"
	end if
End if

function af(a, b)
	a = a & ucase(b) & vbnewline & Request(b) & vbnewline & vbnewline & dashes & vbnewline
end function

%>
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
<!-- InstanceBeginEditable name="head" -->
<style type="text/css">
<!--
.style1 {color: #990000}
-->
</style>
<!-- InstanceEndEditable -->
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
          <!--#include file="../nav.htm"-->
    <!-- InstanceEndEditable --></td>
    <td height="600" align="left" valign="top" background="../../layout-2005/main.gif" bgcolor="#FFFFFF" style="background-repeat: repeat-x; padding-left: 20px; padding-right: 20px;"><!-- InstanceBeginEditable name="PageTitle" -->
      <h1><img src="../../layout/i-green.gif" width="48" height="48" hspace="10" align="absmiddle">Create Advertising Campaign </h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
      <p>Congratulations on your decision to use Bombness Networks advertising for your needs. We take pride in the ease and simplicity of our ad management, as well as the value of our services. We are confident that you will love our advertising system and will  enjoy doing business with us in the future.</p>
      <p>To launch an ad campaign, please fill out this form.</p>
      <%= errtext %>
      <form action="" method="post" name="form1">
        <p><strong>
        Campaign Name<br>
          </strong>For your internal purposes in the ad management system<br>
          <input name="name" type="text" id="name" value="<%= Request("name") %>"> 
        </p>
        <p><strong>Ad Format <br>
          <select name="type" id="type">
            <%
While (NOT ad_types.EOF)
%>
            <option value="<%=(ad_types.Fields.Item("intdata").Value)%>" <%If (Not isNull(Request("type"))) Then If (CStr(ad_types.Fields.Item("intdata").Value) = CStr(Request("type"))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(ad_types.Fields.Item("data").Value)%></option>
            <%
  ad_types.MoveNext()
Wend
If (ad_types.CursorType > 0) Then
  ad_types.MoveFirst
Else
  ad_types.Requery
End If
%>
          </select>
        </strong></p>
        <p><strong>Campaign Details<br>
        </strong>Write the HTML code you want your ad to display, <em>or</em> describe it in as much detail as possible.<br>
        <textarea name="ad_details" cols="64" rows="8" id="ad_details"><%= Request("ad_details") %></textarea>
        </p>
        <p><strong>Pricing Structure<br>
        </strong>
          <input <%If (CStr(Request("pricing")) = CStr("CPM")) Then Response.Write("CHECKED") : Response.Write("")%> name="pricing" type="radio" value="CPM">
CPM<br>
<input <%If (CStr(Request("pricing")) = CStr("CPC")) Then Response.Write("CHECKED") : Response.Write("")%> name="pricing" type="radio" value="CPC"> 
CPC<br>
<input <%If (CStr(Request("pricing")) = CStr("CPA")) Then Response.Write("CHECKED") : Response.Write("")%> name="pricing" type="radio" value="CPA"> 
CPA</p>
        <p><strong>Budget<br>
        </strong>Ad campaigns have a $5 minimum purchase; you can add value at any time <strong>
          <br>
        </strong>
          <input name="budget" type="text" id="budget" value="<%= Request("budget") %>">
</p>
        <p><strong>Submit Request </strong><br>
          We manually process each and every request to create the most personalized experience for all of our advertisers. Within one to two business days, you should receive a reply from us with more details and questions for you. For best results, please do not submit the same request more than once. In rare circumstances, replies from us could take longer than normal. If this is the case, we apologize for the delay. You will not be billed for this campaign until you agree to all terms of it. We reserve the right to hold your campaign until payment is completed, as well as the right to refuse any campaign for any reason, at our sole discretion. Thank you for choosing Bombness Networks! </p>
        <p>
          <input type="submit" name="Submit" value="Submit"> 
        </p>
      </form>
      <p>&nbsp; </p>
      <!-- InstanceEndEditable --></td>
    <td background="../../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table>
</body><!-- InstanceEnd --></html>
<%
ad_types.Close()
Set ad_types = Nothing
%>
