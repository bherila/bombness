<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Configuration Options
	schoolYear = 2006
	connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=e:\webroot\bombness\crsprereg\crsprereg.mdb;Persist Security Info=False;"

' Course Fields
	dim C(47) 
	C(0) = "FA01"
	C(1) = "FA02"
	C(2) = "FA03"
	C(3) = "CS01"
	C(4) = "CS02"
	C(5) = "CSPL"
	C(6) = "EN01"
	C(7) = "HS01"
	C(8) = "HS01B"
	C(9) = "HS02"
	C(10) = "HS02B"
	C(11) = "HS03"
	C(12) = "HS03B"
	C(13) = "HS04"
	C(14) = "HS04B"
	C(15) = "HS05"
	C(16) = "HS05B"
	C(17) = "HS06"
	C(18) = "HS06B"
	C(19) = "LA01"
	C(20) = "LA02"
	C(21) = "LAPL"
	C(22) = "MA01"
	C(23) = "MA02"
	C(24) = "MU01"
	C(25) = "SC01"
	C(26) = "SC01B"
	C(27) = "SC02"
	C(28) = "SC02B"
	C(29) = "SC03"
	C(30) = "SC03B"
	C(31) = "SC04"
	C(32) = "SC04B"
	C(33) = "SC05"
	C(34) = "SC05B"
	C(35) = "SC06"
	C(36) = "SC06B"
	C(37) = "SC07"
	C(38) = "SC07B"
	C(39) = "SC08"
	C(40) = "SC08B"	
	C(41) = "RE01"
	C(42) = "RE01B"
	C(43) = "RE02"
	C(44) = "RE02B"
	C(45) = "REany"
	C(46) = "Comments"

'Editing should not be necessary below this line
'Setting initial variables:
	auth = false
	userid = -1
	page = "Login"
	firstname = ""
'Filling in variables from session cookie
	if len(session("userid")) > 0 then userid = session("userid")
	if len(session("firstname")) > 0 then firstname = session("firstname")
'Open a database connection
	set conn = server.CreateObject("adodb.connection")
	conn.open(connectionstring)
'Processing Requested Actions
	if Request("action") = "login" then
	'Authenticate the user against the database
	cond = "lname = '" & Replace(Request("lastname"), "'", "''") & "' and [sn] = '" & Replace(Request("studentid"), "'", "''") & "'"
	set rs = conn.execute("select count(*) from [results] where " & cond)
		if rs(0) > 0 then
			'User is authenticated
			session("userid") = request("studentid")
			er = false
			auth = true
			page = "Main Menu"
			'Figure out the user's first name
			set rs2 = conn.execute("select top 1 fname from [results] where " & cond)
			firstname = rs2(0)
			session("firstname") = rs2(0)
			'Close the recordset object from finding the user's first name
			rs2.close
			set rs2 = nothing
		else
			'The user is not authenticated.
			session("userid") = ""
			er = true
			auth = false
			page = "Login"
		end if
		'Close the recordset from user validation
		rs.close
		'Free memory allocated to the recordset
		set rs = nothing
	end if
		
	if page = "Main Menu" then
		'We need to figure out if the user has already submitted a pre-registration request.
		sql = "select "
		i = 0
		for i = 0 to ubound(C) - 1
			sql = sql & "len(" & c(i) & "), "
		next
		sql = sql & "0 from [Results] where sn = '" & Replace(session("userid"), "'", "''") & "'"
		set rs = conn.execute(sql)
		v = 0
		for i = 0 to rs.fields.count - 1
			if not isNull(rs(i)) then v = v + cint(rs(i))
		next
		rs.close
		set rs = nothing
		if v > 0 then
			proc = true
			page = "Report"
		else
			proc = false
		end if
	end if
	
	if page = "Report" then
		'User has already submitted a course request. Show them what they have selected.
		report = "<table border=""0"" cellpadding=""3"" cellspacing=""0"" style=""border: 1px solid gray;"" align=""left"" width=""100%"">"
		sql = "select "
		i = 0
		for i = 0 to ubound(C) - 1
			sql = sql & c(i) & ", "
		next
		sql = sql & "0 from [Results] where sn = '" & Replace(session("userid"), "'", "''") & "'"
		set rs = conn.execute(sql)
		for i = 0 to rs.fields.count - 1
			if not isNull(rs(i)) then
				if len(rs(i)) > 0 then
					report = report & vbnewline & "<tr><td width=""80"" align=""left"" valign=""top"" style=""border-right: 1px solid silver; border-bottom: 1px solid #DEDEDE; background-color: #EFEFEF;"">" & rs.fields.item(i).Name & "</td><td>" & rs(i) & "</td></tr>"
				end if
			end if
		next
		rs.close
		set rs = nothing
	end if


'Close and free the connection object
conn.close
set conn = nothing

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
     "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/Template.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<meta http-equiv="content-type" content="text/html; charset=iso-8859-1" />
<meta http-equiv="content-type" content="text/html; charset=iso-8859-1" />
<title>Delbarton School: On Campus</title>
<meta name="keywords" content="independent day school, independent school, college prep school, boys college prep, boys day school, Morristown New Jersey, Abbey Orchestra " />
<meta name="description" content="Delbarton is an independent college preparatory school for boys grades seven through twelve. The School is administered by the Roman Catholic Benedictine monks of St. Mary's Abbey and is rooted in the values of the Christian community and in the monastic tradition of a strong liberal arts education. " />
<script language="JavaScript" type="text/JavaScript" src="/js/si_behaviors.js"></script>
<script language="JavaScript" type="text/JavaScript" src="/js/si_onload.js"></script>
<script language="JavaScript" type="text/JavaScript">
// <![CDATA[
favicon = new Image();
favicon.src="/images/favicon.ico"; 
// ]]>
</script>

<link rel="shortcut icon" href="/images/favicon.ico" />
<link href="/css/print.css" rel="stylesheet" type="text/css" media="print" />
<style type="text/css" media="screen">
/* <![CDATA[ */
@import url(/CrsPreReg/css/modern.css);
/* ]]> */
.style1 {
	color: #3C5045;
	font-weight: bold;
}
</style>
<!--[if IE]><style type="text/css" media="screen">/* <![CDATA[ */ @import url(/CrsPreReg/css/iepc.css); /* ]]> */</style><![endif]-->

</head>
<body class="tier-2 layout-s1-p2 home" id="parents">
	<div id="container">
		<div id="inner-container">
		<!-- header -->
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>

<div id="header">


<div id="masthead">
<table width="400" border="0" cellspacing="0" cellpadding="0">
<tr>
<td><a href="index.html"><img src="imagesSP/hdr-nav-name.gif" alt="Delbarton School" width="349" height="44" border="0" /></a><a href="admissions/admissions.html" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Admissions','','imagesSP/hdr-nav-admissions-on.gif',0)"><img src="imagesSP/hdr-nav-admissions.gif" alt="Admissions" name="Admissions" width="133" height="44" border="0"></a><a href="oncampus/oncampus.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('On Campus','','imagesSP/hdr-nav-oncampus-on.gif',0)"><img src="imagesSP/hdr-nav-oncampus.gif" alt="On Campus" name="On Campus" width="101" height="44" border="0"></a><a href="alumni/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Alumni','','imagesSP/hdr-nav-alumni-on.gif',0)"><img src="imagesSP/hdr-nav-alumni.gif" alt="Alumni" name="Alumni" width="78" height="44" border="0"></a><a href="parents/parents.html" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Parents','','imagesSP/hdr-nav-parents-on.gif',0)"><img src="imagesSP/hdr-nav-parents.gif" alt="Parents" name="Parens" width="108" height="44" border="0"></a></td>
</tr><tr><td><img src="imagesSP/hdr-nav-image.jpg" alt="Delbarton School" width="769" height="46" border="0" /></td></tr>
<tr><td><a href="/announcements" target="_blank" onMouseOver="MM_swapImage('Annoucements','','imagesSP/hdr-nav-annoucements-on.gif',0)" onMouseOut="MM_swapImgRestore()"><img src="imagesSP/hdr-nav-annoucements.gif" alt="Annoucements" name="Annoucements" width="115" height="14" border="0"></a><a href="/calendars" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Calendars','','imagesSP/hdr-nav-calendar-on.gif',0)"><img src="imagesSP/hdr-nav-calendar.gif" alt="Calendars" name="Calendars" width="101" height="14" border="0"></a><a href="oncampus/academics/ac_academics.html" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Academics','','imagesSP/hdr-nav-academics-on.gif',0)"><img src="imagesSP/hdr-nav-academics.gif" alt="Academics" name="Academics" width="101" height="14" border="0"></a><a href="/OnCampus/Student_Life/student_life.html" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Student Life','','imagesSP/hdr-nav-studentlife-on.gif',0)"><img src="imagesSP/hdr-nav-studentlife.gif" alt="Student Life" name="Student Life" width="101" height="14" border="0"></a><a href="http://mail.delbarton.org/exchange" target="_blank" onMouseOver="MM_swapImage('WebMail','','imagesSP/hdr-nav-webmail-on.gif',0)" onMouseOut="MM_swapImgRestore()"><img src="imagesSP/hdr-nav-webmail.gif" alt="WebMail" name="WebMail" width="81" height="14" border="0"></a><a href="oncampus/bookstore/bookstore.html" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Bookstore','','imagesSP/hdr-nav-bookstore-on.gif',0)"><img src="imagesSP/hdr-nav-bookstore.gif" alt="Bookstore" name="Bookstore" width="81" height="14" border="0"></a><a href="/OnCampus/Bookstore/bookstore.html" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Contact','','imagesSP/hdr-nav-contact-on.gif',0)"><img src="imagesSP/hdr-nav-contact.gif" alt="Contact" name="Contact" width="81" height="14" border="0"></a><a href="/OnCampus/contact.html"><img src="imagesSP/hdr-nav-holder.gif" alt=" " width="108" height="14" border="0" /></a></td>
</tr><tr><td><img src="imagesSP/hdr-nav-rule.gif" alt="" width="770" height="3" border="0" /></td>
</tr></table><address><strong>Delbarton School</strong><br >230 Mendham Road | Morristown NJ 07960 | 973.538.3231</address></div><div id="nav">
<div class="hidden"><p><a href="#nav-sub">Skip to Section Subnavigation</a><br /><a href="#primary-content">Skip to Page Content</a></p></div>
<div class="nav-util">

UTIL NAV
</div>
<div id="nav-main">
main-nav
</div><!--# include virtual="/includesHTM/breadcrumbs.htm" -->
</div></div><div id="content"><div class="hidden"><p><a href="#primary-content">Skip to Page Content</a></p></div>

<div id="nav-sub" class="nav-sub">
Sub Nav
</div>	
<div id="inner-content">
<!-- //header -->
<div id="primary-content" class="pc">
<!-- news -->
      <!-- InstanceBeginEditable name="Content" -->
<% if page = "Login" then %>
      <form name="form1" id="form1" method="post" action="">
        <p style="font-weight: bold">Welcome to Delbarton Course Pre-Registration for the <%= SchoolYear %> School Year</p>
        <% if er then %><p style="color: #990000; font-weight: bold;">We could not validate your request. Please check that you correctly entered both your Student ID and your Last Name.</p><% end if %>
        <p>Please enter your Student ID and Last Name to begin.</p>
        <table border="0" cellpadding="10" cellspacing="0" bgcolor="#EBE9ED" style="border: 1px solid gray;">
          <tr>
            <td><table border="0" cellspacing="0" cellpadding="0">
              <tr align="left" valign="top">
                <td width="90" height="28">Student ID: </td>
                <td height="28"><input name="studentID" type="text" id="studentID2" value="<%= Request("studentID") %>"/></td>
              </tr>
              <tr align="left" valign="top">
                <td width="90" height="28">Last Name: </td>
                <td height="28"><input name="lastName" type="text" id="lastName2" value="<%= Request("lastName") %>"/></td>
              </tr>
              <tr>
                <td width="90"><input name="action" type="hidden" id="action" value="login" /></td>
                <td><input type="submit" name="Submit" value="Continue" /></td>
              </tr>
            </table></td>
          </tr>
        </table><p style="">&nbsp;</p>
        <p style="">If you cannot remember or locate your Student ID, please contact the main office for assistance.</p>
      </form>
<% elseif page = "Main Menu" then %>
		<p style="font-weight:bold ">Welcome, <%= firstname %></p><p><%= v %></p>
<% elseif page = "Report" then %>
		<p style="font-weight:bold ">Welcome, <%= firstname %></p>
		<p>You have successfully submitted your course registration for the <%= SchoolYear %> school year. Below, you will find a copy of your placement request.</p>
		<%= report %>
<% end if %>
      <!-- InstanceEndEditable --></div>
<div id="secondary-content" class="sc">

					<div class="sc-daily">
					  <p><strong>Course Pre-Registration</strong></p>
					  <p class="style1">Please complete the Course Pre-Registration wizard to enroll in your next year at Delbarton. 					        </p>
					</div>
					
		</div>
						<!-- footer -->
   </div>
	<!-- /inner-content -->
		</div>
<!-- /content -->
		
		<div id="footer">
			<!-- footer content -->
		</div>
		<!-- //footer -->
		</div>
	</div>
</body>

<!-- InstanceEnd --></html>