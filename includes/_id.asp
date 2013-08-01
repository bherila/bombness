<!--#include virtual="/bombness/includes/_mail.asp"-->
<%

site = Left(Replace(Request.ServerVariables("SERVER_NAME"), "www.", ""), 3)
'Response.Write(site)

if len(Request.Cookies("sid")) < 1 then
	Response.Cookies("sid") = UniqueID()
end if
bh_loggedin = false
userlevel = 0
Set idConn = Server.CreateObject("ADODB.Connection")
idConn.Open("DSN=bombness-mysql; UID=root; pwd=eggbert;")
remoteip = Replace(Request.Cookies("sid"), "'", "''")

ip = Request.ServerVariables("REMOTE_ADDR")
pts = split(cstr(ip), ".", 4)
A = cint(pts(0))
B = cint(pts(1))
C = cint(pts(2))
D = cint(pts(3))
ipn = (A * 16777216) + (B * 65536) + (C * 256) + D
erase pts

set qrs = idconn.execute("select c.country, c.code from ip2nationCountries as c, ip2nation as i where i.ip < " & ipn & " and c.code = i.country order by i.ip desc limit 1")
if qrs.eof and qrs.bof then
	country = "Unknown"
	country2 = "Unknown"
else
	country = qrs(0)
	country2 = qrs(1)
end if
qrs.close
cd = "Unknown/Other"
set qrs = idconn.execute("select Region from country_data where Country = '" & country & "' limit 1")
if not (qrs.eof and qrs.bof) then
	cd = qrs(0)
end if
qrs.close
set qrs = nothing

bn_errortext = ""
bn_error = false

Function UniqueID()
  Randomize
  UniqueID = Request.ServerVariables("REMOTE_ADDR") & ":" & Session.SessionID & Hex(Round(rnd() * 123456789))
End Function

if Request("action") = "bn-signup" then
	Session("username") = Request("username")
	Session("password") = Request("password")
	if len(request("username")) < 1 then bn_errortext = "<li>You must specify a user name.</li>"
	if len(request("password")) < 1 then bn_errortext = bn_errortext & "<li>You must specify a password.</li>"
	if len(request("email")) < 1 then bn_errortext = bn_errortext & "<li>Your e-mail address is required.</li>"
	if len(request("firstname")) < 1 then bn_errortext = bn_errortext & "<li>Please enter your first name.</li>"
	if len(request("lastname")) < 1 then bn_errortext = bn_errortext & "<li>Please enter your last name.</li>"
	if (request("month")) < 1 then bn_errortext = bn_errortext & "<li>Please enter your month of birth.</li>"
	if (request("day")) < 1 then bn_errortext = bn_errortext & "<li>Please enter your day of birth.</li>"
	if (request("year")) < 1 then bn_errortext = bn_errortext & "<li>Please enter your year of birth</li>"
	if (request("gender")) < 0 then bn_errortext = bn_errortext & "<li>Are you male or female? You must select your gender.</li>"
	if (request("country")) = "--" or len(request("country")) < 1 then bn_errortext = bn_errortext & "<li>Plase select your country.</li>"
	if len(request("agree")) < 1 then bn_errortext = bn_errortext & "<li>Your agreement with our terms of service is required to sign up.</li>"
	if Request("password") <> Request("confirmpassword") then bn_errortext = bn_errortext & "<li>Your password and confirmation password do not match. Please correct this.</li>"
	if len(Request("username")) > 0 then
		'Check username
		set srs = idConn.Execute("select count(*) from users where username = '" & Replace(Request("username"), "'", "''") & "'")
		if srs(0) > 0 then bn_errortext = bn_errortext & "<li>The user name you selected is already taken. Please choose a different user name.</li>"
		srs.close
	end if
	if len(Request("email")) > 0 then
		'Check username
		set srs = idConn.Execute("select count(*) from users where email = '" & Replace(Request("email"), "'", "''") & "'")
		if srs(0) > 0 then bn_errortext = bn_errortext & "<li>The e-mail address you entered already has a Bombness Networks Identity associated with it. Please use a different e-mail address.</li>"
		srs.close
	end if
	if request("country") = "us" and len(Request("zipcode")) <> 5 then
		bn_errortext = bn_errortext & "<li>Please enter your five-digit zip code.</li>"
	end if
	if len(request("email")) > 4 then
		dim da(8)
		da(0) = "epals.com"
		da(1) = "meteen.com"
		da(2) = "highschoolhumor.com"
		da(3) = "bombness.com"
		da(4) = "perfectjoke.com"
		da(5) = "bherila.net"
		da(6) = "nobullgames.com"
		da(7) = "mexicoufo.com"
		
		dm = lcase(request("email"))
		for i = lbound(da) to ubound(da)
			if right(dm, len(da(i))) = da(i) and len(da(i)) > 2 then
				bn_errortext = bn_errortext & "<li>E-Mail Addresses from <b>" & da(i) & "</b> are not allowed on Bombness Networks. Please enter a different e-mail address.</li>"
			end if
		next
		erase da
	end if
	if len(bn_errortext) < 1 then
		'Everything is okay, insert the user
		userData = "lastname.visible=false" & vbnewline & "account.level=NORMAL"
		bday = Request("month") & "/" & Request("day") & "/" & Request("year")
		ref = -1
		set irs = idConn.execute("select aid from affiliates where ip = '" & Replace(Request.ServerVariables("REMOTE_HOST"), "'", "''") & "' limit 1")
		if not (irs.eof and irs.bof) then
			ref = cint(irs(0))
			idConn.execute("update users set hp = hp + 100 where id = " & ref)
			idConn.execute("delete from affiliates where ip = '" & Replace(Request.ServerVariables("REMOTE_HOST"), "'", "''") & "'")
		end if
		irs.close
		set irs = nothing
		zip = "'" & Replace(Request("zipcode"), "'", "''") & "'"
		if zip = "'NULL'" then zip = "NULL"
		sql = "insert into users (ref, username, password, email, birthday, school, data, gender, country, firstname, lastname, zipcode) values (" & ref & ", '" & Replace(Request("username"), "'", "''") & "', '" & Replace(Request("password"), "'", "''") & "', '" & Replace(Request("email"), "'", "''") & "', '" & Replace(bday, "'", "''") & "', '" & Replace(Request("schoolID"), "'", "''") & "', '" & Replace(userData, "'", "''") & "', '" & Replace(Request("gender"), "'", "''") & "', '" & Replace(Request("country"), "'", "''") & "', '" & Replace(Request("firstname"), "'", "''") & "', '" & Replace(Request("lastname"), "'", "''") & "', " & zip & ")"
		idConn.Execute(sql)
		'on error resume next
		idConn.Execute("insert into daily_youth (email, ip, conf, [key]) values ('" & Replace(Request("email"), "'", "''") & "', '" & Replace(Request.ServerVariables("REMOTE_HOST"), "'", "''") & "', 1, 'SPAGE')")
		idConn.Close
		Set idConn = Nothing
			mtext = ""
			mtext = mtext & "Welcome to Bombness Networks! This e-mail contains all you" & vbnewline
			mtext = mtext & "need to know about your new Bombness Networks Identity." & vbnewline
			mtext = mtext & "Your Identity is valid on all Bombness Networks sites and" & vbnewline
			mtext = mtext & "many of our affiliates' sites as well. Just a few of the" & vbnewline
			mtext = mtext & "sites your Identity is valid on are:" & vbnewline
			mtext = mtext & "" & vbnewline
			mtext = mtext & "- bombness.com" & vbnewline
			mtext = mtext & "- highschoolhumor.com" & vbnewline
			mtext = mtext & "- perfectjoke.com" & vbnewline
			mtext = mtext & "- bombness.com" & vbnewline
			mtext = mtext & "- meteen.com" & vbnewline
			mtext = mtext & "" & vbnewline
			mtext = mtext & "To confirm your e-mail address, please visit the following" & vbnewline
			mtext = mtext & "URL. You may not be able to use some features of our sites" & vbnewline
			mtext = mtext & "until you confirm your e-mail address." & vbnewline
			mtext = mtext & "" & vbnewline
			mtext = mtext & "http://www.bombness.com/bombness/identity/email-confirm.asp?username=" & Server.URLEncode(Request("username")) & "&email=" & Server.URLEncode(Request("email")) & vbnewline
			mtext = mtext & "" & vbnewline
			mtext = mtext & "We are always working on adding new sites for the future." & vbnewline
			mtext = mtext & "Bombness Networks is committed to providing teens around" & vbnewline
			mtext = mtext & "the world a safe, fun haven on the internet where they can" & vbnewline
			mtext = mtext & "spend hours on end enjoying all we have to offer." & vbnewline
			mtext = mtext & "" & vbnewline
			mtext = mtext & "Your new Identity unlocks many new features on all our" & vbnewline
			mtext = mtext & "web sites and is designed to enhance your experience while" & vbnewline
			mtext = mtext & "online." & vbnewline
			mtext = mtext & "" & vbnewline
			mtext = mtext & "Please save this e-mail for future reference. Your login" & vbnewline
			mtext = mtext & "details follow:" & vbnewline
			mtext = mtext & "" & vbnewline
			mtext = mtext & "  User Name:   " & request("username") & vbnewline
			mtext = mtext & "  Password:    " & request("password") & vbnewline
			mtext = mtext & "" & vbnewline
			mtext = mtext & "" & vbnewline
			mtext = mtext & "Thank you for choosing Bombness Networks. Enjoy." & vbnewline
			sendmail_immed request("email"), "support@bombness.com", "Welcome to Bombness Networks", mtext
			sendmail_immed "official_mail@bombness.com", "official_mail@bombness.com", "New User Signed Up for Bombness Networks", Request("username")
			mtext = ""
		Response.Redirect(Request("dest"))
	else
		bn_errortext = "<p>Some errors occured while processing your request. Please correct the following items, then try again.</p><ul>" & bn_errortext & "</ul>"
	end if
	set srs = nothing
elseif Request("action") = "bn-login" then
	set srs = idConn.Execute("select id from users where username = '" & Replace(Request("username"), "'", "''") & "' and password = '" & Replace(Request("password"), "'", "''") & "' limit 1")
	if srs.eof and srs.bof then
		srs.close
		set srs = nothing
		bn_error = true
		denied = true
	else
		sql = "delete from sessions where site='" & site & "' and uid = " & srs(0)
		idConn.execute(sql)
		sql = "delete from sessions where site='" & site & "' and ip = '" & replace(Remoteip, "'", "''") & "'"
		idConn.execute(sql)
		sql = "insert into sessions (uid, ip, site, pwd) values (" & srs(0) & ", '" & Replace(remoteip, "'", "''") & "', '" & site & "', '" & Replace(Request("password"), "'", "''") & "')"
		idConn.execute(sql)
		srs.close
		set srs = nothing
		if len(Request("dest")) > 0 then
			response.Redirect(Request("dest"))
			idConn.Close
			set idConn = Nothing
		end if
		bh_loggedin = true
	end if

elseif Request("action") = "bn-logout" then
	if len(site) < 1 then site = "???"
	idConn.Execute("delete from sessions where site='" & site & "' and ip = '" & Replace(remoteip, "'", "''") & "'")
end if

x_CurrentUserName = ""
CurrentPassword = ""

function eval_loginStatus()
	set xidconn = server.createobject("adodb.connection")
	xidconn.open("dsn=bombness-mysql;uid=root;pwd=eggbert;")
	sn = "select * from sessions where site='" & site & "' and ip = '" & remoteip & "' limit 1"
	set xidrs = xidconn.execute(sn)
	if not (xidrs.eof and xidrs.bof) then
		bh_loggedin = true
		xidconn.execute("update sessions set expires = DATE_ADD(NOW(), INTERVAL 1 HOUR) where ip = '" & remoteip & "'")
	end if
	xidrs.close
	set xidrs = xidconn.Execute("select username, userlevel, password from users where id = " & CurrentUserID() & " limit 1")
	if xidrs.eof and xidrs.bof then
		x_CurrentUserName = ""
		userlevel = 0
		CurrentPassword = ""
	else
		x_CurrentUserName = xidrs(0)
		userlevel = xidrs(1)
		CurrentPassword = xidrs(2)
	end if
	xidrs.close
	xidconn.close
	set xidconn = nothing
	set xidrs = nothing
end function

rootpath = "/meteen/"
function set_rootPath(path)
	rootpath = path
end function

function LogonState()
	eval_loginStatus()
	if bh_loggedin then
		LogonState = "<a href=""/bombness/identity/sign-out.asp?jump=" & Server.URLEncode(Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString()) & """><img name=""bh_loginstatus"" id=""bh_loginstatus"" src=""/bombness/identity/sign-out.jpg"" border=""0"" width=""88"" height=""31""></a>"
	else
		'LogonState = "<a href=""javascript:void(0);"" onClick=""document.getElementById('loginA').style.visibility = 'visible'; document.getElementById('loginB').style.visibility = 'visible';""><img name=""bh_loginstatus"" id=""bh_loginstatus"" src=""/bombness/identity/sign-in.jpg"" border=""0"" width=""88"" height=""31""></a>"
		LogonState = "<a href=""" & rootpath & "login.asp?jump=" & Server.URLEncode(Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString()) & """><img name=""bh_loginstatus"" id=""bh_loginstatus"" src=""/bombness/identity/sign-in.jpg"" border=""0"" width=""88"" height=""31""></a>"
	end if
end function

function CurrentUserName()
	CurrentUserName = x_CurrentUserName
end function

function CurrentUserID()
	Set fConn = Server.CreateObject("ADODB.Connection")
	fConn.Open("DSN=bombness-mysql; UID=root; PWD=eggbert;")
	sql = "select uid from sessions where ip = '" & Replace(remoteip, "'", "''") & "' limit 1"
	'response.Write(sql)
	set irs = fConn.Execute(sql)
	if irs.eof and irs.bof then
		currentuserid = -1
	else
		CurrentUserID = irs(0)
	end if
	irs.close
	set irs = nothing
	fConn.Close
	Set fConn = Nothing
end function

function PreLogin()
	eval_loginStatus
	if not bh_loggedin then
	%><style type="text/css">
<!--
.style1 {font-weight: bold}
.style2 {color: #333333}
-->
</style>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<div id="loginA" style="visibility: hidden; filter: alpha(opacity=50); -moz-opacity: 0.50; position:absolute; left:0px; top:0px; width:100%; height:100%; z-index:9999; background-color: #000000; layer-background-color: #000000; border: 1px none #000000;">&nbsp;</div>
<div id="loginB" style="visibility: hidden; position:absolute; left:0px; top:0px; width:100%; height:100%; z-index:10000;">
  <table width="100%" height="100%">
  <tr><td width="100%" height="100%" align="center" valign="middle"><table width="400" height="250" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
	<tr>
	  <td bgcolor="#FFFFFF"><iframe src="/bombness/identity/sign-in-small.asp?site=<%= Server.URLEncode(site) %>" width="400" height="250" frameborder="0" scrolling="no">FRAMES</iframe>&nbsp;</td>
	</tr>
  </table></td>
  </tr></table>
</div>
	<%
	end if
end function

function RequireLogin(logonpage)
	if not bh_loggedin then
		url = logonpage
		if inStr(url, "?") then
			url = url & "&"
		else
			url = url & "?"
		end if
		url = url & "dest="
		tp = Request.ServerVariables("SCRIPT_NAME")
		if Len(Request.QueryString()) > 0 then
			tp = tp & "?" & Request.QueryString()
		end if
		url = url & Server.URLEncode(tp)
		Response.Redirect(url)
	end if
end function

function TOS()
%>
<p><strong>Please read the following information carefully. </strong><br>
<br>
Bombness.com and the Bombness Network of web sites are the on-line web sites produced, owned and operated by Bombness Networks, (&quot;Bombness&quot;). Bombness.com and the Bombness Network of web sites includes content, information, services, and merchandise provided by BOMBNESS and third parties. Access to and use of the products and services available through this site (collectively, the &quot;Services&quot;) are subject to the following terms and conditions. <br>
<br>
<strong>All rights reserved </strong><br>
BOMBNESS reserves the right to change the terms, conditions and notices under which the Bombness.com and the Bombness Network of web sites are offered. You are responsible for regularly reviewing these terms and conditions. Unless expressly provided herein, the Services are licensed to you by Bombness.com and the Bombness Network of web sites for your online viewing only. No further reproduction or distribution of the Services is permitted. <br>
<br>
<strong>Proprietary information </strong><br>
Bombness.com and the Bombness Network of web sites contain proprietary information and copyrighted material and trademarks, including text, photos, video, graphics, music and sound. The selection and arrangement of such content, as well as specific content on Bombness.com and the Bombness Network of web sites are considered proprietary by BOMBNESS. <br>
<br>
<strong>Control over Content </strong><br>
Bombness.com and the Bombness Network of web sites also provide content and services supplied by third parties over which it has no editorial or other control. Opinions, statements, services, offers, or other information offered by third parties on Bombness.com and the Bombness Network of web sites are those of the respective third party and not of BOMBNESS. BOMBNESS neither endorses nor is responsible for the accuracy or reliability of any opinion or statement made on Bombness.com and the Bombness Network of web sites by anyone other than authorized BOMBNESS employees or its subsidiaries employees acting in their official capacities. Under no circumstances will BOMBNESS be liable for any loss or damage caused by a user's reliance on information obtained through Bombness.com and the Bombness Network of web sites. It is the responsibility of user to evaluate the accuracy of any information, opinion, offer or other content on Bombness.com and the Bombness Network of web sites. <br>
<br>
<strong>Links to Bombness.com and the Bombness Network of web sites </strong><br>
Bombness.com and the Bombness Network of web sites encourages you to use discretion while browsing the Internet using this service. Bombness.com and the Bombness Network of web sites may produce automated search results or otherwise link you to sites containing information that some people may find inappropriate or offensive. Bombness.com and the Bombness Network of web sites makes no effort to review some of the content of sites listed on some portions of our sites (indeed, given the size of our site and the frequency with which it is updated, such review would be practically impossible). Consequently, Bombness.com and the Bombness Network of web sites is not and cannot be held responsible for the accuracy, copyright compliance, legality or decency of material contained in sites listed in our search results or otherwise linked to the Bombness.com and the Bombness Network of web sites. <br>
<br>
<strong>User Registration </strong><br>
A user who has registered on Bombness.com and the Bombness Network of web sites is responsible for his/her registration name and/or password, for ensuring that use of the same complies with the terms herein and for protecting their confidentiality. The user is also responsible for obtaining and maintaining all computer and communications equipment required for access to and use of Bombness.com and the Bombness Network of web sites. We reserve the right to modify or delete any Bombness Networks account from our system at our discretion for any reason without any compensation or notice to the account holder whatsoever. <br>
<br>
<strong>Disclaimer of Warranties; Accuracy of Information </strong><br>
Neither BOMBNESS nor any of its employees, third party content providers or licensors warrant to the accuracy or reliability of any content, information, service, or merchandise provided through Bombness.com and the Bombness Network of web sites, or that Bombness.com and the Bombness Network of web sites will be uninterrupted or error free. THE SERVICES ARE PROVIDED &quot;AS IS&quot;, WITH NO WARRANTIES WHATSOEVER. ALL EXPRESS, IMPLIED AND STATUTORY WARRANTIES, INCLUDING, WITHOUT LIMITATION, THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, AND NON-INFRINGEMENT OF PROPRIETARY RIGHTS, ARE EXPRESSLY DISCLAIMED TO THE FULLEST EXTENT PERMITTED BY LAW. USER EXPRESSLY AGREES THAT USE OF BOMBNESS.COM AND THE BOMBNESS NETWORK OF WEB SITES IS AT THE USER'S SOLE RISK. <br>
<br>
<strong>LIMITS ON LIABILITY </strong><br>
This disclaimer of liability applies to any damages or injury caused by any failure of performance, error, omission, interruption, deletion, defect, delay in operation or transmission, breach of contract, tortuous behavior, negligence, or under any other cause of action. User specifically acknowledges that BOMBNESS is not liable for the defamatory, offensive or illegal conduct of other users and that the risk of injury from the foregoing rests entirely with the user. UNDER NO CIRCUMSTANCES SHALL BOMBNESS BE LIABLE TO ANY USER ON ACCOUNT OF THAT USER'S USE OF THE SERVICES. SUCH LIMITATION OF LIABILITY SHALL APPLY TO PREVENT RECOVERY OF DIRECT, INDIRECT, INCIDENTAL, CONSEQUENTIAL, SPECIAL AND EXEMPLARY DAMAGES (EVEN IF BOMBNESS HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES), ARISING FROM ANY USE OF THE SERVICES (INCLUDING SUCH DAMAGES INCURRED BY THIRD PARTIES). <br>
<br>
<strong>Changes to Bombness.com and the Bombness Network of web sites service </strong><br>
The user understands that BOMBNESS may, at any time, change or discontinue any feature of Bombness.com and the Bombness Network of web sites, including content, availability, equipment needed for access to Bombness.com and the Bombness Network of web sites, the terms and conditions applicable to Bombness.com and the Bombness Network of web sites or any part thereof. BOMBNESS may also impose conditions, including registration, fees and charges, for the use of some portions of Bombness.com and the Bombness Network of web sites which shall be effective immediately upon notice thereof on Bombness.com and the Bombness Network of web sites or by email to the user. Any use of Bombness.com and the Bombness Network of web sites after such notice shall be deemed to constitute acceptance of such changes, modifications or additions. <br>
<br>
<strong>Restriction on Worldwide access </strong><br>
Although Bombness.com and the Bombness Network of web sites is accessible worldwide, not all services and products are available to persons in all geographic locations or jurisdictions. BOMBNESS reserves the right to limit the availability of any services or products to any person, geographic area or jurisdiction it desires. Any offer for any product or service on Bombness.com and the Bombness Network of web sites is void where prohibited by law. <br>
<br>
<strong>Right to Download </strong><br>
The user may download material on Bombness.com and the Bombness Network of web sites for personal use only. In no event may the user transmit, modify, publish, sell, or in any way commercially exploit any of the content on Bombness.com and the Bombness Network of web sites. Except as permitted under law, the user may not copy, publish or use material on Bombness.com and the Bombness Network of web sites without the express permission of BOMBNESS. In the event material is downloaded and/or copied, no changes shall be made in the trademark, copyrights or attribution of users. The user does not acquire any ownership rights by downloading material located on Bombness.com and the Bombness Network of web sites. <br>
<br>
<strong>Grant to Bombness </strong><br>
In the event the user submits material to Bombness.com and the Bombness Network of web sites, he/she also grants BOMBNESS the right to edit, copy, publish and distribute such material and warrants that either he/she, or the owner of such material, has expressly granted BOMBNESS a royalty- free, worldwide, irrevocable, non-exclusive right and license to use and distribute such material. Material which is protected by copyright, trademark or other proprietary right may not be uploaded on Bombness.com and the Bombness Network of web sites without the express written permission of the owner of that copyright, trademark or other proprietary right. The user shall have the burden of determining that any material is not protected by the same. User shall be solely liable for any damage resulting from any infringement of copyrights, proprietary rights, or any other harm resulting from such a submission. <br>
<br>
<strong>No License </strong><br>
Except as expressly provided, nothing within any of the Services shall be construed as conferring any license under any of BOMBNESS' or any third party's intellectual property rights, whether by estoppel, implication, or otherwise. <br>
<br>
<strong>Copyright, Patent, and Trademark Notices </strong><br>
All marks that appear throughout the Services belong to BOMBNESS or their respective owners, and are protected by U.S. and international copyright and trademark laws. Any use of any of the marks appearing throughout the Services without the express written consent of BOMBNESS or the owner of the mark, as appropriate, is strictly prohibited. <br>
<br>
<strong>Right to Link to Bombness.com and the Bombness Network of web sites </strong><br>
By copying the Bombness.com and the Bombness Network of web sites logos you may be granted a non-assignable, non-transferable and non-exclusive license to use the logo on your website for the limited purpose of creating a link to www.Bombness.com and the other Bombness Network of web sites. You may not use the logo for any other purpose.<br>
  <br>
  <strong>Objectionable Material </strong><br>
  The user shall not place in any area of Bombness.com and the Bombness Network of web sites any material which infringes upon or violates the rights of others, is unlawful, threatening, defamatory, obscene, profane or otherwise objectionable, gives rise to civil liability or otherwise violates any law. Neither may a user place material which, in the absence of a written agreement with BOMBNESS, contains advertising of products or services or any kind of commercial solicitation. BOMBNESS shall not permit conduct which restricts or inhibits any other user from using or enjoying Bombness.com and the Bombness Network of web sites and shall have the right in its sole discretion to edit or remove any material submitted to, or posted on Bombness.com and the Bombness Network of web sites that BOMBNESS, in its sole discretion, finds to be in violation of the provisions hereof or otherwise objectionable. <br>
  <br>
  <strong>Included Beneficiaries </strong><br>
  The provisions herein are for the benefit of BOMBNESS and its third party content providers and licensors, and any assignees of the same, and each shall have the right to assert and enforce such provisions. <br>
  <br>
  <strong>Indemnification </strong><br>
  User agrees to indemnify and hold harmless BOMBNESS, its affiliates, if any, and their respective directors, officers, employees and agents from and against all claims and expenses, including attorneys' fees, arising out of the use of Bombness.com and the Bombness Network of web sites by the user. <br>
  <br>
  <strong>Governing Law </strong><br>
  Unless expressly stated to the contrary elsewhere within the Services, all legal issues arising from or related to the use of the Services shall be construed in accordance with, and all questions with respect thereto shall be determined by, the laws of the State of Arizona applicable to contracts entered into and wholly to be performed within said State. <br>
  <br>
  <br>
  <strong>Privacy Statement </strong><br>
  <br>
  Bombness.com and the Bombness Network of web sites takes seriously the issue of safeguarding your privacy online. Please read the following to understand our views and practices regarding this matter, and how they pertain to you as you make full use of our many offerings. This statement discloses the privacy practices for all of the Bombness Network of web sites. <br>
  <br>
  <strong>Our Privacy Vow </strong><br>
Our goal at Bombness.com and the Bombness Network of web sites is to &quot;unite teens&quot; by providing you with the information and services that are most relevant to you. To achieve this goal, we need to collect information to understand what differentiates you from each of our other unique users. However, we also want to make certain that your wishes are respected, so we have specific ways for you to manage your privacy requests. We collect information in two ways. </p>
<ol>
<li><strong>Registration </strong><br>
We obtain biographical information such as your name, e-mail address and sometimes items such as your postal addresses when you register for certain Bombness.com and the Bombness Network of web sites services such as contests, chat, e-mail, and promotions. This information is used to personalize your Bombness.com and the Bombness Network of web sites experience and verify your identity. Any fee based membership areas are tracked through this registration process. </li>
<li><strong>Cookies </strong><br>
Visits to our site may be tracked by giving visitors a &quot;cookie&quot; when they enter. Cookies are pieces of information that a website transfers to your computer's hard drive for record-keeping purposes. Bombness.com and the Bombness Network of web sites may use cookies to make visiting our site easier since cookies allow us to save passwords and preferences for you so that you won't have to re-enter them the next time you visit. Although most browsers are initially set up to accept cookies, if you prefer, you can reset your browser to notify you when you've received a cookie, or to refuse to accept cookies. However, if you choose that option certain sites will not function properly. Cookies help us collect anonymous click stream data for tracking user trends and patterns so that we can provide you with better services. In addition, third party advertising networks may issue cookies as part of their advertisements. Cookies are now an industry standard and you'll find them used on most major websites. </li>
</ol>
<p>&nbsp; </p>
<strong>Use of Information </strong><br>
All of the information we collect (directly from what you tell us, and indirectly from what we collect through the use of cookies) is &quot;aggregated&quot; so that the information is grouped together. We use the combined information to evaluate which products and services are successful and which ones are not. We also utilize this information to evaluate which new services we should make available on our website. We use only the anonymous click stream data to help our advertisers deliver better targeted advertisements. Bombness.com and the Bombness Network of web sites will not disclose any information about any individual user except to comply with applicable law or valid legal process.  You agree that Bombness Networks may send you e-mail regarding our products or services as well as those of our advertisers and affiliates, even though they will never be provided with your e-mail address or personal information. <br>
<br>
<strong>Sharing of Information </strong><br>
The individually identifiable information that you provide to Bombness.com and the Bombness Network of web sites will be used extensively within the Bombness.com and the Bombness Network of web sites Network to provide a personalized experience to you. For example, we may share this information with related parts of the Bombness.com and the Bombness Network of web sites family of services so you will be able to receive the additional services without having to register again. <br>
<br>
Bombness.com and the Bombness Network of web sites will never willfully disclose individually identifiable information about its users to any third party without first receiving that user's permission. This policy is designed to prevent such information from appearing in unauthorized mailings and other solicitations. <br>
<br>
<strong>Opt-Out </strong><br>
We offer you the ability to opt-out of receiving information on Bombness.com and the Bombness Network of web sites updates and new services at the point of collection. Bombness Networks reserves the right to send members information regarding their account through e-mail, at their registered e-mail address. Such information includes, but is not limited to correspondence with regard to forum posts, information on new Bombness products and services, and a subscriber newsletter that will be sent out no more than twice per week per site. Users not wishing to receive the Bombness Networks newsletters may visit their Identity Control Panel at <a href="http://www.bombness.com/bombness/identities.asp" target="_blank">http://www.bombness.com/bombness/identities.asp</a> and should complete the appropriate form. <br>
<br>
<strong>Termination<br>
</strong>Any user may terminate his or her Bombness Networks account by not logging in for a period of six months, after which point it will automatically be deleted. Even after the deletion of a Bombness Networks Identity, the user still remains under the fullest force and extent of these terms as permitted by law. <br>
<br>
<strong>Correct/Update </strong><br>
To edit your personal information, please visit your account section of Bombness.com and the Bombness Network of web sites.
<%
end function

function MiniSignupForm()
	MiniSignupForm = internal_signupForm(true)
end function
function SignupForm()
	SignupForm = internal_signupForm(false)
end function

function internal_signupForm(ismini)
%>
<style type="text/css"><!--
.tos {
	background-color: #FFFFFF;
	border: 1px solid #96AC80;
	overflow: auto;
	padding: 5px;
	width: 450px;
	height: 230px;
}
-->
</style>
<form action="" method="post" name="frmSignup" id="frmSignup">
<%= bn_errortext %>
<table width="100%"  border="0" align="center" cellpadding="5" cellspacing="0" style="border: 1px solid gray;">
<tr>
<td colspan="2" bgcolor="#EBE9ED" style="border-bottom: 1px solid #A7A6AA;"><strong>User Information </strong></td>
</tr>
<tr>
<td width="130" style="border-right: 1px solid #DDDDDD;" bgcolor="#EEEEEE">Username:</td>
<td bgcolor="#FFFFFF"><input name="username" type="text" id="username" value="<%= Request("username") %>" maxlength="50"></td>
</tr>
<tr>
<td width="130" style="border-right: 1px solid #DDDDDD;" bgcolor="#EEEEEE">Password:</td>
<td bgcolor="#F5F5F5"><input name="password" type="password" id="password" value="<%= Request("password") %>" maxlength="50"></td>
</tr>
<tr>
<td width="130" style="border-right: 1px solid #DDDDDD;" bgcolor="#EEEEEE">Confirm Password: </td>
<td bgcolor="#FFFFFF"><input name="confirmpassword" type="password" id="confirmpassword" value="<%= Request("confirmpassword") %>" maxlength="50"></td>
</tr>
<tr>
<td colspan="2" bgcolor="#EBE9ED" style="border-bottom: 1px solid #A7A6AA; border-top:1px solid #A7A6AA;"><strong>Personal Information </strong></td>
</tr>
<tr>
<td width="130" style="border-right: 1px solid #DDDDDD;" bgcolor="#EEEEEE">E-Mail Address: </td>
<td><input name="email" type="text" id="email" value="<%= Request("email") %>" maxlength="50">
<span class="style4">Private</span></td>
</tr>
<tr>
<td width="130" style="border-right: 1px solid #DDDDDD;" bgcolor="#EEEEEE">First Name: </td>
<td bgcolor="#F5F5F5"><input name="firstname" type="text" id="firstname" value="<%= Request("firstname") %>" maxlength="50">
<span class="style4">Private</span></td>
</tr>
<tr>
<td width="130" style="border-right: 1px solid #DDDDDD;" bgcolor="#EEEEEE">Last Name: </td>
<td><input name="lastname" type="text" id="lastname" value="<%= Request("lastname") %>" maxlength="50">
<span class="style4">Private</span></td>
</tr><% if not ismini then %>
<tr>
<td width="130" style="border-right: 1px solid #DDDDDD;" bgcolor="#EEEEEE">School: </td>
<td bgcolor="#F5F5F5"><table width="250"  border="0" cellspacing="0" cellpadding="0">
<tr>
<td nowrap bgcolor="#FFFFFF" style="border: 1px solid gray;"><div id="selSchool">
<%
if len(Request("school")) > 0 then
Response.Write((Request("school")))
else
%>
<span class="style2">None Selected</span>
<% end if %>
</div></td><td width="16" height="16" bgcolor="#2D5082" style="border: 1px solid #5A7B51;"><input type="button" onClick="MM_openBrWindow('/meteen/pick-school.asp','picker','scrollbars=yes,width=525,height=300')" style="width: 16px; height: 16px;" value="..."></td>
</tr>
</table>
<input name="schoolID" type="hidden" id="schoolID" value="<%

if len(Request("schoolID")) > 0 then
Response.Write(Request("schoolID"))
else
Response.Write("-1")
end if

%>">
<input name="school" type="hidden" id="school" value="<%= Request("school") %>"></td>
</tr><% end if %>
<tr>
<td width="130" style="border-right: 1px solid #DDDDDD;" bgcolor="#EEEEEE">Birthday: </td>
<td bgcolor="#FFFFFF"><select name="month" id="month" height="10">
<option value="-1" <%If (isNull(Request("month"))) Then Response.Write("SELECTED") : Response.Write("")%>>-- Month --</option>
<option value="1" <%If (Not isNull(Request("month"))) Then If ("1" = CStr(Request("month"))) Then Response.Write("SELECTED") : Response.Write("")%>>January</option>
<option value="2" <%If (Not isNull(Request("month"))) Then If ("2" = CStr(Request("month"))) Then Response.Write("SELECTED") : Response.Write("")%>>February</option>
<option value="3" <%If (Not isNull(Request("month"))) Then If ("3" = CStr(Request("month"))) Then Response.Write("SELECTED") : Response.Write("")%>>March</option>
<option value="4" <%If (Not isNull(Request("month"))) Then If ("4" = CStr(Request("month"))) Then Response.Write("SELECTED") : Response.Write("")%>>April</option>
<option value="5" <%If (Not isNull(Request("month"))) Then If ("5" = CStr(Request("month"))) Then Response.Write("SELECTED") : Response.Write("")%>>May</option>
<option value="6" <%If (Not isNull(Request("month"))) Then If ("6" = CStr(Request("month"))) Then Response.Write("SELECTED") : Response.Write("")%>>June</option>
<option value="7" <%If (Not isNull(Request("month"))) Then If ("7" = CStr(Request("month"))) Then Response.Write("SELECTED") : Response.Write("")%>>July</option>
<option value="8" <%If (Not isNull(Request("month"))) Then If ("8" = CStr(Request("month"))) Then Response.Write("SELECTED") : Response.Write("")%>>August</option>
<option value="9" <%If (Not isNull(Request("month"))) Then If ("9" = CStr(Request("month"))) Then Response.Write("SELECTED") : Response.Write("")%>>September</option>
<option value="10" <%If (Not isNull(Request("month"))) Then If ("10" = CStr(Request("month"))) Then Response.Write("SELECTED") : Response.Write("")%>>October</option>
<option value="11" <%If (Not isNull(Request("month"))) Then If ("11" = CStr(Request("month"))) Then Response.Write("SELECTED") : Response.Write("")%>>November</option>
<option value="12" <%If (Not isNull(Request("month"))) Then If ("12" = CStr(Request("month"))) Then Response.Write("SELECTED") : Response.Write("")%>>December</option>
</select>
<select name="day" id="day" height="10">
<option value="-1">-- Day --</option>
<%
dim xd
for xd = 1 to 31
Response.Write("<option value=""" & xd & """")
if request("day") * 1 = 1 * xd then response.Write(" selected")
response.write(""">")
Response.Write(xd)
response.Write("</option>")
next
%>
</select>
<select name="year" id="yeah" height="10">
<option value="-1">-- Year --</option>
<%
for xd = (Year(Now)-5) to (Year(Now)-50) step -1
Response.Write("<option value=""" & xd & """")
if request("year") * 1 = xd * 1 then response.Write(" selected")
response.write(""">")
Response.Write(xd)
response.Write("</option>")
next
%>
</select></td>
</tr>
<% 

xcountry = "--"
if len(Request("country")) > 0 then xcountry = Request("country")					

%>
<tr align="left" valign="top">
<td width="130" style="border-right: 1px solid #DDDDDD;" bgcolor="#EEEEEE">Country:</td>
<td bgcolor="#F5F5F5"><%
if country = "Unknown" then
%><select name="country">
<option value="--" <%If (Not isNull(xCountry)) Then If ("--" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Not Specified</option>
<option value="US" <%If (Not isNull(xCountry)) Then If ("US" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> United States </option>
<option value="AI" <%If (Not isNull(xCountry)) Then If ("AI" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Anguilla </option>
<option value="AR" <%If (Not isNull(xCountry)) Then If ("AR" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Argentina </option>
<option value="AU" <%If (Not isNull(xCountry)) Then If ("AU" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Australia </option>
<option value="AT" <%If (Not isNull(xCountry)) Then If ("AT" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Austria </option>
<option value="BE" <%If (Not isNull(xCountry)) Then If ("BE" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Belgium </option>
<option value="BR" <%If (Not isNull(xCountry)) Then If ("BR" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Brazil </option>
<option value="CA" <%If (Not isNull(xCountry)) Then If ("CA" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Canada </option>
<option value="CL" <%If (Not isNull(xCountry)) Then If ("CL" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Chile </option>
<option value="CN" <%If (Not isNull(xCountry)) Then If ("CN" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> China </option>
<option value="CR" <%If (Not isNull(xCountry)) Then If ("CR" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Costa Rica </option>
<option value="DK" <%If (Not isNull(xCountry)) Then If ("DK" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Denmark </option>
<option value="DO" <%If (Not isNull(xCountry)) Then If ("DO" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Dominican Republic </option>
<option value="EC" <%If (Not isNull(xCountry)) Then If ("EC" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Ecuador </option>
<option value="FI" <%If (Not isNull(xCountry)) Then If ("FI" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Finland </option>
<option value="FR" <%If (Not isNull(xCountry)) Then If ("FR" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> France </option>
<option value="DE" <%If (Not isNull(xCountry)) Then If ("DE" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Germany </option>
<option value="GR" <%If (Not isNull(xCountry)) Then If ("GR" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Greece </option>
<option value="HK" <%If (Not isNull(xCountry)) Then If ("HK" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Hong Kong </option>
<option value="IS" <%If (Not isNull(xCountry)) Then If ("IS" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Iceland </option>
<option value="IN" <%If (Not isNull(xCountry)) Then If ("IN" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> India </option>
<option value="IE" <%If (Not isNull(xCountry)) Then If ("IE" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Ireland </option>
<option value="IL" <%If (Not isNull(xCountry)) Then If ("IL" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Israel </option>
<option value="IT" <%If (Not isNull(xCountry)) Then If ("IT" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Italy </option>
<option value="JM" <%If (Not isNull(xCountry)) Then If ("JM" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Jamaica </option>
<option value="JP" <%If (Not isNull(xCountry)) Then If ("JP" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Japan </option>
<option value="LU" <%If (Not isNull(xCountry)) Then If ("LU" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Luxembourg </option>
<option value="MY" <%If (Not isNull(xCountry)) Then If ("MY" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Malaysia </option>
<option value="MX" <%If (Not isNull(xCountry)) Then If ("MX" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Mexico </option>
<option value="MC" <%If (Not isNull(xCountry)) Then If ("MC" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Monaco </option>
<option value="NL" <%If (Not isNull(xCountry)) Then If ("NL" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Netherlands </option>
<option value="NZ" <%If (Not isNull(xCountry)) Then If ("NZ" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> New Zealand </option>
<option value="NO" <%If (Not isNull(xCountry)) Then If ("NO" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Norway </option>
<option value="PT" <%If (Not isNull(xCountry)) Then If ("PT" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Portugal </option>
<option value="SG" <%If (Not isNull(xCountry)) Then If ("SG" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Singapore </option>
<option value="KR" <%If (Not isNull(xCountry)) Then If ("KR" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> South Korea </option>
<option value="ES" <%If (Not isNull(xCountry)) Then If ("ES" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Spain </option>
<option value="SE" <%If (Not isNull(xCountry)) Then If ("SE" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Sweden </option>
<option value="CH" <%If (Not isNull(xCountry)) Then If ("CH" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Switzerland </option>
<option value="TW" <%If (Not isNull(xCountry)) Then If ("TW" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Taiwan </option>
<option value="TH" <%If (Not isNull(xCountry)) Then If ("TH" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Thailand </option>
<option value="TR" <%If (Not isNull(xCountry)) Then If ("TR" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Turkey </option>
<option value="GB" <%If (Not isNull(xCountry)) Then If ("GB" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> United Kingdom </option>
<option value="UY" <%If (Not isNull(xCountry)) Then If ("UY" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Uruguay </option>
<option value="VE" <%If (Not isNull(xCountry)) Then If ("VE" = CStr(xCountry)) Then Response.Write("SELECTED") : Response.Write("")%>> Venezuela </option>
</select><% else %><%= Country %><input type="hidden" name="country" value="<%= Country2 %>"><% end if %></td>
</tr><% if Country2 = "us" then %>
<tr align="left" valign="top">
  <td style="border-right: 1px solid #DDDDDD;" bgcolor="#EEEEEE">Zip Code: </td>
  <td bgcolor="#FFFFFF"><input name="zipcode" type="text" id="zipcode" value="<%= Request("zipcode") %>" size="20" maxlength="5"></td>
</tr><% else %><input type="hidden" name="zipcode" value="NULL"><% end if %>
<tr align="left" valign="top">
<td width="130" style="border-right: 1px solid #DDDDDD;" bgcolor="#EEEEEE">Gender:</td>
<td bgcolor="#FFFFFF"><select name="gender" id="gender">
<option value="-1" selected <%If (Not isNull(Request("gender"))) Then If ("-1" = CStr(Request("gender"))) Then Response.Write("SELECTED") : Response.Write("")%>>---- Select ----</option>
<option value="0" <%If (Not isNull(Request("gender"))) Then If ("0" = CStr(Request("gender"))) Then Response.Write("SELECTED") : Response.Write("")%>>Male</option>
<option value="1" <%If (Not isNull(Request("gender"))) Then If ("1" = CStr(Request("gender"))) Then Response.Write("SELECTED") : Response.Write("")%>>Female</option>
</select></td>
</tr>
<tr align="left" valign="top">
<td width="130" style="border-right: 1px solid #DDDDDD;" bgcolor="#EEEEEE">Terms of Use: </td>
<td bgcolor="#F5F5F5"><div class="tos"><%= TOS() %></div>
<p>
<input <%If (CStr(Request("agree")) = CStr("Yes")) Then Response.Write("checked") : Response.Write("")%> name="agree" type="checkbox" id="agree" value="Yes">
I Agree to the Terms of Service </p></td>
</tr>
<tr>
<td width="130" bgcolor="#EBE9ED" style="border-top: 1px solid #A7A6AA;"><input name="action" type="hidden" id="action" value="bn-signup">
<input name="dest" type="hidden" id="dest" value="signup-2.asp"></td><td bgcolor="#EBE9ED" style="border-top: 1px solid #A7A6AA;"><input type="submit" name="Submit" value="Submit"></td>
</tr>
</table>
</form>
<%
end function

function loginForm()
	loginForm = loginForm2("loggedin.asp")
end function

function loginForm2(dest)
%>
<p>Please enter your user name and password to continue.</p>
<% if bn_error then %><p class="style1">The user name and password combination you entered was not recognized. Please try again. </p><% end if %>
<form name="frmLogin" method="post" action="">
<table border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="90">User Name: </td>
<td><input name="username" type="text" id="username" value="<%= Request("username") %>" maxlength="50"></td>
</tr>
<tr>
<td width="90" style="padding-top: 5px;">Password:</td>
<td style="padding-top: 5px;"><input name="password" type="password" id="password" maxlength="50"></td></tr>
<tr>
<td><input name="dest" type="hidden" id="dest" value="<%= dest %>">
<input name="action" type="hidden" id="action" value="bn-login"></td>
<td style="padding-top: 13px;"><input name="Submit" type="submit" id="Submit" value="Log In"></td>
</tr>
</table>
<p><a href="signup.asp">Don't have an account? Sign up for a free one!</a> </p>
</form>
<%
end function

on error resume next
	idrs.close
	idconn.close
	set idrs = nothing
	set idconn = nothing
on error goto 0

eval_loginStatus

%>