<%@LANGUAGE="VBSCRIPT"%><!--#include file="../Connections/sql_db.asp" --><!--#include virtual="/bombness/includes/_id.asp"--><%

if not bh_loggedin then
	response.Redirect("default.asp")
end if
allowed = true
dim conn
set conn = server.CreateObject("adodb.connection")
conn.open(MM_sql_db_STRING)
set rs = conn.execute("select owner_of from users where id = '" & Replace(CurrentUserID, "'", "''") & "' limit 1")
if len(rs(0)) < 1 then
	allowed = false
	rs.close
else
	owner = rs(0)
	rs.close
	
	set rs = conn.execute("select `limit` from ads where id = '" & Replace(Request("id"), "'", "''") & "' limit 1")
	if not instr(owner, rs(0)) then
		allowed = false
	end if
	rs.close
end if

set rs = nothing
conn.close
set conn = nothing

if allowed = false then
	response.Redirect("campaigns.asp?" & request.QueryString())
end if

%><%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%><%
' *** Delete Record: declare variables

if (CStr(Request("MM_delete")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_sql_db_STRING
  MM_editTable = "ads"
  MM_editColumn = "id"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "campaigns.asp"

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If
  
End If
%><%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql delete statement
  MM_editQuery = "delete from " & MM_editTable & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%><%
Dim rsCgn__MMColParam
rsCgn__MMColParam = "1"
If (Request.QueryString("id") <> "") Then 
  rsCgn__MMColParam = Request.QueryString("id")
End If
%><%
Dim rsCgn
Dim rsCgn_numRows

Set rsCgn = Server.CreateObject("ADODB.Recordset")
rsCgn.ActiveConnection = MM_sql_db_STRING
rsCgn.Source = "SELECT *  FROM ads  WHERE id = " + Replace(rsCgn__MMColParam, "'", "''") + " limit 1"
rsCgn.CursorType = 2
rsCgn.CursorLocation = 2
rsCgn.LockType = 1
rsCgn.Open()

rsCgn_numRows = 0
%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
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
    <td align="left" valign="top" background="../layout-2005/nav.jpg" style="padding-left: 18px;"><!-- InstanceBeginEditable name="Sidebar" --><!--#include file="nav.htm"--><!-- InstanceEndEditable --></td>
    <td height="600" align="left" valign="top" background="../layout-2005/main.gif" bgcolor="#FFFFFF" style="background-repeat: repeat-x; padding-left: 20px; padding-right: 20px;"><!-- InstanceBeginEditable name="PageTitle" -->
    <h1>Delete Ad Campaign </h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
        <p>Are you absolutely sure you want to permanently delete this campaign? <br>
          <strong>WARNING: Refunds will NOT be issued for incomplete campaigns. </strong></p>
        <p>
          <input type="submit" name="Submit" value="Delete" style="font-weight: bold;"> 
          <input type="button" name="Submit2" value="Don't Delete" onClick="history.back(1);">
        </p>
      
        <input type="hidden" name="MM_delete" value="form1">
        <input type="hidden" name="MM_recordId" value="<%= rsCgn.Fields.Item("id").Value %>">
      </form>
      <p>&nbsp;</p>
    <!-- InstanceEndEditable --></td><td background="../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
<%
rsCgn.Close()
Set rsCgn = Nothing
%>
