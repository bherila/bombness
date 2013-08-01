<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/sql_db.asp" -->
<!--#include virtual="/bombness/includes/_id.asp"-->
<%

if not bh_loggedin then
	response.Redirect("default.asp")
end if
allowed = true
dim conn
set conn = server.CreateObject("adodb.connection")
conn.open(MM_sql_db_STRING)
set rs = conn.execute("select owner_of from users where id = '" & Replace(CurrentUserID, "'", "''") & "' limit 1")
if len(rs(0)) < 1 then
	'allowed = false
	rs.close
else
	owner = rs(0)
	rs.close
	
	set rs = conn.execute("select `limit` from ads where id = '" & Replace(Request("id"), "'", "''") & "' limit 1")
	if not instr(owner, rs(0)) then
		'allowed = false
	end if
	rs.close
end if

set rs = nothing
conn.close
set conn = nothing

if allowed = false then
	response.Redirect("campaigns.asp?" & request.QueryString())
end if

%>
<%
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
%>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_sql_db_STRING
  MM_editTable = "ads"
  MM_editColumn = "id"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "campaigns.asp"
  MM_fieldsStr  = "name|value|html|value|impressionsmax|value|actionsmax|value|active|value|type|value|expires|value|start_date|value|days|value|price|value|limit|value|ip_time|value"
  MM_columnsStr = "name|',none,''|html|',none,''|[impressions.max]|none,none,NULL|[actions.max]|none,none,NULL|active|none,1,0|type|none,none,NULL|expires|none,1,0|start_date|',none,NULL|days|none,none,NULL|price|none,none,NULL|limit|',none,''|ip_time|none,none,NULL"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
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
%>
<%
Dim rsCgn__MMColParam
rsCgn__MMColParam = "1"
If (Request.QueryString("id") <> "") Then 
  rsCgn__MMColParam = Request.QueryString("id")
End If
%>
<%
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
%>
<%
Dim rsCgnTypes
Dim rsCgnTypes_numRows

Set rsCgnTypes = Server.CreateObject("ADODB.Recordset")
rsCgnTypes.ActiveConnection = MM_sql_db_STRING
rsCgnTypes.Source = "SELECT intdata, data  FROM hshdata  WHERE name = 'Ad.Type'  ORDER BY name ASC"
rsCgnTypes.CursorType = 0
rsCgnTypes.CursorLocation = 2
rsCgnTypes.LockType = 1
rsCgnTypes.Open()

rsCgnTypes_numRows = 0
%><%




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
    <h1>Edit Ad Campaign </h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
      <form method="POST" action="<%=MM_editAction%>" name="form1">
        <table border="0" cellpadding="6" cellspacing="0">
          <tr align="left" valign="middle" bgcolor="#EBE9ED">
            <td nowrap>ID:</td>
            <td><%=(rsCgn.Fields.Item("id").Value)%></td>
          </tr>
          <tr align="left" valign="middle">
            <td nowrap>Name:</td>
            <td>
              <input type="text" name="name" value="<%=(rsCgn.Fields.Item("name").Value)%>" size="32">            </td>
          </tr>
          <tr align="left" valign="middle" bgcolor="#EBE9ED">
            <td nowrap>HTML:</td>
            <td>
              <textarea name="html" cols="64" rows="7" wrap="OFF"><%=(rsCgn.Fields.Item("html").Value)%></textarea>            </td>
          </tr>
          <tr align="left" valign="middle">
            <td nowrap>Max Impressions:</td>
            <td>
              <input type="text" name="impressionsmax" value="<%=(rsCgn.Fields.Item("impressions.max").Value)%>" size="32"> 
              (<%=(rsCgn.Fields.Item("impressions").Value)%> total)             </td>
          </tr>
          <tr align="left" valign="middle" bgcolor="#EBE9ED">
            <td nowrap>Max Actions:</td>
            <td><input type="text" name="actionsmax" value="<%=(rsCgn.Fields.Item("actions.max").Value)%>" size="32"> 
            (<%=(rsCgn.Fields.Item("actions").Value)%> total) </td>
          </tr>
          <tr align="left" valign="middle">
            <td nowrap>Active:</td>
            <td>
              <input type="checkbox" name="active" value=1  <%If (CStr(rsCgn.Fields.Item("active").Value) = CStr("1")) Then Response.Write("checked") : Response.Write("")%>>            </td>
          </tr>
          <tr align="left" valign="middle" bgcolor="#EBE9ED">
            <td nowrap>Type:</td>
            <td>
              <select name="type">
                <%
While (NOT rsCgnTypes.EOF)
%>
                <option value="<%=(rsCgnTypes.Fields.Item("intdata").Value)%>" <%If (Not isNull((rsCgn.Fields.Item("type").Value))) Then If (CStr(rsCgnTypes.Fields.Item("intdata").Value) = CStr((rsCgn.Fields.Item("type").Value))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(rsCgnTypes.Fields.Item("data").Value)%></option>
                <%
  rsCgnTypes.MoveNext()
Wend
If (rsCgnTypes.CursorType > 0) Then
  rsCgnTypes.MoveFirst
Else
  rsCgnTypes.Requery
End If
%>
            </select>            </td>
          </tr>
          <tr align="left" valign="middle">
            <td nowrap>Owner:</td>
            <td>BNID-<%=(rsCgn.Fields.Item("sec").Value)%></td>
          </tr>
          <tr align="left" valign="middle" bgcolor="#EBE9ED">
            <td nowrap>Campaign Start:</td>
            <td>
              <input type="text" name="start_date" value="<%=(rsCgn.Fields.Item("start_date").Value)%>" size="32">            </td>
          </tr>
          <tr align="left" valign="middle" bgcolor="#EBE9ED">
            <td nowrap>&nbsp;</td>
            <td><em>
              <input type="checkbox" name="expires" value=1  <%If (CStr(rsCgn.Fields.Item("expires").Value) = CStr("1")) Then Response.Write("checked") : Response.Write("")%>>
              run for 
              <input type="text" name="days" value="<%=(rsCgn.Fields.Item("days").Value)%>" size="10"> 
            days             </em></td>
          </tr>
          <tr align="left" valign="middle">
            <td nowrap>Price:</td>
            <td>
              <input type="text" name="price" value="<%=(rsCgn.Fields.Item("price").Value)%>" size="32"> 
              // records only </td>
          </tr>
          <tr align="left" valign="middle" bgcolor="#EBE9ED">
            <td nowrap>Site:</td>
            <td>
              <input name="limit" type="text" value="<%=(rsCgn.Fields.Item("limit").Value)%>" size="10" maxlength="3">            </td>
          </tr>
          <tr align="left" valign="middle">
            <td nowrap>Frequency Cap:</td>
            <td>
              <input type="text" name="ip_time" value="<%=(rsCgn.Fields.Item("ip_time").Value)%>" size="32"> 
            days            </td>
          </tr>
          <tr align="left" valign="middle" bgcolor="#EBE9ED">
            <td nowrap>&nbsp;</td>
            <td>
              <input type="submit" value="Update record">            </td>
          </tr>
        </table>
        <input type="hidden" name="MM_update" value="form1">
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
<%
rsCgnTypes.Close()
Set rsCgnTypes = Nothing
%>
