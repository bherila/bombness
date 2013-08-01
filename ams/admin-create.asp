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
set rs = conn.execute("select NOW()")
utcnow = rs(0)
rs.close
set rs = conn.execute("select owner_of from users where id = '" & Replace(CurrentUserID, "'", "''") & "' limit 1")
if len(rs(0)) < 1 then
	allowed = false
	rs.close
else
	owner = rs(0)
	rs.close
	
	if not instr(owner, request("id")) then
		allowed = false
	end if
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
%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = MM_sql_db_STRING
  MM_editTable = "ads"
  MM_editRedirectUrl = "campaigns.asp"
  MM_fieldsStr  = "name|value|html|value|impressionsmax|value|actionsmax|value|active|value|type|value|start_date|value|expires|value|days|value|price|value|ip_time|value|limit|value|sec|value"
  MM_columnsStr = "name|',none,''|html|',none,''|[impressions.max]|none,none,NULL|[actions.max]|none,none,NULL|active|none,1,0|type|none,none,NULL|start_date|',none,NULL|expires|none,1,0|days|none,none,NULL|price|none,none,NULL|ip_time|none,none,NULL|limit|',none,''|sec|none,none,NULL"

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
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
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
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
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
    <h1>Create Ad Campaign </h1>
    <!-- InstanceEndEditable -->
      <!-- InstanceBeginEditable name="PageContent" -->
      <form name="form1" method="POST" action="<%=MM_editAction%>">
        <table border="0" cellpadding="6" cellspacing="0">
          <tr align="left" valign="middle">
            <td nowrap>Name:</td>
            <td>
              <input type="text" name="name" size="32">
            </td>
          </tr>
          <tr align="left" valign="middle" bgcolor="#EBE9ED">
            <td nowrap>HTML:</td>
            <td>
              <textarea name="html" cols="64" rows="7" wrap="OFF"></textarea>
            </td>
          </tr>
          <tr align="left" valign="middle">
            <td nowrap>Max Impressions:</td>
            <td>
              <input name="impressionsmax" type="text" value="-1" size="32"> 
              -1 for unlimited
      </td>
          </tr>
          <tr align="left" valign="middle" bgcolor="#EBE9ED">
            <td nowrap>Max Actions:</td>
            <td><input name="actionsmax" type="text" value="-1" size="32"> 
              -1 for unlimited
      </td>
          </tr>
          <tr align="left" valign="middle">
            <td nowrap>Active:</td>
            <td>
              <input name="active" type="checkbox" value=1 checked>
            </td>
          </tr>
          <tr align="left" valign="middle" bgcolor="#EBE9ED">
            <td nowrap>Type:</td>
            <td>
              <select name="type">
                <%
While (NOT rsCgnTypes.EOF)
%>
                <option value="<%=(rsCgnTypes.Fields.Item("intdata").Value)%>" <%If (Not isNull("0")) Then If (CStr(rsCgnTypes.Fields.Item("intdata").Value) = CStr("0")) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(rsCgnTypes.Fields.Item("data").Value)%></option>
                <%
  rsCgnTypes.MoveNext()
Wend
If (rsCgnTypes.CursorType > 0) Then
  rsCgnTypes.MoveFirst
Else
  rsCgnTypes.Requery
End If
%>
              </select>
            </td>
          </tr>
          <tr align="left" valign="middle">
            <td nowrap>Owner:</td>
            <td>BNID-<%=CurrentUserID%></td>
          </tr>
          <tr align="left" valign="middle" bgcolor="#EBE9ED">
            <td nowrap>Campaign Start:</td>
            <td>
              <input name="start_date" type="text" value="<%= utcnow %>" size="32">
            </td>
          </tr>
          <tr align="left" valign="middle" bgcolor="#EBE9ED">
            <td nowrap>&nbsp;</td>
            <td><em>
              <input type="checkbox" name="expires" value=1>
              run for
                  <input name="days" type="text" value="0" size="10">
      days </em></td>
          </tr>
          <tr align="left" valign="middle">
            <td nowrap>Price:</td>
            <td>
              <input name="price" type="text" value="0.00" size="32"> 
              ** This amount will not be billed</td>
          </tr>
          <tr align="left" valign="middle" bgcolor="#EBE9ED">
            <td nowrap>Site:</td>
            <td><%= Server.HTMLEncode(Request("id")) %></td>
          </tr>
          <tr align="left" valign="middle">
            <td nowrap>Frequency Cap:</td>
            <td>
              <input name="ip_time" type="text" value="0" size="32">
      days </td>
          </tr>
          <tr align="left" valign="middle" bgcolor="#EBE9ED">
            <td nowrap><input name="limit" type="hidden" id="limit" value="<%= Request("id") %>"></td>
            <td>
              <input name="submit" type="submit" value="Update record">
              <input name="sec" type="hidden" id="sec" value="<%= CurrentUserID %>">
            </td>
          </tr>
        </table>
      
        <input type="hidden" name="MM_insert" value="form1">
      </form>
    <!-- InstanceEndEditable --></td><td background="../layout-2005/right.gif">&nbsp;</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
<%
rsCgnTypes.Close()
Set rsCgnTypes = Nothing
%>
