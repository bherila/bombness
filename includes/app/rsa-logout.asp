<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%><%
Set cn = Server.CreateObject("ADODB.Connection")
cn.Open("Driver={MySQL ODBC 3.51 Driver};Server=orf-mysql1.brinkster.com;uid=bombness;pwd=brianjef;database=bombness;")
cn.execute("update rsa_users set leaveoff = '" & Replace(Request("leaveoff"), "'", "''") & "' where hardware_key = '" & Replace(Request("key"), "'", "''") & "' limit 1")
cn.close
set cn = nothing
response.End()
%>