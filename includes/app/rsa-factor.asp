<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%><%
Set cn = Server.CreateObject("ADODB.Connection")
cn.Open("Driver={MySQL ODBC 3.51 Driver};Server=orf-mysql1.brinkster.com;uid=bombness;pwd=brianjef;database=bombness;")
cn.execute("insert into rsa_factors (user, target, value) values ('" & Replace(Request("usr"), "'", "''") & "', '" & Replace(Request("target"), "'", "''") & "', '" & Replace(Request("Value"), "'", "''") & "')")
cn.close
set cn = nothing
response.End()
%>