<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%><%
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open("DSN=bombness; UID=bombness; PWD=brianjef;")
Conn.Execute("update messages set [read] = " & cstr(cint(Request("status")) & " where id = " & cstr(cint(Request("id")))))
Conn.Close
Set Conn = NOthing
%>