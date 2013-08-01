<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
Set c = Server.CreateObject("ADODB.Connection")
c.Open("DSN=bombness; UID=bombness; PWD=brianjef;")
c.execute("update ads set actions = actions + 1 where id = '" & Replace(Request("id"), "'", "''") &  "'")
c.close
set c = nothing
Response.Redirect(Request("url"))
%>