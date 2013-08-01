<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bombness/includes/_id.asp"-->
<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open("DSN=bombness; UID=bombness; PWD=brianjef;")
	Conn.Execute("delete from sessions where ip = '" & Replace(remoteip, "'", "''") & "'")
	conn.close
	set conn = nothing
	
	Randomize
	
	j = Request("jump")
	if instr(j, "?") then
		j = j & "&signout="
	else
		j = j & "?signout="
	end if
	j = j & rnd() * 123456789
	
	Response.Redirect(j)
%>