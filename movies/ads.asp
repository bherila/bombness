<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%><%
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1 

i = cint(Application("Last_Movie_Ad"))

dim ads(3)
ads(0) = "bombness_ad_1"
ads(1) = "bombness_ad_2"
ads(2) = "bombness_ad_3"

if i >= ubound(ads)-1 then i = -1
'response.write(i)
i = i + 1
Response.Write("uri=" & Server.URLEncode(ads(i)))
Application("Last_Movie_Ad") = i
erase ads
Response.End()
%>