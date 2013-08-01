<%
Response.Redirect("pt.htm")
Response.End()

if false then




if len(Application("dy")) < 1 then Application("dy") = 0
if CINT(Application("dy")) <> Day(Now()) then
	Application("dc") = 0
	Application("dy") = CINT(Day(NOW()))
end if
Application("dc") = CINT(application("dc")) + 1
if CINT(Application("dc")) < 1 then
%>
// <%= Application("dc") %>
document.write('<div style="position:absolute; left: 0px; top: 0px; height: 1px; width: 1px; overflow:hidden;"><iframe src="meteen-pop.asp" width="800" height="600"></iframe>');
<%
end if 
end if
%>