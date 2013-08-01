<%
Function ShowAd(xxx, siteid)
on error resume next
Application("TodayAds") = Application("TodayAds") + 1
ShowAd = ""
ActiveCondition = "active = 1 and (((impressions < `impressions.max`) or (`impressions.max` < 0)) and ((actions < `actions.max`) or (`actions.max` < 0)))" '"

	Set ads_Conn = Server.CreateObject("ADODB.Connection")
	ads_Conn.Open("DSN=bombness-mysql; UID=bombness; PWD=brianjef;")
	ads_Conn.execute("delete from ad_imps where expires <= NOW()")
	
	if siteid = "HSH" then
		siteid = "HSH' and id <> '86"
	end if 
	
	sql = "select html, id, ip_time, `limit` from ads where ((`limit` = 'ALL') or (`limit` = '" & siteid & "')) and (" & ActiveCondition & ") and (type = '" & xxx & "')"
	'if Len(Session("last" & xxx)) > 0 then sql = sql & " and (id <> " & Session("last" & xxx) & ")"
	
	set irs = ads_Conn.execute("select tid from ad_imps where ip = '" & Replace(Request.ServerVariables("REMOTE_ADDR"), "'", "''") & "'")
	while not irs.eof
		sql = sql & " and (id <> " & irs(0) & ")"
		irs.moveNext
	wend
	irs.close
	set irs = nothing
	
	sql = sql & " order by `limit` desc, rand() limit 1"
	'Response.Write(SQL)
	Set ads_rs = ads_Conn.Execute(sql)
	
	if ads_rs.eof and ads_rs.bof then
		'no ad
		Session("last" & xxx) = ""
		'ShowAd = ShowAd(xxx, siteid)
		ShowAd = ""
	else
		dim id
		ShowAd = ShowAd & (ads_rs(0))
		id = ads_rs(1)
		it = ads_rs(2)
		Session("last" & xxx) = id
		ads_rs.close
		ads_Conn.Execute("update ads set impressions = impressions + 1 where id = " & id)
		ads_Conn.Execute("update hshdata set intdata = intdata + 1 where id = 251")
		if it > 0 then ads_Conn.Execute("insert into ad_imps (ip, tid, expires) values ('" & Replace(Request.ServerVariables("REMOTE_ADDR"), "'", "''") & "', " & id & ", NOW() + " & cstr(it) & ")")
	end if
	
	ads_Conn.Close
	Set ads_Conn = Nothing
	
	Randomize
	ShowAd = Replace(ShowAd, "@@URL@@", Request.ServerVariables("SCRIPT_NAME"))
	ShowAd = Replace(ShowAd, "@@RND@@", Rnd())
	If len(ShowAd) < 1 then ShowAd = "No Ad Available"' <!--" & sql & "-->"
	tries = 0
End Function

Function ShowLinks(site, howMany, prefix, suffix)
	Set ads_Conn = Server.CreateObject("ADODB.Connection")
	ads_Conn.Open("DSN=bombness; UID=bombness; PWD=brianjef;")
	xsql = "select id, html from ads where (type = 4) and (limit = 'ALL' or limit='" & Replace(site, "'", "''") & "') and " & ActiveCondition & " order by ip_time desc, rand() limit " & (howMany * 1) 
	set ads_rs = ads_Conn.Execute(xsql)
	out = ""
	while not ads_rs.eof
		ads_conn.execute("update ads set impressions = impressions + 1 where id = " & ads_rs(0))
		out = out & prefix & ads_rs(1) & suffix
		ads_rs.movenext
	wend	
	ads_rs.close
	set ads_rs = nothing
	ads_Conn.Close
	Set ads_Conn = Nothing
	ShowLinks = out
	'ShowLinks = xsql
End Function

%>

