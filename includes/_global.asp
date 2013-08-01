<%
function GetUserName(id, cn)
	on error resume next
	Err.Number = 0
	sxql = "select top 1 username from users where id = '" & Replace(id, "'", "''") & "'"
	Set RxS = cn.Execute(sxql)
	if rxs.eof and rxs.bof then
		GetUserName = "Anonymous"
	else
		GetUserName = RxS(0)
	end if
	RS.Close
	Set RxS = Nothing
	If err.Number <> 0 then GetUserName = Err.Description
	Err.Number = 0
end function

function addtag(src, tag)
	addtag = src & "|<(?:/)?" & tag & "[^>]*>"
end function

function wordReplace(src, word, replacement)
	dim rx
	set rx = new RegExp
	rx.global = true
	rx.ignorecase = true
	rx.pattern = "((?:[.,!?\s])+)(" & word & ")((?:[.,!?\s])+)"
	wordReplace = rx.replace(src, replacement)
	set rx = nothing
end function

function bhtml(src)
	'on error resume next
	'err.number = 0
	bhtml = src
	Dim emotes(40)
	dim files(40)
	emotes(0)  = "-->"  			: files(0)  = "arrow"
	emotes(1)  = "->"  				: files(1)  = "arrow"
	emotes(2)  = "<downarrow>" 		: files(2)  = "arrowd"
	emotes(3)  = "<--"  			: files(3)  = "arrowl"
	emotes(4)  = "<-"  				: files(4)  = "arrowl"
	emotes(5)  = "<uparrow>"  		: files(5)  = "arrowu"
	emotes(6)  = ":-D" 				: files(6)  = "biggrin"
	emotes(7)  = ":D"				: files(7)  = "biggrin"
	emotes(8)  = ":-B"				: files(8)  = "cheesygrin"
	emotes(9)  = ":B"				: files(9)  = "cheesygrin"
	emotes(10) = ":S"				: files(10) = "confused"
	emotes(11) = ":-S"				: files(11) = "confused"
	emotes(12) = "B-)"				: files(12) = "cool"
	emotes(13) = ":'("				: files(13) = "cry"
	emotes(14) = ":'-("				: files(14) = "cry"
	emotes(15) = "(C)"				: files(15) = "cool"
	emotes(16) = ":-0"				: files(16) = "eek"
	emotes(17) = ":-O"				: files(17) = "eek"
	emotes(18) = ":0"				: files(18) = "eek"
	emotes(19) = ":O"				: files(19) = "eek"
	emotes(20) = ":@"				: files(20) = "mad"
	emotes(21) = ":-@"				: files(21) = "mad"
	emotes(22) = "(6)"				: files(22) = "evil"
	emotes(23) = "(!)"				: files(23) = "exclaim"
	emotes(24) = ":-("				: files(24) = "frown"
	emotes(25) = ":("				: files(25) = "frown"
	emotes(26) = "(idea)"			: files(26) = "idea"
	emotes(27) = "(i)"				: files(27) = "idea"
	emotes(28) = "(lol)"			: files(28) = "lol"
	emotes(29) = ":|"				: files(29) = "neutral"
	emotes(30) = ":-|"				: files(30) = "neutral"
	emotes(31) = "(?)"				: files(31) = "question"
	emotes(32) = ":$"				: files(32) = "redface"
	emotes(33) = ":-$"				: files(33) = "redface"
	emotes(34) = "8)"				: files(34) = "rolleyes"
	emotes(35) = "8-)"				: files(35) = "rolleyes"
	emotes(36) = ":)"				: files(36) = "smile"
	emotes(37) = ":-)"				: files(37) = "smile"
	emotes(38) = ";-)"				: files(38) = "wink"
	emotes(39) = ";)"				: files(39) = "wink"
	
	Dim loRegExp
	Set loRegExp = New RegExp
	loRegExp.Global = True
	loRegExp.IgnoreCase = True

	'Find HTML Tags
	Dim tl
	tl = "<(?:/)?a[^>]*>"
	tl = addtag(tl, "script")
	tl = addtag(tl, "frame")
	tl = addtag(tl, "iframe")
	tl = addtag(tl, "style")
	tl = addtag(tl, "input")
	tl = addtag(tl, "form")
	tl = addtag(tl, "textarea")
	tl = addtag(tl, "%")
	
	loRegExp.Pattern = tl
	bhtml = loRegExp.Replace(bhtml, "")
	bhtml = Replace(bhtml, vbNewLine, "<br>")

	'Remove Superfluous Returns
	loRegExp.Pattern = "(?:<br[^>]*>(?:[\s])*)+"
	bhtml = loRegExp.Replace(bhtml, "<br>")
	loRegExp.Pattern = "(?:<br>)+(?:[\s])*<p[^>]*>"
	bhtml = loRegExp.Replace(bhtml, "<p>")
	loRegExp.Pattern = "(?:<br>)+(?:[\s])*</p[^>]*>"
	bhtml = loRegExp.Replace(bhtml, "</p>")

	'Find Hyperlinks
	loRegExp.Pattern = "((mailto\:|(news|(ht|f)tp(s?))\://){1}[^<\s]+)\b"
	bhtml = loRegExp.Replace(bhtml, "<a href=""$1"" target=""_blank"" class=""aLink"">$1</a>")

	'Find E-Mails
	loRegExp.Pattern = "(\S+@\S+.\.\S\S\S?)"
	bhtml = loRegExp.Replace(bhtml, "<a href=""mailto:$1"" class=""aEml"">$1</A>")
	
	dim i
	for i = lbound(emotes) to ubound(emotes)
		bhtml = replace(bhtml, emotes(i), "<img valign=""middle"" hspace=""1"" src=""/bombness/images/smilies/icon_" & files(i) & ".gif"" border=""0"">")
	next
	
	'if bhtml = o then bhtml = replace(replace(bhtml, vbnewline & vbnewline, "<br><br>"), vbnewline, "<br>")

	Set oRegExp = Nothing
	
	bhtml = replace(bhtml, "mtp:", "/imgsrv?ext=mup&f=")
	
	if err.number <> 0 then
		bhtml = Replace(Server.HTMLEncode(src), vbnewline, "<br>")
	end if
end function


%>
