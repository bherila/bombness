<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="iso-8859-1" %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.Odbc" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Bombness Networks</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body,td,th {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 10pt;
	color: #333333;
}
body {
	background-color: #FFFFFF;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
a:link {
	color: #0066CC;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #0066CC;
}
a:hover {
	text-decoration: underline;
	color: #CC0000;
}
a:active {
	text-decoration: none;
	color: #FF0000;
}
h1 {
	font-size: 18pt;
}
-->
</style></head>

<body>
<table width="100%"  border="0" cellspacing="0" cellpadding="8">
  <tr>
    <td bgcolor="#C6DDFF" style="border-bottom: 1px solid gray;"><strong>Bombness Networks File Manager<br>
    </strong><script runat="server">
	Sub RecurseDir(ByRef SpaceUsed as Integer, Path as String)
		dim f(), fi as System.IO.FileInfo
		dim d(), di, d2 as System.IO.DirectoryInfo
		di = New System.IO.DirectoryInfo(Path)
		f = di.getFiles
		d = di.getDirectories
		for each fi in f
			SpaceUsed += fi.length
		next
		for each d2 in d
			RecurseDir(SpaceUsed, d2.FullName)
		next
	End Sub
</script><%
	
        Dim conn As New Odbc.OdbcConnection("dsn=bombness;uid=bombness;pwd=brianjef;")
        conn.Open()
        Dim cmd As New Odbc.OdbcCommand("select home, quota, can_fserve, can_del, root_url from filemanager where username = '" & Replace(Request("username"), "'", "''") & "' and password = '" & Replace(Request("password"), "'", "''") & "'", conn)
        Dim dr As Odbc.OdbcDataReader = cmd.ExecuteReader()
        If dr.Read Then
            'Response.Write("<html><head><title>File Manager</title></head><body><h1>")
            Dim dir As String = "\"
			dim root_folder as string = dr(0)
			dim quota as integer = dr(1)
			dim can_fserve as boolean = dr(2)
			dim can_del as boolean = dr(3)
			dim root_url as string = dr(4)
			
			Dim di As System.IO.DirectoryInfo
			dim used as integer = 0
			Dim fi As System.IO.FileInfo			
			dim Messages as new System.Text.StringBuilder
			dim refreshSize as boolean = false
            If Not Request("dir") Is Nothing Then dir = Request("dir")
            If Not dir.EndsWith("\") Then dir &= "\"
            If Not dir.StartsWith("\") Then dir = "\" & dir
            Dim path As String = root_folder & dir
			path = path.trimend("\") & "\"
            di = New System.IO.DirectoryInfo(path)

            Dim i As Integer
            For i = 0 To Request.Files.Count - 1
                If Request.Files(i).ContentLength > 0 Then
					if quota > 0 then
						used = 0
						RecurseDir(used, root_folder)
					end if
					if (quota > 0 and used + Request.Files(i).ContentLength < quota) or (quota <= 0) then
						dim filename as string = "unnamed.file"
						filename = New System.IO.FileInfo(Request.Files(i).FileName).Name
                    	Request.Files(i).SaveAs(path.TrimEnd("\") & "\" & filename)
                    	Messages.Append("<font color=""green"">" & Request.Files(i).FileName & " uploaded successfully.</font><br>")
                    	refreshSize = true
					else
                    	Messages.Append("<font color=""maroon"">" & Request.Files(i).FileName & " could not be uploaded because of insufficient space.</font><br>")
					end if
                End If
            Next
			
			if not Request("del") is nothing then
				System.Io.File.Delete(path.TrimEnd("\") & "\" & Request("del"))
				Messages.Append("<font color=""green"">" & Request("del") & " deleted successfully.</font><br>")
				refreshSize = true
			end if

			if RefreshSize orElse Session("space") is nothing then
				used = 0
				RecurseDir(used, root_folder)
			else
				used = Session("space")
			end if
			
			Response.Write("Logged in as " & Server.HtmlEncode(Request("username")) & " - " & FormatNumber(used / 1048576, 2) & " MB Used")
			if quota > 0 then
				Response.Write(" of " & FormatNumber(quota / 1048576, 2) & " MB (" & FormatPercent((quota - used) / quota, 0) & " free)")
			end if
			Response.Write("</td></tr><tr><td align=""left"" valign=""top""><h1>")
			
            Response.Write("Index of: " & dir & "</h1><table border=""0"" cellpadding=""10"" cellspacing=""0"" width=""100%""><tr><td nowrap align=""left"" valign=""top"">")

            If (Not Request("submit") Is Nothing) AndAlso Request("submit") = "Create" Then
                Try
                    di.CreateSubdirectory(Request("foldername"))
                    Response.Write("<font color=""green"">Directory [" & Request("foldername") & "] created.</font><br>")
                Catch ex as Exception
                    Response.Write("<font color=""maroon"">" & ex.message & " ... " & (Request("foldername")) & "</font><br>")
                End Try
            End If

			
			if not Request("del-dir") is nothing then
				System.IO.Directory.Delete(path.TrimEnd("\") & "\" & Request("del-dir"), True)
				Messages.Append("<font color=""green"">" & Request("del-dir") & " and all subitems deleted successfully.</font><br>")
			end if
			
			Response.Write(Messages.ToString)
			
%>
<table width="100%"  border="0" cellspacing="0" cellpadding="4" style="border: 1px solid silver;">
      <tr bgcolor="#EBE9ED">
        <td colspan="2">File Name </td>
        <td>Type</td>
        <td>Size</td>
        <td>Actions</td>
        </tr>
<%
            If Not dir = "\" Then
                Response.Write("<tr><td width=""20"" valign=""middle"" align=""center""><img src=""up.jpg"" align=""absmiddle""></td><td valign=""middle""><a href=""?username=" & Request("username") & "&password=" & Request("password") & "&dir=" & Server.UrlEncode(dir.Substring(0, dir.TrimEnd("\").LastIndexOf("\"))) & """>(Parent Directory)</a></td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>")
            End If
			
            Dim din As System.IO.DirectoryInfo
            For Each din In di.GetDirectories()
                Response.Write("<tr><td width=""20"" valign=""middle"" align=""center""><img src=""folder.jpg"" align=""absmiddle""></td><td valign=""middle""><a href=""?username=" & Request("username") & "&password=" & Request("password") & "&dir=" & Server.UrlEncode(dir.TrimEnd("\") & "\" & din.Name & "\") & """>" & din.Name & "</a></td><td valign=""middle"">Folder</td><td>-</td><td><a href=""fman.aspx?username=" & Request("username") & "&password=" & Request("password") & "&dir=" & Request("dir") & "&del-dir=" & din.name & """ onClick=""return confirm('Are you sure you want to delete " & din.name & "?');"">[Delete]</a></td></tr>")
            Next

            For Each fi In di.GetFiles()
                Response.Write("<tr><td width=""20"" valign=""middle"" align=""center""><img src=""file.jpg"" align=""absmiddle""></td><td valign=""middle"">" & fi.Name & "</td><td>" & fi.Extension & " file</td><td valign=""middle"">" & FormatNumber(fi.Length / 1024, 0) & " KB</td><td valign=""middle"">")
				if len(root_url) > 0 then Response.Write(" <a href=""" & root_url.TrimEnd("/") & dir.Replace("\", "/").TrimEnd("/") & "/" & fi.name & """>[View]</a> ")
				if can_fserve then Response.Write(" [Download] ")
				if can_del then Response.Write(" <a href=""fman.aspx?username=" & Request("username") & "&password=" & Request("password") & "&dir=" & Request("dir") & "&del=" & fi.name & """ onClick=""return confirm('Are you sure you want to delete " & fi.name & "?');"">[Delete]</a> ")
				response.Write("</td></tr>")
            Next
            Response.Write("</table></td><td align=""left"" valign=""top"" width=""250""><form enctype=""multipart/form-data"" method=""post"" action=""fman.aspx""><input type=""hidden"" name=""username"" value=""" & Request("username") & """><input type=""hidden"" name=""password"" value=""" & Request("password") & """><input type=""hidden"" name=""dir"" value=""" & dir & """>New Folder:<br><input disabled type=""text"" name=""foldername"" value=""Not Available""><input type=""submit"" name=""submit"" value=""Create"" disabled><br><br>Upload Files:<br>")
            For i = 1 To 5
                Response.Write("<input type=""file"" name=""file" & i.ToString & """><br>")
            Next
            Response.Write("<input type=""submit"" name=""submit"" value=""Upload""></form></td></tr></table>")
        Else
			Response.Write("Not Logged In</td></tr><tr><td align=""left"" valign=""top"">")
            Response.Write("<form action=""fman.aspx"" method=""post"">Username: <input type=""text"" name=""username""><br>Password: <input type=""password"" name=""password""><br><input type=""submit"" value=""Log In"">")
        End If
        dr.Close()
        cmd.Dispose()
        conn.Close()
        conn.Dispose()

%></td>
  </tr>
</table>
</body>
</html>