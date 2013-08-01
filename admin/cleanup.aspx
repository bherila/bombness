<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="iso-8859-1" %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.Odbc" %>
<script runat="server">
   Const Root As String = "\\premfs16\sites\premium16\bombness\temp\"
   
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Write("<html><head><title>Bombness Networks File Cache Manager</title></head><body>Opening data connection...<br>")
        Response.Flush()
        Dim cn As New System.Data.Odbc.OdbcConnection("Driver={MySQL ODBC 3.51 Driver};Server=mexicoufo.com;uid=mexicouf_bherila;pwd=eggbert;database=mexicouf_bombness;")
        cn.Open()
        Dim cmd As New System.Data.Odbc.OdbcCommand
        cmd.Connection = cn
        cmd.CommandType = CommandType.Text

        Dim f(), fi As System.IO.FileInfo
        Dim di As System.IO.DirectoryInfo
        If Not Request("reset") Is Nothing Then
            di = New System.IO.DirectoryInfo(Root)
            f = di.GetFiles()
            Response.Write("Deleting " & FormatNumber(f.Length, 0) & " files...")
            Response.Flush()
            For Each fi In f
                Try
                    fi.Delete()
                Catch
                End Try
            Next
            Erase f
            Response.Write("Done.<br>")
        End If
        If Not Request("sres") Is Nothing Then
            Response.Write("Collecting files...")
            cmd.CommandText = "select filename from file_cache where cached = 1 and filename not like '%vore' order by score desc limit 200"
			
				Dim al As New ArrayList
				Dim dr As System.Data.Odbc.OdbcDataReader = cmd.ExecuteReader
				While dr.Read
					al.Add(CStr(dr(0)))
				End While
				dr.Close()
			
				cmd.commandText = "select filename from file_cache where cached = 1 and filename not like '%vore' order by created desc limit 200"
				dr = cmd.ExecuteReader
				While dr.Read
					al.Add(CStr(dr(0)))
				End While
				dr.Close()
			
            dr = Nothing
            di = New System.IO.DirectoryInfo(Root)
            f = di.GetFiles()
            Response.Write("Done.<br>Deleting " & FormatNumber(f.Length - al.Count, 0) & " files...")
            Response.Flush()
            For Each fi In f
                If al.Contains(fi.Name) = False Then
                    Try
                        fi.Delete()
                    Catch
                    End Try
                End If
            Next
            Erase f
            Response.Write("Done.<br>")
            al.Clear()
        End If
        If Not Request("rp") Is Nothing Then
            Response.Write("Resetting popularity indices...")
            Response.Flush()
            cmd.CommandText = "update file_cache set hits = hits + score, score = 0"
            cmd.ExecuteNonQuery()
            Response.Write("Done.<br>")
        End If
        Response.Write("Performing inventory...")
        di = New System.IO.DirectoryInfo(Root)
        f = di.GetFiles()
        Response.Flush()
        cmd.CommandText = "update file_cache set cached = 0"
        cmd.ExecuteNonQuery()
        Dim sb As New System.Text.StringBuilder("update file_cache set cached = 1 where (cached = 1)")
        Dim totalsize As Int64 = 0
        Dim files As Integer = 0
        For Each fi In f
            sb.Append(" or (filename = '")
            sb.Append(fi.Name.Replace("'", "''"))
            sb.Append("')")
            totalsize += fi.Length
            files += 1
            If files / 25 = files \ 25 Then
                Response.Write(".")
                Response.Flush()
            End If
        Next
        Response.Write("Done<br>Updating database...")
        Response.Flush()
        cmd.CommandText = sb.ToString
        cmd.ExecuteNonQuery()
        Response.Write("Done.<hr><pre>")
        Response.Write("CACHED FILES:  " & FormatNumber(files, 0) & Environment.NewLine)
        Response.Write("TOTAL SIZE:    " & FormatNumber(totalsize / 1048576, 1) & " MB" & Environment.NewLine)
        Response.Write("</pre><hr><a href=""?reset=y"" onClick=""return confirm('Are you sure you want to delete ALL files in the cache?');"">Clear Cache</a><br><a href=""?sres=y"">Smart Delete (Keep top 200 and newest 200 files)</a><br><a href=""?rp=y"" onClick=""return confirm('Are you sure you want to reset the cache popularity indices?');"">Reset Popularity Indices</a><br><a href=""?ref=y"">Refresh Statistics</a>")
        Response.Write("</body></html>")
        cmd.Dispose()
        cn.Close()
        cn.Dispose()
        Response.End()
    End Sub

    'Private Sub OLD_Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '    Response.Clear()
    '    Dim dconn As New Odbc.OdbcConnection
    '    dconn.ConnectionString = "DSN=bombness; UID=bombness; PWD=brianjef;"
    '    dconn.Open()
    '    Dim cmd As New Odbc.OdbcCommand
    '    cmd.Connection = dconn
    '    cmd.CommandText = "select distinct pic from users where len(pic) > 2"
    '    Dim items As New ArrayList
    '    Dim dr As Odbc.OdbcDataReader = cmd.ExecuteReader
    '    While dr.Read
    '        items.Add(CStr(dr(0)) & ".me")
    '    End While
    '    dr.Close()

    '    cmd.CommandText = "select distinct fname from vore_pics where len(fname) > 2"
    '    dr = cmd.ExecuteReader
    '    While dr.Read
    '        items.Add(CStr(dr(0)) & ".vore")
    '    End While
    '    dr.Close()

    '    cmd.CommandText = "select distinct face from vore_users where len(face) > 2"
    '    dr = cmd.ExecuteReader
    '    While dr.Read
    '        items.Add(CStr(dr(0)) & ".vore")
    '    End While
    '    dr.Close()

    '    cmd.CommandText = "select distinct mouth from vore_users where len(mouth) > 2"
    '    dr = cmd.ExecuteReader
    '    While dr.Read
    '        items.Add(CStr(dr(0)) & ".vore")
    '    End While
    '    dr.Close()

    '    cmd.CommandText = "select distinct belly from vore_users where len(belly) > 2"
    '    dr = cmd.ExecuteReader
    '    While dr.Read
    '        items.Add(CStr(dr(0)) & ".vore")
    '    End While
    '    dr.Close()

    '    cmd.CommandText = "select distinct poo from vore_users where len(poo) > 2"
    '    dr = cmd.ExecuteReader
    '    While dr.Read
    '        items.Add(CStr(dr(0)) & ".vore")
    '    End While
    '    dr.Close()

    '    cmd.CommandText = "select distinct guid from pictures where len(guid) > 2"
    '    dr = cmd.ExecuteReader
    '    While dr.Read
    '        items.Add(CStr(dr(0)) & ".pic")
    '    End While
    '    dr.Close()

    '    cmd.CommandText = "select distinct fname from user_pics where len(fname) > 2"
    '    dr = cmd.ExecuteReader
    '    While dr.Read
    '        items.Add(CStr(dr(0)) & ".mo")
    '    End While
    '    dr.Close()

    '    cmd.Dispose()
    '    dconn.Close()
    '    dconn.Dispose()

    '    Dim ws As Int64 = 0
    '    Dim fs As New System.IO.DirectoryInfo("\\premfs16\sites\premium16\bombness\data")
    '    Dim files() As System.IO.FileInfo = fs.GetFiles("*.*")
    '    Dim fi As System.IO.FileInfo
    '    For Each fi In files
    '        If items.Contains(fi.Name) = False And Not (fi.name.endswith("png") Or fi.name.endswith("jpg") Or fi.name.endswith("jpeg")) Then
    '            Response.Write("<a href=""serve-file.aspx?password=eggbert&fname=" & fi.Name & """>" & fi.Name & "</a>")
    '            ws += fi.Length
    '            Response.Write("<br>")
    '        End If
    '    Next
    '    Response.Write("<b>Total Size: " & FormatNumber(ws / 1024, 2) & " KB</b>")
    '    Response.End()
    'End Sub
	
</script>