<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="iso-8859-1" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>System Information</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body,td,th {
	font-family: Tahoma, Verdana, Arial, Helvetica, sans-serif;
	font-size: 8pt;
}
-->
</style></head>
<body>
<script runat="server">
Sub Page_Load(Src As Object, E As EventArgs)
        'On Error Resume Next
        Dim i As Integer
        Response.Clear()
        Response.Write("<b>Machine Name: </b>" & System.Environment.MachineName)
        Response.Write("<br><br>")
        Response.Write("<b>Processor Information:</b>")
        Response.Flush()
        With Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Hardware\Description\System\CentralProcessor")
            Dim cpu() As String = .GetSubKeyNames()
            For i = 0 To cpu.Length - 1
                Response.Write("<br>Processor " & i.ToString & ":")
                With .OpenSubKey(i.ToString)
                    Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
                    Response.Write(.GetValue("~MHz", "Unavailable") & " MHz, ")
                    Response.Write(.GetValue("ProcessorNameString", ""))
                    Response.Flush()
                    .Close()
                End With
            Next
            .Close()
        End With
        Response.Write("<br><br><b>System Uptime: </b>")

        Dim ts As New TimeSpan(System.Environment.TickCount)
        Response.Write(ts.ToString)
        Response.Write("<br><br>")
        Response.Flush()

        Dim PC As New System.Diagnostics.PerformanceCounter
        PC.ReadOnly = True

        'PC.CategoryName = "Active Server Pages"
        'PC.CounterName = "Requests/Sec"
        'Response.Write("<b>ASP Requests/Sec: </b>" & FormatNumber(PC.NextValue, 0, TriState.True, TriState.False, TriState.True) & "<br>")
        'Response.Flush()

        'PC.CategoryName = "Memory"
        'PC.CounterName = "Available KBytes"
        'Response.Write("<b>Available Memory: </b>" & FormatNumber(PC.NextValue, 0, TriState.True, TriState.False, TriState.True) & " KB<br>")
        'Response.Flush()

        'PC.CategoryName = "Memory"
        'PC.CounterName = "Available KBytes"
        'Response.Write("<b>Available Memory: </b>" & FormatNumber(PC.NextValue, 0, TriState.True, TriState.False, TriState.True) & " KB<br>")
        'Response.Flush()

        Response.Write("<br><b>Operating System: </b>" & System.Environment.OSVersion.ToString())
        Response.Flush()

        Response.Write("<br><b>Logical Drives:</b>")
        Dim drives() As String = System.Environment.GetLogicalDrives()
        For i = 0 To drives.Length - 1
            Response.Write("<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & drives(i))
            Response.Flush()
        Next

        PC.Dispose()
End Sub
</script>
</body>
</html>
