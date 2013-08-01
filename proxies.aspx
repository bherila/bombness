<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="iso-8859-1" %>
<script runat="server">
    Sub Page_Load(ByVal Src As Object, ByVal E As EventArgs)
        'Build Proxy List
        Response.Clear()
        Response.ContentType = "text\plain"
        Dim res As String = GetURL("http://www.samair.ru/proxy/fresh-proxy-list.htm")
        Dim regex As New System.Text.RegularExpressions.Regex("\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b")
        Dim result As System.Text.RegularExpressions.MatchCollection = regex.Matches(res)
        Dim i As Integer
        For i = 0 To result.Count - 1
            Response.Write(result(i).Value)
            Response.Write(Environment.NewLine)
            If i > 50 Then Exit For
        Next
        regex = Nothing
        res = Nothing
        result = Nothing
    End Sub

    Function GetURL(ByVal uri As String) As String
        Dim RQ As New System.Net.WebClient
        Dim b() As Byte = RQ.DownloadData(uri)
        Dim st As String = System.Text.Encoding.ASCII.GetString(b)
        RQ.Dispose()
        Array.Clear(b, 0, b.Length)
        Return st
    End Function
</script>