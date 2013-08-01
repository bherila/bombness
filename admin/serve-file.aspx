<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="iso-8859-1" %>
<script runat="server">
Sub Page_Load(Src As Object, E As EventArgs)
	If Request("password") = "eggbert" Then
		Response.Clear()
		Response.ContentType = "application/octet-stream"
		Response.AppendHeader("content-disposition", "attachment; filename=" & Request("fname"))
		If Request("fname").EndsWith("pg") Then
			Response.ContentType = "image/pjpeg"
		End If
		Dim IO As New System.IO.FileStream("\\premfs16\sites\premium16\bombness\data\" & Request("fname"), System.IO.FileMode.Open, System.IO.FileAccess.Read)
		Dim b(IO.Length) As Byte
		IO.Read(b, 0, IO.Length)
		IO.Close()
		Response.OutputStream.Write(b, 0, b.Length)
		Erase b
		Response.End()
	Else
		Response.Write("Error 500 - Internal Server Error")
	End If
End Sub
</script>