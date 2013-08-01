<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="iso-8859-1" %>
<script runat="server">
	Sub Page_Load(Src As Object, E As EventArgs)
        Dim width As Integer = (Request("b.x") - Request("a.x"))
        Dim height As Integer = (Request("b.y") - Request("a.y"))
        Dim fname As String = "\\premfs16\sites\premium16\bombness\data\" & Request("fname")
        If Request("fname").EndsWith("png") Then
            Response.Write("access denied")
        ElseIf (width < 0) Or (height < 0) Then
            Response.Write("invalid coordinates")
        Else
            Dim bmp As System.Drawing.Bitmap = System.Drawing.Bitmap.FromFile(fname)
            Dim target As New System.Drawing.Bitmap(width, height)
            Dim g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(target)
            Dim x As Integer = 0 - CInt(Request("a.x"))
            Dim y As Integer = 0 - CInt(Request("a.y"))
            bmp.SetResolution(72, 72)
            Dim dest As New System.Drawing.RectangleF(0, 0, width, height)
            Dim source As New System.Drawing.RectangleF(Request("a.x"), Request("a.y"), width, height)
            g.DrawImage(bmp, dest, source, System.Drawing.GraphicsUnit.Pixel)
            g.DrawImage(bmp, x, y, width, height)
            bmp.Dispose()
            If Request("prev") = 1 Then
                target.Save(Response.OutputStream, System.Drawing.Imaging.ImageFormat.Jpeg)
            Else
                SaveJPGWithCompressionSetting(target, fname, 80)
                Response.Write("done!")
            End If
        End If
	End Sub

    Private Sub SaveJPGWithCompressionSetting(ByVal image As _
    System.Drawing.Image, ByVal szFileName As String, ByVal lCompression _
    As Long)
        Dim eps As System.Drawing.Imaging.EncoderParameters = New System.Drawing.Imaging.EncoderParameters(1)
        eps.Param(0) = New System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, _
            lCompression)
        Dim ici As System.Drawing.Imaging.ImageCodecInfo = GetEncoderInfo("image/jpeg")
        image.Save(szFileName, ici, eps)
    End Sub

    Private Function GetEncoderInfo(ByVal mimeType As String) _
        As System.Drawing.Imaging.ImageCodecInfo
        Dim j As Integer
        Dim encoders As System.Drawing.Imaging.ImageCodecInfo()
        encoders = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
        For j = 0 To encoders.Length
            If encoders(j).MimeType = mimeType Then
                Return encoders(j)
            End If
        Next j
        Return Nothing
    End Function
</script>