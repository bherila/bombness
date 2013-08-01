<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="iso-8859-1" %>
<script runat="server">
Sub Page_Load(Src As Object, E As EventArgs)
'try
        Dim m As New System.Web.Mail.MailMessage
        m.Body = Request("body")
        m.To = "bherila@bombness.com"
        m.From = "bherila@bombness.com"
        System.Web.Mail.SmtpMail.SmtpServer = "bombness.com"
        With m
            .Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusing", 2)
            .Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserver", "bombness.com")
            .Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", "bherila@bombness.com")
            .Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", "Benjamin11.")
            .Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", 1)
            .Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserverport", 25)
        End With

        'm.Body = "FROM: " & Request("from") & Environment.NewLine & "PHONE: " & Request("phone") & Environment.NewLine & Request("body")
        m.Subject = "Bombness Networks Contact Form : " & Request("subject")
        Dim i As Integer
            Dim sb As New System.Text.StringBuilder
            For i = 0 To Request.Form.AllKeys.Length - 1
                sb.Append(Request.Form.AllKeys(i))
                sb.Append(Environment.NewLine)
                sb.Append(Request.Form.Item(Request.Form.AllKeys(i)))
                sb.Append(Environment.NewLine)
                sb.Append(Environment.NewLine)
            Next
            m.Body = sb.ToString
        Dim ax As New ArrayList
        For i = 0 To Request.Files.Count - 1
            Try
                Dim tfn As String = "m:\temp\" & New IO.FileInfo(Request.Files(i).FileName).Name
                Request.Files(i).SaveAs(tfn)
                Dim attach As New System.Web.Mail.MailAttachment(tfn)
                m.Attachments.Add(attach)
                ax.Add(tfn)
            Catch
            End Try
        Next
        System.Web.Mail.SmtpMail.Send(m)
        For i = 0 To ax.Count - 1
            Try
                System.IO.File.Delete(ax(i))
            Catch
            End Try
        Next
        Response.Redirect(Request("goto"))
'catch
'        Response.Redirect(Request("errorgoto"))
'end try
End Sub
</script>