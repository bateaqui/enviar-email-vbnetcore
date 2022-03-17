Imports System.Globalization
Imports System.IO
Imports MailKit.Net.Smtp
Imports MimeKit

Namespace BateAquiHost.Services
    Public Class SmtpService
        Public Shared Function Send(ByVal para As String, ByVal paraTitulo As String, ByVal assunto As String, ByVal conteudo As String, ByVal Optional attachments As List(Of String) = Nothing) As Boolean
            Return Send("SEU-EMAIL@SEU-DOMINIO", "SEU-NOME", para, paraTitulo, assunto, conteudo, attachments)
        End Function

        Public Shared Function Send(ByVal de As String, ByVal deTitulo As String, ByVal para As String, ByVal paraTitulo As String, ByVal assunto As String, ByVal conteudo As String, ByVal Optional attachments As List(Of String) = Nothing) As Boolean
            Dim memoryStream As Stream = Nothing

            Try
                Dim FromAddress As String = de
                Dim FromAdressTitle As String = deTitulo
                Dim ToAddress As String = para
                Dim ToAdressTitle As String = paraTitulo
                Dim Subject As String = assunto
                Dim BodyContent As String = conteudo
                Dim SmtpServer As String = "bateaquihost.com.br"
                Dim SmtpPortNumber As Integer = 465
                Dim mimeMessage = New MimeMessage()
                mimeMessage.From.Add(New MailboxAddress(FromAdressTitle, FromAddress))

                For Each item In ToAddress?.Split(";"c)
                    If Not String.IsNullOrEmpty(item?.Trim()) Then mimeMessage.[To].Add(New MailboxAddress(ToAdressTitle, item?.ToLower(CultureInfo.InvariantCulture).Trim()))
                Next

                mimeMessage.Subject = Subject
                Dim multipart = New Multipart("mixed")
                multipart.Add(New TextPart(MimeKit.Text.TextFormat.Html) With {
                    .Text = BodyContent
                })

                If attachments IsNot Nothing Then

                    For Each item In attachments

                        If item.StartsWith("http", StringComparison.InvariantCulture) Then

                            Using webclient = New System.Net.WebClient()
                                Dim downloadDataTaskAsync = webclient.DownloadDataTaskAsync(item)
                                downloadDataTaskAsync.Wait()
                                Dim bytes = downloadDataTaskAsync.Result
                                memoryStream = New MemoryStream(bytes)
                            End Using
                        Else
                            memoryStream = File.OpenRead(item)
                        End If

                        Dim attachmentExtension = item.Split("."c)(item.Split("."c).Length - 1).Trim().ToLowerInvariant()
                        Dim attachment = New MimePart("application", attachmentExtension) With {
                            .Content = New MimeContent(memoryStream, ContentEncoding.[Default]),
                            .ContentDisposition = New ContentDisposition(ContentDisposition.Attachment),
                            .ContentTransferEncoding = ContentEncoding.Base64,
                            .FileName = item.Split("/"c).LastOrDefault().Split("\"c).LastOrDefault()
                        }
                        multipart.Add(attachment)
                    Next
                End If

                mimeMessage.Body = multipart

                Using client = New SmtpClient()
                    client.Connect(SmtpServer, SmtpPortNumber, True)
                    client.Authenticate("SEU-EMAIL@SEU-DOMINIO", "SUA-SENHA")
                    client.SendAsync(mimeMessage).Wait()
                    client.Disconnect(True)
                End Using

                Return True
            Catch ex As Exception
                Return False
            Finally
                If memoryStream IsNot Nothing Then memoryStream.Dispose()
            End Try
        End Function
    End Class
End Namespace
