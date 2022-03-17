Imports System

Module Program
    Sub Main(args As String())
        BateAquiHost.Services.SmtpService.Send("EMAIL@PARA.COM.BR", "TEste", "TEste VB.NET", "Teste VB.Net")
        Console.WriteLine("Hello World!")
    End Sub
End Module
