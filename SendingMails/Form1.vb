Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Net
Imports System.Net.Mail
Imports System.Data
Imports System.Net.Mail.SmtpClient
Imports System
'Imports System.Web.Mail


Imports System.IO
Imports System.Web




Public Class Form1
    Dim cryRpt As New ReportDocument
    '  Dim pdfFile1 As String = "E\Price_exception - RAF.pdf"

    Dim pdfFile1 As String = Application.StartupPath & "\Pending Order Details Report .pdf"



    Dim Attachment1 As System.Net.Mail.Attachment


    'Variable to store the attachments
    ' Dim Mailmsg As New System.Web.Mail.MailMessage()    'Variable to create the message to send
    Dim SmtpServer As New SmtpClient()

    Dim crtableLogoninfos As New CrystalDecisions.Shared.TableLogOnInfos
    Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo

    Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
    Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
    Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table

    Dim body As String
    Dim mail As New MailMessage()

    Dim FileReaderserver As StreamReader
    Dim g_Datapoolip As String
    Dim g_Datapoolport As String
    Dim g_Datapoolpw As String

    Private Sub serverconfig()


        FileReaderserver = New StreamReader(Application.StartupPath & "\server.txt")
        Dim ioLines As String
        ioLines = ""
        ioLines = FileReaderserver.ReadToEnd

        FileReaderserver.Close()
        FileReaderserver.Dispose()

        Dim Seperate_Lines() As String = ioLines.Split(CChar(vbLf))
        Dim Max As Integer = UBound(Seperate_Lines)
        For MyCounter = 0 To Max

            If Seperate_Lines(MyCounter).Contains("[") = True And Seperate_Lines(MyCounter).Contains("]") = True Then


            ElseIf Seperate_Lines(MyCounter).Contains(";") = True Then


            Else
                If Seperate_Lines(MyCounter).Contains("=") = True Then

                    Dim Parts() As String = Seperate_Lines(MyCounter).Split("=")


                    If Parts(0).Contains("ip") = True Then
                        g_Datapoolip = Parts(1)
                    ElseIf Parts(0).Contains("port") = True Then
                        g_Datapoolport = Parts(1)
                    ElseIf Parts(0).Contains("password") = True Then
                        g_Datapoolpw = Parts(1)



                    End If

                End If
            End If


        Next

    End Sub

    Private Sub sendMail()
        Dim FileReader As StreamReader
        Dim FileReadercc As StreamReader
        'obj.SmtpServer = "10.10.9.100:587"

        serverconfig()

        SmtpServer = New SmtpClient(g_Datapoolip.TrimEnd(), g_Datapoolport.TrimEnd())
        SmtpServer.DeliveryMethod = SmtpDeliveryMethod.Network

        SmtpServer.UseDefaultCredentials = False
        SmtpServer.Credentials = New NetworkCredential("ifsreportdist@renukafoods.com", g_Datapoolpw.TrimEnd())

        SmtpServer.EnableSsl = True


        FileReader = New StreamReader(Application.StartupPath & "\email.txt")
        FileReadercc = New StreamReader(Application.StartupPath & "\emailcc.txt")


        Attachment1 = New Attachment(pdfFile1)


        mail = New MailMessage()
        mail.From = New MailAddress("ifsreportdist@renukafoods.com")
        mail.To.Add(FileReader.ReadToEnd())
        mail.To.Add(FileReadercc.ReadToEnd())
        mail.Subject = "Pending Order Details Report"
        mail.Attachments.Add(Attachment1)
        mail.Body = "This is system generated mail!"
        SmtpServer.Send(mail)

        Me.Close()


    End Sub
  
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If System.IO.File.Exists(pdfFile1) Then

            System.IO.File.Delete(pdfFile1)
        End If
        Call CreatePDF(Application.StartupPath & "\ETD ST Report.rpt", pdfFile1)


        sendMail()


    End Sub

    Private Sub CreatePDF(ByVal path As String, ByVal pdfFile As String)

        With crConnectionInfo
            .ServerName = "10.10.10.50"
            .DatabaseName = "COMINV"
            .UserID = "sa"
            .Password = "Renuka@69"
        End With
        cryRpt.Load(path)

        CrTables = cryRpt.Database.Tables

        For Each CrTable In CrTables
            crtableLogoninfo = CrTable.LogOnInfo
            crtableLogoninfo.ConnectionInfo = crConnectionInfo
            CrTable.ApplyLogOnInfo(crtableLogoninfo)
        Next

        CrystalReportViewer1.ReportSource = cryRpt

        Try
            Dim CrExportOptions As ExportOptions
            Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
            Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions
            CrDiskFileDestinationOptions.DiskFileName = pdfFile

            CrExportOptions = cryRpt.ExportOptions

            With CrExportOptions
                .ExportDestinationType = ExportDestinationType.DiskFile
                .ExportFormatType = ExportFormatType.PortableDocFormat
                .DestinationOptions = CrDiskFileDestinationOptions
                .FormatOptions = CrFormatTypeOptions
            End With
            cryRpt.Export()


        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


    End Sub


End Class
