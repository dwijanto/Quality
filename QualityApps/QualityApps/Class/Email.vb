Imports System.Net.Mail
Imports System.Threading

Public Class Email
    Implements IDisposable

    Public Property sendto As String
    Public Property sender As String
    Public Property subject As String
    Public Property body As String
    Public Property cc As String
    Public Property bcc As String
    Public Property isBodyHtml As Boolean
    Public Property attachmentlist As List(Of String)
    Public Property htmlView As AlternateView = Nothing
    Public Property plainView As AlternateView = Nothing
    Dim MyCallBack As System.Net.Mail.SendCompletedEventHandler = AddressOf SendCompletedCallBack
    Dim mailsent As Boolean
    Dim myKey As Object
    Dim mythread As New Thread(AddressOf DoWorkKey)
    Dim myForm As Object
    Public Sub New()

    End Sub

    Public Sub New(ByVal Parent As Object)
        Me.myForm = Parent
    End Sub

    Public Sub New(ByVal CallBack As System.Net.Mail.SendCompletedEventHandler)
        'Me.MyCallBack = CallBack
    End Sub

    Public Sub RunWithKey(ByRef sender As System.Object, ByVal e As GenerateEmailWithKeyEventArgs)
        If Not mythread.IsAlive Then
            Try
                Me.myKey = e.myKey
                mythread = New System.Threading.Thread(New ThreadStart(AddressOf DoWorkKey))
                mythread.SetApartmentState(ApartmentState.MTA)
                mythread.Start()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub
    Sub DoWorkKey()
        send(myKey)
    End Sub
    Public Function send(ByRef message As String) As Boolean
        myForm.ProgressReport(2002, String.Format("{0},Sending email to {0} ...", myKey))
        Dim myresult As Boolean = False
        Try
            Using mail As New MailMessage
                'mail.ReplyToList.Add("sebasiainspectionbk@groupeseb.com")
                'mail.ReplyToList.Add("dwijanto@yahoo.com")
                Dim mysendtos() = sendto.Split(";")
                For Each mysendto In mysendtos
                    If mysendto.Length <> 0 Then
                        mail.To.Add(Trim(mysendto))
                    End If

                Next
                mail.From = New MailAddress(Trim(sender))

                If cc <> "" Then
                    Dim myccs() = cc.Split(";")

                    For Each mycc In myccs
                        If mycc.Length <> 0 Then
                            mail.CC.Add(Trim(mycc))
                        End If

                    Next
                End If
                If bcc <> "" Then
                    Dim mybccs() = bcc.Split(";")
                    For Each mybcc In mybccs
                        If mybcc.Length <> 0 Then
                            mail.Bcc.Add(Trim(mybcc))
                        End If
                    Next
                End If
                ' End If

                mail.Subject = subject
                mail.Body = body
                mail.IsBodyHtml = isBodyHtml

                If Not IsNothing(htmlView) Then
                    mail.AlternateViews.Add(htmlView)
                End If
                If Not IsNothing(plainView) Then
                    mail.AlternateViews.Add(plainView)
                End If

                If Not IsNothing(attachmentlist) Then
                    For Each mystr In attachmentlist
                        mail.Attachments.Add(New Attachment(mystr))
                    Next
                End If

                Using smtp = New SmtpClient(My.Settings.smtpclient)
                    'MessageBox.Show(My.Settings.smtpclient)
                    smtp.Send(mail)
                    myForm.ProgressReport(2002, String.Format("{0},Done. Email Sent", myKey))
                    myForm.progressreport(2005, String.Format("{0}", 1))
                End Using
            End Using


            myresult = True
        Catch ex As Exception
            message = ex.Message
        Finally

        End Try
        Return myresult
    End Function

    Public Function sendAsync(ByRef message As String) As Boolean
        Dim myresult As Boolean = False
        Try
            Using smtp = New SmtpClient(My.Settings.smtpclient)
                'AddHandler smtp.SendCompleted, AddressOf SendCompletedCallBack
                AddHandler smtp.SendCompleted, MyCallBack
                Using mail As New MailMessage
                    Dim mysendtos() = sendto.Split(";")
                    For Each mysendto In mysendtos
                        mail.To.Add(Trim(mysendto))
                    Next
                    mail.From = New MailAddress(Trim(sender))

                    If cc <> "" Then
                        Dim myccs() = cc.Split(";")
                        For Each mycc In myccs
                            mail.CC.Add(Trim(mycc))
                        Next
                    End If
                    If bcc <> "" Then
                        Dim mybccs() = bcc.Split(";")
                        For Each mybcc In mybccs
                            mail.Bcc.Add(Trim(mybcc))
                        Next
                    End If

                    mail.Subject = subject
                    mail.Body = body
                    mail.IsBodyHtml = isBodyHtml
                    If Not IsNothing(attachmentlist) Then
                        For Each mystr In attachmentlist
                            mail.Attachments.Add(New Attachment(mystr))
                        Next
                    End If
                    'Dim userstate As String = "testMessage1"
                    'smtp.SendAsync(mail, userstate)
                    smtp.SendAsync(mail, message)
                End Using
            End Using

            myresult = True
        Catch ex As Exception
            message = ex.Message
        Finally

        End Try
        Return myresult
    End Function

    Private Sub SendCompletedCallBack(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
        Dim token As String = CStr(e.UserState)

        If e.Cancelled Then
            Debug.WriteLine("[{0}] Send canceled.", token)
        End If
        If e.Error IsNot Nothing Then
            Debug.WriteLine("[{0}] {1}", token, e.Error.ToString())
        Else
            Debug.WriteLine("Message sent.")
        End If
        mailsent = True
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

 



End Class
Public Class GenerateEmailWithKeyEventArgs
    Inherits EventArgs
    Public Property myKey As Object
End Class