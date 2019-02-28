Imports Microsoft.Exchange.WebServices.Data
Imports System.Threading
Imports System.IO

Public Class EmailEWS
    Public Property Username As String
    Public Property password As String
    Public Property domain As String
    Dim myParam As ParamAdapter = ParamAdapter.getInstance
    Public Property urlEWS As String = myParam.GeturlEWS("urlEWS") ' = "https://mail-eu.seb.com/ews/exchange.asmx"

    Dim emaildict As Dictionary(Of String, String)
    Dim attachmentdict As Dictionary(Of String, String)
    Dim emaillist As List(Of emailData)

    Dim myForm As Object

    Public ToRecepientsE As EmailAddress
    Public ToRecepients As String
    Public CcRecepients As String
    Public BccRecepients As String
    Public Subject As String
    Public BodyMessage As String
    Public AttachmentList As New List(Of String)

    Dim mythread As New Thread(AddressOf DoWorkKey)
    Private myKey As Object

    Public Sub New(ByVal parent As Object)
        Me.myForm = parent        
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
        createemail(myKey)
    End Sub

    Private Function createemail(ByVal myKey As Object) As Boolean
        Dim myret As Boolean = False
        Dim url As String = urlEWS '"https://mail-eu.seb.com/ews/exchange.asmx"

        Dim service As ExchangeService
        Using myservice As New ClassEWS(url, Username, password, domain, False)
            'service = myservice.CreateConnectionAutoDiscover()
            myForm.ProgressReport(2002, String.Format("{0},Create Connection...", myKey))
            service = myservice.CreateConnection()
            Try
                Dim msg As EmailMessage = New EmailMessage(service)
                emaildict = New Dictionary(Of String, String)
                emaillist = New List(Of emailData)
                attachmentdict = New Dictionary(Of String, String)

                If AttachmentList.Count > 0 Then
                    myForm.ProgressReport(2002, String.Format("{0},Add Attachments...", myKey))
                    For Each a In AttachmentList
                        If File.Exists(a) Then
                            addAttachmentList(a, msg)
                        Else
                            Err.Raise(5001, Description:=String.Format("Attachment {0} is not available.", a))
                        End If

                    Next
                End If

                If Not IsNothing(ToRecepients) Then
                    Dim mye As String() = ToRecepients.Split(";")
                    For i = 0 To mye.Count - 1
                        If mye(i).Length > 0 Then
                            msg.ToRecipients.Add(New EmailAddress(mye(i)))
                        End If
                    Next
                End If

                If Not IsNothing(CcRecepients) Then
                    Dim mye As String() = CcRecepients.Split(";")
                    For i = 0 To mye.Count - 1
                        If mye(i).Length > 0 Then
                            msg.CcRecipients.Add(New EmailAddress(mye(i)))
                        End If
                    Next
                    ' msg.CcRecipients.Add(New EmailAddress(CcRecepients))
                    'msg.CcRecipients.Add(CcRecepients)
                End If
                If Not IsNothing(BccRecepients) Then
                    Dim mye As String() = BccRecepients.Split(";")
                    For i = 0 To mye.Count - 1
                        If mye(i).Length > 0 Then
                            msg.BccRecipients.Add(New EmailAddress(mye(i)))
                        End If
                    Next
                    'msg.BccRecipients.Add(New EmailAddress(BccRecepients))
                    'msg.BccRecipients.Add(BccRecepients)
                End If

                msg.Subject = Subject
                msg.Body = New MessageBody(BodyType.HTML, BodyMessage)
                myForm.ProgressReport(2002, String.Format("{0},Sending email...", myKey))
                'Send Email
                msg.SendAndSaveCopy() 'dlie
                'msg.Save()
                'msg.Send()
                'myForm.ProgressReport(1, String.Format("{0} saving draft... {1}", "ID", "ID2"))

                'msg.Save(WellKnownFolderName.SentItems)
                'myForm.ProgressReport(1, String.Format("{0} end saving draft...", "ID", "ID2"))

                myret = True
                myForm.ProgressReport(2002, String.Format("{0},Sending email...Done", myKey))
            Catch ex As Exception
                Logger.log(String.Format("{0},Error found: {1}", myKey, ex.Message))
                myForm.ProgressReport(2002, String.Format("{0},Error found: {1}", myKey, ex.Message))
            Finally
                myForm.ProgressReport(2003, String.Format("{0},{1}", myKey, True))
                myForm.ProgressReport(2005, String.Format("{0}", myKey))
            End Try

        End Using
        Return myret
    End Function

    Public Function createemailSave() As Boolean
        Dim myret As Boolean = False
        Dim url As String = urlEWS '"https://mail-eu.seb.com/ews/exchange.asmx"

        Dim service As ExchangeService
        Using myservice As New ClassEWS(url, Username, password, domain, False)
            'service = myservice.CreateConnectionAutoDiscover()
            myForm.ProgressReport(1, String.Format("{0},Create Connection...", myKey))
            service = myservice.CreateConnection()
            Try
                Dim msg As EmailMessage = New EmailMessage(service)
                emaildict = New Dictionary(Of String, String)
                emaillist = New List(Of emailData)
                attachmentdict = New Dictionary(Of String, String)

                If AttachmentList.Count > 0 Then
                    myForm.ProgressReport(1, String.Format("{0},Add Attachments...", myKey))
                    For Each a In AttachmentList
                        If File.Exists(a) Then
                            addAttachmentList(a, msg)
                        Else
                            Err.Raise(5001, Description:=String.Format("Attachment {0} is not available.", a))
                        End If

                    Next
                End If

                If Not IsNothing(ToRecepients) Then
                    Dim mye As String() = ToRecepients.Split(";")
                    For i = 0 To mye.Count - 1
                        If mye(i).Length > 0 Then
                            msg.ToRecipients.Add(New EmailAddress(mye(i)))
                        End If
                    Next
                End If

                If Not IsNothing(CcRecepients) Then
                    Dim mye As String() = CcRecepients.Split(";")
                    For i = 0 To mye.Count - 1
                        If mye(i).Length > 0 Then
                            msg.CcRecipients.Add(New EmailAddress(mye(i)))
                        End If
                    Next
                    ' msg.CcRecipients.Add(New EmailAddress(CcRecepients))
                    'msg.CcRecipients.Add(CcRecepients)
                End If
                If Not IsNothing(BccRecepients) Then
                    Dim mye As String() = BccRecepients.Split(";")
                    For i = 0 To mye.Count - 1
                        If mye(i).Length > 0 Then
                            msg.BccRecipients.Add(New EmailAddress(mye(i)))
                        End If
                    Next
                    'msg.BccRecipients.Add(New EmailAddress(BccRecepients))
                    'msg.BccRecipients.Add(BccRecepients)
                End If

                msg.Subject = Subject
                msg.Body = New MessageBody(BodyType.HTML, BodyMessage)
                myForm.ProgressReport(1, String.Format("{0},Saving Draft email...", myKey))
                'Save Email

                msg.Save()
                
                myret = True
                myForm.ProgressReport(1, String.Format("{0},Saving Draft email...Done", myKey))
            Catch ex As Exception
                Logger.log(String.Format("{0},Error found: {1}", myKey, ex.Message))
                myForm.ProgressReport(1, String.Format("{0},Error found: {1}", myKey, ex.Message))
            Finally
                'myForm.ProgressReport(2003, String.Format("{0},{1}", myKey, True))
                'myForm.ProgressReport(2005, String.Format("{0}", myKey))
            End Try

        End Using
        Return myret
    End Function

    Private Sub addEmailList(ByVal recepient As String, ByVal recepientname As String)
        Dim myrecepients() As String = recepient.Split(";")
        For i = 0 To myrecepients.Count - 1
            If Not myrecepients(i).Length = 0 Then
                If Not emaildict.ContainsKey(Trim(myrecepients(i))) Then
                    emaildict.Add(Trim(myrecepients(i)), recepientname)
                    'add email
                    emaillist.Add(New emailData With {.email = Trim(myrecepients(i)), .displayname = recepientname})
                End If
            End If

        Next
    End Sub

    Private Sub addAttachmentList(ByVal filename As String, ByRef msg As EmailMessage)
        If Not attachmentdict.ContainsKey(filename) Then
            attachmentdict.Add(filename, filename)
            msg.Attachments.AddFileAttachment(filename)
        End If
    End Sub

    Private Sub hello()

    End Sub

End Class
Public Class emailData
    Public email As String
    Public displayname As String
End Class