Imports System.Threading
Public Class FormSendEmailConfirmationCP
    Public Shared myform As FormSendEmailConfirmationCP

    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Public myController As New SendEmailControllerCP
    Private TotalProcess As Integer
    Private TotalRowExecuted As Integer
    Dim MyCallBack As System.Net.Mail.SendCompletedEventHandler = AddressOf SendCompletedCallBack
    'Private StartDate As Date
    'Private enddate As Date

    Public Inspectiondate As Date

    Public Shared Function getInstance()
        If myform Is Nothing Then
            myform = New FormSendEmailConfirmationCP
        ElseIf myform.IsDisposed Then
            myform = New FormSendEmailConfirmationCP
        End If
        Return myform
    End Function

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        UcDataGridExcel1.DataGridView1.Columns("Column5").Visible = False 'Column5 HWND

    End Sub

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")
        Try
            If myController.LoadData(Inspectiondate) Then
                ProgressReport(4, "InitData")
                ProgressReport(1, String.Format("Loading Data.Done! Records Count:{0}", myController.BS.Count))
                ProgressReport(5, "Continuous")
            End If
        Catch ex As Exception
            ProgressReport(1, "Loading Data. Error::" & ex.Message)
            ProgressReport(5, "Continuous")
        End Try
    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Try
                Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
                Me.Invoke(d, New Object() {id, message})
            Catch ex As Exception
                ProgressReport(1, ex.Message)
            End Try

        Else
            Try
                Select Case id
                    Case 1
                        ToolStripStatusLabel1.Text = message
                    Case 2
                        ToolStripStatusLabel1.Text = message
                    Case 4
                        'BindingData
                        UcDataGridExcel1.BindingControl(Me, myController.BS)
                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                    Case 2002
                        myController.showMessage(message, "message")
                        UcDataGridExcel1.DataGridView1.Invalidate()
                    Case 2001
                        myController.showMessage(message, "hwnd")
                        UcDataGridExcel1.DataGridView1.Invalidate()
                    Case 2003
                        myController.showMessage(message, "status")
                        UcDataGridExcel1.DataGridView1.Invalidate()


                    Case 2004
                        TotalProcess = message
                        TotalRowExecuted = TotalProcess
                    Case 2005
                        TotalProcess = TotalProcess - 1
                        If TotalProcess = 0 Then
                            'All Completed if TotalRowExecuted eq mycontroller.ds.tables(0).rows.count
                            If TotalRowExecuted = myController.DS.Tables(0).Rows.Count Then
                                Dim myParam As New ParamAdapter
                                myParam.UpdateLastSendDate(myController.enddate)
                            End If

                        End If
                End Select
            Catch ex As Exception
                ProgressReport(1, ex.Message)
            End Try
        End If

    End Sub
    Sub SendCompletedCallBack(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
        Dim token As String = CStr(e.UserState)
        ProgressReport(2002, String.Format("{0},Email Sent.", e.UserState))
        ProgressReport(2005, String.Format("{0}", 1))

    End Sub

    Private Sub FormSendEmailConfirmation_Load(sender As Object, e As EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Private Sub loaddata()
        If Not myThread.IsAlive Then
            'ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click

        If Not myThread.IsAlive Then

            ToolStripStatusLabel1.Text = ""
            UcDataGridExcel1.Validate()
            myThread = New Thread(AddressOf DoCreateEmailEWS)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub DoCreateEmailEWS()
        'Dim myGenerateEmail As New SendEmailConfirmation(Me, MyCallBack, myController.startdate, myController.enddate)
        Dim myGenerateEmail As New SendEmailConfirmationCP(Me, MyCallBack)
        myGenerateEmail.runEWS(myController.DS)

    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        loaddata()
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim myform = FormTimeStampSendingReportCP.getInstance
        myform.show()
    End Sub
End Class