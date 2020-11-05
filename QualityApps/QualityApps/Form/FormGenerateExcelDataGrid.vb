Imports System.Threading

Public Class FormGenerateExcelDataGrid
    Public Shared myForm As FormGenerateExcelDataGrid
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myController As New GenerateExcelController
    Private TotalProcess As Integer
    Dim MyCallBack As System.Net.Mail.SendCompletedEventHandler = AddressOf SendCompletedCallBack
    Dim MyDate As Date = Date.Today
    Dim Mydate2 As Date = Date.Today

    Public Shared Function getInstance()
        If myForm Is Nothing Then
            myForm = New FormGenerateExcelDataGrid
        ElseIf myForm.IsDisposed Then
            myForm = New FormGenerateExcelDataGrid
        End If
        Return myForm
    End Function

    Private Sub loaddata()
        If Not myThread.IsAlive Then
            'ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub FormGenerateExcelDataGrid_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If TotalProcess > 0 Then
            e.Cancel = True
            MessageBox.Show("Please wait until the process is finished.")
        End If
    End Sub


    Private Sub FormGenerateExcelDataGrid_Load(sender As Object, e As EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")
        Try
            If myController.loaddata() Then
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
                    Case 2005
                        TotalProcess = TotalProcess - 1

                End Select
            Catch ex As Exception
                ProgressReport(1, ex.Message)
            End Try
        End If

    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            UcDataGridExcel1.Validate()
            'Get Processing Date
            Dim AskDate As New DialogAskDate
            AskDate.Label1.Text = "Processing Date"
            If AskDate.ShowDialog = Windows.Forms.DialogResult.OK Then
                MyDate = AskDate.DTInscpectionDate.Value.Date                
                myThread = New Thread(AddressOf DoCreateExcel)
                myThread.SetApartmentState(ApartmentState.STA)
                myThread.Start()
            End If
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub DoCreateExcel()
        'Dim myGenerateExcel As New GenerateExcel(Me)
        Dim myGenerateExcel As New GenerateExcel(Me, MyDate)
        myGenerateExcel.Run(myController.DS)
    End Sub

    Private Sub DoCreateEmail()
        Dim myGenerateEmail As New GenerateEmail(Me, MyCallBack)
        myGenerateEmail.run(myController.DS)
    End Sub
    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        loaddata()
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            UcDataGridExcel1.Validate()
            myThread = New Thread(AddressOf DoCreateEmail)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub SendCompletedCallBack(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
        Dim token As String = CStr(e.UserState)
        ProgressReport(2002, String.Format("{0},Email Sent.", e.UserState))
        ProgressReport(2005, String.Format("{0}", 1))

    End Sub

   

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click

        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            UcDataGridExcel1.Validate()
            'Dim AskDate As New DialogAskDate
            Dim AskDate As New DialogAskDateRange
            AskDate.Label1.Text = "Processing Date"
            AskDate.Label2.Text = "Deadline Booking Date"
            If AskDate.ShowDialog = Windows.Forms.DialogResult.OK Then
                'MyDate = AskDate.DTInscpectionDate.Value.Date
                MyDate = AskDate.DateTimePicker1.Value.Date
                Mydate2 = AskDate.DateTimePicker2.Value.Date
                myThread = New Thread(AddressOf DoCreateEmailEWS)
                myThread.SetApartmentState(ApartmentState.STA)
                myThread.Start()
            End If

           
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub
   
    Private Sub DoCreateEmailEWS01()
        'Dim myGenerateEmail As New GenerateEmail(Me, MyCallBack)
        Dim myGenerateEmail As New GenerateEmail(Me, MyCallBack, MyDate)
        myGenerateEmail.runEWS(myController.DS)
    End Sub
    Private Sub DoCreateEmailEWS()
        'Dim myGenerateEmail As New GenerateEmail(Me, MyCallBack)
        Dim myGenerateEmail As New GenerateEmail(Me, MyCallBack, MyDate, Mydate2)
        myGenerateEmail.runEWS(myController.DS)
    End Sub
End Class