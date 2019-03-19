Imports System.Threading
Imports Microsoft.Office.Interop
Public Enum MissingDataEnum
    MissingRemark = 1
    MissingInspDate = 2
End Enum
Public Class FormInspection
   
    WithEvents bs As BindingSource
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim sendEmailThread As New System.Threading.Thread(AddressOf sendEmailWork)
    Dim MyCallBack As System.Net.Mail.SendCompletedEventHandler = AddressOf SendCompletedCallBack

    Dim myController As New InspectionController
    Private Shared myform As FormInspection
    Dim SaveFileDialog1 As New SaveFileDialog
    'Private InspectionDateManual As Date
    Public InspectionDate As Date


    Public Shared Function getInstance()
        If myform Is Nothing Then
            myform = New FormInspection
        ElseIf myform.IsDisposed Then
            myform = New FormInspection
        End If
        
        Return myform
    End Function


    Private Sub loaddata()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            ToolStripStatusLabel2.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")
        Dim sw As New Stopwatch
        sw.Start()
        Try
            If myController.LoadData() Then
                ProgressReport(4, "InitData")
                sw.Stop()
                'ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
                ProgressReport(1, String.Format("Loading Data.Done! Records Count: {0}. Elapsed Time: {1}:{2}.{3}", myController.BS.Count, Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
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
                        DataGridView1.AutoGenerateColumns = False
                        DataGridView1.DataSource = myController.BS
                        bs = myController.BS
                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee


                End Select
            Catch ex As Exception
                ProgressReport(1, ex.Message)
            End Try
        End If

    End Sub


    Private Sub FormInspection_Load(sender As Object, e As EventArgs) Handles Me.Load
        loaddata()
    End Sub


    Private Sub ToolStripTextBox1_Click(sender As Object, e As EventArgs) Handles ToolStripTextBox1.Click

    End Sub

    Private Sub ToolStripTextBox1_TextChanged(sender As Object, e As EventArgs) Handles ToolStripTextBox1.TextChanged
        myController.ApplyFilter = ToolStripTextBox1.Text
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        loaddata()
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        myController.ApplyFilterAll = ToolStripTextBox2.Text
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.RowIndex >= 0 Then
            If e.ColumnIndex = 25 Then

                Dim drv As DataRowView = DataGridView1.Rows(e.RowIndex).DataBoundItem
                Dim myDialog As DialogHistory = DialogHistory.getInstance
                RemoveHandler myDialog.RefreshInspectionDate, AddressOf doRefreshInspectionDate
                AddHandler myDialog.RefreshInspectionDate, AddressOf doRefreshInspectionDate

                RemoveHandler myDialog.RefreshRemark, AddressOf doRefreshRemark
                AddHandler myDialog.RefreshRemark, AddressOf doRefreshRemark

                myDialog.pono = drv.Row.Item("Purch.Doc.")
                myDialog.item = drv.Row.Item("item")
                myDialog.seqn = drv.Row.Item("seqn")
                myDialog.qty = drv.Row.Item("Quantity")
                'myDialog.Id = drv.Row.Item("id")
                myDialog.Run()
                myDialog.Show()

            End If
        End If
        
    End Sub

    Private Sub bs_ListChanged(sender As Object, e As System.ComponentModel.ListChangedEventArgs) Handles bs.ListChanged
        ProgressReport(1, String.Format("Loading Data.Done! Records Count: {0}", myController.BS.Count))
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        SaveFileDialog1.FileName = String.Format("InspectionReport-{0:yyyyMMdd}.xlsx", Date.Today)
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim myReport As ExportToExcelFile = New ExportToExcelFile(Me, myController.SQLSTRExcel, IO.Path.GetDirectoryName(SaveFileDialog1.FileName), IO.Path.GetFileName(SaveFileDialog1.FileName), AddressOf FormatReport, AddressOf PivotCallback, 1, "\templates\excelTemplate.xltx")
            myReport.Run(Me, New EventArgs)
        End If
    End Sub

    Private Sub Callback()
        'Throw New NotImplementedException
    End Sub
    Sub FormatReport(ByRef sender As Object, ByRef e As EventArgs)
        Dim osheet As Excel.Worksheet = DirectCast(sender, Excel.Worksheet)
        osheet.Columns("A:A").Delete()
    End Sub
   

    Private Sub PivotCallback()
        'Throw New NotImplementedException
    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        ''Get Selected Data   
        ''1.Send email to vendor by Grouping datasource based on selected data

        'Dim myquery = From n In myController.DS.Tables(0).AsEnumerable
        '          Where n.Item("Selected")
        '          Group By key = n.Item("Vendor Name") Into Group
        'For Each a In myquery
        '    MessageBox.Show(String.Format("Create Draft for {0}", a.key))
        '    For Each mydata In a.Group
        '        MessageBox.Show(mydata.Item("Purch.Doc."))
        '    Next
        'Next

        ''2.Send email for internal 
        'Dim myqueryInternal = From n In myController.DS.Tables(0).AsEnumerable
        '         Where n.Item("Selected")                
        'For Each a In myqueryInternal
        '    MessageBox.Show(String.Format("Create Draft for Internal Po Number {0}", a.Item("Purch.Doc.")))
        'Next
        Me.Validate()
        Dim myDialog As New DialogAskDate
        If myDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            'InspectionDateManual = myDialog.DTInscpectionDate.Value
            InspectionDate = myDialog.DTInscpectionDate.Value
            sendEmail()
        End If

    End Sub

    Public Sub sendEmail()
        If Not sendEmailThread.IsAlive Then
            'ToolStripStatusLabel1.Text = ""
            sendEmailThread = New Thread(AddressOf sendEmailWork)
            sendEmailThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub


    Sub sendEmailWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Sending Emails..")
        Try
            Dim GoSend As New SendEmailConfirmation(Me, MyCallBack)
            If GoSend.runEWS(myController.DS.Tables(0), InspectionDate) Then
                ProgressReport(1, String.Format("Sending Emails Done!"))
                ProgressReport(5, "Continuous")
            End If
        Catch ex As Exception

            ProgressReport(1, "Sending Emails. Error::" & ex.Message)
            ProgressReport(5, "Continuous")
        End Try
    End Sub

    Sub SendCompletedCallBack(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
        Dim token As String = CStr(e.UserState)
        ProgressReport(2002, String.Format("{0},Email Sent.", e.UserState))
        ProgressReport(2005, String.Format("{0}", 1))
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
    End Sub

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        Dim myform = New FormMissingRemarks(MissingDataEnum.MissingRemark)
        myform.Show()
    End Sub
    Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs) Handles ToolStripButton7.Click
        Dim myform = New FormMissingRemarks(MissingDataEnum.MissingInspDate)
        myform.Show()
    End Sub

    Private Sub doRefreshInspectionDate(po As Long, item As Integer, seq As Integer, NewDate As Date)
        Dim mykey(2) As Object
        mykey(0) = po
        mykey(1) = item
        mykey(2) = seq
        Dim myresult = myController.DS.Tables(0).Rows.Find(mykey)
        If Not IsNothing(myresult) Then
            Dim DOW = {"Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"}
            myresult.Item("Inspection Date") = NewDate
            myresult.Item("dow") = DOW(NewDate.DayOfWeek)
        End If
        DataGridView1.Invalidate()
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        
    End Sub

    Private Sub doRefreshRemark(po As Long, item As Integer, SeqN As Integer, myRemark As String)
        Dim mykey(2) As Object
        mykey(0) = po
        mykey(1) = item
        mykey(2) = SeqN
        Dim myresult = myController.DS.Tables(0).Rows.Find(mykey)
        If Not IsNothing(myresult) Then
            myresult.Item("remarks") = myRemark
        End If
        DataGridView1.Invalidate()
    End Sub


End Class