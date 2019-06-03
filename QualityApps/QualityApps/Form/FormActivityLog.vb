Imports System.Threading

Public Class FormActivityLog
    Dim myRBAC As DbManager = New DbManager
    Private Shared myform As FormActivityLog
    Dim myController As ActivityLogController
    Dim criteria As String = String.Empty
    Private dtPicker1 As New DateTimePicker
    Private dtPicker2 As New DateTimePicker
    Private lbl1 As New Label
    Private WithEvents checkbox1 As New CheckBox
    Private myIdentity As UserController = User.getIdentity



    Public Shared Function getInstance()
        If myform Is Nothing Then
            myform = New FormActivityLog
        ElseIf myform.IsDisposed Then
            myform = New FormActivityLog
        End If
        Return myform
    End Function
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim drv As DataRowView

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        InitializeMenubar()

    End Sub

    Private Sub RefreshMYInterface()
        DataGridView1.Invalidate()
    End Sub

    Sub DoWork()
        myController = New ActivityLogController

        Dim daterange As String = String.Empty
        Dim conjunction As String = String.Empty
        If checkbox1.Checked Then
            daterange = String.Format("activitydate >= '{0:yyyy-MM-dd}' and activitydate <= '{1:yyyy-MM-dd}'", dtPicker1.Value, dtPicker2.Value)
            conjunction = "and"
        End If


        If Not IsNothing(myRBAC.getAssignment("Key User", User.getId)) Then ' is Key User
            criteria = String.Format("where {0}", daterange)
        Else
            criteria = String.Format(" where {0} {1} u.userid in (select quality.getsubordinate('{2}'))", daterange, conjunction, myIdentity.userid.ToLower)
        End If

        'If User.can("View Activity Log All Data") Then
        '    criteria = String.Format("where {0}", daterange)
        'Else
        '    criteria = String.Format(" where {0} {1} u.userid in (select quality.getsubordinate('{2}'))", daterange, conjunction, myIdentity.userid.ToLower)
        'End If

        Try
            ProgressReport(1, "Loading...Please wait.")
            If myController.loaddata(criteria) Then
                ProgressReport(4, "Init Data")
            End If
            ProgressReport(1, String.Format("Loading...Done. Records {0}", myController.BS.Count))
        Catch ex As Exception
            ProgressReport(1, ex.Message)
        End Try

    End Sub



    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    ToolStripStatusLabel1.Text = message
                Case 4
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.DataSource = myController.BS



                    UcUserInfo1.BindingObject()
            End Select
        End If
    End Sub
    Public Sub showTx(ByVal tx As TxEnum)
        If Not myThread.IsAlive Then
            Dim drv As DataRowView = Nothing
            Select Case tx
                Case TxEnum.NewRecord
                    drv = myController.GetNewRecord
                    drv.Item("activitydate") = Today.Date
                    drv.Item("postingdate") = Now
                    drv.Item("userid") = myIdentity.userid
                    drv.Item("username") = myIdentity.username
                Case TxEnum.UpdateRecord
                    drv = myController.GetCurrentRecord
                Case TxEnum.CopyRecord
                    Dim drv2 = myController.GetCurrentRecord
                    drv = myController.GetNewRecord
                    drv.Item("activitydate") = drv2.Item("activitydate")
                    drv.Item("timesession") = drv2.Item("timesession")
                    drv.Item("vendorcode") = drv2.Item("vendorcode")
                    drv.Item("vendorcodename") = drv2.Item("vendorcodename")
                    drv.Item("projectname") = drv2.Item("projectname")
                    drv.Item("activityid") = drv2.Item("activityid")
                    drv.Item("activityname") = drv2.Item("activityname")
                    drv.Item("timesessiondesc") = drv2.Item("timesessiondesc")
                    drv.Item("userid") = drv2.Item("userid")
                    drv.Item("remark") = drv2.Item("remark")
            End Select
            drv.Item("modifiedby") = myIdentity.userid

            Dim myform = New DialogActivityLogNew(drv, myController.GetVendorBS, myController.GetActivityBS, myController.GetTimeSessionBS, myController.GetDataset)
            RemoveHandler myform.RefreshInterface, AddressOf RefreshMYInterface
            AddHandler myform.RefreshInterface, AddressOf RefreshMYInterface
            If myform.ShowDialog() = Windows.Forms.DialogResult.Cancel Then
                If drv.Row.RowState = DataRowState.Detached Or drv.Row.RowState = DataRowState.Added Then
                    myController.RemoveAt(myController.GetCurrentPosition)
                End If
            End If
        End If
    End Sub

    Public Sub showTxOld(ByVal tx As TxEnum)
        If Not myThread.IsAlive Then
            Dim drv As DataRowView = Nothing
            Select Case tx
                Case TxEnum.NewRecord
                    drv = myController.GetNewRecord
                    drv.Item("activitydate") = Today.Date
                    drv.Item("userid") = myIdentity.userid
                Case TxEnum.UpdateRecord
                    drv = myController.GetCurrentRecord
                Case TxEnum.CopyRecord
                    Dim drv2 = myController.GetCurrentRecord
                    drv = myController.GetNewRecord
                    drv.Item("activitydate") = drv2.Item("activitydate")
                    drv.Item("postingdate") = Today.Date
                    drv.Item("timesession") = drv2.Item("timesession")
                    drv.Item("vendorcode") = drv2.Item("vendorcode")
                    drv.Item("vendorcodename") = drv2.Item("vendorcodename")
                    drv.Item("projectname") = drv2.Item("projectname")
                    drv.Item("activityid") = drv2.Item("activityid")
                    drv.Item("activityname") = drv2.Item("activityname")
                    'drv.Item("timesessiondesc") = drv2.Item("timesessiondesc")
                    drv.Item("userid") = drv2.Item("userid")
                    drv.Item("remark") = drv2.Item("remark")
            End Select

            drv.Item("modifiedby") = myIdentity.userid
            Dim myform = New DialogActivityLog(drv, myController.GetVendorBS, myController.GetActivityBS, myController.GetTimeSessionBS, myController.BS2, myController.GetDataset)
            RemoveHandler myform.RefreshInterface, AddressOf RefreshMYInterface
            AddHandler myform.RefreshInterface, AddressOf RefreshMYInterface
            If myform.ShowDialog() = Windows.Forms.DialogResult.Cancel Then
                If drv.Row.RowState = DataRowState.Detached Or drv.Row.RowState = DataRowState.Added Then
                    myController.RemoveAt(myController.GetCurrentPosition)
                End If
            End If
        End If
    End Sub

    Private Sub AddNewToolStripButton1_Click(sender As Object, e As EventArgs) Handles AddToolStripButton.Click
        showTx(TxEnum.NewRecord)
    End Sub

    Private Sub DeleteToolStripButton2_Click(sender As Object, e As EventArgs) Handles DeleteToolStripButton.Click
        If Not IsNothing(myController.GetCurrentRecord) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    If Me.validate(drv) Then
                        myController.RemoveAt(drv.Index)
                    End If


                Next
            End If
        End If
    End Sub

    Public Overloads Function validate(dgvr As DataGridViewRow) As Boolean
        Dim myret As Boolean = False
        myController.BS.Position = dgvr.Index
        Dim drv = myController.BS.Current
        If IsNothing(myRBAC.getAssignment("Key User", User.getId)) Then
            'Not Key User
            Dim ActivityDate As Date = drv.Row.Item("activitydate")
            drv.Row.RowError = ""
            If ActivityDate.Month = Today.Month Then
                myret = True
            ElseIf ActivityDate.Month + 1 = Month(Today) Then
                If Today.Date.Day <= myController.DS.Tables("Cutoff").Rows(0).Item("ivalue") Then
                    myret = True
                Else
                    drv.Row.RowError = String.Format("This Activity date '{0:dd-MMM-yyyy}' has been closed. Cannot be deleted. Please contact Key User!", ActivityDate)
                End If
            Else
                drv.Row.RowError = String.Format("This Activity date '{0:dd-MMM-yyyy}' has been closed. Cannot be deleted. Please contact Key User!", ActivityDate)
            End If
        Else
            myret = True 'Key User
        End If
        Return myret

    End Function

    Private Sub FormActivityLog_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Dim abc = myController.DS.GetChanges()
        If Not IsNothing(abc) Then

            Select Case MessageBox.Show("Save unsaved records?", "Unsaved records", MessageBoxButtons.YesNoCancel)
                Case Windows.Forms.DialogResult.Yes
                    If Me.Validate Then
                        CommitToolStripButton.PerformClick()
                    Else
                        e.Cancel = True
                    End If

                Case Windows.Forms.DialogResult.Cancel
                    e.Cancel = True
            End Select
        End If
    End Sub

    Private Sub FormActivity_Load(sender As Object, e As EventArgs) Handles Me.Load
        LoadData()
    End Sub

    Private Sub LoadData()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub RefreshToolStripButton4_Click(sender As Object, e As EventArgs) Handles RefreshToolStripButton.Click
        LoadData()
    End Sub



    Private Sub CommitToolStripButton3_Click(sender As Object, e As EventArgs) Handles CommitToolStripButton.Click
        Me.Validate()
        myController.save()
    End Sub


    Private Sub ToolStripTextBox1_TextChanged(sender As Object, e As EventArgs) Handles ToolStripTextBox1.TextChanged
        Try
            myController.ApplyFilter = ToolStripTextBox1.Text
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub InitializeMenubar()
        checkbox1.Checked = True
        lbl1.Text = "To "
        dtPicker1.CustomFormat = "dd-MMM-yyyy"
        dtPicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        dtPicker1.Size = New Point(100, 20)
        dtPicker1.Value = CDate(String.Format("{0}-{1}-1", Year(Today), Month(Today)))

        dtPicker2.CustomFormat = "dd-MMM-yyyy"
        dtPicker2.Value = CDate(String.Format("{0}-{1}-{2}", Year(Today), Month(Today), DateTime.DaysInMonth(Year(Today), Month(Today))))
        dtPicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        dtPicker2.Size = New Point(100, 20)

        ToolStrip1.Items.Add(New ToolStripControlHost(checkbox1))
        ToolStrip1.Items.Add(New ToolStripControlHost(dtPicker1))
        ToolStrip1.Items.Add(New ToolStripControlHost(lbl1))
        ToolStrip1.Items.Add(New ToolStripControlHost(dtPicker2))
    End Sub

    Private Sub CopyToolStripButton_Click(sender As Object, e As EventArgs) Handles CopyToolStripButton.Click
        If Not IsNothing(myController.GetCurrentRecord) Then
            showTx(TxEnum.CopyRecord)
        Else
            MessageBox.Show("Nothing to copy.")
        End If

    End Sub

    Private Sub checkbox1_CheckedChanged(sender As Object, e As EventArgs) Handles checkbox1.CheckedChanged
        dtPicker1.Enabled = checkbox1.Checked
        dtPicker2.Enabled = checkbox1.Checked
    End Sub

    Private Sub UpdateToolStripButton_Click(sender As Object, e As EventArgs) Handles UpdateToolStripButton.Click
        If Not IsNothing(myController.GetCurrentRecord) Then
            showTx(TxEnum.UpdateRecord)
        Else
            MessageBox.Show("Nothing to update.")
        End If
    End Sub

    Private Sub DataGridView1_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDoubleClick
        If e.RowIndex > -1 Then UpdateToolStripButton.PerformClick()
    End Sub

    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError

    End Sub


End Class