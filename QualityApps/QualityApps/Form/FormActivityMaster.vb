Imports System.Threading
Public Class FormActivityMaster
    Private Shared myform As FormActivityMaster
    Dim myController As ActivityMasterController
    Public Teams As New List(Of Team)
    Dim TEAMBS As BindingSource
    Public Shared Function getInstance()
        If myform Is Nothing Then
            myform = New FormActivityMaster
        ElseIf myform.IsDisposed Then
            myform = New FormActivityMaster
        End If
        Return myform
    End Function
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)

    'Dim UserInfo1 As UserInfo = UserInfo.getInstance

    Dim drv As DataRowView



    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Teams.Add(New Team With {.TeamId = "1", .TeamName = "QE"})
        Teams.Add(New Team With {.TeamId = "2", .TeamName = "Process Improvement"})
    End Sub

    Private Sub RefreshMYInterface()
        DataGridView1.Invalidate()
    End Sub

    Sub DoWork()
        myController = New ActivityMasterController
        Try
            ProgressReport(1, "Loading...Please wait.")
            If myController.loaddata() Then
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
                    TEAMBS = New BindingSource
                    TEAMBS.DataSource = Teams
                    DirectCast(DataGridView1.Columns("cbTeam"), DataGridViewComboBoxColumn).DataSource = TEAMBS
                    DirectCast(DataGridView1.Columns("cbTeam"), DataGridViewComboBoxColumn).DisplayMember = "TeamName"
                    DirectCast(DataGridView1.Columns("cbTeam"), DataGridViewComboBoxColumn).ValueMember = "TeamId"
                    DirectCast(DataGridView1.Columns("cbTeam"), DataGridViewComboBoxColumn).DataPropertyName = "activitygroup"

                    DataGridView1.DataSource = myController.BS


            End Select
        End If
    End Sub


    Private Sub AddNewToolStripButton1_Click(sender As Object, e As EventArgs) Handles AddToolStripButton.Click
        drv = myController.GetNewRecord
        Me.drv.BeginEdit()
    End Sub

    Private Sub DeleteToolStripButton2_Click(sender As Object, e As EventArgs) Handles DeleteToolStripButton.Click
        If Not IsNothing(myController.GetCurrentRecord) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    myController.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub FormITAssets_Load(sender As Object, e As EventArgs) Handles Me.Load
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



    Private Sub ToolStripTextBox1_Click(sender As Object, e As EventArgs) Handles ToolStripTextBox1.Click

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class

Public Class Team
    Public Property TeamId
    Public Property TeamName

    Public Overrides Function tostring() As String
        Return TeamName
    End Function
End Class