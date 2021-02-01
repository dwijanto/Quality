Imports System.Threading
Public Class FormMissingRemarksCP
    Private Shared myform As FormMissingRemarksCP
    Private MissingData As MissingDataEnum
    Public Shared Function getInstance()
        If myform Is Nothing Then
            myform = New FormMissingRemarksCP
        ElseIf myform.IsDisposed Then
            myform = New FormMissingRemarksCP
        End If
        Return myform
    End Function

    Dim myController As InspectionControllerCP
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)

    'Dim UserInfo1 As UserInfo = UserInfo.getInstance

    Dim drv As DataRowView



    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Sub New(ByVal value As MissingDataEnum)

        ' This call is required by the designer.
        InitializeComponent()
        MissingData = value
        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Private Sub RefreshMYInterface()
        DataGridView1.Invalidate()
    End Sub

    Sub DoWork()
        myController = New InspectionControllerCP
        Try
            ProgressReport(1, "Loading...Please wait.")
            Select Case MissingData
                Case MissingDataEnum.MissingRemark
                    If myController.loaddataAllMising() Then
                        ProgressReport(4, "Init Data")
                    End If
                Case MissingDataEnum.MissingInspDate
                    If myController.loaddataAllMisingInspDate() Then
                        ProgressReport(4, "Init Data")
                    End If
            End Select

            ProgressReport(1, String.Format("Loading...Done. Records {0}", myController.BSMissing.Count))
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
                    DataGridView1.DataSource = myController.BSMissing

            End Select
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


    Private Sub ToolStripTextBox1_TextChanged(sender As Object, e As EventArgs) Handles ToolStripTextBox1.TextChanged
        myController.ApplyFilterMissing = ToolStripTextBox1.Text
    End Sub


    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        If Not IsNothing(myController.GetCurrentRecord) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    myController.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Me.Validate()
        myController.save()
        ProgressReport(1, String.Format("Done. Records {0}", myController.BSMissing.Count))
    End Sub

    Private Sub ToolStripTextBox1_Click(sender As Object, e As EventArgs) Handles ToolStripTextBox1.Click

    End Sub
End Class