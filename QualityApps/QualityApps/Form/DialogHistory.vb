Imports System.Windows.Forms
Imports System.Threading

Public Class DialogHistory
    'Dim model As New InspectionModel
    Dim myController As New InspectionController
    Dim myAdapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance

    Dim myThread As New System.Threading.Thread(AddressOf DoWork) 
    Private Shared myform As DialogHistory



    Public Property pono As Long
    Public Property item As Integer
    Public Property seqn As Integer
    Public Property qty As Decimal
    'Public Property Id As Long
    Dim DS As DataSet
    Dim bs As New BindingSource
    Dim bs2 As New BindingSource

    Public Event RefreshInspectionDate(ByVal po As Long, ByVal item As Integer, ByVal SeqN As Integer, ByVal NewDate As Date)
    Public Event RefreshRemark(ByVal po As Long, ByVal item As Integer, ByVal SeqN As Integer, myRemark As String)

    Public Shared Function getInstance()
        If myform Is Nothing Then
            myform = New DialogHistory
        ElseIf myform.IsDisposed Then
            myform = New DialogHistory
        End If
        Return myform
    End Function

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Sub DoWork()        
        DS = myController.Model.GetRemarkHistory(pono, item, seqn, qty)
        If Not IsNothing(DS) Then
            ProgressReport(4, "Init Data")
        End If
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
                        'ToolStripStatusLabel1.Text = message
                    Case 2
                        'ToolStripStatusLabel1.Text = message
                    Case 4
                        'BindingData
                        DataGridView1.AutoGenerateColumns = False
                        DataGridView2.AutoGenerateColumns = False

                        bs.DataSource = DS.Tables(0)
                        DataGridView1.DataSource = bs


                        bs2.DataSource = DS.Tables(1)
                        DataGridView2.DataSource = bs2

                        Me.Focus()
                    Case 5
                        'ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        'ToolStripProgressBar1.Style = ProgressBarStyle.Marquee


                End Select
            Catch ex As Exception
                ProgressReport(1, ex.Message)
            End Try
        End If

    End Sub
    Public Sub Run()
        If Not myThread.IsAlive Then
            'ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub
    Private Sub DialogHistory_Load(sender As Object, e As EventArgs) Handles Me.Load        
    End Sub

    Private Sub AssignSeqNToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AssignSeqNToolStripMenuItem.Click
        If MessageBox.Show(String.Format("Do you want to assign with seqN {0}?", seqn), "Missing SeqN", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = Windows.Forms.DialogResult.OK Then
            Dim drv As DataRowView = bs2.Current
            Dim myret As Boolean
            myret = myController.Model.UpdateSeqNHistory(drv.Row.Item("id"), seqn)
            If myret Then
                MessageBox.Show("Updated.")
                'get inspdate from the last update one
                Dim myDate = myController.Model.GetLatestHistoryRecord()
                If myDate <> "#12:00:00 AM#" Then
                    RaiseEvent RefreshInspectionDate(pono, item, seqn, myDate)
                End If
                Dim myRemark = myController.Model.GetLatestRemarksHistoryRecord()
                If Not IsNothing(myRemark) Then
                    RaiseEvent RefreshRemark(pono, item, seqn, myRemark)
                End If
                'If Not IsDBNull(drv.Row.Item("inspdate")) Then
                '    RaiseEvent RefreshInspectionDate(pono, item, seqn, drv.Row.Item("inspdate"))
                'End If
                'If Not IsDBNull(drv.Row.Item("remark")) Then
                '    RaiseEvent RefreshRemark(pono, item, seqn, drv.Row.Item("remark"))
                'End If
            Else
                MessageBox.Show(myController.Model.ErrorMessage)
            End If
        End If
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RefreshToolStripMenuItem.Click
        Me.Run()
    End Sub

    Private Sub AddToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddToolStripMenuItem.Click
        'Dim currDrv As DataRowView = bs.Current
        Dim drv As DataRowView = bs.AddNew
        drv.Row.Item("purchdoc") = pono
        drv.Row.Item("item") = item
        drv.Row.Item("seqn") = seqn
        drv.Row.Item("qty") = qty
        drv.Row.Item("docdate") = Today.Date
    End Sub

    Private Sub CommitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CommitToolStripMenuItem.Click
        If myController.savehistory() Then
            'If MessageBox.Show("Refresh Current Record in browser?", "Refresh Inspection Date", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
            Dim myDate = myController.Model.GetLatestHistoryRecord()
            If Not IsNothing(myDate) Then
                RaiseEvent RefreshInspectionDate(pono, item, seqn, myDate)
            End If
            Dim myRemark = myController.Model.GetLatestRemarksHistoryRecord()
            If Not IsNothing(myRemark) Then
                RaiseEvent RefreshRemark(pono, item, seqn, myRemark)
            End If

            'End If
        End If

    End Sub

    Private Sub UpdateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UpdateToolStripMenuItem.Click

    End Sub

    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        If Not IsNothing(myController.GetCurrentHistoryRecord) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    myController.RemoveHistoryAt(drv.Index)
                Next
            End If
        End If
    End Sub
End Class
