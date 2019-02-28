Imports System.Threading
Public Class FormRemarkHistory
    Dim model As New InspectionModel
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    'Dim myController As New HistoryController
    Private Shared myform As FormRemarkHistory
    Public Property pono As Long
    Public Property item As Integer
    Public Property seqn As Integer
    Public Property qty As Decimal

    Public Shared Function getInstance()
        If myform Is Nothing Then
            myform = New FormRemarkHistory
        ElseIf myform.IsDisposed Then
            myform = New FormRemarkHistory
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
        model.GetRemarkHistory(pono, item, seqn, qty)
    End Sub

    Private Sub DialogHistory_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Not myThread.IsAlive Then
            'ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub
End Class