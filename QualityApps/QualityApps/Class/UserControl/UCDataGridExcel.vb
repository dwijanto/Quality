Public Class UCDataGridExcel
    Private WithEvents BS As BindingSource
    Private myform As Object
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If Not IsNothing(BS) Then
            For Each drv As DataRowView In BS.List
                drv.Row.Item("selected") = CheckBox1.Checked
            Next
            BS.EndEdit()
        End If
    End Sub

    Public Sub BindingControl(ByVal myform As Object, ByRef bs As BindingSource)
        Me.myform = myform
        Me.BS = BS
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = BS
    End Sub
End Class
