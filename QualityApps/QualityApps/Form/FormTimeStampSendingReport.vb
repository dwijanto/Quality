Public Class FormTimeStampSendingReport

    Public Shared myForm As FormTimeStampSendingReport
    Private SaveFileName As String = String.Empty
    Private SaveFileDialog1 As New SaveFileDialog
    Dim myController As New ActivityLogController
    Private myIdentity As UserController = User.getIdentity

    Public Shared Function getInstance()
        If myForm Is Nothing Then
            myForm = New FormTimeStampSendingReport
        ElseIf myForm.IsDisposed Then
            myForm = New FormTimeStampSendingReport
        End If
        Return myForm
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        SaveFileDialog1.FileName = String.Format("TimeStampSendingReport-{0:yyyyMMdd}.xlsx", Date.Today)
        Dim criteria = String.Format(" where latestupdate >= '{0:yyyy-MM-dd}' and latestupdate <= '{1:yyyy-MM-dd}'", DateTimePicker1.Value.Date, DateTimePicker2.Value.Date)
        Dim sqlstr As String = String.Format("select * from quality.sendemailtx {0}", criteria)
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim myReport As ExportToExcelFile = New ExportToExcelFile(Me, sqlstr, IO.Path.GetDirectoryName(SaveFileDialog1.FileName), IO.Path.GetFileName(SaveFileDialog1.FileName), AddressOf FormatReport, AddressOf PivotCallback, 1, "\templates\exceltemplate.xltx")
            myReport.Run(Me, New EventArgs)
        End If
    End Sub

    Private Sub FormatReport()
        'Throw New NotImplementedException
    End Sub

    Private Sub PivotCallback(ByRef sender As Object, ByRef e As EventArgs)
        
    End Sub
End Class