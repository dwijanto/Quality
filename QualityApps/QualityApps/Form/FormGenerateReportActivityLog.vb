Public Class FormGenerateReportActivityLog

    Public Shared myForm As FormGenerateReportActivityLog
    Private SaveFileName As String = String.Empty
    Private SaveFileDialog1 As New SaveFileDialog
    Dim myController As New ActivityLogController
    Private myIdentity As UserController = User.getIdentity

    Public Shared Function getInstance()
        If myForm Is Nothing Then
            myForm = New FormGenerateReportActivityLog
        ElseIf myForm.IsDisposed Then
            myForm = New FormGenerateReportActivityLog
        End If
        Return myForm
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        myController = New ActivityLogController
        SaveFileDialog1.FileName = String.Format("ActivityLog-{0:yyyyMMdd}.xlsx", Date.Today)
        Dim criteria = String.Format(" where u.userid in (select quality.getsubordinate('{0}')) and activitydate >= '{1:yyyy-MM-dd}' and activitydate <= '{2:yyyy-MM-dd}'", myIdentity.userid.ToLower, DateTimePicker1.Value.Date, DateTimePicker2.Value.Date)
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim myReport As ExportToExcelFile = New ExportToExcelFile(Me, myController.GetSQLSTRReport(criteria), IO.Path.GetDirectoryName(SaveFileDialog1.FileName), IO.Path.GetFileName(SaveFileDialog1.FileName), AddressOf FormatReport, AddressOf PivotCallback, 1, "\templates\excelTemplate.xltx")
            myReport.Run(Me, New EventArgs)
        End If
    End Sub

    Private Sub FormatReport()
        'Throw New NotImplementedException
    End Sub

    Private Sub PivotCallback()
        'Throw New NotImplementedException
    End Sub

End Class