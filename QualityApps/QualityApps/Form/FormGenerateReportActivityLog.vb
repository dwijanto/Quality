Imports Microsoft.Office.Interop

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
            Dim myReport As ExportToExcelFile = New ExportToExcelFile(Me, myController.GetSQLSTRReport(criteria), IO.Path.GetDirectoryName(SaveFileDialog1.FileName), IO.Path.GetFileName(SaveFileDialog1.FileName), AddressOf FormatReport, AddressOf PivotCallback, 2, "\templates\activitylog.xltx")
            myReport.Run(Me, New EventArgs)
        End If
    End Sub

    Private Sub FormatReport()
        'Throw New NotImplementedException
    End Sub

    Private Sub PivotCallback(ByRef sender As Object, ByRef e As EventArgs)
        Dim oXl As Excel.Application = Nothing
        Dim owb As Excel.Workbook = CType(sender, Excel.Workbook)
        oXl = owb.Parent
        owb.Worksheets(2).select()
        Dim osheet = owb.Worksheets(2)
        Dim orange = osheet.Range("A2")
        If osheet.cells(2, 2).text.ToString = "" Then
            Err.Raise(100, Description:="Data not available.")
        End If
        osheet.Columns("B:B").NumberFormat = "dd-MMM-yyyy"
        osheet.name = "RawData"
        owb.Names.Add("db", RefersToR1C1:="=OFFSET('RawData'!R1C1,0,0,COUNTA('RawData'!C1),COUNTA('RawData'!R1))")

        owb.Worksheets(1).select()
        osheet = owb.Worksheets(1)
        osheet.PivotTables("PivotTable1").ChangePivotCache(owb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, SourceData:="db"))
        'oXl.Run("ShowFG")
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        osheet.pivottables("PivotTable1").SaveData = True
        owb.RefreshAll()
    End Sub

End Class