Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class AutoInspectionReport
    <DllImport("user32.dll")> _
    Public Shared Function EndTask(ByVal hWnd As IntPtr, ByVal fShutDown As Boolean, ByVal fForce As Boolean) As Boolean
    End Function
    Public dbAdapter1 As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance
    Public Property errmsg As String
    Dim myParam As ParamAdapter = ParamAdapter.getInstance
    Dim mypath As String = myParam.GetMailFolder("AUTO REPORT")

    Dim Filename As String = String.Empty
    Dim mycontroller As InspectionController = New InspectionController

    Public Sub New()
        Me.mypath = mypath
        Me.filename = String.Format("InspectionReport-{0:yyyyMMdd}.xlsx", Date.Today)
    End Sub

    Public Function GenerateReport() As Boolean
        Dim myret As Boolean = False


        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        'Dim enddate As Date = Today.Date.AddDays(-1)
        Dim enddate As Date = Today.Date
        Dim hwnd As System.IntPtr
        Dim result As Boolean
        Try
            'Create Object Excel 

            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd

            oXl.Visible = False
            oXl.DisplayAlerts = False

            oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\ExcelTemplate.xltx")

            Dim counter As Integer = 0
            'ProgressReport(2, "Creating Worksheet...")
            'backOrder
            'For i = 0 To 2
            '    oWb.Worksheets.Add(After:=(oWb.Worksheets(3 + i)))
            'Next i

            Dim sqlstr As String = String.Empty
            '
            'Get Filter

            oSheet = oWb.Worksheets(1)
            Dim myfilter As New System.Text.StringBuilder

            sqlstr = mycontroller.Model.sqlstr

            oSheet.Name = "DATA"

            FillWorksheet(oSheet, sqlstr, dbAdapter1)
            Dim lastrow = oSheet.Cells.Find(What:="*", SearchDirection:=Excel.XlSearchDirection.xlPrevious, SearchOrder:=Excel.XlSearchOrder.xlByRows).Row

            If lastrow > 1 Then
                '
                ApplyFormat(oSheet)
                'oXl.Visible = True
                CreatePivotTable(oXl, oWb, 1, enddate)
                'createchart(oWb, 1, errmsg)
            End If

            'remove connection
            For i = 0 To oWb.Connections.Count - 1
                'oWb.Connections(1).Delete()
            Next            
            Filename = String.Format("{0}\{1}", mypath, Filename)


            oWb.SaveAs(filename)
            result = True
        Catch ex As Exception
            errmsg = ex.Message
        Finally
            'clear excel from memory
            oXl.Quit()
            releaseComObject(oSheet)
            releaseComObject(oWb)
            releaseComObject(oXl)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            Try
                'to make sure excel is no longer in memory
                EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try
            Cursor.Current = Cursors.Default
        End Try
        Return result

        Return myret
    End Function

    Public Shared Sub FillWorksheet(ByVal osheet As Excel.Worksheet, ByVal sqlstr As String, ByVal dbAdapter As Object, Optional ByVal Location As String = "A1")
        'Dim oRange As Excel.Range
        'Dim oExCon As String = My.Settings.oExCon '"ODBC;DSN=PostgreSQLhon03nt;"
        'oExCon = oExCon.Insert(oExCon.Length, "UID=admin;Pwd=admin")
        'Dim oRange As Excel.Range
        'oRange = osheet.Range(Location)
        Dim oExCon As String = My.Settings.oExCon ' My.Settings.oExCon.ToString '"ODBC;DSN=PostgreSQL30;"
        'oExCon = oExCon.Insert(oExCon.Length, "UID=" & dbadapter1.userid & ";PWD=" & dbadapter1.password)
        Dim dbAdapter1 = PostgreSQLDBAdapter.getInstance
        oExCon = oExCon.Insert(oExCon.Length, "UID=" & dbAdapter1.UserId & ";PWD=" & dbAdapter1.Password)
        Dim oRange As Excel.Range
        oRange = osheet.Range(Location)
        With osheet.QueryTables.Add(oExCon.Replace("Host=", "Server="), oRange)
            'With osheet.QueryTables.Add(oExCon.Replace("Host=", "Server="), oRange)
            'With osheet.QueryTables.Add(oExCon, osheet.Range("A1"))
            .CommandText = sqlstr
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = True
            .RefreshStyle = Excel.XlCellInsertionMode.xlInsertDeleteCells
            .SavePassword = True
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .PreserveColumnInfo = True
            .Refresh(BackgroundQuery:=False)
            Application.DoEvents()
        End With

        oRange = Nothing
    End Sub

    Private Sub CreatePivotTable(ByVal oxl As Excel.Application, ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal mydate As Date)
        

    End Sub

    Public Shared Sub releaseComObject(ByRef o As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
        Catch ex As Exception
        Finally
            o = Nothing
        End Try
    End Sub

    Private Sub ApplyFormat(oSheet As Excel.Worksheet)
        'Dim osheet As Excel.Worksheet = DirectCast(sender, Excel.Worksheet)
        osheet.Columns("A:A").Delete()
    End Sub

End Class
