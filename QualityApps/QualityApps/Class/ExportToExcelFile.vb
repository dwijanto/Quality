Imports Microsoft.Office.Interop
Imports System.Threading
Imports System.Text

Imports System.IO
Imports System.Runtime.InteropServices

Public Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
Public Delegate Sub FormatReportDelegate(ByRef sender As Object, ByRef e As EventArgs)

Public Class ExportToExcelFile
    <DllImport("user32.dll")> _
    Public Shared Function EndTask(ByVal hWnd As IntPtr, ByVal fShutDown As Boolean, ByVal fForce As Boolean) As Boolean
    End Function
    Public Property sqlstr As String
    Public Property Directory As String
    Public Property ReportName As String
    Public Property Parent As Object
    Public Property FormatReportCallback As FormatReportDelegate
    Dim myThread As New Threading.Thread(AddressOf DoWork)
    Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable
    Dim ValidateWorksheetCallback As FormatReportDelegate = AddressOf PivotTable
    Dim AccessFullPath As String
    Dim AccessTableName As String
    Dim SpecificationName As String
    Dim status As Boolean
    Dim Dataset1 As New DataSet
    Public Property Datasheet As Integer = 1
    Public Property mytemplate As String = "\templates\ExcelTemplate.xltx"
    Public Property Location As String = "A1"
    Public Property QueryList As List(Of QueryWorksheet)
    Public Property AskOpenFile As Boolean = True
    Private myKey As Object

    Public Property DataTableList As List(Of DataTableWorksheet)
    Private DT As DataTable
    Public Property errorMessage As String

    Public Sub New(ByRef parent As Object, ByVal Directory As String, ValidateWorksheetCallback As FormatReportDelegate)
        Me.Parent = parent
        Me.Directory = Directory
        Me.ValidateWorksheetCallback = ValidateWorksheetCallback
    End Sub

    Public Sub New(ByRef Parent As Object, ByRef Sqlstr As String, ByRef Directory As String, ByRef ReportName As String, ByVal FormatReportCallBack As FormatReportDelegate)
        Me.sqlstr = Sqlstr
        Me.Directory = Directory
        Me.ReportName = ReportName
        Me.Parent = Parent
        Me.FormatReportCallback = FormatReportCallBack
    End Sub
    Public Sub New(ByRef Parent As Object, ByRef Sqlstr As String, ByRef Directory As String, ByRef ReportName As String, ByVal FormatReportCallBack As FormatReportDelegate, ByVal PivotCallback As FormatReportDelegate)
        Me.sqlstr = Sqlstr
        Me.Directory = Directory
        Me.ReportName = ReportName
        Me.Parent = Parent
        Me.FormatReportCallback = FormatReportCallBack
        Me.PivotCallback = PivotCallback
    End Sub

    Public Sub New(ByRef Parent As Object, ByRef Sqlstr As String, ByRef Directory As String, ByRef ReportName As String, ByVal FormatReportCallBack As FormatReportDelegate, ByVal PivotCallback As FormatReportDelegate, ByVal AccessFullpath As String, ByVal AccessTableName As String, ByVal SpecificationName As String)
        Me.sqlstr = Sqlstr
        Me.Directory = Directory
        Me.ReportName = ReportName
        Me.Parent = Parent
        Me.FormatReportCallback = FormatReportCallBack
        Me.PivotCallback = PivotCallback
        Me.AccessFullPath = AccessFullpath
        Me.AccessTableName = AccessTableName
        Me.SpecificationName = SpecificationName
    End Sub

    Public Sub New(ByRef Parent As Object, ByRef Sqlstr As String, ByRef Directory As String, ByRef ReportName As String, ByVal FormatReportCallBack As FormatReportDelegate, ByVal PivotCallback As FormatReportDelegate, ByVal datasheet As Integer, ByVal mytemplate As String)
        Me.sqlstr = Sqlstr
        Me.Directory = Directory
        Me.ReportName = ReportName
        Me.Parent = Parent
        Me.FormatReportCallback = FormatReportCallBack
        Me.PivotCallback = PivotCallback
        Me.Datasheet = datasheet
        Me.mytemplate = mytemplate
    End Sub
    Public Sub New(ByRef Parent As Object, ByRef Sqlstr As String, ByRef Directory As String, ByRef ReportName As String, ByVal FormatReportCallBack As FormatReportDelegate, ByVal PivotCallback As FormatReportDelegate, ByVal datasheet As Integer, ByVal mytemplate As String, ByVal Location As String)
        Me.sqlstr = Sqlstr
        Me.Directory = Directory
        Me.ReportName = ReportName
        Me.Parent = Parent
        Me.FormatReportCallback = FormatReportCallBack
        Me.PivotCallback = PivotCallback
        Me.Datasheet = datasheet
        Me.mytemplate = mytemplate
        Me.Location = Location
    End Sub

    Public Sub New(ByRef Parent As Object, ByRef querylist As List(Of QueryWorksheet), ByRef Directory As String, ByRef ReportName As String, ByVal FormatReportCallBack As FormatReportDelegate, ByVal PivotCallback As FormatReportDelegate)
        Me.QueryList = querylist
        Me.Directory = Directory
        Me.ReportName = ReportName
        Me.Parent = Parent
        Me.FormatReportCallback = FormatReportCallBack
        Me.PivotCallback = PivotCallback
    End Sub

    Public Sub New(ByRef Parent As Object, ByVal ReportName As String, ByRef myTemplate As String, ByVal dt As DataTable, ByVal FormatReportCallBack As FormatReportDelegate, ByVal PivotCallback As FormatReportDelegate, Optional openfile As Boolean = True)
        Me.Parent = Parent
        Me.ReportName = ReportName
        Me.mytemplate = myTemplate
        Me.DT = dt
        Me.FormatReportCallback = FormatReportCallBack
        Me.AskOpenFile = openfile
    End Sub


    Public Sub Run(ByRef sender As System.Object, ByVal e As System.EventArgs)

        ' FileName = Application.StartupPath & "\PrintOut"      
        errorMessage = String.Empty
        If Not myThread.IsAlive Then
            Try
                myThread = New System.Threading.Thread(New ThreadStart(AddressOf DoWork))
                myThread.SetApartmentState(ApartmentState.MTA)
                myThread.Start()
            Catch ex As Exception
                'MsgBox(ex.Message)
                errorMessage = ex.Message
            End Try
        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub

    Public Sub RunWithKey(ByRef sender As System.Object, ByVal e As ExportToExcelWithKeyEventArgs)
        If Not myThread.IsAlive Then
            Try
                Me.myKey = e.myKey
                myThread = New System.Threading.Thread(New ThreadStart(AddressOf DoWorkKey))
                'myThread.SetApartmentState(ApartmentState.MTA)
                myThread.SetApartmentState(ApartmentState.STA)
                myThread.Start()
            Catch ex As Exception
                Logger.log(ex.Message)
                MsgBox(ex.Message)
            End Try
        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub

    Private Sub DoWorkKey()
        Dim errMsg As String = String.Empty
        Dim i As Integer = 0
        Dim errSB As New StringBuilder
        Dim sw As New Stopwatch
        sw.Start()
        ProgressReport(2, "Export To Excel..")
        ProgressReport(6, "Marques..")
        status = GenerateReportWithKey(Directory, errMsg, Dataset1)
        ProgressReport(5, "Continues..")
        If status Then


            sw.Stop()
            ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            ProgressReport(3, "")
            If AskOpenFile Then
                If MsgBox("File name: " & Directory & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                    Process.Start(Directory)
                End If
            End If

            ProgressReport(3, "")
            'ProgressReport(4, errSB.ToString)
        Else
            errSB.Append(errMsg) '& vbCrLf)
            ProgressReport(3, errSB.ToString)
        End If
        sw.Stop()
    End Sub

    Public Sub CreateForm(ByRef sender As System.Object, ByVal e As System.EventArgs)

        ' FileName = Application.StartupPath & "\PrintOut"
        If Not myThread.IsAlive Then
            Try
                myThread = New System.Threading.Thread(New ThreadStart(AddressOf DoWorkForm))
                myThread.SetApartmentState(ApartmentState.MTA)
                myThread.Start()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub


    Public Sub DoWork()
        Dim errMsg As String = String.Empty
        Dim i As Integer = 0
        Dim errSB As New StringBuilder
        Dim sw As New Stopwatch
        sw.Start()
        ProgressReport(2, "Export To Excel..")
        ProgressReport(6, "Marques..")
        status = GenerateReport(Directory, errMsg, Dataset1)
        ProgressReport(5, "Continues..")
        If status Then

            sw.Stop()
            ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            ProgressReport(3, "")
            If AskOpenFile Then
                If MsgBox("File name: " & Directory & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                    Process.Start(Directory)
                End If
            End If

            ProgressReport(3, "")
            'ProgressReport(4, errSB.ToString)
        Else
            errSB.Append(errMsg) '& vbCrLf)
            errorMessage = errSB.ToString
            ProgressReport(3, errSB.ToString)
        End If
        sw.Stop()


    End Sub

    Sub DoWorkForm()
        Dim errMsg As String = String.Empty
        Dim i As Integer = 0
        Dim errSB As New StringBuilder
        Dim sw As New Stopwatch
        sw.Start()
        ProgressReport(2, "Export To Excel..")
        ProgressReport(6, "Marques..")
        status = GenerateReport(Directory, errMsg)
        ProgressReport(5, "Continues..")
        If status Then


            sw.Stop()
            ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            ProgressReport(3, "")

            If MsgBox("File name: " & Directory & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                Process.Start(Directory)
            End If
            ProgressReport(3, "")
            'ProgressReport(4, errSB.ToString)
        Else
            errSB.Append(errMsg) '& vbCrLf)
            ProgressReport(3, errSB.ToString)
        End If
        sw.Stop()


    End Sub
    Private Function GenerateReportWithKey(ByRef FileName As String, ByRef errorMsg As String, ByVal dataset1 As DataSet) As Boolean
        Dim myCriteria As String = String.Empty
        Dim result As Boolean = False

        Dim StopWatch As New Stopwatch
        StopWatch.Start()
        'Open Excel
        Application.DoEvents()

        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr

        Try
            'Create Object Excel 
            Parent.ProgressReport(2002, String.Format("{0},CreateObject..", myKey))
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            Parent.ProgressReport(2001, String.Format("{0},{1}", myKey, hwnd))
            'oXl.ScreenUpdating = False
            'oXl.Visible = False
            oXl.DisplayAlerts = False
            Parent.ProgressReport(2002, String.Format("{0},Opening Template...", myKey))
            Parent.ProgressReport(2002, String.Format("{0},Generating records..", myKey))
            If mytemplate.Contains("172") Then
                oWb = oXl.Workbooks.Open(mytemplate)
            Else
                Thread.Sleep(100) 'slowdown the process
                oWb = oXl.Workbooks.Open(Application.StartupPath & mytemplate)
            End If

            oXl.Visible = False
            'For i = 0 To 6
            '    oWb.Worksheets.Add()
            'Next

            'Dim events As New List(Of ManualResetEvent)()
            'Dim counter As Integer = 0
            Parent.ProgressReport(2002, String.Format("{0},Creating Worksheet...", myKey))
            'DATA


            If IsNothing(QueryList) Then
                oWb.Worksheets(Datasheet).select()
                oSheet = oWb.Worksheets(Datasheet)

                If sqlstr <> "" Then
                    Parent.ProgressReport(2002, String.Format("{0},Get records..", myKey))
                    FillWorksheet(oSheet, sqlstr, Location)
                    'Dim orange = oSheet.Range("A1")
                    Dim orange = oSheet.Range(Location)
                    Dim lastrow = GetLastRow(oXl, oSheet, orange)
                    If lastrow > 1 Then
                        'Delegate for modification
                        'oSheet.Columns("A:A").numberformat = "dd-MMM-yyyy"
                        FormatReportCallback.Invoke(oSheet, New EventArgs)
                    End If
                Else
                    FormatReportCallback.Invoke(oSheet, New EventArgs)
                End If
            Else
                'Looping from here
                For i = 0 To QueryList.Count - 1
                    Dim myquery = CType(QueryList(i), QueryWorksheet)
                    oWb.Worksheets(myquery.DataSheet).select()
                    oSheet = oWb.Worksheets(myquery.DataSheet)
                    oSheet.Name = myquery.SheetName
                    Parent.ProgressReport(2002, String.Format("{0},Get records..", myKey))

                    FillWorksheet(oSheet, myquery.Sqlstr, Location)
                    'Dim orange = oSheet.Range("A1")
                    Dim orange = oSheet.Range(Location)
                    Dim lastrow = GetLastRow(oXl, oSheet, orange)


                    If lastrow > 1 Then
                        'Delegate for modification
                        'oSheet.Columns("A:A").numberformat = "dd-MMM-yyyy"
                        'FormatReportCallback.Invoke(oSheet, New EventArgs)
                        FormatReportCallback.Invoke(oSheet, New ExportToExcelFileEventArgs With {.lastRow = lastrow, .SheetNo = myquery.DataSheet})
                    End If
                Next
                oWb.Worksheets(1).select()
                oSheet = oWb.Worksheets(1)

                'End Looping
            End If



            For i = 0 To oWb.Connections.Count - 1
                oWb.Connections(1).Delete()
            Next
            PivotCallback.Invoke(oWb, New EventArgs)
            StopWatch.Stop()

            'FileName = FileName & "\" & String.Format("Report" & ReportName & "-{0}-{1}-{2}.xlsx", Today.Year, Format("00", Today.Month), Format("00", Today.Day))
            FileName = FileName & "\" & String.Format(ReportName)
            'ProgressReport(3, "")
            Parent.ProgressReport(2002, String.Format("{0},Saving File ...{1}", myKey, FileName))
            'oSheet.Name = ReportName


            If FileName.Contains("xlsm") Then
                oWb.SaveAs(FileName, FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)
            Else
                oWb.SaveAs(FileName)
            End If

            Parent.ProgressReport(2002, String.Format("{0},{1}", myKey, "Elapsed Time: " & Format(StopWatch.Elapsed.Minutes, "00") & ":" & Format(StopWatch.Elapsed.Seconds, "00") & "." & StopWatch.Elapsed.Milliseconds.ToString))
            Parent.ProgressReport(2003, String.Format("{0},{1}", myKey, True))
            result = True
        Catch ex As Exception
            ProgressReport(3, ex.Message & FileName)
            Logger.log(ex.Message & FileName)
            errorMsg = ex.Message
        Finally
            'oXl.ScreenUpdating = True
            'clear excel from memory
            Try
                oXl.Quit()
                releaseComObject(oSheet)
                releaseComObject(oWb)
                releaseComObject(oXl)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception
                Logger.log(ex.Message & FileName)
            End Try

            Try
                'to make sure excel is no longer in memory
                EndTask(hwnd, True, True)
                Parent.ProgressReport(2001, String.Format("{0},", myKey))
                Parent.ProgressReport(2005, String.Format("{0}", 1))
            Catch ex As Exception
                Logger.log(ex.Message & FileName)
            End Try

        End Try
        Return result
    End Function


    Private Function GenerateReport(ByRef FileName As String, ByRef errorMsg As String, ByVal dataset1 As DataSet) As Boolean
        Dim myCriteria As String = String.Empty
        Dim result As Boolean = False

        Dim StopWatch As New Stopwatch
        StopWatch.Start()
        'Open Excel
        Application.DoEvents()

        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr

        Try
            'Create Object Excel 
            ProgressReport(2, "CreateObject..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            'oXl.ScreenUpdating = False
            'oXl.Visible = False
            oXl.DisplayAlerts = False
            ProgressReport(2, "Opening Template...")
            ProgressReport(2, "Generating records..")
            If mytemplate.Contains("172") Then
                oWb = oXl.Workbooks.Open(mytemplate)
            Else
                oWb = oXl.Workbooks.Open(Application.StartupPath & mytemplate)
            End If

            oXl.Visible = False
            'For i = 0 To 6
            '    oWb.Worksheets.Add()
            'Next

            'Dim events As New List(Of ManualResetEvent)()
            'Dim counter As Integer = 0
            ProgressReport(2, "Creating Worksheet...")
            'DATA


            If IsNothing(QueryList) Then
                oWb.Worksheets(Datasheet).select()
                oSheet = oWb.Worksheets(Datasheet)

                If sqlstr <> "" Then
                    ProgressReport(2, "Get records..")
                    FillWorksheet(oSheet, sqlstr, Location)
                    'Dim orange = oSheet.Range("A1")
                    Dim orange = oSheet.Range(Location)
                    Dim lastrow = GetLastRow(oXl, oSheet, orange)
                    If lastrow > 1 Then
                        'Delegate for modification
                        'oSheet.Columns("A:A").numberformat = "dd-MMM-yyyy"
                        FormatReportCallback.Invoke(oSheet, New EventArgs)
                    End If
                Else
                    FormatReportCallback.Invoke(oSheet, New EventArgs)
                End If
            Else
                'Looping from here
                For i = 0 To QueryList.Count - 1
                    Dim myquery = CType(QueryList(i), QueryWorksheet)
                    oWb.Worksheets(myquery.DataSheet).select()
                    oSheet = oWb.Worksheets(myquery.DataSheet)
                    oSheet.Name = myquery.SheetName
                    ProgressReport(2, "Get records..")

                    FillWorksheet(oSheet, myquery.Sqlstr, Location)
                    'Dim orange = oSheet.Range("A1")
                    Dim orange = oSheet.Range(Location)
                    Dim lastrow = GetLastRow(oXl, oSheet, orange)


                    If lastrow > 1 Then
                        'Delegate for modification
                        'oSheet.Columns("A:A").numberformat = "dd-MMM-yyyy"
                        'FormatReportCallback.Invoke(oSheet, New EventArgs)
                        FormatReportCallback.Invoke(oSheet, New ExportToExcelFileEventArgs With {.lastRow = lastrow, .SheetNo = myquery.DataSheet})
                    End If
                Next
                oWb.Worksheets(1).select()
                oSheet = oWb.Worksheets(1)

                'End Looping
            End If


            PivotCallback.Invoke(oWb, New EventArgs)
            For i = 0 To oWb.Connections.Count - 1
                oWb.Connections(1).Delete()
            Next
            StopWatch.Stop()

            'FileName = FileName & "\" & String.Format("Report" & ReportName & "-{0}-{1}-{2}.xlsx", Today.Year, Format("00", Today.Month), Format("00", Today.Day))
            FileName = FileName & "\" & String.Format(ReportName)
            ProgressReport(3, "")
            ProgressReport(2, "Saving File ..." & FileName)
            'oSheet.Name = ReportName
            If FileName.Contains("xlsm") Then

                oWb.SaveAs(FileName, FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)
            Else
                oWb.SaveAs(FileName)
            End If

            ProgressReport(2, "Elapsed Time: " & Format(StopWatch.Elapsed.Minutes, "00") & ":" & Format(StopWatch.Elapsed.Seconds, "00") & "." & StopWatch.Elapsed.Milliseconds.ToString)
            result = True
        Catch ex As Exception
            ProgressReport(3, ex.Message & FileName)
            errorMsg = ex.Message
        Finally
            'oXl.ScreenUpdating = True
            'clear excel from memory
            Try
                oXl.Quit()
                releaseComObject(oSheet)
                releaseComObject(oWb)
                releaseComObject(oXl)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception

            End Try

            Try
                'to make sure excel is no longer in memory
                EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try

        End Try
        Return result
    End Function

    Public Function GenerateReport(ByRef FileName As String, ByRef errorMsg As String) As Boolean
        Dim myCriteria As String = String.Empty
        Dim result As Boolean = False

        Dim StopWatch As New Stopwatch
        StopWatch.Start()
        'Open Excel
        Application.DoEvents()

        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr
        Try
            'Create Object Excel 
            ProgressReport(2, "CreateObject..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            'oXl.ScreenUpdating = False
            'oXl.Visible = False
            oXl.DisplayAlerts = False
            ProgressReport(2, "Opening Template...")
            ProgressReport(2, "Generating records..")
            If mytemplate.Contains("172") Then
                oWb = oXl.Workbooks.Open(mytemplate)
            Else
                oWb = oXl.Workbooks.Open(Application.StartupPath & mytemplate)
            End If

            oXl.Visible = False

            ProgressReport(2, "Creating Worksheet...")
            oWb.Worksheets(1).select()
            FormatReportCallback.Invoke(oWb, New EventArgs)

            StopWatch.Stop()

            'FileName = FileName & "\" & String.Format("Report" & ReportName & "-{0}-{1}-{2}.xlsx", Today.Year, Format("00", Today.Month), Format("00", Today.Day))
            FileName = FileName & "\" & String.Format(ReportName)
            ProgressReport(3, "")
            ProgressReport(2, "Saving File ..." & FileName)
            'oSheet.Name = ReportName
            If FileName.Contains("xlsm") Then

                oWb.SaveAs(FileName, FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)
            Else
                oWb.SaveAs(FileName)
            End If

            ProgressReport(2, "Elapsed Time: " & Format(StopWatch.Elapsed.Minutes, "00") & ":" & Format(StopWatch.Elapsed.Seconds, "00") & "." & StopWatch.Elapsed.Milliseconds.ToString)
            result = True
        Catch ex As Exception
            ProgressReport(3, ex.Message & FileName)
            errorMsg = ex.Message
        Finally
            'oXl.ScreenUpdating = True
            'clear excel from memory
            Try
                oXl.Quit()
                releaseComObject(oSheet)
                releaseComObject(oWb)
                releaseComObject(oXl)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception

            End Try

            Try
                'to make sure excel is no longer in memory
                EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try

        End Try
        Return result
    End Function
    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Parent.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Try
                Parent.Invoke(d, New Object() {id, message})
            Catch ex As Exception
                ProgressReport(2, ex.Message)
            End Try

        Else
            Try
                Select Case id
                    Case 2
                        Parent.ToolStripStatusLabel1.Text = message
                    Case 3
                        Parent.ToolStripStatusLabel2.Text = Trim(message)
                    Case 4
                        Parent.close()
                    Case 5
                        Parent.ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        Parent.ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                End Select
            Catch ex As Exception
                ProgressReport(2, ex.Message)
            End Try


        End If

    End Sub

    Public Shared Sub FillWorksheet(ByVal osheet As Excel.Worksheet, ByVal sqlstr As String, Optional ByVal Location As String = "A1")
        'Dim oRange As Excel.Range
        Dim oExCon As String = My.Settings.oExCon ' My.Settings.oExCon.ToString '"ODBC;DSN=PostgreSQL30;"
        'oExCon = oExCon.Insert(oExCon.Length, "UID=" & dbadapter1.userid & ";PWD=" & dbadapter1.password)
        Dim dbAdapter1 = PostgreSQLDBAdapter.getInstance
        oExCon = oExCon.Insert(oExCon.Length, "UID=" & dbAdapter1.UserId & ";PWD=" & dbAdapter1.Password)
        Dim oRange As Excel.Range
        oRange = osheet.Range(Location)
        With osheet.QueryTables.Add(oExCon.Replace("Host=", "Server="), oRange)
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

        oRange = osheet.Range("1:1")
        oRange = osheet.Range(Location)
        oRange.Select()
        osheet.Application.Selection.autofilter()
        osheet.Cells.EntireColumn.AutoFit()
    End Sub

    Public Shared Function GetLastRow(ByVal oxl As Excel.Application, ByVal osheet As Excel.Worksheet, ByVal range As Excel.Range) As Long
        Dim lastrow As Long = 1
        oxl.ScreenUpdating = False
        Try
            lastrow = osheet.Cells.Find("*", range, , , Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious).Row
        Catch ex As Exception
        End Try
        Return lastrow
        oxl.ScreenUpdating = True
    End Function

    Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)

    End Sub

    Sub FinishReport(ByRef sender As Object, ByRef e As EventArgs)

    End Sub

    Public Shared Sub releaseComObject(ByRef o As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
        Catch ex As Exception
        Finally
            o = Nothing
        End Try
    End Sub
    Public Sub RunConvertToTextFile(ByRef sender As System.Object, ByVal e As ExportToExcelWithKeyEventArgs)
        If Not myThread.IsAlive Then
            Try
                Me.myKey = e.myKey
                myThread = New System.Threading.Thread(New ThreadStart(AddressOf DoConvertFile))
                myThread.SetApartmentState(ApartmentState.MTA)
                myThread.Start()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub

    Private Function DoConvertFile()
        Parent.progressreport(1, Directory)
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr
        Dim myret As Boolean = False
        Dim fullname As String = String.Empty
        Try
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            oXl.Visible = False
            oXl.DisplayAlerts = False
            oWb = oXl.Workbooks.Open(Directory)

            'Check File
            Dim myConvertFileEventArgs As New ConvertFileEventArgs
            ValidateWorksheetCallback.Invoke(oWb, myConvertFileEventArgs)

            'fullname = String.Format("{0}\{1}.csv", IO.Path.GetDirectoryName(Directory), IO.Path.GetFileNameWithoutExtension(Directory))
            fullname = String.Format("{0}\{1}.csv", IO.Path.GetDirectoryName(Directory), myConvertFileEventArgs.FileName)
            oWb.SaveAs(Filename:=fullname, FileFormat:=Excel.XlFileFormat.xlUnicodeText, CreateBackup:=False)
            'oWb.SaveAs(Filename:=fullname, FileFormat:=62, CreateBackup:=False) 'Excel.xlFileFormat.xlCSVUTF8
            myret = True
        Catch ex As Exception
            Parent.progressreport(1, ex.Message)
        Finally
            oXl.Quit()
            ExportToExcelFile.releaseComObject(oSheet)
            ExportToExcelFile.releaseComObject(oWb)
            ExportToExcelFile.releaseComObject(oXl)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            Try
                'to make sure excel is no longer in memory
                ExportToExcelFile.EndTask(hwnd, True, True)
                Parent.progressreport(2005, String.Format("{0}", fullname))
            Catch ex As Exception
            End Try
        End Try
        Return myret
    End Function

    Public Sub ExtractFromDataTableUnsyncDT(ByRef sender As Object, ByVal e As System.EventArgs)
        Try
            DoExtractFromDataTableDT()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub DoExtractFromDataTableDT()
        Dim errMsg As String = String.Empty
        Dim i As Integer = 0
        Dim errSB As New StringBuilder
        Dim sw As New Stopwatch
        sw.Start()
        ProgressReport(2, "Export To Excel..")
        ProgressReport(6, "Marques..")
        status = GenerateReportDT(Directory, errMsg)
        ProgressReport(5, "Continues..")
        If status Then


            sw.Stop()
            ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            ProgressReport(3, "")
            If AskOpenFile Then
                If MsgBox("File name: " & Directory & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                    Process.Start(Directory)
                End If
            End If
            ProgressReport(3, "")
            'ProgressReport(4, errSB.ToString)
        Else
            errSB.Append(errMsg) '& vbCrLf)
            ProgressReport(3, errSB.ToString)
        End If
        sw.Stop()
    End Sub

    Private Function GenerateReportDT(ByRef FileName As String, ByRef errorMsg As String) As Boolean
        Dim myCriteria As String = String.Empty
        Dim result As Boolean = False

        Dim StopWatch As New Stopwatch
        StopWatch.Start()
        'Open Excel
        Application.DoEvents()

        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr
        Try
            'Create Object Excel 
            ProgressReport(2, "CreateObject..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            'oXl.ScreenUpdating = False
            'oXl.Visible = False
            oXl.DisplayAlerts = False
            ProgressReport(2, "Opening Template...")
            ProgressReport(2, "Generating records..")
            oWb = oXl.Workbooks.Open(Application.StartupPath & mytemplate)
            oSheet = oWb.Worksheets(1)
            oSheet.Name = "RAWDATA"
            oWb.SaveAs(ReportName)
            'For i = 0 To oWb.Worksheets.Count - 1
            '    oWb.Sheets(1).Delete()
            'Next
            oXl.Visible = False

            oXl.Quit()
            releaseComObject(oSheet)
            releaseComObject(oWb)
            releaseComObject(oXl)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            EndTask(hwnd, True, True)



            'ProgressReport(2, "Creating Worksheet...")
            'DATA

            If IsNothing(DataTableList) Then
                'Print One Worksheet 
                'oWb.Worksheets(Datasheet).select()
                'oSheet = oWb.Worksheets(Datasheet)

                If Me.DT.Rows.Count > 0 Then
                    ProgressReport(2, "Get records..")
                    'FillWorksheet(oSheet, DT)
                    FillWorksheetDT(ReportName, DT)
                    'Dim oXl1 As Excel.Application = Nothing
                    oXl = CType(CreateObject("Excel.Application"), Excel.Application)
                    hwnd = oXl.Hwnd
                    oXl.DisplayAlerts = False
                    'Dim oWb1 As Excel.Workbook = Nothing
                    'Dim oSheet1 As Excel.Worksheet = Nothing
                    oWb = oXl.Workbooks.Open(ReportName)
                    oSheet = oWb.Worksheets(Datasheet)
                    Dim orange = oSheet.Range("A1")
                    Dim lastrow = GetLastRow(oXl, oSheet, orange)
                    If lastrow > 1 Then
                        FormatReportCallback.Invoke(oSheet, New EventArgs)
                    End If
                Else
                    'FormatReportCallback.Invoke(oSheet, New EventArgs)
                End If
            Else
                'Print Multiple Worksheet
                'Looping from here
                For i = 0 To DataTableList.Count - 1
                    Dim MyDataTable = CType(DataTableList(i), DataTableWorksheet)
                    oWb.Worksheets(MyDataTable.DataSheet).select()
                    oSheet = oWb.Worksheets(MyDataTable.DataSheet)
                    oSheet.Name = MyDataTable.SheetName
                    ProgressReport(2, "Get records..")
                    'FillWorksheet(oSheet, MyDataTable.DataTable)
                    FillWorksheetDT(ReportName, MyDataTable.DataTable)
                    Dim orange = oSheet.Range("A1")
                    Dim lastrow = GetLastRow(oXl, oSheet, orange)


                    If lastrow > 1 Then
                        'Delegate for modification
                        'oSheet.Columns("A:A").numberformat = "dd-MMM-yyyy"
                        FormatReportCallback.Invoke(oSheet, New EventArgs)
                    End If
                Next



                'End Looping
            End If


            PivotCallback.Invoke(oWb, New EventArgs)
            'For i = 0 To oWb.Connections.Count - 1
            '    oWb.Connections(1).Delete()
            'Next
            StopWatch.Stop()

            'FileName = FileName & "\" & String.Format("Report" & ReportName & "-{0}-{1}-{2}.xlsx", Today.Year, Format("00", Today.Month), Format("00", Today.Day))
            FileName = String.Format(ReportName)
            ProgressReport(3, "")
            ProgressReport(2, "Saving File ..." & FileName)
            'oSheet.Name = ReportName
            If FileName.Contains("xlsm") Then
                oWb.SaveAs(FileName, FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)
            Else
                oWb.SaveAs(FileName)
            End If

            ProgressReport(2, "Elapsed Time: " & Format(StopWatch.Elapsed.Minutes, "00") & ":" & Format(StopWatch.Elapsed.Seconds, "00") & "." & StopWatch.Elapsed.Milliseconds.ToString)
            result = True
        Catch ex As Exception
            ProgressReport(3, ex.Message & FileName)
            errorMsg = ex.Message
        Finally
            'clear excel from memory
            Try
                oXl.Quit()
                releaseComObject(oSheet)
                releaseComObject(oWb)
                releaseComObject(oXl)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception

            End Try

            Try
                'to make sure excel is no longer in memory
                EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try

        End Try
        Return result
    End Function

    Private Sub FillWorksheetDT(Filename As String, dataTable As DataTable)
        Dim driver As String = "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)"
        Dim ConStr As String = String.Format("Driver={0};Dbq={1};ReadOnly=0;", driver, Filename)

        Using conn As New Odbc.OdbcConnection(ConStr)
            conn.Open()
            'Create Table Query
            Dim strTableQ As String
            Dim colSB As New System.Text.StringBuilder
            'Dim j As Integer = 0
            For j = 0 To dataTable.Columns.Count - 1
                Dim dCol As DataColumn
                dCol = dataTable.Columns(j)
                Dim coltype As String = String.Empty
                Select Case dCol.DataType.Name
                    Case "Boolean"
                        coltype = "int"
                    Case "String"
                        coltype = "TEXT"
                    Case "DateTime"
                        coltype = "DATE"
                    Case "Int32"
                        coltype = "int"
                    Case "Decimal"
                        coltype = "real"
                    Case Else
                End Select
                If colSB.Length > 0 Then
                    colSB.Append(",")
                End If
                colSB.Append(String.Format("[{0}] {1}", dCol.ColumnName, coltype))
            Next
            Dim mysheetname = "RAWDATA"
            'strTableQ = String.Format("CREATE TABLE [Sheet1$]({0})", colSB.ToString)
            strTableQ = String.Format("CREATE TABLE [{0}$]({1})", mysheetname, colSB.ToString)

            Using cmd As New Odbc.OdbcCommand(strTableQ, conn)
                cmd.ExecuteNonQuery()
            End Using

            'Insert Query
            Dim sbInsert As New StringBuilder
            Dim strInsertQ As String
            For k As Integer = 0 To dataTable.Columns.Count - 1
                If sbInsert.Length > 0 Then
                    sbInsert.Append(",")
                End If
                sbInsert.Append("?")
            Next
            'strInsertQ = String.Format("Insert Into [Sheet1$] Values({0})", sbInsert.ToString)
            strInsertQ = String.Format("Insert Into [{0}$] Values({1})", mysheetname, sbInsert.ToString)

            'Parameters Query
            For j = 0 To dataTable.Rows.Count - 1
                Using cmd = New Odbc.OdbcCommand(strInsertQ, conn)
                    For k As Integer = 0 To dataTable.Columns.Count - 1
                        cmd.Parameters.AddWithValue("?", dataTable.Rows(j)(k))
                        Select Case dataTable.Columns(k).DataType.Name
                            Case "String"
                                cmd.Parameters(k).DbType = DbType.String
                            Case "DateTime"
                                cmd.Parameters(k).DbType = DbType.Date
                            Case "Int32"
                                cmd.Parameters(k).DbType = DbType.Int32
                            Case "Decimal"
                                cmd.Parameters(k).DbType = DbType.Double
                            Case "Boolean"
                                cmd.Parameters(k).DbType = DbType.Int32
                            Case Else
                        End Select
                    Next
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                End Using
            Next
        End Using
    End Sub


End Class
Public Class ConvertFileEventArgs
    Inherits EventArgs
    Public Property FileName As String
End Class
Public Class ExportToExcelFileEventArgs
    Inherits EventArgs
    Public Property lastRow As Integer
    Public Property SheetNo As Integer
End Class

Public Class ExportToExcelWithKeyEventArgs
    Inherits EventArgs
    Public Property myKey As Object
End Class

Public Class QueryWorksheet
    Public Property DataSheet As Integer
    Public Property Sqlstr As String
    Public Property SheetName As String
End Class

Public Class DataTableWorksheet
    Public Property DataSheet As Integer
    Public Property DataTable As DataTable
    Public Property SheetName As String
End Class