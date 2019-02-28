Imports Microsoft.Office.Interop
Imports System.Threading
Imports System.Text
Imports System.IO

Public Class GetReply
    Inherits BaseImport
    Dim myParam As ParamAdapter = ParamAdapter.getInstance
    Dim INBOXFolder As String = myParam.GetMailFolder("INBOX")
    Dim totalFile As Integer
    Dim myform As Object
    Dim myTexTFile As New List(Of String)
    Dim Recordsb As New StringBuilder
    Dim deleteSB As New StringBuilder
    Dim myThread As New Thread(AddressOf DoImport)
    Dim myAdapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance

    Public Sub New(ByVal parent As Object)
        Me.myform = parent
    End Sub

    Public Sub run()
        Dim dir As New IO.DirectoryInfo(INBOXFolder)
        Dim arrFI As IO.FileInfo() = dir.GetFiles("*.xls*") '.Union(dir.GetFiles("*.xls")).ToArray
        totalFile = arrFI.Count
        myform.ProgressReport(2004, String.Format("{0}", totalFile))
        'Convert To Text
        For Each fi As IO.FileInfo In arrFI
            Thread.Sleep(100)
            convertToText(fi.FullName)
        Next


    End Sub

    Public Sub ImportFile(TextFile As List(Of String))
        Me.myTexTFile = TextFile
        If Not myThread.IsAlive Then
            myform.ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoImport)
            myThread.Start()
        Else
            myform.ProgressReport(1, "Please wait until the current process finished.")
        End If
    End Sub



    Private Sub convertToText1(FileNameFullPath As String)

        myform.progressreport(1, FileNameFullPath)
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr
        Dim myret As Boolean = False
        Try
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            oXl.Visible = False
            oXl.DisplayAlerts = False
            oWb = oXl.Workbooks.Open(FileNameFullPath)
            'Check FileType
            oWb.Worksheets(1).select()
            oSheet = oWb.Worksheets(1)
            Dim buff = oSheet.Name.Split("_")
            If Not buff(0) = "Schedule" Then
                Exit Sub
            End If
            'Dim fullname As String = String.Format("{0}\{1}.csv", INBOXFolder, IO.Path.GetFileNameWithoutExtension(FileNameFullPath))
            Dim fullname As String = String.Format("{0}\{1}.csv", INBOXFolder, buff(1))

            oWb.SaveAs(Filename:=fullname, FileFormat:=Excel.XlFileFormat.xlUnicodeText, CreateBackup:=False)
            'oWb.SaveAs(Filename:=fullname, FileFormat:=62, CreateBackup:=False) '62 = Excel.xlFileFormat.xlCSVUTF8
        Catch ex As Exception

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
            Catch ex As Exception
            End Try
        End Try
    End Sub

    Private Sub convertToText(Fullname As String)
        Dim myExcel As New ExportToExcelFile(myform, Fullname, AddressOf ValidateWorksheet)
        myExcel.RunConvertToTextFile(myform, New ExportToExcelWithKeyEventArgs)
    End Sub

    Private Sub ValidateWorksheet(ByRef sender As Object, ByRef e As ConvertFileEventArgs)
        Dim owb = DirectCast(sender, Excel.Workbook)
        owb.Worksheets(1).select()
        Dim oSheet = owb.Worksheets(1)
        Dim buff = oSheet.Name.Split("_")
        If Not buff(0) = "Schedule" Then
            Err.Raise("File is not valid.")
        End If
        e.FileName = buff(1)

    End Sub

    Sub DoImport()
        myform.ProgressReport(1, "Importing file....")
        Dim myErrorSB As New StringBuilder
        Recordsb = New StringBuilder
        deleteSB = New StringBuilder
        Try
            For Each t In myTexTFile
                'myform.ProgressReport(1, String.Format("Importing file....{0}", t))
                Dim myrecord() As String
                Dim myList As New List(Of String())
                Dim buff = IO.Path.GetFileNameWithoutExtension(t).Split("-")
                Dim DocDate As String = buff(0)

                Using objTFParser = New FileIO.TextFieldParser(t)
                    With objTFParser
                        .TextFieldType = FileIO.FieldType.Delimited
                        .SetDelimiters(Chr(9))
                        '.SetDelimiters(",")
                        .HasFieldsEnclosedInQuotes = True
                        Dim count As Long = 0
                        myform.ProgressReport(1, "Read Data..")
                        Do Until .EndOfData
                            myrecord = .ReadFields
                            If count > 0 Then
                                If myrecord(17) <> "" Or myrecord(19) <> "" Or myrecord(20) <> "" Then
                                    myList.Add(myrecord)
                                End If
                            Else
                                count += 1
                            End If
                        Loop
                    End With
                    'Do we have data to add?
                    If myList.Count > 0 Then
                        'need to delete first based on DocDate and Vendor
                        deleteSB.Append(String.Format("delete from quality.historytx where docdate = '{0}' and vendor = {1};", buff(0), buff(1)))
                    End If
                    myform.ProgressReport(1, "Build Record..")
                    For i = 0 To myList.Count - 1
                        'CreateRecord(myList(i))
                        If Not myList(i)(0) = "" Then
                            Dim mycheck As Boolean = True
                            If Not checkdate(myList(i)(13)) Then
                                mycheck = False
                                myErrorSB.Append(String.Format("{0} CCETD: {1}{2}", t, myList(i)(13).ToString, vbCrLf))
                            End If

                            If Not checkdate(myList(i)(17)) Then
                                mycheck = False
                                myErrorSB.Append(String.Format("{0} New Inspection Date: {1}{2}", t, myList(i)(17).ToString, vbCrLf))
                            End If
                            If Not checkdate(myList(i)(20)) Then
                                mycheck = False
                                myErrorSB.Append(String.Format("{0} Production End Date: {1}{2}", t, myList(i)(20).ToString, vbCrLf))
                            End If
                            If mycheck Then
                                Recordsb.Append(myList(i)(0) & vbTab &
                                        myList(i)(1) & vbTab &
                                        myList(i)(2) & vbTab &
                                        myList(i)(6) & vbTab &
                                        validDate(myList(i)(13)) & vbTab &
                                        myList(i)(14) & vbTab &
                                        validDate(myList(i)(17)) & vbTab &
                                        validStr(myList(i)(19)) & vbTab &
                                        validDate(buff(0)) & vbTab &
                                        validDate(myList(i)(20)) & vbCrLf)
                            End If

                        End If
                    Next
                End Using


                Kill(t)
            Next

            If myErrorSB.Length > 0 Then
                'Using mystream As New StreamWriter(Application.StartupPath & "\error.txt")
                Using mystream As New StreamWriter(INBOXFolder & "\error.txt")
                    mystream.WriteLine(myErrorSB.ToString)
                End Using
                'Process.Start(Application.StartupPath & "\error.txt")
                Process.Start(INBOXFolder & "\error.txt")              
                myform.ProgressReport(1, "Error Found.")
            Else
                'Copy TX
                If CopyTx() Then
                    myform.ProgressReport(1, "Import Done.")
                End If
            End If



        Catch ex As Exception
            myform.ProgressReport(1, ex.Message)
        End Try
    End Sub
    Private Function checkdate(p1 As String) As Boolean
        Dim myret = True
        If p1.Length > 0 Then
            myret = IsDate(p1)
        End If
        Return myret
    End Function

    Private Sub CreateRecord(myRecord As String())

    End Sub

    Public Function CopyTx() As Boolean

        myform.ProgressReport(1, "Copy records.")

        Dim myret As Boolean = True
        If Recordsb.Length > 0 Then
            Dim Sqlstr As String = String.Empty

            Dim message As String = String.Empty
            If Recordsb.Length > 0 Then

                Sqlstr = deleteSB.ToString & "begin;set statement_timeout to 0;end;copy quality.historytx( purchdoc,item , seqn ,vendor,ccetd, qty , inspdate ,remark, docdate,productionenddate)  from stdin with null as 'Null';"
                message = myAdapter.copy(Sqlstr, Recordsb.ToString, myret)
                If Not myret Then
                    myform.ProgressReport(1, message)
                End If
            End If
        End If
        Return myret

    End Function





End Class
