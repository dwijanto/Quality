Imports System.Threading
Imports Microsoft.Office.Interop
Imports System.Text

Public Class GenerateExcel

    Dim myAdapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance
    Private DS As DataSet
    Dim _ErrorMessage As String
    Dim myform As Object
    Dim myParam As ParamAdapter = ParamAdapter.getInstance
    Dim DirectoryName As String = myParam.GetMailFolder("OUTBOX")
    Dim datespan As Integer = myParam.GetCCETDDateSpan
    Private MyDate As Date = Date.Today

    Public Sub New(ByVal parent As Object)
        Me.myform = parent
    End Sub

    Public Sub New(ByVal parent As Object, mydate As Date)
        Me.myform = parent
        Me.MyDate = mydate
    End Sub

    Public ReadOnly Property ErrorMessage As String
        Get
            Return _ErrorMessage
        End Get
    End Property

    Public Function run(ByVal ds As DataSet) As Boolean
        Dim myret As Boolean = True
        Me.DS = ds
        Dim mycount As Integer

        'Get total Process for control
        For i = 0 To ds.Tables(0).Rows.Count - 1
            If ds.Tables(0).Rows(i).Item("selected") Then
                mycount = mycount + 1
            End If
            ds.Tables(0).Rows(i).Item("status") = False
        Next
        myform.ProgressReport(2004, String.Format("{0}", mycount))

        For i = 0 To ds.Tables(0).Rows.Count - 1
            If ds.Tables(0).Rows(i).Item("selected") Then
                Thread.Sleep(1000) 'slowdown the process 100
                CreateExcel(ds.Tables(0).Rows(i).Item("vendor"), i + 1)
            End If
        Next

        Return myret
    End Function
    Public Function Run() As Boolean
        Dim myret As Boolean = True
        'Get Vendor from dailytask
        DS = New DataSet
        'Dim sqlstr = "select distinct vendor from quality.dailytx where code isnull and vendor < 20000000 order by vendor;" 'Later you can join in master vendor to get the valid Vendor
        Dim sqlstr = "select distinct vendor,code from quality.dailytx tx" &
                     " inner join quality.vendor v on v.vendorcode = tx.vendor" &
                     " where code isnull order by vendor limit 5"
        Try
            If myAdapter.GetDataset(sqlstr, DS) Then
                For i = 0 To DS.Tables(0).Rows.Count - 1
                    CreateExcel(DS.Tables(0).Rows(i).Item("vendor"), i + 1)
                Next
            End If
        Catch ex As Exception
            myret = False
            _ErrorMessage = ex.Message
        End Try

        Return myret
    End Function

    Private Sub CreateExcel(ByVal vendorcode As Long, ByVal count As Integer)
        Dim SBException As New StringBuilder
        Dim SBUExceptionList = myParam.GetSBUExceptionList
        If SBUExceptionList.Length > 0 Then
            SBException.Append(String.Format(" and ( tx.sbu not in ({0}))", SBUExceptionList))
        End If
        'Dim sqlstr = String.Format("select purchdoc as ""Purch.Doc."",item as ""Item"",seqn as ""SeqN"",insplot as ""Insp. Lot"", " &
        '                           " inspector as ""Inspector"",code as ""Inspection Result"",vendor as ""Vendor"",vendorname as ""Vendor Name"",material as ""Material"", materialdesc as ""Material desc""," &
        '                           " custpono as ""Cust PO No"",sbu as ""SBU"",city as ""City"",ccetd as ""Confirmed ETD"",qty as ""Quantity"", qtyoun as	""OUn""," &
        '                           " quality.changesamplesize(inspector,samplesize::integer) as ""Sample size"",null::date as""New Input Inspection Date"",quality.getinspdate(purchdoc,item,seqn,qty) as ""Inspection Date"",null::text as ""Remarks"" from quality.dailytx tx where vendor = {0} and (code isnull or code in ('UF','DF','PP')) and ccetd >= current_date - {1} {2}", vendorcode, datespan, SBException.ToString)
        'Dim sqlstr = String.Format("select tx.purchdoc as ""Purch.Doc."",tx.item as ""Item"",tx.seqn as ""SeqN"",insplot as ""Insp. Lot"", " &
        '                           " inspector as ""Inspector"",code as ""Inspection Result"",vendor as ""Vendor"",vendorname as ""Vendor Name"",material as ""Material"", materialdesc as ""Material desc""," &
        '                           " custpono as ""Cust PO No"",sbu as ""SBU"",city as ""City"",ccetd as ""Confirmed ETD"",qty as ""Quantity"", qtyoun as	""OUn""," &
        '                           " quality.changesamplesize(inspector,samplesize::integer) as ""Sample size"",null::date as""New Input Inspection Date"",quality.getinspdate(tx.purchdoc,tx.item,tx.seqn,qty) as ""Inspection Date"",null::text as ""Remarks"",quality.getproductionenddate(tx.purchdoc,tx.item,tx.seqn,qty)  as ""Production End Date"" from quality.dailytx tx" &
        '                           " where vendor = {0} and (code isnull or code in ('UF','DF','PP')) and ccetd >= current_date - {1} {2}", vendorcode, datespan, SBException.ToString)

        'Dim sqlstr = String.Format("select tx.purchdoc as ""Purch.Doc."",tx.item as ""Item"",tx.seqn as ""SeqN"",insplot as ""Insp. Lot"", " &
        '                           " inspector as ""Inspector"",code as ""Inspection Result"",vendor as ""Vendor"",vendorname as ""Vendor Name"",material as ""Material"", materialdesc as ""Material desc""," &
        '                           " custpono as ""Cust PO No"",sbu as ""SBU"",city as ""City"",ccetd as ""Confirmed ETD"",qty as ""Quantity"", qtyoun as	""OUn""," &
        '                           " quality.changesamplesize(inspector,samplesize::integer) as ""Sample size"",null::date as""New Input Inspection Date"",quality.getinspdate(tx.purchdoc,tx.item,tx.seqn,qty) as ""Inspection Date"",null::text as ""Remarks"",quality.getproductionenddate(tx.purchdoc,tx.item,tx.seqn,qty)  as ""Production End Date"" from quality.dailytx tx" &
        '                           " where vendor = {0} and ccetd >= current_date - {1} {2} order by ccetd asc", vendorcode, datespan, SBException.ToString)
        Dim sqlstr = String.Format("select tx.purchdoc as ""Purch.Doc."",tx.item as ""Item"",tx.seqn as ""SeqN"",insplot as ""Insp. Lot"", " &
                                   " inspector as ""Inspector"",code as ""Inspection Result"",vendor as ""Vendor"",vendorname as ""Vendor Name"",material as ""Material"", materialdesc as ""Material desc""," &
                                   " custpono as ""Cust PO No"",sbu as ""SBU"",city as ""City"",ccetd as ""Confirmed ETD"",qty as ""Quantity"", qtyoun as	""OUn""," &
                                   " quality.changesamplesize(inspector,samplesize::integer) as ""Sample size"",null::date as""New Input Inspection Date"",quality.getinspdate(tx.purchdoc,tx.item,tx.seqn,qty) as ""Inspection Date"",null::text as ""Remarks"",quality.getproductionenddate(tx.purchdoc,tx.item,tx.seqn,qty)  as ""Production End Date"" ,quality.getremarks(tx.purchdoc,tx.item,tx.seqn) as ""Remarks History"" from quality.dailytx tx" &
                                   " where vendor = {0} and ccetd >= current_date - {1} {2} order by tx.purchdoc,tx.item,tx.seqn", vendorcode, datespan, SBException.ToString)

        'Dim ReportName = String.Format("{0:yyyyMMdd}-{1}", Today.Date, vendorcode)
        Dim ReportName = String.Format("{0:yyyyMMdd}-{1}", MyDate.Date, vendorcode)
        Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
        Dim PivotCallback As FormatReportDelegate = AddressOf MyPivotTable

        Dim myreport As New ExportToExcelFile(myform, sqlstr, DirectoryName, ReportName, mycallback, PivotCallback, 1, "\templates\ExcelTemplate.xltx")
        myreport.AskOpenFile = False
        myform.ProgressReport(1, String.Format("Working on Vendorcode: {0}  - {1}/{2}", vendorcode, count, DS.Tables(0).Rows.Count))

        'myreport.Run(myform, New EventArgs)
        'myreport.DoWork()
        myreport.RunWithKey(myform, New ExportToExcelWithKeyEventArgs With {.myKey = vendorcode})
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        Dim osheet As Excel.Worksheet = DirectCast(sender, Excel.Worksheet)
        'osheet.Name = String.Format("Schedule_{0:yyyyMMdd}-{1}", Date.Today, osheet.Range("G2").Value.ToString)
        osheet.Name = String.Format("Schedule_{0:yyyyMMdd}-{1}", MyDate.Date, osheet.Range("G2").Value.ToString)
        osheet.Range("R:R").Locked = False
        With osheet.Columns("R:R").interior
            .Pattern = Excel.Constants.xlSolid
            .PatternColorIndex = Excel.Constants.xlAutomatic
            .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With


        Dim strdata As String = osheet.UsedRange.Address
        Dim maX1 As Integer = CLng(Mid(strdata, InStrRev(strdata, "$")))

        'Data Validation For Date
        osheet.Columns("R:R").Validation.Delete()
        With osheet.Range("R2:R" & maX1).Validation
            .Add(Type:=Excel.XlDVType.xlValidateDate, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop, Operator:=Excel.XlFormatConditionOperator.xlGreaterEqual, Formula1:=String.Format("{0:MM/dd/yyyy}", Today.Date.AddDays(1)))
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = "Value must be in date format and bigger than today."
            .ShowInput = True
            .ShowError = True
        End With

        'osheet.Range("R2").Copy()

        'Dim strdata As String = osheet.UsedRange.Address
        'Dim maX1 As Integer = CLng(Mid(strdata, InStrRev(strdata, "$")))
        ''osheet.Range(osheet.Range("R2"), osheet.Range("R2").End(Excel.XlDirection.xlDown)).Select()
        'osheet.Range(osheet.Cells(2, 18), osheet.Cells(maX1, 18)).Select()
        'osheet.Paste()

        osheet.Range("T:U").Locked = False
        With osheet.Columns("T:U").interior
            .Pattern = Excel.Constants.xlSolid
            .PatternColorIndex = Excel.Constants.xlAutomatic
            .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With

        'Data Validation For Date
        osheet.Columns("U:U").Validation.Delete()
        'With osheet.Range(osheet.Cells(2, 21), osheet.Cells(maX1, 21)).Validation
        With osheet.Range("U2:U" & maX1).Validation
            .Add(Type:=Excel.XlDVType.xlValidateDate, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop, Operator:=Excel.XlFormatConditionOperator.xlGreaterEqual, Formula1:=String.Format("{0:MM/dd/yyyy}", CDate("2010-01-01")))
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = "Value must be in date format and bigger than 2010-01-01"
            .ShowInput = True
            .ShowError = True
        End With

        'osheet.Range("U2").Copy()
        'osheet.Range(osheet.Cells(2, 21), osheet.Cells(maX1, 21)).Select()
        'osheet.Paste()

        osheet.Columns("T:T").ColumnWidth = 40
        osheet.Columns("N:N").NumberFormat = "dd-MMM-yyyy"
        osheet.Columns("R:S").NumberFormat = "dd-MMM-yyyy"
        osheet.Columns("U:U").NumberFormat = "dd-MMM-yyyy"

        osheet.Columns("V:V").WrapText = True
        osheet.Columns("V:V").ColumnWidth = 60
        osheet.UsedRange.RowHeight = 15
        osheet.Protect(Password:="Quality123", AllowSorting:=True, AllowFiltering:=True)

    End Sub

    Private Sub MyPivotTable(ByRef obj As Object, ByRef e As EventArgs)
        Dim owb As Excel.Workbook = DirectCast(obj, Excel.Workbook)
        owb.Protect(Password:="Quality123", [Structure]:=True)
    End Sub

End Class
