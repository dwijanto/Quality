Imports System.Threading
Imports System.Text
Imports Microsoft.Exchange.WebServices.Data
Imports Microsoft.Office.Interop

Public Class SendEmailConfirmation
    Private myForm As Object
    Private DS As DataSet
    Dim myCallback As System.Net.Mail.SendCompletedEventHandler
    Dim ExcelCallback As FormatReportDelegate = AddressOf FormattingReport
    Dim PivotCallback As FormatReportDelegate = AddressOf MyPivotTable

    Dim myParam As ParamAdapter = ParamAdapter.getInstance
    Dim DirectoryName As String = myParam.GetMailFolder("OUTBOX")

    Private startdate As Date
    Private enddate As Date
    Private LastSendDate As Date
    Dim myModel As SendEmailModel
    Dim EWSUser As String() = myParam.GetEWSUser("EWS")
    Private _runEWS As Boolean

    Dim Inspectiondate As Date

    Public Sub New()

    End Sub

    Public Sub New(ByVal parent As Object, ByVal mycallback As System.Net.Mail.SendCompletedEventHandler, ByVal startdate As Date, ByVal enddate As Date)
        Me.myForm = parent
        Me.myCallback = mycallback
        Me.startdate = startdate
        Me.enddate = enddate

    End Sub

    Public Sub New(ByVal parent As Object, ByVal mycallback As System.Net.Mail.SendCompletedEventHandler)
        Me.myForm = parent
        Me.myCallback = mycallback
        Inspectiondate = parent.Inspectiondate
    End Sub

    Function runEWS(DataTable As DataTable, InspectionDate As Date) As Boolean
        Dim myret As Boolean = False
        '1.Send email for Each Vendor
        If Not SendEmailConfirmationVendorManually(DataTable, InspectionDate) Then
            Return myret
        End If

        '2.Send email for internal 
        'DataTable already filtered
        Dim dv As New DataView(DataTable)
        dv.RowFilter = "selected = true"
        Dim myDataTable As DataTable = dv.ToTable


        If Not SendEmailConfirmationInternalManually(myDataTable, InspectionDate) Then
            Return myret
        End If





        myret = True
        Return myret
    End Function

    Public Function runEWS(ByVal ds As DataSet) As Boolean
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
        myForm.ProgressReport(2004, String.Format("{0}", mycount))

        For i = 0 To ds.Tables(0).Rows.Count - 1
            If ds.Tables(0).Rows(i).Item("selected") Then
                Thread.Sleep(100) 'slowdown the process
                'CreateEmailEWS(ds.Tables(0).Rows(i).Item("vendor"), i + 1)
                mymodel = New SendEmailModel(Inspectiondate)
                If ds.Tables(0).Rows(i).Item("vendor") = 0 Then
                    CreateEmailEWSInternal(ds.Tables(0).Rows(i), i + 1)
                Else
                    CreateEmailEWS(ds.Tables(0).Rows(i), i + 1)
                End If

            End If
        Next

        'Save the send Last confirmation date in myForm.ProgressReport (Completed)


        Return myret
    End Function

    Private Sub CreateEmailEWSInternal(dr As DataRow, p2 As Integer)

        Dim myEmail As New EmailEWS(myForm)
        'myEmail.urlEWS = myParam.GeturlEWS("urlEWS")
        'myEmail.ToRecepients = "etai@groupeseb.com;dlie@groupeseb.com" 'myParam.GetInternalEmail("Internal Email (to)") 
        'myEmail.CcRecepients = "etai@groupeseb.com;dlie@groupeseb.com" 'myParam.GetInternalEmail("Internal Email (cc)")

        myEmail.ToRecepients = myParam.GetInternalEmail("Internal Email (to)")
        myEmail.CcRecepients = myParam.GetInternalEmail("Internal Email (cc)")

        myEmail.Username = EWSUser(0)
        myEmail.password = EWSUser(1)
        myEmail.domain = "AS"
        'myEmail.Subject = String.Format(String.Format("Inspection Schedule : {0:dd-MMM-yyyy} - {1:dd-MMM-yyyy}", myForm.myController.lastsenddate, myForm.myController.enddate))
        myEmail.Subject = String.Format(String.Format("Inspection Schedule : {0:dd-MMM-yyyy} ", myForm.myController.startdate))

        myEmail.BodyMessage = GetBodyMessage(dr.Item("vendor"))

        ''Create Attachment
        Dim ReportName As String = "Internal.xlsx"
        Dim myreport As New ExportToExcelFile(myForm, mymodel.InternalSqlStr, DirectoryName, ReportName, ExcelCallback, PivotCallback, 1, "\templates\ExcelTemplate.xltx")
        myreport.AskOpenFile = False
        'myreport.Run(myForm, New EventArgs)
        myreport.DoWork()
        Try

            myEmail.AttachmentList.Add(String.Format("{0}\{1}", DirectoryName, ReportName))
            'Kill(String.Format("{0}\{1}", DirectoryName, ReportName))
        Catch ex As Exception
        End Try

        myEmail.RunWithKey(myForm, New GenerateEmailWithKeyEventArgs With {.myKey = dr.Item("vendor")})
    End Sub

    Private Sub CreateEmailEWS(ByVal obj As Object, p2 As Integer)
        Dim myEmail As New EmailEWS(myForm)
        'myEmail.urlEWS = myParam.GeturlEWS("urlEWS")

        Dim dr As DataRow = DirectCast(obj, DataRow)
        Dim VendorInfo As DataRow = myParam.GetVendorInfo(dr.Item("vendor"))

        'myEmail.ToRecepients = "etai@groupeseb.com;dlie@groupeseb.com" 'VendorInfo.Item("email")

        myEmail.ToRecepients = VendorInfo.Item("email")
        If Not IsDBNull(VendorInfo.Item("spemail")) Then
            myEmail.CcRecepients = VendorInfo.Item("spemail")
        End If

        'If VendorInfo.Item("copytosp") Then
        '    myEmail.CcRecepients = VendorInfo.Item("spemail")
        'End If

        myEmail.Username = EWSUser(0)
        myEmail.password = EWSUser(1)
        myEmail.domain = "AS"
        'myEmail.Subject = String.Format(String.Format("Inspection Schedule : {0:dd-MMM-yyyy} - {1:dd-MMM-yyyy}", myForm.myController.lastsenddate, myForm.myController.enddate))
        myEmail.Subject = String.Format(String.Format("Inspection Schedule : {0:dd-MMM-yyyy} ", myForm.myController.startdate))


        ' If Not IsDBNull(dr.Item("spemail")) Then myEmail.CcRecepients = dr.Item("spemail")

        myEmail.BodyMessage = GetBodyMessage(dr.Item("vendor"))

        myEmail.RunWithKey(myForm, New GenerateEmailWithKeyEventArgs With {.myKey = dr.Item("vendor")})


    End Sub

    Private Function GetBodyMessage(ByVal vendor As Long) As String
        Dim sb As New StringBuilder
        Dim myNotification As String = myParam.GetNotification(Date.Today)

        'Dim myModel As New SendEmailModel

        'Dim myModel As New SendEmailModel(Inspectiondate)

        Dim bs As New BindingSource
        'myModel.GetContent(bs, vendor, startdate, enddate)
        myModel.GetContent(bs, vendor)



        'sb.Append("<!DOCTYPE html><html><head><meta name='description' content='[QualitySystem]'/><meta http-equiv='Content-Type' content='text/html; charset=us-ascii'></head><style>td,th {padding-left:5px;         padding-right:10px; text-align:left;  }  th {background-color:red;    color:white} .defaultfont(){font-size:11.0pt; font-family:'Calibri','sans-serif';}</style><body class='defaultfont'>")
        sb.Append("<!DOCTYPE html><html><head><meta name='description' content='[QualitySystem]'/><meta http-equiv='Content-Type' content='text/html; charset=us-ascii'><style>td,th {padding-left:5px;         padding-right:10px; text-align:left;  }  th {background-color:red;    color:white} body{font-size:11.0pt; font-family:'Calibri','sans-serif';}</style></head><body>")
        If myNotification.Length > 0 Then
            sb.Append(myNotification)
        End If

        If vendor = 0 Then
            sb.Append("<p>Dear Colleagues,</p><p>Please find inspection schedule for your reference.</p>")
        Else
            sb.Append("<p>Dear Vendor,</p><p>Please find inspection schedule for your reference.</p>")
        End If


        sb.Append("<table border=1 cellspacing=0 class='defaultfont'> ")
        sb.Append("<tr><th>Inspection Booking</th><th>Inspector</th><th>Insp.Lot.</th><th>Sample Size</th><th>Purch.Doc.</th><th>Item</th><th>Purchase Order Number</th><th>Supplier Name</th><th>Material</th><th>Material Description</th><th>SBU Description</th><th>Confirmed ETD</th><th>Quantity</th><th>City</th><th>Sold To Party Name</th><th>SP</th></tr>")
        For Each drv In bs.List
            sb.Append(String.Format("<tr><td>{11:dd-MMM-yyyy}</td><td>{12}</td><td>{0}</td><td>{1:#,##0}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9:dd-MMM-yyyy}</td><td>{10:#,##0}</td><td>{13}</td><td>{14}</td><td>{15}</td></tr>", drv.row.item("insplot"), drv.row.item("samplesize"), drv.row.item("purchdoc"), drv.row.item("item"), drv.row.item("custpono"), drv.row.item("vendorname"), drv.row.item("material"), drv.row.item("materialdesc"), drv.row.item("sbu"), drv.row.item("ccetd"), drv.row.item("qty"), drv.row.item("startdate"), drv.row.item("inspector"), drv.row.item("city"), drv.row.item("soldtopartyname"), drv.row.item("sp")))
        Next
        sb.Append("</table> ")        
        sb.Append("<p> Best Regards,<br>WONG Tracy</br></html>")
        Return sb.ToString
    End Function

    Private Function SendEmailConfirmationVendorManually(DataTable As DataTable, InspectionDate As Date) As Boolean
        Dim myret As Boolean = False
        Dim myquery = From n In DataTable.AsEnumerable
                  Where n.Item("Selected")
                  Group By key = n.Item("Vendor") Into Group
        For Each a In myquery            
            Dim VendorInfo As DataRow = myParam.GetVendorInfo(CLng(a.key))
            myForm.ProgressReport(1, String.Format("Create Draft for {0}", VendorInfo.Item("vendorname")))
            Dim myEmail As New EmailEWS(myForm)
            myEmail.urlEWS = myParam.GeturlEWS("urlEWS")
            myEmail.ToRecepients = VendorInfo.Item("email") '"etai@groupeseb.com;dlie@groupeseb.com" '"dlie@groupeseb.com"         
            'If VendorInfo.Item("copytosp") Then
            '    myEmail.CcRecepients = VendorInfo.Item("spemail")
            'End If
            If Not IsDBNull(VendorInfo.Item("spemail")) Then
                myEmail.CcRecepients = VendorInfo.Item("spemail")
            End If
            myEmail.Username = EWSUser(0)
            myEmail.password = EWSUser(1)
            myEmail.domain = "AS"
            myEmail.Subject = String.Format(String.Format("Inspection Schedule - {0:dd-MMM-yyyy}", InspectionDate))


            Dim sb As New StringBuilder
            Dim myNotification As String = myParam.GetNotification(Date.Today)

            'sb.Append("<!DOCTYPE html><html><head><meta name='description' content='[QualitySystem]'/><meta http-equiv='Content-Type' content='text/html; charset=us-ascii'></head><style>td,th {padding-left:5px;         padding-right:10px; text-align:left;  }  th {background-color:red;    color:white} .defaultfont(){font-size:11.0pt; font-family:'Calibri','sans-serif';}</style><body class='defaultfont'>")
            sb.Append("<!DOCTYPE html><html><head><meta name='description' content='[QualitySystem]'/><meta http-equiv='Content-Type' content='text/html; charset=us-ascii'><style>td,th {padding-left:5px;         padding-right:10px; text-align:left;  }  th {background-color:red;    color:white} body{font-size:11.0pt; font-family:'Calibri','sans-serif';}</style></head><body>")
            If myNotification.Length > 0 Then
                sb.Append(myNotification)
            End If
            sb.Append("<p>Dear Vendor,</p><p>Please find inspection schedule for your reference.</p>")

            sb.Append("<table border=1 cellspacing=0 class='defaultfont'> ")
            sb.Append("<tr><th>Insp.Lot.</th><th>Sample Size</th><th>Purch.Doc.</th><th>Item</th><th>Purchase Order Number</th><th>Supplier Name</th><th>Material</th><th>Material Description</th><th>SBU Description</th><th>Confirmed ETD</th><th>Quantity</th><th>Inspection Booking</th><th>Inspector</th></tr>")
            For Each dr As DataRow In a.Group
                sb.Append(String.Format("<tr><td>{0}</td><td>{1:#,##0}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9:dd-MMM-yyyy}</td><td>{10:#,##0}</td><td>{11:dd-MMM-yyyy}</td><td>{12}</td></tr>", dr.Item("Insp. Lot"), dr.Item("Sample size"), dr.Item("Purch.Doc."), dr.Item("Item"), dr.Item("Cust PO No"), dr.Item("Vendor Name"), dr.Item("Material"), dr.Item("Material desc"), dr.Item("SBU"), dr.Item("Confirmed ETD"), dr.Item("Quantity"), dr.Item("startdate"), dr.Item("inspector")))
            Next
            sb.Append("</table> ")
            sb.Append("<p> Best Regards,<br>WONG Tracy</br></html>")

            myEmail.BodyMessage = sb.ToString
            If Not myEmail.createemailSave() Then
                Return myret
            End If
        Next
        myret = True
        Return myret
    End Function

    Private Function SendEmailConfirmationInternalManually(DataTable As DataTable, InspectionDate As Date) As Boolean
        Dim myret As Boolean = False
        Dim myqueryInternal = From n In DataTable.AsEnumerable
                 Where n.Item("Selected")

        myForm.ProgressReport(1, String.Format("Create Draft for {0}", "Internal"))
        Dim myEmail = New EmailEWS(myForm)
        myEmail.urlEWS = myParam.GeturlEWS("urlEWS")
        myEmail.ToRecepients = myParam.GetInternalEmail("Internal Email (to)") '"etai@groupeseb.com;dlie@groupeseb.com" '"dlie@groupeseb.com"            
        myEmail.CcRecepients = myParam.GetInternalEmail("Internal Email (cc)")
        myEmail.Username = EWSUser(0)
        myEmail.password = EWSUser(1)
        myEmail.domain = "AS"
        myEmail.Subject = String.Format(String.Format("Inspection Schedule - {0:dd-MMM-yyyy}", InspectionDate))


        Dim sbInternal As New StringBuilder
        Dim myNotificationInternal As String = myParam.GetNotification(Date.Today)

        'sbInternal.Append("<!DOCTYPE html><html><head><meta name='description' content='[QualitySystem]'/><meta http-equiv='Content-Type' content='text/html; charset=us-ascii'></head><style>td,th {padding-left:5px;         padding-right:10px; text-align:left;  }  th {background-color:red;    color:white} .defaultfont(){font-size:11.0pt; font-family:'Calibri','sans-serif';}</style><body class='defaultfont'>")
        sbInternal.Append("<!DOCTYPE html><html><head><meta name='description' content='[QualitySystem]'/><meta http-equiv='Content-Type' content='text/html; charset=us-ascii'><style>td,th {padding-left:5px;         padding-right:10px; text-align:left;  }  th {background-color:red;    color:white} body{font-size:11.0pt; font-family:'Calibri','sans-serif';}</style></head><body>")
        If myNotificationInternal.Length > 0 Then
            sbInternal.Append(myNotificationInternal)
        End If
        sbInternal.Append("<p>Dear Colleagues,</p><p>Please find inspection schedule for your reference.</p>")

        sbInternal.Append("<table border=1 cellspacing=0 class='defaultfont'> ")
        sbInternal.Append("<tr><th>Insp.Lot.</th><th>Sample Size</th><th>Purch.Doc.</th><th>Item</th><th>Purchase Order Number</th><th>Supplier Name</th><th>Material</th><th>Material Description</th><th>SBU Description</th><th>Confirmed ETD</th><th>Quantity</th><th>Inspection Booking</th><th>Inspector</th><th>City</th><th>Sold To Party Name</th><th>SP</th></tr>")
        For Each dr In myqueryInternal
            'sbInternal.Append(String.Format("<tr><td>{0}</td><td>{1:#,##0}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9:dd-MMM-yyyy}</td><td>{10:#,##0}</td><td>{11:dd-MMM-yyyy}</td><td>{12}</td></tr>", dr.Item("Insp. Lot"), dr.Item("Sample size"), dr.Item("Purch.Doc."), dr.Item("Item"), dr.Item("Cust PO No"), dr.Item("Vendor Name"), dr.Item("Material"), dr.Item("Material desc"), dr.Item("SBU"), dr.Item("Confirmed ETD"), dr.Item("Quantity"), dr.Item("startdate"), dr.Item("inspector")))
            sbInternal.Append(String.Format("<tr><td>{0}</td><td>{1:#,##0}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9:dd-MMM-yyyy}</td><td>{10:#,##0}</td><td>{11:dd-MMM-yyyy}</td><td>{12}</td><td>{13}</td><td>{14}</td><td>{15}</td></tr>", dr.Item("Insp. Lot"), dr.Item("Sample size"), dr.Item("Purch.Doc."), dr.Item("Item"), dr.Item("Cust PO No"), dr.Item("Vendor Name"), dr.Item("Material"), dr.Item("Material desc"), dr.Item("SBU"), dr.Item("Confirmed ETD"), dr.Item("Quantity"), InspectionDate, dr.Item("inspector"), dr.Item("City"), dr.Item("Sold To Party Name"), dr.Item("sp")))
        Next
        sbInternal.Append("</table> ")
        sbInternal.Append("<p> Best Regards,<br>WONG Tracy</br></html>")

        myEmail.BodyMessage = sbInternal.ToString

        'Create attachment

        Dim filename = String.Format("{0}\{1}", DirectoryName, "InternalAll.xlsx")

        Dim datasheet As Integer = 1

        Dim mycallback As FormatReportDelegate = AddressOf FormattingReportManually
        Dim PivotCallback As FormatReportDelegate = AddressOf PivotTableManually
        Dim myTable = DataTable
        Dim myreport As New ExportToExcelFile(myForm, filename, "\templates\ExcelTemplate01.xltx", myTable, mycallback, PivotCallback, False)
        myreport.ExtractFromDataTableUnsyncDT(myForm, New System.EventArgs)

        Try

            myEmail.AttachmentList.Add(String.Format("{0}", filename))
            'Kill(String.Format("{0}\{1}", DirectoryName, ReportName))
        Catch ex As Exception
        End Try

        If Not myEmail.createemailSave() Then
            Return myret
        End If
        myret = True
        Return myret
    End Function

    Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        Dim osheet As Excel.Worksheet = DirectCast(sender, Excel.Worksheet)
        osheet.Columns("A:E").Delete()
        osheet.Columns("AB:AB").Delete()
    End Sub

    Sub MyPivotTable(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
    End Sub

    Private Sub FormattingReportManually(ByRef sender As Object, ByRef e As EventArgs)
        Dim osheet As Excel.Worksheet = DirectCast(sender, Excel.Worksheet)
        osheet.Columns("A:A").Delete()
    End Sub

    Private Sub PivotTableManually(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
    End Sub




End Class
