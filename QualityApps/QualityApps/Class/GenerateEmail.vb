Imports System.Threading
Imports System.Text
Imports Microsoft.Exchange.WebServices.Data

Public Class GenerateEmail

    Private myForm As Object
    Private DS As DataSet
    Dim myCallback As System.Net.Mail.SendCompletedEventHandler

    Dim myParam As ParamAdapter = ParamAdapter.getInstance
    Dim EWSUser As String() = myParam.GetEWSUser("EWS")
    Dim OUTBOXFolder As String = myParam.GetMailFolder("OUTBOX")

    Dim mydate As Date = Date.Today
    Dim mydate2 As Date = Date.Today

    Public Sub New(ByVal parent As Object)
        Me.myForm = parent
    End Sub

    Public Sub New(ByVal parent As Object, ByVal mycallback As System.Net.Mail.SendCompletedEventHandler)
        Me.myForm = parent
        Me.myCallback = mycallback
    End Sub

    Public Sub New(ByVal parent As Object, ByVal mycallback As System.Net.Mail.SendCompletedEventHandler, ByVal mydate As Date)
        Me.myForm = parent
        Me.myCallback = mycallback
        Me.mydate = mydate
    End Sub

    Public Sub New(ByVal parent As Object, ByVal mycallback As System.Net.Mail.SendCompletedEventHandler, ByVal mydate As Date, ByVal mydate2 As Date)
        Me.myForm = parent
        Me.myCallback = mycallback
        Me.mydate = mydate
        Me.mydate2 = mydate2
    End Sub

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
                Thread.Sleep(1000) 'slowdown the process
                'CreateEmailEWS(ds.Tables(0).Rows(i).Item("vendor"), i + 1)
                CreateEmailEWS(ds.Tables(0).Rows(i), i + 1)
            End If
        Next

        Return myret
    End Function

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
        myForm.ProgressReport(2004, String.Format("{0}", mycount))

        For i = 0 To ds.Tables(0).Rows.Count - 1
            If ds.Tables(0).Rows(i).Item("selected") Then
                Thread.Sleep(100) 'slowdown the process
                CreateEmail(ds.Tables(0).Rows(i).Item("vendor"), i + 1)
            End If
        Next

        Return myret
    End Function

    Private Sub CreateEmail(ByVal obj As Object, p2 As Integer)
        Dim myEmail As New Email(myForm)
        myEmail.sendto = "dlie@groupeseb.com"
        myEmail.sender = "dwijanto@yahoo.com.com"
        myEmail.subject = String.Format("Test for {0}", obj)
        myEmail.body = "hello"
        myEmail.RunWithKey(myForm, New GenerateEmailWithKeyEventArgs With {.myKey = obj})
        'myEmail.send(obj.ToString)

    End Sub

    Private Sub CreateEmailEWS(ByVal obj As Object, p2 As Integer)
        Dim myEmail As New EmailEWS(myForm)
        myEmail.urlEWS = myParam.GeturlEWS("urlEWS")

        Dim dr As DataRow = DirectCast(obj, DataRow)
        Dim VendorInfo As DataRow = myParam.GetVendorInfo(dr.Item("vendor"))

        'Dim filename As String = String.Format("{0:yyyyMMdd}-{1}.xlsx", Date.Today, dr.Item("vendor"))
        Dim filename As String = String.Format("{0:yyyyMMdd}-{1}.xlsx", mydate.Date, dr.Item("vendor"))
        'myEmail.ToRecepients = "etai@groupeseb.com;dlie@groupeseb.com" 'VendorInfo.Item("email")
        myEmail.ToRecepients = VendorInfo.Item("email")
        If Not IsDBNull(VendorInfo.Item("spemail")) Then
            myEmail.CcRecepients = VendorInfo.Item("spemail")
        End If
        myEmail.Username = EWSUser(0)
        myEmail.password = EWSUser(1)
        myEmail.domain = "AS"
        myEmail.Subject = String.Format("Inspection Booking - {0}", filename)


        'If Not IsDBNull(dr.Item("spemail")) Then myEmail.CcRecepients = dr.Item("spemail")

        myEmail.BodyMessage = GetBodyMessage(dr.Item("maxdate"))
        Dim AttachmentList As New List(Of String)
        AttachmentList.Add(String.Format("{0}\{1}", OUTBOXFolder, filename))
        myEmail.AttachmentList = AttachmentList               
        myEmail.RunWithKey(myForm, New GenerateEmailWithKeyEventArgs With {.myKey = dr.Item("vendor")})


    End Sub

    Private Function nextday(mydate As Date) As Date
        Dim addition As Integer = 1
        If mydate.DayOfWeek = 5 Then
            addition = 3
        End If
        Return mydate.AddDays(addition)
    End Function

    Private Function GetBodyMessage(maxdate As Date) As String
        Dim sb As New StringBuilder
        Dim myNotification As String = myParam.GetNotification(Date.Today)

        'sb.Append("<!DOCTYPE html><html><head><meta name='description' content='[QualitySystem]'/><meta http-equiv='Content-Type' content='text/html; charset=us-ascii'><style>body{font-size:11.0pt; font-family:'Calibri','sans-serif';}</style></head><body>")
        'sb.Append("<!DOCTYPE html><html><head><meta name='description' content='[QualitySystem]'/><meta http-equiv='Content-Type' content='text/html; charset=us-ascii'></head><style>.defaultfont(){font-size:11.0pt; font-family:'Calibri','sans-serif';}</style><body class='defaultfont'>")
        sb.Append("<!DOCTYPE html><html><head><meta name='description' content='[QualitySystem]'/><meta http-equiv='Content-Type' content='text/html; charset=us-ascii'><style>td,th {padding-left:5px;         padding-right:10px; text-align:left;  }  th {background-color:red;    color:white} body{font-size:11.0pt; font-family:'Calibri','sans-serif';}</style></head><body>")
        If myNotification.Length > 0 Then
            sb.Append(myNotification)
        End If
        ' class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-list:l0 level1 lfo1'
        'sb.Append("<!DOCTYPE html><html><head><meta name='description' content='[QualitySystem]'/><meta http-equiv='Content-Type' content='text/html; charset=us-ascii'></head><style>.defaultfont(){font-size:11.0pt; font-family:'Calibri','sans-serif';}</style><body class='defaultfont'><p>Dear Vendor,</p><p>Enclosed with the CCETD 2-4 weeks PO list for your review. Please check and provide the <u>inspection date</u> for the PO(s) CCETD +14 days.</p><p><span style='color:red'><b>**Important Notice**</b></span></p>" &
        '          "<ol start=1 type=1>" &
        '          "<span style='color:red'>" &
        '          "<li>Vendor is required to provide inspection date and <b><u>MUST</u></b> put it in the &#8220;Inspection Date&#8221; Column, <b><u>DO NOT</u></b> put inspection date in &#8220;Supplier Remark&#8221; column, misplacement will be considered as not an official booking.</li>" &
        '          "<li>Please also be aware that the inspection date should avoid Saturday and Sunday except request for OT. </li>" &
        '          "<li>&#8220;Supplier Remark&#8221; column is for supporting information only, ie. different inspection location, address and contact person, etc.</li>" &
        '          "</span></ol>" &
        '          "<p>Vendor must return this file to SEB Asia Quality Inspection booking team before <b><u><span style='background:yellow'>")
        sb.Append("<p>Dear Vendor,</p><p>Enclosed with the CCETD 2-4 weeks PO list for your review. Please check and provide the <u>inspection date</u> for the PO(s) CCETD +14 days.</p><span style='color:red'><p><b>**Important Notice**</b></p>" &
                  "<ol start=1 type=1>" &
                  "<li>Vendor is required to provide inspection date and <b><u>MUST</u></b> put it in the &#8220;New Input Inspection Date&#8221; Column, <b><u>DO NOT</u></b> put inspection date in &#8220;Supplier Remark&#8221; column, misplacement will be considered as not an official booking.</li>" &
                  "<li>Please also be aware that the inspection date should avoid Saturday and Sunday except request for OT. </li>" &
                  "<li>&#8220;Supplier Remark&#8221; column is for supporting information only, ie. different inspection location, address and contact person, etc.</li>" &
                  "</ol></span>" &
                  "<p>Vendor must return this file to SEB Asia Quality Inspection booking team before <b><u><span style='background:yellow'>")

        'sb.Append(String.Format("{0:dd-MMM-yyyy} ", nextday(mydate2)))
        sb.Append(String.Format("{0:dd-MMM-yyyy} ", mydate2))
        sb.Append("10:30am </b></span></u><span style='background:yellow'> </span></b>.<span style='color:#365F91'> </span>All late incoming inspection booking will be disregarded, and the vendor shall wait for the next day&#8217;s new schedule from SEB Asia.</p><p>This file solely is for Quality inspection allocation purpose only. Any changes of the PO information, vendor must update it from <u><span style='color:#0033CC'><a href='https://asiavendor.groupeseb.com/SCM/login.do'>VENDOR PORTAL</a></span></u></p><span style='color:#746661'><b><p>Best Regards,</p><p>SEB Asia Inspection Booking Team</p></b></span></body></html>")
        Return sb.ToString
    End Function

    Private Function GetBodyMessage1(maxdate As Date) As String
        Dim sb As New StringBuilder
        Dim myNotification As String = myParam.GetNotification(Date.Today)

        'sb.Append("<!DOCTYPE html><html><head><meta name='description' content='[QualitySystem]'/><meta http-equiv='Content-Type' content='text/html; charset=us-ascii'><style>body{font-size:11.0pt; font-family:'Calibri','sans-serif';}</style></head><body>")
        'sb.Append("<!DOCTYPE html><html><head><meta name='description' content='[QualitySystem]'/><meta http-equiv='Content-Type' content='text/html; charset=us-ascii'></head><style>.defaultfont(){font-size:11.0pt; font-family:'Calibri','sans-serif';}</style><body class='defaultfont'>")
        sb.Append("<!DOCTYPE html><html><head><meta name='description' content='[QualitySystem]'/><meta http-equiv='Content-Type' content='text/html; charset=us-ascii'><style>td,th {padding-left:5px;         padding-right:10px; text-align:left;  }  th {background-color:red;    color:white} body{font-size:11.0pt; font-family:'Calibri','sans-serif';}</style></head><body>")
        If myNotification.Length > 0 Then
            sb.Append(myNotification)
        End If
        ' class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-list:l0 level1 lfo1'
        'sb.Append("<!DOCTYPE html><html><head><meta name='description' content='[QualitySystem]'/><meta http-equiv='Content-Type' content='text/html; charset=us-ascii'></head><style>.defaultfont(){font-size:11.0pt; font-family:'Calibri','sans-serif';}</style><body class='defaultfont'><p>Dear Vendor,</p><p>Enclosed with the CCETD 2-4 weeks PO list for your review. Please check and provide the <u>inspection date</u> for the PO(s) CCETD +14 days.</p><p><span style='color:red'><b>**Important Notice**</b></span></p>" &
        '          "<ol start=1 type=1>" &
        '          "<span style='color:red'>" &
        '          "<li>Vendor is required to provide inspection date and <b><u>MUST</u></b> put it in the &#8220;Inspection Date&#8221; Column, <b><u>DO NOT</u></b> put inspection date in &#8220;Supplier Remark&#8221; column, misplacement will be considered as not an official booking.</li>" &
        '          "<li>Please also be aware that the inspection date should avoid Saturday and Sunday except request for OT. </li>" &
        '          "<li>&#8220;Supplier Remark&#8221; column is for supporting information only, ie. different inspection location, address and contact person, etc.</li>" &
        '          "</span></ol>" &
        '          "<p>Vendor must return this file to SEB Asia Quality Inspection booking team before <b><u><span style='background:yellow'>")
        sb.Append("<p>Dear Vendor,</p><p>Enclosed with the CCETD 2-4 weeks PO list for your review. Please check and provide the <u>inspection date</u> for the PO(s) CCETD +14 days.</p><span style='color:red'><p><b>**Important Notice**</b></p>" &
                  "<ol start=1 type=1>" &
                  "<li>Vendor is required to provide inspection date and <b><u>MUST</u></b> put it in the &#8220;New Input Inspection Date&#8221; Column, <b><u>DO NOT</u></b> put inspection date in &#8220;Supplier Remark&#8221; column, misplacement will be considered as not an official booking.</li>" &
                  "<li>Please also be aware that the inspection date should avoid Saturday and Sunday except request for OT. </li>" &
                  "<li>&#8220;Supplier Remark&#8221; column is for supporting information only, ie. different inspection location, address and contact person, etc.</li>" &
                  "</ol></span>" &
                  "<p>Vendor must return this file to SEB Asia Quality Inspection booking team before <b><u><span style='background:yellow'>")

        sb.Append(String.Format("{0:dd-MMM-yyyy} ", nextday(Date.Today)))
        sb.Append("10:30am </b></span></u><span style='background:yellow'> </span></b>.<span style='color:#365F91'> </span>All late incoming inspection booking will be disregarded, and the vendor shall wait for the next day&#8217;s new schedule from SEB Asia.</p><p>This file solely is for Quality inspection allocation purpose only. Any changes of the PO information, vendor must update it from <u><span style='color:#0033CC'><a href='https://asiavendor.groupeseb.com/SCM/login.do'>VENDOR PORTAL</a></span></u></p><span style='color:#746661'><b><p>Best Regards,</p><p>SEB Asia Inspection Booking Team</p></b></span></body></html>")
        Return sb.ToString
    End Function

    Private Function GetBodyMessageold(maxdate As Date) As String
        Dim sb As New StringBuilder
        Dim myNotification As String = myParam.GetNotification(Date.Today)

        sb.Append("<!DOCTYPE html><html><head><meta name='description' content='[QualitySystem]'/><meta http-equiv='Content-Type' content='text/html; charset=us-ascii'></head><style>.defaultfont(){font-size:11.0pt; font-family:'Calibri','sans-serif';}</style><body class='defaultfont'>")
        If myNotification.Length > 0 Then
            sb.Append(myNotification)
        End If
        sb.Append("<p>Dear Vendor,</p><p>To avoid the possibility of missing order lines in this file and inform SEB ASIA timely the latest delivery schedule, please be reminded to review the confirmed ETD 14 days before confirmed ETD.If there is any change, please update the confirmed ETD in Vendor Portal timely.</p><p>Please note that the confirmed ETD to <b><span style='background:yellow'>")
        sb.Append(String.Format("{0:dd-MMM-yyyy} ", maxdate))
        sb.Append("</span></b> in this file.Thus, please review your confirmed ETD at <a href='https://asiavendor.groupeseb.com/SCM/login.do'>VENDOR PORTAL</a> if you cannot find the PO.</p><p>Please provide the <u>production end</u> and <u>closing date</u> for us to arrange the inspection schedule.<br><span style='color:red'>If you have more than one production plant for SEB products, please indicate your plant location at the Supplier Remarks.</span><br>Please reply us before <b><span style='background:yellow'>")
        sb.Append(String.Format("{0:dd-MMM-yyyy} ", Date.Today))
        sb.Append("12:30pm </b></span>.Thanks a lot.</p><p> For both of the below case, you need to enter in next day file.</p><ol start=1 type=1><li class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-list:l0 level1 lfo1'> Any reply after cut off time will not be accepted and will be treated as not replied.</li><li class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-list:l0 level1 lfo1'>Any reply with file format changed will not be accepted and will be treated as not replied.</li></ol><p> Best Regards,<br>WONG Tracy</br></html>")
        Return sb.ToString
    End Function



End Class
