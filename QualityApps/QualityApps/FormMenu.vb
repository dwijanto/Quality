Imports System.Reflection
Public Enum TxEnum
    NewRecord = 1
    CopyRecord = 2
    UpdateRecord = 3
End Enum
Public Class FormMenu

    Private userinfo1 As UserInfo
    Private dbAdapter1 As PostgreSQLDBAdapter
    Dim HasError As Boolean = True
    Private userid As String
    Private myuser As UserController
    Public Sub New()
        myuser = New UserController
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

        Try
            userinfo1 = UserInfo.getInstance
            userinfo1.Userid = Environment.UserDomainName & "\" & Environment.UserName
            userinfo1.computerName = My.Computer.Name
            userinfo1.ApplicationName = "Quality Team Apps"
            userinfo1.Username = "N/A"
            userinfo1.isAuthenticate = False
            userinfo1.Role = 0

            dbAdapter1 = PostgreSQLDBAdapter.getInstance
            dbAdapter1.UserInfo = userinfo1
            dbAdapter1.UserInfo.isAdmin = dbAdapter1.IsAdmin(userinfo1.Userid)

            HasError = False
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Private Sub FormMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If HasError Then
            Me.Close()
            Exit Sub
        End If

        Try
            userid = userinfo1.Userid 'Environment.UserDomainName & "\" & Environment.UserName

            Dim myAD = New ADPrincipalContext
            Dim UserInfo As List(Of ADPrincipalContext) = New List(Of ADPrincipalContext)
            If myAD.GetInfo(userid) Then        
                myuser.Model.ADDUPDUserManager(ADPrincipalContext.ADPrincipalContexts)
            Else
                MessageBox.Show(myAD.ErrorMessage)
                Me.Close()
                Exit Sub
            End If


            Dim mydata As DataSet = myuser.findByUserid(userid.ToLower)
            If mydata.tables(0).rows.count > 0 Then
                Dim identity = myuser.findIdentity(mydata.Tables(0).rows(0).item("id"))
                User.setIdentity(identity)
                User.login(identity)
                User.IdentityClass = myuser
                dbAdapter1 = PostgreSQLDBAdapter.getInstance
                loglogin(userinfo1.Userid)
                Me.Text = GetMenuDesc()
                Me.Location = New Point(300, 10)
                MenuHandles()
            Else
                'disable menubar
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Me.Close()
        End Try

    End Sub
    'Private Sub FormMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    '    If HasError Then
    '        Me.Close()
    '        Exit Sub
    '    End If

    '    Try
    '        userid = Environment.UserDomainName & "\" & Environment.UserName
    '        Dim mydata As DataSet = myuser.findByUserid(userid.ToLower)
    '        If mydata.tables(0).rows.count > 0 Then
    '            Dim identity = myuser.findIdentity(mydata.Tables(0).rows(0).item("id"))
    '            User.setIdentity(identity)
    '            User.login(identity)
    '            User.IdentityClass = myuser
    '            dbAdapter1 = PostgreSQLDBAdapter.getInstance
    '            loglogin(userinfo1.Userid)
    '            Me.Text = GetMenuDesc()
    '            Me.Location = New Point(300, 10)
    '            MenuHandles()
    '        Else
    '            'disable menubar
    '        End If

    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '        Me.Close()
    '    End Try

    'End Sub

    Private Sub loglogin(ByVal userid As String)
        Dim applicationname As String = "Quality Team Apps"
        Dim username As String = Environment.UserDomainName & "\" & Environment.UserName
        Dim computername As String = My.Computer.Name
        Dim time_stamp As DateTime = Now
        dbAdapter1.loglogin(applicationname, userid, username, computername, time_stamp)
    End Sub

    Public Function GetMenuDesc() As String
        Return "App.Version: " & My.Application.Info.Version.ToString & " :: Server: " & dbAdapter1.HOST & ", Database: " & dbAdapter1.Database & ", Userid: " & dbAdapter1.UserInfo.Userid 'HelperClass1.UserId
    End Function

    Private Sub MenuHandles()
        AddHandler ImportSQ01ToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler GenerateExcelForSupplierToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler GenerateExcelForSupplierToolStripMenuItem1.Click, AddressOf ToolStripMenuItem_Click
        AddHandler CollectReplyFromSupplierToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler InspectionReportToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler VendorToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler AnnouncementToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        'AddHandler Send2DaysEmailInAdvanceToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ParameterToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler UserToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler CreateActivityLogToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportDataToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler GenerateCSVToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ReportActivityLogToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ActivityToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        'Admin Function
        'MasterToolStripMenuItem.Visible = userinfo1.isAdmin
        AddHandler VendorAssignmentQEUserToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click


        Dim identity As UserController = User.getIdentity
        TransactionToolStripMenuItem.Visible = User.can("View Actions")
        ReportsToolStripMenuItem.Visible = User.can("View Scheduling")
        MasterToolStripMenuItem.Visible = User.can("View Master")
        AdminToolStripMenuItem.Visible = User.can("View Admin")
        ActivityLogToolStripMenuItem.Visible = User.can("View Activity Log")
        CreateActivityLogToolStripMenuItem.Visible = User.can("View Create Activity Log")
        ReportActivityLogToolStripMenuItem.Visible = User.can("View Report Activity Log")
        OTMToolStripMenuItem.Visible = User.can("View OTM")
        ActivityToolStripMenuItem.Visible = User.can("View Master")


    End Sub
    Private Sub ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ctrl As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        Dim assembly1 As Assembly = Assembly.GetAssembly(GetType(FormMenu))
        Dim frm As Object = CType(assembly1.CreateInstance(assembly1.GetName.Name.ToString & "." & ctrl.Tag.ToString, True), Form)
        Dim myform = frm.GetInstance
        myform.show()
        myform.windowstate = FormWindowState.Normal
        myform.activate()
    End Sub

  

   
    Private Sub Send2DaysEmailInAdvanceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Send2DaysEmailInAdvanceToolStripMenuItem.Click
        Dim AskDate As New DialogAskDate
        If AskDate.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim myform = New FormSendEmailConfirmation
            myform.Inspectiondate = AskDate.DTInscpectionDate.Value
            myform.Text = String.Format("{0} - Inspection Date {1:dd-MMM-yyyy}", myform.Text, myform.Inspectiondate)
            myform.ShowDialog()
        End If
        
    End Sub

    Private Sub RBACToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RBACToolStripMenuItem.Click
        Dim myform = New FormRBAC
        myform.ShowDialog()
    End Sub

    Private Sub TransactionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TransactionToolStripMenuItem.Click

    End Sub

    Private Sub ReportsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReportsToolStripMenuItem.Click

    End Sub

    Private Sub MasterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MasterToolStripMenuItem.Click

    End Sub

    Private Sub AdminToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AdminToolStripMenuItem.Click

    End Sub

    Private Sub GenerateExcelForSupplierToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles GenerateExcelForSupplierToolStripMenuItem1.Click

    End Sub

    Private Sub UserToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UserToolStripMenuItem.Click

    End Sub

    Private Sub ActivityLogToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ActivityLogToolStripMenuItem.Click

    End Sub

    Private Sub CreateActivityLogToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateActivityLogToolStripMenuItem.Click

    End Sub

    Private Sub ReportActivityLogToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReportActivityLogToolStripMenuItem.Click

    End Sub

    Private Sub ImportDataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportDataToolStripMenuItem.Click

    End Sub

    Private Sub OTMToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OTMToolStripMenuItem.Click

    End Sub

    Private Sub GenerateCSVToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GenerateCSVToolStripMenuItem.Click

    End Sub

    Private Sub ActivityToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ActivityToolStripMenuItem.Click

    End Sub

    Private Sub VendorAssignmentQEUserToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VendorAssignmentQEUserToolStripMenuItem.Click

    End Sub
End Class
