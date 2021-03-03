Imports System.Threading
Public Class FormParameterCP
    Private Shared myform As FormParameterCP

    Public Shared Function getInstance()
        If myform Is Nothing Then
            myform = New FormParameterCP
        ElseIf myform.IsDisposed Then
            myform = New FormParameterCP
        End If
        Return myform
    End Function

    Dim myController As ParamAdapter
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Sub DoWork()
        myController = New ParamAdapter
        Try
            ProgressReport(1, "Loading...Please wait.")
            If myController.LoadDataCP() Then
                ProgressReport(4, "Init Data")
            End If
            ProgressReport(1, String.Format("Loading...Done."))
        Catch ex As Exception
            ProgressReport(1, ex.Message)
        End Try

    End Sub

    Private Sub LoadData()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub
    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    ToolStripStatusLabel1.Text = message
                Case 4

                    TextBox1.DataBindings.Clear()
                    TextBox2.DataBindings.Clear()
                    TextBox3.DataBindings.Clear()
                    TextBox4.DataBindings.Clear()
                    DateTimePicker1.DataBindings.Clear()

                    TextBox1.DataBindings.Add(New Binding("Text", myController.BS6, "ivalue", False, DataSourceUpdateMode.OnPropertyChanged))
                    TextBox2.DataBindings.Add(New Binding("Text", myController.BS7, "ivalue", False, DataSourceUpdateMode.OnPropertyChanged))
                    TextBox3.DataBindings.Add(New Binding("Text", myController.BS8, "ivalue", False, DataSourceUpdateMode.OnPropertyChanged))
                    TextBox4.DataBindings.Add(New Binding("Text", myController.BS10, "cvalue", False, DataSourceUpdateMode.OnPropertyChanged))
                    DateTimePicker1.DataBindings.Add(New Binding("Text", myController.BS9, "dvalue", False, DataSourceUpdateMode.OnPropertyChanged, "", "dd-MMM-yyyy"))


                    UcdgvParam1.DataGridView1.ContextMenuStrip = Nothing
                    UcdgvParam1.DataGridView1.Columns(0).DataPropertyName = "paramname"
                    UcdgvParam1.DataGridView1.Columns(1).DataPropertyName = "cvalue"
                    Dim ParentId As Long
                    ParentId = myController.GetParentid("MailFolderCP")
                    UcdgvParam1.BindingControl(Me, myController.BS, ParentId)


                    UcdgvParam2.DataGridView1.Columns(0).DataPropertyName = "paramname"
                    UcdgvParam2.DataGridView1.Columns(1).DataPropertyName = "cvalue"
                    ParentId = myController.GetParentid("SBU Exception CP")
                    UcdgvParam2.BindingControl(Me, myController.BS2, ParentId, "SBU Exception")

                    ParentId = myController.GetParentid("Vendor Exception CP")
                    UcdgvParam3.DataGridView1.Columns(0).DataPropertyName = "paramname"
                    UcdgvParam3.DataGridView1.Columns(1).DataPropertyName = "ivalue"
                    UcdgvParam3.DataGridView1.Columns(1).Width = 100
                    UcdgvParam3.DataGridView1.Columns(1).HeaderText = "Vendor Code"
                    UcdgvParam3.BindingControl(Me, myController.BS3, ParentId, "Vendor Exception")

                    ParentId = myController.GetParentid("Internal Email CP (to)")
                    UcdgvParam4.DataGridView1.Columns(0).HeaderText = "Name"
                    UcdgvParam4.DataGridView1.Columns(0).DataPropertyName = "paramname"
                    UcdgvParam4.DataGridView1.Columns(0).ReadOnly = False
                    UcdgvParam4.DataGridView1.Columns(1).DataPropertyName = "cvalue"
                    UcdgvParam4.DataGridView1.Columns(1).HeaderText = "Email"
                    UcdgvParam4.BindingControl(Me, myController.BS4, ParentId, "")

                    ParentId = myController.GetParentid("Internal Email CP (cc)")
                    UcdgvParam5.DataGridView1.Columns(0).HeaderText = "Name"
                    UcdgvParam5.DataGridView1.Columns(0).DataPropertyName = "paramname"
                    UcdgvParam5.DataGridView1.Columns(0).ReadOnly = False
                    UcdgvParam5.DataGridView1.Columns(1).DataPropertyName = "cvalue"
                    UcdgvParam5.DataGridView1.Columns(1).HeaderText = "Email"
                    UcdgvParam5.BindingControl(Me, myController.BS5, ParentId, "")

                    If Not User.can("ViewUserEWS") Then
                        TabControl1.TabPages.Remove(TabControl1.TabPages("UserEWS"))
                    End If

            End Select
        End If
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Me.Validate()
        myController.save()
    End Sub


    Private Sub FormParameters_Load(sender As Object, e As EventArgs) Handles Me.Load
        LoadData()

    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        LoadData()
    End Sub

End Class