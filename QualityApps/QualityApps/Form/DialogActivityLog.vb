Imports System.Windows.Forms

Public Class DialogActivityLog
    Private drv As DataRowView

    Private TimeSessions As New List(Of TimeSession)
    Private VendorBS As BindingSource
    Private ActivityBS As BindingSource
    Private TimeSessionBS As BindingSource

    Public Event RefreshInterface()

    Public Sub New(DRV As DataRowView, vendorbs As BindingSource, activitybs As BindingSource, timesessionbs As BindingSource)
        InitializeComponent()

        Me.drv = DRV
        Me.drv.BeginEdit()
        Me.VendorBS = vendorbs
        Me.ActivityBS = activitybs
        'Me.TimeSessionBS = timesessionbs
        InitData()
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        drv.EndEdit()
        RaiseEvent RefreshInterface()
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        drv.CancelEdit()
        Me.Close()
    End Sub

    Private Sub InitData()
        TimeSessions.Add(New TimeSession("Full Day", 1))
        TimeSessions.Add(New TimeSession("Half Day", 0.5))
        TimeSessions.Add(New TimeSession("Over Time", 2))

        TimeSessionBS = New BindingSource
        TimeSessionBS.DataSource = TimeSessions

        ComboBox1.DataSource = TimeSessionBS
        ComboBox1.DisplayMember = "description"
        ComboBox1.ValueMember = "myvalue"


        ComboBox2.DisplayMember = "vendorcodename"
        ComboBox2.ValueMember = "vendorcode"
        ComboBox2.DataSource = VendorBS

        ComboBox3.DisplayMember = "activityname"
        ComboBox3.ValueMember = "id"
        ComboBox3.DataSource = ActivityBS

        ComboBox1.DataBindings.Clear()
        ComboBox2.DataBindings.Clear()
        ComboBox3.DataBindings.Clear()
        DateTimePicker1.DataBindings.Clear()
        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()

        ComboBox1.DataBindings.Add(New Binding("SelectedValue", drv, "timesession", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox2.DataBindings.Add(New Binding("SelectedValue", drv, "vendorcode", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox3.DataBindings.Add(New Binding("SelectedValue", drv, "activityid", True, DataSourceUpdateMode.OnPropertyChanged))
        DateTimePicker1.DataBindings.Add(New Binding("Text", drv, "activitydate", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox1.DataBindings.Add(New Binding("Text", drv, "projectname", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox2.DataBindings.Add(New Binding("Text", drv, "remark", True, DataSourceUpdateMode.OnPropertyChanged))

    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        drv.Item("timesessiondesc") = ComboBox1.SelectedItem
        RaiseEvent RefreshInterface()
    End Sub

    Private Sub ComboBox2_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox2.SelectionChangeCommitted
        Dim cbdrv As DataRowView = ComboBox2.SelectedItem
        drv.Item("vendorcodename") = cbdrv.Row.Item("vendorcodename")
        RaiseEvent RefreshInterface()
    End Sub

    Private Sub ComboBox3_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox3.SelectionChangeCommitted
        Dim cbdrv As DataRowView = ComboBox3.SelectedItem
        drv.Item("activityname") = cbdrv.Row.Item("activityname")
        RaiseEvent RefreshInterface()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged
        RaiseEvent RefreshInterface()
    End Sub
End Class


Public Class TimeSession
    Public Property Description As String
    Public Property myvalue As Decimal

    Public Sub New(Description As String, myvalue As Decimal)
        Me.Description = Description
        Me.myvalue = myvalue
    End Sub

    Public Overrides Function ToString() As String
        Return Description
    End Function

End Class