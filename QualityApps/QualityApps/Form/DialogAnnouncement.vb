Imports System.Windows.Forms

Public Class DialogAnnouncement
    Private drv As DataRowView
    Public Event RefreshDataGridView()

    Public Sub New(ByVal drv As DataRowView)
        InitializeComponent()
        Me.drv = drv
        BindingControl()
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        drv.EndEdit()
        RaiseEvent RefreshDataGridView()
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        drv.CancelEdit()
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub BindingControl()
        DateTimePicker1.DataBindings.Clear()
        DateTimePicker2.DataBindings.Clear()
        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()

        TextBox1.DataBindings.Add(New Binding("Text", drv, "title", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox2.DataBindings.Add(New Binding("Text", drv, "bodymessage", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        DateTimePicker1.DataBindings.Add(New Binding("Text", drv, "startdate", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        DateTimePicker2.DataBindings.Add(New Binding("Text", drv, "enddate", True, DataSourceUpdateMode.OnPropertyChanged, ""))

    End Sub

End Class
