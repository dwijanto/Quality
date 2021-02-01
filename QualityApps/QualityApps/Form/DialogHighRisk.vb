Imports System.Windows.Forms

Public Class DialogHighRisk

    Private drv As DataRowView
    Public Event RefreshDataGridView()
    Private CMMFBS As BindingSource
    Private VendorBS As BindingSource
    Private StatusBS As BindingSource
    Public Sub New(ByVal drv As DataRowView, ByVal cmmfbs As BindingSource, ByVal VendorBS As BindingSource, ByVal StatusBS As BindingSource)
        InitializeComponent()
        Me.drv = drv
        Me.CMMFBS = cmmfbs
        Me.VendorBS = VendorBS
        Me.StatusBS = StatusBS
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
        

        ComboBox1.ValueMember = "id"
        ComboBox1.DisplayMember = "statusname"
        ComboBox1.DataSource = StatusBS


        ComboBox1.DataBindings.Clear()

        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        TextBox4.DataBindings.Clear()
        TextBox5.DataBindings.Clear()
        CheckBox1.DataBindings.Clear()
        CheckBox2.DataBindings.Clear()
        CheckBox3.DataBindings.Clear()

        ComboBox1.DataBindings.Add(New Binding("selectedvalue", drv, "status", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox1.DataBindings.Add(New Binding("Text", drv, "cmmfdescription", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox2.DataBindings.Add(New Binding("Text", drv, "vendordescription", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox3.DataBindings.Add(New Binding("Text", drv, "sbu", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox4.DataBindings.Add(New Binding("Text", drv, "remarks", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox5.DataBindings.Add(New Binding("Text", drv, "qe", True, DataSourceUpdateMode.OnPropertyChanged))
        CheckBox1.DataBindings.Add(New Binding("checkstate", drv, "stresult", True, DataSourceUpdateMode.OnPropertyChanged))
        CheckBox2.DataBindings.Add(New Binding("checkstate", drv, "ndresult", True, DataSourceUpdateMode.OnPropertyChanged))
        CheckBox3.DataBindings.Add(New Binding("checkstate", drv, "rdresult", True, DataSourceUpdateMode.OnPropertyChanged))

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim myform = New FormHelper(CMMFBS)
        myform.DataGridView1.Columns(0).DataPropertyName = "description"

        If myform.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim hdrv As DataRowView = CMMFBS.Current
            drv.Row.Item("cmmf") = hdrv.Row.Item("cmmf")
            drv.Row.Item("materialdesc") = hdrv.Row.Item("materialdesc")
            TextBox1.Text = drv.Row.Item("materialdesc")
        End If
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim myform = New FormHelper(VendorBS)
        myform.DataGridView1.Columns(0).DataPropertyName = "description"

        If myform.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim hdrv As DataRowView = VendorBS.Current
            drv.Row.Item("vendorcode") = hdrv.Row.Item("vendorcode")
            drv.Row.Item("vendorname") = hdrv.Row.Item("vendorname")
            TextBox2.Text = drv.Row.Item("vendorname")
        End If
    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        Dim cbdrv As DataRowView = ComboBox1.SelectedItem
        drv.Row.Item("statusname") = cbdrv.Row.Item("statusname")
    End Sub

    Private Sub CheckBox1_CheckStateChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckStateChanged
        RaiseEvent RefreshDataGridView()

    End Sub
End Class
