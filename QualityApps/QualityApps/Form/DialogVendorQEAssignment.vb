Imports System.Windows.Forms

Public Class DialogVendorQEAssignment

    Private drv As DataRowView

    Private TimeSessions As New List(Of TimeSession)
    Private VendorBS As BindingSource
    Private UserBS As BindingSource
    Private QEBS As BindingSource
    Private VendorHelperBS As BindingSource
    Public Event RefreshInterface()

    Public Sub New(DRV As DataRowView, vendorbs As BindingSource, vendorhelperbs As BindingSource, UserBS As BindingSource, QEBS As BindingSource)
        InitializeComponent()

        Me.drv = DRV
        Me.drv.BeginEdit()
        Me.VendorBS = vendorbs
        Me.VendorHelperBS = vendorhelperbs
        Me.UserBS = UserBS
        Me.QEBS = QEBS

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


        ComboBox1.DisplayMember = "description"
        ComboBox1.ValueMember = "vendorcode"
        ComboBox1.DataSource = VendorBS

        'ComboBox2.DisplayMember = "username"
        'ComboBox2.ValueMember = "userid"
        'ComboBox2.DataSource = UserBS

        ComboBox3.DisplayMember = "qename"
        ComboBox3.ValueMember = "QEID"
        ComboBox3.DataSource = QEBS

        ComboBox1.DataBindings.Clear()
        'ComboBox2.DataBindings.Clear()
        ComboBox3.DataBindings.Clear()
       
        ComboBox1.DataBindings.Add(New Binding("SelectedValue", drv, "vendorcode", True, DataSourceUpdateMode.OnPropertyChanged))
        'ComboBox2.DataBindings.Add(New Binding("SelectedValue", drv, "userid", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox1.DataBindings.Add(New Binding("Text", drv, "userid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox3.DataBindings.Add(New Binding("SelectedValue", drv, "fgcp", True, DataSourceUpdateMode.OnPropertyChanged))
       

    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        selectionchangecommitted()
        
    End Sub

    'Private Sub ComboBox2_SelectionChangeCommitted(sender As Object, e As EventArgs)
    '    Dim userdrv As DataRowView = ComboBox2.SelectedItem
    '    drv.Item("username") = userdrv.Row.Item("username")
    '    RaiseEvent RefreshInterface()
    'End Sub

    Private Sub ComboBox3_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox3.SelectionChangeCommitted
        drv.Item("qe") = ComboBox3.SelectedItem
        RaiseEvent RefreshInterface()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim myform = New FormHelper(VendorHelperBS)
        myform.DataGridView1.Columns(0).DataPropertyName = "description"
        If myform.ShowDialog = DialogResult.OK Then
            Dim Helperdrv As DataRowView = VendorHelperBS.Current

          
            drv.Row.Item("vendorcode") = Helperdrv.Row.Item("vendorcode").ToString.Trim
            drv.Row.Item("vendorname") = Helperdrv.Row.Item("vendorname").ToString.Trim
            drv.Row.Item("shortname") = Helperdrv.Row.Item("shortname").ToString.Trim

            'Need bellow code to sync with combobox
            Dim myposition = VendorBS.Find("vendorcode", drv.Row.Item("vendorcode"))
            VendorBS.Position = myposition
            selectionchangecommitted()
        End If
    End Sub

    Private Sub selectionchangecommitted()
        Dim cbdrv As DataRowView = ComboBox1.SelectedItem
        drv.Item("vendorname") = cbdrv.Row.Item("vendorname").ToString.Trim
        drv.Item("shortname") = cbdrv.Row.Item("shortname").ToString.Trim
        RaiseEvent RefreshInterface()
    End Sub

End Class
