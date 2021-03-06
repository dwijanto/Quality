﻿Imports System.Windows.Forms

Public Class DialogVendorSPAssignment

    Private drv As DataRowView
    Private VendorBS As BindingSource
    Private SBUBS As BindingSource
    Private VendorHelperBS As BindingSource
    Public Event RefreshInterface()

    Public Sub New(DRV As DataRowView, vendorbs As BindingSource, vendorhelperbs As BindingSource, SBUBS As BindingSource)
        InitializeComponent()

        Me.drv = DRV
        Me.drv.BeginEdit()
        Me.VendorBS = vendorbs
        Me.VendorHelperBS = vendorhelperbs
        Me.SBUBS = SBUBS
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
        'ComboBox2.DisplayMember = "description"
        'ComboBox2.ValueMember = "sbu"
        'ComboBox2.DataSource = sbubs

     
        'ComboBox2.DataBindings.Clear()


        'ComboBox2.DataBindings.Add(New Binding("SelectedValue", drv, "bu", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox2.DataBindings.Add(New Binding("Text", drv, "description", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox1.DataBindings.Add(New Binding("Text", drv, "spcode", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox3.DataBindings.Add(New Binding("Text", drv, "bck1", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox4.DataBindings.Add(New Binding("Text", drv, "bck2", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox5.DataBindings.Add(New Binding("Text", drv, "bu", True, DataSourceUpdateMode.OnPropertyChanged))
    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs)
        selectionchangecommitted()

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim myform = New FormHelper(VendorHelperBS)
        myform.DataGridView1.Columns(0).DataPropertyName = "description"
        If myform.ShowDialog = DialogResult.OK Then
            Dim Helperdrv As DataRowView = VendorHelperBS.Current
            drv.Row.Item("vendorcode") = Helperdrv.Row.Item("vendorcode").ToString.Trim
            drv.Row.Item("vendorname") = Helperdrv.Row.Item("vendorname").ToString.Trim
            drv.Row.Item("shortname") = Helperdrv.Row.Item("shortname").ToString.Trim
            TextBox2.Text = Helperdrv.Row.Item("description").ToString.Trim
            selectionchangecommitted()
        End If
    End Sub

    Private Sub selectionchangecommitted()
        RaiseEvent RefreshInterface()
    End Sub

    'Private Sub ComboBox2_SelectionChangeCommitted(sender As Object, e As EventArgs)
    '    Dim cbdrv As DataRowView = ComboBox2.SelectedItem
    '    drv.Row.Item("bu") = cbdrv.Row.Item("description")
    '    RaiseEvent RefreshInterface()
    'End Sub

End Class
