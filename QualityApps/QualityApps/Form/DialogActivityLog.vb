Imports System.Windows.Forms
Imports System.Text

Public Class DialogActivityLog
    Private drv As DataRowView

    Private TimeSessions As New List(Of TimeSession)
    Private VendorBS As BindingSource
    Private ActivityBS As BindingSource
    Private TimeSessionBS As BindingSource
    Private DetailBS As BindingSource
    Public Event RefreshInterface()
    Dim DetailCurrDrv As DataRowView
    Private DS As DataSet

    Public Sub New(DRV As DataRowView, vendorbs As BindingSource, activitybs As BindingSource, timesessionbs As BindingSource, dtl As BindingSource, DS As DataSet)
        InitializeComponent()

        Me.drv = DRV
        Me.drv.BeginEdit()
        Me.VendorBS = vendorbs
        Me.ActivityBS = activitybs
        Me.DetailBS = dtl
        Me.DS = DS
        'Me.TimeSessionBS = timesessionbs
        InitData()
    End Sub

    Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        ErrorProvider1.SetError(ComboBox1, "")
        ErrorProvider1.SetError(ComboBox2, "")
        ErrorProvider1.SetError(DataGridView1, "")

        If ComboBox1.SelectedIndex = -1 Then
            ErrorProvider1.SetError(ComboBox1, "Please select from the list.")
            myret = False
        End If

        If ComboBox2.SelectedIndex = -1 Then
            drv.Item("vendorcodename") = ""
            drv.Item("vendorcode") = DBNull.Value
            RaiseEvent RefreshInterface()
            If TextBox2.Text.Length = 0 Then
                ErrorProvider1.SetError(ComboBox2, "Please select the vendor. For other vendors, please put it into ""Remarks"".")
                myret = False
            End If

        End If

        'If ComboBox3.SelectedIndex = -1 Then
        '    ErrorProvider1.SetError(ComboBox3, "Please select from the list.")
        '    myret = False
        'End If
        Dim sb As New StringBuilder
        For i = 0 To DataGridView1.Rows.Count - 1
            If sb.Length > 0 Then
                sb.Append(", ")
            End If
            sb.Append(String.Format("{0}", DataGridView1.Rows(i).Cells(0).FormattedValue))
        Next
        If sb.Length > 0 Then
            drv.Row.Item("activityname") = sb.ToString
        Else
            ErrorProvider1.SetError(DataGridView1, "Please add some Activity.")
            myret = False
        End If

        Return myret
    End Function

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        DetailBS.EndEdit()

        If Me.validate Then
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            drv.EndEdit()
            DetailBS.EndEdit()
            RaiseEvent RefreshInterface()
            'Me.Close()
        End If

    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        If Not drv.Row.RowState = DataRowState.Detached Then
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
            'Use the DataTable to find the Deleted row.
            Try
                For i = 0 To DS.Tables(4).Rows.Count - 1
                    Dim dr = DS.Tables(4).Rows(i)
                    If dr.RowState = DataRowState.Deleted Then
                        If dr.Item("hdid", DataRowVersion.Original) = drv.Row.Item("id") Then
                            dr.RejectChanges()
                        End If
                    Else
                        If dr.Item("hdid") = drv.Row.Item("id") Then
                            dr.RejectChanges()
                        End If
                    End If
                Next                
            Catch ex As Exception

            Finally
                Dim sb As New StringBuilder
                For i = 0 To DataGridView1.Rows.Count - 1
                    If sb.Length > 0 Then
                        sb.Append(", ")
                    End If
                    sb.Append(String.Format("{0}", DataGridView1.Rows(i).Cells(0).FormattedValue))
                Next
                If sb.Length > 0 Then
                    drv.Row.Item("activityname") = sb.ToString
                End If
            End Try

            DetailBS.CancelEdit()
            drv.CancelEdit()
            RaiseEvent RefreshInterface()
        End If
    End Sub

    Private Sub InitData()
        TimeSessions.Add(New TimeSession("Full Day", 1))
        TimeSessions.Add(New TimeSession("Half Day", 0.5))
        TimeSessions.Add(New TimeSession("Over Time", 2))

        TimeSessionBS = New BindingSource
        TimeSessionBS.DataSource = TimeSessions

        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = DetailBS
        With DirectCast(DataGridView1.Columns("Column1"), DataGridViewComboBoxColumn)
            .DataSource = ActivityBS
            .DisplayMember = "activityname"
            .ValueMember = "id"
            .DataPropertyName = "actid" 'be careful, actid data type must be the same with valuemember id data type. otherwise it will display actid not activityname.
        End With

        ComboBox1.DataSource = TimeSessionBS
        ComboBox1.DisplayMember = "description"
        ComboBox1.ValueMember = "myvalue"

        ComboBox2.DisplayMember = "vendorcodename"
        ComboBox2.ValueMember = "vendorcode"
        ComboBox2.DataSource = VendorBS

        'ComboBox3.DisplayMember = "activityname"
        'ComboBox3.ValueMember = "id"
        'ComboBox3.DataSource = ActivityBS

        ComboBox1.DataBindings.Clear()
        ComboBox2.DataBindings.Clear()
        'ComboBox3.DataBindings.Clear()
        DateTimePicker1.DataBindings.Clear()
        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()

        ComboBox1.DataBindings.Add(New Binding("SelectedValue", drv, "timesession", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox2.DataBindings.Add(New Binding("SelectedValue", drv, "vendorcode", True, DataSourceUpdateMode.OnPropertyChanged))
        'ComboBox3.DataBindings.Add(New Binding("SelectedValue", drv, "activityid", True, DataSourceUpdateMode.OnPropertyChanged))
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

    'Private Sub ComboBox3_SelectionChangeCommitted(sender As Object, e As EventArgs)
    '    Dim cbdrv As DataRowView = ComboBox3.SelectedItem
    '    drv.Item("activityname") = cbdrv.Row.Item("activityname")
    '    RaiseEvent RefreshInterface()
    'End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged
        RaiseEvent RefreshInterface()
    End Sub

    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError

    End Sub

    Private Sub AddActivityToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddActivityToolStripMenuItem.Click
        drv.EndEdit() 'set rowset header from detached into added
        DetailBS.AddNew()
    End Sub

    Private Sub DeleteActivityToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteActivityToolStripMenuItem.Click
        If Not IsNothing(DetailBS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    DetailBS.RemoveAt(drv.Index)
                Next
            End If
        End If
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