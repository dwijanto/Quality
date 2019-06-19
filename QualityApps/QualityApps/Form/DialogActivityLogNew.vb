Imports System.Windows.Forms
Imports System.ComponentModel

Public Class DialogActivityLogNew
    Implements INotifyPropertyChanged

    Dim myRBAC As DbManager = New DbManager
    Dim drv As DataRowView
    Dim vendorbs As BindingSource
    Dim ActivityBS As BindingSource
    Dim ds As DataSet
    Public Event RefreshInterface()

    Public Property WorkingOT As Integer
        Get
            If CheckBox1.Checked = True Then
                Return 2
            Else
                Return 0
            End If
        End Get
        Set(value As Integer)
            If value = 2 Then
                CheckBox1.Checked = True
            Else
                CheckBox1.Checked = False
            End If
        End Set
    End Property



    Public Sub New(DRV As DataRowView, vendorbs As BindingSource, activitybs As BindingSource, timesessionbs As BindingSource, DS As DataSet)
        InitializeComponent()

        Me.drv = DRV
        Me.drv.BeginEdit()
        Me.vendorbs = vendorbs
        Me.ActivityBS = activitybs
        Me.ds = DS

        InitData()
    End Sub
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.validate Then
            drv.EndEdit()
            Me.DialogResult = System.Windows.Forms.DialogResult.OK

            Me.Close()
        End If

    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        drv.CancelEdit()
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
    Public Overloads Function validate() As Boolean
        'Dim myret As Boolean = False
        'If IsNothing(myRBAC.getAssignment("Key User", User.getId)) Then
        '    'Not Key User
        '    Dim ActivityDate As Date = drv.Row.Item("activitydate")
        '    ErrorProvider1.SetError(DateTimePicker1, "")
        '    If ActivityDate.Month = Today.Month Then
        '        myret = True
        '    ElseIf ActivityDate.Month + 1 = Month(Today) Then
        '        If Today.Date.Day <= ds.Tables("Cutoff").Rows(0).Item("ivalue") Then
        '            myret = True
        '        Else
        '            ErrorProvider1.SetError(DateTimePicker1, String.Format("This Activity date  '{0:dd-MMM-yyyy}' has been closed. Please contact Key User!", ActivityDate))
        '        End If
        '    Else
        '        ErrorProvider1.SetError(DateTimePicker1, String.Format("This Activity date  '{0:dd-MMM-yyyy}' has been closed. Please contact Key User!", ActivityDate))
        '    End If
        'Else
        '    myret = True 'Key User
        'End If
        'Dim cbdrv As DataRowView = ComboBox1.SelectedItem
        'If IsNothing(cbdrv) Then
        '    ErrorProvider1.SetError(ComboBox1, "Please select from the list.")
        '    myret = False
        'Else
        '    If Not cbdrv.Row.Item("activityname") = "On leave" Then
        '        If IsDBNull(drv.Row.Item("vendorcode")) Then
        '            ErrorProvider1.SetError(TextBox3, "Vendorcode cannot be blank!")
        '            myret = False
        '        End If
        '    End If
        'End If

        'Return myret
        Return True
    End Function
    Private Sub InitData()
        ComboBox1.DataSource = ActivityBS
        ComboBox1.DisplayMember = "activityname"
        ComboBox1.ValueMember = "id"

        ComboBox1.DataBindings.Clear()
        DateTimePicker1.DataBindings.Clear()
        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        TextBox4.DataBindings.Clear()
        CheckBox1.DataBindings.Clear()
        CheckBox2.DataBindings.Clear()
        ComboBox1.DataBindings.Add(New Binding("SelectedValue", drv, "activityid", True, DataSourceUpdateMode.OnPropertyChanged))
        DateTimePicker1.DataBindings.Add(New Binding("Text", drv, "activitydate", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox1.DataBindings.Add(New Binding("Text", drv, "projectname", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox2.DataBindings.Add(New Binding("Text", drv, "remark", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox3.DataBindings.Add(New Binding("Text", drv, "vendorcodename", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox4.DataBindings.Add(New Binding("Text", drv, "username", True, DataSourceUpdateMode.OnPropertyChanged))
        CheckBox2.DataBindings.Add(New Binding("checked", drv, "inoffice", True, DataSourceUpdateMode.OnPropertyChanged))
        Me.DataBindings.Add(New Binding("WorkingOT", drv, "timesession", True, DataSourceUpdateMode.OnPropertyChanged))
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'if Key User can see all user, otherwise depend on getsubordinate.
        Dim mybs As New BindingSource
        If Not IsNothing(myRBAC.getAssignment("Key User", User.getId)) Then
            ' is Key User
            mybs.DataSource = ds.Tables("All Users")
        Else
            'not key user
            mybs.DataSource = ds.Tables("Subordinate")
        End If
        Dim myform = New FormHelper(mybs)
        If myform.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim hdrv As DataRowView = mybs.Current
            drv.Row.Item("userid") = hdrv.Row.Item("userid")
            drv.Row.Item("username") = hdrv.Row.Item("name")
            TextBox4.Text = drv.Row.Item("username")
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim myform = New FormHelper(vendorbs)
        If myform.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim hdrv As DataRowView = vendorbs.Current
            drv.Row.Item("vendorcode") = hdrv.Row.Item("vendorcode")
            drv.Row.Item("vendorcodename") = hdrv.Row.Item("vendorcodename")
            TextBox3.Text = drv.Row.Item("vendorcodename")
        End If
    End Sub

    Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        drv.Row.Item("ot") = CheckBox1.Checked
        RaiseEvent RefreshInterface()
        onPropertyChanged("WorkingOT")
    End Sub

    Private Sub onPropertyChanged(name As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(name))
    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        Dim cbdrv As DataRowView = ComboBox1.SelectedItem
        drv.Row.Item("activityname") = cbdrv.Row.Item("activityname")
        RaiseEvent RefreshInterface()
    End Sub


    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged, TextBox3.TextChanged, TextBox2.TextChanged, DateTimePicker1.ValueChanged
        RaiseEvent RefreshInterface()
    End Sub

End Class
