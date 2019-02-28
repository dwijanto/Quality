Public Class ActivityLogController
    Implements IController
    Implements IToolbarAction

    Public Model As New ActivityLogModel
    Public BS As BindingSource
    Dim DS As DataSet
    Private VendorBS As BindingSource
    Private ActivityBS As BindingSource
    Private TimeSessionBS As BindingSource

    Public ReadOnly Property GetDataset As DataSet
        Get
            Return DS
        End Get
    End Property
    Public ReadOnly Property GetTable As DataTable Implements IController.GetTable
        Get
            Return DS.Tables(Model.tablename).Copy()
        End Get
    End Property

    Public ReadOnly Property GetVendorBS As BindingSource
        Get
            Return VendorBS
        End Get
    End Property

    Public ReadOnly Property GetActivityBS As BindingSource
        Get
            Return ActivityBS
        End Get
    End Property
    Public ReadOnly Property GetTimeSessionBS As BindingSource
        Get
            Return TimeSessionBS
        End Get
    End Property

    Public Function GetSQLSTRReport(criteria) As String
        Return Model.SQLSTRReport(criteria)
    End Function

    Public Function loaddata() As Boolean Implements IController.loaddata
        Dim myret As Boolean = False
        Model = New ActivityLogModel
        DS = New DataSet
        If Model.LoadData(DS) Then
            BS = New BindingSource
            BS.DataSource = DS.Tables(0)
            myret = True
        End If
        Return myret
    End Function

    Public Function loaddata(criteria As String) As Boolean
        Dim myret As Boolean = False
        Model = New ActivityLogModel
        DS = New DataSet
        If Model.LoadData(DS, criteria) Then
            BS = New BindingSource
            BS.DataSource = DS.Tables(0)

            VendorBS = New BindingSource
            VendorBS.DataSource = DS.Tables("Vendor")
            ActivityBS = New BindingSource
            ActivityBS.DataSource = DS.Tables("Activity")
            TimeSessionBS = New BindingSource
            TimeSessionBS.DataSource = DS.Tables("TimeSession")
            myret = True
        End If
        Return myret
    End Function

    Public Function save() As Boolean Implements IController.save
        Dim myret As Boolean = False
        BS.EndEdit()
        Dim ds2 As DataSet = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If save(mye) Then
                    DS.Merge(ds2)
                    DS.AcceptChanges()
                    MessageBox.Show("Saved.")
                    myret = True
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                DS.Merge(ds2)
            End Try
        End If
        Return myret
    End Function

    Public Property ApplyFilter As String Implements IToolbarAction.ApplyFilter
        Get
            Return BS.Filter
        End Get
        Set(ByVal value As String)
            BS.Filter = String.Format(Model.FilterField, value)
        End Set
    End Property

    Public Property ApplyFilterAll As String
        Get
            Return BS.Filter
        End Get
        Set(ByVal value As String)
            BS.Filter = String.Format(Model.FilterField, value)
        End Set
    End Property
    Public Function GetCurrentRecord() As DataRowView Implements IToolbarAction.GetCurrentRecord
        Return BS.Current
    End Function

    Public Function GetNewRecord() As DataRowView Implements IToolbarAction.GetNewRecord
        Return BS.AddNew
    End Function

    Public Sub RemoveAt(value As Integer) Implements IToolbarAction.RemoveAt
        BS.RemoveAt(value)
    End Sub

    Public Function Save(mye As ContentBaseEventArgs) As Boolean Implements IToolbarAction.Save
        Dim myret As Boolean = False
        If Model.save(Me, mye) Then
            myret = True
        End If
        Return myret
    End Function
End Class
