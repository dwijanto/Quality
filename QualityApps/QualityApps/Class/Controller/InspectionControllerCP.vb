Public Class InspectionControllerCP
    Implements IToolbarAction
    Implements IController

    Public Model As New InspectionModelCP
    Public DS As DataSet
    Public DSMissing As DataSet
    Public BS As BindingSource
    Public BSMissing As BindingSource

    Public ReadOnly Property SQLSTR As String
        Get
            Return Model.sqlstr
        End Get
    End Property

    Public ReadOnly Property SQLSTRExcel As String
        Get
            Return Model.sqlstrExcel
        End Get
    End Property

    Public Function LoadData() As Boolean Implements IController.loaddata
        Dim myret As Boolean = True
        DS = New DataSet
        myret = Model.LoadData(DS)
        If myret Then
            Dim pk(2) As DataColumn
            pk(0) = DS.Tables(0).Columns("Purch.Doc.")
            pk(1) = DS.Tables(0).Columns("Item")
            pk(2) = DS.Tables(0).Columns("SeqN")
            DS.Tables(0).PrimaryKey = pk
            BS = New BindingSource
            BS.DataSource = DS.Tables(0)
        End If
        Return myret
    End Function


    Public Property ApplyFilter As String Implements IToolbarAction.ApplyFilter
        Get
            Return BS.Filter
        End Get
        Set(ByVal value As String)
            Try
                BS.Filter = String.Format(Model.FilterField, value)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Set
    End Property

    Public Property ApplyFilterMissing As String
        Get
            Return BSMissing.Filter
        End Get
        Set(ByVal value As String)
            Try
                BSMissing.Filter = String.Format(Model.FilterFieldMissing, value)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Set
    End Property

    Public Property ApplyFilterAll As String
        Get
            Return BS.Filter
        End Get
        Set(ByVal value As String)
            Try
                BS.Filter = String.Format("{0}", value)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Set
    End Property

    Public Function GetCurrentRecord() As DataRowView Implements IToolbarAction.GetCurrentRecord
        Return BSMissing.Current
    End Function

    Public Function GetCurrentHistoryRecord() As DataRowView
        Return Model.BSHistory.Current
    End Function

    Public Function GetNewRecord() As DataRowView Implements IToolbarAction.GetNewRecord
        Throw New NotImplementedException
    End Function

    Public Sub RemoveAt(value As Integer) Implements IToolbarAction.RemoveAt
        BSMissing.RemoveAt(value)
    End Sub

    Public Sub RemoveHistoryAt(value As Integer)
        Model.BSHistory.RemoveAt(value)
    End Sub

    'Public Function Save(mye As ContentBaseEventArgs) As Boolean Implements IToolbarAction.Save
    '    Throw New NotImplementedException
    'End Function

    Public Function Save(mye As ContentBaseEventArgs) As Boolean Implements IToolbarAction.Save
        Dim myret As Boolean = False
        If Model.savehistory(Me, mye) Then
            myret = True


        End If
        Return myret
    End Function

    Function loaddataAllMising() As Boolean
        Dim myret As Boolean = True
        DSMissing = New DataSet

        myret = Model.LoadDataAllMissing(DSMissing)
        If myret Then
            Dim pk(0) As DataColumn
            pk(0) = DSMissing.Tables(0).Columns("id")
            DSMissing.Tables(0).PrimaryKey = pk
            BSMissing = New BindingSource
            BSMissing.DataSource = DSMissing.Tables(0)
        End If
        Return myret
    End Function

    Function loaddataAllMisingInspDate() As Boolean
        Dim myret As Boolean = True
        DSMissing = New DataSet

        myret = Model.LoadDataAllMissingInspDate(DSMissing)
        If myret Then
            Dim pk(0) As DataColumn
            pk(0) = DSMissing.Tables(0).Columns("id")
            DSMissing.Tables(0).PrimaryKey = pk
            BSMissing = New BindingSource
            BSMissing.DataSource = DSMissing.Tables(0)
        End If
        Return myret
    End Function

    Public ReadOnly Property GetTable As DataTable Implements IController.GetTable
        Get
            Return Nothing
        End Get
    End Property



    Public Function save() As Boolean Implements IController.save
        Dim myret As Boolean = False
        BSMissing.EndEdit()
        Dim ds2 As DataSet = DSMissing.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If save(mye) Then
                    DSMissing.Merge(ds2)
                    DSMissing.AcceptChanges()
                    MessageBox.Show("Saved.")
                    myret = True
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                DSMissing.Merge(ds2)
            End Try
        End If
        Return myret
    End Function

    Public Function savehistory() As Boolean
        Dim myret As Boolean = False
        Model.BSHistory.EndEdit()
        Dim ds2 As DataSet = Model.DSHistory.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If Save(mye) Then
                    Model.DSHistory.Merge(ds2)
                    Model.DSHistory.AcceptChanges()
                    MessageBox.Show("Saved.")
                    myret = True
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                Model.DSHistory.Merge(ds2)
            End Try
        End If
        Return myret
    End Function

End Class
