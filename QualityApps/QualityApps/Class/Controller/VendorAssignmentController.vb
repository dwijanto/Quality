Public Class VendorAssignmentController
    Implements IController
    Implements IToolbarAction

    Public Model As New VendorAssignmentModel
    Public BS As BindingSource
    Dim DS As DataSet

    Dim UserBS As BindingSource
    Dim VendorBS As BindingSource
    Dim VendorHelperBS As BindingSource
    Dim EQBS As BindingSource

    Public ReadOnly Property GetDataset As DataSet
        Get
            Return DS
        End Get
    End Property
    Public ReadOnly Property GetVendorBS As BindingSource
        Get
            Return VendorBS
        End Get
    End Property

    Public ReadOnly Property GetVendorHelperBS As BindingSource
        Get
            Return VendorHelperBS
        End Get
    End Property
    Public ReadOnly Property GetUserBS As BindingSource
        Get
            Return UserBS
        End Get
    End Property
    Public ReadOnly Property GetTable As DataTable Implements IController.GetTable
        Get
            Return DS.Tables(Model.tablename).Copy()
        End Get
    End Property



    Public Function loaddata() As Boolean Implements IController.loaddata
        Dim myret As Boolean = False
        Model = New VendorAssignmentModel
        DS = New DataSet
        If Model.LoadData(DS) Then
            BS = New BindingSource
            BS.DataSource = DS.Tables(0)

            VendorBS = New BindingSource
            VendorBS.DataSource = DS.Tables("Vendor")
            VendorHelperBS = New BindingSource
            VendorHelperBS.DataSource = New DataView(DS.Tables("Vendor"))
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
