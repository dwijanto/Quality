Public Class HighRiskController
    Implements IController
    Implements IToolbarAction

    Public Model As New HighRiskModel
    Public BS As BindingSource
    Private CMMFBS As BindingSource
    Private VendorBS As BindingSource
    Private StatusBS As BindingSource
    Private BUBS As BindingSource
    Dim DS As DataSet
    Public ReadOnly Property GetCMMFBS As BindingSource
        Get
            Return CMMFBS
        End Get
    End Property

    Public ReadOnly Property GetVendorBS As BindingSource
        Get
            Return VendorBS
        End Get
    End Property
    Public ReadOnly Property GetStatusBS As BindingSource
        Get
            Return StatusBS
        End Get
    End Property
    Public ReadOnly Property GetBUBS As BindingSource
        Get
            Return BUBS
        End Get
    End Property

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

    Public Function loaddata() As Boolean Implements IController.loaddata
        Dim myret As Boolean = False
        Model = New HighRiskModel
        DS = New DataSet
        If Model.LoadData(DS) Then
            BS = New BindingSource
            BS.DataSource = DS.Tables(0)
            CMMFBS = New BindingSource
            CMMFBS.DataSource = DS.Tables(1)
            VendorBS = New BindingSource
            VendorBS.DataSource = DS.Tables(2)
            StatusBS = New BindingSource
            StatusBS.DataSource = DS.Tables(3)
            BUBS = New BindingSource
            BUBS.DataSource = DS.Tables(4)
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
