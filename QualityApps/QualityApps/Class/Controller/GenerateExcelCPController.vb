Public Class GenerateExcelCPController
    Public myModel As New GenerateExcelCPModel
    Public DS As DataSet
    Public BS As BindingSource

    Public Function LoadData() As Boolean
        Dim myret As Boolean = True
        DS = New DataSet
        myret = myModel.LoadData(DS)
        If myret Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("vendor")
            DS.Tables(0).PrimaryKey = pk
            BS = New BindingSource
            BS.DataSource = DS.Tables(0)
        End If
        Return myret
    End Function


    Public Sub showMessage(ByVal message As String, ByVal fieldname As String)
        Dim myValue() = message.Split(",")
        Dim mykey(0) As Object
        mykey(0) = myValue(0)
        Dim myresult = DS.Tables(0).Rows.Find(mykey)
        If Not IsNothing(myresult) Then
            myresult.Item(fieldname) = myValue(1)
            myresult.EndEdit()
        End If
    End Sub

End Class
