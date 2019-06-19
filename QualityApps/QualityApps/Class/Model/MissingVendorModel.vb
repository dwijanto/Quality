Imports Npgsql
Public Class MissingVendorModel
    Implements IModel
    Dim myadapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance
    Private DS As DataSet

    Public ReadOnly Property tablename As String Implements IModel.tablename
        Get
            Return "quality.missingvendor"
        End Get
    End Property


    Public ReadOnly Property FilterField
        Get
            Return "vendorcode like '%{0}%' or vendorname like '%{0}%' "
        End Get
    End Property


    Public Function LoadData(DS As DataSet) As Boolean Implements IModel.LoadData
        Me.DS = DS
        Dim dataAdapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Dim sqlstr As String = String.Empty
        Using conn As NpgsqlConnection = myadapter.getConnection
            conn.Open()
            sqlstr = String.Format("select * from {0} u order by {1}", tablename, sortField)

            dataAdapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataAdapter.SelectCommand.CommandType = CommandType.Text
            dataAdapter.Fill(DS, tablename)
            PrepareDataSet(DS)
            myret = True
        End Using
        Return myret
    End Function



    Public ReadOnly Property sortField As String Implements IModel.sortField
        Get
            Return "vendorcode"
        End Get
    End Property



    Private Sub PrepareDataSet(ByRef DS As DataSet)
        Dim pk(0) As DataColumn
        pk(0) = DS.Tables(0).Columns("vendorcode")
        DS.Tables(0).PrimaryKey = pk
        DS.Tables(0).TableName = tablename

    End Sub

    Public Function save(obj As Object, mye As ContentBaseEventArgs) As Boolean Implements IModel.save
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        AddHandler dataadapter.RowUpdated, AddressOf myadapter.onRowInsertUpdate
        Dim mytransaction As Npgsql.NpgsqlTransaction
        Using conn As Object = myadapter.getConnection
            conn.Open()
            mytransaction = conn.BeginTransaction

            Dim sqlstr As String
            sqlstr = "quality.sp_deletemissingvendor"
            dataadapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Original
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "quality.sp_insertmissingvendor"
            dataadapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "vendorname").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "shortname").SourceVersion = DataRowVersion.Current


            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "quality.sp_updatemissingvendor"
            dataadapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "vendorname").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "shortname").SourceVersion = DataRowVersion.Current

            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = dataadapter.Update(mye.dataset.Tables(tablename))

            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function
End Class
