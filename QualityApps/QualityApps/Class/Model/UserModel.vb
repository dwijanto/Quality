Imports Npgsql
Public Class UserModel
    Implements IModel



    Dim myadapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance

    Public ReadOnly Property FilterField
        Get
            Return "[userid] like '%{0}%' or [username] like '%{0}%' or [email] like '%{0}%'"
        End Get
    End Property

    Public ReadOnly Property TableName As String Implements IModel.tablename
        Get
            Return "quality.user"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "username"
        End Get
    End Property

    Public Function LoadData(ByVal DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select u.* from {0} u order by {1}", TableName, SortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function

    Public Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean Implements IModel.save
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        AddHandler dataadapter.RowUpdated, AddressOf myadapter.onRowInsertUpdate
        Dim mytransaction As Npgsql.NpgsqlTransaction
        Using conn As Object = myadapter.getConnection

            conn.Open()
            mytransaction = conn.BeginTransaction

            'Update
            Dim sqlstr = "quality.sp_update_user"
            dataadapter.UpdateCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "username").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "email").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isactive").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "quality.sp_insert_user"
            dataadapter.InsertCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "username").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "email").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isactive").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "quality.sp_delete_user"
            dataadapter.DeleteCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.Input
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = dataadapter.Update(mye.dataset.Tables(TableName))

            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function

    Function loglogin(ByVal applicationname As String, ByVal userid As String, ByVal username As String, ByVal computername As String, ByVal time_stamp As Date)
        Dim result As Object
        Using conn As New Npgsql.NpgsqlConnection(myadapter.ConnectionString)
            conn.Open()
            Dim cmd As Npgsql.NpgsqlCommand = New Npgsql.NpgsqlCommand("sp_insertlogonhistory", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = applicationname
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = userid
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = username
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = computername
            result = cmd.ExecuteNonQuery
        End Using
        Return result
    End Function

    Public Function ADDUPDUserManager(ByVal myDATA As List(Of ADPrincipalContext))
        Dim myret As Boolean
        'Manager
        Dim mgrId As Long
        Try
            If myDATA.Count > 1 Then
                'User with ID Manager
                Using conn As New Npgsql.NpgsqlConnection(myadapter.ConnectionString)
                    conn.Open()
                    Dim cmd As Npgsql.NpgsqlCommand = New Npgsql.NpgsqlCommand("quality.sp_addupduser", conn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.Add("iuserid", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = myDATA(1).Userid
                    cmd.Parameters.Add("iusername", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = myDATA(1).DisplayName
                    cmd.Parameters.Add("icompany", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = myDATA(1).Company
                    cmd.Parameters.Add("iemail", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = myDATA(1).Email
                    cmd.Parameters.Add("ititle", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = myDATA(1).Title
                    cmd.Parameters.Add("iemployeenumber", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = myDATA(1).EmployeeNumber
                    cmd.Parameters.Add("idepartment", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = myDATA(1).Department
                    cmd.Parameters.Add("icountry", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = myDATA(1).Country
                    cmd.Parameters.Add("ilocation", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = myDATA(1).Location
                    mgrId = cmd.ExecuteScalar
                End Using
            End If
            Using conn As New Npgsql.NpgsqlConnection(myadapter.ConnectionString)
                conn.Open()
                Dim cmd As Npgsql.NpgsqlCommand = New Npgsql.NpgsqlCommand("quality.sp_addupduser", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("iuserid", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = myDATA(0).Userid
                cmd.Parameters.Add("iusername", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = myDATA(0).UserName
                cmd.Parameters.Add("iparent", NpgsqlTypes.NpgsqlDbType.Bigint, 0).Value = mgrId
                cmd.Parameters.Add("icompany", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = myDATA(0).Company
                cmd.Parameters.Add("iemail", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = myDATA(0).Email
                cmd.Parameters.Add("ititle", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = myDATA(0).Title
                cmd.Parameters.Add("iemployeenumber", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = myDATA(0).EmployeeNumber
                cmd.Parameters.Add("idepartment", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = myDATA(0).Department
                cmd.Parameters.Add("icountry", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = myDATA(0).Country
                cmd.Parameters.Add("ilocation", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = myDATA(0).Location
                cmd.ExecuteScalar()
            End Using
            myret = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            myret = False
        End Try
        Return myret
    End Function

End Class
