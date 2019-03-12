Imports Npgsql
Imports System.Text

Public Class ParamAdapter
    Public Shared myInstance As ParamAdapter
    Dim myAdapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance
    Public DS As DataSet
    Public BS As BindingSource
    Public BS2 As BindingSource
    Public BS3 As BindingSource
    Public BS4 As BindingSource
    Public BS5 As BindingSource
    Public BS6 As BindingSource
    Public BS7 As BindingSource
    Public BS8 As BindingSource
    Public BS9 As BindingSource
    Public BS10 As BindingSource


    Public Shared Function getInstance() As ParamAdapter
        If myInstance Is Nothing Then
            myInstance = New ParamAdapter
        End If
        Return myInstance
    End Function

    Public Function GetMailFolder(ByVal paramName As String) As String
        Dim sqlstr = String.Format("select cvalue from quality.paramdt where paramname =:paramname", paramName)
        Dim myresult As String = String.Empty
        Dim myparam(0) As NpgsqlParameter
        myparam(0) = New NpgsqlParameter("paramname", paramName)
        myAdapter.ExecuteScalar(sqlstr, myparam, recordAffected:=myresult)
        Return myresult
    End Function

    Public Function GetVendorInfo(ByVal vendorcode As Long) As DataRow
        Dim sqlstr = String.Format("select * from quality.vendor v where v.vendorcode = :paramname")
        Dim ds As New DataSet
        Dim sb As New StringBuilder
        Dim myparam(0) As NpgsqlParameter
        myparam(0) = New NpgsqlParameter("paramname", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "", ParameterDirection.Input, True, 0, 0, DataRowVersion.Default, vendorcode)
        If myAdapter.GetDataset(sqlstr, ds, myparam) Then
            Return ds.Tables(0).Rows(0)
        End If
        Return Nothing
    End Function

    Public Function GetInternalEmail(paramname) As String
        Dim sqlstr = String.Format("select pd.cvalue from quality.paramdt pd left join quality.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :paramname")
        Dim ds As New DataSet
        Dim sb As New StringBuilder
        Dim myparam(0) As NpgsqlParameter
        myparam(0) = New NpgsqlParameter("paramname", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "", ParameterDirection.Input, True, 0, 0, DataRowVersion.Default, paramname)
        If myAdapter.GetDataset(sqlstr, ds, myparam) Then
            For Each dr In ds.Tables(0).Rows
                If sb.Length > 1 Then sb.Append(";")
                sb.Append(dr.item(0))
            Next
        End If
        Return sb.ToString
    End Function

    Public Function GetParentid(ByVal paramName As String) As String
        Dim sqlstr = String.Format("select paramhdid from quality.paramhd where paramname =:paramname", paramName)
        Dim myresult As String = String.Empty
        Dim myparam(0) As NpgsqlParameter
        myparam(0) = New NpgsqlParameter("paramname", paramName)
        myAdapter.ExecuteScalar(sqlstr, myparam, recordAffected:=myresult)
        Return myresult
    End Function

    Public Function GetEWSUser(ByVal paramName As String) As String()
        'Dim sqlstr = String.Format("select cvalue from quality.paramhd where paramname =:paramname", paramName)
        Dim sqlstr = String.Format("select pd.cvalue from quality.paramdt pd left join quality.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :paramname")
        Dim myresult As String = String.Empty
        Dim myparam(0) As NpgsqlParameter
        myparam(0) = New NpgsqlParameter("paramname", paramName)
        myAdapter.ExecuteScalar(sqlstr, myparam, recordAffected:=myresult)
        Return myresult.Split(";")
    End Function

    Public Function GeturlEWS(ByVal paramName As String) As String
        Dim sqlstr = String.Format("select cvalue from quality.paramhd where paramname =:paramname", paramName)
        Dim myresult As String = String.Empty
        Dim myparam(0) As NpgsqlParameter
        myparam(0) = New NpgsqlParameter("paramname", paramName)
        myAdapter.ExecuteScalar(sqlstr, myparam, recordAffected:=myresult)
        Return myresult
    End Function

    Public Function GetVendorExceptionList() As String
        Dim sqlstr = String.Format("select pd.ivalue from quality.paramdt pd left join quality.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :paramname")
        Dim ds As New DataSet
        Dim sb As New StringBuilder
        Dim myparam(0) As NpgsqlParameter
        myparam(0) = New NpgsqlParameter("paramname", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "", ParameterDirection.Input, True, 0, 0, DataRowVersion.Default, "Vendor Exception")
        If myAdapter.GetDataset(sqlstr, ds, myparam) Then
            For Each dr In ds.Tables(0).Rows
                If sb.Length > 1 Then sb.Append(",")
                sb.Append(dr.item(0))
            Next
        End If

        Return sb.ToString
    End Function

    Public Function GetSBUExceptionList() As String
        Dim sqlstr = String.Format("select pd.cvalue from quality.paramdt pd left join quality.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :paramname")
        Dim ds As New DataSet
        Dim sb As New StringBuilder
        Dim myparam(0) As NpgsqlParameter
        myparam(0) = New NpgsqlParameter("paramname", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "", ParameterDirection.Input, True, 0, 0, DataRowVersion.Default, "SBU Exception")
        If myAdapter.GetDataset(sqlstr, ds, myparam) Then
            For Each dr In ds.Tables(0).Rows
                If sb.Length > 1 Then sb.Append(",")
                sb.Append(String.Format("'{0}'", dr.item(0)))
            Next
        End If
        Return sb.ToString
    End Function

    Public Function GetNotification(ByVal mydate As Date) As String
        Dim sqlstr = String.Format("select bodymessage from quality.notification where startdate <= :mydate and enddate >= :mydate")
        Dim ds As New DataSet
        Dim sb As New StringBuilder
        Dim myparam(0) As NpgsqlParameter
        myparam(0) = New NpgsqlParameter("mydate", NpgsqlTypes.NpgsqlDbType.Date, 0, "", ParameterDirection.Input, True, 0, 0, DataRowVersion.Default, mydate)
        If myAdapter.GetDataset(sqlstr, ds, myparam) Then
            For Each dr In ds.Tables(0).Rows
                If sb.Length > 1 Then sb.Append("{0}", "<br>")
                sb.Append(String.Format("{0}", dr.item(0)))
            Next
        End If
        Return sb.ToString
    End Function

    Public Function GetCCETDDateSpan() As Integer
        Dim sqlstr = String.Format("select pd.ivalue from quality.paramdt pd left join quality.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :paramname")
        Dim myresult As Integer
        Dim myparam(0) As NpgsqlParameter
        myparam(0) = New NpgsqlParameter("paramname", "Date Span")
        myAdapter.ExecuteScalar(sqlstr, myparam, recordAffected:=myresult)
        Return myresult
    End Function

    Public Function GetStartDate() As Integer
        'Dim sqlstr = String.Format("select ph.ivalue from quality.paramhd ph where ph.paramname = :paramname")
        Dim sqlstr = String.Format("select pd.ivalue from quality.paramdt pd left join quality.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :paramname")
        Dim myresult As Integer
        Dim myparam(0) As NpgsqlParameter
        myparam(0) = New NpgsqlParameter("paramname", "Start Date")
        myAdapter.ExecuteScalar(sqlstr, myparam, recordAffected:=myresult)
        Return myresult
    End Function

    Public Function GetEndDate() As Integer
        'Dim sqlstr = String.Format("select ph.ivalue from quality.paramhd ph where ph.paramname = :paramname")
        Dim sqlstr = String.Format("select pd.ivalue from quality.paramdt pd left join quality.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :paramname")
        Dim myresult As Integer
        Dim myparam(0) As NpgsqlParameter
        myparam(0) = New NpgsqlParameter("paramname", "End Date")
        myAdapter.ExecuteScalar(sqlstr, myparam, recordAffected:=myresult)
        Return myresult
    End Function
    Public Function GetLastSendDate() As Date
        'Dim sqlstr = String.Format("select ph.dvalue from quality.paramhd ph where ph.paramname = :paramname")
        Dim sqlstr = String.Format("select pd.dvalue from quality.paramdt pd left join quality.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :paramname")
        Dim myresult As Date
        Dim myparam(0) As NpgsqlParameter
        myparam(0) = New NpgsqlParameter("paramname", "Last Send Date")
        myAdapter.ExecuteScalar(sqlstr, myparam, recordAffected:=myresult)
        Return myresult
    End Function

    Public Function UpdateLastSendDate(ByVal mydate As Date) As Boolean
        Dim myret As Boolean = False
        Dim sqlstr = String.Format("update quality.paramdt set dvalue = :paramvalue where paramhdid in (select paramhdid from quality.paramhd where paramname = 'Last Send Date')")
        Dim myparam(0) As NpgsqlParameter

        myparam(0) = New NpgsqlParameter("paramvalue", mydate)
        Dim ra As Integer
        Dim message As String = String.Empty
        myAdapter.ExecuteNonQuery(sqlstr, myparam, message, ra)
        Return myret
    End Function

    Public Function LoadData()
        Dim sb As New StringBuilder
        Dim myret As Boolean = False
        sb.Append(String.Format("select pd.* from quality.paramdt pd left join quality.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :pathparam;"))
        sb.Append(String.Format("select pd.* from quality.paramdt pd left join quality.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :sbuparam order by cvalue;"))
        sb.Append(String.Format("select pd.* from quality.paramdt pd left join quality.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :vendorparam order by ivalue;"))
        sb.Append(String.Format("select pd.* from quality.paramdt pd left join quality.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :internalemailtoparam order by paramname;"))
        sb.Append(String.Format("select pd.* from quality.paramdt pd left join quality.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :internalemailccparam order by paramname;"))
        sb.Append(String.Format("select pd.* from quality.paramdt pd left join quality.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :datespan order by paramname;"))
        sb.Append(String.Format("select pd.* from quality.paramdt pd left join quality.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :startdate order by paramname;"))
        sb.Append(String.Format("select pd.* from quality.paramdt pd left join quality.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :enddate order by paramname;"))
        sb.Append(String.Format("select pd.* from quality.paramdt pd left join quality.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :lastsenddate order by paramname;"))
        sb.Append(String.Format("select pd.* from quality.paramdt pd left join quality.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :userpassword order by paramname;"))
        Dim sqlstr = sb.ToString
        DS = New DataSet
        BS = New BindingSource
        BS2 = New BindingSource
        BS3 = New BindingSource
        BS4 = New BindingSource
        BS5 = New BindingSource
        BS6 = New BindingSource
        BS7 = New BindingSource
        BS8 = New BindingSource
        BS9 = New BindingSource
        BS10 = New BindingSource
        Dim myparam(9) As NpgsqlParameter
        myparam(0) = New NpgsqlParameter("pathparam", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "", ParameterDirection.Input, True, 0, 0, DataRowVersion.Default, "MailFolder")
        myparam(1) = New NpgsqlParameter("sbuparam", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "", ParameterDirection.Input, True, 0, 0, DataRowVersion.Default, "SBU Exception")
        myparam(2) = New NpgsqlParameter("vendorparam", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "", ParameterDirection.Input, True, 0, 0, DataRowVersion.Default, "Vendor Exception")
        myparam(3) = New NpgsqlParameter("internalemailtoparam", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "", ParameterDirection.Input, True, 0, 0, DataRowVersion.Default, "Internal Email (to)")
        myparam(4) = New NpgsqlParameter("internalemailccparam", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "", ParameterDirection.Input, True, 0, 0, DataRowVersion.Default, "Internal Email (cc)")
        myparam(5) = New NpgsqlParameter("datespan", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "", ParameterDirection.Input, True, 0, 0, DataRowVersion.Default, "Date Span")
        myparam(6) = New NpgsqlParameter("startdate", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "", ParameterDirection.Input, True, 0, 0, DataRowVersion.Default, "Start Date")
        myparam(7) = New NpgsqlParameter("enddate", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "", ParameterDirection.Input, True, 0, 0, DataRowVersion.Default, "End Date")
        myparam(8) = New NpgsqlParameter("lastsenddate", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "", ParameterDirection.Input, True, 0, 0, DataRowVersion.Default, "Last Send Date")
        myparam(9) = New NpgsqlParameter("userpassword", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "", ParameterDirection.Input, True, 0, 0, DataRowVersion.Default, "EWS")
        If myAdapter.GetDataset(sqlstr, DS, myparam) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("paramdtid")
            DS.Tables(0).PrimaryKey = pk
            BS.DataSource = DS.Tables(0)
            DS.Tables(0).Columns("paramdtid").AutoIncrement = True
            DS.Tables(0).Columns("paramdtid").AutoIncrementSeed = -1
            DS.Tables(0).Columns("paramdtid").AutoIncrementStep = -1



            Dim pk2(0) As DataColumn
            pk2(0) = DS.Tables(1).Columns("paramdtid")
            DS.Tables(1).PrimaryKey = pk2
            BS2.DataSource = DS.Tables(1)
            DS.Tables(1).Columns("paramdtid").AutoIncrement = True
            DS.Tables(1).Columns("paramdtid").AutoIncrementSeed = -1
            DS.Tables(1).Columns("paramdtid").AutoIncrementStep = -1


            Dim pk3(0) As DataColumn
            pk3(0) = DS.Tables(2).Columns("paramdtid")
            DS.Tables(2).PrimaryKey = pk3
            BS3.DataSource = DS.Tables(2)
            DS.Tables(2).Columns("paramdtid").AutoIncrement = True
            DS.Tables(2).Columns("paramdtid").AutoIncrementSeed = -1
            DS.Tables(2).Columns("paramdtid").AutoIncrementStep = -1

            Dim pk4(0) As DataColumn
            pk4(0) = DS.Tables(3).Columns("paramdtid")
            DS.Tables(3).PrimaryKey = pk4
            BS4.DataSource = DS.Tables(3)
            DS.Tables(3).Columns("paramdtid").AutoIncrement = True
            DS.Tables(3).Columns("paramdtid").AutoIncrementSeed = -1
            DS.Tables(3).Columns("paramdtid").AutoIncrementStep = -1

            Dim pk5(0) As DataColumn
            pk5(0) = DS.Tables(4).Columns("paramdtid")
            DS.Tables(4).PrimaryKey = pk5
            BS5.DataSource = DS.Tables(4)
            DS.Tables(4).Columns("paramdtid").AutoIncrement = True
            DS.Tables(4).Columns("paramdtid").AutoIncrementSeed = -1
            DS.Tables(4).Columns("paramdtid").AutoIncrementStep = -1

            BS6.DataSource = DS.Tables(5)
            BS7.DataSource = DS.Tables(6)
            BS8.DataSource = DS.Tables(7)
            BS9.DataSource = DS.Tables(8)
            BS10.DataSource = DS.Tables(9)
            myret = True
        End If
        Return myret
    End Function

    Public Function save() As Boolean
        Dim myret As Boolean = False
        BS.EndEdit()
        BS2.EndEdit()
        BS3.EndEdit()
        BS4.EndEdit()
        BS5.EndEdit()
        BS6.EndEdit()
        BS7.EndEdit()
        BS8.EndEdit()
        BS9.EndEdit()
        BS10.EndEdit()
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

    Public Function Save(mye As ContentBaseEventArgs) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myAdapter.getDbDataAdapter
        Dim myret As Boolean = False
        AddHandler dataadapter.RowUpdated, AddressOf myAdapter.onRowInsertUpdate
        Dim mytransaction As Npgsql.NpgsqlTransaction
        Using conn As Object = myAdapter.getConnection
            conn.Open()
            mytransaction = conn.BeginTransaction

            Dim sqlstr As String
            sqlstr = "quality.sp_deleteparameter"
            dataadapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "paramdtid").SourceVersion = DataRowVersion.Original
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "quality.sp_insertparameter"
            dataadapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "paramhdid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "paramname").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cvalue").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "dvalue").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "ivalue").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nvalue").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Timestamp, 0, "ts").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "paramdtid").Direction = ParameterDirection.InputOutput
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "quality.sp_updateparameter"
            dataadapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "paramdtid").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "paramhdid").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "paramname").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cvalue").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "dvalue").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "ivalue").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nvalue").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Timestamp, 0, "ts").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = dataadapter.Update(mye.dataset.Tables(0))
            mye.ra = dataadapter.Update(mye.dataset.Tables(1))
            mye.ra = dataadapter.Update(mye.dataset.Tables(2))
            mye.ra = dataadapter.Update(mye.dataset.Tables(3))
            mye.ra = dataadapter.Update(mye.dataset.Tables(4))
            mye.ra = dataadapter.Update(mye.dataset.Tables(5))
            mye.ra = dataadapter.Update(mye.dataset.Tables(6))
            mye.ra = dataadapter.Update(mye.dataset.Tables(7))
            mye.ra = dataadapter.Update(mye.dataset.Tables(8))
            mye.ra = dataadapter.Update(mye.dataset.Tables(9))
            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function

    Public Function SaveSendEmailTx() As Boolean
        Dim userinfo1 = UserInfo.getInstance
        Dim myret As Boolean = False
        Try
            Dim sqlstr = String.Format("insert into quality.sendemailtx(userid) values(:userid)")
            Dim myresult As String = String.Empty
            Dim myparam(0) As NpgsqlParameter
            myparam(0) = New NpgsqlParameter("userid", userinfo1.Userid)
            myAdapter.ExecuteScalar(sqlstr, myparam, recordAffected:=myresult)
            myret = IsNothing(myresult)
        Catch ex As Exception
        End Try
        Return myret
    End Function
End Class
