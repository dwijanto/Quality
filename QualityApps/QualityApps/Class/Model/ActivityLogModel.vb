Imports Npgsql
Imports System.Text

Public Class ActivityLogModel
    Implements IModel

    Dim myadapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance
    Private DS As DataSet
    Dim userinfo1 As UserInfo = UserInfo.getInstance

    Public Function SQLSTRReport(ByVal criteria As String) As String
        Dim sb As New StringBuilder
        'sb.Append(String.Format("select *,v.vendorcode::text || ' - ' || v.vendorname as vendorcodename,a.activityname,quality.showtimesession(u.timesession) as timesessiondesc from {0} u" &
        '                           " left join quality.activity a on a.id = u.activityid  " &
        '                           " left join quality.vendor v on v.vendorcode = u.vendorcode  " &
        '                           " {1} order by {2};", tablename, criteria, sortField))
        sb.Append(String.Format("select u.*,us.username,v.vendorcode,v.shortname::text,v.vendorname::text ,a.activityname,quality.showtimesession(u.timesession) as timesessiondesc from {0} u" &
                                   " left join quality.activity a on a.id = u.activityid  " &
                                   " left join vendor v on v.vendorcode = u.vendorcode  " &
                                   " left join quality.user us on lower(us.userid) = lower(u.userid)" &
                                   " {1} order by {2};", tablename, criteria, sortField))
        Return sb.ToString
    End Function

    Public ReadOnly Property FilterField
        Get
            Return "(vendordocname like '%{0}%') or (projectname like '%{0}%') or (activityname like '%{0}%') "
        End Get
    End Property


    Public Function LoadData(DS As DataSet) As Boolean Implements IModel.LoadData
        Me.DS = DS
        Dim sb As New StringBuilder
        Dim dataAdapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Dim sqlstr As String = String.Empty
        Using conn As NpgsqlConnection = myadapter.getConnection
            conn.Open()
            sb.Append(String.Format("select *,v.vendorcode::string || ' - ' || v.vendorname as vendorcodename,a.activityname from {0} u" &
                                   " left join quality.activity a on a.id = u.activityid  " &
                                   " left join quality.vendor v on v.vendorcode = u.vendorcode  " &
                                   "order by {1};", tablename, sortField))
            sb.Append(String.Format("select * from quality.vendor order by vendorname;"))
            sb.Append(String.Format("select * from quality.activity order by activityname;"))

            dataAdapter.SelectCommand = myadapter.getCommandObject(sb.ToString, conn)
            dataAdapter.SelectCommand.CommandType = CommandType.Text
            dataAdapter.Fill(DS, tablename)
            PrepareDataSet(DS)
            myret = True
        End Using
        Return myret
    End Function

    Public Function LoadData(DS As DataSet, criteria As String) As Boolean
        Me.DS = DS
        Dim sb As New StringBuilder
        Dim dataAdapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Dim sqlstr As String = String.Empty
        Dim mygroup = 1

        If User.can("View Activity Group Process Improvement") Then 'Process Improvement
            mygroup = 2
        ElseIf User.can("View Activity Group QE User") Then
            mygroup = 1
        Else
            Err.Raise("Sorry,You are not Process Improvement or QE user. You cannot use this function.")
        End If


        Using conn As NpgsqlConnection = myadapter.getConnection
            conn.Open()

            sb.Append(String.Format("select *,v.vendorcode::text || ' - ' || v.vendorname as vendorcodename,a.activityname,quality.showtimesession(u.timesession) as timesessiondesc from {0} u" &
                                   " left join quality.activity a on a.id = u.activityid  " &
                                   " left join vendor v on v.vendorcode = u.vendorcode  " &
                                   " {1} order by {2};", tablename, criteria, sortField))
            'sb.Append(String.Format("select *,qv.vendorcode::text || ' - ' || v.shortname::text || ' - ' || qv.vendorname as vendorcodename from quality.vendor qv left join vendor v on v.vendorcode = qv.vendorcode order by v.shortname;"))
            sb.Append(String.Format("select *,qv.vendorcode::text || ' - ' || v.shortname::text || ' - ' || v.vendorname as vendorcodename from quality.vendorassignment qv left join vendor v on v.vendorcode = qv.vendorcode where lower(qv.userid) = lower('{0}') order by v.shortname;", userinfo1.Userid.ToLower))
            sb.Append(String.Format("select * from quality.activity where activitygroup = '{0}' order by activityname;", mygroup))
            sb.Append(String.Format("select 0.5 as myvalue,'Half day'::text as description" &
                                    " union all select 0.1 as myvalue,'Full day'::text as description" &
                                    " union all select 2 as myvalue,'OT'::text as description;"))
            dataAdapter.SelectCommand = myadapter.getCommandObject(sb.ToString, conn)
            dataAdapter.SelectCommand.CommandType = CommandType.Text
            dataAdapter.Fill(DS, tablename)
            PrepareDataSet(DS)
            myret = True
        End Using
        Return myret
    End Function

    Public Function LoadDataHistory(DS As DataSet, criteria As String) As Boolean
        Me.DS = DS
        Dim dataAdapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Dim sqlstr As String = String.Empty
        Using conn As NpgsqlConnection = myadapter.getConnection
            conn.Open()
            sqlstr = String.Format("select * from {0} u {1} order by {2}", tablename, criteria, sortField)

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
            Return "u.id"
        End Get
    End Property

    Public ReadOnly Property tablename As String Implements IModel.tablename
        Get
            Return "quality.activitylogtx"
        End Get
    End Property

    Private Sub PrepareDataSet(ByRef DS As DataSet)
        Dim pk(0) As DataColumn
        pk(0) = DS.Tables(0).Columns("id")
        DS.Tables(0).PrimaryKey = pk
        DS.Tables(0).Columns("id").AutoIncrement = True
        DS.Tables(0).Columns("id").AutoIncrementSeed = -1
        DS.Tables(0).Columns("id").AutoIncrementStep = -1
        DS.Tables(0).TableName = tablename
        DS.Tables(1).TableName = "Vendor"
        DS.Tables(2).TableName = "Activity"
        DS.Tables(3).TableName = "TimeSession"
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
            sqlstr = "quality.sp_deleteaactivitylogtx"
            dataadapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "quality.sp_insertactivitylogtx"
            dataadapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "activitydate").SourceVersion = DataRowVersion.Current            
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "timesession").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "projectname").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "activityid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remark").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifiedby").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput

            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "quality.sp_updateactivitylogtx"
            dataadapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "activitydate").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "timesession").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "projectname").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "activityid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remark").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifiedby").SourceVersion = DataRowVersion.Current
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
