Imports Npgsql
Imports System.Text

Public Class ActivityLogModel
    Implements IModel

    Dim myadapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance
    Dim myRBAC As DbManager = New DbManager
    Private DS As DataSet
    Dim userinfo1 As UserInfo = UserInfo.getInstance

    Public Function SQLSTRReport(ByVal criteria As String) As String
        Dim sb As New StringBuilder
        'sb.Append(String.Format("select *,v.vendorcode::text || ' - ' || v.vendorname as vendorcodename,a.activityname,quality.showtimesession(u.timesession) as timesessiondesc from {0} u" &
        '                           " left join quality.activity a on a.id = u.activityid  " &
        '                           " left join quality.vendor v on v.vendorcode = u.vendorcode  " &
        '                           " {1} order by {2};", tablename, criteria, sortField))
        'sb.Append(String.Format("select u.id,u.activitydate,u.timesession / quality.getcountactivity(u.id) as timesession,u.vendorcode,u.userid,u.projectname,u.remark,u.modifiedby,us.username,v.vendorcode,v.shortname::text,v.vendorname::text ,dt.actid as activityid,a.activityname,quality.showtimesession(u.timesession) as timesessiondesc from {0} u" &
        '                        " left join quality.activitylogdtltx dt on dt.hdid = u.id" &
        '                           " left join quality.activity a on a.id = dt.actid  " &
        '                           " left join vendor v on v.vendorcode = u.vendorcode  " &
        '                           " left join quality.user us on lower(us.userid) = lower(u.userid)" &
        '                           " {1} order by {2};", tablename, criteria, "us.username,u.activitydate"))
        'sb.Append(String.Format("select u.id,u.activitydate,u.timesession,u.vendorcode,u.userid,u.projectname,u.remark,u.modifiedby,us.username,v.vendorcode,v.shortname::text,v.vendorname::text ,u.activityid,a.activityname,case u.timesession when 2 then 'OT' else '' end as timesessiondesc,u.postingdate,u.inoffice from {0} u" &
        '                          " left join quality.activity a on a.id = u.activityid  " &
        '                          " left join vendor v on v.vendorcode = u.vendorcode  " &
        '                          " left join quality.user us on lower(us.userid) = lower(u.userid)" &
        '                          " {1} order by {2};", tablename, criteria, "us.username,u.activitydate"))

        sb.Append(String.Format("with ts as (select activitydate,userid,count(activitydate)::numeric from quality.logactivitytx  u {1} group by userid,activitydate) " &
                                " select u.id,u.activitydate,1/ts.count as timesession,u.vendorcode,u.userid,u.projectname,u.remark,u.modifiedby,us.username,v.vendorcode,v.shortname::text,v.vendorname::text ,c.categoryname,u.activityid,a.activityname,case u.timesession when 2 then 'OT' else '' end as timesessiondesc,u.postingdate,u.inoffice from {0} u" &
                                 " left join quality.activity a on a.id = u.activityid  " &
                                 " left join quality.category c on c.id = a.categoryid  " &
                                 " left join vendor v on v.vendorcode = u.vendorcode  " &
                                 " left join quality.user us on lower(us.userid) = lower(u.userid)" &
                                 " left join ts on ts.userid = u.userid and ts.activitydate = u.activitydate" &
                                 " {1} order by {2};", tablename, criteria, "us.username,u.activitydate"))

        Return sb.ToString
    End Function

    Public ReadOnly Property FilterField
        Get
            Return "(vendorcodename like '%{0}%') or (username like '%{0}%') or (projectname like '%{0}%') or (activityname like '%{0}%') "
        End Get
    End Property

    Public Function getActivityDate(ByRef DS As DataSet, ByRef Sqlstr As String) As Boolean
        Dim dataAdapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False

        Using conn As NpgsqlConnection = myadapter.getConnection
            conn.Open()
            dataAdapter.SelectCommand = myadapter.getCommandObject(Sqlstr, conn)
            dataAdapter.SelectCommand.CommandType = CommandType.Text
            dataAdapter.Fill(DS)
            myret = True
        End Using
        Return myret
    End Function


    Public Function LoadData(DS As DataSet) As Boolean Implements IModel.LoadData
        Me.DS = DS
        Dim sb As New StringBuilder
        Dim dataAdapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Dim sqlstr As String = String.Empty
        Using conn As NpgsqlConnection = myadapter.getConnection
            conn.Open()
            sb.Append(String.Format("select *,v.vendorcode::string || ' - ' || v.vendorname as vendorcodename,a.activityname,c.categoryname from {0} u" &
                                   " left join quality.activity a on a.id = u.activityid  " &
                                   " left join quality.category c on c.id = a.categoryid  " &
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
        Dim mygroup As String = String.Empty

        If IsNothing(myRBAC.getAssignment("Key User", User.getId)) Then 'Not Key User
            If User.can("View Activity Group Process Improvement") Then 'Process Improvement
                mygroup = String.Format("where activitygroup = '2'")
            ElseIf User.can("View Activity Group QE User") Then
                mygroup = String.Format("where activitygroup = '1'")
            Else
                Err.Raise("Sorry,You are not Process Improvement or QE user. You cannot use this function.")
            End If
        End If

        'If Not User.can("View Activity Log All Data") Then
        '    If User.can("View Activity Group Process Improvement") Then 'Process Improvement
        '        mygroup = String.Format("where activity group = 2")
        '    ElseIf User.can("View Activity Group QE User") Then
        '        mygroup = String.Format("where activity group = 1")
        '    Else
        '        Err.Raise("Sorry,You are not Process Improvement or QE user. You cannot use this function.")
        '    End If
        'End If

        Using conn As NpgsqlConnection = myadapter.getConnection
            conn.Open()

            sb.Append(String.Format("select u.*,v.vendorcode::text || ' - ' || v.vendorname as vendorcodename,a.activityname,u.timesession = 2 as ot, qu.username,a.categoryid,c.categoryname from {0} u" &
                                   " left join vendor v on v.vendorcode = u.vendorcode  " &
                                   " left join quality.activity a on a.id = u.activityid" &
                                   " left join quality.category c on c.id = a.categoryid  " &
                                   " left join quality.user qu on lower(qu.userid) = lower(u.userid)" &
                                   " {1} order by {2};", tablename, criteria, sortField))
            'sb.Append(String.Format("select *,qv.vendorcode::text || ' - ' || v.shortname::text || ' - ' || v.vendorname as vendorcodename from quality.vendorassignment qv left join vendor v on v.vendorcode = qv.vendorcode where lower(qv.userid) = lower('{0}') order by v.shortname;", userinfo1.Userid.ToLower))
            'sb.Append(String.Format("select v.vendorcode::text || ' - ' || coalesce(v.shortname::text,'') || ' - ' || v.vendorname as name,v.*,v.vendorcode::text || ' - ' || coalesce(v.shortname::text,'') || ' - ' || v.vendorname as vendorcodename from vendor v order by v.shortname;"))
            sb.Append(String.Format("select ''::text as name,null::bigint as vendorcode,''::text as vendorcodename union all (select v.vendorcode::text || ' - ' || coalesce(v.shortname::text,'') || ' - ' || v.vendorname as name,v.vendorcode,v.vendorcode::text || ' - ' || coalesce(v.shortname::text,'') || ' - ' || v.vendorname as vendorcodename from quality.vendorview v order by v.shortname);"))
            'sb.Append(String.Format("select * from quality.activity  {0} order by activitygroup, activityname;", mygroup))
            sb.Append(String.Format("select * from quality.activity where not categoryid isnull order by activityname;"))
            sb.Append(String.Format("select 0.5 as myvalue,'Half day'::text as description" &
                                    " union all select 0.1 as myvalue,'Full day'::text as description" &
                                    " union all select 2 as myvalue,'OT'::text as description;"))
            'sb.Append(String.Format("select dt.*,a.activityname,a.activitygroup from quality.activitylogdtltx dt " &
            '                        " left join quality.activity a on a.id = dt.actid" &
            '                        " left join quality.activitylogtx u on u.id = dt.hdid {0};", criteria))
            sb.Append(String.Format("select u.username as name,u.userid from quality.getsubordinate('{0}') foo left join quality.user u on u.userid = foo order by u.username;", userinfo1.Userid.ToLower))
            sb.Append(String.Format("select u.username as name,u.userid from quality.user u order by u.username;"))
            sb.Append(String.Format("select ivalue from quality.paramhd where paramname = 'cutoffday';"))
            sb.Append(String.Format("select * from quality.category;"))

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
            Return "u.activitydate,u.id"
        End Get
    End Property

    Public ReadOnly Property tablename As String Implements IModel.tablename
        Get
            'Return "quality.activitylogtx"
            Return "quality.logactivitytx"
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
        DS.Tables(4).TableName = "Subordinate"
        DS.Tables(5).TableName = "All Users"
        DS.Tables(6).TableName = "Cutoff"
        DS.Tables(7).TableName = "Category"
        'Dim pk4(0) As DataColumn
        'pk4(0) = DS.Tables(4).Columns("id")
        'DS.Tables(4).PrimaryKey = pk4
        'DS.Tables(4).Columns("id").AutoIncrement = True
        'DS.Tables(4).Columns("id").AutoIncrementSeed = -1
        'DS.Tables(4).Columns("id").AutoIncrementStep = -1

        Dim rel As DataRelation
        Dim hcol As DataColumn
        Dim dcol As DataColumn
        'create relation ds.table(7) and ds.table(2) 'Category as Parent , Activity as Child
        hcol = DS.Tables("Category").Columns("id") 'id in table header
        dcol = DS.Tables("Activity").Columns("categoryid")
        rel = New DataRelation("hdrel", hcol, dcol)
        DS.Relations.Add(rel)

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
            sqlstr = "quality.sp_deletealogactivitytx"
            dataadapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "quality.sp_insertlogactivitytx"
            dataadapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "activitydate").SourceVersion = DataRowVersion.Current            
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "timesession").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "projectname").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "activityid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remark").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifiedby").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Timestamp, 0, "postingdate").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "inoffice").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput

            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "quality.sp_updatelogactivitytx"
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
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "inoffice").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = dataadapter.Update(mye.dataset.Tables(tablename))

            ''Table Detail
            'sqlstr = "quality.sp_deleteaactivitylogdtltx"
            'dataadapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
            'dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            'dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            'sqlstr = "quality.sp_insertactivitylogdtltx"
            'dataadapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            'dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "hdid").SourceVersion = DataRowVersion.Current
            'dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "actid").SourceVersion = DataRowVersion.Current
            'dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput

            'dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            'sqlstr = "quality.sp_updateactivitylogdtltx"
            'dataadapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
            'dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            'dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "hdid").SourceVersion = DataRowVersion.Current
            'dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "actid").SourceVersion = DataRowVersion.Current
            'dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            'dataadapter.InsertCommand.Transaction = mytransaction
            'dataadapter.UpdateCommand.Transaction = mytransaction
            'dataadapter.DeleteCommand.Transaction = mytransaction

            'mye.ra = dataadapter.Update(mye.dataset.Tables("Detail"))


            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function
End Class
