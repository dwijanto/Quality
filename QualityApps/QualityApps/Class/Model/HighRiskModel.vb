Imports Npgsql
Imports System.Text

Public Class HighRiskModel
    Implements IModel

    Dim myadapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance
    Private DS As DataSet

    Public ReadOnly Property FilterField
        Get
            Return "cmmf::texk like '%{0}%' or description like '%{0}%' or vendorcode::text like '%{0}%' or vendorname like '%{0}%' or sbu like '%{0}%'  or qe like '%{0}%'"
        End Get
    End Property


    Public Function LoadData(DS As DataSet) As Boolean Implements IModel.LoadData
        Me.DS = DS
        Dim dataAdapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Dim sqlstr As String = String.Empty
        Dim sb As New StringBuilder

        Using conn As NpgsqlConnection = myadapter.getConnection
            conn.Open()
            sb.Append(String.Format("with st as (select pd.ivalue as id,pd.cvalue as statusname from quality.paramdt pd" &
                                    " left join quality.paramhd ph on ph.paramhdid = pd.paramhdid" &
                                    " where ph.paramname = 'HighRiskStatus') select *,st.statusname,c.materialdesc,v.vendorname,c.cmmf || '-' || c.materialdesc as cmmfdescription,v.vendorcode || '-' || v.vendorname as vendordescription from {0} u " &
                                    " left join cmmf c on c.cmmf = u.cmmf left join vendor v on v.vendorcode = u.vendorcode left join st on st.id = u.status order by {1};;", tablename, sortField))
            sb.Append("select cmmf,materialdesc,cmmf || '-' || materialdesc as description from cmmf order by cmmf;")
            sb.Append("select vendorcode,vendorname,vendorcode || '-' || vendorname as description from vendor" &
                      " where vendorcode >= 10000000 and vendorcode <= 90000000 order by vendorname;")
            sb.Append("select pd.ivalue as id,pd.cvalue as statusname from quality.paramdt pd" &
                      " left join quality.paramhd ph on ph.paramhdid = pd.paramhdid" &
                      " where ph.paramname = 'HighRiskStatus'")
            sqlstr = sb.ToString

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
            Return "quality.highrisk"
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
            sqlstr = "quality.sp_deletehighrisk"
            dataadapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "quality.sp_inserthighrisk"
            dataadapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sbu").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "remarks").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "qe").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "stresult").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "ndresult").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "rdresult").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifiedby").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput

            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "quality.sp_updatehighrisk"
            dataadapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sbu").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "remarks").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "qe").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "stresult").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "ndresult").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "rdresult").SourceVersion = DataRowVersion.Current
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
