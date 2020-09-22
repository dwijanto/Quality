Imports Npgsql
Imports System.Text
Public Class ConsolidationModel
    Dim myAdapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance
    Dim myParam As ParamAdapter = ParamAdapter.getInstance
    Dim datespan As Integer = myParam.GetCCETDDateSpan()

    Private _sqlstr As String
    Private _sqlstrExcel As String

    Public Property BSHistory As BindingSource
    Public Property DSHistory As DataSet

    'Private _po As Long
    'Private _item As Integer
    'Private _seq As Integer
    'Private _qty As Integer


    Public Property po As String
    Public Property item As String
    Public Property seq As String
    Public Property source As String
    Public Property status As String
    Public Property inspectorname As String
    Public Property inspectiondate As String
    
    Public Property sqlstr As String
        Get
            Return _sqlstr
        End Get
        Set(value As String)
            _sqlstr = value
        End Set
    End Property
    Public Property sqlstrExcel As String
        Get
            Return _sqlstrExcel
        End Get
        Set(value As String)
            _sqlstrExcel = value
        End Set
    End Property


    Public ErrorMessage As String = String.Empty

    Public Sub New()
        '2 DataTable 1-> Join Consolidation Data from 4 sources, 2 -> Consolidation assignment

        'Dim sb As New StringBuilder
        'sb.Append("with data as (")
        'sb.Append("(select period::text as ""Period"" , purchdoc::text as ""Purch Doc"",item::text as ""Item"",seqn::text as ""SeqN"",material::text as ""Material"",materialdesc as ""Material Description"",vendorname as ""Vendor Name"",ccetd as ""CCETD"",qty as ""QTY"",brand as ""Brand"",'WOR'::text as ""Source"" from quality.dailybrandtx)")
        'sb.Append(" union all ")
        'sb.Append("(select dt.period::text, hd.orderno,dt.orderpos,dt.subpos,dt.articleno,dt.materialdesc,hd.supplier,reqdeldate,orderqty,null::text as brand,'Panex'::text as source  from quality.cwpxhd hd left join quality.cwpxdt dt on dt.cwpxhdid = hd.id)")
        'sb.Append(" union all ")
        'sb.Append("(select dt.period::text,hd.serno as orderno,dt.regno,dt.gg,dt.regno,dt.materialdesc,hd.vendorname,cetd,qty,null::text as brand,'Czech'::text as source from quality.cwczhd hd left join quality.cwczdt dt on dt.cwczhdid = hd.id order by serno,regno)")
        'sb.Append(" union all ")
        'sb.Append("(select dt.period::text, hd.pono as orderno,1::text as sebasiapono,seq::text,dt.itemno,dt.itemdescription,hd.supplier,coalesce(cetd,fcetd) as etd,dt.orderqty,null::text as brand,'Shanghai' as source from quality.cwshhd hd left join quality.cwshdt dt on dt.hdid = hd.id order by orderno,itemno)")
        'sb.Append(")  select d.*,c.status as ""Inspection Type"",c.inspectorname as ""Inspector Name"",c.inspectiondate as ""Inspection Date"" from data d" &
        '          " left join quality.consolidationtx c on c.purchdoc = d.""Purch Doc"" and c.item = d.""Item"" and c.seq = d.""SeqN"" and c.source = d.""Source"";")
        '' _sqlstr = sb.ToString
        '_sqlstrExcel = sb.ToString '_sqlstr
        'sb.Append("select * from quality.consolidationtx;")
        '_sqlstr = sb.ToString

    End Sub

    Private Sub getSQLSTR()
        Dim sb As New StringBuilder
        sb.Append("with data as (")
        sb.Append("(select period::text as ""Period"" , purchdoc::text as ""Purch Doc"",item::text as ""Item"",seqn::text as ""SeqN"",material::text as ""Material"",materialdesc as ""Material Description"",vendorname as ""Vendor Name"",ccetd as ""CCETD"",qty as ""QTY"",brand as ""Brand"",'WOR'::text as ""Source"" from quality.dailybrandtx)")
        sb.Append(" union all ")
        sb.Append("(select dt.period::text, hd.orderno,dt.orderpos,dt.subpos,dt.articleno,dt.materialdesc,hd.supplier,reqdeldate,orderqty,null::text as brand,'Panex'::text as source  from quality.cwpxhd hd left join quality.cwpxdt dt on dt.cwpxhdid = hd.id)")
        sb.Append(" union all ")
        sb.Append("(select dt.period::text,hd.serno as orderno,dt.regno,dt.gg,dt.regno,dt.materialdesc,hd.vendorname,cetd,qty,null::text as brand,'Czech'::text as source from quality.cwczhd hd left join quality.cwczdt dt on dt.cwczhdid = hd.id order by serno,regno)")
        sb.Append(" union all ")
        sb.Append("(select dt.period::text, hd.pono as orderno,1::text as sebasiapono,seq::text,dt.itemno,dt.itemdescription,hd.supplier,coalesce(cetd,fcetd) as etd,dt.orderqty,null::text as brand,'Shanghai' as source from quality.cwshhd hd left join quality.cwshdt dt on dt.hdid = hd.id order by orderno,itemno)")
        sb.Append(")  select d.*,c.status as ""Inspection Type"",c.inspectorname as ""Inspector Name"",c.inspectiondate as ""Inspection Date"",c.id as cid from data d" &
                  " left join quality.consolidationtx c on c.purchdoc = d.""Purch Doc"" and c.item = d.""Item"" and c.seq = d.""SeqN"" and c.source = d.""Source"";")
        ' _sqlstr = sb.ToString
        _sqlstrExcel = sb.ToString '_sqlstr
        sb.Append("select * from quality.consolidationtx;")
        _sqlstr = sb.ToString        
    End Sub
    Public ReadOnly Property TableName As String
        Get
            Return "Consolidation"
        End Get
    End Property

    Public ReadOnly Property SortField As String
        Get
            Return "vendor"
        End Get
    End Property


    Public Function LoadData(ByVal DS As DataSet) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myAdapter.getDbDataAdapter
        Dim myret As Boolean = False
        getSQLSTR()
        Using conn As Object = myAdapter.getConnection
            conn.Open()
            dataadapter.SelectCommand = myAdapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function


    Public ReadOnly Property FilterField
        Get
            Return "[period] like '%{0}%' or [purch doc] like '%{0}%' or [Vendor Name] like '%{0}%' or [Material] like '%{0}%' or [Material Description] like '%{0}%'  or [Source] like '%{0}%' "
        End Get
    End Property

    Public ReadOnly Property FilterFieldMissing
        Get
            Return "[purchdoc] like '%{0}%'"
        End Get
    End Property

    Public Function GetRemarkHistory(ByVal po As Long, ByVal item As Integer, ByVal seqn As Integer, ByVal qty As Decimal) As DataSet
        '_po = po
        '_item = item
        '_seq = seqn
        '_qty = qty

        'Dim dataadapter As NpgsqlDataAdapter = myAdapter.getDbDataAdapter
        'Dim myret As Boolean = False
        ''Dim ds As New DataSet
        'DSHistory = New DataSet
        'Using conn As Object = myAdapter.getConnection
        '    conn.Open()
        '    'Dim sqlstr = String.Format("select remark,docdate from {0} tx" &
        '    '                           " where purchdoc=:po and item=:item and seqn=:seqn and qty=:qty and not remark isnull order by docdate desc;" &
        '    '                           "with selected as (select id,purchdoc,item,qty,remark,docdate from quality.historytx tx where purchdoc=:po and item=:item and seqn=:seqn and qty=:qty order by docdate desc)," &
        '    '                           " alldata as (select id,purchdoc,item,qty,remark,docdate from quality.historytx tx)," &
        '    '                           " missingdata as (select * from alldata except all select * from selected)" &
        '    '                           " select * from missingdata where purchdoc =:po and item = :item and qty = :qty", "quality.historytx")
        '    'Dim sqlstr = String.Format("select remark,docdate from {0} tx" &
        '    '                           " where purchdoc=:po and item=:item and seqn=:seqn and not remark isnull order by docdate desc;" &
        '    '                           "with selected as (select h.id,h.purchdoc,h.item,h.qty,h.remark,h.docdate from quality.dailytx tx" &
        '    '                           " inner join quality.historytx h on h.purchdoc = tx.purchdoc and h.item = tx.item and h.seqn = tx.seqn" &
        '    '                           " where not h.remark isnull)," &
        '    '                           " alldata as (select id,purchdoc,item,qty,remark,docdate from quality.historytx h where not remark isnull)," &
        '    '                           " missingdata as (select * from alldata except all select * from selected) " &
        '    '                           " select * from missingdata where purchdoc =:po and item = :item;", "quality.historytx")
        '    Dim sqlstr = String.Format("select * from {0} tx" &
        '                               " where purchdoc=:po and item=:item and seqn=:seqn order by docdate desc;" &
        '                               "with selected as (select h.id,h.purchdoc,h.item,h.qty,h.inspdate,h.remark,h.docdate from quality.dailytx tx" &
        '                               " inner join quality.historytx h on h.purchdoc = tx.purchdoc and h.item = tx.item and h.seqn = tx.seqn" &
        '                               " where not (h.remark isnull and h.inspdate isnull))," &
        '                               " alldata as (select id,purchdoc,item,qty,inspdate,remark,docdate from quality.historytx h where not (remark isnull and inspdate isnull))," &
        '                               " missingdata as (select * from alldata except all select * from selected) " &
        '                               " select * from missingdata where purchdoc =:po and item = :item;", "quality.historytx")

        '    dataadapter.SelectCommand = myAdapter.getCommandObject(sqlstr, conn)
        '    dataadapter.SelectCommand.CommandType = CommandType.Text
        '    Dim myParam(3) As NpgsqlParameter
        '    myParam(0) = New NpgsqlParameter("po", po)
        '    myParam(1) = New NpgsqlParameter("item", item)
        '    myParam(2) = New NpgsqlParameter("seqn", seqn)
        '    myParam(3) = New NpgsqlParameter("qty", qty)
        '    dataadapter.SelectCommand.Parameters.AddRange(myParam)
        '    dataadapter.Fill(DSHistory, TableName)
        '    Dim pk(0) As DataColumn
        '    pk(0) = DSHistory.Tables(0).Columns("id")
        '    DSHistory.Tables(0).Columns("id").AutoIncrement = True
        '    DSHistory.Tables(0).Columns("id").AutoIncrementSeed = -1
        '    DSHistory.Tables(0).Columns("id").AutoIncrementStep = -1
        '    DSHistory.Tables(0).PrimaryKey = pk
        '    BSHistory = New BindingSource
        '    BSHistory.DataSource = DSHistory.Tables(0)
        '    myret = True
        'End Using
        'Return DSHistory
        Return Nothing
    End Function

    Public Function UpdateSeqNHistory(ByVal id As Long, ByVal seqn As Integer) As Boolean

        Dim sqlstr = "update quality.historytx set seqn=:seqn where id =:id;"
        Dim myParam(1) As NpgsqlParameter
        myParam(0) = New NpgsqlParameter("seqn", seqn)
        myParam(1) = New NpgsqlParameter("id", id)
        Return myAdapter.ExecuteNonQuery(sqlstr, myParam, ErrorMessage)
    End Function

    Function LoadDataAllMissing(DS As DataSet) As Boolean
        'Dim dataadapter As NpgsqlDataAdapter = myAdapter.getDbDataAdapter
        'Dim myret As Boolean = False
        'Using conn As Object = myAdapter.getConnection
        '    conn.Open()
        '    sqlstr = String.Format("with  selected as (select h.id,h.purchdoc::text,h.item,h.qty,h.remark,h.docdate from quality.dailytx tx" &
        '                           " inner join quality.historytx h on h.purchdoc = tx.purchdoc and h.item = tx.item and h.seqn = tx.seqn" &
        '                           " where not h.remark isnull)," &
        '                           " alldata as (select id,purchdoc::text,item,qty,remark,docdate from quality.historytx h where not remark isnull and purchdoc in (select purchdoc from quality.dailytx))," &
        '                           " missingdata as (select * from alldata except all select * from selected) " &
        '                           " select * from missingdata order by docdate desc,purchdoc")
        '    'sqlstr = String.Format("with  selected as (select h.id,h.purchdoc::text,h.item,h.qty,h.inspdate,h.remark,h.docdate from quality.dailytx tx" &
        '    '                       " inner join quality.historytx h on h.purchdoc = tx.purchdoc and h.item = tx.item and h.seqn = tx.seqn" &
        '    '                       " where not (h.remark isnull and h.inspdate isnull))," &
        '    '                       " alldata as (select id,purchdoc::text,item,qty,inspdate,remark,docdate from quality.historytx h where not (remark isnull and inspdate isnull) )," &
        '    '                       " missingdata as (select * from alldata except all select * from selected) " &
        '    '                       " select * from missingdata order by docdate desc,purchdoc")
        '    dataadapter.SelectCommand = myAdapter.getCommandObject(sqlstr, conn)
        '    dataadapter.SelectCommand.CommandType = CommandType.Text
        '    dataadapter.Fill(DS, TableName)
        '    myret = True
        'End Using
        'Return myret
        Return False
    End Function
    Function LoadDataAllMissingInspDate(DS As DataSet) As Boolean
        'Dim dataadapter As NpgsqlDataAdapter = myAdapter.getDbDataAdapter
        'Dim myret As Boolean = False
        'Using conn As Object = myAdapter.getConnection
        '    conn.Open()
        '    sqlstr = String.Format("with  selected as (select h.id,h.purchdoc::text,h.item,h.qty,h.inspdate,h.docdate from quality.dailytx tx" &
        '                           " inner join quality.historytx h on h.purchdoc = tx.purchdoc and h.item = tx.item and h.seqn = tx.seqn" &
        '                           " where not h.inspdate isnull)," &
        '                           " alldata as (select id,purchdoc::text,item,qty,inspdate,docdate from quality.historytx h where not inspdate isnull and purchdoc in (select purchdoc from quality.dailytx))," &
        '                           " missingdata as (select * from alldata except all select * from selected) " &
        '                           " select * from missingdata order by docdate desc,purchdoc")
        '    'sqlstr = String.Format("with  selected as (select h.id,h.purchdoc::text,h.item,h.qty,h.inspdate,h.remark,h.docdate from quality.dailytx tx" &
        '    '                       " inner join quality.historytx h on h.purchdoc = tx.purchdoc and h.item = tx.item and h.seqn = tx.seqn" &
        '    '                       " where not (h.remark isnull and h.inspdate isnull))," &
        '    '                       " alldata as (select id,purchdoc::text,item,qty,inspdate,remark,docdate from quality.historytx h where not (remark isnull and inspdate isnull) )," &
        '    '                       " missingdata as (select * from alldata except all select * from selected) " &
        '    '                       " select * from missingdata order by docdate desc,purchdoc")
        '    dataadapter.SelectCommand = myAdapter.getCommandObject(sqlstr, conn)
        '    dataadapter.SelectCommand.CommandType = CommandType.Text
        '    dataadapter.Fill(DS, TableName)
        '    myret = True
        'End Using
        'Return myret
        Return False
    End Function

    Public Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myAdapter.getDbDataAdapter
        Dim myret As Boolean = False
        AddHandler dataadapter.RowUpdated, AddressOf myAdapter.onRowInsertUpdate
        Dim mytransaction As Npgsql.NpgsqlTransaction
        Using conn As Object = myAdapter.getConnection

            conn.Open()
            mytransaction = conn.BeginTransaction

            'Update
            Dim sqlstr = "quality.sp_update_consolidationtx"
            dataadapter.UpdateCommand = myAdapter.getCommandObject(sqlstr, conn)

            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "Purch Doc").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "item").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "seqn").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "Inspection Type").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "Inspector Name").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "Inspection Date").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "source").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cid").Direction = ParameterDirection.InputOutput
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "quality.sp_insert_consolidationtxa"
            dataadapter.InsertCommand = myAdapter.getCommandObject(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "purch doc").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "item").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "seqn").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "Inspection Type").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "Inspector Name").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "Inspection Date").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "source").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cid").Direction = ParameterDirection.InputOutput
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "quality.sp_delete_consolidationtx"
            dataadapter.DeleteCommand = myAdapter.getCommandObject(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cid").Direction = ParameterDirection.Input
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

    Public Function GetLatestHistoryRecord() As Date
        Dim myresult As Date
        'Dim dbparams(3) As NpgsqlParameter
        'dbparams(0) = New NpgsqlParameter("purchdoc", _po)
        'dbparams(1) = New NpgsqlParameter("item", _item)
        'dbparams(2) = New NpgsqlParameter("seqn", _seq)
        'dbparams(3) = New NpgsqlParameter("qty", _qty)

        'myresult = myAdapter.RunStoreProcedure("quality.getinspdate", dbparams)
        Return myresult
    End Function

    Public Function GetLatestRemarksHistoryRecord() As String
        Dim myresult As String = String.Empty
        'Dim dbparams(3) As NpgsqlParameter
        'dbparams(0) = New NpgsqlParameter("purchdoc", _po)
        'dbparams(1) = New NpgsqlParameter("item", _item)
        'dbparams(2) = New NpgsqlParameter("seqn", _seq)
        'dbparams(3) = New NpgsqlParameter("qty", _qty)

        'myresult = myAdapter.RunStoreProcedure("quality.getlatestremark", dbparams)
        Return myresult
    End Function
End Class
