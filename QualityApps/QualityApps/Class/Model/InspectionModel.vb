Imports Npgsql
Imports System.Text

Public Class InspectionModel
    Dim myAdapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance
    Dim myParam As ParamAdapter = ParamAdapter.getInstance
    Dim datespan As Integer = myParam.GetCCETDDateSpan()
    'Public sqlstr As String = String.Empty
    Private _sqlstr As String
    Private _sqlstrExcel As String

    Public Property BSHistory As BindingSource
    Public Property DSHistory As DataSet

    Private _po As Long
    Private _item As Integer
    Private _seq As Integer
    Private _qty As Integer

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
        Dim SBExecption As New StringBuilder
        Dim sb As New StringBuilder

        Dim VendorExceptionlist = myParam.GetVendorExceptionList
        If VendorExceptionlist.Length > 0 Then
            SBExecption.Append(String.Format("and (tx.vendor not in ({0}))", VendorExceptionlist))
        End If

        Dim SBUExceptionList = myParam.GetSBUExceptionList
        If SBUExceptionList.Length > 0 Then
            SBExecption.Append(String.Format(" and (tx.sbu not in ({0}))", SBUExceptionList))
        End If
        '_sqlstr = String.Format("select false::boolean as ""Selected"",purchdoc::character varying as ""Purch.Doc."",item as ""Item"",seqn as ""SeqN"",insplot::character varying as ""Insp. Lot"",inspector as ""Inspector"",code as ""Inspection Result"",tx.vendor::character varying as ""Vendor"",tx.vendorname as ""Vendor Name"",material::character varying as ""Material"", materialdesc as ""Material desc""," &
        '                          " custpono as ""Cust PO No"",sbu as ""SBU"",city as ""City"",ccetd as ""Confirmed ETD"",qty as ""Quantity"", qtyoun as	""OUn""," &
        '                          " quality.changesamplesize(inspector,samplesize::integer) as ""Sample size"",quality.getinspdate(purchdoc,item,seqn,qty) as ""Inspection Date"",quality.getlatestremark(purchdoc,item,seqn,qty) as ""Remarks""," &
        '                          " ntsg ,quality.dow(date_part('dow',quality.getinspdate(purchdoc,item,seqn,qty))::integer),v.location as ""Location"",v.groupnumber::character varying as ""Group"",startdate,enddate ,quality.getproductionenddate(purchdoc,item,seqn,qty) as ""Production End Date"", soldtoparty::character varying as ""Sold To Party"",soldtopartyname as ""Sold To Party Name"",reference from {0} tx " &
        '                          " left join quality.vendor v on v.vendorcode = tx.vendor" &
        '                          " where (ccetd >= (current_date -{2})) {1}  order by purchdoc,item,seqn", TableName, SBExecption.ToString, datespan)
        sb.Append("with f as (select distinct * from (select material,first_value(ccetd) over (partition by material order by ccetd asc) as ccetd from quality.dailytx dt" &
                  " inner join quality.firstcmmf f on f.cmmf = dt.material" &
                  " where material not in (select cmmf from quality.firstcmmftx)" &
                  " order by material,ccetd)foo)" &
                  " insert into quality.firstcmmftx(cmmf,po,poitem,ccetd,description) select distinct f.material,purchdoc,item,f.ccetd,'First CMMF PO' from f" &
                  " left join quality.dailytx dt on dt.material = f.material and dt.ccetd = f.ccetd;")
        _sqlstrExcel = String.Format("select false::boolean as ""Selected"",tx.plant::character varying,purchdoc::character varying as ""Purch.Doc."",item as ""Item"",seqn as ""SeqN"",insplot::character varying as ""Insp. Lot"",inspector as ""Inspector"",code as ""Inspection Result"",tx.vendor::character varying as ""Vendor"",tx.vendorname as ""Vendor Name"",material::character varying as ""Material"", materialdesc as ""Material desc""," &
                                  " custpono as ""Cust PO No"",sbu as ""SBU"",city as ""City"",tx.ccetd as ""Confirmed ETD"",qty as ""Quantity"", qtyoun as	""OUn""," &
                                  " quality.changesamplesize(inspector,samplesize::integer) as ""Sample size"",quality.getinspdate(purchdoc,item,seqn,qty) as ""Inspection Date"",quality.getlatestremark(purchdoc,item,seqn,qty) as ""Remarks""," &
                                  " ntsg ,quality.dow(date_part('dow',quality.getinspdate(purchdoc,item,seqn,qty))::integer),v.location as ""Location"",v.groupnumber::character varying as ""Group"",startdate,enddate ,quality.getproductionenddate(purchdoc,item,seqn,qty) as ""Production End Date"", soldtoparty::character varying as ""Sold To Party"",soldtopartyname as ""Sold To Party Name"",reference,f.description,quality.getrisk(f.description,tx.city,tx.ntsg) as risk from {0} tx " &
                                  " left join quality.vendor v on v.vendorcode = tx.vendor" &
                                  " left join quality.firstcmmftx f on f.cmmf = tx.material and f.po = tx.purchdoc and f.poitem = tx.item and f.ccetd = tx.ccetd " &
                                  " where (tx.ccetd >= (current_date -{2})) {1}  order by purchdoc,item,seqn;", TableName, SBExecption.ToString, datespan)
        'sb.Append(_sqlstrExcel)
        '_sqlstr = sb.ToString
        '_sqlstr = String.Format("select * from quality.inspectionreportwrk({0},'{1}');", datespan, SBExecption.ToString.Replace("'", "''"))
        _sqlstr = String.Format("select * from quality.inspectionreportwrk01({0},'{1}');", datespan, SBExecption.ToString.Replace("'", "''"))

    End Sub
    Public ReadOnly Property TableName As String
        Get
            Return "quality.dailytx"
        End Get
    End Property

    Public ReadOnly Property SortField As String
        Get
            Return "vendor"
        End Get
    End Property

    'Form Inspection using this loaddata
    Public Function LoadData(ByVal DS As DataSet) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myAdapter.getDbDataAdapter
        Dim myret As Boolean = False
        Dim SBExecption As New StringBuilder

        Dim VendorExceptionlist = myParam.GetVendorExceptionList
        If VendorExceptionlist.Length > 0 Then
            SBExecption.Append(String.Format("and (tx.vendor not in ({0}))", VendorExceptionlist))
        End If

        Dim SBUExceptionList = myParam.GetSBUExceptionList
        If SBUExceptionList.Length > 0 Then
            SBExecption.Append(String.Format(" and (tx.sbu not in ({0}))", SBUExceptionList))
        End If

        Using conn As Object = myAdapter.getConnection
            conn.Open()
            'Dim sqlstr = String.Format("select tx.* from {0} tx inner join quality.vendor v on v.vendorcode = tx.vendor" &
            '                           " where (code isnull or code in ('UF','DF','PP')) and ccetd >= (current_date -14) order by vendor", TableName, SortField)
            'sqlstr = String.Format("select false::boolean as ""Selected"",purchdoc::character varying as ""Purch.Doc."",item as ""Item"",seqn as ""SeqN"",insplot::character varying as ""Insp. Lot"",inspector as ""Inspector"",code as ""Inspection Result"",tx.vendor::character varying as ""Vendor"",tx.vendorname as ""Vendor Name"",material::character varying as ""Material"", materialdesc as ""Material desc""," &
            '                           " custpono as ""Cust PO No"",sbu as ""SBU"",city as ""City"",ccetd as ""Confirmed ETD"",qty as ""Quantity"", qtyoun as	""OUn""," &
            '                           " quality.changesamplesize(inspector,samplesize::integer) as ""Sample size"",quality.getinspdate(purchdoc,item,seqn,qty) as ""Inspection Date"",quality.getlatestremark(purchdoc,item,seqn,qty) as ""Remarks""," &
            '                           " ntsg ,quality.dow(date_part('dow',quality.getinspdate(purchdoc,item,seqn,qty))::integer),v.location as ""Location"",v.groupnumber::character varying as ""Group"",startdate,enddate ,quality.getproductionenddate(purchdoc,item,seqn,qty) as ""Production End Date"", soldtoparty::character varying as ""Sold To Party"",soldtopartyname as ""Sold To Party Name"" from {0} tx " &
            '                           " left join quality.vendor v on v.vendorcode = tx.vendor" &
            '                           " where (code isnull or code in ('UF','DF','PP')) and (ccetd >= (current_date -{2})) {1}  order by purchdoc,item,seqn", TableName, SBExecption.ToString, datespan)
            'sqlstr = String.Format("select false::boolean as ""Selected"",purchdoc::character varying as ""Purch.Doc."",item as ""Item"",seqn as ""SeqN"",insplot::character varying as ""Insp. Lot"",inspector as ""Inspector"",code as ""Inspection Result"",tx.vendor::character varying as ""Vendor"",tx.vendorname as ""Vendor Name"",material::character varying as ""Material"", materialdesc as ""Material desc""," &
            '                          " custpono as ""Cust PO No"",sbu as ""SBU"",city as ""City"",ccetd as ""Confirmed ETD"",qty as ""Quantity"", qtyoun as	""OUn""," &
            '                          " quality.changesamplesize(inspector,samplesize::integer) as ""Sample size"",quality.getinspdate(purchdoc,item,seqn,qty) as ""Inspection Date"",quality.getlatestremark(purchdoc,item,seqn,qty) as ""Remarks""," &
            '                          " ntsg ,quality.dow(date_part('dow',quality.getinspdate(purchdoc,item,seqn,qty))::integer),v.location as ""Location"",v.groupnumber::character varying as ""Group"",startdate,enddate ,quality.getproductionenddate(purchdoc,item,seqn,qty) as ""Production End Date"", soldtoparty::character varying as ""Sold To Party"",soldtopartyname as ""Sold To Party Name"" from {0} tx " &
            '                          " left join quality.vendor v on v.vendorcode = tx.vendor" &
            '                          " where (ccetd >= (current_date -{2})) {1}  order by purchdoc,item,seqn", TableName, SBExecption.ToString, datespan)
            'dataadapter.SelectCommand = myAdapter.getCommandObject(sqlstr, conn)

            dataadapter.SelectCommand = myAdapter.getCommandObject(sqlstr, conn)

            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function


    Public ReadOnly Property FilterField
        Get
            Return "[plant] like '%{0}%' or [Purch.Doc.] like '%{0}%' or [Vendor] like '%{0}%' or [Insp. Lot] like '%{0}%' or [Inspector] like '%{0}%' or [Inspection Result] like '%{0}%' or [Insp. Lot] like '%{0}%' or [Vendor Name] like '%{0}%' or [Material] like '%{0}%' or [Material desc] like '%{0}%' " &
                "or [Cust PO No] like '%{0}%' or [SBU] like '%{0}%' or [City] like '%{0}%' or [Remarks] like '%{0}%' or [Location] like '%{0}%' or [Group] like '%{0}%'  or [Sold To Party] like '%{0}%'  or [Sold To Party Name] like '%{0}%'"
        End Get
    End Property

    Public ReadOnly Property FilterFieldMissing
        Get
            Return "[purchdoc] like '%{0}%'"
        End Get
    End Property

    Public Function GetRemarkHistory(ByVal po As Long, ByVal item As Integer, ByVal seqn As Integer, ByVal qty As Decimal) As DataSet
        _po = po
        _item = item
        _seq = seqn
        _qty = qty

        Dim dataadapter As NpgsqlDataAdapter = myAdapter.getDbDataAdapter
        Dim myret As Boolean = False
        'Dim ds As New DataSet
        DSHistory = New DataSet
        Using conn As Object = myAdapter.getConnection
            conn.Open()
            'Dim sqlstr = String.Format("select remark,docdate from {0} tx" &
            '                           " where purchdoc=:po and item=:item and seqn=:seqn and qty=:qty and not remark isnull order by docdate desc;" &
            '                           "with selected as (select id,purchdoc,item,qty,remark,docdate from quality.historytx tx where purchdoc=:po and item=:item and seqn=:seqn and qty=:qty order by docdate desc)," &
            '                           " alldata as (select id,purchdoc,item,qty,remark,docdate from quality.historytx tx)," &
            '                           " missingdata as (select * from alldata except all select * from selected)" &
            '                           " select * from missingdata where purchdoc =:po and item = :item and qty = :qty", "quality.historytx")
            'Dim sqlstr = String.Format("select remark,docdate from {0} tx" &
            '                           " where purchdoc=:po and item=:item and seqn=:seqn and not remark isnull order by docdate desc;" &
            '                           "with selected as (select h.id,h.purchdoc,h.item,h.qty,h.remark,h.docdate from quality.dailytx tx" &
            '                           " inner join quality.historytx h on h.purchdoc = tx.purchdoc and h.item = tx.item and h.seqn = tx.seqn" &
            '                           " where not h.remark isnull)," &
            '                           " alldata as (select id,purchdoc,item,qty,remark,docdate from quality.historytx h where not remark isnull)," &
            '                           " missingdata as (select * from alldata except all select * from selected) " &
            '                           " select * from missingdata where purchdoc =:po and item = :item;", "quality.historytx")
            Dim sqlstr = String.Format("select * from {0} tx" &
                                       " where purchdoc=:po and item=:item and seqn=:seqn order by docdate desc;" &
                                       "with selected as (select h.id,h.purchdoc,h.item,h.qty,h.inspdate,h.remark,h.docdate from quality.dailytx tx" &
                                       " inner join quality.historytx h on h.purchdoc = tx.purchdoc and h.item = tx.item and h.seqn = tx.seqn" &
                                       " where not (h.remark isnull and h.inspdate isnull))," &
                                       " alldata as (select id,purchdoc,item,qty,inspdate,remark,docdate from quality.historytx h where not (remark isnull and inspdate isnull))," &
                                       " missingdata as (select * from alldata except all select * from selected) " &
                                       " select * from missingdata where purchdoc =:po and item = :item;", "quality.historytx")

            dataadapter.SelectCommand = myAdapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            Dim myParam(3) As NpgsqlParameter
            myParam(0) = New NpgsqlParameter("po", po)
            myParam(1) = New NpgsqlParameter("item", item)
            myParam(2) = New NpgsqlParameter("seqn", seqn)
            myParam(3) = New NpgsqlParameter("qty", qty)
            dataadapter.SelectCommand.Parameters.AddRange(myParam)
            dataadapter.Fill(DSHistory, TableName)
            Dim pk(0) As DataColumn
            pk(0) = DSHistory.Tables(0).Columns("id")
            DSHistory.Tables(0).Columns("id").AutoIncrement = True
            DSHistory.Tables(0).Columns("id").AutoIncrementSeed = -1
            DSHistory.Tables(0).Columns("id").AutoIncrementStep = -1
            DSHistory.Tables(0).PrimaryKey = pk
            BSHistory = New BindingSource
            BSHistory.DataSource = DSHistory.Tables(0)
            myret = True
        End Using
        Return DSHistory
    End Function

    Public Function UpdateSeqNHistory(ByVal id As Long, ByVal seqn As Integer) As Boolean

        Dim sqlstr = "update quality.historytx set seqn=:seqn where id =:id;"
        Dim myParam(1) As NpgsqlParameter
        myParam(0) = New NpgsqlParameter("seqn", seqn)
        myParam(1) = New NpgsqlParameter("id", id)
        Return myAdapter.ExecuteNonQuery(sqlstr, myParam, ErrorMessage)
    End Function

    Function LoadDataAllMissing(DS As DataSet) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myAdapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myAdapter.getConnection
            conn.Open()
            sqlstr = String.Format("with  selected as (select h.id,h.purchdoc::text,h.item,h.qty,h.remark,h.docdate from quality.dailytx tx" &
                                   " inner join quality.historytx h on h.purchdoc = tx.purchdoc and h.item = tx.item and h.seqn = tx.seqn" &
                                   " where not h.remark isnull)," &
                                   " alldata as (select id,purchdoc::text,item,qty,remark,docdate from quality.historytx h where not remark isnull and purchdoc in (select purchdoc from quality.dailytx))," &
                                   " missingdata as (select * from alldata except all select * from selected) " &
                                   " select * from missingdata order by docdate desc,purchdoc")
            'sqlstr = String.Format("with  selected as (select h.id,h.purchdoc::text,h.item,h.qty,h.inspdate,h.remark,h.docdate from quality.dailytx tx" &
            '                       " inner join quality.historytx h on h.purchdoc = tx.purchdoc and h.item = tx.item and h.seqn = tx.seqn" &
            '                       " where not (h.remark isnull and h.inspdate isnull))," &
            '                       " alldata as (select id,purchdoc::text,item,qty,inspdate,remark,docdate from quality.historytx h where not (remark isnull and inspdate isnull) )," &
            '                       " missingdata as (select * from alldata except all select * from selected) " &
            '                       " select * from missingdata order by docdate desc,purchdoc")
            dataadapter.SelectCommand = myAdapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function
    Function LoadDataAllMissingInspDate(DS As DataSet) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myAdapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myAdapter.getConnection
            conn.Open()
            sqlstr = String.Format("with  selected as (select h.id,h.purchdoc::text,h.item,h.qty,h.inspdate,h.docdate from quality.dailytx tx" &
                                   " inner join quality.historytx h on h.purchdoc = tx.purchdoc and h.item = tx.item and h.seqn = tx.seqn" &
                                   " where not h.inspdate isnull)," &
                                   " alldata as (select id,purchdoc::text,item,qty,inspdate,docdate from quality.historytx h where not inspdate isnull and purchdoc in (select purchdoc from quality.dailytx))," &
                                   " missingdata as (select * from alldata except all select * from selected) " &
                                   " select * from missingdata order by docdate desc,purchdoc")
            'sqlstr = String.Format("with  selected as (select h.id,h.purchdoc::text,h.item,h.qty,h.inspdate,h.remark,h.docdate from quality.dailytx tx" &
            '                       " inner join quality.historytx h on h.purchdoc = tx.purchdoc and h.item = tx.item and h.seqn = tx.seqn" &
            '                       " where not (h.remark isnull and h.inspdate isnull))," &
            '                       " alldata as (select id,purchdoc::text,item,qty,inspdate,remark,docdate from quality.historytx h where not (remark isnull and inspdate isnull) )," &
            '                       " missingdata as (select * from alldata except all select * from selected) " &
            '                       " select * from missingdata order by docdate desc,purchdoc")
            dataadapter.SelectCommand = myAdapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function

    Public Function savehistory(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myAdapter.getDbDataAdapter
        Dim myret As Boolean = False
        AddHandler dataadapter.RowUpdated, AddressOf myAdapter.onRowInsertUpdate
        Dim mytransaction As Npgsql.NpgsqlTransaction
        Using conn As Object = myAdapter.getConnection

            conn.Open()
            mytransaction = conn.BeginTransaction

            'Update
            Dim sqlstr = "quality.sp_update_historytx"
            dataadapter.UpdateCommand = myAdapter.getCommandObject(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "purchdoc").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "item").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "seqn").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendor").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "ccetd").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "qty").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "inspdate").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remark").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "docdate").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "productionenddate").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "quality.sp_insert_historytx"
            dataadapter.InsertCommand = myAdapter.getCommandObject(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "purchdoc").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "item").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "seqn").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendor").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "ccetd").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "qty").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "inspdate").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remark").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "docdate").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "productionenddate").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "quality.sp_delete_historytx"
            dataadapter.DeleteCommand = myAdapter.getCommandObject(sqlstr, conn)
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

    Public Function GetLatestHistoryRecord() As Date
        Dim myresult As Date
        Dim dbparams(3) As NpgsqlParameter
        dbparams(0) = New NpgsqlParameter("purchdoc", _po)
        dbparams(1) = New NpgsqlParameter("item", _item)
        dbparams(2) = New NpgsqlParameter("seqn", _seq)
        dbparams(3) = New NpgsqlParameter("qty", _qty)

        myresult = myAdapter.RunStoreProcedure("quality.getinspdate", dbparams)
        Return myresult
    End Function

    Public Function GetLatestRemarksHistoryRecord() As String
        Dim myresult As String
        Dim dbparams(3) As NpgsqlParameter
        dbparams(0) = New NpgsqlParameter("purchdoc", _po)
        dbparams(1) = New NpgsqlParameter("item", _item)
        dbparams(2) = New NpgsqlParameter("seqn", _seq)
        dbparams(3) = New NpgsqlParameter("qty", _qty)

        myresult = myAdapter.RunStoreProcedure("quality.getlatestremark", dbparams)
        Return myresult
    End Function
End Class
