Imports Npgsql
Imports System.Text
Public Class SendEmailModel
    Dim myAdapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance
    Dim myParam As ParamAdapter = ParamAdapter.getInstance
    Dim Startday As Integer = myParam.GetStartDate
    Dim Endday As Integer = myParam.GetEndDate
    Public LastSendDate As Date = myParam.GetLastSendDate
    Public startdate As Date
    Public enddate As Date
    Dim _InternalSqlstr As String

    Public ReadOnly Property TableName As String
        Get
            Return "quality.dailytx"
        End Get
    End Property

    Public ReadOnly Property SortField As String
        Get
            Return "vendorcode"
        End Get
    End Property

    Public ReadOnly Property InternalSqlStr As String
        Get
            Return _InternalSqlstr
        End Get
    End Property

    Public Sub New()
        'startdate = Date.Today.AddDays(Startday)
        'enddate = Date.Today.AddDays(Endday)

        'Select Case CInt(Date.Today.DayOfWeek)
        '    Case 4
        '        Dim myinc As Integer = 2 + Endday
        '        enddate = Date.Today.AddDays(myinc)
        '    Case 5
        '        Dim myinc1 As Integer = 2 + Startday
        '        Dim myinc2 As Integer = 2 + Endday
        '        startdate = Date.Today.AddDays(myinc1)
        '        enddate = Date.Today.AddDays(myinc2)
        'End Select
    End Sub

    Public Sub New(ByVal Inspectiondate As Date)
        startdate = Inspectiondate
        enddate = Inspectiondate
    End Sub

    Public Function LoadData(ByVal DS As DataSet) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myAdapter.getDbDataAdapter
        Dim SBException As New StringBuilder

        'startdate = Date.Today.AddDays(Startday)
        'enddate = Date.Today.AddDays(Endday)

        'Select Case CInt(Date.Today.DayOfWeek)
        '    Case 4
        '        Dim myinc As Integer = 2 + Endday
        '        enddate = Date.Today.AddDays(myinc)
        '    Case 5
        '        Dim myinc1 As Integer = 2 + Startday
        '        Dim myinc2 As Integer = 2 + Endday
        '        startdate = Date.Today.AddDays(myinc1)               
        '        enddate = Date.Today.AddDays(myinc2)
        'End Select


        Dim VendorExceptionlist = myParam.GetVendorExceptionList
        If VendorExceptionlist.Length > 0 Then
            SBException.Append(String.Format("and tx.vendor not in ({0})", VendorExceptionlist))
        End If

        Dim SBUExceptionList = myParam.GetSBUExceptionList
        If SBUExceptionList.Length > 0 Then
            SBException.Append(String.Format(" and tx.sbu not in ({0})", SBUExceptionList))
        End If
        Dim myret As Boolean = False
        Using conn As Object = myAdapter.getConnection
            conn.Open()
            'Dim sqlstr = String.Format("(select distinct true::boolean as selected,vendor,v.vendorname,null::text as message,false::boolean as status" &
            '                           " from {0} tx inner join quality.vendor v on v.vendorcode = tx.vendor  where tx.startdate >= '{1:yyyy-MM-dd}' and tx.startdate <= '{2:yyyy-MM-dd}' and tx.startdate <> '{3:yyyy-MM-dd}'{4} order by vendorname )" &
            '                           "  union all (select true::boolean,0,'Internal Email',null::text,false::boolean )", TableName, startdate, enddate, LastSendDate, SBException.ToString)
            Dim sqlstr = String.Format("(select distinct true::boolean as selected,vendor,v.vendorname,null::text as message,false::boolean as status" &
                                       " from {0} tx inner join quality.vendor v on v.vendorcode = tx.vendor  where tx.startdate >= '{1:yyyy-MM-dd}' and tx.startdate <= '{2:yyyy-MM-dd}' {3} order by vendorname )" &
                                       "  union all (select true::boolean,0,'Internal Email',null::text,false::boolean )", TableName, startdate, enddate, SBException.ToString)

            dataadapter.SelectCommand = myAdapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function

    Public Function GetContent(ByVal bs As BindingSource, ByVal VendorCode As Long, ByVal startdate As Date, ByVal enddate As Date) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myAdapter.getDbDataAdapter
        Dim SBException As New StringBuilder
        Dim myret As Boolean = False
        Dim ds As New DataSet
        Using conn As Object = myAdapter.getConnection
            conn.Open()
            Dim vendorCriteria As String = String.Empty
            If VendorCode <> 0 Then
                vendorCriteria = String.Format("vendor = {0} and ", VendorCode)
            End If
            'Dim sqlstr = String.Format("select tx.* " &
            'Dim sqlstr = String.Format("select id,porg,cocd,plant,cc,purchdoc,item,seqn,insplot,stage,inspector,code,vendor,vendorname,material,ntsg,materialdesc,custpono,sbu,soldtoparty,soldtopartyname,city,ccetd,qty,qtyoun,quality.changesamplesize(inspector,samplesize::integer) as samplesize,uom,poqty,poqtyoun,startdate,enddate" &
            '                           " from {0} tx  where {1} tx.startdate >= '{2:yyyy-MM-dd}' and tx.startdate <= '{3:yyyy-MM-dd}' and tx.startdate<>{4:yyyy-MM-dd} order by vendorname,ccetd,material", TableName, vendorCriteria, Me.startdate, Me.enddate, Me.LastSendDate)
            'Dim sqlstr = String.Format("select id,porg,cocd,plant,cc,purchdoc,item,seqn,insplot,stage,inspector,code,vendor,vendorname,material,ntsg,materialdesc,custpono,sbu,soldtoparty,soldtopartyname,city,ccetd,qty,qtyoun,quality.changesamplesize(inspector,samplesize::integer) as samplesize,uom,poqty,poqtyoun,startdate,enddate,quality.getproductionenddate(tx.purchdoc,tx.item,tx.seqn,qty)  as productionenddate" &
            '                           " from {0} tx  where {1} tx.startdate >= '{2:yyyy-MM-dd}' and tx.startdate <= '{3:yyyy-MM-dd}' order by vendorname,ccetd,material", TableName, vendorCriteria, Me.startdate, Me.enddate)
            'Dim sqlstr = String.Format("select id,porg,cocd,plant,cc,purchdoc,item,seqn,insplot,stage,inspector,code,vendor,vendorname,material,ntsg,materialdesc,custpono,sbu,soldtoparty,soldtopartyname,city,ccetd,qty,qtyoun,quality.changesamplesize(inspector,samplesize::integer) as samplesize,uom,poqty,poqtyoun,startdate,enddate,quality.getproductionenddate(tx.purchdoc,tx.item,tx.seqn,qty)  as productionenddate" &
            '                           " from {0} tx  where {1} (code isnull or code in ('UF','DF','PP')) and tx.startdate >= '{2:yyyy-MM-dd}' and tx.startdate <= '{3:yyyy-MM-dd}' order by vendorname,ccetd,material", TableName, vendorCriteria, Me.startdate, Me.enddate)
            Dim sqlstr = String.Format("select id,porg,cocd,plant,cc,purchdoc,item,seqn,insplot,stage,inspector,code,vendor,vendorname,material,ntsg,materialdesc,custpono,sbu,soldtoparty,soldtopartyname,city,ccetd,qty,qtyoun,quality.changesamplesize(inspector,samplesize::integer) as samplesize,uom,poqty,poqtyoun,startdate,enddate,quality.getproductionenddate(tx.purchdoc,tx.item,tx.seqn,qty)  as productionenddate" &
                                       " from {0} tx  where {1} tx.startdate >= '{2:yyyy-MM-dd}' and tx.startdate <= '{3:yyyy-MM-dd}' order by vendorname,ccetd,material", TableName, vendorCriteria, Me.startdate, Me.enddate)
            dataadapter.SelectCommand = myAdapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(ds, TableName)
            bs.DataSource = ds.Tables(0)
            myret = True
        End Using
        Return myret
    End Function

    Public Function GetContent(ByVal bs As BindingSource, ByVal VendorCode As Long) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myAdapter.getDbDataAdapter
        Dim SBException As New StringBuilder
        Dim myret As Boolean = False
        Dim ds As New DataSet
        Using conn As Object = myAdapter.getConnection
            conn.Open()
            Dim vendorCriteria As String = String.Empty
            If VendorCode <> 0 Then
                vendorCriteria = String.Format("vendor = {0} and ", VendorCode)
            End If
            'Dim sqlstr = String.Format("select tx.* " &
            'Dim sqlstr = String.Format("select id,porg,cocd,plant,cc,purchdoc,item,seqn,insplot,stage,inspector,code,vendor,vendorname,material,ntsg,materialdesc,custpono,sbu,soldtoparty,soldtopartyname,city,ccetd,qty,qtyoun,quality.changesamplesize(inspector,samplesize::integer) as samplesize,uom,poqty,poqtyoun,startdate,enddate" &
            '                           " from {0} tx  where {1} tx.startdate >= '{2:yyyy-MM-dd}' and tx.startdate <= '{3:yyyy-MM-dd}' and tx.startdate >= '{4:yyyy-MM-dd}'order by tx.startdate,inspector,vendorname,material", TableName, vendorCriteria, Me.startdate, Me.enddate, Me.LastSendDate)
            'Dim sqlstr = String.Format("select id,porg,cocd,plant,cc,purchdoc,item,seqn,insplot,stage,inspector,code,vendor,vendorname,material,ntsg,materialdesc,custpono,sbu,soldtoparty,soldtopartyname,city,ccetd,qty,qtyoun,quality.changesamplesize(inspector,samplesize::integer) as samplesize,uom,poqty,poqtyoun,startdate,enddate,quality.getproductionenddate(tx.purchdoc,tx.item,tx.seqn,qty)  as productionenddate" &
            '                           " from {0} tx  where {1} tx.startdate >= '{2:yyyy-MM-dd}' and tx.startdate <= '{3:yyyy-MM-dd}' order by tx.startdate,inspector,vendorname,material", TableName, vendorCriteria, Me.startdate, Me.enddate)
            'Dim sqlstr = String.Format("select id,porg,cocd,plant,cc,purchdoc,item,seqn,insplot,stage,inspector,code,vendor,vendorname,material,ntsg,materialdesc,custpono,sbu,soldtoparty,soldtopartyname,city,ccetd,qty,qtyoun,quality.changesamplesize(inspector,samplesize::integer) as samplesize,uom,poqty,poqtyoun,startdate,enddate,quality.getproductionenddate(tx.purchdoc,tx.item,tx.seqn,qty)  as productionenddate" &
            '                           " from {0} tx  where {1} (code isnull or code in ('UF','DF','PP')) and tx.startdate >= '{2:yyyy-MM-dd}' and tx.startdate <= '{3:yyyy-MM-dd}' order by tx.startdate,inspector,vendorname,material", TableName, vendorCriteria, Me.startdate, Me.enddate)

            Dim sqlstr = String.Format("select id,porg,cocd,plant,cc,purchdoc,item,seqn,insplot,stage,inspector,code,vendor,vendorname,material,ntsg,materialdesc,custpono,sbu,soldtoparty,soldtopartyname,city,ccetd,qty,qtyoun,quality.changesamplesize(inspector,samplesize::integer) as samplesize,uom,poqty,poqtyoun,startdate,enddate,quality.getproductionenddate(tx.purchdoc,tx.item,tx.seqn,qty)  as productionenddate" &
                                       " from {0} tx  where {1} tx.startdate >= '{2:yyyy-MM-dd}' and tx.startdate <= '{3:yyyy-MM-dd}' order by tx.startdate,inspector,vendorname,material", TableName, vendorCriteria, Me.startdate, Me.enddate)

            _InternalSqlstr = sqlstr

            dataadapter.SelectCommand = myAdapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(ds, TableName)
            bs.DataSource = ds.Tables(0)
            myret = True
        End Using
        Return myret
    End Function
End Class
