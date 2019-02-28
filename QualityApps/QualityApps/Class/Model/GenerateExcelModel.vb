Imports Npgsql
Imports System.Text

Public Class GenerateExcelModel
    Dim myAdapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance
    Dim myParam As ParamAdapter = ParamAdapter.getInstance
    Dim DateSpan As Integer = myParam.GetCCETDDateSpan

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

    Public Function LoadData(ByVal DS As DataSet) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myAdapter.getDbDataAdapter
        Dim SBException As New StringBuilder

        Dim VendorExceptionlist = myParam.GetVendorExceptionList
        If VendorExceptionlist.Length > 0 Then
            SBException.Append(String.Format("and (tx.vendor not in ({0}))", VendorExceptionlist))
        End If

        Dim SBUExceptionList = myParam.GetSBUExceptionList
        If SBUExceptionList.Length > 0 Then
            SBException.Append(String.Format(" and (tx.sbu not in ({0}))", SBUExceptionList))
        End If
        Dim myret As Boolean = False
        Using conn As Object = myAdapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select distinct true::boolean as selected,vendor,v.vendorname,null::text as message,null::text as hwnd,false::boolean as status ,max(ccetd) over (partition by vendor) as maxdate " &
                                       " from {0} tx" &
                                       " inner join quality.vendor v on v.vendorcode = tx.vendor" &
                                       " where (tx.ccetd >= current_date - {3}) and  (code isnull or code in ('UF','DF','PP')) {2} order by vendor", TableName, SortField, SBException.ToString, DateSpan)
            dataadapter.SelectCommand = myAdapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function
End Class
