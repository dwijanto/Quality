﻿Imports Npgsql
Imports System.Text
Public Class ParamDTLAdapter
    'Inherits BaseAdapter
    Public errormessage As String = String.Empty
    Dim DbAdapter1 As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance
    Public DS As DataSet
    Public BS As BindingSource

    Public Function GetParamDTLBS(ByVal paramname) As BindingSource
        Dim sb As New StringBuilder
        sb.Append(String.Format("select pd.* from quality.paramdt pd left join quality.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :paramhdname order by pd.ivalue;"))
        Dim sqlstr = sb.ToString
        DS = New DataSet
        BS = New BindingSource
        Dim myparam(0) As NpgsqlParameter
        myparam(0) = New NpgsqlParameter("paramhdname", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "", ParameterDirection.Input, True, 0, 0, DataRowVersion.Default, paramname)

        If DbAdapter1.GetDataset(sqlstr, DS, myparam) Then
            Dim pk(0) As DataColumn
            BS.DataSource = DS.Tables(0)
        End If
        Return BS
    End Function

    Public Function LoadData(ByVal paramname)
        'If not avail ParamName then Create first
        Dim myparam(0) As NpgsqlParameter
        myparam(0) = New NpgsqlParameter("iparamhdname", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "", ParameterDirection.Input, True, 0, 0, DataRowVersion.Default, paramname)
        DbAdapter1.RunStoreProcedure("quality.validateparamheadername", myparam)

        Dim sb As New StringBuilder
        Dim myret As Boolean = False
        sb.Append(String.Format("select pd.* from quality.paramdt pd left join quality.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :paramhdname;"))
        Dim sqlstr = sb.ToString
        DS = New DataSet
        BS = New BindingSource
        Dim myparam1(0) As NpgsqlParameter
        myparam1(0) = New NpgsqlParameter("paramhdname", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "", ParameterDirection.Input, True, 0, 0, DataRowVersion.Default, paramname)

        If DbAdapter1.GetDataset(sqlstr, DS, myparam1) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("paramdtid")
            DS.Tables(0).PrimaryKey = pk
            BS.DataSource = DS.Tables(0)
            DS.Tables(0).Columns("paramdtid").AutoIncrement = True
            DS.Tables(0).Columns("paramdtid").AutoIncrementSeed = -1
            DS.Tables(0).Columns("paramdtid").AutoIncrementStep = -1
            myret = True
        End If
        Return myret
    End Function

    Public Function save() As Boolean
        Dim myret As Boolean = False
        BS.EndEdit()

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

    Public Function Save(ByVal mye As ContentBaseEventArgs) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = DbAdapter1.getDbDataAdapter
        Dim myret As Boolean = False
        AddHandler dataadapter.RowUpdated, AddressOf DbAdapter1.onRowInsertUpdate
        Dim mytransaction As Npgsql.NpgsqlTransaction
        Using conn As Object = DbAdapter1.getConnection
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

            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function

    Function GetParamHDId(ByVal ParamHDName As String) As Object
        Dim sb As New StringBuilder
        Dim myret As Long
        sb.Append(String.Format("select paramhdid from quality.paramhd ph where ph.paramname = :paramhdname;"))
        Dim sqlstr = sb.ToString

        Dim myparam(0) As NpgsqlParameter
        myparam(0) = New NpgsqlParameter("paramhdname", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "", ParameterDirection.Input, True, 0, 0, DataRowVersion.Default, ParamHDName)

        If Not DbAdapter1.ExecuteScalar(sqlstr, myparam, message:=errormessage, recordAffected:=myret) Then
            Err.Raise(1500, Description:=errormessage)
        End If
        Return myret
    End Function
End Class
