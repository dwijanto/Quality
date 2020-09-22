Imports Npgsql
Imports System.IO

Public Class PostgreSQLDBAdapter
    Implements IDisposable

    Public Property UserInfo As UserInfo
    Public Shared myInstance As PostgreSQLDBAdapter
    Private _connectionstring As String
    Private CopyIn1 As NpgsqlCopyIn
    Dim builder As New NpgsqlConnectionStringBuilder()
    Private _userid As String
    Private _password As String
    Private Sub New()

        DbAdapterInitialize()

    End Sub

    Public ReadOnly Property UserId As String
        Get
            Return _userid
        End Get
    End Property
    Public ReadOnly Property Password As String
        Get
            Return _password
        End Get
    End Property
    Public ReadOnly Property HOST As String
        Get
            Return builder.Host
        End Get
    End Property

    Public ReadOnly Property Database As String
        Get
            Return builder.Database
        End Get
    End Property

    Private Sub DbAdapterInitialize()
        '_connectionstring = String.Format(My.Settings.PostgreSQLCon & "{0}", "User=admin;Password=admin")
        _connectionstring = getConnectionString()
    End Sub

    Private Function getConnectionString() As String
        _userid = "admin"
        _password = "admin"
        builder.ConnectionString = My.Settings.PostgreSQLCon 'My.Settings.ConnectionString1
        builder.Add("User Id", _userid)
        builder.Add("password", _password)
        builder.CommandTimeout = 10000

        builder.Timeout = 1024
        Return builder.ConnectionString
    End Function

    Public Property ConnectionString As String
        Get
            Return _connectionstring
        End Get
        Set(value As String)
            _connectionstring = value
        End Set
    End Property

    Public Shared Function getInstance() As PostgreSQLDBAdapter
        If myInstance Is Nothing Then
            myInstance = New PostgreSQLDBAdapter
        End If
        Return myInstance
    End Function

    Public Overloads Function GetDataset(sqlstr As String, ds As DataSet, Optional params As NpgsqlParameter() = Nothing) As Boolean
        Dim DataAdapter As IDbDataAdapter = New NpgsqlDataAdapter
        Dim myret As Boolean = False
        Using conn As New NpgsqlConnection(_connectionstring)
            conn.Open()
            Using cmd As NpgsqlCommand = New NpgsqlCommand()
                cmd.CommandText = sqlstr
                cmd.Connection = conn
                DataAdapter.SelectCommand = cmd
                If Not IsNothing(params) Then
                    cmd.Parameters.AddRange(params)
                End If
                DataAdapter.Fill(ds)
                myret = True
            End Using
        End Using
        Return myret
    End Function


    'Public Overloads Function GetDataset(sqlstr As String, ds As DataSet, Optional params As List(Of IDataParameter) = Nothing) As Boolean
    '    Dim DataAdapter As IDbDataAdapter = New NpgsqlDataAdapter
    '    Dim myret As Boolean = False
    '    Using conn As New NpgsqlConnection(_connectionstring)
    '        conn.Open()
    '        Using cmd As NpgsqlCommand = New NpgsqlCommand()
    '            cmd.CommandText = sqlstr
    '            cmd.Connection = conn
    '            DataAdapter.SelectCommand = cmd
    '            If Not IsNothing(params) Then
    '                For Each param As IDataParameter In params
    '                    cmd.Parameters.Add(param)
    '                Next
    '            End If
    '            DataAdapter.Fill(ds)
    '            myret = True
    '        End Using
    '    End Using
    '    Return myret
    'End Function

    Public Function ExecuteScalar(ByVal sqlstr As String, Optional ByVal params() As NpgsqlParameter = Nothing, Optional ByRef recordAffected As Object = Nothing, Optional ByRef message As String = "") As Boolean
        Dim myret As Boolean = False
        Using conn As New NpgsqlConnection(_connectionstring)
            conn.Open()
            Using cmd As NpgsqlCommand = New NpgsqlCommand()
                cmd.CommandText = sqlstr
                cmd.Connection = conn
                Try
                    'If params.Length > 0 Then
                    '    cmd.Parameters.AddRange(params)
                    'End If
                    If Not IsNothing(params) Then
                        cmd.Parameters.AddRange(params)
                    End If
                    recordAffected = cmd.ExecuteScalar
                    myret = True
                Catch ex As Exception
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myret
    End Function
    'Public Function ExecuteNonQuery(ByVal sqlstr As String, Optional ByVal params As List(Of IDataParameter) = Nothing, Optional ByRef recordAffected As Int64 = 0, Optional ByRef message As String = "") As Boolean
    '    Dim myret As Boolean = False
    '    Using conn As New NpgsqlConnection(_connectionstring)
    '        conn.Open()
    '        Using cmd As NpgsqlCommand = New NpgsqlCommand()
    '            cmd.CommandText = sqlstr
    '            cmd.Connection = conn
    '            Try
    '                If Not IsNothing(params) Then
    '                    For Each param As IDataParameter In params
    '                        cmd.Parameters.Add(param)
    '                    Next
    '                End If
    '                recordAffected = cmd.ExecuteNonQuery
    '                myret = True
    '            Catch ex As Exception
    '                message = ex.Message
    '            End Try
    '        End Using
    '    End Using
    '    Return myret
    'End Function
    Public Function ExecuteNonQuery(ByVal sqlstr As String, Optional ByVal params() As NpgsqlParameter = Nothing, Optional ByRef message As String = "", Optional ByRef recordAffected As Int64 = 0) As Boolean
        Dim myret As Boolean = False
        Using conn As New NpgsqlConnection(_connectionstring)
            conn.Open()
            Using cmd As NpgsqlCommand = New NpgsqlCommand()
                cmd.CommandText = sqlstr
                cmd.Connection = conn
                Try
                    If Not IsNothing(params) Then
                        For Each param As IDataParameter In params
                            cmd.Parameters.Add(param)
                        Next
                    End If
                    recordAffected = cmd.ExecuteNonQuery
                    myret = True
                Catch ex As Exception
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myret
    End Function

    'Public Function ExecuteScalar(ByVal sqlstr As String, Optional ByVal params As List(Of IDataParameter) = Nothing, Optional ByRef recordAffected As Object = Nothing, Optional ByRef message As String = "") As Boolean
    '    Dim myret As Boolean = False
    '    Using conn As New NpgsqlConnection(_connectionstring)
    '        conn.Open()
    '        Using cmd As NpgsqlCommand = New NpgsqlCommand()
    '            cmd.CommandText = sqlstr
    '            cmd.Connection = conn
    '            Try
    '                If Not IsNothing(params) Then
    '                    For Each param As NpgsqlParameter In params
    '                        cmd.Parameters.Add(param)
    '                    Next
    '                End If
    '                recordAffected = cmd.ExecuteScalar
    '                myret = True
    '            Catch ex As Exception
    '                message = ex.Message
    '            End Try
    '        End Using
    '    End Using
    '    Return myret
    'End Function


   

    Public Function getParam(ByVal ParameterName As String,
                             Optional ByVal value As Object = Nothing,
                         Optional ByVal dbType As DbType = Nothing,
                         Optional ByVal direction As ParameterDirection = ParameterDirection.Input,
                         Optional isNullable As Boolean = False,
                         Optional Precision As Byte = 0,
                         Optional scale As Byte = 0,
                         Optional size As Integer = Integer.MaxValue,
                         Optional SourceColumn As String = "",
                         Optional sourceversion As DataRowVersion = DataRowVersion.Current) As NpgsqlParameter
        Dim myparam = New NpgsqlParameter
        With myparam
            .ParameterName = ParameterName
            .Value = value
            .DbType = dbType
            .Direction = direction
            .IsNullable = isNullable
            .Precision = Precision
            .Scale = scale
            .Size = size
            .SourceColumn = SourceColumn
            .SourceVersion = sourceversion
        End With
        Return myparam
    End Function

    'Public Function isAdmin(ByVal userid As String) As Boolean
    '    Dim sqlstr = String.Format("select * from _user where userid = :userid")
    '    Dim dbparams As New List(Of IDataParameter)
    '    dbparams.Add(getParam("userid", userid, DbType.String))
    '    Dim ra As Object = Nothing
    '    Dim message As String = String.Empty
    '    If ExecuteScalar(sqlstr, dbparams, ra, message) Then
    '        Return Not IsNothing(ra)
    '    End If
    '    Return False
    'End Function

    Function IsAdmin(ByVal p1 As String) As Boolean
        Dim myret As Boolean = False
        Dim sqlstr As String = String.Empty
        Dim Command As New NpgsqlCommand
        Try
            Using conn As New NpgsqlConnection(ConnectionString)
                conn.Open()
                Command = New NpgsqlCommand("quality.sp_isadmin", conn)
                Command.CommandType = CommandType.StoredProcedure
                Command.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").Value = p1.ToLower
                Dim result As Object = Command.ExecuteScalar
                If Not IsNothing(result) Then
                    myret = result
                End If
            End Using
        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            Return False
        End Try
        Return myret
    End Function

    Function RunStoreProcedure(ByVal StoreProcedure As String, params() As NpgsqlParameter) As Object
        Dim myret As Object = Nothing
        Dim sqlstr As String = String.Empty
        Dim cmd As New NpgsqlCommand
        Try
            Using conn As New NpgsqlConnection(ConnectionString)
                conn.Open()
                cmd = New NpgsqlCommand(StoreProcedure, conn)
                cmd.CommandType = CommandType.StoredProcedure
                If Not IsNothing(params) Then
                    cmd.Parameters.AddRange(params)
                End If
                Dim result As Object = cmd.ExecuteScalar
                If Not IsDBNull(result) Then
                    myret = result
                End If
            End Using
        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            Return Nothing
        End Try
        Return myret
    End Function

    Public Function copy(ByVal sqlstr As String, ByVal InputString As String, Optional ByRef result As Boolean = False) As String
        result = False
        Dim myReturn As String = ""
        'Convert string to MemoryStream
        'Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.ASCII.GetBytes(InputString.Replace("\", "\\")))
        'Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.Default.GetBytes(InputString))
        Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(InputString))
        Dim buf(9) As Byte
        Dim CopyInStream As Stream = Nothing
        Dim i As Long
        Using conn = New NpgsqlConnection(_connectionstring)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                CopyIn1 = New NpgsqlCopyIn(command, conn)
                Try
                    CopyIn1.Start()
                    CopyInStream = CopyIn1.CopyStream
                    i = MemoryStream1.Read(buf, 0, buf.Length)
                    While i > 0
                        CopyInStream.Write(buf, 0, i)
                        i = MemoryStream1.Read(buf, 0, buf.Length)
                        Application.DoEvents()
                    End While
                    CopyInStream.Close()
                    result = True
                Catch ex As NpgsqlException
                    Try
                        CopyIn1.Cancel("Undo Copy")
                        myReturn = ex.Message & vbCrLf & ex.Detail & vbCrLf & ex.Where
                    Catch ex2 As NpgsqlException
                        If ex2.Message.Contains("Undo Copy") Then
                            myReturn = ex2.Message & ex.Where
                        End If
                    End Try
                End Try

            End Using
        End Using

        Return myReturn
    End Function


#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

    Sub onRowInsertUpdate(sender As Object, e As NpgsqlRowUpdatedEventArgs)
        'Table with autoincrement
        If e.StatementType = StatementType.Insert Or e.StatementType = StatementType.Update Then
            If e.Status <> UpdateStatus.ErrorsOccurred Then
                e.Status = UpdateStatus.SkipCurrentRow
            End If
        End If
    End Sub



    Public Function getConnection() As NpgsqlConnection
        Return New NpgsqlConnection(_connectionstring)
    End Function

    Public Function getDbDataAdapter() As NpgsqlDataAdapter
        Return New NpgsqlDataAdapter
    End Function

    Public Function getCommandObject() As NpgsqlCommand
        Return New NpgsqlCommand
    End Function

    Public Function getCommandObject(ByVal sqlstr As String, ByVal connection As Object) As NpgsqlCommand
        Return New NpgsqlCommand(sqlstr, connection)
    End Function

    Public Class ContentBaseEventArgs
        Inherits EventArgs
        Public Property dataset As DataSet
        Public Property message As String
        Public Property hasChanges As Boolean
        Public Property ra As Integer
        Public Property continueonerror As Boolean

        Public Sub New(ByVal dataset As DataSet, ByRef haschanges As Boolean, ByRef message As String, ByRef recordaffected As Integer, ByVal continueonerror As Boolean)
            Me.dataset = dataset
            Me.message = message
            Me.ra = ra
            Me.continueonerror = continueonerror
        End Sub
    End Class

    Function loglogin(applicationname As String, userid As String, username As String, computername As String, time_stamp As Date)
        Dim result As Object
        Using conn As New NpgsqlConnection(ConnectionString)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertlogonhistory", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = applicationname
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = userid
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = username
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = computername
            result = cmd.ExecuteNonQuery
        End Using
        Return result
    End Function

End Class
