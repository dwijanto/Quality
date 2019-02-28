Imports System.Text
Public Class ImportOTM
    Inherits BaseImport
    Implements IImport
    Private MyForm As FormImportOTM
    Private TXSB As StringBuilder
    Private myAdapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance

    Public Sub New(ByVal parent As FormImportOTM)
        Me.MyForm = parent
    End Sub

    Public Function CopyTx() As Boolean Implements IImport.CopyTx
        MyForm.ProgressReport(1, "Copy Records..")
        Dim myret As Boolean = True
        If TXSB.Length > 0 Then
            Dim Sqlstr As String = String.Empty

            Dim message As String = String.Empty
            If TXSB.Length > 0 Then
                Sqlstr = "delete from quality.otmtx;"
                Dim ra As Long
                If Not myAdapter.ExecuteNonQuery(Sqlstr, message:=message, recordAffected:=ra) Then
                    MyForm.ProgressReport(1, message)
                    Return False
                End If
                Sqlstr = "begin;set statement_timeout to 0;end;copy quality.otmtx(id,cpo,purchdoc,item,seqn)  from stdin with null as 'Null';"
                message = myAdapter.copy(Sqlstr, TXSB.ToString, myret)
                If Not myret Then
                    MyForm.ProgressReport(1, message)
                End If
            End If
        End If
        Return myret
    End Function

    Public Function RunImport(filename As String) As Boolean Implements IImport.RunImport
        Dim myret As Boolean = True

        Dim myList As New List(Of String())
        Dim myrecord() As String
        TXSB = New StringBuilder
        'Read Text File
        Using objTFParser = New FileIO.TextFieldParser(filename, Encoding.GetEncoding(1252)) '1252 = ANSI
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0
                MyForm.ProgressReport(1, "Read Data..")

                Do Until .EndOfData
                    myrecord = .ReadFields
                    myList.Add(myrecord)
                Loop
            End With
            MyForm.ProgressReport(1, "Build Record..")

            For i = 1 To myList.Count - 1
                CreateRecord(myList(i))
            Next
        End Using

        'Copy TX
        If Not CopyTx() Then
            Return False
        End If

        Return myret
    End Function

    Private Sub CreateRecord(value As String())
        If value(2) = "" Then
            Exit Sub
        End If

        'id character varying NOT NULL,
        'cpo character varying,
        'purchdoc bigint,
        'item integer,
        'seqn integer,
        Dim po = value(2).Split("-")
        TXSB.Append(validStr(value(0)) & vbTab & validStr(value(1)) & vbTab & validLong(po(0)) & vbTab & validint(po(1)) & vbTab & validint(value(3)) & vbCrLf)

    End Sub
End Class
