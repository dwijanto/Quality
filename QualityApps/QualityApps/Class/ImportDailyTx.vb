Imports System.Text

Public Class ImportDailyTx
    Inherits BaseImport
    Implements IImport
    Private MyForm As FormImportDailyExtraction
    Private TXSB As StringBuilder
    Private myAdapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance

    Public Sub New(ByVal parent As FormImportDailyExtraction)
        Me.MyForm = parent
    End Sub

    Public Function CopyTx() As Boolean Implements IImport.CopyTx
        MyForm.ProgressReport(1, "Copy Records..")
        Dim myret As Boolean = True
        If TXSB.Length > 0 Then
            Dim Sqlstr As String = String.Empty

            Dim message As String = String.Empty
            If TXSB.Length > 0 Then
                Sqlstr = "delete from quality.dailytx;select setval('quality.dailytx_id_seq',1,false);"
                Dim ra As Long
                'If Not myAdapter.ExecuteNonQuery(Sqlstr, recordAffected:=ra, message:=message) Then
                '    MyForm.ProgressReport(1, message)
                '    Return False
                'End If
                If Not myAdapter.ExecuteNonQuery(Sqlstr, message:=message, recordAffected:=ra) Then
                    MyForm.ProgressReport(1, message)
                    Return False
                End If
                Sqlstr = "begin;set statement_timeout to 0;end;copy quality.dailytx(porg,cocd,plant,cc,purchdoc,item,seqn,insplot,stage,inspector,code,vendor,vendorname,material,ntsg,materialdesc," &
                         "custpono,sbu,soldtoparty,soldtopartyname,city,ccetd,qty,qtyoun,samplesize,uom,poqty,poqtyoun,startdate, enddate,reference )  from stdin with null as 'Null';"
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
            For i = 4 To myList.Count - 1
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

        'porg character varying,  cocd character varying,  plant character varying,  cc character varying,  purchdoc bigint,
        'item integer,  seqn integer,  insplot integer,  stage integer,  inspector character varying,  code character varying,
        'vendor bigint,  vendorname character varying,  material bigint,  ntsg integer,  materialdesc character varying,
        'custpono character varying,  sbu character varying,  soldtoparty bigint,  soldtopartyname character varying,city character varying,  ccetd date,
        'qty integer,  qtyoun character varying,  samplesize integer,  uom character varying,  poqty integer,  
        'poqtyoun character varying,  startdate date,  enddate date,
        Try

            If value.Length > 33 Then
                TXSB.Append(validStr(value(2)) & vbTab & validStr(value(3)) & vbTab & validStr(value(5)) & vbTab & validStr(value(6)) & vbTab & validLong(value(7)) & vbTab & validint(value(8)) & vbTab &
                    validint(value(9)) & vbTab & validLong(value(10)) & vbTab & validint(value(11)) & vbTab & validStr(value(12)) & vbTab & validStr(value(13)) & vbTab & validLong(value(14)) & vbTab &
                    validStr(value(15)) & vbTab & validLong(value(16)) & vbTab & validint(value(17)) & vbTab & validStr(value(18)) & vbTab & validStr(value(19)) & vbTab & validStr(value(20)) & vbTab &
                    validLong(value(21)) & vbTab & validStr(value(22)) & vbTab & validStr(value(23)) & vbTab & validSAPDate(value(24)) & vbTab & validint(value(25)) & vbTab & validStr(value(26)) & vbTab &
                    validNumeric(value(27)) & vbTab & validStr(value(28)) & vbTab & validint(value(29)) & vbTab & validStr(value(30)) & vbTab & validSAPDate(value(31)) & vbTab & validSAPDate(value(32)) & vbTab & validStr(value(33)) & vbCrLf)
            Else
                TXSB.Append(validStr(value(2)) & vbTab & validStr(value(3)) & vbTab & validStr(value(5)) & vbTab & validStr(value(6)) & vbTab & validLong(value(7)) & vbTab & validint(value(8)) & vbTab &
                    validint(value(9)) & vbTab & validLong(value(10)) & vbTab & validint(value(11)) & vbTab & validStr(value(12)) & vbTab & validStr(value(13)) & vbTab & validLong(value(14)) & vbTab &
                    validStr(value(15)) & vbTab & validLong(value(16)) & vbTab & validint(value(17)) & vbTab & validStr(value(18)) & vbTab & validStr(value(19)) & vbTab & validStr(value(20)) & vbTab &
                    validLong(value(21)) & vbTab & validStr(value(22)) & vbTab & validStr(value(23)) & vbTab & validSAPDate(value(24)) & vbTab & validint(value(25)) & vbTab & validStr(value(26)) & vbTab &
                    validNumeric(value(27)) & vbTab & validStr(value(28)) & vbTab & validint(value(29)) & vbTab & validStr(value(30)) & vbTab & validSAPDate(value(31)) & vbTab & validSAPDate(value(32)) & vbTab & "Null" & vbCrLf)

            End If
        
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

End Class
