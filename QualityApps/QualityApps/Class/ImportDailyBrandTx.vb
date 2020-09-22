Imports System.Text
Public Class ImportDailyBrandTx
    Inherits BaseImport
    Implements IImport
    Private MyForm As FormImportDailyBrandExtraction
    Private TXSB As StringBuilder
    Private TXHistorySB As StringBuilder
    Private myAdapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance
    Private errSB As New StringBuilder
    Private period As Integer
    Dim DS As DataSet
    Public Function getLatestPeriod() As Integer
        Dim sqlstr = "select coalesce(period,0) from quality.dailybrandtx order by period desc limit 1"
        Dim result As Integer
        myAdapter.ExecuteScalar(sqlstr, recordAffected:=result)
        Return result
    End Function

    Public Property errorMessage As String
        Get
            Return errsb.tostring
        End Get
        Set(value As String)
            errSB.Append(value)
        End Set
    End Property


    Public Sub New(ByVal parent As FormImportDailyBrandExtraction)
        Me.MyForm = parent
        period = parent.YearWeek
    End Sub

    Public Function CopyTx() As Boolean Implements IImport.CopyTx
        MyForm.ProgressReport(1, "Copy Records..")
        Dim myret As Boolean = True
        If TXSB.Length > 0 Then
            Dim Sqlstr As String = String.Empty

            Dim message As String = String.Empty
            If TXSB.Length > 0 Then
                'Sqlstr = "delete from quality.dailybrandtx;select setval('quality.dailybrandtx_id_seq',1,false);"
                'Dim ra As Long
                ''If Not myAdapter.ExecuteNonQuery(Sqlstr, recordAffected:=ra, message:=message) Then
                ''    MyForm.ProgressReport(1, message)
                ''    Return False
                ''End If
                'If Not myAdapter.ExecuteNonQuery(Sqlstr, message:=message, recordAffected:=ra) Then
                '    MyForm.ProgressReport(1, message)
                '    Return False
                'End If
                Sqlstr = "begin;set statement_timeout to 0;end;delete from quality.dailybrandtx;select setval('quality.dailybrandtx_id_seq',1,false);copy quality.dailybrandtx(porg,cocd,plant,cc,purchdoc,item,seqn,insplot,stage,inspector,code,vendor,vendorname,material,ntsg,materialdesc," &
                         "custpono,sbu,soldtoparty,soldtopartyname,city,ccetd,qty,qtyoun,samplesize,uom,poqty,poqtyoun,startdate, enddate,reference,brand,period )  from stdin with null as 'Null';"
                message = myAdapter.copy(Sqlstr, TXSB.ToString, myret)
                If Not myret Then
                    errorMessage = message
                    MyForm.ProgressReport(1, message)
                End If
            End If
            If TXHistorySB.Length > 0 Then
                Sqlstr = "copy quality.dailybrandtx_audit(purchdoc,item,seqn,ccetd,qty,period)  from stdin with null as 'Null';"
                message = myAdapter.copy(Sqlstr, TXHistorySB.ToString, myret)
                If Not myret Then
                    errorMessage = message
                    MyForm.ProgressReport(1, message)
                End If
            End If
        End If
        Return myret
    End Function

    Public Function RunImport(filename As String) As Boolean Implements IImport.RunImport
        Dim myret As Boolean = True
        DS = New DataSet
        'Get Existing History
        Dim sqlstr = "select purchdoc,item,seqn,ccetd,qty from quality.dailybrandtx_audit;"
        If myAdapter.GetDataset(sqlstr, DS) Then
            DS.Tables(0).TableName = "History"
            Dim pk1(4) As DataColumn
            pk1(0) = DS.Tables(0).Columns("purchdoc")
            pk1(1) = DS.Tables(0).Columns("item")
            pk1(2) = DS.Tables(0).Columns("seqn")
            pk1(3) = DS.Tables(0).Columns("ccetd")
            pk1(4) = DS.Tables(0).Columns("qty")
            DS.Tables(0).PrimaryKey = pk1
        End If


        Dim myList As New List(Of String())
        Dim myrecord() As String
        TXSB = New StringBuilder
        TXHistorySB = New StringBuilder
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
            If myList(1)(4) <> "PO Confirmation - Brand & SBU Filter" Then
                errorMessage = "Wrong Template."
                Return False
            End If
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
                    validint(value(27)) & vbTab & validStr(value(28)) & vbTab & validint(value(29)) & vbTab & validStr(value(30)) & vbTab & validSAPDate(value(31)) & vbTab & validSAPDate(value(32)) & vbTab & validStr(value(33)) & vbTab & validStr(value(34)) & vbTab &
                    period & vbCrLf)
            Else
                TXSB.Append(validStr(value(2)) & vbTab & validStr(value(3)) & vbTab & validStr(value(5)) & vbTab & validStr(value(6)) & vbTab & validLong(value(7)) & vbTab & validint(value(8)) & vbTab &
                    validint(value(9)) & vbTab & validLong(value(10)) & vbTab & validint(value(11)) & vbTab & validStr(value(12)) & vbTab & validStr(value(13)) & vbTab & validLong(value(14)) & vbTab &
                    validStr(value(15)) & vbTab & validLong(value(16)) & vbTab & validint(value(17)) & vbTab & validStr(value(18)) & vbTab & validStr(value(19)) & vbTab & validStr(value(20)) & vbTab &
                    validLong(value(21)) & vbTab & validStr(value(22)) & vbTab & validStr(value(23)) & vbTab & validSAPDate(value(24)) & vbTab & validint(value(25)) & vbTab & validStr(value(26)) & vbTab &
                    validint(value(27)) & vbTab & validStr(value(28)) & vbTab & validint(value(29)) & vbTab & validStr(value(30)) & vbTab & validSAPDate(value(31)) & vbTab & validSAPDate(value(32)) & vbTab & "Null" & vbTab &
                    period & vbCrLf)

            End If
            'Create History

            'Table Header
            Dim mykey1(4) As Object
            mykey1(0) = value(7) 'POHD
            mykey1(1) = value(8) 'Item
            mykey1(2) = value(9) 'Seq
            mykey1(3) = CDate(validSAPDate(value(24)).Replace("'", "")) 'ccetd
            mykey1(4) = validint(value(25)) 'qty
            Dim myresult = DS.Tables(0).Rows.Find(mykey1)
            If IsNothing(myresult) Then
                Dim dr = DS.Tables(0).NewRow
                dr.Item(0) = value(7)
                dr.Item(1) = value(8)
                dr.Item(2) = value(9)
                dr.Item(3) = CDate(validSAPDate(value(24)).Replace("'", ""))
                dr.Item(4) = validint(value(25))
                DS.Tables(0).Rows.Add(dr)

                TXHistorySB.Append(value(7) & vbTab &
                                   value(8) & vbTab &
                                   value(9) & vbTab &
                                   validSAPDate(value(24)) & vbTab &
                                   validint(value(25)) & vbTab &
                                   period & vbCrLf)
            End If


        Catch ex As Exception
            errorMessage = ex.Message
            'MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class
