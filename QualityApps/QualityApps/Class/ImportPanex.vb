Imports System.Text
Public Class ImportPanex
    Inherits BaseImport
    Implements IImport
    Private myAdapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance
    Private MyForm As FormImportCookwarePanex
    Private HDSB As StringBuilder
    Private DTSB As StringBuilder
    Private UpdHDSB As StringBuilder
    Private UpdDTSB As StringBuilder
    Private hdseq As Long
    Private dtseq As Long
    Private hdseqlatest As Long
    Private dtseqlatest
    Dim DS As DataSet
    Private period As Integer

    Public Property errMsgSB As New StringBuilder

    Public Sub New(ByVal parent As FormImportCookwarePanex)
        Me.MyForm = parent
        period = parent.YearWeek
    End Sub
    Public Function getLatestPeriod() As Integer
        Dim sqlstr = "select coalesce(period,0) from quality.cwpxdt order by period desc limit 1"
        Dim result As Integer
        myAdapter.ExecuteScalar(sqlstr, recordAffected:=result)
        Return result
    End Function

    Public Function CopyTx() As Boolean Implements IImport.CopyTx
        MyForm.ProgressReport(1, "Copy Records..")
        Dim myret As Boolean = True
        'Header
        If HDSB.Length > 0 Then
            Dim Sqlstr = String.Format("begin;set statement_timeout to 0;end;select setval('quality.cwpxhd_id_seq',{0},false);copy quality.cwpxhd(id,orderno,supplier,purchasinggroup,bu,period)  from stdin with null as 'Null'; ", hdseq)
            Dim Message = myAdapter.copy(Sqlstr, HDSB.ToString, myret)
            If Not myret Then
                errMsgSB.Append(Message & vbCrLf)
                MyForm.ProgressReport(1, Message)
                Return False
            End If
        End If


        'Detail
        If DTSB.Length > 0 Then
            Dim Sqlstr = String.Format("begin;set statement_timeout to 0;end;select setval('quality.cwpxdt_id_seq',{0},false);copy quality.cwpxdt(id,cwpxhdid,orderpos,subpos,articleno,materialdesc,orderqty,reqdeldate,createddate,period)  from stdin with null as 'Null';", dtseq)
            Dim Message = myAdapter.copy(Sqlstr, DTSB.ToString, myret)
            If Not myret Then
                errMsgSB.Append(Message & vbCrLf)
                MyForm.ProgressReport(1, Message)
                Return False
            End If
        End If

        ''UpdateHeader
        'If UpdHDSB.Length > 0 Then
        '    'Parent.ProgressReport(2, "Update CMMF")
        '    Dim sqlstr = "update quality.cwczhd cw set deliverydate= foo.deliverydate::date " &
        '                " from (select * from array_to_set2(Array[" & UpdHDSB.ToString &
        '             "]) as tb (id character varying, deliverydate character varying))foo where cw.id = foo.id::bigint;"

        '    Dim errmsg As String = String.Empty
        '    If Not myAdapter.ExecuteNonQuery(sqlstr, message:=errmsg) Then
        '        errMsgSB.Append(errmsg & vbCrLf)
        '        Return False
        '    End If
        'End If

        'UpdateDT
        If UpdDTSB.Length > 0 Then
            'Parent.ProgressReport(2, "Update CMMF")
            Dim sqlstr = "update quality.cwpxdt cw set orderqty= foo.orderqty::integer,reqdeldate = foo.reqdeldate::date,period=foo.period::integer" &
                        " from (select * from array_to_set7(Array[" & UpdDTSB.ToString &
                     "]) as tb (id character varying, orderqty character varying,reqdeldate character varying,period character varying))foo where cw.id = foo.id::bigint;"

            Dim errmsg As String = String.Empty
            If Not myAdapter.ExecuteNonQuery(sqlstr, message:=errmsg) Then
                errMsgSB.Append(errmsg & vbCrLf)
                Return False
            End If
        End If

        Return myret
    End Function

    Public Function RunImport(filename As String) As Boolean Implements IImport.RunImport
        Dim myret As Boolean = True

        Dim myList As New List(Of String())
        Dim myrecord() As String
        HDSB = New StringBuilder
        DTSB = New StringBuilder
        UpdDTSB = New StringBuilder
        UpdHDSB = New StringBuilder
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
            'Check Template File 
            If myList(0)(0) <> "BU" And myList(0)(1) <> "Supplier" Then
                errMsgSB.Append("Sorry wrong file template.")
                Return False
            End If
            MyForm.ProgressReport(1, "Prepare Data..")
            'Get Existing Data

            If Not PopulateExistingData() Then
                Return False
            End If
            MyForm.ProgressReport(1, "Build Record..")
            Dim createhd As Boolean = False
            Try
                For i = 1 To myList.Count - 1
                    'Header
                    'If dtseq = 2 Then
                    '    Debug.Print("test")
                    'End If
                    Dim mymodel As PanexModel = New PanexModel With {.bu = myList(i)(0),
                                                                     .supplier = myList(i)(1),
                                                                     .articleno = myList(i)(2),
                                                                     .materialdesc = myList(i)(3),
                                                                     .purchasinggroup = myList(i)(4),
                                                                     .orderno = myList(i)(5),
                                                                     .orderpos = myList(i)(6),
                                                                     .subpos = myList(i)(7),
                                                                     .orderqty = myList(i)(8),
                                                                     .reqdeldate = myList(i)(9)
                                                                     }
                    'Table Header
                    Dim mykey1(0) As Object
                    mykey1(0) = mymodel.orderno
                    Dim myresult = DS.Tables(0).Rows.Find(mykey1)

                    If Not IsNothing(myresult) Then
                        'UpdateRecordHD(mymodel, myresult)
                        hdseq = myresult.Item("id")
                    Else
                        hdseqlatest = hdseqlatest + 1
                        hdseq = hdseqlatest
                        CreateRecordHD(mymodel)

                    End If

                    'Table Detail
                    Dim mykey1DT(3) As Object
                    mykey1DT(0) = hdseq
                    mykey1DT(1) = mymodel.orderpos
                    mykey1DT(2) = mymodel.subpos
                    mykey1DT(3) = mymodel.articleno
                    Dim myresultDT = DS.Tables(1).Rows.Find(mykey1DT)
                    If Not IsNothing(myresultDT) Then
                        UpdateRecordDT(mymodel, myresultDT)
                        dtseq = myresultDT.Item("id")
                    Else
                        dtseqlatest = dtseqlatest + 1
                        dtseq = dtseqlatest
                        CreateRecordDT(mymodel)
                    End If



                Next
            Catch ex As Exception
                errMsgSB.Append(ex.Message)
                Return False
            End Try

        End Using

        'Copy TX
        If Not CopyTx() Then
            Return False
        End If

        Return myret
    End Function

    Private Sub CreateRecordHD(myModel As PanexModel)
        'id bigserial NOT NULL,
        'serno character varying,
        'vendorname character varying,
        'country character varying,
        'deliverydate date,
        Dim dr As DataRow = DS.Tables(0).NewRow
        dr.Item("id") = hdseq
        dr.Item("orderno") = myModel.orderno
        dr.Item("supplier") = myModel.supplier
        dr.Item("purchasinggroup") = myModel.purchasinggroup
        dr.Item("bu") = myModel.bu
        DS.Tables(0).Rows.Add(dr)

        HDSB.Append(validStr(hdseq) & vbTab &
                    validStr(myModel.orderno) & vbTab &
                    validStr(myModel.supplier) & vbTab &
                    validStr(myModel.purchasinggroup) & vbTab &
                    validStr(myModel.bu) & vbTab &
                    period & vbCrLf)

    End Sub

    'Private Sub UpdateRecordHD(myModel As PanexModel, ByRef myresult As DataRow)
    '    Dim flagUpdate As Boolean = False
    '    If myModel.deliverydate <> myresult.Item("deliverydate") Then
    '        flagUpdate = True
    '        myresult.Item("deliverydate") = myModel.deliverydate
    '    End If
    '    If flagUpdate Then
    '        If UpdHDSB.Length > 0 Then
    '            UpdHDSB.Append(",")
    '        End If
    '        UpdHDSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying]", myresult.Item("id"), myModel.deliverydate))
    '    End If
    'End Sub

    Private Sub UpdateRecordDT(mymodel As PanexModel, ByRef myresultDT As DataRow)
        Dim flagUpdate As Boolean = False
        If validint(mymodel.orderqty) <> myresultDT.Item("orderqty") Then
            flagUpdate = True
        End If

        If IsDBNull(myresultDT.Item("reqdeldate")) Then
            If mymodel.reqdeldate <> "" Then
                flagUpdate = True
            End If

        Else
            If mymodel.reqdeldate <> myresultDT.Item("reqdeldate") Then
                flagUpdate = True
            End If
        End If



        If flagUpdate Then
            If UpdDTSB.Length > 0 Then
                UpdDTSB.Append(",")
            End If
            UpdDTSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{2}'::character varying,'{3}'::character varying]",
                                         myresultDT.Item("id"), validint(mymodel.orderqty), validDate(mymodel.reqdeldate), period))
        End If
    End Sub

    Private Sub CreateRecordDT(mymodel As PanexModel)
        'id bigserial NOT NULL,
        'cwczhdid bigint,
        'gg character varying,
        'regno character varying,
        'materialdesc character varying,
        'qty integer,
        'collectedqty integer,
        'description character varying,
        'cetd date,
        'ceta date,
        'createddate date,


        Dim dr As DataRow = DS.Tables(1).NewRow
        dr.Item("id") = dtseq
        dr.Item("cwpxhdid") = hdseq
        dr.Item("orderpos") = mymodel.orderpos
        dr.Item("subpos") = mymodel.subpos
        dr.Item("articleno") = mymodel.articleno
        dr.Item("materialdesc") = mymodel.materialdesc
        dr.Item("orderqty") = validint(mymodel.orderqty)
        If mymodel.reqdeldate <> "" Then dr.Item("reqdeldate") = CDate(mymodel.reqdeldate)
        dr.Item("createddate") = Today.Date
        DS.Tables(1).Rows.Add(dr)
        DTSB.Append(validStr(dtseq) & vbTab &
                    validStr(hdseq) & vbTab &
                    validStr(mymodel.orderpos) & vbTab &
                    validStr(mymodel.subpos) & vbTab &
                    validStr(mymodel.articleno) & vbTab &
                    validStr(mymodel.materialdesc) & vbTab &
                    validint(mymodel.orderqty) & vbTab &
                    validDate(mymodel.reqdeldate) & vbTab &
                    validDate(Today.Date.ToString) & vbTab &
                    period & vbCrLf)
    End Sub

    Private Function PopulateExistingData() As Boolean
        Dim sb As New StringBuilder
        Dim myret As Boolean = False
        DS = New DataSet
        sb.Append("select * from quality.cwpxhd;") 'Table 1 : cwczhd
        sb.Append("select * from quality.cwpxdt;") 'Table 2 : cwczdt        
        sb.Append("select id from quality.cwpxhd order by id desc limit 1;")
        sb.Append("select id from quality.cwpxdt order by id desc limit 1;")
        'sb.Append("select setval('quality.cwczhd_id_seq',nextval('quality.cwczhd_id_seq')-1,false);select currval('quality.cwczhd_id_seq');") 'Table 3 : seqhd
        'sb.Append("select setval('quality.cwczdt_id_seq',nextval('quality.cwczdt_id_seq'),false);select currval('quality.cwczdt_id_seq');") 'Table 4 : seqdt
        Try
            myAdapter.GetDataset(sb.ToString, DS)
            Dim pk1(0) As DataColumn
            pk1(0) = DS.Tables(0).Columns("orderno")
            DS.Tables(0).PrimaryKey = pk1

            Dim pk2(3) As DataColumn
            pk2(0) = DS.Tables(1).Columns("cwpxhdid")
            pk2(1) = DS.Tables(1).Columns("orderpos")
            pk2(2) = DS.Tables(1).Columns("subpos")
            pk2(3) = DS.Tables(1).Columns("articleno")
            DS.Tables(1).PrimaryKey = pk2
            'Sequence
            hdseq = 0
            If DS.Tables(2).Rows.Count > 0 Then
                hdseqlatest = DS.Tables(2).Rows(0).Item(0)
            End If
            dtseq = 0
            If DS.Tables(3).Rows.Count > 0 Then
                dtseqlatest = DS.Tables(3).Rows(0).Item(0)
            End If

            myret = True
        Catch ex As Exception
            myret = False
            errMsgSB.Append(ex.Message)
        End Try
        Return myret
    End Function


End Class

Public Class PanexModel
    Public Property bu As String
    Public Property supplier As String
    Public Property articleno As String
    Public Property materialdesc As String
    Public Property purchasinggroup As String
    Public Property orderno As String
    Public Property orderpos As String
    Public Property subpos As String
    Public Property orderqty As String
    Public Property reqdeldate As String
    Public Property specialflag As String
End Class