Imports System.Text
Public Class ImportConsolidation
    Inherits BaseImport
    'Implements IImport
    Private myAdapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance
    Private MyForm As FormConsolidation
    'Private HDSB As StringBuilder
    'Private DTSB As StringBuilder
    'Private UpdHDSB As StringBuilder
    'Private UpdDTSB As StringBuilder
    'Private hdseq As Long
    'Private dtseq As Long
    'Private hdseqlatest As Long
    'Private dtseqlatest
    Dim DS As DataSet
    'Public Property ErrorMessage As String
    Public Property errMsgSB As New StringBuilder
    Private period As Integer
    Public Sub New(ByVal parent As FormConsolidation)
        Me.MyForm = parent
        Me.DS = parent.myController.DS
    End Sub


    Public Function getLatestPeriod() As Integer
        Dim sqlstr = "select coalesce(period,0) from quality.cwczdt order by period desc limit 1"
        Dim result As Integer
        myAdapter.ExecuteScalar(sqlstr, recordAffected:=result)
        Return result
    End Function

    'Public Function CopyTx() As Boolean Implements IImport.CopyTx
    '    MyForm.ProgressReport(1, "Copy Records..")
    '    Dim myret As Boolean = True
    '    'Header
    '    If HDSB.Length > 0 Then
    '        Dim Sqlstr = String.Format("begin;set statement_timeout to 0;end;select setval('quality.cwczhd_id_seq',{0},false);copy quality.cwczhd(id,serno,vendorname,country,deliverydate,period)  from stdin with null as 'Null'; ", hdseq)
    '        Dim Message = myAdapter.copy(Sqlstr, HDSB.ToString, myret)
    '        If Not myret Then
    '            errMsgSB.Append(Message & vbCrLf)
    '            MyForm.ProgressReport(1, Message)
    '            Return False
    '        End If
    '    End If


    '    'Detail
    '    If DTSB.Length > 0 Then
    '        Dim Sqlstr = String.Format("begin;set statement_timeout to 0;end;select setval('quality.cwczdt_id_seq',{0},false);copy quality.cwczdt(id,cwczhdid,gg,regno,materialdesc,qty,collectedqty,description,cetd,ceta,createddate,period)  from stdin with null as 'Null';", dtseq)
    '        Dim Message = myAdapter.copy(Sqlstr, DTSB.ToString, myret)
    '        If Not myret Then
    '            errMsgSB.Append(Message & vbCrLf)
    '            MyForm.ProgressReport(1, Message)
    '            Return False
    '        End If
    '    End If

    '    'UpdateHeader
    '    If UpdHDSB.Length > 0 Then
    '        'Parent.ProgressReport(2, "Update CMMF")
    '        Dim sqlstr = "update quality.cwczhd cw set deliverydate= foo.deliverydate::date " &
    '                    " from (select * from array_to_set2(Array[" & UpdHDSB.ToString &
    '                 "]) as tb (id character varying, deliverydate character varying))foo where cw.id = foo.id::bigint;"

    '        Dim errmsg As String = String.Empty
    '        If Not myAdapter.ExecuteNonQuery(sqlstr, message:=errmsg) Then
    '            errMsgSB.Append(errmsg & vbCrLf)
    '            Return False
    '        End If
    '    End If

    '    'UpdateDT
    '    If UpdDTSB.Length > 0 Then
    '        'Parent.ProgressReport(2, "Update CMMF")
    '        Dim sqlstr = "update quality.cwczdt cw set qty= foo.qty::integer,collectedqty = foo.collectedqty::integer,description = foo.description,cetd = foo.cetd::date,ceta = foo.ceta::date ,period = foo.period::integer" &
    '                    " from (select * from array_to_set7(Array[" & UpdDTSB.ToString &
    '                 "]) as tb (id character varying, qty character varying, collectedqty character varying, description character varying, cetd character varying, ceta character varying,period character varying))foo where cw.id = foo.id::bigint;"

    '        Dim errmsg As String = String.Empty
    '        If Not myAdapter.ExecuteNonQuery(sqlstr, message:=errmsg) Then
    '            errMsgSB.Append(errmsg & vbCrLf)
    '            Return False
    '        End If
    '    End If

    '    Return myret
    'End Function

    Public Function RunImport(Filename As String) As Boolean
        'Logic
        'Read Text File
        'Find DS.Tables(0)
        'If not available collect the Error
        'else update the datarow
        'Show Progress Result


        Dim myret As Boolean = True

        Dim myList As New List(Of String())
        Dim myrecord() As String
        'Read Text File
        Using objTFParser = New FileIO.TextFieldParser(Filename, Encoding.GetEncoding(1252)) '1252 = ANSI
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0
                MyForm.ProgressReport(1, "Read Data..")

                Do Until .EndOfData
                    myrecord = .ReadFields
                    If myrecord(11) <> "" Then
                        myList.Add(myrecord)
                    End If

                Loop
            End With
            'Check Template File 
            'If myList(0)(0) <> "Line" And myList(0)(1) <> "Ser.no." Then
            '    errMsgSB.Append("Sorry wrong file template.")
            '    Return False
            'End If
            MyForm.ProgressReport(1, "Prepare Data..")
            'Get Existing Data

            'If Not PopulateExistingData() Then
            '    Return False
            'End If
            MyForm.ProgressReport(1, "Build Record..")
            'Dim createhd As Boolean = False
            Try
                For i = 1 To myList.Count - 1
                    'Header
                    'If dtseq = 2 Then
                    '    Debug.Print("test")
                    'End If
                    Dim mymodel As ConsolidationModel = New ConsolidationModel With {.po = myList(i)(1),
                                                                     .item = myList(i)(2),
                                                                     .seq = myList(i)(3),
                                                                     .source = myList(i)(10),
                                                                     .status = myList(i)(11),
                                                                     .inspectorname = myList(i)(12),
                                                                     .inspectiondate = myList(i)(13)}
                    'Table(0)
                    Dim mykey1(3) As Object
                    mykey1(0) = mymodel.po
                    mykey1(1) = mymodel.item
                    mykey1(2) = mymodel.seq
                    mykey1(3) = mymodel.source
                    Dim myresult = DS.Tables(0).Rows.Find(mykey1)

                    If Not IsNothing(myresult) Then

                        Select Case mymodel.status
                            Case "1"
                                myresult.Item("Inspection Type") = True
                            Case "0"
                                myresult.Item("Inspection Type") = False
                            Case ""
                                myresult.Item("Inspection Type") = DBNull.Value
                        End Select
                        Select Case mymodel.inspectiondate
                            Case ""
                                myresult.Item("Inspection Date") = DBNull.Value
                            Case Else
                                myresult.Item("Inspection Date") = CDate(mymodel.inspectiondate)
                        End Select
                        Select Case mymodel.inspectorname
                            Case ""
                                myresult.Item("Inspector Name") = DBNull.Value
                            Case Else
                                myresult.Item("Inspector Name") = mymodel.inspectorname
                        End Select                        
                    Else
                        'Collect Error
                    End If

                Next
            Catch ex As Exception
                errMsgSB.Append(ex.Message)
                Return False
            End Try
        End Using
        Return myret
    End Function
    'Public Function RunImport(filename As String) As Boolean Implements IImport.RunImport
    '    Dim myret As Boolean = True

    '    Dim myList As New List(Of String())
    '    Dim myrecord() As String
    '    HDSB = New StringBuilder
    '    DTSB = New StringBuilder
    '    UpdDTSB = New StringBuilder
    '    UpdHDSB = New StringBuilder
    '    'Read Text File
    '    Using objTFParser = New FileIO.TextFieldParser(filename, Encoding.GetEncoding(1252)) '1252 = ANSI
    '        With objTFParser
    '            .TextFieldType = FileIO.FieldType.Delimited
    '            .SetDelimiters(Chr(9))
    '            .HasFieldsEnclosedInQuotes = True
    '            Dim count As Long = 0
    '            MyForm.ProgressReport(1, "Read Data..")

    '            Do Until .EndOfData
    '                myrecord = .ReadFields
    '                myList.Add(myrecord)
    '            Loop
    '        End With
    '        'Check Template File 
    '        If myList(0)(0) <> "Line" And myList(0)(1) <> "Ser.no." Then
    '            errMsgSB.Append("Sorry wrong file template.")
    '            Return False
    '        End If
    '        MyForm.ProgressReport(1, "Prepare Data..")
    '        'Get Existing Data

    '        If Not PopulateExistingData() Then
    '            Return False
    '        End If
    '        MyForm.ProgressReport(1, "Build Record..")
    '        Dim createhd As Boolean = False
    '        Try
    '            For i = 1 To myList.Count - 1
    '                'Header
    '                'If dtseq = 2 Then
    '                '    Debug.Print("test")
    '                'End If
    '                Dim mymodel As CzechModel = New CzechModel With {.line = myList(i)(0),
    '                                                                 .serno = myList(i)(1),
    '                                                                 .gg = myList(i)(2),
    '                                                                 .regno = myList(i)(3),
    '                                                                 .materialdesc = myList(i)(4),
    '                                                                 .qty = myList(i)(5),
    '                                                                 .collectedqty = myList(i)(6),
    '                                                                 .vendorname = myList(i)(7),
    '                                                                 .country = myList(i)(8),
    '                                                                 .deliverydate = myList(i)(9),
    '                                                                 .cetd = myList(i)(10),
    '                                                                 .ceta = myList(i)(11),
    '                                                                 .description = myList(i)(12)}
    '                'Table Header
    '                Dim mykey1(0) As Object
    '                mykey1(0) = mymodel.serno
    '                Dim myresult = DS.Tables(0).Rows.Find(mykey1)

    '                If Not IsNothing(myresult) Then
    '                    UpdateRecordHD(mymodel, myresult)
    '                    hdseq = myresult.Item("id")
    '                Else
    '                    hdseqlatest = hdseqlatest + 1
    '                    hdseq = hdseqlatest
    '                    CreateRecordHD(mymodel)

    '                End If

    '                'Table Detail
    '                Dim mykey1DT(2) As Object
    '                mykey1DT(0) = hdseq
    '                mykey1DT(1) = mymodel.gg
    '                mykey1DT(2) = mymodel.regno
    '                Dim myresultDT = DS.Tables(1).Rows.Find(mykey1DT)
    '                If Not IsNothing(myresultDT) Then
    '                    UpdateRecordDT(mymodel, myresultDT)
    '                    dtseq = myresultDT.Item("id")
    '                Else
    '                    dtseqlatest = dtseqlatest + 1
    '                    dtseq = dtseqlatest
    '                    CreateRecordDT(mymodel)
    '                End If



    '            Next
    '        Catch ex As Exception
    '            errMsgSB.Append(ex.Message)
    '            Return False
    '        End Try

    '    End Using

    '    'Copy TX
    '    If Not CopyTx() Then
    '        Return False
    '    End If

    '    Return myret
    'End Function

    'Private Sub CreateRecordHD(myModel As CzechModel)
    '    'id bigserial NOT NULL,
    '    'serno character varying,
    '    'vendorname character varying,
    '    'country character varying,
    '    'deliverydate date,
    '    Dim dr As DataRow = DS.Tables(0).NewRow
    '    dr.Item("id") = hdseq
    '    dr.Item("serno") = myModel.serno
    '    dr.Item("vendorname") = myModel.vendorname
    '    dr.Item("country") = myModel.country
    '    If myModel.deliverydate <> "" Then dr.Item("deliverydate") = myModel.deliverydate
    '    DS.Tables(0).Rows.Add(dr)

    '    HDSB.Append(validStr(hdseq) & vbTab &
    '                validStr(myModel.serno) & vbTab &
    '                validStr(myModel.vendorname) & vbTab &
    '                validStr(myModel.country) & vbTab &
    '                validDate(myModel.deliverydate) & vbTab &
    '                period & vbCrLf)

    'End Sub

    'Private Sub UpdateRecordHD(myModel As CzechModel, ByRef myresult As DataRow)
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

    'Private Sub UpdateRecordDT(mymodel As CzechModel, ByRef myresultDT As DataRow)
    '    Dim flagUpdate As Boolean = False
    '    If validint(mymodel.qty) <> myresultDT.Item("qty") Then
    '        flagUpdate = True
    '    End If
    '    If validint(mymodel.collectedqty) <> myresultDT.Item("collectedqty") Then
    '        flagUpdate = True
    '    End If
    '    If IsDBNull(myresultDT.Item("description")) Then
    '        If mymodel.description <> "" Then
    '            flagUpdate = True
    '        End If

    '    Else
    '        If mymodel.description <> myresultDT.Item("description") Then
    '            flagUpdate = True
    '        End If
    '    End If

    '    If IsDBNull(myresultDT.Item("cetd")) Then
    '        If mymodel.cetd <> "" Then
    '            flagUpdate = True
    '        End If

    '    Else
    '        If mymodel.cetd <> myresultDT.Item("cetd") Then
    '            flagUpdate = True
    '        End If
    '    End If

    '    If IsDBNull(myresultDT.Item("ceta")) Then
    '        If mymodel.ceta <> "" Then
    '            flagUpdate = True
    '        End If
    '    Else
    '        If mymodel.ceta <> myresultDT.Item("ceta") Then
    '            flagUpdate = True
    '        End If
    '    End If


    '    If flagUpdate Then
    '        If UpdDTSB.Length > 0 Then
    '            UpdDTSB.Append(",")
    '        End If
    '        UpdDTSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{2}'::character varying,'{3}'::character varying,{4}::character varying,{5}::character varying,{6}::character varying]",
    '                                     myresultDT.Item("id"), validint(mymodel.qty), validint(mymodel.collectedqty), mymodel.description, validDate(mymodel.cetd), validDate(mymodel.ceta), period))
    '    End If
    'End Sub

    'Private Sub CreateRecordDT(mymodel As CzechModel)
    '    'id bigserial NOT NULL,
    '    'cwczhdid bigint,
    '    'gg character varying,
    '    'regno character varying,
    '    'materialdesc character varying,
    '    'qty integer,
    '    'collectedqty integer,
    '    'description character varying,
    '    'cetd date,
    '    'ceta date,
    '    'createddate date,


    '    Dim dr As DataRow = DS.Tables(1).NewRow
    '    dr.Item("id") = dtseq
    '    dr.Item("cwczhdid") = hdseq
    '    dr.Item("gg") = mymodel.gg
    '    dr.Item("regno") = mymodel.regno
    '    dr.Item("materialdesc") = mymodel.materialdesc
    '    dr.Item("qty") = validint(mymodel.qty)
    '    dr.Item("collectedqty") = validint(mymodel.collectedqty)
    '    dr.Item("description") = mymodel.description
    '    If mymodel.cetd <> "" Then dr.Item("cetd") = CDate(mymodel.cetd)
    '    If mymodel.ceta <> "" Then dr.Item("ceta") = CDate(mymodel.ceta)
    '    dr.Item("createddate") = Today.Date
    '    DS.Tables(1).Rows.Add(dr)
    '    DTSB.Append(validStr(dtseq) & vbTab &
    '                validStr(hdseq) & vbTab &
    '                validStr(mymodel.gg) & vbTab &
    '                validStr(mymodel.regno) & vbTab &
    '                validStr(mymodel.materialdesc) & vbTab &
    '                validint(mymodel.qty) & vbTab &
    '                validint(mymodel.collectedqty) & vbTab &
    '                validStr(mymodel.description) & vbTab &
    '                validDate(mymodel.cetd) & vbTab &
    '                validDate(mymodel.ceta) & vbTab &
    '                validDate(Today.Date.ToString) & vbTab &
    '                period & vbCrLf)
    'End Sub

    'Private Function PopulateExistingData() As Boolean
    '    Dim sb As New StringBuilder
    '    Dim myret As Boolean = False
    '    DS = New DataSet
    '    sb.Append("select * from quality.cwczhd;") 'Table 1 : cwczhd
    '    sb.Append("select * from quality.cwczdt;") 'Table 2 : cwczdt        
    '    sb.Append("select id from quality.cwczhd order by id desc limit 1;")
    '    sb.Append("select id from quality.cwczdt order by id desc limit 1;")
    '    'sb.Append("select setval('quality.cwczhd_id_seq',nextval('quality.cwczhd_id_seq')-1,false);select currval('quality.cwczhd_id_seq');") 'Table 3 : seqhd
    '    'sb.Append("select setval('quality.cwczdt_id_seq',nextval('quality.cwczdt_id_seq'),false);select currval('quality.cwczdt_id_seq');") 'Table 4 : seqdt
    '    Try
    '        myAdapter.GetDataset(sb.ToString, DS)
    '        Dim pk1(0) As DataColumn
    '        pk1(0) = DS.Tables(0).Columns("serno")
    '        DS.Tables(0).PrimaryKey = pk1

    '        Dim pk2(2) As DataColumn
    '        pk2(0) = DS.Tables(1).Columns("cwczhdid")
    '        pk2(1) = DS.Tables(1).Columns("gg")
    '        pk2(2) = DS.Tables(1).Columns("regno")
    '        DS.Tables(1).PrimaryKey = pk2
    '        'Sequence
    '        hdseq = 0
    '        If DS.Tables(2).Rows.Count > 0 Then
    '            hdseqlatest = DS.Tables(2).Rows(0).Item(0)
    '        End If
    '        dtseq = 0
    '        If DS.Tables(3).Rows.Count > 0 Then
    '            dtseqlatest = DS.Tables(3).Rows(0).Item(0)
    '        End If

    '        myret = True
    '    Catch ex As Exception
    '        myret = False
    '        errMsgSB.Append(ex.Message)
    '    End Try
    '    Return myret
    'End Function

    Sub ImportData()
        
    End Sub



End Class
