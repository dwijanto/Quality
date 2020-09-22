Imports System.Text
Public Class ImportShanghai
    Inherits BaseImport
    Implements IImport
    Private myAdapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance
    Private MyForm As FormImportCookwareSH
    Private HDSB As StringBuilder
    Private DTSB As StringBuilder
    Private UpdHDSB As StringBuilder
    Private UpdDTSB As StringBuilder
    Private hdseq As Long
    Private dtseq As Long
    Private hdseqlatest As Long
    Private dtseqlatest
    Dim DS As DataSet
    'Public Property ErrorMessage As String
    Public Property errMsgSB As New StringBuilder
    Private period As Integer

    Public Sub New(ByVal parent As FormImportCookwareSH)
        Me.MyForm = parent
        period = parent.YearWeek
    End Sub
    Public Function getLatestPeriod() As Integer
        Dim sqlstr = "select coalesce(period,0) from quality.cwshdt order by period desc limit 1"
        Dim result As Integer
        myAdapter.ExecuteScalar(sqlstr, recordAffected:=result)
        Return result
    End Function
    Public Function CopyTx() As Boolean Implements IImport.CopyTx
        MyForm.ProgressReport(1, "Copy Records..")
        Dim myret As Boolean = True
        'Header
        If HDSB.Length > 0 Then
            Dim Sqlstr = String.Format("begin;set statement_timeout to 0;end;select setval('quality.cwshhd_id_seq',{0},false);copy quality.cwshhd(id,pono,supplier,currency,planner,orderdate,period)  from stdin with null as 'Null'; ", hdseq)
            Dim Message = myAdapter.copy(Sqlstr, HDSB.ToString, myret)
            If Not myret Then
                errMsgSB.Append(Message & vbCrLf)
                MyForm.ProgressReport(1, Message)
                Return False
            End If
        End If


        'Detail
        If DTSB.Length > 0 Then
            Dim Sqlstr = String.Format("begin;set statement_timeout to 0;end;select setval('quality.cwshdt_id_seq',{0},false);copy quality.cwshdt(id,hdid,seq,itemno,itemdescription,fcetd,orderqty,unitprice,totalamount,cetd,createddate,period)  from stdin with null as 'Null';", dtseq)
            Dim Message = myAdapter.copy(Sqlstr, DTSB.ToString, myret)
            If Not myret Then
                errMsgSB.Append(Message & vbCrLf)
                MyForm.ProgressReport(1, Message)
                Return False
            End If
        End If

        'UpdateHeader
        If UpdHDSB.Length > 0 Then
            'Parent.ProgressReport(2, "Update CMMF")
            Dim sqlstr = "update quality.cwshhd cw set orderdate= foo.orderdate::date " &
                        " from (select * from array_to_set2(Array[" & UpdHDSB.ToString &
                     "]) as tb (id character varying, orderdate character varying))foo where cw.id = foo.id::bigint;"

            Dim errmsg As String = String.Empty
            If Not myAdapter.ExecuteNonQuery(sqlstr, message:=errmsg) Then
                errMsgSB.Append(errmsg & vbCrLf)
                Return False
            End If
        End If

        'UpdateDT
        If UpdDTSB.Length > 0 Then
            'Parent.ProgressReport(2, "Update CMMF")
            Dim sqlstr = "update quality.cwshdt cw set orderqty= foo.orderqty::integer,unitprice = foo.unitprice::numeric,totalamount = foo.totalamount::numeric,cetd = foo.cetd::date, period = foo.period::integer " &
                        " from (select * from array_to_set6(Array[" & UpdDTSB.ToString &
                     "]) as tb (id character varying, orderqty character varying, unitprice character varying, totalamount character varying, cetd character varying,period character varying))foo where cw.id = foo.id::bigint;"

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
            If myList(0)(0) <> "PLANNER" And myList(0)(1) <> "ORDER DATE" Then
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
            Dim i As Integer
            Try
                For i = 1 To myList.Count - 1
                    'Header
                    If i = 57 Then
                        Debug.Print("test")
                    End If
                    Dim mymodel As ShanghaiModel = New ShanghaiModel With {.planner = myList(i)(0),
                                                                     .orderdate = myList(i)(1),
                                                                     .supplier = myList(i)(2),
                                                                     .pono = myList(i)(3),
                                                                     .seq = myList(i)(4),
                                                                     .itemno = myList(i)(5),
                                                                     .itemdescription = myList(i)(6),
                                                                     .orderqty = myList(i)(7),
                                                                     .unitprice = myList(i)(8),
                                                                     .totalamount = myList(i)(9),
                                                                     .currency = myList(i)(10),
                                                                     .fcetd = myList(i)(11),
                                                                     .cetd = myList(i)(12)}
                    'Table Header
                    Dim mykey1(0) As Object
                    mykey1(0) = mymodel.pono
                    Dim myresult = DS.Tables(0).Rows.Find(mykey1)

                    If Not IsNothing(myresult) Then
                        UpdateRecordHD(mymodel, myresult)
                        hdseq = myresult.Item("id")
                    Else
                        hdseqlatest = hdseqlatest + 1
                        hdseq = hdseqlatest
                        CreateRecordHD(mymodel)

                    End If

                    'Table Detail
                    Dim mykey1DT(1) As Object
                    mykey1DT(0) = hdseq
                    mykey1DT(1) = mymodel.seq

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
                errMsgSB.Append(String.Format("Row : {0} : {1}", i, ex.Message))
                Return False
            End Try

        End Using

        'Copy TX
        If Not CopyTx() Then
            Return False
        End If

        Return myret

    End Function

    Private Sub CreateRecordHD(myModel As ShanghaiModel)
 
        Dim dr As DataRow = DS.Tables(0).NewRow
        dr.Item("id") = hdseq
        dr.Item("pono") = myModel.pono
        dr.Item("supplier") = myModel.supplier
        dr.Item("currency") = myModel.currency
        dr.Item("planner") = myModel.planner
        If myModel.orderdate <> "" Then dr.Item("orderdate") = myModel.orderdate
        DS.Tables(0).Rows.Add(dr)

        HDSB.Append(validStr(hdseq) & vbTab &
                    validStr(myModel.pono) & vbTab &
                    validStr(myModel.supplier) & vbTab &
                    validStr(myModel.currency) & vbTab &
                    validStr(myModel.planner) & vbTab &
                    validDate(myModel.orderdate) & vbTab &
                    period & vbCrLf)

    End Sub

    Private Sub UpdateRecordHD(myModel As ShanghaiModel, ByRef myresult As DataRow)
        Dim flagUpdate As Boolean = False
        If myModel.orderdate <> myresult.Item("orderdate") Then
            flagUpdate = True
            myresult.Item("orderdate") = myModel.orderdate
        End If
        If flagUpdate Then
            If UpdHDSB.Length > 0 Then
                UpdHDSB.Append(",")
            End If
            UpdHDSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying]", myresult.Item("id"), myModel.orderdate))
        End If
    End Sub

    Private Sub UpdateRecordDT(mymodel As ShanghaiModel, ByRef myresultDT As DataRow)
        Dim flagUpdate As Boolean = False
        If validint(mymodel.orderqty) <> myresultDT.Item("orderqty") Then
            flagUpdate = True
        End If
        If validint(mymodel.unitprice) <> myresultDT.Item("unitprice") Then
            flagUpdate = True
        End If

        If IsDBNull(myresultDT.Item("itemdescription")) Then
            If mymodel.itemdescription <> "" Then
                flagUpdate = True
            End If

        Else
            If mymodel.itemdescription <> myresultDT.Item("itemdescription") Then
                flagUpdate = True
            End If
        End If

        If IsDBNull(myresultDT.Item("cetd")) Then
            If mymodel.cetd <> "" Then
                flagUpdate = True
            End If

        Else
            If mymodel.cetd <> "" Then
                If mymodel.cetd <> myresultDT.Item("cetd") Then
                    flagUpdate = True
                End If
            Else
                flagUpdate = True
            End If

        End If
       

        If flagUpdate Then
            If UpdDTSB.Length > 0 Then
                UpdDTSB.Append(",")
            End If
            UpdDTSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{2}'::character varying,'{3}'::character varying,{4}::character varying,{5}::character varying]",
                                         myresultDT.Item("id"), validint(mymodel.orderqty), validNumeric(mymodel.unitprice), validNumeric(mymodel.totalamount), validDate(mymodel.cetd), period))
        End If
    End Sub

    Private Sub CreateRecordDT(mymodel As ShanghaiModel)


        Dim dr As DataRow = DS.Tables(1).NewRow
        dr.Item("id") = dtseq
        dr.Item("hdid") = hdseq
        dr.Item("seq") = mymodel.seq
        dr.Item("itemno") = mymodel.itemno
        dr.Item("itemdescription") = mymodel.itemdescription        
        dr.Item("fcetd") = mymodel.fcetd
        dr.Item("orderqty") = validint(mymodel.orderqty)
        dr.Item("unitprice") = validint(mymodel.unitprice)
        dr.Item("totalamount") = mymodel.totalamount
        If mymodel.cetd <> "" Then dr.Item("cetd") = CDate(mymodel.cetd)       
        dr.Item("createddate") = Today.Date
        DS.Tables(1).Rows.Add(dr)
        DTSB.Append(validStr(dtseq) & vbTab &
                    validStr(hdseq) & vbTab &
                    validint(mymodel.seq) & vbTab &
                    validStr(mymodel.itemno) & vbTab &
                    validStr(mymodel.itemdescription) & vbTab &
                    validStr(mymodel.fcetd) & vbTab &
                    validint(mymodel.orderqty) & vbTab &
                    validNumeric(mymodel.unitprice) & vbTab &
                    validNumeric(mymodel.totalamount) & vbTab &                    
                    validDate(mymodel.cetd) & vbTab &
                    validDate(Today.Date.ToString) & vbTab &
                    period & vbCrLf)
    End Sub

    Private Function PopulateExistingData() As Boolean
        Dim sb As New StringBuilder
        Dim myret As Boolean = False
        DS = New DataSet
        sb.Append("select * from quality.cwshhd;") 'Table 1 : cwczhd
        sb.Append("select * from quality.cwshdt;") 'Table 2 : cwczdt        
        sb.Append("select id from quality.cwshhd order by id desc limit 1;")
        sb.Append("select id from quality.cwshdt order by id desc limit 1;")
        'sb.Append("select setval('quality.cwczhd_id_seq',nextval('quality.cwczhd_id_seq')-1,false);select currval('quality.cwczhd_id_seq');") 'Table 3 : seqhd
        'sb.Append("select setval('quality.cwczdt_id_seq',nextval('quality.cwczdt_id_seq'),false);select currval('quality.cwczdt_id_seq');") 'Table 4 : seqdt
        Try
            myAdapter.GetDataset(sb.ToString, DS)
            Dim pk1(0) As DataColumn
            pk1(0) = DS.Tables(0).Columns("pono")
            DS.Tables(0).PrimaryKey = pk1

            Dim pk2(2) As DataColumn
            pk2(0) = DS.Tables(1).Columns("hdid")
            pk2(1) = DS.Tables(1).Columns("seq")
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

Public Class ShanghaiModel
    Public Property planner As String
    Public Property orderdate As String
    Public Property supplier As String
    Public Property pono As String
    Public Property seq As String
    Public Property itemno As String    
    Public Property itemdescription As String
    Public Property orderqty As String
    Public Property unitprice As String
    Public Property totalamount As String
    Public Property currency As String
    Public Property fcetd As String
    Public Property cetd As String


End Class