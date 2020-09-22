Imports System.Text

Public Class VendorAssignmentController
    Implements IController
    Implements IToolbarAction

    Public Model As New VendorAssignmentModel
    Public BS As BindingSource
    Dim DS As DataSet

    Dim SBUBS As BindingSource
    Dim VendorBS As BindingSource
    Dim VendorHelperBS As BindingSource
    Dim MonitoringLevelBS As BindingSource
    Dim QESB As StringBuilder
    Public errMsgSB As StringBuilder
    Dim UpdQESB As StringBuilder

    Dim QETeamDict As Dictionary(Of String, Integer)
    Dim SBUDict As Dictionary(Of String, Integer)
    Dim MonitoringLevelDict As Dictionary(Of String, Integer)

    Public ReadOnly Property ErrorMessage As String
        Get
            Return errmsgSB.tostring
        End Get
    End Property



    Public ReadOnly Property GetDataset As DataSet
        Get
            Return DS
        End Get
    End Property
    Public ReadOnly Property GetVendorBS As BindingSource
        Get
            Return VendorBS
        End Get
    End Property

    Public ReadOnly Property GetVendorHelperBS As BindingSource
        Get
            Return VendorHelperBS
        End Get
    End Property
    Public ReadOnly Property GetSBUBS As BindingSource
        Get
            Return SBUBS
        End Get
    End Property

    Public ReadOnly Property GetMonitoringLevelBS As BindingSource
        Get
            Return monitoringlevelbs
        End Get
    End Property

    Public ReadOnly Property GetTable As DataTable Implements IController.GetTable
        Get
            Return DS.Tables(Model.tablename).Copy()
        End Get
    End Property



    Public Function loaddata() As Boolean Implements IController.loaddata
        Dim myret As Boolean = False
        Model = New VendorAssignmentModel
        DS = New DataSet
        If Model.LoadData(DS) Then
            BS = New BindingSource
            BS.DataSource = DS.Tables(0)
            SBUBS = New BindingSource
            SBUBS.DataSource = DS.Tables("SBUSAP")
            VendorBS = New BindingSource
            VendorBS.DataSource = DS.Tables("Vendor")
            VendorHelperBS = New BindingSource
            VendorHelperBS.DataSource = New DataView(DS.Tables("Vendor"))
            MonitoringLevelBS = New BindingSource
            MonitoringLevelBS.DataSource = DS.Tables("MonitoringLevel")
            myret = True
        End If
        Return myret
    End Function

    Public Function save() As Boolean Implements IController.save
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

    Public Property ApplyFilter As String Implements IToolbarAction.ApplyFilter
        Get
            Return BS.Filter
        End Get
        Set(ByVal value As String)
            BS.Filter = String.Format(Model.FilterField, value)
        End Set
    End Property

    Public Property ApplyFilterAll As String
        Get
            Return BS.Filter
        End Get
        Set(ByVal value As String)
            BS.Filter = String.Format(Model.FilterField, value)
        End Set
    End Property
    Public Function GetCurrentRecord() As DataRowView Implements IToolbarAction.GetCurrentRecord
        Return BS.Current
    End Function

    Public Function GetNewRecord() As DataRowView Implements IToolbarAction.GetNewRecord
        Return BS.AddNew
    End Function

    Public Sub RemoveAt(value As Integer) Implements IToolbarAction.RemoveAt
        BS.RemoveAt(value)
    End Sub

    Public Function Save(mye As ContentBaseEventArgs) As Boolean Implements IToolbarAction.Save
        Dim myret As Boolean = False
        If Model.save(Me, mye) Then
            myret = True
        End If
        Return myret
    End Function

    Public Function ImportTextFile(Parent As Object, filename As String) As Boolean
        errMsgSB = New StringBuilder
        Dim myret As Boolean = False

        Dim myList As New List(Of String())
        Dim myrecord() As String
        QESB = New StringBuilder
        UpdQESB = New StringBuilder

        InitDict()


        'Read Text File
        Using objTFParser = New FileIO.TextFieldParser(filename, Encoding.GetEncoding(1252)) '1252 = ANSI
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0
                Parent.ProgressReport(1, "Read Data..")

                Do Until .EndOfData
                    myrecord = .ReadFields
                    myList.Add(myrecord)
                Loop
            End With
            Parent.ProgressReport(1, "Build Record..")
            Parent.ProgressReport(5, "Set To Continuous")
            For i = 1 To myList.Count - 1

                Dim mydata As New VendorQE With {.id = myList(i)(0),
                                                .vendorcode = myList(i)(1),
                                                .vendorname = myList(i)(2),
                                                 .shortname = myList(i)(3),
                                                 .qe = myList(i)(4),
                                                 .qeteam = myList(i)(5),
                                                 .sbu = myList(i)(6),
                                                 .monitoringlevel = myList(i)(7),
                                                 .location = myList(i)(8)
                                               }
                If mydata.id.Length > 0 Then
                    'Update
                    BuildDataUpdate(mydata)
                Else
                    QESB.Append(mydata.vendorcode & vbTab &
                                mydata.qe & vbTab &
                                getId(mydata.qeteam, QETeamDict) & vbTab &
                                getId(mydata.sbu, SBUDict) & vbTab &
                                getId(mydata.monitoringlevel, MonitoringLevelDict) & vbTab &
                                mydata.location & vbCrLf)
                End If
            Next
        End Using
        If errMsgSB.Length = 0 Then
            If UpdQESB.Length > 0 Then
                Parent.ProgressReport(2, "Update CMMF")
                'cmmf.activitycode, cmmf.brandid, cmmf.cmmftype, cmmf.comfam, cmmf.commercialref, cmmf.materialdesc, cmmf.modelcode, cmmf.plnt, cmmf.rir, cmmf.sbu, cmmf.sorg
                Dim sqlstr = "update quality.vendorassignment set userid= foo.userid,fgcp = foo.fgcp::integer,sbuid = foo.sbuid::integer,monitoringlevelid= case when foo.monitoringlevelid = 'Null' then Null else foo.monitoringlevelid::integer end,factorylocation=foo.factorylocation" &
                            " from (select * from array_to_set6(Array[" & UpdQESB.ToString &
                         "]) as tb (id character varying, userid character varying, fgcp character varying, sbuid character varying,monitoringlevelid character varying, factorylocation character varying ))foo where quality.vendorassignment.id = foo.id::bigint;"
                'Dim ra As Long
                Dim errmsg As String = String.Empty
                If Not Model.myadapter.ExecuteNonQuery(sqlstr, message:=errmsg) Then
                    errMsgSB.Append(errmsg & vbCrLf)
                    Parent.ProgressReport(2, "Update QE" & "::" & errmsg)
                    Return False
                End If
            End If

            If QESB.Length > 0 Then
                Parent.ProgressReport(1, "Start Add New Records")
                'mystr.Append("delete from cmmfspspm;")        
                Dim sqlstr As String = String.Format("begin;set statement_timeout to 0;end;{0};copy quality.vendorassignment(vendorcode,userid,fgcp,sbuid,monitoringlevelid,factorylocation) from stdin with null as 'Null';", "")
                Dim ra As Long = 0
                Try

                    Dim errmessage = Model.myadapter.copy(sqlstr, QESB.ToString, myret)
                    If myret Then
                        Parent.ProgressReport(1, "Add Records Done.")
                    Else
                        errMsgSB.Append("Copy Error " & ErrorMessage & vbCrLf)
                    End If
                Catch ex As Exception
                    errMsgSB.Append(ex.Message & vbCrLf)
                End Try
            End If
            myret = True
        End If
        

        Return myret
    End Function

    Private Sub BuildDataUpdate(mydata As VendorQE)
        'Find Existing Record
        Dim FlagUpdate As Boolean = False
        Dim mycol(0) As Object
        mycol(0) = mydata.id
        Dim result As DataRow = DS.Tables(0).Rows.Find(mycol)
        If Not IsNothing(result) Then
            If Not equalTo(result.Item(9), mydata.qe) Then
                FlagUpdate = True
            End If           
            If Not equalTo(result.Item(4), getId(mydata.qeteam, QETeamDict)) Then
                FlagUpdate = True
            End If
            If Not equalTo(result.Item(6), getId(mydata.sbu, SBUDict)) Then
                FlagUpdate = True
            End If
            If Not equalTo(result.Item(14), getId(mydata.monitoringlevel, MonitoringLevelDict)) Then
                FlagUpdate = True
            End If
            If Not equalTo(result.Item(13), mydata.location) Then
                FlagUpdate = True
            End If

            If FlagUpdate Then
                If UpdQESB.Length > 0 Then
                    UpdQESB.Append(",")
                End If
                UpdQESB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{2}'::character varying,'{3}'::character varying,'{4}'::character varying,'{5}'::character varying]", mydata.id, mydata.qe, getId(mydata.qeteam, QETeamDict), getId(mydata.sbu, SBUDict), getId(mydata.monitoringlevel, MonitoringLevelDict), mydata.location))
            End If
        Else
            'Found Error -> Id is not found.
            errMsgSB.Append(String.Format("BuildDataUpdate: Id {0} not found. cannot be updated.{1}", mydata.id, vbCrLf))
        End If

    End Sub

    Private Sub BuildDataCreate(mydata As VendorQE)
        Throw New NotImplementedException
    End Sub

    Public Sub New()
        QETeamDict = New Dictionary(Of String, Integer)
        QETeamDict.Add("Finished Goods", 1)
        QETeamDict.Add("Component", 2)
        QETeamDict.Add("Cookware", 3)
    End Sub

    Private Function getId(key As String, Dict As Dictionary(Of String, Integer)) As Object
        If key.Length = 0 Then
            Return "Null"
        End If
        Dim myret As Integer = 0
        Dict.TryGetValue(key, myret)
        If myret = 0 Then
            errMsgSB.Append(String.Format("getId: Key {0} not found. cannot be updated.{1}", key, vbCrLf))
        End If
        Return myret
    End Function

    Private Function equalTo(p1 As Object, p2 As String) As Boolean
        Dim myret As Boolean = True
        If IsDBNull(p1) Then
            If p2.Length > 0 Then myret = False
        Else
            If p1 <> p2 Then
                myret = False
            End If
        End If
        Return myret
    End Function

    Private Sub InitDict()
        SBUDict = New Dictionary(Of String, Integer)
        MonitoringLevelDict = New Dictionary(Of String, Integer)
        If Not IsNothing(DS) Then
            For Each dr As DataRow In DS.Tables("SBUSAP").Rows
                SBUDict.Add(dr.Item(1), dr.Item(0))
            Next
            For Each dr As DataRow In DS.Tables("MonitoringLevel").Rows
                MonitoringLevelDict.Add(dr.Item(1), dr.Item(0))
            Next
        End If


    End Sub

    Private Function validchar(p1 As String) As String
        Dim myret As String
        If p1 = "" Then
            myret = "Null"
        Else
            myret = String.Format("'{0}'", p1)
        End If
        Return myret
    End Function

End Class

Public Class VendorQE
    Public Property id As String
    Public Property vendorcode As String
    Public Property vendorname As String
    Public Property shortname As String
    Public Property qe As String
    Public Property qeteam As String
    Public Property sbu As String
    Public Property monitoringlevel As String
    Public Property location As String
End Class
