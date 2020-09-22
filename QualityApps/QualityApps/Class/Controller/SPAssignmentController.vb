Imports System.Text

Public Class SPAssignmentController
    Implements IController
    Implements IToolbarAction

    Public Model As New SPAssignmentModel
    Public BS As BindingSource
    Dim DS As DataSet

    Dim SBUBS As BindingSource
    Dim VendorBS As BindingSource
    Dim VendorHelperBS As BindingSource
    Dim EQBS As BindingSource
    Public errMsgSB As StringBuilder
    Dim SPSB As StringBuilder
    Dim SBUDict As Dictionary(Of String, Integer)
    Dim MonitoringLevelDict As Dictionary(Of String, Integer)
    Dim UpdSPSB As Object

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

    Public ReadOnly Property GetTable As DataTable Implements IController.GetTable
        Get
            Return DS.Tables(Model.tablename).Copy()
        End Get
    End Property



    Public Function loaddata() As Boolean Implements IController.loaddata
        Dim myret As Boolean = False
        Model = New SPAssignmentModel
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
        SPSB = New StringBuilder
        UpdSPSB = New StringBuilder

        'InitDict()


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

                Dim mydata As New VendorSP With {.id = myList(i)(0),
                                                .vendorcode = myList(i)(1),
                                                .vendorname = myList(i)(2),
                                                 .shortname = myList(i)(3),
                                                 .sp = myList(i)(4),
                                                 .backup1 = myList(i)(5),
                                                 .backup2 = myList(i)(6),
                                                 .sbu = myList(i)(7)
                                               }
                If mydata.id.Length > 0 Then
                    'Update
                    BuildDataUpdate(mydata)
                Else
                    SPSB.Append(mydata.vendorcode & vbTab &
                                mydata.sp & vbTab &
                                mydata.backup1 & vbTab &
                                mydata.backup2 & vbTab &
                                mydata.sbu & vbCrLf)
                End If
            Next
        End Using
        If errMsgSB.Length = 0 Then
            If UpdSPSB.Length > 0 Then
                Parent.ProgressReport(2, "Update CMMF")
                'cmmf.activitycode, cmmf.brandid, cmmf.cmmftype, cmmf.comfam, cmmf.commercialref, cmmf.materialdesc, cmmf.modelcode, cmmf.plnt, cmmf.rir, cmmf.sbu, cmmf.sorg
                Dim sqlstr = "update quality.vendorsp set spcode= foo.sp,bck1 = foo.bck1,bck2 = foo.bck2,bu=foo.sbu " &
                            " from (select * from array_to_set5(Array[" & UpdSPSB.ToString &
                         "]) as tb (id character varying, sp character varying, bck1 character varying, bck2 character varying,sbu character varying))foo where quality.vendorsp.id = foo.id::bigint;"
                'Dim ra As Long
                Dim errmsg As String = String.Empty
                If Not Model.myadapter.ExecuteNonQuery(sqlstr, message:=errmsg) Then
                    errMsgSB.Append(errmsg & vbCrLf)
                    Parent.ProgressReport(2, "Update QE" & "::" & errmsg)
                    Return False
                End If
            End If

            If SPSB.Length > 0 Then
                Parent.ProgressReport(1, "Start Add New Records")
                'mystr.Append("delete from cmmfspspm;")        
                Dim sqlstr As String = String.Format("begin;set statement_timeout to 0;end;{0};copy quality.vendorsp(vendorcode,spcode,bck1,bck2,bu) from stdin with null as 'Null';", "")
                Dim ra As Long = 0
                Try

                    Dim errmessage = Model.myadapter.copy(sqlstr, SPSB.ToString, myret)
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

    Private Sub BuildDataUpdate(mydata As VendorSP)
        'Find Existing Record
        Dim FlagUpdate As Boolean = False
        Dim mycol(0) As Object
        mycol(0) = mydata.id
        Dim result As DataRow = DS.Tables(0).Rows.Find(mycol)
        If Not IsNothing(result) Then
            If Not equalTo(result.Item(4), mydata.sp) Then
                FlagUpdate = True
            End If

            If Not equalTo(result.Item(5), mydata.backup1) Then
                FlagUpdate = True
            End If
            If Not equalTo(result.Item(6), mydata.backup2) Then
                FlagUpdate = True
            End If
            If Not equalTo(result.Item(7), mydata.sbu) Then
                FlagUpdate = True
            End If

            If FlagUpdate Then
                If UpdSPSB.Length > 0 Then
                    UpdSPSB.Append(",")
                End If
                UpdSPSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{2}'::character varying,'{3}'::character varying,'{4}'::character varying]", mydata.id, mydata.sp, mydata.backup1, mydata.backup2, mydata.sbu))
            End If
        Else
            'Found Error -> Id is not found.
            errMsgSB.Append(String.Format("BuildDataUpdate: Id {0} not found. cannot be updated.{1}", mydata.id, vbCrLf))
        End If

    End Sub

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
End Class

Public Class VendorSP
    Public Property id As String
    Public Property vendorcode As String
    Public Property vendorname As String
    Public Property shortname As String
    Public Property sp As String
    Public Property backup1 As String
    Public Property backup2 As String
    Public Property sbu As String
End Class
