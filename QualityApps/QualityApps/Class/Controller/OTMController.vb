Imports System.Text
Imports System.IO

Public Class OTMController
    Implements IController
    Dim DS As DataSet

    Dim myAdapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance
    Dim _ErrorMessage As String
    Private SaveFileName As String = String.Empty
    Dim SB As New StringBuilder
    Private _parent As FormGenerateCSVOTM
    Public Sub New(Parent As FormGenerateCSVOTM)
        _parent = Parent
    End Sub

    Public Property ErrorMessage As String
        Get
            Return _ErrorMessage
        End Get
        Set(value As String)
            _ErrorMessage = value
        End Set
    End Property

    Public Property FileName As String
        Get
            Return SaveFileName
        End Get
        Set(value As String)
            SaveFileName = value
        End Set
    End Property


    Public ReadOnly Property GetTable As DataTable Implements IController.GetTable
        Get
            Return DS.Tables(0)
        End Get
    End Property

    Public Function loaddata() As Boolean Implements IController.loaddata
        Dim myret As Boolean = True
        DS = New DataSet
        Dim sqlstr = "select o.id,o.cpo,to_char(startdate,'FMyyyyMMddHHMMSS') as inspdate,t.inspector,t.samplesize,t.insplot,substring(o.id from 1 for 4) as domain from quality.otmtx o" &
                     " left join quality.dailytx t on t.purchdoc = o.purchdoc and o.item = t.item and o.seqn = t.seqn " &
                     " where insplot <> 0"
        Try
            _parent.ProgressReport(1, "Getting Data...")
            If myAdapter.GetDataset(sqlstr, DS) Then
                myret = save()
            End If
        Catch ex As Exception
            myret = False
            _ErrorMessage = ex.Message
        End Try

        Return myret
    End Function

    Public Function save() As Boolean Implements IController.save
        Dim myret As Boolean = False
        'Read Data
        Try
            _parent.ProgressReport(1, "Generating CSV....")
            If DS.Tables(0).Rows.Count > 0 Then
                Dim Header As String = String.Format("ORDER_RELEASE{0}{0}{0}{0}{0}{0}{1}ORDER_RELEASE_GID{0}ORDER_RELEASE_XID{0}INSPECTION_DATE{0}ATTRIBUTE6{0}ATTRIBUTE_NUMBER6{0}ATTRIBUTE_NUMBER7{0}DOMAIN_NAME{1}EXEC SQL ALTER SESSION SET NLS_DATE_FORMAT = 'YYYYMMDDHH24MISS'{0}{0}{0}{0}{0}{0}{1}", ",", vbCrLf)
                SB.Clear()
                SB.Append(Header)
                For Each dr As DataRow In DS.Tables(0).Rows
                    SB.Append(String.Format("{0},{1},{2},{3},{4},{5},{6}{7}", dr.Item("id"), dr.Item("cpo"), dr.Item("inspdate"), dr.Item("inspector"), dr.Item("samplesize"), dr.Item("insplot"), dr.Item("domain"), vbCrLf))
                Next
                'Save to File
                _parent.ProgressReport(1, "Saving CSV....")
                Using mystream As New StreamWriter(SaveFileName)
                    mystream.Write(SB.ToString)
                End Using
                _parent.ProgressReport(1, "Saving CSV Done.")
                myret = True
            Else               
                _ErrorMessage = "Data not available."
            End If
            
        Catch ex As Exception
            _ErrorMessage = ex.Message           
        End Try
               
        Return myret
    End Function
End Class
