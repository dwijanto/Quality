Public Class SendEmailControllerCP
    Public mymodel As SendEmailModelCP
    Public DS As DataSet
    Public BS As BindingSource
    Public ReadOnly Property startdate As Date
        Get
            Return myModel.startdate
        End Get
    End Property

    Public ReadOnly Property enddate As Date
        Get
            Return myModel.enddate
        End Get
    End Property

    Public ReadOnly Property lastsenddate As Date
        Get
            Return myModel.LastSendDate
        End Get
    End Property

    Public Function LoadData(ByVal InspectionDate As Date) As Boolean
        Dim myret As Boolean = True
        DS = New DataSet
        mymodel = New SendEmailModelCP(InspectionDate)
        myret = mymodel.LoadData(DS)
        If myret Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("vendor")
            DS.Tables(0).PrimaryKey = pk
            BS = New BindingSource
            BS.DataSource = DS.Tables(0)
        End If
        Return myret
    End Function


    Public Sub showMessage(ByVal message As String, ByVal fieldname As String)
        Dim myValue() = message.Split(",")
        Dim mykey(0) As Object
        mykey(0) = myValue(0)
        Dim myresult = DS.Tables(0).Rows.Find(mykey)
        If Not IsNothing(myresult) Then
            myresult.Item(fieldname) = myValue(1)
            myresult.EndEdit()
            'Check Send Internal mykey = 0
            If mykey(0) = 0 And myValue(1) = "Sending email...Done" Then
                Dim myparam As ParamAdapter = ParamAdapter.getInstance
                myparam.SaveSendEmailTxCP()
            End If
        End If
    End Sub
End Class
