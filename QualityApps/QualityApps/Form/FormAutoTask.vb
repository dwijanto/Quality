Imports System.Threading
Imports Microsoft.Office.Interop

Public Enum AUTOGENERATE
    AUTO = 0
    NON_AUTO = 1
End Enum
Public Class FormAutoTask
    Public dbAdapter1 As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance
    Private AutoGenerate As AUTOGENERATE
    Dim myThread As New System.Threading.Thread(AddressOf doWork)

    Private Sub FormAutoGenerate_Load(sender As Object, e As EventArgs) Handles Me.Load
        If AutoGenerate = QualityApps.AUTOGENERATE.AUTO Then
            Me.WindowState = FormWindowState.Minimized
            LoadMe()
        End If
    End Sub

    Private Sub LoadMe()
        If Not myThread.IsAlive Then
            Try
                ToolStripStatusLabel1.Text = ""
                myThread = New System.Threading.Thread(AddressOf doWork)
                myThread.TrySetApartmentState(ApartmentState.MTA)
                myThread.Start()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub

    Sub doWork()
        Logger.log("--------Start----------")

        ProgressReport(6, "Start")
        Logger.log("Generate Report.")


        Dim myreport = New AutoInspectionReport()
        If Not myreport.GenerateReport Then
            Logger.log(myreport.errmsg)
        End If

        ProgressReport(5, "End")
        ProgressReport(1, "Close Apps")
        Logger.log("--------End----------")
    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Me.Close()
                Case 2
                    ToolStripStatusLabel1.Text = message
                Case 3

                Case 4

                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7

            End Select
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        LoadMe()
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        AutoGenerate = QualityApps.AUTOGENERATE.AUTO
        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Sub New(status As AUTOGENERATE)

        ' This call is required by the designer.
        InitializeComponent()
        AutoGenerate = status
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub StatusStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles StatusStrip1.ItemClicked

    End Sub


End Class