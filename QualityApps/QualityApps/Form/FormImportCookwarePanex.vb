Imports System.Threading
Public Class FormImportCookwarePanex
    Public Shared myForm As FormImportCookwarePanex
    Public YearWeek As Integer
    Private LatestPeriod As Integer
    Public Shared Function getInstance()
        If myForm Is Nothing Then
            myForm = New FormImportCookwarePanex
        ElseIf myForm.IsDisposed Then
            myForm = New FormImportCookwarePanex
        End If
        Return myForm
    End Function

    Dim myThread As New System.Threading.Thread(AddressOf DoImport)
    Private MyImport As New ImportPanex(Me)
    Private ImportFileName As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.Validate Then
            YearWeek = CInt(TextBox1.Text)
            ImportData()
        End If
    End Sub
    Public Overloads Function validate() As Boolean
        ErrorProvider1.SetError(TextBox1, "")
        If TextBox1.Text.Length = 0 Then
            ErrorProvider1.SetError(TextBox1, "Value cannot be blank.")
            Return False
        End If
        If Not IsNumeric(TextBox1.Text) Then
            ErrorProvider1.SetError(TextBox1, "Only number is allowed.")
            Return False
        End If
        Return True
    End Function
    Private Sub ImportData()
        Dim OpenFileDialog1 As New OpenFileDialog
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                ImportFileName = OpenFileDialog1.FileName
                myThread = New Thread(AddressOf DoImport)
                myThread.Start()
            Else
                ProgressReport(1, "Open file cancelled.")
            End If

        End If
    End Sub

    Sub DoImport()
        ProgressReport(6, "Start Import")
        Try
            Dim myImport = New ImportPanex(Me)
            LatestPeriod = myImport.getLatestPeriod
            'Validate YearWeek
            If Not LatestPeriod <= YearWeek Then
                Err.Raise(1, Description:=String.Format("Wrong Period. Current Period {0}", LatestPeriod))
            End If

            Dim sw As New Stopwatch
            sw.Start()



            If myImport.RunImport(ImportFileName) Then
                sw.Stop()
                ProgressReport(1, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            Else
                ProgressReport(1, String.Format("{0}", myImport.errMsgSB.ToString))
            End If
        Catch ex As Exception
            ProgressReport(1, String.Format("{0}", ex.Message))
        Finally
            ProgressReport(5, "End Import")
        End Try

    End Sub
    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    ToolStripStatusLabel1.Text = message
                Case 4

                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
            End Select
        End If
    End Sub
End Class