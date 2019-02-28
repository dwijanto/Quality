Imports System.Threading

Public Class FormGenerateCSVOTM

    Dim myThread As New System.Threading.Thread(AddressOf DoWork)


    Public Shared myForm As FormGenerateCSVOTM
    Private SaveFileName As String = String.Empty
    Private SaveFileDialog1 As New SaveFileDialog

    Public Shared Function getInstance()
        If myForm Is Nothing Then
            myForm = New FormGenerateCSVOTM
        ElseIf myForm.IsDisposed Then
            myForm = New FormGenerateCSVOTM
        End If
        Return myForm
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DoCSV()
    End Sub

    Private Sub DoCSV()
        'Get Location
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            'Get Location
            SaveFileDialog1.FileName = "OTM.csv"
            If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                SaveFileName = SaveFileDialog1.FileName
                myThread = New Thread(AddressOf DoWork)
                myThread.Start()
            Else
                ProgressReport(1, "Open file cancelled.")
            End If
        Else
            ProgressReport(1, "Please wait until the current process finished.")
        End If
    End Sub

    Sub DoWork()
        ProgressReport(6, "Start Export")
        Try
            Dim myController As New OTMController(Me)
            myController.FileName = SaveFileName
            Dim sw As New Stopwatch
            sw.Start()
            If Not myController.loaddata Then
                ProgressReport(1, myController.ErrorMessage)
                Exit Sub
            Else
                ProgressReport(1, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            End If
        Catch ex As Exception
            ProgressReport(1, ex.Message)
        Finally
            ProgressReport(5, "End Export")
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