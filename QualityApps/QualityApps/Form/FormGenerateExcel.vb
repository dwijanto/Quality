Imports System.Threading

Public Class FormGenerateExcel
    Dim myThread As New System.Threading.Thread(AddressOf DoExport)
    

    Public Shared myForm As FormGenerateExcel

    Public Shared Function getInstance()
        If myForm Is Nothing Then
            myForm = New FormGenerateExcel
        ElseIf myForm.IsDisposed Then
            myForm = New FormGenerateExcel
        End If
        Return myForm
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        CreateExcel()
        'DoExport()
    End Sub

    Private Sub CreateExcel()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoExport)
            myThread.Start()
        Else
            ProgressReport(1, "Please wait until the current process finished.")
        End If
    End Sub

    Sub DoExport()
        ProgressReport(6, "Start Export")
        Try
            Dim myGenerateExcel As New GenerateExcel(Me)

            Dim sw As New Stopwatch
            sw.Start()
            If Not myGenerateExcel.Run Then
                ProgressReport(1, myGenerateExcel.ErrorMessage)
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