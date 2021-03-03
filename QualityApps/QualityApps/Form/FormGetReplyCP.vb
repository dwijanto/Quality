﻿Imports System.Threading
Public Class FormGetReplyCP

    Public Shared myform As FormGetReplyCP
    Private TotalProcess As Integer
    Public TextFileList As New List(Of String)
    Public Shared Function getInstance()
        If myform Is Nothing Then
            myform = New FormGetReplyCP
        ElseIf myform.IsDisposed Then
            myform = New FormGetReplyCP
        End If
        Return myform
    End Function

    Dim myThread As New Thread(AddressOf DoGetReply)
    Private myGetReply As GetReplyCP

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        GetReplyFromSupplier()
    End Sub

    Private Sub GetReplyFromSupplier()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoGetReply)
            myThread.Start()
        Else
            ProgressReport(1, "Please wait until the current process finished.")
        End If
    End Sub

    Private Sub ImportReply()
        myGetReply.ImportFile(TextFileList)
    End Sub

    Sub DoGetReply()
        myGetReply = New GetReplyCP(Me)
        myGetReply.run()
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

                Case 2004
                    TotalProcess = message
                Case 2005
                    'TextFileList.Add()
                    If message <> "" Then
                        TextFileList.Add(message)
                    End If
                    TotalProcess = TotalProcess - 1
                    If TotalProcess = 0 Then
                        ProgressReport(1, "Conversion file completed.")
                        ImportReply()
                    End If


            End Select
        End If
    End Sub

    Private Sub StatusStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles StatusStrip1.ItemClicked

    End Sub
End Class