Imports System.Windows.Forms

Public Class DialogAskDateRange



    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        DateTimePicker2.Value = nextday(Today.Date)
    End Sub

    Private Function nextday(mydate As Date) As Date
        Dim addition As Integer = 1
        If mydate.DayOfWeek = 5 Then
            addition = 3
        End If
        Return mydate.AddDays(addition)
    End Function
End Class
