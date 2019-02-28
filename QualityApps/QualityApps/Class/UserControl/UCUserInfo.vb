Public Class UCUserInfo
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.       
    End Sub
    Public Sub BindingObject()
        Dim myAD = ADPrincipalContext.ADPrincipalContexts(0)
        TextBox1.Text = myAD.DisplayName 'UserName
        TextBox2.Text = myAD.Title 'Title
        TextBox3.Text = myAD.Department 'Department
        TextBox4.Text = myAD.EmployeeNumber 'Employee Number
        TextBox5.Text = myAD.Country 'Country
        TextBox6.Text = myAD.Location 'Location
    End Sub

End Class
