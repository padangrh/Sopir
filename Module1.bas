Attribute VB_Name = "Module1"
Public fMainForm As Form_Login
Public con As New ADODB.Connection
Public username As String
Public userstatus As Integer
Public status

   
Public Sub connect()
    If con.State = adStateClosed Then
        con.ConnectionString = "Provider=MSDASQL.1;Password=yuyu;Persist Security Info=True;User ID=root;Data Source=sopir"
        con.Open
    End If
End Sub



Sub Main()
    Set fMainForm = New Form_Login
    fMainForm.Show
End Sub

