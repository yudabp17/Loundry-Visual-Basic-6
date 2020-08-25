Attribute VB_Name = "Module1"
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Public sql As String

Public Sub connect()
con.ConnectionString = "Provider=MSDASQL.1;Password=;Persist Security Info=True;User ID=root;Data Source=Uts_Loundry"
con.Open
End Sub

