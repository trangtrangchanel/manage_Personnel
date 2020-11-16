Attribute VB_Name = "Module1"
Public cnn As New ADODB.Connection
Public sql As String

Sub ketnoi()
    If cnn.State = 1 Then cnn.Close
    cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\baitapvb.mdb"
    cnn.CursorLocation = adUseClient
    cnn.Open
End Sub
