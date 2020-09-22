Attribute VB_Name = "Module1"
'developer: Shyam Singh Chandel < shyamschandel@rediffmail.com >

Public CN             As ADODB.Connection
Public RS             As ADODB.Recordset

Public Sub connectDB()
On Error GoTo ErrHandler
    Dim CNPath As String
    Set CN = New ADODB.Connection
    CNPath = "Data Source=" & App.Path & "\usdb.mdb"
    CN.Provider = "Microsoft Jet 4.0 OLE DB Provider"
    CN.ConnectionString = CNPath
    CN.Open
   Exit Sub
ErrHandler:
MsgBox "Either Database does not exist or" & vbCrLf & "misplaced location.", vbCritical, "Not Found!"
CN.Close
End
End Sub

