Attribute VB_Name = "Database"
 Public db       As New ADODB.Connection
 Public rs(6)       As New ADODB.Recordset
 Public Sql      As String
 Public TimeLogin As Date
 Public IsAdmin As Boolean

Public Sub ConnectionDatabase()
 On Error Resume Next

 If db.State = adStateOpen Then _
    db.Close
    db.CursorLocation = adUseClient
    db.Provider = "Microsoft.JET.OLEDB.4.0;"
    db.Open App.Path & "\database\Main.mdb"
    Exit Sub
End Sub


