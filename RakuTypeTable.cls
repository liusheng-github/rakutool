VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RakuTypeTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public tblName As String
Public dbName As String
Public action As String
Public selCond As String
Public selOrder As String
Public columns As New Collection
Public values As New Collection

Private Sub Class_Initialize()
    tblName = ""
    dbName = ""
    action = ""
End Sub

Private Sub Class_Terminate()
    For i = 1 To columns.Count
        columns.Remove 1
    Next
    Set columns = Nothing
    
    For i = 1 To values.Count
        values.Remove 1
    Next
    
    Set values = Nothing
End Sub

