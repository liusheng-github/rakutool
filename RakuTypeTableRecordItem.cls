VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RakuTypeTableRecordItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public val As String
Public val_start As String
Public val_end As Long
Public val_cur As Long

Private Sub Class_Initialize()
    val = ""
    val_start = -1
    val_end = -1
    val_cur = -1
End Sub

