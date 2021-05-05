VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RakuTypeTableRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public item_lst As New Collection
Public rec_max_cnt As Long

Private Sub Class_Initizlize()
    rec_max_cnt = 1
End Sub

Private Sub Class_Terminate()
    For i = 1 To item_lst.Count
        item_lst.Remove 1
    Next
    Set item_lst = Nothing
End Sub

