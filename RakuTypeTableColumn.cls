VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RakuTypeTableColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public flg As String
Public name_jp As String
Public name_en As String
Public data_type As String
Public data_size As Integer
Public nullable As String

Private Sub Class_Initialize()
    flg = ""
    name_jp = ""
    name_en = ""
    data_type = ""
    data_size = 0
    nullable = ""
End Sub

