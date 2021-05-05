VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RakuArrayTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private list As Collection
Private CurrentList As Collection

Public Sub Class_Initialize()
    Set list = New Collection
    
End Sub

Public Sub addColumn(obj As Variant)
    list(list.Count).Add obj
End Sub

Public Sub AddRow(Optional ByVal obj As Variant)
    
    If IsMissing(obj) Then
        list.Add New Collection
    Else
        list.Add obj
    End If
    Set CurrentList = list(list.Count)
End Sub

Public Sub clear()
    Set list = Nothing
    Set CurrentList = Nothing
    Set list = New Collection
End Sub

Public Sub Concat(table As RakuArrayTable)
    For r = 1 To table.Nrows
        AddRow
        For c = 1 To table.Ncolumns
            addColumn table.Cell(r, c)
        Next c
    Next r
End Sub

Public Property Get Cell(ROW, col)
    Cell = list(ROW)(col)
End Property

Public Property Get CurrentRecord()
    CurrentRecord = ToArray(CurrentList)
End Property

Private Function ToArr(l As Collection)
    Dim Result As Variant
    ReDim Result(1 To l.Count)
    For i = 1 To l.Count
        Result(i) = l(i)
    Next
    ToArr = Result
End Function

Public Property Get Nrows()
    Nrows = list.Count
End Property

Public Property Get Ncolumns()
    Ncolumns = CurrentList.Count
End Property

Public Property Get columns(col)
    Dim Result
    ReDim Result(1 To list.Count)
    For i = 1 To list.Count
        Result(i) = list(i)(col)
    Next
    columns = Result
End Property

Public Function hasItemCurrentRow(item)
    For Each c In CurrentList
        If c = item Then
            hasItemCurrentRow = True
            Exit Function
        End If
    Next
    hasItemCurrentRow = False
End Function

Public Property Get rows(ROW)
    rows = ToArr(list(ROW))
End Property

Public Property Get ToArray()
    Dim Result As Variant
    ReDim Result(1 To Nrows, 1 To Ncolumns)
    For i = 1 To Nrows
        For j = 1 To Ncolumns
            Result(i, j) = rows(i)(j)
        Next j
    Next i
    ToArray = Result
End Property

