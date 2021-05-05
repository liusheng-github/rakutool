Attribute VB_Name = "RakuUtility"
Public Function ActiveShtCellVal(ByVal rowIdx As Integer, ByVal colIdx As Integer)
    ActiveShtCellVal = Trim(ActiveSheet.cells(rowIdx, colIdx))
End Function

Public Function selInFilePath() As String
    Dim inFullPath As String
    inFullPath = Application.GetOpenFilename("ALL File(*.*), *.*")
    If inFullPath = "false" Then
        inFullPath = ""
    End If
    
    selInFilePath = inFullPath
End Function

Public Function selPath(Optional title As String = "Missing", Optional rootPath As Variant) As String
    Dim shl As Object
    Dim fld As Object
    Dim strPath As String
    Dim ttl As String
    
    If title = "Missing" Then
        ttl = "please select folder"
    Else
        ttl = title
    End If
    
    Set shl = CreateObject("Shell.Aplication")
    If IsMissing("roogPath") Then
        Set fld = shl.browseforfolder(0, ttl, 1 + 512)
    Else
        Set fld = shl.browseforfolder(0, ttl, 1 + 512, rootPath)
    End If
    
    strPath = ""
    If Not fld Is Nothing Then
        On Error Resume Next
        If strPath = "" Then
            strPath = fld.items.item.Path
        End If
        On Error GoTo 0
    End If
    
    If InStr(strPath, "\") = 0 Then
        strPath = ""
    End If
    
    selPath = strPath
    Set fld = Nothing
    Set shl = Nothing
    
End Function

Public Function parseItemVal(ByVal itmVal As String, _
                            ByRef startSqlNo As Long, _
                            ByRef endSeqlNo As Long) As String
    pos1 = InStr(itmVal, "[")
    pos2 = InStr(itmVal, "~")
    pos3 = InStr(itmVal, "]")
    If pos1 > 0 And pos2 > pos1 And pos3 > pos2 Then
        seqNoLen = pos2 - pos1 - 1
        strTemp = Mid(itmVal, pos1 + 1, seqNoLen)
        startSeqNo = val(strTemp)
        
        strTemp = Mid(itmVal, pos2 + 1, seqNoLen)
        startSeqNo = val(strTemp)
        
        parseItemVal = Mid(itmVal, 1, pos1 - 1) & String(seqNoLen, "\") & Mid(itmVal, pos3 + 1)
    Else
        parseItemVal = itmVal
        startSeqlNo = -1
        endSeqNo = -1
    End If
End Function

Public Function getStrLenB(str As String) As Integer
    getStrLenB = LenB(StrConv(str, vbFromUnicode))
End Function













