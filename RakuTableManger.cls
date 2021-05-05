VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RakuTableManger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const ROW_MAX = 1500

Const COL_IDX_TBL_NAME = 1
Const COL_IDX_TBL_NAME_JP = 2
Const COL_IDX_TBL_DB = 3
Const COL_IDX_TBL_OUTFLG = 4
Const COL_IDX_TBL_SEL_COND = 5
Const COL_IDX_TBL_SEL_ACTION = 6
Const COL_IDX_TBL_SEL_ORDER = 7

Const ITM_NAME_JP_ROWIDX = 1
Const ITM_NAME_EN_ROWIDX = 2
Const ITM_TYPE_ROWIDX = 3
Const ITM_SIZE_ROWIDX = 4
Const ITM_NULLABLE_ROWIDX = 5
Const ITM_FLG_ROWIDX = 6
Const ITM_ROW_CNT = 6

Private mCon As New RakuConnection

Private Function getTblDefInfo(ByVal tblName As String) As RakuTypeTable
    Dim sqlStmt As String
    Dim tblInfoArr As RakuArrayTable
    Dim tblColumnItem As RakuTypeTableColumn
    
    Set getTblDefInfo = New RakuTypeTable
    
    tblName = UCase(tblName)
    sqlStmt = "select " & _
                " b.COMMENTS, a.COLUMN_NAME, a.DATA_TYPE, a.DATA_LENGTH, a.DATA_PRECISION, a.NULLABLE " & _
                " from " & _
                "user_tab_columns a, user_col_comments b, " & _
                "(select x.TABLE_NAME " & _
                "from USER_SYNONYMS x " & _
                "where x.SYNONYM_NAME='" & tblName & "' union " & _
                "select y.TABLE_NAME " & _
                "from USER_TAB_COLUMNS y " & _
                "where y.TABLE_NAME='" & tblName & "') c " & _
                "where a.TABLE_NAME=c.TABLE_NAME " & _
                "and b.TABLE_NAME=a.TABLE_NAME and b.COLUMN_NAME=a.COLUMN_NAME " & _
                "order by a.COLUMN_ID"

      Set tblInfoArr = mCon.find(sqlStmt)
      
      getTblDefInfo.tblName = tblName
      
      For i = 1 To tblInfoArr.Nrows
        Set tblColumnItem = New RakuTypeTableColumn
        If i < 3 Then
            tblColumnItem.flg = "¡ø"
        Else
            tblColumnItem.flg = ""
        End If
        tblColumnItem.name_jp = Trim(UCase(tblInfoArr.Cell(i, 1)))
        tblColumnItem.name_en = Trim(UCase(tblInfoArr.Cell(i, 2)))
        tblColumnItem.data_type = Trim(UCase(tblInfoArr.Cell(i, 3)))
        tblColumnItem.data_size = Trim(UCase(tblInfoArr.Cell(i, 4)))
        tblColumnItem.nullable = Trim(UCase(tblInfoArr.Cell(i, 6)))
        
        getTblDefInfo.columns.Add tblColumnItem
     Next
     
     Call tblInfoArr.clear
     Set tblInfoArr = Nothing
                
End Function


Private Function getExcelTblInfo(ByVal rowIdx As Integer) As RakuTypeTable
    Dim tblValRecord As RakuTypeTableRecord
    Dim tblColumnItem As RakuTypeTableColumn
    Dim tblValRecordItem As RakuTypeTableRecordItem
    Dim seqStart As Long, seqEnd As Long
    
    Set getExcelTblInfo = New RakuTypeTable
    
    getExcelTblInfo.tblName = ActiveShtCellVal(rowIdx, COL_IDX_TBL_NAME)
    getExcelTblInfo.dbName = ActiveShtCellVal(rowIdx, COL_IDX_TBL_DB)
    getExcelTblInfo.selCond = ActiveShtCellVal(rowIdx, 5)
    getExcelTblInfo.action = UCase(ActiveShtCellVal(rowIdx, 6))
    getExcelTblInfo.selOrder = ActiveShtCellVal(rowIdx, 7)
    
    colIdx = 2
    While Trim(ActiveSheet.cells(rowIdx + 3, colIdx)) <> ""
        Set tblColumnItem = New RakuTypeTableColumn
        tblColumnItem.name_jp = ActiveShtCellVal(rowIdx + 1, colIdx)
        tblColumnItem.name_en = ActiveShtCellVal(rowIdx + 2, colIdx)
        tblColumnItem.data_type = ActiveShtCellVal(rowIdx + 3, colIdx)
        tblColumnItem.data_size = ActiveShtCellVal(rowIdx + 4, colIdx)
        tblColumnItem.nullable = ActiveShtCellVal(rowIdx + 5, colIdx)
        tblColumnItem.flg = ActiveShtCellVal(rowIdx + 6, colIdx)
        
        getExcelTblInfo.columns.Add tblColumnItem
        colIdx = colIdx + 1
    Wend
    
    colIdx = 2
    If strActionType = "select" Then
        While Trim(ActiveSheet.cells(rowIdx + 7, colIdx)) <> ""
            Set tblValRecord = New RakuTypeTableRecord
            For colIdx = 1 To getExcelTblInfo.columns.Count
                Set tblValRecordItem = New RakuTypeTableRecordItem
                tblValRecordItem.val = parseItemVal(ActiveSheet.cells(rowIdx + 7, colIdx + 1), seqStart, seqEnd)
                tblValRecordItem.val_start = seqStart
                tblValRecordItem.val_end = seqEnd
                tblValRecordItem.val_cur = seqStart

                tblValRecord.item_lst.Add tblValRecordItem

                If tblValRecord.rec_max_cnt < (tblValRecordItem.val_end - tblValRecordItem.val_start + 1) Then
                    tblValRecord.rec_max_cnt = tblValRecordItem.val_end - tblValRecordItem.val_start + 1
                End If
            Next
            getExcelTblInfo.values.Add tblValRecord
            rowIdx = rowIdx + 1
        Wend
    Else
        For rowIdx = Selection.ROW To Selection.ROW + Selection.rows.Count - 1
            If Trim(ActiveSheet.cells(rowIdx, colIdx)) <> "" Then
                Set tblValRecord = New RakuTypeTableRecord
                For colIdx = 1 To getExcelTblInfo.columns.Count
                    Set tblValRecordItem = New RakuTypeTableRecordItem
                    tblValRecordItem.val = parseItemVal(ActiveSheet.cells(rowIdx, colIdx + 1), seqStart, seqEnd)
                    tblValRecordItem.val_start = seqStart
                    tblValRecordItem.val_end = seqEnd
                    tblValRecordItem.val_cur = seqStart

                    tblValRecord.item_lst.Add tblValRecordItem

                    If tblValRecord.rec_max_cnt < (tblValRecordItem.val_end - tblValRecordItem.val_start + 1) Then
                        tblValRecord.rec_max_cnt = tblValRecordItem.val_end - tblValRecordItem.val_start + 1
                    End If
                Next
                getExcelTblInfo.values.Add tblValRecord
            End If
        Next
    End If
    
    Exit Function
End Function

Public Sub refreshExcelTblInfoH()
    Dim tblExcInfo As RakuTypeTable
    Dim tblDefInfo As RakuTypeTable
    Dim dbName As String
    Dim rowIdx As Integer
    Dim ds As DataSource
    
    For Each c In Selection
        Call mCon.Connect
        
        rowIdx = c.Offset(0, 0).ROW
        Set tblExcInfo = getExcelTblInfo(rowIdx)
        
        tblEnName = Trim(c.Offset(0, 0).Value)
        Set tblDefInfo = getTblDefInfo(tblEnName)
        
        colOff = 0
        While Trim(c.Offset(2, colOff).Value) <> ""
            c.Offset(ITM_NAME_JP_ROWIDX, colOff).clear
            c.Offset(ITM_NAME_EN_ROWIDX, colOff).clear
            c.Offset(ITM_TYPE_ROWIDX, colOff).clear
            c.Offset(ITM_SIZE_ROWIDX, colOff).clear
            c.Offset(ITM_NULLABLE_ROWIDX, colOff).clear
            c.Offset(ITM_FLG_ROWIDX, colOff).clear
            colOff = colOff + 1
        Wend

        For i = 1 To tblDefInfo.columns.Count
            c.Offset(ITM_NAME_JP_ROWIDX, i) = tblDefInfo.columns.item(i).name_jp
            c.Offset(ITM_NAME_EN_ROWIDX, i) = tblDefInfo.columns.item(i).name_en
            c.Offset(ITM_TYPE_ROWIDX, i) = tblDefInfo.columns.item(i).data_type
            c.Offset(ITM_SIZE_ROWIDX, i) = tblDefInfo.columns.item(i).data_size
            c.Offset(ITM_NULLABLE_ROWIDX, i) = tblDefInfo.columns.item(i).nullable
            c.Offset(ITM_FLG_ROWIDX, i) = tblDefInfo.columns.item(i).flg
        Next i
        Range(cells(rowIdx + 1, 2), cells(rowIdx + 1, tblDefInfo.columns.Count + 1)).Interior.ColorIndex = 37
        Range(cells(rowIdx + 2, 2), cells(rowIdx + 6, tblDefInfo.columns.Count + 1)).Interior.ColorIndex = 34
        Range(cells(rowIdx + 1, 2), cells(rowIdx + 6, tblDefInfo.columns.Count + 1)).Borders.LineStyle = xlContinuous
        
        
        Call mCon.Disconnect
        Set tblExcInfo = Nothing
        Set tblDefInfo = Nothing
    Next
            
End Sub

Public Sub updData()
    Dim tblInfo As RakuTypeTable
    Dim ds As DataSource
    Dim rowIdx As Integer
    
    On Error GoTo ErrHandling
    For rowIdx = Selection.ROW To 1 Step -1
        If Trim(ActiveSheet.cells(rowIdx, COL_IDX_TBL_NAME)) <> "" Then
            Set tblInfo = getExcelTblInfo(rowIdx)
            If tblInfo.action = "¡ð" Then
                Call mCon.Connect
                
                Call delete(tblInfo)
                
                Call insert(tblInfo)
                
                Call mCon.Disconnect
                
                Exit For
            End If
            
            Set tblInfo = Nothing
        End If
    Next rowIdx
    
    MsgBox ActiveSheet.Name & "  Update completed"
    Exit Sub
ErrHandling:
    MsgBox Err.Description
    Call mCon.Disconnect
    Set tblInfo = Nothing
End Sub

Public Sub delData()
    Dim tblInfo As RakuTypeTable
    Dim ds As DataSource
    Dim rowIdx As Integer
    
    On Error GoTo ErrHandling
    For rowIdx = Selection.ROW To 1 Step -1
        If Trim(ActiveSheet.cells(rowIdx, COL_IDX_TBL_NAME)) <> "" Then
            Set tblInfo = getExcelTblInfo(rowIdx)
            If tblInfo.action = "¡ð" Then
                Call mCon.Connect
                
                Call delete(tblInfo)
                
                Call mCon.Disconnect
                
                Exit For
            End If
            
            Set tblInfo = Nothing
        End If
    Next rowIdx
    
    MsgBox ActiveSheet.Name & "  Delete completed"
    Exit Sub
ErrHandling:
    MsgBox Err.Description
    Call mCon.Disconnect
    Set tblInfo = Nothing
End Sub
Public Sub insData()
    Dim tblInfo As RakuTypeTable
    Dim ds As DataSource
    Dim rowIdx As Integer
    
    On Error GoTo ErrHandling
    For rowIdx = Selection.ROW To 1 Step -1
        If Trim(ActiveSheet.cells(rowIdx, COL_IDX_TBL_NAME)) <> "" Then
            Set tblInfo = getExcelTblInfo(rowIdx)
            If tblInfo.action = "¡ð" Then
                Call mCon.Connect
                
                Call insert(tblInfo)
                
                Call mCon.Disconnect
                
                Exit For
            End If
            
            Set tblInfo = Nothing
        End If
    Next rowIdx
    
    MsgBox ActiveSheet.Name & "  Insert completed"
    Exit Sub
ErrHandling:
    MsgBox Err.Description
    Call mCon.Disconnect
    Set tblInfo = Nothing
End Sub



Private Sub delete(ByRef tbl As RakuTypeTable)
    Dim strSql As String
    Dim strItemVal As String
    Dim tblValRecord As RakuTypeTableRecord
    Dim tblValRecordItem As RakuTypeTableRecordItem
    Dim tblColumnItem As RakuTypeTableColumn
    
    mCon.begin
    
    For i = 1 To tbl.values.Count
        Set tblValRecord = tbl.values.item(i)
        
        strCond = ""
        For j = 1 To tbl.columns.Count
            Set tblColumnItem = tbl.columns.item(j)
            If tblColumnItem.flg <> "" Then
                If strCond <> "" Then
                    strCond = strCond & " AND "
                End If
                
                Set tblValRecordItem = tblValRecord.item_lst.item(j)
                If tblValRecordItem.val_start < 0 Then
                    strCond = strCond & tblColumnItem.name_en & "='" & tblValRecordItem.val & "' "
                Else
                    strItemVal = tblValRecordItem.val
                    strItemVal = Replace(strItemVal, "#######", "%")
                    strItemVal = Replace(strItemVal, "######", "%")
                    strItemVal = Replace(strItemVal, "#####", "%")
                    strItemVal = Replace(strItemVal, "####", "%")
                    strItemVal = Replace(strItemVal, "###", "%")
                    strItemVal = Replace(strItemVal, "##", "%")
                    strItemVal = Replace(strItemVal, "#", "%")
                    
                    strCond = strCond & tblColumnItem.name_en & " = '" & strItemVal & "'"
                End If
            End If
        Next j
        
        If Trim(strCond) <> "" Then
            strSql = "DELETE FROM " & tbl.tblName & " where " & strCond
            mCon.executeSQL strSql
        End If
    Next i
        
    mCon.commit
        
End Sub


Private Sub insert(ByRef tbl As RakuTypeTable)

    Dim seqNo As Long
    Dim tblValRecord As RakuTypeTableRecord
    Dim tblValRecordItem As RakuTypeTableRecordItem
    Dim strSql As String, strTmpSQL As String
    
    columnNames = ""
    For i = 1 To tbl.columns.Count
        If columnNames <> "" Then
            columnNames = columnNames & ","
        End If
        columnNames = columnNames & tbl.columns.item(i).name_en
    Next
    
    For i = 1 To tbl.values.Count
        Set tblValRecord = tbl.values.item(i)
        If tblValRecord.rec_max_cnt > 10000 Then
            If MsgBox(tblName & "exceed 10000") = vbNo Then
                Exit Sub
            End If
        End If
            
        mCon.begin
        For j = 1 To tblValRecord.rec_max_cnt
            itmVals = ""
            For k = 1 To tblValRecord.item_lst.Count
                itmVal = tblValRecord.item_lst.item(k).val
                seqNo = tblValRecord.item_lst.item(k).val_cur
                    
                itemVal = Replace(itemVal, "#####", Format(seqNo Mod 100000, "00000"))
                itemVal = Replace(itemVal, "####", Format(seqNo Mod 10000, "0000"))
                itemVal = Replace(itemVal, "###", Format(seqNo Mod 1000, "000"))
                itemVal = Replace(itemVal, "##", Format(seqNo Mod 100, "00"))
                itemVal = Replace(itemVal, "#", Format(seqNo Mod 10, "0"))
                If seqNo + 1 > tblValRecord.item_lst.item(k).val_end Then
                    tblValRecord.item_lst.item(k).val_cur = tblValRecord.item_lst.item(k).val_start
                Else
                    tblValRecord.item_lst.item(k).val_cur = tblValRecord.item_lst.item(k).val_cur + 1
                End If
                    
                If itmVals <> "" Then
                    itmVals = itmVals & ","
                End If
                If tbl.columns.item(k).data_type = "TIMESTAMP(6)" Then
                    itmVals = itmVals & "to_date('" & itmVal & "','yyyy-mm-dd HH24-mi-ss')"
                ElseIf tbl.columns.item(k).data_type = "DATE" Then
                    itmVals = itmVals & "to_date('" & itmVal & "','yyyy/mm/dd')"
                Else
                    itmVals = itmVals & "'" & itmVal & "'"
                End If
            Next
                
            strSql = "INSERT INTO " & tbl.tblName & "(" & columnNames & ") VALUES( " & itmVals & ")"
                
            mCon.executeSQL strSql
                 
            If (j Mod 2000) = 0 Then
                mCon.commit
                mCon.begin
            End If
        Next
        mCon.commit
    Next
            
End Sub

Public Sub getTblData()
    Dim tblInfo As RakuTypeTable
    Dim ds As DataSource
    Dim rowIdx As Integer, tmpRowIdx As Integer
    Dim dataArr As RakuArrayTable

    For rowIdx = 1 To ROW_MAX
        If Trim(ActiveSheet.cells(rowIdx, COL_IDX_TBL_NAME)) <> "" And Trim(ActiveSheet.cells(rowIdx, COL_IDX_TBL_SEL_ACTION)) = "¡ð" Then
            Set tblInfo = getExcelTblInfo(rowIdx)
            Set dataArr = New RakuArrayTable

                Call mCon.Connect
                
                Set dataArr = Nothing
                Set dataArr = getOneTblData(tblInfo)
                
                If rowIdx + ITM_ROW_CNT + dataArr.Nrows > ROW_MAX Then
                    Err.Raise 1, "RakuTableManager err"
                End If
                
                emptyRowCnt = 0
                For emptyRowIdx = rowIdx + ITM_ROW_CNT + tblInfo.values.Count + 1 To ROW_MAX
                    If Trim(ActiveSheet.cells(emptyRowIdx, COL_IDX_TBL_NAME)) = "" Then
                        emptyRowCnt = emptyRowCnt + 1
                    Else
                        Exit For
                    End If
                Next
                
                If emptyRowCnt - dataArr.Nrows < 2 Then
                    For i = 1 To dataArr.Nrows - emptyRowCnt + 2
                        ActiveSheet.rows(rowIdx + ITM_ROW_CNT + tblInfo.values.Count + 2).insert
                    Next
                End If
                
                For i = 1 To dataArr.Nrows
                    addDataRowId = rowIdx + ITM_ROW_CNT + tblInfo.values.Count + i
                    For j = 1 To dataArr.Ncolumns
                        ActiveSheet.cells(addDataRowId, j + 1) = "'" & dataArr.Cell(i, j)
                    Next
                    If i = 1 Then
                        Range(cells(addDataRowId, 2), cells(addDataRowId, dataArr.Ncolumns + 1)).Interior.ColorIndex = 17
                    End If
                Next
                
                Call mCon.Disconnect
                
            rowIdx = rowIdx + ITM_ROW_CNT + tblInfo.values.Count + dataArr.Nrows
            
            Set tblInfo = Nothing
            Set dataArr = Nothing
        End If
    Next rowIdx
    
    MsgBox ActiveSheet.Name & " Get Data Completed"
    Exit Sub
    
ErrHandling:
    MsgBox Err.Description
    Call mCon.Disconnect
    Set tblInfo = Nothing
    Set dataArr = Nothing
                
End Sub


Private Function getOneTblData(ByRef tbl As RakuTypeTable) As RakuArrayTable

    Dim columnNames As String, strSql As String
    columnNames = ""
    For i = 1 To tbl.columns.Count
        If columnNames <> "" Then
            columnNames = columnNames & ","
        End If
        If tbl.columns.item(i).data_type = "TIMESTAMP(6)" Then
            tmpStr = tbl.columns.item(i).name_en
            columnNames = columnNames & "to_char(" & tmpStr & ",'yyyy-mm-dd HH24-mi-ss')"
        ElseIf tbl.columns.item(i).data_type = "DATE" Then
            tmpStr = tbl.columns.item(i).name_en
            columnNames = columnNames & "to_char(" & tmpStr & ",'yyyy/mm/dd')"
        Else
            columnNames = columnNames & tbl.columns.item(i).name_en
        End If
    Next
       
    If tbl.selCond <> "" Then
        strSql = "SELECT " & columnNames & " FROM " & tbl.tblName & " where " & tbl.selCond
    Else
        strSql = "SELECT " & columnNames & " FROM " & tbl.tblName
    End If
       
    If tbl.selOrder <> "" Then
        strSql = strSql & " order by " & tbl.selOrder
    End If
    
    Set getOneTblData = mCon.find(strSql)
    
End Function


Private Sub Class_Terminate()
    Call mCon.Disconnect
    Set mCon = Nothing
End Sub


Public Function getDataSource2(ByRef ds As DataSource) As String
    
    On Error GoTo ErrorHandler
    
    ds.host = Trim(Sheets("config").Range("O1").Value)
    ds.user = Trim(Sheets("config").Range("P1").Value)
    ds.passwd = Trim(Sheets("config").Range("Q1").Value)
    Exit Function
ErrorHandler:
    Err.Raise 1, "Please Set DB Info"
    
End Function
