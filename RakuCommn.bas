Attribute VB_Name = "RakuCommn"
Public Const PYMN_UT_SI_TOOL_USE_DIV = "SI"
Public Const CONFIG_SHT_NAME = "tool-config"
Public Const RTN_OK = 1
Public Const RTN_NG = 8
Public Const ForReading = 1, ForWriting = 2, FoAppending = 3

Public Type DataSource
    host As String
    user As String
    passwd As String
End Type

Public Type MobilServerInfo
    host As String
    user As String
    passwd As String
    sys_user As String
    sys_passwd As String
    mobills_home As String
    env_file As String
    tlgrm_path As String
    tx_run_cmd As String
    tx_mng_cmd As String
    tx_module_name As String
End Type

Public mTblMng As New RakuTableManger
Public mTgrmMng As New RakuTypeTable
Public mFileIOMng As New RakuTableManger
Public mTxManager As New RakuConnection
Public mFtpTransfer As New RakuTableManger
Public mTtlCmdMng As New RakuArrayTable
Public Const ToolBarName As String = "DBTool"


Public strSID As String   '
Public strUser As String   '
Public strPass As String   '

Public strActionType As String

Sub Auto_Open()
    Call CreateMenubar
End Sub
Sub Auto_Close()
    Call RemoveMenubar
End Sub
Sub RemoveMenubar()
    On Error Resume Next
    Application.CommandBars(ToolBarName).delete
    On Error GoTo 0
End Sub
Sub CreateMenubar()
    Dim iCtr As Long

    Dim MacNames As Variant
    Dim CapNamess As Variant
    Dim TipText As Variant

    Call RemoveMenubar

    MacNames = Array("defineTbl", "UpdData", "selData", "insData", "delData", "setSIDInfo")              ' 1

    CapNamess = Array("Define", "Update", "Select", "Insert", "Delete", "Setting")

    TipText = Array("TBL DEF(&A)", "DATA UPD(&U)", "DATA GET(&X)", "ALL DATA UPD(&M)", "ALL DATA GET(&N)", "SET SID INFO(&S)")

    With Application.CommandBars.Add
        .Name = ToolBarName
        .Left = 200
        .Top = 200
        .Protection = msoBarNoProtection
        .Visible = True
        .Position = msoBarFloating

        For iCtr = LBound(MacNames) To UBound(MacNames)
            With .Controls.Add(Type:=msoControlButton)
                .OnAction = "'" & ThisWorkbook.Name & "'!" & MacNames(iCtr)
                .Caption = CapNamess(iCtr)
                .Style = msoButtonIconAndCaption
                .FaceId = 71 + iCtr
                .TooltipText = TipText(iCtr)
            End With
        Next iCtr
    End With
End Sub

Public Sub defineTbl()
    Call mTblMng.refreshExcelTblInfoH
End Sub

Public Sub updData()
    strActionType = "update"
    Call mTblMng.updData
End Sub

Public Sub selData()
    strActionType = "select"
    Call mTblMng.getTblData
End Sub

Public Sub insData()
    strActionType = "insert"
    Call mTblMng.insData
End Sub

Public Sub delData()
    strActionType = "delete"
    Call mTblMng.delData
End Sub

Public Sub setSIDInfo()


    Dim exitFlg
    exitFlg = 0
    For x = 1 To Sheets.Count
        If Sheets(x).Name = "config" Then
            exitFlg = 1
            Exit For
        End If
    Next x
    If exitFlg = 0 Then
        strSID = "MESST"
        strUser = "messt"
        strPass = "mes"
    Else
        strSID = Sheets("config").Range("O1").Value
        strUser = Sheets("config").Range("P1").Value
        strPass = Sheets("config").Range("Q1").Value
    End If
'    UserForm1.Show
End Sub

Public Sub setDBInfo()

    Dim exitFlg
    exitFlg = 0
    For x = 1 To Sheets.Count
        If Sheets(x).Name = "config" Then
            exitFlg = 1
            Exit For
        End If
    Next x
    If exitFlg = 0 Then
        Sheets.Add
        ActiveSheet.Name = "config"
    End If
    Sheets("config").Range("O1").Value = strSID
    Sheets("config").Range("P1").Value = strUser
    Sheets("config").Range("Q1").Value = strPass
    Sheets("config").Visible = False
End Sub

Public Sub encrypt()
    call_aa2 Selection, "e"
End Sub

Public Sub decrypt()
    call_aa2 Selection, "d"
End Sub

Private Sub call_aa2(cells As Range, mode As String)
    Set ie = CreateObject("InternetExplorer.Application")
    For Each c In cells
        ie.navigate "http://129.172.208.159/cgi-bin/aa2.pl?intext=" & c.Value & ":opt=-" & mode
        While ie.ReadyState <> 4
            While ie.Busy = True
                DoEvents
            Wend
        Wend
        c.Value = ie.Document.ALL.item(21).Value
    Next
End Sub






