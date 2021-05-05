VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RakuConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public errorMsg As String
Public status As Integer

Private con As ADODB.Connection
Private ds As DataSource

Public Sub Connect()
    On Error GoTo ErrorHandler
    Dim Provider As String
    errorMsg = ""
    Set con = CreateObject("ADODB.Connection")
    
    ProviderDrv = "MSDAORA"
    DbSource = "127.0.0.1:1521/orcl"
    Pwd = "mese"
    UserId = "mes"
'    Provider = "Provider=" & ProviderDrv & ";Data Source=" & DbSource & ";Persist Security Info=True;Password=" & Pwd & ";User ID=" & UserId + ";"
'    con.Open "Provider=MSDAORA;Data Source=127.0.0.1:1521/MES;Persist Security Info=True;Password=MESST;User ID=MESST;"
    con.Open "Provider=MSDAORA;Data Source=ORCL;Persist Security Info=True;Password=mes;User ID=mese;"
'    con.Open "DSN=ORCL;UID=mese;PWD=mes;"
'    con.Open _
'      "DRIVER={MySQL ODBC 5.2 Unicode Driver};" & _
'      " SERVER=127.0.0.1;" & _
'      " DATABASE=xiaomimysql;" & _
'      " UID=root;" & _
'      " PWD=admin;"
'    con.ConnectionString = "DSN=xiaomiMysql;UID=root;PWD=admin"
 
'  MsgBox "DB connect success"
    
    Exit Sub
ErrorHandler:
    errorMsg = Err.Description
    status = con.State
    Disconnect
    Err.Raise 1, "RakuConnection", "SQL[select error]." & vbCrLf & errorMsg
End Sub

Public Sub Disconnect()
    con.Close
    Set con = Nothing
End Sub

Public Sub setDataSorce(paramDs As DataSource)
    ds = paramDs
End Sub

Public Sub setDataSorce2(ByVal host As String, ByVal user As String, ByVal passwd As String)
    ds.host = host
    ds.user = user
    ds.passwd = passwd
End Sub

Public Function find(sqlStmt As String) As RakuArrayTable

    On Error GoTo ErrorHandler
    errorMsg = ""
    Dim rs As Object
    Dim recIdx As Long '
    Dim fieldIdx As Integer '
    Dim table As New RakuArrayTable
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlStmt, con
    Do Until rs.EOF
        table.AddRow
        For fieldIdx = 0 To rs.Fields.Count - 1
            unicodeString = getOriginalUnicodeString2(rs.Fields(fieldIdx))
            table.addColumn unicodeString
        Next
        
        recIdx = recIdx + 1
        rs.MoveNext
    Loop
    Set find = table
    Set table = Nothing
    Set table = Nothing
    
    Exit Function
'
    
ErrorHandler:
    errorMsg = Err.Description
    status = con.State
    Disconnect
    Err.Raise 1, "RakuConnection", "SQL[find error]." & vbCrLf & errorMsg
End Function
Private Function getOriginalUnicodeString2(varOriginal As Variant) As String
    
    Dim unicodeString As String
    Dim i As Integer
    
    On Error GoTo skip1:
    unicodeString = varOriginal
    
    Select Case TypeName(varOriginal)
        Case "Byte"
            unicodeString = "&H"
            unicodeString = Format(Hex(CInt(varOriginal)), "00")
        Case "Byte()"
            unicodeString = "&H"
            For i = 0 To UBound(varOriginal)
                unicodeString = unicodeString & Format(Hex(CInt(varOriginal(i))), "00")
            Next
        Case "Field"
            unicodeString = varOriginal
        Case Else
            If unicodeString = "" Then
                unicodeString = ""
            Else
                unicodeString = "'" & unicodeString
            End If
    End Select
    getOriginalUnicodeString2 = unicodeString
    
    Exit Function
    
skip1:
    If IsNull(varOriginal) Then
        getOriginalUnicodeString2 = "" '(null)
    End If

End Function

Public Sub executeSQL(sqlStmt As String)
    On Error GoTo ErrorHandler
    errorMsg = ""
    con.Execute sqlStmt
    Exit Sub
ErrorHandler:
    errorMsg = Err.Description
    status = con.State
    MsgBox errorMsg & vbCrLf
    Disconnect
    MsgBox sqlStmt & vbCrLf
    Err.Raise 1, "RakuConnection", "SQL[executeSQL error]." & vbCrLf & errorMsg
End Sub

Public Sub commit()
    con.CommitTrans
    
End Sub


Public Sub rollback()
    con.RollbackTrans
End Sub


Public Sub begin()
    con.BeginTrans
End Sub

