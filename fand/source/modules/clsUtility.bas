Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private fso As Scripting.FileSystemObject

Private Sub Class_Initialize()

    Set fso = New Scripting.FileSystemObject
    
End Sub

Public Sub AppendTable(Connection As ADODB.Connection, ByVal tableName As String, ByVal TablePath As String)

On Error GoTo Err_Handler

    Dim i As Long
    Dim l As Long
    Dim str As String
    Dim lng As Long
    Dim DataFile As Scripting.TextStream
    Dim strColumn() As String
    Dim strRow() As String
    Dim strValue As String
    Dim rst As ADODB.Recordset
    
    '--Load the file contents into a string
    Set DataFile = fso.OpenTextFile(TablePath, ForReading, False)
    str = DataFile.ReadAll()
    DataFile.Close
    Set DataFile = Nothing

    '--Split the string into an array of rows that are caret delimited
    strRow = Split(str, vbCrLf, , vbTextCompare)
    
    '--Create the recordset for this file
    Set rst = New ADODB.Recordset
    rst.Open tableName, Connection, adOpenStatic, adLockOptimistic, adCmdTable

    For i = 0 To UBound(strRow)
        If Len(strRow(i)) Then
            strColumn = Split(strRow(i), "^", , vbTextCompare)
            rst.AddNew
            For l = 0 To rst.Fields.Count - 2
                If l <= UBound(strColumn) Then
                    If Len(strColumn(l)) Then
                        If rst.Fields(l).name = "SRCode" Then
                            If StrComp(strColumn(l), "~~", vbTextCompare) <> 0 Then
                                strValue = Replace(strColumn(l), "~", vbNullString, 1, 2, vbTextCompare)
                                strValue = Trim$(strValue)
                                If Len(strValue) Then
                                    Select Case Len(strValue)
                                        Case Is > 4
                                        Case Else
                                            strValue = String(5 - Len(strValue), "0") & strValue
                                    End Select
                                    rst.Fields(l).Value = strValue
                                End If
                            End If
                        Else
                            If rst.Fields(l).Type = adVarChar Or rst.Fields(l).Type = adVarWChar Then
                                If StrComp(strColumn(l), "~~", vbTextCompare) <> 0 Then
                                    strValue = Replace(strColumn(l), "~", vbNullString, 1, 2, vbTextCompare)
                                    strValue = Trim$(strValue)
                                    If Len(strValue) Then
                                        rst.Fields(l).Value = strValue
                                    End If
                                End If
                            Else
                                rst.Fields(l).Value = strColumn(l)
                            End If
                        End If
                    End If
                End If
            Next l
            rst.Update
        End If
    Next i
    
    rst.Close
    Set rst = Nothing

    '--Erase the arrays
    Erase strColumn
    Erase strRow

Exit_Sub:
    Exit Sub
Err_Handler:
    Debug.Print tableName, strRow(i), Err.Number, Err.Description
    rst.CancelUpdate
    Resume Next
        
End Sub

Public Sub AppendTableSpecial1(Connection As ADODB.Connection, tableName As String, DataName As String, _
    DataPath As String, SchemaName As String, SchemaPath As String)

On Error GoTo Err_Handler

    Dim l As Long
    Dim lng As Long
    Dim str As String
    Dim strSchema() As String
    Dim strColumn() As String
    Dim strValue As String
    Dim DataFile As Scripting.TextStream
    Dim rst As ADODB.Recordset
    Dim appExcel As Excel.Application
    Dim wbkExcel As Excel.Workbook
    Dim wstExcel As Excel.Worksheet
    
    Set appExcel = New Excel.Application
    With appExcel
        Set wbkExcel = .Workbooks.Open(fso.BuildPath(SchemaPath, SchemaName))
        .Visible = True
    End With
    wbkExcel.Activate
    Set wstExcel = wbkExcel.Worksheets("Sheet1")
    wstExcel.Activate
    
    l = 2
    With wstExcel
        Do Until Len(.Range("A" & CStr(l)).Value) = 0
            .Range("A" & CStr(l)).Activate
            ReDim Preserve strSchema(ArrayIndex(strSchema))
            strSchema(UBound(strSchema)) = Trim$(.Range("A" & CStr(l)).Value) & "^" & _
                Trim$(.Range("B" & CStr(l)).Value) & "^" & _
                Trim$(.Range("D" & CStr(l)).Value)
            l = l + 1
        Loop
    End With
    
    Set wstExcel = Nothing
    If Not (wbkExcel Is Nothing) Then wbkExcel.Close
    Set wbkExcel = Nothing
    If Not (appExcel Is Nothing) Then appExcel.Quit
    Set appExcel = Nothing
    
    '--Open the file
    Set DataFile = fso.OpenTextFile(fso.BuildPath(DataPath, DataName), ForReading, False)

    '--Create the recordset for this file
    Set rst = New ADODB.Recordset
    rst.Open tableName, Connection, adOpenStatic, adLockPessimistic, adCmdTable
    
    Do While Not DataFile.AtEndOfStream
        str = DataFile.ReadLine
        If Len(str) > 0 Then
            rst.AddNew
            For l = 0 To rst.Fields.Count - 2
                strColumn = Split(strSchema(l), "^", 3, vbTextCompare)
                With rst.Fields(l)
                    If .name = strColumn(0) Then
                        strValue = Mid$(str, strColumn(1), strColumn(2))
                        strValue = Trim$(strValue)
                        If Len(strValue) Then .Value = strValue
                    Else
                        Stop
                    End If
                End With
                Erase strColumn()
            Next l
            rst.Update
        End If
    Loop

    rst.Close
    Set rst = Nothing

    '--Close file
    DataFile.Close
    Set DataFile = Nothing
    
    Erase strSchema
        
Exit_Sub:
    Exit Sub
Err_Handler:
    Debug.Print tableName, strValue, Err.Number, Err.Description
    GoTo Exit_Sub
        
End Sub

Public Sub AppendTableSpecial2(Connection As ADODB.Connection, ByVal tableName As String, ByVal fileName As String, _
    ByVal filePath As String, SchemaName As String, SchemaPath As String)

On Error GoTo Err_Handler

    Dim l As Long
    Dim lng As Long
    Dim str As String
    Dim strSchema() As String
    Dim strColumn() As String
    Dim strValue() As String
    Dim DataFile As Scripting.TextStream
    Dim rst As ADODB.Recordset
    Dim appExcel As Excel.Application
    Dim wbkExcel As Excel.Workbook
    Dim wstExcel As Excel.Worksheet
    
    Set appExcel = New Excel.Application
    With appExcel
        Set wbkExcel = .Workbooks.Open(fso.BuildPath(SchemaPath, SchemaName))
        .Visible = True
    End With
    wbkExcel.Activate
    Set wstExcel = wbkExcel.Worksheets(Mid$(tableName, 4))
    wstExcel.Activate
    
    l = 2
    With wstExcel
        Do Until Len(.Range("A" & CStr(l)).Value) = 0
            .Range("A" & CStr(l)).Activate
            If StrComp(Trim$(.Range("I" & CStr(l)).Value), "Y", vbTextCompare) = 0 Then
                ReDim Preserve strSchema(ArrayIndex(strSchema))
                strSchema(UBound(strSchema)) = Trim$(.Range("B" & CStr(l)).Value) & "^" & _
                    Trim$(.Range("D" & CStr(l)).Value) & "^" & _
                    Trim$(.Range("F" & CStr(l)).Value)
            End If
            l = l + 1
        Loop
    End With
    
    Set wstExcel = Nothing
    If Not (wbkExcel Is Nothing) Then wbkExcel.Close
    Set wbkExcel = Nothing
    If Not (appExcel Is Nothing) Then appExcel.Quit
    Set appExcel = Nothing
    
    '--Open the text file
    Set DataFile = fso.OpenTextFile(fso.BuildPath(filePath, fileName))

    '--Create the recordset for this file
    Set rst = New ADODB.Recordset
    rst.Open tableName, Connection, adOpenStatic, adLockPessimistic, adCmdTable
    
    str = DataFile.ReadLine
    Do While Not DataFile.AtEndOfStream
        str = vbNullString
        str = DataFile.ReadLine
        If Len(str) Then
            rst.AddNew
            strValue = Split(str, "^", , vbTextCompare)
            For l = 0 To rst.Fields.Count - 2
                strColumn = Split(strSchema(l), "^", 3, vbTextCompare)
                With rst.Fields(l)
                    If .name = strColumn(0) Then
                        If Len(strValue(l)) Then
                            If StrComp(strValue(l), ".", vbTextCompare) <> 0 Then
                                .Value = strValue(l)
                            End If
                        End If
                    Else
                        Stop
                    End If
                End With
                Erase strColumn()
            Next l
            rst.Update
            Erase strValue()
        End If
    Loop

    rst.Close
    Set rst = Nothing

    '--Close file
    DataFile.Close
    Set DataFile = Nothing
    
    Erase strSchema()
        
Exit_Sub:
    Exit Sub
Err_Handler:
    Debug.Print tableName, strColumn(0), strValue(l), Err.Number, Err.Description
    Resume Next
        
End Sub

Public Function ArrayIndex(var As Variant) As Long

On Error GoTo Err_Handler

    ArrayIndex = UBound(var) + 1

Exit_Function:
    Exit Function
Err_Handler:
    ArrayIndex = 0

End Function

Public Function EscapedString(Text As String) As String

    EscapedString = Replace(Replace(Trim$(Text), "'", "\'"), """", "\""")

End Function

Private Sub Class_Terminate()

    Set fso = Nothing

End Sub