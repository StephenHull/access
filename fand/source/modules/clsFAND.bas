Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const DATABASES_PATH As String = "E:\projects\fand\databases"
Private Const SQL_PATH As String = "E:\projects\fand\databases\fand\sql"

Public Enum FNDDSVersionNumber
    fvnFNDDS1 = 1
    fvnFNDDS2 = 2
    fvnFNDDS3 = 4
    fvnFNDDS4 = 8
    fvnFNDDS5 = 16
    fvnFNDDS6 = 32
    fvnFNDDS7 = 64
End Enum

Private fso As Scripting.FileSystemObject
Private Utility As clsUtility

Private cnnBack As ADODB.Connection
Private cnnFNDDS As ADODB.Connection
'Private cnnMPED As ADODB.Connection
Private cnnSR As ADODB.Connection

Private comAddtlDescr_Lkp As ADODB.Command
Private comCountInDocument_Lkp As ADODB.Command
Private comCountInDocuments_Lkp As ADODB.Command
Private comDocumentCount_Lkp As ADODB.Command
Private comFCDescr_Lkp As ADODB.Command
'Private comFoodDescr_Lkp As ADODB.Command
Private comFoodMatrixA_Lkp As ADODB.Command
Private comFoodMatrixB_Lkp As ADODB.Command
Private comFoodMatrixValue_Lkp As ADODB.Command
'Private comIngredients_Lkp As ADODB.Command
'Private comIngredRecipe_Lkp As ADODB.Command
''Private comIngredSearch_Lkp As ADODB.Command
'Private comModNutrient_Lkp As ADODB.Command
'Private comMPED_Lkp As ADODB.Command
'Private comNutrient_Lkp As ADODB.Command
Private comPortionDescr_Lkp As ADODB.Command
Private comPortions_Lkp As ADODB.Command
Private comRecipeWeight_Lkp As ADODB.Command
Private comRetDescr_Lkp As ADODB.Command
'Private comSimilarRecipe_Lkp As ADODB.Command
Private comSRDescr_Lkp As ADODB.Command
Private comSubcode_Lkp As ADODB.Command
Private comSuggest_Lkp As ADODB.Command
Private comSuggestFoodCount_Lkp As ADODB.Command
Private comSuggestID_Lkp As ADODB.Command
Private comSuggestIngredCount_Lkp As ADODB.Command
Private comTagname_Lkp As ADODB.Command
Private comUpdateWordCount As ADODB.Command
Private comWord_Lkp As ADODB.Command
Private comWordCount_Lkp As ADODB.Command
Private comWordID_Lkp As ADODB.Command
Private comWordsInDoc_Lkp As ADODB.Command

Private rstAddtlDescr_Lkp As ADODB.Recordset
Private rstCountInDocument_Lkp As ADODB.Recordset
Private rstCountInDocuments_Lkp As ADODB.Recordset
Private rstDocumentCount_Lkp As ADODB.Recordset
Private rstFCDescr_Lkp As ADODB.Recordset
'Private rstFoodDescr_Lkp As ADODB.Recordset
Private rstFoodMatrixA_Lkp As ADODB.Recordset
Private rstFoodMatrixB_Lkp As ADODB.Recordset
Private rstFoodMatrixValue_Lkp As ADODB.Recordset
'Private rstIngredients_Lkp As ADODB.Recordset
'Private rstIngredRecipe_Lkp As ADODB.Recordset
''Private rstIngredSearch_Lkp As ADODB.Recordset
'Private rstModNutrient_Lkp As ADODB.Recordset
'Private rstMPED_Lkp As ADODB.Recordset
'Private rstNutrient_Lkp As ADODB.Recordset
Private rstPortionDescr_Lkp As ADODB.Recordset
Private rstPortions_Lkp As ADODB.Recordset
Private rstRecipeWeight_Lkp As ADODB.Recordset
Private rstRetDescr_Lkp As ADODB.Recordset
'Private rstSimilarRecipe_Lkp As ADODB.Recordset
Private rstSRDescr_Lkp As ADODB.Recordset
Private rstSubcode_Lkp As ADODB.Recordset
Private rstSuggest_Lkp As ADODB.Recordset
Private rstSuggestFoodCount_Lkp As ADODB.Recordset
Private rstSuggestID_Lkp As ADODB.Recordset
Private rstSuggestIngredCount_Lkp As ADODB.Recordset
Private rstTagname_Lkp As ADODB.Recordset
Private rstUpdateWordCount As ADODB.Recordset
Private rstWord_Lkp As ADODB.Recordset
Private rstWordCount_Lkp As ADODB.Recordset
Private rstWordID_Lkp As ADODB.Recordset
Private rstWordsInDoc_Lkp As ADODB.Recordset

Private appExcel As Excel.Application
Private wbkExcel1 As Excel.Workbook
Private wstExcel1 As Excel.Worksheet

Private Sub Class_Initialize()
    
    '--Reference back-end database
    Set cnnBack = New ADODB.Connection
    With cnnBack
        .ConnectionString = "Provider=SQLOLEDB;" & _
            "Data Source=SGH-03;" & _
            "Initial Catalog=shull;" & _
            "Integrated Security=SSPI"
        .CursorLocation = adUseServer
        .Open
    End With
    
    '--Reference FNDDS database
    Set cnnFNDDS = New ADODB.Connection
    With cnnFNDDS
        .ConnectionString = "Provider=SQLOLEDB;" & _
            "Data Source=SGH-03;" & _
            "Initial Catalog=FnddsDatabase;" & _
            "Integrated Security=SSPI"
        .CursorLocation = adUseServer
        .Open
    End With
    
    '--Reference MPED database
'    Set cnnMPED = New ADODB.Connection
'    With cnnMPED
'        .ConnectionString = "Provider=SQLOLEDB;" & _
'            "Data Source=SGH-03;" & _
'            "Initial Catalog=mped;" & _
'            "Integrated Security=SSPI"
'        .CursorLocation = adUseServer
'        .Open
'    End With
    
    '--Reference SR database
    Set cnnSR = New ADODB.Connection
    With cnnSR
        .ConnectionString = "Provider=SQLOLEDB;" & _
            "Data Source=SGH-03;" & _
            "Initial Catalog=sr;" & _
            "Integrated Security=SSPI"
        .CursorLocation = adUseServer
        .Open
    End With
    
    Set fso = New Scripting.FileSystemObject
    Set Utility = New clsUtility
    
    Set appExcel = New Excel.Application
    With appExcel
        Set wbkExcel1 = .Workbooks.Open(fso.BuildPath(DATABASES_PATH, "RawData\SR\MissingCodes.xlsx"))
        .Visible = True
'        Set wbkExcel2 = .Workbooks.Open(fso.BuildPath(DATABASES_PATH, "RawData\NHANES\Nutrients.xlsx"))
'        .Visible = True
    End With
    Set wstExcel1 = wbkExcel1.Worksheets("Sheet1")
'    Set wstExcel2 = wbkExcel2.Worksheets("IFF")
'    Set wstExcel3 = wbkExcel2.Worksheets("TOT")

End Sub

Private Sub AppendEquivalentDescr()

    Dim l As Long
    Dim lngDecimals As Long
    Dim lngOrder As Long
    Dim lngVersion As Long
    Dim SQL As String
    Dim strDescription As String
    Dim strTagname As String
    Dim strUnit As String
    Dim fld As ADODB.Field
    Dim rst1 As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    Dim rst3 As ADODB.Recordset
    
    '--New table
    SQL = "SELECT Tagname," & _
        "Version," & _
        "EquivalentDescription," & _
        "Unit," & _
        "Decimals," & _
        "DisplayOrder " & _
        "FROM equivalentdescr " & _
        "WHERE (Tagname IS NULL)"
    Set rst1 = New ADODB.Recordset
    Call rst1.Open(SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText)
    
    '--Old table of equivalent values
    SQL = "SELECT COLUMN_NAME AS Tagname " & _
        "FROM INFORMATION_SCHEMA.Columns " & _
        "WHERE (TABLE_NAME = 'Equivalent') " & _
        "ORDER BY ORDINAL_POSITION"
    Set rst2 = New ADODB.Recordset
    Call rst2.Open(SQL, cnnFNDDS, adOpenStatic, adLockReadOnly, adCmdText)
    
    '--Update equivalents
    For l = 3 To 7
        If l = 1 Then
            lngVersion = 1
        ElseIf l = 2 Then
            lngVersion = 2
        ElseIf l = 3 Then
            lngVersion = 4
        ElseIf l = 4 Then
            lngVersion = 8
        ElseIf l = 5 Then
            lngVersion = 16
        ElseIf l = 6 Then
            lngVersion = 32
        ElseIf l = 7 Then
            lngVersion = 64
        End If
        
        rst2.Requery
        Do Until rst2.EOF
            strTagname = rst2("Tagname")
            Select Case strTagname
                Case "FOODCODE", "MODCODE", "Version", "DESCRIPTION", "Created"
                Case Else
                    strDescription = EquivalentDescription(strTagname)
                    strUnit = EquivalentUnits(strTagname)
'                    If strTagname = "EQUIVFLAG" Then
'                        lngDecimals = 0
'                    Else
'                        lngDecimals = 3
'                    End If
                    lngOrder = EquivalentSortOrder(strTagname)
                    With rst1
                        .AddNew
                        .Fields("Version") = lngVersion
                        .Fields("EquivalentDescription") = strDescription
                        .Fields("Tagname") = strTagname
                        .Fields("Unit") = strUnit
                        .Fields("Decimals") = lngDecimals
                        .Fields("DisplayOrder") = lngOrder
                        .Update
                    End With
            End Select
            rst2.MoveNext
        Loop
        
        '-- Add whole fruit
'        strTagname = "WHOLEFRT"
'        strDescription = EquivalentDescription(strTagname)
'        strUnit = EquivalentUnits(strTagname)
'        lngDecimals = 3
'        lngOrder = EquivalentSortOrder(strTagname)
'        With rst1
'            .AddNew
'            .Fields("Version") = lngVersion
'            .Fields("EquivalentDescription") = strDescription
'            .Fields("Tagname") = strTagname
'            .Fields("Unit") = strUnit
'            .Fields("Decimals") = lngDecimals
'            .Fields("DisplayOrder") = lngOrder
'            .Update
'        End With
        
        '-- Add fruit juice
'        strTagname = "FRTJUICE"
'        strDescription = EquivalentDescription(strTagname)
'        strUnit = EquivalentUnits(strTagname)
'        lngDecimals = 3
'        lngOrder = EquivalentSortOrder(strTagname)
'        With rst1
'            .AddNew
'            .Fields("Version") = lngVersion
'            .Fields("EquivalentDescription") = strDescription
'            .Fields("Tagname") = strTagname
'            .Fields("Unit") = strUnit
'            .Fields("Decimals") = lngDecimals
'            .Fields("DisplayOrder") = lngOrder
'            .Update
'        End With

    Next l
    
    rst2.Close
    Set rst2 = Nothing
    
    rst1.Close
    Set rst1 = Nothing

End Sub

Private Sub AppendEquivalents()

    Dim dblEquivalentValue As Double
    Dim lngFoodCode As Long
    Dim lngModCode As Long
    Dim lngVersion As Long
    Dim SQL As String
    Dim strTagname As String
    Dim fld As ADODB.Field
    Dim rst1 As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    
    '--New table
    SQL = "SELECT FoodCode," & _
        "ModCode," & _
        "Tagname," & _
        "Version," & _
        "EquivalentValue " & _
        "FROM equivalents " & _
        "WHERE (FoodCode = 0)"
    Set rst1 = New ADODB.Recordset
    rst1.Open SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText
    
    '--Food code table
    '-- Exclude versions 1=FNDDS 1.0 and 2=FNDDS 2.0
    SQL = "SELECT FOODCODE," & _
        "0 AS MODCODE," & _
        "Version," & _
        "F_TOTAL," & _
        "F_CITMLB," & _
        "F_OTHER," & _
        "F_JUICE," & _
        "V_TOTAL," & _
        "V_DRKGR," & _
        "V_REDOR_TOTAL," & _
        "V_REDOR_TOMATO," & _
        "V_REDOR_OTHER," & _
        "V_STARCHY_TOTAL," & _
        "V_STARCHY_POTATO," & _
        "V_STARCHY_OTHER," & _
        "V_OTHER," & _
        "V_LEGUMES,"
    SQL = SQL & "G_TOTAL," & _
        "G_WHOLE," & _
        "G_REFINED," & _
        "PF_TOTAL," & _
        "PF_MPS_TOTAL," & _
        "PF_MEAT," & _
        "PF_CUREDMEAT," & _
        "PF_ORGAN," & _
        "PF_POULT," & _
        "PF_SEAFD_HI," & _
        "PF_SEAFD_LOW," & _
        "PF_EGGS," & _
        "PF_SOY," & _
        "PF_NUTSDS," & _
        "PF_LEGUMES,"
    SQL = SQL & "D_TOTAL," & _
        "D_MILK," & _
        "D_YOGURT," & _
        "D_CHEESE," & _
        "OILS," & _
        "SOLID_FATS," & _
        "ADD_SUGARS," & _
        "A_DRINKS " & _
        "FROM Equivalent " & _
        "UNION "
    SQL = "SELECT FOODCODE," & _
        "MODCODE," & _
        "Version," & _
        "F_TOTAL," & _
        "F_CITMLB," & _
        "F_OTHER," & _
        "F_JUICE," & _
        "V_TOTAL," & _
        "V_DRKGR," & _
        "V_REDOR_TOTAL," & _
        "V_REDOR_TOMATO," & _
        "V_REDOR_OTHER," & _
        "V_STARCHY_TOTAL," & _
        "V_STARCHY_POTATO," & _
        "V_STARCHY_OTHER," & _
        "V_OTHER," & _
        "V_LEGUMES,"
    SQL = SQL & "G_TOTAL," & _
        "G_WHOLE," & _
        "G_REFINED," & _
        "PF_TOTAL," & _
        "PF_MPS_TOTAL," & _
        "PF_MEAT," & _
        "PF_CUREDMEAT," & _
        "PF_ORGAN," & _
        "PF_POULT," & _
        "PF_SEAFD_HI," & _
        "PF_SEAFD_LOW," & _
        "PF_EGGS," & _
        "PF_SOY," & _
        "PF_NUTSDS," & _
        "PF_LEGUMES,"
    SQL = SQL & "D_TOTAL," & _
        "D_MILK," & _
        "D_YOGURT," & _
        "D_CHEESE," & _
        "OILS," & _
        "SOLID_FATS," & _
        "ADD_SUGARS," & _
        "A_DRINKS " & _
        "FROM ModEquivalent " & _
        "ORDER BY FOODCODE," & _
        "MODCODE," & _
        "Version"
    Set rst2 = New ADODB.Recordset
    Call rst2.Open(SQL, cnnFNDDS, adOpenStatic, adLockReadOnly, adCmdText)
    
    Do Until rst2.EOF
        lngFoodCode = rst2("FOODCODE")
        lngModCode = rst2("MODCODE")
        lngVersion = rst2("Version")
        For Each fld In rst2.Fields
            strTagname = Trim$(fld.name)
            dblEquivalentValue = -1
            Select Case strTagname
                Case "FOODCODE", "MODCODE", "Version"
'                    Debug.Print strTagname
'                Case "EQUIVFLAG"
'                    dblEquivalentValue = CLng(fld.Value)
                Case Else
                    dblEquivalentValue = CDbl(fld.Value)
            End Select
            If dblEquivalentValue > -1 Then
                With rst1
                    .AddNew
                    .Fields("FoodCode") = lngFoodCode
                    .Fields("ModCode") = lngModCode
                    .Fields("Tagname") = strTagname
                    .Fields("Version") = lngVersion
                    .Fields("EquivalentValue") = dblEquivalentValue
                    .Update
                End With
            End If
        Next fld
        rst2.MoveNext
    Loop
    
    rst2.Close
    Set rst2 = Nothing
    rst1.Close
    Set rst1 = Nothing

End Sub

Private Sub AppendFoodDescr()
    
    Dim lngVersion As Long
    Dim SQL As String
    Dim rst1 As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    
    '--New table
    SQL = "SELECT fooddescr.FoodCode," & _
        "fooddescr.ModCode," & _
        "fooddescr.Version," & _
        "fooddescr.MainDescription," & _
        "fooddescr.AbbrDescription," & _
        "fooddescr.IncludesDescription," & _
        "fooddescr.FortificationCode," & _
        "fooddescr.MoistureChange," & _
        "fooddescr.FatChange," & _
        "fooddescr.FatCode," & _
        "fooddescr.FatDescription," & _
        "fooddescr.WeightInitial," & _
        "fooddescr.WeightChange," & _
        "fooddescr.WeightFinal " & _
        "FROM fooddescr " & _
        "WHERE fooddescr.FoodCode = 0"
    Set rst1 = New ADODB.Recordset
    Call rst1.Open(SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText)

    '--Old table
    SQL = "SELECT MainFoodDesc.FoodCode," & _
        "MainFoodDesc.Version," & _
        "MainFoodDesc.MainFoodDescription," & _
        "MainFoodDesc.AbbreviatedMainFoodDescription," & _
        "MainFoodDesc.FortificationIdentifier," & _
        "MoistNFatAdjust.MoistureChange," & _
        "MoistNFatAdjust.FatChange," & _
        "MoistNFatAdjust.TypeOfFat " & _
        "FROM MainFoodDesc INNER JOIN MoistNFatAdjust ON " & _
        "MainFoodDesc.FoodCode = MoistNFatAdjust.FoodCode AND " & _
        "MainFoodDesc.Version = MoistNFatAdjust.Version " & _
        "ORDER BY MainFoodDesc.FoodCode," & _
        "MainFoodDesc.Version"
    Set rst2 = New ADODB.Recordset
    Call rst2.Open(SQL, cnnFNDDS, adOpenStatic, adLockReadOnly, adCmdText)

    Do Until rst2.EOF
        lngVersion = CLng(rst2("Version"))
        With rst1
            .AddNew
            .Fields("FoodCode") = rst2("FoodCode")
            .Fields("Version") = lngVersion
            .Fields("MainDescription") = rst2("MainFoodDescription")
            If Not IsNull(rst2("AbbreviatedMainFoodDescription")) Then
                .Fields("AbbrDescription") = rst2("AbbreviatedMainFoodDescription")
            End If
            '--Additional description(s)
            Call UpdateAdditionalDescriptions(CLng(rst2("FoodCode")), lngVersion, rst1)
            '--Fortification code
            .Fields("FortificationCode") = rst2("FortificationIdentifier")
            '--Moisture/fat change
            .Fields("MoistureChange") = Format(CDbl(rst2("MoistureChange")) / 100, "0.000")
            .Fields("FatChange") = Format(CDbl(rst2("FatChange")) / 100, "0.000")
            Select Case CLng(rst2("TypeOfFat"))
                Case 0
                Case Is < 10000
                    .Fields("FatCode") = String(5 - Len(CStr(rst2("TypeOfFat"))), "0") & CStr(rst2("TypeOfFat"))
                Case Is < 10000000
                    .Fields("FatCode") = CStr(rst2("TypeOfFat"))
                Case Else
                    .Fields("FatCode") = CStr(rst2("TypeOfFat"))
            End Select
            If CLng(rst2("TypeOfFat")) > 0 Then
                .Fields("FatDescription") = SRDescription(CStr(rst1("FatCode")), lngVersion, "<Missing>")
            End If
            '--Recipe weight
            .Fields("WeightInitial") = InitialWeight(rst2("FoodCode"), lngVersion)
            .Fields("WeightChange") = Format(CDbl(rst1("WeightInitial")) * (CDbl(rst2("MoistureChange") / 100) + CDbl(rst2("FatChange") / 100)), "#,##0.000")
            .Fields("WeightFinal") = Format(CDbl(rst1("WeightInitial")) * (1 + CDbl(rst2("MoistureChange") / 100) + CDbl(rst2("FatChange") / 100)), "#,##0.000")
            .Update
        End With
        rst2.MoveNext
    Loop

    rst2.Close
    Set rst2 = Nothing
    
    '--Old table (mods)
    SQL = "SELECT FoodCode," & _
        "ModificationCode," & _
        "Version," & _
        "ModificationDescription " & _
        "FROM ModDesc " & _
        "ORDER BY FoodCode," & _
        "ModificationCode," & _
        "Version"
    Set rst2 = New ADODB.Recordset
    Call rst2.Open(SQL, cnnFNDDS, adOpenStatic, adLockReadOnly, adCmdText)

    Do Until rst2.EOF
        lngVersion = CLng(rst2("Version"))
        With rst1
            .AddNew
            .Fields("FoodCode") = rst2("FoodCode")
            .Fields("ModCode") = rst2("ModificationCode")
            .Fields("Version") = lngVersion
            .Fields("MainDescription") = rst2("ModificationDescription")
            .Update
        End With
        rst2.MoveNext
    Loop

    rst2.Close
    Set rst2 = Nothing
    rst1.Close
    Set rst1 = Nothing

End Sub

Private Sub AppendFoodSearch()
    
    Dim SQL As String
    Dim rst1 As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
        
    '--New table
    SQL = "SELECT foodsearch.FoodCode," & _
        "foodsearch.ModCode," & _
        "foodsearch.SeqNum," & _
        "foodsearch.Version," & _
        "foodsearch.FoodDescription " & _
        "FROM foodsearch " & _
        "WHERE foodsearch.FoodCode = 0"
    Set rst1 = New ADODB.Recordset
    Call rst1.Open(SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText)

    '--Old table
    SQL = "SELECT MainFoodDesc.FoodCode," & _
        "MainFoodDesc.Version," & _
        "MainFoodDesc.MainFoodDescription " & _
        "FROM MainFoodDesc " & _
        "ORDER BY MainFoodDesc.FoodCode"
    Set rst2 = New ADODB.Recordset
    Call rst2.Open(SQL, cnnFNDDS, adOpenStatic, adLockReadOnly, adCmdText)

    Do Until rst2.EOF
        With rst1
            .AddNew
            .Fields("FoodCode") = rst2("FoodCode")
            .Fields("SeqNum") = 0
            .Fields("Version") = rst2("Version")
            .Fields("FoodDescription") = rst2("MainFoodDescription")
            .Update
        End With
        rst2.MoveNext
    Loop
    
    rst2.Close
    Set rst2 = Nothing
    
    '--Old table (mods)
    SQL = "SELECT FoodCode," & _
        "ModificationCode," & _
        "Version," & _
        "ModificationDescription " & _
        "FROM ModDesc " & _
        "ORDER BY FoodCode," & _
        "ModificationCode"
    Set rst2 = New ADODB.Recordset
    Call rst2.Open(SQL, cnnFNDDS, adOpenStatic, adLockReadOnly, adCmdText)

    Do Until rst2.EOF
        With rst1
            .AddNew
            .Fields("FoodCode") = rst2("FoodCode")
            .Fields("ModCode") = rst2("ModificationCode")
            .Fields("SeqNum") = 0
            .Fields("Version") = rst2("Version")
            .Fields("FoodDescription") = rst2("ModificationDescription")
            .Update
        End With
        rst2.MoveNext
    Loop
    
    rst2.Close
    Set rst2 = Nothing
    
    '--Old table (adds)
    SQL = "SELECT AddFoodDesc.FoodCode," & _
        "AddFoodDesc.SeqNum," & _
        "AddFoodDesc.Version," & _
        "AddFoodDesc.AdditionalFoodDescription " & _
        "FROM AddFoodDesc " & _
        "ORDER BY AddFoodDesc.FoodCode," & _
        "AddFoodDesc.SeqNum"
    Set rst2 = New ADODB.Recordset
    Call rst2.Open(SQL, cnnFNDDS, adOpenStatic, adLockReadOnly, adCmdText)

    Do Until rst2.EOF
        With rst1
            .AddNew
            .Fields("FoodCode") = rst2("FoodCode")
            .Fields("SeqNum") = rst2("SeqNum")
            .Fields("Version") = rst2("Version")
            .Fields("FoodDescription") = rst2("AdditionalFoodDescription")
            .Update
        End With
        rst2.MoveNext
    Loop

    rst2.Close
    Set rst2 = Nothing
    rst1.Close
    Set rst1 = Nothing

End Sub

Private Sub AppendFoodSuggest()

    Dim blnDescribed As Boolean
    Dim l As Long
    Dim lngFoodCode As Long
    Dim lngModCode As Long
    Dim lngTermID As Long
    Dim lngType As Long
    Dim lngVersion As Long
    Dim SQL As String
    Dim strDescription As String
    Dim strTerm As String
    Dim strTerms() As String
    Dim rst1 As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    
    '--New table
    SQL = "SELECT SuggestID," & _
        "SuggestType," & _
        "SuggestDescription " & _
        "FROM suggest " & _
        "WHERE SuggestID = 0"
    Set rst1 = New ADODB.Recordset
    Call rst1.Open(SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText)
    
    '--New food description table
    SQL = "SELECT FoodCode," & _
        "ModCode," & _
        "Version," & _
        "MainDescription," & _
        "IncludesDescription " & _
        "FROM fooddescr " & _
        "ORDER BY FoodCode," & _
        "Version"
    Set rst2 = New ADODB.Recordset
    Call rst2.Open(SQL, cnnBack, adOpenStatic, adLockReadOnly, adCmdText)

    Do Until rst2.EOF
        lngFoodCode = CLng(rst2("FoodCode"))
        lngModCode = CLng(rst2("ModCode"))
        lngVersion = CLng(rst2("Version"))
        
        '--Main food descriptions
        lngType = 1
        blnDescribed = False
        If lngModCode = 0 Then
            Select Case lngFoodCode
                Case 11112000
                    'Milk, cow's, fluid, other than whole, NS as to 2%, 1%, or skim (formerly milk, cow's, fluid, "lowfat", NS as to percent fat)
                    ReDim strTerms(6)
                    strTerms(0) = "Milk"
                    strTerms(1) = "cow's"
                    strTerms(2) = "fluid"
                    strTerms(3) = "other than whole"
                    strTerms(4) = "NS as to 2%, 1%, or skim"
                    strTerms(5) = "formerly milk, cow's, fluid, ""lowfat"""
                    strTerms(6) = "NS as to percent fat"
                Case 11511200
                    'Milk, chocolate, reduced fat milk-based, 2% (formerly "lowfat")
                    ReDim strTerms(3)
                    strTerms(0) = "Milk"
                    strTerms(1) = "chocolate"
                    strTerms(2) = "reduced fat milk-based, 2%"
                    strTerms(3) = "formerly ""lowfat"""
                Case 27114000
                    'Beef with (mushroom) soup (mixture)
                    ReDim strTerms(2)
                    strTerms(0) = "Beef"
                    strTerms(1) = "with (mushroom) soup"
                    strTerms(2) = "mixture"
                Case 27144000
                    'Chicken or turkey with (mushroom) soup (mixture)
                    ReDim strTerms(2)
                    strTerms(0) = "Chicken or turkey"
                    strTerms(1) = "with (mushroom) soup"
                    strTerms(2) = "mixture"
                Case 27213400
                    'Beef and rice with (mushroom) soup (mixture)
                    ReDim strTerms(2)
                    strTerms(0) = "Beef and rice"
                    strTerms(1) = "with (mushroom) soup"
                    strTerms(2) = "mixture"
                Case 27243400
                    'Chicken or turkey and rice with (mushroom) soup (mixture)
                    ReDim strTerms(2)
                    strTerms(0) = "Chicken or turkey and rice"
                    strTerms(1) = "with (mushroom) soup"
                    strTerms(2) = "mixture"
                Case 27250830
                    'Fish and rice with (mushroom) soup
                    ReDim strTerms(0)
                    strTerms(0) = "Fish and rice with (mushroom) soup"
                Case 27250900
                    'Fish and noodles with (mushroom) soup
                    ReDim strTerms(0)
                    strTerms(0) = "Fish and noodles with (mushroom) soup"
                Case 28145410
                    'Turkey with gravy, dressing, potatoes, vegetable, cream of tomato soup, dessert (frozen meal)
                    ReDim strTerms(7)
                    strTerms(0) = "Turkey with gravy"
                    strTerms(1) = "dressing"
                    strTerms(2) = "potatoes"
                    strTerms(3) = "vegetable"
                    strTerms(4) = "cream of"
                    strTerms(5) = "tomato soup"
                    strTerms(6) = "dessert"
                    strTerms(7) = "frozen meal"
                Case 53241500
                    'Cookie, butter or sugar cookie
                    ReDim strTerms(2)
                    strTerms(0) = "Cookie"
                    strTerms(1) = "butter or sugar"
                    strTerms(2) = "cookie"
                Case 53241600
                    'Cookie, butter or sugar cookie, with fruit and/or nuts
                    ReDim strTerms(3)
                    strTerms(0) = "Cookie"
                    strTerms(1) = "butter or sugar"
                    strTerms(2) = "cookie"
                    strTerms(3) = "with fruit and/or nuts"
                Case 54101010
                    'Cracker, animal
                    ReDim strTerms(0)
                    strTerms(0) = "Cracker, animal"
                Case 56205410
                    'Rice, white, cooked with (fat) oil, Puerto Rican style (Arroz blanco)
                    ReDim strTerms(4)
                    strTerms(0) = "Rice"
                    strTerms(1) = "white"
                    strTerms(2) = "cooked with (fat) oil"
                    strTerms(3) = "Puerto Rican style"
                    strTerms(4) = "Arroz blanco"
                Case 58126180
                    'Turnover, meat-, potato-, and vegetable-filled, no gravy
                    ReDim strTerms(2)
                    strTerms(0) = "Turnover"
                    strTerms(1) = "meat-, potato-, and vegetable-filled"
                    strTerms(2) = "no gravy"
                Case 58132310
                    'Spaghetti with tomato sauce and meatballs or spaghetti with meat sauce or spaghetti with meat sauce and meatballs
                    ReDim strTerms(2)
                    strTerms(0) = "Spaghetti with tomato sauce and meatballs"
                    strTerms(1) = "spaghetti with meat sauce"
                    strTerms(2) = "spaghetti with meat sauce and meatballs"
                Case 63320100
                    If lngVersion < 4 Then
                        'Fruit salad, Puerto Rican style (Mixture includes bananas, papayas, oranges, grapefruit, etc.) (Ensalada de frutas tropicales)
                        ReDim strTerms(3)
                        strTerms(0) = "Fruit salad"
                        strTerms(1) = "Puerto Rican style"
                        strTerms(2) = "Mixture includes bananas, papayas, oranges, grapefruit, etc."
                        strTerms(3) = "Ensalada de frutas tropicales"
                    Else
                        'Fruit salad, Puerto Rican style (Mixture includes bananas, papayas, oranges, etc.) (Ensalada de frutas tropicales)
                        ReDim strTerms(3)
                        strTerms(0) = "Fruit salad"
                        strTerms(1) = "Puerto Rican style"
                        strTerms(2) = "Mixture includes bananas, papayas, oranges, etc."
                        strTerms(3) = "Ensalada de frutas tropicales"
                    End If
                Case 75340000
                    'Vegetable combinations, Oriental style, (broccoli, green pepper, water chestnut, etc) cooked, NS as to fat added in cooking
                    ReDim strTerms(4)
                    strTerms(0) = "Vegetable combinations"
                    strTerms(1) = "Oriental style"
                    strTerms(2) = "broccoli, green pepper, water chestnut, etc"
                    strTerms(3) = "cooked"
                    strTerms(4) = "NS as to fat added in cooking"
                Case 75340010
                    'Vegetable combinations, Oriental style, (broccoli, green pepper,  water chestnuts, etc), cooked, fat not added in cooking
                    ReDim strTerms(4)
                    strTerms(0) = "Vegetable combinations"
                    strTerms(1) = "Oriental style"
                    strTerms(2) = "broccoli, green pepper,  water chestnuts, etc"
                    strTerms(3) = "cooked"
                    strTerms(4) = "fat not added in cooking"
                Case 75340020
                    'Vegetable combinations, Oriental style, (broccoli, green pepper, water chestnuts, etc), cooked, fat added in cooking
                    ReDim strTerms(4)
                    strTerms(0) = "Vegetable combinations"
                    strTerms(1) = "Oriental style"
                    strTerms(2) = "broccoli, green pepper, water chestnuts, etc"
                    strTerms(3) = "cooked"
                    strTerms(4) = "fat added in cooking"
                Case 75340100
                    'Vegetable combinations (broccoli, carrots, corn, cauliflower, etc.), cooked, NS as to fat added in cooking
                    ReDim strTerms(3)
                    strTerms(0) = "Vegetable combinations"
                    strTerms(1) = "broccoli, carrots, corn, cauliflower, etc."
                    strTerms(2) = "cooked"
                    strTerms(3) = "NS as to fat added in cooking"
                Case 75340110
                    'Vegetable combinations (broccoli, carrots, corn, cauliflower, etc.), cooked, fat not added in cooking
                    ReDim strTerms(3)
                    strTerms(0) = "Vegetable combinations"
                    strTerms(1) = "broccoli, carrots, corn, cauliflower, etc."
                    strTerms(2) = "cooked"
                    strTerms(3) = "fat not added in cooking"
                Case 75340120
                    'Vegetable combinations (broccoli, carrots, corn, cauliflower, etc.), cooked, fat added in cooking
                    ReDim strTerms(3)
                    strTerms(0) = "Vegetable combinations"
                    strTerms(1) = "broccoli, carrots, corn, cauliflower, etc."
                    strTerms(2) = "cooked"
                    strTerms(3) = "fat added in cooking"
                Case 75340160
                    'Vegetable and pasta combinations with cream or cheese sauce (broccoli, pasta, carrots, corn, zucchini, peppers, cauliflower, peas, etc.), cooked
                    ReDim strTerms(2)
                    strTerms(0) = "Vegetable and pasta combinations with cream or cheese sauce"
                    strTerms(1) = "broccoli, pasta, carrots, corn, zucchini, peppers, cauliflower, peas, etc."
                    strTerms(2) = "cooked"
                Case 75340300
                    'Pinacbet (eggplant with tomatoes, bitter melon, etc.)
                    ReDim strTerms(1)
                    strTerms(0) = "Pinacbet"
                    strTerms(1) = "eggplant with tomatoes, bitter melon, etc."
                Case 81203200
                    'Shortening, animal
                    ReDim strTerms(0)
                    strTerms(0) = "Shortening, animal"
                Case 81302030
                    'Orange sauce (for duck)
                    ReDim strTerms(0)
                    strTerms(0) = "Orange sauce (for duck)"
                Case 82105800
                    'Canola, soybean and sunflower oil
                    ReDim strTerms(0)
                    strTerms(0) = "Canola, soybean and sunflower oil"
                Case 83100100
                    'Salad dressing, NFS, for salads
                    ReDim strTerms(0)
                    strTerms(0) = "Salad dressing, NFS, for salads"
                Case 83100200
                    'Salad dressing, NFS, for sandwiches
                    ReDim strTerms(0)
                    strTerms(0) = "Salad dressing, NFS, for sandwiches"
                Case 91511090
                    'Gelatin dessert, dietetic, with fruit and vegetable(s), sweetened with low calorie sweetener
                    ReDim strTerms(3)
                    strTerms(0) = "Gelatin dessert"
                    strTerms(1) = "dietetic"
                    strTerms(2) = "with fruit and vegetable(s)"
                    strTerms(3) = "sweetened with low calorie sweetener"
                Case 91520100
                    'Yookan (Yokan), a Japanese dessert made with bean paste and sugar
                    ReDim strTerms(3)
                    strTerms(0) = "Yookan"
                    strTerms(1) = "Yokan"
                    strTerms(2) = "Japanese dessert"
                    strTerms(3) = "made with bean paste and sugar"
                Case Else
                    strDescription = FormattedSuggestDescr(rst2("MainDescription"))
                    strTerms = Split(strDescription, ",", , vbTextCompare)
            End Select
        Else
            Select Case lngFoodCode
                Case 27243400
                    If lngModCode = 205515 Then
                        'Chicken or turkey and rice with (mushroom) soup (mixture) W/ VEGETABLE OIL, NFS (INCLUDE OIL, NFS)
                        ReDim strTerms(3)
                        strTerms(0) = "Chicken or turkey and rice"
                        strTerms(1) = "with (mushroom) soup"
                        strTerms(2) = "mixture"
                        strTerms(3) = "W/ VEGETABLE OIL, NFS (INCLUDE OIL, NFS)"
                    ElseIf lngModCode = 207140 Then
                        'Chicken or turkey and rice with (mushroom) soup (mixture) W/O FAT
                        ReDim strTerms(3)
                        strTerms(0) = "Chicken or turkey and rice"
                        strTerms(1) = "with (mushroom) soup"
                        strTerms(2) = "mixture"
                        strTerms(3) = "W/O FAT"
                    End If
                Case 33201500
                    If lngModCode = 205573 Then
                        'Scrambled egg, made from cholesterol-free frozen mixture with vegetables W/O FAT OR W/ NONSTICK SPRAY (INCLUDE PAM...)
                        ReDim strTerms(5)
                        strTerms(0) = "Scrambled egg"
                        strTerms(1) = "made from cholesterol-free frozen mixture"
                        strTerms(2) = "with vegetables"
                        strTerms(3) = "W/O FAT"
                        strTerms(4) = "W/ NONSTICK SPRAY"
                        strTerms(5) = "INCLUDE PAM"
                        blnDescribed = True
                    End If
                Case 58148550
                    If lngModCode = 206182 Then
                        'Pasta or macaroni salad with meat and oil and vinegar-type dressing W/ ITALIAN DRESSING, LOW CALORIE
                        ReDim strTerms(2)
                        strTerms(0) = "Pasta or macaroni salad"
                        strTerms(1) = "with meat and oil and vinegar-type dressing"
                        strTerms(2) = "W/ ITALIAN DRESSING, LOW CALORIE"
                        blnDescribed = True
                    End If
                Case 75340020
                    If lngModCode = 101229 Then
                        'Vegetable combinations, Oriental style, (broccoli, green pepper, chinese cabbage, water chestnuts, etc), cooked, fat added in cooking W/ BUTTER, NFS
                        ReDim strTerms(4)
                        strTerms(0) = "Vegetable combinations"
                        strTerms(1) = "Oriental style"
                        strTerms(2) = "broccoli, green pepper, chinese cabbage, water chestnuts, etc"
                        strTerms(3) = "cooked"
                        strTerms(4) = "fat added in cooking W/ BUTTER, NFS"
                    ElseIf lngModCode = 200421 Then
                        'Vegetable combinations, Oriental style, (broccoli, green pepper, chinese cabbage, water chestnuts, etc), cooked, fat added in cooking W/ VEGETABLE OIL, NFS (INCLUDE OIL, NFS)
                        ReDim strTerms(5)
                        strTerms(0) = "Vegetable combinations"
                        strTerms(1) = "Oriental style"
                        strTerms(2) = "broccoli, green pepper, chinese cabbage, water chestnuts, etc"
                        strTerms(3) = "cooked"
                        strTerms(4) = "fat added in cooking W/ VEGETABLE OIL, NFS"
                        strTerms(5) = "INCLUDE OIL, NFS"
                    End If
                Case 75340120
                    If lngModCode = 200695 Then
                        'Vegetable combinations (broccoli, carrots, corn, cauliflower, etc.), cooked, fat added in cooking W/ BUTTER, NFS
                        ReDim strTerms(3)
                        strTerms(0) = "Vegetable combinations"
                        strTerms(1) = "broccoli, carrots, corn, cauliflower, etc."
                        strTerms(2) = "cooked"
                        strTerms(3) = "fat added in cooking W/ BUTTER, NFS"
                    ElseIf lngModCode = 201702 Then
                        'Vegetable combinations (broccoli, carrots, corn, cauliflower, etc.), cooked, fat added in cooking W/ VEGETABLE OIL, NFS (INCLUDE OIL, NFS)
                        ReDim strTerms(4)
                        strTerms(0) = "Vegetable combinations"
                        strTerms(1) = "broccoli, carrots, corn, cauliflower, etc."
                        strTerms(2) = "cooked"
                        strTerms(3) = "fat added in cooking W/ VEGETABLE OIL, NFS"
                        strTerms(4) = "INCLUDE OIL, NFS"
                    ElseIf lngModCode = 205498 Then
                        'Vegetable combinations (broccoli, carrots, corn, cauliflower, etc.), cooked, fat added in cooking W/ ANIMAL FAT OR MEAT DRIPPINGS
                        ReDim strTerms(3)
                        strTerms(0) = "Vegetable combinations"
                        strTerms(1) = "broccoli, carrots, corn, cauliflower, etc."
                        strTerms(2) = "cooked"
                        strTerms(3) = "fat added in cooking W/ ANIMAL FAT OR MEAT DRIPPINGS"
                    End If
                Case 75340160
                    If lngModCode = 206975 Then
                        'Vegetable and pasta combinations with cream or cheese sauce (broccoli, pasta, carrots, corn, zucchini, peppers, cauliflower, peas, etc.), cooked W/ BUTTER, NFS
                        ReDim strTerms(2)
                        strTerms(0) = "Vegetable and pasta combinations with cream or cheese sauce"
                        strTerms(1) = "broccoli, pasta, carrots, corn, zucchini, peppers, cauliflower, peas, etc."
                        strTerms(2) = "cooked W/ BUTTER, NFS"
                    ElseIf lngModCode = 207000 Then
                        'Vegetable and pasta combinations with cream or cheese sauce (broccoli, pasta, carrots, corn, zucchini, peppers, cauliflower, peas, etc.), cooked W/O FAT
                        ReDim strTerms(2)
                        strTerms(0) = "Vegetable and pasta combinations with cream or cheese sauce"
                        strTerms(1) = "broccoli, pasta, carrots, corn, zucchini, peppers, cauliflower, peas, etc."
                        strTerms(2) = "cooked W/O FAT"
                    ElseIf lngModCode = 207090 Then
                        'Vegetable and pasta combinations with cream or cheese sauce (broccoli, pasta, carrots, corn, zucchini, peppers, cauliflower, peas, etc.), cooked W/ VEGETABLE OIL, NFS (INCLUDE OIL, NFS)
                        ReDim strTerms(3)
                        strTerms(0) = "Vegetable and pasta combinations with cream or cheese sauce"
                        strTerms(1) = "broccoli, pasta, carrots, corn, zucchini, peppers, cauliflower, peas, etc."
                        strTerms(2) = "cooked W/ VEGETABLE OIL, NFS"
                        strTerms(3) = "INCLUDE OIL, NFS"
                    End If
                Case 75340300
                    If lngModCode = 101281 Then
                        'Pinacbet (eggplant with tomatoes, bitter melon, etc.) W/ ANIMAL FAT OR MEAT DRIPPINGS
                        ReDim strTerms(2)
                        strTerms(0) = "Pinacbet"
                        strTerms(1) = "eggplant with tomatoes, bitter melon, etc."
                        strTerms(2) = "W/ ANIMAL FAT OR MEAT DRIPPINGS"
                    End If
                Case 75649010
                    If lngModCode = 100724 Then
                        'Vegetable soup, prepared with water or ready-to-serve MADE FROM CONDENSED W/ 2 CANS OF WATER ADDED OR READY-TO-SERVE WITH 1/2 CAN WATER ADDED
                        ReDim strTerms(4)
                        strTerms(0) = "Vegetable soup"
                        strTerms(1) = "prepared with water or ready-to-serve"
                        strTerms(2) = "MADE FROM CONDENSED"
                        strTerms(3) = "W/ 2 CANS OF WATER ADDED"
                        strTerms(4) = "READY-TO-SERVE WITH 1/2 CAN WATER ADDED"
                        blnDescribed = True
                    Else
                        'Vegetable soup, prepared with water or ready-to-serve MADE FROM CONDENSED W/ 2 CANS OF WATER OR READY-TO-SERVE WITH 1/2 CAN WATER ADDED
                        ReDim strTerms(4)
                        strTerms(0) = "Vegetable soup"
                        strTerms(1) = "prepared with water or ready-to-serve"
                        strTerms(2) = "MADE FROM CONDENSED"
                        strTerms(3) = "W/ 2 CANS OF WATER"
                        strTerms(4) = "READY-TO-SERVE WITH 1/2 CAN WATER ADDED"
                        blnDescribed = True
                    End If
                Case Else
                    'Stop
            End Select
            If Not blnDescribed Then
                strDescription = FormattedSuggestDescr(rst2("MainDescription"))
                strTerms = Split(strDescription, ",", , vbTextCompare)
            End If
        End If
        strTerms = FormattedSuggestTerms(strTerms)
        For l = 0 To UBound(strTerms())
            strTerm = Trim$(LCase(strTerms(l)))
            If Len(strTerm) > 0 Then
'                    Debug.Print strTerm
                lngTermID = SuggestTermExists(lngType, strTerm)
                If lngTermID = 0 Then
                    '--Add term
                    lngTermID = SuggestTermID(lngType) + 1
                    With rst1
                        .AddNew
                        .Fields("SuggestID") = lngTermID
                        .Fields("SuggestType") = lngType
                        .Fields("SuggestDescription") = strTerm
                        .Update
                    End With
                End If
                '--Update count
                Call UpdateFoodSuggestCount(lngFoodCode, lngModCode, lngVersion, lngTermID, lngType)
            End If
        Next l
        
        '--Additional food descriptions
        lngType = 2
        strDescription = vbNullString
        If Not IsNull(rst2("IncludesDescription")) Then
            strDescription = rst2("IncludesDescription")
        End If
        If Len(strDescription) > 0 Then
            Select Case lngFoodCode
                Case 23150270
                    ReDim strTerms(2)
                    strTerms(0) = "barbecued"
                    strTerms(1) = "no sauce added"
                    strTerms(2) = """barbecoa de cabeza, carne de cabra, sin salsa"""
                Case 25221880
                    ReDim strTerms(8)
                    strTerms(0) = "Hillshire Farm"
                    strTerms(1) = "80% Fat Free"
                    strTerms(2) = "Turkey"
                    strTerms(3) = "Pork"
                    strTerms(4) = "Beef Lite Smoked Sausage"
                    strTerms(5) = "Bryan Light"
                    strTerms(6) = "85% Fat Free"
                    strTerms(7) = "Smoked Sausage"
                    strTerms(8) = "Eckrich Reduced Fat Smoked Sausage"
                Case 26141110
                    ReDim strTerms(4)
                    strTerms(0) = "grouper"
                    strTerms(1) = "striped bass"
                    strTerms(2) = "wreakfish"
                    strTerms(3) = "bass"
                    strTerms(4) = "NFS"
                Case 26141120
                    If CLng(rst2("Version")) = 1 Then
                        ReDim strTerms(6)
                        strTerms(0) = "sauteed"
                        strTerms(1) = "fried with no coating"
                        strTerms(2) = "grouper"
                        strTerms(3) = "striped bass"
                        strTerms(4) = "wreakfish"
                        strTerms(5) = "bass"
                        strTerms(6) = "NFS"
                    End If
                Case 26141130
                    ReDim strTerms(4)
                    strTerms(0) = "grouper"
                    strTerms(1) = "striped bass"
                    strTerms(2) = "wreakfish"
                    strTerms(3) = "bass"
                    strTerms(4) = "NFS"
                Case 26141140
                    ReDim strTerms(6)
                    strTerms(0) = "fried"
                    strTerms(1) = "NS as to coating"
                    strTerms(2) = "grouper"
                    strTerms(3) = "striped bass"
                    strTerms(4) = "wreakfish"
                    strTerms(5) = "bass"
                    strTerms(6) = "NFS"
                Case 26141160
                    ReDim strTerms(4)
                    strTerms(0) = "grouper"
                    strTerms(1) = "striped bass"
                    strTerms(2) = "wreakfish"
                    strTerms(3) = "bass"
                    strTerms(4) = "NFS"
                Case 27350110
                    ReDim strTerms(3)
                    strTerms(0) = "seafood stew made with tomato, fish, & shellfish"
                    strTerms(1) = "clams"
                    strTerms(2) = "scallops"
                    strTerms(3) = "shrimp"
                Case 28113010
                    'Lean Cuisine Oriental Beef; or Benihana Oriental Lites Beef and Mushrooms in sauce with Vegetables and Rice
                    ReDim strTerms(2)
                    strTerms(0) = "Lean Cuisine Oriental Beef"
                    strTerms(1) = "Benihana Oriental Lites Beef and Mushrooms in sauce"
                    strTerms(2) = "with Vegetables and Rice"
                Case 28143020
                    'Benihana Oriental Lites Chicken in Spicy Garlic Sauce with Vegetables and Rice
                    ReDim strTerms(1)
                    strTerms(0) = "Benihana Oriental Lites Chicken in Spicy Garlic Sauce"
                    strTerms(1) = "with Vegetables and Rice"
                Case 28340150
                    'chicken broth stock with vegetables, without chicken or rice, for Caldo de Pollo
                    ReDim strTerms(2)
                    strTerms(0) = "chicken broth stock with vegetables"
                    strTerms(1) = "without chicken or rice"
                    strTerms(2) = "Caldo de Pollo"
                Case 32202020
                    ReDim strTerms(0)
                    strTerms(0) = "Hardee's Ham, Egg, & Cheese Biscuit"
                Case 32202070
                    ReDim strTerms(2)
                    strTerms(0) = "Swanson Great Starts Egg, Cheese & Bacon on a Biscuit breakfast sandwich"
                    strTerms(1) = "McDonald's Bacon, Egg, & Cheese Biscuit"
                    strTerms(2) = "Hardee's Bacon, Egg, & Cheese Biscuit"
                Case 32202075
                    ReDim strTerms(0)
                    strTerms(0) = "McDonald's Bacon, Egg, & Cheese McGriddles"
                Case 53452150
                    'nine-layer pudding, a Chinese steamed rice and syrup pudding
                    ReDim strTerms(1)
                    strTerms(0) = "nine-layer pudding"
                    strTerms(1) = "Chinese steamed rice and syrup pudding"
                Case 64100200
                    'Ocean Spray 100% Juice Blends, all flavors; or "cranberry juice, 100% juice"
                    ReDim strTerms(2)
                    strTerms(0) = "Ocean Spray 100% Juice Blends"
                    strTerms(1) = "all flavors"
                    strTerms(2) = """cranberry juice, 100% juice"""
                Case 72302000
                    'cream of broccoli soup
                    ReDim strTerms(1)
                    strTerms(0) = "cream of"
                    strTerms(1) = "broccoli soup"
                Case 75607030
                    'cream of mushroom soup, undiluted
                    ReDim strTerms(2)
                    strTerms(0) = "cream of"
                    strTerms(1) = "mushroom soup"
                    strTerms(2) = "undiluted"
                Case 81101000
                    'stick butter, NS as to salt; butter, seasoned, e.g., garlic butter; salted butter, NS as to stick or tub; or Land O Lakes Salted Stick Butter
                    ReDim strTerms(7)
                    strTerms(0) = "stick butter"
                    strTerms(1) = "NS as to salt"
                    strTerms(2) = "butter"
                    strTerms(3) = "seasoned"
                    strTerms(4) = "e.g., garlic butter"
                    strTerms(5) = "salted butter"
                    strTerms(6) = "NS as to stick or tub"
                    strTerms(7) = "Land O Lakes Salted Stick Butter"
                Case 91305020
                    'creme filling; icing, NFS; or icing with added flavors, e.g., lemon icing, etc.
                    ReDim strTerms(4)
                    strTerms(0) = "creme filling"
                    strTerms(1) = "icing"
                    strTerms(2) = "NFS"
                    strTerms(3) = "icing with added flavors"
                    strTerms(4) = "e.g., lemon icing, etc."
                Case 91703020
                    'Brach's Royal; Caramel Creams; Sugar Babies; Sugar Daddy; Kraft caramels; Jersey's; or Pearson Caramel Nips
                    ReDim strTerms(7)
                    strTerms(0) = "Brach's Royal"
                    strTerms(1) = "Caramel Creams"
                    strTerms(2) = "Sugar Babies"
                    strTerms(3) = "Sugar Daddy"
                    strTerms(4) = "Kraft"
                    strTerms(5) = "caramels"
                    strTerms(6) = "Jersey's"
                    strTerms(7) = "Pearson Caramel Nips"
                Case 91703060
                    ReDim strTerms(4)
                    strTerms(0) = "Goo Goo Cluster(s)"
                    strTerms(1) = "Peanut Chews"
                    strTerms(2) = "Toffifay"
                    strTerms(3) = "Turtles"
                    strTerms(4) = "Reese's NutRageous"
                Case 91715300
                    ReDim strTerms(0)
                    strTerms(0) = "$ 100,000 Bar"
                Case 92410110
                    'tonic water; quinine water; fruit flavors; Clearly Canadian Original, all flavors; Mistic Sparkling Water Beverage with Fruit Flavor and Natural Spring Water; or Penafiel, all flavors
                    ReDim strTerms(8)
                    strTerms(0) = "tonic water"
                    strTerms(1) = "quinine water"
                    strTerms(2) = "fruit flavors"
                    strTerms(3) = "Clearly Canadian Original"
                    strTerms(4) = "all flavors"
                    strTerms(5) = "Mistic Sparkling Water Beverage"
                    strTerms(6) = "with Fruit Flavor and Natural Spring Water"
                    strTerms(7) = "Penafiel"
                    strTerms(8) = "all flavors"
               Case 92511250
                    ReDim strTerms(0)
                    strTerms(0) = "Five (5) Alive Citrus Beverage"
                Case Else
                    strDescription = FormattedSuggestDescr(strDescription, True)
                    strTerms = Split(strDescription, ",", , vbTextCompare)
            End Select
            strTerms = FormattedSuggestTerms(strTerms)
            For l = 0 To UBound(strTerms())
                strTerm = Trim$(LCase(strTerms(l)))
                If Len(strTerm) > 0 Then
'                    Debug.Print strTerm
                    lngTermID = SuggestTermExists(lngType, strTerm)
                    If lngTermID = 0 Then
                        '--Add term
                        lngTermID = SuggestTermID(lngType) + 1
                        With rst1
                            .AddNew
                            .Fields("SuggestID") = lngTermID
                            .Fields("SuggestType") = lngType
                            .Fields("SuggestDescription") = strTerm
                            .Update
                        End With
                    End If
                    '--Update count
                    Call UpdateFoodSuggestCount(lngFoodCode, 0, lngVersion, lngTermID, lngType)
                End If
            Next l
        End If
        rst2.MoveNext
    Loop

    rst1.Close
    Set rst1 = Nothing
    rst2.Close
    Set rst2 = Nothing

End Sub

Private Sub AppendFoodWords()

    Dim l As Long
    Dim lngFoodCode As Long
    Dim lngModCode As Long
    Dim lngWordID As Long
    Dim lngType As Long
    Dim lngVersion As Long
    Dim SQL As String
    Dim strDescription As String
    Dim strWord As String
    Dim strWords() As String
    Dim rst1 As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    
    '--New table
    SQL = "SELECT WordID," & _
        "WordDescription " & _
        "FROM word " & _
        "WHERE WordID = 0"
    Set rst1 = New ADODB.Recordset
    Call rst1.Open(SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText)
    
    '--New food description table
    SQL = "SELECT FoodCode," & _
        "ModCode," & _
        "Version," & _
        "MainDescription," & _
        "IncludesDescription " & _
        "FROM fooddescr " & _
        "ORDER BY FoodCode," & _
        "Version"
    Set rst2 = New ADODB.Recordset
    Call rst2.Open(SQL, cnnBack, adOpenStatic, adLockReadOnly, adCmdText)

    Do Until rst2.EOF
        lngFoodCode = CLng(rst2("FoodCode"))
        lngModCode = CLng(rst2("ModCode"))
        lngVersion = CLng(rst2("Version"))
        Debug.Print lngFoodCode, lngModCode, lngVersion
        
        '--Main food descriptions
        lngType = 1
        Select Case lngFoodCode
            Case 57320500, 57321500
                strDescription = Replace(rst2("MainDescription"), "100 %", "100%", , , vbTextCompare)
                strDescription = Replace(strDescription, "  ", " ", , , vbTextCompare)
            Case Else
                strDescription = Replace(rst2("MainDescription"), "  ", " ", , , vbTextCompare)
        End Select
        strWords = Split(strDescription, " ", , vbTextCompare)
        strWords = FormattedWords(strWords)
        For l = 0 To UBound(strWords())
            strWord = Trim$(LCase(strWords(l)))
            If Len(strWord) > 0 Then
                Debug.Print strWord
                lngWordID = WordExists(strWord)
                If lngWordID = 0 Then
                    '--Add term
                    lngWordID = WordID() + 1
                    With rst1
                        .AddNew
                        .Fields("WordID") = lngWordID
                        .Fields("WordDescription") = strWord
                        .Update
                    End With
                End If
                '--Update count
                Call UpdateWordCount(lngFoodCode, lngModCode, lngVersion, lngWordID, lngType)
            End If
        Next l
        
        '--Additional food descriptions
        lngType = 2
        strDescription = rst2("IncludesDescription")
        If Len(strDescription) > 0 Then
            Select Case lngFoodCode
                Case 64104200
                    strDescription = "Tree Top Apple-Pear 100% Juice; or Apple and Eve Raspberry Cranberry 100% Juice Blend"
                Case 91715300
                    strDescription = "$100,000 Bar"
                Case 91746100
                    strDescription = """M&M's"" Mint Chocolate Candies; or ""M&M's"" Mini Baking Bits"
                Case Else
                    strDescription = Replace(strDescription, "  ", " ", , , vbTextCompare)
            End Select
            Debug.Print strDescription
            strWords = Split(strDescription, " ", , vbTextCompare)
            strWords = FormattedWords(strWords)
            For l = 0 To UBound(strWords())
                strWord = Trim$(LCase(strWords(l)))
                If Len(strWord) > 0 Then
                    Debug.Print strWord
                    lngWordID = WordExists(strWord)
                    If lngWordID = 0 Then
                        '--Add term
                        lngWordID = WordID() + 1
                        With rst1
                            .AddNew
                            .Fields("WordID") = lngWordID
                            .Fields("WordDescription") = strWord
                            .Update
                        End With
                    End If
                    '--Update count
                    Call UpdateWordCount(lngFoodCode, lngModCode, lngVersion, lngWordID, lngType)
                End If
            Next l
        End If
        
        rst2.MoveNext
    Loop

    rst1.Close
    Set rst1 = Nothing
    rst2.Close
    Set rst2 = Nothing

End Sub

Private Sub AppendIngredients()

    Dim dblWeight As Double
    Dim lngFlag As Long
    Dim lngFoodCode As Long
    Dim lngPortionCode As Long
    Dim lngRetentionCode As Long
    Dim lngVersion As Long
    Dim SQL As String
    Dim strDescription As String
    Dim strMeasure As String
    Dim strRetentionCode As String
    Dim strSRCode As String
    Dim strSRDescr As String
    Dim rst1 As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    
    '--New table
    SQL = "SELECT ingredients.FoodCode," & _
        "ingredients.ModCode," & _
        "ingredients.SeqNum," & _
        "ingredients.Version," & _
        "ingredients.SRCode," & _
        "ingredients.SRDescr," & _
        "ingredients.SRDescrAlt," & _
        "ingredients.ChangeTypeToSRCode," & _
        "ingredients.IngredType," & _
        "ingredients.Amount," & _
        "ingredients.Measure," & _
        "ingredients.PortionCode," & _
        "ingredients.PortionDescr," & _
        "ingredients.RetentionCode," & _
        "ingredients.RetentionDescr," & _
        "ingredients.ChangeTypeToRetnCode," & _
        "ingredients.Flag," & _
        "ingredients.Weight," & _
        "ingredients.ChangeTypeToWeight," & _
        "ingredients.Percentage " & _
        "FROM ingredients " & _
        "WHERE ingredients.FoodCode = 0"
    Set rst1 = New ADODB.Recordset
    rst1.Open SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText
    
    '--Old table
    SQL = "SELECT FnddsSrLinks.FoodCode," & _
        "FnddsSrLinks.SeqNum," & _
        "FnddsSrLinks.Version," & _
        "FnddsSrLinks.SRCode," & _
        "FnddsSrLinks.SRDescription," & _
        "FnddsSrLinks.Amount," & _
        "FnddsSrLinks.Measure," & _
        "FnddsSrLinks.PortionCode," & _
        "FnddsSrLinks.RetentionCode," & _
        "FnddsSrLinks.Flag," & _
        "FnddsSrLinks.Weight," & _
        "FnddsSrLinks.ChangeTypeToSRCode," & _
        "FnddsSrLinks.ChangeTypeToWeight," & _
        "FnddsSrLinks.ChangeTypeToRetnCode " & _
        "FROM FnddsSrLinks " & _
        "ORDER BY FnddsSrLinks.FoodCode," & _
        "FnddsSrLinks.Version," & _
        "FnddsSrLinks.SeqNum"
    Set rst2 = New ADODB.Recordset
    rst2.Open SQL, cnnFNDDS, adOpenStatic, adLockReadOnly, adCmdText
    
    Do Until rst2.EOF
        lngFoodCode = CLng(rst2("FoodCode"))
        lngVersion = CLng(rst2("Version"))
        strSRCode = SRCode(rst2("SRCode"))
        strDescription = rst2("SRDescription")
        strSRDescr = SRDescription(strSRCode, lngVersion, strDescription)
        strMeasure = vbNullString
        If Not IsNull(rst2("Measure")) Then
            strMeasure = Trim$(rst2("Measure"))
        End If
        lngPortionCode = CLng(rst2("PortionCode"))
        lngRetentionCode = CLng(rst2("RetentionCode"))
        If IsNull(rst2("Flag")) Then
            lngFlag = 0
        Else
            lngFlag = CLng(rst2("Flag"))
        End If
        dblWeight = CDbl(rst2("Weight"))
        rst1.AddNew
        rst1("FoodCode") = lngFoodCode
        rst1("ModCode") = 0
        rst1("SeqNum") = rst2("SeqNum")
        rst1("Version") = lngVersion
        rst1("SRCode") = strSRCode
        rst1("SRDescr") = strDescription
        If lngFlag = 2 Then
'            Debug.Print strDescription, "->", strSRDescr
            rst1("SRDescrAlt") = strDescription
        Else
            rst1("SRDescrAlt") = strSRDescr
        End If
        If Not IsNull(rst2("ChangeTypeToSRCode")) Then
            rst1("ChangeTypeToSRCode") = rst2("ChangeTypeToSRCode")
        End If
        If Len(strSRCode) = 5 Then
            rst1("IngredType") = 1
        Else
            rst1("IngredType") = 2
        End If
        rst1("Amount") = rst2("Amount")
        If Len(strMeasure) > 0 Then
            rst1("Measure") = MeasureDescription(strMeasure)
        Else
            rst1("Measure") = "N/A"
        End If
        If lngPortionCode > 0 Then
            rst1("PortionCode") = lngPortionCode
            rst1("PortionDescr") = PortionDescr(lngPortionCode, lngVersion)
        End If
        If lngRetentionCode > 0 Then
            Select Case lngRetentionCode
                Case Is < 10
                    strRetentionCode = "000" & CStr(lngRetentionCode)
                Case Is < 100
                    strRetentionCode = "00" & CStr(lngRetentionCode)
                Case Is < 1000
                    strRetentionCode = "0" & CStr(lngRetentionCode)
                Case Else
                    strRetentionCode = CStr(lngRetentionCode)
            End Select
            rst1("RetentionCode") = strRetentionCode
            rst1("RetentionDescr") = RetentionDescription(strRetentionCode)
        End If
        If Not IsNull(rst2("ChangeTypeToRetnCode")) Then
            rst1("ChangeTypeToRetnCode") = rst2("ChangeTypeToRetnCode")
        End If
        rst1("Flag") = lngFlag
        rst1("Weight") = dblWeight
        If Not IsNull(rst2("ChangeTypeToWeight")) Then
            rst1("ChangeTypeToWeight") = rst2("ChangeTypeToWeight")
        End If
        rst1("Percentage") = Format(dblWeight / InitialWeight(lngFoodCode, lngVersion), "##0.00000000")
        rst1.Update
        rst2.MoveNext
    Loop
    
    rst1.Close
    Set rst1 = Nothing

End Sub

Private Sub AppendIngredSearch()

    Dim lngFlag As Long
    Dim lngFoodCode As Long
    Dim lngIngredType As Long
    Dim lngModCode As Long
    Dim lngSeqNum As Long
    Dim lngVersion As Long
    Dim SQL As String
    Dim strDescription As String
    Dim strFoodModKey As String
    Dim strSRCode As String
    Dim strSRDescr As String
    Dim rst1 As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    
    SQL = "SELECT ingredsearch.FoodCode," & _
        "ingredsearch.ModCode," & _
        "ingredsearch.SeqNum," & _
        "ingredsearch.IngredType," & _
        "ingredsearch.IngrCode," & _
        "ingredsearch.IngrDescr," & _
        "ingredsearch.IngrDescrAlt," & _
        "ingredsearch.Version " & _
        "FROM ingredsearch " & _
        "WHERE ingredsearch.FoodCode = 0"
    Set rst1 = New ADODB.Recordset
    Call rst1.Open(SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText)
    
    '--Old table
    SQL = "SELECT FoodCode," & _
        "0 AS ModCode," & _
        "SRCode," & _
        "SRDescription AS Description," & _
        "Flag," & _
        "Version " & _
        "FROM FnddsSrLinks " & _
        "GROUP BY FoodCode," & _
        "SRCode," & _
        "SRDescription," & _
        "Flag," & _
        "Version " & _
        "ORDER BY FoodCode," & _
        "SRCode,Version"
    Set rst2 = New ADODB.Recordset
    Call rst2.Open(SQL, cnnFNDDS, adOpenStatic, adLockReadOnly, adCmdText)
    
    Do Until rst2.EOF
        lngFoodCode = CLng(rst2("FoodCode"))
        lngModCode = CLng(rst2("ModCode"))
        If StrComp(CStr(lngFoodCode) & "_" & CStr(lngModCode), strFoodModKey, vbTextCompare) = 0 Then
            lngSeqNum = lngSeqNum + 1
        Else
            lngSeqNum = 1
            strFoodModKey = CStr(lngFoodCode) & "_" & CStr(lngModCode)
        End If
        strSRCode = SRCode(rst2("SRCode"))
        strDescription = rst2("Description")
        lngVersion = CLng(rst2("Version"))
        strSRDescr = SRDescription(strSRCode, lngVersion, strDescription)
        If IsNull(rst2("Flag")) Then
            lngFlag = 0
        Else
            lngFlag = CLng(rst2("Flag"))
        End If
        With rst1
            .AddNew
            .Fields("FoodCode") = lngFoodCode
            .Fields("ModCode") = lngModCode
            .Fields("SeqNum") = lngSeqNum
            If Len(strSRCode) = 5 Then
                lngIngredType = 1
            Else
                lngIngredType = 2
            End If
            .Fields("IngredType") = lngIngredType
            .Fields("IngrCode") = strSRCode
            .Fields("IngrDescr") = strDescription
            If lngFlag = 2 Then
'                Debug.Print strDescription, "->", strSRDescr
                .Fields("IngrDescrAlt") = strDescription
            Else
                .Fields("IngrDescrAlt") = strSRDescr
            End If
            .Fields("Version") = lngVersion
            .Update
        End With
        If lngIngredType = 2 Then
            Call UpdateIngredSearch(lngFoodCode, lngModCode, CLng(strSRCode), lngSeqNum, lngVersion, 2, rst1)
        End If
        rst2.MoveNext
    Loop
        
    rst2.Close
    Set rst2 = Nothing
    
    rst1.Close
    Set rst1 = Nothing

End Sub

Private Sub AppendIngredSuggest()

    Dim blnDescribed As Boolean
    Dim l As Long
    Dim lngFoodCode As Long
    Dim lngTermID As Long
    Dim lngType As Long
    Dim lngVersion As Long
    Dim SQL As String
    Dim strDescription As String
    Dim strSRCode As String
    Dim strSRDescr As String
    Dim strTerm As String
    Dim strTerms() As String
    Dim rst1 As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    
    '--New table
    SQL = "SELECT SuggestID," & _
        "SuggestType," & _
        "SuggestDescription " & _
        "FROM suggest " & _
        "WHERE SuggestID = 0"
    Set rst1 = New ADODB.Recordset
    Call rst1.Open(SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText)
    
    '--New ingredient table
    SQL = "SELECT SRCode," & _
        "SRDescrAlt AS SRDescr," & _
        "Version " & _
        "FROM ingredients " & _
        "ORDER BY SRCode," & _
        "Version"
    Set rst2 = New ADODB.Recordset
    Call rst2.Open(SQL, cnnBack, adOpenStatic, adLockReadOnly, adCmdText)

    Do Until rst2.EOF
        strSRCode = Trim$(rst2("SRCode"))
        lngFoodCode = CLng(strSRCode)
        strSRDescr = Trim$(rst2("SRDescr"))
        lngVersion = CLng(rst2("Version"))
        
        '--SR descriptions
        If Len(strSRCode) = 5 Then
            lngType = 3
            strDescription = FormattedSuggestDescr(strSRDescr)
            strTerms = Split(strDescription, ",", , vbTextCompare)
        Else
            lngType = 4
            Select Case lngFoodCode
                Case 11112000
                    'Milk, cow's, fluid, other than whole, NS as to 2%, 1%, or skim (formerly milk, cow's, fluid, "lowfat", NS as to percent fat)
                    ReDim strTerms(6)
                    strTerms(0) = "Milk"
                    strTerms(1) = "cow's"
                    strTerms(2) = "fluid"
                    strTerms(3) = "other than whole"
                    strTerms(4) = "NS as to 2%, 1%, or skim"
                    strTerms(5) = "formerly milk, cow's, fluid, ""lowfat"""
                    strTerms(6) = "NS as to percent fat"
                Case 11511200
                    'Milk, chocolate, reduced fat milk-based, 2% (formerly "lowfat")
                    ReDim strTerms(3)
                    strTerms(0) = "Milk"
                    strTerms(1) = "chocolate"
                    strTerms(2) = "reduced fat milk-based, 2%"
                    strTerms(3) = "formerly ""lowfat"""
                Case 27114000
                    'Beef with (mushroom) soup (mixture)
                    ReDim strTerms(2)
                    strTerms(0) = "Beef"
                    strTerms(1) = "with (mushroom) soup"
                    strTerms(2) = "mixture"
                Case 27144000
                    'Chicken or turkey with (mushroom) soup (mixture)
                    ReDim strTerms(2)
                    strTerms(0) = "Chicken or turkey"
                    strTerms(1) = "with (mushroom) soup"
                    strTerms(2) = "mixture"
                Case 27213400
                    'Beef and rice with (mushroom) soup (mixture)
                    ReDim strTerms(2)
                    strTerms(0) = "Beef and rice"
                    strTerms(1) = "with (mushroom) soup"
                    strTerms(2) = "mixture"
                Case 27243400
                    'Chicken or turkey and rice with (mushroom) soup (mixture)
                    ReDim strTerms(2)
                    strTerms(0) = "Chicken or turkey and rice"
                    strTerms(1) = "with (mushroom) soup"
                    strTerms(2) = "mixture"
                Case 27250830
                    'Fish and rice with (mushroom) soup
                    ReDim strTerms(0)
                    strTerms(0) = "Fish and rice with (mushroom) soup"
                Case 27250900
                    'Fish and noodles with (mushroom) soup
                    ReDim strTerms(0)
                    strTerms(0) = "Fish and noodles with (mushroom) soup"
                Case 28145410
                    'Turkey with gravy, dressing, potatoes, vegetable, cream of tomato soup, dessert (frozen meal)
                    ReDim strTerms(7)
                    strTerms(0) = "Turkey with gravy"
                    strTerms(1) = "dressing"
                    strTerms(2) = "potatoes"
                    strTerms(3) = "vegetable"
                    strTerms(4) = "cream of"
                    strTerms(5) = "tomato soup"
                    strTerms(6) = "dessert"
                    strTerms(7) = "frozen meal"
                Case 53241500
                    'Cookie, butter or sugar cookie
                    ReDim strTerms(2)
                    strTerms(0) = "Cookie"
                    strTerms(1) = "butter or sugar"
                    strTerms(2) = "cookie"
                Case 53241600
                    'Cookie, butter or sugar cookie, with fruit and/or nuts
                    ReDim strTerms(3)
                    strTerms(0) = "Cookie"
                    strTerms(1) = "butter or sugar"
                    strTerms(2) = "cookie"
                    strTerms(3) = "with fruit and/or nuts"
                Case 54101010
                    'Cracker, animal
                    ReDim strTerms(0)
                    strTerms(0) = "Cracker, animal"
                Case 56205410
                    'Rice, white, cooked with (fat) oil, Puerto Rican style (Arroz blanco)
                    ReDim strTerms(4)
                    strTerms(0) = "Rice"
                    strTerms(1) = "white"
                    strTerms(2) = "cooked with (fat) oil"
                    strTerms(3) = "Puerto Rican style"
                    strTerms(4) = "Arroz blanco"
                Case 58126180
                    'Turnover, meat-, potato-, and vegetable-filled, no gravy
                    ReDim strTerms(2)
                    strTerms(0) = "Turnover"
                    strTerms(1) = "meat-, potato-, and vegetable-filled"
                    strTerms(2) = "no gravy"
                Case 58132310
                    'Spaghetti with tomato sauce and meatballs or spaghetti with meat sauce or spaghetti with meat sauce and meatballs
                    ReDim strTerms(2)
                    strTerms(0) = "Spaghetti with tomato sauce and meatballs"
                    strTerms(1) = "spaghetti with meat sauce"
                    strTerms(2) = "spaghetti with meat sauce and meatballs"
                Case 63320100
                    If lngVersion < 4 Then
                        'Fruit salad, Puerto Rican style (Mixture includes bananas, papayas, oranges, grapefruit, etc.) (Ensalada de frutas tropicales)
                        ReDim strTerms(3)
                        strTerms(0) = "Fruit salad"
                        strTerms(1) = "Puerto Rican style"
                        strTerms(2) = "Mixture includes bananas, papayas, oranges, grapefruit, etc."
                        strTerms(3) = "Ensalada de frutas tropicales"
                    Else
                        'Fruit salad, Puerto Rican style (Mixture includes bananas, papayas, oranges, etc.) (Ensalada de frutas tropicales)
                        ReDim strTerms(3)
                        strTerms(0) = "Fruit salad"
                        strTerms(1) = "Puerto Rican style"
                        strTerms(2) = "Mixture includes bananas, papayas, oranges, etc."
                        strTerms(3) = "Ensalada de frutas tropicales"
                    End If
                Case 75340000
                    'Vegetable combinations, Oriental style, (broccoli, green pepper, water chestnut, etc) cooked, NS as to fat added in cooking
                    ReDim strTerms(4)
                    strTerms(0) = "Vegetable combinations"
                    strTerms(1) = "Oriental style"
                    strTerms(2) = "broccoli, green pepper, water chestnut, etc"
                    strTerms(3) = "cooked"
                    strTerms(4) = "NS as to fat added in cooking"
                Case 75340010
                    'Vegetable combinations, Oriental style, (broccoli, green pepper,  water chestnuts, etc), cooked, fat not added in cooking
                    ReDim strTerms(4)
                    strTerms(0) = "Vegetable combinations"
                    strTerms(1) = "Oriental style"
                    strTerms(2) = "broccoli, green pepper,  water chestnuts, etc"
                    strTerms(3) = "cooked"
                    strTerms(4) = "fat not added in cooking"
                Case 75340020
                    'Vegetable combinations, Oriental style, (broccoli, green pepper, water chestnuts, etc), cooked, fat added in cooking
                    ReDim strTerms(4)
                    strTerms(0) = "Vegetable combinations"
                    strTerms(1) = "Oriental style"
                    strTerms(2) = "broccoli, green pepper, water chestnuts, etc"
                    strTerms(3) = "cooked"
                    strTerms(4) = "fat added in cooking"
                Case 75340100
                    'Vegetable combinations (broccoli, carrots, corn, cauliflower, etc.), cooked, NS as to fat added in cooking
                    ReDim strTerms(3)
                    strTerms(0) = "Vegetable combinations"
                    strTerms(1) = "broccoli, carrots, corn, cauliflower, etc."
                    strTerms(2) = "cooked"
                    strTerms(3) = "NS as to fat added in cooking"
                Case 75340110
                    'Vegetable combinations (broccoli, carrots, corn, cauliflower, etc.), cooked, fat not added in cooking
                    ReDim strTerms(3)
                    strTerms(0) = "Vegetable combinations"
                    strTerms(1) = "broccoli, carrots, corn, cauliflower, etc."
                    strTerms(2) = "cooked"
                    strTerms(3) = "fat not added in cooking"
                Case 75340120
                    'Vegetable combinations (broccoli, carrots, corn, cauliflower, etc.), cooked, fat added in cooking
                    ReDim strTerms(3)
                    strTerms(0) = "Vegetable combinations"
                    strTerms(1) = "broccoli, carrots, corn, cauliflower, etc."
                    strTerms(2) = "cooked"
                    strTerms(3) = "fat added in cooking"
                Case 75340160
                    'Vegetable and pasta combinations with cream or cheese sauce (broccoli, pasta, carrots, corn, zucchini, peppers, cauliflower, peas, etc.), cooked
                    ReDim strTerms(2)
                    strTerms(0) = "Vegetable and pasta combinations with cream or cheese sauce"
                    strTerms(1) = "broccoli, pasta, carrots, corn, zucchini, peppers, cauliflower, peas, etc."
                    strTerms(2) = "cooked"
                Case 75340300
                    'Pinacbet (eggplant with tomatoes, bitter melon, etc.)
                    ReDim strTerms(1)
                    strTerms(0) = "Pinacbet"
                    strTerms(1) = "eggplant with tomatoes, bitter melon, etc."
                Case 81203200
                    'Shortening, animal
                    ReDim strTerms(0)
                    strTerms(0) = "Shortening, animal"
                Case 81302030
                    'Orange sauce (for duck)
                    ReDim strTerms(0)
                    strTerms(0) = "Orange sauce (for duck)"
                Case 82105800
                    'Canola, soybean and sunflower oil
                    ReDim strTerms(0)
                    strTerms(0) = "Canola, soybean and sunflower oil"
                Case 83100100
                    'Salad dressing, NFS, for salads
                    ReDim strTerms(0)
                    strTerms(0) = "Salad dressing, NFS, for salads"
                Case 83100200
                    'Salad dressing, NFS, for sandwiches
                    ReDim strTerms(0)
                    strTerms(0) = "Salad dressing, NFS, for sandwiches"
                Case 91511090
                    'Gelatin dessert, dietetic, with fruit and vegetable(s), sweetened with low calorie sweetener
                    ReDim strTerms(3)
                    strTerms(0) = "Gelatin dessert"
                    strTerms(1) = "dietetic"
                    strTerms(2) = "with fruit and vegetable(s)"
                    strTerms(3) = "sweetened with low calorie sweetener"
                Case 91520100
                    'Yookan (Yokan), a Japanese dessert made with bean paste and sugar
                    ReDim strTerms(3)
                    strTerms(0) = "Yookan"
                    strTerms(1) = "Yokan"
                    strTerms(2) = "Japanese dessert"
                    strTerms(3) = "made with bean paste and sugar"
                Case Else
                    strDescription = FormattedSuggestDescr(strSRDescr)
                    strTerms = Split(strDescription, ",", , vbTextCompare)
            End Select
        End If
        strTerms = FormattedSuggestTerms(strTerms)
        For l = 0 To UBound(strTerms())
            strTerm = Trim$(LCase(strTerms(l)))
            If Len(strTerm) > 0 Then
                'Debug.Print strTerm
                lngTermID = SuggestTermExists(lngType, strTerm)
                If lngTermID = 0 Then
                    '--Add term
                    lngTermID = SuggestTermID(lngType) + 1
                    With rst1
                        .AddNew
                        .Fields("SuggestID") = lngTermID
                        .Fields("SuggestType") = lngType
                        .Fields("SuggestDescription") = strTerm
                        .Update
                    End With
                End If
                '--Update count
                Call UpdateIngredSuggestCount(strSRCode, lngVersion, lngTermID, lngType)
            End If
        Next l
        rst2.MoveNext
    Loop

    rst1.Close
    Set rst1 = Nothing
    rst2.Close
    Set rst2 = Nothing

End Sub

Private Sub AppendNutrientDescr()

    Dim l As Long
    Dim lngDecimals As Long
    Dim lngNutrientCode As Long
    Dim lngOrder As Long
    Dim lngVersion As Long
    Dim SQL As String
    Dim strDescription As String
    Dim strTagname As String
    Dim strUnit As String
    Dim fld As ADODB.Field
    Dim rst1 As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    Dim rst3 As ADODB.Recordset
    
    '--New table
    SQL = "SELECT Tagname," & _
        "Version," & _
        "NutrientDescription," & _
        "Unit," & _
        "Decimals," & _
        "DisplayOrder " & _
        "FROM nutrientdescr " & _
        "WHERE (Tagname IS NULL)"
    Set rst1 = New ADODB.Recordset
    Call rst1.Open(SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText)
    
    '--Old table of nutrient descriptions
    SQL = "SELECT NutrientCode," & _
        "Version," & _
        "NutrientDescription," & _
        "Tagname," & _
        "Unit," & _
        "Decimals " & _
        "FROM NutDesc " & _
        "ORDER BY NutrientCode," & _
        "Version"
    Set rst2 = New ADODB.Recordset
    Call rst2.Open(SQL, cnnFNDDS, adOpenStatic, adLockReadOnly, adCmdText)
    
    '--Update nutrients
    Do Until rst2.EOF
        lngNutrientCode = CLng(rst2("NutrientCode"))
        lngVersion = CLng(rst2("Version"))
        strDescription = rst2("NutrientDescription")
        strTagname = vbNullString
        If IsNull(rst2("Tagname")) Then
            '-- Take care of 2 nutrients without tagnames
            If lngNutrientCode = 573 Then
                strTagname = "TOCPHA_ADDED"
            ElseIf lngNutrientCode = 578 Then
                strTagname = "VITB12_ADDED"
            Else
                Stop
            End If
        Else
            strTagname = rst2("Tagname")
            '-- Take care of 3 nutrients whose tagnames do not match INFOODS
            If lngNutrientCode = 208 Then
                strTagname = "ENERC"
            ElseIf lngNutrientCode = 320 Then
                strTagname = "VITA"
            ElseIf lngNutrientCode = 430 Then
                strTagname = "VITK"
            ElseIf StrComp(strTagname, "LUTN", vbTextCompare) = 0 Or StrComp(strTagname, "LUT+ZEA", vbTextCompare) = 0 Then
                strTagname = "LUTNZEA"
            End If
        End If
        strUnit = rst2("Unit")
        lngDecimals = CLng(rst2("Decimals"))
        lngOrder = NutrientSortOrder(strTagname)
        With rst1
            .AddNew
            .Fields("Version") = lngVersion
            .Fields("NutrientDescription") = strDescription
            .Fields("Tagname") = strTagname
            .Fields("Unit") = strUnit
            .Fields("Decimals") = lngDecimals
            .Fields("DisplayOrder") = lngOrder
            .Update
        End With
        rst2.MoveNext
    Loop
    
    rst2.Close
    Set rst2 = Nothing
    rst1.Close
    Set rst1 = Nothing

End Sub

Private Sub AppendNutrients()

    Dim dblNutrientValue As Double
    Dim lngFoodCode As Long
    Dim lngModCode As Long
    Dim lngNutrientCode As Long
    Dim lngVersion As Long
    Dim SQL As String
    Dim strTagname As String
    Dim rst1 As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    
    '--New table
    SQL = "SELECT FoodCode," & _
        "ModCode," & _
        "Tagname," & _
        "Version," & _
        "NutrientValue " & _
        "FROM nutrients " & _
        "WHERE FoodCode = 0"
    Set rst1 = New ADODB.Recordset
    Call rst1.Open(SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText)
    
    '--Old table (nutrients)
    SQL = "SELECT FoodCode," & _
        "0 AS ModCode," & _
        "NutrientCode," & _
        "Version," & _
        "NutrientValue " & _
        "FROM FnddsNutVal " & _
        "ORDER BY FoodCode," & _
        "NutrientCode," & _
        "Version"
    Set rst2 = New ADODB.Recordset
    Call rst2.Open(SQL, cnnFNDDS, adOpenStatic, adLockReadOnly, adCmdText)
    
    '--Update nutrients
    Do Until rst2.EOF
        lngFoodCode = CLng(rst2("FoodCode"))
        lngModCode = CLng(rst2("ModCode"))
        lngNutrientCode = CLng(rst2("NutrientCode"))
        lngVersion = CLng(rst2("Version"))
        strTagname = NutrientTagname(lngNutrientCode, lngVersion)
        dblNutrientValue = CDbl(rst2("NutrientValue"))
        
        With rst1
            .AddNew
            .Fields("FoodCode") = lngFoodCode
            .Fields("ModCode") = lngModCode
            .Fields("Tagname") = strTagname
            .Fields("Version") = lngVersion
            .Fields("NutrientValue") = dblNutrientValue
            .Update
        End With
        
        rst2.MoveNext
    Loop
    
    rst2.Close
    Set rst2 = Nothing
    
    '--Old table (modified recipes)
    SQL = "SELECT FoodCode," & _
        "ModificationCode AS ModCode," & _
        "NutrientCode," & _
        "Version," & _
        "NutrientValue " & _
        "FROM ModNutVal " & _
        "ORDER BY FoodCode," & _
        "ModificationCode," & _
        "NutrientCode," & _
        "Version"
    Set rst2 = New ADODB.Recordset
    Call rst2.Open(SQL, cnnFNDDS, adOpenStatic, adLockReadOnly, adCmdText)
    
    '--Update nutrients
    Do Until rst2.EOF
        lngFoodCode = CLng(rst2("FoodCode"))
        lngModCode = CLng(rst2("ModCode"))
        lngNutrientCode = CLng(rst2("NutrientCode"))
        lngVersion = CLng(rst2("Version"))
        strTagname = NutrientTagname(lngNutrientCode, lngVersion)
        dblNutrientValue = CDbl(rst2("NutrientValue"))
        
        With rst1
            .AddNew
            .Fields("FoodCode") = lngFoodCode
            .Fields("ModCode") = lngModCode
            .Fields("Tagname") = strTagname
            .Fields("Version") = lngVersion
            .Fields("NutrientValue") = dblNutrientValue
            .Update
        End With
        
        rst2.MoveNext
    Loop
    
    rst2.Close
    Set rst2 = Nothing
    
    rst1.Close
    Set rst1 = Nothing

End Sub

Private Sub AppendPortions()
    
    Dim lngFoodCode As Long
    Dim lngModCode As Long
    Dim lngPortionCode As Long
    Dim lngSubcode As Long
    Dim lngVersion As Long
    Dim SQL As String
    Dim strPortionChangeType As String
    Dim rst1 As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    
    '--New table
    SQL = "SELECT portions.FoodCode," & _
        "portions.ModCode," & _
        "portions.Subcode," & _
        "portions.SubcodeDescr," & _
        "portions.SeqNum," & _
        "portions.Version," & _
        "portions.PortionCode," & _
        "portions.PortionDescr," & _
        "portions.PortionChangeType," & _
        "portions.Weight," & _
        "portions.WeightChangeType " & _
        "FROM portions " & _
        "WHERE portions.FoodCode = 0"
    Set rst1 = New ADODB.Recordset
    rst1.Open SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText
    
    '--Old table
    SQL = "SELECT FoodWeights.FoodCode," & _
        "FoodWeights.Subcode," & _
        "FoodWeights.SeqNum," & _
        "FoodWeights.PortionCode," & _
        "FoodWeights.Version," & _
        "FoodWeights.PortionWeight," & _
        "FoodWeights.ChangeType " & _
        "FROM FoodWeights " & _
        "ORDER BY FoodWeights.FoodCode," & _
        "FoodWeights.Subcode," & _
        "FoodWeights.SeqNum," & _
        "FoodWeights.Version"
    Set rst2 = New ADODB.Recordset
    rst2.Open SQL, cnnFNDDS, adOpenStatic, adLockReadOnly, adCmdText
    
    Do Until rst2.EOF
        lngPortionCode = CLng(rst2("PortionCode"))
        lngSubcode = CLng(rst2("Subcode"))
        lngVersion = CLng(rst2("Version"))
        rst1.AddNew
        rst1("FoodCode") = rst2("FoodCode")
        rst1("Subcode") = rst2("Subcode")
        rst1("SubcodeDescr") = SubcodeDescr(CLng(rst2("Subcode")), lngVersion)
        rst1("SeqNum") = rst2("SeqNum")
        rst1("Version") = lngVersion
        rst1("PortionCode") = lngPortionCode
        rst1("PortionDescr") = PortionDescr(lngPortionCode, lngVersion)
        strPortionChangeType = PortionChangeType(lngPortionCode, lngVersion)
        If Len(strPortionChangeType) > 0 Then
            rst1("PortionChangeType") = strPortionChangeType
        End If
        If CDbl(rst2("PortionWeight")) > 0 Then
            rst1("Weight") = rst2("PortionWeight")
        Else
            Debug.Print "Invalid Portion Weight", rst1("FoodCode"), rst1("Subcode"), rst1("SeqNum"), rst1("Version"), rst1("PortionCode"), rst2("PortionWeight")
            rst1("Weight") = -1
        End If
        If Not IsNull(rst2("ChangeType")) Then
            rst1("WeightChangeType") = rst2("ChangeType")
        End If
        rst1.Update
        rst2.MoveNext
    Loop
    
    rst2.Close
    Set rst2 = Nothing
    
    '--Recipe Mods
    SQL = "SELECT FoodCode," & _
        "ModCode," & _
        "Version " & _
        "FROM fooddescr " & _
        "WHERE (ModCode > 0) " & _
        "ORDER BY FoodCode," & _
        "ModCode," & _
        "Version"
    Set rst2 = New ADODB.Recordset
    rst2.Open SQL, cnnBack, adOpenStatic, adLockReadOnly, adCmdText
    
    Do Until rst2.EOF
        lngFoodCode = CLng(rst2("FoodCode"))
        lngModCode = CLng(rst2("ModCode"))
        lngVersion = CLng(rst2("Version"))
        comPortions_Lkp("@FoodCode") = lngFoodCode
        comPortions_Lkp("@Version") = lngVersion
        rstPortions_Lkp.Requery
        If rstPortions_Lkp.RecordCount > 0 Then
            Do Until rstPortions_Lkp.EOF
                rst1.AddNew
                rst1("FoodCode") = lngFoodCode
                rst1("ModCode") = lngModCode
                rst1("Subcode") = rstPortions_Lkp("Subcode")
                rst1("SubcodeDescr") = SubcodeDescr(CLng(rstPortions_Lkp("Subcode")), lngVersion)
                rst1("SeqNum") = rstPortions_Lkp("SeqNum")
                rst1("Version") = lngVersion
                lngPortionCode = CLng(rstPortions_Lkp("PortionCode"))
                rst1("PortionCode") = lngPortionCode
                rst1("PortionDescr") = PortionDescr(lngPortionCode, lngVersion)
                strPortionChangeType = PortionChangeType(lngPortionCode, lngVersion)
                If Len(strPortionChangeType) > 0 Then
                    rst1("PortionChangeType") = strPortionChangeType
                End If
                If CDbl(rstPortions_Lkp("Weight")) > 0 Then
                    rst1("Weight") = rstPortions_Lkp("Weight")
                Else
                    Debug.Print "Invalid Portion Weight", lngFoodCode, lngModCode, rstPortions_Lkp("Subcode"), rstPortions_Lkp("SeqNum"), lngVersion, rstPortions_Lkp("PortionCode"), rstPortions_Lkp("Weight")
                    rst1("Weight") = -1
                End If
                If Not IsNull(rstPortions_Lkp("WeightChangeType")) Then
                    rst1("WeightChangeType") = rstPortions_Lkp("WeightChangeType")
                End If
                rst1.Update
                rstPortions_Lkp.MoveNext
            Loop
        End If
        rst2.MoveNext
    Loop
    
    rst2.Close
    Set rst2 = Nothing
    rst1.Close
    Set rst1 = Nothing

End Sub

Private Sub AppendTagname()

    Dim l As Long
    Dim SQL As String
    Dim strComments As String
    Dim strDescription As String
    Dim strExamples As String
    Dim strKeywords As String
    Dim strNotes As String
    Dim strSynonyns As String
    Dim strTables As String
    Dim strTagname As String
    Dim strUnits As String
    Dim rst1 As ADODB.Recordset
    Dim wbk As Excel.Workbook
    Dim wst As Excel.Worksheet
    
    '--New table
    SQL = "SELECT Tagname," & _
        "TagnameDescription," & _
        "Units," & _
        "Tables," & _
        "Synonyms," & _
        "Keywords," & _
        "Examples," & _
        "Comments," & _
        "Notes " & _
        "FROM tagname " & _
        "WHERE (Tagname IS NULL)"
    Set rst1 = New ADODB.Recordset
    Call rst1.Open(SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText)
    
    '--Excel spreadsheet of tagname info
    With appExcel
        Set wbk = .Workbooks.Open(fso.BuildPath(DATABASES_PATH, "RawData\INFOODS\tagnames\Tagnames.xlsm"))
        .Visible = True
    End With
    Set wst = wbk.Worksheets("Sheet1")
    
    '--Update tagnames
    For l = 2 To wst.UsedRange.Rows.Count
        strTagname = wst.Cells(l, 1)
        strDescription = wst.Cells(l, 2)
        strUnits = wst.Cells(l, 3)
        strTables = wst.Cells(l, 4)
        strSynonyns = wst.Cells(l, 5)
        strKeywords = wst.Cells(l, 6)
        strExamples = wst.Cells(l, 7)
        strComments = wst.Cells(l, 8)
        strNotes = wst.Cells(l, 9)
        
        With rst1
            .AddNew
            .Fields("Tagname") = strTagname
            .Fields("TagnameDescription") = strDescription
            .Fields("Units") = strUnits
            If Len(strTables) > 0 Then
                .Fields("Tables") = strTables
            End If
            If Len(strSynonyns) > 0 Then
                .Fields("Synonyms") = strSynonyns
            End If
            If Len(strKeywords) > 0 Then
                .Fields("Keywords") = strKeywords
            End If
            If Len(strExamples) > 0 Then
                .Fields("Examples") = strExamples
            End If
            If Len(strComments) > 0 Then
                .Fields("Comments") = strComments
            End If
            If Len(strNotes) > 0 Then
                .Fields("Notes") = strNotes
            End If
            .Update
        End With
    Next l
    
    rst1.Close
    Set rst1 = Nothing
    
    Set wst = Nothing
    wbk.Close
    Set wbk = Nothing

End Sub

Private Sub CloseCommands()

    If Not (rstAddtlDescr_Lkp Is Nothing) Then
        If rstAddtlDescr_Lkp.State = adStateOpen Then rstAddtlDescr_Lkp.Close
        Set rstAddtlDescr_Lkp = Nothing
    End If
    Set comAddtlDescr_Lkp = Nothing
    If Not (rstCountInDocument_Lkp Is Nothing) Then
        If rstCountInDocument_Lkp.State = adStateOpen Then rstCountInDocument_Lkp.Close
        Set rstCountInDocument_Lkp = Nothing
    End If
    Set comCountInDocument_Lkp = Nothing
    If Not (rstCountInDocuments_Lkp Is Nothing) Then
        If rstCountInDocuments_Lkp.State = adStateOpen Then rstCountInDocuments_Lkp.Close
        Set rstCountInDocuments_Lkp = Nothing
    End If
    Set comCountInDocuments_Lkp = Nothing
    If Not (rstDocumentCount_Lkp Is Nothing) Then
        If rstDocumentCount_Lkp.State = adStateOpen Then rstDocumentCount_Lkp.Close
        Set rstDocumentCount_Lkp = Nothing
    End If
    Set comDocumentCount_Lkp = Nothing
    If Not (rstFCDescr_Lkp Is Nothing) Then
        If rstFCDescr_Lkp.State = adStateOpen Then rstFCDescr_Lkp.Close
        Set rstFCDescr_Lkp = Nothing
    End If
    Set comFCDescr_Lkp = Nothing
'    If Not (rstFoodDescr_Lkp Is Nothing) Then
'        If rstFoodDescr_Lkp.State = adStateOpen Then rstFoodDescr_Lkp.Close
'        Set rstFoodDescr_Lkp = Nothing
'    End If
'    Set comFoodDescr_Lkp = Nothing
    If Not (rstFoodMatrixA_Lkp Is Nothing) Then
        If rstFoodMatrixA_Lkp.State = adStateOpen Then rstFoodMatrixA_Lkp.Close
        Set rstFoodMatrixA_Lkp = Nothing
    End If
    Set comFoodMatrixA_Lkp = Nothing
    If Not (rstFoodMatrixB_Lkp Is Nothing) Then
        If rstFoodMatrixB_Lkp.State = adStateOpen Then rstFoodMatrixB_Lkp.Close
        Set rstFoodMatrixB_Lkp = Nothing
    End If
    Set comFoodMatrixB_Lkp = Nothing
    If Not (rstFoodMatrixValue_Lkp Is Nothing) Then
        If rstFoodMatrixValue_Lkp.State = adStateOpen Then rstFoodMatrixValue_Lkp.Close
        Set rstFoodMatrixValue_Lkp = Nothing
    End If
    Set comFoodMatrixValue_Lkp = Nothing
'    If Not (rstIngredients_Lkp Is Nothing) Then
'        If rstIngredients_Lkp.State = adStateOpen Then rstIngredients_Lkp.Close
'        Set rstIngredients_Lkp = Nothing
'    End If
'    Set comIngredients_Lkp = Nothing
'    If Not (rstIngredRecipe_Lkp Is Nothing) Then
'        If rstIngredRecipe_Lkp.State = adStateOpen Then rstIngredRecipe_Lkp.Close
'        Set rstIngredRecipe_Lkp = Nothing
'    End If
'    Set comIngredRecipe_Lkp = Nothing
'    If Not (rstIngredSearch_Lkp Is Nothing) Then
'        If rstIngredSearch_Lkp.State = adStateOpen Then rstIngredSearch_Lkp.Close
'        Set rstIngredSearch_Lkp = Nothing
'    End If
'    Set comIngredSearch_Lkp = Nothing
'    If Not (rstModNutrient_Lkp Is Nothing) Then
'        If rstModNutrient_Lkp.State = adStateOpen Then rstModNutrient_Lkp.Close
'        Set rstModNutrient_Lkp = Nothing
'    End If
'    Set comModNutrient_Lkp = Nothing
'    If Not (rstMPED_Lkp Is Nothing) Then
'        If rstMPED_Lkp.State = adStateOpen Then rstMPED_Lkp.Close
'        Set rstMPED_Lkp = Nothing
'    End If
'    Set comMPED_Lkp = Nothing
'    If Not (rstNutrient_Lkp Is Nothing) Then
'        If rstNutrient_Lkp.State = adStateOpen Then rstNutrient_Lkp.Close
'        Set rstNutrient_Lkp = Nothing
'    End If
'    Set comNutrient_Lkp = Nothing
    If Not (rstPortionDescr_Lkp Is Nothing) Then
        If rstPortionDescr_Lkp.State = adStateOpen Then rstPortionDescr_Lkp.Close
        Set rstPortionDescr_Lkp = Nothing
    End If
    Set comPortionDescr_Lkp = Nothing
    If Not (rstPortions_Lkp Is Nothing) Then
        If rstPortions_Lkp.State = adStateOpen Then rstPortions_Lkp.Close
        Set rstPortions_Lkp = Nothing
    End If
    Set comPortions_Lkp = Nothing
    If Not (rstRecipeWeight_Lkp Is Nothing) Then
        If rstRecipeWeight_Lkp.State = adStateOpen Then rstRecipeWeight_Lkp.Close
        Set rstRecipeWeight_Lkp = Nothing
    End If
    Set comRecipeWeight_Lkp = Nothing
    If Not (rstRetDescr_Lkp Is Nothing) Then
        If rstRetDescr_Lkp.State = adStateOpen Then rstRetDescr_Lkp.Close
        Set rstRetDescr_Lkp = Nothing
    End If
    Set comRetDescr_Lkp = Nothing
'    If Not (rstSimilarRecipe_Lkp Is Nothing) Then
'        If rstSimilarRecipe_Lkp.State = adStateOpen Then rstSimilarRecipe_Lkp.Close
'        Set rstSimilarRecipe_Lkp = Nothing
'    End If
'    Set comSimilarRecipe_Lkp = Nothing
    If Not (rstSRDescr_Lkp Is Nothing) Then
        If rstSRDescr_Lkp.State = adStateOpen Then rstSRDescr_Lkp.Close
        Set rstSRDescr_Lkp = Nothing
    End If
    Set comSRDescr_Lkp = Nothing
    If Not (rstSubcode_Lkp Is Nothing) Then
        If rstSubcode_Lkp.State = adStateOpen Then rstSubcode_Lkp.Close
        Set rstSubcode_Lkp = Nothing
    End If
    Set comSubcode_Lkp = Nothing
    If Not (rstSuggest_Lkp Is Nothing) Then
        If rstSuggest_Lkp.State = adStateOpen Then rstSuggest_Lkp.Close
        Set rstSuggest_Lkp = Nothing
    End If
    Set comSuggest_Lkp = Nothing
    If Not (rstSuggestID_Lkp Is Nothing) Then
        If rstSuggestID_Lkp.State = adStateOpen Then rstSuggestID_Lkp.Close
        Set rstSuggestID_Lkp = Nothing
    End If
    Set comSuggestID_Lkp = Nothing
    If Not (rstSuggestFoodCount_Lkp Is Nothing) Then
        If rstSuggestFoodCount_Lkp.State = adStateOpen Then rstSuggestFoodCount_Lkp.Close
        Set rstSuggestFoodCount_Lkp = Nothing
    End If
    Set comSuggestFoodCount_Lkp = Nothing
    If Not (rstSuggestIngredCount_Lkp Is Nothing) Then
        If rstSuggestIngredCount_Lkp.State = adStateOpen Then rstSuggestIngredCount_Lkp.Close
        Set rstSuggestIngredCount_Lkp = Nothing
    End If
    Set comSuggestIngredCount_Lkp = Nothing
    If Not (rstUpdateWordCount Is Nothing) Then
        If rstUpdateWordCount.State = adStateOpen Then rstUpdateWordCount.Close
        Set rstUpdateWordCount = Nothing
    End If
    Set comUpdateWordCount = Nothing
    If Not (rstTagname_Lkp Is Nothing) Then
        If rstTagname_Lkp.State = adStateOpen Then rstTagname_Lkp.Close
        Set rstTagname_Lkp = Nothing
    End If
    Set comTagname_Lkp = Nothing
    If Not (rstWord_Lkp Is Nothing) Then
        If rstWord_Lkp.State = adStateOpen Then rstWord_Lkp.Close
        Set rstWord_Lkp = Nothing
    End If
    Set comWord_Lkp = Nothing
    If Not (rstWordID_Lkp Is Nothing) Then
        If rstWordID_Lkp.State = adStateOpen Then rstWordID_Lkp.Close
        Set rstWordID_Lkp = Nothing
    End If
    Set comWordID_Lkp = Nothing
    If Not (rstWordCount_Lkp Is Nothing) Then
        If rstWordCount_Lkp.State = adStateOpen Then rstWordCount_Lkp.Close
        Set rstWordCount_Lkp = Nothing
    End If
    Set comWordCount_Lkp = Nothing
    If Not (rstWordsInDoc_Lkp Is Nothing) Then
        If rstWordsInDoc_Lkp.State = adStateOpen Then rstWordsInDoc_Lkp.Close
        Set rstWordsInDoc_Lkp = Nothing
    End If
    Set comWordsInDoc_Lkp = Nothing

End Sub

Private Function CountInDocuments(WordID As Long, Version As Long, WordType As Long) As Long

    With comCountInDocuments_Lkp
        .Parameters("@WordID") = WordID
        .Parameters("@Version") = Version
        .Parameters("@WordType") = WordType
    End With
    With rstCountInDocuments_Lkp
        .Requery
        If .RecordCount > 0 Then
            CountInDocuments = CLng(.Fields("DocumentCount"))
        End If
    End With

End Function

Public Sub CreateAutoCompleteFiles()

    Dim blnFirstRow As Boolean
    Dim l As Long
    Dim lngFrequency As Long
    Dim lngVersion As Long
    Dim SQL As String
    Dim strFullpath As String
    Dim strTerm As String
    Dim strTermPrevious As String
    
    Dim rst As ADODB.Recordset
    Dim txt As Scripting.TextStream
    
    For l = 1 To 2
        strFullpath = fso.BuildPath("D:\workspace\java\fndds\WebContent\WEB-INF\resources", IIf(l = 1, "foods", "includes") & ".xml")
        
        If fso.FileExists(strFullpath) Then
            Call fso.DeleteFile(strFullpath)
        End If
        
        SQL = "SELECT suggest.SuggestDescription, foodsuggest.Version, SUM(foodsuggest.SuggestCount) AS Frequency " & _
            "FROM suggest INNER JOIN foodsuggest ON suggest.SuggestID = foodsuggest.SuggestID AND " & _
            "suggest.SuggestType = foodsuggest.SuggestType " & _
            "WHERE (suggest.SuggestType = 1) "
            
'        SQL = "SELECT suggest.SuggestDescription, foodsuggest.Version, SUM(foodsuggest.SuggestCount) AS Frequency, foodsuggest.FoodCode " & _
            "FROM suggest INNER JOIN foodsuggest ON suggest.SuggestID = foodsuggest.SuggestID AND " & _
            "suggest.SuggestType = dbo.foodsuggest.SuggestType " & _
            "WHERE (dbo.suggest.SuggestType = 1) "
            
        If l = 2 Then
            SQL = SQL & "OR (suggest.SuggestType = 2) "
        End If
        SQL = SQL & "GROUP BY suggest.SuggestDescription, foodsuggest.Version " & _
            "ORDER BY suggest.SuggestDescription, foodsuggest.Version"
'        SQL = SQL & "GROUP BY suggest.SuggestDescription, foodsuggest.Version, foodsuggest.FoodCode " & _
            "ORDER BY suggest.SuggestDescription, foodsuggest.Version, foodsuggest.FoodCode"

        Set rst = New ADODB.Recordset
        Call rst.Open(SQL, cnnBack, adOpenStatic, adLockReadOnly, adCmdText)
        
        Set txt = fso.CreateTextFile(strFullpath, False)
        
        Call CreateAutoCompleteFileHeader(txt)
        
        blnFirstRow = True
        strTermPrevious = vbNullString
        Do Until rst.EOF
            strTerm = rst("SuggestDescription")
            lngVersion = rst("Version")
            lngFrequency = rst("Frequency")

            strTerm = DescriptionWithoutSpecialCharacters(strTerm)

            Call CreateAutoCompleteFileBody(strTerm, strTermPrevious, lngVersion, lngFrequency, blnFirstRow, txt)
            
            If blnFirstRow Then blnFirstRow = False
            strTermPrevious = strTerm
            rst.MoveNext
        Loop
        
        Call CreateAutoCompleteFileFooter(txt)
        
        txt.Close
        Set txt = Nothing
        
        rst.Close
        Set rst = Nothing
    Next l

    strFullpath = fso.BuildPath("D:\workspace\java\fndds\WebContent\WEB-INF\resources", "ingredients.xml")
    
    If fso.FileExists(strFullpath) Then
        Call fso.DeleteFile(strFullpath)
    End If
    
'    SQL = "SELECT suggest.SuggestDescription, ingredsuggest.Version, SUM(ingredsuggest.SuggestCount) AS Frequency " & _
        "FROM ingredsuggest INNER JOIN suggest ON " & _
        "ingredsuggest.SuggestID = suggest.SuggestID AND ingredsuggest.SuggestType = suggest.SuggestType " & _
        "WHERE (ingredsuggest.SuggestType = 3) OR (ingredsuggest.SuggestType = 4) " & _
        "GROUP BY suggest.SuggestDescription, ingredsuggest.Version " & _
        "ORDER BY suggest.SuggestDescription, ingredsuggest.Version"
    SQL = "SELECT suggest.SuggestDescription, ingredsuggest.Version, SUM(ingredsuggest.SuggestCount) AS Frequency, ingredsuggest.SRCode " & _
        "FROM ingredsuggest INNER JOIN suggest ON " & _
        "ingredsuggest.SuggestID = suggest.SuggestID AND ingredsuggest.SuggestType = suggest.SuggestType " & _
        "WHERE (ingredsuggest.SuggestType = 3) OR (ingredsuggest.SuggestType = 4) " & _
        "GROUP BY suggest.SuggestDescription, ingredsuggest.Version, ingredsuggest.SRCode " & _
        "ORDER BY suggest.SuggestDescription, ingredsuggest.Version, ingredsuggest.SRCode"
    Set rst = New ADODB.Recordset
    Call rst.Open(SQL, cnnBack, adOpenStatic, adLockReadOnly, adCmdText)
    
    Set txt = fso.CreateTextFile(strFullpath, False)
    
    Call CreateAutoCompleteFileHeader(txt)
    
    blnFirstRow = True
    strTermPrevious = vbNullString
    Do Until rst.EOF
        strTerm = rst("SuggestDescription")
        lngVersion = rst("Version")
        lngFrequency = rst("Frequency")
        
        strTerm = DescriptionWithoutSpecialCharacters(strTerm)
        
        Call CreateAutoCompleteFileBody(strTerm, strTermPrevious, lngVersion, lngFrequency, blnFirstRow, txt)
        
        If blnFirstRow Then blnFirstRow = False
        strTermPrevious = strTerm
        rst.MoveNext
    Loop
    
    Call CreateAutoCompleteFileFooter(txt)
    
    txt.Close
    Set txt = Nothing
    
    rst.Close
    Set rst = Nothing

    strFullpath = fso.BuildPath("D:\workspace\java\fndds\WebContent\WEB-INF\resources", "foodcodes.xml")
    
    If fso.FileExists(strFullpath) Then
        Call fso.DeleteFile(strFullpath)
    End If
    
    SQL = "SELECT FoodCode AS SuggestDescription, Version, 1 AS Frequency " & _
        "FROM fooddescr " & _
        "ORDER BY SuggestDescription, Version"
    Set rst = New ADODB.Recordset
    Call rst.Open(SQL, cnnBack, adOpenStatic, adLockReadOnly, adCmdText)
    
    Set txt = fso.CreateTextFile(strFullpath, False)
    
    Call CreateAutoCompleteFileHeader(txt)
    
    blnFirstRow = True
    strTermPrevious = vbNullString
    Do Until rst.EOF
        strTerm = rst("SuggestDescription")
        lngVersion = rst("Version")
        lngFrequency = rst("Frequency")
        
        Call CreateAutoCompleteFileBody(strTerm, strTermPrevious, lngVersion, lngFrequency, blnFirstRow, txt)
        
        If blnFirstRow Then blnFirstRow = False
        strTermPrevious = strTerm
        rst.MoveNext
    Loop
    
    Call CreateAutoCompleteFileFooter(txt)
    
    txt.Close
    Set txt = Nothing
    
    rst.Close
    Set rst = Nothing

    strFullpath = fso.BuildPath("D:\workspace\java\fndds\WebContent\WEB-INF\resources", "ingredcodes.xml")
    
    If fso.FileExists(strFullpath) Then
        Call fso.DeleteFile(strFullpath)
    End If
    
    SQL = "SELECT SRCode AS SuggestDescription, Version, COUNT(FoodCode) AS Frequency " & _
        "FROM ingredients " & _
        "GROUP BY SRCode, Version " & _
        "ORDER BY SuggestDescription, Version"
    Set rst = New ADODB.Recordset
    Call rst.Open(SQL, cnnBack, adOpenStatic, adLockReadOnly, adCmdText)
    
    Set txt = fso.CreateTextFile(strFullpath, False)
    
    Call CreateAutoCompleteFileHeader(txt)
    
    blnFirstRow = True
    strTermPrevious = vbNullString
    Do Until rst.EOF
        strTerm = rst("SuggestDescription")
        lngVersion = rst("Version")
        lngFrequency = rst("Frequency")
        
        Call CreateAutoCompleteFileBody(strTerm, strTermPrevious, lngVersion, lngFrequency, blnFirstRow, txt)
        
        If blnFirstRow Then blnFirstRow = False
        strTermPrevious = strTerm
        rst.MoveNext
    Loop
    
    Call CreateAutoCompleteFileFooter(txt)
    
    txt.Close
    Set txt = Nothing
    
    rst.Close
    Set rst = Nothing

End Sub

Private Sub CreateAutoCompleteFileHeader(TextStream As Scripting.TextStream)

    Call TextStream.WriteLine("<?xml version=""1.0"" encoding=""UTF-8""?>")
    Call TextStream.WriteLine("<java version=""1.6.0_17"" class=""java.beans.XMLDecoder"">")
    Call TextStream.WriteLine("   <object class=""java.util.ArrayList"">")

End Sub

Private Sub CreateAutoCompleteFileFooter(TextStream As Scripting.TextStream)

    Call TextStream.WriteLine("         </object>")
    Call TextStream.WriteLine("      </void>")
    Call TextStream.WriteLine("   </object>")
    Call TextStream.WriteLine("</java>")

End Sub

Private Sub CreateAutoCompleteFileBody(Term As String, TermPrevious As String, Version As Long, Frequency As Long, FirstRow As Boolean, TextStream As Scripting.TextStream)

    If StrComp(Term, TermPrevious, vbTextCompare) = 0 Then
        Call TextStream.WriteLine("            <void method=""updateFrequencies"">")
        Call TextStream.WriteLine("               <int>" & CStr(Version) & "</int>")
        Call TextStream.WriteLine("               <int>" & CStr(Frequency) & "</int>")
        Call TextStream.WriteLine("            </void>")
    Else
        If FirstRow Then
            Call TextStream.WriteLine("      <void method=""add"">")
            Call TextStream.WriteLine("         <object class=""com.foodandnutrientdata.fndds.gui.foods.search.autocomplete.FoodWord"">")
        Else
            Call TextStream.WriteLine("         </object>")
            Call TextStream.WriteLine("      </void>")
            Call TextStream.WriteLine("      <void method=""add"">")
            Call TextStream.WriteLine("         <object class=""com.foodandnutrientdata.fndds.gui.foods.search.autocomplete.FoodWord"">")
        End If
        TermPrevious = Term
        Call TextStream.WriteLine("            <void property=""term"">")
        Call TextStream.WriteLine("               <string>" & EncodedXMLString(Term) & "</string>")
        Call TextStream.WriteLine("            </void>")
        Call TextStream.WriteLine("            <void method=""updateFrequencies"">")
        Call TextStream.WriteLine("               <int>" & CStr(Version) & "</int>")
        Call TextStream.WriteLine("               <int>" & CStr(Frequency) & "</int>")
        Call TextStream.WriteLine("            </void>")
    End If

End Sub

Private Sub CreateTables()
    
    Dim lng As Long
    Dim SQL As String

    '--Create food description table
    SQL = "CREATE TABLE fooddescr" & _
        "(" & _
        "FoodCode               INT," & _
        "ModCode                INT DEFAULT 0," & _
        "Version                INT," & _
        "MainDescription        VARCHAR(240)," & _
        "AbbrDescription        VARCHAR(60)," & _
        "IncludesDescription    VARCHAR(2048)," & _
        "FortificationCode      INT DEFAULT 0," & _
        "MoistureChange         DECIMAL(6,3)," & _
        "FatChange              DECIMAL(6,3)," & _
        "FatCode                VARCHAR(8)," & _
        "FatDescription         VARCHAR(200)," & _
        "WeightInitial          DECIMAL(8,3)," & _
        "WeightChange           DECIMAL(8,3)," & _
        "WeightFinal            DECIMAL(8,3)," & _
        "Created                DATETIME DEFAULT CURRENT_TIMESTAMP," & _
        "CONSTRAINT pk_fooddesrc PRIMARY KEY (FoodCode, ModCode, Version)" & _
        ")"
    cnnBack.Execute SQL, lng, adCmdText
    
    '--Create food search table
    SQL = "CREATE TABLE foodsearch" & _
        "(" & _
        "FoodCode           INT," & _
        "ModCode            INT DEFAULT 0," & _
        "SeqNum             INT," & _
        "Version            INT," & _
        "FoodDescription    VARCHAR(240)," & _
        "Created            DATETIME DEFAULT CURRENT_TIMESTAMP," & _
        "CONSTRAINT pk_foodsearch PRIMARY KEY (FoodCode, ModCode, SeqNum, Version)" & _
        ")"
    cnnBack.Execute SQL, lng, adCmdText
    
    '--Add index(s) to food search table
    SQL = "CREATE INDEX indFoodDescription " & _
        "ON foodsearch (FoodDescription)"
    cnnBack.Execute SQL, lng, adCmdText
    
    '--Add constraint(s) to food search
    SQL = "ALTER TABLE foodsearch " & _
        "ADD CONSTRAINT FK_foodsearch_fooddescr " & _
        "FOREIGN KEY (FoodCode, ModCode, Version) " & _
        "REFERENCES fooddescr (FoodCode, ModCode, Version)"
    cnnBack.Execute SQL, lng, adCmdText

    '--Create portions table
    SQL = "CREATE TABLE portions" & _
        "(" & _
        "FoodCode           INT," & _
        "ModCode            INT DEFAULT 0," & _
        "Subcode            INT," & _
        "SubcodeDescr       VARCHAR(80)," & _
        "SeqNum             INT," & _
        "Version            INT," & _
        "PortionCode        INT," & _
        "PortionDescr       VARCHAR(120)," & _
        "PortionChangeType  VARCHAR(1)," & _
        "Weight             DECIMAL(8,3)," & _
        "WeightChangeType   VARCHAR(1)," & _
        "Created            DATETIME DEFAULT CURRENT_TIMESTAMP," & _
        "CONSTRAINT pk_portions PRIMARY KEY (FoodCode, ModCode, Subcode, SeqNum, Version)" & _
        ")"
    cnnBack.Execute SQL, lng, adCmdText

    '--Add constraint(s) to portions
    SQL = "ALTER TABLE portions " & _
        "ADD CONSTRAINT FK_portions_fooddescr " & _
        "FOREIGN KEY (FoodCode, ModCode, Version) " & _
        "REFERENCES fooddescr (FoodCode, ModCode, Version)"
    cnnBack.Execute SQL, lng, adCmdText
    
    '--Create tagname table (INFOODS tagnames)
    SQL = "CREATE TABLE tagname" & _
        "(" & _
        "Tagname            VARCHAR(15)," & _
        "TagnameDescription VARCHAR(255)," & _
        "Units              VARCHAR(255)," & _
        "Tables             VARCHAR(1280)," & _
        "Synonyms           VARCHAR(255)," & _
        "Keywords           VARCHAR(512)," & _
        "Examples           VARCHAR(1536)," & _
        "Comments           VARCHAR(768)," & _
        "Notes              VARCHAR(512)," & _
        "Created            DATETIME DEFAULT CURRENT_TIMESTAMP," & _
        "CONSTRAINT pk_tagname PRIMARY KEY (Tagname)" & _
        ")"
    cnnBack.Execute SQL, lng, adCmdText

    '--Create nutrient description table
    SQL = "CREATE TABLE nutrientdescr" & _
        "(" & _
        "Tagname                VARCHAR(15)," & _
        "Version                INT," & _
        "NutrientDescription    VARCHAR(45)," & _
        "Unit                   VARCHAR(10)," & _
        "Decimals               INT," & _
        "DisplayOrder           INT," & _
        "Created                DATETIME DEFAULT CURRENT_TIMESTAMP," & _
        "CONSTRAINT pk_nutrientdescr PRIMARY KEY (Tagname, Version)" & _
        ")"
    cnnBack.Execute SQL, lng, adCmdText

    '--Add constraint(s) to nutrientdescr
    SQL = "ALTER TABLE nutrientdescr " & _
        "ADD CONSTRAINT FK_nutrientdescr_tagname " & _
        "FOREIGN KEY (Tagname) " & _
        "REFERENCES tagname (Tagname)"
    cnnBack.Execute SQL, lng, adCmdText
    
    '--Create nutrients table
    SQL = "CREATE TABLE nutrients" & _
        "(" & _
        "FoodCode       INT," & _
        "ModCode        INT," & _
        "Tagname        VARCHAR(15)," & _
        "Version        INT," & _
        "NutrientValue  DECIMAL(10,3)," & _
        "Created        DATETIME DEFAULT CURRENT_TIMESTAMP," & _
        "CONSTRAINT pk_nutrients PRIMARY KEY (FoodCode, ModCode, Tagname, Version)" & _
        ")"
    cnnBack.Execute SQL, lng, adCmdText

    '--Add constraint(s) to nutrients
    SQL = "ALTER TABLE nutrients " & _
        "ADD CONSTRAINT FK_nutrients_nutrientdescr " & _
        "FOREIGN KEY (Tagname, Version) " & _
        "REFERENCES nutrientdescr (Tagname, Version)"
    cnnBack.Execute SQL, lng, adCmdText
    SQL = "ALTER TABLE nutrients " & _
        "ADD CONSTRAINT FK_nutrients_fooddescr " & _
        "FOREIGN KEY (FoodCode, ModCode, Version) " & _
        "REFERENCES fooddescr (FoodCode, ModCode, Version)"
    cnnBack.Execute SQL, lng, adCmdText

    '--Create equivalents description table
    SQL = "CREATE TABLE equivalentdescr" & _
        "(" & _
        "Tagname                VARCHAR(25)," & _
        "Version                INT," & _
        "EquivalentDescription  VARCHAR(512)," & _
        "Unit                   VARCHAR(40)," & _
        "Decimals               INT," & _
        "DisplayOrder           INT," & _
        "Created                DATETIME DEFAULT CURRENT_TIMESTAMP," & _
        "CONSTRAINT pk_equivalentdescr PRIMARY KEY (Tagname, Version)" & _
        ")"
    cnnBack.Execute SQL, lng, adCmdText

    '--Create equivalents table
    SQL = "CREATE TABLE equivalents" & _
        "(" & _
        "FoodCode           INT," & _
        "ModCode            INT," & _
        "Tagname            VARCHAR(25)," & _
        "Version            INT," & _
        "EquivalentValue    DECIMAL(10,3)," & _
        "Created            DATETIME DEFAULT CURRENT_TIMESTAMP," & _
        "CONSTRAINT pk_equivalents PRIMARY KEY (FoodCode, ModCode, Tagname, Version)" & _
        ")"
    cnnBack.Execute SQL, lng, adCmdText
    
    '--Add constraint(s) to equivalents
    SQL = "ALTER TABLE equivalents " & _
        "ADD CONSTRAINT FK_equivalents_equivalentdescr " & _
        "FOREIGN KEY (Tagname, Version) " & _
        "REFERENCES equivalentdescr (Tagname, Version)"
    cnnBack.Execute SQL, lng, adCmdText
    SQL = "ALTER TABLE equivalents " & _
        "ADD CONSTRAINT FK_equivalents_fooddescr " & _
        "FOREIGN KEY (FoodCode, ModCode, Version) " & _
        "REFERENCES fooddescr (FoodCode, ModCode, Version)"
    cnnBack.Execute SQL, lng, adCmdText
    
    '--Create ingredients table
    SQL = "CREATE TABLE ingredients" & _
        "(" & _
        "FoodCode               INT," & _
        "ModCode                INT DEFAULT 0," & _
        "SeqNum                 INT," & _
        "Version                INT," & _
        "SRCode                 VARCHAR(8)," & _
        "SRDescr                VARCHAR(240)," & _
        "SRDescrAlt             VARCHAR(200)," & _
        "ChangeTypeToSRCode     VARCHAR(1)," & _
        "IngredType             INT," & _
        "Amount                 DECIMAL(11,3)," & _
        "Measure                VARCHAR(15)," & _
        "PortionCode            INT," & _
        "PortionDescr           VARCHAR(120)," & _
        "RetentionCode          VARCHAR(4)," & _
        "RetentionDescr         VARCHAR(35)," & _
        "ChangeTypeToRetnCode   VARCHAR(1)," & _
        "Flag                   INT," & _
        "Weight                 DECIMAL(11,3)," & _
        "ChangeTypeToWeight     VARCHAR(1)," & _
        "Percentage             DECIMAL(12,8)," & _
        "Created                DATETIME DEFAULT CURRENT_TIMESTAMP," & _
        "CONSTRAINT pk_ingredients PRIMARY KEY (FoodCode, ModCode, SeqNum, Version)" & _
        ")"
    cnnBack.Execute SQL, lng, adCmdText

    '--Add constraint(s) to ingredients
    SQL = "ALTER TABLE ingredients " & _
        "ADD CONSTRAINT FK_ingredients_fooddescr " & _
        "FOREIGN KEY (FoodCode, ModCode, Version) " & _
        "REFERENCES fooddescr (FoodCode, ModCode, Version)"
    cnnBack.Execute SQL, lng, adCmdText
    
    '--Create ingredsearch table
    SQL = "CREATE TABLE ingredsearch" & _
        "(" & _
        "FoodCode           INT," & _
        "ModCode            INT," & _
        "SeqNum             INT," & _
        "IngredType         INT," & _
        "IngrCode           VARCHAR(8)," & _
        "IngrDescr          VARCHAR(240)," & _
        "IngrDescrAlt       VARCHAR(200)," & _
        "Version            INT," & _
        "Created            DATETIME DEFAULT CURRENT_TIMESTAMP," & _
        "CONSTRAINT pk_includesearch PRIMARY KEY (FoodCode, ModCode, SeqNum, IngredType, IngrCode, Version)" & _
        ")"
    cnnBack.Execute SQL, lng, adCmdText

    '--Add constraint(s) to ingredsearch
    SQL = "ALTER TABLE ingredsearch " & _
        "ADD CONSTRAINT FK_ingredsearch_fooddescr " & _
        "FOREIGN KEY (FoodCode, ModCode, Version) " & _
        "REFERENCES fooddescr (FoodCode, ModCode, Version)"
    cnnBack.Execute SQL, lng, adCmdText

    '--Create word table
    SQL = "CREATE TABLE word" & _
        "(" & _
        "WordID             INT CONSTRAINT pk_word PRIMARY KEY," & _
        "WordDescription    VARCHAR(50)," & _
        "Created            DATETIME DEFAULT CURRENT_TIMESTAMP" & _
        ")"
    cnnBack.Execute SQL, lng, adCmdText
    
    '--Create word document
    SQL = "CREATE TABLE worddocument" & _
        "(" & _
        "WordID             INT," & _
        "WordType           INT," & _
        "Version            INT," & _
        "DocumentCount      INT," & _
        "Created            DATETIME DEFAULT CURRENT_TIMESTAMP," & _
        "CONSTRAINT pk_worddocument PRIMARY KEY (WordID, WordType, Version)" & _
        ")"
    cnnBack.Execute SQL, lng, adCmdText
    
    '--Add constraint(s) to word document
    SQL = "ALTER TABLE worddocument " & _
        "ADD CONSTRAINT FK_worddocument_word " & _
        "FOREIGN KEY (WordID) " & _
        "REFERENCES word (WordID)"
    cnnBack.Execute SQL, lng, adCmdText

    '--Create food word table
    SQL = "CREATE TABLE foodword" & _
        "(" & _
        "FoodCode           INT," & _
        "ModCode            INT DEFAULT 0," & _
        "Version            INT," & _
        "WordID             INT," & _
        "WordType           INT," & _
        "WordCount          INT DEFAULT 1," & _
        "tf_idf             DECIMAL(18,16)," & _
        "Created        DATETIME DEFAULT CURRENT_TIMESTAMP," & _
        "CONSTRAINT pk_foodword PRIMARY KEY (FoodCode, ModCode, Version, WordID, WordType)" & _
        ")"
    cnnBack.Execute SQL, lng, adCmdText
    
    '--Add constraint(s) to foodword
    SQL = "ALTER TABLE foodword " & _
        "ADD CONSTRAINT FK_foodword_word " & _
        "FOREIGN KEY (WordID) " & _
        "REFERENCES word (WordID)"
    cnnBack.Execute SQL, lng, adCmdText
    SQL = "ALTER TABLE foodword " & _
        "ADD CONSTRAINT FK_foodword_fooddescr " & _
        "FOREIGN KEY (FoodCode, ModCode, Version) " & _
        "REFERENCES fooddescr (FoodCode, ModCode, Version)"
    cnnBack.Execute SQL, lng, adCmdText

    '--Create similarity table
    SQL = "CREATE TABLE similarity" & _
        "(" & _
        "FoodCodeA          INT," & _
        "ModCodeA           INT DEFAULT 0," & _
        "FoodCodeB          INT," & _
        "ModCodeB           INT DEFAULT 0," & _
        "Version            INT," & _
        "TypeID             INT," & _
        "Similarity         DECIMAL(18,16)," & _
        "Created            DATETIME DEFAULT CURRENT_TIMESTAMP," & _
        "CONSTRAINT pk_similarity PRIMARY KEY (FoodCodeA, ModCodeA, FoodCodeB, ModCodeB, Version, TypeID)" & _
        ")"
    cnnBack.Execute SQL, lng, adCmdText
    
    '--Add constraint(s) to food search
    SQL = "ALTER TABLE similarity " & _
        "ADD CONSTRAINT FK_similarity_fooddescr " & _
        "FOREIGN KEY (FoodCodeA, ModCodeA, Version) " & _
        "REFERENCES fooddescr (FoodCode, ModCode, Version)"
    cnnBack.Execute SQL, lng, adCmdText

    '--Create suggest table
    SQL = "CREATE TABLE suggest" & _
        "(" & _
        "SuggestID          INT," & _
        "SuggestType        INT," & _
        "SuggestDescription VARCHAR(200)," & _
        "Created            DATETIME DEFAULT CURRENT_TIMESTAMP," & _
        "CONSTRAINT pk_suggest PRIMARY KEY (SuggestID, SuggestType)" & _
        ")"
    cnnBack.Execute SQL, lng, adCmdText
    
    '--Create food suggest table
    SQL = "CREATE TABLE foodsuggest" & _
        "(" & _
        "FoodCode           INT," & _
        "ModCode            INT DEFAULT 0," & _
        "Version            INT," & _
        "SuggestID          INT," & _
        "SuggestType        INT," & _
        "SuggestCount       INT DEFAULT 1," & _
        "Created            DATETIME DEFAULT CURRENT_TIMESTAMP," & _
        "CONSTRAINT pk_foodsuggest PRIMARY KEY (FoodCode, ModCode, Version, SuggestID, SuggestType)" & _
        ")"
    cnnBack.Execute SQL, lng, adCmdText

    '--Add constraint(s) to foodsuggest
    SQL = "ALTER TABLE foodsuggest " & _
        "ADD CONSTRAINT FK_foodsuggest_fooddescr " & _
        "FOREIGN KEY (FoodCode, ModCode, Version) " & _
        "REFERENCES fooddescr (FoodCode, ModCode, Version)"
    cnnBack.Execute SQL, lng, adCmdText
    SQL = "ALTER TABLE foodsuggest " & _
        "ADD CONSTRAINT FK_foodsuggest_suggest " & _
        "FOREIGN KEY (SuggestID, SuggestType) " & _
        "REFERENCES suggest (SuggestID, SuggestType)"
    cnnBack.Execute SQL, lng, adCmdText

    '--Create ingredsuggest table
    SQL = "CREATE TABLE ingredsuggest" & _
        "(" & _
        "SRCode             VARCHAR(8)," & _
        "Version            INT," & _
        "SuggestID          INT," & _
        "SuggestType        INT," & _
        "SuggestCount       INT DEFAULT 1," & _
        "Created            DATETIME DEFAULT CURRENT_TIMESTAMP," & _
        "CONSTRAINT pk_ingredsuggest PRIMARY KEY (SRCode, Version, SuggestID, SuggestType)" & _
        ")"
    cnnBack.Execute SQL, lng, adCmdText

    '--Add constraint(s) to ingredsuggest
    SQL = "ALTER TABLE ingredsuggest " & _
        "ADD CONSTRAINT FK_ingredsuggest_suggest " & _
        "FOREIGN KEY (SuggestID, SuggestType) " & _
        "REFERENCES suggest (SuggestID, SuggestType)"
    cnnBack.Execute SQL, lng, adCmdText

End Sub

Private Function DescriptionWithoutSpecialCharacters(Text As String) As String

    Dim j As Long
    Dim lngCharCode As Long
    Dim lngTextLen As Long
    Dim strChar As String
    Dim strText As String
    
    lngTextLen = Len(Text)
            
    strText = vbNullString
    For j = 1 To lngTextLen
        strChar = Mid$(Text, j, 1)
        lngCharCode = Asc(strChar)
        Select Case lngCharCode
            Case 10 'Line feed
            Case 13 'Carriage return
            Case 32 To 34 'Space!"
            Case 36 To 41 '$%&'()
            Case 43 To 47 '+,-./
            Case 48 To 57 'Numbers
            Case 61 '=
            Case 65 To 90 'Uppercase letters
            Case 92 '\
            Case 97 To 122 'Lowercase letters
            Case 146 '’
                Debug.Print lngCharCode, strChar
                strChar = "'"
            Case 232 'è
                Debug.Print lngCharCode, strChar
                strChar = "e"
            Case 233 'é
                Debug.Print lngCharCode, strChar
                strChar = "e"
            Case Else
                Debug.Print lngCharCode, strChar
                Stop
        End Select
        
        strText = strText & strChar
    Next j
    
    'Debug.Print Text, strText
    
    DescriptionWithoutSpecialCharacters = strText

End Function

Private Function DocumentCount(Version As Long) As Long

    With comDocumentCount_Lkp
        .Parameters("@Version") = Version
    End With
    With rstDocumentCount_Lkp
        .Requery
        If .RecordCount > 0 Then
            DocumentCount = CLng(.Fields("DocumentCount"))
        End If
    End With

End Function

Private Sub DropConstraints()

On Error GoTo Err_Handler

    Dim lng As Long

    With cnnBack
        .Execute "ALTER TABLE equivalents DROP CONSTRAINT FK_equivalents_equivalentdescr", lng, adCmdText
        .Execute "ALTER TABLE equivalents DROP CONSTRAINT FK_equivalents_fooddescr", lng, adCmdText
        .Execute "ALTER TABLE foodsearch DROP CONSTRAINT FK_foodsearch_fooddescr", lng, adCmdText
        .Execute "ALTER TABLE foodsuggest DROP CONSTRAINT FK_foodsuggest_fooddescr", lng, adCmdText
        .Execute "ALTER TABLE foodsuggest DROP CONSTRAINT FK_foodsuggest_suggest", lng, adCmdText
        .Execute "ALTER TABLE foodword DROP CONSTRAINT FK_foodword_fooddescr", lng, adCmdText
        .Execute "ALTER TABLE foodword DROP CONSTRAINT FK_foodword_word", lng, adCmdText
        .Execute "ALTER TABLE ingredients DROP CONSTRAINT FK_ingredients_fooddescr", lng, adCmdText
        .Execute "ALTER TABLE ingredsearch DROP CONSTRAINT FK_ingredsearch_fooddescr", lng, adCmdText
        .Execute "ALTER TABLE ingredsuggest DROP CONSTRAINT FK_ingredsuggest_suggest", lng, adCmdText
        .Execute "ALTER TABLE nutrientdescr DROP CONSTRAINT FK_nutrientdescr_tagname", lng, adCmdText
        .Execute "ALTER TABLE nutrients DROP CONSTRAINT FK_nutrients_fooddescr", lng, adCmdText
        .Execute "ALTER TABLE nutrients DROP CONSTRAINT FK_nutrients_nutrientdescr", lng, adCmdText
        .Execute "ALTER TABLE portions DROP CONSTRAINT FK_portions_fooddescr", lng, adCmdText
        .Execute "ALTER TABLE similarity DROP CONSTRAINT FK_similarity_fooddescr", lng, adCmdText
        .Execute "ALTER TABLE worddocument DROP CONSTRAINT FK_worddocument_word", lng, adCmdText
    End With
    
Exit_Sub:
    Exit Sub
Err_Handler:
    If Err.Number = -2147217865 Or Err.Number = -2147217900 Or Err.Number = -2147467259 Then
        Resume Next
    Else
        MsgBox "Description=" & Err.Description & vbCrLf & vbCrLf & _
            "Number=" & Err.Number, vbCritical, "Untrapped Error in DropConstraints Sub"
        GoTo Exit_Sub
    End If

End Sub

Private Sub DropTables()

On Error GoTo Err_Handler

    Dim lng As Long

    With cnnBack
        .Execute "DROP TABLE equivalentdescr", lng, adCmdText
        .Execute "DROP TABLE equivalents", lng, adCmdText
        .Execute "DROP TABLE fooddescr", lng, adCmdText
        .Execute "DROP TABLE foodsearch", lng, adCmdText
        .Execute "DROP TABLE foodsuggest", lng, adCmdText
        .Execute "DROP TABLE foodword", lng, adCmdText
        .Execute "DROP TABLE ingredients", lng, adCmdText
        .Execute "DROP TABLE ingredsearch", lng, adCmdText
        .Execute "DROP TABLE ingredsuggest", lng, adCmdText
        .Execute "DROP TABLE nutrientdescr", lng, adCmdText
        .Execute "DROP TABLE nutrients", lng, adCmdText
        .Execute "DROP TABLE portions", lng, adCmdText
        .Execute "DROP TABLE similarity", lng, adCmdText
        .Execute "DROP TABLE suggest", lng, adCmdText
        .Execute "DROP TABLE tagname", lng, adCmdText
        .Execute "DROP TABLE word", lng, adCmdText
        .Execute "DROP TABLE worddocument", lng, adCmdText
    End With
    
Exit_Sub:
    Exit Sub
Err_Handler:
    If Err.Number = -2147217865 Then
        Resume Next
    Else
        MsgBox "Description=" & Err.Description & vbCrLf & vbCrLf & _
            "Number=" & Err.Number, vbCritical, "Untrapped Error in DropTables Sub"
        GoTo Exit_Sub
    End If
            
End Sub

Private Function EncodedXMLString(Text As String) As String

    Dim strText As String
    
    strText = Replace(Text, "&", "&amp;", , , vbTextCompare)
    strText = Replace(strText, "<", "&lt;", , , vbTextCompare)
    strText = Replace(strText, ">", "&gt;", , , vbTextCompare)
    strText = Replace(strText, "'", "&apos;", , , vbTextCompare)
    strText = Replace(strText, """", "&quot;", , , vbTextCompare)
    
    EncodedXMLString = strText

End Function

Private Function EquivalentDescription(Tagname As String) As String

    Select Case Tagname
'        Case "EQUIVFLAG"
'            EquivalentDescription = "Equivalent Flag"
        Case "F_TOTAL"
            EquivalentDescription = "Total intact fruits (whole or cut) and fruit juices (cup eq.)"
        Case "F_CITMLB"
            EquivalentDescription = "Intact fruits (whole or cut) of citrus, melons, and berries (cup eq.)"
        Case "F_OTHER"
            EquivalentDescription = "Intact fruits (whole or cut); excluding citrus, melons, and berries (cup eq.)"
        Case "F_JUICE"
            EquivalentDescription = "Fruit juices, citrus and non citrus (cup eq.)"
        Case "V_TOTAL"
            EquivalentDescription = "Total dark green, red and orange, starchy, and other vegetables; excludes legumes (cup eq.)"
        Case "V_DRKGR"
            EquivalentDescription = "Dark green vegetables (cup eq.)"
        Case "V_REDOR_TOTAL"
            EquivalentDescription = "Total red and orange vegetables (tomatoes and tomato products + other red and orange vegetables) (cup eq.)"
        Case "V_REDOR_TOMATO"
            EquivalentDescription = "Tomatoes and tomato products (cup eq.)"
        Case "V_REDOR_OTHER"
            EquivalentDescription = "Other red and orange vegetables, excluding tomatoes and tomato products (cup eq.)"
        Case "V_STARCHY_TOTAL"
            EquivalentDescription = "Total starchy vegetables (white potatoes + other starchy vegetables) (cup eq.)"
        Case "V_STARCHY_POTATO"
            EquivalentDescription = "White potatoes (cup eq.)"
        Case "V_STARCHY_OTHER"
            EquivalentDescription = "Other starchy vegetables, excluding white potatoes (cup eq.)"
        Case "V_OTHER"
            EquivalentDescription = "Other vegetables not in the vegetable components listed above (cup eq.)"
        Case "V_LEGUMES"
            EquivalentDescription = "Beans and peas (legumes) computed as vegetables (cup eq.)"
        Case "G_TOTAL"
            EquivalentDescription = "Total whole and refined grains (oz. eq.)"
        Case "G_WHOLE"
            EquivalentDescription = "Grains defined as whole grains and contain the entire grain kernel ? the bran, germ, and endosperm (oz. eq.)"
        Case "G_REFINED"
            EquivalentDescription = "Refined grains that do not contain all of the components of the entire grain kernel (oz. eq.)"
        Case "PF_TOTAL"
            EquivalentDescription = "Total meat, poultry, organ meat, cured meat, seafood, eggs, soy, and nuts and seeds; excludes legumes (oz. eq.)"
        Case "PF_MPS_TOTAL"
            EquivalentDescription = "Total of meat, poultry, seafood, organ meat, and cured meat (oz. eq.)"
        Case "PF_MEAT"
            EquivalentDescription = "Beef, veal, pork, lamb, and game meat; excludes organ meat and cured meat (oz. eq.)"
        Case "PF_CUREDMEAT"
            EquivalentDescription = "Frankfurters, sausages, corned beef, and luncheon meat that are made from beef, pork, or poultry (oz. eq.)"
        Case "PF_ORGAN"
            EquivalentDescription = "Organ meat from beef, veal, pork, lamb, game, and poultry (oz. eq.)"
        Case "PF_POULT"
            EquivalentDescription = "Chicken, turkey, Cornish hens, duck, goose, quail, and pheasant (game birds); excludes organ meat and cured meat (oz. eq.)"
        Case "PF_SEAFD_HI"
            EquivalentDescription = "Seafood (finfish, shellfish, and other seafood) high in n-3 fatty acids (oz. eq.)"
        Case "PF_SEAFD_LOW"
            EquivalentDescription = "Seafood (finfish, shellfish, and other seafood) low in n-3 fatty acids (oz. eq.)"
        Case "PF_EGGS"
            EquivalentDescription = "Eggs (chicken, duck, goose, quail) and egg substitutes (oz. eq.)"
        Case "PF_SOY"
            EquivalentDescription = "Soy products, excluding calcium fortified soy milk and immature soybeans (oz. eq.)"
        Case "PF_NUTSDS"
            EquivalentDescription = "Peanuts, tree nuts, and seeds; excludes coconut (oz. eq.)"
        Case "PF_LEGUMES"
            EquivalentDescription = "Beans and Peas (legumes) computed as protein foods (oz. eq.)"
        Case "D_TOTAL"
            EquivalentDescription = "Total milk, yogurt, cheese, and whey. For some foods, the total dairy values could be higher than the sum of D_MILK, D_YOGURT, and D_CHEESE because Miscellaneous dairy component composed of whey which is not included in FPED as a separate variable. (cup eq.)"
        Case "D_MILK"
            EquivalentDescription = "Fluid milk, buttermilk, evaporated milk, dry milk, and calcium fortified soy milk (cup eq.)"
        Case "D_YOGURT"
            EquivalentDescription = "Yogurt (cup eq.)"
        Case "D_CHEESE"
            EquivalentDescription = "Cheeses (cup eq.)"
        Case "OILS"
            EquivalentDescription = "Fats naturally present in nuts, seeds, and seafood; unhydrogentated vegetable oils, except palm oil, palm kernel oil, and coconut oils; fat present in avocado and olives above the allowable amount; 50% of fat present in stick and tub margarines and margarine spreads (grams)"
        Case "SOLID_FATS"
            EquivalentDescription = "Fats naturally present in meat, poultry, eggs, and dairy (lard, tallow, and butter); hydrogenated or partially hydrogenated oils; shortening, palm, palm kernel and coconut oils; fat naturally present in coconut meat and cocoa butter; and 50% of fat present in stick and tub margarines and margarine spreads (grams)"
        Case "ADD_SUGARS"
            EquivalentDescription = "Foods defined as added sugars (tsp. eq.)"
        Case "A_DRINKS"
            EquivalentDescription = "Alcoholic beverages and alcohol (ethanol) added to foods after cooking (no. of drinks)"
        Case Else
            Stop
    End Select
    
End Function

Private Function EquivalentSortOrder(Tagname As String) As Long

    Select Case Tagname
'        Case "EQUIVFLAG"
'            EquivalentSortOrder = 0
        Case "F_TOTAL"
            EquivalentSortOrder = 100
        Case "F_CITMLB"
            EquivalentSortOrder = 110
        Case "F_OTHER"
            EquivalentSortOrder = 120
        Case "F_JUICE"
            EquivalentSortOrder = 130
        Case "V_TOTAL"
            EquivalentSortOrder = 200
        Case "V_DRKGR"
            EquivalentSortOrder = 210
        Case "V_REDOR_TOTAL"
            EquivalentSortOrder = 220
        Case "V_REDOR_TOMATO"
            EquivalentSortOrder = 230
        Case "V_REDOR_OTHER"
            EquivalentSortOrder = 235
        Case "V_STARCHY_TOTAL"
            EquivalentSortOrder = 240
        Case "V_STARCHY_POTATO"
            EquivalentSortOrder = 250
        Case "V_STARCHY_OTHER"
            EquivalentSortOrder = 255
        Case "V_OTHER"
            EquivalentSortOrder = 260
        Case "V_LEGUMES"
            EquivalentSortOrder = 270
        Case "G_TOTAL"
            EquivalentSortOrder = 300
        Case "G_WHOLE"
            EquivalentSortOrder = 310
        Case "G_REFINED"
            EquivalentSortOrder = 320
        Case "PF_TOTAL"
            EquivalentSortOrder = 400
        Case "PF_MPS_TOTAL"
            EquivalentSortOrder = 410
        Case "PF_MEAT"
            EquivalentSortOrder = 420
        Case "PF_CUREDMEAT"
            EquivalentSortOrder = 425
        Case "PF_ORGAN"
            EquivalentSortOrder = 430
        Case "PF_POULT"
            EquivalentSortOrder = 440
        Case "PF_SEAFD_HI"
            EquivalentSortOrder = 450
        Case "PF_SEAFD_LOW"
            EquivalentSortOrder = 455
        Case "PF_EGGS"
            EquivalentSortOrder = 460
        Case "PF_SOY"
            EquivalentSortOrder = 470
        Case "PF_NUTSDS"
            EquivalentSortOrder = 480
        Case "PF_LEGUMES"
            EquivalentSortOrder = 490
        Case "D_TOTAL"
            EquivalentSortOrder = 500
        Case "D_MILK"
            EquivalentSortOrder = 510
        Case "D_YOGURT"
            EquivalentSortOrder = 520
        Case "D_CHEESE"
            EquivalentSortOrder = 530
        Case "OILS"
            EquivalentSortOrder = 600
        Case "SOLID_FATS"
            EquivalentSortOrder = 610
        Case "ADD_SUGARS"
            EquivalentSortOrder = 700
        Case "A_DRINKS"
            EquivalentSortOrder = 800
        Case Else
            Stop
    End Select
    
End Function

Private Function EquivalentUnits(Tagname As String) As String

    Select Case Tagname
'        Case "EQUIVFLAG"
'            EquivalentUnits = "N/A"
        Case "F_TOTAL", "F_CITMLB", "F_OTHER", "F_JUICE"
            EquivalentUnits = "cup equivalents"
        Case "V_TOTAL", "V_DRKGR", "V_REDOR_TOTAL", "V_REDOR_TOMATO", "V_REDOR_OTHER", "V_STARCHY_TOTAL", "V_STARCHY_POTATO", "V_STARCHY_OTHER", "V_OTHER", "V_LEGUMES"
            EquivalentUnits = "cup equivalents"
        Case "G_TOTAL", "G_WHOLE", "G_REFINED"
            EquivalentUnits = "ounce equivalents"
        Case "PF_TOTAL", "PF_MPS_TOTAL", "PF_MEAT", "PF_CUREDMEAT", "PF_ORGAN", "PF_POULT", "PF_SEAFD_HI", "PF_SEAFD_LOW", "PF_EGGS", "PF_SOY", "PF_NUTSDS", "PF_LEGUMES"
            EquivalentUnits = "ounce equivalents"
        Case "D_TOTAL", "D_MILK", "D_YOGURT", "D_CHEESE"
            EquivalentUnits = "cup equivalents"
        Case "OILS", "SOLID_FATS"
            EquivalentUnits = "grams"
        Case "ADD_SUGARS"
            EquivalentUnits = "teaspoon equivalents"
        Case "A_DRINKS"
            EquivalentUnits = "number of drinks"
        Case Else
            Stop
    End Select
    
End Function

Public Sub ExportData(Version As FNDDSVersionNumber)
    
    '--Export record(s)
    Call ExportFoodDescr(Version)
    Call ExportFoodSearch(Version)
    ' Call AppendFoodSuggest
    ' Call AppendFoodWords
    Call ExportPortions(Version)
    Call ExportTagname(Version)
    Call ExportNutrientDescr(Version)
    Call ExportNutrients(Version)
    ' Call AppendEquivalentDescr
    ' Call AppendEquivalents
    Call ExportIngredients(Version)
    Call ExportIngredSearch(Version)
    'Call AppendIngredSuggest

End Sub

Private Function ExportFolderName(Version As FNDDSVersionNumber) As String

    Select Case Version
        Case fvnFNDDS1
            ExportFolderName = "v1"
        Case fvnFNDDS2
            ExportFolderName = "v2"
        Case fvnFNDDS3
            ExportFolderName = "v3"
        Case fvnFNDDS4
            ExportFolderName = "v4"
        Case fvnFNDDS5
            ExportFolderName = "v5"
        Case fvnFNDDS6
            ExportFolderName = "v6"
        Case fvnFNDDS7
            ExportFolderName = "v7"
        Case Else
            Stop
    End Select

End Function

Private Sub ExportFoodDescr(Version As FNDDSVersionNumber)

    Dim SQL As String
    Dim lngIndex As Long
    Dim strComma As String
    Dim strFieldName As String
    Dim strFileName As String
    Dim strFolderName As String
    Dim strInsert As String
    Dim strValues As String
    Dim fld As ADODB.Field
    Dim rst As ADODB.Recordset
    Dim txt As Scripting.TextStream
    
    SQL = "SELECT FoodCode," & _
        "ModCode," & _
        "Version," & _
        "MainDescription," & _
        "AbbrDescription," & _
        "IncludesDescription," & _
        "FortificationCode," & _
        "MoistureChange," & _
        "FatChange," & _
        "FatCode," & _
        "FatDescription," & _
        "WeightInitial," & _
        "WeightChange," & _
        "WeightFinal " & _
        "FROM fooddescr " & _
        "WHERE (Version = " & Version & ")" & _
        "ORDER BY FoodCode," & _
        "ModCode"
    Set rst = New ADODB.Recordset
    Call rst.Open(SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText)
    
    strFolderName = fso.BuildPath(SQL_PATH, ExportFolderName(Version))
    If fso.FolderExists(strFolderName) = False Then
        Call fso.CreateFolder(strFolderName)
    End If
    strFileName = fso.BuildPath(strFolderName, "Inserts_01_FoodDescr_v" & Version & ".sql")
    Set txt = fso.CreateTextFile(strFileName)
    
    Call txt.WriteLine("-- ----------------------------------------------------------------------")
    Call txt.WriteLine("-- Food Description Table (Version " & Version & ")")
    Call txt.WriteLine("-- ----------------------------------------------------------------------")
    
    strInsert = "INSERT INTO FoodDescr (FoodCode, ModCode, Version, MainDescription, AbbrDescription, " & _
        "IncludesDescription, FortificationCode, MoistureChange, FatChange, FatCode, FatDescription, " & _
        "WeightInitial, WeightChange, WeightFinal)"
    Call txt.WriteLine(strInsert)
    Call txt.Write("VALUES ")
    
    Call ExportRecordset(txt, rst, strInsert)
    
    txt.Close
    Set txt = Nothing
    
    rst.Close
    Set rst = Nothing

End Sub

Private Sub ExportFoodSearch(Version As FNDDSVersionNumber)

    Dim SQL As String
    Dim lngIndex As Long
    Dim strFileName As String
    Dim strFolderName As String
    Dim strInsert As String
    Dim strValues As String
    Dim fld As ADODB.Field
    Dim rst As ADODB.Recordset
    Dim txt As Scripting.TextStream
    
    SQL = "SELECT FoodCode," & _
        "ModCode," & _
        "SeqNum," & _
        "Version," & _
        "FoodDescription " & _
        "FROM foodsearch " & _
        "WHERE (Version = " & Version & ")" & _
        "ORDER BY FoodCode," & _
        "ModCode," & _
        "SeqNum"
    Set rst = New ADODB.Recordset
    Call rst.Open(SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText)
    
    strFolderName = fso.BuildPath(SQL_PATH, ExportFolderName(Version))
    If fso.FolderExists(strFolderName) = False Then
        Call fso.CreateFolder(strFolderName)
    End If
    strFileName = fso.BuildPath(strFolderName, "Inserts_02_FoodSearch_v" & Version & ".sql")
    Set txt = fso.CreateTextFile(strFileName)
    
    Call txt.WriteLine("-- ----------------------------------------------------------------------")
    Call txt.WriteLine("-- Food Search Table (Version " & Version & ")")
    Call txt.WriteLine("-- ----------------------------------------------------------------------")
    
    strInsert = "INSERT INTO FoodSearch (FoodCode, ModCode, SeqNum, Version, FoodDescription)"
    Call txt.WriteLine(strInsert)
    Call txt.Write("VALUES ")
    
    Call ExportRecordset(txt, rst, strInsert)
    
    txt.Close
    Set txt = Nothing
    
    rst.Close
    Set rst = Nothing

End Sub

Private Sub ExportIngredients(Version As FNDDSVersionNumber)

    Dim SQL As String
    Dim lngIndex As Long
    Dim strFileName As String
    Dim strFolderName As String
    Dim strInsert As String
    Dim strValues As String
    Dim fld As ADODB.Field
    Dim rst As ADODB.Recordset
    Dim txt As Scripting.TextStream
    
    SQL = "SELECT FoodCode," & _
        "ModCode," & _
        "SeqNum," & _
        "Version," & _
        "SRCode," & _
        "SRDescr," & _
        "SRDescrAlt," & _
        "ChangeTypeToSRCode," & _
        "IngredType," & _
        "Amount," & _
        "Measure," & _
        "PortionCode," & _
        "PortionDescr," & _
        "RetentionCode," & _
        "RetentionDescr," & _
        "ChangeTypeToRetnCode," & _
        "Flag," & _
        "Weight," & _
        "ChangeTypeToWeight," & _
        "Percentage " & _
        "FROM ingredients " & _
        "WHERE (Version = " & Version & ")" & _
        "ORDER BY FoodCode," & _
        "ModCode," & _
        "SeqNum"
    Set rst = New ADODB.Recordset
    Call rst.Open(SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText)
    
    strFolderName = fso.BuildPath(SQL_PATH, ExportFolderName(Version))
    If fso.FolderExists(strFolderName) = False Then
        Call fso.CreateFolder(strFolderName)
    End If
    strFileName = fso.BuildPath(strFolderName, "Inserts_09_Ingredients_v" & Version & ".sql")
    Set txt = fso.CreateTextFile(strFileName)
    
    Call txt.WriteLine("-- ----------------------------------------------------------------------")
    Call txt.WriteLine("-- Ingredients Table (Version " & Version & ")")
    Call txt.WriteLine("-- ----------------------------------------------------------------------")
    
    strInsert = "INSERT INTO Ingredients (FoodCode, ModCode, SeqNum, Version, SRCode, SRDescr, SRDescrAlt, " & _
        "ChangeTypeToSRCode, IngredType, Amount, Measure, PortionCode, PortionDescr, RetentionCode, " & _
        "RetentionDescr, ChangeTypeToRetnCode, Flag, Weight, ChangeTypeToWeight, Percentage)"
    Call txt.WriteLine(strInsert)
    Call txt.Write("VALUES ")
    
    Call ExportRecordset(txt, rst, strInsert)
    
    txt.Close
    Set txt = Nothing
    
    rst.Close
    Set rst = Nothing

End Sub

Private Sub ExportIngredSearch(Version As FNDDSVersionNumber)

    Dim SQL As String
    Dim lngIndex As Long
    Dim strFileName As String
    Dim strFolderName As String
    Dim strInsert As String
    Dim strValues As String
    Dim fld As ADODB.Field
    Dim rst As ADODB.Recordset
    Dim txt As Scripting.TextStream
    
    SQL = "SELECT FoodCode," & _
        "ModCode," & _
        "SeqNum," & _
        "IngredType," & _
        "IngrCode," & _
        "IngrDescr," & _
        "IngrDescrAlt," & _
        "Version " & _
        "FROM ingredsearch " & _
        "WHERE (Version = " & Version & ")" & _
        "ORDER BY FoodCode," & _
        "ModCode," & _
        "SeqNum"
    Set rst = New ADODB.Recordset
    Call rst.Open(SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText)
    
    strFolderName = fso.BuildPath(SQL_PATH, ExportFolderName(Version))
    If fso.FolderExists(strFolderName) = False Then
        Call fso.CreateFolder(strFolderName)
    End If
    strFileName = fso.BuildPath(strFolderName, "Inserts_10_IngredSearch_v" & Version & ".sql")
    Set txt = fso.CreateTextFile(strFileName)
    
    Call txt.WriteLine("-- ----------------------------------------------------------------------")
    Call txt.WriteLine("-- Ingredients Search Table (Version " & Version & ")")
    Call txt.WriteLine("-- ----------------------------------------------------------------------")
    
    strInsert = "INSERT INTO IngredSearch (FoodCode, ModCode, SeqNum, IngredType, IngrCode, IngrDescr, IngrDescrAlt, Version)"
    Call txt.WriteLine(strInsert)
    Call txt.Write("VALUES ")
    
    Call ExportRecordset(txt, rst, strInsert)
    
    txt.Close
    Set txt = Nothing
    
    rst.Close
    Set rst = Nothing

End Sub

Private Sub ExportNutrientDescr(Version As FNDDSVersionNumber)

    Dim SQL As String
    Dim lngIndex As Long
    Dim strFileName As String
    Dim strFolderName As String
    Dim strInsert As String
    Dim strValues As String
    Dim fld As ADODB.Field
    Dim rst As ADODB.Recordset
    Dim txt As Scripting.TextStream
    
    SQL = "SELECT Tagname, Version, NutrientDescription, Unit, Decimals, DisplayOrder " & _
        "FROM nutrientdescr " & _
        "WHERE (Version = " & Version & ")" & _
        "ORDER BY Tagname"
    Set rst = New ADODB.Recordset
    Call rst.Open(SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText)
    
    strFolderName = fso.BuildPath(SQL_PATH, ExportFolderName(Version))
    If fso.FolderExists(strFolderName) = False Then
        Call fso.CreateFolder(strFolderName)
    End If
    strFileName = fso.BuildPath(strFolderName, "Inserts_05_NutrientDescr_v" & Version & ".sql")
    Set txt = fso.CreateTextFile(strFileName)
    
    Call txt.WriteLine("-- ----------------------------------------------------------------------")
    Call txt.WriteLine("-- Nutrient Description Table (Version " & Version & ")")
    Call txt.WriteLine("-- ----------------------------------------------------------------------")
    
    strInsert = "INSERT INTO NutrientDescr (Tagname, Version, NutrientDescription, Unit, Decimals, DisplayOrder)"
    Call txt.WriteLine(strInsert)
    Call txt.Write("VALUES ")
    
    Call ExportRecordset(txt, rst, strInsert)
    
    txt.Close
    Set txt = Nothing
    
    rst.Close
    Set rst = Nothing

End Sub

Private Sub ExportNutrients(Version As FNDDSVersionNumber)

    Dim SQL As String
    Dim lngIndex As Long
    Dim lngIndexTotal As Long
    Dim strFileName As String
    Dim strFolderName As String
    Dim strHeaderLine As String
    Dim strHeaderText As String
    Dim strInsert As String
    Dim strValues As String
    Dim fld As ADODB.Field
    Dim rst As ADODB.Recordset
    Dim txt As Scripting.TextStream
    
    SQL = "SELECT FoodCode, ModCode, Tagname, Version, NutrientValue " & _
        "FROM nutrients " & _
        "WHERE (Version = " & Version & ")" & _
        "ORDER BY FoodCode," & _
        "ModCode," & _
        "Tagname"
    Set rst = New ADODB.Recordset
    Call rst.Open(SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText)
    
    strFolderName = fso.BuildPath(SQL_PATH, ExportFolderName(Version))
    If fso.FolderExists(strFolderName) = False Then
        Call fso.CreateFolder(strFolderName)
    End If
    strFileName = fso.BuildPath(strFolderName, "Inserts_06_Nutrients_v" & Version & ".sql")
    Set txt = fso.CreateTextFile(strFileName)
    
    strHeaderLine = "-- ----------------------------------------------------------------------"
    strHeaderText = "-- Nutrients Table (Version " & Version & ")"
    
    Call txt.WriteLine(strHeaderLine)
    Call txt.WriteLine(strHeaderText)
    Call txt.WriteLine(strHeaderLine)
    
    strInsert = "INSERT INTO Nutrients (FoodCode, ModCode, Tagname, Version, NutrientValue)"
    
    Call txt.WriteLine(strInsert)
    Call txt.Write("VALUES ")
    
    Call ExportRecordset(txt, rst, strInsert)
    
    If Not (txt Is Nothing) Then
        txt.Close
    End If
    Set txt = Nothing
    
    rst.Close
    Set rst = Nothing

End Sub

Private Sub ExportPortions(Version As FNDDSVersionNumber)

    Dim SQL As String
    Dim lngIndex As Long
    Dim strFileName As String
    Dim strFolderName As String
    Dim strInsert As String
    Dim strValues As String
    Dim fld As ADODB.Field
    Dim rst As ADODB.Recordset
    Dim txt As Scripting.TextStream
    
    SQL = "SELECT FoodCode," & _
        "ModCode," & _
        "Subcode," & _
        "SubcodeDescr," & _
        "SeqNum," & _
        "Version," & _
        "PortionCode," & _
        "PortionDescr," & _
        "PortionChangeType," & _
        "Weight," & _
        "WeightChangeType " & _
        "FROM portions WHERE (Version = " & Version & ")" & _
        "ORDER BY FoodCode," & _
        "ModCode," & _
        "Subcode," & _
        "SeqNum"
    Set rst = New ADODB.Recordset
    Call rst.Open(SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText)
    
    strFolderName = fso.BuildPath(SQL_PATH, ExportFolderName(Version))
    If fso.FolderExists(strFolderName) = False Then
        Call fso.CreateFolder(strFolderName)
    End If
    strFileName = fso.BuildPath(strFolderName, "Inserts_03_Portions_v" & Version & ".sql")
    Set txt = fso.CreateTextFile(strFileName)
    
    Call txt.WriteLine("-- ----------------------------------------------------------------------")
    Call txt.WriteLine("-- Portions Table (Version " & Version & ")")
    Call txt.WriteLine("-- ----------------------------------------------------------------------")
    
    strInsert = "INSERT INTO Portions (FoodCode, ModCode, Subcode, SubcodeDescr, SeqNum, " & _
        "Version, PortionCode, PortionDescr, PortionChangeType, Weight, WeightChangeType)"
    Call txt.WriteLine(strInsert)
    Call txt.Write("VALUES ")
    
    Call ExportRecordset(txt, rst, strInsert)
    
    txt.Close
    Set txt = Nothing
    
    rst.Close
    Set rst = Nothing

End Sub

Private Sub ExportRecordset(TextFile As Scripting.TextStream, Recordset As ADODB.Recordset, InsertSQL As String)

    Dim lngIndex As Long
    Dim strComma As String
    Dim strValues As String
    Dim fld As ADODB.Field

    Do Until Recordset.EOF
        strComma = vbNullString
        strValues = vbNullString
        
        For Each fld In Recordset.Fields
            If Len(strValues) > 0 Then
                strComma = ", "
            End If
            
            Select Case fld.Type
                Case adInteger '3
                    If IsNull(fld.Value) Then
                        strValues = strValues & strComma & "NULL"
                    Else
                        strValues = strValues & strComma & fld.Value
                    End If
                Case adNumeric '131
                    If IsNull(fld.Value) Then
                        strValues = strValues & strComma & "NULL"
                    Else
                        strValues = strValues & strComma & fld.Value
                    End If
                Case adVarChar '200
                    If IsNull(fld.Value) Then
                        strValues = strValues & strComma & "NULL"
                    Else
                        strValues = strValues & strComma & "'" & Utility.EscapedString(fld.Value) & "'"
                    End If
                Case Else
                    Debug.Print fld.Type
                    Stop
            End Select
        Next fld
        
        Recordset.MoveNext
    
        '-- Add the tab
        If lngIndex > 0 Then
            Call TextFile.Write("   ")
        End If

        If lngIndex > 999 Then
            Call TextFile.WriteLine("(" & strValues & ");")
            Call TextFile.WriteLine(InsertSQL)
            Call TextFile.Write("VALUES ")
            lngIndex = 0
        Else
            If Recordset.EOF Then
                Call TextFile.WriteLine("(" & strValues & ");")
            Else
                Call TextFile.WriteLine("(" & strValues & "),")
            End If
        
            lngIndex = lngIndex + 1
        End If
    Loop

End Sub

Private Sub ExportTagname(Version As FNDDSVersionNumber)

    Dim SQL As String
    Dim lngIndex As Long
    Dim strFileName As String
    Dim strFolderName As String
    Dim strInsert As String
    Dim strValues As String
    Dim fld As ADODB.Field
    Dim rst As ADODB.Recordset
    Dim txt As Scripting.TextStream
    
    SQL = "SELECT Tagname," & _
        "TagnameDescription," & _
        "Units," & _
        "Tables," & _
        "Synonyms," & _
        "Keywords," & _
        "Examples," & _
        "Comments," & _
        "Notes " & _
        "FROM tagname " & _
        "ORDER BY Tagname"
    Set rst = New ADODB.Recordset
    Call rst.Open(SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText)
    
    strFolderName = fso.BuildPath(SQL_PATH, ExportFolderName(Version))
    If fso.FolderExists(strFolderName) = False Then
        Call fso.CreateFolder(strFolderName)
    End If
    strFileName = fso.BuildPath(strFolderName, "Inserts_04_Tagname_v" & Version & ".sql")
    Set txt = fso.CreateTextFile(strFileName)
    
    Call txt.WriteLine("-- ----------------------------------------------------------------------")
    Call txt.WriteLine("-- Tagname Table")
    Call txt.WriteLine("-- ----------------------------------------------------------------------")
    
    strInsert = "INSERT INTO Tagname (Tagname, TagnameDescription, Units, Tables, Synonyms, Keywords, Examples, Comments, Notes)"
    Call txt.WriteLine(strInsert)
    Call txt.Write("VALUES ")
    
    Call ExportRecordset(txt, rst, strInsert)
    
    txt.Close
    Set txt = Nothing
    
    rst.Close
    Set rst = Nothing

End Sub

Private Function FoodMatrixMagnitude(Recordset As ADODB.Recordset) As String

    Dim dblMagnitude As Double
    Dim dblTF_IDF As Double
    
    With Recordset
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                dblTF_IDF = CDbl(.Fields("tf_idf"))
                dblMagnitude = dblMagnitude + (dblTF_IDF ^ 2)
                .MoveNext
            Loop
            FoodMatrixMagnitude = Sqr(dblMagnitude)
        End If
    End With

End Function

Private Function FoodMatrixValue(FoodCode As Long, ModCode As Long, Version As Long, WordID As Long, WordType As Long) As Double

    With comFoodMatrixValue_Lkp
        .Parameters("@FoodCode") = FoodCode
        .Parameters("@ModCode") = ModCode
        .Parameters("@Version") = Version
        .Parameters("@WordID") = WordID
        .Parameters("@WordType1") = 1
        .Parameters("@WordType2") = WordType
    End With
    With rstFoodMatrixValue_Lkp
        .Requery
        If .RecordCount > 0 Then
            If Not IsNull(.Fields("tf_idf")) Then
                FoodMatrixValue = CDbl(.Fields("tf_idf"))
            End If
        End If
    End With

End Function

Private Function FoodMatrixWordIDs(Recordset As ADODB.Recordset) As String

    Dim str As String
    
    With Recordset
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                If Len(str) > 0 Then
                    str = str & ","
                End If
                str = str & .Fields("WordID")
                .MoveNext
            Loop
            FoodMatrixWordIDs = str
        End If
    End With

End Function

Private Function FormattedSuggestDescr(Description As String, Optional Includes As Boolean = False) As String

    Dim l As Long
    Dim lngInStrRev As Long
    Dim lngLen As Long
    Dim strDescription As String
    Dim strBrandNames1() As String
    Dim strBrandNames2() As String
    Dim strFoodWords1() As String
    Dim strFoodWords2() As String
    
    strDescription = Trim$(Description)
    
    If Includes Then
'        Debug.Print strDescription
        strDescription = Replace(strDescription, "; or ", ", ", , , vbTextCompare)
'        Debug.Print strDescription
    End If
    
    '--Replace "(" and ")" and ":" and ";" with a comma
    strDescription = Replace(strDescription, "(", ",", , , vbTextCompare)
    strDescription = Replace(strDescription, ")", ",", , , vbTextCompare)
    strDescription = Replace(strDescription, ":", ",", , , vbTextCompare)
    strDescription = Replace(strDescription, ";", ",", , , vbTextCompare)
    
    '--Remove "..." and ".." and ". . ."
    strDescription = Replace(strDescription, "...", "", , , vbTextCompare)
    strDescription = Replace(strDescription, "..", "", , , vbTextCompare)
    strDescription = Replace(strDescription, ". . .", "", , , vbTextCompare)
    
    '--Add a comma in front of " prepared with "
'    strDescription = Replace(strDescription, " prepared with ", ",made with ", , , vbTextCompare)
    
    '--Add a comma in front of " made from "
'    strDescription = Replace(strDescription, " made from ", ",made from ", , , vbTextCompare)
    
    '--Add a comma in front of " made with "
'    strDescription = Replace(strDescription, " made with ", ",made with ", , , vbTextCompare)
    
    '--Add a comma in front of " w/o "
'    strDescription = Replace(strDescription, " w/o ", ",w/o ", , , vbTextCompare)
    
    '--Add a comma in front of " w/0 "
'    strDescription = Replace(strDescription, " w/0 ", ",w/0 ", , , vbTextCompare)
    
    If InStr(1, strDescription, " made w/ ", vbTextCompare) > 0 Then
        '--Add a comma in front of " made w/ "
'        strDescription = Replace(strDescription, " made w/ ", ",made w/ ", , , vbTextCompare)
    Else
        '--Add a comma in front of " w/ "
        strDescription = Replace(strDescription, " w/ ", ",w/ ", , , vbTextCompare)
    End If
    
    '--Add a comma after brand names
    ReDim strBrandNames1(2)
    strBrandNames1(0) = "Arby's "
    strBrandNames1(1) = "Budget Gourmet "
    strBrandNames1(2) = "Campbell's "
    
    ReDim strBrandNames2(2)
    strBrandNames2(0) = "Arby's,"
    strBrandNames2(1) = "Budget Gourmet,"
    strBrandNames2(2) = "Campbell's,"
    
    For l = 0 To UBound(strBrandNames1())
        strDescription = Replace(strDescription, strBrandNames1(l), strBrandNames2(l), , , vbTextCompare)
    Next l
    
    '--Add a comma after food words
    ReDim strFoodWords1(0)
    strFoodWords1(0) = "Cucumber salad "
    
    ReDim strFoodWords2(0)
    strFoodWords2(0) = "Cucumber salad,"
    
    For l = 0 To UBound(strFoodWords1())
        strDescription = Replace(strDescription, strFoodWords1(l), strFoodWords2(l), , , vbTextCompare)
    Next l
    
    '--Replace ",," with a comma
    strDescription = Replace(strDescription, ",,", ",", , , vbTextCompare)
    
    lngInStrRev = InStrRev(strDescription, ",", , vbTextCompare)
    lngLen = Len(strDescription)
    If lngInStrRev = lngLen Then
        strDescription = Left(strDescription, lngLen - 1)
    End If
    
    FormattedSuggestDescr = strDescription

End Function

Private Function FormattedSuggestTerms(Terms() As String) As String()

    Dim l As Long
    Dim lngDiff As Long
    Dim lngIndex As Long
    Dim lngInStr As Long
    Dim lngInStrRev As Long
    Dim lngLen As Long
    Dim strTerm As String
    Dim strTerms() As String
    
    For l = 0 To UBound(Terms())
        strTerm = Trim$(Terms(l))
        If Len(strTerm) > 0 Then
            '--Remove ending or
            lngLen = Len(strTerm)
            lngInStrRev = InStrRev(strTerm, " or", lngLen, vbTextCompare)
            If lngInStrRev > 0 Then
                lngDiff = lngLen = lngInStrRev
                If lngDiff = 2 Or lngDiff = 3 Then
                    strTerm = Replace(strTerm, " or", vbNullString, 1, 1, vbTextCompare)
                    strTerm = Trim$(strTerm)
                End If
            End If
            
            '-- and
            If InStr(1, strTerm, "and ", vbTextCompare) = 1 Then
                strTerm = Replace(strTerm, "and ", vbNullString, 1, 1, vbTextCompare)
                ReDim Preserve strTerms(Utility.ArrayIndex(strTerms()))
                strTerms(UBound(strTerms())) = Trim$(strTerm)
            '-- and/or
            ElseIf InStr(1, strTerm, "and/or ", vbTextCompare) = 1 Then
                strTerm = Replace(strTerm, "and/or ", vbNullString, 1, 1, vbTextCompare)
                ReDim Preserve strTerms(Utility.ArrayIndex(strTerms()))
                strTerms(UBound(strTerms())) = Trim$(strTerm)
            '-- for
    '        ElseIf InStr(1, strTerm, "for ", vbTextCompare) = 1 Then
    '            strTerm = Replace(strTerm, "for ", vbNullString, 1, 1, vbTextCompare)
    '            ReDim Preserve strTerms(Utility.ArrayIndex(strTerms()))
    '            strTerms(UBound(strTerms())) = Trim$(strTerm)
            '-- formerly called
            ElseIf InStr(1, strTerm, "formerly called ", vbTextCompare) = 1 Then
                ReDim Preserve strTerms(Utility.ArrayIndex(strTerms()))
                strTerms(UBound(strTerms())) = "formerly called"
                strTerm = Replace(strTerm, "formerly called ", vbNullString, 1, 1, vbTextCompare)
                ReDim Preserve strTerms(Utility.ArrayIndex(strTerms()))
                strTerms(UBound(strTerms())) = Trim$(strTerm)
            '-- from
    '        ElseIf InStr(1, strTerm, "from ", vbTextCompare) = 1 Then
    '            strTerm = Replace(strTerm, "from ", vbNullString, 1, 1, vbTextCompare)
    '            ReDim Preserve strTerms(Utility.ArrayIndex(strTerms()))
    '            strTerms(UBound(strTerms())) = Trim$(strTerm)
            '-- in
            ElseIf InStr(1, strTerm, "in ", vbTextCompare) = 1 Then
                strTerm = Replace(strTerm, "in ", vbNullString, 1, 1, vbTextCompare)
                ReDim Preserve strTerms(Utility.ArrayIndex(strTerms()))
                strTerms(UBound(strTerms())) = Trim$(strTerm)
            '-- include
            ElseIf InStr(1, strTerm, "include ", vbTextCompare) = 1 Then
                strTerm = Replace(strTerm, "include ", vbNullString, 1, 1, vbTextCompare)
                ReDim Preserve strTerms(Utility.ArrayIndex(strTerms()))
                strTerms(UBound(strTerms())) = Trim$(strTerm)
            '-- included
            ElseIf InStr(1, strTerm, "included ", vbTextCompare) = 1 Then
                strTerm = Replace(strTerm, "included ", vbNullString, 1, 1, vbTextCompare)
                ReDim Preserve strTerms(Utility.ArrayIndex(strTerms()))
                strTerms(UBound(strTerms())) = Trim$(strTerm)
            '-- includes
            ElseIf InStr(1, strTerm, "includes ", vbTextCompare) = 1 Then
                strTerm = Replace(strTerm, "includes ", vbNullString, 1, 1, vbTextCompare)
                ReDim Preserve strTerms(Utility.ArrayIndex(strTerms()))
                strTerms(UBound(strTerms())) = Trim$(strTerm)
            '-- includes:
            ElseIf InStr(1, strTerm, "includes:", vbTextCompare) = 1 Then
                strTerm = Replace(strTerm, "includes:", vbNullString, 1, 1, vbTextCompare)
                ReDim Preserve strTerms(Utility.ArrayIndex(strTerms()))
                strTerms(UBound(strTerms())) = Trim$(strTerm)
            '-- including
            ElseIf InStr(1, strTerm, "including ", vbTextCompare) = 1 Then
                strTerm = Replace(strTerm, "including ", vbNullString, 1, 1, vbTextCompare)
                ReDim Preserve strTerms(Utility.ArrayIndex(strTerms()))
                strTerms(UBound(strTerms())) = Trim$(strTerm)
            '-- or
            ElseIf InStr(1, strTerm, "or ", vbTextCompare) = 1 Then
                strTerm = Replace(strTerm, "or ", vbNullString, 1, 1, vbTextCompare)
                ReDim Preserve strTerms(Utility.ArrayIndex(strTerms()))
                strTerms(UBound(strTerms())) = Trim$(strTerm)
            Else
                ReDim Preserve strTerms(Utility.ArrayIndex(strTerms()))
                strTerms(UBound(strTerms())) = Trim$(strTerm)
            End If
        End If
    Next l
    
    FormattedSuggestTerms = strTerms

End Function

Private Function FormattedWords(Words() As String) As String()

    Dim l As Long
    Dim m As Long
    Dim lngIndex As Long
    Dim lngInStr As Long
    Dim lngLen As Long
    Dim strWord As String
    Dim strWordList() As String
    Dim strWords() As String
    
    For l = 0 To UBound(Words())
        strWord = Trim$(Words(l))
        '-- (
        strWord = Replace(strWord, "(", " ", , , vbTextCompare)
        '-- )
        strWord = Replace(strWord, ")", " ", , , vbTextCompare)
        '-- "
        strWord = Replace(strWord, """", " ", , , vbTextCompare)
        '-- ;
        strWord = Replace(strWord, ";", " ", , , vbTextCompare)
        '-- :
        strWord = Replace(strWord, ":", " ", , , vbTextCompare)
        '-- ...
        strWord = Replace(strWord, "...", " ", , , vbTextCompare)
        '-- +
        strWord = Replace(strWord, "+", " ", , , vbTextCompare)
        '-- !
        strWord = Replace(strWord, "!", " ", , , vbTextCompare)
        '-- ,
        If InStr(1, strWord, ",", vbTextCompare) > 0 Then
            If Not strWord Like "*#,#*" Then
                strWord = Replace(strWord, ",", " ", , , vbTextCompare)
            End If
        End If
        '-- &
        If InStr(1, strWord, "&", vbTextCompare) > 0 Then
            If Not strWord Like "*m&m*" Then
                strWord = Replace(strWord, "&", " ", , , vbTextCompare)
            End If
        End If
        '-- /
        If InStr(1, strWord, "/", vbTextCompare) > 0 Then
            If Not strWord Like "*#/#*" Then
                strWord = Replace(strWord, "/", " ", , , vbTextCompare)
            End If
        End If
        '-- -
        If InStr(1, strWord, "-", vbTextCompare) > 0 Then
            If Not strWord Like "*#-#*" Then
                strWord = Replace(strWord, "-", " ", , , vbTextCompare)
            End If
        End If
        '-- .
        If InStr(1, strWord, ".", vbTextCompare) > 0 Then
            If strWord Like "*#.#*" Then
            ElseIf strWord Like "*?.?.*" Then
            ElseIf strWord Like "*?.?.?.?.?*" Then
            Else
                strWord = Replace(strWord, ".", " ", , , vbTextCompare)
            End If
        End If
        '-- '
        If InStr(1, strWord, "'", vbTextCompare) > 0 Then
            If strWord Like "*'s*" Then
                If Not strWord Like "*o's*" And Not strWord Like "*m&m's*" Then
                    strWord = Replace(strWord, "'s", " ", , , vbTextCompare)
                End If
            ElseIf strWord Like "*s'*" Then
                If strWord Like "*s'm*" Then
                ElseIf strWord Like "*s'n*" Then
                    If Not strWord Like "*'ner*" Then
                        strWord = Replace(strWord, "'", " ", , , vbTextCompare)
                    End If
                Else
                    strWord = Replace(strWord, "s'", " ", , , vbTextCompare)
                End If
            ElseIf strWord Like "*e'e*" Then
                strWord = Replace(strWord, "'", " ", , , vbTextCompare)
            ElseIf strWord Like "*l'*" Then
                If Not strWord Like "*l'i*" And Not strWord Like "*l's*" Then
                    strWord = Replace(strWord, "'", " ", , , vbTextCompare)
                End If
            ElseIf strWord Like "*in'*" Then
                strWord = Replace(strWord, "'", " ", , , vbTextCompare)
            ElseIf strWord Like "*n'*" Then
                strWord = Replace(strWord, "'", " ", , , vbTextCompare)
            ElseIf strWord Like "*'n*" Then
                If Not strWord Like "*'ner*" Then
                    strWord = Replace(strWord, "'", " ", , , vbTextCompare)
                End If
            ElseIf strWord Like "*o'" Then
                strWord = Replace(strWord, "'", " ", , , vbTextCompare)
            ElseIf strWord Like "'##" Then
'                strWord = Replace(strWord, "'", " ", , , vbTextCompare)
            End If
        End If
        
        '-- Trim
        strWord = Trim$(strWord)
        strWordList() = Split(strWord, " ", , vbTextCompare)
        For m = 0 To UBound(strWordList())
            ReDim Preserve strWords(Utility.ArrayIndex(strWords()))
            strWords(UBound(strWords())) = Trim$(strWordList(m))
        Next m
    Next l
    
    FormattedWords = strWords

End Function

Public Sub ImportData()
    
    '--Append record(s)
    'Call AppendFoodDescr
    'Call AppendFoodSearch
    'Call AppendFoodSuggest
    'Call AppendFoodWords
    'Call AppendPortions
    'Call AppendTagname
    'Call AppendNutrientDescr
    'Call AppendNutrients
    Call AppendEquivalentDescr
    Call AppendEquivalents
    'Call AppendIngredients
    'Call AppendIngredSearch
    'Call AppendIngredSuggest

End Sub

Private Function InitialWeight(FoodCode As Long, Version As Long) As Double

    comRecipeWeight_Lkp("@FoodCode") = FoodCode
    comRecipeWeight_Lkp("@Version") = Version
    With rstRecipeWeight_Lkp
        .Requery
        If .RecordCount > 0 Then
            InitialWeight = CDbl(.Fields("InitialWeight"))
        End If
    End With

End Function

Private Function Log10(x)

    Log10 = Log(x) / Log(10#)
    
End Function

Private Function MeasureDescription(MeasureCode As String) As String

    Select Case MeasureCode
        Case "C"
            MeasureDescription = "Cup(s)"
        Case "CP"
            MeasureDescription = "Cup(s)"
        Case "FO"
            MeasureDescription = "Fluid Ounce(s)"
        Case "GAL"
            MeasureDescription = "Gallon(s)"
        Case "GM"
            MeasureDescription = "Gram(s)"
        Case "LB"
            MeasureDescription = "Pound(s)"
        Case "MG"
            MeasureDescription = "Milligram(s)"
        Case "OZ"
            MeasureDescription = "Ounce(s)"
        Case "PT"
            MeasureDescription = "Pint(s)"
        Case "QT"
            MeasureDescription = "Quart(s)"
        Case "TB"
            MeasureDescription = "Tablespoon(s)"
        Case "TS"
            MeasureDescription = "Teaspoon(s)"
        Case "WO"
            MeasureDescription = "Weight Ounce(s)"
        Case Else
            Stop
    End Select
    
End Function

Private Function NutrientSortOrder(Tagname As String) As Long

    Select Case Tagname
        '-- Proximates
        Case "WATER"
            NutrientSortOrder = 1000
        Case "ENERC"
            NutrientSortOrder = 1010
        Case "PROCNT"
            NutrientSortOrder = 1020
        Case "FAT"
            NutrientSortOrder = 1030
        Case "CHOCDF"
            NutrientSortOrder = 1040
        Case "FIBTG"
            NutrientSortOrder = 1050
        Case "SUGAR"
            NutrientSortOrder = 1060
        '-- Minerals
        Case "CA"
            NutrientSortOrder = 2000
        Case "FE"
            NutrientSortOrder = 2010
        Case "MG"
            NutrientSortOrder = 2020
        Case "P"
            NutrientSortOrder = 2030
        Case "K"
            NutrientSortOrder = 2040
        Case "NA"
            NutrientSortOrder = 2050
        Case "ZN"
            NutrientSortOrder = 2060
        Case "CU"
            NutrientSortOrder = 2070
        Case "SE"
            NutrientSortOrder = 2080
        '-- Vitamins
        Case "VITC"
            NutrientSortOrder = 3000
        Case "THIA"
            NutrientSortOrder = 3010
        Case "RIBF"
            NutrientSortOrder = 3020
        Case "NIA"
            NutrientSortOrder = 3030
        Case "VITB6A"
            NutrientSortOrder = 3040
        Case "FOL"
            NutrientSortOrder = 3050
        Case "FOLAC"
            NutrientSortOrder = 3060
        Case "FOLFD"
            NutrientSortOrder = 3070
        Case "FOLDFE"
            NutrientSortOrder = 3080
        Case "CHOLN"
            NutrientSortOrder = 3090
        Case "VITB12"
            NutrientSortOrder = 3100
        Case "VITB12_ADDED"
            NutrientSortOrder = 3110
        Case "VITA"
            NutrientSortOrder = 3120
        Case "RETOL"
            NutrientSortOrder = 3130
        Case "TOCPHA"
            NutrientSortOrder = 3140
        Case "TOCPHA_ADDED"
            NutrientSortOrder = 3150
        Case "VITD"
            NutrientSortOrder = 3160
        Case "VITK"
            NutrientSortOrder = 3170
        '-- Lipids
        Case "FASAT"
            NutrientSortOrder = 4000
        Case "F4D0"
            NutrientSortOrder = 4010
        Case "F6D0"
            NutrientSortOrder = 4020
        Case "F8D0"
            NutrientSortOrder = 4030
        Case "F10D0"
            NutrientSortOrder = 4040
        Case "F12D0"
            NutrientSortOrder = 4050
        Case "F14D0"
            NutrientSortOrder = 4060
        Case "F16D0"
            NutrientSortOrder = 4070
        Case "F18D0"
            NutrientSortOrder = 4080
        Case "FAMS"
            NutrientSortOrder = 4090
        Case "F16D1"
            NutrientSortOrder = 4100
        Case "F18D1"
            NutrientSortOrder = 4110
        Case "F20D1"
            NutrientSortOrder = 4120
        Case "F22D1"
            NutrientSortOrder = 4130
        Case "FAPU"
            NutrientSortOrder = 4140
        Case "F18D2"
            NutrientSortOrder = 4150
        Case "F18D3"
            NutrientSortOrder = 4160
        Case "F18D4"
            NutrientSortOrder = 4170
        Case "F20D4"
            NutrientSortOrder = 4180
        Case "F20D5"
            NutrientSortOrder = 4190
        Case "F22D5"
            NutrientSortOrder = 4200
        Case "F22D6"
            NutrientSortOrder = 4210
        Case "CHOLE"
            NutrientSortOrder = 4220
        '-- Others
        Case "ALC"
            NutrientSortOrder = 5000
        Case "CAFFN"
            NutrientSortOrder = 5010
        Case "THEBRN"
            NutrientSortOrder = 5020
        Case "CARTB"
            NutrientSortOrder = 5030
        Case "CARTA"
            NutrientSortOrder = 5040
        Case "CRYPX"
            NutrientSortOrder = 5050
        Case "LYCPN"
            NutrientSortOrder = 5060
        Case "LUTNZEA"
            NutrientSortOrder = 5070
        Case Else
            Stop
    End Select

End Function

Private Function NutrientTagname(NutrientCode As Long, Version As Long) As String

    Dim strTagname As String
    
    If NutrientCode > 0 Then
        comTagname_Lkp("@NutrientCode") = NutrientCode
        comTagname_Lkp("@Version") = Version
        With rstTagname_Lkp
            .Requery
            If .RecordCount > 0 Then
                If IsNull(.Fields("Tagname")) Then
                    If NutrientCode = 573 Then
                        strTagname = "TOCPHA_ADDED"
                    ElseIf NutrientCode = 578 Then
                        strTagname = "VITB12_ADDED"
                    Else
                        Stop
                    End If
                Else
                    strTagname = Trim$(.Fields("Tagname"))
                    '-- Take care of 3 nutrients whose tagnames do not match INFOODS
                    If NutrientCode = 208 Then
                        strTagname = "ENERC"
                    ElseIf NutrientCode = 320 Then
                        strTagname = "VITA"
                    ElseIf NutrientCode = 430 Then
                        strTagname = "VITK"
                    ElseIf StrComp(strTagname, "LUTN", vbTextCompare) = 0 Or StrComp(strTagname, "LUT+ZEA", vbTextCompare) = 0 Then
                        strTagname = "LUTNZEA"
                    End If
                End If
            Else
                strTagname = vbNullString
            End If
        End With
    Else
        strTagname = vbNullString
    End If
    
    NutrientTagname = strTagname
        
End Function

Public Sub OpenCommands()

    Dim SQL As String
    Dim prm As ADODB.Parameter
    
    Set comAddtlDescr_Lkp = New ADODB.Command
    With comAddtlDescr_Lkp
        .ActiveConnection = cnnFNDDS
        .CommandText = "SELECT AdditionalFoodDescription " & _
            "FROM AddFoodDesc " & _
            "WHERE (FoodCode = ?) AND " & _
            "(Version = ?) " & _
            "ORDER BY SeqNum"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@FoodCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstAddtlDescr_Lkp = New ADODB.Recordset
    rstAddtlDescr_Lkp.Open comAddtlDescr_Lkp, , adOpenStatic, adLockReadOnly

    Set comCountInDocument_Lkp = New ADODB.Command
    With comCountInDocument_Lkp
        .ActiveConnection = cnnBack
        .CommandText = "SELECT SUM(WordCount) AS CountInDocument " & _
            "FROM foodword " & _
            "WHERE (FoodCode = ?) AND " & _
            "(ModCode = ?) AND " & _
            "(Version = ?) AND " & _
            "(WordID = ?) AND " & _
            "((WordType = ?) OR " & _
            "(WordType = ?))"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@FoodCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@ModCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@WordID", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@WordType1", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@WordType2", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstCountInDocument_Lkp = New ADODB.Recordset
    rstCountInDocument_Lkp.Open comCountInDocument_Lkp, , adOpenStatic, adLockReadOnly
    
    Set comCountInDocuments_Lkp = New ADODB.Command
    With comCountInDocuments_Lkp
        .ActiveConnection = cnnBack
        .CommandText = "SELECT DocumentCount " & _
            "FROM worddocument " & _
            "WHERE (WordID = ?) AND " & _
            "(Version = ?) AND " & _
            "(WordType = ?)"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@WordID", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@WordType", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstCountInDocuments_Lkp = New ADODB.Recordset
    rstCountInDocuments_Lkp.Open comCountInDocuments_Lkp, , adOpenStatic, adLockReadOnly
    
    Set comDocumentCount_Lkp = New ADODB.Command
    With comDocumentCount_Lkp
        .ActiveConnection = cnnBack
        .CommandText = "SELECT COUNT(*) AS DocumentCount " & _
            "FROM fooddescr " & _
            "WHERE (Version = ?)"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstDocumentCount_Lkp = New ADODB.Recordset
    rstDocumentCount_Lkp.Open comDocumentCount_Lkp, , adOpenStatic, adLockReadOnly
    
    Set comFCDescr_Lkp = New ADODB.Command
    With comFCDescr_Lkp
        .ActiveConnection = cnnFNDDS
        .CommandText = "SELECT MainFoodDescription AS Description " & _
            "FROM MainFoodDesc " & _
            "WHERE (FoodCode = ?) AND " & _
            "(Version = ?)"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@FoodCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstFCDescr_Lkp = New ADODB.Recordset
    rstFCDescr_Lkp.Open comFCDescr_Lkp, , adOpenStatic, adLockReadOnly
    
    Set comFoodMatrixA_Lkp = New ADODB.Command
    With comFoodMatrixA_Lkp
        .ActiveConnection = cnnBack
        .CommandText = "SELECT FoodCode," & _
            "ModCode," & _
            "Version," & _
            "WordID," & _
            "WordType," & _
            "tf_idf " & _
            "FROM foodword " & _
            "WHERE (FoodCode = ?) AND " & _
            "(ModCode = ?) AND " & _
            "(Version = ?) AND " & _
            "(WordType = ?)"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@FoodCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@ModCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@WordType", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstFoodMatrixA_Lkp = New ADODB.Recordset
    rstFoodMatrixA_Lkp.Open comFoodMatrixA_Lkp, , adOpenKeyset, adLockOptimistic
    
    Set comFoodMatrixB_Lkp = New ADODB.Command
    With comFoodMatrixB_Lkp
        .ActiveConnection = cnnBack
        .CommandText = "SELECT FoodCode," & _
            "ModCode," & _
            "Version," & _
            "WordID," & _
            "WordType," & _
            "tf_idf " & _
            "FROM foodword " & _
            "WHERE (FoodCode = ?) AND " & _
            "(ModCode = ?) AND " & _
            "(Version = ?) AND " & _
            "(WordType = ?)"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@FoodCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@ModCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@WordType", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstFoodMatrixB_Lkp = New ADODB.Recordset
    rstFoodMatrixB_Lkp.Open comFoodMatrixB_Lkp, , adOpenKeyset, adLockOptimistic
    
    Set comFoodMatrixValue_Lkp = New ADODB.Command
    With comFoodMatrixValue_Lkp
        .ActiveConnection = cnnBack
        .CommandText = "SELECT tf_idf " & _
            "FROM foodword " & _
            "WHERE (FoodCode = ?) AND " & _
            "(ModCode = ?) AND " & _
            "(Version = ?) AND " & _
            "(WordID = ?) AND " & _
            "((WordType = ?) OR " & _
            "(WordType = ?))"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@FoodCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@ModCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@WordID", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@WordType1", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@WordType2", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstFoodMatrixValue_Lkp = New ADODB.Recordset
    rstFoodMatrixValue_Lkp.Open comFoodMatrixValue_Lkp, , adOpenKeyset, adLockOptimistic
    
'    Set comModNutrient_Lkp = New ADODB.Command
'    With comModNutrient_Lkp
'        .ActiveConnection = cnnFNDDS
'        .CommandText = "SELECT NutrientCode," & _
'            "NutrientValue " & _
'            "FROM ModNutVal " & _
'            "WHERE (FoodCode = ?) AND " & _
'            "(ModificationCode = ?) AND " & _
'            "(Version = ?) " & _
'            "ORDER BY NutrientCode"
'        .CommandType = adCmdText
'        Set prm = .CreateParameter("@FoodCode", adBigInt, adParamInput, , 0)
'        .Parameters.Append prm
'        Set prm = .CreateParameter("@ModificationCode", adBigInt, adParamInput, , 0)
'        .Parameters.Append prm
'        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
'        .Parameters.Append prm
'        .Prepared = True
'    End With
'    Set rstModNutrient_Lkp = New ADODB.Recordset
'    rstModNutrient_Lkp.Open comModNutrient_Lkp, , adOpenStatic, adLockReadOnly
    
'    Set comMPED_Lkp = New ADODB.Command
'    With comMPED_Lkp
'        .ActiveConnection = cnnFNDDS
'        SQL = "SELECT G_TOTAL," & _
'            "G_WHL," & _
'            "G_NWHL," & _
'            "V_TOTAL," & _
'            "V_DRKGR," & _
'            "V_ORANGE," & _
'            "V_POTATO," & _
'            "V_STARCY," & _
'            "V_TOMATO," & _
'            "V_OTHER," & _
'            "F_TOTAL," & _
'            "F_CITMLB," & _
'            "F_OTHER," & _
'            "D_TOTAL,"
'        SQL = SQL & "D_MILK," & _
'            "D_YOGURT," & _
'            "D_CHEESE," & _
'            "M_MPF," & _
'            "M_MEAT," & _
'            "M_ORGAN," & _
'            "M_FRANK," & _
'            "M_POULT," & _
'            "M_FISH_HI," & _
'            "M_FISH_LO," & _
'            "M_EGG," & _
'            "M_SOY," & _
'            "M_NUTSD," & _
'            "LEGUMES," & _
'            "DISCFAT_OIL," & _
'            "DISCFAT_SOL," & _
'            "ADD_SUG," & _
'            "A_BEV "
'        SQL = SQL & "FROM Equivalent " & _
'            "WHERE (FOODCODE = ?) AND " & _
'            "(MODCODE = ?) AND " & _
'            "(Version = ?)"
'        .CommandText = SQL
'        .CommandType = adCmdText
'        Set prm = .CreateParameter("@FoodCode", adBigInt, adParamInput, , 0)
'        .Parameters.Append prm
'        Set prm = .CreateParameter("@ModCode", adBigInt, adParamInput, , 0)
'        .Parameters.Append prm
'        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
'        .Parameters.Append prm
'        .Prepared = True
'    End With
'    Set rstMPED_Lkp = New ADODB.Recordset
'    rstMPED_Lkp.Open comMPED_Lkp, , adOpenStatic, adLockReadOnly
    
'    Set comNutrient_Lkp = New ADODB.Command
'    With comNutrient_Lkp
'        .ActiveConnection = cnnFNDDS
'        .CommandText = "SELECT NutrientCode," & _
'            "NutrientValue " & _
'            "FROM AddFoodDesc " & _
'            "WHERE (FoodCode = ?) AND " & _
'            "(Version = ?) " & _
'            "ORDER BY NutrientCode"
'        .CommandType = adCmdText
'        Set prm = .CreateParameter("@FoodCode", adBigInt, adParamInput, , 0)
'        .Parameters.Append prm
'        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
'        .Parameters.Append prm
'        .Prepared = True
'    End With
'    Set rstNutrient_Lkp = New ADODB.Recordset
'    rstNutrient_Lkp.Open comNutrient_Lkp, , adOpenStatic, adLockReadOnly
    
    Set comPortionDescr_Lkp = New ADODB.Command
    With comPortionDescr_Lkp
        .ActiveConnection = cnnFNDDS
        .CommandText = "SELECT PortionDescription," & _
            "ChangeType " & _
            "FROM FoodPortionDesc " & _
            "WHERE (PortionCode = ?) AND " & _
            "(Version = ?)"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@PortionCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstPortionDescr_Lkp = New ADODB.Recordset
    rstPortionDescr_Lkp.Open comPortionDescr_Lkp, , adOpenStatic, adLockReadOnly
    
    Set comPortions_Lkp = New ADODB.Command
    With comPortions_Lkp
        .ActiveConnection = cnnBack
        .CommandText = "SELECT Subcode," & _
            "SubcodeDescr," & _
            "SeqNum," & _
            "PortionCode," & _
            "PortionDescr," & _
            "PortionChangeType," & _
            "Weight," & _
            "WeightChangeType " & _
            "FROM portions " & _
            "WHERE (FoodCode = ?) AND (ModCode = ?) AND (Version = ?) " & _
            "ORDER BY Subcode, SeqNum"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@FoodCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@ModCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstPortions_Lkp = New ADODB.Recordset
    rstPortions_Lkp.Open comPortions_Lkp, , adOpenStatic, adLockReadOnly
    
    Set comRecipeWeight_Lkp = New ADODB.Command
    With comRecipeWeight_Lkp
        .ActiveConnection = cnnFNDDS
        .CommandText = "SELECT SUM(Weight) AS InitialWeight " & _
            "FROM FnddsSrLinks " & _
            "WHERE (FoodCode = ?) AND " & _
            "(Version = ?)"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@FoodCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstRecipeWeight_Lkp = New ADODB.Recordset
    rstRecipeWeight_Lkp.Open comRecipeWeight_Lkp, , adOpenStatic, adLockReadOnly

    Set comRetDescr_Lkp = New ADODB.Command
    With comRetDescr_Lkp
        .ActiveConnection = cnnSR
        .CommandText = "SELECT DISTINCT RetnDesc " & _
            "FROM tblRETENTION " & _
            "WHERE Retn_Code = ?"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@RetCode", adVarChar, adParamInput, 4, adCmdText)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstRetDescr_Lkp = New ADODB.Recordset
    rstRetDescr_Lkp.Open comRetDescr_Lkp, , adOpenStatic, adLockReadOnly

    Set comSRDescr_Lkp = New ADODB.Command
    With comSRDescr_Lkp
        .ActiveConnection = cnnSR
        .CommandText = "SELECT Long_Desc " & _
            "FROM tblFOOD_DES " & _
            "WHERE (NDB_No = ?) AND " & _
            "(Version = ?)"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@SRCode", adVarChar, adParamInput, 5, adCmdText)
        .Parameters.Append prm
        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstSRDescr_Lkp = New ADODB.Recordset
    rstSRDescr_Lkp.Open comSRDescr_Lkp, , adOpenStatic, adLockReadOnly

    Set comSubcode_Lkp = New ADODB.Command
    With comSubcode_Lkp
        .ActiveConnection = cnnFNDDS
        .CommandText = "SELECT SubcodeDescription " & _
            "FROM SubcodeDesc " & _
            "WHERE (Subcode = ?) AND " & _
            "(Version = ?)"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@Subcode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstSubcode_Lkp = New ADODB.Recordset
    rstSubcode_Lkp.Open comSubcode_Lkp, , adOpenStatic, adLockReadOnly
    
    Set comSuggest_Lkp = New ADODB.Command
    With comSuggest_Lkp
        .ActiveConnection = cnnBack
        .CommandText = "SELECT SuggestID " & _
            "FROM suggest " & _
            "WHERE (SuggestType = ?) AND " & _
            "(SuggestDescription = ?)"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@SuggestType", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@Description", adVarChar, adParamInput, 200, vbNullString)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstSuggest_Lkp = New ADODB.Recordset
    rstSuggest_Lkp.Open comSuggest_Lkp, , adOpenStatic, adLockReadOnly
                
    Set comSuggestID_Lkp = New ADODB.Command
    With comSuggestID_Lkp
        .ActiveConnection = cnnBack
        .CommandText = "SELECT MAX(SuggestID) AS SuggestID " & _
            "FROM suggest " & _
            "WHERE (SuggestType = ?)"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@SuggestType", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstSuggestID_Lkp = New ADODB.Recordset
    rstSuggestID_Lkp.Open comSuggestID_Lkp, , adOpenStatic, adLockReadOnly
    
    Set comSuggestFoodCount_Lkp = New ADODB.Command
    With comSuggestFoodCount_Lkp
        .ActiveConnection = cnnBack
        .CommandText = "SELECT FoodCode," & _
            "ModCode," & _
            "Version," & _
            "SuggestID," & _
            "SuggestType," & _
            "SuggestCount " & _
            "FROM foodsuggest " & _
            "WHERE (FoodCode = ?) AND " & _
            "(ModCode = ?) AND " & _
            "(Version = ?) AND " & _
            "(SuggestID = ?) AND " & _
            "(SuggestType = ?)"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@FoodCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@ModCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@SuggestID", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@SuggestType", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstSuggestFoodCount_Lkp = New ADODB.Recordset
    rstSuggestFoodCount_Lkp.Open comSuggestFoodCount_Lkp, , adOpenKeyset, adLockOptimistic
    
    Set comSuggestIngredCount_Lkp = New ADODB.Command
    With comSuggestIngredCount_Lkp
        .ActiveConnection = cnnBack
        .CommandText = "SELECT SRCode," & _
            "Version," & _
            "SuggestID," & _
            "SuggestType," & _
            "SuggestCount " & _
            "FROM ingredsuggest " & _
            "WHERE (SRCode = ?) AND " & _
            "(Version = ?) AND " & _
            "(SuggestID = ?) AND " & _
            "(SuggestType = ?)"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@SRCode", adVarChar, adParamInput, 8, vbNullString)
        .Parameters.Append prm
        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@SuggestID", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@SuggestType", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstSuggestIngredCount_Lkp = New ADODB.Recordset
    rstSuggestIngredCount_Lkp.Open comSuggestIngredCount_Lkp, , adOpenKeyset, adLockOptimistic
    
    Set comTagname_Lkp = New ADODB.Command
    With comTagname_Lkp
        .ActiveConnection = cnnFNDDS
        .CommandText = "SELECT Tagname " & _
            "FROM NutDesc " & _
            "WHERE (NutrientCode = ?) AND " & _
            "(Version = ?)"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@NutrientCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstTagname_Lkp = New ADODB.Recordset
    rstTagname_Lkp.Open comTagname_Lkp, , adOpenStatic, adLockReadOnly
    
    Set comWord_Lkp = New ADODB.Command
    With comWord_Lkp
        .ActiveConnection = cnnBack
        .CommandText = "SELECT WordID " & _
            "FROM word " & _
            "WHERE (WordDescription = ?)"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@Description", adVarChar, adParamInput, 200, vbNullString)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstWord_Lkp = New ADODB.Recordset
    rstWord_Lkp.Open comWord_Lkp, , adOpenStatic, adLockReadOnly
                
    Set comWordID_Lkp = New ADODB.Command
    With comWordID_Lkp
        .ActiveConnection = cnnBack
        .CommandText = "SELECT MAX(WordID) AS WordID FROM word"
        .CommandType = adCmdText
        .Prepared = True
    End With
    Set rstWordID_Lkp = New ADODB.Recordset
    rstWordID_Lkp.Open comWordID_Lkp, , adOpenStatic, adLockReadOnly
    
    Set comUpdateWordCount = New ADODB.Command
    With comUpdateWordCount
        .ActiveConnection = cnnBack
        .CommandText = "SELECT FoodCode," & _
            "ModCode," & _
            "Version," & _
            "WordID," & _
            "WordType," & _
            "WordCount " & _
            "FROM foodword " & _
            "WHERE (FoodCode = ?) AND " & _
            "(ModCode = ?) AND " & _
            "(Version = ?) AND " & _
            "(WordID = ?) AND " & _
            "(WordType = ?)"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@FoodCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@ModCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@WordID", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@WordType", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstUpdateWordCount = New ADODB.Recordset
    rstUpdateWordCount.Open comUpdateWordCount, , adOpenKeyset, adLockOptimistic
    
    Set comWordCount_Lkp = New ADODB.Command
    With comWordCount_Lkp
        .ActiveConnection = cnnBack
        .CommandText = "SELECT SUM(WordCount) AS WordCount " & _
            "FROM foodword " & _
            "WHERE (FoodCode = ?) AND " & _
            "(ModCode = ?) AND " & _
            "(Version = ?) AND " & _
            "(WordID = ?) AND " & _
            "((WordType = ?) OR " & _
            "(WordType = ?))"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@FoodCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@ModCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@WordID", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@WordType1", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@WordType2", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstWordCount_Lkp = New ADODB.Recordset
    rstWordCount_Lkp.Open comWordCount_Lkp, , adOpenStatic, adLockReadOnly
    
    Set comWordsInDoc_Lkp = New ADODB.Command
    With comWordsInDoc_Lkp
        .ActiveConnection = cnnBack
        .CommandText = "SELECT SUM(WordCount) AS WordsInDocument " & _
            "FROM foodword " & _
            "WHERE (FoodCode = ?) AND " & _
            "(ModCode = ?) AND " & _
            "(Version = ?) AND " & _
            "((WordType = ?) OR " & _
            "(WordType = ?))"
        .CommandType = adCmdText
        Set prm = .CreateParameter("@FoodCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@ModCode", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@Version", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@WordType1", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        Set prm = .CreateParameter("@WordType2", adBigInt, adParamInput, , 0)
        .Parameters.Append prm
        .Prepared = True
    End With
    Set rstWordsInDoc_Lkp = New ADODB.Recordset
    rstWordsInDoc_Lkp.Open comWordsInDoc_Lkp, , adOpenStatic, adLockReadOnly

End Sub

Private Function PortionChangeType(PortionCode As Long, Version As Long) As String

    comPortionDescr_Lkp("@PortionCode") = PortionCode
    comPortionDescr_Lkp("@Version") = Version
    With rstPortionDescr_Lkp
        .Requery
        If .RecordCount > 0 Then
            If Not IsNull(.Fields("ChangeType")) Then
                PortionChangeType = .Fields("ChangeType")
            End If
        End If
    End With

End Function

Private Function PortionDescr(PortionCode As Long, Version As Long) As String

    comPortionDescr_Lkp("@PortionCode") = PortionCode
    comPortionDescr_Lkp("@Version") = Version
    With rstPortionDescr_Lkp
        .Requery
        If .RecordCount > 0 Then
            PortionDescr = .Fields("PortionDescription")
        End If
    End With

End Function

Public Sub RebuildTables()

    DropConstraints
    DropTables
    CreateTables

End Sub

Private Function RetentionDescription(RetCode As String) As String

    comRetDescr_Lkp("@RetCode") = RetCode
    With rstRetDescr_Lkp
        .Requery
        If .RecordCount > 0 Then
            RetentionDescription = .Fields("RetnDesc")
        Else
            Stop
        End If
    End With

End Function

Private Function SRCode(Code As String) As String

    Dim lngLen As Long
    Dim strCode As String

    strCode = Trim$(Code)
    lngLen = Len(strCode)
    If lngLen < 5 Then
        SRCode = String(5 - lngLen, "0") & strCode
    Else
        SRCode = strCode
    End If

End Function

Private Function SRDescription(SRCode As String, Version As Long, Abbreviation As String) As String

    Dim blnFound As Boolean
    Dim l As Long
    
    If Len(SRCode) = 5 Then
        comSRDescr_Lkp("@SRCode") = SRCode
        comSRDescr_Lkp("@Version") = Version ^ 2
        With rstSRDescr_Lkp
            .Requery
            If .RecordCount > 0 Then
                SRDescription = .Fields("Long_Desc")
            Else
                '--Look in Missing Codes file
                l = 2
                Do Until Len(wstExcel1.Range("D" & CStr(l)).Value) = 0
                    If StrComp(SRCode, CStr(wstExcel1.Range("D" & CStr(l)).Value), vbTextCompare) = 0 Then
                        SRDescription = Trim$(wstExcel1.Range("E" & CStr(l)).Value)
                        blnFound = True
                        Exit Do
                    End If
                    l = l + 1
                Loop
                If Not blnFound Then
'                    Debug.Print SRCode, Abbreviation, StrConv(Abbreviation, vbProperCase)
'                    SRDescription = StrConv(Abbreviation, vbProperCase)
                    SRDescription = Trim$(Abbreviation)
                End If
            End If
        End With
    Else
        comFCDescr_Lkp("@FoodCode") = CLng(SRCode)
        comFCDescr_Lkp("@Version") = Version
        With rstFCDescr_Lkp
            .Requery
            If .RecordCount > 0 Then
                SRDescription = .Fields("Description")
            Else
                Stop
            End If
        End With
    End If

End Function

Private Function SubcodeDescr(Subcode As Long, Version As Long) As String

    If Subcode > 0 Then
        comSubcode_Lkp("@Subcode") = Subcode
        comSubcode_Lkp("@Version") = Version
        With rstSubcode_Lkp
            .Requery
            If .RecordCount > 0 Then
                SubcodeDescr = .Fields("SubcodeDescription")
            End If
        End With
    Else
        SubcodeDescr = "(Default)"
    End If

End Function

Private Function SuggestTermExists(SuggestType As Long, Term As String) As Long

    comSuggest_Lkp("@SuggestType") = SuggestType
    comSuggest_Lkp("@Description") = Term
    With rstSuggest_Lkp
        .Requery
        If .RecordCount > 0 Then
            SuggestTermExists = CLng(.Fields("SuggestID"))
        End If
    End With

End Function

Private Function SuggestTermID(SuggestType As Long) As Long

    comSuggestID_Lkp("@SuggestType") = SuggestType
    With rstSuggestID_Lkp
        .Requery
        If .RecordCount > 0 Then
            If Not IsNull(.Fields("SuggestID")) Then
                SuggestTermID = CLng(.Fields("SuggestID"))
            End If
        End If
    End With

End Function

Private Sub UpdateAdditionalDescriptions(FoodCode As Long, Version As Long, Recordset As ADODB.Recordset)

    Dim l As Long
    Dim strIncludesDescription As String
    
    comAddtlDescr_Lkp("@FoodCode") = FoodCode
    comAddtlDescr_Lkp("@Version") = Version
    rstAddtlDescr_Lkp.Requery
    
    With rstAddtlDescr_Lkp
        If .RecordCount > 0 Then
            l = 0
            strIncludesDescription = vbNullString
            Do Until .EOF
                strIncludesDescription = strIncludesDescription & .Fields("AdditionalFoodDescription")
                l = l + 1
                .MoveNext
                If Not .EOF Then
                    If l = (.RecordCount - 1) Then
                        strIncludesDescription = strIncludesDescription & "; or "
                    Else
                        strIncludesDescription = strIncludesDescription & "; "
                    End If
                End If
            Loop
        End If
    End With

    If Len(strIncludesDescription) > 0 Then
        Recordset.Fields("IncludesDescription") = strIncludesDescription
    End If
    
End Sub

Public Sub UpdateData()
    
    '--Append record(s)
'    Call UpdateDocumentCount
'    Call UpdateFoodWords
'    Call UpdateSimilarity

End Sub

Private Sub UpdateDocumentCount()

    Dim lngDocumentCount As Long
    Dim lngVersion As Long
    Dim lngWordID As Long
    Dim lngWordType As Long
    Dim SQL As String
    Dim rst1 As ADODB.Recordset
    Dim rst2 As ADODB.Recordset

    SQL = "SELECT DISTINCT WordID," & _
        "Version," & _
        "WordType " & _
        "FROM foodword " & _
        "ORDER BY WordID," & _
        "Version," & _
        "WordType"
    Set rst1 = New ADODB.Recordset
    rst1.Open SQL, cnnBack, adOpenStatic, adLockReadOnly, adCmdText
    
    SQL = "SELECT WordID," & _
        "WordType," & _
        "Version," & _
        "DocumentCount " & _
        "FROM worddocument " & _
        "WHERE (WordID = 0)"
    Set rst2 = New ADODB.Recordset
    rst2.Open SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText
    
    Do Until rst1.EOF
        lngWordID = CLng(rst1("WordID"))
        lngVersion = CLng(rst1("Version"))
        lngWordType = CLng(rst1("WordType"))
        lngDocumentCount = CountInDocuments(lngWordID, lngVersion, lngWordType)
        
        rst2.AddNew
        rst2("WordID") = lngWordID
        rst2("WordType") = lngWordType
        rst2("Version") = lngVersion
        rst2("DocumentCount") = lngDocumentCount
        rst2.Update
        
        rst1.MoveNext
    Loop
    
    rst1.Close
    Set rst1 = Nothing
    rst2.Close
    Set rst2 = Nothing
    
End Sub

Private Sub UpdateFoodWords()

    Dim lngCountInDocuments As Long
    Dim lngDocumentCount As Long
    Dim lngFoodCode As Long
    Dim lngModCode As Long
    Dim lngVersion As Long
    Dim lngWordCount As Long
    Dim lngWordID As Long
    Dim lngWordsInDocument As Long
    Dim lngWordType As Long
    Dim SQL As String
    Dim colDocumentCount As VBA.Collection
    Dim rst As ADODB.Recordset
    
    Set colDocumentCount = New VBA.Collection
    
    lngDocumentCount = DocumentCount(1)
    Call colDocumentCount.Add(lngDocumentCount, "1_")
    lngDocumentCount = DocumentCount(2)
    Call colDocumentCount.Add(lngDocumentCount, "2_")
    lngDocumentCount = DocumentCount(4)
    Call colDocumentCount.Add(lngDocumentCount, "4_")
    lngDocumentCount = DocumentCount(8)
    Call colDocumentCount.Add(lngDocumentCount, "8_")
    
    SQL = "SELECT FoodCode," & _
        "ModCode," & _
        "Version," & _
        "WordID," & _
        "WordType," & _
        "WordCount," & _
        "tf_idf " & _
        "FROM foodword " & _
        "ORDER BY FoodCode," & _
        "ModCode," & _
        "Version," & _
        "WordID"
    Set rst = New ADODB.Recordset
    rst.Open SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText
    Do Until rst.EOF
        lngFoodCode = CLng(rst("FoodCode"))
        lngModCode = CLng(rst("ModCode"))
        lngVersion = CLng(rst("Version"))
        lngWordID = CLng(rst("WordID"))
        lngWordType = CLng(rst("WordType"))
        
        lngWordCount = WordCount(lngFoodCode, lngModCode, lngVersion, lngWordID, lngWordType)
        lngWordsInDocument = WordsInDocument(lngFoodCode, lngModCode, lngVersion, lngWordType)
        lngCountInDocuments = CountInDocuments(lngWordID, lngVersion, lngWordType)
        lngDocumentCount = colDocumentCount.item(CStr(lngVersion) & "_")
        rst("tf_idf") = (lngWordCount / lngWordsInDocument) * Log10(lngDocumentCount / lngCountInDocuments)
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing

    Set colDocumentCount = Nothing

End Sub

Private Sub UpdateIngredSearch(FoodCode As Long, ModCode As Long, RecipeCode As Long, SeqNum As Long, Version As Long, level As Long, Recordset As ADODB.Recordset)

    Dim lngFlag As Long
    Dim lngIngredType As Long
    Dim SQL As String
    Dim strDescription As String
    Dim strSRCode As String
    Dim strSRDescr As String
    Dim rst As ADODB.Recordset
    
    SQL = "SELECT SRCode," & _
        "SRDescription AS Description," & _
        "Flag " & _
        "FROM FnddsSrLinks " & _
        "WHERE (FoodCode = " & CStr(RecipeCode) & ") AND " & _
        "(Version = " & CStr(Version) & ") " & _
        "GROUP BY SRCode," & _
        "SRDescription," & _
        "Flag " & _
        "ORDER BY SRCode"
    Set rst = New ADODB.Recordset
    Call rst.Open(SQL, cnnFNDDS, adOpenKeyset, adLockOptimistic, adCmdText)
    
    Do Until rst.EOF
        SeqNum = SeqNum + 1
        strSRCode = SRCode(rst("SRCode"))
        strDescription = rst("Description")
        strSRDescr = SRDescription(strSRCode, Version, strDescription)
        If IsNull(rst("Flag")) Then
            lngFlag = 0
        Else
            lngFlag = CLng(rst("Flag"))
        End If
        With Recordset
            .AddNew
            .Fields("FoodCode") = FoodCode
            .Fields("ModCode") = ModCode
            .Fields("SeqNum") = SeqNum
            If Len(strSRCode) = 5 Then
                lngIngredType = 1 + level
            Else
                lngIngredType = 2 + level
            End If
            .Fields("IngredType") = lngIngredType
            .Fields("IngrCode") = strSRCode
            .Fields("IngrDescr") = strDescription
            If lngFlag = 2 Then
'                Debug.Print strDescription, "->", strSRDescr
                .Fields("IngrDescrAlt") = strDescription
            Else
                .Fields("IngrDescrAlt") = strSRDescr
            End If
            .Fields("Version") = Version
            .Update
            If lngIngredType Mod 2 = 0 Then
                Call UpdateIngredSearch(FoodCode, ModCode, CLng(strSRCode), SeqNum, Version, lngIngredType, Recordset)
            End If
        End With
        rst.MoveNext
    Loop
    
    rst.Close
    Set rst = Nothing

End Sub

Private Sub UpdateSimilarity()

    Dim dblDotProduct As Double
    Dim dblMagnitudeA As Double
    Dim dblMagnitudeB As Double
    Dim dblMatrixValueA As Double
    Dim dblMatrixValueB As Double
    Dim dblSimilarity As Double
    Dim lngFoodCodeA As Long
    Dim lngFoodCodeB As Long
    Dim lngModCodeA As Long
    Dim lngModCodeB As Long
    Dim lngVersion As Long
    Dim lngWordID As Long
    Dim SQL As String
    Dim strFoodMatrixWordIDs As String
    Dim rst1 As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    Dim rst3 As ADODB.Recordset
    
    '--Open recordset with list of all food codes
    SQL = "SELECT FoodCode," & _
        "ModCode," & _
        "Version " & _
        "FROM fooddescr " & _
        "ORDER BY FoodCode," & _
        "ModCode," & _
        "Version"
    Set rst1 = New ADODB.Recordset
    rst1.Open SQL, cnnBack, adOpenStatic, adLockReadOnly, adCmdText
    
    '--Open recordset for similarity values
    SQL = "SELECT FoodCodeA," & _
        "ModCodeA," & _
        "FoodCodeB," & _
        "ModCodeB," & _
        "Version," & _
        "TypeID," & _
        "Similarity " & _
        "FROM similarity " & _
        "WHERE (FoodCodeA = 0)"
    Set rst2 = New ADODB.Recordset
    rst2.Open SQL, cnnBack, adOpenKeyset, adLockOptimistic, adCmdText

    '--For each food code
    Do Until rst1.EOF
        '--Initialize food code A variables
        lngFoodCodeA = CLng(rst1("FoodCode"))
        lngModCodeA = CLng(rst1("ModCode"))
        lngVersion = CLng(rst1("Version"))
        
        '--Requery food matrix A recordset
        With comFoodMatrixA_Lkp
            .Parameters("@FoodCode") = lngFoodCodeA
            .Parameters("@ModCode") = lngModCodeA
            .Parameters("@Version") = lngVersion
            .Parameters("@WordType") = 1
        End With
        rstFoodMatrixA_Lkp.Requery
        
        '--Get list of word IDs
        strFoodMatrixWordIDs = FoodMatrixWordIDs(rstFoodMatrixA_Lkp)
        '--Calculate the magnitude for food matrix A
        dblMagnitudeA = FoodMatrixMagnitude(rstFoodMatrixA_Lkp)
        
        '--Open recordset with list of food codes containing at least one of the words
        SQL = "SELECT DISTINCT FoodCode," & _
            "ModCode " & _
            "FROM foodword " & _
            "WHERE ((NOT (FoodCode = " & CStr(lngFoodCodeA) & ")) OR " & _
            "(NOT (ModCode = " & CStr(lngModCodeA) & "))) AND " & _
            "(Version = " & CStr(lngVersion) & ") AND " & _
            "(WordID IN (" & strFoodMatrixWordIDs & ")) AND " & _
            "(WordType = 1)"
        Set rst3 = New ADODB.Recordset
        rst3.Open SQL, cnnBack, adOpenStatic, adLockReadOnly, adCmdText
        
        '--For each food code containing at least one of the words
        Do Until rst3.EOF
            '--Initialize food code B variables
            lngFoodCodeB = CLng(rst3("FoodCode"))
            lngModCodeB = CLng(rst3("ModCode"))
            
            dblDotProduct = 0#
            rstFoodMatrixA_Lkp.MoveFirst
            '--For each word in food A
            Do Until rstFoodMatrixA_Lkp.EOF
                '--Initialize word ID
                lngWordID = CLng(rstFoodMatrixA_Lkp("WordID"))
                '--Calculate the dot product
                dblMatrixValueA = CDbl(rstFoodMatrixA_Lkp("tf_idf"))
                dblMatrixValueB = FoodMatrixValue(lngFoodCodeB, lngModCodeB, lngVersion, lngWordID, 1)
                dblDotProduct = dblDotProduct + (dblMatrixValueA * dblMatrixValueB)
                rstFoodMatrixA_Lkp.MoveNext
            Loop
            
            '--Requery food matrix B recordset
            With comFoodMatrixB_Lkp
                .Parameters("@FoodCode") = lngFoodCodeB
                .Parameters("@ModCode") = lngModCodeB
                .Parameters("@Version") = lngVersion
                .Parameters("@WordType") = 1
            End With
            rstFoodMatrixB_Lkp.Requery
            
            '--Calculate the magnitude for food matrix B
            dblMagnitudeB = FoodMatrixMagnitude(rstFoodMatrixB_Lkp)
            
            '--Calculate the similarity between food A and food B
            dblSimilarity = dblDotProduct / (dblMagnitudeA * dblMagnitudeB)
            '--Update the similarity
            Debug.Print lngFoodCodeA, lngFoodCodeB, dblSimilarity
            With rst2
                .AddNew
                .Fields("FoodCodeA") = lngFoodCodeA
                .Fields("ModCodeA") = lngModCodeA
                .Fields("FoodCodeB") = lngFoodCodeB
                .Fields("ModCodeB") = lngModCodeB
                .Fields("Version") = lngVersion
                .Fields("TypeID") = 1
                .Fields("Similarity") = dblSimilarity
                .Update
            End With
            rst3.MoveNext
        Loop
        rst3.Close
        Set rst3 = Nothing
        
        rst1.MoveNext
    Loop
    rst1.Close
    Set rst1 = Nothing
    rst2.Close
    Set rst2 = Nothing

End Sub

Private Sub UpdateFoodSuggestCount(FoodCode As Long, ModCode As Long, Version As Long, SuggestID As Long, SuggestType As Long)
    
    With comSuggestFoodCount_Lkp
        .Parameters("@FoodCode") = FoodCode
        .Parameters("@ModCode") = ModCode
        .Parameters("@Version") = Version
        .Parameters("@SuggestID") = SuggestID
        .Parameters("@SuggestType") = SuggestType
    End With
    With rstSuggestFoodCount_Lkp
        .Requery
        If .RecordCount > 0 Then
            .Fields("SuggestCount") = CLng(.Fields("SuggestCount")) + 1
        Else
            .AddNew
            .Fields("FoodCode") = FoodCode
            .Fields("ModCode") = ModCode
            .Fields("Version") = Version
            .Fields("SuggestID") = SuggestID
            .Fields("SuggestType") = SuggestType
            .Fields("SuggestCount") = 1
        End If
        .Update
    End With

End Sub

Private Sub UpdateIngredSuggestCount(SRCode As String, Version As Long, SuggestID As Long, SuggestType As Long)
    
    With comSuggestIngredCount_Lkp
        .Parameters("@SRCode") = SRCode
        .Parameters("@Version") = Version
        .Parameters("@SuggestID") = SuggestID
        .Parameters("@SuggestType") = SuggestType
    End With
    With rstSuggestIngredCount_Lkp
        .Requery
        If .RecordCount > 0 Then
            .Fields("SuggestCount") = CLng(.Fields("SuggestCount")) + 1
        Else
            .AddNew
            .Fields("SRCode") = SRCode
            .Fields("Version") = Version
            .Fields("SuggestID") = SuggestID
            .Fields("SuggestType") = SuggestType
            .Fields("SuggestCount") = 1
        End If
        .Update
    End With

End Sub

Private Sub UpdateWordCount(FoodCode As Long, ModCode As Long, Version As Long, WordID As Long, WordType As Long)
    
    With comUpdateWordCount
        .Parameters("@FoodCode") = FoodCode
        .Parameters("@ModCode") = ModCode
        .Parameters("@Version") = Version
        .Parameters("@WordID") = WordID
        .Parameters("@WordType") = WordType
    End With
    With rstUpdateWordCount
        .Requery
        If .RecordCount > 0 Then
            .Fields("WordCount") = CLng(.Fields("WordCount")) + 1
        Else
            .AddNew
            .Fields("FoodCode") = FoodCode
            .Fields("ModCode") = ModCode
            .Fields("Version") = Version
            .Fields("WordID") = WordID
            .Fields("WordType") = WordType
            .Fields("WordCount") = 1
        End If
        .Update
    End With

End Sub

Private Function WordCount(FoodCode As Long, ModCode As Long, Version As Long, WordID As Long, WordType As Long) As Long
    
    With comWordCount_Lkp
        .Parameters("@FoodCode") = FoodCode
        .Parameters("@ModCode") = ModCode
        .Parameters("@Version") = Version
        .Parameters("@WordID") = WordID
        .Parameters("@WordType1") = 1
        .Parameters("@WordType2") = WordType
    End With
    With rstWordCount_Lkp
        .Requery
        If .RecordCount > 0 Then
            WordCount = CLng(.Fields("WordCount"))
        End If
    End With

End Function

Private Function WordExists(Word As String) As Long

    comWord_Lkp("@Description") = Word
    With rstWord_Lkp
        .Requery
        If .RecordCount > 0 Then
            WordExists = CLng(.Fields("WordID"))
        End If
    End With

End Function

Private Function WordID() As Long

    With rstWordID_Lkp
        .Requery
        If .RecordCount > 0 Then
            If Not IsNull(.Fields("WordID")) Then
                WordID = CLng(.Fields("WordID"))
            End If
        End If
    End With

End Function

Private Function WordsInDocument(FoodCode As Long, ModCode As Long, Version As Long, WordType As Long) As Long

    With comWordsInDoc_Lkp
        .Parameters("@FoodCode") = FoodCode
        .Parameters("@ModCode") = ModCode
        .Parameters("@Version") = Version
        .Parameters("@WordType1") = 1
        .Parameters("@WordType2") = WordType
    End With
    With rstWordsInDoc_Lkp
        .Requery
        If .RecordCount > 0 Then
            WordsInDocument = .Fields("WordsInDocument")
        End If
    End With

End Function

Public Sub WriteNutrientTooltipMessages()

    Dim SQL As String
    Dim strDescription As String
    Dim strName As String
    Dim strSynonyms As String
    Dim strTables As String
    Dim strTagname As String
    Dim strTitle As String
    Dim strUnits As String
    Dim fso As Scripting.FileSystemObject
    Dim txt1 As Scripting.TextStream
    Dim txt2 As Scripting.TextStream
    Dim rst As ADODB.Recordset
    
    Set fso = New Scripting.FileSystemObject
    Set txt1 = fso.CreateTextFile("E:\projects\fand\databases\rawdata\INFOODS\tagnames\currentNutrients.txt", True)
    Set txt2 = fso.CreateTextFile("E:\projects\fand\databases\rawdata\INFOODS\tagnames\tooltipsNutrients.xhtml", True)
    
    SQL = "SELECT DISTINCT tagname.Tagname, nutrientdescr.NutrientDescription, tagname.TagnameDescription, tagname.Units, tagname.Tables, tagname.Synonyms " & _
        "FROM nutrientdescr INNER JOIN tagname ON nutrientdescr.Tagname = tagname.Tagname " & _
        "WHERE (nutrientdescr.Version = 8) " & _
        "ORDER BY tagname.Tagname"
    Set rst = New ADODB.Recordset
    Call rst.Open(SQL, cnnBack, adOpenStatic, adLockReadOnly, adCmdText)
    
    With rst
        Do Until .EOF
            strTagname = Trim$(.Fields("Tagname"))
            strName = Trim$(.Fields("NutrientDescription"))
            strDescription = Trim$(.Fields("TagnameDescription"))
            strUnits = Trim$(.Fields("Units"))
            strTitle = strName & " (" & strUnits & ")"
            strTables = vbNullString
            If Not IsNull(.Fields("Tables")) Then
                strTables = Trim$(.Fields("Tables"))
            End If
            strSynonyms = vbNullString
            If Not IsNull(.Fields("Synonyms")) Then
                strSynonyms = Trim$(.Fields("Synonyms"))
            End If
            Call txt1.WriteLine("nutrient." & strTagname & ".column.name=" & strName)
            Call txt1.WriteLine("nutrient." & strTagname & ".column.title=" & strTitle)
'            Call txt1.WriteLine("nutrient." & strTagname & ".tooltip.body.tagname=" & strTagname)
            Call txt1.WriteLine("nutrient." & strTagname & ".tooltip.body.description=" & strDescription)
            Call txt1.WriteLine("nutrient." & strTagname & ".tooltip.body.units=" & strUnits)
            If Len(strTables) > 0 Then
                Call txt1.WriteLine("nutrient." & strTagname & ".tooltip.body.tables=" & strTables)
            End If
            If Len(strSynonyms) > 0 Then
                Call txt1.WriteLine("nutrient." & strTagname & ".tooltip.body.synonyms=" & strSynonyms)
            End If
            Call txt2.WriteLine("<ui:include src=""tooltip.xhtml"">")
            Call txt2.Write(vbTab)
            Call txt2.WriteLine("<ui:param name=""tagname"" value=""" & strTagname & """ />")
            If Len(strSynonyms) > 0 Then
                Call txt2.Write(vbTab)
                Call txt2.WriteLine("<ui:param name=""hasSynonyms"" value=""true"" />")
            End If
            If Len(strTables) > 0 Then
                Call txt2.Write(vbTab)
                Call txt2.WriteLine("<ui:param name=""hasTables"" value=""true"" />")
            End If
            Call txt2.Write(vbTab)
            Call txt2.WriteLine("<ui:param name=""styleName"" value="""" />")
            Call txt2.WriteLine("</ui:include>")
            .MoveNext
        Loop
    End With
    
    txt2.Close
    Set txt2 = Nothing
    txt1.Close
    Set txt1 = Nothing
    Set fso = Nothing
    
    rst.Close
    Set rst = Nothing

End Sub

Private Sub Class_Terminate()

    Call CloseCommands
    
    If Not (cnnBack Is Nothing) Then
        cnnBack.Close
        Set cnnBack = Nothing
    End If
    If Not (cnnFNDDS Is Nothing) Then
        cnnFNDDS.Close
        Set cnnFNDDS = Nothing
    End If
'    If Not (cnnMPED Is Nothing) Then
'        cnnMPED.Close
'        Set cnnMPED = Nothing
'    End If
    If Not (cnnSR Is Nothing) Then
        cnnSR.Close
        Set cnnSR = Nothing
    End If
    
    Set wstExcel1 = wbkExcel1.Worksheets("Sheet1")
'    Set wstExcel2 = wbkExcel2.Worksheets("IFF")
'    Set wstExcel3 = wbkExcel2.Worksheets("TOT")
    wbkExcel1.Close
    Set wbkExcel1 = Nothing
'    wbkExcel2.Close
'    Set wbkExcel2 = Nothing
    appExcel.Quit
    Set appExcel = Nothing
    
    Set fso = Nothing
    Set Utility = Nothing

End Sub