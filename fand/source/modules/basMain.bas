Option Compare Database
Option Explicit

Public Sub FoodAndNutrientData()

    Dim FAND As clsFAND
    
    Set FAND = New clsFAND
    With FAND
'        Call .RebuildTables
'        Call .OpenCommands
'        Call .ImportData
'        Call .ExportData(fvnFNDDS1)
'        Call .ExportData(fvnFNDDS2)
'        Call .ExportData(fvnFNDDS3)
'        Call .ExportData(fvnFNDDS4)
'        Call .ExportData(fvnFNDDS5)
'        Call .ExportData(fvnFNDDS6)
'        Call .ExportData(fvnFNDDS7)
'        Call .WriteEquivalentTooltipMessages
'        Call .WriteNutrientTooltipMessages
'        Call .UpdateData
'        Call .CreateAutoCompleteFiles
    End With
    Set FAND = Nothing
    
End Sub