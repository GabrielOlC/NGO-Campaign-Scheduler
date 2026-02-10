Attribute VB_Name = "cm_Buttons"
Option Explicit
Sub vsUpdateAllTokens()
' -----------------------------------------------------------------------------------------
' For each ID on tbASchedule triggers wf_vsUpdateTokensAT_tbASchedule
' -----------------------------------------------------------------------------------------

'Counters
Dim Row As Range

'import classes
Dim clRng As clRange: Set clRng = New clRange

'###############################
    Dim sbC As Long: sbC = clRng.tbASchedule_ID.Count
    For Each Row In clRng.tbASchedule_ID
        Application.StatusBar = "Progresso: " & Round(Row.Row / sbC * 100, 0) & "%"
        
        If Row.Value = "" Then
            Row = Application.WorksheetFunction.Max(clRng.tbASchedule_ID(bvRefresh:=True)) + 1
            
        End If
    
        wf_vsUpdateTokensAT_tbASchedule Row
        
    Next Row

'###############################
Application.StatusBar = ""
Beep
End Sub
