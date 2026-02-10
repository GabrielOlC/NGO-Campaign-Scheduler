VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbASchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
' -----------------------------------------------------------------------------------------
' Handles schedule-sheet double-click actions by processing token operations in t
' he Senhas column (creating an ID when missing and updating tokens) and copying
' confirmation message text to the clipboard from the confirmation column, while
' suppressing Excel’s default edit behavior.”
' -----------------------------------------------------------------------------------------

'import class
Dim clRng As clRange: Set clRng = New clRange

'###############################
    
' -----------------------------------------------------------------------------------------
' Double-click "Senhas" adds an ID to the respective schedule if not present
' and create/cancel tokens based on inline sheet animal type values.
' ----------------------------------------------------------------------------------------
    If Not Intersect(clRng.tbASchedule_Token, Target) Is Nothing Then
    vsSetEnviroment
        
        If Me.Cells(Target.Row, clRng.tbASchedule_ID.Column).Value = "" Then
            Me.Cells(Target.Row, clRng.tbASchedule_ID.Column).Value = Application.WorksheetFunction.Max(clRng.tbASchedule_ID(bvRefresh:=True)) + 1
            
        End If
        
        wf_vsUpdateTokensAT_tbASchedule Me.Cells(Target.Row, clRng.tbASchedule_ID.Column)
    
    Beep
    Cancel = True
    vsResetEnviroment
    End If
    
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
' -----------------------------------------------------------------------------------------
' Open forms for token transfer (update tokens to ensure sycn)
' -----------------------------------------------------------------------------------------

'import class
Dim clRng As clRange: Set clRng = New clRange

'###############################
    If Not Intersect(clRng.tbASchedule_Token, Target) Is Nothing And Target.Count = 1 Then
    vsSetEnviroment
    
        If Me.Cells(Target.Row, clRng.tbASchedule_ID.Column).Value = "" Then
            Me.Cells(Target.Row, clRng.tbASchedule_ID.Column).Value = Application.WorksheetFunction.Max(clRng.tbASchedule_ID(bvRefresh:=True)) + 1
        
        End If
        wf_vsUpdateTokensAT_tbASchedule Me.Cells(Target.Row, clRng.tbASchedule_ID.Column)
        
    'Update the cell values on screen
    vsResetEnviroment
    vsSetEnviroment
    
        vsINITfmTransferToken _
            bvID:=Me.Cells(Target.Row, clRng.tbASchedule_ID.Column), _
            bvName:=Me.Cells(Target.Row, clRng.ASchedule_Name.Column).Value, _
            bvCF:=Me.Cells(Target.Row, clRng.tbASchedule_CF.Column).Value, _
            bvCM:=Me.Cells(Target.Row, clRng.tbASchedule_CM.Column).Value, _
            bvFF:=Me.Cells(Target.Row, clRng.tbASchedule_FF.Column).Value, _
            bvFM:=Me.Cells(Target.Row, clRng.tbASchedule_FM.Column).Value
    
    'Delegated to Forms unload
    'vsResetEnviroment
    Cancel = True
    End If
    
' -----------------------------------------------------------------------------------------
' Double click on "_Mensagem" copy text to clipboard
' -----------------------------------------------------------------------------------------
    If Not Intersect(clRng.tbASchedule_Msg, Target) Is Nothing And Target.Count = 1 Then
        vfCopyToClipboard (Me.Cells(Target.Row, clRng.tbASchedule_Aux_msg.Column).Value)
        
    Beep
    Cancel = True
    End If

End Sub

