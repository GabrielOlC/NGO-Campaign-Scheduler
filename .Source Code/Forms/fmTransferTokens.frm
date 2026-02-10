VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fmTransferTokens 
   Caption         =   "Forms de Transferência de Tokens"
   ClientHeight    =   2772
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4932
   OleObjectBlob   =   "fmTransferTokens.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fmTransferTokens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'cache
Public vOriginalCF As Integer
Public vOriginalCM As Integer
Public vOriginalFF As Integer
Public vOriginalFM As Integer
Public sOldID As Range

'sys
Private DisableEvents As Boolean

'INIT
Private Sub UserForm_Initialize()
    DisableEvents = True
End Sub

'TRIGGERED INIT
Public Sub vsINIT(bvCF As Integer, bvCM As Integer, bvFF As Integer, bvFM As Integer)
    'Load cache
    vOriginalCF = bvCF
    vOriginalCM = bvCM
    vOriginalFF = bvFF
    vOriginalFM = bvFM
    
    'setting spin ranges
    spCF.Max = vOriginalCF: spCF.Min = 0
    spCM.Max = vOriginalCM: spCM.Min = 0
    spFF.Max = vOriginalFF: spFF.Min = 0
    spFM.Max = vOriginalFM: spFM.Min = 0
    
    'Load textbox
    tbCF = vOriginalCF
    tbCM = vOriginalCM
    tbFF = vOriginalFF
    tbFM = vOriginalFM
    
    DisableEvents = False
End Sub

'Close
Private Sub UserForm_Terminate()
    vsResetEnviroment
    
End Sub


'Functions
Private Function vfAddVal(bvSpMax As Integer, bvTbVal As Integer) As Integer
' -----------------------------------------------------------------------------------------
' Add 1 to the given value if inside the allowed range
' -----------------------------------------------------------------------------------------
    Dim v As Integer: v = bvTbVal + 1
    If v > bvSpMax Then vfAddVal = bvTbVal Else vfAddVal = v
    
End Function

Private Function vfSubVal(bvSpMin As Integer, bvTbVal As Integer) As Integer
' -----------------------------------------------------------------------------------------
' Subtract 1 from the given value if inside the allowed range
' -----------------------------------------------------------------------------------------
    Dim v As Integer: v = bvTbVal - 1
    If v < bvSpMin Then vfSubVal = bvTbVal Else vfSubVal = v
    
End Function

Private Function vfCheckChange(bvTbVal As Integer, bvSpMin As Integer, bvSpMax As Integer) As Integer
' -----------------------------------------------------------------------------------------
' Check if the typed value is inside the allowed range
' -----------------------------------------------------------------------------------------
    Select Case bvTbVal
        Case Is < bvSpMin
            vfCheckChange = bvSpMin
            Beep
            
        Case Is > bvSpMax
            vfCheckChange = bvSpMax
            Beep
        
        Case Else
            vfCheckChange = bvTbVal
            
    End Select
End Function

'###############################

Private Sub btOK_Click()
' -----------------------------------------------------------------------------------------
' Validates and Triggers wf_vsTransferIDsAT_tbASchedule
' -----------------------------------------------------------------------------------------
    If Not IsNumeric(tbHeaderTo.Value) Then MsgBox "ID precisa ser um número", vbCritical: Exit Sub
    
    Dim vReceiverID As Range: Set vReceiverID = cm_vfEvaluateReceiverID(CLng(tbHeaderTo.Value))
        If vReceiverID Is Nothing Then MsgBox "ID de transferência não existe na tabela, insira um ID válido", vbCritical: Exit Sub
    
    wf_vsTransferIDsAT_tbASchedule _
        bvOriginalID:=sOldID, _
        bvNewId:=vReceiverID, _
        bvCF:=tbCF, _
        bvCM:=tbCM, _
        bvFF:=tbFF, _
        bvFM:=tbFM, _
        bvFormInstance:=Me
    
    Unload Me
End Sub

Private Sub btCancel_Click()
    Unload Me
End Sub

'# Textbox related functions
'## CF
    Private Sub spCF_SpinUp()
        tbCF.Value = vfAddVal(spCF.Max, tbCF.Value)
    End Sub
    
    Private Sub spCF_SpinDown()
        tbCF.Value = vfSubVal(spCF.Min, tbCF.Value)
    End Sub
    
    Private Sub tbCF_Change()
        If DisableEvents Then Exit Sub
        DisableEvents = True
        
            If IsNumeric(tbCF.Value) Then tbCF.Value = vfCheckChange(tbCF.Value, spCF.Min, spCF.Max) Else tbCF.Value = 0: Beep
        
        DisableEvents = False
    End Sub

'## CM
    Private Sub spCM_SpinUp()
        tbCM.Value = vfAddVal(spCM.Max, tbCM.Value)
    End Sub
    
    Private Sub spCM_SpinDown()
        tbCM.Value = vfSubVal(spCM.Min, tbCM.Value)
    End Sub
    
    Private Sub tbCM_Change()
        If DisableEvents Then Exit Sub
        DisableEvents = True
        
            If IsNumeric(tbCM.Value) Then tbCM.Value = vfCheckChange(tbCM.Value, spCM.Min, spCM.Max) Else tbCM.Value = 0: Beep
        
        DisableEvents = False
    End Sub
    
'## FF
    Private Sub spFF_SpinUp()
        tbFF.Value = vfAddVal(spFF.Max, tbFF.Value)
    End Sub
    
    Private Sub spFF_SpinDown()
        tbFF.Value = vfSubVal(spFF.Min, tbFF.Value)
    End Sub
    
    Private Sub tbFF_Change()
        If DisableEvents Then Exit Sub
        DisableEvents = True
        
            If IsNumeric(tbFF.Value) Then tbFF.Value = vfCheckChange(tbFF.Value, spFF.Min, spFF.Max) Else tbFF.Value = 0: Beep
        
        DisableEvents = False
    End Sub

'## FM
    Private Sub spFM_SpinUp()
        tbFM.Value = vfAddVal(spFM.Max, tbFM.Value)
    End Sub
    
    Private Sub spFM_SpinDown()
        tbFM.Value = vfSubVal(spFM.Min, tbFM.Value)
    End Sub
    
    Private Sub tbFM_Change()
        If DisableEvents Then Exit Sub
        DisableEvents = True
        
            If IsNumeric(tbFM.Value) Then tbFM.Value = vfCheckChange(tbFM.Value, spFM.Min, spFM.Max) Else tbFM.Value = 0: Beep
        
        DisableEvents = False
    End Sub

