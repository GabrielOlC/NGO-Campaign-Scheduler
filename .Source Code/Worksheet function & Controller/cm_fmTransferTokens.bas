Attribute VB_Name = "cm_fmTransferTokens"
Option Explicit

Public Sub cm_vsAddRowTo_tbDBTransfer(bvID As Long, bvFKID_Token As Long, bvFKID_OldSchedule As Long, bvFKID_NewSchedule As Long)
' -----------------------------------------------------------------------------------------
' Add a new row to tbDBTranfer and fill values per demand
' -----------------------------------------------------------------------------------------

'import classes
Dim clRng As clRange: Set clRng = New clRange
    
'###############################
    Dim sNewRow As ListRow: Set sNewRow = clRng.tbDBTransfer.ListRows.Add

    With sNewRow
        'add ID
        'Debug.Print Application.WorksheetFunction.Max(clRng.tbDBTransfer_ID) + 1
        .Range(1, clRng.vfRngToListCol(clRng.tbDBTransfer_ID).Index).Value = bvID
        'add FK_IDSenhas
        .Range(1, clRng.vfRngToListCol(clRng.tbDBTransfer_FKIDToken).Index).Value = bvFKID_Token
        'Add FK_IDAgendamento_Original
        .Range(1, clRng.vfRngToListCol(clRng.tbDBTransfer_FKIDOldSchedule).Index).Value = bvFKID_OldSchedule
        'FK_IDAgendamento_Novo
        .Range(1, clRng.vfRngToListCol(clRng.tbDBTransfer_FKIDNewSchedule).Index).Value = bvFKID_NewSchedule
        
    End With

End Sub
Public Function cm_vfEvaluateReceiverID(bvID As Long) As Range
' -----------------------------------------------------------------------------------------
' Run throughg tbASchedule to check if the receiver ID exists. Returns nothing
' if there isn't an ID, return a range if it was found (to use as reference)
' -----------------------------------------------------------------------------------------

'counters
Dim c As Long

'import class
Dim clRng As clRange: Set clRng = New clRange

'array TbASchedule col index
Dim dtArCol As Object: Set dtArCol = CreateObject("Scripting.Dictionary") 'TODO: Update just like the equivalent on "wf_vbASchedule
    dtArCol.Add "ID", 1

'###############################

    'load TbASchedule to array
    Dim arTbASchedule As Variant: If Not clRng.tbASchedule Is Nothing Then arTbASchedule = clRng.tbASchedule.DataBodyRange.Value
    
    'Find respective receive ID
    For c = UBound(arTbASchedule) To LBound(arTbASchedule) Step -1
        If arTbASchedule(c, dtArCol("ID")) = bvID Then
            Set cm_vfEvaluateReceiverID = clRng.tbASchedule.DataBodyRange.Cells(c, dtArCol("ID"))
            
            Exit Function
        End If
        
    Next c
    
    'if nothing is found
    Set cm_vfEvaluateReceiverID = Nothing
End Function
