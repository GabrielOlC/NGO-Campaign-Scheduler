Attribute VB_Name = "cm_vbASchedule"
Option Explicit
Public Sub cm_vsAddRowTo_tbDBTokens(bvType As String, bvFKIDScheduling As Long, bvStatus As String, bvID As Long)
' -----------------------------------------------------------------------------------------
' Add a new row to tbDBTokens and fill values per demand
' -----------------------------------------------------------------------------------------

'import classes
Dim clRng As clRange: Set clRng = New clRange
    
'###############################
    Dim sNewRow As ListRow: Set sNewRow = clRng.tbDBTokens.ListRows.Add
    
    With sNewRow
        'add ID
        'Debug.Print Application.WorksheetFunction.Max(clRng.tbDBTokens_ID) + 1
        .Range(1, clRng.vfRngToListCol(clRng.tbDBTokens_ID).Index).Value = bvID
        'add Tipo
        .Range(1, clRng.vfRngToListCol(clRng.tbDBTokens_Type).Index).Value = bvType
        'add FK_IDAgendaclRngnto
        .Range(1, clRng.vfRngToListCol(clRng.tbDBTokens_FKIDScheduling).Index).Value = bvFKIDScheduling
        'add Status
        .Range(1, clRng.vfRngToListCol(clRng.tbDBTokens_Status).Index).Value = bvStatus
    End With

End Sub

Sub cm_vfCancelTokenOn_tbDBTokens( _
    bvArTable As Variant, _
    bvDtCol As Object, _
    bvArRow As Collection, _
    bvStatusToCancel As String, _
    bvTkCountToRemove As Integer _
)
' -----------------------------------------------------------------------------------------
' Loop through an array to find a specific ID to cancel
' -----------------------------------------------------------------------------------------

'Counters
Dim c As Integer, Row As Variant

'import classes
Dim clRng As clRange: Set clRng = New clRange
Dim clStr As clString: Set clStr = New clString

'###############################
    Dim arCachedTb As Variant: arCachedTb = bvArTable
    
    'loop through the table array for each respective ID (in a collection)
    c = 1
    For Each Row In bvArRow
        If bvArTable(Row, bvDtCol("AType")) = bvStatusToCancel And c <= bvTkCountToRemove Then
            c = c + 1
            clRng.tbDBTokens_Status.Cells(Row, 1).Value = clStr.Canceled

        End If
    Next Row
    
End Sub
Sub cm_vsUpdateATypeValAT_tbASchedule(bvTarget As Range, bvCF As Integer, bvCM As Integer, bvFF As Integer, bvFM As Integer)
' -----------------------------------------------------------------------------------------
' Update CF, CM, FF, FM vals in tbASchedule per demand
' -----------------------------------------------------------------------------------------

'import class
    Dim clRng As clRange: Set clRng = New clRange
    
'###############################
    
    With clRng.tbASchedule_ID.Worksheet
        'Update CF
        If bvCF <> 0 Then .Cells(bvTarget.Row, clRng.tbASchedule_CF.Column).Value = bvCF Else .Cells(bvTarget.Row, clRng.tbASchedule_CF.Column).Value = ""
        
        'Update CM
        If bvCM <> 0 Then .Cells(bvTarget.Row, clRng.tbASchedule_CM.Column).Value = bvCM Else .Cells(bvTarget.Row, clRng.tbASchedule_CM.Column).Value = ""
        
        'Update FF
        If bvFF <> 0 Then .Cells(bvTarget.Row, clRng.tbASchedule_FF.Column).Value = bvFF Else .Cells(bvTarget.Row, clRng.tbASchedule_FF.Column).Value = ""
        
        'Update FM
        If bvFM <> 0 Then .Cells(bvTarget.Row, clRng.tbASchedule_FM.Column).Value = bvFM Else .Cells(bvTarget.Row, clRng.tbASchedule_FM.Column).Value = ""
    End With


End Sub
