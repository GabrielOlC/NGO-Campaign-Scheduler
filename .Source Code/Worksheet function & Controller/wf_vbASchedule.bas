Attribute VB_Name = "wf_vbASchedule"
Option Explicit

Sub wf_vsUpdateTokensAT_tbASchedule(bvTargetID As Range)
' -----------------------------------------------------------------------------------------
' Evaluate the specified target range and determine whether to  create or
' remove a corresponding token entry in tbDBTokens.
' -----------------------------------------------------------------------------------------

'counters
Dim c As Long, x As Integer, i As Integer

'import classes
Dim clRng As clRange: Set clRng = New clRange
Dim clStr As clString: Set clStr = New clString


'dict to relate 1 - CF; 2 - CM; 3 - FF; 4 - FM
Dim dtAnimaltype As Object: Set dtAnimaltype = CreateObject("Scripting.Dictionary")
Dim arAnimalType As Variant: arAnimalType = Array("CF", "CM", "FF", "FM")
    For c = LBound(arAnimalType) To UBound(arAnimalType)
        dtAnimaltype.Add c + 1, arAnimalType(c)
        
    Next c
    
'array tbDBTokens Col index
Dim dtArCol As Object: Set dtArCol = CreateObject("Scripting.Dictionary") 'TODO: Update for a class reference or equivalent
    dtArCol.Add "FkSchedule", 4
    dtArCol.Add "Status", 5
    dtArCol.Add "AType", 3

'###############################

    'Load tbASchedule animal type values
    Dim arAnimalTypeVal(1 To 4) As Integer
        '1 - CF; 2 - CM; 3 - FF; 4 - FM
        With clRng.tbASchedule_CF.Worksheet
            arAnimalTypeVal(1) = .Cells(bvTargetID.Row, clRng.tbASchedule_CF.Column).Value
            arAnimalTypeVal(2) = .Cells(bvTargetID.Row, clRng.tbASchedule_CM.Column).Value
            arAnimalTypeVal(3) = .Cells(bvTargetID.Row, clRng.tbASchedule_FF.Column).Value
            arAnimalTypeVal(4) = .Cells(bvTargetID.Row, clRng.tbASchedule_FM.Column).Value
        End With
        
    'load database (tbDBTokens) to array
    Dim arTbDBTokens As Variant: If Not clRng.tbDBTokens Is Nothing Then arTbDBTokens = clRng.tbDBTokens.DataBodyRange.Value
    
    'create an index collection for the provided target related IDs
    Dim arAnimalTypeRow() As Collection: ReDim arAnimalTypeRow(1 To dtAnimaltype.Count)
        For c = 1 To dtAnimaltype.Count
            '1 - CF; 2 - CM; 3 - FF; 4 - FM
            Set arAnimalTypeRow(c) = New Collection
            
        Next c
    
        'Load index collections with active TbDBTokens rows that have the target ID
        For c = 1 To UBound(arTbDBTokens)
            If arTbDBTokens(c, dtArCol("FkSchedule")) = bvTargetID.Value And _
                (arTbDBTokens(c, dtArCol("Status")) = clStr.Scheduled Or arTbDBTokens(c, dtArCol("Status")) = clStr.Transferred) Then 'TODO: create a string class for names
                
                For x = 1 To dtAnimaltype.Count 'Find which animal type (from dtAnimaltype) corresponds to the current token row
                    If dtAnimaltype(x) = arTbDBTokens(c, dtArCol("AType")) Then
                        arAnimalTypeRow(x).Add c
                        'Debug.Print c
                        
                        Exit For
                    End If
                Next x
                
            End If
        Next c
    
    'evaluate action type for each animal type
    i = 0 'update id val without the need to reset clRng.tbDBTokens_ID
    For c = 1 To dtAnimaltype.Count '1 - CF; 2 - CM; 3 - FF; 4 - FM
        'Debug.Print dtAnimaltype(c) & ": " & arAnimalTypeVal(c) & " - " & arAnimalTypeRow(c).Count
            
        'arAnimalTypeVal counts from tbASchedule and arAnimalTypeRow counts from tbDBTokens matched array
        If arAnimalTypeVal(c) > arAnimalTypeRow(c).Count Then 'If Scheduling table has more required token than database: add new tokens
            For x = 1 To arAnimalTypeVal(c) - arAnimalTypeRow(c).Count
                cm_vsAddRowTo_tbDBTokens _
                    bvType:=dtAnimaltype(c), _
                    bvFKIDScheduling:=bvTargetID.Value, _
                    bvStatus:=clStr.Scheduled, _
                    bvID:=Application.WorksheetFunction.Max(clRng.tbDBTokens_ID) + 1 + i
                    
                i = i + 1
            
            Next x
                
        ElseIf arAnimalTypeVal(c) < arAnimalTypeRow(c).Count Then 'If there are more tokens then the scheduling table: Cancel tokens
            
            cm_vfCancelTokenOn_tbDBTokens _
                bvArTable:=arTbDBTokens, _
                bvDtCol:=dtArCol, _
                bvArRow:=arAnimalTypeRow(c), _
                bvStatusToCancel:=dtAnimaltype(c), _
                bvTkCountToRemove:=arAnimalTypeRow(c).Count - arAnimalTypeVal(c)

            
        End If
    Next c
    
End Sub
Sub vsINITfmTransferToken(bvID As Range, bvName As String, bvCF As Integer, bvCM As Integer, bvFF As Integer, bvFM As Integer)
' -----------------------------------------------------------------------------------------
' initiate and shows fmTransferToken
' -----------------------------------------------------------------------------------------
        
    Dim frm As fmTransferTokens: Set frm = fmTransferTokens
        
    With frm
        Set .sOldID = bvID
        .lbHeaderFrom.Caption = "Transferir de: " & bvName & " (" & bvID.Value & ")"
        .vsINIT _
            bvCF:=bvCF, _
            bvCM:=bvCM, _
            bvFF:=bvFF, _
            bvFM:=bvFM
        .Show vbModeless
        .tbHeaderTo.SetFocus
        
    End With
    
End Sub



