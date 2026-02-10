Attribute VB_Name = "wf_fmTransferTokens"
Sub wf_vsTransferIDsAT_tbASchedule( _
    bvOriginalID As Range, _
    bvNewId As Range, _
    bvCF As Integer, _
    bvCM As Integer, _
    bvFF As Integer, _
    bvFM As Integer, _
    bvFormInstance As fmTransferTokens _
)
' -----------------------------------------------------------------------------------------
' Will transfer the requested IDs
' -----------------------------------------------------------------------------------------

'counters
Dim c As Long, x As Integer, i As Integer, Row As Variant

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
Dim dtArCol As Object: Set dtArCol = CreateObject("Scripting.Dictionary")
    dtArCol.Add "FkSchedule", 4
    dtArCol.Add "Status", 5
    dtArCol.Add "AType", 3
    dtArCol.Add "ID", 1

'###############################

    'Load tbASchedule animal type values
    Dim arAnimalTypeVal(1 To 4) As Integer
        '1 - CF; 2 - CM; 3 - FF; 4 - FM
        arAnimalTypeVal(1) = bvCF
        arAnimalTypeVal(2) = bvCM
        arAnimalTypeVal(3) = bvFF
        arAnimalTypeVal(4) = bvFM
        
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
            If arTbDBTokens(c, dtArCol("FkSchedule")) = bvOriginalID.Value And _
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
    Dim CountID As Long: CountID = Application.WorksheetFunction.Max(clRng.tbDBTransfer_ID) 'update id val without the need to reset clRng.tbDBTokens_ID
    For c = 1 To dtAnimaltype.Count '1 - CF; 2 - CM; 3 - FF; 4 - FM
        If arAnimalTypeVal(c) > 0 Then
        
            x = 0 'count how many Ids to transfer
            For Each Row In arAnimalTypeRow(c)
                If arTbDBTokens(Row, dtArCol("AType")) = dtAnimaltype(c) And arAnimalTypeVal(c) > x Then
                CountID = CountID + 1
                
                   'replace Current Status with Transferred
                   clRng.tbDBTokens_Status.Cells(Row, 1).Value = clStr.Transferred
                   
                    'add transfer data to tbDBTransfer
                    cm_vsAddRowTo_tbDBTransfer _
                        bvID:=CountID, _
                        bvFKID_Token:=CLng(arTbDBTokens(Row, dtArCol("ID"))), _
                        bvFKID_OldSchedule:=bvOriginalID.Value, _
                        bvFKID_NewSchedule:=bvNewId.Value
                        
                   'add to the col FK_IDTranfers the corresponded Transferred ID
                   clRng.tbDBTokens_FKIDTransfer.Cells(Row, 1).Value = CountID
                   
                   'replace the FK_IDSchedule with the new ID
                   clRng.tbDBTokens_FKIDScheduling.Cells(Row, 1).Value = bvNewId.Value
                
                x = x + 1: i = i + 1
                ElseIf arAnimalTypeVal(c) = x Then
                    Exit For
                
                End If
            Next Row

        End If
    Next c
    
    'Update giver at tbASchedule CF... vals
    With bvFormInstance
        cm_vsUpdateATypeValAT_tbASchedule _
             bvTarget:=bvOriginalID, _
             bvCF:=.vOriginalCF - bvCF, _
             bvCM:=.vOriginalCM - bvCM, _
             bvFF:=.vOriginalFF - bvFF, _
             bvFM:=.vOriginalFM - bvFM
    End With
    
    'Update receiver at tbASchedule CF... vals
    
    With clRng.tbASchedule_ID.Worksheet
        cm_vsUpdateATypeValAT_tbASchedule _
             bvTarget:=bvNewId, _
             bvCF:=.Cells(bvNewId.Row, clRng.tbASchedule_CF.Column).Value + bvCF, _
             bvCM:=.Cells(bvNewId.Row, clRng.tbASchedule_CM.Column).Value + bvCM, _
             bvFF:=.Cells(bvNewId.Row, clRng.tbASchedule_FF.Column).Value + bvFF, _
             bvFM:=.Cells(bvNewId.Row, clRng.tbASchedule_FM.Column).Value + bvFM
    End With

End Sub


    
