VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -----------------------------------------------------------------------------------------
' CLASS clRange
'
' -----------------------------------------------------------------------------------------
Option Explicit

' tbASchedule
Private mcltbASchedule_CF As Range ' "Cadela"
Private mcltbASchedule_Token As Range ' "senhas"
Private mcltbASchedule_ID As Range ' ID
Private mcltbASchedule_CM As Range ' "Cão"
Private mcltbASchedule_FF As Range ' "Gata"
Private mcltbASchedule_FM As Range ' "Gato"
Private mcltbASchedule_Aux_msg As Range ' "_Mensagem"
Private mcltbASchedule_Conf As Range 'check box
Private mclASchedule_Name As Range ' "Nome / tel"
Private mcltbASchedule_Msg As Range ' "Mensagem"]
Private mcltbASchedule As ListObject ' "tbASchedule from tbASchedule_ID
 
'tbDBTokens
Private mcltbDBTokens As ListObject 'tbDBTokens from range sys_cn_tbDBTokensID 'TODO: change from range ref to class just like mcltbASchedule
Private mcltbDBTokens_ID As Range ' ID
Private mcltbDBTokens_Status As Range ' "Status"
Private mcltbDBTokens_Type As Range ' "Tipo"
Private mcltbDBTokens_FKIDScheduling As Range ' "FK_IDAgendamentos"
Private mcltbDBTokens_FKIDTransfer As Range ' "FK_IDTransferencia"

'tbDBTransfer
Private mcltbDBTransfer_ID As Range ' ID
Private mcltbDBTransfer_FKIDToken As Range ' "FK_IDSenhas"
Private mcltbDBTransfer_FKIDOldSchedule As Range ' "FK_IDAgendamento_Original"
Private mcltbDBTransfer_FKIDNewSchedule As Range ' "FK_IDAgendamento_Novo"
Private mcltbDBTransfer As ListObject 'tbDBTokens from range sys_cn_tbDBTransferID 'TODO: change from range ref to class just like mcltbASchedule

'######################
' Functions
'######################

Public Function vfRngToListCol(bvBodyRange As Range) As ListColumn
' -----------------------------------------------------------------------------------------
'Convert DataBodyRange references to ListColumn references.
' -----------------------------------------------------------------------------------------
    Dim sLO As ListObject: Set sLO = bvBodyRange.ListObject
    Set vfRngToListCol = sLO.ListColumns(bvBodyRange.Cells(1, 1).Offset(-1, 0).Value)

End Function

Public Function vfRngToColRef(bvBodyRange As Range) As Integer
' -----------------------------------------------------------------------------------------
'Convert bodyrange references to in table column reference
' -----------------------------------------------------------------------------------------
End Function

'######################
' tbDBTransfer
'######################
Public Property Get tbDBTransfer() As ListObject 'tbDBTransfer
    If mcltbDBTransfer Is Nothing Then
        On Error Resume Next
            Set mcltbDBTransfer = Range("sys_cn_tbDBTransferID").ListObject
            
        On Error GoTo 0
    End If
    
    Set tbDBTransfer = mcltbDBTransfer
End Property

Public Property Get tbDBTransfer_FKIDNewSchedule() As Range ' "FK_IDAgendamento_Novo"
    If mcltbDBTransfer_FKIDNewSchedule Is Nothing Then
        On Error Resume Next
            Set mcltbDBTransfer_FKIDNewSchedule = Range("sys_cn_tbDBTransferFK_IDNewSchedule")
            
        On Error GoTo 0
    End If
    
    Set tbDBTransfer_FKIDNewSchedule = mcltbDBTransfer_FKIDNewSchedule
End Property

Public Property Get tbDBTransfer_FKIDOldSchedule() As Range ' "FK_IDAgendamento_Original"
    If mcltbDBTransfer_FKIDOldSchedule Is Nothing Then
        On Error Resume Next
            Set mcltbDBTransfer_FKIDOldSchedule = Range("sys_cn_tbDBTransferFK_IDOldSchedule")
            
        On Error GoTo 0
    End If
    
    Set tbDBTransfer_FKIDOldSchedule = mcltbDBTransfer_FKIDOldSchedule
End Property

Public Property Get tbDBTransfer_FKIDToken() As Range ' "FK_IDSenhas"
    If mcltbDBTransfer_FKIDToken Is Nothing Then
        On Error Resume Next
            Set mcltbDBTransfer_FKIDToken = Range("sys_cn_tbDBTransferFK_IDToken")
            
        On Error GoTo 0
    End If
    
    Set tbDBTransfer_FKIDToken = mcltbDBTransfer_FKIDToken
End Property

Public Property Get tbDBTransfer_ID() As Range ' ID
    If mcltbDBTransfer_ID Is Nothing Then
        On Error Resume Next
            Set mcltbDBTransfer_ID = Range("sys_cn_tbDBTransferID")
            
        On Error GoTo 0
    End If
    
    Set tbDBTransfer_ID = mcltbDBTransfer_ID
End Property

'######################
' tbASchedule
'######################
Public Property Get tbASchedule() As ListObject ' "tbASchedule from tbASchedule_ID
    If mcltbASchedule Is Nothing Then
        On Error Resume Next
            Set mcltbASchedule = Me.tbASchedule_ID.ListObject
            
        On Error GoTo 0
    End If
    
    Set tbASchedule = mcltbASchedule
End Property

Public Property Get tbDBTokens_FKIDTransfer() As Range ' "FK_IDTransferencia"
    If mcltbDBTokens_FKIDTransfer Is Nothing Then
        On Error Resume Next
            Set mcltbDBTokens_FKIDTransfer = Range("sys_cn_tbDBTokensFK_IDTransfer")
            
        On Error GoTo 0
    End If
    
    Set tbDBTokens_FKIDTransfer = mcltbDBTokens_FKIDTransfer
End Property

Public Property Get tbASchedule_Msg() As Range ' "Mensagem"
    If mcltbASchedule_Msg Is Nothing Then
        On Error Resume Next
            Set mcltbASchedule_Msg = Range("sys_cn_tbAScheduleMsg")
            
        On Error GoTo 0
    End If
    
    Set tbASchedule_Msg = mcltbASchedule_Msg
End Property

Public Property Get ASchedule_Name() As Range ' "Nome / tel"
    If mclASchedule_Name Is Nothing Then
        On Error Resume Next
            Set mclASchedule_Name = Range("sys_cn_tbAScheduleName")
            
        On Error GoTo 0
    End If
    
    Set ASchedule_Name = mclASchedule_Name
End Property

Public Property Get tbASchedule_Conf() As Range 'Check box
    If mcltbASchedule_Conf Is Nothing Then
        On Error Resume Next
            Set mcltbASchedule_Conf = Range("sys_cn_tbAScheduleConf")
            
        On Error GoTo 0
    End If
    
    Set tbASchedule_Conf = mcltbASchedule_Conf
End Property

Public Property Get tbASchedule_Aux_msg() As Range '_Mensagem
    If mcltbASchedule_Aux_msg Is Nothing Then
        On Error Resume Next
            Set mcltbASchedule_Aux_msg = Range("sys_cn_tbASchedule_msg")
            
        On Error GoTo 0
    End If
    
    Set tbASchedule_Aux_msg = mcltbASchedule_Aux_msg
End Property


Public Property Get tbASchedule_Token() As Range 'senhas
    If mcltbASchedule_Token Is Nothing Then
        On Error Resume Next
            Set mcltbASchedule_Token = Range("sys_cn_tbAScheduleToken")
            
        On Error GoTo 0
    End If
    
    Set tbASchedule_Token = mcltbASchedule_Token
End Property

Public Property Get tbASchedule_ID(Optional bvRefresh As Boolean = False) As Range 'ID
    If mcltbASchedule_ID Is Nothing Or bvRefresh Then
        On Error Resume Next
            Set mcltbASchedule_ID = Range("sys_cn_tbAScheduleID")
            
        On Error GoTo 0
    End If
    
    Set tbASchedule_ID = mcltbASchedule_ID
End Property

Public Property Get tbASchedule_CF() As Range 'Cadela
    If mcltbASchedule_CF Is Nothing Then
        On Error Resume Next
            Set mcltbASchedule_CF = Range("sys_cn_tbAScheduleCF")
            
        On Error GoTo 0
    End If
    
    Set tbASchedule_CF = mcltbASchedule_CF
End Property

Public Property Get tbASchedule_CM() As Range 'Cão
    If mcltbASchedule_CM Is Nothing Then
        On Error Resume Next
            Set mcltbASchedule_CM = Range("sys_cn_tbAScheduleCM")
            
        On Error GoTo 0
    End If
    
    Set tbASchedule_CM = mcltbASchedule_CM
End Property

Public Property Get tbASchedule_FF() As Range 'Gata
    If mcltbASchedule_FF Is Nothing Then
        On Error Resume Next
            Set mcltbASchedule_FF = Range("sys_cn_tbAScheduleFF")
            
        On Error GoTo 0
    End If
    
    Set tbASchedule_FF = mcltbASchedule_FF
End Property

Public Property Get tbASchedule_FM() As Range 'Gato Macho
    If mcltbASchedule_FM Is Nothing Then
        On Error Resume Next
            Set mcltbASchedule_FM = Range("sys_cn_tbAScheduleFM")
            
        On Error GoTo 0
    End If
    
    Set tbASchedule_FM = mcltbASchedule_FM
End Property


'######################
' tbDBTokens
'######################

Public Property Get tbDBTokens() As ListObject 'tbDBTokens
    If mcltbDBTokens Is Nothing Then
        On Error Resume Next
            Set mcltbDBTokens = Range("sys_cn_tbDBTokensID").ListObject
        
        On Error GoTo 0
    End If
    
    Set tbDBTokens = mcltbDBTokens
End Property

Public Property Get tbDBTokens_ID() As Range 'ID
    If mcltbDBTokens_ID Is Nothing Then
        On Error Resume Next
            Set mcltbDBTokens_ID = Range("sys_cn_tbDBTokensID")
            'Debug.Print mcltbDBTokens_ID.Address
            
        On Error GoTo 0
    End If
    
    Set tbDBTokens_ID = mcltbDBTokens_ID
End Property

Public Property Get tbDBTokens_Status() As Range 'Status
    If mcltbDBTokens_Status Is Nothing Then
        On Error Resume Next
            Set mcltbDBTokens_Status = Range("sys_cn_tbDBTokensStatus")
            
        On Error GoTo 0
    End If
    
    Set tbDBTokens_Status = mcltbDBTokens_Status
End Property

Public Property Get tbDBTokens_Type() As Range 'Tipo
    If mcltbDBTokens_Type Is Nothing Then
        On Error Resume Next
            Set mcltbDBTokens_Type = Range("sys_cn_tbDBTokensType")
            
        On Error GoTo 0
    End If
    
    Set tbDBTokens_Type = mcltbDBTokens_Type
End Property

Public Property Get tbDBTokens_FKIDScheduling() As Range 'FK_IDAgendamento
    If mcltbDBTokens_FKIDScheduling Is Nothing Then
        On Error Resume Next
            Set mcltbDBTokens_FKIDScheduling = Range("sys_cn_tbDBTokensFK_IDSchedule")
            
        On Error GoTo 0
    End If
    
    Set tbDBTokens_FKIDScheduling = mcltbDBTokens_FKIDScheduling
End Property


