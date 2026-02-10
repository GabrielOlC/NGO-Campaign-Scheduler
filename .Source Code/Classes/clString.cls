VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -----------------------------------------------------------------------------------------
' CLASS clString
'
' -----------------------------------------------------------------------------------------
Option Explicit

'Names' list
Private mclScheduled As String
Private mclTransferred As String
Private mclCanceled As String

'INIT
Private Sub Class_Initialize()
    mclScheduled = "Agendado"
    mclTransferred = "Transferido"
    mclCanceled = "Cancelado"
End Sub


'######################
' Names' list
'######################

Public Property Get Scheduled() As String
    Scheduled = mclScheduled
End Property

Public Property Get Transferred() As String
    Transferred = mclTransferred
End Property

Public Property Get Canceled() As String
    Canceled = mclCanceled
End Property
