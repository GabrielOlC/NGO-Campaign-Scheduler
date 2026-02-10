Attribute VB_Name = "Buttons"
Option Explicit
Sub bt_vsUpdateAllTokens()
' -----------------------------------------------------------------------------------------
' For each ID on tbASchedule triggers wf_vsUpdateTokensAT_tbASchedule
' -----------------------------------------------------------------------------------------

vsSetEnviroment

    vsUpdateAllTokens
    
vsResetEnviroment

End Sub
