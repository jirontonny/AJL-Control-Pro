Attribute VB_Name = "modMain"

Option Explicit

Public Const APP_NAME As String = "AJL Control Pro"
Public Const VERSION As String = "1.0"

Public Sub IniciarSistema()

    MsgBox APP_NAME & vbCrLf & _
           "Versión " & VERSION, _
           vbInformation

End Sub