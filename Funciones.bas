Attribute VB_Name = "Funciones"
'===============================================
'Todo este codigo fue INTEGRAMENTE realizado por
'Modi17 [Modi17@ciudad.com.ar]
'
'Ante cualquier modificación, por favor mandenme
'un e-mail, :)

'MADE IN ARGENTINA CARAJO!
'===============================================
Global Const COLOR_ERROR = vbRed
Global Const COLOR_INTERNO = vbYellow
Global Const COLOR_ENVIADO = vbBlue
Global Const COLOR_RECIBIDO = vbGreen
Sub Status(Texto As String, Color As String)
With frmprincipal.txtestado
    .SelColor = Color
    .SelText = Texto + vbCrLf
    .SelLength = 0
End With
End Sub
