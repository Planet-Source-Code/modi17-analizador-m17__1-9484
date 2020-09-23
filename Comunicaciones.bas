Attribute VB_Name = "Comunicaciones"
'===============================================
'Todo este codigo fue INTEGRAMENTE realizado por
'Modi17 [Modi17@ciudad.com.ar]
'
'Ante cualquier modificación, por favor mandenme
'un e-mail, :)

'MADE IN ARGENTINA CARAJO!
'===============================================

Global Const TIPO_SERVER = "Server"
Global Const TIPO_CLIENTE = "Cliente"

Global tipo As String
Global Conectado As Boolean
Sub Conectar(puerto As String, Optional IP As String)
Select Case tipo
    Case TIPO_SERVER
        frmprincipal.puerto.LocalPort = puerto
        frmprincipal.puerto.RemoteHost = ""
        frmprincipal.puerto.RemotePort = 0
        frmprincipal.puerto.Listen
    Case TIPO_CLIENTE
        frmprincipal.puerto.LocalPort = 0
        frmprincipal.puerto.RemoteHost = IP
        frmprincipal.puerto.RemotePort = puerto
        frmprincipal.puerto.Connect
End Select
End Sub
Sub Enviar(Texto As String)
With frmprincipal
    If .chkcrlf.Value Then
        .puerto.SendData Texto + vbCrLf
    Else
        .puerto.SendData Texto
    End If
End With
Status "ENVIADO: " + Texto, COLOR_ENVIADO
End Sub
