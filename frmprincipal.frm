VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmprincipal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Analizador M17"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4680
   Icon            =   "frmprincipal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dialogo 
      Left            =   2760
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Reloj 
      Interval        =   1
      Left            =   3480
      Top             =   120
   End
   Begin VB.Frame freconectado 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      TabIndex        =   18
      Top             =   2400
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton cmddesconectar 
         Caption         =   "&Desconectar"
         Height          =   495
         Left            =   3120
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox chkcrlf 
         Caption         =   "Con CRLF"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox txtenviar 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   4455
      End
   End
   Begin VB.Frame freconfiguracion 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   0
      TabIndex        =   10
      Top             =   2280
      Width           =   4695
      Begin VB.ComboBox cbotipo 
         Height          =   315
         ItemData        =   "frmprincipal.frx":0442
         Left            =   2010
         List            =   "frmprincipal.frx":044C
         TabIndex        =   1
         Text            =   "Cliente"
         Top             =   120
         Width           =   2535
      End
      Begin VB.Frame frecliente 
         Caption         =   "Cliente"
         Height          =   1335
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   4455
         Begin VB.CommandButton cmdconectar 
            Caption         =   "&Conectar"
            Height          =   495
            Left            =   3120
            TabIndex        =   4
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtip 
            Height          =   285
            Left            =   1200
            TabIndex        =   2
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtpuertocliente 
            Height          =   285
            Left            =   1200
            TabIndex        =   3
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lbldireccionIP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección IP:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   915
         End
         Begin VB.Label lblpuertocliente 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Puerto:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   510
         End
      End
      Begin VB.Frame freserver 
         Caption         =   "Server"
         Height          =   1335
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   4455
         Begin VB.TextBox txtpuertoserver 
            Height          =   285
            Left            =   840
            TabIndex        =   5
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmdcomenzarserver 
            Caption         =   "&Comenzar Server"
            Height          =   495
            Left            =   2760
            TabIndex        =   6
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblpuerto 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Puerto:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   510
         End
      End
      Begin VB.Label lblseleccion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione como actuar:"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   120
         Width           =   1770
      End
   End
   Begin MSWinsockLib.Winsock puerto 
      Left            =   4080
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtestado 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3836
      _Version        =   393217
      BackColor       =   -2147483642
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmprincipal.frx":0461
   End
   Begin VB.Label lblestado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   120
      TabIndex        =   17
      Top             =   4320
      Width           =   4500
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuBajarATxt 
         Caption         =   "Bajar informacion a txt"
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnuNulo 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ayuda"
      Begin VB.Menu mnuQueEsEstePrograma 
         Caption         =   "Que es este programa?"
      End
      Begin VB.Menu mnuRealizadoPor 
         Caption         =   "Realizado por..."
      End
   End
End
Attribute VB_Name = "frmprincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================
'Todo este codigo fue INTEGRAMENTE realizado por
'Modi17 [Modi17@ciudad.com.ar]
'
'Ante cualquier modificación, por favor mandenme
'un e-mail, :)
'MADE IN ARGENTINA CARAJO!
'===============================================

Private Sub cbotipo_Click()
txtpuertocliente.Text = ""
txtpuertoserver.Text = ""
txtip.Text = ""
If cbotipo.Text = "Cliente" Then
    freserver.Visible = False
    frecliente.Visible = True
    tipo = TIPO_CLIENTE
    Status "Modo = Cliente", COLOR_INTERNO
ElseIf cbotipo.Text = "Server" Then
    freserver.Visible = True
    frecliente.Visible = False
    tipo = TIPO_SERVER
    Status "Modo = Server", COLOR_INTERNO
End If
End Sub
Private Sub cmdcomenzarserver_Click()
Conectar txtpuertoserver.Text
Status "Server Iniciado", COLOR_INTERNO
End Sub
Private Sub cmdconectar_Click()
Conectar txtpuertocliente.Text, txtip.Text
End Sub
Private Sub cmddesconectar_Click()
puerto.Close
Status "Desconectado", COLOR_INTERNO
End Sub
Private Sub Form_Load()
tipo = TIPO_CLIENTE
Status "=========================", vbBlue
'Status "=========================", vbBlue
Status "=========================", vbWhite
'Status "=========================", vbWhite
Status "=========================", vbBlue
'Status "=========================", vbBlue
Status "", vbWhite
Status "MADE IN ARGENTINA", vbWhite
End Sub

Private Sub mnuBajarATxt_Click()
If Conectado Then
    Status "ERROR: No puede bajar la informacion a un txt cuando esta conectado", COLOR_ERROR
Else
    dialogo.DialogTitle = "Guardar archivo txt como..."
    dialogo.Filter = "Archivos de texto (*.txt)|*.txt"
    dialogo.ShowSave
    If Dir$(dialogo.FileName) = "" Then
        Open dialogo.FileName For Append As #1
            Print #1, txtestado.Text
        Close #1
    Else
        Open dialogo.FileName For Output As #1
            Print #1, txtestado.Text
        Close #1
    End If
End If
End Sub

Private Sub mnuImprimir_Click()
If Conectado Then
    Status "ERROR: No puede imprimir mientras esta conectado", COLOR_ERROR
Else
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.Print "Analizador M17"
    Printer.FontUnderline = False
    Printer.FontBold = False
    Printer.FontSize = 10
    Printer.Print ""
    Printer.Print txtestado.Text
    Printer.Print ""
    Printer.FontBold = True
    Printer.FontSize = 14
    Printer.Print "MADE IN ARGENTINA"
    Printer.EndDoc
End If
End Sub

Private Sub mnuQueEsEstePrograma_Click()
frmqueesesto.Show
End Sub

Private Sub mnuRealizadoPor_Click()
frmrealizadopor.Show
End Sub

Private Sub mnuSalir_Click()
End
End Sub
Private Sub puerto_Connect()
    Status "Conectado con " + CStr(puerto.RemoteHostIP), COLOR_INTERNO
End Sub

Private Sub puerto_ConnectionRequest(ByVal requestID As Long)
If tipo = TIPO_SERVER Then
    If puerto.State <> 0 Then puerto.Close
    puerto.Accept requestID
End If
End Sub

Private Sub puerto_DataArrival(ByVal bytesTotal As Long)
puerto.GetData datos, vbString
Status "RECIBIDO: " + datos, COLOR_RECIBIDO
End Sub

Private Sub puerto_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Status "ERROR: " + Description, COLOR_ERROR
End Sub
Private Sub Reloj_Timer()
Select Case puerto.State
Case 0
    lblestado.Caption = "Puerto cerrado"
    Conectado = False
    freconfiguracion.Visible = True
    freconectado.Visible = False
Case 1
    lblestado.Caption = "Puerto abierto"
Case 2
    lblestado.Caption = "Esperando conexion..."
Case 3
    lblestado.Caption = "Conexion pendiente"
Case 4
    lblestado.Caption = "Resolviendo host"
Case 5
    lblestado.Caption = "Host resuelto"
Case 6
    lblestado.Caption = "Conectando..."
Case 7
    lblestado.Caption = "Conectado"
    freconfiguracion.Visible = False
    freconectado.Visible = True
    Conectado = True
Case 8
    lblestado.Caption = "El equipo esta cerrando la conexion"
    puerto.Close
    Status tipo + " terminado", COLOR_INTERNO
End Select
End Sub
Private Sub txtenviar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Conectado = True Then
                Enviar txtenviar.Text
                txtenviar.Text = ""
        Else
        Status "ERROR: No puede mandar texto si no esta conectado", COLOR_ERROR
        End If
    End If
End Sub


