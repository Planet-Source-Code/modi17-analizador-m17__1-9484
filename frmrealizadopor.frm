VERSION 5.00
Begin VB.Form frmrealizadopor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Realizado por..."
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblversion 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   75
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   2400
      Top             =   840
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   2400
      Top             =   720
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   2400
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Por favor voten por mi codigo en Planet Source Code!!!!!!!!!!"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3600
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modi17@ciudad.com.ar"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Realizado por Modi17"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Analizador M17"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1605
   End
End
Attribute VB_Name = "frmrealizadopor"
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

Private Sub Cmdok_Click()
Unload Me
End Sub


Private Sub Form_Load()
lblversion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub


