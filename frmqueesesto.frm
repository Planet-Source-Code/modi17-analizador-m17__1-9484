VERSION 5.00
Begin VB.Form frmqueesesto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Que es este programa?"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdgracias 
      Caption         =   "Gracias!, ahora entiendo"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmqueesesto.frx":0000
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmqueesesto"
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
Private Sub cmdgracias_Click()
Unload Me
End Sub
