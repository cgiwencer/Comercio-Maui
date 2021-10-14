VERSION 5.00
Begin VB.Form Cerrado 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2500
      Left            =   0
      Top             =   45
   End
   Begin VB.Image Image3 
      Height          =   870
      Left            =   4005
      Picture         =   "Cerrado.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   825
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   45
      Picture         =   "Cerrado.frx":AA5D
      Stretch         =   -1  'True
      Top             =   2340
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sesión cerrada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   690
      Left            =   45
      TabIndex        =   1
      Top             =   1620
      Width           =   4740
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Acceso autorizado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   690
      Left            =   -405
      TabIndex        =   0
      Top             =   1125
      Width           =   5730
   End
   Begin VB.Image Image2 
      Height          =   3525
      Left            =   -45
      Picture         =   "Cerrado.frx":B59F
      Stretch         =   -1  'True
      Top             =   -45
      Width           =   4920
   End
End
Attribute VB_Name = "Cerrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
If Timer1.Interval = 2500 Then
    Unload Cerrado
    Set Cerrado = Nothing
End If
End Sub
