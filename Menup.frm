VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Menup 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MAUI AND SONS"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17235
   Icon            =   "Menup.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "Menup.frx":CAE1
   ScaleHeight     =   10590
   ScaleWidth      =   17235
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   13095
      Top             =   5265
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowProgressCtls=   0   'False
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   2490
      Left            =   8415
      TabIndex        =   8
      Top             =   1260
      Visible         =   0   'False
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   4392
      _Version        =   196608
      BackStyle       =   1
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Productos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   135
         TabIndex        =   12
         Top             =   1980
         Width           =   2130
      End
      Begin VB.Image Image14 
         Height          =   555
         Left            =   45
         Picture         =   "Menup.frx":195C2
         Stretch         =   -1  'True
         Top             =   1845
         Width           =   2355
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Grupos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   180
         TabIndex        =   11
         Top             =   1395
         Width           =   2130
      End
      Begin VB.Image Image13 
         Height          =   555
         Left            =   45
         Picture         =   "Menup.frx":19C53
         Stretch         =   -1  'True
         Top             =   1260
         Width           =   2355
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Almacenes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   135
         TabIndex        =   10
         Top             =   810
         Width           =   2175
      End
      Begin VB.Image Image12 
         Height          =   555
         Left            =   45
         Picture         =   "Menup.frx":1A2E4
         Stretch         =   -1  'True
         Top             =   675
         Width           =   2355
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Usuarios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   135
         TabIndex        =   9
         Top             =   225
         Width           =   2175
      End
      Begin VB.Image Image11 
         Height          =   555
         Left            =   45
         Picture         =   "Menup.frx":1A975
         Stretch         =   -1  'True
         Top             =   90
         Width           =   2355
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   10140
      Left            =   0
      TabIndex        =   0
      Top             =   1305
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   17886
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   3015
         Top             =   8820
      End
      Begin SSCalendarWidgets_A.SSMonth SSMonth1 
         Height          =   2490
         Left            =   90
         TabIndex        =   21
         Top             =   5490
         Width           =   3345
         _Version        =   65537
         _ExtentX        =   5900
         _ExtentY        =   4392
         _StockProps     =   76
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cerrar Sesión"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1215
         TabIndex        =   36
         Top             =   8820
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Image Image32 
         Height          =   465
         Left            =   450
         Picture         =   "Menup.frx":1B006
         Stretch         =   -1  'True
         Top             =   8730
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reportes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   45
         TabIndex        =   23
         Top             =   4320
         Width           =   3390
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   330
         Left            =   45
         TabIndex        =   19
         Top             =   4995
         Width           =   3345
      End
      Begin VB.Image Image20 
         Height          =   600
         Left            =   0
         Picture         =   "Menup.frx":1C1AD
         Stretch         =   -1  'True
         Top             =   4815
         Width           =   3480
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   585
         Left            =   450
         TabIndex        =   22
         Top             =   8010
         Width           =   2505
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Precios de Venta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   0
         TabIndex        =   7
         Top             =   3735
         Width           =   3435
      End
      Begin VB.Image Image9 
         Height          =   600
         Left            =   0
         Picture         =   "Menup.frx":1C906
         Stretch         =   -1  'True
         Top             =   3555
         Width           =   3480
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Arqueo de Caja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   90
         TabIndex        =   6
         Top             =   3150
         Width           =   3345
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Transferencias"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   0
         TabIndex        =   5
         Top             =   2565
         Width           =   3435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Egreso de Almacenes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   135
         TabIndex        =   4
         Top             =   1980
         Width           =   3210
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso a Almacenes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   90
         TabIndex        =   3
         Top             =   1395
         Width           =   3210
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Clasificadores"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   45
         TabIndex        =   2
         Top             =   765
         Width           =   3390
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1215
         TabIndex        =   1
         Top             =   135
         Width           =   1095
      End
      Begin VB.Image Image6 
         Height          =   600
         Left            =   0
         Picture         =   "Menup.frx":1CF97
         Stretch         =   -1  'True
         Top             =   2970
         Width           =   3480
      End
      Begin VB.Image Image5 
         Height          =   600
         Left            =   0
         Picture         =   "Menup.frx":1D628
         Stretch         =   -1  'True
         Top             =   2385
         Width           =   3480
      End
      Begin VB.Image Image4 
         Height          =   600
         Left            =   0
         Picture         =   "Menup.frx":1DCB9
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   3480
      End
      Begin VB.Image Image3 
         Height          =   600
         Left            =   0
         Picture         =   "Menup.frx":1E34A
         Stretch         =   -1  'True
         Top             =   1215
         Width           =   3480
      End
      Begin VB.Image Image2 
         Height          =   600
         Left            =   0
         Picture         =   "Menup.frx":1E9DB
         Stretch         =   -1  'True
         Top             =   630
         Width           =   3480
      End
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "Menup.frx":1F06C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3480
      End
      Begin VB.Image Image22 
         Height          =   600
         Left            =   0
         Picture         =   "Menup.frx":1F7C5
         Stretch         =   -1  'True
         Top             =   4140
         Width           =   3480
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   1815
      Left            =   5715
      TabIndex        =   13
      Top             =   3915
      Visible         =   0   'False
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   3201
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Devoluciones"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   135
         TabIndex        =   34
         Top             =   1350
         Width           =   2220
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Inv. Inicial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   135
         TabIndex        =   15
         Top             =   225
         Width           =   2220
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Compras"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   135
         TabIndex        =   14
         Top             =   810
         Width           =   2220
      End
      Begin VB.Image Image18 
         Height          =   600
         Left            =   45
         Picture         =   "Menup.frx":1FE56
         Stretch         =   -1  'True
         Top             =   90
         Width           =   2355
      End
      Begin VB.Image Image17 
         Height          =   555
         Left            =   45
         Picture         =   "Menup.frx":204E7
         Stretch         =   -1  'True
         Top             =   675
         Width           =   2355
      End
      Begin VB.Image Image16 
         Height          =   555
         Left            =   45
         Picture         =   "Menup.frx":20B78
         Stretch         =   -1  'True
         Top             =   1215
         Width           =   2355
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   1905
      Left            =   8415
      TabIndex        =   16
      Top             =   5760
      Visible         =   0   'False
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   3360
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Anteriores"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   135
         TabIndex        =   33
         Top             =   855
         Width           =   2175
      End
      Begin VB.Image Image10 
         Height          =   555
         Left            =   45
         Picture         =   "Menup.frx":21209
         Stretch         =   -1  'True
         Top             =   720
         Width           =   2355
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Bajas"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   90
         TabIndex        =   18
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ventas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   135
         TabIndex        =   17
         Top             =   225
         Width           =   2175
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image15 
         Height          =   600
         Left            =   45
         Picture         =   "Menup.frx":2189A
         Stretch         =   -1  'True
         Top             =   90
         Width           =   2355
      End
      Begin VB.Image Image19 
         Height          =   555
         Left            =   45
         Picture         =   "Menup.frx":21F2B
         Stretch         =   -1  'True
         Top             =   1305
         Width           =   2355
      End
   End
   Begin Threed.SSFrame SSFrame5 
      Height          =   1950
      Left            =   8415
      TabIndex        =   24
      Top             =   7695
      Visible         =   0   'False
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   3440
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Inventario Físico"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   315
         TabIndex        =   27
         Top             =   225
         Width           =   1860
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Inventario Valorado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   180
         TabIndex        =   26
         Top             =   810
         Width           =   2085
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Kardex "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   855
         TabIndex        =   25
         Top             =   1395
         Width           =   960
      End
      Begin VB.Image Image26 
         Height          =   555
         Left            =   45
         Picture         =   "Menup.frx":225BC
         Stretch         =   -1  'True
         Top             =   90
         Width           =   2355
      End
      Begin VB.Image Image25 
         Height          =   555
         Left            =   45
         Picture         =   "Menup.frx":22C4D
         Stretch         =   -1  'True
         Top             =   675
         Width           =   2355
      End
      Begin VB.Image Image24 
         Height          =   555
         Left            =   45
         Picture         =   "Menup.frx":232DE
         Stretch         =   -1  'True
         Top             =   1260
         Width           =   2355
      End
   End
   Begin Threed.SSFrame SSFrame6 
      Height          =   1365
      Left            =   10890
      TabIndex        =   30
      Top             =   6255
      Visible         =   0   'False
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   2408
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nueva"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   135
         TabIndex        =   32
         Top             =   225
         Width           =   2130
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Anteriores"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   180
         TabIndex        =   31
         Top             =   810
         Width           =   2085
      End
      Begin VB.Image Image28 
         Height          =   600
         Left            =   45
         Picture         =   "Menup.frx":2396F
         Stretch         =   -1  'True
         Top             =   90
         Width           =   2355
      End
      Begin VB.Image Image29 
         Height          =   555
         Left            =   45
         Picture         =   "Menup.frx":24000
         Stretch         =   -1  'True
         Top             =   675
         Width           =   2355
      End
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rev. 131021"
      Height          =   240
      Left            =   16020
      TabIndex        =   35
      Top             =   10215
      Width           =   1095
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev. 250819"
      Height          =   240
      Left            =   15975
      TabIndex        =   29
      Top             =   11250
      Width           =   1140
   End
   Begin VB.Image Image27 
      Height          =   915
      Left            =   16020
      Picture         =   "Menup.frx":24691
      Stretch         =   -1  'True
      Top             =   9315
      Width           =   1095
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "MAUI AND SONS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   465
      Left            =   14490
      TabIndex        =   28
      Top             =   495
      Width           =   2670
   End
   Begin VB.Image Image23 
      Height          =   915
      Left            =   13185
      Picture         =   "Menup.frx":28D97
      Stretch         =   -1  'True
      Top             =   225
      Width           =   960
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario Inactivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   1350
      TabIndex        =   20
      Top             =   540
      Width           =   4425
   End
   Begin VB.Image Image21 
      Height          =   780
      Left            =   450
      Picture         =   "Menup.frx":306ED
      Stretch         =   -1  'True
      Top             =   225
      Width           =   780
   End
   Begin VB.Image Image8 
      Height          =   10410
      Left            =   3510
      Picture         =   "Menup.frx":32787
      Stretch         =   -1  'True
      Top             =   1305
      Width           =   13740
   End
   Begin VB.Image Image7 
      Height          =   1275
      Left            =   0
      Picture         =   "Menup.frx":3E17F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17250
   End
End
Attribute VB_Name = "Menup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image32_Click()
vusucod = 0
uactivado = 0
vUsuario = ""
vUsuNiv = ""
Image32.Visible = False
Label3.Visible = False
Label20.Caption = "Usuario Inactivo"
Label21.Left = 540
'Image21.Picture = LoadPicture(App.Path & "\imagenes\login.jpg")

End Sub

Private Sub Label10_Click()
Menup.Enabled = False
SSFrame2.Visible = False
Load Usuarios
Usuarios.Show
End Sub

Private Sub Label11_Click()
Menup.Enabled = False
SSFrame2.Visible = False
origen = "almacenes"
Load Almacenes
Almacenes.Show
End Sub

Private Sub Label12_Click()
Menup.Enabled = False
SSFrame2.Visible = False
Load Grupos
Grupos.Show
End Sub

Private Sub Label13_Click()
Menup.Enabled = False
SSFrame2.Visible = False
Load Productos
Productos.Show
End Sub

Private Sub Label14_Click()
Menup.Enabled = False
Load Egreso
Egreso.Show
End Sub

Private Sub Label15_Click()
Menup.Enabled = False
SSFrame3.Visible = False
Load Devoluciones
Devoluciones.Show
End Sub

Private Sub Label16_Click()
If vUsuNiv = "ADMINISTRADOR" Then
    Menup.Enabled = False
    SSFrame3.Visible = False
    Load Ingresos
    Ingresos.Show
Else
    MsgBox "No tiene acceso a este módulo", vbInformation, empresa
End If
End Sub

Private Sub Label17_Click()
If vUsuNiv = "ADMINISTRADOR" Then
    Menup.Enabled = False
    SSFrame3.Visible = False
    Load InvInicial
    InvInicial.Show
Else
    MsgBox "No tiene acceso a este módulo", vbInformation, empresa
End If
End Sub

Private Sub Label18_Click()
Menup.Enabled = False
Load Bajas
Bajas.Show
End Sub

Private Sub Label19_Click()
Unload Menup
Set Menup = Nothing
End Sub

Private Sub Label2_Click()
If uactivado = 1 Then
    If vUsuNiv = "ADMINISTRADOR" Then
        Label2.ForeColor = &HFFFF&
        SSFrame2.Left = 3555
        SSFrame2.Top = 2025
        SSFrame2.Visible = True
        Label4.ForeColor = &H8000000F
        Label5.ForeColor = &H8000000F
        Label6.ForeColor = &H8000000F
        Label7.ForeColor = &H8000000F
        Label8.ForeColor = &H8000000F
'        Label30.ForeColor = &H8000000F
        Label22.ForeColor = &H8000000F
        SSFrame3.Visible = False
        SSFrame4.Visible = False
        SSFrame5.Visible = False
        SSFrame6.Visible = False
    Else
        MsgBox "No tiene acceso a este módulo", vbInformation, empresa
    End If
Else
    MsgBox "Inicie sesion de usuario en el sistema", vbInformation, empresa
    Label20_Click
End If
End Sub

Private Sub Label20_Click()
If Label20.Caption = "Usuario Inactivo" Then
    Menup.Enabled = False
    Load Ingreso
    Ingreso.Show
End If
End Sub

Private Sub Label22_Click()
If uactivado = 1 Then
    If vUsuNiv = "ADMINISTRADOR" Then
        Label22.ForeColor = &HFFFF&
        SSFrame5.Left = 3555
        SSFrame5.Top = 6480
        SSFrame5.Visible = True
        Label2.ForeColor = &H8000000F
        Label4.ForeColor = &H8000000F
        Label5.ForeColor = &H8000000F
        Label6.ForeColor = &H8000000F
        Label7.ForeColor = &H8000000F
        Label8.ForeColor = &H8000000F
'        Label30.ForeColor = &H8000000F
        
        SSFrame2.Visible = False
        SSFrame3.Visible = False
        SSFrame4.Visible = False
        SSFrame6.Visible = False
    Else
        MsgBox "No tiene acceso a este módulo", vbInformation, empresa
    End If
Else
    MsgBox "Inicie sesion de usuario en el sistema", vbInformation, empresa
    Label20_Click
End If
End Sub

Private Sub Label24_Click()
Menup.Enabled = False
Load Kardex
Kardex.Show
End Sub

Private Sub Label25_Click()
'CrystalReport1.ReportFileName = App.Path & "\InvFisVal.rpt"
'CrystalReport1.Action = 1
End Sub

Private Sub Label26_Click()
CrystalReport1.ReportFileName = App.Path & "\InvFisico.rpt"
CrystalReport1.Action = 1
End Sub

Private Sub Label28_Click()
Menup.Enabled = False
Load ProformaAnt
ProformaAnt.Show
End Sub

Private Sub Label29_Click()
Menup.Enabled = False
Load Proforma
Proforma.Show
End Sub

Private Sub Label30_Click()
Label30.ForeColor = &HFFFF&
    Label2.ForeColor = &H8000000F
    Label4.ForeColor = &H8000000F
    Label5.ForeColor = &H8000000F
    Label6.ForeColor = &H8000000F
    Label7.ForeColor = &H8000000F
    Label8.ForeColor = &H8000000F
'    Label9.ForeColor = &H8000000F
    
    SSFrame2.Visible = False
    SSFrame3.Visible = False
    SSFrame4.Visible = False
    SSFrame5.Visible = False
    SSFrame6.Visible = False
    
    Menup.Enabled = False
    Load Utilitarios
    Utilitarios.Show
End Sub

Private Sub Label32_Click()
CrystalReport1.ReportFileName = App.Path & "\etiquetas.rpt"
CrystalReport1.Action = 1
End Sub

Private Sub Label4_Click()
If uactivado = 1 Then
    Label4.ForeColor = &HFFFF&
    SSFrame3.Left = 3510
    SSFrame3.Top = 2655
    SSFrame3.Visible = True
    Label2.ForeColor = &H8000000F
    Label5.ForeColor = &H8000000F
    Label6.ForeColor = &H8000000F
    Label7.ForeColor = &H8000000F
    Label8.ForeColor = &H8000000F
   ' Label30.ForeColor = &H8000000F
    Label22.ForeColor = &H8000000F
    SSFrame2.Visible = False
    SSFrame4.Visible = False
    SSFrame5.Visible = False
    SSFrame6.Visible = False
Else
    MsgBox "Inicie sesion de usuario en el sistema", vbInformation, empresa
    Label20_Click
End If

End Sub
Private Sub Label5_Click()
If uactivado = 1 Then
    Label5.ForeColor = &HFFFF&
    SSFrame4.Left = 3510
    SSFrame4.Top = 3285
    SSFrame4.Visible = True
    Label2.ForeColor = &H8000000F
    Label4.ForeColor = &H8000000F
    Label6.ForeColor = &H8000000F
    Label7.ForeColor = &H8000000F
    Label8.ForeColor = &H8000000F
    'Label30.ForeColor = &H8000000F
    Label22.ForeColor = &H8000000F
    SSFrame2.Visible = False
    SSFrame3.Visible = False
    SSFrame5.Visible = False
    SSFrame6.Visible = False
    
Else
    MsgBox "Inicie sesion de usuario en el sistema", vbInformation, empresa
    Label20_Click
End If

End Sub

Private Sub Label6_Click()
If uactivado = 1 Then
    If vUsuNiv = "ADMINISTRADOR" Then
        Label6.ForeColor = &HFFFF&
        SSFrame6.Left = 3510
        SSFrame6.Top = 3915
        'SSFrame6.Visible = True
        Label2.ForeColor = &H8000000F
        Label4.ForeColor = &H8000000F
        Label5.ForeColor = &H8000000F
        Label7.ForeColor = &H8000000F
        Label8.ForeColor = &H8000000F
       ' Label30.ForeColor = &H8000000F
        Label22.ForeColor = &H8000000F
        SSFrame2.Visible = False
        SSFrame3.Visible = False
        SSFrame4.Visible = False
        SSFrame5.Visible = False
        Load Transferencias
        Transferencias.Show
    Else
        MsgBox "No tiene acceso a este módulo", vbInformation, empresa
    End If
Else
    MsgBox "Inicie sesion de usuario en el sistema", vbInformation, empresa
    Label20_Click
End If

End Sub

Private Sub Label7_Click()
If uactivado = 1 Then
    Label7.ForeColor = &HFFFF&
    Label2.ForeColor = &H8000000F
    Label5.ForeColor = &H8000000F
    Label6.ForeColor = &H8000000F
    Label4.ForeColor = &H8000000F
    Label8.ForeColor = &H8000000F
    'Label30.ForeColor = &H8000000F
    Label22.ForeColor = &H8000000F
    SSFrame2.Visible = False
    SSFrame3.Visible = False
    SSFrame4.Visible = False
    SSFrame5.Visible = False
    SSFrame6.Visible = False
    Menup.Enabled = False
    Load Arqueo
    Arqueo.Show
Else
    MsgBox "Inicie sesion de usuario en el sistema", vbInformation, empresa
    Label20_Click
End If
End Sub

Private Sub Label8_Click()
If uactivado = 1 Then
   If vUsuNiv = "ADMINISTRADOR" Then
        Label8.ForeColor = &HFFFF&
        Label2.ForeColor = &H8000000F
        Label5.ForeColor = &H8000000F
        Label6.ForeColor = &H8000000F
        Label7.ForeColor = &H8000000F
        Label4.ForeColor = &H8000000F
        'Label30.ForeColor = &H8000000F
        Label22.ForeColor = &H8000000F
        SSFrame2.Visible = False
        SSFrame3.Visible = False
        SSFrame4.Visible = False
        SSFrame5.Visible = False
        SSFrame6.Visible = False
        Menup.Enabled = False
        Load Precios
        Precios.Show
    Else
        MsgBox "No tiene acceso a este módulo", vbInformation, empresa
    End If
Else
    MsgBox "Inicie sesion de usuario en el sistema", vbInformation, empresa
    Label20_Click
End If

End Sub
Private Sub Label9_Click()
Menup.Enabled = False
origen = "menup"
Load EgresosAnt
EgresosAnt.Show
End Sub

Private Sub Timer1_Timer()
Label21.Caption = Time$
End Sub
