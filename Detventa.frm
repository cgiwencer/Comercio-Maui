VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form Detventa 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6345
   ClientLeft      =   6975
   ClientTop       =   4605
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2520
      Top             =   3420
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=ferreteria"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ferreteria"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from vventa1 where cegr_id = 0"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   420
      Left            =   7740
      TabIndex        =   0
      Top             =   5760
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
      _Version        =   196608
      ForeColor       =   16777215
      BackColor       =   4210816
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Salir"
      BevelWidth      =   1
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid2 
      Bindings        =   "Detventa.frx":0000
      Height          =   4875
      Left            =   135
      TabIndex        =   1
      Top             =   810
      Width           =   9015
      _Version        =   196616
      ForeColorEven   =   0
      BackColorEven   =   14737632
      BackColorOdd    =   14737632
      RowHeight       =   423
      Columns.Count   =   5
      Columns(0).Width=   2805
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "pro_cod"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "pro_cod"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1376
      Columns(1).Caption=   "Cant."
      Columns(1).Name =   "degr_can"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "degr_can"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      Columns(2).Width=   7117
      Columns(2).Caption=   "Descripción"
      Columns(2).Name =   "pro_des"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "pro_des"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1640
      Columns(3).Caption=   "Costo"
      Columns(3).Name =   "degr_pru"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   1
      Columns(3).DataField=   "degr_pru"
      Columns(3).DataType=   5
      Columns(3).NumberFormat=   "####.#0"
      Columns(3).FieldLen=   256
      Columns(4).Width=   1852
      Columns(4).Caption=   "Total"
      Columns(4).Name =   "degr_prt"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   1
      Columns(4).DataField=   "degr_prt"
      Columns(4).DataType=   5
      Columns(4).NumberFormat=   "####.#0"
      Columns(4).FieldLen=   256
      _ExtentX        =   15901
      _ExtentY        =   8599
      _StockProps     =   79
      Caption         =   " "
      BackColor       =   16777215
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo1 
      Height          =   375
      Left            =   90
      TabIndex        =   2
      Top             =   360
      Width           =   1680
      _Version        =   65537
      _ExtentX        =   2963
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   16777215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   8190
      TabIndex        =   9
      Top             =   360
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   3690
      TabIndex        =   8
      Top             =   360
      Width           =   4290
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   1890
      TabIndex        =   7
      Top             =   360
      Width           =   1725
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Num. Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Left            =   8100
      TabIndex        =   6
      Top             =   135
      Width           =   1155
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Left            =   3735
      TabIndex        =   5
      Top             =   135
      Width           =   600
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NIT Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Left            =   1980
      TabIndex        =   4
      Top             =   135
      Width           =   975
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Fact."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   135
      Width           =   1035
   End
End
Attribute VB_Name = "Detventa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SSCommand6_Click()
Arqueo.Enabled = True
Unload Detventa
Set Detventa = Nothing
End Sub
