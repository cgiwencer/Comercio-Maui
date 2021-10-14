VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Productos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6705
   ClientLeft      =   9390
   ClientTop       =   3975
   ClientWidth     =   13590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   13590
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   12240
      Top             =   2475
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
   Begin Threed.SSFrame ssFrame1 
      Height          =   1680
      Left            =   90
      TabIndex        =   15
      Top             =   4905
      Width           =   13020
      _ExtentX        =   22966
      _ExtentY        =   2963
      _Version        =   196608
      BackStyle       =   1
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   2925
         TabIndex        =   6
         Top             =   1170
         Width           =   825
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   90
         TabIndex        =   0
         Top             =   495
         Width           =   2265
      End
      Begin VB.TextBox Text7 
         Height          =   330
         Left            =   10170
         TabIndex        =   3
         Top             =   450
         Width           =   870
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "Productos.frx":0000
         Left            =   45
         List            =   "Productos.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1170
         Width           =   2805
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "Productos.frx":0004
         Left            =   2430
         List            =   "Productos.frx":000E
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   495
         Width           =   2805
      End
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   11880
         TabIndex        =   13
         Top             =   2880
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "Productos.frx":0028
         Left            =   4320
         List            =   "Productos.frx":002A
         TabIndex        =   10
         Top             =   2610
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Productos.frx":002C
         Left            =   10620
         List            =   "Productos.frx":002E
         TabIndex        =   12
         Top             =   2925
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox Text6 
         Height          =   330
         Left            =   5310
         TabIndex        =   2
         Top             =   450
         Width           =   4785
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   7425
         TabIndex        =   9
         Top             =   2565
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.TextBox Text3 
         Height          =   330
         Left            =   11115
         TabIndex        =   4
         Top             =   450
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Productos.frx":0030
         Left            =   3015
         List            =   "Productos.frx":0032
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   3015
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4365
         TabIndex        =   11
         Top             =   3780
         Width           =   1995
      End
      Begin Threed.SSCommand SSCommand5 
         Height          =   420
         Left            =   6615
         TabIndex        =   16
         Top             =   1035
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
         Caption         =   "&Cancelar"
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand6 
         Height          =   420
         Left            =   5130
         TabIndex        =   7
         Top             =   1035
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   741
         _Version        =   196608
         ForeColor       =   16777215
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Grabar"
         BevelWidth      =   1
      End
      Begin Threed.SSCheck SSCheck1 
         Height          =   330
         Left            =   3915
         TabIndex        =   17
         Top             =   1125
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   582
         _Version        =   196608
         ForeColor       =   128
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Activo"
         Value           =   1
      End
      Begin Threed.SSCheck SSCheck2 
         Height          =   330
         Left            =   8955
         TabIndex        =   34
         Top             =   2970
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   196608
         ForeColor       =   128
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Autogenerar"
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Left            =   810
         TabIndex        =   40
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Talla"
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
         Left            =   2925
         TabIndex        =   39
         Top             =   945
         Width           =   600
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Marca del producto"
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
         Left            =   225
         TabIndex        =   35
         Top             =   945
         Width           =   2355
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Producto"
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
         Left            =   2970
         TabIndex        =   33
         Top             =   270
         Width           =   2355
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Ex. Max."
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
         Left            =   11925
         TabIndex        =   32
         Top             =   2655
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
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
         Left            =   11745
         TabIndex        =   31
         Top             =   225
         Width           =   780
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
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
         Left            =   6975
         TabIndex        =   23
         Top             =   225
         Width           =   1680
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ex. Min."
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
         Left            =   7470
         TabIndex        =   22
         Top             =   2340
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Año"
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
         Left            =   10395
         TabIndex        =   21
         Top             =   225
         Width           =   510
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicación"
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
         Left            =   5175
         TabIndex        =   20
         Top             =   2385
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor"
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
         Left            =   3060
         TabIndex        =   19
         Top             =   2790
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Código Base"
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
         Left            =   4770
         TabIndex        =   18
         Top             =   3555
         Width           =   1365
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2115
      Top             =   2970
      Width           =   2400
      _ExtentX        =   4233
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
      Connect         =   "DSN=comercio"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "comercio"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "cagisa"
      RecordSource    =   "Select * From vproducto ORDER By gru_des"
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
   Begin Threed.SSCommand SSCommand1 
      Height          =   420
      Left            =   12015
      TabIndex        =   8
      Top             =   675
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
      _Version        =   196608
      ForeColor       =   16777215
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Nuevo"
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   420
      Left            =   12015
      TabIndex        =   24
      Top             =   3375
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid1 
      Bindings        =   "Productos.frx":0034
      Height          =   3255
      Left            =   90
      TabIndex        =   25
      Top             =   630
      Width           =   11850
      _Version        =   196616
      AllowUpdate     =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeRow   =   1
      ForeColorEven   =   0
      BackColorEven   =   12632256
      BackColorOdd    =   12632256
      RowHeight       =   423
      Columns.Count   =   6
      Columns(0).Width=   3200
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "pro_cod"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "pro_cod"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3889
      Columns(1).Caption=   "Prenda"
      Columns(1).Name =   "gru_des"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "gru_des"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3995
      Columns(2).Caption=   "Descripción"
      Columns(2).Name =   "pro_Des"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "pro_Des"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1455
      Columns(3).Caption=   "Talla"
      Columns(3).Name =   "ProTLi"
      Columns(3).CaptionAlignment=   0
      Columns(3).DataField=   "ProTLi"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2963
      Columns(4).Caption=   "Color"
      Columns(4).Name =   "pro_tip"
      Columns(4).CaptionAlignment=   0
      Columns(4).DataField=   "pro_tip"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   4392
      Columns(5).Caption=   "Marca"
      Columns(5).Name =   "Mar_des"
      Columns(5).CaptionAlignment=   0
      Columns(5).DataField=   "Mar_des"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      _ExtentX        =   20902
      _ExtentY        =   5741
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
   Begin Threed.SSCommand SSCommand3 
      Height          =   420
      Left            =   12015
      TabIndex        =   26
      Top             =   1215
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
      _Version        =   196608
      ForeColor       =   16777215
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Modificar"
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand4 
      Height          =   420
      Left            =   12015
      TabIndex        =   27
      Top             =   1800
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
      _Version        =   196608
      ForeColor       =   16777215
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Listar"
      BevelWidth      =   1
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   915
      Left            =   90
      TabIndex        =   28
      Top             =   3960
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   1614
      _Version        =   196608
      BackColor       =   14737632
      ClipControls    =   0   'False
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   180
         TabIndex        =   36
         Top             =   315
         Width           =   1950
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   4905
         TabIndex        =   29
         Top             =   315
         Width           =   3480
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Por Descripción"
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
         Left            =   4950
         TabIndex        =   38
         Top             =   90
         Width           =   1545
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Por Código"
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
         Left            =   225
         TabIndex        =   37
         Top             =   90
         Width           =   1545
      End
      Begin VB.Image Image4 
         Height          =   555
         Left            =   2385
         Picture         =   "Productos.frx":0049
         Stretch         =   -1  'True
         Top             =   225
         Width           =   1815
      End
      Begin VB.Image Image2 
         Height          =   555
         Left            =   8640
         Picture         =   "Productos.frx":0D91
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1815
      End
      Begin VB.Image Image3 
         Height          =   555
         Left            =   10935
         Picture         =   "Productos.frx":1AD9
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Image Image1 
      Height          =   60
      Left            =   6705
      Picture         =   "Productos.frx":27E0
      Stretch         =   -1  'True
      Top             =   405
      Width           =   6810
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCTOS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   11700
      TabIndex        =   30
      Top             =   90
      Width           =   1815
   End
End
Attribute VB_Name = "Productos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vpro_id As Integer
Private Sub Combo1_GotFocus()
Dim Cn As New ADODB.Connection
Dim rspr As New ADODB.Recordset   ' Recordset de Proveedores
Cn.ConnectionString = Cadena
Cn.Open

rspr.CursorType = adOpenKeyset
rspr.LockType = adLockOptimistic
rspr.ActiveConnection = Cn
rspr.Source = "Select * from Proveedores Where prv_est = " & 1
rspr.Open

If Not rspr.EOF Then
    Combo1.Clear
    Do While Not rspr.EOF
        Combo1.AddItem rspr!prv_des
        rspr.MoveNext
    Loop
End If
Cn.Close
End Sub
Private Sub Combo1_LostFocus()
SSCommand7_Click
End Sub
Private Sub Combo2_GotFocus()
If Len(Trim(Text6.Text)) > 0 Then
    Dim Cn As New ADODB.Connection
    Dim rsun As New ADODB.Recordset   ' Recordset de Unidad de manejo
    Cn.ConnectionString = Cadena
    Cn.Open
    
    rsun.CursorType = adOpenKeyset
    rsun.LockType = adLockOptimistic
    rsun.ActiveConnection = Cn
    rsun.Source = "Select * from Unidad"
    rsun.Open
    
    If Not rsun.EOF Then
        Combo2.Clear
        Do While Not rsun.EOF
            Combo2.AddItem rsun!Uni_des
            rsun.MoveNext
        Loop
    End If
    Cn.Close
    If Combo2.ListCount > 0 Then
        Combo2.ListIndex = 0
    End If
Else
    MsgBox "Debe ingresar la descripción del producto", vbInformation, empresa
    Text6.SetFocus
End If
End Sub

Private Sub Combo2_LostFocus()
vuni_des = UCase(Combo2.Text)
Dim Cn As New ADODB.Connection
Dim rsun As New ADODB.Recordset   ' Recordset de Unidad de manejo
Cn.ConnectionString = Cadena
Cn.Open

rsun.CursorType = adOpenKeyset
rsun.LockType = adLockOptimistic
rsun.ActiveConnection = Cn
rsun.Source = "Select * from Unidad Where Uni_des = " & "'" & vuni_des & "'"
rsun.Open

If rsun.EOF Then
    If MsgBox("Unidad de manejo no existente. Desea crealo?", vbYesNo, empresa) = vbYes Then
        grabau = "INSERT INTO unidad Set uni_des = " & "'" & vuni_des & "'"
        Cn.Execute grabau
    End If
End If
Cn.Close
Combo2.Text = UCase(Combo2.Text)
End Sub

Private Sub Combo3_GotFocus()
Dim Cn As New ADODB.Connection
Dim rsal As New ADODB.Recordset   ' Recordset de Unidad de almacenes
Cn.ConnectionString = Cadena
Cn.Open

rsal.CursorType = adOpenKeyset
rsal.LockType = adLockOptimistic
rsal.ActiveConnection = Cn
rsal.Source = "Select * from almacen"
rsal.Open

If Not rsal.EOF Then
    Combo3.Clear
    Do While Not rsal.EOF
        Combo3.AddItem rsal!almdes
        rsal.MoveNext
    Loop
End If
Cn.Close
If Combo3.ListCount > 0 Then
    Combo3.ListIndex = 0
End If
End Sub

Private Sub Combo4_GotFocus()
If Len(Trim(Text9.Text)) > 0 Then
    Dim Cn As New ADODB.Connection
    Dim rsgr As New ADODB.Recordset   ' Recordset de Grupos
    Cn.ConnectionString = Cadena
    Cn.Open
    
    rsgr.CursorType = adOpenKeyset
    rsgr.LockType = adLockOptimistic
    rsgr.ActiveConnection = Cn
    rsgr.Source = "Select * from Grupos Where gru_est = " & 1
    rsgr.Open
    
    If Not rsgr.EOF Then
        Combo4.Clear
        Do While Not rsgr.EOF
            Combo4.AddItem rsgr!Gru_des
            rsgr.MoveNext
        Loop
    End If
    Combo4.ListIndex = 0
    Cn.Close
Else
    MsgBox "Debe ingresar el código del producto", vbInformation, empresa
    Text9.SetFocus
End If
End Sub
Private Sub Combo4_LostFocus()
vgru_des = UCase(Combo4.Text)
If Len(Trim(Combo4.Text)) > 0 Then
    Dim Cn As New ADODB.Connection
    Dim rsgr As New ADODB.Recordset   ' Recordset de Grupos
    Cn.ConnectionString = Cadena
    Cn.Open
    
    rsgr.CursorType = adOpenKeyset
    rsgr.LockType = adLockOptimistic
    rsgr.ActiveConnection = Cn
    rsgr.Source = "Select * from Grupos Where Gru_des = " & "'" & vgru_des & "'"
    rsgr.Open
    
    If rsgr.EOF Then
        If MsgBox("Tipo no existente. Desea crealo?", vbYesNo, empresa) = vbYes Then
            grabag = "INSERT INTO grupos Set gru_des = " & "'" & vgru_des & "', gru_est = " & 1
            Cn.Execute grabag
        Else
            Combo4.SetFocus
        End If
    End If
    Cn.Close
    Combo4.Text = UCase(Combo4.Text)
End If
End Sub
Private Sub Combo5_GotFocus()
If Len(Trim(Text3.Text)) > 0 Then
    Dim Cn As New ADODB.Connection
    Dim rsma As New ADODB.Recordset   ' Recordset de Marcas
    Cn.ConnectionString = Cadena
    Cn.Open
    
    rsma.CursorType = adOpenKeyset
    rsma.LockType = adLockOptimistic
    rsma.ActiveConnection = Cn
    rsma.Source = "Select * from Marcas"
    rsma.Open
    
    If Not rsma.EOF Then
        Combo5.Clear
        Do While Not rsma.EOF
            Combo5.AddItem rsma!Mar_Des
            rsma.MoveNext
        Loop
    End If
    Cn.Close
    If Combo5.ListCount > 0 Then
        Combo5.ListIndex = 0
    End If
Else
    MsgBox "Debe ingresar el color del producto", vbInformation, empresa
    Text3.SetFocus
End If
End Sub
Private Sub Combo5_LostFocus()
vmar_des = UCase(Combo5.Text)
Dim Cn As New ADODB.Connection
Dim rsma As New ADODB.Recordset   ' Recordset de Marcas
Cn.ConnectionString = Cadena
Cn.Open

rsma.CursorType = adOpenKeyset
rsma.LockType = adLockOptimistic
rsma.ActiveConnection = Cn
rsma.Source = "Select * from Marcas Where Mar_des = " & "'" & vmar_des & "'"
rsma.Open

If rsma.EOF And Len(Trim(Combo5)) > 0 Then
    If MsgBox("Marca no existente. Desea crealo?", vbYesNo, empresa) = vbYes Then
        grabam = "INSERT INTO Marcas Set mar_des = " & "'" & vmar_des & "'"
        Cn.Execute grabam
    Else
        Combo5.SetFocus
    End If
End If
Cn.Close
Combo5.Text = UCase(Combo5.Text)
End Sub

Private Sub Combo6_GotFocus()
If Len(Trim(Combo5.Text)) > 0 Then
    Dim Cn As New ADODB.Connection
    Dim rsta As New ADODB.Recordset   ' Recordset de Unidad de tallas
    Cn.ConnectionString = Cadena
    Cn.Open
    
    rsta.CursorType = adOpenKeyset
    rsta.LockType = adLockOptimistic
    rsta.ActiveConnection = Cn
    rsta.Source = "Select * from tallaN"
    rsta.Open
    
    If Not rsta.EOF Then
        Combo6.Clear
        Do While Not rsta.EOF
            Combo6.AddItem rsta!tallaN
            rsta.MoveNext
        Loop
    End If
    Cn.Close
    If Combo6.ListCount > 0 Then
        Combo6.ListIndex = 0
    End If
Else
    MsgBox "Debe seleccionar la marca del producto", vbInformation, empresa
    Combo5.SetFocus
End If
End Sub

Private Sub Combo6_LostFocus()
If Len(Trim(Combo6.Text)) > 0 Then
    Combo7.Text = ""
End If
End Sub

Private Sub Combo7_GotFocus()
If Len(Trim(Combo5.Text)) > 0 Then
    Dim Cn As New ADODB.Connection
    Dim rsta As New ADODB.Recordset   ' Recordset de Unidad de tallas
    Cn.ConnectionString = Cadena
    Cn.Open
    
    rsta.CursorType = adOpenKeyset
    rsta.LockType = adLockOptimistic
    rsta.ActiveConnection = Cn
    rsta.Source = "Select * from tallaL"
    rsta.Open
    
    If Not rsta.EOF Then
        Combo7.Clear
        Do While Not rsta.EOF
            Combo7.AddItem rsta!tallaL & ""
            rsta.MoveNext
        Loop
    End If
    Cn.Close
    'If Combo7.ListCount > 0 Then
    '    Combo7.ListIndex = 0
    'End If
Else
    MsgBox "Debe seleccionar la marca del producto", vbInformation, empresa
    Combo5.SetFocus
End If

End Sub
Private Sub Combo7_LostFocus()
If Len(Trim(Combo7.Text)) > 0 Then
    Combo6.Text = ""
End If
End Sub

Private Sub Form_Load()
KeyPreview = True
cuenta
End Sub
Sub ChgEnterToTab(KeyCode As Integer)
If KeyCode = 13 Then
   KeyCode = 0
   SendKeys "{TAB}"
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   SendKeys "{TAB}"
   KeyAscii = 0
End If
End Sub

Private Sub Image2_Click()
If Len(Trim(Text8.Text)) > 0 Then
    Dim db1 As String
    db1 = Text8.Text
    Adodc1.RecordSource = "SELECT * from vproducto WHERE  pro_des LIKE " & "'%" & db1 & "%'"
    Adodc1.Refresh
    Text8.Text = ""
    SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " Productos encontrados"
End If
End Sub

Private Sub Image3_Click()
Text8.Text = ""
Adodc1.RecordSource = "Select * From vproducto ORDER By pro_des"
Adodc1.Refresh
SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " Productos registrados"
End Sub

Private Sub Image4_Click()
If Len(Trim(Text5.Text)) > 0 Then
    Dim db1 As String
    db1 = Text5.Text
    Adodc1.RecordSource = "SELECT * from vproducto WHERE  pro_cod LIKE " & "'%" & db1 & "%'"
    Adodc1.Refresh
    Text5.Text = ""
    SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " Productos encontrados"
End If
End Sub

Private Sub SSCheck2_Click(Value As Integer)
If SSCheck2.Value = -1 Then
    Dim Cn As New ADODB.Connection
    Dim rsac As New ADODB.Recordset
    Cn.ConnectionString = Cadena
    Cn.Open
    rsac.CursorType = adOpenKeyset
    rsac.LockType = adLockOptimistic
    rsac.ActiveConnection = Cn
    rsac.Source = "Select * from autocod"
    rsac.Open
    
    Text4.Text = rsac!autcod + 1
    Text4.Enabled = False
    Cn.Close
Else
    Text4.Enabled = True
    Text4.Text = ""
End If
End Sub

Private Sub SSCommand1_Click()
creatablaP
limpiadatos
modopro = "N"
SSFrame1.Enabled = True
SSCheck1.Value = -1
Text9.SetFocus
End Sub

Private Sub SSCommand2_Click()
If origen = "ingreso" Then
    Ingresos.Enabled = True
ElseIf origen = "InvIni" Then
    InvInicial.Enabled = True
Else
    Menup.Enabled = True
    Menup.Label2.ForeColor = &HE0E0E0
End If
Unload Productos
Set Productos = Nothing
End Sub

Private Sub SSCommand3_Click()
modopro = "M"
SSFrame1.Enabled = True
Text6.SetFocus
End Sub
Private Sub SSCommand4_Click()
CrystalReport1.ReportFileName = App.Path & "\productos.rpt"
CrystalReport1.Action = 1
End Sub
Private Sub SSCommand5_Click()
limpiadatos
End Sub
Private Sub SSCommand6_Click()
If Len(Trim(Text10.Text)) > 0 Then
    Text10.Text = UCase(Text10.Text)
    Dim Cn As New ADODB.Connection
    Dim rsgr As New ADODB.Recordset
    Dim rsprt As New ADODB.Recordset
    Dim rspro As New ADODB.Recordset
    Cn.ConnectionString = Cadena
    Cn.Open
    
    vpro_cod = Text9.Text
    vprotli = Text10.Text
    
    rsgr.CursorType = adOpenKeyset
    rsgr.LockType = adLockOptimistic
    rsgr.ActiveConnection = Cn
    rsgr.Source = "Select * from productos Where pro_cod = " & "'" & vpro_cod & "' AND protli = " & "'" & vprotli & "'"
    rsgr.Open
    
    'If rsgr.EOF Then
        rsgr.Close
        'Id de Grupo
        vgr_des = Combo4.Text
        rsgr.CursorType = adOpenKeyset
        rsgr.LockType = adLockOptimistic
        rsgr.ActiveConnection = Cn
        rsgr.Source = "Select * from grupos Where gru_des = " & "'" & vgr_des & "'"
        rsgr.Open
        
        vgru_id = rsgr!gru_id
        rsgr.Close
        
        'Id Marca
        vmar_des = Combo5.Text
        rsgr.CursorType = adOpenKeyset
        rsgr.LockType = adLockOptimistic
        rsgr.ActiveConnection = Cn
        rsgr.Source = "Select * from marcas Where mar_des = " & "'" & vmar_des & "'"
        rsgr.Open
        
        vmar_id = rsgr!mar_id
        rsgr.Close
        
        vprv_id = 1
        
        vpro_cod = Text9.Text
        vpro_des = Text6.Text
        vpro_uni = Text7.Text
        vpro_tip = Text3.Text
        vprotli = Text10.Text
        
        If SSCheck1.Value = -1 Then
            vpro_est = 1
        Else
            vpro_est = 0
        End If
        
        If modopro = "N" Then
            'Graba en tabla productos
            grabap = "INSERT INTO productos SET pro_des = " & "'" & vpro_des & "', pro_uni = " & "'" & vpro_uni & "', pro_exi = " & 0 & _
            ", pro_tip = " & "'" & vpro_tip & "', prv_id = " & vprv_id & ", gru_id = " & vgru_id & ", pro_est = " & vpro_est & _
            ", mar_id = " & vmar_id & ", proTLi = " & "'" & vprotli & "', pro_cod = " & "'" & vpro_cod & "', usu_id = " & vusucod
            Cn.Execute grabap
            
            'Graba en tabla productos de tienda
            grabap = "INSERT INTO productostie SET pro_des = " & "'" & vpro_des & "', pro_uni = " & "'" & vpro_uni & "', pro_exi = " & 0 & _
            ", pro_tip = " & "'" & vpro_tip & "', prv_id = " & vprv_id & ", gru_id = " & vgru_id & ", pro_est = " & vpro_est & _
            ", mar_id = " & vmar_id & ", proTLi = " & "'" & vprotli & "', pro_cod = " & "'" & vpro_cod & "', usu_id = " & vusucod
            Cn.Execute grabap
            
            rspro.CursorType = adOpenKeyset
            rspro.LockType = adLockOptimistic
            rspro.ActiveConnection = Cn
            rspro.Source = "Select * from productos"
            rspro.Open
              
            If Not rspro.EOF Then
                rspro.MoveLast
                vpro_id = rspro!pro_id
            End If
        ElseIf modopro = "M" Then
            grabap = "UPDATE productos SET pro_des = " & "'" & vpro_des & "', pro_uni = " & "'" & vpro_uni & "', pro_exi = " & 0 & _
            ", pro_tip = " & "'" & vpro_tip & "', prv_id = " & vprv_id & ", gru_id = " & vgru_id & ", pro_est = " & vpro_est & _
            ", mar_id = " & vmar_id & ", proTNu = " & "'" & vprotnu & "', proTLi = " & "'" & vprotli & "', pro_cod = " & "'" & vpro_cod & "' WHERE pro_id = " & vpro_id
            Cn.Execute grabap
            
            grabap = "UPDATE productostie SET pro_des = " & "'" & vpro_des & "', pro_uni = " & "'" & vpro_uni & "', pro_exi = " & 0 & _
            ", pro_tip = " & "'" & vpro_tip & "', prv_id = " & vprv_id & ", gru_id = " & vgru_id & ", pro_est = " & vpro_est & _
            ", mar_id = " & vmar_id & ", proTNu = " & "'" & vprotnu & "', proTLi = " & "'" & vprotli & "', pro_cod = " & "'" & vpro_cod & "' WHERE pro_id = " & vpro_id
            Cn.Execute grabap
            
            'Modifica Inv. Inicial y compras
            grabap = "UPDATE detinvinicial SET pro_des = " & "'" & vpro_des & "', proTLi = " & "'" & vprotli & "', pro_cod = " & "'" & vpro_cod & "', ProCol = " & "'" & vpro_tip & "' WHERE pro_id = " & vpro_id
            Cn.Execute grabap
            
            
        End If
        
        Cn.Close
        MsgBox "Producto registrado", vbInformation, empresa
        cuenta
        limpiadatos
    'Else
    '    MsgBox "Producto ya registrado en la base de datos", vbInformation, empresa
    '    Text9.SetFocus
    'End If
Else
    MsgBox "Debe selecciona la talla", vbInformation, empresa
    Text10.SetFocus
End If
End Sub

Private Sub SSCommand7_Click()
Dim Cn As New ADODB.Connection
Dim rspr As New ADODB.Recordset
Dim rsprt As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open
vpr_des = Combo1.Text
rspr.CursorType = adOpenKeyset
rspr.LockType = adLockOptimistic
rspr.ActiveConnection = Cn
rspr.Source = "Select * from Proveedores Where prv_des = " & "'" & vpr_des & "'"
rspr.Open


rsprt.CursorType = adOpenKeyset
rsprt.LockType = adLockOptimistic
rsprt.ActiveConnection = Cn
rsprt.Source = "Select * from tproveedor" & vusuariot & " Where pr_des = " & "'" & vpr_des & "'"
rsprt.Open

If rsprt.EOF Then
    If Not rspr.EOF Then
        vpr_id = rspr!prv_id
        grabap = "INSERT INTO tproveedor" & vusuariot & " SET pr_id = " & vpr_id & ", pr_des = " & "'" & vpr_des & "'"
        Cn.Execute grabap
'        Adodc2.Refresh
    Else
        MsgBox "Proveedor inexistente, registrelo en el m'odulo de [PROVEEDORES]", vbInformation, empresa
    End If
Else
    MsgBox "Proveedor ya registrado", vbInformation, empresa
End If
Cn.Close
End Sub

Private Sub SSCommand8_Click()
Productos.Enabled = False
origen = "productos"
Load Proveedores
Proveedores.Show

End Sub

Private Sub SSOleDBGrid1_HeadClick(ByVal ColIndex As Integer)
If VarOrder = False Then
  Adodc1.RecordSource = "Select * From vproducto Order By " & _
  SSOleDBGrid1.Columns(ColIndex).DataField & " ASC"
  Adodc1.Refresh
  VarOrder = True
Else
  Adodc1.RecordSource = "Select * From vproducto Order By " & _
  SSOleDBGrid1.Columns(ColIndex).DataField & " DESC"
  Adodc1.Refresh
  VarOrder = False
End If
End Sub
Private Sub SSOleDBGrid1_Click()
Dim Cn As New ADODB.Connection
Dim rsgr As New ADODB.Recordset
Dim rsma As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open
If Adodc1.Recordset.RecordCount > 0 Then
    seleccion = 1
    vpro_id = Adodc1.Recordset.Fields("pro_id")
    Text6.Text = Adodc1.Recordset.Fields("pro_des")
    Text7.Text = Adodc1.Recordset.Fields("pro_uni")
    Text3.Text = Adodc1.Recordset.Fields("pro_tip")
'    Combo6.Text = Adodc1.Recordset.Fields("ProTNu") & ""
    Text10.Text = Adodc1.Recordset.Fields("ProTLi") & ""
    Text9.Text = Adodc1.Recordset.Fields("pro_cod")
    If Adodc1.Recordset.Fields("pro_est") = 1 Then
        SSCheck1.Value = -1
    Else
        SSCheck1.Value = 0
    End If
    
    'Carga Grupo
    vgru_id = Adodc1.Recordset.Fields("gru_id")
    If Len(Trim(vgru_id)) > 0 Then
        rsgr.CursorType = adOpenKeyset
        rsgr.LockType = adLockOptimistic
        rsgr.ActiveConnection = Cn
        rsgr.Source = "Select * from grupos WHERE gru_id = " & vgru_id
        rsgr.Open
        
        If Not rsgr.EOF Then
            Combo4.Text = rsgr!Gru_des
        End If
    End If
    
    'Carga Marca
    vmar_id = Adodc1.Recordset.Fields("mar_id")
    If Len(Trim(vmar_id)) > 0 Then
        rsma.CursorType = adOpenKeyset
        rsma.LockType = adLockOptimistic
        rsma.ActiveConnection = Cn
        rsma.Source = "Select * from marcas WHERE mar_id = " & vmar_id
        rsma.Open
        
        If Not rsma.EOF Then
            Combo5.Text = rsma!Mar_Des
        End If
    End If
    Cn.Close
End If
End Sub
Private Sub SSOleDBGrid2_DblClick()
Adodc2.Recordset.Delete
Adodc2.Refresh
End Sub
Private Sub Text2_GotFocus()
Text2.BackColor = &H80FFFF
End Sub
Private Sub Text2_LostFocus()
Text2.BackColor = &HFFFFFF
Text2.Text = UCase(Text2.Text)
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub Text1_GotFocus()
Text1.BackColor = &H80FFFF
End Sub
Private Sub Text1_LostFocus()
Text1.BackColor = &HFFFFFF
Text1.Text = UCase(Text1.Text)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub Text3_GotFocus()
Text3.BackColor = &H80FFFF
End Sub
Private Sub Text3_LostFocus()
Text3.BackColor = &HFFFFFF
Text3.Text = UCase(Text3.Text)
End Sub
Private Sub Text4_GotFocus()
Text4.BackColor = &H80FFFF
End Sub
Private Sub Text4_LostFocus()
Text4.BackColor = &HFFFFFF
Text4.Text = UCase(Text4.Text)
End Sub
Private Sub Text6_GotFocus()
If Len(Trim(Combo4.Text)) > 0 Then
    Text6.BackColor = &H80FFFF
Else
    MsgBox "Debe seleccionar el tipo de producto", vbInformation, empresa
    Combo4.SetFocus
End If
End Sub
Private Sub Text6_LostFocus()
Text6.BackColor = &HFFFFFF
Text6.Text = UCase(Text6.Text)
End Sub
Private Function creatablaP()
On Error GoTo errort

Dim TaUsu As String
vusuariot = Trim(vusucod)
TaUsu = "CREATE TABLE Tproveedor" & vusuariot & "(" _
& "pr_Id int(6) DEFAULT NULL, " _
& "pr_Des varchar(250) DEFAULT NULL)"

Dim Cn As New ADODB.Connection
Cn.ConnectionString = Cadena
Cn.Open
Cn.Execute TaUsu
Adodc2.RecordSource = "Select * from Tproveedor" & vusuariot
Adodc2.Refresh

errort:
If Err.Number = -2147217900 Then
    borrat = "Drop Table Tproveedor" & vusuariot
    Cn.Execute borrat
    creatablaP
End If
End Function
Private Function limpiadatos()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
Text7.Text = ""
Text9.Text = ""
Text10.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Combo4.Text = ""
Combo5.Text = ""
SSCheck1.Value = 0
SSCheck2.Value = 0
creatablaP
seleccion = 0
SSFrame1.Enabled = False
End Function
Private Function cuenta()
Adodc1.RecordSource = "Select * From vproducto ORDER By gru_des"
Adodc1.Refresh
SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " Productos registrados"
End Function
Private Sub Text7_GotFocus()
If Len(Trim(Text6.Text)) > 0 Then
    Text7.BackColor = &H80FFFF
Else
    MsgBox "Debe ingresar la descripción del producto", vbInformation, empresa
    Text6.SetFocus
End If
End Sub
Private Sub Text7_LostFocus()
Text7.BackColor = &HFFFFFF
End Sub
Private Sub Text9_GotFocus()
Text9.BackColor = &H80FFFF
End Sub
Private Sub Text9_LostFocus()
Text9.BackColor = &HFFFFFF
Text9.Text = UCase(Text9.Text)
End Sub
Private Sub Text10_GotFocus()
Text10.BackColor = &H80FFFF
End Sub
Private Sub Text10_LostFocus()
Text10.BackColor = &HFFFFFF
End Sub

