VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form InvInicial 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9465
   ClientLeft      =   10170
   ClientTop       =   3855
   ClientWidth     =   16950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   16950
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame SSFrame3 
      Height          =   1500
      Left            =   14310
      TabIndex        =   32
      Top             =   6840
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   2646
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo1 
         Height          =   330
         Left            =   90
         TabIndex        =   34
         Top             =   630
         Width           =   1950
         _Version        =   65537
         _ExtentX        =   3440
         _ExtentY        =   582
         _StockProps     =   93
         BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllowNullDate   =   -1  'True
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Inv. Inicial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Top             =   315
         Width           =   1770
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2355
      Left            =   15165
      TabIndex        =   24
      Top             =   990
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   4154
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin Threed.SSCommand SSCommand1 
         Height          =   420
         Left            =   135
         TabIndex        =   10
         Top             =   450
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
      Begin Threed.SSCommand SSCommand11 
         Height          =   420
         Left            =   135
         TabIndex        =   25
         Top             =   1395
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
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   6750
      Top             =   2790
      Width           =   2220
      _ExtentX        =   3916
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
      RecordSource    =   "Select * from detInvInicial where InI_id=0 "
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1125
      Top             =   2655
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
      Connect         =   "DSN=comercio"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "comercio"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "cagisa"
      RecordSource    =   "Select * from valming order by InI_Fec"
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
   Begin Threed.SSFrame SSFrame2 
      Height          =   4965
      Left            =   225
      TabIndex        =   11
      Top             =   4320
      Width           =   13380
      _ExtentX        =   23601
      _ExtentY        =   8758
      _Version        =   196608
      BackStyle       =   1
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   8775
         TabIndex        =   3
         Top             =   3825
         Width           =   1905
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3195
         TabIndex        =   2
         Top             =   3825
         Width           =   1185
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   90
         TabIndex        =   1
         Top             =   3825
         Width           =   3030
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   5490
         TabIndex        =   8
         Top             =   4500
         Width           =   1320
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   135
         TabIndex        =   0
         Top             =   540
         Width           =   4335
      End
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   375
         Left            =   6840
         Top             =   5040
         Visible         =   0   'False
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   661
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
         RecordSource    =   "Select * from productos order by Pro_Cod"
         Caption         =   "Adodc4"
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
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   3870
         TabIndex        =   7
         Top             =   4500
         Width           =   1320
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   2520
         TabIndex        =   6
         Top             =   4500
         Width           =   1185
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1350
         TabIndex        =   5
         Top             =   4500
         Width           =   1005
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   540
         TabIndex        =   4
         Top             =   4500
         Width           =   645
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   3375
         Top             =   2295
         Width           =   2130
         _ExtentX        =   3757
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
         RecordSource    =   "Select * from detinvinicial where InI_id = 0"
         Caption         =   "Adodc2"
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
      Begin Threed.SSCommand SSCommand2 
         Height          =   375
         Left            =   7065
         TabIndex        =   9
         Top             =   4455
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
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
         Caption         =   "Registrar"
         BevelWidth      =   1
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid2 
         Bindings        =   "InvInicial.frx":0000
         Height          =   2490
         Left            =   90
         TabIndex        =   12
         ToolTipText     =   "Para elimina. Doble Click sobre el producto "
         Top             =   945
         Width           =   11415
         _Version        =   196616
         BevelColorFrame =   14737632
         BevelColorHighlight=   14737632
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   1
         ForeColorEven   =   0
         BackColorEven   =   14737632
         BackColorOdd    =   14737632
         RowHeight       =   423
         Columns.Count   =   7
         Columns(0).Width=   2566
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "IteCod"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Pro_cod"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1058
         Columns(1).Caption=   "Cant."
         Columns(1).Name =   "InICan"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Ini_Can"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         Columns(2).Width=   7144
         Columns(2).Caption=   "Descripción"
         Columns(2).Name =   "IteDes"
         Columns(2).CaptionAlignment=   0
         Columns(2).DataField=   "Pro_des"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   2672
         Columns(3).Caption=   "Color"
         Columns(3).Name =   "ProCol"
         Columns(3).CaptionAlignment=   0
         Columns(3).DataField=   "ProCol"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   1746
         Columns(4).Caption=   "Talla"
         Columns(4).Name =   "ProTLi"
         Columns(4).CaptionAlignment=   0
         Columns(4).DataField=   "ProTLi"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   1958
         Columns(5).Caption=   "Pr.Venta"
         Columns(5).Name =   "InI_PrV"
         Columns(5).Alignment=   1
         Columns(5).CaptionAlignment=   1
         Columns(5).DataField=   "InI_PrV"
         Columns(5).DataType=   5
         Columns(5).NumberFormat=   "####.#0"
         Columns(5).FieldLen=   256
         Columns(6).Width=   1958
         Columns(6).Caption=   "Pr.Oferta"
         Columns(6).Name =   "Ini_PrO"
         Columns(6).Alignment=   1
         Columns(6).CaptionAlignment=   1
         Columns(6).DataField=   "Ini_PrO"
         Columns(6).DataType=   5
         Columns(6).NumberFormat=   "####.#0"
         Columns(6).FieldLen=   256
         _ExtentX        =   20135
         _ExtentY        =   4392
         _StockProps     =   79
         Caption         =   " "
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
         Height          =   600
         Left            =   11925
         TabIndex        =   30
         Top             =   1890
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1058
         _Version        =   196608
         CaptionStyle    =   1
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
         Caption         =   "&Grabar Registro"
         BevelWidth      =   1
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Talla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3240
         TabIndex        =   37
         Top             =   3600
         Width           =   435
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Precio de Oferta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5445
         TabIndex        =   36
         Top             =   4275
         Width           =   1410
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Almacén Destino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   35
         Top             =   315
         Width           =   1440
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   4455
         TabIndex        =   29
         Top             =   3825
         Width           =   4245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Precio de venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3870
         TabIndex        =   27
         Top             =   4275
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   26
         Top             =   3600
         Width           =   600
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   7965
         TabIndex        =   22
         Top             =   3465
         Width           =   75
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cant."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   810
         TabIndex        =   21
         Top             =   4950
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   765
         TabIndex        =   20
         Top             =   5220
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   8775
         TabIndex        =   19
         Top             =   3600
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Costo Total "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2520
         TabIndex        =   18
         Top             =   4275
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Costo.Unit."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1350
         TabIndex        =   17
         Top             =   4275
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cant."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   585
         TabIndex        =   16
         Top             =   4275
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4500
         TabIndex        =   13
         Top             =   3645
         Width           =   780
      End
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid1 
      Bindings        =   "InvInicial.frx":0015
      Height          =   3165
      Left            =   225
      TabIndex        =   14
      Top             =   990
      Width           =   4665
      _Version        =   196616
      BevelColorFrame =   14737632
      BevelColorHighlight=   14737632
      AllowUpdate     =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      ForeColorEven   =   0
      BackColorEven   =   14737632
      BackColorOdd    =   14737632
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   953
      Columns(0).Caption=   "Num."
      Columns(0).Name =   "InICod"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   1
      Columns(0).DataField=   "InI_id"
      Columns(0).DataType=   3
      Columns(0).FieldLen=   256
      Columns(1).Width=   2011
      Columns(1).Caption=   "Fecha"
      Columns(1).Name =   "InIFec"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "InI_fec"
      Columns(1).DataType=   7
      Columns(1).NumberFormat=   "dd-mm-yyyy"
      Columns(1).FieldLen=   256
      Columns(2).Width=   4260
      Columns(2).Caption=   "Almacén"
      Columns(2).Name =   "AlmDes"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "AlmDes"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   8229
      _ExtentY        =   5583
      _StockProps     =   79
      Caption         =   " "
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid3 
      Bindings        =   "InvInicial.frx":002A
      Height          =   3210
      Left            =   5025
      TabIndex        =   15
      Top             =   990
      Width           =   9960
      _Version        =   196616
      BevelColorFrame =   14737632
      BevelColorHighlight=   14737632
      AllowUpdate     =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      ForeColorEven   =   0
      BackColorEven   =   14737632
      BackColorOdd    =   14737632
      RowHeight       =   423
      Columns.Count   =   7
      Columns(0).Width=   2434
      Columns(0).Caption=   "Codigo"
      Columns(0).Name =   "IteCod"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Pro_cod"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1032
      Columns(1).Caption=   "Can"
      Columns(1).Name =   "InICan"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Ini_Can"
      Columns(1).DataType=   5
      Columns(1).FieldLen=   256
      Columns(2).Width=   5477
      Columns(2).Caption=   "Descripción"
      Columns(2).Name =   "IteDes"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Pro_des"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2646
      Columns(3).Caption=   "Color"
      Columns(3).Name =   "ProCol"
      Columns(3).CaptionAlignment=   0
      Columns(3).DataField=   "ProCol"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1588
      Columns(4).Caption=   "Talla"
      Columns(4).Name =   "ProTLi"
      Columns(4).CaptionAlignment=   0
      Columns(4).DataField=   "ProTLi"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1746
      Columns(5).Caption=   "Pr.Venta"
      Columns(5).Name =   "InI_PrV"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   1
      Columns(5).DataField=   "InI_PrV"
      Columns(5).DataType=   5
      Columns(5).NumberFormat=   "####.#0"
      Columns(5).FieldLen=   256
      Columns(6).Width=   1588
      Columns(6).Caption=   "Pr.Oferta"
      Columns(6).Name =   "Ini_PrO"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   1
      Columns(6).DataField=   "Ini_PrO"
      Columns(6).DataType=   5
      Columns(6).NumberFormat=   "####.#0"
      Columns(6).FieldLen=   256
      _ExtentX        =   17568
      _ExtentY        =   5662
      _StockProps     =   79
      Caption         =   " "
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
   Begin Threed.SSCommand SSCommand6 
      Height          =   420
      Left            =   15300
      TabIndex        =   28
      Top             =   3510
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
   Begin Threed.SSCommand SSCommand5 
      Height          =   420
      Left            =   14670
      TabIndex        =   31
      Top             =   6300
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
      _Version        =   196608
      ForeColor       =   16777215
      BackColor       =   16576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Productos"
      BevelWidth      =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "INVENTARIO INICIAL"
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
      Left            =   13995
      TabIndex        =   23
      Top             =   270
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   60
      Left            =   10485
      Picture         =   "InvInicial.frx":003F
      Stretch         =   -1  'True
      Top             =   585
      Width           =   6540
   End
End
Attribute VB_Name = "InvInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vini_id, vproidm As Integer

Private Sub Combo1_GotFocus()
buscaprod
If siexiste = 1 Then
    Dim Cn As New ADODB.Connection
    Dim rsit As New ADODB.Recordset   ' Recordset de item
    Cn.ConnectionString = Cadena
    Cn.Open
    Text6.Text = UCase(Text6.Text)
    vcodpro = Text6.Text
        
    rsit.CursorType = adOpenKeyset
    rsit.LockType = adLockOptimistic
    rsit.ActiveConnection = Cn
    rsit.Source = "Select * from productos where Pro_Cod = " & "'" & vcodpro & "'"
    rsit.Open
    
    If Not rsit.EOF Then
        siexiste = 0
        Combo1.Clear
        Do While Not rsit.EOF
            Combo1.AddItem rsit!ProTLi
            rsit.MoveNext
        Loop
    Else
        MsgBox "Producto y talla inexistente", vbInformation, empresa
        Text6.SetFocus
    End If
End If
End Sub

Private Sub Combo3_GotFocus()
If Len(Trim(Combo1.Text)) > 0 Then
    buscaprod
    If siexiste = 1 Then
        Dim Cn As New ADODB.Connection
        Dim rsit As New ADODB.Recordset   ' Recordset de item
        Cn.ConnectionString = Cadena
        Cn.Open
        Text6.Text = UCase(Text6.Text)
        vcodpro = Text6.Text
        vprotli = Combo1.Text
        
        rsit.CursorType = adOpenKeyset
        rsit.LockType = adLockOptimistic
        rsit.ActiveConnection = Cn
        rsit.Source = "Select * from productos where Pro_Cod = " & "'" & vcodpro & "' and ProTli = " & "'" & vprotli & "'"
        rsit.Open
        
        If Not rsit.EOF Then
            siexiste = 0
            Combo3.Clear
            Do While Not rsit.EOF
                Combo3.AddItem rsit!Pro_Tip
                rsit.MoveNext
            Loop
        Else
            MsgBox "Producto y talla inexistente", vbInformation, empresa
            Text6.SetFocus
        End If
    End If
Else
MsgBox "Debe selecciionar la Talla", vbInformation, empresa
Combo1.SetFocus
End If
End Sub

Private Sub Form_Load()
KeyPreview = True
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

Private Sub Combo2_LostFocus()
vgitem = Combo2.Text

Dim Cn As New ADODB.Connection
Dim rsit As New ADODB.Recordset   ' Recordset de item
Cn.ConnectionString = Cadena
Cn.Open
    
rsit.CursorType = adOpenKeyset
rsit.LockType = adLockOptimistic
rsit.ActiveConnection = Cn
rsit.Source = "Select * from productos where Pro_Des = " & "'" & vgitem & "'"
rsit.Open

If Not rsit.EOF Then
    Combo3.Text = rsit!Pro_uni
'    Label13.Caption = rsit!IteUnC
  '  Label15.Caption = rsit!IteUnM
    Combo2.Text = UCase(Combo2.Text)
    Combo2.BackColor = &H80000005
Else

End If
End Sub

Private Sub SSCommand11_Click()
If seleccion = 1 Then
    creatabla
    SSFrame2.Enabled = True
    
    modoini = "M"
    
    'Carga cabecera
    Combo2.Text = SSOleDBGrid1.Columns(2).Value
    Combo2.Enabled = False
    'Fin cabecera
    
    'Carga detalle
    Dim Cn As New ADODB.Connection
    Cn.ConnectionString = Cadena
    Cn.Open
    vinicod = SSOleDBGrid1.Columns(0).Value
    cargad = "Insert into tdetinvinicial" & vusuariot & " Select * from detinvinicial where InI_id = " & vinicod
    Cn.Execute cargad
    Adodc2.Refresh
    'FIn detalle
Else
    MsgBox "Debe seleccionar un ingreso de la lista", vbInformation, empresa
End If

End Sub

Private Sub SSCommand10_Click()
If MsgBox("Luego de cerrar el inventario inicial no podrá ingresar mas items bajo este concepto." & vbCrLf & "" _
& "Desea cerrar el inventario inicial..?", vbYesNo, empresa) = vbYes Then

    Dim Cn As New ADODB.Connection
    Cn.ConnectionString = Cadena
    Cn.Open
    
    cerrar = "UPDATE configuracion set ConInI = " & 1
    Cn.Execute cerrar
    Cn.Close
    MsgBox "Inventario Inicial cerrado", vbInformation, empresa
    SSCommand5_Click
End If

End Sub
Private Sub SSCommand1_Click()
limpiadatos
creatabla
SSFrame1.Enabled = False
'Label8.Caption = Date
SSFrame2.Enabled = True
'Combo2.Enabled = True
'Combo2.SetFocus
modoini = "N"
'Combo2.Enabled = True
Dim Cn As New ADODB.Connection
Dim rsal As New ADODB.Recordset   ' Recordset de Almacenes
Cn.ConnectionString = Cadena
Cn.Open

rsal.CursorType = adOpenKeyset
rsal.LockType = adLockOptimistic
rsal.ActiveConnection = Cn
rsal.Source = "Select * from Almacen WHERE AlmEst = " & 1 & " ORDER BY AlmDes"
rsal.Open

If Not rsal.EOF Then
    Combo2.Clear
    Do While Not rsal.EOF
        If valmori <> rsal!almdes Then
            Combo2.AddItem rsal!almdes
        End If
        rsal.MoveNext
    Loop
End If
Combo2.ListIndex = 0
Cn.Close
Text6.SetFocus
End Sub
Private Sub SSCommand2_Click()
Dim Cn As New ADODB.Connection
Dim rsdt As New ADODB.Recordset   ' Recordset de temporal de detalle
Dim rspr As New ADODB.Recordset   ' Recordset de productos
Cn.ConnectionString = Cadena
Cn.Open
If Len(Trim(Text3.Text)) > 0 Then
    If Val(Text4.Text) > Val(Text2.Text) Then
        vpro_cod = Text6.Text
        vprotli = Combo1.Text
        vprocol = Combo3.Text
        
        rspr.CursorType = adOpenKeyset
        rspr.LockType = adLockOptimistic
        rspr.ActiveConnection = Cn
        rspr.Source = "Select * from productos Where Pro_cod = " & "'" & vpro_cod & "' AND ProTli = " & "'" & vprotli & "' AND Pro_Tip = " & "'" & vprocol & "'"
        rspr.Open
        
        If Not rspr.EOF Then
            vpro_id = rspr!pro_id
        Else
            vpro_id = 0
        End If
        
        vpro_des = Label8.Caption
        vini_can = Val(Text1.Text)
        vini_pru = Val(Text2.Text)
        vIni_PrT = Val(Text3.Text)
        vini_prv = Val(Text4.Text)
        vini_pro = Val(Text5.Text)
        vprocol = Combo3.Text
        vInIFeA = Date
        vInIFeA = Format(vInIFeA, "YYYY-MM-dd")
        
        If SSCommand2.Caption = "Registrar" Then
            'Verifica si el item ya existe
            rsdt.CursorType = adOpenKeyset
            rsdt.LockType = adLockOptimistic
            rsdt.ActiveConnection = Cn
            rsdt.Source = "Select * from detinvinicial Where Pro_cod = " & "'" & vpro_cod & "' AND ProTli = " & "'" & vprotli & "' AND Procol = " & "'" & vprocol & "'"
            rsdt.Open
            
            If Not rsdt.EOF Then
                MsgBox "Este producto ya fue registrado en el Inventario Inicial No. " & rsdt!Ini_Id, vbInformation, empresa
            Else
                rsdt.Close
                rsdt.CursorType = adOpenKeyset
                rsdt.LockType = adLockOptimistic
                rsdt.ActiveConnection = Cn
                rsdt.Source = "Select * from Tdetinvinicial" & vusuariot & " Where Pro_cod = " & "'" & vpro_cod & "' AND ProTli = " & "'" & vprotli & "' AND Procol = " & "'" & vprocol & "'"
                rsdt.Open
            
                If rsdt.EOF And SSCommand2.Caption = "Registrar" Then
                    Text6.Enabled = False
                    nuevoe = "Insert into Tdetinvinicial" & vusuariot & " SET Pro_id = " & vpro_id & ", Pro_Cod = " & "'" & vpro_cod & "'" & ", Pro_Des = " & "'" & vpro_des & "', InI_Can = " & vini_can & _
                    ", InI_PrU = " & vini_pru & ", InI_PrT = " & vIni_PrT & ", InI_Prv = " & vini_prv & ", InI_PrO = " & vini_pro & ", ProTLi = " & "'" & vprotli & "', procol = " & "'" & vprocol & "'"
                    Cn.Execute nuevoe
                    Adodc2.Refresh
                    limpiadatos
                    SSFrame2.Enabled = True
                    Text6.Enabled = True
                    Text6.SetFocus
                Else
                    MsgBox "Item ya seleccionado", vbInformation, empresa
                    limpiadatos
                    Text6.SetFocus
                End If
            End If
         ElseIf SSCommand2.Caption = "Modificar" Then
            vpro_cod1 = SSOleDBGrid2.Columns(0).Value
            borrai = "Delete from Tdetinvinicial" & vusuariot & " Where pro_id = " & vproidm
            Cn.Execute borrai
            
            nuevoe = "Insert into Tdetinvinicial" & vusuariot & " SET Pro_id = " & vpro_id & ", Pro_Cod = " & "'" & vpro_cod & "'" & ", Pro_Des = " & "'" & vpro_des & "', InI_Can = " & vini_can & _
            ", InI_PrU = " & vini_pru & ", InI_PrT = " & vIni_PrT & ", InI_Prv = " & vini_prv & ", InI_PrO = " & vini_pro & ", ProTLi = " & "'" & vprotli & "', procol = " & "'" & vprocol & "'"
            Cn.Execute nuevoe
            Adodc2.Refresh
            limpiadatos
            Text6.SetFocus
            SSCommand2.Caption = "Registrar"
         End If
    Else
        MsgBox "El precio de venta debe ser mayor al de compra", vbInformation, empresa
        Text4.SetFocus
    End If
Else
    MsgBox "Debe completar datos", vbInformation, empresa
    Text1.SetFocus
End If

End Sub

Private Sub SSCommand3_Click()
If modoini = "M" Then
    vini_id = SSOleDBGrid1.Columns(0).Value
    Dim Cn As New ADODB.Connection
    Dim rsir As New ADODB.Recordset   ' Recordset de detalle de inv inicial
    Dim rsit As New ADODB.Recordset   ' Recordset de items
    Cn.ConnectionString = Cadena
    Cn.Open
    
    'Disminuye actualiza Productos
    rsir.CursorType = adOpenKeyset
    rsir.LockType = adLockOptimistic
    rsir.ActiveConnection = Cn
    rsir.Source = "Select * from detinvinicial where InI_id = " & vini_id
    rsir.Open

    If Not rsir.EOF Then
        Do While Not rsir.EOF
            vpro_id = rsir!pro_id
            vini_can = rsir!Ini_Can
            actpro = "UPDATE productos SET pro_exi = pro_exi - " & vini_can & " WHERE pro_id = " & vpro_id
            Cn.Execute actpro
            rsir.MoveNext
        Loop
    End If
    
    ''Borra cabecera
    'borrac = "DELETE FROM cabinvinicial where ini_id = " & vini_id
    'Cn.Execute borrac
    
    'Borra detalle
    borrad = "DELETE FROM detinvinicial where ini_id = " & vini_id
    Cn.Execute borrad
    
End If
grabainvini
grabaalmacen
limpiadatos
creatabla
Adodc1.Refresh
SSFrame2.Enabled = False
SSFrame1.Enabled = True
End Sub
Private Sub SSCommand4_Click()
vinicod = SSOleDBGrid1.Columns(0).Value
If MsgBox("Desea eliminar el ingreso a inventario inicial No: " & vinicod & "..?", vbYesNo, empresa) = vbYes Then
    Dim Cn As New ADODB.Connection
    Dim rsir As New ADODB.Recordset   ' Recordset de detalle de inv inicial
    Dim rsit As New ADODB.Recordset   ' Recordset de items
    Cn.ConnectionString = Cadena
    Cn.Open
    
    'Disminuye actualiza Items
    rsir.CursorType = adOpenKeyset
    rsir.LockType = adLockOptimistic
    rsir.ActiveConnection = Cn
    rsir.Source = "Select * from detinvinicial where InICod = " & vinicod
    rsir.Open
    
    If Not rsir.EOF Then
        Do While Not rsir.EOF
            vcantr = rsir!IniCan
            vitecod = rsir!itecod
            rsit.CursorType = adOpenKeyset
            rsit.LockType = adLockOptimistic
            rsit.ActiveConnection = Cn
            rsit.Source = "Select * from Items where IteCod = " & "'" & vitecod & "'"
            rsit.Open
            
            rsit!IteSal = rsit!IteSal - vcantr
            rsit.Update
            
            rsir.MoveNext
            rsit.Close
        Loop
    End If
    'Fin Disminuye Items
      
    'Borra registro de cabecera
    borracab = "DELETE From cabinvinicial Where IniCod = " & vinicod
    Cn.Execute borracab
    'Fin borra registro de cabecera
    
    'Borra registro de detalle
    borradet = "DELETE From detinvinicial Where IniCod = " & vinicod
    Cn.Execute borradet
    'Fin borra registro de detalle
    
    'Borra cabecera ing egr
    borracie = "DELETE from cabcompra where InICod = " & vinicod
    Cn.Execute borracie
    'Fin borra cabecera ing egr
    
    'Borra registro de detalle ingresos egresos
    borradie = "DELETE from detingegr where InICod = " & vinicod
    Cn.Execute borradie
    'Fin borra ing egr
    MsgBox "Registro eliminado", vbInformation, empresa
    Adodc1.Refresh
    Adodc3.Refresh
End If
End Sub

Private Sub SSCommand5_Click()
origen = "InvIni"
InvInicial.Enabled = False
Load Productos
Productos.Show
End Sub

Private Sub SSCommand6_Click()
Menup.Enabled = True
Unload InvInicial
Set InvInicial = Nothing
Menup.Label4.ForeColor = &HE0E0E0
End Sub

Private Sub SSCommand7_Click()
Dim Cn As New ADODB.Connection
Dim rscie As New ADODB.Recordset   ' Recordset de cabecera de ingresos egresos
Dim rsdie As New ADODB.Recordset   ' Recordset de detalle de ingresos egresos
Dim rsca As New ADODB.Recordset   ' Recordset de cabecera de extras
Dim rsit As New ADODB.Recordset   ' Recordset de items
Dim rsir As New ADODB.Recordset   ' Recordset de detalle real de inventario inicial
Cn.ConnectionString = Cadena
Cn.Open
If modoini = "M" Then
    vinicod = Val(Label2.Caption)

    'actualiza codigos en 0
    borrac = "UPDATE Tdetinvinicial" & vusuariot & " SET InICod = " & 0
    Cn.Execute borrac
    
    'Disminuye actualiza Items
    
    rsir.CursorType = adOpenKeyset
    rsir.LockType = adLockOptimistic
    rsir.ActiveConnection = Cn
    rsir.Source = "Select * from detinvinicial where InICod = " & vinicod
    rsir.Open
    
    
    If Not rsir.EOF Then
        Do While Not rsir.EOF
            vcantr = rsir!IniCan
            vitecod = rsir!itecod
            rsit.CursorType = adOpenKeyset
            rsit.LockType = adLockOptimistic
            rsit.ActiveConnection = Cn
            rsit.Source = "Select * from Items where IteCod = " & "'" & vitecod & "'"
            rsit.Open
            
            rsit!IteSal = rsit!IteSal - vcantr
            rsit.Update
            
            rsir.MoveNext
            rsit.Close
        Loop
    End If
    
        
    'Fin Disminuye Items
      
    'Borra registro de cabecera
    borracab = "DELETE From cabinvinicial Where IniCod = " & vinicod
    Cn.Execute borracab
    'Fin borra registro de cabecera
    
    'Borra registro de detalle
    borradet = "DELETE From detinvinicial Where IniCod = " & vinicod
    Cn.Execute borradet
    'Fin borra registro de detalle
    
    'Borra cabecera ing egr
    borracie = "DELETE from cabcompra where InICod = " & vinicod
    Cn.Execute borracie
    'Fin borra cabecera ing egr
    
    'Borra registro de detalle ingresos egresos
    borradie = "DELETE from detingegr where InICod = " & vinicod
    Cn.Execute borradie
    'Fin borra ing egr
    
    
    
End If
            
If modoini = "N" Then
    'Obtiene numero de boleta
    rsca.CursorType = adOpenKeyset
    rsca.LockType = adLockOptimistic
    rsca.ActiveConnection = Cn
    rsca.Source = "Select * from CabInvinicial ORDER BY InI_id"
    rsca.Open
    
    If Not rsca.EOF Then
        rsca.MoveLast
        vini_id = rsca!Ini_Id + 1
    Else
        vini_id = 1
    End If
    rsca.Close
End If
    
'Asigna codigos a items
nobdet = "Update Tdetinvinicial" & vusuariot & " SET InI_id = " & vini_id
Cn.Execute nobdet
'Fin asigna C


'Graba cabecera
rsca.CursorType = adOpenKeyset
rsca.LockType = adLockOptimistic
rsca.ActiveConnection = Cn
rsca.Source = "Select * from cabinvinicial"
rsca.Open

rsca.AddNew
rsca!Ini_Id = vini_id
rsca!InI_Fec = Format(SSDateCombo2.Text, "yyyy-mm-dd")
rsca!Usu_Id = vusucod
rsca.Update
'Fin graba cabecera

'De Temporal detalle a Real detalle
nuevotr = "Insert into detinvinicial Select * from Tdetinvinicial" & vusuariot
Cn.Execute nuevotr
'Fin de temp a real

''Graba en cabecera de ingreso egreso
'rscie.CursorType = adOpenKeyset
'rscie.LockType = adLockOptimistic
'rscie.ActiveConnection = Cn
'rscie.Source = "Select * from cabcompra order By CabcReg"
'rscie.Open

'If Not rscie.EOF Then
'    rscie.MoveLast
'    vinecod = rscie!cabcreg + 1
'Else
'    vinecod = 1
'End If

''vInEFec = Format(Label8.Caption, "yyyy-mm-dd")
'vCabFec = Format(SSDateCombo2.Text, "yyyy-mm-dd")
'nuevoci = "INSERT INTO Cabcompra SET LCoCod = " & 0 & ", CabcReg = " & vinecod & ", Tipo = 'I', CabFec = " & "'" & vCabFec & "', RazSoc = 'INV.INICIAL', UsuRes =  " & 1 & ", IniCod = " & vinicod
'Cn.Execute nuevoci
''Fin

''Graba en detalle de ingreso egreso
'Adodc2.Recordset.MoveFirst
If Not Adodc2.Recordset.EOF Then
'    rsdie.CursorType = adOpenKeyset
'    rsdie.LockType = adLockOptimistic
'    rsdie.ActiveConnection = Cn
'    rsdie.Source = "Select * from detingegr order By InELot"
'    rsdie.Open

'    If Not rscie.EOF Then
'        rsdie.MoveLast
'        vInELot = rsdie!InELot + 1
'    Else
'        vInELot = 1
'    End If
'    Adodc2.Recordset.MoveFirst
   Do While Not Adodc2.Recordset.EOF
'        rsdie.AddNew
'        rsdie!cabcreg = vinecod
'        rsdie!InELot = vInELot
'        rsdie!itecod = Adodc2.Recordset.Fields("IteCod")
'        rsdie!IteDes = Adodc2.Recordset.Fields("IteDes")
'        rsdie!InEIni = Adodc2.Recordset.Fields("InICan")
'        rsdie!InECaI = Adodc2.Recordset.Fields("InICan")
'        'rsdie!InECaI = 0
'        rsdie!InECaE = 0
'        rsdie!InEsal = Adodc2.Recordset.Fields("InICan")
'        rsdie!InEPrU = Adodc2.Recordset.Fields("InIPru")
'        rsdie!InEtot = Adodc2.Recordset.Fields("InIPrT")
'        rsdie!InEEst = 1
'        rsdie!InETiM = "I"
'        rsdie!inicod = vinicod
'        rsdie!InEFeV = Adodc2.Recordset.Fields("InIFev")
'        rsdie!InEPPP = Adodc2.Recordset.Fields("InIPPP")
'        rsdie.Update
        
        'Actualiza saldo en tabla items
        vpro_cod = Adodc2.Recordset.Fields("Pro_Cod")
        vcant = Adodc2.Recordset.Fields("InI_Can")
        vIteUPC = Adodc2.Recordset.Fields("InI_Pru")
        
        salite = "Update productos set Pro_Sal = " & vcant & " where Pro_Cod = " & "'" & vpro_cod & "'"
        Cn.Execute salite
        'Fin actualiza items
        
        Adodc2.Recordset.MoveNext
        vInELot = vInELot + 1
    Loop
    
End If
'Fin

If modoini = "N" Then
    MsgBox "Ingreso registrado", vbInformation, empresa
ElseIf modoini = "M" Then
    MsgBox "Ingreso modificado", vbInformation, empresa
End If
Adodc1.RecordSource = "Select * from cabinvinicial order By Ini_id Desc"
Adodc1.Refresh

Adodc3.RecordSource = "Select * from detInvInicial where InI_id = " & 0
Adodc3.Refresh




creatabla
limpiadatos
SSFrame1.Enabled = True
SSFrame2.Enabled = False
End Sub

Private Sub SSCommand8_Click()
If seleccioni = 1 Then
    vite = SSOleDBGrid2.Columns(0).Value
    vitedes = SSOleDBGrid2.Columns(2).Value
    If MsgBox("Desea borrar el item " & vite & " ..?", vbYesNo, empresa) = vbYes Then
        Dim Cn As New ADODB.Connection
        Cn.ConnectionString = Cadena
        Cn.Open
        
        borrai = "DELETE from Tdetinvinicial" & vusuariot & " Where IteDes = " & "'" & vitedes & "'"
        Cn.Execute borrai
        Adodc2.Refresh
        MsgBox "Item eliminado", vbInformation, empresa
        SSCommand2.Caption = "Registrar"
    End If
Else
    MsgBox "Debe seleccionar un item de la lista", vbInformation, empresa
End If

End Sub

Private Sub SSCommand9_Click()
creatabla
limpiadatos
Combo2.Enabled = True
Combo2.SetFocus
End Sub
Private Function creatabla()
On Error GoTo errort

Dim TaUsu As String
vusuariot = Trim(vusucod)
TaUsu = "CREATE TABLE Tdetinvinicial" & vusuariot & "(" _
        & "InI_Id int(5) DEFAULT NULL, " _
        & "Pro_id int(5) DEFAULT NULL, " _
        & "Pro_cod varchar(30) DEFAULT NULL, " _
        & "Pro_Des varchar(250) DEFAULT NULL, " _
        & "Ini_Can int(5) DEFAULT NULL, " _
        & "Ini_PrU double(8,2) DEFAULT NULL, " _
        & "Ini_PrT double(8,2) DEFAULT NULL, " _
        & "Ini_PrV double(8,2) DEFAULT NULL, " _
        & "Ini_PrO double(8,2) DEFAULT NULL," _
        & "ProTli varchar(5) DEFAULT NULL," _
        & "ProCol varchar(50) DEFAULT NULL)"
        
Dim Cn As New ADODB.Connection
Cn.ConnectionString = Cadena
Cn.Open
Cn.Execute TaUsu
Adodc2.RecordSource = "Select * from Tdetinvinicial" & vusuariot
Adodc2.Refresh

errort:
If Err.Number = -2147217900 Then
    borrat = "Drop Table Tdetinvinicial" & vusuariot
    Cn.Execute borrat
    creatabla
End If
End Function

Private Sub Text6_GotFocus()
If Len(Trim(Combo2.Text)) > 0 Then
    valm = Combo2.Text
    
    Dim Cn As New ADODB.Connection
    Dim rsal As New ADODB.Recordset   ' Recordset de almacenes
    Cn.ConnectionString = Cadena
    Cn.Open
    
    rsal.CursorType = adOpenKeyset
    rsal.LockType = adLockOptimistic
    rsal.ActiveConnection = Cn
    rsal.Source = "Select * from almacen WHERE AlmDes = " & "'" & valm & "'"
    rsal.Open
    
    If Not rsal.EOF Then
        valmsig = rsal!almsig
        valmid = rsal!almid
    Else
        MsgBox "Debe seleccionar un almacén de la lista", vbInformation, empresa
        Combo2.SetFocus
    End If
    
    Text6.BackColor = &H80FFFF
    Text6.SelStart = 0
    Text6.SelLength = Len(Text6.Text)
Else
    MsgBox "Debe seleccionar un almacén de destino", vbInformation, empresa
    Combo2.SetFocus
End If
End Sub

Private Sub SSOleDBGrid1_Click()
vini_id = SSOleDBGrid1.Columns(0).Value
SSDateCombo1.Text = SSOleDBGrid1.Columns(1).Value
Adodc3.RecordSource = "Select * from detinvinicial where InI_id = " & vini_id
Adodc3.Refresh
seleccion = 1
End Sub

Private Sub SSOleDBGrid2_Click()
selecciond = 1
Text6.Text = Adodc2.Recordset.Fields("Pro_Cod")
Label8.Caption = Adodc2.Recordset.Fields("Pro_Des")
Combo1.Text = Adodc2.Recordset.Fields("ProTli")
Combo3.Text = Adodc2.Recordset.Fields("ProCol")
Text1.Text = Format(Adodc2.Recordset.Fields("InI_Can"), "####.#0")
Text2.Text = Format(Adodc2.Recordset.Fields("InI_PrU"), "####.#0")
Text3.Text = Format(Adodc2.Recordset.Fields("InI_PrT"), "####.#0")
Text4.Text = Format(Adodc2.Recordset.Fields("InI_PrV"), "####.#0")
Text5.Text = Format(Adodc2.Recordset.Fields("InI_PrO"), "####.#0")
vproidm = Adodc2.Recordset.Fields("Pro_id")
Text1.SetFocus
SSCommand2.Caption = "Modificar"
End Sub

Private Sub SSOleDBGrid2_DblClick()
If Adodc2.Recordset.RecordCount > 0 Then
    vcodigo = SSOleDBGrid2.Columns(0).Value
    If MsgBox("Confirma borrar el registro " & vcodigo & "..?", vbYesNo, empresa) = vbYes Then
        Adodc2.Recordset.Delete
        Adodc2.Refresh
    End If
End If
End Sub

Private Sub Text1_GotFocus()
If Len(Trim(Combo3.Text)) > 0 Then
    Text1.BackColor = &H80FFFF
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
Else
    MsgBox "Debe seleccionar el color", vbInformation, empresa
    Combo3.SetFocus
End If
End Sub
Private Sub Text1_LostFocus()
Text1.BackColor = &H80000005
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 46 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub Text2_GotFocus()
If Len(Trim(Text1.Text)) > 0 Then
    Text2.BackColor = &H80FFFF
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
Else
    MsgBox "Debe ingresar la cantidad", vbInformation, empresa
    Text1.SetFocus
End If
End Sub
Private Sub Text2_LostFocus()
If Len(Trim(Text2.Text)) > 0 Then
    Text3.Text = (Format(Val(Text1.Text) * Val(Text2.Text), "####.#0"))
    Text2.Text = Format(Text2.Text, "####.#0")
End If
Text2.BackColor = &H80000005
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 46 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub Text3_GotFocus()
If Len(Trim(Text2.Text)) > 0 Then
    Text3.BackColor = &H80FFFF
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.Text = Format((Val(Text1.Text) * Val(Text2.Text)), "####.#0")
Else
    MsgBox "Debe ingresar el costo unitario", vbInformation, empresa
    Text2.SetFocus
End If
End Sub
Private Sub Text3_LostFocus()
Text3.BackColor = &H80000005
If Len(Trim(Text2.Text)) = 0 Then
    Text2.Text = Format((Val(Text3.Text) / Val(Text1.Text)), "####.#0")
End If
Text3.Text = Format(Text3.Text, "####.#0")
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 46 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Function limpiadatos()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Combo1.Text = ""
Label8.Caption = ""
Combo3.Text = ""
Text6.Text = ""
End Function

Private Sub Text4_GotFocus()
Text4.BackColor = &H80FFFF
Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)
End Sub
Private Sub Text4_LostFocus()
Text4.BackColor = &H80000005
Text4.Text = Format(Text4.Text, "####.#0")
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 46 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Function grabainvini()
On Error GoTo erroin

Dim Cn As New ADODB.Connection
Dim rsit As New ADODB.Recordset   ' Recordset de item
Dim rsci As New ADODB.Recordset   ' Recordset de cabecera de inventario inicial
Dim rsdt As New ADODB.Recordset   ' Recordset de detalle temporal
Cn.ConnectionString = Cadena
Cn.Open
    
rsit.CursorType = adOpenKeyset
rsit.LockType = adLockOptimistic
rsit.ActiveConnection = Cn
rsit.Source = "Select * from productos Where Pro_Cod = " & "'" & vcod_pro & "'"
rsit.Open

vini_fec = Format(SSDateCombo1.Text, "yyyy-mm-dd")

'Graba cabecera de inventario inicial
If modoini = "N" Then
    grabacii = "INSERT INTO cabinvinicial SET ini_fec = " & "'" & vini_fec & "', usu_id = " & vusucod & ", almid = " & valmid
ElseIf modoini = "M" Then
    grabacii = "UPDATE cabinvinicial SET ini_fec = " & "'" & vini_fec & "', usu_id = " & vusucod & " WHERE ini_id = " & vini_id
End If
Cn.Execute grabacii
If modoini = "N" Then
    rsci.CursorType = adOpenKeyset
    rsci.LockType = adLockOptimistic
    rsci.ActiveConnection = Cn
    rsci.Source = "Select * from cabinvinicial"
    rsci.Open
    
    If Not rsci.EOF Then
        rsci.MoveLast
        vini_id = rsci!Ini_Id
    Else
        vini_id = 1
    End If
End If
'Graba el id de cabecera en el detalle temporal
grabaid = "UPDATE Tdetinvinicial" & vusuariot & " SET Ini_Id = " & vini_id
Cn.Execute grabaid
'Graba detalle de inventario inicial
detar = " INSERT INTO detinvinicial SELECT * from Tdetinvinicial" & vusuariot
Cn.Execute detar

'Graba en tabla productos

rsdt.CursorType = adOpenKeyset
rsdt.LockType = adLockOptimistic
rsdt.ActiveConnection = Cn
rsdt.Source = "Select * from Tdetinvinicial" & vusuariot
rsdt.Open

If Not rsdt.EOF Then
    Do While Not rsdt.EOF
        vpro_id = rsdt!pro_id
        vpro_cod = rsdt!pro_cod
        vini_can = rsdt!Ini_Can
        vini_pru = rsdt!ini_pru
        vini_prv = rsdt!ini_prv
        vini_pro = rsdt!ini_pro
        actpro = "UPDATE productos SET pro_exi = pro_exi + " & vini_can & ", pro_pve = " & vini_prv & ", pro_ppp = " & vini_pru & ", pro_pof = " & vini_pro & " WHERE pro_cod = " & "'" & vpro_cod & "'"
        Cn.Execute actpro
        actprotie = "UPDATE productostie SET pro_pve = " & vini_prv & ", pro_ppp = " & vini_pru & ", pro_pof = " & vini_pro & " WHERE pro_cod = " & "'" & vpro_cod & "'"
        Cn.Execute actprotie
        rsdt.MoveNext
    Loop
End If
Adodc1.Refresh
Cn.Close

erroin:
If Err.Number = -2147467259 Then
    MsgBox "Debe ingresar la fecha del inventario Inicial", vbInformation, empresa
    SSDateCombo1.SetFocus
End If
End Function
Private Function grabaalmacen()
Dim Cn As New ADODB.Connection
Dim rsal As New ADODB.Recordset   ' Recordset de almacen
Dim rste As New ADODB.Recordset   ' Recordset de temporal
Cn.ConnectionString = Cadena
Cn.Open
    
rsal.CursorType = adOpenKeyset
rsal.LockType = adLockOptimistic
rsal.ActiveConnection = Cn
rsal.Source = "Select * from " & valmsig
rsal.Open

rste.CursorType = adOpenKeyset
rste.LockType = adLockOptimistic
rste.ActiveConnection = Cn
rste.Source = "Select * from Tdetinvinicial" & vusuariot
rste.Open

If Not rste.EOF Then
    Do While Not rste.EOF
        rsal.AddNew
        rsal!almid = valmid
        rsal!procod = rste!pro_cod
        rsal!procan = rste!Ini_Can
        rsal!UsuRes = vusucod
        rsal.Update
        rste.MoveNext
    Loop
End If


End Function
Private Sub Text5_GotFocus()
If Len(Trim(Text4.Text)) > 0 Then
    Text5.Text = Text4.Text
    Text5.BackColor = &H80FFFF
    Text5.SelStart = 0
    Text5.SelLength = Len(Text5.Text)
Else
    MsgBox "Debe registrar el precio de venta", vbInformation, empresa
    Text4.SetFocus
End If
End Sub
Private Sub Text5_LostFocus()
Text5.BackColor = &H80000005
Text5.Text = Format(Text5.Text, "####.#0")
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 46 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub Text6_LostFocus()
Text6.BackColor = &H80000005
End Sub
Private Function buscaprod()
vcod_pro = Text6.Text
'vprotli = Combo1.Text
    
Dim Cn As New ADODB.Connection
Dim rsit As New ADODB.Recordset   ' Recordset de item
Cn.ConnectionString = Cadena
Cn.Open
    
rsit.CursorType = adOpenKeyset
rsit.LockType = adLockOptimistic
rsit.ActiveConnection = Cn
rsit.Source = "Select * from vproducto Where Pro_Cod = " & "'" & vcod_pro & "'" 'AND proTLi = " & "'" & vprotli & "'"
rsit.Open

If Not rsit.EOF Then
    Label8.Caption = rsit!Gru_des & " " & rsit!Pro_des
    Combo3.Text = rsit!Pro_Tip
    siexiste = 1
Else
    MsgBox "Item inexistente", vbInformation, empresa
    Text6.SetFocus
End If
End Function
