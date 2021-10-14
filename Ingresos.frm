VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form Ingresos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9180
   ClientLeft      =   5295
   ClientTop       =   3045
   ClientWidth     =   12660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   12660
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   2385
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
      RecordSource    =   "Select * from cabingreso order by cingr_id"
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
   Begin Threed.SSFrame SSFrame1 
      Height          =   2355
      Left            =   10755
      TabIndex        =   10
      Top             =   675
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   4154
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin Threed.SSCommand SSCommand1 
         Height          =   420
         Left            =   180
         TabIndex        =   9
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
      Begin Threed.SSCommand SSCommand11 
         Height          =   420
         Left            =   180
         TabIndex        =   11
         Top             =   1440
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
      Left            =   4455
      Top             =   2655
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
      RecordSource    =   "Select * from vingreso where cingr_id=0"
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
   Begin Threed.SSFrame SSFrame2 
      Height          =   4920
      Left            =   225
      TabIndex        =   12
      Top             =   4050
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   8678
      _Version        =   196608
      BackStyle       =   1
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.TextBox Text6 
         Height          =   330
         Left            =   315
         TabIndex        =   2
         Top             =   1260
         Width           =   1995
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   5175
         TabIndex        =   7
         Top             =   1935
         Width           =   1320
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Width           =   4335
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   4230
         Top             =   3600
         Width           =   2265
         _ExtentX        =   3995
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
         RecordSource    =   "Select * from detingreso where cing_id=0"
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid2 
         Bindings        =   "Ingresos.frx":0000
         Height          =   2265
         Left            =   180
         TabIndex        =   31
         Top             =   2430
         Width           =   9015
         _Version        =   196616
         ForeColorEven   =   0
         BackColorEven   =   14737632
         BackColorOdd    =   14737632
         RowHeight       =   423
         Columns.Count   =   6
         Columns(0).Width=   2699
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "pro_cod"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "pro_cod"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1402
         Columns(1).Caption=   "Cant."
         Columns(1).Name =   "ding_can"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "ding_can"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         Columns(2).Width=   5398
         Columns(2).Caption=   "Descripción"
         Columns(2).Name =   "pro_des"
         Columns(2).CaptionAlignment=   0
         Columns(2).DataField=   "pro_des"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   1720
         Columns(3).Caption=   "Costo Un."
         Columns(3).Name =   "ding_pru"
         Columns(3).Alignment=   1
         Columns(3).CaptionAlignment=   1
         Columns(3).DataField=   "ding_pru"
         Columns(3).DataType=   5
         Columns(3).NumberFormat=   "####.#0"
         Columns(3).FieldLen=   256
         Columns(4).Width=   1879
         Columns(4).Caption=   "Total"
         Columns(4).Name =   "ding_Prt"
         Columns(4).Alignment=   1
         Columns(4).CaptionAlignment=   1
         Columns(4).DataField=   "ding_Prt"
         Columns(4).DataType=   5
         Columns(4).NumberFormat=   "####.#0"
         Columns(4).FieldLen=   256
         Columns(5).Width=   1773
         Columns(5).Caption=   "Pr. Venta"
         Columns(5).Name =   "ding_prv"
         Columns(5).Alignment=   1
         Columns(5).CaptionAlignment=   1
         Columns(5).DataField=   "pro_prv"
         Columns(5).DataType=   5
         Columns(5).NumberFormat=   "####.#0"
         Columns(5).FieldLen=   256
         _ExtentX        =   15901
         _ExtentY        =   3995
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
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   315
         TabIndex        =   3
         Top             =   1935
         Width           =   645
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1125
         TabIndex        =   4
         Top             =   1935
         Width           =   1005
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   2295
         TabIndex        =   5
         Top             =   1935
         Width           =   1185
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   3645
         TabIndex        =   6
         Top             =   1935
         Width           =   1320
      End
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   375
         Left            =   8550
         Top             =   4230
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
      Begin Threed.SSCommand SSCommand2 
         Height          =   375
         Left            =   6930
         TabIndex        =   8
         Top             =   1890
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
      Begin Threed.SSCommand SSCommand3 
         Height          =   600
         Left            =   9270
         TabIndex        =   13
         Top             =   2925
         Width           =   1140
         _ExtentX        =   2011
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
         Caption         =   "&Grabar Compra"
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand5 
         Height          =   375
         Left            =   9045
         TabIndex        =   30
         Top             =   1890
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   661
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
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo1 
         Height          =   375
         Left            =   270
         TabIndex        =   0
         Top             =   315
         Width           =   1680
         _Version        =   65537
         _ExtentX        =   2963
         _ExtentY        =   661
         _StockProps     =   93
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Precio de oferta"
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
         Left            =   5175
         TabIndex        =   34
         Top             =   1710
         Width           =   1380
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
         Left            =   2205
         TabIndex        =   33
         Top             =   135
         Width           =   1440
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de compra"
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
         Left            =   270
         TabIndex        =   32
         Top             =   90
         Width           =   1485
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
         Left            =   2565
         TabIndex        =   25
         Top             =   1035
         Width           =   780
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
         Left            =   360
         TabIndex        =   24
         Top             =   1710
         Width           =   465
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
         Left            =   1125
         TabIndex        =   23
         Top             =   1710
         Width           =   960
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
         Left            =   2295
         TabIndex        =   22
         Top             =   1710
         Width           =   1050
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   7560
         TabIndex        =   21
         Top             =   1260
         Width           =   1545
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
         Left            =   7560
         TabIndex        =   20
         Top             =   1035
         Width           =   450
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   9540
         TabIndex        =   19
         Top             =   585
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cant. por"
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
         Left            =   9405
         TabIndex        =   18
         Top             =   345
         Visible         =   0   'False
         Width           =   795
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
         TabIndex        =   17
         Top             =   3465
         Width           =   75
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
         Left            =   360
         TabIndex        =   16
         Top             =   1035
         Width           =   600
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
         Left            =   3645
         TabIndex        =   15
         Top             =   1710
         Width           =   1365
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2520
         TabIndex        =   14
         Top             =   1260
         Width           =   4965
      End
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid3 
      Bindings        =   "Ingresos.frx":0015
      Height          =   3210
      Left            =   2700
      TabIndex        =   27
      Top             =   675
      Width           =   7980
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
      Columns.Count   =   5
      Columns(0).Width=   2408
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "pro_cod"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "pro_cod"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1191
      Columns(1).Caption=   "Cant."
      Columns(1).Name =   "ding_can"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "ding_can"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   5821
      Columns(2).Caption=   "Producto"
      Columns(2).Name =   "pro_Des"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "pro_Des"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1614
      Columns(3).Caption=   "Precio"
      Columns(3).Name =   "ding_pu"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   1
      Columns(3).DataField=   "ding_pru"
      Columns(3).DataType=   5
      Columns(3).NumberFormat=   "####.#0"
      Columns(3).FieldLen=   256
      Columns(4).Width=   1984
      Columns(4).Caption=   "Total"
      Columns(4).Name =   "ding_tot"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   1
      Columns(4).DataField=   "ding_prt"
      Columns(4).DataType=   5
      Columns(4).NumberFormat=   "####.#0"
      Columns(4).FieldLen=   256
      _ExtentX        =   14076
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
      Left            =   10935
      TabIndex        =   28
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
      Bindings        =   "Ingresos.frx":002A
      Height          =   3210
      Left            =   180
      TabIndex        =   26
      Top             =   675
      Width           =   2400
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
      Columns.Count   =   2
      Columns(0).Width=   1005
      Columns(0).Caption=   "Num."
      Columns(0).Name =   "InICod"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   1
      Columns(0).DataField=   "cingr_id"
      Columns(0).DataType=   3
      Columns(0).FieldLen=   256
      Columns(1).Width=   2249
      Columns(1).Caption=   "Fecha"
      Columns(1).Name =   "InIFec"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "cingr_fec"
      Columns(1).DataType=   7
      Columns(1).NumberFormat=   "dd-mm-yyyy"
      Columns(1).FieldLen=   256
      _ExtentX        =   4233
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
   Begin VB.Image Image1 
      Height          =   60
      Left            =   6075
      Picture         =   "Ingresos.frx":003F
      Stretch         =   -1  'True
      Top             =   450
      Width           =   6540
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "INGRESO DE PRODUCTOS"
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
      Left            =   9135
      TabIndex        =   29
      Top             =   135
      Width           =   3345
   End
End
Attribute VB_Name = "Ingresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vcing_id As Integer
Private Sub Combo2_GotFocus()
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
Private Sub SSCommand1_Click()
limpiadatos
creatabladi
ssFrame1.Enabled = False
SSFrame2.Enabled = True
'SSFrame3.Enabled = True
modo = "N"
SSDateCombo1.SetFocus
End Sub

Private Sub SSCommand11_Click()
If seleccion = 1 Then
    creatabladi
    ssFrame1.Enabled = False
    SSFrame2.Enabled = True
'    SSFrame3.Enabled = True
    modo = "M"
    
    'Carga cabecera
    SSDateCombo1.Text = SSOleDBGrid1.Columns(1).Value
    'Fin cabecera
    
    'Carga detalle
    Dim Cn As New ADODB.Connection
    Cn.ConnectionString = Cadena
    Cn.Open
    vini_id = SSOleDBGrid1.Columns(0).Value
    cargad = "Insert into tdetingreso" & vusuariot & " Select * from detingreso where cing_id = " & vini_id
    Cn.Execute cargad
    Adodc2.Refresh
    'FIn detalle
Else
    MsgBox "Debe seleccionar un ingreso de la lista", vbInformation, empresa
End If

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
        
        rspr.CursorType = adOpenKeyset
        rspr.LockType = adLockOptimistic
        rspr.ActiveConnection = Cn
        rspr.Source = "Select * from productos Where Pro_cod = " & "'" & vpro_cod & "'"
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
        ving_pro = Val(Text5.Text)
        
        vInIFeA = Date
        vInIFeA = Format(vInIFeA, "YYYY-MM-dd")
        
        If SSCommand2.Caption = "Registrar" Then
            'Verifica si el item ya existe
            rsdt.CursorType = adOpenKeyset
            rsdt.LockType = adLockOptimistic
            rsdt.ActiveConnection = Cn
            rsdt.Source = "Select * from Tdetingreso" & vusuariot & " Where Pro_Des = " & "'" & vpro_des & "'"
            rsdt.Open
        
            If rsdt.EOF And SSCommand2.Caption = "Registrar" Then
                Text6.Enabled = False
                nuevoe = "Insert into Tdetingreso" & vusuariot & " SET Pro_id = " & vpro_id & ", Pro_Cod = " & "'" & vpro_cod & "'" & ", Pro_Des = " & "'" & vpro_des & "', dIng_Can = " & vini_can & _
                ", ding_PrU = " & vini_pru & ", ding_PrT = " & vIni_PrT & ", Pro_Prv = " & vini_prv & ", ding_Ppp = " & vini_pru & ", ding_Pro = " & ving_pro
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
         ElseIf SSCommand2.Caption = "Modificar" Then
            vpro_cod1 = SSOleDBGrid2.Columns(0).Value
            borrai = "Delete from Tdetingreso" & vusuariot & " Where pro_cod = " & "'" & vpro_cod1 & "'"
            Cn.Execute borrai
            
            nuevoe = "Insert into Tdetingreso" & vusuariot & " SET Pro_id = " & vpro_id & ", Pro_Cod = " & "'" & vpro_cod & "'" & ", Pro_Des = " & "'" & vpro_des & "', dIng_Can = " & vini_can & _
            ", ding_PrU = " & vini_pru & ", ding_PrT = " & vIni_PrT & ", Pro_Prv = " & vini_prv & ", ding_Ppp = " & vini_pru & ", ding_Ppp = " & vini_pru & ", ding_Pro = " & ving_pro
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
If modo = "M" Then
    vcing_id = SSOleDBGrid1.Columns(0).Value
    Dim Cn As New ADODB.Connection
    Dim rsir As New ADODB.Recordset   ' Recordset de detalle de inv inicial
    Dim rsit As New ADODB.Recordset   ' Recordset de items
    Cn.ConnectionString = Cadena
    Cn.Open
    
    'Disminuye actualiza Productos
    rsir.CursorType = adOpenKeyset
    rsir.LockType = adLockOptimistic
    rsir.ActiveConnection = Cn
    rsir.Source = "Select * from detingreso where cing_id = " & vcing_id
    rsir.Open

    If Not rsir.EOF Then
        Do While Not rsir.EOF
            vpro_id = rsir!pro_id
            vini_can = rsir!ding_Can
            actpro = "UPDATE productos SET pro_exi = pro_exi - " & vini_can & " WHERE pro_id = " & vpro_id
            Cn.Execute actpro
            rsir.MoveNext
        Loop
    End If
    
    ''Borra cabecera
    'borrac = "DELETE FROM cabinvinicial where ini_id = " & vini_id
    'Cn.Execute borrac
    
    'Borra detalle
    borrad = "DELETE FROM detingreso where cing_id = " & vcing_id
    Cn.Execute borrad
End If

grabaingreso
grabaalmacen
limpiadatos
creatabladi
SSFrame2.Enabled = False
'SSFrame3.Enabled = False
ssFrame1.Enabled = True
End Sub

Private Sub SSCommand5_Click()
origen = "ingreso"
Ingresos.Enabled = False
Load Productos
Productos.Show
End Sub

Private Sub SSCommand6_Click()
Menup.Enabled = True
Unload Ingresos
Set Ingresos = Nothing
Menup.Label4.ForeColor = &HE0E0E0
End Sub


Private Function creatabladi()
On Error GoTo errort

Dim TaUsu As String
vusuariot = Trim(vusucod)
TaUsu = "CREATE TABLE Tdetingreso" & vusuariot & "(" _
        & "cing_Id int(5) DEFAULT NULL, " _
        & "ding_Id int(5) DEFAULT NULL, " _
        & "Pro_id int(5) DEFAULT NULL, " _
        & "Pro_cod varchar(30) DEFAULT NULL, " _
        & "Pro_Des varchar(250) DEFAULT NULL, " _
        & "ding_Can int(5) DEFAULT NULL, " _
        & "ding_PrU double(8,2) DEFAULT NULL, " _
        & "ding_PrT double(8,2) DEFAULT NULL, " _
        & "pro_PrV double(8,2) DEFAULT NULL, " _
        & "ding_ppp double(8,2) DEFAULT NULL, " _
        & "ding_prO double(8,2) DEFAULT NULL)"
        
Dim Cn As New ADODB.Connection
Cn.ConnectionString = Cadena
Cn.Open
Cn.Execute TaUsu
Adodc2.RecordSource = "Select * from Tdetingreso" & vusuariot
Adodc2.Refresh

errort:
If Err.Number = -2147217900 Then
    borrat = "Drop Table Tdetingreso" & vusuariot
    Cn.Execute borrat
    creatabladi
End If
End Function

Private Sub SSCommand7_Click()
origen = "ingreso"
Ingresos.Enabled = False
Load Proveedores
Proveedores.Show
End Sub

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
Else
    MsgBox "Debe seleccionar un almacén de destino", vbInformation, empresa
    Combo2.SetFocus
End If
End Sub

Private Sub SSOleDBGrid1_Click()
vini_id = SSOleDBGrid1.Columns(0).Value
Adodc3.RecordSource = "Select * from detingreso where cing_id = " & vini_id
Adodc3.Refresh
seleccion = 1
End Sub

Private Sub SSOleDBGrid2_Click()
selecciond = 1
Text6.Text = Adodc2.Recordset.Fields("Pro_Cod")
Label8.Caption = SSOleDBGrid2.Columns(2).Value
Text1_GotFocus
Text1.Text = Format(Adodc2.Recordset.Fields("ding_Can"), "####.#0")
Text2.Text = Format(Adodc2.Recordset.Fields("ding_PrU"), "####.#0")
Text3.Text = Format(Adodc2.Recordset.Fields("ding_PrT"), "####.#0")
Text4.Text = Format(Adodc2.Recordset.Fields("pro_PrV"), "####.#0")
SSCommand2.Caption = "Modificar"

End Sub

Private Sub SSOleDBGrid2_DblClick()
Adodc2.Recordset.Delete
Adodc2.Refresh
limpiadatos
End Sub

Private Sub Text1_GotFocus()
If Len(Trim(Text6.Text)) > 0 Then
    vcod_pro = Text6.Text
        
    Dim Cn As New ADODB.Connection
    Dim rsit As New ADODB.Recordset   ' Recordset de item
    Cn.ConnectionString = Cadena
    Cn.Open
        
    rsit.CursorType = adOpenKeyset
    rsit.LockType = adLockOptimistic
    rsit.ActiveConnection = Cn
    rsit.Source = "Select * from vproducto Where Pro_Cod = " & "'" & vcod_pro & "'"
    rsit.Open
    
    If Not rsit.EOF Then
        Label8.Caption = rsit!Gru_des & " " & rsit!Pro_des
        Label11.Caption = rsit!Pro_uni
        Text2.Text = Format(rsit!pro_ppp, "####.#0")
        Text4.Text = Format(rsit!pro_pve, "####.#0")
        Text5.Text = Format(rsit!pro_pof, "####.#0")
    Else
        MsgBox "Item inexistente", vbInformation, empresa
        Text6.SetFocus
    End If
       
    Text1.BackColor = &H9FBDEA
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
Else
    MsgBox "Debe seleccionar el código de item", vbInformation, empresa
    Text6.SetFocus
End If
End Sub
Private Sub Text1_LostFocus()
Text1.BackColor = &HFFFFFF
Text1.Text = UCase(Text1.Text)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 46 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub Text2_GotFocus()
If Len(Trim(Text1.Text)) > 0 Then
    Text2.BackColor = &H9FBDEA
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
Else
    MsgBox "Debe ingresar la cantidad", vbInformation, empresa
    Text1.SetFocus
End If
End Sub
Private Sub Text2_LostFocus()
Text2.BackColor = &HFFFFFF
Text2.Text = Format(Text2.Text, "####.#0")
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 46 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub Text3_GotFocus()
If Len(Trim(Text2.Text)) > 0 Then
    Text3.BackColor = &H9FBDEA
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.Text = Format((Val(Text1.Text) * Val(Text2.Text)), "####.#0")
Else
    MsgBox "Debe ingresar el costo unitario", vbInformation, empresa
    Text2.SetFocus
End If
End Sub
Private Sub Text3_LostFocus()
Text3.BackColor = &HFFFFFF
Text3.Text = Format(Text3.Text, "#####.#0")
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 46 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub Text4_GotFocus()
Text4.BackColor = &H9FBDEA
End Sub
Private Sub Text4_LostFocus()
Text4.BackColor = &HFFFFFF
Text4.Text = Format(Text4.Text, "#####.#0")
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 46 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub Text5_GotFocus()
Text5.BackColor = &H80FFFF
End Sub
Private Sub Text5_LostFocus()
Text5.BackColor = &HFFFFFF
Text5.Text = Format(Text5.Text, "#####.#0")
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 46 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub Text6_LostFocus()
Text6.BackColor = &HFFFFFF
Text6.Text = UCase(Text6.Text)
End Sub

Private Sub Text7_GotFocus()
Text7.BackColor = &H80FFFF
End Sub
Private Sub Text7_LostFocus()
Text7.BackColor = &HFFFFFF
Text7.Text = UCase(Text7.Text)
End Sub

Private Sub Text8_GotFocus()
Text8.BackColor = &H80FFFF
End Sub
Private Sub Text8_LostFocus()
Text8.BackColor = &HFFFFFF
Text8.Text = UCase(Text8.Text)
End Sub
Private Sub Text9_GotFocus()
Text9.BackColor = &H80FFFF
End Sub
Private Sub Text9_LostFocus()
Text9.BackColor = &HFFFFFF
Text9.Text = UCase(Text9.Text)
End Sub

Private Function limpiadatos()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
'Label2.Caption = ""
Label8.Caption = ""
Label11.Caption = ""
Text6.Text = ""
End Function

Private Function grabaingreso()
Dim Cn As New ADODB.Connection
Dim rsit As New ADODB.Recordset   ' Recordset de item
Dim rsci As New ADODB.Recordset   ' Recordset de cabecera de ingreso
Dim rsdt As New ADODB.Recordset   ' Recordset de detalle temporal
Cn.ConnectionString = Cadena
Cn.Open
    
rsit.CursorType = adOpenKeyset
rsit.LockType = adLockOptimistic
rsit.ActiveConnection = Cn
rsit.Source = "Select * from productos Where Pro_Cod = " & "'" & vcod_pro & "'"
rsit.Open

vini_fec = Format(SSDateCombo1.Text, "yyyy-mm-dd")

'Graba cabecera de ingreso
If modo = "N" Then
    grabacii = "INSERT INTO cabingreso SET cingr_fec = " & "'" & vini_fec & "', usu_id = " & vusucod & ", cing_obs = 'C'"
ElseIf modo = "M" Then
    grabacii = "UPDATE cabingreso SET cingr_fec = " & "'" & vini_fec & "', usu_id = " & vusucod & " WHERE cingr_id = " & vcing_id
End If
Cn.Execute grabacii
If modo = "N" Then
    rsci.CursorType = adOpenKeyset
    rsci.LockType = adLockOptimistic
    rsci.ActiveConnection = Cn
    rsci.Source = "Select * from cabingreso"
    rsci.Open
    
    If Not rsci.EOF Then
        rsci.MoveLast
        vini_id = rsci!cingr_id
    Else
        vini_id = 1
    End If
End If
'Graba el id de cabecera en el detalle temporal
grabaid = "UPDATE Tdetingreso" & vusuariot & " SET cing_Id = " & vini_id
Cn.Execute grabaid

'Graba detalle de ingreso de temporal a real
detar = " INSERT INTO detingreso SELECT * from Tdetingreso" & vusuariot
Cn.Execute detar

'Graba en tabla productos
rsdt.CursorType = adOpenKeyset
rsdt.LockType = adLockOptimistic
rsdt.ActiveConnection = Cn
rsdt.Source = "Select * from Tdetingreso" & vusuariot
rsdt.Open

If Not rsdt.EOF Then
    Do While Not rsdt.EOF
        vpro_id = rsdt!pro_id
        vini_can = rsdt!ding_Can
        vini_pru = rsdt!ding_pru
        vini_prv = rsdt!pro_prv
        actpro = "UPDATE productos SET pro_exi = pro_exi + " & vini_can & ", pro_prC = " & vini_pru & ", pro_pve = " & vini_prv & ", pro_ppp = " & vini_pru & " WHERE pro_id = " & vpro_id
        Cn.Execute actpro
        rsdt.MoveNext
    Loop
End If
Adodc1.Refresh
Cn.Close
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
rste.Source = "Select * from Tdetingreso" & vusuariot
rste.Open

If Not rste.EOF Then
    Do While Not rste.EOF
        rsal.AddNew
        rsal!almid = valmid
        rsal!procod = rste!pro_cod
        rsal!procan = rste!ding_Can
        rsal!UsuRes = vusucod
        rsal.Update
        rste.MoveNext
    Loop
End If


End Function

