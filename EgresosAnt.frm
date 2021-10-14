VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form EgresosAnt 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5085
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15165
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   15165
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5355
      Top             =   225
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowProgressCtls=   0   'False
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   420
      Left            =   13500
      TabIndex        =   0
      Top             =   4095
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2070
      Top             =   2520
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
      RecordSource    =   "Select * from pagoventa order by cegr_id"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   10440
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
      RecordSource    =   "Select * from detegreso where cegr_id=0"
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid3 
      Bindings        =   "EgresosAnt.frx":0000
      Height          =   3210
      Left            =   7740
      TabIndex        =   1
      Top             =   810
      Width           =   7245
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
      Columns(0).Width=   1905
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "pro_cod"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "pro_cod"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1138
      Columns(1).Caption=   "Cant."
      Columns(1).Name =   "degr_can"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "degr_can"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      Columns(2).Width=   5159
      Columns(2).Caption=   "Producto"
      Columns(2).Name =   "pro_des"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "pro_des"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1773
      Columns(3).Caption=   "Costo"
      Columns(3).Name =   "degr_pru"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   1
      Columns(3).DataField=   "degr_pru"
      Columns(3).DataType=   5
      Columns(3).NumberFormat=   "####.#0"
      Columns(3).FieldLen=   256
      Columns(4).Width=   1746
      Columns(4).Caption=   "Total"
      Columns(4).Name =   "degr_prt"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   1
      Columns(4).DataField=   "degr_prt"
      Columns(4).DataType=   5
      Columns(4).NumberFormat=   "####.#0"
      Columns(4).FieldLen=   256
      _ExtentX        =   12779
      _ExtentY        =   5662
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid1 
      Bindings        =   "EgresosAnt.frx":0015
      Height          =   3210
      Left            =   90
      TabIndex        =   2
      Top             =   810
      Width           =   7590
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
      Columns(0).Width=   1693
      Columns(0).Caption=   "No.Venta"
      Columns(0).Name =   "cegr_id"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   1
      Columns(0).DataField=   "cegr_id"
      Columns(0).DataType=   3
      Columns(0).FieldLen=   256
      Columns(1).Width=   2249
      Columns(1).Caption=   "Fecha"
      Columns(1).Name =   "cegr_fec"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Pag_fec"
      Columns(1).DataType=   7
      Columns(1).NumberFormat=   "dd-mm-yyyy"
      Columns(1).FieldLen=   256
      Columns(2).Width=   8334
      Columns(2).Caption=   "Cliente"
      Columns(2).Name =   "cegr_clie"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Pag_RaS"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   13388
      _ExtentY        =   5662
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
   Begin Threed.SSFrame SSFrame2 
      Height          =   915
      Left            =   90
      TabIndex        =   4
      Top             =   4050
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   1614
      _Version        =   196608
      BackColor       =   14737632
      ClipControls    =   0   'False
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   135
         TabIndex        =   5
         Top             =   270
         Width           =   3570
      End
      Begin VB.Image Image3 
         Height          =   555
         Left            =   5715
         Picture         =   "EgresosAnt.frx":002A
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1815
      End
      Begin VB.Image Image2 
         Height          =   555
         Left            =   3825
         Picture         =   "EgresosAnt.frx":0D31
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
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
         Left            =   135
         TabIndex        =   6
         Top             =   45
         Width           =   1545
      End
   End
   Begin VB.Image Image1 
      Height          =   60
      Left            =   8460
      Picture         =   "EgresosAnt.frx":1A79
      Stretch         =   -1  'True
      Top             =   495
      Width           =   6540
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "VENTAS ANTERIORES"
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
      Left            =   12150
      TabIndex        =   3
      Top             =   135
      Width           =   2805
   End
End
Attribute VB_Name = "EgresosAnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image2_Click()
If Len(Trim(Text8.Text)) > 0 Then
    Dim db1 As String
    db1 = Text8.Text
    Adodc1.RecordSource = "SELECT * from vventa1 WHERE  pag_ras LIKE " & "'%" & db1 & "%'"
    Adodc1.Refresh
    Text8.Text = ""
End If

End Sub
Private Sub Image3_Click()
Adodc1.RecordSource = "Select * from vventa1 order by cegr_id"
Adodc1.Refresh
End Sub

Private Sub SSCommand6_Click()
If origen = "menup" Then
    Menup.Enabled = True
ElseIf origen = "devoluciones" Then
    Devoluciones.Enabled = True
End If
Unload EgresosAnt
Set EgresosAnt = Nothing
End Sub
Private Sub SSOleDBGrid1_Click()
vpro_id = SSOleDBGrid1.Columns(0).Value
Adodc3.RecordSource = "Select * from detegreso where cegr_id = " & vpro_id
Adodc3.Refresh
seleccion = 1
End Sub
