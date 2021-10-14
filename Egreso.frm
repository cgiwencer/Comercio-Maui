VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form Egreso 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7815
   ClientLeft      =   4980
   ClientTop       =   3150
   ClientWidth     =   12585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   12585
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame SSFrame3 
      Height          =   825
      Left            =   225
      TabIndex        =   10
      Top             =   630
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1455
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1980
         TabIndex        =   6
         Top             =   270
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   3735
         TabIndex        =   7
         Top             =   270
         Visible         =   0   'False
         Width           =   4605
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   8415
         TabIndex        =   8
         Top             =   270
         Visible         =   0   'False
         Width           =   1680
      End
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo1 
         Height          =   375
         Left            =   225
         TabIndex        =   5
         Top             =   270
         Width           =   1680
         _Version        =   65537
         _ExtentX        =   2963
         _ExtentY        =   661
         _StockProps     =   93
         Enabled         =   0   'False
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
         Left            =   225
         TabIndex        =   14
         Top             =   45
         Width           =   1035
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
         TabIndex        =   13
         Top             =   45
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Razón Social"
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
         TabIndex        =   12
         Top             =   45
         Visible         =   0   'False
         Width           =   1140
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
         Left            =   8460
         TabIndex        =   11
         Top             =   45
         Visible         =   0   'False
         Width           =   1155
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   6135
      Left            =   225
      TabIndex        =   15
      Top             =   1530
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   10821
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2205
         TabIndex        =   1
         Top             =   450
         Width           =   1140
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   90
         TabIndex        =   0
         Top             =   450
         Width           =   2040
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   2475
         TabIndex        =   4
         Top             =   1170
         Width           =   1320
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   7650
         TabIndex        =   31
         Text            =   "0.00"
         Top             =   4860
         Width           =   555
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   7650
         TabIndex        =   29
         Text            =   "0.00"
         Top             =   4455
         Width           =   1275
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   945
         TabIndex        =   3
         Top             =   1170
         Width           =   1320
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   90
         TabIndex        =   2
         Top             =   1170
         Width           =   645
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   4230
         Top             =   3060
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
         RecordSource    =   "Select * from detegreso where cegr_id=0"
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
         Bindings        =   "Egreso.frx":0000
         Height          =   2265
         Left            =   180
         TabIndex        =   16
         Top             =   1710
         Width           =   9015
         _Version        =   196616
         ForeColorEven   =   0
         BackColorEven   =   14737632
         BackColorOdd    =   14737632
         RowHeight       =   423
         Columns.Count   =   5
         Columns(0).Width=   2725
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "pro_cod"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "pro_cod"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1323
         Columns(1).Caption=   "Cantidad"
         Columns(1).Name =   "degr_can"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "degr_can"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         Columns(2).Width=   7064
         Columns(2).Caption=   "Descripción"
         Columns(2).Name =   "pro_des"
         Columns(2).CaptionAlignment=   0
         Columns(2).DataField=   "pro_des"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   1931
         Columns(3).Caption=   "Precio Unit."
         Columns(3).Name =   "degr_pru"
         Columns(3).Alignment=   1
         Columns(3).CaptionAlignment=   1
         Columns(3).DataField=   "degr_pru"
         Columns(3).DataType=   5
         Columns(3).NumberFormat=   "####.#0"
         Columns(3).FieldLen=   256
         Columns(4).Width=   1773
         Columns(4).Caption=   "Total"
         Columns(4).Name =   "degr_prt"
         Columns(4).Alignment=   1
         Columns(4).CaptionAlignment=   1
         Columns(4).DataField=   "degr_prt"
         Columns(4).DataType=   5
         Columns(4).NumberFormat=   "####.#0"
         Columns(4).FieldLen=   256
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
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   375
         Left            =   8550
         Top             =   3510
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
         RecordSource    =   "Select * from tie order by ProCod AND ProCan > 0"
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
         Left            =   4140
         TabIndex        =   9
         Top             =   1125
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
         TabIndex        =   17
         Top             =   2205
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
         Caption         =   "&Grabar Venta"
         BevelWidth      =   1
      End
      Begin VB.Label Label26 
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
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   2475
         TabIndex        =   38
         Top             =   945
         Width           =   1380
      End
      Begin VB.Label Label14 
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
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   2295
         TabIndex        =   37
         Top             =   225
         Width           =   435
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
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   8460
         TabIndex        =   36
         Top             =   225
         Width           =   450
      End
      Begin VB.Label Label11 
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
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   360
         TabIndex        =   35
         Top             =   225
         Width           =   600
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   8415
         TabIndex        =   34
         Top             =   450
         Width           =   2040
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   3420
         TabIndex        =   33
         Top             =   450
         Width           =   4920
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   8235
         TabIndex        =   32
         Top             =   4860
         Width           =   690
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   3
         X1              =   7425
         X2              =   8955
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCTO. %"
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
         Left            =   6570
         TabIndex        =   30
         Top             =   4950
         Width           =   1020
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCTO. BS."
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
         Left            =   6480
         TabIndex        =   28
         Top             =   4545
         Width           =   1185
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   7650
         TabIndex        =   27
         Top             =   5535
         Width           =   1275
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
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
         Left            =   6975
         TabIndex        =   26
         Top             =   5625
         Width           =   615
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   7650
         TabIndex        =   25
         Top             =   4050
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUB TOTAL"
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
         Left            =   6570
         TabIndex        =   24
         Top             =   4140
         Width           =   1050
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
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   945
         TabIndex        =   21
         Top             =   945
         Width           =   1365
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
         TabIndex        =   20
         Top             =   3465
         Width           =   75
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
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   135
         TabIndex        =   19
         Top             =   945
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
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   4995
         TabIndex        =   18
         Top             =   225
         Width           =   780
      End
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   420
      Left            =   11025
      TabIndex        =   22
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "VENTA"
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
      Left            =   11115
      TabIndex        =   23
      Top             =   135
      Width           =   1185
   End
   Begin VB.Image Image1 
      Height          =   60
      Left            =   5805
      Picture         =   "Egreso.frx":0015
      Stretch         =   -1  'True
      Top             =   495
      Width           =   6540
   End
End
Attribute VB_Name = "Egreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vcosto As Double
Dim vexi As Integer

Private Sub Combo1_GotFocus()
buscaprod
If siexiste = 1 Then
    Dim Cn As New ADODB.Connection
    Dim rsit As New ADODB.Recordset   ' Recordset de item
    Cn.ConnectionString = Cadena
    Cn.Open
    Text9.Text = UCase(Text9.Text)
    vcodpro = Text9.Text
        
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
        Text9.SetFocus
    End If
End If
End Sub



Private Sub Form_Load()
KeyPreview = True
creatablade
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

Private Sub SSCheck1_Click(Value As Integer)
If SSCheck1.Value = -1 Then
    Text3.Enabled = True
    Text9.Enabled = True
Else
    Text3.Text = ""
    Text9.Text = ""
    Text3.Enabled = False
    Text9.Enabled = False
    List1.Clear
End If
End Sub

Private Sub SSCommand2_Click()
If Len(Trim(Text3.Text)) > 0 Then
    'Verifica costo vs. precio venta
        If Val(Text3.Text) <= vcosto Then
                MsgBox "Precio de venta igual o menor a costo", vbCritical, empresa
                Text3.SetFocus
        Else
            If vexi >= Val(Text1.Text) Then
                Text2.Text = ""
                Text8.Text = ""
                Label7.Caption = ""
                grabadetventatemp
            Else
                MsgBox "La cantidad no esta disponible en la tienda", vbInformation, empresa
                Text1.SetFocus
            End If
        End If
ElseIf Len(Trim(Text4.Text)) > 0 And Len(Trim(Text3.Text)) = 0 Then
    'Verifica costo vs. precio venta
        If Val(Text3.Text) <= vcosto Then
                MsgBox "Precio de venta igual o menor a costo", vbCritical, empresa
                Text3.SetFocus
        Else
            If vexi >= Val(Text1.Text) Then
                Text2.Text = ""
                Text8.Text = ""
                Label7.Caption = ""
                grabadetventatemp
            Else
                MsgBox "La cantidad no esta disponible en la tienda", vbInformation, empresa
                Text1.SetFocus
            End If
        End If
Else
    MsgBox "Debe ingresar el precio de venta", vbInformation, empresa
    Text4.SetFocus
End If
End Sub

Private Sub SSCommand3_Click()
grabaegreso
limpiadatosf
creatablade
End Sub
Private Sub SSCommand6_Click()
Menup.Enabled = True
Unload Egreso
Set Egreso = Nothing
Menup.Label5.ForeColor = &HE0E0E0
End Sub
Private Sub SSOleDBCombo1_GotFocus()

SSOleDBCombo1.DroppedDown = True
End Sub
Private Sub SSOleDBGrid2_DblClick()
Adodc2.Recordset.Delete
Adodc2.Refresh
recalcula
End Sub
Private Sub Text1_GotFocus()
If Len(Trim(Text9.Text)) > 0 Then
    vcod_pro = Text9.Text
        
    Dim Cn As New ADODB.Connection
    Dim rsit As New ADODB.Recordset   ' Recordset de item
    Cn.ConnectionString = Cadena
    Cn.Open
        
    rsit.CursorType = adOpenKeyset
    rsit.LockType = adLockOptimistic
    rsit.ActiveConnection = Cn
    rsit.Source = "Select * from vproductotie Where Pro_Cod = " & "'" & vcod_pro & "'"
    rsit.Open
    
    If Not rsit.EOF Then
        Label2.Caption = rsit!Gru_des & " " & rsit!Pro_des
        Label8.Caption = rsit!Pro_Tip
'        Label13.Caption = rsit!ProTLi & ""
'        Label24.Caption = rsit!ProTNu & ""
        Text4.Text = Format(rsit!Pro_pve, "####.#0")
        Text3.Text = Format(rsit!Pro_POf, "####.#0")
        vexi = rsit!pro_exi
    Else
        MsgBox "Item inexistente", vbInformation, empresa
        Text9.SetFocus
    End If
       
    Text1.BackColor = &H9FBDEA
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
Else
    MsgBox "Debe seleccionar el producto", vbInformation, empresa
    Text9.SetFocus
End If
End Sub
Private Sub Text1_LostFocus()
Text1.BackColor = &HFFFFFF
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 46 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub Text2_GotFocus()
Text2.BackColor = &H9FBDEA
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub
Private Sub Text2_LostFocus()
Text2.Text = Format(Text2.Text, "####.#0")
If Val(Text2.Text) > 0 Then
    Text8.Text = "0.00"
    Label7.Caption = "0.00"
    Label20.Caption = Format(Val(Label18.Caption) - Val(Text2.Text), "####.#0")
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 46 Or KeyAscii > 57 Then
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
Private Sub Text3_DblClick()
origen = "venta"
Egreso.Enabled = False
Load EgresosAnt
EgresosAnt.Show
End Sub
Private Sub Text8_GotFocus()
Text8.BackColor = &H9FBDEA
Text8.SelStart = 0
Text8.SelLength = Len(Text8.Text)
End Sub
Private Sub Text8_LostFocus()
Text8.Text = Format(Text8.Text, "####.#0")
If Val(Text8.Text) > 0 Then
    Text2.Text = "0.00"
    Label7.Caption = Format((Label18.Caption * Text8.Text) / 100, "####.#0")
    Label20.Caption = Format(Val(Label18.Caption) - Val(Label7.Caption), "####.#0")
End If
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 46 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub


Private Sub Text4_GotFocus()
If Len(Trim(Text1.Text)) > 0 Then
    Text4.BackColor = &H80FFFF
    Text4.SelStart = 0
    Text4.SelLength = Len(Text4.Text)
Else
    MsgBox "Debe ingresar la cantidad", vbCritical, empresa
    Text1.SetFocus
End If
End Sub
Private Sub Text4_LostFocus()
Text4.BackColor = &HFFFFFF
Text4.Text = Format(Text4.Text, "####.#0")
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
Text5.Text = UCase(Text5.Text)
End Sub
Private Sub Text6_GotFocus()
Text6.BackColor = &H80FFFF
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

Private Function creatablade()
On Error GoTo errort

Dim TaUsu As String
vusuariot = Trim(vusucod)
TaUsu = "CREATE TABLE Tdetegreso" & vusuariot & "(" _
        & "cegr_Id int(5) DEFAULT NULL, " _
        & "degr_Id int(5) DEFAULT NULL, " _
        & "Pro_id int(5) DEFAULT NULL, " _
        & "Pro_cod varchar(30) DEFAULT NULL, " _
        & "Pro_Des varchar(250) DEFAULT NULL, " _
        & "degr_Can int(5) DEFAULT NULL, " _
        & "degr_PrU double(8,2) DEFAULT NULL, " _
        & "degr_PrT double(8,2) DEFAULT NULL, " _
        & "degr_Ppp double(8,2) DEFAULT NULL, " _
        & "degr_est varchar(1) DEFAULT NULL)"
        
Dim Cn As New ADODB.Connection
Cn.ConnectionString = Cadena
Cn.Open
Cn.Execute TaUsu
Adodc2.RecordSource = "Select * from Tdetegreso" & vusuariot
Adodc2.Refresh

borrac = "DELETE FROM cambio"
Cn.Execute borrac
Adodc1.Refresh

errort:
If Err.Number = -2147217900 Then
    borrat = "Drop Table Tdetegreso" & vusuariot
    Cn.Execute borrat
    creatablade
End If
End Function
Private Function grabadetventatemp()
Dim Cn As New ADODB.Connection
Dim rsdt As New ADODB.Recordset   ' Recordset de temporal de detalle
Dim rspr As New ADODB.Recordset   ' Recordset de productos
Cn.ConnectionString = Cadena
Cn.Open
vpro_cod = Text9.Text

rspr.CursorType = adOpenKeyset
rspr.LockType = adLockOptimistic
rspr.ActiveConnection = Cn
rspr.Source = "Select * from productostie Where Pro_cod = " & "'" & vpro_cod & "'"
rspr.Open

If Not rspr.EOF Then
    vpro_id = rspr!pro_id
Else
    vpro_id = 0
End If

vprocod = Text9.Text
vegr_can = Val(Text1.Text)
vpro_des = Label2.Caption & " " & Label8.Caption
If Len(Trim(Text3.Text)) > 0 Then
    vegr_pru = Val(Text3.Text)
ElseIf Len(Trim(Text4.Text)) > 0 And Len(Trim(Text3.Text)) = 0 Then
    vegr_pru = Val(Text4.Text)
End If
vtotal = vegr_can * vegr_pru
If SSCommand2.Caption = "Registrar" Then
    'Verifica si el item ya existe
    rsdt.CursorType = adOpenKeyset
    rsdt.LockType = adLockOptimistic
    rsdt.ActiveConnection = Cn
    rsdt.Source = "Select * from Tdetegreso" & vusuariot & " Where Pro_Cod = " & "'" & vprocod & "'"
    rsdt.Open

    If rsdt.EOF And SSCommand2.Caption = "Registrar" Then
        Text9.Enabled = False
        nuevoe = "Insert into Tdetegreso" & vusuariot & " SET Pro_id = " & vpro_id & ", Pro_Cod = " & "'" & vpro_cod & "'" & ", Pro_Des = " & "'" & vpro_des & "', degr_Can = " & vegr_can & _
        ", degr_pru = " & vegr_pru & ", degr_PrT = " & vtotal & ", degr_est = 'V'"
        Cn.Execute nuevoe
        Adodc2.Refresh
        Label18.Caption = Format(Val(Label18.Caption) + vtotal, "####.#0")
        Label20.Caption = Format(Label18.Caption, "####.#0")
        limpiadatos
        SSFrame2.Enabled = True
        Text9.Enabled = True
        Text9.SetFocus
    Else
        MsgBox "Item ya registrado", vbInformation, empresa
        limpiadatos
        Text9.SetFocus
    End If
 ElseIf SSCommand2.Caption = "Modificar" Then
    vpro_cod = SSOleDBGrid2.Columns(0).Value
    borrai = "Delete from Tdetinvinicial" & vusuariot & " Where pro_cod = " & "'" & vpro_cod & "'"
    Cn.Execute borrai
    
    nuevoe = "Insert into Tdetinvinicial" & vusuariot & " SET Pro_id = " & vpro_id & ", Pro_Cod = " & "'" & vpro_cod & "'" & ", Pro_Des = " & "'" & vpro_des & "', InI_Can = " & vini_can & _
    ", InI_PrU = " & vini_pru & ", InI_PrT = " & vIni_PrT & ", InI_Prv = " & vini_prv
    'nuevoe = "Insert into Tdetinvinicial" & vusuariot & " SET IteCod = " & vitecod & ", IteDes = " & "'" & vitedes & "', InICan = " & vIniCan & ", InIPrU = " & vIniPru & ", InIPrT = " & vIniPrT & ", InIPPP = " & vIniPru
    Cn.Execute nuevoe
    Adodc2.Refresh
    limpiadatos
    SSOleDBCombo1.SetFocus
    SSCommand2.Caption = "Registrar"
 End If
End Function
Private Function limpiadatos()
Text1.Text = ""
Text3.Text = ""
Text4.Text = ""
Label2.Caption = ""
Label8.Caption = ""
Text9.Text = ""
Combo1.Text = ""
End Function
Private Function limpiadatosf()
Text2.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Label7.Caption = ""
Label18.Caption = ""
Label20.Caption = ""
SSDateCombo1.Text = ""
End Function

Private Function grabaegreso()
Dim Cn As New ADODB.Connection
Dim rsit As New ADODB.Recordset   ' Recordset de item
Dim rsce As New ADODB.Recordset   ' Recordset de cabecera de egreso
Dim rsdt As New ADODB.Recordset   ' Recordset de detalle temporal
Cn.ConnectionString = Cadena
Cn.Open
    
rsit.CursorType = adOpenKeyset
rsit.LockType = adLockOptimistic
rsit.ActiveConnection = Cn
rsit.Source = "Select * from productos Where Pro_Cod = " & "'" & vcod_pro & "'"
rsit.Open

vini_fec = Format(SSDateCombo1.Text, "yyyy-mm-dd")

'Graba cabecera de egreso
grabacii = "INSERT INTO cabegreso SET cegr_fec = " & "'" & vini_fec & "', usu_id = " & vusucod
Cn.Execute grabacii

rsce.CursorType = adOpenKeyset
rsce.LockType = adLockOptimistic
rsce.ActiveConnection = Cn
rsce.Source = "Select * from cabegreso"
rsce.Open

If Not rsce.EOF Then
    rsce.MoveLast
    vini_id = rsce!cegr_id
Else
    vini_id = 1
End If

'Graba el id de cabecera en el detalle temporal
grabaid = "UPDATE Tdetegreso" & vusuariot & " SET cegr_Id = " & vini_id
Cn.Execute grabaid

'Graba detalle de egreso de temporal a real
detar = " INSERT INTO detegreso SELECT * from Tdetegreso" & vusuariot
Cn.Execute detar

'Graba Pago
vnit = Text5.Text
vras = Text6.Text
vfac = Text7.Text
vsut = Val(Label18.Caption)
vdbs = Val(Text2.Text)
vdpo = Val(Text8.Text)
vpag = Val(Label20.Caption)

grabac = "INSERT INTO pagoventa SET pag_fec = " & "'" & vini_fec & "', pag_nfa = " & "'" & vfac & "', pag_nit = " & "'" & vnit & _
"', pag_ras = " & "'" & vras & "', pag_sut = " & vsut & ", pag_dbs = " & vdbs & ", pag_dpo = " & vdpo & ", pag_mon = " & vpag & ", usu_id = " & vusucod & ", cegr_id = " & vini_id
Cn.Execute grabac

'Graba en tabla productos
rsdt.CursorType = adOpenKeyset
rsdt.LockType = adLockOptimistic
rsdt.ActiveConnection = Cn
rsdt.Source = "Select * from Tdetegreso" & vusuariot
rsdt.Open

If Not rsdt.EOF Then
    Do While Not rsdt.EOF
        vprocod = rsdt!pro_cod
        vini_can = rsdt!degr_can
        vini_pru = rsdt!degr_pru
        actpro = "UPDATE productos SET pro_exi = pro_exi - " & vini_can & " WHERE pro_cod = " & "'" & vprocod & "'"
        Cn.Execute actpro
        rsdt.MoveNext
    Loop
End If

'Graba en tabla de tienda
rsdt.Close
rsdt.CursorType = adOpenKeyset
rsdt.LockType = adLockOptimistic
rsdt.ActiveConnection = Cn
rsdt.Source = "Select * from Tdetegreso" & vusuariot
rsdt.Open

If Not rsdt.EOF Then
    Do While Not rsdt.EOF
        vprocod = rsdt!pro_cod
        vini_can = rsdt!degr_can
        vini_pru = rsdt!degr_pru
        actpro = "UPDATE tie SET procan = procan - " & vini_can & " WHERE procod = " & "'" & vprocod & "'"
        Cn.Execute actpro
        rsdt.MoveNext
    Loop
End If


Adodc2.Refresh
Cn.Close
End Function

Private Sub Text9_GotFocus()
Text9.BackColor = &H9FBDEA
Text9.SelStart = 0
Text9.SelLength = Len(Text9.Text)
End Sub

Private Function recalcula()
Dim Cn As New ADODB.Connection
Dim rsdt As New ADODB.Recordset   ' Recordset de detalle temporal
Cn.ConnectionString = Cadena
Cn.Open


rsdt.CursorType = adOpenKeyset
rsdt.LockType = adLockOptimistic
rsdt.ActiveConnection = Cn
rsdt.Source = "Select * from Tdetegreso" & vusuariot & " Where Pro_Des = " & "'" & vpro_des & "'"
rsdt.Open

If Not rsdt.EOF Then
    Do While Not rsdt.EOF
        total = total + rsdt!Degr_prt
        rsdt.MoveNext
    Loop
    Label18.Caption = Format(total, "####.#0")
    Label20.Caption = Format(total, "####.#0")
End If
Cn.Close
End Function

Private Sub Text9_LostFocus()
Text9.BackColor = &HFFFFFF
End Sub

Private Function buscaprod()
vcod_pro = Text9.Text
'vprotli = Combo1.Text
    
Dim Cn As New ADODB.Connection
Dim rsit As New ADODB.Recordset   ' Recordset de item
Cn.ConnectionString = Cadena
Cn.Open
    
rsit.CursorType = adOpenKeyset
rsit.LockType = adLockOptimistic
rsit.ActiveConnection = Cn
rsit.Source = "Select * from vproductotie Where Pro_Cod = " & "'" & vcod_pro & "'" 'AND proTLi = " & "'" & vprotli & "'"
rsit.Open

If Not rsit.EOF Then
    Label2.Caption = rsit!Gru_des & " " & rsit!Pro_des
    Label8.Caption = rsit!Pro_Tip
    Text4.Text = rsit!Pro_pve
    Text3.Text = rsit!Pro_POf
    siexiste = 1
Else
    MsgBox "Item inexistente", vbInformation, empresa
    Text9.SetFocus
End If
End Function
