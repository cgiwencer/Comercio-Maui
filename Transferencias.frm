VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form Transferencias 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8265
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   16050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   16050
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame SSFrame1 
      Height          =   5955
      Left            =   90
      TabIndex        =   6
      Top             =   1350
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   10504
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   225
         TabIndex        =   2
         Top             =   1080
         Width           =   2265
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   5535
         Top             =   1080
         Visible         =   0   'False
         Width           =   1905
         _ExtentX        =   3360
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
         RecordSource    =   "Select * from almacen order by AlmDes"
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
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4365
         TabIndex        =   3
         Top             =   2610
         Width           =   645
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "AlmDes"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   225
         TabIndex        =   0
         Top             =   450
         Width           =   4335
      End
      Begin Threed.SSCommand SSCommand6 
         Height          =   420
         Left            =   5220
         TabIndex        =   4
         Top             =   2565
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
         Caption         =   "&Transferir"
         BevelWidth      =   1
      End
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   330
         Left            =   3465
         Top             =   1080
         Visible         =   0   'False
         Width           =   1860
         _ExtentX        =   3281
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
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo1 
         Height          =   330
         Left            =   4725
         TabIndex        =   1
         Top             =   405
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
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
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
         Left            =   4770
         TabIndex        =   29
         Top             =   180
         Width           =   540
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3105
         TabIndex        =   28
         Top             =   2610
         Width           =   1050
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Existencia"
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
         Left            =   3150
         TabIndex        =   27
         Top             =   2385
         Width           =   885
      End
      Begin VB.Label Label9 
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
         Left            =   2700
         TabIndex        =   23
         Top             =   1575
         Width           =   780
      End
      Begin VB.Label Label7 
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
         Left            =   4410
         TabIndex        =   22
         Top             =   2385
         Width           =   465
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   225
         TabIndex        =   21
         Top             =   1800
         Width           =   2040
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   6570
         TabIndex        =   20
         Top             =   1800
         Width           =   1635
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   270
         TabIndex        =   19
         Top             =   855
         Width           =   600
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
         Left            =   6570
         TabIndex        =   18
         Top             =   1575
         Width           =   450
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   225
         TabIndex        =   17
         Top             =   2610
         Width           =   1050
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Talla Literal"
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
         Left            =   225
         TabIndex        =   16
         Top             =   2385
         Width           =   1020
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Talla  Numeral"
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
         Left            =   1665
         TabIndex        =   15
         Top             =   2385
         Width           =   1245
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1755
         TabIndex        =   14
         Top             =   2610
         Width           =   1050
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2250
         TabIndex        =   13
         Top             =   1800
         Width           =   4290
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Almacén Origen"
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
         TabIndex        =   10
         Top             =   180
         Width           =   1350
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   5955
      Left            =   8460
      TabIndex        =   8
      Top             =   1350
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   10504
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   225
         TabIndex        =   11
         Top             =   405
         Width           =   4335
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   2250
         Top             =   2250
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
         RecordSource    =   "Select * from dettraspaso where traid=0"
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
         Bindings        =   "Transferencias.frx":0000
         Height          =   4290
         Left            =   135
         TabIndex        =   24
         Top             =   900
         Width           =   7260
         _Version        =   196616
         ForeColorEven   =   0
         BackColorEven   =   14737632
         BackColorOdd    =   14737632
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   3200
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "ProCod"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "ProCod"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   7144
         Columns(1).Caption=   "Producto"
         Columns(1).Name =   "ProDes"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "ProDes"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1402
         Columns(2).Caption=   "Cant."
         Columns(2).Name =   "TraCan"
         Columns(2).Alignment=   1
         Columns(2).CaptionAlignment=   1
         Columns(2).DataField=   "TraCan"
         Columns(2).DataType=   3
         Columns(2).FieldLen=   256
         _ExtentX        =   12806
         _ExtentY        =   7567
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
      Begin Threed.SSCommand SSCommand1 
         Height          =   420
         Left            =   2880
         TabIndex        =   26
         Top             =   5310
         Width           =   2220
         _ExtentX        =   3916
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
         Caption         =   "&Realizar Transferencia"
         BevelWidth      =   1
      End
      Begin VB.Label Label5 
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
         Left            =   270
         TabIndex        =   12
         Top             =   180
         Width           =   1440
      End
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   420
      Left            =   14535
      TabIndex        =   25
      Top             =   7605
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "DESTINO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8460
      TabIndex        =   9
      Top             =   990
      Width           =   7485
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "ORIGEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   90
      TabIndex        =   7
      Top             =   990
      Width           =   8295
   End
   Begin VB.Image Image1 
      Height          =   60
      Left            =   9360
      Picture         =   "Transferencias.frx":0015
      Stretch         =   -1  'True
      Top             =   450
      Width           =   6540
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSFERENCIAS"
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
      Left            =   13545
      TabIndex        =   5
      Top             =   135
      Width           =   2310
   End
End
Attribute VB_Name = "Transferencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim valmo, valmd As Integer
Dim almdes, valmtip As String
Private Sub Combo1_GotFocus()
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
    Combo1.Clear
    Do While Not rsal.EOF
        Combo1.AddItem rsal!almdes
        rsal.MoveNext
    Loop
End If
Combo1.ListIndex = 0
Cn.Close
End Sub

Private Sub Combo2_GotFocus()
If Len(Trim(Combo1.Text)) > 0 Then
    valmori = Combo1.Text
    Dim Cn As New ADODB.Connection
    Dim rsal As New ADODB.Recordset   ' Recordset de Almacenes
    Cn.ConnectionString = Cadena
    Cn.Open
    
    rsal.CursorType = adOpenKeyset
    rsal.LockType = adLockOptimistic
    rsal.ActiveConnection = Cn
    rsal.Source = "Select * from Almacen ORDER BY AlmDes"
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
Else
    MsgBox "No existe almacen de origen", vbInformation, empresa
    Combo1.SetFocus
End If
End Sub

Private Sub Form_Load()
KeyPreview = True
creatablatr
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
Dim Cn As New ADODB.Connection
Dim rsca As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open

'Graba cabecera de Traspaso
vtrafec = Format(SSDateCombo1, "yyyy-mm-dd")
vtraori = Combo1.Text
vtrades = Combo2.Text

grabacab = "INSERT INTO cabtraspaso SET trafec = " & "'" & vtrafec & "', traori = " & "'" & vtraori & _
"', trades = " & "'" & vtrades & "'"
Cn.Execute grabacab

rsca.CursorType = adOpenKeyset
rsca.LockType = adLockOptimistic
rsca.ActiveConnection = Cn
rsca.Source = "Select * from cabtraspaso"
rsca.Open

rsca.MoveLast

vtraid = rsca!traid
rsca.Close

'Graba detalle de temporal a real
grabaid = "UPDATE Tdettraspaso" & vusuariot & " SET TraID = " & vtraid
Cn.Execute grabaid

grabatar = "INSERT INTO dettraspaso Select * from Tdettraspaso" & vusuariot
Cn.Execute grabatar

'Actualiza tabla de productos
rsca.CursorType = adOpenKeyset
rsca.LockType = adLockOptimistic
rsca.ActiveConnection = Cn
rsca.Source = "Select * from Tdettraspaso" & vusuariot
rsca.Open

If Not rsca.EOF Then
    Do While Not rsca.EOF
        vprodcod = rsca!procod
        vtracan = rsca!TraCan
        'Actualiza tabla de productos
        actpro = "UPDATE productos SET Pro_Exi = Pro_exi - " & vtracan & " WHERE Pro_Cod = " & "'" & vprodcod & "'"
        Cn.Execute actpro
        'Actualiza tabla de productos tienda
        acttie = "UPDATE productostie SET Pro_Exi = Pro_exi + " & vtracan & " WHERE Pro_Cod = " & "'" & vprodcod & "'"
        Cn.Execute acttie
        rsca.MoveNext
    Loop
    MsgBox "Se realizó la transferencia", vbInformation, empresa
    SSCommand2_Click
End If
rsca.Close
End Sub

Private Sub SSCommand2_Click()
Unload Transferencias
Set Transferencias = Nothing
End Sub

Private Sub SSCommand6_Click()
If Len(Trim(Text1.Text)) > 0 Then
    If Val(Text1.Text) <= Val(Label15.Caption) Then
        Dim Cn As New ADODB.Connection
        Cn.ConnectionString = Cadena
        Cn.Open
                        
        vprocod = Text2.Text
        vprodes = Label6.Caption & " " & Label25.Caption
        vtracan = Text1.Text
        
        Traspasa = "INSERT INTO Tdettraspaso" & vusuariot & " SET ProCod = " & "'" & vprocod & "', ProDes = " & "'" & vprodes & _
        "', TraCan = " & vtracan
        Cn.Execute Traspasa
        Adodc2.RecordSource = "SELECT * from Tdettraspaso" & vusuariot
        Adodc2.Refresh
        limpiadatos
        Text2.SetFocus
    Else
        MsgBox "La cantidad es mayor a la existencia", vbInformation, empresa
        Text1.SetFocus
    End If
Else
    MsgBox "Debe ingresar la cantidad", vbInformation, empresa
    Text1.SetFocus
End If
End Sub

Private Sub SSOleDBCombo1_GotFocus()
Dim Cn As New ADODB.Connection
Cn.ConnectionString = Cadena
Cn.Open

If Len(Trim(Combo1.Text)) > 0 Then
    vproubi = Combo1.Text
    
    almacenorigen = "SELECT * FROM almacen WHERE     "
    
    
    
    Adodc4.RecordSource = "Select * from productos where Pro_Ubi = " & "'" & vproubi & "'"
    Adodc4.Refresh
    SSOleDBCombo1.DroppedDown = True
Else
    MsgBox "Debe seleccionar un almacen de origen", vbInformation, empresa
    Combo1.SetFocus
End If
End Sub
Private Sub Text1_GotFocus()
If Len(Trim(Text2.Text)) > 0 Then
    vcod_pro = Text2.Text
        
    Dim Cn As New ADODB.Connection
    Dim rsit As New ADODB.Recordset   ' Recordset de item
    Cn.ConnectionString = Cadena
    Cn.Open
        
    'Verifica existencia en detalle temporal
    rsit.CursorType = adOpenKeyset
    rsit.LockType = adLockOptimistic
    rsit.ActiveConnection = Cn
    rsit.Source = "Select * from Tdettraspaso" & vusuariot & " Where ProCod = " & "'" & vcod_pro & "'"
    rsit.Open
    
    If Not rsit.EOF Then
        rsit.Close
        MsgBox "Producto ya registrado en esta traspaso", vbInformation, empresa
        Text2.SetFocus
    Else
        rsit.Close
        rsit.CursorType = adOpenKeyset
        rsit.LockType = adLockOptimistic
        rsit.ActiveConnection = Cn
        If valmtip = "ALMACEN GENERAL" Then
            rsit.Source = "Select * from vproducto Where Pro_Cod = " & "'" & vcod_pro & "'"
        ElseIf valmtip = "ALMACEN SECUNDARIO" Then
            rsit.Source = "Select * from productostie Where Pro_Cod = " & "'" & vcod_pro & "'"
        End If
        rsit.Open
        
        If Not rsit.EOF Then
            Label6.Caption = rsit!Gru_des
            Label25.Caption = rsit!Pro_des
            Label8.Caption = rsit!Pro_Tip
            Label13.Caption = rsit!ProTLi & ""
            Label24.Caption = rsit!ProTNu & ""
            Label15.Caption = rsit!pro_exi & ""
        Else
            MsgBox "Item inexistente", vbInformation, empresa
            Text1.SetFocus
        End If
           
        Text1.BackColor = &H9FBDEA
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
    End If
Else
    MsgBox "Debe seleccionar el código de producto", vbInformation, empresa
    Text2.SetFocus
End If

End Sub
Private Function creatablatr()
On Error GoTo errort

Dim TaUsu As String
vusuariot = Trim(vusucod)
TaUsu = "CREATE TABLE Tdettraspaso" & vusuariot & "(" _
        & "Traid INT(5) DEFAULT NULL, " _
        & "Procod varchar(40) DEFAULT NULL, " _
        & "ProDes varchar(250) DEFAULT NULL, " _
        & "Tracan int(4) DEFAULT NULL)" _

'TaUsu = "CREATE TABLE Tdettraspaso" & vusuariot & "(" _
'        & "Pro_Id int(5) DEFAULT NULL, " _
'        & "Pro_Des varchar(250) DEFAULT NULL, " _
'        & "Pro_Uni varchar(150) DEFAULT NULL, " _
'        & "Pro_Tip varchar(100) DEFAULT NULL, " _
'        & "Pro_ExMi int(5) DEFAULT NULL, " _
'        & "Pro_ExMa int(5) DEFAULT NULL, " _
'        & "Pro_Ubi varchar(200) DEFAULT NULL, " _
'        & "Pro_CodB varchar(30) DEFAULT NULL, " _
'        & "Pro_Cod varchar(40) DEFAULT NULL, " _
'        & "ProTNu varchar(2) DEFAULT NULL, " _
'        & "ProTLi varchar(5) DEFAULT NULL, " _
'        & "Gru_Id int(3) DEFAULT NULL, " _
'        & "Mar_Id int(4) DEFAULT NULL, " _
'        & "Pro_Est int(1) DEFAULT NULL, " _
'        & "Pro_Aut int(1) DEFAULT NULL, " _
'        & "Pro_PrC double(8,2) DEFAULT NULL, " _
'        & "Pro_Sal int(6) DEFAULT NULL, " _
'        & "Prv_Id int(6) DEFAULT NULL, " _
'        & "Pro_Pve double(8,2) DEFAULT NULL, " _
'        & "Pro_Exi double(8,2) DEFAULT NULL, Pro_ppp double(8,2) DEFAULT NULL, Usu_id Int(3) DEFAULT NULL, ProUni varchar(400) DEFAULT NULL, " _
'        & "ProPOf double(8,0) DEFAULT NULL, ProVen Int(4) DEFAULT NULL, TraId Int(5) DEFAULT NULL)"
        
Dim Cn As New ADODB.Connection
Cn.ConnectionString = Cadena
Cn.Open
Cn.Execute TaUsu
Adodc2.RecordSource = "Select * from Tdettraspaso" & vusuariot
Adodc2.Refresh

errort:
If Err.Number = -2147217900 Then
    borrat = "Drop Table Tdettraspaso" & vusuariot
    Cn.Execute borrat
    creatablatr
End If
End Function
Private Sub Text2_LostFocus()
Text2.BackColor = &HFFFFFF
Text2.Text = UCase(Text2.Text)
End Sub
Private Sub Text2_GotFocus()
If Len(Trim(Combo1.Text)) > 0 Then
    
    If Combo1.Text = "ALMACEN GENERAL" Then
        Combo2.Text = "TIENDA"
    Else
            Combo2.Text = "ALMACEN GENERAL"
    End If
    Text2.BackColor = &H80FFFF
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    
    valmdes = Combo1.Text
    Dim Cn As New ADODB.Connection
    Dim rsal As New ADODB.Recordset   ' Recordset de almacen
    Cn.ConnectionString = Cadena
    Cn.Open
        
    rsal.CursorType = adOpenKeyset
    rsal.LockType = adLockOptimistic
    rsal.ActiveConnection = Cn
    rsal.Source = "Select * from almacen Where AlmDes = " & "'" & valmdes & "'"
    rsal.Open
    
    valmtip = rsal!almtip
    Cn.Close
    
Else
    MsgBox "Debe seleccionar el almacén de origen", vbInformation, empresa
    Combo1.SetFocus
End If
End Sub
Private Function limpiadatos()
Label6.Caption = ""
Label8.Caption = ""
Label13.Caption = ""
Label15.Caption = ""
Label24.Caption = ""
Label25.Caption = ""
Text1.Text = ""
Text2.Text = ""
End Function
