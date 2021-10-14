VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Arqueo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8670
   ClientLeft      =   4980
   ClientTop       =   3255
   ClientWidth     =   14385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   14385
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3780
      Top             =   7830
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4950
      Top             =   5310
      Width           =   2310
      _ExtentX        =   4075
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
      RecordSource    =   "Select * from arqueo"
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
      Left            =   13140
      TabIndex        =   0
      Top             =   6930
      Width           =   1185
      _ExtentX        =   2090
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
   Begin Threed.SSFrame SSFrame2 
      Height          =   1515
      Left            =   180
      TabIndex        =   2
      Top             =   855
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   2672
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo2 
         Height          =   375
         Left            =   1995
         TabIndex        =   3
         Top             =   615
         Width           =   1545
         _Version        =   65537
         _ExtentX        =   2725
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483633
      End
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo1 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   615
         Width           =   1635
         _Version        =   65537
         _ExtentX        =   2884
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483633
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   420
         Left            =   3690
         TabIndex        =   7
         Top             =   585
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
         Caption         =   "&Arqueo"
         BevelWidth      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final"
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
         Left            =   2190
         TabIndex        =   6
         Top             =   390
         Width           =   1005
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicial"
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
         Left            =   420
         TabIndex        =   5
         Top             =   390
         Width           =   1110
      End
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid2 
      Bindings        =   "Arqueo.frx":0000
      Height          =   4650
      Left            =   585
      TabIndex        =   8
      Top             =   2835
      Width           =   12450
      _Version        =   196616
      SelectTypeRow   =   1
      ForeColorEven   =   0
      BackColorEven   =   14737632
      BackColorOdd    =   14737632
      RowHeight       =   423
      Columns.Count   =   9
      Columns(0).Width=   1905
      Columns(0).Caption=   "Fecha"
      Columns(0).Name =   "Pag_fec"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   1
      Columns(0).DataField=   "Pag_fec"
      Columns(0).DataType=   7
      Columns(0).NumberFormat=   "dd-mm-yyyy"
      Columns(0).FieldLen=   256
      Columns(1).Width=   1693
      Columns(1).Caption=   "No. Fact."
      Columns(1).Name =   "Pag_NFa"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Pag_NFa"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2302
      Columns(2).Caption=   "Nit"
      Columns(2).Name =   "Pag_Nit"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Pag_Nit"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   5609
      Columns(3).Caption=   "Cliente"
      Columns(3).Name =   "Pag_RaS"
      Columns(3).CaptionAlignment=   0
      Columns(3).DataField=   "Pag_RaS"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1799
      Columns(4).Caption=   "Sub. Total"
      Columns(4).Name =   "Pag_Sut"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   1
      Columns(4).DataField=   "Pag_Sut"
      Columns(4).DataType=   5
      Columns(4).FieldLen=   256
      Columns(5).Width=   1905
      Columns(5).Caption=   "Dscto. Bs."
      Columns(5).Name =   "Pag_DBs"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   1
      Columns(5).DataField=   "Pag_DBs"
      Columns(5).DataType=   5
      Columns(5).FieldLen=   256
      Columns(6).Width=   2064
      Columns(6).Caption=   "Dscto. %"
      Columns(6).Name =   "Pag_Dpo"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   1
      Columns(6).DataField=   "Pag_Dpo"
      Columns(6).DataType=   5
      Columns(6).FieldLen=   256
      Columns(7).Width=   1826
      Columns(7).Caption=   "Monto"
      Columns(7).Name =   "Pag_Mon"
      Columns(7).Alignment=   1
      Columns(7).CaptionAlignment=   1
      Columns(7).DataField=   "Pag_Mon"
      Columns(7).DataType=   5
      Columns(7).NumberFormat=   "#####.#0"
      Columns(7).FieldLen=   256
      Columns(8).Width=   1746
      Columns(8).Caption=   "Devolución"
      Columns(8).Name =   "Pag_Dev"
      Columns(8).Alignment=   1
      Columns(8).CaptionAlignment=   1
      Columns(8).DataField=   "Pag_dev"
      Columns(8).DataType=   5
      Columns(8).NumberFormat=   "#####.#0"
      Columns(8).FieldLen=   256
      _ExtentX        =   21960
      _ExtentY        =   8202
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
      Left            =   5580
      TabIndex        =   9
      Top             =   1440
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
      Caption         =   "&Imprimir"
      BevelWidth      =   1
   End
   Begin VB.Label Label4 
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
      Left            =   11655
      TabIndex        =   13
      Top             =   7920
      Width           =   1140
   End
   Begin VB.Label Label3 
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
      Left            =   11835
      TabIndex        =   12
      Top             =   7515
      Width           =   960
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
      Left            =   9855
      TabIndex        =   11
      Top             =   7605
      Width           =   615
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
      Left            =   10530
      TabIndex        =   10
      Top             =   7515
      Width           =   1275
   End
   Begin VB.Image Image1 
      Height          =   60
      Left            =   7785
      Picture         =   "Arqueo.frx":0015
      Stretch         =   -1  'True
      Top             =   495
      Width           =   6540
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ARQUEO DE CAJA"
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
      Left            =   12060
      TabIndex        =   1
      Top             =   180
      Width           =   2310
   End
End
Attribute VB_Name = "Arqueo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
KeyPreview = True
Dim Cn As New ADODB.Connection
Cn.ConnectionString = Cadena
Cn.Open

borraar = "DELETE FROM arqueo"
Cn.Execute borraar
Adodc1.Refresh
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
vfechaini = SSDateCombo1.Text
vfechafin = SSDateCombo2.Text

CrystalReport1.ReportFileName = App.Path & "\arqueo.rpt"
CrystalReport1.Formulas(0) = "a = " & "'" & vfechaini & "'"
CrystalReport1.Formulas(1) = "de = " & "'" & vfechafin & "'"
CrystalReport1.Action = 1

End Sub

Private Sub SSCommand2_Click()
Dim Cn As New ADODB.Connection
Dim rsar As New ADODB.Recordset ' Recordset de arqueo
Cn.ConnectionString = Cadena
Cn.Open
vfechai = SSDateCombo1.Date
vfechaf = SSDateCombo2.Date
vfechai = Format(vfechai, "YYYY-MM-dd")
vfechaf = Format(vfechaf, "yyyy-mm-dd")

borrat = "DELETE FROM arqueo"
Cn.Execute borrat

grabaarqueo = "INSERT INTO arqueo select * from pagoventa where Pag_Fec BETWEEN " & "'" & vfechai & "'" & " AND " & "'" & vfechaf & "'"
Cn.Execute grabaarqueo

rsar.CursorType = adOpenKeyset
rsar.LockType = adLockOptimistic
rsar.ActiveConnection = Cn
rsar.Source = "Select * from arqueo"
rsar.Open

If Not rsar.EOF Then
    Do While Not rsar.EOF
        If rsar!PAg_mon > 0 Then
        
        Else
            rsar!PAg_mon = 0
            rsar.Update
        End If
        If rsar!Pag_dev > 0 Then
        
        Else
            rsar!Pag_dev = 0
            rsar.Update
        End If
        rsar.MoveNext
    Loop
End If
rsar.Close

rsar.CursorType = adOpenKeyset
rsar.LockType = adLockOptimistic
rsar.ActiveConnection = Cn
rsar.Source = "Select * from arqueo"
rsar.Open

If Not rsar.EOF Then
    Do While Not rsar.EOF
       
            sumaven = sumaven + rsar!PAg_mon
       
            sumadev = sumadev + rsar!Pag_dev
       
        rsar.MoveNext
    Loop
End If

Label20.Caption = Format(sumaven, "####.#0")
Label3.Caption = Format(sumadev, "####.#0")
Label4.Caption = Format(Val(Label20.Caption) - Val(Label3.Caption), "####.#0")
Adodc1.Refresh
Cn.Close
End Sub

Private Sub SSCommand6_Click()
Menup.Enabled = True
Menup.Label7.ForeColor = &H8000000F
Unload Arqueo
Set Arqueo = Nothing
End Sub

Private Sub SSOleDBGrid2_DblClick()
vegr_id = Adodc1.Recordset.Fields("cegr_id")
Dim Cn As New ADODB.Connection
Dim rsve As New ADODB.Recordset   ' Recordset de det det venta
Cn.ConnectionString = Cadena
Cn.Open
    
rsve.CursorType = adOpenKeyset
rsve.LockType = adLockOptimistic
rsve.ActiveConnection = Cn
rsve.Source = "Select * from vventa1 Where cegr_id = " & vegr_id
rsve.Open

If Not rsve.EOF Then
    Load Detventa
    Detventa.SSDateCombo1.Text = rsve!pag_fec
    Detventa.Label8.Caption = rsve!pag_nit
    Detventa.Label1.Caption = rsve!pag_ras
    Detventa.Label2.Caption = rsve!pag_NFa
    Detventa.Adodc1.RecordSource = "Select * from vventa1 where cegr_id = " & vegr_id
    Detventa.Adodc1.Refresh
    Detventa.Show
End If
End Sub

