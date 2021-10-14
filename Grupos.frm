VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Grupos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7005
   ClientLeft      =   11565
   ClientTop       =   4335
   ClientWidth     =   10155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   10155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   8955
      Top             =   2520
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
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2160
      Top             =   3015
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
      RecordSource    =   "Select * From grupos ORDER By gru_des"
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
      Left            =   8595
      TabIndex        =   0
      Top             =   720
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
      Left            =   8595
      TabIndex        =   4
      Top             =   3420
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
   Begin Threed.SSCommand SSCommand3 
      Height          =   420
      Left            =   8595
      TabIndex        =   5
      Top             =   1305
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
      Left            =   8595
      TabIndex        =   6
      Top             =   1845
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
      Left            =   135
      TabIndex        =   7
      Top             =   4050
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1614
      _Version        =   196608
      BackColor       =   14737632
      ClipControls    =   0   'False
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   135
         TabIndex        =   8
         Top             =   315
         Width           =   3480
      End
      Begin VB.Image Image3 
         Height          =   555
         Left            =   6165
         Picture         =   "Grupos.frx":0000
         Stretch         =   -1  'True
         Top             =   225
         Width           =   1815
      End
      Begin VB.Image Image2 
         Height          =   555
         Left            =   3825
         Picture         =   "Grupos.frx":0D07
         Stretch         =   -1  'True
         Top             =   225
         Width           =   1815
      End
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid1 
      Bindings        =   "Grupos.frx":1A4F
      Height          =   3255
      Left            =   1530
      TabIndex        =   9
      Top             =   720
      Width           =   5415
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
      Columns.Count   =   2
      Columns(0).Width=   1138
      Columns(0).Caption=   "No."
      Columns(0).Name =   "gru_id"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   1
      Columns(0).DataField=   "gru_id"
      Columns(0).DataType=   3
      Columns(0).FieldLen=   256
      Columns(1).Width=   7355
      Columns(1).Caption=   "Grupo"
      Columns(1).Name =   "gru_des"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "gru_des"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   9551
      _ExtentY        =   5741
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
   Begin Threed.SSFrame SSFrame1 
      Height          =   1770
      Left            =   135
      TabIndex        =   10
      Top             =   5040
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3122
      _Version        =   196608
      BackStyle       =   1
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.TextBox Text6 
         Height          =   330
         Left            =   135
         TabIndex        =   1
         Top             =   495
         Width           =   4650
      End
      Begin Threed.SSCommand SSCommand5 
         Height          =   420
         Left            =   6525
         TabIndex        =   11
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
         Left            =   5040
         TabIndex        =   3
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
         Left            =   5130
         TabIndex        =   2
         Top             =   495
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
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
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
         TabIndex        =   12
         Top             =   270
         Width           =   2355
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "GRUPOS"
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
      Left            =   8910
      TabIndex        =   13
      Top             =   135
      Width           =   1230
   End
   Begin VB.Image Image1 
      Height          =   60
      Left            =   3645
      Picture         =   "Grupos.frx":1A64
      Stretch         =   -1  'True
      Top             =   450
      Width           =   6540
   End
End
Attribute VB_Name = "Grupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vgru_id As String
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

Private Sub Image2_Click()
If Len(Trim(Text8.Text)) > 0 Then
    Dim db1 As String
    db1 = Text8.Text
    Adodc1.RecordSource = "SELECT * from grupos WHERE  gru_des LIKE " & "'%" & db1 & "%'"
    Adodc1.Refresh
End If
End Sub

Private Sub Image3_Click()
Text8.Text = ""
Adodc1.RecordSource = "Select * From grupos ORDER By gru_des"
Adodc1.Refresh
End Sub

Private Sub SSCommand1_Click()
limpiadatos
modo = "N"
ssFrame1.Enabled = True
SSCheck1.Value = -1
Text6.SetFocus
End Sub

Private Sub SSCommand2_Click()
Menup.Enabled = True
Unload Grupos
Set Grupos = Nothing
Menup.Label2.ForeColor = &HE0E0E0
End Sub

Private Sub SSCommand3_Click()
If seleccion = 1 Then
    modo = "M"
    ssFrame1.Enabled = True
    Text6.SetFocus
Else
    MsgBox "Debe seleccionar un grupo de la lista", vbInformation, empresa
End If
End Sub

Private Sub SSCommand4_Click()
'CrystalReport1.ReportFileName = App.Path & "\grupos.rpt"
'CrystalReport1.Action = 1
End Sub

Private Function verificagru()
Dim Cn As New ADODB.Connection
Dim rspr As New ADODB.Recordset   ' Recordset de grupos
Cn.ConnectionString = Cadena
Cn.Open

rspr.CursorType = adOpenKeyset
rspr.LockType = adLockOptimistic
rspr.ActiveConnection = Cn
rspr.Source = "Select * from grupos where gru_des = " & "'" & vgru_des & "'"
rspr.Open

If Not rspr.EOF Then
    siexiste = 1
Else
    siexiste = 0
End If
End Function

Private Sub SSCommand5_Click()
limpiadatos
End Sub
Private Sub SSCommand6_Click()
If Len(Trim(Text6.Text)) > 0 Then
    Dim Cn As New ADODB.Connection
    Cn.ConnectionString = Cadena
    Cn.Open
    
    vgru_des = Trim(Text6.Text)
    If SSCheck1.Value = -1 Then
        vgru_est = 1
    Else
        vgru_est = 0
    End If
    If modo = "N" Then
        verificagru
    Else
        siexiste = 0
    End If
    If siexiste = 0 Then
        If modo = "N" Then
            grabap = "INSERT INTO grupos SET gru_des = " & "'" & vgru_des & "', gru_est= " & vgru_est
        ElseIf modo = "M" Then
             grabap = "UPDATE grupos SET gru_des = " & "'" & vgru_des & "', gru_est= " & vgru_est & " WHERE gru_id = " & vgru_id
        End If
        Cn.Execute grabap
        Adodc1.Refresh
        limpiadatos
        Cn.Close
        MsgBox "Información registrada", vbInformation, empresa
    ElseIf siexiste = 1 Then
        MsgBox "Grupo existente. No puede grabar", vbInformation, empresa
        Text6.SetFocus
    End If
Else
    MsgBox "Debe ingresar nombre del grupo", vbInformation, empresa
    Text6.SetFocus
End If

End Sub

Private Sub SSOleDBGrid1_Click()
If Adodc1.Recordset.RecordCount > 0 Then
    seleccion = 1
    vgru_id = Adodc1.Recordset.Fields("gru_id")
    Text6.Text = Adodc1.Recordset.Fields("gru_des")
    
    If Adodc1.Recordset.Fields("gru_est") = 1 Then
        SSCheck1.Value = -1
    Else
        SSCheck1.Value = 0
    End If
End If

End Sub

Private Sub Text6_GotFocus()
Text6.BackColor = &H80FFFF
End Sub
Private Sub Text6_LostFocus()
Text6.BackColor = &HFFFFFF
Text6.Text = UCase(Text6.Text)
End Sub

Private Function limpiadatos()
Text6.Text = ""
SSCheck1.Value = 0
seleccion = 0
ssFrame1.Enabled = False
End Function
