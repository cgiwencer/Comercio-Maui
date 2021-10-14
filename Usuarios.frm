VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Usuarios 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7725
   ClientLeft      =   11040
   ClientTop       =   4635
   ClientWidth     =   10140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   10140
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   8955
      Top             =   2475
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
      RecordSource    =   "Select * From usuarios ORDER By Usu_Nom"
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
      Left            =   8595
      TabIndex        =   13
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
   Begin Threed.SSCommand SSCommand3 
      Height          =   420
      Left            =   8595
      TabIndex        =   15
      Top             =   1260
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
      TabIndex        =   16
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
      Left            =   135
      TabIndex        =   17
      Top             =   4005
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1614
      _Version        =   196608
      BackColor       =   14737632
      ClipControls    =   0   'False
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   135
         TabIndex        =   18
         Top             =   315
         Width           =   3480
      End
      Begin VB.Image Image2 
         Height          =   555
         Left            =   3825
         Picture         =   "Usuarios.frx":0000
         Stretch         =   -1  'True
         Top             =   225
         Width           =   1815
      End
      Begin VB.Image Image3 
         Height          =   555
         Left            =   6165
         Picture         =   "Usuarios.frx":0D48
         Stretch         =   -1  'True
         Top             =   225
         Width           =   1815
      End
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid1 
      Bindings        =   "Usuarios.frx":1A4F
      Height          =   3255
      Left            =   135
      TabIndex        =   14
      Top             =   675
      Width           =   8280
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
      Columns.Count   =   3
      Columns(0).Width=   5794
      Columns(0).Caption=   "Nombre"
      Columns(0).Name =   "Usu_Nom"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Usu_Nom"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4551
      Columns(1).Caption=   "Usuario"
      Columns(1).Name =   "Usu_Usu"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Usu_Usu"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Nivel"
      Columns(2).Name =   "Usu_Niv"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Usu_Niv"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   14605
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
      Height          =   2535
      Left            =   135
      TabIndex        =   1
      Top             =   4995
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4471
      _Version        =   196608
      BackStyle       =   1
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.TextBox Text6 
         Height          =   330
         Left            =   135
         TabIndex        =   2
         Top             =   495
         Width           =   3930
      End
      Begin VB.TextBox Text1 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   135
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1125
         Width           =   2850
      End
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   4185
         TabIndex        =   3
         Top             =   495
         Width           =   3480
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Usuarios.frx":1A64
         Left            =   135
         List            =   "Usuarios.frx":1A6E
         TabIndex        =   5
         Top             =   1890
         Width           =   2310
      End
      Begin Threed.SSCommand SSCommand5 
         Height          =   420
         Left            =   6210
         TabIndex        =   8
         Top             =   1800
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
         Left            =   4725
         TabIndex        =   7
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
         Caption         =   "&Grabar"
         BevelWidth      =   1
      End
      Begin Threed.SSCheck SSCheck1 
         Height          =   330
         Left            =   2835
         TabIndex        =   6
         Top             =   1890
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña"
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
         TabIndex        =   11
         Top             =   945
         Width           =   2355
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
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
         Left            =   4230
         TabIndex        =   10
         Top             =   270
         Width           =   2355
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel"
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
         TabIndex        =   9
         Top             =   1665
         Width           =   1140
      End
   End
   Begin VB.Image Image1 
      Height          =   60
      Left            =   3645
      Picture         =   "Usuarios.frx":1A8B
      Stretch         =   -1  'True
      Top             =   405
      Width           =   6540
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIOS"
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
      Left            =   8505
      TabIndex        =   19
      Top             =   90
      Width           =   1455
   End
End
Attribute VB_Name = "Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vusu_nom, vusu_usu, vusu_cla As String
Dim vusu_id As Integer
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
    Adodc1.RecordSource = "SELECT * from usuarios WHERE  usu_nom LIKE " & "'%" & db1 & "%'"
    Adodc1.Refresh
End If
End Sub

Private Sub Image3_Click()
Text8.Text = ""
Adodc1.RecordSource = "Select * From usuarios ORDER By usu_nom"
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
Unload Usuarios
Set Usuarios = Nothing
Menup.Label2.ForeColor = &HE0E0E0
End Sub

Private Sub SSCommand3_Click()
If seleccion = 1 Then
    modo = "M"
    ssFrame1.Enabled = True
    Text6.SetFocus
Else
    MsgBox "Debe seleccionar un usuario de la lista", vbInformation, empresa
End If
End Sub

Private Sub SSCommand4_Click()
'CrystalReport1.ReportFileName = App.Path & "\usuarios.rpt"
'CrystalReport1.Action = 1

End Sub

Private Sub SSCommand5_Click()
limpiadatos
End Sub

Private Sub SSCommand6_Click()
If Len(Trim(Text6.Text)) > 0 Then
    Dim Cn As New ADODB.Connection
    Cn.ConnectionString = Cadena
    Cn.Open
    
    vusu_nom = Trim(Text6.Text)
    vusu_usu = Text2.Text
    vusu_cla = Text1.Text
    vusu_niv = Combo1.Text
    If SSCheck1.Value = -1 Then
        vusu_est = 1
    Else
        vusu_est = 0
    End If
    If modo = "N" Then
        verificausu
    Else
        siexiste = 0
    End If
    If siexiste = 0 Then
        If modo = "N" Then
            grabap = "INSERT INTO usuarios SET usu_nom = " & "'" & vusu_nom & "', usu_usu = " & "'" & vusu_usu & "', usu_cla = " & "'" & vusu_cla & _
            "', usu_niv = " & "'" & vusu_niv & "', usu_est= " & vusu_est
        ElseIf modo = "M" Then
             grabap = "UPDATE usuarios SET usu_nom = " & "'" & vusu_nom & "', usu_usu = " & "'" & vusu_usu & "', usu_cla = " & "'" & vusu_cla & _
            "', usu_niv = " & "'" & vusu_niv & "', usu_est= " & vusu_est & " WHERE usu_id = " & vusu_id
        End If
        Cn.Execute grabap
        Adodc1.Refresh
        limpiadatos
        Cn.Close
        MsgBox "Información registrada", vbInformation, empresa
    ElseIf siexiste = 1 Then
        MsgBox "Usuario existente. No puede grabar", vbInformation, empresa
        Text6.SetFocus
    End If
Else
    MsgBox "Debe ingresar nombre del usuario", vbInformation, empresa
    Text6.SetFocus
End If

End Sub
Private Sub SSOleDBGrid1_Click()
If Adodc1.Recordset.RecordCount > 0 Then
    seleccion = 1
    vusu_id = Adodc1.Recordset.Fields("usu_id")
    Text6.Text = Adodc1.Recordset.Fields("usu_nom")
    Text2.Text = Adodc1.Recordset.Fields("usu_usu")
    Text1.Text = Adodc1.Recordset.Fields("usu_cla")
    Combo1.Text = Adodc1.Recordset.Fields("usu_niv")
    If Adodc1.Recordset.Fields("usu_est") = 1 Then
        SSCheck1.Value = ssCBChecked
    Else
        SSCheck1.Value = ssCBUnchecked
    End If
End If
End Sub
Private Function verificausu()
Dim Cn As New ADODB.Connection
Dim rspr As New ADODB.Recordset   ' Recordset de usuarios
Cn.ConnectionString = Cadena
Cn.Open

rspr.CursorType = adOpenKeyset
rspr.LockType = adLockOptimistic
rspr.ActiveConnection = Cn
rspr.Source = "Select * from usuarios where usu_nom = " & "'" & vusu_nom & "' AND usu_usu = " & "'" & vusu_usu & "'"
rspr.Open

If Not rspr.EOF Then
    siexiste = 1
Else
    siexiste = 0
End If
End Function
Private Function limpiadatos()
Text1.Text = ""
Text2.Text = ""
Text6.Text = ""
Combo1.Text = ""
SSCheck1.Value = 0
seleccion = 0
ssFrame1.Enabled = False
End Function

Private Sub Text6_GotFocus()
Text6.BackColor = &H80FFFF
End Sub
Private Sub Text6_LostFocus()
Text6.BackColor = &HFFFFFF
Text6.Text = UCase(Text6.Text)
End Sub

Private Sub Text1_GotFocus()
If Len(Trim(Text2.Text)) > 0 Then
    Text1.BackColor = &H80FFFF
Else
    MsgBox "Debe ingresar el nombre de usuario", vbInformation, empresa
    Text2.SetFocus
End If
End Sub
Private Sub Text1_LostFocus()
Text1.BackColor = &HFFFFFF
End Sub
Private Sub Text2_GotFocus()
If Len(Trim(Text6.Text)) > 0 Then
    Text2.BackColor = &H80FFFF
Else
    MsgBox "Debe ingresar el nombre de la persona", vbInformation, empresa
    Text6.SetFocus
End If
End Sub
Private Sub Text2_LostFocus()
Text2.BackColor = &HFFFFFF
End Sub




