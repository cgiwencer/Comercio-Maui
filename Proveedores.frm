VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Proveedores 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8760
   ClientLeft      =   10425
   ClientTop       =   4515
   ClientWidth     =   11340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   11340
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   10080
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
   Begin Threed.SSFrame SSFrame1 
      Height          =   3615
      Left            =   90
      TabIndex        =   14
      Top             =   4995
      Width           =   11130
      _ExtentX        =   19632
      _ExtentY        =   6376
      _Version        =   196608
      BackStyle       =   1
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.TextBox Text9 
         Height          =   330
         Left            =   8955
         TabIndex        =   3
         Top             =   495
         Width           =   2040
      End
      Begin VB.TextBox Text7 
         Height          =   330
         Left            =   6030
         TabIndex        =   9
         Top             =   2655
         Width           =   4605
      End
      Begin VB.TextBox Text5 
         Height          =   330
         Left            =   135
         TabIndex        =   8
         Top             =   2655
         Width           =   5820
      End
      Begin VB.TextBox Text4 
         Height          =   330
         Left            =   2520
         TabIndex        =   7
         Top             =   1890
         Width           =   5820
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Proveedores.frx":0000
         Left            =   135
         List            =   "Proveedores.frx":000A
         TabIndex        =   6
         Top             =   1890
         Width           =   2310
      End
      Begin VB.TextBox Text3 
         Height          =   330
         Left            =   6030
         TabIndex        =   5
         Top             =   1170
         Width           =   4965
      End
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   6030
         TabIndex        =   2
         Top             =   495
         Width           =   2805
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   135
         TabIndex        =   4
         Top             =   1170
         Width           =   5820
      End
      Begin VB.TextBox Text6 
         Height          =   330
         Left            =   135
         TabIndex        =   1
         Top             =   495
         Width           =   5820
      End
      Begin Threed.SSCommand SSCommand5 
         Height          =   420
         Left            =   9630
         TabIndex        =   18
         Top             =   3105
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
         Left            =   8145
         TabIndex        =   10
         Top             =   3105
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
         Left            =   9810
         TabIndex        =   27
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
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "NIT"
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
         Left            =   8910
         TabIndex        =   29
         Top             =   270
         Width           =   2355
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
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
         Left            =   6030
         TabIndex        =   26
         Top             =   2430
         Width           =   2355
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
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
         TabIndex        =   25
         Top             =   2430
         Width           =   2355
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Correo electrónico"
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
         Left            =   2520
         TabIndex        =   24
         Top             =   1665
         Width           =   2355
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Origen"
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
         TabIndex        =   23
         Top             =   1665
         Width           =   1140
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Persona de contacto"
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
         Left            =   6030
         TabIndex        =   22
         Top             =   945
         Width           =   2355
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfonos"
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
         Left            =   5985
         TabIndex        =   21
         Top             =   270
         Width           =   2355
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección"
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
         TabIndex        =   20
         Top             =   945
         Width           =   2355
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Razón Social"
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
         TabIndex        =   19
         Top             =   270
         Width           =   2355
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
      Connect         =   "DSN=ferreteria"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ferreteria"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "cagisa"
      RecordSource    =   "Select * From proveedores ORDER By prv_des"
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
      Left            =   9720
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
      Left            =   9720
      TabIndex        =   12
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
      Bindings        =   "Proveedores.frx":0024
      Height          =   3255
      Left            =   90
      TabIndex        =   13
      Top             =   630
      Width           =   9405
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
      Columns.Count   =   4
      Columns(0).Width=   5715
      Columns(0).Caption=   "Razón Social"
      Columns(0).Name =   "Prv_des"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Prv_des"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Telefonos"
      Columns(1).Name =   "Prv_Tel"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Prv_Tel"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   4524
      Columns(2).Caption=   "Contacto"
      Columns(2).Name =   "Prv_Cont"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Prv_Cont"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2117
      Columns(3).Caption=   "Origen"
      Columns(3).Name =   "Prv_Ori"
      Columns(3).CaptionAlignment=   0
      Columns(3).DataField=   "Prv_Ori"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      _ExtentX        =   16589
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
   Begin Threed.SSCommand SSCommand3 
      Height          =   420
      Left            =   9720
      TabIndex        =   15
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
      Left            =   9720
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
      Left            =   90
      TabIndex        =   17
      Top             =   4005
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   1614
      _Version        =   196608
      BackColor       =   14737632
      ClipControls    =   0   'False
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   135
         TabIndex        =   28
         Top             =   315
         Width           =   3480
      End
      Begin VB.Image Image3 
         Height          =   555
         Left            =   6165
         Picture         =   "Proveedores.frx":0039
         Stretch         =   -1  'True
         Top             =   225
         Width           =   1815
      End
      Begin VB.Image Image2 
         Height          =   555
         Left            =   3825
         Picture         =   "Proveedores.frx":0D40
         Stretch         =   -1  'True
         Top             =   225
         Width           =   1815
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDORES"
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
      Left            =   9270
      TabIndex        =   11
      Top             =   90
      Width           =   1950
   End
   Begin VB.Image Image1 
      Height          =   60
      Left            =   4860
      Picture         =   "Proveedores.frx":1A88
      Stretch         =   -1  'True
      Top             =   405
      Width           =   6540
   End
End
Attribute VB_Name = "Proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vprv_id As Integer
Dim vprv_des As String
Private Sub Combo1_GotFocus()
Combo1.BackColor = &H80FFFF
Combo1.ListIndex = 0
End Sub
Private Sub Combo1_LostFocus()
Combo1.BackColor = &HFFFFFF
Combo1.Text = UCase(Combo1.Text)
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
Private Sub Image2_Click()
If Len(Trim(Text8.Text)) > 0 Then
    Dim db1 As String
    db1 = Text8.Text
    Adodc1.RecordSource = "SELECT * from proveedores WHERE  prv_des LIKE " & "'%" & db1 & "%'"
    Adodc1.Refresh
End If
End Sub

Private Sub Image3_Click()
Text8.Text = ""
Adodc1.RecordSource = "Select * From proveedores ORDER By prv_des"
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
If origen = "proveedores" Then
    Menup.Enabled = True
    Menup.Label2.ForeColor = &HE0E0E0
ElseIf origen = "ingreso" Then
    Ingresos.Enabled = True
Else
    Productos.Enabled = True
End If

Unload Proveedores
Set Proveedores = Nothing

End Sub

Private Sub SSCommand3_Click()
If seleccion = 1 Then
    modo = "M"
    ssFrame1.Enabled = True
Else
    MsgBox "Debe seleccionar un proveedor de la lista", vbInformation, empresa
End If
End Sub

Private Sub SSCommand4_Click()
CrystalReport1.ReportFileName = App.Path & "\proveedores.rpt"
CrystalReport1.Action = 1
End Sub

Private Sub SSCommand5_Click()
limpiadatos
End Sub

Private Sub SSCommand6_Click()
If Len(Trim(Text6.Text)) > 0 Then
    Dim Cn As New ADODB.Connection
    Cn.ConnectionString = Cadena
    Cn.Open
    
    vprv_des = Trim(Text6.Text)
    If modo = "N" Then
        verificapro
    Else
        siexiste = 0
    End If
    If siexiste = 0 Then
        vprv_tel = Text2.Text
        vprv_dir = Text1.Text
        vprv_nit = Text9.Text
        vprv_cont = Text3.Text
        vprv_ori = Combo1.Text
        vprv_email = Text4.Text
        vprv_ban = Text5.Text
        vprv_cue = Text7.Text
        If SSCheck1.Value = -1 Then
            vprv_est = 1
        Else
            vprv_est = 0
        End If
        
        If modo = "N" Then
            grabap = "INSERT INTO proveedores SET prv_des = " & "'" & vprv_des & "', prv_tel = " & "'" & vprv_tel & "', prv_dir = " & "'" & vprv_dir & _
            "', prv_cont = " & "'" & vprv_cont & "', prv_ori = " & "'" & vprv_ori & "', prv_email = " & "'" & vprv_email & "', prv_ban = " & "'" & vprv_ban & _
            "', prv_nit = " & "'" & vprv_nit & "', prv_cue = " & "'" & vprv_cue & "', prv_est = " & vprv_est & ", usu_id = " & vusucod
        ElseIf modo = "M" Then
            grabap = "UPDATE proveedores SET prv_des = " & "'" & vprv_des & "', prv_tel = " & "'" & vprv_tel & "', prv_dir = " & "'" & vprv_dir & _
            "', prv_cont = " & "'" & vprv_cont & "', prv_ori = " & "'" & vprv_ori & "', prv_email = " & "'" & vprv_email & "', prv_ban = " & "'" & vprv_ban & _
            "', prv_nit = " & "'" & vprv_nit & "', prv_cue = " & "'" & vprv_cue & "', prv_est = " & vprv_est & " WHERE prv_id = " & vprv_id
        End If
        Cn.Execute grabap
        
        MsgBox "Información registrada", vbInformation, empresa
        Adodc1.Refresh
        limpiadatos
        Cn.Close
        If origen = "productos" Then
            Productos.Enabled = True
            Unload Proveedores
            Set Proveedores = Nothing
        ElseIf origen = "proveedores" Then
            Menup.Enabled = True
            Unload Proveedores
            Set Proveedores = Nothing
        End If
    ElseIf siexiste = 1 Then
        MsgBox "Nombre de proveedor existente. No puede grabar", vbInformation, empresa
        Text6.SetFocus
    End If
Else
    MsgBox "Debe ingresar la Razón Social del proveedor", vbInformation, empresa
    Text6.SetFocus
End If
End Sub

Private Sub SSOleDBGrid1_Click()
If Adodc1.Recordset.RecordCount > 0 Then
    seleccion = 1
    vprv_id = Adodc1.Recordset.Fields("prv_id")
    Text6.Text = Adodc1.Recordset.Fields("prv_des")
    Text2.Text = Adodc1.Recordset.Fields("prv_tel") & ""
    Text1.Text = Adodc1.Recordset.Fields("prv_dir") & ""
    Text3.Text = Adodc1.Recordset.Fields("prv_cont") & ""
    Combo1.Text = Adodc1.Recordset.Fields("prv_ori") & ""
    Text4.Text = Adodc1.Recordset.Fields("prv_email") & ""
    Text5.Text = Adodc1.Recordset.Fields("prv_ban") & ""
    Text7.Text = Adodc1.Recordset.Fields("prv_cue") & ""
    Text9.Text = Adodc1.Recordset.Fields("prv_nit") & ""
    If Adodc1.Recordset.Fields("prv_est") = 1 Then
        SSCheck1.Value = ssCBChecked
    Else
        SSCheck1.Value = ssCBUnchecked
    End If
End If
End Sub
Private Sub Text1_GotFocus()
Text1.BackColor = &H80FFFF
End Sub
Private Sub Text1_LostFocus()
Text1.BackColor = &HFFFFFF
Text1.Text = UCase(Text1.Text)
End Sub
Private Sub Text2_GotFocus()
If Len(Trim(Text6.Text)) > 0 Then
    Text2.BackColor = &H80FFFF
Else
    MsgBox "Debe ingrasar la Razón Social del proveedor", vbInformation, empresa
    Text6.SetFocus
End If
End Sub
Private Sub Text2_LostFocus()
Text2.BackColor = &HFFFFFF
Text2.Text = UCase(Text2.Text)
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
Text4.Text = LCase(Text4.Text)
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

Private Function limpiadatos()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo1.Text = ""
seleccion = 0
ssFrame1.Enabled = False
End Function
Private Function verificapro()
Dim Cn As New ADODB.Connection
Dim rspr As New ADODB.Recordset   ' Recordset de proveedores
Cn.ConnectionString = Cadena
Cn.Open

rspr.CursorType = adOpenKeyset
rspr.LockType = adLockOptimistic
rspr.ActiveConnection = Cn
rspr.Source = "Select * from proveedores where prv_des = " & "'" & vprv_des & "'"
rspr.Open

If Not rspr.EOF Then
    siexiste = 1
Else
    siexiste = 0
End If


End Function
