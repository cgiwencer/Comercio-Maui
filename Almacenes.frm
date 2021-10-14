VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Almacenes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6300
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   8415
      Top             =   2565
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
      Left            =   2295
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
      RecordSource    =   "Select * From almacen ORDER By almdes"
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
      Left            =   8055
      TabIndex        =   0
      Top             =   765
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
      Left            =   8055
      TabIndex        =   5
      Top             =   3465
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
      Left            =   8055
      TabIndex        =   6
      Top             =   1350
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
      Left            =   8055
      TabIndex        =   7
      Top             =   1890
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid1 
      Bindings        =   "Almacenes.frx":0000
      Height          =   3255
      Left            =   495
      TabIndex        =   8
      Top             =   675
      Width           =   7455
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
      Columns(0).Width=   4207
      Columns(0).Caption=   "Almacén"
      Columns(0).Name =   "AlmDes"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "AlmDes"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1773
      Columns(1).Caption=   "Sigla"
      Columns(1).Name =   "AlmSig"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "AlmSig"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   6138
      Columns(2).Caption=   "Dirección"
      Columns(2).Name =   "AlmDir"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "AlmDir"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   13150
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
      Height          =   2085
      Left            =   315
      TabIndex        =   9
      Top             =   4095
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3678
      _Version        =   196608
      BackStyle       =   1
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   4905
         MaxLength       =   6
         TabIndex        =   2
         Top             =   495
         Width           =   1050
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   135
         TabIndex        =   3
         Top             =   1440
         Width           =   4650
      End
      Begin VB.TextBox Text6 
         Height          =   330
         Left            =   135
         TabIndex        =   1
         Top             =   495
         Width           =   4650
      End
      Begin Threed.SSCommand SSCommand5 
         Height          =   420
         Left            =   6435
         TabIndex        =   10
         Top             =   1350
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
         Left            =   4950
         TabIndex        =   4
         Top             =   1350
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
         Left            =   6975
         TabIndex        =   11
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
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "(Max. 6 caracteres)"
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
         Left            =   4905
         TabIndex        =   16
         Top             =   855
         Width           =   2040
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Sigla"
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
         TabIndex        =   15
         Top             =   270
         Width           =   1410
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
         TabIndex        =   14
         Top             =   1215
         Width           =   2355
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Almacén"
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
   Begin VB.Image Image1 
      Height          =   60
      Left            =   3780
      Picture         =   "Almacenes.frx":0015
      Stretch         =   -1  'True
      Top             =   450
      Width           =   5640
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ALMACENES"
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
      Left            =   7830
      TabIndex        =   13
      Top             =   135
      Width           =   1590
   End
End
Attribute VB_Name = "Almacenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim valmid As Integer
Dim valmsig As String
Private Sub SSCommand5_Click()
limpiadatos
End Sub

Private Sub Text1_GotFocus()
If Len(Trim(Text2.Text)) > 0 Then
    Text1.BackColor = &H80FFFF
Else
    MsgBox "Debe ingresar la sigla del almacén", vbInformation, empresa
    Text2.SetFocus
End If
End Sub
Private Sub Text1_LostFocus()
Text1.BackColor = &HFFFFFF
Text1.Text = UCase(Text1.Text)
End Sub

Private Sub Text2_GotFocus()
If Len(Trim(Text6.Text)) > 0 Then
    Text2.BackColor = &H80FFFF
Else
    MsgBox "Debe ingresar la denominación del almacen", vbInformation, empresa
    Text6.SetFocus
End If
End Sub
Private Sub Text2_LostFocus()
Text2.BackColor = &HFFFFFF
Text2.Text = UCase(Text2.Text)
End Sub
Private Sub Text6_GotFocus()
Text6.BackColor = &H80FFFF
End Sub
Private Sub Text6_LostFocus()
Text6.BackColor = &HFFFFFF
Text6.Text = UCase(Text6.Text)
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
Unload Almacenes
Set Almacenes = Nothing
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

Private Sub SSCommand6_Click()
If Len(Trim(Text6.Text)) > 0 Then
    Dim Cn As New ADODB.Connection
    Cn.ConnectionString = Cadena
    Cn.Open
    
    valmdes = Trim(Text6.Text)
    valmsig = Trim(Text2.Text)
    valmdir = Trim(Text1.Text)
    If SSCheck1.Value = -1 Then
        valmest = 1
    Else
        valmest = 0
    End If
    If modo = "N" Then
        verificagru
    Else
        siexiste = 0
    End If
    If siexiste = 0 Then
        If modo = "N" Then
            grabap = "INSERT INTO almacen SET almdes = " & "'" & valmdes & "',almsig = " & "'" & valmsig & "', almdir = " & "'" & valmdir & "', almest= " & valmest
        ElseIf modo = "M" Then
             grabap = "UPDATE almacen SET almdes = " & "'" & valmdes & "',almsig = " & "'" & valmsig & "', almdir = " & "'" & valmdir & "', almest= " & valmest & " WHERE almid = " & valmid
        End If
        Cn.Execute grabap
        Adodc1.Refresh
        
        'Crea tabla del almacen
        creatablaal
        
        limpiadatos
        Cn.Close
        MsgBox "Información registrada", vbInformation, empresa
    ElseIf siexiste = 1 Then
        MsgBox "Almacen existente. No puede grabar", vbInformation, empresa
        Text6.SetFocus
    End If
Else
    MsgBox "Debe ingresar nombre del almacén", vbInformation, empresa
    Text6.SetFocus
End If

End Sub

Private Sub SSOleDBGrid1_Click()
If Adodc1.Recordset.RecordCount > 0 Then
    seleccion = 1
    valmid = Adodc1.Recordset.Fields("almid")
    Text6.Text = Adodc1.Recordset.Fields("almdes")
    Text1.Text = Adodc1.Recordset.Fields("almdir") & ""
    Text2.Text = Adodc1.Recordset.Fields("almsig") & ""
    If Adodc1.Recordset.Fields("almest") = 1 Then
        SSCheck1.Value = -1
    Else
        SSCheck1.Value = 0
    End If
End If

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
Private Function limpiadatos()
Text6.Text = ""
Text1.Text = ""
Text2.Text = ""
SSCheck1.Value = 0
seleccion = 0
ssFrame1.Enabled = False
End Function


Private Function creatablaal()
On Error GoTo errort

Dim TaUsu As String
vusuariot = Trim(vusucod)
TaUsu = "CREATE TABLE " & valmsig & "(" _
        & "AlmId int(5) DEFAULT NULL, " _
        & "ProCod varchar(40) DEFAULT NULL, " _
        & "ProCan int(5) DEFAULT NULL, " _
        & "ProDes varchar(250) DEFAULT NULL, " _
        & "ProCol varchar(150) DEFAULT NULL, " _
        & "ProTNu int(3) DEFAULT NULL, " _
        & "ProTLi varchar(5) DEFAULT NULL, " _
        & "ProPVe double(8,2) DEFAULT NULL, " _
        & "ProPOf double(8,2) DEFAULT NULL, " _
        & "IniId int(5) DEFAULT NULL, " _
        & "UsuRes int(3) DEFAULT NULL)"
        
Dim Cn As New ADODB.Connection
Cn.ConnectionString = Cadena
Cn.Open
Cn.Execute TaUsu
Adodc2.RecordSource = "Select * from Tdettraspaso" & vusuariot
Adodc2.Refresh

errort:
If Err.Number = -2147217900 Then
    borrat = "Drop Table " & valmsig
    Cn.Execute borrat
    creatablaal
End If
End Function


