VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Cambiarclave 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4320
   ClientLeft      =   4530
   ClientTop       =   5145
   ClientWidth     =   5955
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtVNueva 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   315
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2340
      Width           =   2355
   End
   Begin VB.TextBox TxtNueva 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   315
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1755
      Width           =   2400
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2880
      Top             =   1800
      Visible         =   0   'False
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
      CommandType     =   2
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
      Connect         =   "DSN=ginecologia"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ginecologia"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "cagisa"
      RecordSource    =   "historico"
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
      Left            =   135
      TabIndex        =   8
      Top             =   3330
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
      Caption         =   "C&ambiar clave"
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   420
      Left            =   1755
      TabIndex        =   9
      Top             =   3330
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
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "CAMBIO DE CLAVE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3690
      TabIndex        =   7
      Top             =   225
      Width           =   2220
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   315
      TabIndex        =   6
      Top             =   1125
      Width           =   3750
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
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
      Left            =   315
      TabIndex        =   5
      Top             =   900
      Width           =   660
   End
   Begin VB.Image Image2 
      Height          =   780
      Left            =   0
      Picture         =   "Cambiarclave.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6045
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(10 caracteres max.)"
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
      Left            =   540
      TabIndex        =   4
      Top             =   2880
      Width           =   1755
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reescriba su clave"
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
      TabIndex        =   3
      Top             =   2160
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese su nueva clave"
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
      Left            =   315
      TabIndex        =   2
      Top             =   1530
      Width           =   2010
   End
End
Attribute VB_Name = "Cambiarclave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label6.Caption = Ingreso.Text1.Text
End Sub

Private Sub Image3_Click()

End Sub

Private Sub SSCommand1_Click()
On Error GoTo errornc

siexiste = 1
If Len(Trim(TxtNueva.Text)) = 0 Then
    If MsgBox("Su clave de acceso esta en blanco. Desea continuar..?", vbYesNo, "Inlasa") = vbNo Then
        siexiste = 0
        TxtNueva.SetFocus
    End If
End If
If siexiste = 1 Then
    If Trim(TxtNueva.Text) = Trim(TxtVNueva.Text) Then
        Dim Cn As New ADODB.Connection
        Dim rsu As New ADODB.Recordset
        Cn.ConnectionString = Cadena
        Cn.Open
        'Seleccion de tabla de Usuarios
        rsu.CursorType = adOpenKeyset
        rsu.LockType = adLockOptimistic
        rsu.ActiveConnection = Cn
        rsu.Source = "Select * from usuarios"
        rsu.Open
        
        If Not rsu.EOF Then
            rsu.MoveFirst
            Do While Not rsu.EOF
                If vusucod = rsu!Usu_Id Then
                    rsu!Usu_Cla = Trim(TxtNueva.Text)
                    'rsu!UsuLog = 0
                    'ghistorico
                    rsu.Update
                    MsgBox "El usuario " & vUsuario & " cambió la clave de acceso" & vbCrLf & _
                    "Por favor reingrese su clave", vbInformation
                    siexiste = 0
                    intento = 0
                    Exit Do
                End If
                rsu.MoveNext
            Loop
            SSCommand2_Click
            Ingreso.Text2.SetFocus
        Else
            MsgBox "Tabla de Usuarios vacia", vbInformation, "Consulte con Jefatura"
        End If
        SSCommand2_Click
    Else
        MsgBox "Las claves no son iguales", vbInformation, "Por favor revise"
        TxtNueva.SetFocus
    End If
End If

errornc:
If Err.Number = -2147467259 Then
    MsgBox "No se realizó ningún cambio", vbInformation, empresa
ElseIf Err.Number = -2147217887 Then
    MsgBox "Sólo 10 caracteres como máximo", vbInformation, empresa
End If
End Sub
Private Sub SSCommand2_Click()
Ingreso.Enabled = True
Unload Cambiarclave
Set Cambiarclave = Nothing
Ingreso.Text2.SetFocus
End Sub
Private Function ghistorico()
Adodc1.Refresh
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("HisFec") = Date
Adodc1.Recordset.Fields("HisDes") = "El usuario " & Label6.Caption & " cambio su clave "
Adodc1.Recordset.Fields("UsuCod") = 0
Adodc1.Recordset.Fields("HisHor") = Left(Time, 2) & ":" & Mid(Time, 4, 2)
Adodc1.Recordset.Update
End Function

