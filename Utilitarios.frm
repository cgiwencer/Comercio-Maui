VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form Utilitarios 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4620
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame SSFrame2 
      Height          =   1740
      Left            =   4785
      TabIndex        =   0
      Top             =   930
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   3069
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin Threed.SSCommand SSCommand1 
         Height          =   555
         Left            =   135
         TabIndex        =   5
         Top             =   135
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   979
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
         Caption         =   "&Crear copia de seguridad"
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   690
         Left            =   135
         TabIndex        =   6
         Top             =   900
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   1217
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
         Caption         =   "&Restaurar copia de seguridad"
         BevelWidth      =   1
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   990
      Top             =   4230
      Visible         =   0   'False
      Width           =   1950
      _ExtentX        =   3440
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
      RecordSource    =   "Select * from backup"
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid1 
      Bindings        =   "Utilitarios.frx":0000
      Height          =   3435
      Left            =   195
      TabIndex        =   1
      Top             =   780
      Width           =   3915
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
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      ForeColorEven   =   0
      BackColorEven   =   12632256
      BackColorOdd    =   12632256
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   1244
      Columns(0).Caption=   "Numero"
      Columns(0).Name =   "Numero"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   1
      Columns(0).DataField=   "Numero"
      Columns(0).DataType=   3
      Columns(0).FieldLen=   256
      Columns(1).Width=   1799
      Columns(1).Caption=   "Fecha"
      Columns(1).Name =   "Fecha"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Fecha"
      Columns(1).DataType=   7
      Columns(1).NumberFormat=   "dd/mm/yyyy"
      Columns(1).FieldLen=   256
      Columns(2).Width=   2805
      Columns(2).Caption=   "Nombre  del backup"
      Columns(2).Name =   "Nombre"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Nombre"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   6906
      _ExtentY        =   6059
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
   Begin Threed.SSCommand Boton18 
      Height          =   420
      Left            =   4905
      TabIndex        =   4
      Top             =   2880
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "COPIA DE SEGURIDAD"
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
      Left            =   4410
      TabIndex        =   3
      Top             =   90
      Width           =   3120
   End
   Begin VB.Image Image1 
      Height          =   60
      Left            =   225
      Picture         =   "Utilitarios.frx":0015
      Stretch         =   -1  'True
      Top             =   405
      Width           =   7080
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Procesando....por favor espere"
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
      Height          =   675
      Left            =   4170
      TabIndex        =   2
      Top             =   3780
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "Utilitarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmd As String
Private Declare Function OpenProcess Lib "KERNEL32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "KERNEL32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long
     
Private Const SYNCHRONIZE       As Long = &H100000
Private Const INFINITE          As Long = &HFFFF
     
Private Sub execCommand(ByVal cmd As String)
Dim result  As Long
Dim lPid    As Long
Dim lHnd    As Long
Dim lRet    As Long

cmd = "cmd /c " & cmd
result = Shell(cmd, vbHide)

lPid = result
If lPid <> 0 Then
    lHnd = OpenProcess(SYNCHRONIZE, 0, lPid)
    If lHnd <> 0 Then
        lRet = WaitForSingleObject(lHnd, INFINITE)
        CloseHandle (lHnd)
    End If
End If
End Sub

Private Sub Boton18_Click()
Menup.Enabled = True
Unload Utilitarios
Set Utilitarios = Nothing
End Sub

Private Sub SSCommand1_Click()
''MsgBox "Si el sistema funciona en red, las otras terminales debes cerrar el sistema", vbInformation, empresa
Screen.MousePointer = vbHourglass
Label2.Visible = True
DoEvents
     
Fechabk = Format(Date, "ddmmyyyy")
Fechagrid = Format(Date, "dd/mm/yyyy")
nombrebk = Fechabk & ".sql"
     

Set comando1 = CreateObject("WSCript.shell")
comando1.run "cmd /K C: & CD MYSQL\BIN & mysqldump -u root -p ferreteria > C:/" & nombrebk & ".sql"
Set comando1 = Nothing


'cmd = Chr(34) & App.Path & "\mysqldump" & Chr(34) & " -uroot -pcagisa Ferreteria > " & App.Path & "\BKP\" & nombrebk
'Call execCommand(cmd)
  
Screen.MousePointer = vbDefault

Dim Cn As New ADODB.Connection
Dim rsbk As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open

rsbk.CursorType = adOpenKeyset
rsbk.LockType = adLockOptimistic
rsbk.ActiveConnection = Cn
rsbk.Source = "SELECT * FROM backup"
rsbk.Open

If Not rsbk.EOF Then
    rsbk.MoveLast
    vnumero = rsbk!numero + 1
Else
    vnumero = 1
End If

graba = "INSERT INTO backup Set numero = " & vnumero & ", Fecha = " & "'" & Fechagrid & "', Nombre = " & "'" & nombrebk & "'"
Cn.Execute graba

Adodc1.Refresh

MsgBox "Backup realizado con éxito", vbInformation, empresa
Cn.Close
Label2.Visible = False
seleccion = 0
End Sub

Private Sub SSCommand2_Click()
If seleccion = 1 Then
    MsgBox "Si el sistema funciona en red, las otras terminales debes cerrar el sistema", vbInformation, empresa
    vbck = SSOleDBGrid1.Columns(2).Value
    Screen.MousePointer = vbHourglass
    Label2.Visible = True
    DoEvents
    
    cmd = Chr(34) & App.Path & "\mysqldump" & Chr(34) & " -uroot -pcagisa --routines --comments Madic > " & App.Path & "\Backups\" & vbck
    Call execCommand(cmd)
    
    Screen.MousePointer = vbDefault
    MsgBox "Restauración realizada con éxito", vbInformation, empresa
    Label2.Visible = False
    seleccion = 0
Else
    MsgBox "Debe seleccionar un backup de la lista", vbInformation, empresa
End If
End Sub

Private Sub SSOleDBGrid1_Click()
seleccion = 1
End Sub

