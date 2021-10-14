VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Bajas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7440
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   13065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   13065
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2205
      Top             =   2925
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
      RecordSource    =   "Select * From vprodbaj WHERE baj_est = 1 ORDER By baj_fec"
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   12015
      Top             =   2655
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
   Begin Threed.SSCommand SSCommand1 
      Height          =   420
      Left            =   11340
      TabIndex        =   0
      Top             =   1080
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
      Left            =   11340
      TabIndex        =   6
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
      Caption         =   "&Salir"
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand3 
      Height          =   420
      Left            =   11340
      TabIndex        =   7
      Top             =   1665
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
      Caption         =   "&Eliminar"
      BevelWidth      =   1
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3165
      Left            =   135
      TabIndex        =   9
      Top             =   4005
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5583
      _Version        =   196608
      BackStyle       =   1
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.TextBox Text2 
         Height          =   1140
         Left            =   2970
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1305
         Width           =   5055
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   225
         TabIndex        =   2
         Top             =   1305
         Width           =   645
      End
      Begin Threed.SSCommand SSCommand6 
         Height          =   420
         Left            =   6615
         TabIndex        =   5
         Top             =   2610
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
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   375
         Left            =   8280
         Top             =   3960
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
         Connect         =   "DSN=ferreteria"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "ferreteria"
         OtherAttributes =   ""
         UserName        =   "root"
         Password        =   "cagisa"
         RecordSource    =   "Select * from productos order by Pro_Des"
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCombo1 
         Bindings        =   "Bajas.frx":0000
         Height          =   330
         Left            =   225
         TabIndex        =   1
         Top             =   630
         Width           =   7890
         DataFieldList   =   "pro_Des"
         _Version        =   196616
         ForeColorEven   =   0
         BackColorEven   =   12632256
         BackColorOdd    =   12632256
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   8361
         Columns(0).Caption=   "Producto"
         Columns(0).Name =   "pro_Des"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "pro_Des"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Codigo"
         Columns(1).Name =   "pro_cod"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "pro_cod"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1640
         Columns(2).Caption=   "Exist."
         Columns(2).Name =   "pro_exi"
         Columns(2).Alignment=   1
         Columns(2).CaptionAlignment=   1
         Columns(2).DataField=   "pro_exi"
         Columns(2).DataType=   5
         Columns(2).FieldLen=   256
         _ExtentX        =   13917
         _ExtentY        =   582
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "pro_Des"
      End
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo1 
         Height          =   375
         Left            =   1035
         TabIndex        =   3
         Top             =   1305
         Width           =   1680
         _Version        =   65537
         _ExtentX        =   2963
         _ExtentY        =   661
         _StockProps     =   93
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo de la baja"
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
         Left            =   3060
         TabIndex        =   14
         Top             =   1080
         Width           =   1485
      End
      Begin VB.Label Label9 
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
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   1035
         TabIndex        =   13
         Top             =   1080
         Width           =   540
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
         Left            =   225
         TabIndex        =   12
         Top             =   1080
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
         Left            =   270
         TabIndex        =   11
         Top             =   360
         Width           =   780
      End
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid1 
      Bindings        =   "Bajas.frx":0015
      Height          =   3255
      Left            =   135
      TabIndex        =   8
      Top             =   630
      Width           =   11085
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
      Columns.Count   =   5
      Columns(0).Width=   2117
      Columns(0).Caption=   "Fecha"
      Columns(0).Name =   "Baj_fec"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   1
      Columns(0).DataField=   "Baj_fec"
      Columns(0).DataType=   7
      Columns(0).NumberFormat=   "dd-mm-yyyy"
      Columns(0).FieldLen=   256
      Columns(1).Width=   2328
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Pro_cod"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Pro_cod"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   6324
      Columns(2).Caption=   "Descripción"
      Columns(2).Name =   "pro_Des"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "pro_Des"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1270
      Columns(3).Caption=   "Cant."
      Columns(3).Name =   "Baj_can"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   1
      Columns(3).DataField=   "Baj_can"
      Columns(3).DataType=   3
      Columns(3).FieldLen=   256
      Columns(4).Width=   6456
      Columns(4).Caption=   "Motivo"
      Columns(4).Name =   "Baj_mot"
      Columns(4).CaptionAlignment=   0
      Columns(4).DataField=   "Baj_mot"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      _ExtentX        =   19553
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
   Begin VB.Image Image1 
      Height          =   60
      Left            =   6345
      Picture         =   "Bajas.frx":002A
      Stretch         =   -1  'True
      Top             =   360
      Width           =   6540
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "BAJA DE PRODUCTO"
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
      Left            =   10125
      TabIndex        =   10
      Top             =   45
      Width           =   2715
   End
End
Attribute VB_Name = "Bajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vpro_id As Integer
Dim vbaj_can As Integer
Dim vbaj_id As Integer

Private Sub SSCommand1_Click()
ssFrame1.Enabled = True
modo = "N"
SSOleDBCombo1.SetFocus
End Sub

Private Sub SSCommand2_Click()
Menup.Enabled = True
Unload Bajas
Set Bajas = Nothing
End Sub

Private Sub SSCommand3_Click()
If seleccion = 1 Then
    Dim Cn As New ADODB.Connection
    Dim rsce As New ADODB.Recordset
    Cn.ConnectionString = Cadena
    Cn.Open
    If MsgBox("Desea eliminar el registro?", vbYesNo, empresa) = vbYes Then
        'Cambia el estado en la tabla bajas
        borrab = "UPDATE bajas SET baj_est = " & 0 & " WHERE baj_id = " & vbaj_id
        Cn.Execute borrab
        'Borra registro de cabecera egresos
        rsce.CursorType = adOpenKeyset
        rsce.LockType = adLockOptimistic
        rsce.ActiveConnection = Cn
        rsce.Source = "Select * from bajas Where baj_id = " & vbaj_id
        rsce.Open
        
        If Not rsce.EOF Then
            vcegr_id = rsce!cegr_id
            borrac = "DELETE FROM cabegreso Where cegr_id=" & vcegr_id
            Cn.Execute borrac
            
            borrad = "DELETE FROM detegreso Where cegr_id=" & vcegr_id
            Cn.Execute borrad
        End If
        
        'Actualiza Producto
        actpro = "UPDATE productos SET pro_exi = pro_exi + " & vbaj_can & " WHERE pro_Cod = " & "'" & vpro_cod & "'"
        Cn.Execute actpro
        
        Adodc1.Refresh
    End If
Else
    MsgBox "Debe seleccionar un registro de la lista", vbInformation, empresa
End If
End Sub

Private Sub SSCommand6_Click()
Dim Cn As New ADODB.Connection
Dim rsce As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open

vpro_des = SSOleDBCombo1.Text
vbaj_fec = Format(SSDateCombo1, "yyyy-mm-dd")
vbaj_can = Val(Text1.Text)
vbaj_mot = Text2.Text
If modo = "N" Then
    'Actualiza tabla Productos
    actprod = "UPDATE productos SET pro_exi = pro_exi - " & vbaj_can & " WHERE pro_id = " & vpro_id
    Cn.Execute actprod
    
    'Graba cabecera de egreso
    grabacii = "INSERT INTO cabegreso SET cegr_fec = " & "'" & vbaj_fec & "', usu_id = " & vusucod
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

    'Graba detalle de egreso
    grabadet = "INSERT INTO detegreso SET cegr_id = " & vini_id & ", pro_id = " & vpro_id & ", pro_cod = " & "'" & vpro_cod & _
    "', pro_des = " & "'" & vpro_des & "', degr_can = " & vbaj_can & ", degr_pru = " & 0 & ", degr_prt = " & 0 & ",degr_est = 'B'"
    Cn.Execute grabadet
    
    'Graba en tabla Bajas
    grabap = "INSERT INTO bajas SET baj_fec = " & "'" & vbaj_fec & "', pro_cod= " & "'" & vpro_cod & "', baj_can = " & vbaj_can & _
    ", baj_mot = " & "'" & vbaj_mot & "', Usu_id = " & vusucod & ", Baj_est = " & 1 & ", cegr_id = " & vini_id
    Cn.Execute grabap

End If


Adodc1.Refresh
limpiadatos
Cn.Close
MsgBox "Información registrada", vbInformation, empresa

End Sub

Private Sub SSOleDBGrid1_Click()
seleccion = 1
vbaj_id = Adodc1.Recordset.Fields("baj_id")
vpro_cod = Adodc1.Recordset.Fields("Pro_cod")
vbaj_can = SSOleDBGrid1.Columns(3).Value
End Sub
Private Sub Text1_GotFocus()
If Len(Trim(SSOleDBCombo1.Text)) > 0 Then
    vpro_cod = SSOleDBCombo1.Columns(1).Value
        
    Dim Cn As New ADODB.Connection
    Dim rsit As New ADODB.Recordset   ' Recordset de item
    Cn.ConnectionString = Cadena
    Cn.Open
        
    rsit.CursorType = adOpenKeyset
    rsit.LockType = adLockOptimistic
    rsit.ActiveConnection = Cn
    rsit.Source = "Select * from productos Where Pro_Cod = " & "'" & vpro_cod & "'"
    rsit.Open
    
    If Not rsit.EOF Then
        vpro_id = rsit!pro_id
    Else
        MsgBox "Item inexistente", vbInformation, empresa
        SSOleDBCombo1.SetFocus
    End If
       
    Text1.BackColor = &H9FBDEA
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
Else
    MsgBox "Debe seleccionar el producto", vbInformation, empresa
    SSOleDBCombo1.SetFocus
End If
End Sub
Private Sub Text1_LostFocus()
Text1.BackColor = &HFFFFFF
Text1.Text = UCase(Text1.Text)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 46 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub Text2_GotFocus()
Text2.BackColor = &H80FFFF
End Sub
Private Sub Text2_LostFocus()
Text2.BackColor = &HFFFFFF
Text2.Text = UCase(Text2.Text)
End Sub

Private Function limpiadatos()
SSOleDBCombo1.Text = ""
Text1.Text = ""
Text2.Text = "'"
SSDateCombo1.Text = Date
ssFrame1.Enabled = False
End Function
