VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form Kardex 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9495
   ClientLeft      =   6765
   ClientTop       =   4080
   ClientWidth     =   8865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   540
      TabIndex        =   1
      Top             =   1575
      Width           =   1995
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   540
      TabIndex        =   0
      Top             =   945
      Width           =   4335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3600
      Top             =   6705
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
      RecordSource    =   "Select * from kardex order by kar_fec"
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
   Begin Threed.SSCommand SSCommand2 
      Height          =   375
      Left            =   6705
      TabIndex        =   2
      Top             =   2700
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
      Caption         =   "Ver Kardex"
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   420
      Left            =   6840
      TabIndex        =   3
      Top             =   8955
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
   Begin Threed.SSCommand SSCommand5 
      Height          =   420
      Left            =   5400
      TabIndex        =   4
      Top             =   8955
      Visible         =   0   'False
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid2 
      Bindings        =   "Kardex.frx":0000
      Height          =   5460
      Left            =   495
      TabIndex        =   5
      Top             =   3240
      Width           =   7770
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
      SelectTypeRow   =   0
      ForeColorEven   =   0
      BackColorEven   =   14737632
      BackColorOdd    =   14737632
      RowHeight       =   423
      Columns.Count   =   5
      Columns(0).Width=   2540
      Columns(0).Caption=   "Fecha"
      Columns(0).Name =   "kar_fec"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   1
      Columns(0).DataField=   "kar_fec"
      Columns(0).DataType=   7
      Columns(0).NumberFormat=   "dd-mm-yyyy"
      Columns(0).FieldLen=   256
      Columns(1).Width=   2778
      Columns(1).Caption=   "Ingreso"
      Columns(1).Name =   "Kar_canI"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Kar_canI"
      Columns(1).DataType=   5
      Columns(1).NumberFormat=   "####"
      Columns(1).FieldLen=   256
      Columns(2).Width=   2937
      Columns(2).Caption=   "Egreso"
      Columns(2).Name =   "Kar_canE"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   1
      Columns(2).DataField=   "Kar_canE"
      Columns(2).DataType=   5
      Columns(2).NumberFormat=   "####"
      Columns(2).FieldLen=   256
      Columns(3).Width=   3228
      Columns(3).Caption=   "Saldo"
      Columns(3).Name =   "Kar_sal"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   1
      Columns(3).DataField=   "Kar_sal"
      Columns(3).DataType=   5
      Columns(3).NumberFormat=   "####"
      Columns(3).FieldLen=   256
      Columns(4).Width=   1191
      Columns(4).Caption=   "Tipo"
      Columns(4).Name =   "kar_est"
      Columns(4).CaptionAlignment=   0
      Columns(4).DataField=   "kar_tip"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      _ExtentX        =   13705
      _ExtentY        =   9631
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
      Left            =   5265
      Top             =   855
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
      RecordSource    =   "Select * from tie ORDER  by ProCod"
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
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   540
      TabIndex        =   18
      Top             =   2205
      Width           =   2625
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marca"
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
      TabIndex        =   17
      Top             =   1980
      Width           =   540
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Talla Lit."
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
      TabIndex        =   16
      Top             =   1980
      Width           =   765
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   4365
      TabIndex        =   15
      Top             =   2205
      Width           =   825
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Talla Num."
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
      Left            =   3330
      TabIndex        =   14
      Top             =   1980
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   3285
      TabIndex        =   13
      Top             =   2205
      Width           =   960
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Almacén"
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
      Left            =   585
      TabIndex        =   12
      Top             =   720
      Width           =   735
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2700
      TabIndex        =   11
      Top             =   1350
      Width           =   780
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   5355
      TabIndex        =   10
      Top             =   2205
      Width           =   2715
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
      Left            =   5400
      TabIndex        =   9
      Top             =   1980
      Width           =   450
   End
   Begin VB.Label Label2 
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
      Left            =   540
      TabIndex        =   8
      Top             =   1350
      Width           =   600
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   2655
      TabIndex        =   7
      Top             =   1575
      Width           =   5640
   End
   Begin VB.Image Image1 
      Height          =   60
      Left            =   2115
      Picture         =   "Kardex.frx":0015
      Stretch         =   -1  'True
      Top             =   540
      Width           =   6540
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "KARDEX INDIVIDUAL"
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
      Left            =   5895
      TabIndex        =   6
      Top             =   180
      Width           =   2805
   End
End
Attribute VB_Name = "Kardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_LostFocus()
If Len(Trim(Text1.Text)) > 0 Then
    Dim Cn As New ADODB.Connection
    Dim rspr As New ADODB.Recordset
    Cn.ConnectionString = Cadena
    Cn.Open
    
    vpro_cod = Text1.Text
    rspr.CursorType = adOpenKeyset
    rspr.LockType = adLockOptimistic
    rspr.ActiveConnection = Cn
    If Combo2.Text = "ALMACEN GENERAL" Then
        rspr.Source = "Select * from vproducto WHERE Pro_Cod = " & "'" & vpro_cod & "'"
    ElseIf Combo2.Text = "TIENDA" Then
            rspr.Source = "Select * from vproductotie WHERE Pro_Cod = " & "'" & vpro_cod & "'"
    End If
    rspr.Open

    If Not rspr.EOF Then
        Label8.Caption = rspr!Gru_des & " " & rspr!Pro_des
        Label11.Caption = rspr!Pro_Tip
        Label1.Caption = rspr!ProTNu
        Label5.Caption = rspr!ProTLi
        Label13.Caption = rspr!Mar_Des
        Text1.BackColor = &HFFFFFF
        SSCommand2_Click
    Else
        MsgBox "Producto inexistente", vbInformation, empresa
    End If
End If
End Sub
Private Sub Combo2_GotFocus()
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
    Combo2.Clear
    Do While Not rsal.EOF
       Combo2.AddItem rsal!almdes
       rsal.MoveNext
    Loop
End If
Combo2.ListIndex = 0
Cn.Close
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
Private Sub Form_Load()
KeyPreview = True
Dim Cn As New ADODB.Connection
Cn.ConnectionString = Cadena
Cn.Open
borrak = "DELETE FROM kardex"
Cn.Execute borrak
Adodc1.Refresh
Cn.Close
End Sub
Private Sub SSCommand2_Click()
On Error GoTo errork

If Len(Trim(Text1.Text)) > 0 Then
    Dim Cn As New ADODB.Connection
    Dim rska As New ADODB.Recordset   ' Recordset de kardex
    Cn.ConnectionString = Cadena
    Cn.Open
    
    'Carga Inv. Inicial
    cinv = "INSERT INTO kardex(kar_fec, kar_canI, kar_tip) Select Ini_fec,Ini_can, 'I' FROM vinvini WHERE Pro_cod = " & "'" & vpro_cod & "'"
    Cn.Execute cinv
    
    'Carga Ingresos
    cing = "INSERT INTO kardex(kar_fec, kar_canI, kar_tip) Select cingr_fec,ding_can, 'C' FROM vcompra WHERE Pro_cod = " & "'" & vpro_cod & "' AND cing_obs = 'C'"
    Cn.Execute cing
    
    cing = "INSERT INTO kardex(kar_fec, kar_canI, kar_tip) Select cingr_fec,ding_can, 'D' FROM vcompra WHERE Pro_cod = " & "'" & vpro_cod & "' AND cing_obs = 'D'"
    Cn.Execute cing
    
    'Carga Salidas
    csal = "INSERT INTO kardex(kar_fec, kar_canE, kar_tip, kar_est) Select cegr_fec,degr_can, 'V',degr_est FROM vsalida WHERE Pro_cod = " & "'" & vpro_cod & "' AND degr_est = 'V'"
    Cn.Execute csal
    
    csal = "INSERT INTO kardex(kar_fec, kar_canE, kar_tip, kar_est) Select cegr_fec,degr_can, 'B',degr_est FROM vsalida WHERE Pro_cod = " & "'" & vpro_cod & "' AND degr_est = 'B'"
    Cn.Execute csal
    'Calculo
    rska.CursorType = adOpenKeyset
    rska.LockType = adLockOptimistic
    rska.ActiveConnection = Cn
    rska.Source = "Select * from kardex ORDER BY kar_fec"
    rska.Open
    vsaldo = 0
    If Not rska.EOF Then
        If rska!kar_tip = "I" Then
            rska!kar_sal = rska!kar_canI
            rska.Update
            vsaldo = rska!kar_canI
             rska.MoveNext
        End If
        
        Do While Not rska.EOF
            If rska!kar_tip = "C" Or rska!kar_tip = "D" Then
                vkar_id = rska!kar_id
                vkar_canI = rska!kar_canI
                vkar_sal = vsaldo + vkar_canI
                actc = "UPDATE kardex SET kar_sal = " & vkar_sal & " WHERE kar_id = " & vkar_id
                Cn.Execute actc
                'rska!kar_sal = vsaldo + rska!kar_canI
                'rska.Update
                vsaldo = vkar_sal
            End If
            If rska!kar_tip = "V" Or rska!kar_tip = "B" Then
                If rska!kar_est <> "C" Then
                    vkar_id = rska!kar_id
                    vkar_canE = rska!kar_canE
                    vkar_sal = vsaldo - vkar_canE
                    actc = "UPDATE kardex SET kar_sal = " & vkar_sal & " WHERE kar_id = " & vkar_id
                    Cn.Execute actc
                End If
                'rska!kar_sal = vsaldo + rska!kar_canI
                'rska.Update
                vsaldo = vkar_sal
            End If
            rska.MoveNext
        Loop
        Adodc1.Refresh
        
        SSCommand6.SetFocus
    End If
Else
    MsgBox "Debe seleccionar un producto de la lista", vbInformation, empresa
    SSOleDBtext1.SetFocus
End If

errork:
If Err.Number = 3021 Then
    MsgBox "Producto sin movimiento", vbInformation, empresa
End If
End Sub
Private Sub SSCommand6_Click()
Menup.Enabled = True
Unload Kardex
Set Kardex = Nothing
End Sub
Private Sub SSOleDBtext1_GotFocus()
Dim Cn As New ADODB.Connection
Cn.ConnectionString = Cadena
Cn.Open
borrak = "DELETE FROM kardex"
Cn.Execute borrak
kauno = "ALTER TABLE kardex AUTO_INCREMENT = " & 1
Cn.Execute kauno
Label1.Caption = ""
Adodc1.Refresh
Cn.Close
End Sub
Private Sub Text1_GotFocus()
If Len(Trim(Combo2.Text)) > 0 Then
    valm = Combo2.Text
    Text1.BackColor = &H80FFFF
    Dim Cn As New ADODB.Connection
    Dim rsal As New ADODB.Recordset   ' Recordset de almacenes
    Dim rsco As New ADODB.Recordset   ' Recordset de codigos
    Cn.ConnectionString = Cadena
    Cn.Open
    
    rsal.CursorType = adOpenKeyset
    rsal.LockType = adLockOptimistic
    rsal.ActiveConnection = Cn
    rsal.Source = "Select * from almacen WHERE AlmDes = " & "'" & valm & "'"
    rsal.Open
    
    If Not rsal.EOF Then
        valmsig = rsal!almsig
        valmid = rsal!almid
    Else
        MsgBox "Debe seleccionar un almacén de la lista", vbInformation, empresa
        Combo2.SetFocus
    End If
    
    ''Carga codigos del almacen seleccionado
    
    'Adodc4.RecordSource = "SELECT procod FROM " & valmsig & " ORDER BY procod"
    'Adodc4.Refresh
    
    'If Not Adodc4.Recordset.EOF Then
    '    Text1.Clear
    '    Do While Not Adodc4.Recordset.EOF
    '        Text1.AddItem Adodc4.Recordset.Fields("ProCod")
    '        Adodc4.Recordset.MoveNext
    '    Loop
    'Else
    '    MsgBox "Almacén sin productos", vbInformation, empresa
    'End If
Else
    MsgBox "Debe seleccionar un almacén", vbInformation, empresa
    Combo2.SetFocus
End If
End Sub
