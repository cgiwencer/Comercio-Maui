VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form Devoluciones 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5550
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12405
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   12405
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame SSFrame1 
      Height          =   4515
      Left            =   270
      TabIndex        =   3
      Top             =   810
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   7964
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   2790
         TabIndex        =   0
         Top             =   1125
         Width           =   645
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   3780
         TabIndex        =   1
         Top             =   1125
         Width           =   1320
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   2880
         Top             =   3105
         Width           =   2040
         _ExtentX        =   3598
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
         RecordSource    =   "Select * from cambio"
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
         Bindings        =   "Devoluciones.frx":0000
         Height          =   2085
         Left            =   90
         TabIndex        =   4
         Top             =   1755
         Width           =   8565
         _Version        =   196616
         ForeColorEven   =   0
         BackColorEven   =   14737632
         BackColorOdd    =   14737632
         RowHeight       =   423
         Columns.Count   =   5
         Columns(0).Width=   2381
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "pro_cod"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "pro_cod"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   7170
         Columns(1).Caption=   "Producto"
         Columns(1).Name =   "pro_des"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "pro_des"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1005
         Columns(2).Caption=   "Cant."
         Columns(2).Name =   "Pro_can"
         Columns(2).Alignment=   1
         Columns(2).CaptionAlignment=   1
         Columns(2).DataField=   "Pro_can"
         Columns(2).DataType=   3
         Columns(2).FieldLen=   256
         Columns(3).Width=   1799
         Columns(3).Caption=   "Pr.Unit."
         Columns(3).Name =   "pro_pru"
         Columns(3).Alignment=   1
         Columns(3).CaptionAlignment=   1
         Columns(3).DataField=   "pro_pru"
         Columns(3).DataType=   5
         Columns(3).NumberFormat=   "####.#0"
         Columns(3).FieldLen=   256
         Columns(4).Width=   1693
         Columns(4).Caption=   "Total"
         Columns(4).Name =   "pro:_tot"
         Columns(4).Alignment=   1
         Columns(4).CaptionAlignment=   1
         Columns(4).DataField=   "pro_tot"
         Columns(4).DataType=   5
         Columns(4).NumberFormat=   "####.#0"
         Columns(4).FieldLen=   256
         _ExtentX        =   15108
         _ExtentY        =   3678
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
         Left            =   8640
         Top             =   2250
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
      Begin Threed.SSCommand SSCommand2 
         Height          =   375
         Left            =   5895
         TabIndex        =   2
         Top             =   1125
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
         Caption         =   "Registrar"
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   600
         Left            =   8820
         TabIndex        =   8
         Top             =   3195
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
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
         Caption         =   "&Grabar Devolución"
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand8 
         Height          =   600
         Left            =   8955
         TabIndex        =   10
         Top             =   900
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         _Version        =   196608
         CaptionStyle    =   1
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
         Caption         =   "&Ventas Anteriores"
         BevelWidth      =   1
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCombo2 
         Bindings        =   "Devoluciones.frx":0015
         Height          =   330
         Left            =   90
         TabIndex        =   12
         Top             =   405
         Width           =   2130
         DataFieldList   =   "pro_cod"
         _Version        =   196616
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorEven   =   0
         BackColorEven   =   12632256
         BackColorOdd    =   12632256
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   4048
         Columns(0).Caption=   "Codigo"
         Columns(0).Name =   "pro_cod"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "pro_cod"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1217
         Columns(1).Caption=   "Exist."
         Columns(1).Name =   "pro_exi"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "pro_exi"
         Columns(1).DataType=   5
         Columns(1).FieldLen=   256
         _ExtentX        =   3757
         _ExtentY        =   582
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "pro_cod"
      End
      Begin VB.Label Label7 
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
         Left            =   4770
         TabIndex        =   22
         Top             =   180
         Width           =   780
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2295
         TabIndex        =   21
         Top             =   405
         Width           =   2805
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   8190
         TabIndex        =   20
         Top             =   405
         Width           =   2040
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
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   135
         TabIndex        =   19
         Top             =   180
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
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   8235
         TabIndex        =   18
         Top             =   180
         Width           =   450
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   90
         TabIndex        =   17
         Top             =   1125
         Width           =   1050
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T. Literal"
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
         Left            =   135
         TabIndex        =   16
         Top             =   900
         Width           =   780
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T. Numeral"
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
         Left            =   1485
         TabIndex        =   15
         Top             =   900
         Width           =   945
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1440
         TabIndex        =   14
         Top             =   1125
         Width           =   1050
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   5175
         TabIndex        =   13
         Top             =   405
         Width           =   2940
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7290
         TabIndex        =   11
         Top             =   3915
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
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
         Left            =   2790
         TabIndex        =   7
         Top             =   900
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pr.Unitario"
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
         Left            =   3780
         TabIndex        =   6
         Top             =   900
         Width           =   915
      End
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   420
      Left            =   10935
      TabIndex        =   5
      Top             =   4635
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
      BackStyle       =   0  'Transparent
      Caption         =   "DEVOLUCIONES"
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
      Left            =   10215
      TabIndex        =   9
      Top             =   180
      Width           =   2085
   End
   Begin VB.Image Image1 
      Height          =   60
      Left            =   7785
      Picture         =   "Devoluciones.frx":002A
      Stretch         =   -1  'True
      Top             =   495
      Width           =   4515
   End
End
Attribute VB_Name = "Devoluciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vpro_id As Integer
Private Sub Form_Load()
KeyPreview = True
Dim Cn As New ADODB.Connection
Cn.ConnectionString = Cadena
Cn.Open

borrac = "DELETE FROM cambio"
Cn.Execute borrac
Adodc1.Refresh
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

Private Sub SSCommand2_Click()
Dim Cn As New ADODB.Connection
Dim rsdt As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open

vpro_cod = SSOleDBCombo2.Columns(0).Value
vpro_can = Val(Text1.Text)
vpro_des = Label5.Caption & " " & Label25.Caption
vpro_pru = Val(Text4.Text)
vpro_tot = vpro_can * vpro_pru

'Verifica si el item ya existe
rsdt.CursorType = adOpenKeyset
rsdt.LockType = adLockOptimistic
rsdt.ActiveConnection = Cn
rsdt.Source = "Select * from cambio Where Pro_cod = " & "'" & vpro_cod & "'"
rsdt.Open

If rsdt.EOF Then
    grabac = "INSERT INTO cambio SET pro_cod = " & "'" & vpro_cod & "', pro_can = " & vpro_can & ", pro_des = " & "'" & vpro_des & _
    "', pro_pru = " & vpro_pru & ", pro_tot = " & vpro_tot
    Cn.Execute grabac
    
    Adodc1.Refresh
    Cn.Close
    Label4.Caption = Format(Val(Label4.Caption) + vpro_tot, "####.#0")
    limpiadatos
     SSOleDBCombo2.SetFocus
Else
    MsgBox "Item ya registrado", vbInformation, empresa
    limpiadatos
    SSOleDBCombo1.SetFocus
End If
End Sub

Private Sub SSCommand3_Click()
porcambio
End Sub

Private Sub SSCommand6_Click()
Menup.Enabled = True
Unload Devoluciones
Set Devoluciones = Nothing
End Sub

Private Sub SSCommand8_Click()
origen = "devoluciones"
Devoluciones.Enabled = False
Load EgresosAnt
EgresosAnt.Show
End Sub

Private Sub Text1_GotFocus()
If Len(Trim(SSOleDBCombo2.Text)) > 0 Then
    vcod_pro = SSOleDBCombo2.Columns(0).Value
        
    Dim Cn As New ADODB.Connection
    Dim rsit As New ADODB.Recordset   ' Recordset de item
    Cn.ConnectionString = Cadena
    Cn.Open
        
    rsit.CursorType = adOpenKeyset
    rsit.LockType = adLockOptimistic
    rsit.ActiveConnection = Cn
    rsit.Source = "Select * from vproducto Where Pro_Cod = " & "'" & vcod_pro & "'"
    rsit.Open
    
    If Not rsit.EOF Then
        Label5.Caption = rsit!Gru_des
        Label25.Caption = rsit!Pro_des
        Label8.Caption = rsit!Pro_tip
        Label13.Caption = rsit!ProTLi & ""
        Label24.Caption = rsit!ProTNu & ""
        Text4.Text = Format(rsit!pro_pve, "####.#0")
        vcosto = rsit!pro_ppp & ""
        vexi = rsit!pro_exi
    Else
        MsgBox "Item inexistente", vbInformation, empresa
        SSOleDBCombo1.SetFocus
    End If
       
    Text1.BackColor = &H9FBDEA
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
Else
    MsgBox "Debe seleccionar el producto", vbInformation, empresa
    SSOleDBCombo2.SetFocus
End If
End Sub
Private Sub Text1_LostFocus()
Text1.BackColor = &HFFFFFF
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 46 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub Text4_GotFocus()
If Len(Trim(Text1.Text)) > 0 Then
    Text4.BackColor = &H9FBDEA
    Text4.SelStart = 0
    Text4.SelLength = Len(Text4.Text)
Else
    MsgBox "Debe ingresar la cantidad", vbInformation, empresa
    Text1.SetFocus
End If
End Sub
Private Sub Text4_LostFocus()
Text4.BackColor = &HFFFFFF
Text4.Text = Format(Text4.Text, "####.#0")
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 46 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub

Private Function porcambio()
Dim Cn As New ADODB.Connection
Dim rsca As New ADODB.Recordset   ' Recordset de cambio
Dim rsci As New ADODB.Recordset   ' Recordset de cabecera de ingreso
Dim rsdt As New ADODB.Recordset   ' Recordset de detalle temporal
Cn.ConnectionString = Cadena
Cn.Open

rsca.CursorType = adOpenKeyset
rsca.LockType = adLockOptimistic
rsca.ActiveConnection = Cn
rsca.Source = "Select * from cambio"
rsca.Open

'vdegr_id = Val(Text3.Text)
vpro_cod = rsca!pro_cod
vpro_can = rsca!pro_can
vpro_des = rsca!Pro_des
vini_fec = Format(Date, "yyyy-mm-dd")

'Graba en cabecera de ingreso
grabacii = "INSERT INTO cabingreso SET cingr_fec = " & "'" & vini_fec & "', usu_id = " & vusucod & ", cing_obs = 'D'"
Cn.Execute grabacii

rsci.CursorType = adOpenKeyset
rsci.LockType = adLockOptimistic
rsci.ActiveConnection = Cn
rsci.Source = "Select * from cabingreso"
rsci.Open

If Not rsci.EOF Then
    rsci.MoveLast
    vini_id = rsci!cingr_id
Else
    vini_id = 1
End If

'Graba en detalle de ingreso
grabad = "INSERT INTO detingreso SET cing_id = " & vini_id & ", pro_id = " & vpro_id & ", pro_cod = " & "'" & vpro_cod & _
"', pro_des = " & "'" & vpro_des & "', ding_can = " & vpro_can
Cn.Execute grabad

'Actualiza inventario
rsca.MoveFirst
If Not rsca.EOF Then
    Do While Not rsca.EOF
        actpro = "UPDATE productos SET pro_exi = pro_exi + " & vpro_can & " WHERE pro_cod = " & "'" & vpro_cod & "'"
        Cn.Execute actpro
        rsca.MoveNext
    Loop
Else
End If

'Actualiza arqueo
vini_fec = Format(Date, "yyyy-mm-dd")
vpag = Val(Label4.Caption)

grabac = "INSERT INTO pagoventa SET pag_fec = " & "'" & vini_fec & "', pag_dev = " & vpag & ", usu_id = " & vusucod
Cn.Execute grabac
    
    
MsgBox "Proceso realizado", vbInformation, empresa
Cn.Close
SSCommand6_Click
End Function
Private Function limpiadatos()
SSOleDBCombo2.Text = ""
Text1.Text = ""
Text4.Text = ""
Label5.Caption = ""
Label8.Caption = ""
Label13.Caption = ""
Label24.Caption = ""
Label25.Caption = ""
End Function

