VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Proforma 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7545
   ClientLeft      =   6975
   ClientTop       =   4080
   ClientWidth     =   11100
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   11100
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   4455
      Top             =   7020
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
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   5010
      Left            =   225
      TabIndex        =   9
      Top             =   1665
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   8837
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   420
         Left            =   3195
         Top             =   3240
         Visible         =   0   'False
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   741
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
         RecordSource    =   "Select * from detproforma where cegr_id=0"
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
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   315
         TabIndex        =   3
         Top             =   1080
         Width           =   645
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1170
         TabIndex        =   4
         Top             =   1080
         Width           =   1320
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid2 
         Bindings        =   "Proforma.frx":0000
         Height          =   2265
         Left            =   180
         TabIndex        =   10
         Top             =   1980
         Width           =   9015
         _Version        =   196616
         ForeColorEven   =   0
         BackColorEven   =   14737632
         BackColorOdd    =   14737632
         RowHeight       =   423
         Columns.Count   =   5
         Columns(0).Width=   2884
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "pro_cod"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "pro_cod"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1588
         Columns(1).Caption=   "Cant."
         Columns(1).Name =   "degr_can"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "degr_can"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         Columns(2).Width=   6482
         Columns(2).Caption=   "Descripción"
         Columns(2).Name =   "pro_des"
         Columns(2).CaptionAlignment=   0
         Columns(2).DataField=   "pro_des"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   1746
         Columns(3).Caption=   "Pr. Unitario"
         Columns(3).Name =   "degr_pru"
         Columns(3).Alignment=   1
         Columns(3).CaptionAlignment=   1
         Columns(3).DataField=   "degr_pru"
         Columns(3).DataType=   5
         Columns(3).NumberFormat=   "####.#0"
         Columns(3).FieldLen=   256
         Columns(4).Width=   2117
         Columns(4).Caption=   "Total"
         Columns(4).Name =   "degr_prt"
         Columns(4).Alignment=   1
         Columns(4).CaptionAlignment=   1
         Columns(4).DataField=   "degr_prt"
         Columns(4).DataType=   5
         Columns(4).NumberFormat=   "####.#0"
         Columns(4).FieldLen=   256
         _ExtentX        =   15901
         _ExtentY        =   3995
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
         Left            =   8550
         Top             =   3780
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
      Begin Threed.SSCommand SSCommand2 
         Height          =   375
         Left            =   2745
         TabIndex        =   5
         Top             =   1080
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
         Left            =   9270
         TabIndex        =   11
         Top             =   2475
         Width           =   1140
         _ExtentX        =   2011
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
         Caption         =   "&Grabar Proforma"
         BevelWidth      =   1
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCombo1 
         Bindings        =   "Proforma.frx":0015
         Height          =   330
         Left            =   315
         TabIndex        =   2
         Top             =   450
         Width           =   8160
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
         _ExtentX        =   14393
         _ExtentY        =   582
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "pro_Des"
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
         Left            =   360
         TabIndex        =   19
         Top             =   180
         Width           =   780
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
         Left            =   360
         TabIndex        =   16
         Top             =   855
         Width           =   465
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   7965
         TabIndex        =   15
         Top             =   3465
         Width           =   75
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Precio de venta"
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
         Left            =   1170
         TabIndex        =   14
         Top             =   855
         Width           =   1365
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUB TOTAL"
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
         Left            =   6570
         TabIndex        =   13
         Top             =   4410
         Width           =   1050
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   7650
         TabIndex        =   12
         Top             =   4320
         Width           =   1275
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   825
      Left            =   225
      TabIndex        =   6
      Top             =   765
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   1455
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.TextBox Text6 
         Height          =   330
         Left            =   2025
         TabIndex        =   1
         Top             =   315
         Width           =   4605
      End
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo1 
         Height          =   375
         Left            =   225
         TabIndex        =   0
         Top             =   270
         Width           =   1680
         _Version        =   65537
         _ExtentX        =   2963
         _ExtentY        =   661
         _StockProps     =   93
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
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
         Left            =   2070
         TabIndex        =   8
         Top             =   45
         Width           =   600
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
         Left            =   225
         TabIndex        =   7
         Top             =   45
         Width           =   540
      End
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   420
      Left            =   9360
      TabIndex        =   17
      Top             =   6840
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
   Begin VB.Image Image1 
      Height          =   60
      Left            =   4320
      Picture         =   "Proforma.frx":002A
      Stretch         =   -1  'True
      Top             =   540
      Width           =   6540
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NUEVA PROFORMA"
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
      Left            =   8370
      TabIndex        =   18
      Top             =   180
      Width           =   2580
   End
End
Attribute VB_Name = "Proforma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
KeyPreview = True
creatablapro
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
If Len(Trim(Text4.Text)) > 0 Then
    'Text1.Text = ""
    'Text8.Text = ""
'    Label7.Caption = ""
    grabadetproftemp
Else
    MsgBox "Debe ingresar el precio de venta", vbInformation, empresa
    Text4.SetFocus
End If
End Sub

Private Sub SSCommand3_Click()
grabaproforma
limpiadatos
creatablapro

End Sub

Private Sub SSCommand6_Click()
Menup.Enabled = True
Unload Proforma
Set Proforma = Nothing
End Sub
Private Sub SSOleDBCombo1_GotFocus()
SSOleDBCombo1.DroppedDown = True
End Sub

Private Sub SSOleDBGrid2_DblClick()
Adodc1.Recordset.Delete
Adodc1.Refresh
End Sub
Private Sub Text1_GotFocus()
If Len(Trim(SSOleDBCombo1.Text)) > 0 Then
    vcod_pro = SSOleDBCombo1.Columns(1).Value
        
    Dim Cn As New ADODB.Connection
    Dim rsit As New ADODB.Recordset   ' Recordset de item
    Cn.ConnectionString = Cadena
    Cn.Open
        
    rsit.CursorType = adOpenKeyset
    rsit.LockType = adLockOptimistic
    rsit.ActiveConnection = Cn
    rsit.Source = "Select * from productos Where Pro_Cod = " & "'" & vcod_pro & "'"
    rsit.Open
    
    If Not rsit.EOF Then
        
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

Private Sub Text4_GotFocus()
If Len(Trim(Text1.Text)) > 0 Then
    Text4.BackColor = &H80FFFF
    Text4.SelStart = 0
    Text4.SelLength = Len(Text4.Text)
Else
    MsgBox "Debe ingresar la cantidad", vbCritical, empresa
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

Private Sub Text6_GotFocus()
Text6.BackColor = &H80FFFF
End Sub
Private Sub Text6_LostFocus()
Text6.BackColor = &HFFFFFF
Text6.Text = UCase(Text6.Text)
End Sub

Private Function creatablapro()
On Error GoTo errort

Dim TaUsu As String
vusuariot = Trim(vusucod)
TaUsu = "CREATE TABLE Tdetproforma" & vusuariot & "(" _
        & "cegr_Id int(5) DEFAULT NULL, " _
        & "degr_Id int(5) DEFAULT NULL, " _
        & "Pro_id int(5) DEFAULT NULL, " _
        & "Pro_cod varchar(30) DEFAULT NULL, " _
        & "Pro_Des varchar(250) DEFAULT NULL, " _
        & "degr_Can int(5) DEFAULT NULL, " _
        & "degr_PrU double(8,2) DEFAULT NULL, " _
        & "degr_PrT double(8,2) DEFAULT NULL, " _
        & "degr_Ppp double(8,2) DEFAULT NULL)"
        
Dim Cn As New ADODB.Connection
Cn.ConnectionString = Cadena
Cn.Open
Cn.Execute TaUsu
Adodc1.RecordSource = "Select * from Tdetproforma" & vusuariot
Adodc1.Refresh

errort:
If Err.Number = -2147217900 Then
    borrat = "Drop Table Tdetproforma" & vusuariot
    Cn.Execute borrat
    creatablapro
End If
End Function

Private Function grabadetproftemp()
Dim Cn As New ADODB.Connection
Dim rsdt As New ADODB.Recordset   ' Recordset de temporal de detalle
Dim rspr As New ADODB.Recordset   ' Recordset de productos
Cn.ConnectionString = Cadena
Cn.Open
vpro_cod = SSOleDBCombo1.Columns(1).Value

rspr.CursorType = adOpenKeyset
rspr.LockType = adLockOptimistic
rspr.ActiveConnection = Cn
rspr.Source = "Select * from productos Where Pro_cod = " & "'" & vpro_cod & "'"
rspr.Open

If Not rspr.EOF Then
    vpro_id = rspr!pro_id
Else
    vpro_id = 0
End If

vpro_des = SSOleDBCombo1.Columns(0).Value
vegr_can = Val(Text1.Text)
vegr_pru = Val(Text4.Text)
vtotal = vegr_can * vegr_pru
If SSCommand2.Caption = "Registrar" Then
    'Verifica si el item ya existe
    rsdt.CursorType = adOpenKeyset
    rsdt.LockType = adLockOptimistic
    rsdt.ActiveConnection = Cn
    rsdt.Source = "Select * from Tdetproforma" & vusuariot & " Where Pro_Des = " & "'" & vpro_des & "'"
    rsdt.Open

    If rsdt.EOF And SSCommand2.Caption = "Registrar" Then
        SSOleDBCombo1.Enabled = False
        nuevoe = "Insert into Tdetproforma" & vusuariot & " SET Pro_id = " & vpro_id & ", Pro_Cod = " & "'" & vpro_cod & "'" & ", Pro_Des = " & "'" & vpro_des & "', degr_Can = " & vegr_can & _
        ", degr_pru = " & vegr_pru & ", degr_PrT = " & vtotal
        Cn.Execute nuevoe
        Adodc1.Refresh
        Label18.Caption = Format(Val(Label18.Caption) + vtotal, "####.#0")
        
        limpiadatos
        SSFrame2.Enabled = True
        SSOleDBCombo1.Enabled = True
        SSOleDBCombo1.SetFocus
    Else
        MsgBox "Item ya registrado", vbInformation, empresa
        limpiadatos
        SSOleDBCombo1.SetFocus
    End If
 ElseIf SSCommand2.Caption = "Modificar" Then
    vpro_cod = SSOleDBGrid2.Columns(0).Value
    borrai = "Delete from Tdetinvinicial" & vusuariot & " Where pro_cod = " & "'" & vpro_cod & "'"
    Cn.Execute borrai
    
    nuevoe = "Insert into Tdetinvinicial" & vusuariot & " SET Pro_id = " & vpro_id & ", Pro_Cod = " & "'" & vpro_cod & "'" & ", Pro_Des = " & "'" & vpro_des & "', InI_Can = " & vini_can & _
    ", InI_PrU = " & vini_pru & ", InI_PrT = " & vIni_PrT & ", InI_Prv = " & vini_prv
    'nuevoe = "Insert into Tdetinvinicial" & vusuariot & " SET IteCod = " & vitecod & ", IteDes = " & "'" & vitedes & "', InICan = " & vIniCan & ", InIPrU = " & vIniPru & ", InIPrT = " & vIniPrT & ", InIPPP = " & vIniPru
    Cn.Execute nuevoe
    Adodc2.Refresh
    
    
    limpiadatos
    SSOleDBCombo1.SetFocus
    SSCommand2.Caption = "Registrar"
 End If
End Function
Private Function limpiadatos()
Text1.Text = ""
Text4.Text = ""
'Label8.Caption = ""
'Label11.Caption = ""
SSOleDBCombo1.Text = ""
End Function

Private Function grabaproforma()
Dim Cn As New ADODB.Connection
Dim rsci As New ADODB.Recordset   ' Recordset de cabecera de proforma
Dim rsdt As New ADODB.Recordset   ' Recordset de detalle temporal
Cn.ConnectionString = Cadena
Cn.Open
    

vini_fec = Format(Date, "yyyy-mm-dd")
vcliente = Text6.Text
'Graba cabecera de ingreso
grabacii = "INSERT INTO cabproforma SET cegr_fec = " & "'" & vini_fec & "', cegr_clie = " & "'" & vcliente & "', usu_id = " & vusucod
Cn.Execute grabacii

rsci.CursorType = adOpenKeyset
rsci.LockType = adLockOptimistic
rsci.ActiveConnection = Cn
rsci.Source = "Select * from cabproforma"
rsci.Open

If Not rsci.EOF Then
    rsci.MoveLast
    vini_id = rsci!cegr_id
Else
    vini_id = 1
End If

'Graba el id de cabecera en el detalle temporal
grabaid = "UPDATE Tdetproforma" & vusuariot & " SET cegr_Id = " & vini_id
Cn.Execute grabaid

'Graba detalle de ingreso de temporal a real
detar = " INSERT INTO detproforma SELECT * from Tdetproforma" & vusuariot
Cn.Execute detar
Adodc1.Refresh
Cn.Close
Cn.Open
' graba a reporte

borrarep = "DELETE FROM repproforma"
Cn.Execute borrarep

reporte = "INSERT INTO repproforma SELECT * FROM Tdetproforma" & vusuariot
Cn.Execute reporte
CrystalReport1.ReportFileName = App.Path & "\proforma.rpt"
CrystalReport1.Formulas(0) = "cliente = " & "'" & vcliente & "'"
CrystalReport1.Action = 1
End Function
