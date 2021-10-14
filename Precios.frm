VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form Precios 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9720
   ClientLeft      =   6240
   ClientTop       =   3255
   ClientWidth     =   10905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSCommand SSCommand6 
      Height          =   420
      Left            =   9180
      TabIndex        =   0
      Top             =   8595
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4185
      Top             =   2025
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
      RecordSource    =   "Select * From productos ORDER By pro_des"
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
      Bindings        =   "Precios.frx":0000
      Height          =   2490
      Left            =   810
      TabIndex        =   2
      Top             =   720
      Width           =   9840
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
      Columns(0).Width=   2461
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "pro_cod"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "pro_cod"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   7541
      Columns(1).Caption=   "Producto"
      Columns(1).Name =   "pro_Des"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "pro_Des"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2434
      Columns(2).Caption=   "Pr. Venta"
      Columns(2).Name =   "pro_pve"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   1
      Columns(2).DataField=   "pro_pve"
      Columns(2).DataType=   5
      Columns(2).NumberFormat=   "####.#0"
      Columns(2).FieldLen=   256
      _ExtentX        =   17357
      _ExtentY        =   4392
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
   Begin Threed.SSFrame SSFrame3 
      Height          =   2400
      Left            =   405
      TabIndex        =   3
      Top             =   3330
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   4233
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.TextBox Text3 
         Height          =   330
         Left            =   8685
         TabIndex        =   21
         Top             =   1575
         Width           =   1275
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   1140
         Left            =   225
         TabIndex        =   4
         Top             =   135
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   2011
         _Version        =   196608
         BackStyle       =   1
         ClipControls    =   0   'False
         Begin VB.TextBox Text1 
            Height          =   330
            Left            =   405
            TabIndex        =   5
            Top             =   270
            Width           =   2985
         End
         Begin Threed.SSCommand SSCommand2 
            Height          =   330
            Left            =   1125
            TabIndex        =   6
            Top             =   675
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   582
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
            Caption         =   "Buscar"
            BevelWidth      =   1
         End
         Begin VB.Label Label16 
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
            Left            =   270
            TabIndex        =   7
            Top             =   45
            Width           =   600
         End
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   1140
         Left            =   4950
         TabIndex        =   8
         Top             =   135
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   2011
         _Version        =   196608
         BackStyle       =   1
         ClipControls    =   0   'False
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   90
            TabIndex        =   9
            Top             =   270
            Width           =   4875
         End
         Begin Threed.SSCommand SSCommand1 
            Height          =   330
            Left            =   1620
            TabIndex        =   10
            Top             =   720
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   582
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
            Caption         =   "Buscar"
            BevelWidth      =   1
         End
         Begin VB.Label Label2 
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
            Left            =   135
            TabIndex        =   11
            Top             =   45
            Width           =   780
         End
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   330
         Left            =   8685
         TabIndex        =   22
         Top             =   1980
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
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
         Caption         =   "Grabar"
         BevelWidth      =   1
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nuevo precio"
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
         Left            =   8730
         TabIndex        =   20
         Top             =   1350
         Width           =   1155
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   7020
         TabIndex        =   19
         Top             =   1575
         Width           =   1545
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
         Left            =   7065
         TabIndex        =   18
         Top             =   1350
         Width           =   1365
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1980
         TabIndex        =   17
         Top             =   1575
         Width           =   4965
      End
      Begin VB.Label Label12 
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
         Left            =   315
         TabIndex        =   16
         Top             =   1350
         Width           =   600
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   270
         TabIndex        =   15
         Top             =   1575
         Width           =   1545
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
         Left            =   1980
         TabIndex        =   14
         Top             =   1350
         Width           =   780
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   1005
      Left            =   405
      TabIndex        =   12
      Top             =   5805
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   1773
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.TextBox Text5 
         Height          =   330
         Left            =   5940
         TabIndex        =   31
         Top             =   360
         Width           =   1275
      End
      Begin VB.TextBox Text4 
         Height          =   330
         Left            =   4050
         TabIndex        =   28
         Top             =   360
         Width           =   1275
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   270
         TabIndex        =   23
         Top             =   405
         Width           =   2670
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   330
         Left            =   8730
         TabIndex        =   29
         Top             =   360
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
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
         Caption         =   "Grabar"
         BevelWidth      =   1
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Montro en Bs."
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
         Left            =   5985
         TabIndex        =   30
         Top             =   135
         Width           =   1200
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   360
         Left            =   4455
         TabIndex        =   27
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione grupo"
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
         TabIndex        =   24
         Top             =   180
         Width           =   1500
      End
   End
   Begin Threed.SSFrame SSFrame5 
      Height          =   1050
      Left            =   405
      TabIndex        =   13
      Top             =   6840
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   1852
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.TextBox Text7 
         Height          =   330
         Left            =   5985
         TabIndex        =   36
         Top             =   360
         Width           =   1275
      End
      Begin VB.TextBox Text6 
         Height          =   330
         Left            =   4095
         TabIndex        =   33
         Top             =   360
         Width           =   1275
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   270
         TabIndex        =   25
         Top             =   450
         Width           =   2670
      End
      Begin Threed.SSCommand SSCommand5 
         Height          =   330
         Left            =   8730
         TabIndex        =   34
         Top             =   315
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
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
         Caption         =   "Grabar"
         BevelWidth      =   1
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Montro en Bs."
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
         Left            =   6030
         TabIndex        =   35
         Top             =   135
         Width           =   1200
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   360
         Left            =   4500
         TabIndex        =   32
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione marca"
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
         TabIndex        =   26
         Top             =   225
         Width           =   1530
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIOS"
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
      Left            =   9585
      TabIndex        =   1
      Top             =   135
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   60
      Left            =   4365
      Picture         =   "Precios.frx":0015
      Stretch         =   -1  'True
      Top             =   495
      Width           =   6540
   End
End
Attribute VB_Name = "Precios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_GotFocus()
Dim Cn As New ADODB.Connection
Dim rsgr As New ADODB.Recordset   ' Recordset de grupos
Cn.ConnectionString = Cadena
Cn.Open
       
rsgr.CursorType = adOpenKeyset
rsgr.LockType = adLockOptimistic
rsgr.ActiveConnection = Cn
rsgr.Source = "Select * from grupos ORDER BY Gru_des"
rsgr.Open

If Not rsgr.EOF Then
    Combo1.Clear
    Do While Not rsgr.EOF
        Combo1.AddItem rsgr!Gru_des
        rsgr.MoveNext
    Loop
End If
Cn.Close
End Sub

Private Sub Combo2_GotFocus()
Dim Cn As New ADODB.Connection
Dim rsma As New ADODB.Recordset   ' Recordset de marca
Cn.ConnectionString = Cadena
Cn.Open
       
rsma.CursorType = adOpenKeyset
rsma.LockType = adLockOptimistic
rsma.ActiveConnection = Cn
rsma.Source = "Select * from marcas ORDER BY Mar_des"
rsma.Open

If Not rsma.EOF Then
    Combo2.Clear
    Do While Not rsma.EOF
        Combo2.AddItem rsma!Mar_Des
        rsma.MoveNext
    Loop
End If
Cn.Close
End Sub

Private Sub SSCommand1_Click()
Dim indice, db1 As String
db1 = Text2.Text
Adodc1.RecordSource = "SELECT * from productos WHERE Pro_Des LIKE " & "'%" & db1 & "%'"
Adodc1.Refresh
End Sub

Private Sub SSCommand2_Click()
Dim indice, db1 As String
db1 = Text1.Text
Adodc1.RecordSource = "SELECT * from productos WHERE Pro_Des LIKE " & "'%" & db1 & "%'"
Adodc1.Refresh
End Sub

Private Sub SSCommand3_Click()
If Len(Trim(Text3.Text)) > 0 Then
    nprecio = Val(Text3.Text)
    vpro_cod = Label11.Caption
    Dim Cn As New ADODB.Connection
    Cn.ConnectionString = Cadena
    Cn.Open

    cambiapr = "UPDATE productos SET pro_pve = " & nprecio & " WHERE pro_Cod = " & "'" & vpro_cod & "'"
    Cn.Execute cambiapr
    MsgBox "Precio actualizado", vbInformation, empresa
    Cn.Close
    limpiadatospr
    Adodc1.Refresh
End If
End Sub
Private Sub SSCommand4_Click()
Dim Cn As New ADODB.Connection
Dim rspr As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open
      
If Len(Trim(Text4.Text)) > 0 Then
    vmarca = Combo1.Text
    vporc = Val(Text4.Text)
    rspr.CursorType = adOpenKeyset
    rspr.LockType = adLockOptimistic
    rspr.ActiveConnection = Cn
    rspr.Source = "Select * from vproducto Where Gru_Des = " & "'" & vmarca & "'"
    rspr.Open
    If Not rspr.EOF Then
        Do While Not rspr.EOF
            If Not rspr!pro_pve = 0 Then
                porcm = ((rspr!pro_pve * vporc) / 100)
                rspr!pro_pve = rspr!pro_pve + porcm
                rspr.Update
            End If
            rspr.MoveNext
        Loop
        
    End If
End If

If Len(Trim(Text5.Text)) > 0 Then
    vmarca = Combo1.Text
    vbs = Val(Text5.Text)
    grabanp = "UPDATE vproducto SET pro_pve = pro_pve + " & vbs & " WHERE Gru_Des = " & "'" & vmarca & "'"
    Cn.Execute grabanp
End If

MsgBox "Precio actualizado", vbInformation, empresa
Cn.Close
Combo1.Text = ""
Text4.Text = ""
Text5.Text = ""
Adodc1.Refresh
End Sub

Private Sub SSCommand5_Click()
Dim Cn As New ADODB.Connection
Dim rspr As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open
      
If Len(Trim(Text6.Text)) > 0 Then
    vmarca = Combo2.Text
    vporc = Val(Text6.Text)
    rspr.CursorType = adOpenKeyset
    rspr.LockType = adLockOptimistic
    rspr.ActiveConnection = Cn
    rspr.Source = "Select * from vproducto Where Mar_Des = " & "'" & vmarca & "'"
    rspr.Open
    If Not rspr.EOF Then
        Do While Not rspr.EOF
            If Not rspr!pro_pve = 0 Then
                porcm = ((rspr!pro_pve * vporc) / 100)
                rspr!pro_pve = rspr!pro_pve + porcm
                rspr.Update
            End If
            rspr.MoveNext
        Loop
        
    End If
End If

If Len(Trim(Text7.Text)) > 0 Then
    vmarca = Combo2.Text
    vbs = Val(Text7.Text)
    grabanp = "UPDATE vproducto SET pro_pve = pro_pve + " & vbs & " WHERE Mar_Des = " & "'" & vmarca & "'"
    Cn.Execute grabanp
End If
MsgBox "Precio actualizado", vbInformation, empresa
Cn.Close
Combo2.Text = ""
Text6.Text = ""
Text7.Text = ""
Adodc1.Refresh

End Sub

Private Sub SSCommand6_Click()
Menup.Enabled = True
Menup.Label8.ForeColor = &H8000000F
Unload Precios
Set Precios = Nothing
End Sub
Private Sub SSOleDBGrid1_Click()
seleccion = 1
Label11.Caption = Adodc1.Recordset.Fields("Pro_cod")
Label8.Caption = Adodc1.Recordset.Fields("Pro_Des")
Label4.Caption = Format(Adodc1.Recordset.Fields("Pro_Pve"), "####.#0")
End Sub
Private Sub Text1_GotFocus()
Text1.BackColor = &H80FFFF
End Sub
Private Sub Text1_LostFocus()
Text1.BackColor = &HFFFFFF
Text1.Text = UCase(Text1.Text)
End Sub
Private Sub Text2_GotFocus()
Text2.BackColor = &H80FFFF
End Sub
Private Sub Text2_LostFocus()
Text2.BackColor = &HFFFFFF
Text2.Text = UCase(Text2.Text)
End Sub
Private Sub Text3_GotFocus()
Text3.BackColor = &H9FBDEA
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub
Private Sub Text3_LostFocus()
Text3.BackColor = &HFFFFFF
Text3.Text = Format(Text3.Text, "####.#0")
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 46 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub Text4_GotFocus()
If Len(Trim(Combo1.Text)) > 0 Then
    Text4.BackColor = &H9FBDEA
    Text4.SelStart = 0
    Text4.SelLength = Len(Text4.Text)
    Text5.Text = ""
Else
    MsgBox "Debe seleccionar el grupo", vbInformation, empresa
    Combo1.SetFocus
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

Private Sub Text5_GotFocus()
If Len(Trim(Combo1.Text)) > 0 Then
    Text5.BackColor = &H9FBDEA
    Text5.SelStart = 0
    Text5.SelLength = Len(Text5.Text)
    Text4.Text = ""
Else
    MsgBox "Debe seleccionar el grupo", vbInformation, empresa
    Combo1.SetFocus
End If
End Sub
Private Sub Text5_LostFocus()
Text5.BackColor = &HFFFFFF
Text5.Text = Format(Text5.Text, "####.#0")
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 46 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub Text6_GotFocus()
If Len(Trim(Combo2.Text)) > 0 Then
    Text6.BackColor = &H9FBDEA
    Text6.SelStart = 0
    Text6.SelLength = Len(Text6.Text)
    Text7.Text = ""
Else
    MsgBox "Debe seleccionar la marca", vbInformation, empresa
    Combo2.SetFocus
End If
End Sub
Private Sub Text6_LostFocus()
Text6.BackColor = &HFFFFFF
Text6.Text = Format(Text6.Text, "####.#0")
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 46 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub Text7_GotFocus()
If Len(Trim(Combo2.Text)) > 0 Then
    Text7.BackColor = &H9FBDEA
    Text7.SelStart = 0
    Text7.SelLength = Len(Text7.Text)
    Text6.Text = ""
Else
    MsgBox "Debe seleccionar la marca", vbInformation, empresa
    Combo2.SetFocus
End If
End Sub
Private Sub Text7_LostFocus()
Text7.BackColor = &HFFFFFF
Text7.Text = Format(Text7.Text, "####.#0")
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 46 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Function limpiadatospr()
Label4.Caption = ""
Label8.Caption = ""
Label11.Caption = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo1.Text = ""
Combo2.Text = ""
End Function
