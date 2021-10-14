VERSION 5.00
Begin VB.Form Procesos 
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   11580
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Caption         =   "Carga exitencia a tabla productos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   315
      TabIndex        =   1
      Top             =   1035
      Width           =   4245
   End
   Begin VB.Label Label1 
      Caption         =   "Asignacion de codigo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   630
      Width           =   4245
   End
End
Attribute VB_Name = "Procesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
Dim Cn As New ADODB.Connection
Dim rsgr As New ADODB.Recordset
Dim rsprt As New ADODB.Recordset
Dim rspro As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open

rsgr.CursorType = adOpenKeyset
rsgr.LockType = adLockOptimistic
rsgr.ActiveConnection = Cn
rsgr.Source = "Select * from productos"
rsgr.Open


Do While Not rsgr.EOF
    vprotnu = rsgr!ProTNu
    vprotli = rsgr!ProTLi
    vprotip = Left(rsgr!Pro_Tip, 1)
    vprocodb = rsgr!pro_codb
    rsgr!pro_cod = vprocodb & vprotnu & vprotli & vprotip
    rsgr.Update
    rsgr.MoveNext
Loop
MsgBox "Proceso Terminado !!"
End Sub

Private Sub Label2_Click()
Dim Cn As New ADODB.Connection
Dim rsgr As New ADODB.Recordset
Dim rsprt As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open

rsgr.CursorType = adOpenKeyset
rsgr.LockType = adLockOptimistic
rsgr.ActiveConnection = Cn
rsgr.Source = "Select * from tie"
rsgr.Open


Do While Not rsgr.EOF
    vprocod = rsgr!procod
    vprocan = rsgr!procan
    act = "UPDATE productos Set Pro_exi = " & vprocan & " WHERE Pro_Cod = " & "'" & vprocod & "'"
    Cn.Execute act
    rsgr.MoveNext
Loop
MsgBox "Proceso Terminado !!"
End Sub
