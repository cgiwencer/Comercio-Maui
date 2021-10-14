Attribute VB_Name = "Funciones"
Public Function accesos()
Dim Cn As New ADODB.Connection
Dim rspr As New ADODB.Recordset   ' Recordset procesos
Cn.ConnectionString = Cadena
Cn.Open

'Procesos
rspr.CursorType = adOpenKeyset
rspr.LockType = adLockOptimistic
rspr.ActiveConnection = Cn
rspr.Source = "Select * from procesos"
rspr.Open

rspr.MoveFirst
acceso = Abs(rspr.Fields("prcest"))
rspr.MoveNext
Do While Not rspr.EOF
    acceso = acceso & Abs(rspr.Fields("prcest"))
    rspr.MoveNext
Loop
End Function

Public Function verificaacceso()
Dim Cn As New ADODB.Connection
Dim rspr As New ADODB.Recordset   ' Recordset procesos
Dim rsus As New ADODB.Recordset   ' Recordset usuarios
Cn.ConnectionString = Cadena
Cn.Open

'Procesos
rsus.CursorType = adOpenKeyset
rsus.LockType = adLockOptimistic
rsus.ActiveConnection = Cn
rsus.Source = "Select * from usuarios where UsuCod = " & vusucod
rsus.Open

If Mid(rsus!UsuAcc, numproceso, 1) = 1 Then
    siacceso = 1
Else
    siacceso = 0
End If

End Function

Public Function exmin()
Dim Cn As New ADODB.Connection
Dim rsfa As New ADODB.Recordset   ' Recordset de faltantes
Cn.ConnectionString = Cadena
Cn.Open

borra = "DELETE FROM faltantes"
graba = "INSERT INTO faltantes (itecod,itedes,iteUnM,itemin,itesal) SELECT pro_Cod, Pro_Des, Pro_uni, pro_exMi, Pro_exi From productos where Pro_exi < Pro_ExMi"
Cn.Execute borra
Cn.Execute graba


rsfa.CursorType = adOpenKeyset
rsfa.LockType = adLockOptimistic
rsfa.ActiveConnection = Cn
rsfa.Source = "Select * from faltantes"
rsfa.Open

If Not rsfa.EOF Then
    MsgBox "Existen productos con existencia baja...!!!" & vbCrLf & _
    "Por favor revise el listado haciendo click sobre el aviso en la pantalla principal", vbCritical, empresa
    'Menup.Image30.Visible = True
    'Menup.Image8.Picture = LoadPicture(App.Path & "\imagenes\fondotienda.jpg")
End If
End Function
Public Function cierrausuario()
On Error GoTo erroru

Dim Cn As New ADODB.Connection
Cn.ConnectionString = Cadena
Cn.Open

cierrau = "UPDATE usuarios set usulog = " & 0 & " Where UsuId = " & vusucod
Cn.Execute cierrau

Cn.Close

erroru:
If Err.Number = -2147217900 Then
    Unload Menup
    Set Menup = Nothing
End If

End Function
