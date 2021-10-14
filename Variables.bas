Attribute VB_Name = "Variables"
'Variables de coneccion
Public Cn As New ADODB.Connection
'Public Const Cadena = "server = localhost;driver=MySQL ODBC 3.51 Driver;db=comercio;UID=root;PWD=cagisa"
Public Const Cadena = "server = 192.168.0.173;driver=MySQL ODBC 3.51 Driver;db=comercio;UID=root;PWD=cagisa"
Public Const empresa = "MAUI AND SONS"
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public resp As Long

'Variable de caracter
Public vUsuario, vusuariot, origen, origenn, vUsuNiv, vPacMat, vprotip, vInttur, vpro_cod As String
Public vitecod, valmsig As String
Public modo, modoini, modocom, modoven, modopro As String
'Variables numericas
Public seleccion, seleccioni, vusucod, siaccesos, siexiste, venccod, vRegId, vMedId, vTurId, cambioturno As Integer
Public uactivado, vconini, vflag, turno, vini_id, valmid As Integer


'Variables Booleanas
Public VarOrder As Boolean
