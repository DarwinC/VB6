Attribute VB_Name = "CONEXION"
Option Explicit
'CADENA DE CONEXION
Public Function cnn() As String
'PASAR PARAMETROS DE CONEXION REMOTA
Dim path As String 'archivo de la base de datos Access
path = LeerINI
'path = micamino
cnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & ";Persist Security Info=False"
End Function

'Escribe camino de la base de datos
Public Function EscribeINI(ByVal miCamino As String) As Boolean
Dim nro As Integer
Dim path As String
path = App.path & "\ControlDD.ini"
nro = FreeFile
Open path For Output As nro
Print #nro, miCamino
Close nro

EscribeINI = True

End Function

'Lee Hubicacion de la base de datos

Public Function LeerINI() As String

Dim nro As Integer
Dim Camino As String
nro = FreeFile
Dim path As String
path = App.path & "\ControlDD.ini"
On Error GoTo mierror
Open path For Input As #nro
Input #nro, Camino
Close #nro
LeerINI = Camino
Exit Function
mierror:
LeerINI = ""
End Function
'conexion a BD
Public Function CONECTAR() As ADODB.Connection
Dim conex As New ADODB.Connection
conex.Open (cnn)
conex.CursorLocation = adUseClient
Set CONECTAR = conex
End Function

'consultas sql
Public Function miConsulta(ByVal xsql As String) As ADODB.Recordset

Dim rs As New ADODB.Recordset

Set rs = CONECTAR.Execute(xsql)
Set miConsulta = rs

End Function

Public Function pruebaErroresSql(ByVal cadsql As String) As Boolean

On Error GoTo mierror
Dim rs As New ADODB.Recordset
Set rs = CONECTAR.Execute(cadsql)

pruebaErroresSql = False

If CBool(Err.Number) = True Then
mierror:
pruebaErroresSql = True
End If

End Function
