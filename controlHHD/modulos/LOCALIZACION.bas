Attribute VB_Name = "LOCALIZACION"
Option Explicit
'Asigna y actualiza IDs
Public Function AsignaIDLoc() As Double
Dim rs As New ADODB.Recordset
Dim rsid As New ADODB.Recordset
Dim idnueva As Double
Set rs = CONECTAR.Execute("select id from id where nombre='localizacion'")
If Not rs.EOF Then
idnueva = CInt(rs!id) + 1
Set rsid = CONECTAR.Execute("update id set id = " & idnueva & " where nombre = 'localizacion'")

AsignaIDLoc = idnueva 'incompleto
Exit Function
Else
Set rsid = CONECTAR.Execute("insert into id (id, nombre)values(1,'localizacion')")
AsignaIDLoc = 1
Exit Function
End If
End Function
'Asigna y actualiza IDs
Public Function MostrarIDActLoc() As Double
Dim rs As New ADODB.Recordset
Dim rsid As New ADODB.Recordset
Dim idnueva As Double

Set rs = CONECTAR.Execute("select id from id where nombre='localizacion'")
If Not rs.EOF Then
idnueva = CInt(rs!id) + 1
MostrarIDActLoc = idnueva
Exit Function
Else
MostrarIDActLoc = 1
Exit Function
End If
End Function

'Agrega una nueva localizacion
Public Function AgregaLoc(ByVal xnombre As String) As Boolean
Dim varid As Double
Dim varnombre As String
Dim strsql As String
Dim rs As New ADODB.Recordset

varid = AsignaIDLoc ' asigna y actualiza id
varnombre = xnombre

strsql = "insert into localizacion(id,nombre)Values(" & varid & ",'" & varnombre & "')"
Set rs = CONECTAR.Execute(strsql)
AgregaLoc = True

End Function

'Edita localizacion
Public Function EditaLoc(ByVal xid As Double, ByVal xnombre As String) As Boolean

Dim rs As New ADODB.Recordset
Dim rsfind As New ADODB.Recordset

Set rsfind = CONECTAR.Execute("select * from localizacion where id=" & xid)

If Not rsfind.EOF Then
Set rs = CONECTAR.Execute("update localizacion set nombre= '" & xnombre & "' where id=" & xid)
EditaLoc = True
Else
EditaLoc = False
End If

End Function
'Elimina una localizacion
Public Function EliminaLoc(ByVal xid As Double) As Boolean

Dim strsql As String
Dim rs As New ADODB.Recordset
Dim rsfind As New ADODB.Recordset

Set rsfind = CONECTAR.Execute("select * from localizacion where id=" & xid)

If Not rsfind.EOF Then
strsql = "delete * from localizacion where id=" & xid
Set rs = CONECTAR.Execute(strsql)
EliminaLoc = True
Else
EliminaLoc = False
End If

End Function

'Muestra informacion de una localizacion especificada por id
Public Function MostrarLoc(ByVal xidloc As Double) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim varidloc As Double
varidloc = xidloc

Set rs = CONECTAR.Execute("select * from localizacion where id=" & varidloc)

Set MostrarLoc = rs

End Function
'Muestra informacion de tabla localizacion
Public Function MostrarTodoLoc() As ADODB.Recordset
Dim rs As New ADODB.Recordset
Set rs = CONECTAR.Execute("select * from localizacion")
Set MostrarTodoLoc = rs

End Function

