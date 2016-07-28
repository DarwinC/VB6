Attribute VB_Name = "DD"
Option Explicit
'Asigna y actualiza IDs
Public Function AsignaIDDD() As Double
Dim rs As New ADODB.Recordset
Dim rsid As New ADODB.Recordset
Dim idnueva As Double
Set rs = CONECTAR.Execute("select id from id where nombre='disco'")
If Not rs.EOF Then
idnueva = CInt(rs!id) + 1

Set rsid = CONECTAR.Execute("update id set id = " & idnueva & " where nombre = 'disco'")

AsignaIDDD = idnueva
Exit Function
Else
Set rsid = CONECTAR.Execute("insert into id (id, nombre)values(1,'disco')")
AsignaIDDD = 1
Exit Function
End If
End Function
'Asigna y actualiza IDs
Public Function MostrarIDAct() As Double
Dim rs As New ADODB.Recordset
Dim rsid As New ADODB.Recordset
Dim idnueva As Double

Set rs = CONECTAR.Execute("select id from id where nombre='disco'")
If Not rs.EOF Then
idnueva = CInt(rs!id) + 1
MostrarIDAct = idnueva
Exit Function
Else
MostrarIDAct = 1
Exit Function
End If
End Function
'Agrega un nuevo DD
Public Function AgregaDD(ByVal xnroserie As String, ByVal xcapacidad As Double, ByVal xtipo As String) As Boolean
Dim varid, varcapacidad As Double
Dim varnroserie, vartipo, strsql As String
Dim rs As New ADODB.Recordset

varid = AsignaIDDD ' asigna y actualiza id
varnroserie = xnroserie
varcapacidad = xcapacidad
vartipo = xtipo

strsql = "INSERT INTO disco(id,nroserie,capacidad, tipo)Values(" & varid & ",'" & varnroserie & "', " & varcapacidad & ",'" & vartipo & "')"
Set rs = CONECTAR.Execute(strsql)
AgregaDD = True

End Function

'Edita DD
Public Function EditaDD(ByVal xid As Double, ByVal xnroserie As String, ByVal xcapacidad As Double, ByVal xtipo As String) As Boolean
Dim rs As New ADODB.Recordset
Dim rsfind As New ADODB.Recordset

Set rsfind = CONECTAR.Execute("select * from disco where id=" & xid)

If Not rsfind.EOF Then
Set rs = CONECTAR.Execute("update disco set nroserie= '" & xnroserie & "' , capacidad= " & xcapacidad & " , tipo = '" & xtipo & "' where id = " & xid)
EditaDD = True
Else
EditaDD = False
End If
End Function
'Elimina un DD
Public Function EliminaDD(ByVal xid As Double) As Boolean
Dim varid As Double
Dim strsql As String
Dim rs As New ADODB.Recordset
Dim rsfind As New ADODB.Recordset

Set rsfind = CONECTAR.Execute("select * from disco where id=" & xid)

If Not rsfind.EOF Then
strsql = "delete * from disco where id=" & xid
Set rs = CONECTAR.Execute(strsql)
EliminaDD = True
Else
EliminaDD = False
End If

End Function

'Muestra informacion de un DD indicado por id
Public Function MostrarDD(xid) As ADODB.Recordset
Dim rs As New ADODB.Recordset

Set rs = CONECTAR.Execute("select * from disco where id=" & xid)
Set MostrarDD = rs

End Function
'muestra tabla discos
Public Function MostrarTodoDD() As ADODB.Recordset

Dim rs As New ADODB.Recordset

Set rs = CONECTAR.Execute("select * from disco")
If Not rs.EOF Then
Set MostrarTodoDD = rs
Else
Exit Function 'devolver error
End If

End Function

