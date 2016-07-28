Attribute VB_Name = "Historial"
Option Explicit
'Asigna id de historico
Private Function AsignaIDHist() As Double
Dim rs As New ADODB.Recordset
Dim rsid As New ADODB.Recordset
Dim idnueva As Double
Set rs = CONECTAR.Execute("select id from id where nombre='historial'")
If Not rs.EOF Then
idnueva = CInt(rs!id) + 1

Set rsid = CONECTAR.Execute("update id set id = " & idnueva & " where nombre = 'historial'")

AsignaIDHist = idnueva
Exit Function
Else
Set rsid = CONECTAR.Execute("insert into id (id, nombre)values(1,'historial')")
AsignaIDHist = 1
Exit Function
End If
End Function
'Agrega un nuevo registro historico
Public Function AgregaHistorial(ByVal xiddd As Double, ByVal xidloc As Double, ByVal xfech As String) As Boolean

Dim xfechafin, xfechaini, strsql, xvar As String: xvar = 0
Dim xidHistAsig As Double
Dim rsInsert As New ADODB.Recordset
Dim rsUpdate As New ADODB.Recordset
Dim rsfind As New ADODB.Recordset
Set rsfind = CONECTAR.Execute("select * from historial where iddisco =" & xiddd)

'Determinamos la fecha de finalizacion en la localizacion anterior

    If Not rsfind.EOF Then
        rsfind.MoveLast
        If Not xfech = "" Then
            If rsfind!fechafin = "" And CDate(xfech) >= CDate(rsfind!fechaini) Then
            Set rsUpdate = CONECTAR.Execute("update historial set fechafin='" & xfech & "' where idhist=" & CInt(rsfind!idhist))
            Else
            AgregaHistorial = False
            Exit Function
            End If
        Else
            If CDate(rsfind!fechaini) <= Date Then
            Set rsUpdate = CONECTAR.Execute("update historial set fechafin='" & Str(Date) & "' where idhist=" & CInt(rsfind!idhist))
            Else
            AgregaHistorial = False
            Exit Function
            End If
        End If
    End If
'se agrega el nuevo registro historico
If Not xfech = "" Then
    xidHistAsig = AsignaIDHist
    xfechafin = ""
    xfechaini = xfech
    strsql = "INSERT INTO historial(idhist,iddisco,idlocalizacion,fechaini,fechafin)Values( " & xidHistAsig & "," & xiddd & ",'" & xidloc & "', '" & xfechaini & "','" & xfechafin & "')"
    Set rsInsert = CONECTAR.Execute(strsql)
    AgregaHistorial = True
Else
    xidHistAsig = AsignaIDHist
    xfechafin = ""
    xfechaini = Str(Date)
    strsql = "INSERT INTO historial(idhist,iddisco,idlocalizacion,fechaini,fechafin)Values( " & xidHistAsig & "," & xiddd & ",'" & xidloc & "', '" & xfechaini & "','" & xfechafin & "')"
    Set rsInsert = CONECTAR.Execute(strsql)
    AgregaHistorial = True
End If

End Function

'Edita historico seleccionado por id
Public Function EditaHistorial(ByVal xidhist As Double, ByVal xiddd As Double, ByVal xidloc As Double, ByVal xfechaini As String, ByVal xfechafin As String) As Boolean
Dim rs As New ADODB.Recordset
Dim rsfind As New ADODB.Recordset
Set rsfind = CONECTAR.Execute("select * from historial where idhist=" & xidhist)

If Not rsfind.EOF Then
Set rs = CONECTAR.Execute("update historial set iddisco= " & xiddd & " , idlocalizacion= " & xidloc & " , fechaini = '" & xfechaini & "', fechafin ='" & xfechafin & "' where idhist=" & xidhist)
EditaHistorial = True
Else
EditaHistorial = False
End If

End Function

'Elimina un registro historico
Public Function EliminaHistorial(ByVal xidhist As Double) As Boolean

Dim varid As Double
Dim strsql As String
Dim rsfind As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Set rsfind = CONECTAR.Execute("select * from hitorial where idhist=" & xidhist)

If Not rsfind.EOF Then
strsql = "delete * from historial where id=" & xidhist
Set rs = CONECTAR.Execute(strsql)
EliminaHistorial = True
Else
EliminaHistorial = False
End If

End Function

'Muestra historial de un DD indicado por id
Public Function MostrarHistDD(ByVal xiddd As Double) As ADODB.Recordset

Dim rs As New ADODB.Recordset

Set rs = CONECTAR.Execute("select * from historial where iddisco=" & xiddd & " order by idhist desc")

Set MostrarHistDD = rs

End Function
'Muestra historial de una localizacion indicada por id
Public Function MostrarHistLoc(ByVal xidloc As Double) As ADODB.Recordset

Dim rs As New ADODB.Recordset

Set rs = CONECTAR.Execute("select * from historial where idlocalizacion=" & xidloc & " order by idhist desc")

Set MostrarHistLoc = rs

End Function

'mostrar historial indicado por id
Public Function MostrarHistorial(ByVal xidhist As Double) As ADODB.Recordset
Dim rs As New ADODB.Recordset

Set rs = CONECTAR.Execute("select * from historial where idhist=" & xidhist)
Set MostrarHistorial = rs
End Function
'mostrar tabla historial
Public Function MostrarTodoHistorial() As ADODB.Recordset
Dim rs As New ADODB.Recordset
Set rs = CONECTAR.Execute("select * from historial order by idhist")
Set MostrarTodoHistorial = rs
End Function

'mostrar tabla historial con nombres
Public Function MostrarTodoHistorialNombres() As ADODB.Recordset
Dim rs As New ADODB.Recordset
Set rs = CONECTAR.Execute("select a.nombre,c.iddisco,c.fechaini,c.fechafin from localizacion a, disco b, historial c where a.id = c.idlocalizacion and b.id=c.iddisco order by c.idhist desc")
Set MostrarTodoHistorialNombres = rs

End Function
'mostrar ultima fecha de disco en localidad
Public Function UltimaFechaEnLoc(ByVal xiddd As Double, ByVal xidloc As Double) As String
Dim rs As New ADODB.Recordset
Dim xfech As String
xfech = 0
Set rs = CONECTAR.Execute("select fechafin from historial where iddisco=" & xiddd & " and idlocalizacion = " & xidloc)
If Not rs.EOF Then
rs.MoveFirst
While Not rs.EOF
    If rs!fechafin > xfech Then
    xfech = rs!fechafin
    End If
rs.MoveNext
Wend
UltimaFechaEnLoc = xfech
Else
UltimaFechaEnLoc = 0
End If

End Function

