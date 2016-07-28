Attribute VB_Name = "Grabaciones"
Option Explicit
Public Function Imprimime(ByVal xrs As ADODB.Recordset) As String
'imprime
End Function
'Asigna id de grabacion
Public Function AsignaIDGrabacion() As Double
Dim rs As New ADODB.Recordset
Dim rsid As New ADODB.Recordset
Dim idnueva As Double
Set rs = CONECTAR.Execute("select id from id where nombre='grabacion'")
If Not rs.EOF Then
idnueva = CInt(rs!id) + 1
Set rsid = CONECTAR.Execute("update id set id = " & idnueva & " where nombre = 'grabacion'")
AsignaIDGrabacion = idnueva
Exit Function
Else
Set rsid = CONECTAR.Execute("insert into id (id, nombre)values(1,'grabacion')")
AsignaIDGrabacion = 1
Exit Function
End If
End Function
'agrega fechas grabadas
Public Function AgregaRegGrabacion(ByVal xidloc As Integer, ByVal xnrocds As Integer, ByVal xfechaini As String, ByVal xfechafin As String) As Boolean
Dim rs As New ADODB.Recordset
Dim rsfind As New ADODB.Recordset
Dim xsql As String
Dim xidgrabacion As Double

xsql = "select * from grabacion where idlocalizacion = " & xidloc
    If Not (CDate(xfechaini) > CDate(xfechafin)) Then
        Set rsfind = CONECTAR.Execute(xsql)
        If Not rsfind.EOF Then
        rsfind.MoveFirst
            '********************************************************
            'bucle de comprobacion de fechas
            While rsfind.EOF = False
                If CDate(xfechaini) < CDate(rsfind!fechafin) Then
                    If CDate(xfechafin) > CDate(rsfind!fechaini) Then
                    AgregaRegGrabacion = False
                    Exit Sub
                    End If
                End If
            Next
            '*********************************************************
        Else
            xidgrabacion = AsignaIDGrabacion
            Set rs = CONECTAR.Execute("insert into grabacion (idgrabacion, idlocalizacion, nrocds, fechaini, fechafin)values(" & xidgrabacion & "," & xidloc & "," & xnrocds & ",'" & xfechaini & "','" & xfechafin & "')")
            AgregaRegGrabacion = True
            Exit Function
        End If
    Else
        AgregaRegGrabacion = False
        Exit Function
    End If
End Function
'Mostrar todo el registro de grabaciones
Public Function MostrarTodoGrabacion() As ADODB.Recordset
Dim rs As New ADODB.Recordset
Set rs = CONECTAR.Execute("select a.nombre, b.fechaini, b.fechafin, b.nrocds, b.idgrabacion from grabacion b, localizacion a where b.idlocalizacion=a.id order by b.idgrabacion")
Set MostrarTodoGrabacion = rs
End Function
Public Function MostrarRegGrabacion(ByVal xidgrab As Double) As ADODB.Recordset

Dim rsfind As New ADODB.Recordset
Dim xsql As String

xsql = "select * from grabacion where idgrabacion=" & xidgrab

Set rsfind = CONECTAR.Execute(xsql)

Set MostrarRegGrabacion = rsfind

End Function
'elimina registro de grabaciones
Public Function EliminaRegGrabacion(ByVal xidgrabacion As Double) As Boolean
Dim rsfind As New ADODB.Recordset
Dim xsql As String
xsql = "select * from grabacion where idgrabacion=" & xidgrabacion
'set rsfind
End Function

'Edita grabaciones
Public Function EditaRegGrabacion(ByVal xidgrabacion As Integer, ByVal xidloc As Integer, ByVal xnrocds As Integer, ByVal xfechaini As String, ByVal xfechafin As String) As Boolean
Dim rs As New ADODB.Recordset
Dim rsfind As New ADODB.Recordset
Dim xsql As String
xsql = "select * from grabacion where xidgrabacion = " & xidgrabacion
    If Not (CDate(xfechaini) > CDate(xfechafin)) Then
        Set rsfind = CONECTAR.Execute(xsql)
        If Not rsfind.EOF Then
        rsfind.MoveFirst
            If xfechaini >= rsfind!fechafin Then
            Set rs = CONECTAR.Execute("update grabacion set idlocalizacion= " & xidloc & ", nrocds=" & xnrocds & ", fechaini='" & xfechafin & "', fechafin='" & xfechaini & "'")
            EditaRegGrabacion = True
            Else
            EditaRegGrabacion = False
            End If
        Else
            Set rs = CONECTAR.Execute("update grabacion set idlocalizacion= " & xidloc & ", nrocds=" & xnrocds & ", fechaini='" & xfechafin & "', fechafin='" & xfechaini & "'")
            EditaRegGrabacion = True
        End If
    Else
        EditaRegGrabacion = False
    End If
End Function

