Attribute VB_Name = "ERROR"
Option Explicit

Public Type ClaseError
NroError As Integer
DescripcionError As String
Aviso As Boolean
End Type
Public Function mierror(ByVal coderror As String) As String
'Dim mierror As String

End Function
Public Function esNro(ByVal xvar As String) As Boolean

Dim xvarobj As String
xvarobj = xvar

If Val(xvarobj) = 0 Then

esNro = False

Else

esNro = True

End If

End Function
Public Function CorrijeAscii(ByVal KeyAscii As Integer) As Integer
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13 Then
CorrijeAscii = KeyAscii
Else
CorrijeAscii = 0
End If

End Function
Public Function EscribeFecha(ByVal unString As String) As Collection

Dim xlen As Integer
Dim xresultados As New Collection

xlen = Len(unString)
    Select Case Len(unString)
    
    Case 2
    
    unString = unString + "/"
    xresultados.Add (unString)
    xlen = xlen + 1
    xresultados.Add (xlen)
    Set EscribeFecha = xresultados
    
    Case 5
    
    unString = unString + "/"
    
    xresultados.Add (unString)
    xlen = xlen + 1
    xresultados.Add (xlen)
    Set EscribeFecha = xresultados
    
    Case Else
    
    Set EscribeFecha = xresultados
    
    End Select

End Function

Public Function ValidarFecha(ByVal Fecha As String) As Boolean
    Dim Dia, Mes, Año As Integer
    
On Error GoTo MErr

'mejorar la funcion de validar fecha
    
    Dia = Val(Mid(Fecha, 1, 2))
    
    Mes = Val(Mid(Fecha, 4, 2))
    
    If Len(Fecha) = 8 Then
        Año = Val(Mid(Fecha, 7, 2))
    Else
        Año = Val(Mid(Fecha, 7, 4))
    End If
    
    Dim mesa(12) As Integer
        mesa(1) = 31
        mesa(2) = 28 + Abs(Año Mod 4 = 0)
        mesa(3) = 31
        mesa(4) = 30
        mesa(5) = 31
        mesa(6) = 30
        mesa(7) = 31
        mesa(8) = 31
        mesa(9) = 30
        mesa(10) = 31
        mesa(11) = 30
        mesa(12) = 31
    ValidarFecha = Not ((Mes > 12 Or Mes < 1) Or (Dia > mesa(Mes)))
    
    Exit Function
MErr:
    ValidarFecha = False
End Function

