VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ABM_GRABACION 
   Caption         =   "Grabaciones"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdeliminar 
      Caption         =   "Eliminar Registro"
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox txtidgrab 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Text            =   "idgrab"
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox txtnrocds 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3735
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6588
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbloc 
      Height          =   315
      Left            =   720
      TabIndex        =   4
      Text            =   "localizacion"
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox txtidloc 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "idloc"
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Agregar Registro"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox txtfechfin 
      Height          =   285
      Left            =   240
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "fecha fin  "
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtfechini 
      Height          =   285
      Left            =   240
      MaxLength       =   10
      TabIndex        =   0
      Text            =   "fecha inicio"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblloc 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lbldd 
      Height          =   735
      Left            =   840
      TabIndex        =   9
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "cantidad de discos"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
   End
End
Attribute VB_Name = "ABM_GRABACION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub AgregaGrabacion()

If esNro(txtidloc) = True And esNro(txtnrocds) = True And ValidarFecha(txtfechini) = True And ValidarFecha(txtfechfin) = True Then
txtidloc = Val(txtidloc): txtnrocds = Val(txtnrocds)

    If AgregaRegGrabacion(txtidloc, txtnrocds, txtfechini, txtfechfin) = True Then
    MsgBox "Registro de grabacion agregado", vbInformation
    Set DataGrid1.DataSource = MostrarTodoGrabacion
    Exit Sub
    Else
    MsgBox "No se agrego el registro", vbExclamation
    End If
Else
MsgBox "No se agrego el registro", vbExclamation
End If
End Sub
Sub EliminaGrabacion(ByVal xidgrabacion As Double) 'implementar el paso del id

If EliminaRegGrabacion(xidgrabacion) = True Then
MsgBox "Registro eliminado", vbInformation
Set DataGrid1.DataSource = MostrarTodoGrabacion
Exit Sub
Else
MsgBox "No se elimino el registro", vbExclamation
End If

End Sub
Private Sub cmbloc_Click()
Dim nro As Integer
Dim xvar As String

xvar = Mid(cmbloc.Text, 1, Val(InStr(cmbloc.Text, "-") - 1))

nro = Val(xvar)

txtidloc.Text = nro

End Sub

Private Sub cmdAdd_Click()
AgregaGrabacion
End Sub

Private Sub cmdeliminar_Click()
eliminar
End Sub
Sub eliminar()

If EliminaRegGrabacion(txtidhist) = True Then
MsgBox "Se elimino"
End If
End Sub
Private Sub DataGrid1_DblClick()

Dim rs As New ADODB.Recordset
Set rs = MostrarRegGrabacion(DataGrid1.RowBookmark(DataGrid1.Row))
txtidloc = rs!idlocalizacion
txtfechini = rs!fechaini
txtfechfin = rs!fechafin
txtnrocds = rs!nrocds
End Sub

Private Sub Form_Load()
Set rs = MostrarTodoLoc
cmbloc.Clear
If Not rs.EOF Then
rs.MoveFirst
While Not rs.EOF
cmbloc.AddItem (rs!id & " - " & rs!Nombre)
rs.MoveNext
Wend
End If
Set DataGrid1.DataSource = MostrarTodoGrabacion

End Sub

Private Sub txtcantd_KeyPress(KeyAscii As Integer)

KeyAscii = CorrijeAscii(KeyAscii)
If KeyAscii = 13 Then
txtiddd.SetFocus
End If

End Sub

Private Sub txtfechfin_KeyPress(KeyAscii As Integer)
Dim xfech As Collection

8 KeyAscii = CorrijeAscii(KeyAscii)
If KeyAscii <> 8 Then
Set xfech = EscribeFecha(txtfechfin)
    If xfech.Count <> 0 Then
    txtfechfin = xfech.Item(1)
    txtfechfin.SelStart = xfech.Item(2)
    End If
    If KeyAscii = 13 And ValidarFecha(txtfechfin) = True Then
    txtnrocds.SetFocus
    End If
End If

End Sub


Private Sub txtfechini_KeyPress(KeyAscii As Integer)
 Dim xfech As Collection

8 KeyAscii = CorrijeAscii(KeyAscii)
If KeyAscii <> 8 Then
Set xfech = EscribeFecha(txtfechini)
    If xfech.Count <> 0 Then
    txtfechini = xfech.Item(1)
    txtfechini.SelStart = xfech.Item(2)
    End If
    If KeyAscii = 13 And ValidarFecha(txtfechini) = True Then
    txtfechfin.SetFocus
    End If
End If
End Sub

Private Sub txtiddd_KeyPress(KeyAscii As Integer)
Dim rs As New ADODB.Recordset
Dim nroserie As String
Dim capacidad As Double
Dim info As String
Dim tipo As String

If KeyAscii = 13 Then
    If esNro(txtiddd) = True Then
    txtiddd = Val(txtiddd)
    lbldd.Caption = ""
    Set rs = MostrarDD(Val(txtiddd))
        If Not rs.EOF Then
        nroserie = rs!nroserie
        capacidad = rs!capacidad
        tipo = rs!tipo
        info = "Nro serie: " & nroserie & vbCrLf & "Capacidad: " & capacidad & vbCrLf & "Tipo: " & tipo
        lbldd.Caption = info
        Else
        lbldd = "No se encontro registro"
        End If
    Else
txtiddd.Text = ""
    End If
End If

End Sub

Private Sub txtidloc_KeyPress(KeyAscii As Integer)
Dim rs As New ADODB.Recordset
Dim Nombre As String
Dim info As String

If KeyAscii = 13 Then

    If esNro(txtidloc) = True Then
    txtidloc = Val(txtidloc)
        Set rs = MostrarLoc(Val(txtidloc))
            If Not rs.EOF Then
            Nombre = rs!Nombre
            info = "Nombre: " & Nombre
            lblloc.Caption = info
            cmbloc.Text = rs!id & " - " & rs!Nombre
            Else
            lblloc.Caption = "no se encontro localizacion"
            End If
    Else
    txtidloc.Text = ""
    End If
End If
End Sub
 
