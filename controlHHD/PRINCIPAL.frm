VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form PRINCIPAL 
   Caption         =   "Disco Duro Seguimiento"
   ClientHeight    =   5280
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11445
   Icon            =   "PRINCIPAL.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGrab 
      Caption         =   "Fecha Grabada"
      Height          =   495
      Left            =   2400
      TabIndex        =   17
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmddd 
      Caption         =   "disco"
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdconloc 
      Caption         =   "localizacion"
      Height          =   375
      Left            =   2400
      TabIndex        =   15
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Grabaciones"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox txtfech 
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton cmdingresa 
      Caption         =   "In&gresar Estado"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdverHistorial 
      Caption         =   "&Historiales"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdverLoc 
      Caption         =   "&Localizaciones"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   1815
   End
   Begin VB.ComboBox cmbloc 
      Height          =   315
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox txtidloc 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "idloc"
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton cmdverDiscos 
      Caption         =   "&Discos"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox txtiddd 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "id dd"
      Top             =   240
      Width           =   495
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4575
      Left            =   3840
      TabIndex        =   0
      Top             =   480
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8070
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
      Caption         =   "Seguimiento Discos"
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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3000
      Top             =   4200
   End
   Begin VB.Label lblhrs 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   10
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblloc 
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label lbldd 
      Height          =   975
      Left            =   720
      TabIndex        =   6
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label lblfech 
      Caption         =   "fecha:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.Menu mnuArch 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuexpinfo 
         Caption         =   "exportar informe"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edicion"
      Begin VB.Menu mnuDisco 
         Caption         =   "Discos"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuLocalizacion 
         Caption         =   "Localizaciones"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuhistorial 
         Caption         =   "Historiales"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuGrabaciones 
         Caption         =   "Grabaciones"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnugui 
         Caption         =   "_"
      End
   End
   Begin VB.Menu mnuver 
      Caption         =   "&Ver"
      Begin VB.Menu mnuverDiscos 
         Caption         =   "Discos"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuLoc 
         Caption         =   "Localizaciones"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuhistoriales 
         Caption         =   "Historiales"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuVerGrabaciones 
         Caption         =   "Grabaciones"
         Shortcut        =   ^J
      End
   End
   Begin VB.Menu mnuConsulta 
      Caption         =   "C&onsultas"
      Begin VB.Menu mnuConsultaSql 
         Caption         =   "ConsultaSql"
      End
      Begin VB.Menu mnuBuscar 
         Caption         =   "Buscar"
      End
      Begin VB.Menu munguion 
         Caption         =   "_"
      End
   End
End
Attribute VB_Name = "PRINCIPAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbloc_Click()
Dim nro As Integer
Dim xvar As String

xvar = Mid(cmbloc.Text, 1, Val(InStr(cmbloc.Text, "-") - 1))
nro = Val(xvar)
txtidloc.Text = nro

End Sub
Sub salir()
Dim xresp As Integer
xresp = MsgBox("Salir del programa", vbYesNo + vbQuestion, "Atencion")
If xresp = 6 Then
Unload Me
End If
End Sub

Private Sub cmdconloc_Click()
Dim xidloc As String
xidloc = InputBox("Ingrese el nro de la localizacion", "Historial", 0)
If esNro(xidloc) = True Then
Set DataGrid1.DataSource = MostrarHistLoc(Val(xidloc))
End If

End Sub

Private Sub cmddd_Click()
Dim xiddd As String
xiddd = InputBox("Ingrese el nro de disco", "Historial", 0)
If esNro(xiddd) = True Then
Set DataGrid1.DataSource = MostrarHistDD(Val(xiddd))
End If
End Sub

Private Sub cmdGrab_Click()
ABM_GRABACION.Show
End Sub

Private Sub cmdingresa_Click()
Agregarhist
txtiddd.SetFocus
End Sub
Sub Agregarhist()

If esNro(txtiddd) = True And esNro(txtidloc) = True And (ValidarFecha(txtfech) = True Or txtfech = "") Then
txtiddd = Val(txtiddd): txtidloc = Val(txtidloc)
    If AgregaHistorial(Val(txtiddd), Val(txtidloc), txtfech.Text) = True Then
    MsgBox "historial de seguimiento agregado"
    VerHistorial
    End If
    
Else
MsgBox "debe llenar todos los campos con valores validos"
End If

End Sub

Private Sub cmdsalir_Click()
salir
End Sub
Private Sub cmdverDiscos_Click()
VerDiscos
End Sub
Sub VerDiscos()
Set DataGrid1.DataSource = MostrarTodoDD
End Sub

Private Sub cmdverHistorial_Click()
VerHistorial
End Sub

Sub VerHistorial()
Set DataGrid1.DataSource = MostrarTodoHistorialNombres
End Sub

Private Sub cmdverLoc_Click()
VerLocalizaciones
End Sub
Sub VerLocalizaciones()
Set DataGrid1.DataSource = MostrarTodoLoc
End Sub

Private Sub Command1_Click()
VerGrabaciones
End Sub
Sub VerGrabaciones()
Set DataGrid1.DataSource = MostrarTodoGrabacion
End Sub

'esta parte es para correcciones
Sub corrije()
'**********************************************
'funcion para cuando se borraron los id
Dim piff As DAO.Database
Dim pifee As DAO.Recordset
Set piff = OpenDatabase("c:/dd.mdb")
Set pifee = piff.OpenRecordset("select * from historial where fechaini='17/03/2010'")
pifee.MoveFirst

'Dim xvar As Integer: xvar = 0
While Not pifee.EOF
'xvar = xvar + 1
pifee.Edit
pifee!fechaini = "16/03/2010"
pifee.Update
pifee.MoveNext
Wend
pifee.Close
'***********************************************
End Sub
Private Sub DataGrid1_DblClick()
Dim rs As New ADODB.Recordset

MsgBox DataGrid1.Columns(O).Text

'Set rs = MostrarTodoHistorialNombres(DataGrid1.RowBookmark(DataGrid1.Row))

End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset

Set rs = MostrarTodoLoc
cmbloc.Clear
If Not rs.EOF Then
rs.MoveFirst
While Not rs.EOF
cmbloc.AddItem (rs!id & " - " & rs!Nombre)
rs.MoveNext
Wend
End If

Set DataGrid1.DataSource = MostrarTodoHistorialNombres

End Sub

Private Sub mnuBuscar_Click()
BUSQUEDA.Show
End Sub

Private Sub mnuConsultaSql_Click()
CONSULTA.Show
End Sub

Private Sub mnuDisco_Click()
ABM_DD.Show
End Sub

Private Sub mnuGrabaciones_Click()
ABM_GRABACION.Show
End Sub

Private Sub mnuHistorial_Click()
ABM_HISTORIAL.Show
End Sub

Private Sub mnuhistoriales_Click()
VerHistorial
End Sub

Private Sub mnulistDisc_Click()
Set DataGrid1.DataSource = MostrarTodoDD
End Sub

Private Sub mnuLoc_Click()
VerLocalizaciones
End Sub

Private Sub mnuLocalizacion_Click()
ABM_LOCALIZACION.Show
End Sub

Private Sub mnuSalir_Click()
salir
End Sub

Private Sub mnuverDiscos_Click()
VerDiscos
End Sub

Private Sub mnuVerGrabaciones_Click()
VerGrabaciones
End Sub

Private Sub Timer1_Timer()
lblhrs = Time
End Sub

Private Sub txtfech_KeyPress(KeyAscii As Integer)

Dim xfech As Collection

KeyAscii = CorrijeAscii(KeyAscii)
If KeyAscii <> 8 Then
Set xfech = EscribeFecha(txtfech)
    If xfech.Count <> 0 Then
    txtfech = xfech.Item(1)
    txtfech.SelStart = xfech.Item(2)
    End If
    If KeyAscii = 13 Then
        If txtfech = "" And esNro(txtidloc) = True And esNro(txtiddd) = True Then
        txtidloc = Val(txtidloc): txtiddd = Val(txtiddd)
        txtfech = UltimaFechaEnLoc(txtiddd, txtidloc)
        cmdingresa.SetFocus
        End If
    End If
End If
End Sub
Private Sub txtidloc_KeyPress(KeyAscii As Integer)

Dim rs As New ADODB.Recordset
Dim xnombre As String
Dim info As String

If KeyAscii = 13 Then
    
    If esNro(txtidloc) = True Then
    txtidloc = Val(txtidloc)
        Set rs = MostrarLoc(Val(txtidloc))
            If Not rs.EOF Then
            xnombre = rs!Nombre
            info = "Nombre: " & xnombre
            lblloc.Caption = info
            cmbloc.Text = rs!id & " - " & rs!Nombre
            txtfech.SetFocus
            Else
            lblloc.Caption = "no se encontro localizacion"
            End If
    Else
    txtidloc.Text = ""
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
    lbldd.Caption = ""
    txtiddd = Val(txtiddd)
    Set rs = MostrarDD(Val(txtiddd))
        If Not rs.EOF Then
        nroserie = rs!nroserie
        capacidad = rs!capacidad
        tipo = rs!tipo
        info = "Nro serie: " & nroserie & vbCrLf & "Capacidad: " & capacidad & vbCrLf & "Tipo: " & tipo
        lbldd.Caption = info
        txtidloc.SetFocus
        Else
        lbldd = "No se encontro registro"
        End If
    Else
txtiddd.Text = ""
    End If
End If

End Sub
