VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ABM_HISTORIAL 
   Caption         =   "Historial"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdver 
      Caption         =   "Ver historico"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton cmdhistloc 
      Caption         =   "localizacion historial"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtidhist 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   14
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "editar"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdelim 
      Caption         =   "eliminar"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdcan 
      Caption         =   "cancelar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdhistdd 
      Caption         =   "disco historial"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
   Begin VB.ComboBox cmbloc 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3000
      TabIndex        =   9
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox txtidloc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtfechfin 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtfechini 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtiddd 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   3360
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7646
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
            LCID            =   14346
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
            LCID            =   14346
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
   Begin VB.Label lblinfo 
      Height          =   255
      Left            =   3120
      TabIndex        =   20
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label lblloc 
      Height          =   255
      Left            =   5760
      TabIndex        =   19
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label lbldd 
      Height          =   735
      Left            =   3600
      TabIndex        =   18
      Top             =   480
      Width           =   6375
   End
   Begin VB.Label Label5 
      Caption         =   "nro historial"
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "id loc"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "fecha final"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "fecha inicio"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "nro disco"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "ABM_HISTORIAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdedit_Click()
Editar
End Sub
Private Sub cmdelim_Click()
eliminar
End Sub
Private Sub cmdhistdd_Click()
Dim xaux As Double
Dim xiddd As String
xiddd = InputBox("Ingrese el nro de disco", "Historial", 0)
If esNro(xiddd) = True Then
Set DataGrid1.DataSource = MostrarHistDD(Val(xiddd))
End If
End Sub
Sub MostrarHistorialDisco()
'Set DataGrid1.DataSource = MostrarHistDD
End Sub
Private Sub cmdhistloc_Click()
Dim xidloc As String
xidloc = InputBox("Ingrese el nro de la localizacion", "Historial", 0)
If esNro(xidloc) = True Then
Set DataGrid1.DataSource = MostrarHistLoc(Val(xidloc))
End If
'MotrarHistorialLocalizacion
End Sub
Private Sub cmdhistsimple_Click()
MostrarHist2
End Sub
Sub MostrarHist2()
Set DataGrid1.DataSource = MostrarTodoHistorialNombres
End Sub
Private Sub cmdtblhist_Click()
MostrarHist1
End Sub
Sub MostrarHist1()
Set DataGrid1.DataSource = MostrarTodoHistorial
End Sub

Private Sub cmbloc_Click()
Dim nro As Integer
Dim xvar As String
xvar = Mid(cmbloc.Text, 1, Val(InStr(cmbloc.Text, "-") - 1))
nro = Val(xvar)
txtidloc.Text = nro
End Sub

Sub Editar()

If cmdedit.Caption = "editar" Then
txtfechini.Enabled = True
txtfechfin.Enabled = True
txtiddd.Enabled = True
txtidloc.Enabled = True
txtidhist.Enabled = True
cmdcan.Enabled = True
cmdedit.Caption = "aceptar"
cmdelim.Enabled = False
cmbloc.Enabled = True
Else
txtfechini.Enabled = False
txtfechfin.Enabled = False
txtiddd.Enabled = False
txtidloc.Enabled = False
txtidhist.Enabled = False
cmdcan.Enabled = False
cmdedit.Caption = "editar"
cmdelim.Enabled = True
cmbloc.Enabled = False
    If (ValidarFecha(txtfechini) = True) And (ValidarFecha(txtfechfin) = True Or txtfechfin = "") Then
        If EditaHistorial(Val(txtidhist), Val(txtiddd), Val(txtidloc), txtfechini, txtfechfin) = True Then
        MsgBox "Se ha editado con exito", vbInformation
        Set DataGrid1.DataSource = MostrarTodoHistorial
        cmdedit.SetFocus
        
        Else
        MsgBox "No se pudo editar: Datos no validos", vbExclamation
        cmdedit.SetFocus
        End If
    Else
    MsgBox "No se pudo editar: Error en fecha", vbExclamation
    cmdedit.SetFocus
    End If
End If

End Sub
Sub eliminar()

If cmdedit.Caption = "eliminar" Then

txtfechini.Enabled = True
txtfechfin.Enabled = True
txtiddd.Enabled = True
txtidloc.Enabled = True
txtidhist.Enabled = True
cmdcan.Enabled = True
cmdelim.Caption = "aceptar"
cmdedit.Enabled = False
cmbloc.Enabled = True
Else

txtfechini.Enabled = False
txtfechfin.Enabled = False
txtiddd.Enabled = False
txtidloc.Enabled = False
txtidhist.Enabled = False
cmdcan.Enabled = False
cmdedit.Enabled = True
cmdelim.Caption = "eliminar"
cmbloc.Enabled = False
    If EliminaHistorial(Val(txtidhist)) = True Then
    MsgBox "Eliminado"
    cmdelim.SetFocus
    Else
    MsgBox "No se pudo eliminar"
    cmdelim.SetFocus
    End If
End If

End Sub
Sub Cancelar()
txtfechini.Enabled = False
txtfechfin.Enabled = False
txtiddd.Enabled = False
txtidloc.Enabled = False
txtidhist.Enabled = False
cmdcan.Enabled = False
cmdedit.Caption = "editar"
cmdelim.Caption = "eliminar"
End Sub

Private Sub cmdver_Click()
MostrarHist1
End Sub

Private Sub DataGrid1_DblClick()

Dim rs As New ADODB.Recordset
Set rs = MostrarHistorial(DataGrid1.RowBookmark(DataGrid1.Row))
txtidhist = rs!idhist
txtiddd = rs!iddisco
txtidloc = rs!idlocalizacion
txtfechini = rs!fechaini
txtfechfin = rs!fechafin

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

MostrarHist1

End Sub


Private Sub txtfechfin_KeyPress(KeyAscii As Integer)
Dim xfech As Collection

KeyAscii = CorrijeAscii(KeyAscii)
If KeyAscii <> 8 Then
Set xfech = EscribeFecha(txtfechfin)
    If xfech.Count <> 0 Then
    txtfechfin = xfech.Item(1)
    txtfechfin.SelStart = xfech.Item(2)
    End If
    If KeyAscii = 13 And ValidarFecha(txtfechfin) = True Then
        If cmdedit.Enabled = True Then cmdedit.SetFocus
        If cmdedit.Enabled = True Then cmdelim.SetFocus
    End If
End If
End Sub

Private Sub txtfechini_KeyPress(KeyAscii As Integer)
Dim xfech As Collection

KeyAscii = CorrijeAscii(KeyAscii)
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
    lbldd.Caption = ""
    Set rs = MostrarDD(txtiddd)
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

Private Sub txtiddd_LostFocus()
Dim rs As New ADODB.Recordset
Dim nroserie As String
Dim capacidad As Double
Dim info As String
Dim tipo As String

    If esNro(txtiddd) = True Then
    lbldd.Caption = ""
    Set rs = MostrarDD(txtiddd)
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
End Sub
Private Sub txtidhist_KeyPress(KeyAscii As Integer)

Dim rs As New ADODB.Recordset
lblinfo.Caption = ""
If KeyAscii = 13 Then

    If esNro(txtidhist) = True Then
    txtidhist = Val(txtidhist)
        Set rs = MostrarHistorial(Val(txtidhist))
            If Not rs.EOF Then
            lblinfo.Caption = "Existe"
            txtiddd = rs!iddisco
            txtidloc = rs!idlocalizacion
            txtfechini = rs!fechaini
            txtfechfin = rs!fechafin
            txtiddd.SetFocus
            Else
            lblinfo.Caption = "No se encontro historial"
            End If
    Else
    txtidhist.Text = ""
    End If
End If
End Sub

Private Sub txtidhist_LostFocus()

Dim rs As New ADODB.Recordset
lblinfo.Caption = ""

If esNro(txtidhist) = True Then
    txtidhist = Val(txtidhist)
        Set rs = MostrarLoc(Val(txtidhist))
            If Not rs.EOF Then
            Set DataGrid1.DataSource = rs
            lblinfo.Caption = "Existe"
            Else
            lblinfo.Caption = "No se encontro historial"
            End If
    Else
    txtidhist.Text = ""
    End If
End Sub

Private Sub txtidloc_KeyPress(KeyAscii As Integer)

Dim rs As New ADODB.Recordset

If KeyAscii = 13 Then

    If esNro(txtidloc) = True Then
    txtidloc = Val(txtidloc)
        Set rs = MostrarLoc(Val(txtidloc))
            If Not rs.EOF Then
            cmbloc.Text = rs!id & " - " & rs!Nombre
            txtfechini.SetFocus
            lblloc.Caption = ""
            Else
            lblloc.Caption = "no se encontro localizacion"
            End If
    Else
    txtidloc.Text = ""
    End If
End If
End Sub

Private Sub txtidloc_LostFocus()
Dim rs As New ADODB.Recordset
Dim Nombre As String
Dim info As String

 If esNro(txtidloc) = True Then
    txtidloc = Val(txtidloc)
        Set rs = MostrarLoc(Val(txtidloc))
            If Not rs.EOF Then
            Nombre = rs!Nombre
            info = "Nombre: " & Nombre
            lblloc.Caption = info
            Else
            lblloc.Caption = "no se encontro localizacion"
            End If
    Else
    txtidloc.Text = ""
    End If
End Sub
