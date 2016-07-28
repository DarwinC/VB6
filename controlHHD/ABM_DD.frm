VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ABM_DD 
   Caption         =   "Discos Duros"
   ClientHeight    =   5190
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtiddd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Text            =   "id"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdcan 
      Caption         =   "&cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "&editar"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdelim 
      Caption         =   "e&liminar"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "&agregar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3015
      Left            =   0
      TabIndex        =   3
      Top             =   2160
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5318
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
      Caption         =   "Discos Registrados"
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
   Begin VB.ComboBox cmbtipo 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "ABM_DD.frx":0000
      Left            =   2280
      List            =   "ABM_DD.frx":0010
      TabIndex        =   2
      Text            =   "IDE"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtcap 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Text            =   "capacidad"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txtnros 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Text            =   "nro de serie"
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label lbltipodd 
      Alignment       =   1  'Right Justify
      Caption         =   "tipo"
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lblcapdd 
      Alignment       =   1  'Right Justify
      Caption         =   "capacidad"
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblnrsdd 
      Alignment       =   1  'Right Justify
      Caption         =   "nro serie"
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lbliddd 
      Alignment       =   1  'Right Justify
      Caption         =   "id"
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.Menu mnuAccion 
      Caption         =   "Acc&iones"
      Begin VB.Menu mnuAgregar 
         Caption         =   "Agregar"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditar 
         Caption         =   "Editar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuEliminar 
         Caption         =   "Eliminar"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuCancelar 
         Caption         =   "Cancelar"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuguion 
         Caption         =   "_"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "ABM_DD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click() 'AGREGAR
Agregar
End Sub

Sub Agregar()

If cmdAdd.Caption = "&agregar" Then
txtiddd = MostrarIDAct()
    cmdAdd.Caption = "&aceptar"
    cmdelim.Enabled = False
    cmdedit.Enabled = False
    cmdcan.Enabled = True
    txtiddd.Enabled = False
    txtcap.Enabled = True
    cmbtipo.Enabled = True
    txtnros.Enabled = True

Else

    cmdAdd.Caption = "&agregar"
    cmdelim.Enabled = True
    cmdedit.Enabled = True
    cmdcan.Enabled = False
    txtcap.Enabled = False
    cmbtipo.Enabled = False
    txtnros.Enabled = False
    
    If esNro(txtcap) = True Then
        If AgregaDD(txtnros, Val(txtcap), cmbtipo) = True Then
        MsgBox "agregado"
        Else
        MsgBox "no se agrego"
        End If
    End If

End If
Set DataGrid1.DataSource = MostrarTodoDD
End Sub

Private Sub cmdcan_Click()
Cancelar
End Sub

Private Sub cmdedit_Click() 'EDITAR
Editar (Val(txtiddd))
End Sub

Sub Editar(ByVal xiddd As Double)
If cmdedit.Caption = "&editar" Then

    cmdedit.Caption = "&aceptar"
    cmdelim.Enabled = False
    cmdAdd.Enabled = False
    cmdcan.Enabled = True
    txtiddd.Enabled = True
    txtcap.Enabled = True
    cmbtipo.Enabled = True
    txtnros.Enabled = True
    txtiddd = xiddd
Else
    
    cmdedit.Caption = "&editar"
    cmdelim.Enabled = True
    cmdAdd.Enabled = True
    cmdcan.Enabled = False
    txtcap.Enabled = False
    cmbtipo.Enabled = False
    txtnros.Enabled = False
    txtiddd.Enabled = False
    If esNro(txtcap) = True And txtnros <> "" And cmbtipo.Text <> "" Then
        If EditaDD(txtiddd, txtnros, txtcap, cmbtipo.Text) = True Then
        MsgBox "editado"
        Else
        MsgBox "no se edito"
        End If
    End If
End If
Set DataGrid1.DataSource = MostrarTodoDD
End Sub
Private Sub cmdelim_Click() 'BORRAR
Eliminar (Val(txtiddd))
End Sub
Sub Eliminar(ByVal xiddd As Double)

If cmdelim.Caption = "e&liminar" Then

    cmdelim.Caption = "&aceptar"
    cmdAdd.Enabled = False
    cmdedit.Enabled = False
    cmdcan.Enabled = True
    txtiddd.Enabled = True
    txtcap.Enabled = False
    cmbtipo.Enabled = False
    txtnros.Enabled = False
Else

    cmdelim.Caption = "e&liminar"
    cmdAdd.Enabled = True
    cmdedit.Enabled = True
    cmdcan.Enabled = False
    txtcap.Enabled = False
    cmbtipo.Enabled = False
    txtnros.Enabled = False
    txtiddd.Enabled = False

    If esNro(txtcap) = True Then
        If EliminaDD(txtiddd) = True Then
        MsgBox "eliminado"
        Else
        MsgBox "no se elimino"
        End If
    End If
End If

Set DataGrid1.DataSource = MostrarTodoDD

End Sub
Private Sub cmdVerTodo_click()
Set DataGrid1.DataSource = MostrarTodoDD
End Sub

Private Sub DataGrid1_DblClick()

 Dim rs As New ADODB.Recordset
Set rs = MostrarDD(DataGrid1.RowBookmark(DataGrid1.Row))
txtiddd = rs!id
txtcap = rs!capacidad
txtnros = rs!nroserie
cmbtipo = rs!tipo

End Sub

Private Sub Form_Load()
Set DataGrid1.DataSource = MostrarTodoDD
End Sub

Private Sub Label1_Click()

End Sub

Private Sub mnuAgregar_Click()
Agregar
End Sub

Private Sub mnuCancelar_Click()
Cancelar
End Sub
Sub Cancelar()

cmdAdd.Caption = "&agregar"
cmdedit.Caption = "&editar"
cmdelim.Caption = "e&liminar"
cmdAdd.Enabled = True
cmdedit.Enabled = True
cmdelim.Enabled = True
cmdcan.Enabled = False
txtiddd.Enabled = False
txtcap.Enabled = False
txtnros.Enabled = False
cmbtipo.Enabled = False

End Sub
Private Sub mnuEditar_Click()
Editar (Val(txtiddd))
End Sub

Private Sub mnuEliminar_Click()
Eliminar (Val(txtiddd))
End Sub

Private Sub mnuSalir_Click()
Unload Me
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
    Set rs = MostrarDD(Val(txtiddd))
        If Not rs.EOF Then
        nroserie = rs!nroserie
        capacidad = rs!capacidad
        tipo = rs!tipo
        txtnros = nroserie
        txtcap = capacidad
        cmbtipo.Text = tipo
        Else
        txtnros = "no se encontro registro"
        txtcap = ""
        End If
    Else
txtiddd.Text = ""
    End If
End If
End Sub

