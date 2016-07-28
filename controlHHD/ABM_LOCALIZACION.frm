VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ABM_LOCALIZACION 
   Caption         =   "Localizacion"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4560
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txtLoc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   6
      Text            =   "localizacion nombre"
      Top             =   450
      Width           =   3735
   End
   Begin VB.TextBox txtIdloc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Text            =   "id localizacion"
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdcan 
      Caption         =   "&cancelar"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdelim 
      Caption         =   "e&liminar"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "&editar"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "&agregar"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6165
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "nombre"
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "id"
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   480
      Width           =   255
   End
End
Attribute VB_Name = "ABM_LOCALIZACION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()

If cmdAdd.Caption = "&agregar" Then
   txtidloc = MostrarIDActLoc
    cmdAdd.Caption = "&aceptar"
    cmdedit.Enabled = False
    cmdelim.Enabled = False
    cmdcan.Enabled = True
    txtloc.Enabled = True

Else
    cmdAdd.Caption = "&agregar"
    cmdedit.Enabled = True
    cmdelim.Enabled = True
    cmdcan.Enabled = False
    txtloc.Enabled = False
    If AgregaLoc(txtloc) = True Then
    MsgBox "localizacion agregada"
    Else
    MsgBox "no se realizo la operacion"
    End If
End If

Set DataGrid1.DataSource = MostrarTodoLoc
End Sub

Private Sub cmdcan_Click()
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
txtidloc.Enabled = False
txtloc.Enabled = False

End Sub
Private Sub cmdedit_Click()

If cmdedit.Caption = "&editar" Then
    cmdedit.Caption = "&aceptar"
    cmdedit.Enabled = True
    cmdelim.Enabled = False
    cmdAdd.Enabled = False
    cmdcan.Enabled = True
    txtloc.Enabled = True
    txtidloc.Enabled = True
Else
    cmdedit.Caption = "&editar"
    cmdAdd.Enabled = True
    cmdelim.Enabled = True
    cmdcan.Enabled = False
    txtloc.Enabled = False
    txtidloc.Enabled = False
    If esNro(txtidloc) = True Then
        If EditaLoc(txtidloc, txtloc) = True Then
        MsgBox "localizacion editada"
        Else
        MsgBox "no se realizo la operacion"
        End If
    End If
End If

Set DataGrid1.DataSource = MostrarTodoLoc

End Sub

Private Sub cmdelim_Click()

If cmdelim.Caption = "e&liminar" Then
    cmdelim.Caption = "&aceptar"
    cmdedit.Enabled = False
    cmdAdd.Enabled = False
    cmdcan.Enabled = True
    txtloc.Enabled = False
    txtidloc.Enabled = True
Else
    cmdelim.Caption = "e&liminar"
    cmdedit.Enabled = True
    cmdAdd.Enabled = True
    cmdcan.Enabled = False
    txtloc.Enabled = False
    txtidloc.Enabled = False
    If esNro(txtidloc) = True Then
        If EliminaLoc(txtidloc) = True Then
        MsgBox "localizacion eliminada"
        Else
        MsgBox "no se realizo la operacion"
        End If
    End If
End If

Set DataGrid1.DataSource = MostrarTodoLoc

End Sub

Private Sub Command1_Click()
Dim rs As New ADODB.Recordset
Set rs = CONECTAR.Execute("select * from localizacion")
MsgBox rs!id & " - " & rs!Nombre
End Sub

Private Sub DataGrid1_DblClick()
 Dim rs As New ADODB.Recordset
Set rs = MostrarLoc(DataGrid1.RowBookmark(DataGrid1.Row))
txtidloc = rs!id
txtloc = rs!Nombre
End Sub

Private Sub Form_Load()
Dim rsloc As New ADODB.Recordset

Set rsloc = MostrarTodoLoc

Set DataGrid1.DataSource = rsloc

End Sub

Private Sub txtidloc_KeyPress(KeyAscii As Integer)

Dim rs As New ADODB.Recordset
Dim Nombre As String
Dim idloc As String

If KeyAscii = 13 Then

    If esNro(txtidloc) = True Then
    txtidloc = Val(txtidloc)
    Set rs = MostrarLoc(Val(txtidloc))
            If Not rs.EOF Then
            Nombre = rs!Nombre
            txtloc.Text = Nombre
            Else
    txtloc.Text = "no se encontro"
            End If
    
    Else
    txtidloc = ""
    End If
    
End If
End Sub
