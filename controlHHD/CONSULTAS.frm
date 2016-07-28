VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CONSULTA 
   Caption         =   "Consultas"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4215
      Left            =   480
      TabIndex        =   9
      Top             =   3000
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   7435
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
   Begin VB.CommandButton Command2 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   9720
      TabIndex        =   8
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ejecutar consulta"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtcon 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "consulta sql"
      Top             =   480
      Width           =   10575
   End
   Begin VB.Label Label7 
      Caption         =   "Tabla grabaciones: campo ""idsuc"", campo ""fechaini"", campo ""fechafin"""
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   2400
      Width           =   6855
   End
   Begin VB.Label Label6 
      Caption         =   "Tabla historial: campo ""iddisco"", campo ""idlocalizacion"", campo ""fechaini"", campo ""fechafin"", campo ""idhist"""
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   2160
      Width           =   7815
   End
   Begin VB.Label Label5 
      Caption         =   "Tabla id: campo ""id"", campo ""nombre"""
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Tabla localizacion: campo ""id"", campo ""nombre"""
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   1680
      Width           =   5535
   End
   Begin VB.Label Label3 
      Caption         =   "Tabla disco: campo ""id"", campo ""nroserie"", campo ""capacidad"", campo ""tipo"""
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   1440
      Width           =   5775
   End
   Begin VB.Label Label2 
      Caption         =   "Estructura de tablas:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "consulta Sql:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "CONSULTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()

Dim rs As New ADODB.Recordset

If pruebaErroresSql(txtcon) = False Then
Set rs = miConsulta(txtcon)
    If Not rs.EOF Then
    Set DataGrid1.DataSource = rs
    Else
    MsgBox "SIN REGISTROS"
    End If
Else
MsgBox "Error en la sintaxis de la consulta"
End If

End Sub
 
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DataGrid1_DblClick()
MsgBox DataGrid1.Columns(0).Text
MsgBox DataGrid1.se
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
MsgBox DataGrid1.Columns(0).Text
End If
End Sub
