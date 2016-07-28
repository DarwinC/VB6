VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form BUSQUEDA 
   Caption         =   "Busqueda"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbloc 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Top             =   480
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5535
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9763
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
      Caption         =   "Resultados de busqueda"
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
   Begin VB.TextBox txtfechfin 
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar Grabacion"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtfechini 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtidloc 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Hasta:"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Desde:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha/s a buscar"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Localizacion:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "BUSQUEDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
