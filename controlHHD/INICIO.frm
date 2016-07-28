VERSION 5.00
Begin VB.Form INICIO 
   Caption         =   "Form1"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   1560
      TabIndex        =   4
      Top             =   240
      Width           =   3015
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   3600
      Width           =   855
   End
End
Attribute VB_Name = "INICIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
GuardaINI
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Dir1_Change()
File1.path = Dir1.path
End Sub

Private Sub Dir1_Click()
File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1.Drive
File1.path = Dir1.path
End Sub
Private Sub File1_DblClick()
GuardaINI
End Sub
Sub GuardaINI()
EscribeINI (File1.path & "\" & File1.FileName)
PRINCIPAL.Show
Unload Me
End Sub
Private Sub Form_Load()
File1.path = Dir1.path
Dir1.path = Drive1.Drive
End Sub
