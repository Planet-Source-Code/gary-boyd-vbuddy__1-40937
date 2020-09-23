VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "VBuddy"
   ClientHeight    =   1920
   ClientLeft      =   6015
   ClientTop       =   5115
   ClientWidth     =   2640
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   2640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "2. Quit"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1. New Game"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Your Virtual Buddy!"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Welcome to VBuddy 1.0"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmFarm.Show
Unload Me
End Sub

Private Sub Command2_Click()
End
End Sub
