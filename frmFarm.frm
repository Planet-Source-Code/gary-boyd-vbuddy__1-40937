VERSION 5.00
Begin VB.Form frmFarm 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBuddy Farm"
   ClientHeight    =   3315
   ClientLeft      =   5370
   ClientTop       =   765
   ClientWidth     =   4320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   4320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Index           =   200
      Interval        =   1500
      Left            =   240
      Top             =   240
   End
   Begin VB.Frame framVBS 
      Caption         =   "VBuddy Stats"
      Height          =   2775
      Left            =   4080
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label lblAge 
         Caption         =   "0"
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Age:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblNrg 
         Caption         =   "0"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Energy:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblHapy 
         Caption         =   "0"
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Happynes:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblMoney 
         Caption         =   "0"
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Money:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblHP 
         Caption         =   "0"
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Health:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.PictureBox picVBuddy 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   1800
      Picture         =   "frmFarm.frx":0000
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   485
      TabIndex        =   0
      Top             =   1320
      Width           =   485
   End
   Begin VB.Menu mfile 
      Caption         =   "&File"
      Begin VB.Menu mnew 
         Caption         =   "New"
      End
      Begin VB.Menu mdash1 
         Caption         =   "-"
      End
      Begin VB.Menu mquit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mvbuddy 
      Caption         =   "&VBuddy"
      Begin VB.Menu mfeed 
         Caption         =   "Feed"
      End
      Begin VB.Menu mplay 
         Caption         =   "Play"
      End
      Begin VB.Menu mwork 
         Caption         =   "Work"
      End
      Begin VB.Menu mline2 
         Caption         =   "-"
      End
      Begin VB.Menu mstatus 
         Caption         =   "Status"
      End
   End
End
Attribute VB_Name = "frmFarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Age As Integer
Dim Happy As Integer
Dim Energy As Integer
Dim Health As Integer
Dim Money As Integer
Dim TimeRedo As Integer
Dim TimeAge As Integer

Private Sub Command1_Click()
framVBS.Visible = False
End Sub

Sub DeathCheck()
If Health = 0 Then
MsgBox "Your VBuddy has died of bad health.", vbOKOnly, "VBuddy Farm"
If Age <= 30 Then
MsgBox "Your rating: Bad parent", vbOKOnly, "VBuddy Farm"
ElseIf Age <= 50 Then
MsgBox "Your rating: Ok parent", vbOKOnly, "VBuddy Farm"
ElseIf Age <= 76 Then
MsgBox "Your rating: Ok parent", vbOKOnly, "VBuddy Farm"
End If
frmMain.Show
Unload Me
End If
If Age >= 76 Then
MsgBox "Your VBuddy has died of old age", vbOKOnly, "VBuddy Farm"
MsgBox "You are a great parent for keeping him alive so long.", vbOKOnly, "VBuddy Farm"
frmMain.Show
Unload Me
End If
End Sub

Sub UpdateStats()
lblHP.Caption = Health
lblNrg.Caption = Energy
lblHapy.Caption = Happy
lblMoney.Caption = Money
lblAge.Caption = Age
End Sub

Private Sub Form_Load()
Age = 1
Happy = 100
Energy = 100
Health = 100
TimeRedo = 0
TimeAge = 0
Money = 500
End Sub

Private Sub mfeed_Click()
If Money >= 20 Then
Money = Money - 20
If Health <= 99 Then
Health = Health + 1
Else
Health = 100
End If
If Energy <= 96 Then
Energy = Energy + 4
Else
Energy = 100
End If
If Happy <= 95 Then
Happy = Happy + 5
Else
Happy = 100
End If
Else
MsgBox "You dont have enough money to buy food!", vbOKOnly, "Not enough"
End If
UpdateStats
End Sub

Private Sub mnew_Click()
frmMain.Show
Unload Me
End Sub

Private Sub mplay_Click()
If Energy <= 4 Then
Energy = 0
Else
Energy = Energy - 5
End If
If Happy > 90 Then
Happy = 100
Else
Happy = Happy + 10
End If
UpdateStats
DeathCheck
End Sub

Private Sub mquit_Click()
End
End Sub

Private Sub mstatus_Click()
framVBS.Top = 240
framVBS.Left = 360
framVBS.Visible = True
End Sub

Private Sub mwork_Click()
If Happy <= 3 Then
MsgBox "I dont want to work..", vbOKOnly, "Bored"
MsgBox "Im bored...", vbOKOnly, "Bored"
Else
Money = Money + 40
If Energy <= 3 Then
Energy = 0
If Health <= 2 Then
Health = 0
Else
Health = Health - 2
End If
Else
Energy = Energy - 3
End If
If Happy <= 6 Then
Happy = 0
Else
Happy = Happy - 6
End If
End If
UpdateStats
DeathCheck
End Sub


Private Sub picVBuddy_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
If picVBuddy.Top + 250 >= 3000 Then
picVBuddy.Top = 2660
Else
If Happy < 97 Then
Happy = Happy + 3
Else
Happy = 100
End If
If Energy <= 0 Then
Energy = 0
Else
Energy = Energy - 1
End If
End If
picVBuddy.Top = picVBuddy.Top + 250
End If
If KeyCode = 39 Then
If picVBuddy.Left + 650 >= frmFarm.Width Then
picVBuddy.Left = frmFarm.Width - 770
Else
If Happy < 97 Then
Happy = Happy + 3
Else
Happy = 100
End If
If Energy <= 0 Then
Energy = 0
Else
Energy = Energy - 1
End If
End If
picVBuddy.Left = picVBuddy.Left + 250
End If
If KeyCode = 38 Then
If picVBuddy.Top - 250 <= 0 Then
picVBuddy.Top = 170
Else
If Happy < 97 Then
Happy = Happy + 3
Else
Happy = 100
End If
If Energy <= 0 Then
Energy = 0
Else
Energy = Energy - 1
End If
End If
picVBuddy.Top = picVBuddy.Top - 250
End If
If KeyCode = 37 Then
If picVBuddy.Left - 250 <= 0 Then
picVBuddy.Left = 170
Else
If Happy < 97 Then
Happy = Happy + 4
Else
Happy = 100
End If
If Energy <= 0 Then
Energy = 0
Else
Energy = Energy - 1
End If
End If
picVBuddy.Left = picVBuddy.Left - 250
End If
UpdateStats
DeathCheck
End Sub

Private Sub Timer1_Timer(Index As Integer)
If TimeAge = 30 Then
Age = Age + 1
TimeAge = 0
Else
TimeAge = TimeAge + 1
End If
If Happy <= 0 Then
Happy = 0
Else
Happy = Happy - 1
End If
If TimeRedo = 2 And Energy <= 60 Then
If Health <= 0 Then
Health = 0
Else
Health = Health - 1
End If
End If
If TimeRedo = 3 Then
If Energy = 0 Then
Energy = 0
Else
Energy = Energy - 1
End If
TimeRedo = 0
Else
TimeRedo = TimeRedo + 1
End If
UpdateStats
DeathCheck
End Sub
