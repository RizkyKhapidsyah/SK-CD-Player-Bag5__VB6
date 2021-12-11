VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alpha Player"
   ClientHeight    =   2430
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "About the Author!"
      Height          =   375
      Left            =   2040
      TabIndex        =   13
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Track Jump"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "Play"
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Stop"
         Height          =   375
         Left            =   2160
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   9
         Text            =   "1"
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Track:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame frmTime 
      Caption         =   "Time Jump"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4455
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Text            =   "5"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "<<"
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command7 
         Caption         =   ">>"
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Seconds"
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CD len"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "open tray"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "close tray"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   1920
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Snd As CDAudio

Private Sub Command1_Click()
    Snd.SeekCDtoX Val(Text1)
End Sub

Private Sub Command2_Click()
    s$ = Snd.GetCDLength
    MsgBox "Total length of CD: " & s, , "CD len"
End Sub

Private Sub Command3_Click()
    Snd.CloseCD
End Sub

Private Sub Command4_Click()
    Snd.EjectCD
End Sub

Private Sub Command5_Click()
    Snd.StopPlay
End Sub

Private Sub Command6_Click()
    Snd.ReWind Val(Text2) * 1000
End Sub

Private Sub Command7_Click()
    Snd.FastForward Val(Text2) * 1000
End Sub

Private Sub Command8_Click()
MsgBox "The Çrimson§hade" & Chr(10) & "dadams@chesco.com", vbInformation, "About the Author"
End Sub

Private Sub Form_Load()
    Set Snd = New CDAudio
    Snd.ReadyDevice
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Snd.StopPlay
    Snd.UnloadAll
End Sub
