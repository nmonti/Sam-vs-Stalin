VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "SAM vs. STALIN "
   ClientHeight    =   8895
   ClientLeft      =   2820
   ClientTop       =   1620
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   9645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd2Player 
      Caption         =   "Two Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5880
      TabIndex        =   7
      Top             =   7920
      Width           =   2655
   End
   Begin VB.CommandButton cmd1Player 
      Caption         =   "Single Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   5
      Top             =   7920
      Width           =   2655
   End
   Begin VB.Label lblVs 
      Caption         =   "  SAM              VS.          STALIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   6
      Top             =   3360
      Width           =   4935
   End
   Begin VB.Image imgCommies 
      Height          =   1815
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2295
   End
   Begin VB.Image imgStalin 
      Height          =   1455
      Left            =   7320
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Image imgSam 
      Height          =   1335
      Left            =   600
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lblPalin 
      Alignment       =   2  'Center
      Caption         =   $"frmMenu.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   2400
      Width           =   8655
   End
   Begin VB.Label lblI3 
      Caption         =   "3. Stalin shoots back, and if you're hit four times, America is doomed."
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Top             =   7080
      Width           =   3615
   End
   Begin VB.Label lblI2 
      Caption         =   $"frmMenu.frx":00B1
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   6240
      Width           =   3495
   End
   Begin VB.Label lblI1 
      Caption         =   "1. You control Uncle Sam with the arrow keys, and shoot the American flag with the Space bar."
      Height          =   735
      Left            =   3000
      TabIndex        =   1
      Top             =   5400
      Width           =   3375
   End
   Begin VB.Label lblInstructions 
      Caption         =   "INSTRUCTIONS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   4560
      Width           =   3255
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Nick Monti
'Comp sci Game Project
'Main objective is to defend Stalin's invasion of Alaska by
'shooting the american flag and controlling Uncle Sam.
'Infecting Stalin with democracy and hitting him four times will end
'the game and you win. If you are hit four times you lose.
'Two player is same way, but controls are different. Arrows now control
'Stalin and 0 on Num pad shoots, Sam is controlled with WASD and shoots
'with caps lock.


Private Sub cmd1Player_Click()
'Msgbox displays controls for 1 player
MsgBox "Uncle Sam is controlled with the ARROWS and shoots with SPACE BAR", vbCritical

'Unloads menu and shows frm1Player
Unload frmMenu
frm1Player.Show
End Sub

Private Sub cmd2Player_Click()
'Msgbox displays controls for 2 player
MsgBox "Uncle Sam is controled with WASD and shoots with CAPS LOCK. " & Chr(9) & Chr(10) & "Stalin is controlled with the ARROWS and shoots with NUM PAD 0.", vbCritical, "2 Player Controls"

'Unloads menu and shows frm2Player
Unload frmMenu
frm2Player.Show
End Sub

Private Sub Form_Load()
'Loads all pictures
imgSam.Picture = LoadPicture(App.Path & "\sam.gif")
imgStalin.Picture = LoadPicture(App.Path & "\stalin.gif")
imgCommies.Picture = LoadPicture(App.Path & "\commies.jpg")
End Sub

