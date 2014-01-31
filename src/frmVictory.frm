VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmVictory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VICTORY"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReplay 
      Caption         =   "Relive your greatest moment?"
      Height          =   615
      Left            =   1920
      TabIndex        =   4
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   3360
      TabIndex        =   3
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdSpeech 
      Caption         =   "Stalin's death speech"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   4560
      Width           =   1815
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMPSpeech 
      Height          =   4455
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   5775
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   10186
      _cy             =   7858
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "  CONGRATULATIONS!      You have defeated the      communist invasion    and ended the  evil Stalin's life."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmVictory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
'Unloads all active forms
Unload frmVictory
Unload frmMenu
Unload frm1Player
End Sub

Private Sub cmdReplay_Click()
'Reloads menu and unloads other forms
frmMenu.Show
Unload frmVictory
Unload frm1Player
End Sub

Private Sub cmdSpeech_Click()

'Makes WMPSpeech visible and plays Stalin's death speech
WMPSpeech.Visible = True
WMPSpeech.URL = App.Path & "\death speech.wmv"

'Disables cmdSpeech
cmdSpeech.Enabled = False

'Places cmdReplay above cmdSpeech
cmdReplay.Left = 600
cmdReplay.Top = 4560

'Hides cmdSpeech
cmdSpeech.Visible = False
End Sub
