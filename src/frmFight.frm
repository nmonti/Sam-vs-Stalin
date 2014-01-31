VERSION 5.00
Begin VB.Form frm1Player 
   Caption         =   "Sam vs. Stalin"
   ClientHeight    =   7800
   ClientLeft      =   3015
   ClientTop       =   1815
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrUSSR 
      Interval        =   1
      Left            =   7560
      Top             =   1320
   End
   Begin VB.Timer tmrStalinShoot 
      Interval        =   500
      Left            =   7560
      Top             =   1920
   End
   Begin VB.Timer tmrUSA 
      Interval        =   1
      Left            =   1560
      Top             =   1800
   End
   Begin VB.Timer tmrStalinRight 
      Interval        =   20
      Left            =   8520
      Top             =   6360
   End
   Begin VB.Timer tmrStalinDown 
      Interval        =   20
      Left            =   8040
      Top             =   6840
   End
   Begin VB.Timer tmrStalinLeft 
      Interval        =   20
      Left            =   7560
      Top             =   6360
   End
   Begin VB.Timer tmrStalinUp 
      Interval        =   20
      Left            =   8040
      Top             =   5880
   End
   Begin VB.Timer tmrSamRight 
      Interval        =   1
      Left            =   1680
      Top             =   6600
   End
   Begin VB.Timer tmrSamDown 
      Interval        =   1
      Left            =   1200
      Top             =   7080
   End
   Begin VB.Timer tmrSamLeft 
      Interval        =   1
      Left            =   720
      Top             =   6600
   End
   Begin VB.Timer tmrSamUp 
      Interval        =   1
      Left            =   1200
      Top             =   6120
   End
   Begin VB.Timer tmrStalin 
      Interval        =   250
      Left            =   8040
      Top             =   6360
   End
   Begin VB.Label lblStalinHP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label lblSamHP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
   Begin VB.Line LinMid 
      X1              =   4320
      X2              =   4320
      Y1              =   0
      Y2              =   360
   End
   Begin VB.Line LinTop 
      X1              =   0
      X2              =   9360
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Image imgUSSR 
      Height          =   495
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   735
   End
   Begin VB.Image imgUSA 
      Height          =   495
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image imgSam 
      Height          =   1215
      Left            =   360
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Image imgStalin 
      Height          =   1215
      Left            =   7320
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1455
   End
End
Attribute VB_Name = "frm1Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private intStalinMove, intSamHP, intStalinHP As Integer
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'movement is used with timers, on keydown a timer is enabled
If KeyCode = vbKeyRight Then
    tmrSamRight.Enabled = True
ElseIf KeyCode = vbKeyLeft Then
    tmrSamLeft.Enabled = True
ElseIf KeyCode = vbKeyUp Then
    tmrSamUp.Enabled = True
ElseIf KeyCode = vbKeyDown Then
    tmrSamDown.Enabled = True
End If


'When space bar is hit, Flag shoots and if flag is in air you cant hit space
If tmrUSA.Enabled = False Then
    If KeyCode = vbKeySpace Then
        imgUSA.Picture = LoadPicture(App.Path & "\usa.gif")
        imgUSA.Left = imgSam.Left + imgSam.Width
        imgUSA.Top = imgSam.Top
        imgUSA.Visible = True
        tmrUSA.Enabled = True
    End If
ElseIf tmrUSA.Enabled = True Then
    
End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)


'when you release the direction it diables timer

If KeyCode = vbKeyRight Then
    tmrSamRight.Enabled = False
ElseIf KeyCode = vbKeyLeft Then
    tmrSamLeft.Enabled = False
ElseIf KeyCode = vbKeyUp Then
    tmrSamUp.Enabled = False
ElseIf KeyCode = vbKeyDown Then
    tmrSamDown.Enabled = False
End If


End Sub

Private Sub Form_Load()
'randomizes Stalin's movement
Randomize

'Loads BOOM sound
'MMCBoom.FileName = App.Path & "\boom.wav"
'MMCBoom.DeviceType = "waveaudio"
'MMCBoom.Command = "open"


'Disbles tmrUSA on form load
tmrUSA.Enabled = False

'Sets HP to 100 at form load
intSamHP = 100
intStalinHP = 100

'LinTop is 400 twips from top, fits perfectly left to right
LinTop.Y1 = frm1Player.ScaleTop + 300
LinTop.Y2 = frm1Player.ScaleTop + 300
LinTop.X1 = frm1Player.ScaleLeft
LinTop.X2 = frm1Player.ScaleLeft + frm1Player.ScaleWidth

'Linmid is in middle of frm1player, starts at top of frm1player, ends at LinTop
LinMid.Y1 = frm1Player.ScaleTop
LinMid.Y2 = LinTop.Y1
LinMid.X1 = frm1Player.ScaleWidth / 2
LinMid.X2 = frm1Player.ScaleWidth / 2


'enabled Stalin's shoot timer
tmrStalinShoot.Enabled = True

'Stalin shoots
imgUSSR.Picture = LoadPicture(App.Path & "\ussr.gif")
imgUSSR.Left = imgStalin.Left
imgUSSR.Top = imgStalin.Top
imgUSSR.Visible = True


'loads Stalin
imgStalin.Picture = LoadPicture(App.Path & "\Stalin.gif")

'loads Uncle Sam
imgSam.Picture = LoadPicture(App.Path & "\sam.gif")

'Sets position of lblSamHP
lblSamHP.Caption = "Uncle Sam's HP: " & intSamHP
lblSamHP.Top = frm1Player.ScaleTop
lblSamHP.Height = frm1Player.ScaleTop + LinTop.Y1

'Sets position of lblStalinHP
lblStalinHP.Caption = "Stalin's HP: " & intStalinHP
lblStalinHP.Top = frm1Player.ScaleTop
lblStalinHP.Height = frm1Player.ScaleTop + LinTop.Y1


'disables Uncle Sam timers on form load
tmrSamUp.Enabled = False
tmrSamDown.Enabled = False
tmrSamLeft.Enabled = False
tmrSamRight.Enabled = False

'disables Stalin timers on form load
tmrStalinUp.Enabled = False
tmrStalinDown.Enabled = False
tmrStalinLeft.Enabled = False
tmrStalinRight.Enabled = False



End Sub



Private Sub tmrSamDown_Timer()
imgSam.Top = imgSam.Top + 50 'when down is pressed timer starts

'Sam won't go off bottom of screen
If imgSam.Top + imgSam.Height > frm1Player.ScaleHeight Then
     imgSam.Top = frm1Player.ScaleHeight - imgSam.Height
End If
End Sub

Private Sub tmrSamLeft_Timer()
imgSam.Left = imgSam.Left - 50 'when left is pressed timer starts

'Sam won't go off left of screen
If imgSam.Left < frm1Player.ScaleLeft Then
    imgSam.Left = frm1Player.ScaleLeft
End If

End Sub

Private Sub tmrSamRight_Timer()
imgSam.Left = imgSam.Left + 50 'when right is pressed timer starts

'Sam won't go past middle of frm1player
If imgSam.Left + imgSam.Width > frm1Player.ScaleWidth / 2 Then
     imgSam.Left = frm1Player.ScaleWidth / 2 - imgSam.Width
End If

End Sub

Private Sub tmrStalin_Timer()
'Stalin's movement number
intStalinMove = Int(Rnd * 4 + 1)

If intStalinMove = 1 Then 'Stalin moves up
    tmrStalinUp.Enabled = True
    tmrStalinDown.Enabled = False
    tmrStalinLeft.Enabled = False
    tmrStalinRight.Enabled = False
ElseIf intStalinMove = 2 Then 'Stalin moves down
    tmrStalinDown.Enabled = True
    tmrStalinUp.Enabled = False
    tmrStalinLeft.Enabled = False
    tmrStalinRight.Enabled = False
ElseIf intStalinMove = 3 Then 'Stalin moves left
    tmrStalinLeft.Enabled = True
    tmrStalinDown.Enabled = False
    tmrStalinUp.Enabled = False
    tmrStalinRight.Enabled = False
ElseIf intStalinMove = 4 Then 'Stalin moves right
    tmrStalinRight.Enabled = True
    tmrStalinDown.Enabled = False
    tmrStalinLeft.Enabled = False
    tmrStalinUp.Enabled = False

End If

End Sub

Private Sub tmrSamUp_Timer()
imgSam.Top = imgSam.Top - 50 'when up is pressed timer starts

'Sam won't go off top of screen
If imgSam.Top < LinTop.Y1 Then
    imgSam.Top = LinTop.Y1
End If

End Sub

Private Sub tmrStalinDown_Timer()
imgStalin.Top = imgStalin.Top + 100 'Stalin moves down

'Stalin won't go off bottom of screen
If imgStalin.Top + imgStalin.Height > frm1Player.ScaleHeight Then
     imgStalin.Top = frm1Player.ScaleHeight - imgStalin.Height
End If
End Sub

Private Sub tmrStalinLeft_Timer()
imgStalin.Left = imgStalin.Left - 100 'Stalin moves left

'Stalin won't go past left of screen
If imgStalin.Left < frm1Player.ScaleWidth / 2 Then
     imgStalin.Left = frm1Player.ScaleWidth / 2
     tmrStalinLeft.Enabled = False
End If

End Sub

Private Sub tmrStalinRight_Timer()
imgStalin.Left = imgStalin.Left + 100 'Stalin moves right

'Stalin won't go off right of screen
If imgStalin.Left + imgStalin.Width > frm1Player.ScaleWidth Then
    imgStalin.Left = frm1Player.ScaleWidth - imgStalin.Width
    tmrStalinRight.Enabled = False
End If

End Sub

Private Sub tmrStalinShoot_Timer()
'Enabled USSR flag's movement
tmrUSSR.Enabled = True

End Sub

Private Sub tmrStalinUp_Timer()
imgStalin.Top = imgStalin.Top - 100 'Stalin moves up

'Stalin won't go off top
If imgStalin.Top < LinTop.Y1 Then
    imgStalin.Top = LinTop.Y1
    tmrStalinUp.Enabled = False
End If
End Sub

Private Sub tmrUSA_Timer()
imgUSA.Left = imgUSA.Left + 100 'USA flag moves right

'If flag hits Stalin, tmrUSA disabled and 25 HP taken from Stalin
If imgUSA.Left + imgUSA.Width > imgStalin.Left And imgUSA.Top > imgStalin.Top And imgUSA.Top + imgUSA.Height < imgStalin.Top + imgStalin.Height Then
    imgUSA.Visible = False
    tmrUSA.Enabled = False
    intStalinHP = intStalinHP - 25
    lblStalinHP.Caption = "Stalin's HP: " & intStalinHP
'If flag goes off right of frm1player, tmrUSA disabled
ElseIf imgUSA.Left > frm1Player.ScaleLeft + frm1Player.ScaleWidth Then
    tmrUSA.Enabled = False
    imgUSA.Visible = False
End If

'If intStalinHP = 0, all timers disabled, loads frmVictory
If intStalinHP = 0 Then
    tmrUSA.Enabled = False
    tmrUSSR.Enabled = False
    tmrStalinShoot.Enabled = False
    tmrStalin.Enabled = False
    tmrSamUp.Enabled = False
    tmrSamDown.Enabled = False
    tmrSamLeft.Enabled = False
    tmrSamRight.Enabled = False
    Unload Me
    frmVictory.Show
   
End If
End Sub

Private Sub tmrUSSR_Timer()
imgUSSR.Left = imgUSSR.Left - 100 'USSR flag moves left



'If USSR flag hits Sam, 25 HP taken away from Sam, flag resets at Stalin
If imgUSSR.Left < imgSam.Left + imgSam.Width And imgUSSR.Top > imgSam.Top And imgUSSR.Top + imgUSSR.Height < imgSam.Top + imgSam.Height Then
    imgUSSR.Left = imgStalin.Left
    imgUSSR.Top = imgStalin.Top
    imgUSSR.Visible = True
    intSamHP = intSamHP - 25
    lblSamHP.Caption = "Uncle Sam's HP: " & intSamHP
'If USSR flag goes off left of frm1player, flag resets at Stalin
ElseIf imgUSSR.Left < frm1Player.ScaleLeft Then
    imgUSSR.Left = imgStalin.Left
    imgUSSR.Top = imgStalin.Top
    imgUSSR.Visible = True
End If


If intSamHP = 50 Then
   'MMCBoom.Command = "play"
   'MMCBoom.Command = "prev"
'If intSamHP = 0, all timers diabled, loads frmLose
ElseIf intSamHP = 0 Then
    tmrUSA.Enabled = False
    tmrUSSR.Enabled = False
    tmrStalinShoot.Enabled = False
    tmrStalin.Enabled = False
    tmrSamUp.Enabled = False
    tmrSamDown.Enabled = False
    tmrSamLeft.Enabled = False
    tmrSamRight.Enabled = False
    
    Unload Me
    frmLose.Show
   
End If
End Sub
