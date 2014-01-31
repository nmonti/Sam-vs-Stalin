VERSION 5.00
Begin VB.Form frm2Player 
   Caption         =   "Sam vs. Stalin 2 Player"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrStalin 
      Interval        =   250
      Left            =   8040
      Top             =   6360
   End
   Begin VB.Timer tmrSamUp 
      Interval        =   1
      Left            =   1200
      Top             =   6120
   End
   Begin VB.Timer tmrSamLeft 
      Interval        =   1
      Left            =   720
      Top             =   6600
   End
   Begin VB.Timer tmrSamDown 
      Interval        =   1
      Left            =   1200
      Top             =   7080
   End
   Begin VB.Timer tmrSamRight 
      Interval        =   1
      Left            =   1680
      Top             =   6600
   End
   Begin VB.Timer tmrStalinUp 
      Interval        =   20
      Left            =   8040
      Top             =   5880
   End
   Begin VB.Timer tmrStalinLeft 
      Interval        =   20
      Left            =   7560
      Top             =   6360
   End
   Begin VB.Timer tmrStalinDown 
      Interval        =   20
      Left            =   8040
      Top             =   6840
   End
   Begin VB.Timer tmrStalinRight 
      Interval        =   20
      Left            =   8520
      Top             =   6360
   End
   Begin VB.Timer tmrUSA 
      Interval        =   1
      Left            =   1560
      Top             =   1800
   End
   Begin VB.Timer tmrUSSR 
      Interval        =   1
      Left            =   7560
      Top             =   1320
   End
   Begin VB.Image imgStalin 
      Height          =   1215
      Left            =   7320
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Image imgSam 
      Height          =   1215
      Left            =   360
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Image imgUSA 
      Height          =   495
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image imgUSSR 
      Height          =   495
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   735
   End
   Begin VB.Line LinTop 
      X1              =   0
      X2              =   9360
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line LinMid 
      X1              =   4320
      X2              =   4320
      Y1              =   0
      Y2              =   360
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
      TabIndex        =   1
      Top             =   0
      Width           =   3135
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
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frm2Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'StalinMove creates Stalin's movement, SamHP and StalinHP generate health
Private intStalinMove, intSamHP, intStalinHP As Integer
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'Stalin's movement is used with timers, on keydown a timer is enabled
If KeyCode = vbKeyRight Then
    tmrStalinRight.Enabled = True
ElseIf KeyCode = vbKeyLeft Then
    tmrStalinLeft.Enabled = True
ElseIf KeyCode = vbKeyUp Then
    tmrStalinUp.Enabled = True
ElseIf KeyCode = vbKeyDown Then
    tmrStalinDown.Enabled = True
End If

'Sam's movement is used with timer, on keydown a timer is enabled
If KeyCode = vbKeyD Then
    tmrSamRight.Enabled = True
ElseIf KeyCode = vbKeyA Then
    tmrSamLeft.Enabled = True
ElseIf KeyCode = vbKeyW Then
    tmrSamUp.Enabled = True
ElseIf KeyCode = vbKeyS Then
    tmrSamDown.Enabled = True
End If


'When NumPad 0 is hit, USSR flag is shot
'Stalin cannot shoot until flag hits Sam or edge of form
If tmrUSSR.Enabled = False Then
    If KeyCode = vbKeyNumpad0 Then
        imgUSSR.Picture = LoadPicture(App.Path & "\ussr.gif")
        imgUSSR.Left = imgStalin.Left
        imgUSSR.Top = imgStalin.Top
        imgUSSR.Visible = True
        tmrUSSR.Enabled = True
    End If
ElseIf tmrUSSR.Enabled = True Then
    
End If

'When Caps Lock is hit, USA flag is shot
'Sam cannot shoot until flag hits Stalin or edge of form
If tmrUSA.Enabled = False Then
    If KeyCode = vbKeyCapital Then
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


'When you release the direction it diables timer
If KeyCode = vbKeyRight Then
    tmrStalinRight.Enabled = False
ElseIf KeyCode = vbKeyLeft Then
    tmrStalinLeft.Enabled = False
ElseIf KeyCode = vbKeyUp Then
    tmrStalinUp.Enabled = False
ElseIf KeyCode = vbKeyDown Then
    tmrStalinDown.Enabled = False
End If

'When you release the direction it disables timer
If KeyCode = vbKeyD Then
    tmrSamRight.Enabled = False
ElseIf KeyCode = vbKeyA Then
    tmrSamLeft.Enabled = False
ElseIf KeyCode = vbKeyW Then
    tmrSamUp.Enabled = False
ElseIf KeyCode = vbKeyS Then
    tmrSamDown.Enabled = False
End If


End Sub

Private Sub Form_Load()
'randomizes Stalin's movement
Randomize



'Loads frm2Player's Alaskan background
frm2Player.Picture = LoadPicture(App.Path & "\alaska.jpg")


'Disbles tmrUSA on form load
tmrUSA.Enabled = False

'Disables tmrUSSR on form load
tmrUSSR.Enabled = False

'Sets HP to 100 at form load
intSamHP = 100
intStalinHP = 100

'LinTop makes you not able to go above the HP label
LinTop.Y1 = frm2Player.ScaleTop + 300
LinTop.Y2 = frm2Player.ScaleTop + 300
LinTop.X1 = frm2Player.ScaleLeft
LinTop.X2 = frm2Player.ScaleLeft + frm2Player.ScaleWidth

'Linmid separates LinTop
LinMid.Y1 = frm2Player.ScaleTop
LinMid.Y2 = LinTop.Y1
LinMid.X1 = frm2Player.ScaleWidth / 2
LinMid.X2 = frm2Player.ScaleWidth / 2



'loads Stalin
imgStalin.Picture = LoadPicture(App.Path & "\Stalin.gif")

'loads Uncle Sam
imgSam.Picture = LoadPicture(App.Path & "\sam.gif")

'Sets position of lblSamHP
lblSamHP.Caption = "Uncle Sam's HP: " & intSamHP
lblSamHP.Top = frm2Player.ScaleTop
lblSamHP.Height = frm2Player.ScaleTop + LinTop.Y1

'Sets position of lblStalinHP
lblStalinHP.Caption = "Stalin's HP: " & intStalinHP
lblStalinHP.Top = frm2Player.ScaleTop
lblStalinHP.Height = frm2Player.ScaleTop + LinTop.Y1


'disables Uncle Sam direction timers on form load
tmrSamUp.Enabled = False
tmrSamDown.Enabled = False
tmrSamLeft.Enabled = False
tmrSamRight.Enabled = False

'disables Stalin movement timers on form load
tmrStalinUp.Enabled = False
tmrStalinDown.Enabled = False
tmrStalinLeft.Enabled = False
tmrStalinRight.Enabled = False



End Sub



Private Sub tmrSamDown_Timer()
'Sam moves down
imgSam.Top = imgSam.Top + 50

'Sam won't go off bottom of screen
If imgSam.Top + imgSam.Height > frm2Player.ScaleHeight Then
     imgSam.Top = frm2Player.ScaleHeight - imgSam.Height
End If
End Sub

Private Sub tmrSamLeft_Timer()
'Sam moves left
imgSam.Left = imgSam.Left - 50

'Sam won't go off left of screen
If imgSam.Left < frm2Player.ScaleLeft Then
    imgSam.Left = frm2Player.ScaleLeft
End If

End Sub

Private Sub tmrSamRight_Timer()
'Sam moves right
imgSam.Left = imgSam.Left + 50

'Sam won't go past middle of frm2player
If imgSam.Left + imgSam.Width > frm2Player.ScaleWidth / 2 Then
     imgSam.Left = frm2Player.ScaleWidth / 2 - imgSam.Width
End If

End Sub

Private Sub tmrSamUp_Timer()
'Sam moves up
imgSam.Top = imgSam.Top - 50

'Sam won't go off top of screen
If imgSam.Top < LinTop.Y1 Then
    imgSam.Top = LinTop.Y1
End If

End Sub

Private Sub tmrStalinDown_Timer()
'Stalin moves down
imgStalin.Top = imgStalin.Top + 100

'Stalin won't go off bottom of screen
If imgStalin.Top + imgStalin.Height > frm2Player.ScaleHeight Then
     imgStalin.Top = frm2Player.ScaleHeight - imgStalin.Height
End If
End Sub

Private Sub tmrStalinLeft_Timer()
'Stalin moves left
imgStalin.Left = imgStalin.Left - 100

'Stalin won't go past left of screen
If imgStalin.Left < frm2Player.ScaleWidth / 2 Then
     imgStalin.Left = frm2Player.ScaleWidth / 2
     tmrStalinLeft.Enabled = False
End If

End Sub

Private Sub tmrStalinRight_Timer()
'Stalin moves right
imgStalin.Left = imgStalin.Left + 100

'Stalin won't go off right of screen
If imgStalin.Left + imgStalin.Width > frm2Player.ScaleWidth Then
    imgStalin.Left = frm2Player.ScaleWidth - imgStalin.Width
    tmrStalinRight.Enabled = False
End If

End Sub

Private Sub tmrStalinShoot_Timer()
'Enabled USSR flag's movement
tmrUSSR.Enabled = True

End Sub

Private Sub tmrStalinUp_Timer()
'Stalin moves up
imgStalin.Top = imgStalin.Top - 100

'Stalin won't go above LinTop
If imgStalin.Top < LinTop.Y1 Then
    imgStalin.Top = LinTop.Y1
    tmrStalinUp.Enabled = False
End If
End Sub

Private Sub tmrUSA_Timer()
'USA flag moves right
imgUSA.Left = imgUSA.Left + 100

'If flag hits Stalin, tmrUSA disabled and 25 HP taken from Stalin
If imgUSA.Left + imgUSA.Width > imgStalin.Left And imgUSA.Top > imgStalin.Top - 100 And imgUSA.Top + imgUSA.Height < imgStalin.Top + imgStalin.Height + 300 Then
    imgUSA.Visible = False
    tmrUSA.Enabled = False
    intStalinHP = intStalinHP - 25
    lblStalinHP.Caption = "Stalin's HP: " & intStalinHP
'If flag goes off right of frm2player, tmrUSA disabled
ElseIf imgUSA.Left > frm2Player.ScaleLeft + frm2Player.ScaleWidth Then
    tmrUSA.Enabled = False
    imgUSA.Visible = False
End If

'If intStalinHP = 0, all timers disabled, loads frmVictory
If intStalinHP = 0 Then
    tmrUSA.Enabled = False
    tmrUSSR.Enabled = False
    tmrSamUp.Enabled = False
    tmrSamDown.Enabled = False
    tmrSamLeft.Enabled = False
    tmrSamRight.Enabled = False
    Unload Me
    frmVictory.Show
   
End If
End Sub

Private Sub tmrUSSR_Timer()
'USA flag moves right
imgUSSR.Left = imgUSSR.Left - 100

'If flag hits sam, tmrUSSR disabled and 25 HP taken from sam
If imgUSSR.Left < imgSam.Left + imgSam.Width And imgUSSR.Top > imgSam.Top - 100 And imgUSSR.Top + imgUSSR.Height < imgSam.Top + imgSam.Height + 300 Then
    imgUSSR.Visible = False
    tmrUSSR.Enabled = False
    intSamHP = intSamHP - 25
    lblSamHP.Caption = "Uncle Sam's HP: " & intSamHP
'If flag goes off right of frm2player, tmrussr disabled
ElseIf imgUSSR.Left + imgUSSR.Width < frm2Player.ScaleLeft Then
    tmrUSSR.Enabled = False
    imgUSSR.Visible = False
End If




'If intSamHP = 0, all timers diabled, loads frmLose
If intSamHP = 0 Then
    tmrUSA.Enabled = False
    tmrUSSR.Enabled = False
    tmrSamUp.Enabled = False
    tmrSamDown.Enabled = False
    tmrSamLeft.Enabled = False
    tmrSamRight.Enabled = False
    Unload Me
    frmLose.Show
   
End If
End Sub

