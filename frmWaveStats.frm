VERSION 5.00
Begin VB.Form frmWaveStats 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Brian's 911 Version 2    Author: Mr. Brian A. Rouse    BRouse@AAANP.ORG"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   ControlBox      =   0   'False
   Icon            =   "frmWaveStats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmWaveStats.frx":0442
   ScaleHeight     =   5745
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStats 
      BackColor       =   &H00C00000&
      Caption         =   "Level Statistics:"
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   1800
      TabIndex        =   4
      Top             =   1080
      Width           =   5415
      Begin VB.Label lblBonus 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         Caption         =   "()"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   3690
         TabIndex        =   22
         Top             =   1560
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblBonus 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         Caption         =   "()"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   3810
         TabIndex        =   21
         Top             =   1920
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lblBonus 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         Caption         =   "()"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   3810
         TabIndex        =   18
         Top             =   450
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         Caption         =   "Patriot Missiles Launched"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   2670
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         Caption         =   "Anti-Ballistic Missiles ABMs:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1680
         TabIndex        =   14
         Top             =   1200
         Width           =   2940
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         Caption         =   "Success Ratio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1080
         TabIndex        =   13
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         Caption         =   "EMPs Fired:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1440
         TabIndex        =   12
         Top             =   1920
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         Caption         =   "Trade Centers Saved"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   2235
      End
      Begin VB.Label lblTarRem 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   240
         Left            =   2700
         TabIndex        =   10
         Top             =   450
         Width           =   75
      End
      Begin VB.Label lblABMs 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   3240
         TabIndex        =   9
         Top             =   840
         Width           =   75
      End
      Begin VB.Label lblHit 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   4680
         TabIndex        =   8
         Top             =   1200
         Width           =   195
      End
      Begin VB.Label lblRatio 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2700
         TabIndex        =   7
         Top             =   1560
         Width           =   75
      End
      Begin VB.Label lblEMPs 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2700
         TabIndex        =   6
         Top             =   1920
         Width           =   75
      End
      Begin VB.Label lblPct 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2880
         TabIndex        =   5
         Top             =   1560
         Visible         =   0   'False
         Width           =   210
      End
   End
   Begin VB.Timer tmrStatsAnim 
      Interval        =   60
      Left            =   6270
      Top             =   2070
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Retreat"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Clicking this yellow retreat button depicts the cowardice in your heart and quits!"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H000000FF&
      Caption         =   "&Continue..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   600
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Click this button to continue like the red blooded American you are!"
      Top             =   2280
      Width           =   1365
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   1560
      X2              =   5640
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   720
      X2              =   10320
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblBonusScore 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3720
      TabIndex        =   20
      Top             =   3990
      Width           =   180
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Bonus:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2880
      TabIndex        =   19
      Top             =   3990
      Width           =   720
   End
   Begin VB.Label lblTotalScore 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3690
      TabIndex        =   17
      Top             =   4320
      Width           =   75
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Total Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2310
      TabIndex        =   16
      Top             =   4320
      Width           =   1290
   End
   Begin VB.Label lblLevel 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   1800
      TabIndex        =   3
      Top             =   720
      Width           =   180
   End
   Begin VB.Label lblSuccess 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Missile Defense Data.USA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   5985
   End
End
Attribute VB_Name = "frmWaveStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iHitRatio As Integer

Private Sub cmdContinue_Click()

'clear totals etc.
iNumFired = 0
iNumSucceed = 0
iEMPsFired = 0
frmGame.lblScore.Caption = iScore

'proceed to next level
If iLevel < 4 Then
    iLevel = iLevel + 1
Else
    End
End If

'set up level parameters
Select Case iLevel
    Case 2
    'level 2
        iICBMDiam = 1000
        iVertSpeed = 280
        iNumAllowed = 50
        frmGame.lblICBMsLeft.Caption = iNumAllowed
        iEMPsAllowed = 0
        MsgBox "Congratulations, You have just been promoted to Captain, Foreign Intelligence has been infiltrated by U.S. Intelligence...future attacks on Trade Centers have been verified...Missile Control Launch has been Authorized and Authenticated... Good Luck Captain!", vbExclamation, "Bin Laden Stage 2 Brian's 911 Version 2"
    Case 3
    'level 3
        bHoming = True
        iICBMDiam = 1000
        iVertSpeed = 180
        iNumAllowed = 50
        frmGame.lblICBMsLeft.Caption = iNumAllowed
        iEMPsAllowed = 5
        MsgBox "Right Click launches Nuclear War Head" & vbCrLf & "You only get 5!", vbInformation, "Brian's 911 Attacks Level-3"
 Case 4
    'level 4
        bHoming = True
        iICBMDiam = 1000
        iVertSpeed = 380
        iNumAllowed = 75
        frmGame.lblICBMsLeft.Caption = iNumAllowed
        iEMPsAllowed = 10
        MsgBox "Commander...you now have 10 Nuclear Missiles Right Click/ 75 ABM's Left Click" & vbCrLf & "Last level It's all up to you ...this is not an exercise!", vbExclamation, "Brian's 911 Attacks Bonus Level-4!"
End Select

bGameOn = True
bICBMOK = True
bNukeOK = True
frmGame.tmrMissile1Anim.Enabled = True
frmGame.tmrMissile2Anim.Enabled = True
frmGame.tmrMissile3Anim.Enabled = True
frmGame.tmrWaveTime.Enabled = True
Unload Me

End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub Form_Load()

'set positions
With Me
    '.Width = Screen.Width / 2
    '.Height = Screen.Height / 2
End With
With lblSuccess
    .Left = (Me.Width / 2) - (.Width / 2)
    .Top = 100
End With
With cmdContinue
    .Left = (Me.Width / 2) - (.Width / 2)
    .Top = Me.Height - (.Height + 200)
End With
With cmdExit
    .Left = Me.Width - (.Width + 200)
    .Top = Me.Height - (.Height + 200)
End With

'set caption
lblLevel.Caption = "Red Alert " & iLevel

If (iNumSucceed > 0) And (iNumFired > 0) Then
    iHitRatio = Fix((iNumSucceed / iNumFired) * 100)
End If

End Sub


Private Sub tmrStatsAnim_Timer()
Static iBonus As Integer
Dim iTemp As Integer
Dim rc As Long

'bang!
If bSounds Then
    rc = sndPlaySound(App.Path & "\explosion.wav", 1)
End If

'loop through each totals incrementing then onto next
If Val(lblTarRem.Caption) < (iNumTargets - iNumHit) Then
    lblTarRem.Caption = Val(lblTarRem.Caption) + 1
    Exit Sub
End If
If lblBonus(0).Visible = False Then
    iTemp = ((iNumTargets - iNumHit) * 50)
    lblBonus(0).Caption = "(+" & iTemp & ")"
    lblBonus(0).Visible = True
    iBonus = iBonus + iTemp
    iTemp = 0
End If

If Val(lblABMs.Caption) < iNumFired Then
    lblABMs.Caption = Val(lblABMs.Caption) + 1
    Exit Sub
End If

If Val(lblHit.Caption) < iNumSucceed Then
    lblHit.Caption = Val(lblHit.Caption) + 1
    Exit Sub
End If


lblRatio.Caption = iHitRatio
lblPct.Visible = True
If lblBonus(1).Visible = False Then
    Select Case iHitRatio
        Case 75 To 90
            iTemp = iHitRatio * 1
        Case 90 To 100
            iTemp = iHitRatio * 3
        Case Is > 100
            iTemp = iHitRatio * 5
    End Select
    lblBonus(1).Caption = "(+" & iTemp & ")"
    lblBonus(1).Visible = True
    iBonus = iBonus + iTemp
    iTemp = 0
End If

If Val(lblEMPs.Caption) < iEMPsFired Then
    lblEMPs.Caption = Val(lblEMPs.Caption) + 1
    Exit Sub
End If

'accumulate bonus score
iScore = iScore + iBonus

lblBonusScore.Caption = iBonus
lblTotalScore.Caption = iScore
cmdContinue.Enabled = True
tmrStatsAnim.Enabled = False
iBonus = 0
End Sub
