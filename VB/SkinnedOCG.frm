VERSION 5.00
Begin VB.Form SkinnedOCG 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "OneClick Go!"
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5775
   Icon            =   "SkinnedOCG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5775
   Begin VB.Frame FrontImg 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   2400
      TabIndex        =   8
      Top             =   1830
      Width           =   1005
      Begin VB.Image PlayPauseImg 
         Height          =   1005
         Left            =   0
         Top             =   0
         Width           =   1005
      End
   End
   Begin VB.Timer NPSVanisher 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   120
   End
   Begin VB.Timer FadeFX 
      Interval        =   10
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer NPSfx 
      Interval        =   10
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer NPSVvanisher 
      Interval        =   5000
      Left            =   120
      Top             =   120
   End
   Begin VB.Frame NPSinfo 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   1140
      Width           =   5775
      Begin VB.Label SeekBack 
         BackStyle       =   0  'Transparent
         Height          =   450
         Left            =   630
         TabIndex        =   2
         Top             =   1695
         Width           =   3975
      End
      Begin VB.Label Rank 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   810
         TabIndex        =   1
         Top             =   150
         Width           =   120
      End
      Begin VB.Image NPSBlockerR 
         Height          =   525
         Left            =   5100
         Top             =   0
         Width           =   675
      End
      Begin VB.Image NPSBlockerL 
         Height          =   660
         Left            =   0
         Top             =   0
         Width           =   1515
      End
      Begin VB.Shape SeekFX 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         Height          =   45
         Left            =   615
         Shape           =   4  'Rounded Rectangle
         Top             =   1905
         Width           =   15
      End
      Begin VB.Label NextBtn 
         BackStyle       =   0  'Transparent
         Height          =   615
         Left            =   4425
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.Label PrevBtn 
         BackStyle       =   0  'Transparent
         Height          =   615
         Left            =   390
         TabIndex        =   6
         Top             =   1140
         Width           =   975
      End
      Begin VB.Image PrevNextImg 
         Height          =   990
         Left            =   0
         Top             =   690
         Width           =   5760
      End
      Begin VB.Label StopBtn 
         BackStyle       =   0  'Transparent
         Height          =   300
         Left            =   4800
         TabIndex        =   5
         Top             =   1770
         Width           =   600
      End
      Begin VB.Label SeekTime 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2115
         TabIndex        =   4
         Top             =   1740
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label NPS 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   3300
         TabIndex        =   3
         Top             =   150
         Width           =   120
      End
      Begin VB.Image StopImg 
         Height          =   2160
         Left            =   0
         Top             =   0
         Width           =   5760
      End
   End
End
Attribute VB_Name = "SkinnedOCG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright © 2011 ANIco.in
'Welcome to the source code of OneClick Go! Skin Mode 1
'The code in this module deals with the first skin mode of OCG.
'The modification and resdistribution of the code is completely permitted.
'---------------------------------------------------------------------------

Option Explicit

Private Sub Form_Load()
    Refresher
    SetWindowPos hWnd, -1, 0, 0, 0, 0, 1 Or 2
    SetWindowLongA hWnd, -20, GetWindowLongA(hWnd, -20) Or &H80000
    SetLayeredWindowAttributes hWnd, 0, 255 + 0&, &H2& + &H1 'Setting the SkinnedOCG as transparent to black color and always on top
    OpacityNow = 255
    Left = GtSt("SkinLeft", Screen.Width - Width)
    Top = GtSt("SkinTop", Screen.Height - Height)
    SeekFX.BackColor = RGB(255, 173, 80)
    Rank.ForeColor = RGB(1, 0, 0)
    PrevNextImg.Picture = LoadPicture(OfName("PrevNextDefault"))
    StopImg.Picture = LoadPicture(OfName("StopDefault"))
    NPSBlockerR.Picture = LoadPicture(OfName("NPSBlockerR"))
    NPSBlockerL.Picture = LoadPicture(OfName("NPSBlockerL"))
    If OneClickGo.mnuPlay.Caption = "&Pause" Then
        PlayPauseImg.Picture = LoadPicture(OfName("PlayPauseADefault"))
    Else
        PlayPauseImg.Picture = LoadPicture(OfName("PlayPauseLDefault"))
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then 'Disable Skin Mode
        Unload Me
        OneClickGo.mnuSkinMode.Checked = False
        OneClickGo.Show
    ElseIf KeyCode = vbKeyF3 Then 'Cycle through skins
        SkinnedOCG2.Show
        OneClickGo.mnuFreezedBlue.Checked = True
        OneClickGo.mnuRockstarGold.Checked = False
        SvSt "CurrentSkin", "1"
        Unload SkinnedOCG
    ElseIf KeyCode = vbKeyF4 Then 'Close on Alt+F4
        If Shift = 4 Then
            Unload Me
            Unload OneClickGo
        End If
    ElseIf KeyCode = vbKeyF7 Then 'Volume Down
        OneClickGo.mnuVolumeDown_Click
    ElseIf KeyCode = vbKeyF8 Then 'Volume Up
        OneClickGo.mnuVolumeUp_Click
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then 'Mute/Unmute
        OneClickGo.mnuMute_Click
    End If
End Sub

Private Sub FadeFX_Timer() 'Timer for Fading in and Fading out
    If Not IsOut Then
        If OpacityNow > 204 Then
            SetLayeredWindowAttributes hWnd, 0, (OpacityNow - 5) + 0&, &H2& + &H1
            OpacityNow = OpacityNow - 5
        Else
            FadeFX.Enabled = False
        End If
    Else
        If OpacityNow < 255 Then
            SetLayeredWindowAttributes hWnd, 0, (OpacityNow + 5) + 0&, &H2& + &H1
            OpacityNow = OpacityNow + 5
        Else
            FadeFX.Enabled = False
        End If
    End If
End Sub

Private Sub NextBtn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then
        Exit Sub
    End If
    PrevNextImg.Picture = LoadPicture(OfName("PrevNextNPressed"))
End Sub

Private Sub NextBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IsOut = True
    FadeFX.Enabled = True
    NPSVanisher.Enabled = False
    NPSVanisher.Enabled = True
    NPSVvanisher.Enabled = False
    NPSVvanisher.Enabled = True
End Sub

Private Sub NextBtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then
        Exit Sub
    End If
    PlayPauseImg.Picture = LoadPicture(OfName("PlayPauseADefault"))
    PrevNextImg.Picture = LoadPicture(OfName("PrevNextDefault"))
    OneClickGo.mnuNextSong_Click
End Sub

Private Sub NPSfx_Timer() 'Generating Skin effects using a timer
    NPSfx.Interval = 10
    SetLayeredWindowAttributes hWnd, 0, OpacityNow + 0&, &H2& + &H1
    If OneClickGo.mnuPlay.Caption = "&Pause" Then
        If OneClickGo.MyPlayer(ActiveIndex).currentMedia.duration <> 0 Then
            SeekCtrl2 = (OneClickGo.MyPlayer(ActiveIndex).Controls.currentPosition / OneClickGo.MyPlayer(ActiveIndex).currentMedia.duration) * SeekBack.Width
        End If
        If SeekCtrl2 > SeekBack.Width Then
            SeekCtrl2 = SeekBack.Width
        End If
        SeekFX.Width = SeekCtrl2
    End If
    If NPS.Width > NPSBlockerR.Left - NPSBlockerL.Width Then
        If Rotator Then
            NPS.Left = NPS.Left + 10
        Else
            NPS.Left = NPS.Left - 10
        End If
        If NPS.Left + NPS.Width <= NPSBlockerR.Left Then
            Rotator = True
            NPSfx.Interval = 500
        End If
        If NPS.Left >= NPSBlockerL.Width Then
            Rotator = False
            NPSfx.Interval = 500
        End If
    End If
End Sub

Private Sub NPSVanisher_Timer() 'Vanishing the NPSinfo frame
    IsOut = False
    FadeFX.Enabled = True
    NPSVanisher.Enabled = False
End Sub

Private Sub NPSVvanisher_Timer()
    SeekTime.Visible = False
    NPSinfo.Visible = False
    SeekFX.BackColor = RGB(255, 173, 80)
    NPSVvanisher.Enabled = False
End Sub

Private Sub PlayPauseImg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then
        Exit Sub
    End If
    Xinit = X
    Yinit = Y
    If OneClickGo.mnuPlay.Caption = "&Pause" Then
        PlayPauseImg.Picture = LoadPicture(OfName("PlayPauseAPressed"))
    Else
        PlayPauseImg.Picture = LoadPicture(OfName("PlayPauseLPressed"))
    End If
End Sub

Private Sub PlayPauseImg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    NPSinfo.Visible = True
    IsOut = True
    FadeFX.Enabled = True
    NPSVanisher.Enabled = False
    NPSVanisher.Enabled = True
    NPSVvanisher.Enabled = False
    NPSVvanisher.Enabled = True
    If Button <> 1 Then
        Exit Sub
    End If
    Move Left + X - Xinit, Top + Y - Yinit
    If X - Xinit <> 0 Or Y - Yinit <> 0 Then
        IamMoved = True
        SvSt "SkinLeft", Left
        SvSt "SkinTop", Top
    End If
End Sub

Private Sub PlayPauseImg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then
        Exit Sub
    End If
    If OneClickGo.mnuPlay.Caption = "&Pause" Then
        If Not IamMoved Then
            OneClickGo.MyPlayer(ActiveIndex).Controls.pause
            OneClickGo.mnuPlay.Caption = "&Play"
            PlayPauseImg.Picture = LoadPicture(OfName("PlayPauseLDefault"))
        Else
            PlayPauseImg.Picture = LoadPicture(OfName("PlayPauseADefault"))
        End If
    Else
        If Not IamMoved Then
            OneClickGo.MyPlayer(ActiveIndex).Controls.play
            OneClickGo.mnuPlay.Caption = "&Pause"
            PlayPauseImg.Picture = LoadPicture(OfName("PlayPauseADefault"))
        Else
            PlayPauseImg.Picture = LoadPicture(OfName("PlayPauseLDefault"))
        End If
    End If
    IamMoved = False
End Sub

Private Sub PrevBtn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then
        Exit Sub
    End If
    PrevNextImg.Picture = LoadPicture(OfName("PrevNextPPressed"))
End Sub

Private Sub PrevBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IsOut = True
    FadeFX.Enabled = True
    NPSVanisher.Enabled = False
    NPSVanisher.Enabled = True
    NPSVvanisher.Enabled = False
    NPSVvanisher.Enabled = True
End Sub

Private Sub PrevBtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then
        Exit Sub
    End If
    PlayPauseImg.Picture = LoadPicture(OfName("PlayPauseADefault"))
    PrevNextImg.Picture = LoadPicture(OfName("PrevNextDefault"))
    OneClickGo.mnuPrevSong.Checked = True
    OneClickGo.mnuNextSong_Click
End Sub

Private Sub SeekBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TempSeekTime As Integer
    IsOut = True
    FadeFX.Enabled = True
    NPSVanisher.Enabled = False
    NPSVanisher.Enabled = True
    NPSVvanisher.Enabled = False
    NPSVvanisher.Enabled = True
    SeekFX.BackColor = vbWhite
    If X > 0 And X < SeekBack.Width Then
        SeekTime.Caption = DurationStr(OneClickGo.MyPlayer(ActiveIndex).currentMedia.duration * X / SeekBack.Width) & "/" & LTrim$(OneClickGo.MyPlayer(ActiveIndex).currentMedia.durationString)
    End If
    TempSeekTime = SeekFX.Left + X - (SeekTime.Width / 2)
    If TempSeekTime < SeekBack.Left Then
        SeekTime.Left = SeekBack.Left
    ElseIf TempSeekTime > SeekBack.Left + SeekBack.Width - SeekTime.Width Then
        SeekTime.Left = SeekBack.Width + SeekBack.Left - SeekTime.Width
    Else
        SeekTime.Left = SeekFX.Left + X - (SeekTime.Width / 2)
    End If
    SeekTime.Visible = True
End Sub

Private Sub SeekBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > 0 And X < SeekBack.Width And OneClickGo.MyPlayer(ActiveIndex).URL <> vbNullString Then
        SeekFX.Width = X
        OneClickGo.MyPlayer(ActiveIndex).Controls.currentPosition = OneClickGo.MyPlayer(ActiveIndex).currentMedia.duration * X / SeekBack.Width
    End If
End Sub

Private Sub StopBtn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then
        Exit Sub
    End If
    StopImg.Picture = LoadPicture(OfName("StopPressed"))
End Sub

Private Sub StopBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IsOut = True
    FadeFX.Enabled = True
    NPSVanisher.Enabled = False
    NPSVanisher.Enabled = True
    NPSVvanisher.Enabled = False
    NPSVvanisher.Enabled = True
End Sub

Private Sub StopBtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then
        Exit Sub
    End If
    OneClickGo.MyPlayer(ActiveIndex).Controls.currentPosition = 0
    OneClickGo.MyPlayer(ActiveIndex).Controls.pause
    OneClickGo.mnuPlay.Caption = "&Play"
    SeekFX.Width = 0
    PlayPauseImg.Picture = LoadPicture(OfName("PlayPauseLDefault"))
    StopImg.Picture = LoadPicture(OfName("StopDefault"))
End Sub

Private Sub StopImg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SeekTime.Visible = False
    SeekFX.BackColor = RGB(255, 173, 80)
    IsOut = True
    FadeFX.Enabled = True
    NPSVanisher.Enabled = False
    NPSVanisher.Enabled = True
    NPSVvanisher.Enabled = False
    NPSVvanisher.Enabled = True
End Sub
