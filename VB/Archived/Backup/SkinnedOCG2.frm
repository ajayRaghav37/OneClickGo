VERSION 5.00
Begin VB.Form SkinnedOCG2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "OneClick Go!"
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "SkinnedOCG2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4680
   Begin VB.Timer NPSVanisher 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   600
   End
   Begin VB.Timer FadeFX 
      Interval        =   10
      Left            =   120
      Top             =   120
   End
   Begin VB.Timer NPSfx 
      Interval        =   10
      Left            =   120
      Top             =   1560
   End
   Begin VB.Frame NPSinfo 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   1080
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   3600
      Begin VB.Label SeekTime 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2:33/5:10"
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
         Left            =   2400
         TabIndex        =   8
         Top             =   1230
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Image NPSback2 
         Height          =   1800
         Left            =   0
         Top             =   0
         Width           =   540
      End
      Begin VB.Image NPSback1 
         Height          =   1800
         Left            =   3060
         Top             =   0
         Width           =   540
      End
      Begin VB.Label SeekBack 
         BackStyle       =   0  'Transparent
         Height          =   90
         Left            =   570
         TabIndex        =   5
         Top             =   1440
         Width           =   2490
      End
      Begin VB.Shape SeekLimit 
         Height          =   90
         Left            =   570
         Shape           =   4  'Rounded Rectangle
         Top             =   1440
         Width           =   2490
      End
      Begin VB.Shape SeekFX 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         Height          =   90
         Left            =   570
         Shape           =   4  'Rounded Rectangle
         Top             =   1440
         Width           =   15
      End
      Begin VB.Label NowPlayingSkin 
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
         Left            =   1755
         TabIndex        =   7
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Rank 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   1755
         TabIndex        =   6
         Top             =   840
         Width           =   120
      End
      Begin VB.Image NPSback 
         Height          =   1800
         Left            =   540
         Top             =   0
         Width           =   2520
      End
   End
   Begin VB.Timer NPSVvanisher 
      Interval        =   5000
      Left            =   120
      Top             =   1080
   End
   Begin VB.Label PrevBtn 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   1680
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label PlayBtn 
      BackStyle       =   0  'Transparent
      Height          =   795
      Left            =   2520
      TabIndex        =   0
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label StopBtn 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   2700
      Width           =   495
   End
   Begin VB.Label NextBtn 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   3240
      TabIndex        =   2
      Top             =   2040
      Width           =   855
   End
   Begin VB.Image ImgBack 
      Height          =   1140
      Left            =   1440
      Top             =   1920
      Width           =   2880
   End
End
Attribute VB_Name = "SkinnedOCG2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright © 2011 ANIco.in
'Welcome to the source code of OneClick Go! Skin Mode 2
'The code in this module deals with the second skin mode of OCG.
'The modification and resdistribution of the code is completely permitted.
'---------------------------------------------------------------------------

Option Explicit

Private Sub Form_Load()
    Refresher2
    SetWindowPos hWnd, -1, 0, 0, 0, 0, 1 Or 2
    SetWindowLongA hWnd, -20, GetWindowLongA(hWnd, -20) Or &H80000
    SetLayeredWindowAttributes hWnd, 0, 255 + 0&, &H2& + &H1 'Setting the SkinnedOCG as transparent to black color and always on top
    OpacityNow = 255
    Left = GtSt("SkinLeft", Screen.Width - Width)
    Top = GtSt("SkinTop", Screen.Height - Height)
    NowPlayingSkin.ForeColor = RGB(160, 174, 193)
    Rank.ForeColor = RGB(112, 126, 145)
    SeekLimit.BorderColor = RGB(128, 142, 161)
    SeekFX.BackColor = RGB(48, 55, 64)
    SeekTime.ForeColor = RGB(128, 142, 161)
    ImgBack.Picture = LoadPicture(OfName(IIf(OneClickGo.mnuPlay.Caption = "&Pause", "Pause", "Play") & "Default"))
    NPSback.Picture = LoadPicture(OfName("NPSback"))
    NPSback1.Picture = LoadPicture(OfName("NPSback1"))
    NPSback2.Picture = LoadPicture(OfName("NPSback2"))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then 'Disable Skin mode
        Unload Me
        OneClickGo.mnuSkinMode.Checked = False
        OneClickGo.Show
    ElseIf KeyCode = vbKeyF3 Then 'Cycle through skins
        SkinnedOCG.Show
        OneClickGo.mnuFreezedBlue.Checked = False
        OneClickGo.mnuRockstarGold.Checked = True
        SvSt "CurrentSkin", "0"
        Unload SkinnedOCG2
    ElseIf KeyCode = vbKeyF4 Then 'Close on Alt+F4
        If Shift = 4 Then
            Unload Me
            Unload OneClickGo
        End If
    ElseIf KeyCode = vbKeyF7 Then 'Volume down
        OneClickGo.mnuVolumeDown_Click
    ElseIf KeyCode = vbKeyF8 Then 'Volume up
        OneClickGo.mnuVolumeUp_Click
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then 'Mute/Unmute
        OneClickGo.mnuMute_Click
    End If
End Sub

Private Sub FadeFX_Timer() 'Timer for fading in and fading out
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
    Xinit = X
    Yinit = Y
    ImgBack.Picture = LoadPicture(OfName("NextPressed" & IIf(OneClickGo.mnuPlay.Caption = "&Pause", "Pause", "Play")))
End Sub

Private Sub NextBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsMove Button, X, Y
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
    If OneClickGo.mnuPlay.Caption = "&Pause" Then
        If Not IamMoved Then
            OneClickGo.mnuNextSong_Click
            ImgBack.Picture = LoadPicture(OfName("PauseDefault"))
        Else
            ImgBack.Picture = LoadPicture(OfName("PauseDefault"))
        End If
    Else
        If Not IamMoved Then
            OneClickGo.mnuNextSong_Click
            ImgBack.Picture = LoadPicture(OfName("PauseDefault"))
        Else
            ImgBack.Picture = LoadPicture(OfName("PlayDefault"))
        End If
    End If
    IamMoved = False
End Sub

Private Sub NPSback_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'Clearing things when you are not hovering the seek bar
    SeekTime.Visible = False
    SeekLimit.BorderWidth = 1
    SeekFX.BackColor = RGB(48, 55, 64)
    IsOut = True
    FadeFX.Enabled = True
    NPSVanisher.Enabled = False
    NPSVanisher.Enabled = True
    NPSVvanisher.Enabled = False
    NPSVvanisher.Enabled = True
End Sub

Private Sub NPSback1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SeekTime.Visible = False
    SeekLimit.BorderWidth = 1
    SeekFX.BackColor = RGB(48, 55, 64)
    IsOut = True
    FadeFX.Enabled = True
    NPSVanisher.Enabled = False
    NPSVanisher.Enabled = True
    NPSVvanisher.Enabled = False
    NPSVvanisher.Enabled = True
End Sub

Private Sub NPSback2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SeekTime.Visible = False
    SeekLimit.BorderWidth = 1
    SeekFX.BackColor = RGB(48, 55, 64)
    IsOut = True
    FadeFX.Enabled = True
    NPSVanisher.Enabled = False
    NPSVanisher.Enabled = True
    NPSVvanisher.Enabled = False
    NPSVvanisher.Enabled = True
End Sub

Private Sub NPSfx_Timer() 'Generate effects in the skin through a timer
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
    If NowPlayingSkin.Width > NPSback1.Left - NPSback2.Width Then
        NowPlayingSkin.Left = IIf(Rotator, NowPlayingSkin.Left + 10, NowPlayingSkin.Left - 10)
        If NowPlayingSkin.Left + NowPlayingSkin.Width <= NPSback1.Left Then
            Rotator = True
            NPSfx.Interval = 500
        End If
        If NowPlayingSkin.Left >= NPSback2.Width Then
            Rotator = False
            NPSfx.Interval = 500
        End If
    End If
End Sub

Private Sub NPSVanisher_Timer() 'Vanishing NPSinfo after a finite time
    IsOut = False
    FadeFX.Enabled = True
    NPSVanisher.Enabled = False
End Sub

Private Sub NPSVvanisher_Timer()
    SeekTime.Visible = False
    NPSinfo.Visible = False
    SeekLimit.BorderWidth = 1
    SeekFX.BackColor = RGB(48, 55, 64)
    NPSVvanisher.Enabled = False
End Sub

Private Sub PlayBtn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then
        Exit Sub
    End If
    Xinit = X
    Yinit = Y
    ImgBack.Picture = LoadPicture(OfName(IIf(OneClickGo.mnuPlay.Caption = "&Pause", "Pause", "Play") & "Pressed"))
End Sub

Private Sub PlayBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsMove Button, X, Y
    NPSinfo.Visible = True
    IsOut = True
    FadeFX.Enabled = True
    NPSVanisher.Enabled = False
    NPSVanisher.Enabled = True
    NPSVvanisher.Enabled = False
    NPSVvanisher.Enabled = True
End Sub

Private Sub PlayBtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then
        Exit Sub
    End If
    If OneClickGo.mnuPlay.Caption = "&Pause" Then
        If Not IamMoved Then
            OneClickGo.MyPlayer(ActiveIndex).Controls.pause
            OneClickGo.mnuPlay.Caption = "&Play"
            ImgBack.Picture = LoadPicture(OfName("PlayDefault"))
        Else
            ImgBack.Picture = LoadPicture(OfName("PauseDefault"))
        End If
    Else
        If Not IamMoved Then
            OneClickGo.MyPlayer(ActiveIndex).Controls.play
            OneClickGo.mnuPlay.Caption = "&Pause"
            ImgBack.Picture = LoadPicture(OfName("PauseDefault"))
        Else
            ImgBack.Picture = LoadPicture(OfName("PlayDefault"))
        End If
    End If
    IamMoved = False
End Sub

Private Sub PrevBtn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then
        Exit Sub
    End If
    Xinit = X
    Yinit = Y
    ImgBack.Picture = LoadPicture(OfName("PrevPressed" & IIf(OneClickGo.mnuPlay.Caption = "&Pause", "Pause", "Play")))
End Sub

Private Sub PrevBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsMove Button, X, Y
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
    If OneClickGo.mnuPlay.Caption = "&Pause" Then
        If Not IamMoved Then
            OneClickGo.mnuPrevSong.Checked = True
            OneClickGo.mnuNextSong_Click
            ImgBack.Picture = LoadPicture(OfName("PauseDefault"))
        Else
            ImgBack.Picture = LoadPicture(OfName("PauseDefault"))
        End If
    Else
        If Not IamMoved Then
            OneClickGo.mnuPrevSong.Checked = True
            OneClickGo.mnuNextSong_Click
            ImgBack.Picture = LoadPicture(OfName("PauseDefault"))
        Else
            ImgBack.Picture = LoadPicture(OfName("PlayDefault"))
        End If
    End If
    IamMoved = False
End Sub

Private Sub SeekBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IsOut = True
    FadeFX.Enabled = True
    NPSinfo.Visible = True
    NPSVanisher.Enabled = False
    NPSVanisher.Enabled = True
    NPSVvanisher.Enabled = False
    NPSVvanisher.Enabled = True
    SeekLimit.BorderWidth = 2
    SeekFX.BackColor = RGB(200, 200, 200)
    If X > 0 And X < SeekBack.Width Then
        SeekTime.Caption = IIf(Int(OneClickGo.MyPlayer(ActiveIndex).currentMedia.duration * X / (60 * SeekBack.Width)) < 10, "0", vbNullString) & LTrim$(Str$(Int(OneClickGo.MyPlayer(ActiveIndex).currentMedia.duration * X / (60 * SeekBack.Width)))) & ":" & IIf(Int(OneClickGo.MyPlayer(ActiveIndex).currentMedia.duration * X / SeekBack.Width) Mod 60 < 10, "0", vbNullString) & LTrim$(Str$(Int(OneClickGo.MyPlayer(ActiveIndex).currentMedia.duration * X / SeekBack.Width) Mod 60)) & "/" & LTrim$(OneClickGo.MyPlayer(ActiveIndex).currentMedia.durationString)
    End If
    If SeekFX.Left + X - (SeekTime.Width / 2) < NPSback.Left Then
        SeekTime.Left = NPSback.Left
    ElseIf SeekFX.Left + X + (SeekTime.Width / 2) > NPSback1.Left Then
        SeekTime.Left = NPSback1.Left - SeekTime.Width
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
    Xinit = X
    Yinit = Y
    ImgBack.Picture = LoadPicture(OfName("StopPressed" & IIf(OneClickGo.mnuPlay.Caption = "&Pause", "Pause", "Play")))
End Sub

Private Sub StopBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsMove Button, X, Y
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
    If OneClickGo.mnuPlay.Caption = "&Pause" Then
        If Not IamMoved Then
            OneClickGo.MyPlayer(ActiveIndex).Controls.currentPosition = 0
            OneClickGo.MyPlayer(ActiveIndex).Controls.pause
            OneClickGo.mnuPlay.Caption = "&Play"
            SeekFX.Width = 0
            ImgBack.Picture = LoadPicture(OfName("PlayDefault"))
        Else
            ImgBack.Picture = LoadPicture(OfName("PauseDefault"))
        End If
    Else
        If Not IamMoved Then
            OneClickGo.MyPlayer(ActiveIndex).Controls.currentPosition = 0
            OneClickGo.MyPlayer(ActiveIndex).Controls.pause
            OneClickGo.mnuPlay.Caption = "&Play"
            SeekFX.Width = 0
        End If
        ImgBack.Picture = LoadPicture(OfName("PlayDefault"))
    End If
    IamMoved = False
End Sub

Private Sub MsMove(Button As Integer, X As Single, Y As Single) 'Event to handle window move on mouse move
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
