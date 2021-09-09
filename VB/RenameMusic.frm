VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form RenameMusic 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  Auto Rename Music"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   ControlBox      =   0   'False
   Icon            =   "RenameMusic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5790
   Begin VB.Timer TimOut2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1800
      Top             =   1800
   End
   Begin VB.Timer TimOut 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1320
      Top             =   1800
   End
   Begin VB.Timer RenTimer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   1800
   End
   Begin VB.Timer RenTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   360
      Top             =   1800
   End
   Begin VB.CheckBox ChkMyMusic 
      BackColor       =   &H80000005&
      Caption         =   "My Music"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1500
   End
   Begin VB.CheckBox ChkAllMusic 
      BackColor       =   &H80000005&
      Caption         =   "All Music"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1500
   End
   Begin VB.OptionButton CrtLT 
      BackColor       =   &H80000005&
      Caption         =   "Album - Title"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      ToolTipText     =   "Example: Mylo Xyloto - Paradise"
      Top             =   840
      Value           =   -1  'True
      Width           =   1500
   End
   Begin VB.OptionButton CrtRT 
      BackColor       =   &H80000005&
      Caption         =   "Artist - Title"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      ToolTipText     =   "Example: Coldplay - Paradise"
      Top             =   840
      Width           =   1500
   End
   Begin VB.CommandButton BtnOK 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Start"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton BtnCancel 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin WMPLibCtl.WindowsMediaPlayer RenPlayer2 
      Height          =   30
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   30
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
      volume          =   0
      mute            =   -1  'True
      uiMode          =   "invisible"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   53
      _cy             =   53
   End
   Begin WMPLibCtl.WindowsMediaPlayer RenPlayer 
      Height          =   30
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   30
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
      volume          =   0
      mute            =   -1  'True
      uiMode          =   "invisible"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   53
      _cy             =   53
   End
   Begin VB.Label JunkLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Select renaming criteria and music to be renamed and click OK. Renaming will take some time and OCG will disappear."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   3
      Left            =   660
      TabIndex        =   9
      Top             =   240
      Width           =   5055
   End
   Begin VB.Image QstBtn 
      Height          =   450
      Left            =   120
      Picture         =   "RenameMusic.frx":000C
      Top             =   240
      Width           =   450
   End
   Begin VB.Label JunkLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rename songs in"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   660
      TabIndex        =   8
      Top             =   1200
      Width           =   1350
   End
   Begin VB.Label JunkLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Renaming criteria"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   660
      TabIndex        =   7
      Top             =   840
      Width           =   1395
   End
   Begin VB.Label JunkLbl 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Index           =   0
      Left            =   -120
      TabIndex        =   6
      Top             =   1680
      Width           =   6375
   End
End
Attribute VB_Name = "RenameMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CurrentItem As Integer
Dim NewName1 As String
Dim FirstName As String
Dim TitleName As String
Dim CurrentItem2 As Integer
Dim NewName2 As String
Dim FirstName2 As String
Dim TitleName2 As String
Dim InitTime As Double
Dim Error1 As Integer
Dim Error2 As Integer

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub BtnOK_Click()
    InitTime = Timer
    If ChkMyMusic.Value = 1 And OneClickGo.MyMusicBox.ListCount > 0 Then
    RenPlayer.URL = MyMusicFolder & "\" & OneClickGo.MyMusicBox.List(0) & ".mp3"
    RenPlayer.Controls.play
    RenTimer.Enabled = True
    End If
    If ChkAllMusic.Value = 1 And OneClickGo.AllMusicBox.ListCount > 0 Then
    RenPlayer2.URL = AllMusicFolder & "\" & OneClickGo.AllMusicBox.List(0) & ".mp3"
    RenPlayer2.Controls.play
    RenTimer2.Enabled = True
    End If
    Hide
End Sub

Private Sub ChkMyMusic_Click()
    If ChkMyMusic.Value = 0 And ChkAllMusic.Value = 0 Then
        BtnOK.Enabled = False
    Else
        BtnOK.Enabled = True
    End If
End Sub

Private Sub ChkAllMusic_Click()
    If ChkMyMusic.Value = 0 And ChkAllMusic.Value = 0 Then
        BtnOK.Enabled = False
    Else
        BtnOK.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    CanRestore = False
    Left = OneClickGo.Left + OneClickGo.Width / 2 - Width / 2
    Top = OneClickGo.Top + OneClickGo.Height / 2 - Height / 2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    CurrentItem = 0
    NewName1 = vbNullString
    FirstName = vbNullString
    TitleName = vbNullString
    FirstName = vbNullString
    CurrentItem2 = 0
    NewName2 = vbNullString
    FirstName2 = vbNullString
    TitleName2 = vbNullString
    FirstName2 = vbNullString
    InitTime = 0
    Error1 = 0
    Error2 = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CanRestore = True
End Sub

Private Sub RenTimer_Timer()
    On Error GoTo ErrorHandler1
    TimOut.Enabled = True
    If RenPlayer.currentMedia.duration <= 0 Then
        Exit Sub
    End If
    RenPlayer.Controls.stop
    If CrtLT.Value Then
        FirstName = LTrim$(UCase$(RenPlayer.currentMedia.getItemInfo("Album")))
    Else
        FirstName = LTrim$(UCase$(RenPlayer.currentMedia.getItemInfo("Artist")))
    End If
    TitleName = UCase$(RenPlayer.currentMedia.getItemInfo("Title"))
    If (FirstName = vbNullString) Or (TitleName = vbNullString) Then
        GoTo ExitSub
    End If
    If CrtLT.Value Then
        TitleName = Replace(TitleName, FirstName, "TITLE")
    End If
    RectifyTitle TitleName
    NewName1 = StrConv(FirstName & " - " & TitleName, vbProperCase)
    RectifyNewName NewName1
    NewName1 = MyMusicFolder & "\" & NewName1
    If NewName1 = Mid$(RenPlayer.currentMedia.sourceURL, 1, Len(RenPlayer.currentMedia.sourceURL) - 4) Then
        GoTo ExitSub
    End If
    If FileSystem.FileExists(NewName1 & ".mp3") Then
        TempNum = 1
        Do
            TempNum = TempNum + 1
        Loop While FileSystem.FileExists(NewName1 & CStr(TempNum) & ".mp3")
        NewName1 = NewName1 & CStr(TempNum)
    End If
    Name RenPlayer.currentMedia.sourceURL As NewName1 & ".mp3"
    OneClickGo.MyMusicBox.List(CurrentItem) = Mid$(NewName1, InStrRev(NewName1, "\") + 1, Len(NewName1) + 1 - InStrRev(NewName1, "\"))
    SvSt CStr(CurrentItem + 1), OneClickGo.MyMusicBox.List(CurrentItem)
    TimOut.Enabled = False
    GoTo ExitSub
    Exit Sub
ErrorHandler1:
    Error1 = Error1 + 1
    CurrentItem = CurrentItem + 1
    RenTimer.Enabled = False
    RenTimer.Enabled = True
    Exit Sub
ExitSub:
    CurrentItem = CurrentItem + 1
    If OneClickGo.MyMusicBox.ListCount = CurrentItem Then
        If RenTimer2.Enabled = False Then
            FinishedRenaming
        Else
            TimOut.Enabled = False
            RenTimer.Enabled = False
        End If
    Else
        RenPlayer.URL = MyMusicFolder & "\" & OneClickGo.MyMusicBox.List(CurrentItem) & ".mp3"
        RenPlayer.Controls.play
    End If
End Sub

Private Sub RenTimer2_Timer()
    On Error GoTo ErrorHandler2
    TimOut2.Enabled = True
    If RenPlayer2.currentMedia.duration <= 0 Then
        Exit Sub
    End If
    RenPlayer2.Controls.stop
    If CrtLT.Value Then
        FirstName2 = LTrim$(UCase$(RenPlayer2.currentMedia.getItemInfo("Album")))
    Else
        FirstName2 = LTrim$(UCase$(RenPlayer2.currentMedia.getItemInfo("Artist")))
    End If
    TitleName2 = UCase$(RenPlayer2.currentMedia.getItemInfo("Title"))
    If (FirstName2 = vbNullString) Or (TitleName2 = vbNullString) Then
        GoTo ExitSub
    End If
    If CrtLT.Value Then
        TitleName2 = Replace(TitleName2, FirstName2, "TITLE")
    End If
    RectifyTitle TitleName2
    NewName2 = StrConv(FirstName2 & " - " & TitleName2, vbProperCase)
    RectifyNewName NewName2
    NewName2 = AllMusicFolder & "\" & NewName2
    If NewName2 = Mid$(RenPlayer2.currentMedia.sourceURL, 1, Len(RenPlayer2.currentMedia.sourceURL) - 4) Then
        GoTo ExitSub
    End If
    If FileSystem.FileExists(NewName2 & ".mp3") Then
        TempNum = 1
        Do
            TempNum = TempNum + 1
        Loop While FileSystem.FileExists(NewName2 & CStr(TempNum) & ".mp3")
        NewName2 = NewName2 & CStr(TempNum)
    End If
    Name RenPlayer2.currentMedia.sourceURL As NewName2 & ".mp3"
    OneClickGo.AllMusicBox.List(CurrentItem2) = Mid$(NewName2, InStrRev(NewName2, "\") + 1, Len(NewName2) + 1 - InStrRev(NewName2, "\"))
    TimOut2.Enabled = False
    GoTo ExitSub
    Exit Sub
ErrorHandler2:
    Error2 = Error2 + 1
    CurrentItem2 = CurrentItem2 + 1
    RenTimer2.Enabled = False
    RenTimer2.Enabled = True
    Exit Sub
ExitSub:
    CurrentItem2 = CurrentItem2 + 1
    If OneClickGo.AllMusicBox.ListCount = CurrentItem2 Then
        If RenTimer.Enabled = False Then
            FinishedRenaming
        Else
            TimOut2.Enabled = False
            RenTimer2.Enabled = False
        End If
    Else
        RenPlayer2.URL = AllMusicFolder & "\" & OneClickGo.AllMusicBox.List(CurrentItem2) & ".mp3"
        RenPlayer2.Controls.play
    End If
End Sub

Private Sub TimOut_Timer()
    RenTimer.Enabled = False
    CurrentItem = CurrentItem + 1
    RenTimer.Enabled = True
End Sub

Private Sub TimOut2_Timer()
    RenTimer2.Enabled = False
    CurrentItem2 = CurrentItem2 + 1
    RenTimer2.Enabled = True
End Sub

Private Sub FinishedRenaming()
    Dim SongsCount As Long
    If ChkMyMusic.Value = 1 Then
        SongsCount = OneClickGo.MyMusicBox.ListCount
    End If
    If ChkAllMusic.Value = 1 Then
        SongsCount = SongsCount + OneClickGo.AllMusicBox.ListCount
    End If
    OneClickGo.RecentAction.Caption = CStr(SongsCount) & " songs renamed in " & CStr(Round((Timer - InitTime), 2)) & " seconds [" & CStr(Error1 + Error2) & " errors encountered]"
    Unload Me
End Sub

Private Sub RectifyTitle(ByRef TitleStr As String)
    TitleStr = Replace(TitleStr, "BOUNCE MIX", "Remix")
    TitleStr = Replace(TitleStr, "BOUNCE", "Remix")
    TitleStr = Replace(TitleStr, "CLUB MIX", "Remix")
    TitleStr = Replace(TitleStr, "UNPLUGGED", "Remix")
    TitleStr = Replace(TitleStr, "SAD", "Live")
    TitleStr = Replace(TitleStr, "REPRISE", "Live")
    TitleStr = Replace(TitleStr, "ACOUSTIC", "Live")
    TitleStr = Replace(TitleStr, "INSTRUMENTAL", "Live")
    TitleStr = Replace(TitleStr, "THEME", "Live")
    TitleStr = Replace(TitleStr, "_", " ")
    TitleStr = Replace(TitleStr, " - WWW.SONGS.PK", vbNullString)
    TitleStr = Replace(TitleStr, " - SONGS.PK", vbNullString)
    TitleStr = Replace(TitleStr, "SONGS.PK", vbNullString)
    TitleStr = Replace(TitleStr, "(", vbNullString)
    TitleStr = Replace(TitleStr, ")", vbNullString)
    TitleStr = Replace(TitleStr, "[", vbNullString)
    TitleStr = Replace(TitleStr, "]", vbNullString)
    TitleStr = Replace(TitleStr, "{", vbNullString)
    TitleStr = Replace(TitleStr, "}", vbNullString)
End Sub
    
Private Sub RectifyNewName(ByRef NameStr As String) 'Coping Invalid FileName errors
    NameStr = Replace(NameStr, Chr$(60), vbNullString)
    NameStr = Replace(NameStr, Chr$(62), vbNullString)
    NameStr = Replace(NameStr, Chr$(58), vbNullString)
    NameStr = Replace(NameStr, Chr$(34), vbNullString)
    NameStr = Replace(NameStr, Chr$(47), vbNullString)
    NameStr = Replace(NameStr, Chr$(92), vbNullString)
    NameStr = Replace(NameStr, Chr$(124), vbNullString)
    NameStr = Replace(NameStr, Chr$(63), vbNullString)
    NameStr = Replace(NameStr, Chr$(42), vbNullString)
End Sub
