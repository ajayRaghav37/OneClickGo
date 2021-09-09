VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{8ACD47D5-E321-474C-9A53-A800D522CE74}#1.0#0"; "MCLHotkey.OCX"
Begin VB.Form OneClickGo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "OneClick Go!"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11670
   Icon            =   "OneClickGo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   11670
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin MCLHotkey.VBHotKey HotKeyOCG 
      Index           =   0
      Left            =   4560
      Top             =   7080
      _ExtentX        =   794
      _ExtentY        =   794
      CtrlKey         =   -1  'True
      VKey            =   118
      WinKey          =   0   'False
   End
   Begin VB.Timer CFtimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7440
      Top             =   7080
   End
   Begin VB.CommandButton HideBtn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Hide"
      Height          =   330
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CancelBtn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox SearchAll 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   8520
      TabIndex        =   28
      TabStop         =   0   'False
      Text            =   "Search All Music"
      Top             =   90
      Width           =   3015
   End
   Begin VB.TextBox SearchMy 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   2760
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   "Search My Music"
      Top             =   90
      Width           =   3015
   End
   Begin VB.ListBox AllSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      ItemData        =   "OneClickGo.frx":139D9
      Left            =   8520
      List            =   "OneClickGo.frx":139DB
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   390
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ListBox MySearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      ItemData        =   "OneClickGo.frx":139DD
      Left            =   2760
      List            =   "OneClickGo.frx":139DF
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   390
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSComDlg.CommonDialog MyDialog 
      Left            =   7920
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame ColorSchemeBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Customize Display"
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
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Visible         =   0   'False
      Width           =   5655
      Begin VB.Label CCSC 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   7
         Left            =   3360
         TabIndex        =   22
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label CCSC 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   1920
         TabIndex        =   21
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label CCSC 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   3360
         TabIndex        =   20
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label CCSC 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   19
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label CCSC 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   18
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label CCSC 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   17
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label CCSC 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   16
         Top             =   960
         Width           =   255
      End
      Begin VB.Label CCSC 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   15
         Top             =   960
         Width           =   255
      End
      Begin VB.Label CCSF 
         BackStyle       =   0  'Transparent
         Caption         =   "All Music List"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label CCSF 
         BackStyle       =   0  'Transparent
         Caption         =   "My Music List"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label CCSF 
         BackStyle       =   0  'Transparent
         Caption         =   "Status Bar"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label CCSF 
         BackStyle       =   0  'Transparent
         Caption         =   "Main Window"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label CCS3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ForeColor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3000
         TabIndex        =   10
         Top             =   480
         Width           =   825
      End
      Begin VB.Label CCS2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BackColor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1680
         TabIndex        =   9
         Top             =   480
         Width           =   885
      End
      Begin VB.Label CCS1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Font"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   390
      End
   End
   Begin VB.Timer MainTimer 
      Interval        =   100
      Left            =   8400
      Top             =   7080
   End
   Begin VB.ListBox MyMusicBox 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   6405
      IntegralHeight  =   0   'False
      ItemData        =   "OneClickGo.frx":139E1
      Left            =   5895
      List            =   "OneClickGo.frx":139E3
      TabIndex        =   1
      Top             =   450
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.ListBox AllMusicBox 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3780
      IntegralHeight  =   0   'False
      ItemData        =   "OneClickGo.frx":139E5
      Left            =   120
      List            =   "OneClickGo.frx":139E7
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   450
      Visible         =   0   'False
      Width           =   5655
   End
   Begin MCLHotkey.VBHotKey HotKeyOCG 
      Index           =   1
      Left            =   5040
      Top             =   7080
      _ExtentX        =   794
      _ExtentY        =   794
      CtrlKey         =   -1  'True
      VKey            =   119
      WinKey          =   0   'False
   End
   Begin MCLHotkey.VBHotKey HotKeyOCG 
      Index           =   2
      Left            =   5520
      Top             =   7080
      _ExtentX        =   794
      _ExtentY        =   794
      CtrlKey         =   -1  'True
      VKey            =   120
      WinKey          =   0   'False
   End
   Begin MCLHotkey.VBHotKey HotKeyOCG 
      Index           =   3
      Left            =   6000
      Top             =   7080
      _ExtentX        =   794
      _ExtentY        =   794
      CtrlKey         =   -1  'True
      VKey            =   121
      WinKey          =   0   'False
   End
   Begin MCLHotkey.VBHotKey HotKeyOCG 
      Index           =   4
      Left            =   6480
      Top             =   7080
      _ExtentX        =   794
      _ExtentY        =   794
      CtrlKey         =   -1  'True
      VKey            =   122
      WinKey          =   0   'False
   End
   Begin MCLHotkey.VBHotKey HotKeyOCG 
      Index           =   5
      Left            =   6960
      Top             =   7080
      _ExtentX        =   794
      _ExtentY        =   794
      CtrlKey         =   -1  'True
      VKey            =   123
      WinKey          =   0   'False
   End
   Begin VB.Label SeekFX 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Left            =   120
      TabIndex        =   24
      Top             =   7080
      Width           =   11415
   End
   Begin VB.Label RecentAction 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   240
      TabIndex        =   4
      Top             =   7185
      UseMnemonic     =   0   'False
      Width           =   10815
   End
   Begin VB.Label MyMusicBoxLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "  My Music (0/0)"
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
      Height          =   255
      Left            =   5880
      TabIndex        =   6
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label AllMusicBoxLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "  All Music (0/0)"
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
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label SeekBack 
      BackColor       =   &H00C0C0FF&
      Height          =   405
      Left            =   120
      TabIndex        =   23
      Top             =   7080
      Width           =   15
   End
   Begin VB.Label StatusBar 
      BackColor       =   &H00000000&
      Height          =   405
      Left            =   120
      TabIndex        =   3
      Top             =   7080
      Width           =   11415
   End
   Begin WMPLibCtl.WindowsMediaPlayer MyPlayer 
      Height          =   30
      Index           =   0
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6900
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
      volume          =   100
      mute            =   0   'False
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
   Begin WMPLibCtl.WindowsMediaPlayer MyPlayer 
      Height          =   30
      Index           =   1
      Left            =   240
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   7050
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
      volume          =   100
      mute            =   0   'False
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
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Begin VB.Menu mnuTransferNow 
         Caption         =   "&Move to My Music"
      End
      Begin VB.Menu mnuRenameFile 
         Caption         =   "&Rename Song"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuDeleteFile 
         Caption         =   "&Delete Song"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuSeparator0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save My Music"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFolderChange 
         Caption         =   "&Change Folders and Reload"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^{F7}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuSkinMode 
         Caption         =   "&Skin Mode"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuDND 
         Caption         =   "&Do Not Disturb"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuHit 
         Caption         =   "&Copy Music Chart"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuFindSong 
         Caption         =   "&Find Song"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuGaming 
         Caption         =   "&Gaming Mode"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStart 
         Caption         =   "&Auto Start and Hide"
      End
      Begin VB.Menu mnuMusicLeft 
         Caption         =   "&My Music On Left"
      End
      Begin VB.Menu mnuColor 
         Caption         =   "C&ustomize Display"
      End
      Begin VB.Menu mnuSkinChooser 
         Caption         =   "S&kin Chooser"
         Begin VB.Menu mnuRockstarGold 
            Caption         =   "&Rockstar Gold"
         End
         Begin VB.Menu mnuFreezedBlue 
            Caption         =   "&Freezed Blue"
         End
      End
   End
   Begin VB.Menu mnuMedia 
      Caption         =   "&Media"
      Begin VB.Menu mnuPlay 
         Caption         =   "&Play"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuStop 
         Caption         =   "&Stop"
      End
      Begin VB.Menu mnuPrevSong 
         Caption         =   "P&revious Song"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuNextSong 
         Caption         =   "&Next Song"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLinear 
         Caption         =   "&Linear"
      End
      Begin VB.Menu mnuLinRev 
         Caption         =   "&Reverse Linear"
      End
      Begin VB.Menu mnuShuffle 
         Caption         =   "S&huffle"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuRepeat 
         Caption         =   "&Auto Repeat"
      End
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCrossfade 
         Caption         =   "&Crossfade Music"
      End
      Begin VB.Menu mnuPlaySpeed 
         Caption         =   "Pla&y Speed (100%)"
         Begin VB.Menu mnuPSreset 
            Caption         =   "&Reset"
            Shortcut        =   +{F6}
         End
         Begin VB.Menu mnuPSslower 
            Caption         =   "&Slower"
            Shortcut        =   +{F7}
         End
         Begin VB.Menu mnuPSfaster 
            Caption         =   "&Faster"
            Shortcut        =   +{F8}
         End
      End
      Begin VB.Menu mnuBalance 
         Caption         =   "&Balance (C)"
         Begin VB.Menu mnuBcentre 
            Caption         =   "&Centre"
            Shortcut        =   +{F9}
         End
         Begin VB.Menu mnuBleft 
            Caption         =   "&Left Weighted"
            Shortcut        =   +{F11}
         End
         Begin VB.Menu mnuBright 
            Caption         =   "&Right Weighted"
            Shortcut        =   +{F12}
         End
      End
      Begin VB.Menu mnuVolume 
         Caption         =   "&Volume"
         Begin VB.Menu mnuMute 
            Caption         =   "&Mute"
            Shortcut        =   {F6}
         End
         Begin VB.Menu mnuVolumeDown 
            Caption         =   "&Decrease"
            Shortcut        =   {F7}
         End
         Begin VB.Menu mnuVolumeUp 
            Caption         =   "&Increase"
            Shortcut        =   {F8}
         End
      End
   End
   Begin VB.Menu mnuPlaylist 
      Caption         =   "&Playlist"
      Begin VB.Menu mnuUp 
         Caption         =   "&Increase Rating (Move Up)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuDown 
         Caption         =   "&Decrease Rating (Move Down)"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRating 
         Caption         =   "&Custom Rating"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuSeparator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBriefHistory 
         Caption         =   "&Brief History"
      End
      Begin VB.Menu mnuAutoRename 
         Caption         =   "&Auto Rename Music"
      End
      Begin VB.Menu mnuSend 
         Caption         =   "&Send Music to Device"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Tools"
      Begin VB.Menu mnuBackup 
         Caption         =   "&Backup Playlist and Settings"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore Playlist and Settings"
      End
      Begin VB.Menu mnuReset 
         Caption         =   "&Factory Reset"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuSupport 
         Caption         =   "&Support"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuLicense 
         Caption         =   "&View License"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "OneClickGo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright © 2011 ANIco.in
'Welcome to the source code of OneClick Go!
'The code in this module deals with the working of main OCG window.
'The modification of the code is completely permitted.
'---------------------------------------------------------------------------
Option Explicit

Dim PlaySpd As Integer
Dim TempStr2 As String
Dim AutoRepeated As Boolean
Dim ActiveCS As Integer
Dim ASvolin As Boolean
Dim ClickedNext As Boolean
Dim CurrentSong As Integer
Dim ASradio As Boolean
Dim PrevRectify As Boolean
Dim PassiveIndex As Integer
Dim IsMaxWin As Boolean
Dim Reported As Boolean
Dim TimesCount As Integer
Dim IsContinue As Boolean
Dim TempStr As String
Dim MyControl As Variant

Private Sub Form_Load()
    Dim RemExt As Integer
    On Error GoTo DriveSel 'Trapping an obvious error when the user selects a drive as AllMusicFolder or MyMusicFolder
    
    'Loading the Visual Interface
    mnuStop.Caption = "&Stop" & Chr$(9) & "F10"
    mnuTransferNow.Caption = "&Move to My Music" & Chr$(9) & "Space"
    ColorScheme GtSt("Color0", vbButtonFace), GtSt("Color1", vbButtonText), GtSt("Color2", vbWindowBackground), GtSt("Color3", vbWindowText), GtSt("Color4", vbWindowBackground), GtSt("Color5", vbWindowText), GtSt("Color6", vbButtonFace), GtSt("Color7", vbButtonText)
    
    'Loading Settings from Registry
    mnuLinear.Checked = GtSt("Linear", True)
    mnuLinRev.Checked = GtSt("LinRev", False)
    mnuShuffle.Checked = GtSt("Shuffle", False)
    mnuLinear.Enabled = IIf(mnuLinear.Checked, False, True)
    mnuLinRev.Enabled = IIf(mnuLinRev.Checked, False, True)
    mnuShuffle.Enabled = IIf(mnuShuffle.Checked, False, True)
    mnuRepeat.Checked = GtSt("Repeat", False)
    mnuMusicLeft.Checked = GtSt("MusicLeft", False)
    mnuStart.Checked = GtSt("AutoStart", False)
    mnuCrossfade.Checked = GtSt("Crossfade", False)
    CurrentSkin = Val(LTrim$(GtSt("CurrentSkin", "0")))
    Select Case CurrentSkin
        Case 0
            mnuRockstarGold.Checked = True
        Case 1
            mnuFreezedBlue.Checked = True
    End Select
    Width = Screen.Width * 0.7
    Height = Width * 0.7
    PlaySpd = 100

    'Checking if OCG is being run for the first time, if so, ask AllMusicFolder and MyMusicFolder
    If GtSt("UsedBefore", "0") = "0" Then
        OldName = "Select the folder that you would like to use as your All Music folder."
        Do
            Set ShellOpener = ShellSystem.BrowseForFolder(0, OldName, &H1, 17)
            OldName = "Select the folder that you would like to use as your All Music folder."
            If Not (ShellOpener Is Nothing) Then
                TempStr = ShellOpener.ParentFolder.ParseName(ShellOpener.Title).Path
                If FileSystem.folderexists(TempStr) Then
                    Exit Do
                Else
                    OldName = OldName & " [Invalid Folder]"
                End If
            Else
                If FileSystem.folderexists(GtSt("AllMusicFolder", vbNullString)) Then
                    TempStr = GtSt("AllMusicFolder", vbNullString)
                    Exit Do
                Else
                    OldName = OldName & " [Cannot be cancelled]"
                End If
            End If
        Loop
        SvSt "AllMusicFolder", TempStr
        OldName = "Select the folder that you would like to use as your My Music folder."
        Do
            Set ShellOpener = ShellSystem.BrowseForFolder(0, OldName, &H1, 17)
            OldName = "Select the folder that you would like to use as your My Music folder."
            If Not (ShellOpener Is Nothing) Then
                TempStr = ShellOpener.ParentFolder.ParseName(ShellOpener.Title).Path
                If TempStr <> GtSt("AllMusicFolder") And FileSystem.folderexists(TempStr) Then
                    Exit Do
                Else
                    OldName = IIf(TempStr = GtSt("AllMusicFolder"), OldName & " [Cannot be same as All Music Folder]", OldName & " [Invalid Folder]")
                End If
            Else
                If FileSystem.folderexists(GtSt("MyMusicFolder", vbNullString)) Then
                    TempStr = GtSt("MyMusicFolder", vbNullString)
                    Exit Do
                Else
                    OldName = OldName & " [Cannot be cancelled]"
                End If
            End If
        Loop
        SvSt "MyMusicFolder", TempStr
        If TempStr <> GtSt("MyMusicFolder", vbNullString) And GtSt("MyMusicFolder", vbNullString) <> vbNullString Then
            MyMsgBox = MsgBox("You are about to change your 'My Music' directory that might result in corruption of your playlist." & Chr$(13) & "Are you sure you want to continue?", vbYesNo + vbCritical + vbSystemModal, "WARNING")
            If MyMsgBox = vbYes Then
                SvSt "MyMusicFolder", TempStr
            End If
        End If
        SvSt "UsedBefore", "1"
    End If
    AllMusicFolder = GtSt("AllMusicFolder")
    AllMusicBoxLabel.ToolTipText = AllMusicFolder
    MyMusicFolder = GtSt("MyMusicFolder")
    MyMusicBoxLabel.ToolTipText = MyMusicFolder

    'Retrieving My Music list
    
    Set MyFiles = FileSystem.GetFolder(MyMusicFolder)
    Set MySongs = MyFiles.Files
    FilNum = 1
    Do
        If GtSt(Str$(FilNum), vbNullString) <> vbNullString Then
            If FileSystem.FileExists(MyMusicFolder & "\" & GtSt(Str$(FilNum)) & ".mp3") Then
                MyMusicBox.AddItem GtSt(Str$(FilNum)) & ".mp3"
            ElseIf FileSystem.FileExists(AllMusicFolder & "\" & GtSt(Str$(FilNum)) & ".mp3") Then
                FileSystem.movefile AllMusicFolder & "\" & GtSt(Str$(FilNum)) & ".mp3", MyMusicFolder & "\" & GtSt(Str$(FilNum)) & ".mp3"
                MyMusicBox.AddItem GtSt(Str$(FilNum)) & ".mp3"
            Else
                MyMsgBox = MsgBox("This song has been deleted or moved to another location:" & Chr$(13) & GtSt(Str$(FilNum)), vbCritical, "Song Not Found")
            End If
        Else
            Exit Do
        End If
        FilNum = FilNum + 1
    Loop
    
    'Searching for new songs in My Music directory
    FilNum = 0
    For Each MySong In MySongs
        If UCase$(Right(MySong.Name, 3)) = "MP3" Then
            FilNum = FilNum + 1
        End If
    Next
    If FilNum <> MyMusicBox.ListCount Then

        'Adding new songs found in My Music directory
        TempNum = 0
        For Each MySong In MySongs
            If UCase$(Right(MySong.Name, 3)) = "MP3" Then
                NewEntry = True
                If SendMessage(MyMusicBox.hWnd, &H18F, -1, ByVal MySong.Name) <> -1 Then
                    NewEntry = False
                End If
                If NewEntry Then
                    MyMusicBox.AddItem MySong.Name
                    TempNum = TempNum + 1
                End If
            End If
        Next
    End If

    'Clean Up .mp3 from all songs
    For RemExt = 0 To MyMusicBox.ListCount - 1
        MyMusicBox.List(RemExt) = Mid$(MyMusicBox.List(RemExt), 1, (Len(MyMusicBox.List(RemExt)) - 4))
    Next

    'Retrieving All Music list
    Set AllFiles = FileSystem.GetFolder(AllMusicFolder)
    Set AllSongs = AllFiles.Files
    For Each AllSong In AllSongs
        If UCase$(Right(AllSong.Name, 3)) = "MP3" Then
            AllMusicBox.AddItem Mid$(AllSong.Name, 1, Len(AllSong.Name) - 4)
        End If
    Next

    'Updating Interface
    AllMusicBox.Visible = True
    MyMusicBox.Visible = True
    AllMusicBoxLabel.Caption = "  All Music (" & Trim$(Str$(AllMusicBox.ListIndex + 1) & "/" & Trim$(Str$(AllMusicBox.ListCount)) & ")")
    MyMusicBoxLabel.Caption = "  My Music (" & Trim$(Str$(MyMusicBox.ListIndex & 1) & "/" & Trim$(Str$(MyMusicBox.ListCount)) & ")")
    If Not mnuStart.Checked Then
        RecentAction.Caption = "Welcome to OneClick Go! - Loaded successfully" & IIf(TempNum > 0, " (" & LTrim$(Str$(TempNum)) & " new " & IIf(TempNum = 1, "song", "songs") & " added to My Music)", vbNullString)
        Show
    Else
        'In case of Auto-Start
        If GtSt("CantGame", "0") = "0" Then
            If mnuLinear.Checked Then
                MyMusicBox.ListIndex = 0
            ElseIf mnuLinRev.Checked Then
                MyMusicBox.ListIndex = MyMusicBox.ListCount - 1
            Else
                Randomize
                MyMusicBox.ListIndex = Round((MyMusicBox.ListCount - 1) * Rnd, 0)
                ASradio = True
                ASvolin = True
                MyPlayer(ActiveIndex).settings.volume = 0
            End If
            mnuGaming_Click
            CanRestore = True
        Else
            SvSt "CantGame", "0"
        End If
    End If
    Exit Sub
DriveSel:
    TempStr = Mid$(ShellOpener.Title, Len(ShellOpener.Title) - 2, 2) & "\"
    Resume Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Checking if the user have made changes to playlist
    Dim MadeChanges As Boolean
    For FilNum = 1 To MyMusicBox.ListCount
        If MyMusicBox.List(FilNum - 1) <> GtSt(Str$(FilNum)) Then
            MadeChanges = True
            Exit For
        End If
    Next
    If GtSt(Str$(MyMusicBox.ListCount + 1)) <> vbNullString Then
        MadeChanges = True
    End If
    'Asking the user if he wants to save the changes if made any
    If MadeChanges Then
        MyMsgBox = MsgBox("Do you want to save the changes you made to your music list?", vbYesNoCancel + vbQuestion + vbDefaultButton1, "Save Playlist")
        If MyMsgBox = vbYes Or MyMsgBox = vbNo Then
            If MyMsgBox = vbYes Then
                mnuSave_Click
            End If
            MadeChanges = False
            If mnuFolderChange.Checked Then
                Exit Sub
            End If
            Cancel = 0
        ElseIf MyMsgBox = vbCancel Then
            Cancel = 1
            mnuFolderChange.Checked = False
            If mnuSkinMode.Checked Then
                If mnuRockstarGold.Checked Then
                    SkinnedOCG.Show
                Else
                    SkinnedOCG2.Show
                End If
            End If
        End If
    Else
        If mnuFolderChange.Checked Then
            Exit Sub
        End If
        Cancel = 0
    End If
End Sub

Private Sub Form_Resize()
    If WindowState = 2 Then
        IsMaxWin = True
    End If
    If WindowState = 0 Then
        IsMaxWin = False
    End If
    If WindowState = 0 Then
        If Width < 10000 Then
            Width = 10000
        End If
        If Height < 6000 Then
            Height = 6000
        End If
    End If
    If WindowState <> 1 Then
        AllMusicBox.Width = (ScaleWidth - 360) / 2
        SearchAll.Width = AllMusicBox.Width / 2
        AllSearch.Width = SearchAll.Width
        MySearch.Width = SearchAll.Width
        SearchMy.Width = SearchAll.Width
        AllMusicBox.Height = IIf(ColorSchemeBox.Visible, ScaleHeight - 960 - ColorSchemeBox.Height - 120, ScaleHeight - 960)
        MyMusicBox.Width = AllMusicBox.Width
        MyMusicBox.Height = ScaleHeight - 960
        MyMusicBox.Left = IIf(mnuMusicLeft.Checked, 120, MyMusicBox.Width + 240)
        AllMusicBox.Left = IIf(mnuMusicLeft.Checked, AllMusicBox.Width + 240, 120)
        AllMusicBoxLabel.Left = AllMusicBox.Left
        AllMusicBoxLabel.Width = AllMusicBox.Width
        MyMusicBoxLabel.Left = MyMusicBox.Left
        MyMusicBoxLabel.Width = MyMusicBox.Width
        AllSearch.Left = AllMusicBox.Left + (AllMusicBox.Width / 2)
        SearchAll.Left = AllSearch.Left
        MySearch.Left = MyMusicBox.Left + (MyMusicBox.Width / 2)
        SearchMy.Left = MySearch.Left
        ColorSchemeBox.Left = AllMusicBox.Left
        ColorSchemeBox.Top = AllMusicBox.Top + AllMusicBox.Height + 120
        ColorSchemeBox.Width = AllMusicBox.Width
        StatusBar.Top = MyMusicBox.Top + MyMusicBox.Height + 60
        StatusBar.Width = ScaleWidth - 240
        RecentAction.Top = StatusBar.Top + 75
        RecentAction.Width = StatusBar.Width - 630
        SeekFX.Top = StatusBar.Top
        SeekFX.Width = StatusBar.Width
        SeekBack.Top = StatusBar.Top
        CancelBtn.Left = StatusBar.Left + StatusBar.Width - CancelBtn.Width - 30
        HideBtn.Left = CancelBtn.Left - HideBtn.Width
        CancelBtn.Top = StatusBar.Top + 45
        HideBtn.Top = CancelBtn.Top
        CCS3.Left = 120 + ColorSchemeBox.Width - 1500
        CCS2.Left = CCS3.Left - 1200
        For FilNum = 0 To 7
            CCSC(FilNum).Left = IIf(FilNum Mod 2 = 0, CCS2.Left + (CCS2.Width - CCSC(FilNum).Width) / 2, CCS3.Left + (CCS3.Width - CCSC(FilNum).Width) / 2)
        Next
    End If
End Sub

Private Sub AllMusicBox_Click()
    MyBoxesClick AllMusicBox
End Sub

Private Sub AllMusicBox_DblClick()
    MyBoxesDblClick AllMusicBox
End Sub

Private Sub AllMusicBox_KeyUp(KeyCode As Integer, Shift As Integer)
    MyBoxesKeyUp AllMusicBox, KeyCode
End Sub

Private Sub AllSearch_DblClick()
    If AllSearch.ListIndex <> -1 Then
        AllMusicBox.ListIndex = Val(Mid$(AllSearch.List(AllSearch.ListIndex), InStrRev(AllSearch.List(AllSearch.ListIndex), "(") + 1, Len(AllSearch.List(AllSearch.ListIndex)) - InStrRev(AllSearch.List(AllSearch.ListIndex), "(") - 1)) - 1
    End If
    SearchAll.Text = vbNullString
    AllMusicBox.SetFocus
    SearchAll_LostFocus
End Sub

Private Sub AllSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp And AllSearch.ListIndex = 0 Then
        SearchAll.SetFocus
    End If
End Sub

Private Sub AllSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        AllSearch_DblClick
    End If
End Sub

Private Sub CancelBtn_Click() 'User clicked cancel when the music was being sent
    CancelBtn.Visible = False
    HideBtn.Visible = False
    Unload SendMusic
    mnuSend.Caption = "&Send Music to Device"
End Sub

Private Sub CCSC_Click(Index As Integer) 'Changing the visual interface as per user's choice
    Dim NewColor As Long
    On Error GoTo GetOut
    MyDialog.Color = CCSC(Index).BackColor
    MyDialog.ShowColor
    NewColor = MyDialog.Color
    Select Case Index
        Case 0
            BackColor = NewColor
            ColorSchemeBox.BackColor = NewColor
        Case 1
            AllMusicBoxLabel.ForeColor = NewColor
            MyMusicBoxLabel.ForeColor = NewColor
            For FilNum = 0 To 7
                CCSC(FilNum).ForeColor = NewColor
            Next
            For FilNum = 0 To 3
                CCSF(FilNum).ForeColor = NewColor
            Next
            ColorSchemeBox.ForeColor = NewColor
            CCS1.ForeColor = NewColor
            CCS2.ForeColor = NewColor
            CCS3.ForeColor = NewColor
        Case 2
            AllMusicBox.BackColor = NewColor
            AllSearch.BackColor = NewColor
            SearchAll.BackColor = NewColor
        Case 3
            AllMusicBox.ForeColor = NewColor
            AllSearch.ForeColor = NewColor
            SearchAll.ForeColor = NewColor
        Case 4
            MyMusicBox.BackColor = NewColor
            MySearch.BackColor = NewColor
            SearchMy.BackColor = NewColor
        Case 5
            MyMusicBox.ForeColor = NewColor
            MySearch.ForeColor = NewColor
            SearchMy.ForeColor = NewColor
        Case 6
            StatusBar.BackColor = NewColor
            ColorCodeToRGB StatusBar.BackColor
            SeekBack.BackColor = RGB(IIf(cRed < 128, cRed + 16, cRed - 16), IIf(cGreen < 128, cGreen + 16, cGreen - 16), IIf(cBlue < 128, cBlue + 16, cBlue - 16))
            ColorCodeToRGB AllMusicBox.BackColor
            SearchAll.BackColor = RGB(IIf(cRed < 128, cRed + 16, cRed - 16), IIf(cGreen < 128, cGreen + 16, cGreen - 16), IIf(cBlue < 128, cBlue + 16, cBlue - 16))
            ColorCodeToRGB MyMusicBox.BackColor
            SearchMy.BackColor = RGB(IIf(cRed < 128, cRed + 16, cRed - 16), IIf(cGreen < 128, cGreen + 16, cGreen - 16), IIf(cBlue < 128, cBlue + 16, cBlue - 16))
        Case 7
            RecentAction.ForeColor = NewColor
    End Select
    CCSC(Index).BackColor = NewColor
    SvSt "Color" & Trim$(Str$(Index)), Str$(NewColor)
    Exit Sub
GetOut:
    Exit Sub
End Sub

Private Sub CCSF_Click(Index As Integer)
    Dim NFN As String
    Dim NFB As Boolean
    Dim NFI As Boolean
    Dim NFS As Single
    On Error GoTo GetOut
    MyDialog.FontName = CCSF(Index).FontName
    MyDialog.FontBold = CCSF(Index).FontBold
    MyDialog.FontItalic = CCSF(Index).FontItalic
    MyDialog.FontSize = CCSF(Index).FontSize
    MyDialog.ShowFont
    NFN = MyDialog.FontName
    NFB = MyDialog.FontBold
    NFI = MyDialog.FontItalic
    NFS = MyDialog.FontSize
    Select Case Index
        Case 0
            NewFont AllMusicBoxLabel, NFN, NFB, NFI, NFS
            NewFont MyMusicBoxLabel, NFN, NFB, NFI, NFS
            NewFont ColorSchemeBox, NFN, NFB, NFI, NFS
        Case 1
            NewFont AllMusicBox, NFN, NFB, NFI, NFS
        Case 2
            NewFont MyMusicBox, NFN, NFB, NFI, NFS
        Case 3
            NewFont RecentAction, NFN, NFB, NFI, NFS
    End Select
    NewFont CCSF(Index), NFN, NFB, NFI, NFS
    SvSt "FN" & Trim$(Str$(Index)), NFN
    SvSt "FB" & Trim$(Str$(Index)), NFB
    SvSt "FI" & Trim$(Str$(Index)), NFI
    SvSt "FS" & Trim$(Str$(Index)), NFS
    Exit Sub
GetOut:
    Exit Sub
End Sub

Private Sub CFtimer_Timer() 'Crossfading Process
    Dim CFvol As Single
    TimesCount = TimesCount + 1
    If Not ClickedNext Then
        PrevRectify = False
        CFtimer.Interval = (MyPlayer(ActiveIndex).currentMedia.duration - 1 - MyPlayer(ActiveIndex).Controls.currentPosition) * 10
        mnuNextSong_Click
        ClickedNext = True
    End If
    CFvol = (Val(GtSt("Volume", "67")) * (100 - TimesCount)) / 100
    If CFvol < 0 Then
        CFvol = 0
    End If
    If CFvol > 100 Then
        CFvol = 100
    End If
    MyPlayer(ActiveIndex).settings.volume = Round(CFvol, 0)
    MyPlayer(PassiveIndex).settings.volume = Val(GtSt("Volume", "67")) - CFvol
    If MyPlayer(ActiveIndex).playState = wmppsStopped Then
        CFtimer.Enabled = False
        MyPlayer(ActiveIndex).settings.volume = 0
        MyPlayer(PassiveIndex).settings.volume = Val(GtSt("Volume", "67"))
        ClickedNext = False
        MyPlayer(ActiveIndex).URL = vbNullString
        Reported = False
        ActiveIndex = IIf(ActiveIndex = 0, 1, 0)
        PassiveIndex = IIf(PassiveIndex = 1, 0, 1)
        IsContinue = True
        If PlayingAll Then
            If Not mnuDND.Checked Then
                AllMusicBox.ListIndex = NowPlaying
            End If
            RecentAction.Caption = "Now Playing : " & AllMusicBox.List(NowPlaying)
        Else
            If Not mnuDND.Checked Then
                MyMusicBox.ListIndex = NowPlaying
            End If
            RecentAction.Caption = "Now Playing : " & MyMusicBox.List(NowPlaying)
        End If
        IsContinue = False
        PrevRectify = False
        TimesCount = 0
        ActiveCS = CurrentSong
    End If
End Sub

Private Sub MainTimer_Timer() 'The main Timer of OCG that performs many tasks simultaneously
    Dim SeekCtrl As Integer
    On Error Resume Next
    
    'Play Speed setting
    MyPlayer(0).settings.Rate = PlaySpd / 100
    MyPlayer(1).settings.Rate = PlaySpd / 100

    'Awaiting Quick Restore
    If GtSt("QuickRestore", "0") = "1" Then
        HotKeyOCG_HotkeyPressed 1
        SvSt "QuickRestore", "0"
    End If

    'Crossfading functionality
    If MyPlayer(ActiveIndex).URL <> vbNullString And mnuCrossfade.Checked And Not Reported Then
        If MyPlayer(ActiveIndex).Controls.currentPosition >= MyPlayer(ActiveIndex).currentMedia.duration - 11 And MyPlayer(ActiveIndex).currentMedia.duration > 15 Then
            If mnuRepeat.Checked Then
                AutoRepeated = True
            End If
            CFtimer.Interval = 100 / PlaySpd
            CFtimer.Enabled = True
            Reported = True
        End If
    End If

    'Tuning in the Auto-Start functionality
    If MyPlayer(ActiveIndex).Controls.currentPosition > 0 Then
        If ASradio Then
            SeekFX_MouseUp 1, 0, Rnd * SeekFX.Width, 45
            ASradio = False
        End If
        If ASvolin Then
            MyPlayer(ActiveIndex).settings.volume = MyPlayer(ActiveIndex).settings.volume + Val(GtSt("Volume", "67")) / 20
            If MyPlayer(ActiveIndex).settings.volume >= Val(GtSt("Volume", "67")) Then
                MyPlayer(ActiveIndex).settings.volume = Val(GtSt("Volume", "67"))
                ASvolin = False
            End If
        End If
    End If

    'Change the song as soon as one song is done playing
    If MyPlayer(ActiveIndex).playState = wmppsStopped And Not mnuCrossfade.Checked Then
        If mnuRepeat.Checked Then
            MyPlayer(ActiveIndex).Controls.play
        Else
            mnuNextSong_Click
        End If
    End If

    'Updating the seek bar as the song proceeds
    If Not mnuSkinMode.Checked And Not mnuGaming.Checked Then
        Randomize
        'If mnuPlay.Caption = "&Pause" Then
        If MyPlayer(ActiveIndex).URL <> vbNullString Then
            If MyPlayer(ActiveIndex).currentMedia.duration > 15 Then
                If MyPlayer(ActiveIndex).currentMedia.duration <> 0 Then
                    SeekCtrl = (MyPlayer(ActiveIndex).Controls.currentPosition / MyPlayer(ActiveIndex).currentMedia.duration) * SeekFX.Width
                End If
            End If
            If SeekCtrl > SeekFX.Width Then
                SeekCtrl = SeekFX.Width
            End If
            SeekBack.Width = SeekCtrl
        End If

        'MainTimer Timer as a Menu Managager
        mnuTransferNow.Enabled = IIf(AllMusicBox.ListIndex = -1 And MyMusicBox.ListIndex = -1, False, True)
        mnuDeleteFile.Enabled = IIf(AllMusicBox.ListIndex = -1 And MyMusicBox.ListIndex = -1, False, True)
        mnuRenameFile.Enabled = IIf(AllMusicBox.ListIndex = -1 And MyMusicBox.ListIndex = -1, False, True)
        mnuHit.Enabled = IIf(MyMusicBox.ListCount = 0, False, True)
        mnuUp.Enabled = IIf(MyMusicBox.ListIndex = -1 Or TransferFrom = 0 Or MyMusicBox.ListIndex = 0, False, True)
        mnuDown.Enabled = IIf(MyMusicBox.ListIndex = -1 Or TransferFrom = 0 Or MyMusicBox.ListIndex = MyMusicBox.ListCount - 1, False, True)
        mnuRating.Enabled = IIf(MyMusicBox.ListIndex = -1 Or TransferFrom = 0, False, True)
        mnuMedia.Enabled = IIf(MyPlayer(ActiveIndex).URL = vbNullString, False, True)
        mnuSkinMode.Enabled = IIf(MyPlayer(ActiveIndex).URL = vbNullString, False, True)
        mnuSend.Enabled = IIf(MyMusicBox.ListCount = 0, False, True)
        mnuSave.Enabled = IIf(MyMusicBox.ListCount = 0, False, True)
        mnuDND.Enabled = IIf(MyPlayer(0).URL = vbNullString And MyPlayer(1).URL = vbNullString, False, True)
        mnuGaming.Enabled = IIf(MyPlayer(0).URL = vbNullString And MyPlayer(1).URL = vbNullString, False, True)
    End If
End Sub

Private Sub HideBtn_Click() 'Hiding the extra buttons when sending music to device
    CancelBtn.Visible = False
    HideBtn.Visible = False
End Sub

Private Sub HotKeyOCG_HotkeyPressed(Index As Integer)
    On Error Resume Next
    If HotKeyOCG(Index).Enabled Then
        Select Case Index
            Case 0
                Unload AboutWin
                Unload RenameMusic
                Unload SendMusic
                Unload SkinnedOCG
                Unload SkinnedOCG2
                Unload OneClickGo
            Case 1
                If mnuGaming.Checked Then
                    On Error Resume Next
                    For Each MyControl In Controls
                        If InStr(MyControl.Tag, "E") <> 0 Then
                            MyControl.Enabled = True
                        End If
                        If InStr(MyControl.Tag, "V") <> 0 Then
                            MyControl.Visible = True
                        End If
                        If InStr(MyControl.Tag, "A") <> 0 Then
                            MyControl.AutoRedraw = True
                        End If
                    Next
                    Enabled = True
                    AutoRedraw = True
                    Show
                    If MyPlayer(0).URL <> vbNullString Or MyPlayer(1).URL <> vbNullString Then
                        RecentAction.Caption = "Now Playing : " & IIf(PlayingAll, AllMusicBox.List(NowPlaying), MyMusicBox.List(NowPlaying))
                        If Not mnuDND.Checked Then
                            If PlayingAll Then
                                If AllMusicBox.ListIndex <> NowPlaying Then
                                    IsContinue = True
                                    AllMusicBox.ListIndex = NowPlaying
                                End If
                            Else
                                If MyMusicBox.ListIndex <> NowPlaying Then
                                    IsContinue = True
                                    MyMusicBox.ListIndex = NowPlaying
                                End If
                            End If
                        End If
                    End If
                    mnuGaming.Checked = False
                    If TransferFrom = 1 Then
                        MyMusicBox.SetFocus
                    Else
                        AllMusicBox.SetFocus
                    End If
                Else
                    If CanRestore Then
                        Unload SkinnedOCG
                        Unload SkinnedOCG2
                        mnuSkinMode.Checked = False
                        Show
                        WindowState = IIf(IsMaxWin, 2, 0)
                    End If
                End If
            Case 2
                mnuPlay_Click
                If mnuSkinMode.Checked Then
                    If mnuPlay.Caption = "&Pause" Then
                        If mnuRockstarGold.Checked Then
                            SkinnedOCG.PlayPauseImg.Picture = LoadPicture(OfName("PlayPauseADefault"))
                        Else
                            SkinnedOCG2.ImgBack.Picture = LoadPicture(OfName("PauseDefault"))
                        End If
                    Else
                        If mnuRockstarGold.Checked Then
                            SkinnedOCG.PlayPauseImg.Picture = LoadPicture(OfName("PlayPauseLDefault"))
                        Else
                            SkinnedOCG2.ImgBack.Picture = LoadPicture(OfName("PlayDefault"))
                        End If
                    End If
                End If
            Case 3
                mnuStop_Click
                If mnuSkinMode.Checked Then
                    If mnuRockstarGold.Checked Then
                        SkinnedOCG.PlayPauseImg.Picture = LoadPicture(OfName("PlayPauseLDefault"))
                    Else
                        SkinnedOCG2.ImgBack.Picture = LoadPicture(OfName("PlayDefault"))
                    End If
                End If
            Case 4
                mnuPrevSong_Click
                If mnuSkinMode.Checked Then
                    If mnuRockstarGold.Checked Then
                        SkinnedOCG.PlayPauseImg.Picture = LoadPicture(OfName("PlayPauseADefault"))
                        SkinnedOCG.NPSVvanisher.Enabled = True
                    Else
                        SkinnedOCG2.ImgBack.Picture = LoadPicture(OfName("PauseDefault"))
                        SkinnedOCG2.NPSVvanisher.Enabled = True
                    End If
                End If
            Case 5
                mnuNextSong_Click
                If mnuSkinMode.Checked Then
                    If mnuRockstarGold.Checked Then
                        SkinnedOCG.PlayPauseImg.Picture = LoadPicture(OfName("PlayPauseADefault"))
                        SkinnedOCG.NPSVvanisher.Enabled = True
                    Else
                        SkinnedOCG2.ImgBack.Picture = LoadPicture(OfName("PauseDefault"))
                        SkinnedOCG2.NPSVvanisher.Enabled = True
                    End If
                End If
        End Select
    End If
End Sub

Private Sub mnuAbout_Click()
    AboutWin.Show vbModal
End Sub

Private Sub mnuAutoRename_Click() 'Auto Rename Menu
    RenameMusic.Show vbModal
End Sub

Private Sub mnuBackup_Click() 'BackUp Menu
    TempStr2 = InputBox("Where do you want OCG to create back up? OCG will create a file BackUp.ocgb in this location and overwrite any existing file.", Replace(mnuBackup.Caption, "&", vbNullString), GtSt("BackUpDir", App.Path))
    If FileSystem.folderexists(TempStr2) Then
        SvSt "BackUpDir", TempStr2
        TempStr2 = "REG EXPORT " & Chr$(34) & "HKCU\Software\VB and VBA Program Settings\ANIco.in\OneClick Go!" & Chr$(34) & " " & Chr$(34) & TempStr2 & "\BackUp.ocgb" & Chr$(34) & " /y"
        Shell TempStr2, vbHide
        RecentAction.Caption = "Playlist and settings backed up"
    Else
        If TempStr2 = vbNullString Then
            RecentAction.Caption = "Back up aborted by user"
        Else
            RecentAction.Caption = "Back up failed: Invalid directory input"
        End If
    End If
End Sub

Private Sub mnuBcentre_Click()
    MyPlayer(0).settings.balance = 0
    MyPlayer(1).settings.balance = 0
    mnuBalance.Caption = "&Balance (C)"
End Sub

Private Sub mnuBleft_Click()
    MyPlayer(0).settings.balance = MyPlayer(0).settings.balance - 1
    MyPlayer(1).settings.balance = MyPlayer(1).settings.balance - 1
    If MyPlayer(0).settings.balance < 0 Then
        mnuBalance.Caption = "&Balance (L)"
    End If
    If MyPlayer(0).settings.balance = 0 Then
        mnuBalance.Caption = "&Balance (C)"
    End If
End Sub

Private Sub mnuBright_Click()
    MyPlayer(0).settings.balance = MyPlayer(0).settings.balance + 1
    MyPlayer(1).settings.balance = MyPlayer(1).settings.balance + 1
    If MyPlayer(0).settings.balance > 0 Then
        mnuBalance.Caption = "&Balance (R)"
    End If
    If MyPlayer(0).settings.balance = 0 Then
        mnuBalance.Caption = "&Balance (C)"
    End If
End Sub

Private Sub mnuPSfaster_Click()
    If PlaySpd < 800 Then
        PlaySpd = PlaySpd + 5
    End If
    mnuPlaySpeed.Caption = "Pl&ay Speed (" & Trim$(Str$(PlaySpd)) & "%)"
End Sub

Private Sub mnuPSreset_Click()
    PlaySpd = 100
    mnuPlaySpeed.Caption = "Pl&ay Speed (" & Trim$(Str$(PlaySpd)) & "%)"
End Sub

Private Sub mnuPSslower_Click()
    If PlaySpd > 50 Then
        PlaySpd = PlaySpd - 5
    End If
    mnuPlaySpeed.Caption = "Pl&ay Speed (" & Trim$(Str$(PlaySpd)) & "%)"
End Sub

Private Sub mnuReset_Click() 'Reset Menu
    MyMsgBox = MsgBox("Are you sure you want to reset program settings and change them to default? The process will also delete all information about your music chart i.e. your playlist and automatically restart with default settings. If you are unsure, click 'No' else click 'Yes' to continue.", vbYesNo + vbCritical + vbDefaultButton2, Replace(mnuReset.Caption, "&", vbNullString))
    If MyMsgBox = vbYes Then
        DeleteSetting "ANIco.in", "OneClick Go!"
        DeleteSetting "ANIco.in", "SongCache"
        SvSt "AllowRun", "1"
        Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
        End
    Else
        RecentAction.Caption = "Reset aborted by user"
    End If
End Sub

Private Sub mnuRestore_Click() 'Restore Menu
    TempStr2 = InputBox("Where did you save OCG back up file BackUp.ocgb? OCG will restore it's playlist and settings from that file and overwrite any existing setting and playlist. The program will restart automatically.", Replace(mnuRestore.Caption, "&", vbNullString), GtSt("BackUpDir", App.Path))
    If FileSystem.FileExists(TempStr2 & "\BackUp.ocgb") Then
        SvSt "BackUpDir", TempStr2
        TempStr2 = "REG IMPORT " & Chr$(34) & TempStr2 & "\BackUp.ocgb" & Chr$(34)
        Shell TempStr2, vbHide
        SvSt "AllowRun", "1"
        Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
        End
    Else
        If TempStr2 = vbNullString Then
            RecentAction.Caption = "Restore aborted by user"
        Else
            RecentAction.Caption = "Restore failed: Invalid directory input"
        End If
    End If
End Sub

Private Sub mnuGaming_Click() 'Code for Gaming Mode
    If GtSt("GamingIntro", "0") = "0" Then
        MyMsgBox = MsgBox("If you enter Gaming Mode, OCG consumes lesser resources than it does generally. Remember, that you can restore the window again by pressing Ctrl+F8 and can end the program by pressing Ctrl+F7. It's best if you are about to play a heavy game or even if you are about to initialize a heavy application. You can still use hotkeys to change songs.", vbInformation, "Understanding Gaming Mode")
        SvSt "GamingIntro", "1"
    End If
    mnuGaming.Checked = True
    On Error Resume Next
    For Each MyControl In Controls
        MyControl.Tag = vbNullString
        If MyControl.Enabled Then
            MyControl.Enabled = False
            MyControl.Tag = "E"
        End If
        If MyControl.Visible Then
            MyControl.Visible = False
            MyControl.Tag = MyControl.Tag & "V"
        End If
        If MyControl.AutoRedraw Then
            MyControl.AutoRedraw = False
            MyControl.Tag = MyControl.Tag & "A"
        End If
    Next
    HotKeyOCG(0).Enabled = True
    HotKeyOCG(1).Enabled = True
    HotKeyOCG(2).Enabled = True
    HotKeyOCG(3).Enabled = True
    HotKeyOCG(4).Enabled = True
    HotKeyOCG(5).Enabled = True
    MainTimer.Enabled = True
    mnuMedia.Enabled = True
    Enabled = False
    AutoRedraw = False
    Hide
End Sub

Private Sub MyMusicBox_Click()
    MyBoxesClick MyMusicBox
End Sub

Private Sub MyMusicBox_DblClick()
    MyBoxesDblClick MyMusicBox
End Sub

Private Sub MyMusicBox_KeyUp(KeyCode As Integer, Shift As Integer)
    MyBoxesKeyUp MyMusicBox, KeyCode
End Sub

Private Sub MySearch_DblClick()
    If MySearch.ListIndex <> -1 Then
        MyMusicBox.ListIndex = Val(Mid$(MySearch.List(MySearch.ListIndex), InStrRev(MySearch.List(MySearch.ListIndex), "(") + 1, Len(MySearch.List(MySearch.ListIndex)) - InStrRev(MySearch.List(MySearch.ListIndex), "(") - 1)) - 1
    End If
    SearchMy.Text = vbNullString
    MyMusicBox.SetFocus
    SearchMy_LostFocus
End Sub

Private Sub MySearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp And MySearch.ListIndex = 0 Then
        SearchMy.SetFocus
    End If
End Sub

Private Sub MySearch_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        MySearch_DblClick
    End If
End Sub

Private Sub SearchAll_Change() 'Searching through the search boxes
    AllSearch.Enabled = True
    AllSearch.Visible = IIf(SearchAll.Text = vbNullString, False, True)
    AllSearch.Clear
    For TempNum = 0 To AllMusicBox.ListCount - 1
        If InStr(1, AllMusicBox.List(TempNum), SearchAll.Text, vbTextCompare) <> 0 Then
            AllSearch.AddItem AllMusicBox.List(TempNum) & " (" & LTrim$(Str$(TempNum + 1)) & ")"
        End If
    Next
    If AllSearch.ListCount >= 7 Then
        AllSearch.Height = 1500
    Else
        AllSearch.Height = 225 * AllSearch.ListCount
    End If
    If AllSearch.ListCount = 0 Then
        AllSearch.AddItem "No Search Results"
        AllSearch.Enabled = False
    End If
End Sub

Private Sub SearchAll_GotFocus()
    SearchAll.FontItalic = False
    If SearchAll.Text = "Search All Music" Then
        SearchAll.Text = vbNullString
    End If
    MySearch.Visible = False
    SearchAll.BackColor = AllMusicBox.BackColor
    SearchAll.SelStart = 0
    SearchAll.SelLength = Len(SearchAll.Text)
End Sub

Private Sub SearchAll_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        AllSearch.Visible = False
    End If
    If KeyCode = vbKeyDown And AllSearch.Visible And AllSearch.List(0) <> "No Search Results" Then
        AllSearch.SetFocus
        AllSearch.ListIndex = 0
    End If
End Sub

Private Sub SearchAll_LostFocus()
    If SearchAll.Text = vbNullString Then
        SearchAll.Text = "Search All Music"
        AllSearch.Visible = False
        SearchAll.FontItalic = True
    End If
    ColorCodeToRGB AllMusicBox.BackColor
    SearchAll.BackColor = RGB(IIf(cRed < 128, cRed + 16, cRed - 16), IIf(cGreen < 128, cGreen + 16, cGreen - 16), IIf(cBlue < 128, cBlue + 16, cBlue - 16))
End Sub

Private Sub SearchMy_Change()
    MySearch.Enabled = True
    MySearch.Visible = IIf(SearchMy.Text = vbNullString, False, True)
    MySearch.Clear
    For TempNum = 0 To MyMusicBox.ListCount - 1
        If InStr(1, MyMusicBox.List(TempNum), SearchMy.Text, vbTextCompare) <> 0 Then
            MySearch.AddItem MyMusicBox.List(TempNum) & " (" & LTrim$(Str$(TempNum + 1)) & ")"
        End If
    Next
    If MySearch.ListCount >= 7 Then
        MySearch.Height = 1500
    Else
        MySearch.Height = 225 * MySearch.ListCount
    End If
    If MySearch.ListCount = 0 Then
        MySearch.AddItem "No Search Results"
        MySearch.Enabled = False
    End If
End Sub

Private Sub SearchMy_GotFocus()
    SearchMy.FontItalic = False
    If SearchMy.Text = "Search My Music" Then
        SearchMy.Text = vbNullString
    End If
    AllSearch.Visible = False
    SearchMy.BackColor = MyMusicBox.BackColor
    SearchMy.SelStart = 0
    SearchMy.SelLength = Len(SearchMy.Text)
End Sub

Private Sub SearchMy_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        MySearch.Visible = False
    End If
    If KeyCode = vbKeyDown And MySearch.Visible And MySearch.List(0) <> "No Search Results" Then
        MySearch.SetFocus
        MySearch.ListIndex = 0
    End If
End Sub

Private Sub SearchMy_LostFocus()
    If SearchMy.Text = vbNullString Then
        SearchMy.Text = "Search My Music"
        MySearch.Visible = False
        SearchMy.FontItalic = True
    End If
    ColorCodeToRGB MyMusicBox.BackColor
    SearchMy.BackColor = RGB(IIf(cRed < 128, cRed + 16, cRed - 16), IIf(cGreen < 128, cGreen + 16, cGreen - 16), IIf(cBlue < 128, cBlue + 16, cBlue - 16))
End Sub

Private Sub SeekFX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MyPlayer(ActiveIndex).URL <> vbNullString Then
        SeekFX.ToolTipText = IIf(Int(MyPlayer(ActiveIndex).currentMedia.duration * X / (60 * SeekFX.Width)) < 10, "0", vbNullString) & LTrim$(Str$(Int(MyPlayer(ActiveIndex).currentMedia.duration * X / (60 * SeekFX.Width)))) & ":" & IIf(Int(MyPlayer(ActiveIndex).currentMedia.duration * X / SeekFX.Width) Mod 60 < 10, "0", vbNullString) & LTrim$(Str$(Int(MyPlayer(ActiveIndex).currentMedia.duration * X / SeekFX.Width) Mod 60)) & "/" & LTrim$(MyPlayer(ActiveIndex).currentMedia.durationString)
    End If
End Sub

Public Sub SeekFX_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > 0 And X < SeekFX.Width And MyPlayer(ActiveIndex).URL <> vbNullString Then
        SeekBack.Width = X
        MyPlayer(ActiveIndex).Controls.currentPosition = MyPlayer(ActiveIndex).currentMedia.duration * X / SeekFX.Width
        If Not ASradio Then
            MyPlayer(ActiveIndex).settings.volume = Val(GtSt("Volume", "67"))
        End If
        If CFtimer.Enabled Then
            If Not mnuRepeat.Checked Or mnuShuffle.Checked Then
                CurrentSong = ActiveCS - 1
            End If
            If mnuRepeat.Checked And Not mnuShuffle.Checked Then
                AutoRepeated = True
            End If
            mnuNextSong_Click
            TimesCount = 0
            CFtimer.Enabled = False
            PrevRectify = False
            ClickedNext = False
            MyPlayer(PassiveIndex).URL = vbNullString
            Reported = False
        End If
    End If
End Sub

Private Sub mnuBriefHistory_Click() 'Breif History menu
    If Len(BriefHistory) > 850 Then
        BriefHistory = Mid$(BriefHistory, 1, InStrRev(BriefHistory, Chr$(187)) - 1)
    End If
    If BriefHistory <> vbNullString Then
        MyMsgBox = MsgBox("BRIEF HISTORY" & Chr$(13) & String(30, "—") & Chr$(13) & Chr$(13) & BriefHistory, vbInformation, "Brief History")
    Else
        MyMsgBox = MsgBox("No major action performed in this session.", vbInformation, "Brief History")
    End If
End Sub

Private Sub mnuColor_Click() 'Customize Display menu
    mnuColor.Checked = IIf(mnuColor.Checked, False, True)
    ColorSchemeBox.Visible = mnuColor.Checked
    Form_Resize
End Sub

Private Sub mnuDeleteFile_Click() 'Delete Song menu
    If TransferFrom = 1 Then
        OldName = MyMusicBox.List(MyMusicBox.ListIndex)
        MyMsgBox = MsgBox("Are you sure you want to delete this song permanently?" & Chr$(13) & Chr$(13) & Trim$(OldName), vbQuestion + vbYesNo + vbDefaultButton2, "Song Deletion")
        If MyMsgBox = vbNo Then
            Exit Sub
        End If
        If MyPlayer(ActiveIndex).URL = MyMusicFolder & "\" & Trim$(OldName) & ".mp3" Then
            MyPlayer(ActiveIndex).URL = vbNullString
        End If
        FileSystem.deletefile MyMusicFolder & "\" & Trim$(OldName) & ".mp3", True
        MyMusicBox.RemoveItem MyMusicBox.ListIndex
        mnuPlay.Caption = "&Play"
        SeekBack.Width = 0
        RecentAction.Caption = "Junk Song Removed : " & Trim$(OldName)
        BriefHistory = Chr$(187) & " " & Str$(Now) & Chr$(13) & RecentAction.Caption & Chr$(13) & Chr$(13) & BriefHistory
    Else
        OldName = AllMusicBox.List(AllMusicBox.ListIndex)
        MyMsgBox = MsgBox("Are you sure you want to delete this song permanently?" & Chr$(13) & Chr$(13) & Trim$(OldName), vbQuestion + vbYesNo + vbDefaultButton2, "Song Deletion")
        If MyMsgBox = vbNo Then
            Exit Sub
        End If
        If MyPlayer(ActiveIndex).URL = AllMusicFolder & "\" & Trim$(OldName) & ".mp3" Then
            MyPlayer(ActiveIndex).URL = vbNullString
        End If
        FileSystem.deletefile AllMusicFolder & "\" & Trim$(OldName) & ".mp3", True
        AllMusicBox.RemoveItem AllMusicBox.ListIndex
        mnuPlay.Caption = "&Play"
        SeekBack.Width = 0
        RecentAction.Caption = "Junk Song Removed : " & Trim$(OldName)
        BriefHistory = Chr$(187) & " " & Str$(Now) & Chr$(13) & RecentAction.Caption & Chr$(13) & Chr$(13) & BriefHistory
    End If
End Sub

Private Sub mnuDND_Click() 'Do Not Disturb menu
    mnuDND.Checked = IIf(mnuDND.Checked, False, True)
    RecentAction.Caption = "Do Not Disturb mode turned " & IIf(mnuDND.Checked, "on", "off")
End Sub

Private Sub mnuDown_Click() 'Move Down(Decrease Rating) menu
    OldName = MyMusicBox.List(MyMusicBox.ListIndex)
    MyMusicBox.List(MyMusicBox.ListIndex) = MyMusicBox.List(MyMusicBox.ListIndex + 1)
    MyMusicBox.List(MyMusicBox.ListIndex + 1) = OldName
    IsContinue = True
    MyMusicBox.ListIndex = MyMusicBox.ListIndex + 1
    NowPlaying = NowPlaying + 1
    If MyMusicBox.ListIndex = MyMusicBox.ListCount - 1 Then
        mnuDown.Enabled = False
    End If
    mnuUp.Enabled = True
End Sub

Private Sub mnuExit_Click() 'Exit Menu
    Unload Me
    Unload SkinnedOCG
    Unload SkinnedOCG2
End Sub

Private Sub mnuFindSong_Click() 'Find song menu
    If TransferFrom = 0 Then
        SearchAll.SetFocus
    Else
        SearchMy.SetFocus
    End If
End Sub

Private Sub mnuFolderChange_Click() 'Change Folders and reload menu
    SvSt "UsedBefore", "0"
    mnuFolderChange.Checked = True
    SvSt "CantGame", "1"
    Unload Me
    Show
    RecentAction.Caption = "Welcome to OneClick Go! - Loaded successfully"
End Sub

Private Sub mnuFreezedBlue_Click() 'Change skin to Freezed Blue
    mnuFreezedBlue.Checked = True
    mnuRockstarGold.Checked = False
    SvSt "CurrentSkin", "1"
End Sub

Private Sub mnuHit_Click() 'Copy Music Chart menu
    Dim HitList As Long
    HitList = Val(InputBox("How many songs are to be selected for your Music Chart?", "Share Music Chart", IIf(MyMusicBox.ListCount >= 10, "10", LTrim$(MyMusicBox.ListCount))))
    If (HitList <= MyMusicBox.ListCount And HitList > 0) Then
        OldName = vbNullString
        Clipboard.Clear
        For FilNum = 0 To IIf(MyMusicBox.ListCount > HitList - 1, HitList - 1, MyMusicBox.ListCount - 1)
            OldName = OldName & "#" & LTrim$(Str$(FilNum + 1)) & ". " & MyMusicBox.List(FilNum) & Chr$(13)
        Next
        Clipboard.SetText OldName
        RecentAction.Caption = "Your Music Chart is ready to be shared. Paste (Ctrl+V) to use in any textbox"
    Else
        RecentAction.Caption = "Copying Music Chart cancelled by user or user made an invalid entry"
    End If
End Sub

Private Sub mnuLicense_Click() 'About Menu
    ShellExecute 0, "OPEN", App.Path & "\Documents\END USER LICENSE AGREEMENT.rtf", vbNullString, App.Path & "\Documents\", 1
End Sub

Private Sub mnuLinear_Click() 'Play songs linearly menu
    mnuShuffle.Checked = False
    mnuLinRev.Checked = False
    mnuLinear.Checked = True
    mnuLinear.Enabled = False
    mnuShuffle.Enabled = True
    mnuLinRev.Enabled = True
    SongCacheReset
End Sub

Private Sub mnuLinRev_Click() 'Play songs reverse to the linear order
    mnuShuffle.Checked = False
    mnuLinear.Checked = False
    mnuLinRev.Checked = True
    mnuLinear.Enabled = True
    mnuShuffle.Enabled = True
    mnuLinRev.Enabled = False
    SongCacheReset
End Sub

Private Sub mnuMusicLeft_Click() 'Show My Music on Left menu
    mnuMusicLeft.Checked = IIf(mnuMusicLeft.Checked, False, True)
    SvSt "MusicLeft", mnuMusicLeft.Checked
    Form_Resize
End Sub

Public Sub mnuMute_Click() 'Mute menu
    mnuMute.Checked = IIf(mnuMute.Checked, False, True)
    If mnuMute.Checked Then
        MyPlayer(ActiveIndex).settings.mute = True
        MyPlayer(PassiveIndex).settings.mute = True
        RecentAction.Caption = "Volume - " & GtSt("Volume", "67") & " (Muted)"
    Else
        MyPlayer(ActiveIndex).settings.mute = False
        MyPlayer(PassiveIndex).settings.mute = False
        RecentAction.Caption = "Volume - " & GtSt("Volume", "67")
    End If
    SvSt "Mute", mnuMute.Checked
End Sub

Public Sub mnuNextSong_Click() 'The algorithm for Next Song is quite mathematical and logical
    'Managing CurrentSong
    If Not mnuRepeat.Checked Then
        CurrentSong = IIf(mnuPrevSong.Checked, CurrentSong - IIf(CFtimer.Enabled, 2, 1), CurrentSong + 1)
    Else
        If Not AutoRepeated Then
            CurrentSong = IIf(mnuPrevSong.Checked, CurrentSong - 1, CurrentSong + 1)
        End If
    End If
    If PrevRectify And CFtimer.Enabled And mnuPrevSong.Checked And Not mnuRepeat.Checked Then
        CurrentSong = CurrentSong + 1
    Else
        If mnuPrevSong.Checked Then
            PrevRectify = True
        End If
    End If
    If Not CFtimer.Enabled Then
        ActiveCS = CurrentSong
    End If
    
    'Calculating the next song
    If Not AutoRepeated Then
        If GtSt2(Str$(CurrentSong), vbNullString) = vbNullString Then
            Do
                If mnuShuffle.Checked Then
                    NowPlaying = Round(Rnd * (IIf(PlayingAll, AllMusicBox.ListCount, MyMusicBox.ListCount) - 1), 0)
                Else
                    NowPlaying = IIf((mnuLinear.Checked And mnuPrevSong.Checked = False) Or (mnuLinRev.Checked And mnuPrevSong.Checked), (NowPlaying + 1) Mod IIf(PlayingAll, AllMusicBox.ListCount, MyMusicBox.ListCount), (IIf(PlayingAll, AllMusicBox.ListCount, MyMusicBox.ListCount) - 1) - (((IIf(PlayingAll, AllMusicBox.ListCount, MyMusicBox.ListCount) + 1) - (NowPlaying + 1)) Mod IIf(PlayingAll, AllMusicBox.ListCount, MyMusicBox.ListCount)))
                End If
                If IIf(PlayingAll, AllMusicBox.ListCount, MyMusicBox.ListCount) = 1 Then
                    Exit Do
                End If
            Loop While NowPlaying = Val(GtSt2(Str$(IIf(mnuPrevSong.Checked, CurrentSong + 1, CurrentSong - 1))))
            SvSt2 Str$(CurrentSong), Str$(NowPlaying)
        Else
            NowPlaying = Val(GtSt2(Str$(CurrentSong)))
        End If
    Else
        AutoRepeated = False
    End If
    MyPlayer(IIf(CFtimer.Enabled, PassiveIndex, ActiveIndex)).URL = IIf(PlayingAll, AllMusicFolder, MyMusicFolder) & "\" & IIf(PlayingAll, AllMusicBox.List(NowPlaying), MyMusicBox.List(NowPlaying)) & ".mp3"
    If Not mnuGaming.Checked And Not CFtimer.Enabled Then
        RecentAction.Caption = "Now Playing : " & IIf(PlayingAll, AllMusicBox.List(NowPlaying), MyMusicBox.List(NowPlaying))
        If Not mnuDND.Checked Then
            IsContinue = True
            If PlayingAll Then
                AllMusicBox.ListIndex = NowPlaying
            Else
                MyMusicBox.ListIndex = NowPlaying
            End If
        End If
    End If
    mnuPrevSong.Checked = False
    
    'Sending the information to skins
    If mnuSkinMode.Checked Then
        If mnuRockstarGold.Checked Then
            Refresher
        Else
            Refresher2
        End If
    End If
End Sub

Private Sub mnuPlay_Click() 'Play menu
    If MyPlayer(ActiveIndex).URL <> vbNullString Or MyPlayer(PassiveIndex).URL <> vbNullString Then
        If mnuPlay.Caption = "&Play" Then
            mnuPlay.Caption = "&Pause"
            MyPlayer(ActiveIndex).Controls.play
            If MyPlayer(ActiveIndex).URL <> vbNullString And MyPlayer(PassiveIndex).URL <> vbNullString Then
                CFtimer.Enabled = True
                MyPlayer(PassiveIndex).Controls.play
                RecentAction.Caption = "Multiple Media Resumed"
            Else
                RecentAction.Caption = "Now Playing : " & IIf(PlayingAll, AllMusicBox.List(NowPlaying), MyMusicBox.List(NowPlaying))
            End If
        Else
            mnuPlay.Caption = "&Play"
            MyPlayer(ActiveIndex).Controls.pause
            If MyPlayer(ActiveIndex).URL <> vbNullString And MyPlayer(PassiveIndex).URL <> vbNullString Then
                CFtimer.Enabled = False
                MyPlayer(PassiveIndex).Controls.pause
                RecentAction.Caption = "Multiple Media Paused"
            Else
                RecentAction.Caption = "Media Paused : " & IIf(PlayingAll, AllMusicBox.List(NowPlaying), MyMusicBox.List(NowPlaying))
            End If
        End If
    End If
End Sub

Private Sub mnuPrevSong_Click() 'Previous Song Menu
    mnuPrevSong.Checked = True
    mnuNextSong_Click 'Calling everything of the next song but in reverse order
End Sub

Private Sub mnuRating_Click() 'Custom Rank menu
    FilNum = Val(InputBox("Enter the desired rank of this song", "OneClick Go!", Trim$(Str$(MyMusicBox.ListIndex + 1))))
    If FilNum <= MyMusicBox.ListCount And FilNum > 0 And FilNum <> MyMusicBox.ListIndex + 1 Then
        If Not mnuDND.Checked Then
            NowPlaying = FilNum - 1
        Else
            If FilNum - 1 < NowPlaying And MyMusicBox.ListIndex > NowPlaying Then
                NowPlaying = NowPlaying + 1
            End If
            If FilNum - 1 > NowPlaying And MyMusicBox.ListIndex < NowPlaying Then
                NowPlaying = NowPlaying - 1
            End If
        End If
        OldName = MyMusicBox.List(MyMusicBox.ListIndex)
        MyMusicBox.RemoveItem MyMusicBox.ListIndex
        MyMusicBox.AddItem OldName, FilNum - 1
    End If
End Sub

Private Sub mnuRenameFile_Click() 'Rename song menu
    If TransferFrom = 1 Then
        OldName = MyMusicBox.List(MyMusicBox.ListIndex)
        NewName = InputBox("Give a new name for the song" & Chr$(13) & Chr$(13) & OldName, "Rename Song", OldName)
        If NewName = vbNullString Or NewName = OldName Then
            Exit Sub
        End If
        If FileSystem.FileExists(MyMusicFolder & "\" & NewName & ".mp3") Then
            TempNum = 1
            Do
                TempNum = TempNum + 1
            Loop While FileSystem.FileExists(MyMusicFolder & "\" & NewName & Str$(TempNum) & ".mp3")
            NewName = NewName & Str$(TempNum)
        End If
        Name MyMusicFolder & "\" & OldName & ".mp3" As MyMusicFolder & "\" & NewName & ".mp3"
        MyMusicBox.List(MyMusicBox.ListIndex) = NewName
        SvSt Str$(MyMusicBox.ListIndex + 1), NewName
    Else
        OldName = AllMusicBox.List(AllMusicBox.ListIndex)
        NewName = InputBox("Give a new name for the song" & Chr$(13) & Chr$(13) & OldName, "Rename Song", OldName)
        If NewName = vbNullString Or NewName = OldName Then
            Exit Sub
        End If
        If FileSystem.FileExists(AllMusicFolder & "\" & NewName & ".mp3") Then
            TempNum = 1
            Do
                TempNum = TempNum + 1
            Loop While FileSystem.FileExists(AllMusicFolder & "\" & NewName & Str$(TempNum) & ".mp3")
            NewName = NewName & Str$(TempNum)
        End If
        Name AllMusicFolder & "\" & OldName & ".mp3" As AllMusicFolder & "\" & NewName & ".mp3"
        AllMusicBox.List(AllMusicBox.ListIndex) = NewName
    End If
    RecentAction.Caption = "Song Renamed : " & OldName
    BriefHistory = Chr$(187) & " " & Str$(Now) & Chr$(13) & RecentAction.Caption & Chr$(13) & Chr$(13) & BriefHistory
End Sub

Private Sub mnuRepeat_Click() 'Auto-Repeat menu
    mnuRepeat.Checked = IIf(mnuRepeat.Checked, False, True)
    SvSt "Repeat", mnuRepeat.Checked
End Sub

Private Sub mnuCrossfade_Click() 'Crossfade menu
    mnuCrossfade.Checked = IIf(mnuCrossfade.Checked, False, True)
    SvSt "Crossfade", mnuCrossfade.Checked
End Sub

Private Sub mnuRockstarGold_Click() 'Select Rockstar Gold skin
    mnuRockstarGold.Checked = True
    mnuFreezedBlue.Checked = False
    SvSt "CurrentSkin", "0"
End Sub

Private Sub mnuSave_Click() 'Save playlist menu
    For FilNum = 1 To MyMusicBox.ListCount
        SvSt Str$(FilNum), MyMusicBox.List(FilNum - 1)
    Next
    FilNum = MyMusicBox.ListCount + 1
    Do
        If GtSt(Str$(FilNum)) <> vbNullString Then
            DlSt Str$(FilNum)
        Else
            Exit Do
        End If
        FilNum = FilNum + 1
    Loop
    If GtSt("MaxCount", MyMusicBox.ListCount) < MyMusicBox.ListCount Then
        SvSt "MaxCount", MyMusicBox.ListCount
    End If
    RecentAction.Caption = "My Music Saved at " & Str$(Now)
    BriefHistory = Chr$(187) & " " & Str$(Now) & Chr$(13) & RecentAction.Caption & Chr$(13) & Chr$(13) & BriefHistory
End Sub

Private Sub mnuSend_Click() 'Send Music to device menu
    If mnuSend.Caption = "&Send Music to Device" Then
        SendMusic.Show
        Hide
        mnuSend.Caption = "Cancel (Pause) &Sending Music"
    Else
        Unload SendMusic
        CancelBtn.Visible = False
        HideBtn.Visible = False
        mnuSend.Caption = "&Send Music to Device"
    End If
End Sub

Private Sub mnuShuffle_Click() 'Shuffle menu
    mnuShuffle.Checked = True
    mnuLinear.Checked = False
    mnuLinRev.Checked = False
    mnuLinear.Enabled = True
    mnuShuffle.Enabled = False
    mnuLinRev.Enabled = True
    SongCacheReset
End Sub

Private Sub mnuSkinMode_Click() 'Switch to skin mode menu
    mnuSkinMode.Checked = True
    If mnuRockstarGold.Checked Then
        SkinnedOCG.Show
    Else
        SkinnedOCG2.Show
    End If
    Hide
End Sub

Private Sub mnuStart_Click() 'Auto-Start menu
    mnuStart.Checked = IIf(mnuStart.Checked, False, True)
    RecentAction.Caption = "Auto Start " & IIf(mnuStart.Checked, "enabled", "disabled")
    SvSt "AutoStart", mnuStart.Checked
End Sub

Private Sub mnuStop_Click() 'Stop menu
    If MyPlayer(ActiveIndex).URL <> vbNullString Then
        MyPlayer(ActiveIndex).Controls.currentPosition = 0
        MyPlayer(ActiveIndex).Controls.pause
        mnuPlay.Caption = "&Play"
        If CFtimer.Enabled Then
            If Not mnuRepeat.Checked Or mnuShuffle.Checked Then
                CurrentSong = ActiveCS - 1
            End If
            If mnuRepeat.Checked And Not mnuShuffle.Checked Then
                AutoRepeated = True
            End If
            mnuNextSong_Click
            TimesCount = 0
            CFtimer.Enabled = False
            PrevRectify = False
            ClickedNext = False
            MyPlayer(PassiveIndex).URL = vbNullString
            Reported = False
            MyPlayer(ActiveIndex).Controls.stop
            MyPlayer(ActiveIndex).settings.volume = Val(GtSt("Volume", "67"))
        End If
        If Not mnuGaming.Checked Then
            SeekBack.Width = 0
            RecentAction.Caption = "Media Stopped : " & IIf(PlayingAll, AllMusicBox.List(NowPlaying), MyMusicBox.List(NowPlaying))
        End If
    End If
End Sub

Private Sub mnuSupport_Click() 'Support Menu
    ShellExecute 0, "OPEN", App.Path & "\Documents\Help.pdf", vbNullString, App.Path & "\Documents\", 1
End Sub

Private Sub mnuTransferNow_Click() 'Transfer now menu
    If TransferFrom = 1 Then
        MyBoxesDblClick MyMusicBox
    Else
        MyBoxesDblClick AllMusicBox
    End If
End Sub

Private Sub mnuUp_Click() 'Move Up(Increase Rating) menu
    OldName = MyMusicBox.List(MyMusicBox.ListIndex)
    MyMusicBox.List(MyMusicBox.ListIndex) = MyMusicBox.List(MyMusicBox.ListIndex - 1)
    MyMusicBox.List(MyMusicBox.ListIndex - 1) = OldName
    IsContinue = True
    MyMusicBox.ListIndex = MyMusicBox.ListIndex - 1
    NowPlaying = NowPlaying - 1
    If MyMusicBox.ListIndex = 0 Then
        mnuUp.Enabled = False
    End If
    mnuDown.Enabled = True
End Sub

Public Sub mnuVolumeDown_Click() 'Volume Down
    mnuVolumeUp.Enabled = True
    mnuMute.Checked = False
    SvSt "Mute", mnuMute.Checked
    SvSt "Volume", Val(GtSt("Volume", "67")) - 1
    MyPlayer(ActiveIndex).settings.mute = False
    MyPlayer(PassiveIndex).settings.mute = False
    If Val(GtSt("Volume", "67")) = 0 Then
        mnuVolumeDown.Enabled = False
    End If
    RecentAction.Caption = "Volume - " & GtSt("Volume", "67")
    If Not CFtimer.Enabled Then
        MyPlayer(ActiveIndex).settings.volume = Val(GtSt("Volume", "67"))
    End If
End Sub

Public Sub mnuVolumeUp_Click() 'Volume Up
    mnuVolumeDown.Enabled = True
    mnuMute.Checked = False
    SvSt "Mute", mnuMute.Checked
    SvSt "Volume", Val(GtSt("Volume", "67")) + 1
    MyPlayer(ActiveIndex).settings.mute = False
    MyPlayer(PassiveIndex).settings.mute = False
    If Val(GtSt("Volume", "67")) = 100 Then
        mnuVolumeUp.Enabled = False
    End If
    RecentAction.Caption = "Volume - " & GtSt("Volume", "67")
    If Not CFtimer.Enabled Then
        MyPlayer(ActiveIndex).settings.volume = Val(GtSt("Volume", "67"))
    End If
End Sub

Private Sub ColorScheme(Col0 As Long, Col1 As Long, Col2 As Long, Col3 As Long, Col4 As Long, Col5 As Long, Col6 As Long, Col7 As Long) 'Interface loading procedure
    BackColor = Col0
    ColorSchemeBox.BackColor = Col0
    CCSC(0).BackColor = Col0
    AllMusicBoxLabel.ForeColor = Col1
    MyMusicBoxLabel.ForeColor = Col1
    ColorSchemeBox.ForeColor = Col1
    CCS1.ForeColor = Col1
    CCS2.ForeColor = Col1
    CCS3.ForeColor = Col1
    CCSC(1).BackColor = Col1
    CCSF(0).ForeColor = Col1
    CCSF(1).ForeColor = Col1
    CCSF(2).ForeColor = Col1
    CCSF(3).ForeColor = Col1
    AllMusicBox.BackColor = Col2
    AllSearch.BackColor = Col2
    SearchAll.BackColor = Col2
    CCSC(2).BackColor = Col2
    AllMusicBox.ForeColor = Col3
    AllSearch.ForeColor = Col3
    SearchAll.ForeColor = Col3
    CCSC(3).BackColor = Col3
    MyMusicBox.BackColor = Col4
    MySearch.BackColor = Col4
    SearchMy.BackColor = Col4
    CCSC(4).BackColor = Col4
    MyMusicBox.ForeColor = Col5
    MySearch.ForeColor = Col5
    SearchMy.ForeColor = Col5
    CCSC(5).BackColor = Col5
    StatusBar.BackColor = Col6
    CCSC(6).BackColor = Col6
    RecentAction.ForeColor = Col7
    CCSC(7).BackColor = Col7
    ColorCodeToRGB StatusBar.BackColor
    SeekBack.BackColor = RGB(IIf(cRed < 128, cRed + 16, cRed - 16), IIf(cGreen < 128, cGreen + 16, cGreen - 16), IIf(cBlue < 128, cBlue + 16, cBlue - 16))
    ColorCodeToRGB AllMusicBox.BackColor
    SearchAll.BackColor = RGB(IIf(cRed < 128, cRed + 16, cRed - 16), IIf(cGreen < 128, cGreen + 16, cGreen - 16), IIf(cBlue < 128, cBlue + 16, cBlue - 16))
    ColorCodeToRGB MyMusicBox.BackColor
    SearchMy.BackColor = RGB(IIf(cRed < 128, cRed + 16, cRed - 16), IIf(cGreen < 128, cGreen + 16, cGreen - 16), IIf(cBlue < 128, cBlue + 16, cBlue - 16))
    NewFont AllMusicBoxLabel, GtSt("FN0", "MS Sans Serif"), GtSt("FB0", True), GtSt("FI0", False), Val(GtSt("FS0", "13"))
    NewFont MyMusicBoxLabel, GtSt("FN0", "MS Sans Serif"), GtSt("FB0", True), GtSt("FI0", False), Val(GtSt("FS0", "13"))
    NewFont AllMusicBox, GtSt("FN1", "MS Sans Serif"), GtSt("FB1", False), GtSt("FI1", False), Val(GtSt("FS1", "7"))
    NewFont MyMusicBox, GtSt("FN2", "MS Sans Serif"), GtSt("FB2", False), GtSt("FI2", False), Val(GtSt("FS2", "10"))
    NewFont RecentAction, GtSt("FN3", "MS Sans Serif"), GtSt("FB3", True), GtSt("FI3", False), Val(GtSt("FS3", "10"))
    NewFont CCSF(0), AllMusicBoxLabel.FontName, AllMusicBoxLabel.FontBold, AllMusicBoxLabel.FontItalic, AllMusicBoxLabel.FontSize
    NewFont ColorSchemeBox, AllMusicBoxLabel.FontName, AllMusicBoxLabel.FontBold, AllMusicBoxLabel.FontItalic, AllMusicBoxLabel.FontSize
    NewFont CCSF(1), AllMusicBox.FontName, AllMusicBox.FontBold, AllMusicBox.FontItalic, AllMusicBox.FontSize
    NewFont CCSF(2), MyMusicBox.FontName, MyMusicBox.FontBold, MyMusicBox.FontItalic, MyMusicBox.FontSize
    NewFont CCSF(3), RecentAction.FontName, RecentAction.FontBold, RecentAction.FontItalic, RecentAction.FontSize
End Sub

Private Sub MyBoxesClick(BoxName As ListBox) 'Clicking on a music box calls MyBoxesClick
    'Updating GUI
    AllSearch.Visible = False
    MySearch.Visible = False
    TransferFrom = IIf(BoxName.Name = "AllMusicBox", 0, 1)
    mnuTransferNow.Caption = IIf(TransferFrom = 0, "&Move to My Music" & Chr$(9) & "Space", "&Move to All Music" & Chr$(9) & "Space")
    If Not IsContinue And Not mnuDND.Checked Then
        SvSt2 "JustToClear", vbNullString
        DeleteSetting "ANIco.in", "SongCache"
        CurrentSong = 0
        SvSt2 Str$(CurrentSong), Str$(BoxName.ListIndex)
        If Not CFtimer.Enabled Then
            ActiveCS = CurrentSong
        End If
        NowPlaying = BoxName.ListIndex
        If Not ASradio Then
            MyPlayer(0).settings.volume = Val(GtSt("Volume", "67"))
        End If
        MyPlayer(1).settings.volume = 0
        MyPlayer(1).URL = vbNullString
        ActiveIndex = 0
        PassiveIndex = 1
        If MyPlayer(0).settings.volume = 100 Then
            mnuVolumeUp.Enabled = False
        End If
        If MyPlayer(0).settings.volume = 0 Then
            mnuVolumeDown.Enabled = False
        End If
        mnuMute.Checked = GtSt("Mute", False)
        MyPlayer(ActiveIndex).settings.mute = mnuMute.Checked
        MyPlayer(PassiveIndex).settings.mute = mnuMute.Checked
        If BoxName.Name = "AllMusicBox" Then
            MyPlayer(IIf(CFtimer.Enabled, PassiveIndex, ActiveIndex)).URL = AllMusicFolder & "\" & AllMusicBox.List(AllMusicBox.ListIndex) & ".mp3"
            PlayingAll = True
        Else
            MyPlayer(IIf(CFtimer.Enabled, PassiveIndex, ActiveIndex)).URL = MyMusicFolder & "\" & MyMusicBox.List(MyMusicBox.ListIndex) & ".mp3"
            PlayingAll = False
        End If
        mnuPlay.Caption = "&Pause"
    End If
    If Not mnuDND.Checked And Not mnuGaming.Checked And Not CFtimer.Enabled Then
        RecentAction.Caption = "Now Playing : " & BoxName.List(BoxName.ListIndex)
    End If
    IsContinue = False
    AllMusicBoxLabel.Caption = "  All Music (" & Trim$(Str$(AllMusicBox.ListIndex + 1) & "/" & Trim$(Str$(AllMusicBox.ListCount)) & ")")
    MyMusicBoxLabel.Caption = "  My Music (" & Trim$(Str$(MyMusicBox.ListIndex + 1) & "/" & Trim$(Str$(MyMusicBox.ListCount)) & ")")
End Sub

Private Sub MyBoxesDblClick(BoxName As ListBox) 'Double Clicking on a music box calls MyBoxesDblClick
    Dim TempN2 As Integer
    TempN2 = 0
    If BoxName.ListIndex <> -1 Then
        If BoxName.Name = "AllMusicBox" Then
            TempStr = InputBox("Enter the new rank of this song", "Adding New Song", Trim$(Str$(MyMusicBox.ListCount + 1)))
            If TempStr <> vbNullString And Val(TempStr) > 0 And Val(TempStr) <= MyMusicBox.ListCount + 1 And Val(TempStr) - Int(Val(TempStr)) = 0 Then
                If SendMessage(MyMusicBox.hWnd, &H18F, -1, ByVal AllMusicBox.List(AllMusicBox.ListIndex)) = -1 Then
                    MyMusicBox.AddItem AllMusicBox.List(AllMusicBox.ListIndex), Val(TempStr) - 1
                Else
                    TempN2 = 1
                    Do
                        TempN2 = TempN2 + 1
                    Loop While SendMessage(MyMusicBox.hWnd, &H18F, -1, ByVal AllMusicBox.List(AllMusicBox.ListIndex) & Str$(TempN2)) <> -1
                    MyMusicBox.AddItem AllMusicBox.List(AllMusicBox.ListIndex) & Str$(TempN2), Val(TempStr) - 1
                End If
            Else
                If TempStr <> vbNullString Then
                    RecentAction.Caption = "Song cannot be selected for My Music : Invalid Rank"
                End If
                Exit Sub
            End If
            FileSystem.movefile AllMusicFolder & "\" & AllMusicBox.List(AllMusicBox.ListIndex) & ".mp3", MyMusicFolder & "\" & MyMusicBox.List(Val((TempStr) - 1)) & ".mp3"
            OldName = "Song selected for My Music " & IIf(TempN2 <> 0, "and renamed ", vbNullString) & ": "
        Else
            If SendMessage(AllMusicBox.hWnd, &H18F, -1, ByVal MyMusicBox.List(MyMusicBox.ListIndex)) = -1 Then
                AllMusicBox.AddItem MyMusicBox.List(MyMusicBox.ListIndex)
                FileSystem.movefile MyMusicFolder & "\" & MyMusicBox.List(MyMusicBox.ListIndex) & ".mp3", AllMusicFolder & "\" & MyMusicBox.List(MyMusicBox.ListIndex) & ".mp3"
            Else
                TempN2 = 1
                Do
                    TempN2 = TempN2 + 1
                Loop While SendMessage(AllMusicBox.hWnd, &H18F, -1, ByVal MyMusicBox.List(MyMusicBox.ListIndex) & Str$(TempN2)) <> -1
                AllMusicBox.AddItem MyMusicBox.List(MyMusicBox.ListIndex) & Str$(TempN2)
                FileSystem.movefile MyMusicFolder & "\" & MyMusicBox.List(MyMusicBox.ListIndex) & ".mp3", AllMusicFolder & "\" & MyMusicBox.List(MyMusicBox.ListIndex) & Str$(TempN2) & ".mp3"
            End If
            OldName = "Song removed from My Music " & IIf(TempN2 <> 0, "and renamed ", vbNullString) & ": "
        End If
        RecentAction.Caption = OldName & BoxName.List(BoxName.ListIndex)
        BriefHistory = Chr$(187) & " " & Str$(Now) & Chr$(13) & RecentAction.Caption & Chr$(13) & Chr$(13) & BriefHistory
        BoxName.RemoveItem BoxName.ListIndex
    End If
    AllMusicBoxLabel.Caption = "  All Music (" & Trim$(Str$(AllMusicBox.ListIndex + 1) & "/" & Trim$(Str$(AllMusicBox.ListCount)) & ")")
    MyMusicBoxLabel.Caption = "  My Music (" & Trim$(Str$(MyMusicBox.ListIndex + 1) & "/" & Trim$(Str$(MyMusicBox.ListCount)) & ")")
End Sub

Private Sub MyBoxesKeyUp(BoxName As ListBox, KeyCode As Integer) 'Resolving the issue with making F10 as shortcut for Stop Button and making it more easier to change the music box of a song by pressing Space
    If KeyCode = vbKeySpace Then
        MyBoxesDblClick BoxName
    End If
    If KeyCode = vbKeyF10 And Not mnuSkinMode.Checked Then
        SendKeys "{ESC}", True
        mnuStop_Click
    End If
End Sub

Private Sub SongCacheReset()
    SvSt "Linear", mnuLinear.Checked
    SvSt "LinRev", mnuLinRev.Checked
    SvSt "Shuffle", mnuShuffle.Checked
    SvSt2 "JustToClear", vbNullString
    DeleteSetting "ANIco.in", "SongCache"
    CurrentSong = 0
    If Not CFtimer.Enabled Then
        ActiveCS = CurrentSong
    End If
    NowPlaying = IIf(PlayingAll, AllMusicBox.ListIndex, MyMusicBox.ListIndex)
    SvSt2 Str$(CurrentSong), Str$(NowPlaying)
End Sub
