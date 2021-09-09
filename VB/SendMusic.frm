VERSION 5.00
Begin VB.Form SendMusic 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Send Music to Device"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11730
   Icon            =   "SendMusic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   Begin VB.Timer VFXupdater 
      Interval        =   1
      Left            =   9480
      Top             =   7080
   End
   Begin VB.CheckBox MyCheck 
      Height          =   195
      Left            =   165
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   150
      Width           =   195
   End
   Begin VB.CheckBox AllCheck 
      Height          =   195
      Left            =   5940
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   150
      Width           =   195
   End
   Begin VB.CommandButton BrowseBtn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "..."
      Height          =   330
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   7110
      Width           =   375
   End
   Begin VB.CommandButton SendNow 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Send Now"
      Height          =   330
      Left            =   10305
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7110
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
      Height          =   315
      Left            =   8500
      TabIndex        =   9
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
      Height          =   315
      Left            =   2700
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Search My Music"
      Top             =   90
      Width           =   3015
   End
   Begin VB.ListBox AllSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   225
      Left            =   -10000
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   390
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ListBox MySearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   225
      Left            =   -10000
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   390
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ListBox MyMusicBox 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   6180
      IntegralHeight  =   0   'False
      ItemData        =   "SendMusic.frx":FCC9
      Left            =   5895
      List            =   "SendMusic.frx":FCCB
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   450
      Width           =   5655
   End
   Begin VB.ListBox AllMusicBox 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   6165
      IntegralHeight  =   0   'False
      ItemData        =   "SendMusic.frx":FCCD
      Left            =   120
      List            =   "SendMusic.frx":FCCF
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   450
      Width           =   5655
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
      TabIndex        =   3
      Top             =   7185
      UseMnemonic     =   0   'False
      Width           =   10095
   End
   Begin VB.Label MyMusicBoxLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "     My Music (0/0)"
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
      TabIndex        =   5
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label AllMusicBoxLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "     All Music (0/0)"
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
      TabIndex        =   4
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label StatusBar 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   7080
      Width           =   11415
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuSendNow 
         Caption         =   "&Send Now"
      End
      Begin VB.Menu mnuFindSong 
         Caption         =   "&Find Song"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuRenameFile 
         Caption         =   "&Rename Song"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuDeleteFile 
         Caption         =   "&Delete Song"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuSeparator10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSynchronize 
         Caption         =   "S&ynchronize Folder"
      End
      Begin VB.Menu mnuRankWrite 
         Caption         =   "&Write Rank Before Name"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Cancel"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuSupport 
         Caption         =   "&Support"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "SendMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright © 2011 ANIco.in
'Welcome to the source code of OneClick Go! SendMusic
'The code in this module deals with the send music feature of OCG.
'The modification and resdistribution of the code is completely permitted.
'---------------------------------------------------------------------------

Option Explicit
Dim DeviceFolder As String
Dim ISentIt As Boolean

Private Sub Form_Load()
    CanRestore = False
    
    'Cleaning Cache of previous sessions
    SaveSetting "ANIco.in", "SendMusicCache", "1", "1"
    DeleteSetting "ANIco.in", "SendMusicCache"
    
    'Creating the visual interface
    mnuExit.Caption = "&Cancel" & Chr$(9) & "Esc"
    ColorScheme GtSt("Color0", vbButtonFace), GtSt("Color1", vbButtonText), GtSt("Color2", vbWindowBackground), GtSt("Color3", vbWindowText), GtSt("Color4", vbWindowBackground), GtSt("Color5", vbWindowText), GtSt("Color6", vbButtonFace), GtSt("Color7", vbButtonText)
    If OneClickGo.WindowState <> 2 Then
        Width = OneClickGo.Width
        Height = OneClickGo.Height
        Left = OneClickGo.Left
        Top = OneClickGo.Top
    Else
        WindowState = 2
    End If

    'Loading Settings from registry and applying them
    AllMusicFolder = GtSt("AllMusicFolder")
    AllMusicBoxLabel.ToolTipText = AllMusicFolder
    MyMusicFolder = GtSt("MyMusicFolder")
    MyMusicBoxLabel.ToolTipText = MyMusicFolder

    'Updating AllMusicBox and MyMusicBox with main window
    For FilNum = 0 To OneClickGo.MyMusicBox.ListCount - 1
        MyMusicBox.AddItem OneClickGo.MyMusicBox.List(FilNum)
    Next
    For FilNum = 0 To OneClickGo.AllMusicBox.ListCount - 1
        AllMusicBox.AddItem OneClickGo.AllMusicBox.List(FilNum)
    Next
    MyCheck.Value = 1
    ISentIt = True
    MyCheck_Click
    AllMusicBoxLabel.Caption = "     All Music (" & CStr(AllMusicBox.SelCount) & "/" & CStr(AllMusicBox.ListCount) & ")"
    MyMusicBoxLabel.Caption = "     My Music (" & CStr(MyMusicBox.SelCount) & "/" & CStr(MyMusicBox.ListCount) & ")"
    DeviceFolder = GtSt("DeviceFolder", "[No Default/Last Used Device]")
    RecentAction.Caption = "Select the songs you want to send to your device: " & DeviceFolder
    mnuRankWrite.Checked = GtSt("RankWrite", False) 'Determine whether the user wants 001,002,..., 045 before his songs' name
    mnuSynchronize.Checked = GtSt("Synchronize", False)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    OneClickGo.CancelBtn.Visible = False
    OneClickGo.HideBtn.Visible = False
    OneClickGo.mnuSend.Caption = "&Send Music to Device"
    If Not OneClickGo.mnuSkinMode.Checked Then
        OneClickGo.Show
    End If
End Sub

Private Sub AllCheck_Click() 'Selecting/Deselecting all songs in AllMusicBox
    If ISentIt Then
        VFXupdater.Enabled = False
        AllMusicBox.Visible = False
        For FilNum = 0 To AllMusicBox.ListCount - 1
            AllMusicBox.Selected(FilNum) = Not CBool(AllCheck.Value - 1)
        Next
        AllMusicBox.Visible = True
        AllMusicBox.ListIndex = -1
        VFXupdater.Enabled = True
    End If
End Sub

Private Sub AllMusicBox_Click() 'Click and keyup event AllMusicBox
    MyBoxesClick AllMusicBox
End Sub

Private Sub AllMusicBox_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        mnuExit_Click
    End If
End Sub

Private Sub AllMusicBoxLabel_Click() 'Checking/Unchecking AllCheck checkbox through label
    If AllCheck.Value Then
        AllCheck.Value = 0
    Else
        AllCheck.Value = 1
    End If
    AllCheck_Click
End Sub

Private Sub AllSearch_DblClick() 'Events for search results of AllMusicBox
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

Private Sub BrowseBtn_Click() 'Changing the destination device folder
    Dim RemCheck As Variant
    On Error GoTo OnError
    Hide
    OldName = "Select the folder in your device that you would use as the destination for transfering your music."
    Do
        Set ShellOpener = ShellSystem.BrowseForFolder(0, OldName, &H1, 17)
        OldName = "Select the folder in your device that you would use as the destination for transfering your music."
        If Not (ShellOpener Is Nothing) Then
            DeviceFolder = ShellOpener.ParentFolder.ParseName(ShellOpener.Title).Path
            If FileSystem.folderexists(DeviceFolder) Then
                Set RemCheck = FileSystem.GetDrive(Mid$(DeviceFolder, 1, 3))
                If RemCheck.DriveType = 2 Then
                    MyMsgBox = MsgBox("The device you selected is fixed to your computer. It can cause redundancy in your files." & vbCrLf & "Are you sure you want to send music?", vbExclamation + vbYesNo + vbDefaultButton2, "Fixed drive selected")
                    If MyMsgBox = vbYes Then
                        Exit Do
                    Else
                        OldName = OldName & " [Select An External Device]"
                    End If
                Else
                    Exit Do
                End If
            Else
                OldName = OldName & " [Device Not Ready]"
            End If
        Else
            Show
            OneClickGo.RecentAction.Caption = "Operation cancelled by the user"
            Exit Sub
        End If
    Loop
    SvSt "DeviceFolder", DeviceFolder
    RecentAction.Caption = "Select the songs you want to send to your device: " & GtSt("DeviceFolder", "[No Default/Last Used Device]")
    Show
    Exit Sub
OnError:
    If Err.Number = 91 Then
        DeviceFolder = Mid$(ShellOpener.Title, Len(ShellOpener.Title) - 2, 2) & "\"
        Resume Next
    End If
End Sub

Private Sub Form_Resize()
    AllMusicBox.Width = (ScaleWidth - 360) / 2
    AllMusicBox.Left = IIf(OneClickGo.mnuMusicLeft.Checked, AllMusicBox.Width + 240, 120)
    AllMusicBox.Height = ScaleHeight - 960
    MyMusicBox.Width = AllMusicBox.Width
    MyMusicBox.Height = AllMusicBox.Height
    MyMusicBox.Left = IIf(OneClickGo.mnuMusicLeft.Checked, 120, MyMusicBox.Width + 240)
    AllSearch.Left = AllMusicBox.Left + (AllMusicBox.Width / 2)
    AllSearch.Width = AllMusicBox.Width / 2
    SearchAll.Left = AllSearch.Left
    SearchAll.Width = AllSearch.Width
    MySearch.Left = MyMusicBox.Left + (MyMusicBox.Width / 2)
    MySearch.Width = AllSearch.Width
    SearchMy.Left = MySearch.Left
    SearchMy.Width = SearchAll.Width
    AllMusicBoxLabel.Left = AllMusicBox.Left
    AllMusicBoxLabel.Width = AllMusicBox.Width
    MyMusicBoxLabel.Left = MyMusicBox.Left
    MyMusicBoxLabel.Width = MyMusicBox.Width
    StatusBar.Top = MyMusicBox.Top + MyMusicBox.Height + 60
    StatusBar.Width = ScaleWidth - 240
    RecentAction.Top = StatusBar.Top + 75
    RecentAction.Width = StatusBar.Width - 630
    SendNow.Left = StatusBar.Left + StatusBar.Width - SendNow.Width - 30
    SendNow.Top = StatusBar.Top + 45
    BrowseBtn.Left = SendNow.Left - BrowseBtn.Width
    BrowseBtn.Top = SendNow.Top
    MyCheck.Left = MyMusicBox.Left + 60
    AllCheck.Left = AllMusicBox.Left + 60
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CanRestore = True
End Sub

Private Sub MyCheck_Click() 'Selecting/Deselecting all items in MyMusicBox
    If ISentIt Then
        VFXupdater.Enabled = False
        MyMusicBox.Visible = False
        For FilNum = 0 To MyMusicBox.ListCount - 1
            MyMusicBox.Selected(FilNum) = IIf(MyCheck.Value = 1, True, False)
        Next
        MyMusicBox.Visible = True
        MyMusicBox.ListIndex = -1
        VFXupdater.Enabled = True
    End If
End Sub

Private Sub MyMusicBox_Click() 'Click and KeyUp events of MyMusicBox
    MyBoxesClick MyMusicBox
End Sub

Private Sub MyMusicBox_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        mnuExit_Click
    End If
End Sub

Private Sub MyMusicBoxLabel_Click() 'Checking/Unchecking MyCheck through label
    MyCheck.Value = IIf(MyCheck.Value = 0, 1, 0)
    MyCheck_Click
End Sub

Private Sub MySearch_DblClick() 'Events for search results of MyMusicBox
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

Private Sub SearchAll_Change() 'Search songs
    AllSearch.Enabled = True
    AllSearch.Visible = IIf(SearchAll.Text = vbNullString, False, True)
    AllSearch.Clear
    For FilNum = 0 To AllMusicBox.ListCount - 1
        If InStr(1, AllMusicBox.List(FilNum), SearchAll.Text, vbTextCompare) <> 0 Then
            AllSearch.AddItem AllMusicBox.List(FilNum) & " (" & CStr(FilNum + 1) & ")"
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
    SearchAll.BackColor = RGB(MakeDiff(cRed), MakeDiff(cGreen), MakeDiff(cBlue))
End Sub

Private Sub SearchMy_Change()
    MySearch.Enabled = True
    MySearch.Visible = IIf(SearchMy.Text = vbNullString, False, True)
    MySearch.Clear
    For FilNum = 0 To MyMusicBox.ListCount - 1
        If InStr(1, MyMusicBox.List(FilNum), SearchMy.Text, vbTextCompare) <> 0 Then
            MySearch.AddItem MyMusicBox.List(FilNum) & " (" & CStr(FilNum + 1) & ")"
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
    SearchMy.BackColor = RGB(MakeDiff(cRed), MakeDiff(cGreen), MakeDiff(cBlue))
End Sub

Private Sub SendNow_Click() 'Send Now!
    mnuSendNow_Click
End Sub

Private Sub mnuDeleteFile_Click() 'Delete Song Menu
    If TransferFrom = 1 Then
        OldName = MyMusicBox.List(MyMusicBox.ListIndex)
        MyMsgBox = MsgBox("Are you sure you want to delete this song permanently?" & vbCrLf & vbCrLf & LTrim$(OldName), vbQuestion + vbYesNo + vbDefaultButton2, "Song Deletion")
        If MyMsgBox = vbNo Then
            Exit Sub
        End If
        FileSystem.deletefile MyMusicFolder & "\" & LTrim$(OldName) & ".mp3", True
        MyMusicBox.RemoveItem MyMusicBox.ListIndex
        RecentAction.Caption = "Junk Song Removed: " & LTrim$(OldName)
        BriefHistory = Chr$(187) & " " & CStr(Now) & vbCrLf & RecentAction.Caption & vbCrLf & vbCrLf & BriefHistory
    Else
        OldName = AllMusicBox.List(AllMusicBox.ListIndex)
        MyMsgBox = MsgBox("Are you sure you want to delete this song permanently?" & vbCrLf & vbCrLf & LTrim$(OldName), vbQuestion + vbYesNo + vbDefaultButton2, "Song Deletion")
        If MyMsgBox = vbNo Then
            Exit Sub
        End If
        FileSystem.deletefile AllMusicFolder & "\" & LTrim$(OldName) & ".mp3", True
        AllMusicBox.RemoveItem AllMusicBox.ListIndex
        RecentAction.Caption = "Junk Song Removed: " & LTrim$(OldName)
        BriefHistory = Chr$(187) & " " & CStr(Now) & vbCrLf & RecentAction.Caption & vbCrLf & vbCrLf & BriefHistory
    End If
End Sub

Private Sub mnuExit_Click() 'Exit Menu
    OneClickGo.RecentAction.Caption = "Sending failed: Operation cancelled by the user"
    Unload Me
    OneClickGo.Show
End Sub

Private Sub mnuFindSong_Click() 'Find Song Menu
    If TransferFrom = 0 Then
        SearchAll.SetFocus
    Else
        SearchMy.SetFocus
    End If
End Sub

Private Sub mnuRankWrite_Click() 'Determine whether the user wants 001,002,..., 045 before his songs' name
    mnuRankWrite.Checked = IIf(mnuRankWrite.Checked, False, True)
    SvSt "RankWrite", mnuRankWrite.Checked
End Sub

Private Sub mnuSynchronize_Click() 'Determine whether to synchronize the folder to remove junk content
    mnuSynchronize.Checked = IIf(mnuSynchronize.Checked, False, True)
    SvSt "Synchronize", mnuSynchronize.Checked
End Sub

Private Sub mnuRenameFile_Click() 'Rename song Menu
    If TransferFrom = 1 Then
        OldName = MyMusicBox.List(MyMusicBox.ListIndex)
        NewName = InputBox("Give a new name for the song" & vbCrLf & vbCrLf & OldName, "Rename Song", OldName)
        If NewName = vbNullString Or NewName = OldName Then
            Exit Sub
        End If
        Name MyMusicFolder & "\" & OldName & ".mp3" As MyMusicFolder & "\" & NewName & ".mp3"
        MyMusicBox.List(MyMusicBox.ListIndex) = NewName
        SvSt CStr(MyMusicBox.ListIndex + 1), NewName
    Else
        OldName = AllMusicBox.List(AllMusicBox.ListIndex)
        NewName = InputBox("Give a new name for the song" & vbCrLf & vbCrLf & OldName, "Rename Song", OldName)
        If NewName = vbNullString Or NewName = OldName Then
            Exit Sub
        End If
        Name AllMusicFolder & "\" & OldName & ".mp3" As AllMusicFolder & "\" & NewName & ".mp3"
        AllMusicBox.List(AllMusicBox.ListIndex) = NewName
    End If
    RecentAction.Caption = "Song Renamed: " & OldName
    BriefHistory = Chr$(187) & " " & CStr(Now) & vbCrLf & RecentAction.Caption & vbCrLf & vbCrLf & BriefHistory
End Sub

Private Sub mnuSendNow_Click() 'Send Music to the device
    Dim DestFile As String
    Dim DestFile1 As String
    Dim TempN3 As Integer
    On Error GoTo OnError
    VFXupdater.Enabled = False
    Hide
    OneClickGo.CancelBtn.Visible = True
    OneClickGo.HideBtn.Visible = True
    OneClickGo.Show
    DoEvents

'Things doesn't work out well when Sync is turned on

    If mnuSynchronize.Checked Then 'Deleting extra songs before sending is started
        
        'Selected Count
        TempNum = 0
        
        'For all MyMusic list items
        For FilNum = 0 To MyMusicBox.ListCount - 1
            
            'Checking if it was selected before sending the song
            If MyMusicBox.Selected(FilNum) Then
                TempNum = TempNum + 1
                NewEntry = True
                
                'MaxCount is the maximum number of items ever existed in MyMusic list
                For TempN3 = 0 To Val(GtSt("MaxCount", CStr(MyMusicBox.ListCount)))
                    
                    'Example values stored in DestFile1
                    'If TempN3 = 0, then "
                    'Else,
                    DestFile1 = IIf(TempN3 = 0, vbNullString, String(Len(CStr(MyMusicBox.ListCount)) - Len(CStr(TempN3)), "0") & CStr(TempN3) & " ") & MyMusicBox.List(FilNum) & ".mp3"
                    DestFile = DeviceFolder & "\" & DestFile1
                    If FileSystem.FileExists(DestFile) Then
                        NewEntry = False
                        Exit For
                    End If
                Next

                If Not NewEntry Then
                    SaveSetting "ANIco.in", "SendMusicCache", DestFile1, "1"
                End If
            End If
        Next

        For FilNum = 0 To AllMusicBox.ListCount - 1
            If AllMusicBox.Selected(FilNum) Then
                TempNum = TempNum + 1
                NewEntry = True
    
                For TempN3 = 0 To Val(GtSt("MaxCount", CStr(MyMusicBox.ListCount)))
                    DestFile1 = IIf(TempN3 = 0, vbNullString, String(Len(CStr(MyMusicBox.ListCount)) - Len(CStr(TempN3)), "0") & CStr(TempN3) & " ") & AllMusicBox.List(FilNum) & ".mp3"
                    DestFile = DeviceFolder & "\" & DestFile1
                    If FileSystem.FileExists(DestFile) Then
                        NewEntry = False
                        Exit For
                    End If
                Next

                If Not NewEntry Then
                    SaveSetting "ANIco.in", "SendMusicCache", "DestFile1", "1"
                End If
            End If
        Next

        Set AllFiles = FileSystem.GetFolder(DeviceFolder)
        Set AllSongs = AllFiles.Files
        For Each AllSong In AllSongs
            If GtSt(AllSong.Name, vbNullString) = vbNullString Then
                FileSystem.deletefile DeviceFolder & "\" & AllSong.Name, True
            End If
        Next
    End If

    OneClickGo.RecentAction.Caption = "Sending your music to your device... 0% (Click cancel to pause)"
    DoEvents
    TempNum = 0
    For FilNum = 0 To MyMusicBox.ListCount - 1
        If MyMusicBox.Selected(FilNum) Then 'Checking if it was selected before sending the song
            TempNum = TempNum + 1
            NewEntry = True
                
            For TempN3 = 0 To Val(GtSt("MaxCount", CStr(MyMusicBox.ListCount)))
                DestFile = DeviceFolder & "\" & IIf(TempN3 = 0, vbNullString, String(Len(CStr(MyMusicBox.ListCount)) - Len(CStr(TempN3)), "0") & CStr(TempN3) & " ") & MyMusicBox.List(FilNum) & ".mp3"
                If FileSystem.FileExists(DestFile) Then
                    NewEntry = False
                    Exit For
                End If
            Next

            If NewEntry Then
                FileSystem.copyfile MyMusicFolder & "\" & MyMusicBox.List(FilNum) & ".mp3", DeviceFolder & "\" & IIf(mnuRankWrite.Checked, String(Len(CStr(MyMusicBox.ListCount)) - Len(CStr(FilNum + 1)), "0") & CStr(FilNum + 1) & " ", vbNullString) & MyMusicBox.List(FilNum) & ".mp3"
            Else
                Name DestFile As DeviceFolder & "\" & IIf(mnuRankWrite.Checked, String(Len(CStr(MyMusicBox.ListCount)) - Len(CStr(FilNum + 1)), "0") & CStr(FilNum + 1) & " ", vbNullString) & MyMusicBox.List(FilNum) & ".mp3"
            End If
        End If

        OneClickGo.RecentAction.Caption = "Sending your music to your device... " & CStr(Round(TempNum * 100 / (AllMusicBox.SelCount + MyMusicBox.SelCount), 0)) & "%" & IIf(OneClickGo.CancelBtn.Visible, " (Click cancel to pause)", vbNullString)
        DoEvents
    Next
    For FilNum = 0 To AllMusicBox.ListCount - 1
        If AllMusicBox.Selected(FilNum) Then
            TempNum = TempNum + 1
            NewEntry = True

            For TempN3 = 0 To Val(GtSt("MaxCount", CStr(MyMusicBox.ListCount)))
                DestFile = DeviceFolder & "\" & IIf(TempN3 = 0, vbNullString, String(Len(CStr(MyMusicBox.ListCount)) - Len(CStr(TempN3)), "0") & CStr(TempN3)) & " " & AllMusicBox.List(FilNum) & ".mp3"
                If FileSystem.FileExists(DestFile) Then
                    NewEntry = False
                    Exit For
                End If
            Next
    
            If NewEntry Then
                FileSystem.copyfile AllMusicFolder & "\" & AllMusicBox.List(FilNum) & ".mp3", DeviceFolder & "\" & IIf(mnuRankWrite.Checked, CStr(MyMusicBox.ListCount + 1) & " ", vbNullString) & AllMusicBox.List(FilNum) & ".mp3"
            Else
                Name DestFile As DeviceFolder & "\" & IIf(mnuRankWrite.Checked, CStr(MyMusicBox.ListCount + 1) & " ", vbNullString) & AllMusicBox.List(FilNum) & ".mp3"
            End If
        End If

        OneClickGo.RecentAction.Caption = "Sending your music to your device... " & CStr(Round(TempNum * 100 / (AllMusicBox.SelCount + MyMusicBox.SelCount), 0)) & "%" & IIf(OneClickGo.CancelBtn.Visible, " (Click cancel to pause)", vbNullString)
        DoEvents
    Next
    OneClickGo.RecentAction.Caption = "Your music sent to device successfully"
    BriefHistory = Chr$(187) & " " & CStr(Now) & vbCrLf & OneClickGo.RecentAction.Caption & vbCrLf & vbCrLf & BriefHistory
    Unload Me
    Exit Sub
OnError:
    If Err.Number = -2147024784 Then
        OneClickGo.RecentAction.Caption = "Sending failed: Disk full, delete some files to make room for your music"
    ElseIf Err.Number = 53 Then
        OneClickGo.RecentAction.Caption = "Sending paused: Cancelled by user"
    Else
        OneClickGo.RecentAction.Caption = "Sending failed: " & Err.Description
    End If
    Unload Me
End Sub

Private Sub mnuSupport_Click() 'Support Menu
    ShellExecute 0, "OPEN", App.Path & "\Documents\Help.pdf", vbNullString, App.Path & "\Documents\", 1
End Sub

Private Sub ColorScheme(Col0 As Long, Col1 As Long, Col2 As Long, Col3 As Long, Col4 As Long, Col5 As Long, Col6 As Long, Col7 As Long) 'Procedure to control visual interface
    BackColor = Col0
    AllMusicBoxLabel.ForeColor = Col1
    MyMusicBoxLabel.ForeColor = Col1
    AllMusicBox.BackColor = Col2
    AllSearch.BackColor = Col2
    SearchAll.BackColor = Col2
    AllMusicBox.ForeColor = Col3
    AllSearch.ForeColor = Col3
    SearchAll.ForeColor = Col3
    MyMusicBox.BackColor = Col4
    MySearch.BackColor = Col4
    SearchMy.BackColor = Col4
    MyMusicBox.ForeColor = Col5
    MySearch.ForeColor = Col5
    SearchMy.ForeColor = Col5
    StatusBar.BackColor = Col6
    RecentAction.ForeColor = Col7
    ColorCodeToRGB AllMusicBox.BackColor
    SearchAll.BackColor = RGB(MakeDiff(cRed), MakeDiff(cGreen), MakeDiff(cBlue))
    ColorCodeToRGB MyMusicBox.BackColor
    SearchMy.BackColor = RGB(MakeDiff(cRed), MakeDiff(cGreen), MakeDiff(cBlue))
    NewFont AllMusicBoxLabel, GtSt("FN0", "MS Sans Serif"), GtSt("FB0", True), GtSt("FI0", False), Val(GtSt("FS0", "13"))
    NewFont MyMusicBoxLabel, GtSt("FN0", "MS Sans Serif"), GtSt("FB0", True), GtSt("FI0", False), Val(GtSt("FS0", "13"))
    NewFont AllMusicBox, GtSt("FN1", "MS Sans Serif"), GtSt("FB1", False), GtSt("FI1", False), Val(GtSt("FS1", "7"))
    NewFont MyMusicBox, GtSt("FN2", "MS Sans Serif"), GtSt("FB2", False), GtSt("FI2", False), Val(GtSt("FS2", "10"))
    NewFont RecentAction, GtSt("FN3", "MS Sans Serif"), GtSt("FB3", True), GtSt("FI3", False), Val(GtSt("FS3", "10"))
End Sub

Private Sub MyBoxesClick(BoxName As ListBox) 'Click on Music Boxes
    TransferFrom = IIf(BoxName.Name = "AllMusicBox", 0, 1)
    AllSearch.Visible = False
    MySearch.Visible = False
    AllMusicBoxLabel.Caption = "     All Music (" & CStr(AllMusicBox.SelCount) & "/" & CStr(AllMusicBox.ListCount) & ")"
    MyMusicBoxLabel.Caption = "     My Music (" & CStr(MyMusicBox.SelCount) & "/" & CStr(MyMusicBox.ListCount) & ")"
End Sub

Private Sub VFXupdater_Timer() 'Timer to update music boxes' checkbox on top
    ISentIt = False
    Select Case AllMusicBox.SelCount
        Case 0
            AllCheck.Value = 0
        Case AllMusicBox.ListCount
            AllCheck.Value = 1
        Case Else
            AllCheck.Value = 2
    End Select
    Select Case MyMusicBox.SelCount
        Case 0
            MyCheck.Value = 0
        Case MyMusicBox.ListCount
            MyCheck.Value = 1
        Case Else
            MyCheck.Value = 2
    End Select
    If AllMusicBox.SelCount + MyMusicBox.SelCount = 0 Or Not FileSystem.folderexists(DeviceFolder) Then
        mnuSendNow.Enabled = False
        SendNow.Enabled = False
    Else
        mnuSendNow.Enabled = True
        SendNow.Enabled = True
    End If
    ISentIt = True
End Sub
