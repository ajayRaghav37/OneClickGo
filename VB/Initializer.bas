Attribute VB_Name = "Initializer"
'Copyright © 2011 ANIco.in
'Welcome to the source code of OneClick Go! Skin Mode 2
'The code in this module deals with the main module of OCG that contains important functions and procedures.
'The modification and resdistribution of the code is completely permitted.
'---------------------------------------------------------------------------

'The declaration of the variables are made in this module so that they can be used throughout the program.

Option Explicit

Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Public CanRestore As Boolean
Public ActiveIndex As Integer
Public FileSystem As Variant
Public ShellSystem As Variant
Public ShellOpener As Variant
Public AllFiles As Variant
Public AllSongs As Variant
Public AllSong As Variant
Public MyFiles As Variant
Public MySongs As Variant
Public MySong As Variant
Public NewName As String
Public OldName As String
Public AllMusicFolder As String
Public MyMusicFolder As String
Public FilNum As Integer
Public TransferFrom As Integer
Public MyMsgBox As VbMsgBoxResult
Public TempNum As Integer
Public NewEntry As Boolean
Public cRed As Byte
Public cGreen As Byte
Public cBlue As Byte
Public NowPlaying As Integer
Public PlayingAll As Boolean
Public CurrentSkin As Integer
Public Xinit As Single
Public Yinit As Single
Public IamMoved As Boolean
Public Rotator As Boolean
Public SeekCtrl2 As Integer
Public OpacityNow As Integer
Public IsOut As Boolean
Public BriefHistory As String

'The powerful API (Application Programming Interface) procedures' and functions' declarations for skin effects and other.
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crey As Single, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Sub Main() 'Sub Main is the startup procedure of OneClick Go!
    If App.PrevInstance And GtSt("AllowRun", "0") = "0" Then 'Checking for Previous Instances running in background
        MyMsgBox = MsgBox("OneClick Go! is already running." & vbCrLf & String(30, "—") & vbCrLf & "Running multiple sessions simultaneously might result in corruption of your playlist. If you think it's an error, you can feel free to continue with a new session. If you select 'No', we will try to restore the window of your current session." & vbCrLf & String(30, "—") & vbCrLf & "Are you sure you want to continue?", vbYesNoCancel + vbExclamation + vbDefaultButton2, "Identical Process Detected")
        If MyMsgBox = vbNo Then
            SvSt "QuickRestore", "1"
        End If
        If MyMsgBox <> vbYes Then
            End
        End If
    End If
    
    'Initializing FileSystemObject
    Set FileSystem = CreateObject("Scripting.filesystemobject")
    Set ShellSystem = CreateObject("Shell.Application")
    
    CheckIntegrity 'Checking if any file is missing
    SvSt "AllowRun", "0" 'Disallow running of multiple sessions
    Load OneClickGo
End Sub

Public Sub ColorCodeToRGB(Col As Long) 'Procedure to find RGB values of a long color
    If Col < 0 Then
        Col = GetSysColor(Col And &HFF)
    End If
    cRed = Col And &HFF
    cGreen = (Col \ &H100) And &HFF
    cBlue = (Col \ &H10000) And &HFF
End Sub

Public Sub DlSt(Optional SettingName As String = "") 'Procedure to delete setting
    If SettingName = vbNullString Then
        DeleteSetting "ANIco.in", "OneClick Go!"
    Else
        DeleteSetting "ANIco.in", "OneClick Go!", SettingName
    End If
End Sub

Public Sub NewFont(ObjectName As Control, FontName As String, FontBold As Boolean, FontItalic As Boolean, FontSize As Single) 'Procedure to change fonts of objects
    ObjectName.FontName = FontName
    ObjectName.FontBold = FontBold
    ObjectName.FontItalic = FontItalic
    ObjectName.FontSize = FontSize
End Sub

Public Sub SvSt(SettingName As String, SettingValue As Variant) 'Procedure to save settings
    SaveSetting "ANIco.in", "OneClick Go!", SettingName, SettingValue
End Sub

Public Sub SvSt2(SettingName As String, SettingValue As Variant) 'Procedure to save settings for Shuffle Cache
    SaveSetting "ANIco.in", "SongCache", SettingName, SettingValue
End Sub

Public Function GtSt(SettingName As String, Optional DefaultValue As String) As String 'Function to get setting
    GtSt = GetSetting("ANIco.in", "OneClick Go!", SettingName, DefaultValue)
End Function

Public Function GtSt2(SettingName As String, Optional DefaultValue As String) As String 'Function to get setting of shuffle cache
    GtSt2 = GetSetting("ANIco.in", "SongCache", SettingName, DefaultValue)
End Function

Public Function OfName(PictureName As String) As String 'Function to get full path of a picture
    OfName = App.Path & "\Skin\" & PictureName & ".BMP"
End Function

Public Sub Refresher() 'Procedure to change the interface of SkinnedOCG as soon as the song changes
    If PlayingAll Then
        SkinnedOCG.NPS.Caption = OneClickGo.AllMusicBox.List(NowPlaying)
        SkinnedOCG.Rank.Caption = "N/A"
    Else
        SkinnedOCG.NPS.Caption = OneClickGo.MyMusicBox.List(NowPlaying)
        SkinnedOCG.Rank.Caption = CStr(NowPlaying + 1)
    End If
    If SkinnedOCG.NPS.Width > SkinnedOCG.NPSBlockerR.Left - SkinnedOCG.NPSBlockerL.Width Then
        SkinnedOCG.NPS.Left = SkinnedOCG.NPSBlockerL.Width
    Else
        SkinnedOCG.NPS.Left = (SkinnedOCG.NPSBlockerR.Left + SkinnedOCG.NPSBlockerL.Width - SkinnedOCG.NPS.Width) / 2
    End If
    SkinnedOCG.NPSinfo.Visible = True
    SkinnedOCG.NPSVvanisher.Enabled = True
End Sub

Public Sub Refresher2() 'Procedure to change the interface of SkinnedOCG2 as soon as the song changes
    If SkinnedOCG2.NowPlayingSkin.Width > SkinnedOCG2.NPSback1.Left - SkinnedOCG2.NPSback2.Width Then
        SkinnedOCG2.NowPlayingSkin.Left = SkinnedOCG2.NPSback2.Width
    Else
        SkinnedOCG2.NowPlayingSkin.Left = (SkinnedOCG2.NPSback1.Left + SkinnedOCG2.NPSback2.Width - SkinnedOCG2.NowPlayingSkin.Width) / 2
    End If
    If PlayingAll Then
        SkinnedOCG2.NowPlayingSkin.Caption = OneClickGo.AllMusicBox.List(NowPlaying)
        SkinnedOCG2.Rank.Caption = "(Unrated)"
        SkinnedOCG2.Rank.FontSize = 10
    Else
        SkinnedOCG2.NowPlayingSkin.Caption = OneClickGo.MyMusicBox.List(NowPlaying)
        SkinnedOCG2.Rank.Caption = "#" & CStr(NowPlaying + 1)
        SkinnedOCG2.Rank.FontSize = 14
    End If
    SkinnedOCG2.NPSinfo.Visible = True
    SkinnedOCG2.NPSVvanisher.Enabled = True
End Sub

Private Sub CheckIntegrity()
    If Not FileSystem.FileExists(App.Path & "\Documents\END USER LICENSE AGREEMENT.rtf") Then
        MyMsgBox = MsgBox("File is missing: END USER LICENSE AGREEMENT.rtf" & vbCrLf & vbCrLf & "The file is required to review end user license agreement. It is not a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton1, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Documents\Help.pdf") Then
        MyMsgBox = MsgBox("File is missing: Help.pdf" & vbCrLf & vbCrLf & "The file is required to view help. It is not a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton1, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\NextPressedPause.BMP") Then
        MyMsgBox = MsgBox("File is missing: NextPressedPause.BMP" & vbCrLf & vbCrLf & "The file is required by Freezed Blue Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\NextPressedPlay.BMP") Then
        MyMsgBox = MsgBox("File is missing: NextPressedPlay.BMP" & vbCrLf & vbCrLf & "The file is required by Freezed Blue Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\NPSback.BMP") Then
        MyMsgBox = MsgBox("File is missing: NPSback.BMP" & vbCrLf & vbCrLf & "The file is required by Freezed Blue Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\NPSback1.BMP") Then
        MyMsgBox = MsgBox("File is missing: NPSback1.BMP" & vbCrLf & vbCrLf & "The file is required by Freezed Blue Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\NPSback2.BMP") Then
        MyMsgBox = MsgBox("File is missing: NPSback2.BMP" & vbCrLf & vbCrLf & "The file is required by Freezed Blue Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\NPSBlockerL.BMP") Then
        MyMsgBox = MsgBox("File is missing: NPSBlockerL.BMP" & vbCrLf & vbCrLf & "The file is required by Rockstar Gold Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\NPSBlockerR.BMP") Then
        MyMsgBox = MsgBox("File is missing: NPSBlockerR.BMP" & vbCrLf & vbCrLf & "The file is required by Rockstar Gold Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\PauseDefault.BMP") Then
        MyMsgBox = MsgBox("File is missing: PauseDefault.BMP" & vbCrLf & vbCrLf & "The file is required by Freezed Blue Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\PausePressed.BMP") Then
        MyMsgBox = MsgBox("File is missing: PausePressed.BMP" & vbCrLf & vbCrLf & "The file is required by Freezed Blue Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\PlayDefault.BMP") Then
        MyMsgBox = MsgBox("File is missing: PlayDefault.BMP" & vbCrLf & vbCrLf & "The file is required by Freezed Blue Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\PlayPauseADefault.BMP") Then
        MyMsgBox = MsgBox("File is missing: PlayPauseADefault.BMP" & vbCrLf & vbCrLf & "The file is required by Rockstar Gold Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\PlayPauseAPressed.BMP") Then
        MyMsgBox = MsgBox("File is missing: PlayPauseAPressed.BMP" & vbCrLf & vbCrLf & "The file is required by Rockstar Gold Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\PlayPauseLDefault.BMP") Then
        MyMsgBox = MsgBox("File is missing: PlayPauseLDefault.BMP" & vbCrLf & vbCrLf & "The file is required by Rockstar Gold Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\PlayPauseLPressed.BMP") Then
        MyMsgBox = MsgBox("File is missing: PlayPauseLPressed.BMP" & vbCrLf & vbCrLf & "The file is required by Rockstar Gold Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\PlayPressed.BMP") Then
        MyMsgBox = MsgBox("File is missing: PlayPressed.BMP" & vbCrLf & vbCrLf & "The file is required by Freezed Blue Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\PrevNextDefault.BMP") Then
        MyMsgBox = MsgBox("File is missing: PrevNextDefault.BMP" & vbCrLf & vbCrLf & "The file is required by Rockstar Gold Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\PrevNextNPressed.BMP") Then
        MyMsgBox = MsgBox("File is missing: PrevNextNPressed.BMP" & vbCrLf & vbCrLf & "The file is required by Rockstar Gold Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\PrevNextPPressed.BMP") Then
        MyMsgBox = MsgBox("File is missing: PrevNextPPressed.BMP" & vbCrLf & vbCrLf & "The file is required by Rockstar Gold Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\PrevPressedPause.BMP") Then
        MyMsgBox = MsgBox("File is missing: PrevPressedPause.BMP" & vbCrLf & vbCrLf & "The file is required by Freezed Blue Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\PrevPressedPlay.BMP") Then
        MyMsgBox = MsgBox("File is missing: PrevPressedPlay.BMP" & vbCrLf & vbCrLf & "The file is required by Freezed Blue Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\StopDefault.BMP") Then
        MyMsgBox = MsgBox("File is missing: StopDefault.BMP" & vbCrLf & vbCrLf & "The file is required by Rockstar Gold Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\StopPressed.BMP") Then
        MyMsgBox = MsgBox("File is missing: StopPressed.BMP" & vbCrLf & vbCrLf & "The file is required by Rockstar Gold Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\StopPressedPause.BMP") Then
        MyMsgBox = MsgBox("File is missing: StopPressedPause.BMP" & vbCrLf & vbCrLf & "The file is required by Freezed Blue Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
    If Not FileSystem.FileExists(App.Path & "\Skin\StopPressedPlay.BMP") Then
        MyMsgBox = MsgBox("File is missing: StopPressedPlay.BMP" & vbCrLf & vbCrLf & "The file is required by Freezed Blue Skin. It is a compulsary file" & vbCrLf & vbCrLf & "If you want this error to not be displayed each time, copy the original file to (or create a fake file in) " & App.Path & "\Documents" & vbCrLf & vbCrLf & "Do you want to continue loading OneClick Go! without this file?", vbCritical + vbYesNo + vbDefaultButton2, "File Not Found")
        If MyMsgBox = vbNo Then
            End
        End If
    End If
End Sub

Public Function MakeDiff(cColor As Byte)
    If cColor < 128 Then
        MakeDiff = cColor + 16
    Else
        MakeDiff = cColor - 16
    End If
End Function

Public Function DurationStr(nSeconds As Long) As String
    DurationStr = LeadZero((nSeconds \ 60) Mod 60) & ":" & LeadZero(nSeconds Mod 60)  'Writing Minutes & Seconds
    If nSeconds > 3600 Then
        DurationStr = LeadZero(nSeconds \ 3600) & ":" & DurationStr           'Writing Hours
    End If
End Function

Public Function LeadZero(nNumber As Long, Optional nDig = 2) As String
    LeadZero = CStr(nNumber)
    If nNumber < (10 ^ (nDig - 1)) Then
        LeadZero = String$(nDig - Len(LeadZero), "0") & LeadZero
    End If
End Function
