Attribute VB_Name = "Module1"
'registry reading data
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

Public Const SW_SHOWNORMAL = 1
Public Const SND_ASYNC = &H1
Const READ_CONTROL = &H20000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const KEY_QUERY_VALUE = &H1&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const REG_SZ = 1&
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const ERROR_SUCCESS = 0&

Dim hKey As Long
Dim sBuffer As String
Dim lBufferSize As Long
'end reg reading data

'sound playing data
Public Declare Function mciSendStringA Lib "winmm.dll" (ByVal lpstrCommand As String, ByVal lpstrRtnString As Any, ByVal wRtnLength As Integer, ByVal hCallback As Integer) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'end sound playing data

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Type ReceiveInfoStructure
    Nick As String
    Filename As String
    IP As String
    Port As Long
    Filelength As Long
    UsableFileLength As String
    ReceivedBytes As Long
    InUse As Boolean
    StartTime As Date
    CPS As Long
    TimeRemaining As Integer
    LastTimeReceivedData As Date
    Resumed As Boolean
End Type

Type PreferencesStructure
    VersionReply As String
    Nickname As String
    Server As String
    ServerPort As Integer
    DownloadPath As String
    NagDelay As Integer
    CommandDelay As Integer
    Botshare As Integer
    ChannelListen As Integer
    Sounds As Integer
    OpenOnDownload As Integer
    DCCRelay As Integer
    DCCRelayNick As String
    UseMSG As Boolean
    UseCTCP As Boolean
    Channels() As Variant
End Type

Global Preferences As PreferencesStructure
Global ReceiveInfo() As ReceiveInfoStructure
Global ReceiveBox() As New ReceiveBoxForm
Global NameAndVersion As String
Global RealNameAndVersion As String
Global BotSwapChannel As String
Global UsingNickname As String
Global ShouldIBeep As Boolean

Public Function ReadableFilesize(ThisFilesize)

'work out the size of the file so that it can be displayed
Select Case ThisFilesize
'Case Is > 1000000
    'ReadableFilesize = Format((ThisFilesize / 1000000), "0.00") + " MB"
Case Is > 1000
    ReadableFilesize = LTrim(Str(Fix(ThisFilesize / 1000)) + " K")
Case Is < 1001
    ReadableFilesize = LTrim(Str(ThisFilesize)) + " Bytes"
End Select

End Function

Public Function DirectoryStatus(ThisDirectory As String)

On Error GoTo NotFound

DirectoryStatus = GetAttr(ThisDirectory)
Exit Function

NotFound:
DirectoryStatus = 999

End Function
Public Sub ReadINI()

On Error GoTo ErrorHandler

Dim DataString As String

If DirectoryStatus(App.Path + "\ircft.ini") = 999 Then
    MainForm.Visible = False
    LegalForm.Show
    Exit Sub
End If

Open App.Path + "\ircft.ini" For Input As #1

Input #1, DataString
Do While Left(DataString, 8) = "channel="
    PrefsForm.ChannelList.AddItem Right(DataString, Len(DataString) - InStr(1, DataString, "="))
    Input #1, DataString
Loop

PrefsForm.UseMSGOption.Value = Right(DataString, Len(DataString) - InStr(1, DataString, "="))
Input #1, DataString
PrefsForm.UseCTCPOption.Value = Right(DataString, Len(DataString) - InStr(1, DataString, "="))
Input #1, DataString
PrefsForm.ServerName.Text = Right(DataString, Len(DataString) - InStr(1, DataString, "="))
Input #1, DataString
PrefsForm.ServerPort.Text = Right(DataString, Len(DataString) - InStr(1, DataString, "="))
Input #1, DataString
PrefsForm.DownloadPath.Text = Right(DataString, Len(DataString) - InStr(1, DataString, "="))
Input #1, DataString
PrefsForm.SoundsCheck.Value = Val(Right(DataString, Len(DataString) - InStr(1, DataString, "=")))
Input #1, DataString
PrefsForm.Nickname.Text = Right(DataString, Len(DataString) - InStr(1, DataString, "="))
Input #1, DataString
PrefsForm.HeresMyBotsCheck.Value = Val(Right(DataString, Len(DataString) - InStr(1, DataString, "=")))
Input #1, DataString
PrefsForm.VersionText.Text = Right(DataString, Len(DataString) - InStr(1, DataString, "="))
Input #1, DataString
MainForm.Left = Val(Right(DataString, Len(DataString) - InStr(1, DataString, "=")))
Input #1, DataString
MainForm.Top = Val(Right(DataString, Len(DataString) - InStr(1, DataString, "=")))
Input #1, DataString
MainForm.Height = Val(Right(DataString, Len(DataString) - InStr(1, DataString, "=")))
Input #1, DataString
MainForm.Width = Val(Right(DataString, Len(DataString) - InStr(1, DataString, "=")))
Input #1, DataString
PrefsForm.ChanListenCheck.Value = Val(Right(DataString, Len(DataString) - InStr(1, DataString, "=")))
Input #1, DataString
PrefsForm.TimeBetweenCommandsText.Text = Val(Right(DataString, Len(DataString) - InStr(1, DataString, "=")))
Input #1, DataString
PrefsForm.TimeBetweenNagsText.Text = Val(Right(DataString, Len(DataString) - InStr(1, DataString, "=")))
Input #1, DataString
PrefsForm.OpenOnDownloadCheck.Value = Val(Right(DataString, Len(DataString) - InStr(1, DataString, "=")))
Input #1, DataString
PrefsForm.DCCRelayCheck.Value = Val(Right(DataString, Len(DataString) - InStr(1, DataString, "=")))
Input #1, DataString
PrefsForm.DCCRelayNickText.Text = Right(DataString, Len(DataString) - InStr(1, DataString, "="))

Close #1

If PrefsForm.ChanListenCheck.Value = 0 Then MainForm.PartChannelsButton.Visible = False

'check that there's a sensible time between commands
If Val(PrefsForm.TimeBetweenCommandsText.Text) < 200 Then
    PrefsForm.TimeBetweenCommandsText.Text = "200"
End If

'check that there's a sensible time between nags
If Val(PrefsForm.TimeBetweenNagsText.Text) < 3000 Then
    PrefsForm.TimeBetweenNagsText.Text = "3000"
End If

SetPreferences

Exit Sub

ErrorHandler:

Close #1
WriteINI

End Sub

Public Sub WriteINI()

If DirectoryStatus(App.Path + "\ircft.ini") <> 999 Then
    Kill App.Path + "\ircft.ini"
End If
Open App.Path + "\ircft.ini" For Output As #1
Print #1, "msgoption=True"
Print #1, "ctcpoption=False"
Print #1, "server=irc.pacbell.net"
Print #1, "port=6667"
Print #1, "path=c:\"
Print #1, "sounds=1"
Print #1, "nick=Nickname"
Print #1, "botshare=1"
Print #1, "version=IRC Filetool v1.2"
Print #1, "left=100"
Print #1, "top=100"
Print #1, "height=6000"
Print #1, "width=8000"
Print #1, "chanlisten=0"
Print #1, "commandpause=250"
Print #1, "nagpause=3000"
Print #1, "openondownload=0"
Print #1, "dccrelay=0"
Print #1, "dccrelaynick=Nickname"
Close #1

SetPreferences

End Sub

Public Function GetWord(ThisLine As String, WordNumber As Integer)

On Error GoTo NotEnoughWordsInString

Dim i, LastSpacePos, NextSpacePos As Integer
Dim ThisWord As String

LastSpacePos = 1
For i = 1 To WordNumber
    NextSpacePos = InStr(LastSpacePos, ThisLine, " ")
    While Mid(ThisLine, NextSpacePos + 1, 1) = " "
        NextSpacePos = NextSpacePos + 1
    Wend
    If NextSpacePos = 0 Then NextSpacePos = Len(ThisLine) + 1
    ThisWord = Mid(ThisLine, LastSpacePos, NextSpacePos - LastSpacePos)
    LastSpacePos = NextSpacePos + 1
Next i

GetWord = ThisWord

Exit Function

NotEnoughWordsInString:
GetWord = "%error%"

End Function
Public Function DecodeIP(LongIPString As String)

Dim LongIP, Remainder, FirstPart, SecondPart, ThirdPart, FourthPart As Double
Dim FirstPartString, SecondPartString, ThirdPartString, FourthPartString, TempIP As String
Dim i As Integer

'remove brackets
For i = 1 To Len(LongIPString)
    If (Mid(LongIPString, i, 1) <> "(") And (Mid(LongIPString, i, 1) <> ")") Then TempIP = TempIP + Mid(LongIPString, i, 1)
Next i

LongIP = Val(TempIP)
FirstPart = Fix(LongIP / 16777216)
Remainder = LongIP - (FirstPart * 16777216)
SecondPart = Fix(Remainder / 65536)
Remainder = Remainder - (SecondPart * 65536)
ThirdPart = Fix(Remainder / 256)
Remainder = Remainder - (ThirdPart * 256)
FourthPart = Remainder

FirstPartString = LTrim(Str(FirstPart))
SecondPartString = LTrim(Str(SecondPart))
ThirdPartString = LTrim(Str(ThirdPart))
FourthPartString = LTrim(Str(FourthPart))

DecodeIP = FirstPartString + "." + SecondPartString + "." + ThirdPartString + "." + FourthPartString

End Function

Public Function EncodeIP(IPAddress As String)

Dim Nibble(4) As Integer
Dim FirstPart As Double
Dim SecondPart As Double
Dim ThirdPart As Double
Dim FourthPart As Double
Dim LongIP As Double
Dim ThisPos, NextPos, i As Integer

IPAddress = IPAddress + "."

ThisPos = 0
For i = 1 To 4
    NextPos = InStr(ThisPos + 1, IPAddress, ".")
    Nibble(i) = Val(Mid(IPAddress, ThisPos + 1, NextPos - ThisPos - 1))
    ThisPos = NextPos
Next i

FirstPart = Nibble(1) * 16777216#
SecondPart = Nibble(2) * 65536#
ThirdPart = Nibble(3) * 256#
FourthPart = Nibble(4)

LongIP = FirstPart + SecondPart + ThirdPart + FourthPart

EncodeIP = LTrim(Str(LongIP))

End Function

Public Function GetFreeSlot()

Dim j As Integer
Dim found As Boolean

found = False
For j = 10 To UBound(ReceiveInfo)
    If ReceiveInfo(j).InUse = False Then
        ReceiveInfo(j).InUse = True
        found = True
        Exit For
    End If
Next
If found = False Then
    j = UBound(ReceiveInfo) + 1
    ReDim Preserve ReceiveInfo(j) As ReceiveInfoStructure
    ReDim Preserve ReceiveBox(j) As New ReceiveBoxForm
    ReceiveInfo(j).InUse = True
End If

GetFreeSlot = j
Debug.Print "Assigned slot:"; j

End Function

Public Function GetRegInfo(Item As String, Area As String)

On Error GoTo Unknown

sBuffer = Space(255)
lBufferSize = Len(sBuffer)

'result = RegOpenKeyEx(&H80000001, "Software\Microsoft\MS Setup (ACME)\User Info", 0, KEY_READ, hKey)
result = RegOpenKeyEx(&H80000001, Area, 0, KEY_READ, hKey)
'Debug.Print "ok? "; ERROR_SUCCESS = result
result = RegQueryValueEx(hKey, Item, 0, REG_SZ, sBuffer, lBufferSize)
'Debug.Print "ok? "; ERROR_SUCCESS = result
result = RegCloseKey(hKey)
'Debug.Print "ok? "; ERROR_SUCCESS = result
sBuffer = RTrim(sBuffer)
sBuffer = Left(sBuffer, Len(sBuffer) - 1)
If sBuffer = "" Then GoTo Unknown
GetRegInfo = sBuffer

Exit Function

Unknown:
GetRegInfo = "Unknown"

End Function
Public Sub AppendLog(ThisLine As String)

MainForm.LogList.AddItem Format(Now, "hh:mm") & " " & ThisLine
If MainForm.LogList.ListCount > 3 Then
    MainForm.LogList.TopIndex = MainForm.LogList.ListCount - 3
End If

End Sub

Public Sub SendToServerImmediately(ThisLine As String)

If MainForm.ServerConnection.State = sckConnected Then
    MainForm.ServerConnection.SendData ThisLine + vbCrLf
    Debug.Print ThisLine
Else
    Debug.Print "Error sending to server"
End If

End Sub
Public Function AssignFilename(ThisFilename As String)

Dim Addition As String
Dim Identifier As Integer

Identifier = 1
Addition = ""
While LCase(Dir(Preferences.DownloadPath + Addition + ThisFilename)) = LCase(Addition + ThisFilename)
    Addition = "duplicate_" + LTrim(Str(Identifier)) + "_of_"
    Identifier = Identifier + 1
Wend

AssignFilename = Addition + ThisFilename

End Function

Public Sub PlayWAV(ThisSample As String)

Dim RetVal As Long

RetVal = mciSendStringA("play " + ThisSample, "", 255, 0)

End Sub

Public Sub SetPreferences()

Dim i As Integer

Preferences.VersionReply = PrefsForm.VersionText.Text
Preferences.Nickname = PrefsForm.Nickname.Text
Preferences.Server = PrefsForm.ServerName.Text
Preferences.ServerPort = Val(PrefsForm.ServerPort.Text)
Preferences.DownloadPath = PrefsForm.DownloadPath.Text
Preferences.NagDelay = Val(PrefsForm.TimeBetweenNagsText.Text)
Preferences.CommandDelay = Val(PrefsForm.TimeBetweenCommandsText.Text)
Preferences.Botshare = PrefsForm.HeresMyBotsCheck.Value
Preferences.ChannelListen = PrefsForm.ChanListenCheck.Value
Preferences.Sounds = PrefsForm.SoundsCheck.Value
Preferences.OpenOnDownload = PrefsForm.OpenOnDownloadCheck.Value
Preferences.DCCRelay = PrefsForm.DCCRelayCheck.Value
Preferences.DCCRelayNick = PrefsForm.DCCRelayNickText.Text
Preferences.UseMSG = PrefsForm.UseMSGOption
Preferences.UseCTCP = PrefsForm.UseCTCPOption


ReDim Preserve Preferences.Channels(PrefsForm.ChannelList.ListCount)
For i = 1 To PrefsForm.ChannelList.ListCount
    Preferences.Channels(i) = PrefsForm.ChannelList.List(i - 1)
Next i
'Debug.Print "ubound: "; UBound(Preferences.Channels)
'For i = 1 To UBound(Preferences.Channels)
'    Debug.Print i; " here: "; Preferences.Channels(i)
'Next i

NameAndVersion = PrefsForm.VersionText.Text

If Preferences.ChannelListen = 0 Then
    MainForm.PartChannelsButton.Visible = False
Else
    MainForm.PartChannelsButton.Visible = True
End If

MainForm.CommandsTimer.Interval = Preferences.CommandDelay * 10

End Sub
