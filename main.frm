VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form MainForm 
   Caption         =   "IRC Filetool"
   ClientHeight    =   8430
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11790
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   11790
   Begin VB.ListBox MyFilesList 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   5040
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Timer ChannelsTimer 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   4440
      Top             =   7680
   End
   Begin VB.Timer CommandsTimer 
      Left            =   3960
      Top             =   7680
   End
   Begin VB.Timer ServerDataChecker 
      Interval        =   10
      Left            =   3480
      Top             =   7680
   End
   Begin VB.ListBox ServerDataList 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5760
      TabIndex        =   20
      Top             =   7800
      Visible         =   0   'False
      Width           =   4395
   End
   Begin VB.Timer ShouldIBeepTimer 
      Interval        =   1000
      Left            =   3000
      Top             =   7680
   End
   Begin VB.CommandButton GetURLButton 
      Caption         =   "Go To URL"
      Height          =   500
      Left            =   5520
      TabIndex        =   19
      Top             =   120
      Width           =   1000
   End
   Begin VB.CommandButton HelpButton 
      Caption         =   "Help"
      Height          =   500
      Left            =   9600
      TabIndex        =   18
      Top             =   120
      Width           =   1000
   End
   Begin VB.Timer NagTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Tag             =   "0"
      Top             =   7680
   End
   Begin VB.CommandButton NagForOffersButton 
      Caption         =   "Nag For Offers"
      Height          =   495
      Left            =   2280
      TabIndex        =   17
      Top             =   120
      Width           =   1000
   End
   Begin VB.Timer AntiIdleTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Tag             =   "0"
      Top             =   7680
   End
   Begin VB.Timer BroadcastBotsTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Tag             =   "0"
      Top             =   7680
   End
   Begin MSWinsockLib.Winsock IdentWinsock 
      Left            =   11280
      Top             =   7680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton FindButton 
      Caption         =   "Find"
      Height          =   495
      Left            =   1200
      TabIndex        =   16
      Top             =   120
      Width           =   1000
   End
   Begin VB.CommandButton PartChannelsButton 
      Caption         =   "Leave Channels"
      Height          =   495
      Left            =   6600
      TabIndex        =   15
      Top             =   120
      Width           =   1000
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   8175
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame OffersFrame 
      Caption         =   "Offers"
      Height          =   4035
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   4815
      Begin VB.ListBox OffersList 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         ItemData        =   "main.frx":08CA
         Left            =   120
         List            =   "main.frx":08CC
         MultiSelect     =   2  'Extended
         TabIndex        =   13
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame LogFrame 
      Caption         =   "Messages"
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   660
      Width           =   11535
      Begin VB.ListBox LogList 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   690
         ItemData        =   "main.frx":08CE
         Left            =   120
         List            =   "main.frx":08D0
         TabIndex        =   11
         Top             =   240
         Width           =   11295
      End
   End
   Begin VB.ListBox WaitBotsList 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   10200
      TabIndex        =   9
      Top             =   6000
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.CommandButton ExitButton 
      Caption         =   "Exit"
      Height          =   500
      Left            =   10680
      TabIndex        =   8
      Top             =   120
      Width           =   1000
   End
   Begin VB.Timer WaitBeforeAskingTimer 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   1080
      Tag             =   "0"
      Top             =   7680
   End
   Begin VB.CommandButton PreferencesButton 
      Caption         =   "Preferences"
      Height          =   500
      Left            =   4440
      TabIndex        =   7
      Top             =   120
      Width           =   1000
   End
   Begin VB.CommandButton RequestFilesButton 
      Caption         =   "Request Files"
      Height          =   500
      Left            =   3360
      TabIndex        =   6
      Top             =   120
      Width           =   1000
   End
   Begin MSWinsockLib.Winsock ReceiveConnection 
      Index           =   0
      Left            =   10800
      Top             =   7680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox RequestedBotsList 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   8760
      TabIndex        =   5
      Top             =   6000
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.ListBox RespondedBotsList 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   7320
      TabIndex        =   4
      Top             =   6000
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.ListBox CommandsList 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   120
      TabIndex        =   3
      Top             =   6000
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Timer WaitTimer 
      Left            =   2040
      Tag             =   "0"
      Top             =   7680
   End
   Begin VB.ListBox PossibleBotsList 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   5880
      TabIndex        =   2
      Top             =   6000
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.Timer JoinChannelTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Tag             =   "0"
      Top             =   7680
   End
   Begin VB.CommandButton TestButton 
      Caption         =   "Test"
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.CommandButton ConnectButton 
      Caption         =   "Connect"
      Height          =   500
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1000
   End
   Begin MSWinsockLib.Winsock ServerConnection 
      Left            =   10320
      Top             =   7680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'all code by david_amor@hotmail.com unless stated otherwise
'this project along with the tripod website was given up for
'adoption some time ago and although one new version surfaced,
'development has not continued.

'I'd be more than happy if someone adopted ircft and continued
'development

Option Explicit
Dim DebugMode
Dim ThisLineFromServer As String
Dim CurrentChannelNumber As Integer


Private Sub AntiIdleTimer_Timer()

AntiIdleTimer.Tag = Str(Val(AntiIdleTimer.Tag) + 1)

If Val(AntiIdleTimer.Tag) = 60 Then
    If RespondedBotsList.ListCount > 0 Then
        SendToServerImmediately "privmsg " + BotSwapChannel + " :Detected" + Str(RespondedBotsList.ListCount) + " bots and" + Str(OffersList.ListCount - (RespondedBotsList.ListCount - 1)) + " offers." + Str(Int((RespondedBotsList.ListCount / PossibleBotsList.ListCount) * 100)) + "% response rate."
    End If
    AntiIdleTimer.Tag = "0"
End If

End Sub

Private Sub BroadcastBotsTimer_Timer()

BroadcastBotsTimer.Tag = Str(Val(BroadcastBotsTimer.Tag) + 1)

If Val(BroadcastBotsTimer.Tag) = 300 Then
    BroadcastBots
    BroadcastBotsTimer.Tag = "0"
End If

End Sub

Private Sub ChannelsTimer_Timer()

If UBound(Preferences.Channels) >= CurrentChannelNumber Then
    'Debug.Print "part " + Preferences.Channels(CurrentChannelNumber)
    If Preferences.ChannelListen = 0 Then CommandsList.AddItem "part " + Preferences.Channels(CurrentChannelNumber)
    CurrentChannelNumber = CurrentChannelNumber + 1
    ChannelsTimer.Enabled = False
End If

JoinChannels

End Sub

Private Sub CommandsTimer_Timer()

If ServerConnection.State = sckConnected Then
    If CommandsList.ListCount > 0 Then
        ServerConnection.SendData CommandsList.List(0) & vbCrLf
        Debug.Print CommandsList.List(0)
        CommandsList.RemoveItem 0
    End If
End If

End Sub

Private Sub ConnectButton_Click()

If ServerConnection.State = sckClosed Then
    AppendLog "Attempting connect to " + Preferences.Server
    ServerConnection.RemoteHost = Preferences.Server
    ServerConnection.RemotePort = Preferences.ServerPort
    ServerConnection.Connect
    ConnectButton.Caption = "Disconnect"
Else
    While ServerConnection.State <> sckClosed
        ServerConnection.Close
        DoEvents
    Wend
    ConnectButton.Caption = "Connect"
End If

End Sub

Private Sub ExitButton_Click()

If ServerConnection.State = sckConnected Then
    If InStr(LCase(Preferences.VersionReply), "irc filetool") = 0 Then
        ServerConnection.SendData ("QUIT :Leaving") + vbCrLf
    Else
        ServerConnection.SendData ("QUIT :" + "http://members.tripod.com/~IRC_Filetool/") + vbCrLf
    End If
    AppendLog "Logging off server. Please wait"
    While ServerConnection.State <> sckClosed
        DoEvents
    Wend
    End
Else
    End
End If

End Sub

Private Sub FindButton_Click()

FindForm.Show

End Sub

Private Sub Form_DblClick()

Select Case OffersList.FontSize
Case 6.75
    OffersList.FontSize = 8.25
    OffersFrame.Height = MainForm.Height - 2700
Case 8.25
    OffersList.FontSize = 9.75
    OffersFrame.Height = MainForm.Height - 2700
Case 9.75
    OffersList.FontSize = 6.75
    OffersFrame.Height = MainForm.Height - 2700
End Select

End Sub

Private Sub Form_Load()

DebugMode = False

Call InitialSetup

End Sub

Private Sub Form_Resize()

If DebugMode = False Then

    If MainForm.WindowState = 1 Then Exit Sub

    If MainForm.Width < 8000 Then MainForm.Width = 8000
    If MainForm.Height < 6000 Then MainForm.Height = 6000

    If MainForm.Width > 10000 Then
        LogFrame.Width = MainForm.Width - 360
        LogFrame.Top = 660
        LogList.Width = LogFrame.Width - 240
        OffersFrame.Top = 1800
        OffersFrame.Width = MainForm.Width - 360
        OffersList.Width = OffersFrame.Width - 240
        OffersFrame.Height = MainForm.Height - 2550
        OffersList.Height = OffersFrame.Height - 240
        RequestFilesButton.Left = 3360
        RequestFilesButton.Top = 120
        PreferencesButton.Left = 4440
        PreferencesButton.Top = 120
        GetURLButton.Left = 5520
        GetURLButton.Top = 120
        PartChannelsButton.Left = 6600
        PartChannelsButton.Top = 120
        HelpButton.Left = MainForm.Width - (1340 + HelpButton.Width)
        HelpButton.Top = 120
        ExitButton.Left = MainForm.Width - (240 + ExitButton.Width)
    Else
        LogFrame.Width = MainForm.Width - 360
        LogFrame.Top = 1260
        LogList.Width = LogFrame.Width - 240
        OffersFrame.Top = 2400
        OffersFrame.Width = MainForm.Width - 360
        OffersList.Width = OffersFrame.Width - 240
        OffersFrame.Height = MainForm.Height - 3150
        OffersList.Height = OffersFrame.Height - 240
        RequestFilesButton.Left = 120
        RequestFilesButton.Top = 720
        PreferencesButton.Left = 1200
        PreferencesButton.Top = 720
        GetURLButton.Left = 2280
        GetURLButton.Top = 720
        PartChannelsButton.Left = 3360
        PartChannelsButton.Top = 720
        HelpButton.Left = MainForm.Width - (240 + HelpButton.Width)
        HelpButton.Top = 720
        ExitButton.Left = MainForm.Width - (240 + ExitButton.Width)
    End If

End If

'OffersFrame.Height = 4000

End Sub

Private Sub Form_Unload(Cancel As Integer)

End

End Sub

Private Sub GetURLButton_Click()

Dim ThisWord As String
Dim i, j As Integer
Dim iret As Long

For i = 0 To (OffersList.ListCount - 1)
    If (OffersList.Selected(i) = True) And (OffersList.List(i) <> "") Then
        j = 1
        ThisWord = "null"
        Do While ThisWord <> "%error%"
            ThisWord = GetWord(OffersList.List(i), j)
            If InStr(LCase(ThisWord), "http") <> 0 Or InStr(LCase(ThisWord), "www") <> 0 Then
                If Mid(ThisWord, 1, 1) = "(" And Mid(ThisWord, Len(ThisWord), 1) = ")" Then ThisWord = Mid(ThisWord, 2, Len(ThisWord) - 2)
                If Mid(ThisWord, 1, 1) = "[" And Mid(ThisWord, Len(ThisWord), 1) = "]" Then ThisWord = Mid(ThisWord, 2, Len(ThisWord) - 2)
                If Mid(ThisWord, 1, 1) = "{" And Mid(ThisWord, Len(ThisWord), 1) = "}" Then ThisWord = Mid(ThisWord, 2, Len(ThisWord) - 2)
                iret = ShellExecute(Me.hwnd, vbNullString, ThisWord, vbNullString, "c:\", SW_SHOWNORMAL)
            End If
            j = j + 1
        Loop
    End If
Next i

End Sub

Private Sub HelpButton_Click()

HelpForm.Show

End Sub

Private Sub IdentWinsock_ConnectionRequest(ByVal requestID As Long)

Dim SendLine As String

On Error GoTo ErrorHandler

Do While IdentWinsock.State <> sckClosed
    IdentWinsock.Close
    DoEvents
Loop
IdentWinsock.Accept requestID
AppendLog ("Ident connected")
SendLine = Str(ServerConnection.LocalPort) + "," + Str(ServerConnection.RemotePort) + " : USERID : UNIX : " + Preferences.Nickname + vbCrLf
IdentWinsock.SendData SendLine
AppendLog ("Ident info dispatched")
Exit Sub

ErrorHandler:
MsgBox Error, vbOKOnly + vbInformation, "Error"

End Sub

Private Sub IdentWinsock_DataArrival(ByVal bytesTotal As Long)

Dim Data As String

On Error GoTo ErrorHandler

IdentWinsock.GetData Data, vbString

ErrorHandler:
MsgBox Error, vbOKOnly + vbInformation, "Error"

End Sub

Private Sub IdentWinsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

MsgBox "Error #" & CStr(Number) & vbCrLf & Description, vbOKOnly + vbInformation, "Error"

End Sub

Private Sub JoinChannelTimer_Timer()

JoinChannelTimer.Tag = Str(Val(JoinChannelTimer.Tag) + 1)

If Val(JoinChannelTimer.Tag) = 300 Then
    If Preferences.ChannelListen = 1 Then
        PartChannels
    End If
    CurrentChannelNumber = 1
    JoinChannels
    JoinChannelTimer.Tag = "0"
End If

End Sub

Private Sub NagForOffersButton_Click()

If NagTimer.Enabled = False Then
    Call RequestFilesButton_Click
    NagTimer.Enabled = True
    NagForOffersButton.Caption = "Stop Nagging"
Else
    NagTimer.Enabled = False
    NagForOffersButton.Caption = "Nag For Offers"
End If

End Sub

Private Sub NagTimer_Timer()

NagTimer.Tag = Str(Val(NagTimer.Tag) + 1)

If Val(NagTimer.Tag) = Int(Preferences.NagDelay / 100) Then
    Call RequestFilesButton_Click
    NagTimer.Tag = "0"
End If

End Sub

Private Sub OffersList_DblClick()

Call RequestFilesButton_Click

End Sub

Private Sub PartChannelsButton_Click()

PartChannels
PartChannelsButton.Visible = False
PrefsForm.ChanListenCheck.Value = 0
Preferences.ChannelListen = 0

End Sub



Private Sub PreferencesButton_Click()

'AppendLog "All " + LTrim(Str(PossibleBotsList.ListCount)) + " bots XDCC Listed"

MainForm.ConnectButton.Enabled = True
If Preferences.DownloadPath = "" Then
    PrefsForm.DriveSelect.Drive = "c:\"
    PrefsForm.DirSelect.Path = "c:\"
Else
    PrefsForm.DriveSelect.Drive = Preferences.DownloadPath
    PrefsForm.DirSelect.Path = Preferences.DownloadPath
End If

PrefsForm.Show

End Sub

Private Sub ReceiveConnection_Close(Index As Integer)

If Preferences.Sounds = 1 Then PlayWAV (App.Path + "\error.wav")

While ReceiveConnection(Index).State <> sckClosed
    ReceiveConnection(Index).Close
    DoEvents
Wend
'a known bug forces me to do this:
ReceiveConnection(Index).LocalPort = 0

End Sub

Private Sub ReceiveConnection_Connect(Index As Integer)

Dim ThisSlot As Integer

ThisSlot = Index

'Debug.Print "Connection request from: "; Index

ReceiveInfo(ThisSlot).StartTime = Now
ReceiveInfo(ThisSlot).Filename = AssignFilename(ReceiveInfo(ThisSlot).Filename)

ReceiveBox(ThisSlot).Caption = ReceiveInfo(ThisSlot).Filename + " (" + ReceiveInfo(ThisSlot).UsableFileLength + ")"
ReceiveBox(ThisSlot).CPSTimer.Enabled = True
ReceiveBox(ThisSlot).KickstartTimer = True
ReceiveBox(ThisSlot).StatusBar1.SimpleText = "Download from " + ReceiveInfo(ThisSlot).Nick + " initiated at" + Str(ReceiveInfo(ThisSlot).StartTime)
Open Preferences.DownloadPath + ReceiveInfo(ThisSlot).Filename For Binary Access Read Write As Index

End Sub

Private Sub ReceiveConnection_DataArrival(Index As Integer, ByVal bytesTotal As Long)

'get the data and write it to the file
Dim SendData() As Byte
Dim i As Integer
Dim RetVal As Double
Dim Hexdata As String

ReceiveConnection(Index).GetData SendData, vbByte

Put Index, , SendData

'update the receivebytes property
ReceiveInfo(Index).ReceivedBytes = ReceiveInfo(Index).ReceivedBytes + bytesTotal
ReceiveInfo(Index).LastTimeReceivedData = Now
'and then reflect the data on the progress bar
'first check that the bot isn't sending more than it said it would
If ReceiveInfo(Index).ReceivedBytes >= ReceiveBox(Index).DataProgressBar.Max Then ReceiveBox(Index).DataProgressBar.Max = ReceiveInfo(Index).ReceivedBytes
ReceiveBox(Index).DataProgressBar.Value = ReceiveInfo(Index).ReceivedBytes
ReceiveBox(Index).InfoReceivedBytes.Caption = ReadableFilesize(ReceiveInfo(Index).ReceivedBytes)

'send back acknowledgement
Hexdata = Hex(LOF(Index))
Hexdata = String$(8 - Len(Hexdata), "0") & Hexdata
ReDim SendBackData(3) As Byte
For i = 1 To Len(Hexdata) Step 2
    SendBackData((i - 1) / 2) = Val("&H" & Mid(Hexdata, i, 2))
Next
ReceiveConnection(Index).SendData SendBackData

'has it finished?
If ReceiveInfo(Index).Filelength = ReceiveInfo(Index).ReceivedBytes Then
    'close the winsock control and check it's closed
    While ReceiveConnection(Index).State <> sckClosed
        ReceiveConnection(Index).Close
        DoEvents
    Wend
    'unload the control
    Unload ReceiveConnection(Index)
    'hide the form
    ReceiveBox(Index).Hide
    'flag that it's no longer in use
    ReceiveInfo(Index).InUse = False
    'stop the timers
    ReceiveBox(Index).CPSTimer.Enabled = False
    ReceiveBox(Index).KickstartTimer.Enabled = False
    'close the file
    Close Index
    'tell the user that download was successful
    AppendLog "Successfully downloaded " + ReceiveInfo(Index).Filename + " at" + Str(ReceiveInfo(Index).CPS) + " CPS"
    'play a sound
    If Preferences.Sounds = 1 Then PlayWAV (App.Path + "\finish.wav")
    'open it?
    If Preferences.OpenOnDownload = 1 Then
        RetVal = ShellExecute(Me.hwnd, vbNullString, (Preferences.DownloadPath + ReceiveInfo(Index).Filename), vbNullString, "c:\", SW_SHOWNORMAL)
    End If
    'reset data
    ReceiveInfo(Index).CPS = 0
    ReceiveInfo(Index).ReceivedBytes = 0
    ReceiveBox(Index).InfoTimeRemaining.Caption = "Unknown"
    ReceiveBox(Index).InfoReceivedBytes.Caption = "0"
    ReceiveBox(Index).InfoCPS.Caption = "Unknown"
    ReceiveBox(Index).DataProgressBar.Value = 0
    'sndPlaySound App.Path + "\finish.wav", SND_ASYNC
End If

End Sub

Private Sub RequestFilesButton_Click()

RequestFileFromNotice

End Sub

Private Sub ServerConnection_Close()

AppendLog "Disconnected from server"

End Sub

Private Sub ServerConnection_Connect()

UsingNickname = Preferences.Nickname
SendToServerImmediately "nick " + Preferences.Nickname
SendToServerImmediately "user " + Preferences.Nickname + " " + ServerConnection.LocalHostName + " " + ServerConnection.RemoteHost & " :" + NameAndVersion
    
CommandsTimer.Enabled = True
    
End Sub

Private Sub ServerConnection_DataArrival(ByVal bytesTotal As Long)

'shamefully ripped from bodebot - http://www.felmlee.com/bodebot/default.asp

Dim s$, buf$, a%, i%, ii%
Dim endline$
buf$ = String$(1024, " ")
ServerConnection.GetData buf$, vbString, 1024
a% = Len(buf$)
endline$ = vbLf
If a% > 0 Then
    'SockReadBuffer$ = SockReadBuffer$ + RTrim$(buf$)
    ThisLineFromServer = ThisLineFromServer + Left$(buf$, a%)
    While InStr(ThisLineFromServer, endline$) <> 0
        i = InStr(ThisLineFromServer, endline$)
        If i <> 0 Then
            If i < Len(ThisLineFromServer) Then
                s$ = Left$(ThisLineFromServer, i - 1)
                ThisLineFromServer = Mid$(ThisLineFromServer, i + 1)
                If InStr(s$, Chr$(13)) Then
                    s$ = Left$(s$, InStr(s$, Chr$(13)) - 1)
                ElseIf InStr(s$, Chr$(10)) Then
                    s$ = Left$(s$, InStr(s$, Chr$(10)) - 1)
                End If
                ServerDataList.AddItem s$
                'Debug.Print "|"; s$; "|"
            Else
                s$ = ThisLineFromServer
                ThisLineFromServer = ""
                If InStr(s$, Chr$(13)) Then
                    s$ = Left$(s$, InStr(s$, Chr$(13)) - 1)
                ElseIf InStr(s$, Chr$(10)) Then
                    s$ = Left$(s$, InStr(s$, Chr$(10)) - 1)
                End If
                ServerDataList.AddItem s$
                'Debug.Print "|"; s$; "|"
            End If
        End If
    Wend
End If


'Static PreProcessedString As String
'Dim i, ThisASC, CRPosition, LFPosition As Integer
'Dim ThisString As String

'ServerConnection.GetData PreProcessedString, vbString, 1024 'Get the data when it comes
'DoEvents
'For i = 1 To Len(PreProcessedString)
'    DoEvents
'    If Mid(PreProcessedString, i, 1) = vbCr Or Mid(PreProcessedString, i, 1) = vbLf Then
'        ServerDataList.AddItem ThisLineFromServer
'        ThisLineFromServer = ""
'    Else
'        ThisASC = Asc(Mid(PreProcessedString, i, 1))
'        If (ThisASC <> 2) And (ThisASC <> 31) And (ThisASC <> 22) And (ThisASC <> 3) And (ThisASC <> 15) And (ThisASC <> 38) And (ThisASC <> 13) Then
'            ThisLineFromServer = ThisLineFromServer & Mid(PreProcessedString, i, 1)
'        End If
'    End If
'Next

'CRPosition = InStr(PreProcessedString, vbCr)
'LFPosition = InStr(PreProcessedString, vbLf)

'If CRPosition <> 0 Then
'    ThisLineFromServer = ThisLineFromServer + Left(PreProcessedString, CRPosition - 1)
'    ParseServerData ThisLineFromServer
'    ThisLineFromServer = ""
'    PreProcessedString = Right(PreProcessedString, Len(PreProcessedString) - CRPosition)
'End If
'If LFPosition <> 0 Then
'    ThisLineFromServer = ThisLineFromServer + Left(PreProcessedString, LFPosition - 1)
'    ParseServerData ThisLineFromServer
'    ThisLineFromServer = ""
'    PreProcessedString = Right(PreProcessedString, Len(PreProcessedString) - LFPosition)
'End If
'ThisLineFromServer = ThisLineFromServer + PreProcessedString

End Sub

Private Sub ServerConnection_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

MsgBox "Error #" & CStr(Number) & vbCrLf & Description, vbOKOnly + vbInformation, "Error"
AppendLog "Error: " & Description
Call ConnectButton_Click

End Sub

Private Sub ParseServerData(ThisLine As String)

Dim Word1, Word2, Word3, Word4, Word5 As String

'Debug.Print ThisLine

Word1 = GetWord(ThisLine, 1)
Word2 = GetWord(ThisLine, 2)
'Word3 = GetWord(ThisLine, 3)
Word4 = GetWord(ThisLine, 4)
'Word5 = GetWord(ThisLine, 5)

Select Case Word1
    Case "PING"
        SendToServerImmediately "PONG " & Preferences.Server
        Call AppendLog("Ping? Pong! " & Now)
        Exit Sub
    Case "ERROR"
        Call LoginError(ThisLine)
        Exit Sub
End Select

Select Case Word2
    Case "001"
        'successfully logged in
        Call LoginSuccess
        Exit Sub
    Case "353"
        'list of nicks in channel
        Call Process353(ThisLine)
        Exit Sub
    Case "473", "461", "475", "471", "476", "403", "405", "474"
        'join response
        Call JoinResponse(ThisLine)
        Exit Sub
    Case "433", "436", "432"
        'bad nick
        Call BadNick(ThisLine)
        Exit Sub
        Case "401"
        'no such nick
        Call NoSuchNick(ThisLine)
        Exit Sub
    Case "366"
        'end of names list
        Call Process366(ThisLine)
        Exit Sub
    Case "311"
        'whois
        Call WhoIs(ThisLine)
        Exit Sub
    Case "NOTICE"
        Call ProcessNotice(ThisLine)
    Case "PART"
        Call PartChannel(ThisLine)
        Exit Sub
    Case "PRIVMSG"
        If Preferences.ChannelListen = 1 Then Call ProcessPrivmsg(ThisLine)
    Case "KICK"
        'you were kicked
        Call KickedFromChannel(ThisLine)
        Exit Sub
End Select

Select Case Word4
    Case ":VERSION"
        Call VersionRequest(ThisLine)
    Case ":REALVERSION"
        Call RealVersionRequest(ThisLine)
    Case ":FINGER"
        Call FingerRequest(ThisLine)
    Case ":USERINFO"
        Call UserinfoRequest(ThisLine)
    Case ":USERINFO_REG"
        Call RegUserinfoRequest(ThisLine)
    Case ":TIME"
        Call TimeRequest(ThisLine)
    Case ":PING"
        Call PingRequest(ThisLine)
    Case ":DCC"
        Call RequestSend(ThisLine)
    Case ":HeresMyBots:"
        If Preferences.Botshare = 1 Then Call ProcessBotsBroadcast(ThisLine)
End Select

End Sub

Private Sub JoinChannels()

If UBound(Preferences.Channels) >= CurrentChannelNumber Then
    'Debug.Print "join " + Preferences.Channels(CurrentChannelNumber)
    CommandsList.AddItem "join " + Preferences.Channels(CurrentChannelNumber)
    ChannelsTimer.Enabled = True
End If

End Sub

Private Sub PartChannels()

Dim i As Integer

For i = 0 To PrefsForm.ChannelList.ListCount - 1
    CommandsList.AddItem "part " + PrefsForm.ChannelList.List(i)
Next i

End Sub

Private Sub Process353(ThisLine As String)

Dim i, j, StartPosition As Integer
Dim ThisNick As String
Dim found As Boolean

StartPosition = InStr(2, ThisLine, ":") + 1
For i = StartPosition To Len(ThisLine)
    If Mid(ThisLine, i, 1) = " " Then
        ThisNick = LCase(ThisNick)
        If Left(ThisNick, 1) = "@" Then ThisNick = Right(ThisNick, Len(ThisNick) - 1)
        If Left(ThisNick, 1) = "+" Then ThisNick = Right(ThisNick, Len(ThisNick) - 1)
        'is it a bot?
        If InStr(1, ThisNick, "dcc") <> 0 Then
            'do I already have it?
            found = False
            For j = 0 To PossibleBotsList.ListCount - 1
                If ThisNick = PossibleBotsList.List(j) Then found = True
            Next j
            If found = False Then
                PossibleBotsList.AddItem ThisNick
                'If Preferences.ChannelListen = 0 Then
                    If Preferences.UseMSG = True Then
                        CommandsList.AddItem "privmsg " + ThisNick + " :xdcc list"
                    Else
                        CommandsList.AddItem "privmsg " + ThisNick + " :" + Chr$(1) + "xdcc list" + Chr$(1)
                    End If
                'End If
            End If
        End If
        ThisNick = ""
    Else
        ThisNick = ThisNick & Mid(ThisLine, i, 1)
    End If
Next i

End Sub

Private Sub WaitFor(HundrethsOfSeconds)

WaitTimer.Interval = HundrethsOfSeconds * 10
WaitTimer.Enabled = True
    
While WaitTimer.Enabled = True
    DoEvents
Wend

End Sub

Private Sub ServerDataChecker_Timer()

Dim StrippedString As String
Dim i As Integer
Dim ThisASC As Integer

If ServerDataList.ListCount > 0 Then
    For i = 1 To Len(ServerDataList.List(0))
        ThisASC = Asc(Mid(ServerDataList.List(0), i, 1))
        If (ThisASC <> 2) And (ThisASC <> 31) And (ThisASC <> 22) And (ThisASC <> 3) And (ThisASC <> 15) And (ThisASC <> 38) And (ThisASC <> 13) Then
            StrippedString = StrippedString & Mid(ServerDataList.List(0), i, 1)
        End If
    Next i
    ParseServerData StrippedString
    ServerDataList.RemoveItem (0)
End If

End Sub

Private Sub ShouldIBeepTimer_Timer()

ShouldIBeepTimer.Tag = Str(Val(ShouldIBeepTimer.Tag) + 1)

If Val(ShouldIBeepTimer.Tag) = 600 Then
    ShouldIBeep = True
    ShouldIBeepTimer.Enabled = False
End If

End Sub

Private Sub TestButton_Click()

Dim MyPath As String
Dim MyName As String

MyPath = Preferences.DownloadPath   ' Set the path.
MyName = Dir(MyPath, vbDirectory)   ' Retrieve the first entry.
Do While MyName <> ""   ' Start the loop.
   ' Ignore the current directory and the encompassing directory.
   If MyName <> "." And MyName <> ".." Then
      ' Use bitwise comparison to make sure MyName is a directory.
      If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
         AddThisDirectorysFiles (MyName)  ' Display entry only if it
      End If   ' it represents a directory.
   End If
   MyName = Dir   ' Get next entry.
Loop

End Sub
Private Sub AddThisDirectorysFiles(ThisDir)

Dim MyPath As String
Dim MyName As String

MyPath = ThisDir   ' Set the path.
MyName = Dir(MyPath, vbNormal)   ' Retrieve the first entry.
Do While MyName <> ""   ' Start the loop.
   ' Ignore the current directory and the encompassing directory.
   If MyName <> "." And MyName <> ".." Then
      ' Use bitwise comparison to make sure MyName is a directory.
      If (GetAttr(MyPath & MyName) And vbNormal) = vbNormal Then
         MyFilesList.AddItem MyName   ' Display entry only if it
      End If   ' it represents a directory.
   End If
   MyName = Dir   ' Get next entry.
Loop

End Sub

Private Sub WaitBeforeAskingTimer_Timer()

Dim i As Integer

For i = 0 To WaitBotsList.ListCount - 1
    If Preferences.UseMSG = True Then
        CommandsList.AddItem "privmsg " & WaitBotsList.List(i) & " :xdcc list"
    Else
        CommandsList.AddItem "privmsg " & WaitBotsList.List(i) & " :" + Chr$(1) + "xdcc list" + Chr$(1)
    End If
Next i
WaitBeforeAskingTimer.Enabled = False
WaitBotsList.Clear

End Sub

Private Sub WaitTimer_Timer()

WaitTimer.Enabled = False

End Sub

Private Sub ProcessNotice(ThisLine As String)

Dim ThisNick As String
Dim ThisOffer As String
Dim NickPadder, OfferLine As String
Dim i, j, PositionToPlace As Integer
Dim found, FoundPlace As Boolean

'extract the useful parts
If InStr(ThisLine, "!") <> 0 Then
    ThisNick = Mid(ThisLine, 2, InStr(ThisLine, "!") - 2)
Else
    ThisNick = "%IRCAdmn%"
End If
NickPadder = Left("          ", 10 - Len(ThisNick))
ThisOffer = Right(ThisLine, Len(ThisLine) - InStr(2, ThisLine, ":"))
'get rid of the leading space
If Len(ThisOffer) > 1 Then
    If Mid(ThisOffer, 1, 1) = " " And Mid(ThisOffer, 2, 1) <> " " Then
        ThisOffer = Right(ThisOffer, Len(ThisOffer) - 1)
    End If
End If
OfferLine = ThisNick & NickPadder & ThisOffer

'does the response come from a possible bot?
found = False
For i = 0 To PossibleBotsList.ListCount - 1
    If LCase(ThisNick) = PossibleBotsList.List(i) Then found = True
Next i

If found = True Then
    Call PlaceOffer(OfferLine, ThisNick)
End If

'does the response come from a bot we've asked for a pack from?
found = False
For i = 0 To PossibleBotsList.ListCount - 1
    If LCase(ThisNick) = RequestedBotsList.List(i) Then found = True
Next i
If found = True Then AppendLog (OfferLine)

End Sub

Private Sub BroadcastBots()

Dim BroadcastLine(50) As String
Dim i As Integer

If RespondedBotsList.ListCount > 0 Then
    For i = 0 To RespondedBotsList.ListCount - 1
        If BroadcastLine(Int(i / 10)) = "" Then BroadcastLine(Int(i / 10)) = "HeresMyBots:"
        BroadcastLine(Int(i / 10)) = BroadcastLine(Int(i / 10)) + " " + RespondedBotsList.List(i)
    Next i
End If

For i = 0 To 50
    If BroadcastLine(i) <> "" Then
        SendToServerImmediately "privmsg " + BotSwapChannel + " :" + BroadcastLine(i)
        WaitFor (50)
    End If
Next i

End Sub

Private Sub ProcessBotsBroadcast(ThisLine As String)

Dim ThisWord As String
Dim i, j, X As Integer
Dim found As Boolean

X = 5
ThisWord = GetWord(ThisLine, 5)
Do While ThisWord <> "%error%"
    found = False
    For j = 0 To PossibleBotsList.ListCount
        If LCase(PossibleBotsList.List(j)) = LCase(ThisWord) Then found = True
    Next j
    If found = False Then
        PossibleBotsList.AddItem ThisWord
        WaitBotsList.AddItem ThisWord
        Randomize
        WaitBeforeAskingTimer.Interval = Int((60000 * Rnd) + 1)
        WaitBeforeAskingTimer.Enabled = True
    End If
    X = X + 1
    ThisWord = GetWord(ThisLine, X)
Loop

End Sub

Private Sub RequestFileFromNotice()

Dim FoundPack, found As Boolean
Dim i, j, k, ThisASC As Integer
Dim ThisPack, ThisOffer, ThisCharacter, ThisNick As String

For i = 0 To (OffersList.ListCount - 1)
    If (OffersList.Selected(i) = True) And (OffersList.List(i) <> "") Then
        'work out the pack #
        FoundPack = False
        ThisPack = ""
        ThisOffer = Right(OffersList.List(i), Len(OffersList.List(i)) - 10)
        
        j = 1
        Do
            ThisCharacter = Mid(ThisOffer, j, 1)
            ThisASC = Asc(ThisCharacter)
            If ((ThisASC > 47) And (ThisASC < 58)) Then
                FoundPack = True
                'check to see if it's a > 1 digit pack #
                Do
                    ThisPack = ThisPack + ThisCharacter
                    j = j + 1
                    ThisCharacter = Mid(ThisOffer, j, 1)
                    ThisASC = Asc(ThisCharacter)
                Loop Until ((ThisASC < 47) Or (ThisASC > 58) Or (j = Len(ThisOffer)))
            End If
            j = j + 1
        Loop Until (FoundPack = True) Or (j >= Len(ThisOffer))
    
        ThisNick = GetWord(OffersList.List(i), 1)
        
        'add this nick to RequestedBotsList (which is bots we want to hear from)
        found = False
        For k = 0 To RequestedBotsList.ListCount - 1
            If RequestedBotsList.List(k) = ThisNick Then found = True
        Next k
        If found = False Then RequestedBotsList.AddItem LCase(ThisNick)
                
        If ServerConnection.State <> sckClosed Then AppendLog "Requesting " + ThisOffer
        If Preferences.UseMSG = True Then
            SendToServerImmediately "privmsg " + ThisNick + " :xdcc send #" + ThisPack
        Else
            SendToServerImmediately "privmsg " + ThisNick + " :" + Chr$(1) + "xdcc send #" + ThisPack + Chr$(1)
        End If
        WaitFor (100)
        
     End If
     
    Next i
    
End Sub

Private Sub RequestSend(ThisLine As String)

Dim RecieveSlot As Integer
Dim SenderNick As String
Dim SenderFilename As String
Dim SenderIP As String
Dim SenderPort As Long
Dim SenderFileLength As Long
Dim i, ReceiveSlot As Integer
Dim found As Boolean
Dim SendString As String

SenderNick = Mid(ThisLine, 2, InStr(1, ThisLine, "!") - 2)

'check it's from a nick you requested something from
found = False
For i = 0 To RequestedBotsList.ListCount - 1
    If LCase(SenderNick) = LCase(RequestedBotsList.List(i)) Then found = True
Next i
If found = False Then
    AppendLog "Unauthorised DCC send detected from " & SenderNick
    Exit Sub
End If

Select Case GetWord(ThisLine, 5)
    Case "CHAT"
        AppendLog "Unauthorised DCC chat detected from " & SenderNick
        Exit Sub
    Case "ACCEPT"
        AppendLog "DCC resume is not supported"
        Exit Sub
    Case "RESUME"
        AppendLog "DCC resume is not supported"
        Exit Sub
    Case "REJECT"
        AppendLog "DCC error"
        Exit Sub
End Select

SenderFilename = GetWord(ThisLine, 6)
If InStr(SenderFilename, "\") <> 0 Then SenderFilename = "BadFilename"
If InStr(SenderFilename, "/") <> 0 Then SenderFilename = "BadFilename"

SenderIP = DecodeIP(GetWord(ThisLine, 7))
SenderPort = Val(GetWord(ThisLine, 8))
SenderFileLength = Val(GetWord(ThisLine, 9))
If SenderFileLength = 0 Then SenderFileLength = 1

If Preferences.DCCRelay = 1 Then
    SendString = "PRIVMSG " & Preferences.DCCRelayNick & " :DCC SEND " & SenderFilename & Chr(32) & GetWord(ThisLine, 7) & Chr(32) & GetWord(ThisLine, 8) & Chr(32) & GetWord(ThisLine, 9) & ""
    SendToServerImmediately SendString
    SendToServerImmediately "notice " & Preferences.DCCRelayNick & " :DCC relaying from " + SenderNick
    AppendLog SenderFilename + " from " + SenderNick + " auto-forwarded to " + Preferences.DCCRelayNick
    Exit Sub
End If

'check directory validity?

'find a place to put it
ReceiveSlot = GetFreeSlot

SenderFilename = AssignFilename(SenderFilename)
    
ReceiveInfo(ReceiveSlot).Resumed = False
ReceiveInfo(ReceiveSlot).Nick = SenderNick
ReceiveInfo(ReceiveSlot).Filename = SenderFilename
ReceiveInfo(ReceiveSlot).IP = SenderIP
ReceiveInfo(ReceiveSlot).Port = SenderPort
ReceiveInfo(ReceiveSlot).Filelength = SenderFileLength
ReceiveInfo(ReceiveSlot).ReceivedBytes = 0
ReceiveInfo(ReceiveSlot).InUse = True
ReceiveInfo(ReceiveSlot).StartTime = Now

'work out the size of the file so that it can be displayed
ReceiveInfo(ReceiveSlot).UsableFileLength = ReadableFilesize(SenderFileLength)

'alter the receive box
ReceiveBox(ReceiveSlot).Caption = ReceiveInfo(ReceiveSlot).Filename + " (" + ReceiveInfo(ReceiveSlot).UsableFileLength + ")"
If SenderFileLength > 0 Then
    ReceiveBox(ReceiveSlot).DataProgressBar.Max = SenderFileLength
Else
    ReceiveBox(ReceiveSlot).DataProgressBar.Max = 1000000
End If
ReceiveBox(ReceiveSlot).CPSTimer.Enabled = True
ReceiveBox(ReceiveSlot).Tag = Str(ReceiveSlot)
ReceiveBox(ReceiveSlot).KickstartTimer.Enabled = True
ReceiveBox(ReceiveSlot).StatusBar1.SimpleText = "Download from " + ReceiveInfo(ReceiveSlot).Nick + " initiated at" + Str(ReceiveInfo(ReceiveSlot).StartTime)
ReceiveBox(ReceiveSlot).InfoReceivedBytes = ReadableFilesize(ReceiveInfo(ReceiveSlot).ReceivedBytes)
    
ReceiveBox(ReceiveSlot).Show

'got overflow message here once
'connect to it
Call ConnectToSender(ReceiveSlot, SenderIP, Val(SenderPort))

End Sub

Private Sub InitialSetup()

'RequestedBotsList.AddItem "Rice2"
'PossibleBotsList.AddItem "Rice2"

On Error GoTo ErrorHandler

RealNameAndVersion = "IRC Filetool v1.2"
'BotSwapChannel = "#ircft-botswap"
BotSwapChannel = Chr(35) + Chr(105) + Chr(114) + Chr(99) + Chr(102) + Chr(116) + Chr(45) + Chr(98) + Chr(111) + Chr(116) + Chr(115) + Chr(119) + Chr(97) + Chr(112)
IdentWinsock.LocalPort = 113
IdentWinsock.Listen

ReDim Preserve ReceiveInfo(10) As ReceiveInfoStructure
ReDim Preserve ReceiveBox(10) As New ReceiveBoxForm
CurrentChannelNumber = 1

ReadINI

MainForm.Caption = RealNameAndVersion + ""

ErrorHandler:
If Err.Number = 10048 Then
    AppendLog ("Ident port in use")
    Resume Next
End If

End Sub

Private Sub ConnectToSender(ReceiveSlot As Integer, SenderIP As String, SenderPort As Long)

'On Error GoTo bad_details

Load ReceiveConnection(ReceiveSlot)
ReceiveConnection(ReceiveSlot).RemoteHost = SenderIP
ReceiveConnection(ReceiveSlot).RemotePort = SenderPort
ReceiveConnection(ReceiveSlot).Connect

GoTo EndOfSub

BadDetails:
ReceiveBox(ReceiveSlot).InfoTimeRemaining = "Bad Details!"
ReceiveBox(ReceiveSlot).InfoReceivedBytes = "Bad Details!"
ReceiveBox(ReceiveSlot).InfoCPS = "Bad Details!"

EndOfSub:
End Sub

Private Function RipOutNumbers(ThisLine As String)

Dim i As Integer
Dim RippedString As String
Dim ShouldIStrip As Boolean

ShouldIStrip = False
If InStr(LCase(ThisLine), "offered") <> 0 Then ShouldIStrip = True
If InStr(LCase(ThisLine), "leeched") <> 0 Then ShouldIStrip = True
If InStr(LCase(ThisLine), "queue") <> 0 Then ShouldIStrip = True
If InStr(LCase(ThisLine), "slots") <> 0 Then ShouldIStrip = True
If InStr(LCase(ThisLine), "record") <> 0 Then ShouldIStrip = True
If InStr(LCase(ThisLine), "gets") <> 0 Then ShouldIStrip = True
If InStr(GetWord(ThisLine, 3), "x") <> 0 Then ShouldIStrip = True

If ShouldIStrip = True Then
    For i = 1 To Len(ThisLine)
        If (Asc(Mid(ThisLine, i, 1)) > 57) Or (Asc(Mid(ThisLine, i, 1)) < 48) Then
            If Asc(Mid(ThisLine, i, 1)) <> 32 Then RippedString = RippedString + Mid(ThisLine, i, 1)
        End If
    Next i
Else
    RippedString = ThisLine
End If

RipOutNumbers = RippedString

End Function

Private Sub ProcessPrivmsg(ThisLine As String)

Dim ThisNick As String
Dim ThisOffer As String
Dim NickPadder, OfferLine As String
Dim i As Integer
Dim found As Boolean

'check that it's a privmsg to a channel
If Left(GetWord(ThisLine, 3), 1) <> "#" Then Exit Sub

'extract the useful parts
ThisNick = Mid(ThisLine, 2, InStr(ThisLine, "!") - 2)
NickPadder = Left("          ", 10 - Len(ThisNick))
ThisOffer = LTrim(Right(ThisLine, Len(ThisLine) - InStr(2, ThisLine, ":")))
'get rid of the "% "
If Len(ThisOffer) > 1 Then
    If Mid(ThisOffer, 1, 2) = "% " Then
        ThisOffer = Right(ThisOffer, Len(ThisOffer) - 2)
    End If
End If

OfferLine = ThisNick & NickPadder & ThisOffer

'does the response come from a possible bot?
found = False
For i = 0 To PossibleBotsList.ListCount - 1
    If LCase(ThisNick) = PossibleBotsList.List(i) Then found = True
Next i

If found = True Then Call PlaceOffer(OfferLine, ThisNick)

End Sub

Private Sub BadNick(ThisLine As String)

Dim ErrorCode As String

ErrorCode = GetWord(ThisLine, 2)

Select Case ErrorCode
    Case "433"
        AppendLog ("Duplicate nickname. Generating random nickname")
    Case "436"
        AppendLog ("Nickname collision. Generating random nickname")
    Case "432"
        AppendLog ("Invalid nickname. Generating random nickname")
End Select

Call CreateNick

End Sub

Private Function CreateNick()

Dim MyNick As String

Randomize
MyNick = "ift-"
MyNick = MyNick + Chr(Int((25 * Rnd) + 97))
MyNick = MyNick + Chr(Int((25 * Rnd) + 97))
MyNick = MyNick + Chr(Int((25 * Rnd) + 97))
MyNick = MyNick + Chr(Int((25 * Rnd) + 97))
UsingNickname = MyNick
CommandsList.AddItem ("nick " + MyNick)

End Function

Private Sub LoginSuccess()

AppendLog ("Successfully connected to server using nick " + UsingNickname)

SendToServerImmediately "join " + BotSwapChannel
SendToServerImmediately "mode " + BotSwapChannel + " +s"

ShouldIBeep = False
ShouldIBeepTimer.Enabled = True
JoinChannels
JoinChannelTimer.Tag = "0"
JoinChannelTimer.Enabled = True
BroadcastBotsTimer.Enabled = True
AntiIdleTimer.Enabled = True

End Sub

Private Sub JoinResponse(ThisLine As String)

Dim JoinResponseCode As String
Dim ThisChannel As String

JoinResponseCode = GetWord(ThisLine, 2)
ThisChannel = GetWord(ThisLine, 4)
        
If LCase(ThisChannel) = BotSwapChannel Then Exit Sub
        
Select Case JoinResponseCode
    Case "474"
        AppendLog "Banned from channel " + ThisChannel
    Case "473"
        AppendLog "Invite only channel " + ThisChannel
    Case "471"
        AppendLog "Channel " + ThisChannel + " is full"
    Case "405"
        AppendLog "Error joining " + ThisChannel + " Too many channels"
    Case "461", "475", "476", "403"
        AppendLog "Error joining " + ThisChannel
End Select

End Sub

Private Sub PlaceOffer(ThisOffer As String, ThisNick As String)

Dim i, j, PositionToPlace As Integer
Dim found, DetectedIdentical As Boolean

'if we don't have it already, add this nick to the RespondedBotsList
found = False
For i = 0 To RespondedBotsList.ListCount - 1
    If RespondedBotsList.List(i) = LCase(ThisNick) Then found = True
Next i
If found = False Then RespondedBotsList.AddItem LCase(ThisNick)
    
'find position to place
PositionToPlace = 0
DetectedIdentical = False
For i = 0 To OffersList.ListCount
    If RTrim(Left(OffersList.List(i), 9)) = ThisNick Then
        j = 0
        Do While OffersList.List(i + j) <> ""
            'the line below doesn't quite work
            'If RipOutNumbers(OffersList.List(i + j)) = RipOutNumbers(ThisOffer) Then DetectedIdentical = True
            j = j + 1
        Loop
        PositionToPlace = i + j
    End If
Next i
'has it hasn't found another offer the same bot...
If PositionToPlace = 0 Then
    If OffersList.ListCount = 0 Then
        PositionToPlace = OffersList.ListCount
    Else
        OffersList.AddItem ""
        PositionToPlace = OffersList.ListCount
    End If
    If ShouldIBeep = True And Preferences.Sounds = 1 Then Call PlayWAV(App.Path + "\newbot.wav")
End If
    
If DetectedIdentical = False Then
    OffersList.AddItem ThisOffer, PositionToPlace
    StatusBar.SimpleText = "Detected" + Str(RespondedBotsList.ListCount) + " bots and" + Str(OffersList.ListCount - (RespondedBotsList.ListCount - 1)) + " offers." + Str(Int((RespondedBotsList.ListCount / PossibleBotsList.ListCount) * 100)) + "% response rate."
End If

End Sub

Private Sub LoginError(ThisLine As String)

Dim Reason As String

If InStr(ThisLine, NameAndVersion) <> 0 Then
    AppendLog ("Successfully disconnected from ") + Preferences.Server
    ServerConnection.Close
Else
    If InStr(ThisLine, "(") <> 0 Then
        Reason = Mid(ThisLine, InStr(ThisLine, "(") + 1, Len(ThisLine) - InStr(ThisLine, "(") - 1)
    Else
        Reason = ThisLine
    End If
    AppendLog "Disconnected: " + Reason
    While ServerConnection.State <> sckClosed
        ServerConnection.Close
        DoEvents
    Wend
    ConnectButton.Caption = "Connect"
End If

End Sub

Private Sub KickedFromChannel(ThisLine As String)

Dim KickedNick, KickerNick, Channel, Reason As String

':BatyCoda!tez@scorpio.dmv.com KICK #exceed oc_aybn :banned: DNR: Narks SUCK!

'check that the kick isn't for someone else
KickedNick = GetWord(ThisLine, 4)
If KickedNick <> UsingNickname Then Exit Sub

KickerNick = Mid(ThisLine, 2, InStr(1, ThisLine, "!") - 2)
Channel = GetWord(ThisLine, 3)

Reason = Right(ThisLine, Len(ThisLine) - InStr(2, ThisLine, ":"))
AppendLog "Kicked from " + Channel + " by " + KickerNick + ". " + Reason

End Sub

Private Sub VersionRequest(ThisLine As String)

Dim Nickname As String

Nickname = Mid(ThisLine, 2, InStr(1, ThisLine, "!") - 2)
SendToServerImmediately ("NOTICE " & Nickname & " :" & Chr$(1) & "VERSION" & " " & NameAndVersion & Chr$(1))

End Sub
Private Sub RealVersionRequest(ThisLine As String)

Dim Nickname As String

Nickname = Mid(ThisLine, 2, InStr(1, ThisLine, "!") - 2)
SendToServerImmediately ("NOTICE " & Nickname & " :" & Chr$(1) & "VERSION" & " " & RealNameAndVersion & Chr$(1))

End Sub

Private Sub UserinfoRequest(ThisLine As String)

Dim Nickname As String

Nickname = Mid(ThisLine, 2, InStr(1, ThisLine, "!") - 2)
SendToServerImmediately ("NOTICE " & Nickname & " :" & Chr$(1) & "USERINFO" + " " + UsingNickname + " (" + UsingNickname + ") " + "Idle" + Str(Int(20 * Rnd) + 1) + " seconds" + Chr$(1))

End Sub
Private Sub RegUserinfoRequest(ThisLine As String)

Dim Nickname As String

Nickname = Mid(ThisLine, 2, InStr(1, ThisLine, "!") - 2)
SendToServerImmediately ("NOTICE " & Nickname & " :" & Chr$(1) & "USERINFO" + " " + GetRegInfo("DefName", "Software\Microsoft\MS Setup (ACME)\User Info"))
SendToServerImmediately ("NOTICE " & Nickname & " :" & Chr$(1) & "USERINFO" + " " + GetRegInfo("DefCompany", "Software\Microsoft\MS Setup (ACME)\User Info"))

End Sub

Private Sub PingRequest(ThisLine As String)

Dim Nickname As String
Dim ThisNumber As String

ThisNumber = Left(GetWord(ThisLine, 5), Len(GetWord(ThisLine, 5)) - 1)
Nickname = Mid(ThisLine, 2, InStr(1, ThisLine, "!") - 2)
SendToServerImmediately ("NOTICE " & Nickname & " :" & Chr$(1) & "PING" + " " + ThisNumber + Chr$(1))

':Rice2!~amo@s0.bbnplanet.net PRIVMSG ircft :PING 907856883


End Sub

Private Sub TimeRequest(ThisLine As String)

Dim Nickname, ThisTime As String

ThisTime = Format(Date, "ddd mmm dd ") + Format(Time, "hh:mm:ss ") + Format(Date, "yyyy")
Nickname = Mid(ThisLine, 2, InStr(1, ThisLine, "!") - 2)
SendToServerImmediately ("NOTICE " & Nickname & " :" & Chr$(1) & "TIME" + " " + ThisTime + Chr$(1))

End Sub
Private Sub FingerRequest(ThisLine As String)

Dim Nickname As String

Nickname = Mid(ThisLine, 2, InStr(1, ThisLine, "!") - 2)
SendToServerImmediately ("NOTICE " & Nickname & " :" & Chr$(1) & "FINGER" + " " + UsingNickname + " (" + UsingNickname + ") " + "Idle" + Str(Int(20 * Rnd) + 1) + " seconds" + Chr$(1))

End Sub

Private Function GetTodaysPwd()

Dim Pwd As String
Dim i, j, k As String
Dim NumPart As String

i = Month(Now)
j = Day(Now)
k = Weekday(Now)

NumPart = Right(Str((((i Xor j) + k) * (j + 19)) Xor 21), Len(Str((((i Xor j) + k) * (j + 19)) Xor 21)) - 1)
'Pwd = "Ldf"
Pwd = Chr(76) + Chr(100) + Chr(102)
Pwd = Pwd + Chr(Int(78 + (i * k * j) / 155))
Pwd = Pwd + Chr(Int(72 + (i * k * j) / 150))
Pwd = Pwd + Chr(Int(65 + (i * k * j) / 101))
Pwd = Pwd + NumPart
Pwd = Pwd + "b"

GetTodaysPwd = Pwd

End Function

Private Sub NoSuchNick(ThisLine As String)

AppendLog GetWord(ThisLine, 4) + " is not on IRC"

End Sub

Private Sub WhoIs(ThisLine As String)

AppendLog GetWord(ThisLine, 4) + " is on IRC"

End Sub

Private Sub PartChannel(ThisLine As String)

Dim ThisNick As String

':zeit!~zeit@4-21-39-65.com PART #mpeg3files
ThisNick = Mid(ThisLine, 2, InStr(1, ThisLine, "!") - 2)

If LCase(ThisNick) = LCase(UsingNickname) Then
    If InStr(ThisLine, BotSwapChannel) = 0 Then
        AppendLog "Left channel " + GetWord(ThisLine, 3)
    End If
End If

End Sub

Private Sub Process366(ThisLine As String)

':irc.best.net 366 zeit #ircft-botswap :End of /NAMES list.

If InStr(ThisLine, BotSwapChannel) = 0 Then
    AppendLog "Successfully joined " + GetWord(ThisLine, 4)
End If

End Sub
