VERSION 5.00
Begin VB.Form PrefsForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "IRC Filetool Preferences"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame7 
      Caption         =   "DCC Relay"
      Height          =   915
      Left            =   4080
      TabIndex        =   35
      Top             =   3960
      Width           =   3495
      Begin VB.CommandButton DCCRelayNickOnlineButton 
         Caption         =   "Online?"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2580
         TabIndex        =   39
         Top             =   360
         Width           =   795
      End
      Begin VB.TextBox DCCRelayNickText 
         BackColor       =   &H80000013&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   37
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox DCCRelayCheck 
         Caption         =   "Use DCC Relay"
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label7 
         Caption         =   "Nickname"
         Height          =   195
         Left            =   1320
         TabIndex        =   38
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Delays (Hundreths of seconds)"
      Height          =   1095
      Left            =   120
      TabIndex        =   30
      Top             =   4800
      Width           =   3800
      Begin VB.TextBox TimeBetweenCommandsText 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         TabIndex        =   34
         Text            =   "250"
         Top             =   600
         Width           =   620
      End
      Begin VB.TextBox TimeBetweenNagsText 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         TabIndex        =   33
         Text            =   "3000"
         Top             =   240
         Width           =   620
      End
      Begin VB.Label Label6 
         Caption         =   "Time Between Server Commands"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Time Between Nags"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Client Details"
      Height          =   975
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   3800
      Begin VB.TextBox Nickname 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   25
         Top             =   480
         Width           =   1140
      End
      Begin VB.TextBox VersionText 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Text            =   "IRC Filetool v1.2"
         Top             =   480
         Width           =   2300
      End
      Begin VB.Label Label4 
         Caption         =   "Nickname"
         Height          =   255
         Left            =   2520
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Version Reply"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton WriteLogButton 
      Caption         =   "Save Log"
      Height          =   400
      Left            =   4080
      TabIndex        =   20
      Top             =   4980
      Width           =   1300
   End
   Begin VB.TextBox DownloadPath 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7080
      TabIndex        =   19
      Top             =   8040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton WriteOffersButton 
      Caption         =   "Save Offers"
      Height          =   400
      Left            =   4080
      TabIndex        =   18
      Top             =   5460
      Width           =   1300
   End
   Begin VB.CommandButton CloseButton 
      Caption         =   "Save and Close"
      Height          =   420
      Left            =   120
      TabIndex        =   16
      Top             =   6060
      Width           =   7455
   End
   Begin VB.Frame Frame8 
      Caption         =   "Options"
      Height          =   1755
      Left            =   4080
      TabIndex        =   14
      Top             =   2040
      Width           =   3495
      Begin VB.CheckBox OpenOnDownloadCheck 
         Caption         =   "Open after download"
         Height          =   300
         Left            =   180
         TabIndex        =   29
         Top             =   1320
         Width           =   1890
      End
      Begin VB.CheckBox SoundsCheck 
         Caption         =   "Sounds Enabled"
         Height          =   300
         Left            =   180
         TabIndex        =   28
         Top             =   960
         Width           =   1530
      End
      Begin VB.CheckBox ChanListenCheck 
         Caption         =   "Listen in channels for offers."
         Height          =   300
         Left            =   180
         TabIndex        =   17
         Top             =   600
         Width           =   2355
      End
      Begin VB.CheckBox HeresMyBotsCheck 
         Caption         =   "Use other IRCFT user's bots"
         Height          =   300
         Left            =   180
         TabIndex        =   15
         Top             =   240
         Width           =   2370
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bot Comms"
      Height          =   795
      Left            =   5520
      TabIndex        =   11
      Top             =   5040
      Width           =   2055
      Begin VB.OptionButton UseMSGOption 
         Caption         =   "Use MSG "
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton UseCTCPOption 
         Caption         =   "Use CTCP"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1100
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Download Path"
      Height          =   2475
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   3800
      Begin VB.DirListBox DirSelect 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1770
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   3600
      End
      Begin VB.DriveListBox DriveSelect 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   3600
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Channel Details"
      Height          =   1695
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   3495
      Begin VB.ListBox ChannelList 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         ItemData        =   "prefs.frx":0000
         Left            =   1800
         List            =   "prefs.frx":0002
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton RemoveChannelButton 
         Caption         =   "Remove Channel"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton AddChannelButton 
         Caption         =   "Add This Channel"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Channel 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Server Details"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   3800
      Begin VB.TextBox ServerName 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2300
      End
      Begin VB.TextBox ServerPort 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   1
         Top             =   480
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "IRC Server"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Port"
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "PrefsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AddChannelButton_Click()

Dim i As Integer
Dim found As Boolean

If Len(Channel.Text) > 0 Then
    If Left(Channel.Text, 1) <> "#" Then Channel.Text = "#" + Channel.Text
    found = False
    For i = 0 To ChannelList.ListCount - 1
        If LCase(ChannelList.List(i)) = LCase(Channel.Text) Then found = True
    Next i
    If found = False Then
        ChannelList.AddItem Channel.Text
    End If
    Channel.Text = ""
End If

End Sub

Private Sub Channel_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then Call AddChannelButton_Click

End Sub

Private Sub CloseButton_Click()

'check all the data
'ensure there's a backslash at the end of the download path
If DownloadPath = "" Then DownloadPath = "c:\"
If (Asc(Right(DownloadPath, 1))) <> 92 Then
    DownloadPath = DownloadPath + "\"
End If

'check that nick is < 9 chars
If Nickname = "" Then Nickname = "Nickname"
If Len(Nickname) > 9 Then
    Nickname = Mid(Nickname, 1, 9)
End If

'check that theres a port #
If ServerPort.Text = "" Then
    ServerPort.Text = "6667"
End If

'check that there's a sensible time between commands
If Val(TimeBetweenCommandsText.Text) < 200 Then
    TimeBetweenCommandsText.Text = "200"
End If
If Val(TimeBetweenCommandsText.Text) > 5999 Then
    TimeBetweenCommandsText.Text = "250"
End If

'check that there's a sensible time between nags
If Val(TimeBetweenNagsText.Text) < 3000 Then
    TimeBetweenNagsText.Text = "3000"
End If

'check there's a server
If ServerName.Text = "" Then
    ServerName.Text = "irc.frontiernet.net"
End If

'check dcc relay nick
If DCCRelayNickText = "" Then DCCRelayNickText = "Nickname"
If Len(DCCRelayNickText) > 9 Then
    DCCRelayNickText = Mid(DCCRelayNickText, 1, 9)
End If


On Error GoTo ErrorHandler

If DirectoryStatus(App.Path + "\ircft.ini") <> 999 Then
    Kill App.Path + "\ircft.ini"
End If
Open App.Path + "\ircft.ini" For Output As #1
For i = 0 To (ChannelList.ListCount - 1)
    Print #1, "channel=" + ChannelList.List(i)
Next i
Print #1, "msgoption=" + Str(UseMSGOption.Value)
Print #1, "ctcpoption=" + Str(UseCTCPOption.Value)
Print #1, "server=" + ServerName.Text
Print #1, "port=" + ServerPort.Text
Print #1, "path=" + DownloadPath.Text
Print #1, "sounds=" + LTrim(Str(SoundsCheck.Value))
Print #1, "nick=" + Nickname
Print #1, "botshare=" + LTrim(Str(HeresMyBotsCheck.Value))
Print #1, "version=" + VersionText.Text
Print #1, "left=" + LTrim(MainForm.Left)
Print #1, "top=" + LTrim(MainForm.Top)
Print #1, "height=" + LTrim(MainForm.Height)
Print #1, "width=" + LTrim(MainForm.Width)
Print #1, "chanlisten=" + LTrim(Str(ChanListenCheck.Value))
Print #1, "commandpause=" + LTrim(Str(TimeBetweenCommandsText.Text))
Print #1, "nagpause=" + LTrim(Str(TimeBetweenNagsText.Text))
Print #1, "openondownload=" + LTrim(Str(OpenOnDownloadCheck.Value))
Print #1, "dccrelay=" + LTrim(Str(DCCRelayCheck.Value))
Print #1, "dccrelaynick=" + LTrim(DCCRelayNickText)
Close #1

PrefsForm.Hide

SetPreferences

Exit Sub

ErrorHandler:
    
Call WriteINI
Resume

End Sub

Private Sub DCCRelayCheck_Click()

If DCCRelayCheck.Value = 1 Then
    DCCRelayNickText.Enabled = True
    DCCRelayNickOnlineButton.Enabled = True
    DCCRelayNickText.BackColor = &H80000005
Else
    DCCRelayNickText.Enabled = False
    DCCRelayNickOnlineButton.Enabled = False
    DCCRelayNickText.BackColor = &H80000013
End If

End Sub

Private Sub DCCRelayNickOnlineButton_Click()

SendToServerImmediately "whois " + DCCRelayNickText.Text

End Sub

Private Sub DirSelect_Change()

If Right(DownloadPath.Text, 1) <> "\" Then
    DownloadPath.Text = DirSelect.Path + "\"
Else
    DownloadPath.Text = DirSelect.Path
End If

End Sub

Private Sub DriveSelect_Change()

DirSelect.Path = DriveSelect

End Sub

Private Sub Form_Load()

PrefsForm.Top = MainForm.Top + 240
PrefsForm.Left = MainForm.Left + 240
DirSelect.Path = DownloadPath.Text

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Cancel = -1

End Sub

Private Sub RemoveChannelButton_Click()

Dim i As Integer

For i = 0 To ChannelList.ListCount - 1
    If ChannelList.Selected(i) = True Then
        ChannelList.RemoveItem (i)
        Exit Sub
    End If
Next i

End Sub

Private Sub WriteLogButton_Click()

Dim Filename As String
Dim i As Integer

Filename = AssignFilename("ircft_log_" + Format(Date, "yymmdd") + "_" + Format(Time, "hhmm") + ".txt")

Open Preferences.DownloadPath + Filename For Output As 3

Print #3, "IRC Filetool Log Dump - " + Format(Now, "dddd, mmm d yyyy") + " - " + Format(Now, "hh:mm:ss AMPM")
Print #3, ""

For i = 0 To MainForm.LogList.ListCount - 1
    Print #3, MainForm.LogList.List(i)
Next i
Close 3

AppendLog "Log saved as " + Filename

End Sub

Private Sub WriteOffersButton_Click()

Dim Filename As String
Dim i As Integer

Filename = AssignFilename("ircft_offers_" + Format(Date, "yymmdd") + "_" + Format(Time, "hhmm") + ".txt")

Open Preferences.DownloadPath + Filename For Output As 2

Print #2, "IRC Filetool Log Dump - " + Format(Now, "dddd, mmm d yyyy") + " - " + Format(Now, "hh:mm:ss AMPM")
Print #2, ""

For i = 0 To MainForm.OffersList.ListCount - 1
    Print #2, MainForm.OffersList.List(i)
Next i
Close 2

AppendLog "Offers saved as " + Filename

End Sub
