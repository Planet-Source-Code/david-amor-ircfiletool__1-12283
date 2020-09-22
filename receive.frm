VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ReceiveBoxForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receiving"
   ClientHeight    =   1845
   ClientLeft      =   3270
   ClientTop       =   4680
   ClientWidth     =   4605
   Icon            =   "receive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4605
   Begin ComctlLib.ProgressBar DataProgressBar 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   1590
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer KickstartTimer 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   960
      Top             =   1680
   End
   Begin VB.CommandButton TestButton 
      Caption         =   "Test"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Progress"
      Height          =   1000
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2655
      Begin VB.Label Label3 
         Caption         =   "Download CPS:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Received Data:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Estimated Time Left:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label InfoTimeRemaining 
         Caption         =   "Unknown"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Width           =   900
      End
      Begin VB.Label InfoCPS 
         Caption         =   "Unknown"
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   720
         Width           =   900
      End
      Begin VB.Label InfoReceivedBytes 
         Caption         =   "0"
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.CommandButton CancelAndDeleteButton 
      Caption         =   "Cancel and Delete"
      Height          =   400
      Left            =   2880
      TabIndex        =   2
      Top             =   1080
      Width           =   1620
   End
   Begin VB.Timer CPSTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   1680
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   400
      Left            =   2880
      TabIndex        =   10
      Top             =   600
      Width           =   1620
   End
   Begin VB.Label KickstartDot 
      Height          =   255
      Left            =   4550
      TabIndex        =   11
      Top             =   200
      Width           =   135
   End
   Begin VB.Label messages 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "ReceiveBoxForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelAndDeleteButton_Click()

Dim PathPlusFilename As String
Dim Index As Integer

Index = Val(Tag)

GenericCancel

'delete the file although check that *something* has been written and check that the file is still there to delete
PathPlusFilename = Preferences.DownloadPath + ReceiveInfo(Index).Filename
Filename = ReceiveInfo(Index).Filename
If DirectoryStatus(PathPlusFilename) <> 999 And DirectoryStatus(PathPlusFilename) <> 16 And DirectoryStatus(PathPlusFilename) <> 1 Then
    Kill Preferences.DownloadPath + ReceiveInfo(Index).Filename
    'tell the user that download was cancelled and file was deleted
    AppendLog "Cancelled and Deleted " + ReceiveInfo(Index).Filename
Else
    AppendLog "Cancelled. Error deleting " + ReceiveInfo(Index).Filename
End If


End Sub

Private Sub CancelButton_Click()

Dim Index As Integer

Index = Val(Tag)

GenericCancel

'tell the user that download was cancelled
AppendLog "Cancelled " + ReceiveInfo(Index).Filename

End Sub

Private Sub CPSTimer_Timer()

Dim Index, BytesRemaining, TotalSecondsRemaining, SecondsRemaining, MinutesRemaining, HoursRemaining As Integer

Index = Val(Tag)

If ReceiveInfo(Index).ReceivedBytes > 0 Then
    
    ReceiveInfo(Index).CPS = Fix(ReceiveInfo(Index).ReceivedBytes / (DateDiff("s", ReceiveInfo(Index).StartTime, Now) + 1))
    'then display it
    ReceiveBox(Index).InfoCPS.Caption = LTrim(Str(ReceiveInfo(Index).CPS))

    'work out time remaining
    BytesRemaining = ReceiveInfo(Index).Filelength - ReceiveInfo(Index).ReceivedBytes
    TotalSecondsRemaining = Fix(BytesRemaining / ReceiveInfo(Index).CPS)
    HoursRemaining = Fix(TotalSecondsRemaining / 3600)
    If HoursRemaining > 0 Then
        MinutesRemaining = TotalSecondsRemaining Mod 60
        InfoTimeRemaining.Caption = LTrim(Str(HoursRemaining)) & "h " & LTrim(Str(MinutesRemaining)) & "m "
    Else
        MinutesRemaining = Fix(TotalSecondsRemaining / 60)
        SecondsRemaining = TotalSecondsRemaining Mod 60
        InfoTimeRemaining.Caption = LTrim(Str(MinutesRemaining)) & "m " & LTrim(Str(SecondsRemaining)) & "s"
    End If
End If

'check the connection is still active
If ReceiveInfo(Index).ReceivedBytes <> ReceiveInfo(Index).Filelength And ReceiveInfo(Index).InUse = True Then
    If MainForm.ReceiveConnection(Index).State <> sckConnected And ReceiveInfo(Index).ReceivedBytes > 0 Then
        'stop the timer
        CPSTimer.Enabled = False
        'alert the user
        Debug.Print "Connection lost receiving " + ReceiveInfo(Index).Filename

        'reset data
        InfoTimeRemaining.Caption = "Unknown"
        InfoCPS.Caption = "Error!"
    End If
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Cancel = -1

End Sub

Private Sub KickStartTimer_Timer()

Dim Index, i As Integer
Dim Hexdata As String
Dim SendBackData() As Byte

Index = Val(Tag)
If DateDiff("s", ReceiveInfo(Index).LastTimeReceivedData, Now) > 2 And ReceiveInfo(Index).ReceivedBytes > 0 Then
    'if we haven't received data for a while, give it a kick up the arse
    'give a visual indication
    If KickstartDot.Caption = "." Then
        KickstartDot.Caption = " "
    Else
        KickstartDot.Caption = "."
    End If
    Hexdata = Hex(LOF(Index))
    Hexdata = String$(8 - Len(Hexdata), "0") & Hexdata
    ReDim SendBackData(3) As Byte
    For i = 1 To Len(Hex_Data) Step 2
        SendBackData((i - 1) / 2) = Val("&H" & Mid(Hexdata, i, 2))
    Next
    If MainForm.ReceiveConnection(Index).State = 7 Then MainForm.ReceiveConnection(Index).SendData SendBackData
    'Debug.Print "Kickstarted!"
End If

End Sub

Private Sub GenericCancel()

Dim Index As Integer

Index = Val(Tag)

While MainForm.ReceiveConnection(Index).State <> sckClosed
    MainForm.ReceiveConnection(Index).Close
    DoEvents
Wend

Unload MainForm.ReceiveConnection(Index)
CPSTimer.Enabled = False
KickstartTimer.Enabled = False

'hide the form
ReceiveBox(Index).Hide
'close the file
Close Index

'reset data
ReceiveInfo(Index).CPS = 0
ReceiveInfo(Index).ReceivedBytes = 0
InfoTimeRemaining.Caption = "Unknown"
InfoReceivedBytes.Caption = "0"
InfoCPS.Caption = "Unknown"
DataProgressBar.Value = 0

'flag that it's no longer in use
ReceiveInfo(Index).InUse = False

End Sub
