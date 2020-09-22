VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form HelpForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "IRC Filetool Help"
   ClientHeight    =   11790
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11790
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   5355
      Left            =   6000
      TabIndex        =   6
      Top             =   6120
      Width           =   5500
      Begin VB.Label Label26 
         Caption         =   $"help.frx":0000
         Height          =   615
         Left            =   240
         TabIndex        =   13
         Top             =   4500
         Width           =   5000
      End
      Begin VB.Label Label25 
         Caption         =   $"help.frx":00BC
         Height          =   615
         Left            =   240
         TabIndex        =   12
         Top             =   3840
         Width           =   5000
      End
      Begin VB.Label Label21 
         Caption         =   "Find. Lets you find a specific word in the offer list."
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1740
         Width           =   5000
      End
      Begin VB.Label Label14 
         Caption         =   $"help.frx":0162
         Height          =   1035
         Left            =   240
         TabIndex        =   10
         Top             =   660
         Width           =   4995
      End
      Begin VB.Label Label13 
         Caption         =   "Main Window"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   5000
      End
      Begin VB.Label Label16 
         Caption         =   $"help.frx":0282
         Height          =   855
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   5000
      End
      Begin VB.Label Label15 
         Caption         =   $"help.frx":0374
         Height          =   795
         Left            =   240
         TabIndex        =   7
         Top             =   2940
         Width           =   5000
      End
   End
   Begin VB.Frame Frame4 
      Height          =   5355
      Left            =   6000
      TabIndex        =   5
      Top             =   480
      Width           =   5500
      Begin VB.Label Label31 
         Caption         =   $"help.frx":0494
         Height          =   675
         Left            =   240
         TabIndex        =   37
         Top             =   4080
         Width           =   4875
      End
      Begin VB.Label Label30 
         Caption         =   "Save Offers. Save the main window's offer list to a text file in the download directory."
         Height          =   495
         Left            =   240
         TabIndex        =   36
         Top             =   3540
         Width           =   4995
      End
      Begin VB.Label Label29 
         Caption         =   "Save Log. Saves the main window's log to a text file in the download directory."
         Height          =   435
         Left            =   240
         TabIndex        =   35
         Top             =   3000
         Width           =   4875
      End
      Begin VB.Label Label19 
         Caption         =   $"help.frx":0530
         Height          =   855
         Left            =   240
         TabIndex        =   34
         Top             =   2100
         Width           =   5055
      End
      Begin VB.Label Label20 
         Caption         =   "Open After Download. IRCFT will automatically open the download once it has finished downloading."
         Height          =   495
         Left            =   240
         TabIndex        =   33
         Top             =   1560
         Width           =   4995
      End
      Begin VB.Label Label28 
         Caption         =   $"help.frx":0627
         Height          =   855
         Left            =   240
         TabIndex        =   32
         Top             =   660
         Width           =   4995
      End
      Begin VB.Label Label18 
         Caption         =   "Preferences"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   1875
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5355
      Left            =   240
      TabIndex        =   2
      Top             =   6120
      Width           =   5500
      Begin VB.Label Label27 
         Caption         =   $"help.frx":070C
         Height          =   855
         Left            =   240
         TabIndex        =   24
         Top             =   4140
         Width           =   4995
      End
      Begin VB.Label Label11 
         Caption         =   "Channel Details. Add and remove channels you wish to join here."
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   3780
         Width           =   4995
      End
      Begin VB.Label Label12 
         Caption         =   $"help.frx":07EB
         Height          =   675
         Left            =   240
         TabIndex        =   22
         Top             =   3060
         Width           =   4995
      End
      Begin VB.Label Label10 
         Caption         =   "Delay Between Nags. Hundredths of seconds between nagging bots for an offer. Minimum 30 seconds."
         Height          =   435
         Left            =   240
         TabIndex        =   21
         Top             =   2520
         Width           =   4995
      End
      Begin VB.Label Label9 
         Caption         =   "Download Path. Sets the path that files will be downloaded to."
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2640
         Width           =   4995
      End
      Begin VB.Label Label8 
         Caption         =   "IRC Server Port. Sets the port of the server to connect to (usually 6666 or 6667)"
         Height          =   435
         Left            =   240
         TabIndex        =   19
         Top             =   1980
         Width           =   4995
      End
      Begin VB.Label Label7 
         Caption         =   "IRC Server. Sets the server name to connect to."
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1620
         Width           =   4995
      End
      Begin VB.Label Label6 
         Caption         =   "Nickname. Sets the name that IRCFT attempts to log on the server with."
         Height          =   435
         Left            =   240
         TabIndex        =   17
         Top             =   1140
         Width           =   4995
      End
      Begin VB.Label Label5 
         Caption         =   "Version Reply. If you feel the need to pretend you client is something other than IRC Filetool, enter the name of it here."
         Height          =   435
         Left            =   240
         TabIndex        =   16
         Top             =   660
         Width           =   4995
      End
      Begin VB.Label Label3 
         Caption         =   "Preferences"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   4995
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5355
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5500
      Begin VB.CommandButton OkayButton 
         Caption         =   "Okay"
         Height          =   375
         Left            =   2160
         TabIndex        =   38
         Top             =   4740
         Width           =   1275
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Web Site"
         Height          =   375
         Left            =   3300
         TabIndex        =   26
         Top             =   3480
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   660
         TabIndex        =   25
         Text            =   "http://members.tripod.com/~IRC_Filetool/index.html"
         Top             =   4200
         Width           =   3855
      End
      Begin VB.Label Label17 
         Caption         =   "Bugs and Suggestions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   2340
         Width           =   3735
      End
      Begin VB.Label Label22 
         Caption         =   "Bugs and suggestions can be e-mailed to me by going to my webpage or by e-mailing me direct at ircft@hotmail.com"
         Height          =   495
         Left            =   240
         TabIndex        =   29
         Top             =   2640
         Width           =   4755
      End
      Begin VB.Label Label23 
         Caption         =   "Updates"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label24 
         Caption         =   "News and Updates are a button push away. "
         Height          =   495
         Left            =   240
         TabIndex        =   27
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   $"help.frx":08A4
         Height          =   1245
         Left            =   240
         TabIndex        =   4
         Top             =   660
         Width           =   5000
      End
      Begin VB.Label Label1 
         Caption         =   "IRC Filetool"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   5865
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   10345
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Overview"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Main Window"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Preferences Part 1"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Preferences Part 2"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   495
      Left            =   5160
      TabIndex        =   15
      Top             =   5520
      Width           =   1215
   End
End
Attribute VB_Name = "HelpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()

Dim iret As Long
iret = ShellExecute(Me.hwnd, vbNullString, "http://members.tripod.com/~IRC_Filetool/index.html", vbNullString, "c:\", SW_SHOWNORMAL)

End Sub

Private Sub Form_Load()
    
    HelpForm.Width = 6150
    HelpForm.Height = 6400
    
    Frame1.Left = 240
    Frame2.Left = 240
    Frame3.Left = 240
    Frame4.Left = 240
    Frame1.Top = 480
    Frame2.Top = 480
    Frame3.Top = 480
    Frame4.Top = 480
    
    Frame4.Visible = False
    Frame3.Visible = False
    Frame2.Visible = False
    Frame1.Visible = True
    
End Sub

Private Sub OkayButton_Click()

HelpForm.Hide

End Sub

Private Sub TabStrip1_Click()

If TabStrip1.SelectedItem.Index = 1 Then
    Frame4.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame1.Visible = True
End If

If TabStrip1.SelectedItem.Index = 2 Then
    Frame4.Visible = False
    Frame1.Visible = False
    Frame3.Visible = False
    Frame2.Visible = True
End If

If TabStrip1.SelectedItem.Index = 3 Then
    Frame4.Visible = False
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = True
End If

If TabStrip1.SelectedItem.Index = 4 Then
    Frame4.Visible = True
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
End If

End Sub
