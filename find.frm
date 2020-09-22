VERSION 5.00
Begin VB.Form FindForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find Text"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   3420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   1500
   End
   Begin VB.CommandButton FindButton 
      Caption         =   "Find"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1500
   End
   Begin VB.TextBox FindText 
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
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "FindForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()

FindForm.Hide

End Sub

Private Sub FindButton_Click()

Dim i As Integer
Dim found As Boolean

found = False
For i = (Val(FindButton.Tag) + 1) To MainForm.OffersList.ListCount - 1
    If InStr(LCase(MainForm.OffersList.List(i)), LCase(FindText.Text)) <> 0 Then
        MainForm.OffersList.TopIndex = i
        FindButton.Tag = Str(i)
        FindButton.Caption = "Find Next"
        found = True
        Exit For
    End If
Next i

If found = False Then FindButton.Tag = "-1"

End Sub

Private Sub FindText_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then Call FindButton_Click

End Sub

Private Sub Form_Load()

FindForm.Top = MainForm.Top + 240
FindForm.Left = MainForm.Left + 240

FindButton.Caption = "Find"
FindButton.Tag = "-1"

End Sub
