VERSION 5.00
Begin VB.Form LegalForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "License Agreement"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "0"
   Begin VB.CommandButton Command2 
      Caption         =   "I Decline"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "I Agree"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"legal.frx":0000
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "LegalForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

WriteINI
ReadINI

LegalForm.Hide
MainForm.Visible = True

End Sub

Private Sub Command2_Click()

End

End Sub

Private Sub Form_Load()

LegalForm.Top = MainForm.Top + 240
LegalForm.Left = MainForm.Left + 240

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Cancel = -1

End Sub


