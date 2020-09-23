VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forced Banner Click"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Support 
      Caption         =   "Support"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Enter 
      Caption         =   "Enter"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "To use this, click the support button and follow the instructions there."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "This little code will get you lots of money cash for your visual basic programs."
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   2040
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Forced Banner Click"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Enter_Click()
Form1.Show
Main.Hide
End Sub

Private Sub Support_Click()
On Error Resume Next
Banner.Show
End Sub

Private Sub Form_Load()
On Error Resume Next
    Enter.Enabled = False
    Enter.Visible = True
    Load Banner

End Sub


