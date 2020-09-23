VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Banner 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Forced Banner Click"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   Icon            =   "banner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ok 
      BackColor       =   &H00C0C0C0&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4320
      Width           =   855
   End
   Begin VB.Timer Time1 
      Interval        =   500
      Left            =   6720
      Top             =   4440
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   4080
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   7995
      ExtentX         =   14102
      ExtentY         =   7197
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label LblReg 
      BackColor       =   &H00000000&
      Caption         =   $"banner.frx":08CA
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   6255
   End
End
Attribute VB_Name = "Banner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
ok.Enabled = False
On Error Resume Next
    web.Navigate "http://www.geocities.com/insomniac55/banner1.htm"
End Sub

Private Sub ok_Click()
On Error Resume Next
    Main.Support.Enabled = False
    Main.Enter.Enabled = True
Unload Banner
End Sub
Private Sub time1_Timer()
On Error Resume Next
    If LCase(web.LocationURL) = LCase("http://www.modchip.com/") Or LCase(web.LocationURL) = LCase("http://www.modchip.com/") Then
        ok.Enabled = True
    End If
End Sub
