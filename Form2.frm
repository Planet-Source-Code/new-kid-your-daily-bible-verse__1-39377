VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search"
   ClientHeight    =   840
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5730
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "Search by Keyword"
      Top             =   480
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   255
      Left            =   4320
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Search by Passage"
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set mIE = CreateObject("InternetExplorer.Application")
go = "http://www.biblegateway.com/cgi-bin/bible?passage=" & Me.Text1.Text '& "&search=&version=ALL&language=english&optional.x=12&optional.y=4"


    With mIE
        .Navigate go
        '.Width = 420
        '.Height = 360
        .Visible = True
    End With


'r = Shell("C:\windows\explorer.exe " + go, vbNormalFocus)
End Sub

Private Sub Command2_Click()
Set mIE = CreateObject("InternetExplorer.Application")

go = "http://www.biblegateway.com/cgi-bin/bible?passage=&search=" & Me.Text2.Text & "&version=ALL&language=english&optional.x=0&optional.y=0"

    With mIE
        .Navigate go
        '.Width = 420
        '.Height = 360
        .Visible = True
    End With
    
End Sub
