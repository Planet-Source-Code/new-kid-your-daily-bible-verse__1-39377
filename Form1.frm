VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " "
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Close"
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   3480
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   6480
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   975
      ScaleWidth      =   1095
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Search"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   3480
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      Picture         =   "Form1.frx":0852
      ScaleHeight     =   1455
      ScaleWidth      =   7815
      TabIndex        =   4
      Top             =   0
      Width           =   7815
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplied by Gosplecom.net"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   " Verse of the Day"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Today's Verse"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   3480
      Width           =   3975
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4260
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Form1.frx":6708
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   135
      Left            =   3960
      TabIndex        =   1
      Top             =   5400
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.RichTextBox1.Text = ""
frmBrowser.rtbBrowser.Text = ""
frmBrowser.JamieHTMLParser1.ParseHTML (frmBrowser.Inet1.OpenURL(Trim("http://www.gospelcom.net/")))
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Form2.Show
End Sub

Private Sub Form_Load()
If Command = "start" Then
Form1.Height = 3825
End If
Top = Screen.Height / 2.7
Left = Screen.Width / 3.7
Me.RichTextBox1.Text = ""
frmBrowser.rtbBrowser.Text = ""
frmBrowser.JamieHTMLParser1.ParseHTML (frmBrowser.Inet1.OpenURL(Trim("http://www.gospelcom.net/")))
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label3_Click()
Set mIE = CreateObject("InternetExplorer.Application")

    With mIE
        .Navigate "http://www.gospelcom.net/"
        '.Width = 420
        '.Height = 360
        .Visible = True
    End With
End Sub
