VERSION 5.00
Begin VB.Form frmSpla 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrPause 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   2625
      Top             =   2835
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "INTERNET DIALLER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   3645
   End
End
Attribute VB_Name = "frmSpla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub Form_Load()
    Dim WindowRegion As Long

    Me.BorderStyle = vbBSNone

    SetWindowRgn Me.hwnd, WindowRegion, True
    Me.Top = Screen.Height / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2
    DoEvents
    'Me.Show
    tmrPause.Enabled = True
    End Sub
Private Sub tmrPause_Timer()
Load frmdial
frmdial.Show
DoEvents
Unload Me
End Sub


