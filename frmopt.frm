VERSION 5.00
Begin VB.Form frmopt 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Dial Net Options..."
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmopt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option7 
      Caption         =   "Connection Details"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   120
      Width           =   3255
   End
   Begin VB.OptionButton Option6 
      Caption         =   "General"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   120
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   480
      ScaleHeight     =   5295
      ScaleWidth      =   6615
      TabIndex        =   3
      Top             =   480
      Width           =   6615
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   360
         Picture         =   "frmopt.frx":030A
         ScaleHeight     =   615
         ScaleWidth      =   735
         TabIndex        =   45
         Top             =   120
         Width           =   735
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "  Optimize  my  internet  connection "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   360
         TabIndex        =   31
         Top             =   3960
         Width           =   5895
         Begin VB.OptionButton win98se 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Windows 98 SE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3120
            TabIndex        =   37
            Top             =   840
            Width           =   1455
         End
         Begin VB.OptionButton winnt 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Windows NT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3120
            TabIndex        =   36
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Optimize"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4680
            TabIndex        =   34
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton win95 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Windows 95"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1800
            TabIndex        =   33
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton win98 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Windows 98"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1800
            TabIndex        =   32
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Operating System :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   360
            TabIndex        =   35
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Applications to start when connection is detected "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1575
         Left            =   360
         TabIndex        =   10
         Top             =   2280
         Width           =   5895
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            ItemData        =   "frmopt.frx":2AAC
            Left            =   120
            List            =   "frmopt.frx":2AAE
            TabIndex        =   13
            Top             =   360
            Width           =   3975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Add Application"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   4200
            TabIndex        =   12
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Remove"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   4200
            TabIndex        =   11
            Top             =   840
            Width           =   1332
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Default Connections "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   5895
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Pulse Dial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3120
            TabIndex        =   27
            Top             =   1080
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tone Dial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1800
            TabIndex        =   26
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton Command4 
            Caption         =   "C&hange"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4440
            TabIndex        =   7
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox Combo3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2400
            TabIndex        =   6
            Top             =   480
            Width           =   1575
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Dial Using :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   600
            TabIndex        =   28
            Top             =   1080
            Width           =   810
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Default User name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   360
            TabIndex        =   9
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Default Phone No:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2520
            TabIndex        =   8
            Top             =   240
            Width           =   1320
         End
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   480
      ScaleHeight     =   5295
      ScaleWidth      =   6615
      TabIndex        =   14
      Top             =   600
      Width           =   6615
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   240
         Picture         =   "frmopt.frx":2AB0
         ScaleHeight     =   615
         ScaleWidth      =   735
         TabIndex        =   46
         Top             =   120
         Width           =   735
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0C0C0&
         Caption         =   " When connected to the internet "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   240
         TabIndex        =   38
         Top             =   3840
         Width           =   5895
         Begin VB.CommandButton Command13 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4320
            Picture         =   "frmopt.frx":2EF2
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   480
            Width           =   375
         End
         Begin VB.CommandButton Command12 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4800
            TabIndex        =   42
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   600
            TabIndex        =   40
            Top             =   480
            Width           =   3255
         End
         Begin VB.CommandButton Command11 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3840
            Picture         =   "frmopt.frx":307C
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Leave it empty if you do  not want any music to be played."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   480
            TabIndex        =   43
            Top             =   960
            Width           =   4095
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Play : "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   480
            Width           =   435
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Correct The Counters "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   240
         TabIndex        =   20
         Top             =   2400
         Width           =   5895
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   960
            TabIndex        =   23
            Text            =   "0:0:0"
            Top             =   720
            Width           =   1092
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   2640
            TabIndex        =   22
            Text            =   "0:0:0"
            Top             =   720
            Width           =   1092
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Chan&ge"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4200
            TabIndex        =   21
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Month Time:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   960
            TabIndex        =   25
            Top             =   360
            Width           =   870
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Time:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2640
            TabIndex        =   24
            Top             =   360
            Width           =   810
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Add / Remove Username "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1335
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   5895
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            TabIndex        =   18
            Top             =   600
            Width           =   2655
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Add User"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            TabIndex        =   17
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Remove User"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            TabIndex        =   16
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Username:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   600
            Width           =   765
         End
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Apply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   6000
      Width           =   1215
   End
End
Attribute VB_Name = "frmopt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error Resume Next
'add application to the list
Dim sFile As String, sAlias As String
    sFile = ShowOpenDialog(Me.hwnd, "Add new application", "Applications (*.exe)" + Chr$(0) + "*.exe" + Chr$(0))
    If Len(sFile) > 0 Then
        sAlias = InputBox("Enter a description for the program you have just added. If you don't enter any text, the application's path will be used as identifier." & vbCr & vbCr & "Example: Internet Explorer", "Enter Title", "")
        If sAlias = "" Then sAlias = sFile
        List1.AddItem sAlias
        iStartAppCount = iStartAppCount + 1
        StartApp(iStartAppCount).strPath = sFile
        StartApp(iStartAppCount).strAlias = sAlias
    End If
    List1.SetFocus
End Sub

Private Sub Command10_Click()
'Optimizing modem settings
Dim os As String
os = ""
If win95 Then os = "Windows 95"
If win98 Then os = "Windows 98"
If winnt Then os = "Windows NT"
If win98se Then os = "Windows 98 SE"
If os = "" Then MsgBox "Please select your Operating System.", , "Dial Net": Exit Sub
s = MsgBox("Click yes to optimize your settings for " _
& os & ". Click no if you do not want to " _
& "proceed or if you do not have " & os & _
" installed on your PC.", vbYesNo, "Dial Net")
If s = 6 Then
Call optimize(os)
MsgBox "Restart your PC to experience the effects", , "Dial Net"
End If
End Sub
Private Sub Command11_Click()
'show open dialog to select sounds
Text4.Text = ShowOpenDialog(Me.hwnd, "Choose the sound file", "Wave sounds (*.wav)" & Chr$(0) & "*.wav" & Chr$(0) & "All files (*.*)" & Chr$(0) & "*.*" & Chr(0))
End Sub
Private Sub Command12_Click()
'save the sound file address
WriteString HKEY_CURRENT_USER, sReg, "Sound", Text4
End Sub
Private Sub Command13_Click()
'test the sound
On Error Resume Next
If Not Text4 = "" Then blare (Text4)
End Sub
Private Sub Command2_Click()
On Error Resume Next
'remove application from the applications list.
Dim i As Integer
    If MsgBox("Are you sure you want to remove " & List1.List(List1.ListIndex) & " from your applications list?", vbQuestion + vbOKCancel, "Confirm delete") = vbCancel Then
        List1.SetFocus
        Exit Sub
    End If
    If List1.ListIndex = List1.ListCount - 1 Then
        List1.RemoveItem List1.ListIndex
    Else
        For i = List1.ListIndex + 2 To iStartAppCount
            StartApp(i - 1).strPath = StartApp(i).strPath
            StartApp(i - 1).strAlias = StartApp(i).strAlias
        Next
        List1.RemoveItem List1.ListIndex
    End If
    iStartAppCount = iStartAppCount - 1
    Command2.Enabled = False
    List1.SetFocus
End Sub
Private Sub Command3_Click()
On Error Resume Next
'saving tone / pulse dialing
If Option1.Value = True Then
tp = 1
Else
tp = 2
End If
WriteInteger HKEY_CURRENT_USER, sReg, "T/P", tp
'Saving applications
    WriteInteger HKEY_CURRENT_USER, sReg & "\Applications", "Count", iStartAppCount
    If iStartAppCount > 0 Then
        For i = 1 To iStartAppCount
            WriteString HKEY_CURRENT_USER, sReg & "\Applications", "Start" & CStr(i), StartApp(i).strPath & Chr(1) & StartApp(i).strAlias
        Next
    End If
End Sub
Private Sub Command4_Click()
On Error Resume Next
'write default pwd, uid
WriteString HKEY_CURRENT_USER, sReg, "Defuid", Combo2.Text
WriteString HKEY_CURRENT_USER, sReg, "Defpwd", Combo3.Text
End Sub
Private Sub Command5_Click()
On Error Resume Next
'change the time settings
If Not ValidateTime(Text1.Text, False) Then
        MsgBox "The time you entered in one of the fields is invalid. You must enter the time in the HH:MM:SS pattern." & vbCrLf & vbCrLf & "Example of correct time: 12:45:00", vbCritical + vbOKOnly, "Invalid format"
        Text1.SetFocus
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
        Exit Sub
    End If
    If Not ValidateTime(Text3.Text, False) Then
        MsgBox "The time you entered in one of the fields is invalid. You must enter the time in the HH:MM:SS pattern." & vbCrLf & vbCrLf & "Example of correct time: 12:45:00", vbCritical + vbOKOnly, "Invalid format"
        Text3.SetFocus
        Text3.SelStart = 0
        Text3.SelLength = Len(Text3.Text)
        Exit Sub
    End If
    If Len(Text2.Text) < 1 Or Val(Text2.Text) > 100000 Then
        MsgBox "Please enter a valid number for the total calls value.", vbCritical + vbOKOnly, "Invalid value"
        Text2.SetFocus
        Exit Sub
    End If
    iTotalHours = GetTimePart(Text1.Text, TIME_HOURS)
    iTotalMinutes = GetTimePart(Text1.Text, TIME_MINUTES)
    iTotalSeconds = GetTimePart(Text1.Text, TIME_SECONDS)
    iMonthHours = GetTimePart(Text3.Text, TIME_HOURS)
    iMonthMinutes = GetTimePart(Text3.Text, TIME_MINUTES)
    iMonthSeconds = GetTimePart(Text3.Text, TIME_SECONDS)
    frmdial.lblTotal.Caption = AddZero(iTotalSeconds, iTotalMinutes, iTotalHours)
    frmdial.lblMonth.Caption = AddZero(iMonthSeconds, iMonthMinutes, iMonthHours)
End Sub
Private Sub Command6_Click()
On Error Resume Next
'add user
Call Savevalues
End Sub
Private Sub Command7_Click()
On Error Resume Next
'remove user
If Combo1.ListIndex = -1 Then MsgBox "Please Select the User to be removed."
For i = 1 To 100
If ValueExists(HKEY_CURRENT_USER, sReg, "U" & i) Then
s = ReadString(HKEY_CURRENT_USER, sReg, "U" & i, vbNull)
If StrComp(s, Combo1.Text, vbTextCompare) = 0 Then
DeleteValue HKEY_CURRENT_USER, sReg, "U" & i
DeleteValue HKEY_CURRENT_USER, sReg, "P" & i
If ValueExists(HKEY_CURRENT_USER, sReg, "R" & i) Then DeleteValue HKEY_CURRENT_USER, sReg, "R" & i
Combo1.RemoveItem (Combo1.ListIndex)
If Combo1.ListCount <> 0 Then Combo1.Text = Combo1.List(Combo1.ListCount - 1)
MsgBox "User Deleted.."
Call Loadvalues
Exit Sub
End If
Else
Call Loadvalues
Exit Sub
End If
Next i
End Sub
Private Sub Command8_Click()
On Error Resume Next
Unload Me
End Sub
Private Sub Command9_Click()
'ok button
'apply the changes and unload
On Error Resume Next
Command3_Click
Unload Me
End Sub
Private Sub Form_Load()
'read the option values for the options form
On Error Resume Next
Option7.Value = True
'sound filename
Text4 = strsound
'tone / pulse
If tp = 1 Then
Option1.Value = True
Else
option2.Value = True
End If
'default connections
For i = 0 To frmdial.Combo2.ListCount - 1
Combo3.AddItem (frmdial.Combo2.List(i))
Combo3.Text = frmdial.Combo2.List(i)
Next i
For i = 0 To frmdial.Combo3.ListCount - 1
Combo1.AddItem (frmdial.Combo3.List(i))
Combo1.Text = frmdial.Combo3.List(i)
Combo2.AddItem (frmdial.Combo3.List(i))
Combo2.Text = frmdial.Combo3.List(i)
Next i
s = ReadString(HKEY_CURRENT_USER, sReg, "Defuid", Combo2.Text)
Combo2.Text = s
s = ReadString(HKEY_CURRENT_USER, sReg, "Defpwd", Combo3.Text)
Combo3.Text = s
End Sub

Private Sub Option7_Click()
'the general tab
On Error Resume Next
Picture1.Visible = True
Picture2.Visible = False
Option6.Value = False
End Sub
Private Sub Option6_Click()
'the connections tab
On Error Resume Next
Picture2.Visible = True
Picture1.Visible = False
Option7.Value = False
End Sub
