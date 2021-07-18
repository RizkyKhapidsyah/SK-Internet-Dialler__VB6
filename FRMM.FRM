VERSION 5.00
Begin VB.Form frmdial 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Dial Net "
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   Icon            =   "FRMM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton opt 
      Caption         =   "Options"
      Height          =   375
      Left            =   4200
      TabIndex        =   20
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   1440
   End
   Begin VB.Timer tms 
      Interval        =   1
      Left            =   0
      Top             =   480
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   960
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Connection Statistics "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   240
      TabIndex        =   13
      Top             =   5400
      Width           =   5175
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Session Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   600
         TabIndex        =   19
         Top             =   300
         Width           =   1185
      End
      Begin VB.Label lblSession 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   840
         TabIndex        =   18
         Top             =   600
         Width           =   750
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3960
         TabIndex        =   17
         Top             =   600
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3885
         TabIndex        =   16
         Top             =   300
         Width           =   900
      End
      Begin VB.Label lblMonth 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2400
         TabIndex        =   15
         Top             =   600
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Month Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2280
         TabIndex        =   14
         Top             =   300
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Connection Status... "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   240
      TabIndex        =   11
      Top             =   3840
      Width           =   5175
      Begin VB.Label status 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.CommandButton exi 
      Caption         =   "Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Connect 
      Caption         =   "Connect "
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Auto Dial"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Remember Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2280
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   600
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   4560
      Picture         =   "FRMM.frx":030A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   690
      Left            =   480
      Picture         =   "FRMM.frx":097C
      Stretch         =   -1  'True
      Top             =   960
      Width           =   945
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   4
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   3
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   2
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   720
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   1
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   840
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   10
      Top             =   3300
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   960
      TabIndex        =   9
      Top             =   2280
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   960
      TabIndex        =   8
      Top             =   1860
      Width           =   945
   End
   Begin VB.Menu s 
      Caption         =   "down"
      Visible         =   0   'False
      Begin VB.Menu dial 
         Caption         =   "Dialer"
      End
      Begin VB.Menu dis 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu optio 
         Caption         =   "Options"
      End
      Begin VB.Menu conn 
         Caption         =   "Connection Details"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu ab 
         Caption         =   "About"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmdial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim starttime As String, alcon As Integer
Dim i As Integer, k As Integer
Private Sub ab_Click()
Load frmAbout
frmAbout.Show vbModal, Me
End Sub
Private Sub Combo2_Change()
On Error Resume Next
'phone number combo
If Len(Combo2.Text) = 0 Then
Status.Caption = "Please Enter a Phone Number to dial..."
Else
Status.Caption = "Press Connect to begin dialing..."
End If
End Sub
Private Sub Combo3_Change()
On Error Resume Next
'username combo
If Len(Combo3.Text) = 0 Then
Status.Caption = "Please Enter Your User Name.."
End If
End Sub
Private Sub Combo3_LostFocus()
On Error Resume Next
'refresh the values for the password if
'the remember password feature is enabled
Call Refreshvalues
End Sub
Private Sub conn_Click()
Load frmset
frmset.Show
End Sub
Private Sub Connect_Click()
On Error Resume Next
'connect to the given telephone number
If Len(Combo2.Text) = 0 Then
MsgBox "Please enter a username.."
Exit Sub
ElseIf Len(Text1.Text) = 0 Then
MsgBox "Please enter your password.."
Exit Sub
ElseIf Len(Combo3.Text) = 0 Then
MsgBox "Please enter a number to connect to..."
Exit Sub
End If
Timer2.Enabled = True
Connect.Enabled = False
opt.Enabled = False
exi.Caption = "Cancel"
Combo2.Enabled = False
Combo3.Enabled = False
Text1.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Call Savevalues
Call savephnos
dialnum (Combo1.List(Combo1.ListIndex))
Status.Caption = "Attempting to establish a connection..."
End Sub
Private Sub dial_Click()
s.Visible = False
Me.Height = 6735
Me.Width = 5715
Me.Visible = True
SetWindowPos Me.hwnd, conSwpShowWindow, 100, 100, 6735, 5715, conHwndTopmost
End Sub
Private Sub dis_Click()
Call RasDisconnect
End Sub
Private Sub exi_Click()
On Error Resume Next
'end the session.
If exi.Caption = "Cancel" Then
Shape1(i).Visible = False
WriteInteger HKEY_CURRENT_USER, sReg, "Month", Month(Date)
Status.Caption = " Please Click Connect to begin Dialing.."
Connect.Enabled = True
opt.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Text1.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
exi.Caption = "Close"
RasDisconnect
Else
Call Savedata
Call Startrek(Me)
unlo = 0
Unload Me
End If
End Sub
Private Sub exit_Click()
unlo = 0
Unload Me
End Sub
Private Sub Form_Load()
On Error Resume Next
Dim conn As Integer
'load the default values
Call Createreg
Call RasLoadEntries(frmdial.Combo1)
Call Loadvalues
Call Loadphnos
Call Loaddata
Call Defaultdata
snd = 0
lblMonth.Caption = AddZero(iMonthSeconds, iMonthMinutes, iMonthHours)
lblTotal.Caption = AddZero(iTotalSeconds, iTotalMinutes, iTotalHours)
    'Add icon to system tray
    With nfIconData
        .cbSize = Len(nfIconData)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = 512
        .hIcon = Me.Icon
        .szTip = "Dial Net" + Chr(0)
    End With
Call Shell_NotifyIcon(NIM_ADD, nfIconData)

If iautodial = 1 Then Connect_Click
alcon = 0
Call Refreshvalues
i = k = 0
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim tmpLong As Single
    tmpLong = X / Screen.TwipsPerPixelX

    Select Case tmpLong 'For system tray icon
    Case WM_LBUTTONDOWN
        s.Visible = False
    Case WM_RBUTTONUP
        s.Visible = False
        If Me.Visible = False Then
        s.Visible = True
        PopupMenu s, 2
        End If
    End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Cancel = unlo
If Cancel = 0 Then
If IsConnected Then
Writelog Date, starttime, Time, frmset.l1.Caption
End If
Call Savedata
Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
End
End If
Me.Visible = False
End Sub
Private Sub Opt_Click()
On Error Resume Next
'show the options
Load frmopt
frmopt.Show vbModal, Me
End Sub

Private Sub optio_Click()
Call Opt_Click
End Sub

Private Sub Text1_Change()
'if the password is empty, then ask the user to
'give the password.
If Len(Text1.Text) = 0 Then
Status.Caption = "Please Enter Your Pass Word.."
ElseIf Len(Combo2.Text) = 0 Then
Status.Caption = "Please Enter a Phone Number to dial.."
End If
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
'check for the connection.
    On Error Resume Next
    If IsConnected Then
    If Len(strsound) > 0 And snd = 0 Then blare (strsound)
        Timer2.Enabled = False
        tmp.Enabled = True
        tmu.Enabled = True
        Me.Hide
        If alcon = 0 Then
        tms.Enabled = True
        alcon = 1
        End If
        If iSeconds > 58 Then
            iSeconds = iSeconds - 59
            iMinutes = iMinutes + 1
            If iMinutes = 60 Then
                iHours = iHours + 1
                iMinutes = 0
            End If
        Else
        iSeconds = iSeconds + 1
        End If
        If iTotalSeconds > 58 Then
            iTotalSeconds = iTotalSeconds - 59
            iTotalMinutes = iTotalMinutes + 1
            If iTotalMinutes = 60 Then
                iTotalHours = iTotalHours + 1
                iTotalMinutes = 0
            End If
        Else
        iTotalSeconds = iTotalSeconds + 1
        End If
        If iMonthSeconds > 58 Then
            iMonthSeconds = iMonthSeconds - 59
            iMonthMinutes = iMonthMinutes + 1
            If iMonthMinutes = 60 Then
                iMonthHours = iMonthHours + 1
                iMonthMinutes = 0
            End If
        Else
        iMonthSeconds = iMonthSeconds + 1
        End If
        frmset.l3.Caption = "Connected"
    Else
        snd = 0
        tmu.Enabled = False
        frmset.l3.Caption = " Not Connected"
        di1.Caption = "&Connect"
        If alcon = 1 Then
        Status.Caption = "Disconnected"
        Connect.Enabled = True
        opt.Enabled = True
        Combo2.Enabled = True
        Combo3.Enabled = True
        Text1.Enabled = True
        Check1.Enabled = True
        Check2.Enabled = True
        exi.Caption = "Close"
        lblMonth.Caption = AddZero(iMonthSeconds, iMonthMinutes, iMonthHours)
        lblTotal.Caption = AddZero(iTotalSeconds, iTotalMinutes, iTotalHours)
        Writelog Date, starttime, Time, frmset.l1.Caption
        alcon = 0
        End If
    End If
    lblSession.Caption = AddZero(iSeconds, iMinutes, iHours)
    frmset.l4.Caption = lblSession.Caption
    Writelog Date, starttime, Time, frmset.l1.Caption
    With nfIconData
        .szTip = "Dial Net - " & Status.Caption & lblSession.Caption & Chr$(0)
        .uFlags = NIF_TIP
    End With
    Call Shell_NotifyIcon(NIM_MODIFY, nfIconData)
End Sub
Private Sub Timer2_Timer()
On Error Resume Next
'dialing animation.
If i > k Then
If k > 0 Then k = k - 1
i = i - 1
Shape1(i).Visible = True
Shape1(i).FillColor = vbBlue
Else
k = 3
i = i + 1
Shape1(i).FillColor = vbRed
Shape1(i).Visible = True
End If
For r = 0 To 4
If i <> r Then Shape1(r).Visible = False
Next r
End Sub
Private Sub tms_Timer()
On Error Resume Next
'on detection of connection,
'start the applications.
starttime = Time
Call StartApplications
frmset.l1.Caption = RasGetConnectionSpeed
frmset.l2.Caption = RasGetConnectedEntry
tms.Enabled = False
End Sub
