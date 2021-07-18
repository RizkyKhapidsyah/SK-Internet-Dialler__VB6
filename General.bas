Attribute VB_Name = "General"

'sound constants
     Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
     Public Const SND_ASYNC = &H1 ' Return immediately after the sound starts.
     Public Const SND_NODEFAULT = &H2 ' If the sound file is not found, do NOT play default sound.
     Public Const SND_NOSTOP = &H10 ' Don't stop current sound to play another.
Public snd As Integer
Public strsound As String
'encryption constants
Public Const lEncryptKey = 68475297
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (lpString1 As Any, ByVal lpString2 As String) As Long
Public Sub Inc(iVariable As Integer, Optional iStep As Integer = 1)
    On Error Resume Next
    'incrementing
    iVariable = iVariable + iStep
End Sub
Public Function Encrypt(ByVal s As String, salt As Boolean) As String
    On Error Resume Next
    'encryption algorithm
    Dim n As Long, i As Long, ss As String
    Dim k1 As Long, k2 As Long, k3 As Long, k4 As Long, t As Long
    Static saltvalue As String * 4

    If salt Then
        For i = 1 To 4
            t = 100 * (1 + Asc(Mid(saltvalue, i, 1))) * Rnd() * (Timer + 1)
            Mid(saltvalue, i, 1) = Chr(t Mod 256)
        Next
        s = Mid(saltvalue, 1, 2) & s & Mid(saltvalue, 3, 2)
    End If

    n = Len(s)
    ss = Space(n)
    ReDim sn(n) As Long
    
    k1 = 11 + (lEncryptKey Mod 233): k2 = 7 + (lEncryptKey Mod 239)
    k3 = 5 + (lEncryptKey Mod 241): k4 = 3 + (lEncryptKey Mod 251)
    
    For i = 1 To n: sn(i) = Asc(Mid(s, i, 1)): Next i

    For i = 2 To n: sn(i) = sn(i) Xor sn(i - 1) Xor ((k1 * sn(i - 1)) Mod 256): Next
    For i = n - 1 To 1 Step -1: sn(i) = sn(i) Xor sn(i + 1) Xor (k2 * sn(i + 1)) Mod 256: Next
    For i = 3 To n: sn(i) = sn(i) Xor sn(i - 2) Xor (k3 * sn(i - 1)) Mod 256: Next
    For i = n - 2 To 1 Step -1: sn(i) = sn(i) Xor sn(i + 2) Xor (k4 * sn(i + 1)) Mod 256: Next
    
    For i = 1 To n: Mid(ss, i, 1) = Chr(sn(i)): Next i

    Encrypt = ss
    saltvalue = Mid(ss, Len(ss) / 2, 4)
End Function
Public Function Decrypt(ByVal s As String, salt As Boolean) As String
Dim n As Long, i As Long, ss As String
Dim k1 As Long, k2 As Long, k3 As Long, k4 As Long
On Error Resume Next
'Decryption algo
n = Len(s)
ss = Space(n)
ReDim sn(n) As Long

k1 = 11 + (lEncryptKey Mod 233): k2 = 7 + (lEncryptKey Mod 239)
k3 = 5 + (lEncryptKey Mod 241): k4 = 3 + (lEncryptKey Mod 251)

For i = 1 To n: sn(i) = Asc(Mid(s, i, 1)): Next

For i = 1 To n - 2: sn(i) = sn(i) Xor sn(i + 2) Xor (k4 * sn(i + 1)) Mod 256: Next
For i = n To 3 Step -1: sn(i) = sn(i) Xor sn(i - 2) Xor (k3 * sn(i - 1)) Mod 256: Next
For i = 1 To n - 1: sn(i) = sn(i) Xor sn(i + 1) Xor (k2 * sn(i + 1)) Mod 256: Next
For i = n To 2 Step -1: sn(i) = sn(i) Xor sn(i - 1) Xor (k1 * sn(i - 1)) Mod 256: Next

For i = 1 To n: Mid(ss, i, 1) = Chr(sn(i)): Next i

If salt Then Decrypt = Mid(ss, 3, Len(ss) - 4) Else Decrypt = ss
End Function
Function Savevalues()
On Error Resume Next
'saving the values for uid pwd
Dim s As String, i As Integer
WriteInteger HKEY_CURRENT_USER, sReg, "Autodial", frmdial.Check2.Value
For i = 1 To 100
If ValueExists(HKEY_CURRENT_USER, sReg, "U" & i) Then
s = ReadString(HKEY_CURRENT_USER, sReg, "U" & i, vbNull)
If StrComp(s, frmdial.Combo3.Text, vbTextCompare) = 0 Then
If frmdial.Check1.Value = 1 Then
WriteString HKEY_CURRENT_USER, sReg, "R" & i, frmdial.Check1.Value
s = Encrypt(frmdial.Text1.Text, False)
s = Encrypt(s, False)
WriteString HKEY_CURRENT_USER, sReg, "P" & i, s
Else
WriteString HKEY_CURRENT_USER, sReg, "R" & i, frmdial.Check1.Value
End If
Exit Function
End If
Else
WriteString HKEY_CURRENT_USER, sReg, "U" & i, frmdial.Combo3.Text
s = Encrypt(frmdial.Text1.Text, False)
s = Encrypt(s, False)
WriteString HKEY_CURRENT_USER, sReg, "P" & i, s
If frmdial.Check1.Value = 1 Then
WriteString HKEY_CURRENT_USER, sReg, "R" & i, frmdial.Check1.Value
End If
frmdial.Combo3.Clear
Call Loadvalues
Exit Function
End If
Next i
End Function
Function Loadvalues()
On Error Resume Next
'load uid, pwd
Dim s As String, i As Integer
For i = 1 To 100
If ValueExists(HKEY_CURRENT_USER, sReg, "U" & i) Then
s = ReadString(HKEY_CURRENT_USER, sReg, "U" & i, vbNull)
frmdial.Combo3.AddItem (s)
frmdial.Combo3.Text = s
s = ReadString(HKEY_CURRENT_USER, sReg, "R" & i, vbNull)
If s = 1 Then
frmdial.Check1.Value = Val(ReadString(HKEY_CURRENT_USER, sReg, "R" & i, 0))
s = ReadString(HKEY_CURRENT_USER, sReg, "P" & i, 0)
s = Decrypt(s, False)
frmdial.Text1.Text = Decrypt(s, False)
Else
frmdial.Check1.Value = 0
frmdial.Text1.Text = ""
End If
Else
Exit Function
End If
Next i
End Function
Function Loadphnos()
On Error Resume Next
'loading phone nos
Dim i As Integer, s As String
For i = 1 To 100
If ValueExists(HKEY_CURRENT_USER, sReg, "Ph" & i) Then
s = ReadString(HKEY_CURRENT_USER, sReg, "Ph" & i, vbNull)
frmdial.Combo2.AddItem (s)
frmdial.Combo2.Text = s
Else
Exit Function
End If
Next i
End Function
Function savephnos()
On Error Resume Next
'saving phone nos
Dim s As String, i As Integer
For i = 1 To 100
If ValueExists(HKEY_CURRENT_USER, sReg, "Ph" & i) Then
s = ReadString(HKEY_CURRENT_USER, sReg, "Ph" & i, vbNull)
If StrComp(s, frmdial.Combo2.Text, vbTextCompare) = 0 Then Exit Function
Else
WriteString HKEY_CURRENT_USER, sReg, "Ph" & i, frmdial.Combo2.Text
Exit Function
End If
Next i
End Function
Function Refreshvalues()
On Error Resume Next
'refresh values for uid & pwd
Dim s As String, i As Integer
For i = 1 To 100
If ValueExists(HKEY_CURRENT_USER, sReg, "U" & i) Then
s = ReadString(HKEY_CURRENT_USER, sReg, "U" & i, vbNull)
If StrComp(s, frmdial.Combo3.Text, vbTextCompare) = 0 Then
s = ReadString(HKEY_CURRENT_USER, sReg, "R" & i, vbNull)
If s = 1 Then
frmdial.Check1.Value = 1
s = ReadString(HKEY_CURRENT_USER, sReg, "P" & i, 0)
s = Decrypt(s, False)
frmdial.Text1.Text = Decrypt(s, False)
Else
frmdial.Check1.Value = 0
frmdial.Text1.Text = ""
Exit Function
End If
End If
Else
Exit Function
End If
Next i
End Function
Function Loaddata()
On Error Resume Next
'load other data
Dim sTime As String
'tone / pulse
tp = ReadInteger(HKEY_CURRENT_USER, sReg, "T/P", 1)
'autostart applications
iStartAppCount = ReadInteger(HKEY_CURRENT_USER, sReg & "\Applications", "Count", 0)
    If iStartAppCount > 0 Then
        For i = 1 To iStartAppCount
            sTime = ReadString(HKEY_CURRENT_USER, sReg & "\Applications", "Start" & CStr(i), "")
            If Len(sTime) > 0 And InStr(1, sTime, Chr(1)) > 0 Then
                StartApp(i).strPath = Mid(sTime, 1, InStr(1, sTime, Chr(1)) - 1)
                StartApp(i).strAlias = Mid(sTime, InStr(1, sTime, Chr(1)) + 1)
            Else: iStartAppCount = iStartAppCount - 1
            End If
            If Len(Dir(StartApp(i).strPath)) = 0 Or Len(StartApp(i).strPath) = 0 Then iStartAppCount = iStartAppCount - 1
        Next
    End If
'sound filename(frmopt)
strsound = ReadString(HKEY_CURRENT_USER, sReg, "Sound", "")
'dialer settings
iautodial = ReadInteger(HKEY_CURRENT_USER, sReg, "Autodial", 0)
iTotalSeconds = ReadInteger(HKEY_CURRENT_USER, sReg, "Seconds", 0)
iTotalMinutes = ReadInteger(HKEY_CURRENT_USER, sReg, "Minutes", 0)
iTotalHours = ReadInteger(HKEY_CURRENT_USER, sReg, "Hours", 0)
'Reading month time
sTime = ReadString(HKEY_CURRENT_USER, sReg, "MonthTime", "0:0:0")
iMonthHours = Val(Mid(sTime, 1, InStr(1, sTime, ":") - 1))
iMonthMinutes = Val(Mid(sTime, InStr(1, sTime, ":") + 1, InStr(InStr(1, sTime, ":"), sTime, ":")))
sTime = Mid(sTime, Len(sTime) - 1)
If Left$(sTime, 1) = ":" Then sTime = Right$(sTime, 1)
iMonthSeconds = Val(sTime)
frmdial.Check2.Value = iautodial
End Function
Function Savedata()
On Error Resume Next
'dialer settings
WriteInteger HKEY_CURRENT_USER, sReg, "Seconds", iTotalSeconds
WriteInteger HKEY_CURRENT_USER, sReg, "Minutes", iTotalMinutes
WriteInteger HKEY_CURRENT_USER, sReg, "Hours", iTotalHours
WriteString HKEY_CURRENT_USER, sReg, "MonthTime", CStr(iMonthHours) & ":" & CStr(iMonthMinutes) & ":" & CStr(iMonthSeconds)
End Function
Function Createreg()
On Error Resume Next
'registry codes
If Not KeyExists(HKEY_CURRENT_USER, sReg) Then
CreateKey HKEY_CURRENT_USER, sReg
End If
End Function
Sub Startrek(frm As Form)
On Error Resume Next
'just another animation
gotoval = frm.Height / 2
For Gointo = 1 To gotoval
DoEvents
frm.Height = frm.Height - 100
frm.Top = (Screen.Height - frm.Height) \ 2
If frm.Height <= 500 Then Exit For
Next Gointo
horiz:
frm.Height = 30
gotoval = frm.Width / 2
For Gointo = 1 To gotoval
DoEvents
frm.Width = frm.Width - 100
frm.Left = (Screen.Width - frm.Width) \ 2
If frm.Width <= 2000 Then Exit For
Next Gointo
End Sub
Function Writelog(sd As String, st As String, et As String, co As String)
On Error Resume Next
'write log
s = Dir("c:\ilog.log")
If Not s = "" Then
Open "c:\ilog.log" For Input As #1
Do While Not EOF(1)
Input #1, s
Loop
Close #1
End If
If s = "  Date  || Started At || Finished At|| Connection " Then
Open "c:\ilog.log" For Append As #1
Print #1, sd, st, et, co
Close #1
Else
Open "c:\ilog.log" For Output As #1
Print #1, "  Date  || Started At || Finished At|| Connection "
Print #1, sd, st, et, co
Close #1
End If
End Function
Function StartApplications()
On Error Resume Next
'start applications
Dim i As Integer
    On Error Resume Next
    If iStartAppCount > 0 Then
        For i = 1 To iStartAppCount
            Call Shell(StartApp(i).strPath, vbNormalFocus)
        Next
    End If
End Function
Function Defaultdata()
On Error Resume Next
'default data
If ValueExists(HKEY_CURRENT_USER, sReg, "Defuid") Then
s = ReadString(HKEY_CURRENT_USER, sReg, "Defuid", frmdial.Combo2.Text)
frmdial.Combo3.Text = s
End If
If ValueExists(HKEY_CURRENT_USER, sReg, "Defpwd") Then
s = ReadString(HKEY_CURRENT_USER, sReg, "Defpwd", frmdial.Combo3.Text)
frmdial.Combo2.Text = s
End If
End Function
Public Function GetTimePart(ByVal strTime As String, ByVal iTimePart As Integer) As Integer
Dim iTime As Integer, s As String, i As Integer
    Select Case iTimePart
        Case TIME_HOURS
            iTime = Val(Mid(strTime, 1, InStr(1, strTime, ":") - 1))
        Case TIME_MINUTES
            iTime = Val(Mid(strTime, InStr(1, strTime, ":") + 1, InStr(InStr(1, strTime, ":"), strTime, ":")))
        Case TIME_SECONDS
            i = InStr(1, strTime, ":") + 1
            i = InStr(i, strTime, ":")
            s = Mid(strTime, i)
            If Left$(s, 1) = ":" Then s = Right$(s, Len(s) - 1)
            iTime = Val(s)
    End Select
    GetTimePart = iTime
End Function
Public Sub blare(strName As String)
sndPlaySound strName, SND_ASYNC Or SND_NODEFAULT
snd = 1
End Sub
