Attribute VB_Name = "Public"

Option Explicit
Public Type Applications
    strPath As String
    strAlias As String
End Type
Public hRasConn As Long, lConnectHandle As Long
Public StartApp(1 To 100) As Applications, iStartAppCount As Integer
Public tp As Integer

Public Const SUCCESS = 0&

Public iSeconds As Integer, iMinutes As Integer, iHours As Integer, iTotalSeconds As Integer, iTotalMinutes As Integer, iTotalHours As Integer
Public iMonthSeconds As Integer, iMonthMinutes As Integer, iMonthHours As Integer
Public iautodial As Integer

Public lngRASErrorNumber As Long
Public Const sReg As String = "Software\Harish Works\Net Buddy"
Public unlo As Integer

Public Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'shell icon constants & func.
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1
Public Const MAX_TOOLTIP As Integer = 64


Type NOTIFYICONDATA
     cbSize As Long
     hwnd As Long
     uId As Long
     uFlags As Long
     uCallbackMessage As Long
     hIcon As Long
     szTip As String * MAX_TOOLTIP
End Type
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public nfIconData As NOTIFYICONDATA

Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201

'translucent image for spalsh screen
Public Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "GDI32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Public Const RGN_OR = 2
Public Const WM_NCLBUTTONDOWN = &HA1

'set window on top
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const conHwndTopmost = -1
Public Const conSwpNoActivate = &H10
Public Const conSwpShowWindow = &H40

'********************************
'  Win32 Function Declarations
'********************************
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Function MakeRegion(picSkin As PictureBox) As Long
With frmSpla
    ' Make a windows "region" based on a given picture box'
    ' picture. This done by passing on the picture line-
    ' by-line and for each sequence of non-transparent
    ' pixels a region is created that is added to the
    ' complete region. I tried to optimize it so it's
    ' fairly fast, but some more optimizations can
    ' always be done - mainly storing the transparency
    ' data in advance, since what takes the most time is
    ' the GetPixel calls, not Create/CombineRgn
    
    Dim X As Long, Y As Long, StartLineX As Long
    Dim FullRegion As Long, LineRegion As Long
    Dim TransparentColor As Long
    Dim InFirstRegion As Boolean
    Dim InLine As Boolean  ' Flags whether we are in a non-tranparent pixel sequence
    Dim hDC As Long
    Dim PicWidth As Long
    Dim PicHeight As Long
    
    

    
    InFirstRegion = True: InLine = False
    X = Y = StartLineX = 0
    
    ' The transparent color is always the color of the
    ' top-left pixel in the picture. If you wish to
    ' bypass this constraint, you can set the tansparent
    ' color to be a fixed color (such as pink), or
    ' user-configurable
    TransparentColor = GetPixel(hDC, 0, 0)
    
    For Y = 0 To PicHeight - 1
        For X = 0 To PicWidth - 1
            
            If GetPixel(hDC, X, Y) = TransparentColor Or X = PicWidth Then
                ' We reached a transparent pixel
                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, Y, X, Y + 1)
                    
                    If InFirstRegion Then
                        FullRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FullRegion, FullRegion, LineRegion, RGN_OR
                        ' Always clean up your mess
                        DeleteObject LineRegion
                    End If
                End If
            Else
                ' We reached a non-transparent pixel
                If Not InLine Then
                    InLine = True
                    StartLineX = X
                End If
            End If
        Next
    Next
End With
    MakeRegion = FullRegion
End Function

Public Function IsConnected() As Boolean
On Error Resume Next
'check for connection
Dim TRasCon(255) As RASCONN
Dim lg As Long, retval As Long
Dim lpcon As Long
Dim Tstatus As RASCONNSTATUS
    On Error Resume Next
    TRasCon(0).dwSize = 412
    lg = 256 * TRasCon(0).dwSize
    retval = RasEnumConnections(TRasCon(0), lg, lpcon)
    Tstatus.dwSize = 160
    Call RasGetConnectStatus(TRasCon(0).hRasConn, Tstatus)
    If Tstatus.RASCONNSTATE = RASCS_Connected Then
        frmdial.status.Caption = "Connected."
        IsConnected = True
        lConnectHandle = TRasCon(0).hRasConn
        hRasConn = 0
    ElseIf Tstatus.RASCONNSTATE = RASCS_PortOpened Then
        frmdial.status.Caption = "Opening port for connection..."
        IsConnected = False
    ElseIf Tstatus.RASCONNSTATE = RASCS_LogonNetwork Then
        frmdial.status.Caption = "Logging on to network..."
        IsConnected = False
    ElseIf Tstatus.RASCONNSTATE = RASCS_Authenticate Then
        frmdial.status.Caption = "Verifying username and password..."
        IsConnected = False
    ElseIf Tstatus.RASCONNSTATE = RASCS_Authenticated Then
        Dim ispeed As Long
        ispeed = ReadLong(HKEY_DYN_DATA, "PerfStats\StatData", "Dial-Up Adapter\ConnectSpeed", 0)
        If ispeed > 0 Then
        frmdial.status.Caption = "Logon successful."
        hRasConn = 0
        Else
        frmdial.status.Caption = "Disconnected."
        IsConnected = False
        lConnectHandle = 0
        End If
    ElseIf Tstatus.RASCONNSTATE = RASCS_ConnectDevice Then
        IsConnected = False
        frmdial.status.Caption = "Dialing " & frmdial.Combo2.Text & " ..."
    End If
End Function
Public Function AddZero(intSeconds As Integer, intMinutes As Integer, intHours As Integer) As String
On Error Resume Next
'adding zeros in the time
Dim sSeconds As String, sMinutes As String, sHours As String
    sSeconds = CStr(intSeconds)
    sMinutes = CStr(intMinutes)
    sHours = CStr(intHours)
    If Len(sSeconds) = 1 Then
        sSeconds = "0" & sSeconds
    End If
    If Len(sMinutes) = 1 Then
        sMinutes = "0" & sMinutes
    End If
    If Len(sHours) = 1 Then
        sHours = "0" & sHours
    End If
    AddZero = sHours & ":" & sMinutes & ":" & sSeconds
End Function
Public Function ClearNulls(ByVal strSource As String) As String
On Error Resume Next
'clear nulls
Dim iPos As Integer
    iPos = InStr(strSource, Chr$(0))
    If iPos <> 0 Then
        ClearNulls = Left$(strSource, iPos - 1)
    End If
End Function
Public Function ShowOpenDialog(ByVal hOwner As Long, ByVal strTitle As String, ByVal strFilter As String) As String
Dim ofn As OPENFILENAME, lResult As Long, strFileName As String
    With ofn
        .lStructSize = Len(ofn)
        .hWndOwner = hOwner
        .hInstance = App.hInstance
        .lpstrFilter = strFilter
        .lpstrFile = Space$(255)
        .nMaxFile = 255
        .lpstrTitle = strTitle
        .flags = 4 + 4096
    End With
    lResult = GetOpenFileName(ofn)
    If lResult = 0 Then
        strFileName = ""
    Else
        strFileName = ofn.lpstrFile
        strFileName = ClearNulls(strFileName)
    End If
    ShowOpenDialog = strFileName
End Function
Public Function ValidateTime(ByVal strTime As String, ByVal bClockTime As Boolean) As Boolean
Dim i As Integer, iNr As Integer
    'Step 1
    If Len(strTime) > 8 Then
        ValidateTime = False
        Exit Function
    End If
    'Step 2
    For i = 1 To 8
        If Not IsNumeric(Mid(strTime, i, 1)) And Mid(strTime, i, 1) <> ":" Then
            ValidateTime = False
            Exit Function
        End If
    Next
    'Step 3
    For i = 1 To 8
        If Mid(strTime, i, 1) = ":" Then iNr = iNr + 1
    Next
    If iNr <> 2 Then
        ValidateTime = False
        Exit Function
    End If
    'Step 4
    If Mid(strTime, 3, 1) <> ":" Or Mid(strTime, 6, 1) <> ":" Then
        ValidateTime = False
        Exit Function
    End If
    'Step 5
    If Not bClockTime Then strTime = "23" & Mid(strTime, 3)
    If Not IsNumeric(Format(strTime, "hhmmss")) Then
        ValidateTime = False
        Exit Function
    End If
    ValidateTime = True
End Function
Function optimize(os As String)
'The optimization of RcvWindow and DefaultTTL along
'with other registry settings such as MaxMTU and MaxMSS
'can speed up TCP/IP modem networking connections (eg. Internet connections).
'RWIN (Receive WINdow) is the buffer your machine waits to fill with data
'before attending to whatever other TCP transactions are occurring on the
'other threads and sockets WinSock has open while a connection is in progress.
'The value of TTL (Time To Live) defines how long a packet can stay active before being discarded. The default value is '32'.

Select Case os
Case "Windows 95":
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "DefaultRcvWindow", "64240", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "DefaultTTL", "128", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "PMTUDiscovery", "0", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "PMTUBlackHoleDetect", "0", REG_SZ
    CreateNewKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings", HKEY_CURRENT_USER
    SetKeyValue "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MaxConnectionsPer1_0Server", "dword:00000010", REG_SZ
    CreateNewKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings", HKEY_CURRENT_USER
    SetKeyValue "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MaxConnectionsPerServer", "dword:00000008", REG_SZ
Case "Windows 98":
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "DefaultRcvWindow", "372300", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "DefaultTTL", "128", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "PMTUDiscovery", "0", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "PMTUBlackHoleDetect", "0", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "PMTUDiscovery", "0", REG_SZ
    '(string var, recommended setting is 3. The possible settings are 0 - No Windowscaling and Timestamp Options, 1 - Window scaling but no Timestamp options, 3 - Window scaling and Time stamp options.)
    CreateNewKey "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", "Tcp1323Opts", "3", REG_SZ
    '(string var, recommended setting is 1. Possible settings are 0 - No Sack options or 1 - Sack Option enabled)
    CreateNewKey "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", "SackOpts", "1", REG_SZ
    '(DWORD decimal var, taking integer values from 2 to N. Recommended setting is 3)
    CreateNewKey "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", "MaxDupAcks", "3", REG_SZ
    CreateNewKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings", HKEY_CURRENT_USER
    SetKeyValue "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MaxConnectionsPer1_0Server", "dword:00000010", REG_SZ
    CreateNewKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings", HKEY_CURRENT_USER
    SetKeyValue "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MaxConnectionsPerServer", "dword:00000008", REG_SZ
Case "Windows NT":
    CreateNewKey "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", HKEY_CURRENT_USER
    SetKeyValue "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", "TcpWindowSize", "64240", REG_SZ
    CreateNewKey "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", HKEY_CURRENT_USER
    SetKeyValue "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", "DefaultTTL", "128", REG_SZ
    CreateNewKey "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", HKEY_CURRENT_USER
    SetKeyValue "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", "EnablePMTUDiscovery", "0", REG_SZ
    CreateNewKey "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", HKEY_CURRENT_USER
    SetKeyValue "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", "EnablePMTUBHDetect", "0", REG_SZ
    CreateNewKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings", HKEY_CURRENT_USER
    SetKeyValue "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MaxConnectionsPer1_0Server", "dword:00000010", REG_SZ
    CreateNewKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings", HKEY_CURRENT_USER
    SetKeyValue "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MaxConnectionsPerServer", "dword:00000008", REG_SZ
Case "Windows 98 SE":
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "DefaultRcvWindow", "372300", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "DefaultTTL", "128", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "PMTUDiscovery", "0", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "PMTUBlackHoleDetect", "0", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "PMTUDiscovery", "0", REG_SZ
    '(string var, recommended setting is 3. The possible settings are 0 - No Windowscaling and Timestamp Options, 1 - Window scaling but no Timestamp options, 3 - Window scaling and Time stamp options.)
    CreateNewKey "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", "Tcp1323Opts", "3", REG_SZ
    '(string var, recommended setting is 1. Possible settings are 0 - No Sack options or 1 - Sack Option enabled)
    CreateNewKey "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", "SackOpts", "1", REG_SZ
    '(DWORD decimal var, taking integer values from 2 to N. Recommended setting is 3)
    CreateNewKey "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", "MaxDupAcks", "3", REG_SZ
    CreateNewKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings", HKEY_CURRENT_USER
    SetKeyValue "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MaxConnectionsPer1_0Server", "dword:00000010", REG_SZ
    CreateNewKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings", HKEY_CURRENT_USER
    SetKeyValue "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MaxConnectionsPerServer", "dword:00000008", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\ICSharing\Settings\General", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\ICSharing\Settings\General", "internetMTU", "1500", REG_SZ
End Select
End Function

