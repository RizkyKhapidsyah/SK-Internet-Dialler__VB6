Attribute VB_Name = "Ras"

Option Explicit

'*****************************
'  RAS Functions - Constants
'*****************************

Public Const RASCS_Connected = &H2000
Public Const RASCS_Disconnected = &H2001

Public Const RAS_NOTIFY_HWND = &HFFFFFFFF

Private Const ERROR_INVALID_HANDLE = 6

Public Const UNLEN = 256
Public Const DNLEN = 15
Public Const PWLEN = 256

Public Const RAS_MaxPhoneNumber = 128
Public Const RAS_MaxEntryName = 256
Public Const RAS_MaxCallbackNumber = RAS_MaxPhoneNumber
Public Const RAS_MaxDeviceType = 16
Public Const RAS_MaxDeviceName = 128
Public Enum RASCONNSTATE
    RASCS_OpenPort = 0
    RASCS_PortOpened = 1
    RASCS_ConnectDevice = 2
    RASCS_DeviceConnected = 3
    RASCS_AllDevicesConnected = 4
    RASCS_Authenticate = 5
    RASCS_AuthNotify = 6
    RASCS_AuthRetry = 7
    RASCS_AuthCallback = 8
    RASCS_AuthChangePassword = 9
    RASCS_AuthProject = 10
    RASCS_AuthLinkSpeed = 11
    RASCS_AuthAck = 12
    RASCS_ReAuthenticate = 13
    RASCS_Authenticated = 14
    RASCS_PrepareForCallback = 15
    RASCS_WaitForModemReset = 16
    RASCS_WaitForCallback = 17
    RASCS_Projected = 18
    RASCS_StartAuthentication = 19
    RASCS_CallbackComplete = 20
    RASCS_LogonNetwork = 21
    RASCS_SubEntryConnected = 22
    RASCS_SubEntryDisconnected = 23
    RASCS_Interactive = &H1000
    RASCS_RetryAuthentication = &H1001
    RASCS_CallbackSetByCaller = &H1002
    RASCS_PasswordExpired = &H1003
    RASCS_InvokeEapUI = &H1004
End Enum

'******************************
'  RAS Functions - Structures
'******************************

Private Type RASENTRYNAME
    dwSize As Long
    szEntryName(RAS_MaxEntryName) As Byte
End Type

Public Type RASCONN
    dwSize As Long
    hRasConn As Long
    szEntryName(RAS_MaxEntryName) As Byte
    szDeviceType(RAS_MaxDeviceType) As Byte
    szDeviceName(RAS_MaxDeviceName) As Byte
End Type

Public Type RASCONNSTATUS
    dwSize As Long
    RASCONNSTATE As Long
    dwError As Long
    szDeviceType(RAS_MaxDeviceType) As Byte
    szDeviceName(RAS_MaxDeviceName) As Byte
End Type

Public Type RASDIALPARAMS
    dwSize As Long
    szEntryName(RAS_MaxEntryName) As Byte
    szPhoneNumber(RAS_MaxPhoneNumber) As Byte
    szCallbackNumber(RAS_MaxCallbackNumber) As Byte
    szUserName(UNLEN) As Byte
    szPassword(PWLEN) As Byte
    szDomain(DNLEN) As Byte
End Type

'****************************
'  RAS Functions - Declares
'****************************
Public Declare Function RasGetErrorString Lib "rasapi32.dll" Alias "RasGetErrorStringA" (ByVal uErrorValue As Long, ByVal lpszErrorString As String, ByVal cBufSize As Long) As Long
Public Declare Function RasGetConnectStatus Lib "rasapi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long
Public Declare Function RasDial Lib "rasapi32.dll" Alias "RasDialA" (lpRasDialExtensions As Any, ByVal lpszPhonebook As String, lpRasDialParams As Any, ByVal dwNotifierType As Long, ByVal hwndNotifier As Long, lphRasConn As Long) As Long
Public Declare Function RasHangUp Lib "rasapi32.dll" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long
Private Declare Function RasGetEntryDialParams Lib "rasapi32.dll" Alias "RasGetEntryDialParamsA" (ByVal lpszPhonebook As String, lpRasDialParams As Any, blnPasswordRetrieved As Long) As Long
Private Declare Function RasEnumEntries Lib "rasapi32.dll" Alias "RasEnumEntriesA" (ByVal lpStrNull As String, ByVal lpszPhonebook As String, lprasentryname As RASENTRYNAME, lpCb As Long, lpCEntries As Long) As Long
Public Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasConn As Any, lpCb As Long, lpcConnections As Long) As Long

Public Function RasGetConnectionSpeed() As Long
    On Error GoTo ErrorHandler
    If IsConnected = False Then GoSub ErrorHandler
    RasGetConnectionSpeed = ReadLong(HKEY_DYN_DATA, "PerfStats\StatData", "Dial-Up Adapter\ConnectSpeed", 0)
    Exit Function
ErrorHandler:
    RasGetConnectionSpeed = 0
End Function

Public Sub RasDisconnect(Optional hRas As Long = -1)
Dim rasConInfo(64) As RASCONN, lResult As Long
Dim lConnections As Long, lSize As Long, i As Integer
Dim iCalled As Integer
    On Error Resume Next    'make sure it hangs up
    rasConInfo(0).dwSize = LenB(rasConInfo(0))
    lSize = rasConInfo(0).dwSize * 64
    lResult = RasEnumConnections(rasConInfo(0), lSize, lConnections)
    If lResult <> 0 Then
        lngRASErrorNumber = lResult
        frmdial.Status.Caption = GetDunError
    Else
        If lConnections > 0 Then
            For i = 0 To lConnections - 1
                If hRas = -1 Or hRas = rasConInfo(i).hRasConn Then
                    iCalled = 0
                    Do
                        DoEvents
                        Inc iCalled
                        lResult = RasHangUp(rasConInfo(i).hRasConn)
                    Loop Until lResult = ERROR_INVALID_HANDLE Or iCalled = 1000
                End If
            Next
        End If
        'Setting hRasConn to 0 - no active connections
        hRasConn = 0
    End If
    'If hRas <> -1 Then hRas = 0
    frmdial.Timer2.Enabled = False
End Sub
Public Function dialnum(ByVal strEntry As String) As Boolean
On Error Resume Next
'dial th given number
Dim rasParams As RASDIALPARAMS, lResult As Long, hRas As Long
    lResult = RasGetDialParams(strEntry, rasParams)
    'Take care of alternate phone numbers
    lstrcpy rasParams.szUserName(0), frmdial.Combo3.Text
    lstrcpy rasParams.szPassword(0), frmdial.Text1.Text
    If tp = 1 Then
    lstrcpy rasParams.szPhoneNumber(0), frmdial.Combo2.Text
    Else
    lstrcpy rasParams.szPhoneNumber(0), "P" & frmdial.Combo2.Text
    End If
    Select Case lResult
        Case SUCCESS
            If RasDial(ByVal 0&, vbNullString, rasParams, RAS_NOTIFY_HWND, frmdial.hwnd, hRas) Then
            
            Else
            frmdial.Status.Caption = "Unable To Contact the Modem... Please Check The Connections..."
                dialnum = True
                hRasConn = hRas
            End If
        Case Else
            lngRASErrorNumber = lResult
            frmdial.Status.Caption = GetDunError
            RasDisconnect
    End Select
End Function

Public Function RasGetDialParams(strEntryName As String, rdp As RASDIALPARAMS, Optional blnPassword As Long) As Long
On Error Resume Next
'get the dial parameters
Dim bPassword As Long
    rdp.dwSize = LenB(rdp)
    lstrcpy rdp.szEntryName(0), strEntryName
    RasGetDialParams = RasGetEntryDialParams(vbNullString, rdp, bPassword)
    blnPassword = bPassword
End Function

Public Sub RasLoadEntries(Combo As ComboBox)
On Error Resume Next
'load data for dialing properties
Dim lResult As Long, lConns As Long, lSize As Long
Dim i As Integer, bexists As Integer
ReDim rasentry(64) As RASENTRYNAME
    rasentry(0).dwSize = LenB(rasentry(0))
    lSize = rasentry(0).dwSize * 64
    lResult = RasEnumEntries(0&, 0&, rasentry(0), lSize, lConns)
    Combo.Clear
    For i = 0 To lConns - 1
        Combo.AddItem ClearNulls(StrConv(rasentry(i).szEntryName, vbUnicode))
    Next
    Combo.ListIndex = 0
    'Check to see if there is at least one entry
    bexists = frmdial.Combo1.ListCount
If bexists < 1 Then
MsgBox "Please create a connection in the " & _
"Dial-Up Networking Folder, before attempting to connect to" + _
"the internet.", vbCritical + vbApplicationModal, "DialNet"
End
End If
End Sub
Public Function RasGetConnectedEntry() As String
On Error Resume Next
'get the state of the connection, no. of connections
Dim TRasCon(255) As RASCONN, lg As Long, lpcon As Long
Dim Tstatus As RASCONNSTATUS
    On Error Resume Next
    TRasCon(0).dwSize = LenB(TRasCon(0))
    lg = 256 * TRasCon(0).dwSize
    Call RasEnumConnections(TRasCon(0), lg, lpcon)
    Tstatus.dwSize = 160
    Call RasGetConnectStatus(TRasCon(0).hRasConn, Tstatus)
    If Tstatus.RASCONNSTATE <> RASCS_Disconnected And Tstatus.RASCONNSTATE <> 0 Then
        RasGetConnectedEntry = ClearNulls(StrConv(TRasCon(0).szEntryName, vbUnicode))
    Else
        RasGetConnectedEntry = ""
    End If
End Function
Public Function GetDunError() As String
On Error Resume Next
'get any DUN errors.
Dim lngRetCode As Long, iNullPos As Long
Dim strRASErrorString As String
    strRASErrorString = Space$(256)
    'lngRASErrorNumber is the RAS error number in class decl
    Select Case lngRASErrorNumber
        Case Is >= 600
            lngRetCode = RasGetErrorString(lngRASErrorNumber, strRASErrorString, 256&)
            If lngRetCode Then
                'We should never see this
                GetDunError = "Error: Unable to retrieve error message."
            Else
                'Return string
                 GetDunError = ClearNulls(strRASErrorString)
                 If GetDunError = "Unknown error." Then
                      ' An unknown error has occured
                      GetDunError = GetDunError + Str(lngRASErrorNumber)
                      ' See if is a common error
                      Select Case lngRASErrorNumber
                          Case 635
                              GetDunError = "Incorrect password. Server connection canceled."
                      End Select
                 End If
            End If
        Case Else
            GetDunError = "Unexpected Error. Error code" + Str(lngRASErrorNumber) + "."
     End Select
     iNullPos = InStr(GetDunError, Chr(0))
     If iNullPos > 1 Then GetDunError = Left$(GetDunError, iNullPos - 1)
End Function
Public Function RasGetConnectionState(ByVal hRas As Long) As Long
Dim lResult As Long, rasStatus As RASCONNSTATUS
    rasStatus.dwSize = LenB(rasStatus)
    lResult = RasGetConnectStatus(hRas, rasStatus)
    If lResult = 0 Then
        RasGetConnectionState = rasStatus.RASCONNSTATE
    Else
        RasGetConnectionState = RASCS_Disconnected
    End If
End Function
