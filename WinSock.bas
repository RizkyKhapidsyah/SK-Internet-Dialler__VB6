Attribute VB_Name = "WinSock"

'*********************************
'  WinSock Functions - Constants
'*********************************

Private Const POP3_GREETING As Integer = 0
Private Const POP3_USER As Integer = 1
Private Const POP3_PASS As Integer = 2
Private Const POP3_STAT As Integer = 3

Private Const SOCK_STREAM As Integer = 1
Private Const AF_INET As Integer = 2
Private Const INADDR_NONE = &HFFFFFFFF
Private Const FIONREAD = &H4004667F

Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&

'**********************************
'  WinSock Functions - Structures
'**********************************

Private Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type

Private Type HostEnt
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 256) As Byte
    szSystemStatus(0 To 128) As Byte
    wMaxSockets As Long
    wMaxUDPDG As Long
    dwVendorInfo As Long
End Type

Private Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type
   
Private Type ICMP_ECHO_REPLY
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    DataSize        As Integer
    Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type

'********************************
'  WinSock Functions - Declares
'********************************

Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal timeout As Long) As Long
Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequired As Long, lpWSAData As WSADATA) As Long
Private Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Private Declare Function gethostbyaddr Lib "ws2_32.dll" (lIPAddress As Long, ByVal iLen As Long, ByVal iType As Long) As Long
Private Declare Function inet_addr Lib "ws2_32.dll" (ByVal sIPAddress As String) As Long
Private Declare Function htons Lib "ws2_32.dll" (ByVal A As Integer) As Long
Private Declare Function ioctlsocket Lib "ws2_32.dll" (ByVal Sock As Long, ByVal cmd As Long, argp As Any) As Long
Private Declare Function socket Lib "ws2_32.dll" (ByVal afinet As Integer, ByVal socktype As Integer, ByVal protocol As Integer) As Long
Private Declare Function connect Lib "ws2_32.dll" (ByVal Sock As Integer, sockstruct As sockaddr, ByVal structlen As Long) As Integer
Private Declare Function send Lib "ws2_32.dll" (ByVal Sock As Long, ByVal Msg As String, ByVal msglen As Integer, ByVal flag As Integer) As Long
Private Declare Function Receive Lib "ws2_32.dll" Alias "recv" (ByVal Sock As Long, ByVal Msg As String, ByVal msglen As Integer, ByVal flag As Integer) As Long
Private Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Integer

Public Function IsIP(ByVal strIPAddress As String) As Boolean
Dim s As String, i As Integer, pos As Integer
    IsIP = False
    s = Format(strIPAddress, "############")
    If Not IsNumeric(s) Or Len(s) <> Len(strIPAddress) - 3 Then Exit Function
    If Mid(strIPAddress, 1, 1) = "." Or Mid(strIPAddress, Len(strIPAddress), 1) = "." Then Exit Function
    i = 0
    pos = 0
    Do
        pos = InStr(pos + 1, strIPAddress, ".")
        If pos <> 0 Then i = i + 1
    Loop Until i = 3 Or pos = 0
    If i <> 3 Then Exit Function
    If inet_addr(strIPAddress) = INADDR_NONE Then Exit Function
    IsIP = True
End Function

Private Function SocketsInitialize() As Boolean
Dim WSAD As WSADATA, X As Integer, szLoByte As String, szHiByte As String, szBuf As String
    X = WSAStartup(WS_VERSION_REQD, WSAD)
    If X <> 0 Then
        MsgBox "Windows Sockets For 32 bit Windows environments is not successfully responding."
        Exit Function
    End If
    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
        szHiByte = Trim$(Str$(HiByte(WSAD.wVersion)))
        szLoByte = Trim$(Str$(LoByte(WSAD.wVersion)))
        szBuf = "Windows Sockets Version " & szLoByte & "." & szHiByte
        szBuf = szBuf & " is not supported by Windows " & _
        "Sockets For 32 bit Windows environments."
        MsgBox szBuf, vbExclamation
        Exit Function
    End If
    If WSAD.wMaxSockets < 1 Then
        szBuf = "This application requires a minimum of " & _
        Trim$(Str$(1)) & " supported sockets."
        MsgBox szBuf, vbExclamation
        Exit Function
    End If
    SocketsInitialize = True
End Function

Private Sub SocketsCleanup()
Dim X As Long
    X = WSACleanup()
    If X <> 0 Then
        MsgBox "Windows Sockets Error " & Trim$(Str$(X)) & " occurred in Cleanup.", vbExclamation
    End If
End Sub

Private Function HiByte(ByVal wParam As Long) As Integer
    HiByte = wParam \ &H100 And &HFF&
End Function

Private Function LoByte(ByVal wParam As Long) As Integer
    LoByte = wParam And &HFF&
End Function

Public Function ResolveHostIP(ByVal HostName As String) As Collection

    Dim hostent_addr As Long
    Dim Host As HostEnt
    Dim hostip_addr As Long
    Dim temp_ip_address() As Byte
    Dim i As Integer
    Dim ip_address As String
    Dim Count As Integer


    If SocketsInitialize() Then
        
        Set ResolveHostIP = New Collection
        hostent_addr = gethostbyname(HostName)


        If hostent_addr = 0 Then
            SocketsCleanup
            Exit Function
        End If

        RtlMoveMemory Host, hostent_addr, LenB(Host)
        RtlMoveMemory hostip_addr, Host.hAddrList, 4
        'get all of the IP address if machine is
        '     multi-homed


        Do
            ReDim temp_ip_address(1 To Host.hLength)
            RtlMoveMemory temp_ip_address(1), hostip_addr, Host.hLength


            For i = 1 To Host.hLength
                ip_address = ip_address & temp_ip_address(i) & "."
            Next

            ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
            ResolveHostIP.Add ip_address
            ip_address = ""
            Host.hAddrList = Host.hAddrList + LenB(Host.hAddrList)
            RtlMoveMemory hostip_addr, Host.hAddrList, 4
        Loop While (hostip_addr <> 0)

        
        SocketsCleanup
    End If

End Function
Public Function GetHostAlias(ByVal sIPAddress As String) As String
Dim addr As Long, Host As HostEnt, lResult As Long
Dim sHost As String
    If SocketsInitialize() Then
        addr = inet_addr(sIPAddress)
        lResult = gethostbyaddr(addr, 4, 4)
        If lResult <> 0 Then
            RtlMoveMemory Host, lResult, Len(Host)
            sHost = String$(255, Chr(0))
            RtlMoveMemory ByVal sHost, ByVal Host.hName, 255
        Else: Exit Function
        End If
        SocketsCleanup
    End If
    GetHostAlias = ClearNulls(sHost)
End Function

Public Function CreateSocket() As Integer
    CreateSocket = socket(AF_INET, SOCK_STREAM, 0)
End Function

Public Function ConnectSocket(ByVal nSocket As Long, ByVal sServer As String, ByVal iPort As Integer) As Integer
Dim sa As sockaddr, sAddr As Long
    sAddr = GetNetworkIP(sServer)
    With sa
        .sin_family = AF_INET
        .sin_zero = String$(8, 0)
        .sin_port = htons(iPort)
        .sin_addr = sAddr
    End With
    DoEvents
    ConnectSocket = connect(nSocket, sa, Len(sa))
End Function

Private Function SendData(ByVal nSocket As Long, sBuffer As String) As Integer
    DoEvents
    SendData = send(nSocket, ByVal sBuffer, Len(sBuffer), 0)
End Function

Private Function ReceiveData(ByVal nSocket As Long, sBuffer As String) As Long
Dim s As String * 25, i As Long
    sBuffer = ""
    Do
        If GetInputState() Then DoEvents
        i = Receive(nSocket, ByVal s, Len(s), 0)
        If i > 0 Then
            sBuffer = sBuffer & Left$(s, i)
        End If
    Loop Until Not IsThereData(nSocket)
    ReceiveData = i
End Function

Private Function GetNetworkIP(ByVal sAddress As String) As Long
Dim lHostent As Long, he As HostEnt
Dim iaddr As Long, lIP As Long
Dim b(1 To 4) As Byte, i As Integer, s As String
Dim c As Collection
    If IsIP(sAddress) Then
        lIP = inet_addr(sAddress)
    Else
        Set c = ResolveHostIP(sAddress)
        If c.Count > 0 Then s = c.Item(1)
        lIP = inet_addr(s)
    End If
    GetNetworkIP = lIP
End Function

Private Function IsThereData(iSocket As Long) As Boolean
Dim iResult As Long, lBuffer As Long
    iResult = ioctlsocket(iSocket, FIONREAD, lBuffer)
    If iResult Then
        IsThereData = False
    Else
        IsThereData = (lBuffer <> 0)
    End If
End Function

Public Sub CheckMail()
Dim nSocket As Long, lResult As Long, sData As String
Dim sSend As String, iCount As Integer, bRetry As Boolean
    If SocketsInitialize Then
        'Creating socket
        nSocket = CreateSocket
        If nSocket < 1 Then
            frmMain.sb1.Panels.Item(1).Text = "Unable to create socket."
            GoSub CleanupStuff
        End If
        'Connecting socket
        frmMain.sb1.Panels.Item(1).Text = "Connecting to " & sMailServer
        If ConnectSocket(nSocket, sMailServer, iMailPort) Then
            If MsgBox("Unable to connect to " & sMailServer & " on port " & CStr(iMailPort) & ". Check your Internet connection and make sure the settings are correct.", vbCritical + vbRetryCancel, "Unable to connect") = vbRetry Then bRetry = True
            GoSub CleanupStuff
        End If
        Call ReceiveData(nSocket, sData)
        If CheckPOP3Response(POP3_GREETING, sData) = False Then
            frmMain.sb1.Panels.Item(1).Text = "Server is not ready to accept requests."
            GoSub CleanupStuff
        End If
        frmMain.sb1.Panels.Item(1).Text = "Sending Username..."
        sSend = "USER " & sMailAccount & vbCrLf
        If SendData(nSocket, sSend) <> Len(sSend) Then
            frmMain.sb1.Panels.Item(1).Text = "Unable to communicate with server."
            GoSub CleanupStuff
        End If
        Call ReceiveData(nSocket, sData)
        If CheckPOP3Response(POP3_USER, sData) = False Then
            frmMain.sb1.Panels.Item(1).Text = "The account name you entered is invalid."
            GoSub CleanupStuff
        End If
        sSend = "PASS " & sMailPassword & vbCrLf
        frmMain.sb1.Panels.Item(1).Text = "Sending Password..."
        If SendData(nSocket, sSend) <> Len(sSend) Then
            frmMain.sb1.Panels.Item(1).Text = "Unable to communicate with server."
            GoSub CleanupStuff
        End If
        Call ReceiveData(nSocket, sData)
        Do While CheckPOP3Response(POP3_PASS, sData) = False
                frmMain.sb1.Panels.Item(1).Text = "Invalid Password. Please reenter u r password"
                sSend = InputBox("Enter your email account password:", "Password required", sMailPassword)
                frmMain.sb1.Panels.Item(1).Text = "Sending Password..."
        
            If StrPtr(sSend) = 0 Then
                frmMain.sb1.Panels.Item(1).Text = "Unable to check for new nessages."
                GoSub CleanupStuff
            End If
            sMailPassword = sSend
            sSend = "PASS " & sSend & vbCrLf
            If SendData(nSocket, sSend) <> Len(sSend) Then
                frmMain.sb1.Panels.Item(1).Text = "Unable to communicate with server."
                GoSub CleanupStuff
            End If
            Call ReceiveData(nSocket, sData)
        Loop
        sSend = "STAT " & vbCrLf
        If SendData(nSocket, sSend) <> Len(sSend) Then
            frmMain.sb1.Panels.Item(1).Text = "Unable to communicate with server."
            GoSub CleanupStuff
        End If
        Call ReceiveData(nSocket, sData)
        If CheckPOP3Response(POP3_STAT, sData, iCount) = False Then
            frmMain.sb1.Panels.Item(1).Text = "You Have" & CStr(iCount) & " new message(s)."
            GoSub CleanupStuff
        Else
            frmMain.sb1.Panels.Item(1).Text = "You have " & CStr(iCount) & " new message(s) in your mailbox."
        End If
        '
        '
        '
        '
        frmMain.sb1.Panels.Item(1).Text = "Retrieving message " & CStr(1) & ". . . "
        MsgBuffer = MsgBuffer & Data
        If InStr(1, MsgBuffer, vbLf & "." & vbCrLf) > 0 Then
                       vHeader = Split(MsgBuffer, vbCrLf)
                For Each vField In vHeader
                        msgfield = CStr(vField)
                        iPos = InStr(1, msgfield, ":")
                        If iPos Then
                            msgtest = LCase(Left(msgfield, iPos - 1))
                        Else
                            msgtest = ""
                        End If
                Select Case msgtest
                        Case "from"
                            msgfrom = Mid$(msgfield, iPos + 1)
                        Case "subject"
                            msgsubject = Mid$(msgfield, iPos + 1)
                        End Select
                    Next
                    MsgBuffer = ""
                    
                    
                End If
                'msg list
                msgSize = Val(Mid$(sData, 5 + InStr(1, Mid$(sData, 5), " ")))
                
                Dim lvitem
                Set lvitem = frmchk.lvMsg.ListItems.Add
                lvitem.Key = "msg" & currmsg
                lvitem.Text = msgfrom
                lvitem.SubItems(1) = msgsubject
                lvitem.SubItems(2) = CStr(msgSize)
                 
                msgfrom = ""
                msgsubject = ""
                currmsg = currmsg + 1
                If currmsg > nummsg Then frmMain.sb1.Panels.Item(1) = "Done"
                
                
                
            sSend = "QUIT " & vbCrLf
        If SendData(nSocket, sSend) <> Len(sSend) Then
            frmMain.sb1.Panels.Item(1).Text = "Unable to send disconnect data."
            GoSub CleanupStuff
        End If
        Call ReceiveData(nSocket, sData)
    Else
        frmdial.status.Caption = "Unable to initialize WinSock module."
    End If
CleanupStuff:
    Call closesocket(nSocket)
    Call SocketsCleanup
    If bRetry Then Call CheckMail
End Sub

Private Function CheckPOP3Response(ByVal iResponseType As Integer, ByVal sResponse As String, Optional iValue As Integer) As Boolean
Dim iPos As Integer, s As String
    If Len(sResponse) = 0 Then Exit Function
    Select Case iResponseType
        Case POP3_STAT
            If Mid(sResponse, 1, 1) = "+" Then
                iPos = InStr(1, sResponse, " ")
                iValue = Val(Mid(sResponse, iPos + 1, InStr(iPos + 1, sResponse, " ") - 5))
                s = Mid(sResponse, iPos + 1, InStr(iPos + 1, sResponse, " ") - 5)
                CheckPOP3Response = True
            End If
        Case Else
            If Mid(sResponse, 1, 1) = "+" Then CheckPOP3Response = True
    End Select
End Function
