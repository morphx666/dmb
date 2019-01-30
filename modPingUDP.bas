Attribute VB_Name = "modPingUDP"
Option Explicit

Public Function PingUDP(ByVal hostName As String) As Single

    Dim sO As Long
    Dim ao As sockaddr
    Dim bufo As String
    
    Dim si As Long
    Dim ai As sockaddr
    Dim bufi As String
    Dim rai As sockaddr
    
    Dim sIP As String
    Dim eip As String
    
    Dim ttl As Long
    Dim r As Long
    Dim ha As Long
    
    Dim t As Single
    Dim tt As Single
    Dim j As Integer
      
    sIP = GetIPFromHostName(hostName, True)
    If sIP = "0.0.0.0" Then GoTo ExitWithErr
    ha = GetAddressFromHost(hostName)
    bufo = String(32, "a")
    
    ttl = 255
    
    With ao
        .sin_family = AF_INET
        .sin_port = htons(33434)
        .sin_addr = ha
    End With
    sO = socket(PF_INET, SOCK_DGRAM, IPPROTO_UDP)
    If sO < 0 Then GoTo ExitWithErr
    r = setsockopt(sO, SOL_SOCKET, SO_RCVTIMEO, 300, Len(ttl))
    If r > 0 Then GoTo ExitWithErr
    r = setsockopt(sO, IPPROTO_IP, IP_TTL, ttl, Len(ttl))
    If r > 0 Then GoTo ExitWithErr
    
    With ai
        .sin_family = AF_INET
        .sin_port = htons(33434)
        .sin_addr = htons(INADDR_ANY)
    End With
    si = socket(PF_INET, SOCK_RAW, IPPROTO_ICMP)
    If sO < 0 Then GoTo ExitWithErr
    r = bind(si, ai, Len(ai))
    If r > 0 Then GoTo ExitWithErr
    r = setsockopt(si, SOL_SOCKET, SO_RCVTIMEO, 300, Len(ttl))
    If r > 0 Then GoTo ExitWithErr

    tt = 0
    For j = 1 To 3
        t = Timer
            r = sendto(sO, bufo, Len(bufo), 0&, ao, Len(ao))
            
            bufi = Space(512)
            r = recvfrom(si, bufi, Len(bufi), 0&, rai, Len(rai))
        tt = tt + (Timer - t)
    Next j
    tt = Round((tt / 3) * 1000, 2)
    If tt < 0 Then tt = 0
    
    closesocket si
    closesocket sO
    
    eip = GetIPFromAddress(rai.sin_addr)
        
    PingUDP = tt
    
    'ResolveHops
    
CleanUp:
    SocketsCleanup
    
    Exit Function
    
ExitWithErr:
    GoTo CleanUp

End Function
