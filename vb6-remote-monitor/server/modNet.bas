Attribute VB_Name = "modNet"
'=============================================================================
' modNet.bas  -  Saf Winsock API (ws2_32.dll) ile TCP soket yardimcilari
' OCX YOK. MSWINSCK.OCX kullanilmaz. Tum soket isleri API ile yapilir.
' Not: Bu modul Server ve Client projelerinde ayni sekilde kullanilir.
'=============================================================================
Option Explicit

Public Const AF_INET As Long = 2
Public Const SOCK_STREAM As Long = 1
Public Const IPPROTO_TCP As Long = 6
Public Const INADDR_ANY As Long = 0
Public Const INADDR_NONE As Long = &HFFFFFFFF
Public Const SOCKET_ERROR As Long = -1
Public Const INVALID_SOCKET As Long = -1
Public Const SD_BOTH As Long = 2
Public Const FIONBIO As Long = &H8004667E
Public Const WSAEWOULDBLOCK As Long = 10035
Public Const SOMAXCONN As Long = 5

Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * 257
    szSystemStatus As String * 129
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Type sockaddr_in
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(0 To 7) As Byte
End Type

Private Type fd_set
    fd_count As Long
    fd_array(0 To 63) As Long
End Type

Private Type timeval
    tv_sec As Long
    tv_usec As Long
End Type

Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
Private Declare Function api_socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal sType As Long, ByVal protocol As Long) As Long
Private Declare Function api_bind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, addr As sockaddr_in, ByVal namelen As Long) As Long
Private Declare Function api_listen Lib "ws2_32.dll" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function api_accept Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, addr As Any, addrlen As Any) As Long
Private Declare Function api_connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, addr As sockaddr_in, ByVal namelen As Long) As Long
Private Declare Function api_closesocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long
Private Declare Function api_send Lib "ws2_32.dll" Alias "send" (ByVal s As Long, buf As Any, ByVal length As Long, ByVal flags As Long) As Long
Private Declare Function api_recv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, buf As Any, ByVal length As Long, ByVal flags As Long) As Long
Private Declare Function api_ioctlsocket Lib "ws2_32.dll" Alias "ioctlsocket" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long
Private Declare Function api_select Lib "ws2_32.dll" Alias "select" (ByVal nfds As Long, readfds As Any, writefds As Any, exceptfds As Any, timeout As timeval) As Long
Private Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
Private Declare Function gethostbyname Lib "ws2_32.dll" (ByVal hostname As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal numBytes As Long)

Private mStarted As Boolean

'--- Winsock baslat (uygulama acilisinda bir kere cagir) ------------------
Public Function NetStartup() As Boolean
    Dim wsa As WSADATA
    If mStarted Then NetStartup = True: Exit Function
    If WSAStartup(&H202, wsa) = 0 Then
        mStarted = True
        NetStartup = True
    End If
End Function

Public Sub NetCleanup()
    If mStarted Then
        WSACleanup
        mStarted = False
    End If
End Sub

Public Function NetLastError() As Long
    NetLastError = WSAGetLastError()
End Function

'--- Soketi bloklamayan (non-blocking) moda al ---------------------------
Public Sub SetNonBlocking(ByVal s As Long, ByVal nonBlocking As Boolean)
    Dim arg As Long
    If nonBlocking Then arg = 1 Else arg = 0
    api_ioctlsocket s, FIONBIO, arg
End Sub

Public Sub CloseSock(ByRef s As Long)
    If s <> INVALID_SOCKET And s <> 0 Then
        api_closesocket s
    End If
    s = INVALID_SOCKET
End Sub

'--- Port numarasini network byte order (big-endian) olarak doldur -------
Private Sub FillPort(ByRef sa As sockaddr_in, ByVal port As Long)
    Dim b(0 To 1) As Byte
    b(0) = (port \ 256) And &HFF   ' yuksek byte once
    b(1) = port And &HFF
    CopyMemory sa.sin_port, b(0), 2
End Sub

'--- Dinleyen (listen) soket olustur -------------------------------------
Public Function CreateListener(ByVal port As Long) As Long
    Dim s As Long, sa As sockaddr_in
    s = api_socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If s = INVALID_SOCKET Then CreateListener = INVALID_SOCKET: Exit Function

    sa.sin_family = AF_INET
    FillPort sa, port
    sa.sin_addr = INADDR_ANY

    If api_bind(s, sa, LenB(sa)) = SOCKET_ERROR Then
        CloseSock s: CreateListener = INVALID_SOCKET: Exit Function
    End If
    If api_listen(s, SOMAXCONN) = SOCKET_ERROR Then
        CloseSock s: CreateListener = INVALID_SOCKET: Exit Function
    End If

    SetNonBlocking s, True
    CreateListener = s
End Function

'--- Bekleyen baglantiyi kabul et (yoksa INVALID_SOCKET doner) -----------
Public Function AcceptClient(ByVal listener As Long) As Long
    Dim s As Long
    s = api_accept(listener, ByVal 0&, ByVal 0&)
    AcceptClient = s
End Function

'--- Bir host adini (IP veya DynDNS) network-order IP'ye cevir -----------
Private Function ResolveHost(ByVal host As String) As Long
    Dim addr As Long
    addr = inet_addr(host)
    If addr <> INADDR_NONE Then
        ResolveHost = addr
        Exit Function
    End If
    ' Isim cozumleme (gethostbyname)
    Dim pHostent As Long, pAddrList As Long, pAddr As Long, ip As Long
    pHostent = gethostbyname(host)
    If pHostent = 0 Then ResolveHost = INADDR_NONE: Exit Function
    ' HOSTENT ofset 12 -> h_addr_list (char**)
    CopyMemory pAddrList, ByVal (pHostent + 12), 4
    If pAddrList = 0 Then ResolveHost = INADDR_NONE: Exit Function
    CopyMemory pAddr, ByVal pAddrList, 4   ' ilk adres pointer'i
    If pAddr = 0 Then ResolveHost = INADDR_NONE: Exit Function
    CopyMemory ip, ByVal pAddr, 4          ' 4 byte IP (zaten network order)
    ResolveHost = ip
End Function

'--- Sunucuya baglan (non-blocking). Donus: soket veya INVALID_SOCKET ----
' connect non-blocking ise genelde WSAEWOULDBLOCK verir; bu normaldir.
' Baglanti tamamlandi mi diye SockWritable ile kontrol edilir.
Public Function ConnectTo(ByVal host As String, ByVal port As Long) As Long
    Dim s As Long, sa As sockaddr_in, ip As Long
    ip = ResolveHost(host)
    If ip = INADDR_NONE Then ConnectTo = INVALID_SOCKET: Exit Function

    s = api_socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If s = INVALID_SOCKET Then ConnectTo = INVALID_SOCKET: Exit Function

    SetNonBlocking s, True
    sa.sin_family = AF_INET
    FillPort sa, port
    sa.sin_addr = ip

    Dim r As Long
    r = api_connect(s, sa, LenB(sa))
    ' r = 0 (hemen baglandi) ya da SOCKET_ERROR + WSAEWOULDBLOCK (devam ediyor)
    ConnectTo = s
End Function

'--- select ile okunabilirlik kontrolu -----------------------------------
Public Function SockReadable(ByVal s As Long) As Boolean
    Dim fds As fd_set, tv As timeval, r As Long
    fds.fd_count = 1
    fds.fd_array(0) = s
    tv.tv_sec = 0
    tv.tv_usec = 0
    r = api_select(0, fds, ByVal 0&, ByVal 0&, tv)
    SockReadable = (r > 0)
End Function

'--- select ile yazilabilirlik (baglanti kuruldu mu) kontrolu -----------
Public Function SockWritable(ByVal s As Long) As Boolean
    Dim fds As fd_set, tv As timeval, r As Long
    fds.fd_count = 1
    fds.fd_array(0) = s
    tv.tv_sec = 0
    tv.tv_usec = 0
    r = api_select(0, ByVal 0&, fds, ByVal 0&, tv)
    SockWritable = (r > 0)
End Function

'--- Tum tamponu gonder. Basari: True ------------------------------------
Public Function SendAll(ByVal s As Long, ByRef data() As Byte, ByVal length As Long) As Boolean
    Dim sent As Long, n As Long, guard As Long
    sent = 0
    Do While sent < length
        n = api_send(s, data(sent), length - sent, 0)
        If n = SOCKET_ERROR Then
            If WSAGetLastError() = WSAEWOULDBLOCK Then
                ' Gonderme tamponu dolu; kisa bekle ve tekrar dene
                guard = guard + 1
                If guard > 100000 Then SendAll = False: Exit Function
            Else
                SendAll = False
                Exit Function
            End If
        ElseIf n = 0 Then
            SendAll = False
            Exit Function
        Else
            sent = sent + n
        End If
    Loop
    SendAll = True
End Function

'--- Bir kareyi [4 byte uzunluk + veri] olarak tek seferde gonder ---------
Public Function SendFrame(ByVal s As Long, ByRef data() As Byte, ByVal length As Long) As Boolean
    Dim buf() As Byte
    ReDim buf(0 To length + 3)
    buf(0) = length And &HFF
    buf(1) = (length \ 256) And &HFF
    buf(2) = (length \ 65536) And &HFF
    buf(3) = (length \ 16777216) And &HFF
    CopyMemory buf(4), data(0), length
    SendFrame = SendAll(s, buf, length + 4)
End Function

'--- Gelen veriyi oku. Donus: okunan byte (>0), 0=kapandi, -1=simdilik yok
Public Function RecvSome(ByVal s As Long, ByRef buf() As Byte, ByVal maxLen As Long) As Long
    Dim n As Long
    n = api_recv(s, buf(0), maxLen, 0)
    If n = SOCKET_ERROR Then
        If WSAGetLastError() = WSAEWOULDBLOCK Then
            RecvSome = -1    ' su an veri yok
        Else
            RecvSome = 0     ' hata -> kapandi say
        End If
    Else
        RecvSome = n         ' n=0 ise karsi taraf kapatti
    End If
End Function
