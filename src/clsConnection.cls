VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================================
' clsConnection.cls
' Tek bir TCP socket'ini saran asenkron baglanti nesnesi.
'  - Listen / ConnectTo / AttachAccepted
'  - Alinan baytlardan paketleri cerceveler (tasma korumali parser)
'  - Gonderimde kismi send'leri kuyruklar (UI bloke olmaz)
'  - Olaylar: Connected, Disconnected, SocketErr, PacketReceived,
'    ClientAccepted
' Kaynak yonetimi: CloseNow tum socket/buffer'lari serbest birakir.
'=====================================================================
Option Explicit

Public Event Connected()
Public Event Disconnected()
Public Event SocketErr(ByVal code As Long, ByVal desc As String)
Public Event PacketReceived(ByVal pkt As clsPacket)
Public Event ClientAccepted(ByVal client As clsConnection)

Private Const RECV_CHUNK As Long = 8192
Private Const EVENT_MASK_CLIENT As Long = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE
Private Const EVENT_MASK_SERVER As Long = FD_ACCEPT Or FD_CLOSE

Private mSock As Long
Private mState As ConnState
Private mIsServer As Boolean
Private mPeerIP As String
Private mPeerPort As Long

' Alim (receive) tamponu
Private mRx() As Byte
Private mRxLen As Long
Private mRxCap As Long

' Gonderim (transmit) tamponu
Private mTx() As Byte
Private mTxStart As Long
Private mTxEnd As Long

Private Sub Class_Initialize()
    mSock = INVALID_SOCKET
    mState = CS_IDLE
    mRxCap = RECV_CHUNK
    ReDim mRx(0 To mRxCap - 1)
    mRxLen = 0
    ReDim mTx(0 To RECV_CHUNK - 1)
    mTxStart = 0
    mTxEnd = 0
End Sub

Private Sub Class_Terminate()
    CloseNow
End Sub

'--- Ozellikler ---
Public Property Get State() As ConnState
    State = mState
End Property
Public Property Get IsServer() As Boolean
    IsServer = mIsServer
End Property
Public Property Get PeerIP() As String
    PeerIP = mPeerIP
End Property
Public Property Get PeerPort() As Long
    PeerPort = mPeerPort
End Property
Public Property Get IsConnected() As Boolean
    IsConnected = (mState = CS_CONNECTED)
End Property

'=====================================================================
' Dinlemeye basla (Host modu)
'=====================================================================
Public Function Listen(ByVal port As Long) As Boolean
    On Error GoTo Fail
    mIsServer = True

    mSock = api_socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If mSock = INVALID_SOCKET Then GoTo Fail

    Dim opt As Long
    opt = 1
    api_setsockopt mSock, SOL_SOCKET, SO_REUSEADDR, opt, 4

    Dim sa As sockaddr_in
    sa.sin_family = AF_INET
    sa.sin_port = api_htons(port)
    sa.sin_addr = INADDR_ANY
    If api_bind(mSock, sa, Len(sa)) = SOCKET_ERROR Then
        LogWrite LVL_ERROR, "bind basarisiz, port " & port & " (WSA " & WSAGetLastError() & ")"
        GoTo Fail
    End If

    If api_listen(mSock, SOMAXCONN) = SOCKET_ERROR Then GoTo Fail

    NetRegisterSocket mSock, Me
    WSAAsyncSelect mSock, NetMsgHwnd(), WM_SOCKET, EVENT_MASK_SERVER
    mState = CS_LISTENING
    LogWrite LVL_INFO, "Dinleniyor: port " & port
    Listen = True
    Exit Function
Fail:
    CloseNow
    Listen = False
End Function

'=====================================================================
' Uzak host'a baglan (Viewer modu). ip = noktali IPv4 ("192.168.1.50")
'=====================================================================
Public Function ConnectTo(ByVal ip As String, ByVal port As Long) As Boolean
    On Error GoTo Fail
    mIsServer = False

    Dim addr As Long
    addr = api_inet_addr(ip)
    If addr = INADDR_NONE Then
        RaiseEvent SocketErr(0, "Gecersiz IP adresi: " & ip)
        ConnectTo = False
        Exit Function
    End If

    mSock = api_socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If mSock = INVALID_SOCKET Then GoTo Fail

    Dim nd As Long
    nd = 1
    api_setsockopt mSock, IPPROTO_TCP, TCP_NODELAY, nd, 4

    NetRegisterSocket mSock, Me
    ' Once async-select (FD_CONNECT icin), sonra connect.
    WSAAsyncSelect mSock, NetMsgHwnd(), WM_SOCKET, EVENT_MASK_CLIENT

    Dim sa As sockaddr_in
    sa.sin_family = AF_INET
    sa.sin_port = api_htons(port)
    sa.sin_addr = addr

    mPeerIP = ip
    mPeerPort = port
    mState = CS_CONNECTING

    Dim r As Long
    r = api_connect(mSock, sa, Len(sa))
    ' Non-blocking socket'te connect genellikle WSAEWOULDBLOCK doner;
    ' sonuc FD_CONNECT olayinda gelir. Bu beklenen durumdur.
    If r = SOCKET_ERROR Then
        Dim e As Long
        e = WSAGetLastError()
        If e <> WSAEWOULDBLOCK Then
            RaiseEvent SocketErr(e, "Baglanti baslatilamadi (WSA " & e & ")")
            GoTo Fail
        End If
    End If
    LogWrite LVL_INFO, "Baglaniliyor: " & ip & ":" & port
    ConnectTo = True
    Exit Function
Fail:
    CloseNow
    ConnectTo = False
End Function

'=====================================================================
' accept() ile gelen hazir socket'i sar (listener tarafindan cagrilir)
'=====================================================================
Friend Sub AttachAccepted(ByVal sock As Long, ByVal ip As String, ByVal port As Long)
    mSock = sock
    mIsServer = False
    mPeerIP = ip
    mPeerPort = port

    Dim nd As Long
    nd = 1
    api_setsockopt mSock, IPPROTO_TCP, TCP_NODELAY, nd, 4

    NetRegisterSocket mSock, Me
    WSAAsyncSelect mSock, NetMsgHwnd(), WM_SOCKET, EVENT_MASK_CLIENT
    mState = CS_CONNECTED
End Sub

'=====================================================================
' Gonderim yardimcilari
'=====================================================================
Public Sub SendBytes(ByVal pktType As Long, ByRef payload() As Byte, ByVal payloadLen As Long, Optional ByVal flags As Long = 0)
    If mState <> CS_CONNECTED Then Exit Sub
    Dim frame() As Byte
    frame = NetBuildFrame(pktType, 0, flags, payload, payloadLen)
    EnqueueTx frame, UBound(frame) - LBound(frame) + 1
End Sub

Public Sub SendString(ByVal pktType As Long, ByVal s As String, Optional ByVal flags As Long = 0)
    Dim b() As Byte, n As Long
    b = StringToUtf8(s, n)
    SendBytes pktType, b, n, flags
End Sub

Public Sub SendEmpty(ByVal pktType As Long, Optional ByVal flags As Long = 0)
    Dim b(0 To 0) As Byte
    SendBytes pktType, b, 0, flags
End Sub

'=====================================================================
' Socket olay geri cagrimlari (modNetwork.DispatchSocketEvent'ten)
'=====================================================================
Friend Sub OnConnect(ByVal errCode As Long)
    If errCode <> 0 Then
        RaiseEvent SocketErr(errCode, ErrText(errCode))
        CloseNow
        Exit Sub
    End If
    mState = CS_CONNECTED
    LogWrite LVL_INFO, "Baglandi: " & mPeerIP & ":" & mPeerPort
    RaiseEvent Connected
    FlushTx
End Sub

Friend Sub OnAccept()
    On Error GoTo SafeExit
    ' Bekleyen tum baglantilari kabul et.
    Do
        Dim sa As sockaddr_in, alen As Long, s As Long
        alen = Len(sa)
        s = api_accept(mSock, sa, alen)
        If s = INVALID_SOCKET Then Exit Do

        Dim client As clsConnection
        Set client = New clsConnection
        client.AttachAccepted s, IpToString(sa.sin_addr), api_ntohs(sa.sin_port)
        LogWrite LVL_INFO, "Gelen baglanti: " & client.PeerIP
        RaiseEvent ClientAccepted(client)
    Loop
    Exit Sub
SafeExit:
    LogWrite LVL_WARNING, "OnAccept hata: " & Err.Description
End Sub

Friend Sub OnRead()
    On Error GoTo SafeExit
    Dim tmp() As Byte, got As Long
    ReDim tmp(0 To RECV_CHUNK - 1)
    Do
        got = api_recv(mSock, tmp(0), RECV_CHUNK, 0)
        If got > 0 Then
            AppendRx tmp, got
        ElseIf got = 0 Then
            ' Karsi taraf duzgun kapatti.
            OnClose 0
            Exit Sub
        Else
            Dim e As Long
            e = WSAGetLastError()
            If e = WSAEWOULDBLOCK Then Exit Do   ' su an baska veri yok
            RaiseEvent SocketErr(e, ErrText(e))
            CloseNow
            Exit Sub
        End If
    Loop
    ParseBuffer
    Exit Sub
SafeExit:
    LogWrite LVL_ERROR, "OnRead hata: " & Err.Description
End Sub

Friend Sub OnWrite()
    FlushTx
End Sub

Friend Sub OnClose(ByVal errCode As Long)
    If mState = CS_CLOSED Then Exit Sub
    CloseNow
    RaiseEvent Disconnected
End Sub

'=====================================================================
' Alinan tamponu paketlere ayir (tasma korumali)
'=====================================================================
Private Sub ParseBuffer()
    Do While mRxLen >= PKT_HEADER_SIZE
        ' Magic + surum dogrula; bozuksa baglanti guvenli degildir -> kapat.
        If mRx(0) <> PKT_MAGIC0 Or mRx(1) <> PKT_MAGIC1 Then
            LogWrite LVL_WARNING, "Gecersiz paket imzasi; baglanti kapatiliyor."
            CloseNow
            RaiseEvent Disconnected
            Exit Sub
        End If

        Dim pType As Long, sessId As Long, pLen As Long
        pType = GetLong(mRx, 4)
        sessId = GetLong(mRx, 8)
        pLen = GetLong(mRx, 12)

        ' Payload uzunlugu siniri: tasma/bellek tuketimi korumasi.
        If pLen < 0 Or pLen > PKT_MAX_PAYLOAD Then
            LogWrite LVL_WARNING, "Gecersiz payload uzunlugu (" & pLen & "); baglanti kapatiliyor."
            CloseNow
            RaiseEvent Disconnected
            Exit Sub
        End If

        Dim need As Long
        need = PKT_HEADER_SIZE + pLen
        If mRxLen < need Then Exit Do     ' tam paket icin daha fazla veri bekle

        Dim pkt As clsPacket
        Set pkt = New clsPacket
        pkt.PType = pType
        pkt.SessionID = sessId
        pkt.Flags = mRx(3)
        If pLen > 0 Then
            Dim pay() As Byte
            ReDim pay(0 To pLen - 1)
            CopyMemory pay(0), mRx(PKT_HEADER_SIZE), pLen
            pkt.SetPayloadBytes pay, pLen
        End If

        RemoveFront need
        RaiseEvent PacketReceived(pkt)
    Loop
End Sub

'=====================================================================
' Tampon yonetimi
'=====================================================================
Private Sub AppendRx(ByRef src() As Byte, ByVal n As Long)
    If n <= 0 Then Exit Sub
    Dim newLen As Long
    newLen = mRxLen + n
    EnsureRxCap newLen
    CopyMemory mRx(mRxLen), src(0), n
    mRxLen = newLen
End Sub

Private Sub EnsureRxCap(ByVal need As Long)
    If need <= mRxCap Then Exit Sub
    Do While mRxCap < need
        mRxCap = mRxCap * 2
    Loop
    ReDim Preserve mRx(0 To mRxCap - 1)
End Sub

Private Sub RemoveFront(ByVal n As Long)
    If n >= mRxLen Then
        mRxLen = 0
    Else
        CopyMemory mRx(0), mRx(n), mRxLen - n
        mRxLen = mRxLen - n
    End If
End Sub

Private Sub EnqueueTx(ByRef data() As Byte, ByVal n As Long)
    If n <= 0 Then Exit Sub
    EnsureTxCap mTxEnd + n
    CopyMemory mTx(mTxEnd), data(LBound(data)), n
    mTxEnd = mTxEnd + n
    FlushTx
End Sub

Private Sub EnsureTxCap(ByVal need As Long)
    Dim cap As Long
    cap = UBound(mTx) + 1
    If need <= cap Then Exit Sub
    Do While cap < need
        cap = cap * 2
    Loop
    ReDim Preserve mTx(0 To cap - 1)
End Sub

Private Sub FlushTx()
    If mState <> CS_CONNECTED Then Exit Sub
    Do While mTxStart < mTxEnd
        Dim sent As Long
        sent = api_send(mSock, mTx(mTxStart), mTxEnd - mTxStart, 0)
        If sent = SOCKET_ERROR Then
            Dim e As Long
            e = WSAGetLastError()
            If e = WSAEWOULDBLOCK Then Exit Sub    ' FD_WRITE'ta devam edilir
            RaiseEvent SocketErr(e, ErrText(e))
            CloseNow
            RaiseEvent Disconnected
            Exit Sub
        ElseIf sent = 0 Then
            Exit Sub
        Else
            mTxStart = mTxStart + sent
        End If
    Loop
    ' Hepsi gonderildi -> tamponu sifirla.
    mTxStart = 0
    mTxEnd = 0
End Sub

'=====================================================================
' Kapanis: socket ve tamponlari serbest birak.
'=====================================================================
Public Sub CloseNow()
    On Error Resume Next
    If mSock <> INVALID_SOCKET Then
        WSAAsyncSelect mSock, NetMsgHwnd(), 0, 0   ' bildirimleri durdur
        NetUnregisterSocket mSock
        api_closesocket mSock
        mSock = INVALID_SOCKET
    End If
    mRxLen = 0
    mTxStart = 0
    mTxEnd = 0
    mState = CS_CLOSED
End Sub

'--- IPv4 Long -> "a.b.c.d" (ag bayt sirasi) ---
Private Function IpToString(ByVal ipLong As Long) As String
    Dim b(0 To 3) As Byte
    CopyMemory b(0), ipLong, 4
    IpToString = b(0) & "." & b(1) & "." & b(2) & "." & b(3)
End Function

Private Function ErrText(ByVal code As Long) As String
    Select Case code
        Case WSAECONNREFUSED: ErrText = "Baglanti reddedildi (karsi taraf dinlemiyor?)"
        Case WSAECONNRESET:   ErrText = "Baglanti karsi tarafca sifirlandi"
        Case WSAETIMEDOUT:    ErrText = "Baglanti zaman asimina ugradi"
        Case Else:            ErrText = "Ag hatasi (WSA " & code & ")"
    End Select
End Function
