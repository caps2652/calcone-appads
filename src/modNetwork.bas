Attribute VB_Name = "modNetwork"
'=====================================================================
' modNetwork.bas
' Ag katmaninin cekirdegi (OCX'siz, ham Winsock 2 API):
'  - Winsock baslatma/kapatma
'  - Asenkron socket bildirimleri icin "message-only" pencere + subclass
'  - WndProc (AddressOf) -> olaylari ilgili clsConnection'a dagitir
'  - Paket cerceve olusturma, UTF-8 ve byte yardimcilari
' Boylece UI thread'i hicbir zaman ag yuzunden bloke olmaz
' (WSAAsyncSelect socket'i otomatik non-blocking yapar).
'
' Katman ayrik tutulur ki ileride Direct LAN / Internet Relay /
' Cloud Server / NAT Traversal eklenebilsin.
'=====================================================================
Option Explicit

'--- Protokol sabitleri ---
Public Const PKT_MAGIC0 As Byte = &H52        ' "R"
Public Const PKT_MAGIC1 As Byte = &H53        ' "S"
Public Const PKT_VERSION As Byte = 1
Public Const PKT_HEADER_SIZE As Long = 16
Public Const PKT_MAX_PAYLOAD As Long = 16777216   ' 16 MB ust sinir (tasma korumasi)

'--- Paket turleri (protokol boyunca sabit anlamlar) ---
Public Enum PacketType
    PT_HELLO = 1
    PT_WELCOME = 2
    PT_AUTH_REQUEST = 10
    PT_AUTH_CHALLENGE = 11
    PT_AUTH_RESPONSE = 12
    PT_AUTH_RESULT = 13
    PT_PERMISSION_REQUEST = 20
    PT_PERMISSION_GRANT = 21
    PT_PERMISSION_DENY = 22
    PT_SCREEN_FRAME = 30
    PT_MOUSE_MOVE = 40
    PT_MOUSE_CLICK = 41
    PT_MOUSE_WHEEL = 42
    PT_KEYBOARD = 50
    PT_CLIPBOARD = 60
    PT_FILE_META = 70
    PT_FILE_CHUNK = 71
    PT_FILE_ACK = 72
    PT_CHAT = 80
    PT_PING = 90
    PT_PONG = 91
    PT_DISCONNECT = 99
End Enum

'--- Baglanti durumlari ---
Public Enum ConnState
    CS_IDLE = 0
    CS_LISTENING = 1
    CS_CONNECTING = 2
    CS_CONNECTED = 3
    CS_CLOSED = 4
End Enum

Private Const CP_UTF8 As Long = 65001

Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

Private mStarted As Boolean
Private mMsgHwnd As Long
Private mOldProc As Long
Private mConns As Collection       ' CStr(socket) -> clsConnection

'=====================================================================
' Baslatma / kapatma
'=====================================================================
Public Function NetInit() As Boolean
    On Error GoTo Fail
    Dim wsa As WSADATA
    If WSAStartup(&H202, wsa) <> 0 Then     ' Winsock 2.2
        LogWrite LVL_ERROR, "WSAStartup basarisiz."
        NetInit = False
        Exit Function
    End If

    Set mConns = New Collection

    ' Socket bildirimlerini alacak gizli (message-only) pencere olustur.
    Dim hInst As Long
    hInst = GetModuleHandleW(0)
    Dim cls As String
    cls = "STATIC"
    mMsgHwnd = CreateWindowExW(0, StrPtr(cls), 0, 0, 0, 0, 0, 0, HWND_MESSAGE, 0, hInst, ByVal 0&)
    If mMsgHwnd = 0 Then
        LogWrite LVL_ERROR, "Mesaj penceresi olusturulamadi."
        WSACleanup
        NetInit = False
        Exit Function
    End If

    ' Pencereyi subclass'la: socket mesajlarini kendi WndProc'umuza yonlendir.
    mOldProc = SetWindowLongW(mMsgHwnd, GWL_WNDPROC, AddressOf SocketWndProc)

    mStarted = True
    NetInit = True
    LogWrite LVL_INFO, "Ag alt sistemi hazir (Winsock 2.2)."
    Exit Function
Fail:
    LogWrite LVL_ERROR, "NetInit hata: " & Err.Description
    NetInit = False
End Function

Public Sub NetShutdown()
    On Error Resume Next
    If Not mStarted Then Exit Sub

    ' Kayitli tum socket'leri kapat.
    If Not mConns Is Nothing Then
        Dim v As Variant, c As clsConnection
        For Each v In mConns
            Set c = v
            If Not c Is Nothing Then c.CloseNow
        Next v
        Set mConns = Nothing
    End If

    ' Subclass'i geri al ve pencereyi yok et.
    If mMsgHwnd <> 0 Then
        If mOldProc <> 0 Then SetWindowLongW mMsgHwnd, GWL_WNDPROC, mOldProc
        DestroyWindow mMsgHwnd
        mMsgHwnd = 0
    End If

    WSACleanup
    mStarted = False
    LogWrite LVL_INFO, "Ag alt sistemi kapatildi."
End Sub

Public Function NetMsgHwnd() As Long
    NetMsgHwnd = mMsgHwnd
End Function

'=====================================================================
' Socket kayit defteri (socket -> clsConnection eslestirme)
'=====================================================================
Public Sub NetRegisterSocket(ByVal sock As Long, ByVal conn As clsConnection)
    On Error Resume Next
    mConns.Remove CStr(sock)           ' varsa eskisini temizle
    On Error GoTo 0
    mConns.Add conn, CStr(sock)
End Sub

Public Sub NetUnregisterSocket(ByVal sock As Long)
    On Error Resume Next
    mConns.Remove CStr(sock)
End Sub

Private Function FindConn(ByVal sock As Long) As clsConnection
    On Error Resume Next
    Set FindConn = mConns(CStr(sock))
End Function

'=====================================================================
' WndProc: socket olaylarini ilgili baglantiya dagitir.
' lParam'in dusuk kelimesi = olay, yuksek kelimesi = hata kodu.
'=====================================================================
Public Function SocketWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_SOCKET Then
        DispatchSocketEvent wParam, lParam
        SocketWndProc = 0
        Exit Function
    End If
    SocketWndProc = CallWindowProcW(mOldProc, hWnd, uMsg, wParam, lParam)
End Function

Private Sub DispatchSocketEvent(ByVal sock As Long, ByVal lParam As Long)
    On Error GoTo SafeExit
    Dim evt As Long, errCode As Long, c As clsConnection
    evt = lParam And &HFFFF&
    errCode = (lParam \ &H10000) And &HFFFF&

    Set c = FindConn(sock)
    If c Is Nothing Then Exit Sub

    Select Case evt
        Case FD_CONNECT: c.OnConnect errCode
        Case FD_ACCEPT:  c.OnAccept
        Case FD_READ:    c.OnRead
        Case FD_WRITE:   c.OnWrite
        Case FD_CLOSE:   c.OnClose errCode
    End Select
    Exit Sub
SafeExit:
    LogWrite LVL_ERROR, "Socket olay dagitim hatasi: " & Err.Description
End Sub

'=====================================================================
' Paket cerceve olusturma
'=====================================================================
Public Function NetBuildFrame(ByVal pktType As Long, ByVal sessionId As Long, ByVal flags As Long, ByRef payload() As Byte, ByVal payloadLen As Long) As Byte()
    Dim f() As Byte, total As Long
    If payloadLen < 0 Then payloadLen = 0
    total = PKT_HEADER_SIZE + payloadLen
    ReDim f(0 To total - 1)
    f(0) = PKT_MAGIC0
    f(1) = PKT_MAGIC1
    f(2) = PKT_VERSION
    f(3) = CByte(flags And &HFF)
    PutLong f, 4, pktType
    PutLong f, 8, sessionId
    PutLong f, 12, payloadLen
    If payloadLen > 0 Then CopyMemory f(PKT_HEADER_SIZE), payload(LBound(payload)), payloadLen
    NetBuildFrame = f
End Function

'=====================================================================
' Byte yardimcilari (little-endian Long)
'=====================================================================
Public Sub PutLong(ByRef b() As Byte, ByVal offset As Long, ByVal value As Long)
    CopyMemory b(offset), value, 4
End Sub

Public Function GetLong(ByRef b() As Byte, ByVal offset As Long) As Long
    Dim v As Long
    CopyMemory v, b(offset), 4
    GetLong = v
End Function

'=====================================================================
' UTF-8 <-> VB string donusumleri (Unicode guvenli ag iletisimi)
'=====================================================================
Public Function StringToUtf8(ByVal s As String, ByRef outLen As Long) As Byte()
    Dim b() As Byte, wlen As Long, n As Long
    wlen = Len(s)
    If wlen = 0 Then
        outLen = 0
        ReDim b(0 To 0)
        StringToUtf8 = b
        Exit Function
    End If
    n = WideCharToMultiByte(CP_UTF8, 0, StrPtr(s), wlen, 0, 0, 0, 0)
    If n <= 0 Then
        outLen = 0
        ReDim b(0 To 0)
        StringToUtf8 = b
        Exit Function
    End If
    ReDim b(0 To n - 1)
    WideCharToMultiByte CP_UTF8, 0, StrPtr(s), wlen, VarPtr(b(0)), n, 0, 0
    outLen = n
    StringToUtf8 = b
End Function

Public Function Utf8ToString(ByRef b() As Byte, ByVal n As Long) As String
    If n <= 0 Then
        Utf8ToString = ""
        Exit Function
    End If
    Dim wlen As Long, s As String
    wlen = MultiByteToWideChar(CP_UTF8, 0, VarPtr(b(LBound(b))), n, 0, 0)
    If wlen <= 0 Then
        Utf8ToString = ""
        Exit Function
    End If
    s = String$(wlen, vbNullChar)
    MultiByteToWideChar CP_UTF8, 0, VarPtr(b(LBound(b))), n, StrPtr(s), wlen
    Utf8ToString = s
End Function
