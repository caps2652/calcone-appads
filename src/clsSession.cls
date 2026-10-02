VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================================
' clsSession.cls
' Host/Viewer oturum yoneticisi. Yerel kimlik (ID + erisim sifresi),
' dinleyici (Host) ve aktif baglanti (Viewer/kabul edilen istemci)
' burada yonetilir. UI, bu sinifin olaylarini dinler.
'
' PHASE 1 kapsami: LAN pipe kurulumu + HELLO/WELCOME el sikismasi +
' PING/PONG. Kimlik dogrulama (sifre challenge) ve Allow/Deny izni
' PHASE 2'de bu iskelet uzerine eklenecektir.
'
' GUVENLIK: Erisim sifresi bellekte tutulur, loglanmaz; kalici
' saklama (Phase 9) yalnizca hash ile yapilacaktir.
'=====================================================================
Option Explicit

' UI icin olaylar
Public Event StatusChanged(ByVal text As String, ByVal kind As StatusKind)
Public Event IncomingPeer(ByVal peerIP As String)      ' Phase 2: Allow/Deny burada
Public Event PeerConnected(ByVal peerIP As String)
Public Event PeerDisconnected()
Public Event IdentityChanged()                          ' ID/sifre degisti -> UI tazele

Private WithEvents mListener As clsConnection
Attribute mListener.VB_VarHelpID = -1
Private WithEvents mConn As clsConnection
Attribute mConn.VB_VarHelpID = -1

Private mLocalID As String
Private mLocalPwd As String
Private mLocalPwdHash As String
Private mRequirePwd As Boolean          ' erisim sifresi isteniyor mu?
Private mHostName As String
Private mStatusText As String
Private mStatusKind As StatusKind

'=====================================================================
' Yerel kimlik
'=====================================================================
Public Sub InitLocalIdentity()
    mLocalID = GenerateConnectionID()
    mHostName = LocalHostName()
    mRequirePwd = ConfigGetRequirePassword()

    ' Erisim sifresi yalnizca istenirse uretilir; aksi halde bos birakilir.
    If mRequirePwd Then
        RegeneratePassword
    Else
        mLocalPwd = ""
        mLocalPwdHash = ""
    End If

    SetStatus "Hazir", SK_READY
    LogWrite LVL_INFO, "Yerel kimlik olusturuldu. ID=" & mLocalID & _
             " (sifre " & IIf(mRequirePwd, "acik", "kapali") & ")"   ' sifre loglanmaz
End Sub

Public Sub RegeneratePassword()
    mLocalPwd = GenerateAccessPassword()
    ' Phase 2 challenge-response icin referans hash (duz metin saklanmaz).
    mLocalPwdHash = Sha256HexFromString(mLocalID & ":" & mLocalPwd)
    RaiseEvent IdentityChanged
End Sub

Public Property Get RequirePassword() As Boolean
    RequirePassword = mRequirePwd
End Property

' Erisim sifresini ac/kapat. Acilinca sifre uretilir, kapaninca temizlenir.
Public Sub SetRequirePassword(ByVal b As Boolean)
    If mRequirePwd = b Then Exit Sub
    mRequirePwd = b
    ConfigSetRequirePassword b
    If b Then
        RegeneratePassword          ' olusturur ve IdentityChanged yayinlar
    Else
        mLocalPwd = ""
        mLocalPwdHash = ""
        RaiseEvent IdentityChanged
    End If
    LogWrite LVL_INFO, "Erisim sifresi " & IIf(b, "acildi", "kapatildi")
End Sub

Public Property Get LocalID() As String
    LocalID = mLocalID
End Property
Public Property Get LocalIDFormatted() As String
    LocalIDFormatted = FormatConnectionID(mLocalID)
End Property
Public Property Get LocalPassword() As String
    LocalPassword = mLocalPwd         ' Kopyala butonu icin; UI'da maskeli gosterilir
End Property
Public Property Get StatusText() As String
    StatusText = mStatusText
End Property
Public Property Get StatusKindValue() As StatusKind
    StatusKindValue = mStatusKind
End Property
Public Property Get IsHosting() As Boolean
    IsHosting = (Not mListener Is Nothing)
End Property
Public Property Get HasPeer() As Boolean
    HasPeer = (Not mConn Is Nothing)
End Property

'=====================================================================
' Host modu: dinlemeye basla
'=====================================================================
Public Sub StartHosting()
    If Not mListener Is Nothing Then Exit Sub
    Set mListener = New clsConnection
    If mListener.Listen(ConfigGetPort()) Then
        SetStatus "Hazir (gelen baglantilar dinleniyor)", SK_READY
    Else
        Set mListener = Nothing
        SetStatus "Dinleme baslatilamadi (port " & ConfigGetPort() & " kullanimda olabilir)", SK_ERROR
    End If
End Sub

Public Sub StopHosting()
    If mListener Is Nothing Then Exit Sub
    mListener.CloseNow
    Set mListener = Nothing
End Sub

'=====================================================================
' Viewer modu: partner'a baglan.
' target: "192.168.1.50" veya "192.168.1.50:5000"
' (9 haneli ID ile LAN kesfi sonraki fazda; simdilik IP:Port.)
'=====================================================================
Public Sub ConnectToPartner(ByVal target As String)
    Dim ip As String, port As Long
    If Not ParseTarget(target, ip, port) Then
        SetStatus "Gecersiz hedef. IP veya IP:Port girin (orn. 192.168.1.50:7070).", SK_ERROR
        Exit Sub
    End If

    ' Onceki baglanti varsa temizle.
    If Not mConn Is Nothing Then
        mConn.CloseNow
        Set mConn = Nothing
    End If

    Set mConn = New clsConnection
    SetStatus "Baglaniliyor... (" & ip & ":" & port & ")", SK_BUSY
    If Not mConn.ConnectTo(ip, port) Then
        Set mConn = Nothing
        ' Hata mesaji SocketErr olayindan gelir; burada ek durum birakmayalim.
    End If
End Sub

Public Sub Disconnect()
    If mConn Is Nothing Then Exit Sub
    On Error Resume Next
    mConn.SendEmpty PT_DISCONNECT
    mConn.CloseNow
    Set mConn = Nothing
    SetStatus "Baglanti sonlandirildi", SK_READY
    RaiseEvent PeerDisconnected
End Sub

Public Sub ShutdownAll()
    On Error Resume Next
    If Not mConn Is Nothing Then
        mConn.CloseNow
        Set mConn = Nothing
    End If
    If Not mListener Is Nothing Then
        mListener.CloseNow
        Set mListener = Nothing
    End If
End Sub

'=====================================================================
' Dinleyici olaylari
'=====================================================================
Private Sub mListener_ClientAccepted(ByVal client As clsConnection)
    ' PHASE 1: pipe'i kabul et ve HELLO bekle.
    ' PHASE 2: burada once Allow/Deny ve sifre dogrulamasi yapilacak.
    If Not mConn Is Nothing Then
        ' Zaten aktif bir es var; yeni geleni nazikce kapat (tek oturum).
        client.CloseNow
        Exit Sub
    End If
    Set mConn = client
    RaiseEvent IncomingPeer(client.PeerIP)
    SetStatus "Gelen baglanti: " & client.PeerIP, SK_BUSY
End Sub

'=====================================================================
' Aktif baglanti olaylari
'=====================================================================
Private Sub mConn_Connected()
    ' Viewer tarafi: baglanti kuruldu, kendimizi tanitalim.
    SetStatus "Kimlik dogrulaniyor...", SK_BUSY
    SendHello
End Sub

Private Sub mConn_Disconnected()
    Set mConn = Nothing
    SetStatus IIf(IsHosting, "Hazir (gelen baglantilar dinleniyor)", "Baglanti kesildi"), _
              IIf(IsHosting, SK_READY, SK_ERROR)
    RaiseEvent PeerDisconnected
End Sub

Private Sub mConn_SocketErr(ByVal code As Long, ByVal desc As String)
    SetStatus desc, SK_ERROR
    LogWrite LVL_WARNING, "Socket hatasi: " & desc
End Sub

Private Sub mConn_PacketReceived(ByVal pkt As clsPacket)
    Select Case pkt.PType
        Case PT_HELLO
            ' Host tarafi: karsi taraf tanitti -> WELCOME ile yanitla.
            LogWrite LVL_DEBUG, "HELLO alindi: " & pkt.PayloadString()
            SendWelcome
            CompletePeerConnected

        Case PT_WELCOME
            ' Viewer tarafi: el sikismasi tamam.
            LogWrite LVL_DEBUG, "WELCOME alindi: " & pkt.PayloadString()
            CompletePeerConnected

        Case PT_PING
            If Not mConn Is Nothing Then mConn.SendEmpty PT_PONG

        Case PT_PONG
            ' keep-alive yaniti; su an ek islem yok

        Case PT_DISCONNECT
            Disconnect

        Case Else
            ' Diger paket turleri sonraki fazlarda islenecek.
            LogWrite LVL_DEBUG, "Islenmeyen paket turu: " & pkt.PType
    End Select
End Sub

'=====================================================================
' El sikismasi yardimcilari
'=====================================================================
Private Sub SendHello()
    If mConn Is Nothing Then Exit Sub
    mConn.SendString PT_HELLO, "ver=" & PROTOCOL_VERSION & ";id=" & mLocalID & ";name=" & mHostName
End Sub

Private Sub SendWelcome()
    If mConn Is Nothing Then Exit Sub
    mConn.SendString PT_WELCOME, "ver=" & PROTOCOL_VERSION & ";id=" & mLocalID & ";name=" & mHostName
End Sub

Private Sub CompletePeerConnected()
    SetStatus "Baglandi: " & mConn.PeerIP, SK_READY
    RaiseEvent PeerConnected(mConn.PeerIP)
End Sub

'=====================================================================
' Yardimcilar
'=====================================================================
Private Sub SetStatus(ByVal text As String, ByVal kind As StatusKind)
    mStatusText = text
    mStatusKind = kind
    RaiseEvent StatusChanged(text, kind)
End Sub

'--- "ip" veya "ip:port" ayristir ---
Private Function ParseTarget(ByVal s As String, ByRef ip As String, ByRef port As Long) As Boolean
    s = Trim$(s)
    If Len(s) = 0 Then ParseTarget = False: Exit Function

    Dim p As Long
    p = InStr(s, ":")
    If p > 0 Then
        ip = Left$(s, p - 1)
        port = CLng(Val(Mid$(s, p + 1)))
    Else
        ip = s
        port = ConfigGetPort()
    End If

    ' Kaba IPv4 dogrulamasi: dort oktet, her biri 0..255.
    If Not IsValidIPv4(ip) Then ParseTarget = False: Exit Function
    If port < 1 Or port > 65535 Then ParseTarget = False: Exit Function
    ParseTarget = True
End Function

Private Function IsValidIPv4(ByVal s As String) As Boolean
    Dim parts() As String, i As Long, v As Long
    parts = Split(s, ".")
    If UBound(parts) <> 3 Then Exit Function
    For i = 0 To 3
        If Len(parts(i)) = 0 Or Not IsNumeric(parts(i)) Then Exit Function
        v = CLng(Val(parts(i)))
        If v < 0 Or v > 255 Then Exit Function
    Next i
    IsValidIPv4 = True
End Function

Private Function LocalHostName() As String
    Dim buf As String, n As Long
    buf = String$(256, vbNullChar)
    If api_gethostname(buf, 256) = 0 Then
        n = InStr(buf, vbNullChar)
        If n > 1 Then LocalHostName = Left$(buf, n - 1) Else LocalHostName = "PC"
    Else
        LocalHostName = "PC"
    End If
End Function
