VERSION 5.00
Begin VB.Form frmClient
   Caption         =   "Yonetici Panel (Client) - Is Yeri"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picEkran
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5055
      Left            =   120
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   544
      TabIndex        =   7
      Top             =   1320
      Width           =   8055
   End
   Begin VB.CommandButton cmdDisconnect
      Caption         =   "Baglantiyi Kes"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdConnect
      Caption         =   "Baglan"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtPort
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3180
      TabIndex        =   3
      Text            =   "5900"
      Top             =   525
      Width           =   855
   End
   Begin VB.TextBox txtHost
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "192.168.1.10"
      Top             =   525
      Width           =   1935
   End
   Begin VB.Timer tmrPoll
      Enabled         =   0   'False
      Interval        =   40
      Left            =   7680
      Top             =   840
   End
   Begin VB.Label lblStatus
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Hazir"
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   8055
   End
   Begin VB.Label lblPortC
      Caption         =   "Port:"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   570
      Width           =   495
   End
   Begin VB.Label lblHostC
      Caption         =   "Sunucu:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   570
      Width           =   735
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================
' frmClient  -  Is yerinde calisir. Ev PC'sindeki sunucuya baglanir,
' gelen canli JPEG karelerini cozup ekranda gosterir. OCX YOK.
'=============================================================================
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal numBytes As Long)

Private Const ST_DISCONNECTED As Long = 0
Private Const ST_CONNECTING As Long = 1
Private Const ST_CONNECTED As Long = 2

Private Const CHUNK As Long = 16384
Private Const MAX_FRAME As Long = 4000000   ' guvenlik: makul ust sinir

Private mSock As Long
Private mState As Long
Private mBuf() As Byte        ' gelen veri tamponu
Private mBufLen As Long       ' tampondaki gecerli byte sayisi
Private mBufCap As Long       ' tampon kapasitesi
Private mConnectTick As Long  ' baglanma zaman asimi icin
Private mFpsCounter As Long
Private mLastSecond As Long

Private Sub Form_Load()
    mSock = INVALID_SOCKET
    mState = ST_DISCONNECTED
    mBufCap = CHUNK
    ReDim mBuf(0 To mBufCap - 1)
    mBufLen = 0
    If Not NetStartup() Then MsgBox "Winsock baslatilamadi (ws2_32).", vbCritical
    If Not GfxStartup() Then MsgBox "GDI+ baslatilamadi (gdiplus).", vbCritical
    SetStatus "Hazir. Sunucu IP ve port girip 'Baglan' deyin."
End Sub

Private Sub cmdConnect_Click()
    Dim host As String, port As Long
    host = Trim$(txtHost.Text)
    port = Val(txtPort.Text)
    If Len(host) = 0 Then MsgBox "Sunucu adresi girin.", vbExclamation: Exit Sub
    If port < 1 Or port > 65535 Then MsgBox "Gecerli port girin.", vbExclamation: Exit Sub

    mSock = ConnectTo(host, port)
    If mSock = INVALID_SOCKET Then
        MsgBox "Soket olusturulamadi veya adres cozulemedi.", vbCritical
        Exit Sub
    End If

    mState = ST_CONNECTING
    mConnectTick = Timer
    mBufLen = 0
    tmrPoll.Enabled = True
    cmdConnect.Enabled = False
    cmdDisconnect.Enabled = True
    txtHost.Enabled = False
    txtPort.Enabled = False
    SetStatus "Baglaniliyor " & host & ":" & port & " ..."
End Sub

Private Sub cmdDisconnect_Click()
    Disconnect "Kullanici baglantiyi kesti."
End Sub

Private Sub tmrPoll_Timer()
    Select Case mState
        Case ST_CONNECTING
            If SockWritable(mSock) Then
                mState = ST_CONNECTED
                SetStatus "Baglandi. Canli yayin bekleniyor..."
            ElseIf Timer - mConnectTick > 10 Then
                Disconnect "Baglanti zaman asimina ugradi."
            End If

        Case ST_CONNECTED
            DrainSocket
    End Select
End Sub

'--- Sokette biriken tum veriyi oku ve kareleri isle -----------------------
Private Sub DrainSocket()
    Dim tmp() As Byte, n As Long, guard As Long
    ReDim tmp(0 To CHUNK - 1)

    Do
        n = RecvSome(mSock, tmp, CHUNK)
        If n = -1 Then
            Exit Do                    ' su an baska veri yok
        ElseIf n = 0 Then
            Disconnect "Sunucu baglantiyi kapatti."
            Exit Sub
        Else
            AppendToBuffer tmp, n
        End If
        guard = guard + 1
        If guard > 256 Then Exit Do    ' bir turda cok fazla takilma
    Loop

    ParseFrames
End Sub

'--- Gelen veriyi tampona ekle ---------------------------------------------
Private Sub AppendToBuffer(ByRef data() As Byte, ByVal count As Long)
    If mBufLen + count > mBufCap Then
        Do
            mBufCap = mBufCap * 2
        Loop While mBufLen + count > mBufCap
        ReDim Preserve mBuf(0 To mBufCap - 1)
    End If
    CopyMemory mBuf(mBufLen), data(0), count
    mBufLen = mBufLen + count
End Sub

'--- Tampondan tam kareleri ayikla, goster ---------------------------------
Private Sub ParseFrames()
    Dim frameLen As Long, total As Long

    Do While mBufLen >= 4
        If mBuf(3) >= 128 Then
            Disconnect "Bozuk veri (gecersiz kare boyu)."
            Exit Sub
        End If
        frameLen = CLng(mBuf(0)) + CLng(mBuf(1)) * 256 + CLng(mBuf(2)) * 65536 + CLng(mBuf(3)) * 16777216

        If frameLen <= 0 Or frameLen > MAX_FRAME Then
            Disconnect "Bozuk veri (kare boyu sinir disi)."
            Exit Sub
        End If

        total = 4 + frameLen
        If mBufLen < total Then Exit Do    ' kare henuz tam gelmedi

        ' Kareyi ayir
        Dim frame() As Byte
        ReDim frame(0 To frameLen - 1)
        CopyMemory frame(0), mBuf(4), frameLen

        DisplayFrame frame, frameLen

        ' Isaretlenen kareyi tampondan cikar (kalan veriyi one kaydir)
        Dim remaining As Long
        remaining = mBufLen - total
        If remaining > 0 Then CopyMemory mBuf(0), mBuf(total), remaining
        mBufLen = remaining
    Loop
End Sub

'--- Bir JPEG kareyi coz ve PictureBox'a ciz -------------------------------
Private Sub DisplayFrame(ByRef frame() As Byte, ByVal length As Long)
    Dim hbm As Long, w As Long, h As Long
    hbm = DecodeJpegToHBitmap(frame, length, w, h)
    If hbm <> 0 Then
        DrawHBitmapStretched picEkran.hdc, hbm, picEkran.ScaleWidth, picEkran.ScaleHeight, w, h
        picEkran.Refresh
        DeleteBitmap hbm
        mFpsCounter = mFpsCounter + 1
        UpdateFps
    End If
End Sub

Private Sub UpdateFps()
    Dim nowSec As Long
    nowSec = Timer
    If nowSec <> mLastSecond Then
        SetStatus "Canli yayin - " & mFpsCounter & " kare/sn"
        mFpsCounter = 0
        mLastSecond = nowSec
    End If
End Sub

Private Sub Disconnect(ByVal reason As String)
    tmrPoll.Enabled = False
    CloseSock mSock
    mState = ST_DISCONNECTED
    mBufLen = 0
    cmdConnect.Enabled = True
    cmdDisconnect.Enabled = False
    txtHost.Enabled = True
    txtPort.Enabled = True
    SetStatus reason
End Sub

Private Sub SetStatus(ByVal s As String)
    lblStatus.Caption = " " & s
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picEkran.Move picEkran.Left, picEkran.Top, _
                  Me.ScaleWidth - picEkran.Left * 2, _
                  Me.ScaleHeight - picEkran.Top - 120
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Disconnect "Kapatiliyor."
    GfxShutdown
    NetCleanup
End Sub
