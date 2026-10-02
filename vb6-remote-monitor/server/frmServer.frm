VERSION 5.00
Begin VB.Form frmServer
   Caption         =   "Ekran Sunucu (Server) - Ev / Cocuk PC"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLog
      Height          =   1935
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   2640
      Width           =   5535
   End
   Begin VB.CheckBox chkGray
      Caption         =   "Gri tonlama (daha kucuk/hizli)"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.TextBox txtQuality
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4200
      TabIndex        =   7
      Text            =   "35"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtScale
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Text            =   "50"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtFps
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4200
      TabIndex        =   3
      Text            =   "3"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtPort
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "5900"
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdStop
      Caption         =   "Durdur"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2940
      TabIndex        =   9
      Top             =   1740
      Width           =   2715
   End
   Begin VB.CommandButton cmdStart
      Caption         =   "Baslat (Dinlemeye gec)"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1740
      Width           =   2715
   End
   Begin VB.Timer tmrMain
      Enabled         =   0   'False
      Interval        =   333
      Left            =   5280
      Top             =   1740
   End
   Begin VB.Label lblStatus
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Hazir"
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   5535
   End
   Begin VB.Label lblQuality
      Caption         =   "Kalite (1-100):"
      Height          =   255
      Left            =   2940
      TabIndex        =   6
      Top             =   885
      Width           =   1215
   End
   Begin VB.Label lblScale
      Caption         =   "Olcek (%):"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   885
      Width           =   1095
   End
   Begin VB.Label lblFps
      Caption         =   "FPS (1-10):"
      Height          =   255
      Left            =   2940
      TabIndex        =   2
      Top             =   525
      Width           =   1215
   End
   Begin VB.Label lblPort
      Caption         =   "Dinleme portu:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   525
      Width           =   1215
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================
' frmServer  -  Ev / cocuk PC'de calisir. Ekrani yakalayip baglanan
' panele (client) canli JPEG kareleri gonderir. OCX YOK, diske yazma YOK.
'=============================================================================
Option Explicit

Private mListener As Long
Private mClient As Long
Private mRunning As Boolean
Private mFrameCount As Long
Private mFpsCounter As Long
Private mLastSecond As Long

Private Sub Form_Load()
    mListener = INVALID_SOCKET
    mClient = INVALID_SOCKET
    If Not NetStartup() Then
        MsgBox "Winsock baslatilamadi (ws2_32).", vbCritical
    End If
    If Not GfxStartup() Then
        MsgBox "GDI+ baslatilamadi (gdiplus).", vbCritical
    End If
    Log "Hazir. Port girip 'Baslat' deyin."
End Sub

Private Sub cmdStart_Click()
    Dim port As Long
    port = Val(txtPort.Text)
    If port < 1 Or port > 65535 Then
        MsgBox "Gecerli bir port girin (1-65535).", vbExclamation
        Exit Sub
    End If

    mListener = CreateListener(port)
    If mListener = INVALID_SOCKET Then
        MsgBox "Dinleme baslatilamadi. Port kullaniliyor olabilir. Hata: " & NetLastError(), vbCritical
        Exit Sub
    End If

    ApplyInterval
    mRunning = True
    tmrMain.Enabled = True
    cmdStart.Enabled = False
    cmdStop.Enabled = True
    txtPort.Enabled = False
    SetStatus "Dinleniyor (port " & port & ") - panel baglantisi bekleniyor..."
    Log "Dinleme basladi, port " & port & "."
End Sub

Private Sub cmdStop_Click()
    StopAll
End Sub

Private Sub txtFps_Change()
    If mRunning Then ApplyInterval
End Sub

Private Sub ApplyInterval()
    Dim fps As Long
    fps = Val(txtFps.Text)
    If fps < 1 Then fps = 1
    If fps > 10 Then fps = 10
    tmrMain.Interval = 1000 \ fps
End Sub

Private Sub tmrMain_Timer()
    If Not mRunning Then Exit Sub

    If mClient = INVALID_SOCKET Then
        ' Henuz panel baglanmadi -> bekleyen baglantiyi kabul etmeye calis
        Dim s As Long
        s = AcceptClient(mListener)
        If s <> INVALID_SOCKET Then
            mClient = s
            SetNonBlocking mClient, False   ' gonderimi basit tutmak icin blocking
            SetStatus "Panel baglandi. Canli yayin suruyor..."
            Log "Panel baglandi."
        End If
        Exit Sub
    End If

    ' Panel bagli -> kareyi yakala ve gonder
    Dim jpg() As Byte, n As Long
    n = CaptureFrameJpeg(Val(txtQuality.Text), Val(txtScale.Text), (chkGray.Value = 1), jpg)
    If n > 0 Then
        If SendFrame(mClient, jpg, n) Then
            mFrameCount = mFrameCount + 1
            mFpsCounter = mFpsCounter + 1
            UpdateFps
        Else
            Log "Gonderim hatasi - baglanti kapandi. Yeni panel bekleniyor."
            CloseSock mClient
            SetStatus "Panel baglantisi koptu. Yeniden bekleniyor..."
        End If
    End If
End Sub

Private Sub UpdateFps()
    Dim nowSec As Long
    nowSec = Timer
    If nowSec <> mLastSecond Then
        SetStatus "Canli yayin - " & mFpsCounter & " kare/sn - toplam " & mFrameCount & " kare"
        mFpsCounter = 0
        mLastSecond = nowSec
    End If
End Sub

Private Sub StopAll()
    mRunning = False
    tmrMain.Enabled = False
    CloseSock mClient
    CloseSock mListener
    cmdStart.Enabled = True
    cmdStop.Enabled = False
    txtPort.Enabled = True
    SetStatus "Durduruldu."
    Log "Durduruldu."
End Sub

Private Sub SetStatus(ByVal s As String)
    lblStatus.Caption = " " & s
End Sub

Private Sub Log(ByVal s As String)
    txtLog.Text = txtLog.Text & Format$(Now, "hh:nn:ss") & "  " & s & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    StopAll
    GfxShutdown
    NetCleanup
End Sub
