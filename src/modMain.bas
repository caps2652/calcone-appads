Attribute VB_Name = "modMain"
'=====================================================================
' modMain.bas
' Uygulama giris noktasi (Sub Main), uygulama genelindeki sabitler
' ve tekil (singleton) nesneler. Global degisken sayisi minimumda
' tutulur; sadece uygulama omru boyunca yasayan cekirdek nesneler.
'=====================================================================
Option Explicit

'--- Uygulama kimligi / surum ---
Public Const APP_TITLE As String = "Remote Support"
Public Const APP_INTERNAL As String = "RemoteSupport"
Public Const APP_VERSION As String = "0.1.0"        ' PHASE 1
Public Const PROTOCOL_VERSION As Long = 1

'--- Ag varsayilanlari ---
Public Const DEFAULT_PORT As Long = 7070            ' ayarlardan degistirilebilir
Public Const DEFAULT_TIMEOUT_MS As Long = 15000

'--- Durum gostergesi turu (UI renk/ikon secimi icin) ---
Public Enum StatusKind
    SK_READY = 0
    SK_BUSY = 1
    SK_ERROR = 2
End Enum

'--- Uygulama genelindeki tekil nesneler ---
Public gTheme As clsTheme                           ' aktif renk paleti
Public gSession As clsSession                       ' host/viewer oturum yoneticisi

' Uygulama yolu (sondaki ters-bolu ile) - log/config icin
Public Function AppPath() As String
    If Right$(App.Path, 1) = "\" Then
        AppPath = App.Path
    Else
        AppPath = App.Path & "\"
    End If
End Function

'=====================================================================
' Giris noktasi. Proje "Sub Main" ile baslatilir.
'=====================================================================
Public Sub Main()
    On Error GoTo Fatal

    ' DPI: islemi DPI-aware yap ki Windows gorseli bulaniklastirmasin.
    ' (Vista+ fonksiyonu; basarisiz olursa sorun degil.)
    SetProcessDPIAware

    ' Alt sistemleri sirayla baslat.
    LogInit
    LogWrite LVL_INFO, "=== " & APP_TITLE & " v" & APP_VERSION & " baslatiliyor ==="

    ConfigLoad
    UIInit                      ' DPI olcek faktoru + font onbellegi

    Set gTheme = New clsTheme
    gTheme.ApplyMode ConfigGetThemeMode()

    If Not NetInit() Then
        ' Winsock baslatilamadiysa kullanici dostu mesajla cik.
        MsgBox "Ag alt sistemi baslatilamadi (Winsock)." & vbCrLf & _
               "Uygulama kapatilacak.", vbCritical, APP_TITLE
        GoTo CleanExit
    End If

    Set gSession = New clsSession
    gSession.InitLocalIdentity        ' baglanti ID'si + erisim sifresi uretir

    ' Ana pencereyi goster.
    Load frmMain
    frmMain.Show

    Exit Sub

Fatal:
    LogWrite LVL_ERROR, "Main() beklenmeyen hata: " & Err.Number & " - " & Err.Description
    MsgBox "Uygulama baslatilirken bir sorun olustu." & vbCrLf & _
           "Ayrinti icin Logs\app.log dosyasina bakin.", vbCritical, APP_TITLE
CleanExit:
    ShutdownApp
End Sub

'=====================================================================
' Duzenli kapanis: tum kaynaklari serbest birak (socket/thread/GDI).
' Her cikis yolundan tek bir yerden gecilir.
'=====================================================================
Public Sub ShutdownApp()
    On Error Resume Next
    If Not gSession Is Nothing Then
        gSession.ShutdownAll
        Set gSession = Nothing
    End If
    NetShutdown
    UICleanup
    LogWrite LVL_INFO, "=== Uygulama kapandi ==="
End Sub
