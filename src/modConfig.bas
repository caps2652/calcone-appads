Attribute VB_Name = "modConfig"
'=====================================================================
' modConfig.bas
' Ayarlarin yuklenmesi/kaydedilmesi (config.ini).
' Unicode guvenli olmasi icin W (wide) INI API'leri kullanilir.
' GUVENLIK: Sifreler burada DUZ METIN olarak saklanmaz. Sadece
' sifre HASH'i (Phase 2'de kullanilacak) ve sifresiz ayarlar tutulur.
'=====================================================================
Option Explicit

Private Const CFG_FILE As String = "config.ini"
Private Const SEC_GENERAL As String = "General"
Private Const SEC_APPEARANCE As String = "Appearance"
Private Const SEC_CONNECTION As String = "Connection"

' Tema modlari
Public Enum ThemeMode
    TM_DARK = 0
    TM_LIGHT = 1
    TM_SYSTEM = 2
End Enum

Private mFile As String
Private mThemeMode As ThemeMode
Private mPort As Long
Private mTimeoutMs As Long

Public Sub ConfigLoad()
    mFile = AppPath() & CFG_FILE

    mThemeMode = IniGetLong(SEC_APPEARANCE, "Theme", TM_DARK)
    If mThemeMode < TM_DARK Or mThemeMode > TM_SYSTEM Then mThemeMode = TM_DARK

    mPort = IniGetLong(SEC_CONNECTION, "Port", DEFAULT_PORT)
    If mPort < 1 Or mPort > 65535 Then mPort = DEFAULT_PORT

    mTimeoutMs = IniGetLong(SEC_CONNECTION, "TimeoutMs", DEFAULT_TIMEOUT_MS)
    If mTimeoutMs < 1000 Then mTimeoutMs = DEFAULT_TIMEOUT_MS

    LogSetDebug (IniGetLong(SEC_GENERAL, "Debug", 0) <> 0)
End Sub

'--- Erisimciler (accessor) ---
Public Function ConfigGetThemeMode() As ThemeMode
    ConfigGetThemeMode = mThemeMode
End Function

Public Sub ConfigSetThemeMode(ByVal m As ThemeMode)
    mThemeMode = m
    IniPutLong SEC_APPEARANCE, "Theme", m
End Sub

Public Function ConfigGetPort() As Long
    ConfigGetPort = mPort
End Function

Public Sub ConfigSetPort(ByVal p As Long)
    If p >= 1 And p <= 65535 Then
        mPort = p
        IniPutLong SEC_CONNECTION, "Port", p
    End If
End Sub

Public Function ConfigGetTimeoutMs() As Long
    ConfigGetTimeoutMs = mTimeoutMs
End Function

'=====================================================================
' INI yardimcilari (Unicode / W API)
'=====================================================================
Private Function IniGetLong(ByVal section As String, ByVal key As String, ByVal defVal As Long) As Long
    IniGetLong = CLng(Val(IniGetString(section, key, CStr(defVal))))
End Function

Private Sub IniPutLong(ByVal section As String, ByVal key As String, ByVal value As Long)
    IniPutString section, key, CStr(value)
End Sub

Private Function IniGetString(ByVal section As String, ByVal key As String, ByVal defVal As String) As String
    Dim buf As String, n As Long
    buf = String$(512, vbNullChar)
    n = GetPrivateProfileStringW(StrPtr(section), StrPtr(key), StrPtr(defVal), StrPtr(buf), 512, StrPtr(mFile))
    If n > 0 Then
        IniGetString = Left$(buf, n)
    Else
        IniGetString = defVal
    End If
End Function

Private Sub IniPutString(ByVal section As String, ByVal key As String, ByVal value As String)
    WritePrivateProfileStringW StrPtr(section), StrPtr(key), StrPtr(value), StrPtr(mFile)
End Sub
