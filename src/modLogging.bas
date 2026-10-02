Attribute VB_Name = "modLogging"
'=====================================================================
' modLogging.bas
' Basit, guvenli log sistemi -> Logs\app.log
' GUVENLIK: Sifre, kimlik dogrulama sirri, pano icerigi ve dosya
' icerikleri ASLA loglanmaz. Cagiran taraf bunlari gondermemelidir;
' ayrica burada da acik bir "maskeleme" notu birakilmistir.
'=====================================================================
Option Explicit

Public Enum LogLevel
    LVL_DEBUG = 0
    LVL_INFO = 1
    LVL_WARNING = 2
    LVL_ERROR = 3
End Enum

Private Const LOG_DIR As String = "Logs"
Private Const LOG_FILE As String = "app.log"
Private Const MAX_LOG_BYTES As Long = 1048576      ' 1 MB -> uzerine donulur

Private mPath As String
Private mDebugEnabled As Boolean
Private mReady As Boolean

'--- Log alt sistemini hazirla (Logs klasorunu olustur) ---
Public Sub LogInit()
    On Error Resume Next
    Dim d As String
    d = AppPath() & LOG_DIR
    If Dir$(d, vbDirectory) = "" Then MkDir d
    mPath = d & "\" & LOG_FILE
    mDebugEnabled = False
    mReady = True
End Sub

'--- Debug seviyesini ac/kapat (Settings ekranindan kontrol edilir) ---
Public Sub LogSetDebug(ByVal enabled As Boolean)
    mDebugEnabled = enabled
End Sub

Public Function LogIsDebug() As Boolean
    LogIsDebug = mDebugEnabled
End Function

'--- Ana log yazma fonksiyonu ---
Public Sub LogWrite(ByVal level As LogLevel, ByVal message As String)
    On Error Resume Next
    If Not mReady Then Exit Sub
    If level = LVL_DEBUG And Not mDebugEnabled Then Exit Sub

    RollIfTooBig

    Dim ff As Integer
    ff = FreeFile
    Open mPath For Append As #ff
    Print #ff, FormatDate(Now) & " [" & LevelName(level) & "] " & message
    Close #ff
End Sub

'--- Log dosyasi buyuduyse sade bir "rollover" uygula ---
Private Sub RollIfTooBig()
    On Error Resume Next
    If Dir$(mPath) = "" Then Exit Sub
    If FileLen(mPath) < MAX_LOG_BYTES Then Exit Sub

    Dim bak As String
    bak = Left$(mPath, Len(mPath) - 4) & ".1.log"
    If Dir$(bak) <> "" Then Kill bak
    Name mPath As bak
End Sub

Private Function LevelName(ByVal level As LogLevel) As String
    Select Case level
        Case LVL_DEBUG:   LevelName = "DEBUG"
        Case LVL_INFO:    LevelName = "INFO"
        Case LVL_WARNING: LevelName = "WARN"
        Case LVL_ERROR:   LevelName = "ERROR"
        Case Else:        LevelName = "?"
    End Select
End Function

Private Function FormatDate(ByVal d As Date) As String
    FormatDate = Format$(d, "yyyy-mm-dd hh:nn:ss")
End Function
