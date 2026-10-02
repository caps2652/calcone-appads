VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================================
' clsTheme.cls
' Windows 11 tarzi dark/light renk paleti. Renkler COLORREF (BGR Long)
' olarak RGB() ile doner; hem VB hem GDI API ayni degeri bekler.
'=====================================================================
Option Explicit

' Sistem temasini okumak icin registry (sadece bu sinifa ozel).
Private Declare Function RegOpenKeyExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const KEY_READ As Long = &H20019

Private mDark As Boolean

Public Property Get IsDark() As Boolean
    IsDark = mDark
End Property

'--- Istenen moda gore paleti ayarla ---
Public Sub ApplyMode(ByVal mode As ThemeMode)
    Select Case mode
        Case TM_LIGHT: mDark = False
        Case TM_DARK:  mDark = True
        Case Else:     mDark = SystemUsesDark()   ' TM_SYSTEM
    End Select
End Sub

'=====================================================================
' Renk paleti (dark varsayilan). Her renk tek bir yerden gelir ki
' tutarli kalsin ve ileride kolay degistirilsin.
'=====================================================================
Public Property Get WindowBg() As Long
    WindowBg = IIf(mDark, RGB(32, 32, 32), RGB(243, 243, 243))
End Property

Public Property Get TitleBarBg() As Long
    TitleBarBg = IIf(mDark, RGB(28, 28, 28), RGB(238, 238, 238))
End Property

Public Property Get CardBg() As Long
    CardBg = IIf(mDark, RGB(43, 43, 43), RGB(255, 255, 255))
End Property

Public Property Get FieldBg() As Long
    FieldBg = IIf(mDark, RGB(55, 55, 55), RGB(248, 248, 248))
End Property

Public Property Get Separator() As Long
    Separator = IIf(mDark, RGB(58, 58, 58), RGB(224, 224, 224))
End Property

Public Property Get TextPrimary() As Long
    TextPrimary = IIf(mDark, RGB(240, 240, 240), RGB(26, 26, 26))
End Property

Public Property Get TextSecondary() As Long
    TextSecondary = IIf(mDark, RGB(165, 165, 165), RGB(100, 100, 100))
End Property

Public Property Get Accent() As Long
    Accent = IIf(mDark, RGB(76, 194, 255), RGB(0, 103, 192))
End Property

Public Property Get AccentHover() As Long
    AccentHover = IIf(mDark, RGB(102, 205, 255), RGB(16, 118, 208))
End Property

Public Property Get AccentPressed() As Long
    AccentPressed = IIf(mDark, RGB(60, 170, 230), RGB(0, 90, 170))
End Property

Public Property Get AccentText() As Long
    ' Accent dolgulu buton uzerindeki yazi rengi
    AccentText = IIf(mDark, RGB(0, 0, 0), RGB(255, 255, 255))
End Property

Public Property Get BtnBg() As Long
    BtnBg = IIf(mDark, RGB(55, 55, 55), RGB(255, 255, 255))
End Property

Public Property Get BtnHover() As Long
    BtnHover = IIf(mDark, RGB(64, 64, 64), RGB(245, 245, 245))
End Property

Public Property Get BtnPressed() As Long
    BtnPressed = IIf(mDark, RGB(48, 48, 48), RGB(235, 235, 235))
End Property

Public Property Get BtnBorder() As Long
    BtnBorder = IIf(mDark, RGB(70, 70, 70), RGB(220, 220, 220))
End Property

Public Property Get StatusReady() As Long
    StatusReady = IIf(mDark, RGB(72, 201, 134), RGB(16, 160, 90))
End Property

Public Property Get StatusBusy() As Long
    StatusBusy = IIf(mDark, RGB(240, 180, 60), RGB(200, 140, 20))
End Property

Public Property Get StatusError() As Long
    StatusError = IIf(mDark, RGB(235, 95, 95), RGB(200, 50, 50))
End Property

'--- Sistem acik tema mi kullaniyor? (TM_SYSTEM icin) ---
Private Function SystemUsesDark() As Boolean
    On Error GoTo Fallback
    Dim hKey As Long, subKey As String
    subKey = "Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    If RegOpenKeyExW(HKEY_CURRENT_USER, StrPtr(subKey), 0, KEY_READ, hKey) = 0 Then
        Dim valName As String, data As Long, cb As Long, t As Long
        valName = "AppsUseLightTheme"
        cb = 4
        If RegQueryValueExW(hKey, StrPtr(valName), 0, t, data, cb) = 0 Then
            RegCloseKey hKey
            SystemUsesDark = (data = 0)   ' 0 = koyu, 1 = acik
            Exit Function
        End If
        RegCloseKey hKey
    End If
Fallback:
    SystemUsesDark = True   ' varsayilan koyu
End Function
