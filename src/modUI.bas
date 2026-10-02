Attribute VB_Name = "modUI"
'=====================================================================
' modUI.bas
' GDI tabanli modern cizim yardimcilari + DPI olcekleme + font onbellegi.
' Hicbir OCX/ActiveX kullanilmaz; tum gorsel ogeler Windows API ile cizilir.
' GDI nesneleri onbelleklenir ve UICleanup ile serbest birakilir (leak yok).
'=====================================================================
Option Explicit

Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Const NULL_BRUSH As Long = 5
Private Const NULL_PEN As Long = 8

Private mDpi As Long          ' yatay DPI (96=100%, 120=125%, 144=150%, 192=200%)
Private mScale As Single      ' DPI olcek faktoru
Private mFonts As Collection  ' HFONT onbellegi ("pt|weight" -> handle)
Private mInited As Boolean

'--- Alt sistemi hazirla: sistem DPI'sini oku, font onbellegini kur ---
Public Sub UIInit()
    Dim hdc As Long
    hdc = GetDC(0)
    mDpi = GetDeviceCaps(hdc, LOGPIXELSX)
    ReleaseDC 0, hdc
    If mDpi < 96 Then mDpi = 96
    mScale = mDpi / 96#
    Set mFonts = New Collection
    mInited = True
End Sub

Public Sub UICleanup()
    On Error Resume Next
    If Not mFonts Is Nothing Then
        Dim v As Variant
        For Each v In mFonts
            DeleteObject CLng(v)
        Next v
        Set mFonts = Nothing
    End If
    mInited = False
End Sub

'--- DPI olcekleme: tasarim pikselini gercek piksele cevir ---
Public Function Dp(ByVal px As Long) As Long
    If Not mInited Then Dp = px Else Dp = CLng(px * mScale)
End Function

Public Function ScaleFactor() As Single
    ScaleFactor = mScale
End Function

'=====================================================================
' Font onbellegi: Segoe UI, verilen punto ve agirlik icin HFONT doner.
' Punto -> piksel yukseklik DPI'ye gore hesaplanir.
'=====================================================================
Public Function UIFont(ByVal pt As Single, Optional ByVal weight As Long = FW_NORMAL) As Long
    Dim key As String
    key = CStr(pt) & "|" & CStr(weight)

    On Error Resume Next
    Dim h As Long
    h = mFonts(key)
    On Error GoTo 0
    If h <> 0 Then
        UIFont = h
        Exit Function
    End If

    Dim px As Long
    px = -CLng(pt * mDpi / 72#)       ' negatif = karakter yuksekligi

    Dim face As String
    face = "Segoe UI"
    h = CreateFontW(px, 0, 0, 0, weight, 0, 0, 0, _
                    DEFAULT_CHARSET, 0, 0, CLEARTYPE_QUALITY, _
                    DEFAULT_PITCH Or FF_DONTCARE, StrPtr(face))
    If h <> 0 Then mFonts.Add h, key
    UIFont = h
End Function

'=====================================================================
' Cizim yardimcilari
'=====================================================================

'--- Duz dikdortgen dolgu ---
Public Sub FillRectC(ByVal hdc As Long, ByVal l As Long, ByVal t As Long, ByVal r As Long, ByVal b As Long, ByVal color As Long)
    Dim rc As RECT, hBr As Long
    rc.Left = l: rc.Top = t: rc.Right = r: rc.Bottom = b
    hBr = CreateSolidBrush(color)
    FillRect hdc, rc, hBr
    DeleteObject hBr
End Sub

'--- Yuvarlatilmis dolgulu dikdortgen (kenarliksiz) ---
Public Sub FillRoundRectC(ByVal hdc As Long, ByVal l As Long, ByVal t As Long, ByVal r As Long, ByVal b As Long, ByVal radius As Long, ByVal color As Long)
    Dim hBr As Long, hPen As Long, oBr As Long, oPen As Long
    hBr = CreateSolidBrush(color)
    hPen = CreatePen(PS_SOLID, 1, color)
    oBr = SelectObject(hdc, hBr)
    oPen = SelectObject(hdc, hPen)
    RoundRect hdc, l, t, r, b, radius, radius
    SelectObject hdc, oBr
    SelectObject hdc, oPen
    DeleteObject hBr
    DeleteObject hPen
End Sub

'--- Yuvarlatilmis dikdortgen kenarlik (dolgusuz) ---
Public Sub StrokeRoundRectC(ByVal hdc As Long, ByVal l As Long, ByVal t As Long, ByVal r As Long, ByVal b As Long, ByVal radius As Long, ByVal color As Long, Optional ByVal penWidth As Long = 1)
    Dim hPen As Long, oPen As Long, oBr As Long
    hPen = CreatePen(PS_SOLID, penWidth, color)
    oPen = SelectObject(hdc, hPen)
    oBr = SelectObject(hdc, GetStockObject(NULL_BRUSH))
    RoundRect hdc, l, t, r, b, radius, radius
    SelectObject hdc, oPen
    SelectObject hdc, oBr
    DeleteObject hPen
End Sub

'--- Dolu daire (durum noktasi / ikon icin) ---
Public Sub FillCircleC(ByVal hdc As Long, ByVal cx As Long, ByVal cy As Long, ByVal radius As Long, ByVal color As Long)
    Dim hBr As Long, hPen As Long, oBr As Long, oPen As Long
    hBr = CreateSolidBrush(color)
    hPen = CreatePen(PS_SOLID, 1, color)
    oBr = SelectObject(hdc, hBr)
    oPen = SelectObject(hdc, hPen)
    Ellipse hdc, cx - radius, cy - radius, cx + radius, cy + radius
    SelectObject hdc, oBr
    SelectObject hdc, oPen
    DeleteObject hBr
    DeleteObject hPen
End Sub

'--- Ince cizgi (separator) ---
Public Sub DrawLineC(ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal color As Long, Optional ByVal penWidth As Long = 1)
    Dim hPen As Long, oPen As Long
    hPen = CreatePen(PS_SOLID, penWidth, color)
    oPen = SelectObject(hdc, hPen)
    MoveToEx hdc, x1, y1, ByVal 0&
    LineTo hdc, x2, y2
    SelectObject hdc, oPen
    DeleteObject hPen
End Sub

'--- Unicode metin ciz (DrawTextW). Dikdortgen icine hizalanir. ---
' Not: flags varsayilani = DT_LEFT Or DT_SINGLELINE Or DT_VCENTER Or DT_NOPREFIX
' (VB6 Optional varsayilani sabit olmali, bu yuzden onceden hesaplanmis literal.)
Public Sub DrawTextC(ByVal hdc As Long, ByVal text As String, ByVal l As Long, ByVal t As Long, ByVal r As Long, ByVal b As Long, ByVal color As Long, ByVal hFont As Long, Optional ByVal flags As Long = &H824&)
    Dim rc As RECT, oFont As Long
    rc.Left = l: rc.Top = t: rc.Right = r: rc.Bottom = b
    oFont = SelectObject(hdc, hFont)
    SetBkMode hdc, TRANSPARENT
    SetTextColor hdc, color
    DrawTextW hdc, StrPtr(text), -1, rc, flags
    SelectObject hdc, oFont
End Sub

'--- Metnin piksel genisligini olc ---
Public Function MeasureTextWidth(ByVal hdc As Long, ByVal text As String, ByVal hFont As Long) As Long
    Dim rc As RECT, oFont As Long
    rc.Right = 10000
    oFont = SelectObject(hdc, hFont)
    DrawTextW hdc, StrPtr(text), -1, rc, DT_CALCRECT Or DT_SINGLELINE Or DT_NOPREFIX
    SelectObject hdc, oFont
    MeasureTextWidth = rc.Right - rc.Left
End Function

'--- Nokta bir dikdortgenin icinde mi? (hit-test) ---
Public Function PtInRect(ByVal px As Long, ByVal py As Long, ByRef rc As RECT) As Boolean
    PtInRect = (px >= rc.Left And px < rc.Right And py >= rc.Top And py < rc.Bottom)
End Function

'--- GetDC/ReleaseDC yerel bildirimleri (sadece DPI olcumu icin) ---
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
