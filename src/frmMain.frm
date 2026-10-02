VERSION 5.00
Begin VB.Form frmMain
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Remote Support"
   ClientHeight    =   8400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   560
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   880
   ShowInTaskbar   =   -1  'True
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrToast
      Enabled         =   0   'False
      Interval        =   0
      Left            =   120
      Top             =   120
   End
   Begin VB.Timer tmrConn
      Enabled         =   0   'False
      Interval        =   0
      Left            =   600
      Top             =   120
   End
   Begin VB.TextBox txtPartnerID
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   0
      Top             =   2400
      Width           =   3000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=====================================================================
' frmMain.frm
' Modern (Windows 11 tarzi) ana pencere. TAMAMEN GDI ile cizilir;
' klasik gri form/CommandButton/CheckBox gorunumu kullanilmaz.
' Butonlar owner-drawn'dir ve hit-test ile yonetilir.
'=====================================================================
Option Explicit

'--- Buton kimlikleri ---
Private Enum BtnId
    BID_NONE = -1
    BID_MIN = 1
    BID_CLOSE = 2
    BID_NEWPWD = 3
    BID_COPYPWD = 4
    BID_COPYID = 5
    BID_REVEAL = 6
    BID_SETTINGS = 7
    BID_MANAGE = 8
    BID_CONNECT = 9
    BID_THEME = 10
    BID_TOGGLE_PWD = 11
End Enum

'--- Buton stilleri ---
Private Enum BtnStyle
    BS_NORMAL = 0       ' kenarlikli yumusak buton
    BS_ACCENT = 1       ' vurgu (accent) dolgulu buton
    BS_TITLE = 2        ' baslik cubugu ikon butonu
    BS_TITLECLOSE = 3   ' baslik kapat (hover kirmizi)
    BS_GHOST = 4        ' arka plansiz hafif buton
    BS_TOGGLE = 5       ' modern ac/kapa anahtari (label + switch)
End Enum

'--- Basit vektorel ikon (glyph) turleri ---
Private Enum Glyph
    GL_NONE = 0
    GL_LOGO = 1
    GL_REFRESH = 2
    GL_COPY = 3
    GL_ID = 4
    GL_GEAR = 5
    GL_LIST = 6
    GL_ARROW = 7
    GL_EYE = 8
    GL_SUN = 9
    GL_MOON = 10
    GL_MIN = 11
    GL_CLOSE = 12
End Enum

Private Type UiButton
    id As BtnId
    style As BtnStyle
    gl As Glyph
    caption As String
    rc As RECT
    enabled As Boolean
    onState As Boolean      ' BS_TOGGLE icin acik/kapali durumu
End Type

Private Const MAX_BTNS As Long = 32
Private mBtn(0 To MAX_BTNS - 1) As UiButton
Private mBtnCount As Long

Private mHover As BtnId
Private mPressed As BtnId
Private mRevealPwd As Boolean

' Partner alani (textbox'in arkasina cizilecek) dikdortgeni
Private mFieldRc As RECT

' Baglanti zaman asimi takibi
Private mConnecting As Boolean
Private mConnectStart As Long

' Toast bildirim durumu
Private mToastText As String
Private mToastKind As StatusKind
Private mToastUntil As Long

' Oturum olaylarini dinle
Private WithEvents mSession As clsSession
Attribute mSession.VB_VarHelpID = -1

'=====================================================================
' Form yasam dongusu
'=====================================================================
Private Sub Form_Load()
    On Error GoTo Fail

    ' Pencereyi DPI'ye gore boyutlandir ve ortala.
    Dim w As Long, h As Long
    w = Dp(880): h = Dp(560)
    Me.Move (Screen.Width - w * Screen.TwipsPerPixelX) \ 2, _
            (Screen.Height - h * Screen.TwipsPerPixelY) \ 2, _
            w * Screen.TwipsPerPixelX, h * Screen.TwipsPerPixelY

    ApplyThemeToControls
    Set mSession = gSession

    ' Host modunu baslat (gelen baglantilari dinle). Kontrol devri
    ' PHASE 2'de kullanici onayina (Allow/Deny) baglanacaktir.
    gSession.StartHosting

    mHover = BID_NONE
    mPressed = BID_NONE

    LayoutAndPaint
    ApplyRoundedCorners
    Exit Sub
Fail:
    LogWrite LVL_ERROR, "Form_Load hata: " & Err.Description
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    LayoutAndPaint
    ApplyRoundedCorners
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mSession = Nothing
    ShutdownApp
End Sub

'=====================================================================
' Yerlesim + cizim
'=====================================================================
Private Sub LayoutAndPaint()
    BuildLayout
    PaintAll
End Sub

Private Sub BuildLayout()
    Dim W As Long, H As Long
    W = Me.ScaleWidth: H = Me.ScaleHeight
    mBtnCount = 0

    Dim titleH As Long, statusH As Long, pad As Long, gap As Long
    titleH = Dp(44): statusH = Dp(40): pad = Dp(20): gap = Dp(16)

    '--- Baslik cubugu butonlari (sagda) ---
    Dim tb As Long
    tb = Dp(44)
    AddBtn BID_CLOSE, BS_TITLECLOSE, GL_CLOSE, "", W - tb, 0, W, titleH
    AddBtn BID_MIN, BS_TITLE, GL_MIN, "", W - tb * 2, 0, W - tb, titleH

    '--- Icerik kartlari ---
    Dim cTop As Long, cBot As Long, cardW As Long
    cTop = titleH + pad
    cBot = H - statusH - pad
    cardW = (W - pad * 2 - gap) \ 2

    Dim lRc As RECT, rRc As RECT
    lRc.Left = pad: lRc.Top = cTop: lRc.Right = pad + cardW: lRc.Bottom = cBot
    rRc.Left = pad + cardW + gap: rRc.Top = cTop: rRc.Right = W - pad: rRc.Bottom = cBot

    LayoutLeftCard lRc
    LayoutRightCard rRc

    '--- Durum cubugu tema butonu (sag alt) ---
    Dim thW As Long, thH As Long
    thW = Dp(96): thH = Dp(26)
    AddBtn BID_THEME, BS_GHOST, IIf(gTheme.IsDark, GL_SUN, GL_MOON), _
           IIf(gTheme.IsDark, "Acik tema", "Koyu tema"), _
           W - pad - thW, H - statusH + (statusH - thH) \ 2, W - pad, H - statusH + (statusH - thH) \ 2 + thH
End Sub

Private Sub LayoutLeftCard(ByRef rc As RECT)
    Dim ip As Long, bx As Long, bw As Long, bh As Long, y As Long
    Dim halfGap As Long, halfW As Long
    ip = Dp(22)                          ' ic bosluk
    bx = rc.Left + ip
    bw = (rc.Right - rc.Left) - ip * 2
    bh = Dp(38)
    halfGap = Dp(10)
    halfW = (bw - halfGap) \ 2

    ' Erisim sifresi ac/kapa anahtari (her zaman gorunur, varsayilan KAPALI).
    AddBtn BID_TOGGLE_PWD, BS_TOGGLE, GL_NONE, "Erisim sifresi iste", _
           bx, rc.Top + Dp(162), rc.Right - ip, rc.Top + Dp(196), _
           gSession.RequirePassword

    If gSession.RequirePassword Then
        ' Sifre acik: goz (reveal) + iki satir buton (sifre islemleri dahil).
        Dim pwRowY As Long
        pwRowY = rc.Top + Dp(228)
        AddBtn BID_REVEAL, BS_GHOST, GL_EYE, "", rc.Right - ip - Dp(34), pwRowY, rc.Right - ip, pwRowY + Dp(28)

        y = rc.Bottom - ip - bh * 2 - halfGap
        AddBtn BID_NEWPWD, BS_NORMAL, GL_REFRESH, "Yeni Sifre", bx, y, bx + halfW, y + bh
        AddBtn BID_COPYPWD, BS_NORMAL, GL_COPY, "Sifreyi Kopyala", bx + halfW + halfGap, y, bx + bw, y + bh
        y = y + bh + halfGap
        AddBtn BID_COPYID, BS_NORMAL, GL_ID, "ID'yi Kopyala", bx, y, bx + halfW, y + bh
        AddBtn BID_SETTINGS, BS_NORMAL, GL_GEAR, "Ayarlar", bx + halfW + halfGap, y, bx + bw, y + bh
    Else
        ' Sifre kapali: tek satir buton (ID kopyala + ayarlar).
        y = rc.Bottom - ip - bh
        AddBtn BID_COPYID, BS_NORMAL, GL_ID, "ID'yi Kopyala", bx, y, bx + halfW, y + bh
        AddBtn BID_SETTINGS, BS_NORMAL, GL_GEAR, "Ayarlar", bx + halfW + halfGap, y, bx + bw, y + bh
    End If
End Sub

Private Sub LayoutRightCard(ByRef rc As RECT)
    Dim ip As Long, bx As Long, bw As Long
    ip = Dp(22)
    bx = rc.Left + ip
    bw = (rc.Right - rc.Left) - ip * 2

    ' Partner alani (textbox'in arkasindaki cerceve)
    mFieldRc.Left = bx
    mFieldRc.Top = rc.Top + Dp(96)
    mFieldRc.Right = bx + bw
    mFieldRc.Bottom = mFieldRc.Top + Dp(44)

    ' Gercek textbox'i cercevenin icine yerlestir (ince ic boslukla).
    ' Form ScaleMode=Pixel oldugundan tum olculer pikseldir.
    Dim tin As Long, txtH As Long
    tin = Dp(12)
    txtH = Dp(24)
    txtPartnerID.Move mFieldRc.Left + tin, _
                      mFieldRc.Top + ((mFieldRc.Bottom - mFieldRc.Top) - txtH) \ 2, _
                      (mFieldRc.Right - mFieldRc.Left) - tin * 2, _
                      txtH

    ' BAGLAN butonu
    Dim byy As Long, bh As Long
    bh = Dp(44)
    byy = mFieldRc.Bottom + Dp(16)
    AddBtn BID_CONNECT, BS_ACCENT, GL_ARROW, "BAGLAN", bx, byy, bx + bw, byy + bh

    ' Baglantilari Yonet (tam genislik, normal) - alt kisim
    Dim my As Long
    my = rc.Bottom - ip - Dp(38)
    AddBtn BID_MANAGE, BS_NORMAL, GL_LIST, "Baglantilari Yonet", bx, my, bx + bw, my + Dp(38)
End Sub

Private Sub AddBtn(ByVal id As BtnId, ByVal style As BtnStyle, ByVal gl As Glyph, ByVal caption As String, ByVal l As Long, ByVal t As Long, ByVal r As Long, ByVal b As Long, Optional ByVal onState As Boolean = False)
    If mBtnCount >= MAX_BTNS Then Exit Sub
    With mBtn(mBtnCount)
        .id = id: .style = style: .gl = gl: .caption = caption
        .rc.Left = l: .rc.Top = t: .rc.Right = r: .rc.Bottom = b
        .enabled = True
        .onState = onState
    End With
    mBtnCount = mBtnCount + 1
End Sub

'=====================================================================
' Ana cizim
'=====================================================================
Private Sub PaintAll()
    On Error Resume Next
    Dim hdc As Long, W As Long, H As Long
    hdc = Me.hdc
    W = Me.ScaleWidth: H = Me.ScaleHeight

    ' Arka plan
    FillRectC hdc, 0, 0, W, H, gTheme.WindowBg

    DrawTitleBar hdc, W
    DrawLeftCard hdc
    DrawRightCard hdc
    DrawStatusBar hdc, W, H

    ' Butonlar
    Dim i As Long
    For i = 0 To mBtnCount - 1
        DrawButton hdc, mBtn(i)
    Next i

    If Len(mToastText) > 0 Then DrawToast hdc, W, H

    Me.Refresh
End Sub

Private Sub DrawTitleBar(ByVal hdc As Long, ByVal W As Long)
    Dim titleH As Long
    titleH = Dp(44)
    FillRectC hdc, 0, 0, W, titleH, gTheme.TitleBarBg

    ' Logo glyph + uygulama adi
    DrawGlyph hdc, GL_LOGO, Dp(16), (titleH - Dp(18)) \ 2, Dp(18), gTheme.Accent
    DrawTextC hdc, APP_TITLE, Dp(44), 0, Dp(300), titleH, gTheme.TextPrimary, _
              UIFont(10.5, FW_SEMIBOLD)

    DrawLineC hdc, 0, titleH, W, titleH, gTheme.Separator
End Sub

Private Sub DrawLeftCard(ByVal hdc As Long)
    Dim rc As RECT
    rc = CardRect(True)
    DrawCardBase hdc, rc, "Bu Bilgisayar"

    Dim ip As Long, x As Long, w As Long
    ip = Dp(22): x = rc.Left + ip: w = (rc.Right - rc.Left) - ip * 2

    ' Connection ID
    DrawTextC hdc, "Baglanti ID", x, rc.Top + Dp(50), rc.Right - ip, rc.Top + Dp(70), _
              gTheme.TextSecondary, UIFont(9)
    DrawTextC hdc, gSession.LocalIDFormatted, x, rc.Top + Dp(70), rc.Right - ip, rc.Top + Dp(108), _
              gTheme.TextPrimary, UIFont(21, FW_SEMIBOLD)

    ' Durum (dot + metin)
    Dim dotY As Long
    dotY = rc.Top + Dp(132)
    FillCircleC hdc, x + Dp(5), dotY, Dp(5), StatusColor(gSession.StatusKindValue)
    DrawTextC hdc, gSession.StatusText, x + Dp(18), dotY - Dp(12), rc.Right - ip, dotY + Dp(12), _
              gTheme.TextPrimary, UIFont(9.5)

    ' (Erisim sifresi ac/kapa anahtari BS_TOGGLE butonu olarak cizilir.)

    If gSession.RequirePassword Then
        ' Sifre acik: etiket + maskeli/acik deger.
        DrawTextC hdc, "Erisim Sifresi", x, rc.Top + Dp(204), rc.Right - ip, rc.Top + Dp(224), _
                  gTheme.TextSecondary, UIFont(9)
        Dim pwText As String
        If mRevealPwd Then
            pwText = gSession.LocalPassword
        Else
            pwText = String$(Len(gSession.LocalPassword), ChrW$(&H2022))   ' bullet
        End If
        DrawTextC hdc, pwText, x, rc.Top + Dp(224), rc.Right - ip - Dp(40), rc.Top + Dp(256), _
                  gTheme.TextPrimary, UIFont(15, FW_SEMIBOLD)
        DrawLineC hdc, x, rc.Bottom - Dp(120), rc.Right - ip, rc.Bottom - Dp(120), gTheme.Separator
    Else
        ' Sifre kapali: katilimli erisim aciklamasi.
        DrawTextC hdc, "Sifre kapali. Baglantilar yalnizca sizin onayinizla" & vbCrLf & _
                  "(Allow / Deny) kurulur. Gozetimsiz erisim icin sifreyi acin.", _
                  x, rc.Top + Dp(204), rc.Right - ip, rc.Top + Dp(252), _
                  gTheme.TextSecondary, UIFont(8.5), DT_LEFT Or DT_WORDBREAK Or DT_NOPREFIX
        DrawLineC hdc, x, rc.Bottom - Dp(72), rc.Right - ip, rc.Bottom - Dp(72), gTheme.Separator
    End If
End Sub

Private Sub DrawRightCard(ByVal hdc As Long)
    Dim rc As RECT
    rc = CardRect(False)
    DrawCardBase hdc, rc, "Uzak Bilgisayara Baglan"

    Dim ip As Long, x As Long
    ip = Dp(22): x = rc.Left + ip

    DrawTextC hdc, "Partner ID / IP", x, rc.Top + Dp(66), rc.Right - ip, rc.Top + Dp(86), _
              gTheme.TextSecondary, UIFont(9)

    ' Partner alani cercevesi (textbox'in arkasi)
    FillRoundRectC hdc, mFieldRc.Left, mFieldRc.Top, mFieldRc.Right, mFieldRc.Bottom, Dp(8), gTheme.FieldBg
    StrokeRoundRectC hdc, mFieldRc.Left, mFieldRc.Top, mFieldRc.Right, mFieldRc.Bottom, Dp(8), gTheme.BtnBorder

    ' Yardim metni
    DrawTextC hdc, "Uzak bilgisayara baglanmak icin karsi bilgisayarin" & vbCrLf & _
              "IP adresini girin (orn. 192.168.1.50:7070).", _
              x, mFieldRc.Bottom + Dp(70), rc.Right - ip, mFieldRc.Bottom + Dp(118), _
              gTheme.TextSecondary, UIFont(8.5), DT_LEFT Or DT_WORDBREAK Or DT_NOPREFIX
End Sub

Private Sub DrawStatusBar(ByVal hdc As Long, ByVal W As Long, ByVal H As Long)
    Dim sTop As Long
    sTop = H - Dp(40)
    DrawLineC hdc, 0, sTop, W, sTop, gTheme.Separator
    FillCircleC hdc, Dp(24), sTop + Dp(20), Dp(4), StatusColor(gSession.StatusKindValue)
    DrawTextC hdc, gSession.StatusText, Dp(38), sTop, Dp(560), H, gTheme.TextSecondary, UIFont(8.5)
End Sub

Private Sub DrawCardBase(ByVal hdc As Long, ByRef rc As RECT, ByVal title As String)
    FillRoundRectC hdc, rc.Left, rc.Top, rc.Right, rc.Bottom, Dp(12), gTheme.CardBg
    StrokeRoundRectC hdc, rc.Left, rc.Top, rc.Right, rc.Bottom, Dp(12), gTheme.Separator
    DrawTextC hdc, title, rc.Left + Dp(22), rc.Top + Dp(16), rc.Right - Dp(22), rc.Top + Dp(42), _
              gTheme.TextPrimary, UIFont(12, FW_SEMIBOLD)
End Sub

Private Function CardRect(ByVal isLeft As Boolean) As RECT
    Dim W As Long, H As Long, titleH As Long, statusH As Long, pad As Long, gap As Long, cardW As Long
    W = Me.ScaleWidth: H = Me.ScaleHeight
    titleH = Dp(44): statusH = Dp(40): pad = Dp(20): gap = Dp(16)
    cardW = (W - pad * 2 - gap) \ 2
    Dim r As RECT
    r.Top = titleH + pad
    r.Bottom = H - statusH - pad
    If isLeft Then
        r.Left = pad: r.Right = pad + cardW
    Else
        r.Left = pad + cardW + gap: r.Right = W - pad
    End If
    CardRect = r
End Function

'=====================================================================
' Buton cizimi
'=====================================================================
Private Sub DrawButton(ByVal hdc As Long, ByRef btn As UiButton)
    Dim bg As Long, fg As Long, hovered As Boolean, pressedNow As Boolean
    hovered = (mHover = btn.id)
    pressedNow = (mPressed = btn.id And hovered)

    ' Toggle anahtari kendi ciziminde yonetilir.
    If btn.style = BS_TOGGLE Then DrawToggleButton hdc, btn: Exit Sub

    Dim l As Long, t As Long, r As Long, b As Long
    l = btn.rc.Left: t = btn.rc.Top: r = btn.rc.Right: b = btn.rc.Bottom

    Select Case btn.style
        Case BS_ACCENT
            bg = gTheme.Accent
            If pressedNow Then bg = gTheme.AccentPressed ElseIf hovered Then bg = gTheme.AccentHover
            fg = gTheme.AccentText
            FillRoundRectC hdc, l, t, r, b, Dp(8), bg

        Case BS_NORMAL
            bg = gTheme.BtnBg
            If pressedNow Then bg = gTheme.BtnPressed ElseIf hovered Then bg = gTheme.BtnHover
            fg = gTheme.TextPrimary
            FillRoundRectC hdc, l, t, r, b, Dp(8), bg
            StrokeRoundRectC hdc, l, t, r, b, Dp(8), gTheme.BtnBorder

        Case BS_GHOST
            fg = gTheme.TextSecondary
            If hovered Then
                FillRoundRectC hdc, l, t, r, b, Dp(6), gTheme.BtnHover
                fg = gTheme.TextPrimary
            End If

        Case BS_TITLE
            fg = gTheme.TextSecondary
            If hovered Then
                FillRectC hdc, l, t, r, b, gTheme.BtnHover
                fg = gTheme.TextPrimary
            End If

        Case BS_TITLECLOSE
            fg = gTheme.TextSecondary
            If hovered Then
                FillRectC hdc, l, t, r, b, RGB(232, 17, 35)   ' Windows kapat kirmizisi
                fg = RGB(255, 255, 255)
            End If
    End Select

    ' Ikon + metin yerlesimi
    Dim cx As Long, cy As Long, gsz As Long
    cy = (t + b) \ 2
    gsz = Dp(16)

    If btn.style = BS_TITLE Or btn.style = BS_TITLECLOSE Then
        DrawGlyph hdc, btn.gl, (l + r) \ 2 - gsz \ 2, cy - gsz \ 2, gsz, fg
        Exit Sub
    End If

    If Len(btn.caption) = 0 Then
        ' Yalnizca ikon (orn. reveal)
        DrawGlyph hdc, btn.gl, (l + r) \ 2 - gsz \ 2, cy - gsz \ 2, gsz, fg
        Exit Sub
    End If

    ' Ikon + metni birlikte ortala
    Dim hFont As Long, tw As Long, totalW As Long, startX As Long
    hFont = UIFont(IIf(btn.style = BS_ACCENT, 10.5, 9.5), IIf(btn.style = BS_ACCENT, FW_SEMIBOLD, FW_NORMAL))
    tw = MeasureTextWidth(hdc, btn.caption, hFont)
    Dim hasGlyph As Boolean
    hasGlyph = (btn.gl <> GL_NONE)
    totalW = tw + IIf(hasGlyph, gsz + Dp(8), 0)
    startX = (l + r) \ 2 - totalW \ 2

    If hasGlyph Then
        DrawGlyph hdc, btn.gl, startX, cy - gsz \ 2, gsz, fg
        DrawTextC hdc, btn.caption, startX + gsz + Dp(8), t, r, b, fg, hFont, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER Or DT_NOPREFIX
    Else
        DrawTextC hdc, btn.caption, l, t, r, b, fg, hFont, DT_CENTER Or DT_SINGLELINE Or DT_VCENTER Or DT_NOPREFIX
    End If
End Sub

'--- Modern ac/kapa anahtari (label solda, switch sagda) ---
Private Sub DrawToggleButton(ByVal hdc As Long, ByRef btn As UiButton)
    Dim l As Long, t As Long, r As Long, b As Long
    l = btn.rc.Left: t = btn.rc.Top: r = btn.rc.Right: b = btn.rc.Bottom

    If mHover = btn.id Then FillRoundRectC hdc, l, t, r, b, Dp(8), gTheme.BtnHover

    ' Etiket
    DrawTextC hdc, btn.caption, l + Dp(4), t, r - Dp(56), b, gTheme.TextPrimary, UIFont(9.5)

    ' Anahtar govdesi
    Dim sw As Long, sh As Long, sx As Long, sy As Long, cy As Long
    sw = Dp(44): sh = Dp(24)
    sx = r - sw - Dp(4)
    sy = (t + b) \ 2 - sh \ 2
    cy = sy + sh \ 2

    Dim trackCol As Long
    If btn.onState Then
        trackCol = gTheme.Accent
    Else
        trackCol = IIf(gTheme.IsDark, RGB(80, 80, 80), RGB(200, 200, 200))
    End If
    FillRoundRectC hdc, sx, sy, sx + sw, sy + sh, sh \ 2, trackCol

    ' Topuz (knob)
    Dim kr As Long, kcx As Long
    kr = sh \ 2 - Dp(3)
    If btn.onState Then kcx = sx + sw - sh \ 2 Else kcx = sx + sh \ 2
    FillCircleC hdc, kcx, cy, kr, RGB(255, 255, 255)
End Sub

'=====================================================================
' Basit vektorel ikonlar (line-art). Dis OCX gerektirmez.
'=====================================================================
Private Sub DrawGlyph(ByVal hdc As Long, ByVal gl As Glyph, ByVal x As Long, ByVal y As Long, ByVal sz As Long, ByVal color As Long)
    Dim cx As Long, cy As Long, pw As Long
    cx = x + sz \ 2: cy = y + sz \ 2
    pw = Dp(2): If pw < 1 Then pw = 1

    Select Case gl
        Case GL_LOGO
            ' Yuvarlatilmis kare + nokta (uzak baglanti simgesi)
            StrokeRoundRectC hdc, x, y, x + sz, y + sz, Dp(4), color, pw
            FillCircleC hdc, cx, cy, Dp(2), color

        Case GL_CLOSE
            DrawLineC hdc, x + Dp(3), y + Dp(3), x + sz - Dp(3), y + sz - Dp(3), color, pw
            DrawLineC hdc, x + sz - Dp(3), y + Dp(3), x + Dp(3), y + sz - Dp(3), color, pw

        Case GL_MIN
            DrawLineC hdc, x + Dp(3), cy, x + sz - Dp(3), cy, color, pw

        Case GL_REFRESH
            ' Yaklasik dairesel ok
            StrokeRoundRectC hdc, x + Dp(2), y + Dp(2), x + sz - Dp(2), y + sz - Dp(2), Dp(6), color, pw
            DrawLineC hdc, x + sz - Dp(4), y + Dp(2), x + sz - Dp(2), y + Dp(5), color, pw

        Case GL_COPY
            StrokeRoundRectC hdc, x + Dp(2), y + Dp(4), x + sz - Dp(5), y + sz - Dp(1), Dp(2), color, pw
            StrokeRoundRectC hdc, x + Dp(5), y + Dp(1), x + sz - Dp(1), y + sz - Dp(4), Dp(2), color, pw

        Case GL_ID
            StrokeRoundRectC hdc, x + Dp(1), y + Dp(3), x + sz - Dp(1), y + sz - Dp(3), Dp(3), color, pw
            DrawLineC hdc, x + Dp(4), cy, x + sz - Dp(4), cy, color, pw

        Case GL_GEAR
            StrokeRoundRectC hdc, x + Dp(3), y + Dp(3), x + sz - Dp(3), y + sz - Dp(3), Dp(6), color, pw
            FillCircleC hdc, cx, cy, Dp(2), color

        Case GL_LIST
            DrawLineC hdc, x + Dp(2), y + Dp(4), x + sz - Dp(2), y + Dp(4), color, pw
            DrawLineC hdc, x + Dp(2), cy, x + sz - Dp(2), cy, color, pw
            DrawLineC hdc, x + Dp(2), y + sz - Dp(4), x + sz - Dp(2), y + sz - Dp(4), color, pw

        Case GL_ARROW
            DrawLineC hdc, x + Dp(2), cy, x + sz - Dp(2), cy, color, pw
            DrawLineC hdc, x + sz - Dp(6), y + Dp(3), x + sz - Dp(2), cy, color, pw
            DrawLineC hdc, x + sz - Dp(6), y + sz - Dp(3), x + sz - Dp(2), cy, color, pw

        Case GL_EYE
            StrokeRoundRectC hdc, x + Dp(1), y + Dp(4), x + sz - Dp(1), y + sz - Dp(4), Dp(6), color, pw
            FillCircleC hdc, cx, cy, Dp(2), color

        Case GL_SUN
            FillCircleC hdc, cx, cy, Dp(3), color
            DrawLineC hdc, cx, y, cx, y + Dp(3), color, pw
            DrawLineC hdc, cx, y + sz - Dp(3), cx, y + sz, color, pw
            DrawLineC hdc, x, cy, x + Dp(3), cy, color, pw
            DrawLineC hdc, x + sz - Dp(3), cy, x + sz, cy, color, pw

        Case GL_MOON
            FillCircleC hdc, cx, cy, Dp(6), color
            FillCircleC hdc, cx + Dp(3), cy - Dp(2), Dp(5), gTheme.WindowBg
    End Select
End Sub

Private Sub DrawToast(ByVal hdc As Long, ByVal W As Long, ByVal H As Long)
    Dim hFont As Long, tw As Long, boxW As Long, boxH As Long, l As Long, t As Long
    hFont = UIFont(9.5, FW_SEMIBOLD)
    tw = MeasureTextWidth(hdc, mToastText, hFont)
    boxW = tw + Dp(56): boxH = Dp(40)
    l = (W - boxW) \ 2
    t = H - Dp(40) - boxH - Dp(14)

    FillRoundRectC hdc, l, t, l + boxW, t + boxH, Dp(10), gTheme.CardBg
    StrokeRoundRectC hdc, l, t, l + boxW, t + boxH, Dp(10), gTheme.Separator
    FillCircleC hdc, l + Dp(20), t + boxH \ 2, Dp(5), StatusColor(mToastKind)
    DrawTextC hdc, mToastText, l + Dp(36), t, l + boxW - Dp(12), t + boxH, gTheme.TextPrimary, hFont
End Sub

Private Function StatusColor(ByVal k As StatusKind) As Long
    Select Case k
        Case SK_READY: StatusColor = gTheme.StatusReady
        Case SK_BUSY:  StatusColor = gTheme.StatusBusy
        Case Else:     StatusColor = gTheme.StatusError
    End Select
End Function

'=====================================================================
' Fare etkilesimi (hit-test + surukleme)
'=====================================================================
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim id As BtnId
    id = HitTest(CLng(X), CLng(Y))
    If id <> mHover Then
        mHover = id
        SetCursor LoadCursorW(0, IIf(id <> BID_NONE, IDC_HAND, IDC_ARROW))
        PaintAll
    ElseIf id <> BID_NONE Then
        SetCursor LoadCursorW(0, IDC_HAND)
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    Dim id As BtnId
    id = HitTest(CLng(X), CLng(Y))
    If id <> BID_NONE Then
        mPressed = id
        PaintAll
        Exit Sub
    End If

    ' Baslik cubugundaysa pencereyi surukle.
    If Y < Dp(44) Then
        ReleaseCapture
        SendMessageW Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    Dim id As BtnId
    id = HitTest(CLng(X), CLng(Y))
    Dim act As BtnId
    act = mPressed
    mPressed = BID_NONE
    PaintAll
    If act <> BID_NONE And act = id Then DoAction act
End Sub

Private Function HitTest(ByVal X As Long, ByVal Y As Long) As BtnId
    Dim i As Long
    For i = 0 To mBtnCount - 1
        If mBtn(i).enabled Then
            If PtInRect(X, Y, mBtn(i).rc) Then
                HitTest = mBtn(i).id
                Exit Function
            End If
        End If
    Next i
    HitTest = BID_NONE
End Function

'=====================================================================
' Buton aksiyonlari
'=====================================================================
Private Sub DoAction(ByVal id As BtnId)
    On Error GoTo Fail
    Select Case id
        Case BID_CLOSE
            Unload Me

        Case BID_MIN
            Me.WindowState = vbMinimized

        Case BID_NEWPWD
            gSession.RegeneratePassword
            ShowToast "Yeni sifre olusturuldu", SK_READY

        Case BID_COPYPWD
            Clipboard.Clear
            Clipboard.SetText gSession.LocalPassword
            ShowToast "Sifre panoya kopyalandi", SK_READY

        Case BID_COPYID
            Clipboard.Clear
            Clipboard.SetText gSession.LocalID
            ShowToast "ID panoya kopyalandi", SK_READY

        Case BID_REVEAL
            mRevealPwd = Not mRevealPwd
            PaintAll

        Case BID_SETTINGS
            ShowToast "Ayarlar penceresi PHASE 9'da gelecek", SK_BUSY

        Case BID_MANAGE
            ShowToast "Baglanti gecmisi PHASE 9'da gelecek", SK_BUSY

        Case BID_CONNECT
            StartConnect

        Case BID_THEME
            ToggleTheme

        Case BID_TOGGLE_PWD
            gSession.SetRequirePassword Not gSession.RequirePassword
            If gSession.RequirePassword Then
                ShowToast "Erisim sifresi acildi", SK_READY
            Else
                ShowToast "Erisim sifresi kapatildi", SK_BUSY
            End If
    End Select
    Exit Sub
Fail:
    LogWrite LVL_ERROR, "DoAction(" & id & ") hata: " & Err.Description
    ShowToast "Islem gerceklestirilemedi", SK_ERROR
End Sub

Private Sub StartConnect()
    Dim target As String
    target = Trim$(txtPartnerID.Text)
    If Len(target) = 0 Then
        ShowToast "Lutfen bir IP adresi girin", SK_ERROR
        Exit Sub
    End If
    mConnecting = True
    mConnectStart = GetTickCount()
    tmrConn.Interval = 1000
    tmrConn.Enabled = True
    gSession.ConnectToPartner target
End Sub

Private Sub ToggleTheme()
    Dim newMode As ThemeMode
    If gTheme.IsDark Then newMode = TM_LIGHT Else newMode = TM_DARK
    gTheme.ApplyMode newMode
    ConfigSetThemeMode newMode
    ApplyThemeToControls
    LayoutAndPaint
    ApplyRoundedCorners
End Sub

'=====================================================================
' Oturum olaylari
'=====================================================================
Private Sub mSession_StatusChanged(ByVal text As String, ByVal kind As StatusKind)
    ' Baglanti denemesi sirasinda hata gelirse zaman asimini beklemeden dur.
    If mConnecting And kind = SK_ERROR Then
        mConnecting = False
        tmrConn.Enabled = False
    End If
    PaintAll
End Sub

Private Sub mSession_IdentityChanged()
    ' Sifre acilip kapaninca sol kart duzeni (butonlar) degisir -> yeniden kur.
    LayoutAndPaint
End Sub

Private Sub mSession_IncomingPeer(ByVal peerIP As String)
    ' PHASE 2: burada Allow/Deny onay penceresi acilacak.
    ShowToast "Gelen baglanti: " & peerIP, SK_BUSY
End Sub

Private Sub mSession_PeerConnected(ByVal peerIP As String)
    mConnecting = False
    tmrConn.Enabled = False
    ShowToast ChrW$(&H2713) & " Baglandi: " & peerIP, SK_READY   ' check mark
End Sub

Private Sub mSession_PeerDisconnected()
    mConnecting = False
    tmrConn.Enabled = False
    ShowToast "Baglanti kesildi", SK_ERROR
End Sub

'=====================================================================
' Zamanlayicilar
'=====================================================================
Private Sub tmrConn_Timer()
    ' Baglanti zaman asimi denetimi.
    If Not mConnecting Then
        tmrConn.Enabled = False
        Exit Sub
    End If
    If GetTickCount() - mConnectStart > ConfigGetTimeoutMs() Then
        mConnecting = False
        tmrConn.Enabled = False
        gSession.Disconnect
        ShowToast "Baglanti zaman asimina ugradi", SK_ERROR
    End If
End Sub

Private Sub tmrToast_Timer()
    If GetTickCount() > mToastUntil Then
        mToastText = ""
        tmrToast.Enabled = False
        PaintAll
    End If
End Sub

'=====================================================================
' Yardimcilar
'=====================================================================
Private Sub ShowToast(ByVal text As String, Optional ByVal kind As StatusKind = SK_READY)
    mToastText = text
    mToastKind = kind
    mToastUntil = GetTickCount() + 2600
    tmrToast.Interval = 200
    tmrToast.Enabled = True
    PaintAll
End Sub

Private Sub ApplyThemeToControls()
    txtPartnerID.BackColor = gTheme.FieldBg
    txtPartnerID.ForeColor = gTheme.TextPrimary
End Sub

'--- Pencere koselerini yuvarla (Win11 DWM + tum surumler icin region) ---
Private Sub ApplyRoundedCorners()
    On Error Resume Next
    Dim radius As Long
    radius = Dp(10)

    ' Win11: DWM ile yumusak kose (varsa).
    Dim pref As Long
    pref = DWMWCP_ROUND
    DwmSetWindowAttribute Me.hwnd, DWMWA_WINDOW_CORNER_PREFERENCE, pref, 4

    ' Tum surumler: region ile kirp.
    Dim pxW As Long, pxH As Long
    pxW = Me.ScaleWidth
    pxH = Me.ScaleHeight
    Dim hRgn As Long
    hRgn = CreateRoundRectRgn(0, 0, pxW + 1, pxH + 1, radius, radius)
    SetWindowRgn Me.hwnd, hRgn, True
End Sub
