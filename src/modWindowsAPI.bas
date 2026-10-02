Attribute VB_Name = "modWindowsAPI"
'=====================================================================
' modWindowsAPI.bas
' Tum Windows API bildirimleri ve paylasilan Type/Const tanimlari.
' Kod kalitesi kurali: API declaration'lari tek bir yerde tutulur.
' Unicode icin mumkun oldugunca "W" (wide) fonksiyonlar kullanilir.
'=====================================================================
Option Explicit

'--------------------------------------------------------------------
' Genel yardimci tipler
'--------------------------------------------------------------------
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type SIZEAPI
    cx As Long
    cy As Long
End Type

'--------------------------------------------------------------------
' Winsock (ws2_32.dll) tipleri
'--------------------------------------------------------------------
Public Type sockaddr_in
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(0 To 7) As Byte
End Type

Private Const WSADESCRIPTION_LEN As Long = 256
Private Const WSASYS_STATUS_LEN As Long = 128

Public Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADESCRIPTION_LEN) As Byte
    szSystemStatus(0 To WSASYS_STATUS_LEN) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

'--------------------------------------------------------------------
' Winsock sabitleri
'--------------------------------------------------------------------
Public Const AF_INET As Long = 2
Public Const SOCK_STREAM As Long = 1
Public Const IPPROTO_TCP As Long = 6
Public Const INADDR_ANY As Long = 0
Public Const INADDR_NONE As Long = &HFFFFFFFF
Public Const SOCKET_ERROR As Long = -1
Public Const INVALID_SOCKET As Long = -1
Public Const SOMAXCONN As Long = 5

' setsockopt
Public Const SOL_SOCKET As Long = &HFFFF&
Public Const SO_REUSEADDR As Long = &H4
Public Const TCP_NODELAY As Long = &H1

' WSAAsyncSelect event bayraklari
Public Const FD_READ As Long = &H1
Public Const FD_WRITE As Long = &H2
Public Const FD_ACCEPT As Long = &H8
Public Const FD_CONNECT As Long = &H10
Public Const FD_CLOSE As Long = &H20

' Winsock hata kodlari (ihtiyac duyulanlar)
Public Const WSAEWOULDBLOCK As Long = 10035
Public Const WSAECONNRESET As Long = 10054
Public Const WSAECONNREFUSED As Long = 10061
Public Const WSAETIMEDOUT As Long = 10060

'--------------------------------------------------------------------
' Winsock API
'--------------------------------------------------------------------
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, lpWSAData As WSADATA) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Public Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
Public Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Public Declare Function api_socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal sType As Long, ByVal protocol As Long) As Long
Public Declare Function api_closesocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long
Public Declare Function api_bind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, addr As sockaddr_in, ByVal namelen As Long) As Long
Public Declare Function api_listen Lib "ws2_32.dll" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Public Declare Function api_accept Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, addr As sockaddr_in, addrlen As Long) As Long
Public Declare Function api_connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, addr As sockaddr_in, ByVal namelen As Long) As Long
Public Declare Function api_send Lib "ws2_32.dll" Alias "send" (ByVal s As Long, buf As Any, ByVal length As Long, ByVal flags As Long) As Long
Public Declare Function api_recv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, buf As Any, ByVal length As Long, ByVal flags As Long) As Long
Public Declare Function api_setsockopt Lib "ws2_32.dll" Alias "setsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Public Declare Function api_htons Lib "ws2_32.dll" Alias "htons" (ByVal hostshort As Long) As Integer
Public Declare Function api_ntohs Lib "ws2_32.dll" Alias "ntohs" (ByVal netshort As Integer) As Long
Public Declare Function api_inet_addr Lib "ws2_32.dll" Alias "inet_addr" (ByVal cp As String) As Long
Public Declare Function api_gethostname Lib "ws2_32.dll" Alias "gethostname" (ByVal name As String, ByVal namelen As Long) As Long

'--------------------------------------------------------------------
' Pencere / mesaj API (user32.dll) - subclassing ve mesaj penceresi
'--------------------------------------------------------------------
Public Const WM_USER As Long = &H400
Public Const WM_SOCKET As Long = WM_USER + 1      ' kendi socket bildirim mesajimiz
Public Const HWND_MESSAGE As Long = -3            ' message-only pencere parent

Public Const GWL_WNDPROC As Long = -4
Public Const WM_NCLBUTTONDOWN As Long = &HA1
Public Const HTCAPTION As Long = 2
Public Const WM_SETCURSOR As Long = &H20

Public Declare Function RegisterWindowMessageW Lib "user32" (ByVal lpString As Long) As Long
Public Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DefWindowProcW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProcW Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function LoadCursorW Lib "user32" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Const IDC_ARROW As Long = 32512
Public Const IDC_HAND As Long = 32649
Public Const IDC_IBEAM As Long = 32513

'--------------------------------------------------------------------
' DPI farkindaligi
'--------------------------------------------------------------------
Public Const LOGPIXELSX As Long = 88
Public Declare Function SetProcessDPIAware Lib "user32" () As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

'--------------------------------------------------------------------
' Pencere kose yuvarlatma: GDI region (tum Windows) + DWM (Win11)
'--------------------------------------------------------------------
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

' DWM Win11 kose tercihi (dwmapi.dll her sistemde olmayabilir -> hata yakalanir)
Public Const DWMWA_WINDOW_CORNER_PREFERENCE As Long = 33
Public Const DWMWCP_ROUND As Long = 2
Public Declare Function DwmSetWindowAttribute Lib "dwmapi.dll" (ByVal hWnd As Long, ByVal dwAttribute As Long, pvAttribute As Any, ByVal cbAttribute As Long) As Long

'--------------------------------------------------------------------
' GDI cizim API
'--------------------------------------------------------------------
Public Const PS_SOLID As Long = 0
Public Const BS_SOLID As Long = 0
Public Const TRANSPARENT As Long = 1
Public Const OPAQUE As Long = 2

' DrawTextW bayraklari
Public Const DT_LEFT As Long = &H0
Public Const DT_CENTER As Long = &H1
Public Const DT_RIGHT As Long = &H2
Public Const DT_TOP As Long = &H0
Public Const DT_VCENTER As Long = &H4
Public Const DT_SINGLELINE As Long = &H20
Public Const DT_WORDBREAK As Long = &H10
Public Const DT_END_ELLIPSIS As Long = &H8000
Public Const DT_NOPREFIX As Long = &H800
Public Const DT_CALCRECT As Long = &H400

' Font sabitleri
Public Const FW_NORMAL As Long = 400
Public Const FW_SEMIBOLD As Long = 600
Public Const FW_BOLD As Long = 700
Public Const DEFAULT_CHARSET As Long = 1
Public Const CLEARTYPE_QUALITY As Long = 5
Public Const DEFAULT_PITCH As Long = 0
Public Const FF_DONTCARE As Long = 0

Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function CreateFontW Lib "gdi32" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As Long) As Long

'--------------------------------------------------------------------
' Kernel / bellek / INI / zaman
'--------------------------------------------------------------------
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal length As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function WritePrivateProfileStringW Lib "kernel32" (ByVal lpAppName As Long, ByVal lpKeyName As Long, ByVal lpString As Long, ByVal lpFileName As Long) As Long
Public Declare Function GetPrivateProfileStringW Lib "kernel32" (ByVal lpAppName As Long, ByVal lpKeyName As Long, ByVal lpDefault As Long, ByVal lpReturnedString As Long, ByVal nSize As Long, ByVal lpFileName As Long) As Long

'--------------------------------------------------------------------
' CryptoAPI (advapi32.dll) - guvenli rastgele deger ve SHA-256
'--------------------------------------------------------------------
Public Const PROV_RSA_AES As Long = 24
Public Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Public Const CALG_SHA_256 As Long = &H800C&
Public Const HP_HASHVAL As Long = 2

Public Declare Function CryptAcquireContextW Lib "advapi32.dll" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptGenRandom Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwLen As Long, pbBuffer As Any) As Long
Public Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal algId As Long, ByVal hKey As Long, ByVal dwFlags As Long, phHash As Long) As Long
Public Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
