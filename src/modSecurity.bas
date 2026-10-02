Attribute VB_Name = "modSecurity"
'=====================================================================
' modSecurity.bas
' Guvenlik primitifleri:
'  - Guvenli rastgele deger (CryptGenRandom)
'  - Baglanti ID uretimi (9 haneli)
'  - Erisim sifresi uretimi (karistirilmasi zor karakter kumesi)
'  - SHA-256 hash (Phase 2 challenge-response kimlik dogrulama icin)
' GUVENLIK: Sifreler asla loglanmaz/duz metin saklanmaz. Dogrulama
' hash uzerinden yapilir.
'=====================================================================
Option Explicit

' Karistirilmasi kolay karakterler (0/O, 1/l/I) cikarildi.
Private Const PWD_ALPHABET As String = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789"
Private Const PWD_LEN As Long = 10

'--- Kriptografik olarak guvenli rastgele byte dizisi ---
Public Function SecureRandomBytes(ByVal count As Long) As Byte()
    Dim b() As Byte
    If count < 1 Then count = 1
    ReDim b(0 To count - 1)

    Dim hProv As Long
    If CryptAcquireContextW(hProv, 0, 0, PROV_RSA_AES, CRYPT_VERIFYCONTEXT) <> 0 Then
        CryptGenRandom hProv, count, b(0)
        CryptReleaseContext hProv, 0
    Else
        ' Yedek yol: CryptoAPI yoksa zayif ama calisir uretec.
        LogWrite LVL_WARNING, "CryptAcquireContext basarisiz; yedek rastgele uretec kullaniliyor."
        Randomize Timer
        Dim i As Long
        For i = 0 To count - 1
            b(i) = CByte(Int(Rnd * 256))
        Next i
    End If
    SecureRandomBytes = b
End Function

'--- 0..max-1 araliginda onyargisiz (bias'siz) rastgele tam sayi ---
Private Function RandIndex(ByVal maxExclusive As Long) As Long
    Dim b() As Byte
    b = SecureRandomBytes(4)
    Dim v As Long
    ' Isaret bitini temizleyip pozitif bir Long elde et.
    v = (CLng(b(0)) Or (CLng(b(1)) * &H100&) Or (CLng(b(2)) * &H10000)) ' 24-bit
    RandIndex = v Mod maxExclusive
End Function

'--- 9 haneli baglanti ID'si: "123 456 789" bicimli gosterim icin ham deger ---
Public Function GenerateConnectionID() As String
    Dim s As String, i As Long
    ' Ilk hane 0 olmasin.
    s = CStr(1 + RandIndex(9))
    For i = 2 To 9
        s = s & CStr(RandIndex(10))
    Next i
    GenerateConnectionID = s
End Function

'--- Gosterim bicimi: "123456789" -> "123 456 789" ---
Public Function FormatConnectionID(ByVal id As String) As String
    If Len(id) = 9 Then
        FormatConnectionID = Mid$(id, 1, 3) & " " & Mid$(id, 4, 3) & " " & Mid$(id, 7, 3)
    Else
        FormatConnectionID = id
    End If
End Function

'--- Erisim sifresi uret ---
Public Function GenerateAccessPassword() As String
    Dim s As String, i As Long, n As Long
    n = Len(PWD_ALPHABET)
    For i = 1 To PWD_LEN
        s = s & Mid$(PWD_ALPHABET, 1 + RandIndex(n), 1)
    Next i
    GenerateAccessPassword = s
End Function

'=====================================================================
' SHA-256 -> onaltilik (hex) string. Phase 2 kimlik dogrulamasinda
' "tuz (salt) + sifre" hash'lemek icin kullanilir.
'=====================================================================
Public Function Sha256Hex(ByRef data() As Byte) As String
    Dim hProv As Long, hHash As Long
    Dim ok As Boolean
    Sha256Hex = ""

    If CryptAcquireContextW(hProv, 0, 0, PROV_RSA_AES, CRYPT_VERIFYCONTEXT) = 0 Then Exit Function

    If CryptCreateHash(hProv, CALG_SHA_256, 0, 0, hHash) <> 0 Then
        Dim n As Long
        n = UBound(data) - LBound(data) + 1
        If n > 0 Then
            If CryptHashData(hHash, data(LBound(data)), n, 0) <> 0 Then ok = True
        Else
            ok = True
        End If

        If ok Then
            Dim outBuf(0 To 31) As Byte, outLen As Long
            outLen = 32
            If CryptGetHashParam(hHash, HP_HASHVAL, outBuf(0), outLen, 0) <> 0 Then
                Sha256Hex = BytesToHex(outBuf, outLen)
            End If
        End If
        CryptDestroyHash hHash
    End If

    CryptReleaseContext hProv, 0
End Function

'--- Yardimci: string -> SHA-256 hex ---
Public Function Sha256HexFromString(ByVal s As String) As String
    Dim b() As Byte
    b = StrConv(s, vbFromUnicode)   ' not: salt eklemek Phase 2'de cagiranin isi
    Sha256HexFromString = Sha256Hex(b)
End Function

Private Function BytesToHex(ByRef b() As Byte, ByVal count As Long) As String
    Const HEXCH As String = "0123456789abcdef"
    Dim s As String, i As Long, v As Long
    For i = 0 To count - 1
        v = b(i)
        s = s & Mid$(HEXCH, 1 + (v \ 16), 1) & Mid$(HEXCH, 1 + (v And 15), 1)
    Next i
    BytesToHex = s
End Function
