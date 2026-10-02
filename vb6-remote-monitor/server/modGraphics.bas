Attribute VB_Name = "modGraphics"
'=============================================================================
' modGraphics.bas  -  Ekran yakalama + bellekte JPEG kodlama/cozme (GDI+)
' OCX YOK. gdi32 (BitBlt) + gdiplus (JPEG) + ole32 (IStream) API ile calisir.
' Diske HICBIR dosya yazilmaz; tum islem bellekte (RAM) yapilir.
' Server ve Client projelerinde ayni modul kullanilir.
'=============================================================================
Option Explicit

'--- GDI / cizim -----------------------------------------------------------
Private Const SRCCOPY As Long = &HCC0020
Private Const CAPTUREBLT As Long = &H40000000
Private Const HALFTONE As Long = 4

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

'--- Bellek / COM stream ---------------------------------------------------
Private Const GMEM_MOVEABLE As Long = &H2
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal numBytes As Long)
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Long) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal pstm As Long, phglobal As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Long, pclsid As GUID_T) As Long

'--- GDI+ ------------------------------------------------------------------
Private Type GUID_T
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type EncoderParameter
    GUID As GUID_T
    NumberOfValues As Long
    ParamType As Long
    Value As Long
End Type

Private Type EncoderParameters
    Count As Long
    Parameter As EncoderParameter
End Type

Private Const PixelFormat24bppRGB As Long = &H21808
Private Const UnitPixel As Long = 2
Private Const InterpolationModeLowQuality As Long = 1
Private Const ColorAdjustTypeDefault As Long = 0
Private Const ColorMatrixFlagsDefault As Long = 0
Private Const EncoderParameterValueTypeLong As Long = 4

Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hpal As Long, bitmap As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal pxFormat As Long, ByVal scan0 As Long, bitmap As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal image As Long, graphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As Long, ByVal mode As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (imageattr As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageattr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageattr As Long, ByVal clrAdjType As Long, ByVal enableFlag As Long, clrMatrix As Any, ByVal grayMatrix As Long, ByVal flags As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstw As Long, ByVal dsth As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcw As Long, ByVal srch As Long, ByVal srcUnit As Long, ByVal imageAttr As Long, ByVal callback As Long, ByVal callbackData As Long) As Long
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal image As Long, ByVal stream As Long, clsidEncoder As GUID_T, encoderParams As Any) As Long
Private Declare Function GdipCreateBitmapFromStream Lib "gdiplus" (ByVal stream As Long, bitmap As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal bitmap As Long, hbmReturn As Long, ByVal background As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, Height As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdipFlush Lib "gdiplus" (ByVal graphics As Long, ByVal intention As Long) As Long

Private mToken As Long
Private mStarted As Boolean

'=== Baslatma / kapatma ====================================================
Public Function GfxStartup() As Boolean
    Dim si As GdiplusStartupInput
    If mStarted Then GfxStartup = True: Exit Function
    si.GdiplusVersion = 1
    If GdiplusStartup(mToken, si, 0) = 0 Then
        mStarted = True
        GfxStartup = True
    End If
End Function

Public Sub GfxShutdown()
    If mStarted Then
        GdiplusShutdown mToken
        mStarted = False
    End If
End Sub

'=== SERVER: ekrani yakala, kucult, (istege bagli) grileştir, JPEG uret ====
' quality 1-100, scalePercent orn. 50, grayscale True/False
' Basari: JPEG byte sayisi (outBytes doldurulur). Hata: 0
Public Function CaptureFrameJpeg(ByVal quality As Long, ByVal scalePercent As Long, ByVal grayscale As Boolean, ByRef outBytes() As Byte) As Long
    Dim hdcScreen As Long, hdcMem As Long, hbm As Long, hbmOld As Long
    Dim scrW As Long, scrH As Long, dstW As Long, dstH As Long
    Dim srcBmp As Long, dstBmp As Long, g As Long, imgAttr As Long
    Dim pStream As Long, result As Long

    If Not mStarted Then CaptureFrameJpeg = 0: Exit Function
    If scalePercent < 5 Then scalePercent = 5
    If scalePercent > 100 Then scalePercent = 100
    If quality < 1 Then quality = 1
    If quality > 100 Then quality = 100

    scrW = GetSystemMetrics(0)          ' SM_CXSCREEN
    scrH = GetSystemMetrics(1)          ' SM_CYSCREEN
    dstW = (scrW * scalePercent) \ 100
    dstH = (scrH * scalePercent) \ 100
    If dstW < 1 Then dstW = 1
    If dstH < 1 Then dstH = 1

    ' 1) Ekrani bellek DC'ye kopyala
    hdcScreen = GetDC(0)
    hdcMem = CreateCompatibleDC(hdcScreen)
    hbm = CreateCompatibleBitmap(hdcScreen, scrW, scrH)
    hbmOld = SelectObject(hdcMem, hbm)
    BitBlt hdcMem, 0, 0, scrW, scrH, hdcScreen, 0, 0, SRCCOPY
    SelectObject hdcMem, hbmOld   ' hbm'i DC'den cikar ki GDI+ kullanabilsin

    ' 2) GDI+ kaynak bitmap
    GdipCreateBitmapFromHBITMAP hbm, 0, srcBmp

    ' 3) Hedef (kucuk) bitmap + grafik baglami
    GdipCreateBitmapFromScan0 dstW, dstH, 0, PixelFormat24bppRGB, 0, dstBmp
    GdipGetImageGraphicsContext dstBmp, g
    GdipSetInterpolationMode g, InterpolationModeLowQuality   ' hiz icin

    ' 4) Grileştirme (istege bagli) -> ColorMatrix
    '    NOT: GDI+ matrisi satir-oncelikli (row-major) ister. VB6 cok boyutlu
    '    diziyi sutun-oncelikli sakladigindan, matrisi DUZ 1B dizi ile tam
    '    native sirada (index = satir*5 + sutun) veriyoruz. Boylece transpoze
    '    karisikligi olmaz.
    If grayscale Then
        Dim cm(0 To 24) As Single
        ' satir0 (R girisi): tum cikis kanallarina 0.299
        cm(0) = 0.299: cm(1) = 0.299: cm(2) = 0.299
        ' satir1 (G girisi): 0.587
        cm(5) = 0.587: cm(6) = 0.587: cm(7) = 0.587
        ' satir2 (B girisi): 0.114
        cm(10) = 0.114: cm(11) = 0.114: cm(12) = 0.114
        ' satir3 (Alpha) -> aynen koru
        cm(18) = 1
        ' satir4 (sabit/translate)
        cm(24) = 1
        GdipCreateImageAttributes imgAttr
        GdipSetImageAttributesColorMatrix imgAttr, ColorAdjustTypeDefault, 1, cm(0), 0, ColorMatrixFlagsDefault
    End If

    ' 5) Kaynagi kucultup hedefe ciz
    GdipDrawImageRectRectI g, srcBmp, 0, 0, dstW, dstH, 0, 0, scrW, scrH, UnitPixel, imgAttr, 0, 0
    GdipFlush g, 1   ' cizimi bitmap'e yaz (FlushIntentionSync)

    ' 6) Hedef bitmap'i bellekte JPEG'e kodla
    result = EncodeToJpegBytes(dstBmp, quality, outBytes)

    ' 7) Temizlik
    If imgAttr <> 0 Then GdipDisposeImageAttributes imgAttr
    If g <> 0 Then GdipDeleteGraphics g
    If dstBmp <> 0 Then GdipDisposeImage dstBmp
    If srcBmp <> 0 Then GdipDisposeImage srcBmp
    DeleteObject hbm
    DeleteDC hdcMem
    ReleaseDC 0, hdcScreen

    CaptureFrameJpeg = result
End Function

'--- Bir GDI+ image'i JPEG byte dizisine kodla (bellek icinde) -----------
Private Function EncodeToJpegBytes(ByVal image As Long, ByVal quality As Long, ByRef outBytes() As Byte) As Long
    Dim jpegClsid As GUID_T, qualGuid As GUID_T
    Dim ep As EncoderParameters, qv As Long
    Dim pStream As Long, hMem As Long, pData As Long, sz As Long

    ' JPEG encoder CLSID: {557CF401-1A04-11D3-9A73-0000F81EF32E}
    CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), jpegClsid
    ' EncoderQuality GUID: {1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}
    CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), qualGuid

    qv = quality
    ep.Count = 1
    ep.Parameter.GUID = qualGuid
    ep.Parameter.NumberOfValues = 1
    ep.Parameter.ParamType = EncoderParameterValueTypeLong
    ep.Parameter.Value = VarPtr(qv)

    If CreateStreamOnHGlobal(0, 1, pStream) <> 0 Then EncodeToJpegBytes = 0: Exit Function

    If GdipSaveImageToStream(image, pStream, jpegClsid, ep) = 0 Then
        GetHGlobalFromStream pStream, hMem
        If hMem <> 0 Then
            sz = GlobalSize(hMem)
            pData = GlobalLock(hMem)
            If sz > 0 And pData <> 0 Then
                ReDim outBytes(0 To sz - 1)
                CopyMemory outBytes(0), ByVal pData, sz
                EncodeToJpegBytes = sz
            End If
            GlobalUnlock hMem
        End If
    End If

    ReleaseStream pStream
End Function

'=== CLIENT: JPEG byte dizisini cozup HBITMAP uret (ekranda gostermek icin)
' Basari: HBITMAP handle (kullandiktan sonra DeleteObject ile sil). Hata: 0
Public Function DecodeJpegToHBitmap(ByRef inBytes() As Byte, ByVal length As Long, ByRef outW As Long, ByRef outH As Long) As Long
    Dim hMem As Long, pData As Long, pStream As Long
    Dim bmp As Long, hbm As Long

    If Not mStarted Then DecodeJpegToHBitmap = 0: Exit Function
    If length <= 0 Then DecodeJpegToHBitmap = 0: Exit Function

    hMem = GlobalAlloc(GMEM_MOVEABLE, length)
    If hMem = 0 Then DecodeJpegToHBitmap = 0: Exit Function
    pData = GlobalLock(hMem)
    CopyMemory ByVal pData, inBytes(0), length
    GlobalUnlock hMem

    ' fDeleteOnRelease=1 -> stream serbest birakilinca hMem de serbest kalir
    If CreateStreamOnHGlobal(hMem, 1, pStream) <> 0 Then DecodeJpegToHBitmap = 0: Exit Function

    If GdipCreateBitmapFromStream(pStream, bmp) = 0 Then
        GdipGetImageWidth bmp, outW
        GdipGetImageHeight bmp, outH
        GdipCreateHBITMAPFromBitmap bmp, hbm, 0
        GdipDisposeImage bmp
    End If

    ReleaseStream pStream
    DecodeJpegToHBitmap = hbm
End Function

'--- CLIENT: bir HBITMAP'i hedef hDC'ye gererek ciz (PictureBox icin) -----
Public Sub DrawHBitmapStretched(ByVal destDC As Long, ByVal hbm As Long, ByVal destW As Long, ByVal destH As Long, ByVal srcW As Long, ByVal srcH As Long)
    Dim memDC As Long, oldBmp As Long
    If hbm = 0 Or srcW <= 0 Or srcH <= 0 Then Exit Sub
    memDC = CreateCompatibleDC(destDC)
    oldBmp = SelectObject(memDC, hbm)
    SetStretchBltMode destDC, HALFTONE
    StretchBlt destDC, 0, 0, destW, destH, memDC, 0, 0, srcW, srcH, SRCCOPY
    SelectObject memDC, oldBmp
    DeleteDC memDC
End Sub

Public Sub DeleteBitmap(ByVal hbm As Long)
    If hbm <> 0 Then DeleteObject hbm
End Sub

'--- IStream pointer'ini guvenle serbest birak (COM Release) --------------
Private Sub ReleaseStream(ByVal pStream As Long)
    Dim u As IUnknown
    If pStream = 0 Then Exit Sub
    ' Ham pointer'i IUnknown'a tasimaci yap (AddRef yapmadan sahiplen),
    ' Nothing atayinca tam olarak bir kez Release edilir.
    CopyMemory u, pStream, 4
    Set u = Nothing
End Sub
