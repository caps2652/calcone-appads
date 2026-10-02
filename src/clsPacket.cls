VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================================
' clsPacket.cls
' Tek bir protokol paketini temsil eden veri tasiyicisi.
' Paket cerceve bicimi (16 bayt baslik + payload):
'   [0..1]  magic  "RS" (&H52,&H53)
'   [2]     version
'   [3]     flags
'   [4..7]  type       (Long, little-endian)
'   [8..11] sessionId  (Long, little-endian)
'   [12..15]payloadLen (Long, little-endian)  -> [0 .. PKT_MAX_PAYLOAD]
' Cerceve olusturma/ayristirma modNetwork + clsConnection tarafindadir;
' bu sinif yalnizca alanlari ve payload yardimcilarini tutar.
'=====================================================================
Option Explicit

Public PType As Long
Public SessionID As Long
Public Flags As Long

Private mPayload() As Byte
Private mLen As Long

Private Sub Class_Initialize()
    mLen = 0
End Sub

'--- Payload'i ham baytlardan ayarla (ilk n bayt) ---
Public Sub SetPayloadBytes(ByRef b() As Byte, ByVal n As Long)
    If n <= 0 Then
        mLen = 0
        Erase mPayload
        Exit Sub
    End If
    ReDim mPayload(0 To n - 1)
    CopyMemory mPayload(0), b(LBound(b)), n
    mLen = n
End Sub

'--- Payload'i UTF-8 metinden ayarla ---
Public Sub SetPayloadString(ByVal s As String)
    Dim b() As Byte, n As Long
    b = StringToUtf8(s, n)
    SetPayloadBytes b, n
End Sub

Public Property Get PayloadLen() As Long
    PayloadLen = mLen
End Property

'--- Payload'in kopyasini (mLen bayt) dondur ---
Public Function PayloadBytes() As Byte()
    Dim b() As Byte
    If mLen > 0 Then
        ReDim b(0 To mLen - 1)
        CopyMemory b(0), mPayload(0), mLen
    Else
        ReDim b(0 To 0)   ' VB6'da bos dizi dondurmemek icin tek elemanli (len=0 kontrol edin)
    End If
    PayloadBytes = b
End Function

'--- Payload'i UTF-8 string olarak dondur ---
Public Function PayloadString() As String
    If mLen <= 0 Then
        PayloadString = ""
    Else
        PayloadString = Utf8ToString(mPayload, mLen)
    End If
End Function
