Attribute VB_Name = "Module"
Option Explicit

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Type tBits
    Bits(0 To 7) As Single
End Type

'Public vHex, vBin

'file header, total 14 bytes
Type winBMPFileHeader
     strFileType As String * 2 ' file type always 4D42h or "BM"
     lngFileSize As Long       'size in bytes ussually 0 for uncompressed
     bytReserved1 As Integer   ' always 0
     bytReserved2 As Integer   ' always 0
     lngBitmapOffset As Long   'starting position of image data in bytes
End Type


'image header, total 40 bytes
Type BITMAPINFOHEADER
     biSize As Long          'Size of this header
     biWidth As Long         'width of your image
     biHeight As Long        'height of your image
     biPlanes As Integer     'always 1
     byBitCount As Integer   'number of bits per pixel 1, 4, 8, or 24
     biCompression As Long   '0 data is not compressed
     biSizeImage As Long     'size of bitmap in bytes, typicaly 0 when uncompressed
     biXPelsPerMeter As Long 'preferred resolution in pixels per meter
     biYPelsPerMeter As Long 'preferred resolution in pixels per meter
     biClrUsed As Long       'number of colors that are actually used (can be 0)
     biClrImportant As Long  'which color is most important (0 means all of them)
End Type

'palette, 4 bytes * 256 = 1024
Type BITMAPPalette
     lngBlue As Byte
     lngGreen As Byte
     lngRed As Byte
     lngReserved As Byte
End Type

Function Binary2String(laData() As tBits)
Dim ArrEnd() As Byte
Dim strEnd$, I&
    
    ReDim ArrEnd(0 To UBound(laData()))
    strEnd$ = ""
    For I = 0 To UBound(laData())
        strEnd = strEnd & Chr(Bin2Asc(laData(I)))
    Next I
strEnd = VBA.Left$(strEnd, Len(strEnd) - 1)
 Binary2String = strEnd
End Function

'Function String2Binary(theStr As String, RetArray() As tBits, olbStatus As Label, ProgBar As ProgressBar)
'Dim I&, hexRes$, LenBy&
'Dim arrHexStr() As String, arrHexBy() As Byte, BinRes() As tBits
'
'    For I = 1 To Len(theStr)
'        hexRes = hexRes & Asc(Mid(theStr, I, 1)) & ","
'    Next I
'
'    arrHexStr() = Split(hexRes, ",")
'
'    LenBy = UBound(arrHexStr)
'    ReDim arrHexBy(0 To LenBy)
'
'    For I = 0 To LenBy - 1
'        arrHexBy(I) = CByte(arrHexStr(I))
'        ProgBar.Value = I * 100 / LenBy
'    Next I
'
'    Convert2BinaryArray arrHexBy(), BinRes(), olbStatus, ProgBar
'
'RetArray = BinRes()
'End Function

'Function Bin2Hex(Bin As tBits)
'Dim Nibble1$, Nibble2$, I&
'Dim Res$
'    For I = 0 To 3
'        Nibble1 = Nibble1 & Bin.Bits(I)
'    Next I
'    For I = 4 To 7
'        Nibble2 = Nibble2 & Bin.Bits(I)
'    Next I
'    For I = 0 To UBound(vBin)
'        If (Nibble1 = vBin(I)) And Res = "" Then
'            Res = vHex(I)
'            I = 0
'        End If
'        If (Nibble2 = vBin(I)) And Res <> "" Then
'            Res = Res & vHex(I)
'            Exit For
'        End If
'    Next I
'Bin2Hex = Res
'End Function

Function Bin2Asc(Bin As tBits) As Integer
Dim num As Integer
Dim Fact%, I&
    num = 0
    Fact = 128
    For I = 0 To 7
        num = num + Bin.Bits(I) * Fact
        Fact = Fact / 2
    Next I
    Bin2Asc = num
End Function

Function Convert2BinaryArray(laData() As Byte, RetArray() As tBits, olbStatus As Label, ProgBar As ProgressBar)
Dim LenArray&, I&, J&, K&
Dim strBin$, Arr() As String
Dim arrBinary() As tBits
Dim Bits8 As tBits
    olbStatus = "Convert Hex to Binary..."
    
    LenArray = UBound(laData())
     
     ReDim arrBinary(0 To LenArray)
     
     For I = 0 To LenArray
        'strBin = ByteToBinary(laData(I))
        Bits8 = ByteToBinary(laData(I))
        K = 1
        For J = 0 To 7
            arrBinary(I).Bits(J) = Mid(strBin, K, 1)
            K = K + 1
        Next J
        
        ProgBar.Value = I * 100 / LenArray
        DoEvents
     Next I

RetArray = arrBinary

End Function

Function ByteToBinary(ByVal data As Byte) As tBits
Dim tmpBit As tBits
    Dim I As Long, J&
    
    I = &H80 '10000000
    
    While I
        tmpBit.Bits(J) = IIf(data And I, "1", "0")
        I = I \ 2
        J = J + 1
    Wend
ByteToBinary = tmpBit
    
End Function
