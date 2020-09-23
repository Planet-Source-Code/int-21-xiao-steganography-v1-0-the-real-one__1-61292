VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsStegano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*************************************
'BMP Header Struct
Private BmpHead As winBMPFileHeader
Private BmpInfo As BITMAPINFOHEADER
Private bmpPalette() As BITMAPPalette
Private DeepColor&
'*************************************

'*************************************
'BinaryAttach carried the data for each file added in binary format
'BinaryImg() carried tha data for the main image in binary format
Dim BinaryAttach() As tBits, BinaryImg() As tBits
'bAttachdata carried the data for each file added in byte format
'bImgData() carried tha data for the main image in byte format
Dim bImgData() As Byte, bAttachData() As Byte
'very hard to figure it out that?
'*************************************

Dim mImageFile$ 'Image Filename
Dim mOutputImageFile$ 'New Image Filename

Dim mFilesAdded& 'Count files added
Dim OutFile& 'Pointer to file
Dim mBytesLimit& 'Bytes limit to be added
Dim mBytesAdded& 'Bytes to attach

Dim colFiles As Collection 'My files's collection

Dim cTAG() As Byte 'the main tag to identify if the file carried any file attached

Event StatusChanged(prcDone As Long, strStatus As String) 'Raise this event to notify what whe are doing
Event SomeError(strDescription As String) 'Raise this event to notify when some error ocurr

' NewEnum tiene que devolver la interfaz IUnknown del
' enumerador de una colección.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
   Set NewEnum = colFiles.[_NewEnum]
End Function

Public Property Get ImageFile() As String
    ImageFile = mImageFile
End Property

Public Property Let ImageFile(ByVal vNewValue As String)
    mImageFile = vNewValue
    mBytesLimit = FileLen(mImageFile) / 8
End Property

Public Property Get FilesAdded() As Long
    FilesAdded = mFilesAdded
End Property
'AddFile
'strFile:the filename will be attach
'strTitle the Shortname fot this file, must be the same name with out the extension and the large path
'Key:the unique identify key for this file
Public Function AddFile(strFile As String, strTitle As String, Key As String) As Boolean
Dim tmpFile As ClsFile
On Local Error GoTo AddErr
    Set tmpFile = New ClsFile
    'fill data
    If FileExist(strFile) Then
        With tmpFile
            .KeyFile = Key
            .FileName = strFile
            .FileTitle = strTitle
            .LenBytes = FileLen(strFile) 'get len in bytes
            .TypeFile = VBA.Right$(strFile, 3) 'get type. (.exe,.txt,.bmp...)
            
            mBytesAdded = mBytesAdded + .LenBytes
            If mBytesAdded > mBytesLimit Then 'if the files to attach is too long, can't be carried
                mBytesAdded = mBytesAdded - .LenBytes
                Err.Raise 9001, "AddFile", "The File can't be add. Too long to be attach!"
            End If
        End With
        
    End If
    colFiles.Add tmpFile, Key
    mFilesAdded = mFilesAdded + 1
    AddFile = True
Exit Function
AddErr:
    RaiseEvent SomeError(Err.Description & " in " & Err.Source)
End Function

Public Function RemoveFile(Key As String) As Boolean
On Local Error GoTo AddErr
Dim tmpFile As ClsFile
    Set tmpFile = colFiles(Key) 'remove form the collection the file added
    mBytesAdded = mBytesAdded - tmpFile.LenBytes 'rest the bytes added too
    Set tmpFile = Nothing 'Free memory
    colFiles.Remove Key 'remove item
    RemoveFile = True
    mFilesAdded = mFilesAdded - 1
Exit Function
AddErr:
    RaiseEvent SomeError(Err.Description)
    Err.Clear
End Function

Public Function GetFile(Key As String) As ClsFile
Attribute GetFile.VB_UserMemId = 0
On Local Error GoTo GetErr
    Set GetFile = colFiles(Key) 'return info about any file added
Exit Function
GetErr:
    RaiseEvent SomeError(Err.Description)
    Err.Clear
End Function

Private Sub Class_Initialize()
    Set colFiles = New Collection
    cTAG() = StrConv("TAG:Int21", vbFromUnicode)
End Sub

Public Function EncodeIt() As Boolean
On Local Error GoTo EncodeErr
    
    If FileExist(mImageFile) Then 'Validate filename exist
        
        
        Dim tmpPalette As BITMAPPalette ' To calculate len of struct
    
        'Process data Image
        Call ReadImg_
        'convert image data to binary
        Call Convert2BinaryArray_(bImgData(), BinaryImg())
        
        DoEvents

        RaiseEvent StatusChanged(0, "Preparing data to be write...")
        
        OutFile = FreeFile 'The Main Buffer file
        'in this files we going to put all the data, TAG, and each file added
        Open "c:\tmp_C23F41AA.dat" For Binary As #OutFile
            
            Put #OutFile, , cTAG()
            Put #OutFile, , mFilesAdded
            
            RaiseEvent StatusChanged(0, "Please Wait...")
            ReadAttach_
            
        Close #OutFile
        
        ConvertAttach_
        Join_Img_Files_
        
        Kill "c:\tmp_C23F41AA.dat" 'delete buffer file
        
        RaiseEvent StatusChanged(100, "Encode done!")
        
    Else
        RaiseEvent SomeError("File doesn't exist") 'Dumb !!
    End If
    
Exit Function
EncodeErr:
    RaiseEvent SomeError(Err.Description)
    Err.Clear
    Close
End Function

Public Function DecodeIt() As Boolean
    If Not ReadTag_ Then 'Look for tag
        RaiseEvent SomeError("The selected image no contain any data to extract or haven't a Xiao format")
    Else
        ExtractData_
        DecodeIt = True 'return successful
    End If
End Function

Public Sub Save2Image()
Dim strDone$
If OutputImageFile <> "" Then
    
    RaiseEvent StatusChanged(0, "Saving file...")
    
    If Not SaveImg_() Then strDone = "Some error saving to new image" Else strDone = "Files was saved!"
    RaiseEvent StatusChanged(100, strDone)
    
Else
        RaiseEvent StatusChanged(0, "Image to save not was found!")
End If
    
End Sub

Private Function SaveImg_() As Boolean
Dim I&, J&, xFil&, lngCounter&
Dim maxArr&
On Local Error GoTo SaveImgErr
    'save to new file in disc our image with the file added
    maxArr = UBound(bImgData()) 'get max data image
    
    For J = 0 To UBound(BinaryImg()) 'Len image in binary format, must be equal LenImageInBytes * 8
        
        bImgData(I) = Bin2Asc(BinaryImg(J)) 'Convert the binary data to byte, 11111111 = 255
        
        I = I + 1
        
        If I > maxArr Then
            Exit For
        End If
            
        RaiseEvent StatusChanged(J * 100 / maxArr, "Saving new image...")
            
        DoEvents
    Next J
    
    xFil = FreeFile 'prepare our file to be write
    Open mOutputImageFile For Binary As #xFil
        Put #xFil, , BmpHead 'write header 1st
        Put #xFil, , BmpInfo '2th, write info
        'write the image data with the files hiden
        For lngCounter = 1 To DeepColor 'if exist..write palette data
            Put #1, , bmpPalette(lngCounter)
        Next lngCounter
        
        Put #xFil, , bImgData 'finally write the new data with our hide data
        
    Close #xFil 'end of the magic....=)
    SaveImg_ = True
Exit Function
SaveImgErr:
    RaiseEvent SomeError(Err.Description)
    Err.Clear
End Function

Private Sub ReadAttach_()
Dim xFil&, I&, lenBy&
Dim It As ClsFile
Dim vData() As Byte, strOut() As Byte
Dim Str3 As String * 3, Str10 As String * 10
Dim strShort$
On Local Error GoTo ReadAttachErr
    
xFil = FreeFile
'Read attach file
RaiseEvent StatusChanged(0, "Reading file to attach...")
I = 0
For Each It In colFiles 'read the files added in the image
    Open It.FileName For Binary As #xFil ' for each file added, build a new temp file in disc
        
        vData = InputB(LOF(xFil), #xFil)
        
        Str3 = It.TypeFile 'txt, bmp, jpg, gif, png
        Str10 = It.FileTitle 'the short name
        
        I = I + 1
        RaiseEvent StatusChanged((I * 100 / mFilesAdded), "Reading file to attach..." & Str10)
        
        strOut() = StrConv(Str3, vbFromUnicode)
        Put #OutFile, , strOut()
        Put #OutFile, , It.LenBytes
        strOut() = StrConv(Str10, vbFromUnicode)
        Put #OutFile, , strOut()
        
        Put #OutFile, , vData()
        
        
        DoEvents
    Close #xFil
Next
Exit Sub
ReadAttachErr:
RaiseEvent SomeError(Err.Description)
Err.Clear
End Sub

Private Sub ReadHeadImg_(pFile&)
Dim tmpPalette As BITMAPPalette
Dim I&
    'teh 1st step is read al header for the bitmap, and skip it, to going directly to the image data
    Get #pFile, , BmpHead 'fill head struct
    Get #pFile, , BmpInfo 'fill info struct
    
    'calculate deepcolor
    DeepColor = ((BmpHead.lngBitmapOffset - 54) / Len(tmpPalette))
        
    If DeepColor > 0 Then ReDim bmpPalette(1 To DeepColor) 'Rezise
    
    For I = 1 To DeepColor
        Get #pFile, , bmpPalette(I)
     Next I
    
End Sub

Private Sub ReadImg_()
Dim xFil&, LenBytes&, lngCounter&
Dim bytColor As Byte

'Read the Img File
xFil = FreeFile
Open mImageFile For Binary As #xFil

    RaiseEvent StatusChanged(0, "Reading Header...")
    
    ReadHeadImg_ xFil
    
    'Calculate len image data, without headers
    ReDim bImgData(0 To (BmpHead.lngFileSize - BmpHead.lngBitmapOffset) - 1)
    
    RaiseEvent StatusChanged(0, "Reading Image Data...")
    
    LenBytes = UBound(bImgData())
    
    For lngCounter = 0 To LenBytes ' this is the data where we going to hide our files
        If Not EOF(xFil) Then
            Get #xFil, , bytColor 'Read each rgb byte info
            bImgData(lngCounter) = bytColor
        End If
        
        RaiseEvent StatusChanged(lngCounter * 100 / LenBytes, "Reading Image Data...")
        
        DoEvents
    Next lngCounter
    
'ReadImg = DatImg()
Close #xFil
    
End Sub
'Look for our tag in the image file, if doesn't exist skip all
Private Function ReadTag_() As Boolean
Dim binData() As tBits, binTag() As tBits
Dim I&, J&, Cur&, bytColor As Byte
Dim strMyTag As String * 9
Dim lenStruct&, xFil&
Dim xyTb(0) As tBits
    
    RaiseEvent StatusChanged(0, "Searching header...")
    
    lenStruct = 9 'the len for the tag is always 9 bytes
    ReDim binTag(0 To lenStruct)
    
    xFil = FreeFile
    Open mImageFile For Binary As #xFil
        
        ReadHeadImg_ xFil  'Read header for bitmap
        
        lenStruct = 72 '8 bytes = 1 extra-byte, TAG= 9 bytes * 8 bytes = 72 bytes
        
        ReDim bImgData(0 To lenStruct)
        
        For I = 0 To lenStruct
            If Not EOF(xFil) Then
                Get #xFil, , bytColor
                bImgData(I) = bytColor
            End If
        Next I
    
    Close #xFil
    
    Call Convert2BinaryArray_(bImgData(), binData())
    
    lenStruct = UBound(binTag()) 'len data in binary
    
    Cur = 0
    lenStruct = 9 'the len tag is alway 9bytes
    For I = 0 To lenStruct
        For J = 0 To 7
            binTag(I).Bits(J) = binData(Cur).Bits(7)
            Cur = Cur + 1
            
        Next J
        If Cur >= 72 Then Exit For
    Next I
    
    strMyTag = Binary2String(binTag)
    
    ReadTag_ = (strMyTag = "TAG:Int21")
    
End Function

Private Sub ExtractData_()
Dim OutFile&, ImgFile&
Dim tmpFile&
Dim dataOut() As Byte
Dim BinOut() As tBits
Dim Bytes2Read&, Cur&, I&, J&
Dim bytColor As Byte
Dim sTAg$, lFA&, sTF$, lLF&, sNF$

    ImgFile& = FreeFile
    
    Open mImageFile$ For Binary As #ImgFile 'open the main image
        
        Call ReadHeadImg_(ImgFile)
        'skip the bmp header, to get the real image data
        Bytes2Read = LOF(ImgFile) - Loc(ImgFile)
        
        ReDim dataOut(0 To Bytes2Read)
        For I = 0 To (Bytes2Read)
            If Not EOF(ImgFile) Then
                Get #ImgFile, , bytColor
                dataOut(I) = bytColor
            End If
        Next I
    Close #ImgFile

        Call Convert2BinaryArray_(dataOut(), BinaryImg())
        
        Bytes2Read = UBound(BinaryImg()) 'len image in binary
        ReDim dataOut(0 To Bytes2Read)
        ReDim BinOut(0 To Bytes2Read)
        
        Cur = 0
        'we going to read the bytes 7 for each byte in the image data
        'and put it in other array to extract the hide data
        For I = 0 To (Bytes2Read)
            For J = 0 To 7
                If Cur >= Bytes2Read Then Exit For
                BinOut(I).Bits(J) = BinaryImg(Cur).Bits(7)
                Cur = Cur + 1
            Next J
            dataOut(I) = Bin2Asc(BinOut(I)) 'convert the binary hide in bytes
        Next I
        
        OutFile = FreeFile
        Open "c:\tmp_DD2741C.dat" For Binary As #OutFile 'tmp file to read data
            Put #OutFile, , dataOut()
        Close OutFile
        
        OutFile = FreeFile
        
        Open "c:\tmp_DD2741C.dat" For Binary As #OutFile 'tmp file to read data
            
            sTAg = ExtractItem_(OutFile, 9, 0, 1) 'Read the main tag
            lFA = ExtractItem_(OutFile, 4, 0, 0) 'read the number of files added
            Dim strFile$
            Dim It As ClsFile
            For I = 1 To lFA
                sTF = ExtractItem_(OutFile, 3, 0, 1) 'Read the type file(txt,bmp,gif,jpg,png)
                lLF = ExtractItem_(OutFile, 4, 0, 0) 'read the len in bytes for this file
                sNF = ExtractItem_(OutFile, 10, 0, 1) 'read the short name for this file
                strFile = "c:\" & sNF & "DD2741C." & sTF 'build the buffer filename
                
                tmpFile = FreeFile
                Open strFile For Binary As tmpFile
                    dataOut() = InputB(lLF, OutFile) 'read n-bytes, the len for this file
                    Put tmpFile, , dataOut() 'write in disc
                Close tmpFile
                
                AddFile strFile, sNF, CStr("c0" & I) 'add in the class
                'mBytesAdded = mBytesAdded + lLF 'counter the bytes added in the image
                
            Next
        
        Close OutFile
        
        Kill "c:\tmp_DD2741C.dat"
        
    
End Sub

Private Function ExtractItem_(pFile As Long, Bytes2Read As Long, Bytes2Look As Long, RetType As Integer)
Dim Memo() As Byte
Dim lLong&
Dim strEnd$
    Memo() = InputB(Bytes2Read, pFile) 'read n-bytes from disc
    
If RetType = 0 Then 'Numeric
    CopyMemory lLong, Memo(0), Len(lLong)
    ExtractItem_ = lLong
ElseIf RetType = 1 Then 'String
    strEnd = Memo()
    ExtractItem_ = StrConv(strEnd, vbUnicode)
End If

End Function

Private Sub ConvertAttach_()
Dim byt As Byte
Dim LenF&, I&
'Read all files added and convert to binary
    OutFile = FreeFile
    LenF = FileLen("c:\tmp_C23F41AA.dat")
    ReDim bAttachData(0 To LenF)
    
    Open "c:\tmp_C23F41AA.dat" For Binary As #OutFile
        Do While Not EOF(OutFile)
            Get OutFile, , byt
            bAttachData(I) = byt
            I = I + 1
        Loop
        
    Close #OutFile

    Call Convert2BinaryArray_(bAttachData(), BinaryAttach())
    
End Sub
'the magic function, joing image and files to attach in only one file
Private Sub Join_Img_Files_()
Dim I&, J&, K&, LenImg&, LenF&
    
    LenImg = UBound(BinaryImg()) 'len in binary of image
    LenF = UBound(BinaryAttach()) 'len in binary for files to attach
    I = 0
    
    For J = 0 To LenF
        For K = 0 To 7
            BinaryImg(I).Bits(7) = BinaryAttach(J).Bits(K) 'put one bit from binary data to hide in the bit 7
            I = I + 1
        Next K
        If I >= LenImg Then Exit For
        RaiseEvent StatusChanged((I * 100 / LenImg), "Joining files with image...")
        DoEvents
    Next J
    
End Sub
'Convert2BinaryArray_
'Source(): the file data in bytes
'retArray(): the Binary data to be return
Private Sub Convert2BinaryArray_(Source() As Byte, RetArray() As tBits)
Dim LenArray&, I&
Dim arrBinary() As tBits
Dim Bits8 As tBits

    LenArray = UBound(Source())
    
     ReDim arrBinary(0 To LenArray)
     
     For I = 0 To LenArray
        Bits8 = ByteToBinary(Source(I)) 'convert 1 byte to binary
        arrBinary(I) = Bits8
        RaiseEvent StatusChanged((I * 100 / LenArray), "Convert Hex to Binary...")
        
        DoEvents

     Next I

RetArray = arrBinary

End Sub

Private Function FileExist(strFile As String) As Boolean
Dim Rs$, Tama As Boolean
Dim Tm&
    Rs = Dir(strFile)
    FileExist = (Len(Rs))
End Function

Public Property Get OutputImageFile() As String
    OutputImageFile = mOutputImageFile
End Property

Public Property Let OutputImageFile(ByVal sNewFile As String)
    mOutputImageFile = sNewFile
End Property

Private Sub Class_Terminate()
Dim tmpClass As ClsFile
For Each tmpClass In colFiles
    Set tmpClass = Nothing
Next
Set colFiles = Nothing
End Sub

Public Property Get BytesAdded() As Long
    BytesAdded = mBytesAdded
End Property


Public Property Get BytesTotal() As Variant
    BytesTotal = mBytesLimit
End Property

