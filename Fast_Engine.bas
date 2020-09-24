Attribute VB_Name = "Fast_Engine"

Option Explicit

Type BITMAP
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

'Used to ensure quality stretching of color images
Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dX As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

'Standard pixel data
Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbAlpha As Byte
End Type

'Full-size bitmap header
Type BITMAPINFOHEADER
        bmSize As Long
        bmWidth As Long
        bmHeight As Long
        bmPlanes As Integer
        bmBitCount As Integer
        bmCompression As Long
        bmSizeImage As Long
        bmXPelsPerMeter As Long
        bmYPelsPerMeter As Long
        bmClrUsed As Long
        bmClrImportant As Long
End Type

'Extended header for 8-bit images
Type BITMAPINFO
        bmHeader As BITMAPINFOHEADER
        bmColors(0 To 255) As RGBQUAD
End Type

'Convert to absolute byte values (Long-type)
Public Sub ByteMe(ByRef TempVar As Long)
    If TempVar > 255 Then TempVar = 255
    If TempVar < 0 Then TempVar = 0
End Sub

'Get an image's pixel information into an array dimensioned (x * 3 + bgr, y), with the option to get it in its true orientation
Public Sub GetImageData2D(SrcPictureBox As PictureBox, ImageData() As Byte, Optional ByVal CorrectOrientation As Boolean = False)
    Dim bm As BITMAP
    'Get the picture box information
    GetObject SrcPictureBox.Image, Len(bm), bm
    'Build a correctly sized array
    Erase ImageData()
    'Generate a correctly-dimensioned array (for 2-dimensional access)
    Dim ArrayWidth As Long
    ArrayWidth = (bm.bmWidth * 3) - 1
    ArrayWidth = ArrayWidth + (bm.bmWidth Mod 4)  '4-bit alignment
    ReDim ImageData(0 To ArrayWidth, 0 To bm.bmHeight) As Byte
    ReDim tmpData(0 To ArrayWidth, 0 To bm.bmHeight) As Byte
    
    'Create a temporary header to pass to the GetDIBits call
    Dim bmi As BITMAPINFO
    bmi.bmHeader.bmWidth = bm.bmWidth
    bmi.bmHeader.bmHeight = bm.bmHeight
    bmi.bmHeader.bmSize = 40                'Size, in bytes, of the header
    bmi.bmHeader.bmPlanes = 1               'Number of planes (always one for this instance)
    bmi.bmHeader.bmBitCount = 24            'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.bmCompression = 0          'Compression :standard/none or RLE
    
    'Get the image data into our array
    If CorrectOrientation = False Then
        GetDIBits SrcPictureBox.hdc, SrcPictureBox.Image, 0, bm.bmHeight, ImageData(0, 0), bmi, 0
    Else
        GetDIBits SrcPictureBox.hdc, SrcPictureBox.Image, 0, bm.bmHeight, tmpData(0, 0), bmi, 0
    End If
    
    'This code is to orient the image data correctly in the array (i.e. (0,0) as top-left, (max,max) as bottom right)
    ' (if this option is enabled, we must set the DIB height to negative in the SetImageData routine below)
    If CorrectOrientation = True Then
    
        Dim x As Long, y As Long, z As Long
        Dim QuickVal As Long
        For x = 0 To bm.bmWidth - 1
            QuickVal = x * 3
         For y = 0 To bm.bmHeight - 1
          For z = 0 To 2
            ImageData(QuickVal + z, y) = tmpData(QuickVal + z, bm.bmHeight - y)
          Next z
         Next y
        Next x
        
    End If
    
    'Save memory...?
    Erase tmpData
End Sub

'Set an image's pixel information from an array dimensioned (x * 3 + bgr, y)
Public Sub SetImageData2D(DstPictureBox As PictureBox, OriginalWidth As Long, OriginalHeight As Long, ImageData() As Byte, Optional ByVal CorrectOrientation As Boolean = False)
    Dim bm As BITMAP
    'Get the picture box information
    GetObject DstPictureBox.Image, Len(bm), bm
    'Create a temporary header to pass to the StretchDIBits call
    Dim bmi As BITMAPINFO
    bmi.bmHeader.bmWidth = OriginalWidth
    If CorrectOrientation = False Then
        bmi.bmHeader.bmHeight = OriginalHeight
    Else
        bmi.bmHeader.bmHeight = -OriginalHeight
    End If
    bmi.bmHeader.bmSize = 40                'Size, in bytes, of the header
    bmi.bmHeader.bmPlanes = 1               'Number of planes (always one for this instance)
    bmi.bmHeader.bmBitCount = 24            'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.bmCompression = 0          'Compression :standard/none or RLE
    'Assume color images and set the corresponding best stretch mode
    SetStretchBltMode DstPictureBox.hdc, 3&
    'Send the array to the picture box and draw it accordingly
    StretchDIBits DstPictureBox.hdc, 0, 0, bm.bmWidth, bm.bmHeight, 0, 0, OriginalWidth, OriginalHeight, ImageData(0, 0), bmi, 0, vbSrcCopy
    'Since this doesn't automatically initialize AutoRedraw, we have to do it manually
    If DstPictureBox.AutoRedraw = True Then
        DstPictureBox.Picture = DstPictureBox.Image
        DstPictureBox.Refresh
    End If
    'Always good to manually halt for external processes after heavy API usage
    DoEvents
End Sub

Sub Fast_Adjust(hPictureBoxSource As PictureBox, hExportPictureBox As PictureBox, Optional AdjustAll As Single = 5, Optional AdjustRed As Single = 5, Optional AdjustGreen As Single = 5, Optional AdjustBlue As Single = 5, Optional LoopEffect As Single = 1)
Dim PicInfo As BITMAP
GetObject hPictureBoxSource.Image, Len(PicInfo), PicInfo
Dim ImageData() As Byte
Dim x As Long, y As Long
Dim z As Integer
Dim iWidth As Long, iHeight As Long
    iWidth = PicInfo.bmWidth
    iHeight = PicInfo.bmHeight
    
GetImageData2D hPictureBoxSource, ImageData()

Dim R As Long, G As Long, B As Long, tR As Long, tG As Long, tB As Long

    'Run a quick loop through the image, adjusting pixel values as we go
    Dim QuickX As Long
    For z = 1 To LoopEffect
    For x = 0 To iWidth - 1
        QuickX = x * 3
    For y = 0 To iHeight - 1
        R = ImageData(QuickX + 2, y)
        G = ImageData(QuickX + 1, y)
        B = ImageData(QuickX, y)
        tR = R * R * AdjustAll / 1000 + R * (255 - R) * AdjustRed * 0.1 / 255
        tG = G * G * AdjustAll / 1000 + G * (255 - G) * AdjustGreen * 0.1 / 255
        tB = B * B * AdjustAll / 1000 + B * (255 - B) * AdjustBlue * 0.1 / 255
        ByteMe tR
        ByteMe tG
        ByteMe tB
        ImageData(QuickX + 2, y) = tR
        ImageData(QuickX + 1, y) = tG
        ImageData(QuickX, y) = tB
    Next y
    Next x
    Next z

SetImageData2D hExportPictureBox, iWidth, iHeight, ImageData()
End Sub
