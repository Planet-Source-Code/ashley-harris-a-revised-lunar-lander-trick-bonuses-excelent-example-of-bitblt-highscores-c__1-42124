Attribute VB_Name = "ashleysbitblt"
'the magic of fast graphics, part of gdi32.dll that comes with win95
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'super fast timer used to put a reasonable cap on frame rates (30fps)
Public Declare Function GetTickCount Lib "kernel32" () As Long

'creates a null picture in memory
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

'creates a picture device in memory compatible with the desktop
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

'converts an objects hwnd (Handle of Window) to it's Hdc
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

'links hdcs to pictures created in memory
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

'removes pictures from memory
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'removes devices from memory
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

'returns the colour value of one dot in a device (millilon times faster then vb's 'point'
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

'sets a single pixel to the specified coulour in a device (million times faster then vb's 'pset
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

'the pointer to the pointer to the backbuffer
Public backbuffer As Long

'used to tell the status of keys
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'loads a graphic into memory
Public Function LoadGraphicDC(sFileName As String) As Long
    On Error Resume Next
    Dim LoadGraphicDCTEMP As Long
    LoadGraphicDCTEMP = CreateCompatibleDC(GetDC(0))
    SelectObject LoadGraphicDCTEMP, LoadPicture(sFileName)
    lib(fso.GetFileName(sFileName)) = LoadGraphicDCTEMP
    LoadGraphicDC = LoadGraphicDCTEMP
End Function

'create a new backbuffer device in memory
Public Sub setupbackbuffer()
    mybackbuffer = CreateCompatibleDC(GetDC(0))
    myBufferBMP = CreateCompatibleBitmap(GetDC(0), 640, 640)
    SelectObject mybackbuffer, myBufferBMP
    BitBlt mybackbuffer, 0, 0, 640, 640, 0, 0, 0, vbWhiteness
    backbuffer = mybackbuffer
End Sub

