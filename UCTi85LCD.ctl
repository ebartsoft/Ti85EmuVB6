VERSION 5.00
Begin VB.UserControl UCTi85LCD 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   ClientHeight    =   2145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3315
   ScaleHeight     =   143
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   221
End
Attribute VB_Name = "UCTi85LCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const LCDW              As Long = 128
Private Const LCDH              As Long = 64
Private Const LCDMARGIN         As Long = 8

Private Const BI_RGB            As Long = 0
Private Const DIB_RGB_COLORS    As Long = 0
Private Const COLORONCOLOR      As Long = 3

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type BITMAPINFO
    bmiHeader                   As BITMAPINFOHEADER
    bmiColors(0 To 2)           As Long
End Type

Private Type SAFEARRAYBOUND
    cElements                   As Long
    lLbound                     As Long
End Type

Private Type SAFEARRAY2D
    cDims                       As Integer
    fFeatures                   As Integer
    cbElements                  As Long
    cLocks                      As Long
    pvData                      As Long
    Bounds(1)                   As SAFEARRAYBOUND
End Type

Private Declare Sub MemoryCopy Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "MSVBVM60.DLL" Alias "VarPtr" (ByRef Ptr() As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal xDest As Long, ByVal yDest As Long, ByVal dxDest As Long, ByVal dyDest As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dxSrc As Long, ByVal dySrc As Long, ByVal lpBits As Long, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal hSection As Long, ByVal dwOffset As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal iMode As Long) As Long

Private mBits1D         As SAFEARRAY2D
Private mScreenBmi      As BITMAPINFO
Private mDirty          As Boolean
Private mBitMask()      As Long
Private mScreenBmp      As Long
Private mScreenBits     As Long
Private mLastContrast   As Long
Private mBits1Dsa       As Long
Private mBffWidth       As Long
Private mBffHeight      As Long
Private mPrev()         As Byte

'================================================
' UserControl_Initialize
'================================================
Private Sub UserControl_Initialize()
    InitScreenBuffer
End Sub

'================================================
' UserControl_Terminate
'================================================
Private Sub UserControl_Terminate()
    FreeScreenBuffer
End Sub

'================================================
' Invalidate
'================================================
Public Sub Invalidate()
    mDirty = True
End Sub

'================================================
' BitsLockLong1D
'================================================
Private Sub BitsLockLong1D(ByRef pData() As Long)
    MemoryCopy ByVal ArrPtr(pData()), mBits1Dsa, 4
End Sub

'================================================
' BitsUnlockLong1D
'================================================
Private Sub BitsUnlockLong1D(ByRef pData() As Long)
    MemoryCopy ByVal ArrPtr(pData()), 0&, 4
End Sub

'================================================
' TakeScreenshot
'================================================
Public Sub TakeScreenshot(ByRef pFileName As String)
    Dim mPixels() As Long
    BitsLockLong1D mPixels
    Open pFileName For Output As 1: Close 1                 'Clear existing file
    Open pFileName For Binary As 1
    Put 1, , CInt(&H4D42)                                   'bfType
    Put 1, , CLng(54 + mScreenBmi.bmiHeader.biSizeImage)    'bfSize
    Put 1, , CLng(0)                                        'bfReserved
    Put 1, , CLng(54)                                       'bfOffBits
    Put 1, , mScreenBmi.bmiHeader
    Put 1, , mPixels
    Close 1
    BitsUnlockLong1D mPixels
End Sub

'================================================
' InitScreenBuffer
'================================================
Private Sub InitScreenBuffer()
    Dim i As Long
    
    ReDim mPrev(LCDW - 1, LCDH - 1)
    mBffHeight = LCDH + LCDMARGIN
    mBffWidth = LCDW + LCDMARGIN
    mLastContrast = -1
    mDirty = True
    
    ReDim mBitMask(7)
    For i = 0 To 7
        mBitMask(i) = 2 ^ i
    Next

    With mScreenBmi.bmiHeader
        .biSize = Len(mScreenBmi.bmiHeader)
        .biWidth = mBffWidth
        .biHeight = -mBffHeight
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
        .biSizeImage = mBffWidth * mBffHeight * 4
    End With

    mScreenBmp = CreateDIBSection(hDC, mScreenBmi, DIB_RGB_COLORS, mScreenBits, 0, 0)
    
    With mBits1D
        .cDims = 2
        .cbElements = 4
        .pvData = mScreenBits
        .Bounds(0).cElements = mBffHeight
        .Bounds(1).cElements = mBffWidth
    End With
    mBits1Dsa = VarPtr(mBits1D)

End Sub

'================================================
' FreeScreenBuffer
'================================================
Private Sub FreeScreenBuffer()
    If mScreenBmp <> 0 Then
        DeleteObject mScreenBmp
    End If
    mScreenBmp = 0
    mScreenBits = 0
End Sub

'================================================
' FillBorders
'================================================
Private Sub FillBorders(ByRef mPixels() As Long, ByVal pColor As Long)
    Dim pX As Long
    Dim pY As Long
    
    For pX = 0 To mBffWidth - 1
        For pY = 0 To 3
            mPixels(pX, pY) = pColor
        Next
        For pY = LCDH + 4 To LCDH + 8 - 1
            mPixels(pX, pY) = pColor
        Next
    Next
    
    For pY = 0 To mBffHeight - 1
        For pX = 0 To 3
            mPixels(pX, pY) = pColor
        Next
        For pX = LCDW + 4 To LCDW + 8 - 1
            mPixels(pX, pY) = pColor
        Next
    Next
    
End Sub

'================================================
' FxBackColor
'================================================
Private Function FxBackColor(ByVal x As Single) As Single
    FxBackColor = x * x * 1.1 - 0.17
End Function

'================================================
' FxForeColor
'================================================
Private Function FxForeColor(ByVal x As Single) As Single
    FxForeColor = x * x * 5 + 0.05
End Function

'================================================
' LerpColor
'================================================
Private Function LerpColor(ByVal pFrom As Long, ByVal pTo As Long, ByVal pFactor As Single) As Long
    Dim a(3) As Byte
    Dim b(3) As Byte
    Dim c(3) As Byte
    If pFactor < 0 Then pFactor = 0
    If pFactor > 1 Then pFactor = 1
    MemoryCopy a(0), pFrom, 4
    MemoryCopy b(0), pTo, 4
    c(2) = Int(a(0) + (0& + b(0) - a(0)) * pFactor)
    c(1) = Int(a(1) + (0& + b(1) - a(1)) * pFactor)
    c(0) = Int(a(2) + (0& + b(2) - a(2)) * pFactor)
    MemoryCopy LerpColor, c(0), 4
End Function

'================================================
' Redraw
'================================================
Public Sub Redraw()
    Dim pX As Long
    Dim pY As Long
    Dim byteX As Long
    Dim bitNo As Long
    Dim Addr As Long
    Dim v As Long
    Dim isOn As Boolean
    Dim onColor  As Long
    Dim offColor As Long
    Dim bdColor  As Long
    Dim t As Single
    Dim changed As Boolean
    
    Dim vBrightness As Single
    Dim vPixels()   As Long
    Dim vContrast   As Long

    'vContrast = MEM(&H8007&) And &H1F&
    vContrast = DisplayContrast
    If vContrast <> mLastContrast Then
        mLastContrast = vContrast
        mDirty = True
    End If

    vBrightness = vContrast / 31
    offColor = LerpColor(MainTheme.BkgColorOff, MainTheme.BkgColorOn, FxBackColor(vBrightness))
    onColor = LerpColor(MainTheme.FntColorOff, MainTheme.FntColorOn, FxForeColor(vBrightness))
    bdColor = MainTheme.BdrColor

    BitsLockLong1D vPixels
    FillBorders vPixels, LerpColor(bdColor, 0, 0)

    changed = mDirty
    Addr = VidOffset And &HFFFF&
    
    If Not Ti85IsDisplayOn() Then
        For pY = 0 To LCDH - 1
            For pX = 0 To LCDW - 1
                If mDirty Or mPrev(pX, pY) <> 0 Then
                    mPrev(pX, pY) = 0
                    vPixels(pX + 4, pY + 4) = offColor
                    changed = True
                End If
            Next
        Next
    Else
        For pY = 0 To LCDH - 1
            For byteX = 0 To 15
                v = MEM(Addr + pY * 16 + byteX)
                For bitNo = 7 To 0 Step -1
                    pX = byteX * 8 + (7 - bitNo)
                    isOn = (v And mBitMask(bitNo)) <> 0
                    If mDirty Or mPrev(pX, pY) <> Abs(isOn) Then
                        If isOn Then
                            mPrev(pX, pY) = 1
                            vPixels(pX + 4, pY + 4) = onColor
                        Else
                            mPrev(pX, pY) = 0
                            vPixels(pX + 4, pY + 4) = offColor
                        End If
                        changed = True
                    End If
                Next
            Next
        Next
    End If
    
    BitsUnlockLong1D vPixels

    If changed And (mScreenBits <> 0) Then
        SetStretchBltMode hDC, COLORONCOLOR
        StretchDIBits hDC, 0, 0, ScaleWidth, ScaleHeight, 0, 0, mBffWidth, mBffHeight, mScreenBits, mScreenBmi, 0, vbSrcCopy
    End If

    mDirty = False
End Sub
