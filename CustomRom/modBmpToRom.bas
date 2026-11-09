Attribute VB_Name = "modBmpToRom"
Option Explicit

'============================================================
' BMP -> TI-85 monochrome ROM builder
'
' The generated ROM contains a fixed boot/ON-OFF stub followed by
' 1024 bytes of 1-bpp screen data (128 x 64, 16 bytes per row,
' bit 7 = leftmost pixel).
'============================================================

Private Const SCREEN_W As Long = 128
Private Const SCREEN_H As Long = 64
Private Const SCREEN_BYTES As Long = 1024
Private Const DARK_LEVEL As Long = 160

' Fixed bytecode for the reference program in CustomTI85.asm.
' The IMAGE address operand is patched to the actual image offset
' when the ROM is generated.
Private Const STUB_HEX As String = _
    "F33100F03E3CD3003E12D3023207803E08D3033E01320080" & _
    "CD4A00CD4300CD3C00CD43003A0080EE0132008028093E08" & _
    "D303CD4A0018E4AFD30318DFDB03E60820FAC9DB03E60828" & _
    "FAC92156001100FC010004EDB0C9"

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

'============================================================
' BuildRomFromBmp
'============================================================
Public Function BuildRomFromBmp( _
    ByVal BmpPath As String, _
    ByVal RomPath As String, _
    ByRef Mono() As Byte, _
    ByRef BlackPixels As Long, _
    ByRef ErrMsg As String) As Boolean

    Dim stub() As Byte
    Dim rom() As Byte
    Dim base As Long
    Dim i As Long

    On Error GoTo Fail

    If Not BmpToMono(BmpPath, Mono, ErrMsg) Then Exit Function

    stub = HexToBytes(STUB_HEX)
    base = UBound(stub) + 1

    ' The LD HL,IMAGE operand sits at offset $004B in the stub.
    If base <> &H56& Then
        ErrMsg = "Internal error: stub layout changed."
        Exit Function
    End If
    stub(&H4B&) = CByte(base And &HFF&)
    stub(&H4C&) = CByte((base \ 256&) And &HFF&)

    ReDim rom(0 To base + SCREEN_BYTES - 1)
    For i = 0 To UBound(stub)
        rom(i) = stub(i)
    Next i

    BlackPixels = 0
    For i = 0 To SCREEN_BYTES - 1
        rom(base + i) = Mono(i)
        BlackPixels = BlackPixels + CountBits(Mono(i))
    Next i

    SaveRom RomPath, rom
    BuildRomFromBmp = True
    Exit Function

Fail:
    ErrMsg = "Error: " & Err.Description
End Function

'============================================================
' BmpToMono
'
' Supports uncompressed 1, 4, 8, 24 and 32 bpp BMP files.
' Images larger than 128 x 64 are scaled down; smaller images
' are centred at their original size.
'============================================================
Public Function BmpToMono(ByVal BmpPath As String, ByRef Mono() As Byte, ByRef ErrMsg As String) As Boolean

    Dim data() As Byte
    Dim pal() As Long
    Dim img() As Byte

    Dim w As Long
    Dim h As Long
    Dim bpp As Long
    Dim planes As Long
    Dim compression As Long
    Dim colors As Long
    Dim pixelOffset As Long
    Dim headerSize As Long
    Dim stride As Long
    Dim topDown As Boolean

    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim sx As Long
    Dim sy As Long
    Dim destW As Long
    Dim destH As Long
    Dim offX As Long
    Dim offY As Long
    Dim rowOfs As Long
    Dim zoom As Single

    Dim p As Long
    Dim r As Long
    Dim g As Long
    Dim bl As Long

    On Error GoTo Fail

    If Not ReadAllBytes(BmpPath, data) Then
        ErrMsg = "Cannot read file: " & BmpPath
        Exit Function
    End If

    If UBound(data) < 54 Then
        ErrMsg = "Not a valid BMP file."
        Exit Function
    End If
    If data(0) <> 66 Or data(1) <> 77 Then
        ErrMsg = "Not a BMP file."
        Exit Function
    End If

    pixelOffset = GetLong(data, 10)
    headerSize = GetLong(data, 14)
    If headerSize < 40 Or headerSize > UBound(data) - 13 Then
        ErrMsg = "Unsupported BMP header."
        Exit Function
    End If

    w = GetLong(data, 18)
    h = GetLong(data, 22)
    planes = GetInt(data, 26)
    bpp = GetInt(data, 28)
    compression = GetLong(data, 30)

    If w <= 0 Or h = 0 Then
        ErrMsg = "Invalid BMP dimensions."
        Exit Function
    End If
    If planes <> 1 Then
        ErrMsg = "Invalid BMP plane count."
        Exit Function
    End If
    If compression <> 0 Then
        ErrMsg = "Compressed BMP files are not supported."
        Exit Function
    End If

    topDown = (h < 0)
    If topDown Then h = -h
    If h <= 0 Then
        ErrMsg = "Invalid BMP height."
        Exit Function
    End If

    Select Case bpp
    Case 1
        stride = ((w + 31&) \ 32&) * 4&
    Case 4
        stride = ((w * 4& + 31&) \ 32&) * 4&
    Case 8
        stride = ((w + 3&) \ 4&) * 4&
    Case 24
        stride = ((w * 3& + 3&) \ 4&) * 4&
    Case 32
        stride = w * 4&
    Case Else
        ErrMsg = "Unsupported BMP colour depth: " & bpp & " bpp."
        Exit Function
    End Select

    If pixelOffset < 14 + headerSize Then
        ErrMsg = "Invalid BMP pixel offset."
        Exit Function
    End If
    If pixelOffset + stride * h > UBound(data) + 1 Then
        ErrMsg = "Truncated BMP pixel data."
        Exit Function
    End If

    If bpp <= 8 Then
        colors = GetLong(data, 46)
        If colors <= 0 Or colors > 256 Then colors = CLng(2 ^ bpp)
        ReDim pal(0 To colors - 1)
        p = 14 + headerSize
        If p + colors * 4& > UBound(data) + 1 Then
            ErrMsg = "Truncated BMP colour table."
            Exit Function
        End If
        For i = 0 To colors - 1
            bl = data(p + i * 4&)
            g = data(p + i * 4& + 1)
            r = data(p + i * 4& + 2)
            pal(i) = r Or (g * 256&) Or (bl * 65536&)
        Next i
    Else
        ReDim pal(0 To 0)
    End If

    ReDim img(0 To SCREEN_BYTES - 1)

    ' Fit larger images into the LCD; keep smaller images at 1:1.
    zoom = 1
    If w > SCREEN_W Or h > SCREEN_H Then
        zoom = CSng(SCREEN_W) / CSng(w)
        If CSng(SCREEN_H) / CSng(h) < zoom Then
            zoom = CSng(SCREEN_H) / CSng(h)
        End If
    End If

    destW = CLng(Int(CSng(w) * zoom + 0.5))
    destH = CLng(Int(CSng(h) * zoom + 0.5))
    If destW < 1 Then destW = 1
    If destH < 1 Then destH = 1
    If destW > SCREEN_W Then destW = SCREEN_W
    If destH > SCREEN_H Then destH = SCREEN_H
    offX = (SCREEN_W - destW) \ 2
    offY = (SCREEN_H - destH) \ 2

    For y = 0 To SCREEN_H - 1
        If y >= offY And y < offY + destH Then
            For x = 0 To SCREEN_W - 1
                If x >= offX And x < offX + destW Then
                    sx = ((x - offX) * w) \ destW
                    sy = ((y - offY) * h) \ destH
                    If sx >= w Then sx = w - 1
                    If sy >= h Then sy = h - 1

                    If topDown Then
                        rowOfs = pixelOffset + sy * stride
                    Else
                        rowOfs = pixelOffset + (h - 1 - sy) * stride
                    End If

                    If PixelIsDark(data, rowOfs, sx, bpp, pal) Then
                        i = y * 16 + x \ 8
                        img(i) = img(i) Or BitMask(x And 7)
                    End If
                End If
            Next x
        End If
    Next y

    Mono = img
    BmpToMono = True
    Exit Function

Fail:
    If ErrMsg = "" Then ErrMsg = "Error: " & Err.Description
End Function

'============================================================
' PixelIsDark
'============================================================
Private Function PixelIsDark( _
    ByRef data() As Byte, _
    ByVal rowOfs As Long, _
    ByVal x As Long, _
    ByVal bpp As Long, _
    ByRef pal() As Long) As Boolean

    Dim p As Long
    Dim idx As Long
    Dim color As Long
    Dim r As Long
    Dim g As Long
    Dim bl As Long

    Select Case bpp
    Case 1
        p = rowOfs + (x \ 8)
        idx = 0
        If (data(p) And BitMask(7 - (x And 7))) <> 0 Then idx = 1
        color = pal(idx)
        r = color And &HFF&
        g = (color \ 256&) And &HFF&
        bl = (color \ 65536&) And &HFF&
    Case 4
        p = rowOfs + (x \ 2)
        If (x And 1) = 0 Then
            idx = (data(p) \ 16) And &HF&
        Else
            idx = data(p) And &HF&
        End If
        color = pal(idx)
        r = color And &HFF&
        g = (color \ 256&) And &HFF&
        bl = (color \ 65536&) And &HFF&
    Case 8
        p = rowOfs + x
        color = pal(data(p))
        r = color And &HFF&
        g = (color \ 256&) And &HFF&
        bl = (color \ 65536&) And &HFF&
    Case 24
        p = rowOfs + x * 3&
        bl = data(p)
        g = data(p + 1)
        r = data(p + 2)
    Case 32
        p = rowOfs + x * 4&
        bl = data(p)
        g = data(p + 1)
        r = data(p + 2)
    End Select

    PixelIsDark = ((r * 299& + g * 587& + bl * 114&) \ 1000&) < DARK_LEVEL
End Function

'============================================================
' Helpers
'============================================================
Private Function BitMask(ByVal n As Long) As Byte
    Select Case n
    Case 0: BitMask = 128
    Case 1: BitMask = 64
    Case 2: BitMask = 32
    Case 3: BitMask = 16
    Case 4: BitMask = 8
    Case 5: BitMask = 4
    Case 6: BitMask = 2
    Case 7: BitMask = 1
    End Select
End Function

Private Function CountBits(ByVal v As Byte) As Long
    Dim n As Long
    n = v
    Do While n <> 0
        CountBits = CountBits + (n And 1)
        n = n \ 2
    Loop
End Function

Private Function HexToBytes(ByVal s As String) As Byte()
    Dim b() As Byte
    Dim n As Long
    Dim i As Long
    n = Len(s) \ 2
    ReDim b(0 To n - 1)
    For i = 0 To n - 1
        b(i) = CByte(CLng("&H" & Mid$(s, i * 2 + 1, 2)) And &HFF&)
    Next i
    HexToBytes = b
End Function

Private Function ReadAllBytes(ByVal Path As String, ByRef Data() As Byte) As Boolean
    Dim f As Integer
    Dim n As Long
    On Error GoTo Fail

    f = FreeFile
    Open Path For Binary As #f
    n = LOF(f)
    If n <= 0 Then
        Close #f
        Exit Function
    End If
    ReDim Data(0 To n - 1)
    Get #f, , Data
    Close #f
    ReadAllBytes = True
    Exit Function

Fail:
    On Error Resume Next
    Close #f
End Function

Private Sub SaveRom(ByVal Path As String, ByRef Data() As Byte)
    Dim f As Integer
    On Error Resume Next
    Kill Path
    On Error GoTo 0
    f = FreeFile
    Open Path For Binary As #f
    Put #f, , Data
    Close #f
End Sub

Private Function GetLong(ByRef b() As Byte, ByVal off As Long) As Long
    Dim v As Long
    CopyMemory v, b(off), 4
    GetLong = v
End Function

Private Function GetInt(ByRef b() As Byte, ByVal off As Long) As Long
    Dim v As Integer
    CopyMemory v, b(off), 2
    GetInt = CLng(v)
End Function
