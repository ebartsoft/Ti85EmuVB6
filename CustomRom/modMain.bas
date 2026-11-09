Attribute VB_Name = "modMain"
Option Explicit

'============================================================
' Command-line entry point
'
' Usage:
'   CustomRomBuilder.exe <input.bmp> [output.rom]
'
' If output.rom is omitted, CustomTI85.rom is written next to
' the executable.
'============================================================
Public Sub Main()
    Dim args() As String
    Dim bmpPath As String
    Dim romPath As String
    Dim mono() As Byte
    Dim blackPixels As Long
    Dim errMsg As String

    args = SplitCommand(Command$)
    If Len(Trim$(args(0))) = 0 Then
        MsgBox "Usage:" & vbCrLf & vbCrLf & _
               "CustomRomBuilder.exe <input.bmp> [output.rom]", _
               vbInformation, "TI-85 Custom ROM Builder"
        Exit Sub
    End If

    bmpPath = args(0)
    If UBound(args) >= 1 Then
        romPath = args(1)
    Else
        romPath = App.Path & "\CustomTI85.rom"
    End If

    If Not modBmpToRom.BuildRomFromBmp(bmpPath, romPath, mono, blackPixels, errMsg) Then
        MsgBox "Could not create ROM:" & vbCrLf & vbCrLf & errMsg, _
               vbCritical, "TI-85 Custom ROM Builder"
        Exit Sub
    End If
End Sub

'============================================================
' Split the command line into quoted arguments.
'============================================================
Private Function SplitCommand(ByVal s As String) As String()
    Dim out() As String
    Dim count As Long
    Dim i As Long
    Dim ch As String
    Dim token As String
    Dim inQuote As Boolean

    ReDim out(0 To 7)
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = """" Then
            inQuote = Not inQuote
        ElseIf (ch = " " Or ch = vbTab) And Not inQuote Then
            If Len(token) > 0 Then
                If count > UBound(out) Then ReDim Preserve out(0 To count + 7)
                out(count) = token
                count = count + 1
                token = ""
            End If
        Else
            token = token & ch
        End If
        i = i + 1
    Loop

    If Len(token) > 0 Then
        If count > UBound(out) Then ReDim Preserve out(0 To count + 7)
        out(count) = token
        count = count + 1
    End If

    If count = 0 Then
        ReDim out(0 To 0)
        out(0) = ""
    Else
        ReDim Preserve out(0 To count - 1)
    End If
    SplitCommand = out
End Function
