Attribute VB_Name = "modTi85"
Option Explicit

'============================================================
' TI-85 hardware emulation
' Memory, keypad, LCD and port handling.
'============================================================

Public Const IFF_1          As Long = &H1&
Public Const IFF_IM1        As Long = &H2&
Public Const IFF_IM2        As Long = &H4&
Public Const IFF_2          As Long = &H8&
Public Const IFF_EI         As Long = &H20&
Public Const IFF_HALT       As Long = &H80&

Public Const INT_IRQ        As Long = &HFF&

Private Declare Sub MemoryCopy Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

Private ROM()               As Byte
Private CurrBank            As Long

Public Parity()             As Long
Public ShiftRight()         As Long
Public glMemAddrDiv256()    As Long
Public MEM()                As Byte

Public VidOffset            As Long
Public KeypadMask           As Long
Public DisplayContrast      As Long
Public Port3                As Byte
Public PowerReg             As Byte
Public WasInterrupt         As Long
Public Keyboard(7)          As Byte
Public OnPressed            As Boolean

Public Cycles As Long
Public SliceTicks As Long
Public TotalCycles As Double
'Public AutoOnDone As Boolean
'Public AutoOnTimer As Long

'================================================
' Full emulator state file
'================================================
Private Const TI85_STATE_MAGIC   As String = "TI85ST01"
Private Const TI85_STATE_VERSION As Long = 1

Private Type Ti85StateHeader
    Magic   As String * 8
    Version As Long
End Type

Private Type Ti85CpuState
    regA     As Long
    regHL    As Long
    regB     As Long
    regC     As Long
    regDE    As Long
    regAF_   As Long
    regHL_   As Long
    regBC_   As Long
    regDE_   As Long
    regIX    As Long
    regIY    As Long
    regID    As Long
    regSP    As Long
    regPC    As Long
    fC       As Long
    fN       As Long
    fPV      As Long
    f3       As Long
    fH       As Long
    f5       As Long
    fZ       As Long
    fS       As Long
    intI     As Long
    intR     As Long
    intRTemp As Long
    intIFF   As Long
    IBackup  As Long
    ICount   As Long
End Type

Private Type Ti85PortState
    VidOffset       As Long
    KeypadMask      As Long
    DisplayContrast As Long
    Port3           As Byte
    PowerReg        As Byte
    WasInterrupt    As Long
    Keyboard(0 To 7) As Byte
    OnPressed       As Boolean
    CurrBank        As Long
    Cycles          As Long
    SliceTicks      As Long
    TotalCycles     As Double
End Type

'================================================
' Ti85LoadRom
'================================================
Public Sub Ti85LoadRom(ByRef pFileName As String)
    Dim vFile   As Integer
    Dim vAttr   As Long
    Dim vSize   As Long
    Dim vData() As Byte
    
    ReDim ROM(&H1FFFF)
    
    vAttr = GetFileAttributes(pFileName)
    If (vAttr = -1) Or ((vAttr And &H10) = &H10) Then
        vData = LoadResData(101, "CUSTOM")
        MemoryCopy ROM(0), vData(0), UBound(vData) + 1
    Else
        vFile = FreeFile
        Open pFileName For Binary As vFile
        vSize = LOF(vFile)
        If vSize > 0 Then
            If vSize > &H20000 Then vSize = &H20000
            ReDim vData(vSize - 1)
            Get vFile, , vData
            MemoryCopy ROM(0), vData(0), vSize
        End If
        Close vFile
    End If
    
End Sub

'================================================
' Ti85MapRom
'================================================
Private Sub Ti85MapRom(ByVal pBank As Long)
    Dim i As Long
    pBank = pBank And 7
    If CurrBank <> pBank Then
        CurrBank = pBank
        For i = 0 To &H3FFF&
            MEM(&H4000& + i) = ROM(&H4000& * pBank + i)
        Next
    End If
End Sub

'================================================
' Ti85UpdateRamChecksum
'================================================
Public Sub Ti85UpdateRamChecksum()
    Dim i As Long
    Dim sum As Long

    ' Same checksum used by the TI-85 ROM:
    ' sum bytes from 8BF7h to FA6Fh, add 8, then two's complement.
    For i = &H8BF7& - &H8000& To &HFA6F& - &H8000&
        sum = (sum + MEM(&H8000& + i)) And &HFFFF&
    Next
    sum = (sum + 8) And &HFFFF&
    sum = (&H10000 - sum) And &HFFFF&
    MEM(&H81BE&) = sum And &HFF&
    MEM(&H81BF&) = (sum \ 256) And &HFF&
End Sub

''================================================
'' Ti85RamStateLoad
''================================================
'Public Sub Ti85RamStateLoad()
'    Dim f As Integer
'    Dim n As Long
'    Dim i As Long
'    Dim b() As Byte
'
'    If dir$(MainConfig.RamFilePath) = "" Then Exit Sub
'
'    On Error GoTo LoadFail
'    f = FreeFile
'    Open MainConfig.RamFilePath For Binary As #f
'    n = LOF(f)
'    If n <> 32768 Then
'        Close #f
'        Exit Sub
'    End If
'    ReDim b(0 To 32767)
'    Get #f, , b
'    Close #f
'    For i = 0 To 32767
'        MEM(&H8000& + i) = b(i)
'    Next
'    Ti85UpdateRamChecksum
'    'DisplayContrast = MEM(&H8007&) And &H1F&
'    Exit Sub
'
'LoadFail:
'    If f > 0 Then Close #f
'End Sub

''================================================
'' Ti85RamStateSave
''================================================
'Public Sub Ti85RamStateSave()
'    Dim f As Integer
'    Dim i As Long
'    Dim b() As Byte
'
'    On Error GoTo SaveFail
'    'MEM(&H8007&) = DisplayContrast And &H1F& 'Save DisplayContrast
'    Ti85UpdateRamChecksum
'    ReDim b(0 To 32767)
'    For i = 0 To 32767
'        b(i) = MEM(&H8000& + i)
'    Next
'    f = FreeFile
'    Open MainConfig.RamFilePath For Binary As #f
'    Put #f, , b
'    Close #f
'    Exit Sub
'
'SaveFail:
'    If f > 0 Then Close #f
'End Sub

'================================================
' Ti85StateSave
'
' Save the complete emulator state:
'   - all 64 KB of MEM, including screen/RAM
'   - all Z80 registers, flags and interrupt state
'   - all modTi85 port/keypad/timing state
'
' Call this while the CPU is between Execute calls, not while
' modCPU.Execute is running.
'================================================
Public Function Ti85StateSave(ByRef pFileName As String) As Boolean
    Dim f As Integer
    Dim h As Ti85StateHeader
    Dim c As Ti85CpuState
    Dim p As Ti85PortState
    Dim i As Long

    On Error GoTo Fail

    h.Magic = TI85_STATE_MAGIC
    h.Version = TI85_STATE_VERSION

    c.regA = regA
    c.regHL = regHL
    c.regB = regB
    c.regC = regC
    c.regDE = regDE
    c.regAF_ = regAF_
    c.regHL_ = regHL_
    c.regBC_ = regBC_
    c.regDE_ = regDE_
    c.regIX = regIX
    c.regIY = regIY
    c.regID = regID
    c.regSP = regSP
    c.regPC = regPC
    c.fC = fC
    c.fN = fN
    c.fPV = fPV
    c.f3 = f3
    c.fH = fH
    c.f5 = f5
    c.fZ = fZ
    c.fS = fS
    c.intI = intI
    c.intR = intR
    c.intRTemp = intRTemp
    c.intIFF = intIFF
    c.IBackup = IBackup
    c.ICount = ICount

    p.VidOffset = VidOffset
    p.KeypadMask = KeypadMask
    p.DisplayContrast = DisplayContrast
    p.Port3 = Port3
    p.PowerReg = PowerReg
    p.WasInterrupt = WasInterrupt
    For i = 0 To 7
        p.Keyboard(i) = Keyboard(i)
    Next i
    p.OnPressed = OnPressed
    p.CurrBank = CurrBank
    p.Cycles = Cycles
    p.SliceTicks = SliceTicks
    p.TotalCycles = TotalCycles

    f = FreeFile
    Open pFileName For Binary As #f
    Put #f, , h
    Put #f, , MEM
    Put #f, , c
    Put #f, , p
    Close #f

    Ti85StateSave = True
    Exit Function

Fail:
    On Error Resume Next
    Close #f
    Ti85StateSave = False
End Function

'================================================
' Ti85StateLoad
'
' Restore a state written by Ti85StateSave. The current state is
' only changed after the whole file has been read successfully.
'================================================
Public Function Ti85StateLoad(ByRef pFileName As String) As Boolean
    Dim f As Integer
    Dim h As Ti85StateHeader
    Dim c As Ti85CpuState
    Dim p As Ti85PortState
    Dim m() As Byte
    Dim i As Long

    On Error GoTo Fail

    f = FreeFile
    Open pFileName For Binary As #f

    Get #f, , h
    If Left$(h.Magic, 8) <> TI85_STATE_MAGIC Then GoTo Fail
    If h.Version <> TI85_STATE_VERSION Then GoTo Fail

    ReDim m(0 To 65535)
    Get #f, , m
    Get #f, , c
    Get #f, , p
    Close #f

    MEM = m

    regA = c.regA
    regHL = c.regHL
    regB = c.regB
    regC = c.regC
    regDE = c.regDE
    regAF_ = c.regAF_
    regHL_ = c.regHL_
    regBC_ = c.regBC_
    regDE_ = c.regDE_
    regIX = c.regIX
    regIY = c.regIY
    regID = c.regID
    regSP = c.regSP
    regPC = c.regPC
    fC = c.fC
    fN = c.fN
    fPV = c.fPV
    f3 = c.f3
    fH = c.fH
    f5 = c.f5
    fZ = c.fZ
    fS = c.fS
    intI = c.intI
    intR = c.intR
    intRTemp = c.intRTemp
    intIFF = c.intIFF
    IBackup = c.IBackup
    ICount = c.ICount

    VidOffset = p.VidOffset
    KeypadMask = p.KeypadMask
    DisplayContrast = p.DisplayContrast
    Port3 = p.Port3
    PowerReg = p.PowerReg
    WasInterrupt = p.WasInterrupt
    For i = 0 To 7
        Keyboard(i) = p.Keyboard(i)
    Next i
    OnPressed = p.OnPressed
    CurrBank = p.CurrBank
    Cycles = p.Cycles
    SliceTicks = p.SliceTicks
    TotalCycles = p.TotalCycles

    Ti85StateLoad = True
    Exit Function

Fail:
    On Error Resume Next
    Close #f
    Ti85StateLoad = False
End Function

'================================================
' Ti85IsDisplayOn
'================================================
Public Function Ti85IsDisplayOn() As Boolean
    Ti85IsDisplayOn = ((Port3 And &H8&) <> 0)
End Function

'================================================
' Ti85Reset
'================================================
Public Sub Ti85Reset()
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim z As Long

    'LoadRom RomFile
    
    ReDim MEM(65535)
    ReDim Parity(256)
    ReDim glMemAddrDiv256(65535)
    ReDim ShiftRight(600, 1 To 10)
    
    'For i = &H8000& To &HFFFF&
    '    MEM(i) = 0
    'Next
    
    For i = 0 To &H3FFF&
        MEM(i) = ROM(i)
        MEM(&H4000& + i) = ROM(&H4000& + i)
    Next

    For i = 0 To 65535
        glMemAddrDiv256(i) = i \ 256
    Next
    
    For i = 0 To 255
        y = 1
        For z = 0 To 7
            If (i And (2 ^ z)) <> 0 Then y = 1 - y
        Next
        Parity(i) = y
    Next
    
    For i = 1 To 599
        z = 2
        For x = 1 To 10
            ShiftRight(i, x) = i \ z
            z = z * 2
        Next
    Next

    VidOffset = 0
    KeypadMask = 0
    DisplayContrast = MEM(&H8007&) And &H1F& '18 '0
    Port3 = 0
    PowerReg = 0
    LinkReset
    CurrBank = 1
    WasInterrupt = 0
    OnPressed = False
    ReleaseAllKeys
    Cycles = 0
    SliceTicks = 0
    TotalCycles = 0
    'AutoOnDone = False
    'AutoOnTimer = 0

    Z80Reset
    regPC = 0
    regSP = &HF000&
End Sub

'================================================
' Ti85Tick
'================================================
Public Sub Ti85Tick()
    TotalCycles = TotalCycles + 30000

    'If AutoOnTimer > 0 Then
    '    AutoOnTimer = AutoOnTimer - 1
    '    If AutoOnTimer = 0 Then ReleaseOn
    'End If

    'If (Not AutoOnDone) And ((Port3 And 1) <> 0) And (TotalCycles > 40000000) Then
    '    AutoOnDone = True
    '    PressOn
    '    AutoOnTimer = 30
    'End If

    If (Port3 And &H2&) <> 0 Then WasInterrupt = WasInterrupt Or 4
End Sub

Public Function Shl(ByVal v As Long, ByVal n As Long) As Long
    Shl = (v * (2& ^ n)) And &HFFFFFFFF
End Function

Public Function PeekB(ByVal Addr As Long) As Byte
    PeekB = MEM(Addr And &HFFFF&)
End Function

Public Sub PokeB(ByVal Addr As Long, ByVal Value As Byte)
    If Addr >= &H8000& Then MEM(Addr) = Value
End Sub

Public Function inn(ByVal Port As Long) As Byte
    Dim j As Long
    Port = Port And &HFF&
    Select Case Port
    Case 1
        j = &HFF&
        If (KeypadMask And 1) = 0 Then j = j And Keyboard(0)
        If (KeypadMask And 2) = 0 Then j = j And Keyboard(1)
        If (KeypadMask And 4) = 0 Then j = j And Keyboard(2)
        If (KeypadMask And 8) = 0 Then j = j And Keyboard(3)
        If (KeypadMask And 16) = 0 Then j = j And Keyboard(4)
        If (KeypadMask And 32) = 0 Then j = j And Keyboard(5)
        If (KeypadMask And 64) = 0 Then j = j And Keyboard(6)
        inn = j And &HFF&
    Case 2
        inn = DisplayContrast And &H1F&
    Case 3
        j = WasInterrupt
        WasInterrupt = 0
        If j = 1 Then Port3 = Port3 Or &H4&
        If OnPressed Then
            ' ON key is being held
        Else
            j = j Or &H8&
        End If
        If (Port3 And &H8&) <> 0 Then j = j Or (Port3 And &H2&)
        inn = j And &HFF&
    Case 5
        inn = CurrBank
    Case 6
        inn = PowerReg
    Case 7
        inn = LinkReadPort()
    Case Else
        inn = &HFF&
    End Select
End Function

Public Sub out(ByVal Port As Long, ByVal Value As Byte)
    Dim v As Long
    Port = Port And &HFF&
    v = Value And &HFF&
    Select Case Port
    Case 0
        VidOffset = ((v And &H3F&) + &HC0&) * 256&
    Case 1
        KeypadMask = v
    Case 2
        DisplayContrast = v
    Case 3
        Port3 = v
    Case 4
    Case 5
        Ti85MapRom v
    Case 6
        PowerReg = v
    Case 7
        LinkWritePort v
    End Select
End Sub

Public Sub KeyDown(ByVal Row As Long, ByVal Col As Long)
    If Row < 0 Or Row > 7 Then Exit Sub
    Keyboard(Row) = Keyboard(Row) And Not (Shl(1, Col) And &HFF&)
End Sub

Public Sub KeyUp(ByVal Row As Long, ByVal Col As Long)
    If Row < 0 Or Row > 7 Then Exit Sub
    Keyboard(Row) = Keyboard(Row) Or (Shl(1, Col) And &HFF&)
End Sub

Public Sub ReleaseAllKeys()
    Dim i As Long
    For i = 0 To 7
        Keyboard(i) = &HFF&
    Next
    OnPressed = False
End Sub

Public Sub PressOn()
    If OnPressed Then Exit Sub
    OnPressed = True
    'AutoOnDone = True
    If (Port3 And 1) <> 0 Then WasInterrupt = 1
End Sub

Public Sub ReleaseOn()
    OnPressed = False
End Sub

'================================================
' Ti85Interrupt
'================================================
Public Sub Ti85Interrupt(ByVal Vector As Long)
    Dim w As Long
    
    If intIFF And IFF_HALT Then
        regPC = regPC + 1
        intIFF = intIFF And Not IFF_HALT
    End If
    
    If (intIFF And IFF_1) Or Vector = &HFFFD& Then
    
        pushpc
        If Vector = &HFFFD& Then
            intIFF = intIFF And Not (IFF_1 Or IFF_EI)
            regPC = &H66&
            Exit Sub
        End If
        
        intIFF = intIFF And Not (IFF_1 Or IFF_2 Or IFF_EI)
        If intIFF And IFF_IM2 Then
            w = (Vector And &HFF&) Or Shl(intI, 8)
            regPC = peekw(w + 1)
            Exit Sub
        End If
        
        If intIFF And IFF_IM1 Then
            regPC = &H38&
            Exit Sub
        End If
        
        Select Case Vector
        Case &H0&:  regPC = &H0&
        Case &H8&:  regPC = &H8&
        Case &H10&: regPC = &H10&
        Case &H18&: regPC = &H18&
        Case &H20&: regPC = &H20&
        Case &H28&: regPC = &H28&
        Case &H30&: regPC = &H30&
        Case &HFF&: regPC = &H38&
        End Select
        
    End If
    
End Sub
