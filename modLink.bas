Attribute VB_Name = "modLink"
Option Explicit

'============================================================
' TI-85 link emulation (PC <-> emulated calculator)
'
' The TI-85 ROM itself implements the link packet protocol.
' This module emulates the two open-collector link wires and
' speaks the PC side of the manual-transfer protocol so that
' .85b, .85s and .85p files can be sent to and received from
' the emulated calculator.
'
' The byte engine is deliberately non-blocking: it is driven
' by the calculator's writes to port 7, not by re-entering the
' Z80 core. This fits the single-threaded VB6 main loop.
'============================================================

' Set to False to require "Link -> Receive File..." before the
' calculator may send data to the PC.
Private Const LINK_AUTO_RECEIVE As Boolean = True

Private Const LINK_TIMEOUT_MS As Long = 5000
Private Const LINK_ARM_TIMEOUT_MS As Long = 30000
Private Const LINK_BACKUP_EOT_GRACE_MS As Long = 1000
Private Const LINK_QCAP As Long = 131072     ' must be a power of two

' Machine ID used when this emulator sends TI-85 data.
Private Const LINK_ID_TI85 As Byte = &H85

' Packet command bytes
Private Const LCMD_VAR  As Byte = &H6
Private Const LCMD_CTS  As Byte = &H9
Private Const LCMD_DATA As Byte = &H15
Private Const LCMD_SKIP As Byte = &H36
Private Const LCMD_ACK  As Byte = &H56
Private Const LCMD_ERR  As Byte = &H5A
Private Const LCMD_EOT  As Byte = &H92

' Session types
Private Const LS_NONE As Long = 0
Private Const LS_SEND As Long = 1
Private Const LS_RECV As Long = 2

' Send states
Private Const SS_WAIT_ACK_CTS As Long = 1
Private Const SS_WAIT_DATA_ACK As Long = 2
Private Const SS_NEXT As Long = 3
Private Const SS_WAIT_EOT_ACK As Long = 4
Private Const SS_BACKUP_ACK1 As Long = 5
Private Const SS_BACKUP_ACK2 As Long = 6
Private Const SS_BACKUP_ACK3 As Long = 7

' Receive states
Private Const RS_WAIT_VAR As Long = 1
Private Const RS_WAIT_ACK As Long = 2
Private Const RS_WAIT_DATA As Long = 3
Private Const RS_WAIT_NEXT As Long = 4
Private Const RS_WAIT_BACKUP1 As Long = 5
Private Const RS_WAIT_BACKUP2 As Long = 6
Private Const RS_WAIT_BACKUP3 As Long = 7
Private Const RS_WAIT_BACKUP_EOT As Long = 8

' TX bit phases
Private Const TXP_WAIT_ACK As Long = 1
Private Const TXP_WAIT_CLEAR As Long = 2

' RX bit phases
Private Const RXP_WAIT_START As Long = 1
Private Const RXP_WAIT_CLEAR As Long = 2

' Wire state.  Out = driven by the calculator, In = driven by us.
Public LinkOutRed As Long
Public LinkOutWhite As Long
Public LinkInRed As Long
Public LinkInWhite As Long

' Byte queues
Private mTxBuf(0 To LINK_QCAP - 1) As Byte
Private mTxHead As Long, mTxTail As Long, mTxCount As Long
Private mRxBuf(0 To LINK_QCAP - 1) As Byte
Private mRxHead As Long, mRxTail As Long, mRxCount As Long

' TX engine
Private mTxActive As Boolean
Private mTxByte As Long
Private mTxBit As Long
Private mTxBitVal As Long
Private mTxPhase As Long

' RX engine
Private mRxActive As Boolean
Private mRxPhase As Long
Private mRxByte As Long
Private mRxBit As Long
Private mRxBitVal As Long

' Session
Private mSession As Long
Private mSessionActive As Boolean
Private mLastActivity As Long
Private mPeerID As Byte
Private mCtsHandled As Boolean
Private mRecvGotData As Boolean
Private mBackupGraceUntil As Long

' Send-side file data
Private Type LinkEntry
    dataLen As Long
    VarType As Byte
    NameLen As Long
    Name As String * 8
    Data As Variant
End Type

Private mSendEntries() As LinkEntry
Private mSendCount As Long
Private mSendIndex As Long
Private mSendState As Long
Private mSendIsBackup As Boolean
Private mSendFileName As String
Private mBackupSize1 As Long, mBackupSize2 As Long, mBackupSize3 As Long
Private mBackupAddr As Long
Private mBackupData1() As Byte
Private mBackupData2() As Byte
Private mBackupData3() As Byte

' Receive-side file data
Private mRecvEntries() As LinkEntry
Private mRecvCount As Long
Private mRecvState As Long
Private mRecvIsBackup As Boolean
Private mRecvCurEntry As Long
Private mRecvBackupSize1 As Long, mRecvBackupSize2 As Long, mRecvBackupSize3 As Long
Private mRecvBackupAddr As Long
Private mRecvBackup1() As Byte
Private mRecvBackup2() As Byte
Private mRecvBackup3() As Byte

' Public status / notification used by frmMain
Public LinkStatusText As String
Public LinkNotifyText As String
Public LinkNotifyError As Boolean
Public LinkNotifyPending As Boolean
Public LinkLastError As String

Private Declare Function GetTickCount Lib "kernel32" () As Long

'============================================================
' Public interface
'============================================================
Public Function LinkIsBusy() As Boolean
    LinkIsBusy = mSessionActive
End Function

Public Sub LinkReset()
    LinkOutRed = 0
    LinkOutWhite = 0
    LinkInRed = 1
    LinkInWhite = 1

    mTxHead = 0: mTxTail = 0: mTxCount = 0
    mRxHead = 0: mRxTail = 0: mRxCount = 0

    mTxActive = False
    mTxPhase = 0
    mTxByte = 0
    mTxBit = 0
    mTxBitVal = 0

    mRxActive = True
    mRxPhase = RXP_WAIT_START
    mRxByte = 0
    mRxBit = 0
    mRxBitVal = 0

    mSession = LS_NONE
    mSessionActive = False
    mLastActivity = GetTickCount
    mPeerID = 0
    mCtsHandled = False
    mRecvGotData = False
    mBackupGraceUntil = 0

    mSendCount = 0
    mSendIndex = 0
    mSendState = 0
    mSendIsBackup = False
    mSendFileName = ""

    mRecvCount = 0
    mRecvState = 0
    mRecvIsBackup = False
    mRecvCurEntry = 0

    Erase mSendEntries
    Erase mRecvEntries
    Erase mBackupData1
    Erase mBackupData2
    Erase mBackupData3
    Erase mRecvBackup1
    Erase mRecvBackup2
    Erase mRecvBackup3

    LinkStatusText = ""
    LinkNotifyText = ""
    LinkNotifyError = False
    LinkNotifyPending = False
    LinkLastError = ""
End Sub

Public Function LinkReadPort() As Long
    Dim v As Long

    If LinkOutWhite <> 0 Then v = v Or 8
    If LinkOutRed <> 0 Then v = v Or 4
    If (LinkInWhite <> 0) And (LinkOutWhite = 0) Then v = v Or 2
    If (LinkInRed <> 0) And (LinkOutRed = 0) Then v = v Or 1
    LinkReadPort = v
End Function

Public Sub LinkWritePort(ByVal Value As Long)
    LinkOutRed = (Value \ 4) And 1
    LinkOutWhite = (Value \ 8) And 1

    If mTxActive Then LinkTxStep
    If mRxActive Then LinkRxStep
    If mTxCount > 0 Then LinkKickTx
End Sub

Public Sub LinkPoll()
    Dim elapsed As Long
    Dim limit As Long

    If Not mSessionActive Then
        If LINK_AUTO_RECEIVE Then
            If mRxCount > 0 Then LinkBeginReceive True
        Else
            If mRxCount > 0 Then RxFlush
        End If
    End If

    If Not mSessionActive Then Exit Sub

    If mSession = LS_RECV Then
        If mRecvState = RS_WAIT_BACKUP_EOT Then
            If GetTickCount >= mBackupGraceUntil Then
                FinishReceive
                Exit Sub
            End If
        End If
    End If

    elapsed = GetTickCount - mLastActivity
    If mSession = LS_SEND Then
        limit = LINK_TIMEOUT_MS
        ' A backup makes the receiving calculator display a
        ' "Memory Backup" confirmation before it sends CTS.
        If mSendIsBackup And mSendState = SS_WAIT_ACK_CTS Then
            limit = LINK_ARM_TIMEOUT_MS
        End If
    Else
        If mRecvGotData Then
            limit = LINK_TIMEOUT_MS
        Else
            limit = LINK_ARM_TIMEOUT_MS
        End If
    End If
    If elapsed > limit Then
        LinkAbort "Link timeout"
        Exit Sub
    End If

    Do While mRxCount >= 4
        If Not ParseOnePacket() Then Exit Do
        If Not mSessionActive Then Exit Sub
    Loop

    LinkKickTx
End Sub

Public Sub LinkCancel()
    If Not mSessionActive Then Exit Sub
    LinkAbort "Transfer cancelled"
End Sub

Public Function LinkArmReceive() As Boolean
    If mSessionActive Then LinkCancel
    RxFlush
    TxFlush
    LinkBeginReceive False
    LinkArmReceive = True
End Function

Public Function LinkStartSendFile(ByVal FilePath As String) As Boolean
    Dim ok As Boolean

    LinkLastError = ""
    LinkReset
    ok = LoadLinkFile(FilePath)
    If Not ok Then
        LinkStartSendFile = False
        Exit Function
    End If

    mSendFileName = FilePath
    mSession = LS_SEND
    mSessionActive = True
    mSendIndex = 0
    mSendState = SS_WAIT_ACK_CTS
    mPeerID = 0
    mCtsHandled = False
    mLastActivity = GetTickCount
    LinkStatusText = "Sending " & FileBaseName(FilePath)

    LinkSendCurrentVar
    LinkKickTx
    LinkStartSendFile = True
End Function

'============================================================
' Wire byte engine
'============================================================
Private Sub LinkTxStep()
    If mTxPhase = TXP_WAIT_ACK Then
        If (mTxBitVal = 1 And LinkOutRed = 1) Or _
           (mTxBitVal = 0 And LinkOutWhite = 1) Then
            LinkInRed = 1
            LinkInWhite = 1
            mTxPhase = TXP_WAIT_CLEAR
            mLastActivity = GetTickCount
        End If
    ElseIf mTxPhase = TXP_WAIT_CLEAR Then
        If LinkOutRed = 0 And LinkOutWhite = 0 Then
            mTxBit = mTxBit + 1
            mLastActivity = GetTickCount
            If mTxBit >= 8 Then
                mTxBit = 0
                If mTxCount > 0 Then
                    mTxByte = TxPop()
                    LinkTxAssertBit
                Else
                    mTxActive = False
                    mTxPhase = 0
                    mRxActive = True
                    mRxPhase = RXP_WAIT_START
                    mRxByte = 0
                    mRxBit = 0
                End If
            Else
                LinkTxAssertBit
            End If
        End If
    End If
End Sub

Private Sub LinkTxAssertBit()
    mTxBitVal = (mTxByte \ (2 ^ mTxBit)) And 1
    If mTxBitVal = 1 Then
        LinkInRed = 1
        LinkInWhite = 0
    Else
        LinkInRed = 0
        LinkInWhite = 1
    End If
    mTxPhase = TXP_WAIT_ACK
    mLastActivity = GetTickCount
End Sub

Private Sub LinkRxStep()
    If mRxPhase = RXP_WAIT_START Then
        If (LinkOutRed <> 0) Or (LinkOutWhite <> 0) Then
            mRxBitVal = LinkOutWhite
            LinkInWhite = LinkOutWhite
            LinkInRed = LinkOutRed
            mRxPhase = RXP_WAIT_CLEAR
            mLastActivity = GetTickCount
        End If
    ElseIf mRxPhase = RXP_WAIT_CLEAR Then
        If LinkOutRed = 0 And LinkOutWhite = 0 Then
            LinkInRed = 1
            LinkInWhite = 1
            If mRxBitVal <> 0 Then mRxByte = mRxByte Or (2 ^ mRxBit)
            mRxBit = mRxBit + 1
            mLastActivity = GetTickCount
            mRecvGotData = True
            If mRxBit >= 8 Then
                RxPush CByte(mRxByte And &HFF&)
                mRxByte = 0
                mRxBit = 0
            End If
            mRxPhase = RXP_WAIT_START
        End If
    End If
End Sub

Private Sub LinkStartTxByte()
    If mTxCount = 0 Then Exit Sub
    mTxByte = TxPop()
    mTxBit = 0
    mTxActive = True
    mRxActive = False
    LinkTxAssertBit
End Sub

Private Sub LinkKickTx()
    If mTxCount = 0 Then Exit Sub
    If mTxActive Then Exit Sub
    If mRxActive And (mRxPhase <> RXP_WAIT_START) Then Exit Sub
    LinkStartTxByte
End Sub

'============================================================
' Byte queues
'============================================================
Private Sub TxFlush()
    mTxHead = 0
    mTxTail = 0
    mTxCount = 0
End Sub

Private Sub RxFlush()
    mRxHead = 0
    mRxTail = 0
    mRxCount = 0
End Sub

Private Sub TxPush(ByVal b As Byte)
    If mTxCount >= LINK_QCAP Then
        LinkAbort "Transmit queue overflow"
        Exit Sub
    End If
    mTxBuf(mTxTail) = b
    mTxTail = (mTxTail + 1) And (LINK_QCAP - 1)
    mTxCount = mTxCount + 1
End Sub

Private Function TxPop() As Byte
    TxPop = mTxBuf(mTxHead)
    mTxHead = (mTxHead + 1) And (LINK_QCAP - 1)
    mTxCount = mTxCount - 1
End Function

Private Sub RxPush(ByVal b As Byte)
    If mRxCount >= LINK_QCAP Then
        ' Should not happen: the protocol consumes packets promptly.
        mRxHead = 0
        mRxTail = 0
        mRxCount = 0
    End If
    mRxBuf(mRxTail) = b
    mRxTail = (mRxTail + 1) And (LINK_QCAP - 1)
    mRxCount = mRxCount + 1
End Sub

Private Function RxPeek(ByVal Offset As Long) As Byte
    RxPeek = mRxBuf((mRxHead + Offset) And (LINK_QCAP - 1))
End Function

Private Sub RxDropBytes(ByVal n As Long)
    If n > mRxCount Then n = mRxCount
    mRxHead = (mRxHead + n) And (LINK_QCAP - 1)
    mRxCount = mRxCount - n
End Sub

'============================================================
' Packet layer
'============================================================
Private Function PeerID() As Byte
    If mPeerID <> 0 Then
        PeerID = mPeerID
    Else
        PeerID = LINK_ID_TI85
    End If
End Function

Private Sub QueueCmd(ByVal MachineID As Byte, ByVal Command As Byte)
    If Not mSessionActive Then Exit Sub
    If mTxCount + 4 > LINK_QCAP Then
        LinkAbort "Transmit queue overflow"
        Exit Sub
    End If
    TxPush MachineID
    TxPush Command
    TxPush 0
    TxPush 0
End Sub

Private Sub QueuePacket(ByVal MachineID As Byte, ByVal Command As Byte, ByRef Data() As Byte)
    Dim n As Long, i As Long, csum As Long

    If Not mSessionActive Then Exit Sub
    n = DataLength(Data)
    If mTxCount + 4 + n + IIf(n > 0, 2, 0) > LINK_QCAP Then
        LinkAbort "Transmit queue overflow"
        Exit Sub
    End If

    TxPush MachineID
    TxPush Command
    TxPush CByte(n And &HFF&)
    TxPush CByte((n \ 256) And &HFF&)
    For i = 0 To n - 1
        TxPush Data(i)
        csum = (csum + Data(i)) And &HFFFF&
    Next
    If n > 0 Then
        TxPush CByte(csum And &HFF&)
        TxPush CByte((csum \ 256) And &HFF&)
    End If
End Sub

Private Function ParseOnePacket() As Boolean
    Dim id As Long, cmd As Long, ln As Long, got As Long
    Dim i As Long, cs As Long, stored As Long
    Dim d() As Byte
    Dim hasData As Boolean

    If mRxCount < 4 Then Exit Function

    id = RxPeek(0)
    cmd = RxPeek(1)
    ln = CLng(RxPeek(2)) + 256& * CLng(RxPeek(3))

    ' ACK, CTS, EOT and ERR carry no data. Their length field is
    ' unreliable (the TI-85 may echo the last data length in EOT),
    ' so the packet is always just the four header bytes.
    Select Case cmd
    Case LCMD_ACK, LCMD_CTS, LCMD_EOT, LCMD_ERR
        hasData = False
    Case Else
        hasData = True
    End Select

    If hasData Then
        got = 4 + ln + 2
    Else
        got = 4
    End If

    If got > LINK_QCAP Then
        LinkAbort "Link packet too large"
        Exit Function
    End If
    If mRxCount < got Then Exit Function

    If hasData And ln > 0 Then
        ReDim d(0 To ln - 1)
        cs = 0
        For i = 0 To ln - 1
            d(i) = RxPeek(4 + i)
            cs = (cs + d(i)) And &HFFFF&
        Next
        stored = CLng(RxPeek(4 + ln)) + 256& * CLng(RxPeek(5 + ln))
        If cs <> stored Then
            RxDropBytes got
            LinkAbort "Link checksum error"
            Exit Function
        End If
    Else
        ReDim d(0 To 0)
        d(0) = 0
    End If

    RxDropBytes got
    mPeerID = CByte(id And &HFF&)
    mRecvGotData = True
    HandlePacket CByte(cmd And &HFF&), d, ln
    ParseOnePacket = True
End Function

Private Sub HandlePacket(ByVal Command As Byte, ByRef Data() As Byte, ByVal n As Long)
    If mSession = LS_SEND Then
        HandleSendPacket Command, Data, n
    ElseIf mSession = LS_RECV Then
        HandleReceivePacket Command, Data, n
    End If
End Sub

'============================================================
' Send session
'============================================================
Private Sub HandleSendPacket(ByVal Command As Byte, ByRef Data() As Byte, ByVal n As Long)
    Select Case mSendState
    Case SS_WAIT_ACK_CTS
        Select Case Command
        Case LCMD_CTS
            If Not mCtsHandled Then
                mCtsHandled = True
                QueueCmd PeerID, LCMD_ACK
                If mSendIsBackup Then
                    QueuePacket PeerID, LCMD_DATA, mBackupData1
                    mSendState = SS_BACKUP_ACK1
                Else
                    Dim d() As Byte
                    d = mSendEntries(mSendIndex).Data
                    QueuePacket PeerID, LCMD_DATA, d
                    mSendState = SS_WAIT_DATA_ACK
                End If
            End If
        Case LCMD_SKIP
            If Not mCtsHandled Then
                mCtsHandled = True
                QueueCmd PeerID, LCMD_ACK
                If n > 0 And Data(0) = 1 Then
                    LinkAbort "Calculator cancelled the transfer"
                Else
                    mSendState = SS_NEXT
                    LinkSendAdvance
                End If
            End If
        Case LCMD_ERR
            LinkAbort "Calculator reported a checksum error"
        End Select

    Case SS_WAIT_DATA_ACK
        If Command = LCMD_ACK Then
            mSendState = SS_NEXT
            LinkSendAdvance
        ElseIf Command = LCMD_ERR Then
            LinkAbort "Calculator reported a checksum error"
        End If

    Case SS_BACKUP_ACK1
        If Command = LCMD_ACK Then
            QueuePacket PeerID, LCMD_DATA, mBackupData2
            mSendState = SS_BACKUP_ACK2
        End If

    Case SS_BACKUP_ACK2
        If Command = LCMD_ACK Then
            QueuePacket PeerID, LCMD_DATA, mBackupData3
            mSendState = SS_BACKUP_ACK3
        End If

    Case SS_BACKUP_ACK3
        If Command = LCMD_ACK Then
            ' VTI sends no EOT for backups; the transfer is complete
            ' after the third section has been acknowledged.
            FinishSend "Backup sent"
        End If

    Case SS_WAIT_EOT_ACK
        If Command = LCMD_ACK Then
            FinishSend "File sent"
        End If
    End Select
End Sub

Private Sub LinkSendAdvance()
    mSendIndex = mSendIndex + 1
    If mSendIndex >= mSendCount Then
        QueueCmd PeerID, LCMD_EOT
        mSendState = SS_WAIT_EOT_ACK
    Else
        mCtsHandled = False
        mSendState = SS_WAIT_ACK_CTS
        LinkStatusText = "Sending " & FileBaseName(mSendFileName) & _
                         " (" & CStr(mSendIndex + 1) & "/" & CStr(mSendCount) & ")"
        LinkSendCurrentVar
    End If
    LinkKickTx
End Sub

Private Sub LinkSendCurrentVar()
    Dim d() As Byte

    mCtsHandled = False
    mSendState = SS_WAIT_ACK_CTS
    If mSendIsBackup Then
        MakeBackupVarData d
    Else
        MakeVarData mSendEntries(mSendIndex), d
    End If
    QueuePacket LINK_ID_TI85, LCMD_VAR, d
End Sub

Private Sub MakeVarData(ByRef e As LinkEntry, ByRef d() As Byte)
    Dim i As Long

    ReDim d(0 To 3 + e.NameLen)
    d(0) = CByte(e.dataLen And &HFF&)
    d(1) = CByte((e.dataLen \ 256) And &HFF&)
    d(2) = e.VarType
    d(3) = CByte(e.NameLen And &HFF&)
    For i = 0 To e.NameLen - 1
        d(4 + i) = CByte(Asc(Mid$(e.Name, i + 1, 1)) And &HFF&)
    Next
End Sub

Private Sub MakeBackupVarData(ByRef d() As Byte)
    ReDim d(0 To 8)
    d(0) = CByte(mBackupSize1 And &HFF&)
    d(1) = CByte((mBackupSize1 \ 256) And &HFF&)
    d(2) = &H1D
    d(3) = CByte(mBackupSize2 And &HFF&)
    d(4) = CByte((mBackupSize2 \ 256) And &HFF&)
    d(5) = CByte(mBackupSize3 And &HFF&)
    d(6) = CByte((mBackupSize3 \ 256) And &HFF&)
    d(7) = CByte(mBackupAddr And &HFF&)
    d(8) = CByte((mBackupAddr \ 256) And &HFF&)
End Sub

Private Sub FinishSend(ByVal Message As String)
    mSessionActive = False
    mSession = LS_NONE
    LinkStatusText = ""
    LinkNotifyText = Message
    LinkNotifyError = False
    LinkNotifyPending = True
    LinkInRed = 1
    LinkInWhite = 1
    mTxActive = False
    mRxActive = True
    mRxPhase = RXP_WAIT_START
    TxFlush
    RxFlush
End Sub

'============================================================
' Receive session
'============================================================
Private Sub LinkBeginReceive(ByVal AutoMode As Boolean)
    If mSessionActive Then Exit Sub
    If Not AutoMode Then RxFlush

    mSession = LS_RECV
    mSessionActive = True
    mRecvState = RS_WAIT_VAR
    mRecvCount = 0
    mRecvIsBackup = False
    mRecvCurEntry = 0
    mRecvGotData = False
    mLastActivity = GetTickCount
    mPeerID = 0
    mCtsHandled = False
    Erase mRecvEntries
    Erase mRecvBackup1
    Erase mRecvBackup2
    Erase mRecvBackup3

    If AutoMode Then
        LinkStatusText = "Receiving data..."
    Else
        LinkStatusText = "Waiting for calculator data..."
    End If
End Sub

Private Sub HandleReceivePacket(ByVal Command As Byte, ByRef Data() As Byte, ByVal n As Long)
    Select Case mRecvState
    Case RS_WAIT_VAR, RS_WAIT_NEXT
        If Command = LCMD_VAR Then
            If ParseIncomingVar(Data, n) Then
                QueueCmd PeerID, LCMD_ACK
                QueueCmd PeerID, LCMD_CTS
                mRecvState = RS_WAIT_ACK
                mRecvGotData = True
            Else
                LinkAbort "Invalid variable header"
            End If
        ElseIf Command = LCMD_EOT Then
            QueueCmd PeerID, LCMD_ACK
            FinishReceive
        End If

    Case RS_WAIT_ACK
        If Command = LCMD_ACK Then
            If mRecvIsBackup Then
                mRecvState = RS_WAIT_BACKUP1
            Else
                mRecvState = RS_WAIT_DATA
            End If
        End If

    Case RS_WAIT_DATA
        If Command = LCMD_DATA Then
            StoreRecvData Data, n
            QueueCmd PeerID, LCMD_ACK
            mRecvState = RS_WAIT_NEXT
        End If

    Case RS_WAIT_BACKUP1
        If Command = LCMD_DATA Then
            CopyByteArray Data, mRecvBackup1, n
            QueueCmd PeerID, LCMD_ACK
            mRecvState = RS_WAIT_BACKUP2
        End If

    Case RS_WAIT_BACKUP2
        If Command = LCMD_DATA Then
            CopyByteArray Data, mRecvBackup2, n
            QueueCmd PeerID, LCMD_ACK
            mRecvState = RS_WAIT_BACKUP3
        End If

    Case RS_WAIT_BACKUP3
        If Command = LCMD_DATA Then
            CopyByteArray Data, mRecvBackup3, n
            QueueCmd PeerID, LCMD_ACK
            mRecvState = RS_WAIT_BACKUP_EOT
            mBackupGraceUntil = GetTickCount + LINK_BACKUP_EOT_GRACE_MS
        End If

    Case RS_WAIT_BACKUP_EOT
        If Command = LCMD_EOT Then
            QueueCmd PeerID, LCMD_ACK
            FinishReceive
        End If
    End Select
End Sub

Private Function ParseIncomingVar(ByRef d() As Byte, ByVal n As Long) As Boolean
    Dim dl As Long, typ As Byte, nl As Long, nm As String

    If n = 9 And d(2) = &H1D Then
        mRecvIsBackup = True
        mRecvBackupSize1 = CLng(d(0)) + 256& * CLng(d(1))
        mRecvBackupSize2 = CLng(d(3)) + 256& * CLng(d(4))
        mRecvBackupSize3 = CLng(d(5)) + 256& * CLng(d(6))
        mRecvBackupAddr = CLng(d(7)) + 256& * CLng(d(8))
        mRecvCount = 0
        Erase mRecvEntries
        ParseIncomingVar = True
        Exit Function
    End If

    If n < 4 Then Exit Function
    dl = CLng(d(0)) + 256& * CLng(d(1))
    typ = d(2)
    nl = CLng(d(3))
    If nl > 8 Or (4 + nl) > n Then Exit Function

    nm = BytesToText(d, 4, nl)
    ReDim Preserve mRecvEntries(0 To mRecvCount)
    With mRecvEntries(mRecvCount)
        .dataLen = dl
        .VarType = typ
        .NameLen = nl
        .Name = nm
        .Data = Empty
    End With
    mRecvCurEntry = mRecvCount
    mRecvCount = mRecvCount + 1
    mRecvIsBackup = False
    ParseIncomingVar = True
End Function

Private Sub StoreRecvData(ByRef d() As Byte, ByVal n As Long)
    Dim x() As Byte

    CopyByteArray d, x, n
    mRecvEntries(mRecvCurEntry).Data = x
    mRecvEntries(mRecvCurEntry).dataLen = n
End Sub

Private Sub CopyByteArray(ByRef src() As Byte, ByRef dst() As Byte, ByVal n As Long)
    Dim i As Long

    If n <= 0 Then
        ReDim dst(0 To 0)
        Exit Sub
    End If
    ReDim dst(0 To n - 1)
    For i = 0 To n - 1
        dst(i) = src(i)
    Next
End Sub

Private Sub FinishReceive()
    Dim saved As String

    mSessionActive = False
    mSession = LS_NONE
    saved = SaveReceivedData
    LinkStatusText = ""

    If Len(saved) > 0 Then
        LinkNotifyText = "Received data saved to:" & vbCrLf & saved
        LinkNotifyError = False
    Else
        If Len(LinkLastError) = 0 Then LinkLastError = "Receive failed"
        LinkNotifyText = LinkLastError
        LinkNotifyError = True
    End If
    LinkNotifyPending = True

    LinkInRed = 1
    LinkInWhite = 1
    mTxActive = False
    mRxActive = True
    mRxPhase = RXP_WAIT_START
    TxFlush
    RxFlush
End Sub

'============================================================
' Abort
'============================================================
Public Sub LinkAbort(ByVal Message As String)
    If Len(Message) > 0 Then LinkLastError = Message

    mSessionActive = False
    mSession = LS_NONE
    mTxActive = False
    mRxActive = True
    mRxPhase = RXP_WAIT_START
    mRxByte = 0
    mRxBit = 0
    LinkInRed = 1
    LinkInWhite = 1
    LinkStatusText = ""
    TxFlush
    RxFlush

    If Len(Message) > 0 Then
        LinkNotifyText = Message
        LinkNotifyError = True
        LinkNotifyPending = True
    End If
End Sub

'============================================================
' File loading (send side)
'============================================================
Private Function LoadLinkFile(ByVal FilePath As String) As Boolean
    Dim f As Integer, n As Long
    Dim b() As Byte
    Dim off As Long, hdr As Long, dl As Long
    Dim typ As Byte, nl As Long
    Dim nm As String
    Dim d() As Byte
    Dim s1 As Long, s2 As Long, s3 As Long, Addr As Long
    Dim o As Long

    On Error GoTo Fail
    f = FreeFile
    Open FilePath For Binary As #f
    n = LOF(f)
    If n < 57 Then
        Close #f
        LinkLastError = "File is too small"
        Exit Function
    End If
    ReDim b(0 To n - 1)
    Get #f, , b
    Close #f

    If BytesToText(b, 0, 8) <> "**TI85**" Then
        LinkLastError = "Not a TI-85 file"
        Exit Function
    End If

    mSendCount = 0
    mSendIsBackup = False
    Erase mSendEntries

    off = 55
    hdr = WordAt(b, off)
    dl = WordAt(b, off + 2)
    typ = b(off + 4)

    If typ = &H1D Then
        ' Backup file: 9-byte backup header + three length-prefixed sections.
        s1 = WordAt(b, off + 2)
        s2 = WordAt(b, off + 5)
        s3 = WordAt(b, off + 7)
        Addr = WordAt(b, off + 9)
        o = off + 11

        If o + 2 > n Then GoTo Fail
        If WordAt(b, o) <> s1 Then GoTo Fail
        o = o + 2
        If o + s1 > n Then GoTo Fail
        CopyRange b, o, s1, mBackupData1
        o = o + s1

        If o + 2 > n Then GoTo Fail
        If WordAt(b, o) <> s2 Then GoTo Fail
        o = o + 2
        If o + s2 > n Then GoTo Fail
        CopyRange b, o, s2, mBackupData2
        o = o + s2

        If o + 2 > n Then GoTo Fail
        If WordAt(b, o) <> s3 Then GoTo Fail
        o = o + 2
        If o + s3 > n Then GoTo Fail
        CopyRange b, o, s3, mBackupData3

        mBackupSize1 = s1
        mBackupSize2 = s2
        mBackupSize3 = s3
        mBackupAddr = Addr
        mSendIsBackup = True
        mSendCount = 1
        LoadLinkFile = True
        Exit Function
    End If

    ' Normal variable file.  Entries start at offset 55.
    off = 55
    Do While off + 8 <= n
        hdr = WordAt(b, off)
        dl = WordAt(b, off + 2)
        typ = b(off + 4)
        nl = CLng(b(off + 5))

        If hdr = 0 Or hdr = &HFFFF& Then Exit Do
        If nl > 8 Then GoTo Fail
        If off + 8 + nl > n Then GoTo Fail
        If off + 8 + nl + dl > n Then GoTo Fail

        nm = BytesToText(b, off + 6, nl)
        ReDim Preserve mSendEntries(0 To mSendCount)
        With mSendEntries(mSendCount)
            .dataLen = dl
            .VarType = typ
            .NameLen = nl
            .Name = nm
            .Data = Empty
        End With
        CopyRange b, off + 8 + nl, dl, d
        mSendEntries(mSendCount).Data = d
        mSendCount = mSendCount + 1

        off = off + 8 + nl + dl
    Loop

    If mSendCount = 0 Then
        LinkLastError = "No variables found in file"
        Exit Function
    End If

    LoadLinkFile = True
    Exit Function

Fail:
    On Error Resume Next
    If f > 0 Then Close #f
    If Len(LinkLastError) = 0 Then LinkLastError = "Invalid TI-85 file"
End Function

Private Sub CopyRange(ByRef src() As Byte, ByVal Start As Long, ByVal count As Long, ByRef dst() As Byte)
    Dim i As Long

    If count <= 0 Then
        ReDim dst(0 To 0)
        Exit Sub
    End If
    ReDim dst(0 To count - 1)
    For i = 0 To count - 1
        dst(i) = src(Start + i)
    Next
End Sub

'============================================================
' File saving (receive side)
'============================================================
Private Function SaveReceivedData() As String
    Dim path As String, dir As String
    Dim b() As Byte
    Dim nm As String

    dir = ReceiveDir()
    If Len(dir) = 0 Then
        LinkLastError = "Cannot create receive folder"
        Exit Function
    End If

    If mRecvIsBackup Then
        path = dir & "\backup_" & Format$(Now, "yyyymmdd_hhnnss") & ".85b"
        BuildBackupFile b
    Else
        If mRecvCount = 0 Then
            LinkLastError = "No data received"
            Exit Function
        End If
        If mRecvCount = 1 Then
            nm = Left$(mRecvEntries(0).Name, mRecvEntries(0).NameLen)
            path = dir & "\" & SafeFileName(nm) & "." & ExtForType(mRecvEntries(0).VarType)
        Else
            path = dir & "\group_" & Format$(Now, "yyyymmdd_hhnnss") & ".85g"
        End If
        BuildVariableFile b
    End If

    path = UniquePath(path)
    If Not WriteByteFile(path, b) Then Exit Function
    SaveReceivedData = path
End Function

Private Sub BuildVariableFile(ByRef out() As Byte)
    Dim i As Long, k As Long, j As Long
    Dim dataLen As Long, nl As Long, dl As Long, csum As Long
    Dim d() As Byte
    Dim sig As String

    For i = 0 To mRecvCount - 1
        nl = mRecvEntries(i).NameLen
        d = mRecvEntries(i).Data
        dl = DataLength(d)
        dataLen = dataLen + 8 + nl + dl
    Next

    ReDim out(0 To 55 + dataLen + 2 - 1)
    sig = "**TI85**"
    For i = 0 To 7
        out(i) = CByte(Asc(Mid$(sig, i + 1, 1)))
    Next
    out(8) = &H1A
    out(9) = &HC
    out(10) = 0
    out(53) = CByte(dataLen And &HFF&)
    out(54) = CByte((dataLen \ 256) And &HFF&)

    j = 55
    For i = 0 To mRecvCount - 1
        nl = mRecvEntries(i).NameLen
        d = mRecvEntries(i).Data
        dl = DataLength(d)

        out(j) = CByte((6 + nl) And &HFF&)
        out(j + 1) = CByte(((6 + nl) \ 256) And &HFF&)
        out(j + 2) = CByte(dl And &HFF&)
        out(j + 3) = CByte((dl \ 256) And &HFF&)
        out(j + 4) = mRecvEntries(i).VarType
        out(j + 5) = CByte(nl And &HFF&)
        For k = 0 To nl - 1
            out(j + 6 + k) = CByte(Asc(Mid$(mRecvEntries(i).Name, k + 1, 1)) And &HFF&)
        Next
        out(j + 6 + nl) = CByte(dl And &HFF&)
        out(j + 7 + nl) = CByte((dl \ 256) And &HFF&)
        For k = 0 To dl - 1
            out(j + 8 + nl + k) = d(k)
        Next
        j = j + 8 + nl + dl
    Next

    csum = 0
    For i = 55 To 55 + dataLen - 1
        csum = (csum + out(i)) And &HFFFF&
    Next
    out(55 + dataLen) = CByte(csum And &HFF&)
    out(55 + dataLen + 1) = CByte((csum \ 256) And &HFF&)
End Sub

Private Sub BuildBackupFile(ByRef out() As Byte)
    Dim s1 As Long, s2 As Long, s3 As Long
    Dim dataLen As Long, j As Long, i As Long, csum As Long
    Dim sig As String

    s1 = DataLength(mRecvBackup1)
    s2 = DataLength(mRecvBackup2)
    s3 = DataLength(mRecvBackup3)
    If s1 = 0 Then s1 = mRecvBackupSize1
    If s2 = 0 Then s2 = mRecvBackupSize2
    If s3 = 0 Then s3 = mRecvBackupSize3

    dataLen = 17 + s1 + s2 + s3
    ReDim out(0 To 55 + dataLen + 2 - 1)

    sig = "**TI85**"
    For i = 0 To 7
        out(i) = CByte(Asc(Mid$(sig, i + 1, 1)))
    Next
    out(8) = &H1A
    out(9) = &HC
    out(10) = 0
    out(53) = CByte(dataLen And &HFF&)
    out(54) = CByte((dataLen \ 256) And &HFF&)

    j = 55
    PutWord out, j, 9: j = j + 2
    PutWord out, j, s1: j = j + 2
    out(j) = &H1D: j = j + 1
    PutWord out, j, s2: j = j + 2
    PutWord out, j, s3: j = j + 2
    PutWord out, j, mRecvBackupAddr: j = j + 2

    PutWord out, j, s1: j = j + 2
    For i = 0 To s1 - 1: out(j + i) = mRecvBackup1(i): Next
    j = j + s1
    PutWord out, j, s2: j = j + 2
    For i = 0 To s2 - 1: out(j + i) = mRecvBackup2(i): Next
    j = j + s2
    PutWord out, j, s3: j = j + 2
    For i = 0 To s3 - 1: out(j + i) = mRecvBackup3(i): Next

    csum = 0
    For i = 55 To 55 + dataLen - 1
        csum = (csum + out(i)) And &HFFFF&
    Next
    out(55 + dataLen) = CByte(csum And &HFF&)
    out(55 + dataLen + 1) = CByte((csum \ 256) And &HFF&)
End Sub

Private Sub PutWord(ByRef b() As Byte, ByVal Offset As Long, ByVal Value As Long)
    b(Offset) = CByte(Value And &HFF&)
    b(Offset + 1) = CByte((Value \ 256) And &HFF&)
End Sub

Private Function ReceiveDir() As String
    Dim p As String, q As String

    On Error Resume Next
    p = App.path & "\link"
    If dir$(p, vbDirectory) = "" Then MkDir p
    q = p & "\receive"
    If dir$(q, vbDirectory) = "" Then MkDir q
    On Error GoTo 0
    ReceiveDir = q
End Function

Private Function UniquePath(ByVal FilePath As String) As String
    Dim base As String, ext As String, n As Long, i As Long, p As String

    If dir$(FilePath) = "" Then
        UniquePath = FilePath
        Exit Function
    End If

    n = InStrRev(FilePath, ".")
    If n > 0 Then
        base = Left$(FilePath, n - 1)
        ext = Mid$(FilePath, n)
    Else
        base = FilePath
        ext = ""
    End If

    i = 1
    Do
        p = base & "_" & CStr(i) & ext
        If dir$(p) = "" Then
            UniquePath = p
            Exit Function
        End If
        i = i + 1
    Loop
End Function

Private Function WriteByteFile(ByVal FilePath As String, ByRef b() As Byte) As Boolean
    Dim f As Integer

    On Error GoTo Fail
    f = FreeFile
    Open FilePath For Binary As #f
    Put #f, , b
    Close #f
    WriteByteFile = True
    Exit Function

Fail:
    On Error Resume Next
    If f > 0 Then Close #f
    LinkLastError = "Cannot write " & FilePath
End Function

Private Function ExtForType(ByVal typ As Byte) As String
    Select Case typ
    Case &H0: ExtForType = "85n"
    Case &H1: ExtForType = "85c"
    Case &H2, &H3: ExtForType = "85v"
    Case &H4, &H5: ExtForType = "85l"
    Case &H6, &H7: ExtForType = "85m"
    Case &H8, &H9: ExtForType = "85k"
    Case &HA: ExtForType = "85e"
    Case &HC: ExtForType = "85s"
    Case &HD To &H10: ExtForType = "85d"
    Case &H11: ExtForType = "85i"
    Case &H12: ExtForType = "85p"
    Case &H17 To &H1B: ExtForType = "85r"
    Case &H1D: ExtForType = "85b"
    Case Else: ExtForType = "85x"
    End Select
End Function

Private Function SafeFileName(ByVal s As String) As String
    Dim i As Long, ch As String, out As String

    s = Trim$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If InStr(1, "\/:*?""<>|", ch, vbBinaryCompare) = 0 Then out = out & ch
    Next
    If Len(out) = 0 Then out = "var"
    If Len(out) > 32 Then out = Left$(out, 32)
    SafeFileName = out
End Function

Private Function FileBaseName(ByVal FilePath As String) As String
    Dim i As Long

    i = InStrRev(FilePath, "\")
    If i > 0 Then
        FileBaseName = Mid$(FilePath, i + 1)
    Else
        FileBaseName = FilePath
    End If
End Function

'============================================================
' Helpers
'============================================================
Private Function WordAt(ByRef b() As Byte, ByVal Offset As Long) As Long
    WordAt = CLng(b(Offset)) + 256& * CLng(b(Offset + 1))
End Function

Private Function BytesToText(ByRef b() As Byte, ByVal Start As Long, ByVal count As Long) As String
    Dim i As Long, s As String

    For i = 0 To count - 1
        s = s & Chr$(b(Start + i))
    Next
    BytesToText = s
End Function

Private Function DataLength(ByRef d() As Byte) As Long
    On Error GoTo NoArray
    DataLength = UBound(d) - LBound(d) + 1
    Exit Function
NoArray:
    DataLength = 0
End Function
