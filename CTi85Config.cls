VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTi85Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long

Private mFirstTime      As Boolean
Private mRomFilePath    As String
Private mStateFilePath  As String
Private mThemeName      As String
Private mLCDName        As String
Private mWndLeft        As Single
Private mWndTop         As Single
Private mWndHeight      As Single
Private mWndWidth       As Single

'================================================
' Class_Initialize
'================================================
Private Sub Class_Initialize()
    mStateFilePath = ReadIniString("StateFilePath", App.path & "\TI85.STATE")
    mRomFilePath = ReadIniString("RomFilePath", App.path & "\TI85.ROM")
    mLCDName = ReadIniString("LCDName", "BLUEONGREEN")
    mThemeName = ReadIniString("ThemeName", "DARK")
    mFirstTime = val(ReadIniString("FirstTime", "1"))
    mWndTop = val(ReadIniString("WndTop", "0"))
    mWndLeft = val(ReadIniString("WndLeft", "0"))
    mWndWidth = val(ReadIniString("WndWidth", "0"))
    mWndHeight = val(ReadIniString("WndHeight", "0"))
End Sub

'================================================
' ReadIniString
'================================================
Private Function ReadIniString(ByRef pKey As String, Optional ByRef pDefault As String = "") As String
    Dim vTmp As String
    Dim vLen As Long
    vLen = 260
    vTmp = String$(vLen, 0)
    vLen = GetPrivateProfileString("Ti85Emu", pKey, pDefault, vTmp, vLen, ConfigFilePath)
    ReadIniString = Left$(vTmp, vLen)
End Function

'================================================
' WriteIniString
'================================================
Private Sub WriteIniString(ByRef pKey As String, ByRef pValue As String)
    WritePrivateProfileString "Ti85Emu", pKey, pValue, ConfigFilePath
End Sub

'================================================
' ConfigFilePath
'================================================
Public Property Get ConfigFilePath() As String
    ConfigFilePath = App.path & "\TI85.INI"
End Property

'================================================
' ThemeName
'================================================
Public Property Get ThemeName() As String
    ThemeName = mThemeName
End Property

'================================================
' ThemeName
'================================================
Public Property Let ThemeName(ByRef RHS As String)
    If StrComp(RHS, mThemeName) = 0 Then Exit Property
    WriteIniString "ThemeName", RHS
    mThemeName = RHS
End Property

'================================================
' LCDName
'================================================
Public Property Get LCDName() As String
    LCDName = mLCDName
End Property

'================================================
' LCDName
'================================================
Public Property Let LCDName(ByRef RHS As String)
    If StrComp(RHS, mLCDName) = 0 Then Exit Property
    WriteIniString "LCDName", RHS
    mLCDName = RHS
End Property

'================================================
' StateFilePath
'================================================
Public Property Get StateFilePath() As String
    StateFilePath = mStateFilePath
End Property

'================================================
' StateFilePath
'================================================
Public Property Let StateFilePath(ByRef RHS As String)
    If StrComp(RHS, mStateFilePath) = 0 Then Exit Property
    WriteIniString "StateFilePath", RHS
    mStateFilePath = RHS
End Property

'================================================
' RomFilePath
'================================================
Public Property Get RomFilePath() As String
    RomFilePath = mRomFilePath
End Property

'================================================
' RomFilePath
'================================================
Public Property Let RomFilePath(ByRef RHS As String)
    If StrComp(RHS, mRomFilePath) = 0 Then Exit Property
    WriteIniString "RomFilePath", RHS
    mRomFilePath = RHS
End Property

'================================================
' FirstTime
'================================================
Public Property Get FirstTime() As Boolean
    FirstTime = mFirstTime
End Property

'================================================
' FirstTime
'================================================
Public Property Let FirstTime(ByVal RHS As Boolean)
    If RHS = mFirstTime Then Exit Property
    WriteIniString "FirstTime", Abs(RHS)
    mFirstTime = RHS
End Property

'================================================
' WndLeft
'================================================
Public Property Get WndLeft() As Single
    WndLeft = mWndLeft
End Property

'================================================
' WndLeft
'================================================
Public Property Let WndLeft(ByVal RHS As Single)
    If RHS = mWndLeft Then Exit Property
    WriteIniString "WndLeft", CStr(RHS)
    mWndLeft = RHS
End Property

'================================================
' WndTop
'================================================
Public Property Get WndTop() As Single
    WndTop = mWndTop
End Property

'================================================
' WndTop
'================================================
Public Property Let WndTop(ByVal RHS As Single)
    If RHS = mWndTop Then Exit Property
    WriteIniString "WndTop", CStr(RHS)
    mWndTop = RHS
End Property

'================================================
' WndWidth
'================================================
Public Property Get WndWidth() As Single
    WndWidth = mWndWidth
End Property

'================================================
' WndWidth
'================================================
Public Property Let WndWidth(ByVal RHS As Single)
    If RHS = mWndWidth Then Exit Property
    WriteIniString "WndWidth", CStr(RHS)
    mWndWidth = RHS
End Property

'================================================
' WndHeight
'================================================
Public Property Get WndHeight() As Single
    WndHeight = mWndHeight
End Property

'================================================
' WndHeight
'================================================
Public Property Let WndHeight(ByVal RHS As Single)
    If RHS = mWndHeight Then Exit Property
    WriteIniString "WndHeight", CStr(RHS)
    mWndHeight = RHS
End Property
