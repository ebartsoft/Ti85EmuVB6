VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTi85Theme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public CalculatorBackColor  As Long
Public ScreenBorderColor    As Long
Public BtnBlackBackColor    As Long
Public BtnYellowBackColor   As Long
Public BtnBlueBackColor     As Long
Public BtnGrayBackColor     As Long
Public BtnBorderReleased    As Long
Public BtnBorderPressed     As Long
Public BtnCaption           As Long
Public BtnCaption2nd        As Long
Public BtnCaptionAlpha      As Long
Public BkgColorOn           As Long
Public BkgColorOff          As Long
Public FntColorOn           As Long
Public FntColorOff          As Long
Public BdrColor             As Long

'================================================
' Class_Initialize
'================================================
Private Sub Class_Initialize()
    SetDarkMode
    SetLCDBlueOnGreen
End Sub

'================================================
' SetLCDGreenOnBlack
'================================================
Public Sub SetLCDGreenOnBlack()
    BkgColorOn = &HFF00&
    BkgColorOff = &H0
    FntColorOn = &HFF00&
    FntColorOff = &H0
    BdrColor = &H0
End Sub

'================================================
' SetLCDBlueOnGreen
'================================================
Public Sub SetLCDBlueOnGreen()
    BkgColorOn = &H7B634A
    BkgColorOff = &HB0CCAE
    FntColorOn = &H7B634A
    FntColorOff = &HB0CCAE
    BdrColor = &HB0CCAE
End Sub

'================================================
' SetLCDBrownOnOrange
'================================================
Public Sub SetLCDBrownOnOrange()
    BkgColorOn = &H102951
    BkgColorOff = &H3890F8
    FntColorOn = &H102951
    FntColorOff = &H3890F8
    BdrColor = &H3890F8
End Sub

'================================================
' SetLCDBlackOnWhite
'================================================
Public Sub SetLCDBlackOnWhite()
    BkgColorOn = 0
    BkgColorOff = &HFFFFFF
    FntColorOn = 0
    FntColorOff = &HFFFFFF
    BdrColor = &HFFFFFF
End Sub

'================================================
' SetLCDWhiteOnBlack
'================================================
Public Sub SetLCDWhiteOnBlack()
    BkgColorOn = &HFFFFFF
    BkgColorOff = 0
    FntColorOn = &HFFFFFF
    FntColorOff = &HFFFFFF
    BdrColor = 0
End Sub

'================================================
' SetDarkMode
'================================================
Public Sub SetDarkMode()
    CalculatorBackColor = &H43413E
    ScreenBorderColor = 0
    BtnBlackBackColor = &H272A28
    BtnYellowBackColor = &H88BCF1
    BtnBlueBackColor = &HBBA099
    BtnGrayBackColor = &H8F8F95
    BtnBorderReleased = &H0
    BtnBorderPressed = &H808080
    BtnCaption = &HF0F0F0
    BtnCaption2nd = &HBECFDC
    BtnCaptionAlpha = &HBBA099
End Sub

'================================================
' SetLightMode
'================================================
Public Sub SetLightMode()
    CalculatorBackColor = &HF0F0F0
    ScreenBorderColor = 0
    BtnBlackBackColor = &HFFFFFF
    BtnYellowBackColor = &H88BCF1
    BtnBlueBackColor = &HBBA099
    BtnGrayBackColor = &HBBBBBB
    BtnBorderReleased = &H808080
    BtnBorderPressed = &HBFFF&
    BtnCaption = &H333333
    BtnCaption2nd = &H8AC4&
    BtnCaptionAlpha = &HBB7664
End Sub

'================================================
' SetThemeByName
'================================================
Public Sub SetThemeByName(ByRef pName As String)
    Select Case pName
    Case "DARK":  SetDarkMode
    Case "LIGHT": SetLightMode
    End Select
End Sub

'================================================
' SetCLDByName
'================================================
Public Sub SetLCDByName(ByRef pName As String)
    Select Case pName
    Case "BLUEONGREEN":   SetLCDBlueOnGreen
    Case "BROWNONORANGE": SetLCDBrownOnOrange
    Case "GREENONBLACK":  SetLCDGreenOnBlack
    Case "BLACKONWHITE":  SetLCDBlackOnWhite
    Case "WHITEONBLACK":  SetLCDWhiteOnBlack
    End Select
End Sub
