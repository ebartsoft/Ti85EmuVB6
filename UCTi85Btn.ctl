VERSION 5.00
Begin VB.UserControl UCTi85Btn 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H0045484D&
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1095
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00BECFDC&
   ScaleHeight     =   41
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   73
End
Attribute VB_Name = "UCTi85Btn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum Ti85ButtonTypeEnum
    tiBtnBlack
    tiBtnYellow
    tiBtnBlue
    tiBtnGray
End Enum

Private Type POINTAPI
    x As Long
    y As Long
End Type

' The Win32 API declaration for DPtoLP
Private Declare Function DPtoLP Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByRef lpPoint As POINTAPI, _
    ByVal nCount As Long) As Long
    
' --- Windows API Declarations ---
Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Private Const PI            As Double = 3.14159265358979
Private Const DIPI          As Double = PI * 2
Private Const MIPI          As Double = PI / 2

' Pen Styles for ExtCreatePen
Private Const PS_GEOMETRIC  As Long = &H1000
Private Const PS_SOLID      As Long = 0

' Join Styles (The Corner Shapes)
Private Const PS_JOIN_ROUND As Long = &H1000  ' Rounded corners
Private Const PS_JOIN_MITER As Long = &H2000  ' Sharp, mitered corners
Private Const PS_JOIN_BEVEL As Long = &H10000 ' Flat, beveled corners

' End Cap Styles
Private Const PS_ENDCAP_ROUND As Long = &H0
Private Const PS_ENDCAP_SQUARE As Long = &H100
Private Const PS_ENDCAP_FLAT As Long = &H200

Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Private Declare Function Polygon Lib "gdi32" (ByVal pHdc As Long, ByRef lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, ByRef lplb As LOGBRUSH, ByVal dwStyleCount As Long, ByVal lpStyle As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BeginPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function StrokePath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FillPath Lib "gdi32" (ByVal pHdc As Long) As Long

Private mBtnType        As Ti85ButtonTypeEnum
Private mCaption2nd     As String
Private mCaptionAlpha   As String
Private mCaption        As String
Private mKeyCode        As Long
Private mBtnState       As Long

'================================================
' UserControl_Initialize
'================================================
Private Sub UserControl_Initialize()
    'NTD
End Sub

'================================================
' UserControl_InitProperties
'================================================
Private Sub UserControl_InitProperties()
    mCaption = "SIN"
    mCaption2nd = "SIN-1"
    mCaptionAlpha = "B"
    mBtnType = tiBtnBlack
    mKeyCode = 32
End Sub

'================================================
' UserControl_WriteProperties
'================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", mCaption
    PropBag.WriteProperty "Caption2nd", mCaption2nd
    PropBag.WriteProperty "CaptionAlpha", mCaptionAlpha
    PropBag.WriteProperty "BtnType", mBtnType
    PropBag.WriteProperty "KeyCode", mKeyCode
End Sub

'================================================
' UserControl_ReadProperties
'================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mCaption = PropBag.ReadProperty("Caption", "SIN")
    mCaption2nd = PropBag.ReadProperty("Caption2nd", "SIN-1")
    mCaptionAlpha = PropBag.ReadProperty("CaptionAlpha", "B")
    mBtnType = PropBag.ReadProperty("BtnType", tiBtnBlack)
    mKeyCode = PropBag.ReadProperty("KeyCode", 32)
End Sub

'================================================
' UserControl_MouseDown
'================================================
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
    mBtnState = 1
    Redraw
    'Debug.Print "UserControl_MouseDown"
End Sub

'================================================
' UserControl_MouseUp
'================================================
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
    mBtnState = 2
    Redraw
    'Debug.Print "UserControl_MouseUp"
End Sub

'================================================
' UserControl_DblClick
'================================================
Private Sub UserControl_DblClick()
    UserControl_MouseDown 0, 0, 0, 0
    'Debug.Print "UserControl_DblClick"
End Sub

'================================================
' KeyCode
'================================================
Public Property Get KeyCode() As Long
    KeyCode = mKeyCode
End Property

'================================================
' KeyCode
'================================================
Public Property Let KeyCode(ByVal RHS As Long)
    mKeyCode = RHS
End Property

'================================================
' BtnType
'================================================
Public Property Get BtnType() As Ti85ButtonTypeEnum
    BtnType = mBtnType
End Property

'================================================
' BtnType
'================================================
Public Property Let BtnType(ByVal RHS As Ti85ButtonTypeEnum)
    mBtnType = RHS
    RedrawIfVisible
End Property

'================================================
' Caption
'================================================
Public Property Get Caption() As String
    Caption = mCaption
End Property

'================================================
' Caption
'================================================
Public Property Let Caption(ByRef RHS As String)
    mCaption = RHS
    RedrawIfVisible
End Property
   
'================================================
' Caption2nd
'================================================
Public Property Get Caption2nd() As String
    Caption2nd = mCaption2nd
End Property

'================================================
' Caption2nd
'================================================
Public Property Let Caption2nd(ByRef RHS As String)
    mCaption2nd = RHS
    RedrawIfVisible
End Property
  
'================================================
' CaptionAlpha
'================================================
Public Property Get CaptionAlpha() As String
    CaptionAlpha = mCaptionAlpha
End Property

'================================================
' CaptionAlpha
'================================================
Public Property Let CaptionAlpha(ByRef RHS As String)
    mCaptionAlpha = RHS
    RedrawIfVisible
End Property
  
'================================================
' UserControl_Resize
'================================================
Private Sub UserControl_Resize()
    SetFontSize
    RedrawIfVisible
End Sub

'================================================
' UserControl_Show
'================================================
Private Sub UserControl_Show()
    RedrawIfVisible
End Sub

'================================================
' SetFontSize
'================================================
Private Sub SetFontSize()
    On Error Resume Next
    FontSize = GetScale() * 7
End Sub

'================================================
' RedrawIfVisible
'================================================
Private Sub RedrawIfVisible()
    If IsWindowVisible(hWnd) = 0 Then Exit Sub
    Redraw
End Sub

'================================================
' GetScale
'================================================
Private Function GetScale() As Single
    Dim vScaleX As Single
    Dim vScaleY As Single
    'vScaleY = ScaleHeight / 41
    vScaleX = ScaleWidth / 73
    'If vScaleX > vScaleY Then
        GetScale = vScaleX
    'Else
    '    GetScale = vScaleY
    'End If
End Function

'================================================
' Redraw
'================================================
Public Sub Redraw()
    Dim vTheme     As CTi85Theme
    Dim vScale     As Single
    Dim vScaleY    As Single
    Dim vScaleX    As Single
    Dim vRound     As Single
    Dim vFillColor As Long
    
    Set vTheme = MainTheme
    
    vScale = GetScale()
    vScale = Int(vScale * 2)
    If vScale < 1 Then vScale = 1
    
    vRound = ScaleWidth / 73
    vRound = Int(vRound * 8)
    If vRound < 1 Then vRound = 1
    
    AutoRedraw = True
    
    Line (0, 0)-(ScaleWidth, ScaleHeight), vTheme.CalculatorBackColor, BF
    
    Select Case mBtnType
    Case Ti85ButtonTypeEnum.tiBtnBlack:  vFillColor = vTheme.BtnBlackBackColor
    Case Ti85ButtonTypeEnum.tiBtnYellow: vFillColor = vTheme.BtnYellowBackColor
    Case Ti85ButtonTypeEnum.tiBtnBlue:   vFillColor = vTheme.BtnBlueBackColor
    Case Ti85ButtonTypeEnum.tiBtnGray:   vFillColor = vTheme.BtnGrayBackColor
    End Select
    
    DrawText 4, Int(ScaleHeight / 3 / 2), mCaption2nd, vTheme.BtnCaption2nd, 4
    DrawText ScaleWidth - 4, Int(ScaleHeight / 3 / 2), mCaptionAlpha, vTheme.BtnCaptionAlpha, 4 Or 2
    If mBtnState = 1 Then
        DrawRoundRect 2, 2 + Int(ScaleHeight / 3), ScaleWidth, ScaleHeight, vRound, vFillColor, vTheme.BtnBorderPressed, vScale
        DrawArrow Int(ScaleWidth / 2) + 1, Int(ScaleHeight / 3) + Int(ScaleHeight * (2 / 3 / 2)) + 1, mCaption, vTheme.BtnCaption, 1 Or 4, vRound
    Else
        DrawRoundRect 2, 2 + Int(ScaleHeight / 3), ScaleWidth, ScaleHeight, vRound, vFillColor, vTheme.BtnBorderReleased, vScale
        DrawArrow Int(ScaleWidth / 2), Int(ScaleHeight / 3) + Int(ScaleHeight * (2 / 3 / 2)), mCaption, vTheme.BtnCaption, 1 Or 4, vRound
    End If
    
    AutoRedraw = False
    
End Sub

'================================================
' DrawText
'================================================
Private Sub DrawText(ByVal pX As Long, ByVal pY As Long, ByRef pText As String, ByVal pColor As Long, Optional ByVal pAlign As Long = 0)
    ForeColor = pColor
    If (pAlign And 1) = 1 Then
        pX = pX - Int(TextWidth(pText) / 2)
    ElseIf (pAlign And 2) = 2 Then
        pX = pX - TextWidth(pText)
    End If
    If (pAlign And 4) = 4 Then
        pY = pY - Int(TextHeight(pText) / 2)
    ElseIf (pAlign And 8) = 8 Then
        pY = pY - TextHeight(pText)
    End If
    TextOut hDC, pX, pY, pText, Len(pText)
End Sub

'================================================
' DrawArrow
'================================================
Private Sub DrawArrow(ByVal pX As Long, ByVal pY As Long, ByRef pText As String, ByVal pColor As Long, ByVal pAlign As Long, ByVal pRound As Single)
    Dim p(2)    As POINTAPI
    Dim vPhase  As Single
    Dim vMid    As Long
    
    Select Case pText
    Case "a^": vPhase = MIPI * 0
    Case "a>": vPhase = MIPI * 1
    Case "av": vPhase = MIPI * 2
    Case "a<": vPhase = MIPI * 3
    Case Else: DrawText pX, pY, pText, pColor, pAlign: Exit Sub
    End Select
    
    vPhase = vPhase - MIPI
    vMid = Int(pRound / 2) * 3
    
    p(0).x = pX + Cos(vPhase) * vMid
    p(0).y = pY + Sin(vPhase) * vMid
    p(1).x = pX + Cos(DIPI / 3 + vPhase) * vMid
    p(1).y = pY + Sin(DIPI / 3 + vPhase) * vMid
    p(2).x = pX + Cos(DIPI / 3 * 2 + vPhase) * vMid
    p(2).y = pY + Sin(DIPI / 3 * 2 + vPhase) * vMid
    
    FillColor = pColor
    FillStyle = vbFSSolid
    
    BeginPath hDC
    Polygon hDC, p(0), 3
    EndPath hDC
    FillPath hDC
    
    FillStyle = vbFSTransparent

End Sub

'================================================
' DrawRoundRect
'================================================
Private Sub DrawRoundRect(ByVal pX1 As Long, ByVal pY1 As Long, ByVal pX2 As Long, ByVal pY2 As Long, ByVal pRound As Long, ByVal pFillColor As Long, ByVal pForeColor As Long, ByVal pWidth As Single)
    FillStyle = FillStyleConstants.vbFSSolid
    ForeColor = pForeColor
    FillColor = pFillColor
    DrawWidth = pWidth
    
    pRound = 0
    
''    Dim brushInfo   As LOGBRUSH
''    Dim hPen        As Long
''    Dim hOldPen     As Long
''
''    ' Set up brush info for the geometric pen
''    brushInfo.lbStyle = 0 ' Solid brush
''    brushInfo.lbHatch = 0
''    brushInfo.lbColor = pForeColor
''
''    'Dim vPt As POINTAPI
''    'vPt.X = 3
''    'Debug.Print DPtoLP(hDC, vPt, 1)
''
''    ' Create the custom pen combining style, cap, and corner join properties
''    'hPen = ExtCreatePen(PS_GEOMETRIC Or PS_SOLID Or PS_JOIN_ROUND Or PS_ENDCAP_ROUND, 3, brushInfo, 0, 0)
''    hPen = CreatePen(PS_GEOMETRIC Or PS_SOLID Or PS_JOIN_ROUND Or PS_ENDCAP_ROUND, 3, pForeColor)
''    hOldPen = SelectObject(hDC, hPen)
    
    'BeginPath hdc
    RoundRect hDC, pX1, pY1, pX2 - pWidth, pY2 - pWidth, pRound, pRound
    'EndPath hdc
    'StrokePath hdc
    
''    SelectObject hDC, hOldPen
''    DeleteObject hPen
    
    FillStyle = FillStyleConstants.vbFSTransparent
    DrawWidth = 1

    'Dim i As Long
    'For i = 0 To pWidth - 1
    '    RoundRect hdc, pX1 + i, pY1 + i, pX2 - i, pY2 - i, pRound, pRound
    'Next

End Sub

