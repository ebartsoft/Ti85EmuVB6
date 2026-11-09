VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTi85Layout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type LAYOUT_ITEM
    Left   As Single
    Top    As Single
    Width  As Single
    Height As Single
End Type

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private mItems() As LAYOUT_ITEM

'================================================
' Class_Initialize
'================================================
Private Sub Class_Initialize()

    ReDim mItems(51)
    
    mItems(0).Width = 411
    mItems(0).Height = 688
    
    mItems(1).Top = 16
    mItems(1).Left = 16
    mItems(1).Width = 377
    mItems(1).Height = 200
    
    mItems(2).Top = 224
    mItems(2).Left = 16
    mItems(2).Width = 72
    mItems(2).Height = 40
    
    mItems(3).Top = 224
    mItems(3).Left = 92
    mItems(3).Width = 72
    mItems(3).Height = 40
    
    mItems(4).Top = 224
    mItems(4).Left = 168
    mItems(4).Width = 72
    mItems(4).Height = 40
    
    mItems(5).Top = 224
    mItems(5).Left = 244
    mItems(5).Width = 72
    mItems(5).Height = 40
    
    mItems(6).Top = 224
    mItems(6).Left = 320
    mItems(6).Width = 72
    mItems(6).Height = 40
    
    mItems(7).Top = 280
    mItems(7).Left = 16
    mItems(7).Width = 72
    mItems(7).Height = 40
    
    mItems(8).Top = 280
    mItems(8).Left = 92
    mItems(8).Width = 72
    mItems(8).Height = 40
    
    mItems(9).Top = 280
    mItems(9).Left = 168
    mItems(9).Width = 72
    mItems(9).Height = 40
    
    mItems(10).Top = 296
    mItems(10).Left = 244
    mItems(10).Width = 48
    mItems(10).Height = 40
    
    mItems(11).Top = 296
    mItems(11).Left = 344
    mItems(11).Width = 48
    mItems(11).Height = 40
    
    mItems(12).Top = 272
    mItems(12).Left = 294
    mItems(12).Width = 48
    mItems(12).Height = 40
    
    mItems(13).Top = 320
    mItems(13).Left = 294
    mItems(13).Width = 48
    mItems(13).Height = 40
    
    mItems(14).Top = 324
    mItems(14).Left = 16
    mItems(14).Width = 72
    mItems(14).Height = 40
    
    mItems(15).Top = 324
    mItems(15).Left = 92
    mItems(15).Width = 72
    mItems(15).Height = 40
    
    mItems(16).Top = 324
    mItems(16).Left = 168
    mItems(16).Width = 72
    mItems(16).Height = 40
    
    mItems(17).Top = 368
    mItems(17).Left = 16
    mItems(17).Width = 72
    mItems(17).Height = 40
    
    mItems(18).Top = 368
    mItems(18).Left = 92
    mItems(18).Width = 72
    mItems(18).Height = 40
    
    mItems(19).Top = 368
    mItems(19).Left = 168
    mItems(19).Width = 72
    mItems(19).Height = 40
    
    mItems(20).Top = 368
    mItems(20).Left = 244
    mItems(20).Width = 72
    mItems(20).Height = 40
    
    mItems(21).Top = 368
    mItems(21).Left = 320
    mItems(21).Width = 72
    mItems(21).Height = 40
    
    mItems(22).Top = 412
    mItems(22).Left = 16
    mItems(22).Width = 72
    mItems(22).Height = 40
    
    mItems(23).Top = 412
    mItems(23).Left = 92
    mItems(23).Width = 72
    mItems(23).Height = 40
    
    mItems(24).Top = 412
    mItems(24).Left = 168
    mItems(24).Width = 72
    mItems(24).Height = 40
    
    mItems(25).Top = 412
    mItems(25).Left = 244
    mItems(25).Width = 72
    mItems(25).Height = 40
    
    mItems(26).Top = 412
    mItems(26).Left = 320
    mItems(26).Width = 72
    mItems(26).Height = 40
    
    mItems(27).Top = 456
    mItems(27).Left = 16
    mItems(27).Width = 72
    mItems(27).Height = 40
    
    mItems(28).Top = 456
    mItems(28).Left = 92
    mItems(28).Width = 72
    mItems(28).Height = 40
    
    mItems(29).Top = 456
    mItems(29).Left = 168
    mItems(29).Width = 72
    mItems(29).Height = 40
    
    mItems(30).Top = 456
    mItems(30).Left = 244
    mItems(30).Width = 72
    mItems(30).Height = 40
    
    mItems(31).Top = 456
    mItems(31).Left = 320
    mItems(31).Width = 72
    mItems(31).Height = 40
    
    mItems(32).Top = 500
    mItems(32).Left = 16
    mItems(32).Width = 72
    mItems(32).Height = 40
    
    mItems(33).Top = 500
    mItems(33).Left = 92
    mItems(33).Width = 72
    mItems(33).Height = 40
    
    mItems(34).Top = 500
    mItems(34).Left = 168
    mItems(34).Width = 72
    mItems(34).Height = 40
    
    mItems(35).Top = 500
    mItems(35).Left = 244
    mItems(35).Width = 72
    mItems(35).Height = 40
    
    mItems(36).Top = 500
    mItems(36).Left = 320
    mItems(36).Width = 72
    mItems(36).Height = 40
    
    mItems(37).Top = 544
    mItems(37).Left = 16
    mItems(37).Width = 72
    mItems(37).Height = 40
    
    mItems(38).Top = 544
    mItems(38).Left = 92
    mItems(38).Width = 72
    mItems(38).Height = 40
    
    mItems(39).Top = 544
    mItems(39).Left = 168
    mItems(39).Width = 72
    mItems(39).Height = 40
    
    mItems(40).Top = 544
    mItems(40).Left = 244
    mItems(40).Width = 72
    mItems(40).Height = 40
    
    mItems(41).Top = 544
    mItems(41).Left = 320
    mItems(41).Width = 72
    mItems(41).Height = 40
    
    mItems(42).Top = 588
    mItems(42).Left = 16
    mItems(42).Width = 72
    mItems(42).Height = 40
    
    mItems(43).Top = 588
    mItems(43).Left = 92
    mItems(43).Width = 72
    mItems(43).Height = 40
    
    mItems(44).Top = 588
    mItems(44).Left = 168
    mItems(44).Width = 72
    mItems(44).Height = 40
    
    mItems(45).Top = 588
    mItems(45).Left = 244
    mItems(45).Width = 72
    mItems(45).Height = 40
    
    mItems(46).Top = 588
    mItems(46).Left = 320
    mItems(46).Width = 72
    mItems(46).Height = 40
    
    mItems(47).Top = 632
    mItems(47).Left = 16
    mItems(47).Width = 72
    mItems(47).Height = 40
    
    mItems(48).Top = 632
    mItems(48).Left = 92
    mItems(48).Width = 72
    mItems(48).Height = 40
    
    mItems(49).Top = 632
    mItems(49).Left = 168
    mItems(49).Width = 72
    mItems(49).Height = 40
    
    mItems(50).Top = 632
    mItems(50).Left = 244
    mItems(50).Width = 72
    mItems(50).Height = 40
    
    mItems(51).Top = 632
    mItems(51).Left = 320
    mItems(51).Width = 72
    mItems(51).Height = 40

End Sub

'================================================
' Apply
'================================================
Public Sub Apply()
    Dim vBtn          As VBControlExtender 'UCTi85Btn
    Dim vScr          As VBControlExtender 'UCTi85LCD
    Dim vForm         As frmMain
    Dim vClientWidth  As Single
    Dim vClientHeight As Single
    Dim vLayoutWidth  As Single
    Dim vLayoutHeight As Single
    Dim vCenterLeft   As Single
    Dim vCenterTop    As Single
    Dim vRatio        As Single
    Dim i             As Long
    
    On Error Resume Next
    
    Set vForm = frmMain
    vClientWidth = vForm.ScaleWidth
    vClientHeight = vForm.ScaleHeight
    
    vRatio = MagicRatio(mItems(0).Width, mItems(0).Height, vClientWidth, vClientHeight)
    
    vLayoutWidth = mItems(0).Width * vRatio
    vLayoutHeight = mItems(0).Height * vRatio
    
    vCenterLeft = (vClientWidth - vLayoutWidth) * 0.5
    vCenterTop = (vClientHeight - vLayoutHeight) * 0.5
    
    Set vScr = vForm.UCTi85LCD1
    vScr.Move _
    vCenterLeft + mItems(1).Left * vRatio, _
    vCenterTop + mItems(1).Top * vRatio, _
    mItems(1).Width * vRatio, _
    mItems(1).Height * vRatio
    
    For i = 0 To vForm.cmdKey.UBound
        Set vBtn = vForm.cmdKey(i)
        vBtn.Move _
        vCenterLeft + mItems(2 + i).Left * vRatio, _
        vCenterTop + mItems(2 + i).Top * vRatio, _
        mItems(2 + i).Width * vRatio, _
        mItems(2 + i).Height * vRatio
    Next
    
End Sub

'================================================
' Resize
'================================================
Public Sub Resize(ByVal pRatio As Single)
    Dim vWidth  As Single
    Dim vHeight As Single
    vWidth = 1 + Int(pRatio * (mItems(0).Width / (mItems(1).Height / 72)) + 0.5)
    vHeight = 1 + Int(pRatio * (mItems(0).Height / (mItems(1).Height / 72)) + 0.5)
    ResizeWindow frmMain, vWidth, vHeight
End Sub

'================================================
' MagicRatio
'================================================
Private Function MagicRatio(ByVal pDstW As Single, ByVal pDstH As Single, ByVal pSrcW As Single, ByVal pSrcH As Single) As Single
    Dim a As Single
    Dim b As Single
    If pDstW = 0 Then Exit Function
    If pDstH = 0 Then Exit Function
    a = pSrcW / pDstW
    b = pSrcH / pDstH
    If a < b Then
        MagicRatio = a
    Else
        MagicRatio = b
    End If
End Function

'================================================
' ResizeWindow
'================================================
Private Function ResizeWindow(ByRef pForm As Form, ByVal pWidth As Long, ByVal pHeight As Long) As Boolean
    Dim vMarginX As Long
    Dim vMarginY As Long
    On Error GoTo Xe
    vMarginX = pForm.ScaleX(pForm.Width, vbTwips, vbPixels) - pForm.ScaleX(pForm.ScaleWidth, pForm.ScaleMode, vbPixels)
    vMarginY = pForm.ScaleY(pForm.Height, vbTwips, vbPixels) - pForm.ScaleY(pForm.ScaleHeight, pForm.ScaleMode, vbPixels)
    'ResizeWindow = SetWindowPos(pForm.hWnd, HWND_TOP, 0, 0, pWidth + vMarginX, pHeight + vMarginY, SWP_NOZORDER Or SWP_NOMOVE)
    ResizeWindow = SetWindowPos(pForm.hWnd, 0, 0, 0, pWidth + vMarginX, pHeight + vMarginY, 4 Or 2)
Xi: Exit Function
Xe: Resume Xi
End Function

''================================================
'' SaveLayout
''================================================
'Public Sub SaveLayout(ByRef pFileName As String)
'    Dim vBtn  As VBControlExtender 'UCTi85Btn
'    Dim vScr  As VBControlExtender 'UCTi85LCD
'    Dim vForm As frmMain
'    Dim vText As String
'    Dim i     As Long
'
'    Set vForm = frmMain
'    vText = vText & "    mItems(0).Width = " & vForm.ScaleWidth & vbCrLf
'    vText = vText & "    mItems(0).Height = " & vForm.ScaleHeight & vbCrLf
'    vText = vText & vbCrLf
'
'    Set vScr = vForm.UCTi85LCD1
'    vText = vText & "    mItems(1).Top = " & vScr.Top & vbCrLf
'    vText = vText & "    mItems(1).Left = " & vScr.Left & vbCrLf
'    vText = vText & "    mItems(1).Width = " & vScr.Width & vbCrLf
'    vText = vText & "    mItems(1).Height = " & vScr.Height & vbCrLf
'    vText = vText & vbCrLf
'
'    For i = 0 To vForm.cmdKey.UBound
'        Set vBtn = vForm.cmdKey(i)
'        vText = vText & "    mItems(" & i + 2 & ").Top = " & vBtn.Top & vbCrLf
'        vText = vText & "    mItems(" & i + 2 & ").Left = " & vBtn.Left & vbCrLf
'        vText = vText & "    mItems(" & i + 2 & ").Width = " & vBtn.Width & vbCrLf
'        vText = vText & "    mItems(" & i + 2 & ").Height = " & vBtn.Height & vbCrLf
'        vText = vText & vbCrLf
'    Next
'
'    Open pFileName For Output As #1
'    Print #1, vText
'    Close #1
'
'End Sub
