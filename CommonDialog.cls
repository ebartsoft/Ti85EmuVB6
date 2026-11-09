VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'
' O    O    O
'  \__/ \__/
'  /=||=||=\   oouuuunnnnnnnnmmmmmmmmmmmmmm\
' // ||_||                 DrBacter®        \
' \\ /\ #\     oouuuunnnnnnnnmmmmmmmmmmmmmmmm\
' /=(  \  )==>       Coded by EBArtSoft@      \
'//  \O_\/            Copyright © 2004         \
'\\  || ||        email ebartsoft@hotmail.com   \
' \==||=||==/  oouuuunnnnnnnnmmmmmmmmmmmmmmmmmmmm\
' ===========
'==== E.B ====
'
' ALL RIGHTS RESERVED ::..
' Permission  to  use,  copy,  modify,  and  distribute this software for
' any  purpose and  without  fee  is  hereby  granted,  provided that the
' above copyright notice appear in all copies and that both the copyright
' notice  and  this permission notice appear in supporting documentation.
'
' THE  MATERIAL  EMBODIED  ON  THIS  SOFTWARE IS PROVIDED TO YOU "AS-IS"
' AND  WITHOUT  WARRANTY  OF  ANY  KIND,  EXPRESS, IMPLIED OR OTHERWISE,
' INCLUDING  WITHOUT  LIMITATION,  ANY  WARRANTY  OF  MERCHANTABILITY OR
' FITNESS  FOR  A  PARTICULAR  PURPOSE.  IN  NO EVENT SHALL WE BE LIABLE
' TO  YOU  OR  ANYONE ELSE FOR ANY DIRECT, SPECIAL, INCIDENTAL, INDIRECT
' OR  CONSEQUENTIAL  DAMAGES  OF  ANY  KIND,  OR ANY DAMAGES WHATSOEVER,
' INCLUDING  WITHOUT  LIMITATION,  LOSS  OF PROFIT, LOSS OF USE, SAVINGS
' OR REVENUE, OR THE CLAIMS OF THIRD PARTIES, WHETHER OR NOT WE HAS BEEN
' ADVISED  OF  THE  POSSIBILITY  OF  SUCH  LOSS,  HOWEVER  CAUSED AND ON
' ANY  THEORY  OF  LIABILITY,  ARISING  OUT OF OR IN CONNECTION WITH THE
' POSSESSION, USE OR PERFORMANCE OF THIS SOFTWARE.
'
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Enum OFN_Flags
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000&
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
    ' Win98/NT5 only...
    OFN_ENABLEINCLUDENOTIFY = &H400000
    OFN_ENABLESIZING = &H800000
End Enum

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As OFN_Flags
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const dv_FileName = ""
Private Const dv_FilterIndex = 0
Private Const dv_FileTitle = ""
Private Const dv_Filter = ""
Private Const dv_InitDir = ""
Private Const dv_Flags = 0
Private Const dv_MaxFileSize = 260
Private Const dv_Parent = 0
Private Const dv_DefaultExt = ""

Dim mv_FilterIndex As Integer
Dim mv_MaxFileSize As Integer
Dim mv_FileTitle   As String
Dim mv_Filter      As String
Dim mv_InitDir     As String
Dim mv_DialogTitle As String
Dim mv_FileName    As String
Dim mv_DefaultExt  As String
Dim mv_Flags       As Long
Dim mv_Parent      As Long

Private Sub Class_Initialize()
    mv_FileName = dv_FileName
    mv_FilterIndex = dv_FilterIndex
    mv_FileTitle = dv_FileTitle
    mv_Filter = dv_Filter
    mv_MaxFileSize = dv_MaxFileSize
End Sub

Public Property Get DefaultExt() As String
    DefaultExt = mv_DefaultExt
End Property

Public Property Let DefaultExt(ByVal vNewValue As String)
    mv_DefaultExt = vNewValue
End Property

Public Property Get Parent() As Long
    Parent = mv_Parent
End Property

Public Property Let Parent(ByVal vNewValue As Long)
    mv_Parent = vNewValue
End Property

Public Property Get MaxFileSize() As Integer
    MaxFileSize = mv_MaxFileSize
End Property

Public Property Let MaxFileSize(ByVal vNewValue As Integer)
    mv_MaxFileSize = Max(dv_MaxFileSize, vNewValue)
End Property

Public Property Get Flags() As OFN_Flags
    Flags = mv_Flags
End Property

Public Property Let Flags(ByVal vNewValue As OFN_Flags)
    mv_Flags = vNewValue
End Property

Public Property Get DialogTitle() As String
    DialogTitle = mv_DialogTitle
End Property

Public Property Let DialogTitle(ByVal vNewValue As String)
    mv_DialogTitle = vNewValue
End Property

Public Property Get InitDir() As String
    InitDir = mv_InitDir
End Property

Public Property Let InitDir(ByVal vNewValue As String)
    mv_InitDir = vNewValue
End Property

Public Property Get Filter() As String
    Filter = mv_Filter
End Property

Public Property Let Filter(ByVal vNewValue As String)
    mv_Filter = vNewValue
End Property

Public Property Get FilterIndex() As Integer
    FilterIndex = mv_FilterIndex
End Property

Public Property Let FilterIndex(ByVal vNewValue As Integer)
    mv_FilterIndex = vNewValue
End Property

Public Property Get FileName() As String
    FileName = mv_FileName
End Property

Public Property Let FileName(ByVal vNewValue As String)
    mv_FileName = vNewValue
End Property

Public Property Get FileTitle() As String
    FileTitle = mv_FileTitle
End Property

Function ShowSave() As Boolean
    Dim ofn As OPENFILENAME
    With ofn
        .lStructSize = Len(ofn)
        .Flags = mv_Flags
        .hWndOwner = mv_Parent
        .nMaxFile = mv_MaxFileSize
        .lpstrDefExt = mv_DefaultExt
        .nFilterIndex = mv_FilterIndex
        .lpstrInitialDir = mv_InitDir & vbNullChar
        .lpstrTitle = mv_DialogTitle & vbNullChar
        .lpstrFile = mv_FileName & String$(Max(mv_MaxFileSize - Len(mv_FileName), 0), vbNullChar)
        .lpstrFilter = Replace(mv_Filter, "|", vbNullChar) & vbNullChar & vbNullChar
    End With
    If GetSaveFileName(ofn) Then
        mv_FileName = Left(ofn.lpstrFile, lstrlen(ofn.lpstrFile))
        mv_FilterIndex = ofn.nFilterIndex
        mv_FileTitle = Mid$(ofn.lpstrFile, ofn.nFileOffset + 1, InStr(ofn.lpstrFile, vbNullChar) - (ofn.nFileOffset + 1))
        ShowSave = True
    End If
End Function

Function ShowOpen() As Boolean
    Dim ofn As OPENFILENAME
    With ofn
        .lStructSize = Len(ofn)
        .Flags = mv_Flags
        .hWndOwner = mv_Parent
        .nMaxFile = mv_MaxFileSize
        .lpstrDefExt = mv_DefaultExt
        .nFilterIndex = mv_FilterIndex
        .lpstrInitialDir = mv_InitDir & vbNullChar
        .lpstrTitle = mv_DialogTitle & vbNullChar
        .lpstrFile = mv_FileName & String$(Max(mv_MaxFileSize - Len(mv_FileName), 0), vbNullChar)
        .lpstrFilter = Replace(mv_Filter, "|", vbNullChar) & vbNullChar & vbNullChar
    End With
    If GetOpenFileName(ofn) Then
        mv_FileName = Left(ofn.lpstrFile, lstrlen(ofn.lpstrFile))
        mv_FilterIndex = ofn.nFilterIndex
        mv_FileTitle = Mid$(ofn.lpstrFile, ofn.nFileOffset + 1, InStr(ofn.lpstrFile, vbNullChar) - (ofn.nFileOffset + 1))
        ShowOpen = True
    End If
End Function

Private Function Max(v1, v2)
    Max = IIf(v1 > v2, v1, v2)
End Function
