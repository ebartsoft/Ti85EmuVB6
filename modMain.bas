Attribute VB_Name = "modMain"
Option Explicit

Private mMainConfig As CTi85Config
Private mMainLayout As CTi85Layout
Private mMainTheme  As CTi85Theme

'================================================
' Main
'================================================
Public Sub Main()
    frmMain.RunEmulator
End Sub

'================================================
' MainConfig
'================================================
Public Property Get MainConfig() As CTi85Config
    If mMainConfig Is Nothing Then
        Set mMainConfig = New CTi85Config
    End If
    Set MainConfig = mMainConfig
End Property

'================================================
' MainLayout
'================================================
Public Property Get MainLayout() As CTi85Layout
    If mMainLayout Is Nothing Then
        Set mMainLayout = New CTi85Layout
    End If
    Set MainLayout = mMainLayout
End Property

'================================================
' MainTheme
'================================================
Public Property Get MainTheme() As CTi85Theme
    If mMainTheme Is Nothing Then
        Set mMainTheme = New CTi85Theme
        mMainTheme.SetThemeByName MainConfig.ThemeName
        mMainTheme.SetLCDByName MainConfig.LCDName
    End If
    Set MainTheme = mMainTheme
End Property
