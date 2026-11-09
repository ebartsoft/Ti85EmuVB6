VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H0043413E&
   Caption         =   "TI-85 Emulator"
   ClientHeight    =   10320
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   6165
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   688
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   411
   Begin VB.Timer tmrKeys 
      Interval        =   30
      Left            =   120
      Top             =   120
   End
   Begin Ti85EmuVB6.UCTi85LCD UCTi85LCD1 
      Height          =   3000
      Left            =   240
      TabIndex        =   50
      Top             =   240
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5292
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   3360
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "F1"
      Caption2nd      =   "M1"
      CaptionAlpha    =   ""
      BtnType         =   0
      KeyCode         =   52
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   1
      Left            =   1380
      TabIndex        =   1
      Top             =   3360
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "F2"
      Caption2nd      =   "M2"
      CaptionAlpha    =   ""
      BtnType         =   0
      KeyCode         =   51
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   2
      Left            =   2520
      TabIndex        =   2
      Top             =   3360
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "F3"
      Caption2nd      =   "M3"
      CaptionAlpha    =   ""
      BtnType         =   0
      KeyCode         =   50
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   3
      Left            =   3660
      TabIndex        =   3
      Top             =   3360
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "F4"
      Caption2nd      =   "M4"
      CaptionAlpha    =   ""
      BtnType         =   0
      KeyCode         =   49
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   4
      Left            =   4800
      TabIndex        =   4
      Top             =   3360
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "F5"
      Caption2nd      =   "M5"
      CaptionAlpha    =   ""
      BtnType         =   0
      KeyCode         =   48
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   4200
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "2nd"
      Caption2nd      =   ""
      CaptionAlpha    =   ""
      BtnType         =   1
      KeyCode         =   53
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   6
      Left            =   1380
      TabIndex        =   6
      Top             =   4200
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "EXIT"
      Caption2nd      =   "QUIT"
      CaptionAlpha    =   ""
      BtnType         =   0
      KeyCode         =   54
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   7
      Left            =   2520
      TabIndex        =   7
      Top             =   4200
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "MORE"
      Caption2nd      =   "MODE"
      CaptionAlpha    =   ""
      BtnType         =   0
      KeyCode         =   55
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   10
      Left            =   4410
      TabIndex        =   10
      Top             =   4080
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      Caption         =   "a^"
      Caption2nd      =   ""
      CaptionAlpha    =   ""
      BtnType         =   3
      KeyCode         =   3
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   11
      Left            =   4410
      TabIndex        =   11
      Top             =   4800
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      Caption         =   "av"
      Caption2nd      =   ""
      CaptionAlpha    =   ""
      BtnType         =   3
      KeyCode         =   0
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   12
      Left            =   240
      TabIndex        =   12
      Top             =   4860
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "ALPHA"
      Caption2nd      =   "alpha"
      CaptionAlpha    =   ""
      BtnType         =   2
      KeyCode         =   47
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   13
      Left            =   1380
      TabIndex        =   13
      Top             =   4860
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "x-VAR"
      Caption2nd      =   "LINK"
      CaptionAlpha    =   "x"
      BtnType         =   0
      KeyCode         =   39
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   14
      Left            =   2520
      TabIndex        =   14
      Top             =   4860
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "DEL"
      Caption2nd      =   "INS"
      CaptionAlpha    =   ""
      BtnType         =   0
      KeyCode         =   31
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   15
      Left            =   240
      TabIndex        =   15
      Top             =   5520
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "GRAPH"
      Caption2nd      =   "SOLVER"
      CaptionAlpha    =   ""
      BtnType         =   0
      KeyCode         =   46
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   16
      Left            =   1380
      TabIndex        =   16
      Top             =   5520
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "STAT"
      Caption2nd      =   "SIMULT"
      CaptionAlpha    =   ""
      BtnType         =   0
      KeyCode         =   38
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   17
      Left            =   2520
      TabIndex        =   17
      Top             =   5520
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "PRGM"
      Caption2nd      =   "POLY"
      CaptionAlpha    =   ""
      BtnType         =   0
      KeyCode         =   30
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   18
      Left            =   3660
      TabIndex        =   18
      Top             =   5520
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "CUSTOM"
      Caption2nd      =   "CATALOG"
      CaptionAlpha    =   ""
      BtnType         =   0
      KeyCode         =   22
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   19
      Left            =   4800
      TabIndex        =   19
      Top             =   5520
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "CLEAR"
      Caption2nd      =   "TOLER"
      CaptionAlpha    =   ""
      BtnType         =   0
      KeyCode         =   14
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   20
      Left            =   240
      TabIndex        =   20
      Top             =   6180
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "LOG"
      Caption2nd      =   "10^x"
      CaptionAlpha    =   "A"
      BtnType         =   0
      KeyCode         =   45
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   21
      Left            =   1380
      TabIndex        =   21
      Top             =   6180
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "SIN"
      Caption2nd      =   "SIN^-1"
      CaptionAlpha    =   "B"
      BtnType         =   0
      KeyCode         =   37
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   22
      Left            =   2520
      TabIndex        =   22
      Top             =   6180
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "COS"
      Caption2nd      =   "COS^-1"
      CaptionAlpha    =   "C"
      BtnType         =   0
      KeyCode         =   29
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   23
      Left            =   3660
      TabIndex        =   23
      Top             =   6180
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "TAN"
      Caption2nd      =   "TAN^-1"
      CaptionAlpha    =   "D"
      BtnType         =   0
      KeyCode         =   21
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   24
      Left            =   4800
      TabIndex        =   24
      Top             =   6180
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "^"
      Caption2nd      =   "PI"
      CaptionAlpha    =   "E"
      BtnType         =   0
      KeyCode         =   13
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   25
      Left            =   240
      TabIndex        =   25
      Top             =   6840
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "LN"
      Caption2nd      =   "e^x"
      CaptionAlpha    =   "F"
      BtnType         =   0
      KeyCode         =   44
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   26
      Left            =   1380
      TabIndex        =   26
      Top             =   6840
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "EE"
      Caption2nd      =   "X^-1"
      CaptionAlpha    =   "G"
      BtnType         =   0
      KeyCode         =   36
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   27
      Left            =   2520
      TabIndex        =   27
      Top             =   6840
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "("
      Caption2nd      =   "["
      CaptionAlpha    =   "H"
      BtnType         =   0
      KeyCode         =   28
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   28
      Left            =   3660
      TabIndex        =   28
      Top             =   6840
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   ")"
      Caption2nd      =   "]"
      CaptionAlpha    =   "I"
      BtnType         =   0
      KeyCode         =   20
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   29
      Left            =   4800
      TabIndex        =   29
      Top             =   6840
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "/"
      Caption2nd      =   "CALC"
      CaptionAlpha    =   "J"
      BtnType         =   0
      KeyCode         =   12
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   30
      Left            =   240
      TabIndex        =   30
      Top             =   7500
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "x^2"
      Caption2nd      =   "SQRT"
      CaptionAlpha    =   "K"
      BtnType         =   0
      KeyCode         =   43
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   31
      Left            =   1380
      TabIndex        =   31
      Top             =   7500
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "7"
      Caption2nd      =   "MATRX"
      CaptionAlpha    =   "L"
      BtnType         =   3
      KeyCode         =   35
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   32
      Left            =   2520
      TabIndex        =   32
      Top             =   7500
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "8"
      Caption2nd      =   "VECTR"
      CaptionAlpha    =   "M"
      BtnType         =   3
      KeyCode         =   27
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   33
      Left            =   3660
      TabIndex        =   33
      Top             =   7500
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "9"
      Caption2nd      =   "CPLX"
      CaptionAlpha    =   "N"
      BtnType         =   3
      KeyCode         =   19
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   34
      Left            =   4800
      TabIndex        =   34
      Top             =   7500
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "*"
      Caption2nd      =   "MATH"
      CaptionAlpha    =   "O"
      BtnType         =   0
      KeyCode         =   11
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   35
      Left            =   240
      TabIndex        =   35
      Top             =   8160
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   ","
      Caption2nd      =   "ANGLE"
      CaptionAlpha    =   "P"
      BtnType         =   0
      KeyCode         =   42
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   36
      Left            =   1380
      TabIndex        =   36
      Top             =   8160
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "4"
      Caption2nd      =   "CONS"
      CaptionAlpha    =   "Q"
      BtnType         =   3
      KeyCode         =   34
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   37
      Left            =   2520
      TabIndex        =   37
      Top             =   8160
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "5"
      Caption2nd      =   "CONV"
      CaptionAlpha    =   "R"
      BtnType         =   3
      KeyCode         =   26
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   38
      Left            =   3660
      TabIndex        =   38
      Top             =   8160
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "6"
      Caption2nd      =   "STRNG"
      CaptionAlpha    =   "S"
      BtnType         =   3
      KeyCode         =   18
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   39
      Left            =   4800
      TabIndex        =   39
      Top             =   8160
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "-"
      Caption2nd      =   "LIST"
      CaptionAlpha    =   "T"
      BtnType         =   0
      KeyCode         =   10
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   40
      Left            =   240
      TabIndex        =   40
      Top             =   8820
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "STO"
      Caption2nd      =   "RCL"
      CaptionAlpha    =   "="
      BtnType         =   0
      KeyCode         =   41
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   41
      Left            =   1380
      TabIndex        =   41
      Top             =   8820
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "1"
      Caption2nd      =   "BASE"
      CaptionAlpha    =   "U"
      BtnType         =   3
      KeyCode         =   33
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   42
      Left            =   2520
      TabIndex        =   42
      Top             =   8820
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "2"
      Caption2nd      =   "TEST"
      CaptionAlpha    =   "V"
      BtnType         =   3
      KeyCode         =   25
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   43
      Left            =   3660
      TabIndex        =   43
      Top             =   8820
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "3"
      Caption2nd      =   "VARS"
      CaptionAlpha    =   "W"
      BtnType         =   3
      KeyCode         =   17
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   44
      Left            =   4800
      TabIndex        =   44
      Top             =   8820
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "+"
      Caption2nd      =   "MEM"
      CaptionAlpha    =   "X"
      BtnType         =   0
      KeyCode         =   9
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   45
      Left            =   240
      TabIndex        =   45
      Top             =   9480
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "ON"
      Caption2nd      =   "OFF"
      CaptionAlpha    =   ""
      BtnType         =   0
      KeyCode         =   255
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   46
      Left            =   1380
      TabIndex        =   46
      Top             =   9480
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "0"
      Caption2nd      =   "CHAR"
      CaptionAlpha    =   "Y"
      BtnType         =   3
      KeyCode         =   32
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   47
      Left            =   2520
      TabIndex        =   47
      Top             =   9480
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "."
      Caption2nd      =   ":"
      CaptionAlpha    =   "Z"
      BtnType         =   3
      KeyCode         =   24
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   48
      Left            =   3660
      TabIndex        =   48
      Top             =   9480
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "(-)"
      Caption2nd      =   "ANS"
      CaptionAlpha    =   "|_|"
      BtnType         =   3
      KeyCode         =   16
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   49
      Left            =   4800
      TabIndex        =   49
      Top             =   9480
      Width           =   1080
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "ENTER"
      Caption2nd      =   "ENTRY"
      CaptionAlpha    =   ""
      BtnType         =   0
      KeyCode         =   8
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   9
      Left            =   5160
      TabIndex        =   9
      Top             =   4440
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      Caption         =   "a>"
      Caption2nd      =   ""
      CaptionAlpha    =   ""
      BtnType         =   3
      KeyCode         =   2
   End
   Begin Ti85EmuVB6.UCTi85Btn cmdKey 
      Height          =   600
      Index           =   8
      Left            =   3660
      TabIndex        =   8
      Top             =   4440
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      Caption         =   "a<"
      Caption2nd      =   ""
      CaptionAlpha    =   ""
      BtnType         =   3
      KeyCode         =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Emulator"
      Begin VB.Menu mnuLoadRom 
         Caption         =   "Load ROM File..."
      End
      Begin VB.Menu mnuStateSaveAs 
         Caption         =   "Save State As..."
      End
      Begin VB.Menu mnuStateLoad 
         Caption         =   "Load State..."
      End
      Begin VB.Menu mnuReset 
         Caption         =   "Reset"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTheme 
      Caption         =   "Display"
      Begin VB.Menu mnuZoom 
         Caption         =   "Zoom"
         Begin VB.Menu mnuZooms 
            Caption         =   "100%"
            Index           =   0
         End
         Begin VB.Menu mnuZooms 
            Caption         =   "150%"
            Index           =   1
         End
         Begin VB.Menu mnuZooms 
            Caption         =   "200%"
            Index           =   2
         End
         Begin VB.Menu mnuZooms 
            Caption         =   "250%"
            Index           =   3
         End
         Begin VB.Menu mnuZooms 
            Caption         =   "300%"
            Index           =   4
         End
      End
      Begin VB.Menu mnuThemeBtn 
         Caption         =   "Buttons"
         Begin VB.Menu mnuThemes 
            Caption         =   "Dark"
            Index           =   0
         End
         Begin VB.Menu mnuThemes 
            Caption         =   "Light"
            Index           =   1
         End
      End
      Begin VB.Menu mnuLCD 
         Caption         =   "Screen"
         Begin VB.Menu mnuLCDs 
            Caption         =   "Blue / Green"
            Index           =   0
         End
         Begin VB.Menu mnuLCDs 
            Caption         =   "Brown / Orange"
            Index           =   1
         End
         Begin VB.Menu mnuLCDs 
            Caption         =   "Green / Black"
            Index           =   2
         End
         Begin VB.Menu mnuLCDs 
            Caption         =   "Black / White"
            Index           =   3
         End
         Begin VB.Menu mnuLCDs 
            Caption         =   "White / Black"
            Index           =   4
         End
      End
      Begin VB.Menu mnuScreenshot 
         Caption         =   "Capture Screen"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuLink 
      Caption         =   "Link"
      Begin VB.Menu mnuLinkSend 
         Caption         =   "Send File..."
      End
      Begin VB.Menu mnuLinkReceive 
         Caption         =   "Receive File..."
      End
      Begin VB.Menu mnuLinkCancel 
         Caption         =   "Cancel Transfer"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private mRunning            As Boolean
Private mOnDown             As Boolean
Private mKeyHeld(0 To 49)   As Boolean
Private mLastLinkStatus     As String
Private mSliceCount         As Long
Private mDownSlice(0 To 49) As Long
Private mReleaseAt(0 To 49) As Long

'============================================================
' Form_Load
'============================================================
Private Sub Form_Load()

    Ti85LoadRom MainConfig.RomFilePath
    Ti85Reset
    Ti85StateLoad MainConfig.StateFilePath
    
    BackColor = MainTheme.CalculatorBackColor
    
    If Not MainConfig.FirstTime Then
        Move MainConfig.WndLeft, MainConfig.WndTop, MainConfig.WndWidth, MainConfig.WndHeight
    Else
        Move (Screen.Width - Width) * 0.5, (Screen.Height - Height) * 0.5
    End If
    
End Sub

'============================================================
' Form_Unload
'============================================================
Private Sub Form_Unload(Cancel As Integer)
    If WindowState = vbNormal Then
        MainConfig.WndTop = Top
        MainConfig.WndLeft = Left
        MainConfig.WndWidth = Width
        MainConfig.WndHeight = Height
        MainConfig.FirstTime = False
    End If
    mRunning = False
    'ReleaseAllKeys
    'LinkReset
    Ti85StateSave MainConfig.StateFilePath
End Sub

'================================================
' Form_Resize
'================================================
Private Sub Form_Resize()
    If WindowState = vbMinimized Then Exit Sub
    MainLayout.Apply
    UCTi85LCD1.Invalidate
End Sub

'============================================================
' Main emulation loop
'============================================================
Public Sub RunEmulator()
    Dim sliceStart As Long
    
    Show
    DoEvents

    mRunning = True
    Do While mRunning
        sliceStart = GetTickCount()
        modCPU.Execute
        mSliceCount = mSliceCount + 1

        ' Keep the emulation close to the real 6 MHz speed.
        Do While mRunning And (GetTickCount() - sliceStart) < 30
            Sleep 1
        Loop

        UCTi85LCD1.Redraw
        LinkPoll
        UpdateLinkStatus
        DoEvents
    Loop
    
End Sub

Private Sub UpdateLinkStatus()
    Dim s As String

    s = LinkStatusText
    If s <> mLastLinkStatus Then
        mLastLinkStatus = s
        If Len(s) = 0 Then
            Me.Caption = "TI-85 Emulator"
        Else
            Me.Caption = "TI-85 Emulator - " & s
        End If
    End If

    If LinkNotifyPending Then
        LinkNotifyPending = False
        If LinkNotifyError Then
            MsgBox LinkNotifyText, vbExclamation, "TI-85 Link"
        Else
            MsgBox LinkNotifyText, vbInformation, "TI-85 Link"
        End If
    End If
End Sub

'================================================
' mnuAbout_Click
'================================================
Private Sub mnuAbout_Click()
    MsgBox _
    "=================" & vbCrLf & _
    "   Ti-85 Emulator v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
    "=================" & vbCrLf & _
    "by EBArtSoft@ - 2026" & vbCrLf & _
    "Created with Microsoft Visual Basic 6" & vbCrLf & _
    "AI assistance DeepSeek V4.1 Flash", vbInformation
End Sub

'================================================
' mnuScreenshot_Click
'================================================
Private Sub mnuScreenshot_Click()
    Dim vFileName As String
    vFileName = App.path & "\" & Format(Now, "yyyy-MM-dd-hh-mm-ss") & ".bmp"
    UCTi85LCD1.TakeScreenshot vFileName
End Sub

'================================================
' mnuThemes_Click
'================================================
Private Sub mnuThemes_Click(Index As Integer)
    Dim i As Long
    Select Case Index
    Case 0: MainConfig.ThemeName = "DARK"
    Case 1: MainConfig.ThemeName = "LIGHT"
    End Select
    MainTheme.SetThemeByName MainConfig.ThemeName
    For i = 0 To cmdKey.UBound
        cmdKey(i).Redraw
    Next
    BackColor = MainTheme.CalculatorBackColor
End Sub

'================================================
' mnuLCDs_Click
'================================================
Private Sub mnuLCDs_Click(Index As Integer)
    Dim i As Long
    Select Case Index
    Case 0: MainConfig.LCDName = "BLUEONGREEN"
    Case 1: MainConfig.LCDName = "BROWNONORANGE"
    Case 2: MainConfig.LCDName = "GREENONBLACK"
    Case 3: MainConfig.LCDName = "BLACKONWHITE"
    Case 4: MainConfig.LCDName = "WHITEONBLACK"
    End Select
    MainTheme.SetLCDByName MainConfig.LCDName
    UCTi85LCD1.Invalidate
End Sub

'================================================
' mnuZooms_Click
'================================================
Private Sub mnuZooms_Click(Index As Integer)
    Select Case Index
    Case 0: MainLayout.Resize 1
    Case 1: MainLayout.Resize 1.5
    Case 2: MainLayout.Resize 2
    Case 3: MainLayout.Resize 2.5
    Case 4: MainLayout.Resize 3
    End Select
End Sub

'================================================
' tmrKeys_Timer
'================================================
Private Sub tmrKeys_Timer()
    Dim i As Long
    Dim nowTick As Long

    nowTick = GetTickCount
    For i = 0 To 49
        If mKeyHeld(i) Then
            If (nowTick >= mReleaseAt(i)) And (mSliceCount > mDownSlice(i)) Then
                mKeyHeld(i) = False
                ActivateKeyCode cmdKey(i).KeyCode, False
            End If
        End If
    Next
End Sub

'================================================
' cmdKey_MouseDown
'================================================
Private Sub cmdKey_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index < 0 Or Index > 49 Then Exit Sub
    If mKeyHeld(Index) Then Exit Sub
    mKeyHeld(Index) = True
    mDownSlice(Index) = mSliceCount
    mReleaseAt(Index) = GetTickCount + 200
    ActivateKeyCode cmdKey(Index).KeyCode, True
End Sub

'================================================
' cmdKey_MouseUp
'================================================
Private Sub cmdKey_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index < 0 Or Index > 49 Then Exit Sub
    If mKeyHeld(Index) Then mReleaseAt(Index) = GetTickCount + 200
    Me.SetFocus
End Sub

'================================================
' ActivateKeyCode
'================================================
Private Sub ActivateKeyCode(ByVal kc As Long, ByVal Down As Boolean)
    Dim r As Long
    Dim c As Long
    If kc = 255 Then
        If Down Then OnPress Else OnRelease
    Else
        r = kc \ 8
        c = kc Mod 8
        If Down Then KeyDown r, c Else KeyUp r, c
    End If
End Sub

'================================================
' OnPress
'================================================
Private Sub OnPress()
    If mOnDown Then Exit Sub
    mOnDown = True
    PressOn
End Sub

'================================================
' OnRelease
'================================================
Private Sub OnRelease()
    If Not mOnDown Then Exit Sub
    mOnDown = False
    ReleaseOn
End Sub

'============================================================
' Menu
'============================================================
Private Sub mnuReset_Click()
    If MsgBox("This action will clear the memory;" & vbCrLf & "do you want to continue?", vbQuestion Or vbYesNo) = vbYes Then
        Ti85Reset
        mOnDown = False
        UCTi85LCD1.Invalidate
        UCTi85LCD1.Redraw
        'Ti85StateLoad MainConfig.RamFilePath
        Me.SetFocus
    End If
End Sub

'================================================
' mnuExit_Click
'================================================
Private Sub mnuExit_Click()
    Unload Me
End Sub

'================================================
' mnuLoadRom_Click
'================================================
Private Sub mnuLoadRom_Click()
    Dim vCdg As CommonDialog
    Set vCdg = New CommonDialog
    vCdg.Parent = hWnd
    vCdg.InitDir = App.path
    vCdg.DialogTitle = "Send file to TI-85"
    vCdg.Filter = "ROM File (.ROM)|*.ROM|All files (*.*)|*.*"
    vCdg.Flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST 'Or OFN_EXPLORER
    If vCdg.ShowOpen() Then
        MainConfig.RomFilePath = vCdg.FileName
        Ti85LoadRom MainConfig.RomFilePath
        mnuReset_Click
    End If
End Sub

'================================================
' mnuStateSaveAs_Click
'================================================
Private Sub mnuStateSaveAs_Click()
    Dim vCdg As CommonDialog
    Set vCdg = New CommonDialog
    vCdg.Parent = hWnd
    vCdg.InitDir = App.path
    vCdg.DefaultExt = "STATE"
    vCdg.DialogTitle = "Save TI-85 State"
    vCdg.Filter = "State File (.STATE)|*.STATE|All files (*.*)|*.*"
    vCdg.Flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST 'Or OFN_EXPLORER
    If vCdg.ShowSave() Then
        Ti85StateSave vCdg.FileName
    End If
End Sub

'================================================
' mnuStateLoad_Click
'================================================
Private Sub mnuStateLoad_Click()
    Dim vCdg As CommonDialog
    Set vCdg = New CommonDialog
    vCdg.Parent = hWnd
    vCdg.InitDir = App.path
    vCdg.DefaultExt = "STATE"
    vCdg.DialogTitle = "Load TI-85 State"
    vCdg.Filter = "State File (.STATE)|*.STATE|All files (*.*)|*.*"
    vCdg.Flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST 'Or OFN_EXPLORER
    If vCdg.ShowOpen() Then
        Ti85LoadRom MainConfig.RomFilePath
        Ti85Reset
        If Not Ti85StateLoad(vCdg.FileName) Then
            MsgBox "State load failed", vbCritical
        End If
    End If
End Sub

'================================================
' LinkBrowseFile
'================================================
Private Function LinkBrowseFile() As String
    Dim vDlg As CommonDialog
    Set vDlg = New CommonDialog
    vDlg.InitDir = App.path
    vDlg.Parent = frmMain.hWnd
    vDlg.DialogTitle = "Send file to TI-85"
    vDlg.Filter = "All TI-85 Files (*.85?)|*.85?|All Files (*.*)|*.*"
    vDlg.Flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST 'Or OFN_EXPLORER
    If vDlg.ShowOpen() Then
        LinkBrowseFile = vDlg.FileName
    End If
End Function

'================================================
' mnuLinkSend_Click
'================================================
Private Sub mnuLinkSend_Click()
    Dim FilePath As String
    Dim Prompt As String

    If LinkIsBusy() Then
        MsgBox "A link transfer is already in progress.", vbExclamation, "TI-85 Link"
        Exit Sub
    End If

    FilePath = LinkBrowseFile()
    If Len(FilePath) = 0 Then Exit Sub

    Prompt = "Put the calculator in receive mode (2nd -> LINK -> RECV)," & vbCrLf & "then click OK."
    If LCase$(Right$(FilePath, 4)) = ".85b" Then
        Prompt = Prompt & vbCrLf & vbCrLf & _
                 "The calculator will show WARNING Memory Backup." & vbCrLf & _
                 "Select CONT to start the transfer. This replaces its memory."
    End If

    If MsgBox(Prompt, vbOKCancel + vbInformation, "TI-85 Link") <> vbOK Then
        Exit Sub
    End If

    If Not LinkStartSendFile(FilePath) Then
        MsgBox LinkLastError, vbExclamation, "TI-85 Link"
    End If
End Sub

'================================================
' mnuLinkReceive_Click
'================================================
Private Sub mnuLinkReceive_Click()
    If LinkIsBusy() Then
        MsgBox "A link transfer is already in progress.", vbExclamation, "TI-85 Link"
        Exit Sub
    End If

    LinkArmReceive
    MsgBox "Receiver armed. On the calculator, press 2nd -> LINK," & vbCrLf & _
           "choose SEND and select the data to transfer.", vbInformation, "TI-85 Link"
End Sub

'================================================
' mnuLinkCancel_Click
'================================================
Private Sub mnuLinkCancel_Click()
    LinkCancel
End Sub

''============================================================
'' Keyboard
''============================================================
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    HandleKey KeyCode, True
'End Sub
'
'Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'    HandleKey KeyCode, False
'End Sub

'Private Sub HandleKey(ByVal KeyCode As Integer, ByVal Down As Boolean)
'    Select Case KeyCode
'    Case vbKey0: SetMatrix 4, 7, Down
'    Case vbKey1: SetMatrix 4, 6, Down
'    Case vbKey2: SetMatrix 3, 6, Down
'    Case vbKey3: SetMatrix 2, 6, Down
'    Case vbKey4: SetMatrix 4, 5, Down
'    Case vbKey5: SetMatrix 3, 5, Down
'    Case vbKey6: SetMatrix 2, 5, Down
'    Case vbKey7: SetMatrix 4, 4, Down
'    Case vbKey8: SetMatrix 3, 4, Down
'    Case vbKey9: SetMatrix 2, 4, Down
'
'    Case 190, vbKeyDecimal: SetMatrix 3, 7, Down
'    Case vbKeyAdd, 187: SetMatrix 1, 6, Down
'    Case vbKeySubtract, 189: SetMatrix 1, 5, Down
'    Case vbKeyMultiply: SetMatrix 1, 4, Down
'    Case vbKeyDivide, 191: SetMatrix 1, 3, Down
'    Case vbKeyReturn: SetMatrix 1, 7, Down
'
'    Case vbKeyBack, vbKeyDelete: SetMatrix 3, 0, Down
'    Case vbKeyEscape: SetMatrix 1, 1, Down
'    Case vbKeyTab: SetMatrix 5, 6, Down
'
'    Case vbKeyUp: SetMatrix 0, 4, Down
'    Case vbKeyDown: SetMatrix 0, 7, Down
'    Case vbKeyLeft: SetMatrix 0, 6, Down
'    Case vbKeyRight: SetMatrix 0, 5, Down
'
'    Case vbKeyF1: SetMatrix 6, 3, Down
'    Case vbKeyF2: SetMatrix 6, 4, Down
'    Case vbKeyF3: SetMatrix 6, 5, Down
'    Case vbKeyF4: SetMatrix 6, 6, Down
'    Case vbKeyF5: SetMatrix 6, 7, Down
'
'    Case vbKeyCapital: SetMatrix 4, 0, Down
'    Case vbKeyShift, vbKeyControl: SetMatrix 6, 2, Down
'
'    Case vbKeyP: SetMatrix 3, 1, Down
'    Case vbKeyG: SetMatrix 5, 1, Down
'    Case vbKeyS: SetMatrix 4, 2, Down
'    Case vbKeyC: SetMatrix 3, 2, Down
'    Case vbKeyT: SetMatrix 2, 2, Down
'    Case vbKeyL: SetMatrix 5, 2, Down
'    Case vbKeyN: SetMatrix 5, 3, Down
'    Case vbKeyR: SetMatrix 5, 4, Down
'    Case 219: SetMatrix 3, 3, Down
'    Case 221: SetMatrix 2, 3, Down
'    Case 188: SetMatrix 5, 5, Down
'    Case vbKeyM: SetMatrix 6, 0, Down
'    Case vbKeyX: SetMatrix 6, 1, Down
'
'    Case vbKeyQ, vbKeyF12
'        If Down Then OnPress Else OnRelease
'    End Select
'End Sub

'Private Sub SetMatrix(ByVal Row As Long, ByVal Col As Long, ByVal Down As Boolean)
'    ' The TI-85 keypad matrix columns are wired in reverse bit order.
'    If Down Then
'        KeyDown Row, 7 - Col
'    Else
'        KeyUp Row, 7 - Col
'    End If
'End Sub

