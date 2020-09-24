VERSION 5.00
Begin VB.Form frmCPUMeter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'Kein
   Caption         =   "CPU Meter"
   ClientHeight    =   315
   ClientLeft      =   1335
   ClientTop       =   1185
   ClientWidth     =   2145
   Icon            =   "CpuMeter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   21
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   143
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picStatus 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      FillStyle       =   0  'Ausgefüllt
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   75
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   60
      Width           =   1530
   End
   Begin VB.Timer tmrCpuStatus 
      Left            =   -60
      Top             =   -60
   End
   Begin VB.Timer tmrFormMove 
      Left            =   330
      Top             =   -60
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1620
      TabIndex        =   1
      Top             =   45
      Width           =   465
   End
   Begin VB.Shape shpBack 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Undurchsichtig
      Height          =   240
      Left            =   30
      Top             =   30
      Width           =   2070
   End
End
Attribute VB_Name = "frmCPUMeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' -----------------------------------------------------------------------------
' CPU Meter Version 1.0 by X-LEAD
' -----------------------------------------------------------------------------
' mail me any comments, updates or errors to: xlead@drakemedia.de
' -----------------------------------------------------------------------------
' most api declarations are picked from the API Guide ... get it for free at:
' http://kpdweb.cjb.net/apiguide/index.htm ... this program is a must have for
' vb-programmers ...
' -----------------------------------------------------------------------------
' if anyone knows how to draw the statusbar with GDI or how to place the
' complete CPU Meter in the systemtray next to the clock ... please let me know
' -----------------------------------------------------------------------------
'


Option Explicit

' api declaration to get the cursors position
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

' declare type to store the coordinates
Private Type POINTAPI
    X As Long
    Y As Long
End Type

' api declarations for our CPU meter
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long

Const REG_DWORD = 4
Const HKEY_DYN_DATA = &H80000006

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

' api declarations to raise our form
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

Const DFC_BUTTON = 4
Const DFCS_BUTTON3STATE = &H10

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' api declarations to make form stay on top
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_SHOWWINDOW = &H40
Const HWND_TOPMOST = -1
Const HWND_NOTTOPMOST = -2


Private Sub Form_Load()

    ' set the two timer intervals
    tmrFormMove.Interval = 1
    tmrCpuStatus.Interval = 500 'used 500 cause our program needs resources too
    
    ' color the background shape and picturebox
    shpBack.BackColor = RGB(0, 10, 90)
    shpBack.BorderColor = RGB(0, 10, 90)
    picStatus.BackColor = RGB(130, 130, 170)
    
    ' raise our form
    RaiseForm
    
    ' initialize our CPU meter
    InitCPU
    
End Sub


Private Sub RaiseForm()
    
    Dim R As RECT
    
    Me.ScaleMode = vbPixels
    SetRect R, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    DrawFrameControl Me.hdc, R, DFC_BUTTON, DFCS_BUTTON3STATE
    OffsetRect R, 0, 22

End Sub


Private Sub InitCPU()

    Dim lData As Long
    Dim lType As Long
    Dim lSize As Long
    Dim hKey As Long
    Dim Qry As String
    
    Qry = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StartStat", hKey)
    
    If Qry <> 0 Then
            MsgBox "Could not open registery!"
        End
    End If
    
    lType = REG_DWORD
    lSize = 4
    
    Qry = RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, lType, lData, lSize)
    Qry = RegCloseKey(hKey)

End Sub


Private Sub tmrFormMove_Timer()
    
    Dim Point As POINTAPI
    
    ' get the cursorposition
    GetCursorPos Point

    ' multiply the coordinates to convert twips to pixel and place the form
    Me.Left = Point.X * 15 + 200
    Me.Top = Point.Y * 15 + 150
    
     ' make our form stay on top
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE
    
End Sub


Private Sub tmrCpuStatus_Timer()

    Dim lData As Long
    Dim lType As Long
    Dim lSize As Long
    Dim hKey As Long
    Dim Qry As String
    Dim Status As Long
                  
    Qry = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StatData", hKey)
                
    If Qry <> 0 Then
            MsgBox "Could not open registery!"
        End
    End If
                
    lType = REG_DWORD
    lSize = 4
                
    Qry = RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, lType, lData, lSize)
    
    Status = Int(lData)

    ' show CPU usage in Label
    lblStatus.Caption = Status & "%"
    
    ' show CPU usage in our selfmade progressbar
    ' when CPU usage is over 80% then color the status red
    If Status < 80 Then
    picStatus.Line (Status, 0)-(0, 10), RGB(255, 245, 85), BF
    Else
    picStatus.Line (Status, 0)-(0, 10), RGB(245, 10, 0), BF
    End If
    picStatus.Line (Status, 0)-(100, 10), RGB(130, 130, 170), BF
    
    Qry = RegCloseKey(hKey)

End Sub
