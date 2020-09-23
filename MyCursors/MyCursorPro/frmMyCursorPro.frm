VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "My Cursors Pro"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   Icon            =   "frmMyCursorPro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   153
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3510
      Top             =   1845
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      Left            =   2055
      Style           =   2  'Dropdown List
      TabIndex        =   17
      ToolTipText     =   "BLUE Color"
      Top             =   1950
      Width           =   960
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   1095
      Style           =   2  'Dropdown List
      TabIndex        =   16
      ToolTipText     =   "GREEN Color"
      Top             =   1950
      Width           =   960
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   135
      Style           =   2  'Dropdown List
      TabIndex        =   15
      ToolTipText     =   "RED Color"
      Top             =   1950
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Wait"
      Height          =   495
      Index           =   14
      Left            =   3555
      TabIndex        =   14
      Top             =   1350
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Up"
      Height          =   495
      Index           =   13
      Left            =   2700
      TabIndex        =   13
      Top             =   1350
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Size WE"
      Height          =   495
      Index           =   12
      Left            =   1845
      TabIndex        =   12
      Top             =   1350
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Size NWSE"
      Height          =   495
      Index           =   11
      Left            =   990
      TabIndex        =   11
      Top             =   1350
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Size NS"
      Height          =   495
      Index           =   10
      Left            =   135
      TabIndex        =   10
      Top             =   1350
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Size NESW"
      Height          =   495
      Index           =   9
      Left            =   3555
      TabIndex        =   9
      Top             =   735
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pen"
      Height          =   495
      Index           =   8
      Left            =   2700
      TabIndex        =   8
      Top             =   735
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "No"
      Height          =   495
      Index           =   7
      Left            =   1845
      TabIndex        =   7
      Top             =   735
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Move"
      Height          =   495
      Index           =   6
      Left            =   990
      TabIndex        =   6
      Top             =   735
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "IBeam"
      Height          =   495
      Index           =   5
      Left            =   135
      TabIndex        =   5
      Top             =   735
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hourglass"
      Height          =   495
      Index           =   4
      Left            =   3555
      TabIndex        =   4
      Top             =   135
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Help"
      Height          =   495
      Index           =   3
      Left            =   2700
      TabIndex        =   3
      Top             =   135
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hand"
      Height          =   495
      Index           =   2
      Left            =   1845
      TabIndex        =   2
      Top             =   135
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cross"
      Height          =   495
      Index           =   1
      Left            =   990
      TabIndex        =   1
      Top             =   135
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Arrow"
      Height          =   495
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   870
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3060
      TabIndex        =   18
      ToolTipText     =   "Click me!!!"
      Top             =   1950
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***********************
'  Written by GioRock  *
'***********************
'***********************
'      Completely      *
'  Created by GioRock  *
'***********************

' WELL!!!: How to use custom cursor with 16777215 colors
'          in a simple way

' The OSVERSIONINFO data structure contains operating system version information.
' The information includes major and minor version numbers, a build number,
' a platform identifier, and descriptive text about the operating system.
' This structure is used with the GetVersionEx function.
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 ' Maintenance string for PSS usage
End Type
' The GetVersionEx function obtains extended information about the version of the
' operating system that is currently running.
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

' My Custom Cursors
' P.S. Resource is compiled with type CURSOR!!!
' All cursors are created 32x32 with 16 colors
Private Enum eMyCursors
    ARROW_CUR = 1001
    CROSS_CUR = 1002
    HAND_CUR = 1003
    HELP_CUR = 1004
    HOURGLASS_CUR = 1005
    I_BEAM_CUR = 1006
    MOVE_CUR = 1007
    NO_CUR = 1008
    PEN_CUR = 1009
    SIZE_NESW_CUR = 1010
    SIZE_NS_CUR = 1011
    SIZE_NWSE_CUR = 1012
    SIZE_WE_CUR = 1013
    UP_CUR = 1014
    WAIT_CUR = 1015
End Enum

' The ICONINFO structure contains information about an icon or a cursor.
Private Type ICONINFO
    fIcon As Long       ' Specifies whether this structure defines an icon or a cursor.
                        ' A value of TRUE specifies an icon; FALSE specifies a cursor.
    xHotspot As Long    ' Specifies the x-coordinate of a cursor's hot spot.
    yHotspot As Long    ' Specifies the y-coordinate of a cursor's hot spot.
                        ' If these structures defines an icon, the hot spot is always
                        ' in the center of the icon, and this member is ignored.
    hbmMask As Long     ' Specifies the icon bitmask bitmap.
    hbmColor As Long    ' Identifies the icon color bitmap.
End Type
' The GetIconInfo function retrieves information about the specified icon or cursor.
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
' The CreateIconIndirect function creates an icon or cursor from an ICONINFO structure.
Private Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long
' The DestroyIcon function destroys an icon and frees any memory the icon occupied
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

' The SelectObject function selects an object into the specified device context.
' The new object replaces the previous object of the same type.
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
' The CreateCompatibleDC function creates a memory device context (DC)
' compatible with the specified device.
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
' The DeleteDC function deletes the specified device context (DC).
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
' The SetPixel function sets the pixel at the specified coordinates to the specified color.
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
' The GetPixel function retrieves the red, green, blue (RGB) color value of the pixel
' at the specified coordinates.
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
' The CreateCompatibleBitmap function creates a bitmap compatible with the device
' that is associated with the specified device context.
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
' The DeleteObject function deletes a logical pen, brush, font, bitmap, region, or palette,
' freeing all system resources associated with the object.
' After the object is deleted, the specified handle is no longer valid.
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
' The DrawIcon function draws an icon in the client area of the window of the specified
' device context.
Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long

' The SetClassLong function replaces the specified 32-bit (long) value
' at the specified offset into the extra class memory or the WNDCLASS structure
' for the class to which the specified window belongs.
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GCL_HCURSOR = (-12)

' Cursor Handle created
Private hCur As Long
' Old Window Cursor Handle to restore
Private hOldCursor As Long
' Old Window Index to restore
Private lOldIndex As Long

Private bLoad As Boolean
Private bSystem9X As Boolean

Private Function GetColorizedCursor(ByVal IDCursor As eMyCursors, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte) As Long
Dim pIF As ICONINFO
Dim p As StdPicture
Dim hDCIcon As Long, hDCMask As Long
Dim X As Long, Y As Long
    
    ' Load Cursor Resource inside picture object
    Set p = LoadResPicture(IDCursor, vbResCursor)

    ' Get Icon Informations by picture handle
    GetIconInfo p.Handle, pIF

    ' Create compatible device context
    hDCIcon = CreateCompatibleDC(0)
    If bSystem9X Then: hDCMask = CreateCompatibleDC(0)
    
    ' Since pIF.hbmColor could be returns always zero
    ' you must create a valid color bitmap handle
    ' NOTE: On WinNT or higher this project must be
    ' compiled before use it!!!
    ' To see some result at design time, set hDC below to zero
    ' No cursor color will be displayed...default only.
    pIF.hbmColor = CreateCompatibleBitmap(hDC, 32, 32)
    
    ' Select Object in device context
    SelectObject hDCIcon, pIF.hbmColor
    If bSystem9X Then: SelectObject hDCMask, pIF.hbmMask
           
    ' Draw color bitmap
    DrawIcon hDCIcon, 0, 0, p.Handle
    
    ' Replace default white color by one specified through RGB values
    ' P.S. You can treat and manipulate cursor as bitmap before creation here!!!
    For X = 0 To 31
        For Y = 0 To 31
            ' Ensure to paint white color only and not black borders
            If IIf(bSystem9X, GetPixel(hDCMask, X, Y) = vbBlack, 1) And GetPixel(hDCIcon, X, Y) <> vbBlack Then
                SetPixel hDCIcon, X, Y, RGB(R, G, B)
                ' Otherwise set all black
            Else
                SetPixel hDCIcon, X, Y, vbBlack
            End If
        Next Y
    Next X

    ' Create new color cursor handle by ICONINFO structure
    GetColorizedCursor = CreateIconIndirect(pIF)
    
    ' Clean up unusued memory
    DeleteObject pIF.hbmColor
    DeleteObject pIF.hbmMask
    DeleteDC hDCIcon
    If bSystem9X Then: DeleteDC hDCMask
    
    Set p = Nothing
     
End Function

Private Sub Combo1_Click(Index As Integer)
    If Not bLoad Then
        Label1.BackColor = RGB(Combo1(0).ListIndex, Combo1(1).ListIndex, Combo1(2).ListIndex)
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    
    ' Icon or cursor creation must be always destroyed
    DestroyOldCursor

    ' Get new colorized cursor handle
    hCur = GetColorizedCursor(Index + 1001, Combo1(0).ListIndex, Combo1(1).ListIndex, Combo1(2).ListIndex)

    ' Undo last cursor change if exist
    RestoreOldCursor

    ' Sets new custom cursor handle and returns the previous
    hOldCursor = SetClassLong(Command1(Index).hWnd, GCL_HCURSOR, hCur)

    ' Remember last window index to restore
    lOldIndex = Index

End Sub

Private Sub Form_Load()
    bLoad = True
    FillComboBox
    bLoad = False
    bSystem9X = GetSystem9X()
    If Not bSystem9X Then
        Debug.Print "Warning - This project must be compiled!!!"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Undo last cursor change if exist
    RestoreOldCursor
    ' Icon or cursor creation must be always destroyed
    DestroyOldCursor
    ' Terminate and free program memory
    End
    Set Form1 = Nothing
End Sub

Private Sub RestoreOldCursor()
    If hOldCursor Then
        ' You must restore original window cursor before apply changes
        Call SetClassLong(Command1(lOldIndex).hWnd, GCL_HCURSOR, hOldCursor)
    End If
End Sub

Private Sub FillComboBox()
Dim i As Integer, j As Integer

    ' Fill ComboBox with RGB values
    For i = 0 To 255
        For j = Combo1.LBound To Combo1.UBound
            Combo1(j).AddItem CStr(i)
        Next j
    Next i

    ' Set Default Color to vbWhite
    For j = Combo1.LBound To Combo1.UBound
        Combo1(j).ListIndex = 255
    Next j

End Sub

Private Sub DestroyOldCursor()
    If hCur Then
        ' Destroy unusued Icon or Cursor
        ' if no longer needed
        Call DestroyIcon(hCur)
        hCur = 0
    End If
End Sub

Private Sub Label1_Click()
Dim lOldColor As Long

    On Error GoTo Out
    
    With CD1
        lOldColor = Label1.BackColor
        .Color = lOldColor
        .Flags = cdlCCFullOpen Or cdlCCRGBInit
        .ShowColor
        Combo1(0).ListIndex = (.Color And &HFF)
        Combo1(1).ListIndex = (.Color \ &H100) And &HFF
        Combo1(2).ListIndex = (.Color \ &H10000) And &HFF
'        Command1_Click CInt(lOldIndex)
    End With
    
    Exit Sub

Out:

    Label1.BackColor = lOldColor
    
End Sub



Private Function GetSystem9X() As Boolean
Dim OSI As OSVERSIONINFO
Dim sPlatform As String

    OSI.dwOSVersionInfoSize = Len(OSI)
    
    GetVersionEx OSI
    
    GoSub GetPlatForm
    
    Debug.Print "Platform:", sPlatform
    Debug.Print "Version :", CStr(OSI.dwMajorVersion) + "." + CStr(OSI.dwMinorVersion) + "." + _
        CStr(IIf((OSI.dwPlatformId = VER_PLATFORM_WIN32_NT), OSI.dwBuildNumber, (OSI.dwBuildNumber Mod &H1000) And &HFFF))

    GetSystem9X = Not CBool(OSI.dwPlatformId = VER_PLATFORM_WIN32_NT)
      
    Exit Function
    
GetPlatForm:
    
    Select Case OSI.dwPlatformId
        Case VER_PLATFORM_WIN32s
            sPlatform = "Win32s on Windows 3.1"
        Case VER_PLATFORM_WIN32_WINDOWS
            sPlatform = "Win32 on Windows 9X"
        Case VER_PLATFORM_WIN32_NT
            sPlatform = "Win32 on Windows NT"
    End Select
    
    Return
    
End Function
