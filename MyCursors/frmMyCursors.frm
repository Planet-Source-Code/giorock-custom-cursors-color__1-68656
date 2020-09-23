VERSION 5.00
Begin VB.Form Form1
   Caption         =   "My Cursors"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   Icon            =   "frmMyCursors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   153
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   2  'CenterScreen
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

' My Custom Cursors
' P.S. Resource is compiled with type CUSTOM
'      not CURSOR!!!
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

' The LoadCursorFromFile function creates a cursor based on data contained in a file.
' The file is specified by its name or by a system cursor identifier.
' The function returns a handle to the newly created cursor.
' Files containing cursor data may be in either cursor (.CUR)
' or animated cursor (.ANI) format.
' All cursor formats are supported: 2Color, 16Color, 256Color, etc.
Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long

' The SetClassLong function replaces the specified 32-bit (long) value
' at the specified offset into the extra class memory or the WNDCLASS structure
' for the class to which the specified window belongs.
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GCL_HCURSOR = (-12)

' Old Window Cursor Handle to restore
Private hOldCursor As Long
' Old Window Index to restore
Private lOldIndex As Long

Private bLoad As Boolean
Private Sub ColorizeCursor(ByVal IDCursor As eMyCursors, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte)
Dim hff As Integer
Dim bArr() As Byte

    ' LoadResData returns a Byte Array
    ' so is easy to manipulate
    bArr = LoadResData(IDCursor, "CUSTOM")

    ' I replace White Palette only
    ' with custom Color
    If RGB(R, G, B) <> vbWhite Then
        ' Blue
        bArr(122) = B
        ' Green
        bArr(123) = G
        ' Red
        bArr(124) = R
    End If

    ' Create a temporary cursor file by resource data
    hff = FreeFile
    Open App.Path + "\TempCur.cur" For Binary Access Write As #hff
        Put #hff, , bArr()
    Close #hff

    Erase bArr

End Sub

Private Sub Combo1_Click(Index As Integer)
    If Not bLoad Then
        Label1.BackColor = RGB(Combo1(0).ListIndex, Combo1(1).ListIndex, Combo1(2).ListIndex)
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim hCur As Long

    ColorizeCursor Index + 1001, Combo1(0).ListIndex, Combo1(1).ListIndex, Combo1(2).ListIndex

    hCur = LoadCursorFromFile(App.Path + "\TempCur.cur")

    RestoreOldCursor

    hOldCursor = SetClassLong(Command1(Index).hWnd, GCL_HCURSOR, hCur)

    lOldIndex = Index

End Sub

Private Sub Form_Load()
    bLoad = True
    FillComboBox
    bLoad = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RestoreOldCursor
    End
    Set Form1 = Nothing
End Sub

Private Sub RestoreOldCursor()
    ' Delete temporary cursor file if exist
    If Dir$(App.Path + "\TempCur.cur") <> "" Then: Kill App.Path + "\TempCur.cur"
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
