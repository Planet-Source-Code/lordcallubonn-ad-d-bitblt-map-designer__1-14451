VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Floor Tool Bar"
   ClientHeight    =   5445
   ClientLeft      =   150
   ClientTop       =   315
   ClientWidth     =   2100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   48
      Left            =   1560
      Picture         =   "tollbar2.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   40
      ToolTipText     =   "Archway East/West"
      Top             =   3000
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   47
      Left            =   1560
      Picture         =   "tollbar2.frx":033A
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   39
      ToolTipText     =   "Archway North/South"
      Top             =   2520
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   46
      Left            =   1560
      Picture         =   "tollbar2.frx":0677
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   38
      ToolTipText     =   "Archway 4ways"
      Top             =   2040
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   45
      Left            =   1560
      Picture         =   "tollbar2.frx":09C1
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   37
      ToolTipText     =   "Archway South & West"
      Top             =   1560
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   44
      Left            =   1560
      Picture         =   "tollbar2.frx":0CFD
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   36
      ToolTipText     =   "Archway South & East"
      Top             =   1080
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   43
      Left            =   1560
      Picture         =   "tollbar2.frx":103E
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   35
      ToolTipText     =   "Archway North & West"
      Top             =   600
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   42
      Left            =   1560
      Picture         =   "tollbar2.frx":1357
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   34
      ToolTipText     =   "Archway North & East"
      Top             =   120
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   37
      Left            =   1080
      Picture         =   "tollbar2.frx":1553
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   33
      ToolTipText     =   "Corner Upper Left"
      Top             =   3000
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   38
      Left            =   1080
      Picture         =   "tollbar2.frx":1750
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   32
      ToolTipText     =   "Corner Lower Left"
      Top             =   3480
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   36
      Left            =   1080
      Picture         =   "tollbar2.frx":1971
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   31
      ToolTipText     =   "Corner Lower Right"
      Top             =   2520
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   35
      Left            =   1080
      Picture         =   "tollbar2.frx":1BCD
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   30
      ToolTipText     =   "Corner Upper Right"
      Top             =   2040
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   19
      Left            =   120
      Picture         =   "tollbar2.frx":1DE4
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   29
      ToolTipText     =   "Noth/South Stairs"
      Top             =   3000
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   18
      Left            =   120
      Picture         =   "tollbar2.frx":201C
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   28
      ToolTipText     =   "East/West Stairs"
      Top             =   3480
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   17
      Left            =   1080
      Picture         =   "tollbar2.frx":2251
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   27
      ToolTipText     =   "Secret Passage West"
      Top             =   600
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   16
      Left            =   1080
      Picture         =   "tollbar2.frx":245F
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   26
      ToolTipText     =   "Secret Passage South"
      Top             =   1560
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   15
      Left            =   1080
      Picture         =   "tollbar2.frx":27CA
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   25
      ToolTipText     =   "Secret Passage North"
      Top             =   1080
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   14
      Left            =   1080
      Picture         =   "tollbar2.frx":2B1E
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   24
      ToolTipText     =   "Secret Passage East"
      Top             =   120
      Width           =   435
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Tile"
      ClipControls    =   0   'False
      Height          =   975
      Left            =   600
      TabIndex        =   14
      Top             =   4320
      Width           =   975
      Begin VB.PictureBox picCurrent 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   420
         Left            =   240
         ScaleHeight     =   360
         ScaleWidth      =   405
         TabIndex        =   15
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "X, Y:"
         Height          =   195
         Left            =   1200
         TabIndex        =   23
         Top             =   360
         Width           =   345
      End
      Begin VB.Label lblXY 
         AutoSize        =   -1  'True
         Caption         =   "1,2"
         Height          =   195
         Left            =   2400
         TabIndex        =   22
         Top             =   360
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   195
         Left            =   1200
         TabIndex        =   21
         Top             =   600
         Width           =   405
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Left            =   2400
         TabIndex        =   20
         Top             =   600
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ImgNum:"
         Height          =   195
         Left            =   1200
         TabIndex        =   19
         Top             =   840
         Width           =   630
      End
      Begin VB.Label lblImgNum 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   195
         Left            =   2400
         TabIndex        =   18
         Top             =   840
         Width           =   90
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   1200
         TabIndex        =   17
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   195
         Left            =   2400
         TabIndex        =   16
         Top             =   1080
         Width           =   90
      End
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   120
      Picture         =   "tollbar2.frx":2E83
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   13
      ToolTipText     =   "Main Floor"
      Top             =   120
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   1
      Left            =   120
      Picture         =   "tollbar2.frx":305A
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   12
      ToolTipText     =   "Bottom Door"
      Top             =   1080
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   2
      Left            =   120
      Picture         =   "tollbar2.frx":328C
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   11
      ToolTipText     =   "Upper Door"
      Top             =   1560
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   3
      Left            =   120
      Picture         =   "tollbar2.frx":3476
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   10
      ToolTipText     =   "Right Door"
      Top             =   2520
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   4
      Left            =   120
      Picture         =   "tollbar2.frx":36A0
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   9
      ToolTipText     =   "Empty Floor"
      Top             =   600
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   5
      Left            =   120
      Picture         =   "tollbar2.frx":3804
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   8
      ToolTipText     =   "Left Door"
      Top             =   2040
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   6
      Left            =   600
      Picture         =   "tollbar2.frx":3A06
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   7
      ToolTipText     =   "Archway North"
      Top             =   120
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   7
      Left            =   600
      Picture         =   "tollbar2.frx":3D15
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   6
      ToolTipText     =   "Archway West"
      Top             =   1080
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   8
      Left            =   600
      Picture         =   "tollbar2.frx":3EF3
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   5
      ToolTipText     =   "Archway East"
      Top             =   1560
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   9
      Left            =   600
      Picture         =   "tollbar2.frx":40E7
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   4
      ToolTipText     =   "Archway South"
      Top             =   600
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   10
      Left            =   600
      Picture         =   "tollbar2.frx":42E3
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   3
      ToolTipText     =   "Thick Wall North"
      Top             =   2040
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   11
      Left            =   600
      Picture         =   "tollbar2.frx":44C1
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   2
      ToolTipText     =   "Thick Wall South"
      Top             =   2520
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   12
      Left            =   600
      Picture         =   "tollbar2.frx":46E3
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   1
      ToolTipText     =   "Thick Wall East"
      Top             =   3480
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   13
      Left            =   600
      Picture         =   "tollbar2.frx":4903
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   0
      ToolTipText     =   "Thick Wall West"
      Top             =   3000
      Width           =   435
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

' OnTop Sample by Matt Hart - vbhelp@matthart.com
' http://matthart.com
'
' This app shows you how to make a ""floating"" window which will stay on top
' even when not activated.  Note that you MUST compile this and run as an EXE
' file before it will stay on top of all other running programs.  That is because
' when it is run in the VB environment, it is considered a ""child"" window of VB,
' and will thus only float on top of other VB windows.

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
' SetWindowPos Flags
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOZORDER = &H4
Const SWP_NOREDRAW = &H8
Const SWP_NOACTIVATE = &H10
Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Const SWP_SHOWWINDOW = &H40
Const SWP_HIDEWINDOW = &H80
Const SWP_NOCOPYBITS = &H100
Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering

Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
' SetWindowPos() hwndInsertAfter values
Const HWND_TOP = 0
Const HWND_BOTTOM = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Private Sub Form_Load()
    Dim i
    i = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

End Sub

Private Sub picTile_Click(Index As Integer)
     frmMap.lblName = Tiles(Index).Name
        'view the name of the current tile
    frmMap.lblType = Tiles(Index).Type
        'view the type of the current tile
    frmMap.lblImgNum = Tiles(Index).ImageNum
        'view the ImageNum of the current tile
   ret = BitBlt(picCurrent.hDC, 0, 0, TileSize, TileSize, Tiles(Index).Src, 0, 0, SRCCOPY)
        'BitBlt the tile that user wanted to paint with
    picCurrent.Refresh
End Sub


Private Sub PaintTile(SrcX As Long, SrcY As Long, ImageNum As Integer)
    ret = BitBlt(picMap.hDC, SrcX, SrcY, TileSize, TileSize, Tiles(ImageNum).Src, 0, 0, SRCCOPY)
        'this BitBlts to the picMap, with the information
        'given when this sub was called
    frmMap.picMap.Refresh
        'refresh the picturebox, so that it could be
        'redrawn
        'sets the value of the coordinate printed,
        'this is essential because this is the information
        'that is being saved as a *.map file
End Sub

