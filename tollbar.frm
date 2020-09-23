VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Room Tool Bar"
   ClientHeight    =   5445
   ClientLeft      =   150
   ClientTop       =   315
   ClientWidth     =   1650
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   435
      Index           =   51
      Left            =   1080
      Picture         =   "tollbar.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   30
      ToolTipText     =   "Spiral Stairs"
      Top             =   2040
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   435
      Index           =   50
      Left            =   1080
      Picture         =   "tollbar.frx":0417
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   29
      ToolTipText     =   "Wild Magic Area"
      Top             =   1560
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   435
      Index           =   49
      Left            =   1080
      Picture         =   "tollbar.frx":062E
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   28
      ToolTipText     =   "Dead Magic Area"
      Top             =   1080
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   435
      Index           =   41
      Left            =   1080
      Picture         =   "tollbar.frx":0823
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   27
      ToolTipText     =   "Dirt Ground"
      Top             =   120
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   435
      Index           =   40
      Left            =   600
      Picture         =   "tollbar.frx":0A1E
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   26
      ToolTipText     =   "Teleport Area"
      Top             =   3480
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   435
      Index           =   39
      Left            =   600
      Picture         =   "tollbar.frx":0C22
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   25
      ToolTipText     =   "Treasure Chest"
      Top             =   3000
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   435
      Index           =   34
      Left            =   600
      Picture         =   "tollbar.frx":0EAE
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   24
      ToolTipText     =   "Treasure"
      Top             =   2520
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   435
      Index           =   33
      Left            =   600
      Picture         =   "tollbar.frx":11B8
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   23
      ToolTipText     =   "Curtian North/South"
      Top             =   2040
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   435
      Index           =   32
      Left            =   600
      Picture         =   "tollbar.frx":141D
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   22
      ToolTipText     =   "Curtian East/West"
      Top             =   1560
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   435
      Index           =   31
      Left            =   600
      Picture         =   "tollbar.frx":167C
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   21
      ToolTipText     =   "Rubble"
      Top             =   1080
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   435
      Index           =   30
      Left            =   600
      Picture         =   "tollbar.frx":18E9
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   20
      ToolTipText     =   "Sarcophagus Open"
      Top             =   600
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   435
      Index           =   28
      Left            =   600
      Picture         =   "tollbar.frx":1B9B
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   19
      ToolTipText     =   "Sarcophagus"
      Top             =   120
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   435
      Index           =   29
      Left            =   120
      Picture         =   "tollbar.frx":1DBC
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   18
      ToolTipText     =   "Alter"
      Top             =   3000
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   435
      Index           =   27
      Left            =   120
      Picture         =   "tollbar.frx":2039
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   17
      ToolTipText     =   "Trap"
      Top             =   3480
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   435
      Index           =   26
      Left            =   1080
      Picture         =   "tollbar.frx":227C
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   16
      ToolTipText     =   "Water"
      Top             =   600
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   435
      Index           =   25
      Left            =   120
      Picture         =   "tollbar.frx":2447
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   15
      ToolTipText     =   "Well"
      Top             =   2520
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   435
      Index           =   24
      Left            =   120
      Picture         =   "tollbar.frx":26E7
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   14
      ToolTipText     =   "Pool/Fountian"
      Top             =   2040
      Width           =   435
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Tile"
      ClipControls    =   0   'False
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   4320
      Width           =   975
      Begin VB.PictureBox picCurrent 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   420
         Left            =   240
         ScaleHeight     =   360
         ScaleWidth      =   405
         TabIndex        =   5
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "X, Y:"
         Height          =   195
         Left            =   1200
         TabIndex        =   13
         Top             =   360
         Width           =   345
      End
      Begin VB.Label lblXY 
         AutoSize        =   -1  'True
         Caption         =   "1,2"
         Height          =   195
         Left            =   2400
         TabIndex        =   12
         Top             =   360
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   195
         Left            =   1200
         TabIndex        =   11
         Top             =   600
         Width           =   405
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Left            =   2400
         TabIndex        =   10
         Top             =   600
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ImgNum:"
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   840
         Width           =   630
      End
      Begin VB.Label lblImgNum 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   195
         Left            =   2400
         TabIndex        =   8
         Top             =   840
         Width           =   90
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   1200
         TabIndex        =   7
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   195
         Left            =   2400
         TabIndex        =   6
         Top             =   1080
         Width           =   90
      End
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   435
      Index           =   20
      Left            =   120
      Picture         =   "tollbar.frx":29A1
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   3
      ToolTipText     =   "Statue"
      Top             =   120
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   435
      Index           =   22
      Left            =   120
      Picture         =   "tollbar.frx":2C91
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   2
      ToolTipText     =   "Closed Pit"
      Top             =   1080
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   435
      Index           =   23
      Left            =   120
      Picture         =   "tollbar.frx":2F57
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   1
      ToolTipText     =   "Pillar"
      Top             =   1560
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   435
      Index           =   21
      Left            =   120
      Picture         =   "tollbar.frx":31E5
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   0
      ToolTipText     =   "Open Pit"
      Top             =   600
      Width           =   435
   End
End
Attribute VB_Name = "Form1"
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

