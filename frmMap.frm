VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMap 
   AutoRedraw      =   -1  'True
   Caption         =   "LCS_Dungeon Map Maker.."
   ClientHeight    =   6435
   ClientLeft      =   405
   ClientTop       =   495
   ClientWidth     =   9480
   ClipControls    =   0   'False
   Icon            =   "frmMap.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   549
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      Height          =   8265
      Left            =   0
      ScaleHeight     =   547
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   786
      TabIndex        =   0
      Top             =   0
      Width           =   11850
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6120
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Tile"
      Height          =   1575
      Left            =   8280
      TabIndex        =   1
      Top             =   4680
      Width           =   1095
      Begin VB.PictureBox picCurrent 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   780
         Left            =   120
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   2
         Top             =   480
         Width           =   780
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   195
         Left            =   2400
         TabIndex        =   10
         Top             =   1080
         Width           =   90
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   1080
         Width           =   465
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ImgNum:"
         Height          =   195
         Left            =   1200
         TabIndex        =   7
         Top             =   840
         Width           =   630
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Left            =   2400
         TabIndex        =   6
         Top             =   600
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   195
         Left            =   1200
         TabIndex        =   5
         Top             =   600
         Width           =   405
      End
      Begin VB.Label lblXY 
         AutoSize        =   -1  'True
         Caption         =   "1,2"
         Height          =   195
         Left            =   2400
         TabIndex        =   4
         Top             =   360
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "X, Y:"
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   345
      End
   End
   Begin VB.Menu mnuopt 
      Caption         =   "&Options"
      Begin VB.Menu mnunew 
         Caption         =   "&New"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuopen 
         Caption         =   "&Open"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnusave 
         Caption         =   "&Save"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuthis23 
         Caption         =   "-"
      End
      Begin VB.Menu mnutollbar 
         Caption         =   "&ToolBars"
         Begin VB.Menu mnufloor 
            Caption         =   "&Floor"
            Begin VB.Menu mnuhide 
               Caption         =   "&Hide"
               Shortcut        =   {F4}
            End
            Begin VB.Menu mnushow 
               Caption         =   "&Show"
               Shortcut        =   {F5}
            End
         End
         Begin VB.Menu mnuroom 
            Caption         =   "&Room"
            Begin VB.Menu mnurhide 
               Caption         =   "&Hide"
               Shortcut        =   {F6}
            End
            Begin VB.Menu mnurshow 
               Caption         =   "&Show"
               Shortcut        =   {F7}
            End
         End
      End
      Begin VB.Menu nmuthis 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XPos As Long
    'this carries the X information which the tiles
    'will be BitBlted on
Dim YPos As Long
    'this carries the Y information which the tiles
    'will be BitBlted on

Private Sub Form_Load()

Form2.Show
Form1.Show
Form1.Visible = False

    Call CreateTiles
        'call the sub "CreateTiles" to create tile info
    picTile_Click (0)
        'call the sub that activates when it's clicked
'    picCurrent.Picture = Form2.picTile(0).Picture
        'set the current picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload Form2
 Unload Form1
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuhide_Click()
Form2.Visible = False
End Sub

Private Sub mnunew_Click()
picMap.Picture = LoadPicture("")
End Sub

Private Sub mnuopen_Click()
On Error GoTo openerr

CommonDialog1.Filter = "BitMap Files (*.bmp)| *.bmp"
CommonDialog1.ShowOpen
picMap.Picture = LoadPicture(CommonDialog1.FileName)
frmMap.Caption = "LCS_Dungeon Map Maker.. " & CommonDialog1.FileName

openerr:
Exit Sub

End Sub

Private Sub mnurhide_Click()
Form1.Visible = False
End Sub

Private Sub mnurshow_Click()
Form1.Visible = True
End Sub

Private Sub mnusave_Click()
Whereisgenerated:
 Randomize Timer
 Mypicture = App.Path & "\Map" & CStr(Int(Rnd * 500)) & ".bmp"
  If Dir(Mypicture) <> "" Then GoTo Whereisgenerated
  SavePicture picMap.Image, Mypicture
 frmMap.Caption = "LCS_Dungeon Map Maker.. " & Mypicture

End Sub


Private Sub mnushow_Click()
Form2.Visible = True
End Sub

Private Sub picMap_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim retval As String
    'used for InputBox
    
Dim i As Integer
    'set i as integer

    i = Val(lblImgNum)
        'set i's value, it turns the string in
        'lblImgNum into an integer
    
    XPos = XPos * TileSize
        'XPos * TileSize because each tile will be
        'BitBlted in a different X value, so whereever
        'the mouse's X position is, it was divided
        'by 64, but in integer only (no decimals), and
        'then multiply it by the TileSize (64), because
        'remember, we need to BitBlt it to the right
        'spot, making it right next to each other
        
    YPos = YPos * TileSize
        'samething as the above description, except
        'replace X with Y

Select Case Val(lblType)
Case 2
    retval = InputBox("Please type in the map this door will lead to:")
        'the input box asking for user's response
    Call PaintTile(XPos, YPos, i)
        'call the sub PaintTile
    WorldMap(XPos \ TileSize, YPos \ TileSize).Goto = retval
        'set the Goto value, makes sure it loads
        'the map
Case Else
    Call PaintTile(XPos, YPos, i)
        'call the sub PaintTile
End Select
End Sub

Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
XPos = (x \ TileSize)
    'XPos is the current mouse's X value, but divided
    'by the TileSize, returning INTEGER only, that means
    'no decimals! That is what the sign "\" is for,
    'instead of "/" that you usually see
    
YPos = (y \ TileSize)
    'same as above, replace X with Y

    lblXY = XPos & ", " & YPos
        'this gives the label the information of the
        'current square the mouse is over

If Button = 1 Then
    'if the user has the mouse button down, then
    Call PaintTile(XPos * TileSize, YPos * TileSize, Val(lblImgNum))
        'call the sub PaintTile
Else    'otherwise
    Exit Sub
        'exit current sub
End If

End Sub

Private Sub picTile_Click(Index As Integer)
     lblName = Tiles(Index).Name
        'view the name of the current tile
    lblType = Tiles(Index).Type
        'view the type of the current tile
    lblImgNum = Tiles(Index).ImageNum
        'view the ImageNum of the current tile
    ret = BitBlt(Form2.picCurrent.hDC, 0, 0, TileSize, TileSize, Tiles(Index).Src, 0, 0, SRCCOPY)
ret = BitBlt(Form1.picCurrent.hDC, 0, 0, TileSize, TileSize, Tiles(Index).Src, 0, 0, SRCCOPY)        'BitBlt the tile that user wanted to paint with
                picCurrent.Refresh
                'picCurrent.Picture = Form2.picCurrent.Image
        'refreshes the new image
End Sub

Private Sub PaintTile(SrcX As Long, SrcY As Long, ImageNum As Integer)
    ret = BitBlt(picMap.hDC, SrcX, SrcY, TileSize, TileSize, Tiles(ImageNum).Src, 0, 0, SRCCOPY)
        'this BitBlts to the picMap, with the information
        'given when this sub was called
    picMap.Refresh
        'refresh the picturebox, so that it could be
        'redrawn
        'sets the value of the coordinate printed,
        'this is essential because this is the information
        'that is being saved as a *.map file
End Sub
