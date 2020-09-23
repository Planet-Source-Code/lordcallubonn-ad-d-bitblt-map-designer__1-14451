Attribute VB_Name = "modMap"
Dim ret As Long

Public Const TileSize = 25
    'the size of each tile

Type Tile
    Type As String
        'type, (ie: 0 = non walkable)
    ImageNum As Integer
        'the index of the picturebox
    Name As String
        'string that labels the tile (ie: grass)
    Src As Long
        'the handle to bitblt from
    Goto As String
        'uses only if it's a door
End Type

Public Tiles(0 To 51) As Tile
    'total of 3 tiles being represented

Public WorldMap(8, 8) As Tile
    'the world map consists of 8 x 8

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    'bitblt api
Public Const SRCCOPY = &HCC0020
    'used to copy from hDC to another

Public Sub CreateTiles()
'this sub just basically creates all the necessary tiles
'and their info. The map is basically saved as the
'picture's index number, that way, we could also load
'from it
        
    Tiles(0).Src = Form2.picTile(0).hDC
        'this is the hDC that will be BitBlted from

    Tiles(1).ImageNum = 1
        Tiles(1).Name = "Water"
        'sets the name of the tile as "Water"
    Tiles(1).Src = Form2.picTile(1).hDC
        'this is the hDC that will be BitBlted from
        
    Tiles(2).ImageNum = 2
    
    Tiles(2).Src = Form2.picTile(2).hDC
        'this is the hDC that will be BitBlted from
        
        Tiles(3).ImageNum = 3
    
    Tiles(3).Src = Form2.picTile(3).hDC
    
    Tiles(4).ImageNum = 4
    
    Tiles(4).Src = Form2.picTile(4).hDC
    
        Tiles(5).ImageNum = 5
    
    Tiles(5).Src = Form2.picTile(5).hDC
    
    Tiles(6).ImageNum = 6
    
    Tiles(6).Src = Form2.picTile(6).hDC
    
        Tiles(7).ImageNum = 7
    
    Tiles(7).Src = Form2.picTile(7).hDC
    
    Tiles(8).ImageNum = 8
    
    Tiles(8).Src = Form2.picTile(8).hDC
    
    Tiles(9).ImageNum = 9
    
    Tiles(9).Src = Form2.picTile(9).hDC
    
    Tiles(10).ImageNum = 10
    
    Tiles(10).Src = Form2.picTile(10).hDC
    
    Tiles(11).ImageNum = 11
    
    Tiles(11).Src = Form2.picTile(11).hDC
    
    Tiles(12).ImageNum = 12
    
    Tiles(12).Src = Form2.picTile(12).hDC
    
        Tiles(13).ImageNum = 13
    
    Tiles(13).Src = Form2.picTile(13).hDC

        Tiles(14).ImageNum = 14
    
    Tiles(14).Src = Form2.picTile(14).hDC
    
        Tiles(15).ImageNum = 15
    
    Tiles(15).Src = Form2.picTile(15).hDC
    
            Tiles(16).ImageNum = 16
    
    Tiles(16).Src = Form2.picTile(16).hDC

            Tiles(17).ImageNum = 17
    
    Tiles(17).Src = Form2.picTile(17).hDC
    
            Tiles(18).ImageNum = 18
    
    Tiles(18).Src = Form2.picTile(18).hDC
    
                Tiles(19).ImageNum = 19
    
    Tiles(19).Src = Form2.picTile(19).hDC
            
            Tiles(20).ImageNum = 20
    
    Tiles(20).Src = Form1.picTile(20).hDC

            Tiles(21).ImageNum = 21
    
    Tiles(21).Src = Form1.picTile(21).hDC
    
            Tiles(22).ImageNum = 22
    
    Tiles(22).Src = Form1.picTile(22).hDC

            Tiles(23).ImageNum = 23
    
    Tiles(23).Src = Form1.picTile(23).hDC
    
            Tiles(24).ImageNum = 24
    
    Tiles(24).Src = Form1.picTile(24).hDC
    
                Tiles(25).ImageNum = 25
    
    Tiles(25).Src = Form1.picTile(25).hDC

                Tiles(26).ImageNum = 26
    
    Tiles(26).Src = Form1.picTile(26).hDC


                Tiles(27).ImageNum = 27
    
    Tiles(27).Src = Form1.picTile(27).hDC

                Tiles(28).ImageNum = 28
    
    Tiles(28).Src = Form1.picTile(28).hDC

                Tiles(29).ImageNum = 29
    
    Tiles(29).Src = Form1.picTile(29).hDC

                Tiles(30).ImageNum = 30
    
    Tiles(30).Src = Form1.picTile(30).hDC

Tiles(31).ImageNum = 31
    
    Tiles(31).Src = Form1.picTile(31).hDC

Tiles(32).ImageNum = 32
    
    Tiles(32).Src = Form1.picTile(32).hDC

Tiles(33).ImageNum = 33
    
    Tiles(33).Src = Form1.picTile(33).hDC

Tiles(34).ImageNum = 34
    
    Tiles(34).Src = Form1.picTile(34).hDC

Tiles(35).ImageNum = 35
    
    Tiles(35).Src = Form2.picTile(35).hDC

Tiles(36).ImageNum = 36
    
    Tiles(36).Src = Form2.picTile(36).hDC

Tiles(37).ImageNum = 37
    
    Tiles(37).Src = Form2.picTile(37).hDC

Tiles(38).ImageNum = 38
    
    Tiles(38).Src = Form2.picTile(38).hDC


Tiles(39).ImageNum = 39
    
    Tiles(39).Src = Form1.picTile(39).hDC


Tiles(40).ImageNum = 40
    
    Tiles(40).Src = Form1.picTile(40).hDC


Tiles(41).ImageNum = 41
    
    Tiles(41).Src = Form1.picTile(41).hDC

Tiles(42).ImageNum = 42
    
    Tiles(42).Src = Form2.picTile(42).hDC

Tiles(43).ImageNum = 43
    
    Tiles(43).Src = Form2.picTile(43).hDC

Tiles(44).ImageNum = 44
    
    Tiles(44).Src = Form2.picTile(44).hDC

Tiles(45).ImageNum = 45
    
    Tiles(45).Src = Form2.picTile(45).hDC

Tiles(46).ImageNum = 46
    
    Tiles(46).Src = Form2.picTile(46).hDC


Tiles(47).ImageNum = 47
    
    Tiles(47).Src = Form2.picTile(47).hDC

Tiles(48).ImageNum = 48
    
    Tiles(48).Src = Form2.picTile(48).hDC

Tiles(49).ImageNum = 49
    
    Tiles(49).Src = Form1.picTile(49).hDC
    
    Tiles(50).ImageNum = 50
    
    Tiles(50).Src = Form1.picTile(50).hDC

Tiles(51).ImageNum = 51
    
    Tiles(51).Src = Form1.picTile(51).hDC

End Sub
