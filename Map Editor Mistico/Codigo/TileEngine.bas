Attribute VB_Name = "TileEngine"
'Argentum Online 0.9.83
'Copyright (C) 2001 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based in Baronsoft's VB6 Online RPG
'Engine 9/08/2000 http://www.baronsoft.com/
'aaron@baronsoft.com
'
'Contact info:
'Pablo Ignacio Márquez
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900


Option Explicit



'********** CONSTANTS ***********
'Heading Constants
Public Const NORTH = 1
Public Const EAST = 2
Public Const SOUTH = 3
Public Const WEST = 4

'Map sizes in tiles
Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

'Object Constants
Public Const MAX_INVENORY_OBJS = 10000

'bltbit constant
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

'Sound flag constants
Public Const SND_SYNC = &H0 ' play synchronously (default)
Public Const SND_ASYNC = &H1 ' play asynchronously
Public Const SND_NODEFAULT = &H2 ' silence not default, if sound not found
Public Const SND_LOOP = &H8 ' loop the sound until next sndPlaySound
Public Const SND_NOSTOP = &H10 ' don't stop any currently playing sound

Public Const NumSoundBuffers = 7

'********** TYPES ***********

'Bitmap header
Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

'Bitmap info header
Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

'Holds a local position
Public Type Position
    x As Integer
    y As Integer
End Type

'Holds a world position
Public Type WorldPos
    Map As Integer
    x As Integer
    y As Integer
End Type

'Holds data about where a bmp can be found,
'How big it is and animation info
Public Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames(1 To 25) As Integer
    Speed As Integer
End Type

'Points to a grhData and keeps animation info
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Byte
    SpeedCounter As Byte
    Started As Byte
End Type

'Bodies list
Public Type BodyData
    Walk(1 To 4) As Grh
    HeadOffset As Position
End Type

'Heads list
Public Type HeadData
    Head(0 To 4) As Grh
End Type

'Hold info about a character
Public Type Char
    Active As Byte
    Heading As Byte
    pos As Position

    Body As BodyData
    Head As HeadData
    
    Moving As Byte
    MoveOffset As Position
    
End Type

'Holds info about a object
Public Type Obj
    objindex As Integer
    Amount As Integer
End Type

'Holds info about each tile position
Public Type MapBlock
    
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    Trigger As Integer
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
End Type

'Hold info about each map
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    Rest As String
    Zona As String
    Terreno As String
    
    'ME Only
    Changed As Byte
End Type




'Where the map borders are.. Set during load
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

Public gDespX As Integer
Public gDespY As Integer

'User status vars
Public CurMap As Integer 'Current map loaded
Public UserIndex As Integer
Public UserMoving As Byte
Global UserBody As Integer
Global UserHead As Integer
Public UserPos As Position 'Holds current user pos
Public AddtoUserPos As Position 'For moving user
Public UserCharIndex As Integer

Public EngineRun As Boolean
Public FramesPerSec As Integer
Public FramesPerSecCounter As Long

'Main view size size in tiles
Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

'Pixel offset of main view screen from 0,0
Public MainViewTop As Integer
Public MainViewLeft As Integer

'How many tiles the engine "looks ahead" when
'drawing the screen
Public TileBufferSize As Integer

'Handle to where all the drawing is going to take place
Public DisplayFormhWnd As Long

'Tile size in pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Map editor variables
Public WalkMode As Boolean
Public DrawGrid As Boolean
Public DrawBlock As Boolean

'Totals
Public NumMaps As Integer 'Number of maps
Public NumBodies As Integer
Public NumHeads As Integer
Public NumGrhFiles As Integer 'Number of bmps
Public NumGrhs As Integer 'Number of Grhs
Global NumChars As Integer
Global LastChar As Integer

'********** Direct X ***********
Public MainViewRect As RECT
Public MainViewWidth As Integer
Public MainViewHeight As Integer
Public BackBufferRect As RECT

Public DirectX As New DirectX7
Public DirectDraw As DirectDraw7

Public PrimarySurface As DirectDrawSurface7
Public PrimaryClipper As DirectDrawClipper
Public BackBufferSurface As DirectDrawSurface7
Public SurfaceDB() As DirectDrawSurface7

'Sound
Dim DirectSound As DirectSound
Dim DSBuffers(1 To NumSoundBuffers) As DirectSoundBuffer
Dim LastSoundBufferUsed As Integer

'********** Public ARRAYS ***********
Public GrhData() As GrhData 'Holds all the grh data

Public BodyData() As BodyData
Public HeadData() As HeadData

Public MapData() As MapBlock 'Holds map data for current map
Public MapInfo As MapInfo 'Holds map info for current map
Public CharList(1 To 10000) As Char 'Holds info about all characters on map


'********** OUTSIDE FUNCTIONS ***********
'Good old BitBlt
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

'Sound stuff
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uRetrunLength As Long, ByVal hwndCallback As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Function LoadWavetoDSBuffer(DS As DirectSound, DSB As DirectSoundBuffer, sfile As String) As Boolean


    '========================================================================
    '- Step4 CREATE SOUND BUFFER FROM FILE.
    '  we use the DSBUFFERDESC type to indicate
    '  what features we want the sound to have.
    '  The lFlags member can be used to enable 3d support,
    '  frequency changes, and volume changes.
    '  The DSBCAPS flags indicates we will allow
    '  volume changes, frequency changes, and pan changes
    '  the DDSBCAPS_STATIC -(which is optional in this release
    '  since all  buffers loaded by this method are static) indicates
    '  that we want the entire file loaded into memory.
    '
    '  The function fills in the other members of bufferDesc which lets
    '  us know how large the buffer is.  It also fills in the wave Format
    '  type giving information about the waves quality and if it supports
    '  stereo the function returns an initialized SoundBuffer
    '=========================================================================
    
    Dim bufferDesc As DSBUFFERDESC
    Dim waveFormat As WAVEFORMATEX
    
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2
    waveFormat.lSamplesPerSec = 22050
    waveFormat.nBitsPerSample = 16
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    Set DSB = DS.CreateSoundBufferFromFile(sfile, bufferDesc, waveFormat)
    
    '========================================
    '- Step 5 make sure we have no errors
    '========================================
    
    If Err.Number <> 0 Then
        'MsgBox "unable to find " + sfile
        'End
        Exit Function
    End If
    
    LoadWavetoDSBuffer = True
    
End Function
Sub LoadHeadData()
'*****************************************************************
'Loads Dats/Head.dat
'*****************************************************************

Dim LoopC As Integer

'Get Number of heads
NumHeads = Val(GetVar(IniDats & "Head.dat", "INIT", "NumHeads"))

'Resize array
ReDim HeadData(0 To NumHeads) As HeadData

'Fill List
For LoopC = 1 To NumHeads
    InitGrh HeadData(LoopC).Head(1), Val(GetVar(IniDats & "Head.dat", "Head" & LoopC, "Head1")), 0
    InitGrh HeadData(LoopC).Head(2), Val(GetVar(IniDats & "Head.dat", "Head" & LoopC, "Head2")), 0
    InitGrh HeadData(LoopC).Head(3), Val(GetVar(IniDats & "Head.dat", "Head" & LoopC, "Head3")), 0
    InitGrh HeadData(LoopC).Head(4), Val(GetVar(IniDats & "Head.dat", "Head" & LoopC, "Head4")), 0
Next LoopC

End Sub

Sub LoadBodyData()
'*****************************************************************
'Loads Dats/Body.dat
'*****************************************************************

Dim LoopC As Integer

'Get number of bodies
NumBodies = Val(GetVar(IniDats & "Body.dat", "INIT", "NumBodies"))

'Resize array
ReDim BodyData(1 To NumBodies) As BodyData

'Fill list
For LoopC = 1 To NumBodies
    InitGrh BodyData(LoopC).Walk(1), Val(GetVar(IniDats & "Body.dat", "Body" & LoopC, "Walk1")), 0
    InitGrh BodyData(LoopC).Walk(2), Val(GetVar(IniDats & "Body.dat", "Body" & LoopC, "Walk2")), 0
    InitGrh BodyData(LoopC).Walk(3), Val(GetVar(IniDats & "Body.dat", "Body" & LoopC, "Walk3")), 0
    InitGrh BodyData(LoopC).Walk(4), Val(GetVar(IniDats & "Body.dat", "Body" & LoopC, "Walk4")), 0

    BodyData(LoopC).HeadOffset.x = Val(GetVar(IniDats & "Body.dat", "Body" & LoopC, "HeadOffsetX"))
    BodyData(LoopC).HeadOffset.y = Val(GetVar(IniDats & "Body.dat", "Body" & LoopC, "HeadOffsetY"))

Next LoopC

End Sub
Sub ConvertCPtoTP(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal CX As Single, ByVal CY As Single, tX As Integer, tY As Integer)
'******************************************
'Converts where the user clicks in the main window
'to a tile position
'******************************************
Dim HWindowX As Integer
Dim HWindowY As Integer

CX = CX - StartPixelLeft
CY = CY - StartPixelTop

HWindowX = (WindowTileWidth \ 2)
HWindowY = (WindowTileHeight \ 2)

'Figure out X and Y tiles
CX = (CX \ TilePixelWidth)
CY = (CY \ TilePixelHeight)

If CX > HWindowX Then
    CX = (CX - HWindowX)

Else
    If CX < HWindowX Then
        CX = (0 - (HWindowX - CX))
    Else
        CX = 0
    End If
End If

If CY > HWindowY Then
    CY = (0 - (HWindowY - CY))
Else
    If CY < HWindowY Then
        CY = (CY - HWindowY)
    Else
        CY = 0
    End If
End If

tX = UserPos.x + CX
tY = UserPos.y + CY

End Sub




Function DeInitTileEngine() As Boolean
'*****************************************************************
'Shutsdown engine
'*****************************************************************
Dim LoopC As Integer

EngineRun = False

'****** Clear DirectX objects ******
Set PrimarySurface = Nothing
Set PrimaryClipper = Nothing
Set BackBufferSurface = Nothing

'Clear GRH memory
For LoopC = 1 To NumGrhFiles
    Set SurfaceDB(LoopC) = Nothing
Next LoopC
Set DirectDraw = Nothing

'Reset any channels that are done
For LoopC = 1 To NumSoundBuffers
    Set DSBuffers(LoopC) = Nothing
Next LoopC
Set DirectSound = Nothing

Set DirectX = Nothing

DeInitTileEngine = True

End Function

Sub MakeChar(CharIndex As Integer, Body As Integer, Head As Integer, Heading As Byte, x As Integer, y As Integer)
'*****************************************************************
'Makes a new character and puts it on the map
'*****************************************************************

'Update LastChar
If CharIndex > LastChar Then LastChar = CharIndex
NumChars = NumChars + 1

'Update head, body, ect.
CharList(CharIndex).Body = BodyData(Body)
CharList(CharIndex).Head = HeadData(Head)
CharList(CharIndex).Heading = Heading

'Reset moving stats
CharList(CharIndex).Moving = 0
CharList(CharIndex).MoveOffset.x = 0
CharList(CharIndex).MoveOffset.y = 0

'Update position
CharList(CharIndex).pos.x = x
CharList(CharIndex).pos.y = y

'Make active
CharList(CharIndex).Active = 1

'Plot on map
MapData(x, y).CharIndex = CharIndex

End Sub



Sub EraseChar(CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
If CharIndex = 0 Then Exit Sub
'Make un-active
CharList(CharIndex).Active = 0

'Update lastchar
If CharIndex = LastChar Then
    Do Until CharList(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If

'Remove from map
MapData(CharList(CharIndex).pos.x, CharList(CharIndex).pos.y).CharIndex = 0

'Update NumChars
NumChars = NumChars - 1

End Sub

Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
If GrhIndex = 0 Then Exit Sub
Grh.GrhIndex = GrhIndex

If Started = 2 Then
    If GrhData(Grh.GrhIndex).NumFrames > 1 Then
        Grh.Started = 1
    Else
        Grh.Started = 0
    End If
Else
    Grh.Started = Started
End If

Grh.FrameCounter = 1
Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed

End Sub

Sub MoveCharbyHead(CharIndex As Integer, nHeading As Byte)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
Dim addX As Integer
Dim addY As Integer
Dim x As Integer
Dim y As Integer
Dim nX As Integer
Dim nY As Integer

x = CharList(CharIndex).pos.x
y = CharList(CharIndex).pos.y

'Figure out which way to move
Select Case nHeading

    Case NORTH
        addY = -1

    Case EAST
        addX = 1

    Case SOUTH
        addY = 1
    
    Case WEST
        addX = -1
        
End Select

nX = x + addX
nY = y + addY

MapData(nX, nY).CharIndex = CharIndex
CharList(CharIndex).pos.x = nX
CharList(CharIndex).pos.y = nY
MapData(x, y).CharIndex = 0

CharList(CharIndex).MoveOffset.x = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.y = -1 * (TilePixelHeight * addY)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nHeading

End Sub

Sub MoveCharbyPos(CharIndex As Integer, nX As Integer, nY As Integer)
'*****************************************************************
'Starts the movement of a character to nX,nY
'*****************************************************************
Dim x As Integer
Dim y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nHeading As Byte

x = CharList(CharIndex).pos.x
y = CharList(CharIndex).pos.y

addX = nX - x
addY = nY - y

If Sgn(addX) = 1 Then
    nHeading = EAST
End If

If Sgn(addX) = -1 Then
    nHeading = WEST
End If

If Sgn(addY) = -1 Then
    nHeading = NORTH
End If

If Sgn(addY) = 1 Then
    nHeading = SOUTH
End If

MapData(nX, nY).CharIndex = CharIndex
CharList(CharIndex).pos.x = nX
CharList(CharIndex).pos.y = nY
MapData(x, y).CharIndex = 0

CharList(CharIndex).MoveOffset.x = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.y = -1 * (TilePixelHeight * addY)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nHeading

End Sub

Sub MoveScreen(Heading As Byte)
'******************************************
'Starts the screen moving in a direction
'******************************************
Dim x As Integer
Dim y As Integer
Dim tX As Integer
Dim tY As Integer

'Figure out which way to move
Select Case Heading

    Case NORTH
        y = -1

    Case EAST
        x = 1

    Case SOUTH
        y = 1
    
    Case WEST
        x = -1
        
End Select

'Fill temp pos
tX = UserPos.x + x
tY = UserPos.y + y

'Check to see if its out of bounds
If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
    Exit Sub
Else
    'Start moving... MainLoop does the rest
    AddtoUserPos.x = x
    UserPos.x = tX
    AddtoUserPos.y = y
    UserPos.y = tY
    UserMoving = 1
End If

End Sub


Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
Dim LoopC As Integer

LoopC = 1
Do While CharList(LoopC).Active
    LoopC = LoopC + 1
Loop

NextOpenChar = LoopC

End Function


Sub LoadGrhData()
'*****************************************************************
'Loads Grh.dat
'*****************************************************************

On Error GoTo ErrorHandler

Dim Grh As Integer
Dim Frame As Integer
Dim TempInt As Integer

'Get Number of Graphics
GrhPath = GetVar(IniPath & "Grh.ini", "INIT", "Path")
NumGrhs = Val(GetVar(IniPath & "Grh.ini", "INIT", "NumGrhs"))

'Resize arrays
ReDim GrhData(1 To NumGrhs) As GrhData

'Open files
Open IniPath & "Graficos.ind" For Binary Access Read As #1
Seek #1, 1

Get #1, , MiCabecera
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt

'Fill Grh List

'Get first Grh Number
Get #1, , Grh

Do Until Grh <= 0
        
    'Get number of frames
    Get #1, , GrhData(Grh).NumFrames
    If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
    
    If GrhData(Grh).NumFrames > 1 Then
    
        'Read a animation GRH set
        For Frame = 1 To GrhData(Grh).NumFrames
        
            Get #1, , GrhData(Grh).Frames(Frame)
            If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > NumGrhs Then GoTo ErrorHandler
        
        Next Frame
    
        Get #1, , GrhData(Grh).Speed
        If GrhData(Grh).Speed <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
        If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
        If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler
    
    Else
    
        'Read in normal GRH data
        Get #1, , GrhData(Grh).FileNum
        If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sX
        If GrhData(Grh).sX < 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sY
        If GrhData(Grh).sY < 0 Then GoTo ErrorHandler
            
        Get #1, , GrhData(Grh).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / TilePixelHeight
        GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / TilePixelWidth
        
        GrhData(Grh).Frames(1) = Grh
            
    End If

    'Get Next Grh Number
    Get #1, , Grh

Loop
'************************************************

Close #1

Exit Sub

ErrorHandler:
Close #1
MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh

End Sub

Function LegalPos(x As Integer, y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************

'Check to see if its out of bounds
'If X - 8 < 1 Or X - 8 > 100 Or Y - 6 < 1 Or Y - 6 > 100 Then
'    LegalPos = False
'    Exit Function
'End If

'Check to see if its blocked
'If MapData(X, Y).Blocked = 1 Then
'    LegalPos = False
'    Exit Function
'End If

'Check for character
'If MapData(X, Y).CharIndex > 0 Then
'    LegalPos = False
'    Exit Function
'End If

LegalPos = True

End Function




Function InMapLegalBounds(x As Integer, y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps
'LEGAL/Walkable bounds
'*****************************************************************

If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

Function InMapBounds(x As Integer, y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************

If x < XMinMapSize Or x > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function
Sub DDrawGrhtoSurface(Surface As DirectDrawSurface7, Grh As Grh, x As Integer, y As Integer, Center As Byte, Animate As Byte)
'*****************************************************************
'Draws a Grh at the X and Y positions
'*****************************************************************
Dim CurrentGrh As Grh
Dim DestRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2
On Error Resume Next
'Check to make sure it is legal
If Grh.GrhIndex < 1 Then
    Exit Sub
End If
If GrhData(Grh.GrhIndex).NumFrames < 1 Then
    Exit Sub
End If

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                End If
            End If
        End If
    End If
End If

'Figure out what frame to draw (always 1 if not animated)
If Grh.FrameCounter = 0 Then Exit Sub
CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If Center Then
    If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
        x = x - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
        y = y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With DestRect
    .Left = x
    .Top = y
    .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
    .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
End With
    
Surface.GetSurfaceDesc SurfaceDesc

'Draw

If DestRect.Left >= 0 And DestRect.Top >= 0 And DestRect.Right <= SurfaceDesc.lWidth And DestRect.Bottom <= SurfaceDesc.lHeight Then
    
    With SourceRect
        .Left = GrhData(CurrentGrh.GrhIndex).sX
        .Top = GrhData(CurrentGrh.GrhIndex).sY
        .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
    End With
    
    Surface.BltFast DestRect.Left, DestRect.Top, SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum), SourceRect, DDBLTFAST_WAIT
    
End If

End Sub

Sub DDrawTransGrhtoSurface(Surface As DirectDrawSurface7, Grh As Grh, x As Integer, y As Integer, Center As Byte, Animate As Byte)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
Dim CurrentGrh As Grh
Dim DestRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

'Check to make sure it is legal
If Grh.GrhIndex < 1 Then
    Exit Sub
End If
If GrhData(Grh.GrhIndex).NumFrames < 1 Then
    Exit Sub
End If

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                End If
            End If
        End If
    End If
End If

'Figure out what frame to draw (always 1 if not animated)
CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If Center Then
    If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
        x = x - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
        y = y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With DestRect
    .Left = x
    .Top = y
    .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
    .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
End With

Surface.GetSurfaceDesc SurfaceDesc

'Draw
If DestRect.Left >= 0 And DestRect.Top >= 0 And DestRect.Right <= SurfaceDesc.lWidth And DestRect.Bottom <= SurfaceDesc.lHeight Then
    With SourceRect
        .Left = GrhData(CurrentGrh.GrhIndex).sX
        .Top = GrhData(CurrentGrh.GrhIndex).sY
        .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
    End With
    
    Surface.BltFast DestRect.Left, DestRect.Top, SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If

End Sub

Sub DrawBackBufferSurface()
'*****************************************************************
'Copies backbuffer to primarysurface
'*****************************************************************
Dim SourceRect As RECT

With SourceRect
    .Left = (TilePixelWidth * TileBufferSize) - TilePixelWidth
    .Top = (TilePixelHeight * TileBufferSize) - TilePixelHeight
    .Right = .Left + MainViewWidth
    .Bottom = .Top + MainViewHeight
End With

PrimarySurface.Blt MainViewRect, BackBufferSurface, SourceRect, DDBLT_WAIT
'PrimarySurface.Flip Nothing, DDFLIP_WAIT

End Sub

Function GetBitmapDimensions(BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'*****************************************************************
'Gets the dimensions of a bmp
'*****************************************************************
Dim BMHeader As BITMAPFILEHEADER
Dim BINFOHeader As BITMAPINFOHEADER

Open BmpFile For Binary Access Read As #1
Get #1, , BMHeader
Get #1, , BINFOHeader
Close #1
bmWidth = BINFOHeader.biWidth
bmHeight = BINFOHeader.biHeight
End Function



Sub DrawGrhtoHdc(DestHdc As Long, Grh As Grh, x As Integer, y As Integer, Center As Byte, Animate As Byte, ROP As Long)
On Error Resume Next

'*****************************************************************
'Draws a Grh at the X and Y positions
'*****************************************************************
Dim retcode As Long
Dim CurrentGrh As Grh
Dim SourceHdc As Long


'Check to make sure it is legal
If Grh.GrhIndex < 1 Then
    Exit Sub
End If
If GrhData(Grh.GrhIndex).NumFrames < 1 Then
    Exit Sub
End If

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                End If
            End If
        End If
    End If
End If

'Figure out what frame to draw (always 1 if not animated)
CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If Center Then
    If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
        x = x - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
        y = y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

SourceHdc = SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum).GetDC

retcode = BitBlt(DestHdc, x, y, GrhData(CurrentGrh.GrhIndex).pixelWidth, GrhData(CurrentGrh.GrhIndex).pixelHeight, SourceHdc, GrhData(CurrentGrh.GrhIndex).sX, GrhData(CurrentGrh.GrhIndex).sY, ROP)

SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum).ReleaseDC SourceHdc

End Sub

Sub PlayWaveDS(File As String)

    'Cylce through avaiable sound buffers
    LastSoundBufferUsed = LastSoundBufferUsed + 1
    If LastSoundBufferUsed > NumSoundBuffers Then
        LastSoundBufferUsed = 1
    End If
    
    If LoadWavetoDSBuffer(DirectSound, DSBuffers(LastSoundBufferUsed), File) Then
        DSBuffers(LastSoundBufferUsed).Play DSBPLAY_DEFAULT
    End If

End Sub

Sub PlayWaveAPI(File As String)
'*****************************************************************
'Plays a Wave using windows APIs
'*****************************************************************
Dim rc As Integer

rc = sndPlaySound(File, SND_ASYNC)

End Sub
Sub RenderScreen(TileX As Integer, TileY As Integer, PixelOffsetX As Integer, PixelOffsetY As Integer)
'***********************************************
'Draw current visible to scratch area based on TileX and TileY
'***********************************************
Dim y As Integer    'Keeps track of where on map we are
Dim x As Integer
Dim minY As Integer 'Start Y pos on current map
Dim maxY As Integer 'End Y pos on current map
Dim minX As Integer 'Start X pos on current map
Dim maxX As Integer 'End X pos on current map
Dim ScreenX As Integer 'Keeps track of where to place tile on screen
Dim ScreenY As Integer
Dim PixelOffsetXTemp As Integer 'For centering grhs
Dim PixelOffsetYTemp As Integer
Dim Moved As Byte
Dim Grh As Grh 'Temp Grh for show tile and blocked
Dim TempChar As Char
Dim r As RECT
BackBufferSurface.BltColorFill r, 0

'Figure out Ends and Starts of screen
minY = (TileY - (WindowTileHeight \ 2)) - TileBufferSize
maxY = (TileY + (WindowTileHeight \ 2)) + TileBufferSize
minX = (TileX - (WindowTileWidth \ 2)) - TileBufferSize
maxX = (TileX + (WindowTileWidth \ 2)) + TileBufferSize

'Draw floor layer
ScreenY = 0
For y = minY To maxY
    ScreenX = 0
    For x = minX To maxX
        
        'Check to see if in bounds
        If InMapBounds(x, y) Then
    
            'Layer 1 **********************************
            
            PixelOffsetXTemp = PixelPos(ScreenX) + PixelOffsetX
            PixelOffsetYTemp = PixelPos(ScreenY) + PixelOffsetY
            
            'Draw
            Call DDrawGrhtoSurface(BackBufferSurface, MapData(x, y).Graphic(1), PixelPos(ScreenX) + PixelOffsetX, PixelPos(ScreenY) + PixelOffsetY, 0, 1)
            '**********************************
            
        End If
    
        ScreenX = ScreenX + 1
    Next x
    ScreenY = ScreenY + 1
Next y

'Draw floor layer 2
ScreenY = 0
For y = minY To maxY
    ScreenX = 0
    For x = minX To maxX

        'Check to see if in bounds
        If InMapBounds(x, y) Then

            'Layer 2 **********************************
            If MapData(x, y).Graphic(2).GrhIndex > 0 Then
            
                PixelOffsetXTemp = PixelPos(ScreenX) + PixelOffsetX
                PixelOffsetYTemp = PixelPos(ScreenY) + PixelOffsetY
            
                'Draw
                Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(x, y).Graphic(2), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                
            End If
            '**********************************
        End If
    
        ScreenX = ScreenX + 1
    Next x
    ScreenY = ScreenY + 1
Next y


'Draw transparent layers
ScreenY = 0
For y = minY To maxY
    ScreenX = 0
    For x = minX To maxX

        'Check to see if in bounds
        If InMapBounds(x, y) Then

            'Object Layer **********************************
            If MapData(x, y).ObjGrh.GrhIndex > 0 Then
            
                PixelOffsetXTemp = PixelPos(ScreenX) + PixelOffsetX
                PixelOffsetYTemp = PixelPos(ScreenY) + PixelOffsetY
            
            
                Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(x, y).ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
            
            End If
            '**********************************
            
            
             'Char layer **********************************
            If (MapData(x, y).CharIndex > 0) Then
            
                TempChar = CharList(MapData(x, y).CharIndex)
            
                PixelOffsetXTemp = PixelOffsetX
                PixelOffsetYTemp = PixelOffsetY
                
                Moved = 0
                'If needed, move left and right
                If TempChar.MoveOffset.x <> 0 Then
                        TempChar.Body.Walk(TempChar.Heading).Started = 1
                        PixelOffsetXTemp = PixelOffsetXTemp + TempChar.MoveOffset.x
                        TempChar.MoveOffset.x = TempChar.MoveOffset.x - (8 * Sgn(TempChar.MoveOffset.x))
                        Moved = 1
                End If
          
                'If needed, move up and down
                If TempChar.MoveOffset.y <> 0 Then
                        TempChar.Body.Walk(TempChar.Heading).Started = 1
                        PixelOffsetYTemp = PixelOffsetYTemp + TempChar.MoveOffset.y
                        TempChar.MoveOffset.y = TempChar.MoveOffset.y - (8 * Sgn(TempChar.MoveOffset.y))
                        Moved = 1
                End If
                
                'If done moving stop animation
                If Moved = 0 And TempChar.Moving = 1 Then
                    TempChar.Moving = 0
                    TempChar.Body.Walk(TempChar.Heading).FrameCounter = 1
                    TempChar.Body.Walk(TempChar.Heading).Started = 0
                End If
               
              'Dibuja solamente players
              If TempChar.Head.Head(TempChar.Heading).GrhIndex <> 0 Then
                'Draw Body
                Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Body.Walk(TempChar.Heading), (PixelPos(ScreenX) + PixelOffsetXTemp), PixelPos(ScreenY) + PixelOffsetYTemp, 1, 1)
                'Draw Head
                Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Head.Head(TempChar.Heading), (PixelPos(ScreenX) + PixelOffsetXTemp) + TempChar.Body.HeadOffset.x, PixelPos(ScreenY) + PixelOffsetYTemp + TempChar.Body.HeadOffset.y, 1, 0)
              Else: Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Body.Walk(TempChar.Heading), (PixelPos(ScreenX) + PixelOffsetXTemp), PixelPos(ScreenY) + PixelOffsetYTemp, 1, 1)
              End If
              
                
              
                'Refresh charlist
                CharList(MapData(x, y).CharIndex) = TempChar
                
            End If
            '**********************************
            
            
            'Layer 3 **********************************
            If MapData(x, y).Graphic(3).GrhIndex > 0 Then
            
                PixelOffsetXTemp = PixelPos(ScreenX) + PixelOffsetX
                PixelOffsetYTemp = PixelPos(ScreenY) + PixelOffsetY
            
                'Draw
                Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(x, y).Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
            
            End If
            '**********************************
            
        End If
    
        ScreenX = ScreenX + 1
    Next x
    ScreenY = ScreenY + 1
Next y


'Draw blocked tiles and grid
ScreenY = 0
For y = minY To maxY
    ScreenX = 0
    For x = minX To maxX
            
        'Check to see if in bounds
        If InMapBounds(x, y) Then
                                
            'Layer 4 **********************************
            If (MapData(x, y).Graphic(4).GrhIndex > 0) _
                 And (frmMain.Mostar4layer.value = vbChecked) Then
               
               PixelOffsetXTemp = PixelPos(ScreenX) + PixelOffsetX
               PixelOffsetYTemp = PixelPos(ScreenY) + PixelOffsetY
               Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(x, y).Graphic(4), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
               
            End If
            '**********************************
                                
            'Draw exit
            If MapData(x, y).TileExit.Map > 0 Then
                
                PixelOffsetXTemp = PixelPos(ScreenX) + PixelOffsetX
                PixelOffsetYTemp = PixelPos(ScreenY) + PixelOffsetY
                
                Grh.GrhIndex = 1
                Grh.FrameCounter = 1
                Grh.Started = 0
                Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, PixelOffsetXTemp, PixelOffsetYTemp, 0, 0)
            End If
                
            

            'Show blocked tiles
            If DrawBlock = True Then
                If MapData(x, y).Blocked = 1 Then
                
                    PixelOffsetXTemp = PixelPos(ScreenX) + PixelOffsetX
                    PixelOffsetYTemp = PixelPos(ScreenY) + PixelOffsetY
                
                    Grh.GrhIndex = 4
                    Grh.FrameCounter = 1
                    Grh.Started = 0
                    Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, PixelOffsetXTemp, PixelOffsetYTemp, 0, 0)
                End If
            End If
            
            If frmMain.Check3.value = vbChecked Then
                  Call DrawText(PixelPos(ScreenX), PixelPos(ScreenY), Str(MapData(x, y).Trigger), vbRed)
            End If
            

        End If
    
        ScreenX = ScreenX + 1
    Next x
    ScreenY = ScreenY + 1
Next y

End Sub

Public Sub DrawText(lngXPos As Integer, lngYPos As Integer, strText As String, lngColor As Long)
   If strText <> "" Then
    BackBufferSurface.SetFontTransparency True                    'Set the transparency flag to true
    BackBufferSurface.SetForeColor vbBlack                       'Set the color of the text to the color passed to the sub
    BackBufferSurface.SetFont frmMain.Font                    'Set the font used to the font on the form
    BackBufferSurface.DrawText lngXPos - 2, lngYPos - 1, strText, False 'Draw the text on to the screen, in the coordinates specified
    
    
    BackBufferSurface.SetFontTransparency True                    'Set the transparency flag to true
    BackBufferSurface.SetForeColor lngColor                       'Set the color of the text to the color passed to the sub
    BackBufferSurface.SetFont frmMain.Font                'Set the font used to the font on the form
    BackBufferSurface.DrawText lngXPos, lngYPos, strText, False   'Draw the text on to the screen, in the coordinates specified
   End If
End Sub

Function HayUserAbajo(x As Integer, y As Integer, GrhIndex) As Boolean
HayUserAbajo = _
    CharList(UserCharIndex).pos.x >= x - (GrhData(GrhIndex).TileWidth \ 2) _
And CharList(UserCharIndex).pos.x <= x + (GrhData(GrhIndex).TileWidth \ 2) _
And CharList(UserCharIndex).pos.y >= y - (GrhData(GrhIndex).TileHeight - 1) _
And CharList(UserCharIndex).pos.y <= y
End Function
Sub DibujarRect(ByVal x As Long, ByVal y As Long)


BackBufferSurface.SetForeColor (RGB(Val(frmGrilla.Col(0)), Val(frmGrilla.Col(1)), Val(frmGrilla.Col(2))))

BackBufferSurface.DrawLine x + gDespX, y + gDespY, x + Val(frmGrilla.Ancho) + gDespX, y + gDespY
BackBufferSurface.DrawLine x + gDespX, y + gDespY, x + gDespX, y + Val(frmGrilla.Alto) + gDespY
BackBufferSurface.DrawLine x + gDespX, y + Val(frmGrilla.Alto) + gDespY, x + gDespX, y + Val(frmGrilla.Alto) + gDespY
BackBufferSurface.DrawLine x + Val(frmGrilla.Ancho) + gDespX, y + gDespY, x + Val(frmGrilla.Ancho) + gDespX, y + Val(frmGrilla.Alto) + gDespY

End Sub

Sub HacerGrid()
Dim j As Integer, i As Integer
Dim x As Long, y As Long
Dim canty As Integer, cantx As Integer
canty = MainViewHeight \ Val(frmGrilla.Alto) + 4
cantx = MainViewWidth \ Val(frmGrilla.Ancho) + 4
y = 352
For j = 1 To canty
    x = 352

    For i = 1 To cantx
       Call DibujarRect(x, y)
       x = x + Val(frmGrilla.Ancho)
       DoEvents
    Next i
    y = y + Val(frmGrilla.Alto)
Next j

'BackBufferSurface.DrawBox 350, 350, 480, 480

End Sub

Function PixelPos(x As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'*****************************************************************

PixelPos = (TilePixelWidth * x) - TilePixelWidth

End Function

Sub LoadGraphics()
'*****************************************************************
'Loads all the sprites and tiles from the gif or bmp files
'*****************************************************************
Dim LoopC As Integer
Dim SurfaceDesc As DDSURFACEDESC2
Dim ddck As DDCOLORKEY
Dim ddsd As DDSURFACEDESC2

NumGrhFiles = Val(GetVar(IniPath & "Grh.ini", "INIT", "NumGrhFiles"))
ReDim SurfaceDB(1 To NumGrhFiles)



'Load the GRHx.bmps into memory
For LoopC = 1 To NumGrhFiles

    If FileExist(App.Path & GrhPath & LoopC & ".bmp", vbNormal) Then
        
        With ddsd
            .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        End With
        
        GetBitmapDimensions App.Path & GrhPath & LoopC & ".bmp", ddsd.lWidth, ddsd.lHeight
        
        Set SurfaceDB(LoopC) = DirectDraw.CreateSurfaceFromFile(App.Path & GrhPath & LoopC & ".bmp", ddsd)
        'Set color key
        ddck.low = 0
        ddck.high = 0
        SurfaceDB(LoopC).SetColorKey DDCKEY_SRCBLT, ddck
    End If
 
Next LoopC

End Sub
Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer, Optional re As Boolean) As Boolean
'*****************************************************************
'InitEngine
'*****************************************************************

Dim SurfaceDesc As DDSURFACEDESC2
Dim ddck As DDCOLORKEY


'Set intial user position
UserPos.x = MinXBorder
UserPos.y = MinYBorder

'Fill startup variables

DisplayFormhWnd = setDisplayFormhWnd
MainViewTop = setMainViewTop
MainViewLeft = setMainViewLeft
TilePixelWidth = setTilePixelWidth
TilePixelHeight = setTilePixelHeight
WindowTileHeight = setWindowTileHeight
WindowTileWidth = setWindowTileWidth
TileBufferSize = setTileBufferSize

'frmMain.HScroll1.Min = (WindowTileWidth \ 2) + 1
'frmMain.VScroll1.Min = (WindowTileHeight \ 2) + 1
'frmMain.HScroll1.Max = 100 - (WindowTileWidth \ 2)
'frmMain.VScroll1.Max = 100 - (WindowTileHeight \ 2)

MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

MainViewWidth = TilePixelWidth * WindowTileWidth
MainViewHeight = TilePixelHeight * WindowTileHeight

'Resize mapdata array
ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

'****** INIT DirectDraw ******
' Create the root DirectDraw object
Set DirectDraw = DirectX.DirectDrawCreate("")
DirectDraw.SetCooperativeLevel DisplayFormhWnd, DDSCL_NORMAL

'Primary Surface
' Fill the surface description structure
With SurfaceDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
End With
' Create the surface
Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)

'Create Primary Clipper
Set PrimaryClipper = DirectDraw.CreateClipper(0)
PrimaryClipper.SetHWnd frmMain.hWnd
PrimarySurface.SetClipper PrimaryClipper

'Back Buffer Surface
With BackBufferRect
    .Left = 0
    .Top = 0
    .Right = TilePixelWidth * (WindowTileWidth + (2 * TileBufferSize))
    .Bottom = TilePixelHeight * (WindowTileHeight + (2 * TileBufferSize))
End With
With SurfaceDesc
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    .lHeight = BackBufferRect.Bottom
    .lWidth = BackBufferRect.Right
End With

' Create surface
Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)

'Set color key
ddck.low = 0
ddck.high = 0
BackBufferSurface.SetColorKey DDCKEY_SRCBLT, ddck

If Not re Then
'Load graphic data into memory
Call LoadGrhData
Call LoadBodyData
Call LoadHeadData
Call LoadMapData
Call LoadGraphics
End If

'Wave Sound
Set DirectSound = DirectX.DirectSoundCreate("")
DirectSound.SetCooperativeLevel DisplayFormhWnd, DSSCL_PRIORITY
LastSoundBufferUsed = 1

InitTileEngine = True
EngineRun = True

End Function

Sub LoadMapData()
'*****************************************************************
'Load Map.dat
'*****************************************************************

'Get Number of Maps
NumMaps = Val(GetVar(IniDats & "Map.dat", "INIT", "NumMaps"))
MapPath = GetVar(IniDats & "Map.dat", "INIT", "MapPath")

End Sub
Sub ShowNextFrame(DisplayFormTop As Integer, DisplayFormLeft As Integer)
'***********************************************
'Updates and draws next frame to screen
'***********************************************
    Static OffsetCounterX As Integer
    Static OffsetCounterY As Integer

    '****** Set main view rectangle ******
    With MainViewRect
        .Left = (DisplayFormLeft / Screen.TwipsPerPixelX) + MainViewLeft
        .Top = (DisplayFormTop / Screen.TwipsPerPixelY) + MainViewTop
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With

    '***** Check if engine is allowed to run ******
    If EngineRun Then
        'Make sure noone goes above 30 FPS
        'If FramesPerSec <= 30 Then
        
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.x <> 0 Then
                OffsetCounterX = (OffsetCounterX - (8 * Sgn(AddtoUserPos.x)))
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.x) Then
                    OffsetCounterX = 0
                    AddtoUserPos.x = 0
                    UserMoving = 0
                End If
            End If

            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.y <> 0 Then
                OffsetCounterY = OffsetCounterY - (8 * Sgn(AddtoUserPos.y))
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.y = 0
                    UserMoving = 0
                End If
            End If

            '****** Update screen ******
            Call RenderScreen(UserPos.x - AddtoUserPos.x, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)
            'Draw grid
            If frmMain.DrawGridChk = vbChecked Then
                  Call HacerGrid
            End If
            DrawBackBufferSurface
            FramesPerSecCounter = FramesPerSecCounter + 1

        'End If
    End If

End Sub


