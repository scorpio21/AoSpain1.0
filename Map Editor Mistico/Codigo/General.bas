Attribute VB_Name = "General"
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

'For Get and Write Var
Declare Function writeprivateprofilestring Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function getprivateprofilestring Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'For KeyInput
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer


Sub SwitchMap(Map As String)
'*****************************************************************
'Loads and switches to a new room
'*****************************************************************
Dim LoopC As Integer
Dim TempInt As Integer
Dim Body As Integer
Dim Head As Integer
Dim Heading As Byte
Dim y As Integer
Dim x As Integer
   
'Change mouse icon
frmMain.MousePointer = 11
   
'Open files
Open Map For Binary As #1


Seek #1, 1

Map = Left(Map, Len(Map) - 4)
Map = Map & ".inf"
Open Map For Binary As #2
Seek #2, 1

'Cabecera map
Get #1, , MapInfo.MapVersion
Get #1, , MiCabecera
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt

'Cabecera inf
Get #2, , TempInt
Get #2, , TempInt
Get #2, , TempInt
Get #2, , TempInt
Get #2, , TempInt


'Load arrays
For y = YMinMapSize To YMaxMapSize
    For x = XMinMapSize To XMaxMapSize

        '.map file
        Get #1, , MapData(x, y).Blocked
        For LoopC = 1 To 4
            Get #1, , MapData(x, y).Graphic(LoopC).GrhIndex
            
            'Set up GRH
            If MapData(x, y).Graphic(LoopC).GrhIndex > 0 Then

                InitGrh MapData(x, y).Graphic(LoopC), MapData(x, y).Graphic(LoopC).GrhIndex

            End If
        
        Next LoopC
        'Trigger
        Get #1, , MapData(x, y).Trigger
        
        Get #1, , TempInt
        '.inf file
        
        'Tile exit
        Get #2, , MapData(x, y).TileExit.Map
        Get #2, , MapData(x, y).TileExit.x
        Get #2, , MapData(x, y).TileExit.y
                      
        'make NPC
        Get #2, , MapData(x, y).NPCIndex
        If MapData(x, y).NPCIndex > 0 Then
            
            If MapData(x, y).NPCIndex > 499 Then
                Body = Val(GetVar(IniDats & "NPCs-HOSTILES.dat", "NPC" & MapData(x, y).NPCIndex, "Body"))
                Head = Val(GetVar(IniDats & "NPCs-HOSTILES.dat", "NPC" & MapData(x, y).NPCIndex, "Head"))
                Heading = Val(GetVar(IniDats & "NPCs-HOSTILES.dat", "NPC" & MapData(x, y).NPCIndex, "Heading"))
            Else
                Body = Val(GetVar(IniDats & "NPCs.dat", "NPC" & MapData(x, y).NPCIndex, "Body"))
                Head = Val(GetVar(IniDats & "NPCs.dat", "NPC" & MapData(x, y).NPCIndex, "Head"))
                Heading = Val(GetVar(IniDats & "NPCs.dat", "NPC" & MapData(x, y).NPCIndex, "Heading"))
            End If
            
            Call MakeChar(NextOpenChar(), Body, Head, Heading, x, y)
        End If
        
        'Make obj
        Get #2, , MapData(x, y).OBJInfo.objindex
        Get #2, , MapData(x, y).OBJInfo.Amount
        If MapData(x, y).OBJInfo.objindex > 0 Then
            InitGrh MapData(x, y).ObjGrh, Val(GetVar(IniDats & "OBJ.dat", "OBJ" & MapData(x, y).OBJInfo.objindex, "GrhIndex"))
        End If
        
        'Empty place holders for future expansion
        Get #2, , TempInt
        Get #2, , TempInt
             
    Next x
Next y

'Close files
Close #1
Close #2



Map = Right(Map, Len(Map) - Len(IniMaps))

Map = Left(Map, Len(Map) - 4) & ".dat"

MapInfo.Name = GetVar(IniMaps & Map, Left(Map, Len(Map) - 4), "Name")
MapInfo.Music = GetVar(IniMaps & Map, Left(Map, Len(Map) - 4), "MusicNum")
MapInfo.StartPos.Map = Val(ReadField(1, GetVar(IniMaps & Map, Left(Map, Len(Map) - 4), "StartPos"), 45))
MapInfo.StartPos.x = Val(ReadField(2, GetVar(Map, Left(Map, Len(Map) - 4), "StartPos"), 45))
MapInfo.StartPos.y = Val(ReadField(3, GetVar(Map, Left(Map, Len(Map) - 4), "StartPos"), 45))
frmMain.Text2.Text = MapInfo.Music


MapInfo.Terreno = GetVar(IniMaps & Map, Left(Map, Len(Map) - 4), "Terreno")

If MapInfo.Terreno = "BOSQUE" Then
 frmCarac.Option1(0).value = True
ElseIf MapInfo.Terreno = "DESIERTO" Then
 frmCarac.Option1(1).value = True
Else
 frmCarac.Option1(2).value = True
End If

MapInfo.Zona = GetVar(IniMaps & Map, Left(Map, Len(Map) - 4), "Zona")
If MapInfo.Zona = "CIUDAD" Then
 frmCarac.Option2(0).value = True
ElseIf MapInfo.Zona = "DUNGEON" Then
 frmCarac.Option2(1).value = True
Else
 frmCarac.Option2(2).value = True
End If

MapInfo.Rest = GetVar(IniMaps & Map, Left(Map, Len(Map) - 4), "Restringir")

If MapInfo.Rest = "Si" Then
    frmCarac.Check1.value = vbChecked
Else
    frmCarac.Check1.value = vbUnchecked
End If

'CurMap = Map
frmMain.Text1.Text = MapInfo.Name
frmMain.Vers.Text = MapInfo.MapVersion

'Set changed flag
MapInfo.Changed = 0

'Change mouse icon
frmMain.MousePointer = 0
MapaCargado = True
frmMain.mAncho.SetFocus
End Sub
Sub ActualizaDespGrilla()
If UserPos.x - 8 < 1 Or UserPos.y - 6 < 1 Then Exit Sub
Dim i As Integer, j As Integer
gDespX = 0
gDespY = 0
j = 1
   
If UserPos.y - 6 <> 1 Then
   Do While j <> UserPos.y - 6
        gDespY = gDespY - 32
        If gDespY = -Val(frmGrilla.Alto) Then gDespY = 0
        j = j + 1
   Loop
End If

i = 1
If UserPos.x - 8 <> 1 Then
    Do While i <> UserPos.x - 8
        gDespX = gDespX - 32
        If gDespX = -Val(frmGrilla.Ancho) Then gDespX = 0
        i = i + 1
    Loop
End If
End Sub
Sub CheckKeys()


    If GetKeyState(vbKeyUp) < 0 Then
        
        If LegalPos(UserPos.x, UserPos.y - 1) Then
            UserPos.y = UserPos.y - 1
            ActualizaDespGrilla
            frmMain.Apuntador.Move UserPos.x - 8, UserPos.y - 6
            frmMain.SetFocus
        End If
        
        Exit Sub
    End If

    If GetKeyState(vbKeyRight) < 0 Then
        If LegalPos(UserPos.x + 1, UserPos.y) Then
            UserPos.x = UserPos.x + 1
            ActualizaDespGrilla
            frmMain.Apuntador.Move UserPos.x - 8, UserPos.y - 6
            frmMain.SetFocus
        End If
        Exit Sub
    End If

    If GetKeyState(vbKeyDown) < 0 Then
        If LegalPos(UserPos.x, UserPos.y + 1) Then
            UserPos.y = UserPos.y + 1
            ActualizaDespGrilla
            frmMain.Apuntador.Move UserPos.x - 8, UserPos.y - 6
            frmMain.SetFocus
        End If
        Exit Sub
    End If

    If GetKeyState(vbKeyLeft) < 0 Then
        If LegalPos(UserPos.x - 1, UserPos.y) Then
            UserPos.x = UserPos.x - 1
            ActualizaDespGrilla
            frmMain.Apuntador.Move UserPos.x - 8, UserPos.y - 6
            frmMain.SetFocus
        End If
        Exit Sub
    End If
    

End Sub



Sub ReacttoMouseClick(Button As Integer, tX As Integer, tY As Integer)

'*****************************************************************
'React to mouse button
'*****************************************************************
Dim LoopC As Integer
Dim NPCIndex As Integer
Dim objindex As Integer
Dim Head As Integer
Dim Body As Integer
Dim Heading As Byte
On Error GoTo ErrHandler

'Right
If Button = vbRightButton Then
    
    'Show Info
    
    'Position
    frmMain.StatTxt.Text = frmMain.StatTxt.Text & ENDL & "Posicion " & tX & "," & tY & "  Bloqueada=" & MapData(tX, tY).Blocked
    
    'Exits
    If MapData(tX, tY).TileExit.Map > 0 Then
        frmMain.StatTxt.Text = frmMain.StatTxt.Text & ENDL & "Salida a: " & MapData(tX, tY).TileExit.Map & "," & MapData(tX, tY).TileExit.x & "," & MapData(tX, tY).TileExit.y
    End If
    
    'NPCs
    If MapData(tX, tY).NPCIndex > 0 Then
        If MapData(tX, tY).NPCIndex > 499 Then
            frmMain.StatTxt.Text = frmMain.StatTxt.Text & ENDL & "NPC: " & GetVar(IniDats & "NPCs-HOSTILES.dat", "NPC" & MapData(tX, tY).NPCIndex, "Name")
        Else
            frmMain.StatTxt.Text = frmMain.StatTxt.Text & ENDL & "NPC: " & GetVar(IniDats & "NPCs.dat", "NPC" & MapData(tX, tY).NPCIndex, "Name")
        End If
    End If
    
    'OBJs
    If MapData(tX, tY).OBJInfo.objindex > 0 Then
        frmMain.StatTxt.Text = frmMain.StatTxt.Text & ENDL & "OBJ" & MapData(tX, tY).OBJInfo.objindex & ": " & GetVar(IniDats & "OBJ.dat", "OBJ" & MapData(tX, tY).OBJInfo.objindex, "Name") & "   Cantidad=" & MapData(tX, tY).OBJInfo.Amount
    End If
    
    'Append
    frmMain.StatTxt.Text = frmMain.StatTxt.Text & ENDL
    frmMain.StatTxt.SelStart = Len(frmMain.StatTxt.Text)
    
    If frmMain.EM.value = 1 Then
        EY = tY
        EX = tX
    End If
    
    
    Exit Sub
End If


'Left click
If Button = vbLeftButton Then

    '************** Place grh
    If frmMain.PlaceGrhCmd.Enabled = False Then
      
        
        'Erase 2-3
        If frmMain.EraseAllchk.value = 1 Then
            For LoopC = 2 To 3
                MapData(tX, tY).Graphic(LoopC).GrhIndex = 0
            Next LoopC
            Exit Sub
        End If

        'Erase layer
        If frmMain.Erasechk.value = 1 Then
        
            If Val(frmMain.Layertxt.Text) = 1 Then
                MsgBox "No puedo borrar el layer 1!"
                Exit Sub
            End If
            
       MapData(tX, tY).Graphic(Val(frmMain.Layertxt.Text)).GrhIndex = 0
            Exit Sub
        End If
        If frmMain.MOSAICO.value = vbChecked Then
          Dim aux As Integer
          Dim dy As Integer
          Dim DX As Integer
          
          
         If frmMain.EM.value = 1 Then
            MapData(EX, EY).TileExit.Map = Val(frmHerramientas.MapExitTxt.Text)
            MapData(EX, EY).TileExit.x = Val(tX)
            MapData(EX, EY).TileExit.y = Val(tY)
            objindex = 378
            InitGrh MapData(EX, EY).ObjGrh, Val(GetVar(IniDats & "OBJ.dat", "OBJ" & objindex, "GrhIndex"))
            MapData(EX, EY).OBJInfo.objindex = objindex
            MapData(EX, EY).OBJInfo.Amount = 1
            Exit Sub
        End If
          
          If frmMain.DespMosaic.value = vbChecked Then
                        dy = Val(frmMain.DMLargo)
                        DX = Val(frmMain.DMAncho.Text)
          Else
                    dy = 0
                    DX = 0
          End If
                
          If frmMain.Completar = vbUnchecked Then
                aux = Val(frmMain.Grhtxt.Text) + _
                (((tY + dy) Mod frmMain.mLargo) * frmMain.mAncho) + ((tX + DX) Mod frmMain.mAncho)
                 MapData(tX, tY).Blocked = frmMain.Blockedchk.value
                 MapData(tX, tY).Graphic(Val(frmMain.Layertxt.Text)).GrhIndex = aux
                 InitGrh MapData(tX, tY).Graphic(Val(frmMain.Layertxt.Text)), aux
          Else
            Dim tXX As Integer, tYY As Integer, i As Integer, j As Integer, desptile As Integer
            tXX = tX
            tYY = tY
            desptile = 0
            For i = 1 To frmMain.mLargo
                For j = 1 To frmMain.mAncho
                    aux = Val(frmMain.Grhtxt.Text) + desptile
                     
                     MapData(tXX, tYY).Blocked = frmMain.Blockedchk.value
                     
                     MapData(tXX, tYY).Graphic(Val(frmMain.Layertxt.Text)).GrhIndex = aux
                     
                     InitGrh MapData(tXX, tYY).Graphic(Val(frmMain.Layertxt.Text)), aux
                     tXX = tXX + 1
                     desptile = desptile + 1
                Next
                tXX = tX
                tYY = tYY + 1
            Next
            tYY = tY
                
                
          End If
          
        Else
            'Else Place graphic
            If tX < 1 Or tX > 100 Then Exit Sub
            If tY < 1 Or tY > 100 Then Exit Sub
            MapData(tX, tY).Blocked = frmMain.Blockedchk.value
            MapData(tX, tY).Graphic(Val(frmMain.Layertxt.Text)).GrhIndex = Val(frmMain.Grhtxt.Text)
            
            
            'Setup GRH
    
            InitGrh MapData(tX, tY).Graphic(Val(frmMain.Layertxt.Text)), Val(frmMain.Grhtxt.Text)
        End If
        
        
        
        
    End If
    
    '************** Place blocked tile
    If frmMain.PlaceBlockCmd.Enabled = False Then
        MapData(tX, tY).Blocked = frmMain.Blockedchk.value
    End If
    
    '4x4 Sin borrar nada
    If frmMain.Blo4.value = 1 Then
        Dim valo1 As Byte
        Dim valo2 As Byte
        
        For valo1 = 1 To 4
            For valo2 = 1 To 4
                MapData(tX + (valo1 - 1), tY + (valo2 - 1)).Blocked = 1
            Next
        Next
    End If
    
    '4x4 de bloqueos Con Vacio
    If frmMain.ByV.value = 1 Then
        Dim cante As Byte
        Dim cante2 As Byte
        
        For cante = 1 To 4
            For cante2 = 1 To 4
                MapData(tX + (cante - 1), tY + (cante2 - 1)).Blocked = 1
                MapData(tX + (cante - 1), tY + (cante2 - 1)).Graphic(1).GrhIndex = 1
                MapData(tX + (cante - 1), tY + (cante2 - 1)).Graphic(2).GrhIndex = 0
                MapData(tX + (cante - 1), tY + (cante2 - 1)).Graphic(3).GrhIndex = 0
                MapData(tX + (cante - 1), tY + (cante2 - 1)).Graphic(4).GrhIndex = 0
            Next
        Next
    End If
    
    '************** Place exit
  
    If frmHerramientas.PlaceExitCmd.Enabled = False Then
        If frmHerramientas.EraseExitChk.value = 0 Then
            If frmHerramientas.Adya = vbChecked Then
                MapData(tX, tY).TileExit.Map = Val(frmHerramientas.MapExitTxt.Text)
                If tX = 92 Then
                          MapData(tX, tY).TileExit.x = 10
                          MapData(tX, tY).TileExit.y = tY
                ElseIf tX = 9 Then
                    MapData(tX, tY).TileExit.x = 91
                    MapData(tX, tY).TileExit.y = tY
                End If
                
                If tY = 94 Then
                         MapData(tX, tY).TileExit.y = 8
                         MapData(tX, tY).TileExit.x = tX
                ElseIf tY = 7 Then
                    MapData(tX, tY).TileExit.y = 93
                    MapData(tX, tY).TileExit.x = tX
                End If
                        
            Else
                MapData(tX, tY).TileExit.Map = Val(frmHerramientas.MapExitTxt.Text)
                MapData(tX, tY).TileExit.x = Val(frmHerramientas.XExitTxt.Text)
                MapData(tX, tY).TileExit.y = Val(frmHerramientas.YExitTxt.Text)
            End If
        Else
            MapData(tX, tY).TileExit.Map = 0
            MapData(tX, tY).TileExit.x = 0
            MapData(tX, tY).TileExit.y = 0
        End If
    End If

    '************** Place NPC
    If frmHerramientas.PlaceNPCCmd.Enabled = False Then
        If frmHerramientas.EraseNPCChk.value = 0 Then
            If frmHerramientas.NPCLst.ListIndex >= 0 Then
                NPCIndex = frmHerramientas.NPCLst.ListIndex + 1
                Body = Val(GetVar(IniDats & "NPCs.dat", "NPC" & NPCIndex, "Body"))
                Head = Val(GetVar(IniDats & "NPCs.dat", "NPC" & NPCIndex, "Head"))
                Heading = Val(GetVar(IniDats & "NPCs.dat", "NPC" & NPCIndex, "Heading"))
                Call MakeChar(NextOpenChar(), Body, Head, Heading, tX, tY)
                MapData(tX, tY).NPCIndex = NPCIndex
            End If
            
        Else
            If MapData(tX, tY).NPCIndex > 0 Then
                MapData(tX, tY).NPCIndex = 0
                Call EraseChar(MapData(tX, tY).CharIndex)
            End If
        End If
    End If
    
    If frmHerramientas.PlaceNPCHOSTCmd.Enabled = False Then
        If frmHerramientas.EraseNPCHOSTChk.value = 0 Then
            If frmHerramientas.NPCHOSTLst.ListIndex >= 0 Then
                NPCIndex = frmHerramientas.NPCHOSTLst.ListIndex + 1 + 499
                Body = Val(GetVar(IniDats & "NPCs-HOSTILES.dat", "NPC" & NPCIndex, "Body"))
                Head = Val(GetVar(IniDats & "NPCs-HOSTILES.dat", "NPC" & NPCIndex, "Head"))
                Heading = Val(GetVar(IniDats & "NPCs-HOSTILES.dat", "NPC" & NPCIndex, "Heading"))
                Call MakeChar(NextOpenChar(), Body, Head, Heading, tX, tY)
                MapData(tX, tY).NPCIndex = NPCIndex
            End If
         
            
        Else
            If MapData(tX, tY).NPCIndex > 0 Then
                MapData(tX, tY).NPCIndex = 0
                Call EraseChar(MapData(tX, tY).CharIndex)
            End If
        End If
    End If
    
    '************** Place OBJ
    If frmHerramientas.PlaceObjCmd.Enabled = False Then
        MapData(tX, tY).Blocked = frmMain.Blockedchk.value
        If frmHerramientas.EraseObjChk.value = 0 Then
            If frmHerramientas.ObjLst.ListIndex >= 0 Then
                objindex = frmHerramientas.ObjLst.ListIndex + 1
                InitGrh MapData(tX, tY).ObjGrh, Val(GetVar(IniDats & "OBJ.dat", "OBJ" & objindex, "GrhIndex"))
                MapData(tX, tY).OBJInfo.objindex = objindex
                MapData(tX, tY).OBJInfo.Amount = Val(frmHerramientas.OBJAmountTxt)
            End If
        Else
            MapData(tX, tY).OBJInfo.objindex = 0
            MapData(tX, tY).OBJInfo.Amount = 0
            MapData(tX, tY).ObjGrh.GrhIndex = 0
        End If
    End If
    
    If frmHerramientas.CmdTrigger.Enabled = False Then
    'Trigguer 4x4
    If frmHerramientas.trig4x4.value = 1 Then
        Dim vals1 As Byte
        Dim vals2 As Byte
        
        For vals1 = 1 To 4
            For vals2 = 1 To 4
                MapData(tX + (vals1 - 1), tY + (vals2 - 1)).Trigger = frmHerramientas.triggerlist.ListIndex
            Next
        Next
    End If
    Else
            MapData(tX, tY).Trigger = frmHerramientas.triggerlist.ListIndex
    End If
    
    'Set changed flag
    MapInfo.Changed = 1
End If
Exit Sub
ErrHandler:
Resume Next
End Sub

Function FileExist(File As String, FileType As VbFileAttribute) As Boolean
'*****************************************************************
'Checks to see if a file exists
'*****************************************************************

If Dir(File, FileType) = "" Then
    FileExist = False
Else
    FileExist = True
End If

End Function

Public Function ReadField(pos As Integer, Text As String, SepASCII As Integer) As String
'*****************************************************************
'Gets a field from a string
'*****************************************************************
Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String

Seperator = Chr(SepASCII)
LastPos = 0
FieldNum = 0

For i = 1 To Len(Text)
    CurChar = Mid(Text, i, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = pos Then
            ReadField = Mid(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next i
FieldNum = FieldNum + 1

If FieldNum = pos Then
    ReadField = Mid(Text, LastPos + 1)
End If

End Function

Private Sub AbrirBD()
Err.Clear
On Error GoTo fin

With Conexion
    .Provider = "Microsoft.Jet.OLEDB.3.51"
    .ConnectionString = "Data Source=" & IniBase & "grhindex.mdb"
    .Open
End With

Exit Sub



fin:
If Err Then
    Debug.Print Err.Description
    MsgBox "No se puede abrir la base de datos"
    End
End If

End Sub
Private Sub CargarIndices()
Dim rs As New Recordset
rs.Open "GrhIndex", Conexion, , adLockReadOnly, adCmdTable
Do While Not rs.EOF
    frmMain.lCelda.AddItem _
    rs!Nombre
    rs.MoveNext
Loop
rs.Close
End Sub
Sub rutas()
IniPath = App.Path & "\Init\"
IniDats = App.Path & "\Dats\"
IniBase = App.Path & "\BaseDatos\"
IniMaps = App.Path & "\IniMaps\"
End Sub

Sub Main()

Call rutas
Call IniciarCabecera(MiCabecera)

NumMidi = GetVar(IniPath & "grh.ini", "INIT", "NumMidi")
GrhPath = GetVar(IniPath & "grh.ini", "INIT", "Path")
Dim i
For i = 1 To NumMidi
    frmMusica.List1.AddItem "Mus" & i & ".mid"
Next

IniciarDirectSound
frmCargando.Show
frmCargando.Picture1.Picture = LoadPicture(App.Path & GrhPath & "\logo.jpg")
AbrirBD
CargarIndices
frmMain.Dialog.InitDir = IniMaps

'*****************************************************************
'Main
'*****************************************************************
Dim LoopC As Integer

'***************************************************
'Start up
'***************************************************

'****** INIT vars ******
ENDL = Chr(13) & Chr(10)
'Start up engine
frmCargando.Visible = False
'Display form handle, View window offset from 0,0 of display form, Tile Size, Display size in tiles, Screen buffer
If InitTileEngine(frmMain.hWnd, frmMain.MainViewShp.Top + 40, frmMain.MainViewShp.Left + 4, 32, 32, 13, 17, 12) Then

    '****** Load files into memory ******
    Call LoadBodyData
    Call LoadHeadData
    Call LoadNPCData
    Call LoadOBJData
    Call LoadOBJData2
    Call LoadTriggers

End If

CargarMIDI App.Path & MidiDir & "Mus1.mid"
'****** Show frmmain ******
frmMain.Show



'***************************************************
'Main Loop
'***************************************************
prgRun = True
Do While prgRun

    'Show Next Frame
    If frmMain.WindowState = 2 And MapaCargado Then
        ShowNextFrame frmMain.Top, frmMain.Left
    End If

    Call CheckKeys
   
    
    '****** Draw currently selected Grh in ShowPic ******
    If CurrentGrh.GrhIndex = 0 Then
        InitGrh CurrentGrh, 1
    End If
    
    Call MostrarGrh
    
    
    '****** Go do other events ******
    DoEvents
'    If Play Then
'        If Not EstaSonandoVieja Then
'            If frmMusica.Check1 = vbUnchecked Then
'                    Play_Midi
'            ElseIf frmMusica.Check1 = vbChecked Then
'                Dim N As Integer
'                N = RandomNumber(1, frmMusica.List1.ListCount)
'                stopmidi
'                CargarMIDI App.Path & MidiDir & "Mus" & N & ".mid"
'                CurMidi = "Mus" & N & ".mid"
'                frmMusica.MIdiAct.Caption = CurMidi
'                Play_Midi
'            End If
'
'        End If
'    End If
Loop
    

'*****************************************************************
'Close Down
'*****************************************************************

'****** Check if map is saved ******
If MapInfo.Changed = 1 Then
    If MsgBox("Este mapa há sido modificado. Vas a perder todos los cambios si no lo grabas. Lo queres grabar ahora?", vbYesNo) = vbYes Then
        Call SaveMapData(frmMain.Dialog.FileName)
    End If
End If

'Unload engine
DeInitTileEngine
LiberarDirectSound
'****** Unload forms and end******
Dim f
For Each f In Forms
    Unload f
Next
End

End Sub



Function GetVar(File As String, Main As String, Var As String) As String
'*****************************************************************
'Get a var to from a text file
'*****************************************************************
Dim l As Integer
Dim Char As String
Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

szReturn = ""

sSpaces = Space(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish


getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File

GetVar = RTrim(sSpaces)
GetVar = Left(GetVar, Len(GetVar) - 1)

End Function

Sub WriteVar(File As String, Main As String, Var As String, value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************

writeprivateprofilestring Main, Var, value, File

End Sub

Sub ToggleWalkMode()
'*****************************************************************
'Toggle walk mode on or off
'*****************************************************************
On Error GoTo fin:
If WalkMode = False Then
    WalkMode = True
Else
    WalkMode = False
End If

If WalkMode = False Then
    'Erase character
    Call EraseChar(UserCharIndex)
    MapData(UserPos.x, UserPos.y).CharIndex = 0
Else
    'MakeCharacter
    If LegalPos(UserPos.x, UserPos.y) Then
        Call MakeChar(NextOpenChar(), 1, 1, SOUTH, UserPos.x, UserPos.y)
        UserCharIndex = MapData(UserPos.x, UserPos.y).CharIndex
    Else
        MsgBox "Error: ubicacion ilegal."
        'frmMain.WalkModeChk.value = 0
    End If
End If
fin:
End Sub



Sub FixCoasts(ByVal GrhIndex As Integer, ByVal x As Integer, ByVal y As Integer)

If GrhIndex = 7284 Or GrhIndex = 7290 Or GrhIndex = 7291 Or GrhIndex = 7297 Or _
   GrhIndex = 7300 Or GrhIndex = 7301 Or GrhIndex = 7302 Or GrhIndex = 7303 Or _
   GrhIndex = 7304 Or GrhIndex = 7306 Or GrhIndex = 7308 Or GrhIndex = 7310 Or _
   GrhIndex = 7311 Or GrhIndex = 7313 Or GrhIndex = 7314 Or GrhIndex = 7315 Or _
   GrhIndex = 7316 Or GrhIndex = 7317 Or GrhIndex = 7319 Or GrhIndex = 7321 Or _
   GrhIndex = 7325 Or GrhIndex = 7326 Or GrhIndex = 7327 Or GrhIndex = 7328 Or GrhIndex = 7332 Or _
   GrhIndex = 7338 Or GrhIndex = 7339 Or GrhIndex = 7345 Or GrhIndex = 7348 Or _
   GrhIndex = 7349 Or GrhIndex = 7350 Or GrhIndex = 7351 Or GrhIndex = 7352 Or _
   GrhIndex = 7349 Or GrhIndex = 7350 Or GrhIndex = 7351 Or _
   GrhIndex = 7354 Or GrhIndex = 7357 Or GrhIndex = 7358 Or GrhIndex = 7360 Or _
   GrhIndex = 7362 Or GrhIndex = 7363 Or GrhIndex = 7365 Or GrhIndex = 7366 Or _
   GrhIndex = 7367 Or GrhIndex = 7368 Or GrhIndex = 7369 Or GrhIndex = 7371 Or _
   GrhIndex = 7373 Or GrhIndex = 7375 Or GrhIndex = 7376 Then MapData(x, y).Graphic(2).GrhIndex = 0

End Sub

Sub SaveMapData(SaveAs As String)
Dim LoopC As Integer
Dim TempInt As Integer
Dim y As Integer
Dim x As Integer

On Error Resume Next

If SaveAs = "" Then

        Exit Sub

End If
If FileExist(SaveAs, vbNormal) = True Then
    If MsgBox("¿Desea sobrescribir" & SaveAs & ".x archivos?", vbYesNo) = vbNo Then
        Exit Sub
    End If
End If

'Change mouse icon
frmMain.MousePointer = 11

If FileExist(SaveAs, vbNormal) = True Then
    Kill SaveAs
End If

If FileExist(Left(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) = True Then
    Kill Left(SaveAs, Len(SaveAs) - 4) & ".inf"
End If

'Open .map file
Open SaveAs For Binary As #1
Seek #1, 1


SaveAs = Left(SaveAs, Len(SaveAs) - 4)
SaveAs = SaveAs & ".inf"
'Open .inf file
Open SaveAs For Binary As #2
Seek #2, 1
'map Header
If frmMain.Vers.Text = "" Then
        Put #1, , 0
Else
        Put #1, , CInt(frmMain.Vers.Text)
End If

Put #1, , MiCabecera

Put #1, , TempInt
Put #1, , TempInt
Put #1, , TempInt
Put #1, , TempInt

'inf Header
Put #2, , TempInt
Put #2, , TempInt
Put #2, , TempInt
Put #2, , TempInt
Put #2, , TempInt

'Write .map file
For y = YMinMapSize To YMaxMapSize
    For x = XMinMapSize To XMaxMapSize
        
        '.map file
        Put #1, , MapData(x, y).Blocked
        For LoopC = 1 To 4
            If LoopC = 2 Then Call FixCoasts(MapData(x, y).Graphic(LoopC).GrhIndex, x, y)
            Put #1, , MapData(x, y).Graphic(LoopC).GrhIndex
        Next LoopC
        
        Put #1, , MapData(x, y).Trigger
        
        Put #1, , TempInt
        
        '.inf file
        'Tile exit
        Put #2, , MapData(x, y).TileExit.Map
        Put #2, , MapData(x, y).TileExit.x
        Put #2, , MapData(x, y).TileExit.y
        
        'NPC
        Put #2, , MapData(x, y).NPCIndex
        
        'Object
        Put #2, , MapData(x, y).OBJInfo.objindex
        Put #2, , MapData(x, y).OBJInfo.Amount
        
        'Empty place holders for future expansion
        Put #2, , TempInt
        Put #2, , TempInt
        
    Next x
Next y

'Close .map file
Close #1

'Close .inf file
Close #2

'write .dat file
'SaveAs = Left(SaveAs, Len(SaveAs) - 4) & ".dat"
SaveAs = Right(SaveAs, Len(SaveAs) - Len(IniMaps) + 1)

SaveAs = Left(SaveAs, Len(SaveAs) - 4) & ".dat"

Call WriteVar(IniMaps & SaveAs, Left(SaveAs, Len(SaveAs) - 4), "Name", MapInfo.Name)
Call WriteVar(IniMaps & SaveAs, Left(SaveAs, Len(SaveAs) - 4), "MusicNum", frmMain.Text2.Text)
Call WriteVar(IniMaps & SaveAs, Left(SaveAs, Len(SaveAs) - 4), "StartPos", MapInfo.StartPos.Map & "-" & MapInfo.StartPos.x & "-" & MapInfo.StartPos.y)


If frmCarac.Option1(0).value Then
    Call WriteVar(IniMaps & SaveAs, Left(SaveAs, Len(SaveAs) - 4), "Terreno", frmCarac.Option1(0).Caption)
ElseIf frmCarac.Option1(1).value Then
    Call WriteVar(IniMaps & SaveAs, Left(SaveAs, Len(SaveAs) - 4), "Terreno", frmCarac.Option1(1).Caption)
ElseIf frmCarac.Option1(2).value Then
    Call WriteVar(IniMaps & SaveAs, Left(SaveAs, Len(SaveAs) - 4), "Terreno", frmCarac.Option1(2).Caption)
End If

If frmCarac.Option2(0).value Then
ElseIf frmCarac.Option2(1).value Then
    Call WriteVar(IniMaps & SaveAs, Left(SaveAs, Len(SaveAs) - 4), "Zona", frmCarac.Option2(1).Caption)
ElseIf frmCarac.Option2(2).value Then
    Call WriteVar(IniMaps & SaveAs, Left(SaveAs, Len(SaveAs) - 4), "Zona", frmCarac.Option2(2).Caption)
End If

If frmCarac.Check1 = vbChecked Then
        Call WriteVar(IniMaps & SaveAs, Left(SaveAs, Len(SaveAs) - 4), "Restringir", "Si")
Else
        Call WriteVar(IniMaps & SaveAs, Left(SaveAs, Len(SaveAs) - 4), "Restringir", "No")
End If

Call WriteVar(IniMaps & SaveAs, Left(SaveAs, Len(SaveAs) - 4), "MusicNum", frmMain.Text2.Text)
Call WriteVar(IniMaps & SaveAs, Left(SaveAs, Len(SaveAs) - 4), "StartPos", MapInfo.StartPos.Map & "-" & MapInfo.StartPos.x & "-" & MapInfo.StartPos.y)

'Change mouse icon
frmMain.MousePointer = 0

MsgBox ("Mapa guardado como " & Left(SaveAs, Len(SaveAs) - 4) & ".map")

End Sub



Sub LoadOBJData()
'*****************************************************************
'Setup OBJ list
'*****************************************************************
Dim NumOBJs As Integer
Dim Obj As Integer

'Get Number of Maps
NumOBJs = Val(GetVar(IniDats & "OBJ.dat", "INIT", "NumOBJs"))

'Add OBJs to the OBJ list
For Obj = 1 To NumOBJs
     frmHerramientas.ObjLst.AddItem Val(Obj) & " (" & GetVar(IniDats & "OBJ.dat", "OBJ" & Obj, "Name") & ")"
Next Obj

End Sub

Sub LoadTriggers()

Dim Numt As Integer
Dim t As Integer


Numt = Val(GetVar(IniDats & "Triggers.dat", "INIT", "NumTriggers"))

'Add OBJs to the OBJ list
For t = 1 To Numt
     frmHerramientas.triggerlist.AddItem GetVar(IniDats & "Triggers.dat", "Trig" & t, "Name")
Next t

End Sub

Sub LoadNPCData()

Dim NumNPCs As Integer
Dim NumNPCsHOST As Integer
Dim NPC As Integer

NumNPCs = Val(GetVar(IniDats & "NPCs.dat", "INIT", "NumNPCs"))
For NPC = 1 To NumNPCs
    frmHerramientas.NPCLst.AddItem GetVar(IniDats & "NPCs.dat", "NPC" & NPC, "Name")
Next NPC

NumNPCsHOST = Val(GetVar(IniDats & "NPCs-HOSTILES.dat", "INIT", "NumNPCs"))
For NPC = 1 To NumNPCsHOST
    frmHerramientas.NPCHOSTLst.AddItem GetVar(IniDats & "NPCs-HOSTILES.dat", "NPC" & NPC + 499, "Name")
Next NPC

End Sub

Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound

End Function

Private Sub MostrarGrh()
frmGrafico.ShowPic = frmGrafico.Picture1
    If frmMain.MOSAICO = vbUnchecked Then
        Call DrawGrhtoHdc(frmGrafico.ShowPic.hdc, CurrentGrh, 0, 0, 0, 0, SRCCOPY)
    Else
        Dim x As Integer, y As Integer, j As Integer, i As Integer
        Dim cont As Integer
        For i = 1 To CInt(Val(frmMain.mLargo))
            For j = 1 To CInt(Val(frmMain.mAncho))
                Call DrawGrhtoHdc(frmGrafico.ShowPic.hdc, CurrentGrh, (j - 1) * 32, (i - 1) * 32, 0, 0, SRCCOPY)
                If cont < CInt(Val(frmMain.mLargo)) * CInt(Val(frmMain.mAncho)) Then
                    cont = cont + 1
                    CurrentGrh.GrhIndex = CurrentGrh.GrhIndex + 1
                End If
            Next
        Next
        CurrentGrh.GrhIndex = CurrentGrh.GrhIndex - cont
    End If
     
frmGrafico.ShowPic.Picture = frmGrafico.ShowPic.Image
frmMain.Picture3 = frmGrafico.ShowPic

    

End Sub
