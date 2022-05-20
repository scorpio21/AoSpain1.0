Attribute VB_Name = "DS"
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


Public Const NumSoundBuffers = 20

Public DirectSound As DirectSound
Public DSBuffers(1 To NumSoundBuffers) As DirectSoundBuffer
Public LastSoundBufferUsed As Integer


Public Sub IniciarDirectSound()
Err.Clear
On Error GoTo fin
    
    '<----------------Direct Music--------------->
    Set Perf = DirectX.DirectMusicPerformanceCreate()
    Call Perf.Init(Nothing, 0)
    Perf.SetPort -1, 80
    Call Perf.SetMasterAutoDownload(True)
    '<------------------------------------------->
    
    Set DirectSound = DirectX.DirectSoundCreate("")
    If Err Then
        MsgBox "Error iniciando DirectSound"
        End
    End If
    
    LastSoundBufferUsed = 1
    'CrearMIDILoader DX
    
    
    Exit Sub
fin:
End
End Sub

Public Sub LiberarDirectSound()
Dim cloop As Integer
For cloop = 1 To NumSoundBuffers
    Set DSBuffers(cloop) = Nothing
Next cloop
Set DirectSound = Nothing
End Sub

