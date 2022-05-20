Attribute VB_Name = "MoD_MIDI"
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

Public Const MidiDir = "\MIDI\"


Public CurMidi As String
Public LoopMidi As Byte
Public IsPlayingCheck As Boolean

Public GetStartTime As Long
Public Offset As Long
Public mtTime As Long
Public mtLength As Double
Public dTempo As Double


Public Perf As DirectMusicPerformance
Public Seg As DirectMusicSegment
Public SegState As DirectMusicSegmentState
Public Loader As DirectMusicLoader

Dim timesig As DMUS_TIMESIGNATURE
Dim portcaps As DMUS_PORTCAPS

Dim msg As String
Dim time As Double
Dim Offset2 As Long
Dim ElapsedTime2 As Double
Dim fIsPaused As Boolean


Public Sub CargarMIDI(Archivo As String)
'If Musica = 1 Then Exit Sub
On Error GoTo fin
    
    If Loader Is Nothing Then Set Loader = DirectX.DirectMusicLoaderCreate()
    Set Seg = Loader.LoadSegment(Archivo)
    Set Loader = Nothing 'Liberamos el cargador
    Exit Sub
fin:
    MsgBox ("Error producido en 'Public Sub CargarMIDI' " & Err.Description & " " & Err.Number & " " & Archivo)

End Sub

Public Sub Stop_Midi()

If IsPlayingCheck Then
     Offset = 0
     If Perf.IsPlaying(Seg, SegState) = True Then
            Call Perf.Stop(Seg, SegState, 0, 0)
     End If
     'LoopMidi = False
     IsPlayingCheck = False
     Seg.SetStartPoint (0)
     'Seg.Unload
     Call Perf.Reset(0)
End If
End Sub

Public Sub Play_Midi()
On Error GoTo fin
        If IsPlayingCheck Then Stop_Midi
        Offset = 0
        If Perf.IsPlaying(Seg, SegState) = True Then
            Call Perf.Stop(Seg, SegState, 0, 0)
        End If
        Seg.SetStartPoint (0)
        Set SegState = Perf.PlaySegment(Seg, 0, 0)
        IsPlayingCheck = True
        Exit Sub
fin:
    MsgBox "Error producido en Public Sub Play_Midi()"

End Sub

Function Sonando()
Sonando = (Perf.IsPlaying(Seg, SegState) = True)
End Function




