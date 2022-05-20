VERSION 5.00
Begin VB.Form frmMusica 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MUSICA ;-)"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      BackColor       =   &H8000000D&
      Caption         =   "Shuffle"
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   1485
      TabIndex        =   6
      Top             =   3330
      Width           =   1320
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stop"
      Height          =   330
      Left            =   1395
      TabIndex        =   4
      Top             =   4815
      Width           =   915
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Hide me"
      Height          =   330
      Left            =   90
      TabIndex        =   3
      Top             =   4815
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   330
      Left            =   2655
      TabIndex        =   2
      Top             =   4815
      Width           =   1005
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   855
      TabIndex        =   0
      Top             =   630
      Width           =   2085
   End
   Begin VB.Label MIdiAct 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1530
      TabIndex        =   5
      Top             =   4410
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Canciones:"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1350
      TabIndex        =   1
      Top             =   270
      Width           =   795
   End
End
Attribute VB_Name = "frmMusica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Private Sub Command1_Click()
Play = True
End Sub

Private Sub Command2_Click()
Command2.Enabled = False
Command3.Enabled = True
End Sub

Private Sub Command3_Click()
Command2.Enabled = True
Command3.Enabled = False
End Sub

Private Sub Command4_Click()
Me.Visible = False
End Sub

Private Sub Command5_Click()
If EstaSonandoVieja Then
    Stop_Midi
End If
Play = False
End Sub

Private Sub Command6_Click()
Command2.Enabled = True
Command3.Enabled = True
End Sub

Private Sub List1_Click()
'Command5_Click
CurMidi = List1.List(List1.ListIndex)
MIdiAct.Caption = CurMidi
CargarMIDI App.Path & MidiDir & List1.List(List1.ListIndex)
'Play = True
Play_Midi
Play = True
End Sub

Private Sub Timer1_Timer()

End Sub
