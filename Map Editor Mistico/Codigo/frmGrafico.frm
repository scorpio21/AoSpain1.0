VERSION 5.00
Begin VB.Form frmGrafico 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grafico"
   ClientHeight    =   5385
   ClientLeft      =   600
   ClientTop       =   7590
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5355
      Left            =   0
      ScaleHeight     =   355
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   455
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6855
      Begin VB.PictureBox ShowPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   5355
         Left            =   0
         ScaleHeight     =   355
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   455
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmGrafico"
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


Private Sub Form_Deactivate()
Me.Visible = False
End Sub

