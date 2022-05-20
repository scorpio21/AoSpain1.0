VERSION 5.00
Begin VB.Form frmCarac 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Caracteristicas"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   1680
      TabIndex        =   5
      Top             =   0
      Width           =   1575
      Begin VB.OptionButton Option2 
         Caption         =   "CIUDAD"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "DUNGEON"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "CAMPO"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Prohibir entrada"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "NIEVE"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "DESIERTO"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "BOSQUE"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Left            =   120
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   120
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmCarac"
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
Me.Visible = False
End Sub

