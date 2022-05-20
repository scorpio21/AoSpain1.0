VERSION 5.00
Begin VB.Form frmGrilla 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grilla"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   345
      Left            =   270
      TabIndex        =   10
      Top             =   1980
      Width           =   3960
   End
   Begin VB.TextBox Col 
      Height          =   285
      Index           =   2
      Left            =   3255
      TabIndex        =   8
      Text            =   "255"
      Top             =   1305
      Width           =   465
   End
   Begin VB.TextBox Col 
      Height          =   285
      Index           =   1
      Left            =   2190
      TabIndex        =   6
      Text            =   "255"
      Top             =   1305
      Width           =   465
   End
   Begin VB.TextBox Col 
      Height          =   285
      Index           =   0
      Left            =   1125
      TabIndex        =   4
      Text            =   "255"
      Top             =   1305
      Width           =   465
   End
   Begin VB.TextBox Alto 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Text            =   "128"
      Top             =   330
      Width           =   465
   End
   Begin VB.TextBox Ancho 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1410
      TabIndex        =   1
      Text            =   "128"
      Top             =   315
      Width           =   495
   End
   Begin VB.Shape Shape2 
      Height          =   645
      Left            =   285
      Top             =   1125
      Width           =   3990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Azul"
      Height          =   195
      Index           =   2
      Left            =   2880
      TabIndex        =   9
      Top             =   1335
      Width           =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Verde"
      Height          =   195
      Index           =   1
      Left            =   1755
      TabIndex        =   7
      Top             =   1335
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Rojo"
      Height          =   195
      Index           =   0
      Left            =   750
      TabIndex        =   5
      Top             =   1335
      Width           =   330
   End
   Begin VB.Shape Shape1 
      Height          =   780
      Left            =   255
      Top             =   120
      Width           =   4020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Alto"
      Height          =   195
      Left            =   2250
      TabIndex        =   2
      Top             =   375
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ancho"
      Height          =   195
      Left            =   900
      TabIndex        =   0
      Top             =   375
      Width           =   465
   End
End
Attribute VB_Name = "frmGrilla"
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

Private Sub Form_Deactivate()
Me.Visible = False
End Sub

