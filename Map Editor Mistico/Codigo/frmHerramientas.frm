VERSION 5.00
Begin VB.Form frmHerramientas 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Herramientas"
   ClientHeight    =   7005
   ClientLeft      =   8145
   ClientTop       =   1785
   ClientWidth     =   6075
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   6075
   Begin VB.CheckBox trig4x4 
      BackColor       =   &H8000000D&
      Caption         =   "trig4x4"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3945
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5152
      Width           =   1635
   End
   Begin VB.CommandButton CmdTrigger 
      Caption         =   "PonerTrigger"
      Height          =   285
      Left            =   3990
      TabIndex        =   30
      Top             =   5715
      Width           =   1515
   End
   Begin VB.ListBox triggerlist 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1620
      Left            =   3945
      TabIndex        =   29
      Top             =   3510
      Width           =   1545
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2820
      Left            =   3780
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      TabStop         =   0   'False
      Text            =   "frmHerramientas.frx":0000
      Top             =   465
      Width           =   2130
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3015
      TabIndex        =   26
      Text            =   "1"
      Top             =   6090
      Width           =   555
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1050
      TabIndex        =   24
      Text            =   "1"
      Top             =   6135
      Width           =   555
   End
   Begin VB.CheckBox Adya 
      BackColor       =   &H8000000D&
      Caption         =   "Mapa Adyacente"
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   210
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2475
      Width           =   1530
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Agregar al azar"
      Height          =   285
      Left            =   2115
      TabIndex        =   22
      Top             =   2835
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Agregar al azar"
      Height          =   285
      Left            =   2160
      TabIndex        =   21
      Top             =   5715
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar al azar"
      Height          =   285
      Left            =   180
      TabIndex        =   20
      Top             =   5715
      Width           =   1500
   End
   Begin VB.CheckBox EraseNPCChk 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "Borrar NPC"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   210
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5175
      Width           =   1215
   End
   Begin VB.CommandButton PlaceNPCCmd 
      Caption         =   "Poner NPC"
      Height          =   255
      Left            =   180
      TabIndex        =   18
      Top             =   5400
      Width           =   1515
   End
   Begin VB.ListBox NPCLst 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1620
      Left            =   180
      TabIndex        =   17
      Top             =   3510
      Width           =   1575
   End
   Begin VB.CheckBox EraseNPCHOSTChk 
      BackColor       =   &H8000000D&
      Caption         =   "Borrar NPC"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2145
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5175
      Width           =   1215
   End
   Begin VB.CommandButton PlaceNPCHOSTCmd 
      Caption         =   "Poner NPC"
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   5415
      Width           =   1515
   End
   Begin VB.ListBox NPCHOSTLst 
      Height          =   1620
      Left            =   2115
      TabIndex        =   14
      Top             =   3510
      Width           =   1545
   End
   Begin VB.CommandButton command3 
      Caption         =   "Ocultar"
      Height          =   285
      Left            =   300
      TabIndex        =   13
      Top             =   90
      Width           =   5325
   End
   Begin VB.CommandButton PlaceExitCmd 
      BackColor       =   &H8000000D&
      Caption         =   "Poner salida"
      Height          =   255
      Left            =   210
      TabIndex        =   8
      Top             =   2790
      Width           =   1515
   End
   Begin VB.CheckBox EraseExitChk 
      BackColor       =   &H8000000D&
      Caption         =   "Borrar salida"
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   225
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2190
      Width           =   1215
   End
   Begin VB.TextBox MapExitTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   945
      TabIndex        =   6
      Text            =   "1"
      Top             =   615
      Width           =   795
   End
   Begin VB.TextBox XExitTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   945
      TabIndex        =   5
      Text            =   "1"
      Top             =   1095
      Width           =   795
   End
   Begin VB.TextBox YExitTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   930
      TabIndex        =   4
      Text            =   "1"
      Top             =   1575
      Width           =   795
   End
   Begin VB.ListBox ObjLst 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1230
      Left            =   2040
      TabIndex        =   3
      Top             =   510
      Width           =   1575
   End
   Begin VB.CommandButton PlaceObjCmd 
      Caption         =   "Poner OBJ"
      Height          =   255
      Left            =   2115
      TabIndex        =   2
      Top             =   2520
      Width           =   1515
   End
   Begin VB.CheckBox EraseObjChk 
      BackColor       =   &H8000000D&
      Caption         =   "Borrar OBJ"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2115
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox OBJAmountTxt 
      Height          =   285
      Left            =   2940
      TabIndex        =   0
      Text            =   "1"
      Top             =   1770
      Width           =   555
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   2
      Height          =   2715
      Index           =   2
      Left            =   150
      Top             =   495
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2175
      TabIndex        =   27
      Top             =   6090
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   210
      TabIndex        =   25
      Top             =   6135
      Width           =   810
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   2
      Height          =   3135
      Index           =   0
      Left            =   135
      Top             =   3420
      Width           =   1695
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   2
      Height          =   3135
      Index           =   1
      Left            =   2055
      Top             =   3420
      Width           =   1695
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   405
      TabIndex        =   12
      Top             =   1140
      Width           =   225
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   405
      TabIndex        =   11
      Top             =   1620
      Width           =   225
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "MAPA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   165
      TabIndex        =   10
      Top             =   675
      Width           =   735
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000D&
      BorderWidth     =   2
      Height          =   3165
      Left            =   3810
      Top             =   3405
      Width           =   1890
   End
   Begin VB.Shape Shape5 
      BorderWidth     =   2
      Height          =   2715
      Left            =   2010
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2100
      TabIndex        =   9
      Top             =   1770
      Width           =   810
   End
End
Attribute VB_Name = "frmHerramientas"
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



Private Sub CmdTrigger_Click()
frmMain.PlaceGrhCmd.Enabled = True
frmMain.PlaceBlockCmd.Enabled = True
PlaceExitCmd.Enabled = True
PlaceNPCHOSTCmd.Enabled = True
PlaceNPCCmd.Enabled = True
PlaceObjCmd.Enabled = True

CmdTrigger.Enabled = False

End Sub

Private Sub Command1_Click()
Call PonerAlAzar(CInt(Text1.Text), 1)
End Sub

Private Sub Command2_Click()
Call PonerAlAzar(CInt(Text2.Text), 2)
End Sub

Private Sub Command3_Click()
Me.Visible = False
frmGrafico.Visible = False
End Sub


Private Sub PonerAlAzar(ByVal n As Integer, t As Byte)
Dim x, y, i
Dim Head As Integer
Dim Body As Integer
Dim Heading As Byte
i = n
Do While i > 0
    x = CInt(RandomNumber(XMinMapSize, XMaxMapSize - 1))
    y = CInt(RandomNumber(YMinMapSize, YMaxMapSize - 1))
    
    Select Case t
        Case 0
            If MapData(x, y).OBJInfo.objindex = 0 Then
                  i = i - 1
                  MapData(x, y).Blocked = frmMain.Blockedchk.value
                  If frmHerramientas.ObjLst.ListIndex >= 0 Then
                      objindex = frmHerramientas.ObjLst.ListIndex + 1
                      InitGrh MapData(x, y).ObjGrh, Val(GetVar(IniDats & "OBJ.dat", "OBJ" & objindex, "GrhIndex"))
                      MapData(x, y).OBJInfo.objindex = objindex
                      MapData(x, y).OBJInfo.Amount = Val(frmHerramientas.OBJAmountTxt)
                  End If
            End If
        Case 1
           If MapData(x, y).Blocked = 0 Then
                  i = i - 1
                  If frmHerramientas.NPCLst.ListIndex >= 0 Then
                        NPCIndex = frmHerramientas.NPCLst.ListIndex + 1
                        Body = Val(GetVar(IniDats & "NPCs.dat", "NPC" & NPCIndex, "Body"))
                        Head = Val(GetVar(IniDats & "NPCs.dat", "NPC" & NPCIndex, "Head"))
                        Heading = Val(GetVar(IniDats & "NPCs.dat", "NPC" & NPCIndex, "Heading"))
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, CInt(x), CInt(y))
                        MapData(x, y).NPCIndex = NPCIndex
                  End If
            End If
        Case 2
           If MapData(x, y).Blocked = 0 Then
                  i = i - 1
                  If frmHerramientas.NPCHOSTLst.ListIndex >= 0 Then
                        NPCIndex = frmHerramientas.NPCHOSTLst.ListIndex + 1 + 499
                        Body = Val(GetVar(IniDats & "NPCs-HOSTILES.dat", "NPC" & NPCIndex, "Body"))
                        Head = Val(GetVar(IniDats & "NPCs-HOSTILES.dat", "NPC" & NPCIndex, "Head"))
                        Heading = Val(GetVar(IniDats & "NPCs-HOSTILES.dat", "NPC" & NPCIndex, "Heading"))
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, CInt(x), CInt(y))
                        MapData(x, y).NPCIndex = NPCIndex
                  End If
           End If
        End Select
Loop
End Sub

Private Sub Command4_Click()
Call PonerAlAzar(CInt(OBJAmountTxt.Text), 0)
End Sub

Private Sub NPCHOSTLst_Click()
Call PlaceNPCHOSTCmd_Click
Text3.Text = GetVar(IniDats & "NPCs-HOSTILES.dat", "NPC" & (499 + NPCHOSTLst.ListIndex + 1), "INFO")
End Sub

Private Sub NPCHOSTLst_DblClick()
Call PlaceNPCHOSTCmd_Click
End Sub

Private Sub NPCLst_DblClick()
Call PlaceNPCCmd_Click
End Sub

Private Sub PlaceNPCHOSTCmd_Click()
frmMain.PlaceGrhCmd.Enabled = True
frmMain.PlaceBlockCmd.Enabled = True
PlaceExitCmd.Enabled = True
PlaceNPCHOSTCmd.Enabled = False
PlaceNPCCmd.Enabled = True
PlaceObjCmd.Enabled = True
CmdTrigger.Enabled = True
End Sub



Private Sub XExitTxt_Change()

If Val(XExitTxt.Text) < XMinMapSize Then
  XExitTxt.Text = XMinMapSize
End If

If Val(XExitTxt.Text) > XMaxMapSize Then
  XExitTxt.Text = XMaxMapSize
End If

Call PlaceExitCmd_Click

End Sub




Private Sub YExitTxt_Change()

If Val(YExitTxt.Text) < YMinMapSize Then
  YExitTxt.Text = YMinMapSize
End If

If Val(YExitTxt.Text) > YMaxMapSize Then
  YExitTxt.Text = YMaxMapSize
End If

Call PlaceExitCmd_Click

End Sub

Private Sub MapExitTxt_Change()
frmMain.MapExitTxt.Text = MapExitTxt.Text
If Val(MapExitTxt.Text) < 1 Then
  MapExitTxt.Text = 1
End If

If Val(MapExitTxt.Text) > NumMaps Then
  MapExitTxt.Text = NumMaps
End If

Call PlaceExitCmd_Click

End Sub




Private Sub NPCLst_Click()
Call PlaceNPCCmd_Click
Text3.Text = GetVar(IniDats & "NPCs.dat", "NPC" & NPCLst.ListIndex + 1, "INFO")
End Sub

Private Sub OBJAmountTxt_Change()

If Val(OBJAmountTxt.Text) > MAX_INVENORY_OBJS Then
    OBJAmountTxt.Text = 0
End If

If Val(OBJAmountTxt.Text) < 1 Then
    OBJAmountTxt.Text = MAX_INVENORY_OBJS
End If

End Sub

Private Sub ObjLst_Click()

Call PlaceObjCmd_Click
Text3.Text = ObjData(ObjLst.ListIndex + 1).info
End Sub



Public Sub PlaceExitCmd_Click()

frmMain.PlaceGrhCmd.Enabled = True
frmMain.PlaceBlockCmd.Enabled = True
PlaceExitCmd.Enabled = False
PlaceNPCHOSTCmd.Enabled = True
PlaceNPCCmd.Enabled = True
PlaceObjCmd.Enabled = True
CmdTrigger.Enabled = True

End Sub




Public Sub PlaceNPCCmd_Click()

frmMain.PlaceGrhCmd.Enabled = True
frmMain.PlaceBlockCmd.Enabled = True
PlaceExitCmd.Enabled = True
PlaceNPCCmd.Enabled = False
PlaceNPCHOSTCmd.Enabled = True
PlaceObjCmd.Enabled = True
CmdTrigger.Enabled = True
End Sub


Public Sub PlaceObjCmd_Click()

frmMain.PlaceGrhCmd.Enabled = True
frmMain.PlaceBlockCmd.Enabled = True
PlaceExitCmd.Enabled = True
PlaceNPCCmd.Enabled = True
PlaceNPCHOSTCmd.Enabled = True
PlaceObjCmd.Enabled = False
CmdTrigger.Enabled = True

End Sub






Private Sub Check2_Click()

End Sub






