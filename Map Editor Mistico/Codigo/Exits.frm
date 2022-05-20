VERSION 5.00
Begin VB.Form Exits 
   BackColor       =   &H00808000&
   Caption         =   "Exits"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CANCELAR 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3495
      TabIndex        =   10
      Top             =   2475
      Width           =   1065
   End
   Begin VB.CommandButton ACEPTAR 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   2130
      TabIndex        =   9
      Top             =   2475
      Width           =   1155
   End
   Begin VB.TextBox Oeste 
      Height          =   450
      Left            =   450
      TabIndex        =   6
      Top             =   1080
      Width           =   900
   End
   Begin VB.TextBox Este 
      Height          =   450
      Left            =   3060
      TabIndex        =   5
      Top             =   1080
      Width           =   900
   End
   Begin VB.TextBox Sur 
      Height          =   450
      Left            =   1785
      TabIndex        =   1
      Top             =   1770
      Width           =   900
   End
   Begin VB.TextBox Norte 
      Height          =   450
      Left            =   1785
      TabIndex        =   0
      Top             =   255
      Width           =   900
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808000&
      Caption         =   "Este"
      Height          =   255
      Left            =   3338
      TabIndex        =   8
      Top             =   810
      Width           =   345
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "Oeste"
      Height          =   255
      Left            =   660
      TabIndex        =   7
      Top             =   810
      Width           =   480
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808000&
      Caption         =   "Norte"
      Height          =   255
      Left            =   2018
      TabIndex        =   4
      Top             =   30
      Width           =   435
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "Sur"
      Height          =   255
      Left            =   2130
      TabIndex        =   3
      Top             =   1515
      Width           =   315
   End
   Begin VB.Label Map 
      BackStyle       =   0  'Transparent
      Caption         =   "MAP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   1845
      TabIndex        =   2
      Top             =   1005
      Width           =   825
   End
End
Attribute VB_Name = "Exits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ACEPTAR_Click()
Dim cantidad As Integer

If Norte.Text <> "" Then
    For cantidad = 1 To 100
        MapData(cantidad, 7).TileExit.Map = Val(Norte.Text)
        MapData(cantidad, 7).TileExit.x = cantidad
        MapData(cantidad, 7).TileExit.y = 93
    Next
End If
If Sur.Text <> "" Then
    For cantidad = 1 To 100
        MapData(cantidad, 94).TileExit.Map = Val(Sur.Text)
        MapData(cantidad, 94).TileExit.x = cantidad
        MapData(cantidad, 94).TileExit.y = 8
    Next
End If
If Este.Text <> "" Then
    For cantidad = 1 To 100
        MapData(92, cantidad).TileExit.Map = Val(Este.Text)
        MapData(92, cantidad).TileExit.x = 10
        MapData(92, cantidad).TileExit.y = cantidad
    Next
End If
If Oeste.Text <> "" Then
    For cantidad = 1 To 100
        MapData(9, cantidad).TileExit.Map = Val(Oeste.Text)
        MapData(9, cantidad).TileExit.x = 91
        MapData(9, cantidad).TileExit.y = cantidad
Next
End If
Unload Me
End Sub

Private Sub CANCELAR_Click()
Unload Me
End Sub

Private Sub Form_Load()
Map.Caption = MapInfo.Name
Norte.Text = "1"
Sur.Text = "1"
Este.Text = "1"
Oeste.Text = "1"

End Sub

