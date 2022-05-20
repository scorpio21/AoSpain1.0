VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10635
   ClientLeft      =   390
   ClientTop       =   960
   ClientWidth     =   14040
   Icon            =   "frmMain2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   709
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   936
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Blo4 
      BackColor       =   &H00808000&
      Caption         =   "Bloq 4x4"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   240
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   5715
      Width           =   1740
   End
   Begin VB.CheckBox ByV 
      BackColor       =   &H00808000&
      Caption         =   "Blq y Vacio 4x4"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   225
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1740
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5355
      Left            =   15
      ScaleHeight     =   355
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   295
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   6750
      Width           =   4455
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   3690
         Left            =   -15
         ScaleHeight     =   3630
         ScaleWidth      =   4350
         TabIndex        =   36
         Top             =   0
         Width           =   4410
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   2565
      Top             =   2340
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Caption         =   "&H00000000&"
      Height          =   2115
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   13920
      Begin VB.CommandButton Command2 
         Caption         =   "H&erramientas"
         Height          =   375
         Left            =   4230
         TabIndex        =   59
         Top             =   1440
         Width           =   1500
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Bl&oq map"
         Height          =   375
         Index           =   0
         Left            =   1110
         TabIndex        =   58
         Top             =   1440
         Width           =   1500
      End
      Begin VB.CommandButton Command5 
         Caption         =   "D&esbloq map"
         Height          =   375
         Index           =   0
         Left            =   2670
         TabIndex        =   57
         Top             =   1440
         Width           =   1500
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00808000&
         Caption         =   "Cambiar vista pero se delira mal"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4215
         TabIndex        =   54
         Top             =   1875
         Width           =   3240
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
         Left            =   8805
         TabIndex        =   52
         Text            =   "1"
         Top             =   1470
         Width           =   795
      End
      Begin VB.CheckBox EM 
         BackColor       =   &H00808000&
         Caption         =   "Telep Mode"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   7515
         TabIndex        =   51
         Top             =   1710
         Width           =   1335
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
         Left            =   8820
         TabIndex        =   50
         Text            =   "1"
         Top             =   765
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
         Left            =   8820
         TabIndex        =   49
         Text            =   "1"
         Top             =   150
         Width           =   795
      End
      Begin VB.CommandButton MasY 
         Caption         =   "+"
         Height          =   300
         Index           =   1
         Left            =   7635
         TabIndex        =   46
         Top             =   915
         Width           =   375
      End
      Begin VB.CommandButton MenosY 
         Caption         =   "-"
         Height          =   300
         Index           =   0
         Left            =   8085
         TabIndex        =   45
         Top             =   915
         Width           =   375
      End
      Begin VB.CommandButton MenosX 
         Caption         =   "-"
         Height          =   300
         Index           =   1
         Left            =   8085
         TabIndex        =   44
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton MasX 
         Caption         =   "+"
         Height          =   300
         Index           =   0
         Left            =   7605
         TabIndex        =   43
         Top             =   300
         Width           =   375
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00808000&
         Caption         =   "Loop"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   345
         Left            =   90
         TabIndex        =   41
         Top             =   1770
         Width           =   720
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   90
         TabIndex        =   39
         Text            =   "1"
         Top             =   1335
         Width           =   750
      End
      Begin VB.TextBox Vers 
         Height          =   330
         Left            =   4305
         TabIndex        =   37
         Text            =   "         "
         Top             =   840
         Width           =   1485
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000007&
         Height          =   1680
         Left            =   5895
         ScaleHeight     =   108
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   99
         TabIndex        =   24
         Top             =   30
         Width           =   1545
         Begin VB.Label Apuntador 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "O"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   30
            TabIndex        =   25
            Top             =   0
            Width           =   120
         End
      End
      Begin VB.CheckBox Completar 
         BackColor       =   &H00808000&
         Caption         =   "AutoCompletar"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   45
         TabIndex        =   21
         Top             =   690
         Width           =   1665
      End
      Begin VB.CommandButton Command6 
         Caption         =   "+"
         Height          =   240
         Index           =   3
         Left            =   3960
         TabIndex        =   20
         Top             =   870
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "-"
         Height          =   240
         Index           =   2
         Left            =   2865
         TabIndex        =   19
         Top             =   870
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "-"
         Height          =   240
         Index           =   1
         Left            =   2865
         TabIndex        =   18
         Top             =   465
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "+"
         Height          =   240
         Index           =   0
         Left            =   3960
         TabIndex        =   17
         Top             =   465
         Width           =   240
      End
      Begin VB.TextBox DMLargo 
         Height          =   330
         Left            =   3135
         TabIndex        =   16
         Text            =   "0"
         Top             =   825
         Width           =   780
      End
      Begin VB.TextBox DMAncho 
         Height          =   330
         Left            =   3135
         TabIndex        =   15
         Text            =   "0"
         Top             =   420
         Width           =   780
      End
      Begin VB.CheckBox DespMosaic 
         BackColor       =   &H00808000&
         Caption         =   "DespMosaico"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   2775
         TabIndex        =   14
         Top             =   105
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   4305
         TabIndex        =   13
         Top             =   315
         Width           =   1485
      End
      Begin VB.TextBox mLargo 
         Height          =   330
         Left            =   1470
         TabIndex        =   11
         Text            =   "1"
         Top             =   300
         Width           =   1140
      End
      Begin VB.TextBox mAncho 
         Height          =   330
         Left            =   105
         TabIndex        =   10
         Text            =   "1"
         Top             =   300
         Width           =   1140
      End
      Begin VB.CheckBox MOSAICO 
         BackColor       =   &H00808000&
         Caption         =   "Mosaico"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   1695
         TabIndex        =   9
         Top             =   705
         Width           =   1065
      End
      Begin VB.TextBox StatTxt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   1995
         Left            =   9765
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "frmMain2.frx":030A
         Top             =   30
         Width           =   1965
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Mapa"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   3
         Left            =   7875
         TabIndex        =   53
         Top             =   1455
         Width           =   465
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00000000&
         Height          =   645
         Index           =   3
         Left            =   7485
         Top             =   1365
         Width           =   2205
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Salidas Y"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   2
         Left            =   7680
         TabIndex        =   48
         Top             =   705
         Width           =   855
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00000000&
         Height          =   645
         Index           =   2
         Left            =   7485
         Top             =   690
         Width           =   2205
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Salidas X"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   1
         Left            =   7680
         TabIndex        =   47
         Top             =   45
         Width           =   855
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00000000&
         Height          =   660
         Index           =   1
         Left            =   7485
         Top             =   45
         Width           =   2205
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MIDI"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   225
         TabIndex        =   40
         Top             =   1050
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   4740
         TabIndex        =   38
         Top             =   630
         Width           =   690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Alto"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   1875
         TabIndex        =   23
         Top             =   60
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Ancho"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   413
         TabIndex        =   22
         Top             =   60
         Width           =   555
      End
      Begin VB.Shape Shape4 
         Height          =   1170
         Index           =   0
         Left            =   4260
         Top             =   45
         Width           =   1620
      End
      Begin VB.Shape Shape3 
         Height          =   1170
         Left            =   2760
         Top             =   45
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del mapa"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   0
         Left            =   4305
         TabIndex        =   12
         Top             =   90
         Width           =   1560
      End
      Begin VB.Shape Shape6 
         Height          =   930
         Left            =   30
         Top             =   45
         Width           =   2745
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4620
      Left            =   30
      TabIndex        =   2
      Top             =   2130
      Width           =   4395
      Begin VB.CheckBox Check3 
         BackColor       =   &H00808000&
         Caption         =   "Ver triggers"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   195
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   2865
         Width           =   1815
      End
      Begin VB.CheckBox Erasechk 
         BackColor       =   &H00808000&
         Caption         =   "Borrar Layer"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2400
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   3525
         Width           =   1425
      End
      Begin VB.CheckBox EraseAllchk 
         BackColor       =   &H00808000&
         Caption         =   "Borrar todo"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2400
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3870
         Width           =   1335
      End
      Begin VB.TextBox Layertxt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   480
         Left            =   2322
         TabIndex        =   30
         Text            =   "1"
         Top             =   2985
         Width           =   800
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Down"
         Height          =   255
         Left            =   3210
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Up"
         Height          =   255
         Left            =   3210
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Grhtxt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   480
         Left            =   2322
         TabIndex        =   27
         Text            =   "1"
         Top             =   2280
         Width           =   800
      End
      Begin VB.CommandButton PlaceGrhCmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Poner Grh"
         Height          =   255
         Left            =   2310
         TabIndex        =   26
         Top             =   4140
         Width           =   1515
      End
      Begin VB.ListBox lCelda 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1620
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   0
         Width           =   4065
      End
      Begin VB.CheckBox Mostar4layer 
         BackColor       =   &H00808000&
         Caption         =   "4º layer"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   345
         Left            =   945
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2550
         Width           =   1185
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00808000&
         Caption         =   "Mostrar Blocked"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   420
         Left            =   150
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2175
         Width           =   2010
      End
      Begin VB.CheckBox DrawGridChk 
         BackColor       =   &H00808000&
         Caption         =   "Grilla"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   945
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1860
         Width           =   915
      End
      Begin VB.CommandButton PlaceBlockCmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cambiar bloqueado"
         Height          =   255
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4140
         Width           =   1515
      End
      Begin VB.CheckBox Blockedchk 
         BackColor       =   &H00808000&
         Caption         =   "Bloqueado"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   210
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   3330
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "Grh"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   2565
         TabIndex        =   34
         Top             =   2025
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "Layer"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   2475
         TabIndex        =   33
         Top             =   2775
         Width           =   495
      End
      Begin VB.Shape Shape2 
         Height          =   2535
         Left            =   2175
         Top             =   1950
         Width           =   1785
      End
      Begin VB.Shape Shape1 
         Height          =   1200
         Left            =   75
         Top             =   3285
         Width           =   2115
      End
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H000000C0&
      Height          =   6240
      Left            =   4455
      Top             =   2175
      Width           =   8160
   End
   Begin VB.Menu FileMnu 
      Caption         =   "Archivo"
      Begin VB.Menu mnunuevo 
         Caption         =   "Nuevo mapa"
      End
      Begin VB.Menu mnunuevo3 
         Caption         =   "Rellenar en 4x4"
      End
      Begin VB.Menu mnuNuevo2 
         Caption         =   "Rellenar en 2x2"
      End
      Begin VB.Menu mnuCargar 
         Caption         =   "Cargar Mapa"
      End
      Begin VB.Menu SaveMnu 
         Caption         =   "Grabar"
      End
      Begin VB.Menu SaveNewMnu 
         Caption         =   "Grabar como mapa nuevo"
      End
      Begin VB.Menu nAbout 
         Caption         =   "Acerca de"
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu OptionMnu 
      Caption         =   "Mapa"
      Begin VB.Menu borratri 
         Caption         =   "Borrar Triggers"
      End
      Begin VB.Menu Actmapas 
         Caption         =   "Actualizar mapas"
      End
      Begin VB.Menu mnuCarac 
         Caption         =   "Caracteristicas"
      End
      Begin VB.Menu ClsRoomMnu 
         Caption         =   "Borrar Mapa"
      End
      Begin VB.Menu ClsBordMnu 
         Caption         =   "Borrar Borde"
      End
      Begin VB.Menu mnuGrilla 
         Caption         =   "Grilla"
      End
      Begin VB.Menu mnuMusica 
         Caption         =   "Musica"
      End
      Begin VB.Menu mnuborrarArboles 
         Caption         =   "Borrar Arboles"
      End
      Begin VB.Menu mnuborrarNpcs 
         Caption         =   "Borrar Npcs"
      End
      Begin VB.Menu mnuBloquear 
         Caption         =   "Bloquear Bordes"
      End
      Begin VB.Menu mnuExits 
         Caption         =   "Poner Exits"
      End
   End
End
Attribute VB_Name = "frmMain"
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


Option Explicit



Function DameGrhIndex(Nombre As String) As Integer
Dim rs As New Recordset
rs.Open "GrhIndex", Conexion, , adLockReadOnly, adCmdTable

Do While Not rs.EOF And rs!Nombre <> Nombre
    rs.MoveNext
Loop
DameGrhIndex = rs!GrhIndice
If rs!Ancho > 0 Then
    MOSAICO.value = vbChecked
    mAncho.Text = rs!Ancho
    mLargo.Text = rs!Alto
Else
    MOSAICO.value = vbUnchecked
    mAncho.Text = ""
    mLargo.Text = ""
End If
        
rs.Close

End Function


Private Sub Actmapas_Click()
Dim i As Integer
For i = 2 To 52
    Call SwitchMap("map" & i & ".map")
    Call SaveMapData("Map" & i & ".map")
Next
End Sub

Private Sub Blockedchk_Click()

Call PlaceBlockCmd_Click

End Sub

Private Sub ObtenerNombreArchivo(Guardar As Boolean)
With Dialog
    .Filter = "Mapas|*.map"
    
    If Guardar Then
            .DialogTitle = "Guardar"
            .DefaultExt = ".txt"
            .FileName = ""
            .flags = cdlOFNPathMustExist
            .ShowSave
           
    Else
        .DialogTitle = "Cargar"
        '.FileName = ""
        
        .flags = cdlOFNFileMustExist
        .ShowOpen
    End If
End With

End Sub

Private Sub borratri_Click()
Dim y As Integer
Dim x As Integer
For y = YMinMapSize To YMaxMapSize
    For x = XMinMapSize To XMaxMapSize
        MapData(x, y).Trigger = 0
    Next x
Next y

End Sub

Private Sub Check1_Click()

If DrawBlock = True Then
    DrawBlock = False
Else
    DrawBlock = True
End If

End Sub

Private Sub Check4_Click()
If Check4.value = 1 Then
    'Call DeInitTileEngine
    '****** Clear DirectX objects ******

    Call InitTileEngine(frmMain.hWnd, frmMain.MainViewShp.Top + 40, frmMain.MainViewShp.Left + 4, 32, 32, 18, 22, 12, True)
Else
    'Call DeInitTileEngine
    '****** Clear DirectX objects ******

    Call InitTileEngine(frmMain.hWnd, frmMain.MainViewShp.Top + 40, frmMain.MainViewShp.Left + 4, 32, 32, 13, 17, 12, True)
End If
End Sub

Private Sub Command4_Click(Index As Integer)
If MsgBox("Cuidado, con este comando podes arruinar el mapa.¿Estas seguro que queres hacer esto?", vbYesNo) = vbNo Then
        Exit Sub
End If
Dim y As Integer
Dim x As Integer

If Not MapaCargado Then
    Exit Sub
End If

For y = YMinMapSize To YMaxMapSize
    For x = XMinMapSize To XMaxMapSize
        MapData(x, y).Blocked = 1
    Next x
Next y

MapInfo.Changed = 1
End Sub

Private Sub Command5_Click(Index As Integer)
If MsgBox("Cuidado, con este comando podes arruinar el mapa.¿Estas seguro que queres hacer esto?", vbYesNo) = vbNo Then
        Exit Sub
End If
Dim y As Integer
Dim x As Integer

If Not MapaCargado Then
    Exit Sub
End If

For y = YMinMapSize To YMaxMapSize
    For x = XMinMapSize To XMaxMapSize
        MapData(x, y).Blocked = 0
    Next x
Next y

MapInfo.Changed = 0
End Sub

Private Sub Command6_Click(Index As Integer)
On Error Resume Next
Select Case Index
        Case 0
            DMAncho.Text = Str(Val(DMAncho.Text) + 1)
        Case 1
            DMAncho.Text = Str(Val(DMAncho.Text) - 1)
        Case 2
            DMLargo.Text = Str(Val(DMLargo.Text) - 1)
        Case 3
            DMLargo.Text = Str(Val(DMLargo.Text) + 1)
End Select
End Sub


Private Sub DespMosaic_Click()
If DMAncho.Text = "" Then DMAncho.Text = "0"
If DMLargo.Text = "" Then DMLargo.Text = "0"
End Sub

Private Sub EM_Click()
If EM.value = 1 Then
    EY = 0
    EX = 0
    'EM.value = 0
'Else
    'EM.value = 1
End If
End Sub

Private Sub lCelda_Click()
Grhtxt.Text = DameGrhIndex(lCelda.List(lCelda.ListIndex))
'If frmGrafico.Visible = False Then frmGrafico.Visible = True
Call PlaceGrhCmd_Click
End Sub

Private Sub lCelda_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
Grhtxt.SetFocus
End Sub

Private Sub lCelda_KeyPress(KeyAscii As Integer)
KeyAscii = 0
Grhtxt.SetFocus
End Sub

Private Sub lCelda_KeyUp(KeyCode As Integer, Shift As Integer)
KeyCode = 0
Grhtxt.SetFocus
End Sub

Private Sub MapExitTxt_Change()
frmHerramientas.MapExitTxt.Text = MapExitTxt.Text
End Sub

Private Sub MasX_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim CoorX As Integer

CoorX = Val(frmHerramientas.XExitTxt.Text)

'Sumo en uno el valor X
If Button = vbLeftButton Then
    CoorX = CoorX + 1
End If

If Button = vbRightButton Then
    CoorX = CoorX + 10
End If

frmHerramientas.XExitTxt.Text = CoorX
CoorX = Val(frmHerramientas.XExitTxt.Text)
XExitTxt = frmHerramientas.XExitTxt.Text
End Sub

Private Sub MenosX_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim CoorX As Integer

CoorX = Val(frmHerramientas.XExitTxt.Text)

'resto en uno el valor X
If Button = vbLeftButton Then
    CoorX = CoorX - 1
End If

If Button = vbRightButton Then
    CoorX = CoorX - 10
End If

frmHerramientas.XExitTxt.Text = CoorX
CoorX = Val(frmHerramientas.XExitTxt.Text)
XExitTxt = frmHerramientas.XExitTxt.Text
End Sub

Private Sub MasY_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim CoorY As Integer

CoorY = Val(frmHerramientas.YExitTxt.Text)

'Sumo en uno el valor X
If Button = vbLeftButton Then
    CoorY = CoorY + 1
End If

If Button = vbRightButton Then
    CoorY = CoorY + 10
End If

frmHerramientas.YExitTxt.Text = CoorY
CoorY = Val(frmHerramientas.YExitTxt.Text)
YExitTxt = frmHerramientas.YExitTxt.Text
End Sub

Private Sub MenosY_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim CoorY As Integer

CoorY = Val(frmHerramientas.YExitTxt.Text)

'Sumo en uno el valor X
If Button = vbLeftButton Then
    CoorY = CoorY - 1
End If

If Button = vbRightButton Then
    CoorY = CoorY - 10
End If

frmHerramientas.YExitTxt.Text = CoorY
CoorY = Val(frmHerramientas.YExitTxt.Text)

YExitTxt = frmHerramientas.YExitTxt.Text
End Sub

Private Sub mnuBloquear_Click()

Dim y As Integer
Dim x As Integer

If Not MapaCargado Then
    Exit Sub
End If

For y = YMinMapSize To YMaxMapSize
    For x = XMinMapSize To XMaxMapSize

        If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
            MapData(x, y).Blocked = 1
        End If
    Next x
Next y

'Set changed flag
MapInfo.Changed = 1

End Sub

Private Sub mnuborrarArboles_Click()
Dim y As Integer
Dim x As Integer

If Not MapaCargado Then
    Exit Sub
End If

For y = 1 To 100
    For x = 1 To 100
    
        If MapData(x, y).OBJInfo.objindex > 0 Then
            
            Dim c As String
            c = GetVar(IniDats & "OBJ.dat", "OBJ" & MapData(x, y).OBJInfo.objindex, "Objtype")
            If c = "" Then c = 0
            If Val(c) = 4 Then
                    MapData(x, y).OBJInfo.objindex = 0
                    MapData(x, y).Blocked = 0
                    MapData(x, y).OBJInfo.Amount = 0
                End If
            End If
'                If MapData(x, y).OBJInfo.objindex = 4 Or _
'                   MapData(x, y).OBJInfo.objindex = 5 Or _
'                   MapData(x, y).OBJInfo.objindex = 6 Then
'
'                    MapData(x, y).Blocked = 0
'                    MapData(x, y).OBJInfo.objindex = 0
'                    MapData(x, y).OBJInfo.Amount = 0
'                    MapData(x, y).ObjGrh.Grhindex = 0
'                    MapData(x, y).ObjGrh.FrameCounter = 0
'                    MapData(x, y).ObjGrh.SpeedCounter = 0
'                    MapData(x, y).ObjGrh.Started = 0
'
'
'
'                End If
        
    Next
Next

End Sub
Private Sub mnuborrarNpcs_Click()
Dim y As Integer
Dim x As Integer

If Not MapaCargado Then
    Exit Sub
End If

For y = 1 To 100
    For x = 1 To 100
    
        If MapData(x, y).OBJInfo.objindex > 0 Then
            
            Dim c As String
            c = GetVar(IniDats & "NPCs-HOSTILES.dat", "NPC" & MapData(x, y).NPCIndex, "Objtype")
            If c = "" Then c = 0
            If Val(c) = 4 Then
                    MapData(x, y).OBJInfo.objindex = 0
                    MapData(x, y).Blocked = 0
                    MapData(x, y).OBJInfo.Amount = 0
                End If
            End If
            
            
            If MapData(x, y).NPCIndex > 0 Then
            EraseChar MapData(x, y).CharIndex
            MapData(x, y).NPCIndex = 0
        End If
        
    Next
Next

End Sub

Private Sub mnuCarac_Click()
frmCarac.Visible = True
End Sub

Private Sub mnuCargar_Click()
'frmCargar.Visible = True
Dialog.CancelError = True
On Error GoTo ErrHandler
Call ObtenerNombreArchivo(False)


If MapInfo.Changed = 1 Then
    If MsgBox("Este mapa há sido modificado. Vas a perder todos los cambios si no lo grabas. Lo queres grabar ahora?", vbYesNo) = vbYes Then
        Call SaveMapData(Dialog.FileName)
    End If
End If


    UserPos.x = (WindowTileWidth \ 2) + 1
    
    UserPos.y = (WindowTileHeight \ 2) + 1
    
    Call mnunuevo_Click
    Call SwitchMap(Dialog.FileName)
    EngineRun = True
Exit Sub

ErrHandler:
MsgBox Err.Description
End Sub

Private Sub mnuExits_Click()
Exits.Show
End Sub

Private Sub mnuGrilla_Click()
frmGrilla.Visible = True
End Sub

Private Sub mnuMusica_Click()
frmMusica.Show
End Sub

Private Sub mnunuevo_Click()


Dim y As Integer
Dim x As Integer

Call borratri_Click

frmMain.MousePointer = 11
For y = YMinMapSize To YMaxMapSize
    For x = XMinMapSize To XMaxMapSize
        MapData(x, y).Graphic(1).GrhIndex = 1
        'Change blockes status
        MapData(x, y).Blocked = frmMain.Blockedchk.value

        'Erase layer 2 and 3
        MapData(x, y).Graphic(2).GrhIndex = 0
        MapData(x, y).Graphic(3).GrhIndex = 0
        MapData(x, y).Graphic(4).GrhIndex = 0

        'Erase NPCs
        If MapData(x, y).NPCIndex > 0 Then
            EraseChar MapData(x, y).CharIndex
            MapData(x, y).NPCIndex = 0
        End If

        'Erase Objs
        MapData(x, y).OBJInfo.objindex = 0
        MapData(x, y).OBJInfo.Amount = 0
        MapData(x, y).ObjGrh.GrhIndex = 0

        'Clear exits
        MapData(x, y).TileExit.Map = 0
        MapData(x, y).TileExit.x = 0
        MapData(x, y).TileExit.y = 0
         
         
        MapData(x, y).Blocked = frmMain.Blockedchk.value
        MapData(x, y).Graphic(Val(frmMain.Layertxt.Text)).GrhIndex = Val(frmMain.Grhtxt.Text)
            
        InitGrh MapData(x, y).Graphic(Val(frmMain.Layertxt.Text)), Val(frmMain.Grhtxt.Text)
        
    Next x
Next y


MapInfo.Changed = 1
MapInfo.MapVersion = 0

Text1.Text = "Nuevo Mapa"
UserPos.x = (WindowTileWidth \ 2) + 1

UserPos.y = (WindowTileHeight \ 2) + 1

'CurMap = frmCargar.MapLst.ListCount
MapaCargado = True
EngineRun = True
frmMain.MousePointer = 0
End Sub

Private Sub mnunuevo3_Click()


Dim y As Integer
Dim x As Integer

Call borratri_Click

frmMain.MousePointer = 11
For y = YMinMapSize To YMaxMapSize Step 4
    For x = XMinMapSize To XMaxMapSize Step 4
        MapData(x, y).Graphic(1).GrhIndex = 3
        'Change blockes status
        MapData(x, y).Blocked = frmMain.Blockedchk.value

        'Erase layer 2 and 3
        'If y - 1 Mod 4 = 0 And x - 1 Mod 4 = 0 Or x = 1 Or y = 1 Then
        Dim tXX As Integer, tYY As Integer, i As Integer, j As Integer, desptile As Integer
        Dim aux As Integer
            tXX = x
            tYY = y
            desptile = 0
         
            For i = 1 To 4
                For j = 1 To 4
                    aux = Val(frmMain.Grhtxt.Text) + desptile
                     
                     MapData(tXX, tYY).Blocked = frmMain.Blockedchk.value
                     'Exit Sub
                     MapData(tXX, tYY).Graphic(1).GrhIndex = aux
                     
                     InitGrh MapData(tXX, tYY).Graphic(1), aux
                      If tXX < 100 Then tXX = tXX + 1
                     desptile = desptile + 1
                Next
                tXX = x
                If tYY < 100 Then tYY = tYY + 1
            Next
       ' End If
        
         
         
        MapData(x, y).Blocked = frmMain.Blockedchk.value
        'MapData(x, y).Graphic(Val(frmMain.Layertxt.Text)).GrhIndex = Val(frmMain.Grhtxt.Text)
            
        'InitGrh MapData(x, y).Graphic(Val(frmMain.Layertxt.Text)), Val(frmMain.Grhtxt.Text)
        
    Next x
Next y


MapInfo.Changed = 1
MapInfo.MapVersion = 0

Text1.Text = "Nuevo Mapa"
UserPos.x = (WindowTileWidth \ 2) + 1

UserPos.y = (WindowTileHeight \ 2) + 1

'CurMap = frmCargar.MapLst.ListCount
MapaCargado = True
EngineRun = True
frmMain.MousePointer = 0
End Sub
Private Sub mnuNuevo2_Click()


Dim y As Integer
Dim x As Integer

Call borratri_Click

frmMain.MousePointer = 11
For y = YMinMapSize To YMaxMapSize
    For x = XMinMapSize To XMaxMapSize
        'MapData(x, y).Graphic(1).GrhIndex = 3
        'Change blockes status
        MapData(x, y).Blocked = frmMain.Blockedchk.value

        'Erase layer 2 and 3
        If y Mod 2 <> 0 And x Mod 2 <> 0 Then
        Dim tXX As Integer, tYY As Integer, i As Integer, j As Integer, desptile As Integer
        Dim aux As Integer
            tXX = x
            tYY = y
            desptile = 0
         
            For i = 1 To 2
                For j = 1 To 2
                    aux = Val(frmMain.Grhtxt.Text) + desptile
                     
                     MapData(tXX, tYY).Blocked = frmMain.Blockedchk.value
                     'Exit Sub
                     MapData(tXX, tYY).Graphic(1).GrhIndex = aux
                     
                     InitGrh MapData(tXX, tYY).Graphic(1), aux
                      If tXX < 100 Then tXX = tXX + 1
                     desptile = desptile + 1
                Next
                tXX = x
                If tYY < 100 Then tYY = tYY + 1
            Next
        End If
       
         
        MapData(x, y).Blocked = frmMain.Blockedchk.value
        'MapData(x, y).Graphic(Val(frmMain.Layertxt.Text)).GrhIndex = Val(frmMain.Grhtxt.Text)
            
        'InitGrh MapData(x, y).Graphic(Val(frmMain.Layertxt.Text)), Val(frmMain.Grhtxt.Text)
        
    Next x
Next y


MapInfo.Changed = 1
MapInfo.MapVersion = 0

Text1.Text = "Nuevo Mapa"
UserPos.x = (WindowTileWidth \ 2) + 1

UserPos.y = (WindowTileHeight \ 2) + 1

'CurMap = frmCargar.MapLst.ListCount
MapaCargado = True
EngineRun = True
frmMain.MousePointer = 0
End Sub

Private Sub MOSAICO_Click()
If mAncho.Text = "" Then mAncho.Text = "1"
If mLargo.Text = "" Then mLargo.Text = "1"
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Apuntador.Move x, y
UserPos.x = x
UserPos.y = y
Call ActualizaDespGrilla
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
MiRadarX = x
MiRadarY = y
End Sub

Public Sub PlaceBlockCmd_Click()

PlaceGrhCmd.Enabled = True
PlaceBlockCmd.Enabled = False
frmHerramientas.PlaceExitCmd.Enabled = True
frmHerramientas.PlaceNPCHOSTCmd.Enabled = True
frmHerramientas.PlaceNPCCmd.Enabled = True
frmHerramientas.PlaceObjCmd.Enabled = True

End Sub


Private Sub Grhtxt_Change()

If Val(Grhtxt.Text) < 1 Then
  Grhtxt.Text = NumGrhs
  Exit Sub
End If

If Val(Grhtxt.Text) > NumGrhs Then
  Grhtxt.Text = 1
  Exit Sub
End If

'Change CurrentGrh
CurrentGrh.GrhIndex = Val(Grhtxt.Text)
CurrentGrh.Started = 1
CurrentGrh.FrameCounter = 1
CurrentGrh.SpeedCounter = GrhData(CurrentGrh.GrhIndex).Speed

End Sub

Private Sub ClsBordMnu_Click()

'*****************************************************************
'Clears a border in a room with current GRH
'*****************************************************************

Dim y As Integer
Dim x As Integer

If Not MapaCargado Then
    Exit Sub
End If

For y = YMinMapSize To YMaxMapSize
    For x = XMinMapSize To XMaxMapSize

        If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then

          If frmMain.MOSAICO.value = vbChecked Then
            Dim aux As Integer
            aux = Val(frmMain.Grhtxt.Text) + _
            ((y Mod frmMain.mLargo) * frmMain.mAncho) + (x Mod frmMain.mAncho)
             MapData(x, y).Blocked = frmMain.Blockedchk.value
             MapData(x, y).Graphic(Val(frmMain.Layertxt.Text)).GrhIndex = aux
            'Setup GRH
            InitGrh MapData(x, y).Graphic(Val(frmMain.Layertxt.Text)), aux
          Else
            'Else Place graphic
            MapData(x, y).Blocked = frmMain.Blockedchk.value
            MapData(x, y).Graphic(Val(frmMain.Layertxt.Text)).GrhIndex = Val(frmMain.Grhtxt.Text)
            
            'Setup GRH
    
            InitGrh MapData(x, y).Graphic(Val(frmMain.Layertxt.Text)), Val(frmMain.Grhtxt.Text)
        End If
             'Erase NPCs
            If MapData(x, y).NPCIndex > 0 Then
                EraseChar MapData(x, y).CharIndex
                MapData(x, y).NPCIndex = 0
            End If

            'Erase Objs
            MapData(x, y).OBJInfo.objindex = 0
            MapData(x, y).OBJInfo.Amount = 0
            MapData(x, y).ObjGrh.GrhIndex = 0

            'Clear exits
            MapData(x, y).TileExit.Map = 0
            MapData(x, y).TileExit.x = 0
            MapData(x, y).TileExit.y = 0

        End If

    Next x
Next y

'Set changed flag
MapInfo.Changed = 1

End Sub

Private Sub ClsRoomMnu_Click()
'*****************************************************************
'Clears all layers
'*****************************************************************

Dim y As Integer
Dim x As Integer

If Not MapaCargado Then
    Exit Sub
End If

For y = YMinMapSize To YMaxMapSize
    For x = XMinMapSize To XMaxMapSize
        MapData(x, y).Graphic(1).GrhIndex = 3
        'Change blockes status
        MapData(x, y).Blocked = frmMain.Blockedchk.value

        'Erase layer 2 and 3
        MapData(x, y).Graphic(2).GrhIndex = 0
        MapData(x, y).Graphic(3).GrhIndex = 0
        MapData(x, y).Graphic(4).GrhIndex = 0

        'Erase NPCs
        If MapData(x, y).NPCIndex > 0 Then
            EraseChar MapData(x, y).CharIndex
            MapData(x, y).NPCIndex = 0
        End If

        'Erase Objs
        MapData(x, y).OBJInfo.objindex = 0
        MapData(x, y).OBJInfo.Amount = 0
        MapData(x, y).ObjGrh.GrhIndex = 0

        'Clear exits
        MapData(x, y).TileExit.Map = 0
        MapData(x, y).TileExit.x = 0
        MapData(x, y).TileExit.y = 0

        If frmMain.MOSAICO.value = vbChecked Then
            Dim aux As Integer
            aux = Val(frmMain.Grhtxt.Text) + _
            ((y Mod frmMain.mLargo) * frmMain.mAncho) + (x Mod frmMain.mAncho)
             MapData(x, y).Blocked = frmMain.Blockedchk.value
             MapData(x, y).Graphic(Val(frmMain.Layertxt.Text)).GrhIndex = aux
            'Setup GRH
            InitGrh MapData(x, y).Graphic(Val(frmMain.Layertxt.Text)), aux
        Else
            'Else Place graphic
            MapData(x, y).Blocked = frmMain.Blockedchk.value
            MapData(x, y).Graphic(Val(frmMain.Layertxt.Text)).GrhIndex = Val(frmMain.Grhtxt.Text)
            
            'Setup GRH
    
            InitGrh MapData(x, y).Graphic(Val(frmMain.Layertxt.Text)), Val(frmMain.Grhtxt.Text)
        End If

    Next x
Next y

'Set changed flag
MapInfo.Changed = 1

End Sub





Private Sub Command2_Click()
If frmHerramientas.Visible Then frmHerramientas.SetFocus _
Else: frmHerramientas.Visible = Not frmHerramientas.Visible
End Sub

Private Sub DrawGridChk_Click()

If DrawGrid = True Then
    DrawGrid = False
Else
    DrawGrid = True
End If

End Sub

Private Sub EraseAllchk_Click()
Call PlaceGrhCmd_Click
Erasechk.value = False
End Sub

Private Sub Erasechk_Click()

'Set Place GRh mode
Call PlaceGrhCmd_Click

EraseAllchk.value = False

End Sub

Private Sub EraseExitChk_Click()

Call frmHerramientas.PlaceExitCmd_Click

End Sub

Private Sub EraseNPCChk_Click()

Call frmHerramientas.PlaceNPCCmd_Click

End Sub

Private Sub EraseObjChk_Click()

Call frmHerramientas.PlaceObjCmd_Click

End Sub



Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim Grh As Integer

'Set Place GRh mode
Call PlaceGrhCmd_Click

Grh = Val(Grhtxt.Text)

'Add to current Grh number
If Button = vbLeftButton Then
   Grh = Grh - 1
End If

If Button = vbRightButton Then
    Grh = Grh - 10
End If

'Update Grhtxt
Grhtxt.Text = Grh
Grh = Val(Grhtxt)

'If blank find next valid Grh
If GrhData(Grh).NumFrames = 0 Then
    
    Do Until GrhData(Grh).NumFrames > 0
        Grh = Grh - 1
        If Grh < 1 Then
            Grh = NumGrhs
        End If
    Loop
    
End If

'Update Grhtxt
Grhtxt.Text = Grh

End Sub

Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim Grh As Integer

'Set Place GRh mode
Call PlaceGrhCmd_Click

Grh = Val(Grhtxt.Text)

'Add to current Grh number
If Button = vbLeftButton Then
   Grh = Grh + 1
End If

If Button = vbRightButton Then
    Grh = Grh + 10
End If

'Update Grhtxt
Grhtxt.Text = Grh
Grh = Val(Grhtxt)

'If blank find next valid Grh
If GrhData(Grh).NumFrames = 0 Then
    
    Do Until GrhData(Grh).NumFrames > 0
        Grh = Grh + 1
        If Grh > NumGrhs Then
            Grh = 1
        End If
    Loop
    
End If

'Update Grhtxt
Grhtxt.Text = Grh

End Sub

Private Sub Layertxt_Change()

If Val(Layertxt.Text) < 1 Then
  Layertxt.Text = 1
End If

If Val(Layertxt.Text) > 4 Then
  Layertxt.Text = 4
End If

Call PlaceGrhCmd_Click

End Sub




Private Sub Form_Load()
frmMain.Caption = frmMain.Caption & " V " & App.Major & "." & App.Minor & "." & App.Revision

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim tX As Integer
Dim tY As Integer

If Not MapaCargado Then Exit Sub

If x <= MainViewShp.Left Or x >= MainViewShp.Left + MainViewWidth Or y <= MainViewShp.Top Or y >= MainViewShp.Top + MainViewHeight Then
    Exit Sub
End If

ConvertCPtoTP MainViewShp.Left, MainViewShp.Top, x, y, tX, tY

ReacttoMouseClick Button, tX, tY

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim tX As Integer
Dim tY As Integer

'Make sure map is loaded
If Not MapaCargado Then Exit Sub

'Make sure click is in view window
If x <= MainViewShp.Left Or x >= MainViewShp.Left + MainViewWidth Or y <= MainViewShp.Top Or y >= MainViewShp.Top + MainViewHeight Then
    Exit Sub
End If

ConvertCPtoTP MainViewShp.Left, MainViewShp.Top, x, y, tX, tY

ReacttoMouseClick Button, tX, tY

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'Allow MainLoop to close program
If prgRun = True Then
    prgRun = False
    Cancel = 1
End If

End Sub

Private Sub nAbout_Click()
frmAbout1.Show
End Sub

Public Sub PlaceGrhCmd_Click()
PlaceGrhCmd.Enabled = False
PlaceBlockCmd.Enabled = True
frmHerramientas.PlaceExitCmd.Enabled = True
frmHerramientas.PlaceNPCCmd.Enabled = True
frmHerramientas.PlaceObjCmd.Enabled = True

End Sub

Private Sub Salir_Click()
Unload Me
End Sub

Private Sub SaveMnu_Click()



Call SaveMapData(Dialog.FileName)

'Set changed flag
MapInfo.Changed = 0

End Sub


Private Sub SaveNewMnu_Click()
Dialog.CancelError = True
On Error GoTo ErrHandler
Call ObtenerNombreArchivo(True)
Call SaveMapData(Dialog.FileName)
    'frmCargar.MapLst.AddItem "Map " & NumMaps, NumMaps - 1
Exit Sub

ErrHandler:
MsgBox Err.Description

End Sub


Private Sub Text1_Change()
MapInfo.Name = Text1.Text
End Sub

Private Sub WalkModeChk_Click()

'ToggleWalkMode

End Sub

Private Sub VScroll1_Scroll()

End Sub

