VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Ver Graficos"
   ClientHeight    =   7830
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12525
   Icon            =   "Encode.frx":0000
   MouseIcon       =   "Encode.frx":10CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   7830
   ScaleWidth      =   12525
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraCarpetas 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "Carpetas"
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   2460
      TabIndex        =   54
      Top             =   5190
      Width           =   2550
      Begin VB.TextBox GrhTex 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   885
         Width           =   735
      End
      Begin VB.TextBox txtIni 
         Height          =   285
         Left            =   750
         TabIndex        =   56
         Text            =   "Ini"
         Top             =   270
         Width           =   1740
      End
      Begin VB.TextBox txtGraficos 
         Height          =   285
         Left            =   735
         TabIndex        =   55
         Text            =   "Graficos"
         Top             =   540
         Width           =   1755
      End
      Begin VB.Label lblGRH 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GRH:"
         Height          =   255
         Left            =   165
         TabIndex        =   64
         Top             =   915
         Width           =   495
      End
      Begin VB.Label lblCarpetaInds 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Inds:"
         Height          =   195
         Left            =   105
         TabIndex        =   58
         Top             =   285
         Width           =   345
      End
      Begin VB.Label lblCarpetaDe 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Graficos:"
         Height          =   195
         Left            =   90
         TabIndex        =   57
         Top             =   585
         Width           =   630
      End
   End
   Begin VB.Frame FraDatosIndice 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "Datos Indice Elegido"
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   1035
      TabIndex        =   47
      Top             =   5190
      Visible         =   0   'False
      Width           =   3270
      Begin VB.TextBox TxtReferencias 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   540
         Width           =   705
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   870
         TabIndex        =   59
         Top             =   960
         Visible         =   0   'False
         Width           =   2010
      End
      Begin VB.TextBox TxtGrhIndex 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   2250
         TabIndex        =   53
         Top             =   255
         Width           =   840
      End
      Begin VB.TextBox txtAncho 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   780
         TabIndex        =   52
         Top             =   495
         Width           =   630
      End
      Begin VB.TextBox txtAlto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   780
         TabIndex        =   51
         Top             =   225
         Width           =   630
      End
      Begin VB.Label lblReferencias 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referencias:"
         Height          =   195
         Left            =   1470
         TabIndex        =   61
         Top             =   555
         Width           =   900
      End
      Begin VB.Label lblDescripcion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   210
         TabIndex        =   60
         Top             =   990
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label lblGrhIndex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GrhIndex:"
         Height          =   195
         Left            =   1545
         TabIndex        =   50
         Top             =   255
         Width           =   690
      End
      Begin VB.Label lblAncho 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ancho:"
         Height          =   195
         Left            =   240
         TabIndex        =   49
         Top             =   540
         Width           =   510
      End
      Begin VB.Label lblAlto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alto:"
         Height          =   195
         Left            =   430
         TabIndex        =   48
         Top             =   225
         Width           =   315
      End
   End
   Begin ChamaleonButton.ChameleonBtn BotonI 
      Height          =   255
      Index           =   0
      Left            =   45
      TabIndex        =   34
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Graficos"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Encode.frx":1D94
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CommandButton CmbResourceFile 
      Caption         =   "Limpiar Memoria"
      Height          =   255
      Left            =   1080
      MouseIcon       =   "Encode.frx":1DB0
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CheckBox Checkcabeza 
      BackColor       =   &H00C0FFC0&
      Caption         =   "cabeza"
      Height          =   195
      Left            =   3840
      TabIndex        =   32
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "f. verde"
      Height          =   315
      Left            =   4335
      MouseIcon       =   "Encode.frx":2A7A
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   4680
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4590
      Top             =   6165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox CDibujarWalk 
      Height          =   315
      ItemData        =   "Encode.frx":3744
      Left            =   3840
      List            =   "Encode.frx":3757
      MouseIcon       =   "Encode.frx":377B
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton BotonBorrrar 
      Caption         =   "borrar"
      Height          =   255
      Left            =   4440
      MouseIcon       =   "Encode.frx":4445
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton NuevoGhr 
      Caption         =   "Nuevo/buscar"
      Height          =   255
      Left            =   960
      MouseIcon       =   "Encode.frx":510F
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Guardar Graficos.ind"
      Height          =   255
      Left            =   2610
      MouseIcon       =   "Encode.frx":5DD9
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   4380
      Width           =   2415
   End
   Begin VB.CommandButton BotonGuardar 
      Caption         =   "Guardar"
      Height          =   255
      Left            =   2640
      MouseIcon       =   "Encode.frx":6AA3
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Timer Dibujado 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   240
      Top             =   120
   End
   Begin VB.TextBox TextDatos 
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   3840
      TabIndex        =   11
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   3840
      TabIndex        =   10
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   7
      Left            =   3840
      TabIndex        =   9
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   6
      Left            =   3840
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   5
      Left            =   3840
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   4
      Left            =   3840
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   3
      Left            =   3840
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   1
      Left            =   3240
      TabIndex        =   3
      Top             =   0
      Width           =   8955
   End
   Begin VB.TextBox TextDatos 
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   3840
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.ListBox Lista 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FF0000&
      Height          =   4740
      ItemData        =   "Encode.frx":776D
      Left            =   960
      List            =   "Encode.frx":776F
      MouseIcon       =   "Encode.frx":7771
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.PictureBox Visor 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   6195
      Left            =   5115
      MouseIcon       =   "Encode.frx":843B
      MousePointer    =   99  'Custom
      ScaleHeight     =   6135
      ScaleWidth      =   5655
      TabIndex        =   0
      Top             =   480
      Width           =   5715
   End
   Begin ChamaleonButton.ChameleonBtn BotonI 
      Height          =   255
      Index           =   1
      Left            =   45
      TabIndex        =   35
      Top             =   1080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Bodys"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Encode.frx":9105
      PICN            =   "Encode.frx":9121
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn BotonI 
      Height          =   255
      Index           =   2
      Left            =   45
      TabIndex        =   36
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Cabezas"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Encode.frx":963B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn BotonI 
      Height          =   255
      Index           =   3
      Left            =   45
      TabIndex        =   37
      Top             =   1800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Cascos"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Encode.frx":9657
      PICN            =   "Encode.frx":9673
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn BotonI 
      Height          =   255
      Index           =   4
      Left            =   45
      TabIndex        =   38
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Escudos"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Encode.frx":9B95
      PICN            =   "Encode.frx":9BB1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn BotonI 
      Height          =   255
      Index           =   5
      Left            =   45
      TabIndex        =   39
      Top             =   2520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Armas"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Encode.frx":A19F
      PICN            =   "Encode.frx":A1BB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn BotonI 
      Height          =   255
      Index           =   6
      Left            =   45
      TabIndex        =   40
      Top             =   3600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Alas"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Encode.frx":A7A9
      PICN            =   "Encode.frx":A7C5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn BotonI 
      Height          =   255
      Index           =   7
      Left            =   45
      TabIndex        =   41
      Top             =   6030
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Capas"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Encode.frx":AC7F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn BotonI 
      Height          =   255
      Index           =   8
      Left            =   45
      TabIndex        =   42
      Top             =   2880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Fx"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Encode.frx":AC9B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn BotonI 
      Height          =   495
      Index           =   9
      Left            =   45
      TabIndex        =   43
      Top             =   4320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Ver Graficos"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Encode.frx":ACB7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn BotonI 
      Height          =   255
      Index           =   10
      Left            =   45
      TabIndex        =   44
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Botas"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Encode.frx":ACD3
      PICN            =   "Encode.frx":ACEF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn EncodeMapas 
      Height          =   495
      Left            =   45
      TabIndex        =   45
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Encode Mapas"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Encode.frx":B109
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn ChaGenerarIndices 
      Height          =   495
      Left            =   30
      TabIndex        =   46
      Top             =   5445
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Generar Indices"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   8454016
      FCOL            =   0
      FCOLO           =   255
      MCOL            =   12582912
      MPTR            =   1
      MICON           =   "Encode.frx":B125
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label DescripcionAyuda 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   2760
      TabIndex        =   30
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label LUlitError 
      BackColor       =   &H80000004&
      Height          =   975
      Left            =   105
      TabIndex        =   28
      Top             =   6705
      Width           =   12120
   End
   Begin VB.Label LGHRnumeroA 
      BackColor       =   &H00C0FFC0&
      Caption         =   "0"
      Height          =   255
      Left            =   3240
      TabIndex        =   26
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label LNumActual 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ghr:"
      Height          =   255
      Left            =   2520
      TabIndex        =   25
      Top             =   360
      Width           =   495
   End
   Begin VB.Label LTexto 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ancho Titles:"
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   21
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label LTexto 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Alto Titles:"
      Height          =   255
      Index           =   8
      Left            =   2520
      TabIndex        =   20
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label LTexto 
      BackColor       =   &H00C0FFC0&
      Caption         =   "PosicionY:"
      Height          =   255
      Index           =   7
      Left            =   2520
      TabIndex        =   19
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label LTexto 
      BackColor       =   &H00C0FFC0&
      Caption         =   "PosicionX:"
      Height          =   255
      Index           =   6
      Left            =   2520
      TabIndex        =   18
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label LTexto 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Velocidad:"
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   17
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label LTexto 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ancho:"
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   16
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label LTexto 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Alto:"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   15
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label LTexto 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Numero Frames:"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   14
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label LTexto 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Frames:"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   13
      Top             =   0
      Width           =   735
   End
   Begin VB.Label LTexto 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Numero BMP:"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
   Begin VB.Menu MenuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu MenuArchivoGuardar 
         Caption         =   "Guardar"
         Shortcut        =   ^S
      End
      Begin VB.Menu MenuBotonGuardarP 
         Caption         =   "Guardar..."
      End
      Begin VB.Menu MenuArchivoCargar 
         Caption         =   "Cargar"
         Shortcut        =   ^O
      End
      Begin VB.Menu MenuBotonCargarP 
         Caption         =   "Cargar..."
      End
      Begin VB.Menu MenuArchivoGuardado 
         Caption         =   "Indice de guardado"
         Begin VB.Menu MenuIndiceGuardado 
            Caption         =   "Sobreescribir"
            Index           =   0
         End
         Begin VB.Menu MenuIndiceGuardado 
            Caption         =   "1"
            Index           =   1
         End
         Begin VB.Menu MenuIndiceGuardado 
            Caption         =   "2"
            Index           =   2
         End
         Begin VB.Menu MenuIndiceGuardado 
            Caption         =   "3"
            Index           =   3
         End
         Begin VB.Menu MenuIndiceGuardado 
            Caption         =   "4"
            Index           =   4
         End
         Begin VB.Menu MenuIndiceGuardado 
            Caption         =   "5"
            Index           =   5
         End
         Begin VB.Menu MenuIndiceGuardado 
            Caption         =   "6"
            Index           =   6
         End
         Begin VB.Menu MenuIndiceGuardado 
            Caption         =   "7"
            Index           =   7
         End
         Begin VB.Menu MenuIndiceGuardado 
            Caption         =   "8"
            Index           =   8
         End
         Begin VB.Menu MenuIndiceGuardado 
            Caption         =   "9"
            Index           =   9
         End
      End
   End
   Begin VB.Menu Medicion 
      Caption         =   "Edicion"
      Begin VB.Menu MenuEdicionNuevo 
         Caption         =   "Nuevo/Ir A"
         Shortcut        =   ^F
      End
      Begin VB.Menu menuEdicionMover 
         Caption         =   "Mover"
      End
      Begin VB.Menu MenuEdicionCopiar 
         Caption         =   "Copiar"
      End
      Begin VB.Menu MenuEdicionBorrar 
         Caption         =   "Borrar"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu MenuEdicionClonar 
         Caption         =   "Clonar..."
      End
      Begin VB.Menu menuEdicionColor 
         Caption         =   "Color de fondo..."
      End
   End
   Begin VB.Menu MenuHerramientas 
      Caption         =   "Herramientas"
      Begin VB.Menu MenuHerramientasBG 
         Caption         =   "Buscar Grh Con bmp..."
      End
      Begin VB.Menu MenuHerramientasNI 
         Caption         =   "Buscar Bmps sin indexar (Cry)"
      End
      Begin VB.Menu MenuHerramientasBN 
         Caption         =   "Buscar siguiente"
         Shortcut        =   {F3}
         Visible         =   0   'False
      End
      Begin VB.Menu MenuHerramientasAAnim 
         Caption         =   "Autoindexador"
      End
      Begin VB.Menu MenuHerramientasBR 
         Caption         =   "Buscar Grh Repetidos"
      End
   End
   Begin VB.Menu MenuSuperficie 
      Caption         =   "Indexar Superficies"
      Begin VB.Menu MenuHerramientasWE 
         Caption         =   "Generar Indices Para WE"
      End
   End
   Begin VB.Menu MenuAyuda 
      Caption         =   "Ayuda"
      Begin VB.Menu MenuAcercaDe 
         Caption         =   "Acerca de..."
      End
   End
   Begin VB.Menu mnuautoI 
      Caption         =   "Indexar como..."
      Visible         =   0   'False
      Begin VB.Menu IAnim 
         Caption         =   "indexar como Animacion"
      End
      Begin VB.Menu mnIgeneral 
         Caption         =   "indexar como grafico individual"
      End
      Begin VB.Menu mnuibody 
         Caption         =   "indexar como body"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DibujarGHRVisor(ByVal GhrIndex As Integer)

    On Error Resume Next

    If Not GrHCambiando Then
        If GhrIndex <= 0 Then Exit Sub
        Call dibujarGrh2(Grhdata(GhrIndex))
        frmMain.Visor.Refresh
    Else
        Call dibujarGrh2(TempGrh)
        frmMain.Visor.Refresh

    End If

End Sub

Private Sub DibujarBMPVisor(ByVal GhrIndex As Integer)

    Dim SR    As RECT, DR As RECT

    Dim Alto  As Long

    Dim Ancho As Long

    frmMain.Visor.Cls

    Dim dummy As Integer

    If GhrIndex <= 0 Then Exit Sub
    '    Call GetTamañoBMP(GhrIndex, Alto, Ancho, dummy)
    '    SR.Left = 0
    '    SR.Top = 0
    '    SR.Right = Ancho
    '    SR.Bottom = Alto
    '
    '    DR.Left = 0
    '    DR.Top = 0
    '    DR.Right = SR.Right
    '    DR.Bottom = SR.Bottom
    '    Call DrawGrhtoHdc(frmMain.Visor.hWnd, frmMain.Visor.hDC, GhrIndex, SR, DR)
    Call dibujarBMP2(GhrIndex)
    frmMain.Visor.Refresh

End Sub

Private Sub DibujarDataIndex(ByRef DataIndex As BodyData, _
                             Optional ByVal Frame As Integer = 1, _
                             Optional ByVal Index As Byte = 0)

    On Error Resume Next

    Dim SR                 As RECT, DR As RECT

    Dim r                  As RECT

    Dim sourceRect         As RECT, destRect As RECT

    Dim i                  As Long

    Dim curX               As Long

    Dim curY               As Long

    Dim GhrIndex(1 To 4)   As Grh

    Dim Posiciones(1 To 4) As Position

    Dim tGrhIndex          As Long

    curX = 0
    curY = 0

    If EstadoIndexador = e_EstadoIndexador.Fx Then
        Index = 1

    End If

    With sourceRect
        .Bottom = 500
        .Left = 0
        .Right = 500
        .Top = 0

    End With

    If (Index > 0 And Index < 5) Or EstadoIndexador = e_EstadoIndexador.Fx Then
        If DataIndex.Walk(Index).GrhIndex <= 0 Then DataIndex.Walk(Index).GrhIndex = 1
        If Grhdata(DataIndex.Walk(Index).GrhIndex).NumFrames > 1 Then
            tGrhIndex = Grhdata(DataIndex.Walk(Index).GrhIndex).Frames(Frame)
        Else
            tGrhIndex = DataIndex.Walk(Index).GrhIndex

        End If

        If tGrhIndex <= 0 Then Exit Sub
        
        '        SR.Left = Grhdata(tGrhIndex).sX
        '        SR.Top = Grhdata(tGrhIndex).sY
        '        SR.Right = Grhdata(tGrhIndex).sX + Grhdata(tGrhIndex).pixelWidth
        '        SR.Bottom = Grhdata(tGrhIndex).sY + Grhdata(tGrhIndex).pixelHeight
        '
        '        DR.Left = CurX
        '        DR.Top = CurY
        '        DR.Right = CurX + Grhdata(tGrhIndex).pixelWidth
        '        DR.Bottom = CurY + Grhdata(tGrhIndex).pixelHeight
        Call dibujarGrh2(Grhdata(tGrhIndex))
        'Call DrawGrhtoHdc(frmMain.Visor.hWnd, frmMain.Visor.hDC, Grhdata(tGrhIndex).FileNum, SR, DR)
        frmMain.Visor.Refresh
    Else

        If DibujarFondo Then
            BackBufferSurface.BltColorFill r, ColorFondo
        Else
            BackBufferSurface.BltColorFill r, 0

        End If

        Call CalcularPosiciones(DataIndex, Posiciones)

        For i = 1 To 4

            If DataIndex.Walk(i).GrhIndex <= 0 Then DataIndex.Walk(i).GrhIndex = 1
            If Grhdata(DataIndex.Walk(i).GrhIndex).NumFrames > 1 Then
                tGrhIndex = Grhdata(DataIndex.Walk(i).GrhIndex).Frames(Frame)
            Else
                tGrhIndex = DataIndex.Walk(i).GrhIndex

            End If

            If tGrhIndex <= 0 Then Exit Sub
        
            SR.Left = Grhdata(tGrhIndex).sX
            SR.Top = Grhdata(tGrhIndex).sY
            SR.Right = Grhdata(tGrhIndex).sX + Grhdata(tGrhIndex).pixelWidth
            SR.Bottom = Grhdata(tGrhIndex).sY + Grhdata(tGrhIndex).pixelHeight
        
            DR.Left = Posiciones(i).X
            DR.Top = Posiciones(i).Y
            DR.Right = Posiciones(i).X + Grhdata(tGrhIndex).pixelWidth
            DR.Bottom = Posiciones(i).Y + Grhdata(tGrhIndex).pixelHeight

            Call dibujapjESpecial(BackBufferSurface, Grhdata(tGrhIndex), DR.Left, DR.Top)
        
        Next i

        If EstadoIndexador = e_EstadoIndexador.Body And cabezaActual <> 0 Then
            If cabezaActual > 0 And cabezaActual <= MAXGrH Then
                Call dibujapjESpecial(BackBufferSurface, Grhdata(cabezaActual), Posiciones(3).X + (Grhdata(Grhdata(DataIndex.Walk(3).GrhIndex).Frames(Frame)).pixelWidth / 2) - (Grhdata(cabezaActual).pixelWidth / 2) + DataIndex.HeadOffset.X, Posiciones(3).Y + Grhdata(Grhdata(DataIndex.Walk(3).GrhIndex).Frames(Frame)).pixelHeight - Grhdata(cabezaActual).pixelHeight + DataIndex.HeadOffset.Y - 1)

            End If

        End If

        If EstadoIndexador = e_EstadoIndexador.Cabezas Then
            If frmMain.Checkcabeza.value = vbChecked Then
                cabezaActual = DataIndex.Walk(3).GrhIndex

            End If

        End If

        'SecundaryClipper.SetHWnd frmMain.Visor.hWnd
        If DataIndex.Walk(4).GrhIndex > 0 Then
            sourceRect.Right = Posiciones(2).X + Grhdata(DataIndex.Walk(4).GrhIndex).pixelWidth
        Else
            sourceRect.Right = Posiciones(2).X * 2

        End If

        If DataIndex.Walk(3).GrhIndex > 0 Then
            sourceRect.Bottom = Posiciones(3).Y + Grhdata(DataIndex.Walk(3).GrhIndex).pixelHeight
        Else
            sourceRect.Bottom = Posiciones(3).Y * 2

        End If

        destRect = sourceRect
        BackBufferSurface.BltToDC frmMain.Visor.hDC, sourceRect, destRect
    
        frmMain.Visor.Refresh

    End If

End Sub

Private Sub DibujarTempGHRVisor(ByVal loopAnim As Integer)

    On Error Resume Next

    Dim SR       As RECT, DR As RECT

    Dim GhrIndex As Integer

    GhrIndex = loopAnim

    If GhrIndex <= 0 Then Exit Sub
    
    SR.Left = Grhdata(GhrIndex).sX
    SR.Top = Grhdata(GhrIndex).sY
    SR.Right = Grhdata(GhrIndex).sX + Grhdata(GhrIndex).pixelWidth
    SR.Bottom = Grhdata(GhrIndex).sY + Grhdata(GhrIndex).pixelHeight
    
    DR.Left = 0
    DR.Top = 0
    DR.Right = Grhdata(GhrIndex).pixelWidth
    DR.Bottom = Grhdata(GhrIndex).pixelHeight
    Call DrawGrhtoHdc(frmMain.Visor.hwnd, frmMain.Visor.hDC, Grhdata(GhrIndex).FileNum, SR, DR)
    frmMain.Visor.Refresh

End Sub

Private Sub GetInfoGHR(ByVal GrhIndex As Long)

    If GrhIndex <= 0 Then Exit Sub
    LoadingNew = True

    Dim i           As Long

    Dim Ancho       As Long

    Dim Alto        As Long

    Dim PrimerAlto  As Long

    Dim PrimerAncho As Long

    Dim dumYin      As Integer

    TextDatos(0).Text = Grhdata(GrhIndex).FileNum
    TextDatos(1).Text = ""
    TextDatos(2).Text = Grhdata(GrhIndex).NumFrames

    TextDatos(3).Text = Grhdata(GrhIndex).pixelHeight
    TextDatos(4).Text = Grhdata(GrhIndex).pixelWidth
    TextDatos(5).Text = Grhdata(GrhIndex).Speed
    TextDatos(6).Text = Grhdata(GrhIndex).sX
    TextDatos(7).Text = Grhdata(GrhIndex).sY
    TextDatos(8).Text = Grhdata(GrhIndex).TileHeight
    TextDatos(9).Text = Grhdata(GrhIndex).TileWidth
    LUlitError.Caption = ""

    If Grhdata(GrhIndex).NumFrames = 1 Then
        TextDatos(1).BackColor = vbWhite
        TextDatos(1).Text = Grhdata(GrhIndex).Frames(1)
        Call GetTamañoBMP(Grhdata(GrhIndex).FileNum, Alto, Ancho, dumYin)
        frmMain.Dibujado.Enabled = False
        TextDatos(1).Enabled = False

        For i = 3 To 4
            TextDatos(i).Enabled = True
        Next i

        TextDatos(5).Enabled = False

        For i = 6 To 7
            TextDatos(i).Enabled = True
        Next i

    Else
        TextDatos(1).BackColor = vbWhite

        For i = 1 To Grhdata(GrhIndex).NumFrames

            If i = 1 Then
                TextDatos(1).Text = Grhdata(GrhIndex).Frames(i)
            Else
                TextDatos(1).Text = TextDatos(1).Text & "-" & Grhdata(GrhIndex).Frames(i)

            End If

        Next i

        If Grhdata(GrhIndex).Speed > 0 Then ' pervenimos division por 0
            frmMain.Dibujado.Interval = 50 * Grhdata(GrhIndex).Speed
        Else
            frmMain.Dibujado.Interval = 100

        End If

        frmMain.Dibujado.Enabled = True
        TextDatos(1).Enabled = True

        For i = 3 To 4
            TextDatos(i).Enabled = False
            TextDatos(i).BackColor = vbWhite
        Next i

        TextDatos(5).Enabled = True

        For i = 6 To 7
            TextDatos(i).Enabled = False
            TextDatos(i).BackColor = vbWhite
        Next i

    End If
If FraDatosIndice.Visible = True Then
    If frmMain.TextDatos(3).Text = 32 Then
        frmMain.TextDatos(3).Text = 4
        frmMain.txtAlto.Text = TextDatos(3).Text
        frmMain.TextDatos(3).Text = 32
        ElseIf frmMain.TextDatos(3).Text > 283 And frmMain.TextDatos(3).Text < 301 Then
            frmMain.TextDatos(3).Text = 0
            frmMain.txtAlto.Text = TextDatos(3).Text
            frmMain.TextDatos(3).Text = Grhdata(GrhIndex).pixelHeight
    Else
        frmMain.TextDatos(3).Text = Grhdata(GrhIndex).pixelHeight
        frmMain.txtAlto.Text = TextDatos(3).Text

    End If

    If frmMain.TextDatos(4).Text = 32 Then
        frmMain.TextDatos(4).Text = 4
        frmMain.txtAncho.Text = TextDatos(4).Text
        frmMain.TextDatos(4) = 32
     ElseIf frmMain.TextDatos(4).Text > 255 And frmMain.TextDatos(4).Text < 321 Then
        frmMain.TextDatos(4).Text = 0
        frmMain.txtAncho.Text = TextDatos(4).Text
        frmMain.TextDatos(4) = Grhdata(GrhIndex).pixelWidth
     Else
        frmMain.TextDatos(4).Text = Grhdata(GrhIndex).pixelWidth
        frmMain.txtAncho.Text = TextDatos(4).Text

    End If
        
    frmMain.TxtGrhIndex.Text = GRHActual
End If
    GrHCambiando = False
    LNumActual.Caption = "Ghr:"
    BotonGuardar.Visible = False
    LoadingNew = False

End Sub

Private Sub GetInfoBmp(ByVal GrhIndex As Long)

    If GrhIndex <= 0 Then Exit Sub

    Dim i             As Long

    Dim Ancho         As Long

    Dim Alto          As Long

    Dim PrimerAlto    As Long

    Dim PrimerAncho   As Long

    Dim BitCount      As Integer

    Dim existenciaBMP As Byte

    Dim ResourceS     As String

    existenciaBMP = ExisteBMP(GrhIndex)

    If existenciaBMP = 0 Then Exit Sub
    If existenciaBMP = 1 And ResourceFile = 3 Then
        If GrhIndex > 0 And GrhIndex <= UBound(ResourceF.graficos) Then
            If ResourceF.graficos(GrhIndex).tamaño > 0 Then ResourceS = "+ ResF"

        End If

    End If

    Call GetTamañoBMP(GrhIndex, Alto, Ancho, BitCount)
     

Dim tamano As Long
tamano = FileLen(App.Path & "\" & CarpetaGraficos & "\" & GrhIndex & ".bmp")
    If existenciaBMP = 2 Then TextDatos(0).Text = ResourceF.graficos(GrhIndex).tamaño
    TextDatos(0).Text = tamano
    TextDatos(1).Text = ""
    TextDatos(2).Text = Alto

    TextDatos(3).Text = Ancho
    TextDatos(4).Text = BitCount
    TextDatos(5).Text = StringRecurso(existenciaBMP)

    If ResourceS <> vbNullString Then TextDatos(5).Text = TextDatos(5).Text & ResourceS

    LNumActual.Caption = "BMP:"
    BotonGuardar.Visible = False

End Sub

Private Sub GetInfoDataIndex(ByVal DataIndex As Integer)

    If DataIndex <= 0 Then Exit Sub

    Dim i           As Long

    Dim Ancho       As Long

    Dim Alto        As Long

    Dim PrimerAlto  As Long

    Dim PrimerAncho As Long

    LoadingNew = True

    Dim GhrIndex(1 To 4) As Integer

    Dim tGrhIndex        As Long

    TextDatos(5).Visible = False
    LTexto(5).Visible = False
    TextDatos(5).Text = ""
    LUlitError.Caption = ""

    For i = 1 To 4

        If EstadoIndexador = e_EstadoIndexador.Body Then
            GhrIndex(i) = BodyData(DataIndex).Walk(i).GrhIndex
            tempDataIndex.Walk(i).GrhIndex = BodyData(DataIndex).Walk(i).GrhIndex
        ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
            GhrIndex(i) = HeadData(DataIndex).Head(i).GrhIndex
            tempDataIndex.Walk(i).GrhIndex = HeadData(DataIndex).Head(i).GrhIndex
        ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
            GhrIndex(i) = CascoAnimData(DataIndex).Head(i).GrhIndex
            tempDataIndex.Walk(i).GrhIndex = CascoAnimData(DataIndex).Head(i).GrhIndex
        ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
            GhrIndex(i) = ShieldAnimData(DataIndex).ShieldWalk(i).GrhIndex
            tempDataIndex.Walk(i).GrhIndex = ShieldAnimData(DataIndex).ShieldWalk(i).GrhIndex
        ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
            GhrIndex(i) = WeaponAnimData(DataIndex).WeaponWalk(i).GrhIndex
            tempDataIndex.Walk(i).GrhIndex = WeaponAnimData(DataIndex).WeaponWalk(i).GrhIndex
        ElseIf EstadoIndexador = e_EstadoIndexador.Alas Then
            GhrIndex(i) = AlaData(DataIndex).Alas(i).GrhIndex
            tempDataIndex.Walk(i).GrhIndex = AlaData(DataIndex).Alas(i).GrhIndex
        ElseIf EstadoIndexador = e_EstadoIndexador.Capas Then
            GhrIndex(i) = EspaldaAnimData(DataIndex).Head(i).GrhIndex
            tempDataIndex.Walk(i).GrhIndex = EspaldaAnimData(DataIndex).Head(i).GrhIndex
        ElseIf EstadoIndexador = e_EstadoIndexador.Fx Then
            GhrIndex(i) = FxData(DataIndex).Fx.GrhIndex
            tempDataIndex.Walk(i).GrhIndex = FxData(DataIndex).Fx.GrhIndex
        ElseIf EstadoIndexador = e_EstadoIndexador.Botas Then
            GhrIndex(i) = BotaData(DataIndex).Bota(i).GrhIndex
            tempDataIndex.Walk(i).GrhIndex = BotaData(DataIndex).Bota(i).GrhIndex

        End If

    Next i

    TextDatos(0).Text = GhrIndex(1)
    TextDatos(2).Text = GhrIndex(2)

    TextDatos(3).Text = GhrIndex(3)
    TextDatos(4).Text = GhrIndex(4)

    If EstadoIndexador = e_EstadoIndexador.Body Then
        TextDatos(5).Text = BodyData(DataIndex).HeadOffset.Y & "º" & BodyData(DataIndex).HeadOffset.X
        tempDataIndex.HeadOffset.X = BodyData(DataIndex).HeadOffset.X
        tempDataIndex.HeadOffset.Y = BodyData(DataIndex).HeadOffset.Y
        TextDatos(5).Visible = True
        LTexto(5).Visible = True
    ElseIf EstadoIndexador = e_EstadoIndexador.Fx Then
        TextDatos(2).Text = FxData(DataIndex).OffsetY & "º" & FxData(DataIndex).OffsetX
        tempDataIndex.HeadOffset.X = FxData(DataIndex).OffsetX
        tempDataIndex.HeadOffset.Y = FxData(DataIndex).OffsetY
        TextDatos(2).Visible = True
        LTexto(2).Visible = True

    End If

    GrHCambiando = False
    BotonGuardar.Visible = False
    
    LoadingNew = False

End Sub

Private Sub BotonBorrrar_Click()
    Call SBotonBorrrar

End Sub

Public Sub CambiarEstado(ByVal Index As Integer)
    ' Cambia el estado del indexador entre las distintas secciones. Oculta/cambia labels

    Dim i As Long

    EstadoIndexador = Index
    Dibujado.Enabled = False
    Visor.Cls
    Lista.Clear
    GrHCambiando = False
    CDibujarWalk.Visible = False
    LUlitError.Caption = ""
    MenuEdicionClonar.Visible = False
    MenuHerramientas.Visible = False
    Command10.Visible = True
    BotonBorrrar.Visible = True
    DescripcionAyuda.Visible = False
    Checkcabeza.Visible = False

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh
            Call RenuevaListaGrH   'mostramos lista de grhs

            For i = 0 To 9
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                LTexto(i).Visible = True
            Next i

            MenuHerramientas.Visible = True
            MenuHerramientasBN.Visible = False
            MenuEdicionClonar.Visible = True
            LNumActual.Caption = "Grh: "
            LTexto(0).Caption = "Numero BMP:"
            LTexto(1).Caption = "Frames:"
            LTexto(2).Caption = "Numero Frames:"
            LTexto(3).Caption = "Alto:"
            LTexto(4).Caption = "Ancho:"
            LTexto(5).Caption = "Velocidad:"
            LTexto(6).Caption = "PosicionX:"
            LTexto(7).Caption = "PosicionY:"
            LTexto(8).Caption = "Alto Titles:"
            LTexto(9).Caption = "Ancho Titles:"
            LUlitError.Caption = ""

        Case e_EstadoIndexador.Body
            Checkcabeza.Visible = True
            CDibujarWalk.Visible = True
            CDibujarWalk.listIndex = 0
            Call RenuevaListaBodys

            For i = 0 To 5
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i

            LTexto(1).Visible = False
            TextDatos(1).Visible = False

            For i = 6 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i

            LNumActual.Caption = "Body: "
            LTexto(0).Caption = "Subir:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "derecha:"
            LTexto(3).Caption = "Abajo:"
            LTexto(4).Caption = "izquierda:"
            LTexto(5).Caption = "Offset"
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""

        Case e_EstadoIndexador.Cabezas
            CDibujarWalk.Visible = True
            CDibujarWalk.listIndex = 0
            Call RenuevaListaCabezas

            For i = 0 To 4
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i

            LTexto(1).Visible = False
            TextDatos(1).Visible = False

            For i = 5 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i

            LNumActual.Caption = "Cabeza: "
            LTexto(0).Caption = "Subir:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "derecha:"
            LTexto(3).Caption = "Abajo:"
            LTexto(4).Caption = "izquierda:"
            LTexto(5).Caption = ""
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""

        Case e_EstadoIndexador.Cascos
            CDibujarWalk.Visible = True
            CDibujarWalk.listIndex = 0
            Call RenuevaListaCascos

            For i = 0 To 4
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i

            LTexto(1).Visible = False
            TextDatos(1).Visible = False

            For i = 5 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i

            LNumActual.Caption = "Casco: "
            LTexto(0).Caption = "Subir:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "derecha:"
            LTexto(3).Caption = "Abajo:"
            LTexto(4).Caption = "izquierda:"
            LTexto(5).Caption = ""
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""

        Case e_EstadoIndexador.Escudos
            CDibujarWalk.Visible = True
            CDibujarWalk.listIndex = 0
            Call RenuevaListaEscudos

            For i = 0 To 4
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i

            LTexto(1).Visible = False
            TextDatos(1).Visible = False

            For i = 5 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i

            LNumActual.Caption = "Escudo: "
            LTexto(0).Caption = "Subir:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "derecha:"
            LTexto(3).Caption = "Abajo:"
            LTexto(4).Caption = "izquierda:"
            LTexto(5).Caption = ""
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""

        Case e_EstadoIndexador.Armas
            CDibujarWalk.Visible = True
            CDibujarWalk.listIndex = 0
            Call RenuevaListaArmas

            For i = 0 To 4
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i

            LTexto(1).Visible = False
            TextDatos(1).Visible = False

            For i = 5 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i

            LNumActual.Caption = "Armas: "
            LTexto(0).Caption = "Subir:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "derecha:"
            LTexto(3).Caption = "Abajo:"
            LTexto(4).Caption = "izquierda:"
            LTexto(5).Caption = ""
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""

        Case e_EstadoIndexador.Alas
            Checkcabeza.Visible = True
            CDibujarWalk.Visible = True
            CDibujarWalk.listIndex = 0
            Call RenuevaListaAlas

            For i = 0 To 5
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i

            LTexto(1).Visible = False
            TextDatos(1).Visible = False

            For i = 6 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i

            LNumActual.Caption = "Alas: "
            LTexto(0).Caption = "Subir:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "derecha:"
            LTexto(3).Caption = "Abajo:"
            LTexto(4).Caption = "izquierda:"
            LTexto(5).Caption = "Offset"
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""

        Case e_EstadoIndexador.Capas
            CDibujarWalk.Visible = True
            CDibujarWalk.listIndex = 0
            Call RenuevaListaCapas

            For i = 0 To 4
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i

            LTexto(1).Visible = False
            TextDatos(1).Visible = False

            For i = 5 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i

            LNumActual.Caption = "Capa: "
            LTexto(0).Caption = "Subir:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "derecha:"
            LTexto(3).Caption = "Abajo:"
            LTexto(4).Caption = "izquierda:"
            LTexto(5).Caption = ""
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""

        Case e_EstadoIndexador.Fx
            Call RenuevaListaFX

            For i = 0 To 2
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i

            LTexto(1).Visible = False
            TextDatos(1).Visible = False

            For i = 3 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i

            LNumActual.Caption = "Fx: "
            LTexto(0).Caption = "NumGrh:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "Offset:"
            LTexto(3).Caption = ""
            LTexto(4).Caption = ""
            LTexto(5).Caption = ""
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""

        Case e_EstadoIndexador.Botas
            CDibujarWalk.Visible = True
            CDibujarWalk.listIndex = 0
            Call RenuevaListaBotas

            For i = 0 To 4
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i

            LTexto(1).Visible = False
            TextDatos(1).Visible = False

            For i = 5 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i

            LNumActual.Caption = "Botas: "
            LTexto(0).Caption = "Subir:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "derecha:"
            LTexto(3).Caption = "Abajo:"
            LTexto(4).Caption = "izquierda:"
            LTexto(5).Caption = ""
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""
            
        Case e_EstadoIndexador.Resource
            Call RenuevaListaResource

            For i = 0 To 5
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i

            LTexto(1).Visible = False
            TextDatos(1).Visible = False

            For i = 6 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i

            LNumActual.Caption = "Crypt: "
            LTexto(0).Caption = "Tamaño:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "Alto:"
            LTexto(3).Caption = "Ancho:"
            LTexto(4).Caption = "Bits:"
            LTexto(5).Caption = "Situacion:"
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""
            Command10.Visible = False
            BotonBorrrar.Visible = False
            DescripcionAyuda.Visible = True
            DescripcionAyuda.Caption = "N:Si estan disponible el BMP y el archivo de recursos, se usa el bmp"

    End Select

    Call CambiarcaptionCommand10

End Sub

Private Sub MoverGrh(ByVal numGRH As Integer, _
                     ByVal OrigenGRH As Integer, _
                     ByVal BorrarOriginal As Boolean)

    Dim tempLong  As Long

    Dim cadena    As String

    Dim respuesta As Byte

    Dim GrhVacio  As Grhdata

    Dim looPero   As Long

    tempLong = ListaindexGrH(OrigenGRH)

    If tempLong <= 0 Then
        LUlitError.Caption = "grafico incorrecto"
        Exit Sub

    End If

    tempLong = ListaindexGrH(numGRH)

    If tempLong > 0 Then
        respuesta = MsgBox("El grafico " & numGRH & " ya existe, ¿deseas sobreescribirlo?", 4, "Aviso")

        If respuesta = vbYes Then
            Grhdata(numGRH) = Grhdata(OrigenGRH)

            If BorrarOriginal Then
                Grhdata(OrigenGRH) = GrhVacio

            End If

            GRHActual = Val(numGRH)
            LOOPActual = 1
            'frmMain.Visor.Cls
            'Call DibujarGHRVisor(GRHActual)
            'Call GetInfoGHR(GRHActual)
            LGHRnumeroA.Caption = GRHActual
            tempLong = ListaindexGrH(GRHActual)
            frmMain.Lista.listIndex = tempLong
            EstadoNoGuardado(e_EstadoIndexador.Grh) = True

        End If

    Else
        Grhdata(numGRH) = Grhdata(OrigenGRH)

        If BorrarOriginal Then
            Grhdata(OrigenGRH) = GrhVacio

        End If

        GRHActual = numGRH
        LOOPActual = 1
        'frmMain.Visor.Cls
        'Call DibujarGHRVisor(GRHActual)
        'Call GetInfoGHR(GRHActual)
        LGHRnumeroA.Caption = GRHActual
        tempLong = ListaindexGrH(GRHActual)
        frmMain.Lista.listIndex = tempLong
        EstadoNoGuardado(e_EstadoIndexador.Grh) = True

    End If
    
End Sub

Private Sub SBotonMover(ByVal BorrarOriginal As Boolean, _
                        Optional ByVal CantidadM As Integer = 1)

    Dim tempLong  As Long

    Dim cadena    As String

    Dim respuesta As Byte

    Dim GrhVacio  As Grhdata

    Dim LooPer    As Long

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh

            If GrHCambiando Then
                GrHCambiando = False
                LNumActual.Caption = "Ghr:"
                BotonGuardar.Visible = False

            End If

            cadena = InputBox("Introduzca número de GHR al que quieres mover el grafico " & GRHActual, "Mover Grafico")

            If IsNumeric(cadena) Then
                If Val(cadena) > 0 And Val(cadena) < MAXGrH Then
                    Call MoverGrh(Val(cadena), GRHActual, BorrarOriginal)
                    Call RenuevaListaGrH
                    tempLong = ListaindexGrH(Val(cadena))
                    frmMain.Lista.listIndex = tempLong
                Else
                    LUlitError.Caption = "introduzca un numero correcto"

                End If

            Else
                LUlitError.Caption = "introduzca un numero"

            End If

        Case Else

            Dim StringCaso As String

            If EstadoIndexador = Body Then
                StringCaso = "Body"
            ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
                StringCaso = "Cabeza"
            ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
                StringCaso = "Casco"
            ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
                StringCaso = "Escudo"
            ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
                StringCaso = "Arma"
            ElseIf EstadoIndexador = e_EstadoIndexador.Alas Then
                StringCaso = "Alas"
            ElseIf EstadoIndexador = e_EstadoIndexador.Capas Then
                StringCaso = "Capa"
            ElseIf EstadoIndexador = e_EstadoIndexador.Fx Then
                StringCaso = "Fx"
            ElseIf EstadoIndexador = e_EstadoIndexador.Botas Then
                StringCaso = "Bota"
            ElseIf EstadoIndexador = e_EstadoIndexador.Resource Then
                Exit Sub

            End If

            cadena = InputBox("Introduzca numero de " & StringCaso & " al que quieres mover", "Mover " & StringCaso)

            If IsNumeric(cadena) And (Val(cadena) < 31000) Then
                If EstadoIndexador = e_EstadoIndexador.Body Then
                    Call mueveBody(Val(cadena), DataIndexActual, BorrarOriginal)
                ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
                    Call MueveCabeza(Val(cadena), DataIndexActual, BorrarOriginal)
                ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
                    Call MueveCasco(Val(cadena), DataIndexActual, BorrarOriginal)
                ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
                    Call MueveEscudo(Val(cadena), DataIndexActual, BorrarOriginal)
                ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
                    Call MueveArma(Val(cadena), DataIndexActual, BorrarOriginal)
                ElseIf EstadoIndexador = e_EstadoIndexador.Alas Then
                    Call MueveAlas(Val(cadena), DataIndexActual, BorrarOriginal)
                ElseIf EstadoIndexador = e_EstadoIndexador.Capas Then
                    Call MueveCapa(Val(cadena), DataIndexActual, BorrarOriginal)
                ElseIf EstadoIndexador = e_EstadoIndexador.Fx Then
                    Call MueveFX(Val(cadena), DataIndexActual, BorrarOriginal)
                ElseIf EstadoIndexador = e_EstadoIndexador.Botas Then
                    Call MueveBota(Val(cadena), DataIndexActual, BorrarOriginal)

                End If

                DataIndexActual = Val(cadena)
                LOOPActual = 1
                frmMain.Visor.Cls
                Call GetInfoDataIndex(DataIndexActual)
                Dibujado.Interval = 100
                Dibujado.Enabled = True
                LGHRnumeroA.Caption = DataIndexActual
                tempLong = ListaindexGrH(DataIndexActual)
                frmMain.Lista.listIndex = tempLong
                EstadoNoGuardado(EstadoIndexador) = True
            Else
                LUlitError.Caption = "introduzca un numero valido"

            End If

    End Select

End Sub

Private Sub BotonI_Click(Index As Integer)
     If FraDatosIndice.Visible = True Then
        ChaGenerarIndices.Visible = False
        lblDescripcion.Visible = False
        txtNombre.Visible = False
        FraDatosIndice.Visible = False
    Else
        ChaGenerarIndices.Visible = True
        lblDescripcion.Visible = True
        txtNombre.Visible = True
        FraDatosIndice.Visible = True
    End If
    
    Call CambiarEstado(Index)

    Call ComprobarIndexLista

End Sub

Private Sub CDibujarWalk_Click()
    DibujarWalk = CDibujarWalk.listIndex
    Visor.Cls

End Sub

Private Sub ChaGenerarIndices_Click()

    
    If txtNombre.Text = "" Then MsgBox "Debes ponerle un Nombre": Exit Sub
    If txtAlto.Text = "" Then MsgBox "Elige el terreno de la lista de los Indices": Exit Sub
    'Crea el ficherocon los datos

    GRHActual = Val(ReadField(1, Lista.List(Lista.listIndex), Asc(" ")))
 
    If Salto = False Then
        Call WriteVar(PathDat & "Indices.ini", "INIT", "Referencias", ReferenciaPrincipal & vbCrLf)
        Salto = True
    Else
        Call WriteVar(PathDat & "Indices.ini", "INIT", "Referencias", ReferenciaPrincipal)

    End If

    'Poner un salto
    
    Call WriteVar(PathDat & "Indices.ini", "REFERENCIA" & ReferenciaDos, "Nombre ", txtNombre)
    Call WriteVar(PathDat & "Indices.ini", "REFERENCIA" & ReferenciaDos, "GrhIndice ", LGHRnumeroA)
    Call WriteVar(PathDat & "Indices.ini", "REFERENCIA" & ReferenciaDos, "Alto ", frmMain.txtAlto.Text) 'frmMain.TextDatos(3))
    Call WriteVar(PathDat & "Indices.ini", "REFERENCIA" & ReferenciaDos, "Ancho ", frmMain.txtAncho.Text & vbCrLf) 'frmMain.TextDatos(4) & vbCrLf)

    'Poner otro salto

    ReferenciaPrincipal = ReferenciaPrincipal + 1
    ReferenciaDos = ReferenciaDos + 1
    TxtReferencias.Text = ReferenciaDos
End Sub

Private Sub Checkcabeza_Click()

    If Checkcabeza.value = vbChecked Then
        cabezaActual = 3008
    Else
        cabezaActual = 0

    End If

End Sub

Private Sub CmbResourceFile_Click()
    'Borramos todas las surfaces de la memoria. Sirve por si se hacen cambios en los BMPs y se necesita obligar a recargarlos
    
    If Not IniciadoTodo Then
        Call SurfaceDB.BorrarTodo
    Else
        IniciadoTodo = False

    End If
    
    Call CambiarEstado(EstadoIndexador)
    
    Call ComprobarIndexLista

End Sub

Private Sub SBotonBorrrar()

    Dim respuesta As Byte

    Dim tempLong  As Long

    tempLong = frmMain.Lista.listIndex

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh

            If GRHActual = 0 Then Exit Sub
            respuesta = MsgBox("ATENCION ¿Estas segudo de borrar el Grh " & GRHActual & " ?", 4, "¡¡ADVERTENCIA!!")

            If respuesta = vbYes Then
                Grhdata(GRHActual).NumFrames = 0
                'Call RenuevaListaGrH
                frmMain.Lista.RemoveItem tempLong

            End If

        Case e_EstadoIndexador.Body

            If DataIndexActual = 0 Then Exit Sub
            respuesta = MsgBox("ATENCION ¿Estas segudo de borrar el body " & DataIndexActual & " ?", 4, "¡¡ADVERTENCIA!!")

            If respuesta = vbYes Then
                BodyData(DataIndexActual).Walk(1).GrhIndex = 0
                BodyData(DataIndexActual).Walk(2).GrhIndex = 0
                BodyData(DataIndexActual).Walk(3).GrhIndex = 0
                BodyData(DataIndexActual).Walk(4).GrhIndex = 0
                'Call RenuevaListaBodys
                frmMain.Lista.RemoveItem tempLong

            End If

        Case e_EstadoIndexador.Cabezas

            If DataIndexActual = 0 Then Exit Sub
            respuesta = MsgBox("ATENCION ¿Estas segudo de borrar la Cabeza " & DataIndexActual & " ?", 4, "¡¡ADVERTENCIA!!")

            If respuesta = vbYes Then
                HeadData(DataIndexActual).Head(1).GrhIndex = 0
                HeadData(DataIndexActual).Head(2).GrhIndex = 0
                HeadData(DataIndexActual).Head(3).GrhIndex = 0
                HeadData(DataIndexActual).Head(4).GrhIndex = 0
                'Call RenuevaListaCabezas
                frmMain.Lista.RemoveItem tempLong

            End If

        Case e_EstadoIndexador.Cascos

            If DataIndexActual = 0 Then Exit Sub
            respuesta = MsgBox("ATENCION ¿Estas segudo de borrar el casco " & DataIndexActual & " ?", 4, "¡¡ADVERTENCIA!!")

            If respuesta = vbYes Then
                CascoAnimData(DataIndexActual).Head(1).GrhIndex = 0
                CascoAnimData(DataIndexActual).Head(2).GrhIndex = 0
                CascoAnimData(DataIndexActual).Head(3).GrhIndex = 0
                CascoAnimData(DataIndexActual).Head(4).GrhIndex = 0
                'Call RenuevaListaCascos
                frmMain.Lista.RemoveItem tempLong

            End If

        Case e_EstadoIndexador.Escudos

            If DataIndexActual = 0 Then Exit Sub
            respuesta = MsgBox("ATENCION ¿Estas segudo de borrar el escudo " & DataIndexActual & " ?", 4, "¡¡ADVERTENCIA!!")

            If respuesta = vbYes Then
                ShieldAnimData(DataIndexActual).ShieldWalk(1).GrhIndex = 0
                ShieldAnimData(DataIndexActual).ShieldWalk(2).GrhIndex = 0
                ShieldAnimData(DataIndexActual).ShieldWalk(3).GrhIndex = 0
                ShieldAnimData(DataIndexActual).ShieldWalk(4).GrhIndex = 0
                frmMain.Lista.RemoveItem tempLong

                'Call RenuevaListaEscudos
            End If

        Case e_EstadoIndexador.Armas

            If DataIndexActual = 0 Then Exit Sub
            respuesta = MsgBox("ATENCION ¿Estas segudo de borrar el arma " & DataIndexActual & " ?", 4, "¡¡ADVERTENCIA!!")

            If respuesta = vbYes Then
                WeaponAnimData(DataIndexActual).WeaponWalk(1).GrhIndex = 0
                WeaponAnimData(DataIndexActual).WeaponWalk(2).GrhIndex = 0
                WeaponAnimData(DataIndexActual).WeaponWalk(3).GrhIndex = 0
                WeaponAnimData(DataIndexActual).WeaponWalk(4).GrhIndex = 0
                'Call RenuevaListaArmas
                frmMain.Lista.RemoveItem tempLong

            End If

        Case e_EstadoIndexador.Alas

            If DataIndexActual = 0 Then Exit Sub
            respuesta = MsgBox("ATENCION ¿Estas segudo de borrar el ala " & DataIndexActual & " ?", 4, "¡¡ADVERTENCIA!!")

            If respuesta = vbYes Then
                AlaData(DataIndexActual).Alas(1).GrhIndex = 0
                AlaData(DataIndexActual).Alas(2).GrhIndex = 0
                AlaData(DataIndexActual).Alas(3).GrhIndex = 0
                AlaData(DataIndexActual).Alas(4).GrhIndex = 0
                'Call RenuevaListaBotas
                frmMain.Lista.RemoveItem tempLong

            End If

        Case e_EstadoIndexador.Capas

            If DataIndexActual = 0 Then Exit Sub
            respuesta = MsgBox("ATENCION ¿Estas segudo de borrar el capa " & DataIndexActual & " ?", 4, "¡¡ADVERTENCIA!!")

            If respuesta = vbYes Then
                EspaldaAnimData(DataIndexActual).Head(1).GrhIndex = 0
                EspaldaAnimData(DataIndexActual).Head(2).GrhIndex = 0
                EspaldaAnimData(DataIndexActual).Head(3).GrhIndex = 0
                EspaldaAnimData(DataIndexActual).Head(4).GrhIndex = 0
                'Call RenuevaListaCapas
                frmMain.Lista.RemoveItem tempLong

            End If

        Case e_EstadoIndexador.Fx

            If DataIndexActual = 0 Then Exit Sub
            respuesta = MsgBox("ATENCION ¿Estas segudo de borrar el FX " & DataIndexActual & " ?", 4, "¡¡ADVERTENCIA!!")

            If respuesta = vbYes Then
                FxData(DataIndexActual).Fx.GrhIndex = 0
                FxData(DataIndexActual).OffsetX = 0
                FxData(DataIndexActual).OffsetY = 0
                'Call RenuevaListaFX
                frmMain.Lista.RemoveItem tempLong

            End If

        Case e_EstadoIndexador.Botas

            If DataIndexActual = 0 Then Exit Sub
            respuesta = MsgBox("ATENCION ¿Estas segudo de borrar la Botas " & DataIndexActual & " ?", 4, "¡¡ADVERTENCIA!!")

            If respuesta = vbYes Then
                BotaData(DataIndexActual).Bota(1).GrhIndex = 0
                BotaData(DataIndexActual).Bota(2).GrhIndex = 0
                BotaData(DataIndexActual).Bota(3).GrhIndex = 0
                BotaData(DataIndexActual).Bota(4).GrhIndex = 0
                'Call RenuevaListaBotas
                frmMain.Lista.RemoveItem tempLong

            End If

    End Select

    If tempLong < frmMain.Lista.ListCount Then
        frmMain.Lista.listIndex = tempLong
    Else
        frmMain.Lista.listIndex = frmMain.Lista.ListCount - 1

    End If

End Sub

Public Function StringGuardadoActual(PEstado As e_EstadoIndexador) As String

    Dim elq As String

    Select Case PEstado

        Case e_EstadoIndexador.Grh

            If SavePath = 0 Then
                elq = "Graficos"
            Else
                elq = "Graficos" & SavePath

            End If

        Case e_EstadoIndexador.Body

            If SavePath = 0 Then
                elq = "personajes"
            Else
                elq = "personajes" & SavePath

            End If

        Case e_EstadoIndexador.Cabezas

            If SavePath = 0 Then
                elq = "cabezas"
            Else
                elq = "cabezas" & SavePath

            End If

        Case e_EstadoIndexador.Cascos

            If SavePath = 0 Then
                elq = "cascos"
            Else
                elq = "cascos" & SavePath

            End If

        Case e_EstadoIndexador.Escudos

            If SavePath = 0 Then
                elq = "escudos"
            Else
                elq = "escudos" & SavePath

            End If

        Case e_EstadoIndexador.Armas

            If SavePath = 0 Then
                elq = "armas"
            Else
                elq = "armas" & SavePath

            End If

        Case e_EstadoIndexador.Alas

            If SavePath = 0 Then
                elq = "alas"
            Else
                elq = "alas" & SavePath

            End If

        Case e_EstadoIndexador.Capas

            If SavePath = 0 Then
                elq = "capas"
            Else
                elq = "capas" & SavePath

            End If

        Case e_EstadoIndexador.Fx

            If SavePath = 0 Then
                elq = "fxs"
            Else
                elq = "fxs" & SavePath

            End If

        Case e_EstadoIndexador.Botas

            If SavePath = 0 Then
                elq = "bota"
            Else
                elq = "bota" & SavePath

            End If

        Case e_EstadoIndexador.Resource
            elq = ""

    End Select

    StringGuardadoActual = elq

End Function

Private Sub CambiarcaptionCommand10()
    Command10.Caption = "Guardar " & StringGuardadoActual(EstadoIndexador)
    MenuArchivoGuardar.Caption = "Guardar " & StringGuardadoActual(EstadoIndexador)
    MenuArchivoCargar.Caption = "Cargar " & StringGuardadoActual(EstadoIndexador)

    If EstadoIndexador = e_EstadoIndexador.Escudos Or EstadoIndexador = e_EstadoIndexador.Armas Then
        Command10.Caption = Command10.Caption & ".dat"
        MenuArchivoGuardar.Caption = MenuArchivoGuardar.Caption & ".dat"
        MenuArchivoCargar.Caption = MenuArchivoCargar.Caption & ".dat"
    Else
        Command10.Caption = Command10.Caption & ".ind"
        MenuArchivoGuardar.Caption = MenuArchivoGuardar.Caption & ".ind"
        MenuArchivoCargar.Caption = MenuArchivoCargar.Caption & ".ind"

    End If

End Sub

Private Sub Command1_Click()
    DibujarFondo = Not DibujarFondo
    Call ClickEnLista

End Sub

Private Sub Command10_Click()
    'Boton de guardado en disco
    Call BotonGuardado

End Sub

Private Sub Command4_Click()

    On Error GoTo ErrHandler

    Call CargarTips

    Dim n As Integer, i As Integer

    n = FreeFile

    Open App.Path & "\" & CarpetaDeInis & "\Tips.ayu" For Binary As #n
    'Escribimos la cabecera
    Put #n, , MiCabecera
    'Guardamos las cabezas
    Put #n, , NumTips

    For i = 1 To NumTips
        Put #n, , Tips(i)
    Next i

    Close #n
    Call MsgBox("Listo, encode ok!!")

    Exit Sub
ErrHandler:
    Call MsgBox("Error en tip " & i)

End Sub

Private Sub EncodeMapas_Click()

    On Error GoTo ErrHandler

    Call CargarMapas

    Dim n As Integer, i As Integer

    n = FreeFile
    Open App.Path & "\" & CarpetaDeInis & "\FK.ind" For Binary As #n
    'Escribimos la cabecera
    Put #n, , MiCabecera
    'Guardamos las cabezas
    Put #n, , NumMapas

    For i = 1 To NumMapas
        Put #n, , Mapas(i)
    Next i

    Close #n

    Call MsgBox("Listo, encode ok!!")

    Exit Sub

ErrHandler:
    Call MsgBox("Error en Mapas " & i)

End Sub

Private Sub BotonGuardar_Click()

    'boton de guardado en memoria
    Dim i As Long

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh
            'guardando un grafico
        
            If GRHActual = 0 Then Exit Sub
    
            If Val(TextDatos(2).Text) <= 0 Then ' numframes = 0
                MsgBox "numero de frames incorrecto"
                Exit Sub

            End If
    
            If Val(TextDatos(2).Text) = 1 Then ' si no es animacion se comprueba si existe el BMP
                If ExisteBMP(Val(TextDatos(0).Text)) = ResourceFile Or (ResourceFile And ExisteBMP(Val(TextDatos(0).Text)) > 0) Then
                Else
                    LUlitError.Caption = "No existe el archivo del grafico"
                    Exit Sub

                End If

            End If
        
            Grhdata(GRHActual).FileNum = Val(TextDatos(0).Text)
            Grhdata(GRHActual).NumFrames = Val(TextDatos(2).Text)

            If Grhdata(GRHActual).NumFrames = 1 Then
                Grhdata(GRHActual).Frames(1) = GRHActual
                Grhdata(GRHActual).pixelHeight = Val(TextDatos(3).Text)
                Grhdata(GRHActual).pixelWidth = Val(TextDatos(4).Text)
                Grhdata(GRHActual).Speed = Val(TextDatos(5).Text)
                Grhdata(GRHActual).sX = Val(TextDatos(6).Text)
                Grhdata(GRHActual).sY = Val(TextDatos(7).Text)
        
                Grhdata(GRHActual).TileHeight = Grhdata(GRHActual).pixelHeight / TilePixelHeight
                Grhdata(GRHActual).TileWidth = Grhdata(GRHActual).pixelWidth / TilePixelWidth
            Else

                For i = 1 To Grhdata(GRHActual).NumFrames

                    If Val(ReadField(i, TextDatos(1).Text, Asc("-"))) < MAXGrH Then
                        Grhdata(GRHActual).Frames(i) = Val(ReadField(i, TextDatos(1).Text, Asc("-")))

                    End If

                Next i

                Grhdata(GRHActual).Speed = Val(TextDatos(5).Text)

                If Grhdata(GRHActual).Frames(1) > 0 Then
                    Grhdata(GRHActual).pixelHeight = Grhdata(Grhdata(GRHActual).Frames(1)).pixelHeight
                    Grhdata(GRHActual).pixelWidth = Grhdata(Grhdata(GRHActual).Frames(1)).pixelWidth
                    Grhdata(GRHActual).sX = Grhdata(Grhdata(GRHActual).Frames(1)).sX
                    Grhdata(GRHActual).sY = Grhdata(Grhdata(GRHActual).Frames(1)).sY
                    Grhdata(GRHActual).TileHeight = Grhdata(Grhdata(GRHActual).Frames(1)).TileHeight
                    Grhdata(GRHActual).TileWidth = Grhdata(Grhdata(GRHActual).Frames(1)).TileWidth
                Else
                    Grhdata(GRHActual).pixelHeight = Val(TextDatos(3).Text)
                    Grhdata(GRHActual).pixelWidth = Val(TextDatos(4).Text)
                    Grhdata(GRHActual).sX = Val(TextDatos(6).Text)
                    Grhdata(GRHActual).sY = Val(TextDatos(7).Text)
                    Grhdata(GRHActual).TileHeight = Grhdata(GRHActual).pixelHeight / TilePixelHeight
                    Grhdata(GRHActual).TileWidth = Grhdata(GRHActual).pixelWidth / TilePixelWidth

                End If

            End If
        
            Call GetInfoGHR(GRHActual)
            frmMain.Visor.Cls
            Call DibujarGHRVisor(GRHActual)

        Case e_EstadoIndexador.Body
            BodyData(DataIndexActual).HeadOffset.Y = Val(ReadField(1, TextDatos(5).Text, Asc("º")))
            BodyData(DataIndexActual).HeadOffset.X = Val(ReadField(2, TextDatos(5).Text, Asc("º")))
            BodyData(DataIndexActual).Walk(1).GrhIndex = Val(TextDatos(0).Text)
            BodyData(DataIndexActual).Walk(2).GrhIndex = Val(TextDatos(2).Text)
            BodyData(DataIndexActual).Walk(3).GrhIndex = Val(TextDatos(3).Text)
            BodyData(DataIndexActual).Walk(4).GrhIndex = Val(TextDatos(4).Text)

        Case e_EstadoIndexador.Cabezas
            HeadData(DataIndexActual).Head(1).GrhIndex = Val(TextDatos(0).Text)
            HeadData(DataIndexActual).Head(2).GrhIndex = Val(TextDatos(2).Text)
            HeadData(DataIndexActual).Head(3).GrhIndex = Val(TextDatos(3).Text)
            HeadData(DataIndexActual).Head(4).GrhIndex = Val(TextDatos(4).Text)

        Case e_EstadoIndexador.Cascos
            CascoAnimData(DataIndexActual).Head(1).GrhIndex = Val(TextDatos(0).Text)
            CascoAnimData(DataIndexActual).Head(2).GrhIndex = Val(TextDatos(2).Text)
            CascoAnimData(DataIndexActual).Head(3).GrhIndex = Val(TextDatos(3).Text)
            CascoAnimData(DataIndexActual).Head(4).GrhIndex = Val(TextDatos(4).Text)

        Case e_EstadoIndexador.Armas
            WeaponAnimData(DataIndexActual).WeaponWalk(1).GrhIndex = Val(TextDatos(0).Text)
            WeaponAnimData(DataIndexActual).WeaponWalk(2).GrhIndex = Val(TextDatos(2).Text)
            WeaponAnimData(DataIndexActual).WeaponWalk(3).GrhIndex = Val(TextDatos(3).Text)
            WeaponAnimData(DataIndexActual).WeaponWalk(4).GrhIndex = Val(TextDatos(4).Text)

        Case e_EstadoIndexador.Escudos
            ShieldAnimData(DataIndexActual).ShieldWalk(1).GrhIndex = Val(TextDatos(0).Text)
            ShieldAnimData(DataIndexActual).ShieldWalk(2).GrhIndex = Val(TextDatos(2).Text)
            ShieldAnimData(DataIndexActual).ShieldWalk(3).GrhIndex = Val(TextDatos(3).Text)
            ShieldAnimData(DataIndexActual).ShieldWalk(4).GrhIndex = Val(TextDatos(4).Text)

        Case e_EstadoIndexador.Botas
            BotaData(DataIndexActual).HeadOffset.Y = Val(ReadField(1, TextDatos(5).Text, Asc("º")))
            BotaData(DataIndexActual).HeadOffset.X = Val(ReadField(2, TextDatos(5).Text, Asc("º")))
            BotaData(DataIndexActual).Bota(1).GrhIndex = Val(TextDatos(0).Text)
            BotaData(DataIndexActual).Bota(2).GrhIndex = Val(TextDatos(2).Text)
            BotaData(DataIndexActual).Bota(3).GrhIndex = Val(TextDatos(3).Text)
            BotaData(DataIndexActual).Bota(4).GrhIndex = Val(TextDatos(4).Text)

        Case e_EstadoIndexador.Capas
            EspaldaAnimData(DataIndexActual).Head(1).GrhIndex = Val(TextDatos(0).Text)
            EspaldaAnimData(DataIndexActual).Head(2).GrhIndex = Val(TextDatos(2).Text)
            EspaldaAnimData(DataIndexActual).Head(3).GrhIndex = Val(TextDatos(3).Text)
            EspaldaAnimData(DataIndexActual).Head(4).GrhIndex = Val(TextDatos(4).Text)

        Case e_EstadoIndexador.Fx
            FxData(DataIndexActual).Fx.GrhIndex = Val(TextDatos(0).Text)
            FxData(DataIndexActual).OffsetX = Val(ReadField(2, TextDatos(2).Text, Asc("º")))
            FxData(DataIndexActual).OffsetY = Val(ReadField(1, TextDatos(2).Text, Asc("º")))

        Case e_EstadoIndexador.Alas
            AlaData(DataIndexActual).HeadOffset.Y = Val(ReadField(1, TextDatos(5).Text, Asc("º")))
            AlaData(DataIndexActual).HeadOffset.X = Val(ReadField(2, TextDatos(5).Text, Asc("º")))
            AlaData(DataIndexActual).Alas(1).GrhIndex = Val(TextDatos(0).Text)
            AlaData(DataIndexActual).Alas(2).GrhIndex = Val(TextDatos(2).Text)
            AlaData(DataIndexActual).Alas(3).GrhIndex = Val(TextDatos(3).Text)
            AlaData(DataIndexActual).Alas(4).GrhIndex = Val(TextDatos(4).Text)

    End Select

    If EstadoIndexador <> e_EstadoIndexador.Grh Then
        Call GetInfoDataIndex(DataIndexActual)

    End If

    EstadoNoGuardado(EstadoIndexador) = True

End Sub

Private Sub Dibujado_Timer()

    On Error Resume Next

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh

            If Not GrHCambiando Then
                If GRHActual <= 0 Then Exit Sub
                If LOOPActual > Grhdata(GRHActual).NumFrames Then LOOPActual = 1
                Call DibujarGHRVisor(Grhdata(GRHActual).Frames(LOOPActual))
                LOOPActual = LOOPActual + 1
            Else

                If LOOPActual > TempGrh.NumFrames Then LOOPActual = 1
                Call DibujarTempGHRVisor(TempGrh.Frames(LOOPActual))
                LOOPActual = LOOPActual + 1

            End If

        Case e_EstadoIndexador.Resource

        Case Else

            If DataIndexActual <= 0 Then Exit Sub
            If tempDataIndex.Walk(1).GrhIndex = 0 Then Exit Sub
            If LOOPActual > Grhdata(tempDataIndex.Walk(1).GrhIndex).NumFrames Then LOOPActual = 1
            Call DibujarDataIndex(tempDataIndex, LOOPActual, DibujarWalk)
            LOOPActual = LOOPActual + 1

    End Select

End Sub

Private Sub Form_close()
    End

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim i          As Long

    Dim HayCambios As Boolean

    Dim respuesta  As Byte

    Dim Tstr       As String

Inicio:
    HayCambios = False

    For i = e_EstadoIndexador.Grh To Resource

        If EstadoNoGuardado(i) Then
            HayCambios = True
            Tstr = Tstr & StringGuardadoActual(i) & vbCrLf

        End If

    Next i

    If HayCambios Then
        respuesta = MsgBox("Hay cambios sin Guardar en :" & vbCrLf & vbCrLf & Tstr & vbCrLf & "¿Quieres GUARDAR los cambios antes de salir?" & vbCrLf & "(Si pulsas NO se perderan estos cambios)" & vbCrLf, 3, "Aviso")

        If respuesta = vbCancel Then
            Cancel = 1 ' cancelamos la salida
            Exit Sub
        ElseIf respuesta = vbYes Then

            For i = e_EstadoIndexador.Grh To e_EstadoIndexador.Resource ' guardamos

                If EstadoNoGuardado(i) Then
                    EstadoIndexador = i
                    Call BotonGuardado

                End If

            Next i

            Tstr = vbNullString
            GoTo Inicio ' weno el goto es el alien d la programacion estructurada pero paso d romperme la cabeza xD asi se ve mejor

            ' volvemos a comprobar si algo no se guardo
        End If
        
    End If

End Sub

Private Sub Form_resize()
    Visor.Height = Abs(frmMain.Height - Visor.Top - LUlitError.Height - 810)
    Visor.Width = Abs(frmMain.Width - Visor.Left - 120)
    LUlitError.Top = Abs(frmMain.Height - LUlitError.Height - 705)
    LUlitError.Width = Abs(frmMain.Width - 155)
    Call ClickEnLista

End Sub

Private Sub Form_Load()
    'indices.ini
    ReferenciaPrincipal = 1
    ReferenciaDos = 0
    Salto = False
    'configuracion inicial:
    SavePath = 0
    LoadingNew = False ' variable que evita redibujado excesibo
    IniciadoTodo = True
    ColorFondo = vbGreen
    CarpetaDeInis = GetVar(App.Path & "\Datos\Conf.ini", "Config", "CarpetaDeInis")
    CarpetaGraficos = GetVar(App.Path & "\Datos\Conf.ini", "Config", "CarpetaGraficos")
    CarpetaDedatos = GetVar(App.Path & "\Datos\Conf.ini", "Config", "CarpetaDedatos")
    PathDat = App.Path & "\" & CarpetaDedatos & "\"
   
    ResourceFile = 1 ' siempre cargamos lo bmps, esta deshabilitado el archivo de recursos.
    GrhTex.Text = MAXGrH

    If ResourceFile <= 0 Then ResourceFile = 1
    If CarpetaDeInis = vbNullString Then CarpetaDeInis = "INIT"
    If CarpetaDedatos = vbNullString Then CarpetaDedatos = "Datos"
    If CarpetaGraficos = vbNullString Then CarpetaGraficos = "graficos"
    
    txtIni.Text = CarpetaDeInis
    txtGraficos.Text = CarpetaGraficos
    Call IniciarCabecera(MiCabecera)
    
    Call IniciarObjetosDirectX
    Set SurfaceDB = New clsSurfaceManDyn
    Call InitTileEngine(frmMain.hwnd, 155, 16, 32, 32, 13, 17, 9)
    
    Call CargarAnimsExtra
    Call CargarTips
    
    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos

    EstadoIndexador = e_EstadoIndexador.Grh

    Dim Lister As Long

    For Lister = 0 To 9
        MenuIndiceGuardado(Lister).Checked = False
    Next Lister

    MenuIndiceGuardado(0).Checked = True
    Call CambiarcaptionCommand10
    Call Existeindices

End Sub

Private Sub CargarMapas()

    Dim loopc As Integer

    NumMapas = Val(GetVar(App.Path & "\" & CarpetaDeInis & "\encode\mapas.dat", "INIT", "NumMaps"))

    ReDim Mapas(0 To NumMapas + 1) As Byte

    For loopc = 1 To NumMapas
        Mapas(loopc) = Val(GetVar(App.Path & "\" & CarpetaDeInis & "\encode\mapas.dat", "Map" & loopc, "Lluvia"))
    Next loopc

End Sub

Private Sub CargarTips()

    Dim loopc As Integer

    NumTips = Val(GetVar(App.Path & "\encode\tips.dat", "INIT", "Tips"))

    ReDim Tips(0 To NumTips + 1) As String * 255

    For loopc = 1 To NumTips
        Tips(loopc) = GetVar(App.Path & "\encode\tips.dat", "Tip" & loopc, "Tip")
    Next loopc

End Sub

Private Sub IAnim_Click()

    Dim textoActual As String

    Dim bmpstring   As String

    Dim BMPBuscado  As Long

    bmpstring = ReadField(1, Lista.List(Lista.listIndex), Asc(" "))

    If IsNumeric(bmpstring) Then
        BMPBuscado = Val(bmpstring)

        If BMPBuscado > 0 And BMPBuscado <= MAXGrH Then
            If FormAuto.Visible Then
                FormAuto.SetFocus
            Else
                FormAuto.Show , frmMain

            End If

            FormAuto.FrameAnim(0).Visible = True
            FormAuto.FrameAnim(1).Visible = False
            FormAuto.FrameAnim(2).Visible = False
            FormAuto.TextDatos(4).Text = BMPBuscado

        End If

    End If

End Sub

Private Sub Lista_Click()
    Call ClickEnLista

End Sub

Public Sub ClickEnLista()

    On Error Resume Next

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh
            GRHActual = Val(ReadField(1, Lista.List(Lista.listIndex), Asc(" ")))
            LOOPActual = 1
            Call GetInfoGHR(GRHActual)
            Call DibujarGHRVisor(GRHActual)
            LGHRnumeroA.Caption = GRHActual

        Case e_EstadoIndexador.Resource
            GRHActual = Val(Lista.List(Lista.listIndex))
            frmMain.Visor.Cls

            If ExisteBMP(GRHActual) = 0 Then Exit Sub
            Call GetInfoBmp(GRHActual)
            Call DibujarBMPVisor(GRHActual)
            LGHRnumeroA.Caption = GRHActual

        Case Else
            frmMain.Visor.Cls
            DataIndexActual = Val(Lista.List(Lista.listIndex))
            LOOPActual = 1
            Call GetInfoDataIndex(DataIndexActual)
        
            LGHRnumeroA.Caption = DataIndexActual

    End Select

    UltimoindexE(EstadoIndexador) = Lista.listIndex

End Sub

Private Sub Lista_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton And EstadoIndexador = Resource Then
        Call Me.PopupMenu(Me.mnuautoI)

    End If

End Sub

Private Sub MenuAcercaDe_Click()
    MsgBox "Indexador Creado por columdrum" & vbCrLf & vbCrLf & "Email: columdrum@gmail.com" & vbCrLf & vbCrLf & "Version: " & VERSION_ACTUAL

End Sub

Private Sub MenuArchivoCargar_Click()
    Call BotonCargado

End Sub

Private Sub MenuArchivoGuardar_Click()
    Call BotonGuardado

End Sub

Private Sub MenuBotonCargarP_Click()

    On Error GoTo Cancelar

    With CommonDialog1
        .Filter = "Binario(.ind)|*.ind|Archivo de texto DAT(*.dat)|*.dat"

        If EstadoIndexador = e_EstadoIndexador.Armas Or EstadoIndexador = e_EstadoIndexador.Escudos Then
            .Filter = "Archivo de texto DAT(*.dat)|*.dat"

        End If

        .CancelError = True
        .flags = cdlOFNFileMustExist
        .FileName = StringGuardadoActual(EstadoIndexador)
        .ShowOpen

    End With
 
    Select Case UCase(Right(CommonDialog1.FileName, 3))

        Case "DAT"
            Call BotonCargadoDat(CommonDialog1.FileName)

        Case "IND"
            Call BotonCargado(CommonDialog1.FileName)

        Case Else
            Exit Sub

    End Select

    Exit Sub
Cancelar:

End Sub

Private Sub MenuBotonGuardarP_Click()

    On Error GoTo Cancelar

    With CommonDialog1
        .Filter = "Binario(.ind)|*.ind|Archivo de texto DAT(*.dat)|*.dat"

        If EstadoIndexador = e_EstadoIndexador.Armas Or EstadoIndexador = e_EstadoIndexador.Escudos Then
            .Filter = "Archivo de texto DAT(*.dat)|*.dat"

        End If

        .CancelError = True
        .flags = cdlOFNOverwritePrompt
        .FileName = StringGuardadoActual(EstadoIndexador)
        .ShowSave

    End With
 
    Select Case UCase(Right(CommonDialog1.FileName, 3))

        Case "DAT"
            Call BotonGuardadoDat(CommonDialog1.FileName)

        Case "IND"
            Call BotonGuardado(CommonDialog1.FileName)

        Case Else
            Exit Sub

    End Select

    Exit Sub
Cancelar:

End Sub

Private Sub MenuEdicionBorrar_Click()
    Call SBotonBorrrar

End Sub

Private Sub MenuEdicionClonar_Click()
    Call SbotonClonar

End Sub

Private Sub menuEdicionColor_Click()

    With CommonDialog1
        .DialogTitle = "Seleccionar color para el fondo"
        .ShowColor

    End With

    ColorFondo = CommonDialog1.Color
    Call ClickEnLista

End Sub

Private Sub MenuEdicionCopiar_Click()
    Call SBotonMover(False)

End Sub

Private Sub menuEdicionMover_Click()
    Call SBotonMover(True)

End Sub

Public Sub SbotonClonar()

    Dim tempLong  As Long

    Dim cadena    As String

    Dim respuesta As Byte

    Dim LooPer    As Long

    Dim Inicial   As Long

    Dim Final     As Long

    Dim Origen    As Long

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh

            If GrHCambiando Then
                GrHCambiando = False
                LNumActual.Caption = "Ghr:"
                BotonGuardar.Visible = False

            End If

            If GRHActual < 0 Or GRHActual > MAXGrH Then Exit Sub
            cadena = InputBox("Introduzca el primer Numero de GHR al que quieres mover el grafico " & GRHActual, "Clonar Grafico")

            If IsNumeric(cadena) Then
                Inicial = Val(cadena)

                If Inicial > 0 And Inicial < MAXGrH Then
                    cadena = InputBox("Introduzca Cantidad de veces que quieres clonar el grafico " & GRHActual & " a partir de la posicion: " & Inicial, "Clonar Grafico")

                    If IsNumeric(cadena) Then
                        Final = Val(cadena) + Inicial

                        If Final > 0 And Final < MAXGrH Then
                            Origen = GRHActual

                            For LooPer = Inicial To Final
                                Call MoverGrh(LooPer, Origen, False)
                            Next LooPer

                            Call RenuevaListaGrH
                            tempLong = ListaindexGrH(Inicial)
                            frmMain.Lista.listIndex = tempLong
                            EstadoNoGuardado(e_EstadoIndexador.Grh) = True
                        Else
                            MsgBox "Fuera de los limites"

                        End If

                    End If

                Else
                    MsgBox "numero incorrecto"

                End If

            End If

            '    Case Else
            '        Dim StringCaso As String
            '        If EstadoIndexador = Body Then
            '            StringCaso = "Body"
            '        ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
            '            StringCaso = "Cabeza"
            '        ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
            '            StringCaso = "Casco"
            '        ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
            '            StringCaso = "Escudo"
            '        ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
            '            StringCaso = "Arma"
            '        ElseIf EstadoIndexador = e_EstadoIndexador.Botas Then
            '            StringCaso = "Bota"
            '        ElseIf EstadoIndexador = e_EstadoIndexador.Capas Then
            '            StringCaso = "Capa"
            '        ElseIf EstadoIndexador = e_EstadoIndexador.Fx Then
            '            StringCaso = "Fx"
            '        End If
            '        cadena = InputBox("Introduzca numero de " & StringCaso & " al que quieres mover", "Mover " & StringCaso)
            '        If IsNumeric(cadena) And (Val(cadena) < 31000) Then
            '            If EstadoIndexador = e_EstadoIndexador.Body Then
            '                Call mueveBody(Val(cadena), DataIndexActual, False)
            '            ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
            '                Call MueveCabeza(Val(cadena), DataIndexActual, False)
            '            ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
            '                Call MueveCasco(Val(cadena), DataIndexActual, False)
            '            ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
            '                Call MueveEscudo(Val(cadena), DataIndexActual, False)
            '            ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
            '                Call MueveArma(Val(cadena), DataIndexActual, False)
            '            ElseIf EstadoIndexador = e_EstadoIndexador.Botas Then
            '                Call MueveBota(Val(cadena), DataIndexActual, False)
            '            ElseIf EstadoIndexador = e_EstadoIndexador.Capas Then
            '                Call MueveCapa(Val(cadena), DataIndexActual, False)
            '            ElseIf EstadoIndexador = e_EstadoIndexador.Fx Then
            '                Call MueveFX(Val(cadena), DataIndexActual, False)
            '            End If
            '                DataIndexActual = Val(cadena)
            '                LOOPActual = 1
            '                frmMain.Visor.Cls
            '                Call GetInfoDataIndex(DataIndexActual)
            '                Dibujado.Interval = 200
            '                Dibujado.Enabled = True
            '                LGHRnumeroA.Caption = DataIndexActual
            '                templong = ListaindexGrH(DataIndexActual)
            '                frmMain.Lista.listIndex = templong
            '        Else
            '            MsgBox "introduzca un numero valido"
            '        End If
    End Select

End Sub

Private Sub MenuEdicionNuevo_Click()
    Call BotonNuevoGRH

End Sub

Private Sub MenuHerramientasAAnim_Click()

    If FormAuto.Visible Then
        FormAuto.SetFocus
    Else
        FormAuto.Show , frmMain

    End If

End Sub

Private Sub MenuHerramientasBG_Click()

    Dim i      As Long

    Dim cadena As String

    LastFound = 0
    cadena = InputBox("Introduzca número de Bmp a buscar", "Nuevo Grafico")

    If Val(cadena) > 0 And Val(cadena) <= MAXGrH Then
        BMPBuscado = Val(cadena)

        For i = 1 To MAXGrH

            If Grhdata(i).FileNum = BMPBuscado Then
                Call BuscarNuevoF(i)
                LastFound = i
                MenuHerramientasBN.Visible = True
                LUlitError.Caption = "F3 para continuar la busqueda"
                Exit Sub

            End If

        Next i

        LUlitError.Caption = "BMP no encontrado"
        MenuHerramientasBN.Visible = False

    End If

End Sub

Private Sub MenuHerramientasBN_Click()

    Dim i As Long

    If LastFound = 0 Or BMPBuscado = 0 Then Exit Sub

    For i = LastFound + 1 To MAXGrH

        If Grhdata(i).FileNum = BMPBuscado Then
            Call BuscarNuevoF(i)
            LastFound = i
            LUlitError.Caption = "F3 para continuar la busqueda"
            Exit Sub

        End If

    Next i

    LUlitError.Caption = " Se termino la busqueda"
    MenuHerramientasBN.Visible = False
    LastFound = 0
    BMPBuscado = 0

End Sub

Private Sub MenuHerramientasBR_Click()

    If FrmSearch.Visible Then
        FrmSearch.SetFocus
    Else
        FrmSearch.Show , frmMain

    End If

    Call FrmSearch.HacerBusquedaR

End Sub

Private Sub MenuHerramientasNI_Click()

    If FrmSearch.Visible Then
        FrmSearch.SetFocus
    Else
        FrmSearch.Show , frmMain

    End If

    Call FrmSearch.HacerBusquedaNI

End Sub

Private Sub MenuHerramientasWE_Click()
    If FraDatosIndice.Visible = True Then
        ChaGenerarIndices.Visible = False
        lblDescripcion.Visible = False
        txtNombre.Visible = False
        FraDatosIndice.Visible = False
        EncodeMapas.Visible = True
        BotonI(9).Visible = True
        CmbResourceFile.Visible = True
        FraCarpetas.Visible = True
        Command10.Visible = True
        BotonBorrrar.Visible = True
        Command1.Visible = True
        txtNombre.Text = ""
        txtAlto.Text = ""
        txtAncho.Text = ""
        TxtGrhIndex.Text = ""
        TxtReferencias.Text = ""
    Else
        ChaGenerarIndices.Visible = True
        lblDescripcion.Visible = True
        txtNombre.Visible = True
        FraDatosIndice.Visible = True
        EncodeMapas.Visible = False
        BotonI(9).Visible = False
        CmbResourceFile.Visible = False
        TxtReferencias.Text = MaxSup
        FraCarpetas.Visible = False
        Command10.Visible = False
        BotonBorrrar.Visible = False
        Command1.Visible = False
    End If

End Sub

Private Sub MenuIndiceGuardado_Click(Index As Integer)

    Dim i As Long

    For i = 0 To 9
        MenuIndiceGuardado(i).Checked = False
    Next i

    MenuIndiceGuardado(Index).Checked = True
    SavePath = Index
    Call CambiarcaptionCommand10

End Sub

Private Sub mnIgeneral_Click()

    Dim textoActual As String

    Dim bmpstring   As String

    Dim BMPBuscado  As Long

    bmpstring = ReadField(1, Lista.List(Lista.listIndex), Asc(" "))

    If IsNumeric(bmpstring) Then
        BMPBuscado = Val(bmpstring)

        If BMPBuscado > 0 And BMPBuscado <= MAXGrH Then
            If FormAuto.Visible Then
                FormAuto.SetFocus
            Else
                FormAuto.Show , frmMain

            End If

            FormAuto.FrameAnim(0).Visible = False
            FormAuto.FrameAnim(1).Visible = False
            FormAuto.FrameAnim(2).Visible = True
            FormAuto.TextDatos3(4).Text = BMPBuscado
            FormAuto.TextDatos3(5).Text = BuscarGrHlibres(1)

        End If

    End If

End Sub

Private Sub mnuibody_Click()

    Dim textoActual As String

    Dim bmpstring   As String

    Dim BMPBuscado  As Long

    bmpstring = ReadField(1, Lista.List(Lista.listIndex), Asc(" "))

    If IsNumeric(bmpstring) Then
        BMPBuscado = Val(bmpstring)

        If BMPBuscado > 0 And BMPBuscado <= MAXGrH Then
            If FormAuto.Visible Then
                FormAuto.SetFocus
            Else
                FormAuto.Show , frmMain

            End If

            FormAuto.FrameAnim(1).Visible = True
            FormAuto.FrameAnim(0).Visible = False
            FormAuto.FrameAnim(2).Visible = False
            FormAuto.Combo2.Visible = False
            FormAuto.Labelbody.Visible = False
            FormAuto.Labelbody1.Visible = False
            FormAuto.Labelbody2.Visible = False
            FormAuto.Loff.Visible = False
            FormAuto.Loffx.Visible = False
            FormAuto.Loffy.Visible = False
            FormAuto.TextDatos2(7).Visible = False
            FormAuto.TextDatos2(8).Visible = False
            FormAuto.TextDatos2(0).Enabled = False
            FormAuto.TextDatos2(1).Enabled = False
            FormAuto.TextDatos2(6).Enabled = False
            FormAuto.Text1.Visible = False
            FormAuto.Text2.Visible = False
            FormAuto.CheckAuto.Visible = False
            FormAuto.Optiondimension(0).Visible = False
            FormAuto.Optiondimension(1).Visible = False
            FormAuto.Optiondimension(2).Visible = False
            FormAuto.Label5.Visible = False
            FormAuto.Label6.Visible = False
            FormAuto.TextDatos2(0).Text = 6
            FormAuto.TextDatos2(1).Text = 4
            FormAuto.TextDatos2(4).Text = BMPBuscado
            FormAuto.TextDatos2(6).Text = 22
            FormAuto.TextDatos2(2).Text = 46
            FormAuto.TextDatos2(3).Text = 26
            FormAuto.Optiondimension(0).value = True
            FormAuto.Labelbody.Visible = True
            FormAuto.Labelbody1.Visible = True
            FormAuto.Labelbody2.Visible = True
            FormAuto.Text1.Visible = True
            FormAuto.Text1.Enabled = False
            FormAuto.Text2.Visible = True
            FormAuto.Text2.Enabled = False
            FormAuto.CheckAuto.Visible = True
            FormAuto.CheckAuto.value = vbUnchecked
            FormAuto.Text1.Text = UBound(BodyData) + 1
            FormAuto.Text2.Text = "-38º0"
            FormAuto.Combo2.Visible = True
            FormAuto.Combo2.listIndex = 0
            FormAuto.Optiondimension(0).Visible = True
            FormAuto.Optiondimension(1).Visible = True
            FormAuto.Optiondimension(2).Visible = True
            FormAuto.Label5.Visible = True
            FormAuto.Label6.Visible = True

        End If

    End If

End Sub

Private Sub NuevoGhr_Click()
    Call BotonNuevoGRH

End Sub

Public Sub BotonNuevoGRH()

    Dim cadena    As String

    Dim respuesta As Byte

    Dim tempLong  As Long

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh

            If GrHCambiando Then
                GrHCambiando = False
                LNumActual.Caption = "Ghr:"
                BotonGuardar.Visible = False

            End If

            cadena = InputBox("Introduzca el número de GHR (0 Para encontrar un hueco libre)", "Nuevo Grafico")

            If IsNumeric(cadena) Then
                Call BuscarNuevoF(Val(cadena))
            Else
                LUlitError.Caption = "introduzca un numero"

            End If

        Case e_EstadoIndexador.Resource
            cadena = InputBox("Introduzca el número de BMP", "Nuevo Grafico")

            If IsNumeric(cadena) Then
                Call BuscarNuevoF(Val(cadena))
            Else
                LUlitError.Caption = "introduzca un numero"

            End If

            Exit Sub

        Case Else

            Dim StringCaso As String

            If EstadoIndexador = Body Then
                StringCaso = "Body"
            ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
                StringCaso = "Cabeza"
            ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
                StringCaso = "Casco"
            ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
                StringCaso = "Escudo"
            ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
                StringCaso = "Arma"
            ElseIf EstadoIndexador = e_EstadoIndexador.Alas Then
                StringCaso = "Alas"
            ElseIf EstadoIndexador = e_EstadoIndexador.Capas Then
                StringCaso = "Capa"
            ElseIf EstadoIndexador = e_EstadoIndexador.Fx Then
                StringCaso = "Fx"
            ElseIf EstadoIndexador = e_EstadoIndexador.Botas Then
                StringCaso = "Bota"

            End If
        
            cadena = InputBox("Introduzca " & StringCaso & " (0 Para encontrar un hueco libre)", "Nuevo " & StringCaso & "/buscar")

            If IsNumeric(cadena) Then
                Call BuscarNuevoF(Val(cadena))
            Else
                LUlitError.Caption = "introduzca un numero"

            End If

    End Select

End Sub

Public Sub BuscarNuevoF(ByVal Index As Long)

    Dim cadena    As String

    Dim respuesta As Byte

    Dim tempLong  As Long

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh

            If GrHCambiando Then
                GrHCambiando = False
                LNumActual.Caption = "Ghr:"
                BotonGuardar.Visible = False

            End If

            If Index > 0 And Index < MAXGrH Then
                tempLong = ListaindexGrH(Index)

                If tempLong >= 0 Then
                    GRHActual = Index
                    LOOPActual = 1
                    frmMain.Visor.Cls
                    'Call DibujarGHRVisor(GRHActual)
                    'Call GetInfoGHR(GRHActual)
                    LGHRnumeroA.Caption = GRHActual
                    frmMain.Lista.listIndex = tempLong
                Else
                    respuesta = MsgBox("El grafico no existe,¿ quieres crearlo? ", 4, "Aviso")

                    If respuesta = vbYes Then
                        GRHActual = Index
                        'GrhData(GRHActual).NumFrames = 0
                        Call AgregaGrH(GRHActual)
                        LOOPActual = 1
                        frmMain.Visor.Cls
                        'Call DibujarGHRVisor(GRHActual)
                        Call GetInfoGHR(GRHActual)
                        LGHRnumeroA.Caption = GRHActual
                        EstadoNoGuardado(e_EstadoIndexador.Grh) = True

                    End If

                End If

            ElseIf Index = 0 Then
                GRHActual = BuscarGrHlibre()

                If GRHActual > 0 And GRHActual <= MAXGrH Then
                    Call AgregaGrH(GRHActual)
                    LOOPActual = 1
                    frmMain.Visor.Cls
                    'Call DibujarGHRVisor(GRHActual)
                    Call GetInfoGHR(GRHActual)
                    LGHRnumeroA.Caption = GRHActual
                Else
                    LUlitError.Caption = "No Se encontro hueco"

                End If

            Else
                LUlitError.Caption = "Valor no valido"

            End If

        Case e_EstadoIndexador.Resource

            If Index > 0 And Index <= MAXGrH Then
                tempLong = ListaindexGrH(Index)

                If tempLong >= 0 Then
                    GRHActual = Index
                    LOOPActual = 1
                    frmMain.Visor.Cls
                    'Call DibujarBMPVisor(GRHActual)
                    Call GetInfoBmp(GRHActual)
                    Call DibujarBMPVisor(GRHActual)
                    LGHRnumeroA.Caption = GRHActual
                    frmMain.Lista.listIndex = tempLong
                Else
                    LUlitError.Caption = "Bmp no existe"

                End If

            Else
                LUlitError.Caption = "Valor no valido"

            End If

            Exit Sub

        Case Else

            Dim StringCaso As String

            If EstadoIndexador = Body Then
                StringCaso = "Body"
            ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
                StringCaso = "Cabeza"
            ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
                StringCaso = "Casco"
            ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
                StringCaso = "Escudo"
            ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
                StringCaso = "Arma"
            ElseIf EstadoIndexador = e_EstadoIndexador.Alas Then
                StringCaso = "Alas"
            ElseIf EstadoIndexador = e_EstadoIndexador.Capas Then
                StringCaso = "Capa"
            ElseIf EstadoIndexador = e_EstadoIndexador.Fx Then
                StringCaso = "Fx"
            ElseIf EstadoIndexador = e_EstadoIndexador.Botas Then
                StringCaso = "Bota"

            End If

            If Index > 0 And Index < MAXGrH Then
                tempLong = ListaindexGrH(Index)

                If tempLong >= 0 Then
                    DataIndexActual = Index
                    LOOPActual = 1
                    frmMain.Visor.Cls
                    Call GetInfoDataIndex(DataIndexActual)
                    Dibujado.Interval = 100
                    Dibujado.Enabled = True
                    LGHRnumeroA.Caption = DataIndexActual
                    Lista.listIndex = tempLong
                Else
                    respuesta = MsgBox("El " & StringCaso & " no existe,¿ quieres crearlo? ", 4, "Aviso")

                    If respuesta = vbYes Then
                        DataIndexActual = Index
                        LOOPActual = 1
                        frmMain.Visor.Cls
                        Dibujado.Interval = 100
                        Dibujado.Enabled = True
                        LGHRnumeroA.Caption = DataIndexActual

                        If EstadoIndexador = e_EstadoIndexador.Body Then
                            Call AgregaBody(DataIndexActual)
                        ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
                            Call AgregaCabeza(DataIndexActual)
                        ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
                            Call AgregaCasco(DataIndexActual)
                        ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
                            Call AgregaEscudo(DataIndexActual)
                        ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
                            Call AgregaArma(DataIndexActual)
                        ElseIf EstadoIndexador = e_EstadoIndexador.Alas Then
                            Call AgregaAlas(DataIndexActual)
                        ElseIf EstadoIndexador = e_EstadoIndexador.Capas Then
                            Call AgregaCapa(DataIndexActual)
                        ElseIf EstadoIndexador = e_EstadoIndexador.Fx Then
                            Call AgregaFx(DataIndexActual)
                        ElseIf EstadoIndexador = e_EstadoIndexador.Botas Then
                            Call AgregaBota(DataIndexActual)

                        End If

                        Call GetInfoDataIndex(DataIndexActual)
                        EstadoNoGuardado(EstadoIndexador) = True

                    End If

                End If

                '            ElseIf Val(cadena) = 0 Then
                '                GRHActual = BuscarGrHlibre()
                '                If GRHActual > 0 And GRHActual <= MAXGrH Then
                '                    Call AgregaGrH(GRHActual)
                '                    LOOPActual = 1
                '                    frmMain.Visor.Cls
                '                    Call DibujarGHRVisor(GRHActual)
                '                    Call GetInfoGHR(GRHActual)
                '                    LGHRnumeroA.Caption = GRHActual
                '                Else
                '                    MsgBox "No Se encontro hueco"
                '                End If
            Else
                LUlitError.Caption = "Valor no valido"

            End If

    End Select

End Sub

Private Sub Text1_Change()
    CarpetaDeInis = Text1.Text
    Call WriteVar(App.Path & "\Datos\Conf.ini", "Config", "CarpetaDeInis", CarpetaDeInis)

End Sub

Private Sub Text2_Change()
    CarpetaGraficos = Text2.Text
    Call WriteVar(App.Path & "\Datos\Conf.ini", "Config", "CarpetaGraficos", Text2.Text)

End Sub

Private Sub TextDatos_DblClick(Index As Integer)

    If EstadoIndexador = e_EstadoIndexador.Grh Or Index > 4 Or (EstadoIndexador = e_EstadoIndexador.Fx) And Index > 0 Then Exit Sub

    If Val(TextDatos(Index).Text) > 0 And Val(TextDatos(Index).Text) < MAXGrH Then

        If EstadoIndexador <> e_EstadoIndexador.Grh Then Call CambiarEstado(e_EstadoIndexador.Grh)
        Call BuscarNuevoF(TextDatos(Index).Text)

    End If

End Sub

Private Sub TextDatos_Change(Index As Integer)
    'Comprueba que los datos introducidos son correctos

    Dim Ancho       As Long

    Dim Alto        As Long

    Dim PrimerAncho As Long

    Dim PrimerAlto  As Long

    Dim i           As Long

    Dim Algun_Error As Boolean

    Dim ErroresGrh  As ErroresGrh

    Dim tdouble1    As Double, tdouble2 As Double

    If EstadoIndexador = e_EstadoIndexador.Resource Then Exit Sub

2   For i = 0 To 7

        If i <> 1 And ((i <> 5) Or EstadoIndexador <> Body) And ((i <> 2) Or EstadoIndexador <> Fx) Then ' el 1 son los frames y el 5 se usa para offset
            If Val(TextDatos(i).Text) > MAXGrH Then
                TextDatos(i).Text = MAXGrH

            End If

        ElseIf ((i = 5) And EstadoIndexador = Body) Or ((i = 2) And EstadoIndexador = Fx) Then
            tdouble1 = Val(ReadField(1, TextDatos(i).Text, Asc("º")))
            tdouble2 = Val(ReadField(2, TextDatos(i).Text, Asc("º")))

            If tdouble1 < -MAXGrH Or tdouble1 > MAXGrH Then
                TextDatos(i).Text = "0º" & tdouble2
                tdouble1 = 0

            End If
        
            If tdouble2 < -MAXGrH Or tdouble2 > MAXGrH Then
                TextDatos(i).Text = tdouble1 & "º0"

            End If

        End If

        ErroresGrh.colores(i) = vbWhite
    Next i

    ErroresGrh.colores(8) = vbWhite
    ErroresGrh.colores(9) = vbWhite

    LUlitError.Caption = ""

    Dim resul        As Long

    Dim MensageError As String

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh

            If Not GrHCambiando Then
                GrHCambiando = True
                TempGrh = Grhdata(GRHActual)
                LNumActual.Caption = "**Ghr:"
                BotonGuardar.Visible = True

            End If

            If Val(TextDatos(5).Text) > MAXGrH Then
                TextDatos(5).Text = MAXGrH

            End If

            If Val(TextDatos(2).Text) > 25 Then 'numframes > 25
                TextDatos(2).Text = 25
            ElseIf Val(TextDatos(2).Text) < 1 Then 'numframes < 1
                TextDatos(2).Text = 1

            End If
            
            If Val(TextDatos(2).Text) = 1 Then ' Es grh normal
                TextDatos(1).Enabled = False

                For i = 3 To 4
                    TextDatos(i).Enabled = True
                Next i

                TextDatos(5).Enabled = False

                For i = 6 To 7
                    TextDatos(i).Enabled = True
                Next i

            ElseIf Val(TextDatos(2).Text) > 1 Then ' es animacion
                TextDatos(1).Enabled = True

                For i = 3 To 4
                    TextDatos(i).Enabled = False
                Next i

                TextDatos(5).Enabled = True

                For i = 6 To 7
                    TextDatos(i).Enabled = False
                Next i

            End If

            TempGrh.FileNum = Val(TextDatos(0).Text)
            TempGrh.NumFrames = Val(TextDatos(2).Text)

            If TempGrh.NumFrames = 1 Then
                TempGrh.Frames(1) = Val(ReadField(1, TextDatos(1).Text, Asc("-")))
            Else

                For i = 1 To TempGrh.NumFrames

                    If Val(ReadField(i, TextDatos(1).Text, Asc("-"))) < MAXGrH Then
                        TempGrh.Frames(i) = Val(ReadField(i, TextDatos(1).Text, Asc("-")))

                    End If

                Next i

            End If

            TempGrh.pixelHeight = Val(TextDatos(3).Text)
            TempGrh.pixelWidth = Val(TextDatos(4).Text)
            TempGrh.Speed = Val(TextDatos(5).Text)
            TempGrh.sX = Val(TextDatos(6).Text)
            TempGrh.sY = Val(TextDatos(7).Text)
            TempGrh.TileHeight = Grhdata(GRHActual).pixelHeight / TilePixelHeight
            TextDatos(8).Text = TempGrh.TileHeight
            TempGrh.TileWidth = Grhdata(GRHActual).pixelWidth / TilePixelWidth
            TextDatos(9).Text = TempGrh.TileWidth

            resul = GrhCorrecto(TempGrh, MensageError, ErroresGrh)
            LUlitError.Caption = MensageError
            
            For i = 0 To 9
                TextDatos(i).BackColor = ErroresGrh.colores(i)
            Next i
            
            If ErroresGrh.ErrorCritico Then
                BotonGuardar.Visible = False
                Exit Sub
            Else
                BotonGuardar.Visible = True

            End If
            
            frmMain.Visor.Cls

            If Not LoadingNew Then Call DibujarGHRVisor(GRHActual)
            If TempGrh.NumFrames = 1 Then
                frmMain.Dibujado.Enabled = False
            ElseIf TempGrh.NumFrames > 1 Then

                If TempGrh.Speed > 0 Then ' pervenimos division por 0
                    frmMain.Dibujado.Interval = 50 * TempGrh.Speed
                Else
                    frmMain.Dibujado.Interval = 100

                End If

                frmMain.Dibujado.Enabled = True
            Else
                frmMain.Dibujado.Enabled = False

            End If

        Case Else

            If Not GrHCambiando Then
                GrHCambiando = True
                BotonGuardar.Visible = True

            End If
        
            If Not LoadingNew Then frmMain.Visor.Cls ' Si no estamos cargando limpiamos
            
            Dibujado.Interval = 100
            Dibujado.Enabled = True

            If EstadoIndexador = e_EstadoIndexador.Body Then
                tempDataIndex.HeadOffset.Y = Val(ReadField(1, TextDatos(5).Text, Asc("º")))
                tempDataIndex.HeadOffset.X = Val(ReadField(2, TextDatos(5).Text, Asc("º")))

            End If

            Dim III  As Long

            Dim Tstr As String

            Algun_Error = False

            For i = 1 To 4

                If i = 1 Then
                    III = 0
                Else
                    III = i

                End If

                If i = 1 Then
                    If EstadoIndexador = e_EstadoIndexador.Fx Then
                        Tstr = "FX"
                    Else
                        Tstr = "Subir"

                    End If

                ElseIf i = 2 Then
                    Tstr = "Derecha"
                ElseIf i = 3 Then
                    Tstr = "Abajo"
                ElseIf i = 4 Then
                    Tstr = "Izquierda"

                End If

                If i = 1 Or EstadoIndexador <> e_EstadoIndexador.Fx Then
                    tempDataIndex.Walk(i).GrhIndex = Val(TextDatos(III).Text)

                    If tempDataIndex.Walk(i).GrhIndex > 1 Then
                        MensageError = ""
                        resul = GrhCorrecto(Grhdata(tempDataIndex.Walk(i).GrhIndex), MensageError, ErroresGrh)

                        If ErroresGrh.ErrorCritico Then
                            Algun_Error = True
                            TextDatos(III).BackColor = vbRed
                            LUlitError.Caption = LUlitError.Caption & "(" & Tstr & ") " & MensageError & vbCrLf
                        Else

                            If EstadoIndexador = e_EstadoIndexador.Cabezas Or EstadoIndexador = e_EstadoIndexador.Cascos Then
                                If ErroresGrh.EsAnimacion Then
                                    TextDatos(III).BackColor = vbYellow
                                    LUlitError.Caption = LUlitError.Caption & "(" & Tstr & ") Es una animacion" & vbCrLf
                                Else
                                    TextDatos(III).BackColor = vbWhite

                                End If

                            Else

                                If Not ErroresGrh.EsAnimacion Then
                                    TextDatos(III).BackColor = vbYellow
                                    LUlitError.Caption = LUlitError.Caption & "(" & Tstr & ") No es una animacion" & vbCrLf
                                Else
                                    TextDatos(III).BackColor = vbWhite

                                End If

                            End If

                        End If

                    End If

                End If

            Next i

            If Algun_Error Then
                BotonGuardar.Visible = False
            Else
                BotonGuardar.Visible = True

            End If
        
    End Select

End Sub

