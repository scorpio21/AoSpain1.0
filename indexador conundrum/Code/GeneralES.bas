Attribute VB_Name = "GeneralES"
Option Explicit
'********************************************************************************
'********************************************************************************
'********************************************************************************
'*********************** Funciones de Carga *************************************
'********************************************************************************
'********************************************************************************

Public Sub LoadGrhData(Optional ByVal FileNamePath As String = vbNullString)
    '*****************************************************************
    'Loads Grh.dat
    '*****************************************************************

    On Error GoTo ErrorHandler

    Dim Grh          As Long

    Dim Frame        As Integer

    Dim TempInt      As Integer

    Dim ArchivoAbrir As String

    'Resize arrays
    ReDim Grhdata(1 To MAXGrH) As Grhdata

    'Open files

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = IniPath & "Graficos.ind"
        Else
            ArchivoAbrir = IniPath & "Graficos" & SavePath & ".ind"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not FileExist(ArchivoAbrir, vbNormal) Then
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
        Exit Sub

    End If

    Open ArchivoAbrir For Binary Access Read As #1
    Seek #1, 1

    Get #1, , MiCabecera
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt

    'Fill Grh List

    'Get first Grh Number
    If UsarGrhLong Then
        Get #1, , Grh
    Else
        Get #1, , TempInt
        Grh = TempInt

    End If

    Do Until Grh <= 0

        'Get number of frames
        Get #1, , Grhdata(Grh).NumFrames
    
        If Grhdata(Grh).NumFrames <= 0 Then GoTo ErrorHandler
    
        If Grhdata(Grh).NumFrames > 1 Then
            frmMain.Lista.AddItem Grh & " (animacion)"

            'Read a animation GRH set
            For Frame = 1 To Grhdata(Grh).NumFrames

                If UsarGrhLong Then
                    Get #1, , Grhdata(Grh).Frames(Frame)
                Else
                    Get #1, , TempInt
                    Grhdata(Grh).Frames(Frame) = TempInt

                End If
            
                If Grhdata(Grh).Frames(Frame) <= 0 Or Grhdata(Grh).Frames(Frame) > MAXGrH Then
                    GoTo ErrorHandler

                End If
        
            Next Frame
    
            Get #1, , Grhdata(Grh).Speed

            If Grhdata(Grh).Speed <= 0 Then MsgBox Grh & " velocidad <= 0 ", , "advertencia"
        
            'Compute width and height
            Grhdata(Grh).pixelHeight = Grhdata(Grhdata(Grh).Frames(1)).pixelHeight

            If Grhdata(Grh).pixelHeight <= 0 Then MsgBox Grh & " ancho <= 0 ", , "advertencia"
        
            Grhdata(Grh).pixelWidth = Grhdata(Grhdata(Grh).Frames(1)).pixelWidth

            If Grhdata(Grh).pixelWidth <= 0 Then MsgBox Grh & " ancho <= 0 ", , "advertencia"
        
            Grhdata(Grh).TileWidth = Grhdata(Grhdata(Grh).Frames(1)).TileWidth

            If Grhdata(Grh).TileWidth <= 0 Then MsgBox Grh & " anchoT <= 0 ", , "advertencia"
        
            Grhdata(Grh).TileHeight = Grhdata(Grhdata(Grh).Frames(1)).TileHeight

            If Grhdata(Grh).TileHeight <= 0 Then MsgBox Grh & " altoT <= 0 ", , "advertencia"
    
        Else
            frmMain.Lista.AddItem Grh
            'Read in normal GRH data
            Get #1, , Grhdata(Grh).FileNum

            If Grhdata(Grh).FileNum <= 0 Then MsgBox Grh & " tiene bmp = 0 ", , "advertencia"
           
            Get #1, , Grhdata(Grh).sX

            If Grhdata(Grh).sX < 0 Then MsgBox Grh & " tiene Sx <= 0 ", , "advertencia"
        
            Get #1, , Grhdata(Grh).sY

            If Grhdata(Grh).sY < 0 Then MsgBox Grh & " tiene Sy <= 0 ", , "advertencia"
            
            Get #1, , Grhdata(Grh).pixelWidth

            If Grhdata(Grh).pixelWidth <= 0 Then MsgBox Grh & " ancho <= 0 ", , "advertencia"
        
            Get #1, , Grhdata(Grh).pixelHeight

            If Grhdata(Grh).pixelHeight <= 0 Then MsgBox Grh & " alto <= 0 ", , "advertencia"
        
            'Compute width and height
            Grhdata(Grh).TileWidth = Grhdata(Grh).pixelWidth / TilePixelHeight
            Grhdata(Grh).TileHeight = Grhdata(Grh).pixelHeight / TilePixelWidth
        
            Grhdata(Grh).Frames(1) = Grh
            
        End If

        'Get Next Grh Number
        If UsarGrhLong Then
            Get #1, , Grh
        Else
            Get #1, , TempInt
            Grh = TempInt

        End If

    Loop
    '************************************************

    Close #1

    Exit Sub

ErrorHandler:
    Close #1
    MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh

End Sub

Sub CargarAlas(Optional ByVal FileNamePath As String = vbNullString)

    On Error Resume Next

    Dim n             As Integer, i As Integer

    Dim NumAlas       As Integer

    Dim MisAlas()     As tIndiceAlas

    Dim MisAlasLong() As tIndiceAlasLong

    Dim ArchivoAbrir  As String

    n = FreeFile

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Alas.ind"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Alas" & SavePath & ".ind"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If
    
    If Not FileExist(ArchivoAbrir, vbNormal) Then
           
        frmMain.BotonI(10).Visible = False
        If UBound(AlaData()) = 0 Then
            ReDim AlaData(1) As AlaData

        End If

        Exit Sub

    End If
    
    If Not FileExist(ArchivoAbrir, vbNormal) Then
    
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
        frmMain.BotonI(10).Visible = False
        If UBound(AlaData()) = 0 Then
            ReDim AlaData(1) As AlaData

        End If

        Exit Sub

    End If

    Open ArchivoAbrir For Binary Access Read As #n

    'cabecera
    Get #n, , MiCabecera

    'num de cabezas
    Get #n, , NumAlas

    'Resize array
    ReDim AlaData(0 To NumAlas) As AlaData
    ReDim MisAlas(0 To NumAlas + 1) As tIndiceAlas
    ReDim MisAlasLong(0 To NumAlas + 1) As tIndiceAlasLong

    If UsarGrhLong Then

        For i = 1 To NumAlas
            Get #n, , MisAlasLong(i)
            InitGrh AlaData(i).Alas(1), MisAlasLong(i).Alas(1), 0
            InitGrh AlaData(i).Alas(2), MisAlasLong(i).Alas(2), 0
            InitGrh AlaData(i).Alas(3), MisAlasLong(i).Alas(3), 0
            InitGrh AlaData(i).Alas(4), MisAlasLong(i).Alas(4), 0
            AlaData(i).HeadOffset.X = MisAlasLong(i).HeadOffsetX
            AlaData(i).HeadOffset.Y = MisAlasLong(i).HeadOffsetY
        Next i

    Else

        For i = 1 To NumAlas
            Get #n, , MisAlas(i)
            InitGrh AlaData(i).Alas(1), MisAlas(i).Alas(1), 0
            InitGrh AlaData(i).Alas(2), MisAlas(i).Alas(2), 0
            InitGrh AlaData(i).Alas(3), MisAlas(i).Alas(3), 0
            InitGrh AlaData(i).Alas(4), MisAlas(i).Alas(4), 0
            AlaData(i).HeadOffset.X = MisAlas(i).HeadOffsetX
            AlaData(i).HeadOffset.Y = MisAlas(i).HeadOffsetY
        Next i

    End If

    Close #n

End Sub

Public Sub CargarAlasDat(Optional ByVal FileNamePath As String = vbNullString)

    On Error Resume Next

    Dim n            As Integer, i As Integer

    Dim NumAlas      As Integer

    Dim MisAlas()    As tIndiceAlasLong

    Dim ArchivoAbrir As String

    Dim loopc        As Integer

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\encode\Alas.dat"
        Else
            ArchivoAbrir = App.Path & "\encode\Alas" & SavePath & ".dat"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not FileExist(ArchivoAbrir, vbNormal) Then
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"

        Exit Sub

    End If

    Dim Leer As New clsIniReader

    Call Leer.Initialize(ArchivoAbrir)

    NumAlas = Val(Leer.GetValue("INIT", "NumAlas"))
    ReDim MisAlas(0 To NumAlas + 1) As tIndiceAlasLong

    ReDim AlaData(0 To NumAlas) As AlaData

    For loopc = 1 To NumAlas
        InitGrh AlaData(loopc).Alas(1), Val(Leer.GetValue("Alas" & loopc, "ALAS1")), 0
        InitGrh AlaData(loopc).Alas(2), Val(Leer.GetValue("Alas" & loopc, "ALAS2")), 0
        InitGrh AlaData(loopc).Alas(3), Val(Leer.GetValue("Alas" & loopc, "ALAS3")), 0
        InitGrh AlaData(loopc).Alas(4), Val(Leer.GetValue("Alas" & loopc, "ALAS4")), 0
        AlaData(loopc).HeadOffset.X = Val(Leer.GetValue("Alas" & loopc, "HeadOffsetX"))
        AlaData(loopc).HeadOffset.Y = Val(Leer.GetValue("Alas" & loopc, "HeadOffsety"))
    Next loopc

    Set Leer = Nothing

End Sub

Public Sub CargarCuerposDat(Optional ByVal FileNamePath As String = vbNullString)

    On Error Resume Next

    Dim n            As Integer, i As Integer

    Dim NumCuerpos   As Integer

    Dim MisCuerpos() As tIndiceCuerpoLong

    Dim ArchivoAbrir As String

    Dim loopc        As Integer

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\encode\Body.dat"
        Else
            ArchivoAbrir = App.Path & "\encode\Body" & SavePath & ".dat"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not FileExist(ArchivoAbrir, vbNormal) Then
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"

        Exit Sub

    End If

    Dim Leer As New clsIniReader

    Call Leer.Initialize(ArchivoAbrir)

    NumCuerpos = Val(Leer.GetValue("INIT", "NumBodies"))
    ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpoLong

    ReDim BodyData(0 To NumCuerpos) As BodyData

    For loopc = 1 To NumCuerpos
        InitGrh BodyData(loopc).Walk(1), Val(Leer.GetValue("Body" & loopc, "WALK1")), 0
        InitGrh BodyData(loopc).Walk(2), Val(Leer.GetValue("Body" & loopc, "WALK2")), 0
        InitGrh BodyData(loopc).Walk(3), Val(Leer.GetValue("Body" & loopc, "WALK3")), 0
        InitGrh BodyData(loopc).Walk(4), Val(Leer.GetValue("body" & loopc, "WALK4")), 0
        BodyData(loopc).HeadOffset.X = Val(Leer.GetValue("body" & loopc, "HeadOffsetX"))
        BodyData(loopc).HeadOffset.Y = Val(Leer.GetValue("body" & loopc, "HeadOffsety"))
    Next loopc

    Set Leer = Nothing

End Sub

Public Sub CargarCabezasdat(Optional ByVal FileNamePath As String = vbNullString)

    On Error Resume Next

    Dim n            As Integer, i As Integer, Index As Integer

    Dim ArchivoAbrir As String

    Dim loopc        As Long

    Dim Miscabezas() As tIndiceCabezaLong

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\encode\Head.dat"
        Else
            ArchivoAbrir = App.Path & "\encode\Head" & SavePath & ".dat"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not FileExist(ArchivoAbrir, vbNormal) Then
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
        Exit Sub

    End If

    Dim Leer As New clsIniReader

    Call Leer.Initialize(ArchivoAbrir)

    Numheads = Val(Leer.GetValue("INIT", "NumHeads"))

    ReDim HeadData(0 To Numheads) As HeadData

    For i = 1 To Numheads
        InitGrh HeadData(i).Head(1), Val(Leer.GetValue("Head" & i, "Head1")), 0
        InitGrh HeadData(i).Head(2), Val(Leer.GetValue("Head" & i, "Head2")), 0
        InitGrh HeadData(i).Head(3), Val(Leer.GetValue("Head" & i, "Head3")), 0
        InitGrh HeadData(i).Head(4), Val(Leer.GetValue("Head" & i, "Head4")), 0
        DoEvents
        frmMain.LUlitError.Caption = "cabeza: " & i
    Next i

    Set Leer = Nothing

End Sub

Public Sub CargarEspaldaDat(Optional ByVal FileNamePath As String = vbNullString)

    On Error Resume Next

    Dim n            As Integer, i As Integer, NumEspalda As Integer, Index As Integer

    Dim ArchivoAbrir As String

    Dim MisE()       As tIndiceCabezaLong

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\encode\Capas.dat"
        Else
            ArchivoAbrir = App.Path & "\encode\Capas" & SavePath & ".dat"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not FileExist(ArchivoAbrir, vbNormal) Then
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
        Exit Sub

    End If

    'Resize array
    Numcapas = Val(GetVar(ArchivoAbrir, "INIT", "NumCapas"))

    ReDim EspaldaAnimData(0 To NumEspalda) As HeadData

    For i = 1 To NumEspalda
        InitGrh EspaldaAnimData(i).Head(1), Val(GetVar(ArchivoAbrir, "Capa" & i, "Capa1")), 0
        InitGrh EspaldaAnimData(i).Head(2), Val(GetVar(ArchivoAbrir, "Capa" & i, "Capa2")), 0
        InitGrh EspaldaAnimData(i).Head(3), Val(GetVar(ArchivoAbrir, "Capa" & i, "Capa3")), 0
        InitGrh EspaldaAnimData(i).Head(4), Val(GetVar(ArchivoAbrir, "Capa" & i, "Capa4")), 0
    Next i

End Sub

Public Sub CargarBotasDat(Optional ByVal FileNamePath As String = vbNullString)

    On Error Resume Next

    Dim n            As Integer, i As Integer

    Dim NumBotas     As Integer

    Dim MisBotas()   As tIndiceBotasLong

    Dim ArchivoAbrir As String

    Dim loopc        As Integer

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\encode\Botas.dat"
        Else
            ArchivoAbrir = App.Path & "\encode\Botas" & SavePath & ".dat"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not FileExist(ArchivoAbrir, vbNormal) Then
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"

        Exit Sub

    End If

    Dim Leer As New clsIniReader

    Call Leer.Initialize(ArchivoAbrir)

    NumBotas = Val(Leer.GetValue("INIT", "NumBotas"))
    ReDim MisBotas(0 To NumBotas + 1) As tIndiceBotasLong

    ReDim BotaData(0 To NumBotas) As BotaData

    For loopc = 1 To NumBotas
        InitGrh BotaData(loopc).Bota(1), Val(Leer.GetValue("Bota" & loopc, "BOTA1")), 0
        InitGrh BotaData(loopc).Bota(2), Val(Leer.GetValue("Bota" & loopc, "BOTA2")), 0
        InitGrh BotaData(loopc).Bota(3), Val(Leer.GetValue("Bota" & loopc, "BOTA3")), 0
        InitGrh BotaData(loopc).Bota(4), Val(Leer.GetValue("Bota" & loopc, "BOTA4")), 0
        BotaData(loopc).HeadOffset.X = Val(Leer.GetValue("Bota" & loopc, "HeadOffsetX"))
        BotaData(loopc).HeadOffset.Y = Val(Leer.GetValue("Bota" & loopc, "HeadOffsety"))
    Next loopc

    Set Leer = Nothing

End Sub

Public Sub CargarCascosDat(Optional ByVal FileNamePath As String = vbNullString)

    On Error Resume Next

    Dim ArchivoAbrir As String

    Dim n            As Integer, i As Integer, NumCascos As Integer, Index As Integer

    Dim Miscabezas() As tIndiceCabezaLong

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\encode\Cascos.dat"
        Else
            ArchivoAbrir = App.Path & "\encode\Cascos" & SavePath & ".dat"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not FileExist(ArchivoAbrir, vbNormal) Then
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
        Exit Sub

    End If

    NumCascos = Val(GetVar(ArchivoAbrir, "INIT", "NumCascos"))
    'Resize array
    ReDim CascoAnimData(0 To NumCascos + 1) As HeadData

    For i = 1 To NumCascos
        InitGrh CascoAnimData(i).Head(1), Val(GetVar(ArchivoAbrir, "Casco" & i, "Head1")), 0
        InitGrh CascoAnimData(i).Head(2), Val(GetVar(ArchivoAbrir, "Casco" & i, "Head2")), 0
        InitGrh CascoAnimData(i).Head(3), Val(GetVar(ArchivoAbrir, "Casco" & i, "Head3")), 0
        InitGrh CascoAnimData(i).Head(4), Val(GetVar(ArchivoAbrir, "Casco" & i, "Head4")), 0
    Next i

End Sub

Public Sub CargarFxsDat(Optional ByVal FileNamePath As String = vbNullString)

    On Error Resume Next

    Dim n            As Integer, i As Integer

    Dim MisFxs()     As tIndiceFxLong

    Dim ArchivoAbrir As String

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\encode\fx.dat"
        Else
            ArchivoAbrir = App.Path & "\encode\fx" & SavePath & ".dat"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not FileExist(ArchivoAbrir, vbNormal) Then
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
        Exit Sub

    End If

    numfxs = Val(GetVar(ArchivoAbrir, "INIT", "NumFxs"))
    ReDim FxData(0 To numfxs) As FxData

    For i = 1 To numfxs
        Call InitGrh(FxData(i).Fx, Val(GetVar(ArchivoAbrir, "Fx" & i, "Animacion")), 1)
        FxData(i).OffsetX = Val(GetVar(ArchivoAbrir, "Fx" & i, "OffsetX"))
        FxData(i).OffsetY = Val(GetVar(ArchivoAbrir, "Fx" & i, "OffsetY"))
    Next i

End Sub

Public Sub CargarCabezas(Optional ByVal FileNamePath As String = vbNullString)

    On Error Resume Next

    Dim n                As Integer, i As Integer, Index As Integer

    Dim ArchivoAbrir     As String

    Dim Miscabezas()     As tIndiceCabeza

    Dim MiscabezasLong() As tIndiceCabezaLong

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Cabezas.ind"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Cabezas" & SavePath & ".ind"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not FileExist(ArchivoAbrir, vbNormal) Then
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"

        If UBound(HeadData()) = 0 Then
            ReDim HeadData(1) As HeadData

        End If
    
        Exit Sub

    End If

    n = FreeFile
    Open ArchivoAbrir For Binary Access Read As #n

    'cabecera
    Get #n, , MiCabecera

    'num de cabezas
    Get #n, , Numheads

    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads + 1) As tIndiceCabeza
    ReDim MiscabezasLong(0 To Numheads + 1) As tIndiceCabezaLong

    If UsarGrhLong Then

        For i = 1 To Numheads
            Get #n, , MiscabezasLong(i)
            InitGrh HeadData(i).Head(1), MiscabezasLong(i).Head(1), 0
            InitGrh HeadData(i).Head(2), MiscabezasLong(i).Head(2), 0
            InitGrh HeadData(i).Head(3), MiscabezasLong(i).Head(3), 0
            InitGrh HeadData(i).Head(4), MiscabezasLong(i).Head(4), 0
        Next i

    Else

        For i = 1 To Numheads
            Get #n, , Miscabezas(i)
            InitGrh HeadData(i).Head(1), Miscabezas(i).Head(1), 0
            InitGrh HeadData(i).Head(2), Miscabezas(i).Head(2), 0
            InitGrh HeadData(i).Head(3), Miscabezas(i).Head(3), 0
            InitGrh HeadData(i).Head(4), Miscabezas(i).Head(4), 0
        Next i

    End If

    Close #n

End Sub

Public Sub CargarCascos(Optional ByVal FileNamePath As String = vbNullString)

    On Error Resume Next

    Dim ArchivoAbrir     As String

    Dim n                As Integer, i As Integer, NumCascos As Integer, Index As Integer

    Dim Miscabezas()     As tIndiceCabeza

    Dim MiscabezasLong() As tIndiceCabezaLong

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Cascos.ind"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Cascos" & SavePath & ".ind"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not FileExist(ArchivoAbrir, vbNormal) Then
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"

        If UBound(CascoAnimData()) = 0 Then
            ReDim CascoAnimData(1) As HeadData

        End If

        Exit Sub

    End If

    n = FreeFile
    Open ArchivoAbrir For Binary Access Read As #n

    'cabecera
    Get #n, , MiCabecera

    'num de cabezas
    Get #n, , NumCascos

    'Resize array
    ReDim CascoAnimData(0 To NumCascos + 1) As HeadData
    ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza
    ReDim MiscabezasLong(0 To Numheads + 1) As tIndiceCabezaLong

    If UsarGrhLong Then

        For i = 1 To NumCascos
            Get #n, , MiscabezasLong(i)
            InitGrh CascoAnimData(i).Head(1), MiscabezasLong(i).Head(1), 0
            InitGrh CascoAnimData(i).Head(2), MiscabezasLong(i).Head(2), 0
            InitGrh CascoAnimData(i).Head(3), MiscabezasLong(i).Head(3), 0
            InitGrh CascoAnimData(i).Head(4), MiscabezasLong(i).Head(4), 0
        Next i

    Else

        For i = 1 To NumCascos
            Get #n, , Miscabezas(i)
            InitGrh CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0
            InitGrh CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0
            InitGrh CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0
            InitGrh CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0
        Next i

    End If

    Close #n

End Sub

Public Sub CargarEspalda(Optional ByVal FileNamePath As String = vbNullString)

    On Error Resume Next

    Dim n            As Integer, i As Integer, NumEspalda As Integer, Index As Integer

    Dim ArchivoAbrir As String

    Dim MisE()       As tIndiceCabeza

    Dim MisELong()   As tIndiceCabezaLong

    n = FreeFile

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Capas.ind"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Capas" & SavePath & ".ind"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not FileExist(ArchivoAbrir, vbNormal) Then
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"

        If UBound(EspaldaAnimData()) = 0 Then
            ReDim EspaldaAnimData(1) As HeadData

        End If

        Exit Sub

    End If

    Open ArchivoAbrir For Binary Access Read As #n

    'cabecera
    Get #n, , MiCabecera

    'num de cabezas
    Get #n, , NumEspalda

    'Resize array
    ReDim EspaldaAnimData(0 To NumEspalda) As HeadData
    ReDim MisE(0 To NumEspalda + 1) As tIndiceCabeza
    ReDim MisELong(0 To NumEspalda + 1) As tIndiceCabezaLong

    If UsarGrhLong Then

        For i = 1 To NumEspalda
            Get #n, , MisELong(i)
            InitGrh EspaldaAnimData(i).Head(1), MisELong(i).Head(1), 0
            InitGrh EspaldaAnimData(i).Head(2), MisELong(i).Head(2), 0
            InitGrh EspaldaAnimData(i).Head(3), MisELong(i).Head(3), 0
            InitGrh EspaldaAnimData(i).Head(4), MisELong(i).Head(4), 0
        Next i

    Else

        For i = 1 To NumEspalda
            Get #n, , MisE(i)
            InitGrh EspaldaAnimData(i).Head(1), MisE(i).Head(1), 0
            InitGrh EspaldaAnimData(i).Head(2), MisE(i).Head(2), 0
            InitGrh EspaldaAnimData(i).Head(3), MisE(i).Head(3), 0
            InitGrh EspaldaAnimData(i).Head(4), MisE(i).Head(4), 0
        Next i

    End If

    Close #n

End Sub

Public Sub CargarBotas(Optional ByVal FileNamePath As String = vbNullString)

    On Error Resume Next

    Dim n              As Integer, i As Integer

    Dim NumBotas       As Integer

    Dim MisBotas()     As tIndiceBotas

    Dim MisBotasLong() As tIndiceBotasLong

    Dim ArchivoAbrir   As String

    n = FreeFile

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Botas.ind"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Botas" & SavePath & ".ind"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If
    
     If Not FileExist(ArchivoAbrir, vbNormal) Then
        
        frmMain.BotonI(6).Visible = False
        If UBound(BotaData()) = 0 Then
            ReDim BotaData(1) As BotaData

        End If

        Exit Sub

    End If


    If Not FileExist(ArchivoAbrir, vbNormal) Then
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
        frmMain.BotonI(6).Visible = False
        If UBound(BotaData()) = 0 Then
            ReDim BotaData(1) As BotaData

        End If

        Exit Sub

    End If

    Open ArchivoAbrir For Binary Access Read As #n

    'cabecera
    Get #n, , MiCabecera

    'num de cabezas
    Get #n, , NumBotas

    'Resize array
    ReDim BotaData(0 To NumBotas) As BotaData
    ReDim MisBotas(0 To NumBotas + 1) As tIndiceBotas
    ReDim MisBotasLong(0 To NumBotas + 1) As tIndiceBotasLong

    If UsarGrhLong Then

        For i = 1 To NumBotas
            Get #n, , MisBotasLong(i)
            InitGrh BotaData(i).Bota(1), MisBotasLong(i).Bota(1), 0
            InitGrh BotaData(i).Bota(2), MisBotasLong(i).Bota(2), 0
            InitGrh BotaData(i).Bota(3), MisBotasLong(i).Bota(3), 0
            InitGrh BotaData(i).Bota(4), MisBotasLong(i).Bota(4), 0
            BotaData(i).HeadOffset.X = MisBotasLong(i).HeadOffsetX
            BotaData(i).HeadOffset.Y = MisBotasLong(i).HeadOffsetY
        Next i

    Else

        For i = 1 To NumBotas
            Get #n, , MisBotas(i)
            InitGrh BotaData(i).Bota(1), MisBotas(i).Bota(1), 0
            InitGrh BotaData(i).Bota(2), MisBotas(i).Bota(2), 0
            InitGrh BotaData(i).Bota(3), MisBotas(i).Bota(3), 0
            InitGrh BotaData(i).Bota(4), MisBotas(i).Bota(4), 0
            BotaData(i).HeadOffset.X = MisBotas(i).HeadOffsetX
            BotaData(i).HeadOffset.Y = MisBotas(i).HeadOffsetY
        Next i

    End If

    Close #n

End Sub

Sub CargarCuerpos(Optional ByVal FileNamePath As String = vbNullString)

    On Error Resume Next

    Dim n                As Integer, i As Integer

    Dim NumCuerpos       As Integer

    Dim MisCuerpos()     As tIndiceCuerpo

    Dim MisCuerposLong() As tIndiceCuerpoLong

    Dim ArchivoAbrir     As String

    n = FreeFile

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Personajes.ind"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Personajes" & SavePath & ".ind"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not FileExist(ArchivoAbrir, vbNormal) Then
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"

        If UBound(BodyData()) = 0 Then
            ReDim BodyData(1) As BodyData

        End If

        Exit Sub

    End If

    Open ArchivoAbrir For Binary Access Read As #n

    'cabecera
    Get #n, , MiCabecera

    'num de cabezas
    Get #n, , NumCuerpos

    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpo
    ReDim MisCuerposLong(0 To NumCuerpos + 1) As tIndiceCuerpoLong

    If UsarGrhLong Then

        For i = 1 To NumCuerpos
            Get #n, , MisCuerposLong(i)
            InitGrh BodyData(i).Walk(1), MisCuerposLong(i).Body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerposLong(i).Body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerposLong(i).Body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerposLong(i).Body(4), 0
            BodyData(i).HeadOffset.X = MisCuerposLong(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerposLong(i).HeadOffsetY
        Next i

    Else

        For i = 1 To NumCuerpos
            Get #n, , MisCuerpos(i)
            InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        Next i

    End If

    Close #n

End Sub

Public Sub CargarFxs(Optional ByVal FileNamePath As String = vbNullString)

    On Error Resume Next

    Dim n            As Integer, i As Integer

    Dim MisFxs()     As tIndiceFx

    Dim MisFxslong() As tIndiceFxLong

    Dim ArchivoAbrir As String

    n = FreeFile

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Fxs.ind"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Fxs" & SavePath & ".ind"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not FileExist(ArchivoAbrir, vbNormal) Then
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"

        If UBound(FxData()) = 0 Then
            ReDim FxData(1) As FxData

        End If

        Exit Sub

    End If

    Open ArchivoAbrir For Binary Access Read As #n

    'cabecera
    Get #n, , MiCabecera

    'num de cabezas
    Get #n, , numfxs

    'Resize array
    ReDim FxData(0 To numfxs) As FxData
    ReDim MisFxs(0 To numfxs + 1) As tIndiceFx
    ReDim MisFxslong(0 To numfxs + 1) As tIndiceFxLong

    If UsarGrhLong Then

        For i = 1 To numfxs
            Get #n, , MisFxslong(i)
            Call InitGrh(FxData(i).Fx, MisFxslong(i).Animacion, 1)
            FxData(i).OffsetX = MisFxslong(i).OffsetX
            FxData(i).OffsetY = MisFxslong(i).OffsetY
        Next i

    Else

        For i = 1 To numfxs
            Get #n, , MisFxs(i)
            Call InitGrh(FxData(i).Fx, MisFxs(i).Animacion, 1)
            FxData(i).OffsetX = MisFxs(i).OffsetX
            FxData(i).OffsetY = MisFxs(i).OffsetY
        Next i

    End If

    Close #n

End Sub

'********************************************************************************
'********************************************************************************
'********************************************************************************
'********************** Funciones de guardado ***********************************
'********************************************************************************
'********************************************************************************

Public Sub GuardarCabezas(Optional ByVal FileNamePath As String = vbNullString)

    On Error GoTo ErrHandler

    Dim n As Integer, i As Integer

    n = FreeFile

    Dim ArchivoAbrir As String

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Cabezas.ind"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Cabezas" & SavePath & ".ind"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

    Open ArchivoAbrir For Binary As #n
    'Escribimos la cabecera
    Put #n, , MiCabecera
    'Guardamos las cabezas

    Put #n, , CInt(UBound(HeadData)) 'numheads

    Dim Miscabezas() As tIndiceCabeza

    ReDim Miscabezas(0 To UBound(HeadData) + 1) As tIndiceCabeza
    ReDim MiscabezasLong(0 To UBound(HeadData) + 1) As tIndiceCabezaLong

    If UsarGrhLong Then

        For i = 1 To UBound(HeadData)
            MiscabezasLong(i).Head(1) = HeadData(i).Head(1).GrhIndex
            MiscabezasLong(i).Head(2) = HeadData(i).Head(2).GrhIndex
            MiscabezasLong(i).Head(3) = HeadData(i).Head(3).GrhIndex
            MiscabezasLong(i).Head(4) = HeadData(i).Head(4).GrhIndex
            Put #n, , MiscabezasLong(i)
        Next i

    Else

        For i = 1 To UBound(HeadData)
            Miscabezas(i).Head(1) = HeadData(i).Head(1).GrhIndex
            Miscabezas(i).Head(2) = HeadData(i).Head(2).GrhIndex
            Miscabezas(i).Head(3) = HeadData(i).Head(3).GrhIndex
            Miscabezas(i).Head(4) = HeadData(i).Head(4).GrhIndex
            Put #n, , Miscabezas(i)
        Next i

    End If

    Close #n

    frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

    EstadoNoGuardado(e_EstadoIndexador.Cabezas) = False

    Exit Sub
ErrHandler:
    Call MsgBox("Error en cabeza" & i)

End Sub

Public Sub GuardarCabezasDat(Optional ByVal FileNamePath As String = vbNullString)

    On Error GoTo ErrHandler

    Dim n As Integer, i As Integer

    n = FreeFile

    Dim ArchivoAbrir As String

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\encode\Head.dat"
        Else
            ArchivoAbrir = App.Path & "\encode\Head" & SavePath & ".dat"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

    Call WriteVar(ArchivoAbrir, "INIT", "NumHeads", CInt(UBound(HeadData)))

    For i = 1 To UBound(HeadData)

        If HeadData(i).Head(1).GrhIndex > 0 Then
            Call WriteVar(ArchivoAbrir, "Head" & i, "Head1", HeadData(i).Head(1).GrhIndex)
            Call WriteVar(ArchivoAbrir, "Head" & i, "Head2", HeadData(i).Head(2).GrhIndex)
            Call WriteVar(ArchivoAbrir, "Head" & i, "Head3", HeadData(i).Head(3).GrhIndex)
            Call WriteVar(ArchivoAbrir, "Head" & i, "Head4", HeadData(i).Head(4).GrhIndex)
            DoEvents
            frmMain.LUlitError.Caption = "cabeza: " & i

        End If

    Next i

    frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

    EstadoNoGuardado(e_EstadoIndexador.Cabezas) = False

    Exit Sub

ErrHandler:
    Call MsgBox("Error en cabeza" & i)

End Sub

Public Sub GuardarFxs(Optional ByVal FileNamePath As String = vbNullString)

    On Error GoTo ErrHandler

    Dim MisFxslong() As tIndiceFxLong

    numfxs = UBound(FxData)
    ReDim FxDataI(0 To numfxs + 1) As tIndiceFx
    ReDim MisFxslong(0 To numfxs + 1) As tIndiceFxLong

    Dim n As Integer, i As Integer

    n = FreeFile

    Dim ArchivoAbrir As String

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Fxs.ind"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Fxs" & SavePath & ".ind"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

    Open ArchivoAbrir For Binary As #n

    'Escribimos la cabecera
    Put #n, , MiCabecera
    'Guardamos las cabezas
    Put #n, , numfxs

    If UsarGrhLong Then

        For i = 1 To numfxs
            MisFxslong(i).Animacion = FxData(i).Fx.GrhIndex
            MisFxslong(i).OffsetX = FxData(i).OffsetX
            MisFxslong(i).OffsetY = FxData(i).OffsetY
            Put #n, , MisFxslong(i)
        Next i

    Else

        For i = 1 To numfxs
            FxDataI(i).Animacion = FxData(i).Fx.GrhIndex
            FxDataI(i).OffsetX = FxData(i).OffsetX
            FxDataI(i).OffsetY = FxData(i).OffsetY
            Put #n, , FxDataI(i)
        Next i

    End If

    Close #n

    frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

    EstadoNoGuardado(e_EstadoIndexador.Fx) = False

    Exit Sub
ErrHandler:
    Call MsgBox("Error en FX " & i)

End Sub

Public Sub GuardarFxsDat(Optional ByVal FileNamePath As String = vbNullString)

    On Error GoTo ErrHandler

    numfxs = UBound(FxData)

    Dim n As Integer, i As Integer

    n = FreeFile

    Dim ArchivoAbrir As String

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\encode\fx.dat"
        Else
            ArchivoAbrir = App.Path & "\encode\fx" & SavePath & ".dat"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

    Call WriteVar(ArchivoAbrir, "INIT", "NumFxs", numfxs)

    For i = 1 To numfxs

        If FxData(i).Fx.GrhIndex > 0 Then
            Call WriteVar(ArchivoAbrir, "Fx" & i, "Animacion", FxData(i).Fx.GrhIndex)
            Call WriteVar(ArchivoAbrir, "Fx" & i, "OffsetX", FxData(i).OffsetX)
            Call WriteVar(ArchivoAbrir, "Fx" & i, "OffsetY", FxData(i).OffsetY)
            frmMain.LUlitError.Caption = "Fx : " & i
            DoEvents

        End If

    Next i

    frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

    EstadoNoGuardado(e_EstadoIndexador.Fx) = False

    Exit Sub
ErrHandler:
    Call MsgBox("Error en FX " & i)

End Sub

Public Sub GuardarBotas(Optional ByVal FileNamePath As String = vbNullString)

    On Error GoTo ErrHandler

    Dim n As Integer, i As Integer

    n = FreeFile

    Dim ArchivoAbrir As String

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Botas.ind"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Botas" & SavePath & ".ind"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

    Open ArchivoAbrir For Binary As #n
    'Escribimos la cabecera
    Put #n, , MiCabecera
    'Guardamos las cabezas

    Dim MisBotas()     As tIndiceBotas

    Dim MisBotasLong() As tIndiceBotasLong

    ReDim MisBotas(0 To UBound(BotaData) + 1) As tIndiceBotas

    ReDim MisBotasLong(0 To UBound(BotaData) + 1) As tIndiceBotasLong

    Put #n, , CInt(UBound(BotaData)) 'numheads

    If UsarGrhLong Then

        For i = 1 To UBound(BotaData)
            MisBotasLong(i).Bota(1) = BotaData(i).Bota(1).GrhIndex
            MisBotasLong(i).Bota(2) = BotaData(i).Bota(2).GrhIndex
            MisBotasLong(i).Bota(3) = BotaData(i).Bota(3).GrhIndex
            MisBotasLong(i).Bota(4) = BotaData(i).Bota(4).GrhIndex
            MisBotasLong(i).HeadOffsetX = BotaData(i).HeadOffset.X
            MisBotasLong(i).HeadOffsetY = BotaData(i).HeadOffset.Y
            Put #n, , MisBotas(i)
        Next i

    Else

        For i = 1 To UBound(BotaData)
            MisBotas(i).Bota(1) = BotaData(i).Bota(1).GrhIndex
            MisBotas(i).Bota(1) = BotaData(i).Bota(2).GrhIndex
            MisBotas(i).Bota(1) = BotaData(i).Bota(3).GrhIndex
            MisBotas(i).Bota(1) = BotaData(i).Bota(4).GrhIndex
            MisBotas(i).HeadOffsetX = BotaData(i).HeadOffset.X
            MisBotas(i).HeadOffsetY = BotaData(i).HeadOffset.Y
            Put #n, , MisBotas(i)
        Next i

    End If

    Close #n

    frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

    EstadoNoGuardado(e_EstadoIndexador.Botas) = False

    Exit Sub
ErrHandler:
    Call MsgBox("Error en bota " & i)

End Sub

Public Sub GuardarBotasDat(Optional ByVal FileNamePath As String = vbNullString)

    On Error GoTo ErrHandler

    Dim n As Integer, i As Integer

    n = FreeFile

    Dim ArchivoAbrir As String

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\encode\Botas.dat"
        Else
            ArchivoAbrir = App.Path & "\encode\Botas" & SavePath & ".dat"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

    Call WriteVar(ArchivoAbrir, "INIT", "NumBotas", CInt(UBound(BotaData)))

    For i = 1 To UBound(BotaData)

        If BotaData(i).Bota(1).GrhIndex > 0 Then
            Call WriteVar(ArchivoAbrir, "Bota" & i, "Bota1", BotaData(i).Bota(1).GrhIndex)
            Call WriteVar(ArchivoAbrir, "Bota" & i, "Bota2", BotaData(i).Bota(2).GrhIndex)
            Call WriteVar(ArchivoAbrir, "Bota" & i, "Bota3", BotaData(i).Bota(3).GrhIndex)
            Call WriteVar(ArchivoAbrir, "Bota" & i, "Bota4", BotaData(i).Bota(4).GrhIndex)

        End If

    Next i

    frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

    EstadoNoGuardado(e_EstadoIndexador.Botas) = False

    Exit Sub
ErrHandler:
    Call MsgBox("Error en bota " & i)

End Sub

Public Sub GuardarCapas(Optional ByVal FileNamePath As String = vbNullString)

    On Error GoTo ErrHandler

    Dim n As Integer, i As Integer

    n = FreeFile

    Dim ArchivoAbrir As String

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Capas.ind"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Capas" & SavePath & ".ind"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

    Open ArchivoAbrir For Binary As #n

    Put #n, , MiCabecera

    Put #n, , CInt(UBound(EspaldaAnimData))  'numheads

    Dim Miscabezas() As tIndiceCabeza

    ReDim Miscabezas(0 To UBound(EspaldaAnimData) + 1) As tIndiceCabeza
    ReDim MiscabezasLong(0 To UBound(EspaldaAnimData) + 1) As tIndiceCabezaLong

    If UsarGrhLong Then

        For i = 1 To UBound(EspaldaAnimData)
            MiscabezasLong(i).Head(1) = EspaldaAnimData(i).Head(1).GrhIndex
            MiscabezasLong(i).Head(2) = EspaldaAnimData(i).Head(2).GrhIndex
            MiscabezasLong(i).Head(3) = EspaldaAnimData(i).Head(3).GrhIndex
            MiscabezasLong(i).Head(4) = EspaldaAnimData(i).Head(4).GrhIndex
            Put #n, , MiscabezasLong(i)
        Next i

    Else

        For i = 1 To UBound(EspaldaAnimData)
            Miscabezas(i).Head(1) = EspaldaAnimData(i).Head(1).GrhIndex
            Miscabezas(i).Head(2) = EspaldaAnimData(i).Head(2).GrhIndex
            Miscabezas(i).Head(3) = EspaldaAnimData(i).Head(3).GrhIndex
            Miscabezas(i).Head(4) = EspaldaAnimData(i).Head(4).GrhIndex
            Put #n, , Miscabezas(i)
        Next i

    End If

    Close #n
    frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

    EstadoNoGuardado(e_EstadoIndexador.Capas) = False

    Exit Sub
ErrHandler:
    Call MsgBox("Error en capa " & i)

End Sub

Public Sub GuardarCapasDat(Optional ByVal FileNamePath As String = vbNullString)

    On Error GoTo ErrHandler

    Dim n As Integer, i As Integer

    n = FreeFile

    Dim ArchivoAbrir As String

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\encode\Capas.dat"
        Else
            ArchivoAbrir = App.Path & "\encode\Capas" & SavePath & ".dat"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

    'Resize array
    Call WriteVar(ArchivoAbrir, "INIT", "NumCapas", CInt(UBound(EspaldaAnimData)))

    For i = 1 To UBound(EspaldaAnimData)

        If EspaldaAnimData(i).Head(1).GrhIndex > 0 Then
            Call WriteVar(ArchivoAbrir, "Capa" & i, "Capa1", EspaldaAnimData(i).Head(1).GrhIndex)
            Call WriteVar(ArchivoAbrir, "Capa" & i, "Capa2", EspaldaAnimData(i).Head(2).GrhIndex)
            Call WriteVar(ArchivoAbrir, "Capa" & i, "Capa3", EspaldaAnimData(i).Head(3).GrhIndex)
            Call WriteVar(ArchivoAbrir, "Capa" & i, "Capa4", EspaldaAnimData(i).Head(4).GrhIndex)

        End If

    Next i

    frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

    EstadoNoGuardado(e_EstadoIndexador.Capas) = False

    Exit Sub
ErrHandler:
    Call MsgBox("Error en capa " & i)

End Sub

Public Sub GuardarAlas(Optional ByVal FileNamePath As String = vbNullString)

    On Error GoTo ErrHandler

    Dim n As Integer, i As Integer

    n = FreeFile

    Dim ArchivoAbrir As String

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Alas.ind"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Alas" & SavePath & ".ind"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

    Open ArchivoAbrir For Binary As #n
    'Escribimos la cabecera
    Put #n, , MiCabecera
    'Guardamos las cabezas

    Dim MisAlas()     As tIndiceAlas

    Dim MisAlasLong() As tIndiceAlasLong

    ReDim MisAlas(0 To UBound(AlaData) + 1) As tIndiceAlas

    ReDim MisAlasLong(0 To UBound(AlaData) + 1) As tIndiceAlasLong

    Put #n, , CInt(UBound(AlaData)) 'numheads

    If UsarGrhLong Then

        For i = 1 To UBound(AlaData)
            MisAlasLong(i).Alas(1) = AlaData(i).Alas(1).GrhIndex
            MisAlasLong(i).Alas(2) = AlaData(i).Alas(2).GrhIndex
            MisAlasLong(i).Alas(3) = AlaData(i).Alas(3).GrhIndex
            MisAlasLong(i).Alas(4) = AlaData(i).Alas(4).GrhIndex
            MisAlasLong(i).HeadOffsetX = AlaData(i).HeadOffset.X
            MisAlasLong(i).HeadOffsetY = AlaData(i).HeadOffset.Y
            Put #n, , MisAlas(i)
        Next i

    Else

        For i = 1 To UBound(AlaData)
            MisAlas(i).Alas(1) = AlaData(i).Alas(1).GrhIndex
            MisAlas(i).Alas(2) = AlaData(i).Alas(2).GrhIndex
            MisAlas(i).Alas(3) = AlaData(i).Alas(3).GrhIndex
            MisAlas(i).Alas(4) = AlaData(i).Alas(4).GrhIndex
            MisAlas(i).HeadOffsetX = AlaData(i).HeadOffset.X
            MisAlas(i).HeadOffsetY = AlaData(i).HeadOffset.Y
            Put #n, , MisAlas(i)
        Next i

    End If

    Close #n

    frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

    EstadoNoGuardado(e_EstadoIndexador.Alas) = False

    Exit Sub
ErrHandler:
    Call MsgBox("Error en Alas " & i & " . " & Err.Description)

End Sub

Public Sub GuardarAlasDat(Optional ByVal FileNamePath As String = vbNullString)

    On Error GoTo ErrHandler

    Dim n As Integer, i As Integer

    n = FreeFile

    Dim ArchivoAbrir As String

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\encode\Alas.dat"
        Else
            ArchivoAbrir = App.Path & "\encode\Alas" & SavePath & ".dat"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

    Call WriteVar(ArchivoAbrir, "INIT", "NumAlas", CInt(UBound(AlaData))) 'numheads

    For i = 1 To UBound(AlaData)

        If AlaData(i).Alas(1).GrhIndex > 0 Then
            Call WriteVar(ArchivoAbrir, "Alas" & i, "ALAS1", AlaData(i).Alas(1).GrhIndex)
            Call WriteVar(ArchivoAbrir, "Alas" & i, "ALAS2", AlaData(i).Alas(2).GrhIndex)
            Call WriteVar(ArchivoAbrir, "Alas" & i, "ALAS3", AlaData(i).Alas(3).GrhIndex)
            Call WriteVar(ArchivoAbrir, "Alas" & i, "ALAS4", AlaData(i).Alas(4).GrhIndex)
            Call WriteVar(ArchivoAbrir, "Alas" & i, "HeadOffsetX", AlaData(i).HeadOffset.X)
            Call WriteVar(ArchivoAbrir, "Alas" & i, "HeadOffsety", AlaData(i).HeadOffset.Y)
            frmMain.LUlitError.Caption = "Alas : " & i
            DoEvents

        End If

    Next i

    frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir
    Exit Sub

    EstadoNoGuardado(e_EstadoIndexador.Alas) = False

ErrHandler:
    Call MsgBox("Error en Alas " & i & " . " & Err.Description)

End Sub

Public Sub GuardarBodys(Optional ByVal FileNamePath As String = vbNullString)

    On Error GoTo ErrHandler

    Dim n As Integer, i As Integer

    n = FreeFile

    Dim ArchivoAbrir As String

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Personajes.ind"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Personajes" & SavePath & ".ind"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

    Open ArchivoAbrir For Binary As #n
    'Escribimos la cabecera
    Put #n, , MiCabecera
    'Guardamos las cabezas

    Dim MisCuerpos()     As tIndiceCuerpo

    Dim MisCuerposLong() As tIndiceCuerpoLong

    ReDim MisCuerpos(0 To UBound(BodyData) + 1) As tIndiceCuerpo

    ReDim MisCuerposLong(0 To UBound(BodyData) + 1) As tIndiceCuerpoLong

    Put #n, , CInt(UBound(BodyData)) 'numheads

    If UsarGrhLong Then

        For i = 1 To UBound(BodyData)
            MisCuerposLong(i).Body(1) = BodyData(i).Walk(1).GrhIndex
            MisCuerposLong(i).Body(2) = BodyData(i).Walk(2).GrhIndex
            MisCuerposLong(i).Body(3) = BodyData(i).Walk(3).GrhIndex
            MisCuerposLong(i).Body(4) = BodyData(i).Walk(4).GrhIndex
            MisCuerposLong(i).HeadOffsetX = BodyData(i).HeadOffset.X
            MisCuerposLong(i).HeadOffsetY = BodyData(i).HeadOffset.Y
            Put #n, , MisCuerpos(i)
        Next i

    Else

        For i = 1 To UBound(BodyData)
            MisCuerpos(i).Body(1) = BodyData(i).Walk(1).GrhIndex
            MisCuerpos(i).Body(2) = BodyData(i).Walk(2).GrhIndex
            MisCuerpos(i).Body(3) = BodyData(i).Walk(3).GrhIndex
            MisCuerpos(i).Body(4) = BodyData(i).Walk(4).GrhIndex
            MisCuerpos(i).HeadOffsetX = BodyData(i).HeadOffset.X
            MisCuerpos(i).HeadOffsetY = BodyData(i).HeadOffset.Y
            Put #n, , MisCuerpos(i)
        Next i

    End If

    Close #n

    frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

    EstadoNoGuardado(e_EstadoIndexador.Body) = False

    Exit Sub
ErrHandler:
    Call MsgBox("Error en cuerpo " & i & " . " & Err.Description)

End Sub

Public Sub GuardarBodysDat(Optional ByVal FileNamePath As String = vbNullString)

    On Error GoTo ErrHandler

    Dim n As Integer, i As Integer

    n = FreeFile

    Dim ArchivoAbrir As String

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\encode\Body.dat"
        Else
            ArchivoAbrir = App.Path & "\encode\Body" & SavePath & ".dat"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

    Call WriteVar(ArchivoAbrir, "INIT", "NumBodies", CInt(UBound(BodyData))) 'numheads

    For i = 1 To UBound(BodyData)

        If BodyData(i).Walk(1).GrhIndex > 0 Then
            Call WriteVar(ArchivoAbrir, "Body" & i, "WALK1", BodyData(i).Walk(1).GrhIndex)
            Call WriteVar(ArchivoAbrir, "Body" & i, "WALK2", BodyData(i).Walk(2).GrhIndex)
            Call WriteVar(ArchivoAbrir, "Body" & i, "WALK3", BodyData(i).Walk(3).GrhIndex)
            Call WriteVar(ArchivoAbrir, "Body" & i, "WALK4", BodyData(i).Walk(4).GrhIndex)
            Call WriteVar(ArchivoAbrir, "Body" & i, "HeadOffsetX", BodyData(i).HeadOffset.X)
            Call WriteVar(ArchivoAbrir, "Body" & i, "HeadOffsety", BodyData(i).HeadOffset.Y)
            frmMain.LUlitError.Caption = "body : " & i
            DoEvents

        End If

    Next i

    frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir
    Exit Sub

    EstadoNoGuardado(e_EstadoIndexador.Body) = False

ErrHandler:
    Call MsgBox("Error en cuerpo " & i & " . " & Err.Description)

End Sub

Public Sub GuardarCascos(Optional ByVal FileNamePath As String = vbNullString)

    On Error GoTo ErrHandler

    Dim n As Integer, i As Integer

    n = FreeFile

    Dim ArchivoAbrir As String

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Cascos.ind"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Cascos" & SavePath & ".ind"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

    Open ArchivoAbrir For Binary As #n
    'Escribimos la cabecera
    Put #n, , MiCabecera
    'Guardamos las cabezas

    Dim Miscabezas() As tIndiceCabeza

    ReDim Miscabezas(0 To UBound(CascoAnimData) + 1) As tIndiceCabeza
    ReDim MiscabezasLong(0 To UBound(CascoAnimData) + 1) As tIndiceCabezaLong

    Put #n, , CInt(UBound(CascoAnimData)) 'numheads

    If UsarGrhLong Then

        For i = 1 To UBound(CascoAnimData)
            MiscabezasLong(i).Head(1) = CascoAnimData(i).Head(1).GrhIndex
            MiscabezasLong(i).Head(2) = CascoAnimData(i).Head(2).GrhIndex
            MiscabezasLong(i).Head(3) = CascoAnimData(i).Head(3).GrhIndex
            MiscabezasLong(i).Head(4) = CascoAnimData(i).Head(4).GrhIndex
            Put #n, , MiscabezasLong(i)
        Next i

    Else
    
        For i = 1 To UBound(CascoAnimData)
            Miscabezas(i).Head(1) = CascoAnimData(i).Head(1).GrhIndex
            Miscabezas(i).Head(2) = CascoAnimData(i).Head(2).GrhIndex
            Miscabezas(i).Head(3) = CascoAnimData(i).Head(3).GrhIndex
            Miscabezas(i).Head(4) = CascoAnimData(i).Head(4).GrhIndex
            Put #n, , Miscabezas(i)
        Next i

    End If

    Close #n

    frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

    EstadoNoGuardado(e_EstadoIndexador.Cascos) = False

    Exit Sub
ErrHandler:
    Call MsgBox("Error en casco " & i)

End Sub

Public Sub GuardarCascosDat(Optional ByVal FileNamePath As String = vbNullString)

    On Error GoTo ErrHandler

    Dim n As Integer, i As Integer

    n = FreeFile

    Dim ArchivoAbrir As String

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\encode\Cascos.dat"
        Else
            ArchivoAbrir = App.Path & "\encode\Cascos" & SavePath & ".dat"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

    Call WriteVar(ArchivoAbrir, "INIT", "NumCascos", CInt(UBound(CascoAnimData)))

    For i = 1 To UBound(CascoAnimData)

        If CascoAnimData(i).Head(1).GrhIndex > 0 Then
            Call WriteVar(ArchivoAbrir, "Casco" & i, "Head1", CascoAnimData(i).Head(1).GrhIndex)
            Call WriteVar(ArchivoAbrir, "Casco" & i, "Head2", CascoAnimData(i).Head(2).GrhIndex)
            Call WriteVar(ArchivoAbrir, "Casco" & i, "Head3", CascoAnimData(i).Head(3).GrhIndex)
            Call WriteVar(ArchivoAbrir, "Casco" & i, "Head4", CascoAnimData(i).Head(4).GrhIndex)

        End If

    Next i

    frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

    EstadoNoGuardado(e_EstadoIndexador.Cascos) = False

    Exit Sub
ErrHandler:
    Call MsgBox("Error en casco " & i)

End Sub

Public Sub GuardarArmas(Optional ByVal FileNamePath As String = vbNullString)

    On Error GoTo ErrHandler

    Dim Narchivo     As String

    Dim n            As Integer, i As Integer

    Dim ArchivoAbrir As String

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\armas.dat"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\armas" & SavePath & ".dat"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

    Narchivo = ArchivoAbrir
    Call WriteVar(Narchivo, "INIT", "NumArmas", UBound(WeaponAnimData))

    For i = 1 To UBound(WeaponAnimData)

        If WeaponAnimData(i).WeaponWalk(1).GrhIndex > 0 Then
            Call WriteVar(Narchivo, "ARMA" & i, "Dir1", WeaponAnimData(i).WeaponWalk(1).GrhIndex)
            Call WriteVar(Narchivo, "ARMA" & i, "Dir2", WeaponAnimData(i).WeaponWalk(2).GrhIndex)
            Call WriteVar(Narchivo, "ARMA" & i, "Dir3", WeaponAnimData(i).WeaponWalk(3).GrhIndex)
            Call WriteVar(Narchivo, "ARMA" & i, "Dir4", WeaponAnimData(i).WeaponWalk(4).GrhIndex)

        End If

    Next i

    frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

    EstadoNoGuardado(e_EstadoIndexador.Armas) = False

    Exit Sub
ErrHandler:
    Call MsgBox("Error en arma " & i)

End Sub

Public Sub GuardarEscudos(Optional ByVal FileNamePath As String = vbNullString)

    On Error GoTo ErrHandler

    Dim Narchivo     As String

    Dim n            As Integer, i As Integer

    Dim ArchivoAbrir As String

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Escudos.dat"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Escudos" & SavePath & ".dat"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

    Narchivo = ArchivoAbrir
    Call WriteVar(Narchivo, "INIT", "NumEscudos", UBound(ShieldAnimData))

    For i = 1 To UBound(ShieldAnimData)

        If ShieldAnimData(i).ShieldWalk(1).GrhIndex > 0 Then
            Call WriteVar(Narchivo, "ESC" & i, "Dir1", ShieldAnimData(i).ShieldWalk(1).GrhIndex)
            Call WriteVar(Narchivo, "ESC" & i, "Dir2", ShieldAnimData(i).ShieldWalk(2).GrhIndex)
            Call WriteVar(Narchivo, "ESC" & i, "Dir3", ShieldAnimData(i).ShieldWalk(3).GrhIndex)
            Call WriteVar(Narchivo, "ESC" & i, "Dir4", ShieldAnimData(i).ShieldWalk(4).GrhIndex)

        End If

    Next i

    frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

    EstadoNoGuardado(e_EstadoIndexador.Escudos) = False

    Exit Sub
ErrHandler:
    Call MsgBox("Error en escudo " & i)

End Sub

'********************************************************************************
'********************************************************************************
'********************************************************************************
'******************************** Botones ***************************************
'********************************************************************************
'********************************************************************************
'********************************************************************************

Public Sub BotonGuardado(Optional ByVal FileNamePath As String = vbNullString)

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh
            Call SaveGrhData(FileNamePath)

        Case e_EstadoIndexador.Body
            Call GuardarBodys(FileNamePath)

        Case e_EstadoIndexador.Cabezas
            Call GuardarCabezas(FileNamePath)

        Case e_EstadoIndexador.Cascos
            Call GuardarCascos(FileNamePath)

        Case e_EstadoIndexador.Escudos
            Call GuardarEscudos(FileNamePath)

        Case e_EstadoIndexador.Armas
            Call GuardarArmas(FileNamePath)

        Case e_EstadoIndexador.Alas
            Call GuardarAlas(FileNamePath)

        Case e_EstadoIndexador.Botas
            Call GuardarBotas(FileNamePath)

        Case e_EstadoIndexador.Capas
            Call GuardarCapas(FileNamePath)

        Case e_EstadoIndexador.Fx
            Call GuardarFxs(FileNamePath)

    End Select

End Sub

Public Sub BotonGuardadoDat(Optional ByVal FileNamePath As String = vbNullString)

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh
            Call SaveGrhDataDat(FileNamePath)

        Case e_EstadoIndexador.Body
            Call GuardarBodysDat(FileNamePath)

        Case e_EstadoIndexador.Cabezas
            Call GuardarCabezasDat(FileNamePath)

        Case e_EstadoIndexador.Cascos
            Call GuardarCascosDat(FileNamePath)

        Case e_EstadoIndexador.Escudos
            Call GuardarEscudos(FileNamePath)

        Case e_EstadoIndexador.Armas
            Call GuardarArmas(FileNamePath)

        Case e_EstadoIndexador.Alas
            Call GuardarAlasDat(FileNamePath)

        Case e_EstadoIndexador.Botas
            Call GuardarBotasDat(FileNamePath)

        Case e_EstadoIndexador.Capas
            Call GuardarCapasDat(FileNamePath)

        Case e_EstadoIndexador.Fx
            Call GuardarFxsDat(FileNamePath)

    End Select

End Sub

Public Sub BotonCargado(Optional ByVal FileNamePath As String = vbNullString)

    Dim respuesta As Byte

    Dim tempLong  As Long

    respuesta = MsgBox("ATENCION Si contunias perderas los cambios no guardados", 4, "¡¡ADVERTENCIA!!")

    If respuesta <> vbYes Then
        Exit Sub

    End If
        
    frmMain.Visor.Cls

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh
            Call LoadGrhData(FileNamePath)
            Call RenuevaListaGrH

        Case e_EstadoIndexador.Body
            Call CargarCuerpos(FileNamePath)
            Call RenuevaListaBodys

        Case e_EstadoIndexador.Cabezas
            Call CargarCabezas(FileNamePath)
            Call RenuevaListaCabezas

        Case e_EstadoIndexador.Cascos
            Call CargarCascos(FileNamePath)
            Call RenuevaListaCascos

        Case e_EstadoIndexador.Escudos
            Call CargarAnimEscudos(FileNamePath)
            Call RenuevaListaEscudos

        Case e_EstadoIndexador.Armas
            Call CargarAnimArmas(FileNamePath)
            Call RenuevaListaArmas

        Case e_EstadoIndexador.Alas
            Call CargarAlas(FileNamePath)
            Call RenuevaListaAlas

        Case e_EstadoIndexador.Capas
            Call CargarEspalda(FileNamePath)
            Call RenuevaListaCapas

        Case e_EstadoIndexador.Fx
            Call CargarFxs(FileNamePath)
            Call RenuevaListaFX

        Case e_EstadoIndexador.Botas
            Call CargarBotas(FileNamePath)
            Call RenuevaListaBotas

    End Select

    If EstadoIndexador = e_EstadoIndexador.Grh Then
        tempLong = ListaindexGrH(GRHActual)

        If tempLong >= frmMain.Lista.ListCount Then tempLong = 0
        frmMain.Lista.listIndex = tempLong
    Else
        tempLong = ListaindexGrH(DataIndexActual)

        If tempLong >= frmMain.Lista.ListCount Then tempLong = 0
        frmMain.Lista.listIndex = tempLong

    End If

End Sub

Public Sub BotonCargadoDat(Optional ByVal FileNamePath As String = vbNullString)

    Dim respuesta As Byte

    Dim tempLong  As Long

    respuesta = MsgBox("ATENCION Si contunias perderas los cambios no guardados", 4, "¡¡ADVERTENCIA!!")

    If respuesta <> vbYes Then
        Exit Sub

    End If
        
    frmMain.Visor.Cls

    Select Case EstadoIndexador

        Case e_EstadoIndexador.Grh
            Call LoadGrhDataDat(FileNamePath)
            Call RenuevaListaGrH

        Case e_EstadoIndexador.Body
            Call CargarCuerposDat(FileNamePath)
            Call RenuevaListaBodys

        Case e_EstadoIndexador.Cabezas
            Call CargarCabezasdat(FileNamePath)
            Call RenuevaListaCabezas

        Case e_EstadoIndexador.Cascos
            Call CargarCascosDat(FileNamePath)
            Call RenuevaListaCascos

        Case e_EstadoIndexador.Escudos
            Call CargarAnimEscudos(FileNamePath)
            Call RenuevaListaEscudos

        Case e_EstadoIndexador.Armas
            Call CargarAnimArmas(FileNamePath)
            Call RenuevaListaArmas

        Case e_EstadoIndexador.Alas
            Call CargarAlasDat(FileNamePath)
            Call RenuevaListaAlas

        Case e_EstadoIndexador.Capas
            Call CargarEspaldaDat(FileNamePath)
            Call RenuevaListaCapas

        Case e_EstadoIndexador.Fx
            Call CargarFxsDat(FileNamePath)
            Call RenuevaListaFX

        Case e_EstadoIndexador.Botas
            Call CargarBotasDat(FileNamePath)
            Call RenuevaListaBotas

    End Select

    If EstadoIndexador = e_EstadoIndexador.Grh Then
        tempLong = ListaindexGrH(GRHActual)
        frmMain.Lista.listIndex = tempLong
    Else
        tempLong = ListaindexGrH(DataIndexActual)
        frmMain.Lista.listIndex = tempLong

    End If

End Sub

