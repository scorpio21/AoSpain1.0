Attribute VB_Name = "Indices"
Option Explicit

Public CascoSData() As tIndiceCabeza

Public CapasData()  As tIndiceCabeza

'Public BotasData() As tIndiceCabeza
Public headataI()   As tIndiceCabeza

Public Mapas()      As Byte

Public CuerpoData() As tIndiceCuerpo

Public FxDataI()    As tIndiceFx

Public Numheads     As Integer

Public NumCascos    As Integer

'Public NumBotas As Integer
Public Numcapas     As Integer

Public NumCuerpos   As Integer

Public NumTips      As Integer

Public NumMapas     As Integer

Public Declare Function getprivateprofilestring _
               Lib "kernel32" _
               Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                 ByVal lpKeyname As Any, _
                                                 ByVal lpdefault As String, _
                                                 ByVal lpreturnedstring As String, _
                                                 ByVal nSize As Long, _
                                                 ByVal lpfilename As String) As Long

Function GetVar(file As String, Main As String, Var As String) As String
    '*****************************************************************
    'Gets a Var from a text file
    '*****************************************************************

    Dim l        As Integer

    Dim Char     As String

    Dim sSpaces  As String ' This will hold the input that the program will retrieve

    Dim szReturn As String ' This will be the defaul value if the string is not found

    szReturn = ""

    sSpaces = Space(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish

    getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), file

    GetVar = RTrim(sSpaces)
    GetVar = Left(GetVar, Len(GetVar) - 1)

End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")

End Function

Sub CargarAnimArmas(Optional ByVal FileNamePath As String = vbNullString)

    On Error Resume Next

    Dim loopc        As Long

    Dim arch         As String

    Dim ArchivoAbrir As String
    
    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\" & "armas.dat"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\" & "armas" & SavePath & ".dat"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If
    
    If Not FileExist(ArchivoAbrir, vbNormal) Then
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"

        If UBound(WeaponAnimData()) = 0 Then
            ReDim WeaponAnimData(1) As WeaponAnimData

        End If

        Exit Sub

    End If
    
    arch = ArchivoAbrir
    
    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopc = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopc, "Dir1")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopc, "Dir2")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopc, "Dir3")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopc, "Dir4")), 0
    Next loopc

End Sub

Sub CargarAnimEscudos(Optional ByVal FileNamePath As String = vbNullString)

    On Error Resume Next

    Dim loopc        As Long

    Dim arch         As String

    Dim ArchivoAbrir As String

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\" & "escudos.dat"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\" & "escudos" & SavePath & ".dat"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If
    
    If Not FileExist(ArchivoAbrir, vbNormal) Then
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"

        If UBound(ShieldAnimData()) = 0 Then
            ReDim ShieldAnimData(1) As ShieldAnimData

        End If

        Exit Sub

    End If
    
    arch = ArchivoAbrir
    
    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopc, "Dir1")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopc, "Dir2")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopc, "Dir3")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopc, "Dir4")), 0
    Next loopc

End Sub

Public Function ReadField(ByVal Pos As Integer, _
                          ByVal Text As String, _
                          ByVal SepASCII As Integer) As String

    '*****************************************************************
    'Gets a field from a string
    '*****************************************************************
    Dim i         As Integer

    Dim LastPos   As Integer

    Dim CurChar   As String * 1

    Dim FieldNum  As Integer

    Dim Seperator As String
    
    Seperator = Chr$(SepASCII)
    LastPos = 0
    FieldNum = 0
    
    For i = 1 To Len(Text)
        CurChar = mid$(Text, i, 1)

        If CurChar = Seperator Then
            FieldNum = FieldNum + 1

            If FieldNum = Pos Then
                ReadField = mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
                Exit Function

            End If

            LastPos = i

        End If

    Next i

    FieldNum = FieldNum + 1
    
    If FieldNum = Pos Then
        ReadField = mid$(Text, LastPos + 1)

    End If

End Function

Sub LoadGrhDataDat(Optional ByVal FileNamePath As String = vbNullString)
    '*****************************************************************
    'Loads Grh.dat
    '*****************************************************************

    On Error GoTo ErrorHandler

    Dim Grh          As Integer

    Dim Frame        As Integer

    Dim TempInt      As Integer

    Dim ArchivoAbrir As String

    Dim StringGrh    As String

    'Resize arrays
    ReDim Grhdata(1 To MAXGrH) As Grhdata

    'Open files

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = IniPath & "Graficos.dat"
        Else
            ArchivoAbrir = IniPath & "Graficos" & SavePath & ".dat"

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

    Do Until Grh > MAXGrH
    
        'Get number of frames
        StringGrh = Leer.GetValue("Graphics", "Grh" & Grh)

        If StringGrh <> vbNullString Then
        
            Grhdata(Grh).NumFrames = Val(ReadField(1, StringGrh, Asc("-")))
    
            If Grhdata(Grh).NumFrames <= 0 Then GoTo ErrorHandler
        
            If Grhdata(Grh).NumFrames > 1 Then

                'Read a animation GRH set
                For Frame = 1 To Grhdata(Grh).NumFrames
                    Grhdata(Grh).Frames(Frame) = Val(ReadField(1 + Frame, StringGrh, Asc("-")))

                    If Grhdata(Grh).Frames(Frame) <= 0 Or Grhdata(Grh).Frames(Frame) > MAXGrH Then
                        GoTo ErrorHandler

                    End If
            
                Next Frame
        
                Grhdata(Grh).Speed = Val(ReadField(1 + Frame, StringGrh, Asc("-")))

                'Compute width and height
                Grhdata(Grh).pixelHeight = Grhdata(Grhdata(Grh).Frames(1)).pixelHeight
            
                Grhdata(Grh).pixelWidth = Grhdata(Grhdata(Grh).Frames(1)).pixelWidth
            
                Grhdata(Grh).TileWidth = Grhdata(Grhdata(Grh).Frames(1)).TileWidth
            
                Grhdata(Grh).TileHeight = Grhdata(Grhdata(Grh).Frames(1)).TileHeight
        
            Else
                'Read in normal GRH data
                Grhdata(Grh).FileNum = Val(ReadField(2, StringGrh, Asc("-")))

                If Grhdata(Grh).FileNum <= 0 Then GoTo ErrorHandler
    
                Grhdata(Grh).sX = Val(ReadField(3, StringGrh, Asc("-")))
            
                Grhdata(Grh).sY = Val(ReadField(4, StringGrh, Asc("-")))
                
                Grhdata(Grh).pixelWidth = Val(ReadField(5, StringGrh, Asc("-")))

                Grhdata(Grh).pixelHeight = Val(ReadField(6, StringGrh, Asc("-")))
            
                'Compute width and height
                Grhdata(Grh).TileWidth = Grhdata(Grh).pixelWidth / TilePixelHeight
                Grhdata(Grh).TileHeight = Grhdata(Grh).pixelHeight / TilePixelWidth
            
                Grhdata(Grh).Frames(1) = Grh
                
            End If
        
        End If

        'Get Next Grh Number
        Grh = Grh + 1

    Loop
    '************************************************
    Set Leer = Nothing

    Exit Sub

ErrorHandler:
    Close #1
    MsgBox "Error while loading " & ArchivoAbrir & " Stopped at GRH number: " & Grh

End Sub

Sub SaveGrhData(Optional ByVal FileNamePath As String = vbNullString)
    '*****************************************************************
    'Loads Grh.dat
    '*****************************************************************

    On Error GoTo ErrorHandler

    Dim Grh          As Integer

    Dim Frame        As Integer

    Dim TempInt      As Integer

    Dim n            As Integer

    Dim ArchivoAbrir As String

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = IniPath & "Graficos.ind"
        Else
            ArchivoAbrir = IniPath & "Graficos" & SavePath & ".ind"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

    n = FreeFile
Call IniciarCabecera(MiCabecera)
    'Open files
    Open ArchivoAbrir For Binary As #n
    Seek #1, 1

    Put #n, , MiCabecera
    Put #1, , TempInt
    Put #1, , TempInt
    Put #1, , TempInt
    Put #1, , TempInt
    Put #1, , TempInt

    'Fill Grh List
    For Grh = 1 To MAXGrH

        If Grhdata(Grh).NumFrames <= 0 Then GoTo aqui2
        'Get first Grh Number
    
        If UsarGrhLong Then
            Put #1, , Grh
        Else
            TempInt = Grh
            Put #1, , TempInt

        End If

        'Get number of frames
        Put #1, , Grhdata(Grh).NumFrames
    
        If Grhdata(Grh).NumFrames > 1 Then

            'Read a animation GRH set
            For Frame = 1 To Grhdata(Grh).NumFrames

                If UsarGrhLong Then
                    Put #1, , Grhdata(Grh).Frames(Frame)
                Else
                    TempInt = Grhdata(Grh).Frames(Frame)
                    Put #1, , TempInt

                End If
            
                If Grhdata(Grh).Frames(Frame) <= 0 Or Grhdata(Grh).Frames(Frame) > MAXGrH Then
                    frmMain.LUlitError.Caption = frmMain.LUlitError.Caption & " .frame incorrecto(" & Grh & ")"

                End If
        
            Next Frame
        
            If Grhdata(Grh).Speed <= 0 Then Grhdata(Grh).Speed = 1
            Put #1, , Grhdata(Grh).Speed
        
            'Compute width and height
            Grhdata(Grh).pixelHeight = Grhdata(Grhdata(Grh).Frames(1)).pixelHeight

            If Grhdata(Grh).pixelHeight <= 0 Then frmMain.LUlitError.Caption = frmMain.LUlitError.Caption & " .Alto incorrecto(" & Grh & ")"
        
            Grhdata(Grh).pixelWidth = Grhdata(Grhdata(Grh).Frames(1)).pixelWidth

            If Grhdata(Grh).pixelWidth <= 0 Then frmMain.LUlitError.Caption = frmMain.LUlitError.Caption & " .Ancho incorrecto(" & Grh & ")"
        
            Grhdata(Grh).TileWidth = Grhdata(Grhdata(Grh).Frames(1)).TileWidth
            'If Grhdata(Grh).TileWidth <= 0 Then frmMain.LUlitError.Caption = frmMain.LUlitError.Caption & " .Ancho incorrecto(" & Grh & ")"
        
            Grhdata(Grh).TileHeight = Grhdata(Grhdata(Grh).Frames(1)).TileHeight
            'If Grhdata(Grh).TileHeight <= 0 Then frmMain.LUlitError.Caption = frmMain.LUlitError.Caption & " .Ancho incorrecto(" & Grh & ")"
    
        Else
            'Read in normal GRH data
            Put #1, , Grhdata(Grh).FileNum

            If Grhdata(Grh).FileNum <= 0 Then frmMain.LUlitError.Caption = frmMain.LUlitError.Caption & " .Bmp incorrecto(" & Grh & ")"
        
            Put #1, , Grhdata(Grh).sX
            'If Grhdata(Grh).sX < 0 Then GoTo ErrorHandler
        
            Put #1, , Grhdata(Grh).sY
            'If Grhdata(Grh).sY < 0 Then GoTo ErrorHandler
            
            Put #1, , Grhdata(Grh).pixelWidth

            If Grhdata(Grh).pixelWidth <= 0 Then frmMain.LUlitError.Caption = frmMain.LUlitError.Caption & " .Ancho incorrecto(" & Grh & ")"
        
            Put #1, , Grhdata(Grh).pixelHeight

            If Grhdata(Grh).pixelHeight <= 0 Then frmMain.LUlitError.Caption = frmMain.LUlitError.Caption & " .Alto incorrecto(" & Grh & ")"
        
            'Compute width and height
            Grhdata(Grh).TileWidth = Grhdata(Grh).pixelWidth / TilePixelHeight
            Grhdata(Grh).TileHeight = Grhdata(Grh).pixelHeight / TilePixelWidth
        
            Grhdata(Grh).Frames(1) = Grh
            
        End If

        'Get Next Grh Number
aqui2:
    Next Grh

    '************************************************

    Close #1

    EstadoNoGuardado(e_EstadoIndexador.Grh) = False
    frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir
    Exit Sub

ErrorHandler:
    Close #1
    MsgBox "Error while saving the " & ArchivoAbrir & " ! Stopped at GRH number: " & Grh

End Sub

Sub SaveGrhDataDat(Optional ByVal FileNamePath As String = vbNullString)
    '*****************************************************************
    'Loads Grh.dat
    '*****************************************************************

    On Error GoTo ErrorHandler

    Dim Grh          As Integer

    Dim Frame        As Integer

    Dim TempInt      As Integer

    Dim ArchivoAbrir As String

    Dim StringGrh    As String

    Dim LastGrh      As Long

    Dim TotalString  As String

    'Resize arrays

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = IniPath & "Graficos.dat"
        Else
            ArchivoAbrir = IniPath & "Graficos" & SavePath & ".dat"

        End If

    Else
        ArchivoAbrir = FileNamePath

    End If

    If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

    'TotalString = "[Graphics]" & vbCrLf & vbCrLf
    Grh = 1

    Do Until Grh > MAXGrH
    
        'Get number of frames
        If Grhdata(Grh).NumFrames >= 1 Then
            StringGrh = Grhdata(Grh).NumFrames & "-"

            If Grhdata(Grh).NumFrames > 1 Then

                'Read a animation GRH set
                For Frame = 1 To Grhdata(Grh).NumFrames
                    StringGrh = StringGrh & Grhdata(Grh).Frames(Frame) & "-"

                    If Grhdata(Grh).Frames(Frame) <= 0 Or Grhdata(Grh).Frames(Frame) > MAXGrH Then
                        GoTo ErrorHandler

                    End If
            
                Next Frame
        
                StringGrh = StringGrh & Grhdata(Grh).Speed

                If Grhdata(Grh).Speed <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                StringGrh = StringGrh & Grhdata(Grh).FileNum & "-"
    
                StringGrh = StringGrh & Grhdata(Grh).sX & "-"
        
                StringGrh = StringGrh & Grhdata(Grh).sY & "-"
                
                StringGrh = StringGrh & Grhdata(Grh).pixelWidth & "-"
            
                StringGrh = StringGrh & Grhdata(Grh).pixelHeight

            End If

            Call WriteVar(ArchivoAbrir, "Graphics", "Grh" & Grh, StringGrh)
            'TotalString = TotalString & "Grh" & Grh & "=" & StringGrh & vbCrLf
            LastGrh = Grh
            DoEvents

        End If

        'Get Next Grh Number
        Grh = Grh + 1

        frmMain.LUlitError.Caption = "Grh: " & Grh
    Loop
    '************************************************
    Call WriteVar(ArchivoAbrir, "INIT", "NumGrh", LastGrh)
    'TotalString = TotalString & vbCrLf & "[INIT]" & vbCrLf & "numGRH" & "=" & LastGrh

    '    Dim N As Integer
    '    N = FreeFile
    
    '    Open ArchivoAbrir For Binary As #N
    '        Put #N, , TotalString
    '    Close #N

    EstadoNoGuardado(e_EstadoIndexador.Grh) = False
    frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir
    Exit Sub

ErrorHandler:

    MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh

End Sub

Public Function ListaindexGrH(ByVal numGRH As Integer) As Integer

    Dim i As Long

    ListaindexGrH = -1

    For i = 0 To frmMain.Lista.ListCount

        If numGRH = Val(ReadField(1, frmMain.Lista.List(i), Asc(" "))) Then
            ListaindexGrH = i
            Exit Function

        End If

    Next i

End Function

Public Function ComprobarSobreescribir(ByVal ArchivoAbrir As String) As Boolean
    ' Comprueba si el archvo existe y advierte de sobreescritura. Si se acepta ya lo borra

    If FileExist(ArchivoAbrir, vbArchive) Then

        Dim respuesta As Byte

        respuesta = MsgBox("ATENCION Si contunias sobrescribiras el archivo existente" & vbCrLf & ArchivoAbrir, 4, "��ADVERTENCIA!!")

        If respuesta <> vbYes Then
            ComprobarSobreescribir = False
            Exit Function

        End If

        Kill ArchivoAbrir

    End If

    ComprobarSobreescribir = True

End Function

Public Sub ComprobarIndexLista()

    If UltimoindexE(EstadoIndexador) < 0 Then
        If UltimoindexE(EstadoIndexador) <> -1 Then
            frmMain.Lista.listIndex = 0
        Else
            frmMain.Lista.listIndex = -1

        End If

    ElseIf UltimoindexE(EstadoIndexador) >= frmMain.Lista.ListCount Then
        frmMain.Lista.listIndex = frmMain.Lista.ListCount - 1
    Else
        frmMain.Lista.listIndex = UltimoindexE(EstadoIndexador)

    End If

End Sub

Public Function BuscarGrHlibre() As Integer

    Dim i As Long

    For i = 1 To MAXGrH

        If Grhdata(i).NumFrames = 0 Then
            BuscarGrHlibre = i
            Exit Function

        End If

    Next i

End Function

Public Function BuscarGrHlibres(ByVal hTotales As Integer) As Integer

    Dim i       As Long

    Dim Primero As Integer

    Dim Cuenta  As Integer

    For i = 1 To MAXGrH

        If Cuenta = hTotales Then
            BuscarGrHlibres = Primero
            Exit Function

        End If

        If Grhdata(i).NumFrames = 0 Then
            If Primero = 0 Then
                Primero = i
                Cuenta = 1
            Else
                Cuenta = Cuenta + 1

            End If

        Else
            Cuenta = 0
            Primero = 0

        End If

    Next i

End Function

Public Function hayGrHlibres(ByVal Primero As Integer, _
                             ByVal hTotales As Integer) As Boolean

    Dim i      As Long

    Dim Cuenta As Integer

    If Primero <= 0 Or Primero > MAXGrH Then Exit Function

    For i = Primero To Primero + hTotales - 1

        If Grhdata(i).NumFrames > 0 Then
            hayGrHlibres = False
            Exit Function

        End If

    Next i

    hayGrHlibres = True

End Function

Public Sub AgregaGrH(ByVal numGRH As Integer)

    Dim i           As Long

    Dim EsteIndex   As Long

    Dim CuentaIndex As Long

    Grhdata(numGRH).FileNum = 1
    Grhdata(numGRH).NumFrames = 1
    Grhdata(numGRH).pixelHeight = 32
    Grhdata(numGRH).pixelWidth = 32
    Grhdata(numGRH).Frames(1) = numGRH

    CuentaIndex = -1
    frmMain.Lista.Clear

    For i = 1 To MAXGrH

        If Grhdata(i).NumFrames = 1 Then
            frmMain.Lista.AddItem i
            CuentaIndex = CuentaIndex + 1
        ElseIf Grhdata(i).NumFrames > 1 Then
            frmMain.Lista.AddItem i & " (animacion)"
            CuentaIndex = CuentaIndex + 1

        End If

        If i = numGRH Then
            EsteIndex = CuentaIndex

        End If

    Next i

    frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub AgregaBody(ByVal Numbody As Integer, _
                      Optional ByVal RefreshList As Boolean = True)

    Dim i           As Long

    Dim EsteIndex   As Long

    Dim CuentaIndex As Long

    If Numbody > UBound(BodyData) Then ReDim Preserve BodyData(0 To Numbody) As BodyData

    BodyData(Numbody).Walk(1).GrhIndex = 1

    If Not RefreshList Then Exit Sub

    CuentaIndex = -1
    frmMain.Lista.Clear

    For i = 1 To UBound(BodyData)

        If BodyData(i).Walk(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i
            CuentaIndex = CuentaIndex + 1

        End If

        If i = Numbody Then
            EsteIndex = CuentaIndex

        End If

    Next i

    frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub mueveBody(ByVal Numbody As Integer, _
                     ByVal origenBody As Integer, _
                     Optional ByVal BorrarOriginal As Boolean = True)

    Dim i           As Long

    Dim EsteIndex   As Long

    Dim CuentaIndex As Long

    Dim BodyVacio   As BodyData

    Dim respuesta   As Byte

    If Numbody > UBound(BodyData) Then ReDim Preserve BodyData(0 To Numbody) As BodyData
    If BodyData(Numbody).Walk(1).GrhIndex > 0 Then
        respuesta = MsgBox("El body " & Numbody & " ya existe, �deseas sobreescribirlo?", 4, "Aviso")

        If respuesta = vbYes Then
            BodyData(Numbody) = BodyData(origenBody)

            If BorrarOriginal Then BodyData(origenBody) = BodyVacio

        End If

    Else
        BodyData(Numbody) = BodyData(origenBody)

        If BorrarOriginal Then BodyData(origenBody) = BodyVacio

    End If

    CuentaIndex = -1
    frmMain.Lista.Clear

    For i = 1 To UBound(BodyData)

        If BodyData(i).Walk(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i
            CuentaIndex = CuentaIndex + 1

        End If

        If i = Numbody Then
            EsteIndex = CuentaIndex

        End If

    Next i

    frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub MueveCabeza(ByVal NumHead As Integer, _
                       ByVal origenHead As Integer, _
                       Optional ByVal BorrarOriginal As Boolean = True)

    Dim i           As Long

    Dim EsteIndex   As Long

    Dim CuentaIndex As Long

    Dim respuesta   As Byte

    Dim headVacia   As HeadData

    If NumHead > UBound(HeadData) Then ReDim Preserve HeadData(0 To NumHead) As HeadData
    If HeadData(NumHead).Head(1).GrhIndex > 0 Then
        respuesta = MsgBox("La cabeza " & NumHead & " ya existe, �deseas sobreescribirla?", 4, "Aviso")

        If respuesta = vbYes Then
            HeadData(NumHead) = HeadData(origenHead)

            If BorrarOriginal Then HeadData(origenHead) = headVacia

        End If

    Else
        HeadData(NumHead) = HeadData(origenHead)

        If BorrarOriginal Then HeadData(origenHead) = headVacia

    End If

    CuentaIndex = -1
    frmMain.Lista.Clear

    For i = 1 To UBound(HeadData)

        If HeadData(i).Head(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i
            CuentaIndex = CuentaIndex + 1

        End If

        If i = NumHead Then
            EsteIndex = CuentaIndex

        End If

    Next i

    frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub AgregaCabeza(ByVal NumHead As Integer)

    Dim i           As Long

    Dim EsteIndex   As Long

    Dim CuentaIndex As Long

    If NumHead > UBound(HeadData) Then ReDim Preserve HeadData(0 To NumHead) As HeadData

    HeadData(NumHead).Head(1).GrhIndex = 1

    CuentaIndex = -1
    frmMain.Lista.Clear

    For i = 1 To UBound(HeadData)

        If HeadData(i).Head(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i
            CuentaIndex = CuentaIndex + 1

        End If

        If i = NumHead Then
            EsteIndex = CuentaIndex

        End If

    Next i

    frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub AgregaCasco(ByVal NumCasco As Integer)

    Dim i           As Long

    Dim EsteIndex   As Long

    Dim CuentaIndex As Long

    If NumCasco > UBound(CascoAnimData) Then ReDim Preserve CascoAnimData(0 To NumCasco) As HeadData

    CascoAnimData(NumCasco).Head(1).GrhIndex = 1

    CuentaIndex = -1
    frmMain.Lista.Clear

    For i = 1 To UBound(CascoAnimData)

        If CascoAnimData(i).Head(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i
            CuentaIndex = CuentaIndex + 1

        End If

        If i = NumCasco Then
            EsteIndex = CuentaIndex

        End If

    Next i

    frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub MueveCasco(ByVal NumCasco As Integer, _
                      ByVal OrigenCasco As Integer, _
                      Optional ByVal BorrarOriginal As Boolean = True)

    Dim i           As Long

    Dim EsteIndex   As Long

    Dim CuentaIndex As Long

    Dim respuesta   As Byte

    Dim headVacia   As HeadData

    If NumCasco > UBound(CascoAnimData) Then ReDim Preserve CascoAnimData(0 To NumCasco) As HeadData

    If CascoAnimData(NumCasco).Head(1).GrhIndex > 0 Then
        respuesta = MsgBox("El casco " & NumCasco & " ya existe, �deseas sobreescribirlo?", 4, "Aviso")

        If respuesta = vbYes Then
            CascoAnimData(NumCasco) = CascoAnimData(OrigenCasco)

            If BorrarOriginal Then CascoAnimData(OrigenCasco) = headVacia

        End If

    Else
        CascoAnimData(NumCasco) = CascoAnimData(OrigenCasco)

        If BorrarOriginal Then CascoAnimData(OrigenCasco) = headVacia

    End If

    CuentaIndex = -1
    frmMain.Lista.Clear

    For i = 1 To UBound(CascoAnimData)

        If CascoAnimData(i).Head(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i
            CuentaIndex = CuentaIndex + 1

        End If

        If i = NumCasco Then
            EsteIndex = CuentaIndex

        End If

    Next i

    frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub MueveEscudo(ByVal NumEscudo As Integer, _
                       ByVal origenEscudo As Integer, _
                       Optional ByVal BorrarOriginal As Boolean = True)

    Dim i           As Long

    Dim EsteIndex   As Long

    Dim CuentaIndex As Long

    Dim respuesta   As Byte

    Dim escudoVacio As ShieldAnimData

    escudoVacio.ShieldWalk(1).GrhIndex = 0

    If NumEscudo > UBound(ShieldAnimData) Then ReDim Preserve ShieldAnimData(1 To NumEscudo) As ShieldAnimData

    If ShieldAnimData(NumEscudo).ShieldWalk(1).GrhIndex > 0 Then
        respuesta = MsgBox("El escudo " & NumEscudo & " ya existe, �deseas sobreescribirlo?", 4, "Aviso")

        If respuesta = vbYes Then
            ShieldAnimData(NumEscudo) = ShieldAnimData(origenEscudo)

            If BorrarOriginal Then ShieldAnimData(origenEscudo) = escudoVacio

        End If

    Else
        ShieldAnimData(NumEscudo) = ShieldAnimData(origenEscudo)

        If BorrarOriginal Then ShieldAnimData(origenEscudo) = escudoVacio

    End If

    CuentaIndex = -1
    frmMain.Lista.Clear

    For i = 1 To UBound(ShieldAnimData)

        If ShieldAnimData(i).ShieldWalk(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i
            CuentaIndex = CuentaIndex + 1

        End If

        If i = NumEscudo Then
            EsteIndex = CuentaIndex

        End If

    Next i

    frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub AgregaEscudo(ByVal NumEscudo As Integer, _
                        Optional ByVal RefreshList As Boolean = True)

    Dim i           As Long

    Dim EsteIndex   As Long

    Dim CuentaIndex As Long

    If NumEscudo > UBound(ShieldAnimData) Then ReDim Preserve ShieldAnimData(1 To NumEscudo) As ShieldAnimData

    ShieldAnimData(NumEscudo).ShieldWalk(1).GrhIndex = 1

    If Not RefreshList Then Exit Sub
    CuentaIndex = -1
    frmMain.Lista.Clear

    For i = 1 To UBound(ShieldAnimData)

        If ShieldAnimData(i).ShieldWalk(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i
            CuentaIndex = CuentaIndex + 1

        End If

        If i = NumEscudo Then
            EsteIndex = CuentaIndex

        End If

    Next i

    frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub AgregaArma(ByVal NumArma As Integer, _
                      Optional ByVal RefreshList As Boolean = True)

    Dim i           As Long

    Dim EsteIndex   As Long

    Dim CuentaIndex As Long

    If NumArma > UBound(WeaponAnimData) Then ReDim Preserve WeaponAnimData(1 To NumArma) As WeaponAnimData

    WeaponAnimData(NumArma).WeaponWalk(1).GrhIndex = 1

    If Not RefreshList Then Exit Sub
    CuentaIndex = -1
    frmMain.Lista.Clear

    For i = 1 To UBound(WeaponAnimData)

        If WeaponAnimData(i).WeaponWalk(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i
            CuentaIndex = CuentaIndex + 1

        End If

        If i = NumArma Then
            EsteIndex = CuentaIndex

        End If

    Next i

    frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub MueveArma(ByVal NumArma As Integer, _
                     ByVal OrigenArma As Integer, _
                     Optional ByVal BorrarOriginal As Boolean = True)

    Dim i           As Long

    Dim EsteIndex   As Long

    Dim CuentaIndex As Long

    Dim respuesta   As Byte

    Dim armaVacia   As WeaponAnimData

    armaVacia.WeaponWalk(1).GrhIndex = 0

    If NumArma > UBound(WeaponAnimData) Then ReDim Preserve WeaponAnimData(1 To NumArma) As WeaponAnimData

    If WeaponAnimData(NumArma).WeaponWalk(1).GrhIndex > 0 Then
        respuesta = MsgBox("El arma " & NumArma & " ya existe, �deseas sobreescribirla?", 4, "Aviso")

        If respuesta = vbYes Then
            WeaponAnimData(NumArma) = WeaponAnimData(OrigenArma)

            If BorrarOriginal Then WeaponAnimData(OrigenArma) = armaVacia

        End If

    Else
        WeaponAnimData(NumArma) = WeaponAnimData(OrigenArma)

        If BorrarOriginal Then WeaponAnimData(OrigenArma) = armaVacia

    End If

    CuentaIndex = -1
    frmMain.Lista.Clear

    For i = 1 To UBound(WeaponAnimData)

        If WeaponAnimData(i).WeaponWalk(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i
            CuentaIndex = CuentaIndex + 1

        End If

        If i = NumArma Then
            EsteIndex = CuentaIndex

        End If

    Next i

    frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub MueveBota(ByVal NumBotas As Integer, _
                     ByVal OrigenBota As Integer, _
                     Optional ByVal BorrarOriginal As Boolean = True)

    Dim i           As Long

    Dim EsteIndex   As Long

    Dim CuentaIndex As Long

    Dim BotaVacio   As BotaData

    Dim respuesta   As Byte

    If NumBotas > UBound(BotaData) Then ReDim Preserve BotaData(0 To NumBotas) As BotaData
    If BotaData(NumBotas).Bota(1).GrhIndex > 0 Then
        respuesta = MsgBox("La bota " & NumBotas & " ya existe, �deseas sobreescribirlo?", 4, "Aviso")

        If respuesta = vbYes Then
            BotaData(NumBotas) = BotaData(OrigenBota)

            If BorrarOriginal Then BotaData(OrigenBota) = BotaVacio

        End If

    Else
        BotaData(NumBotas) = BotaData(OrigenBota)

        If BorrarOriginal Then BotaData(OrigenBota) = BotaVacio

    End If

    CuentaIndex = -1
    frmMain.Lista.Clear

    For i = 1 To UBound(BotaData)

        If BotaData(i).Bota(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i
            CuentaIndex = CuentaIndex + 1

        End If

        If i = NumBotas Then
            EsteIndex = CuentaIndex

        End If

    Next i

    frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub AgregaBota(ByVal NumBotas As Integer, _
                      Optional ByVal RefreshList As Boolean = True)

    Dim i           As Long

    Dim EsteIndex   As Long

    Dim CuentaIndex As Long

    If NumBotas > UBound(BotaData) Then ReDim Preserve BotaData(0 To NumBotas) As BotaData

    BotaData(NumBotas).Bota(1).GrhIndex = 1

    If Not RefreshList Then Exit Sub

    CuentaIndex = -1
    frmMain.Lista.Clear

    For i = 1 To UBound(BotaData)

        If BotaData(i).Bota(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i
            CuentaIndex = CuentaIndex + 1

        End If

        If i = NumBotas Then
            EsteIndex = CuentaIndex

        End If

    Next i

    frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub AgregaCapa(ByVal NumCapa As Integer, _
                      Optional ByVal RefreshList As Boolean = True)

    Dim i           As Long

    Dim EsteIndex   As Long

    Dim CuentaIndex As Long

    If NumCapa > UBound(EspaldaAnimData) Then ReDim Preserve EspaldaAnimData(0 To NumCapa) As HeadData

    EspaldaAnimData(NumCapa).Head(1).GrhIndex = 1

    If Not RefreshList Then Exit Sub
    CuentaIndex = -1
    frmMain.Lista.Clear

    For i = 1 To UBound(EspaldaAnimData)

        If EspaldaAnimData(i).Head(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i
            CuentaIndex = CuentaIndex + 1

        End If

        If i = NumCapa Then
            EsteIndex = CuentaIndex

        End If

    Next i

    frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub AgregaFx(ByVal FxCapa As Integer)

    Dim i           As Long

    Dim EsteIndex   As Long

    Dim CuentaIndex As Long

    If FxCapa > UBound(FxData) Then ReDim Preserve FxData(0 To FxCapa) As FxData

    FxData(FxCapa).Fx.GrhIndex = 1

    CuentaIndex = -1
    frmMain.Lista.Clear

    For i = 1 To UBound(FxData)

        If FxData(i).Fx.GrhIndex > 0 Then
            frmMain.Lista.AddItem i
            CuentaIndex = CuentaIndex + 1

        End If

        If i = FxCapa Then
            EsteIndex = CuentaIndex

        End If

    Next i

    frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub MueveCapa(ByVal NumCapa As Integer, _
                     ByVal origenCapa As Integer, _
                     Optional ByVal BorrarOriginal As Boolean = True)

    Dim i           As Long

    Dim EsteIndex   As Long

    Dim CuentaIndex As Long

    Dim respuesta   As Byte

    Dim CapaVacia   As HeadData

    If NumCapa > UBound(EspaldaAnimData) Then ReDim Preserve EspaldaAnimData(0 To NumCapa) As HeadData

    If EspaldaAnimData(NumCapa).Head(1).GrhIndex > 0 Then
        respuesta = MsgBox("La capa " & NumCapa & " ya existe, �deseas sobreescribirla?", 4, "Aviso")

        If respuesta = vbYes Then
            EspaldaAnimData(NumCapa) = EspaldaAnimData(origenCapa)

            If BorrarOriginal Then EspaldaAnimData(origenCapa) = CapaVacia

        End If

    Else
        EspaldaAnimData(NumCapa) = EspaldaAnimData(origenCapa)

        If BorrarOriginal Then EspaldaAnimData(origenCapa) = CapaVacia

    End If

    CuentaIndex = -1
    frmMain.Lista.Clear

    For i = 1 To UBound(EspaldaAnimData)

        If EspaldaAnimData(i).Head(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i
            CuentaIndex = CuentaIndex + 1

        End If

        If i = NumCapa Then
            EsteIndex = CuentaIndex

        End If

    Next i

    frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub MueveFX(ByVal NumFx As Integer, _
                   ByVal origenFx As Integer, _
                   Optional ByVal BorrarOriginal As Boolean = True)

    Dim i           As Long

    Dim EsteIndex   As Long

    Dim CuentaIndex As Long

    Dim respuesta   As Byte

    Dim fxVacio     As FxData

    If NumFx > UBound(FxData) Then ReDim Preserve FxData(0 To NumFx) As FxData

    If FxData(NumFx).Fx.GrhIndex > 0 Then
        respuesta = MsgBox("El fx " & NumFx & " ya existe, �deseas sobreescribirlo?", 4, "Aviso")

        If respuesta = vbYes Then
            FxData(NumFx) = FxData(origenFx)

            If BorrarOriginal Then FxData(origenFx) = fxVacio

        End If

    Else
        FxData(NumFx) = FxData(origenFx)

        If BorrarOriginal Then FxData(origenFx) = fxVacio

    End If

    CuentaIndex = -1
    frmMain.Lista.Clear

    For i = 1 To UBound(FxData)

        If FxData(i).Fx.GrhIndex > 0 Then
            frmMain.Lista.AddItem i
            CuentaIndex = CuentaIndex + 1

        End If

        If i = NumFx Then
            EsteIndex = CuentaIndex

        End If

    Next i

    frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub AgregaAlas(ByVal NumAlas As Integer, _
                      Optional ByVal RefreshList As Boolean = True)

    Dim i           As Long

    Dim EsteIndex   As Long

    Dim CuentaIndex As Long

    If NumAlas > UBound(AlaData) Then ReDim Preserve AlaData(0 To NumAlas) As AlaData

    AlaData(NumAlas).Alas(1).GrhIndex = 1

    If Not RefreshList Then Exit Sub

    CuentaIndex = -1
    frmMain.Lista.Clear

    For i = 1 To UBound(AlaData)

        If AlaData(i).Alas(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i
            CuentaIndex = CuentaIndex + 1

        End If

        If i = NumAlas Then
            EsteIndex = CuentaIndex

        End If

    Next i

    frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub MueveAlas(ByVal NumAlas As Integer, _
                     ByVal origenAlas As Integer, _
                     Optional ByVal BorrarOriginal As Boolean = True)

    Dim i           As Long

    Dim EsteIndex   As Long

    Dim CuentaIndex As Long

    Dim AlaVacio    As AlaData

    Dim respuesta   As Byte

    If NumAlas > UBound(AlaData) Then ReDim Preserve AlaData(0 To NumAlas) As AlaData
    If AlaData(NumAlas).Alas(1).GrhIndex > 0 Then
        respuesta = MsgBox("El ala " & NumAlas & " ya existe, �deseas sobreescribirlo?", 4, "Aviso")

        If respuesta = vbYes Then
            AlaData(NumAlas) = AlaData(origenAlas)

            If BorrarOriginal Then AlaData(origenAlas) = AlaVacio

        End If

    Else
        AlaData(NumAlas) = AlaData(origenAlas)

        If BorrarOriginal Then AlaData(origenAlas) = AlaVacio

    End If

    CuentaIndex = -1
    frmMain.Lista.Clear

    For i = 1 To UBound(AlaData)

        If AlaData(i).Alas(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i
            CuentaIndex = CuentaIndex + 1

        End If

        If i = NumAlas Then
            EsteIndex = CuentaIndex

        End If

    Next i

    frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub RenuevaListaGrH()

    Dim i As Long

    frmMain.Lista.Clear

    For i = 1 To MAXGrH

        If Grhdata(i).NumFrames = 1 Then
            frmMain.Lista.AddItem i
        ElseIf Grhdata(i).NumFrames > 1 Then
            frmMain.Lista.AddItem i & " (animacion)"

        End If

    Next i

End Sub

Public Sub RenuevaListaAlas()

    Dim i As Long

    frmMain.Lista.Clear

    For i = 1 To UBound(AlaData)

        If AlaData(i).Alas(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i

        End If

    Next i

End Sub

Public Sub RenuevaListaBodys()

    Dim i As Long

    frmMain.Lista.Clear

    For i = 1 To UBound(BodyData)

        If BodyData(i).Walk(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i

        End If

    Next i

End Sub

Public Sub RenuevaListaCabezas()

    Dim i As Long

    frmMain.Lista.Clear

    For i = 1 To UBound(HeadData)

        If HeadData(i).Head(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i

        End If

    Next i

End Sub

Public Sub RenuevaListaCascos()

    Dim i As Long

    frmMain.Lista.Clear

    For i = 1 To UBound(CascoAnimData)

        If CascoAnimData(i).Head(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i

        End If

    Next i

End Sub

Public Sub RenuevaListaEscudos()

    Dim i As Long

    frmMain.Lista.Clear

    For i = 1 To UBound(ShieldAnimData)

        If ShieldAnimData(i).ShieldWalk(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i

        End If

    Next i

End Sub

Public Sub RenuevaListaArmas()

    Dim i As Long

    frmMain.Lista.Clear

    For i = 1 To UBound(WeaponAnimData)

        If WeaponAnimData(i).WeaponWalk(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i

        End If

    Next i

End Sub

Public Sub RenuevaListaBotas()

    Dim i As Long

    frmMain.Lista.Clear

    For i = 1 To UBound(BotaData)

        If BotaData(i).Bota(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i

        End If

    Next i

End Sub

Public Sub RenuevaListaCapas()

    Dim i As Long

    frmMain.Lista.Clear

    For i = 1 To UBound(EspaldaAnimData)

        If EspaldaAnimData(i).Head(1).GrhIndex > 0 Then
            frmMain.Lista.AddItem i

        End If

    Next i

End Sub

Public Sub RenuevaListaFX()

    Dim i As Long

    frmMain.Lista.Clear

    For i = 1 To UBound(FxData)

        If FxData(i).Fx.GrhIndex > 0 Then
            frmMain.Lista.AddItem i

        End If

    Next i

End Sub

Public Sub RenuevaListaResource()

    Dim i As Long

    frmMain.Lista.Clear

    For i = 1 To MAXGrH

        If ExisteBMP(i) > 0 Then
            frmMain.Lista.AddItem i

        End If

    Next i

End Sub

Public Function GrhCorrecto(ByRef GrhT As Grhdata, _
                            ByRef ErrorMSG As String, _
                            ByRef ErroresGrh As ErroresGrh) As Long

    ' Comprueba que un grafico es correcto
    Dim Alto           As Long

    Dim Ancho          As Long

    Dim i              As Long

    Dim DumyString     As String

    Dim PrimerAlto     As Long

    Dim PrimerAncho    As Long

    Dim dumyErroresGrh As ErroresGrh

    ErroresGrh.ErrorCritico = False

    If GrhT.NumFrames <= 0 Then
        ErrorMSG = "N� de frames incorrecto"
        GrhCorrecto = 0
        ErroresGrh.ErrorCritico = True
        ErroresGrh.colores(2) = vbRed
        Exit Function

    End If

    If GrhT.NumFrames = 1 Then
        'si es solo un frame lo comprobamos
        GrhCorrecto = GrhCorrectoNormal(GrhT, ErrorMSG, ErroresGrh)
        ErroresGrh.EsAnimacion = False
    Else
        ErroresGrh.EsAnimacion = True

        ' si es una animacion, comprobamos frame a frame
        For i = 1 To GrhT.NumFrames

            If GrhT.Frames(i) > 0 Then
                If Grhdata(GrhT.Frames(i)).NumFrames <> 1 Or (GrhCorrectoNormal(Grhdata(GrhT.Frames(i)), DumyString, dumyErroresGrh) < 2) Then
                    ErrorMSG = ErrorMSG & "El frame n� " & i & " es incorrecto. "
                    ErroresGrh.ErrorCritico = True
                    GrhCorrecto = 1
                    ErroresGrh.colores(1) = vbRed
                Else

                    If i = 1 Then
                        PrimerAlto = Grhdata(GrhT.Frames(i)).pixelHeight
                        PrimerAncho = Grhdata(GrhT.Frames(i)).pixelWidth
                    Else
                        Alto = Grhdata(GrhT.Frames(i)).pixelHeight
                        Ancho = Grhdata(GrhT.Frames(i)).pixelWidth

                        If Alto <> PrimerAlto Then
                            ErrorMSG = ErrorMSG & "El frame n� " & i & " distintas dimensiones. "
                            ErroresGrh.colores(1) = vbYellow
                        ElseIf Ancho <> PrimerAncho Then
                            ErrorMSG = ErrorMSG & "El frame n� " & i & " distintas dimensiones. "
                            ErroresGrh.colores(1) = vbYellow

                        End If

                    End If

                End If

            Else
                ErrorMSG = ErrorMSG & "Falta frame n� " & i & ". "
                ErroresGrh.ErrorCritico = True
                ErroresGrh.colores(1) = vbRed

            End If

        Next i

    End If

End Function

Public Function GrhCorrectoNormal(ByRef GrhT As Grhdata, _
                                  ByRef ErrorMSG As String, _
                                  ByRef ErroresGrh As ErroresGrh) As Long

    Dim Alto   As Long

    Dim Ancho  As Long

    Dim dumYin As Integer

    'Comprueba que el grh es correcto. Ademas pone en rojo los texboxes con datos incorrectos.

    If GrhT.NumFrames <= 0 Then
        ErrorMSG = "N� de frames incorrecto"
        GrhCorrectoNormal = 0
        ErroresGrh.colores(2) = vbRed
        ErroresGrh.ErrorCritico = True
        Exit Function

    End If
    
    If ExisteBMP(GrhT.FileNum) = ResourceFile Or (ResourceFile = 3 And ExisteBMP(GrhT.FileNum) > 0) Then
        Call GetTama�oBMP(GrhT.FileNum, Alto, Ancho, dumYin)
    Else
        ErrorMSG = "El archivo " & GrhT.FileNum & ".bmp no existe"
        GrhCorrectoNormal = 1
        ErroresGrh.colores(0) = vbRed
        ErroresGrh.ErrorCritico = True
        Exit Function

    End If
    
    GrhCorrectoNormal = 2 'mascara d bits, bit de grafico existente
    
    If GrhT.sX > Ancho Or GrhT.sY > Alto Then
        If GrhT.sX > Ancho Then
            ErrorMSG = ErrorMSG & "Posicion X fuera del BMP. "
            GrhCorrectoNormal = GrhCorrectoNormal + 8 'mascara d bits , bit de error 2
            ErroresGrh.colores(6) = vbRed

        End If

        If GrhT.sY > Alto Then
            ErrorMSG = ErrorMSG & "Posicion Y fuera del BMP. "
            GrhCorrectoNormal = GrhCorrectoNormal + 4 'mascara d bits , bit de error 1
            ErroresGrh.colores(7) = vbRed

        End If

    Else

        If GrhT.sY + GrhT.pixelHeight > Alto Then
            ErrorMSG = ErrorMSG & "Alto fuera del BMP. "
            GrhCorrectoNormal = GrhCorrectoNormal + 16 'mascara d bits , bit de error 3
            ErroresGrh.colores(3) = vbYellow

        End If

        If GrhT.sX + GrhT.pixelWidth > Ancho Then
            ErrorMSG = ErrorMSG & "Ancho fuera del BMP. "
            GrhCorrectoNormal = GrhCorrectoNormal + 32 'mascara d bits , bit de error 4
            ErroresGrh.colores(4) = vbYellow

        End If

    End If

End Function
