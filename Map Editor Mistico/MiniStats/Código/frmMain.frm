VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{60CC5D62-2D08-11D0-BDBE-00AA00575603}#1.1#0"; "SYSTRAY.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lite Stats Server"
   ClientHeight    =   3345
   ClientLeft      =   2085
   ClientTop       =   2430
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4755
   Begin VB.Frame frmCorreo 
      Caption         =   "Servidor de correo saliente"
      Height          =   1455
      Left            =   0
      TabIndex        =   8
      Top             =   1800
      Width           =   4695
      Begin VB.TextBox txtSTMP 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Text            =   "smtp.miserver.com"
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtMailFrom 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Text            =   "Staff de Mi server"
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Text            =   "staff@miserver.com"
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Direccion SMTP:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Su nombre:"
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Su e-mail:"
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   600
      TabIndex        =   6
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Frame frmDatos 
      Caption         =   "Server Data"
      Height          =   1335
      Left            =   840
      TabIndex        =   2
      Top             =   0
      Width           =   3855
      Begin SysTrayCtl.cSysTray SysTray 
         Left            =   3120
         Top             =   240
         _ExtentX        =   900
         _ExtentY        =   900
         InTray          =   0   'False
         TrayIcon        =   "frmMain.frx":0000
         TrayTip         =   "Lite Stats Server"
      End
      Begin VB.Label lblLastError 
         Caption         =   "Último error/aviso:"
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
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label lblEmail 
         Caption         =   "E-mail enviados:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label lblRequests 
         Caption         =   "Peticiones:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmdSystray 
      Caption         =   "Sys Tray"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdOnOff 
      Caption         =   "Iniciar el servidor"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
   Begin MSWinsockLib.Winsock Wsck 
      Index           =   0
      Left            =   2760
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin MSWinsockLib.Winsock Wsck 
      Index           =   1
      Left            =   2760
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin MSWinsockLib.Winsock Wsck 
      Index           =   2
      Left            =   2760
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin VB.Label lblIP 
      Caption         =   "Su IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Menu mnu_SysTray 
      Caption         =   "SysTray"
      Visible         =   0   'False
      Begin VB.Menu mnu_mostrar 
         Caption         =   "Mostrar"
      End
      Begin VB.Menu mnu_salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Private Declare Function gettickcount Lib "kernel32" Alias "GetTickCount" () As Long
Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal T As Long, ByVal R As String)

Private WithEvents poSendMail As vbSendMail.clsSendMail
Attribute poSendMail.VB_VarHelpID = -1

Private Type SockCtl
    RecivedData As Boolean
    Data As String
    Datos As String
    Timeout As Long
End Type

Dim IDRequest As Integer
Dim PAGINA As String
Dim SockCtl(2) As SockCtl
Dim ServerIniciado As Boolean
Dim Request As Long
Dim email As Long
Dim LastError As String
Dim LastClean As Date


Private Sub cmdOnOff_Click()
Dim Error As Integer
On Error GoTo err:

If ServerIniciado = False Then
Error = 1
txtIP.Locked = True
Error = 2
Wsck(0).LocalPort = 80
Wsck(0).Listen
Error = 3
cmdOnOff.Caption = "Detener el servidor"
ServerIniciado = True
LogServer ("Servidor iniciado")
Else

txtIP.Locked = False
Error = 4
Dim LoopA As Integer
For LoopA = 0 To 2
Wsck(LoopA).Close
SockCtl(LoopA).Data = ""
SockCtl(LoopA).Datos = ""
SockCtl(LoopA).RecivedData = False
SockCtl(LoopA).Timeout = 0
Next
Error = 3
cmdOnOff.Caption = "Iniciar el servidor"
ServerIniciado = False
LogServer ("Detenido")
End If

ActualizaDatos
Exit Sub

err:
Select Case Error:
    Case 1: LastError = "Error en la dirección IP."
    Case 2: LastError = "Error al poner socket a la escucha."
    Case 3: LastError = "Error desconocido."
    Case 4: LastError = "Error al detener el socket."
End Select
ActualizaDatos
End Sub

Private Sub cmdSystray_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Dim LoopA As Integer
SysTray.InTray = True
ServerIniciado = False
LastClean = Date
Request = 0
email = 0
txtIP = Wsck(0).LocalIP
For LoopA = 0 To 2
Wsck(LoopA).Close
Next
LogServer ("Programa iniciado")
End Sub

Private Sub ActualizaDatos()
lblRequests = "Peticiones: " & Request
lblEmail = "E-mail enviados:" & email
lblLastError = "Último error/aviso:" & vbCrLf & LastError
End Sub

 

Private Sub Wsck_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    On Error GoTo fin:
    'Si ya paso el dia... borramos todo...
    Request = Request + 1
    If LastClean <> Date Then
        Kill (App.Path & "\Actions\*.dat")
        LastClean = Date
        Request = 0
        email = 0
        IDRequest = 0
        LogServer ("Peticiones eliminadas")
    End If
    
    If IDRequest = 10000 Then IDRequest = 0
    
    SockCtl(Index).RecivedData = False
    If Wsck(Index).State <> sckClosed Then Wsck(Index).Close
    If LastSockListening(Index) = -1 And LastSockClosed(Index) <> -1 Then
        Wsck(LastSockClosed(Index)).LocalPort = 80
        Wsck(LastSockClosed(Index)).Listen
    End If
   
    Wsck(Index).LocalPort = 0
    Wsck(Index).Accept requestID
    SockCtl(Index).Timeout = gettickcount
    While Not SockCtl(Index).RecivedData
        DoEvents
        If SockCtl(Index).Timeout + 1000 < gettickcount Then GoTo fin
   Wend
   If InStr(SockCtl(Index).Data, "/") <> 0 Or InStr(SockCtl(Index).Data, ".") <> 0 _
   Or InStr(SockCtl(Index).Data, "\") <> 0 Then
   Wsck(Index).SendData "<html><font face='Verdana' size='1'><b>Los singos / y . no estan admitidos.</b></font></html>"
   Exit Sub
   End If
      
  If SockCtl(Index).Data = "" Then SockCtl(Index).Data = "inicio"
   
   If Left$(SockCtl(Index).Data, 9) = "DoConfirm" Then
        Dim auxstr() As String
        auxstr = Split(Right$(SockCtl(Index).Data, Len(SockCtl(Index).Data) - 9), "%C2%AC")
        Call Confirm(Index, auxstr(0), auxstr(1))
        Exit Sub
   End If
    
    If SockCtl(Index).Data = "RecPas" Then
    Dim Personaje As String
    Dim E_mail As String
    'Destripamos para obtener el nombre y email
    Personaje = Mid$(SockCtl(Index).Datos, 4, InStr(SockCtl(Index).Datos, "&") - 4)
    E_mail = Mid$(SockCtl(Index).Datos, Len(Personaje) + 11, Len(SockCtl(Index).Datos) - Len(Personaje) - 12)
    If InStr(Personaje, "/") <> 0 Or InStr(Personaje, ".") <> 0 _
    Or InStr(Personaje, "\") <> 0 Or Not (ExisteArchivo(App.Path & "\Charfile\" & Personaje & ".chr")) Then
        Wsck(Index).SendData "<html><font face='Verdana' size='1'><b>¡¡¡ El personaje no existe !!!</b></font></html>"
        Exit Sub
    End If
    If Not (GetVar(App.Path & "\Charfile\" & Personaje & ".chr", "CONTACTO", "Email") = E_mail) Then
         Wsck(Index).SendData "<html><font face='Verdana' size='1'><b>¡¡¡ El e-mail no coincide con el ingresado cuando creó el personaje !!!</b></font></html>"
        Exit Sub
    End If
    IDRequest = IDRequest + 1
    'Creamos el password, activamos la página de confirmación y le enviamos el mail !
    Dim NewPassword As String
    Dim MailToSend As String
    Dim AuxString As String
    NewPassword = Format((999999 - 1 + 1) * Rnd + 1, "000000")
    'Comenzamos a crear el mail
    Call WriteVar(App.Path & "\Actions\" & Format(IDRequest, "0000") & ".dat", "Action", "Action", "2")
    Call WriteVar(App.Path & "\Actions\" & Format(IDRequest, "0000") & ".dat", "Action", "Password", MD5String(NewPassword))
    Call WriteVar(App.Path & "\Actions\" & Format(IDRequest, "0000") & ".dat", "Action", "PJ", Personaje)
    AuxString = Format((999999999 - 1 + 1) * Rnd + 1, "000000000")
    Call WriteVar(App.Path & "\Actions\" & Format(IDRequest, "0000") & ".dat", "Confirmacion", "ConfirmString", AuxString)
    MailToSend = "La nueva clave de tu personaje " & Personaje & _
    " es " & NewPassword & "." & vbCrLf & "Para activarlo debes visitar http://" & _
    txtIP & "/DoConfirm" & Format(IDRequest, "0000") & "¬" & AuxString & vbCrLf & txtMailFrom
    'Ya tenemos el mail, y el archivo creados
    Wsck(Index).SendData "<html><font face='Verdana' size='1'><b>¡¡¡ Has iniciado el tramite de recuperación de password!!!<br>Para hacerlo definitivo deberás revisar tu e-mail y confirmar la modificación de contraseña.</b></font></html>"
    Call SendeMail(E_mail, MailToSend)
    Exit Sub
   End If
    
    If SockCtl(Index).Data = "BorraPJ" Then
    Dim Pj As String
    Dim mail As String
    Dim Password As String
    
    'Destripamos para obtener el nombre, email y password
    Pj = Mid$(SockCtl(Index).Datos, 4, InStr(SockCtl(Index).Datos, "&email=") - 4)
    mail = Mid$(SockCtl(Index).Datos, Len(Pj) + 11, InStr(SockCtl(Index).Datos, "&password=") - InStr(SockCtl(Index).Datos, "&email=") - 7)
    Password = Mid$(SockCtl(Index).Datos, Len(Pj) + Len(mail) + 21, Len(SockCtl(Index).Datos) - Len(Pj) - Len(mail) - 22)
    
    If InStr(Personaje, "/") <> 0 Or InStr(Personaje, ".") <> 0 _
    Or InStr(Personaje, "\") <> 0 Or Not (ExisteArchivo(App.Path & "\Charfile\" & Pj & ".chr")) Then
        Wsck(Index).SendData "<html><font face='Verdana' size='1'><b>¡¡¡ El personaje no existe !!!</b></font></html>"
        Exit Sub
    End If
    
    If Not (GetVar(App.Path & "\Charfile\" & Pj & ".chr", "CONTACTO", "Email") = mail) Then
         Wsck(Index).SendData "<html><font face='Verdana' size='1'><b>¡¡¡ El e-mail no coincide con el ingresado cuando creó el personaje !!!</b></font></html>"
        Exit Sub
    End If
    
    If Not GetVar(App.Path & "\Charfile\" & Pj & ".chr", "INIT", "Password") = MD5String(Password) Then
         Wsck(Index).SendData "<html><font face='Verdana' size='1'><b>¡¡¡ La contraseña es incorrecta !!!</b></font></html>"
        Exit Sub
    End If
    IDRequest = IDRequest + 1
    'Activamos la página de confirmación y le enviamos el mail !
    Call WriteVar(App.Path & "\Actions\" & Format(IDRequest, "0000") & ".dat", "Action", "Action", "1")
    AuxString = Format((999999999 - 1 + 1) * Rnd + 1, "000000000")
    Call WriteVar(App.Path & "\Actions\" & Format(IDRequest, "0000") & ".dat", "Action", "PJ", Pj)
    Call WriteVar(App.Path & "\Actions\" & Format(IDRequest, "0000") & ".dat", "Confirmacion", "ConfirmString", AuxString)
    MailToSend = "Para borrar a tu personaje " & Personaje & _
    vbCrLf & "Para borrarlo definitivamente debes visitar http://" & _
    txtIP & "/DoConfirm" & Format(IDRequest, "0000") & "¬" & AuxString & vbCrLf & txtMailFrom
    'Ya tenemos el mail, y el archivo creados
    Wsck(Index).SendData "<html><font face='Verdana' size='1'><b>¡¡¡ Has iniciado el tramite de borrado de tu personaje!!!<br>Para hacerlo definitivo deberás revisar tu e-mail y confirmar la eliminación.</b></font></html>"
    Call SendeMail(mail, MailToSend)
    Exit Sub
   End If
   
    If Left$(SockCtl(Index).Data, 8) = "SeeStats" Then
            Dim STRAUX As String
            STRAUX = Right$(SockCtl(Index).Data, Len(SockCtl(Index).Data) - 12)
            If Not ExisteArchivo(App.Path & "\Charfile\" & STRAUX & ".chr") Then
                Wsck(Index).SendData "<html><font face='Verdana' size='1'><b>¡¡¡ NO EXISTE EL PERSONAJE !!!</b></font></html>"
                Exit Sub
            End If
            Wsck(Index).SendData GetMiniStats(UCase$(STRAUX))
            Exit Sub
    End If

   If ExisteArchivo(App.Path & "\www\" & SockCtl(Index).Data) Then
      Wsck(Index).SendData AbrirArchivo(App.Path & "/www/" & SockCtl(Index).Data)
   Else
      Wsck(Index).SendData "<html><font face='Verdana' size='1'><b>NO SE ENCUENTRA LA PÁGINA</b></font></html>"
   End If
fin:
    ActualizaDatos
End Sub

Private Sub Wsck_DataArrival(Index As Integer, ByVal bytesTotal As Long)
  On Error Resume Next
  Dim Data As String
  Dim STRAUX As String
  Dim BUSCAR_ENVIAR As String
  Dim BUSCAR_RECIBIR As String
  Dim tempint As Integer
  Wsck(Index).GetData Data
  If Mid(Data, 1, 3) = "GET" Then
     BUSCAR_ENVIAR = InStr(Data, "GET ")
     STRAUX = InStr(BUSCAR_ENVIAR + 5, Data, " ")
     SockCtl(Index).Data = Mid(Data, BUSCAR_ENVIAR + 5, STRAUX - (BUSCAR_ENVIAR + 5))
  ElseIf Mid(Data, 1, 4) = "POST" Then
     BUSCAR_RECIBIR = InStr(Data, "POST ")
     STRAUX = InStr(BUSCAR_RECIBIR + 5, Data, " ")
     SockCtl(Index).Data = Mid(Data, BUSCAR_RECIBIR + 6, STRAUX - (BUSCAR_RECIBIR + 6))
     SockCtl(Index).Datos = Right$(Data, Len(Data) - InStr(Data, "pj=") + 1)
  End If
  SockCtl(Index).RecivedData = True

End Sub

Private Sub Wsck_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   
LastError = Description
ActualizaDatos

Wsck(Index).Close
'Si no hay sockets escuchando... Let´s go!
    If LastSockListening(Index) = -1 Then
        Wsck(Index).LocalPort = 80
        Wsck(Index).Listen
    End If



End Sub

Private Sub Wsck_SendComplete(Index As Integer)

Wsck(Index).Close

If LastSockListening(-1) = -1 Then
Wsck(Index).LocalPort = 80
Wsck(Index).Listen
End If

ActualizaDatos
End Sub

Private Function LastSockClosed(ByVal ExcludeSock As Integer) As Integer

LastSockClosed = -1
Dim LoopA As Integer
For LoopA = 0 To 2
If Wsck(LoopA).State = sckClosed And LoopA <> ExcludeSock Then LastSockClosed = LoopA 'LastSockClosed =
Next

If LastSockClosed = -1 Then LastError = "¡¡¡ No hay sockets libres!!!"
ActualizaDatos

Exit Function

End Function

Private Function LastSockListening(ByVal ExcludeSock As Integer) As Integer

LastSockListening = -1
Dim LoopA As Integer
For LoopA = 0 To 2
If Wsck(LoopA).State = sckListening And LoopA <> ExcludeSock Then LastSockListening = LoopA
Next

End Function

Private Function AbrirArchivo(Archivo As String) As String
On Error GoTo fin:
Dim file As Integer
Dim TextAUX As String
Dim CHAR As String * 1

On Error Resume Next
file = FreeFile
TextAUX = ""

If ExisteArchivo(Archivo) Then
    If Len(Archivo) Then
        Open Archivo For Input As #file
        Do While Not EOF(file)
            CHAR = Input(1, #file)
            TextAUX = "" & TextAUX & CHAR
        Loop
        Close #file
    End If
AbrirArchivo = TextAUX
Else
AbrirArchivo = ""
End If
fin:
End Function

Private Function ExisteArchivo(ByVal Archivo As String) As Integer
    Dim I As Integer
    On Error Resume Next

    I = Len(Dir(Archivo))
    
    If err Or I = 0 Then
       ExisteArchivo = False
    Else
       ExisteArchivo = True
    End If
End Function

Private Sub Confirm(Index As Integer, ActionID As String, ConfirmString As String)
   On Error GoTo Error
If GetVar(App.Path & "\Actions\" & ActionID & ".dat", "Confirmacion", "ConfirmString") = ConfirmString Then
    If CInt(GetVar(App.Path & "\Actions\" & ActionID & ".dat", "Action", "Action")) = 1 Then
    Dim PjName As String
    PjName = GetVar(App.Path & "\Actions\" & ActionID & ".dat", "Action", "PJ")
    'Borramos el personaje...
    Name App.Path & "\Charfile\" & PjName & ".chr" As App.Path & "\DeletedChars\" & PjName & ".chr"
    Wsck(Index).SendData "<html><font face='Verdana' size='1'><b>EL PERSONAJE " & PjName & " FUE BORRADO</b></font></html>"
    Kill App.Path & "\Actions\" & ActionID & ".dat"
    LogServer (PjName & " borrado desde el IP " & Wsck(Index).RemoteHostIP)
    End If
    
    If CInt(GetVar(App.Path & "\Actions\" & ActionID & ".dat", "Action", "Action")) = 2 Then
    PjName = GetVar(App.Path & "\Actions\" & ActionID & ".dat", "Action", "PJ")
    'Activamos la nueva clave
    Call WriteVar(App.Path & "\Charfile\" & PjName & ".chr", "INIT", "Password", GetVar(App.Path & "\Actions\" & ActionID & ".dat", "Action", "Password"))
    Wsck(Index).SendData "<html><font face='Verdana' size='1'><b>LA NUEVA CONTRASEÑA PARA " & PjName & " HA SIDO ACTIVADA</b></font></html>"
    Kill App.Path & "\Actions\" & ActionID & ".dat"
    LogServer (PjName & " cambia password desde el IP " & Wsck(Index).RemoteHostIP)
    End If
    Else
    Wsck(Index).SendData "<html><font face='Verdana' size='1'><b>OCURRIÓ UN ERROR. LA SOLICITUD NO EXISTE, ES INVALIDA O HA EXPIRADO!!!</b></font></html>"
    End If
Exit Sub
Error:
    Wsck(Index).SendData "<html><font face='Verdana' size='1'><b>OCURRIÓ UN ERROR. LA SOLICITUD NO EXISTE, ES INVALIDA O HA EXPIRADO!!!</b></font></html>"
    Wsck(Index).Close
    If LastSockListening(-1) = -1 Then
    Wsck(Index).LocalPort = 80
    Wsck(Index).Listen
    End If
End Sub
Function GetMiniStats(Personaje As String) As String

Dim Temp As String

Temp = "<html><title>Mini Estadisticas</title><font face='Verdana' size='3'><b>Estadísticas para " & Personaje & "</b></font><Br><font face='Times New Roman' size='2'>Facción: "
If GetVar(App.Path & "\Charfile\" & Personaje & ".chr", "FACCIONES", "EjercitoReal") = "1" Then
    Temp = Temp & "Armada real<br>"
ElseIf GetVar(App.Path & "\Charfile\" & Personaje & ".chr", "FACCIONES", "EjercitoCaos") = "1" Then
    Temp = Temp & "Armada del caos<br>"
Else
    Temp = Temp & "No pertenece a ninguna<br>"
End If

Temp = Temp & "Clan: "

If GetVar(App.Path & "\Charfile\" & Personaje & ".chr", "GUILD", "GuildName") <> "" Then
    Temp = Temp & GetVar(App.Path & "\Charfile\" & Personaje & ".chr", "GUILD", "GuildName") & "<br>"
Else
    Temp = Temp & "No pertenece a ningún clan.<br>"
End If

Temp = Temp & "Hogar: " & _
GetVar(App.Path & "\Charfile\" & Personaje & ".chr", "INIT", "Hogar") & "<br>"
    
Temp = Temp & "Ciudadanos Matados: " & _
GetVar(App.Path & "\Charfile\" & Personaje & ".chr", "FACCIONES", "CiudMatados") & "<br>"

Temp = Temp & "Criminales Matados: " & _
GetVar(App.Path & "\Charfile\" & Personaje & ".chr", "FACCIONES", "CrimMatados") & "<br>"

Temp = Temp & "NPCs Matados: " & _
GetVar(App.Path & "\Charfile\" & Personaje & ".chr", "MUERTES", "NpcsMuertes") & "<br>"

Temp = Temp & "Baneado: "
If GetVar(App.Path & "\Charfile\" & Personaje & ".chr", "FLAGS", "Ban") = "1" Then
    Temp = Temp & "Si. <br> Baneado por:" & GetVar(App.Path & "\Logs\" & "BanDetail.dat", Personaje, "BannedBy") & _
    "<br> Causa:" & GetVar(App.Path & "\logs\" & "BanDetail.dat", Personaje, "Reason")
Else
    Temp = Temp & "No."
End If
    Temp = Temp & "</font></html>"
    GetMiniStats = Temp
End Function
Sub SendeMail(email As String, Data As String)
Set poSendMail = New vbSendMail.clsSendMail
poSendMail.SMTPHost = txtSTMP
poSendMail.From = txtEmail
poSendMail.FromDisplayName = txtMailFrom
poSendMail.Recipient = email
poSendMail.RecipientDisplayName = "Usuario"
poSendMail.ReplyToAddress = ""
poSendMail.Subject = "Control de personajes"
poSendMail.Message = Data
poSendMail.Send
Set poSendMail = Nothing

End Sub
Private Sub SysTray_MouseDown(Button As Integer, Id As Long)
    If Button = 2 Then PopupMenu mnu_SysTray
End Sub
Private Sub mnu_mostrar_Click()
    Me.Show
End Sub

Private Sub mnu_Salir_Click()
    End
End Sub

'Sacadas del server de AO ;)
Private Function GetVar(file As String, Main As String, Var As String) As String

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
  
szReturn = ""
  
sSpaces = Space(5000) ' This tells the computer how long the longest string can be
  
  
getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), file
  
GetVar = RTrim(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function
Private Sub WriteVar(file As String, Main As String, Var As String, Value As String)
'*****************************************************************
'Escribe VAR en un archivo
'*****************************************************************

writeprivateprofilestring Main, Var, Value, file
    
End Sub

Public Function MD5String(P As String) As String
' compute MD5 digest on a given string, returning the result
    Dim R As String * 32, T As Long
    R = Space(32)
    T = Len(P)
    MDStringFix P, T, R
    MD5String = R
End Function
'Un poco cambiada ;)
Public Sub LogServer(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\MiniStats.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub
