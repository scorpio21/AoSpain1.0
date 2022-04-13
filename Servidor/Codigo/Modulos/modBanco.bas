Attribute VB_Name = "modBanco"
Option Explicit

'MODULO PROGRAMADO POR NEB
'Kevin Birmingham
'kbneb@hotmail.com

Sub IniciarDeposito(ByVal UserIndex As Integer)
On Error GoTo errhandler

'Hacemos un Update del inventario del usuario
Call UpdateBanUserInv(True, UserIndex, 0)
'Atcualizamos el dinero
Call SendUserStatsBox(UserIndex)
'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
SendData ToIndex, UserIndex, 0, "INITBANCO"
UserList(UserIndex).Flags.Comerciando = True

errhandler:

End Sub

Sub SendBanObj(UserIndex As Integer, SLOT As Byte, Object As UserOBJ)


UserList(UserIndex).BancoInvent.Object(SLOT) = Object

If Object.ObjIndex > 0 Then

    Call SendData(ToIndex, UserIndex, 0, "SBO" & SLOT & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
    & ObjData(Object.ObjIndex).ObjType & "," _
    & ObjData(Object.ObjIndex).MaxHIT & "," _
    & ObjData(Object.ObjIndex).MinHIT & "," _
    & ObjData(Object.ObjIndex).MaxDef)

Else

    Call SendData(ToIndex, UserIndex, 0, "SBO" & SLOT & "," & "0" & "," & "(None)" & "," & "0" & "," & "0")

End If


End Sub

Sub UpdateBanUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal SLOT As Byte)

Dim NullObj As UserOBJ
Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).BancoInvent.Object(SLOT).ObjIndex > 0 Then
        Call SendBanObj(UserIndex, SLOT, UserList(UserIndex).BancoInvent.Object(SLOT))
    Else
        Call SendBanObj(UserIndex, SLOT, NullObj)
    End If

Else

'Actualiza todos los slots
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS

        'Actualiza el inventario
        If UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex > 0 Then
            Call SendBanObj(UserIndex, LoopC, UserList(UserIndex).BancoInvent.Object(LoopC))
        Else
            
            Call SendBanObj(UserIndex, LoopC, NullObj)
            
        End If

    Next LoopC

End If

End Sub

Sub UserRetiraItem(ByVal UserIndex As Integer, ByVal i As Integer, ByVal Cantidad As Integer)
On Error GoTo errhandler


If Cantidad < 1 Then Exit Sub


Call SendUserStatsBox(UserIndex)

   
       If UserList(UserIndex).BancoInvent.Object(i).Amount > 0 Then
         
            If Cantidad > UserList(UserIndex).BancoInvent.Object(i).Amount Then Cantidad = UserList(UserIndex).BancoInvent.Object(i).Amount
            'Agregamos el obj que compro al inventario
            Call UserReciveObj(UserIndex, CInt(i), Cantidad)
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, UserIndex, 0)
            'Actualizamos el banco
            Call UpdateBanUserInv(True, UserIndex, 0)
            'Actualizamos la ventana de comercio
            Call UpdateVentanaBanco(i, 0, UserIndex)
       End If



errhandler:

End Sub

Sub UserReciveObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)

Dim SLOT As Integer
Dim obji As Integer


If UserList(UserIndex).BancoInvent.Object(ObjIndex).Amount <= 0 Then Exit Sub

obji = UserList(UserIndex).BancoInvent.Object(ObjIndex).ObjIndex


'�Ya tiene un objeto de este tipo?
SLOT = 1
Do Until UserList(UserIndex).Invent.Object(SLOT).ObjIndex = obji And _
   UserList(UserIndex).Invent.Object(SLOT).Amount + Cantidad <= MAX_INVENTORY_OBJS
    
    SLOT = SLOT + 1
    If SLOT > MAX_INVENTORY_SLOTS Then
        Exit Do
    End If
Loop

'Sino se fija por un slot vacio
If SLOT > MAX_INVENTORY_SLOTS Then
        SLOT = 1
        Do Until UserList(UserIndex).Invent.Object(SLOT).ObjIndex = 0
            SLOT = SLOT + 1

            If SLOT > MAX_INVENTORY_SLOTS Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes tener mas objetos." & FONTTYPE_INFO)
                Exit Sub
            End If
        Loop
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
End If



'Mete el obj en el slot
If UserList(UserIndex).Invent.Object(SLOT).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
    
    'Menor que MAX_INV_OBJS
    UserList(UserIndex).Invent.Object(SLOT).ObjIndex = obji
    UserList(UserIndex).Invent.Object(SLOT).Amount = UserList(UserIndex).Invent.Object(SLOT).Amount + Cantidad
    
    Call QuitarBancoInvItem(UserIndex, CByte(ObjIndex), Cantidad)
Else
    Call SendData(ToIndex, UserIndex, 0, "||No puedes tener mas objetos." & FONTTYPE_INFO)
End If


End Sub

Sub QuitarBancoInvItem(ByVal UserIndex As Integer, ByVal SLOT As Byte, ByVal Cantidad As Integer)



Dim ObjIndex As Integer
ObjIndex = UserList(UserIndex).BancoInvent.Object(SLOT).ObjIndex

    'Quita un Obj

       UserList(UserIndex).BancoInvent.Object(SLOT).Amount = UserList(UserIndex).BancoInvent.Object(SLOT).Amount - Cantidad
        
        If UserList(UserIndex).BancoInvent.Object(SLOT).Amount <= 0 Then
            UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems - 1
            UserList(UserIndex).BancoInvent.Object(SLOT).ObjIndex = 0
            UserList(UserIndex).BancoInvent.Object(SLOT).Amount = 0
        End If

    
    
End Sub

Sub UpdateVentanaBanco(ByVal SLOT As Integer, ByVal NpcInv As Byte, ByVal UserIndex As Integer)
 
 
 Call SendData(ToIndex, UserIndex, 0, "BANCOOK" & SLOT & "," & NpcInv)
 
End Sub

Sub UserDepositaItem(ByVal UserIndex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer)

On Error GoTo errhandler

'El usuario deposita un item
Call SendUserStatsBox(UserIndex)
   
If UserList(UserIndex).Invent.Object(Item).Amount > 0 And UserList(UserIndex).Invent.Object(Item).Equipped = 0 Then
'ULISES NO SE PUEDEN DEPOSITAR EN BOVEDA 922 CABALLO 923 UNICORNIO
If ((UserList(UserIndex).Invent.Object(Item).ObjIndex = 923) Or (UserList(UserIndex).Invent.Object(Item).ObjIndex = 922)) Then
             Call SendData(ToIndex, UserIndex, 0, "||No Compro Caballos!." & FONTTYPE_WARNING)
             Call UpdateVentanaBanco(Item, 1, UserIndex)
             Exit Sub
            End If
            
            If Cantidad > 0 And Cantidad > UserList(UserIndex).Invent.Object(Item).Amount Then Cantidad = UserList(UserIndex).Invent.Object(Item).Amount
            'Agregamos el obj que compro al inventario
            Call UserDejaObj(UserIndex, CInt(Item), Cantidad)
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, UserIndex, 0)
            'Actualizamos el inventario del banco
            Call UpdateBanUserInv(True, UserIndex, 0)
            'Actualizamos la ventana del banco
            
            Call UpdateVentanaBanco(Item, 1, UserIndex)
            
End If

errhandler:

End Sub

Sub UserDejaObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)

Dim SLOT As Integer
Dim obji As Integer

If Cantidad < 1 Then Exit Sub

obji = UserList(UserIndex).Invent.Object(ObjIndex).ObjIndex

'�Ya tiene un objeto de este tipo?
SLOT = 1
Do Until UserList(UserIndex).BancoInvent.Object(SLOT).ObjIndex = obji And _
         UserList(UserIndex).BancoInvent.Object(SLOT).Amount + Cantidad <= MAX_INVENTORY_OBJS
            SLOT = SLOT + 1
        
            If SLOT > MAX_BANCOINVENTORY_SLOTS Then
                Exit Do
            End If
Loop

'Sino se fija por un slot vacio antes del slot devuelto
If SLOT > MAX_BANCOINVENTORY_SLOTS Then
        SLOT = 1
        Do Until UserList(UserIndex).BancoInvent.Object(SLOT).ObjIndex = 0
            SLOT = SLOT + 1

            If SLOT > MAX_BANCOINVENTORY_SLOTS Then
                Call SendData(ToIndex, UserIndex, 0, "||No tienes mas espacio en el banco!!" & FONTTYPE_INFO)
                Exit Sub
                Exit Do
            End If
        Loop
        If SLOT <= MAX_BANCOINVENTORY_SLOTS Then UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems + 1
        
        
End If

If SLOT <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido
    'Mete el obj en el slot
    If UserList(UserIndex).BancoInvent.Object(SLOT).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
        
        'Menor que MAX_INV_OBJS
        UserList(UserIndex).BancoInvent.Object(SLOT).ObjIndex = obji
        UserList(UserIndex).BancoInvent.Object(SLOT).Amount = UserList(UserIndex).BancoInvent.Object(SLOT).Amount + Cantidad
        
        Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)

    Else
        Call SendData(ToIndex, UserIndex, 0, "||El banco no puede cargar tantos objetos." & FONTTYPE_INFO)
    End If

Else
    Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)
End If

End Sub


