Attribute VB_Name = "API"
Option Explicit

Public IniPath As String
Public Conexion As New Connection
Public LastGrh As Integer
Public FSO As New FileSystemObject
Public Changed As Boolean

'********** OUTSIDE FUNCTIONS ***********'
Declare Function writeprivateprofilestring Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function getprivateprofilestring Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Function BiggestVal() As Integer
On Error Resume Next
Dim i As Integer
Dim max As Integer
Dim cad As String
max = -1900
For i = 0 To frmAdmin.TXTList.ListCount - 1
    cad = ReadField(2, frmAdmin.TXTList.List(i), 45)
    If Val(Right(cad, Len(cad) - 3)) > max Then
          max = Val(Right(cad, Len(cad) - 3))
    End If
Next
BiggestVal = max + 1
    
End Function


Public Sub DelBD(Index As String)
On Error GoTo errorhandler
Dim rs As New Recordset
rs.Open "GrhIndex", Conexion, , adLockOptimistic, adCmdTable
    
Do While Not rs.EOF And rs!GrhIndice <> Index
    
    
    rs.MoveNext
    
Loop

If rs!GrhIndice = Index Then
    rs.Delete
    rs.MoveFirst
End If
rs.Close
Set rs = Nothing

errorhandler:

End Sub

Public Sub AddBD(Nombre As String, Index As String, X As Integer, Y As Integer)
Dim rs As New Recordset
rs.Open "GrhIndex", Conexion, , adLockOptimistic, adCmdTable
    
    rs.AddNew
    rs!Nombre = Nombre
    If Left(Index, 3) = "Grh" Then Index = Right(Index, Len(Index) - 3)
    rs!GrhIndice = Index
    rs!Ancho = X
    rs!alto = Y
    rs.Update
    rs.MoveNext

rs.Close
Set rs = Nothing
End Sub


Public Sub UpdateBD(Index As String, cad As String, X As Integer, Y As Integer)
Dim rs As New Recordset
rs.Open "GrhIndex", Conexion, , adLockOptimistic, adCmdTable

Do While Not rs.EOF And rs!GrhIndice <> Index
    
    
    rs.MoveNext
    
Loop

If rs!GrhIndice = Index Then
    rs!Nombre = ReadField(1, cad, 45)
    rs!GrhIndice = Right(ReadField(2, cad, 45), Len(ReadField(2, cad, 45)) - 3)
    'rs!Ancho = x
    'rs!alto = y
    rs.Update
End If

rs.Close
End Sub

Public Function ReadField(Pos As Integer, Text As String, SepASCII As Integer) As String

'*****************************************************************
'Gets a field from a string
'*****************************************************************

Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String

Seperator = Chr(SepASCII)
LastPos = 0
FieldNum = 0
For i = 1 To Len(Text)
    CurChar = Mid(Text, i, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = Pos Then
            ReadField = Mid(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next i
FieldNum = FieldNum + 1
If FieldNum = Pos Then
    ReadField = Mid(Text, LastPos + 1)
End If


End Function

Function FileExist(File As String, FileType As VbFileAttribute) As Boolean
'*****************************************************************
'Checks to see if a file exists
'*****************************************************************

If Dir(File, FileType) = "" Then
    FileExist = False
Else
    FileExist = True
End If

End Function

Function GetVar(File As String, Main As String, Var As String) As String
'*****************************************************************
'Get a variable from a a text file
'*****************************************************************
Dim l As Integer
Dim Char As String
Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

szReturn = ""

sSpaces = Space(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish


getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File

GetVar = RTrim(sSpaces)
GetVar = Left(GetVar, Len(GetVar) - 1)

End Function

Sub WriteVar(File As String, Main As String, Var As String, Value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************

writeprivateprofilestring Main, Var, Value, File
    
End Sub

Public Sub BuscarDescrip(cad As String)
Dim i As Integer

For i = 0 To frmAdmin.TXTList.ListCount - 1
    If UCase(ReadField(1, frmAdmin.TXTList.List(i), 45)) = UCase(cad) Then
        Changed = False
        frmAdmin.TXTList.ListIndex = i
        Exit For
    End If
Next
End Sub

