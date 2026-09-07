VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "frmLogin"
   ClientHeight    =   3900
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4716
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intentos As Integer
'========================================
' INICIAR FORMULARIO
'========================================
Private Sub UserForm_Initialize()

    On Error Resume Next

    txtUsuario.Value = ""
    txtPassword.Value = ""

    txtUsuario.Visible = True
    txtPassword.Visible = True

    txtUsuario.Enabled = True
    txtPassword.Enabled = True

End Sub
'========================================
' BOTÓN ENTRAR
'========================================
Private Sub btnEntrar_Click()

    On Error GoTo ErrorLogin

    Dim hoja As Worksheet
    Dim i As Long
    Dim ultimaFila As Long

    Dim usuario As String
    Dim password As String

    Dim usuarioHoja As String
    Dim passwordHoja As String

    Dim encontrado As Boolean

    Set hoja = ThisWorkbook.Sheets("Usuarios")

    usuario = Trim(CStr(txtUsuario.Value))
    password = Trim(CStr(txtPassword.Value))

    If usuario = "" Or password = "" Then

        MsgBox "Ingresa usuario y contraseña", vbExclamation
        Exit Sub

    End If

    ultimaFila = hoja.Cells(hoja.Rows.Count, 1).End(xlUp).Row

    encontrado = False

    For i = 2 To ultimaFila

        usuarioHoja = Trim(CStr(hoja.Cells(i, 1).Value))
        passwordHoja = Trim(CStr(hoja.Cells(i, 2).Value))
        
        'MsgBox usuarioHoja & " | " & passwordHoja

        If StrComp(usuarioHoja, usuario, vbTextCompare) = 0 And _
           StrComp(passwordHoja, password, vbTextCompare) = 0 Then

            UsuarioActivo = usuarioHoja
            RolActivo = hoja.Cells(i, 3).Value

            encontrado = True

            Exit For

        End If

    Next i

    If encontrado Then

        intentos = 0

        Call RegistrarAcceso(usuario, "Acceso correcto")

        MsgBox "Bienvenido " & UsuarioActivo, vbInformation

        Unload Me

        FormCapturaFrenos.Show

    Else

        intentos = intentos + 1

        Call RegistrarAcceso(usuario, "Intento fallido")

        MsgBox "Usuario o contraseña incorrectos", vbCritical

        txtPassword.Value = ""

        If intentos >= 3 Then

            MsgBox "Sistema bloqueado", vbCritical

            ThisWorkbook.Close SaveChanges:=False

        End If

    End If

    Exit Sub

ErrorLogin:

    MsgBox "Error: " & Err.Description, vbCritical

End Sub
'========================================
' CERRAR FORMULARIO
'========================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    ' SI CIERRAN CON LA X
    If CloseMode = 0 Then

        MsgBox "Debes iniciar sesión", vbCritical

        ThisWorkbook.Close SaveChanges:=False

    End If

End Sub
