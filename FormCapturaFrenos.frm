VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormCapturaFrenos 
   Caption         =   "Sistema Aduanas - Captura Frenos"
   ClientHeight    =   9468.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   15660
   OleObjectBlob   =   "FormCapturaFrenos.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "FormCapturaFrenos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filaSeleccionada As Long
Private Sub TextBox5_Change()

End Sub
Private Sub CommandButton4_Click()

End Sub

Private Sub btnBuscarGeneral_Click()

End Sub
Private Sub Label23_Click()

End Sub
Private Sub UserForm_Initialize()

    ' Turnos
    cboTurno.AddItem "Matutino"
    cboTurno.AddItem "Vespertino"

    ' Líneas
    Dim i As Integer
    For i = 1 To 5
        cboLinea.AddItem i
    Next i

    ' Cargar partes desde catálogo
    Dim hoja As Worksheet
    Dim x As Long
    
    Set hoja = Sheets("Catalogo Partes")
    
    For x = 2 To hoja.Cells(Rows.Count, 1).End(xlUp).Row
        cboParte.AddItem hoja.Cells(x, 1).Value
    Next x
    
    ' FILTROS
    cboFiltro.AddItem "ID"
    cboFiltro.AddItem "Cliente"
    cboFiltro.AddItem "Parte"
    cboFiltro.AddItem "Turno"
    cboFiltro.AddItem "Linea"
    cboFiltro.AddItem "S.O."
    cboFiltro.AddItem "Fecha"
    
    cboFiltro.ListIndex = 0
    
    Call CargarTodosLosRegistros

End Sub
Private Sub btnGuardar_Click()

    Dim hoja As Worksheet
    Dim fila As Long
    
    Set hoja = Sheets("Frenos")
    
    fila = hoja.Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    hoja.Cells(fila, 1).Value = GenerarID("F-", "Frenos")
    hoja.Cells(fila, 2).Value = txtFecha.Value
    hoja.Cells(fila, 3).Value = cboReferencia.Value
    hoja.Cells(fila, 4).Value = cboTurno.Value
    hoja.Cells(fila, 5).Value = cboLinea.Value
    hoja.Cells(fila, 6).Value = cboParte.Value
    hoja.Cells(fila, 7).Value = txtQty.Value
    hoja.Cells(fila, 8).Value = cboRev.Value
    hoja.Cells(fila, 9).Value = txtSO.Value
    hoja.Cells(fila, 10).Value = cboCustomer.Value
    hoja.Cells(fila, 11).Value = cboLocation.Value
    hoja.Cells(fila, 12).Value = txtExport.Value
    
    ' ?? NUEVO: fuente y notas
    hoja.Cells(fila, 16).Value = txtFuente.Value
    hoja.Cells(fila, 17).Value = txtNotas.Value
    
    ' ?? FECHA CAPTURA
    hoja.Cells(fila, 18).Value = Now
    
    ' ?? USUARIO LOGUEADO (AQUÍ ESTÁ EL CAMBIO IMPORTANTE)
    hoja.Cells(fila, 19).Value = UsuarioActivo

    ' ?? CONEXIÓN AUTOMÁTICA
    Call RegistrarProduccion
    Call RegistrarBDProduccion
    Call RegistrarControl

    MsgBox "Guardado y conectado correctamente", vbInformation

End Sub

' ===============================
' BUSCAR
' ===============================
Private Sub btnBuscar_Click()

    Dim hoja As Worksheet
    Dim i As Long
    
    Set hoja = Sheets("Frenos")
    
    For i = 2 To hoja.Cells(Rows.Count, 1).End(xlUp).Row
    
        If hoja.Cells(i, 1).Value = txtID.Value Then
            
            filaSeleccionada = i
            
            txtFecha.Value = hoja.Cells(i, 2).Value
            cboReferencia.Value = hoja.Cells(i, 3).Value
            cboTurno.Value = hoja.Cells(i, 4).Value
            cboLinea.Value = hoja.Cells(i, 5).Value
            cboParte.Value = hoja.Cells(i, 6).Value
            txtQty.Value = hoja.Cells(i, 7).Value
            cboRev.Value = hoja.Cells(i, 8).Value
            txtSO.Value = hoja.Cells(i, 9).Value
            cboCustomer.Value = hoja.Cells(i, 10).Value
            cboLocation.Value = hoja.Cells(i, 11).Value
            txtExport.Value = hoja.Cells(i, 12).Value
            txtProd1.Value = hoja.Cells(i, 13).Value
            txtProd2.Value = hoja.Cells(i, 14).Value
            txtBalance.Value = hoja.Cells(i, 15).Value
            txtFuente.Value = hoja.Cells(i, 16).Value
            txtNotas.Value = hoja.Cells(i, 17).Value
            
            MsgBox "Registro encontrado"
            Exit Sub
            
        End If
        
    Next i

    MsgBox "No encontrado"

End Sub
' ===============================
' ACTUALIZAR
' ===============================
Private Sub btnActualizar_Click()

    If filaSeleccionada = 0 Then
        MsgBox "Busca un registro primero"
        Exit Sub
    End If
    
    Dim hoja As Worksheet
    Set hoja = Sheets("Frenos")
    
    hoja.Cells(filaSeleccionada, 2).Value = txtFecha.Value
    hoja.Cells(filaSeleccionada, 3).Value = cboReferencia.Value
    hoja.Cells(filaSeleccionada, 4).Value = cboTurno.Value
    hoja.Cells(filaSeleccionada, 5).Value = cboLinea.Value
    hoja.Cells(filaSeleccionada, 6).Value = cboParte.Value
    hoja.Cells(filaSeleccionada, 7).Value = txtQty.Value
    hoja.Cells(filaSeleccionada, 8).Value = cboRev.Value
    hoja.Cells(filaSeleccionada, 9).Value = txtSO.Value
    hoja.Cells(filaSeleccionada, 10).Value = cboCustomer.Value
    hoja.Cells(filaSeleccionada, 11).Value = cboLocation.Value
    hoja.Cells(filaSeleccionada, 12).Value = txtExport.Value
    hoja.Cells(filaSeleccionada, 13).Value = txtProd1.Value
    hoja.Cells(filaSeleccionada, 14).Value = txtProd2.Value
    hoja.Cells(filaSeleccionada, 15).Value = txtBalance.Value
    hoja.Cells(filaSeleccionada, 16).Value = txtFuente.Value
    hoja.Cells(filaSeleccionada, 17).Value = txtNotas.Value

    MsgBox "Registro actualizado"

End Sub
' ===============================
' ELIMINAR
' ===============================
Private Sub btnEliminar_Click()

    'If filaSeleccionada = 0 Then
        'MsgBox "Busca un registro primero"
        'Exit Sub
    'End If
    
    'Sheets("Frenos").Rows(filaSeleccionada).Delete
    
    'MsgBox "Registro eliminado"
    
    'Call LimpiarCampos
    


    ' ?? VALIDAR PERMISOS
    If RolActivo <> "Admin" Then
        MsgBox "Solo el administrador puede eliminar registros", vbCritical
        Exit Sub
    End If

    ' ?? VALIDAR QUE HAYA REGISTRO SELECCIONADO
    If filaSeleccionada = 0 Then
        MsgBox "Busca un registro primero", vbExclamation
        Exit Sub
    End If

    ' ?? CONFIRMACIÓN
    If MsgBox("¿Seguro que deseas eliminar este registro?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If

    ' ??? ELIMINAR
    Sheets("Frenos").Rows(filaSeleccionada).Delete

    MsgBox "Registro eliminado correctamente", vbInformation

    ' ?? LIMPIAR FORMULARIO
    Call LimpiarCampos

End Sub
' ===============================
' LIMPIAR
' ===============================
Private Sub btnLimpiar_Click()
    Call LimpiarCampos
End Sub

Private Sub LimpiarCampos()

    txtID = ""
    txtFecha = ""
    cboReferencia = ""
    cboTurno = ""
    cboLinea = ""
    cboParte = ""
    txtQty = ""
    cboRev = ""
    txtSO = ""
    cboCustomer = ""
    cboLocation = ""
    txtExport = ""
    txtProd1 = ""
    txtProd2 = ""
    txtBalance = ""
    txtFuente = ""
    txtNotas = ""
    
End Sub
Sub RegistrarProduccion()

    Dim hoja As Worksheet
    Dim fila As Long
    
    Set hoja = Sheets("Producción")
    
    fila = hoja.Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    hoja.Cells(fila, 1).Value = txtFecha.Value
    hoja.Cells(fila, 2).Value = cboReferencia.Value
    hoja.Cells(fila, 3).Value = cboTurno.Value
    hoja.Cells(fila, 4).Value = cboLinea.Value
    hoja.Cells(fila, 5).Value = "Freno"
    hoja.Cells(fila, 6).Value = cboParte.Value
    hoja.Cells(fila, 7).Value = txtQty.Value
    hoja.Cells(fila, 8).Value = cboCustomer.Value
    hoja.Cells(fila, 9).Value = txtSO.Value
    hoja.Cells(fila, 10).Value = cboRev.Value
    hoja.Cells(fila, 11).Value = txtFuente.Value

End Sub
Sub RegistrarBDProduccion()

    Dim hoja As Worksheet
    Dim fila As Long
    
    Set hoja = Sheets("bd_producción")
    
    fila = hoja.Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    hoja.Cells(fila, 1).Value = GenerarID("FGA-", "bd_producción")
    hoja.Cells(fila, 2).Value = txtFecha.Value
    hoja.Cells(fila, 3).Value = cboCustomer.Value
    hoja.Cells(fila, 4).Value = cboParte.Value
    hoja.Cells(fila, 5).Value = cboParte.Value
    hoja.Cells(fila, 6).Value = txtQty.Value
    hoja.Cells(fila, 7).Value = cboLinea.Value
    hoja.Cells(fila, 8).Value = txtFuente.Value
    hoja.Cells(fila, 9).Value = "Pendiente"

End Sub
Sub RegistrarControl()

    Dim hoja As Worksheet
    Dim fila As Long
    
    Set hoja = Sheets("Control Captura")
    
    fila = hoja.Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    hoja.Cells(fila, 1).Value = txtFecha.Value
    hoja.Cells(fila, 2).Value = cboReferencia.Value
    hoja.Cells(fila, 3).Value = Now
    hoja.Cells(fila, 4).Value = UsuarioActivo

End Sub
Sub CargarTodosLosRegistros()

    Dim ws As Worksheet
    Dim i As Long
    Dim ultimaFila As Long

    Set ws = Sheets("BD_FRENOS")

    lstResultados.Clear

    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To ultimaFila

        lstResultados.AddItem ws.Cells(i, 1).Value
        lstResultados.List(lstResultados.ListCount - 1, 1) = ws.Cells(i, 2).Value
        lstResultados.List(lstResultados.ListCount - 1, 2) = ws.Cells(i, 6).Value
        lstResultados.List(lstResultados.ListCount - 1, 3) = ws.Cells(i, 10).Value
        lstResultados.List(lstResultados.ListCount - 1, 4) = ws.Cells(i, 7).Value
        lstResultados.List(lstResultados.ListCount - 1, 5) = ws.Cells(i, 4).Value
        lstResultados.List(lstResultados.ListCount - 1, 6) = ws.Cells(i, 5).Value
        lstResultados.List(lstResultados.ListCount - 1, 7) = ws.Cells(i, 9).Value
        lstResultados.List(lstResultados.ListCount - 1, 8) = ws.Cells(i, 12).Value
        lstResultados.List(lstResultados.ListCount - 1, 9) = ws.Cells(i, 17).Value

    Next i

End Sub
