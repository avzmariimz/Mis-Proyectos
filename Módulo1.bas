Attribute VB_Name = "Módulo1"
Public UsuarioActivo As String
Public RolActivo As String
Sub GuardarProduccion()

    Dim prod As Worksheet
    Set prod = Sheets("BD_PRODUCCION")
    
    Dim fila As Long
    fila = prod.Cells(prod.Rows.Count, 1).End(xlUp).Row + 1
    
    prod.Cells(fila, 1).Value = txtFecha
    prod.Cells(fila, 2).Value = txtReferencia
    prod.Cells(fila, 3).Value = cmbTurno
    prod.Cells(fila, 4).Value = cmbLinea
    prod.Cells(fila, 5).Value = lblTipo.Caption
    prod.Cells(fila, 6).Value = cmbParte
    prod.Cells(fila, 7).Value = txtCantidad
    prod.Cells(fila, 8).Value = cmbCliente
    prod.Cells(fila, 9).Value = txtOrder
    prod.Cells(fila, 10).Value = txtREV
    prod.Cells(fila, 11).Value = Now

End Sub
Sub GuardarConversionFG()

    Dim fg As Worksheet
    Set fg = Sheets("bd_producción – Conversión FG a ASSY")
    
    Dim fila As Long
    fila = fg.Cells(fg.Rows.Count, 1).End(xlUp).Row + 1
    
    fg.Cells(fila, 1).Value = "FGA-" & Format(fila - 1, "0000")
    fg.Cells(fila, 2).Value = txtFecha
    fg.Cells(fila, 3).Value = cmbCliente
    fg.Cells(fila, 4).Value = cmbParte
    fg.Cells(fila, 5).Value = cmbParte
    fg.Cells(fila, 6).Value = txtCantidad
    fg.Cells(fila, 7).Value = cmbLinea
    fg.Cells(fila, 8).Value = Now
    fg.Cells(fila, 9).Value = "Pendiente"

End Sub
Sub GuardarReporteAduanas()

    Dim rep As Worksheet
    Set rep = Sheets("REPORTES_ADUANAS")
    
    Dim fila As Long
    fila = rep.Cells(rep.Rows.Count, 1).End(xlUp).Row + 1
    
    rep.Cells(fila, 1).Value = "FGA-" & Format(fila - 1, "0000")
    rep.Cells(fila, 2).Value = txtFecha
    rep.Cells(fila, 3).Value = txtReferencia
    rep.Cells(fila, 4).Value = cmbCliente
    rep.Cells(fila, 5).Value = txtExportDate
    rep.Cells(fila, 6).Value = "Pendiente"
    rep.Cells(fila, 7).Value = "Factura vinculada"
    rep.Cells(fila, 8).Value = "Tracking #"
    rep.Cells(fila, 9).Value = "Exportación programada"

End Sub
Sub ActualizarControlPartes()

    Dim ws As Worksheet
    Set ws = Sheets("Control de Partes")
    
    Dim i As Long
    Dim ultima As Long
    
    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To ultima
        
        ' Si cantidad = 0 ? agotado
        If ws.Cells(i, 4).Value <= 0 Then
            ws.Cells(i, 5).Value = "AGOTADO"
        End If
        
        ' Si hay piezas ? disponible
        If ws.Cells(i, 4).Value > 0 Then
            ws.Cells(i, 5).Value = "DISPONIBLE"
        End If
        
    Next i

End Sub
Sub AlertasVencimiento()

    Dim ws As Worksheet
    Set ws = Sheets("BD_FRENOS")
    
    Dim i As Long
    Dim fechaExp As Date
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        If ws.Cells(i, 12).Value <> "" Then 'Columna Export Date
            
            fechaExp = ws.Cells(i, 12).Value
            
            If fechaExp - Date <= 5 Then
                MsgBox "? Pedido próximo a vencer: " & ws.Cells(i, 6).Value, vbExclamation
            End If
            
        End If
        
    Next i

End Sub
Sub BuscarTrazabilidad()

    Dim parte As String
    parte = InputBox("Ingresa No. de Parte")
    
    Dim ws As Worksheet
    Dim i As Long
    
    For Each ws In ThisWorkbook.Worksheets
        
        For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            If ws.Cells(i, 1).Value = parte Or ws.Cells(i, 6).Value = parte Then
                
                MsgBox "Encontrado en: " & ws.Name & " fila: " & i
                
            End If
            
        Next i
        
    Next ws

End Sub
Sub DashboardProduccion()

    Dim ws As Worksheet
    Set ws = Sheets("BD_PRODUCCION")
    
    Dim total As Double
    Dim i As Long
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        total = total + ws.Cells(i, 7).Value
    Next i
    
    Sheets("REPORTES_ADUANAS").Range("A1").Value = "Total Producción"
    Sheets("REPORTES_ADUANAS").Range("B1").Value = total
    
    MsgBox "Producción total: " & total
    
End Sub
Sub RegistrarCaptura()

    Dim ws As Worksheet
    Set ws = Sheets("CONTROL_CAPTURA")
    
    Dim fila As Long
    fila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ws.Cells(fila, 1).Value = Date
    ws.Cells(fila, 2).Value = "Captura sistema"
    ws.Cells(fila, 3).Value = Now
    ws.Cells(fila, 4).Value = Environ("Username")
    ws.Cells(fila, 5).Value = "OK"

End Sub
Sub BuscarRegistro()

    Dim id As String
    id = InputBox("Ingrese ID")
    
    Dim ws As Worksheet
    Set ws = Sheets("BD_FRENOS")
    
    Dim i As Long
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        If ws.Cells(i, 1).Value = id Then
            
            MsgBox "Encontrado en fila " & i
            
        End If
        
    Next i

End Sub
Sub EliminarRegistro()

    Dim id As String
    id = InputBox("ID a eliminar")
    
    Dim ws As Worksheet
    Set ws = Sheets("BD_FRENOS")
    
    Dim i As Long
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        If ws.Cells(i, 1).Value = id Then
            
            ws.Rows(i).Delete
            MsgBox "Eliminado"
            Exit Sub
            
        End If
        
    Next i

End Sub
Sub GuardarTodo()

    Call GuardarProduccion
    Call GuardarConversionFG
    Call GuardarReporteAduanas
    Call RegistrarCaptura
    Call AlertasVencimiento
    
    MsgBox "Registro completo y trazabilidad actualizada", vbInformation

End Sub
Sub AbrirFormulario()
    UserForm1.Show
End Sub
Sub IrControlPartes()
    Sheets("Control de Partes").Activate
End Sub
Sub IrReportes()
    Sheets("REPORTES_ADUANAS").Activate
End Sub
Sub ActualizarTodo()
    Call DashboardProduccion
    Call AlertasVencimiento
    Call ActualizarControlPartes
    MsgBox "Sistema actualizado"
End Sub
Sub LimpiarFormulario()

    Dim ctrl As Object
    
    For Each ctrl In UserForm1.Controls
        If TypeName(ctrl) = "TextBox" Then ctrl.Value = ""
        If TypeName(ctrl) = "ComboBox" Then ctrl.Value = ""
    Next ctrl

End Sub
Sub IrFrenos()
    Sheets("BD_FRENOS").Activate
End Sub

Sub IrBases()
    Sheets("BD_BASES").Activate
End Sub

Sub IrProduccion()
    Sheets("BD_PRODUCCION").Activate
End Sub
Sub IrCatalogo()
    Sheets("CATALOGO_PARTES").Activate
End Sub
Sub NuevoFreno()
    Range("A2:Q2").ClearContents
End Sub
Sub GuardarFreno()

Dim fila As Long
fila = Sheets("Frenos").Cells(Rows.Count, 1).End(xlUp).Row + 1

Cells(fila, 1).Value = "F-" & Format(fila, "0000")
Cells(fila, 2).Value = Range("B2").Value
Cells(fila, 3).Value = Range("C2").Value
Cells(fila, 4).Value = Range("D2").Value
Cells(fila, 5).Value = Range("E2").Value
Cells(fila, 6).Value = Range("F2").Value
Cells(fila, 7).Value = Range("G2").Value
Cells(fila, 8).Value = Range("H2").Value

MsgBox "Registro guardado", vbInformation

End Sub
Sub EnviarAProduccion()

Dim fila As Long
fila = Sheets("BD_PRODUCCION").Cells(Rows.Count, 1).End(xlUp).Row + 1

Sheets("Produccion").Cells(fila, 1).Value = Sheets("Frenos").Range("B2").Value
Sheets("Produccion").Cells(fila, 2).Value = Sheets("Frenos").Range("C2").Value

End Sub
Sub RevisarVencimientos()

Dim i As Long
For i = 2 To Sheets("Control de Partes").Cells(Rows.Count, 1).End(xlUp).Row

    If Sheets("Control partes").Cells(i, 4).Value < Date Then
        MsgBox "Material vencido en fila " & i
    End If

Next i

End Sub
Sub AbrirFrenos()
    FormCapturaFrenos.Show
End Sub
Private Sub UserForm_Initialize()

    Dim hoja As Worksheet
    Dim i As Long
    
    Set hoja = Sheets("Catalogo Partes")
    
    For i = 2 To hoja.Cells(Rows.Count, 1).End(xlUp).Row
        cboParte.AddItem hoja.Cells(i, 1).Value
    Next i

End Sub
Sub Alertas()

    Dim hoja As Worksheet
    Dim i As Long
    
    Set hoja = Sheets("Control Partes")
    
    For i = 2 To hoja.Cells(Rows.Count, 1).End(xlUp).Row
        
        If hoja.Cells(i, 5).Value = "Vencido" Then
            MsgBox "Parte vencida: " & hoja.Cells(i, 1).Value
        End If
        
    Next i

End Sub
Sub ActualizarDashboard()
    ThisWorkbook.RefreshAll
    MsgBox "Dashboard actualizado"
End Sub
Sub VerificarVencimientos()

    Dim hoja As Worksheet
    Dim i As Long
    
    Set hoja = Sheets("Frenos")
    
    For i = 2 To hoja.Cells(Rows.Count, 1).End(xlUp).Row
    
        If hoja.Cells(i, 12).Value < Date Then
            hoja.Cells(i, 12).Interior.Color = RGB(255, 0, 0)
        End If
        
    Next i

End Sub
'Function GenerarID(prefijo As String, hoja As String) As String
Function GenerarID(prefijo As String, hojaDestino As String) As String
    Dim ultimaFila As Long
    ultimaFila = Sheets(hojaDestino).Cells(Rows.Count, 1).End(xlUp).Row
    ' Genera un número consecutivo basado en la última fila
    GenerarID = prefijo & Format(ultimaFila, "0000")
End Function
Sub RegistrarAcceso(usuario As String, estado As String)

    Dim hoja As Worksheet
    Dim fila As Long
    
    Set hoja = Sheets("Accesos")
    
    fila = hoja.Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    hoja.Cells(fila, 1).Value = Now
    hoja.Cells(fila, 2).Value = usuario
    hoja.Cells(fila, 3).Value = estado

End Sub
Sub AbrirMenu()
    frmMenu.Show
End Sub
Sub AbrirCatalogo()
    FormCatalogo.Show
End Sub
