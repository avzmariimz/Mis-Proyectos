VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3228
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   15876
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnGuardar_Click()

    Dim tipo As String
    tipo = lblTipo.Caption
    
    Dim fila As Long
    Dim id As String

'================ FRENOS =================
If tipo = "Freno" Then

    Dim frenos As Worksheet
    Set frenos = Sheets("BD_FRENOS")
    
    fila = frenos.Cells(frenos.Rows.Count, 1).End(xlUp).Row + 1
    
    id = "F-" & Format(fila - 1, "0000")
    
    frenos.Cells(fila, 1).Value = id
    frenos.Cells(fila, 2).Value = txtFecha
    frenos.Cells(fila, 3).Value = txtReferencia
    frenos.Cells(fila, 4).Value = cmbTurno
    frenos.Cells(fila, 5).Value = cmbLinea
    frenos.Cells(fila, 6).Value = cmbParte
    frenos.Cells(fila, 7).Value = txtCantidad
    frenos.Cells(fila, 8).Value = txtREV
    frenos.Cells(fila, 9).Value = txtOrder
    frenos.Cells(fila, 10).Value = cmbCliente
    frenos.Cells(fila, 11).Value = txtLocation
    frenos.Cells(fila, 12).Value = txtExportDate
    frenos.Cells(fila, 16).Value = "Sistema VBA"
    frenos.Cells(fila, 17).Value = txtNotas
    frenos.Cells(fila, 18).Value = Now
    frenos.Cells(fila, 19).Value = txtCapturado

Else

'================ BASES =================

    Dim bases As Worksheet
    Set bases = Sheets("BD_BASES")
    
    fila = bases.Cells(bases.Rows.Count, 1).End(xlUp).Row + 1
    
    id = "B-" & Format(fila - 1, "0000")
    
    bases.Cells(fila, 1).Value = id
    bases.Cells(fila, 2).Value = txtFecha
    bases.Cells(fila, 3).Value = txtReferencia
    bases.Cells(fila, 4).Value = cmbTurno
    bases.Cells(fila, 5).Value = cmbLinea
    bases.Cells(fila, 6).Value = cmbParte
    bases.Cells(fila, 7).Value = txtCantidad
    bases.Cells(fila, 8).Value = txtOrder
    bases.Cells(fila, 9).Value = cmbCliente
    bases.Cells(fila, 10).Value = txtLocation
    bases.Cells(fila, 11).Value = txtREV
    bases.Cells(fila, 14).Value = "Sistema VBA"
    bases.Cells(fila, 15).Value = Now
    bases.Cells(fila, 16).Value = txtCapturado

End If

Call GuardarProduccion
Call GuardarConversionFG
Call GuardarReporteAduanas
Call ActualizarControlPartes
Call AlertasVencimiento
Call RegistrarCaptura
MsgBox "Guardado completo ??"

End Sub

Private Sub cmbParte_Change()

    Dim cat As Worksheet
    Dim i As Long
    
    Set cat = Sheets("CATALOGO_PARTES")
    
    For i = 2 To cat.Cells(cat.Rows.Count, 1).End(xlUp).Row
        
        If cat.Cells(i, 1).Value = cmbParte.Value Then
            
            lblTipo.Caption = cat.Cells(i, 2).Value
            
        End If
        
    Next i
    
End Sub
Private Sub Label1_Click()

End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label13_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub UserForm_Initialize()

    cmbTurno.AddItem "Matutino"
    cmbTurno.AddItem "Vespertino"
    cmbTurno.AddItem "Nocturno"

    cmbLinea.AddItem "1"
    cmbLinea.AddItem "2"
    cmbLinea.AddItem "3"
    cmbLinea.AddItem "4"
    

    Dim cat As Worksheet
    Dim i As Long
    
    Set cat = Sheets("CATALOGO_PARTES")
    
    For i = 2 To cat.Cells(cat.Rows.Count, 1).End(xlUp).Row
        cmbParte.AddItem cat.Cells(i, 1).Value
    Next i
    
End Sub
