Attribute VB_Name = "Módulo2"
Sub Listadesplegable10_Cambiar()
cmbCliente.List = Sheets("LISTAS").Range("A2:A100").Value
End Sub
Sub AgregarCliente()

    Dim ws As Worksheet
    Set ws = Sheets("LISTAS")
    
    Dim fila As Long
    fila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ws.Cells(fila, 1).Value = InputBox("Nuevo cliente")

End Sub
