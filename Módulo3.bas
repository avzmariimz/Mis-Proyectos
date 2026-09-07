Attribute VB_Name = "Módulo3"
Function GenerarID(prefijo As String, hoja As String) As String

    Dim ultimaFila As Long
    ultimaFila = Sheets(hoja).Cells(Rows.Count, 1).End(xlUp).Row
    
    GenerarID = prefijo & Format(ultimaFila, "0000")

End Function
