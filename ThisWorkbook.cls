VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    'PROTEGER HOJA
    'ThisWorkbook.Sheets("Frenos").Protect password:="1234"

    'ACTUALIZAR SISTEMA
    'Call AlertasVencimiento
    'Call ActualizarControlPartes

    'ABRIR LOGIN
    frmLogin.Show

End Sub
