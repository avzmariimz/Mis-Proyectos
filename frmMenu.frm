VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenu 
   Caption         =   "frmMenu"
   ClientHeight    =   2352
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10764
   OleObjectBlob   =   "frmMenu.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCaptura_Click()
    UserForm1.Show
End Sub
Private Sub btnControl_Click()
    Sheets("Control de Partes").Activate
End Sub
Private Sub btnDashboard_Click()
    UserForm1.Show
End Sub
Private Sub btnSalir_Click()
    Unload Me
End Sub
Private Sub btnTrazabilidad_Click()
    Call DashboardProduccion
End Sub
Private Sub UserForm_Click()

End Sub
