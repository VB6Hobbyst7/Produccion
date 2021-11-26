Attribute VB_Name = "MCargaMenu"

Public Type TMatmenu

    nId As Integer
    sCodigo As String
    sName As String
    sCaption As String
    sIndex As String
    nNumHijos As Integer
    bCheck As Boolean
    nPuntDer As Integer
    nPuntAbajo As Integer
    nNivel As Integer
End Type
Public MatMenuItems() As TMatmenu
Public MatOperac(2000, 5) As String
Public NroRegOpe As Integer




