VERSION 5.00
Begin VB.Form frmOpePagoProvSel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de Proveedor"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4635
   Icon            =   "frmOpePagoProvSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboProveedor 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   4140
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "&Seleccionar"
      Height          =   320
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000E&
      Height          =   1065
      Left            =   105
      Top             =   105
      Width           =   4410
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H8000000C&
      Height          =   1035
      Left            =   120
      Top             =   120
      Width           =   4365
   End
End
Attribute VB_Name = "frmOpePagoProvSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'** Nombre : frmOpePagoProvSel
'** Descripción : Formulario para la selección de Proveedores segun ERS062-2013
'** Creación : EJVG, 20131120 11:00:00 AM
'******************************************************************************
Option Explicit
Dim fMatProveedor() As String
Dim fnTpoPago As LogTipoPagoComprobante
Dim fsOpcion As String

Private Sub cmdSeleccionar_Click()
    If cboProveedor.ListIndex = -1 Then
        MsgBox "Ud. debe de elegir una opción de la lista de Proveedores", vbInformation, "Aviso"
        Exit Sub
    End If
    fsOpcion = Trim(Right(cboProveedor.Text, 13))
    Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    Dim I As Integer
    cboProveedor.Clear
    If UBound(fMatProveedor, 2) >= 1 Then
        If fnTpoPago = gPagoCuentaCMAC Or fnTpoPago = gPagoTransferencia Then 'ListIndex=0
            cboProveedor.AddItem "TODOS" & Space(150) & "TODOS"
        End If
        For I = 1 To UBound(fMatProveedor, 2)
            cboProveedor.AddItem fMatProveedor(1, I) & Space(150) & fMatProveedor(0, I)
        Next
        If fnTpoPago = gPagoCuentaCMAC Or fnTpoPago = gPagoTransferencia Then
            cboProveedor.ListIndex = 0
        End If
        cmdSeleccionar.Default = True
    End If
End Sub
Public Function Inicio(ByRef pMatProveedor() As String, ByVal pnTpoPago As LogTipoPagoComprobante) As String
    Dim I As Integer
    fMatProveedor = pMatProveedor
    fnTpoPago = pnTpoPago
    Show 1
    Inicio = fsOpcion
End Function
Private Sub cboProveedor_KeyPress(KeyAscii As Integer)
    If cboProveedor.ListIndex > -1 Then
        If KeyAscii = 13 Then
            cmdSeleccionar.SetFocus
        End If
    End If
End Sub
