VERSION 5.00
Begin VB.Form frmCapRegVouDepBus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscando Voucher"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5010
   Icon            =   "frmCapRegVouDepBus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtMoneda 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox txtCliente 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   4095
   End
   Begin VB.ComboBox cboVoucher 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.ComboBox cboIF 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label5 
      Caption         =   "Monto"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Moneda"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Voucher"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Banco"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmCapRegVouDepBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*** Nombre : frmCapRegVouDepBus
'*** Descripción : Formulario para registrar el vouchert de Depósito.
'*** Creación : ELRO, 20120530 07:40:21 PM, según OYP-RFC024-2012
'********************************************************************
Option Explicit

Private fsMoneda As String
Private fnTipMot As Integer

Private fsGlosa As String
Private fsIF As String
Private fsVoucher As String
Private fnMonto As Currency
Private fsPersCod As String
Private fsCodOpe As String
Private fnMovNroRVD As Long
Private fnMovNroPen As Long
Private fsPersNombre As String
Private fsPersDireccion As String
Private fsPersDocumento As String
Private fsDetalle As String 'EJVG20130912
Private fsDetalleIng As String 'EJVG20130914
Private fsIFTipo As String 'RIRO20170714

Public Sub iniciarFormulario(ByVal psMoneda As String, _
                             ByVal pnTipMot As Integer, _
                             ByRef psGlosa As String, _
                             ByRef psIF As String, _
                             ByRef psVoucher As String, _
                             ByRef pnMonto As Currency, _
                             ByRef psPersCod As String, _
                             ByRef pnMovNroRVD As Long, _
                             ByRef pnMovNroPen As Long, _
                             Optional ByRef psDetalle As String = "")
fsMoneda = psMoneda
fnTipMot = pnTipMot
Me.Show 1
psGlosa = fsGlosa
psIF = fsIF
psVoucher = fsVoucher
pnMonto = fnMonto
psPersCod = fsPersCod
pnMovNroRVD = fnMovNroRVD
pnMovNroPen = fnMovNroPen
psDetalle = fsDetalle 'EJVG20130912
limpiarCampos
fsMoneda = ""
fnTipMot = 0
End Sub

Public Sub iniciarFormularioDeposito(ByVal psMoneda As String, _
                                     ByVal pnTipMot As Integer, _
                                     ByRef psGlosa As String, _
                                     ByRef psIF As String, _
                                     ByRef psVoucher As String, _
                                     ByRef pnMonto As Currency, _
                                     ByRef psPersCod As String, _
                                     ByRef pnMovNroRVD As Long, _
                                     ByRef pnMovNroPen As Long, _
                                     ByRef psPersNombre As String, _
                                     ByRef psPersDireccion As String, _
                                     ByRef psPersDocumento As String, _
                                     Optional ByVal psDetalleIng As String = "", _
                                     Optional ByRef psIFTipo As String = "")
fsMoneda = psMoneda
fnTipMot = pnTipMot
fsDetalleIng = psDetalleIng 'EJVG20130914
Me.Show 1
psGlosa = fsGlosa
psIF = fsIF
psVoucher = fsVoucher
pnMonto = fnMonto
psPersCod = fsPersCod
pnMovNroRVD = fnMovNroRVD
pnMovNroPen = fnMovNroPen
psPersNombre = fsPersNombre
psPersDireccion = fsPersDireccion
psPersDocumento = fsPersDocumento
psIFTipo = fsIFTipo 'RIRO20170714
limpiarCampos
fsMoneda = ""
fnTipMot = 0
End Sub

Private Sub cargarIFs()
    Dim oNCOMCaptaGenerales As NCOMCaptaGenerales
    Set oNCOMCaptaGenerales = New NCOMCaptaGenerales
    Dim rsIFs As ADODB.Recordset
    Set rsIFs = New ADODB.Recordset
    
    'Set rsIFs = oNCOMCaptaGenerales.obtenerVoucherIFSinOperacion(Right(gsCodAge, 2), fsMoneda, fnTipMot)
    Set rsIFs = oNCOMCaptaGenerales.obtenerVoucherIFSinOperacion(Right(gsCodAge, 2), fsMoneda, fnTipMot, fsDetalleIng) 'EJVG20130914
    
    cboIF.Clear
    Do While Not rsIFs.EOF
        cboIF.AddItem rsIFs!cIF
        rsIFs.MoveNext
    Loop
    cboIF.ListIndex = -1
    
    Set rsIFs = Nothing
    Set oNCOMCaptaGenerales = Nothing
End Sub

Private Sub cargarVouchers()
    Dim oNCOMCaptaGenerales As NCOMCaptaGenerales
    Set oNCOMCaptaGenerales = New NCOMCaptaGenerales
    Dim rsVouchers As ADODB.Recordset
    Set rsVouchers = New ADODB.Recordset
    Dim lcIF As String
    
    lcIF = Trim(Right(cboIF, 50))
    
    'Set rsVouchers = oNCOMCaptaGenerales.obtenerVoucherNroIFSinOperacion(Right(gsCodAge, 2), fsMoneda, fnTipMot, Left(lcIF, 2), Mid(lcIF, 4, 13))
    Set rsVouchers = oNCOMCaptaGenerales.obtenerVoucherNroIFSinOperacion(Right(gsCodAge, 2), fsMoneda, fnTipMot, Left(lcIF, 2), Mid(lcIF, 4, 13), fsDetalleIng) 'EJVG20130914
    
    cboVoucher.Clear
    Do While Not rsVouchers.EOF
        cboVoucher.AddItem rsVouchers!cNroVou & Space(100) & rsVouchers!cCtaIFCod
        rsVouchers.MoveNext
    Loop
    cboVoucher.ListIndex = -1
    
    lcIF = ""
    Set rsVouchers = Nothing
    Set oNCOMCaptaGenerales = Nothing
End Sub

Private Sub limpiarCampos()
    txtCliente = ""
    txtMoneda = ""
    txtMonto = "0.00"
    fsGlosa = ""
    fsIF = ""
    fsVoucher = ""
    fnMonto = 0#
    fsPersCod = ""
    fsCodOpe = ""
    fnMovNroRVD = 0
    fnMovNroPen = 0
    fsDetalle = "" 'EJVG20130912
    fsPersNombre = ""
    fsPersDireccion = ""
    fsPersDocumento = ""
    fsIFTipo = "" 'RIRO20170714
End Sub



Private Sub CargarCliente()
    Dim oNCOMCaptaGenerales As NCOMCaptaGenerales
    Set oNCOMCaptaGenerales = New NCOMCaptaGenerales
    Dim rsCliente As ADODB.Recordset
    Set rsCliente = New ADODB.Recordset
    Dim lcIF, lcVoucher As String
    
    lcIF = Trim(Right(cboIF, 50))
    lcVoucher = Trim(Right(cboVoucher, 50))
    
    Set rsCliente = oNCOMCaptaGenerales.obtenerVoucherClienteNroIFSinOperacion(Right(gsCodAge, 2), fsMoneda, fnTipMot, Left(lcIF, 2), Mid(lcIF, 4, 13), lcVoucher, cboVoucher)
    
    If Not rsCliente.BOF And Not rsCliente.EOF Then
        txtCliente = rsCliente!cPersNombre
        txtMoneda = rsCliente!cMoneda
        txtMonto = Format$(rsCliente!nMonVou, "#,##0.00")
        
        fsGlosa = rsCliente!cGlosa
        fsIF = Trim(Left(cboIF, 30))
        fsVoucher = cboVoucher
        fnMonto = rsCliente!nMonVou
        fsPersCod = rsCliente!cPersCod
        fsCodOpe = rsCliente!cOpecod
        fnMovNroRVD = rsCliente!nMovNroRVD
        fnMovNroPen = rsCliente!nMovNroPen
        fsDetalle = rsCliente!lsDetalle 'EJVG20130912
        fsPersNombre = rsCliente!cPersNombre
        fsPersDireccion = rsCliente!cPersDireccDomicilio
        fsPersDocumento = rsCliente!cPersIDnro
        fsIFTipo = lcIF 'RIRO20171013
    Else
        limpiarCampos
    End If
    
    lcIF = ""
    Set rsCliente = Nothing
    Set oNCOMCaptaGenerales = Nothing
End Sub

Private Function Validar() As Boolean

Validar = False

If cboIF.ListIndex = -1 Then
    MsgBox "Elija un Banco.", vbInformation, "Aviso"
    cboVoucher.SetFocus
    Exit Function
End If

If cboVoucher.ListIndex = -1 Then
    MsgBox "Elija un Voucher.", vbInformation, "Aviso"
    cboVoucher.SetFocus
    Exit Function
End If

Validar = True
End Function

Private Sub cboIF_Click()
    cargarVouchers
    limpiarCampos
End Sub

Private Sub cboVoucher_Click()
   CargarCliente
End Sub

Private Sub cmdAceptar_Click()
    If Validar = False Then Exit Sub
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    fsMoneda = ""
    fnTipMot = 0
    limpiarCampos
    Unload Me
End Sub

Private Sub Form_Load()
    limpiarCampos
    cargarIFs
    cargarVouchers
    CargarCliente
End Sub
