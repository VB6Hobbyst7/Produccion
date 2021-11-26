VERSION 5.00
Begin VB.Form frmChequeBusqueda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscando Cheque"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5040
   Icon            =   "frmChequeBusqueda.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboIFi 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.ComboBox cboCheque 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtCliente 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   4095
   End
   Begin VB.TextBox txtMoneda 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Banco"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   200
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Cheque"
      Height          =   255
      Left            =   150
      TabIndex        =   10
      Top             =   550
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   900
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Moneda"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1230
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Monto"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1600
      Width           =   495
   End
End
Attribute VB_Name = "frmChequeBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************
'** Nombre : frmChequeBusqueda
'** Descripción : Para Busqueda de Cheques en Operaciones segun TI-ERS126-2013
'** Creación : EJVG, 20140131 11:30:00 AM
'*****************************************************************************
Option Explicit

Dim fnMoneda As Moneda
Dim fnOperacion As TipoOperacionCheque
Dim fnMonto As Currency
Dim fsDetalle As String
Dim fsDetalleIng As String
Dim fsGlosa As String
Dim oDocRec As UDocRec

Private Sub Form_Load()
    Set oDocRec = New UDocRec
    limpiarCampos
    CargarIFis
    CargarCheques
    CargarCliente
End Sub
'Función para operaciones que esperan el 'DATO'->psDetalle como operaciones de apertura Ctas Ahorro, DPF y CTS
Public Function Iniciar(ByVal pnMoneda As Moneda, ByVal pnOperacion As TipoOperacionCheque) As UDocRec
    fnMoneda = pnMoneda
    fnOperacion = pnOperacion
    Show 1
    Set Iniciar = oDocRec
    Set oDocRec = Nothing
End Function
'Función para operaciones que ingresan el 'DATO'->psDetalleIng como busqueda como operaciones de deposito Ctas Ahorro/DPF, aumemento capital DPF y Pago Credito
Public Function iniciarBusqueda(ByVal pnMoneda As Moneda, ByVal pnOperacion As TipoOperacionCheque, ByVal psDetalleIng As String) As UDocRec
    fnMoneda = pnMoneda
    fnOperacion = pnOperacion
    fsDetalleIng = psDetalleIng
    Show 1
    Set iniciarBusqueda = oDocRec
    Set oDocRec = Nothing
End Function
Private Sub CargarIFis()
    Dim oDR As New NCOMDocRec
    Dim oRS As New ADODB.Recordset

    On Error GoTo ErrCargarIFis
    Screen.MousePointer = 11
    Set oRS = oDR.ListaIFiChequexBusqueda(fnMoneda, Right(gsCodAge, 2), fnOperacion, fsDetalleIng)
    cboIFi.Clear
    Do While Not oRS.EOF
        cboIFi.AddItem oRS!cPersNombre & space(200) & oRS!cperscod
        oRS.MoveNext
    Loop
    Screen.MousePointer = 0
    Set oRS = Nothing
    Set oDR = Nothing
    Exit Sub
ErrCargarIFis:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub CargarCheques()
    Dim oDR As New NCOMDocRec
    Dim oRS As New ADODB.Recordset
    Dim lsIFiCod As String
    
    On Error GoTo ErrCargarCheques
    Screen.MousePointer = 11
    lsIFiCod = Trim(Right(cboIFi.Text, 20))
    Set oRS = oDR.ListaChequexBusqueda(fnMoneda, Right(gsCodAge, 2), fnOperacion, lsIFiCod, fsDetalleIng)
    cboCheque.Clear
    lsIFiCod = Trim(Right(cboIFi.Text, 20))
    
    If Len(Trim(lsIFiCod)) > 0 Then
        Set oRS = oDR.ListaChequexBusqueda(fnMoneda, Right(gsCodAge, 2), fnOperacion, lsIFiCod, fsDetalleIng)
        Do While Not oRS.EOF
            cboCheque.AddItem oRS!cNroDoc & space(200) & oRS!nTpoDoc & "|" & oRS!cNroDoc & "|" & oRS!cperscod & "|" & oRS!cIFTpo & "|" & oRS!cIFCta
            oRS.MoveNext
        Loop
    End If
    
    Screen.MousePointer = 0
    Set oRS = Nothing
    Set oDR = Nothing
    Exit Sub
ErrCargarCheques:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub CargarCliente()
    Dim oDR As New NCOMDocRec
    Dim oRS As New ADODB.Recordset
    Dim lsIFiCod As String
    Dim Voucher() As String

    On Error GoTo ErrCargarCliente
    setDatosCliente
    Voucher = Split(Trim(Right(cboCheque.Text, 100)), "|")
    If UBound(Voucher) < 0 Then Exit Sub
    Screen.MousePointer = 11
    Set oRS = oDR.ChequexBusqueda(Voucher(0), Voucher(1), Voucher(2), Voucher(3), Voucher(4))
    If Not oRS.EOF Then
        setDatosCliente oRS!cGiradorNombre, oRS!cMoneda, oRS!nMonto, oRS!cGlosa, oRS!cDetalle
        cmdAceptar.Default = True
    End If
    Screen.MousePointer = 0
    Set oRS = Nothing
    Set oDR = Nothing
    Exit Sub
ErrCargarCliente:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub cboCheque_Click()
    CargarCliente
End Sub
Private Sub cboIFi_Click()
    cmdAceptar.Default = False
    CargarCheques
    CargarCliente
End Sub
Private Sub cboIFi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboCheque.Visible And cboCheque.Enabled Then cboCheque.SetFocus
    End If
End Sub
Private Sub cmdAceptar_Click()
    Dim Voucher() As String
    Dim lsPersCod As String, lsPersNombre As String
    On Error GoTo ErrCmdAceptar
    If Not ValidarSeleccion Then Exit Sub
    lsPersNombre = Trim(Left(cboIFi.Text, Len(cboIFi.Text) - 20))
    Voucher = Split(Trim(Right(cboCheque.Text, 100)), "|")
    Set oDocRec = New UDocRec
    oDocRec.CargaDatos Voucher(0), Voucher(1), Voucher(2), Voucher(3), Voucher(4), lsPersNombre, fsDetalle, fnMonto, fsGlosa
    Unload Me
    Exit Sub
ErrCmdAceptar:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdCancelar_Click()
    Set oDocRec = New UDocRec
    oDocRec.CargaDatos 0, "", "", "", "", "", "", 0#, ""
    Unload Me
End Sub
Private Sub limpiarCampos()
    cboIFi.ListIndex = -1
    cboCheque.ListIndex = -1
    txtCliente.Text = ""
    txtMoneda.Text = ""
    txtMonto.Text = "0.00"
End Sub
Private Function ValidarSeleccion() As Boolean
    ValidarSeleccion = True
    If cboIFi.ListIndex = -1 Then
        ValidarSeleccion = False
        MsgBox "Ud. debe seleccionar la Institución Financiera", vbInformation, "Aviso"
        If cboIFi.Visible And cboIFi.Enabled Then cboIFi.SetFocus
        Exit Function
    End If
    If cboCheque.ListIndex = -1 Then
        ValidarSeleccion = False
        MsgBox "Ud. debe seleccionar el Documento", vbInformation, "Aviso"
        If cboCheque.Visible And cboCheque.Enabled Then cboCheque.SetFocus
        Exit Function
    End If
End Function
Private Sub setDatosCliente(Optional ByVal psGiradorNombre As String = "", Optional ByVal psMoneda As String = "", Optional ByVal pnMonto As Currency = 0#, Optional ByVal psGlosa As String = "", Optional ByVal psDetalle As String = "")
    txtCliente.Text = psGiradorNombre
    txtMoneda.Text = psMoneda
    txtMonto.Text = Format(pnMonto, gsFormatoNumeroView)
    fnMonto = pnMonto
    fsGlosa = psGlosa
    fsDetalle = psDetalle
End Sub

