VERSION 5.00
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredVerdeDet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Sub Destino "
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   ForeColor       =   &H80000008&
   Icon            =   "frmCredVerdeDet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameBeneficiados 
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   8415
      Begin Spinner.uSpinner spnBeneficiados 
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   5040
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton btnSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   6720
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lbBeneficiados 
         Caption         =   "N° de Beneficiarios"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.TextBox txtMontoS 
      Height          =   405
      Left            =   6600
      MaxLength       =   15
      TabIndex        =   7
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtMontoP 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.ComboBox cmbSubDestino 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox txtDestinoPrincipal 
      DragIcon        =   "frmCredVerdeDet.frx":030A
      Enabled         =   0   'False
      Height          =   405
      Left            =   1680
      TabIndex        =   0
      Top             =   210
      Width           =   2775
   End
   Begin VB.Label lblMontoMinimo 
      Caption         =   "* Monto mínimo Sub Destino S/300"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   13
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label lbMontoSub 
      Caption         =   "Monto Sub Destino"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lbMontoPrincipal 
      Caption         =   "Monto Disponible"
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lbSubDestino 
      Caption         =   "Sub Destino"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lbDestinoPrincipal 
      Caption         =   "Destino Principal"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmCredVerdeDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************
'** Nombre : frmCredAguaSaneamiento
'** Descripción : Paraincluir el destino de agua y saneamiento para discrimar que porcentaje
'** del monto solicitado apunta hacia un destino determinado creado segun TI-ERS052-2018
'** Creación : EAAS, 20180727 04:30:00 PM
'************************************************************************************************
Option Explicit
Dim dMontoSolicitado As Double
Private Enum eCreditoVerde
    CreditoVerdeRegistrar = 1
    CreditoVerdeModificar = 2
End Enum

Dim fnInicio As eCreditoVerde
Dim fvCreditoVerde As TCreditoVerde
Dim fvListaCreditoVerde() As TCreditoVerde
Dim fnIndexValida As Integer
Dim fvListaValida() As TCreditoVerde

Dim fbAceptar As Boolean
Dim fnTpoProducto As Integer
Dim fnFormTamanioIni As Double, fnFormTamanioActiva As Double
Dim sCtaCod As String
Dim nSuma As Double
Dim nVarTemp As Double
Dim nPresEnter As Integer
Dim nMontoSubDestinoEditar As Double 'EAAS20191401 SEGUN 018-GM-DI_CMACM
Private Sub Form_Load()
    fnFormTamanioIni = 4650
    fnFormTamanioActiva = 6300
    
    fbAceptar = False
    CargaControles
    
    If fnInicio = CreditoVerdeModificar Then

        
        txtDestinoPrincipal.Text = fvCreditoVerde.sDestinoDescripcion & "                                       " & fvCreditoVerde.nDestinoCod
        cmbSubDestino.ListIndex = IndiceListaCombo(cmbSubDestino, fvCreditoVerde.nSubDestinoCod)
        txtMontoP.Text = dMontoSolicitado 'EAAS20191401 SEGUN 018-GM-DI_CMACM - nSuma
        txtMontoS.Text = fvCreditoVerde.nMontoS
        txtMontoP.Text = Format(txtMontoP.Text, "#,##0.00")
        txtMontoS.Text = Format(txtMontoS.Text, "#,##0.00")
        spnBeneficiados.valor = fvCreditoVerde.nBeneficia

    End If
End Sub

Public Sub Inicio(ByRef pvCreditoVerde() As TCreditoVerde, ByVal pDestCred As String, ByRef pMontoSol As Double)
    fvListaCreditoVerde = pvCreditoVerde
    txtDestinoPrincipal = pDestCred
    txtMontoP.ForeColor = vbBlue
    txtMontoS.ForeColor = vbBlue
    txtMontoP.ForeColor = pMontoSol
    dMontoSolicitado = pMontoSol
    CargaSubDestino
    Show 1
    pvCreditoVerde = fvListaCreditoVerde
End Sub

Private Sub CargaSubDestino()
Dim oCred As COMDCredito.DCOMCredito
Dim ssql As String
Dim RTemp As ADODB.Recordset
    On Error GoTo ERRORCargaSubDestino
    Set oCred = New COMDCredito.DCOMCredito
    Set RTemp = oCred.RecuperaCargaSubDestinoCreditoVerde()
    Set oCred = Nothing
    cmbSubDestino.Clear
    Do While Not RTemp.EOF
        cmbSubDestino.AddItem RTemp!cConsDescripcion & Space(250) & RTemp!nConsValor
        RTemp.MoveNext
    Loop
    RTemp.Close
    Set RTemp = Nothing
    Call CambiaTamañoCombo(cmbSubDestino, 250)
    Exit Sub
    
ERRORCargaSubDestino:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub
Private Sub btnSalir_Click()
    Unload Me
End Sub

Private Sub CmdAceptar_Click()
    On Error GoTo ErrAceptar
    Dim bCreditoAnt As Boolean
    Dim i As Integer
    If Len(Trim(txtMontoS.Text)) = 0 Then
         txtMontoP.Text = "0.00"
    End If
    txtMontoS.Text = Format(txtMontoS.Text, "#0.00")
    Call TxtMontoS_KeyPress(13)
    
    If (nPresEnter = 0) Then
        Exit Sub
    End If
    
    If (cmbSubDestino.Text = "") Then
    MsgBox "Seleccione Sub Destino", vbInformation, "Aviso"
    Exit Sub
    End If
    spnBeneficiados.valor = 1
    If (CInt(spnBeneficiados.valor) = 0) Then
    MsgBox "Digite al menos un beneficiario", vbInformation, "Aviso"
    Exit Sub
    End If
    'INICIO EAAS20190410 SEGUN 018-GM-DI_CMACM
    If (txtMontoS.Text = "") Then
    MsgBox "Ingrese Monto Sub Destino", vbInformation, "Aviso"
    Exit Sub
    End If
    'FIN EAAS20190410 SEGUN 018-GM-DI_CMACM
    If (CInt(txtMontoS.Text) < 300) Then
    MsgBox "El monto mínimo del Sub Destino debe ser S/ 300.00", vbInformation, "Aviso"
    Exit Sub
    End If
   
    If (nMontoSubDestinoEditar = 0) Then
    dMontoSolicitado = dMontoSolicitado - txtMontoS.Text
    Else
    dMontoSolicitado = dMontoSolicitado + (nMontoSubDestinoEditar - txtMontoS.Text)
    End If
    fvCreditoVerde.nDestinoCod = Right(txtDestinoPrincipal.Text, 2)
    fvCreditoVerde.sDestinoDescripcion = txtDestinoPrincipal
    fvCreditoVerde.nSubDestinoCod = Right(cmbSubDestino.Text, 2)
    fvCreditoVerde.sSubDestinoDescripcion = cmbSubDestino.Text
    fvCreditoVerde.nMontoP = txtMontoP.Text
    fvCreditoVerde.nMontoS = Format(txtMontoS.Text, "#,##0.00")
    fvCreditoVerde.nBeneficia = CInt(spnBeneficiados.valor)
    fvCreditoVerde.sNroCredito = sCtaCod
    For i = 1 To UBound(fvListaValida)
    fvListaValida(i).nBeneficia = CInt(spnBeneficiados.valor)
    Next
    fvCreditoVerde.nMontoS = Format(fvCreditoVerde.nMontoS, "#0.00")
    fbAceptar = True
    Unload Me
    Exit Sub
ErrAceptar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub txtMontoP_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMontoP, KeyAscii)
    Dim dMontoSubTipo As Double
    If KeyAscii = 13 Then
        If (txtMontoP.Text <> "") Then
            If (CDbl(txtMontoP.Text) > dMontoSolicitado) Then
                MsgBox "Monto Mayor a lo solitado", vbInformation, "Aviso"
                txtMontoS.Text = 0
                txtMontoP.Text = 0
                txtMontoP_LostFocus
                'txtMontoS_LostFocus
                txtMontoP.SetFocus
            ElseIf (CDbl(txtMontoP.Text) <= 0) Then
                MsgBox "Monto tiene que ser mayor a Cero", vbInformation, "Aviso"
                txtMontoS.Text = 0
                txtMontoP.Text = 0
                txtMontoP_LostFocus
                'txtMontoS_LostFocus
                txtMontoP.SetFocus
            Else
                dMontoSubTipo = dMontoSolicitado - CDbl(txtMontoP.Text)
                txtMontoS.Text = CStr(dMontoSubTipo)
                spnBeneficiados.SetFocus
            End If
        End If
    End If
End Sub

Private Sub TxtMontoS_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMontoS, KeyAscii)
    Dim nMontoPTemp As Double
    If KeyAscii = 13 Then
        If (txtMontoP.Text <> "") Then
            nMontoPTemp = CDbl(txtMontoP.Text)
            nMontoPTemp = Format(nMontoPTemp, "#,##0.00")
            If (nMontoSubDestinoEditar = 0) Then
            txtMontoP = Format(dMontoSolicitado - nMontoSubDestinoEditar - IIf(txtMontoS.Text = "", 0, txtMontoS.Text), "#,##0.00")
            Else
            txtMontoP = Format(dMontoSolicitado + (nMontoSubDestinoEditar - IIf(txtMontoS.Text = "", 0, txtMontoS.Text)), "#,##0.00")
            End If
            txtMontoS.Text = Format(txtMontoS.Text, "#,##0.00")
            
            If (CDbl(txtMontoP.Text) < 0) Then
                MsgBox "Monto Mayor a lo solitado", vbInformation, "Aviso"
                txtMontoP = Format(nMontoPTemp, "#,##0.00")
                txtMontoS.SetFocus
                nPresEnter = 0
                Exit Sub
'            ElseIf (CDbl(txtMontoP.Text) <= 0 And Right(txtDestinoPrincipal.Text, 2) <> 26) Then
'                MsgBox "El Monto Sub Destino no puede ser igual al Monto solicitado.", vbInformation, "Aviso"
'                txtMontoS.SetFocus
'                nPresEnter = 0
'                Exit Sub
            Else
                If (spnBeneficiados.Enabled = True) Then
                'spnBeneficiados.SetFocus
                End If
            End If
        End If
        
    End If
    
    nPresEnter = 1
End Sub

Private Sub txtMontoP_LostFocus()
    If Len(Trim(txtMontoP.Text)) = 0 Then
         txtMontoP.Text = "0.00"
    End If
    txtMontoP.Text = Format(txtMontoP.Text, "#0.00")
    'txtMontoS_LostFocus
End Sub
Public Function Registrar(ByRef pvCreditoVerde As TCreditoVerde, ByRef pvListaValida() As TCreditoVerde, ByVal pnTpoProducto As Integer, ByVal pDestCred As String, ByRef pMontoSol As Double, ByVal psCtaCod As String, Optional ByRef pnbeneficiarios As Integer = 0) As Boolean
    txtMontoP.Text = Format(txtMontoP.Text, "#0.00")
    txtDestinoPrincipal = pDestCred
    txtMontoP.ForeColor = vbBlue
    txtMontoS.ForeColor = vbBlue
    txtMontoP = pMontoSol
    txtMontoP.Text = Format(txtMontoP.Text, "#,##0.00")
    If (pnbeneficiarios <> 0) Then
    spnBeneficiados.valor = pnbeneficiarios
    spnBeneficiados.Enabled = False
    End If
    dMontoSolicitado = pMontoSol
    CargaSubDestino
    fvCreditoVerde = pvCreditoVerde
    fvListaValida = pvListaValida
    fnTpoProducto = pnTpoProducto
    sCtaCod = psCtaCod
    Show 1
    Registrar = fbAceptar
    pMontoSol = dMontoSolicitado
    pvCreditoVerde = fvCreditoVerde
End Function
Public Function Modificar(ByRef pvCreditoVerde As TCreditoVerde, ByVal pnIndexValida As Integer, ByRef pvListaValida() As TCreditoVerde, ByVal pnTpoProducto As Integer, ByRef pdMontoSol As Double, ByVal psCtaCod As String, ByVal pnSuma As Double, ByRef pnMontoSubDestinoEditar As Double) As Boolean
    fnInicio = CreditoVerdeModificar
    dMontoSolicitado = pdMontoSol
    fvCreditoVerde = pvCreditoVerde
    fnIndexValida = pnIndexValida
    fvListaValida = pvListaValida
    fnTpoProducto = pnTpoProducto
    sCtaCod = psCtaCod
    nMontoSubDestinoEditar = pnMontoSubDestinoEditar 'EAAS20191401 SEGUN 018-GM-DI_CMACM
'    nSuma = pnSuma
'    nVarTemp = nSuma - CDbl(txtMontoS.Text)
    Show 1
    Modificar = fbAceptar
    pvListaValida = fvListaValida
    pvCreditoVerde = fvCreditoVerde
    pdMontoSol = dMontoSolicitado

End Function

Private Sub CargaControles()
Dim oCred As COMDCredito.DCOMCredito
Dim ssql As String
Dim RTemp As ADODB.Recordset
    On Error GoTo ErrCargaControles
    Set oCred = New COMDCredito.DCOMCredito
    Set RTemp = oCred.RecuperaCargaSubDestinoCreditoVerde()
    Set oCred = Nothing
    cmbSubDestino.Clear
    Do While Not RTemp.EOF
        cmbSubDestino.AddItem RTemp!cConsDescripcion & Space(250) & RTemp!nConsValor
        RTemp.MoveNext
    Loop
    RTemp.Close
    Set RTemp = Nothing
    Call CambiaTamañoCombo(cmbSubDestino, 250)
    Exit Sub
ErrCargaControles:
    MsgBox Err.Description, vbCritical, "Aviso"
    Screen.MousePointer = 0
End Sub
