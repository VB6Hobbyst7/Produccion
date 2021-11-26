VERSION 5.00
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredAguaSaneamientoDet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Sub Destino "
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   ForeColor       =   &H80000008&
   Icon            =   "frmCredAguaSaneamientoDet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameBeneficiados 
      Caption         =   "Beneficiados"
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
      DragIcon        =   "frmCredAguaSaneamientoDet.frx":030A
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
      Caption         =   "Monto Solicitado"
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
Attribute VB_Name = "frmCredAguaSaneamientoDet"
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
Private Enum eAguaSaneamiento
    AguaSaneamientoRegistrar = 1
    AguaSaneamientoModificar = 2
End Enum

Dim fnInicio As eAguaSaneamiento
Dim fvAguaSaneamiento As TAguaSaneamiento
Dim fvListaAguaSaneamiento() As TAguaSaneamiento
Dim fnIndexValida As Integer
Dim fvListaValida() As TAguaSaneamiento

Dim fbAceptar As Boolean
Dim fnTpoProducto As Integer
Dim fnFormTamanioIni As Double, fnFormTamanioActiva As Double
Dim sCtaCod As String
Dim nSuma As Double
Dim nVarTemp As Double
Dim nPresEnter As Integer
Dim nCentinela As Integer 'EAAS20191401 SEGUN 018-GM-DI_CMACM
Dim nMontoSubDestinoEditar As Double 'EAAS20191401 SEGUN 018-GM-DI_CMACM
Private Sub Form_Load()
    fnFormTamanioIni = 4650
    fnFormTamanioActiva = 6300
    
    fbAceptar = False
    CargaControles
    
    If fnInicio = AguaSaneamientoModificar Then

        
        txtDestinoPrincipal.Text = fvAguaSaneamiento.sDestinoDescripcion & "                                       " & fvAguaSaneamiento.nDestinoCod
        cmbSubDestino.ListIndex = IndiceListaCombo(cmbSubDestino, fvAguaSaneamiento.nSubDestinoCod)
        txtMontoP.Text = dMontoSolicitado 'EAAS20191401 SEGUN 018-GM-DI_CMACM SE COMENTO- nSuma
        txtMontoS.Text = fvAguaSaneamiento.nMontoS
        txtMontoP.Text = Format(txtMontoP.Text, "#,##0.00")
        txtMontoS.Text = Format(txtMontoS.Text, "#,##0.00")
        spnBeneficiados.valor = fvAguaSaneamiento.nBeneficia

    End If
End Sub

Public Sub Inicio(ByRef pvListaAguaSaneamiento() As TAguaSaneamiento, ByVal pDestCred As String, ByRef pMontoSol As Double)
    fvListaAguaSaneamiento = pvListaAguaSaneamiento
    txtDestinoPrincipal = pDestCred
    txtMontoP.ForeColor = vbBlue
    txtMontoS.ForeColor = vbBlue
    txtMontoP.ForeColor = pMontoSol
    dMontoSolicitado = pMontoSol
    CargaSubDestino
    Show 1
    pvListaAguaSaneamiento = fvListaAguaSaneamiento
End Sub

Private Sub CargaSubDestino()
Dim oCred As COMDCredito.DCOMCredito
Dim ssql As String
Dim RTemp As ADODB.Recordset
    On Error GoTo ERRORCargaSubDestino
    Set oCred = New COMDCredito.DCOMCredito
    Set RTemp = oCred.RecuperaCargaSubDestinoAguaSaneamiento()
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
    'INICIO EAAS20191401 SEGUN 018-GM-DI_CMACM
    If (nMontoSubDestinoEditar = 0) Then
    dMontoSolicitado = dMontoSolicitado - txtMontoS.Text
    Else
    dMontoSolicitado = dMontoSolicitado + (nMontoSubDestinoEditar - txtMontoS.Text)
    End If
    'FIN EAAS20191401 SEGUN 018-GM-DI_CMACM
    fvAguaSaneamiento.nDestinoCod = Right(txtDestinoPrincipal.Text, 2)
    fvAguaSaneamiento.sDestinoDescripcion = txtDestinoPrincipal
    fvAguaSaneamiento.nSubDestinoCod = Right(cmbSubDestino.Text, 2)
    fvAguaSaneamiento.sSubDestinoDescripcion = cmbSubDestino.Text
    fvAguaSaneamiento.nMontoP = txtMontoP.Text
    fvAguaSaneamiento.nMontoS = Format(txtMontoS.Text, "#,##0.00")
    fvAguaSaneamiento.nBeneficia = CInt(spnBeneficiados.valor)
    fvAguaSaneamiento.sNroCredito = sCtaCod
    For i = 1 To UBound(fvListaValida)
    fvListaValida(i).nBeneficia = CInt(spnBeneficiados.valor)
    Next
    fvAguaSaneamiento.nMontoS = Format(fvAguaSaneamiento.nMontoS, "#0.00")
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
            'txtMontoP = Format(dMontoSolicitado - nVarTemp - IIf(txtMontoS.Text = "", 0, txtMontoS.Text), "#,##0.00")//COMENTO 'EAAS20191401 SEGUN 018-GM-DI_CMACM
            'INICIO EAAS20191401 SEGUN 018-GM-DI_CMACM
            If (nMontoSubDestinoEditar = 0) Then
            txtMontoP = Format(dMontoSolicitado - nMontoSubDestinoEditar - IIf(txtMontoS.Text = "", 0, txtMontoS.Text), "#,##0.00")
            Else
            txtMontoP = Format(dMontoSolicitado + (nMontoSubDestinoEditar - IIf(txtMontoS.Text = "", 0, txtMontoS.Text)), "#,##0.00")
            End If
            'FIN EAAS20191401 SEGUN 018-GM-DI_CMACM
            txtMontoS.Text = Format(txtMontoS.Text, "#,##0.00")
            If (CDbl(txtMontoP.Text) < 0) Then
                MsgBox "Monto Mayor a lo solitado", vbInformation, "Aviso"
                txtMontoP = Format(nMontoPTemp, "#,##0.00")
                txtMontoS.SetFocus
                nPresEnter = 0
                Exit Sub
                'EAAS20191401 SEGUN 018-GM-DI_CMACM nCentinela
            ElseIf (CDbl(txtMontoP.Text) <= 0 And Right(txtDestinoPrincipal.Text, 2) <> 26 And nCentinela = 0) Then
                MsgBox "El Monto Sub Destino no puede ser igual al Monto solicitado.", vbInformation, "Aviso"
                txtMontoS.SetFocus
                nPresEnter = 0
                Exit Sub
            Else
                If (spnBeneficiados.Enabled = True) Then
                spnBeneficiados.SetFocus
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
Public Function Registrar(ByRef pvAguaSaneamiento As TAguaSaneamiento, ByRef pvListaValida() As TAguaSaneamiento, ByVal pnTpoProducto As Integer, ByVal pDestCred As String, ByRef pMontoSol As Double, ByVal psCtaCod As String, Optional ByRef pnbeneficiarios As Integer = 0, Optional ByRef pnCentinela As Integer) As Boolean
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
    fvAguaSaneamiento = pvAguaSaneamiento
    fvListaValida = pvListaValida
    fnTpoProducto = pnTpoProducto
    sCtaCod = psCtaCod
    nCentinela = pnCentinela 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    Show 1
    Registrar = fbAceptar
    pMontoSol = dMontoSolicitado
    pvAguaSaneamiento = fvAguaSaneamiento
End Function
Public Function Modificar(ByRef pvAguaSaneamiento As TAguaSaneamiento, ByVal pnIndexValida As Integer, ByRef pvListaValida() As TAguaSaneamiento, ByVal pnTpoProducto As Integer, ByRef pdMontoSol As Double, ByVal psCtaCod As String, ByVal pnSuma As Double, ByRef pnMontoSubDestinoEditar As Double, Optional ByRef pnCentinela As Integer) As Boolean 'EAAS20190410 SEGUN 018-GM-DI_CMACM pnCentinela
    fnInicio = AguaSaneamientoModificar
    dMontoSolicitado = pdMontoSol
    fvAguaSaneamiento = pvAguaSaneamiento
    fnIndexValida = pnIndexValida
    fvListaValida = pvListaValida
    fnTpoProducto = pnTpoProducto
    sCtaCod = psCtaCod
    nCentinela = pnCentinela 'EAAS20190410 SEGUN 018-GM-DI_CMACM
    nMontoSubDestinoEditar = pnMontoSubDestinoEditar 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    nSuma = pnSuma
    nVarTemp = nSuma - CDbl(txtMontoS.Text)
    Show 1
    Modificar = fbAceptar
    pvListaValida = fvListaValida
    pvAguaSaneamiento = fvAguaSaneamiento
    pdMontoSol = dMontoSolicitado

End Function

Private Sub CargaControles()
Dim oCred As COMDCredito.DCOMCredito
Dim ssql As String
Dim RTemp As ADODB.Recordset
    On Error GoTo ErrCargaControles
    Set oCred = New COMDCredito.DCOMCredito
    Set RTemp = oCred.RecuperaCargaSubDestinoAguaSaneamiento()
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
