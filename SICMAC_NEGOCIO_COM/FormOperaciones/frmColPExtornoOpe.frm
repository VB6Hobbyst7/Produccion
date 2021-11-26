VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColPExtornoOpe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio - Extorno de Operaciones"
   ClientHeight    =   5610
   ClientLeft      =   735
   ClientTop       =   1950
   ClientWidth     =   9660
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmColPExtornoOpe.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmMotExtorno 
      Caption         =   "Motivos del Extorno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2700
      Left            =   3480
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   2845
      Begin VB.CommandButton cmdExtContinuar 
         Caption         =   "&Continuar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   855
         TabIndex        =   14
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtDetExtorno 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   240
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox cmbMotivos 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmColPExtornoOpe.frx":030A
         Left            =   240
         List            =   "frmColPExtornoOpe.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalles del Extorno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1900
      End
      Begin VB.Label lblExtCmb 
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CheckBox chkOpeCMACLlamada 
      Caption         =   "Operaciones LLamada CMAC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7680
      TabIndex        =   6
      Top             =   0
      Width           =   1695
      Begin VB.CommandButton cmdExtorno 
         Caption         =   "&Extornar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1245
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1245
      End
   End
   Begin VB.Frame fraBuscar 
      Caption         =   "Buscar Por"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   180
      TabIndex        =   3
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton opt 
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   1
         Top             =   720
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.OptionButton opt 
         Caption         =   "Nro Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   0
         Top             =   360
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   240
      Width           =   1005
   End
   Begin MSComctlLib.ListView lstExtorno 
      Height          =   3540
      Left            =   180
      TabIndex        =   4
      Top             =   1320
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   6244
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   17
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColPExtornoOpe.frx":030E
            Key             =   "Cuenta"
         EndProperty
      EndProperty
   End
   Begin SICMACT.ActXCodCta AXCodCta 
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      Texto           =   "Crédito"
      EnabledProd     =   -1  'True
   End
   Begin VB.Label lblMensaje 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Por favor, sea muy cuidadoso(a) al utilizar los EXTORNOS. No hay forma de volver a realizar el proceso del Extorno."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   3120
      TabIndex        =   5
      Top             =   4920
      Width           =   6360
   End
End
Attribute VB_Name = "frmColPExtornoOpe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'EXTORNO DE OPERACIONES DE CONTRATO PIGNORATICIO
'Archivo:  frmColPExtornoOpe.frm
'LAYG   :  25/07/2001.
'ICA  : 25/09/2004 - layg
'Resumen:  Nos permite realizar los extornos de Prendario
Option Explicit
Dim vNroContrato As String

Dim fsFechaTransac As String
Dim fnOperacion As COMDConstantes.ColocPOperaciones
Dim fsListaOpe As String

Dim fsRemateCadaAgencia As String
Dim fsAdjudicaCadaAgencia As String


Public Sub Inicio(ByVal pnOperacion As COMDConstantes.ColocPOperaciones, ByVal PsOperacion As String)

fnOperacion = pnOperacion

Me.Caption = "Pignoraticio - Extornos - " & PsOperacion
fsListaOpe = "('')"
Select Case fnOperacion
    Case gColPOpeExtDesemb, gColPOpeExtDesembAmpliado
        '*** PEAC 20090608 - se agregó las operaciones ligadas a desembolso en cuenta

        'gColPOpeDesembolsoAboCta = 120202
        'gColPOpeDesembolsoAboCtaOtraAge = 120203
        'gColPOpeDesembolsoAboCtaNva = 120204

        'fsListaOpe = "('" & gColPOpeDesembolsoEFE & "') "
        fsListaOpe = "('" & gColPOpeDesembolsoEFE & "','" & gColPOpeDesembolsoAboCta & "','" & gColPOpeDesembolsoAboCtaNva & "','" & gColPOpeDesembolsoAboCtaOtraAge & "') "
    Case gColPOpeExtRenov
        fsListaOpe = "('" & gColPOpeRenovNorEFE & "','" & gColPOpeRenovMorEFE & "','" & gColPOpeRenovNorCHQ & "','" & gColPOpeRenovMorCHQ & "','" & gColPOpeRenovNorVoucher & "','" & gColPOpeRenovMorVoucher _
            & "','" & gColPOpeRenovNorCargoCta & "','" & gColPOpeRenovMorCargoCta & "')"
        'CTI4 ERS0112020 Add:gColPOpeRenovNorVoucher, gColPOpeRenovMorVoucher,gColPOpeRenovNorCargoCta,gColPOpeRenovMorCargoCta
    Case gColPOpeExtCance
        fsListaOpe = "('" & gColPOpeCancelNorEFE & "','" & gColPOpeCancelMorEFE & "','" & gColPOpeCancelNorCHQ & "','" & gColPOpeCancelMorCHQ & "','" & gColPOpeCancelNorVoucher & "','" & gColPOpeCancelMorVoucher & "','" & gColPOpeCancelNorCargoCta & "','" & gColPOpeCancelMorCargoCta & "')"
        'CTI4 ERS0112020 Add:gColPOpeCancelNorVoucher, gColPOpeCancelMorVoucher, gColPOpeCancelNorCargoCta, gColPOpeCancelMorCargoCta
    '*** PEAC 20161212
    Case gColPOpeExtPagoParc
        fsListaOpe = "('" & gColPOpePagoParcialNorEfectivo & "','" & gColPOpePagoParcialNorCheque & "','" & gColPOpePagoParcialNorVoucher & "','" & gColPOpePagoParcialNorCargoCta & "')"
        'CTI4 ERS0112020 Add: gColPOpePagoParcialNorVoucher, gColPOpePagoParcialNorCargoCta
    Case gColPOpeExtDevJoyas
        fsListaOpe = "('" & gColPOpeDevJoyas & "')"
    Case gColPOpeExtCustodDifer
        fsListaOpe = "('" & gColPOpeCobCusDiferida & "')"
    Case gColPOpeExtDupli
        fsListaOpe = "('" & gColPOpeImpDuplicado & "')"
    Case "129700"
        fsListaOpe = "('" & gColPOpeAmortNorEFE & "','" & gColPOpeAmortNorCHQ & "')"
    Case "129701" ' Venta Remate
        fsListaOpe = "('" & gColPOpeVtaRemate & "')"
    Case "129702" ' Pago Sobrante
         fsListaOpe = "('" & gColPOpePagSobrante & "')"
    Case "129703" ' Venta Adjudicado
         fsListaOpe = "('" & gColPOpeVtaSubasta & "')"
    Case "129704" ' Recuperacion Adjudicado
         fsListaOpe = "('122900','" & gColPOpeRecuperaContratoAdjVoucher & "','" & gColPOpeRecuperaContratoAdjCargoCta & "')"
         'CTI4 ERS0112020 Add:gColPOpeRecuperaContratoAdjVoucher, gColPOpeRecuperaContratoAdjCargoCta
         
    '*** PEAC 20090316 - Pago sobrante adjudicado
    Case "129705"
         fsListaOpe = "('" & gColPOpePagSobraAdjudicado & "')"
         
    Case "129801"
         fsListaOpe = "('" & gColPOpeRenovNorEnOtCjEFE & "','" & gColPOpeRenovMorEnOtCjEFE & "')"
    Case "129802"
         fsListaOpe = "('" & gColPOpeCanceNorEnOtCjEFE & "','" & gColPOpeCanceMorEnOtCjEFE & "')"
    Case "129803"
         fsListaOpe = "('" & gColPOpeAmortNorEnOtCjEFE & "')"
    
'    Case gColPOpeExtRenovEOCmac
'        fsListaOpe = "('" & gColPOpeRenovNorEnOtCjEFE & "','" & gColPOpeRenovMorEnOtCjEFE & "','" & gColPOpeRenovNorEnOtCjCHQ & "','" & gColPOpeRenovMorEnOtCjCHQ & "')"
'    Case gColPOpeExtCanceEOCmac
'        fsListaOpe = "('" & gColPOpeCanceNorEnOtCjEFE & "','" & gColPOpeCanceMorEnOtCjEFE & "','" & gColPOpeCanceNorEnOtCjCHQ & "','" & gColPOpeCanceMorEnOtCjCHQ & "' )"
'    Case "129803"
'        fsListaOpe = "('" & gColPOpeAmortNorEnOtCjEFE & "','" & gColPOpeAmortNorEnOtCjCHQ & "')"
'    Case "129803"
'        fsListaOpe = "('" & gColPOpeAmortNorEnOtCjEFE & "','" & gColPOpeAmortNorEnOtCjCHQ & "')"
'
End Select
    
Me.Show 1
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdBuscar.SetFocus
End Sub

Private Sub cmdBuscar_Click()

Dim lrBusca As ADODB.Recordset
Dim loValContrato As COMNColoCPig.NCOMColPValida
Dim lsmensaje As String
'On Error GoTo ControlError
    'Valida Contrato
    Limpiar
    Set lrBusca = New ADODB.Recordset
    Set loValContrato = New COMNColoCPig.NCOMColPValida
        If Me.chkOpeCMACLlamada.value = 1 Then ' Operaciones LLamada CMAC
            Set lrBusca = loValContrato.nBuscaOperacionesCredPigParaExtornoLLamadaCMAC(fsFechaTransac, , , lsmensaje)
        Else
            If Me.opt(0).value = True Then ' Busca por Codigo
                Set lrBusca = loValContrato.nBuscaOperacionesCredPigParaExtorno(fsFechaTransac, fsListaOpe, Me.AXCodCta.NroCuenta, , lsmensaje)
            Else
                Set lrBusca = loValContrato.nBuscaOperacionesCredPigParaExtorno(fsFechaTransac, fsListaOpe, , , lsmensaje)
            End If
        End If
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Exit Sub
        End If
    Set loValContrato = Nothing
    
    If lrBusca Is Nothing Then ' Hubo un Error
        Set lrBusca = Nothing
        Exit Sub
    End If
    If lrBusca.BOF And lrBusca.EOF Then
        MsgBox "No Existen Operaciones para EXTORNAR", vbInformation, "Aviso"
        Exit Sub
    End If
    
    lstExtorno.ListItems.Clear
    
    Call LLenaLista(lrBusca)
    
    Set lrBusca = Nothing
        
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & err.Number & " " & err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

'***CTI3 (ferimoro) modificado
Private Sub cmdExtContinuar_Click()
'On Error GoTo ControlError

Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabarExt As COMNColoCPig.NCOMColPContrato

Dim lsMovNro As String
Dim lsFechaHoraGrab As String

Dim lsNroContrato As String
Dim lsOperacion As String, lsOperacionDesc As String
Dim lnMovNroAExt As Long
Dim lnSaldo As Currency
Dim lnMonto As Currency
Dim lnITFOpeVoucher As Currency 'CTI4 ERS0112020
Dim Fecha As String
Dim lsCliente As String, lsCMACOpe As String
Dim loImprime As COMNColoCPig.NCOMColPImpre
Dim loPrevio As previo.clsprevio
Dim lsCadImprimir As String
Dim lsImpCargoCta As String 'CTI4 ERS0112020
Dim lsmensaje As String


'*** PEAC 20081002
Dim lbResultadoVisto As Boolean
Dim sPersVistoCod  As String
Dim sPersVistoCom As String
Dim loVistoElectronico As frmVistoElectronico
Set loVistoElectronico = New frmVistoElectronico

Dim lsNroProceso As String
Dim lsCtaAhorroRemate As String
Dim lsCtaAhorroAdjudica As String

Dim loConec As COMConecta.DCOMConecta
Dim lrdatosrem As ADODB.Recordset
Dim lrdatosAdj As ADODB.Recordset
Dim lsSQL As String
Dim lnITFAho As Currency, lsCtaAhoExt As String, lbProcedeExtAho As Boolean
Dim lsClienteAhoExt As String  'CTI4 ERS0112020
Dim lnMontoAhoExt As Currency 'CTI4 ERS0112020

'If lstExtorno.ListItems.count = 0 Then
'    cmdExtorno.Enabled = False
'    Exit Sub
'End If

If cmbMotivos.ListIndex = -1 Or Len(txtDetExtorno.Text) <= 0 Then
    MsgBox "Debe ingresar el motivo y/o detalle del Extorno", vbInformation, "Aviso"
    Exit Sub
End If

'***CTI3 (ferimoro)    02102018
frmMotExtorno.Visible = False
Dim DatosExtorna(1) As String
DatosExtorna(0) = cmbMotivos.Text
DatosExtorna(1) = txtDetExtorno.Text
'***********************************
'*** PEAC 20081001 - visto electronico ******************************************************
'*** en estos extornos de operaciones pedirá visto electrónico

    ' *** RIRO SEGUN TI-ERS108-2013 ***
           
        Dim nMovNroOperacion As Long
        nMovNroOperacion = 0
        If Not lstExtorno.SelectedItem Is Nothing Then
        nMovNroOperacion = CDbl(Val(lstExtorno.SelectedItem.ListSubItems(5)))
        End If
           
    ' *** FIN RIRO ***

           Select Case fnOperacion
                Case "129100", "129400", "129700", "129702", "129703", "129705"
            
                    lbResultadoVisto = loVistoElectronico.Inicio(3, fnOperacion, , , nMovNroOperacion) 'RIRO SEGUN TI-ERS108-2013/ Se agrego parametro nMovNroOperacion
                    If Not lbResultadoVisto Then
                    
                        '******CTI3 (ferimoro) 27092018
                        frmMotExtorno.Visible = False
                        Me.cmbMotivos.ListIndex = -1
                        Me.txtDetExtorno.Text = ""
                        fraBuscar.Enabled = True
                        AXCodCta.Enabled = True
                        cmdBuscar.Enabled = True
                        cmdExtorno.Enabled = False
                        '******************************
                    
                        Exit Sub
                    End If
           End Select

'*** FIN PEAC ************************************************************

If lstExtorno.SelectedItem.SubItems(2) <> "1" Then
    MsgBox " Debe Extornar el último movimiento del Contrato ", vbInformation, " Aviso "
    Exit Sub
Else
    If MsgBox(" Esta Ud seguro de Extornar dicha Operación ? ", vbQuestion + vbYesNo + vbDefaultButton2, " Aviso ") = vbNo Then
        '******CTI3 (ferimoro) 27092018
        frmMotExtorno.Visible = False
        Me.cmbMotivos.ListIndex = -1
        Me.txtDetExtorno.Text = ""
        fraBuscar.Enabled = True
        AXCodCta.Enabled = True
        cmdBuscar.Enabled = True
        cmdExtorno.Enabled = False
        '******************************
        Exit Sub
    Else
        MsgBox " Prepare la impresora para imprimir " & vbCr & _
        " el recibo del Extorno", vbInformation, " Aviso "
    End If
End If


'*** Obtiene Datos de Operacion
lsNroContrato = Trim(lstExtorno.SelectedItem)
lsOperacion = Right(lstExtorno.SelectedItem.ListSubItems(1), 6)
lnMovNroAExt = CCur(lstExtorno.SelectedItem.ListSubItems(5))
'lnSaldo = CCur(lstExtorno.SelectedItem.ListSubItems(2))
lnMonto = CCur(lstExtorno.SelectedItem.ListSubItems(3))
Fecha = lstExtorno.SelectedItem.ListSubItems(4)
lsCliente = lstExtorno.SelectedItem.ListSubItems(7)
lsCMACOpe = lstExtorno.SelectedItem.ListSubItems(8)
lnITFOpeVoucher = lstExtorno.SelectedItem.ListSubItems(9) 'CTI4 ERS0112020
'*** Genera el Mov Nro
Set loContFunct = New COMNContabilidad.NCOMContFunciones
    lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
Set loContFunct = Nothing

lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
    
Set loGrabarExt = New COMNColoCPig.NCOMColPContrato
    
    Select Case lsOperacion
        '** Extornar un DESEMBOLSO PIG

        '*** PEAC 20090609
'       gColPOpeDesembolsoAboCta = 120202
'       gColPOpeDesembolsoAboCtaOtraAge = 120203
'       gColPOpeDesembolsoAboCtaNva = 120204
        
        'Case gColPOpeDesembolsoEFE
        Case gColPOpeDesembolsoEFE, gColPOpeDesembolsoAboCta, gColPOpeDesembolsoAboCtaOtraAge, gColPOpeDesembolsoAboCtaNva
        
           'RECO20140208 ERS002******************************************************************************
                'Call loGrabarExt.nExtornoDesembolsoCredPig(lsNroContrato, lsFechaHoraGrab, _
                lsMovNro, lnMovNroAExt, lnMonto, False, lbProcedeExtAho, lnITFAho, lsCtaAhoExt)
                'lsOperacionDesc = "DESEMBOLSO"
            If fnOperacion = gColPOpeExtDesembAmpliado Then
                Call loGrabarExt.nExtornoDesembolsoCredPigAmpliado(lsNroContrato, lsFechaHoraGrab, _
                lsMovNro, lnMovNroAExt, lnMonto, False, lbProcedeExtAho, lnITFAho, lsCtaAhoExt, , lsmensaje, DatosExtorna)
                lsOperacionDesc = "DESEMBOLSO X AMPLIACIÓN"
            Else
                Call loGrabarExt.nExtornoDesembolsoCredPig(lsNroContrato, lsFechaHoraGrab, _
                lsMovNro, lnMovNroAExt, lnMonto, False, lbProcedeExtAho, lnITFAho, lsCtaAhoExt, , DatosExtorna)
                lsOperacionDesc = "DESEMBOLSO"
            End If
            'RECO FIN****
        
        '** Extornar un DUPLICADO CONTRATO PIG
        Case geColPImpDuplicado
            Call loGrabarExt.nExtornoDuplicadoContratoCredPig(lsNroContrato, lsFechaHoraGrab, _
                 lsMovNro, lnMovNroAExt, lnMonto, False, DatosExtorna)
            lsOperacionDesc = "DUPLICADO CONTRATO"
        
        '** Extornar una RENOVACION
        Case gColPOpeRenovNorEFE, gColPOpeRenovMorEFE, gColPOpeRenovNorCHQ, gColPOpeRenovMorCHQ, gColPOpeRenovNorEnOtCjEFE, gColPOpeRenovMorEnOtCjEFE, gColPOpeRenovNorEnOtCjCHQ, gColPOpeRenovMorEnOtCjCHQ, "129801", _
            gColPOpeRenovNorVoucher, gColPOpeRenovMorVoucher, gColPOpeRenovNorCargoCta, gColPOpeRenovMorCargoCta
            If lsOperacion = gColPOpeRenovNorCargoCta Or lsOperacion = gColPOpeRenovMorCargoCta Then Sleep 1000  'CTI4 ERS0112020
            Set loContFunct = New COMNContabilidad.NCOMContFunciones
            Call loGrabarExt.nExtornoRenovacionCredPig(lsNroContrato, lsFechaHoraGrab, _
                 lsMovNro, lnMovNroAExt, lnMonto, lsmensaje, False, DatosExtorna, lsOperacion, lnITFOpeVoucher, loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), _
                 lsCtaAhoExt, lsClienteAhoExt, lnMontoAhoExt)
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
        End If
        lsOperacionDesc = "RENOVACION"
        
        '** Extornar una CANCELACION
        Case gColPOpeCancelNorEFE, gColPOpeCancelMorEFE, gColPOpeCancelNorCHQ, gColPOpeCancelMorCHQ, gColPOpeCanceNorEnOtCjEFE, gColPOpeCanceMorEnOtCjEFE, gColPOpeCanceNorEnOtCjCHQ, gColPOpeCanceMorEnOtCjCHQ, "129802", _
            gColPOpeCancelNorVoucher, gColPOpeCancelMorVoucher, gColPOpeCancelNorCargoCta, gColPOpeCancelMorCargoCta
            If lsOperacion = gColPOpeCancelNorCargoCta Or lsOperacion = gColPOpeCancelMorCargoCta Then Sleep 1000  'CTI4 ERS0112020
            Set loContFunct = New COMNContabilidad.NCOMContFunciones
            Call loGrabarExt.nExtornoCancelacionCredPig(lsNroContrato, lsFechaHoraGrab, _
                 lsMovNro, lnMovNroAExt, lnMonto, lsmensaje, False, DatosExtorna, lsOperacion, lnITFOpeVoucher, loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), _
                 lsCtaAhoExt, lsClienteAhoExt, lnMontoAhoExt)
                 
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
        End If

        lsOperacionDesc = "CANCELACION"
        
        '** Extornar una AMORTIZACION
        Case gColPOpeAmortNorEFE, gColPOpeAmortNorCHQ, gColPOpeAmortNorEnOtCjEFE, gColPOpeAmortNorEnOtCjCHQ, "129803"
            Call loGrabarExt.nExtornoAmortizacionCredPig(lsNroContrato, lsFechaHoraGrab, _
                 lsMovNro, lnMovNroAExt, lnMonto, lsOperacion, lsmensaje, False, DatosExtorna)
        lsOperacionDesc = "AMORTIZACION"
        
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
        End If
    
        '*** PEAC 20161212
        '** Extornar una PAGO PARCIAL
        Case gColPOpePagoParcialNorEfectivo, gColPOpePagoParcialNorCheque, gColPOpePagoParcialNorVoucher, gColPOpePagoParcialNorCargoCta
            If lsOperacion = gColPOpePagoParcialNorCargoCta Then Sleep 1000 'CTI4 ERS0112020
            Set loContFunct = New COMNContabilidad.NCOMContFunciones
            Call loGrabarExt.nExtornoPagoParcialCredPig(lsNroContrato, lsFechaHoraGrab, _
                 lsMovNro, lnMovNroAExt, lnMonto, lsOperacion, lsmensaje, False, DatosExtorna, lsOperacion, lnITFOpeVoucher, loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), _
                 lsCtaAhoExt, lsClienteAhoExt, lnMontoAhoExt)
        lsOperacionDesc = "PAGO PARCIAL"
        
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
        End If
    
        '** Extornar una DEVOLUCION DE JOYAS
        Case geColPDevJoyas
            Call loGrabarExt.nExtornoRescateJoyaCredPig(lsNroContrato, lsFechaHoraGrab, _
                 lsMovNro, lnMovNroAExt, lnMonto, False, DatosExtorna)
        lsOperacionDesc = "DEVOLUCION JOYA"
        
        '** Extornar ADJUDICACION
        Case geColPAdjudica
'            Call loGrabarExt.nExtornoRescateJoyaCredPig(lsNroContrato, lsFechaHoraGrab, _
'                 lsMovNro, lnMovNroAExt, lnMonto, False)

        '** Extornar COBRO DE CUSTODIA DIFERIDA
        Case geColPCobCusDiferida
            Call loGrabarExt.nExtornoCustodiaDiferidaCredPig(lsNroContrato, lsFechaHoraGrab, _
                 lsMovNro, lnMovNroAExt, lnMonto, False, DatosExtorna)
        lsOperacionDesc = "CUSTODIA DIFERIDA"
        
        '** Extorno VENTA EN REMATE
        Case geColPVtaRemate
            Call loGrabarExt.nExtornoRemateVentaCredPig(lsNroContrato, lsFechaHoraGrab, _
                 lsMovNro, lnMovNroAExt, lnMonto, False, DatosExtorna)
        lsOperacionDesc = "VENTA EN REMATE"
        
        '** Extorno VENTA EN SUBASTA
        Case geColPVtaSubasta
            Call loGrabarExt.nExtornoSubastaVentaCredPig(lsNroContrato, lsFechaHoraGrab, _
                lsMovNro, lnMovNroAExt, lnMonto, False, DatosExtorna)
        lsOperacionDesc = "VENTA EN SUBASTA"
        
        '**Extorno RECUPERACION ADJUDICADO
        Case "122900", gColPOpeRecuperaContratoAdjVoucher, gColPOpeRecuperaContratoAdjCargoCta
            lsFechaHoraGrab = loGrabarExt.ObtenerFechaEstado(lsNroContrato, 2108)
            If lsOperacion = gColPOpeRecuperaContratoAdjCargoCta Then Sleep 1000  'CTI4 ERS0112020
            Set loContFunct = New COMNContabilidad.NCOMContFunciones
           Call loGrabarExt.nExtornoRecuperacionCredPig(lsNroContrato, lsFechaHoraGrab, lsMovNro, lnMovNroAExt, lnMonto, False, DatosExtorna, _
           lsOperacion, lnITFOpeVoucher, loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), _
           lsCtaAhoExt, lsClienteAhoExt, lnMontoAhoExt)
        lsOperacionDesc = "RECUPERACION ADJUDICADO"
             
       
        '** TODOCOMPLETA Extorno PAGO DE SOBRANTE
        Case geColPPagSobrante

            Set loConec = New COMConecta.DCOMConecta
                loConec.AbreConexion
                Set lrdatosrem = New ADODB.Recordset
                lsSQL = "SELECT cNroProceso From ColocPigRGDet Where cTpoProceso='R' and cCtaCod ='" & lsNroContrato & "' And nRGDetEstado = 4 "
                'loConec.CargaRecordSet lsSQL
                lrdatosrem.Open lsSQL, loConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
                If lrdatosrem Is Nothing Then
                    MsgBox "No se ubico el nro de Remate del Credito, Comunicarse con Sistemas", vbInformation, "Aviso"
                    Exit Sub
                End If
                lsNroProceso = lrdatosrem!cNroProceso

                lrdatosrem.Close
            Set loConec = Nothing

            Dim loDatRem As COMNColoCPig.NCOMColPRecGar '  NColPRecGar
            Set loDatRem = New COMNColoCPig.NCOMColPRecGar 'NColPRecGar
                lsCtaAhorroRemate = loDatRem.nObtieneCtaSobranteRemate(fsRemateCadaAgencia, lsmensaje)
                If Trim(lsmensaje) <> "" Then
                     MsgBox lsmensaje, vbInformation, "Aviso"
                     Exit Sub
                End If
            Set loDatRem = Nothing
            If lsCtaAhorroRemate = "" Then
                MsgBox "No se ubico la Cta de Sobrante de Remate, Comunicarse con Sistemas", vbInformation, "Aviso"
                Exit Sub
            End If

            Call loGrabarExt.nExtornoPagoSobranteRemate(lsNroContrato, lsFechaHoraGrab, _
                 lsMovNro, lnMovNroAExt, lnMonto, sLpt, lsNroProceso, lsCtaAhorroRemate, False, DatosExtorna)
        lsOperacionDesc = "PAGO SOBRANTE REMATE"
        
        '*** PEAC 20090316
        Case geColPPagSobraAdjudica

            Set loConec = New COMConecta.DCOMConecta
                loConec.AbreConexion
                Set lrdatosAdj = New ADODB.Recordset
                lsSQL = "SELECT cNroProceso From ColocPigRGDet Where cTpoProceso='A' and cCtaCod ='" & lsNroContrato & "' And nRGDetEstado = 4 "
                lrdatosAdj.Open lsSQL, loConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
                If lrdatosAdj Is Nothing Then
                    MsgBox "No se ubico el nro de Adjudicacion del Credito, Comunicarse con Sistemas", vbInformation, "Aviso"
                    Exit Sub
                End If
                lsNroProceso = lrdatosAdj!cNroProceso
                lrdatosAdj.Close
            Set loConec = Nothing

            Dim loDatAdj As COMNColoCPig.NCOMColPRecGar
            Set loDatAdj = New COMNColoCPig.NCOMColPRecGar
                lsCtaAhorroAdjudica = loDatAdj.nObtieneCtaSobranteAdjudicado(fsAdjudicaCadaAgencia, lsmensaje)
                If Trim(lsmensaje) <> "" Then
                     MsgBox lsmensaje, vbInformation, "Aviso"
                     Exit Sub
                End If
            Set loDatAdj = Nothing
            If lsCtaAhorroAdjudica = "" Then
                MsgBox "No se ubico la Cta de Sobrante de Adjudicado, Comunicarse con Sistemas", vbInformation, "Aviso"
                Exit Sub
            End If

            Call loGrabarExt.nExtornoPagoSobranteAdjudica(lsNroContrato, lsFechaHoraGrab, _
                 lsMovNro, lnMovNroAExt, lnMonto, sLpt, lsNroProceso, lsCtaAhorroAdjudica, False, DatosExtorna)
        lsOperacionDesc = "PAGO SOBRANTE ADJUDICADO"
                        
    End Select
      
    loVistoElectronico.RegistraVistoElectronico (lnMovNroAExt)

Set loGrabarExt = Nothing



Dim clsMov As COMDMov.DCOMMov, sCodUserBus As String, sMovNroBus As String
Set clsMov = New COMDMov.DCOMMov
sMovNroBus = "": sCodUserBus = ""
sMovNroBus = clsMov.GetcMovNro(lnMovNroAExt)
sCodUserBus = Right(sMovNroBus, 4)

'*************************** PEAC 20090703 - Imprime recibo de extorno de abono en cta

If lbProcedeExtAho And Len(lsCtaAhoExt) > 0 Then

    Set loImprime = New COMNColoCPig.NCOMColPImpre
        lsCadImprimir = loImprime.nPrintReciboExtorAboCta(gsNomAge, lsFechaHoraGrab, lsNroContrato, lsCtaAhoExt, lnITFAho, _
            lsCliente, lsOperacionDesc, lnMonto, 0, lnMovNroAExt, gsCodUser, lsCMACOpe, "", sCodUserBus, gImpresora, gbImpTMU)

    Set loImprime = Nothing
    Set loPrevio = New previo.clsprevio
        loPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
        Do While True
            If MsgBox("Reimprimir Recibo de Extorno del Abono a Cuenta ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                loPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
            Else
                Set loPrevio = Nothing
                Exit Do
            End If
        Loop
     Set loPrevio = Nothing
End If

'*************************** FIN PEAC

Set loImprime = New COMNColoCPig.NCOMColPImpre
    lsCadImprimir = loImprime.nPrintReciboExtorno(gsNomAge, lsFechaHoraGrab, lsNroContrato, _
        lsCliente, lsOperacionDesc, lnMonto, 0, lnMovNroAExt, gsCodUser, lsCMACOpe, "", sCodUserBus, gImpresora, gbImpTMU)

'CTI4 ERS0112020
If lsOperacion = gColPOpeRenovNorCargoCta Or lsOperacion = gColPOpeRenovMorCargoCta _
    Or lsOperacion = gColPOpeCancelNorCargoCta Or lsOperacion = gColPOpeCancelMorCargoCta _
    Or lsOperacion = gColPOpePagoParcialNorCargoCta Or lsOperacion = gColPOpeRecuperaContratoAdjCargoCta Then
    lsImpCargoCta = loImprime.nPrintReciboExtorCargoCta(gsNomAge, lsFechaHoraGrab, lsNroContrato, lsCtaAhoExt, 0, PstaNombre(lsClienteAhoExt), _
    lsOperacionDesc, lnMontoAhoExt, 0, lnMovNroAExt, gsCodUser, lsCMACOpe, "", sCodUserBus, gImpresora, gbImpTMU)
End If
'END

Set loImprime = Nothing
Set loPrevio = New previo.clsprevio
    loPrevio.PrintSpool sLpt, lsCadImprimir & lsImpCargoCta, False, 22
    Do While True
        If MsgBox("Reimprimir Recibo de Extorno ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
            loPrevio.PrintSpool sLpt, lsCadImprimir & lsImpCargoCta, False, 22
        Else
            Set loPrevio = Nothing
            Exit Do
        End If
    Loop
 Set loPrevio = Nothing
 
 '******CTI3 (ferimoro) 27092018
'frmMotExtorno.Visible = False
Me.cmbMotivos.ListIndex = -1
Me.txtDetExtorno.Text = ""
fraBuscar.Enabled = True
AXCodCta.Enabled = True
cmdBuscar.Enabled = True
cmdExtorno.Enabled = False
'******************************

'***************************

Me.lstExtorno.ListItems.Clear
If lstExtorno.ListItems.count = 0 Then
    cmdExtorno.Enabled = False
End If
If Me.opt(3).value = True Then
    opt_KeyPress 3, 13
Else
    cmdBuscar_Click
End If

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
        
End Sub

Private Sub cmdExtorno_Click()

If lstExtorno.ListItems.count = 0 Then
    cmdExtorno.Enabled = False
    Exit Sub
End If

'******CTI3 (ferimoro) 27092018
 frmMotExtorno.Visible = True
 fraBuscar.Enabled = False
 AXCodCta.Enabled = False
 cmdBuscar.Enabled = False
 cmdExtorno.Enabled = False
 cmbMotivos.SetFocus
'******************************

End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    fsFechaTransac = Mid(Format$(gdFecSis, "dd/mm/yyyy"), 7, 4) & Mid(Format$(gdFecSis, "dd/mm/yyyy"), 4, 2) & Mid(Format$(gdFecSis, "dd/mm/yyyy"), 1, 2)
    Call CargaControles 'CTI3
    lstExtorno.ColumnHeaders.Add , , "NroCuenta", 2000
    lstExtorno.ColumnHeaders.Add , , "Operación", 2200
    lstExtorno.ColumnHeaders.Add , , "OpcExt.", 750, lvwColumnCenter
    lstExtorno.ColumnHeaders.Add , , "Monto", 1100, lvwColumnRight
    lstExtorno.ColumnHeaders.Add , , "Fecha de Movimiento", 1750, lvwColumnCenter
    lstExtorno.ColumnHeaders.Add , , "N°Tran", 800, lvwColumnCenter
    lstExtorno.ColumnHeaders.Add , , "Usuario", 800, lvwColumnCenter
    lstExtorno.ColumnHeaders.Add , , "Cliente", 1600, lvwColumnLeft
    lstExtorno.ColumnHeaders.Add , , "CMACOpe", 600, lvwColumnLeft
    lstExtorno.ColumnHeaders.Add , , "ITFOpeVoucher", 0, lvwColumnLeft 'CTI4 ERS0112020
    lstExtorno.View = lvwReport
    Limpiar
    CargaParametros
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
End Sub

Private Sub LLenaLista(myRs As Recordset)
Dim litmX As ListItem
Dim lsCtaCodAnterior As String

Do While Not myRs.EOF
    Set litmX = lstExtorno.ListItems.Add(, , myRs!cCtaCod, , "Cuenta")           'Nro de Cred Pig
        litmX.SubItems(1) = Mid(myRs!cOpeDesc, 1, 30) & space(10) & myRs!cOpecod 'Operacion
        litmX.SubItems(3) = Format(myRs!nMonto, "#0.00")                         'Monto Operacion
        litmX.SubItems(4) = fgFechaHoraGrab(myRs!cMovNro)                        'Fecha/hora Operacion
        litmX.SubItems(5) = Str(myRs!nMovNro)                                    'Nro Movimiento(nMovNro)
        litmX.SubItems(6) = Mid(myRs!cMovNro, 22, 4)                             'Usuario
        litmX.SubItems(7) = Trim(myRs!cCliente)                                  'Cliente
        litmX.SubItems(8) = Trim(myRs!cCMACOpe)                                  'CMAC Operacion
        litmX.SubItems(9) = Trim(myRs!nITFOpeVoucher)                            'CTI4 ERS0112020 ITF de Operacion con Voucher
    If myRs!cCtaCod = lsCtaCodAnterior Then
        litmX.SubItems(2) = "0"
    Else
        litmX.SubItems(2) = "1"
    End If
    lsCtaCodAnterior = myRs!cCtaCod
    myRs.MoveNext
Loop

End Sub

'Valida el ListView lstExtorno
Private Sub lstExtorno_GotFocus()
If lstExtorno.ListItems.count >= 0 Then
   cmdExtorno.Enabled = True
End If
End Sub

Private Sub lstExtorno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     If lstExtorno.ListItems.count > 0 Then
        cmdExtorno.Enabled = True
        cmdExtorno.SetFocus
     End If
End If
End Sub

Private Sub opt_Click(index As Integer)
Limpiar

Select Case index
    Case 0
        AXCodCta.Visible = True
        AXCodCta.EnabledAge = True
        AXCodCta.EnabledCta = True
       
    Case 3
        AXCodCta.Visible = False
End Select
cmdBuscar.Visible = True
End Sub

Private Sub opt_KeyPress(index As Integer, KeyAscii As Integer)
Select Case index
    Case 0
        If KeyAscii = 13 Then
            AXCodCta.SetFocusCuenta
            
        End If
    Case 3
        If KeyAscii = 13 Then
            cmdBuscar.Enabled = True
            cmdBuscar.SetFocus
        End If
End Select
Me.Caption = "Crédito Pignoraticio : Extornos"
End Sub

'Inicializa variables
Private Sub Limpiar()
    'Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    lstExtorno.ListItems.Clear
End Sub

Private Sub CargaParametros()
Dim loConstSis As COMDConstSistema.NCOMConstSistema
Dim lnProcesoCadaAgencia As Integer

Set loConstSis = New COMDConstSistema.NCOMConstSistema
    lnProcesoCadaAgencia = loConstSis.LeeConstSistema(121)  ' gConstSistPigRemateCadaAg
    If lnProcesoCadaAgencia = 1 Then  ' En cada agencia
        fsRemateCadaAgencia = gsCodCMAC & gsCodAge
        fsAdjudicaCadaAgencia = gsCodCMAC & gsCodAge
    Else
        fsRemateCadaAgencia = gsCodCMAC & "00"
        fsAdjudicaCadaAgencia = gsCodCMAC & "00"
    End If
Set loConstSis = Nothing

End Sub


'******CTI3 (ferimoro) 18102018
Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

Set oCons = New COMDConstantes.DCOMConstantes

Set R = oCons.ObtenerConstanteExtornoMotivo

Set oCons = Nothing
Call Llenar_Combo_MotivoExtorno(R, cmbMotivos)

End Sub

Private Sub txtDetExtorno_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
'If KeyAscii = 13 Then SendKeys "{TAB}": Exit Sub
End Sub
