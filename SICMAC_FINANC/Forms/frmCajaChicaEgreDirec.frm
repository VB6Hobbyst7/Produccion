VERSION 5.00
Begin VB.Form frmCajaChicaEgreDirec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Egresos Directos de Caja Chica"
   ClientHeight    =   5235
   ClientLeft      =   1110
   ClientTop       =   2205
   ClientWidth     =   9735
   Icon            =   "frmCajaChicaEgreDirec.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDocumento 
      Caption         =   "&Documento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5880
      TabIndex        =   21
      Top             =   4680
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdExtDesemb 
      Caption         =   "&Extornar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7125
      TabIndex        =   20
      Top             =   4680
      Width           =   1245
   End
   Begin VB.TextBox txtMovDesc 
      Height          =   585
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   4560
      Width           =   5565
   End
   Begin VB.CommandButton cmdDesembolsar 
      Caption         =   "&Desembolsar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7125
      TabIndex        =   18
      Top             =   4680
      Width           =   1245
   End
   Begin VB.CommandButton cmdExtorno 
      Caption         =   "&Extornar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7125
      TabIndex        =   12
      Top             =   4680
      Width           =   1245
   End
   Begin VB.CommandButton cmdAtender 
      Caption         =   "&Atender"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7125
      TabIndex        =   11
      Top             =   4680
      Width           =   1245
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8355
      TabIndex        =   8
      Top             =   4680
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3720
      Left            =   90
      TabIndex        =   6
      Top             =   780
      Width           =   9540
      Begin Sicmact.FlexEdit fgListaCH 
         Height          =   2205
         Left            =   90
         TabIndex        =   7
         Top             =   210
         Width           =   9390
         _ExtentX        =   16563
         _ExtentY        =   3889
         Cols0           =   11
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "N°-Ok-Tipo-Nro Doc.-Fecha Doc-Solicitante-Importe-Concepto-cMovNro-cTpoDoc-cMovNroSol"
         EncabezadosAnchos=   "450-450-450-1300-1200-4000-1200-2500-0-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-4-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-L-C-L-R-L-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-2-0-0-0-0"
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbOrdenaCol     =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   450
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label lblSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   285
         Left            =   810
         TabIndex        =   17
         Top             =   3315
         Width           =   1605
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   6615
         TabIndex        =   13
         Top             =   3345
         Width           =   735
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   360
         Left            =   7590
         TabIndex        =   14
         Top             =   3255
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   6495
         TabIndex        =   15
         Top             =   3255
         Width           =   2805
      End
   End
   Begin VB.Frame fraCajaChica 
      Caption         =   "Caja Chica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   705
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   9540
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8145
         TabIndex        =   10
         Top             =   217
         Width           =   1245
      End
      Begin Sicmact.TxtBuscar txtBuscarAreaCH 
         Height          =   345
         Left            =   1200
         TabIndex        =   1
         Top             =   225
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   609
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
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Area/Agencia : "
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   300
         Width           =   1125
      End
      Begin VB.Label lblCajaChicaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2280
         TabIndex        =   4
         Top             =   225
         Width           =   4425
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Proceso :"
         Height          =   195
         Left            =   6765
         TabIndex        =   3
         Top             =   285
         Width           =   675
      End
      Begin VB.Label lblNroProcCH 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7470
         TabIndex        =   2
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "&Rechazar"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7125
      TabIndex        =   9
      Top             =   4680
      Width           =   1245
   End
End
Attribute VB_Name = "frmCajaChicaEgreDirec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oArendir As NARendir
Dim oCH As nCajaChica
Dim lnArendirFase As ARendirFases
Dim lsCtaArendir As String
Dim lsCtaPendiente As String
Dim lsCtaFondofijo As String
Dim lsCtaFondofijoF As String
Dim lsCtaDesembolso As String
Dim lsTipoDoc As String
Dim lbSalir As Boolean
Dim lnTipoProcCH As CHTipoProc
'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Public Sub Inicio(Optional ByVal pnArendirFase As ARendirFases = ArendirRechazo, Optional ByVal pnTipoProcCH As CHTipoProc = gCHTipoProcEgresoDirecto)
lnArendirFase = pnArendirFase
lnTipoProcCH = pnTipoProcCH
Me.Show 1
End Sub
Private Sub cmdAtender_Click()
Dim i As Integer
Dim oCH As nCajaChica
Dim oCon As NContFunciones
Dim lsMovNroAtenc As String
Dim lsMovNroSol As String
Dim lnImporte As Currency

On Error GoTo ErrcmdAtender

cmdAtender.Enabled = False
Set oCH = New nCajaChica
Set oCon = New NContFunciones
If fgListaCH.TextMatrix(1, 0) = "" Then Exit Sub
If Val(lblTotal) = 0 Then
    MsgBox "Seleccione alguna solicitud por favor", vbInformation, "Aviso"
    cmdAtender.Enabled = True
    Exit Sub
End If
If ValidaCajaChica = False Then
    cmdAtender.Enabled = True
    Exit Sub
End If
If MsgBox("Desea Realizar la Atención de las solicitudes seleccionadas??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    For i = 1 To fgListaCH.Rows - 1
        lsMovNroAtenc = ""
        lsMovNroSol = ""
        If fgListaCH.TextMatrix(i, 1) <> "" Then
            lsMovNroAtenc = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            lsMovNroSol = fgListaCH.TextMatrix(i, 8)
            lnImporte = CCur(fgListaCH.TextMatrix(i, 6))
            If oCH.GrabaAtenEgresoDirec(gsFormatoFecha, lsMovNroAtenc, gsOpeCod, _
                                        gsOpeDesc, lsMovNroSol, Mid(txtBuscarAreaCH, 1, 3), _
                                        Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lnImporte) = 0 Then
                
            End If
        End If
    Next
    CargaLista
    lblSaldo = Format(oCH.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2)), "#,#0.00")
    MsgBox "Atencion de solicitudes se ha realizado con éxito", vbInformation, "Aviso"
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Atendio Solicitud del Agencia/Area : " & lblCajaChicaDesc & " |Monto :" & lblTotal
            Set objPista = Nothing
            '*******
    CalculaTotal
End If
Set oCH = Nothing
Set oCon = Nothing
cmdAtender.Enabled = True

Exit Sub
ErrcmdAtender:
    MsgBox "Error N° [" & Err.Number & "]" & TextErr(Err.Description), vbInformation, "Aviso"

End Sub
Private Function ValidaCajaChica() As Boolean
Dim oCajaChica As nCajaChica
Dim lnSaldo As Currency
Set oCajaChica = New nCajaChica
Dim lnTope As Currency
ValidaCajaChica = True
lnSaldo = oCajaChica.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2))
If lnSaldo = 0 Then
    MsgBox "Caja Chica Sin Saldo. Es necesario Solicitar Autorización o Desembolso", vbInformation, "Aviso"
    ValidaCajaChica = False
    Exit Function
End If
If lnSaldo < CCur(lblTotal) Then
    MsgBox "Egreso no puede ser mayor que " & Format(lnSaldo, gsFormatoNumeroView), vbInformation, "Aviso"
    ValidaCajaChica = False
    Exit Function
End If
If oCajaChica.VerificaTopeCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2)) = True Then
    MsgBox "No puede realizar Egreso porque el saldo de esta Caja Chica es menor que el permitido." & oImpresora.gPrnSaltoLinea _
            & "Por favor es necesario que realice Rendición", vbInformation, "Aviso"
    ValidaCajaChica = False
    If MsgBox(" ¿ Desea continuar con el proceso de Egreso de Caja Chica ? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
         ValidaCajaChica = True
    End If
    Exit Function
End If
End Function


Private Sub cmdDesembolsar_Click()
Dim oDocPago As clsDocPago
Dim oCon As NContFunciones

Dim lsNroDoc As String
Dim lsNroVoucher As String
Dim lsFechaDoc As String
Dim lsDocumento As String
Dim lsGlosa As String
Dim lsMovNro As String
Dim lsPersCod As String
Dim lsPersNombre As String
Dim lsPersDireccion As String
Dim lsUbigeo    As String
Dim lsCuentaAho As String
Dim lsCadBol As String
Dim lsNroMovAut As String
Dim lnImporte As Currency
Dim rs As ADODB.Recordset
Dim rsAut As ADODB.Recordset
Dim lnMontoDif      As Currency
Dim lsCtaDiferencia As String
Dim lsCtaContDebeITF As String
Dim lsCtaContHaberITF As String
'Dim lsMontoITF As Currency
Dim lsMontoITF As Double '*** PEAC 20110331
Dim lnNroProceso As Integer '*** PEAC 20100128
Dim oOpe As New DOperacion

Set oCon = New NContFunciones
Set oDocPago = New clsDocPago
If Me.fgListaCH.TextMatrix(1, 0) = "" Then Exit Sub

If Trim(Len(txtMovDesc)) = 0 Then
    MsgBox "Descripción de Operación no válida", vbInformation, "Aviso"
    Me.txtMovDesc.SetFocus
    Exit Sub
End If
Set rs = New ADODB.Recordset
If Val(lblTotal) = 0 Then
    MsgBox "Total Importe de Operación no válido. Seleccione alguna autorizacion para realizar el Desembolso", vbInformation, "Aviso"
    Exit Sub
End If

lsPersCod = fgListaCH.TextMatrix(fgListaCH.row, 8)
lsPersNombre = fgListaCH.TextMatrix(fgListaCH.row, 4)
lsNroMovAut = fgListaCH.TextMatrix(fgListaCH.row, 7)

'*** PEAC 20100128
'Dim oOpe As New DOperacion
lnNroProceso = Val(lblNroProcCH)

If gsOpeCod = "401322" Then
    If lnNroProceso > 1 Then
        lsCtaFondofijo = oOpe.EmiteOpeCta(gsOpeCod, "D", "1")
    Else
        lsCtaFondofijo = oOpe.EmiteOpeCta(gsOpeCod, "D")
    End If
Else
    lsCtaFondofijo = oOpe.EmiteOpeCta(gsOpeCod, "D")
End If
'*** FIN PEAC

lsNroDoc = ""
lsNroVoucher = ""
lsFechaDoc = ""
lsDocumento = ""
lsGlosa = ""
lsGlosa = txtMovDesc
lnImporte = CCur(lblTotal)
Set rsAut = fgListaCH.GetRsNew
If nVal(lsTipoDoc) = TpoDocOrdenPago Then
    oDocPago.InicioOrdenPago lsNroDoc, False, lsPersCod, gsOpeCod, lsPersNombre, Me.Caption, txtMovDesc, CCur(lblTotal), gdFecSis, "", , gsCodAge
    If oDocPago.vbOk Then
        lsNroDoc = oDocPago.vsNroDoc
        lsNroVoucher = oDocPago.vsNroVoucher
        lsFechaDoc = oDocPago.vdFechaDoc
        lsDocumento = oDocPago.vsFormaDoc
        lsGlosa = oDocPago.vsGlosa
    Else
        Exit Sub
    End If
ElseIf nVal(lsTipoDoc) = TpoDocNotaAbono Then
    Dim oImp As New NContImprimir
    lsTipoDoc = TpoDocNotaAbono
    
    frmNotaCargoAbono.Inicio lsTipoDoc, lnImporte, gdFecSis, txtMovDesc, gsOpeCod, False, lsPersCod, lsPersNombre
    If frmNotaCargoAbono.vbOk Then
        lsNroDoc = frmNotaCargoAbono.NroNotaCA
        txtMovDesc = frmNotaCargoAbono.Glosa
        lsDocumento = frmNotaCargoAbono.NotaCargoAbono
        lsPersNombre = frmNotaCargoAbono.PersNombre
        lsPersDireccion = frmNotaCargoAbono.PersDireccion
        lsUbigeo = frmNotaCargoAbono.PersUbigeo
        lsCuentaAho = frmNotaCargoAbono.CuentaAhoNro
        lsFechaDoc = frmNotaCargoAbono.FechaNotaCA
'        lsDocumento = oImp.ImprimeNotaCargoAbono(lsNroDoc, txtMovDesc, CCur(frmNotaCargoAbono.Monto), _
'                            lsPersNombre, lsPersDireccion, lsUbigeo, gdFecSis, Mid(gsOpeCod, 3, 1), lsCuentaAho, lsTipoDoc, gsNomAge, gsCodUser)
         lsDocumento = oImp.ImprimeNotaAbono(lsFechaDoc, lnImporte, txtMovDesc, lsCuentaAho, lsPersNombre)
        Dim oDis As New NRHProcesosCierre
        lsCadBol = oDis.ImprimeBoletaCad(CDate(lsFechaDoc), "ABONO CAJA GENERAL", "Depósito CAJA GENERAL*Nro." & lsNroDoc, "", lnImporte, lsPersNombre, lsCuentaAho, "", 0, 0, "Nota Abono", 0, 0, False, False, , , , True, , , , False, gsNomAge) & oImpresora.gPrnSaltoPagina
    Else
        Exit Sub
    End If
ElseIf nVal(lsTipoDoc) = gnDocCuentaPendiente Then
'    Dim oOpe As New DOperacion
 
    lsNroVoucher = oCon.GeneraDocNro(gnDocCuentaPendiente, , Mid(gsOpeCod, 3, 1), gsCodAge)
    lnMontoDif = frmArendirEfectivo.vnDiferencia
    'lnMontoDif = fgListaCH.TextMatrix(fgListaCH.Row, 6)
    lsFechaDoc = gdFecSis
'    lblSaldo = 0
    lsCtaDiferencia = oOpe.EmiteOpeCta(gsOpeCod, IIf(lnMontoDif > 0, "D", "H"), "2")
'    Set oOpe = Nothing
Else
    frmArendirEfectivo.Inicio gArendirTipoCajaChica, txtBuscarAreaCH & "-" & lblNroProcCH, Mid(gsOpeCod, 3, 1), lblCajaChicaDesc, _
                        CCur(lblTotal), lsPersCod, lsPersNombre, , "Caja Chica N°"
    If frmArendirEfectivo.lbOk Then
       Set rs = frmArendirEfectivo.rsEfectivo
        If frmArendirEfectivo.vnDiferencia <> 0 Then
            'Dim oOpe As New DOperacion
            lnMontoDif = frmArendirEfectivo.vnDiferencia
            lsCtaDiferencia = oOpe.EmiteOpeCta(gsOpeCod, IIf(lnMontoDif > 0, "D", "H"), "2")
            Set oOpe = Nothing
        End If
    Else
        Exit Sub
    End If
End If
Set oOpe = Nothing
If MsgBox("Desea Realizar el Desembolso de Caja chica??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    
    lsMovNro = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    'edpyme - sgte reglon para fondos fijos por agencia
    Dim lsSubCta As String
    lsSubCta = oCon.GetFiltroObjetos(ObjCMACAgenciaArea, lsCtaFondofijo, Trim(txtBuscarAreaCH), False)
 
    If lsSubCta <> "" Then
       If Mid(txtBuscarAreaCH, 1, 3) = "067" And Mid(txtBuscarAreaCH, 4, 2) = "01" Then ' Principal
              'lsSubCta = "01"
              'lsCtaFondofijoF = lsCtaFondofijo & "02" & lsSubCta
              lsCtaFondofijoF = lsCtaFondofijo & lsSubCta
       ElseIf Mid(txtBuscarAreaCH, 1, 3) = "043" Then  'Secretaria
              'lsSubCta = "02"
              'lsCtaFondofijoF = lsCtaFondofijo & lsSubCta & "01"
              lsCtaFondofijoF = lsCtaFondofijo & lsSubCta
       ElseIf Mid(txtBuscarAreaCH, 1, 3) = "023" Then  'Logistica
              'lsSubCta = "02"
              'lsCtaFondofijoF = lsCtaFondofijo & lsSubCta & "01"
              lsCtaFondofijoF = lsCtaFondofijo & lsSubCta
       Else
              'lsCtaFondofijoF = lsCtaFondofijo + IIf(CCur(lsSubCta) > 29, "01", "02") & lsSubCta Comentado for GITU 16-10-2009
              lsCtaFondofijoF = lsCtaFondofijo & lsSubCta
       End If
    Else
        MsgBox "Sub Cuenta no Definida", vbInformation, "Aviso"
        Exit Sub
    End If
    
    lsCtaContDebeITF = oOpe.EmiteOpeCta(gsOpeCod, "D", 2)
    lsCtaContHaberITF = oOpe.EmiteOpeCta(gsOpeCod, "H", 2)
    
    lsMontoITF = fgTruncar(lnImporte * gnImpITF, 2)

    If oCH.GrabaDesembolsoCH(lsMovNro, lsNroMovAut, lsPersCod, gsFormatoFecha, rs, gsOpeCod, lsGlosa, lnImporte, _
                              Mid(Me.txtBuscarAreaCH, 1, 3), Mid(Me.txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), _
                            lsTipoDoc, lsNroDoc, lsFechaDoc, lsNroVoucher, rsAut, lsCtaDiferencia, lnMontoDif, lsCtaFondofijoF, lsCtaDesembolso, CCur(lblSaldo), lsCuentaAho, gbBitCentral, , lsCtaContDebeITF, lsCtaContHaberITF, lsMontoITF) = 0 Then
        
        If gsOpeCod = "401322" Or gsOpeCod = "402322" Then
            'ImprimeAsientoContableNew lsMovNro, lsNroVoucher, lsTipoDoc, lsDocumento, IIf(lsNroDoc = "", True, False), False, lsGlosa, lsPersCod, _
            'lnImporte, gArendirTipoCajaChica, , , , , "17", , , lsCadBol, Mid(gsOpeCod, 3, 1)
            ImprimeAsientoContableUltimo lsMovNro, lsNroVoucher, lsTipoDoc, lsDocumento, IIf(lsNroDoc = "", True, False), False, lsGlosa, lsPersCod, _
            lnImporte, gArendirTipoCajaChica, , , , , "17", , , lsCadBol, Mid(gsOpeCod, 3, 1)
            
        Else
            ImprimeAsientoContable lsMovNro, lsNroVoucher, lsTipoDoc, lsDocumento, IIf(lsNroDoc = "", True, False), False, lsGlosa, lsPersCod, _
            lnImporte, gArendirTipoCajaChica, , , , , "17", , , lsCadBol
        End If
        
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Autorizo Desembolos de Chica del  Agencia/Area : " & lblCajaChicaDesc & " |Encargado : " & lsPersNombre _
            & " |Monto : " & lblTotal
            Set objPista = Nothing
            '*******
        cmdProcesar_Click
    End If
End If
Set oCon = Nothing
Set oDocPago = Nothing

End Sub

Private Sub cmdDocumento_Click()
Dim lsRec As String
Dim oImp As NContImprimir
Set oImp = New NContImprimir

If fgListaCH.TextMatrix(1, 1) = "" Then
    MsgBox "No existen datos para Documento", vbInformation, "Aviso"
    Exit Sub
End If
If fgListaCH.TextMatrix(fgListaCH.row, 9) = TpoDocRecEgreso Then
    oImp.Inicio gsNomCmac, gsNomAge, Format(gdFecSis, gsFormatoFechaView)
    lsRec = oImp.ImprimeRecibo(fgListaCH.TextMatrix(fgListaCH.row, 10), True)
End If
Set oImp = Nothing
If lsRec = "" Then
    MsgBox "No existe Formato para Tipo de Documento", vbInformation, "¡Aviso!"
Else
    EnviaPrevio lsRec, "IMPRESION DE DOCUMENTO", gnLinPage, False
End If
End Sub
Private Sub cmdExtDesemb_Click()
Dim oCon As NContFunciones
Dim lsMovNro As String
Dim lnImporte As Currency
Dim ldFechaDes As Date
Dim lsMovNroDesemb As String
Dim lsImp As String
Dim oContImp As NContImprimir
Dim i As Integer

If Trim(Len(txtMovDesc)) = 0 Then
    MsgBox "Descripción de Operación no válida", vbInformation, "Aviso"
    Me.txtMovDesc.SetFocus
    Exit Sub
End If
If Val(lblTotal) = 0 Then
    MsgBox "Tota Importe de Operación no válido. Seleccione algun Desembolso para realizar el Extorno", vbInformation, "Aviso"
    Exit Sub
End If
Set oContImp = New NContImprimir
Set oCon = New NContFunciones
If MsgBox("Desea Realizar el Extorno de Desembolso de la Caja chica Seleccionada??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsImp = ""
    For i = 1 To fgListaCH.Rows - 1
        If fgListaCH.TextMatrix(i, 1) <> "" Then
            ldFechaDes = CDate(fgListaCH.TextMatrix(fgListaCH.row, 4))
            lsMovNroDesemb = fgListaCH.TextMatrix(fgListaCH.row, 9)
            lnImporte = CCur(fgListaCH.TextMatrix(fgListaCH.row, 7))
            
            lsMovNro = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            If oCH.GrabaExtornoDesembolso(gdFecSis, ldFechaDes, gsFormatoFecha, lsMovNroDesemb, lsMovNro, gsOpeCod, _
                                          txtMovDesc, Mid(Me.txtBuscarAreaCH, 1, 3), Mid(Me.txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), _
                                          lnImporte, IIf(Mid(gsOpeCod, 3, 1) = "1", gCHAutorizaDesembMN, gCHAutorizaDesembME)) = 0 Then
                                          '***Agregado por ELRO el gCHAutorizaDesembMN o gCHAutorizaDesembME el 20120619, según OYP-RFC047-2012
                
                lsImp = lsImp + oContImp.ImprimeAsientoContable(lsMovNro, gnLinPage, gnColPage)
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            Dim lsPersNombre As String
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, gIngresarSistema, Me.Caption & " Se Extorno Aprobacion de Apertura de Chica del  Agencia/Area : " & lblCajaChicaDesc & " |Monto : " & lblTotal
            Set objPista = Nothing
            '*******
                
            End If
        End If
    Next
    If lsImp <> "" Then
        EnviaPrevio lsImp, Me.Caption, gnLinPage, False
    End If
    cmdProcesar_Click
End If
Set oCon = Nothing
End Sub

Private Sub cmdExtorno_Click()
Dim i As Integer
Dim oCH As nCajaChica
Dim oCon As NContFunciones
Dim lsMovNroAtenc As String
Dim lsMovNroSol As String
Dim lnImporte As Currency
Dim lsMovNroExt As String
On Error GoTo ErrcmdAtender

Set oCH = New nCajaChica
Set oCon = New NContFunciones
If fgListaCH.TextMatrix(1, 0) = "" Then Exit Sub
If Val(lblTotal) = 0 Then
    MsgBox "Seleccione alguna solicitud por favor", vbInformation, "Aviso"
    Exit Sub
End If
If MsgBox("Desea Realizar el Extorno de las Atenciones seleccionadas??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    For i = 1 To fgListaCH.Rows - 1
        lsMovNroAtenc = ""
        lsMovNroSol = ""
        If fgListaCH.TextMatrix(i, 1) <> "" Then
            lsMovNroExt = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            lsMovNroAtenc = fgListaCH.TextMatrix(i, 8)
            lsMovNroSol = fgListaCH.TextMatrix(i, 10)
            lnImporte = CCur(fgListaCH.TextMatrix(i, 6))
            If oCH.GrabaExtornoAtenEgresoDirec(gsFormatoFecha, lsMovNroExt, gsOpeCod, _
                                        gsOpeDesc, lsMovNroAtenc, lsMovNroSol, Mid(txtBuscarAreaCH, 1, 3), _
                                        Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lnImporte) = 0 Then
                
            End If
        End If
    Next
    CargaLista
    lblSaldo = Format(oCH.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2)), "#,#0.00")
    MsgBox "Extorno de Atencion de Solicitudes se ha realizado con éxito", vbInformation, "Aviso"
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", Me.Caption & " Se Extorno Solicitud del Agencia/Area : " & lblCajaChicaDesc & " |Monto :" & lblTotal
            Set objPista = Nothing
            '*******
    CalculaTotal
End If
Set oCH = Nothing
Set oCon = Nothing

Exit Sub
ErrcmdAtender:
    MsgBox "Error N° [" & Err.Number & "]" & TextErr(Err.Description), vbInformation, "Aviso"


End Sub
Private Sub cmdProcesar_Click()
If CargaLista = False Then
    If lnTipoProcCH <> gCHTipoProcDesembolso Then
        MsgBox "No se encuentra información solicitada para el proceso", vbInformation, "Aviso"
    Else
        If lnArendirFase = ArendirAtencion Then
            MsgBox "No se encuentran Autorizaciones pendientes de Desembolso", vbInformation, "Aviso"
        Else
            MsgBox "No se encuentran Desembolsos " & IIf(lsTipoDoc = "", " en Efectivo ", "Con Orden de Pago") & " para la Caja Chica Seleccionada ", vbInformation, "Aviso"
        End If
    End If
    If txtBuscarAreaCH.Enabled Then
        Me.txtBuscarAreaCH.SetFocus
    End If
End If
End Sub

Function CargaLista() As Boolean
Dim rs As ADODB.Recordset
Dim lnTipoRend As RendicionTipo
Set rs = New ADODB.Recordset
fgListaCH.Clear
fgListaCH.FormaCabecera
fgListaCH.Rows = 2
txtMovDesc = ""
CargaLista = False
Me.MousePointer = 11
If lnTipoProcCH <> gCHTipoProcDesembolso Then
    Select Case lnArendirFase
    
        Case ArendirRechazo, ArendirAtencion
            Set rs = oCH.GetSolEgresoDirectoPend(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lsCtaFondofijo)
        Case ArendirExtornoAtencion
            Set rs = oCH.GetSolEgresoDirectoAtend(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lsCtaFondofijo)
    End Select
Else
    If lnArendirFase = ArendirAtencion Then
        Set rs = oCH.GetCHAutorizadas(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH))
    Else
        If gsOpeCod = "401333" Then
            Set rs = oCH.GetCHDesembolsadas(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), "19111909", False, 80)
        Else
            '*** PEAC 20100208
            'Set rs = oCH.GetCHDesembolsadas(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lsCtaFondofijo, IIf(nVal(lsTipoDoc) = 0, True, False))
            
            If gsOpeCod = "401332" Then
                If lblNroProcCH > 1 Then
                    Set rs = oCH.GetCHDesembolsadas(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), "251419", IIf(nVal(lsTipoDoc) = 0, True, False))
                Else
                    Set rs = oCH.GetCHDesembolsadas(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lsCtaFondofijo, IIf(nVal(lsTipoDoc) = 0, True, False))
                End If
            '***Agregado por ELRO el 20120604, según OYP-RFC047-2012
            ElseIf gsOpeCod = CStr(gCHExtApropacionApeMN) Or gsOpeCod = CStr(gCHExtApropacionApeME) Then
                If lblNroProcCH = 1 Then
                    Set rs = oCH.devolverCHAprobadas(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lsCtaFondofijo, IIf(nVal(lsTipoDoc) = 0, True, False))
                Else
                    Set rs = oCH.devolverCHAprobadas(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), "251419", IIf(nVal(lsTipoDoc) = 0, True, False))
                End If
            '***Fin Agregado por ELRO*******************************
            Else
                Set rs = oCH.GetCHDesembolsadas(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lsCtaFondofijo, IIf(nVal(lsTipoDoc) = 0, True, False))
            End If
            '*** FIN PEAC
        End If
    End If
End If
If Not rs.EOF And Not rs.BOF Then
    Set fgListaCH.Recordset = rs
    If lnTipoProcCH = gCHTipoProcDesembolso Then
        If lnArendirFase = ArendirAtencion Then fgListaCH.FormatoPersNom (4)
    Else
        fgListaCH.FormatoPersNom (5)
    End If
    Me.fgListaCH.SetFocus
    CargaLista = True
End If
rs.Close
Set rs = Nothing
Me.MousePointer = 0
End Function
Private Sub cmdRechazar_Click()
Dim i As Integer
Dim oCH As nCajaChica
Dim oCon As NContFunciones
Dim lsMovNroRech As String
Dim lsMovNroSol As String


On Error GoTo ErrCmdRechazar

Set oCH = New nCajaChica
Set oCon = New NContFunciones

If fgListaCH.TextMatrix(1, 0) = "" Then Exit Sub
If Val(lblTotal) = 0 Then
    MsgBox "Seleccione alguna solicitud por favor", vbInformation, "Aviso"
    Exit Sub
End If
If Val(lblTotal) = 0 Then
    MsgBox "Seleccione alguna solicitud por favor", vbInformation, "Aviso"
    Exit Sub
End If
If MsgBox("Desea Realizar el Rechazo de las solicitudes seleccionadas??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    For i = 1 To fgListaCH.Rows - 1
        lsMovNroRech = ""
        lsMovNroSol = ""
        If fgListaCH.TextMatrix(i, 1) <> "" Then
            lsMovNroRech = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            lsMovNroSol = fgListaCH.TextMatrix(i, 8)
            If oCH.GrabaRechazoSolEgrDir(lsMovNroRech, lsMovNroSol, gsFormatoFecha, _
                                      gsOpeCod, gsOpeDesc) = 0 Then
                
            End If
        End If
    Next
    CargaLista
    lblSaldo = Format(oCH.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2)), "#,#0.00")
    MsgBox "Rechazo de solicitudes se ha realizado con éxito", vbInformation, "Aviso"
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Rechazo Desembolos de Chica del  Agencia/Area : " & lblCajaChicaDesc & _
            " |Monto : " & lblTotal
            Set objPista = Nothing
            '*******
    CalculaTotal
End If
Set oCH = Nothing
Set oCon = Nothing

Exit Sub
ErrCmdRechazar:
    MsgBox "Error N° [" & Err.Number & "]" & TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub fgListaCH_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtMovDesc.Visible Then
        txtMovDesc.SetFocus
    End If
End If
End Sub

Private Sub fgListaCH_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
CalculaTotal
End Sub
Private Sub Form_Activate()
If lbSalir Then
    Unload Me
End If
End Sub
Private Sub Form_Load()
Set oArendir = New NARendir
Set oCH = New nCajaChica
Dim oOpe As DOperacion
CentraForm Me
lbSalir = False
Me.Caption = gsOpeDesc
Set oOpe = New DOperacion
Dim cCodCarJefCon As String 'Agregado por ELRO el 20120529, según OYP-RFC047-2012

cmdAtender.Visible = False
cmdRechazar.Visible = False
cmdExtorno.Visible = False
cmdDesembolsar.Visible = False
cmdExtDesemb.Visible = False
Me.txtMovDesc.Visible = False

'***Agregado por ELRO el 20120808, según OYP-RFC047-2012
cCodCarJefCon = oArendir.devolverCodigoJefeContabilidad
'***Fin Agregado por ELRO*****************************

lsTipoDoc = "0"
If lnTipoProcCH <> gCHTipoProcDesembolso Then
    Select Case lnArendirFase
        Case ArendirRechazo
            cmdRechazar.Visible = True

            '*** PEAC 20100208
            If gsOpeCod = "401372" Then
                lsCtaFondofijo = oOpe.EmiteOpeCta(gsOpeCod, "D", "1")
                'lsCtaFondofijo = oOpe.EmiteOpeCta(gsOpeCod, "D")
            Else
                lsCtaFondofijo = oOpe.EmiteOpeCta(gsOpeCod, "D")
            End If
            '***FIN PEAC
            
            If lsCtaFondofijo = "" Then
                MsgBox "No se han definido correctamente las cuentas Contables de Operacion", vbInformation, "Aviso"
                lbSalir = True
                Exit Sub
            End If
        Case ArendirAtencion
            cmdAtender.Visible = True
            cmdDocumento.Visible = True
            
            '*** PEAC 20100129
            If gsOpeCod = "401373" Then
                lsCtaFondofijo = oOpe.EmiteOpeCta(gsOpeCod, "D", "1")
                'lsCtaFondofijo = oOpe.EmiteOpeCta(gsOpeCod, "D")
            Else
                lsCtaFondofijo = oOpe.EmiteOpeCta(gsOpeCod, "D")
            End If
            '*** FIN PEAC
            
            If lsCtaFondofijo = "" Then
                MsgBox "No se han definido correctamente las cuentas Contables de Operacion", vbInformation, "Aviso"
                lbSalir = True
                Exit Sub
            End If
        Case ArendirExtornoAtencion
            cmdExtorno.Visible = True
            '*** PEAC 20100208
            If gsOpeCod = "401374" Then
                lsCtaFondofijo = oOpe.EmiteOpeCta(gsOpeCod, "D", "1")
                'lsCtaFondofijo = oOpe.EmiteOpeCta(gsOpeCod, "D")
            Else
                lsCtaFondofijo = oOpe.EmiteOpeCta(gsOpeCod, "D")
            End If
            '*** FIN PEAC
            If lsCtaFondofijo = "" Then
                MsgBox "No se han definido correctamente las cuentas Contables de Operacion", vbInformation, "Aviso"
                lbSalir = True
                Exit Sub
            End If
    End Select
Else
    If lnArendirFase = ArendirAtencion Then
        Me.cmdDesembolsar.Visible = True
        Me.txtMovDesc.Visible = True
        Me.fgListaCH.EncabezadosNombres = "N°-Ok-Fecha Aut-Aut.Por-Encargado-ImporteAux-Importe-cMovNroAut-cPersCod"
        Me.fgListaCH.EncabezadosAnchos = "450-400-1200-900-4500-1200-0-0"
        Me.fgListaCH.EncabezadosAlineacion = "C-C-C-C-L-R-C-C-C"
        Me.fgListaCH.FormatosEdit = "0-0-0-0-0-2-2-0-0-0-0"
        lsCtaFondofijo = oOpe.EmiteOpeCta(gsOpeCod, "D")
        lsCtaDesembolso = oOpe.EmiteOpeCta(gsOpeCod, "H")
        lsTipoDoc = oOpe.EmiteDocOpe(gsOpeCod, OpeDocEstObligatorioDebeExistir, OpeDocMetAutogenerado)
    Else
        '***Agregado por ELRO el 20120808, según OYP-RFC047-2012
        If gsCodCargo <> cCodCarJefCon Then
            MsgBox "Solo el(la) Jefe(a) de Contabilidad puede realizar el extorno.", vbInformation, "Aviso"
            Exit Sub
        End If
        '***Fin Agregado por ELRO*******************************
        Me.fgListaCH.EncabezadosNombres = "N°-Ok-Tipo-Número-Fecha-Concepto-MontoDesemb-Importe-Saldo-cMovNroDesemb"
        Me.fgListaCH.EncabezadosAnchos = "450-400-600-1200-1000-4500-0-1200-0-0"
        Me.fgListaCH.EncabezadosAlineacion = "C-C-C-C-C-L-R-R-C-C-C"
        Me.fgListaCH.FormatosEdit = "0-0-0-0-0-2-2-2-0-0-0"
        cmdExtDesemb.Visible = True
        Me.txtMovDesc.Visible = True
        
        lsCtaFondofijo = oOpe.EmiteOpeCta(gsOpeCod, "D")
        
        lsTipoDoc = oOpe.EmiteDocOpe(gsOpeCod, OpeDocEstObligatorioDebeExistir, OpeDocMetAutogenerado)
    End If
End If
CentraForm Me
txtBuscarAreaCH.psRaiz = "CAJAS CHICAS"
txtBuscarAreaCH.rs = oArendir.EmiteCajasChicas
Set oOpe = Nothing
End Sub


Private Sub txtBuscarAreaCH_EmiteDatos()
Dim oCajaCH As nCajaChica
Set oCajaCH = New nCajaChica
lblCajaChicaDesc = txtBuscarAreaCH.psDescripcion
lblNroProcCH = oCajaCH.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), NroCajaChica)
fgListaCH.Clear
fgListaCH.FormaCabecera
fgListaCH.Rows = 2
If oCajaCH.GetMovCHProceso(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), gCHTipoProcRendicion) <> "" Then
    MsgBox "Caja Chica se encuentra en proceso de rendicion. Consulte con Sistemas", vbInformation, "Aviso"
    txtBuscarAreaCH = ""
    lblNroProcCH = ""
    lblCajaChicaDesc = ""
    lblTotal = "0.00"
    lblSaldo = "0.00"
    Exit Sub
End If
If lblCajaChicaDesc <> "" Then
    lblSaldo = Format(oCajaCH.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2)), "#,#0.00")
    If cmdProcesar.Visible Then
        cmdProcesar.SetFocus
    End If
End If
Set oCajaCH = Nothing
End Sub


Private Sub txtBuscarAreaCH_Validate(Cancel As Boolean)
If txtBuscarAreaCH = "" Then
    Cancel = True
End If
End Sub
Sub CalculaTotal()
Dim lnTotal As Currency
Dim i As Integer
Dim oCajaChica As nCajaChica
Set oCajaChica = New nCajaChica
Dim lnTope As Currency

lnTotal = 0
If fgListaCH.TextMatrix(1, 0) <> "" Then
    For i = 1 To fgListaCH.Rows - 1
        If fgListaCH.TextMatrix(i, 1) <> "" Then
            lnTotal = lnTotal + CCur(IIf(fgListaCH.TextMatrix(i, 6) = "", "0", fgListaCH.TextMatrix(i, 6)))
        End If
    Next i
End If
lblTotal = Format(lnTotal, gsFormatoNumeroView)
'--------John--------
'If gsOpeCod = 401373 Then
'    lnTope = oCajaChica.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), MontoTope)
'    If lblTotal > lnTope Then
'       MsgBox "Monto Total de Documentos es mayor a Tope de Caja Chica" & oImpresora.gPrnSaltoLinea & "Debe Rechazar Documentos antes de Atender", vbInformation, "Aviso"
'       cmdAtender.Enabled = False
'    Else
'       cmdAtender.Enabled = True
'    End If
'End If
'-------------------
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If lnTipoProcCH = gCHTipoProcDesembolso Then
        If lnArendirFase = ArendirAtencion Then
            cmdDesembolsar.SetFocus
        Else
            Me.cmdExtDesemb.SetFocus
        End If
    End If
End If
End Sub
