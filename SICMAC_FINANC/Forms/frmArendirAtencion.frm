VERSION 5.00
Begin VB.Form frmArendirAtencion 
   Caption         =   "Operaciones: "
   ClientHeight    =   5250
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10875
   Icon            =   "frmArendirAtencion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   10875
   Begin VB.PictureBox vFormPagoAge 
      Height          =   5145
      Left            =   60
      ScaleHeight     =   5085
      ScaleWidth      =   960
      TabIndex        =   15
      Top             =   60
      Width           =   1020
   End
   Begin VB.TextBox txtMovDesc 
      Height          =   465
      Left            =   1125
      Locked          =   -1  'True
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3045
      Width           =   9690
   End
   Begin VB.TextBox txtMovDescAt 
      Height          =   690
      Left            =   1140
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3780
      Width           =   9690
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   9495
      TabIndex        =   6
      Top             =   4785
      Width           =   1215
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "&Rechazar"
      Height          =   360
      Left            =   8280
      TabIndex        =   5
      Top             =   4785
      Width           =   1215
   End
   Begin VB.CommandButton cmdDoc 
      Caption         =   "&Emitir"
      Enabled         =   0   'False
      Height          =   360
      Left            =   8280
      TabIndex        =   8
      Top             =   4785
      Width           =   1215
   End
   Begin VB.PictureBox vFormPago 
      Height          =   5145
      Left            =   60
      ScaleHeight     =   5085
      ScaleWidth      =   960
      TabIndex        =   7
      Top             =   60
      Width           =   1020
   End
   Begin Sicmact.FlexEdit FlexARendir 
      Height          =   2340
      Left            =   1125
      TabIndex        =   1
      Top             =   645
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   4128
      Cols0           =   13
      HighLight       =   2
      AllowUserResizing=   1
      EncabezadosNombres=   "Nº-Número-Area-Agencia-Persona Solicitante-Fecha-Importe-Concepto-cCodArea-cPersCod-nMovNro-cDocTpo-cCodAge"
      EncabezadosAnchos=   "300-1100-1800-1800-2400-900-1100-0-0-0-0-0-0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L-L-C-R-C-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-2-0-0-0-0-0-0"
      TextArray0      =   "Nº"
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
   End
   Begin VB.Frame fraEntidad 
      Caption         =   "Entidad Pagadora"
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
      Height          =   660
      Left            =   1125
      TabIndex        =   9
      Top             =   4530
      Width           =   6960
      Begin Sicmact.TxtBuscar txtBuscaEntidad 
         Height          =   360
         Left            =   105
         TabIndex        =   4
         Top             =   210
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   635
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCtaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2610
         TabIndex        =   10
         Top             =   210
         Width           =   4245
      End
   End
   Begin VB.Frame fraArendir 
      Height          =   615
      Left            =   1125
      TabIndex        =   12
      Top             =   -15
      Width           =   9660
      Begin Sicmact.TxtBuscar txtBuscarArendir 
         Height          =   345
         Left            =   1140
         TabIndex        =   0
         Top             =   188
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         ForeColor       =   -2147483647
      End
      Begin VB.Label lblDescArendir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   315
         Left            =   2325
         TabIndex        =   14
         Top             =   210
         Width           =   5715
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "A Rendir de :"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   240
         Width           =   930
      End
   End
   Begin VB.Label lblConcepto 
      AutoSize        =   -1  'True
      Caption         =   "Concepto de Operación :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1125
      TabIndex        =   11
      Top             =   3570
      Width           =   2145
   End
End
Attribute VB_Name = "frmArendirAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**************************************************************************************
'*************************************OBJETOS **********************************
Dim oDCtaIF As DCajaCtasIF
Dim oNArendir As NARendir
Dim oNContFunc As NContFunciones
Dim oOpe As DOperacion

Dim objPista As COMManejador.Pista
'************************************************************************************
Dim lSalir As Boolean
Dim lsMoney As String
Dim lsSimbolo As String
Dim lbMN As Boolean
Dim lsDocTpo As String
Dim lsDocTpoVoucher  As String
Dim lsDocVoucher  As String
Dim lsEntidad As String
Dim lsFecha As String
Dim lnTipoArendir  As ArendirTipo
Dim lsFileCarta As String
Dim lsDocNro As String
Dim lnArendirFase As ARendirFases
Public Sub Inicio(pnTipoArendir As ArendirTipo, ByVal pnArendirFase As ARendirFases)
lnArendirFase = pnArendirFase
lnTipoArendir = pnTipoArendir
Me.Show 1
End Sub
Private Function ValidaInterfaz() As Boolean
ValidaInterfaz = True
If Len(Trim(txtMovDescAt)) = 0 Then
    MsgBox "Glosa o descripcion de Operación no válida", vbInformation, "Aviso"
    txtMovDescAt.SetFocus
    ValidaInterfaz = False
    Exit Function
End If
If fraEntidad.Visible Then
    If txtBuscaEntidad.Text = "" Then
        MsgBox "Cuenta de Institución Financiera no válida", vbInformation, "Aviso"
        ValidaInterfaz = False
        Exit Function
    End If
End If
End Function

Private Sub cmdDoc_Click()
Dim lsEntidadOrig As String
Dim lsCtaEntidadOrig As String
Dim lsGlosa As String
Dim lsPersNombre As String
Dim lsPersDireccion As String
Dim lsUbigeo    As String
Dim lsCuentaAho As String

Dim lnImporte As Currency
Dim oDocPago As clsDocPago
Dim lsSubCuentaIF As String
Dim lsPersCod As String
Dim lsMovNro As String
Dim lsDocumento As String
Dim lsOpeCod As String
Dim lsCtaBanco As String
Dim lsCtaContDebe As String
Dim lsCtaContHaber As String
Dim lsPersCodIf As String
Dim lsMovAnt As String
Dim lsCtaContHaberGen As String
Dim lnTrans As Integer
Dim rsBilletaje As ADODB.Recordset
Dim lsOPSave As String
Dim lsAreaCod As String
Dim lsAreaDesc As String
Dim lsAgeCod As String
Dim lsAgeDesc As String
Dim lbEfectivo As Boolean
Dim oCtasIF As NCajaCtaIF
Dim oOpe As DOperacion
Dim lsNroDocViatico As String
Dim lsTpoIf As String
Dim lsCadBol As String

On Error GoTo NoGrabo

Set oCtasIF = New NCajaCtaIF
Set oOpe = New DOperacion
If ValidaInterfaz = False Then Exit Sub
Set oDocPago = New clsDocPago

lsCtaEntidadOrig = Trim(lblCtaDesc)
lsTpoIf = Mid(txtBuscaEntidad, 1, 2)
lsCtaBanco = Mid(txtBuscaEntidad, 18, Len(Me.txtBuscaEntidad))
lsPersCodIf = Mid(txtBuscaEntidad, 4, 13)
lsEntidadOrig = oDCtaIF.NombreIF(lsPersCodIf)
lsSubCuentaIF = oCtasIF.SubCuentaIF(lsPersCodIf)
lsGlosa = Trim(txtMovDescAt)
lsPersNombre = FlexARendir.TextMatrix(FlexARendir.row, 4)
lsPersCod = FlexARendir.TextMatrix(FlexARendir.row, 9)
lsAreaCod = FlexARendir.TextMatrix(FlexARendir.row, 8)
lsAreaDesc = FlexARendir.TextMatrix(FlexARendir.row, 2)
lsAgeCod = FlexARendir.TextMatrix(FlexARendir.row, 12)
lsAgeDesc = FlexARendir.TextMatrix(FlexARendir.row, 3)

If lsDocTpo = "" Then
    MsgBox "Seleccione Forma de pago para el Arendir siguiente", vbInformation, "Aviso"
    Exit Sub
End If
If lnTipoArendir = gArendirTipoViaticos Then
    lsNroDocViatico = FlexARendir.TextMatrix(FlexARendir.row, 1)
Else
    lsNroDocViatico = ""
End If

lnImporte = CCur(FlexARendir.TextMatrix(FlexARendir.row, 6))
lsMovAnt = FlexARendir.TextMatrix(FlexARendir.row, 10)
lsDocVoucher = ""
lsDocNro = ""
lbEfectivo = False
If lsDocTpo = "-1" Then
    If lnTipoArendir = gArendirTipoAgencias Then
        MsgBox "Atención con Efectivo sólo puede realizarlo Caja General", vbInformation
        Exit Sub
    End If
    frmArendirEfectivo.Inicio lnTipoArendir, FlexARendir.TextMatrix(FlexARendir.row, 1), Mid(gsOpeCod, 3, 1), IIf(FlexARendir.TextMatrix(FlexARendir.row, 3) = "", FlexARendir.TextMatrix(FlexARendir.row, 2), FlexARendir.TextMatrix(FlexARendir.row, 3)), lnImporte, lsPersCod, lsPersNombre
    
    Set rsBilletaje = frmArendirEfectivo.rsEfectivo
    Set frmArendirEfectivo = Nothing
    Unload frmArendirEfectivo
    If rsBilletaje Is Nothing Then
        Exit Sub
    End If
    lbEfectivo = True
ElseIf lsDocTpo = TpoDocNotaAbono Then
    Dim oImp As New NContImprimir
    lsDocTpo = TpoDocNotaAbono
    
    frmNotaCargoAbono.Inicio lsDocTpo, lnImporte, gdFecSis, txtMovDesc, gsOpeCod, False, lsPersCod, lsPersNombre
    If frmNotaCargoAbono.vbOk Then
    
        lsDocNro = frmNotaCargoAbono.NroNotaCA
        txtMovDesc = frmNotaCargoAbono.Glosa
        lsDocumento = frmNotaCargoAbono.NotaCargoAbono
        lsPersNombre = frmNotaCargoAbono.PersNombre
        lsPersDireccion = frmNotaCargoAbono.PersDireccion
        lsUbigeo = frmNotaCargoAbono.PersUbigeo
        lsCuentaAho = frmNotaCargoAbono.CuentaAhoNro
        lsFecha = frmNotaCargoAbono.FechaNotaCA
'        lsDocumento = oImp.ImprimeNotaCargoAbono(lsDocNRo, txtMovDesc, CCur(frmNotaCargoAbono.Monto), _
'                            lsPersNombre, lsPersDireccion, lsUbigeo, gdFecSis, Mid(gsOpeCod, 3, 1), lsCuentaAho, lsDocTpo, gsNomAge, gsCodUser)
        lsDocumento = oImp.ImprimeNotaAbono(lsFecha, lnImporte, txtMovDesc, lsCuentaAho, lsPersNombre)
        Dim oDis As New NRHProcesosCierre
        lsCadBol = oDis.ImprimeBoletaCad(CDate(lsFecha), "ABONO CAJA GENERAL", "Depósito CAJA GENERAL*Nro." & lsDocNro, "", lnImporte, lsPersNombre, lsCuentaAho, "", 0, 0, "Nota Abono", 0, 0, False, False, , , , True, , , , False, gsNomAge) & oImpresora.gPrnSaltoPagina
    Else
        Exit Sub
    End If
Else
    If lsDocTpo = TpoDocCheque Then
        If lnTipoArendir = gArendirTipoAgencias Then
            MsgBox "Atención con Cheque sólo puede realizarlo Caja General", vbInformation
            Exit Sub
        End If
        lsDocVoucher = oNContFunc.GeneraDocNro(TpoDocVoucherEgreso, , Mid(gsOpeCod, 3, 1), Right(gsCodAge, 2))
        
        'oDocPago.InicioCheque lsDocNRo, True, lsPersCod, gsOpeCod, lsPersNombre, gsOpeCod, lsGlosa, lnImporte, gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lsEntidadOrig, lsCtaEntidadOrig, lsDocVoucher, False, , lsCtaBanco
        oDocPago.InicioCheque lsDocNro, True, lsPersCod, gsOpeCod, lsPersNombre, gsOpeCod, lsGlosa, lnImporte, gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lsEntidadOrig, lsCtaEntidadOrig, lsDocVoucher, False, , lsCtaBanco, , lsTpoIf, lsPersCodIf, lsCtaBanco 'EJVG20121130
 
    End If
    If lsDocTpo = TpoDocOrdenPago Then
        Screen.MousePointer = 11
        lsDocVoucher = oNContFunc.GeneraDocNro(TpoDocVoucherEgreso, , Mid(gsOpeCod, 3, 1), Right(gsCodAge, 2))
        oDocPago.InicioOrdenPago lsDocNro, True, lsPersCod, gsOpeCod, lsPersNombre, gsOpeCod, lsGlosa, lnImporte, gdFecSis, lsDocVoucher, False, Right(gsCodAge, 2)
        Screen.MousePointer = 0
    End If
    If lsDocTpo = TpoDocCarta Then
        If lnTipoArendir = gArendirTipoAgencias Then
            MsgBox "Atención con Carta sólo puede realizarlo Caja General", vbInformation
            Exit Sub
        End If
       oDocPago.InicioCarta lsDocNro, lsPersCod, gsOpeCod, gsOpeCod, lsGlosa, lsFileCarta, lnImporte, gdFecSis, lsEntidadOrig, lsCtaEntidadOrig, lsPersNombre, "", lsMovNro
    End If
    If oDocPago.vbOk Then    'Se ingresó dato de Cheque u Orden de Pago
       txtMovDescAt = oDocPago.vsGlosa
       lsOpeCod = gsOpeCod
       lsFecha = oDocPago.vdFechaDoc
       lsDocTpo = oDocPago.vsTpoDoc
       lsDocNro = oDocPago.vsNroDoc
       lsDocVoucher = oDocPago.vsNroVoucher
       lsDocumento = oDocPago.vsFormaDoc
    Else
        Exit Sub
    End If
End If
lsOpeCod = oOpe.EmiteOpeDoc(Mid(gsOpeCod, 1, 5), Val(lsDocTpo))
If lsOpeCod = "" Then
    MsgBox "No se asignó Documentos de Referencia a Operación de Atención", vbInformation, "Aviso"
    Exit Sub
End If
lsCtaContHaber = ""
lsCtaContDebe = ""

lsCtaContDebe = oOpe.EmiteOpeCta(gsOpeCod, "D", , txtBuscaEntidad, ObjEntidadesFinancieras)
If lsDocTpo = TpoDocOrdenPago Then
    lsCtaContHaber = oOpe.EmiteOpeCta(lsOpeCod, "H", , TxtBuscarArendir, ObjCMACAgenciaArea)
Else
    lsCtaContHaber = oOpe.EmiteOpeCta(lsOpeCod, "H", , txtBuscaEntidad, ObjEntidadesFinancieras)
End If

'lsCtaContHaberGen = oOpe.EmiteOpeCta(lsOpeCod, "H", , txtBuscaEntidad, CtaOBjFiltroIF, False)
If lsCtaContDebe = "" Or lsCtaContHaber = "" Then
    MsgBox "Cuentas Contables no determinadas correctamente." & oImpresora.gPrnSaltoLinea & "consulte con sistemas", vbInformation, "Aviso"
    Exit Sub
End If

Dim lsCtaITFD As String
Dim lsCtaITFH As String

lsCtaITFD = oOpe.EmiteOpeCta(gsOpeCod, "D", 2)
lsCtaITFH = oOpe.EmiteOpeCta(gsOpeCod, "H", 2)

If lsCtaITFD = "" Or lsCtaITFH = "" Then
    MsgBox "Cuentas Contables ITF no determinadas correctamente." & oImpresora.gPrnSaltoLinea & "consulte con sistemas", vbInformation, "Aviso"
    Exit Sub
End If

'If gsOpeCod = 401130 Or gsOpeCod = 401230 Then
'    Dim oConect As DConecta
'    Set oConect = New DConecta
'    Dim subCta As String
'    Dim sql1 As String
'    Dim rs2 As ADODB.Recordset
'    'Set rs2 = New ADODB.Recordset
'    Dim codarea As String
'    codarea = gsCodArea
'
'    If oConect.AbreConexion = False Then MsgBox "No hay conexion a la Base de Datos"
'    sql1 = "Select cSubCtacod FROM Areas where cAreaCod = '" & codarea & "'"
'    Set rs2 = oConect.CargaRecordSet(sql1)
'    subCta = rs2!cSubctacod
'    lsCtaITFD = oOpe.EmiteOpeCta(gsOpeCod, "D", 2) & subCta
'    lsCtaITFH = oOpe.EmiteOpeCta(gsOpeCod, "H", 2) & subCta
'End If


If MsgBox("Desea Grabar la Información", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsMovNro = oNContFunc.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    
    If oNArendir.GrabaAtencionArendir(lnTipoArendir, lsMovNro, lsOpeCod, txtMovDescAt, lsCtaContDebe, _
                                    lsCtaContHaber, lsPersCod, lnImporte, lsTpoIf, lsPersCodIf, lsCtaBanco, _
                                    rsBilletaje, lsDocTpo, lsDocNro, lsFecha, lsDocTpoVoucher, lsDocVoucher, lsMovAnt, lsCuentaAho, gbBitCentral, , lsCtaITFD, lsCtaITFH, gnImpITF) = 0 Then
        
        If lsOpeCod = "401132" Or lsOpeCod = "402132" Or lsOpeCod = "401232" Or lsOpeCod = "402232" Then
           'ImprimeAsientoContableNew lsMovNro, lsDocVoucher, lsDocTpo, lsDocumento, lbEfectivo, _
           '                     False, txtMovDescAt, lsPersCod, lnImporte, lnTipoArendir, lsNroDocViatico, , , , "17", , , lsCadBol, Mid(lsOpeCod, 3, 1)
           ImprimeAsientoContableUltimo lsMovNro, lsDocVoucher, lsDocTpo, lsDocumento, lbEfectivo, False, txtMovDesc, lsPersCod, lnImporte, , , , 1, , "17", , , lsCadBol, Mid(lsOpeCod, 3, 1)
                                
        Else
           ImprimeAsientoContable lsMovNro, lsDocVoucher, lsDocTpo, lsDocumento, lbEfectivo, _
                                False, txtMovDescAt, lsPersCod, lnImporte, lnTipoArendir, lsNroDocViatico, , , , "17", , , lsCadBol
        End If
        objPista.InsertarPista lsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Atencion de A Rendir o A Rendir Viaticos"
        FlexARendir.EliminaFila FlexARendir.row
        If FlexARendir.TextMatrix(1, 0) = "" Then
            Unload Me
            Exit Sub
        End If
        txtMovDesc = ""
        txtMovDescAt = ""
        lsDocTpo = ""
        lsDocNro = ""
        lsDocVoucher = ""
        lsDocumento = ""
        txtBuscaEntidad = ""
        lblCtaDesc = ""
    End If
End If
   
Exit Sub
NoGrabo:
  MsgBox TextErr(Err.Description), vbInformation, "Error de Actualización"
End Sub

Private Sub cmdRechazar_Click()
On Error GoTo ErrCmdRechazar
Dim lsMovNro As String
If FlexARendir.TextMatrix(1, 1) = "" Then
    MsgBox "No existen Solicitudes para Rechazar...!", vbInformation, "Aviso"
    Exit Sub
End If
If Len(Trim(txtMovDescAt)) = 0 Then
    MsgBox "Descripcion de la Operación no Válida", vbInformation, "Aviso"
    txtMovDescAt.SetFocus
    Exit Sub
End If
If MsgBox(" ¿ Seguro de Rechazar A Rendir N° :" & FlexARendir.TextMatrix(FlexARendir.row, 1) & vbCrLf & vbCrLf & "Solicitado por :" & FlexARendir.TextMatrix(FlexARendir.row, 4) & vbCrLf & "De :" & FlexARendir.TextMatrix(FlexARendir.row, 2) & " - " & FlexARendir.TextMatrix(FlexARendir.row, 3) & vbCrLf, vbQuestion + vbYesNo, "Confirmación") = vbYes Then
    'Actualizamos el Estado del Recibo de Egresos a Rechazado
    lsMovNro = oNContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    If oNArendir.GrabaRechazoSolARendir(lsMovNro, gsOpeCod, FlexARendir.TextMatrix(FlexARendir.row, 10), Trim(txtMovDescAt)) = 0 Then
        'Eliminado del ListView el Recibo Rechazado
        'ARLO 201710214
        objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Rechazo a Rendir Cuentas"
        '******
        FlexARendir.EliminaFila FlexARendir.row
        FlexARendir.SetFocus
        Me.txtMovDesc = ""
        Me.txtMovDescAt = ""
    End If
End If
Exit Sub
ErrCmdRechazar:
    MsgBox "Error N° [" & Err.Number & "]" & TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub FlexARendir_Click()
RefrescaDatosArendir
End Sub

Private Sub RefrescaDatosArendir()
If FlexARendir.TextMatrix(1, 1) <> "" Then
    If lnArendirFase = ArendirAtencion Then
       txtMovDescAt = FlexARendir.TextMatrix(FlexARendir.row, 7)
    End If
    txtMovDesc = FlexARendir.TextMatrix(FlexARendir.row, 7)
Else
    txtMovDescAt = ""
    txtMovDesc = ""
End If
End Sub
Private Sub FlexARendir_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtMovDescAt.Visible And txtMovDescAt.Enabled Then
        txtMovDescAt.SetFocus
    Else
        txtMovDesc.SetFocus
    End If
End If
End Sub

Private Sub FlexARendir_OnRowChange(pnRow As Long, pnCol As Long)
RefrescaDatosArendir
End Sub
Private Sub FlexARendir_RowColChange()
RefrescaDatosArendir
End Sub

Private Sub Form_Activate()
If lSalir Then
   Unload Me
End If
End Sub

Private Sub Form_Load()
Dim lvItem As ListItem
Dim rsPer As New ADODB.Recordset
Dim oAreas As DActualizaDatosArea
Dim oOpe As DOperacion
Set oDCtaIF = New DCajaCtasIF
Set oAreas = New DActualizaDatosArea
Set objPista = New COMManejador.Pista

CentraForm Me
lsFileCarta = App.path & gsDirPlantillas & gsOpeCod & ".TXT"
Set oOpe = New DOperacion
txtBuscaEntidad.rs = oOpe.GetRsOpeObj(gsOpeCod, "1") '  oDCtaIF.CargaCtasIF(Mid(gsOpeCod, 3, 1), gTpoIFBanco + gTpoCtaIFCtaCte + gTpoCtaIFCtaAho)
Set oOpe = Nothing
Me.Caption = Me.Caption & gsOpeDesc
'AbreConexion
If Mid(gsOpeCod, 3, 1) = gMonedaNacional Then  'Identificación de Tipo de Moneda
   lbMN = True
   lsMoney = gcMN
   lsSimbolo = gcMN
Else
   lbMN = False
   lsMoney = gcME
   lsSimbolo = gcME
   If gnTipCambio = 0 Then
      If Not GetTipCambio(gdFecSis) Then
         lSalir = True
         Exit Sub
      End If
   End If
End If
Set oNContFunc = New NContFunciones
Set oNArendir = New NARendir
Set oOpe = New DOperacion
TxtBuscarArendir.psRaiz = "A Rendir de..."
'txtBuscarArendir.rs = oAreas.GetAgenciasAreas(, 1)
If lnTipoArendir = gArendirTipoAgencias Then
    Set rsPer = oOpe.CargaOpeObj(gCGArendirCtaSolMNAge, 1)
    vFormPago.Visible = False
    vFormPagoAge.Visible = True
Else
    Set rsPer = oOpe.CargaOpeObj(gCGArendirCtaSolMN, 1)
    vFormPago.Visible = True
    vFormPagoAge.Visible = False
End If
If Not rsPer.EOF Then
    TxtBuscarArendir.rs = oAreas.GetAgenciasAreas(rsPer!cOpeObjFiltro, 1)
End If

lsDocTpo = oOpe.EmiteDocOpe(gsOpeCod, OpeDocEstObligatorioDebeExistir, OpeDocMetAutogenerado)
lsDocTpoVoucher = oOpe.EmiteDocOpe(gsOpeCod, OpeDocEstOpcionalNoDebeExistir, OpeDocMetAutogenerado)
    
If lnArendirFase = ArendirRechazo Then
    lblConcepto = "Motivo de Rechazo: "
    Me.cmdDoc.Visible = False
    Me.fraEntidad.Visible = False
    cmdRechazar.Visible = True
    Me.fraArendir.Left = 100
    FlexARendir.Left = 100
    txtMovDesc.Left = 100
    txtMovDescAt.Left = 100
    lblConcepto.Left = 100
    cmdDoc.Left = 7300
    cmdRechazar.Left = 7300
    cmdSalir.Left = 8550
    vFormPago.Visible = False
    vFormPagoAge.Visible = False
    Me.Width = 10500
Else
    cmdRechazar.Visible = False
    cmdDoc.Visible = True
    fraEntidad.Visible = False
    vFormPago.Visible = True
    cmdDoc.Enabled = True ' arlo
End If
lsDocTpo = "-1"
Set oAreas = Nothing
End Sub
Private Sub CargaPendientes()
Dim lsArea As String
Dim lsDescArea As String
Dim rs As ADODB.Recordset
Dim N As Integer
Dim lvItem As ListItem
FlexARendir.Clear
FlexARendir.FormaCabecera
FlexARendir.Rows = 2
Set rs = New ADODB.Recordset
Set rs = oNArendir.ARendirPendientes(lnTipoArendir, TpoDocRecArendirCuenta, Mid(gsOpeCod, 3, 1), Mid(TxtBuscarArendir, 1, 3), Mid(TxtBuscarArendir, 4, 2))
lSalir = False
If rs.EOF And rs.BOF Then
   MsgBox "No existe Solicitudes Pendientes", vbInformation, "Aviso"
   rs.Close
   Set rs = Nothing
   RefrescaDatosArendir
   Exit Sub
End If
If Not rs.EOF And Not rs.BOF Then
    Set FlexARendir.Recordset = rs
    FlexARendir.FormatoPersNom 4
    If FlexARendir.Visible Then
        FlexARendir.SetFocus
    End If
End If
FlexARendir.FormateaColumnas
RSClose rs
RefrescaDatosArendir

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oDCtaIF = Nothing
Set oNArendir = Nothing
Set oNContFunc = Nothing
End Sub


Private Sub txtBuscaEntidad_EmiteDatos()
Dim oNCtasIf As NCajaCtaIF
Set oNCtasIf = New NCajaCtaIF
If txtBuscaEntidad.Text <> "" Then
    lblCtaDesc = oNCtasIf.EmiteTipoCuentaIF(Mid(Me.txtBuscaEntidad.Text, 18, Len(txtBuscaEntidad.Text))) & " " & txtBuscaEntidad.psDescripcion
    Set oNCtasIf = Nothing
    cmdDoc.Enabled = True
    cmdDoc.SetFocus
    cmdDoc_Click
End If
End Sub

Private Sub TxtBuscarArendir_EmiteDatos()
lblDescArendir = Trim(TxtBuscarArendir.psDescripcion)
CargaPendientes
End Sub

Private Sub txtMovDescAt_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras(KeyAscii)
If KeyAscii = 13 Then
    KeyAscii = 0
    If cmdRechazar.Visible Then
        cmdRechazar.SetFocus
    Else
        If cmdDoc.Visible And cmdDoc.Enabled Then
            cmdDoc.SetFocus
        End If
    End If
End If
End Sub
Private Sub vFormPago_MenuItemClick(MenuNumber As Long, MenuItem As Long)
fraEntidad.Visible = False
If TxtBuscarArendir = "" Then
    MsgBox "Seleccione a quien solicito el Arendir", vbInformation, "Aviso"
    TxtBuscarArendir.SetFocus
    Exit Sub
End If
If FlexARendir.TextMatrix(1, 0) = "" Then
    MsgBox "Solicitudes no encontradas por favor vuelva a procesar...", vbInformation, "Aviso"
    Exit Sub
End If
txtBuscaEntidad.Text = ""
lblCtaDesc = ""
Select Case MenuItem
    Case 1: fraEntidad.Visible = False
           cmdDoc.Enabled = True
           lsDocTpo = "-1"
           cmdDoc_Click
    Case 2:
           fraEntidad.Visible = True
           cmdDoc.Enabled = False
           lsDocTpo = TpoDocCarta  ' TpoDocCarta
   Case 3:
           cmdDoc.Enabled = True
           fraEntidad.Visible = False
           lsDocTpo = TpoDocOrdenPago  ' TpoDocOrdenPago
           cmdDoc_Click
   Case 4: 'Cheque
            fraEntidad.Visible = True
            lsDocTpo = TpoDocCheque
            cmdDoc.Enabled = False
   Case 5:  'Nota de Abono
            fraEntidad.Visible = False
            lsDocTpo = TpoDocNotaAbono
            cmdDoc_Click
End Select
End Sub


Private Sub vFormPagoAge_MenuItemClick(MenuNumber As Long, MenuItem As Long)
fraEntidad.Visible = False
If TxtBuscarArendir = "" Then
    MsgBox "Seleccione a quien solicito el Arendir", vbInformation, "Aviso"
    TxtBuscarArendir.SetFocus
    Exit Sub
End If
If FlexARendir.TextMatrix(1, 0) = "" Then
    MsgBox "Solicitudes no encontradas por favor vuelva a procesar...", vbInformation, "Aviso"
    Exit Sub
End If
txtBuscaEntidad.Text = ""
lblCtaDesc = ""
Select Case MenuItem
   Case 1:
           cmdDoc.Enabled = True
           fraEntidad.Visible = False
           lsDocTpo = TpoDocOrdenPago  ' TpoDocOrdenPago
           cmdDoc_Click
End Select
End Sub
