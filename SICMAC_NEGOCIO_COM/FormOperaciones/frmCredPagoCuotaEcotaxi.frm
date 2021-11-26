VERSION 5.00
Begin VB.Form frmCredPagoCuotaEcotaxi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAGO NORMAL CUOTAS ECOTAXI"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   11970
   Icon            =   "frmCredPagoCuotaEcotaxi.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   11970
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "2. Pagar Cuotas"
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
      Left            =   2240
      TabIndex        =   6
      Top             =   3585
      Width           =   2145
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
      Left            =   10740
      TabIndex        =   5
      Top             =   3585
      Width           =   1185
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "1.  Buscar Créditos >"
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
      Left            =   60
      MaskColor       =   &H00C0FFFF&
      TabIndex        =   4
      Top             =   3585
      Width           =   2145
   End
   Begin VB.Frame Frame1 
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
      Height          =   3600
      Left            =   50
      TabIndex        =   0
      Top             =   -40
      Width           =   11895
      Begin VB.OptionButton OptSelec 
         Caption         =   "&Todos"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   180
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton OptSelec 
         Caption         =   "&Ninguno"
         Height          =   255
         Index           =   1
         Left            =   1005
         TabIndex        =   8
         Top             =   180
         Width           =   1065
      End
      Begin SICMACT.FlexEdit feCreditosEcotaxi 
         Height          =   2610
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   4604
         Cols0           =   13
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   "N°-OK-Afecta Ope.Garant.-Cliente-Cta Crédito-Cuota-Monto Pago-ITF-Fecha Pago-Cta Recaudo-Cubre Recaudo-Cta Operador-Cubre Operador"
         EncabezadosAnchos=   "350-500-1500-2500-1800-600-1200-800-1200-1800-1250-1800-1200"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
         ColumnasAEditar =   "X-1-2-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-4-3-0-0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-L-L-C-C-R-R-C-C-R-C-R"
         FormatosEdit    =   "0-0-0-0-0-3-2-2-0-0-2-0-2"
         AvanceCeldas    =   1
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total de Registros:"
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
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   3210
         Width           =   1635
      End
      Begin VB.Label lblNumRegistros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Height          =   300
         Left            =   1800
         TabIndex        =   2
         Top             =   3195
         Width           =   825
      End
   End
   Begin VB.Label lblProgreso 
      Caption         =   "Espere un momento.. Se estan procesando los pagos Ecotaxi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4920
      TabIndex        =   7
      Top             =   3675
      Visible         =   0   'False
      Width           =   5415
   End
End
Attribute VB_Name = "frmCredPagoCuotaEcotaxi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oNCredito As COMNCredito.NCOMCredito
Dim fsOpeCod As String

Private Sub Form_Load()
    CentraForm Me
    Set oNCredito = New COMNCredito.NCOMCredito
    'Para que puedan escoger cobertura
    If fsOpeCod = gCredPagoCuotasEcotaxi Then
        feCreditosEcotaxi.ColWidth(2) = 0
        feCreditosEcotaxi.ColWidth(11) = 0
        feCreditosEcotaxi.ColWidth(12) = 0
    Else
        feCreditosEcotaxi.ColWidth(2) = 1500
        feCreditosEcotaxi.ColWidth(11) = 1800
        feCreditosEcotaxi.ColWidth(12) = 1300
    End If
End Sub
Public Sub Inicio(ByVal psOpeCod As String, ByVal psOpeDesc As String)
    fsOpeCod = psOpeCod
    Caption = psOpeCod & " " & UCase(psOpeDesc)
    If Not HoyEsFechaPagoCuotasEcotaxi Then
        Unload Me
        Exit Sub
    End If
    Me.Show 0, MDISicmact
End Sub
Private Sub cmdBuscar_Click()
    Dim oCredito As COMDCredito.DCOMCredito
    Dim oConsSist As New COMDConstSistema.NCOMConstSistema
    Dim rsCreditos As New ADODB.Recordset
    Dim lnITF As Double, lnMontoAPagar As Double, lnTotal As Double
    Dim lnITFCtaRecaudo As Double
    Dim lnMontoCtaAhoCliente As Double, lnMontoCtaAhoOperador As Double
    Dim lnMontoCtaAhoClienteCubre As Double, lnMontoCtaAhoOperadorCubre As Double
    Dim ldFechaPago As Date
    Dim lnNroRegistros As Long

On Error GoTo ErrBuscar
    Screen.MousePointer = 11
    If fsOpeCod = gCredPagoCuotasEcotaxi Then
        ldFechaPago = CDate(Format(oConsSist.LeeConstSistema(gConstSistDiaPagoCalendEcotaxi), "00") & "/" & Format(Month(gdFecSis), "00") & "/" & Year(gdFecSis))
    Else
        ldFechaPago = CDate(Format(oConsSist.LeeConstSistema(gConstSistDiaPagoCalendEcotaxiCoberturaOG), "00") & "/" & Format(Month(gdFecSis), "00") & "/" & Year(gdFecSis))
    End If

    Set oCredito = New COMDCredito.DCOMCredito
    'Set rsCreditos = oCredito.RecuperaCuotasAPagarCreditoEcotaxi(Right(gsCodAge, 2), gdFecSis)
    Set rsCreditos = oCredito.RecuperaCuotasAPagarCreditoEcotaxi(Right(gsCodAge, 2), ldFechaPago) 'EJVG20130611
    
    Call LimpiarCampos
    
    If RSVacio(rsCreditos) Then
        'MsgBox "No existen Créditos Ecotaxi con pago de cuota para el día " & Format(gdFecSis, "dd/mm/yyyy") & " en la " & gsNomAge, vbInformation, "Aviso"
        MsgBox "No existen Créditos Ecotaxi con pago de cuota pendiente hasta el día " & Format(ldFechaPago, "dd/mm/yyyy") & " en la " & gsNomAge, vbInformation, "Aviso"
        lblNumRegistros.Caption = 0
        Screen.MousePointer = 0
        Exit Sub
    End If

    Do While Not rsCreditos.EOF
        lnITF = 0
        lnMontoAPagar = 0
        lnTotal = 0
        lnMontoCtaAhoCliente = 0
        lnITFCtaRecaudo = 0
        lnMontoCtaAhoCliente = 0
        lnMontoCtaAhoClienteCubre = 0
        lnMontoCtaAhoOperadorCubre = 0

        lnMontoAPagar = CDbl(rsCreditos!nMontoPagar)
        lnITF = oNCredito.DameMontoITF(lnMontoAPagar)
        lnTotal = lnMontoAPagar + lnITF

        lnMontoCtaAhoCliente = CDbl(rsCreditos!nSaldoRecaudo)
        lnITFCtaRecaudo = oNCredito.DameMontoITF(lnMontoCtaAhoCliente)
        
        'Para que cobre ITF en la debitación
        If lnITFCtaRecaudo > 0 Then
            lnMontoCtaAhoCliente = lnMontoCtaAhoCliente - lnITFCtaRecaudo
        End If

        If lnTotal <= lnMontoCtaAhoCliente Then
            lnMontoCtaAhoClienteCubre = lnTotal
            lnMontoCtaAhoOperadorCubre = 0
        Else
            lnMontoCtaAhoClienteCubre = lnMontoCtaAhoCliente
            If rsCreditos!cCtaCodGarante <> "" Then
            lnMontoCtaAhoOperadorCubre = lnTotal - lnMontoCtaAhoCliente
            End If
        End If
        
        'gCredPagoCuotasEcotaxi: Se mostrarán todos las créditos, pero ninguno podrá ser coberturado
        'gCredPagoCuotasEcotaxiCoberturaOG: Se mostrarán los que no se pagaron, pueden ser coberturados
        If fsOpeCod = gCredPagoCuotasEcotaxi Then
            lnMontoCtaAhoOperadorCubre = 0
        End If

        feCreditosEcotaxi.AdicionaFila
        feCreditosEcotaxi.TextMatrix(feCreditosEcotaxi.row, 1) = "1"
        feCreditosEcotaxi.TextMatrix(feCreditosEcotaxi.row, 2) = "SI" & Space(75) & "1"
        feCreditosEcotaxi.TextMatrix(feCreditosEcotaxi.row, 3) = rsCreditos!cPersNombre
        feCreditosEcotaxi.TextMatrix(feCreditosEcotaxi.row, 4) = rsCreditos!cCtaCod
        feCreditosEcotaxi.TextMatrix(feCreditosEcotaxi.row, 5) = rsCreditos!nCuota
        feCreditosEcotaxi.TextMatrix(feCreditosEcotaxi.row, 6) = Format(lnMontoAPagar, "##,##0.00")
        feCreditosEcotaxi.TextMatrix(feCreditosEcotaxi.row, 7) = Format(lnITF, "##,##0.00")
        feCreditosEcotaxi.TextMatrix(feCreditosEcotaxi.row, 8) = Format(rsCreditos!dVenc, "dd/mm/yyyy")
        feCreditosEcotaxi.TextMatrix(feCreditosEcotaxi.row, 9) = rsCreditos!cCtaCodRecaudo
        feCreditosEcotaxi.TextMatrix(feCreditosEcotaxi.row, 10) = Format(lnMontoCtaAhoClienteCubre, "##,##0.00")
        feCreditosEcotaxi.TextMatrix(feCreditosEcotaxi.row, 11) = rsCreditos!cCtaCodGarante
        feCreditosEcotaxi.TextMatrix(feCreditosEcotaxi.row, 12) = Format(lnMontoCtaAhoOperadorCubre, "##,##0.00")
        rsCreditos.MoveNext
    Loop
    'lblNumRegistros.Caption = rsCreditos.RecordCount
    If fsOpeCod = gCredPagoCuotasEcotaxi Then
        feCreditosEcotaxi.ColWidth(2) = 0
        feCreditosEcotaxi.ColWidth(11) = 0
        feCreditosEcotaxi.ColWidth(12) = 0
    Else
        feCreditosEcotaxi.ColWidth(2) = 1500
        feCreditosEcotaxi.ColWidth(11) = 1800
        feCreditosEcotaxi.ColWidth(12) = 1300
    End If
    feCreditosEcotaxi.TopRow = 1
    feCreditosEcotaxi.row = 1
    lnNroRegistros = IIf(feCreditosEcotaxi.TextMatrix(1, 0) = "", 0, feCreditosEcotaxi.Rows - 1)
    lblNumRegistros.Caption = Format(lnNroRegistros, "##,###,##0")
    OptSelec.iTem(0).value = True
    If lnNroRegistros = 0 Then Exit Sub
    
    CmdBuscar.Enabled = False
    cmdGrabar.Enabled = True
    
    Set oCredito = Nothing
    Set oConsSist = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrBuscar:
    Screen.MousePointer = 0
    MsgBox TextErr(err.Description), vbCritical, "Aviso"
    CmdBuscar.Enabled = False
    cmdGrabar.Enabled = False
End Sub
Private Sub cmdGrabar_Click()
    Dim vPrevio As previo.clsprevio
    Dim MatCreditos As Variant
    Dim lsMsgError As String, lsCadCreditosPagados As String, lsCadCreditosNoCoberturados As String
    Dim i As Long, iMat As Long
    Dim lsImpreBoleta As String
    
On Error GoTo ErrGrabarPago
    
    ReDim MatCreditos(1 To 11, 0 To 0)
    iMat = 0
    For i = 1 To feCreditosEcotaxi.Rows - 1
        'If feCreditosEcotaxi.TextMatrix(i, 1) = "." Then
        If (feCreditosEcotaxi.TextMatrix(i, 1) = ".") And (CDbl(Trim(feCreditosEcotaxi.TextMatrix(i, 10))) + CDbl(Trim(feCreditosEcotaxi.TextMatrix(i, 12))) > 0) Then 'Valida que tenga monto de cobertura
            iMat = iMat + 1
            ReDim Preserve MatCreditos(1 To 11, 0 To iMat)
            MatCreditos(1, iMat) = Trim(feCreditosEcotaxi.TextMatrix(i, 3)) 'Nombre de Cliente
            MatCreditos(2, iMat) = Trim(feCreditosEcotaxi.TextMatrix(i, 4)) 'CtaCod Credito
            MatCreditos(3, iMat) = CInt(Trim(feCreditosEcotaxi.TextMatrix(i, 5))) 'NroCuota
            MatCreditos(4, iMat) = CDbl(Trim(feCreditosEcotaxi.TextMatrix(i, 6))) 'MontoCuota
            MatCreditos(5, iMat) = Trim(feCreditosEcotaxi.TextMatrix(i, 7)) 'ITF
            MatCreditos(6, iMat) = Trim(feCreditosEcotaxi.TextMatrix(i, 9)) 'CtaCodRecaudo
            MatCreditos(7, iMat) = CDbl(Trim(feCreditosEcotaxi.TextMatrix(i, 10))) 'SaldoRecaudoCubre
            MatCreditos(8, iMat) = Trim(feCreditosEcotaxi.TextMatrix(i, 11)) 'CtaCodOperador
            MatCreditos(9, iMat) = CDbl(Trim(feCreditosEcotaxi.TextMatrix(i, 12))) 'SaldoOperadorCubre
            MatCreditos(10, iMat) = 0 'Bit de Pago satisfactorio
            MatCreditos(11, iMat) = 0 'Bit de NO cobertura del Operador Garante
        End If
    Next
    If UBound(MatCreditos, 2) = 0 Then
        'MsgBox "Ud. debe de seleccionar las cuotas que se van a Pagar", vbCritical, "Aviso"
        MsgBox "No existen registros para procesar los pagos, revise que haya seleccionado los registros" & Chr(10) & "o que los Montos de Coberturas sean mayor a cero", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("¿Esta seguro de grabar los pagos de las cuotas Ecotaxi?", vbYesNo + vbInformation, "Aviso") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = 11
    lblProgreso.Visible = True
    cmdGrabar.Enabled = False

    'lsMsgError = oNCredito.grabarPagoCuotasEcotaxi(MatCreditos, gsCodUser, Right(gsCodAge, 2), gdFecSis, lsImpreBoleta)
    lsMsgError = oNCredito.grabarPagoCuotasEcotaxi(MatCreditos, gsCodUser, Right(gsCodAge, 2), gdFecSis, lsImpreBoleta, fsOpeCod)
    lblProgreso.Visible = False
    Screen.MousePointer = 0
    
    If Len(lsMsgError) > 0 Then
        MsgBox lsMsgError, vbCritical, "Aviso"
    End If

    Set vPrevio = New previo.clsprevio

    lsCadCreditosPagados = RecuperaCadenaCreditosPagados(MatCreditos) 'Creditos Pagados
    If Len(lsCadCreditosPagados) > 0 Then
        MsgBox "Se realizaron los sgtes Pagos de cuotas de los Créditos Ecotaxi", vbInformation, "Aviso"
        vPrevio.Show lsCadCreditosPagados, "Pago de Créditos Ecotaxi", False
    Else
        MsgBox "No se realizó ningún Pago, verifique..", vbExclamation, "Aviso"
    End If

    'If Len(lsImpreBoleta) > 0 Then
    '    vPrevio.Show lsImpreBoleta, "Pago de Créditos Ecotaxi", False
    'End If

    Set MatCreditos = Nothing
    Set vPrevio = Nothing
    
    Call LimpiarCampos
    Exit Sub
ErrGrabarPago:
    Screen.MousePointer = 0
    MsgBox "Ha sucedido un error al realizar los Pagos de cuotas Ecotaxi", vbCritical, "Aviso"
    lblProgreso.Visible = False
    cmdGrabar.Enabled = True
End Sub
Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub FormatearGrillaCreditosEcotaxi()
    feCreditosEcotaxi.Clear
    feCreditosEcotaxi.FormaCabecera
    feCreditosEcotaxi.Rows = 2
End Sub
Private Function HoyEsFechaPagoCuotasEcotaxi() As Boolean
    Dim oConsSist As COMDConstSistema.NCOMConstSistema
    Set oConsSist = New COMDConstSistema.NCOMConstSistema
    Dim nDiaPagoCuotaEcotaxi As Integer
    If fsOpeCod = gCredPagoCuotasEcotaxi Then
        nDiaPagoCuotaEcotaxi = oConsSist.LeeConstSistema(gConstSistDiaPagoCuotaEcotaxi)
    Else
        nDiaPagoCuotaEcotaxi = oConsSist.LeeConstSistema(gConstSistDiaPagoCuotaEcotaxiCoberturaOG)
    End If
    HoyEsFechaPagoCuotasEcotaxi = True
    If nDiaPagoCuotaEcotaxi <> Day(gdFecSis) Then
        MsgBox "Los días de Pago de Cuotas de los Créditos de Ecotaxi son los " & nDiaPagoCuotaEcotaxi & " de cada mes", vbInformation, "Aviso"
        HoyEsFechaPagoCuotasEcotaxi = False
    End If
End Function
'EJVG20121126 ***
Private Function RecuperaCadenaCreditosNoCoberturados(ByRef pMatCreditos As Variant)
    Dim oFun As New COMFunciones.FCOMImpresion
    Dim oAlinear As New COMFunciones.FCOMCadenas
    Dim lsCadena As String
    Dim lsItem As String * 4
    Dim lsCtaCod As String * 18
    Dim lsNombreCliente As String * 40
    Dim lsNroCuota As String * 5
    Dim lsImporteCuota As String * 9
    Dim lsITFCuota As String * 6
    Dim lsEstadoNoProcesado As String * 12
    Dim i As Integer
    Dim lnLinea As Long, lnNroPagina As Long

    lnLinea = 57
    lnNroPagina = 1
    For i = 1 To UBound(pMatCreditos, 2)
        If pMatCreditos(10, i) = 0 Then
            If lnLinea > 56 Then
                If i > 1 Then
                    lsCadena = lsCadena & Chr(12)
                    lnNroPagina = lnNroPagina + 1
                End If
                lsCadena = lsCadena & oFun.CabeceraPagina("CREDITOS ECOTAXI NO PROCESADOS FECHA " & Format(gdFecSis, "dd/mm/yyyy"), lnNroPagina - 1, 1, gsNomAge, gsInstCmac, gdFecSis, , False) '4 lineas
                lsCadena = lsCadena & String(98, "-") & Chr(10)
                lsCadena = lsCadena & "ITEM" & Space(1) & "   NRO CUENTA    " & Space(1) & "              NOMBRE CLIENTE            " & Space(1) & "CUOTA" & Space(1) & "  IMPORTE " & Space(1) & "  ITF " & Space(1) & "   ESTADO   " & Chr(10)
                lsCadena = lsCadena & String(98, "-") & Chr(10)
                lnLinea = 7
            End If

            lsItem = oAlinear.AlinearTexto(i, 4, Centro)
            lsNombreCliente = oAlinear.AlinearTexto(Left(Trim(pMatCreditos(1, i)), 40), 40, izquierda)
            lsCtaCod = oAlinear.AlinearTexto(Trim(pMatCreditos(2, i)), 18, Centro)
            lsNroCuota = oAlinear.AlinearTexto(Trim(pMatCreditos(3, i)), 5, Centro)
            lsImporteCuota = oAlinear.AlinearTexto(Format(Trim(pMatCreditos(4, i)), "##,##0.00"), 9, Derecha)
            lsITFCuota = oAlinear.AlinearTexto(Format(Trim(pMatCreditos(5, i)), "##,##0.00"), 6, Derecha)
            lsEstadoNoProcesado = IIf(pMatCreditos(11, i) = 1, "NO COBERTURA", "NO PAGADO")
        
            lsCadena = lsCadena & lsItem & Space(1) & lsCtaCod & Space(1) & lsNombreCliente & Space(1) & lsNroCuota & Space(1) & lsImporteCuota & Space(1) & lsITFCuota & Space(1) & lsEstadoNoProcesado & Chr(10)
            lnLinea = lnLinea + 1
        End If
    Next
    RecuperaCadenaCreditosNoCoberturados = lsCadena
    Set oFun = Nothing
    Set oAlinear = Nothing
End Function
Private Function RecuperaCadenaCreditosPagados(ByRef pMatCreditos As Variant) As String
    Dim oFun As New COMFunciones.FCOMImpresion
    Dim oAlinear As New COMFunciones.FCOMCadenas
    Dim lsCadena As String
    Dim lsItem As String * 4
    Dim lsCtaCod As String * 18
    Dim lsNombreCliente As String * 40
    Dim lsNroCuota As String * 5
    Dim lsImporteCuota As String * 9
    Dim lsITFCuota As String * 6
    Dim i As Integer
    Dim lnLinea As Long, lnNroPagina As Long

    lnLinea = 57
    lnNroPagina = 1
    For i = 1 To UBound(pMatCreditos, 2)
        If pMatCreditos(10, i) = 1 Then
            If lnLinea > 56 Then
                If i > 1 Then
                    lsCadena = lsCadena & Chr(12)
                    lnNroPagina = lnNroPagina + 1
                End If
                lsCadena = lsCadena & oFun.CabeceraPagina("PAGO CUOTAS CREDITOS ECOTAXI FECHA " & Format(gdFecSis, "dd/mm/yyyy"), lnNroPagina - 1, 1, gsNomAge, gsInstCmac, gdFecSis, , False) '4 lineas
                lsCadena = lsCadena & String(88, "-") & Chr(10)
                lsCadena = lsCadena & "ITEM" & Space(1) & "   NRO CUENTA    " & Space(1) & "              NOMBRE CLIENTE            " & Space(1) & "CUOTA" & Space(1) & "  IMPORTE " & Space(1) & "  ITF " & Space(1) & Chr(10)
                lsCadena = lsCadena & String(88, "-") & Chr(10)
                lnLinea = 7
            End If

            lsItem = oAlinear.AlinearTexto(i, 4, Centro)
            lsNombreCliente = oAlinear.AlinearTexto(Left(Trim(pMatCreditos(1, i)), 40), 40, izquierda)
            lsCtaCod = oAlinear.AlinearTexto(Trim(pMatCreditos(2, i)), 18, Centro)
            lsNroCuota = oAlinear.AlinearTexto(Trim(pMatCreditos(3, i)), 5, Centro)
            lsImporteCuota = oAlinear.AlinearTexto(Format(Trim(pMatCreditos(4, i)), "##,##0.00"), 9, Derecha)
            lsITFCuota = oAlinear.AlinearTexto(Format(Trim(pMatCreditos(5, i)), "##,##0.00"), 6, Derecha)
        
            lsCadena = lsCadena & lsItem & Space(1) & lsCtaCod & Space(1) & lsNombreCliente & Space(1) & lsNroCuota & Space(1) & lsImporteCuota & Space(1) & lsITFCuota & Space(1) & Chr(10)
            lnLinea = lnLinea + 1
        End If
    Next
    RecuperaCadenaCreditosPagados = lsCadena
    Set oFun = Nothing
    Set oAlinear = Nothing
End Function
Private Sub Form_Unload(Cancel As Integer)
    Set oNCredito = Nothing
End Sub
Private Sub LimpiarCampos()
    FormatearGrillaCreditosEcotaxi
    lblNumRegistros.Caption = 0
    cmdGrabar.Enabled = False
    CmdBuscar.Enabled = True
    lblProgreso.Visible = False
End Sub
'END EJVG *******
Private Sub feCreditosEcotaxi_OnChangeCombo()
    Call EstableCoberturaOperadorGarante(feCreditosEcotaxi.row)
End Sub
Private Sub feCreditosEcotaxi_RowColChange()
    Dim rsOpt As New ADODB.Recordset
    If feCreditosEcotaxi.Col = 2 Then
        With rsOpt
            .Fields.Append "desc", adVarChar, 10
            .Fields.Append "value", adVarChar, 2
            .Open
            .AddNew
            .Fields("desc") = "SI"
            .Fields("value") = "1"
            .AddNew
            .Fields("desc") = "NO"
            .Fields("value") = "2"
        End With
        rsOpt.MoveFirst
        feCreditosEcotaxi.CargaCombo rsOpt
    End If
    Set rsOpt = Nothing
End Sub
Private Sub feCreditosEcotaxi_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(feCreditosEcotaxi.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub
Private Sub EstableCoberturaOperadorGarante(ByVal pnRow As Integer)
    Dim lnMontoCtaAhoClienteCubre As Double, lnMontoCtaAhoOperadorCubre As Double
    Dim lnMontoCuota As Double, lnITF As Double
    Dim lsAfectaOG As String

    lsAfectaOG = Trim(Right(feCreditosEcotaxi.TextMatrix(pnRow, 2), 2))
    If lsAfectaOG = "1" Then 'SI
        If feCreditosEcotaxi.TextMatrix(feCreditosEcotaxi.row, 11) <> "" Then 'Verifica que el crédito cuente con operador garante
        lnMontoCuota = CDbl(feCreditosEcotaxi.TextMatrix(pnRow, 6))
        lnITF = CDbl(feCreditosEcotaxi.TextMatrix(pnRow, 7))
        lnMontoCtaAhoClienteCubre = CDbl(feCreditosEcotaxi.TextMatrix(pnRow, 10))
        lnMontoCtaAhoOperadorCubre = (lnMontoCuota + lnITF) - lnMontoCtaAhoClienteCubre
        End If
    Else
        lnMontoCtaAhoOperadorCubre = 0
    End If
    feCreditosEcotaxi.TextMatrix(pnRow, 12) = Format(lnMontoCtaAhoOperadorCubre, "##,##0.00")
End Sub
Private Sub OptSelec_Click(Index As Integer)
    Dim lsSelecciona As String
    Dim i As Long
    If Index = 0 Then
        lsSelecciona = "1"
    Else
        lsSelecciona = "0"
    End If
    If feCreditosEcotaxi.TextMatrix(1, 0) <> "" Then
        For i = 1 To feCreditosEcotaxi.Rows - 1
            feCreditosEcotaxi.TextMatrix(i, 1) = lsSelecciona
        Next
    End If
End Sub
