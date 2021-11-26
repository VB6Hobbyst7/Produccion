VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMColRecImpre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public gsImpresora As Impresoras

'Cabecera de Impresiones de Reportes de Pignoraticio
Public Function nImprimeCabeceraReportes(ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, _
        ByVal psFechaSis As String, ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String, _
        Optional ByVal psImpresora As Impresoras = gEPSON) As String
        
  Dim lsCabe01 As String, lsCabe02 As String
  Dim lsCabe03 As String, lsCabe04 As String
  Dim lsCabRepo As String
  Dim oFun As New COMFunciones.FCOMImpresion
  Dim oFunI As New COMFunciones.FCOMVarImpresion
  oFunI.Inicia psImpresora

  lsCabRepo = ""
  ' Cabecera 1
  lsCabe01 = oFun.FillText(Trim(UCase(psNomCmac)), 55, " ")
  lsCabe01 = lsCabe01 & Space(pnAnchoLinea - 55 - 25)
  lsCabe01 = lsCabe01 & "Pag.  : " & Str(pnPagina) & "  -  " & psCodUser & Chr(10)
  'lsCabe01 = lsCabe01 & IIf(pbCiereDia = True, IIf(VerifSiCierreDia(), "DC", "AC"), "") & ofunI.gPrnSaltoLinea
  ' Cabecera 2
  lsCabe01 = lsCabe01 & oFun.FillText(Trim(UCase(psNomAgencia)), 35, " ")
  lsCabe01 = lsCabe01 & Space(pnAnchoLinea - 35 - 25)
  lsCabe01 = lsCabe01 & "Fecha : " & Format(psFechaSis & " " & Time, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
  ' Titulo
  lsCabe02 = String(Int((pnAnchoLinea - Len(psTitulo)) / 2), " ") & psTitulo & oFunI.gPrnSaltoLinea
  ' SubTitulo
  lsCabe03 = String(Int((pnAnchoLinea - Len(psSubTitulo)) / 2), " ") & psSubTitulo & oFunI.gPrnSaltoLinea
  ' Comenta
  lsCabe04 = IIf(Len(psComenta) > 0, psComenta & oFunI.gPrnSaltoLinea, "")
  ' ***
  lsCabRepo = lsCabRepo & lsCabe01 & lsCabe02
  lsCabRepo = lsCabRepo & lsCabe03 & lsCabe04
  nImprimeCabeceraReportes = lsCabRepo
End Function

Public Function nPrintGastosRecuperaciones(ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, _
        ByVal psFecSis As String, ByVal psNomCliente As String, ByVal prGastos As ADODB.Recordset, _
        ByVal psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
    
Dim lsCadImp As String
Dim oFun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
    Dim lnNroPag As Integer, lnNroLinea As Integer
    Dim lsTitulo As String, lsSubtitulo As String, lsComenta As String
    
    
    lnNroPag = 1
    lsTitulo = "LISTADO DE GASTOS DE CREDITO EN RECUPERACIONES"
    lsSubtitulo = ""
        lsCadImp = "" & Chr(10)
        lsCadImp = lsCadImp & nImprimeCabeceraReportes(psNomCmac, psNomAgencia, psCodUser, psFecSis, lsTitulo, lsSubtitulo, lnNroPag, 110, lsComenta)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & Space(2) & "Cliente : " & psNomCliente & Chr(10)
        'lsCadImp = lsCadImp & Space(2) & "Credito : " & psCredito
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & Chr(10)
        lsCadImp = lsCadImp & Space(2) & String(90, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(2) & "Nro  FECHA        GASTO                                MONTO GASTO  MONTO PAGADO  ESTADO " & Chr(10)
        lsCadImp = lsCadImp & Space(2) & String(90, "-") & Chr(10)
        lnNroLinea = 8
        Do While Not prGastos.EOF
            lnNroLinea = lnNroLinea + 1
            lsCadImp = lsCadImp & oFun.ImpreFormat(prGastos("nNroGastoCta"), 4, 0) & Space(2) & Format(prGastos("dAsigna"), "dd/mm/yyyy") & Space(2) & oFun.ImpreFormat(Mid(prGastos("cMotivoGasto"), 1, 30), 30) _
                        & Space(2) & oFun.ImpreFormat(prGastos("nMonto"), 12, 2) & oFun.ImpreFormat(prGastos("nMontoPagado"), 12, 2) & Space(2)
            Select Case prGastos("nColocRecGastoEstado")
                Case gColRecGastoEstEliminado
                    lsCadImp = lsCadImp & "Eliminado"
                Case gColRecGastoEstPagado
                    lsCadImp = lsCadImp & "Pagado   "
                Case gColRecGastoEstPendiente
                    lsCadImp = lsCadImp & "Pendiente"
            End Select
            lsCadImp = lsCadImp & Chr(10)
            If lnNroLinea > 60 Then
                lsCadImp = lsCadImp & Chr(10)
                lnNroPag = lnNroPag + 1
                lsCadImp = lsCadImp & nImprimeCabeceraReportes(psNomCmac, psNomAgencia, psCodUser, psFecSis, lsTitulo, lsSubtitulo, lnNroPag, 90, lsComenta)
                lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
                lsCadImp = lsCadImp & Space(2) & "Cliente : " & psNomCliente & Chr(10)
                lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
                lsCadImp = lsCadImp & Space(2) & String(90, "-") & Chr(10)
                lsCadImp = lsCadImp & Space(10) & " FECHA                       GASTO                 MONTO GASTO     MONTO PAGADO   " & Chr(10)
                lsCadImp = lsCadImp & Space(2) & String(90, "-") & Chr(10)
                lnNroLinea = 8
            End If
            prGastos.MoveNext
        Loop
        lsCadImp = lsCadImp & Chr(10) & Space(2) & String(90, "-") & Chr(10)
                            
        lsCadImp = lsCadImp & " " & oFun.ImpreFormat(psmensaje, 40) _
                            & Chr(10)
        
        nPrintGastosRecuperaciones = lsCadImp
End Function

Public Function nPrintReciboPagoCredRecupLima(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodcta As String, ByVal psNombCliente As String, _
        ByVal pnMontoPago As Currency, _
        ByVal psCodUser As String, ByVal pnCapital As Currency, _
        ByVal pnInteres As Currency, ByVal pnMora As Currency, _
        ByVal pnGastos As Currency, Optional ByVal psmensaje As String = "", Optional ByVal psImpresora As Impresoras = gEPSON) As String
Dim lsCadImp As String
Dim vEspacio As Integer
Dim oFun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

lsCadImp = lsCadImp & Chr$(27) & Chr$(64)
lsCadImp = lsCadImp & Chr$(27) & Chr$(65)   'espaciamiento lineas 1/6 pulg.
lsCadImp = lsCadImp & Chr$(27) & Chr$(67) & Chr$(22)  'Longitud de página a 22 líneas'
lsCadImp = lsCadImp & Chr$(27) & Chr$(77)   'Tamaño 10 cpi
lsCadImp = lsCadImp & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
lsCadImp = lsCadImp & Chr$(27) + Chr$(18) ' cancela condensada
lsCadImp = lsCadImp & Chr$(27) + Chr$(72) ' desactiva negrita

    vEspacio = 9
        lsCadImp = "" & Chr(10) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & " Recuperaciones - Pago " & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & " " & oFun.ImpreFormat(psNomAgencia, 26, 0) & Format(psFecTransac, "dd/mm/yyyy hh:mm:ss") & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & " Credito   : " & Space(9) & Mid(psCodcta, 1, 5) & "-" & Mid(psCodcta, 6) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & " CLIENTE   : " & oFun.ImpreFormat(psNombCliente, 27) & Chr(10)
                            
        lsCadImp = lsCadImp & Chr(10) & Chr(10)
        lsCadImp = lsCadImp & " ==========  Comprobante Pago ========== " & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & " CAPITAL          :" & oFun.ImpreFormat(pnCapital, 18, , True) & Chr(10)
        lsCadImp = lsCadImp & " INTERES COMP.    :" & oFun.ImpreFormat(pnInteres, 18, , True) & Chr(10)
        lsCadImp = lsCadImp & " MORA             :" & oFun.ImpreFormat(pnMora, 18, , True) & Chr(10)
        lsCadImp = lsCadImp & " GASTOS           :" & oFun.ImpreFormat(pnGastos, 18, , True) & Chr(10)
        lsCadImp = lsCadImp & " NETO A PAGAR     :" & oFun.ImpreFormat(pnMontoPago, 18, , True) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & " ===================================== " & Chr(10)
        lsCadImp = lsCadImp & " " & Format(psCodUser, "@@@@") & Chr(10)
        lsCadImp = lsCadImp & " " & oFun.ImpreFormat(psmensaje, 40) & Chr(12)
        nPrintReciboPagoCredRecupLima = lsCadImp
End Function

'Procedimiento de impresión del recibo de renovación
Public Function nPrintReciboPagoCredRecup(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodcta As String, ByVal psNombCliente As String, _
        ByVal pnMontoPago As Currency, _
        ByVal psCodUser As String, Optional ByVal psmensaje As String = "", Optional ByVal pnITF As Double = 0#, _
        Optional ByVal psImpresora As Impresoras = gEPSON, Optional ByVal pbImpTMU As Boolean, _
        Optional ByVal psOpeCod As String = "") As String 'FRHU 20150603 ERS022-2015
    Dim lsCadImp As String
    Dim vEspacio As Integer

    Dim oFun As New COMFunciones.FCOMImpresion
    Dim oFunI As New COMFunciones.FCOMVarImpresion
    Dim sMoneda As String
    oFunI.Inicia psImpresora
    lsCadImp = ""
    lsCadImp = lsCadImp & Chr$(27) & Chr$(64)
    sMoneda = IIf(Mid(psCodcta, 9, 1) = 1, "SOLES", "DOLARES")
    'JUEZ 20131021 *******************************************************************************
    Dim oGen As New COMDConstSistema.DCOMGeneral, nEstadoCampDatos As Integer, lsCadImpCamp As String
    nEstadoCampDatos = oGen.LeeConstSistema(gConstSistEstadoCampActualizaDatos)
    If nEstadoCampDatos = 1 Then
        lsCadImpCamp = ""
        If pbImpTMU = False Then
            lsCadImpCamp = lsCadImpCamp & Chr(10) & Chr$(27) & Chr$(69)
            lsCadImpCamp = lsCadImpCamp & "                            * *         " & Space(13) & "                            * *         " & Chr(10)
            lsCadImpCamp = lsCadImpCamp & "                         *      *       " & Space(13) & "                         *      *       " & Chr(10)
            lsCadImpCamp = lsCadImpCamp & "                        *               " & Space(13) & "                        *               " & Chr(10)
            lsCadImpCamp = lsCadImpCamp & "      * *              *          *     " & Space(13) & "      * *              *          *     " & Chr(10)
            lsCadImpCamp = lsCadImpCamp & "    *     *      *   *             *    " & Space(13) & "    *     *      *   *             *    " & Chr(10)
            lsCadImpCamp = lsCadImpCamp & "   *        *  *   *                    " & Space(13) & "   *        *  *   *                    " & Chr(10)
            lsCadImpCamp = lsCadImpCamp & "  *           *                     *   " & Space(13) & "  *           *                     *   " & Chr(10)
            lsCadImpCamp = lsCadImpCamp & " *                                   *  " & Space(13) & " *                                   *  " & Chr(10)
            lsCadImpCamp = lsCadImpCamp & "*                                     * " & Space(13) & "*                                     * " & Chr(10)
            lsCadImpCamp = lsCadImpCamp & "ACTUALIZA Y GANA UN PASAJE DOBLE A CUSCO" & Space(13) & "ACTUALIZA Y GANA UN PASAJE DOBLE A CUSCO" & Chr$(27) & Chr$(70) & Chr(10)
            lsCadImpCamp = lsCadImpCamp & Chr(10)
        Else
            lsCadImpCamp = lsCadImpCamp & Chr(10) & Chr$(27) & Chr$(69) & "C"
            lsCadImpCamp = lsCadImpCamp & "                            * *" & Chr(10)
            lsCadImpCamp = lsCadImpCamp & "                         *      *" & Chr(10)
            lsCadImpCamp = lsCadImpCamp & "                        *         " & Chr(10)
            lsCadImpCamp = lsCadImpCamp & "      * *              *          *" & Chr(10)
            lsCadImpCamp = lsCadImpCamp & "    *     *      *   *             *" & Chr(10)
            lsCadImpCamp = lsCadImpCamp & "   *        *  *   *                " & Chr(10)
            lsCadImpCamp = lsCadImpCamp & "  *           *                     *" & Chr(10)
            lsCadImpCamp = lsCadImpCamp & " *                                   *" & Chr(10)
            lsCadImpCamp = lsCadImpCamp & "*                                     *" & Chr(10)
            lsCadImpCamp = lsCadImpCamp & "ACTUALIZA Y GANA UN PASAJE DOBLE A CUSCO" & Chr$(27) & Chr$(70) & Chr(10)
            lsCadImpCamp = lsCadImpCamp & Chr(10)
        End If
    End If
    'END JUEZ ************************************************************************************
    
    If pbImpTMU = False Then
        'lsCadImp = lsCadImp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
        'lsCadImp = lsCadImp & Chr$(27) & Chr$(67) & Chr$(22)  'Longitud de página a 22 líneas'
        lsCadImp = lsCadImp & oFunI.gPrnTamLetra10CPI 'Chr$(27) & Chr$(77)   'Tamaño 10 cpi
        lsCadImp = lsCadImp & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
        'lsCadImp = lsCadImp & Chr$(27) + Chr$(18) ' cancela condensada
        'lsCadImp = lsCadImp & Chr$(27) + Chr$(72) ' desactiva negrita
    
        vEspacio = 16
        'lsCadImp = "" & Chr(10) & Chr(10) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        'FRHU 20150603 ERS022-2015
        'lsCadImp = lsCadImp & " Recuperaciones - Pago " & Space(10) & Space(12 + vEspacio) _
        '                    & " Recuperaciones - Pago " & Chr(10)
        If psOpeCod = gColRecOpePagTransfFocMacmEfe Or psOpeCod = gColRecOpePagTransfFocMacmVou Or psOpeCod = gColRecOpePagTransfFocMacmChq Then
            'lsCadImp = lsCadImp & " Recuperaciones - Pago Transf FOCMAC " & Space(10) & Space(12 + vEspacio) _ 'FRHU 20150611 Observacion

            lsCadImp = lsCadImp & " Recuperaciones - Pago Transf FOCMAC " & Space(10) & Space(vEspacio - 11) _
                                & " Recuperaciones - Pago Transf FOCMAC " & Chr(10)
        Else
            lsCadImp = lsCadImp & " Recuperaciones - Pago " & Space(10) & Space(12 + vEspacio) _
                                & " Recuperaciones - Pago " & Chr(10)
        End If
        'FIN FRHU 20150603
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & " " & oFun.ImpreFormat(psNomAgencia, 24, 0) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Space(vEspacio - 5) _
                            & " " & oFun.ImpreFormat(psNomAgencia, 24, 0) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Chr(10)
        lsCadImp = lsCadImp & " " & oFun.ImpreFormat(sMoneda, 24, 0) & Space(vEspacio + 12) _
                            & " " & oFun.ImpreFormat(sMoneda, 24, 0) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & "Credito   : " & Space(9) & Mid(psCodcta, 1, 5) & "-" & Mid(psCodcta, 6) & Space(vEspacio - 3) _
                            & "Credito   : " & Mid(psCodcta, 1, 5) & "-" & Mid(psCodcta, 6) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & "CLIENTE   : " & oFun.ImpreFormat(psNombCliente, 27) & Space(vEspacio - 4) _
                            & "CLIENTE   : " & oFun.ImpreFormat(psNombCliente, 27) & Chr(10)
                            
        lsCadImp = lsCadImp & Chr(10) & Chr(10)
        
        lsCadImp = lsCadImp & " ==========  Comprobante Pago ========== " & Space(vEspacio - 5) _
                            & " ==========  Comprobante Pago ========== " & Chr(10)
                            
                            
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        
        lsCadImp = lsCadImp & "I.T.F.         :" & oFun.ImpreFormat(pnITF, 18, , True) & Space(vEspacio) _
                            & "I.T.F.         :" & oFun.ImpreFormat(pnITF, 18, , True) & Chr(10)
        lsCadImp = lsCadImp & "PAGO           :" & oFun.ImpreFormat(pnMontoPago, 18, , True) & Space(vEspacio) _
                            & "PAGO           :" & oFun.ImpreFormat(pnMontoPago, 18, , True) & Chr(10)
        lsCadImp = lsCadImp & "NETO A PAGAR   :" & oFun.ImpreFormat(pnMontoPago + pnITF, 18, , True) & Space(vEspacio) _
                            & "NETO A PAGAR   :" & oFun.ImpreFormat(pnMontoPago + pnITF, 18, , True) & Chr(10)
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        
        lsCadImp = lsCadImp & " ======================================= " & Space(vEspacio - 5) _
                            & " ======================================= " & Chr(10)
        lsCadImp = lsCadImp & " " & Format(psCodUser, "@@@@") & Space(30) & Space(vEspacio) _
                            & " " & Format(psCodUser, "@@@@") & Chr(10)
        
        lsCadImp = lsCadImp & " " & oFun.ImpreFormat(psmensaje, 40) & Space(vEspacio) _
                            & " " & oFun.ImpreFormat(psmensaje, 40) & Chr(10)
        'lsCadImp = lsCadImp & Chr(10) & lsCadImp & Chr(10) & Chr(10) & Chr(10)
        lsCadImp = lsCadImpCamp & lsCadImp & Chr(10) & lsCadImp & Chr(10) & Chr(10) & Chr(10) 'JUEZ 20131021
    Else
        'lsCadImp = lsCadImp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
        'lsCadImp = lsCadImp & Chr$(27) & Chr$(67) & Chr$(22)  'Longitud de página a 22 líneas'
        lsCadImp = lsCadImp & oFunI.gPrnTamLetra10CPI 'Chr$(27) & Chr$(77)   'Tamaño 10 cpi
        lsCadImp = lsCadImp & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
        'lsCadImp = lsCadImp & Chr$(27) + Chr$(18) ' cancela condensada
        'lsCadImp = lsCadImp & Chr$(27) + Chr$(72) ' desactiva negrita
    
        vEspacio = 16
        'lsCadImp = "" & Chr(10) & Chr(10) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        'FRHU 20150603 ERS022-2015
        'lsCadImp = lsCadImp & "RRecuperaciones - Pago " & Chr(10)
        If psOpeCod = gColRecOpePagTransfFocMacmEfe Or psOpeCod = gColRecOpePagTransfFocMacmVou Or psOpeCod = gColRecOpePagTransfFocMacmChq Then
            lsCadImp = lsCadImp & "RRecuperaciones - Pago Transf FOCMAC " & Chr(10)
        Else
            lsCadImp = lsCadImp & "RRecuperaciones - Pago " & Chr(10)
        End If
        'FIN ERS022-2015
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & oFun.ImpreFormat(psNomAgencia, 24, 0) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Chr(10)
        lsCadImp = lsCadImp & oFun.ImpreFormat(sMoneda, 24, 0) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & "CCredito: " & Space(4) & Mid(psCodcta, 1, 5) & "-" & Mid(psCodcta, 6) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & "CLIENTE: " & oFun.ImpreFormat(psNombCliente, 27) & Chr(10)
                            
        lsCadImp = lsCadImp & Chr(10)
        
        lsCadImp = lsCadImp & "==========  Comprobante Pago ==========" & Chr(10)
                            
                            
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        
        lsCadImp = lsCadImp & "II.T.F.       :" & oFun.ImpreFormat(pnITF, 18, , True) & Chr(10)
        lsCadImp = lsCadImp & "PAGO         :" & oFun.ImpreFormat(pnMontoPago, 18, , True) & Chr(10)
        lsCadImp = lsCadImp & "NETO A PAGAR :" & oFun.ImpreFormat(pnMontoPago + pnITF, 18, , True) & Chr(10)
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        
        lsCadImp = lsCadImp & "======================================= " & Chr(10)
        lsCadImp = lsCadImp & Format(psCodUser, "@@@@") & Chr(10)
        
        lsCadImp = lsCadImp & " " & oFun.ImpreFormat(psmensaje, 40) & Chr(10)
        'lsCadImp = lsCadImp & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(27) & "m" & Chr(10)
        lsCadImp = lsCadImpCamp & lsCadImp & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(27) & "m" & Chr(10) 'JUEZ 20131021
    End If
    nPrintReciboPagoCredRecup = lsCadImp
End Function

' IMPRIME  ACTA DE TRANSFERENCIA
' AVMM 09/12/2005
Public Function nImprimeActaTransferencia(ByVal psTextoCarta As String, ByVal psFecha As String, ByVal Rs As ADODB.Recordset, Optional ByVal psImpresora As Impresoras = gEPSON) As String
Dim lsCadImp As String
Dim lsCartaModelo As String
Dim lsFechaHoraGrab As String
Dim liItem As Integer
Dim lsListaTransfer As String
Dim lsTotalTransfer As String
Dim lnTotalTransfer As Double
Dim I As Integer

Dim oFun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

On Error GoTo dError


    lsListaTransfer = lsListaTransfer & Space(8) & oFun.ImpreFormat("CREDITO", 20) & oFun.ImpreFormat("CLIENTE", 34)
    lsListaTransfer = lsListaTransfer & oFun.ImpreFormat("SALDO", 5) & oFun.ImpreFormat("ATRASO", 15) & oFunI.gPrnSaltoLinea
    lsListaTransfer = lsListaTransfer & Space(2) & String(80, "-") & oFunI.gPrnSaltoLinea
    I = 1
    If Not (Rs.EOF And Rs.BOF) Then
       Rs.MoveFirst
       Do Until Rs.EOF
            If Rs!cTran = "S" Then
                lsListaTransfer = lsListaTransfer & Space(2) & oFun.ImpreFormat(I, 4, 0, False) & oFun.ImpreFormat(Rs!cCtaCod, 20)
                lsListaTransfer = lsListaTransfer & oFun.ImpreFormat(Rs!cCliente, 30)
                lsListaTransfer = lsListaTransfer & oFun.ImpreFormat(CDbl(Rs!nSaldo), 10, , True)
                lsListaTransfer = lsListaTransfer & oFun.ImpreFormat(Rs!nDiasA, 8) & oFunI.gPrnSaltoLinea
                'lsListaTransfer = lsListaTransfer & ImpreFormat(Me.FECreditos.TextMatrix(i, 7), 10)
                'Acumulo el total
                lnTotalTransfer = lnTotalTransfer + oFun.ImpreFormat(CDbl(Rs!nSaldo), 10, , True)
            End If
            I = I + 1
         Rs.MoveNext
       Loop
    End If
    lsListaTransfer = lsListaTransfer & Space(2) & String(80, "-") & oFunI.gPrnSaltoLinea
    lsTotalTransfer = "Total Creditos : " & oFun.ImpreFormat(I - 1, 5, 0, False) & Space(20) & "Saldo Capital : " & oFun.ImpreFormat(lnTotalTransfer, 12, 2, True)

    'Llena cartas
    lsCartaModelo = psTextoCarta
    lsCartaModelo = Replace(lsCartaModelo, "<<FECHAC>>", Format(psFecha, "dd/mm/yyyy"), , 1, vbTextCompare)
    lsCartaModelo = Replace(lsCartaModelo, "<<FECHAL>>", Format(psFecha, "dddd,d mmmm yyyy"), , 1, vbTextCompare)
    lsCartaModelo = Replace(lsCartaModelo, "<<LISTA>>", lsListaTransfer, , 1, vbTextCompare)
    lsCartaModelo = Replace(lsCartaModelo, "<<TOTAL>>", lsTotalTransfer, , 1, vbTextCompare)
    
    lsCadImp = lsCadImp & lsCartaModelo & Chr(12)
    nImprimeActaTransferencia = lsCadImp

Exit Function
   
dError:
    Err.Raise Err.Number, "Obtiene Datos Imprime Listado Aviso Vencimiento  <<nImprimeAvisoVencimiento>>", Err.Description
    
End Function

' IMPRESION DE TRANSFERENCIA
' AVMM 09/12/2005 - MADM 20110515
Public Function ImpresionTransferencia(ByVal Rs As ADODB.Recordset, ByVal gsNomCmac As String, ByVal gdFecSis As String, _
                                   ByVal gsNomAge As String, ByVal gsCodUser As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
Dim sCad As String
Dim I As Integer
'Dim rs As ADODB.Recordset
Dim oFun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
'Set rs = New ADODB.recorset
   
    sCad = oFunI.gPrnSaltoLinea
    sCad = sCad & Space(2) & gsNomCmac & Space(85 - Len(gsNomCmac)) & gdFecSis & oFunI.gPrnSaltoLinea
    sCad = sCad & Space(2) & gsNomAge & Space(85 - Len(gsNomAge)) & gsCodUser & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    sCad = sCad & Space(40) & " TRANSFERENCIA A JUDICIAL " & oFunI.gPrnSaltoLinea
    sCad = sCad & Space(40) & String(30, "-") & oFunI.gPrnSaltoLinea
    sCad = sCad & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    sCad = sCad & Space(6) & oFun.ImpreFormat("CREDITO", 20) & oFun.ImpreFormat("DEMANDA", 7) & oFun.ImpreFormat("CLIENTE", 34)
    sCad = sCad & oFun.ImpreFormat("PRESTAMO", 12) & oFun.ImpreFormat("SALDO", 5) & oFun.ImpreFormat("REFINAN", 10) & oFunI.gPrnSaltoLinea
    sCad = sCad & Space(5) & String(110, "-") & oFunI.gPrnSaltoLinea
   I = 1
   If Not (Rs.EOF And Rs.BOF) Then
      Rs.MoveFirst
      Do Until Rs.EOF
        If Rs!cTran = "S" Then
            sCad = sCad & Space(2) & oFun.ImpreFormat(I, 4, 0, False) & oFun.ImpreFormat(Rs!cCtaCod, 20)
            sCad = sCad & oFun.ImpreFormat(Rs!cCond, 7)
            sCad = sCad & oFun.ImpreFormat(Rs!cCliente, 30)
            sCad = sCad & oFun.ImpreFormat(CDbl(Rs!nMonto), 10, , True)
            sCad = sCad & oFun.ImpreFormat(CDbl(Rs!nSaldo), 10, , True)
            sCad = sCad & oFun.ImpreFormat(Rs!cEstado, 10) & oFunI.gPrnSaltoLinea
        End If
        Rs.MoveNext
        I = I + 1
      Loop
    End If
    sCad = sCad & Space(5) & String(110, "-") & oFunI.gPrnSaltoLinea
    ImpresionTransferencia = sCad
    
End Function

Public Function ImprimeNegociacion(ByVal gsNomCmac As String, ByVal gdFecSis As Date, ByVal gsNomAge As String, ByVal gsCodUser As String, _
                                   ByVal psCtaCod As String, ByVal psnomclie As String, ByVal psnomestudio As String, ByVal pnNegCuotas As Currency, ByVal Rs As ADODB.Recordset, _
                                   ByVal pnTipo As Integer, Optional ByVal psImpresora As Impresoras = gEPSON) As String
    
    Dim oFun As New COMFunciones.FCOMImpresion
    Dim oFunI As New COMFunciones.FCOMVarImpresion
    Dim nMonto As Double
    oFunI.Inicia psImpresora
    
    Dim lscadena As String
    'Set Rs = New ADODB.Recordset
    If Not (Rs.EOF And Rs.BOF) Then
       Rs.MoveFirst
       nMonto = CDbl(Rs!nMonto)
    End If
    
    
    lscadena = oFunI.gPrnSaltoLinea
    lscadena = lscadena & Space(3) & gsNomCmac & Space(50) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
    lscadena = lscadena & Space(3) & gsNomAge & Space(70) & gsCodUser & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & Chr$(27) & Chr$(69)
    lscadena = lscadena & Space(35) & IIf(pnTipo = 1, "SIMULACION DE ", "") & oFunI.gPrnSaltoLinea
    lscadena = lscadena & Space(25) & "NEGOCIACION DE CREDITO EN COBRANZA JUDICIAL" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    lscadena = lscadena & Space(5) & "Credito         : " & psCtaCod & Space(10) & IIf(Mid(psCtaCod, 9, 1) = "1", "Nuevos Soles", "Dolares") & oFunI.gPrnSaltoLinea
    lscadena = lscadena & Space(5) & "Cliente         : " & Trim(psnomclie) & Chr$(27) + Chr$(72) & oFunI.gPrnSaltoLinea
    lscadena = lscadena & Space(5) & "Abogado         : " & Trim(psnomestudio) & oFunI.gPrnSaltoLinea
    lscadena = lscadena & Space(5) & "Monto Negociado : " & IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "US$") & " " & Format(CDbl(nMonto), "#,##0.00")
    lscadena = lscadena & Space(15) & "Nro Cuotas    : " & pnNegCuotas & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    
    lscadena = lscadena & Space(5) & " * En caso de incumplimiento con algunas de las cuotas pactadas, este  convenio" & oFunI.gPrnSaltoLinea
    lscadena = lscadena & Space(5) & "   quedará resuelto automáticamente y se continuará con las acciones judiciales" & oFunI.gPrnSaltoLinea
    lscadena = lscadena & Space(5) & "   pertinentes. " & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    
    lscadena = lscadena & Space(30) & "CALENDARIO DE PAGOS" & oFunI.gPrnSaltoLinea
    lscadena = lscadena & Space(5) & String(85, "-") & oFunI.gPrnSaltoLinea
    lscadena = lscadena & "        Nro   Fecha Pago                   Monto  " & oFunI.gPrnSaltoLinea
    lscadena = lscadena & Space(5) & String(85, "-") & oFunI.gPrnSaltoLinea
    
    
    If Not (Rs.EOF And Rs.BOF) Then
       Rs.MoveFirst
       Do Until Rs.EOF
        
        lscadena = lscadena & Space(8) & oFun.ImpreFormat(Trim(Rs!nNro), 5, 0) & Space(1)
        lscadena = lscadena & oFun.ImpreCarEsp(Format(Trim(Rs!dFecha), "ddd, d mmm yyyy"))
        lscadena = lscadena & oFun.ImpreFormat(CDbl(Rs!nMonto), 14, 2, True)
        lscadena = lscadena & oFunI.gPrnSaltoLinea
        Rs.MoveNext
       Loop
    End If
    Rs.Close
    lscadena = lscadena & Space(5) & String(85, "-") & oFunI.gPrnSaltoLinea
    
    ImprimeNegociacion = lscadena
End Function


Public Function ImprimeRecMemos(ByVal psNomCmac As String, ByVal psNomAge As String, ByVal psCodUser As String, ByVal pdFecSis As Date, _
                                ByVal psCodcta As String, ByVal psNomPers As String, psText As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
    Dim lsCadImp As String
    Dim oFunI As New COMFunciones.FCOMVarImpresion
    oFunI.Inicia psImpresora
    
    lsCadImp = lsCadImp & psNomCmac & Space(40) & Format(pdFecSis & " " & Time(), "dd/mm/yyyy hh:mm:ss") & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & psNomAge & Space(49) & psCodUser & oFunI.gPrnSaltoLinea
    
    lsCadImp = lsCadImp & Space(25) & "Reporte" & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    
    lsCadImp = lsCadImp & " Credito : " & psCodcta & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & " Cliente : " & psNomPers & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & String(60, "*") & oFunI.gPrnSaltoLinea
    
    lsCadImp = lsCadImp & psText
    
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & String(60, "*") & oFunI.gPrnSaltoLinea
    ImprimeRecMemos = lsCadImp
End Function

'Public Sub fImpresora(Optional ByVal psImpresora As Impresoras = gEPSON)
'    Dim oFunI As New COMFunciones.FCOMVarImpresion
'        oFunI.Inicia psImpresora
'    Set oFunI = Nothing
'End Sub
'
'Private Sub Class_Initialize()
'    gsImpresora = "gEPSON"
'    fImpresora gsImpresora
'End Sub
Public Function ImpresionTransferenciaAdjudicados(ByVal Rs As ADODB.Recordset, ByVal gsNomCmac As String, ByVal gdFecSis As String, _
                                   ByVal gsNomAge As String, ByVal gsCodUser As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
Dim sCad As String
Dim I As Integer
Dim oFun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
   
    sCad = oFunI.gPrnSaltoLinea
    sCad = sCad & Space(2) & gsNomCmac & Space(100 - Len(gsNomCmac)) & gdFecSis & oFunI.gPrnSaltoLinea
    sCad = sCad & Space(2) & gsNomAge & Space(100 - Len(gsNomAge)) & gsCodUser & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    sCad = sCad & Space(40) & " TRANSFERENCIA A ADJUDICADOS " & oFunI.gPrnSaltoLinea
    sCad = sCad & Space(40) & String(30, "-") & oFunI.gPrnSaltoLinea
    sCad = sCad & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea

    sCad = sCad & Space(6) & oFun.ImpreFormat("CLIENTE", 34) & oFun.ImpreFormat("CREDITO", 20) & oFun.ImpreFormat("GARANTIA", 10)
    sCad = sCad & oFun.ImpreFormat("TIPO", 35) & oFun.ImpreFormat("DIRECCION", 35) & oFunI.gPrnSaltoLinea
    sCad = sCad & Space(5) & String(150, "-") & oFunI.gPrnSaltoLinea
   I = 1
   If Not (Rs.EOF And Rs.BOF) Then
      Rs.MoveFirst
      Do Until Rs.EOF
        If Rs!cTran = "R" Then
            sCad = sCad & Space(2) & oFun.ImpreFormat(I, 4, 0, False) & oFun.ImpreFormat(Rs!cCliente, 34)
            sCad = sCad & oFun.ImpreFormat(Rs!cCtaCod, 20)
            sCad = sCad & oFun.ImpreFormat(Rs!cNumGarant, 10)
            sCad = sCad & oFun.ImpreFormat(Rs!cDesTipo, 35)
            sCad = sCad & oFun.ImpreFormat(Rs!cDireccion, 35) & oFunI.gPrnSaltoLinea
        End If
        Rs.MoveNext
        I = I + 1
      Loop
    End If
    sCad = sCad & Space(5) & String(150, "-") & oFunI.gPrnSaltoLinea
    ImpresionTransferenciaAdjudicados = sCad
    
End Function

' MADM 20110627 - 20110513
Public Function ImpresionCastigo(ByVal Rs As ADODB.Recordset, ByVal gsNomCmac As String, ByVal gdFecSis As String, _
                                   ByVal gsNomAge As String, ByVal gsCodUser As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
Dim sCad As String
Dim I As Integer
Dim oFun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
   
    sCad = oFunI.gPrnSaltoLinea
    sCad = sCad & Space(2) & gsNomCmac & Space(100 - Len(gsNomCmac)) & gdFecSis & oFunI.gPrnSaltoLinea
    sCad = sCad & Space(2) & gsNomAge & Space(100 - Len(gsNomAge)) & gsCodUser & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    sCad = sCad & Space(40) & " REPORTE CASTIGO DE CREDITOS " & oFunI.gPrnSaltoLinea
    sCad = sCad & Space(40) & String(30, "-") & oFunI.gPrnSaltoLinea
    sCad = sCad & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea

    sCad = sCad & Space(6) & oFun.ImpreFormat("CLIENTE", 34) & oFun.ImpreFormat("CREDITO", 20) & oFun.ImpreFormat("SAL K", 10)
    sCad = sCad & oFun.ImpreFormat("SAL ICOMP", 10) & oFun.ImpreFormat("SALD IMORA", 10)
    sCad = sCad & oFun.ImpreFormat("SALD GASTO", 10) & oFun.ImpreFormat("INTCOMPGEN", 10) & oFunI.gPrnSaltoLinea
    sCad = sCad & Space(5) & String(120, "-") & oFunI.gPrnSaltoLinea
   I = 1
   If Not (Rs.EOF And Rs.BOF) Then
      Rs.MoveFirst
      Do Until Rs.EOF
        
        sCad = sCad & Space(2) & oFun.ImpreFormat(I, 4, 0, False) & oFun.ImpreFormat(Rs!nNombre, 34)
        sCad = sCad & oFun.ImpreFormat(Rs!cCtaCod, 18)
        sCad = sCad & oFun.ImpreFormat(Rs!nSaldoCap, 10)
        sCad = sCad & oFun.ImpreFormat(Rs!nSaldoIntCom, 10)
        sCad = sCad & oFun.ImpreFormat(Rs!nSaldoIntMor, 8)
        sCad = sCad & oFun.ImpreFormat(Rs!nSaldoGasto, 8)
        sCad = sCad & oFun.ImpreFormat(Rs!nIntComGen, 8) & oFunI.gPrnSaltoLinea
        Rs.MoveNext
        I = I + 1
      Loop
    End If
    sCad = sCad & Space(5) & String(120, "-") & oFunI.gPrnSaltoLinea
    ImpresionCastigo = sCad
'Dim sCad As String
'Dim I As Integer, J As Integer
'Dim oFun As New COMFunciones.FCOMImpresion
'Dim oFunI As New COMFunciones.FCOMVarImpresion
'Dim nChangeMoneda As Integer
'Dim nChangeAgencia As Integer
'Dim vMoney As String
'Dim vage As String
'
'Dim ApExcel As Variant
'Set ApExcel = CreateObject("Excel.application")
'
'    'Agrega un nuevo Libro
'    ApExcel.Workbooks.Add
'
'    'Poner Titulos
'    ApExcel.Cells(2, 2).Formula = "REPORTE CASTIGO DE CREDITOS - RECUPERACIONES DE " & Format(gdFecSis, "dd/MM/YYYY")
'    'ApExcel.cells(2, 2).Font.Size = 12
'    ApExcel.Range("B2", "I2").MergeCells = True
'    ApExcel.Cells(2, 2).Font.Bold = True
'    ApExcel.Cells(2, 2).HorizontalAlignment = 3
'
''    ApExcel.Cells(5, 2).Formula = "TIPO CAMBIO : " & ImpreFormat(pnTipCam, 5, 3)
''    ApExcel.Cells(5, 2).Font.Bold = True
'
'    ApExcel.Cells(8, 2).Formula = "Nº"
'    ApExcel.Cells(8, 3).Formula = "CREDITO"
'    ApExcel.Cells(8, 4).Formula = "CLIENTE"
'    ApExcel.Cells(8, 5).Formula = "SALDO K"
'    ApExcel.Cells(8, 6).Formula = "SALD INT COMP"
'    ApExcel.Cells(8, 7).Formula = "SALD INT MORA" '**
'    ApExcel.Cells(8, 8).Formula = "SALD GASTO"
'    ApExcel.Cells(8, 9).Formula = "INT COMP GEN"
'
'    ApExcel.Range("B7", "I8").Interior.Color = RGB(10, 190, 160) 'MADM 20110503
'    ApExcel.Range("B7", "I8").Font.Bold = True                   'MADM 20110503
'    ApExcel.Range("B7", "I8").HorizontalAlignment = 3            'MADM 20110503
'
'    I = 8
'    Rs.MoveFirst
'
'    Do While Not Rs.EOF
'    I = I + 1
'    ApExcel.Cells(I, 2).Formula = "COD AGENCIA : " & Rs!cCodAge
'    ApExcel.Cells(I, 2).Font.Bold = True
'    vage = Rs!cCodAge
'
'        Do While Rs!cCodAge = vage
'                I = I + 1
'                ApExcel.Cells(I, 2).Formula = "MONEDA : " & IIf(Mid(Trim(Rs!cCtaCod), 9, 1) = "1", "SOLES", "DOLARES")
'                ApExcel.Cells(I, 2).Font.Bold = True
'                vMoney = IIf(Mid(Trim(Rs!cCtaCod), 9, 1) = "1", "SOLES", "DOLARES")
'                     J = 0
'                    Do While IIf(Mid(Trim(Rs!cCtaCod), 9, 1) = "1", "SOLES", "DOLARES") = vMoney And Rs!cCodAge = vage
'                        J = J + 1
'                        I = I + 1
'                        ApExcel.Cells(I, 2).Formula = J
'                        ApExcel.Cells(I, 3).Formula = "'" & Rs!cCtaCod
'                        ApExcel.Cells(I, 4).Formula = "'" & Rs!nNombre
'                        ApExcel.Cells(I, 5).Formula = Rs!nSaldoCap
'                        ApExcel.Cells(I, 6).Formula = Rs!nSaldoIntCom
'                        ApExcel.Cells(I, 7).Formula = Rs!nSaldoIntMor
'                        ApExcel.Cells(I, 8).Formula = Rs!nSaldoGasto
'                        ApExcel.Cells(I, 9).Formula = Rs!nIntComGen
'
'                        ApExcel.Range("E" & Trim(Str(I)) & ":" & "I" & Trim(Str(I))).NumberFormat = "#,##0.00"
'                        'ApExcel.Range("B" & Trim(str(i)) & ":" & "W" & Trim(str(i))).Borders.LineStyle = 1
'                        ApExcel.Range("B" & Trim(Str(I)) & ":" & "I" & Trim(Str(I))).Borders.LineStyle = 1
'
'                        Rs.MoveNext
'                        If Rs.EOF Then
'                            Exit Do
'                        End If
'
'            Loop
'
'            I = I + 1
'            ApExcel.Cells(I, 5).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
'            ApExcel.Cells(I, 6).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
'            ApExcel.Cells(I, 7).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
'            ApExcel.Cells(I, 8).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
'            ApExcel.Cells(I, 9).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
'
'            I = I + 1
'
'            If Rs.EOF Then
'                Exit Do
'            End If
'
'        Loop
'    Loop
'
'    Rs.Close
'    Set Rs = Nothing
'
'    ApExcel.Cells.Select
'    ApExcel.Cells.EntireColumn.AutoFit
'    ApExcel.Columns("B:B").ColumnWidth = 6#
'    ApExcel.Range("B2").Select
'
'    Screen.MousePointer = 0
'
'    ApExcel.Visible = True
'    Set ApExcel = Nothing
End Function
