VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NColRecImpre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Cabecera de Impresiones de Reportes de Pignoraticio
Public Function nImprimeCabeceraReportes(ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, _
        ByVal psFechaSis As String, ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String _
        ) As String
        
  Dim lsCabe01 As String, lsCabe02 As String
  Dim lsCabe03 As String, lsCabe04 As String
  Dim lsCabRepo As String
  
  lsCabRepo = ""
  ' Cabecera 1
  lsCabe01 = FillText(Trim(UCase(psNomCmac)), 55, " ")
  lsCabe01 = lsCabe01 & Space(pnAnchoLinea - 55 - 25)
  lsCabe01 = lsCabe01 & "Pag.  : " & Str(pnPagina) & "  -  " & psCodUser & Chr(10)
  'lsCabe01 = lsCabe01 & IIf(pbCiereDia = True, IIf(VerifSiCierreDia(), "DC", "AC"), "") & chr$(10)
  ' Cabecera 2
  lsCabe01 = lsCabe01 & FillText(Trim(UCase(psNomAgencia)), 35, " ")
  lsCabe01 = lsCabe01 & Space(pnAnchoLinea - 35 - 25)
  lsCabe01 = lsCabe01 & "Fecha : " & Format(psFechaSis & " " & Time, "dd/mm/yyyy hh:mm") & Chr$(10)
  ' Titulo
  lsCabe02 = String(Int((pnAnchoLinea - Len(psTitulo)) / 2), " ") & psTitulo & Chr$(10)
  ' SubTitulo
  lsCabe03 = String(Int((pnAnchoLinea - Len(psSubTitulo)) / 2), " ") & psSubTitulo & Chr$(10)
  ' Comenta
  lsCabe04 = IIf(Len(psComenta) > 0, psComenta & Chr$(10), "")
  ' ***
  lsCabRepo = lsCabRepo & lsCabe01 & lsCabe02
  lsCabRepo = lsCabRepo & lsCabe03 & lsCabe04
  nImprimeCabeceraReportes = lsCabRepo
End Function

Public Function nPrintGastosRecuperaciones(ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, _
        ByVal psFecSis As String, ByVal psNomCliente As String, ByVal prGastos As ADODB.Recordset, _
        ByVal psMensaje As String) As String
    
Dim lsCadImp As String


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
            lsCadImp = lsCadImp & ImpreFormat(prGastos("nNroGastoCta"), 4, 0) & Space(2) & Format(prGastos("dAsigna"), "dd/mm/yyyy") & Space(2) & ImpreFormat(Mid(prGastos("cMotivoGasto"), 1, 30), 30) _
                        & Space(2) & ImpreFormat(prGastos("nMonto"), 12, 2) & ImpreFormat(prGastos("nMontoPagado"), 12, 2) & Space(2)
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
                            
        lsCadImp = lsCadImp & " " & ImpreFormat(psMensaje, 40) _
                            & Chr(10)
        
        nPrintGastosRecuperaciones = lsCadImp
End Function

Public Function nPrintReciboPagoCredRecupLima(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, _
        ByVal pnMontoPago As Currency, _
        ByVal psCodUser As String, ByVal pnCapital As Currency, _
        ByVal pnInteres As Currency, ByVal pnMora As Currency, _
        ByVal pnGastos As Currency, Optional ByVal psMensaje As String = "") As String
Dim lsCadImp As String
Dim vEspacio As Integer


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
        lsCadImp = lsCadImp & " " & ImpreFormat(psNomAgencia, 26, 0) & Format(psFecTransac, "dd/mm/yyyy hh:mm:ss") & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & " Credito   : " & Space(9) & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & " CLIENTE   : " & ImpreFormat(psNombCliente, 27) & Chr(10)
                            
        lsCadImp = lsCadImp & Chr(10) & Chr(10)
        lsCadImp = lsCadImp & " ==========  Comprobante Pago ========== " & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & " CAPITAL          :" & ImpreFormat(pnCapital, 18, , True) & Chr(10)
        lsCadImp = lsCadImp & " INTERES COMP.    :" & ImpreFormat(pnInteres, 18, , True) & Chr(10)
        lsCadImp = lsCadImp & " MORA             :" & ImpreFormat(pnMora, 18, , True) & Chr(10)
        lsCadImp = lsCadImp & " GASTOS           :" & ImpreFormat(pnGastos, 18, , True) & Chr(10)
        lsCadImp = lsCadImp & " NETO A PAGAR     :" & ImpreFormat(pnMontoPago, 18, , True) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & " ===================================== " & Chr(10)
        lsCadImp = lsCadImp & " " & Format(psCodUser, "@@@@") & Chr(10)
        lsCadImp = lsCadImp & " " & ImpreFormat(psMensaje, 40) & Chr(12)
        nPrintReciboPagoCredRecupLima = lsCadImp
End Function

'Procedimiento de impresión del recibo de renovación
Public Function nPrintReciboPagoCredRecup(ByVal psNomAgencia As String, ByVal psFecTransac As String, _
        ByVal psCodCta As String, ByVal psNombCliente As String, _
        ByVal pnMontoPago As Currency, _
        ByVal psCodUser As String, Optional ByVal psMensaje As String = "", Optional ByVal pnITF As Double = 0#) As String
Dim lsCadImp As String
Dim vEspacio As Integer

lsCadImp = lsCadImp & Chr$(27) & Chr$(64)

lsCadImp = lsCadImp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
lsCadImp = lsCadImp & Chr$(27) & Chr$(67) & Chr$(22)  'Longitud de página a 22 líneas'
lsCadImp = lsCadImp & Chr$(27) & Chr$(77)   'Tamaño 10 cpi
lsCadImp = lsCadImp & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
lsCadImp = lsCadImp & Chr$(27) + Chr$(18) ' cancela condensada
lsCadImp = lsCadImp & Chr$(27) + Chr$(72) ' desactiva negrita

    vEspacio = 9
        lsCadImp = "" & Chr(10) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)      'Activa Negrita
        lsCadImp = lsCadImp & " Recuperaciones - Pago " & Space(10) & Space(10 + vEspacio) _
                            & " Recuperaciones - Pago " & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)     'Desactiva Negrita
        lsCadImp = lsCadImp & " " & ImpreFormat(psNomAgencia, 26, 0) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Space(vEspacio) _
                            & " " & ImpreFormat(psNomAgencia, 26, 0) & Format(psFecTransac, "dd/mm/yyyy hh:mm") & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        lsCadImp = lsCadImp & " Credito   : " & Space(9) & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Space(vEspacio) _
                            & " Credito   : " & Space(9) & Mid(psCodCta, 1, 5) & "-" & Mid(psCodCta, 6) & Chr(10)
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & " CLIENTE   : " & ImpreFormat(psNombCliente, 27) & Space(vEspacio) _
                            & " CLIENTE   : " & ImpreFormat(psNombCliente, 27) & Chr(10)
                            
        lsCadImp = lsCadImp & Chr(10) & Chr(10)
        
        lsCadImp = lsCadImp & " ==========  Comprobante Pago ========== " & Space(vEspacio) _
                            & " ==========  Comprobante Pago ========== " & Chr(10)
                            
                            
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)    'Activa Negrita
        
        lsCadImp = lsCadImp & " I.T.F.           :" & ImpreFormat(pnITF, 18, , True) & Space(vEspacio) _
                            & " I.T.F.           :" & ImpreFormat(pnITF, 18, , True) & Chr(10)
        lsCadImp = lsCadImp & " PAGO             :" & ImpreFormat(pnMontoPago, 18, , True) & Space(vEspacio) _
                            & " PAGO             :" & ImpreFormat(pnMontoPago, 18, , True) & Chr(10)
        lsCadImp = lsCadImp & " NETO A PAGAR     :" & ImpreFormat(pnMontoPago + pnITF, 18, , True) & Space(vEspacio) _
                            & " NETO A PAGAR     :" & ImpreFormat(pnMontoPago + pnITF, 18, , True) & Chr(10)
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        
        lsCadImp = lsCadImp & " ===================================== " & Space(vEspacio) _
                            & " ===================================== " & Chr(10)
        lsCadImp = lsCadImp & " " & Format(psCodUser, "@@@@") & Space(30) & Space(vEspacio) _
                            & " " & Format(psCodUser, "@@@@") & Chr(10)
        
        lsCadImp = lsCadImp & " " & ImpreFormat(psMensaje, 40) & Space(vEspacio) _
                            & " " & ImpreFormat(psMensaje, 40) & Chr(10)

        nPrintReciboPagoCredRecup = lsCadImp
End Function
