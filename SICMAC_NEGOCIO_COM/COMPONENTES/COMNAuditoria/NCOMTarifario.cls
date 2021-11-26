VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMTarifario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim csNomCMAC As String
Dim csNomAgencia As String
Dim csCodUser As String
Dim csFechaSis As String

Public Sub inicio(ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psFechaSis As String)
    csNomCMAC = psNomCmac
    csNomAgencia = psNomAgencia
    csCodUser = psCodUser
    csFechaSis = psFechaSis
End Sub

Public Function DarTarifarioGastos(ByVal psProducto As String, ByRef psMensaje As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMTarifario
    Set objDAuditoria = New COMDAuditoria.DCOMTarifario
    Set DarTarifarioGastos = objDAuditoria.CargarTarifarioGastos(psProducto, psMensaje)
End Function

Public Function ReporteTarifarioGastos(ByVal lsProducto As String, ByVal psProducto As String, _
        Optional ByRef psMensaje As String, Optional psImpresora As Impresoras = gEPSON) As String
        
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim ofun As New COMFunciones.FCOMImpresion

Dim P As String

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

    Set lrDataRep = DarTarifarioGastos(psProducto, psMensaje)
    
    If psMensaje = "" Then
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("TARIFARIO DE GASTOS ", "VIGENTES AL " & Format(csFechaSis, "dd/mm/yyyy"), lnPage, 150, "")

        lnIndice = 0:  lnLineas = 7

        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                
                Dim lsAplMon As String
                Dim lsAplPro As String
                Dim lsTpoValor As String
                
                If !cAplicaMonto = "D" Then
                    lsAplMon = "Préstamo"
                End If
                
                If !cAplicaMonto = "C" Then
                    lsAplMon = "Cuota"
                End If
                
                If !cAplicaMonto = "H" Then
                    lsAplMon = "Garantia"
                End If
                
                If !cAplicaMonto = "S" Then
                lsAplMon = "Saldo Cap"
                End If
                
                If !cAplicaMonto = "T" Then
                    lsAplMon = "Interés Comp"
                End If
                
                If !cAplicaMonto = "I" Then
                lsAplMon = "Saldo Cap Ant"
                End If
                
                If IsNull(!cAplicaProceso) Or !cAplicaProceso = "DE" Then
                lsAplPro = "Aprob."
                End If
                
                If !cAplicaProceso = "PA" Then
                lsAplPro = "Pagos"
                End If
                
                If !cAplicaProceso = "CA" Then
                lsAplPro = "Canal"
                End If
                    
                If !cAplicaProceso = "RE" Then
                lsAplPro = "Refin."
                End If
                
                If !cAplicaProceso = "PP" Then
                lsAplPro = "Pre-Pag"
                End If
                
                If !cAplicaProceso = "CD" Then
                lsAplPro = "Cierre Día"
                End If
                
                If !cAplicaProceso = "MA" Then
                lsAplPro = "Asig. Manual"
                End If
                
                If !nTpoValor = "1" Then
                lsTpoValor = "Valor"
                End If
                
                If !nTpoValor = "2" Then
                lsTpoValor = "Porcentaje"
                End If
                
                If !nTpoValor = "3" Then
                lsTpoValor = "Personalizado"
                End If
            
                lsCadImp = lsCadImp & ofun.ImpreFormat(!nPrdConceptoCod, 5, 0) & Space(1) & ofun.ImpreFormat(!cDescripcion, 42, 0) & Space(1) & ofun.ImpreFormat(!nInicial, 4, 2) & Space(1) & ofun.ImpreFormat(!nFinal, 7, 2) & Space(1) & ofun.ImpreFormat(!nMontoMin, 4, 2) & Space(1) & ofun.ImpreFormat(!nMontoMax, 7, 2) & Space(1) & ofun.ImpreFormat(!nValor, 4, 4) & Space(2) & ofun.ImpreFormat(IIf(!nMoneda = "1", "NS", "USD"), 2, 0) & Space(2) & ofun.ImpreFormat(lsAplMon, 13, 0) & Space(2) & ofun.ImpreFormat(lsAplPro, 13, 0) & Space(2) & ofun.ImpreFormat(lsTpoValor, 10, 0)
                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea

                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If

                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS ", "VIGENTES AL " & Format(csFechaSis, "dd/mm/yyyy"), lnPage, 150, "")
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "Total de Tarifario de Gastos: " & ofun.ImpreFormat(lnIndice, 5, 0)

        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina

    End If
    ReporteTarifarioGastos = lsCadBuffer
Set oFunI = Nothing
End Function

Public Function nRepoCabecera(ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String, _
        Optional ByVal psImpresora As Impresoras = gEPSON) As String
        
Dim lsCadImp As String
Dim loImprimeCab As COMNColoCPig.NCOMColPImpre
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

    Set loImprimeCab = New COMNColoCPig.NCOMColPImpre
        lsCadImp = loImprimeCab.nImprimeCabeceraReportes(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)
    Set loImprimeCab = Nothing
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & String(pnAnchoLinea, "-") & oFunI.gPrnSaltoLinea
    
            lsCadImp = lsCadImp & "Cod.  Descripcion                                   R.Ini.  R. Fin.   M. Min.  M. Max.    Valor  Mon. Apl. Monto     Apl. Proceso   Tpo Valor" & oFunI.gPrnSaltoLinea
    lsCadImp = lsCadImp & String(pnAnchoLinea, "-") & oFunI.gPrnSaltoLinea
nRepoCabecera = lsCadImp
Set oFunI = Nothing
End Function

