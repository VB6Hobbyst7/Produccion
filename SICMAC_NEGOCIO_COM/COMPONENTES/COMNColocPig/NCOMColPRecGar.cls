VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMColPRecGar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mnEjecutaBatch As Integer
Dim mbTrans As Boolean

Public Sub nRecGarGrabaDatosPreparaCredPignoraticio(ByVal psTpoProceso As String, ByVal psNroProceso As String, _
            ByVal pnRGEstado As Integer, ByVal psFecProceso As String, ByVal psProcesoCadaAgencia As String, _
            Optional ByVal pnPrec14 As Currency = -1, Optional ByVal pnPrec16 As Currency = -1, Optional ByVal pnPrec18 As Currency = -1, _
            Optional ByVal pnPrec21 As Currency = -1, Optional ByVal pnPrecioOro As Currency = -1, Optional ByVal pnTipoCambio As Currency = -1, _
            Optional ByVal psFecParaAdjud As String = "@", _
            Optional pbEjecBatch As Boolean = False, _
            Optional ByVal pnTasaCustodiaVencida As Currency = -1, Optional ByVal pnTasaPreparacionRemate As Currency = -1, _
            Optional ByVal pnTasaImpuesto As Currency = -1, Optional ByVal pnFactorPrecioVenta As Currency = -1, Optional bActualiza As Boolean = False, Optional psFechaCorte As String = "")
'** Actualiza ColocPigRecGar
'** Actualiza ColocPigRGPrecio
'************************************

Dim lsSQL As String
Dim loRegPig As New COMDColocPig.DCOMColPActualizaBD
Dim sSql As String

'On Error GoTo ErrorModPig

    '*** Cambio para guardar la fecha de Corte de Remate -- AVMM 09-06-2007 ****
    If psFechaCorte <> "" Then
        Dim oCon As COMConecta.DCOMConecta
        Set oCon = New COMConecta.DCOMConecta
            If oCon.AbreConexion Then
               sSql = "Update Constante Set cConsDescripcion='" & Format(psFechaCorte, "dd/mm/yyyy") & "' Where nConsCod='3208' and nConsValor='" & Trim(Right(psProcesoCadaAgencia, 2)) & "' "
               oCon.Ejecutar sSql
            End If
        Set oCon = Nothing
    End If
    '***************************************************************************
    
    '** Actualiza ColocPigRecGar
    
    If bActualiza = False Then
        If psTpoProceso = "A" Then
            Call loRegPig.dUpdateColocPigRecGar(psTpoProceso, psNroProceso, psProcesoCadaAgencia, psFecProceso, pnRGEstado, psFecParaAdjud, True)
        Else
            Call loRegPig.dUpdateColocPigRecGar(psTpoProceso, psNroProceso, psProcesoCadaAgencia, psFecProceso, pnRGEstado, , True)
        End If
        '** Elimina ColocPigRGPrecio
        Call loRegPig.dDeleteColocPigRGPrecio(psTpoProceso, psNroProceso, psProcesoCadaAgencia, True)
            
        '** Actualiza ColocPigRGPrecio
        If pnPrec14 > 0 Then
            Call loRegPig.dInsertColocPigRGPrecio(psTpoProceso, psNroProceso, "14", pnPrec14, psProcesoCadaAgencia, True)
        End If
        If pnPrec16 > 0 Then
            Call loRegPig.dInsertColocPigRGPrecio(psTpoProceso, psNroProceso, "16", pnPrec16, psProcesoCadaAgencia, True)
        End If
        If pnPrec18 > 0 Then
            Call loRegPig.dInsertColocPigRGPrecio(psTpoProceso, psNroProceso, "18", pnPrec18, psProcesoCadaAgencia, True)
        End If
        If pnPrec21 > 0 Then
            Call loRegPig.dInsertColocPigRGPrecio(psTpoProceso, psNroProceso, "21", pnPrec21, psProcesoCadaAgencia, True)
        End If
        If pnPrecioOro > 0 Then
            Call loRegPig.dInsertColocPigRGPrecio(psTpoProceso, psNroProceso, "PO", pnPrecioOro, psProcesoCadaAgencia, True)
        End If
        If pnTipoCambio > 0 Then
            Call loRegPig.dInsertColocPigRGPrecio(psTpoProceso, psNroProceso, "TC", pnTipoCambio, psProcesoCadaAgencia, True)
        End If
        If pnTasaCustodiaVencida >= 0 Then
            Call loRegPig.dInsertColocPigRGPrecio(psTpoProceso, psNroProceso, gConsColPTasaCustodiaVencida, pnTasaCustodiaVencida, psProcesoCadaAgencia, True)
        End If
        If pnTasaPreparacionRemate >= 0 Then
            Call loRegPig.dInsertColocPigRGPrecio(psTpoProceso, psNroProceso, gConsColPTasaPreparaRemate, pnTasaPreparacionRemate, psProcesoCadaAgencia, True)
        End If
        If pnTasaImpuesto >= 0 Then
            Call loRegPig.dInsertColocPigRGPrecio(psTpoProceso, psNroProceso, gConsColPTasaImpuesto, pnTasaImpuesto, psProcesoCadaAgencia, True)
        End If
        If pnFactorPrecioVenta >= 0 Then
            Call loRegPig.dInsertColocPigRGPrecio(psTpoProceso, psNroProceso, gConsColPFactorPrecioBaseRemate, pnFactorPrecioVenta, psProcesoCadaAgencia, True)
        End If
    Else
        If pnTasaCustodiaVencida >= 0 Then
            Call loRegPig.dUpdateColocPigRGPrecio(psTpoProceso, psNroProceso, gConsColPTasaCustodiaVencida, pnTasaCustodiaVencida, False)
        End If
        If pnTasaPreparacionRemate >= 0 Then
            Call loRegPig.dUpdateColocPigRGPrecio(psTpoProceso, psNroProceso, gConsColPTasaPreparaRemate, pnTasaPreparacionRemate, False)
        End If
        If pnTasaImpuesto >= 0 Then
            Call loRegPig.dUpdateColocPigRGPrecio(psTpoProceso, psNroProceso, gConsColPTasaImpuesto, pnTasaImpuesto, False)
        End If
        If pnFactorPrecioVenta >= 0 Then
            Call loRegPig.dUpdateColocPigRGPrecio(psTpoProceso, psNroProceso, gConsColPFactorPrecioBaseRemate, pnFactorPrecioVenta, False)
        End If
    
    End If


    mnEjecutaBatch = loRegPig.dEjecutaBatch
Set loRegPig = Nothing

Exit Sub
ErrorModPig:
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Sub

Public Function nObtieneNroUltimoProceso(ByVal psTpoProceso As String, ByVal psAgencia As String, ByRef psmensaje As String, Optional ByVal psFecSis As String = "@") As String
'Obtiene Nro Ultimo Remate Credito Pignoraticio
'*** PEAC 20190719 - Se agregó "psFecSis"

Dim loValida As New COMDColocPig.DCOMColPFunciones
Dim lrs As New ADODB.Recordset
Dim lsUltRem As String
Dim lsSQL As String
Dim lsAgencia As String



On Error GoTo dError

     '**APRI 20170516 ********************************************************
    'lsSQL = "Exec sp_ObtieneNroUltimoProceso '" & psTpoProceso & "','" & psAgencia & "'"
    lsSQL = "Exec sp_ObtieneNroUltimoProceso '" & psTpoProceso & "','" & psAgencia & "','" & psFecSis & "' " '*** PEAC 20190718
    '************************************************************************

    Set lrs = loValida.dObtieneRecordSet(lsSQL)
    Set loValida = Nothing
    If lrs.BOF And lrs.EOF Then
        psmensaje = "No se tiene Información del Ultimo Proceso"
    Else
        lsUltRem = lrs!cNroProceso
    End If
    Set lrs = Nothing

    nObtieneNroUltimoProceso = lsUltRem
    Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
End Function

Public Function nObtieneNroProcesoFecha(ByVal psTpoProceso As String, ByVal psAgencia As String, ByVal psFecha As String, ByRef psmensaje As String) As String
'Obtiene Nro Ultimo Remate Credito Pignoraticio
Dim loValida As New COMDColocPig.DCOMColPFunciones
Dim lrs As New ADODB.Recordset
Dim lsUltRem As String
Dim lsSQL As String
Dim lsAgencia As String

On Error GoTo dError

lsSQL = "SELECT RG.cTpoProceso, RG.cNroProceso " _
    & " FROM ColocPigRecGar RG  WHERE cTpoProceso ='" & psTpoProceso _
    & "' AND cCodAge ='" & psAgencia & "' ORDER BY cNroProceso DESC "

'Set lrs = New Recordset
   
    Set lrs = loValida.dObtieneRecordSet(lsSQL)
    Set loValida = Nothing
    If lrs.BOF And lrs.EOF Then
        psmensaje = "No se tiene Información del Ultimo Proceso"
    Else
        lsUltRem = lrs!cNroProceso
    End If
Set lrs = Nothing

nObtieneNroProcesoFecha = lsUltRem
Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function

Public Function nObtieneDatosProcesoRGCredPig(ByVal psTpoProceso As String, ByVal psNroProceso As String, ByVal psAgencia As String, ByRef psmensaje As String) As ADODB.Recordset
'Obtiene Datos de Remate Credito Pignoraticio
Dim loValida As COMDColocPig.DCOMColPFunciones
Dim lrs As New ADODB.Recordset
Dim lsSQL As String
Dim lsAgencia As String

On Error GoTo dError

'    If psTpoProceso = "A" Then
'        'lsSQL = "SELECT RG.cTpoProceso, RG.cNroProceso, RG.dProceso, RG.nRGEstado, RG.dParaAdjudicar, " _
'            & " sum( case when cTpoPrecio ='PO' then nPrecio else 0  end ) as  nPrecioOro ," _
'            & " sum( case when cTpoPrecio ='TC' then nPrecio else 0  end ) as  nTipoCambio " _
'            & " FROM ColocPigRecGar RG LEFT JOIN ColocPigRGPrecio RGPrec " _
'            & " ON RG.cTpoProceso = RGPrec.cTpoProceso AND RG.cNroProceso = RGPrec.cNroProceso  " _
'            & " WHERE RG.cTpoProceso ='" & psTpoProceso & "' and RG.cNroProceso ='" & psNroProceso & "' AND RG.cCodAge ='" & psAgencia & "'  " _
'            & " GROUP BY RG.cTpoProceso, RG.cNroProceso, RG.dProceso, RG.nRGEstado, RG.dParaAdjudicar "
'        lsSQL = "SELECT RG.cTpoProceso, RG.cNroProceso, RG.dProceso, RG.nRGEstado, RG.dParaAdjudicar, " _
'            & " sum( case when cTpoPrecio ='14' then nPrecio else 0  end ) as  nPrecioK14 ," _
'            & " sum( case when cTpoPrecio ='16' then nPrecio else 0  end ) as  nPrecioK16 ," _
'            & " sum( case when cTpoPrecio ='18' then nPrecio else 0  end ) as  nPrecioK18 ," _
'            & " sum( case when cTpoPrecio ='21' then nPrecio else 0  end ) as  nPrecioK21 " _
'            & " FROM ColocPigRecGar RG LEFT JOIN ColocPigRGPrecio RGPrec " _
'            & " ON RG.cTpoProceso = RGPrec.cTpoProceso AND RG.cNroProceso = RGPrec.cNroProceso  " _
'            & " WHERE RG.cTpoProceso ='" & psTpoProceso & "' and RG.cNroProceso ='" & psNroProceso & "' AND RG.cCodAge ='" & psAgencia & "' " _
'            & " GROUP BY RG.cTpoProceso, RG.cNroProceso, RG.dProceso, RG.nRGEstado, RG.dParaAdjudicar "
'    Else
'        lsSQL = "SELECT RG.cTpoProceso, RG.cNroProceso, RG.dProceso, RG.nRGEstado, " _
'            & " sum( case when cTpoPrecio ='14' then nPrecio else 0  end ) as  nPrecioK14 ," _
'            & " sum( case when cTpoPrecio ='16' then nPrecio else 0  end ) as  nPrecioK16 ," _
'            & " sum( case when cTpoPrecio ='18' then nPrecio else 0  end ) as  nPrecioK18 ," _
'            & " sum( case when cTpoPrecio ='21' then nPrecio else 0  end ) as  nPrecioK21 " _
'            & " FROM ColocPigRecGar RG LEFT JOIN ColocPigRGPrecio RGPrec " _
'            & " ON RG.cTpoProceso = RGPrec.cTpoProceso AND RG.cNroProceso = RGPrec.cNroProceso  " _
'            & " WHERE RG.cTpoProceso ='" & psTpoProceso & "' and RG.cNroProceso ='" & psNroProceso & "' AND RG.cCodAge ='" & psAgencia & "' " _
'            & " GROUP BY RG.cTpoProceso, RG.cNroProceso, RG.dProceso, RG.nRGEstado "
'    End If
    
    '*** PEAC 20080714
    lsSQL = " Exec stp_sel_ObtieneDatosProcesoRGCredPig '" & psTpoProceso & "','" & psNroProceso & "','" & psAgencia & "'"
    
    Set loValida = New COMDColocPig.DCOMColPFunciones
        Set lrs = loValida.dObtieneRecordSet(lsSQL)
    Set loValida = Nothing
    If lrs.BOF And lrs.EOF Then
        psmensaje = "No se encuentra información del Ultimo Proceso"
    Else
        Set nObtieneDatosProcesoRGCredPig = lrs
    End If
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function

'Imprime Aviso de Vencimiento de Remate
Public Function nRemImprimeAvisoVencimiento(ByVal psTextoCarta As String, ByVal psFecRemate As String, _
       ByVal pnDiasAtrasoCartaVenc As Integer, ByVal psAgencia As String, ByVal pnPrevioMax As Long, _
       ByVal psFechaSis As String, ByVal pnCostoCorrespondencia As Double, ByVal psCodAge As String, _
       ByVal psCodUsu As String, ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
Dim LALISTA() As String
Dim lnFilas As Long, lnItem As Long
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCartaModelo As String
Dim loRegPig As COMDColocPig.DCOMColPActualizaBD
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lnMovNro As Long
Dim lsOpeCod As String
Dim lsDireAge As String '*** PEAC 20190718
Dim lsFecRemateReal As String

Dim ofun As New COMFunciones.FCOMFechas
Dim ofcar As New COMFunciones.FCOMCadenas
lsOpeCod = "121700"
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


On Error GoTo dError

'ARCV 23-06-2007
lsFecRemateReal = psFecRemate
'-----
    '*****************************************************
    'Obtener Fecha de Corte
    Dim oCon As New COMConecta.DCOMConecta
    Dim rs1 As ADODB.Recordset
    Dim rs2 As ADODB.Recordset '*** PEAC 20190718
    Set oCon = New COMConecta.DCOMConecta
        If oCon.AbreConexion Then
           Set rs1 = New ADODB.Recordset
           lsSQL = "Select cConsDescripcion From Constante Where nConsCod='3208' and nConsValor='" & Trim(psAgencia) & "' "
           Set rs1 = oCon.CargaRecordSet(lsSQL)
           If Not (rs1.EOF And rs1.BOF) Then
             psFecRemate = Format(rs1("cConsDescripcion"), "mm/dd/yyyy")
           End If
           Set rs1 = Nothing
           
            '*** PEAC 20181227 - obtiene direccion de la agencia actual
           Set rs2 = New ADODB.Recordset
           lsSQL = "select rtrim(cagedireccion) cagedireccion from agencias where cagecod = '" & Trim(psAgencia) & "' "
           Set rs2 = oCon.CargaRecordSet(lsSQL)
           If Not (rs2.EOF And rs2.BOF) Then
             lsDireAge = rs2("cagedireccion")
           End If
           Set rs2 = Nothing
           
        End If
    Set oCon = Nothing
    '*******************************************************
    

    ldFecAviso = DateAdd("d", -pnDiasAtrasoCartaVenc, psFecRemate)


    lsSQL = "SELECT Per.cPersNombre, Per.cPersDireccDomicilio, Per.cPersTelefono, Per.cPersDireccUbiGeo, " _
        & " Ubigeo.cUbiGeoDescripcion, C.cCtaCod , C.dVenc, P.nPrdEstado, nSaldo" _
        & " FROM Producto P JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod " _
        & " JOIN ProductoPersona PP ON c.cCtaCod = PP.cCtaCod " _
        & " JOIN Persona Per ON Per.cPersCod = PP.cPersCod " _
        & " LEFT JOIN UbicacionGeografica Ubigeo ON Per.cPersDireccUbiGeo = Ubigeo.cUbiGeoCod " _
        & " WHERE DATEDIFF(dd, C.dVenc, '" & Format(ldFecAviso, "mm/dd/yyyy") & "') >= 0  AND " _
        & " P.nPrdEstado in (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ") " _
        & " And Substring(P.cCtaCod,4,2) ='" & psAgencia & "' ORDER BY Per.cPersNombre "

    '*** PEAC 20190718
    lsSQL = "exec stp_sel_AvisoVencimientoPigno '" & Format(ldFecAviso, "yyyymmdd") & "','" & psAgencia & "' "

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing


    If lrDataRep Is Nothing Then
        psmensaje = " No existen Créditos Pignoraticios para Aviso de Vencimiento "
        Exit Function
    ElseIf lrDataRep.BOF And lrDataRep.EOF Then
        psmensaje = " No existen Créditos Pignoraticios para Aviso de Vencimiento "
        Exit Function
    Else
        If pnCostoCorrespondencia > 0 Then
            'Asigno el gasto a los creditos
            Set loContFunct = New COMNContabilidad.NCOMContFunciones
                lsMovNro = loContFunct.GeneraMovNro(psFechaSis, psCodAge, psCodUsu)
            Set loContFunct = Nothing
        
            lsFechaHoraGrab = ofun.fgFechaHoraGrab(lsMovNro)
            Set loRegPig = New COMDColocPig.DCOMColPActualizaBD
        
            '** Inserta Mov
            Call loRegPig.dInsertMov(lsMovNro, lsOpeCod, "Costo Correspondencia Cred Pign", gMovEstContabNoContable, gMovFlagVigente, False)
            ' Obtiene nMovNro
            lnMovNro = loRegPig.dGetnMovNro(lsMovNro)
            
            Do While Not lrDataRep.EOF
                '** Inserta MovCol
                Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, lrDataRep!cCtaCod, 1, pnCostoCorrespondencia, 0, "", 0, 0, 0, False)
            
                '** Inserta MovColDet -  Costo Duplicado
                Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, lrDataRep!cCtaCod, 1, 2215, 1, pnCostoCorrespondencia, False)
                
                
                'Actualiza ColocCalendDetPig
                lsSQL = "UPDATE ColocCalendDetPig Set nMonto = nMonto + " & pnCostoCorrespondencia & _
                        " WHERE cCtaCod ='" & lrDataRep!cCtaCod & "' And nNroCalen = 1 And nColocCalendApl = 0 " & _
                        " And nCuota = 1 And nPrdConceptoCod = 2215 "
                        
                loRegPig.dEjecutar lsSQL
                lrDataRep.MoveNext
            Loop
            Set loRegPig = Nothing
        End If
        lnFilas = 0
        ReDim LALISTA(lrDataRep.RecordCount, 5)
        lrDataRep.MoveFirst
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnFilas = lnFilas + 1
                If !cPersNombre = LALISTA(lnFilas - 1, 0) Then
                    LALISTA(lnFilas - 1, 3) = LALISTA(lnFilas - 1, 3) & " - " & !cCtaCod
                    LALISTA(lnFilas - 1, 4) = LALISTA(lnFilas - 1, 4) & " - " & Format(!dVenc, "dd/mm/yyyy")
                    lnFilas = lnFilas - 1
                Else
                    LALISTA(lnFilas, 0) = !cPersNombre
                    LALISTA(lnFilas, 1) = !cPersDireccDomicilio & " - " & !cUbiGeoDescripcion '  ClienteZona(!cCodZon) & " - " & ClienteCiudad(!cCodZon)
                    LALISTA(lnFilas, 2) = !cPersTelefono & ""
                    LALISTA(lnFilas, 3) = !cCtaCod
                    LALISTA(lnFilas, 4) = Format(!dVenc, "dd/mm/yyyy")
                    LALISTA(lnFilas, 5) = Format(!nSaldo, "0.00")
                End If
                .MoveNext
            Loop
        End With
        lrDataRep.Close
        Set lrDataRep = Nothing
    End If
    'Llena cartas
    For lnItem = 1 To lnFilas
        lsCartaModelo = psTextoCarta
        lsCartaModelo = Replace(lsCartaModelo, "<<FECHA>>", Format(psFechaSis, "dddd,d mmmm yyyy"), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<NOMBRE>>", ofcar.PstaNombre(LALISTA(lnItem, 0), False), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<FECVENCIM>>", LALISTA(lnItem, 4), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<CONTRATO>>", LALISTA(lnItem, 3), , 2, vbTextCompare)
        
        lsCartaModelo = Replace(lsCartaModelo, "<<DIREAGE>>", lsDireAge, , 1, vbTextCompare)  '*** PEAC 20190718
        
        'ARCV 23-06-2007
        'lsCartaModelo = Replace(lsCartaModelo, "<<FECREMATE>>", Format(CDate(psFecRemate), "dd.MM.yyyy"), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<FECREMATE>>", Format(CDate(lsFecRemateReal), "dd.MM.yyyy"), , 1, vbTextCompare)
        '----------
        lsCartaModelo = Replace(lsCartaModelo, "<<DIRECCION>>", LALISTA(lnItem, 1), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<TELEFONO>>", LALISTA(lnItem, 2), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<MONTO>>", LALISTA(lnItem, 5), , 1, vbTextCompare)
        
        lsCadImp = lsCadImp & lsCartaModelo & oFunI.gPrnSaltoPagina
        
    Next lnItem
    nRemImprimeAvisoVencimiento = lsCadImp

Exit Function
   
dError:
    Err.Raise Err.Number, "Obtiene Datos Imprime Listado Aviso Vencimiento  <<nImprimeAvisoVencimiento>>", Err.Description
    
End Function

'Imprime Aviso de Remate
Public Function nRemImprimeAvisoRemate(ByVal psTextoCarta As String, ByVal psFecRemate As String, _
       ByVal pnDiasVctoParaRemate As Integer, ByVal psAgencia As String, ByVal pnPrevioMax As Long, _
       ByVal psFechaSis As String, ByVal pnCostoCorrespondencia As Double, ByVal psCodAge As String, _
       ByVal psCodUsu As String, ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
Dim LALISTA() As String
Dim lnFilas As Long, lnItem As Long
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCartaModelo As String

Dim loRegPig As COMDColocPig.DCOMColPActualizaBD

Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lnMovNro As Long
Dim lsOpeCod As String
lsOpeCod = "121700"

Dim ofun As New COMFunciones.FCOMFechas
Dim ofcar As New COMFunciones.FCOMCadenas
'On Error GoTo dError
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

    ldFecAviso = DateAdd("d", -pnDiasVctoParaRemate, psFecRemate)

    lsSQL = "SELECT Per.cPersNombre, Per.cPersDireccDomicilio, Per.cPersTelefono, Per.cPersDireccUbiGeo, " _
        & " Ubigeo.cUbiGeoDescripcion, C.cCtaCod , C.dVenc, P.nPrdEstado" _
        & " FROM Producto P JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod " _
        & " JOIN ProductoPersona PP ON c.cCtaCod = PP.cCtaCod " _
        & " JOIN Persona Per ON Per.cPersCod = PP.cPersCod " _
        & " LEFT JOIN UbicacionGeografica Ubigeo ON Per.cPersDireccUbiGeo = Ubigeo.cUbiGeoCod " _
        & " WHERE DATEDIFF(dd, C.dVenc, '" & Format(ldFecAviso, "mm/dd/yyyy") & "') >= 0  AND " _
        & " P.nPrdEstado in (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ") " _
        & " And Substring(P.cCtaCod,4,2) ='" & psAgencia & "' ORDER BY Per.cPersNombre "

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing


    If lrDataRep Is Nothing Then
        psmensaje = " No existen Créditos Pignoraticios para Aviso de Remate "
        Exit Function
    ElseIf lrDataRep.BOF And lrDataRep.EOF Then
        psmensaje = " No existen Créditos Pignoraticios para Aviso de Remate "
        Exit Function
    Else
    
        If pnCostoCorrespondencia > 0 Then
            'Asigno el gasto a los creditos
            Set loContFunct = New COMNContabilidad.NCOMContFunciones
                lsMovNro = loContFunct.GeneraMovNro(psFechaSis, psCodAge, psCodUsu)
            Set loContFunct = Nothing
        
            lsFechaHoraGrab = ofun.fgFechaHoraGrab(lsMovNro)
            Set loRegPig = New COMDColocPig.DCOMColPActualizaBD
        
            '** Inserta Mov
            Call loRegPig.dInsertMov(lsMovNro, lsOpeCod, "Costo Correspondencia Remat Cred Pign", gMovEstContabNoContable, gMovFlagVigente, False)
            ' Obtiene nMovNro
            lnMovNro = loRegPig.dGetnMovNro(lsMovNro)
            
            Do While Not lrDataRep.EOF
                '** Inserta MovCol
                Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, lrDataRep!cCtaCod, 1, pnCostoCorrespondencia, 0, "", 0, 0, 0, False)
            
                '** Inserta MovColDet -  Costo Duplicado
                Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, lrDataRep!cCtaCod, 1, 2215, 1, pnCostoCorrespondencia, False)
                
                
                'Actualiza ColocCalendDetPig
                lsSQL = "UPDATE ColocCalendDetPig Set nMonto = nMonto + " & pnCostoCorrespondencia & _
                        " WHERE cCtaCod ='" & lrDataRep!cCtaCod & "' And nNroCalen = 1 And nColocCalendApl = 0 " & _
                        " And nCuota = 1 And nPrdConceptoCod = 2215 "
                        
                loRegPig.dEjecutar lsSQL
                lrDataRep.MoveNext
            Loop
            Set loRegPig = Nothing
        End If
    
        lnFilas = 0
        
        ReDim LALISTA(lrDataRep.RecordCount, 5)
        lrDataRep.MoveFirst
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnFilas = lnFilas + 1
                If !cPersNombre = LALISTA(lnFilas - 1, 0) Then
                    LALISTA(lnFilas - 1, 3) = LALISTA(lnFilas - 1, 3) & " - " & !cCtaCod
                    LALISTA(lnFilas - 1, 4) = LALISTA(lnFilas - 1, 4) & " - " & Format(!dVenc, "dd/mm/yyyy")
                    lnFilas = lnFilas - 1
                Else
                    LALISTA(lnFilas, 0) = !cPersNombre
                    LALISTA(lnFilas, 1) = !cPersDireccDomicilio & " - " & !cUbiGeoDescripcion '  ClienteZona(!cCodZon) & " - " & ClienteCiudad(!cCodZon)
                    LALISTA(lnFilas, 2) = !cPersTelefono & ""
                    LALISTA(lnFilas, 3) = !cCtaCod
                    LALISTA(lnFilas, 4) = Format(!dVenc, "dd/mm/yyyy")
                End If
                .MoveNext
            Loop
        End With
        lrDataRep.Close
        Set lrDataRep = Nothing
    End If
    'Llena cartas
    For lnItem = 1 To lnFilas
        lsCartaModelo = psTextoCarta
        lsCartaModelo = Replace(lsCartaModelo, "<<FECHA>>", Format(psFechaSis, "dddd,d mmmm yyyy"), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<NOMBRE>>", ofcar.PstaNombre(LALISTA(lnItem, 0), False), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<FECVENCIM>>", LALISTA(lnItem, 4), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<CONTRATO>>", LALISTA(lnItem, 3), , 2, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<FECREMATE>>", Format(CDate(psFecRemate), "dd.MM.yyyy"), , 1, vbTextCompare)
        'RTFTmp = Replace(RTFTmp, "<<HORREMATE>>", vHorRem, , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<DIRECCION>>", LALISTA(lnItem, 1), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<TELEFONO>>", LALISTA(lnItem, 2), , 1, vbTextCompare)
        
        lsCadImp = lsCadImp & lsCartaModelo & oFunI.gPrnSaltoPagina
        'prgList.Value = vCont
    Next lnItem
    
    Erase LALISTA
    nRemImprimeAvisoRemate = lsCadImp

Exit Function
   
dError:
    Err.Raise Err.Number, "Obtiene Datos Imprime Listado Aviso Vencimiento  <<nImprimeAvisoVencimiento>>", Err.Description
End Function

Public Function nImprimeRemateReporteCabec(ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, _
        ByVal psFechaSis As String, ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String, _
        ByVal pnNroRepor As Integer, Optional ByVal psImpresora As Impresoras = gEPSON) As String
        
Dim lsCadImp As String
Dim loImprimeCab As NCOMColPImpre
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

    
    Set loImprimeCab = New NCOMColPImpre
        lsCadImp = loImprimeCab.nImprimeCabeceraReportes(psNomCmac, psNomAgencia, psCodUser, psFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)
    Set loImprimeCab = Nothing
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & String(pnAnchoLinea, "-") & oFunI.gPrnSaltoLinea
    Select Case pnNroRepor
        Case 1 ' PLANILLA DE CONTRATOS PARA ADJUDICAR
            ''*** PEAC 20090227
            'lsCadImp = lsCadImp & Space(1) & "Item    Contrato    ContratoAnt      Fec.Venc.     Cliente              Gramos    Tasacion   Prestamo  SaldoCap  Interes  Custodia Impuesto Prep.Rem   Deuda   P.Base  Boveda" & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Space(1) & "Item    Contrato    ContratoAnt      Fec.Venc.     Cliente              Gramos    Tasacion   Prestamo  SaldoCap  Interes   Int.Vcdo.  Int.Morat.  Notif.  Custodia  Deuda   Boveda" & oFunI.gPrnSaltoLinea
            'pnAnchoLinea = pnAnchoLinea + 35
        Case 2 ' Listados para Remate
            ''*** PEAC 20090227
            'lsCadImp = lsCadImp & Space(1) & "Item       Contrato         ContratoAnt      Piezas     Descripción                                   14Kl.    16Kl.    18Kl.    21Kl.  Prec.Base " & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Space(1) & "Item       Contrato         ContratoAnt      Piezas     Descripción                                   14Kl.    16Kl.    18Kl.    21Kl." & oFunI.gPrnSaltoLinea
            
            'pnAnchoLinea = pnAnchoLinea + 35
        Case 3 ' Planilla de Prendas Vendidas
            lsCadImp = lsCadImp & Space(1) & "Item    Contrato    Vencimiento     Tasacion   Prestamo   SaldoCap      Deuda    Venta B.   IGV  Comision   Vta Neta   Sobrante OroNeto" & oFunI.gPrnSaltoLinea
        Case 4 ' Planilla de Prendas No Vendidas
            lsCadImp = lsCadImp & Space(1) & "Item    ContratoAnt     Contrato     Vencimiento     Tasacion   Prestamo   SaldoCap      Deuda   OroNeto" & oFunI.gPrnSaltoLinea
            
        Case 9 ' Planilla de Remate
            lsCadImp = lsCadImp & Space(1) & "Item  ContratoAnt         Contrato          Fec.Venc.   Gramos    Tasacion   Prestamo  SaldoCap  Interes  Custodia Impuesto Prep.Rem   Deuda   P.Base " & oFunI.gPrnSaltoLinea
        Case 10 ' Listados para Remate
            lsCadImp = lsCadImp & Space(1) & "Item  ContratoAnt        Contrato            Piezas     Descripción                                    14Kl.    16Kl.    18Kl.    21Kl.  Prec.Base " & oFunI.gPrnSaltoLinea
        Case 11 ' Listados para Remate
            lsCadImp = lsCadImp & Space(1) & "Item  ContratoAnt        Contrato            Piezas     Descripción                                  14Kl.    16Kl.    18Kl.    21Kl.  Prec.Base " & oFunI.gPrnSaltoLinea
    End Select
    lsCadImp = lsCadImp & String(pnAnchoLinea, "-") & oFunI.gPrnSaltoLinea
nImprimeRemateReporteCabec = lsCadImp
Set oFunI = Nothing
End Function

'Private Sub CargaParametros()
'Dim loParam As COMDColocPig.DCOMColPCalculos
'Set loParam = New COMDColocPig.DCOMColPCalculos
'    fnVarTasaPreparacionRemate = loParam.dObtieneColocParametro(gConsColPTasaPreparaRemate)
'    fnVarTasaImpuesto = loParam.dObtieneColocParametro(gConsColPTasaImpuesto)
'    fnVarTasaCustodiaVencida = loParam.dObtieneColocParametro(gConsColPTasaCustodiaVencida)
'    fnVarTasaCustodia = loParam.dObtieneColocParametro(gConsColPTasaCustodia)
'    fnVarDiasCambCart = loParam.dObtieneColocParametro(gConsColPDiasCambioCartera)
'    fnVarTopRenovaciones = loParam.dObtieneColocParametro(gConsColPMaxNroRenovac)
'Set loParam = Nothing
'End Sub


'Imprime Planilla de Remate

Public Function bEsParaRemate(ByVal psCodCta As String, ByVal psFecRemate As String, ByVal pnDiasVctoParaRemate As Integer, ByVal psFechaSis As String, ByVal psNroRemate As String) As Boolean
Dim ldFecAviso As Date, RSTEMP As New ADODB.Recordset, lsSQL As String
Dim conex As New COMConecta.DCOMConecta
Set RSTEMP = New Recordset
    
    bEsParaRemate = False
conex.AbreConexion
    ldFecAviso = DateAdd("d", -pnDiasVctoParaRemate, psFecRemate)

    lsSQL = " SELECT valor=count(P.cCtaCod) " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod  " _
          & " WHERE P.nPrdEstado in (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ") " _
          & " AND P.cCtaCod ='" & psCodCta & "' AND DATEDIFF(dd, C.dVenc, '" & Format(ldFecAviso, "mm/dd/yyyy") & "') >= 0  " _
    
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open lsSQL, conex.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTEMP.EOF And RSTEMP.BOF) Then
            If RSTEMP!Valor > 0 Then
                bEsParaRemate = True
            Else
                bEsParaRemate = False
            End If
        
    End If
    

Set RSTEMP = Nothing

conex.CierraConexion
Set conex = Nothing

 
End Function


Public Function nImprimePlanillaParaRemate(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal pnDiasVctoParaRemate As Integer, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, ByVal pnFactorPrecioVenta As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroRemate As String, ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
Dim fnVarIGV As Double


Dim loCalculos As NCOMColPCalculos

Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency
    
Dim lsNombre As String * 28
Dim vSumGramos As Currency, vSumTasaci As Currency, vSumPresta As Currency
Dim vSumIntVen As Currency, vSumCosCus As Currency, vSumImpues As Currency
Dim vSumCoPrRe As Currency, vSumPreDeu As Currency, vSumPreVen As Currency
Dim vSumSalCap As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double

Dim ofcar As New COMFunciones.FCOMCadenas
Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion

'**DAOR 20070714******************
Dim lnEstProRem As Integer
Dim lrEstRem As ADODB.Recordset
'**********************************

oFunI.Inicia psImpresora
    '*****************************************************
    'Obtener Fecha de Corte
    Dim oCon As New COMConecta.DCOMConecta
    Dim rs1 As ADODB.Recordset
    Set oCon = New COMConecta.DCOMConecta
        If oCon.AbreConexion Then
           Set rs1 = New ADODB.Recordset
           lsSQL = "Select cConsDescripcion From Constante Where nConsCod='3208' and nConsValor='" & Trim(psAgencia) & "' "
           Set rs1 = oCon.CargaRecordSet(lsSQL)
           If Not (rs1.EOF And rs1.BOF) Then
             psFecRemate = Format(rs1("cConsDescripcion"), "mm/dd/yyyy")
           End If
           Set rs1 = Nothing
        End If
    Set oCon = Nothing
    '*******************************************************
       
    ldFecAviso = DateAdd("d", -pnDiasVctoParaRemate, psFecRemate)

    lsSQL = " SELECT P.cCtaCod, P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred,  " _
          & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , " _
          & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , " _
          & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , " _
          & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , " _
          & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  " _
          & "               WHERE cLineaCred = C.cLineaCred  " _
          & "               AND nColocLinCredTasaTpo  = " & gColocLineaCredTasasIntMoratNormal & " ),    " _
          & " cNomCliente = (SELECT MAX(cPersNombre) FROM Persona Per INNER JOIN ProductoPersona PPer " _
          & "               ON Per.cPersCod = PPer.cPersCod WHERE PPer.cCtaCod = P.cCtaCod ),  " _
          & " CuentaAntigua=ISNULL(cCtaCodAnt,'                  ')" _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " left JOIN RelCuentas R ON P.cCtaCod=R.cCtacod " _
          & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod  " _
          & " WHERE P.nPrdEstado in (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ") " _
          & " AND Substring(P.cCtaCod,4,2) ='" & psAgencia & "' AND DATEDIFF(dd, C.dVenc, '" & Format(ldFecAviso, "mm/dd/yyyy") & "') >= 0  " _
          & " GROUP BY P.cCtaCod,cCtaCodAnt, P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred  " _
          & " ORDER BY P.cCtaCod "

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
 

    If lrDataRep Is Nothing Then
        psmensaje = " No existen Créditos Pignoraticios para Aviso de Vencimiento "
        'MuestraImpresion = False

        Exit Function
    Else
        '**DAOR 20070714**************************
        'lsSQL = "select nRGEstado from ColocPigRecGar " & _
        '        "where cCodAge='109" & psAgencia & "' and cTpoProceso='R' and cNroProceso='" & psNroRemate & "'"
        '
        'Set loDataRep = New COMDColocPig.DCOMColPFunciones
        '    Set lrEstRem = loDataRep.dObtieneRecordSet(lsSQL)
        'Set loDataRep = Nothing
        'If (lrEstRem Is Nothing) Or lrEstRem.EOF Then
        '    lnEstProRem = 0
        'Else
        '    lnEstProRem = lrEstRem("nRGEstado")
        'End If
        'Set lrEstRem = Nothing
        '*****************************************
    
        lnLineas = 0
        lnPage = 1

        
        'CABECERA
        lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "PLANILLA DE CONTRATOS PARA REMATE NRO " & psNroRemate, "Del " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 190, "", 1)

        lnIndice = 1:        lnLineas = 7
        
        vSumGramos = 0:        vSumTasaci = 0
        vSumPresta = 0:        vSumIntVen = 0
        vSumCosCus = 0:        vSumImpues = 0
        vSumCoPrRe = 0:        vSumPreDeu = 0
        vSumPreVen = 0:        vSumSalCap = 0
        'constante para el igv
        '3007
        
        fnVarIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)

        Dim pnDeudaCosto As Double, pn15DeVenta As Double
        
        With lrDataRep
            Do While Not lrDataRep.EOF
'                If lrDataRep!cCtacod = "108013051000170429" Then Stop
                v14 = !nK14: v14 = Round(v14, 2)
                v16 = !nK16: v16 = Round(v16, 2)
                v18 = !nK18: v18 = Round(v18, 2)
                v21 = !nK21: v21 = Round(v21, 2)
                
                pn15DeVenta = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * 0.15
              
                'lnPreBase = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * (fnVarIGV + 1)
                lnPreBase = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21))
            

                lnGramos = v14 + v16 + v18 + v21
                
                lnPreBase = Round(lnPreBase, 2)
                
                lsNombre = ofcar.PstaNombre(!cNomCliente)
        
                vdiasAtraso = DateDiff("d", !dVenc, psFecRemate)
                
                Set loCalculos = New NCOMColPCalculos
                    If vdiasAtraso <= 0 Then
                        vdiasAtraso = 0
                        vInteresVencido = 0
                        vCostoCustodiaMoratorio = 0
                        vImpuesto = 0
                    Else
                        vInteresVencido = loCalculos.nCalculaInteresMoratorio(!nSaldo, !nTasaIntVenc, vdiasAtraso)
                        vInteresVencido = Round(vInteresVencido, 2)
                        vCostoCustodiaMoratorio = loCalculos.nCalculaCostoCustodiaMoratorio(!nTasacion, pnTasaCustodiaVencida, vdiasAtraso)
                        vCostoCustodiaMoratorio = Round(vCostoCustodiaMoratorio, 2)
                        vImpuesto = (vInteresVencido + vCostoCustodiaMoratorio) * pnTasaImpuesto
                        vImpuesto = Round(vImpuesto, 2)
                    End If
                    
                    vCostoPreparacionRemate = 0
                    '**Comentado por DAOR 20070713
                    If !nPrdEstado = gColPEstPRema Then    ' Si esta en via de Remate
                        vCostoPreparacionRemate = pnTasaPreparacionRemate * !nTasacion
                        vCostoPreparacionRemate = Round(vCostoPreparacionRemate, 2)
                    End If
                    lnDeuda = !nSaldo + vInteresVencido + vCostoCustodiaMoratorio + vImpuesto + vCostoPreparacionRemate
                    
                    '**DAOR 20070714***********************
                    'If !nPrdEstado = gColPEstPRema And lnEstProRem = 1 Then  ' Si esta en via de Remate y además el proceso se encuentra iniciado
                    '    vCostoPreparacionRemate = pnTasaPreparacionRemate * !nTasacion
                    '    vCostoPreparacionRemate = Round(vCostoPreparacionRemate, 2)
                    'End If
                    'lnDeuda = !nSaldo + vInteresVencido + vCostoCustodiaMoratorio + vImpuesto + vCostoPreparacionRemate
                    '***************************************
                    
                    'pnDeudaCosto = Round((lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta) * (fnVarIGV + 1), 2)
                    pnDeudaCosto = Round((lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta), 2)
                    
                    lnPreVenta = loCalculos.nCalculaPrecioVentaRemate(pnDeudaCosto, lnPreBase, pnFactorPrecioVenta)
                    lnPreVenta = CalculaPrecioVentaRemate(Round(lnPreVenta, 0))
                    
                    
                Set loCalculos = Nothing

                'vInteresVencido = CalculaInteresMoratorio(!nSaldoCap, !nTasaIntVenc, vDiasAtraso)
                'vInteresVencido = Round(vInteresVencido, 2)
                'vCostoCustodia = CalculaCostoCustodiaMoratorio(!nvaltasac, pTasaCustodiaVencida, vDiasAtraso)
                'vCostoCustodia = Round(vCostoCustodia, 2)
                'vCostoPreparacionRemate = pTasaPreparacionRemate * !nvaltasac
                'vCostoPreparacionRemate = Round(vCostoPreparacionRemate, 2)
                'vImpuesto = (vInteresVencido + vCostoCustodia + vCostoPreparacionRemate) * pTasaImpuesto
                'vImpuesto = Round(vImpuesto, 2)
                'lnPreDeuda = CalculaDeudaPrendario(!nSaldoCap, !dFecVenc, !nvaltasac, !nTasaIntVenc, _
                          !nCostCusto, pTasaImpuesto, "6", pTasaPreparacionRemate, txtFecRemate)
                'lnPreDeuda = Round(lnPreDeuda, 2)
                'lnPreVenta = IIf((lnPreDeuda * pFactorPrecioBaseRemate) > vPreBas, CalculaPrecioVentaRemate(lnPreDeuda * pFactorPrecioBaseRemate), CalculaPrecioVentaRemate(vPreBas))
                'lnPreVenta = Round(lnPreVenta, 2)
                
                

                lnLineas = lnLineas + 1
                    
                vSumGramos = vSumGramos + lnGramos
                vSumTasaci = vSumTasaci + !nTasacion
                vSumPresta = vSumPresta + !nMontoCol
                vSumSalCap = vSumSalCap + !nSaldo
                vSumIntVen = vSumIntVen + vInteresVencido
                vSumCosCus = vSumCosCus + vCostoCustodiaMoratorio
                vSumImpues = vSumImpues + vImpuesto
                vSumCoPrRe = vSumCoPrRe + vCostoPreparacionRemate
                vSumPreDeu = vSumPreDeu + lnDeuda
                vSumPreVen = vSumPreVen + lnPreVenta
                                        
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 4, 0) & Space(1) & !cCtaCod & Space(1) & !CuentaAntigua & Space(1) & Format(!dVenc, "dd/mm/yyyy") _
                    & ofun.ImpreFormat(ofun.ImpreCarEsp(lsNombre), 20, 1) _
                    & ofun.ImpreFormat(lnGramos, 6) & ofun.ImpreFormat(!nTasacion, 8) & ofun.ImpreFormat(!nMontoCol, 8) _
                    & ofun.ImpreFormat(!nSaldo, 8) & ofun.ImpreFormat(vInteresVencido, 6) & ofun.ImpreFormat(vCostoCustodiaMoratorio, 5) _
                    & ofun.ImpreFormat(vImpuesto, 5) & ofun.ImpreFormat(vCostoPreparacionRemate, 6) & ofun.ImpreFormat(lnDeuda, 8) _
                    & ofun.ImpreFormat(lnPreVenta, 8) & Space(2) & !cAgeBovedA _
                    & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                        
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "PLANILLA DE CONTRATOS PARA REMATE NRO " & psNroRemate, "Del " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 155, "", 1)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(32) & "RESUMEN " & ofun.ImpreFormat(vSumGramos, 41) & ofun.ImpreFormat(vSumTasaci, 8) _
                & ofun.ImpreFormat(vSumPresta, 8) & ofun.ImpreFormat(vSumSalCap, 8) & ofun.ImpreFormat(vSumIntVen, 6) _
                & ofun.ImpreFormat(vSumCosCus, 5) & ofun.ImpreFormat(vSumImpues, 5) & ofun.ImpreFormat(vSumCoPrRe, 6) _
                & ofun.ImpreFormat(vSumPreDeu, 8) & ofun.ImpreFormat(vSumPreVen, 8) _
                & oFunI.gPrnSaltoLinea
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimePlanillaParaRemate = lsCadBuffer
    Set oFunI = Nothing
End Function

Public Function nImprimePlanillaParaRemateConSiaf(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal pnDiasVctoParaRemate As Integer, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, ByVal pnFactorPrecioVenta As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroRemate As String, ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
Dim fnVarIGV As Double

Dim loCalculos As NCOMColPCalculos

Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency
    
Dim lsNombre As String * 28
Dim vSumGramos As Currency, vSumTasaci As Currency, vSumPresta As Currency
Dim vSumIntVen As Currency, vSumCosCus As Currency, vSumImpues As Currency
Dim vSumCoPrRe As Currency, vSumPreDeu As Currency, vSumPreVen As Currency
Dim vSumSalCap As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double
Dim ofcar As New COMFunciones.FCOMCadenas
Dim ofun As New COMFunciones.FCOMImpresion

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

    
    ldFecAviso = DateAdd("d", -pnDiasVctoParaRemate, psFecRemate)

    lsSQL = " SELECT CODESIAF=isnull(RE.cctacodant,'                  ') ,P.cCtaCod, P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred,  " _
          & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , " _
          & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , " _
          & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , " _
          & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , " _
          & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  " _
          & "               WHERE cLineaCred = C.cLineaCred  " _
          & "               AND nColocLinCredTasaTpo  = " & gColocLineaCredTasasIntMoratNormal & " ),    " _
          & " cNomCliente = (SELECT MAX(cPersNombre) FROM Persona Per INNER JOIN ProductoPersona PPer " _
          & "               ON Per.cPersCod = PPer.cPersCod WHERE PPer.cCtaCod = P.cCtaCod )  " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " LEFT JOIN RELCUENTAS RE ON RE.CCTACOD=P.CCTACOD  " _
          & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod  " _
          & " WHERE P.nPrdEstado in (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ") " _
          & " AND Substring(P.cCtaCod,4,2) ='" & psAgencia & "' AND DATEDIFF(dd, C.dVenc, '" & Format(ldFecAviso, "mm/dd/yyyy") & "') >= 0  " _
          & " GROUP BY P.cCtaCod, P.nPrdestado,isnull(RE.cctacodant,'                  '),C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred  " _
          & " ORDER BY P.cCtaCod "

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
 

    If lrDataRep Is Nothing Then
        psmensaje = " No existen Créditos Pignoraticios para Aviso de Vencimiento "
        'MuestraImpresion = False

        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        
        'CABECERA
        lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "PLANILLA DE CONTRATOS PARA REMATE NRO " & psNroRemate, "Del " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 155, "", 9)

        lnIndice = 1:        lnLineas = 7
        
        vSumGramos = 0:        vSumTasaci = 0
        vSumPresta = 0:        vSumIntVen = 0
        vSumCosCus = 0:        vSumImpues = 0
        vSumCoPrRe = 0:        vSumPreDeu = 0
        vSumPreVen = 0:        vSumSalCap = 0
        'constante para el igv
        '3007
        
        fnVarIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)

        Dim pnDeudaCosto As Double, pn15DeVenta As Double
        
        With lrDataRep
            Do While Not lrDataRep.EOF
'                If lrDataRep!cCtacod = "108013051000170429" Then Stop
                v14 = !nK14: v14 = Round(v14, 2)
                v16 = !nK16: v16 = Round(v16, 2)
                v18 = !nK18: v18 = Round(v18, 2)
                v21 = !nK21: v21 = Round(v21, 2)
                
                pn15DeVenta = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * 0.15
              
                lnPreBase = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) '* (fnVarIGV + 1)
                
                lnGramos = v14 + v16 + v18 + v21
                
                lnPreBase = Round(lnPreBase, 2)
                
                lsNombre = ofcar.PstaNombre(!cNomCliente)
        
                vdiasAtraso = DateDiff("d", !dVenc, psFecRemate)
                
                Set loCalculos = New NCOMColPCalculos
                    If vdiasAtraso <= 0 Then
                        vdiasAtraso = 0
                        vInteresVencido = 0
                        vCostoCustodiaMoratorio = 0
                        vImpuesto = 0
                    Else
                        vInteresVencido = loCalculos.nCalculaInteresMoratorio(!nSaldo, !nTasaIntVenc, vdiasAtraso)
                        vInteresVencido = Round(vInteresVencido, 2)
                        vCostoCustodiaMoratorio = loCalculos.nCalculaCostoCustodiaMoratorio(!nTasacion, pnTasaCustodiaVencida, vdiasAtraso)
                        vCostoCustodiaMoratorio = Round(vCostoCustodiaMoratorio, 2)
                        vImpuesto = (vInteresVencido + vCostoCustodiaMoratorio) * pnTasaImpuesto
                        vImpuesto = Round(vImpuesto, 2)
                    End If
                    
                    vCostoPreparacionRemate = 0
                    If !nPrdEstado = gColPEstPRema Then    ' Si esta en via de Remate
                        vCostoPreparacionRemate = pnTasaPreparacionRemate * !nTasacion
                        vCostoPreparacionRemate = Round(vCostoPreparacionRemate, 2)
                    End If
                    lnDeuda = !nSaldo + vInteresVencido + vCostoCustodiaMoratorio + vImpuesto + vCostoPreparacionRemate
                    
                    pnDeudaCosto = Round((lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta), 2)
                    
                    'pnDeudaCosto = Round((lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta) * (fnVarIGV + 1), 2)
                    lnPreVenta = loCalculos.nCalculaPrecioVentaRemate(pnDeudaCosto, lnPreBase, pnFactorPrecioVenta)
                    lnPreVenta = Round(lnPreVenta, 0)
                    
                    
                    
                    
                Set loCalculos = Nothing

                'vInteresVencido = CalculaInteresMoratorio(!nSaldoCap, !nTasaIntVenc, vDiasAtraso)
                'vInteresVencido = Round(vInteresVencido, 2)
                'vCostoCustodia = CalculaCostoCustodiaMoratorio(!nvaltasac, pTasaCustodiaVencida, vDiasAtraso)
                'vCostoCustodia = Round(vCostoCustodia, 2)
                'vCostoPreparacionRemate = pTasaPreparacionRemate * !nvaltasac
                'vCostoPreparacionRemate = Round(vCostoPreparacionRemate, 2)
                'vImpuesto = (vInteresVencido + vCostoCustodia + vCostoPreparacionRemate) * pTasaImpuesto
                'vImpuesto = Round(vImpuesto, 2)
                'lnPreDeuda = CalculaDeudaPrendario(!nSaldoCap, !dFecVenc, !nvaltasac, !nTasaIntVenc, _
                          !nCostCusto, pTasaImpuesto, "6", pTasaPreparacionRemate, txtFecRemate)
                'lnPreDeuda = Round(lnPreDeuda, 2)
                'lnPreVenta = IIf((lnPreDeuda * pFactorPrecioBaseRemate) > vPreBas, CalculaPrecioVentaRemate(lnPreDeuda * pFactorPrecioBaseRemate), CalculaPrecioVentaRemate(vPreBas))
                'lnPreVenta = Round(lnPreVenta, 2)
                
                

                lnLineas = lnLineas + 1
                    
                vSumGramos = vSumGramos + lnGramos
                vSumTasaci = vSumTasaci + !nTasacion
                vSumPresta = vSumPresta + !nMontoCol
                vSumSalCap = vSumSalCap + !nSaldo
                vSumIntVen = vSumIntVen + vInteresVencido
                vSumCosCus = vSumCosCus + vCostoCustodiaMoratorio
                vSumImpues = vSumImpues + vImpuesto
                vSumCoPrRe = vSumCoPrRe + vCostoPreparacionRemate
                vSumPreDeu = vSumPreDeu + lnDeuda
                vSumPreVen = vSumPreVen + lnPreVenta
                                        
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 4, 0) & Space(1) & !codesiaf & Space(1) & !cCtaCod & Space(1) & Format(!dVenc, "dd/mm/yyyy") _
                           & ofun.ImpreFormat(lnGramos, 6) & ofun.ImpreFormat(!nTasacion, 8) & ofun.ImpreFormat(!nMontoCol, 8) _
                           & ofun.ImpreFormat(!nSaldo, 8) & ofun.ImpreFormat(vInteresVencido, 6) & ofun.ImpreFormat(vCostoCustodiaMoratorio, 5) _
                           & ofun.ImpreFormat(vImpuesto, 5) & ofun.ImpreFormat(vCostoPreparacionRemate, 6) & ofun.ImpreFormat(lnDeuda, 8) _
                           & ofun.ImpreFormat(lnPreVenta, 8) & oFunI.gPrnSaltoLinea   '& Space(2) & !cAgeBoveda _

                    '& ofun.ImpreFormat(oFun.ImpreCarEsp(lsNombre), 20, 1) _

                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                        
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "PLANILLA DE CONTRATOS PARA REMATE NRO " & psNroRemate, "Del " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 155, "", 9)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(10) & "RESUMEN " & ofun.ImpreFormat(vSumGramos, 41) & ofun.ImpreFormat(vSumTasaci, 8) _
                  & ofun.ImpreFormat(vSumPresta, 8) & ofun.ImpreFormat(vSumSalCap, 8) & ofun.ImpreFormat(vSumIntVen, 6) _
                  & ofun.ImpreFormat(vSumCosCus, 5) & ofun.ImpreFormat(vSumImpues, 5) & ofun.ImpreFormat(vSumCoPrRe, 6) _
                  & ofun.ImpreFormat(vSumPreDeu, 8) & ofun.ImpreFormat(vSumPreVen, 8) & oFunI.gPrnSaltoLinea
                  '--& ofun.ImpreFormat(vSumPreVen, 8) & oFunI.gPrnSaltoLinea
                  '& oFunI.gPrnSaltoLinea
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimePlanillaParaRemateConSiaf = lsCadBuffer
    Set oFunI = Nothing
End Function


'Imprime Listado para Remate
Public Function nImprimeListadoParaRemate(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal pnDiasVctoParaRemate As Integer, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, ByVal pnFactorPrecioVenta As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroRemate As String, _
        ByVal pnJoyasDetalle As Boolean, ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
Dim fnVarIGV As Double



Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim loCalculos As NCOMColPCalculos

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double
    
Dim lrCredPigJoyasDet As New ADODB.Recordset
Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
Dim item As Integer, itemTMP As Integer
Dim lmDetalle(20) As String

Dim ofun As New COMFunciones.FCOMImpresion

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

    '*****************************************************
    'Obtener Fecha de Corte
    Dim oCon As New COMConecta.DCOMConecta
    Dim rs1 As ADODB.Recordset
    Set oCon = New COMConecta.DCOMConecta
        If oCon.AbreConexion Then
           Set rs1 = New ADODB.Recordset
           lsSQL = "Select cConsDescripcion From Constante Where nConsCod='3208' and nConsValor='" & Trim(psAgencia) & "' "
           Set rs1 = oCon.CargaRecordSet(lsSQL)
           If Not (rs1.EOF And rs1.BOF) Then
             psFecRemate = Format(rs1("cConsDescripcion"), "mm/dd/yyyy")
           End If
           Set rs1 = Nothing
        End If
    Set oCon = Nothing
    '*******************************************************
    

    ldFecAviso = DateAdd("d", -pnDiasVctoParaRemate, psFecRemate)

    lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote,   " _
          & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , " _
          & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , " _
          & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , " _
          & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , " _
          & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  " _
          & "               WHERE cLineaCred = C.cLineaCred  " _
          & "               AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ),    " _
          & " CuentaAntigua=ISNULL(cCtaCodAnt,'                  ')" _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " left JOIN RelCuentas R ON P.cCtaCod=R.cCtacod " _
          & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod  " _
          & " WHERE P.nPrdEstado in (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ") " _
          & " AND Substring(P.cCtaCod,4,2) = '" & psAgencia & "' AND DATEDIFF(dd, C.dVenc, '" & Format(ldFecAviso, "mm/dd/yyyy") & "') >= 0  " _
          & " GROUP BY P.cCtaCod,cCtaCodAnt, P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote  " _
          & " ORDER BY P.cCtaCod "

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing





    If lrDataRep Is Nothing Then
        psmensaje = " No existen Créditos Pignoraticios para Aviso de Vencimiento "
        'MuestraImpresion = False

        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "LISTADO DE CONTRATOS PARA REMATE NRO " & psNroRemate, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 155, "", 2)

        lnIndice = 1:        lnLineas = 7
        

        'Set loCalculos = New COMDColocPig.DCOMColPCalculos
        
      
        
                
        fnVarIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)

        Dim pnDeudaCosto As Double, pn15DeVenta As Double

        
        With lrDataRep
            Do While Not lrDataRep.EOF
                
                v14 = !nK14: v14 = Round(v14, 2)
                v16 = !nK16: v16 = Round(v16, 2)
                v18 = !nK18: v18 = Round(v18, 2)
                v21 = !nK21: v21 = Round(v21, 2)
                vCuenta = 0
                
                pn15DeVenta = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * 0.15
              
                lnPreBase = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) ' * (fnVarIGV + 1)
                
                'lnPreBase = (v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)
                
                lnGramos = v14 + v16 + v18 + v21
                
                lnPreBase = Round(lnPreBase, 2)
                
                vdiasAtraso = DateDiff("d", !dVenc, psFecRemate)
                
                Set loCalculos = New NCOMColPCalculos
                    If vdiasAtraso <= 0 Then
                        vdiasAtraso = 0
                        vInteresVencido = 0
                        vCostoCustodiaMoratorio = 0
                        vImpuesto = 0
                    Else
                        vInteresVencido = loCalculos.nCalculaInteresMoratorio(!nSaldo, !nTasaIntVenc, vdiasAtraso)
                        vInteresVencido = Round(vInteresVencido, 2)
                        vCostoCustodiaMoratorio = loCalculos.nCalculaCostoCustodiaMoratorio(!nTasacion, pnTasaCustodiaVencida, vdiasAtraso)
                        vCostoCustodiaMoratorio = Round(vCostoCustodiaMoratorio, 2)
                        vImpuesto = (vInteresVencido + vCostoCustodiaMoratorio) * pnTasaImpuesto
                        vImpuesto = Round(vImpuesto, 2)
                    End If
                    vCostoPreparacionRemate = 0
                    If !nPrdEstado = gColPEstPRema Then    ' Si esta en via de Remate
                        vCostoPreparacionRemate = pnTasaPreparacionRemate * !nTasacion
                        vCostoPreparacionRemate = Round(vCostoPreparacionRemate, 2)
                    End If
                    lnDeuda = !nSaldo + vInteresVencido + vCostoCustodiaMoratorio + vImpuesto + vCostoPreparacionRemate
                    
                    
                    pnDeudaCosto = Round((lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta), 2)
                    ' pnDeudaCosto = Round((lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta) * (fnVarIGV + 1), 2)
                    'Calcula el Precio Venta para Remate
                    lnPreVenta = loCalculos.nCalculaPrecioVentaRemate(pnDeudaCosto, lnPreBase, pnFactorPrecioVenta)
                    lnPreVenta = CalculaPrecioVentaRemate(Round(lnPreVenta, 2))
                Set loCalculos = Nothing
                
                'Joyas Detalladas
                lnLineas = lnLineas + 1
                If pnJoyasDetalle = True Then

                    Set loMuestraContrato = New COMDColocPig.DCOMColPContrato
                        Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(!cCtaCod)
                    Set loMuestraContrato = Nothing
                    
                    item = 0
                    Do While Not lrCredPigJoyasDet.EOF
                        lmDetalle(item) = ofun.ImpreFormat(lrCredPigJoyasDet!nPiezas, 4, 0) & " " & ofun.ImpreFormat(lrCredPigJoyasDet!cDescrip, 27) & " " _
                                    & lrCredPigJoyasDet!cKilataje & "K " & ofun.ImpreFormat(lrCredPigJoyasDet!npesoneto, 4, 2) & "Gr"
                        item = item + 1
                        lrCredPigJoyasDet.MoveNext
                    Loop
                    
                    lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & Mid(!cCtaCod, 1, 5) & "-" & Mid(!cCtaCod, 6) & Space(1) & !CuentaAntigua & Space(1) & ofun.ImpreFormat(!nPiezas, 3, 0) & Space(1) _
                                        & lmDetalle(0) & ofun.ImpreFormat(v14, 6) & ofun.ImpreFormat(v16, 6) & ofun.ImpreFormat(v18, 6) & ofun.ImpreFormat(v21, 6) _
                                        & ofun.ImpreFormat(lnPreVenta, 8) & oFunI.gPrnSaltoLinea
                    
                    For itemTMP = 1 To item - 1
                            lsCadImp = lsCadImp & Space(49) & lmDetalle(itemTMP) & oFunI.gPrnSaltoLinea
                            lnLineas = lnLineas + 1
                    Next
                    

                Else ' **
                    lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !cCtaCod & Space(1) & !CuentaAntigua & Space(1) & ofun.ImpreFormat(!nPiezas, 3, 0) & Space(1) _
                                        & ofun.ImpreFormat(ofun.QuiebreTexto(!cLote, 1), 42, 1) & ofun.ImpreFormat(v14, 6) & ofun.ImpreFormat(v16, 6) & ofun.ImpreFormat(v18, 6) & ofun.ImpreFormat(v21, 6) _
                                        & ofun.ImpreFormat(lnPreVenta, 8) & oFunI.gPrnSaltoLinea
                    
                    If lnIndice Mod 300 = 0 Then
                        lsCadBuffer = lsCadBuffer & lsCadImp
                        lsCadImp = ""
                    End If
                                            
    
                    Do While vCuenta < 15
                        vCuenta = vCuenta + 1
                        If Len(ofun.QuiebreTexto(!cLote, vCuenta + 1)) > 0 Then
                            lsCadImp = lsCadImp & Space(30) & ofun.QuiebreTexto(!cLote, vCuenta + 1) & oFunI.gPrnSaltoLinea
                            lnLineas = lnLineas + 1
                        End If
                    Loop
                
                End If
                lnLineas = lnLineas + 1
                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
                
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "LISTADO DE CONTRATOS PARA REMATE NRO " & psNroRemate, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 125, "", 2)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(155, "*") & oFunI.gPrnSaltoLinea
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimeListadoParaRemate = lsCadBuffer
    Set oFunI = Nothing
End Function

'Imprime Listado para Remate
Public Function nImprimeListadoParaRemateConSIAF(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal pnDiasVctoParaRemate As Integer, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, ByVal pnFactorPrecioVenta As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroRemate As String, _
        ByVal pnJoyasDetalle As Boolean, ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
Dim fnVarIGV As Double



Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim loCalculos As NCOMColPCalculos

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double
    
Dim lrCredPigJoyasDet As New ADODB.Recordset
Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
Dim item As Integer, itemTMP As Integer
Dim lmDetalle(20) As String
Dim ofun As New COMFunciones.FCOMImpresion

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora



    ldFecAviso = DateAdd("d", -pnDiasVctoParaRemate, psFecRemate)

    lsSQL = " SELECT CODESIAF=ISNULL(RE.CCTACODANT,'                  '),P.cCtaCod, P.nPrdEstado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote,   " _
          & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , " _
          & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , " _
          & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , " _
          & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , " _
          & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  " _
          & "               WHERE cLineaCred = C.cLineaCred  " _
          & "               AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " )    " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " LEFT JOIN RELCUENTAS RE ON RE.CCTACOD=P.CCTACOD " _
          & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod  " _
          & " WHERE P.nPrdEstado in (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ") " _
          & " AND Substring(P.cCtaCod,4,2) = '" & psAgencia & "' AND DATEDIFF(dd, C.dVenc, '" & Format(ldFecAviso, "mm/dd/yyyy") & "') >= 0  " _
          & " GROUP BY P.cCtaCod, P.nPrdestado,ISNULL(RE.CCTACODANT,'                  ') ,C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote  " _
          & " ORDER BY P.cCtaCod "

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing





    If lrDataRep Is Nothing Then
        psmensaje = " No existen Créditos Pignoraticios para Aviso de Vencimiento "
        'MuestraImpresion = False

        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "LISTADO DE CONTRATOS PARA REMATE NRO " & psNroRemate, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 150, "", 10)

        lnIndice = 1:        lnLineas = 7
        

        'Set loCalculos = New COMDColocPig.DCOMColPCalculos
        
      
        
                
        fnVarIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)

        Dim pnDeudaCosto As Double, pn15DeVenta As Double

        
        With lrDataRep
            Do While Not lrDataRep.EOF
                
                v14 = !nK14: v14 = Round(v14, 2)
                v16 = !nK16: v16 = Round(v16, 2)
                v18 = !nK18: v18 = Round(v18, 2)
                v21 = !nK21: v21 = Round(v21, 2)
                vCuenta = 0
                
                pn15DeVenta = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * 0.15
              
                'lnPreBase = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) '* (fnVarIGV + 1)
                
                lnPreBase = (v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)
                
                
                lnGramos = v14 + v16 + v18 + v21
                
                lnPreBase = Round(lnPreBase, 2)
                
                vdiasAtraso = DateDiff("d", !dVenc, psFecRemate)
                
                Set loCalculos = New NCOMColPCalculos
                    If vdiasAtraso <= 0 Then
                        vdiasAtraso = 0
                        vInteresVencido = 0
                        vCostoCustodiaMoratorio = 0
                        vImpuesto = 0
                    Else
                        vInteresVencido = loCalculos.nCalculaInteresMoratorio(!nSaldo, !nTasaIntVenc, vdiasAtraso)
                        vInteresVencido = Round(vInteresVencido, 2)
                        vCostoCustodiaMoratorio = loCalculos.nCalculaCostoCustodiaMoratorio(!nTasacion, pnTasaCustodiaVencida, vdiasAtraso)
                        vCostoCustodiaMoratorio = Round(vCostoCustodiaMoratorio, 2)
                        vImpuesto = (vInteresVencido + vCostoCustodiaMoratorio) * pnTasaImpuesto
                        vImpuesto = Round(vImpuesto, 2)
                    End If
                    vCostoPreparacionRemate = 0
                    If !nPrdEstado = gColPEstPRema Then    ' Si esta en via de Remate
                        vCostoPreparacionRemate = pnTasaPreparacionRemate * !nTasacion
                        vCostoPreparacionRemate = Round(vCostoPreparacionRemate, 2)
                    End If
                    lnDeuda = !nSaldo + vInteresVencido + vCostoCustodiaMoratorio + vImpuesto + vCostoPreparacionRemate
                    
                    
                   ' pnDeudaCosto = lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta
                     pnDeudaCosto = Round((lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta), 2)
                     
                     'pnDeudaCosto = Round((lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta) * (fnVarIGV + 1), 2)
                     
                    'Calcula el Precio Venta para Remate
                    lnPreVenta = loCalculos.nCalculaPrecioVentaRemate(pnDeudaCosto, lnPreBase, pnFactorPrecioVenta)
                    lnPreVenta = Round(lnPreVenta, 2)
                Set loCalculos = Nothing
                
                'Joyas Detalladas
                lnLineas = lnLineas + 1
                If pnJoyasDetalle = True Then

                    Set loMuestraContrato = New COMDColocPig.DCOMColPContrato
                        Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(!cCtaCod)
                    Set loMuestraContrato = Nothing
                    
                    item = 0
                    Do While Not lrCredPigJoyasDet.EOF
                        lmDetalle(item) = ofun.ImpreFormat(lrCredPigJoyasDet!nPiezas, 3, 0) & "" & ofun.ImpreFormat(lrCredPigJoyasDet!cDescrip, 27) & " " _
                                    & lrCredPigJoyasDet!cKilataje & "K " & ofun.ImpreFormat(lrCredPigJoyasDet!nPesoBruto, 3, 2) & "g" & ofun.ImpreFormat(lrCredPigJoyasDet!npesoneto, 3, 2) & "g"
                        item = item + 1
                        lrCredPigJoyasDet.MoveNext
                    Loop
                  
                    lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !codesiaf & Space(1) & Mid(!cCtaCod, 1, 5) & "-" & Mid(!cCtaCod, 6) & Space(1) & ofun.ImpreFormat(!nPiezas, 3, 0) & Space(1) _
                                        & lmDetalle(0) & ofun.ImpreFormat(v14, 6) & ofun.ImpreFormat(v16, 6) & ofun.ImpreFormat(v18, 6) & ofun.ImpreFormat(v21, 6) _
                                        & ofun.ImpreFormat(lnPreVenta, 8) & oFunI.gPrnSaltoLinea
                    
                    For itemTMP = 1 To item - 1
                            lsCadImp = lsCadImp & Space(49) & lmDetalle(itemTMP) & oFunI.gPrnSaltoLinea
                            lnLineas = lnLineas + 1
                    Next
                    

                Else ' **
                    lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !codesiaf & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!nPiezas, 3, 0) & Space(1) _
                                        & ofun.ImpreFormat(ofun.QuiebreTexto(!cLote, 1), 42, 1) & ofun.ImpreFormat(v14, 6) & ofun.ImpreFormat(v16, 6) & ofun.ImpreFormat(v18, 6) & ofun.ImpreFormat(v21, 6) _
                                        & ofun.ImpreFormat(lnPreVenta, 8) & oFunI.gPrnSaltoLinea
                    
                    If lnIndice Mod 300 = 0 Then
                        lsCadBuffer = lsCadBuffer & lsCadImp
                        lsCadImp = ""
                    End If
                                            
    
                    Do While vCuenta < 15
                        vCuenta = vCuenta + 1
                        If Len(ofun.QuiebreTexto(!cLote, vCuenta + 1)) > 0 Then
                            lsCadImp = lsCadImp & Space(49) & ofun.QuiebreTexto(!cLote, vCuenta + 1) & oFunI.gPrnSaltoLinea
                            lnLineas = lnLineas + 1
                        End If
                    Loop
                
                End If
                lnLineas = lnLineas + 1
                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
                
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "LISTADO DE CONTRATOS PARA REMATE NRO " & psNroRemate, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 150, "", 10)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(150, "*") & oFunI.gPrnSaltoLinea
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimeListadoParaRemateConSIAF = lsCadBuffer
    Set oFunI = Nothing
End Function

Public Function nImpListadoRemateAnteriores(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal pnDiasVctoParaRemate As Integer, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, ByVal pnFactorPrecioVenta As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroRemate As String, _
        ByVal pnJoyasDetalle As Boolean, ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
Dim fnVarIGV As Double



Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim loCalculos As NCOMColPCalculos

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double
    
Dim lrCredPigJoyasDet As New ADODB.Recordset
Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
Dim item As Integer, itemTMP As Integer
Dim lmDetalle(20) As String
Dim ldFecAnt As Date

Dim ofun As New COMFunciones.FCOMImpresion

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


    ldFecAviso = DateAdd("d", -pnDiasVctoParaRemate, psFecRemate)
    
    ldFecAnt = DateAdd("d", -1, psFecRemate)

'    lsSQL = " SELECT CODESIAF=ISNULL(RE.CCTACODANT,'                  '),P.cCtaCod, P.nPrdEstado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
'          & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote,   " _
'          & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , " _
'          & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , " _
'          & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , " _
'          & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , " _
'          & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  " _
'          & "               WHERE cLineaCred = C.cLineaCred  " _
'          & "               AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " )    " _
'          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
'          & " LEFT JOIN RELCUENTAS RE ON RE.CCTACOD=P.CCTACOD " _
'          & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
'          & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod  " _
'          & " WHERE P.nPrdEstado in (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ") " _
'          & " AND Substring(P.cCtaCod,4,2) = '" & PSAGENCIA & "' AND DATEDIFF(dd, C.dVenc, '" & Format(ldFecAviso, "mm/dd/yyyy") & "') >= 0  " _
'          & " GROUP BY P.cCtaCod, P.nPrdestado,ISNULL(RE.CCTACODANT,'                  ') ,C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
'          & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote  " _
'          & " ORDER BY P.cCtaCod "


 lsSQL = "           SELECT  CODESIAF=ISNULL(RE.CCTACODANT,'                  '),P.cCtaCod, C.dVenc, CP.nTasacion,"
 lsSQL = lsSQL & " C.nMontoCol,  P.nSaldocap, A.NREMSUBBASEVTA, C.cLineaCred, CP.nPiezas, CP.cLote,"
 lsSQL = lsSQL & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 ,"
 lsSQL = lsSQL & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 ,"
 lsSQL = lsSQL & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 ,"
 lsSQL = lsSQL & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 ,"
 lsSQL = lsSQL & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa   WHERE cLineaCred = C.cLineaCred"
 lsSQL = lsSQL & "                  AND nColocLinCredTasaTpo = 3 ), "
 lsSQL = lsSQL & " cNomCliente = (SELECT MAX(cPersNombre) FROM Persona Per INNER JOIN ProductoPersona PPer   ON Per.cPersCod = PPer.cPersCod WHERE PPer.cCtaCod = P.cCtaCod ) "
 lsSQL = lsSQL & " FROM (select * from colocacsaldo where convert(char(8),dfecha,112)='" & Format(ldFecAnt, "yyyymmdd") & "' and cctacod like '108__305%' ) P"
 lsSQL = lsSQL & " LEFT JOIN RELCUENTAS RE ON RE.CCTACOD=P.CCTACOD "
 lsSQL = lsSQL & " Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod"
 lsSQL = lsSQL & " inner join ( select r.*,d.cctacod,d.NREMSUBBASEVTA from  colocpigrgdet d join colocpigrecgar r on r.cnroproceso=d.cnroproceso and d.ctpoproceso=r.ctpoproceso where r.ctpoproceso='R' and   convert(char(8),dproceso,112)='" & Format(psFecRemate, "yyyymmdd") & "'"
 lsSQL = lsSQL & "        ) a on a.cctacod=p.cctacod"
 lsSQL = lsSQL & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod"
 lsSQL = lsSQL & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod"
 lsSQL = lsSQL & " where substring(p.cctacod,4,2)='" & psAgencia & "'"
 lsSQL = lsSQL & " GROUP BY RE.CCTACODANT,P.cCtaCod, C.dVenc, CP.nTasacion, A.NREMSUBBASEVTA,C.nMontoCol, P.nSaldocap,  CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote"
 lsSQL = lsSQL & " ORDER BY P.cCtaCod"



    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Then
        psmensaje = " No existen Créditos Pignoraticios para Aviso de Vencimiento "
        'MuestraImpresion = False

        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "LISTADO DE CONTRATOS PARA REMATE NRO " & psNroRemate, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 145, "", 11)

        lnIndice = 1:        lnLineas = 7
        

        'Set loCalculos = New COMDColocPig.DCOMColPCalculos
             
        fnVarIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)

        Dim pnDeudaCosto As Double, pn15DeVenta As Double

        
        With lrDataRep
            Do While Not lrDataRep.EOF
                
                v14 = !nK14: v14 = Round(v14, 2)
                v16 = !nK16: v16 = Round(v16, 2)
                v18 = !nK18: v18 = Round(v18, 2)
                v21 = !nK21: v21 = Round(v21, 2)
                vCuenta = 0
                
                pn15DeVenta = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * 0.15
              
                'lnPreBase = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * (fnVarIGV + 1)
                
                lnPreBase = (v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)
                
                
                lnGramos = v14 + v16 + v18 + v21
                
                lnPreBase = Round(lnPreBase, 2)
                
                vdiasAtraso = DateDiff("d", !dVenc, psFecRemate)
                

                lnLineas = lnLineas + 1
                If pnJoyasDetalle = True Then

                    Set loMuestraContrato = New COMDColocPig.DCOMColPContrato
                        Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(!cCtaCod)
                    Set loMuestraContrato = Nothing
                    
                    item = 0
                    Do While Not lrCredPigJoyasDet.EOF
                        lmDetalle(item) = ofun.ImpreFormat(lrCredPigJoyasDet!nPiezas, 4, 0) & " " & ofun.ImpreFormat(lrCredPigJoyasDet!cDescrip, 27) & " " _
                                    & lrCredPigJoyasDet!cKilataje & "K " & ofun.ImpreFormat(lrCredPigJoyasDet!npesoneto, 4, 2) & "Gr"
                        item = item + 1
                        lrCredPigJoyasDet.MoveNext
                    Loop
                  
                    lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & Trim(Left(!codesiaf, 18)) & Space(19 - Len(Trim(Left(!codesiaf, 18)))) & Mid(!cCtaCod, 1, 5) & "-" & Mid(!cCtaCod, 6) & Space(1) & ofun.ImpreFormat(!nPiezas, 3, 0) & Space(1) _
                                        & lmDetalle(0) & ofun.ImpreFormat(v14, 6) & ofun.ImpreFormat(v16, 6) & ofun.ImpreFormat(v18, 6) & ofun.ImpreFormat(v21, 6) _
                                        & ofun.ImpreFormat(!NREMSUBBASEVTA, 8) & oFunI.gPrnSaltoLinea
                    
                    For itemTMP = 1 To item - 1
                            lsCadImp = lsCadImp & Space(49) & lmDetalle(itemTMP) & oFunI.gPrnSaltoLinea
                            lnLineas = lnLineas + 1
                    Next
                    

                Else ' **
                    lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & Trim(Left(!cNomCliente, 18)) & Space(19 - Len(Trim(Left(!cNomCliente, 18)))) & !cCtaCod & Space(1) & ofun.ImpreFormat(!nPiezas, 3, 0) & Space(1) _
                                        & ofun.ImpreFormat(ofun.QuiebreTexto(!cLote, 1), 42, 1) & ofun.ImpreFormat(v14, 6) & ofun.ImpreFormat(v16, 6) & ofun.ImpreFormat(v18, 6) & ofun.ImpreFormat(v21, 6) _
                                        & ofun.ImpreFormat(lnPreVenta, 8) & oFunI.gPrnSaltoLinea
                    
                    If lnIndice Mod 300 = 0 Then
                        lsCadBuffer = lsCadBuffer & lsCadImp
                        lsCadImp = ""
                    End If
                                            
    
                    Do While vCuenta < 15
                        vCuenta = vCuenta + 1
                        If Len(ofun.QuiebreTexto(!cLote, vCuenta + 1)) > 0 Then
                            lsCadImp = lsCadImp & Space(49) & ofun.QuiebreTexto(!cLote, vCuenta + 1) & oFunI.gPrnSaltoLinea
                            lnLineas = lnLineas + 1
                        End If
                    Loop
                
                End If
                lnLineas = lnLineas + 1
                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
                
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "LISTADO DE CONTRATOS PARA REMATE NRO " & psNroRemate, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 145, "", 11)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(145, "*") & oFunI.gPrnSaltoLinea
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImpListadoRemateAnteriores = lsCadBuffer
    Set oFunI = Nothing
End Function

Public Function nListadoParaRemConSIAF(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal pnDiasVctoParaRemate As Integer, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, ByVal pnFactorPrecioVenta As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroRemate As String, _
        ByVal pnJoyasDetalle As Boolean) As ADODB.Recordset
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
Dim fnVarIGV As Double



Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim loCalculos As NCOMColPCalculos

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double
    
Dim lrCredPigJoyasDet As New ADODB.Recordset
Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
Dim item As Integer, itemTMP As Integer
Dim lmDetalle(20) As String
Dim ldFecAnt As Date


    ldFecAviso = DateAdd("d", -pnDiasVctoParaRemate, psFecRemate)
    
    ldFecAnt = DateAdd("d", -1, psFecRemate)



 lsSQL = "           SELECT CODESIAF=ISNULL(RE.CCTACODANT,'                  '),P.cCtaCod, C.dVenc, CP.nTasacion,"
 lsSQL = lsSQL & " C.nMontoCol,  P.nSaldocap, A.NREMSUBBASEVTA, C.cLineaCred, CP.nPiezas, CP.cLote,"
 lsSQL = lsSQL & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 ,"
 lsSQL = lsSQL & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 ,"
 lsSQL = lsSQL & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 ,"
 lsSQL = lsSQL & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 ,"
 lsSQL = lsSQL & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa   WHERE cLineaCred = C.cLineaCred"
 lsSQL = lsSQL & "                  AND nColocLinCredTasaTpo = 3 ), "
 lsSQL = lsSQL & " cNomCliente = (SELECT MAX(cPersNombre) FROM Persona Per INNER JOIN ProductoPersona PPer   ON Per.cPersCod = PPer.cPersCod WHERE PPer.cCtaCod = P.cCtaCod ) "
 lsSQL = lsSQL & " FROM (select * from colocacsaldo where convert(char(8),dfecha,112)='" & Format(ldFecAnt, "yyyymmdd") & "' and cctacod like '108__305%' ) P"
 lsSQL = lsSQL & " Inner Join Colocaciones C ON  C.cCtaCod=P.cCtaCod "
 lsSQL = lsSQL & " inner join ( select r.*,d.cctacod,d.NREMSUBBASEVTA from  colocpigrgdet d join colocpigrecgar r on r.cnroproceso=d.cnroproceso and d.ctpoproceso=r.ctpoproceso where r.ctpoproceso='R' and   convert(char(8),dproceso,112)='" & Format(psFecRemate, "yyyymmdd") & "'"
 lsSQL = lsSQL & "        ) a on a.cctacod=p.cctacod"
 lsSQL = lsSQL & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod"
 lsSQL = lsSQL & " LEFT JOIN RELCUENTAS RE ON RE.CCTACOD=P.CCTACOD "
 lsSQL = lsSQL & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod"
 lsSQL = lsSQL & " where substring(p.cctacod,4,2)='" & psAgencia & "'"
 lsSQL = lsSQL & " GROUP BY RE.CCTACODANT,P.cCtaCod, C.dVenc, CP.nTasacion, A.NREMSUBBASEVTA,C.nMontoCol, P.nSaldocap,  CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote"
 lsSQL = lsSQL & " ORDER BY P.cCtaCod"



    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)

        Set nListadoParaRemConSIAF = lrDataRep
        
        
    Set loDataRep = Nothing
    Set lrDataRep = Nothing




End Function




'Imprime Listado para Adj
Public Function nImprimeListadoParaAdjConSIAF(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal psNroProcesoAdjud As String, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal pnJoyasDetalle As Boolean, ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
Dim fnVarIGV As Double



Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim loCalculos As NCOMColPCalculos

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double
    
Dim lrCredPigJoyasDet As New ADODB.Recordset
Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
Dim item As Integer, itemTMP As Integer
Dim lmDetalle(20) As String

Dim ofun As New COMFunciones.FCOMImpresion

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

Dim lsProcesoListar As String
lsProcesoListar = IIf(psNroProcesoAdjud = "0000", "", "AND CRGDet.cNroProceso ='" & psNroProcesoAdjud & "'")

    lsSQL = " SELECT CODESIAF=ISNULL(RE.CCTACODANT,'                  '),P.cCtaCod, P.nPrdEstado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote, CRGDet.nAdjValRegistro ,  " _
          & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , " _
          & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , " _
          & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , " _
          & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , " _
          & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  " _
          & "               WHERE cLineaCred = C.cLineaCred  " _
          & "               AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " )    " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " LEFT JOIN RELCUENTAS RE ON RE.CCTACOD=P.CCTACOD " _
          & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod  " _
          & " INNER JOIN ColocPigRGDet CRGDet On P.cCtaCod = CRGDet.cCtaCod  " _
          & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod  " _
          & " WHERE CRGDet.cTpoProceso ='A' AND Substring(P.cCtaCod,4,2) = '" & psAgencia & "'" _
          & lsProcesoListar _
          & " And P.nPrdEstado= " & gColPEstAdjud _
          & " GROUP BY P.cCtaCod, P.nPrdestado,ISNULL(RE.CCTACODANT,'                  ') ,C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote,CRGDet.nAdjValRegistro  " _
          & " ORDER BY P.cCtaCod "

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing





    If lrDataRep Is Nothing Then
         psmensaje = " No existen Créditos Pignoraticios para Aviso de Vencimiento "
        'MuestraImpresion = False

        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nImprimeAdjudicaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "LISTADO DE CONTRATOS PARA ADJUDICACION NRO " & psNroProcesoAdjud, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 144, "", 11)

        lnIndice = 1:        lnLineas = 7
        

        'Set loCalculos = New COMDColocPig.DCOMColPCalculos
        
      
        
                
        fnVarIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)

        Dim pnDeudaCosto As Double, pn15DeVenta As Double

        
        With lrDataRep
            Do While Not lrDataRep.EOF
                
                v14 = !nK14: v14 = Round(v14, 2)
                v16 = !nK16: v16 = Round(v16, 2)
                v18 = !nK18: v18 = Round(v18, 2)
                v21 = !nK21: v21 = Round(v21, 2)
                vCuenta = 0
                
'                pn15DeVenta = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * 0.15
'
'                lnPreBase = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * (fnVarIGV + 1)
'
'                'lnPreBase = (v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)
'
'
'                lnGramos = v14 + v16 + v18 + v21
'
'                lnPreBase = Round(lnPreBase, 2)
'
'                vdiasAtraso = DateDiff("d", !DVENC, psFecRemate)
'
'                Set loCalculos = New NCOMColPCalculos
'                    If vdiasAtraso <= 0 Then
'                        vdiasAtraso = 0
'                        vInteresVencido = 0
'                        vCostoCustodiaMoratorio = 0
'                        vImpuesto = 0
'                    Else
'                        vInteresVencido = loCalculos.nCalculaInteresMoratorio(!nSaldo, !nTasaIntVenc, vdiasAtraso)
'                        vInteresVencido = Round(vInteresVencido, 2)
'                        vCostoCustodiaMoratorio = loCalculos.nCalculaCostoCustodiaMoratorio(!nTasacion, pnTasaCustodiaVencida, vdiasAtraso)
'                        vCostoCustodiaMoratorio = Round(vCostoCustodiaMoratorio, 2)
'                        vImpuesto = (vInteresVencido + vCostoCustodiaMoratorio) * pnTasaImpuesto
'                        vImpuesto = Round(vImpuesto, 2)
'                    End If
'                    vCostoPreparacionRemate = 0
'                    If !nPrdEstado = gColPEstPRema Then    ' Si esta en via de Remate
'                        vCostoPreparacionRemate = pnTasaPreparacionRemate * !nTasacion
'                        vCostoPreparacionRemate = Round(vCostoPreparacionRemate, 2)
'                    End If
'                    lnDeuda = !nSaldo + vInteresVencido + vCostoCustodiaMoratorio + vImpuesto + vCostoPreparacionRemate
'
'
'                   ' pnDeudaCosto = lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta
'                     pnDeudaCosto = Round((lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta) * (fnVarIGV + 1), 2)
'                    'Calcula el Precio Venta para Remate
'                    'lnPreVenta = loCalculos.nCalculaPrecioVentaRemate(pnDeudaCosto, lnPreBase, pnFactorPrecioVenta)
                
                lnPreVenta = !nAdjValRegistro
                Set loCalculos = Nothing
                
                'Joyas Detalladas
                lnLineas = lnLineas + 1
                If pnJoyasDetalle = True Then

                    Set loMuestraContrato = New COMDColocPig.DCOMColPContrato
                        Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(!cCtaCod)
                    Set loMuestraContrato = Nothing
                    
                    item = 0
                    Do While Not lrCredPigJoyasDet.EOF
                        lmDetalle(item) = ofun.ImpreFormat(lrCredPigJoyasDet!nPiezas, 4, 0) & " " & ofun.ImpreFormat(lrCredPigJoyasDet!cDescrip, 27) & " " _
                                    & lrCredPigJoyasDet!cKilataje & "K " & ofun.ImpreFormat(lrCredPigJoyasDet!npesoneto, 4, 2) & "Gr"
                        item = item + 1
                        lrCredPigJoyasDet.MoveNext
                    Loop
                  
                    lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !codesiaf & Space(1) & Mid(!cCtaCod, 1, 5) & "-" & Mid(!cCtaCod, 6) & Space(1) & ofun.ImpreFormat(!nPiezas, 3, 0) & Space(1) _
                                        & lmDetalle(0) & ofun.ImpreFormat(v14, 6) & ofun.ImpreFormat(v16, 6) & ofun.ImpreFormat(v18, 6) & ofun.ImpreFormat(v21, 6) _
                                        & ofun.ImpreFormat(lnPreVenta, 8) & oFunI.gPrnSaltoLinea
                    
                    For itemTMP = 1 To item - 1
                            lsCadImp = lsCadImp & Space(49) & lmDetalle(itemTMP) & oFunI.gPrnSaltoLinea
                            lnLineas = lnLineas + 1
                    Next
                    

                Else ' **
                    lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !codesiaf & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!nPiezas, 3, 0) & Space(1) _
                                        & ofun.ImpreFormat(ofun.QuiebreTexto(!cLote, 1), 42, 1) & ofun.ImpreFormat(v14, 6) & ofun.ImpreFormat(v16, 6) & ofun.ImpreFormat(v18, 6) & ofun.ImpreFormat(v21, 6) _
                                        & ofun.ImpreFormat(lnPreVenta, 8) & oFunI.gPrnSaltoLinea
                    
                    If lnIndice Mod 300 = 0 Then
                        lsCadBuffer = lsCadBuffer & lsCadImp
                        lsCadImp = ""
                    End If
                                            
    
                    Do While vCuenta < 15
                        vCuenta = vCuenta + 1
                        If Len(ofun.QuiebreTexto(!cLote, vCuenta + 1)) > 0 Then
                            lsCadImp = lsCadImp & Space(40) & ofun.QuiebreTexto(!cLote, vCuenta + 1) & oFunI.gPrnSaltoLinea
                            lnLineas = lnLineas + 1
                        End If
                    Loop
                
                End If
                lnLineas = lnLineas + 1
                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
                
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nImprimeAdjudicaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "LISTADO DE CONTRATOS PARA ADJUDICACION NRO " & psNroProcesoAdjud, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 144, "", 11)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(120, "*") & oFunI.gPrnSaltoLinea
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimeListadoParaAdjConSIAF = lsCadBuffer
    Set oFunI = Nothing
End Function


Public Function nListadoParaAdjConSIAF(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal psNroProcesoAdjud As String, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal pnJoyasDetalle As Boolean) As ADODB.Recordset
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
Dim fnVarIGV As Double



Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim loCalculos As NCOMColPCalculos

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double
    
Dim lrCredPigJoyasDet As New ADODB.Recordset
Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
Dim item As Integer, itemTMP As Integer
Dim lmDetalle(20) As String


Dim lsProcesoListar As String
lsProcesoListar = IIf(psNroProcesoAdjud = "0000", "", "AND CRGDet.cNroProceso ='" & psNroProcesoAdjud & "'")


    'ldFecAviso = DateAdd("d", -pnDiasVctoParaRemate, psFecRemate)
    

    lsSQL = " SELECT CODESIAF=ISNULL(RE.CCTACODANT,'                  '),P.cCtaCod, P.nPrdEstado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote, CRGDet.nAdjValRegistro ,  " _
          & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , " _
          & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , " _
          & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , " _
          & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , " _
          & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  " _
          & "               WHERE cLineaCred = C.cLineaCred  " _
          & "               AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " )    " _
          & " , cNomCliente = (SELECT MAX(cPersNombre) FROM Persona Per INNER JOIN ProductoPersona PPer   ON Per.cPersCod = PPer.cPersCod WHERE PPer.cCtaCod = P.cCtaCod ) " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " LEFT JOIN RELCUENTAS RE ON RE.CCTACOD=P.CCTACOD " _
          & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod  " _
          & " INNER JOIN ColocPigRGDet CRGDet On P.cCtaCod = CRGDet.cCtaCod  " _
          & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod  " _
          & " WHERE CRGDet.cTpoProceso ='A' AND Substring(P.cCtaCod,4,2) = '" & psAgencia & "'" _
          & lsProcesoListar _
          & " GROUP BY P.cCtaCod, P.nPrdestado,ISNULL(RE.CCTACODANT,'                  ') ,C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote,CRGDet.nAdjValRegistro  " _
          & " ORDER BY P.cCtaCod "
          '& " And P.nPrdEstado= " & gColPEstAdjud _

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
        
        
        Set nListadoParaAdjConSIAF = lrDataRep
        
        
    Set loDataRep = Nothing
    Set lrDataRep = Nothing




End Function

Public Function nIniciaRemate(ByVal psNroRemate As String, ByVal psFecRemate As String, _
        ByVal psProcesoCadaAgencia As String, ByVal pnDiasVctoParaRemate As Integer, ByVal pnFactorPrecioVenta As Double, ByVal pnPrecioOro14 As Double, ByVal pnPrecioOro16 As Double, _
        ByVal pnPrecioOro18 As Double, ByVal pnPrecioOro21 As Double, ByVal pnTasaCustodiaVenc As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, Optional ByVal psImpresora As Impresoras = gEPSON) As Boolean

Dim lsSQL As String
Dim loDataIniRem As COMDColocPig.DCOMColPFunciones
Dim lrDataInicia As New ADODB.Recordset
              
Dim loCalcula As NCOMColPCalculos
Dim loRegPig As COMDColocPig.DCOMColPActualizaBD
Dim ldFecLimiteRemate As Date

Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
Dim fnVarIGV As Double
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency


Dim lnPreBase As Currency, lnPreDeuda As Currency, lnPreVenta As Currency

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


ldFecLimiteRemate = DateAdd("d", -pnDiasVctoParaRemate, Format(psFecRemate, "mm/dd/yyyy"))

'Arturo
lsSQL = " SELECT c.cCtacod, c.dVenc, cp.nTasacion, c.nMontoCol, p.nSaldo, p.nprdEstado, C.cLineaCred, CP.nNroRemate, " _
    & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , " _
    & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , " _
    & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , " _
    & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , " _
    & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  " _
    & "               WHERE cLineaCred = C.cLineaCred  " _
    & "               AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ) " _
    & " FROM Producto p Inner Join Colocaciones c on p.cctacod = c.cctacod " _
    & " Inner join ColocPignoraticio cp on c.cctacod = cp.cctacod " _
    & " Left Join ColocPigjoya cpjoy on cp.cctacod = cpjoy.cctacod " _
    & " WHERE DATEDIFF(dd, C.dVenc, '" & Format(ldFecLimiteRemate, "mm/dd/yyyy") & "') >= 0 " _
    & " And p.nPrdEstado IN (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & " )  " _
    & IIf(Mid(psProcesoCadaAgencia, 4, 2) = "00", "", " And Substring(p.cCtaCod,1,5) = '" & psProcesoCadaAgencia & "'") & " " _
    & " GROUP By c.cCtacod, c.dVenc, cp.nTasacion, c.nMontoCol, p.nSaldo, p.nprdEstado, C.cLineaCred, CP.nNroRemate  "
    
Set loDataIniRem = New COMDColocPig.DCOMColPFunciones
    Set lrDataInicia = loDataIniRem.dObtieneRecordSet(lsSQL)
Set loDataIniRem = Nothing

If lrDataInicia Is Nothing Then

Else
    lnPreBase = 0: lnPreDeuda = 0: lnPreVenta = 0
    
        fnVarIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)
        Dim pnDeudaCosto As Double, pn15DeVenta As Double
    
    With lrDataInicia
        Set loRegPig = New COMDColocPig.DCOMColPActualizaBD
        
        Do While Not .EOF
        
                v14 = !nK14: v14 = Round(v14, 2)
                v16 = !nK16: v16 = Round(v16, 2)
                v18 = !nK18: v18 = Round(v18, 2)
                v21 = !nK21: v21 = Round(v21, 2)
        
        
            'Precio Base
            pn15DeVenta = ((v14 * pnPrecioOro14) + (v16 * pnPrecioOro16) + (v18 * pnPrecioOro18) + (v21 * pnPrecioOro21)) * 0.15
            lnPreBase = ((v14 * pnPrecioOro14) + (v16 * pnPrecioOro16) + (v18 * pnPrecioOro18) + (v21 * pnPrecioOro21)) '* (fnVarIGV + 1)

            lnPreBase = Round(lnPreBase, 2)

            '** Actualiza Producto
            Call loRegPig.dUpdateProducto(!cCtaCod, , , gColPEstPRema, , , False)         ' (-2) Cambia el estado a Para Remate

            Set loCalcula = New NCOMColPCalculos
                'Deuda
                lnPreDeuda = loCalcula.nCalculaDeudaPignoraticio(!nSaldo, !dVenc, !nTasacion, !nTasaIntVenc, _
                    pnTasaCustodiaVenc, pnTasaImpuesto, gColPEstPRema, pnTasaPreparacionRemate, Format(psFecRemate, "mm/dd/yyyy"))
                
                lnPreDeuda = Round(lnPreDeuda, 2)
                'pnDeudaCosto = lnPreDeuda + (lnPreDeuda * pnFactorPrecioVenta) + pn15DeVenta
                
                lnPreVenta = loCalcula.nCalculaPrecioVentaRemate(pnDeudaCosto, lnPreBase, pnFactorPrecioVenta)
                lnPreVenta = CalculaPrecioVentaRemate(Round(lnPreVenta, 0))
                
                'Calcula el Precio Venta para Remate
                'lnPreVenta = loCalcula.nCalculaPrecioVentaRemate(lnPreDeuda, lnPreBase, pnFactorPrecioVenta)
                'lnPreVenta = Round(lnPreVenta, 2)
            Set loCalcula = Nothing

            'Grabar información en Detalle de Remate
            
            '** Actualiza ColocPignoraticio
            Call loRegPig.dUpdateColocPignoraticio(!cCtaCod, , , , , , !nNroRemate + 1, , , , , False)
            
            Call loRegPig.dInsertColocPigRGDet("R", psNroRemate, !cCtaCod, gColPRecGarVentaEnProceso, 0, lnPreVenta, 0, 0, lnPreDeuda, psProcesoCadaAgencia, False)
                
            .MoveNext
        Loop
        .Close
        Set lrDataInicia = Nothing
    End With
            
    Call loRegPig.dUpdateColocPigRecGar("R", psNroRemate, psProcesoCadaAgencia, , gColPRecGarEstIniciado, , False)
     
    Set loRegPig = Nothing
    Set oFunI = Nothing
End If

End Function

Public Function nFinalizaRemate(ByVal psNroRemate As String, ByVal psFecProceso As String, ByVal psCtaAhorroRemate As String, _
        ByVal pnMontoSobrante As Currency, ByVal psProcesoCadaAgencia As String, ByVal psCodUser As String, _
        ByVal psLpt As String) As Boolean

Dim lsSQL As String
Dim lsNextRemate As String

Dim loRegPig As COMDColocPig.DCOMColPActualizaBD
Dim loNroPro As COMDColocPig.DCOMColPFunciones
Dim oCont As COMNContabilidad.NCOMContFunciones

Dim Rs As New ADODB.Recordset
   
    'lsNextRemate = String(4 - Len(Trim(Str(Val(psNroRemate) + 1))), "0") & Trim(Str(Val(psNroRemate) + 1))
    
    'Genera el sgte Nro de Remate
    Set loNroPro = New COMDColocPig.DCOMColPFunciones
        lsSQL = "SELECT max(cNroProceso) nNextProceso FROM ColocPigRecGar RG  WHERE cTpoProceso ='R'"
        Set Rs = loNroPro.dObtieneRecordSet(lsSQL)
        If Not (Rs.EOF And Rs.BOF) Then
            lsNextRemate = String(4 - Len(Trim(Str(Val(Rs!nNextProceso) + 1))), "0") & Trim(Str(Val(Rs!nNextProceso) + 1)) 'Rs!nNextProceso
        End If
    Set loNroPro = Nothing

    Set loRegPig = New COMDColocPig.DCOMColPActualizaBD
        'Cambia estado de los que NO se han vendido
        lsSQL = " UPDATE ColocPigRGDet SET nRGDetEstado = " & gColPRecGarVentaNoVendido _
               & " WHERE cTpoProceso = 'R' AND cNroProceso = '" & psNroRemate & "' " _
               & " AND nRGDetEstado = " & gColPRecGarVentaEnProceso & " AND cCodAge = '" & psProcesoCadaAgencia & "' "
        Call loRegPig.dEjecutar(lsSQL)

        'Termina el remate
        Call loRegPig.dUpdateColocPigRecGar("R", psNroRemate, psProcesoCadaAgencia, , gColPRecGarEstTerminado, , False)

        'Inserta el Sgte remate
        Call loRegPig.dInsertColocPigRecGar("R", lsNextRemate, Format(DateAdd("m", 1, Format(psFecProceso, "mm/dd/yyyy")) & " 08:00", "mm/dd/yyyy hh:mm"), gColPRecGarEstNoIniciado, psProcesoCadaAgencia, , False)

        'Deposita a la Cta de Ahorros de sobrante
        '****
        Dim oCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
        Dim sMovNroReg As String
        Dim nMovNroReg As Long
        Set oCont = New COMNContabilidad.NCOMContFunciones
            sMovNroReg = oCont.GeneraMovNro(psFecProceso, psProcesoCadaAgencia, psCodUser)
        Set oCont = Nothing
        Set oCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
            'oCapMov.CapAbonoCuentaAho psCtaAhorroRemate, pnMontoSobrante, gAhoDepEfec, sMovNroReg, "Dep. Sobrante Remate-" & psNroRemate & " - " & psProcesoCadaAgencia, , , , , , , , , , psLpt, , True
            oCapMov.CapAbonoCuentaAho psCtaAhorroRemate, pnMontoSobrante, gAhoDepSobRemate, sMovNroReg, "Dep. Sobrante Remate-" & psNroRemate & " - " & psProcesoCadaAgencia, , , , , , , , , , psLpt, , True
        Set oCapMov = Nothing
        'nMovNroReg = loRegPig.dGetnMovNro(sMovNroReg)
        'Call loRegPig.dInsertMovRef(lnMovNro, nMovNroReg)
        '****
    Set loRegPig = Nothing

End Function

'Imprime Planilla de Ventas en Remate
Public Function nImprimePlanillaVentaRemate(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal pnDiasVctoParaRemate As Integer, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, ByVal pnFactorPrecioVenta As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroRemate As String, ByRef psmensaje As String, _
        Optional ByVal psImpresora As Impresoras = gEPSON) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim loCalculos As NCOMColPCalculos

Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency
    
Dim lsNombre As String * 28
Dim vSumGramos As Currency, vSumTasaci As Currency, vSumPresta As Currency
Dim vSumVtaBru As Currency, vSumIGV As Currency, vSumComisRem As Currency
Dim vSumVtaNet As Currency, vSumSobrante As Currency, vSumSalCap As Currency
Dim vSumDeuda As Currency
Dim vSumOroNeto As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double

Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
    
    
' , CRGDet.nSobrante
    lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, CP.nOroNeto, CRGDet.nMontoProceso, CRGDet.nDeuda,  " _
          & " SUM( case when nPrdConceptoCod  =" & gColPConceptoCodComisionRemate & " then nMonto else 0  end ) as nComisionRemate , " _
          & " SUM( case when nPrdConceptoCod  =" & gColPConceptoCodSobranteRemate & " then nMonto else 0  end ) as nSobranteRemate , " _
          & " SUM( case when nPrdConceptoCod  = 2216 then nMonto else 0  end ) as nIGV   " _
          & " " _
          & " " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " INNER JOIN ColocPigRGDet CRGDet ON P.cCtaCod = CRGDet.cCtaCod " _
          & " INNER JOIN MovColDet MCDet On CP.cCtaCod = MCDet.cCtaCod And MCDet.nMovNro = CRGDet.nMovVta  " _
          & " WHERE CRGDet.cTpoProceso ='R' AND CRGDet.cNroProceso='" & psNroRemate & "' " _
          & " AND CRGDet.nRGDetEstado in (" & gColPRecGarVentaVendido & "," & gColPRecGarVentaSobranteCobrado & " ) " _
          & " AND Substring(P.cCtaCod,4,2) ='" & psAgencia & "'  " _
          & " GROUP BY P.cCtaCod, P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda,CP.nOroNeto, CRGDet.nMontoProceso, CRGDet.nDeuda " _
          & " ORDER BY P.cCtaCod "

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing


    If lrDataRep Is Nothing Then
         psmensaje = " No existen Créditos Pignoraticios para Aviso de Vencimiento "
        'MuestraImpresion = False

        Exit Function
    Else
        lnLineas = 0
        lnPage = 1
        
        'CABECERA
        lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "PLANILLA DE PRENDAS VENDIDAS EN REMATE NRO " & psNroRemate, "Del " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 140, "", 3)

        lnIndice = 1:        lnLineas = 7
        
        vSumGramos = 0:        vSumTasaci = 0
        vSumPresta = 0:        vSumVtaBru = 0
        vSumVtaNet = 0:        vSumIGV = 0
        vSumComisRem = 0:      vSumSobrante = 0
        vSumSalCap = 0:        vSumDeuda = 0
        vSumOroNeto = 0
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                
                lnLineas = lnLineas + 1
                    
                vSumGramos = vSumGramos + lnGramos
                vSumTasaci = vSumTasaci + !nTasacion
                vSumPresta = vSumPresta + !nMontoCol
                vSumSalCap = vSumSalCap + !nSaldo
                vSumDeuda = vSumDeuda + !nDeuda
                vSumVtaBru = vSumVtaBru + !nMontoProceso
                vSumIGV = vSumIGV + !nIGV
                vSumComisRem = vSumComisRem + !nComisionRemate
                vSumVtaNet = vSumVtaNet + !nMontoProceso - !nComisionRemate
                vSumSobrante = vSumSobrante + !nSobranteRemate
                vSumOroNeto = vSumOroNeto + !nOroNeto
                                        
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 4, 0) & Space(1) & !cCtaCod & Space(1) & Format(!dVenc, "dd/mm/yyyy") _
                    & ofun.ImpreFormat(!nTasacion, 8) & ofun.ImpreFormat(!nMontoCol, 8) & ofun.ImpreFormat(!nSaldo, 8) _
                    & ofun.ImpreFormat(!nDeuda, 8) & ofun.ImpreFormat(!nMontoProceso, 8) _
                    & ofun.ImpreFormat(!nIGV, 5) _
                    & ofun.ImpreFormat(!nComisionRemate, 6) & ofun.ImpreFormat(!nMontoProceso - !nComisionRemate, 8) _
                    & ofun.ImpreFormat(!nSobranteRemate, 8) & ofun.ImpreFormat(!nOroNeto, 4) _
                    & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                        
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "PLANILLA DE PRENDAS VENDIDAS EN REMATE NRO " & psNroRemate, "Del " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 140, "", 3)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(12) & "RESUMEN " & Space(14) & ofun.ImpreFormat(vSumTasaci, 8) _
                & ofun.ImpreFormat(vSumPresta, 8) & ofun.ImpreFormat(vSumSalCap, 8) & ofun.ImpreFormat(vSumDeuda, 8) _
                & ofun.ImpreFormat(vSumVtaBru, 8) & ofun.ImpreFormat(vSumIGV, 5) _
                & ofun.ImpreFormat(vSumComisRem, 6) & ofun.ImpreFormat(vSumVtaBru - vSumComisRem, 8) & ofun.ImpreFormat(vSumSobrante, 8) _
                & ofun.ImpreFormat(vSumOroNeto, 4) & oFunI.gPrnSaltoLinea
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimePlanillaVentaRemate = lsCadBuffer
    Set oFunI = Nothing
End Function

'Imprime Planilla de Lotes No vendidos en Remate
Public Function nImprimePlanillaNoVendidosRemate(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal pnDiasVctoParaRemate As Integer, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, ByVal pnFactorPrecioVenta As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroRemate As String, ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
    
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
    
Dim lsNombre As String * 28
Dim vSumTasaci As Currency, vSumPresta As Currency
Dim vSumSalCap As Currency, vSumDeuda As Currency, vSumOroNeto As Currency

Dim ofun As New COMFunciones.FCOMImpresion

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
    

    lsSQL = " SELECT  CODESIAF=ISNULL(RE.CCTACODANT,'                  '),P.cCtaCod, P.nPrdEstado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, CP.nOroNeto, CRGDet.nMontoProceso, CRGDet.nDeuda   " _
          & " " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " LEFT JOIN RELCUENTAS RE ON RE.CCTACOD=P.CCTACOD " _
          & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " INNER JOIN ColocPigRGDet CRGDet ON P.cCtaCod = CRGDet.cCtaCod " _
          & "  " _
          & " WHERE CRGDet.cTpoProceso ='R' AND CRGDet.cNroProceso='" & psNroRemate & "' " _
          & " AND CRGDet.nRGDetEstado in (" & gColPRecGarVentaNoVendido & " ) " _
          & " AND Substring(P.cCtaCod,4,2) ='" & psAgencia & "'  " _
          & " GROUP BY P.cCtaCod,ISNULL(RE.CCTACODANT,'                  '), P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda,CP.nOroNeto, CRGDet.nMontoProceso, CRGDet.nDeuda " _
          & " ORDER BY P.cCtaCod "

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing


    If lrDataRep Is Nothing Then
        psmensaje = " No existen Créditos Pignoraticios No Rematados en la Agencia"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1
        
        'CABECERA
        lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "PLANILLA DE PRENDAS NO VENDIDAS DEL REMATE NRO " & psNroRemate, "Del " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 105, "", 4)

        lnIndice = 1:        lnLineas = 7
        
        vSumTasaci = 0
        vSumPresta = 0:
        vSumSalCap = 0:
        vSumDeuda = 0: vSumOroNeto = 0
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                
                lnLineas = lnLineas + 1
                    
                vSumTasaci = vSumTasaci + !nTasacion
                vSumPresta = vSumPresta + !nMontoCol
                vSumSalCap = vSumSalCap + !nSaldo
                vSumDeuda = vSumDeuda + !nDeuda
                vSumOroNeto = vSumOroNeto + !nOroNeto
                                        
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 4, 0) & Space(1) & !codesiaf & Space(1) & !cCtaCod & Space(1) & Format(!dVenc, "dd/mm/yyyy") _
                    & ofun.ImpreFormat(!nTasacion, 8) & ofun.ImpreFormat(!nMontoCol, 8) & ofun.ImpreFormat(!nSaldo, 8) _
                    & ofun.ImpreFormat(!nDeuda, 8) & ofun.ImpreFormat(!nOroNeto, 5) _
                    & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                        
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "PLANILLA DE PRENDAS NO VENDIDAS DEL REMATE NRO " & psNroRemate, "Del " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 105, "", 4)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(12) & "RESUMEN " & Space(33) & ofun.ImpreFormat(vSumTasaci, 8) _
                & ofun.ImpreFormat(vSumPresta, 8) & ofun.ImpreFormat(vSumSalCap, 8) & ofun.ImpreFormat(vSumDeuda, 8) _
                & ofun.ImpreFormat(vSumOroNeto, 5) _
                & oFunI.gPrnSaltoLinea
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nImprimePlanillaNoVendidosRemate = lsCadBuffer
    Set oFunI = Nothing
End Function


Public Function nObtieneListadoProcesosRG(ByVal psTpoProceso As String, ByVal psProcesoCadaAgencia As String, _
                Optional pbOrdenAsc As Boolean = True) As ADODB.Recordset
'Obtiene Datos de Remate Credito Pignoraticio
Dim loValida As COMDColocPig.DCOMColPFunciones
Dim lrs As New ADODB.Recordset
Dim lsSQL As String
Dim lsAgencia As String

On Error GoTo dError
    lsSQL = "SELECT RG.cTpoProceso, RG.cNroProceso, RG.dProceso, RG.nRGEstado " _
        & " FROM ColocPigRecGar RG " _
        & " WHERE RG.cTpoProceso ='" & psTpoProceso & "' And cCodAge = '" & psProcesoCadaAgencia & "' " _
        & " ORDER BY RG.cNroProceso " & IIf(pbOrdenAsc = True, "ASC", "DESC")
    Set loValida = New COMDColocPig.DCOMColPFunciones
        Set lrs = loValida.dObtieneRecordSet(lsSQL)
    Set loValida = Nothing
    Set nObtieneListadoProcesosRG = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function

'Imprime Aviso de Sobrante de Remate
Public Function nImprimeAvisoSobranteRemate(ByVal psNumRemate As String, ByVal psAgencia As String, ByVal psTextoCarta As String, _
       ByVal pnPrevioMax As Long, ByVal psFechaSis As String, ByRef psmensaje As String, Optional psImpresora As Impresoras = gEPSON) As String
Dim LALISTA() As String
Dim lnFilas As Long, lnItem As Long
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCartaModelo As String
Dim ofcar As New COMFunciones.FCOMCadenas

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


On Error GoTo dError

    
    lsSQL = "SELECT Per.cPersNombre, Per.cPersDireccDomicilio, Per.cPersTelefono, Per.cPersDireccUbiGeo, " _
        & " Ubigeo.cUbiGeoDescripcion, C.cCtaCod, P.nPrdEstado, CRRVta.nSobrante " _
        & " FROM Producto P JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod " _
        & " JOIN ProductoPersona PP ON c.cCtaCod = PP.cCtaCod " _
        & " JOIN Persona Per ON Per.cPersCod = PP.cPersCod " _
        & " JOIN ColocPigRGDet CRGDet ON P.cCtaCod = CRGDet.cCtaCod " _
        & " JOIN ColocPigRecupVta CRRVta ON P.cCtaCod = CRRVta.cCtaCod " _
        & " LEFT JOIN UbicacionGeografica Ubigeo ON Per.cPersDireccUbiGeo = Ubigeo.cUbiGeoCod " _
        & " WHERE CRGDet.cTpoProceso ='R' AND CRGDet.cNroProceso='" & psNumRemate & "' " _
        & " AND CRGDet.nRGDetEstado in (" & gColPRecGarVentaVendido & " ) " _
        & " And Substring(P.cCtaCod,4,2) ='" & psAgencia & "' ORDER BY Per.cPersNombre "

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing


    If lrDataRep Is Nothing Then
        psmensaje = " No existen Créditos Pignoraticios para Carta de Sobrante"
        Exit Function
    Else
        lnFilas = 0
        ReDim LALISTA(lrDataRep.RecordCount, 5)
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnFilas = lnFilas + 1
                If !cPersNombre = LALISTA(lnFilas - 1, 0) Then
                    LALISTA(lnFilas - 1, 3) = LALISTA(lnFilas - 1, 3) & " - " & !cCtaCod
                    LALISTA(lnFilas - 1, 4) = LALISTA(lnFilas - 1, 4) & " - " & Format(!nSobrante, "#,##0.00")
                    lnFilas = lnFilas - 1
                Else
                    LALISTA(lnFilas, 0) = !cPersNombre
                    LALISTA(lnFilas, 1) = !cPersDireccDomicilio & " - " & !cUbiGeoDescripcion '  ClienteZona(!cCodZon) & " - " & ClienteCiudad(!cCodZon)
                    LALISTA(lnFilas, 2) = !cPersTelefono & ""
                    LALISTA(lnFilas, 3) = !cCtaCod
                    LALISTA(lnFilas, 4) = !nSobrante
                End If
                .MoveNext
            Loop
        End With
        lrDataRep.Close
        Set lrDataRep = Nothing
    End If
    'Llena cartas
    For lnItem = 1 To lnFilas
        lsCartaModelo = psTextoCarta
        lsCartaModelo = Replace(lsCartaModelo, "<<FECHA>>", Format(psFechaSis, "dddd,d mmmm yyyy"), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<NOMBRE>>", ofcar.PstaNombre(LALISTA(lnItem, 0), False), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<SOBRANTE>>", LALISTA(lnItem, 4), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<CONTRATO>>", LALISTA(lnItem, 3), , 2, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<FECREMATE>>", psNumRemate, , 1, vbTextCompare)
        'RTFTmp = Replace(RTFTmp, "<<HORREMATE>>", vHorRem, , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<DIRECCION>>", LALISTA(lnItem, 1), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<TELEFONO>>", LALISTA(lnItem, 2), , 1, vbTextCompare)
        
        lsCadImp = lsCadImp & lsCartaModelo & oFunI.gPrnSaltoPagina
        'prgList.Value = vCont
    Next lnItem
    nImprimeAvisoSobranteRemate = lsCadImp

Exit Function
   
dError:
    Err.Raise Err.Number, "Obtiene Datos Imprime Listado Aviso Vencimiento  <<nImprimeAvisoVencimiento>>", Err.Description
    
End Function


Public Function nImprimeAdjudicaReporteCabec(ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, _
        ByVal psFechaSis As String, ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String, _
        ByVal pnNroRepor As Integer, Optional ByVal psImpresora As Impresoras = gEPSON) As String
        
Dim lsCadImp As String
Dim loImprimeCab As NCOMColPImpre
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

    
    Set loImprimeCab = New NCOMColPImpre
        lsCadImp = loImprimeCab.nImprimeCabeceraReportes(psNomCmac, psNomAgencia, psCodUser, psFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)
    Set loImprimeCab = Nothing
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & String(pnAnchoLinea, "-") & oFunI.gPrnSaltoLinea
    Select Case pnNroRepor
        Case 1 ' Planilla para Adjudicacion
            ''** PEAC 20090227
            'lsCadImp = lsCadImp & Space(1) & "Item    Contrato      Fec.Venc.     Cliente                 14kl    16kl    18kl     21kl     Gramos    Tasacion    SaldoCap      Deuda     ValRegistro" & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Space(1) & "Item    Contrato      Fec.Venc.     Cliente                 14kl    16kl    18kl     21kl     Gramos    Tasacion    SaldoCap      Deuda" & oFunI.gPrnSaltoLinea
        Case 2 '
            lsCadImp = lsCadImp & Space(1) & "Item    ContratoAnt     Contrato          Cliente                   14kl      16kl     18kl     21kl     Gramos    Tasacion   SaldoCap  ValReg" & oFunI.gPrnSaltoLinea
        Case 9
            lsCadImp = lsCadImp & Space(1) & "Item    ContratoAnt     Contrato           Fec.Venc.     Cliente               14kl   16kl    18kl     21kl     Gramos     Tasacion    SaldoCap     ValReg" & oFunI.gPrnSaltoLinea
        Case 10
            lsCadImp = lsCadImp & Space(1) & "Item    ContratoAnt     Contrato           Cliente                    14kl    16kl    18kl     21kl     Gramos    Tasacion    SaldoCap      Deuda     ValRegistro " & oFunI.gPrnSaltoLinea
        Case 11
            lsCadImp = lsCadImp & Space(1) & "Item    ContratoAnt     Contrato            Piezas     Descripción                                 14Kl.    16Kl.    18Kl.    21Kl.  ValRegistro " & oFunI.gPrnSaltoLinea
    
    End Select
    lsCadImp = lsCadImp & String(pnAnchoLinea, "-") & oFunI.gPrnSaltoLinea
nImprimeAdjudicaReporteCabec = lsCadImp
Set oFunI = Nothing
End Function

'Imprime Planilla para Adjudicacion
Public Function nImprimePlanillaParaAdjudicacion(ByVal psFecProceso As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal pnNroRemParaAdjud As Integer, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroProceso As String, ByVal gdFecSis As Date, ByRef psmensaje As String, _
        Optional ByVal psImpresora As Impresoras = gEPSON, Optional pdFecCorte As Date) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim loCalculos As NCOMColPCalculos

Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer

Dim lnGra14 As Double, lnGra16 As Double, lnGra18 As Double, lnGra21 As Double
Dim lnGramos As Currency, lnTasaci As Currency, lnSalCap As Currency, lnDeuda As Currency
Dim lnPreMerc As Currency, lnValRegis As Currency

Dim lsNombre As String * 28

Dim vSumGra14 As Currency, vSumGra16 As Currency, vSumGra18 As Currency, vSumGra21 As Currency
Dim vSumGramos As Currency, vSumTasaci As Currency, vSumSalCap As Currency, vSumDeuda As Currency
Dim vSumPreMerc As Currency, vSumValRegis As Currency
Dim loCalcula As NCOMColPCalculos
Set loCalcula = New NCOMColPCalculos
    
Dim vdiasAtraso As Integer

Dim ofcar As New COMFunciones.FCOMCadenas
Dim ofun As New COMFunciones.FCOMImpresion

Dim pdFechaLey As Date

''*** PEAC 20090226 ***********************
Dim lnDiasAtraso As Integer
Dim lnIntVencido As Double
Dim lnIntMoratorio As Double
Dim lnIntAdelantado As Double
''*****************************************
Dim lnValorCosCustoDif As Double '*** PEAC 20190725

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


'    lsSQL = " SELECT P.cCtaCod, P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred,  " _
          & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , " _
          & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , " _
          & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , " _
          & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , " _
          & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  " _
          & "               WHERE cLineaCred = C.cLineaCred  " _
          & "               AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ),    " _
          & " cNomCliente = (SELECT MAX(cPersNombre) FROM Persona Per INNER JOIN ProductoPersona PPer " _
          & "               ON Per.cPersCod = PPer.cPersCod WHERE PPer.cCtaCod = P.cCtaCod )  " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod  " _
          & " WHERE P.nPrdEstado in (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ") " _
          & " AND Substring(P.cCtaCod,4,2) ='" & psAgencia & "' AND nNroRemate = " & pnNroRemParaAdjud _
          & " GROUP BY P.cCtaCod, P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda,C.cLineaCred   " _
          & " ORDER BY P.cCtaCod "

    pdFechaLey = "01/06/2006"

    lsSQL = " exec stp_sel_PlanillaParaAdjudicacion '" & psAgencia & "','" & Format(pdFechaLey, "yyyymmdd") & "','" & Format(pdFecCorte, "yyyymmdd") & "','" & psNroProceso & "'"

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Then
        psmensaje = " No existen Créditos Pignoraticios para Aviso de Vencimiento "
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1
        
        'CABECERA
        lsCadImp = lsCadImp & nImprimeAdjudicaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "PLANILLA PARA ADJUDICACION NRO " & psNroProceso, "Del " & Format(psFecProceso, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 155, "", 1)

        lnIndice = 1:        lnLineas = 7
        
        vSumGra14 = 0:  vSumGra16 = 0:  vSumGra18 = 0:  vSumGra21 = 0
        vSumGramos = 0: vSumTasaci = 0: vSumSalCap = 0: vSumDeuda = 0
        vSumPreMerc = 0: vSumValRegis = 0
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                
                lnGra14 = !nK14:  lnGra14 = Round(lnGra14, 2)
                lnGra16 = !nK16:  lnGra16 = Round(lnGra16, 2)
                lnGra18 = !nK18:  lnGra18 = Round(lnGra18, 2)
                lnGra21 = !nK21:  lnGra21 = Round(lnGra21, 2)
                
                lnPreMerc = (lnGra14 * pnPreOro14) + (lnGra16 * pnPreOro16) + (lnGra18 * pnPreOro18) + (lnGra21 * pnPreOro21)
                lnPreMerc = Round(lnPreMerc, 2)
                
                lnGramos = lnGra14 + lnGra16 + lnGra18 + lnGra21
                lsNombre = ofcar.PstaNombre(!cNomCliente)
                
                'vdiasAtraso = DateDiff("d", !dVenc, psFecProceso)
                vdiasAtraso = DateDiff("d", !dVenc, gdFecSis)
                
                Set loCalculos = New NCOMColPCalculos
                If vdiasAtraso <= 0 Then
                    vdiasAtraso = 0
                    lnIntVencido = 0
                    lnIntAdelantado = 0
                    lnIntMoratorio = 0
                  Else
                    ''*** PEAC 20090226 ***************************************************
                    lnIntAdelantado = loCalculos.nCalculaInteresAlVencimiento(!nSaldo, !nTasaIntVenc, 30)
                    lnIntAdelantado = Round(lnIntAdelantado, 2)
                    
                    lnIntMoratorio = loCalculos.nCalculaInteresMoratorio(!nSaldo, !nTasaIntMora, vdiasAtraso)
                    lnIntMoratorio = Round(lnIntMoratorio, 2)
                    
                    'lnIntVencido = loCalculos.nCalculaInteresMoratorio(!nSaldo, !nTasaIntVenc, vdiasAtraso)
                    lnIntVencido = loCalculos.nCalculaInteresMoratorio(!nSaldo, !nTasaIntVenc, vdiasAtraso, lnIntAdelantado) '*** PEAC 20190725
                    lnIntVencido = Round(lnIntVencido, 2)
                    
                    ''*********************************************************************
                End If
                Set loCalculos = Nothing
                
                'Set loCalculos = New NCOMColPCalculos
                
                '*** PEAC 20190725
                lnValorCosCustoDif = 0
                If !nSaldo = 0 And !nPrdEstado = 2102 Then
                    lnValorCosCustoDif = !monto_custo
                End If
                
                ''*** PEAC 20090227 *******************************
                'lnDeuda = loCalcula.nCalculaDeudaPignoraticio(!nsaldo, !dVenc, !nTasacion, !ntasaintvenc, pnTasaCustodiaVencida, pnTasaImpuesto, !nPrdEstado, pnTasaPreparacionRemate, gdFecSis)
                'lnDeuda = !nSaldo + Round(lnIntMoratorio, 2) + Round(lnIntVencido, 2) + Round(lnIntAdelantado, 2) + !nNotificacion
                lnDeuda = !nSaldo + Round(lnIntMoratorio, 2) + Round(lnIntVencido, 2) + Round(lnIntAdelantado, 2) + !nNotificacion + lnValorCosCustoDif '*** PEAC 20190725
                lnDeuda = Round(lnDeuda, 2)
                
                ''**************************************************
                    
                    ' Calcula el Valor Adjudicacion
                    lnValRegis = loCalcula.nCalculaValorAdjudicacion(!nSaldo, lnPreMerc)
                    lnValRegis = Round(lnValRegis, 2)
              

                lnLineas = lnLineas + 1
                    
                vSumGra14 = vSumGra14 + lnGra14
                vSumGra16 = vSumGra16 + lnGra16
                vSumGra18 = vSumGra18 + lnGra18
                vSumGra21 = vSumGra21 + lnGra21
                vSumGramos = vSumGramos + lnGramos
                vSumTasaci = vSumTasaci + !nTasacion
                vSumSalCap = vSumSalCap + !nSaldo
                vSumDeuda = vSumDeuda + lnDeuda
                vSumPreMerc = vSumPreMerc + lnPreMerc
                vSumValRegis = vSumValRegis + lnValRegis
                
                ''** PEAC 20090227
                'lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 4, 0) & Space(1) & !cctacod & Space(1) & Format(!dVenc, "dd/mm/yyyy") _
                    & ofun.ImpreFormat(ofun.ImpreCarEsp(lsNombre), 20, 1) _
                    & ofun.ImpreFormat(lnGra14, 6) & ofun.ImpreFormat(lnGra16, 6) & ofun.ImpreFormat(lnGra16, 6) & ofun.ImpreFormat(lnGra18, 6) _
                    & ofun.ImpreFormat(lnGramos, 6) & ofun.ImpreFormat(!nTasacion, 10) & ofun.ImpreFormat(!nsaldo, 10) _
                    & ofun.ImpreFormat(lnDeuda, 10) _
                    & ofun.ImpreFormat(lnValRegis, 10) _
                    & oFunI.gPrnSaltoLinea
                
                ''** PEAC 20090227
                '*** Comentado por JGPA20180523
                'lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 4, 0) & Space(1) & !cCtaCod & Space(1) & Format(!dVenc, "dd/mm/yyyy") _
                    & ofun.ImpreFormat(ofun.ImpreCarEsp(lsNombre), 20, 1) _
                    & ofun.ImpreFormat(lnGra14, 6) & ofun.ImpreFormat(lnGra16, 6) & ofun.ImpreFormat(lnGra16, 6) & ofun.ImpreFormat(lnGra18, 6) _
                    & ofun.ImpreFormat(lnGramos, 6) & ofun.ImpreFormat(!nTasacion, 10) & ofun.ImpreFormat(!nSaldo, 10) _
                    & ofun.ImpreFormat(lnDeuda, 10) _
                    & oFunI.gPrnSaltoLinea
                
                '*** JGPA20180523 según ACTA 028-2018
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 4, 0) & Space(1) & !cCtaCod & Space(1) & Format(!dVenc, "dd/mm/yyyy") _
                    & ofun.ImpreFormat(ofun.ImpreCarEsp(lsNombre), 20, 1) _
                    & ofun.ImpreFormat(lnGra14, 6) & ofun.ImpreFormat(lnGra16, 6) & ofun.ImpreFormat(lnGra18, 6) & ofun.ImpreFormat(lnGra21, 6) _
                    & ofun.ImpreFormat(lnGramos, 6) & ofun.ImpreFormat(!nTasacion, 10) & ofun.ImpreFormat(!nSaldo, 10) _
                    & ofun.ImpreFormat(lnDeuda, 10) _
                    & oFunI.gPrnSaltoLinea
                '***End JGPA20180523
                    
                '& ofun.ImpreFormat(lnPreMerc, 10)
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                        
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nImprimeAdjudicaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "PLANILLA PARA ADJUDICACION NRO " & psNroProceso, "Del " & Format(psFecProceso, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 155, "", 1)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        
        ''** PEAC 20090227
        'lsCadImp = lsCadImp & Space(50) & "RESUMEN" & ofun.ImpreFormat(vSumGramos, 40) & ofun.ImpreFormat(vSumTasaci, 10) _
                & ofun.ImpreFormat(vSumSalCap, 10) & ofun.ImpreFormat(vSumDeuda, 10) _
                & ofun.ImpreFormat(vSumValRegis, 10) _
                & "" _
                & oFunI.gPrnSaltoLinea
                    
        ''** PEAC 20090227
        lsCadImp = lsCadImp & Space(50) & "RESUMEN" & ofun.ImpreFormat(vSumGramos, 40) & ofun.ImpreFormat(vSumTasaci, 10) _
                & ofun.ImpreFormat(vSumSalCap, 10) & ofun.ImpreFormat(vSumDeuda, 10) _
                & "" _
                & oFunI.gPrnSaltoLinea
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimePlanillaParaAdjudicacion = lsCadBuffer
End Function


Public Function nImprimePlanillaParaAdjudicacionConSIAF(ByVal psFecProceso As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal pnNroRemParaAdjud As Integer, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroProceso As String, ByVal gdFecSis As Date, _
        ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim loCalculos As NCOMColPCalculos

Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer

Dim lnGra14 As Double, lnGra16 As Double, lnGra18 As Double, lnGra21 As Double
Dim lnGramos As Currency, lnTasaci As Currency, lnSalCap As Currency, lnDeuda As Currency
Dim lnPreMerc As Currency, lnValRegis As Currency

Dim lsNombre As String * 28

Dim vSumGra14 As Currency, vSumGra16 As Currency, vSumGra18 As Currency, vSumGra21 As Currency
Dim vSumGramos As Currency, vSumTasaci As Currency, vSumSalCap As Currency, vSumDeuda As Currency
Dim vSumPreMerc As Currency, vSumValRegis As Currency
Dim loCalcula As NCOMColPCalculos
Set loCalcula = New NCOMColPCalculos
    
Dim vdiasAtraso As Integer

Dim ofcar As New COMFunciones.FCOMCadenas
Dim ofun As New COMFunciones.FCOMImpresion

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

    lsSQL = " SELECT CODESIAF=ISNULL(RE.CCTACODANT,'                   ') ,P.cCtaCod, P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred,  " _
          & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , " _
          & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , " _
          & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , " _
          & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , " _
          & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  " _
          & "               WHERE cLineaCred = C.cLineaCred  " _
          & "               AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ),    " _
          & " cNomCliente = (SELECT MAX(cPersNombre) FROM Persona Per INNER JOIN ProductoPersona PPer " _
          & "               ON Per.cPersCod = PPer.cPersCod WHERE PPer.cCtaCod = P.cCtaCod )  " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " LEFT JOIN RELCUENTAS RE ON RE.CCTACOD=P.CCTACOD " _
          & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod  " _
          & " WHERE P.nPrdEstado in (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ") " _
          & " AND Substring(P.cCtaCod,4,2) ='" & psAgencia & "' AND nNroRemate = " & pnNroRemParaAdjud _
          & " GROUP BY P.cCtaCod,ISNULL(RE.CCTACODANT,'                   ') , P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda,C.cLineaCred   " _
          & " ORDER BY P.cCtaCod "

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Then
        psmensaje = " No existen Créditos Pignoraticios para Aviso de Vencimiento "
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1
        
        'CABECERA
        lsCadImp = lsCadImp & nImprimeAdjudicaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "PLANILLA PARA ADJUDICACION NRO " & psNroProceso, "Del " & Format(psFecProceso, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 155, "", 9)

        lnIndice = 1:        lnLineas = 7
        
        vSumGra14 = 0:  vSumGra16 = 0:  vSumGra18 = 0:  vSumGra21 = 0
        vSumGramos = 0: vSumTasaci = 0: vSumSalCap = 0: vSumDeuda = 0
        vSumPreMerc = 0: vSumValRegis = 0
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                
                lnGra14 = !nK14:  lnGra14 = Round(lnGra14, 2)
                lnGra16 = !nK16:  lnGra16 = Round(lnGra16, 2)
                lnGra18 = !nK18:  lnGra18 = Round(lnGra18, 2)
                lnGra21 = !nK21:  lnGra21 = Round(lnGra21, 2)
                
                lnPreMerc = (lnGra14 * pnPreOro14) + (lnGra16 * pnPreOro16) + (lnGra18 * pnPreOro18) + (lnGra21 * pnPreOro21)
                lnPreMerc = Round(lnPreMerc, 2)
                
                lnGramos = lnGra14 + lnGra16 + lnGra18 + lnGra21
                lsNombre = ofcar.PstaNombre(!cNomCliente)
        
                'Set loCalculos = New NCOMColPCalculos
                    lnDeuda = loCalcula.nCalculaDeudaPignoraticio(!nSaldo, !dVenc, !nTasacion, !nTasaIntVenc, pnTasaCustodiaVencida, pnTasaImpuesto, !nPrdEstado, pnTasaPreparacionRemate, gdFecSis)
                    
                    ' Calcula el Valor Adjudicacion
                    lnValRegis = loCalcula.nCalculaValorAdjudicacion(!nSaldo, lnPreMerc)
                    lnValRegis = Round(lnValRegis, 2)
              

                lnLineas = lnLineas + 1
                    
                vSumGra14 = vSumGra14 + lnGra14
                vSumGra16 = vSumGra16 + lnGra16
                vSumGra18 = vSumGra18 + lnGra18
                vSumGra21 = vSumGra21 + lnGra21
                vSumGramos = vSumGramos + lnGramos
                vSumTasaci = vSumTasaci + !nTasacion
                vSumSalCap = vSumSalCap + !nSaldo
                vSumDeuda = vSumDeuda + lnDeuda
                vSumPreMerc = vSumPreMerc + lnPreMerc
                vSumValRegis = vSumValRegis + lnValRegis
                
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 4, 0) & Space(1) & !codesiaf & Space(1) & !cCtaCod & Space(1) & Format(!dVenc, "dd/mm/yyyy") _
                    & ofun.ImpreFormat(ofun.ImpreCarEsp(lsNombre), 20, 1) _
                    & ofun.ImpreFormat(lnGra14, 6) & ofun.ImpreFormat(lnGra16, 6) & ofun.ImpreFormat(lnGra18, 6) & ofun.ImpreFormat(lnGra21, 6) _
                    & ofun.ImpreFormat(lnGramos, 6) & ofun.ImpreFormat(!nTasacion, 10) & ofun.ImpreFormat(!nSaldo, 10) _
                    & ofun.ImpreFormat(lnValRegis, 10) _
                    & oFunI.gPrnSaltoLinea
                    
                    '& ofun.ImpreFormat(lnDeuda, 10) _
                '& ofun.ImpreFormat(lnPreMerc, 10)
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                        
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nImprimeAdjudicaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "PLANILLA PARA ADJUDICACION NRO " & psNroProceso, "Del " & Format(psFecProceso, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 155, "", 9)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(50) & "RESUMEN" & ofun.ImpreFormat(vSumGramos, 40) & ofun.ImpreFormat(vSumTasaci, 10) _
                & ofun.ImpreFormat(vSumSalCap, 10) & ofun.ImpreFormat(vSumDeuda, 10) _
                & ofun.ImpreFormat(vSumValRegis, 10) _
                & "" _
                & oFunI.gPrnSaltoLinea
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimePlanillaParaAdjudicacionConSIAF = lsCadBuffer
    Set oFunI = Nothing
End Function

Public Function nAdjudicaLotes(ByVal psMovNro As String, ByVal psFechaHora As String, _
        ByVal psTipoAdju As String, ByVal psNroAdju As String, _
        ByVal pnPreOroInternacional As Currency, ByVal pnTipoCambio As Double, _
        ByVal psFecAdjudica As String, ByVal pnPrecioOro14 As Double, ByVal pnPrecioOro16 As Double, _
        ByVal pnPrecioOro18 As Double, ByVal pnPrecioOro21 As Double, ByVal psFecParaAdjudica As String, _
        ByVal psProcesoCadaAgencia As String, ByVal pnNroRemParaAdjud As Integer, ByRef psmensaje As String, Optional pdFecCorte As Date, Optional pdFecSis As Date) As Boolean

Dim lsSQL As String
Dim loDataAdj As COMDColocPig.DCOMColPFunciones
Dim lrDataAdj As New ADODB.Recordset
Dim lnMovNro As Long
Dim loCalcula As NCOMColPCalculos
Dim loRegPig As COMDColocPig.DCOMColPActualizaBD
Dim lsOpeCod As String
Dim lsNextAdjud As String
Dim pdFechaLey As String

'**DAOR 20080131 *****************************************
Dim loNroPro As COMDColocPig.DCOMColPFunciones
Dim lrs As ADODB.Recordset
'*********************************************************

''*** PEAC 20090226 ***********************
Dim loCalculos As NCOMColPCalculos
Dim vdiasAtraso As Integer
Dim lnIntVencido As Double
Dim lnIntMoratorio As Double
Dim lnIntAdelantado As Double, vCostoCustodiaMoratorio As Currency
Dim lnDeuda As Currency
''*****************************************
'*** PEAC 20190326
Dim lnTasacionCustodia As Double
Dim lnCustDofDeuda As Double

''*** GIPO 20171013 ***********************

lsOpeCod = geColPAdjudica

Dim lnPreMercado As Currency, lnPreDeuda As Currency, lnPreAdj As Currency

pdFechaLey = "01/06/2006"

lsSQL = " exec stp_sel_AdjudicaLotes '" & Right(psProcesoCadaAgencia, 2) & "','" & Format(pdFechaLey, "yyyymmdd") & "','" & Format(pdFecCorte, "yyyymmdd") & "','" & psNroAdju & "'"

Set loDataAdj = New COMDColocPig.DCOMColPFunciones
    Set lrDataAdj = loDataAdj.dObtieneRecordSet(lsSQL)
Set loDataAdj = Nothing
        
If lrDataAdj Is Nothing Then
    psmensaje = "No hay creditos para Adjudicar"
    
Else
    lnPreMercado = 0: lnPreDeuda = 0: lnPreAdj = 0
    
    
    Set loRegPig = New COMDColocPig.DCOMColPActualizaBD
        
        'Inicializa Transaccion
        loRegPig.dBeginTrans
        mbTrans = True
        
        '** Inserta Mov
        Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Adjudicacion Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
        
        ' Obtiene nMovNro
        lnMovNro = loRegPig.dGetnMovNro(psMovNro)
        
        
         '**Registro de Montos de Devolución de Sobrantes Adjud. ERS048-2017 12-10-2017
'        nMovNroRegSobrantes = loRegPig.InsertMovSobrantes(psMovNro, cOpeCodRegSobrantes, cDescripRegSobrantes, gMovEstContabMovContable, gMovFlagVigente, False)
        
        With lrDataAdj
            
            Do While Not .EOF
                'Precio Base
                lnPreMercado = !nK14 * pnPrecioOro14 + !nK16 * pnPrecioOro16 + !nK18 * pnPrecioOro18 + !nK21 * pnPrecioOro21
                lnPreMercado = Round(lnPreMercado, 2)

                'Calcula Deuda
                'Set loCalcula = New NCOMColPCalculos
                '    lnPreDeuda = loCalcula.nCalculaDeudaPignoraticio(!nSaldo, !dFecVenc, !nTasacion, !nTasaIntVenc, _
                '        !nCostCusto, pnTasaImpuesto, !nPrdEstado, pnTasaPreparacionRemate, Format(psFecRemate, "mm/dd/yyyy"))
                '
                '    lnPreDeuda = Round(lnPreDeuda, 2)
                '    Set loCalcula = Nothing

'*********************************************************

                vdiasAtraso = DateDiff("d", !dVenc, pdFecSis)
                'vdiasAtraso = DateDiff("d", !dVenc, CDate(psFecAdjudica))
                Set loCalculos = New NCOMColPCalculos
                If vdiasAtraso <= 0 Then
                    vdiasAtraso = 0
                    lnIntVencido = 0
                    lnIntAdelantado = 0
                    lnIntMoratorio = 0
                    vCostoCustodiaMoratorio = 0
                  Else
                                      
                    '*** PEAC 20190715
                    lnIntAdelantado = loCalculos.nCalculaInteresAlVencimiento(!nSaldo, !nTasaIntVenc, 30)
                    lnIntAdelantado = Round(lnIntAdelantado, 2)
                  
                    lnIntMoratorio = loCalculos.nCalculaInteresMoratorio(!nSaldo, !nTasaIntMora, vdiasAtraso)
                    lnIntMoratorio = Round(lnIntMoratorio, 2)
                    
                    lnIntVencido = loCalculos.nCalculaInteresMoratorio(!nSaldo, !nTasaIntVenc, vdiasAtraso, lnIntAdelantado)
                    lnIntVencido = Round(lnIntVencido, 2)
                    '*** PEAC 20190715 - SE AGREGO "lnIntAdelantado"

'                    lnIntAdelantado = loCalculos.nCalculaInteresAlVencimiento(!nSaldo, !nTasaIntVenc, 30)
'                    lnIntAdelantado = Round(lnIntAdelantado, 2)

                End If
                Set loCalculos = Nothing

                '*** PEAC 20190719
                lnCustDofDeuda = IIf(!nSaldo = 0 And !nPrdEstado = 2102, !nTasacion, 0)

                ''*** PEAC 20090227 *******************************
                'lnDeuda = loCalcula.nCalculaDeudaPignoraticio(!nsaldo, !dVenc, !nTasacion, !ntasaintvenc, pnTasaCustodiaVencida, pnTasaImpuesto, !nPrdEstado, pnTasaPreparacionRemate, gdFecSis)
                lnDeuda = !nSaldo + Round(lnIntMoratorio, 2) + Round(lnIntVencido, 2) + Round(lnIntAdelantado, 2) + !nNotificacion + lnCustDofDeuda '*** PEAC - Se agrego lnCustDofDeuda
                lnDeuda = Round(lnDeuda, 2)

'*********************************************************
                      
                        

                'NOTA. Se adjudica al menor entre SALDO DE CAPITAL y valor de mercado
                lnPreAdj = IIf(!nSaldo > lnPreMercado, lnPreMercado, !nSaldo)
                lnPreAdj = Round(lnPreAdj, 2)
                
                '** Actualiza Producto
                Call loRegPig.dUpdateProducto(!cCtaCod, , , gColPEstAdjud, psFechaHora, -2, False)
                
                '** Actualiza Colocaciones
                Call loRegPig.dUpdateColocaciones(!cCtaCod, , , , , psMovNro, , False)
                
                '** Actualiza ColocPignoraticio
                'Call loRegPig.dUpdateColocPignoraticio(!cCtaCod, , , , , , , , , , , False)
                
                '** Insert ColocEstado
                '*** PEAC 20090302
                'Call loRegPig.dInsertColocacEstado(!cctacod, psFechaHora, gColPEstAdjud, 1, lnPreAdj, "Adjudicacion de Joyas", _
                     gColocCalendCodFFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
                Call loRegPig.dInsertColocacEstado(!cCtaCod, psFechaHora, gColPEstAdjud, 1, lnDeuda, "Adjudicacion de Joyas", _
                     gColocCalendCodFFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
                                
                '** Inserta RGDet
                '*** PEAC 20090302
                'Call loRegPig.dInsertColocPigRGDet("A", psNroAdju, !cctacod, gColPRecGarAdjudicado, lnPreAdj, 0, lnPreMercado, lnPreAdj, 0, psProcesoCadaAgencia, False)
                Call loRegPig.dInsertColocPigRGDet("A", psNroAdju, !cCtaCod, gColPRecGarAdjudicado, lnDeuda, 0, lnPreMercado, lnDeuda, 0, psProcesoCadaAgencia, False)
                
                                                
                '** Inserta MovCol
                '*** PEAC 20090302
                'Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, !cctacod, 1, lnPreAdj, 0, "", 0, gColPEstAdjud, !nsaldo, False)
                Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, !cCtaCod, 1, lnDeuda, 0, "", 0, gColPEstAdjud, !nSaldo, False)
            
                '**************************************** PEAC 20090302
                            
                Call loRegPig.dUpdateColocPigNotifi(!cCtaCod, , 1, psMovNro, False)

                '** Inserta MovColDet -  Interes Compensatorio (al vencimiento)
                If lnIntAdelantado > 0 Then
                    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, !cCtaCod, 1, gColPConceptoCodInteresCompensatorio, 1, lnIntAdelantado, False)
                End If

                '** Inserta MovColDet -  Interes Vencido
                If lnIntVencido > 0 Then
                    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, !cCtaCod, 1, gColPConceptoCodInteresVencido, 1, lnIntVencido, False)
                End If

                 '** Inserta MovColDet -  Interes Moratorio
                If lnIntMoratorio > 0 Then
                    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, !cCtaCod, 1, gColPConceptoCodInteresMoratorio, 1, lnIntMoratorio, False)
                End If

                '*** inserta en MovColDet - costo de notificacion
                If !nNotificacion > 0 Then
                    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, !cCtaCod, 1, "2103", 1, !nNotificacion, False)
                End If

                '*** inserta en MovColDet - 2218-Monto por cobrar custodia dif Y 2219-Adjudica Joya Valor Tasacion
                If lnCustDofDeuda > 0 Then '*** PEAC 20190719
                    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, !cCtaCod, 1, "2218", 1, lnCustDofDeuda, False)
                    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, !cCtaCod, 1, "2219", 1, lnCustDofDeuda, False)
                End If

                '****************************************
            
                '** Inserta MovColDet -  Capital
                Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, !cCtaCod, 1, gColPConceptoCodCapital, 1, !nSaldo, False)
            
                '** Inserta MovColDet -  Valor Tasacion
                Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, !cCtaCod, 1, gColPConceptoCodTasacion, 1, !nTasacion, False)
                
                ' Gramos de Oro
                If !nK14 > 0 Then
                    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, !cCtaCod, 1, gColPConceptoCodOro14, 1, !nK14, False)
                End If
                If !nK16 > 0 Then
                    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, !cCtaCod, 1, gColPConceptoCodOro16, 1, !nK16, False)
                End If
                If !nK18 > 0 Then
                    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, !cCtaCod, 1, gColPConceptoCodOro18, 1, !nK18, False)
                End If
                If !nK21 > 0 Then
                    Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, !cCtaCod, 1, gColPConceptoCodOro21, 1, !nK21, False)
                End If
                
                'Registro de los montos a devolver ERS048-2017 GIPO
                Call loRegPig.RegistrarSobrantesPorCuenta(!cCtaCod, lnMovNro, psNroAdju, !nTasacion - lnDeuda) 'sobrante a devolver
                
                .MoveNext
            Loop
            .Close
        End With
        'Actualiza el Proceso Adjudicacion
        Call loRegPig.dUpdateColocPigRecGar("A", psNroAdju, psProcesoCadaAgencia, , gColPRecGarEstTerminado, psFecAdjudica, False)
        
        '**DAOR 20080131, Genera el sgte Nro de Adjudicacion ****************************
        Set loNroPro = New COMDColocPig.DCOMColPFunciones
            lsSQL = "SELECT max(cNroProceso) nNextProceso FROM ColocPigRecGar RG  WHERE cTpoProceso ='A'"
            Set lrs = loNroPro.dObtieneRecordSet(lsSQL)
            If Not (lrs.EOF And lrs.BOF) Then
                lsNextAdjud = String(4 - Len(Trim(Str(Val(lrs!nNextProceso) + 1))), "0") & Trim(Str(Val(lrs!nNextProceso) + 1))
            End If
        Set loNroPro = Nothing
        '*********************************************************************************

        'Insert el Sgte proceso de Adjudicacion
        'lsNextAdjud = String(4 - Len(Trim(Str(Val(psNroAdju) + 1))), "0") & Trim(Str(Val(psNroAdju) + 1))
        
        Call loRegPig.dInsertColocPigRecGar("A", lsNextAdjud, Format(DateAdd("m", 1, psFecAdjudica) & " 08:00", "mm/dd/yyyy hh:mm"), gColPRecGarEstNoIniciado, psProcesoCadaAgencia, , False)
               
        '*******************************************************************
        
        loRegPig.dCommitTrans
        'Generar e Insertar los códigos de las Joyas adjudicadas - GIPO 20190102
        Call loRegPig.InsertCodigosJoyasAdjud(psNroAdju, lnMovNro)        
    Set loRegPig = Nothing
End If
Set lrDataAdj = Nothing
psmensaje = "Se han registrado los montos de devoluciones pendientes. Para efectuar las devoluciones " & _
            " a Cuenta diríjase al SICMACM Web e ingrese a la ruta: " & vbNewLine & "Operaciones > Devolución de Sobrantes Adjud.> 123400 - Devoluciones a Cuenta"
Exit Function

ErrorModPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error nRegistraContratoPignoraticio", "Error en Funcion de Registro de Contrato "

End Function

'Imprime Listado de Adjudicados
Public Function nImprimeListadoAdjudicados(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal psNroProcesoAdjud As String, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByRef psmensaje As String, _
        Optional ByVal psImpresora As Impresoras = gEPSON) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim loCalculos As NCOMColPCalculos

Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double
Dim lsProcesoListar As String

'*** PEAC 20080714
'lsProcesoListar = IIf(psNroProcesoAdjud = "0000", "", "AND CRGDet.cNroProceso ='" & psNroProcesoAdjud & "'")
lsProcesoListar = IIf(psNroProcesoAdjud = "0000", "", psNroProcesoAdjud)
'*** END PEAC 20080714

Dim vSumGra14 As Currency, vSumGra16 As Currency, vSumGra18 As Currency, vSumGra21 As Currency
Dim vSumGramos As Currency, vSumTasaci As Currency, vSumSalCap As Currency, vSumDeuda As Currency
Dim vSumPreMerc As Currency, vSumValRegis As Currency
Dim lsNombre As String * 28

Dim ofcar As New COMFunciones.FCOMCadenas
Dim ofun As New COMFunciones.FCOMImpresion

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


    'lsSQL = " SELECT CODESIAF=ISNULL(R.CCTACODANT,'                  '),P.cCtaCod, P.nPrdEstado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred, CRGDet.nDeuda, CRGDet.nAdjValRegistro,  " _
          & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , " _
          & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , " _
          & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , " _
          & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , " _
          & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  " _
          & "               WHERE cLineaCred = C.cLineaCred  " _
          & "               AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ), " _
          & " cNomCliente = (SELECT MAX(cPersNombre) FROM Persona Per INNER JOIN ProductoPersona PPer " _
          & "               ON Per.cPersCod = PPer.cPersCod WHERE PPer.cCtaCod = P.cCtaCod )  " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " LEFT JOIN RELCUENTAS R ON R.CCTACOD=P.CCTACOD " _
          & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " INNER JOIN ColocPigRGDet CRGDet On P.cCtaCod = CRGDet.cCtaCod " _
          & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod  " _
          & " WHERE CRGDet.cTpoProceso ='A' AND Substring(P.cCtaCod,4,2) = '" & psAgencia & "' " _
          & lsProcesoListar _
          & " And P.nPrdEstado= " & gColPEstAdjud _
          & " GROUP BY P.cCtaCod,R.CCTACODANT,P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred, CRGDet.nDeuda, CRGDet.nAdjValRegistro " _
          & " ORDER BY P.cCtaCod "

    'Cambio Miguel And P.nPrdEstado=gColPEstAdjud  (2108)

    '*** PEAC 20080714
    If Len(lsProcesoListar) = 0 Then
        lsSQL = " exec stp_sel_ImprimeListadoAdjudicados '" & psAgencia & "'"
    Else
        lsSQL = " exec stp_sel_ImprimeListadoAdjudicados '" & psAgencia & "','" & lsProcesoListar & "'"
    End If

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing


    If lrDataRep Is Nothing Then
        psmensaje = " No existen Créditos Pignoraticios Adjudicados "
        'MuestraImpresion = False

        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        If lsProcesoListar = "0000" Then
            lsCadImp = lsCadImp & nImprimeAdjudicaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "LISTADO DE CONTRATOS ADJUDICADOS  CONSOLIDADO " & " - Agencia " & psAgencia, "", lnPage, 144, "", 2)
        Else
            lsCadImp = lsCadImp & nImprimeAdjudicaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "LISTADO DE CONTRATOS PARA ADJUDICACION NRO " & psNroProcesoAdjud, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 144, "", 2)
        End If

        lnIndice = 1:        lnLineas = 7
        

        vSumGra14 = 0:  vSumGra16 = 0:  vSumGra18 = 0:  vSumGra21 = 0
        vSumGramos = 0: vSumTasaci = 0: vSumSalCap = 0: vSumDeuda = 0
        vSumPreMerc = 0: vSumValRegis = 0
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                                
                v14 = !nK14: v14 = Round(v14, 2)
                v16 = !nK16: v16 = Round(v16, 2)
                v18 = !nK18: v18 = Round(v18, 2)
                v21 = !nK21: v21 = Round(v21, 2)
                vCuenta = 0
                'lnPreBase = (v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)
                lnGramos = v14 + v16 + v18 + v21
                lsNombre = ofcar.PstaNombre(!cNomCliente)
                
                'lnPreBase = Round(lnPreBase, 2)
                
                lnDeuda = !nDeuda

                lnLineas = lnLineas + 1
                                        
                vSumGra14 = vSumGra14 + v14
                vSumGra16 = vSumGra16 + v16
                vSumGra18 = vSumGra18 + v18
                vSumGra21 = vSumGra21 + v21
                vSumGramos = vSumGramos + lnGramos
                vSumTasaci = vSumTasaci + !nTasacion
                vSumSalCap = vSumSalCap + !nSaldo
                vSumDeuda = vSumDeuda + lnDeuda
                'vSumPreMerc = vSumPreMerc + lnPreMerc
                vSumValRegis = vSumValRegis + !nAdjValRegistro
                                        
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !codesiaf & Space(1) & !cCtaCod & Space(1) _
                    & ofun.ImpreFormat(ofun.ImpreCarEsp(lsNombre), 20, 1) _
                    & ofun.ImpreFormat(v14, 6) & ofun.ImpreFormat(v16, 6) & ofun.ImpreFormat(v18, 6) & ofun.ImpreFormat(v21, 6) _
                    & ofun.ImpreFormat(lnGramos, 6) & ofun.ImpreFormat(!nTasacion, 8) & ofun.ImpreFormat(!nSaldo, 8) _
                    & ofun.ImpreFormat(!nAdjValRegistro, 8) _
                    & oFunI.gPrnSaltoLinea
                   ' & ofun.ImpreFormat(lnDeuda, 10) _

                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                        
                'lnLineas = lnLineas + 1
                'lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
                
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    If lsProcesoListar = "0000" Then
                        lsCadImp = lsCadImp & nImprimeAdjudicaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "LISTADO DE CONTRATOS ADJUDICADOS  CONSOLIDADO " & " - Agencia " & psAgencia, " ", lnPage, 120, "", 2)
                    Else
                        lsCadImp = lsCadImp & nImprimeAdjudicaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "LISTADO DE CONTRATOS PARA ADJUDICACION NRO " & psNroProcesoAdjud, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 120, "", 2)
                    End If
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & String(144, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(37) & "RESUMEN " & ofun.ImpreFormat(vSumGramos, 40) & ofun.ImpreFormat(vSumTasaci, 10) _
                & ofun.ImpreFormat(vSumSalCap, 10) & ofun.ImpreFormat(vSumDeuda, 10) _
                & ofun.ImpreFormat(vSumValRegis, 10) _
                & "" _
                & oFunI.gPrnSaltoLinea
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimeListadoAdjudicados = lsCadBuffer
    Set oFunI = Nothing
End Function

'Imprime Listado de Adjudicados
Public Function nImprimeListadoAdjudicadosConSIAF(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal psNroProcesoAdjud As String, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByRef psmensaje As String, _
        Optional ByVal psImpresora As Impresoras = gEPSON) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim loCalculos As NCOMColPCalculos

Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double
Dim lsProcesoListar As String
lsProcesoListar = IIf(psNroProcesoAdjud = "0000", "", "AND CRGDet.cNroProceso ='" & psNroProcesoAdjud & "'")


Dim vSumGra14 As Currency, vSumGra16 As Currency, vSumGra18 As Currency, vSumGra21 As Currency
Dim vSumGramos As Currency, vSumTasaci As Currency, vSumSalCap As Currency, vSumDeuda As Currency
Dim vSumPreMerc As Currency, vSumValRegis As Currency
Dim lsNombre As String * 28

Dim ofcar As New COMFunciones.FCOMCadenas
Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


    lsSQL = " SELECT CODESIAF=ISNULL(RE.CCTACODANT,'                  '),P.cCtaCod, P.nPrdEstado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred, CRGDet.nDeuda, CRGDet.nAdjValRegistro,  " _
          & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , " _
          & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , " _
          & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , " _
          & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , " _
          & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  " _
          & "               WHERE cLineaCred = C.cLineaCred  " _
          & "               AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ), " _
          & " cNomCliente = (SELECT MAX(cPersNombre) FROM Persona Per INNER JOIN ProductoPersona PPer " _
          & "               ON Per.cPersCod = PPer.cPersCod WHERE PPer.cCtaCod = P.cCtaCod )  " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " LEFT JOIN RELCUENTAS RE ON RE.CCTACOD=P.CCTACOD " _
          & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " INNER JOIN ColocPigRGDet CRGDet On P.cCtaCod = CRGDet.cCtaCod " _
          & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod  " _
          & " WHERE CRGDet.cTpoProceso ='A' AND Substring(P.cCtaCod,4,2) = '" & psAgencia & "' " _
          & lsProcesoListar _
          & " And P.nPrdEstado= " & gColPEstAdjud _
          & " GROUP BY P.cCtaCod,ISNULL(RE.CCTACODANT,'                  ') ,P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred, CRGDet.nDeuda, CRGDet.nAdjValRegistro " _
          & " ORDER BY P.cCtaCod "

    'Cambio Miguel And P.nPrdEstado=gColPEstAdjud  (2108)

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing


    If lrDataRep Is Nothing Then
         psmensaje = " No existen Créditos Pignoraticios Adjudicados "
        'MuestraImpresion = False

        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        If lsProcesoListar = "0000" Then
            lsCadImp = lsCadImp & nImprimeAdjudicaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "LISTADO DE CONTRATOS ADJUDICADOS  CONSOLIDADO " & " - Agencia " & psAgencia, "", lnPage, 161, "", 10)
        Else
            lsCadImp = lsCadImp & nImprimeAdjudicaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "LISTADO DE CONTRATOS PARA ADJUDICACION NRO " & psNroProcesoAdjud, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 161, "", 10)
        End If

        lnIndice = 1:        lnLineas = 7
        

        vSumGra14 = 0:  vSumGra16 = 0:  vSumGra18 = 0:  vSumGra21 = 0
        vSumGramos = 0: vSumTasaci = 0: vSumSalCap = 0: vSumDeuda = 0
        vSumPreMerc = 0: vSumValRegis = 0
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                
                v14 = !nK14: v14 = Round(v14, 2)
                v16 = !nK16: v16 = Round(v16, 2)
                v18 = !nK18: v18 = Round(v18, 2)
                v21 = !nK21: v21 = Round(v21, 2)
                vCuenta = 0
                'lnPreBase = (v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)
                lnGramos = v14 + v16 + v18 + v21
                lsNombre = ofcar.PstaNombre(!cNomCliente)
                
                'lnPreBase = Round(lnPreBase, 2)
                
                lnDeuda = !nDeuda

                lnLineas = lnLineas + 1
                                        
                vSumGra14 = vSumGra14 + v14
                vSumGra16 = vSumGra16 + v16
                vSumGra18 = vSumGra18 + v18
                vSumGra21 = vSumGra21 + v21
                vSumGramos = vSumGramos + lnGramos
                vSumTasaci = vSumTasaci + !nTasacion
                vSumSalCap = vSumSalCap + !nSaldo
                vSumDeuda = vSumDeuda + lnDeuda
                'vSumPreMerc = vSumPreMerc + lnPreMerc
                vSumValRegis = vSumValRegis + !nAdjValRegistro
                                        
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !codesiaf & Space(1) & !cCtaCod & Space(1) _
                    & ofun.ImpreFormat(ofun.ImpreCarEsp(lsNombre), 20, 1) _
                    & ofun.ImpreFormat(v14, 6) & ofun.ImpreFormat(v16, 6) & ofun.ImpreFormat(v18, 6) & ofun.ImpreFormat(v21, 6) _
                    & ofun.ImpreFormat(lnGramos, 6) & ofun.ImpreFormat(!nTasacion, 10) & ofun.ImpreFormat(!nSaldo, 10) _
                    & ofun.ImpreFormat(lnDeuda, 10) _
                    & ofun.ImpreFormat(!nAdjValRegistro, 10) _
                    & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                        
                'lnLineas = lnLineas + 1
                'lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
                
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    If lsProcesoListar = "0000" Then
                        lsCadImp = lsCadImp & nImprimeAdjudicaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "LISTADO DE CONTRATOS ADJUDICADOS  CONSOLIDADO " & " - Agencia " & psAgencia, " ", lnPage, 161, "", 10)
                    Else
                        lsCadImp = lsCadImp & nImprimeAdjudicaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "LISTADO DE CONTRATOS PARA ADJUDICACION NRO " & psNroProcesoAdjud, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 161, "", 10)
                    End If
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & String(161, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(55) & "RESUMEN " & ofun.ImpreFormat(vSumGramos, 40) & ofun.ImpreFormat(vSumTasaci, 10) _
                & ofun.ImpreFormat(vSumSalCap, 10) & ofun.ImpreFormat(vSumDeuda, 10) _
                & ofun.ImpreFormat(vSumValRegis, 10) _
                & "" _
                & oFunI.gPrnSaltoLinea
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimeListadoAdjudicadosConSIAF = lsCadBuffer
    
End Function

'Imprime Aviso de Remate
Public Function nImprimeAvisoSubasta(ByVal psTextoCarta As String, ByVal psFecSubasta As String, _
       ByVal psAgencia As String, ByVal pnPrevioMax As Long, _
       ByVal psFechaSis As String, ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
Dim LALISTA() As String
Dim lnFilas As Long, lnItem As Long
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCartaModelo As String

Dim ofcar As New COMFunciones.FCOMCadenas
Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

'On Error GoTo dError

    'ldFecAviso = DateAdd("d", -pnDiasVctoParaRemate, psFecRemate)

    lsSQL = "SELECT Per.cPersNombre, Per.cPersDireccDomicilio, Per.cPersTelefono, Per.cPersDireccUbiGeo, " _
        & " Ubigeo.cUbiGeoDescripcion, C.cCtaCod , C.dVenc, P.nPrdEstado" _
        & " FROM Producto P JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod " _
        & " JOIN ProductoPersona PP ON c.cCtaCod = PP.cCtaCod " _
        & " JOIN Persona Per ON Per.cPersCod = PP.cPersCod " _
        & " LEFT JOIN UbicacionGeografica Ubigeo ON Per.cPersDireccUbiGeo = Ubigeo.cUbiGeoCod " _
        & " WHERE  " _
        & " P.nPrdEstado in (" & gColPEstAdjud & ") " _
        & " ORDER BY Per.cPersNombre "

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing


    If lrDataRep Is Nothing Then
        psmensaje = " No existen Créditos Pignoraticios para Aviso de Subasta "
        'MuestraImpresion = False

        Exit Function
    Else
        lnFilas = 0
        
        ReDim LALISTA(lrDataRep.RecordCount, 5)
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnFilas = lnFilas + 1
                If !cPersNombre = LALISTA(lnFilas - 1, 0) Then
                    LALISTA(lnFilas - 1, 3) = LALISTA(lnFilas - 1, 3) & " - " & !cCtaCod
                    LALISTA(lnFilas - 1, 4) = LALISTA(lnFilas - 1, 4) & " - " & Format(!dVenc, "dd/mm/yyyy")
                    lnFilas = lnFilas - 1
                Else
                    LALISTA(lnFilas, 0) = !cPersNombre
                    LALISTA(lnFilas, 1) = !cPersDireccDomicilio & " - " & !cUbiGeoDescripcion '  ClienteZona(!cCodZon) & " - " & ClienteCiudad(!cCodZon)
                    LALISTA(lnFilas, 2) = !cPersTelefono & ""
                    LALISTA(lnFilas, 3) = !cCtaCod
                    LALISTA(lnFilas, 4) = Format(!dVenc, "dd/mm/yyyy")
                End If
                .MoveNext
            Loop
        End With
        lrDataRep.Close
        Set lrDataRep = Nothing
    End If
    'Llena cartas
    For lnItem = 1 To lnFilas
        lsCartaModelo = psTextoCarta
        
        lsCartaModelo = Replace(lsCartaModelo, "<<FECHA>>", Format(psFechaSis, "dddd,d mmmm yyyy"), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<NOMBRE>>", ofcar.PstaNombre(LALISTA(lnItem, 0), False), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<CONTRATO>>", LALISTA(lnItem, 3), , 2, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<FECSUBASTA>>", psFecSubasta, , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<DIRECCION>>", LALISTA(lnItem, 1), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<TELEFONO>>", LALISTA(lnItem, 2), , 1, vbTextCompare)
        lsCadImp = lsCadImp & lsCartaModelo & oFunI.gPrnSaltoPagina
        'prgList.Value = vCont
    Next lnItem
    
    Erase LALISTA
    nImprimeAvisoSubasta = lsCadImp
    Set oFunI = Nothing
Exit Function
   
dError:
    Err.Raise Err.Number, "Obtiene Datos Imprime Listado Aviso Vencimiento  <<nImprimeAvisoVencimiento>>", Err.Description
End Function

Public Function nImprimeSubastaReporteCabec(ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, _
        ByVal psFechaSis As String, ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String, _
        ByVal pnNroRepor As Integer, Optional ByVal psImpresora As Impresoras = gEPSON) As String
        
Dim lsCadImp As String
Dim loImprimeCab As NCOMColPImpre
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
    
    Set loImprimeCab = New NCOMColPImpre
        lsCadImp = loImprimeCab.nImprimeCabeceraReportes(psNomCmac, psNomAgencia, psCodUser, psFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)
    Set loImprimeCab = Nothing
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & String(pnAnchoLinea, "-") & oFunI.gPrnSaltoLinea
    Select Case pnNroRepor
        Case 1 ' Planilla de Remate
            lsCadImp = lsCadImp & Space(1) & "Item    Contrato      Fec.Venc.     Cliente             Gramos    Tasacion   Prestamo  SaldoCap  Interes  Custodia Impuesto Prep.Rem   Deuda   P.Base Boveda " & oFunI.gPrnSaltoLinea
        Case 2 ' Listados para Remate
            lsCadImp = lsCadImp & Space(1) & "Item    Contrato    Piezas     Descripción                                      14Kl.    16Kl.    18Kl.    21Kl.  Prec.Base " & oFunI.gPrnSaltoLinea
    End Select
    lsCadImp = lsCadImp & String(pnAnchoLinea, "-") & oFunI.gPrnSaltoLinea
nImprimeSubastaReporteCabec = lsCadImp
End Function

'Imprime Planilla de Remate
Public Function nImprimePlanillaParaSubasta(ByVal psFecSubasta As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal psFechaSis As String, ByVal pnTasaIGV As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroSubasta As String, ByRef psmensaje As String, _
        Optional ByVal psImpresora As Impresoras = gEPSON) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
Dim lmDetalle(20) As String

Dim loCalculos As NCOMColPCalculos

Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency
    
Dim lsNombre As String * 28
Dim vSumGramos As Currency, vSumTasaci As Currency, vSumPresta As Currency
Dim vSumIntVen As Currency, vSumCosCus As Currency, vSumImpues As Currency
Dim vSumCoPrRe As Currency, vSumPreDeu As Currency, vSumPreVen As Currency
Dim vSumSalCap As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double
Dim item As Integer
Dim itemTMP As Integer

Dim ofcar As New COMFunciones.FCOMCadenas
Dim ofun As New COMFunciones.FCOMImpresion

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

    
    lsSQL = " SELECT P.cCtaCod, P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred,  " _
          & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , " _
          & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , " _
          & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , " _
          & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , " _
          & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  " _
          & "               WHERE cLineaCred = C.cLineaCred  " _
          & "               AND nColocLinCredTasaTpo = 2 ),    " _
          & " cNomCliente = (SELECT MAX(cPersNombre) FROM Persona Per INNER JOIN ProductoPersona PPer " _
          & "               ON Per.cPersCod = PPer.cPersCod WHERE PPer.cCtaCod = P.cCtaCod ) , " _
          & " dAdjudica = (SELECT CPRG.dProceso FROM ColocPigRecGar CPRG INNER JOIN ColocPigRGDet CPRGDet " _
          & "               ON CPRG.cTpoProceso = CPRGDet.cTpoProceso And  CPRG.cNroProceso = CPRGDet.cNroProceso WHERE CPRGDet.cCtaCod = P.cCtaCod And CPRG.cTpoProceso = 'A' )  " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod  " _
          & " WHERE P.nPrdEstado in (" & gColPEstAdjud & ") " _
          & " " _
          & " GROUP BY P.cCtaCod, P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred  " _
          & " ORDER BY P.cCtaCod "

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing


    If lrDataRep Is Nothing Then
        psmensaje = " No existen Créditos Pignoraticios para Aviso de Vencimiento "
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1
        
        'CABECERA
        lsCadImp = lsCadImp & nImprimeSubastaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "PLANILLA DE PRENDAS PARA SUBASTA NRO " & psNroSubasta, "", lnPage, 120, "", 1)

        lnIndice = 1:        lnLineas = 7
        
        vSumGramos = 0:        vSumTasaci = 0
        vSumPresta = 0:        vSumIntVen = 0
        vSumCosCus = 0:        vSumImpues = 0
        vSumCoPrRe = 0:        vSumPreDeu = 0
        vSumPreVen = 0:        vSumSalCap = 0
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                
                v14 = !nK14: v14 = Round(v14, 2)
                v16 = !nK16: v16 = Round(v16, 2)
                v18 = !nK18: v18 = Round(v18, 2)
                v21 = !nK21: v21 = Round(v21, 2)
                
                lnPreBase = (v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)
                lnPreBase = lnPreBase * (1 + pnTasaIGV)
                lnPreBase = Round(nRedondeaPrecioSubasta(lnPreBase), 2)
                
                lnPreVenta = lnPreBase
                lnGramos = v14 + v16 + v18 + v21
                
                
                lsNombre = ofcar.PstaNombre(!cNomCliente)
        

                lnLineas = lnLineas + 1
                    
                vSumGramos = vSumGramos + lnGramos
                vSumTasaci = vSumTasaci + !nTasacion
                vSumPresta = vSumPresta + !nMontoCol
                vSumSalCap = vSumSalCap + !nSaldo
                vSumIntVen = vSumIntVen + vInteresVencido
                vSumCosCus = vSumCosCus + vCostoCustodiaMoratorio
                vSumImpues = vSumImpues + vImpuesto
                vSumCoPrRe = vSumCoPrRe + vCostoPreparacionRemate
                vSumPreDeu = vSumPreDeu + lnDeuda
                vSumPreVen = vSumPreVen + lnPreVenta
                
'                    Set loMuestraContrato = New COMDColocPig.DCOMColPContrato
'                        Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(!cCtaCod)
'                    Set loMuestraContrato = Nothing
'
'                    item = 0
'                    Do While Not lrCredPigJoyasDet.EOF
'                        lmDetalle(item) = ofun.ImpreFormat(lrCredPigJoyasDet!npiezas, 4, 0) & " " & ofun.ImpreFormat(lrCredPigJoyasDet!cDescrip, 27) & " " _
'                                    & lrCredPigJoyasDet!cKilataje & "K " & ofun.ImpreFormat(lrCredPigJoyasDet!npesoneto, 4, 2) & "Gr"
'                        item = item + 1
'                        lrCredPigJoyasDet.MoveNext
'                    Loop
                  
                                        
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 4, 0) & Space(1) & !cCtaCod & Space(1) & Format(!dVenc, "dd/mm/yyyy") _
                    & ofun.ImpreFormat(ofun.ImpreCarEsp(lsNombre), 20, 1) _
                    & ofun.ImpreFormat(lnGramos, 6) & ofun.ImpreFormat(!nTasacion, 8) & ofun.ImpreFormat(!nMontoCol, 8) _
                    & Format(!dVenc, "dd/mm/yyyy") & ofun.ImpreFormat(lnPreVenta, 8) & Space(2) & Mid(!cAgeBovedA, 1, 2) _
                    & oFunI.gPrnSaltoLinea
                
'                 For itemTMP = 1 To item - 1
'                            lsCadImp = lsCadImp & Space(49) & lmDetalle(itemTMP) & oFunI.gPrnSaltoLinea
'                            lnLineas = lnLineas + 1
'                 Next
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                        
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nImprimeSubastaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "PLANILLA DE PRENDAS PARA SUBASTA NRO " & psNroSubasta, "", lnPage, 120, "", 1)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(15) & "RESUMEN " & ofun.ImpreFormat(vSumGramos, 41) & ofun.ImpreFormat(vSumTasaci, 8) _
                & ofun.ImpreFormat(vSumPresta, 8) & ofun.ImpreFormat(vSumSalCap, 8) & ofun.ImpreFormat(vSumIntVen, 6) _
                & ofun.ImpreFormat(vSumCosCus, 6) & ofun.ImpreFormat(vSumImpues, 6) & ofun.ImpreFormat(vSumCoPrRe, 6) _
                & ofun.ImpreFormat(vSumPreDeu, 8) & ofun.ImpreFormat(vSumPreVen, 8) _
                & oFunI.gPrnSaltoLinea
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimePlanillaParaSubasta = lsCadBuffer
    Set oFunI = Nothing
End Function

Private Function nRedondeaPrecioSubasta(ByVal pValor As Currency) As Currency
If (pValor - Int(pValor)) = 0 Then  ' Si es entero
    If (pValor Mod 5) = 0 Then
        nRedondeaPrecioSubasta = pValor
    Else
        nRedondeaPrecioSubasta = (Int(pValor / 5) * 5) + 5
    End If
Else    'Si no es entero
    nRedondeaPrecioSubasta = (Int(pValor / 5) * 5) + 5
End If
End Function

'Imprime Listado para Subasta
Public Function nImprimeListadoParaSubasta(ByVal psFecSubasta As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal psFechaSis As String, ByVal pnTasaIGV As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroSubasta As String, _
        ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim lrCredPigJoyasDet As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
Dim lmDetalle(20) As String
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim loCalculos As NCOMColPCalculos

Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer

Dim item As Integer
Dim itemTMP As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double
    
Dim ofcar As New COMFunciones.FCOMCadenas
Dim ofun As New COMFunciones.FCOMImpresion

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

    lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote,   " _
          & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , " _
          & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , " _
          & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , " _
          & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , " _
          & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  " _
          & "               WHERE cLineaCred = C.cLineaCred  " _
          & "               AND nColocLinCredTasaTpo = 2 )    " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod  " _
          & " WHERE P.nPrdEstado in (" & gColPEstAdjud & ") " _
          & "  " _
          & " GROUP BY P.cCtaCod, P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote  " _
          & " ORDER BY P.cCtaCod "

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing


    If lrDataRep Is Nothing Then
        psmensaje = " No existen Créditos Pignoraticios para Aviso de Vencimiento "
        'MuestraImpresion = False

        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nImprimeSubastaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "LISTADO DE CONTRATOS PARA SUBASTA NRO " & psNroSubasta, "DEL " & Format(psFecSubasta, "dd/mm/yyyy"), lnPage, 130, "", 2)

        lnIndice = 1:        lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                
                v14 = !nK14: v14 = Round(v14, 2)
                v16 = !nK16: v16 = Round(v16, 2)
                v18 = !nK18: v18 = Round(v18, 2)
                v21 = !nK21: v21 = Round(v21, 2)
                vCuenta = 0
                lnPreBase = (v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)
                lnPreBase = lnPreBase * (1 + pnTasaIGV)
                lnPreBase = Round(nRedondeaPrecioSubasta(lnPreBase), 2)
                
                lnPreVenta = lnPreBase
                lnGramos = v14 + v16 + v18 + v21
                
                'Set loCalculos = New NCOMColPCalculos
                '    If vDiasAtraso <= 0 Then
                '        vDiasAtraso = 0
                '        vInteresVencido = 0
                '        vCostoCustodiaMoratorio = 0
                '        vImpuesto = 0
                '    Else
                '        vInteresVencido = loCalculos.nCalculaInteresMoratorio(!nSaldo, !nTasaIntVenc, vDiasAtraso)
                '        vInteresVencido = Round(vInteresVencido, 2)
                '        vCostoCustodiaMoratorio = loCalculos.nCalculaCostoCustodiaMoratorio(!nTasacion, pnTasaCustodiaVencida, vDiasAtraso)
                '        vCostoCustodiaMoratorio = Round(vCostoCustodiaMoratorio, 2)
                '        vImpuesto = (vInteresVencido + vCostoCustodiaMoratorio) * pnTasaImpuesto
                '        vImpuesto = Round(vImpuesto, 2)
                '   End If
                '   vCostoPreparacionRemate = 0
                '    If !nPrdEstado = gColPEstPRema Then    ' Si esta en via de Remate
                '        vCostoPreparacionRemate = pnTasaPreparacionRemate * !nTasacion
                '        vCostoPreparacionRemate = Round(vCostoPreparacionRemate, 2)
                '    End If
                '    lnDeuda = !nSaldo + vInteresVencido + vCostoCustodiaMoratorio + vImpuesto + vCostoPreparacionRemate
        
                '    If (lnDeuda * pnFactorPrecioVenta) > lnPreBase Then
                '        lnPreVenta = loCalculos.nCalculaPrecioVentaRemate(lnDeuda * pnFactorPrecioVenta)
                '    Else
                '        lnPreVenta = loCalculos.nCalculaPrecioVentaRemate(lnPreBase)
                '    End If
                '    lnPreVenta = Round(lnPreVenta, 2)
                'Set loCalculos = Nothing

                lnLineas = lnLineas + 1
                                       
                Set loMuestraContrato = New COMDColocPig.DCOMColPContrato
                    Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(!cCtaCod)
                Set loMuestraContrato = Nothing

                item = 0
                Do While Not lrCredPigJoyasDet.EOF
                    lmDetalle(item) = ofun.ImpreFormat(lrCredPigJoyasDet!nPiezas, 4, 0) & " " & ofun.ImpreFormat(lrCredPigJoyasDet!cDescrip, 27) & " " _
                                & lrCredPigJoyasDet!cKilataje & "K " & ofun.ImpreFormat(lrCredPigJoyasDet!npesoneto, 4, 2) & "Gr"
                    item = item + 1
                    lrCredPigJoyasDet.MoveNext
                Loop
                                        
                 lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !cCtaCod & Space(1) & ofun.ImpreFormat(!nPiezas, 3, 0) & Space(2) & lmDetalle(0) _
                    & ofun.ImpreFormat(v14, 6) & ofun.ImpreFormat(v16, 6) & ofun.ImpreFormat(v18, 6) & ofun.ImpreFormat(v21, 6) _
                    & ofun.ImpreFormat(lnPreVenta, 8) _
                    & oFunI.gPrnSaltoLinea
                    
                 For itemTMP = 1 To item - 1
                            lsCadImp = lsCadImp & Space(30) & lmDetalle(itemTMP) & oFunI.gPrnSaltoLinea
                            lnLineas = lnLineas + 1
                 Next

                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                        

                Do While vCuenta < 15
                    vCuenta = vCuenta + 1
                    If Len(ofun.QuiebreTexto(!cLote, vCuenta + 1)) > 0 Then
                        lsCadImp = lsCadImp & Space(28) & ofun.QuiebreTexto(!cLote, vCuenta + 1) & oFunI.gPrnSaltoLinea
                        lnLineas = lnLineas + 1
                    End If
                Loop
                
                lnLineas = lnLineas + 1
                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
                
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nImprimeSubastaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "LISTADO DE CONTRATOS PARA SUBASTA NRO " & psNroSubasta, "DEL " & Format(psFecSubasta, "dd/mm/yyyy"), lnPage, 120, "", 2)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(120, "*") & oFunI.gPrnSaltoLinea
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimeListadoParaSubasta = lsCadBuffer
    Set oFunI = Nothing
End Function
Public Function nIniciaSubasta(ByVal psNroSubasta As String, ByVal psFecSubasta As String, _
        ByVal pnTasaIGV As Double, ByVal pnPrecioOro14 As Double, ByVal pnPrecioOro16 As Double, _
        ByVal pnPrecioOro18 As Double, ByVal pnPrecioOro21 As Double, ByVal psProcesoAgencia As String) As Boolean

Dim lsSQL As String
Dim loDataIniSub As COMDColocPig.DCOMColPFunciones
Dim lrDataInicia As New ADODB.Recordset
              
Dim loCalcula As NCOMColPCalculos
Dim loRegPig As COMDColocPig.DCOMColPActualizaBD


Dim lnPreBase As Currency


lsSQL = " SELECT c.cCtacod, c.dVenc, cp.nTasacion, c.nMontoCol, p.nSaldo, p.nprdEstado, C.cLineaCred, " _
    & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , " _
    & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , " _
    & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , " _
    & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , " _
    & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  " _
    & "               WHERE cLineaCred = C.cLineaCred  " _
    & "               AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ) " _
    & " FROM Producto p Inner Join Colocaciones c on p.cctacod = c.cctacod " _
    & " Inner join ColocPignoraticio cp on c.cctacod = cp.cctacod " _
    & " Left Join ColocPigjoya cpjoy on cp.cctacod = cpjoy.cctacod " _
    & " WHERE P.nPrdEstado IN (" & gColPEstAdjud & " )  " _
    & " GROUP By c.cCtacod, c.dVenc, cp.nTasacion, c.nMontoCol, p.nSaldo, p.nprdEstado, C.cLineaCred  "

Set loDataIniSub = New COMDColocPig.DCOMColPFunciones
    Set lrDataInicia = loDataIniSub.dObtieneRecordSet(lsSQL)
Set loDataIniSub = Nothing

If lrDataInicia Is Nothing Then

Else
    lnPreBase = 0
    With lrDataInicia
        Set loRegPig = New COMDColocPig.DCOMColPActualizaBD
        
        Do While Not .EOF
            'Precio Base
            lnPreBase = !nK14 * pnPrecioOro14 + !nK16 * pnPrecioOro16 + !nK18 * pnPrecioOro18 + !nK21 * pnPrecioOro21
            lnPreBase = lnPreBase * (1 + pnTasaIGV)
            lnPreBase = Round(lnPreBase, 2)
            'Precio Deuda
            Set loCalcula = New NCOMColPCalculos
                lnPreBase = loCalcula.nCalculaValorAdjudicacion(!nSaldo, lnPreBase)
            Set loCalcula = Nothing
            lnPreBase = Round(lnPreBase, 2)

            'Grabar información en Detalle de Subasta
            
            '** Actualiza ColocPignoraticio
            'Call loRegPig.dUpdateColocPignoraticio(!cCtaCod, , , , , , , , , , , False)
            
            Call loRegPig.dInsertColocPigRGDet("S", psNroSubasta, !cCtaCod, gColPRecGarVentaEnProceso, 0, lnPreBase, 0, 0, 0, psProcesoAgencia, False)
                
            .MoveNext
        Loop
        .Close
        Set lrDataInicia = Nothing
    End With
            
    Call loRegPig.dUpdateColocPigRecGar("S", psNroSubasta, psProcesoAgencia, , gColPRecGarEstIniciado, , False)
     
    Set loRegPig = Nothing
    
End If

End Function

Public Function nFinalizaSubasta(ByVal psNroSubasta As String, ByVal psFecProceso As String, ByVal psProcesoCadaAgencia As String) As Boolean

Dim lsSQL As String
Dim lsNextSubasta As String
              
Dim loRegPig As COMDColocPig.DCOMColPActualizaBD

    'Genera el sgte Nro de Subasta
    lsNextSubasta = String(4 - Len(Trim(Str(Val(psNroSubasta) + 1))), "0") & Trim(Str(Val(psNroSubasta) + 1))
    
    Set loRegPig = New COMDColocPig.DCOMColPActualizaBD
    
        'Cambia estado de los que NO se han vendido
        lsSQL = " UPDATE ColocPigRGDet SET nRGDetEstado = " & gColPRecGarVentaNoVendido & " WHERE cTpoProceso = 'S' AND cNroProceso = '" & psNroSubasta & "' AND nRGDetEstado = " & gColPRecGarVentaEnProceso
        Call loRegPig.dEjecutar(lsSQL)
    
        'Termina el Subasta
        Call loRegPig.dUpdateColocPigRecGar("S", psNroSubasta, psProcesoCadaAgencia, , gColPRecGarEstTerminado, , False)
        
        'Inserta la Sgte subasta
        Call loRegPig.dInsertColocPigRecGar("S", lsNextSubasta, Format(DateAdd("m", 1, Format(psFecProceso, "mm/dd/yyyy")) & " 08:00", "mm/dd/yyyy hh:mm"), gColPRecGarEstNoIniciado, psProcesoCadaAgencia, , False)
     
    Set loRegPig = Nothing

End Function

Public Function nObtieneSobranteRemateProceso(ByVal psNroProceso As String, ByVal psAgencia As String, ByRef psmensaje As String) As Currency
'Obtiene el Sobrante del Ultimo Remate
Dim loValida As COMDColocPig.DCOMColPFunciones
Dim lrs As New ADODB.Recordset
Dim lnSobrante As Currency
Dim lsSQL As String
Dim lsAgencia As String

On Error GoTo dError

lsSQL = "Select Isnull(sum(nsobrante),0) nSobrante " _
    & " FROM ColocPigRGDet Rd join colocpigrecupVta Rv " _
    & " On Rd.cTpoProceso = Rv.cTpoProceso and Rd.cNroProceso = Rv.cNroProceso and Rd.cCtaCod = Rv.cCtaCod " _
    & " WHERE Rd.cTpoProceso ='R' And Rd.cNroProceso ='" & psNroProceso _
    & "' And nRGDetEstado ='1' AND cCodAge ='" & psAgencia & "' "

'Set lrs = New Recordset
    Set loValida = New COMDColocPig.DCOMColPFunciones
        Set lrs = loValida.dObtieneRecordSet(lsSQL)
    Set loValida = Nothing
    If lrs.BOF And lrs.EOF Then
        psmensaje = "No se tiene Información de Sobrante de Remate"
    Else
        lnSobrante = lrs!nSobrante
    End If
Set lrs = Nothing

nObtieneSobranteRemateProceso = lnSobrante
Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function


'****
Public Function nObtieneRemateIniciado() As String

Dim loConec As COMConecta.DCOMConecta
Dim lrDatosRem As New ADODB.Recordset
Dim lsUltRemate As String
Dim lsSQL As String

lsUltRemate = ""
Set loConec = New COMConecta.DCOMConecta
    loConec.AbreConexion
    Set lrDatosRem = New Recordset
    

    lsSQL = "SELECT RG.cTpoProceso, RG.cNroProceso, RG.nRGEstado " _
    & " FROM ColocPigRecGar RG  WHERE cTpoProceso ='R' " _
    & " ORDER BY cNroProceso DESC "
    loConec.CargaRecordSet lsSQL
    
    lrDatosRem.Open lsSQL, loConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    If lrDatosRem Is Nothing Then
        Exit Function
    End If
    'lrDatosRem.MoveFirst
    If lrDatosRem.RecordCount = 0 Then
        lsUltRemate = ""
    Else
        If lrDatosRem!nRGEstado = gColPRecGarEstIniciado Then
            lsUltRemate = lrDatosRem!cNroProceso
        End If
    End If
    lrDatosRem.Close
Set loConec = Nothing
nObtieneRemateIniciado = lsUltRemate
End Function

Public Function nObtieneCtaSobranteRemate(ByVal psRemateCadaAgencia As String, ByRef psmensaje As String) As String
Dim loConstSis As COMDConstSistema.NCOMConstSistema
Dim PObjConec As COMDConstantes.DCOMConstantes
Dim lrCtaAho As New ADODB.Recordset
Dim lsCtaSobranteRemate As String
Dim lsnConscod As ConstanteCabecera
    ' Obtiene la cuenta de Ahorros de Sobrante de Remate  ****
    If Right(psRemateCadaAgencia, 2) = "00" Then  ' Remate Centralizado
        Set loConstSis = New COMDConstSistema.NCOMConstSistema
            lsCtaSobranteRemate = loConstSis.LeeConstSistema(61)  ' Cuenta SobranteRemate
        Set loConstSis = Nothing
    Else ' Remate en Cada Agencia
        Set PObjConec = New COMDConstantes.DCOMConstantes
        lsnConscod = "3207"
        lsCtaSobranteRemate = PObjConec.DameDescripcionConstante(lsnConscod, Right(psRemateCadaAgencia, 2))
        If lsCtaSobranteRemate = "" Then
           psmensaje = "No se han configurado las Ctas de Ahorro de Sobrantes"
           lsCtaSobranteRemate = ""
           Exit Function
        End If
'        Do While Not lrCtaAho.EOF
'           If Val(lrCtaAho!nConsValor) = Val(Right(psRemateCadaAgencia, 2)) Then
'               lsCtaSobranteRemate = Trim(lrCtaAho!cConsDescripcion)
'               Exit Do
'           End If
'           lrCtaAho.MoveNext
'        Loop
'        Set lrCtaAho = Nothing
    End If
    nObtieneCtaSobranteRemate = lsCtaSobranteRemate
End Function


Public Function ImprimirComision(ByVal pCodCta As String, ByVal pNomAdj As String, pMonCom As Currency, Optional ByVal PVENBRUTA As Currency, Optional ByVal PVENNETA As Currency, Optional ByVal PIGV As Currency, _
                                 Optional ByVal pdfecha As Date, Optional ByVal psUser As String, Optional ByVal psAge As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
    Dim vNombre As String * 27
    Dim vEspacio As Integer
    Dim lsCadImp As String
    Dim ofun As New COMFunciones.FCOMImpresion
    Dim oFunI As New COMFunciones.FCOMVarImpresion
    oFunI.Inicia psImpresora

    vNombre = pNomAdj
    
    'vEspacio = 9
    vEspacio = 34
    'ImpreBegChe True, 22
        
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69)
        lsCadImp = lsCadImp & Space(15) & " CREDITO PIGNORATICIO" & Space(23 + vEspacio) & "CREDITO PIGNORATICIO" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
        lsCadImp = lsCadImp & Space(5) & ofun.ImpreFormat(psAge, 26, 0) & Format(pdfecha & " " & Time, "dd/mm/yyyy hh:mm") & Space(vEspacio) & ofun.ImpreFormat(psAge, 26, 0) & Format(pdfecha & " " & Time, "dd/mm/yyyy hh:mm") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(69) 'Activa Negrita
        lsCadImp = lsCadImp & Space(4) & " CONTRATO     :" & Space(28 - Len(pCodCta)) & pCodCta & Space(vEspacio) & "CONTRATO     :" & Space(28 - Len(pCodCta)) & pCodCta & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(4) & "ADJUDICATARIO: " & vNombre & Space(vEspacio) & "ADJUDICATARIO: " & vNombre & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(4) & "COMISION DE REMATE:" & ofun.ImpreFormat(pMonCom, 20, , True) & Space(vEspacio) & "COMISION DE REMATE:" & ofun.ImpreFormat(pMonCom, 20, , True) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(4) & "SUBTOTAL          :" & ofun.ImpreFormat(PVENBRUTA, 20, , True) & Space(vEspacio) & "SUBTOTAL          :" & ofun.ImpreFormat(PVENBRUTA, 20, , True) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(4) & "ITF               :" & ofun.ImpreFormat(PIGV, 20, , True) & Space(vEspacio) & "ITF              :" & ofun.ImpreFormat(PIGV, 20, , True) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(4) & "VENTA NETA        :" & ofun.ImpreFormat(PVENNETA, 20, , True) & Space(vEspacio) & "VENTA NETA        :" & ofun.ImpreFormat(PVENNETA, 20, , True) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(4) & "Venta en Remate   " & Space(20) & Format(psUser, "@@@@") & Space(vEspacio) & "Venta en Remate   " & Space(20) & Format(psUser, "@@@@") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Chr$(27) & Chr$(70)    'Desactiva Negrita
    'ImpreEnd
    ImprimirComision = lsCadImp
    Set oFunI = Nothing
End Function

Public Function CalculaPrecioVentaRemate(vValor As Double) As Double
    If (vValor - Int(vValor)) = 0 Then  ' Si es entero
        If (vValor Mod 5) = 0 Then
            CalculaPrecioVentaRemate = vValor
        Else
            CalculaPrecioVentaRemate = (Int(vValor / 5) * 5) + 5
        End If
    Else    'Si no es entero
        CalculaPrecioVentaRemate = (Int(vValor / 5) * 5) + 5
    End If
End Function

'**DAOR 20080131, Obtiene adjudicaciones para venta en lote *******
Public Function ObtieneAdjudicacionesParaVentaEnLote(ByVal psCodAge As String) As ADODB.Recordset
Dim loDCOMColPFunc As COMDColocPig.DCOMColPFunciones
Dim lsSQL As String

On Error GoTo ObtieneAdjudicacionesParaVentaEnLoteErr

    lsSQL = " exec stp_sel_AdjudicacionesParaVentaEnLote '" & psCodAge & "'"
    
    Set loDCOMColPFunc = New COMDColocPig.DCOMColPFunciones
    Set ObtieneAdjudicacionesParaVentaEnLote = loDCOMColPFunc.dObtieneRecordSet(lsSQL)
    Set loDCOMColPFunc = Nothing
                
    Exit Function

ObtieneAdjudicacionesParaVentaEnLoteErr:
    Err.Raise Err.Number, "Obtiene adjudicaciones para venta en lote <<ObtieneAdjudicacionesParaVentaEnLote>>", Err.Description
End Function

'**DAOR 20080131, Obtiene creditos para venta en lote *******
Public Function ObtieneCreditosParaVentaEnLote(ByVal psListaAdj As String) As ADODB.Recordset
Dim loDCOMColPFunc As COMDColocPig.DCOMColPFunciones
Dim lsSQL As String

On Error GoTo ObtieneCreditosParaVentaEnLoteErr

    lsSQL = " exec stp_sel_CreditosParaVentaEnLote '" & psListaAdj & "'"
    
    Set loDCOMColPFunc = New COMDColocPig.DCOMColPFunciones
    Set ObtieneCreditosParaVentaEnLote = loDCOMColPFunc.dObtieneRecordSet(lsSQL)
    Set loDCOMColPFunc = Nothing
                
    Exit Function

ObtieneCreditosParaVentaEnLoteErr:
    Err.Raise Err.Number, "Obtiene creditos para venta en lote <<ObtieneCreditosParaVentaEnLote>>", Err.Description
End Function

'**DAOR 20080131, Venta en Lote
Public Function VentaEnLote(ByVal psMovNro As String, ByVal psFechaHora As String, ByVal pMatCuentas As Variant, _
        ByVal psNroVent As String, ByVal psFecVentaLote As String, _
        ByVal pnPrecioOro14 As Double, ByVal pnPrecioOro16 As Double, ByVal pnPrecioOro18 As Double, ByVal pnPrecioOro21 As Double, _
        ByVal psProcesoCadaAgencia As String, ByRef psmensaje As String) As Boolean
Dim loRegPig As COMDColocPig.DCOMColPActualizaBD
Dim loNroPro As COMDColocPig.DCOMColPFunciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String
Dim lnMovNro As Long, i As Integer
Dim lsOpeCod As String
Dim lsSgteVenta As String
Dim lnPreMercado As Currency, lnPreDeuda As Currency, lnPreVen As Currency
'*** PEAC 20190807
Dim lnTasac As Double
Dim lrDataAdj As ADODB.Recordset
Dim loDataRecup As New COMDColocPig.DCOMColPFunciones
'*** FIN PEAC

On Error GoTo VentaEnLoteErr

    '**DAOR 20080606, Corrección debido a problema de compilación***
    'lsOpeCod = geColPVentaLote
    lsOpeCod = gColPOpeVentaLote
    '****************************************************************
    lnPreMercado = 0: lnPreDeuda = 0: lnPreVen = 0
    
    Set loRegPig = New COMDColocPig.DCOMColPActualizaBD
    loRegPig.dBeginTrans
    mbTrans = True
        
    Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Venta en Lote Cred. Pign. Adjudicados", gMovEstContabNoContable, gMovFlagVigente, False)
    lnMovNro = loRegPig.dGetnMovNro(psMovNro)
        
    For i = 0 To UBound(pMatCuentas) - 1
        
        '*** PEAC 20190807
        lnTasac = 0
        
        lsSQL = "exec stp_sel_DatosParaRecupContraAdjudiCustDif '" & pMatCuentas(i, 0) & "'"
        
        Set loDataRecup = New COMDColocPig.DCOMColPFunciones
            Set lrDataAdj = loDataRecup.dObtieneRecordSet(lsSQL)
        Set loDataRecup = Nothing
        
        If Not (lrDataAdj.BOF And lrDataAdj.EOF) Then
            lnTasac = lrDataAdj!nTasacion
        End If
        '*** FIN PEAC

        lnPreMercado = pMatCuentas(i, 3) * pnPrecioOro14 + pMatCuentas(i, 4) * pnPrecioOro16 + pMatCuentas(i, 5) * pnPrecioOro18 + pMatCuentas(i, 6) * pnPrecioOro21
        lnPreMercado = Round(lnPreMercado, 2)
                
        lnPreVen = IIf(pMatCuentas(i, 1) > lnPreMercado, lnPreMercado, pMatCuentas(i, 1))
        lnPreVen = Round(lnPreVen, 2)
        
        Call loRegPig.dUpdateProducto(pMatCuentas(i, 0), , , gColPEstChafa, psFechaHora, -2, False)
        Call loRegPig.dUpdateColocaciones(pMatCuentas(i, 0), , , , , psMovNro, , False)
        
        '*** PEAC 20090918
        'Call loRegPig.dInsertColocacEstado(pMatCuentas(i, 0), psFechaHora, gColPEstChafa, 1, lnPreVen, "Vendido (Proceso venta en lote)", gColocCalendCodFFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
        Call loRegPig.dInsertColocacEstado(pMatCuentas(i, 0), psFechaHora, gColPEstChafa, 1, pMatCuentas(i, 7), "Vendido (Proceso venta en lote)", gColocCalendCodFFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
        
        'Call loRegPig.dInsertColocPigRGDet("V", psNroVent, pMatCuentas(i, 0), gColPRecGarVentaVendido, lnPreVen, 0, lnPreMercado, lnPreVen, 0, psProcesoCadaAgencia, False)
        Call loRegPig.dInsertColocPigRGDet("V", psNroVent, pMatCuentas(i, 0), gColPRecGarVentaVendido, pMatCuentas(i, 7), 0, lnPreMercado, pMatCuentas(i, 7), 0, psProcesoCadaAgencia, False)
        
        'Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, pMatCuentas(i, 0), 1, lnPreVen, 0, "", 0, gColPEstChafa, pMatCuentas(i, 1), False)
        Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, pMatCuentas(i, 0), 1, pMatCuentas(i, 7), 0, "", 0, gColPEstChafa, pMatCuentas(i, 1), False)
        
        'Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, pMatCuentas(i, 0), 1, gColPConceptoCodCapital, 1, pMatCuentas(i, 1), False)
        'Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, pMatCuentas(i, 0), 1, gColPConceptoCodTasacion, 1, pMatCuentas(i, 2), False)
        Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, pMatCuentas(i, 0), 1, gColPConceptoCodCapitalVencido, 1, pMatCuentas(i, 7), False)
        '*** FIN PEAC
        
        '*** PEAC 20190807
        If lnTasac > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, pMatCuentas(i, 0), 1, "2218", 1, lnTasac, False)
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, pMatCuentas(i, 0), 1, "2219", 1, lnTasac, False)
        End If
        '*** FIN PEAC
        
        If pMatCuentas(i, 3) > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, pMatCuentas(i, 0), 1, gColPConceptoCodOro14, 1, pMatCuentas(i, 3), False)
        End If
        If pMatCuentas(i, 4) > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, pMatCuentas(i, 0), 1, gColPConceptoCodOro16, 1, pMatCuentas(i, 4), False)
        End If
        If pMatCuentas(i, 5) > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, pMatCuentas(i, 0), 1, gColPConceptoCodOro18, 1, pMatCuentas(i, 5), False)
        End If
        If pMatCuentas(i, 6) > 0 Then
            Call loRegPig.dInsertMovColDet(lnMovNro, lsOpeCod, pMatCuentas(i, 0), 1, gColPConceptoCodOro21, 1, pMatCuentas(i, 6), False)
        End If
    Next i

    Call loRegPig.dUpdateColocPigRecGar("V", psNroVent, psProcesoCadaAgencia, , gColPRecGarEstTerminado, , False)

    Set loNroPro = New COMDColocPig.DCOMColPFunciones
        lsSQL = " exec stp_sel_UltimoNumProcesoPignoraticio 'V'"
        Set lrs = loNroPro.dObtieneRecordSet(lsSQL)
        If Not (lrs.EOF And lrs.BOF) Then
            lsSgteVenta = String(4 - Len(Trim(Str(Val(lrs!cNroProceso) + 1))), "0") & Trim(Str(Val(lrs!cNroProceso) + 1))
        End If
    Set loNroPro = Nothing
        
    Call loRegPig.dInsertColocPigRecGar("V", lsSgteVenta, Format(DateAdd("m", 1, psFecVentaLote) & " 08:00", "mm/dd/yyyy hh:mm"), gColPRecGarEstNoIniciado, psProcesoCadaAgencia, , False)
        
    loRegPig.dCommitTrans
    Set loRegPig = Nothing
    Exit Function
    
VentaEnLoteErr:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error VentaEnLote", "Error en Funcion de Registro de Venta en Lote "

End Function


'**DAOR 20080131, Reporte de Creditos de Venta en Lote
'Imprime Listado de Adjudicados
Public Function nImprimeListadoVentaEnLote(ByVal psFecProceso As String, ByVal psAgencia As String, _
        ByVal psNroProceso As String, ByVal psFechaSis As String, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByRef psmensaje As String, _
        Optional ByVal psImpresora As Impresoras = gEPSON) As String
    
Dim lsSQL As String
Dim loColPFunc As COMDColocPig.DCOMColPFunciones
Dim lrDataRep As New ADODB.Recordset
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loCalculos As NCOMColPCalculos
Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer
Dim lnDeuda As Double
Dim lnGramos As Currency
Dim vSumGra14 As Currency, vSumGra16 As Currency, vSumGra18 As Currency, vSumGra21 As Currency
Dim vSumGramos As Currency, vSumTasaci As Currency, vSumSalCap As Currency, vSumDeuda As Currency
Dim vSumPreMerc As Currency, vSumValRegis As Currency
Dim lsNombre As String * 28
Dim loFCOMCade As New COMFunciones.FCOMCadenas
Dim loFun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion

    oFunI.Inicia psImpresora

    lsSQL = " exec stp_sel_CreditosDeVentaEnLote '" & psNroProceso & "'"

    Set loColPFunc = New COMDColocPig.DCOMColPFunciones
    Set lrDataRep = loColPFunc.dObtieneRecordSet(lsSQL)
    Set loColPFunc = Nothing


    If lrDataRep Is Nothing Then
        psmensaje = " No existen créditos pignoraticios adjudicados vendidos en este proceso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1
        'Cabecera
        lsCadImp = lsCadImp & nImprimeVentaLoteReporteCabecera(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                        "Listado de Contratos de Venta en Lote Nro " & psNroProceso, "Del " & Format(psFecProceso, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 144, "", 1)
        lnIndice = 1
        lnLineas = 7
        vSumGra14 = 0:  vSumGra16 = 0:  vSumGra18 = 0:  vSumGra21 = 0
        vSumGramos = 0: vSumTasaci = 0: vSumSalCap = 0: vSumDeuda = 0
        vSumPreMerc = 0: vSumValRegis = 0
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnLineas = lnLineas + 1
                v14 = !nK14: v14 = Round(v14, 2)
                v16 = !nK16: v16 = Round(v16, 2)
                v18 = !nK18: v18 = Round(v18, 2)
                v21 = !nK21: v21 = Round(v21, 2)
                vCuenta = 0
                lnGramos = v14 + v16 + v18 + v21
                lsNombre = loFCOMCade.PstaNombre(!cPersNombre)
                lnDeuda = !nDeuda
                
                vSumGra14 = vSumGra14 + v14
                vSumGra16 = vSumGra16 + v16
                vSumGra18 = vSumGra18 + v18
                vSumGra21 = vSumGra21 + v21
                vSumGramos = vSumGramos + lnGramos
                vSumTasaci = vSumTasaci + !nTasacion
                vSumSalCap = vSumSalCap + !nSaldo
                vSumDeuda = vSumDeuda + lnDeuda
                                        
                lsCadImp = lsCadImp & loFun.ImpreFormat(lnIndice, 5, 0)
                lsCadImp = lsCadImp & Space(1) & !CodSiafc
                lsCadImp = lsCadImp & Space(1) & !cCtaCod
                lsCadImp = lsCadImp & Space(1) & loFun.ImpreFormat(loFun.ImpreCarEsp(lsNombre), 20, 1)
                lsCadImp = lsCadImp & Space(1) & loFun.ImpreFormat(v14, 6)
                lsCadImp = lsCadImp & Space(1) & loFun.ImpreFormat(v16, 6)
                lsCadImp = lsCadImp & Space(1) & loFun.ImpreFormat(v18, 6)
                lsCadImp = lsCadImp & Space(1) & loFun.ImpreFormat(v21, 6)
                lsCadImp = lsCadImp & Space(1) & loFun.ImpreFormat(lnGramos, 6)
                lsCadImp = lsCadImp & Space(1) & loFun.ImpreFormat(!nTasacion, 8)
                lsCadImp = lsCadImp & Space(1) & loFun.ImpreFormat(!nSaldo, 8)
                lsCadImp = lsCadImp & Space(1) & loFun.ImpreFormat(!cNroAdjudica, 8)
                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
               
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                                        
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nImprimeVentaLoteReporteCabecera(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                        "Listado de Contratos de Venta en Lote Nro " & psNroProceso, "Del " & Format(psFecProceso, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 144, "", 1)
                    lnLineas = 7
                End If
                
                .MoveNext
            Loop
        End With
        
        lsCadImp = lsCadImp & String(144, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(64) & "RESUMEN "
        lsCadImp = lsCadImp & loFun.ImpreFormat(vSumGramos, 40)
        lsCadImp = lsCadImp & loFun.ImpreFormat(vSumTasaci, 9)
        lsCadImp = lsCadImp & loFun.ImpreFormat(vSumSalCap, 9)
        lsCadImp = lsCadImp & "" & oFunI.gPrnSaltoLinea
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nImprimeListadoVentaEnLote = lsCadBuffer
    Set oFunI = Nothing
End Function

'**DAOR 20080131
Public Function nImprimeVentaLoteReporteCabecera(ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, _
        ByVal psFechaSis As String, ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String, _
        ByVal pnNroRepor As Integer, Optional ByVal psImpresora As Impresoras = gEPSON) As String
Dim lsCadImp As String
Dim loImprimeCab As NCOMColPImpre
Dim oFunI As New COMFunciones.FCOMVarImpresion
    
    oFunI.Inicia psImpresora
    Set loImprimeCab = New NCOMColPImpre
        lsCadImp = loImprimeCab.nImprimeCabeceraReportes(psNomCmac, psNomAgencia, psCodUser, psFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)
    Set loImprimeCab = Nothing
    
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & String(pnAnchoLinea, "-") & oFunI.gPrnSaltoLinea
    
    Select Case pnNroRepor
        Case 1 ' Planilla de Venta en Lote
            lsCadImp = lsCadImp & Space(1) & "Item    ContratoAnt     Contrato          Cliente                   14kl      16kl     18kl     21kl     Gramos    Tasacion   SaldoCap  NroAdju" & oFunI.gPrnSaltoLinea
    End Select
    
    lsCadImp = lsCadImp & String(pnAnchoLinea, "-") & oFunI.gPrnSaltoLinea
    nImprimeVentaLoteReporteCabecera = lsCadImp
    Set oFunI = Nothing
End Function

'*** PEAC 20080708
Public Function nImprimePlanillaParaAdjudicar(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal pnDiasVctoParaRemate As Integer, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, ByVal pnFactorPrecioVenta As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroRemate As String, ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
Dim fnVarIGV As Double


Dim loCalculos As NCOMColPCalculos

Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency
    
Dim lsNombre As String * 28
Dim vSumGramos As Currency, vSumTasaci As Currency, vSumPresta As Currency
Dim vSumIntVen As Currency, vSumCosCus As Currency, vSumImpues As Currency
Dim vSumCoPrRe As Currency, vSumPreDeu As Currency, vSumPreVen As Currency
Dim vSumSalCap As Currency

'*** PEAC 20090226
Dim vSumIntCom As Currency, vSumIntMor As Currency, vSumNotifi As Currency


Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double

Dim ofcar As New COMFunciones.FCOMCadenas
Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion

'**DAOR 20070714******************
Dim lnEstProRem As Integer
Dim lrEstRem As ADODB.Recordset
'**********************************

''*** PEAC 20090226 ***********************
Dim lnDiasAtraso As Integer
Dim lnIntVencido As Double
Dim lnIntMoratorio As Double
Dim lnIntAdelantado As Double
''*****************************************

'*** PEAC 20190415
Dim lnValorCosCust As Double
Dim lnValorCosCustTasacion As Double
Dim lnValorCosNotif As Double
Dim lnValorCosCustoDif As Double

oFunI.Inicia psImpresora
    '*****************************************************
    'Obtener Fecha de Corte
    Dim oCon As New COMConecta.DCOMConecta
    Dim rs1 As ADODB.Recordset
    Set oCon = New COMConecta.DCOMConecta
        If oCon.AbreConexion Then
           Set rs1 = New ADODB.Recordset
           lsSQL = "Select cConsDescripcion From Constante Where nConsCod='3208' and nConsValor='" & Trim(psAgencia) & "' "
           Set rs1 = oCon.CargaRecordSet(lsSQL)
           If Not (rs1.EOF And rs1.BOF) Then
             psFecRemate = Format(rs1("cConsDescripcion"), "mm/dd/yyyy")
           End If
           Set rs1 = Nothing
        End If
    Set oCon = Nothing
    '*******************************************************
       
    ldFecAviso = DateAdd("d", -pnDiasVctoParaRemate, psFecRemate)

    'lsSQL = " SELECT P.cCtaCod, P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred,  " _
          & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , " _
          & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , " _
          & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , " _
          & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , " _
          & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  " _
          & "               WHERE cLineaCred = C.cLineaCred  " _
          & "               AND nColocLinCredTasaTpo  = " & gColocLineaCredTasasIntMoratNormal & " ),    " _
          & " cNomCliente = (SELECT MAX(cPersNombre) FROM Persona Per INNER JOIN ProductoPersona PPer " _
          & "               ON Per.cPersCod = PPer.cPersCod WHERE PPer.cCtaCod = P.cCtaCod ),  " _
          & " CuentaAntigua=ISNULL(cCtaCodAnt,'                  ')" _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " left JOIN RelCuentas R ON P.cCtaCod=R.cCtacod " _
          & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod  " _
          & " WHERE P.nPrdEstado in (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ") " _
          & " AND Substring(P.cCtaCod,4,2) ='" & psAgencia & "' AND DATEDIFF(dd, C.dVenc, '" & Format(ldFecAviso, "mm/dd/yyyy") & "') >= 0  " _
          & " GROUP BY P.cCtaCod,cCtaCodAnt, P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred  " _
          & " ORDER BY P.cCtaCod "

    lsSQL = " exec stp_sel_PlanillaParaAdjudicar '" & psAgencia & "','" & psNroRemate & "'"

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
 

    If lrDataRep Is Nothing Then
        psmensaje = " No existen Créditos Pignoraticios para Aviso de Vencimiento "
        'MuestraImpresion = False

        Exit Function
    Else
        '**DAOR 20070714**************************
        'lsSQL = "select nRGEstado from ColocPigRecGar " & _
        '        "where cCodAge='109" & psAgencia & "' and cTpoProceso='R' and cNroProceso='" & psNroRemate & "'"
        '
        'Set loDataRep = New COMDColocPig.DCOMColPFunciones
        '    Set lrEstRem = loDataRep.dObtieneRecordSet(lsSQL)
        'Set loDataRep = Nothing
        'If (lrEstRem Is Nothing) Or lrEstRem.EOF Then
        '    lnEstProRem = 0
        'Else
        '    lnEstProRem = lrEstRem("nRGEstado")
        'End If
        'Set lrEstRem = Nothing
        '*****************************************
    
        lnLineas = 0
        lnPage = 1

        
        'CABECERA
        lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "PLANILLA DE CONTRATOS PARA ADJUDICAR " & psNroRemate, "Del " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 190, "", 1)

        lnIndice = 1:        lnLineas = 7
        
        vSumGramos = 0:        vSumTasaci = 0
        vSumPresta = 0:        vSumIntVen = 0
        vSumCosCus = 0:        vSumImpues = 0
        vSumCoPrRe = 0:        vSumPreDeu = 0
        vSumPreVen = 0:        vSumSalCap = 0
        
        vSumIntCom = 0: vSumIntMor = 0: vSumNotifi = 0  ''*** PEAC 20090226
        
        'constante para el igv
        '3007
        
        fnVarIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)

        Dim pnDeudaCosto As Double, pn15DeVenta As Double
        
        With lrDataRep
            Do While Not lrDataRep.EOF
'                If lrDataRep!cCtacod = "108013051000170429" Then Stop
                v14 = !nK14: v14 = Round(v14, 2)
                v16 = !nK16: v16 = Round(v16, 2)
                v18 = !nK18: v18 = Round(v18, 2)
                v21 = !nK21: v21 = Round(v21, 2)
                
                pn15DeVenta = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * 0.15
              
                'lnPreBase = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * (fnVarIGV + 1)
                lnPreBase = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21))

                lnGramos = v14 + v16 + v18 + v21
                
                lnPreBase = Round(lnPreBase, 2)
                
                lsNombre = ofcar.PstaNombre(!cNomCliente)
        
                '*** PEAC 20090226
                'vdiasAtraso = DateDiff("d", !dVenc, psFecRemate)
                vdiasAtraso = DateDiff("d", !dVenc, psFechaSis)
                
                Set loCalculos = New NCOMColPCalculos
                    If vdiasAtraso <= 0 Then
                        vdiasAtraso = 0
                        vInteresVencido = 0
                        lnIntVencido = 0
                        lnIntAdelantado = 0
                        lnIntMoratorio = 0
                        vCostoCustodiaMoratorio = 0
                        vImpuesto = 0
                    Else
                    
                        ''*** PEAC 20090226 - COMENTADO PARA CREAR OTRO
                        'vInteresVencido = loCalculos.nCalculaInteresMoratorio(!nSaldo, !nTasaIntVenc, vdiasAtraso)
                        'vInteresVencido = Round(vInteresVencido, 2)

                        '*** PEAC 20190605 - se traslado para calcular int venc
                        lnIntAdelantado = loCalculos.nCalculaInteresAlVencimiento(!nSaldo, !nTasaIntVenc, 30)
                        lnIntAdelantado = Round(lnIntAdelantado, 2)


                        ''*** PEAC 20090226
                        lnIntMoratorio = loCalculos.nCalculaInteresMoratorio(!nSaldo, !nTasaIntMora, vdiasAtraso)
                        lnIntMoratorio = Round(lnIntMoratorio, 2)
                        
                        '*** PEAC 20190605 - se agregó el parametro lnIntAdelantado
                        lnIntVencido = loCalculos.nCalculaInteresMoratorio(!nSaldo, !nTasaIntVenc, vdiasAtraso, lnIntAdelantado)
                        lnIntVencido = Round(lnIntVencido, 2)
                        
'                        lnIntVencido = loCalculos.nCalculaInteresMoratorio(!nSaldo, !nTasaIntVenc, vdiasAtraso)
'                        lnIntVencido = Round(lnIntVencido, 2)
                        
'                        lnIntAdelantado = loCalculos.nCalculaInteresAlVencimiento(!nSaldo, !nTasaIntVenc, 30)
'                        lnIntAdelantado = Round(lnIntAdelantado, 2)
                        
                        vCostoCustodiaMoratorio = loCalculos.nCalculaCostoCustodiaMoratorio(!nTasacion, pnTasaCustodiaVencida, vdiasAtraso)
                        vCostoCustodiaMoratorio = Round(vCostoCustodiaMoratorio, 2)

                        ''*** PEAC 20090226
'                        vImpuesto = (vInteresVencido + vCostoCustodiaMoratorio) * pnTasaImpuesto
'                        vImpuesto = Round(vImpuesto, 2)
                    End If
                    
                    vCostoPreparacionRemate = 0
                    '**Comentado por DAOR 20070713
                    If !nPrdEstado = gColPEstPRema Then    ' Si esta en via de Remate
                        vCostoPreparacionRemate = pnTasaPreparacionRemate * !nTasacion
                        vCostoPreparacionRemate = Round(vCostoPreparacionRemate, 2)
                    End If
                    
                    '*** PEAC 20190719
                    lnValorCosCustoDif = 0
                    If !nSaldo = 0 And !nPrdEstado = 2102 Then
                        lnValorCosCustoDif = !monto_custo
                    End If
                    
                    ''*** PEAC 20090226 **************
                    'lnDeuda = !nSaldo + vInteresVencido + vCostoCustodiaMoratorio + vImpuesto + vCostoPreparacionRemate
                     'lnDeuda = !nSaldo + Round(lnIntMoratorio, 2) + Round(lnIntVencido, 2) + Round(lnIntAdelantado, 2) + !nNotificacion
                     lnDeuda = !nSaldo + Round(lnIntMoratorio, 2) + Round(lnIntVencido, 2) + Round(lnIntAdelantado, 2) + !nNotificacion + lnValorCosCustoDif '*** PEAC 20190719
                     lnDeuda = Round(lnDeuda, 2)
                    '**********************************
                    
                    
                    '**DAOR 20070714***********************
                    'If !nPrdEstado = gColPEstPRema And lnEstProRem = 1 Then  ' Si esta en via de Remate y además el proceso se encuentra iniciado
                    '    vCostoPreparacionRemate = pnTasaPreparacionRemate * !nTasacion
                    '    vCostoPreparacionRemate = Round(vCostoPreparacionRemate, 2)
                    'End If
                    'lnDeuda = !nSaldo + vInteresVencido + vCostoCustodiaMoratorio + vImpuesto + vCostoPreparacionRemate
                    '***************************************
                    
                    'pnDeudaCosto = Round((lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta) * (fnVarIGV + 1), 2)
                    pnDeudaCosto = Round((lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta), 2)
                    
                    lnPreVenta = loCalculos.nCalculaPrecioVentaRemate(pnDeudaCosto, lnPreBase, pnFactorPrecioVenta)
                    lnPreVenta = CalculaPrecioVentaRemate(Round(lnPreVenta, 0))
                    
                    
                Set loCalculos = Nothing

                'vInteresVencido = CalculaInteresMoratorio(!nSaldoCap, !nTasaIntVenc, vDiasAtraso)
                'vInteresVencido = Round(vInteresVencido, 2)
                'vCostoCustodia = CalculaCostoCustodiaMoratorio(!nvaltasac, pTasaCustodiaVencida, vDiasAtraso)
                'vCostoCustodia = Round(vCostoCustodia, 2)
                'vCostoPreparacionRemate = pTasaPreparacionRemate * !nvaltasac
                'vCostoPreparacionRemate = Round(vCostoPreparacionRemate, 2)
                'vImpuesto = (vInteresVencido + vCostoCustodia + vCostoPreparacionRemate) * pTasaImpuesto
                'vImpuesto = Round(vImpuesto, 2)
                'lnPreDeuda = CalculaDeudaPrendario(!nSaldoCap, !dFecVenc, !nvaltasac, !nTasaIntVenc, _
                          !nCostCusto, pTasaImpuesto, "6", pTasaPreparacionRemate, txtFecRemate)
                'lnPreDeuda = Round(lnPreDeuda, 2)
                'lnPreVenta = IIf((lnPreDeuda * pFactorPrecioBaseRemate) > vPreBas, CalculaPrecioVentaRemate(lnPreDeuda * pFactorPrecioBaseRemate), CalculaPrecioVentaRemate(vPreBas))
                'lnPreVenta = Round(lnPreVenta, 2)

                lnLineas = lnLineas + 1
                    
                    
''*** PEAC 20090226 ******************************************************
'                vSumGramos = vSumGramos + lnGramos
'                vSumTasaci = vSumTasaci + !nTasacion
'                vSumPresta = vSumPresta + !nMontoCol
'                vSumSalCap = vSumSalCap + !nSaldo
'                vSumIntVen = vSumIntVen + vInteresVencido
'                vSumCosCus = vSumCosCus + vCostoCustodiaMoratorio
'                vSumImpues = vSumImpues + vImpuesto
'                vSumCoPrRe = vSumCoPrRe + vCostoPreparacionRemate
'                vSumPreDeu = vSumPreDeu + lnDeuda
'                vSumPreVen = vSumPreVen + lnPreVenta
                                        
                vSumGramos = vSumGramos + lnGramos
                vSumTasaci = vSumTasaci + !nTasacion
                vSumPresta = vSumPresta + !nMontoCol
                vSumSalCap = vSumSalCap + !nSaldo
                vSumIntCom = vSumIntCom + lnIntAdelantado
                vSumIntVen = vSumIntVen + lnIntVencido
                vSumIntMor = vSumIntMor + lnIntMoratorio
                vSumNotifi = vSumNotifi + !nNotificacion
                'vSumCosCus = vSumCosCus + vCostoCustodiaMoratorio
                vSumCosCus = vSumCosCus + vCostoCustodiaMoratorio + lnValorCosCustoDif '*** PEAC 20190719
                vSumPreDeu = vSumPreDeu + lnDeuda
''*************************************************************************
                                        
                ''*** PEAC 20090226
'                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 4, 0) & Space(1) & !cctacod & Space(1) & !CuentaAntigua & Space(1) & Format(!dVenc, "dd/mm/yyyy") _
                    & ofun.ImpreFormat(ofun.ImpreCarEsp(lsNombre), 20, 1) _
                    & ofun.ImpreFormat(lnGramos, 6) & ofun.ImpreFormat(!nTasacion, 8) & ofun.ImpreFormat(!nMontoCol, 8) _
                    & ofun.ImpreFormat(!nSaldo, 8) & ofun.ImpreFormat(vInteresVencido, 6) & ofun.ImpreFormat(vCostoCustodiaMoratorio, 5) _
                    & ofun.ImpreFormat(vImpuesto, 5) & ofun.ImpreFormat(vCostoPreparacionRemate, 6) & ofun.ImpreFormat(lnDeuda, 8) _
                    & ofun.ImpreFormat(lnPreVenta, 8) & Space(2) & !cAgeBovedA _
                    & oFunI.gPrnSaltoLinea
                    
                ''*** PEAC 20090226
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 4, 0) & Space(1) & !cCtaCod & Space(1) & !CuentaAntigua & Space(1) & Format(!dVenc, "dd/mm/yyyy") _
                    & ofun.ImpreFormat(ofun.ImpreCarEsp(lsNombre), 20, 1) _
                    & ofun.ImpreFormat(lnGramos, 6) & ofun.ImpreFormat(!nTasacion, 8) & ofun.ImpreFormat(!nMontoCol, 8) _
                    & ofun.ImpreFormat(!nSaldo, 8) _
                    & ofun.ImpreFormat(lnIntAdelantado, 6) _
                    & ofun.ImpreFormat(lnIntVencido, 6) _
                    & ofun.ImpreFormat(lnIntMoratorio, 6) _
                    & ofun.ImpreFormat(!nNotificacion, 6) _
                    & ofun.ImpreFormat(vCostoCustodiaMoratorio, 5) _
                    & ofun.ImpreFormat(lnDeuda, 8) _
                    & Space(2) & !cAgeBovedA _
                    & oFunI.gPrnSaltoLinea
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                        
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "PLANILLA DE CONTRATOS ADJUDICAR " & psNroRemate, "Del " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 155, "", 1)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        
        ''*** PEAC 20090226
'        lsCadImp = lsCadImp & Space(32) & "RESUMEN " & ofun.ImpreFormat(vSumGramos, 41) & ofun.ImpreFormat(vSumTasaci, 8) _
'                & ofun.ImpreFormat(vSumPresta, 8) & ofun.ImpreFormat(vSumSalCap, 8) & ofun.ImpreFormat(vSumIntVen, 6) _
'                & ofun.ImpreFormat(vSumCosCus, 5) & ofun.ImpreFormat(vSumImpues, 5) & ofun.ImpreFormat(vSumCoPrRe, 6) _
'                & ofun.ImpreFormat(vSumPreDeu, 8) & ofun.ImpreFormat(vSumPreVen, 8) _
'                & oFunI.gPrnSaltoLinea
                    
        lsCadImp = lsCadImp & Space(32) & "RESUMEN " & ofun.ImpreFormat(vSumGramos, 41) & ofun.ImpreFormat(vSumTasaci, 8) _
                & ofun.ImpreFormat(vSumPresta, 8) & ofun.ImpreFormat(vSumSalCap, 8) & ofun.ImpreFormat(vSumIntCom, 6) _
                & ofun.ImpreFormat(vSumIntVen, 6) & ofun.ImpreFormat(vSumIntMor, 6) & ofun.ImpreFormat(vSumNotifi, 6) _
                & ofun.ImpreFormat(vSumCosCus, 5) & ofun.ImpreFormat(vSumPreDeu, 8) & oFunI.gPrnSaltoLinea
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimePlanillaParaAdjudicar = lsCadBuffer
    Set oFunI = Nothing
End Function

'*** PEAC 20080708
Public Function nImprimeListadoParaAdjudicar(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal pnDiasVctoParaRemate As Integer, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, ByVal pnFactorPrecioVenta As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroRemate As String, _
        ByVal pnJoyasDetalle As Boolean, ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
Dim fnVarIGV As Double



Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim loCalculos As NCOMColPCalculos

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double
    
Dim lrCredPigJoyasDet As New ADODB.Recordset
Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
Dim item As Integer, itemTMP As Integer
'Dim lmDetalle(20) As String
Dim lmDetalle() As String '***PEAC 20190605 ANTES ERA (20)

Dim ofun As New COMFunciones.FCOMImpresion

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

    '*****************************************************
    'Obtener Fecha de Corte
    Dim oCon As New COMConecta.DCOMConecta
    Dim rs1 As ADODB.Recordset
    Set oCon = New COMConecta.DCOMConecta
        If oCon.AbreConexion Then
           Set rs1 = New ADODB.Recordset
           lsSQL = "Select cConsDescripcion From Constante Where nConsCod='3208' and nConsValor='" & Trim(psAgencia) & "' "
           Set rs1 = oCon.CargaRecordSet(lsSQL)
           If Not (rs1.EOF And rs1.BOF) Then
             psFecRemate = Format(rs1("cConsDescripcion"), "mm/dd/yyyy")
           End If
           Set rs1 = Nothing
        End If
    Set oCon = Nothing
    '*******************************************************

    ldFecAviso = DateAdd("d", -pnDiasVctoParaRemate, psFecRemate)

    'lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote,   " _
          & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , " _
          & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , " _
          & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , " _
          & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , " _
          & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  " _
          & "               WHERE cLineaCred = C.cLineaCred  " _
          & "               AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " ),    " _
          & " CuentaAntigua=ISNULL(cCtaCodAnt,'                  ')" _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " left JOIN RelCuentas R ON P.cCtaCod=R.cCtacod " _
          & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod  " _
          & " WHERE P.nPrdEstado in (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ") " _
          & " AND Substring(P.cCtaCod,4,2) = '" & psAgencia & "' AND DATEDIFF(dd, C.dVenc, '" & Format(ldFecAviso, "mm/dd/yyyy") & "') >= 0  " _
          & " GROUP BY P.cCtaCod,cCtaCodAnt, P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote  " _
          & " ORDER BY P.cCtaCod "

    lsSQL = " stp_sel_ListadoParaAdjudicar '" & psAgencia & "','" & psNroRemate & "'"

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Then
        psmensaje = " No existen Créditos Pignoraticios para Aviso de Vencimiento "
        'MuestraImpresion = False

        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "LISTADO DE CONTRATOS PARA ADJUDICAR NRO " & psNroRemate, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 155, "", 2)

        lnIndice = 1:        lnLineas = 7
        

        'Set loCalculos = New COMDColocPig.DCOMColPCalculos
        
      
        
                
        fnVarIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)

        Dim pnDeudaCosto As Double, pn15DeVenta As Double

        
        With lrDataRep
            Do While Not lrDataRep.EOF
                
                v14 = !nK14: v14 = Round(v14, 2)
                v16 = !nK16: v16 = Round(v16, 2)
                v18 = !nK18: v18 = Round(v18, 2)
                v21 = !nK21: v21 = Round(v21, 2)
                vCuenta = 0
                
                pn15DeVenta = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * 0.15
              
                lnPreBase = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) ' * (fnVarIGV + 1)
                
                'lnPreBase = (v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)
                
                lnGramos = v14 + v16 + v18 + v21
                
                lnPreBase = Round(lnPreBase, 2)
                
                vdiasAtraso = DateDiff("d", !dVenc, psFecRemate)
                
                Set loCalculos = New NCOMColPCalculos
                    If vdiasAtraso <= 0 Then
                        vdiasAtraso = 0
                        vInteresVencido = 0
                        vCostoCustodiaMoratorio = 0
                        vImpuesto = 0
                    Else
                        vInteresVencido = loCalculos.nCalculaInteresMoratorio(!nSaldo, !nTasaIntVenc, vdiasAtraso)
                        vInteresVencido = Round(vInteresVencido, 2)
                        vCostoCustodiaMoratorio = loCalculos.nCalculaCostoCustodiaMoratorio(!nTasacion, pnTasaCustodiaVencida, vdiasAtraso)
                        vCostoCustodiaMoratorio = Round(vCostoCustodiaMoratorio, 2)
                        vImpuesto = (vInteresVencido + vCostoCustodiaMoratorio) * pnTasaImpuesto
                        vImpuesto = Round(vImpuesto, 2)
                    End If
                    vCostoPreparacionRemate = 0
                    If !nPrdEstado = gColPEstPRema Then    ' Si esta en via de Remate
                        vCostoPreparacionRemate = pnTasaPreparacionRemate * !nTasacion
                        vCostoPreparacionRemate = Round(vCostoPreparacionRemate, 2)
                    End If
                    lnDeuda = !nSaldo + vInteresVencido + vCostoCustodiaMoratorio + vImpuesto + vCostoPreparacionRemate
                    
                    
                    pnDeudaCosto = Round((lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta), 2)
                    ' pnDeudaCosto = Round((lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta) * (fnVarIGV + 1), 2)
                    'Calcula el Precio Venta para Remate
                    lnPreVenta = loCalculos.nCalculaPrecioVentaRemate(pnDeudaCosto, lnPreBase, pnFactorPrecioVenta)
                    lnPreVenta = CalculaPrecioVentaRemate(Round(lnPreVenta, 2))
                Set loCalculos = Nothing
                
                'Joyas Detalladas
                lnLineas = lnLineas + 1
                If pnJoyasDetalle = True Then

                    Set loMuestraContrato = New COMDColocPig.DCOMColPContrato
                        Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(!cCtaCod)
                    Set loMuestraContrato = Nothing
                    
                    ReDim lmDetalle(lrCredPigJoyasDet.RecordCount) '*** PEAC 20190605
                    
                    item = 0
                    Do While Not lrCredPigJoyasDet.EOF
                        lmDetalle(item) = ofun.ImpreFormat(lrCredPigJoyasDet!nPiezas, 4, 0) & " " & ofun.ImpreFormat(lrCredPigJoyasDet!cDescrip, 27) & " " _
                                    & lrCredPigJoyasDet!cKilataje & "K " & ofun.ImpreFormat(lrCredPigJoyasDet!npesoneto, 4, 2) & "Gr"
                        item = item + 1
                        lrCredPigJoyasDet.MoveNext
                    Loop
                    
                    ''*** PEAC 20090227
                    'lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & Mid(!cctacod, 1, 5) & "-" & Mid(!cctacod, 6) & Space(1) & !CuentaAntigua & Space(1) & ofun.ImpreFormat(!npiezas, 3, 0) & Space(1) _
                                        & lmDetalle(0) & ofun.ImpreFormat(v14, 6) & ofun.ImpreFormat(v16, 6) & ofun.ImpreFormat(v18, 6) & ofun.ImpreFormat(v21, 6) _
                                        & ofun.ImpreFormat(lnPreVenta, 8) & oFunI.gPrnSaltoLinea
                    
                    '*** PEAC 20090227 - SE QUITO lnPreVenta YA QUE ESTO EST REFERIDO A REMATE EL CUAL YA NO EXISTE
                    lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & Mid(!cCtaCod, 1, 5) & "-" & Mid(!cCtaCod, 6) & Space(1) & !CuentaAntigua & Space(1) & ofun.ImpreFormat(!nPiezas, 3, 0) & Space(1) _
                                        & lmDetalle(0) & ofun.ImpreFormat(v14, 6) & ofun.ImpreFormat(v16, 6) & ofun.ImpreFormat(v18, 6) & ofun.ImpreFormat(v21, 6) _
                                        & oFunI.gPrnSaltoLinea
                    
                    For itemTMP = 1 To item - 1
                            lsCadImp = lsCadImp & Space(49) & lmDetalle(itemTMP) & oFunI.gPrnSaltoLinea
                            lnLineas = lnLineas + 1
                    Next
                    

                Else ' **
                
                    ''*** PEAC 20090227
                    'lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !cctacod & Space(1) & !CuentaAntigua & Space(1) & ofun.ImpreFormat(!npiezas, 3, 0) & Space(1) _
                                        & ofun.ImpreFormat(ofun.QuiebreTexto(!cLote, 1), 42, 1) & ofun.ImpreFormat(v14, 6) & ofun.ImpreFormat(v16, 6) & ofun.ImpreFormat(v18, 6) & ofun.ImpreFormat(v21, 6) _
                                        & ofun.ImpreFormat(lnPreVenta, 8) & oFunI.gPrnSaltoLinea
                                        
                    '*** PEAC 20090227 - SE QUITO lnPreVenta YA QUE ESTO EST REFERIDO A REMATE EL CUAL YA NO EXISTE
                    lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !cCtaCod & Space(1) & !CuentaAntigua & Space(1) & ofun.ImpreFormat(!nPiezas, 3, 0) & Space(1) _
                                        & ofun.ImpreFormat(ofun.QuiebreTexto(!cLote, 1), 42, 1) & ofun.ImpreFormat(v14, 6) & ofun.ImpreFormat(v16, 6) & ofun.ImpreFormat(v18, 6) & ofun.ImpreFormat(v21, 6) _
                                        & oFunI.gPrnSaltoLinea
                    
                    If lnIndice Mod 300 = 0 Then
                        lsCadBuffer = lsCadBuffer & lsCadImp
                        lsCadImp = ""
                    End If
                                            
    
                    Do While vCuenta < 15
                        vCuenta = vCuenta + 1
                        If Len(ofun.QuiebreTexto(!cLote, vCuenta + 1)) > 0 Then
                            lsCadImp = lsCadImp & Space(30) & ofun.QuiebreTexto(!cLote, vCuenta + 1) & oFunI.gPrnSaltoLinea
                            lnLineas = lnLineas + 1
                        End If
                    Loop
                
                End If
                lnLineas = lnLineas + 1
                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
                
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "LISTADO DE CONTRATOS PARA ADJUDICAR NRO " & psNroRemate, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 125, "", 2)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(155, "*") & oFunI.gPrnSaltoLinea
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimeListadoParaAdjudicar = lsCadBuffer
    Set oFunI = Nothing
End Function

'*** PEAC 20090313
Public Function nObtieneCtaSobranteAdjudicado(ByVal psAdjudicaCadaAgencia As String, ByRef psmensaje As String) As String
Dim loConstSis As COMDConstSistema.NCOMConstSistema
Dim PObjConec As COMDConstantes.DCOMConstantes
Dim lrCtaAho As New ADODB.Recordset
Dim lsCtaSobranteAdjudica As String
Dim lsnConscod As ConstanteCabecera
    ' Obtiene la cuenta de Ahorros de Sobrante de Adjudicado  ****
    If Right(psAdjudicaCadaAgencia, 2) = "00" Then
        Set loConstSis = New COMDConstSistema.NCOMConstSistema
            lsCtaSobranteAdjudica = loConstSis.LeeConstSistema(61)  ' Cuenta SobranteAdjudica
        Set loConstSis = Nothing
    Else ' Adjudica en Cada Agencia
        Set PObjConec = New COMDConstantes.DCOMConstantes
        lsnConscod = "3207"
        lsCtaSobranteAdjudica = PObjConec.DameDescripcionConstante(lsnConscod, Right(psAdjudicaCadaAgencia, 2))
        If lsCtaSobranteAdjudica = "" Then
           psmensaje = "No se han configurado las Ctas de Ahorro de Sobrantes"
           lsCtaSobranteAdjudica = ""
           Exit Function
        End If
    End If
    nObtieneCtaSobranteAdjudicado = lsCtaSobranteAdjudica
End Function

''*** PEAC 20090521
'RecuperaDiferidosPorLote
Public Function nObtieneDiferidosPorLote(ByVal psAgencias As String) As ADODB.Recordset

Dim loValida As COMDColocPig.DCOMColPFunciones
Dim lrs As New ADODB.Recordset
Dim lsSQL As String
Dim lsAgencia As String

On Error GoTo dError
    lsSQL = "exec stp_sel_DetalleCredDiferidosPorLote '" & psAgencias & "'"
    
    Set loValida = New COMDColocPig.DCOMColPFunciones
        Set lrs = loValida.dObtieneRecordSet(lsSQL)
    Set loValida = Nothing
    Set nObtieneDiferidosPorLote = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, " <<nObtieneDiferidosPorLote>> ", Err.Description
    
End Function


''*** PEAC 20090521
Public Function nObtieneAdjudicadosPorLote(ByVal psAgencias As String, ByVal pdFecIni As Date, ByVal pdFecFin As Date) As ADODB.Recordset

Dim loValida As COMDColocPig.DCOMColPFunciones
Dim lrs As New ADODB.Recordset
Dim lsSQL As String
Dim lsAgencia As String

On Error GoTo dError
    lsSQL = "exec stp_sel_DetalleCredAdjudicadosPorLote '" & psAgencias & "','" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "'"
    
    Set loValida = New COMDColocPig.DCOMColPFunciones
        Set lrs = loValida.dObtieneRecordSet(lsSQL)
    Set loValida = Nothing
    Set nObtieneAdjudicadosPorLote = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, " <<nObtieneAdjudicadosPorLote>> ", Err.Description
End Function

''*** PEAC 20090521
Public Function nObtieneVigentesPorLote(ByVal psAgencias As String, ByVal pdFecIni As Date, ByVal pdFecFin As Date) As ADODB.Recordset

Dim loValida As COMDColocPig.DCOMColPFunciones
Dim lrs As New ADODB.Recordset
Dim lsSQL As String
Dim lsAgencia As String

On Error GoTo dError
    lsSQL = "exec stp_sel_DetalleCredVigentesPorLote '" & psAgencias & "','" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "'"
                                                                            '  Format(pdIni, "yyyymmdd")
    Set loValida = New COMDColocPig.DCOMColPFunciones
        Set lrs = loValida.dObtieneRecordSet(lsSQL)
    Set loValida = Nothing
    Set nObtieneVigentesPorLote = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, " <<nObtieneVigentesPorLote>> ", Err.Description
End Function
