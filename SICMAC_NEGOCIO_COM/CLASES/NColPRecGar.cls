VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NColPRecGar"
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
            Optional ByVal pnTasaImpuesto As Currency = -1, Optional ByVal pnFactorPrecioVenta As Currency = -1, Optional bActualiza As Boolean = False)

'** Actualiza ColocPigRecGar
'** Actualiza ColocPigRGPrecio
'************************************

Dim lsSQL As String
Dim loRegPig As DColPActualizaBD

'On Error GoTo ErrorModPig
Set loRegPig = New DColPActualizaBD
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

Public Function nObtieneNroUltimoProceso(ByVal psTpoProceso As String, ByVal psAgencia As String) As String
'Obtiene Nro Ultimo Remate Credito Pignoraticio
Dim loValida As dColPFunciones
Dim lrs As ADODB.Recordset
Dim lsUltRem As String
Dim lsSQL As String
Dim lsAgencia As String

On Error GoTo dError

lsSQL = "SELECT RG.cTpoProceso, RG.cNroProceso " _
    & " FROM ColocPigRecGar RG  WHERE cTpoProceso ='" & psTpoProceso _
    & "' AND cCodAge ='" & psAgencia & "' ORDER BY cNroProceso DESC "

'Set lrs = New Recordset
    Set loValida = New dColPFunciones
        Set lrs = loValida.dObtieneRecordSet(lsSQL)
    Set loValida = Nothing
    If lrs.BOF And lrs.EOF Then
        MsgBox "No se tiene Información del Ultimo Proceso", vbInformation, "Aviso"
    Else
        lsUltRem = lrs!cNroProceso
    End If
Set lrs = Nothing

nObtieneNroUltimoProceso = lsUltRem
Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function

Public Function nObtieneNroProcesoFecha(ByVal psTpoProceso As String, ByVal psAgencia As String, ByVal psFecha As String) As String
'Obtiene Nro Ultimo Remate Credito Pignoraticio
Dim loValida As dColPFunciones
Dim lrs As ADODB.Recordset
Dim lsUltRem As String
Dim lsSQL As String
Dim lsAgencia As String

On Error GoTo dError

lsSQL = "SELECT RG.cTpoProceso, RG.cNroProceso " _
    & " FROM ColocPigRecGar RG  WHERE cTpoProceso ='" & psTpoProceso _
    & "' AND cCodAge ='" & psAgencia & "' ORDER BY cNroProceso DESC "

'Set lrs = New Recordset
    Set loValida = New dColPFunciones
        Set lrs = loValida.dObtieneRecordSet(lsSQL)
    Set loValida = Nothing
    If lrs.BOF And lrs.EOF Then
        MsgBox "No se tiene Información del Ultimo Proceso", vbInformation, "Aviso"
    Else
        lsUltRem = lrs!cNroProceso
    End If
Set lrs = Nothing

nObtieneNroProcesoFecha = lsUltRem
Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
    
End Function

Public Function nObtieneDatosProcesoRGCredPig(ByVal psTpoProceso As String, ByVal psNroProceso As String, ByVal psAgencia As String) As Recordset
'Obtiene Datos de Remate Credito Pignoraticio
Dim loValida As dColPFunciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String
Dim lsAgencia As String

On Error GoTo dError
    If psTpoProceso = "A" Then
        'lsSQL = "SELECT RG.cTpoProceso, RG.cNroProceso, RG.dProceso, RG.nRGEstado, RG.dParaAdjudicar, " _
            & " sum( case when cTpoPrecio ='PO' then nPrecio else 0  end ) as  nPrecioOro ," _
            & " sum( case when cTpoPrecio ='TC' then nPrecio else 0  end ) as  nTipoCambio " _
            & " FROM ColocPigRecGar RG LEFT JOIN ColocPigRGPrecio RGPrec " _
            & " ON RG.cTpoProceso = RGPrec.cTpoProceso AND RG.cNroProceso = RGPrec.cNroProceso  " _
            & " WHERE RG.cTpoProceso ='" & psTpoProceso & "' and RG.cNroProceso ='" & psNroProceso & "' AND RG.cCodAge ='" & psAgencia & "'  " _
            & " GROUP BY RG.cTpoProceso, RG.cNroProceso, RG.dProceso, RG.nRGEstado, RG.dParaAdjudicar "
        lsSQL = "SELECT RG.cTpoProceso, RG.cNroProceso, RG.dProceso, RG.nRGEstado, RG.dParaAdjudicar, " _
            & " sum( case when cTpoPrecio ='14' then nPrecio else 0  end ) as  nPrecioK14 ," _
            & " sum( case when cTpoPrecio ='16' then nPrecio else 0  end ) as  nPrecioK16 ," _
            & " sum( case when cTpoPrecio ='18' then nPrecio else 0  end ) as  nPrecioK18 ," _
            & " sum( case when cTpoPrecio ='21' then nPrecio else 0  end ) as  nPrecioK21 " _
            & " FROM ColocPigRecGar RG LEFT JOIN ColocPigRGPrecio RGPrec " _
            & " ON RG.cTpoProceso = RGPrec.cTpoProceso AND RG.cNroProceso = RGPrec.cNroProceso  " _
            & " WHERE RG.cTpoProceso ='" & psTpoProceso & "' and RG.cNroProceso ='" & psNroProceso & "' AND RG.cCodAge ='" & psAgencia & "' " _
            & " GROUP BY RG.cTpoProceso, RG.cNroProceso, RG.dProceso, RG.nRGEstado, RG.dParaAdjudicar "
    Else
        lsSQL = "SELECT RG.cTpoProceso, RG.cNroProceso, RG.dProceso, RG.nRGEstado, " _
            & " sum( case when cTpoPrecio ='14' then nPrecio else 0  end ) as  nPrecioK14 ," _
            & " sum( case when cTpoPrecio ='16' then nPrecio else 0  end ) as  nPrecioK16 ," _
            & " sum( case when cTpoPrecio ='18' then nPrecio else 0  end ) as  nPrecioK18 ," _
            & " sum( case when cTpoPrecio ='21' then nPrecio else 0  end ) as  nPrecioK21 " _
            & " FROM ColocPigRecGar RG LEFT JOIN ColocPigRGPrecio RGPrec " _
            & " ON RG.cTpoProceso = RGPrec.cTpoProceso AND RG.cNroProceso = RGPrec.cNroProceso  " _
            & " WHERE RG.cTpoProceso ='" & psTpoProceso & "' and RG.cNroProceso ='" & psNroProceso & "' AND RG.cCodAge ='" & psAgencia & "' " _
            & " GROUP BY RG.cTpoProceso, RG.cNroProceso, RG.dProceso, RG.nRGEstado "
    End If
    Set loValida = New dColPFunciones
        Set lrs = loValida.dObtieneRecordSet(lsSQL)
    Set loValida = Nothing
    If lrs.BOF And lrs.EOF Then
        MsgBox "No se encuentra información del Ultimo Proceso", vbInformation, "Aviso"
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
       ByVal psCodUsu As String) As String
Dim LALISTA() As String
Dim lnFilas As Long, lnItem As Long
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCartaModelo As String
Dim loRegPig As DColPActualizaBD
Dim loContFunct As NContFunciones
Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lnMovNro As Long
Dim lsOpeCod As String
lsOpeCod = "121700"

On Error GoTo dError

    ldFecAviso = DateAdd("d", -pnDiasAtrasoCartaVenc, psFecRemate)


    lsSQL = "SELECT Per.cPersNombre, Per.cPersDireccDomicilio, Per.cPersTelefono, Per.cPersDireccUbiGeo, " _
        & " Ubigeo.cUbiGeoDescripcion, C.cCtaCod , C.dVenc, P.nPrdEstado" _
        & " FROM Producto P JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod " _
        & " JOIN ProductoPersona PP ON c.cCtaCod = PP.cCtaCod " _
        & " JOIN Persona Per ON Per.cPersCod = PP.cPersCod " _
        & " LEFT JOIN UbicacionGeografica Ubigeo ON Per.cPersDireccUbiGeo = Ubigeo.cUbiGeoCod " _
        & " WHERE DATEDIFF(dd, C.dVenc, '" & Format(ldFecAviso, "mm/dd/yyyy") & "') >= 0  AND " _
        & " P.nPrdEstado in (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ") " _
        & " And Substring(P.cCtaCod,4,2) ='" & psAgencia & "' ORDER BY Per.cPersNombre "

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing


    If lrDataRep Is Nothing Then
        MsgBox " No existen Créditos Pignoraticios para Aviso de Vencimiento ", vbInformation, " Aviso"
        Exit Function
    ElseIf lrDataRep.BOF And lrDataRep.EOF Then
        MsgBox " No existen Créditos Pignoraticios para Aviso de Vencimiento ", vbInformation, " Aviso"
        Exit Function
    Else
        If pnCostoCorrespondencia > 0 Then
            'Asigno el gasto a los creditos
            Set loContFunct = New NContFunciones
                lsMovNro = loContFunct.GeneraMovNro(psFechaSis, psCodAge, psCodUsu)
            Set loContFunct = Nothing
        
            lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
            Set loRegPig = New DColPActualizaBD
        
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
        lsCartaModelo = Replace(lsCartaModelo, "<<NOMBRE>>", PstaNombre(LALISTA(lnItem, 0), False), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<FECVENCIM>>", LALISTA(lnItem, 4), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<CONTRATO>>", LALISTA(lnItem, 3), , 2, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<FECREMATE>>", Format(CDate(psFecRemate), "dd.MM.yyyy"), , 1, vbTextCompare)
        'RTFTmp = Replace(RTFTmp, "<<HORREMATE>>", vHorRem, , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<DIRECCION>>", LALISTA(lnItem, 1), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<TELEFONO>>", LALISTA(lnItem, 2), , 1, vbTextCompare)
        
        lsCadImp = lsCadImp & lsCartaModelo & Chr(12)
        'prgList.Value = vCont
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
       ByVal psCodUsu As String) As String
Dim LALISTA() As String
Dim lnFilas As Long, lnItem As Long
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCartaModelo As String

Dim loRegPig As DColPActualizaBD

Dim loContFunct As NContFunciones
Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lnMovNro As Long
Dim lsOpeCod As String
lsOpeCod = "121700"

'On Error GoTo dError

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

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing


    If lrDataRep Is Nothing Then
        MsgBox " No existen Créditos Pignoraticios para Aviso de Remate ", vbInformation, " Aviso"
        Exit Function
    ElseIf lrDataRep.BOF And lrDataRep.EOF Then
        MsgBox " No existen Créditos Pignoraticios para Aviso de Remate ", vbInformation, " Aviso"
        Exit Function
    Else
    
        If pnCostoCorrespondencia > 0 Then
            'Asigno el gasto a los creditos
            Set loContFunct = New NContFunciones
                lsMovNro = loContFunct.GeneraMovNro(psFechaSis, psCodAge, psCodUsu)
            Set loContFunct = Nothing
        
            lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
            Set loRegPig = New DColPActualizaBD
        
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
        lsCartaModelo = Replace(lsCartaModelo, "<<NOMBRE>>", PstaNombre(LALISTA(lnItem, 0), False), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<FECVENCIM>>", LALISTA(lnItem, 4), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<CONTRATO>>", LALISTA(lnItem, 3), , 2, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<FECREMATE>>", Format(CDate(psFecRemate), "dd.MM.yyyy"), , 1, vbTextCompare)
        'RTFTmp = Replace(RTFTmp, "<<HORREMATE>>", vHorRem, , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<DIRECCION>>", LALISTA(lnItem, 1), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<TELEFONO>>", LALISTA(lnItem, 2), , 1, vbTextCompare)
        
        lsCadImp = lsCadImp & lsCartaModelo & Chr(12)
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
        ByVal pnNroRepor As Integer) As String
        
Dim lsCadImp As String
Dim loImprimeCab As NColPImpre
    
    Set loImprimeCab = New NColPImpre
        lsCadImp = loImprimeCab.nImprimeCabeceraReportes(psNomCmac, psNomAgencia, psCodUser, psFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)
    Set loImprimeCab = Nothing
    lsCadImp = lsCadImp & Chr(10) & String(pnAnchoLinea, "-") & Chr(10)
    Select Case pnNroRepor
        Case 1 ' Planilla de Remate
            lsCadImp = lsCadImp & Space(1) & "Item    Contrato      Fec.Venc.     Cliente              Gramos    Tasacion   Prestamo  SaldoCap  Interes  Custodia Impuesto Prep.Rem   Deuda   P.Base  Boveda" & Chr(10)
        Case 2 ' Listados para Remate
            lsCadImp = lsCadImp & Space(1) & "Item    Contrato    Piezas     Descripción                                       14Kl.    16Kl.    18Kl.    21Kl.  Prec.Base " & Chr(10)
    
        Case 3 ' Planilla de Prendas Vendidas
            lsCadImp = lsCadImp & Space(1) & "Item    Contrato    Vencimiento     Tasacion   Prestamo   SaldoCap      Deuda    Venta B.   IGV  Comision   Vta Neta   Sobrante OroNeto" & Chr(10)
        Case 4 ' Planilla de Prendas No Vendidas
            lsCadImp = lsCadImp & Space(1) & "Item    ContratoAnt     Contrato     Vencimiento     Tasacion   Prestamo   SaldoCap      Deuda   OroNeto" & Chr(10)
            
        Case 9 ' Planilla de Remate
            lsCadImp = lsCadImp & Space(1) & "Item  ContratoAnt         Contrato          Fec.Venc.   Gramos    Tasacion   Prestamo  SaldoCap  Interes  Custodia Impuesto Prep.Rem   Deuda   P.Base " & Chr(10)
        Case 10 ' Listados para Remate
            lsCadImp = lsCadImp & Space(1) & "Item  ContratoAnt        Contrato            Piezas     Descripción                                    14Kl.    16Kl.    18Kl.    21Kl.  Prec.Base " & Chr(10)
        Case 11 ' Listados para Remate
            lsCadImp = lsCadImp & Space(1) & "Item  ContratoAnt        Contrato            Piezas     Descripción                                  14Kl.    16Kl.    18Kl.    21Kl.  Prec.Base " & Chr(10)
    End Select
    lsCadImp = lsCadImp & String(pnAnchoLinea, "-") & Chr(10)
nImprimeRemateReporteCabec = lsCadImp
End Function

'Private Sub CargaParametros()
'Dim loParam As DColPCalculos
'Set loParam = New DColPCalculos
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
Dim ldFecAviso As Date, RSTEMP As Recordset, lsSQL As String
Dim conex As DConecta
Set conex = New DConecta
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
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroRemate As String) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As DColPCalculos
Set loParam = New DColPCalculos
Dim fnVarIGV As Double


Dim loCalculos As NColPCalculos

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
          & "               ON Per.cPersCod = PPer.cPersCod WHERE PPer.cCtaCod = P.cCtaCod )  " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod  " _
          & " WHERE P.nPrdEstado in (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ") " _
          & " AND Substring(P.cCtaCod,4,2) ='" & psAgencia & "' AND DATEDIFF(dd, C.dVenc, '" & Format(ldFecAviso, "mm/dd/yyyy") & "') >= 0  " _
          & " GROUP BY P.cCtaCod, P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred  " _
          & " ORDER BY P.cCtaCod "

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
 

    If lrDataRep Is Nothing Then
        MsgBox " No existen Créditos Pignoraticios para Aviso de Vencimiento ", vbInformation, " Aviso"
        'MuestraImpresion = False

        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        
        'CABECERA
        lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "PLANILLA DE CONTRATOS PARA REMATE NRO " & psNroRemate, "Del " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 155, "", 1)

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
              
                lnPreBase = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * (fnVarIGV + 1)
                
                lnGramos = v14 + v16 + v18 + v21
                
                lnPreBase = Round(lnPreBase, 2)
                
                lsNombre = PstaNombre(!cNomCliente)
        
                vdiasAtraso = DateDiff("d", !dVenc, psFecRemate)
                
                Set loCalculos = New NColPCalculos
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
                    
                    pnDeudaCosto = Round((lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta) * (fnVarIGV + 1), 2)
                    
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
                                        
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 4, 0) & Space(1) & !cCtaCod & Space(1) & Format(!dVenc, "dd/mm/yyyy") _
                    & ImpreFormat(ImpreCarEsp(lsNombre), 20, 1) _
                    & ImpreFormat(lnGramos, 6) & ImpreFormat(!nTasacion, 8) & ImpreFormat(!nMontoCol, 8) _
                    & ImpreFormat(!nSaldo, 8) & ImpreFormat(vInteresVencido, 6) & ImpreFormat(vCostoCustodiaMoratorio, 5) _
                    & ImpreFormat(vImpuesto, 5) & ImpreFormat(vCostoPreparacionRemate, 6) & ImpreFormat(lnDeuda, 8) _
                    & ImpreFormat(lnPreVenta, 8) & Space(2) & !cAgeBoveda _
                    & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                        
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "PLANILLA DE CONTRATOS PARA REMATE NRO " & psNroRemate, "Del " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 155, "", 1)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & Chr(10)
        lsCadImp = lsCadImp & Space(12) & "RESUMEN " & ImpreFormat(vSumGramos, 41) & ImpreFormat(vSumTasaci, 8) _
                & ImpreFormat(vSumPresta, 8) & ImpreFormat(vSumSalCap, 8) & ImpreFormat(vSumIntVen, 6) _
                & ImpreFormat(vSumCosCus, 5) & ImpreFormat(vSumImpues, 5) & ImpreFormat(vSumCoPrRe, 6) _
                & ImpreFormat(vSumPreDeu, 8) & ImpreFormat(vSumPreVen, 8) _
                & Chr(10)
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimePlanillaParaRemate = lsCadBuffer
End Function

Public Function nImprimePlanillaParaRemateConSiaf(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal pnDiasVctoParaRemate As Integer, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, ByVal pnFactorPrecioVenta As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroRemate As String) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As DColPCalculos
Set loParam = New DColPCalculos
Dim fnVarIGV As Double


Dim loCalculos As NColPCalculos

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

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
 

    If lrDataRep Is Nothing Then
        MsgBox " No existen Créditos Pignoraticios para Aviso de Vencimiento ", vbInformation, " Aviso"
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
              
                lnPreBase = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * (fnVarIGV + 1)
                
                lnGramos = v14 + v16 + v18 + v21
                
                lnPreBase = Round(lnPreBase, 2)
                
                lsNombre = PstaNombre(!cNomCliente)
        
                vdiasAtraso = DateDiff("d", !dVenc, psFecRemate)
                
                Set loCalculos = New NColPCalculos
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
                    
                    pnDeudaCosto = Round((lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta) * (fnVarIGV + 1), 2)
                    
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
                                        
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 4, 0) & Space(1) & !codesiaf & Space(1) & !cCtaCod & Space(1) & Format(!dVenc, "dd/mm/yyyy") _
                           & ImpreFormat(lnGramos, 6) & ImpreFormat(!nTasacion, 8) & ImpreFormat(!nMontoCol, 8) _
                           & ImpreFormat(!nSaldo, 8) & ImpreFormat(vInteresVencido, 6) & ImpreFormat(vCostoCustodiaMoratorio, 5) _
                           & ImpreFormat(vImpuesto, 5) & ImpreFormat(vCostoPreparacionRemate, 6) & ImpreFormat(lnDeuda, 8) _
                           & ImpreFormat(lnPreVenta, 8) & Chr(10)  '& Space(2) & !cAgeBoveda _

                    '& ImpreFormat(ImpreCarEsp(lsNombre), 20, 1) _

                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                        
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "PLANILLA DE CONTRATOS PARA REMATE NRO " & psNroRemate, "Del " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 155, "", 9)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & Chr(10)
        lsCadImp = lsCadImp & Space(10) & "RESUMEN " & ImpreFormat(vSumGramos, 41) & ImpreFormat(vSumTasaci, 8) _
                  & ImpreFormat(vSumPresta, 8) & ImpreFormat(vSumSalCap, 8) & ImpreFormat(vSumIntVen, 6) _
                  & ImpreFormat(vSumCosCus, 5) & ImpreFormat(vSumImpues, 5) & ImpreFormat(vSumCoPrRe, 6) _
                  & ImpreFormat(vSumPreDeu, 8) & ImpreFormat(vSumPreVen, 8) & Chr(10)
                  '--& ImpreFormat(vSumPreVen, 8) & Chr(10)
                  '& Chr(10)
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimePlanillaParaRemateConSiaf = lsCadBuffer
End Function


'Imprime Listado para Remate
Public Function nImprimeListadoParaRemate(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal pnDiasVctoParaRemate As Integer, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, ByVal pnFactorPrecioVenta As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroRemate As String, _
        ByVal pnJoyasDetalle As Boolean) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As DColPCalculos
Set loParam = New DColPCalculos
Dim fnVarIGV As Double



Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim loCalculos As NColPCalculos

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double
    
Dim lrCredPigJoyasDet As ADODB.Recordset
Dim loMuestraContrato As DColPContrato
Dim item As Integer, itemTMP As Integer
Dim lmDetalle(20) As String



    ldFecAviso = DateAdd("d", -pnDiasVctoParaRemate, psFecRemate)

    lsSQL = " SELECT P.cCtaCod, P.nPrdEstado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote,   " _
          & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , " _
          & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , " _
          & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , " _
          & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , " _
          & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  " _
          & "               WHERE cLineaCred = C.cLineaCred  " _
          & "               AND nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " )    " _
          & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
          & " INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod " _
          & " LEFT JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod  " _
          & " WHERE P.nPrdEstado in (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & ") " _
          & " AND Substring(P.cCtaCod,4,2) = '" & psAgencia & "' AND DATEDIFF(dd, C.dVenc, '" & Format(ldFecAviso, "mm/dd/yyyy") & "') >= 0  " _
          & " GROUP BY P.cCtaCod, P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
          & " CP.cAgeBoveda, C.cLineaCred, CP.nPiezas, CP.cLote  " _
          & " ORDER BY P.cCtaCod "

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing





    If lrDataRep Is Nothing Then
        MsgBox " No existen Créditos Pignoraticios para Aviso de Vencimiento ", vbInformation, " Aviso"
        'MuestraImpresion = False

        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "LISTADO DE CONTRATOS PARA REMATE NRO " & psNroRemate, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 125, "", 2)

        lnIndice = 1:        lnLineas = 7
        

        'Set loCalculos = New DColPCalculos
        
      
        
                
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
              
                lnPreBase = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * (fnVarIGV + 1)
                
                'lnPreBase = (v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)
                
                
                lnGramos = v14 + v16 + v18 + v21
                
                lnPreBase = Round(lnPreBase, 2)
                
                vdiasAtraso = DateDiff("d", !dVenc, psFecRemate)
                
                Set loCalculos = New NColPCalculos
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
                     pnDeudaCosto = Round((lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta) * (fnVarIGV + 1), 2)
                    'Calcula el Precio Venta para Remate
                    lnPreVenta = loCalculos.nCalculaPrecioVentaRemate(pnDeudaCosto, lnPreBase, pnFactorPrecioVenta)
                    lnPreVenta = Round(lnPreVenta, 2)
                Set loCalculos = Nothing
                
                'Joyas Detalladas
                lnLineas = lnLineas + 1
                If pnJoyasDetalle = True Then

                    Set loMuestraContrato = New DColPContrato
                        Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(!cCtaCod)
                    Set loMuestraContrato = Nothing
                    
                    item = 0
                    Do While Not lrCredPigJoyasDet.EOF
                        lmDetalle(item) = ImpreFormat(lrCredPigJoyasDet!npiezas, 4, 0) & " " & ImpreFormat(lrCredPigJoyasDet!cDescrip, 27) & " " _
                                    & lrCredPigJoyasDet!cKilataje & "K " & ImpreFormat(lrCredPigJoyasDet!npesoneto, 4, 2) & "Gr"
                        item = item + 1
                        lrCredPigJoyasDet.MoveNext
                    Loop
                    
                    lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & Mid(!cCtaCod, 1, 5) & "-" & Mid(!cCtaCod, 6) & Space(1) & ImpreFormat(!npiezas, 3, 0) & Space(1) _
                                        & lmDetalle(0) & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & ImpreFormat(v18, 6) & ImpreFormat(v21, 6) _
                                        & ImpreFormat(lnPreVenta, 8) & Chr(10)
                    
                    For itemTMP = 1 To item - 1
                            lsCadImp = lsCadImp & Space(30) & lmDetalle(itemTMP) & Chr(10)
                            lnLineas = lnLineas + 1
                    Next
                    

                Else ' **
                    lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !cCtaCod & Space(1) & ImpreFormat(!npiezas, 3, 0) & Space(1) _
                                        & ImpreFormat(QuiebreTexto(!cLote, 1), 42, 1) & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & ImpreFormat(v18, 6) & ImpreFormat(v21, 6) _
                                        & ImpreFormat(lnPreVenta, 8) & Chr(10)
                    
                    If lnIndice Mod 300 = 0 Then
                        lsCadBuffer = lsCadBuffer & lsCadImp
                        lsCadImp = ""
                    End If
                                            
    
                    Do While vCuenta < 15
                        vCuenta = vCuenta + 1
                        If Len(QuiebreTexto(!cLote, vCuenta + 1)) > 0 Then
                            lsCadImp = lsCadImp & Space(30) & QuiebreTexto(!cLote, vCuenta + 1) & Chr(10)
                            lnLineas = lnLineas + 1
                        End If
                    Loop
                
                End If
                lnLineas = lnLineas + 1
                lsCadImp = lsCadImp & Chr(10)
                
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "LISTADO DE CONTRATOS PARA REMATE NRO " & psNroRemate, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 125, "", 2)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & Chr(10)
        lsCadImp = lsCadImp & String(120, "*") & Chr(10)
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimeListadoParaRemate = lsCadBuffer
End Function

'Imprime Listado para Remate
Public Function nImprimeListadoParaRemateConSIAF(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal pnDiasVctoParaRemate As Integer, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, ByVal pnFactorPrecioVenta As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroRemate As String, _
        ByVal pnJoyasDetalle As Boolean) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As DColPCalculos
Set loParam = New DColPCalculos
Dim fnVarIGV As Double



Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim loCalculos As NColPCalculos

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double
    
Dim lrCredPigJoyasDet As ADODB.Recordset
Dim loMuestraContrato As DColPContrato
Dim item As Integer, itemTMP As Integer
Dim lmDetalle(20) As String



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

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing





    If lrDataRep Is Nothing Then
        MsgBox " No existen Créditos Pignoraticios para Aviso de Vencimiento ", vbInformation, " Aviso"
        'MuestraImpresion = False

        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "LISTADO DE CONTRATOS PARA REMATE NRO " & psNroRemate, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 150, "", 10)

        lnIndice = 1:        lnLineas = 7
        

        'Set loCalculos = New DColPCalculos
        
      
        
                
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
              
                lnPreBase = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * (fnVarIGV + 1)
                
                'lnPreBase = (v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)
                
                
                lnGramos = v14 + v16 + v18 + v21
                
                lnPreBase = Round(lnPreBase, 2)
                
                vdiasAtraso = DateDiff("d", !dVenc, psFecRemate)
                
                Set loCalculos = New NColPCalculos
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
                     pnDeudaCosto = Round((lnDeuda + (lnDeuda * pnFactorPrecioVenta) + pn15DeVenta) * (fnVarIGV + 1), 2)
                    'Calcula el Precio Venta para Remate
                    lnPreVenta = loCalculos.nCalculaPrecioVentaRemate(pnDeudaCosto, lnPreBase, pnFactorPrecioVenta)
                    lnPreVenta = Round(lnPreVenta, 2)
                Set loCalculos = Nothing
                
                'Joyas Detalladas
                lnLineas = lnLineas + 1
                If pnJoyasDetalle = True Then

                    Set loMuestraContrato = New DColPContrato
                        Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(!cCtaCod)
                    Set loMuestraContrato = Nothing
                    
                    item = 0
                    Do While Not lrCredPigJoyasDet.EOF
                        lmDetalle(item) = ImpreFormat(lrCredPigJoyasDet!npiezas, 3, 0) & "" & ImpreFormat(lrCredPigJoyasDet!cDescrip, 27) & " " _
                                    & lrCredPigJoyasDet!cKilataje & "K " & ImpreFormat(lrCredPigJoyasDet!npesoBRUTO, 3, 2) & "g" & ImpreFormat(lrCredPigJoyasDet!npesoneto, 3, 2) & "g"
                        item = item + 1
                        lrCredPigJoyasDet.MoveNext
                    Loop
                  
                    lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !codesiaf & Space(1) & Mid(!cCtaCod, 1, 5) & "-" & Mid(!cCtaCod, 6) & Space(1) & ImpreFormat(!npiezas, 3, 0) & Space(1) _
                                        & lmDetalle(0) & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & ImpreFormat(v18, 6) & ImpreFormat(v21, 6) _
                                        & ImpreFormat(lnPreVenta, 8) & Chr(10)
                    
                    For itemTMP = 1 To item - 1
                            lsCadImp = lsCadImp & Space(49) & lmDetalle(itemTMP) & Chr(10)
                            lnLineas = lnLineas + 1
                    Next
                    

                Else ' **
                    lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !codesiaf & Space(1) & !cCtaCod & Space(1) & ImpreFormat(!npiezas, 3, 0) & Space(1) _
                                        & ImpreFormat(QuiebreTexto(!cLote, 1), 42, 1) & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & ImpreFormat(v18, 6) & ImpreFormat(v21, 6) _
                                        & ImpreFormat(lnPreVenta, 8) & Chr(10)
                    
                    If lnIndice Mod 300 = 0 Then
                        lsCadBuffer = lsCadBuffer & lsCadImp
                        lsCadImp = ""
                    End If
                                            
    
                    Do While vCuenta < 15
                        vCuenta = vCuenta + 1
                        If Len(QuiebreTexto(!cLote, vCuenta + 1)) > 0 Then
                            lsCadImp = lsCadImp & Space(49) & QuiebreTexto(!cLote, vCuenta + 1) & Chr(10)
                            lnLineas = lnLineas + 1
                        End If
                    Loop
                
                End If
                lnLineas = lnLineas + 1
                lsCadImp = lsCadImp & Chr(10)
                
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "LISTADO DE CONTRATOS PARA REMATE NRO " & psNroRemate, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 150, "", 10)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & Chr(10)
        lsCadImp = lsCadImp & String(150, "*") & Chr(10)
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimeListadoParaRemateConSIAF = lsCadBuffer
End Function

Public Function nImpListadoRemateAnteriores(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal pnDiasVctoParaRemate As Integer, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, ByVal pnFactorPrecioVenta As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroRemate As String, _
        ByVal pnJoyasDetalle As Boolean) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As DColPCalculos
Set loParam = New DColPCalculos
Dim fnVarIGV As Double



Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim loCalculos As NColPCalculos

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double
    
Dim lrCredPigJoyasDet As ADODB.Recordset
Dim loMuestraContrato As DColPContrato
Dim item As Integer, itemTMP As Integer
Dim lmDetalle(20) As String
Dim ldFecAnt As Date


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



    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing





    If lrDataRep Is Nothing Then
        MsgBox " No existen Créditos Pignoraticios para Aviso de Vencimiento ", vbInformation, " Aviso"
        'MuestraImpresion = False

        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "LISTADO DE CONTRATOS PARA REMATE NRO " & psNroRemate, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 145, "", 11)

        lnIndice = 1:        lnLineas = 7
        

        'Set loCalculos = New DColPCalculos
        
      
        
                
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
              
                lnPreBase = ((v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)) * (fnVarIGV + 1)
                
                'lnPreBase = (v14 * pnPreOro14) + (v16 * pnPreOro16) + (v18 * pnPreOro18) + (v21 * pnPreOro21)
                
                
                lnGramos = v14 + v16 + v18 + v21
                
                lnPreBase = Round(lnPreBase, 2)
                
                vdiasAtraso = DateDiff("d", !dVenc, psFecRemate)
                

                lnLineas = lnLineas + 1
                If pnJoyasDetalle = True Then

                    Set loMuestraContrato = New DColPContrato
                        Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(!cCtaCod)
                    Set loMuestraContrato = Nothing
                    
                    item = 0
                    Do While Not lrCredPigJoyasDet.EOF
                        lmDetalle(item) = ImpreFormat(lrCredPigJoyasDet!npiezas, 4, 0) & " " & ImpreFormat(lrCredPigJoyasDet!cDescrip, 27) & " " _
                                    & lrCredPigJoyasDet!cKilataje & "K " & ImpreFormat(lrCredPigJoyasDet!npesoneto, 4, 2) & "Gr"
                        item = item + 1
                        lrCredPigJoyasDet.MoveNext
                    Loop
                  
                    lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & Trim(Left(!codesiaf, 18)) & Space(19 - Len(Trim(Left(!codesiaf, 18)))) & Mid(!cCtaCod, 1, 5) & "-" & Mid(!cCtaCod, 6) & Space(1) & ImpreFormat(!npiezas, 3, 0) & Space(1) _
                                        & lmDetalle(0) & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & ImpreFormat(v18, 6) & ImpreFormat(v21, 6) _
                                        & ImpreFormat(!NREMSUBBASEVTA, 8) & Chr(10)
                    
                    For itemTMP = 1 To item - 1
                            lsCadImp = lsCadImp & Space(49) & lmDetalle(itemTMP) & Chr(10)
                            lnLineas = lnLineas + 1
                    Next
                    

                Else ' **
                    lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & Trim(Left(!cNomCliente, 18)) & Space(19 - Len(Trim(Left(!cNomCliente, 18)))) & !cCtaCod & Space(1) & ImpreFormat(!npiezas, 3, 0) & Space(1) _
                                        & ImpreFormat(QuiebreTexto(!cLote, 1), 42, 1) & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & ImpreFormat(v18, 6) & ImpreFormat(v21, 6) _
                                        & ImpreFormat(lnPreVenta, 8) & Chr(10)
                    
                    If lnIndice Mod 300 = 0 Then
                        lsCadBuffer = lsCadBuffer & lsCadImp
                        lsCadImp = ""
                    End If
                                            
    
                    Do While vCuenta < 15
                        vCuenta = vCuenta + 1
                        If Len(QuiebreTexto(!cLote, vCuenta + 1)) > 0 Then
                            lsCadImp = lsCadImp & Space(49) & QuiebreTexto(!cLote, vCuenta + 1) & Chr(10)
                            lnLineas = lnLineas + 1
                        End If
                    Loop
                
                End If
                lnLineas = lnLineas + 1
                lsCadImp = lsCadImp & Chr(10)
                
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "LISTADO DE CONTRATOS PARA REMATE NRO " & psNroRemate, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 145, "", 11)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & Chr(10)
        lsCadImp = lsCadImp & String(145, "*") & Chr(10)
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImpListadoRemateAnteriores = lsCadBuffer
End Function

Public Function nListadoParaRemConSIAF(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal pnDiasVctoParaRemate As Integer, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, ByVal pnFactorPrecioVenta As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroRemate As String, _
        ByVal pnJoyasDetalle As Boolean) As ADODB.Recordset
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As DColPCalculos
Set loParam = New DColPCalculos
Dim fnVarIGV As Double



Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim loCalculos As NColPCalculos

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double
    
Dim lrCredPigJoyasDet As ADODB.Recordset
Dim loMuestraContrato As DColPContrato
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



    Set loDataRep = New dColPFunciones
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
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal pnJoyasDetalle As Boolean) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As DColPCalculos
Set loParam = New DColPCalculos
Dim fnVarIGV As Double



Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim loCalculos As NColPCalculos

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double
    
Dim lrCredPigJoyasDet As ADODB.Recordset
Dim loMuestraContrato As DColPContrato
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

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing





    If lrDataRep Is Nothing Then
        MsgBox " No existen Créditos Pignoraticios para Aviso de Vencimiento ", vbInformation, " Aviso"
        'MuestraImpresion = False

        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nImprimeAdjudicaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "LISTADO DE CONTRATOS PARA ADJUDICACION NRO " & psNroProcesoAdjud, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 144, "", 11)

        lnIndice = 1:        lnLineas = 7
        

        'Set loCalculos = New DColPCalculos
        
      
        
                
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
'                Set loCalculos = New NColPCalculos
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

                    Set loMuestraContrato = New DColPContrato
                        Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(!cCtaCod)
                    Set loMuestraContrato = Nothing
                    
                    item = 0
                    Do While Not lrCredPigJoyasDet.EOF
                        lmDetalle(item) = ImpreFormat(lrCredPigJoyasDet!npiezas, 4, 0) & " " & ImpreFormat(lrCredPigJoyasDet!cDescrip, 27) & " " _
                                    & lrCredPigJoyasDet!cKilataje & "K " & ImpreFormat(lrCredPigJoyasDet!npesoneto, 4, 2) & "Gr"
                        item = item + 1
                        lrCredPigJoyasDet.MoveNext
                    Loop
                  
                    lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !codesiaf & Space(1) & Mid(!cCtaCod, 1, 5) & "-" & Mid(!cCtaCod, 6) & Space(1) & ImpreFormat(!npiezas, 3, 0) & Space(1) _
                                        & lmDetalle(0) & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & ImpreFormat(v18, 6) & ImpreFormat(v21, 6) _
                                        & ImpreFormat(lnPreVenta, 8) & Chr(10)
                    
                    For itemTMP = 1 To item - 1
                            lsCadImp = lsCadImp & Space(49) & lmDetalle(itemTMP) & Chr(10)
                            lnLineas = lnLineas + 1
                    Next
                    

                Else ' **
                    lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !codesiaf & Space(1) & !cCtaCod & Space(1) & ImpreFormat(!npiezas, 3, 0) & Space(1) _
                                        & ImpreFormat(QuiebreTexto(!cLote, 1), 42, 1) & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & ImpreFormat(v18, 6) & ImpreFormat(v21, 6) _
                                        & ImpreFormat(lnPreVenta, 8) & Chr(10)
                    
                    If lnIndice Mod 300 = 0 Then
                        lsCadBuffer = lsCadBuffer & lsCadImp
                        lsCadImp = ""
                    End If
                                            
    
                    Do While vCuenta < 15
                        vCuenta = vCuenta + 1
                        If Len(QuiebreTexto(!cLote, vCuenta + 1)) > 0 Then
                            lsCadImp = lsCadImp & Space(40) & QuiebreTexto(!cLote, vCuenta + 1) & Chr(10)
                            lnLineas = lnLineas + 1
                        End If
                    Loop
                
                End If
                lnLineas = lnLineas + 1
                lsCadImp = lsCadImp & Chr(10)
                
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nImprimeAdjudicaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "LISTADO DE CONTRATOS PARA ADJUDICACION NRO " & psNroProcesoAdjud, "DEL " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 144, "", 11)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & Chr(10)
        lsCadImp = lsCadImp & String(120, "*") & Chr(10)
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimeListadoParaAdjConSIAF = lsCadBuffer
End Function


Public Function nListadoParaAdjConSIAF(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal psNroProcesoAdjud As String, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal pnJoyasDetalle As Boolean) As ADODB.Recordset
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim loParam As DColPCalculos
Set loParam = New DColPCalculos
Dim fnVarIGV As Double



Dim lnIndice As Long  'contador de Item
Dim lnLineas As Integer
Dim lnPage As Integer
    
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
Dim vCuenta As Integer

Dim loCalculos As NColPCalculos

Dim lnPreBase As Double
Dim lnDeuda As Double
Dim lnPreVenta As Currency
Dim lnGramos As Currency

Dim vInteresVencido As Currency, vCostoCustodiaMoratorio As Currency, vImpuesto As Currency
Dim vCostoPreparacionRemate As Currency
Dim vdiasAtraso As Double
    
Dim lrCredPigJoyasDet As ADODB.Recordset
Dim loMuestraContrato As DColPContrato
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

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
        
        
        Set nListadoParaAdjConSIAF = lrDataRep
        
        
    Set loDataRep = Nothing
    Set lrDataRep = Nothing




End Function




Public Function nIniciaRemate(ByVal psNroRemate As String, ByVal psFecRemate As String, _
        ByVal psProcesoCadaAgencia As String, ByVal pnDiasVctoParaRemate As Integer, ByVal pnFactorPrecioVenta As Double, ByVal pnPrecioOro14 As Double, ByVal pnPrecioOro16 As Double, _
        ByVal pnPrecioOro18 As Double, ByVal pnPrecioOro21 As Double, ByVal pnTasaCustodiaVenc As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double) As Boolean

Dim lsSQL As String
Dim loDataIniRem As dColPFunciones
Dim lrDataInicia As ADODB.Recordset
              
Dim loCalcula As NColPCalculos
Dim loRegPig As DColPActualizaBD
Dim ldFecLimiteRemate As Date

Dim loParam As DColPCalculos
Set loParam = New DColPCalculos
Dim fnVarIGV As Double
Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency


Dim lnPreBase As Currency, lnPreDeuda As Currency, lnPreVenta As Currency

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
    
Set loDataIniRem = New dColPFunciones
    Set lrDataInicia = loDataIniRem.dObtieneRecordSet(lsSQL)
Set loDataIniRem = Nothing

If lrDataInicia Is Nothing Then

Else
    lnPreBase = 0: lnPreDeuda = 0: lnPreVenta = 0
    
        fnVarIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)
        Dim pnDeudaCosto As Double, pn15DeVenta As Double
    
    With lrDataInicia
        Set loRegPig = New DColPActualizaBD
        
        Do While Not .EOF
        
                v14 = !nK14: v14 = Round(v14, 2)
                v16 = !nK16: v16 = Round(v16, 2)
                v18 = !nK18: v18 = Round(v18, 2)
                v21 = !nK21: v21 = Round(v21, 2)
        
        
            'Precio Base
            pn15DeVenta = ((v14 * pnPrecioOro14) + (v16 * pnPrecioOro16) + (v18 * pnPrecioOro18) + (v21 * pnPrecioOro21)) * 0.15
            lnPreBase = ((v14 * pnPrecioOro14) + (v16 * pnPrecioOro16) + (v18 * pnPrecioOro18) + (v21 * pnPrecioOro21)) * (fnVarIGV + 1)

            lnPreBase = Round(lnPreBase, 2)

            '** Actualiza Producto
            Call loRegPig.dUpdateProducto(!cCtaCod, , , gColPEstPRema, , , False)         ' (-2) Cambia el estado a Para Remate

            Set loCalcula = New NColPCalculos
                'Deuda
                lnPreDeuda = loCalcula.nCalculaDeudaPignoraticio(!nSaldo, !dVenc, !nTasacion, !nTasaIntVenc, _
                    pnTasaCustodiaVenc, pnTasaImpuesto, gColPEstPRema, pnTasaPreparacionRemate, Format(psFecRemate, "mm/dd/yyyy"))
                
                lnPreDeuda = Round(lnPreDeuda, 2)
                'pnDeudaCosto = lnPreDeuda + (lnPreDeuda * pnFactorPrecioVenta) + pn15DeVenta
                
                lnPreVenta = loCalcula.nCalculaPrecioVentaRemate(pnDeudaCosto, lnPreBase, pnFactorPrecioVenta)
                lnPreVenta = Round(lnPreVenta, 0)
                
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
    
End If

End Function

Public Function nFinalizaRemate(ByVal psNroRemate As String, ByVal psFecProceso As String, ByVal psCtaAhorroRemate As String, _
        ByVal pnMontoSobrante As Currency, ByVal psProcesoCadaAgencia As String, ByVal psCodUser As String, _
        ByVal psLpt As String) As Boolean

Dim lsSQL As String
Dim lsNextRemate As String
              
Dim loRegPig As DColPActualizaBD
Dim oBase As DCredActualizaBD
Dim oCont As NContFunciones

    'Genera el sgte Nro de Remate
    lsNextRemate = String(4 - Len(Trim(Str(Val(psNroRemate) + 1))), "0") & Trim(Str(Val(psNroRemate) + 1))
    
    Set loRegPig = New DColPActualizaBD
    
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
        Dim oCapMov As NCapMovimientos
        Dim sMovNroReg As String
        Dim nMovNroReg As Long
        Set oCont = New NContFunciones
            sMovNroReg = oCont.GeneraMovNro(psFecProceso, psProcesoCadaAgencia, psCodUser)
        Set oCont = Nothing
        Set oCapMov = New NCapMovimientos
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
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroRemate As String) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim loCalculos As NColPCalculos

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

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing


    If lrDataRep Is Nothing Then
        MsgBox " No existen Créditos Pignoraticios para Aviso de Vencimiento ", vbInformation, " Aviso"
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
                vSumOroNeto = vSumOroNeto + !noroneto
                                        
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 4, 0) & Space(1) & !cCtaCod & Space(1) & Format(!dVenc, "dd/mm/yyyy") _
                    & ImpreFormat(!nTasacion, 8) & ImpreFormat(!nMontoCol, 8) & ImpreFormat(!nSaldo, 8) _
                    & ImpreFormat(!nDeuda, 8) & ImpreFormat(!nMontoProceso, 8) _
                    & ImpreFormat(!nIGV, 5) _
                    & ImpreFormat(!nComisionRemate, 6) & ImpreFormat(!nMontoProceso - !nComisionRemate, 8) _
                    & ImpreFormat(!nSobranteRemate, 8) & ImpreFormat(!noroneto, 4) _
                    & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                        
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "PLANILLA DE PRENDAS VENDIDAS EN REMATE NRO " & psNroRemate, "Del " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 140, "", 3)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & Chr(10)
        lsCadImp = lsCadImp & Space(12) & "RESUMEN " & Space(14) & ImpreFormat(vSumTasaci, 8) _
                & ImpreFormat(vSumPresta, 8) & ImpreFormat(vSumSalCap, 8) & ImpreFormat(vSumDeuda, 8) _
                & ImpreFormat(vSumVtaBru, 8) & ImpreFormat(vSumIGV, 5) _
                & ImpreFormat(vSumComisRem, 6) & ImpreFormat(vSumVtaBru - vSumComisRem, 8) & ImpreFormat(vSumSobrante, 8) _
                & ImpreFormat(vSumOroNeto, 4) & Chr(10)
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimePlanillaVentaRemate = lsCadBuffer
End Function

'Imprime Planilla de Lotes No vendidos en Remate
Public Function nImprimePlanillaNoVendidosRemate(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal pnDiasVctoParaRemate As Integer, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, ByVal pnFactorPrecioVenta As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroRemate As String) As String
    
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
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

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing


    If lrDataRep Is Nothing Then
        MsgBox " No existen Créditos Pignoraticios No Rematados en la Agencia", vbInformation, " Aviso"
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
                vSumOroNeto = vSumOroNeto + !noroneto
                                        
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 4, 0) & Space(1) & !codesiaf & Space(1) & !cCtaCod & Space(1) & Format(!dVenc, "dd/mm/yyyy") _
                    & ImpreFormat(!nTasacion, 8) & ImpreFormat(!nMontoCol, 8) & ImpreFormat(!nSaldo, 8) _
                    & ImpreFormat(!nDeuda, 8) & ImpreFormat(!noroneto, 5) _
                    & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                        
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nImprimeRemateReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "PLANILLA DE PRENDAS NO VENDIDAS DEL REMATE NRO " & psNroRemate, "Del " & Format(psFecRemate, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 105, "", 4)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & Chr(10)
        lsCadImp = lsCadImp & Space(12) & "RESUMEN " & Space(33) & ImpreFormat(vSumTasaci, 8) _
                & ImpreFormat(vSumPresta, 8) & ImpreFormat(vSumSalCap, 8) & ImpreFormat(vSumDeuda, 8) _
                & ImpreFormat(vSumOroNeto, 5) _
                & Chr(10)
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        
    End If
    nImprimePlanillaNoVendidosRemate = lsCadBuffer
End Function


Public Function nObtieneListadoProcesosRG(ByVal psTpoProceso As String, ByVal psProcesoCadaAgencia As String, _
                Optional pbOrdenAsc As Boolean = True) As Recordset
'Obtiene Datos de Remate Credito Pignoraticio
Dim loValida As dColPFunciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String
Dim lsAgencia As String

On Error GoTo dError
    lsSQL = "SELECT RG.cTpoProceso, RG.cNroProceso, RG.dProceso, RG.nRGEstado " _
        & " FROM ColocPigRecGar RG " _
        & " WHERE RG.cTpoProceso ='" & psTpoProceso & "' And cCodAge = '" & psProcesoCadaAgencia & "' " _
        & " ORDER BY RG.cNroProceso " & IIf(pbOrdenAsc = True, "ASC", "DESC")
    Set loValida = New dColPFunciones
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
       ByVal pnPrevioMax As Long, ByVal psFechaSis As String) As String
Dim LALISTA() As String
Dim lnFilas As Long, lnItem As Long
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCartaModelo As String

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

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing


    If lrDataRep Is Nothing Then
        MsgBox " No existen Créditos Pignoraticios para Carta de Sobrante", vbInformation, " Aviso"
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
        lsCartaModelo = Replace(lsCartaModelo, "<<NOMBRE>>", PstaNombre(LALISTA(lnItem, 0), False), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<SOBRANTE>>", LALISTA(lnItem, 4), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<CONTRATO>>", LALISTA(lnItem, 3), , 2, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<FECREMATE>>", psNumRemate, , 1, vbTextCompare)
        'RTFTmp = Replace(RTFTmp, "<<HORREMATE>>", vHorRem, , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<DIRECCION>>", LALISTA(lnItem, 1), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<TELEFONO>>", LALISTA(lnItem, 2), , 1, vbTextCompare)
        
        lsCadImp = lsCadImp & lsCartaModelo & Chr(12)
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
        ByVal pnNroRepor As Integer) As String
        
Dim lsCadImp As String
Dim loImprimeCab As NColPImpre
    
    Set loImprimeCab = New NColPImpre
        lsCadImp = loImprimeCab.nImprimeCabeceraReportes(psNomCmac, psNomAgencia, psCodUser, psFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)
    Set loImprimeCab = Nothing
    lsCadImp = lsCadImp & Chr(10) & String(pnAnchoLinea, "-") & Chr(10)
    Select Case pnNroRepor
        Case 1 ' Planilla para Adjudicacion
            lsCadImp = lsCadImp & Space(1) & "Item    Contrato      Fec.Venc.     Cliente                 14kl    16kl    18kl     21kl     Gramos    Tasacion    SaldoCap      Deuda     ValRegistro" & Chr(10)
        Case 2 '
            lsCadImp = lsCadImp & Space(1) & "Item    ContratoAnt     Contrato          Cliente                   14kl      16kl     18kl     21kl     Gramos    Tasacion   SaldoCap  ValReg" & Chr(10)
        Case 9
            lsCadImp = lsCadImp & Space(1) & "Item    ContratoAnt     Contrato           Fec.Venc.     Cliente               14kl   16kl    18kl     21kl     Gramos     Tasacion    SaldoCap     ValReg" & Chr(10)
        Case 10
            lsCadImp = lsCadImp & Space(1) & "Item    ContratoAnt     Contrato           Cliente                    14kl    16kl    18kl     21kl     Gramos    Tasacion    SaldoCap      Deuda     ValRegistro " & Chr(10)
        Case 11
            lsCadImp = lsCadImp & Space(1) & "Item    ContratoAnt     Contrato            Piezas     Descripción                                 14Kl.    16Kl.    18Kl.    21Kl.  ValRegistro " & Chr(10)
    
    End Select
    lsCadImp = lsCadImp & String(pnAnchoLinea, "-") & Chr(10)
nImprimeAdjudicaReporteCabec = lsCadImp
End Function

'Imprime Planilla para Adjudicacion
Public Function nImprimePlanillaParaAdjudicacion(ByVal psFecProceso As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal pnNroRemParaAdjud As Integer, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroProceso As String) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim loCalculos As NColPCalculos

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
Dim loCalcula As NColPCalculos
Set loCalcula = New NColPCalculos
    
Dim vdiasAtraso As Integer
    lsSQL = " SELECT P.cCtaCod, P.nPrdestado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
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

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Then
        MsgBox " No existen Créditos Pignoraticios para Aviso de Vencimiento ", vbInformation, " Aviso"
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
                lsNombre = PstaNombre(!cNomCliente)
        
                'Set loCalculos = New NColPCalculos
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
                
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 4, 0) & Space(1) & !cCtaCod & Space(1) & Format(!dVenc, "dd/mm/yyyy") _
                    & ImpreFormat(ImpreCarEsp(lsNombre), 20, 1) _
                    & ImpreFormat(lnGra14, 6) & ImpreFormat(lnGra16, 6) & ImpreFormat(lnGra16, 6) & ImpreFormat(lnGra18, 6) _
                    & ImpreFormat(lnGramos, 6) & ImpreFormat(!nTasacion, 10) & ImpreFormat(!nSaldo, 10) _
                    & ImpreFormat(lnDeuda, 10) _
                    & ImpreFormat(lnValRegis, 10) _
                    & Chr(10)
                '& ImpreFormat(lnPreMerc, 10)
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                        
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nImprimeAdjudicaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "PLANILLA PARA ADJUDICACION NRO " & psNroProceso, "Del " & Format(psFecProceso, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 155, "", 1)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & Chr(10)
        lsCadImp = lsCadImp & Space(50) & "RESUMEN" & ImpreFormat(vSumGramos, 40) & ImpreFormat(vSumTasaci, 10) _
                & ImpreFormat(vSumSalCap, 10) & ImpreFormat(vSumDeuda, 10) _
                & ImpreFormat(vSumValRegis, 10) _
                & "" _
                & Chr(10)
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimePlanillaParaAdjudicacion = lsCadBuffer
End Function


Public Function nImprimePlanillaParaAdjudicacionConSIAF(ByVal psFecProceso As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal pnNroRemParaAdjud As Integer, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroProceso As String) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim loCalculos As NColPCalculos

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
Dim loCalcula As NColPCalculos
Set loCalcula = New NColPCalculos
    
Dim vdiasAtraso As Integer
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

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Then
        MsgBox " No existen Créditos Pignoraticios para Aviso de Vencimiento ", vbInformation, " Aviso"
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
                lsNombre = PstaNombre(!cNomCliente)
        
                'Set loCalculos = New NColPCalculos
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
                
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 4, 0) & Space(1) & !codesiaf & Space(1) & !cCtaCod & Space(1) & Format(!dVenc, "dd/mm/yyyy") _
                    & ImpreFormat(ImpreCarEsp(lsNombre), 20, 1) _
                    & ImpreFormat(lnGra14, 6) & ImpreFormat(lnGra16, 6) & ImpreFormat(lnGra18, 6) & ImpreFormat(lnGra21, 6) _
                    & ImpreFormat(lnGramos, 6) & ImpreFormat(!nTasacion, 10) & ImpreFormat(!nSaldo, 10) _
                    & ImpreFormat(lnValRegis, 10) _
                    & Chr(10)
                    
                    '& ImpreFormat(lnDeuda, 10) _
                '& ImpreFormat(lnPreMerc, 10)
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                        
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nImprimeAdjudicaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "PLANILLA PARA ADJUDICACION NRO " & psNroProceso, "Del " & Format(psFecProceso, "dd/mm/yyyy") & " - Agencia " & psAgencia, lnPage, 155, "", 9)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & Chr(10)
        lsCadImp = lsCadImp & Space(50) & "RESUMEN" & ImpreFormat(vSumGramos, 40) & ImpreFormat(vSumTasaci, 10) _
                & ImpreFormat(vSumSalCap, 10) & ImpreFormat(vSumDeuda, 10) _
                & ImpreFormat(vSumValRegis, 10) _
                & "" _
                & Chr(10)
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimePlanillaParaAdjudicacionConSIAF = lsCadBuffer
End Function




Public Function nAdjudicaLotes(ByVal psMovNro As String, ByVal psFechaHora As String, _
        ByVal psTipoAdju As String, ByVal psNroAdju As String, _
        ByVal pnPreOroInternacional As Currency, ByVal pnTipoCambio As Double, _
        ByVal psFecAdjudica As String, ByVal pnPrecioOro14 As Double, ByVal pnPrecioOro16 As Double, _
        ByVal pnPrecioOro18 As Double, ByVal pnPrecioOro21 As Double, ByVal psFecParaAdjudica As String, _
        ByVal psProcesoCadaAgencia As String, ByVal pnNroRemParaAdjud As Integer) As Boolean

Dim lsSQL As String
Dim loDataAdj As dColPFunciones
Dim lrDataAdj As ADODB.Recordset
Dim lnMovNro As Long
Dim loCalcula As NColPCalculos
Dim loRegPig As DColPActualizaBD
Dim lsOpeCod As String
Dim lsNextAdjud As String

lsOpeCod = geColPAdjudica

Dim lnPreMercado As Currency, lnPreDeuda As Currency, lnPreAdj As Currency

lsSQL = " SELECT c.cCtacod, c.dVenc, cp.nTasacion, c.nMontoCol, p.nSaldo, p.nPrdEstado, C.cLineaCred, " _
    & " SUM( case when cKilataje ='14' then nPesoOro else 0  end ) as  nK14 , " _
    & " SUM( case when cKilataje ='16' then nPesoOro else 0  end ) as  nK16 , " _
    & " SUM( case when cKilataje ='18' then nPesoOro else 0  end ) as  nK18 , " _
    & " SUM( case when cKilataje ='21' then nPesoOro else 0  end ) as  nK21 , " _
    & " nTasaIntVenc = (SELECT nTasaIni FROM ColocLineaCreditoTasa  " _
    & "               WHERE cLineaCred = C.cLineaCred  " _
    & "               AND nColocLinCredTasaTpo = '02' ) " _
    & " FROM Producto p Inner Join Colocaciones c on p.cctacod = c.cctacod " _
    & " Inner join ColocPignoraticio cp on c.cctacod = cp.cctacod " _
    & " Left Join ColocPigjoya cpjoy on cp.cctacod = cpjoy.cctacod " _
    & " WHERE p.nPrdEstado IN (" & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov & " )  " _
    & IIf(Mid(psProcesoCadaAgencia, 4, 2) = "00", "", " AND Substring(P.cCtaCod,4,2) ='" & Right(psProcesoCadaAgencia, 2) & "' ") & " AND nNroRemate = " & pnNroRemParaAdjud _
    & " GROUP By c.cCtacod, c.dVenc, cp.nTasacion, c.nMontoCol, p.nSaldo, p.nprdEstado, C.cLineaCred  "

        
Set loDataAdj = New dColPFunciones
    Set lrDataAdj = loDataAdj.dObtieneRecordSet(lsSQL)
Set loDataAdj = Nothing
        
If lrDataAdj Is Nothing Then
    MsgBox "No hay creditos para Adjudicar", vbInformation, "Aviso"
    
Else
    lnPreMercado = 0: lnPreDeuda = 0: lnPreAdj = 0
    
    
    Set loRegPig = New DColPActualizaBD
        
        'Inicializa Transaccion
        loRegPig.dBeginTrans
        mbTrans = True
        
        '** Inserta Mov
        Call loRegPig.dInsertMov(psMovNro, lsOpeCod, "Adjudicacion Cred Pign", gMovEstContabMovContable, gMovFlagVigente, False)
        
        ' Obtiene nMovNro
        lnMovNro = loRegPig.dGetnMovNro(psMovNro)
        
        With lrDataAdj
            
            Do While Not .EOF
                'Precio Base
                lnPreMercado = !nK14 * pnPrecioOro14 + !nK16 * pnPrecioOro16 + !nK18 * pnPrecioOro18 + !nK21 * pnPrecioOro21
                lnPreMercado = Round(lnPreMercado, 2)
                
                'Calcula Deuda
                'Set loCalcula = New NColPCalculos
                '    lnPreDeuda = loCalcula.nCalculaDeudaPignoraticio(!nSaldo, !dFecVenc, !nTasacion, !nTasaIntVenc, _
                '        !nCostCusto, pnTasaImpuesto, !nPrdEstado, pnTasaPreparacionRemate, Format(psFecRemate, "mm/dd/yyyy"))
                '
                '    lnPreDeuda = Round(lnPreDeuda, 2)
                '    Set loCalcula = Nothing
                        
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
                Call loRegPig.dInsertColocacEstado(!cCtaCod, psFechaHora, gColPEstAdjud, 1, lnPreAdj, "Adjudicacion de Joyas", _
                     gColocCalendCodFFCF, 0, 0, 0, "0", 0, 0, 0, 0, False)
                
                '** Inserta RGDet
                Call loRegPig.dInsertColocPigRGDet("A", psNroAdju, !cCtaCod, gColPRecGarAdjudicado, lnPreAdj, 0, lnPreMercado, lnPreAdj, 0, psProcesoCadaAgencia, False)
                
                '** Inserta MovCol
                Call loRegPig.dInsertMovCol(lnMovNro, lsOpeCod, !cCtaCod, 1, lnPreAdj, 0, "", 0, gColPEstAdjud, !nSaldo, False)
            
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
                
                .MoveNext
            Loop
            .Close
        End With
        'Actualiza el Proceso Adjudicacion
        Call loRegPig.dUpdateColocPigRecGar("A", psNroAdju, psProcesoCadaAgencia, , gColPRecGarEstTerminado, psFecAdjudica, False)
        
        'Insert el Sgte proceso de Adjudicacion
        lsNextAdjud = String(4 - Len(Trim(Str(Val(psNroAdju) + 1))), "0") & Trim(Str(Val(psNroAdju) + 1))
        Call loRegPig.dInsertColocPigRecGar("A", lsNextAdjud, Format(DateAdd("m", 1, psFecAdjudica) & " 08:00", "mm/dd/yyyy hh:mm"), gColPRecGarEstNoIniciado, psProcesoCadaAgencia, , False)
        
        loRegPig.dCommitTrans
        
    Set loRegPig = Nothing
End If
Set lrDataAdj = Nothing

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
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim loCalculos As NColPCalculos

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


    lsSQL = " SELECT CODESIAF=ISNULL(R.CCTACODANT,'                  '),P.cCtaCod, P.nPrdEstado, C.dVenc, CP.nTasacion, C.nMontoCol, P.nSaldo, " _
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

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing


    If lrDataRep Is Nothing Then
        MsgBox " No existen Créditos Pignoraticios Adjudicados ", vbInformation, " Aviso"
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
                lsNombre = PstaNombre(!cNomCliente)
                
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
                                        
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !codesiaf & Space(1) & !cCtaCod & Space(1) _
                    & ImpreFormat(ImpreCarEsp(lsNombre), 20, 1) _
                    & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & ImpreFormat(v18, 6) & ImpreFormat(v21, 6) _
                    & ImpreFormat(lnGramos, 6) & ImpreFormat(!nTasacion, 8) & ImpreFormat(!nSaldo, 8) _
                    & ImpreFormat(!nAdjValRegistro, 8) _
                    & Chr(10)
                   ' & ImpreFormat(lnDeuda, 10) _

                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                        
                'lnLineas = lnLineas + 1
                'lsCadImp = lsCadImp & Chr(10)
                
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
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
        
        lsCadImp = lsCadImp & String(144, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(37) & "RESUMEN " & ImpreFormat(vSumGramos, 40) & ImpreFormat(vSumTasaci, 10) _
                & ImpreFormat(vSumSalCap, 10) & ImpreFormat(vSumDeuda, 10) _
                & ImpreFormat(vSumValRegis, 10) _
                & "" _
                & Chr(10)
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimeListadoAdjudicados = lsCadBuffer
End Function

'Imprime Listado de Adjudicados
Public Function nImprimeListadoAdjudicadosConSIAF(ByVal psFecRemate As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal psNroProcesoAdjud As String, ByVal psFechaSis As String, ByVal pnTasaCustodiaVencida As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnTasaPreparacionRemate As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim loCalculos As NColPCalculos

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

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing


    If lrDataRep Is Nothing Then
        MsgBox " No existen Créditos Pignoraticios Adjudicados ", vbInformation, " Aviso"
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
                lsNombre = PstaNombre(!cNomCliente)
                
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
                                        
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !codesiaf & Space(1) & !cCtaCod & Space(1) _
                    & ImpreFormat(ImpreCarEsp(lsNombre), 20, 1) _
                    & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & ImpreFormat(v18, 6) & ImpreFormat(v21, 6) _
                    & ImpreFormat(lnGramos, 6) & ImpreFormat(!nTasacion, 10) & ImpreFormat(!nSaldo, 10) _
                    & ImpreFormat(lnDeuda, 10) _
                    & ImpreFormat(!nAdjValRegistro, 10) _
                    & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                        
                'lnLineas = lnLineas + 1
                'lsCadImp = lsCadImp & Chr(10)
                
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
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
        
        lsCadImp = lsCadImp & String(161, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(55) & "RESUMEN " & ImpreFormat(vSumGramos, 40) & ImpreFormat(vSumTasaci, 10) _
                & ImpreFormat(vSumSalCap, 10) & ImpreFormat(vSumDeuda, 10) _
                & ImpreFormat(vSumValRegis, 10) _
                & "" _
                & Chr(10)
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimeListadoAdjudicadosConSIAF = lsCadBuffer
End Function




'Imprime Aviso de Remate
Public Function nImprimeAvisoSubasta(ByVal psTextoCarta As String, ByVal psFecSubasta As String, _
       ByVal psAgencia As String, ByVal pnPrevioMax As Long, _
       ByVal psFechaSis As String) As String
Dim LALISTA() As String
Dim lnFilas As Long, lnItem As Long
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCartaModelo As String

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

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing


    If lrDataRep Is Nothing Then
        MsgBox " No existen Créditos Pignoraticios para Aviso de Subasta ", vbInformation, " Aviso"
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
        lsCartaModelo = Replace(lsCartaModelo, "<<NOMBRE>>", PstaNombre(LALISTA(lnItem, 0), False), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<CONTRATO>>", LALISTA(lnItem, 3), , 2, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<FECSUBASTA>>", psFecSubasta, , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<DIRECCION>>", LALISTA(lnItem, 1), , 1, vbTextCompare)
        lsCartaModelo = Replace(lsCartaModelo, "<<TELEFONO>>", LALISTA(lnItem, 2), , 1, vbTextCompare)
        lsCadImp = lsCadImp & lsCartaModelo & Chr(12)
        'prgList.Value = vCont
    Next lnItem
    
    Erase LALISTA
    nImprimeAvisoSubasta = lsCadImp

Exit Function
   
dError:
    Err.Raise Err.Number, "Obtiene Datos Imprime Listado Aviso Vencimiento  <<nImprimeAvisoVencimiento>>", Err.Description
End Function

Public Function nImprimeSubastaReporteCabec(ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, _
        ByVal psFechaSis As String, ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String, _
        ByVal pnNroRepor As Integer) As String
        
Dim lsCadImp As String
Dim loImprimeCab As NColPImpre
    
    Set loImprimeCab = New NColPImpre
        lsCadImp = loImprimeCab.nImprimeCabeceraReportes(psNomCmac, psNomAgencia, psCodUser, psFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)
    Set loImprimeCab = Nothing
    lsCadImp = lsCadImp & Chr(10) & String(pnAnchoLinea, "-") & Chr(10)
    Select Case pnNroRepor
        Case 1 ' Planilla de Remate
            lsCadImp = lsCadImp & Space(1) & "Item    Contrato      Fec.Venc.     Cliente             Gramos    Tasacion   Prestamo  SaldoCap  Interes  Custodia Impuesto Prep.Rem   Deuda   P.Base Boveda " & Chr(10)
        Case 2 ' Listados para Remate
            lsCadImp = lsCadImp & Space(1) & "Item    Contrato    Piezas     Descripción                                 14Kl.    16Kl.    18Kl.    21Kl.  Prec.Base " & Chr(10)
    End Select
    lsCadImp = lsCadImp & String(pnAnchoLinea, "-") & Chr(10)
nImprimeSubastaReporteCabec = lsCadImp
End Function

'Imprime Planilla de Remate
Public Function nImprimePlanillaParaSubasta(ByVal psFecSubasta As String, ByVal psAgencia As String, pnPrevioMax As Long, _
        ByVal psFechaSis As String, ByVal pnTasaIGV As Double, _
        ByVal pnPreOro14 As Currency, ByVal pnPreOro16 As Currency, ByVal pnPreOro18 As Currency, ByVal pnPreOro21 As Currency, _
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroSubasta As String) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim loCalculos As NColPCalculos

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

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing


    If lrDataRep Is Nothing Then
        MsgBox " No existen Créditos Pignoraticios para Aviso de Vencimiento ", vbInformation, " Aviso"
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
                
                
                lsNombre = PstaNombre(!cNomCliente)
        

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
                                        
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 4, 0) & Space(1) & !cCtaCod & Space(1) & Format(!dVenc, "dd/mm/yyyy") _
                    & ImpreFormat(ImpreCarEsp(lsNombre), 20, 1) _
                    & ImpreFormat(lnGramos, 6) & ImpreFormat(!nTasacion, 8) & ImpreFormat(!nMontoCol, 8) _
                    & Format(!dVenc, "dd/mm/yyyy") & ImpreFormat(lnPreVenta, 8) & Space(2) & Mid(!cAgeBoveda, 1, 2) _
                    & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                        
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nImprimeSubastaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "PLANILLA DE PRENDAS PARA SUBASTA NRO " & psNroSubasta, "", lnPage, 120, "", 1)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & Chr(10)
        lsCadImp = lsCadImp & Space(15) & "RESUMEN " & ImpreFormat(vSumGramos, 41) & ImpreFormat(vSumTasaci, 8) _
                & ImpreFormat(vSumPresta, 8) & ImpreFormat(vSumSalCap, 8) & ImpreFormat(vSumIntVen, 6) _
                & ImpreFormat(vSumCosCus, 6) & ImpreFormat(vSumImpues, 6) & ImpreFormat(vSumCoPrRe, 6) _
                & ImpreFormat(vSumPreDeu, 8) & ImpreFormat(vSumPreVen, 8) _
                & Chr(10)
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimePlanillaParaSubasta = lsCadBuffer
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
        ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, ByVal psNroSubasta As String) As String
    
Dim ldFecAviso As Date
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim loCalculos As NColPCalculos

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

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing


    If lrDataRep Is Nothing Then
        MsgBox " No existen Créditos Pignoraticios para Aviso de Vencimiento ", vbInformation, " Aviso"
        'MuestraImpresion = False

        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nImprimeSubastaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                            "LISTADO DE CONTRATOS PARA SUBASTA NRO " & psNroSubasta, "DEL " & Format(psFecSubasta, "dd/mm/yyyy"), lnPage, 120, "", 2)

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
                
                'Set loCalculos = New NColPCalculos
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
                                        
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !cCtaCod & Space(1) & ImpreFormat(!npiezas, 3, 0) & Space(2) & ImpreFormat(QuiebreTexto(!cLote, 1), 42, 1) _
                    & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & ImpreFormat(v18, 6) & ImpreFormat(v21, 6) _
                    & ImpreFormat(lnPreVenta, 8) _
                    & Chr(10)
                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                                        

                Do While vCuenta < 15
                    vCuenta = vCuenta + 1
                    If Len(QuiebreTexto(!cLote, vCuenta + 1)) > 0 Then
                        lsCadImp = lsCadImp & Space(28) & QuiebreTexto(!cLote, vCuenta + 1) & Chr(10)
                        lnLineas = lnLineas + 1
                    End If
                Loop
                
                lnLineas = lnLineas + 1
                lsCadImp = lsCadImp & Chr(10)
                
                lnIndice = lnIndice + 1
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nImprimeSubastaReporteCabec(psNomCmac, psNomAgencia, psCodUser, psFechaSis, _
                                        "LISTADO DE CONTRATOS PARA SUBASTA NRO " & psNroSubasta, "DEL " & Format(psFecSubasta, "dd/mm/yyyy"), lnPage, 120, "", 2)
                    lnLineas = 7
                End If
                'vCont = vCont + 1
                'prgList.Value = vCont
                .MoveNext
        
            Loop
        End With
        
        lsCadImp = lsCadImp & Chr(10)
        lsCadImp = lsCadImp & String(120, "*") & Chr(10)
                    
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        'prgList.Visible = False
        'prgList.Value = 0
        
    End If
    nImprimeListadoParaSubasta = lsCadBuffer
End Function
Public Function nIniciaSubasta(ByVal psNroSubasta As String, ByVal psFecSubasta As String, _
        ByVal pnTasaIGV As Double, ByVal pnPrecioOro14 As Double, ByVal pnPrecioOro16 As Double, _
        ByVal pnPrecioOro18 As Double, ByVal pnPrecioOro21 As Double, ByVal psProcesoAgencia As String) As Boolean

Dim lsSQL As String
Dim loDataIniSub As dColPFunciones
Dim lrDataInicia As ADODB.Recordset
              
Dim loCalcula As NColPCalculos
Dim loRegPig As DColPActualizaBD


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

Set loDataIniSub = New dColPFunciones
    Set lrDataInicia = loDataIniSub.dObtieneRecordSet(lsSQL)
Set loDataIniSub = Nothing

If lrDataInicia Is Nothing Then

Else
    lnPreBase = 0
    With lrDataInicia
        Set loRegPig = New DColPActualizaBD
        
        Do While Not .EOF
            'Precio Base
            lnPreBase = !nK14 * pnPrecioOro14 + !nK16 * pnPrecioOro16 + !nK18 * pnPrecioOro18 + !nK21 * pnPrecioOro21
            lnPreBase = lnPreBase * (1 + pnTasaIGV)
            lnPreBase = Round(lnPreBase, 2)
            'Precio Deuda
            Set loCalcula = New NColPCalculos
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
              
Dim loRegPig As DColPActualizaBD

    'Genera el sgte Nro de Subasta
    lsNextSubasta = String(4 - Len(Trim(Str(Val(psNroSubasta) + 1))), "0") & Trim(Str(Val(psNroSubasta) + 1))
    
    Set loRegPig = New DColPActualizaBD
    
        'Cambia estado de los que NO se han vendido
        lsSQL = " UPDATE ColocPigRGDet SET nRGDetEstado = " & gColPRecGarVentaNoVendido & " WHERE cTpoProceso = 'S' AND cNroProceso = '" & psNroSubasta & "' AND nRGDetEstado = " & gColPRecGarVentaEnProceso
        Call loRegPig.dEjecutar(lsSQL)
    
        'Termina el Subasta
        Call loRegPig.dUpdateColocPigRecGar("S", psNroSubasta, psProcesoCadaAgencia, , gColPRecGarEstTerminado, , False)
        
        'Inserta la Sgte subasta
        Call loRegPig.dInsertColocPigRecGar("S", lsNextSubasta, Format(DateAdd("m", 1, Format(psFecProceso, "mm/dd/yyyy")) & " 08:00", "mm/dd/yyyy hh:mm"), gColPRecGarEstNoIniciado, psProcesoCadaAgencia, , False)
     
    Set loRegPig = Nothing

End Function

Public Function nObtieneSobranteRemateProceso(ByVal psNroProceso As String, ByVal psAgencia As String) As Currency
'Obtiene el Sobrante del Ultimo Remate
Dim loValida As dColPFunciones
Dim lrs As ADODB.Recordset
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
    Set loValida = New dColPFunciones
        Set lrs = loValida.dObtieneRecordSet(lsSQL)
    Set loValida = Nothing
    If lrs.BOF And lrs.EOF Then
        MsgBox "No se tiene Información de Sobrante de Remate", vbInformation, "Aviso"
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

Dim loConec As DConecta
Dim lrDatosRem As Recordset
Dim lsUltRemate As String
Dim lsSQL As String

lsUltRemate = ""
Set loConec = New DConecta
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
    If lrDatosRem!nRGEstado = gColPRecGarEstIniciado Then
        lsUltRemate = lrDatosRem!cNroProceso
    End If
    lrDatosRem.Close
Set loConec = Nothing
nObtieneRemateIniciado = lsUltRemate
End Function

Public Function nObtieneCtaSobranteRemate(ByVal psRemateCadaAgencia As String) As String
Dim loConstSis As NConstSistemas
Dim PObjConec As DConstante
Dim lrCtaAho As ADODB.Recordset
Dim lsCtaSobranteRemate As String

    ' Obtiene la cuenta de Ahorros de Sobrante de Remate  ****
    If Right(psRemateCadaAgencia, 2) = "00" Then  ' Remate Centralizado
        Set loConstSis = New NConstSistemas
            lsCtaSobranteRemate = loConstSis.LeeConstSistema(61)  ' Cuenta SobranteRemate
        Set loConstSis = Nothing
    Else ' Remate en Cada Agencia
        Set PObjConec = New DConstante
        Set lrCtaAho = PObjConec.RecuperaConstantes(3207, , "C.nConsValor")
        If lrCtaAho.BOF And lrCtaAho.EOF Then
           MsgBox "No se han configurado las Ctas de Ahorro de Sobrantes", vbInformation, "Aviso"
           lsCtaSobranteRemate = ""
           Exit Function
        End If
        Do While Not lrCtaAho.EOF
           If Val(lrCtaAho!nConsValor) = Val(Right(psRemateCadaAgencia, 2)) Then
               lsCtaSobranteRemate = Trim(lrCtaAho!CCONSDESCRIPCION)
               Exit Do
           End If
           lrCtaAho.MoveNext
        Loop
        Set lrCtaAho = Nothing
    End If
    nObtieneCtaSobranteRemate = lsCtaSobranteRemate
End Function
