VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DPigActualizaBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim oConn As DConecta
Dim lsSQL As String

'***********************************
' CMPCL - Inicializacion de la Clase
' CAFF - 14/08/2002
'***********************************
Private Sub Class_Initialize()
Dim oIni As COMConecta.DCOMClasIni
Dim sConn As String

    Set oIni = New COMConecta.DCOMClasIni
    sConn = oIni.CadenaConexion
    Set oIni = Nothing
    
Set oConn = New DConecta
If oConn.AbreConexion(sConn) = False Then
    Call RaiseError(MyUnhandledError, "DPigActualizaBD:Initialize. Error en Conexion a Base de datos")
End If

End Sub

Private Sub Class_Terminate()
    oConn.CierraConexion
    Set oConn = Nothing
End Sub

Public Sub dBeginTrans()
    oConn.BeginTrans
End Sub

Public Sub dRollbackTrans()
    oConn.RollbackTrans
End Sub

Public Sub dCommitTrans()
    oConn.CommitTrans
End Sub

Public Function GeneraMovNro(ByVal pdFecha As Date, Optional ByVal psCodAge As String = "07", Optional ByVal psUser As String = "SIST", Optional psMovNro As String = "") As String

    On Error GoTo GeneraMovNroErr
    Dim rs As Recordset
    Dim sql As String
    
    If oConn.AbreConexion = False Then Exit Function
    If psMovNro = "" Or Len(psMovNro) <> 25 Then
       sql = "sp_GeneraMovNro '" & Format(pdFecha & " " & oConn.GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "','" & Right(psCodAge, 2) & "','" & psUser & "'"
    Else
       sql = "sp_GeneraMovNro '','','','" & psMovNro & "'"
    End If
    
    Set rs = oConn.Ejecutar(sql)
    If Not rs.EOF Then
        GeneraMovNro = rs.Fields(0)
    End If
    
    Set rs = Nothing
    
    Exit Function
GeneraMovNroErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:GeneraMovNro Method")
End Function

Public Sub dInsertProducto(ByVal psctacod As String, ByVal pnTasaInteres As Double, ByVal pnSaldo As Currency, _
       ByVal pnEstado As ColocEstado, ByVal psFecEstado As String, ByVal pnTransac As Integer, Optional ByVal pbEjecBatch As Boolean = False)

On Error GoTo ErrordInsertProducto

   lsSQL = "INSERT Producto (cCtaCod, nTasaInteres, nSaldo, nPrdEstado, dPrdEstado, nTransacc ) " & _
         "VALUES ('" & psctacod & "'," & pnTasaInteres & "," & pnSaldo & "," & pnEstado & ",'" & psFecEstado & "'," & pnTransac & ")"
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
    
Exit Sub

ErrordInsertProducto:
    Err.Raise Err.Number, "Error En Proceso dInsertProducto", Err.Description

End Sub

Public Sub dInsertProductoPersona(ByVal psCuenta As String, ByVal psPersona As String, ByVal pnRelacion As ColocRelacPers, _
       Optional ByVal pbEjecBatch As Boolean = False)

On Error GoTo ErrordInsertProductoPersona

    lsSQL = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac) " & _
          "VALUES ('" & psCuenta & "','" & psPersona & "'," & pnRelacion & ")"
    
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
    
Exit Sub
 
ErrordInsertProductoPersona:
    Err.Raise Err.Number, "Error En Proceso dInsertProductoPersona", Err.Description
    
End Sub

Public Sub dInsertProductoTasaInteres(ByVal psCuenta As String, ByVal pnPrdTasaInteres As ColocTipoTasa, ByVal pnTasaInteres As Double, _
       Optional ByVal pbEjecBatch As Boolean = False)

On Error GoTo ErrordInsertProductoTasaInteres

    lsSQL = "INSERT ProductoTasaInteres (cCtaCod, nPrdTasaInteres, nTasaInteres) " & _
          "VALUES ('" & psCuenta & "'," & pnPrdTasaInteres & "," & pnTasaInteres & ")"
    
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
    
Exit Sub

ErrordInsertProductoTasaInteres:
    Err.Raise Err.Number, "Error En Proceso dInsertProductoTasaInteres", Err.Description

End Sub

Public Sub dInsertColocaciones(ByVal psctacod As String, ByVal pnPlazo As Integer, ByVal psFecVenc As String, ByVal pnMontoCol As Currency, _
       ByVal pnTipoCalend As ColocTipoCalend, ByVal psMovNro As String, ByVal psLineaCred As String, ByVal psFecVigencia As String, _
       Optional ByVal pbEjecBatch As Boolean = False)

On Error GoTo ErrordInsertColocaciones

    lsSQL = "INSERT Colocaciones (cCtaCod,nPlazo,dVenc,nMontoCol,nColocCalendCod,cUltimaActualizacion,cLineaCred, dVigencia) " _
        & "VALUES ('" & psctacod & "'," & pnPlazo & ",'" & psFecVenc & "'," & pnMontoCol & "," _
        & pnTipoCalend & ",'" & psMovNro & "','" & psLineaCred & "','" & psFecVigencia & "')"
        
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
    
Exit Sub

ErrordInsertColocaciones:
    Err.Raise Err.Number, "Error En Proceso dInsertColocaciones", Err.Description

End Sub

Public Sub dInsertColocacEstado(ByVal psctacod As String, ByVal psFechaPrdEstado As String, ByVal psPrdEstado As ColocEstado, _
       ByVal pnCuotas As Integer, ByVal pnMonto As Currency, ByVal psDescripcion As String, ByVal pnTipoCalendario As ColocTipoCalend, _
       ByVal pnPeriodoFechaFija As Integer, ByVal pnPeriodoGracia As Integer, ByVal pnPlazo As Integer, ByVal pnTipoGracia As Integer, _
       ByVal pnTipoDesembolso As Integer, ByVal pnProxMes As Integer, ByVal pnCalendDinamico As Integer, ByVal pnMotivoRechazo As Integer, _
       Optional pbEjecBatch As Boolean = False)

On Error GoTo ErrordInsertColocacEstado

    lsSQL = "INSERT ColocacEstado (cCtaCod, dPrdEstado, nPrdEstado, nCuotas, nMonto, cDescripcion, nColocCalendCod, nPeriodoFechaFija, " & _
        "nPeriodoGracia, nPlazo, nTipoGracia, nTipoDesembolso, nProxMes, nCalendDinamico, nMotivoRechazo ) " & _
        "VALUES ('" & psctacod & "','" & psFechaPrdEstado & "'," & psPrdEstado & "," & pnCuotas & "," & pnMonto & ",'" & _
        psDescripcion & "'," & pnTipoCalendario & "," & pnPeriodoFechaFija & "," & pnPeriodoGracia & "," & pnPlazo & "," & _
        pnTipoGracia & "," & pnTipoDesembolso & "," & pnProxMes & "," & pnCalendDinamico & "," & pnMotivoRechazo & ") "

    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If

Exit Sub

ErrordInsertColocacEstado:
    Err.Raise Err.Number, "Error En Proceso dInsertColocacEstado", Err.Description

End Sub

Public Sub dInsertColocCalendario(ByVal psctacod As String, ByVal pnNroCalend As Integer, ByVal pnColocCalendApl As ColocCalendApl, _
       ByVal pnCuota As Integer, ByVal psFecVenc As String, ByVal pnColocCalendEstado As Integer, ByVal psDescrip As String, _
       ByVal psColocCalenFlag As String, ByVal pnColocCalendProc As ColocCalendConceptoProc, Optional ByVal pbEjecBatch As Boolean = False)

On Error GoTo ErrordInsertColocCalendario

    lsSQL = "INSERT ColocCalendario (cCtaCod,nNroCalen,nColocCalendApl,nCuota,dVenc,nColocCalendEstado,cDescripcion,cColocCalenFlag, nCalendProc) " _
         & "VALUES ('" & psctacod & "'," & pnNroCalend & "," & pnColocCalendApl & "," & pnCuota & ",'" & psFecVenc & "'," & pnColocCalendEstado & ",'" _
         & psDescrip & "','" & psColocCalenFlag & "'," & pnColocCalendProc & ")"
    
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
    
Exit Sub

ErrordInsertColocCalendario:
    Err.Raise Err.Number, "Error En Proceso dInsertColocCalendario", Err.Description
    
End Sub

Public Sub dInsertColocCalendDetPig(ByVal psctacod As String, ByVal pnNroCalend As Integer, ByVal pnColocCalendApl As ColocCalendApl, _
        ByVal pnCuota As Integer, ByVal pnColocConceptoCod As ColocConcepto, ByVal pnMonto As Currency, ByVal pnMontoPagado As Currency, _
        ByVal psFlag As String, Optional ByVal pbEjecBatch As Boolean = False)

    On Error GoTo ErrordInsertColocCalendDetPig
    
    lsSQL = "INSERT ColocCalendDetPig (cCtaCod,nNroCalen,nColocCalendApl,nCuota,nPrdConceptoCod,nMonto,nMontoPagado,cFlag) " _
         & "VALUES ('" & psctacod & "'," & pnNroCalend & "," & pnColocCalendApl & "," & pnCuota & "," & pnColocConceptoCod & "," & pnMonto & "," _
         & pnMontoPagado & ",'" & psFlag & "' ) "
    
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
    
Exit Sub

ErrordInsertColocCalendDetPig:
    Err.Raise Err.Number, "Error En Proceso dInsertColocCalendDetPig", Err.Description
    
End Sub
'*********************************************
'CAFF - 16/08/2002 -  Insert ColocPigno
'*********************************************
Public Sub dInsertColocPigno(ByVal psctacod As String, ByVal pnNroUsoLinea As Integer, ByVal pnPiezas As Integer, ByVal pnEstadoCont As ColocEstado, _
        ByVal pnUbicaLote As ColocPigUbica, ByVal pnNumCalend As Integer, ByVal pnTipoTasacion As ColocPigTipoTasacion, ByVal pnNroDuplic As Integer, _
        ByVal pnDiasAtraso As Integer, ByVal pnTotalBalanza As Currency, ByVal pnNumCalendDesem As Integer, Optional ByVal pbEjecBatch As Boolean = False)

    On Error GoTo ErrordInsertColocPigno

    lsSQL = "INSERT ColocPigno (cCtaCod, nUsoLineaNro, nPiezas, nEstadoCont, nUbicaLote, nNumCalend, nTipoTasacion, nNroDuplic, " & _
         "nDiasAtraso, nNroAmort, nTotalBalanza, nNumCalendDesem) VALUES ('" & psctacod & "'," & pnNroUsoLinea & "," & pnPiezas & "," & pnEstadoCont & "," & pnUbicaLote & "," & _
         pnNumCalend & "," & pnTipoTasacion & "," & pnNroDuplic & "," & pnDiasAtraso & ",0," & pnTotalBalanza & "," & pnNumCalendDesem & ")"

    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If

Exit Sub

ErrordInsertColocPigno:
    Err.Raise Err.Number, "Error En Proceso dInsertColocPigno", Err.Description

End Sub

'*********************************************
  'CAFF - 16/08/2002 -  Insert ColocPigno
'*********************************************
Public Sub dInsertColocPigJoyaTasacion(ByVal psctacod As String, ByVal pnTipoTasacion As ColocPigTipoTasacion, ByVal pnItemPieza As Integer, _
        ByVal pnTipoMaterial As Integer, ByVal pnTipoJoya As Integer, ByVal pnSubTipoJoya As Integer, _
        ByVal pnestJoya As Integer, pnPesoBruto As Double, pnPesoNeto As Double, pnTasacion As Currency, ByVal pnTasacionAdic As Currency, _
        ByVal psDescrip As String, ByVal psDescripAdic As String, ByVal pnUbicaPieza As ColocPigUbica, ByVal psCodPers As String, _
        ByVal psMovNro As String, Optional ByVal pbEjecBatch As Boolean = False)

    On Error GoTo ErrordInsertColocPigJoyaTasacion
    
    lsSQL = "INSERT ColocPigJoyaTasacion (cCtaCod, nTipoTasacion, nItemPieza, nMaterial, nTipoJoya, nSubTipoJoya, nConservacion, nPesoBruto, " & _
        "nPesoNeto, nTasacion, nTasacionAdicional, cDescripcion, cDescripcionAdic, nUbicaPieza, cPersCod, cUltimaActualizacion) VALUES ('" & _
        psctacod & "'," & pnTipoTasacion & "," & pnItemPieza & "," & pnTipoMaterial & "," & pnTipoJoya & "," & pnSubTipoJoya & "," & _
        pnestJoya & "," & pnPesoBruto & "," & pnPesoNeto & "," & pnTasacion & "," & pnTasacionAdic & ",'" & psDescrip & "','" & _
        psDescripAdic & "'," & pnUbicaPieza & ",'" & psCodPers & "','" & psMovNro & "')"
    
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
    
Exit Sub

ErrordInsertColocPigJoyaTasacion:
    Err.Raise Err.Number, "Error En Proceso dInsertColocPigJoyaTasacion", Err.Description
    
End Sub

Public Sub dInsertMov(ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psMovDesc As String, ByVal pnMovEstado As MovEstado, _
    Optional pnMovFlag As MovFlag = gMovFlagVigente, Optional pbEjecBatch As Boolean = False)
   
On Error GoTo InsertMovErr
    
    lsSQL = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado, nMovFlag) " _
          & "VALUES ('" & psMovNro & "','" & psOpeCod & "','" & Replace(psMovDesc, "'", "''") & "'," _
          & pnMovEstado & "," & pnMovFlag & ")"
    
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
    Exit Sub

InsertMovErr:
    Call RaiseError(MyUnhandledError, "DPigActualizaBD Insert Mov")
End Sub

'***************************************************
'** CMCPL - Obtiene el nMovNro a partir del cMovNro
'***************************************************
Public Function dGetnMovNro(ByVal psMovNro As String) As Long
Dim rs As Recordset

lsSQL = "SELECT nMovNro FROM Mov WHERE cMovNro ='" & psMovNro & "'"

Set rs = oConn.CargaRecordSet(lsSQL)

If Not rs.EOF And Not rs.BOF Then
    dGetnMovNro = rs!nMovNro
End If

rs.Close
Set rs = Nothing

End Function

'***************************************************
'CAFF - 22/08/2002 -  Insert MovCol
'***************************************************

Public Sub dInsertMovCol(ByVal pnMovNro As Long, ByVal psOperacion As String, ByVal psCuenta As String, ByVal pnNroCalend As Integer, _
        ByVal pnMonto As Currency, ByVal pnDiasMora As Integer, ByVal psMetLiq As String, ByVal pnPlazo As Integer, _
        ByVal pnCredEstado As ColocEstado, ByVal pnSaldoCap As Currency, Optional pbEjecBatch As Boolean = False)

On Error GoTo ErrInsertMovCol

    lsSQL = "INSERT MovCol (nMovNro,cOpeCod,cCtaCod,nNroCalen,nMonto,nDiasMora,cMetLiq,nPlazo, nCredEstado, nSaldoCap) " _
        & "VALUES (" & pnMovNro & ",'" & psOperacion & "','" & psCuenta & "'," & pnNroCalend & "," _
        & pnMonto & "," & pnDiasMora & ",'" & psMetLiq & "'," & pnPlazo & "," & pnCredEstado & "," & pnSaldoCap & ") "

    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If

Exit Sub

ErrInsertMovCol:
    Call RaiseError(MyUnhandledError, "DPigActualizaBD Insert Mov Col")

End Sub

Public Sub dInsertMovColDet(ByVal pnMovNro As Long, ByVal psOperacion As String, ByVal psctacod As String, pnNroCalend As Integer, _
        ByVal pnConcepto As ColocConcepto, ByVal pnNroCuota As Integer, ByVal pnMonto As Currency, _
        Optional pbEjecBatch As Boolean = False)

On Error GoTo ErrInsertMovColDet
    
    lsSQL = "INSERT MovColDet (nMovNro,cOpeCod,cCtaCod,nNroCalen,nPrdConceptoCod,nNroCuota,nMonto) " _
        & "VALUES (" & pnMovNro & ",'" & psOperacion & "','" & psctacod & "'," & pnNroCalend & "," _
        & pnConcepto & "," & pnNroCuota & "," & pnMonto & " ) "

    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If

Exit Sub

ErrInsertMovColDet:
    Call RaiseError(MyUnhandledError, "DPigActualizaBD Insert MovColDet")

End Sub

Public Sub dInsertProductoBloqueos(ByVal psctacod As String, ByVal pnBlqTpo As Integer, ByVal pnBlqMotivo As String, _
        ByVal psMovNro As String, ByVal psComentario As String, Optional ByVal psMovNroDbl As String = "@", _
        Optional pbEjecBatch As Boolean = False)

    lsSQL = "INSERT ProductoBloqueos (cCtaCod,nBlqTpo,nBlqMotivo,cMovNro,cComentario) " _
        & "VALUES ('" & psctacod & "'," & pnBlqTpo & "," & pnBlqMotivo & ",'" & psMovNro & "','" _
        & psComentario & "') "

    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If

End Sub

Public Sub dInsertMovRef(ByVal pnMovNro As Long, ByVal pnMovNroRef As Long, _
    Optional pbEjecBatch As Boolean = False)
    
On Error GoTo InsertMovRefErr
    
    lsSQL = "INSERT MovRef (nMovNro,nMovNroRef) " _
          & "VALUES (" & pnMovNro & "," & pnMovNroRef & ")"
    
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
    Exit Sub

InsertMovRefErr:
    Call RaiseError(MyUnhandledError, "DPigActualizaBD Insert Mov Ref")
End Sub

'**************************************************
'** CAFF -  03/08/02 - INSERT COLOCPIGEVALCLIENTE
'**************************************************
Public Sub dInsertColocPigEvalCliente(ByVal psCodPers As String, ByVal pnCalifCte As Integer, ByVal pnEvalSBSCte As Integer, _
    Optional pbEjecBatch As Boolean = False)
    
On Error GoTo InsertColocPigEvalCliente
    
    lsSQL = "INSERT ColocPigEvalCliente (cPersCod,cCalifiCliente, cEvalSBSCliente, cCalifiClienteAnt) " _
          & "VALUES ('" & psCodPers & "'," & pnCalifCte & "," & pnEvalSBSCte & "," & pnCalifCte & ")"
    
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
    Exit Sub

InsertColocPigEvalCliente:
    Call RaiseError(MyUnhandledError, "DPigActualizaBD Insert ColocPigEvalCliente")
End Sub

'******************************************
'** CMCPL - UPDATE PRODUCTO  '** CRSF -  28/08/02
'*******************************************
Public Sub dUpdateProducto(ByVal psctacod As String, Optional ByVal pnTasaInteres As Double = -1, _
        Optional ByVal pnSaldo As Currency = -1, Optional ByVal pnPrdEstado As ColocEstado = -1, _
        Optional ByVal psFechaHora As String = "@", Optional ByVal pnTransac As Integer = -1, _
        Optional pbEjecBatch As Boolean = False)

    lsSQL = "UPDATE Producto SET "

    If pnTasaInteres <> -1 Then
        lsSQL = lsSQL & " nTasaInteres = " & pnTasaInteres & ","
    End If
    If pnSaldo <> -1 Then
        lsSQL = lsSQL & " nSaldo = " & pnSaldo & ","
    End If
    If pnPrdEstado <> -1 Then
        lsSQL = lsSQL & " nPrdEstado = " & pnPrdEstado & ","
    End If
    If psFechaHora <> "@" Then
        lsSQL = lsSQL & " dPrdEstado = '" & psFechaHora & "',"
    End If
    If pnTransac = -2 Then ' Si se le envia (-2) aumenta el nro de transaccion en uno
        lsSQL = lsSQL & " nTransacc = nTransacc + 1,"
    ElseIf pnTransac <> -1 Then
        lsSQL = lsSQL & " nTransacc = " & pnTransac & ","
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod ='" & psctacod & "' "
    
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If

End Sub


'************************************************
'** CRSF - 28/02/2002 - UPDATE COLOCACIONES
'************************************************
Public Sub dUpdateColocaciones(ByVal psctacod As String, Optional ByVal pnPlazo As Integer = -1, _
        Optional ByVal psFecVenc As String = "@", Optional ByVal pnMontoCol As Currency = -1, Optional ByVal pnColocCalendCod As Integer = -1, _
        Optional ByVal psUltimaActualizacion As String = "@", Optional ByVal psLineaCred As String = "@", Optional ByVal psFecVig As String = "@", _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "UPDATE Colocaciones SET "

    If pnPlazo <> -1 Then
        lsSQL = lsSQL & " nPlazo = " & pnPlazo & ","
    End If
    If psFecVenc <> "@" Then
        lsSQL = lsSQL & " dVenc = '" & psFecVenc & "',"
    End If
    If psFecVig <> "@" Then
        lsSQL = lsSQL & " dVigencia = '" & psFecVig & "',"
    End If
    If pnMontoCol <> -1 Then
        lsSQL = lsSQL & " nMontoCol = " & pnMontoCol & ","
    End If
    If pnColocCalendCod <> -1 Then
        lsSQL = lsSQL & " nColocCalendCod = " & pnColocCalendCod & ","
    End If
    If psUltimaActualizacion <> "@" Then
        lsSQL = lsSQL & " cUltimaActualizacion = '" & psUltimaActualizacion & "',"
    End If
    If psLineaCred <> "@" Then
        lsSQL = lsSQL & " cLineaCred = '" & psLineaCred & "',"
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod ='" & psctacod & "' "

    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If

End Sub
'************************************************
'** EAFA- 28/02/2002 - UPDATE CALENDARIO
'************************************************
Public Sub dUpdateColocCalendario(ByVal psctacod As String, ByVal pnNroCalend As Integer, ByVal pnColocCalendApl As Integer, _
                                  Optional ByVal pnColocCalendEstado As Integer = -1, Optional ByVal pnColocCalendFlag As Integer = -1, _
                                  Optional ByVal pnCalendProc As Integer = -1, Optional ByVal pdFecPago As String = "", _
                                  Optional pbEjecBatch As Boolean = False)
                                  
Dim lsSQL As String
    lsSQL = "UPDATE ColocCalendario SET "

    If pnColocCalendEstado <> -1 Then
        lsSQL = lsSQL & " nColocCalendEstado = " & pnColocCalendEstado & ","
    End If
    If pnColocCalendFlag <> -1 Then
        lsSQL = lsSQL & " cColocCalendFlag = " & pnColocCalendFlag & ","
    End If
    If pnCalendProc <> -1 Then
        lsSQL = lsSQL & " nCalendProc = " & pnCalendProc & ","
    End If
    If pdFecPago <> "" Then
        lsSQL = lsSQL & " dPago = '" & Format(pdFecPago, "mm/dd/yyyy") & "',"
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE          cCtaCod ='" & psctacod & "'" _
                         & "    AND          nNroCalen = " & pnNroCalend _
                         & "    AND nColocCalendApl = " & pnColocCalendApl

    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If

End Sub
'************************************************
'** CRSF - 28/02/2002 - UPDATE CALENDARIODET
'************************************************
Public Sub dUpdateColocCalendDetPig(ByVal psctacod As String, ByVal pnNroCalend As Integer, ByVal pnColocCalendApl As Integer, _
                                                       ByVal pnColocConcepto As Integer, Optional ByVal pnNewMonto As Currency = -1, _
                                                       Optional ByVal pnMontoPag As Currency = -1, Optional ByVal pnFlag As Integer = -1, _
                                                       Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
    lsSQL = "UPDATE ColocCalendDetPig SET "

    If pnNewMonto <> -1 Then
        lsSQL = lsSQL & " nMonto = " & pnNewMonto & ","
    End If
    If pnMontoPag <> -1 Then
        lsSQL = lsSQL & " nMontoPagado = " & pnMontoPag & ","
    End If
    If pnFlag <> -1 Then
        lsSQL = lsSQL & " cFlag = " & pnFlag & ","
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE     cCtaCod ='" & psctacod & "'" _
                         & "    AND  nNroCalen = " & pnNroCalend _
                         & "    AND  nColocCalendApl = " & pnColocCalendApl _
                         & "    AND nPrdConceptoCod = " & pnColocConcepto

    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If

End Sub

'************************************************
'** CAFF - 05/09/2002 - UPDATE COLOCPIGNO
'************************************************
Public Sub dUpdateColocPigno(ByVal psctacod As String, Optional ByVal pnUsoLinea As Integer = -1, Optional ByVal pnEstadoCont As ColocEstado = -1, _
        Optional pnUbicaLote As Integer = -1, Optional ByVal pnNumCalen As Integer = -1, Optional ByVal pnTipoTasac As Integer = -1, _
        Optional ByVal pnNroDuplic As Integer = -1, Optional ByVal pnNumAmort As Integer = -1, Optional pbEjecBatch As Boolean = False)

    lsSQL = "UPDATE COLOCPIGNO SET "
    
    If pnEstadoCont <> -1 Then
        lsSQL = lsSQL & " nEstadoCont = " & pnEstadoCont & ","
    End If

    If pnUsoLinea <> -1 Then
        lsSQL = lsSQL & " nUsoLineaNro = " & pnUsoLinea & ","
    End If

    If pnUbicaLote <> -1 Then
        lsSQL = lsSQL & " nUbicaLote = " & pnUbicaLote & ","
    End If

    If pnNumCalen <> -1 Then
        lsSQL = lsSQL & " nNumCalend = " & pnNumCalen & ","
    End If

    If pnTipoTasac <> -1 Then
        lsSQL = lsSQL & " nTipoTasacion = " & pnTipoTasac & ","
    End If
    
    If pnNroDuplic <> -1 Then
        lsSQL = lsSQL & " nNroDuplic = " & pnNroDuplic & ","
    End If
       
    If pnNumAmort = -2 Then  ' Si se le envia (-2) aumenta el nro de amortizacion
        lsSQL = lsSQL & " nNroAmort  = pnNumAmort  + 1,"
    ElseIf pnNumAmort <> -1 Then
        lsSQL = lsSQL & " nNroAmort = " & pnNumAmort & ","
    End If
           
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod ='" & psctacod & "'"
    
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
    
End Sub

'***************************************************
'CAFF - 02/09/2002 -  Update Mov
'***************************************************
Public Sub dUpdateMov(ByVal pnMovNro As Long, Optional ByVal psOpeCod As String = "@", _
    Optional ByVal psMovDesc As String = "@", Optional ByVal pnMovEstado As MovEstado = -1, _
    Optional ByVal pnMovFlag As MovFlag = -1, Optional pbEjecBatch As Boolean = False)

    lsSQL = "UPDATE Mov SET "

    If psOpeCod <> "@" Then
        lsSQL = lsSQL & " cOpeCod = '" & psOpeCod & "',"
    End If
    If psMovDesc <> "@" Then
        lsSQL = lsSQL & " cMovDesc = '" & psMovDesc & "',"
    End If
    If pnMovEstado <> -1 Then
        lsSQL = lsSQL & " nMovEstado = " & pnMovEstado & ","
    End If
    If pnMovFlag <> -1 Then
        lsSQL = lsSQL & " nMovFlag = " & pnMovFlag & ","
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE nMovNro =" & pnMovNro & " "

    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If

End Sub
'*******************************************************************************
' CAFF  -  02/09/2002 -  UPDATE PRODUCTOSBLOQUEO PARA DESBLOQUEO DE CONTRATOS
'*******************************************************************************
Public Sub dUpdateProductoBloqueos(ByVal psctacod As String, ByVal psMovNro As String, _
        Optional ByVal psMovNroDbl As String = "@", Optional pbEjecBatch As Boolean = False)

    lsSQL = "UPDATE ProductoBloqueos SET "
    
    If psMovNroDbl <> "@" Then
        lsSQL = lsSQL & " cMovNroDbl = '" & psMovNroDbl & "',"
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod ='" & psctacod & "' AND cMovNro ='" & psMovNro & "' "
    
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If

End Sub

'****************************************************
'** CAFF - 28/02/2002 - UPDATE COLOCPIGJOYATASACION
'****************************************************
 Public Sub dUpdateColocPigJoyaTasacion(ByVal psctacod As String, ByVal pnTipoTasacion As Integer, ByVal pnItemPieza As Integer, _
                           Optional ByVal psMovNro As String = "@", Optional ByVal pnTipoMaterial As Integer = -1, Optional ByVal pnTipoJoya As Integer = -1, _
                           Optional ByVal pnSubTipoJoya As Integer = -1, Optional ByVal pnestJoya As Integer = -1, Optional pnPesoBruto As Double = -1, _
                           Optional pnPesoNeto As Double = -1, Optional pnTasacion As Currency = -1, Optional ByVal pnTasacionAdic As Currency = -1, _
                           Optional ByVal psDescrip As String = "@", Optional ByVal psDescripAdic As String = "@", Optional ByVal pnUbicaPieza As Integer = -1, _
                           Optional ByVal psCodPers As String = "@", Optional ByVal pbEjecBatch As Boolean = False)

    lsSQL = "UPDATE ColocPigJoyaTasacion SET "

    If pnTipoMaterial <> -1 Then
        lsSQL = lsSQL & " nMaterial = " & pnTipoMaterial & ","
    End If
    If pnTipoJoya <> -1 Then
        lsSQL = lsSQL & " nTipoJoya = " & pnTipoJoya & ","
    End If
    If pnSubTipoJoya <> -1 Then
        lsSQL = lsSQL & " nSubTipoJoya = " & pnSubTipoJoya & ","
    End If
    If pnestJoya <> -1 Then
        lsSQL = lsSQL & " nConservacion = " & pnestJoya & ","
    End If
    If pnPesoBruto <> -1 Then
        lsSQL = lsSQL & " nPesoBruto = " & pnPesoBruto & ","
    End If
    If pnPesoNeto <> -1 Then
        lsSQL = lsSQL & " nPesoNeto = " & pnPesoNeto & ","
    End If
    If pnTasacion <> -1 Then
        lsSQL = lsSQL & " nTasacion = " & pnTasacion & ","
    End If
    If pnTasacionAdic <> -1 Then
        lsSQL = lsSQL & " nTasacionAdicional = " & pnTasacionAdic & ","
    End If
    If psDescrip <> "@" Then
        lsSQL = lsSQL & " cDescripcion = '" & psDescrip & "',"
    End If
    If psDescripAdic <> "@" Then
        lsSQL = lsSQL & " cDescripcionAdic = '" & psDescripAdic & "',"
    End If
    If pnUbicaPieza <> -1 Then
        lsSQL = lsSQL & " nUbicaPieza = " & pnUbicaPieza & ","
    End If
    If psCodPers <> "@" Then
        lsSQL = lsSQL & " cPersCod = '" & psCodPers & "',"
    End If
    
    If psMovNro <> "@" Then
        lsSQL = lsSQL & " cUltimaActualizacion = '" & psMovNro & "',"
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cCtaCod = '" & psctacod & "'" _
                         & " AND nTipoTasacion = " & pnTipoTasacion _
                         & " AND nItemPieza = " & pnItemPieza

    oConn.Ejecutar lsSQL


End Sub

'******************************************************
'** CAFF - 07/09/2002 - INSERT COLOCPIGPRECIOMATERIAL
'******************************************************
Public Sub dInsertColocPigPrecioMaterial(ByVal pnMaterial As Integer, ByVal pnMoneda As Integer, ByVal pnTipoValor As Integer, _
        ByVal psFecha As String, ByVal pnValor As Currency, Optional pbEjecBatch = False)

On Error GoTo ErrInsertTarifario

    lsSQL = "INSERT COLOCPIGPRECIOMATERIAL(nMaterial, nMoneda, nTipoValor, dInicial, nValor) VALUES ( " _
        & pnMaterial & "," & pnMoneda & "," & pnTipoValor & ",'" & Format(psFecha, "mm/dd/yyyy") & "'," & pnValor & ")"

    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
    
    Exit Sub
    
ErrInsertTarifario:
    If Err.Number = -2147211505 Then
       MsgBox "Dato ya existe", vbInformation, "Aviso"
    End If
End Sub
'********************************************************
'** CAFF - 09/09/2002 - DELETE COLOCPIGPRECIOMATERIAL
'********************************************************
Public Sub dDeleteColocPigPrecioMaterial(ByVal pnProceso As Integer, ByVal pnMaterial As Integer, ByVal pnMoneda As Integer, _
        ByVal psFecha As String, Optional pbEjecBatch = False)

    lsSQL = "DELETE ColocPigPrecioMaterial WHERE nTipoValor = " & pnProceso & " AND nMaterial =" & pnMaterial & " AND nMoneda = " _
            & pnMoneda & " AND datediff(dd,dInicial, '" & Format(psFecha, "mm/dd/yyyy") & "') = 0"
            
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
    
    Exit Sub
    
End Sub
'********************************************************
'** CAFF - 09/09/2002 - UPDATE COLOCPIGPRECIOMATERIAL
'********************************************************
Public Sub dUpdateColocPigPrecioMaterial(ByVal pnProceso As Integer, ByVal pnMaterial As Integer, ByVal pnMoneda As Integer, _
        ByVal psFecha As String, ByVal pnValor As Currency, ByVal pnMonedaAnt As Integer)

    lsSQL = "Update ColocPigPrecioMaterial Set nValor = " & pnValor & ", nMoneda = " & pnMoneda _
            & " WHERE nTipoValor = " & pnProceso & " AND nMaterial = " & pnMaterial _
            & " AND nMoneda = " & pnMonedaAnt & " AND datediff(dd, dInicial, '" & Format(psFecha, "mm/dd/yyyy") & "') = 0"

    oConn.Ejecutar lsSQL
    Exit Sub
    
End Sub
'***********************************************************************
' CRSF - 10/09/02 - UPDATE COLOCACESTADO PARA EL COBRO DE LA COMISION
'***********************************************************************
Public Function dUpdateColocEstado(ByVal psCodigo As String, ByVal pnFechaAct As String, _
                    ByVal pnEstado As ColocEstado, Optional pbEjecBatch As Boolean = False)
  
  lsSQL = "UPDATE COLOCACESTADO SET "
   lsSQL = lsSQL & " dPrdEstado= '" & pnFechaAct & "',"
   lsSQL = Left(lsSQL, Len(lsSQL) - 1)
   lsSQL = lsSQL & " WHERE cCtaCod  ='" & psCodigo & "'  and nPrdEstado = " & pnEstado & ""
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
End Function

'****************************************************************
' CRSF - 03/09/02 - UPDATE EL COLOPIGEVALCLIENTE  EL DATO ACTUAL
'****************************************************************
Public Function dUpdateEval(ByVal psCodigo As String, ByVal pnCalif As Integer, Optional pbEjecBatch As Boolean = False)
    
    lsSQL = "UPDATE COLOCPIGEVALCLIENTE SET "
    'Actualizacion de Calendario
    lsSQL = lsSQL & " cCalifiClienteAnt = cCalifiCliente ,"
    lsSQL = lsSQL & " cCalifiCliente       = '" & pnCalif & "',"
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cPersCod  ='" & psCodigo & "'"
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If

End Function

'*********************************************
' EAFA - 03/09/02 - DELETE COLOC CALENDARIO
'*********************************************

Public Sub dDeleteColocCalendario(ByVal psCuenta As String, ByVal pnNroCalen As Integer, _
                                                  ByVal pnColocCalendApl As ColocCalendApl, Optional ByVal pnCuota As Integer = -1, _
                                                  Optional ByVal pbEjecBatch As Boolean = False)
    
    lsSQL = " DELETE ColocCalendario " _
             & " WHERE         cCtaCod = '" & psCuenta & "' " _
             & "   AND          nNroCalen = " & pnNroCalen & " " _
             & "   AND nColocCalendApl = " & pnColocCalendApl
             
    If pnCuota <> -1 Then
       lsSQL = lsSQL & " AND nCuota = " & pnCuota
    End If
    
    oConn.Ejecutar lsSQL
    
End Sub

'***************************************************
' EAFA - 03/09/02 - DELETE COLOC CALENDARIO DET
'***************************************************
Public Sub dDeleteColocCalendDetPig(ByVal psCuenta As String, ByVal pnNroCalen As Integer, _
                                                      ByVal pnColocCalendApl As ColocCalendApl, Optional ByVal pnCuota As Integer = -1, _
                                                      Optional ByVal pnConcepto As ColocConcepto = -1, Optional ByVal pbEjecBatch As Boolean = False)

    lsSQL = " DELETE ColocCalendDetPig " _
             & " WHERE         cCtaCod = '" & psCuenta & "' " _
             & "   AND          nNroCalen = " & pnNroCalen & " " _
             & "   AND nColocCalendApl = " & pnColocCalendApl
             
    If pnCuota <> -1 Then
       lsSQL = lsSQL & " AND nCuota = " & pnCuota
    End If
    
    If pnConcepto <> -1 Then
       lsSQL = lsSQL & " AND nPrdConceptoCod = " & pnConcepto
    End If
    
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
    
End Sub

'*****************************************************
' EAFA - 03/09/02 - DELETE COLOC PIG JOYA TASACION
'*****************************************************
Public Sub dDeleteColocPigJoyaTasacion(ByVal psctacod As String, ByVal pnTipoTasacion As ColocPigTipoTasacion, _
                                       Optional ByVal pnItemPieza As Integer = -1, Optional ByVal pbEjecBatch As Boolean = False)
    
    On Error GoTo ErrordDeleteColocPigJoyaTasacion
    
    lsSQL = " DELETE ColocPigJoyaTasacion  " _
             & " WHERE      cCtaCod = '" & psctacod & "'" _
             & "    AND nTipoTasacion = " & pnTipoTasacion
    
    If pnItemPieza <> -1 Then
       lsSQL = lsSQL & " AND nItemPieza = " & pnItemPieza
    End If
    
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
    
Exit Sub
    
ErrordDeleteColocPigJoyaTasacion:
    Err.Raise Err.Number, "Error En Proceso dDeleteColocPigJoyaTasacion", Err.Description
    
End Sub

'*********************************************
'** CAFF - 18/09/2002 - INSERT COLOCPIGGUIA
'*********************************************
Public Sub dInsertColocPigGuia(ByVal psNumDoc As String, ByVal pnMotivo As Integer, ByVal pnOrigen As Integer, _
        ByVal pnDestino As Integer, ByVal pnTipoGuia As Integer, ByVal pnTotalItem As Long, ByVal pnPBruto As Double, _
        ByVal pnPNeto As Double, ByVal pnTotalTasac As Currency, ByVal pnValorAdj As Currency, ByVal psMov As String, _
        Optional ByVal pbEjecBatch As Boolean = False)

    lsSQL = "INSERT COLOCPIGGUIA (cNumDoc, cFlag, nMotivo, nOrigen, nDestino, nTipoGuia, nTotItem, nPesoBruto, " _
            & "nPesoNeto, nTotTasacion, nValorAdju, cUltimaActualizacion) " _
            & "VALUES ('" & psNumDoc & "',1," & pnMotivo & "," & pnOrigen & "," & pnDestino & ", " & pnTipoGuia _
                        & "," & pnTotalItem & "," & pnPBruto & "," & pnPNeto & "," & pnTotalTasac & "," & pnValorAdj & ",'" _
                        & psMov & "')"

    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
       
End Sub

'***********************************************
'** CAFF - 18/09/2002 - INSERT COLOCPIGGUIADET
'***********************************************
Public Sub dInsertColocPigGuiaDet(ByVal psNumDoc As String, ByVal pnItemGuia As Integer, ByVal psctacod As String, _
        ByVal pnItemPieza As Integer, ByVal pnTotalPiezas As Integer, Optional ByVal pbEjecBatch = False)

    lsSQL = "INSERT COLOCPIGGUIADET(cNumDoc, nItemGuia, cCtaCod, nItemPieza, nTotPiezasCont, cConfirma) " _
            & "VALUES ('" & psNumDoc & "'," & pnItemGuia & ",'" & psctacod & "'," & pnItemPieza & "," _
            & pnTotalPiezas & ",0)"
            
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
    
End Sub

'************************************************
'** CAFF - 19/09/2002 - INSERT COLOCPIGGUIAETAPA
'************************************************
Public Sub dInsertColocPigGuiaEtapa(ByVal psNumDoc As String, ByVal pnEtapa As Integer, ByVal psNumMov As String, _
            Optional ByVal pbEjecBatch = False)

    lsSQL = "INSERT COLOCPIGGUIAETAPA (cNumDoc, nEtapa, cUltimaActualizacion) " _
                & "VALUES ('" & psNumDoc & "'," & pnEtapa & ",'" & psNumMov & "')"
    
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
        
End Sub

'********************************************
'** CAFF - 19/09/2002 - UPDATE COLOCPIGGUIA
'********************************************
Public Sub dUpdateColocPigGuia(ByVal psNumDoc As String, Optional ByVal psFlag As String = "@", _
                Optional ByVal pnMotivo As Integer = -1, Optional ByVal pnOrigen As Integer = -1, _
                Optional ByVal pnDestino As Integer = -1, Optional ByVal psPersCod As String = "@", _
                Optional ByVal pnTipoGuia As Integer = -1, Optional ByVal pnTotalItem As Integer = -1, _
                Optional ByVal pnPBruto As Double = -1, Optional pnPNeto As Double = -1, Optional pnTasacion As Currency = -1, _
                Optional ByVal pnValorAdj As Currency = -1, Optional ByVal psSerieSunat As String = "@", _
                Optional ByVal psNumDocSunat As String = "@", Optional ByVal psMovNro As String = "@", _
                Optional ByVal pbEjecBatch As Boolean = False)
    
    lsSQL = "UPDATE ColocPigGuia SET "

    If psFlag <> "@" Then
        lsSQL = lsSQL & " cFlag = '" & psFlag & "',"
    End If
    If pnMotivo <> -1 Then
        lsSQL = lsSQL & " nMotivo = " & pnMotivo & ","
    End If
    If pnOrigen <> -1 Then
        lsSQL = lsSQL & " nOrigen = " & pnOrigen & ","
    End If
    If pnDestino <> -1 Then
        lsSQL = lsSQL & " nDestino= " & pnDestino & ","
    End If
    If psPersCod <> "@" Then
        lsSQL = lsSQL & " cPersCod = '" & psPersCod & "',"
    End If
    If pnTipoGuia <> -1 Then
        lsSQL = lsSQL & " nTipoGuia = " & pnTipoGuia & ","
    End If
    If pnTotalItem <> -1 Then
        lsSQL = lsSQL & " nTipItem = " & pnTasacion & ","
    End If
    If pnPBruto <> -1 Then
        lsSQL = lsSQL & " nPesoBruto  = " & pnPBruto & ","
    End If
    If pnPNeto <> -1 Then
        lsSQL = lsSQL & " nPesoNeto = " & pnPNeto & ","
    End If
    If pnTasacion <> -1 Then
        lsSQL = lsSQL & " nTotTasacion = '" & pnTasacion & ","
    End If
    If pnValorAdj <> -1 Then
        lsSQL = lsSQL & " nValorAdj = " & pnValorAdj & ","
    End If
    If psSerieSunat <> "@" Then
        lsSQL = lsSQL & " cSerDocSunat = '" & psSerieSunat & "',"
    End If
    If psNumDocSunat <> "@" Then
        lsSQL = lsSQL & " cNumDocSunat = '" & psNumDocSunat & "',"
    End If
    If psMovNro <> "@" Then
        lsSQL = lsSQL & " cUltimaActualizacion = '" & psMovNro & "',"
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cNumDoc = '" & psNumDoc & "'"

    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If

End Sub

'***********************************************
'** CAFF - 19/09/2002 - UPDATE COLOCPIGGUIADET
'***********************************************
Public Sub dUpdateColocPigGuiaDet(ByVal psNumDoc As String, Optional ByVal pnItemGuia As Integer = -1, _
                Optional ByVal psctacod As String = "@", Optional ByVal pnItemPieza As Integer = -1, _
                Optional ByVal pnTotalPiezas As Integer = -1, Optional ByVal psConfirma As String = "@", _
                Optional ByVal pbEjecBatch As Boolean = False)
    
    lsSQL = "UPDATE ColocPigGuiaDet SET "

    If psctacod <> "@" Then
        lsSQL = lsSQL & " cCtaCod = '" & psctacod & "',"
    End If
    If pnItemPieza <> -1 Then
        lsSQL = lsSQL & " nItemPieza = " & pnItemPieza & ","
    End If
    If pnTotalPiezas <> -1 Then
        lsSQL = lsSQL & " nTotPiezasCont = " & pnTotalPiezas & ","
    End If
    If psConfirma <> "@" Then
        lsSQL = lsSQL & " cConfirma = '" & psConfirma & "',"
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cNumDoc = '" & psNumDoc & "'"
    
    If pnItemGuia <> -1 Then
        lsSQL = lsSQL & " AND nItemGuia = " & pnItemGuia
    End If

    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If

End Sub

'******************************************
'** CAFF - 24/09/2002 - UPDATE COLOCPIGNO
'******************************************
Public Sub dUpdateColocPignoLote(ByVal psNumDoc As String, ByVal pnUbicaLote As Integer)

    
On Error GoTo dErrorInsert

    lsSQL = " UPDATE ColocPigno SET nUbicaLote = " & pnUbicaLote & "WHERE cCtaCod IN " _
            & "(SELECT cCtaCod FROM ColocPigGuiaDet WHERE cNumDoc = '" & psNumDoc & "' AND cCtaCod = ColocPigno.cCtaCod)"

    oConn.Ejecutar lsSQL
    
    Exit Sub

dErrorInsert:

    Err.Raise Err.Number, "Error en dUpdateColocPignoLote", Err.Description
    
End Sub

'**************************************************************************************************************************
'** CAFF - 24/09/2002 - UPDATE COLOCPIGJOYATASACION cUbicaPieza en Transito para el tema de los despachos de una remesa
'**************************************************************************************************************************
Public Sub dUpdateColocPigJoyaTasacionPieza(ByVal psNumDoc As String, ByVal pnUbicaLote As Integer, ByVal pnTipoGuia As Integer)
    
On Error GoTo dErrorInsert

    If pnTipoGuia = 1 Then
        lsSQL = "UPDATE ColocPigJoyaTasacion SET nUbicaPieza = " & pnUbicaLote & " WHERE cCtaCod IN " _
            & "(SELECT cCtaCod FROM ColocPigGuiaDet WHERE cNumDoc = '" & psNumDoc & "')"
    Else
        lsSQL = "UPDATE ColocPigJoyaTasacion SET nUbicaPieza = " & pnUbicaLote & " WHERE cCtaCod IN " _
            & "(SELECT cCtaCod FROM ColocPigGuiaDet WHERE cNumDoc = '" & psNumDoc & "' AND nItemPieza = ColocPigJoyaTasacion.nItemPieza)"
    End If
    
    oConn.Ejecutar lsSQL
    
    Exit Sub

dErrorInsert:

    Err.Raise Err.Number, "Error en dUpdateColocPigJoyaTasacionPieza", Err.Description
    
End Sub

'**************************************************************
'** CAFF - 25/09/2002 - Elimina los datos de la Tabla Temporal
'**************************************************************
Public Sub dDeletePiezasAdjColocPigGuiaDet(ByVal pnOrigen As Integer, ByVal pnDestino As Integer)
Dim sql As String
On Error GoTo dError
sql = "DELETE TMPCOLOCPIGPIEZAS WHERE nOrigen = " & pnOrigen & " AND nDestino = " & pnDestino

    oConn.Ejecutar sql
    
    Exit Sub
    
dError:

    Err.Raise Err.Number, "Elimina Datos en dDeletePiezasAdjColocPigGuiaDet ", Err.Description
    
End Sub
'**************************************************************
'** CAFF - 27/09/2002 - Inserta en la Tabla ColocPigProceso
'**************************************************************
Public Sub dInsertColocPigProceso(ByVal pnRemate As Integer, ByVal pnTipoProceso As Integer, ByVal psctacod As String, _
            ByVal pnItemPieza As Integer, ByVal pnValorProceso As Currency, ByVal pnValorDeuda As Currency, ByVal pnSituaPieza As Integer, _
            ByVal pnValorDeudaCompen As Currency, ByVal pnValorVenta As Currency, ByVal psMovNro As String)

Dim sql As String

    On Error GoTo dInsertError
    
    sql = "INSERT COLOCPIGPROCESO (nRemate, nTipoProceso, cCtaCod, nItemPieza, nValorProceso, nValorDeuda, nSituacionPieza, " _
                                & "nValorDeudaComp, nValorVenta, cUltimaActualizacion) " _
                & "VALUES (" & pnRemate & "," & pnTipoProceso & ",'" & psctacod & "'," & pnItemPieza & "," & pnValorProceso & "," _
                            & pnValorDeuda & "," & pnSituaPieza & "," & pnValorDeudaCompen & "," & "," & pnValorVenta & ",'" & _
                            psMovNro & "')"
    
    oConn.Ejecutar sql

    Exit Sub
    
dInsertError:

    Err.Raise Err.Number, "Insertar en ColocPigProceso <<dInsertColocPigProceso>>", Err.Description
    
End Sub

'**************************************************************
'** CAFF - 27/09/2002 - Inserta en la Tabla ColocPigProcesoDet
'**************************************************************
Public Sub dInsertColocPigProcesoDet(ByVal pnRemate As Integer, ByVal pnTipoProceso As Integer, ByVal psctacod As String, _
            ByVal pnItemPieza As Integer, ByVal pnTipoAplica As Integer, ByVal pnCodConcepto As Integer, ByVal pnMonto As Currency)
Dim sql As String

    On Error GoTo dInsertError
    
    sql = "INSERT COLOCPIGPROCESODET (nRemate, nTipoProceso, cCtaCod, nItemPieza, nTipAplica, nCodConcepto, nMonto) " _
                & "VALUES (" & pnRemate & "," & pnTipoProceso & ",'" & psctacod & "'," & pnItemPieza & "," & pnTipoAplica & "," _
                            & pnCodConcepto & "," & pnMonto & ")"
    
    oConn.Ejecutar sql

    Exit Sub
    
dInsertError:

    Err.Raise Err.Number, "Insertar en ColocPigProcesoDet <<dInsertColocPigProcesoDet>>", Err.Description
    
End Sub
'**************************************************************
'** CAFF - 27/09/2002 - Inserta en la Tabla ColocPigTroy
'**************************************************************
Public Sub dInsertColocPigTroy(ByVal pnRemate As Integer, ByVal psctacod As String, ByVal pnItemPieza As Integer, ByVal pnValTroy As Currency)
Dim sql As String
    On Error GoTo dInsertError
    
    sql = "INSERT COLOCPIGTROY (nRemate, cCtaCod, nItemPieza, nValTroyMin) " _
                & "VALUES (" & pnRemate & ",'" & psctacod & "'," & pnItemPieza & "," & pnValTroy & ")"
    
    oConn.Ejecutar sql

    Exit Sub
    
dInsertError:

    Err.Raise Err.Number, "Insertar en ColocPigTroy <<dInsertColocPigTroy>>", Err.Description

End Sub

'***********************************************
' CRSF - 23/09/02 INSERTA UN REGISTRO DE REMATE
'***********************************************
Public Sub dInsertRemate(ByVal pnRemate As Long, ByVal pnTipoProceso As Integer, ByVal pnUbicacion As Long, _
                                             ByVal pdInicio As Date, pdFin As Date, ByVal pdReferencia As Date, _
                                             ByVal psCodigoPer As String, ByVal pdAviso As Date)

On Error GoTo ErrordInsertRemate

    lsSQL = "INSERT ColocPigRemate(nRemate,nTipoProceso,cUbicacion,dInicio,dFin, dReferencia,  " _
                   & " cPersCod, dAviso) " _
            & "VALUES (" & pnRemate & "," & pnTipoProceso & "," & pnUbicacion & " , '" & Format(pdInicio, "m/dd/yyyy") _
            & "' , '" & Format(pdFin, "mm/dd/yyyy") & "', '" & Format(pdReferencia, "mm/dd/yyyy") _
            & "','" & psCodigoPer & "','" & Format(pdAviso, "mm/dd/yyyy") & "') "
    
    oConn.Ejecutar lsSQL
    
Exit Sub
 
ErrordInsertRemate:
    Err.Raise Err.Number, "Error En Proceso dInsertaRemate ", Err.Description

End Sub

'******************************************************************************************
' CRSF - 25/09/02 - INSERTAR REGISTRO DE SELECCION CON TIPO TASACION 3 -RETASACION MANUAL
'******************************************************************************************
Public Sub dInsertaRegistroSel(ByVal pdFecIniRem As String, ByVal pnDiasVenc As Integer, ByVal psFecAct As String)

On Error GoTo ErrordRegistroSeleccion
     
     lsSQL = " INSERT COLOCPIGJOYATASACION " & _
                    " Select J.cCtaCod, nTipoTasacion = 3, nItemPieza, nMaterial, nTipoJoya, nSubTipoJoya, nConservacion, " & _
                    "  nPesoBruto,nPesoNeto,nTasacion, nTasacionAdicional,cDescripcion, cDescripcionAdic, nUbicaPieza, " & _
                    " cPersCod, cUltimaActualizacion = '" & psFecAct & "'" & _
                    " FROM COLOCPIGJOYATASACION J INNER JOIN COLOCPIGNO C ON J.cCtaCod = C.cCtaCod " & _
                    " WHERE J.nTipoTasacion = C.nTipoTasacion " & _
                    " AND J.cCtaCod In (" _
                        & " SELECT C.cCtaCod " _
                        & "FROM COLOCPIGNO CP INNER JOIN COLOCACIONES C ON CP.CCTACOD = C.CCTACOD " _
                                                & "INNER JOIN PRODUCTO P ON P.CCTACOD = C.CCTACOD " _
                        & "WHERE DATEDIFF(dd, dVenc, '" & Format(pdFecIniRem, "mm/dd/yyyy") & "') >= " & pnDiasVenc _
                        & " AND nPrdEstado IN (" & gPigEstAmortiz & "," & gPigEstDesemb & "," & gPigEstReusoLin & "))" _
                   & "AND J.cCtaCod Not In " _
                        & "(SELECT DISTINCT(cCtaCod) FROM COLOCPIGJOYATASACION WHERE nTipoTasacion = 3)"
                
        
        oConn.Ejecutar lsSQL
Exit Sub
 
ErrordRegistroSeleccion:
    Err.Raise Err.Number, "Error En Proceso ErrordregistroSeleccion", Err.Description

End Sub

Public Sub dUpdateColocPignoSel(ByVal pnTipoTasac As Integer, ByVal pdFecIniRem As String, ByVal pnDiasVenc As Integer)

    lsSQL = "UPDATE COLOCPIGNO SET nTipoTasacion = " & pnTipoTasac & "" _
        & " WHERE cCtaCod IN (" _
        & " SELECT CP.cCtaCod " _
                & "FROM COLOCPIGNO CP INNER JOIN COLOCACIONES C ON CP.CCTACOD = C.CCTACOD " _
                & "INNER JOIN PRODUCTO P ON P.CCTACOD = CP.CCTACOD " _
                & "WHERE DATEDIFF(dd, dVenc, '" & Format(pdFecIniRem, "mm/dd/yyyy") & "') >= " & pnDiasVenc _
                & " AND nPrdEstado IN (" & gPigEstAmortiz & "," & gPigEstDesemb & "," & gPigEstReusoLin & "))" _
    
    oConn.Ejecutar lsSQL
    
End Sub

Public Sub dInsertaSobrante(ByVal psCodCta As String, ByVal pnRemate As Integer, ByVal psCtaAhorro As String, _
            ByVal pnSobrante As Currency, ByVal pnMovNro As Long, ByVal pnEstadoSob As Integer)

On Error GoTo ErrordInsertaSobrante
    
    lsSQL = "INSERT COLOCPIGSOBRANTE(cCtaCod, nRemate, cCtaAbono, nEstadoSobrante, nSobrante, nMovNro) " _
        & "VALUES ('" & psCodCta & "'," & pnRemate & ",'" & psCtaAhorro & "'," & pnEstadoSob & "," & pnSobrante & "," & pnMovNro & ")"
        
    oConn.Ejecutar lsSQL
    
Exit Sub

ErrordInsertaSobrante:
    Err.Raise Err.Number, "Error En Proceso ErrordInsertaSobrante", Err.Description
        
End Sub
'*************************
'CAFF - 17/10/2002
'*************************
Public Function ValidaSiCtaAhorro(ByVal psCodPers As String) As String
Dim rs As Recordset
Dim sql As String

    sql = "SELECT cCtaAbono FROM COLOCPIGSOBRANTE S INNER JOIN PRODUCTO P ON S.cCtaAbono = P.cCtaCod " _
        & "WHERE nPrdEstado NOT IN  (1300, 1301, 1400, 1401) WHERE cPersCod = '" & psCodPers & "'"
        
Set rs = oConn.CargaRecordSet(sql)
        
If Not rs.EOF And Not rs.BOF Then
    ValidaSiCtaAhorro = IIf(IsNull(rs!cCtaAbono), "", rs!cCtaAbono)
Else
    ValidaSiCtaAhorro = rs!cCtaAbono
End If

Set rs = Nothing

End Function

'*********************************************************************************
' CRSF - 16/10/2002 - ACTUALIZA EL CONTRATO POR PIEZAS, LA SITUACION DE LA PIEZA
'*********************************************************************************
Public Sub dUpdateSituacionPieza(ByVal psCuenta As String, ByVal pnItemPieza As Integer, ByVal psSituaPieza As ColocPigSituacionPieza, _
                                 ByVal pnTipoProceso As Integer, ByVal psCliente As String, ByVal pnValorVta As Currency)
    
On Error GoTo dErrorInsert

    lsSQL = " UPDATE ColocPigProceso SET nSituacionPieza =" & psSituaPieza & " , cComprador = '" & psCliente & "', nValorVenta = " _
            & pnValorVta & "" _
            & " WHERE  cCtaCod = '" & psCuenta & "' AND nItemPieza = " & pnItemPieza & " and nTipoProceso = " & pnTipoProceso & "  "

    oConn.Ejecutar lsSQL
    
    Exit Sub

dErrorInsert:
    Err.Raise Err.Number, "Error en dUpdateColocPignoLote", Err.Description
    
End Sub

'***********************************
'* CMPCL -  Insert MovDocPig '* EAFA - 09/10/2002
'***********************************
Public Sub dInsertMovDocPig(ByVal pnTipDoc As Integer, ByVal psDocumento As String, ByVal pnMovNro As Long, _
                                                  ByVal pnTotVenta As Currency, ByVal psCliente As String, ByVal psMotivo As String, _
                                                  ByVal pnMoneda As Integer, ByVal psFechaHora, Optional ByVal pbEjecBatch As Boolean = False)

    On Error GoTo ErrordInsertMovDocPig
    
    lsSQL = "INSERT MovDocPig (nCodTipo, cDocumento, nNroMov, nMonto, cPersCod, nMotivo, nMoneda, cUltimaActualizacion ) VALUES (" & _
        pnTipDoc & ",'" & psDocumento & "'," & pnMovNro & "," & pnTotVenta & ",'" & psCliente & "'," & psMotivo & "," & _
        pnMoneda & ",'" & psFechaHora & "')"
    
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
    
Exit Sub

ErrordInsertMovDocPig:
    Err.Raise Err.Number, "Error En Proceso dInsertMovDocPig", Err.Description
    
End Sub

'***************************************************
'* CMPCL -  Insert MovDocPigDet * EAFA - 09/10/2002
'***************************************************
Public Sub dInsertMovDocPigDet(ByVal pnMovNro As Long, ByVal pnTipDoc As Integer, ByVal psDocumento As String, _
                                                        ByVal pnItemDoc As Integer, ByVal psctacod As String, ByVal pnItemPieza As Integer, _
                                                        ByVal pnConcepto As Integer, ByVal pnMonto, Optional ByVal pbEjecBatch As Boolean = False)

    On Error GoTo ErrordInsertMovDocPigDet
    
    lsSQL = "INSERT MovDocPigDet (nNroMov, nCodTipo, cDocumento,  nItemDoc, cCtaCod, nPieza, " _
                & " nCodConcepto, nMonto ) VALUES (" & pnMovNro & "," & pnTipDoc & ",'" & psDocumento & "'," _
                & pnItemDoc & ",'" & psctacod & "'," & pnItemPieza & "," & pnConcepto & "," & pnMonto & ")"
    
    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
    
Exit Sub

ErrordInsertMovDocPigDet:
    Err.Raise Err.Number, "Error En Proceso dInsertMovDocPigDet", Err.Description
    
End Sub

'************************************************
'** CMCPL - Update ColocPigProceso
'** EAFA - 09/10/2002
'************************************************
Public Sub dUpdateColocPigProceso(ByVal pnRemate As Integer, ByVal pnTipoProceso As Integer, ByVal psctacod As String, _
                                         ByVal pnItemPieza As Integer, ByVal psFechaHora As String, Optional ByVal pnValProceso As Currency = -1, _
                                         Optional ByVal pnValDeuda As Currency = -1, Optional ByVal pnSituacionPieza As Currency = -1, _
                                         Optional ByVal pnValDeudaComp As Currency = -1, Optional ByVal pnValVenta As Currency = -1, _
                                         Optional ByVal pnValAdicional As Currency = -1, Optional pbEjecBatch As Boolean = False)
Dim lsSQL As String
    lsSQL = "UPDATE ColocPigProceso SET "

    If pnValProceso <> -1 Then
        lsSQL = lsSQL & " nValorProceso = " & pnValProceso & ","
    End If
    If pnValDeuda <> -1 Then
        lsSQL = lsSQL & " nValorDeuda = " & pnValDeuda & ","
    End If
    If pnSituacionPieza <> -1 Then
        lsSQL = lsSQL & " nSituacionPieza = " & pnSituacionPieza & ","
    End If
    If pnValDeudaComp <> -1 Then
        lsSQL = lsSQL & " nValorDeudaComp = " & pnValDeudaComp & ","
    End If
    If pnValVenta <> -1 Then
        lsSQL = lsSQL & " nValorVenta = " & pnValVenta & ","
    End If
    If pnValAdicional <> -1 Then
        lsSQL = lsSQL & " nValorAdicional = " & pnValAdicional & ","
    End If
    
    lsSQL = lsSQL & " cUltimaActualizacion = '" & psFechaHora & "'"
    lsSQL = lsSQL & " WHERE        nRemate = " & pnRemate _
                                 & "    AND nTipoProceso = " & pnTipoProceso _
                                 & "    AND           cCtaCod = '" & psctacod & "'" _
                                 & "    AND       nItemPieza = " & pnItemPieza

    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If

End Sub

'**************************************************************
'** CMCPL - Update a la tabla ColocPigEstadoJoya
'** EAFA - 02/11/2002
'**************************************************************
Public Function dUpdateColocPigEstadoJoya(ByVal psctacod As String, ByVal pnItemPieza As Long, ByVal psFechaHora As String, _
                                                                                  Optional ByVal pnRemate As Long = -1, Optional ByVal pnEstadoJoya As Integer = -1, _
                                                                                  Optional ByVal psFlag As String = "@", Optional pbEjecBatch = False)
Dim lsSQL As String

    lsSQL = "UPDATE ColocPigEstadoJoya SET "

    If pnRemate <> -1 Then
        lsSQL = lsSQL & " nRemate = " & pnRemate & ","
    End If
    If pnEstadoJoya <> -1 Then
        lsSQL = lsSQL & " nEstadoJoya = " & pnEstadoJoya & ","
    End If
    If psFlag <> "@" Then
        lsSQL = lsSQL & " cFlag = " & psFlag & ","
    End If
    
    lsSQL = lsSQL & " cUltimaActualizacion = '" & psFechaHora & "'"
    lsSQL = lsSQL & " WHERE     cCtaCod = " & psctacod _
                                 & "    AND nItemPieza = " & pnItemPieza

    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If

End Function

Public Sub dUpdateColocPigRemate(ByVal pnRemate As Long, ByVal pnUbica As Integer, _
        ByVal pdFecIni As String, ByVal pdFecFin As String, ByVal pdFecRef As String, ByVal pdFecAviso As String, _
        ByVal psMartillero As String)

    lsSQL = "UPDATE COLOCPIGREMATE SET cUbicacion = " & pnUbica & "," _
            & "dInicio = '" & Format(pdFecIni, "mm/dd/yyyy") & "', dFin = '" & Format(pdFecFin, "mm/dd/yyyy") _
            & "', dReferencia = '" & Format(pdFecRef, "mm/dd/yyyy") & "', dAviso = '" & Format(pdFecAviso, "mm/dd/yyyy") _
            & "', cPersCod = '" & psMartillero & "' WHERE nRemate = " & pnRemate
    
    oConn.Ejecutar lsSQL
    
End Sub

Public Function dGetColocPigProcesoDet(ByVal psCodCta As String, ByVal pnPieza As Integer, _
                                        Optional ByVal pnRemate As Integer) As Recordset
                                        
Dim rs As Recordset

    On Error GoTo dError
    
    lsSQL = "SELECT nCodConcepto, nMonto FROM COLOCPIGPROCESODET CRD " _
        & "INNER JOIN COLOCPIGPROCESO CR ON CRD.nRemate = CR.nRemate AND CR.nTipoProceso = CRD.nTipoProceso " _
        & "AND CRD.cCtaCod = CR.cCtaCod  AND CR.nItemPieza = CRD.nItemPieza " _
        & "WHERE nSituacionPieza = " & gPigSituacionPendFacturar & " AND CR.cCtaCod = '" & psCodCta & "' " _
        & "AND CR.nItemPieza = " & pnPieza & " AND CR.nRemate = " & pnRemate
    
    Set rs = oConn.CargaRecordSet(lsSQL)
    Set dGetColocPigProcesoDet = rs
    
    Exit Function
        
dError:
    Err.Raise Err.Number, "Obtiene detalle de Remate <<dObtieneColocPigProcesoDet>>", Err.Description
    
End Function

Public Sub dDeleteColocEstadoExtorno(ByVal psCodCta As String, ByVal pnEstado As Integer)

    On Error GoTo dError
    
    lsSQL = " DELETE ColocacEstado WHERE cCtaCod = '" & psCodCta & "' AND nPrdEstado = " & pnEstado _
            & " AND dPrdEstado = " _
            & "(SELECT Max(dPrdEstado) FROM ColocacEstado WHERE cCtaCod = '" & psCodCta & "' AND nPrdEstado = " & pnEstado & ")"

    oConn.Ejecutar lsSQL
    
    Exit Sub
        
dError:
    Err.Raise Err.Number, "Elimina Estados de un Contrato <<dDeleteColocEstadoExtorno>>", Err.Description

End Sub


'***********************************************************************
' DFAO - 03/07/02 - INSERTA REGISTRO DE CALIFICACION MANUAL
'***********************************************************************
Public Sub dInsertEvalPigManual(ByVal psCodigo As String, ByVal pnCalif As String, ByVal pnCalifAnt As String, ByVal psUsuario As String, ByVal psFec As Date, ByVal pncalifica As String, ByVal pbEjecBatch As Boolean)

On Error GoTo ErrInsertTarifario

    lsSQL = "INSERT COLOCPIGEVALMANUAL(cPersCod, cEvalPigno, cEvalPignoAnterior, cUsuarioEvalPigno, fEvalPigno,iCalifica) VALUES ( " _
        & psCodigo & "," & pnCalif & "," & pnCalifAnt & ",'" & psUsuario & "','" & Format(psFec & " " & oConn.GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "','" & pncalifica & "')"

    If pbEjecBatch Then
        oConn.AdicionaCmdBatch lsSQL
    Else
        oConn.Ejecutar lsSQL
    End If
    
    Exit Sub
    
ErrInsertTarifario:
    If Err.Number = -2147211505 Then
       MsgBox "Dato ya existe", vbInformation, "Aviso"
    End If
End Sub

Public Sub dEvalPigno(ByVal dFecha As String)
On Error GoTo ErrorEvaluacionPignoraticia

lsSQL = "exec ColocPigEvalClientes '" & dFecha & "'"
oConn.Ejecutar lsSQL

    
Exit Sub

ErrorEvaluacionPignoraticia:
    Err.Raise Err.Number, "Error En Proceso EvaluacionPignoraticia", Err.Description


End Sub


