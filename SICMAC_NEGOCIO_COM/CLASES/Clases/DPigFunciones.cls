VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DPigFunciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim sql As String
Dim oConn As DConecta

Private Sub Class_Initialize()
Dim oIni As ClasIni
Dim sConn As String

    Set oIni = New ClasIni
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

Public Function dObtieneRecordSet(ByVal psSql As String) As ADODB.Recordset
Dim lrs As ADODB.Recordset

On Error GoTo dError

'Set lrs = New ADODB.Recordset
Set dObtieneRecordSet = oConn.CargaRecordSet(psSql)

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Recordset <<dObtieneRecordSet>>", Err.Description

End Function

'***************************************************************
' CAFF - Devuelve la Descripcion + el Valor de una Constante
'***************************************************************
Public Function GetConstante(ByVal psCodCons As String) As Recordset


On Error GoTo ErrorConstante
sql = "SELECT cConsDescripcion + space(25) +  convert(varchar(4), nConsValor) AS Constante, cConsDescripcion, nConsValor FROM CONSTANTE WHERE nConsCod LIKE '" & psCodCons & "' AND nConsValor <> '" & psCodCons & "'"

    Set GetConstante = oConn.CargaRecordSet(sql)
    
    Exit Function

ErrorConstante:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

'***************************************************************
' CAFF - Devuelve el Valor de un Parametro de ColocParametro
'***************************************************************

Public Function GetParamValor(ByVal psCodParam As ColocPParametros) As Double
Dim Rs As Recordset

On Error GoTo ErrorConstante
sql = "SELECT nParamValor FROM ColocParametro WHERE nParamVar = " & psCodParam

    Set Rs = oConn.CargaRecordSet(sql)

    GetParamValor = Rs!nParamValor
    
    Set Rs = Nothing
    
    Exit Function

ErrorConstante:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function
'****************************************************************
'CAFF  - 20/08/2002- Obtiene la Calificacion interna del Cliente
'****************************************************************
Public Function GetEvalCliente(ByVal psCodPers) As Recordset

On Error GoTo ErrorEvalCliente

sql = "SELECT cCalifiCliente, cConsDescripcion FROM ColocPigEvalCliente E INNER JOIN Constante C ON " & _
    " E.cCalifiCliente = C.nConsValor WHERE nConsCod = " & gColocPigCalifCte & " AND cPersCod = '" & psCodPers & "'"
    
    Set GetEvalCliente = oConn.CargaRecordSet(sql)
    
    Exit Function

ErrorEvalCliente:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function
'***************************************************************
'CAFF  - 20/08/2002 - Obtiene la Calificacion SBS del Cliente
'***************************************************************
Public Function GetCalifCliente(ByVal psCodPers) As Recordset

On Error GoTo ErrorCalifCliente

sql = "SELECT cEvalSBSCliente, cConsDescripcion FROM ColocPigEvalCliente E INNER JOIN Constante C ON " & _
    " E.cCalifiCliente = C.nConsValor WHERE nConsCod = " & gColocPigCalifSbs & " AND cPersCod = '" & psCodPers & "'"
    
    Set GetCalifCliente = oConn.CargaRecordSet(sql)
    
    Exit Function

ErrorCalifCliente:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'***************************************************************************
' CAFF - Obtiene el Prescio del Material dependiendo del Tipo de Proceso
'***************************************************************************

Public Function GetPrecioMaterial(ByVal pnTipoVal, ByVal pnTipoMat As Integer, pnTipoMon As Integer) As Currency
Dim rsTemp As Recordset

On Error GoTo ErrorPrecioMaterial

sql = "SELECT TOP 1 nValor FROM ColocPigPrecioMaterial WHERE nTipoValor = " & pnTipoVal & " AND nMaterial = " & pnTipoMat & _
    " AND nMoneda = " & pnTipoMon & " AND DATEDIFF(d, dInicial, '" & Format(Date, "mm/dd/yyyy") & "') > 0 ORDER BY dInicial Desc"
    Set rsTemp = oConn.CargaRecordSet(sql)
    
    If rsTemp.EOF And rsTemp.BOF Then
        GetPrecioMaterial = 0
    Else
        GetPrecioMaterial = rsTemp!nValor
    End If
   
    Exit Function
    
ErrorPrecioMaterial:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'*********************************************************************
' CAFF - Devuelve el Valor de un Concepto de la Tabla ColocConcepto
'*********************************************************************
Public Function GetConceptoValor(ByVal pnCodConcep As ColocConcepto) As Recordset

On Error GoTo ErrorConcepto

sql = "SELECT * FROM ProductoConcepto WHERE nPrdConceptoCod = " & pnCodConcep

    Set GetConceptoValor = oConn.CargaRecordSet(sql)
    
    Exit Function

ErrorConcepto:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'*****************************************************************************
' CAFF - Devuelve LineaCredito, Tasa Compensatoria y Moratoria de la Linea
'*****************************************************************************
Public Function GetLineaCredito(ByVal pnCapital As Currency) As Recordset

On Error GoTo ErrorLineaCredito

sql = "SELECT cLineaCred, " & _
    "TasaComp = (SELECT nTasaIni FROM colocLineaCreditoTasa WHERE cLineaCred = LC.cLineaCred and nColocLinCredTasaTpo = 1), " & _
    "TasaMora = (SELECT nTasaIni FROM colocLineaCreditoTasa WHERE cLineaCred = LC.cLineaCred and nColocLinCredTasaTpo = 3) " & _
    "FROM ColocLineaCredito LC " & _
    "WHERE Substring(cLineaCred,7,3) = '305' AND (" & pnCapital & " Between nMontoMin AND nMontoMax) AND bEstado = 1"

    Set GetLineaCredito = oConn.CargaRecordSet(sql)
    
    Exit Function
    
ErrorLineaCredito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function dObtieneMovNroAnterior(ByVal psCtaCod As String, ByVal psOpeCod As String) As Long
Dim Rs As ADODB.Recordset
Dim lnMovNro As String

On Error GoTo dError

Set Rs = New ADODB.Recordset

    sql = "SELECT nMovNro FROM MovCol WHERE cCtaCod = '" & psCtaCod & "' AND cOpecod ='" & psOpeCod & "' "
    Set Rs = oConn.CargaRecordSet(sql)
    
    If Rs.BOF And Rs.EOF Then
        lnMovNro = 0
    Else
        Do While Not Rs.EOF
            lnMovNro = Rs!nMovNro
            Rs.MoveNext
        Loop
    End If
    Rs.Close

Set Rs = Nothing
dObtieneMovNroAnterior = lnMovNro

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Nro de Movimiento anterior <<dObtieneMovNroAnterior>>", Err.Description

End Function

'***********************************************************************
' CAFF    - 03/09/2002 - Devuelve el Numero Telefonico de una Agencia
'***********************************************************************

Public Function GetFonoAge(ByVal psCodAge As String) As String
Dim Rs As ADODB.Recordset
Dim lsFonoAge As String

On Error GoTo dError

    sql = "SELECT cAgeTelefono FROM Agencias WHERE cAgeCod = '" & psCodAge & "'"
    
    Set Rs = oConn.CargaRecordSet(sql)
    
    If Rs.BOF And Rs.EOF Then
        lsFonoAge = ""
    Else
        lsFonoAge = Rs!cAgeTelefono
    End If
    
Set Rs = Nothing

GetFonoAge = lsFonoAge

Exit Function

dError:
    Err.Raise Err.Number, "Error en GetFonoAge <<GetFonoAge>>", Err.Description

End Function

'***************************************************
'CAFF - 03/09/2002 - Devuelve datos del tarifario
'***************************************************
Public Function GetTarifario() As Recordset

On Error GoTo ErrorTarifario

sql = "SELECT " _
    & "Proceso = (SELECT cConsDescripcion FROM Constante C WHERE nConsCod = " & gColocPigTipoProcesoTar & " AND C.nConsValor = T.nTipoValor), " _
    & "Material = (SELECT cConsDescripcion FROM Constante C WHERE nConsCod = " & gColocPigMaterial & " AND C.nConsValor = T.nMaterial), " _
    & "Moneda = CASE nMoneda WHEN 1 THEN 'SOLES' ELSE 'DOLARES' END, convert(varchar(10), dInicial, 103) Fecha, nValor " _
    & "FROM colocPigPrecioMaterial T"

    Set GetTarifario = oConn.CargaRecordSet(sql)
    
    Exit Function
    
ErrorTarifario:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function GetNumGuia(ByVal psCodAge As String) As String
Dim Rs As Recordset
Dim lsCorre As String
Dim lsNumGuia As String
Dim sCorrDig As String

    sql = "SELECT ISNULL(CONVERT(INT,MAX(SUBSTRING(cNumDoc,4,9))),0) + 1 Correlativo FROM ColocPigGuia " _
            & "WHERE CONVERT(INT,SUBSTRING(cNumDoc,1,3)) = " & CInt(psCodAge)

    Set Rs = oConn.CargaRecordSet(sql)
   
    sCorrDig = FillNum(Trim(Rs!Correlativo), 9, "0")
    lsNumGuia = FillNum(Trim(psCodAge), 3, "0") & sCorrDig

    Set Rs = Nothing
    
    GetNumGuia = lsNumGuia
    
End Function

'************************************************************************
' CMPCL   - Devuelve el Maximo Numero de Remate de un Contrato
' EAFA     - 02/10/2002
'************************************************************************
Public Function GetRemate(ByVal psCtaCod As String) As Integer
Dim Rs As ADODB.Recordset

On Error GoTo ErrorRemate

sql = " SELECT max(nRemate) nRemate " _
       & " FROM ColocPigProceso " _
       & " WHERE cCtaCod = " & psCtaCod

    Set Rs = oConn.CargaRecordSet(sql)
    
    GetRemate = Rs!nRemate
    
    Exit Function
    
ErrorRemate:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'************************************************************************
' CMPCL  - Devuelve la Serie de un Documento (Poliza, Boleta, Factura)
' EAFA     - 09/10/2002
'************************************************************************
Public Function GetSerDocumento(ByVal pnUbicacion As Integer, ByVal pnTipDocu As Integer) As String
Dim Rs As ADODB.Recordset

On Error GoTo ErrorSerDocumento

sql = " SELECT cSerDocumento " _
       & " FROM ColocPigDocumento " _
       & " WHERE nUbicacion = " & pnUbicacion _
       & "     AND nCodTipo = " & pnTipDocu

    Set Rs = oConn.CargaRecordSet(sql)
    
    If Not Rs.EOF And Not Rs.BOF Then
        GetSerDocumento = Right("000" & Trim(IIf(IsNull(Rs!cSerDocumento), "0", Rs!cSerDocumento)), 3)
    Else
        GetSerDocumento = Right("001", 3)
    End If
    
    Exit Function
    
ErrorSerDocumento:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


'**************************************************************************************
' CMPCL  - Devuelve el Maximo numero de Documento (Poliza, Boleta, Factura) de una Serie y Ubicación
' EAFA     - 09/10/2002
' CRSF    - 16/06/2003 , modificacion de cDocumento,1,4 a 1,3
'**************************************************************************************
Public Function GetNumDocumento(ByVal pnTipDocu As Integer, ByVal pnSerDocumento As Integer) As String
Dim Rs As ADODB.Recordset

On Error GoTo ErrorNumDocumento

sql = " SELECT isNull(Max(Convert(Int,Substring(cDocumento,5,8))),0) + 1 cNumDocumento " _
       & " FROM MovDocPig " _
       & " WHERE nCodTipo = " & pnTipDocu _
       & "    AND Convert(Int,Substring(cDocumento,1,3)) = " & pnSerDocumento

    Set Rs = oConn.CargaRecordSet(sql)
    
    GetNumDocumento = Right("00000000" & Trim(Rs!cNumDocumento), 8)
    
    Exit Function
    
ErrorNumDocumento:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


'***************************************************
' CMCPL - OBTIENE LA LISTA DE REMATES ADJUDICADOS
' EAFA - 24/10/02
'***************************************************
Public Function GetListaRemates() As Recordset
Dim lrs As ADODB.Recordset

On Error GoTo dError

sql = " SELECT  nRemate " _
        & "   FROM ColocPigRemate " _
        & "WHERE nAdjudicado = 1 Order by nRemate"
        
Set lrs = oConn.CargaRecordSet(sql)

Set GetListaRemates = lrs
Set lrs = Nothing

Exit Function
dError:
    Err.Raise Err.Number, "Obtiene Lista de Remate en <<GetListaRemates>>", Err.Description
End Function

'***************************************************
' CMCPL - OBTIENE LA LISTA DE AGENCIAS
' EAFA - 24/10/02
'***************************************************
Public Function GetListaAgencias() As Recordset
Dim lrs As ADODB.Recordset

On Error GoTo dError

sql = " SELECT  cAgeCod, cAgeDescripcion " _
        & "   FROM Agencias " _
        & "WHERE nAgeEspecial = 0 Order by cAgeCod"
        
Set lrs = oConn.CargaRecordSet(sql)

Set GetListaAgencias = lrs
Set lrs = Nothing

Exit Function
dError:
    Err.Raise Err.Number, "Obtiene Lista de Agencias en <<GetListaAgencias>>", Err.Description
End Function


'************************************************************************
' CMPCL   - Devuelve la Ubicacion de la Agencia
' JAPP    - 04/07/2003
'************************************************************************
Public Function GetUbicacion(ByVal psCodAge As String) As String
Dim Rs As ADODB.Recordset
On Error GoTo ErrorUbicacion

sql = " SELECT cUbiCod " _
       & " FROM ColocPigUbicacion " _
       & " WHERE cAgeCod = '" & psCodAge & "'"

    Set Rs = oConn.CargaRecordSet(sql)
    GetUbicacion = Rs!cUbiCod
    
    Exit Function
    
ErrorUbicacion:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function
'******************************************************************
' CMPCL   - Devuelve el Numero de Piezas de un Contrato
' JAPP    - 08/07/2003
'******************************************************************
Public Function GetNroPieza(ByVal pnCodCta As String) As Integer
Dim Rs As ADODB.Recordset

On Error GoTo ErrorItemPieza

sql = "SELECT t1.nPiezas npiezas " _
      & "From ColocPigno t1 " _
      & "Where t1.cCtaCod = '" & pnCodCta & "'"

    Set Rs = oConn.CargaRecordSet(sql)
    
    GetNroPieza = Rs!npiezas
    
    Exit Function
    
ErrorItemPieza:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


