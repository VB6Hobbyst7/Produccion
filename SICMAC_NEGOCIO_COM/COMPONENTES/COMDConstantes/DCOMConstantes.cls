VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMConstantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Public bError As Boolean
Public sMsgError As String
Private sCentralCom As String

Private lsServerComunes As String
Private lsServerPersona As String
Private lsServerAdministracion As String
Private lsServerNegocio As String
Private lsServerImagenes As String

Dim oError As New COMConecta.COMErrorHandling

Public Sub inicio(psServerComunes As String, psServerPersona As String, psServerAdministracion As String, psServerNegocio As String, psServerImagenes As String)
    lsServerComunes = psServerComunes
    lsServerPersona = psServerPersona
    lsServerAdministracion = psServerAdministracion
    lsServerNegocio = psServerNegocio
    lsServerImagenes = psServerImagenes
End Sub



Public Function GetSector() As ADODB.Recordset
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim Co As New COMConecta.DCOMConecta
    Set Co = New COMConecta.DCOMConecta
    Sql = "select nConsValor, cConsDescripcion from Constante where nConsCod = 1013 and nConsValor<>1013"
    Co.AbreConexion
    Set rs = Co.CargaRecordSet(Sql)
    Co.CierraConexion
    Set GetSector = rs
    Set rs = Nothing
End Function

'Public Function GetConstante(psGrupo As ConstanteCabecera, Optional pbOrdenCol1 As Boolean = True, Optional pbAscendente As Boolean = True) As Recordset
 Public Function GetConstante(ByVal psGrupo As ConstanteCabecera, Optional pbOrdenCol1 As Boolean = True, Optional pbAscendente As Boolean = True, Optional pbEstructuraArbol As Boolean = False, Optional psSubArbol As String = "", Optional psValorExcluido As String = "", Optional pbUnido As Boolean = False) As Recordset
    Dim sqlC As String
    Dim oCon As COMConecta.DCOMConecta
    Dim lsOrdena As String
    Set oCon = New COMConecta.DCOMConecta
    
    On Error GoTo GetConstanteErr
    
    lsOrdena = " Order by " & IIf(pbOrdenCol1, " Columna1 ", " Columna2 ") & IIf(pbAscendente, " ASC ", " DESC ")
    If Not pbEstructuraArbol Then
        If psValorExcluido = "" Then
            If pbUnido Then
                sqlC = "Select cConsDescripcion + space(50) + Str(nConsValor) As Columna1  From " & lsServerComunes & "Constante where Left(nConsCod," & Len(Trim(psGrupo)) & ") = '" & psGrupo & "' " & lsOrdena
            Else
                sqlC = "Select cConsDescripcion as Columna1, nConsValor as Columna2 From " & lsServerComunes & "Constante where Left(nConsCod," & Len(Trim(psGrupo)) & ") = '" & psGrupo & "' " & lsOrdena
            End If
        Else
            If pbUnido Then
                sqlC = "Select cConsDescripcion + space(50) + Str(nConsValor) As Columna1 From " & lsServerComunes & "Constante where Left(nConsCod," & Len(Trim(psGrupo)) & ") = '" & psGrupo & "' And nConsValor not in ('" & psValorExcluido & "') " & lsOrdena
            Else
                sqlC = "Select cConsDescripcion as Columna1, nConsValor as Columna2 From " & lsServerComunes & "Constante where Left(nConsCod," & Len(Trim(psGrupo)) & ") = '" & psGrupo & "' And nConsValor not in ('" & psValorExcluido & "') " & lsOrdena
            End If
        End If
    Else
        If psSubArbol = "" Then
            If psValorExcluido = "" Then
                sqlC = "Select nConsValor as Columna2, cConsDescripcion as Columna1,  Len(nConsValor) Nivel From " & lsServerComunes & "Constante where Left(nConsCod," & Len(Trim(psGrupo)) & ") = '" & psGrupo & "' " & lsOrdena
            Else
                sqlC = "Select nConsValor as Columna2, cConsDescripcion as Columna1,  Len(nConsValor) Nivel From " & lsServerComunes & "Constante where Left(nConsCod," & Len(Trim(psGrupo)) & ") = '" & psGrupo & "' And nConsValor not in ('" & psValorExcluido & "')" & lsOrdena
            End If
        Else
            If psValorExcluido = "" Then
                sqlC = "Select nConsValor as Columna2, cConsDescripcion as Columna1,  Len(nConsValor) Nivel From " & lsServerComunes & "Constante where Left(nConsCod," & Len(Trim(psGrupo)) & ") = '" & psGrupo & "' And nConsValor like '" & psSubArbol & "%'" & lsOrdena
            Else
                sqlC = "Select nConsValor as Columna2, cConsDescripcion as Columna1,  Len(nConsValor) Nivel From " & lsServerComunes & "Constante where Left(nConsCod," & Len(Trim(psGrupo)) & ") = '" & psGrupo & "' And nConsValor like '" & psSubArbol & "%' And nConsValor not in ('" & psValorExcluido & "')" & lsOrdena
            End If
        End If
    End If
    
    If oCon.AbreConexion Then
        Set GetConstante = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    
    Exit Function
GetConstanteErr:
   oError.RaiseError oError.MyUnhandledError, "DConstantes:GetConstante Method"
End Function

Public Function GetPeriodosNoLabTpo(Optional psSubArbol As String = "", Optional pbOrdenCol1 As Boolean = True, Optional pbAscendente As Boolean = True, Optional pbSoloAprovacionRechazo As Boolean = False) As Recordset
    Dim sqlC As String
    Dim oCon As COMConecta.DCOMConecta
    Dim lsOrdena As String
    Set oCon = New COMConecta.DCOMConecta
    
    On Error GoTo GetConstanteErr
    
    lsOrdena = " Order by " & IIf(pbOrdenCol1, " Columna1 ", " Columna2 ") & IIf(pbAscendente, " ASC ", " DESC ")
    If pbSoloAprovacionRechazo Then
        sqlC = " Select nConsValor as Columna2, ltrim(rtrim(str(nRHPeriodoUnidad))) + '-' + cConsDescripcion as Columna1,  Len(nConsValor) Nivel From " & lsServerComunes & "Constante CO" _
             & " Inner Join RHPeriodosNolaboradosTpo RHT ON CO.nConsValor = RHT.nRHPeriodoTpo where nConsCod like '6006' And nConsValor like '" & psSubArbol & "%' And nConsValor > 0 " & lsOrdena
    Else
        sqlC = " Select nConsValor as Columna2, ltrim(rtrim(str(nRHPeriodoUnidad))) + '-' + cConsDescripcion as Columna1,  Len(nConsValor) Nivel From " & lsServerComunes & "Constante CO" _
             & " Inner Join RHPeriodosNolaboradosTpo RHT ON CO.nConsValor = RHT.nRHPeriodoTpo where nConsCod like '6006' And nConsValor like '" & psSubArbol & "%'" & lsOrdena
    End If
    
    If oCon.AbreConexion Then
        Set GetPeriodosNoLabTpo = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    
    Exit Function
GetConstanteErr:
    oError.RaiseError oError.MyUnhandledError, "DConstantes:GetConstante Method"
End Function


'##ModelId=3AC2BBDD023F
Public Function GetAgencias(Optional pbOrdenCol1 As Boolean = False, Optional pbAscendente As Boolean = True, Optional pbTipoArbol As Boolean = False, Optional psArea As String = "", Optional psPersCod As String = "") As Recordset
    Dim sqlC As String
    Dim oCon As COMConecta.DCOMConecta
    Dim lsOrdena As String
    Set oCon = New COMConecta.DCOMConecta
    
    On Error GoTo GetAgenciasErr
    
    lsOrdena = " Order by " & IIf(pbOrdenCol1, " Columna1 ", " Columna2 ") & IIf(pbAscendente, " ASC ", " DESC ")
    If Not pbTipoArbol Then
        sqlC = "Select cAgeDescripcion as Columna1, cAgeCod as Columna2 From " & lsServerComunes & "Agencias" & lsOrdena
    Else
        If psArea <> "" Then
            sqlC = " Select AA.cUbicaCod as Columna2, cAgeDescripcion as Columna1, Len(AG.cAgeCod) Nivel From " & lsServerComunes & "Agencias AG" _
                 & " Inner Join AreaAgencia AA On AG.cAgeCod = AA.cUbicaCod And AA.cAreaCod  = '" & psArea & "' " & lsOrdena
        ElseIf psPersCod <> "" Then
            sqlC = " Select Cargo.cAreaCod + AA.cUbicaCod as Columna2, cAgeDescripcion as Columna1, Len(AG.cAgeCod) Nivel From " & lsServerComunes & "Agencias AG" _
                 & " Inner Join AreaAgencia AA On AG.cAgeCod = AA.cUbicaCod" _
                 & " Inner Join (Select cAreaCod from RRHH where cPersCod = '" & psPersCod & "') Cargo On Cargo.cAreaCod = AA.cAreaCod  " _
                 & " " & lsOrdena
        Else
            sqlC = "Select cAgeCod as Columna2, cAgeDescripcion as Columna1, Len(cAgeCod) Nivel  From " & lsServerComunes & "Agencias" & lsOrdena
        End If
    End If
    If oCon.AbreConexion Then
        Set GetAgencias = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetAgenciasErr:
    oError.RaiseError oError.MyUnhandledError, "DConstantes:GetConstante Method"
End Function

'##ModelId=3AC2BBDD023F
Public Function GetAreas(Optional pbOrdenCol1 As Boolean = True, Optional pbAscendente As Boolean = True) As Recordset
    Dim sqlC As String
    Dim oCon As COMConecta.DCOMConecta
    Dim lsOrdena As String
    Set oCon = New COMConecta.DCOMConecta
    On Error GoTo GetAreasErr
    
    lsOrdena = " Order by " & IIf(pbOrdenCol1, " Columna1 ", " Columna2 ") & IIf(pbAscendente, " ASC ", " DESC ")
    sqlC = "Select cAreaDescripcion as Columna1, cAreaCod as Columna2 From " & lsServerComunes & "Areas " & lsOrdena
    
    If oCon.AbreConexion Then
        Set GetAreas = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetAreasErr:
    oError.RaiseError oError.MyUnhandledError, "DConstantes:GetConstante Method"
End Function

'##ModelId=3AC2BBDD023F
Public Function GetRHTipoContrato() As Recordset
    Dim sqlC As String
    Dim oCon As COMConecta.DCOMConecta
    Dim lsOrdena As String
    Set oCon = New COMConecta.DCOMConecta
    
    On Error GoTo GetAreasErr
    
    sqlC = " Select  Case nConsValor When 2 then 'L' " _
         & "    When 3 then 'F'" _
         & "    When 4 then 'P'" _
         & "    When 5 then 'S'" _
         & "    end Valor, cConsDescripcion Descrip , 2 Nivel from constante" _
         & " where nconscod like '6012%' and nConsvalor > 1" _
         & " Union" _
         & " Select 'E' Valor, 'LABORAL' Descrip , 2 Nivel Order By Valor"
    
    If oCon.AbreConexion Then
        Set GetRHTipoContrato = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetAreasErr:
    oError.RaiseError oError.MyUnhandledError, "DConstantes:GetConstante Method"
End Function

Public Function RecuperaConstantes(ByVal psConstanteCab As ConstanteCabecera, _
    Optional ByVal pnFiltro As Integer = -1, Optional ByVal psCodOrden As String = "C.nConsValor", _
    Optional ByVal pnEstado As Integer = 1, Optional ByVal pcTpProd As String = "") As ADODB.Recordset
    'Optional ByVal pcTpProd As String = "" - JOEP CP ERS034 - 2018
    
    Dim Conn As COMConecta.DCOMConecta
    Dim sSQL As String
    
    On Error GoTo ErrorRecuperaConstantes
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    
    sSQL = "Exec stp_sel_RecuperaConstante " & psConstanteCab & "," & pnFiltro & ",'" & psCodOrden & "'," & pnEstado & ",'" & pcTpProd & "'"  'agrego JOEP catalogo
'Comento JOEP Catalogo
'    sSQL = "Select C.cConsDescripcion,C.nConsValor from " & sCentralCom & "Constante C "
'    If pnFiltro <> -1 Then
'        sSQL = sSQL & " Inner join ColocCredConsFiltro CF ON C.nConsValor = CF.nConsValor AND nCodFiltro = " & pnFiltro
'    End If
'    'sSQL = sSQL & " where C.nConsValor <> '" & psConstanteCab & "' And  C.nConsCod = '" & psConstanteCab & "' order by " & psCodOrden
'    'modificado 30032006
'    sSQL = sSQL & " where C.nConsValor <> '" & psConstanteCab & "' And  C.nConsCod = '" & psConstanteCab & "' "
'
'    Select Case pnEstado
'        Case 0, 1
'            sSQL = sSQL & " And bEstado = " & pnEstado
'        Case 2
'            sSQL = sSQL & " And bEstado In (0, 1)"
'    End Select
'    sSQL = sSQL & " order by " & psCodOrden
'Comento JOEP Catalogo
    Set RecuperaConstantes = Conn.CargaRecordSet(sSQL)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function

ErrorRecuperaConstantes:
    Err.Raise Err.Number, "Error En Proceso RecuperaConstantes", Err.Description
    
End Function

Public Function RecuperaConstantes_2(ByVal psConstanteCab As Integer, _
    Optional ByVal pnFiltro As Integer = -1, Optional ByVal psCodOrden As String = "C.nConsValor") As ADODB.Recordset
    Dim Conn As COMConecta.DCOMConecta
    Dim sSQL As String
    
    On Error GoTo ErrorRecuperaConstantes
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    sSQL = "Select C.cConsDescripcion,C.nConsValor from " & sCentralCom & "Constante C "
    If pnFiltro <> -1 Then
        sSQL = sSQL & " Inner join ColocCredConsFiltro CF ON C.nConsValor = CF.nConsValor AND nCodFiltro = " & pnFiltro
    End If
    sSQL = sSQL & " where C.nConsValor <> '" & psConstanteCab & "' And  C.nConsCod = '" & psConstanteCab & "' order by " & psCodOrden
    Set RecuperaConstantes_2 = Conn.CargaRecordSet(sSQL)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function

ErrorRecuperaConstantes:
    Err.Raise Err.Number, "Error En Proceso RecuperaConstantes", Err.Description
    
End Function

Public Function DameDescripcionConstante(ByVal pnCodigo As ConstanteCabecera, ByVal pnValor As Long) As String
Dim sSQL As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    On Error GoTo ErrorDameDescripcionConstante
    sSQL = "Select cConsDescripcion from Constante where nConsCod = " & pnCodigo & " AND nConsValor = " & pnValor
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If Not R.BOF And Not R.EOF Then
        DameDescripcionConstante = Trim(R!cConsDescripcion)
    Else
        DameDescripcionConstante = ""
    End If
    R.Close
    Set R = Nothing
    Exit Function

ErrorDameDescripcionConstante:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function ObtenerVarRecuperaciones(ByVal pnValor As Integer) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    On Error GoTo ErrorDameDescripcionConstante
    sSQL = " select cConsDescripcion + space(20) + str(nConsValor) from Constante " & _
           " where nConsCod = " & pnValor & " AND nConsValor <> " & pnValor & ""
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If Not R.BOF And Not R.EOF Then
       Set ObtenerVarRecuperaciones = R
    Else
       Set ObtenerVarRecuperaciones = Nothing
    End If
  
    Exit Function

ErrorDameDescripcionConstante:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function ObtenerFeridos(ByVal pdFechaI As Date, ByVal pdFechaF As Date) As Integer
Dim sSQL As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    On Error GoTo ErrorDameDescripcionConstante
    sSQL = " select count(*) from Feriado where dferiado between '" & Format(pdFechaI, "YYYYMMDD") & "' and '" & Format(pdFechaF, "YYYYMMDD") & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If Not R.BOF And Not R.EOF Then
       ObtenerFeridos = R(0)
    Else
       ObtenerFeridos = 0
    End If
  
Exit Function

ErrorDameDescripcionConstante:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function ObtieneTipoPinPad(ByVal psNombrePc As String, ByVal lnPeriferico As Long) As ADODB.Recordset
    Dim lsql As String
    Dim rs As New ADODB.Recordset
    Dim oconec As COMConecta.DCOMConecta
    Set oconec = New COMConecta.DCOMConecta
    oconec.AbreConexion
    lsql = "SELECT nMarca FROM PERIFERICO where cPCNombre = '" & psNombrePc & "' and  nPeriferico = " & lnPeriferico
        Set rs = oconec.CargaRecordSet(lsql)
        Set ObtieneTipoPinPad = rs
    oconec.CierraConexion
    
End Function

Private Sub Class_Initialize()
    Dim ClsIni As New COMConecta.DCOMClasIni
    sCentralCom = ClsIni.BaseComunes
    Set ClsIni = Nothing
End Sub
'JACA 20110426********************************************************************
Public Function RecuperaMagnitudPersona(ByVal psConstanteCab As ConstanteCabecera, ByVal pnPersoneria As Integer) As ADODB.Recordset

    Dim Conn As COMConecta.DCOMConecta
    Dim sSQL As String
    
    On Error GoTo ErrorRecuperaMagnitudPersona
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    sSQL = "Select C.cConsDescripcion,C.nConsValor from " & sCentralCom & "Constante C "
    
    sSQL = sSQL & " where C.nConsValor <> '" & psConstanteCab & "' And  C.nConsCod = '" & psConstanteCab & "' "
    
    If pnPersoneria = 1 Then
        sSQL = sSQL & " and C.nConsValor>=4  "
    Else
        sSQL = sSQL & " and C.nConsValor<4  "
    End If
    
    sSQL = sSQL & " And bEstado = 1"
    sSQL = sSQL & " order by nConsValor"
    Set RecuperaMagnitudPersona = Conn.CargaRecordSet(sSQL)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function

ErrorRecuperaMagnitudPersona:
    Err.Raise Err.Number, "Error En Proceso RecuperaMagnitudPersona", Err.Description
    
End Function
'JACA END****************************************************************************

'*** BRGO 20110511 ***************************
Public Function ObtieneTipoBienGarantia() As ADODB.Recordset
    Dim lsql As String
    Dim rs As New ADODB.Recordset
    Dim oconec As COMConecta.DCOMConecta
    Set oconec = New COMConecta.DCOMConecta
    oconec.AbreConexion
    lsql = "Select ntipobienCod Codigo, Rtrim(cLiteral) + ' - ' + RTRIM(cDescripcion) Descripcion," _
         & " LEN(ntipobienCod) Nivel From colocrecuptipobienadjudicacion "
        Set rs = oconec.CargaRecordSet(lsql)
        Set ObtieneTipoBienGarantia = rs
    oconec.CierraConexion
End Function
'***********************************************************************

'*** JIPR20190520 ***************************
Public Function ObtieneTipoBienEmbargo() As ADODB.Recordset
    Dim sSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    sSQL = "exec stp_ObtieneTipoBienEmbargo"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneTipoBienEmbargo = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
End Function
'***********************************************************************

'MAVM 20110428 ***
Public Function ObtenerCargosArea(Optional psRHCargoCod As String = "") As ADODB.Recordset 'JUEZ 20121128 Se agregó psRHCargoCod
    Dim sSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    sSQL = "exec stp_sel_ObtenerCargosArea '" & psRHCargoCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerCargosArea = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
End Function

Public Function ObtenerNivelAprob() As ADODB.Recordset
    Dim sSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    sSQL = "exec stp_sel_ObtenerNivelAprob"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerNivelAprob = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
End Function

Public Function ObtenerRiesgoAprob() As ADODB.Recordset
    Dim sSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    sSQL = "exec stp_sel_ObtenerRiesgoAprob"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerRiesgoAprob = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
End Function
'***
'MIOL 20120919, SEGUN RQ12257 ******************************************
Public Function ObtenerConstReqSUNAT() As ADODB.Recordset
    Dim sSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    sSQL = "exec stp_sel_ConstReqSUNAT"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerConstReqSUNAT = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
End Function
'END MIOL **************************************************************

'JUEZ 20121129 *********************************************************
Public Function ObtenerUsuariosArea(Optional psPersCodUser As String = "") As ADODB.Recordset
    Dim sSQL As String
    Dim oConecta As COMConecta.DCOMConecta
    sSQL = "exec stp_sel_ObtenerUsuariosArea '" & psPersCodUser & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerUsuariosArea = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
End Function
'END JUEZ **************************************************************
'***Agregado por ELRO el 20130722, según TI-ERS079-2013****
Public Function devolverMediosRetiros() As ADODB.Recordset
    Dim lsSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    lsSql = "exec stp_sel_ERS0792013_DevolverMediosRetiros "
    
    If oConecta.AbreConexion = False Then Exit Function
    Set devolverMediosRetiros = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion
    
     Set oConecta = Nothing
End Function
'***Fin Agregado por ELRO el 20130722, según TI-ERS079-2013
'FRHU 20140905 ANEXO 01 TI-ERS106-2014
Public Function ObtenerParaOficialCumplimientoCIIU() As ADODB.Recordset
    Dim lsSql As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSql = "exec stp_sel_OfCumplimientoCIIU"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerParaOficialCumplimientoCIIU = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function ObtenerNivelRiesgoVariable(ByVal pTipo As Integer) As ADODB.Recordset
    Dim lsSql As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSql = "exec stp_sel_NivelRiesgoVariables " & pTipo
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerNivelRiesgoVariable = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function GetOficCumpTotalPersonasPorOcupacion() As ADODB.Recordset
    Dim lsSql As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSql = "exec stp_sel_ofCumpTotalPersonaPorOcupacion"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set GetOficCumpTotalPersonasPorOcupacion = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function GetOficCumpTotalPersonasPorCIIU() As ADODB.Recordset
    Dim lsSql As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSql = "exec stp_sel_OfCumpTotalPersonasPorCIIU"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set GetOficCumpTotalPersonasPorCIIU = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'FIN FRHU 20140905
'FRHU 20141202 ERS048-2014 *******************************************************
Public Function GetConceptoNotaCargoAbono(ByVal psOpeCod As String) As ADODB.Recordset
    Dim lsSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    lsSql = "exec stp_sel_GetConceptosNotaAbonoCargo '" & psOpeCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set GetConceptoNotaCargoAbono = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function GetProductoConcepto(ByVal pnPrdConcepto As Integer) As String
    Dim lsSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rsConcepto As ADODB.Recordset
    
    lsSql = "exec stp_sel_GetProductoConcepto " & pnPrdConcepto
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rsConcepto = oConecta.CargaRecordSet(lsSql)
    If Not rsConcepto.EOF And Not rsConcepto.BOF Then
        GetProductoConcepto = rsConcepto("cDescripcion")
    Else
        GetProductoConcepto = ""
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'FIN FRHU 20141202 ***************************************************
'ALPA 20150205****************************************
Public Function GetCargo() As ADODB.Recordset
    Dim lsSql As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSql = "exec stp_sel_cargos"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set GetCargo = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'*****************************************************
'RECO20150318 ERS010-2015*****************************
Public Function ObtieneConstanteFiltroXCodValor(ByVal pnConsCod As Integer, ByVal psConsValor As String) As ADODB.Recordset
    Dim lsSql As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSql = "exec stp_sel_ObtieneConstanteFiltroXCodValor " & pnConsCod & ",'" & psConsValor & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtieneConstanteFiltroXCodValor = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'RECO FIN*********************************************

'*******VARIABLES DE MONEDA****************
'MARG ERS044-2016

Public Function gcPEN_SINGULAR() As String
    Dim R As ADODB.Recordset
    Dim lsSql As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSql = "exec stp_sel_gcPEN_SINGULAR"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    gcPEN_SINGULAR = R!cConsDescripcion
    
End Function
Public Function gcPEN_PLURAL() As String
    Dim R As New ADODB.Recordset
    Dim lsSql As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSql = "exec stp_sel_gcPEN_PLURAL"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    gcPEN_PLURAL = R!cConsDescripcion
End Function
Public Function gcPEN_SIMBOLO() As String
    Dim R As ADODB.Recordset
    Dim lsSql As String
    Dim oConecta As COMConecta.DCOMConecta
    lsSql = "exec stp_sel_gcPEN_SIMBOLO"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    gcPEN_SIMBOLO = R!cConsDescripcion
End Function
'******************************************

'ANDE 2017-07-12
Public Function ObtenerConstante(ByVal pnValor As Integer) As ADODB.Recordset
Dim sSQL As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    On Error GoTo ErrorDameDescripcionConstante
    sSQL = " select cConsDescripcion + space(20) + str(nConsValor) as cConstante from Constante " & _
           " where nConsCod = " & pnValor & " AND nConsValor <> " & pnValor & ""
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If Not R.BOF And Not R.EOF Then
       Set ObtenerConstante = R
    Else
       Set ObtenerConstante = Nothing
    End If
  
    Exit Function

ErrorDameDescripcionConstante:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function
'END ANDE

'CTI1 20180730 ***
Public Function RecuperaConstantesGeneral(ByVal psConstanteCab As ConstanteCabecera, _
    Optional ByVal nConsValor As Integer = -1) As ADODB.Recordset

    Dim Conn As COMConecta.DCOMConecta
    Dim sSQL As String
    
    On Error GoTo ErrorRecuperaConstantes
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    sSQL = "EXEC stp_sel_ObtenerConstanteGen " & psConstanteCab & "," & nConsValor

    Set RecuperaConstantesGeneral = Conn.CargaRecordSet(sSQL)
    Conn.CierraConexion
    Set Conn = Nothing
    
    Exit Function

ErrorRecuperaConstantes:
    Err.Raise Err.Number, "Error En Proceso RecuperaConstantes", Err.Description
End Function
Public Function VerificaExisteConstante(ByVal psConstanteCab As ConstanteCabecera, _
    Optional ByVal nConsValor As Integer = -1) As Boolean
    Dim Conn As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo ErrorVerificaExisteConstante
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    sSQL = "EXEC stp_sel_ObtenerConstanteGen " & psConstanteCab & "," & nConsValor

    Set rs = Conn.CargaRecordSet(sSQL)
    VerificaExisteConstante = False
    If Not (rs.BOF And rs.EOF) Then
        VerificaExisteConstante = True
    End If
    Conn.CierraConexion
    Set Conn = Nothing
    
    Exit Function

ErrorVerificaExisteConstante:
    VerificaExisteConstante = False
    Err.Raise Err.Number, "Error En Proceso RecuperaConstantes", Err.Description
End Function
'CTI1 FIN ********
'***CTI3 (ferimoro)   18102018
Public Function ObtenerConstanteExtornoMotivo() As ADODB.Recordset
Dim sSQL As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    On Error GoTo ErrorDameDescripcionConstante
    sSQL = "select * from movMotivosExtorno"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    If Not R.BOF And Not R.EOF Then
       Set ObtenerConstanteExtornoMotivo = R
    Else
       Set ObtenerConstanteExtornoMotivo = Nothing
    End If
  
    Exit Function

ErrorDameDescripcionConstante:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

'APRI20190601 RFC1902040001
Public Function ObtenerMovCancelacion() As ADODB.Recordset
    Dim lsSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    lsSql = "exec USP_SEL_Constante " & 10713
    
    If oConecta.AbreConexion = False Then Exit Function
    Set ObtenerMovCancelacion = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion
     Set oConecta = Nothing
End Function
'END APRI
'ADD JHCU EncuestaPinPads
Public Function EncuestaPinPads(ByRef sUser As String, ByRef sCodAge As String, ByRef sCodEncuesta As String, ByRef sCodOpe As String) As ADODB.Recordset
    Dim lsSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrEncuestaPinPads
   
    Set oConecta = New COMConecta.DCOMConecta
    lsSql = "exec stp_sel_validaencuesta '" & sUser & "','" & sCodAge & "','" & sCodEncuesta & "','" & sCodOpe & "'"
    
    If oConecta.AbreConexion = False Then Exit Function
    Set EncuestaPinPads = oConecta.CargaRecordSet(lsSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrEncuestaPinPads:
     Err.Raise Err.Number, "Error En Proceso de cálculo de Encuesta PinPads", Err.Description
     Set EncuestaPinPads = Nothing
     Set oConecta = Nothing
End Function
Public Sub EncuestaPinPadsRes(ByRef sUser As String, ByRef sCodAge As String, ByRef sCodEncuesta As String, ByRef sCodOpe As String, ByRef cRes As Integer, Optional psUserVisto As String = "")
  
    Dim lsSql As String
    Dim oConecta As COMConecta.DCOMConecta
    
     On Error GoTo InsertaPinPadsRes
     Set oConecta = New COMConecta.DCOMConecta
     lsSql = "exec stp_ins_encuestarespuesta '" & sUser & "','" & sCodAge & "','" & sCodEncuesta & "','" & sCodOpe & "'," & cRes & ",'" & psUserVisto & "'"
    
     If oConecta.AbreConexion = False Then Exit Sub
     oConecta.Ejecutar (lsSql)
     oConecta.CierraConexion
     Set oConecta = Nothing
     Exit Sub
InsertaPinPadsRes:
     Err.Raise Err.Number, "Error En Proceso de Registro de Respuesta de Encuesta PinPads", Err.Description
     Set oConecta = Nothing
End Sub
'END JHCU

'JOEP20200928 Reprogramacion OCM
'Public Function RecuperaConstanteReprogaramacion(ByVal nConsValor As Integer, ByVal nFiltro As Integer) As ADODB.Recordset
Public Function RecuperaConstanteReprogaramacion(ByVal nConsValor As Integer, ByVal nFiltro As Integer, Optional ByVal nOpMenu As Integer = -1) As ADODB.Recordset 'Add JOEP20210306 garantia covid
    Dim Conn As COMConecta.DCOMConecta
    Dim sSQL As String
    
    On Error GoTo ErrorRecuperaConstanteReprogaramacion
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    'sSQL = "EXEC stp_Sel_ReprogramacionRecuperaConstante " & nConsValor & "," & nFiltro
    sSQL = "EXEC stp_Sel_ReprogramacionRecuperaConstante " & nConsValor & "," & nFiltro & "," & nOpMenu & "" 'Add JOEP20210306 garantia covid
    
    Set RecuperaConstanteReprogaramacion = Conn.CargaRecordSet(sSQL)
    Conn.CierraConexion
    Set Conn = Nothing
    
    Exit Function

ErrorRecuperaConstanteReprogaramacion:
    Err.Raise Err.Number, "Error En Proceso RecuperaConstantes", Err.Description
End Function
'JOEP20200928 Reprogramacion OCM

'CTI2 FERIMORO : ERS034-2019 : 04112020
Public Function RecuperaDoi(ByVal psConstanteCab As ConstanteCabecera) As ADODB.Recordset
        
    Dim Conn As COMConecta.DCOMConecta
    Dim sSQL As String
    
    On Error GoTo ErrorRecuperaConstantes
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    
    sSQL = "Exec stp_sel_ERS034_2019_recuperaDoi " & psConstanteCab & ""

    Set RecuperaDoi = Conn.CargaRecordSet(sSQL)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function

ErrorRecuperaConstantes:
    Err.Raise Err.Number, "Error En Proceso RecuperaDoi", Err.Description
    
End Function
