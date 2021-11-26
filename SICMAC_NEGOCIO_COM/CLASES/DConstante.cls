VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DConstante"
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

Public Sub inicio(psServerComunes As String, psServerPersona As String, psServerAdministracion As String, psServerNegocio As String, psServerImagenes As String)
    lsServerComunes = psServerComunes
    lsServerPersona = psServerPersona
    lsServerAdministracion = psServerAdministracion
    lsServerNegocio = psServerNegocio
    lsServerImagenes = psServerImagenes
End Sub



Public Function GetSector() As ADODB.Recordset
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Co As New DConecta
    Set Co = New DConecta
    sql = "select nConsValor, cConsDescripcion from Constante where nConsCod = 1013 and nConsValor<>1013"
    Co.AbreConexion
    Set rs = Co.CargaRecordSet(sql)
    Co.CierraConexion
    Set GetSector = rs
    Set rs = Nothing
End Function

'Public Function GetConstante(psGrupo As ConstanteCabecera, Optional pbOrdenCol1 As Boolean = True, Optional pbAscendente As Boolean = True) As Recordset
 Public Function GetConstante(ByVal psGrupo As ConstanteCabecera, Optional pbOrdenCol1 As Boolean = True, Optional pbAscendente As Boolean = True, Optional pbEstructuraArbol As Boolean = False, Optional psSubArbol As String = "", Optional psValorExcluido As String = "", Optional pbUnido As Boolean = False) As Recordset
    Dim sqlC As String
    Dim oCon As DConecta
    Dim lsOrdena As String
    Set oCon = New DConecta
    
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
    Call RaiseError(MyUnhandledError, "DConstantes:GetConstante Method")
End Function

Public Function GetPeriodosNoLabTpo(Optional psSubArbol As String = "", Optional pbOrdenCol1 As Boolean = True, Optional pbAscendente As Boolean = True, Optional pbSoloAprovacionRechazo As Boolean = False) As Recordset
    Dim sqlC As String
    Dim oCon As DConecta
    Dim lsOrdena As String
    Set oCon = New DConecta
    
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
    Call RaiseError(MyUnhandledError, "DConstantes:GetConstante Method")
End Function


'##ModelId=3AC2BBDD023F
Public Function GetAgencias(Optional pbOrdenCol1 As Boolean = False, Optional pbAscendente As Boolean = True, Optional pbTipoArbol As Boolean = False, Optional psArea As String = "", Optional psPersCod As String = "") As Recordset
    Dim sqlC As String
    Dim oCon As DConecta
    Dim lsOrdena As String
    Set oCon = New DConecta
    
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
    Call RaiseError(MyUnhandledError, "DConstantes:GetConstante Method")
End Function

'##ModelId=3AC2BBDD023F
Public Function GetAreas(Optional pbOrdenCol1 As Boolean = True, Optional pbAscendente As Boolean = True) As Recordset
    Dim sqlC As String
    Dim oCon As DConecta
    Dim lsOrdena As String
    Set oCon = New DConecta
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
    Call RaiseError(MyUnhandledError, "DConstantes:GetConstante Method")
End Function

'##ModelId=3AC2BBDD023F
Public Function GetRHTipoContrato() As Recordset
    Dim sqlC As String
    Dim oCon As DConecta
    Dim lsOrdena As String
    Set oCon = New DConecta
    
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
    Call RaiseError(MyUnhandledError, "DConstantes:GetConstante Method")
End Function

Public Function RecuperaConstantes(ByVal psConstanteCab As ConstanteCabecera, _
    Optional ByVal pnFiltro As Integer = -1, Optional ByVal psCodOrden As String = "C.nConsValor") As ADODB.Recordset
    Dim Conn As DConecta
    Dim sSql As String
    
    On Error GoTo ErrorRecuperaConstantes
    Set Conn = New DConecta
    Conn.AbreConexion
    sSql = "Select C.cConsDescripcion,C.nConsValor from " & sCentralCom & "Constante C "
    If pnFiltro <> -1 Then
        sSql = sSql & " Inner join ColocCredConsFiltro CF ON C.nConsValor = CF.nConsValor AND nCodFiltro = " & pnFiltro
    End If
    sSql = sSql & " where C.nConsValor <> '" & psConstanteCab & "' And  C.nConsCod = '" & psConstanteCab & "' order by " & psCodOrden
    Set RecuperaConstantes = Conn.CargaRecordSet(sSql)
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Function

ErrorRecuperaConstantes:
    Err.Raise Err.Number, "Error En Proceso RecuperaConstantes", Err.Description
    
End Function

Public Function DameDescripcionConstante(ByVal pnCodigo As ConstanteCabecera, ByVal pnValor As Long) As String
Dim sSql As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

    On Error GoTo ErrorDameDescripcionConstante
    sSql = "Select cConsDescripcion from Constante where nConsCod = " & pnCodigo & " AND nConsValor = " & pnValor
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
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

Private Sub Class_Initialize()
Dim ClsIni As COMConecta.DCOMClasIni
    Set ClsIni = New COMConecta.DCOMClasIni
    sCentralCom = ClsIni.BaseComunes
    Set ClsIni = Nothing
End Sub
