VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function RecuperaGarantiasxDatos(ByVal psTipoGar As String, ByVal psTipoDoc As String, ByVal psNumDoc As String) As ADODB.Recordset
Dim oConecta As DConecta
Dim sSql As String

    On Error GoTo ErrorRecuperaGarantiasxDatos
    sSql = " Select DISTINCT G.nEstado, C3.cConsDescripcion as cEstadoGar, G.cNumGarant, P.cPersNombre, G.cTpoDoc, G.cNroDoc, G.nTpoGarantia, C.cConsDescripcion as cTpoGarDescripcion, "
    sSql = sSql & " G.cDescripcion, G.nMoneda, C2.cConsdescripcion as cMonedaDesc,  G.nTasacion, G.nRealizacion, G.nGravament, G.nPorGravar, G.cPersCodEmisor, "
    sSql = sSql & " G.nPorGravar - G.nGravament as nDisponible, G.cPersCodEmisor, D.cDocDesc,G.nGarClase, G.nGarTpoRealiz, "
    sSql = sSql & " G.cBancoPersCod, PB.cPersNombre as cBanco,G.nGarantReal"
    sSql = sSql & " From Garantias G Inner Join PersGarantia PG ON G.cNumGarant = PG.cNumGarant "
    sSql = sSql & "                  Inner Join Constante C ON G.nTpoGarantia = C.nConsValor AND C.nConsCod = " & gPersGarantia
    sSql = sSql & "                  Inner Join Constante C2 ON G.nMoneda = C2.nConsValor AND C2.nConsCod = " & gMoneda
    sSql = sSql & "                  Inner Join Persona P ON P.cPersCod = G.cPersCodEmisor "
    sSql = sSql & "                  Left Join Persona PB ON PB.cPersCod = G.cBancoPersCod "
    sSql = sSql & "                  Inner Join GarantDoc GD ON G.nTpoGarantia = GD.nConsValor "
    sSql = sSql & "                  Inner Join Documento D ON D.nDocTpo = GD.nDocTpo "
    sSql = sSql & "                  Inner Join Constante C3 ON G.nEstado = C3.nConsValor AND C3.nConsCod = 1030 "
    sSql = sSql & " WHERE G.nTpoGarantia = " & psTipoGar & " AND G.cTpoDoc = '" & psTipoDoc & "' AND G.cNroDoc = '" & psNumDoc & "' "
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGarantiasxDatos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaGarantiasxDatos:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


Public Sub ExtornaGarantiaReal(ByVal Nrog As String, ByVal NroMov As String)
Dim sSql As String
Dim sSql2 As String
Dim oConecta As DConecta

    
    
           sSql = " Update Mov "
    sSql = sSql & " Set nMovFlag=" & gMovFlagExtornado
    sSql = sSql & " Where cMovNro =" & "'" & NroMov & "'"
    'sSql = sSql & "                "
    sSql2 = sSql2 & " Update Garantias"
    sSql2 = sSql2 & " Set nEstado = 4"
    sSql2 = sSql2 & " Where cNumGarant =" & Nrog

    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.BeginTrans
    oConecta.Ejecutar (sSql)
    oConecta.Ejecutar (sSql2)
    oConecta.CommitTrans
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

End Sub
Public Function RecuperaExtornoGarantReal(ByVal psFecha As String, ByVal opt As Integer, Optional psUser As String)
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecupera
    
    sSql = "Select MG.cNumGarant as Nro_Garantia, MG.cOpeCod as Cod_Operacion, M.cMovNro as Transaccion From MovGarantia MG"
    sSql = sSql & " Inner Join Mov M on M.nMovNro=MG.nMovNro"
    sSql = sSql & " Where M.nmovFlag=0 and MG.cOpeCod=" & gGarantLevanta
    If opt = 0 Then
        
        sSql = sSql & " and M.cMovNro like '" & psFecha & "%'"
    Else
        sSql = sSql & " and M.cMovNro like '" & psFecha & "%" & psUser & "'"
    End If
    sSql = sSql & " ORDER BY M.cMovNro"

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaExtornoGarantReal = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecupera:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function
Public Function RecuperaTitularGarantReal(ByVal Nrog As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecupera
    
 
sSql = "Select PG.cPersCod as CodPers, P.cPersNombre as Nombre, GR.nGarantReal From Garantias GR"
sSql = sSql & " Inner Join PersGarantia PG on PG.cNumGarant=GR.cNumGarant"
sSql = sSql & " Inner Join Persona P on P.cPersCod=Pg.cPersCod"
sSql = sSql & " Where GR.cNumGarant ='" & Nrog & "'"
sSql = sSql & " And PG.nRelacion = 1"

    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaTitularGarantReal = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecupera:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaTotalLiberacion(ByVal pdValorRealiza As Double, ByVal nMoneda As String, _
                ByVal psCodAge As String)
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecupera
    
    sSql = "Select PC.nPrdConceptoCod as ConceptoCod, PC.cDescripcion as De, "
    sSql = sSql & " Monto = case"
    sSql = sSql & "         when PC.nTpoValor=1 then PC.nValor"
    sSql = sSql & "         when PC.nTpoValor=2 then"
    sSql = sSql & "          Case"
    sSql = sSql & "           when " & pdValorRealiza & " *nvalor/1000 < Pc.nMontoMin then Pc.nMontoMin"
    sSql = sSql & "           when " & pdValorRealiza & " *nvalor/1000 > Pc.nMontoMax then Pc.nMontoMax"
    sSql = sSql & "          End"
    sSql = sSql & " End"
    sSql = sSql & " from ProductoConcepto PC"
    sSql = sSql & " Inner Join ProductoConceptoFiltro PCF on PC.nPrdConceptoCod=PCF.nPrdConceptoCod"
    sSql = sSql & " Where nColocCred=2 and PC.nMoneda = " & nMoneda & " and PCF.cAgeCod=" & psCodAge & " and PCF.nProdCod=901 and"
    sSql = sSql & " PC.nInicial<= " & pdValorRealiza & " and PC.nFinal>=" & pdValorRealiza

    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaTotalLiberacion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecupera:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function InsertaLiberacion(ByVal Nrog As String, ByVal Comentario As String, _
ByVal cGLevanta As String, _
pdHoy As Date, psCodAge As String, psCodUser As String) As Long

Dim oCredito As DCredActualizaBD
Dim oFunciones As NContFunciones
Dim oGarantia As DGarantia
Dim rs As ADODB.Recordset

Dim sMovNro As String
Dim nMovNro As Double
Dim Total As Double
Dim Realizacion As Double
Dim Moneda As String
Dim lsDetalle As String
Dim TipoG As Integer
lsDetalle = ""

Set oFunciones = New NContFunciones
sMovNro = oFunciones.GeneraMovNro(pdHoy, psCodAge, psCodUser)
Set oFunciones = Nothing

Set oCredito = New DCredActualizaBD
Set oGarantia = New DGarantia
Set rs = oGarantia.RecuperaGarantia(Nrog)
Moneda = rs!nMoneda
Realizacion = rs!nRealizacion
TipoG = rs!nGarantReal
If TipoG = 1 Then
    'Solo para Garantias Relales
    Set rs = oGarantia.RecuperaTotalLiberacion(Realizacion, Moneda, psCodAge)
    Total = 0
    While Not rs.EOF
         Total = Total + rs!Monto
         rs.MoveNext
    Wend
    If rs.RecordCount > 0 Then rs.MoveFirst
'------->
End If

With oCredito
    .dBeginTrans
    Call .dInsertMov(sMovNro, cGLevanta, Comentario, gMovEstContabMovContable, gMovFlagVigente, False)
    nMovNro = oCredito.GetnMovNro(sMovNro)
    Call .dInsertMovGarantia(nMovNro, Nrog, cGLevanta, Total, "")
    If TipoG = 1 Then
      While Not rs.EOF
             Call .dInsertMovGarantDet(nMovNro, Nrog, cGLevanta, rs!ConceptoCod, rs!Monto)
          rs.MoveNext
        Wend
    End If
  'tipo accion 4 actualiza el estado de la garantia
    Call .dUpdateGarantias("", "", 0, 4, False, gPersGarantEstadoLevantada, Nrog)
   .dCommitTrans
End With
InsertaLiberacion = nMovNro
Set rs = Nothing
Set oCredito = Nothing
Set oGarantia = Nothing
End Function
Public Function RecuperaPesona_X_Credito(ByVal psCodPers As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    On Error GoTo ErrorRecupera
    
    sSql = " Select Pro.cCtaCod as Credito, "
    sSql = sSql & " (Select cConsDescripcion From Constante Where Pro.nPrdestado=nConsValor and nConsCod=3001) as Estado"
    sSql = sSql & " from Producto Pro "
    sSql = sSql & " Inner Join ColocacCred CC on CC.cCtaCod=Pro.cCtaCod "
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=CC.cCtaCod "
    sSql = sSql & " Where PP.nPrdPersRelac =" & gColRelPersTitular
    sSql = sSql & " and Pro.nPrdEstado in (" & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & ")"
    sSql = sSql & " and PP.cPersCod=" & "'" & psCodPers & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaPesona_X_Credito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecupera:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Error En Proceso", Err.Description


End Function
Public Function RecuperaNroGaran_X_Credito(ByVal Nrog As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    On Error GoTo ErrorRecupera
    
    
sSql = "SElect Pro.cCtaCod,"
sSql = sSql & " (Select cConsDescripcion From Constante Where Pro.nPrdestado=nConsValor and nConsCod=3001) Estado"
sSql = sSql & " from Garantias GR"
sSql = sSql & " Inner Join ColocGarantia CG on CG.cNumGarant = GR.cNumGarant"
sSql = sSql & " Inner Join ColocacCred CCC on CCC.cCtacod=CG.cCtaCod"
sSql = sSql & " Inner Join Producto Pro on CCC.cCtacod=Pro.cCtaCod"
sSql = sSql & " Where Pro.nPrdEstado in (" & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & ")"
sSql = sSql & " and GR.cNumGarant=" & "'" & Nrog & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaNroGaran_X_Credito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecupera:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


Public Function RecuperaGarantiasRealPersona(ByVal psCodPers As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    On Error GoTo ErrorRecuperaTiposDocumGarantias
    'sSql = " Select GR.cNumGarant as NroGarantia, P.cPersNombre as Nombre from PersGarantia PG"
    'sSql = sSql & " Inner Join Persona P on P.cPersCod=PG.cPersCod"
    'sSql = sSql & " Inner Join GarantReal GR on GR.cNumgarant=PG.cNumGarant"
    'sSql = sSql & " right outer Join Garantias G on G.cNumgarant=GR.cNumGarant"
    'sSql = sSql & " Where PG.nRelacion = 1 and G.nEstado=4" 'in (1,2,3,4)
    'sSql = sSql & " and P.cPersCod='" & psCodPers & "'"
    sSql = " Select G.cNumGarant as NroGarantia, dbo.pstanombre(P.cPersNombre,0) as Nombre"
    sSql = sSql & " from Garantias G"
    sSql = sSql & " Left outer Join GarantReal GR on GR.cNumgarant=G.cNumgarant"
    sSql = sSql & " Inner Join PersGarantia PG on PG.cNumgarant=G.cNumgarant"
    sSql = sSql & " Inner Join Persona P on P.cPersCod=PG.cPersCod"
    sSql = sSql & " Where PG.nRelacion = 1 and G.nEstado=4 and P.cPersCod='" & psCodPers & "'"

    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGarantiasRealPersona = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaTiposDocumGarantias:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function RecuperaGaratRealPersoRelac(ByVal Nrog As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaTiposDocumGarantias
    'sSql = " Select P.cPersCod as Codigo_Cliente, P.cPersNombre as Nombre,"
    'sSql = sSql & " (Select cConsDescripcion from Constante where PG.nRelacion=nConsValor and nConsCod=1021) as Relacion"
    'sSql = sSql & " from PersGarantia PG"
    'sSql = sSql & " Inner Join Persona P on P.cPersCod=PG.cPersCod"
    'sSql = sSql & " Inner Join GarantReal GR on GR.cNumgarant=PG.cNumGarant"
    'sSql = sSql & " right outer Join Garantias G on G.cNumgarant=GR.cNumGarant"
    'sSql = sSql & " where PG.cNumGarant='" & Nrog & "'"
    'sSql = sSql & " and G.nEstado=4"
    sSql = " Select P.cPersCod as Codigo_Cliente, dbo.pstanombre(P.cPersNombre,0) as Nombre,"
    sSql = sSql & " (Select cConsDescripcion from Constante where PG.nRelacion=nConsValor and nConsCod=1021) as Relacion"
    sSql = sSql & " from Garantias G"
    sSql = sSql & " Left outer Join GarantReal GR on GR.cNumgarant=G.cNumgarant"
    sSql = sSql & " Inner Join PersGarantia PG on PG.cNumgarant=G.cNumgarant"
    sSql = sSql & " Inner Join Persona P on P.cPersCod=PG.cPersCod"
    sSql = sSql & " Where G.cNumGarant='" & Nrog & "' and G.nEstado=4"
        
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGaratRealPersoRelac = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaTiposDocumGarantias:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


'------------------------
'------------------------

Public Function PerteneceACredito(ByVal psNumGarant As String) As Boolean
Dim sSql As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

    sSql = "Select CG.cNumgarant from ColocGarantia CG "
    sSql = sSql & " Inner Join Garantias G ON G.cNumgarant = CG.cNumgarant "
    sSql = sSql & " Inner Join Producto P ON P.cCtaCod = CG.cCtaCod "
    sSql = sSql & " Where G.nEstado not in (1,4,5) AND CG.cNumgarant = '" & psNumGarant & "' "
    sSql = sSql & " AND P.nPrdEstado >= 2020 "
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    If R.RecordCount > 0 Then
        PerteneceACredito = True
    Else
        PerteneceACredito = False
    End If
    R.Close
    Set R = Nothing

End Function

Public Function PerteneceACreditoVigente(ByVal psNumGarant As String) As Boolean
Dim sSql As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

    sSql = "Select CG.cNumgarant from ColocGarantia CG "
    sSql = sSql & " Inner Join Garantias G ON G.cNumgarant = CG.cNumgarant "
    sSql = sSql & " Inner Join Producto P ON P.cCtaCod = CG.cCtaCod "
    sSql = sSql & " Where G.nEstado not in (1,4,5) AND CG.cNumgarant = '" & psNumGarant & "' "
    sSql = sSql & " AND P.nPrdEstado in (" & gColocEstRefNorm & "," & gColocEstRefMor & "," & gColocEstRefVenc & "," & gColocEstVigNorm & "," & gColocEstVigMor & "," & gColocEstVigVenc & ") "
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    If R.RecordCount > 0 Then
        PerteneceACreditoVigente = True
    Else
        PerteneceACreditoVigente = False
    End If
    R.Close
    Set R = Nothing

End Function


Public Function PerteneceAGarantiaPendienteAsiento(ByVal psNumGarant As String, ByVal pdFecSis As Date) As Boolean
Dim sSql As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

    sSql = "Select CG.cNumgarant from ColocGarantia CG "
    sSql = sSql & " Inner Join Garantias G ON G.cNumgarant = CG.cNumgarant "
    sSql = sSql & " Inner Join Colocaciones C ON C.cCtaCod = CG.cCtaCod "
    sSql = sSql & " Inner Join Producto P ON P.cCtaCod = CG.cCtaCod "
    sSql = sSql & " Where G.nEstado not in (1,4,5) AND CG.cNumgarant = '" & psNumGarant & "' "
    sSql = sSql & " AND DATEDIFF(day,C.dVigencia,'" & Format(pdFecSis, "mm/dd/yyyy") & "') = 0 "
    sSql = sSql & " AND P.nPrdEstado in (" & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & ")"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    If R.RecordCount > 0 Then
        PerteneceAGarantiaPendienteAsiento = True
    Else
        PerteneceAGarantiaPendienteAsiento = False
    End If
    R.Close
    Set R = Nothing

End Function


Public Function RecuperaTiposDocumGarantias(ByVal pnTipoGarantia As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaTiposDocumGarantias
    
    sSql = "select D.cDocDesc, D.nDocTpo "
    sSql = sSql & " from Documento D Inner Join GarantDoc O ON D.nDocTpo = O.nDocTpo "
    sSql = sSql & " Where O.nConsValor =  " & pnTipoGarantia
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaTiposDocumGarantias = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaTiposDocumGarantias:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function


Public Function RecuperaTiposDocumentosGarantia() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaTiposDocumentosGarantia
    
    sSql = "select D.cDocDesc, D.nDocTpo "
    sSql = sSql & " from Documento D Inner Join GarantDoc O ON D.nDocTpo = O.nDocTpo "
    
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaTiposDocumentosGarantia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaTiposDocumentosGarantia:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function


'*****************************************************************************************
'***     Rutina:           RecuperaGarantiasPersona
'***     Descripcion:      Recupera todas las Garantias que Posee una Persona
'***     Modificado por:        NSSE
'***     Fecha-Tiempo:         11/06/2001 12:27:28 PM
'***     Ultima Modificacion: Creacion de la Rutina
'*****************************************************************************************
Public Function RecuperaGarantiasPersona(ByVal psPersCod As String, Optional ByVal pbMantenimiento As Boolean = False, _
Optional ByVal pbLiberacion As Boolean = False) As ADODB.Recordset

Dim oConecta As DConecta
Dim sSql As String

    On Error GoTo ErrorRecuperaGarantiasPersona
    sSql = " Select DISTINCT G.nEstado, C3.cConsDescripcion as cEstadoGar, G.cNumGarant, P.cPersNombre, G.cTpoDoc, G.cNroDoc, G.nTpoGarantia, C.cConsDescripcion as cTpoGarDescripcion, "
    sSql = sSql & " G.cDescripcion, G.nMoneda, C2.cConsdescripcion as cMonedaDesc,  G.nTasacion, G.nRealizacion, G.nGravament, G.nPorGravar, G.nPorGravar - G.nGravament as nDisponible, G.cPersCodEmisor, "
    sSql = sSql & " D.cDocDesc,G.nGarClase, G.nGarTpoRealiz, G.cBancoPersCod, PB.cPersNombre as cBanco,G.nGarantReal"
    sSql = sSql & " From Garantias G Inner Join PersGarantia PG ON G.cNumGarant = PG.cNumGarant  "
    sSql = sSql & "                  Inner Join Constante C ON G.nTpoGarantia = C.nConsValor AND C.nConsCod = " & gPersGarantia
    sSql = sSql & "                  Inner Join Constante C2 ON G.nMoneda = C2.nConsValor AND C2.nConsCod = " & gMoneda
    sSql = sSql & "                  Inner Join Persona P ON P.cPersCod = G.cPersCodEmisor "
    sSql = sSql & "                  Left Join Persona PB ON PB.cPersCod = G.cBancoPersCod "
    sSql = sSql & "                  Inner Join GarantDoc GD ON G.nTpoGarantia = GD.nConsValor AND GD.nDocTpo = CONVERT(int,G.cTpoDoc) "
    sSql = sSql & "                  Inner Join Documento D ON D.nDocTpo = GD.nDocTpo "
    sSql = sSql & "                  Inner Join Constante C3 ON G.nEstado = C3.nConsValor AND C3.nConsCod = 1030 "
    sSql = sSql & " WHERE PG.cPersCod = '" & psPersCod & "' AND PG.nRelacion = " & Trim(Str(gPersRelGarantiaTitular))
    If Not pbLiberacion Then
        If Not pbMantenimiento Then
            sSql = sSql & " AND G.nEstado not in (" & gPersGarantEstadoBloqueada & "," & gPersGarantEstadoLevantada & ") "
        Else
            sSql = sSql & " AND G.nEstado not in (" & gPersGarantEstadoBloqueada & ") "
        End If
    Else
        sSql = sSql & " AND G.nEstado not in (" & gPersGarantEstadoLevantada & ") "
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGarantiasPersona = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaGarantiasPersona:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


Public Function RecuperaMontoGarantiaCredito(ByVal psctacod As String) As Double
Dim sSql As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

    
    'sSql = " Select SUM(nTasacion) as nMonto From Garantias G "
    'sSql = sSql & " Inner Join Colocgarantia CG ON G.cnumGarant = CG.cNumGarant "
    'sSql = sSql & " Where CG.cCtaCod = '" & psCtaCod & "' AND G.nEstado = 1"
    
    'Aplicado al Valor de Construccion
    sSql = " Select SUM(ISNULL(GR.nValorConstruccion,0)) as nMonto From Garantias G "
    sSql = sSql & " Inner Join Colocgarantia CG ON G.cnumGarant = CG.cNumGarant "
    sSql = sSql & " Left Join GarantReal GR ON GR.cnumGarant = G.cNumGarant "
    sSql = sSql & " Where CG.cCtaCod = '" & psctacod & "' AND G.nEstado in (2,3)"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    RecuperaMontoGarantiaCredito = IIf(IsNull(R!nMonto), 0, R!nMonto)
    
End Function



Public Function RecuperaGarantiaCreditoDatosRefinan(ByVal psCtaCods As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaGarantiaCreditoRef
    sSql = "Select Distinct G.cNumgarant, CN.nConsValor as nTpoGar, G.cNroDoc, CN.cConsDescripcion cTpoGarantia, G.cDescripcion, D.cDocDesc, CG.nGravado, CN2.cConsDescripcion cMoneda, G.nPorGravar, G.nGravaMent, CN3.cConsDescripcion cMonedaGar, CG.nMoneda, CG.nEstado"
    sSql = sSql & " From ColocGarantia CG Inner Join Garantias G ON CG.cNumGarant = G.cNumGarant "
    sSql = sSql & "                       Inner Join Constante CN ON G.nTpoGarantia = CN.nConsValor AND CN.nConsCod = " & gPersGarantia
    sSql = sSql & "                       Inner Join Constante CN2 ON CN2.nConsValor = CG.nMoneda AND CN2.nConsCod = " & gMoneda
    sSql = sSql & "                       Inner Join Constante CN3 ON CN3.nConsValor = G.nMoneda AND CN3.nConsCod = " & gMoneda
    sSql = sSql & "                       Inner Join Documento D ON D.nDocTpo = Convert(int,G.cTpoDoc) "
    sSql = sSql & "                       Inner Join GarantDoc O ON D.nDocTpo = O.nDocTpo "
    sSql = sSql & " WHERE CG.cCtaCod in (" & psCtaCods & ")"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGarantiaCreditoDatosRefinan = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaGarantiaCreditoRef:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function



Public Function RecuperaGarantiaCreditoDatosVigente(ByVal psNumGarant As String)
Dim sSql As String
Dim oConecta As DConecta

    sSql = "Select CN1.cConsDescripcion cEstado, CG.cCtaCod, CN.cConsDescripcion as cMoneda, CG.nGravado, P.nSaldo, "
    sSql = sSql & " CE.nCuotas , C.nNroProxCuota"
    sSql = sSql & " From ColocGarantia CG Inner Join Producto P ON CG.cCtaCod = P.cCtaCod"
    sSql = sSql & " Inner Join Constante CN ON CN.nConsValor = convert(int,SUBSTRING(CG.cCtaCod,9,1)) AND CN.nConsCod = 1011"
    sSql = sSql & " Left Join ColocacEstado CE ON CG.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = 2002"
    sSql = sSql & " Inner Join ColocacCred C ON C.cCtaCod = CG.cCtaCod "
    sSql = sSql & " Inner Join Constante CN1 ON CN1.nConsCod = 3001 AND CN1.nConsValor = P.nPrdEstado "
    sSql = sSql & " Where cNumGarant = '" & psNumGarant & "' AND nEstado = 3"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGarantiaCreditoDatosVigente = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaGarantiaCredito(ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaGarantiaCredito
    sSql = "Select CN.nConsValor as nTpoGar, G.cNroDoc, CN.cConsDescripcion cTpoGarantia, G.cDescripcion, D.cDocDesc, CG.nGravado, CN2.cConsDescripcion cMoneda, G.nPorGravar, G.nGravaMent, CN3.cConsDescripcion cMonedaGar, G.nRealizacion as nRealizacion"
    sSql = sSql & " From ColocGarantia CG Inner Join Garantias G ON CG.cNumGarant = G.cNumGarant "
    sSql = sSql & "                       Inner Join Constante CN ON G.nTpoGarantia = CN.nConsValor AND CN.nConsCod = " & gPersGarantia
    sSql = sSql & "                       Inner Join Constante CN2 ON CN2.nConsValor = CG.nMoneda AND CN2.nConsCod = " & gMoneda
    sSql = sSql & "                       Inner Join Constante CN3 ON CN3.nConsValor = G.nMoneda AND CN3.nConsCod = " & gMoneda
    sSql = sSql & "                       Inner Join Documento D ON D.nDocTpo = Convert(int,G.cTpoDoc) "
    sSql = sSql & "                       Inner Join GarantDoc O ON D.nDocTpo = O.nDocTpo and O.nConsValor=G.nTpoGarantia"
    sSql = sSql & " WHERE CG.cCtaCod = '" & psctacod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGarantiaCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaGarantiaCredito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaGarantiaCreditoDatos(ByVal psctacod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaGarantiaCredito
    sSql = "Select CN.nConsValor as nTpoGar, G.cNroDoc, CN.cConsDescripcion cTpoGarantia, G.cDescripcion, D.cDocDesc, CG.nGravado, CN2.cConsDescripcion cMoneda, G.nPorGravar, G.nGravaMent, CN3.cConsDescripcion cMonedaGar,G.cNumgarant, CG.nMoneda, CG.nEstado"
    sSql = sSql & " From ColocGarantia CG Inner Join Garantias G ON CG.cNumGarant = G.cNumGarant "
    sSql = sSql & "                       Inner Join Constante CN ON G.nTpoGarantia = CN.nConsValor AND CN.nConsCod = " & gPersGarantia
    sSql = sSql & "                       Inner Join Constante CN2 ON CN2.nConsValor = CG.nMoneda AND CN2.nConsCod = " & gMoneda
    sSql = sSql & "                       Inner Join Constante CN3 ON CN3.nConsValor = G.nMoneda AND CN3.nConsCod = " & gMoneda
    sSql = sSql & "                       Inner Join Documento D ON D.nDocTpo = Convert(int,G.cTpoDoc) "
    sSql = sSql & "                       Inner Join GarantDoc O ON D.nDocTpo = O.nDocTpo  and O.nConsValor=G.nTpoGarantia"
    sSql = sSql & " WHERE CG.cCtaCod = '" & psctacod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGarantiaCreditoDatos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaGarantiaCredito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function



Public Function RecuperaGarantia(ByVal psNumGarant As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaGarantia
    
    sSql = "Select cTpoDoc, cNroDoc, nTpoGarantia, IdSupGarant, dGarantia, cDescripcion, cZona, nMoneda, "
    sSql = sSql & " nTasacion, nRealizacion, nGravament, nPorGravar, nEstado, cFlag, "
    sSql = sSql & " cComentario, cBancoPersCod, nGarantReal, nGarClase, nGarTpoRealiz, CN.cConsDescripcion cEstado "
    sSql = sSql & " From Garantias G Inner Join Constante CN ON G.nEstado = CN.nConsValor AND CN.nConsCod = 1030"
    sSql = sSql & " Where cNumGarant = '" & psNumGarant & "' "
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGarantia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaGarantia:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function RecuperaGarantiaReal(ByVal psNumGarant As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
            
    sSql = "Select cNumGarant, cPersCodSeguro, dEscritura, cRegistro, nRangHip, "
    sSql = sSql & " cPersCodVend, cPersVendTelef, cPersCodTasador, nTipVivienda,   "
    sSql = sSql & " cNroHip, cPersNotaria, PSeg.cPersNombre cSegPersNombre, "
    sSql = sSql & " PVend.cPersNombre cVendPersNombre,PTasa.cPersNombre cTasaPersNombre, PNota.cPersNombre cNotaPersNombre, "
    sSql = sSql & " GR.nCuotaInicial, GR.nMontoHipoteca, GR.nPrecioVenta, GR.nValorConstruccion, nNroPoliza, "
    sSql = sSql & " dVigenciaPol, nMontoPoliza, dConstitucion, dTasacion "
    sSql = sSql & " From GarantReal GR Left Join Persona PSeg ON GR.cPersCodSeguro = PSeg.cPersCod "
    sSql = sSql & "                     Left Join Persona PVend ON GR.cPersCodVend = PVend.cPersCod "
    sSql = sSql & "                     Left Join Persona PTasa ON GR.cPersCodTasador = PTasa.cPersCod "
    sSql = sSql & "                     Left Join Persona PNota ON GR.cPersNotaria = PNota.cPersCod "
    sSql = sSql & " Where cNumGarant = '" & psNumGarant & "' "
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGarantiaReal = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

'*****************************************************************************************
'***     Rutina:           RecuperaRelacPersonaGarantia
'***     Descripcion:       Recupera todas las Personas Relacionadas con la Garantia
'***     Modificado por:        NSSE
'***     Fecha-Tiempo:         11/06/2001 11:20:42 AM
'***     Ultima Modificacion: Creacion
'*****************************************************************************************
Public Function RecuperaRelacPersonaGarantia(ByVal psNumGarant As String)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaRelacPersonaGarantia
    sSql = "select P.cPersCod, P.cPersNombre, PG.cNumgarant, C.cConsDescripcion + space(50) + convert(varchar(2),PG.nRelacion)  as cRelacion "
    sSql = sSql & " from persgarantia PG Inner Join Persona P ON PG.cPersCod  = P.cPersCod "
    sSql = sSql & " inner join Constante C ON PG.nRelacion = C.nConsValor "
    sSql = sSql & " where C.nConsCod = '" & gPersRelGarantia & "' and PG.cNumGarant = '" & psNumGarant & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaRelacPersonaGarantia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaRelacPersonaGarantia:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function RecuperaGarantDeclaracionJur(ByVal psNumGarant As String)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaGarantDeclaracionJur
    sSql = "select DJ.cNumgarant, DJ.nItem, DJ.cGarDJDescripcion, DJ.nGarDJCantidad, DJ.nGarDJPrecioUnit, DJ.cGarDJTpoDoc, "
    sSql = sSql & " DJ.cGarDJNroDoc, C.cConsDescripcion + space(50) + DJ.cGarDJTpoDoc as cGarDJTpoDocDes "
    sSql = sSql & " from GarantDeclaracionJur DJ "
    sSql = sSql & " inner join Constante C ON DJ.cGarDJTpoDoc = C.nConsValor "
    sSql = sSql & " where C.nConsCod = '" & gColocPigTipoDocumento & "' and DJ.cNumGarant = '" & psNumGarant & "' ORDER BY DJ.cNumgarant, DJ.nItem"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGarantDeclaracionJur = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaGarantDeclaracionJur:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ExisteGarantia(ByVal psTipoDocGar As String, ByVal psNroDoc As String, _
 ByVal psPersCodEmi As String) As Boolean
Dim sql As String
Dim Co As DConecta
Dim rs As New ADODB.Recordset
Set Co = New DConecta
sql = " select * from Garantias " & _
      " where cTpoDoc='" & psTipoDocGar & "'" & _
      " and cNroDoc='" & psNroDoc & "'" & _
      " and cPersCodEmisor='" & psPersCodEmi & "'"
Co.AbreConexion
Set rs = Co.CargaRecordSet(sql)
Co.CierraConexion
If rs.EOF And rs.BOF Then
    ExisteGarantia = False
Else
    ExisteGarantia = True
End If
Set rs = Nothing
Set Co = Nothing
End Function

Public Sub NuevaGarantia(ByVal psTipoDocGar As String, ByVal psNroDoc As String, _
                ByVal psTipoGar As String, ByVal pnSTipoGarant As Integer, _
                ByVal psMoneda As String, ByVal psDescrip As String, _
                ByVal psZona As String, ByVal pnMontoTasac As Double, ByVal pnMontoReal As Double, _
                ByVal pnMontoxGrav As Double, ByVal psComent As String, RelPers As Variant, _
                ByVal dFecGar As Date, ByVal psPersCodEmi As String, ByVal psBancoPersCod As String, _
                ByVal pnGarantReal As Integer, ByVal pnGarClase As Integer, ByVal pnGarTpoRealiz As Integer, _
                ByVal psPersCodSeguro As String, ByVal pdEscritura As Date, _
                ByVal psRegistro As String, ByVal pnRangHip As Integer, ByVal psPersCodVend As String, _
                ByVal psPersVendTelef As String, ByVal psPersCodTasador As String, ByVal pnTipVivienda As Integer, _
                ByVal psNroHip As String, ByVal psPersNotaria As String, GarDetDJ As Variant, Optional ByVal pnCuotaInicial As Double = 0#, _
                Optional ByVal pnMontoHipoteca As Double = 0#, Optional ByVal pnPrecioVenta As Double = 0#, _
                Optional ByVal pnValorConstruccion As Double = 0#, Optional ByVal psNroPoliza As String, _
                Optional ByVal pdVigenciaPol As Date, Optional ByVal pnMontoPoliza As Double, _
                Optional ByVal pdConstitucion As Date, Optional ByVal dTasacion As Date)
                
                
Dim sSql As String
Dim oConecta As DConecta
Dim bTran As Boolean
Dim i As Integer
Dim sNumGarant As String
Dim R As ADODB.Recordset

                On Error GoTo ErrorNuevaGarantia
                
                bTran = False
                Set oConecta = New DConecta
                oConecta.AbreConexion
                oConecta.ConexionActiva.BeginTrans
                bTran = True
                
                sSql = "Select Max(cNumGarant) as cNumGarant From Garantias"
                Set R = oConecta.CargaRecordSet(sSql)
                sNumGarant = Format(CLng(IIf(IsNull(R!cNumGarant), "0", R!cNumGarant)) + 1, "00000000")
                
                sSql = "INSERT INTO Garantias(cNumGarant,cTpoDoc,cNroDoc,nTpoGarantia,dGarantia,cDescripcion,"
                sSql = sSql & "cZona,nMoneda,nTasacion,nRealizacion,nGravament,nPorGravar,nEstado,cComentario,cPersCodEmisor, "
                sSql = sSql & " cBancoPersCod, nGarantReal, nGarClase, nGarTpoRealiz,IdSupGarant)"
                sSql = sSql & " VALUES('" & sNumGarant & "','" & psTipoDocGar & "','" & psNroDoc & "','" & psTipoGar & "','"
                sSql = sSql & Format(dFecGar, "mm/dd/yyyy") & "','" & psDescrip & "','" & psZona & "',"
                sSql = sSql & psMoneda & "," & Format(pnMontoTasac, "#0.00") & "," & Format(pnMontoReal, "#0.00") & ","
                sSql = sSql & "0.00," & Format(pnMontoxGrav, "#0.00") & "," & gPersGarantEstadoRegistrado & ",'" & psComent & "','" & psPersCodEmi & "',"
                sSql = sSql & "'" & psBancoPersCod & "'," & pnGarantReal & "," & pnGarClase & "," & pnGarTpoRealiz & "," & pnSTipoGarant & ")"
                oConecta.ConexionActiva.Execute sSql
                
                If pnGarantReal = 1 Then
                    sSql = "INSERT INTO GarantReal(cNumGarant, cPersCodSeguro, dEscritura, "
                    sSql = sSql & " cRegistro, nRangHip, cPersCodVend, cPersVendTelef, cPersCodTasador, "
                    sSql = sSql & " nTipVivienda, cNroHip, cPersNotaria,nCuotaInicial,nMontoHipoteca,nPrecioVenta,nValorConstruccion,nNroPoliza, "
                    sSql = sSql & " dVigenciaPol,nMontoPoliza,dConstitucion,dTasacion )"
                    sSql = sSql & " VALUES('" & sNumGarant & "','" & psPersCodSeguro & "','" & Format(pdEscritura, "mm/dd/yyyy") & "','"
                    sSql = sSql & Trim(psRegistro) & "'," & pnRangHip & ",'" & psPersCodVend & "','" & psPersVendTelef & "','" & psPersCodTasador & "',"
                    sSql = sSql & pnTipVivienda & ",'" & psNroHip & "','" & psPersNotaria & "',"
                    sSql = sSql & Format(pnCuotaInicial, "#0.00") & "," & Format(pnMontoHipoteca, "#0.00") & "," & Format(pnPrecioVenta, "#0.00") & "," & Format(pnValorConstruccion, "#0.00")
                    sSql = sSql & ",'" & psNroPoliza & "','" & Format(pdVigenciaPol, "mm/dd/yyyy") & "'," & Format(pnMontoPoliza, "#0.00")
                    sSql = sSql & ",'" & Format(pdConstitucion, "mm/dd/yyyy") & "','" & Format(dTasacion, "mm/dd/yyyy") & "')"
                    
                    oConecta.ConexionActiva.Execute sSql
                End If
                
                ' CMACICA_CSTS - 25/11/2003 --------------------------------------------------------------------------
                'If psTipoGar = gPersGarantiaOtrasGarantias Then
                If psTipoDocGar = 15 Then ' Si es declaracion jurada
                   For i = 0 To UBound(GarDetDJ) - 1
                       sSql = "INSERT INTO GARANTDECLARACIONJUR(cNumGarant, nItem, cGarDJDescripcion, nGarDJCantidad, nGarDJPrecioUnit, cGarDJTpoDoc, cGarDJNroDoc)"
                       sSql = sSql & " VALUES('" & sNumGarant & "', "
                       sSql = sSql & "  " & GarDetDJ(i, 0) & ", "
                       sSql = sSql & "  '" & GarDetDJ(i, 1) & "', "
                       sSql = sSql & " " & GarDetDJ(i, 2) & ", "
                       sSql = sSql & " " & CCur(GarDetDJ(i, 3)) & ", "
                       sSql = sSql & " '" & GarDetDJ(i, 4) & "', "
                       sSql = sSql & "  '" & GarDetDJ(i, 5) & "') "
                       oConecta.ConexionActiva.Execute sSql
                   Next i
                End If
                ' ----------------------------------------------------------------------------------------------------
                
                
                For i = 0 To UBound(RelPers) - 1
                    sSql = "INSERT INTO PERSGARANTIA(cPersCod, cNumGarant, nRelacion)"
                    sSql = sSql & " VALUES('" & RelPers(i, 0) & "', "
                    sSql = sSql & " '" & sNumGarant & "', "
                    sSql = sSql & " " & RelPers(i, 3) & ") "
                    oConecta.ConexionActiva.Execute sSql
                Next i
                oConecta.ConexionActiva.CommitTrans
                oConecta.CierraConexion
                Set oConecta = Nothing
                
                
                Exit Sub

ErrorNuevaGarantia:
        If bTran Then
            oConecta.ConexionActiva.RollbackTrans
        End If
        Set oConecta = Nothing
        If Err.Number = -2147217873 Then
            MsgBox "Garantia Duplicada"
            Exit Sub
        End If
        MsgBox Err.Description, vbCritical, "Aviso"
                
End Sub

Public Function GarantiaEnUso(ByVal psNumGarant As String) As Boolean
Dim sSql As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

    sSql = "Select * From ColocGarantia Where cNumGarant  = '" & psNumGarant & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    If R.RecordCount > 0 Then
        GarantiaEnUso = True
    Else
        GarantiaEnUso = False
    End If
    R.Close
    Set R = Nothing
End Function

Public Sub EliminarGraantia(ByVal psNumGarant As String)
Dim sSql As String
Dim oConecta As DConecta
Dim bTran As Boolean

    On Error GoTo ErrorEliminarGraantia
    bTran = False
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.BeginTrans
    bTran = True
    'Elimina Relacion de Garantias de Credito
    sSql = "DELETE ColocGarantia WHERE cNumGarant = '" & psNumGarant & "' "
    oConecta.ConexionActiva.Execute sSql
    'Elimina garantia Real
    sSql = "DELETE GarantReal WHERE cNumGarant = '" & psNumGarant & "' "
    oConecta.ConexionActiva.Execute sSql
    
    ' CMACICA_CSTS - 25/11/2003 ------------------------------------------------
    'Elimina garantia DECLARACION JURADA
    sSql = "DELETE GarantDeclaracionJur WHERE cNumGarant = '" & psNumGarant & "' "
    oConecta.ConexionActiva.Execute sSql
    '---------------------------------------------------------------------------
    
    'Elimina PersGarantia
    sSql = "DELETE PersGarantia WHERE cNumGarant = '" & psNumGarant & "' "
    oConecta.ConexionActiva.Execute sSql
    'Elimina Garantias
    sSql = "DELETE Garantias WHERE cNumGarant = '" & psNumGarant & "'"
    oConecta.ConexionActiva.Execute sSql
    oConecta.ConexionActiva.CommitTrans
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorEliminarGraantia:
    If bTran Then
        oConecta.ConexionActiva.RollbackTrans
        Set oConecta = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso EliminarGraantia", Err.Description

End Sub

Public Sub ActualizaGarantia(ByVal psNumGarant As String, ByVal psTipoDocGar As String, ByVal psNroDoc As String, _
                ByVal psTipoGar As String, ByVal pnSTipoGarant As Integer, _
                ByVal psMoneda As String, ByVal psDescrip As String, _
                ByVal psZona As String, ByVal pnMontoTasac As Double, ByVal pnMontoReal As Double, _
                ByVal pnMontoxGrav As Double, ByVal psComent As String, RelPers As Variant, _
                ByVal dFecGar As Date, ByVal sPersCodEmi As String, ByVal psBancoPersCod As String, _
                ByVal pnGarantReal As Integer, ByVal pnGarClase As Integer, ByVal pnGarTpoRealiz As Integer, _
                ByVal psPersCodSeguro As String, ByVal pdEscritura As Date, _
                ByVal psRegistro As String, ByVal pnRangHip As Integer, ByVal psPersCodVend As String, _
                ByVal psPersVendTelef As String, ByVal psPersCodTasador As String, ByVal pnTipVivienda As Integer, _
                ByVal psNroHip As String, ByVal psPersNotaria As String, GarDetDJ As Variant, Optional ByVal pnCuotaInicial As Double = 0#, _
                Optional ByVal pnMontoHipoteca As Double = 0#, Optional ByVal pnPrecioVenta As Double = 0#, _
                Optional ByVal pnValorConstruccion As Double = 0#, Optional ByVal psNroPoliza As String, _
                Optional ByVal pdVigenciaPol As Date, Optional ByVal pnMontoPoliza As Double, _
                Optional ByVal pdConstitucion As Date, Optional ByVal dTasacion As Date)
                
Dim sSql As String
Dim oConecta As DConecta
Dim bTran As Boolean
Dim i As Integer
            On Error GoTo ErrorActualizaGarantia
            bTran = False
            Set oConecta = New DConecta
            oConecta.AbreConexion
            oConecta.ConexionActiva.BeginTrans
            bTran = True
            sSql = "UPDATE Garantias SET "
            sSql = sSql & " nTpoGarantia = " & psTipoGar & ", "
            sSql = sSql & " cTpoDoc = " & psTipoDocGar & ", "
            sSql = sSql & " cNroDoc = '" & psNroDoc & "', "
            sSql = sSql & " dGarantia = '" & Format(dFecGar, "mm/dd/yyyy") & "', "
            sSql = sSql & " cDescripcion = '" & psDescrip & "',"
            sSql = sSql & " cZona = '" & psZona & "', "
            sSql = sSql & " nMoneda = " & psMoneda & ", "
            sSql = sSql & " nTasacion = " & Format(pnMontoTasac, "#0.00") & ", "
            sSql = sSql & " nRealizacion = " & Format(pnMontoReal, "#0.00") & ", "
            'sSql = sSql & " nGravament = " & Format(pnMontoxGrav, "#0.00") & ", "
            sSql = sSql & " nPorGravar = " & Format(pnMontoxGrav, "#0.00") & ", "
            sSql = sSql & " cComentario = ' " & psComent & "', "
            sSql = sSql & " cBancoPersCod = '" & psBancoPersCod & "',"
            sSql = sSql & " nGarantReal = " & pnGarantReal & ","
            sSql = sSql & " nGarClase = " & pnGarClase & ","
            sSql = sSql & " nGarTpoRealiz = " & pnGarTpoRealiz & ", "
            sSql = sSql & " cPersCodEmisor = '" & sPersCodEmi & "', "
            sSql = sSql & "IdSupGarant=" & pnSTipoGarant
            sSql = sSql & " WHERE cNumGarant = '" & psNumGarant & "' "
            oConecta.ConexionActiva.Execute sSql
            
            If pnGarantReal = 1 Then
                sSql = "UPDATE GarantReal SET "
                sSql = sSql & " cPersCodSeguro = '" & psPersCodSeguro & "',"
                sSql = sSql & " dEscritura = '" & Format(pdEscritura, "mm/dd/yyyy") & "',"
                sSql = sSql & " cRegistro = '" & Trim(psRegistro) & "', "
                sSql = sSql & " nRangHip = " & pnRangHip & ", "
                sSql = sSql & " cPersCodVend = '" & psPersCodVend & "', "
                sSql = sSql & " cPersVendTelef = '" & psPersVendTelef & "', "
                sSql = sSql & " cPersCodTasador = '" & psPersCodTasador & "', "
                sSql = sSql & " nTipVivienda = " & pnTipVivienda & ", "
                sSql = sSql & " cNroHip = '" & psNroHip & "', "
                sSql = sSql & " cPersNotaria = '" & psPersNotaria & "', "
                sSql = sSql & " nCuotaInicial = " & Format(pnCuotaInicial, "#0.00") & ","
                sSql = sSql & " nMontoHipoteca = " & Format(pnMontoHipoteca, "#0.00") & ","
                sSql = sSql & " nPrecioVenta = " & Format(pnPrecioVenta, "#0.00") & ","
                sSql = sSql & " nValorConstruccion = " & Format(pnValorConstruccion, "#0.00") & ","
                sSql = sSql & " nNroPoliza = '" & psNroPoliza & "',"
                sSql = sSql & " dVigenciaPol = '" & Format(pdVigenciaPol, "mm/dd/yyyy") & "',"
                sSql = sSql & " nMontoPoliza = " & Format(pnMontoPoliza, "#0.00") & ","
                sSql = sSql & " dConstitucion = '" & Format(pdConstitucion, "mm/dd/yyyy") & "',"
                sSql = sSql & " dTasacion = '" & Format(dTasacion, "mm/dd/yyyy") & "'"
                sSql = sSql & " WHERE cNumGarant = '" & psNumGarant & "' "
                
                oConecta.ConexionActiva.Execute sSql
            End If
            
            
            sSql = "DELETE PersGarantia WHERE cNumGarant = '" & psNumGarant & "'"
            oConecta.ConexionActiva.Execute sSql
            
            For i = 0 To UBound(RelPers) - 1
                sSql = "INSERT INTO PERSGARANTIA(cPersCod, cNumGarant, nRelacion) "
                sSql = sSql & " VALUES('" & RelPers(i, 0) & "', "
                sSql = sSql & " '" & psNumGarant & "', "
                sSql = sSql & " " & RelPers(i, 3) & ") "
                oConecta.ConexionActiva.Execute sSql
            Next i
            
            ' CMACICA_CSTS - 25/11/2003 --------------------------------------------------------------------
            'If psTipoGar = gPersGarantiaOtrasGarantias Then
            If psTipoDocGar = 15 Then ' declaracion jurada
               sSql = "DELETE GarantDeclaracionJur WHERE cNumGarant = '" & psNumGarant & "'"
               oConecta.ConexionActiva.Execute sSql
                
               For i = 0 To UBound(GarDetDJ) - 1
                   sSql = "INSERT INTO GARANTDECLARACIONJUR(cNumGarant, nItem, cGarDJDescripcion, nGarDJCantidad, nGarDJPrecioUnit, cGarDJTpoDoc, cGarDJNroDoc)"
                   sSql = sSql & " VALUES('" & psNumGarant & "', "
                   sSql = sSql & "  " & GarDetDJ(i, 0) & ", "
                   sSql = sSql & "  '" & GarDetDJ(i, 1) & "', "
                   sSql = sSql & " " & GarDetDJ(i, 2) & ", "
                   sSql = sSql & " " & CCur(GarDetDJ(i, 3)) & ", "
                   sSql = sSql & " '" & GarDetDJ(i, 4) & "', "
                   sSql = sSql & "  '" & GarDetDJ(i, 5) & "') "
                   oConecta.ConexionActiva.Execute sSql
               Next i
            End If
            '------------------------------------------------------------------------------------------------
            
            oConecta.ConexionActiva.CommitTrans
            oConecta.CierraConexion
            Set oConecta = Nothing
    Exit Sub

ErrorActualizaGarantia:
        If bTran Then
            oConecta.ConexionActiva.RollbackTrans
        End If
        Set oConecta = Nothing
        Err.Raise Err.Number, "Error En Proceso ActualizaGarantia", Err.Description
End Sub

Public Function ListaGarantias(ByVal psCodPersona As String) As Recordset
    'Lista a las garantias donde la persona es el titular y muestra la descripcion
     Dim oConecta As DConecta
     Dim strSQL As String
     On Error GoTo ErrHandler
        Set oConecta = New DConecta
        oConecta.AbreConexion
        strSQL = "select c.cDescripcion,c.cNumGarant "
        strSQL = strSQL & " From Persona a "
        strSQL = strSQL & "inner join PersGarantia b on a.cPersCod=b.cPersCod "
        strSQL = strSQL & " inner join Garantias c on b.cNumGarant=c.cNumGarant "
        strSQL = strSQL & " where  b.nRelacion=1 and a.cPersCod='" & psCodPersona & "'"
        
        Set ListaGarantias = oConecta.CargaRecordSet(strSQL)
        
        oConecta.CierraConexion
        Set oConecta = Nothing
    Exit Function
ErrHandler:
    If Not oConecta Is Nothing Then Set oConecta = Nothing
    Set ListaGarantias = Nothing
End Function


Public Function DescripGarantia(ByVal cNumGarant As String) As Recordset
    Dim oConecta As DConecta
    Dim strSQL As String
    On Error GoTo ErrHandler
        Set oConecta = New DConecta
        oConecta.AbreConexion
        strSQL = "select a.nGravament,a.nTasacion,a.nRealizacion, "
        strSQL = strSQL & " cEstado=(Select cConsDescripcion From Constante where nconscod='1030' and nConsValor=a.nEstado) , "
        strSQL = strSQL & " a.nEstado, p.cPersCod, p.cPersNombre  "
        strSQL = strSQL & " From garantias a "
        strSQL = strSQL & " Inner join PersGarantia pg on pg.cNumGarant=a.cNumGarant "
        strSQL = strSQL & " Inner join Persona p on pg.cPersCod=p.cPersCod "
        strSQL = strSQL & " where a.cNumGarant='" & cNumGarant & "'"
        
         Set DescripGarantia = oConecta.CargaRecordSet(strSQL)
         oConecta.CierraConexion
         Set oConecta = Nothing
    Exit Function
ErrHandler:
    If Not oConecta Is Nothing Then Set oConecta = Nothing
    Set DescripGarantia = Nothing
End Function

Public Function InsertarGarantiaPreferidas(ByVal pcMovNro As String, ByVal pCnumGarant As String, _
                                           ByVal nMonto As Double, ByVal pcPersCod As String, _
                                           ByVal nEstado As Integer) As Boolean
    Dim oConecta As DConecta
    Dim strSQL As String
    Dim strMovDesc As String
    Dim intMovNro As Long
    Dim rs As ADODB.Recordset
    Dim bTran As Boolean
    On Error GoTo ErrHandler
        bTran = False
        strMovDesc = "Movimiento por Garantias Preferidas"
        Set oConecta = New DConecta
        oConecta.AbreConexion
        oConecta.BeginTrans
        bTran = True
        'Insertando movimiento
        strSQL = "Insert into Mov(cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag)values('"
        strSQL = strSQL & pcMovNro & "'," & 172700 & ",'" & strMovDesc & "'," & 0 & "," & 0 & ")"
        oConecta.Ejecutar (strSQL)
        
        Set rs = New ADODB.Recordset
        strSQL = "Select @@Identity as MovNro"
        Set rs = oConecta.CargaRecordSet(strSQL)
        If Not rs.EOF And Not rs.BOF Then
            intMovNro = IIf(IsNull(rs!MovNro), 0, rs!MovNro)
        End If
        Set rs = Nothing
        
        If intMovNro <> 0 Then
            ' se inserta en movgarantia
            strSQL = "Insert Into MovGarantia(nMovNro,cNumGarant,cOpeCod,nMonto,cDescripcion)Values( "
            strSQL = strSQL & intMovNro & ",'" & pCnumGarant & "'," & 172700 & "," & nMonto & ", 'Movimiento por Garantia Preferida')"
            oConecta.Ejecutar (strSQL)
            
            strSQL = "Insert Into ColocCalifGarantia (cPersCod, cNumGarant, nEstadoGar, nMovNro) Values('" & pcPersCod & "','" & pCnumGarant & "'," & nEstado & "," & intMovNro & ")"
            oConecta.Ejecutar (strSQL)
        Else
            If bTran Then
                oConecta.RollbackTrans
            End If
            InsertarGarantiaPreferidas = False
            If oConecta Is Nothing Then Set oConecta = Nothing
            Exit Function
        End If
        oConecta.CommitTrans
        oConecta.CierraConexion
        Set oConecta = Nothing
        InsertarGarantiaPreferidas = True
    Exit Function
ErrHandler:
    If Not oConecta Is Nothing Then Set oConecta = Nothing
     InsertarGarantiaPreferidas = False
End Function


Public Function ListaSuperGarantias() As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    On Error GoTo ErrHandler
    
    sSql = "Select nConsValor,cConsDescripcion from constante where nconscod='1048' and nConsValor<>'1048'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaSuperGarantias = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ListaSuperGarantias = Null
End Function

Public Function CargarRelGarantia(ByVal pIsSuperTipo As Integer) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    On Error GoTo ErrHandler
        sSql = "Select C.nConsValor,C.cConsDescripcion,R.IdInfGarant "
        sSql = sSql & " From relgarant R"
        sSql = sSql & " Inner join Constante C on C.nConsValor=R.IdInfGarant"
        sSql = sSql & " Where IdSupGarant=" & pIsSuperTipo & " and C.nConsCod='1027' ORDER BY cConsDescripcion "
     Set oConec = New DConecta
     oConec.AbreConexion
     Set CargarRelGarantia = oConec.CargaRecordSet(sSql)
     oConec.CierraConexion
     Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set CargarRelGarantia = Null
End Function

Public Function ObtenerIdSuperGarantia(ByVal pIdTipoGarantia As Integer) As Integer
    Dim oConec As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
        On Error GoTo ErrHandler
        sSql = "Select IdSupGarant From relgarant Where IdInfGarant=" & pIdTipoGarantia
        
        Set oConec = New DConecta
        oConec.AbreConexion
        Set rs = oConec.CargaRecordSet(sSql)
        If Not rs.EOF And Not rs.BOF Then
            ObtenerIdSuperGarantia = rs!IdSupGarant
        End If
        Set rs = Nothing
        oConec.CierraConexion
        Set oConec = Nothing
        Exit Function
ErrHandler:
        If Not oConec Is Nothing Then Set oConec = Nothing
        If Not rs Is Nothing Then Set rs = Nothing
        ObtenerIdSuperGarantia = -1
End Function


Public Function ValidaPFCTS(ByVal psctacod As String) As Boolean
    Dim oConec As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim nValor As Integer
    
    sSql = "Select Count(*) as nCantidad"
    sSql = sSql & " From Producto"
    sSql = sSql & " Where SubString(cCtaCod,6,3) in ('233','234') and nPrdEstado not in (1300,1400)"
    sSql = sSql & " and cCtaCod='" & psctacod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    
    If Not rs.EOF And Not rs.BOF Then
        nValor = rs!nCantidad
    End If
    Set rs = Nothing
    
    If nValor > 0 Then
        ValidaPFCTS = True
    Else
       ValidaPFCTS = False
    End If
End Function

