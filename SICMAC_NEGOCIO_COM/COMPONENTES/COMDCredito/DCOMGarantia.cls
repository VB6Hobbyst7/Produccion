VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim csConexion As String
Dim csNegocio As String
Dim csCentralPer As String
Dim csCentralCom As String
Dim csCentralImg As String
Dim csAdminist As String

Dim oError As COMConecta.COMErrorHandling
Public coConex As COMConecta.DCOMConecta

Public Function RecuperaGarantiasxDatos(ByVal psTipoGar As String, ByVal psTipoDoc As String, ByVal psNumDoc As String) As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
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
    
    Set oConecta = New COMConecta.DCOMConecta
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
Dim oConecta As COMConecta.DCOMConecta

    
    
           sSql = " Update Mov "
    sSql = sSql & " Set nMovFlag=" & gMovFlagExtornado
    sSql = sSql & " Where cMovNro =" & "'" & NroMov & "'"
    'sSql = sSql & "                "
    sSql2 = sSql2 & " Update Garantias"
    sSql2 = sSql2 & " Set nEstado = 4"
    sSql2 = sSql2 & " Where cNumGarant =" & Nrog

    Set oConecta = New COMConecta.DCOMConecta
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
Dim oConecta As COMConecta.DCOMConecta


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

    Set oConecta = New COMConecta.DCOMConecta
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
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecupera
    
 
sSql = "Select PG.cPersCod as CodPers, P.cPersNombre as Nombre, GR.nGarantReal From Garantias GR"
sSql = sSql & " Inner Join PersGarantia PG on PG.cNumGarant=GR.cNumGarant"
sSql = sSql & " Inner Join Persona P on P.cPersCod=Pg.cPersCod"
sSql = sSql & " Where GR.cNumGarant ='" & Nrog & "'"
sSql = sSql & " And PG.nRelacion = 1"

    
    Set oConecta = New COMConecta.DCOMConecta
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
Dim oConecta As COMConecta.DCOMConecta

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

    
    Set oConecta = New COMConecta.DCOMConecta
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

Dim oCredito As DCOMCredActBD
Dim oFunciones As COMNContabilidad.NCOMContFunciones
Dim oGarantia As DCOMGarantia
Dim rs As ADODB.Recordset

Dim sMovNro As String
Dim nMovNro As Double
Dim Total As Double
Dim Realizacion As Double
Dim Moneda As String
Dim lsDetalle As String
Dim TipoG As Integer
lsDetalle = ""

Set oFunciones = New COMNContabilidad.NCOMContFunciones
sMovNro = oFunciones.GeneraMovNro(pdHoy, psCodAge, psCodUser)
Set oFunciones = Nothing

Set oCredito = New DCOMCredActBD
Set oGarantia = New DCOMGarantia
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
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorRecupera
    
    sSql = " Select Pro.cCtaCod as Credito, "
    sSql = sSql & " (Select cConsDescripcion From Constante Where Pro.nPrdestado=nConsValor and nConsCod=3001) as Estado"
    sSql = sSql & " from Producto Pro "
    sSql = sSql & " Inner Join ColocacCred CC on CC.cCtaCod=Pro.cCtaCod "
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=CC.cCtaCod "
    sSql = sSql & " Where PP.nPrdPersRelac =" & gColRelPersTitular
    sSql = sSql & " and Pro.nPrdEstado in (" & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & ")"
    sSql = sSql & " and PP.cPersCod=" & "'" & psCodPers & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
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
Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorRecupera
    
    
sSql = "SElect Pro.cCtaCod,"
sSql = sSql & " (Select cConsDescripcion From Constante Where Pro.nPrdestado=nConsValor and nConsCod=3001) Estado"
sSql = sSql & " from Garantias GR"
sSql = sSql & " Inner Join ColocGarantia CG on CG.cNumGarant = GR.cNumGarant"
sSql = sSql & " Inner Join ColocacCred CCC on CCC.cCtacod=CG.cCtaCod"
sSql = sSql & " Inner Join Producto Pro on CCC.cCtacod=Pro.cCtaCod"
sSql = sSql & " Where Pro.nPrdEstado in (" & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & ")"
sSql = sSql & " and GR.cNumGarant=" & "'" & Nrog & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
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
Dim oConecta As COMConecta.DCOMConecta
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

    
    Set oConecta = New COMConecta.DCOMConecta
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
Dim oConecta As COMConecta.DCOMConecta

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
        
    Set oConecta = New COMConecta.DCOMConecta
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
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    sSql = "Select CG.cNumgarant from ColocGarantia CG "
    sSql = sSql & " Inner Join Garantias G ON G.cNumgarant = CG.cNumgarant "
    sSql = sSql & " Inner Join Producto P ON P.cCtaCod = CG.cCtaCod "
    sSql = sSql & " Where G.nEstado not in (1,4,5) AND CG.cNumgarant = '" & psNumGarant & "' "
    sSql = sSql & " AND P.nPrdEstado >= 2020 "
    
    Set oConecta = New COMConecta.DCOMConecta
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
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    sSql = "Select CG.cNumgarant from ColocGarantia CG "
    sSql = sSql & " Inner Join Garantias G ON G.cNumgarant = CG.cNumgarant "
    sSql = sSql & " Inner Join Producto P ON P.cCtaCod = CG.cCtaCod "
    sSql = sSql & " Where G.nEstado not in (1,4,5) AND CG.cNumgarant = '" & psNumGarant & "' "
    sSql = sSql & " AND P.nPrdEstado in (" & gColocEstRefNorm & "," & gColocEstRefMor & "," & gColocEstRefVenc & "," & gColocEstVigNorm & "," & gColocEstVigMor & "," & gColocEstVigVenc & ") "
    
    Set oConecta = New COMConecta.DCOMConecta
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
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    sSql = "Select CG.cNumgarant from ColocGarantia CG "
    sSql = sSql & " Inner Join Garantias G ON G.cNumgarant = CG.cNumgarant "
    sSql = sSql & " Inner Join Colocaciones C ON C.cCtaCod = CG.cCtaCod "
    sSql = sSql & " Inner Join Producto P ON P.cCtaCod = CG.cCtaCod "
    sSql = sSql & " Where G.nEstado not in (1,4,5) AND CG.cNumgarant = '" & psNumGarant & "' "
    sSql = sSql & " AND DATEDIFF(day,C.dVigencia,'" & Format(pdFecSis, "mm/dd/yyyy") & "') = 0 "
    sSql = sSql & " AND P.nPrdEstado in (" & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & ")"
    
    Set oConecta = New COMConecta.DCOMConecta
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
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaTiposDocumGarantias
    
    sSql = "select D.cDocDesc, D.nDocTpo "
    sSql = sSql & " from Documento D Inner Join GarantDoc O ON D.nDocTpo = O.nDocTpo "
    sSql = sSql & " Where O.nConsValor =  " & pnTipoGarantia
    
    Set oConecta = New COMConecta.DCOMConecta
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
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaTiposDocumentosGarantia
    
    sSql = "select D.cDocDesc, D.nDocTpo "
    sSql = sSql & " from Documento D Inner Join GarantDoc O ON D.nDocTpo = O.nDocTpo "
    
    
    Set oConecta = New COMConecta.DCOMConecta
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
'RecuperaDatosGarantiasPersona
Public Function RecuperaGarantiasPersona(ByVal psPersCod As String, Optional ByVal pbMantenimiento As Boolean = False, _
                                        Optional ByVal pbLiberacion As Boolean = False, Optional pbGarantClase As Boolean = False) As ADODB.Recordset

Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String

    On Error GoTo ErrorRecuperaGarantiasPersona
    sSql = " Select DISTINCT G.nEstado, C3.cConsDescripcion as cEstadoGar, G.cNumGarant, P.cPersNombre, G.cTpoDoc, G.cNroDoc, G.nTpoGarantia, C.cConsDescripcion as cTpoGarDescripcion, "
    If pbGarantClase Then
        sSql = sSql & " (G.cDescripcion + ' ' + G.cDireccion) cDescripcion, "
    Else
        sSql = sSql & " G.cDescripcion, "
    End If
    sSql = sSql & " G.nMoneda, C2.cConsdescripcion as cMonedaDesc,  G.nTasacion, G.nRealizacion, G.nGravament, G.nPorGravar, G.nPorGravar - G.nGravament as nDisponible, G.cPersCodEmisor, "
    sSql = sSql & " D.cDocDesc,G.nGarClase, G.nGarTpoRealiz, G.cBancoPersCod, PB.cPersNombre as cBanco,G.nGarantReal"
    'By Capi 12082008
    'sSql = sSql & " ,G.cNumGarant" 'ARCV 11-07-2006
    sSql = sSql & " ,G.cNumGarant,Isnull(C4.cConsDescripcion,'') as PFestado" 'ARCV 11-07-2006
    '
    'ALPA 20080811****
    If pbGarantClase Then
    sSql = sSql & " ,CCG.cCtaCod,isnull(nEstadoAdju,0) as nEstadoAdju,isnull(dEstadoAdju,'1900/01/01') dEstadoAdju,isnull(cUsuariAdju,'') cUsuariAdju "
    End If
    '*****************
    sSql = sSql & " From Garantias G Inner Join PersGarantia PG ON G.cNumGarant = PG.cNumGarant  "
    sSql = sSql & "                  Inner Join Constante C ON G.nTpoGarantia = C.nConsValor AND C.nConsCod = " & gPersGarantia
    sSql = sSql & "                  Inner Join Constante C2 ON G.nMoneda = C2.nConsValor AND C2.nConsCod = " & gMoneda
    sSql = sSql & "                  Inner Join Persona P ON P.cPersCod = G.cPersCodEmisor "
    'ALPA 20080811*****
    If pbGarantClase Then
    sSql = sSql & " inner join ColocGarantia CCG on (PG.cNumGarant=CCG.cNumGarant)"
    sSql = sSql & " inner join ProductoPersona PP3 on (PG.cPersCod=PP3.cPersCod)"
    sSql = sSql & " and (CCG.cCtaCod=PP3.cCtaCod) and PP3.nPrdPersRelac=20 "
    End If
    '*****************
    sSql = sSql & "                  Left Join Persona PB ON PB.cPersCod = G.cBancoPersCod "
    sSql = sSql & "                  Inner Join GarantDoc GD ON G.nTpoGarantia = GD.nConsValor AND GD.nDocTpo = CONVERT(int,G.cTpoDoc) "
    sSql = sSql & "                  Inner Join Documento D ON D.nDocTpo = GD.nDocTpo "
    sSql = sSql & "                  Inner Join Constante C3 ON G.nEstado = C3.nConsValor AND C3.nConsCod = 1030 "
    'By Capi 12082008 para que obtenga el estado del plazo fijo
    sSql = sSql & "                  Left Join Producto Pr On Pr.cCtaCod=RTrim(G.cNroDoc)"
    sSql = sSql & "                  Left Join Constante C4 On Pr.nPrdEstado=C4.nConsValor And C4.nConsCod=2001"
    '
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
    If pbGarantClase Then
    'sSql = sSql & " AND G.nGarClase = 1 and nEstadoAdju in (7,8,9) "
    sSql = sSql & " AND G.nTpoGarantia in (1,2,3,4,5) and nEstadoAdju in (7,8,9) "
    End If
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaGarantiasPersona = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaGarantiasPersona:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


Public Function RecuperaMontoGarantiaCredito(ByVal psCtaCod As String, Optional ByVal pbConstruccion As Boolean = False) As Double
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    
    'sSql = " Select SUM(nTasacion) as nMonto From Garantias G "
    'sSql = sSql & " Inner Join Colocgarantia CG ON G.cnumGarant = CG.cNumGarant "
    'sSql = sSql & " Where CG.cCtaCod = '" & psCtaCod & "' AND G.nEstado = 1"
    'EJVG20150702 ***
    'If pbConstruccion Then
    '    'Aplicado al Valor de Construccion
    '    sSql = " Select SUM(ISNULL(GR.nValorConstruccion,0)) as nMonto From Garantias G "
    '    sSql = sSql & " Inner Join Colocgarantia CG ON G.cnumGarant = CG.cNumGarant "
    '    sSql = sSql & " Left Join GarantRealTasacion GR ON GR.cnumGarant = G.cNumGarant "
    '    sSql = sSql & " Where CG.cCtaCod = '" & psCtaCod & "' AND G.nEstado in (2,3)"
    'Else
    '    sSql = " Select SUM(nTasacion) as nMonto From Garantias G "
    '    sSql = sSql & " Inner Join Colocgarantia CG ON G.cnumGarant = CG.cNumGarant "
    '    sSql = sSql & " Where CG.cCtaCod = '" & psCtaCod & "' AND G.nEstado = 1"
    'End If
    sSql = "EXEC stp_sel_ERS0632014_RecuperaMontoGarantiaCredito '" & psCtaCod & "'," & IIf(pbConstruccion, 1, 0)
    'END EJVG *******
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    RecuperaMontoGarantiaCredito = IIf(IsNull(R!nMonto), 0, R!nMonto)
    
End Function

'**PEAC 20080111
Public Function RecuperaPolizaParaGenerarGastos(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
   
    sSql = "exec stp_sel_RecuperaPolizaParaGenerarGastos '" & psCtaCod & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaPolizaParaGenerarGastos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
        
End Function
Public Function RecuperaGarantiaCreditoDatosRefinan(ByVal psCtaCods As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaGarantiaCreditoRef
    sSql = "Select Distinct G.cNumgarant, CN.nConsValor as nTpoGar, G.cNroDoc, CN.cConsDescripcion cTpoGarantia, G.cDescripcion, D.cDocDesc, CG.nGravado, CN2.cConsDescripcion cMoneda, G.nPorGravar, G.nGravaMent, CN3.cConsDescripcion cMonedaGar, CG.nMoneda, CG.nEstado"
    sSql = sSql & " From ColocGarantia CG Inner Join Garantias G ON CG.cNumGarant = G.cNumGarant "
    sSql = sSql & "                       Inner Join Constante CN ON G.nTpoGarantia = CN.nConsValor AND CN.nConsCod = " & gPersGarantia
    sSql = sSql & "                       Inner Join Constante CN2 ON CN2.nConsValor = CG.nMoneda AND CN2.nConsCod = " & gMoneda
    sSql = sSql & "                       Inner Join Constante CN3 ON CN3.nConsValor = G.nMoneda AND CN3.nConsCod = " & gMoneda
    sSql = sSql & "                       Inner Join Documento D ON D.nDocTpo = Convert(int,G.cTpoDoc) "
    sSql = sSql & "                       Inner Join GarantDoc O ON D.nDocTpo = O.nDocTpo "
    sSql = sSql & " WHERE CG.cCtaCod in (" & psCtaCods & ")"
    Set oConecta = New COMConecta.DCOMConecta
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
Dim oConecta As COMConecta.DCOMConecta

    'sSql = "Select CN1.cConsDescripcion cEstado, CG.cCtaCod, CN.cConsDescripcion as cMoneda, CG.nGravado, P.nSaldo, "
    'sSql = sSql & " CE.nCuotas , C.nNroProxCuota"
    'sSql = sSql & " From ColocGarantia CG Inner Join Producto P ON CG.cCtaCod = P.cCtaCod"
    'sSql = sSql & " Inner Join Constante CN ON CN.nConsValor = convert(int,SUBSTRING(CG.cCtaCod,9,1)) AND CN.nConsCod = 1011"
    'sSql = sSql & " Left Join ColocacEstado CE ON CG.cCtaCod = CE.cCtaCod AND CE.nPrdEstado = 2002"
    'sSql = sSql & " Inner Join ColocacCred C ON C.cCtaCod = CG.cCtaCod "
    'sSql = sSql & " Inner Join Constante CN1 ON CN1.nConsCod = 3001 AND CN1.nConsValor = P.nPrdEstado "
    'sSql = sSql & " Where cNumGarant = '" & psNumGarant & "' AND nEstado = 3"
    sSql = "EXEC stp_sel_ERS0632014_RecuperaGarantiaCreditoDatosVigente '" & psNumGarant & "'" 'EJVG20150702
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaGarantiaCreditoDatosVigente = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

Public Function RecuperaGarantiaCredito(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaGarantiaCredito
    'sSql = "Select G.cNumGarant cNumGarantP, CN.nConsValor as nTpoGar, G.cNroDoc, CN.cConsDescripcion cTpoGarantia, G.cDescripcion, D.cDocDesc, CG.nGravado, CN2.cConsDescripcion cMoneda, G.nPorGravar, G.nGravaMent, CN3.cConsDescripcion cMonedaGar, G.nRealizacion as nRealizacion"
    'sSql = sSql & " From ColocGarantia CG Inner Join Garantias G ON CG.cNumGarant = G.cNumGarant "
    'sSql = sSql & "                       Inner Join Constante CN ON G.nTpoGarantia = CN.nConsValor AND CN.nConsCod = " & gPersGarantia
    'sSql = sSql & "                       Inner Join Constante CN2 ON CN2.nConsValor = CG.nMoneda AND CN2.nConsCod = " & gMoneda
    'sSql = sSql & "                       Inner Join Constante CN3 ON CN3.nConsValor = G.nMoneda AND CN3.nConsCod = " & gMoneda
    'sSql = sSql & "                       Inner Join Documento D ON D.nDocTpo = Convert(int,G.cTpoDoc) "
    'sSql = sSql & "                       Inner Join GarantDoc O ON D.nDocTpo = O.nDocTpo and O.nConsValor=G.nTpoGarantia"
    'sSql = sSql & " WHERE CG.cCtaCod = '" & psCtaCod & "'"
    sSql = "EXEC stp_sel_ERS0632014_RecuperaGarantiaCredito '" & psCtaCod & "'" 'EJVG20150724
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaGarantiaCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaGarantiaCredito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaGarantiaCreditoDatos(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

    On Error GoTo ErrorRecuperaGarantiaCredito
    sSql = "Select CN.nConsValor as nTpoGar, G.cNroDoc, CN.cConsDescripcion cTpoGarantia, G.cDescripcion, D.cDocDesc, CG.nGravado, CN2.cConsDescripcion cMoneda, G.nPorGravar, G.nGravaMent, CN3.cConsDescripcion cMonedaGar,G.cNumgarant, CG.nMoneda, CG.nEstado"
    sSql = sSql & " From ColocGarantia CG Inner Join Garantias G ON CG.cNumGarant = G.cNumGarant "
    sSql = sSql & "                       Inner Join Constante CN ON G.nTpoGarantia = CN.nConsValor AND CN.nConsCod = " & gPersGarantia
    sSql = sSql & "                       Inner Join Constante CN2 ON CN2.nConsValor = CG.nMoneda AND CN2.nConsCod = " & gMoneda
    sSql = sSql & "                       Inner Join Constante CN3 ON CN3.nConsValor = G.nMoneda AND CN3.nConsCod = " & gMoneda
    sSql = sSql & "                       Inner Join Documento D ON D.nDocTpo = Convert(int,G.cTpoDoc) "
    sSql = sSql & "                       Inner Join GarantDoc O ON D.nDocTpo = O.nDocTpo  and O.nConsValor=G.nTpoGarantia"
    sSql = sSql & " WHERE CG.cCtaCod = '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
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
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaGarantia
    
'*** PEAC 20090724

'    sSql = "Select cTpoDoc, cNroDoc, nTpoGarantia, IdSupGarant,  dGarantia, cDescripcion, cZona, nMoneda, "
'    sSql = sSql & " nTasacion, nRealizacion, nGravament, nPorGravar, nEstado, cFlag, "
'    sSql = sSql & " cComentario, cBancoPersCod, nGarantReal,isnull(nGarantPoliza,0) nGarantPoliza, nGarClase, nGarTpoRealiz, CN.cConsDescripcion cEstado "
'    sSql = sSql & ",cDireccion=ISNULL(cDireccion,'')" 'CUSCO
'    sSql = sSql & " From Garantias G Inner Join Constante CN ON G.nEstado = CN.nConsValor AND CN.nConsCod = 1030"
'    sSql = sSql & " Where cNumGarant = '" & psNumGarant & "' "
    
    sSql = " exec stp_sel_RecuperaGarantia '" & psNumGarant & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
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
Dim oConecta As COMConecta.DCOMConecta

    '*** PEAC 20080523
'    sSql = "Select GR.cNumGarant, isnull(gr.dFecVctoPol,'')dFecVctoPol,isnull(gr.nValorMerca,0)nValorMerca,isnull(gr.cDireAlma,'')cDireAlma,isnull(gr.cTipoMerca,'')cTipoMerca, cPersCodSeguro, dEscritura, cRegistro, nRangHip, "
'    sSql = sSql & " cPersCodVend, cPersVendTelef, cPersCodTasador, nTipVivienda,   "
'    sSql = sSql & " cNroHip, cPersNotaria, PSeg.cPersNombre cSegPersNombre, "
'    sSql = sSql & " PVend.cPersNombre cVendPersNombre,PTasa.cPersNombre cTasaPersNombre, PNota.cPersNombre cNotaPersNombre, "
'    sSql = sSql & " GR.nCuotaInicial, GR.nMontoHipoteca,GR.nValorGravado, nPrecioVenta, nValorConstruccion, nNroPoliza, "
'    sSql = sSql & " dVigenciaPol, nMontoPoliza, dConstitucion, dTasacion, isNull(GR.cdirgarregpubli,'') cdirgarregpubli,GR.CTIPOCONTRATO,nValorEdificacion =ISNULL(nValorEdificacion,0),nVRM = ISNULL(nVRM,0)"
'    '* CUSCO *
'    sSql = sSql & ", cPlacaAuto=ISNULL(cPlacaAuto,''),nEstadoTasacion=ISNULL(nEstadoTasacion,0),dTasacionInmueble=ISNULL(dTasacionInmueble,0),dCertifGravamen=ISNULL(dCertifGravamen,0)"
'    '*********
'    'ARCV 20-01-2007
'    'sSQL = sSQL & " From GarantReal GR Left Join Persona PSeg ON GR.cPersCodSeguro = PSeg.cPersCod "
'    sSql = sSql & " From GarantReal GR INNER JOIN GarantRealTasacion GRT ON GR.cNumGarant = GRT.cNumGarant AND GRT.dTasacion= (SELECT MAX(dTasacion)FROM GarantRealTasacion WHERE cNumGarant='" & psNumGarant & "')"
'    sSql = sSql & "                     Left Join Persona PSeg ON GR.cPersCodSeguro = PSeg.cPersCod "
'    '--------
'    sSql = sSql & "                     Left Join Persona PVend ON GR.cPersCodVend = PVend.cPersCod "
'    sSql = sSql & "                     Left Join Persona PTasa ON cPersCodTasador = PTasa.cPersCod "
'    sSql = sSql & "                     Left Join Persona PNota ON GR.cPersNotaria = PNota.cPersCod "
'    sSql = sSql & " Where GR.cNumGarant = '" & psNumGarant & "' "
    
    sSql = "exec stp_sel_RecuperaGarantiaReal '" & psNumGarant & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaGarantiaReal = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function
'EJVG20130131 ***
Public Function RecuperaGarantPendConstituirySgdoPreferente(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    sSql = "exec stp_sel_GarantiaPendConstituirySgdoPreferente '" & psPersCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaGarantPendConstituirySgdoPreferente = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'END EJVG *******
'peac 20071222
Public Function RecuperaDatosInmueblePoliza(ByVal psNumGarant As String) As ADODB.Recordset
'RecuperaDatosTasador
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    
    oConecta.AbreConexion
    sSql = " Execute stp_sel_ObtieneDatosInmueblePoliza '" & psNumGarant & "'"
    Set RecuperaDatosInmueblePoliza = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

End Function
'*** BRGO 20111205 *******************************************
Public Function RecuperaDatosMueblePoliza(ByVal psNumGarant As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    
    oConecta.AbreConexion
    sSql = " Execute stp_sel_ObtieneDatosMueblePoliza '" & psNumGarant & "'"
    Set RecuperaDatosMueblePoliza = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function RecuperaDatosGarantDocCompra(ByVal psNumGarant As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    
    oConecta.AbreConexion
    sSql = " Execute stp_sel_ObtieneDatosGarantDocCompra '" & psNumGarant & "'"
    Set RecuperaDatosGarantDocCompra = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
'*** END BRGO ************************************************



'*****************************************************************************************
'***     Rutina:           RecuperaRelacPersonaGarantia
'***     Descripcion:       Recupera todas las Personas Relacionadas con la Garantia
'***     Modificado por:        NSSE
'***     Fecha-Tiempo:         11/06/2001 11:20:42 AM
'***     Ultima Modificacion: Creacion
'*****************************************************************************************
Public Function RecuperaRelacPersonaGarantia(ByVal psNumGarant As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaRelacPersonaGarantia
    sSql = "select P.cPersCod, P.cPersNombre, PG.cNumgarant, C.cConsDescripcion + space(50) + convert(varchar(2),PG.nRelacion)  as cRelacion "
    sSql = sSql & " from persgarantia PG Inner Join Persona P ON PG.cPersCod  = P.cPersCod "
    sSql = sSql & " inner join Constante C ON PG.nRelacion = C.nConsValor "
    sSql = sSql & " where C.nConsCod = '" & gPersRelGarantia & "' and PG.cNumGarant = '" & psNumGarant & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
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
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaGarantDeclaracionJur
    sSql = "select DJ.cNumgarant, DJ.nItem, DJ.cGarDJDescripcion, DJ.nGarDJCantidad, DJ.nGarDJPrecioUnit, DJ.cGarDJTpoDoc, "
    sSql = sSql & " DJ.cGarDJNroDoc, C.cConsDescripcion + space(50) + DJ.cGarDJTpoDoc as cGarDJTpoDocDes "
    sSql = sSql & " from GarantDeclaracionJur DJ "
    sSql = sSql & " inner join Constante C ON DJ.cGarDJTpoDoc = C.nConsValor "
    sSql = sSql & " where C.nConsCod = '" & gColocPigTipoDocumento & "' and DJ.cNumGarant = '" & psNumGarant & "' ORDER BY DJ.cNumgarant, DJ.nItem"
    
    Set oConecta = New COMConecta.DCOMConecta
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
Dim SQL As String
Dim Co As COMConecta.DCOMConecta
Dim rs As New ADODB.Recordset
Set Co = New COMConecta.DCOMConecta
SQL = " select * from Garantias " & _
      " where cTpoDoc='" & psTipoDocGar & "'" & _
      " and cNroDoc='" & psNroDoc & "'" & _
      " and cPersCodEmisor='" & psPersCodEmi & "'"
Co.AbreConexion
Set rs = Co.CargaRecordSet(SQL)
Co.CierraConexion
If rs.EOF And rs.BOF Then
    ExisteGarantia = False
Else
    ExisteGarantia = True
End If
Set rs = Nothing
Set Co = Nothing
End Function

'ARCV 11-07-2006: prsTablaValores,pnTipoTabla

Public Sub NuevaGarantia(ByVal psTipoDocGar As String, ByVal psNroDoc As String, ByVal psTipoGar As String, ByVal pnSTipoGarant As Integer, ByVal psMoneda As String, ByVal psDescrip As String, ByVal psZona As String, ByVal pnMontoTasac As Double, ByVal pnMontoReal As Double, _
                ByVal pnMontoxGrav As Double, ByVal psComent As String, RelPers As Variant, ByVal dFecGar As Date, ByVal psPersCodEmi As String, ByVal psBancoPersCod As String, _
                ByVal pnGarClase As Integer, ByVal pnGarTpoRealiz As Integer, ByVal psPersCodSeguro As String, ByVal pdEscritura As Date, _
                ByVal psRegistro As String, ByVal pnRangHip As Integer, ByVal psPersCodVend As String, ByVal psPersVendTelef As String, ByVal psPersCodTasador As String, ByVal pnTipVivienda As Integer, _
                ByVal psNroHip As String, ByVal psPersNotaria As String, GarDetDJ As Variant, Optional ByVal pnCuotaInicial As Double = 0#, _
                Optional ByVal pnMontoHipoteca As Double = 0#, Optional ByVal pnPrecioVenta As Double = 0#, Optional ByVal pnValorConstruccion As Double = 0#, Optional ByVal psNroPoliza As String, _
                Optional ByVal pdVigenciaPol As Date, Optional ByVal pnMontoPoliza As Double, Optional ByVal pdConstitucion As Date, Optional ByVal dTasacion As Date, _
                Optional ByVal psNumGarantia As String = "", Optional ByVal prsDJ As ADODB.Recordset = Nothing, Optional ByVal pdFecSis As Date = 0, Optional ByVal psDireccion As String = "", _
                Optional ByVal pcPlacaAuto As String = "", Optional pnEstadoTasacion As Integer = 0, Optional ByVal pdTasacionInmueble As Date = 0, _
                Optional ByVal psProyectoActual As String = "", Optional ByVal psCodUser As String = "", Optional ByVal psCodCmac As String = "", _
                Optional ByVal psCodAge As String = "", Optional ByVal prsTablaValores As ADODB.Recordset = Nothing, Optional ByVal pnTipoTabla As Integer = 0, Optional psDirRegPubli As String = "", _
                Optional ByVal pPoli As Variant, Optional ByVal pDatosGar As Variant, Optional ByVal pdFecVctoPol As Date = 0, Optional ByVal dFecEmision As Date, Optional ByVal psPersCodEmisor As String = "", _
                Optional ByVal pnTpoDoc As Integer = -1, Optional psNroDocCompra As String = "", Optional ByVal pnValDocCompra As Double = 0)
                'PEAC 20071122 se agrego "Optional ByVal pnVRM As Double = 0"
                'BRGO 20111205 Se agregó parámetros de datos  de Documento de Compra: "dFecEmision,psPersCodEmisor,pnTpoDoc,psNroDoc"

'Public Sub NuevaGarantia(ByVal psTipoDocGar As String, ByVal psNroDoc As String, ByVal psTipoGar As String, ByVal pnSTipoGarant As Integer, ByVal psMoneda As String, ByVal psDescrip As String, ByVal psZona As String, ByVal pnMontoTasac As Double, ByVal pnMontoReal As Double, _
'                ByVal pnMontoxGrav As Double, ByVal psComent As String, RelPers As Variant, ByVal dFecGar As Date, ByVal psPersCodEmi As String, ByVal psBancoPersCod As String, _
'                ByVal pnGarantReal As Integer, ByVal pnGarClase As Integer, ByVal pnGarTpoRealiz As Integer, ByVal psPersCodSeguro As String, ByVal pdEscritura As Date, _
'                ByVal psRegistro As String, ByVal pnRangHip As Integer, ByVal psPersCodVend As String, ByVal psPersVendTelef As String, ByVal psPersCodTasador As String, ByVal pnTipVivienda As Integer, _
'                ByVal psNroHip As String, ByVal psPersNotaria As String, GarDetDJ As Variant, Optional ByVal pnCuotaInicial As Double = 0#, _
'                Optional ByVal pnMontoHipoteca As Double = 0#, Optional ByVal pnPrecioVenta As Double = 0#, Optional ByVal pnValorConstruccion As Double = 0#, Optional ByVal psNroPoliza As String, _
'                Optional ByVal pdVigenciaPol As Date, Optional ByVal pnMontoPoliza As Double, Optional ByVal pdConstitucion As Date, Optional ByVal dTasacion As Date, _
'                Optional ByVal psNumGarantia As String = "", Optional ByVal prsDJ As ADODB.Recordset = Nothing, Optional ByVal pdFecSis As Date = 0, Optional ByVal psDireccion As String = "", _
'                Optional ByVal pcPlacaAuto As String = "", Optional pnEstadoTasacion As Integer = 0, Optional ByVal pdTasacionInmueble As Date = 0, _
'                Optional ByVal psProyectoActual As String = "", Optional ByVal psCodUser As String = "", Optional ByVal psCodCmac As String = "", _
'                Optional ByVal psCodAge As String = "", Optional ByVal prsTablaValores As ADODB.Recordset = Nothing, Optional ByVal pnTipoTabla As Integer = 0, Optional psDirRegPubli As String = "", Optional ByVal psTipoC As String = "", _
'                Optional ByVal pnValorEdificacion As Double = 0, Optional ByVal pnVRM As Double = 0, Optional ByVal pdCertifGravamen As Date = 0, Optional ByVal pnValorGravado As Double = 0#, Optional ByVal pPoli As Variant)
'                'PEAC 20071122 se agrego "Optional ByVal pnVRM As Double = 0"

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim bTran As Boolean
Dim i As Integer
Dim sNumGarant As String
Dim R As ADODB.Recordset

                On Error GoTo ErrorNuevaGarantia
                
                bTran = False
                Set oConecta = New COMConecta.DCOMConecta
                oConecta.AbreConexion
                oConecta.ConexionActiva.BeginTrans
                bTran = True

                sSql = "Select Max(cNumGarant) as cNumGarant From Garantias"
                Set R = oConecta.CargaRecordSet(sSql)
                sNumGarant = Format(CLng(IIf(IsNull(R!cNumGarant), "0", R!cNumGarant)) + 1, "00000000")

                sSql = "INSERT INTO Garantias(cNumGarant,cTpoDoc,cNroDoc,nTpoGarantia,dGarantia,cDescripcion,"
                sSql = sSql & "cZona,nMoneda,nTasacion,nRealizacion,nGravament,nPorGravar,nEstado,cComentario,cPersCodEmisor, "
                sSql = sSql & " cBancoPersCod, nGarantReal,nGarantPoliza, nGarClase, nGarTpoRealiz,IdSupGarant,cDireccion"
                sSql = sSql & ",bTablaValores,nTasador)"   'ARCV 11-07-2006
                sSql = sSql & " VALUES('" & sNumGarant & "','" & psTipoDocGar & "','" & psNroDoc & "','" & psTipoGar & "','"
                sSql = sSql & Format(dFecGar, "mm/dd/yyyy") & "','" & psDescrip & "','" & psZona & "',"
                sSql = sSql & psMoneda & "," & Format(pnMontoTasac, "#0.00") & "," & Format(pnMontoReal, "#0.00") & ","
                sSql = sSql & "0.00," & Format(pnMontoxGrav, "#0.00") & "," & gPersGarantEstadoRegistrado & ",'" & psComent & "','" & psPersCodEmi & "',"
                sSql = sSql & "'" & psBancoPersCod & "'," & pDatosGar(1, 0) & "," & pDatosGar(1, 1) & "," & pnGarClase & "," & pnGarTpoRealiz & "," & pnSTipoGarant & ","
                sSql = sSql & "'" & psDireccion '& "')" 'CUSCO
                sSql = sSql & "'," & IIf(pnTipoTabla > 0, 1, 0) & "," & pDatosGar(1, 14) & " )" 'ARCV 11-07-2006
                '*** PEAC 20090724 - SE AGREGO pDatosGar(1, 14)
                oConecta.ConexionActiva.Execute sSql


'                sSql = "INSERT INTO Garantias(cNumGarant,cTpoDoc,cNroDoc,nTpoGarantia,dGarantia,cDescripcion,"
'                sSql = sSql & "cZona,nMoneda,nTasacion,nRealizacion,nGravament,nPorGravar,nEstado,cComentario,cPersCodEmisor, "
'                sSql = sSql & " cBancoPersCod, nGarantReal, nGarClase, nGarTpoRealiz,IdSupGarant,cDireccion"
'                sSql = sSql & ",bTablaValores,nGarantPoliza)"   'ARCV 11-07-2006
'                sSql = sSql & " VALUES('" & sNumGarant & "','" & psTipoDocGar & "','" & psNroDoc & "','" & psTipoGar & "','"
'                sSql = sSql & Format(dFecGar, "mm/dd/yyyy") & "','" & psDescrip & "','" & psZona & "',"
'                sSql = sSql & psMoneda & "," & Format(pnMontoTasac, "#0.00") & "," & Format(pnMontoReal, "#0.00") & ","
'                sSql = sSql & "0.00," & Format(pnMontoxGrav, "#0.00") & "," & gPersGarantEstadoRegistrado & ",'" & psComent & "','" & psPersCodEmi & "',"
'                sSql = sSql & "'" & psBancoPersCod & "'," & pnGarantReal & "," & pnGarClase & "," & pnGarTpoRealiz & "," & pnSTipoGarant & ","
'                sSql = sSql & "'" & psDireccion '& "')" 'CUSCO
'                sSql = sSql & "'," & IIf(pnTipoTabla > 0, 1, 0) & "," & pnGarantPoliza & " )" 'ARCV 11-07-2006
'                oConecta.ConexionActiva.Execute sSql

                'If pnGarantReal = 1 Then
                
                If pDatosGar(1, 0) = 1 Then
'ARCV 20-01-2007
'                    sSQL = "INSERT INTO GarantReal(cNumGarant, cPersCodSeguro, dEscritura, "
'                    sSQL = sSQL & " cRegistro, nRangHip, cPersCodVend, cPersVendTelef, cPersCodTasador, "
'                    sSQL = sSQL & " nTipVivienda, cNroHip, cPersNotaria,nCuotaInicial,nMontoHipoteca,nPrecioVenta,nValorConstruccion,nNroPoliza, "
'                    sSQL = sSQL & " dVigenciaPol,nMontoPoliza,dConstitucion,dTasacion ,cPlacaAuto,nEstadoTasacion,dTasacionInmueble,cdirgarregpubli,CTIPOCONTRATO)"
'                    sSQL = sSQL & " VALUES('" & sNumGarant & "','" & psPersCodSeguro & "','" & Format(pdEscritura, "mm/dd/yyyy") & "','"
'                    sSQL = sSQL & Trim(psRegistro) & "'," & pnRangHip & ",'" & psPersCodVend & "','" & psPersVendTelef & "','" & psPersCodTasador & "',"
'                    sSQL = sSQL & pnTipVivienda & ",'" & psNroHip & "','" & psPersNotaria & "',"
'                    sSQL = sSQL & Format(pnCuotaInicial, "#0.00") & "," & Format(pnMontoHipoteca, "#0.00") & "," & Format(pnPrecioVenta, "#0.00") & "," & Format(pnValorConstruccion, "#0.00")
'                    sSQL = sSQL & ",'" & psNroPoliza & "','" & Format(pdVigenciaPol, "mm/dd/yyyy") & "'," & Format(pnMontoPoliza, "#0.00")
'                    sSQL = sSQL & ",'" & Format(pdConstitucion, "mm/dd/yyyy") & "','" & Format(dTasacion, "mm/dd/yyyy") & "'"
'****************
                                        
                    '*** PEAC 20080522 "dFecVctoPol"
                    'EJVG20130129 Se agregó nTipoInscripcion
                    sSql = "INSERT INTO GarantReal(cNumGarant,cDescripBien,cNumMotor,cNumSerie,dAnioFab,nValorMerca,cDireAlma,cTipoMerca,dFecVctoPol, cPersCodSeguro, dEscritura, "
                    'sSql = "INSERT INTO GarantReal(cNumGarant,nValorMerca,cDireAlma,cTipoMerca,dFecVctoPol, cPersCodSeguro, dEscritura, "
                    sSql = sSql & " cRegistro, nRangHip, cPersCodVend, cPersVendTelef, "
                    sSql = sSql & " nTipVivienda, cNroHip, cPersNotaria,nCuotaInicial,nMontoHipoteca,nNroPoliza, "
                    sSql = sSql & " dVigenciaPol,nMontoPoliza,dConstitucion,cPlacaAuto,nEstadoTasacion,dTasacionInmueble,cdirgarregpubli,CTIPOCONTRATO,dCertifGravamen,nValorGravado,dFechaBloqueo,nTipoInscripcion)"
                    sSql = sSql & " VALUES('" & sNumGarant & "','" & Trim(pDatosGar(1, 10)) & "','" & Trim(pDatosGar(1, 11)) & "','" & Trim(pDatosGar(1, 12)) & "','" & Format(pDatosGar(1, 13), "yyyymmdd") & "'," & Format(pDatosGar(1, 7), "#0.00") & ",'" & Trim(pDatosGar(1, 8)) & "','" & Trim(pDatosGar(1, 9)) & "','" & Format(pdFecVctoPol, "yyyymmdd") & "','" & psPersCodSeguro & "','" & Format(pdEscritura, "mm/dd/yyyy") & "','"
                    'sSql = sSql & " VALUES('" & sNumGarant & "'," & Format(pDatosGar(1, 7), "#0.00") & ",'" & Trim(pDatosGar(1, 8)) & "','" & Trim(pDatosGar(1, 9)) & "','" & Format(pdFecVctoPol, "yyyymmdd") & "','" & psPersCodSeguro & "','" & Format(pdEscritura, "mm/dd/yyyy") & "','"
                    sSql = sSql & Trim(psRegistro) & "'," & pnRangHip & ",'" & psPersCodVend & "','" & psPersVendTelef & "', "
                    sSql = sSql & pnTipVivienda & ",'" & psNroHip & "','" & psPersNotaria & "',"
                    sSql = sSql & Format(pnCuotaInicial, "#0.00") & "," & Format(pnMontoHipoteca, "#0.00")
                    sSql = sSql & " ,'" & psNroPoliza & "','" & Format(pdVigenciaPol, "mm/dd/yyyy") & "'," & Format(pnMontoPoliza, "#0.00")
                    sSql = sSql & ",'" & Format(pdConstitucion, "mm/dd/yyyy") & "'"
                    'CUSCO
                    sSql = sSql & ",'" & pcPlacaAuto & "'," & pnEstadoTasacion & ",'" & Format(pdTasacionInmueble, "mm/dd/yyyy") & "',"
                    'Maynas
                    sSql = sSql & "'" & psDirRegPubli & "','" & pDatosGar(1, 6) & "','" & Format(pDatosGar(1, 3), "mm/dd/yyyy") & "'," & Format(pDatosGar(1, 2), "#0.00") & ",'" & Format(pDatosGar(1, 16), "mm/dd/yyyy") & "'," & IIf(pDatosGar(1, 17) = -1, "NULL", pDatosGar(1, 17)) & ")"
                    
                    oConecta.ConexionActiva.Execute sSql
                    
                    '*** PEAC 20090724
                    If pDatosGar(1, 14) = 1 Then
                        'ARCV 20-01-2007
                        If psProyectoActual = "M" Then
                            'peac 20071123 se agrego "nVRM"
                            sSql = "INSERT INTO GarantRealTasacion (cNumGarant,dTasacion,cPersCodTasador,nPrecioVenta,nValorConstruccion,nValorEdificacion,nVRM)" & _
                               " VALUES ('" & sNumGarant & "','" & Format(dTasacion, "mm/dd/yyyy") & "','" & psPersCodTasador & "'," & pnPrecioVenta & "," & pnValorConstruccion & "," & pDatosGar(1, 5) & "," & pDatosGar(1, 4) & ")"
                        
                            oConecta.ConexionActiva.Execute sSql
                        End If
                    End If
                    '*** FIN PEAC
                    
                    '*** BRGO 20111205
                    If pDatosGar(1, 15) = 1 Then
                        If psProyectoActual = "M" Then
                            sSql = "INSERT INTO GarantRealDocumento (cNumGarant,dFecEmision,cPersCodEmisor,nTpoDoc,cNroDoc,nValor)" & _
                               " VALUES ('" & sNumGarant & "','" & Format(dFecEmision, "mm/dd/yyyy") & "','" & psPersCodEmisor & "'," & pnTpoDoc & "," & psNroDocCompra & "," & pnValDocCompra & ")"
                            oConecta.ConexionActiva.Execute sSql
                            
                            'Se agrega el registro en GarantRealTasacion para mantener la integridad de datos
                            sSql = "INSERT INTO GarantRealTasacion (cNumGarant,dTasacion,cPerscodTasador,nPrecioVenta,nValorConstruccion, nValorEdificacion)" & _
                            " VALUES ('" & sNumGarant & "','" & Format(dFecEmision, "mm/dd/yyyy") & "','" & psPersCodEmisor & "',0.00,0.00,0.00)"
                            oConecta.ConexionActiva.Execute sSql
                        End If
                    End If
                    '*** FIN BRGO
                End If
                
                'peac 20071221 - ***************************************
                If pPoli(1, 0) = 1 Then
                    'ALPA20140203 se agregó los campo nPolizaBF,dFechaPBF
                    sSql = " insert into GarantInmueble"
                    sSql = sSql & " (cNumGarant,cAgeCod,nInmueble,nCategoria,nNumLocales,nNumPisos,nNumSotanos,nAnioConstruc,nPolizaBF,dFechaPBF)"
                    sSql = sSql & " values ('" & sNumGarant & "','" & psCodAge & "'," & pPoli(1, 1) & "," & pPoli(1, 2) & "," & pPoli(1, 3) & "," & pPoli(1, 4) & "," & pPoli(1, 5) & "," & pPoli(1, 6) & "," & pPoli(1, 7) & "," & IIf(pPoli(1, 7) = 1, "'" & Format(pPoli(1, 8), "YYYY/MM/DD") & "'", "NULL") & ")"
                    oConecta.ConexionActiva.Execute sSql
                End If
'                '**********************************************************
                
                '*** BRGO 20111205 *******************************************
                If pPoli(1, 0) = 2 Then
                    sSql = " insert into GarantMobiliario"
                    sSql = sSql & " (cNumGarant,cAgeCod,nMobiliario,nAnioFabric)"
                    sSql = sSql & " values ('" & sNumGarant & "','" & psCodAge & "'," & pPoli(1, 1) & "," & pPoli(1, 2) & ")"
                    oConecta.ConexionActiva.Execute sSql
                End If
                '*** END BRGO ************************************************
                
                ' CMACICA_CSTS - 25/11/2003 --------------------------------------------------------------------------
                'If psTipoGar = gPersGarantiaOtrasGarantias Then
                '*** PEAC 20080425
                If psTipoDocGar = 93 Then ' Si es declaracion jurada
                'If psTipoDocGar = 15 Then ' Si es declaracion jurada

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
                                                
                If psProyectoActual = "H" Then  'Guardar un historico
                    Dim oCred As COMDCredito.DCOMCredito
                    Set oCred = New COMDCredito.DCOMCredito
                    
                    sSql = "INSERT INTO GarantMontos(cNumGarant,nItem,nTasacion,nRealizacion,nPorGravar,cUltimaActualizacion)" & _
                            " VALUES('" & sNumGarant & "',1," & pnMontoTasac & "," & pnMontoReal & "," & pnMontoxGrav & ",'" & oCred.GeneraMovNroActualiza(pdFecSis, psCodUser, psCodCmac, psCodAge) & "')"
                    
                    Set oCred = Nothing
                    oConecta.ConexionActiva.Execute sSql
                    
                End If
                
                'ARCV 11-07-2006
                'oConecta.ConexionActiva.CommitTrans
                'oConecta.CierraConexion
                'Set oConecta = Nothing
                
                'Para la Impresion en Linea
                'psNumGarantia = ObtenerMaxcNumGarant
                'Set prsDJ = DJ(psNumGarantia, pdFecSis)
                
                'ARCV 11-07-2006
                If Not prsTablaValores Is Nothing Then
                    With prsTablaValores
                        While Not .EOF
                        'CInt(.TextMatrix(.Row, 0))
                        'CInt(Trim(Right(.TextMatrix(.Row, 2), 20)))
                        
                            If prsTablaValores.Fields(2) <> "" And Left(prsTablaValores.Fields(2), 7) <> "NINGUNO" Then
                                sSql = "INSERT INTO Scoring_Garantias" & _
                                        " VALUES(" & pnTipoTabla & "," & CInt(prsTablaValores.Fields(0)) & "," & CInt(Trim(Right(prsTablaValores.Fields(2), 20))) & ",'" & sNumGarant & "')"

                                oConecta.ConexionActiva.Execute sSql
                            End If
                            .MoveNext
                        Wend
                    End With
                End If

                oConecta.ConexionActiva.CommitTrans
                oConecta.CierraConexion
                Set oConecta = Nothing

                'Para la Impresion en Linea
                psNumGarantia = ObtenerMaxcNumGarant
                Set prsDJ = DJ(psNumGarantia, pdFecSis)
                '----------------
                Exit Sub

ErrorNuevaGarantia:
        If bTran Then
            oConecta.ConexionActiva.RollbackTrans
        End If
        Set oConecta = Nothing
        'If Err.Number = -2147217873 Then
        '    MsgBox "Garantia Duplicada"
        '    Exit Sub
        'End If
        'MsgBox Err.Description, vbCritical, "Aviso"
         Err.Raise Err.Number, "Error", Err.Description
End Sub

Public Function GarantiaEnUso(ByVal psNumGarant As String) As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

    sSql = "Select * From ColocGarantia Where cNumGarant  = '" & psNumGarant & "'"
    Set oConecta = New COMConecta.DCOMConecta
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
Dim oConecta As COMConecta.DCOMConecta
Dim bTran As Boolean

    On Error GoTo ErrorEliminarGraantia
    bTran = False
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.BeginTrans
    bTran = True
    'Elimina Relacion de Garantias de Credito
    sSql = "DELETE ColocGarantia WHERE cNumGarant = '" & psNumGarant & "' "
    oConecta.ConexionActiva.Execute sSql
    
'ARCV 23-06-2007
    'Elimina TASACION DE GARANTIA REAL
    'MAVM 20120808 ***
    sSql = "DELETE GarantPoliza WHERE cNumGarant = '" & psNumGarant & "' "
    oConecta.ConexionActiva.Execute sSql
    '***
    
    sSql = "DELETE GarantRealTasacion WHERE cNumGarant = '" & psNumGarant & "' "
    oConecta.ConexionActiva.Execute sSql
    'Elimina MONTOS HISTORICOS
    sSql = "DELETE GarantMontos WHERE cNumGarant = '" & psNumGarant & "' "
    oConecta.ConexionActiva.Execute sSql
    
'---------

    'Elimina garantia Real
    sSql = "DELETE GarantReal WHERE cNumGarant = '" & psNumGarant & "' "
    oConecta.ConexionActiva.Execute sSql
       
    'Elimina garantia poliza
    sSql = "DELETE GarantInmueble WHERE cNumGarant = '" & psNumGarant & "' "
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

'ARCV 11-07-2006

Public Sub ActualizaGarantia(ByVal psNumGarant As String, ByVal psTipoDocGar As String, ByVal psNroDoc As String, _
                ByVal psTipoGar As String, ByVal pnSTipoGarant As Integer, ByVal psMoneda As String, ByVal psDescrip As String, _
                ByVal psZona As String, ByVal pnMontoTasac As Double, ByVal pnMontoReal As Double, ByVal pnMontoxGrav As Double, ByVal psComent As String, RelPers As Variant, _
                ByVal dFecGar As Date, ByVal sPersCodEmi As String, ByVal psBancoPersCod As String, ByVal pnGarantReal As Integer, ByVal pnGarClase As Integer, ByVal pnGarTpoRealiz As Integer, _
                ByVal psPersCodSeguro As String, ByVal pdEscritura As Date, ByVal psRegistro As String, ByVal pnRangHip As Integer, ByVal psPersCodVend As String, _
                ByVal psPersVendTelef As String, ByVal psPersCodTasador As String, ByVal pnTipVivienda As Integer, ByVal psNroHip As String, ByVal psPersNotaria As String, GarDetDJ As Variant, Optional ByVal pnCuotaInicial As Double = 0#, _
                Optional ByVal pnMontoHipoteca As Double = 0#, Optional ByVal pnPrecioVenta As Double = 0#, Optional ByVal pnValorConstruccion As Double = 0#, Optional ByVal psNroPoliza As String, _
                Optional ByVal pdVigenciaPol As Date, Optional ByVal pnMontoPoliza As Double, Optional ByVal pdConstitucion As Date, Optional ByVal dTasacion As Date, _
                Optional ByVal prsDJ As ADODB.Recordset = Nothing, Optional ByVal pdFecSis As Date = 0, Optional ByVal psDireccion As String = "", Optional ByVal pcPlacaAuto As String = "", Optional pnEstadoTasacion As Integer = 0, Optional ByVal pdTasacionInmueble As Date = 0, _
                Optional ByVal psProyectoActual As String = "", Optional ByVal psCodUser As String = "", Optional ByVal psCodCmac As String = "", Optional ByVal psCodAge As String = "", Optional ByRef pbVerificaDescobertura As Boolean = False, Optional ByVal prsTablaValores As ADODB.Recordset = Nothing, _
                Optional ByVal pnTipoTabla As Integer = 0, Optional ByVal psDirecRegPubli As String = "", Optional ByVal pPoli As Variant, Optional ByVal pDatosGar As Variant, Optional ByVal pdFecVctoPol As Date = 0, Optional pnNroPolizaNew As String = "", Optional pDatosGarMob As Variant)
                'peac 20071123 se agrego "Optional ByVal pnVRM As Double = 0"
'Public Sub ActualizaGarantia(ByVal psNumGarant As String, ByVal psTipoDocGar As String, ByVal psNroDoc As String, _
'                ByVal psTipoGar As String, ByVal pnSTipoGarant As Integer, ByVal psMoneda As String, ByVal psDescrip As String, _
'                ByVal psZona As String, ByVal pnMontoTasac As Double, ByVal pnMontoReal As Double, ByVal pnMontoxGrav As Double, ByVal psComent As String, RelPers As Variant, _
'                ByVal dFecGar As Date, ByVal sPersCodEmi As String, ByVal psBancoPersCod As String, ByVal pnGarantReal As Integer, ByVal pnGarClase As Integer, ByVal pnGarTpoRealiz As Integer, _
'                ByVal psPersCodSeguro As String, ByVal pdEscritura As Date, ByVal psRegistro As String, ByVal pnRangHip As Integer, ByVal psPersCodVend As String, _
'                ByVal psPersVendTelef As String, ByVal psPersCodTasador As String, ByVal pnTipVivienda As Integer, ByVal psNroHip As String, ByVal psPersNotaria As String, GarDetDJ As Variant, Optional ByVal pnCuotaInicial As Double = 0#, _
'                Optional ByVal pnMontoHipoteca As Double = 0#, Optional ByVal pnPrecioVenta As Double = 0#, Optional ByVal pnValorConstruccion As Double = 0#, Optional ByVal psNroPoliza As String, _
'                Optional ByVal pdVigenciaPol As Date, Optional ByVal pnMontoPoliza As Double, Optional ByVal pdConstitucion As Date, Optional ByVal dTasacion As Date, _
'                Optional ByVal prsDJ As ADODB.Recordset = Nothing, Optional ByVal pdFecSis As Date = 0, Optional ByVal psDireccion As String = "", Optional ByVal pcPlacaAuto As String = "", Optional pnEstadoTasacion As Integer = 0, Optional ByVal pdTasacionInmueble As Date = 0, _
'                Optional ByVal psProyectoActual As String = "", Optional ByVal psCodUser As String = "", Optional ByVal psCodCmac As String = "", Optional ByVal psCodAge As String = "", Optional ByRef pbVerificaDescobertura As Boolean = False, Optional ByVal prsTablaValores As ADODB.Recordset = Nothing, _
'                Optional ByVal pnTipoTabla As Integer = 0, Optional ByVal psDirecRegPubli As String = "", Optional ByVal psTipo As String = "", Optional ByVal pnValorEdificacion As Double = 0, Optional ByVal pnVRM As Double = 0, Optional ByVal pdCertifGravamen As Date = 0, Optional ByVal pnValorGravado As Double = 0#, Optional ByVal pPoli As Variant)
'                'peac 20071123 se agrego "Optional ByVal pnVRM As Double = 0"

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim bTran As Boolean
Dim i As Integer

            'On Error GoTo ErrorActualizaGarantia
            
            bTran = False
            Set oConecta = New COMConecta.DCOMConecta
            oConecta.AbreConexion
                                   
            '*********************************************
            If psProyectoActual = "H" Or psProyectoActual = "M" Then   'Verifica descobertura
                Dim rs As ADODB.Recordset
                Dim r1 As ADODB.Recordset
                'ARCV 05-07-2007
                'sSql = "SELECT cNumGarant,Total=SUM(nGravado)FROM ColocGarantia WHERE cNumGarant='" & psNumGarant & "'" & _
                       " GROUP BY cNumGarant "
                
'                sSql = "SELECT cNumGarant,Total=SUM(nGravado)FROM ColocGarantia WHERE cNumGarant='" & psNumGarant & "'" & _
'                       " AND nEstado=1 GROUP BY cNumGarant "
                
'                'MADM 20110926
                sSql = " stp_sel_ValorTotalGravadoGarantia '" & psNumGarant & "'"
'                 sSql = "SELECT cNumGarant,Total=SUM(nGravado)FROM ColocGarantia CG Inner Join Producto P on CG.cCtaCod = P.cCtaCod  WHERE cNumGarant='" & psNumGarant & "'" & _
'                       " AND nEstado=1 AND nPrdEstado Not in (2050,2060) GROUP BY cNumGarant "
                'end madm
                
                '------
                Set rs = oConecta.CargaRecordSet(sSql)
                If rs.EOF Then
                    pbVerificaDescobertura = False
                Else
                    If pnMontoxGrav <= rs!Total Then
                        pbVerificaDescobertura = True
                        Exit Sub
                    Else
                        pbVerificaDescobertura = False
                    End If
                End If
            End If
            '*********************************************
                        
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
            sSql = sSql & " nGarantReal = " & pDatosGar(1, 0) & ","
            sSql = sSql & " nGarantPoliza = " & pDatosGar(1, 1) & ","
            sSql = sSql & " nGarClase = " & pnGarClase & ","
            sSql = sSql & " nGarTpoRealiz = " & pnGarTpoRealiz & ", "
            sSql = sSql & " cPersCodEmisor = '" & sPersCodEmi & "', "
            sSql = sSql & "IdSupGarant=" & pnSTipoGarant
            sSql = sSql & ",cDireccion='" & psDireccion '& "'" 'CUSCO
            sSql = sSql & "',bTablaValores=" & IIf(pnTipoTabla > 0, 1, 0) 'ARCV 11-07-2006
            sSql = sSql & ",nTasador = " & pDatosGar(1, 14) '*** PEAC 20090724
            sSql = sSql & " WHERE cNumGarant = '" & psNumGarant & "' "
            
'            sSql = "UPDATE Garantias SET "
'            sSql = sSql & " nTpoGarantia = " & psTipoGar & ", "
'            sSql = sSql & " cTpoDoc = " & psTipoDocGar & ", "
'            sSql = sSql & " cNroDoc = '" & psNroDoc & "', "
'            sSql = sSql & " dGarantia = '" & Format(dFecGar, "mm/dd/yyyy") & "', "
'            sSql = sSql & " cDescripcion = '" & psDescrip & "',"
'            sSql = sSql & " cZona = '" & psZona & "', "
'            sSql = sSql & " nMoneda = " & psMoneda & ", "
'            sSql = sSql & " nTasacion = " & Format(pnMontoTasac, "#0.00") & ", "
'            sSql = sSql & " nRealizacion = " & Format(pnMontoReal, "#0.00") & ", "
'            'sSql = sSql & " nGravament = " & Format(pnMontoxGrav, "#0.00") & ", "
'            sSql = sSql & " nPorGravar = " & Format(pnMontoxGrav, "#0.00") & ", "
'            sSql = sSql & " cComentario = ' " & psComent & "', "
'            sSql = sSql & " cBancoPersCod = '" & psBancoPersCod & "',"
'            sSql = sSql & " nGarantReal = " & pnGarantReal & ","
'            sSql = sSql & " nGarantPoliza = " & pnGarantPoliza & ","
'            sSql = sSql & " nGarClase = " & pnGarClase & ","
'            sSql = sSql & " nGarTpoRealiz = " & pnGarTpoRealiz & ", "
'            sSql = sSql & " cPersCodEmisor = '" & sPersCodEmi & "', "
'            sSql = sSql & "IdSupGarant=" & pnSTipoGarant
'            sSql = sSql & ",cDireccion='" & psDireccion '& "'" 'CUSCO
'            sSql = sSql & "',bTablaValores=" & IIf(pnTipoTabla > 0, 1, 0) 'ARCV 11-07-2006
'            sSql = sSql & " WHERE cNumGarant = '" & psNumGarant & "' "
            oConecta.ConexionActiva.Execute sSql
            '*** PEAC 20080522 "dFecVctoPol"
            'If pnGarantReal = 1 Then
            If pDatosGar(1, 0) = 1 Then
                sSql = "UPDATE GarantReal SET "
                sSql = sSql & " cPersCodSeguro = '" & psPersCodSeguro & "',"
                
                sSql = sSql & " cDescripBien = '" & Trim(pDatosGar(1, 10)) & "', "
                sSql = sSql & " cNumMotor = '" & Trim(pDatosGar(1, 11)) & "', "
                sSql = sSql & " cNumSerie = '" & Trim(pDatosGar(1, 12)) & "', "
                sSql = sSql & " dAnioFab = '" & Format(pDatosGar(1, 13), "yyyymmdd") & "',"
                
                sSql = sSql & " dEscritura = '" & Format(pdEscritura, "mm/dd/yyyy") & "',"
                sSql = sSql & " dFecVctoPol = '" & Format(pdFecVctoPol, "yyyymmdd") & "',"
                sSql = sSql & " cRegistro = '" & Trim(psRegistro) & "', "
                sSql = sSql & " nRangHip = " & pnRangHip & ", "
                sSql = sSql & " cPersCodVend = '" & psPersCodVend & "', "
                sSql = sSql & " cPersVendTelef = '" & psPersVendTelef & "', "
                'sSQL = sSQL & " cPersCodTasador = '" & psPersCodTasador & "', "
                sSql = sSql & " nTipVivienda = " & pnTipVivienda & ", "
                sSql = sSql & " cNroHip = '" & psNroHip & "', "
                sSql = sSql & " cPersNotaria = '" & psPersNotaria & "', "
                sSql = sSql & " nCuotaInicial = " & Format(pnCuotaInicial, "#0.00") & ","
                sSql = sSql & " nMontoHipoteca = " & Format(pnMontoHipoteca, "#0.00") & ","
                '*** PEAC 20080523
                sSql = sSql & " nValorMerca = " & Format(pDatosGar(1, 7), "#0.00") & ","
                sSql = sSql & " cDireAlma = '" & Trim(pDatosGar(1, 8)) & "', "
                sSql = sSql & " cTipoMerca = '" & Trim(pDatosGar(1, 9)) & "', "

                'sSQL = sSQL & " nPrecioVenta = " & Format(pnPrecioVenta, "#0.00") & ","
                'sSQL = sSQL & " nValorConstruccion = " & Format(pnValorConstruccion, "#0.00") & ","
                sSql = sSql & " nNroPoliza = '" & psNroPoliza & "',"
                sSql = sSql & " dVigenciaPol = '" & Format(pdVigenciaPol, "mm/dd/yyyy") & "',"
                sSql = sSql & " nMontoPoliza = " & Format(pnMontoPoliza, "#0.00") & ","
                'ALPA 20120322***********************
                sSql = sSql & " dFechaBloqueo = '" & Format(pDatosGar(1, 16), "mm/dd/yyyy") & "',"
                '************************************
                sSql = sSql & " dConstitucion = '" & Format(pdConstitucion, "mm/dd/yyyy") & "'"
                'sSQL = sSQL & " dTasacion = '" & Format(dTasacion, "mm/dd/yyyy") & "'"
                'CUSCO
                'sSql = sSql & ", cPlacaAuto='" & pcPlacaAuto & "',nEstadoTasacion=" & pnEstadoTasacion & ",dTasacionInmueble='" & Format(pdTasacionInmueble, "mm/dd/yyyy") & "',dCertifGravamen='" & Format(pdCertifGravamen, "mm/dd/yyyy") & "',nValorGravado = " & Format(pnValorGravado, "#0.00") & " "
                sSql = sSql & ", cPlacaAuto='" & pcPlacaAuto & "',nEstadoTasacion=" & pnEstadoTasacion & ",dTasacionInmueble='" & Format(pdTasacionInmueble, "mm/dd/yyyy") & "',dCertifGravamen='" & Format(pDatosGar(1, 3), "mm/dd/yyyy") & "',nValorGravado = " & Format(pDatosGar(1, 2), "#0.00") & " "
                
                '***********
                sSql = sSql & " ,cdirgarregpubli ='" & psDirecRegPubli & "',"
                'sSql = sSql & " CTIPOCONTRATO ='" & psTipo & "'"
                sSql = sSql & " CTIPOCONTRATO ='" & pDatosGar(1, 6) & "'"
                sSql = sSql & ",nTipoInscripcion = " & IIf(pDatosGar(1, 17) = -1, "NULL", pDatosGar(1, 17))  'EJVG20130131
                sSql = sSql & " WHERE cNumGarant = '" & psNumGarant & "' "
                
                oConecta.ConexionActiva.Execute sSql
                
                '*** BRGO 20120222
                If pDatosGar(1, 15) = 1 Then
                    If psProyectoActual = "M" Then
                        sSql = "UPDATE GarantRealDocumento SET "
                        sSql = sSql & " dFecEmision = '" & Format(pDatosGarMob(0), "mm/dd/yyyy") & "', "
                        sSql = sSql & " cPersCodEmisor = '" & pDatosGarMob(1) & "', "
                        sSql = sSql & " nTpoDoc = " & pDatosGarMob(2) & ", "
                        sSql = sSql & " cNroDoc='" & pDatosGarMob(3) & "', "
                        sSql = sSql & " nValor = " & pDatosGarMob(4)
                        sSql = sSql & " WHERE cNumGarant = '" & psNumGarant & "'"
                        oConecta.ConexionActiva.Execute sSql
                            
                        'Se actualiza el registro en GarantRealTasacion para mantener la integridad de datos
                        
                        sSql = "UPDATE GarantRealTasacion " & _
                            " SET dTasacion = '" & Format(pDatosGarMob(0), "mm/dd/yyyy") & "', cPerscodTasador = '" & pDatosGarMob(1) & "' " & _
                            " WHERE cNumGarant = '" & psNumGarant & "'"
                        oConecta.ConexionActiva.Execute sSql
                    End If
                End If
                '*** FIN BRGO
                
                
                'peac 20080111 - **************** modifica los datos del inmueble-poliza
                
                If pPoli(1, 0) = 1 Then
                    Set r1 = oConecta.CargaRecordSet("select * from GarantInmueble where cNumGarant='" & psNumGarant & "'")
                    If r1.RecordCount = 0 Then
                        sSql = " insert into GarantInmueble"
                        sSql = sSql & " (cNumGarant,cAgeCod,nInmueble,nCategoria,nNumLocales,nNumPisos,nNumSotanos,nAnioConstruc)"
                        sSql = sSql & " values ('" & psNumGarant & "','" & psCodAge & "'," & pPoli(1, 1) & "," & pPoli(1, 2) & "," & pPoli(1, 3) & "," & pPoli(1, 4) & "," & pPoli(1, 5) & "," & pPoli(1, 6) & ")"
                    Else
                        sSql = " update GarantInmueble set "
                        sSql = sSql & " cAgeCod      ='" & psCodAge & "', "
                        sSql = sSql & " nInmueble    =" & pPoli(1, 1) & ", "
                        sSql = sSql & " nCategoria   =" & pPoli(1, 2) & ", "
                        sSql = sSql & " nNumLocales  =" & pPoli(1, 3) & ", "
                        sSql = sSql & " nNumPisos    =" & pPoli(1, 4) & ", "
                        sSql = sSql & " nNumSotanos  =" & pPoli(1, 5) & ", "
                        sSql = sSql & " nAnioConstruc=" & pPoli(1, 6) & ", "
                        'ALPA20140204*******************************************************************
                        sSql = sSql & " nPolizaBF=" & pPoli(1, 7) & ", "
                        sSql = sSql & " dFechaPBF=" & IIf(pPoli(1, 7) = 1, "'" & Format(pPoli(1, 8), "YYYY/MM/DD") & "'", "NULL") & " "
                        sSql = sSql & " Where cNumGarant = '" & psNumGarant & "' "
                        '*******************************************************************************
                     End If
                    oConecta.ConexionActiva.Execute sSql
                End If
                '**********************************************************
                
                '*** BRGO 20120222 *******************************************
                If pPoli(1, 0) = 2 Then
                    Set r1 = oConecta.CargaRecordSet("select * from GarantMobiliario where cNumGarant='" & psNumGarant & "'")
                    If r1.RecordCount = 0 Then
                        sSql = " insert into GarantMobiliario"
                        sSql = sSql & " (cNumGarant,cAgeCod,nMobiliario,nAnioFabric)"
                        sSql = sSql & " values ('" & psNumGarant & "','" & psCodAge & "'," & pPoli(1, 1) & "," & pPoli(1, 2) & ")"
                   Else
                        sSql = " update GarantMobiliario set "
                        sSql = sSql & " cAgeCod      ='" & psCodAge & "', "
                        sSql = sSql & " nMobiliario  =" & pPoli(1, 1) & ", "
                        sSql = sSql & " nAnioFabric  =" & pPoli(1, 2) & " "
                        sSql = sSql & " Where cNumGarant = '" & psNumGarant & "' "
                   End If
                    oConecta.ConexionActiva.Execute sSql
                End If
                '*** END BRGO ************************************************
                
            End If
            'pcPlacaAuto pnEstadoTasacion pdTasacionInmueble
            
            
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
            '*** PEAC 0080412
            'If psTipoDocGar = 15 Then ' declaracion jurada
            If psTipoDocGar = 93 Then ' declaracion jurada
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
            
            '***********************************************************
            If psProyectoActual = "H" Or psProyectoActual = "M" Then   'Guardar un historico
                Dim oCred As COMDCredito.DCOMCredito
                Set oCred = New COMDCredito.DCOMCredito
                    
                sSql = "INSERT INTO GarantMontos(cNumGarant,nItem,nTasacion,nRealizacion,nPorGravar,cUltimaActualizacion)" & _
                       " SELECT '" & psNumGarant & "',ISNULL(MAX(nItem),0)+1," & pnMontoTasac & "," & pnMontoReal & "," & pnMontoxGrav & ",'" & oCred.GeneraMovNroActualiza(pdFecSis, psCodUser, psCodCmac, psCodAge) & "'" & _
                       " FROM GarantMontos WHERE cNumGarant='" & psNumGarant & "'"
                    
                Set oCred = Nothing
                oConecta.ConexionActiva.Execute sSql
                    
            End If
            '***********************************************************
            
            'ARCV 20-01-2007
            'If pnGarantReal = 1 Then
            If pDatosGar(1, 0) = 1 Then
                If psProyectoActual = "M" Then
                    Set rs = oConecta.CargaRecordSet("SELECT*FROM GarantReal WHERE cNumGarant='" & psNumGarant & "'")
                    If rs.EOF Then
                        sSql = "INSERT INTO GarantReal(cNumGarant,cDescripBien,cNumMotor,cNumSerie,dAnioFab,nValorMerca,cDireAlma,cTipoMerca,dFecVctoPol, cPersCodSeguro, dEscritura, "
                        'sSql = "INSERT INTO GarantReal(cNumGarant,nValorMerca,cDireAlma,cTipoMerca,dFecVctoPol, cPersCodSeguro, dEscritura, "
                        sSql = sSql & " cRegistro, nRangHip, cPersCodVend, cPersVendTelef, "
                        sSql = sSql & " nTipVivienda, cNroHip, cPersNotaria,nCuotaInicial,nMontoHipoteca,nNroPoliza, "
                        sSql = sSql & " dVigenciaPol,nMontoPoliza,dConstitucion,cPlacaAuto,nEstadoTasacion,dTasacionInmueble,cdirgarregpubli,CTIPOCONTRATO,dCertifGravamen,nValorGravado)"
                        sSql = sSql & " VALUES('" & psNumGarant & "','" & Trim(pDatosGar(1, 10)) & "','" & Trim(pDatosGar(1, 11)) & "','" & Trim(pDatosGar(1, 12)) & "','" & Format(pDatosGar(1, 13), "yyyymmdd") & "'," & Format(pDatosGar(1, 7), "#0.00") & ",'" & Trim(pDatosGar(1, 8)) & "','" & Trim(pDatosGar(1, 9)) & "','" & Format(pdFecVctoPol, "yyyymmdd") & "','" & psPersCodSeguro & "','" & Format(pdEscritura, "mm/dd/yyyy") & "','"
                        'sSql = sSql & " VALUES('" & psNumGarant & "'," & Format(pDatosGar(1, 7), "#0.00") & ",'" & Trim(pDatosGar(1, 8)) & "','" & Trim(pDatosGar(1, 9)) & "','" & Format(pdFecVctoPol, "yyyymmdd") & "','" & psPersCodSeguro & "','" & Format(pdEscritura, "mm/dd/yyyy") & "','"
                        sSql = sSql & Trim(psRegistro) & "'," & pnRangHip & ",'" & psPersCodVend & "','" & psPersVendTelef & "', "
                        sSql = sSql & pnTipVivienda & ",'" & psNroHip & "','" & psPersNotaria & "',"
                        sSql = sSql & Format(pnCuotaInicial, "#0.00") & "," & Format(pnMontoHipoteca, "#0.00")
                        sSql = sSql & " ,'" & psNroPoliza & "','" & Format(pdVigenciaPol, "mm/dd/yyyy") & "'," & Format(pnMontoPoliza, "#0.00")
                        sSql = sSql & ",'" & Format(pdConstitucion, "mm/dd/yyyy") & "'"
                        'CUSCO
                        sSql = sSql & ",'" & pcPlacaAuto & "'," & pnEstadoTasacion & ",'" & Format(pdTasacionInmueble, "mm/dd/yyyy") & "',"
                        'Maynas
                        'sSql = sSql & "'" & psDirecRegPubli & "','" & psTipoGar & "','" & Format(pdCertifGravamen, "mm/dd/yyyy") & "'," & Format(pnValorGravado, "#0.00") & ")"
                        'ALPA 20091216****************************
                        'sSql = sSql & "'" & psDirecRegPubli & "','" & psTipoGar & "','" & Format(pDatosGar(1, 3), "mm/dd/yyyy") & "'," & Format(pDatosGar(1, 2), "#0.00") & ")"
                        sSql = sSql & "'" & psDirecRegPubli & "','" & pDatosGar(1, 6) & "','" & Format(pDatosGar(1, 3), "mm/dd/yyyy") & "'," & Format(pDatosGar(1, 2), "#0.00") & ")"
                        '*****************************************
                        
                        '
                        oConecta.ConexionActiva.Execute sSql
                    End If
                    
                    
                    If pDatosGar(1, 14) = 1 Then
                        Set rs = oConecta.CargaRecordSet("SELECT*FROM GarantRealTasacion WHERE cNumGarant='" & psNumGarant & "' AND dTasacion='" & Format(dTasacion, "mm/dd/yyyy") & "'")
                        'Set rs = oConecta.CargaRecordSet("SELECT*FROM GarantRealTasacion WHERE cNumGarant='" & psNumGarant & "'")
    
                        'rs.MoveFirst 'peac 20071128
                        
                        If (rs.EOF And rs.BOF) Then
                            'rs.RecordCount = 0
                            
                        'If rs.EOF Then  'peac 20071128
                            ' peac 20071123 se agrego "nVRM"
                            'sSql = "INSERT INTO GarantRealTasacion (cNumGarant,dTasacion,cPersCodTasador,nPrecioVenta,nValorConstruccion,nValorEdificacion,nVRM)" & _
                               " VALUES ('" & psNumGarant & "','" & Format(dTasacion, "mm/dd/yyyy") & "','" & psPersCodTasador & "'," & pnPrecioVenta & "," & pnValorConstruccion & "," & pnValorEdificacion & "," & pnVRM & ")"
                        
                            sSql = "INSERT INTO GarantRealTasacion (cNumGarant,dTasacion,cPersCodTasador,nPrecioVenta,nValorConstruccion,nValorEdificacion,nVRM)" & _
                               " VALUES ('" & psNumGarant & "','" & Format(dTasacion, "mm/dd/yyyy") & "','" & psPersCodTasador & "'," & pnPrecioVenta & "," & pnValorConstruccion & "," & pDatosGar(1, 5) & "," & pDatosGar(1, 4) & ")"
                            
                            oConecta.ConexionActiva.Execute sSql
                        Else
                            'peac 20071128
                            Set r1 = oConecta.CargaRecordSet("select cnumgarant,dTasacion from GarantPoliza where cNumGarant='" & psNumGarant & "'AND dTasacion='" & Format(dTasacion, "mm/dd/yyyy") & "'")
                            If (r1.EOF And r1.BOF) Then
                                'r1.RecordCount = 0
                                
                                sSql = " Update GarantRealTasacion set dTasacion='" & Format(dTasacion, "mm/dd/yyyy") & "',cPersCodTasador=" & psPersCodTasador & ","
                                sSql = sSql & " nPrecioVenta = " & pnPrecioVenta & ",nValorConstruccion=" & pnValorConstruccion & " ,nValorEdificacion=" & pDatosGar(1, 5) & ",nVRM=" & pDatosGar(1, 4)
                                sSql = sSql & " where cNumGarant='" & psNumGarant & "' AND dTasacion='" & Format(dTasacion, "mm/dd/yyyy") & "'"
                            Else
                                sSql = " Update GarantRealTasacion set cPersCodTasador=" & psPersCodTasador & ","
                                sSql = sSql & " nPrecioVenta = " & pnPrecioVenta & ",nValorConstruccion=" & pnValorConstruccion & " ,nValorEdificacion=" & pDatosGar(1, 5) & ",nVRM=" & pDatosGar(1, 4)
                                sSql = sSql & " where cNumGarant='" & psNumGarant & "' AND dTasacion='" & Format(dTasacion, "mm/dd/yyyy") & "'"
                            End If
                            oConecta.ConexionActiva.Execute sSql
                        End If
                    End If
                End If
            End If
            '***********************
            
            'ARCV 11-07-2006
            If Not prsTablaValores Is Nothing Then
                sSql = "DELETE FROM Scoring_Garantias WHERE cNumGarant='" & psNumGarant & "'"
                oConecta.ConexionActiva.Execute sSql
                With prsTablaValores
                    While Not .EOF
                        If prsTablaValores.Fields(2) <> "" And Left(prsTablaValores.Fields(2), 7) <> "NINGUNO" Then
                            sSql = "INSERT INTO Scoring_Garantias" & _
                                    " VALUES(" & pnTipoTabla & "," & CInt(prsTablaValores.Fields(0)) & "," & CInt(Trim(Right(prsTablaValores.Fields(2), 20))) & ",'" & psNumGarant & "')"
                                    
                            oConecta.ConexionActiva.Execute sSql
                        End If
                        .MoveNext
                    Wend
                End With
            End If
            '------------------------------
            oConecta.ConexionActiva.CommitTrans
            oConecta.CierraConexion
            
            Set prsDJ = DJ(psNumGarant, pdFecSis)
            
            Set oConecta = Nothing
    Exit Sub

ErrorActualizaGarantia:
        If bTran Then
            oConecta.ConexionActiva.RollbackTrans
        End If
        Set oConecta = Nothing
        Err.Raise Err.Number, "Error En Proceso ActualizaGarantia", Err.Description
End Sub

Public Function ListaGarantias(ByVal psCodPersona As String) As ADODB.Recordset
    'Lista a las garantias donde la persona es el titular y muestra la descripcion
     Dim oConecta As COMConecta.DCOMConecta
     Dim strSQL As String
     On Error GoTo ErrHandler
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
'        strSQL = "select c.cDescripcion,c.cNumGarant "
'        strSQL = strSQL & " From Persona a "
'        strSQL = strSQL & "inner join PersGarantia b on a.cPersCod=b.cPersCod "
'        strSQL = strSQL & " inner join Garantias c on b.cNumGarant=c.cNumGarant "
'        strSQL = strSQL & " where  b.nRelacion=1 and a.cPersCod='" & psCodPersona & "'"
        strSQL = "exec stp_sel_ListarGarantias '" & psCodPersona & "'"
        Set ListaGarantias = oConecta.CargaRecordSet(strSQL)
        
        oConecta.CierraConexion
        Set oConecta = Nothing
    Exit Function
ErrHandler:
    If Not oConecta Is Nothing Then Set oConecta = Nothing
    Set ListaGarantias = Nothing
End Function

'JACA 20110627 SE AGREGO LA VARIABLE cCtaCod
Public Function DescripGarantia(ByVal cNumGarant As String, ByVal cCtaCod As String) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim strSQL As String
    On Error GoTo ErrHandler
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        'JACA 20110627*******************************************************
'        strSQL = "select a.nGravament,a.nTasacion,a.nRealizacion, "
'        strSQL = strSQL & " cEstado=(Select cConsDescripcion From Constante where nconscod='1030' and nConsValor=a.nEstado) , "
'        strSQL = strSQL & " a.nEstado, p.cPersCod, p.cPersNombre  "
'        strSQL = strSQL & " From garantias a "
'        strSQL = strSQL & " Inner join PersGarantia pg on pg.cNumGarant=a.cNumGarant "
'        strSQL = strSQL & " Inner join Persona p on pg.cPersCod=p.cPersCod "
'        strSQL = strSQL & " where a.cNumGarant='" & cNumGarant & "'"
        
        strSQL = "exec stp_sel_DescripGarantia '" & cNumGarant & "','" & cCtaCod & "'"
        'JACA END************************************************************
        
         Set DescripGarantia = oConecta.CargaRecordSet(strSQL)
         oConecta.CierraConexion
         Set oConecta = Nothing
    Exit Function
ErrHandler:
    If Not oConecta Is Nothing Then Set oConecta = Nothing
    Set DescripGarantia = Nothing
End Function

Public Function InsertarGarantiaPreferidas(ByVal pcMovNro As String, ByVal pcNumGarant As String, _
                                           ByVal nMonto As Double, ByVal pcPersCod As String, _
                                           ByVal nEstado As Integer, Optional ByVal psCtaCod As String) As Boolean
    Dim oConecta As COMConecta.DCOMConecta
    Dim strSQL As String
    Dim strMovDesc As String
    Dim intMovNro As Long
    Dim rs As ADODB.Recordset
    Dim bTran As Boolean
    On Error GoTo ErrHandler
        bTran = False
        strMovDesc = "Movimiento por Garantias Preferidas"
        Set oConecta = New COMConecta.DCOMConecta
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
            strSQL = strSQL & intMovNro & ",'" & pcNumGarant & "'," & 172700 & "," & nMonto & ", 'Movimiento por Garantia Preferida')"
            oConecta.Ejecutar (strSQL)
            
            'MODIFICADO X JACA 20110627************************************************
            'strSQL = "Insert Into ColocCalifGarantia (cPersCod, cNumGarant, nEstadoGar, nMovNro) Values('" & pcPersCod & "','" & pcNumGarant & "'," & nEstado & "," & intMovNro & ")"
             strSQL = "exec stp_upd_GarantiaPreferida '" & pcNumGarant & "'," & nEstado & ",'" & psCtaCod & "'"
            'JACA END ****************************************************************
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


Public Function ListaSuperGarantias() As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler

    sSql = "Select nConsValor,cConsDescripcion from constante where nconscod='1048' and nConsValor<>'1048'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaSuperGarantias = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ListaSuperGarantias = Null
End Function

'peac 20071116 'para activarlo despues
Public Function ListaClaseInmueble() As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler

    sSql = "Select nConsValor,cConsDescripcion from constante where nconscod='1052' and nConsValor<>'1052'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaClaseInmueble = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ListaClaseInmueble = Null
End Function



'peac 20071116  'para activarlo despues

Public Function ListaCategoria() As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler

    sSql = "Select nConsValor,cConsDescripcion from constante where nconscod='1054' and nConsValor<>'1054'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaCategoria = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ListaCategoria = Null
End Function




Public Function CargarRelGarantia(ByVal pIsSuperTipo As Integer) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrHandler
        sSql = "Select C.nConsValor,C.cConsDescripcion,R.IdInfGarant "
        sSql = sSql & " From relgarant R"
        sSql = sSql & " Inner join Constante C on C.nConsValor=R.IdInfGarant"
        sSql = sSql & " Where IdSupGarant=" & pIsSuperTipo & " and C.nConsCod='1027' ORDER BY cConsDescripcion "
     Set oConec = New COMConecta.DCOMConecta
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
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    
        On Error GoTo ErrHandler
        sSql = "Select IdSupGarant From relgarant Where IdInfGarant=" & pIdTipoGarantia
        
        Set oConec = New COMConecta.DCOMConecta
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


Public Function ValidaPFCTS(ByVal psCtaCod As String) As Boolean
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim nValor As Integer
    
    sSql = "Select Count(*) as nCantidad"
    sSql = sSql & " From Producto"
    sSql = sSql & " Where SubString(cCtaCod,6,3) in ('233','234') and nPrdEstado not in (1300,1400)"
    sSql = sSql & " and cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
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

'******************************************************
'   Funciones de la Clase DMantGarantia
'*****************************************************


Public Function ConsultarTpoGarantia(ByVal pnTpoGarantia As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    On Error GoTo ErrHandler
        sSql = "SELECT * FROM ("
        sSql = sSql & " SELECT 1 AS ASIGNACION,D.nDocTpo,D.cDocDesc"
        sSql = sSql & " FROM DOCUMENTO D"
        sSql = sSql & " INNER JOIN GARANTDOC GD ON D.NDOCTPO=GD.NDOCTPO"
        sSql = sSql & " Where GD.NCONSVALOR = " & pnTpoGarantia
        sSql = sSql & " Union"
        sSql = sSql & " SELECT 0 AS ASIGNACION,D.nDocTpo,D.cDocDesc"
        sSql = sSql & " FROM DOCUMENTO D"
        sSql = sSql & " WHERE NOT EXISTS(SELECT NDOCTPO FROM GARANTDOC "
        sSql = sSql & " WHERE NDOCTPO=D.NDOCTPO AND NCONSVALOR=" & pnTpoGarantia & " ))X"
        sSql = sSql & " ORDER BY ASIGNACION DESC,ndoctpo asc"
        
        Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
        Set ConsultarTpoGarantia = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ConsultarTpoGarantia = Null
End Function

Public Function CargarTpoGarantia() As ADODB.Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta

    sSql = "Select nConsValor,cConsDescripcion From Constante Where nConsCod=1027 and nConsValor<>1027 "
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set CargarTpoGarantia = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    On Error GoTo ErrHandler
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set CargarTpoGarantia = Null
End Function

Public Function VerificaGarantDoc(ByVal pTpoGarantia As Integer, ByVal pDocTpo As Integer) As Boolean
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    On Error GoTo ErrHandler
        sSql = "Select Count(*) as Cantidad"
        sSql = sSql & " From GarantDoc Where nConsValor=" & pTpoGarantia & " and nDocTpo=" & pDocTpo
        
        Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        Set oConec = Nothing
        If Not rs.BOF And Not rs.EOF Then
            If rs!cantidad > 0 Then
                VerificaGarantDoc = True
            Else
                VerificaGarantDoc = False
            End If
        End If
        Set rs = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    VerificaGarantDoc = False
End Function


Public Function ActualizarGarantDoc(ByVal prs As ADODB.Recordset, ByVal nTpoGarant As Integer) As Boolean
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    On Error GoTo ErrHandler
        Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
        oConec.ConexionActiva.BeginTrans
        If Not prs.BOF And Not prs.EOF Then
            prs.MoveFirst
        End If
        Do Until prs.EOF
            If val(prs!Asignacion) = 1 Then
                ' entonces se registra o actualiza
                   If VerificaGarantDoc(nTpoGarant, prs!TpoDoc) = False Then
                       'SE INSERTA
                        sSql = "INSERT INTO  GARANTDOC VALUES(" & nTpoGarant & "," & prs!TpoDoc & ")"
                        oConec.ConexionActiva.Execute sSql
                   End If
            Else
                If VerificaGarantDoc(prs!TpoGarantia, prs!TpoDoc) = True Then
                      sSql = "DELETE FROM GARANTDOC WHERE nConsValor=" & nTpoGarant & " AND nDocTpo=" & prs!TpoDoc
                       oConec.ConexionActiva.Execute sSql
                End If
            End If
            prs.MoveNext
        Loop
        oConec.ConexionActiva.CommitTrans
        oConec.CierraConexion
        Set oConec = Nothing
        ActualizarGarantDoc = True
    Exit Function
ErrHandler:
    oConec.ConexionActiva.RollbackTrans
    If Not oConec Is Nothing Then Set oConec = Nothing
    ActualizarGarantDoc = False
End Function

Public Function CargarNombrePersona(ByVal psCtaCod As String) As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim sNombre As String
    
    
    Set oConecta = New COMConecta.DCOMConecta
    sSql = "Select cPersNombre From Persona Where cPersCod='" & psCtaCod & "'"
    
    oConecta.AbreConexion
    Set rs = New ADODB.Recordset
    Set rs = oConecta.CargaRecordSet(sSql)
     
    If Not rs.EOF And Not rs.BOF Then
        sNombre = IIf(IsNull(rs!cPersNombre), "", rs!cPersNombre)
    End If
    Set rs = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    CargarNombrePersona = sNombre
End Function


Public Function VerificarCreditoAutomatico(ByVal psCtaCod As String) As Boolean
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim nCantidad As Integer
    ' verifica si el credito es automatico o no?
    
    sSql = "Select count(*) as Cantidad "
    sSql = sSql & " From Producto P"
    sSql = sSql & " Inner Join ColocacCred CC on P.cCtaCod=CC.cCtaCod"
    sSql = sSql & " Where P.cCtaCod='" & psCtaCod & "' and P.nPrdEstado=2000 and"
    sSql = sSql & " CC.nColocCondicion2 = " & gCredAutomatico

     Set oConecta = New COMConecta.DCOMConecta
     oConecta.AbreConexion
     Set rs = oConecta.CargaRecordSet(sSql)
     oConecta.CierraConexion
     Set oConecta = Nothing
     
     
     If Not rs.EOF And Not rs.BOF Then
        nCantidad = rs!cantidad
     End If
      

    Set rs = Nothing
    
    If nCantidad = 0 Then
        VerificarCreditoAutomatico = False
    Else
        VerificarCreditoAutomatico = True
    End If
End Function


Public Function ListaCreditosGarantias(ByVal pcPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    
'    ssql = "Select P.cCtaCod,G.cNumGarant,G.cDescripcion as cGarantias,PE.cPersNombre,CG.nGravado as nMonto,"
'    ssql = ssql & " G.nPorGravar as nDisponible"
'    ssql = ssql & " From Producto P"
'    ssql = ssql & " Inner Join ProductoPersona PP on PP.cCtaCod=P.cCtaCod and PP.nPrdPersRelac=" & gColRelPersTitular
'    ssql = ssql & " Inner Join ColocGarantia CG on CG.cCtaCod=P.cCtaCod"
'    ssql = ssql & " Inner Join Garantias G on G.cNumGarant=CG.cNumGarant"
'    ssql = ssql & "  Inner Join Persona PE on PE.cPersCod=PP.cPersCod"
'    ssql = ssql & " Where PP.cPersCod='" & pcPersCod & "'"

    'peac 20080111
    sSql = "exec stp_sel_ListadoRelaGarantiaConCreditos '" & pcPersCod & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaCreditosGarantias = oConec.CargaRecordSet(sSql)
    
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerMaxcNumGarant() As String
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sNumGarant As String
    
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    sSql = "Select Max(cNumGarant) as cNumGarant From Garantias"
    Set rs = oConec.CargaRecordSet(sSql)
    sNumGarant = Format(CLng(IIf(IsNull(rs!cNumGarant), "0", rs!cNumGarant)), "00000000")
    oConec.CierraConexion
    Set oConec = Nothing
    Set rs = Nothing
    ObtenerMaxcNumGarant = sNumGarant
End Function

Public Function DJ(ByVal psNumGarant As String, ByVal pdFecSis As Date) As ADODB.Recordset
    'Dim Rs As ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Dim SQL  As String
    Set oCon = New COMConecta.DCOMConecta
    'Set Rs = New ADODB.Recordset

    oCon.AbreConexion
    SQL = "SELECT GETDATE() "
    Set DJ = oCon.CargaRecordSet(SQL)
    oCon.CierraConexion
    Set oCon = Nothing
        
End Function


'Public Function CargarObjetosControles(ByRef prsBancos As ADODB.Recordset, _
'                                       ByRef prsTInmue As ADODB.Recordset, _
'                                       ByRef prsUbic As ADODB.Recordset, _
'                                       ByRef prsTGaran As ADODB.Recordset, _
'                                       ByRef prsMoneda As ADODB.Recordset, _
'                                       ByRef prsRelac As ADODB.Recordset, _
'                                       ByRef prsTDocum As ADODB.Recordset, _
'                                       ByRef prsSuperG As ADODB.Recordset, _
'                                       ByRef prsTContrato As ADODB.Recordset)
    
'peac 20071130 para activarlo despues
Public Function CargarObjetosControles(ByRef prsBancos As ADODB.Recordset, _
                                       ByRef prsTInmue As ADODB.Recordset, _
                                       ByRef prsUbic As ADODB.Recordset, _
                                       ByRef prsTGaran As ADODB.Recordset, _
                                       ByRef prsMoneda As ADODB.Recordset, _
                                       ByRef prsRelac As ADODB.Recordset, _
                                       ByRef prsTDocum As ADODB.Recordset, _
                                       ByRef prsSuperG As ADODB.Recordset, _
                                       ByRef prsTContrato As ADODB.Recordset, _
                                       ByRef prsClaseInmueble As ADODB.Recordset, _
                                       ByRef prsCategoria As ADODB.Recordset, _
                                       Optional ByRef prsClaseMueble As ADODB.Recordset = Nothing, _
                                       Optional ByRef prsTpoInscripcion As ADODB.Recordset = Nothing) 'EJVG20130128
    
'--*** peac 20071116 se agrego al compomente
'prsClaseInmueble
'prsCategoria
'--***
    
    Dim sSql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim oFinan As COMDPersona.DCOMInstFinac
    Dim oCons As COMDConstantes.DCOMConstantes
    Dim oUbic As COMDPersona.DCOMPersonas
    
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrorCargarObjetosControles
    
    Set oFinan = New COMDPersona.DCOMInstFinac
    Set prsBancos = oFinan.RecuperaBancos(True)
    Set oFinan = Nothing
    
    Set oUbic = New COMDPersona.DCOMPersonas
    Set prsUbic = oUbic.CargarUbicacionesGeograficas(, 1)
    Set oUbic = Nothing
    
    Set oCons = New COMDConstantes.DCOMConstantes
    Set prsTInmue = oCons.RecuperaConstantes(gGarantTpoInmueb)
        
    Set prsTGaran = oCons.RecuperaConstantes(gPersGarantia)
        
    Set prsMoneda = oCons.RecuperaConstantes(gMoneda)
        
    Set prsRelac = oCons.RecuperaConstantes(gPersRelGarantia)
        
    Set prsTDocum = oCons.RecuperaConstantes(gColocPigTipoDocumento)
        
    Set prsTContrato = oCons.RecuperaConstantes(1051)
    
    Set prsClaseMueble = oCons.RecuperaConstantes(1056) 'BRGO 20111205
    Set prsTpoInscripcion = oCons.RecuperaConstantes(10012) 'EJVG20130128
        
    Set oCons = Nothing
    
    Set prsSuperG = ListaSuperGarantias
    
    Set prsClaseInmueble = ListaClaseInmueble  'peac 20071116 para activarlo despues
    Set prsCategoria = ListaCategoria  'peac 20071116 para activarlo despues
    
    
    
    
    Exit Function
ErrorCargarObjetosControles:
    Set oConec = Nothing
    
End Function

Public Function CargarDatosGarantia(ByVal psNumGarant As String, ByRef prsGarantia As ADODB.Recordset, ByRef prsRelGarantia As ADODB.Recordset, ByRef prsGarantReal As ADODB.Recordset, ByRef prsGarantDJ As ADODB.Recordset, ByRef pbAsignadoACredito As Boolean, _
                                    Optional ByRef prsTablaValores As ADODB.Recordset = Nothing, _
                                    Optional ByRef pInmueblePoliza As ADODB.Recordset = Nothing, _
                                    Optional ByRef pTasador As ADODB.Recordset = Nothing, Optional ByRef pbGarantiaLegal As Boolean = False, _
                                    Optional ByRef pMueblePoliza As ADODB.Recordset = Nothing, Optional ByRef pDocumentoCompra As ADODB.Recordset = Nothing)
    'MADM 20110415 - rsGarantiaLegal
    'BRGO 20111205 - pMueblePoliza / pDocumentoCompra
    
    On Error GoTo ErrorCargarDatosGarantia
    
    pbGarantiaLegal = ExisteGarantiaLegal(psNumGarant)
    Set prsGarantia = RecuperaGarantia(psNumGarant)
    Set prsRelGarantia = RecuperaRelacPersonaGarantia(psNumGarant)
    Set prsGarantReal = RecuperaGarantiaReal(psNumGarant)
    Set prsGarantDJ = RecuperaGarantDeclaracionJur(psNumGarant)
    pbAsignadoACredito = PerteneceACredito(psNumGarant)
    Set prsTablaValores = RecuperaTablaValores_x_Garantia(psNumGarant)
    Set pInmueblePoliza = RecuperaDatosInmueblePoliza(psNumGarant)
    Set pMueblePoliza = RecuperaDatosMueblePoliza(psNumGarant) 'BRGO 20111205
    Set pDocumentoCompra = RecuperaDatosGarantDocCompra(psNumGarant) 'BRGO 20111205
    
    'Set pTasador = RecuperaDatosTasador(psNumGarant) '*** PEAC 20090722
    Set pTasador = Nothing
    
    Exit Function

ErrorCargarDatosGarantia:
    Err.Raise Err.Number, "Error", Err.Description
                                
End Function

Public Function CargarDatosGarantCredito(ByVal psCtaCod As String, _
                                        ByVal psPersCod As String, _
                                        ByVal pdFecSis As Date, _
                                        ByRef prsDatosGarCred As ADODB.Recordset, _
                                        ByRef prsPersGarant As ADODB.Recordset, _
                                        ByRef prsGarantias As ADODB.Recordset, _
                                        ByRef prsGarCred As ADODB.Recordset, _
                                        ByRef prsGarCF As ADODB.Recordset, _
                                        ByRef pnTipoCamb As Double, _
                                        ByRef pnMontoGravado As Double)
    
    Dim oCredito As COMDCredito.DCOMCredito
    Dim oCartaF As COMDCartaFianza.DCOMCartaFianza
    Dim oTipoCam  As COMDConstSistema.NCOMTipoCambio
    
    On Error GoTo ErrorCargarDatosGarantCredito
    
    Set oTipoCam = New COMDConstSistema.NCOMTipoCambio
    pnTipoCamb = oTipoCam.EmiteTipoCambio(pdFecSis, TCFijoMes)
    Set oTipoCam = Nothing
    
    Set oCredito = New COMDCredito.DCOMCredito
    
    'Verifica el Monto Gravado
    pnMontoGravado = oCredito.RecuperaMontoGarantiaCredito(psCtaCod, pdFecSis)
    
    Set prsDatosGarCred = oCredito.RecuperaDatosGarantiaCred(psCtaCod)
    Set prsPersGarant = oCredito.RecuperaRelacPers(psCtaCod)
    Set prsGarCred = oCredito.RecuperaGarantiasCredito(psCtaCod)
    Set oCredito = Nothing
    
    '****** Agregado para cargar automaticamente por el cCtaCod
    If psPersCod = "" And Not prsPersGarant.EOF Then
        psPersCod = prsPersGarant!cPersCod
    End If
    '**************************************************
    Set prsGarantias = RecuperaGarantiasPersona(psPersCod)
    
    Set oCartaF = New COMDCartaFianza.DCOMCartaFianza
    Set prsGarCF = oCartaF.RecuperaDatosGarantiaCF(psCtaCod)
    Set oCartaF = Nothing
    
    Exit Function

ErrorCargarDatosGarantCredito:
    Err.Raise Err.Number, "Cargar Datos Garantia Credito", Err.Description
                                
End Function

Public Function CargarControlesLiberaGarantia(ByRef prsDocum As ADODB.Recordset, _
                                       ByRef prsTipoGar As ADODB.Recordset)
    
    Dim oCons As COMDConstantes.DCOMConstantes
    
    On Error GoTo ErrorCargarControlesLiberaGarantia
    
    Set oCons = New COMDConstantes.DCOMConstantes
    Set prsTipoGar = oCons.RecuperaConstantes(gPersGarantia)
    Set oCons = Nothing
    
    Set prsDocum = RecuperaTiposDocumentosGarantia()
    
    Exit Function
    
ErrorCargarControlesLiberaGarantia:
    Err.Raise Err.Number, "Carga Controles Libera Garantia", Err.Description
End Function

Public Function EsSubTipoGarantiaTablaValores(ByVal pnSubTipoGarant As Integer) As Boolean

Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

    On Error GoTo ErrorEsSubTipoGarantiaTablaValores
    
    sSql = "SELECT nCodfiltro=ISNULL(nCodfiltro,0) FROM FiltroGarant WHERE nCodValor=" & pnSubTipoGarant
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    If rs.EOF Then
        EsSubTipoGarantiaTablaValores = False
        Exit Function
    End If
    
    If rs!nCodfiltro = 500 Then
        EsSubTipoGarantiaTablaValores = True
    Else
        EsSubTipoGarantiaTablaValores = False
    End If
    
    Exit Function

ErrorEsSubTipoGarantiaTablaValores:
    Err.Raise Err.Number, "Error", Err.Description
End Function

'Se Agrego para manejar los Tipos de Garantias
Public Function RecuperaDatosTipoGarantias(ByVal pnTipoGarantia As Integer, _
                                            ByRef prsDocum As ADODB.Recordset, _
                                            ByRef pnTipoRealizacion As Integer)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

    On Error GoTo ErrorRecuperaDatosTipoGarantias
    
    Set prsDocum = RecuperaTiposDocumGarantias(pnTipoGarantia)
    
    sSql = "SELECT nCodFiltro FROM FiltroGarant WHERE nCodValor=" & pnTipoGarantia & " AND nCodFiltro IN(300,400)"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    If rs.EOF Then
        pnTipoRealizacion = 0
    Else
        pnTipoRealizacion = rs!nCodfiltro
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaDatosTipoGarantias:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function EsSubTipoGarantiaInmueble(ByVal pnSubTipoGarant As Integer) As Boolean

Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

    On Error GoTo ErrorEsSubTipoGarantiaInmueble
    
    sSql = "SELECT nCodfiltro=ISNULL(nCodfiltro,0) FROM FiltroGarant WHERE nCodValor=" & pnSubTipoGarant
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
    
    If rs.EOF Then
        EsSubTipoGarantiaInmueble = False
        Exit Function
    End If
    
    If rs!nCodfiltro = 100 Then
        EsSubTipoGarantiaInmueble = True
    Else
        EsSubTipoGarantiaInmueble = False
    End If
    
    Exit Function

ErrorEsSubTipoGarantiaInmueble:
    Err.Raise Err.Number, "Error", Err.Description
End Function

Public Function RecuperaTablaValores(ByVal pnTipoCredito As Integer) As ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = "SELECT nTipoEval,cConsDescripcion from Scoring S INNER JOIN Constante C " & _
       " ON S.nTipoEval=C.nConsValor AND nConsCod=9063 AND nTipoTabla=" & pnTipoCredito

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set RecuperaTablaValores = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function RecuperaValoresTablaValores(ByVal pnTipoCredito As Integer, _
                                            ByVal pnTipoEval As Integer) As ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = "SELECT SG.cDescripcion,SG.nCodItem from Scoring_Detalle SG " & _
       " WHERE nTipoTabla=" & pnTipoCredito & " AND nTipoEval=" & pnTipoEval
'ARCV 14-08-2006
sSql = sSql & " UNION SELECT 'NINGUNO',0 "

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set RecuperaValoresTablaValores = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function RecuperaPuntajeTablaValores(ByVal pnTipoCredito As Integer, _
                                            ByVal pnTipoEval As Integer, _
                                            ByVal pnCodItem As Integer) As Double
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

sSql = "SELECT nValor FROM Scoring_Detalle " & _
       " WHERE nTipoTabla=" & pnTipoCredito & " AND nTipoEval=" & pnTipoEval & " AND nCodItem=" & pnCodItem

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSql)
If rs.EOF Then
    RecuperaPuntajeTablaValores = 0
Else
    RecuperaPuntajeTablaValores = rs!nValor
End If
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function RecuperaTablaValores_x_Garantia(ByVal psNumGarant As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = "SELECT SG.nTipoTabla,SG.nTipoEval,cConsDescripcion,SD.cDescripcion,SD.nValor,SD.nCodItem FROM Scoring_Garantias SG INNER JOIN  Scoring_Detalle SD" & _
       " ON SG.nTipoTabla =SD.nTipoTabla AND SG.nTipoEval =SD.nTipoEval AND SG.nCodItem =SD.nCodItem INNER JOIN Constante C " & _
       " ON C.nConsValor= SG.nTipoEval AND C.nConscod= 9063 WHERE SG.cNumGarant='" & psNumGarant & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set RecuperaTablaValores_x_Garantia = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function RecuperaDatosTablaValores_x_Garantia(ByVal psNumGarant As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = "SELECT Valor=ISNULL(SUM(nValor)*1000,0) FROM Garantias G INNER JOIN Scoring_Garantias SG ON G.cNumGarant =SG.cNumGarant " & _
       " INNER JOIN  Scoring_Detalle SD ON SG.nTipoTabla =SD.nTipoTabla AND SG.nTipoEval =SD.nTipoEval " & _
       " AND SG.nCodItem =SD.nCodItem INNER JOIN Constante C  ON C.nConsValor= SG.nTipoEval AND C.nConscod= 9063 " & _
       " WHERE SG.cNumGarant='" & psNumGarant & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set RecuperaDatosTablaValores_x_Garantia = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function
'ALPA**20080811*********************
Public Function RecuperaDatosGarantiaSanemiento(ByVal psNumGarant As String, ByVal psCtaCod As String, ByVal pnTESan As Integer) As ADODB.Recordset
'Public Function RecuperaDatosGarantiaSanemiento(ByVal psNumGarant As String, ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = "exec stp_sel_RecuperaDatosGarantiaSanemiento '" & psNumGarant & "','" & psCtaCod & "'," & pnTESan
'sSql = "exec stp_sel_RecuperaDatosGarantiaSanemiento '" & psNumGarant & "','" & psCtaCod & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set RecuperaDatosGarantiaSanemiento = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Sub InsertarGarantiaSaneamiento(ByVal psNumGarant As String, ByVal psCtaCod As String, ByVal pnTipoSan As Integer, ByVal pnPeriSan As Integer, ByVal pnMontSan As Double, ByVal pdFecSan As Date, ByVal psUsuariAdju As String, ByVal pnContador As Integer, ByVal pnMoneda As Integer, ByVal nTipoSaneame As Integer)
'Public Sub InsertarGarantiaSaneamiento(ByVal psNumGarant As String, ByVal psCtaCod As String, ByVal pnTipoSan As Integer, ByVal pnPeriSan As Integer, ByVal pnMontSan As Double, ByVal pdFecSan As Date, ByVal psUsuariAdju As String, ByVal pnContador As Integer)
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = "exec stp_ins_GarantiaSaneamiento '" & psNumGarant & "','" & psCtaCod & "'," & pnTipoSan & "," & pnPeriSan & "," & pnMontSan & ", '" & Format(pdFecSan, "YYYY/MM/DD") & "', '" & psUsuariAdju & "', " & pnContador & ", " & pnMoneda & "," & nTipoSaneame
'sSql = "exec stp_ins_GarantiaSaneamiento '" & psNumGarant & "','" & psCtaCod & "'," & pnTipoSan & "," & pnPeriSan & "," & pnMontSan & ", '" & Format(pdFecSan, "YYYY/MM/DD") & "', '" & psUsuariAdju & "', " & pnContador

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
oCon.CargaRecordSet (sSql)
oCon.CierraConexion
Set oCon = Nothing
End Sub

Public Function RecuperaDatosGarantiaRemate(ByVal psNumGarant As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = "exec stp_sel_RecuperaDatosGarantiaRemate '" & psNumGarant & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set RecuperaDatosGarantiaRemate = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Sub InsertarGarantiaRemate(ByVal psNumGarant As String, ByVal psCtaCod As String, ByVal nNumRemate As Integer, ByVal cLugarRemate As String, ByVal dFechaRemate As Date, ByVal dFechaSis As Date, ByVal nEstadoAdju As Integer, ByVal cPersCod As String, ByVal nMonto As Double, ByVal nContador As Integer, ByVal psMovAct As String, ByVal pnMoneda As Integer, ByVal psPerCod2 As String, ByVal nInteres As Double)
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = "exec stp_ins_GarantiaRemate '" & psNumGarant & "','" & psCtaCod & "', " & nNumRemate & ", '" & cLugarRemate & "', '" & Format(dFechaRemate, "yyyy/mm/dd") & "', '" & Format(dFechaSis, "yyyy/mm/dd") & "', " & nEstadoAdju & ", '" & cPersCod & "'," & nMonto & "," & nContador & ",'" & psMovAct & "'," & pnMoneda & ",'" & psPerCod2 & "'," & nInteres

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
oCon.CargaRecordSet (sSql)
oCon.CierraConexion
Set oCon = Nothing
End Sub
Public Function Reporte_Adjudicados138005(ByVal pdFecha As Date) As ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = "exec stp_sel_reporteGarantiasAdjudicadasxPeriodo '" & Format(pdFecha, "YYYY/MM/DD") & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set Reporte_Adjudicados138005 = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function ReporteGarantiasAdjudicadasContabilidad(ByVal pdFecha As Date, ByVal psAgencias As Variant) As ADODB.Recordset
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta

sSql = "exec stp_sel_ReporteGarantiasAdjudicadasContabilidad '" & Format(pdFecha, "YYYY/MM/DD") & "','" & psAgencias & "'"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set ReporteGarantiasAdjudicadasContabilidad = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function
'*************************************
'**ALPA 20080806**************************************
Public Sub dUpdateGarantiasAdjudicados(ByVal psGarantiaCod As String, ByVal pdFecha As Date, ByVal pnEstadoAdj As Integer, ByVal psUsuario As String)

Dim lsSQL As String

    On Error GoTo ErrordUpdateGarantiasAdjudicados
    lsSQL = "exec stp_upd_GarantiaAdjudicado '" & psGarantiaCod & "','" & Format(pdFecha, "YYYY/MM/DD") & "'," & pnEstadoAdj & ",'" & psUsuario & "'"
    'coConex.Ejecutar lsSQL
    coConex.CargaRecordSet (lsSQL)
    Exit Sub

ErrordUpdateGarantiasAdjudicados:
    Err.Raise Err.Number, "Error En Proceso dUpdateColocaciones", Err.Description
End Sub
'******************************************************
Public Sub dBeginTrans()
    'If pbConcurrencia Then
    '    coConex.ConexionActiva.Execute "SET TRANSACTION ISOLATION LEVEL SERIALIZABLE"
    'End If
    coConex.BeginTrans
End Sub
Public Sub dRollbackTrans()
    coConex.RollbackTrans
    coConex.Ejecutar "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
End Sub
Public Sub dCommitTrans()
    coConex.CommitTrans
    'coConex.Ejecutar "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
End Sub

Private Sub Class_Initialize()
     Dim loIni As COMConecta.DCOMClasIni

    Set loIni = New COMConecta.DCOMClasIni
        csConexion = loIni.CadenaConexion
        csNegocio = loIni.BaseNegocio
        csCentralPer = loIni.BasePersonas
        csCentralCom = loIni.BaseComunes
        csCentralImg = loIni.BaseImagenes
        csAdminist = loIni.BaseAdministracion
    Set loIni = Nothing

Set coConex = New COMConecta.DCOMConecta
Set oError = New COMConecta.COMErrorHandling

If coConex.AbreConexion(csConexion) = False Then
    Call oError.RaiseError(oError.MyUnhandledError, "DCOMGarantia:Initialize. Error en Conexion a Base de datos")
End If
End Sub

''*** PEAC 20090722
'Public Function RecuperaDatosTasador(ByVal psNumGarant As String) As ADODB.Recordset
'
'    Dim sSql As String
'    Dim oConecta As COMConecta.DCOMConecta
'    Set oConecta = New COMConecta.DCOMConecta
'
'    oConecta.AbreConexion
'    sSql = " Execute stp_sel_ObtieneDatosDelTasador '" & psNumGarant & "'"
'    Set RecuperaDatosTasador = oConecta.CargaRecordSet(sSql)
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'
'End Function

'*** PEAC 20090725
Public Function RecuperaDatosGarantiasPersona(ByVal psNumGarantia As String) As ADODB.Recordset

Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String

    On Error GoTo ErrorRecuperaDatosGarantiasPersona
    
    sSql = " exec stp_sel_RecuperaDatosGarantiasPersona '" & psNumGarantia & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaDatosGarantiasPersona = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function
    
ErrorRecuperaDatosGarantiasPersona:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function


'madm 20100511 -----------------------------
Public Function RecuperaCondcionCreditoApro(ByVal pcCtaCod As String, pdFecha As Date) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaCondcionCreditoApro
        
    sSql = " exec stp_sel_ObtieneCondicionCreditoApro '" & pcCtaCod & "' , '" & Format(pdFecha, "yyyymmdd") & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCondcionCreditoApro = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCondcionCreditoApro:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function
Public Function RecuperaCondicionCreditoDes(ByVal pcCtaCod As String, pdFecha As Date) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    On Error GoTo ErrorRecuperaCondicionCreditoDes
        
    sSql = " exec stp_sel_ObtieneCondicionCreditoDesp '" & pcCtaCod & "', '" & Format(pdFecha, "yyyymmdd") & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCondicionCreditoDes = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaCondicionCreditoDes:
    Set oConecta = Nothing
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

'madm 20110415 -
Public Function ExisteGarantiaLegal(ByVal psNumGar As String) As Boolean
Dim SQL As String
Dim Co As COMConecta.DCOMConecta
Dim rs As New ADODB.Recordset
Set Co = New COMConecta.DCOMConecta

  SQL = " exec stp_sel_GarantiaLegal '" & psNumGar & "'"

Co.AbreConexion
Set rs = Co.CargaRecordSet(SQL)
Co.CierraConexion
If rs.EOF And rs.BOF Then
    ExisteGarantiaLegal = False
ElseIf rs!nApruebaLegal = 1 Then
    ExisteGarantiaLegal = True
Else
    ExisteGarantiaLegal = False
End If
Set rs = Nothing
Set Co = Nothing
End Function

Public Sub dUpdateGarantiasLegal(ByVal psGarantiaCod As String, ByVal pvalor As Integer)
Dim Co As COMConecta.DCOMConecta
Dim lsSQL As String

    On Error GoTo ErrordUpdateGarantiasLegal
    Set Co = New COMConecta.DCOMConecta
    
    lsSQL = "exec stp_sel_UpdGarantiaLegal   '" & psGarantiaCod & "', " & pvalor & " "
    Co.AbreConexion
    Co.CargaRecordSet (lsSQL)
    Co.CierraConexion
    Set Co = Nothing
    Exit Sub

ErrordUpdateGarantiasLegal:
    Err.Raise Err.Number, "Error En Proceso dUpdateColocaciones", Err.Description
End Sub

Public Function CargarDatosGarantiaLegal(ByVal psNumGarant As String, ByRef prsGarantia As ADODB.Recordset, ByRef prsGarantReal As ADODB.Recordset, Optional ByRef pbGarantiaLegal As Boolean = False)
    'MADM 20110415 - rsGarantiaLegal
    On Error GoTo ErrorCargarDatosGarantiaLegal
    
    pbGarantiaLegal = ExisteGarantiaLegal(psNumGarant)
    Set prsGarantia = RecuperaGarantia(psNumGarant)
    Set prsGarantReal = RecuperaGarantiaRealMax(psNumGarant)
    Exit Function

ErrorCargarDatosGarantiaLegal:
    Err.Raise Err.Number, "Error", Err.Description
                                
End Function

'MADM 20110429
Public Sub dUpdateGarantiasVerificaLegal(ByVal psGarantiaCod As String, ByVal pvalor As Integer)
Dim Co As COMConecta.DCOMConecta
Dim lsSQL As String

    On Error GoTo ErrordUpdateGarantiasLegal
    Set Co = New COMConecta.DCOMConecta
    
    lsSQL = "exec stp_sel_UpdGarantiaVerificaLegal   '" & psGarantiaCod & "', " & pvalor & " "
    Co.AbreConexion
    Co.CargaRecordSet (lsSQL)
    Co.CierraConexion
    Set Co = Nothing
    Exit Sub

ErrordUpdateGarantiasLegal:
    Err.Raise Err.Number, "Error En Proceso dUpdateColocaciones", Err.Description
End Sub

Public Function CargarDatosGarantiaLegalVerifica(ByVal psNumGarant As String, ByRef prsGarantia As ADODB.Recordset, ByRef prsGarantReal As ADODB.Recordset, Optional pbGarantiaVerificaLegal As Boolean = False)
    On Error GoTo ErrorCargarDatosGarantiaLegalVerifica
    
    pbGarantiaVerificaLegal = ExisteGarantiaVerificaLegal(psNumGarant)
    Set prsGarantia = RecuperaGarantia(psNumGarant)
    Set prsGarantReal = RecuperaGarantiaRealMax(psNumGarant)
    Exit Function

ErrorCargarDatosGarantiaLegalVerifica:
    Err.Raise Err.Number, "Error", Err.Description
                                
End Function

'madm 20110415 -
Public Function ExisteGarantiaVerificaLegal(ByVal psNumGar As String) As Boolean
Dim SQL As String
Dim Co As COMConecta.DCOMConecta
Dim rs As New ADODB.Recordset
Set Co = New COMConecta.DCOMConecta
ExisteGarantiaVerificaLegal = False

  SQL = " exec stp_sel_GarantiaVerificaLegal '" & psNumGar & "'"

Co.AbreConexion
Set rs = Co.CargaRecordSet(SQL)
Co.CierraConexion
If rs.EOF And rs.BOF Then
    ExisteGarantiaVerificaLegal = False
ElseIf rs!nVerificaLegal = 1 Then
    ExisteGarantiaVerificaLegal = True
Else
    ExisteGarantiaVerificaLegal = False
End If
Set rs = Nothing
Set Co = Nothing
End Function
'END MADM

Public Function RecuperaGarantiaRealMax(ByVal psNumGarant As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_RecuperaGarantiaRealMax '" & psNumGarant & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaGarantiaRealMax = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

'MADM 20110601
Public Function RecuperaGarantiaRealMaxAprobacion(ByVal pcCtaCod As String, ByVal psNumGarant As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
    sSql = "exec stp_sel_RecuperaGarantiaRealMaxAprobacion '" & pcCtaCod & "','" & psNumGarant & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaGarantiaRealMaxAprobacion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Function

'end madm
''ALPA
'Public Function ValidaAhorro(ByVal psCtaCod As String) As Boolean
'    Dim oConec As COMConecta.DCOMConecta
'    Dim sSql As String
'    Dim rs As ADODB.Recordset
'    Dim nValor As Integer
'
'    sSql = "Select Count(*) as nCantidad"
'    sSql = sSql & " From Producto P inner join Captaciones C on P.cCtaCod=C.cCtaCod"
'    sSql = sSql & " Where SubString(P.cCtaCod,6,3) in ('232') and P.nPrdEstado not in (1300,1400)"
'    sSql = sSql & " and P.cCtaCod='" & psCtaCod & "'"
'
'    Set oConec = New COMConecta.DCOMConecta
'    oConec.AbreConexion
'    Set rs = oConec.CargaRecordSet(sSql)
'    oConec.CierraConexion
'
'    If Not rs.EOF And Not rs.BOF Then
'        nValor = rs!nCantidad
'    End If
'    Set rs = Nothing
'
'    If nValor > 0 Then
'        ValidaAhorro = True
'    Else
'       ValidaAhorro = False
'    End If
'End Function
''*********
'***Agregado por ELRO el 20120329, según OYP-RFC022-2012
Public Sub actulizarGarantiaPF(ByVal pvFecha As String, ByVal pcCtaCod As String)
Dim oDCOMConecta As COMConecta.DCOMConecta
Dim lsSQL As String

    On Error GoTo ErroractulizarGarantiaPF
    Set oDCOMConecta = New COMConecta.DCOMConecta

    lsSQL = "exec stp_upd_ActualizarGarantiaPF '" & pvFecha & "', '" & pcCtaCod & "' "
    oDCOMConecta.AbreConexion
    oDCOMConecta.CargaRecordSet (lsSQL)
    oDCOMConecta.CierraConexion
    Set oDCOMConecta = Nothing
    lsSQL = ""
    Exit Sub

ErroractulizarGarantiaPF:
    Err.Raise Err.Number, "Error En Proceso actulizarGarantiaPF", Err.Description
End Sub

Public Function devolverFechaGarantiaPF(ByVal pcCtaCod As String) As Date
Dim oDCOMConecta As COMConecta.DCOMConecta
Dim rsFecha As ADODB.Recordset
Dim lsSQL As String
Set oDCOMConecta = New COMConecta.DCOMConecta
Set rsFecha = New ADODB.Recordset

    On Error GoTo ErroractulizarGarantiaPF

    lsSQL = "exec stp_sel_DevolverFechaGarantiaPF '" & pcCtaCod & "' "
    oDCOMConecta.AbreConexion
    Set rsFecha = oDCOMConecta.CargaRecordSet(lsSQL)
    If Not rsFecha.BOF And Not rsFecha.EOF Then
        devolverFechaGarantiaPF = rsFecha!dGarantia
    Else
        devolverFechaGarantiaPF = CDate("01/01/1900")
    End If
    oDCOMConecta.CierraConexion
    Set oDCOMConecta = Nothing
    lsSQL = ""
    Exit Function

ErroractulizarGarantiaPF:
    Err.Raise Err.Number, "Error En Proceso actulizarGarantiaPF", Err.Description
End Function
'***Fin Agregado por ELRO***************************
'WIOR 20120616 ***************************************************
Public Sub RegistrarMovimientoGarantia(ByVal pcNumGarant As String, ByVal pcUltimaActualizacion As String)
Dim oDCOMConecta As COMConecta.DCOMConecta
Dim lsSQL As String

    On Error GoTo ErrorRegistrarMovimientoGarantia
    Set oDCOMConecta = New COMConecta.DCOMConecta

    lsSQL = "insert into GarantAct(cNumGarant,cUltimaActualizacion) values('" & pcNumGarant & "','" & pcUltimaActualizacion & "')"
    oDCOMConecta.AbreConexion
    oDCOMConecta.CargaRecordSet (lsSQL)
    oDCOMConecta.CierraConexion
    Set oDCOMConecta = Nothing
    lsSQL = ""
    Exit Sub

ErrorRegistrarMovimientoGarantia:
    Err.Raise Err.Number, "Error En Proceso RegistrarMovimientoGarantia", Err.Description
End Sub

Public Sub ActualizarMovimientoGarantia(ByVal pnTipo As Integer, ByVal pcNumGarant As String, Optional ByVal pcUltimaActualizacion As String = "", Optional ByVal pcMovVerificacion As String = "")
Dim oDCOMConecta As COMConecta.DCOMConecta
Dim lsSQL As String

    On Error GoTo ErrorActualizarMovimientoGarantia
    Set oDCOMConecta = New COMConecta.DCOMConecta
    If pnTipo = 1 Then
        lsSQL = "update GarantAct set cUltimaActualizacion='" & pcUltimaActualizacion & "'," & _
                "cMovVerificacion='' " & _
                " where cNumGarant='" & pcNumGarant & "'"
    ElseIf pnTipo = 2 Then
        lsSQL = "update GarantAct set cMovVerificacion='" & pcMovVerificacion & "'" & _
                 " where cNumGarant='" & pcNumGarant & "'"
    End If
    oDCOMConecta.AbreConexion
    oDCOMConecta.CargaRecordSet (lsSQL)
    oDCOMConecta.CierraConexion
    Set oDCOMConecta = Nothing
    lsSQL = ""
    Exit Sub

ErrorActualizarMovimientoGarantia:
    Err.Raise Err.Number, "Error En Proceso ActualizarMovimientoGarantia", Err.Description
End Sub

Public Function ObtenerUltimaActualizacion(ByVal psNumGarant As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
 On Error GoTo ErrorObtenerUltimaActualizacion
 
 sSql = "select cNumGarant,isnull(cUltimaActualizacion,'') cUltimaActualizacion," & _
    "isnull(cMovVerificacion ,'') cMovVerificacion " & _
    "from GarantAct where cNumGarant='" & psNumGarant & "' "
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ObtenerUltimaActualizacion = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerUltimaActualizacion:
    Err.Raise Err.Number, "Error En Proceso ObtenerUltimaActualizacion", Err.Description
End Function
'WIOR FIN ********************************************************
'WIOR 20130122 ***************************************************
Public Function ActualizarGarAdmCred(ByVal psNumGarant As String, Optional ByVal pbActivado As Boolean = False, Optional ByVal pvPoliza As Variant) As Boolean
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
On Error GoTo ErrorActualizarGarAdmCred
Set oConecta = New COMConecta.DCOMConecta

oConecta.AbreConexion
If Trim(pvPoliza(1, 0)) = "1" Then
    sSql = "exec stp_del_QuitarPolizaGar " & 1 & ",'" & psNumGarant & "'"
    oConecta.Ejecutar sSql
    If pbActivado Then
        sSql = "exec stp_ins_RegistrarPolizaInmueble '" & psNumGarant & "','" & pvPoliza(1, 1) & "'," & pvPoliza(1, 2) & "," & pvPoliza(1, 3) & "," & pvPoliza(1, 4) & "," & pvPoliza(1, 5) & "," & pvPoliza(1, 6) & "," & pvPoliza(1, 7)
        oConecta.Ejecutar sSql
    End If
    sSql = "exec stp_upd_ActualizaPolizaInmueble '" & psNumGarant & "'," & IIf(pbActivado, 1, 0)
    oConecta.Ejecutar sSql
ElseIf Trim(pvPoliza(1, 0)) = "2" Then
    sSql = "exec stp_del_QuitarPolizaGar " & 2 & ",'" & psNumGarant & "'"
    oConecta.Ejecutar sSql
    If pbActivado Then
        sSql = "exec  stp_ins_RegistrarPolizaMobiliaria '" & psNumGarant & "','" & pvPoliza(1, 1) & "'," & pvPoliza(1, 2) & "," & pvPoliza(1, 3)
        oConecta.Ejecutar sSql
    End If
End If
ActualizarGarAdmCred = True
oConecta.CierraConexion

Set oConecta = Nothing

Exit Function
ErrorActualizarGarAdmCred:
    ActualizarGarAdmCred = False
    Err.Raise Err.Number, "Error En Proceso ActualizarGarAdmCred", Err.Description
End Function
'WIOR FIN ********************************************************
'ALPA20140301 ***
Public Function RecuperaGarantiasReemplazadas(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    sSql = "exec stp_garantiaReemplazada '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaGarantiasReemplazadas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function RecuperaGarantiasReemplazante(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    sSql = "exec stp_garantiaReemplazante '" & psCtaCod & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaGarantiasReemplazante = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Function RecuperaCantidadCreditosxGarantia(ByVal psNumGarant As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    sSql = "exec stp_sel_ObtenerCantidadCreditosxGarantia '" & psNumGarant & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCantidadCreditosxGarantia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Sub GuardarReemplazoDeGarantia(ByVal psCtaCod As String, ByVal psNumGarantAnt As String, ByVal psNumGarantNueva As String, ByVal pnMontoCoberturaMN As Currency, ByVal pnMontoCoberturaME As Currency, ByVal psMoneda As String, ByVal pnTipoCambio As Currency)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    sSql = "exec stp_ins_CambiarGarantia '" & psCtaCod & "', '" & psNumGarantAnt & "', '" & psNumGarantNueva & "', " & pnMontoCoberturaMN & "," & pnMontoCoberturaME & ", '" & psMoneda & "', " & pnTipoCambio & ""
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Sub
'******************************************************************
'EJVG20140422 ***
Public Sub GenerarDocsPendiente(ByVal pbFiltraAge As Boolean, ByVal psAgeCod As String, ByVal pdFecha As Date, ByVal psTblNombre As String)
    Dim lsSQL As String
    lsSQL = "Exec stp_ins_ERS0342014_DocsPendiente " & IIf(pbFiltraAge, 1, 0) & ", '" & psAgeCod & "', '" & Format(pdFecha, "yyyymmdd") & "','" & psTblNombre & "'"
    coConex.Ejecutar (lsSQL)
End Sub
Public Function ConsultarDocsPendiente(ByVal pnTblId As Integer, ByVal psTblNombre As String) As ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "Exec stp_sel_ERS0342014_DocsPendiente " & pnTblId & ", '" & psTblNombre & "'"
    Set ConsultarDocsPendiente = coConex.CargaRecordSet(lsSQL)
End Function
Public Sub EliminarDocsPendiente(ByVal psTblNombre As String)
    Dim lsSQL As String
    lsSQL = "Exec stp_del_ERS0342014_DocsPendiente '" & psTblNombre & "'"
    coConex.Ejecutar (lsSQL)
End Sub
Public Function MostrarDocsPendiente(ByVal pdFecha As Date) As Boolean
    Dim oRs As New ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "SELECT dbo.fnc_ERS0342014_MostrarDocsPendiente('" & Format(pdFecha, "yyyymmdd") & "') bMostrar"
    Set oRs = coConex.CargaRecordSet(lsSQL)
    If Not oRs.EOF Then
        MostrarDocsPendiente = oRs!bMostrar
    End If
    oRs.Close
    Set oRs = Nothing
End Function
Public Function NivelVerDocsPendiente(ByVal psGrupoUsu As String, ByVal psUserCod As String, ByVal psRHCargoCod As String) As Integer
    Dim oRs As New ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "SELECT dbo.fnc_ERS0342014_NivelVerDocsPendiente('" & psGrupoUsu & "','" & psUserCod & "','" & psRHCargoCod & "') nNivel"
    Set oRs = coConex.CargaRecordSet(lsSQL)
    If Not oRs.EOF Then
        NivelVerDocsPendiente = oRs!nNivel
    End If
    oRs.Close
    Set oRs = Nothing
End Function
'END EJVG *******
'RECO20150309 ERS149-2015*************************************************************
Public Function RecuperaListaGarantPolizaPersona(ByVal psPersCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
     On Error GoTo ErrorRecuperaListaGarantPolizaPersona
     
     sSql = "stp_sel_RecuperaListaGarantPolizaPersona '" & psPersCod & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set RecuperaListaGarantPolizaPersona = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorRecuperaListaGarantPolizaPersona:
        Err.Raise Err.Number, "Error En Proceso RecuperaListaGarantPolizaPersona", Err.Description
End Function
Public Function RegistraSolicitudCoberturaSeguro(ByVal psPersCod As String, ByVal psNumGaran As String, ByVal psFecSiniestro As String, ByVal psGlosa As String _
                                                 , ByVal psNumCarta As String, ByVal psFecEmision As String, ByVal pnTipoSeguro As TipoSeguro _
                                                 , Optional ByVal psPersGestiona As String = "", Optional ByVal pnPersRelac As Integer = 0, Optional ByVal psAgeCodSiniestro As String = "" _
                                                 ) As Integer
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oRs As New Recordset
     On Error GoTo ErrorRegistraSolicitudCoberturaSeguro
     
     sSql = "stp_ins_RegistraSolicitudCoberturaSeguro '" & psPersCod & "','" & psNumGaran & "','" & psFecSiniestro & "','" & psGlosa _
                                                 & "','" & psNumCarta & "','" & Format(psFecEmision, "yyyyMMdd") & "'," & pnTipoSeguro _
                                                 & ",'" & psPersGestiona & "'," & pnPersRelac & ",'" & psAgeCodSiniestro & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set oRs = oConecta.CargaRecordSet(sSql)
        If Not (oRs.EOF And oRs.BOF) Then
            RegistraSolicitudCoberturaSeguro = oRs!nId
        End If
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorRegistraSolicitudCoberturaSeguro:
        Err.Raise Err.Number, "Error En Proceso RegistraSolicitudCoberturaSeguro", Err.Description
End Function
Public Function ObtieneDatosSolicitudCobert(ByVal psNumGarant As String, ByVal pnTpoSeguro As TipoSeguro) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oRs As New Recordset
     On Error GoTo ErrorObtieneDatosSolicitudCobert
     
        sSql = "stp_sel_ObtieneDatosSolicitudCobert '" & psNumGarant & "'," & pnTpoSeguro
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set ObtieneDatosSolicitudCobert = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorObtieneDatosSolicitudCobert:
        Err.Raise Err.Number, "Error En Proceso ObtieneDatosSolicitudCobert", Err.Description
End Function
Public Function ObtieneDocAdjuntosSolicitud(ByVal pnIdSolic As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oRs As New Recordset
     On Error GoTo ErrorObtieneDocAdjuntosSolicitud
     
     sSql = "stp_sel_ObtieneDocAdjuntosSolicitud " & pnIdSolic
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set ObtieneDocAdjuntosSolicitud = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorObtieneDocAdjuntosSolicitud:
        Err.Raise Err.Number, "Error En Proceso ObtieneDocAdjuntosSolicitud", Err.Description
End Function
Public Function ObtieneNumCartaSolicCobert(ByVal psPersCod As String) As Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oRs As New Recordset
     On Error GoTo ErrorObtieneNumCartaSolicCobert
     
     sSql = "stp_sel_ObtieneNumCartaSolicCobert '" & psPersCod & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set ObtieneNumCartaSolicCobert = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorObtieneNumCartaSolicCobert:
        Err.Raise Err.Number, "Error En Proceso ObtieneNumCartaSolicCobert", Err.Description
End Function
Public Sub RegistraResolucionSolicitud(ByVal pnIdSolicitud As Integer, ByVal pnTpoDoc As Integer, ByVal pnMoneda As Integer, ByVal pnMonto As Double _
                                             , ByVal psFecEmision As String, ByVal psFecDeposito As String, ByVal psNumero As String, ByVal psCtaCodBanco As String _
                                             , ByVal pnEstado As Integer, ByVal psGlosa As String, ByVal pnTpoSeguro As TipoSeguro)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
On Error GoTo ErrorRegistraResolucionSolicitud
Set oConecta = New COMConecta.DCOMConecta

oConecta.AbreConexion
sSql = "stp_ins_RegistraResolucionSolicitud " & pnIdSolicitud & "," & pnTpoDoc & "," & pnMoneda & "," & pnMonto & ",'" & psFecEmision & "','" & _
                                                     psFecDeposito & "','" & psNumero & "','" & psCtaCodBanco & "','" & psGlosa & "'," & pnEstado & "," & pnTpoSeguro
oConecta.Ejecutar sSql
oConecta.CierraConexion

Set oConecta = Nothing

Exit Sub
ErrorRegistraResolucionSolicitud:
    Err.Raise Err.Number, "Error En Proceso RegistraResolucionSolicitud", Err.Description
End Sub
Public Sub RegsitraHistorialSegSolic(ByVal pnIdSolicitud As Integer, ByVal psDescripcion As String, ByVal psNumCarta As String, ByVal psGlosa As String _
                                    , ByVal psFecRespuesta As String, ByVal psFecEstado As String, ByVal pnEstado As Integer, ByVal pnTpoSeguro As TipoSeguro)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
On Error GoTo ErrorRegsitraHistorialSegSolic
Set oConecta = New COMConecta.DCOMConecta

oConecta.AbreConexion
sSql = "stp_ins_RegsitraHistorialSegSolic " & pnIdSolicitud & ",'" & psDescripcion & "','" & psNumCarta & "','" & psGlosa & "','" & psFecRespuesta & "','" & psFecEstado & "'," & pnEstado & "," & pnTpoSeguro
oConecta.Ejecutar sSql
oConecta.CierraConexion

Set oConecta = Nothing

Exit Sub
ErrorRegsitraHistorialSegSolic:
    Err.Raise Err.Number, "Error En Proceso RegsitraHistorialSegSolic", Err.Description
End Sub
Public Function BuscaSolicitudCobert(ByVal psNumCarta As String, ByVal psEstados As String, ByVal pnTpoSeguro As TipoSeguro) As Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oRs As New Recordset
     On Error GoTo ErrorBuscaSolicitudCobert
     
     sSql = "stp_sel_BuscaSolicitudCobert '" & psNumCarta & "','" & psEstados & "'," & pnTpoSeguro
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set BuscaSolicitudCobert = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorBuscaSolicitudCobert:
        Err.Raise Err.Number, "Error En Proceso BuscaSolicitudCobert", Err.Description
End Function
Public Sub RegistraRechazoSolicitud(ByVal pnIdSolicitud As Integer, ByVal psFecRespuesta As String, ByVal psGlosa As String, ByVal psNumCarta As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
On Error GoTo ErrorRegistraRechazoSolicitud
Set oConecta = New COMConecta.DCOMConecta

oConecta.AbreConexion
sSql = "stp_ins_RegistraRechazoSolicitud " & pnIdSolicitud & ",'" & psFecRespuesta & "','" & psGlosa & "','" & psNumCarta & "'"
oConecta.Ejecutar sSql
oConecta.CierraConexion

Set oConecta = Nothing

Exit Sub
ErrorRegistraRechazoSolicitud:
    Err.Raise Err.Number, "Error En Proceso RegistraRechazoSolicitud", Err.Description
End Sub
Public Sub ActualizaEstadoSolicCobert(ByVal pnEstado As Integer, ByVal pnIdSolicitud As Integer, ByVal psFecReg As String, ByVal pnTpoSeguro As TipoSeguro)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
On Error GoTo ErrorActualizaEstadoSolicCobert
Set oConecta = New COMConecta.DCOMConecta

oConecta.AbreConexion
sSql = "stp_upd_ActualizaEstadoSolicCobert " & pnEstado & "," & pnIdSolicitud & ",'" & psFecReg & "'," & pnTpoSeguro
oConecta.Ejecutar sSql
oConecta.CierraConexion

Set oConecta = Nothing

Exit Sub
ErrorActualizaEstadoSolicCobert:
    Err.Raise Err.Number, "Error En Proceso RegistraRechazoSolicitud", Err.Description
End Sub
Public Function ObtieneSolicitCoberXGarant(ByVal psNumGarant As String) As Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oRs As New Recordset
     On Error GoTo ErrorObtieneSolicitCoberXGarant
     
     sSql = "stp_sel_ObtieneSolicitCoberXGarant '" & psNumGarant & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set ObtieneSolicitCoberXGarant = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorObtieneSolicitCoberXGarant:
        Err.Raise Err.Number, "Error En Proceso ObtieneSolicitCoberXGarant", Err.Description
End Function
Public Function ObtieneDadosSolicitusContraIncendioExtornar(ByVal psPersCod As String, ByVal psFecDesde As String, ByVal psFecHasta As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oRs As New Recordset
     On Error GoTo ErrorObtieneDadosSolicitusContraIncendioExtornar
     
     sSql = "stp_sel_ObtieneDadosSolicitusContraIncendioExtornar '" & psPersCod & "','" & psFecDesde & "','" & psFecHasta & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set ObtieneDadosSolicitusContraIncendioExtornar = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorObtieneDadosSolicitusContraIncendioExtornar:
        Err.Raise Err.Number, "Error En Proceso ObtieneDadosSolicitusContraIncendioExtornar", Err.Description
End Function
Public Function ObtieneDatosExtornarSeg(ByVal psPersCod As String, ByVal psFecDesde As String, ByVal psFecHasta As String, ByVal pnTpoSeguro As TipoSeguro)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oRs As New Recordset
     On Error GoTo ErrorObtieneDatosExtornarSeg
     
        sSql = "stp_sel_ObtieneDatosExtornarSeg '" & psPersCod & "','" & psFecDesde & "','" & psFecHasta & "'," & pnTpoSeguro
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set ObtieneDatosExtornarSeg = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorObtieneDatosExtornarSeg:
        Err.Raise Err.Number, "Error En Proceso ObtieneDatosExtornarSegCI", Err.Description
End Function
Public Sub ExtornoSolicitud(ByVal pnIdSolicitud As Integer, ByVal pnEstado As Integer, ByVal psFecReg As String, ByVal pnTpoSeguro As TipoSeguro)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
On Error GoTo ErrorExtornoSolicitud
Set oConecta = New COMConecta.DCOMConecta

oConecta.AbreConexion
sSql = "stp_upd_ExtornoSolicitud " & pnIdSolicitud & "," & pnEstado & ",'" & psFecReg & "'," & pnTpoSeguro
oConecta.Ejecutar sSql
oConecta.CierraConexion

Set oConecta = Nothing

Exit Sub
ErrorExtornoSolicitud:
    Err.Raise Err.Number, "Error En Proceso ExtornoSolicitud", Err.Description
End Sub
Public Sub EliminaRegistroEditar(ByVal pnIdSolic As Integer, ByVal pnTpoSeguro As TipoSeguro)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorEliminaRegistroEditar
    Set oConecta = New COMConecta.DCOMConecta
    
    oConecta.AbreConexion
    sSql = "stp_del_EliminaRegistroEditar " & pnIdSolic & "," & pnTpoSeguro
    oConecta.Ejecutar sSql
    oConecta.CierraConexion
    
    Set oConecta = Nothing
    
    Exit Sub
ErrorEliminaRegistroEditar:
        Err.Raise Err.Number, "Error En Proceso EliminaRegistroEditar", Err.Description
End Sub
Public Sub ActulizaSolicitudSeg(ByVal pnIdSolic As Integer, ByVal psFecSiniestro As String, ByVal psGlosa As String, ByVal pnTpoSeguro As TipoSeguro)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    On Error GoTo ErrorActulizaSolicitudSeg
    Set oConecta = New COMConecta.DCOMConecta
    
    oConecta.AbreConexion
    sSql = "stp_upd_ActulizaSolicitudSeg " & pnIdSolic & ",'" & psFecSiniestro & "','" & psGlosa & "'," & pnTpoSeguro
    oConecta.Ejecutar sSql
    oConecta.CierraConexion
    
    Set oConecta = Nothing
    
    Exit Sub
ErrorActulizaSolicitudSeg:
        Err.Raise Err.Number, "Error En Proceso ActulizaSolicitudSeg", Err.Description
End Sub
Public Function SegSolicitudHistorial(ByVal pnIdSolicitud As Long, ByVal pnTpoSeguro As TipoSeguro) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrorSegSolicitudHistorial
    sSql = "EXEC stp_sel_SegSolicitudHistorial '" & pnIdSolicitud & "'," & pnTpoSeguro
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set SegSolicitudHistorial = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorSegSolicitudHistorial:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'RECO FIN*********************************************************
'EJVG20150326 ***
Public Function ObtenerDocsGarantias(Optional ByVal pnCodConfig As Long = 0, Optional ByVal pnClasificacion As Long = 0) As ADODB.Recordset
    Dim lsSQL As String
    'lsSQL = "EXEC stp_sel_GarantConfigDocs " & pnCodConfig
    lsSQL = "EXEC stp_sel_GarantConfigDocs " & pnCodConfig & "," & pnClasificacion
    Set ObtenerDocsGarantias = coConex.CargaRecordSet(lsSQL)
End Function
Public Function ObtenerDocsGarantiasXID(Optional ByVal pnCodConfig As Long = 0) As ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "EXEC stp_sel_GarantDocsXID " & pnCodConfig
    Set ObtenerDocsGarantiasXID = coConex.CargaRecordSet(lsSQL)
End Function
Public Function ObtenerConfigGarantParam(Optional ByVal pnClasificacion As Long = 0) As ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "EXEC stp_sel_GarantConfigParamClas " & pnClasificacion
    Set ObtenerConfigGarantParam = coConex.CargaRecordSet(lsSQL)
End Function
Public Function ObtenerConfigGarantParamTpoValor(Optional ByVal pnClasificacion As Long = 0) As ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "EXEC stp_sel_GarantConfigParamTpoValor " & pnClasificacion
    Set ObtenerConfigGarantParamTpoValor = coConex.CargaRecordSet(lsSQL)
End Function
Public Function ObtenerConfigGarantTpoValor(Optional ByVal pnCodConfig As Long = 0) As ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "EXEC stp_sel_GarantTpoValorXID " & pnCodConfig
    Set ObtenerConfigGarantTpoValor = coConex.CargaRecordSet(lsSQL)
End Function
Public Function ObtenerConfigGarant() As ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "EXEC stp_sel_GarantRegistradas"
    Set ObtenerConfigGarant = coConex.CargaRecordSet(lsSQL)
End Function
Public Function ObtenerConfigGarantXID(ByVal pnGarantConfigID As Long) As ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "EXEC stp_sel_GarantConfigXID " & pnGarantConfigID
    Set ObtenerConfigGarantXID = coConex.CargaRecordSet(lsSQL)
End Function
Public Function RegistarActConfigGarant(ByVal pnCod As Long, ByVal pnClasificacion As Integer, ByVal pnObjGarantCod As Integer, ByVal pcObjGarantDesc As String, _
                                    ByVal pnConstitucionCod As Integer, ByVal pcConstitucionDesc As String, ByVal pnLeasingCod As Integer, _
                                    ByVal pcLeasingDesc As String, ByVal pnSinTramite As Integer, ByVal pbActivado As Boolean) As Long

    Dim oRs As New ADODB.Recordset
    Dim lsSQL As String
    
    lsSQL = "EXEC stp_ins_upd_GarantConfig " & pnCod & "," & pnClasificacion & "," & pnObjGarantCod & ",'" & pcObjGarantDesc & "'," & _
                                        pnConstitucionCod & ",'" & pcConstitucionDesc & "'," & pnLeasingCod & ",'" & _
                                        pcLeasingDesc & "'," & pnSinTramite & "," & IIf(pbActivado, 1, 0)
    
    Set oRs = coConex.CargaRecordSet(lsSQL)
    If Not (oRs.EOF And oRs.BOF) Then
        RegistarActConfigGarant = CLng(oRs!nGarantConfigID)
    End If
    Set oRs = Nothing
End Function
Public Sub RegistrarConfigGarantDocs(ByVal pnGarantConfigID As Long, ByVal pnDocTpo As Long)
    Dim lsSQL As String
    lsSQL = "EXEC stp_ins_GarantConfigDocs " & pnGarantConfigID & "," & pnDocTpo
    coConex.Ejecutar (lsSQL)
End Sub
Public Sub RegistrarConfigGarantTpoValor(ByVal pnGarantConfigID As Long, ByVal pnTpoValor As Long)
    Dim lsSQL As String
    lsSQL = "EXEC stp_ins_GarantConfigTpoValor " & pnGarantConfigID & "," & pnTpoValor
    coConex.Ejecutar (lsSQL)
End Sub
Public Sub EliminaConfigGarantDocs(ByVal pnGarantConfigID As Long)
    Dim lsSQL As String
    lsSQL = "EXEC stp_del_GarantConfigDocs " & pnGarantConfigID
    coConex.Ejecutar (lsSQL)
End Sub
Public Sub EliminaConfigGarantTpoValor(ByVal pnGarantConfigID As Long)
    Dim lsSQL As String
    lsSQL = "EXEC stp_del_GarantConfigTpoValor " & pnGarantConfigID
    coConex.Ejecutar (lsSQL)
End Sub

Public Function GarantValidaCod(ByVal pnCod As Long, ByVal pnTipo As Long) As ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "EXEC stp_sel_GarantValidaCod " & pnCod & "," & pnTipo
    Set GarantValidaCod = coConex.CargaRecordSet(lsSQL)
End Function

Public Function ObtenerConfigGarantClas(ByVal pnClasificacion As Long) As ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "EXEC stp_sel_GarantConfigXClasificacion " & pnClasificacion
    Set ObtenerConfigGarantClas = coConex.CargaRecordSet(lsSQL)
End Function

Private Sub Class_Terminate()
coConex.CierraConexion
Set coConex = Nothing
End Sub
'END EJVG *******
'RECO20150416 ERS001-2015***************************
Public Function ObtieneCreditosPorGarantia(ByVal psNumGarant As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oRs As New Recordset
     On Error GoTo ErrorObtieneCreditosPorGarantia
     
        sSql = "stp_sel_ObtieneCreditosPorGarantia '" & psNumGarant & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set ObtieneCreditosPorGarantia = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorObtieneCreditosPorGarantia:
        Err.Raise Err.Number, "Error En Proceso ObtieneDatosExtornarSegCI", Err.Description
End Function
'RECO FIN ******************************************
'RECO20150509 ERS023-2015*************************************************************************
Public Function ObtieneGarantMultiRiesgoMYPE(ByVal psPersCod As String, Optional ByVal psCtaCod As String = "")
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oRs As New Recordset
     On Error GoTo ErrorObtieneGarantMultiRiesgoMYPE
     
        sSql = "stp_sel_ObtieneGarantMultiRiesgoMYPE '" & psPersCod & "','" & psCtaCod & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set ObtieneGarantMultiRiesgoMYPE = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorObtieneGarantMultiRiesgoMYPE:
        Err.Raise Err.Number, "Error En Proceso ObtieneGarantMultiRiesgoMYPE", Err.Description
End Function
Public Sub RegistroGastoMicroSegMYPE(ByVal psCtaCod As String, ByVal psNumGarant As String, ByVal pnValorAseg As Double, ByVal psFecReg As String)
    Dim lsSQL As String
    lsSQL = "EXEC stp_ins_RegistroGastoMicroSegMYPE '" & psCtaCod & "','" & psNumGarant & "','" & pnValorAseg & "','" & psFecReg & "'"
    coConex.Ejecutar (lsSQL)
End Sub
Public Function RecuperaDatosClienteDJMYPE(ByVal psCtaCod As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oRs As New Recordset
     On Error GoTo ErrorRecuperaDatosClienteDJMYPE
     
        sSql = "stp_sel_RecuperaDatosClienteDJMYPE '" & psCtaCod & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set RecuperaDatosClienteDJMYPE = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorRecuperaDatosClienteDJMYPE:
        Err.Raise Err.Number, "Error En Proceso RecuperaDatosClienteDJMYPE", Err.Description
End Function
Public Function RecuperaDatosGarantDJMYPE(ByVal psNumGarant As String)
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oRs As New Recordset
     On Error GoTo ErrorRecuperaDatosGarantDJMYPE
     
        sSql = "stp_sel_RecuperaDatosGarantDJMYPE '" & psNumGarant & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set RecuperaDatosGarantDJMYPE = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorRecuperaDatosGarantDJMYPE:
        Err.Raise Err.Number, "Error En Proceso RecuperaDatosGarantDJMYPE", Err.Description
End Function
Public Sub RegistrarNumPoliza(ByVal psCtaCod As String, ByVal psNumPoliza As String)
    Dim lsSQL As String
    lsSQL = "EXEC stp_upd_RegistrarNumPoliza'" & psCtaCod & "','" & psNumPoliza & "'"
    coConex.Ejecutar (lsSQL)
End Sub
'RECO FIN*****************************************************************************************
'RECO20150715 ERS034-2015*************************************
Public Function RecuperaListaGarantMoviliariaPersona(ByVal psPersCod As String, Optional ByVal psCtaCod As String = "") As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
     On Error GoTo ErrorRecuperaListaGarantPolizaPersona
     
        sSql = "stp_sel_RecuperaListaGarantMoviliariaPersona '" & psPersCod & "','" & psCtaCod & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set RecuperaListaGarantMoviliariaPersona = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorRecuperaListaGarantPolizaPersona:
        Err.Raise Err.Number, "Error En Proceso RecuperaListaGarantPolizaPersona", Err.Description
End Function

'RECO FIN*****************************************************
'RECO20160125 ERS001-2016*************************************
Public Function ObtieneDatosGarantRealVenc(ByVal psFecha As String, ByVal psAgeCad As String, ByVal psCadTpoCred As String, ByVal psMoneda As String, ByVal psEstado As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
     On Error GoTo ErrorObtieneDatosGarantRealVenc
     
        sSql = "stp_sel_ObtieneDatosGarantRealVenc '" & psFecha & "','" & psAgeCad & "','" & psCadTpoCred & "','" & psMoneda & "','" & psEstado & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set ObtieneDatosGarantRealVenc = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorObtieneDatosGarantRealVenc:
        Err.Raise Err.Number, "Error En Proceso ObtieneDatosGarantRealVenc", Err.Description
End Function

Public Function GarantRealRegistraCartaVenc(ByVal psCtaCod As String, ByVal psUser As String) As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim lrDataRep As New ADODB.Recordset
    Dim sSql As String
    On Error GoTo ErrorGarantRealRegistraCartaVenc
     
        sSql = "stp_ins_GarantReal_RegistraCartaVenc '" & psCtaCod & "','" & psUser & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set lrDataRep = oConecta.CargaRecordSet(sSql)
        
        If Not (lrDataRep.EOF And lrDataRep.BOF) Then
            GarantRealRegistraCartaVenc = lrDataRep!cNroCarta
        End If
        Set lrDataRep = Nothing
    
        oConecta.CierraConexion
        Set oConecta = Nothing
        Exit Function
ErrorGarantRealRegistraCartaVenc:
        Err.Raise Err.Number, "Error En Proceso GarantRealRegistraCartaVenc", Err.Description
End Function

'RECO END*****************************************************
'EJVG20150326 ***
Public Function ListaCuentasxGarantiaAutoLiq(ByVal psPersCod As String, ByVal psNumGarant As String, Optional ByVal psCtaCod As String = "") As ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "EXEC stp_sel_ERS0632014_ListaCuentasxGarAutoLiq '" & psPersCod & "','" & psNumGarant & "','" & psCtaCod & "'"
    Set ListaCuentasxGarantiaAutoLiq = coConex.CargaRecordSet(lsSQL)
End Function
'Public Function ObtenerCuentaxGarantiaAutoLiq(ByVal psCtaCod As String) As ADODB.Recordset
'    Dim lsSql As String
'    lsSql = "EXEC stp_sel_ERS0632014_ObtenerCuentasxGarAutoLiq '" & psCtaCod & "'"
'    Set ObtenerCuentaxGarantiaAutoLiq = coConex.CargaRecordSet(lsSql)
'End Function
Public Function RegistrarGarantia(ByVal pnMoneda As Moneda, ByVal pnClasificacion As Integer, ByVal pnBienGarantia As Integer, ByVal psPersCod As String, ByVal pnDocTpo As Integer, ByVal psDocNro As String, ByVal pdConstatacion As Date, ByVal psUltimaActualizacion As String, ByVal pnTpoBienContrato As Integer, ByVal pnContratoId As Integer) As String
    Dim rs As New ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "EXEC stp_ins_ERS0632014_Garantia " & pnMoneda & "," & pnClasificacion & "," & pnBienGarantia & ",'" & psPersCod & "'," & pnDocTpo & ",'" & psDocNro & "','" & Format(pdConstatacion, "yyyymmdd") & "','" & psUltimaActualizacion & "'," & pnTpoBienContrato & "," & IIf(pnContratoId > 0, pnContratoId, "NULL")
    Set rs = coConex.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        RegistrarGarantia = rs!cNumGarant
    End If
End Function
Public Sub ActualizarGarantia(ByVal psNumGarant As String, ByVal pnMoneda As Moneda, ByVal pnClasificacion As Integer, ByVal pnBienGarantia As Integer, ByVal psPersCod As String, ByVal pnDocTpo As Integer, ByVal psDocNro As String, ByVal pdConstatacion As Date, ByVal psUltimaActualizacion As String, ByVal pnTpoBienContrato As Integer, ByVal pnContratoId As Integer)
    Dim lsSQL As String
    lsSQL = "EXEC stp_upd_ERS0632014_Garantia '" & psNumGarant & "'," & pnMoneda & "," & pnClasificacion & "," & pnBienGarantia & ",'" & psPersCod & "'," & pnDocTpo & ",'" & psDocNro & "','" & Format(pdConstatacion, "yyyymmdd") & "','" & psUltimaActualizacion & "'," & pnTpoBienContrato & "," & IIf(pnContratoId > 0, pnContratoId, "NULL")
    coConex.Ejecutar lsSQL
End Sub
Public Sub RegistrarGarantiaPersona(ByVal psNumGarant As String, ByVal psPersCod As String, ByVal pnRelacion As Integer)
    Dim lsSQL As String
    lsSQL = "EXEC stp_ins_ERS0632014_GarantiaPersona '" & psNumGarant & "','" & psPersCod & "'," & pnRelacion
    coConex.Ejecutar lsSQL
End Sub
Public Sub EliminarGarantiaPersonas(ByVal psNumGarant As String)
    Dim lsSQL As String
    lsSQL = "EXEC stp_del_ERS0632014_GarantiaPersonas '" & psNumGarant & "'"
    coConex.Ejecutar lsSQL
End Sub
Public Sub RegistrarValorizacion(ByVal psNumGarant As String, ByVal pdRegistro As Date, ByVal psGlosa As String, ByVal psUltimaActualizacion As String)
    Dim lsSQL As String
    lsSQL = "EXEC stp_ins_ERS0632014_GarantiaValoriza '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "','" & psGlosa & "','" & psUltimaActualizacion & "',0"
    coConex.Ejecutar lsSQL
End Sub
Public Sub ActualizarValorizacion(ByVal psNumGarant As String, ByVal pdRegistro As Date, ByVal psGlosa As String, ByVal psUltimaActualizacion As String)
    Dim lsSQL As String
    lsSQL = "EXEC stp_upd_ERS0632014_GarantiaValoriza '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "','" & psGlosa & "','" & psUltimaActualizacion & "'"
    coConex.Ejecutar lsSQL
End Sub
Public Sub EliminarValorizacion(ByVal psNumGarant As String, ByVal pdRegistro As Date)
    Dim lsSQL As String
    lsSQL = "EXEC stp_del_ERS0632014_GarantiaValoriza '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'"
    coConex.Ejecutar lsSQL
End Sub
Public Sub RegistrarValorizacionDirectaTotal(ByVal psNumGarant As String, ByVal pdRegistro As Date, ByRef poTotal As tValorDirectoTotal)
    Dim lsSQL As String
    lsSQL = "EXEC stp_ins_ERS0632014_GarantiaValorizaTot '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'," & poTotal.nVRM
    coConex.Ejecutar lsSQL
End Sub
Public Sub ActualizarValorizacionDirectaTotal(ByVal psNumGarant As String, ByVal pdRegistro As Date, ByRef poTotal As tValorDirectoTotal)
    Dim lsSQL As String
    lsSQL = "EXEC stp_upd_ERS0632014_GarantiaValorizaTot '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'," & poTotal.nVRM
    coConex.Ejecutar lsSQL
End Sub
Public Sub EliminarValorizacionDirectaTotal(ByVal psNumGarant As String, ByVal pdRegistro As Date)
    Dim lsSQL As String
    lsSQL = "EXEC stp_del_ERS0632014_GarantiaValorizaTot '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'"
    coConex.Ejecutar lsSQL
End Sub
Public Sub RegistrarValorizacionDirectaDet(ByVal psNumGarant As String, ByVal pdRegistro As Date, ByVal pnItem As Integer, ByRef poDetalle As tValorDirectoDetallado)
    Dim lsSQL As String
    lsSQL = "EXEC stp_ins_ERS0632014_GarantiaValorizaDet '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'," & pnItem & ",'" & poDetalle.sDescripcion & "'," & poDetalle.nCantidad & "," & poDetalle.nValor & "," & poDetalle.nDocTpo & ",'" & poDetalle.sDocNro & "'"
    coConex.Ejecutar lsSQL
End Sub
Public Sub EliminarValorizacionDirectaDet(ByVal psNumGarant As String, ByVal pdRegistro As Date)
    Dim lsSQL As String
    lsSQL = "EXEC stp_del_ERS0632014_GarantiaValorizaDet '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'"
    coConex.Ejecutar lsSQL
End Sub
Public Sub RegistrarValorizacionAutoLiquidable(ByVal psNumGarant As String, ByVal pdRegistro As Date, ByRef poAutoLiquidable As tValorAutoLiquidable)
    Dim lsSQL As String
    lsSQL = "EXEC stp_ins_ERS0632014_GarantiaValorizaAutoLiq '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "','" & poAutoLiquidable.sCtaCod & "'," & poAutoLiquidable.nFormaRetiro & "," & poAutoLiquidable.nMontoCuenta & "," & poAutoLiquidable.nMontoCuentaDisp & "," & poAutoLiquidable.nMontoDisponible
    coConex.Ejecutar lsSQL
End Sub
Public Sub ActualizarValorizacionAutoLiquidable(ByVal psNumGarant As String, ByVal pdRegistro As Date, ByRef poAutoLiquidable As tValorAutoLiquidable)
    Dim lsSQL As String
    lsSQL = "EXEC stp_upd_ERS0632014_GarantiaValorizaAutoLiq '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "','" & poAutoLiquidable.sCtaCod & "'," & poAutoLiquidable.nFormaRetiro & "," & poAutoLiquidable.nMontoCuenta & "," & poAutoLiquidable.nMontoCuentaDisp & "," & poAutoLiquidable.nMontoDisponible
    coConex.Ejecutar lsSQL
End Sub
Public Sub EliminarValorizacionAutoLiquidable(ByVal psNumGarant As String, ByVal pdRegistro As Date)
    Dim lsSQL As String
    lsSQL = "EXEC stp_del_ERS0632014_GarantiaValorizaAutoLiq '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'"
    coConex.Ejecutar lsSQL
End Sub
Public Sub RegistrarValorizacionTasacInmobiliaria(ByVal psNumGarant As String, ByVal pdRegistro As Date, ByRef poInmueble As tValorTasacionInmobiliaria)
    Dim lsSQL As String
    lsSQL = "EXEC stp_ins_ERS0632014_GarantiaValorizaTasacInmobil '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "','" & poInmueble.sUbicacionGeografica & "','" & poInmueble.sDireccion & "'," & poInmueble.nValorComercial & "," & poInmueble.nValorEdificacion & "," & poInmueble.nVRM & "," & poInmueble.nClase & "," & poInmueble.nCategoria & "," & poInmueble.nNroLocales & "," & poInmueble.nNroPisos & "," & poInmueble.nNroSotanos & "," & poInmueble.nAnioConstruccion & ",'" & Format(poInmueble.dTasacion, "yyyymmdd") & "'," & poInmueble.nTasacionTC & ",'" & poInmueble.sTasadorCod & "','" & poInmueble.sMz & "','" & poInmueble.sLt & "','" & poInmueble.Set & "'"
    coConex.Ejecutar lsSQL
End Sub
Public Sub ActualizarValorizacionTasacInmobiliaria(ByVal psNumGarant As String, ByVal pdRegistro As Date, ByRef poInmueble As tValorTasacionInmobiliaria)
    Dim lsSQL As String
    lsSQL = "EXEC stp_upd_ERS0632014_GarantiaValorizaTasacInmobil '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "','" & poInmueble.sUbicacionGeografica & "','" & poInmueble.sDireccion & "'," & poInmueble.nValorComercial & "," & poInmueble.nValorEdificacion & "," & poInmueble.nVRM & "," & poInmueble.nClase & "," & poInmueble.nCategoria & "," & poInmueble.nNroLocales & "," & poInmueble.nNroPisos & "," & poInmueble.nNroSotanos & "," & poInmueble.nAnioConstruccion & ",'" & Format(poInmueble.dTasacion, "yyyymmdd") & "'," & poInmueble.nTasacionTC & ",'" & poInmueble.sTasadorCod & "','" & poInmueble.sMz & "','" & poInmueble.sLt & "','" & poInmueble.Set & "'"
    coConex.Ejecutar lsSQL
End Sub
Public Sub EliminarValorizacionTasacInmobiliaria(ByVal psNumGarant As String, ByVal pdRegistro As Date)
    Dim lsSQL As String
    lsSQL = "EXEC stp_del_ERS0632014_GarantiaValorizaTasacInmobil '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'"
    coConex.Ejecutar lsSQL
End Sub
Public Sub RegistrarValorizacionTasacVehicular(ByVal psNumGarant As String, ByVal pdRegistro As Date, ByRef poVehiculo As tValorTasacionVehicular)
    Dim lsSQL As String
    lsSQL = "EXEC stp_ins_ERS0632014_GarantiaValorizaTasacVehicular '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "','" & poVehiculo.sDescripcion & "','" & poVehiculo.sCarroceria & "','" & poVehiculo.sMarca & "','" & poVehiculo.sModelo & "','" & poVehiculo.sNroPlaca & "','" & poVehiculo.sNroMotor & "','" & poVehiculo.sNroChasis & "'," & poVehiculo.nAnioFabricacion & "," & poVehiculo.nAnioAdquisicionTienda & "," & poVehiculo.nValorComercial & "," & poVehiculo.nVRM & IIf(poVehiculo.sTasadorCod <> "", ",'" & Format(poVehiculo.dTasacion, "yyyymmdd") & "'," & poVehiculo.nTasacionTC & ",'" & poVehiculo.sTasadorCod & "'", ",NULL,NULL,NULL")
    coConex.Ejecutar lsSQL
End Sub
Public Sub ActualizarValorizacionTasacVehicular(ByVal psNumGarant As String, ByVal pdRegistro As Date, ByRef poVehiculo As tValorTasacionVehicular)
    Dim lsSQL As String
    lsSQL = "EXEC stp_upd_ERS0632014_GarantiaValorizaTasacVehicular '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "','" & poVehiculo.sDescripcion & "','" & poVehiculo.sCarroceria & "','" & poVehiculo.sMarca & "','" & poVehiculo.sModelo & "','" & poVehiculo.sNroPlaca & "','" & poVehiculo.sNroMotor & "','" & poVehiculo.sNroChasis & "'," & poVehiculo.nAnioFabricacion & "," & poVehiculo.nAnioAdquisicionTienda & "," & poVehiculo.nValorComercial & "," & poVehiculo.nVRM & IIf(poVehiculo.sTasadorCod <> "", ",'" & Format(poVehiculo.dTasacion, "yyyymmdd") & "'," & poVehiculo.nTasacionTC & ",'" & poVehiculo.sTasadorCod & "'", ",NULL,NULL,NULL")
    coConex.Ejecutar lsSQL
End Sub
Public Sub EliminarValorizacionTasacVehicular(ByVal psNumGarant As String, ByVal pdRegistro As Date)
    Dim lsSQL As String
    lsSQL = "EXEC stp_del_ERS0632014_GarantiaValorizaTasacVehicular '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'"
    coConex.Ejecutar lsSQL
End Sub
Public Sub RegistrarValorizacionOtrasTasacMobiliarias(ByVal psNumGarant As String, ByVal pdRegistro As Date, ByRef poOtras As tValorTasacionMobiliariaOtras)
    Dim lsSQL As String
    lsSQL = "EXEC stp_ins_ERS0632014_GarantiaValorizaTasacMobilOtras '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "','" & poOtras.sDescripcion & "','" & poOtras.sNroSerie & "','" & poOtras.sMarca & "','" & poOtras.sModelo & "'," & poOtras.nAnioFabricacion & "," & poOtras.nValorComercial & "," & poOtras.nVRM & ",'" & Format(poOtras.dTasacion, "yyyymmdd") & "'," & poOtras.nTasacionTC & ",'" & poOtras.sTasadorCod & "'"
    coConex.Ejecutar lsSQL
End Sub
Public Sub ActualizarValorizacionOtrasTasacMobiliarias(ByVal psNumGarant As String, ByVal pdRegistro As Date, ByRef poOtras As tValorTasacionMobiliariaOtras)
    Dim lsSQL As String
    lsSQL = "EXEC stp_upd_ERS0632014_GarantiaValorizaTasacMobilOtras '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "','" & poOtras.sDescripcion & "','" & poOtras.sNroSerie & "','" & poOtras.sMarca & "','" & poOtras.sModelo & "'," & poOtras.nAnioFabricacion & "," & poOtras.nValorComercial & "," & poOtras.nVRM & ",'" & Format(poOtras.dTasacion, "yyyymmdd") & "'," & poOtras.nTasacionTC & ",'" & poOtras.sTasadorCod & "'"
    coConex.Ejecutar lsSQL
End Sub
Public Sub EliminarValorizacionOtrasTasacMobiliarias(ByVal psNumGarant As String, ByVal pdRegistro As Date)
    Dim lsSQL As String
    lsSQL = "EXEC stp_del_ERS0632014_GarantiaValorizaTasacMobilOtras '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'"
    coConex.Ejecutar lsSQL
End Sub
'EJVG20160226 ***
Public Sub RegistrarValorizacionGravamenOtraIFi(ByVal psNumGarant As String, ByVal pdRegistro As Date, ByRef poIFi As tValorGravamenFavorOtraIFi)
    Dim lsSQL As String
    Dim i As Integer
    lsSQL = "EXEC stp_ins_ERS0022016_GarantiaValorizaGravamenOtraIFi '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "','" & poIFi.sDescripcion & "'," & poIFi.nValorComercial
    coConex.Ejecutar lsSQL
    For i = 1 To UBound(poIFi.vValorGravamenFavorOtraIFiDet)
        If poIFi.nOpTpRegPub = 1 Then 'JOEP20180524 ERS009-2018
            RegistrarValorizacionGravamenOtraIFiDet psNumGarant, pdRegistro, i, poIFi.vValorGravamenFavorOtraIFiDet(i)
        ElseIf poIFi.nOpTpRegPub = 2 Then 'JOEP20180524 ERS009-2018
            RegistrarValorizacionGravamenPerNatJudDet psNumGarant, pdRegistro, i, poIFi.vValorGravamenFavorOtraIFiDet(i) 'JOEP20180524 ERS009-2018
        End If 'JOEP20180524 ERS009-2018
    Next
End Sub
'JOEP20180524 ERS009-2018
Public Sub RegistrarValorizacionGravamenPerNatJudDet(ByVal psNumGarant As String, ByVal pdRegistro As Date, ByVal pnItem As Integer, ByRef poIFiDet As tValorGravamenFavorOtraIFiDet)
    Dim lsSQL As String
    lsSQL = "EXEC stp_ins_ERS0092018_GarantiaValorizaGravamenPerNatJudDet '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'," & pnItem & ",'" & poIFiDet.sIFiCod & "'," & poIFiDet.nGravado
    coConex.Ejecutar lsSQL
End Sub
Public Sub EliminarValorizacionGravamenPerNatJudDet(ByVal psNumGarant As String, ByVal pdRegistro As Date)
    Dim lsSQL As String
    lsSQL = "EXEC stp_del_ERS0092018_GarantiaValorizaGravamenPerNatJudDet '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'"
    coConex.Ejecutar lsSQL
End Sub
'JOEP20180524 ERS009-2018
Public Sub RegistrarValorizacionGravamenOtraIFiDet(ByVal psNumGarant As String, ByVal pdRegistro As Date, ByVal pnItem As Integer, ByRef poIFiDet As tValorGravamenFavorOtraIFiDet)
    Dim lsSQL As String
    lsSQL = "EXEC stp_ins_ERS0022016_GarantiaValorizaGravamenOtraIFiDet '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'," & pnItem & ",'" & poIFiDet.sIFiCod & "','" & poIFiDet.sIFTpo & "'," & poIFiDet.nGravado
    coConex.Ejecutar lsSQL
End Sub
Public Sub EliminarValorizacionGravamenOtraIFi(ByVal psNumGarant As String, ByVal pdRegistro As Date)
    Dim lsSQL As String
    lsSQL = "EXEC stp_del_ERS0022016_GarantiaValorizaGravamenOtraIFi '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'"
    coConex.Ejecutar lsSQL
End Sub
'END EJVG *******
Public Function RecuperaGarantiasxPersona(ByVal psPersCod As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaGarantiasxPersona
    lsSQL = "EXEC stp_sel_ERS0632014_ListaGarantiasxPersona '" & psPersCod & "'"
    Set RecuperaGarantiasxPersona = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaGarantiasxPersona:
    Err.Raise Err.Number, "Error al recuperar las Garantías de una Persona", Err.Description
End Function
Public Function RecuperaGarantiaxID(ByVal psNumGarant As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaGarantiaxID
    lsSQL = "EXEC stp_sel_ERS0632014_GarantiaxID '" & psNumGarant & "'"
    Set RecuperaGarantiaxID = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaGarantiaxID:
    Err.Raise Err.Number, "Error al recuperar la Garantia con identificador " & psNumGarant, Err.Description
End Function
Public Function RecuperaGarantiaPersonaxID(ByVal psNumGarant As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaGarantiaPersonaxID
    lsSQL = "EXEC stp_sel_ERS0632014_GarantiaPersonaxID '" & psNumGarant & "'"
    Set RecuperaGarantiaPersonaxID = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaGarantiaPersonaxID:
    Err.Raise Err.Number, "Error al recuperar los Propietarios de la Garantia con identificador " & psNumGarant, Err.Description
End Function
Public Function RecuperaGarantiaValorizaxID(ByVal psNumGarant As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaGarantiaPersonaxID
    lsSQL = "EXEC stp_sel_ERS0632014_GarantiaValorizaxID '" & psNumGarant & "'"
    Set RecuperaGarantiaValorizaxID = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaGarantiaPersonaxID:
    Err.Raise Err.Number, "Error al recuperar las Valorizaciones de la Garantia con identificador " & psNumGarant, Err.Description
End Function
Public Function RecuperaGarantiaValorizaxTipo(ByVal psNumGarant As String, ByVal pdRegistro As Date, ByVal pnTpoValoriza As eGarantiaTipoValorizacion) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaGarantiaPersonaxID
    
    If pnTpoValoriza = ValorDirectoTotal Then
        lsSQL = "EXEC stp_sel_ERS0632014_GarantiaValorizaTot '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'"
    ElseIf pnTpoValoriza = ValorDirectoDetallado Then
        lsSQL = "EXEC stp_sel_ERS0632014_GarantiaValorizaDet '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'"
    ElseIf pnTpoValoriza = GarantiaAutoliquidable Then
        lsSQL = "EXEC stp_sel_ERS0632014_GarantiaValorizaAutoLiquidable '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'"
    ElseIf pnTpoValoriza = TasacionInmobiliaria Then
        lsSQL = "EXEC stp_sel_ERS0632014_GarantiaValorizaTasacInmobiliaria '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'"
    ElseIf pnTpoValoriza = TasacionVehicular Then
        lsSQL = "EXEC stp_sel_ERS0632014_GarantiaValorizaTasacVehicular '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'"
    ElseIf pnTpoValoriza = OtrasTasacionesMobiliarias Then
        lsSQL = "EXEC stp_sel_ERS0632014_GarantiaValorizaTasacInmobiliariaOtras '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'"
    ElseIf pnTpoValoriza = GravamenFavorOtraIFi Then 'EJVG20160227
        lsSQL = "EXEC stp_sel_ERS0632014_GarantiaValorizaGravamenOtraIFi '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'"
    End If
    
    Set RecuperaGarantiaValorizaxTipo = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaGarantiaPersonaxID:
    Err.Raise Err.Number, "Error al recuperar la Valorización con identificador " & psNumGarant & "y Fecha '" & Format(pdRegistro, "dd/mm/yyyy hh:mm:ss") & "'", Err.Description
End Function
Public Function ExisteGarantiaNew(ByVal psPersCodEmisor As String, ByVal pnDocTpo As Integer, ByVal psDocNro As String, ByVal psNumGarant As String, ByRef psNumGarantExiste As String) As Boolean
    Dim rsGarant As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrExisteGarantia
    
    lsSQL = "EXEC stp_sel_ERS0632014_ExisteGarantia '" & psPersCodEmisor & "'," & pnDocTpo & ",'" & psDocNro & "','" & psNumGarant & "'"
        
    Set rsGarant = coConex.CargaRecordSet(lsSQL)
    If Not rsGarant.EOF Then
        psNumGarantExiste = rsGarant!cNumGarant
        ExisteGarantiaNew = True
    End If
    Exit Function
ErrExisteGarantia:
    Err.Raise Err.Number, "Error al verificar existencia de garantía", Err.Description
End Function
Public Function RecuperaZonaRegistral() As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaZonaRegistral
    lsSQL = "EXEC stp_sel_ERS0632014_ZonaRegistral"
    Set RecuperaZonaRegistral = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaZonaRegistral:
    Err.Raise Err.Number, "Error al recuperar las Zonas Registrales", Err.Description
End Function
'Public Function RecuperaOficinaZonaRegistral() As ADODB.Recordset
'    Dim lsSql As String
'    On Error GoTo ErrRecuperaOficinaZonaRegistral
'    lsSql = "EXEC stp_sel_ERS0632014_ListaOficinaZonaRegistral "
'    Set RecuperaOficinaZonaRegistral = coConex.CargaRecordSet(lsSql)
'    Exit Function
'ErrRecuperaOficinaZonaRegistral:
'    Err.Raise Err.Number, "Error al recuperar las Oficinas-Zonas Registrales", Err.Description
'End Function
Public Function RecuperaOficinaRegistralxZona(ByVal pnZonaID As Integer) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaOficinaRegistralxZona
    lsSQL = "EXEC stp_sel_ERS0632014_OficinaRegistralxZona " & pnZonaID
    Set RecuperaOficinaRegistralxZona = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaOficinaRegistralxZona:
    Err.Raise Err.Number, "Error al recuperar las Oficinas de las Zonas Registrales", Err.Description
End Function
Public Function RecuperaProxTramiteLegal(ByVal psNumGarant As String, ByVal pdTramiteLegal_ULT As Date, ByVal pnAccion As Integer) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaProxTramiteLegal
    lsSQL = "EXEC stp_sel_ERS0632014_ListaTipoTramite '" & psNumGarant & "','" & Format(pdTramiteLegal_ULT, "yyyymmdd hh:mm:ss") & "'," & pnAccion
    Set RecuperaProxTramiteLegal = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaProxTramiteLegal:
    Err.Raise Err.Number, "Error al recuperar los Próximos Tramites Legales según actual situación Garantía", Err.Description
End Function
Public Function ExisteNroAsientoGarantia(ByVal psNumGarant As String, ByVal pdFecha As Date, ByVal psNroAsiento As String) As Boolean
    Dim rsGarant As New ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaProxTramiteLegal
    lsSQL = "SELECT bExiste = dbo.fnc_ERS0632014_ExisteNroAsientoGarantia('" & psNumGarant & "','" & Format(pdFecha, "yyyymmdd hh:mm:ss") & "','" & psNroAsiento & "')"
    Set rsGarant = coConex.CargaRecordSet(lsSQL)
    If Not rsGarant.EOF Then
        ExisteNroAsientoGarantia = rsGarant!bExiste
    End If
    Exit Function
ErrRecuperaProxTramiteLegal:
    Err.Raise Err.Number, "Error al recuperar existencia de Nro. de Asiento de Tramite Legal de Garantía", Err.Description
End Function
Public Sub RegistrarTramiteLegal(ByVal psNumGarant As String, ByRef pvTramiteLegal As tTramiteLegal)
    Dim lsSQL As String
    lsSQL = "EXEC stp_ins_ERS0632014_TramiteLegal '" & psNumGarant & "','" & Format(pvTramiteLegal.dFecha, "yyyymmdd hh:mm:ss") & "','" _
            & pvTramiteLegal.sNroPartidaRegistral & "'," & pvTramiteLegal.nTipoTramite & "," & pvTramiteLegal.nGravamen & "," _
            & pvTramiteLegal.nZonaRegistralID & "," & pvTramiteLegal.nOficinaRegistralID & ",'" & pvTramiteLegal.sNotariaCod & "'," & pvTramiteLegal.nEstado & "," _
            & IIf(pvTramiteLegal.nTipoTramite = BloqueoRegistral, "NULL", pvTramiteLegal.nTipoCobertura) & "," _
            & IIf(pvTramiteLegal.nEstado = Pendiente, "NULL", "'" & pvTramiteLegal.sNroAsiento & "'") & "," _
            & IIf(pvTramiteLegal.nEstado = Pendiente, "NULL", "'" & Format(pvTramiteLegal.dInscripcion, "yyyymmdd hh:mm:ss") & "'") & ",'" _
            & pvTramiteLegal.sUltimaActualizacion & "',0"
    
    coConex.Ejecutar lsSQL
End Sub
Public Sub ActualizarTramiteLegal(ByVal psNumGarant As String, ByRef pvTramiteLegal As tTramiteLegal)
    Dim lsSQL As String
    lsSQL = "EXEC stp_upd_ERS0632014_TramiteLegal '" & psNumGarant & "','" & Format(pvTramiteLegal.dFecha, "yyyymmdd hh:mm:ss") & "','" _
            & pvTramiteLegal.sNroPartidaRegistral & "'," & pvTramiteLegal.nTipoTramite & "," & pvTramiteLegal.nGravamen & "," _
            & pvTramiteLegal.nZonaRegistralID & "," & pvTramiteLegal.nOficinaRegistralID & ",'" & pvTramiteLegal.sNotariaCod & "'," & pvTramiteLegal.nEstado & "," _
            & IIf(pvTramiteLegal.nTipoTramite = BloqueoRegistral, "NULL", pvTramiteLegal.nTipoCobertura) & "," _
            & IIf(pvTramiteLegal.nEstado = Pendiente, "NULL", "'" & pvTramiteLegal.sNroAsiento & "'") & "," _
            & IIf(pvTramiteLegal.nEstado = Pendiente, "NULL", "'" & Format(pvTramiteLegal.dInscripcion, "yyyymmdd hh:mm:ss") & "'") & ",'" _
            & pvTramiteLegal.sUltimaActualizacion & "'"
    
    coConex.Ejecutar lsSQL
End Sub
Public Sub EliminarTramiteLegal(ByVal psNumGarant As String, ByRef pdFecha As Date)
    Dim lsSQL As String
    lsSQL = "EXEC stp_del_ERS0632014_TramiteLegal '" & psNumGarant & "','" & Format(pdFecha, "yyyymmdd hh:mm:ss") & "'"
    coConex.Ejecutar lsSQL
End Sub
Public Function RecuperaTramiteLegalxID(ByVal psNumGarant As String, ByVal pbUltVinculadoVAL As Boolean, ByVal pbPuedeAgregarVAL As Boolean) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaTramiteLegalxID
    lsSQL = "EXEC stp_sel_ERS0632014_GarantiaTramiteLegalxID '" & psNumGarant & "'," & IIf(pbUltVinculadoVAL, 1, 0) & "," & IIf(pbPuedeAgregarVAL, 1, 0)
    Set RecuperaTramiteLegalxID = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaTramiteLegalxID:
    Err.Raise Err.Number, "Error al recuperar los Trámites Legales de la Garantia con identificador " & psNumGarant, Err.Description
End Function
Public Function RecuperaProductoCoberturaConfig() As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaProductoCoberturaConfig
    lsSQL = "EXEC stp_sel_ERS0632014_ListaProductoCoberturaConfig"
    Set RecuperaProductoCoberturaConfig = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaProductoCoberturaConfig:
    Err.Raise Err.Number, "Error al recuperar el listado de productos para configuración de coberturas", Err.Description
End Function
Public Function RecuperaCoberturaConfig(ByVal psTpoProdCod As String, ByVal pbCliPreferencial As Boolean) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaCoberturaConfig
    lsSQL = "EXEC stp_sel_ERS0632014_CoberturaConfig '" & psTpoProdCod & "'," & IIf(pbCliPreferencial, 1, 0)
    Set RecuperaCoberturaConfig = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaCoberturaConfig:
    Err.Raise Err.Number, "Error al recuperar las coberturas ya antes configuradas", Err.Description
End Function
Public Sub EliminarCoberturaConfig(ByVal psTpoProdCod As String, ByVal pbCliPreferencial As Boolean)
    Dim lsSQL As String
    On Error GoTo ErrEliminarCoberturaConfig
    lsSQL = "EXEC stp_del_ERS0632014_CoberturaConfig '" & psTpoProdCod & "'," & IIf(pbCliPreferencial, 1, 0)
    coConex.Ejecutar (lsSQL)
    Exit Sub
ErrEliminarCoberturaConfig:
    Err.Raise Err.Number, "Error al eliminar las coberturas anteriores", Err.Description
End Sub
Public Sub RegistrarCoberturaConfig(ByVal psTpoProdCod As String, ByVal pbCliPreferencial As Boolean, ByVal pnBienGarantia As Integer, ByVal pnCoberturaPreferida As Currency, ByVal pnCoberturaNoPreferida As Currency)
    Dim lsSQL As String
    On Error GoTo ErrEliminarCoberturaConfig
    lsSQL = "EXEC stp_ins_ERS0632014_CoberturaConfig '" & psTpoProdCod & "'," & IIf(pbCliPreferencial, 1, 0) & "," & pnBienGarantia & "," & pnCoberturaPreferida & "," & pnCoberturaNoPreferida
    coConex.Ejecutar (lsSQL)
    Exit Sub
ErrEliminarCoberturaConfig:
    Err.Raise Err.Number, "Error al registrar las coberturas especificadas", Err.Description
End Sub
Public Function RecuperaCoberturaConfigRapiFlash() As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaCoberturaConfigEx
    lsSQL = "EXEC stp_sel_ERS0632014_CoberturaConfigRapiFlash"
    Set RecuperaCoberturaConfigRapiFlash = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaCoberturaConfigEx:
    Err.Raise Err.Number, "Error al recuperar la configuración de cobertura para RapiFlash", Err.Description
End Function
Public Function RecuperaCoberturaConfigRapiFlashCamp() As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaCoberturaConfigEx
    lsSQL = "EXEC stp_sel_ERS0632014_CoberturaConfigRapiFlashCamp"
    Set RecuperaCoberturaConfigRapiFlashCamp = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaCoberturaConfigEx:
    Err.Raise Err.Number, "Error al recuperar la configuración de cobertura para RapiFlash con Campaña", Err.Description
End Function
Public Sub ActualizarCoberturaConfig(ByVal pnRetiroTpo As Integer, ByVal pnRetiroMonto As Currency)
    Dim lsSQL As String
    On Error GoTo ErrActualizarCoberturaConfig
    lsSQL = "EXEC stp_upd_ERS0632014_CoberturaConfigEx " & pnRetiroTpo & "," & pnRetiroMonto
    coConex.Ejecutar lsSQL
    Exit Sub
ErrActualizarCoberturaConfig:
    Err.Raise Err.Number, "Error al actualizar las excepciones de la configuración de coberturas", Err.Description
End Sub
'Cambiar a DCOMCredito
Public Function RecuperaPersonaCreditoxGravamen(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaPersonaCreditoxGravamen
    lsSQL = "EXEC stp_sel_ERS0632014_PersonaCredito '" & psCtaCod & "'"
    Set RecuperaPersonaCreditoxGravamen = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaPersonaCreditoxGravamen:
    Err.Raise Err.Number, "Error al recuperar las personas intervinientes en el crédito", Err.Description
End Function
'Cambiar a DCOMCredito
Public Function RecuperaDatosCreditoxGravamen(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaDatosCreditoxGravamen
    lsSQL = "EXEC stp_sel_ERS0632014_DatosCredxGravamen '" & psCtaCod & "'"
    
    Set RecuperaDatosCreditoxGravamen = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaDatosCreditoxGravamen:
    Err.Raise Err.Number, "Error al recuperar datos del crédito solicitado", Err.Description
End Function
'Cambiar a DCOMCredito
Public Function RecuperaDatosCFxGravamen(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaDatosCFxGravamen
    lsSQL = "EXEC stp_sel_ERS0632014_DatosCFxGravamen '" & psCtaCod & "'"
    
    Set RecuperaDatosCFxGravamen = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaDatosCFxGravamen:
    Err.Raise Err.Number, "Error al recuperar datos de la Carta Fianza solicitada", Err.Description
End Function
'Cambiar a DCOMCredito
Public Function RecuperaGarantiasPersonaxGravamen(ByVal psCtaCod As String, ByVal psPersCodLista As String, ByVal pbCliPreferencial As Boolean, ByVal psTpoProdCod As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaGarantiasPersonaxGravamen
    lsSQL = "EXEC stp_sel_ERS0632014_GarantiasPersona '" & psCtaCod & "','" & psPersCodLista & "'," & IIf(pbCliPreferencial, 1, 0) & ",'" & psTpoProdCod & "'"
    Set RecuperaGarantiasPersonaxGravamen = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaGarantiasPersonaxGravamen:
    Err.Raise Err.Number, "Error al recuperar las Garantías de las Personas intervinientes del crédito", Err.Description
End Function
'Cambiar a DCOMCredito
Public Function RecuperaGarantiasCreditoxGravamen(ByVal psCtaCod As String, Optional ByVal psNumGarantLista As String = "") As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaGarantiasCreditoxGravamen
    lsSQL = "EXEC stp_sel_ERS0632014_GarantiasCredito '" & psCtaCod & "','" & psNumGarantLista & "'"
    Set RecuperaGarantiasCreditoxGravamen = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaGarantiasCreditoxGravamen:
    Err.Raise Err.Number, "Error al recuperar las Garantías del préstamo " & psCtaCod, Err.Description
End Function
Public Function RecuperaGarantiasCreditoAnterior(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaGarantiasCreditoAnterior
    lsSQL = "EXEC stp_sel_ERS0632014_GarantiaCreditoAnterior '" & psCtaCod & "'"
    Set RecuperaGarantiasCreditoAnterior = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaGarantiasCreditoAnterior:
    Err.Raise Err.Number, "Error al recuperar las Garantías de los créditos anteriores" & psCtaCod, Err.Description
End Function
Public Function RecuperaGarantiasCreditoxAjusteGravamen(ByVal psCtaCod As String, ByVal pbCliPreferencial As Boolean, ByVal psTpoProdCod As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaGarantiasCreditoxAjusteGravamen
    lsSQL = "EXEC stp_sel_ERS0632014_GarantiasCreditoxAjuste '" & psCtaCod & "'," & pbCliPreferencial & ",'" & psTpoProdCod & "'"
    Set RecuperaGarantiasCreditoxAjusteGravamen = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaGarantiasCreditoxAjusteGravamen:
    Err.Raise Err.Number, "Error al recuperar las Garantías del préstamo " & psCtaCod, Err.Description
End Function
Public Function RecuperaColocGravamen(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaColocGravamen
    lsSQL = "EXEC stp_sel_ERS0632014_ColocGravamen '" & psCtaCod & "'"
    Set RecuperaColocGravamen = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaColocGravamen:
    Err.Raise Err.Number, "Error al recuperar las Garantías del préstamo " & psCtaCod, Err.Description
End Function
'Public Sub RegistrarColocGravamen(ByVal psCtaCod As String, ByVal psTpoProdCod As String, ByVal pnMontoColocado As Currency)
'    Dim lsSql As String
'    On Error GoTo ErrRegistrarColocGravamen
'    lsSql = "EXEC stp_ins_ERS0632014_ColocGravamen '" & psCtaCod & "','" & psTpoProdCod & "'," & pnMontoColocado
'    coConex.Ejecutar lsSql
'    Exit Sub
'ErrRegistrarColocGravamen:
'    Err.Raise Err.Number, "Error al registrar datos del gravamen para el producto con código " & psCtaCod, Err.Description
'End Sub
'Public Sub RegistrarCoberturaGarantia(ByVal psNumGarant As String, ByVal pnMoneda As Integer, ByVal psCtaCod As String, ByVal pdValorizacion As Date, _
'            ByVal pnSaldoCobertura As Currency, ByVal pnRatio As Currency, ByVal pnCoberturaMax As Currency, ByVal pnCobertura As Currency, _
'            Optional ByVal pdTramiteLegal As Date = CDate("1900-01-01"))
'    Dim lsSql As String
'    On Error GoTo ErrRegistrarCoberturaGarantia
'    lsSql = "EXEC stp_ins_ERS0632014_ColocGarantia '" & psNumGarant & "'," & pnMoneda & ",'" & psCtaCod & "','" & Format(pdValorizacion, "yyyymmdd hh:mm:ss") & "'," _
'            & IIf(DateDiff("d", pdTramiteLegal, "1900-01-01") = 0, "NULL", "'" & Format(pdTramiteLegal, "yyyymmdd hh:mm:ss") & "'") & "," & pnSaldoCobertura & "," _
'            & pnRatio & "," & pnCoberturaMax & "," & pnCobertura
'    coConex.Ejecutar lsSql
'    Exit Sub
'ErrRegistrarCoberturaGarantia:
'    Err.Raise Err.Number, "Error al registrar la cobertura para el producto con código " & psCtaCod, Err.Description
'End Sub
'Public Sub EliminarCoberturaGarantia(ByVal psCtaCod As String)
'    Dim lsSql As String
'    On Error GoTo ErrEliminarCoberturaGarantia
'    lsSql = "EXEC stp_del_ERS0632014_ColocGarantia '" & psCtaCod & "'"
'    coConex.Ejecutar (lsSql)
'    Exit Sub
'ErrEliminarCoberturaGarantia:
'    Err.Raise Err.Number, "Error al eliminar las coberturas relaciondas al producto con código " & psCtaCod, Err.Description
'End Sub
Public Function CambiosAfectanGravamen(ByVal psCtaCod As String, ByVal psTpoProdCodNEW As String, ByVal pnMonto As Currency) As Boolean
    Dim lsSQL As String
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrCambiosAfectanGravamen
    lsSQL = "SELECT bAfecta = dbo.fnc_ERS0632014_CambiosAfectanGravamen('" & psCtaCod & "','" & psTpoProdCodNEW & "'," & pnMonto & ")"
    Set rs = coConex.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        CambiosAfectanGravamen = rs!bAfecta
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrCambiosAfectanGravamen:
    Err.Raise Err.Number, "Error al verificar si cambios afectan Gravamen", Err.Description
End Function
Public Function RecuperaSolicitudPoliza(ByVal psAgeCod As String, Optional ByVal psCtaCod As String = "") As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaSolicitudPoliza
    'lsSql = "EXEC stp_sel_ERS0632014_SolicitudPoliza '" & psAgeCod & "','" & psCtaCod & "'"
    lsSQL = "SELECT * FROM DBO.fnc_ERS0632014_SolicitudPoliza('" & psAgeCod & "','" & psCtaCod & "')"
    Set RecuperaSolicitudPoliza = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaSolicitudPoliza:
    Err.Raise Err.Number, "Error al recuperar los créditos por solicitud de poliza", Err.Description
End Function
Public Function RecuperaGarantiasxSolicitudPoliza(ByVal psCtaCod As String, ByVal psTpoCredCod As String, ByVal pnMonto As Currency, ByVal pdVencimiento As Date) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaGarantiasxSolicitudPoliza
    lsSQL = "EXEC stp_sel_ERS0632014_GarantiasxSolicitudPoliza '" & psCtaCod & "','" & psTpoCredCod & "'," & pnMonto & ",'" & Format(pdVencimiento, "yyyymmdd") & "'"
    Set RecuperaGarantiasxSolicitudPoliza = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaGarantiasxSolicitudPoliza:
    Err.Raise Err.Number, "Error al recuperar las garantías de un crédito para registro de Poliza", Err.Description
End Function
Public Function RecuperaInmueblexSolicitudPoliza(ByVal psNumGarant As String, ByVal pdValorizacion As Date) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaInmueblexSolicitudPoliza
    lsSQL = "EXEC stp_sel_ERS0632014_InmueblexSolicitudPoliza '" & psNumGarant & "','" & Format(pdValorizacion, "yyyymmdd hh:mm:ss") & "'"
    Set RecuperaInmueblexSolicitudPoliza = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaInmueblexSolicitudPoliza:
    Err.Raise Err.Number, "Error al recuperar el Inmueble para poder modificarlo", Err.Description
End Function
Public Sub ActualizarGarantiaInmobiliaria(ByVal psNumGarant As String, ByVal pdRegistro As Date, ByVal pnClase As Integer, _
                                                ByVal pnNroLocales As Integer, ByVal pnNroPisos As Integer, ByVal pnCategoria As Integer, _
                                                ByVal pnNroSotanos As Integer, ByVal pnAnioConstruccion As Integer)
    Dim lsSQL As String
    On Error GoTo ErrActualizarCoberturaConfig
    lsSQL = "EXEC stp_upd_ERS0632014_InmueblexSolicitudPoliza '" & psNumGarant & "','" & Format(pdRegistro, "yyyymmdd hh:mm:ss") & "'," & pnClase & "," & pnNroLocales & "," & pnNroPisos & "," & pnCategoria & "," & pnNroSotanos & "," & pnAnioConstruccion
    coConex.Ejecutar lsSQL
    Exit Sub
ErrActualizarCoberturaConfig:
    Err.Raise Err.Number, "Error al actualizar las excepciones de la configuración de coberturas", Err.Description
End Sub
Public Function RecuperaTipoPoliza(ByVal pbExterna As Boolean, ByVal pbTasacInmobil As Boolean, ByVal pbTasacMobil As Boolean) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaTipoPoliza
    lsSQL = "EXEC stp_sel_ERS0632014_TipoPrimaPoliza " & IIf(pbExterna, 1, 0) & "," & IIf(pbTasacInmobil, 1, 0) & "," & IIf(pbTasacMobil, 1, 0)
    Set RecuperaTipoPoliza = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaTipoPoliza:
    Err.Raise Err.Number, "Error al recuperar los tipos de Poliza", Err.Description
End Function
Public Function RecuperaUltPolizaxGarantia(ByVal psNumGarant As String, ByVal psCtaCod As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaUltPolizaxGarantia
    lsSQL = "SELECT * FROM DBO.fnc_ERS0632014_GarantiaPolizaVence('" & psNumGarant & "','" & psCtaCod & "')"
    Set RecuperaUltPolizaxGarantia = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaUltPolizaxGarantia:
    Err.Raise Err.Number, "Error al recuperar última Poliza por Garantía", Err.Description
End Function
Public Function RecuperaDatosGarantiaxCalculoPoliza(ByVal psNumGarant As String, ByVal pdValorizacion As Date, ByVal pnMonedaGar As Moneda, ByVal pnMonedaCred As Integer) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaDatosGarantiaxCalculoPoliza
    lsSQL = "EXEC stp_sel_ERS0632014_DatosGarantiaxCalculoPoliza '" & psNumGarant & "','" & Format(pdValorizacion, "yyyymmdd hh:mm:ss") & "'," & pnMonedaGar & "," & pnMonedaCred
    Set RecuperaDatosGarantiaxCalculoPoliza = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaDatosGarantiaxCalculoPoliza:
    Err.Raise Err.Number, "Error al recuperar datos para el calculo de Poliza para Garantía Inmobiliaria/Mobiliaria/Otras Mobiliareas", Err.Description
End Function
Public Function RecuperaPoliza(ByVal psNumPoliza As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaPoliza
    lsSQL = "EXEC stp_sel_ERS0632014_PolizaxID '" & psNumPoliza & "'"
    Set RecuperaPoliza = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaPoliza:
    Err.Raise Err.Number, "Error al recuperar datos de la Póliza N° " & psNumPoliza, Err.Description
End Function
Public Function DameAlternativaPolizaMobiliaria(ByVal pnTpoPoliza As Integer) As Integer
    Dim lsSQL As String
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrDameAlternativaPolizaMobiliaria
    lsSQL = "EXEC stp_sel_ERS0632014_AlternativaPolizaMobiliaria " & pnTpoPoliza
    Set rs = coConex.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        DameAlternativaPolizaMobiliaria = rs!nAlternativa
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrDameAlternativaPolizaMobiliaria:
    Err.Raise Err.Number, "Error al recuperar alternativa para Polizas Mobiliareas", Err.Description
End Function
Public Function RecuperaDatosCalculoPoliza(ByVal pnMonedaPOL As Moneda, ByVal pnSumaAsegurada As Currency, ByVal pbInmobiliaria As Boolean, ByVal psAgeCod As String, ByVal pnTpoAltPoliza As Integer) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaDatosCalculoPoliza
    lsSQL = "EXEC stp_sel_ERS0632014_DatosCalculoPoliza " & pnMonedaPOL & "," & pnSumaAsegurada & "," & IIf(pbInmobiliaria, 1, 0) & ",'" & psAgeCod & "'," & pnTpoAltPoliza
    Set RecuperaDatosCalculoPoliza = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaDatosCalculoPoliza:
    Err.Raise Err.Number, "Error al recuperar datos para el cálculo de Poliza para Garantía Inmobiliaria/Mobiliaria/Otras Mobiliareas", Err.Description
End Function
Public Function RecuperaDatosCreditoxPoliza(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaDatosCreditoxPoliza
    lsSQL = "EXEC stp_sel_ERS0632014_DatosCredxPoliza '" & psCtaCod & "'"
    
    Set RecuperaDatosCreditoxPoliza = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaDatosCreditoxPoliza:
    Err.Raise Err.Number, "Error al recuperar datos del crédito para Polizas", Err.Description
End Function
Public Function RecuperaCalendarioCreditoxPolizaInt(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaCalendarioCreditoxPolizaInt
    lsSQL = "EXEC stp_sel_ERS0632014_CalendarioCredxPolizaInt '" & psCtaCod & "'"
    
    Set RecuperaCalendarioCreditoxPolizaInt = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaCalendarioCreditoxPolizaInt:
    Err.Raise Err.Number, "Error al recuperar calendario de pago del crédito para Polizas Internas", Err.Description
End Function
Public Function RecuperaGarantiasCreditoxPolizaInt(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaGarantiasCreditoxPolizaInt
    lsSQL = "EXEC stp_sel_ERS0632014_GarantiasCredxPolizaInt '" & psCtaCod & "'" ' & IIf(pbRenovacionPoliza, 1, 0)
    
    Set RecuperaGarantiasCreditoxPolizaInt = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaGarantiasCreditoxPolizaInt:
    Err.Raise Err.Number, "Error al recuperar las garantías del crédito para Polizas Internas", Err.Description
End Function
Public Function RecuperaGarantiasCreditoxPolizaExt(ByVal psCtaCod As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaGarantiasCreditoxPolizaExt
    lsSQL = "EXEC stp_sel_ERS0632014_GarantiasCredxPolizaExt '" & psCtaCod & "'"
    
    Set RecuperaGarantiasCreditoxPolizaExt = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaGarantiasCreditoxPolizaExt:
    Err.Raise Err.Number, "Error al recuperar las garantías del crédito para Polizas Externas", Err.Description
End Function
Public Sub ActualizarPeriodoPoliza(ByVal psNumPoliza As String, ByVal pdVigencia As Date, ByVal pdVencimiento As Date)
    Dim lsSQL As String
    On Error GoTo ErrActualizarPeriodoPoliza
    lsSQL = "EXEC stp_upd_ERS0632014_PeriodoPoliza '" & psNumPoliza & "','" & Format(pdVigencia, "yyyymmdd") & "','" & Format(pdVencimiento, "yyyymmdd") & "'"
    coConex.Ejecutar lsSQL
    Exit Sub
ErrActualizarPeriodoPoliza:
    Err.Raise Err.Number, "Error al actualizar el Periodo de la Poliza", Err.Description
End Sub
Public Function RecuperaGarantiaxVerificacionLegal(ByVal psNumGarant As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaGarantiaxVerificacionLegal
    lsSQL = "EXEC stp_sel_ERS0632014_GarantiaxVerificacionLegal '" & psNumGarant & "'"
    
    Set RecuperaGarantiaxVerificacionLegal = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaGarantiaxVerificacionLegal:
    Err.Raise Err.Number, "Error en el proceso <RecuperaGarantiaxVerificacionLegal>", Err.Description
End Function
Public Function RecuperaGarantiaxBloqueoLegal(ByVal psNumGarant As String) As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrRecuperaGarantiaxBloqueoLegal
    lsSQL = "EXEC stp_sel_ERS0632014_GarantiaxBloqueoLegal '" & psNumGarant & "'"
    
    Set RecuperaGarantiaxBloqueoLegal = coConex.CargaRecordSet(lsSQL)
    Exit Function
ErrRecuperaGarantiaxBloqueoLegal:
    Err.Raise Err.Number, "Error en el proceso <RecuperaGarantiaxBloqueoLegal>", Err.Description
End Function
Public Function CadenaPermiteModificarCobertura(ByVal psCtaCod As String, ByVal pbMoneyEqual As Boolean, ByVal pnMontoCRE As Currency, ByVal pnMontoGAR As Currency, ByRef pbBloqueaGarantiaPersona As Boolean) As String
    Dim rs As New ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrCadenaPermiteModificarCobertura
    lsSQL = "EXEC stp_sel_ERS0632014_PermiteModificarCobertura '" & psCtaCod & "'," & IIf(pbMoneyEqual, 1, 0) & "," & pnMontoCRE & "," & pnMontoGAR
    
    Set rs = coConex.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        CadenaPermiteModificarCobertura = rs!cMensaje
        pbBloqueaGarantiaPersona = rs!bBloqueGarantiaPersona
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrCadenaPermiteModificarCobertura:
    Err.Raise Err.Number, "Error en el proceso <PermiteModificarCobertura>", Err.Description
End Function
Public Function CadenaGarantiasPendienteMigracion(ByVal psCtaCod As String) As String
    Dim rs As New ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrCadenaGarantiasPendienteMigracion
    lsSQL = "SELECT cCadena = dbo.fnc_ERS0632014_CadenaGarantiasPendienteMigracion('" & psCtaCod & "')"
    
    Set rs = coConex.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        CadenaGarantiasPendienteMigracion = rs!cCadena
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrCadenaGarantiasPendienteMigracion:
    Err.Raise Err.Number, "Error en el proceso <CadenaGarantiasPendienteMigracion>", Err.Description
End Function
Public Sub ConsolidarGarantia(Optional ByVal psNumGarant As String = "")
    Dim lsSQL As String
    On Error GoTo ErrConsolidaGarantia
    lsSQL = "EXEC stp_upd_ERS0632014_ConsolidaGarantia '" & psNumGarant & "'"
    coConex.Ejecutar lsSQL
    Exit Sub
ErrConsolidaGarantia:
    Err.Raise Err.Number, "Error al consolidar datos de la Garantía", Err.Description
End Sub
Public Sub ConsolidarGarantiaValoriza(ByVal psNumGarant As String, ByVal pdValorizacion As Date)
    Dim lsSQL As String
    On Error GoTo ErrConsolidarGarantiaValoriza
    lsSQL = "EXEC stp_upd_ERS0632014_ConsolidaGarantiaValoriza '" & psNumGarant & "','" & Format(pdValorizacion, "yyyymmdd hh:mm:ss") & "'"
    coConex.Ejecutar lsSQL
    Exit Sub
ErrConsolidarGarantiaValoriza:
    Err.Raise Err.Number, "Error al consolidar las valorizaciones de la Garantia", Err.Description
End Sub
Public Function CadenaCoberturaGarantia(ByVal psNumGarant As String, ByVal pbTramiteLegal As Boolean, ByVal pnTipoTramite As Integer, ByVal pnVRM As Currency, ByVal pnGravamen As Currency) As String
    Dim rs As New ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrCadenaCoberturaGarantia
    lsSQL = "SELECT cCadena = dbo.fnc_ERS0632014_CadenaCoberturaGarantia('" & psNumGarant & "'," & IIf(pbTramiteLegal, 1, 0) & "," & pnTipoTramite & "," & pnVRM & "," & pnGravamen & " )"
    
    Set rs = coConex.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        CadenaCoberturaGarantia = rs!cCadena
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrCadenaCoberturaGarantia:
    Err.Raise Err.Number, "Error en el proceso <CadenaCoberturaGarantia>", Err.Description
End Function
Public Function RecuperaGarantiaxVerCredito(ByVal psNumGarant As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oRs As New Recordset
    On Error GoTo ErrRecuperaGarantiaxVerCredito
    
    sSql = "EXEC stp_sel_ERS0632014_GarantiaxVerCredito '" & psNumGarant & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaGarantiaxVerCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrRecuperaGarantiaxVerCredito:
    Err.Raise Err.Number, "Error en obtener créditos de garantía", Err.Description
End Function
Public Function RecuperaCreditosGarantiaxVerCredito(ByVal psNumGarant As String) As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oRs As New Recordset
    On Error GoTo ErrRecuperaCreditosGarantiaxVerCredito
    
    sSql = "EXEC stp_sel_ERS0632014_CreditosGarantiaxVerCredito '" & psNumGarant & "'"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaCreditosGarantiaxVerCredito = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrRecuperaCreditosGarantiaxVerCredito:
    Err.Raise Err.Number, "Error en obtener créditos de garantía", Err.Description
End Function
Public Function RecuperaIFixGravamenOtraIFi() As ADODB.Recordset
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oRs As New Recordset
    On Error GoTo ErrRecuperaIFixGravamenOtraIFi
    
    sSql = "EXEC stp_sel_ERS0022016_IFisxValorizacionGravamenOtraIFi"
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set RecuperaIFixGravamenOtraIFi = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrRecuperaIFixGravamenOtraIFi:
    Err.Raise Err.Number, "Error en obtener lista de instituciones financieras", Err.Description
End Function
'END EJVG *******
'EJVG20160318 ***
Public Function RecuperaGarantiaxVerificarActRegCobertura(ByVal psNumGarant As String, ByVal psCtaCod As String, ByVal psTpoProdCod As String, ByVal pbCliPreferencial As Boolean) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String

    On Error GoTo ErrVerificarGarantiaxRegistroCobertura
    
    sSql = "EXEC stp_sel_ERS0633_RecuperaGarantiaxVerificarActRegCobertura '" & psNumGarant & "','" & psCtaCod & "','" & psTpoProdCod & "'," & IIf(pbCliPreferencial, 1, 0)
    oConecta.AbreConexion
    Set RecuperaGarantiaxVerificarActRegCobertura = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrVerificarGarantiaxRegistroCobertura:
    Err.Raise Err.Number, "Error al recuperar datos de la garantía para verificar si se ha actualizado Valorización y/o Trámite Legal", Err.Description
End Function
Public Sub ActualizarCoberturaConfigRapiFlash(ByVal pnCampana As Integer, ByVal pnRetiroTpo As Integer, ByVal pnRetiroMonto As Currency)
    Dim lsSQL As String
    On Error GoTo ErrActualizarCoberturaConfigRapiFlash
    If pnCampana = 88 Then
        lsSQL = "EXEC stp_upd_ERS0632014_CoberturaConfigRapiFlashCamp " & pnRetiroTpo & "," & pnRetiroMonto
    Else
        lsSQL = "EXEC stp_upd_ERS0632014_CoberturaConfigRapiFlash " & pnRetiroTpo & "," & pnRetiroMonto
    End If
    coConex.Ejecutar lsSQL
    Exit Sub
ErrActualizarCoberturaConfigRapiFlash:
    Err.Raise Err.Number, "Error al actualizar las excepciones de la configuración de coberturas RapiFlash", Err.Description
End Sub
'END EJVG *******
'RECO20160419 ERS013-2016 ************************************************
Public Function ObtieneGarantLiq(ByVal psNumGarant As String) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String

    On Error GoTo ErrObtieneGarantLiq
    
    sSql = "EXEC stp_sel_ObtieneGarantLiq '" & psNumGarant & "'"
    oConecta.AbreConexion
    Set ObtieneGarantLiq = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrObtieneGarantLiq:
    Err.Raise Err.Number, "Error al obtener la clasificacion de la garantia", Err.Description
End Function
'RECO FIN ****************************************************************
'EJVG20160420 ***
Public Sub EliminarRegCoberturaCredProceso(ByVal psNumGarant As String)
    Dim lsSQL As String
    On Error GoTo ErrEliminarRegCoberturaCredProceso
    lsSQL = "EXEC stp_del_ERS0632014_ColocGarantiaXCredProceso '" & psNumGarant & "'"
    coConex.Ejecutar lsSQL
    Exit Sub
ErrEliminarRegCoberturaCredProceso:
    Err.Raise Err.Number, "Error al eliminar Registro de Cobertura", Err.Description
End Sub
Public Sub ActivaExclusionValidaDescobertura(ByVal psNumGarant As String, ByVal pbActiva As Boolean)
    Dim lsSQL As String
    On Error GoTo ErrActivaExclusionValidaDescobertura
    lsSQL = "EXEC stp_upd_ERS0632014_ActivaExclusionValidaDescobertura '" & psNumGarant & "'," & IIf(pbActiva, 1, 0)
    coConex.Ejecutar lsSQL
    Exit Sub
ErrActivaExclusionValidaDescobertura:
    Err.Raise Err.Number, "Error al activar exclusion para grabar datos así esté descoberturado", Err.Description
End Sub
Public Function ObviarDescobertura(ByVal psNumGarant As String) As Boolean
    Dim lsSQL As String
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrObviarDescobertura
    lsSQL = "EXEC stp_sel_ERS0632014_ObviarDescobertura '" & psNumGarant & "'"
    Set rs = coConex.CargaRecordSet(lsSQL)
    If rs.RecordCount > 0 Then
        ObviarDescobertura = True
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrObviarDescobertura:
    Err.Raise Err.Number, "Error al consultar si se va a obviar la descobertura", Err.Description
End Function
'END EJVG *******
'***JGPA TI-ERS026-2018------------------------------------------------------------------
Public Sub DarDeBajaGarantia(ByVal psNumGarant As String)
    Dim lsSQL As String
    On Error GoTo ErrEliminaGarantia
    lsSQL = "EXEC stp_udp_ERS026_2018_DarDeBajaGarantia'" & psNumGarant & "'"
    coConex.Ejecutar lsSQL
    Exit Sub
ErrEliminaGarantia:
    Err.Raise Err.Number, "Error al eliminar la garantía", Err.Description
End Sub
Public Function ObtenerCreditosEnProcesoxGarantia(ByVal psNumGarant As String) As ADODB.Recordset
Dim lsSQL As String
Dim oConecta As New COMConecta.DCOMConecta
Dim rs As New ADODB.Recordset
    On Error GoTo ErrObtenerCredEnProceso
    lsSQL = "EXEC stp_sel_ERS026_2018_ObtenerCreditosEnProcesoxGarantia '" & psNumGarant & "'"
    oConecta.AbreConexion
    Set ObtenerCreditosEnProcesoxGarantia = oConecta.CargaRecordSet(lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrObtenerCredEnProceso:
    Err.Raise Err.Number, "Error al consultar créditos en proceso", Err.Description
End Function
Public Function VerificaCargoLegal(ByVal psUser As String) As ADODB.Recordset
Dim lsSQL As String
Dim oConecta As New COMConecta.DCOMConecta
Dim rs As New ADODB.Recordset
    On Error GoTo ErrVerificaCargoLegal
    lsSQL = "EXEC stp_sel_ERS026_2018_VerificarCargoLegal '" & psUser & "'"
    oConecta.AbreConexion
    Set VerificaCargoLegal = oConecta.CargaRecordSet(lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrVerificaCargoLegal:
    Err.Raise Err.Number, "Error al verificar cargo legal.", Err.Description
End Function

'***END JGPA-----------------------------------------------------------------------------------
'CTI5 ERS001-2020******************************************************************************
Public Function ObtenerTpoBienFuturo() As ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "EXEC stp_sel_GarantiaContratoBienFuturo "
    Set ObtenerTpoBienFuturo = coConex.CargaRecordSet(lsSQL)
End Function
Public Function ObtenerDocumentoIdentidadTitularGarantia(ByVal psPersCod As String) As ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "EXEC stp_sel_ObtenerDocumentoIdentidadTitularGarantia '" & psPersCod & "'"
    Set ObtenerDocumentoIdentidadTitularGarantia = coConex.CargaRecordSet(lsSQL)
End Function
Public Function ObtenertTpoDocumentoBienFuturoxGarantia(ByVal cNumGarant As String, ByVal cNumDoc As String) As ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "EXEC stp_sel_GarantiaContratoBienFuturoxGarantia '" & cNumGarant & "','" & cNumDoc & "'"
    Set ObtenertTpoDocumentoBienFuturoxGarantia = coConex.CargaRecordSet(lsSQL)
End Function

Public Function ObtenerGarantiaCreditoDesemboldado(ByVal cNumGarant As String) As ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "EXEC stp_sel_GarantiaBienFuturoDesembolsado '" & cNumGarant & "'"
    Set ObtenerGarantiaCreditoDesemboldado = coConex.CargaRecordSet(lsSQL)
End Function

Public Function ObtenerGarantiaBienFuturoCreditoHipo(ByVal cNumGarant As String) As ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "EXEC stp_sel_GarantiaBienFuturoCredHipotecario '" & cNumGarant & "'"
    Set ObtenerGarantiaBienFuturoCreditoHipo = coConex.CargaRecordSet(lsSQL)
End Function

Public Function ObtenerGarantiaBienFuturo(ByVal nContratoId As Integer) As ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "EXEC stp_sel_GarantiaBienFuturo " & nContratoId & ""
    Set ObtenerGarantiaBienFuturo = coConex.CargaRecordSet(lsSQL)
End Function
Public Function ObtenerGarantiaColocaciones(ByVal cNumGarant As String) As ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "EXEC stp_sel_GarantiaColocaciones '" & cNumGarant & "'"
    Set ObtenerGarantiaColocaciones = coConex.CargaRecordSet(lsSQL)
End Function
Public Function VerificaCargosGarantiasDesembolsadosBienFuturo(ByVal psUser As String) As ADODB.Recordset
Dim lsSQL As String
Dim oConecta As New COMConecta.DCOMConecta
Dim rs As New ADODB.Recordset
    On Error GoTo ErrVerificaCargoLegal
    lsSQL = "EXEC stp_sel_VerificarVariosCargosGarantiasDesembolsadosBienFuturo '" & psUser & "'"
    oConecta.AbreConexion
    Set VerificaCargosGarantiasDesembolsadosBienFuturo = oConecta.CargaRecordSet(lsSQL)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrVerificaCargoLegal:
    Err.Raise Err.Number, "Error al verificar cargo legal.", Err.Description
End Function
Public Sub GrabarPolizaBienFuturo(ByVal psNumGarant As String, ByVal pdFecha As Date, ByVal psMovNroPoliza As String)
    On Error GoTo ErrGrabarPolizaBienFuturo
        Dim lsSQL As String
        lsSQL = "EXEC STP_GenerarPolizaHIPOTECARIO_FMV_TP '" & psNumGarant & "','" & Format(pdFecha, "yyyymmdd") & "','" & psMovNroPoliza & "'"
        coConex.Ejecutar lsSQL
        Exit Sub
ErrGrabarPolizaBienFuturo:
        Err.Raise Err.Number, "Error al grabar poliza.", Err.Description
End Sub
Public Sub ActualizaTCEA(ByVal psCtaCod As String, pnTCEA As Double)
Dim lsSQL As String
lsSQL = "EXEC sp_sel_ActualizarTCEA '" & psCtaCod & "'," & pnTCEA & ""
coConex.Ejecutar (lsSQL)
End Sub

Public Function ObtenerCalendarioMIVIENDA(ByVal pcCtaCod As String) As ADODB.Recordset
Dim sSql As String
    sSql = "Exec stp_sel_CalendarioCredito '" & pcCtaCod & "'"
    Set ObtenerCalendarioMIVIENDA = coConex.CargaRecordSet(sSql)
End Function
Public Function ObtenerDatosCredito(ByVal psNumGarant As String) As ADODB.Recordset
Dim sSql As String
    sSql = "Exec stp_sel_ObtenerDatosCreditoxGarantia '" & psNumGarant & "'"
    Set ObtenerDatosCredito = coConex.CargaRecordSet(sSql)
End Function
Public Function IdentificarTipoPeriodoMIVIENDA(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    
    sSql = "Exec sp_sel_IdentTipodePeriodo '" & psCtaCod & "'"
    Set IdentificarTipoPeriodoMIVIENDA = coConex.CargaRecordSet(sSql)

End Function
Public Function ObtenerGarantiaColocacionesCobertura(ByVal cNumGarant As String) As ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "EXEC stp_sel_GarantiaColocacionesCobertura '" & cNumGarant & "'"
    Set ObtenerGarantiaColocacionesCobertura = coConex.CargaRecordSet(lsSQL)
End Function
'**********************************************************************************************
