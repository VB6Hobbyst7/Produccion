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

Public Function PerteneceACredito(ByVal pnTpoDoc As Integer, ByVal psNroDoc As String) As Boolean
Dim sSql As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

    sSql = "Select cTpoDoc, cNroDoc from ColocGarantia Where cTpoDoc = " & pnTpoDoc & " AND cNroDoc = '" & psNroDoc & "'"
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

Public Function RecuperaTiposDocumGarantias() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaTiposDocumGarantias
    
    sSql = "select D.cDocDesc, D.nDocTpo "
    sSql = sSql & " from Documento D Inner join OpeDoc O ON D.nDocTpo = O.nDocTpo "
    sSql = sSql & " Where cOpeCod = '300100' "
    
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

'*******************************************************///*************************************
'***     Rutina:           RecuperaGarantiasPersona
'***     Descripcion:      Recupera todas las Garantias que Posee una Persona
'***     Modificado por:        NSSE
'***     Fecha-Tiempo:         11/06/2001 12:27:28 PM
'***     Ultima Modificacion: Creacion de la Rutina
'*****************************************************************************************
Public Function RecuperaGarantiasPersona(ByVal psPersCod As String) As ADODB.Recordset
Dim oConecta As DConecta
Dim sSql As String

    On Error GoTo ErrorRecuperaGarantiasPersona
    sSql = " Select G.cTpoDoc, G.cNroDoc, G.nTpoGarantia, C.cConsDescripcion as cTpoGarDescripcion, "
    sSql = sSql & " G.cDescripcion, G.nMoneda, C2.cConsdescripcion as cMonedaDesc,  G.nTasacion, G.nRealizacion, G.nGravament, G.nPorGravar, G.nPorGravar - G.nGravament as nDisponible "
    sSql = sSql & " From Garantias G Inner Join PersGarantia PG ON G.cTpoDoc = PG.cTpoDoc And G.cNroDoc = PG.cNroDoc "
    sSql = sSql & "                  Inner Join Constante C ON G.nTpoGarantia = C.nConsValor AND C.nConsCod = " & gPersGarantia
    sSql = sSql & "                  Inner Join Constante C2 ON G.nMoneda = C2.nConsValor AND C2.nConsCod = " & gMoneda
    sSql = sSql & " WHERE PG.cPersCod = '" & psPersCod & "' AND PG.nRelacion = " & Trim(Str(gPersRelGarantiaTitular))
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGarantiasPersona = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaGarantiasPersona:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaGarantiaCredito(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaGarantiaCredito
    sSql = "Select CG.cNroDoc, CN.cConsDescripcion cTpoGarantia, G.cDescripcion, D.cDocDesc, CG.nGravado, CN2.cConsDescripcion cMoneda, G.nPorGravar, G.nGravaMent, CN3.cConsDescripcion cMonedaGar"
    sSql = sSql & " From ColocGarantia CG Inner Join Garantias G ON CG.cTpoDoc = G.cTpoDoc AND CG.cNroDoc = G.cNroDoc "
    sSql = sSql & "                       Inner Join Constante CN ON G.nTpoGarantia = CN.nConsValor AND CN.nConsCod = " & gPersGarantia
    sSql = sSql & "                       Inner Join Constante CN2 ON CN2.nConsValor = CG.nMoneda AND CN2.nConsCod = " & gMoneda
    sSql = sSql & "                       Inner Join Constante CN3 ON CN3.nConsValor = G.nMoneda AND CN3.nConsCod = " & gMoneda
    sSql = sSql & "                       Inner Join Documento D ON D.nDocTpo = Convert(int,G.cTpoDoc) "
    sSql = sSql & "                       Inner Join OpeDoc O ON D.nDocTpo = O.nDocTpo AND O.cOpeCod = '300100'"
    sSql = sSql & " WHERE CG.cCtaCod = '" & psCtaCod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGarantiaCredito = oConecta.CargaRecordSet(sSql)
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
    
    sSql = "Select cTpoDoc, cNroDoc, nTpoGarantia, dGarantia, cDescripcion, cZona, nMoneda, "
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
'*****************************************************************************************
'***     Rutina:           RecuperaRelacPersonaGarantia
'***     Descripcion:       Recupera todas las Personas Relacionadas con la Garantia
'***     Modificado por:        NSSE
'***     Fecha-Tiempo:         11/06/2001 11:20:42 AM
'***     Ultima Modificacion: Creacion
'*****************************************************************************************
Public Function RecuperaRelacPersonaGarantia(ByVal psTipoDoc As String, ByVal psNroDoc As String)
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaRelacPersonaGarantia
    sSql = "select P.cPersCod, P.cPersNombre, PG.cNroDoc, C.cConsDescripcion + space(50) + convert(varchar(2),PG.nRelacion)  as cRelacion "
    sSql = sSql & " from persgarantia PG Inner Join Persona P ON PG.cPersCod  = P.cPersCod "
    sSql = sSql & " inner join Constante C ON PG.nRelacion = C.nConsValor "
    sSql = sSql & " where C.nConsCod = '" & gPersRelGarantia & "' and PG.cTpoDoc = '" & psTipoDoc & "' and PG.cNroDoc='" & psNroDoc & "'"
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaRelacPersonaGarantia = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function

ErrorRecuperaRelacPersonaGarantia:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Sub NuevaGarantia(ByVal psTipoDocGar As String, ByVal psNroDoc As String, _
                ByVal psTipoGar As String, ByVal psMoneda As String, ByVal psDescrip As String, _
                ByVal psZona As String, ByVal pnMontoTasac As Double, ByVal pnMontoReal As Double, _
                ByVal pnMontoxGrav As Double, ByVal psComent As String, RelPers As Variant, _
                ByVal dFecGar As Date)
                
Dim sSql As String
Dim oConecta As DConecta
Dim bTran As Boolean
Dim I As Integer
                On Error GoTo ErrorNuevaGarantia
                
                bTran = False
                Set oConecta = New DConecta
                oConecta.AbreConexion
                oConecta.ConexionActiva.BeginTrans
                bTran = True
                
                sSql = "INSERT INTO Garantias(cTpoDoc,cNroDoc,nTpoGarantia,dGarantia,cDescripcion,"
                sSql = sSql & "cZona,nMoneda,nTasacion,nRealizacion,nGravament,nPorGravar,cEstado,cComentario) "
                sSql = sSql & " VALUES('" & psTipoDocGar & "','" & psNroDoc & "'," & psTipoGar & ",'"
                sSql = sSql & Format(dFecGar, "mm/dd/yyyy") & "','" & psDescrip & "','" & psZona & "',"
                sSql = sSql & psMoneda & "," & Format(pnMontoTasac, "#0.00") & "," & Format(pnMontoReal, "#0.00") & ","
                sSql = sSql & "0.00," & Format(pnMontoxGrav, "#0.00") & ",'','" & psComent & "')"
                oConecta.ConexionActiva.Execute sSql
                
                sSql = "DELETE PersGarantia WHERE cTpoDoc = '" & psTipoDocGar & "' AND cNroDoc = '" & psNroDoc & "'"
                oConecta.ConexionActiva.Execute sSql
                
                For I = 0 To UBound(RelPers) - 1
                    sSql = "INSERT INTO PERSGARANTIA(cPersCod, cTpoDoc, cNroDoc, nRelacion)"
                    sSql = sSql & " VALUES('" & RelPers(I, 0) & "', "
                    sSql = sSql & " '" & RelPers(I, 1) & "', "
                    sSql = sSql & " '" & RelPers(I, 2) & "', "
                    sSql = sSql & " " & RelPers(I, 3) & ")"
                    oConecta.ConexionActiva.Execute sSql
                Next I
                oConecta.ConexionActiva.CommitTrans
                oConecta.CierraConexion
                Set oConecta = Nothing
                
                
                Exit Sub

ErrorNuevaGarantia:
        If bTran Then
            oConecta.ConexionActiva.RollBackTrans
        End If
        Set oConecta = Nothing
        MsgBox Err.Description, vbCritical, "Aviso"
                
End Sub

Public Function GarantiaEnUso(ByVal psTipoDocGar As String, ByVal psNroDoc As String) As Boolean
Dim sSql As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

    sSql = "Select * From ColocGarantia Where cTpoDoc = '" & psTipoDocGar & "' AND psNroDoc = '" & psNroDoc & "'"
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

Public Sub EliminarGraantia(ByVal psTipoDocGar As String, ByVal psNroDoc As String)
Dim sSql As String
Dim oConecta As DConecta
Dim bTran As Boolean

    On Error GoTo ErrorEliminarGraantia
    bTran = False
    Set oConecta = New DConecta
    oConecta.AbreConexion
    oConecta.ConexionActiva.BeginTrans
    bTran = True
    'Elimina PersGarantia
    sSql = "DELETE PersGarantia WHERE cTpoDoc = '" & psTipoDocGar & "' AND cNroDoc = '" & psNroDoc & "'"
    oConecta.ConexionActiva.Execute sSql
    'Elimina Garantias
    sSql = "DELETE Garantias WHERE cTpoDoc = '" & psTipoDocGar & "' AND cNroDoc = '" & psNroDoc & "'"
    oConecta.ConexionActiva.Execute sSql
    oConecta.ConexionActiva.CommitTrans
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub

ErrorEliminarGraantia:
    If bTran Then
        oConecta.ConexionActiva.RollBackTrans
        Set oConecta = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso EliminarGraantia", Err.Description

End Sub
Public Sub ActualizaGarantia(ByVal psTipoDocGar As String, ByVal psNroDoc As String, _
                ByVal psTipoGar As String, ByVal psMoneda As String, ByVal psDescrip As String, _
                ByVal psZona As String, ByVal pnMontoTasac As Double, ByVal pnMontoReal As Double, _
                ByVal pnMontoxGrav As Double, ByVal psComent As String, RelPers As Variant, _
                ByVal dFecGar As Date)
    
Dim sSql As String
Dim oConecta As DConecta
Dim bTran As Boolean
Dim I As Integer
            On Error GoTo ErrorActualizaGarantia
            bTran = False
            Set oConecta = New DConecta
            oConecta.AbreConexion
            oConecta.ConexionActiva.BeginTrans
            bTran = True
            sSql = "UPDATE Garantias SET "
            sSql = sSql & " nTpoGarantia = " & psTipoGar & ", "
            sSql = sSql & " dGarantia = '" & Format(dFecGar, "mm/dd/yyyy") & "', "
            sSql = sSql & " cDescripcion = '" & psDescrip & "',"
            sSql = sSql & " cZona = '" & psZona & "', "
            sSql = sSql & " nMoneda = " & psMoneda & ", "
            sSql = sSql & " nTasacion = " & Format(pnMontoTasac, "#0.00") & ", "
            sSql = sSql & " nRealizacion = " & Format(pnMontoReal, "#0.00") & ", "
            'sSql = sSql & " nGravament = " & Format(pnMontoxGrav, "#0.00") & ", "
            sSql = sSql & " nPorGravar = " & Format(pnMontoxGrav, "#0.00") & ", "
            sSql = sSql & " cEstado = '', "
            sSql = sSql & " cComentario = ' " & psComent & "'"
            sSql = sSql & " WHERE cTpoDoc = '" & psTipoDocGar & "' AND cNroDoc = '" & psNroDoc & "'"
            oConecta.ConexionActiva.Execute sSql
            
            sSql = "DELETE PersGarantia WHERE cTpoDoc = '" & psTipoDocGar & "' AND cNroDoc = '" & psNroDoc & "'"
            oConecta.ConexionActiva.Execute sSql
            
            For I = 0 To UBound(RelPers) - 1
                sSql = "INSERT INTO PERSGARANTIA(cPersCod, cTpoDoc, cNroDoc, nRelacion)"
                sSql = sSql & " VALUES('" & RelPers(I, 0) & "', "
                sSql = sSql & " '" & RelPers(I, 1) & "', "
                sSql = sSql & " '" & RelPers(I, 2) & "', "
                sSql = sSql & " " & RelPers(I, 3) & ")"
                oConecta.ConexionActiva.Execute sSql
            Next I
            oConecta.ConexionActiva.CommitTrans
            oConecta.CierraConexion
            Set oConecta = Nothing
    Exit Sub

ErrorActualizaGarantia:
        If bTran Then
            oConecta.ConexionActiva.RollBackTrans
        End If
        Set oConecta = Nothing
        Err.Raise Err.Number, "Error En Proceso ActualizaGarantia", Err.Description
End Sub

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing


End Sub

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

Public Function RecuperaGarantiaCreditoDatos(ByVal psCtaCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta

    On Error GoTo ErrorRecuperaGarantiaCredito
    sSql = "Select CN.nConsValor as nTpoGar, G.cNroDoc, CN.cConsDescripcion cTpoGarantia, G.cDescripcion, D.cDocDesc, CG.nGravado, CN2.cConsDescripcion cMoneda, G.nPorGravar, G.nGravaMent, CN3.cConsDescripcion cMonedaGar,G.cNumgarant, CG.nMoneda, CG.nEstado"
    sSql = sSql & " From ColocGarantia CG Inner Join Garantias G ON CG.cNumGarant = G.cNumGarant "
    sSql = sSql & "                       Inner Join Constante CN ON G.nTpoGarantia = CN.nConsValor AND CN.nConsCod = " & gPersGarantia
    sSql = sSql & "                       Inner Join Constante CN2 ON CN2.nConsValor = CG.nMoneda AND CN2.nConsCod = " & gMoneda
    sSql = sSql & "                       Inner Join Constante CN3 ON CN3.nConsValor = G.nMoneda AND CN3.nConsCod = " & gMoneda
    sSql = sSql & "                       Inner Join Documento D ON D.nDocTpo = Convert(int,G.cTpoDoc) "
    sSql = sSql & "                       Inner Join GarantDoc O ON D.nDocTpo = O.nDocTpo "
    sSql = sSql & " WHERE CG.cCtaCod = '" & psCtaCod & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGarantiaCreditoDatos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    Exit Function

ErrorRecuperaGarantiaCredito:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function
'ALPA ******20080818*****************
Public Function RecuperaGarantiasPersonaLogistica(ByVal psPersCod As String, Optional ByVal pbMantenimiento As Boolean = False, _
                                        Optional ByVal pbLiberacion As Boolean = False, Optional pbGarantClase As Boolean = False) As ADODB.Recordset

Dim oConecta As DConecta
Dim sSql As String

    On Error GoTo ErrorRecuperaGarantiasPersonaLogistica
    sSql = " Select DISTINCT G.nEstado, C3.cConsDescripcion as cEstadoGar, G.cNumGarant, P.cPersNombre, G.cTpoDoc, G.cNroDoc, G.nTpoGarantia, C.cConsDescripcion as cTpoGarDescripcion, "
    sSql = sSql & " G.cDescripcion, G.nMoneda, C2.cConsdescripcion as cMonedaDesc,  G.nTasacion, G.nRealizacion, G.nGravament, G.nPorGravar, G.nPorGravar - G.nGravament as nDisponible, G.cPersCodEmisor, "
    sSql = sSql & " D.cDocDesc,G.nGarClase, G.nGarTpoRealiz, G.cBancoPersCod, PB.cPersNombre as cBanco,G.nGarantReal"
    'By Capi 12082008
    'sSql = sSql & " ,G.cNumGarant" 'ARCV 11-07-2006
    sSql = sSql & " ,G.cNumGarant,C4.cConsDescripcion PFestado" 'ARCV 11-07-2006
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
    'ALPA 20080811****
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
    sSql = sSql & " AND G.nGarClase = 1 and nEstadoAdju in (8,10) "
    End If
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set RecuperaGarantiasPersonaLogistica = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing

    Exit Function

ErrorRecuperaGarantiasPersonaLogistica:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function RecuperaDatosGarantiaLogistica(ByVal psNumGarant As String, ByVal pnPeriodo As Integer) As ADODB.Recordset
Dim sSql As String
Dim oCon As DConecta

sSql = "exec stp_sel_RecuperaDatosGarantiaLogistica '" & psNumGarant & "'," & pnPeriodo & ""

Set oCon = New DConecta
oCon.AbreConexion
Set RecuperaDatosGarantiaLogistica = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Sub InsertarGarantiaLogistica(ByVal psNumGarant As String, ByVal pnPeriodo As Integer, ByVal pnTValor As Integer, ByVal pnMontValor As Double, ByVal pcDesValor As String, ByVal pnMoneda As Integer, ByVal pnEstado As Integer, ByVal psMovNro As String, ByVal psPersCod As String, ByVal pnContador As Integer)
Dim sSql As String
Dim oCon As DConecta

sSql = "exec stp_ins_garantiaLogistica '" & psNumGarant & "'," & pnPeriodo & "," & pnTValor & "," & pnMontValor & ",'" & pcDesValor & "', " & pnMoneda & ",1, '" & psMovNro & "', '" & psPersCod & "', " & pnContador

Set oCon = New DConecta
oCon.AbreConexion
oCon.CargaRecordSet (sSql)
oCon.CierraConexion
Set oCon = Nothing
End Sub

Public Sub dUpdateGarantiasAdjudicados(ByVal psGarantiaCod As String, ByVal pdFecha As Date, ByVal pnEstadoAdj As Integer, ByVal psUsuario As String)
Dim oCon As DConecta
Dim lsSQL As String
    
    lsSQL = "exec stp_upd_GarantiaAdjudicado '" & psGarantiaCod & "','" & Format(pdFecha, "YYYY/MM/DD") & "'," & pnEstadoAdj & ",'" & psUsuario & "'"
    Set oCon = New DConecta
    oCon.AbreConexion
        oCon.CargaRecordSet (lsSQL)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'****************************
Public Sub ActualizargarantiasxVentaLogistica(ByVal psNumGarant As String, ByVal pnEstadoAdju As Integer, ByVal pdEstadoAdju As Date, ByVal psUsuariAdju As String, ByVal pdFechaCompra As Date, ByVal pnMonedaAdju As Integer, ByVal pnTipoCambio As Double, ByVal psPersCodComprador As String, ByVal pnVendido As Integer)
Dim sSql As String
Dim oCon As DConecta

sSql = "exec stp_upd_ActualizargarantiasxVentaLogistica '" & psNumGarant & "'," & pnEstadoAdju & ",'" & Format(pdEstadoAdju, "YYYY/MM/DD") & "','" & psUsuariAdju & "','" & Format(pdFechaCompra, "YYYY/MM/DD") & "', " & pnMonedaAdju & "," & pnTipoCambio & ", '" & psPersCodComprador & "', " & pnVendido

Set oCon = New DConecta
oCon.AbreConexion
oCon.CargaRecordSet (sSql)
oCon.CierraConexion
Set oCon = Nothing
End Sub
Public Function RecuperaDatosGarantiaSanemiento(ByVal psNumGarant As String, ByVal psCtaCod As String, ByVal pnTESan As Integer) As ADODB.Recordset
Dim sSql As String
Dim oCon As DConecta

sSql = "exec stp_sel_RecuperaDatosGarantiaSanemiento '" & psNumGarant & "','" & psCtaCod & "'," & pnTESan

Set oCon = New DConecta
oCon.AbreConexion
Set RecuperaDatosGarantiaSanemiento = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function
Public Sub InsertarGarantiaSaneamiento(ByVal psNumGarant As String, ByVal psCtaCod As String, ByVal pnTipoSan As Integer, ByVal pnPeriSan As Integer, ByVal pnMontSan As Double, ByVal pdFecSan As Date, ByVal psUsuariAdju As String, ByVal pnContador As Integer, ByVal pnMoneda As Integer, ByVal nTipoSaneame As Integer)
Dim sSql As String
Dim oCon As DConecta

sSql = "exec stp_ins_GarantiaSaneamiento '" & psNumGarant & "','" & psCtaCod & "'," & pnTipoSan & "," & pnPeriSan & "," & pnMontSan & ", '" & Format(pdFecSan, "YYYY/MM/DD") & "', '" & psUsuariAdju & "', " & pnContador & ", " & pnMoneda & "," & nTipoSaneame

Set oCon = New DConecta
oCon.AbreConexion
oCon.CargaRecordSet (sSql)
oCon.CierraConexion
Set oCon = Nothing
End Sub
Public Function ReporteGarantiasLogisticaxPeriodo(ByVal pdFecha As Date) As ADODB.Recordset
Dim sSql As String
Dim oCon As DConecta

sSql = "exec stp_sel_reporteGarantiasLogisticaxPeriodo '" & Format(pdFecha, "yyyy/mm/dd") & "'"

Set oCon = New DConecta
oCon.AbreConexion
Set ReporteGarantiasLogisticaxPeriodo = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function
