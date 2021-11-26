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

'*****************************************************************************************
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

Public Function RecuperaGarantia(ByVal psTipoDoc As String, ByVal psNroDoc As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
    
    On Error GoTo ErrorRecuperaGarantia
    
    sSql = "Select cTpoDoc, cNroDoc, nTpoGarantia, dGarantia, cDescripcion, cZona, nMoneda, nTasacion, nRealizacion, nGravament, nPorGravar, cEstado, cCodPersSeguro, nRangHip, cFlag, cRegistro, cComentario"
    sSql = sSql & " From Garantias Where cTpoDoc = '" & psTipoDoc & "' AND cNroDoc = '" & psNroDoc & "'"
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
Dim i As Integer
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
                
                For i = 0 To UBound(RelPers) - 1
                    sSql = "INSERT INTO PERSGARANTIA(cPersCod, cTpoDoc, cNroDoc, nRelacion)"
                    sSql = sSql & " VALUES('" & RelPers(i, 0) & "', "
                    sSql = sSql & " '" & RelPers(i, 1) & "', "
                    sSql = sSql & " '" & RelPers(i, 2) & "', "
                    sSql = sSql & " " & RelPers(i, 3) & ")"
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
        MsgBox Err.Description, vbCritical, "Aviso"
                
End Sub
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
        oConecta.ConexionActiva.RollbackTrans
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
Dim i As Integer
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
            
            For i = 0 To UBound(RelPers) - 1
                sSql = "INSERT INTO PERSGARANTIA(cPersCod, cTpoDoc, cNroDoc, nRelacion)"
                sSql = sSql & " VALUES('" & RelPers(i, 0) & "', "
                sSql = sSql & " '" & RelPers(i, 1) & "', "
                sSql = sSql & " '" & RelPers(i, 2) & "', "
                sSql = sSql & " " & RelPers(i, 3) & ")"
                oConecta.ConexionActiva.Execute sSql
            Next i
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

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

End Sub

'->***** LUCV20161010
Public Function ObtenerGarantiasJoyasRpt(ByVal pdFecha As Date, ByVal pnTipoGar As Integer) As ADODB.Recordset
    Dim lsSQL As String
    Dim oCon As New DConecta

    On Error GoTo ErrorObtenerGarantiasJoyasRpt
    lsSQL = "Exec stp_sel_GarantiasJoyasRpt '" & Format(pdFecha, "yyyymmdd") & "', " & pnTipoGar & ""
    Set oCon = New DConecta
    oCon.AbreConexion
    Set ObtenerGarantiasJoyasRpt = oCon.CargaRecordSet(lsSQL)
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function

ErrorObtenerGarantiasJoyasRpt:
    Set oCon = Nothing
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function
'<-***** LUCV20161010
