VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DMantGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Function ConsultarTpoGarantia(ByVal pnTpoGarantia As Integer) As Recordset
    Dim sSql As String
    Dim oConec As DConecta
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
        
        Set oConec = New DConecta
        oConec.AbreConexion
        Set ConsultarTpoGarantia = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ConsultarTpoGarantia = Null
End Function

Public Function CargarTpoGarantia() As Recordset
    Dim sSql As String
    Dim oConec As DConecta

    sSql = "Select nConsValor,cConsDescripcion From Constante Where nConsCod=1027 and nConsValor<>1027 "
    Set oConec = New DConecta
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
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    On Error GoTo ErrHandler
        sSql = "Select Count(*) as Cantidad"
        sSql = sSql & " From GarantDoc Where nConsValor=" & pTpoGarantia & " and nDocTpo=" & pDocTpo
        
        Set oConec = New DConecta
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


Public Function ActualizarGarantDoc(ByVal pRs As ADODB.Recordset, ByVal nTpoGarant As Integer) As Boolean
    Dim sSql As String
    Dim oConec As DConecta
    On Error GoTo ErrHandler
        Set oConec = New DConecta
        oConec.AbreConexion
        oConec.ConexionActiva.BeginTrans
        If Not pRs.BOF And Not pRs.EOF Then
            pRs.MoveFirst
        End If
        Do Until pRs.EOF
            If Val(pRs!Asignacion) = 1 Then
                ' entonces se registra o actualiza
                   If VerificaGarantDoc(nTpoGarant, pRs!TpoDoc) = False Then
                       'SE INSERTA
                        sSql = "INSERT INTO  GARANTDOC VALUES(" & nTpoGarant & "," & pRs!TpoDoc & ")"
                        oConec.ConexionActiva.Execute sSql
                   End If
            Else
                If VerificaGarantDoc(pRs!TpoGarantia, pRs!TpoDoc) = True Then
                      sSql = "DELETE FROM GARANTDOC WHERE nConsValor=" & nTpoGarant & " AND nDocTpo=" & pRs!TpoDoc
                       oConec.ConexionActiva.Execute sSql
                End If
            End If
            pRs.MoveNext
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

Public Function CargarNombrePersona(ByVal psctacod As String) As String
    Dim oConecta As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim sNombre As String
    
    
    Set oConecta = New DConecta
    sSql = "Select cPersNombre From Persona Where cPersCod='" & psctacod & "'"
    
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


Public Function VerificarCreditoAutomatico(ByVal psctacod As String) As Boolean
    Dim oConecta As DConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim nCantidad As Integer
    ' verifica si el credito es automatico o no?
    
    sSql = "Select count(*) as Cantidad "
    sSql = sSql & " From Producto P"
    sSql = sSql & " Inner Join ColocacCred CC on P.cCtaCod=CC.cCtaCod"
    sSql = sSql & " Where P.cCtaCod='" & psctacod & "' and P.nPrdEstado=2000 and"
    sSql = sSql & " CC.nColocCondicion2 = " & gCredAutomatico

     Set oConecta = New DConecta
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


Public Function ListaCreditosGarantias(ByVal pcPersCod As String) As Recordset
    Dim sSql As String
    Dim oConec As DConecta
    
    sSql = "Select P.cCtaCod,G.cNumGarant,G.cDescripcion as cGarantias,PE.cPersNombre,CG.nGravado as nMonto,"
    sSql = sSql & " G.nPorGravar as nDisponible"
    sSql = sSql & " From Producto P"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=P.cCtaCod and PP.nPrdPersRelac=" & gColRelPersTitular
    sSql = sSql & " Inner Join ColocGarantia CG on CG.cCtaCod=P.cCtaCod"
    sSql = sSql & " Inner Join Garantias G on G.cNumGarant=CG.cNumGarant"
    sSql = sSql & "  Inner Join Persona PE on PE.cPersCod=PP.cPersCod"
    sSql = sSql & " Where PP.cPersCod='" & pcPersCod & "'"

    Set oConec = New DConecta
    Set ListaCreditosGarantias = oConec.CargaRecordSet(sSql)
    
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerMaxcNumGarant() As String
    Dim sSql As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sNumGarant As String
    
    
    Set oConec = New DConecta
    oConec.AbreConexion
    sSql = "Select Max(cNumGarant) as cNumGarant From Garantias"
    Set rs = oConec.CargaRecordSet(sSql)
    sNumGarant = Format(CLng(IIf(IsNull(rs!cNumGarant), "0", rs!cNumGarant)), "00000000")
    oConec.CierraConexion
    Set oConec = Nothing
    Set rs = Nothing
    ObtenerMaxcNumGarant = sNumGarant
End Function

Public Sub DJ(ByVal psNumGarant As String, ByVal pdFecSis As Date)
    Dim rs As ADODB.Recordset
    Dim oCon As DConecta
    Dim sql  As String
    Set oCon = New DConecta
    Set rs = New ADODB.Recordset

    oCon.AbreConexion
    sql = "SELECT GETDATE() "
    Set rs = oCon.CargaRecordSet(sql)
    oCon.CierraConexion
    Set oCon = Nothing
    
     With DRDJ
        Set .DataSource = rs
            .DataMember = ""
       .Orientation = rptOrientPortrait
       .inicio psNumGarant, pdFecSis
      .Refresh
      .Show vbModal
End With
           
    
End Sub
