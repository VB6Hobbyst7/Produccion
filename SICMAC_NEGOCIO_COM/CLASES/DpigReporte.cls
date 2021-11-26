VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DpigReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Actualiza en Base Autorización Fisica
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB903050003
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB90305014E
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB903050112
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

Public Function GetRep(psql As String) As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If oCon.AbreConexion Then
        Set GetRep = oCon.CargaRecordSet(psql)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function



#If DebugMode Then
    '##ModelId=3AB90305005E
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If



Public Function RepInventarioJoyas(psFecini) As ADODB.Recordset
    Dim sqlE As String
    Dim lsFecha As String
    Dim oCon As DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oCon = New DConecta
    Dim sSql As String
    
sSql = " SELECT SUBSTRING(t1.cCtaCod, 4, 2) AS Agencia, ag.cAgeDescripcion, t1.cCtaCod, t3.dVigencia, pe.cPersNombre, t2.nPiezas, "
sSql = sSql & "     (SELECT SUM(cpjt.nPesoNeto) nPesoNeto "
sSql = sSql & "                    FROM   ColocPigJoyaTasacion cpjt "
sSql = sSql & "                    WHERE cpjt.nTipoTasacion = 1 AND cpjt.cctacod = t1.cctacod  "
sSql = sSql & "                    GROUP BY cpjt.cctacod) AS nPesoNeto,  "
sSql = sSql & "                   (SELECT SUM(nTasacion) + SUM(nTasacionAdicional) nTasacionAdicional  "
sSql = sSql & "                    FROM   dbo.ColocPigJoyaTasacion  "
sSql = sSql & "                    WHERE nTipoTasacion = 1 AND cctacod = t1.cctacod  "
sSql = sSql & "                    GROUP BY cctacod) AS ntasacion, ccdp.nMonto AS capital,  "
sSql = sSql & "                   (SELECT detpig.nmonto  "
sSql = sSql & "                    FROM   ColocCalendDetPig detpig  "
sSql = sSql & "                    WHERE detpig.cctacod = t1.cctacod AND detpig.cctacod = ccdp.cctacod AND detpig.nPrdConceptoCod = '8205' AND detpig.ncoloccalendapl = 1 AND   "
sSql = sSql & "                                   detpig.ncuota = 1 AND detpig.nNroCalen = t2.nnumcalend) AS Otros, "
sSql = sSql & "                   (SELECT detpig.nmonto "
sSql = sSql & "                    FROM   ColocCalendDetPig detpig "
sSql = sSql & "                    WHERE detpig.cctacod = t1.cctacod AND detpig.cctacod = ccdp.cctacod AND detpig.nPrdConceptoCod = '8100' AND detpig.ncoloccalendapl = 1 AND  "
sSql = sSql & "                                   detpig.ncuota = 1 AND detpig.nNroCalen = t2.nnumcalend) AS interescomp, "
sSql = sSql & "                   (SELECT detpig.nmonto "
sSql = sSql & "                    FROM   ColocCalendDetPig detpig "
sSql = sSql & "                    WHERE detpig.cctacod = t1.cctacod AND detpig.cctacod = ccdp.cctacod AND detpig.nPrdConceptoCod = '8101' AND detpig.ncoloccalendapl = 1 AND  "
sSql = sSql & "                                   detpig.ncuota = 1 AND detpig.nNroCalen = t2.nnumcalend) AS Mora, "
sSql = sSql & "                   (SELECT isnull(detpig.nmonto, 0) "
sSql = sSql & "                    FROM   ColocCalendDetPig detpig "
sSql = sSql & "                    WHERE detpig.cctacod = t1.cctacod AND detpig.cctacod = ccdp.cctacod AND detpig.nPrdConceptoCod = '8206' AND detpig.ncoloccalendapl = 1 AND  "
sSql = sSql & "                                   detpig.ncuota = 1 AND detpig.nNroCalen = t2.nnumcalend) AS PreparaRemate "
sSql = sSql & " FROM  ProductoPersona t1 INNER JOIN "
sSql = sSql & "               Agencias ag ON SUBSTRING(t1.cCtaCod, 4, 2) = ag.cAgeCod INNER JOIN "
sSql = sSql & "               ColocPigno t2 ON t1.cCtaCod = t2.cCtaCod INNER JOIN "
sSql = sSql & "               Colocaciones t3 ON t1.cCtaCod = t3.cCtaCod INNER JOIN "
sSql = sSql & "               Persona pe ON t1.cPersCod = pe.cPersCod INNER JOIN "
sSql = sSql & "               ColocCalendDetPig ccdp ON t1.cCtaCod = ccdp.cCtaCod AND t2.cCtaCod = ccdp.cCtaCod AND t1.cCtaCod = ccdp.cCtaCod AND  "
sSql = sSql & "               t2.nNumCalend = ccdp.nNroCalen and convert(varchar(10),t3.dVigencia,112) = '" & Format(psFecini, gsFormatoMovFecha) & "' "
sSql = sSql & " WHERE (ccdp.nColocCalendApl = 1) AND (ccdp.nCuota = 1) AND (ccdp.nPrdConceptoCod = '8000') order by SUBSTRING(t1.cCtaCod, 4, 2)"
'AND (t1.cCtaCod IN ('102043051000280593',  "
'ssql = ssql & "               '102053051000000060', '102083051000026921', '102013051000335304', '102013051000335410', '102013051000335924', '102013051000335991','102013051000336831'))  "


    
         
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sSql)
        Set RepInventarioJoyas = rs
        oCon.CierraConexion
    End If
    
    If rs Is Nothing Then
         MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
    
    If rs.EOF And rs.BOF Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
    Set oCon = Nothing

End Function


Public Function RepVentaMaterial(psAgencia As String, psFecha As String, nTipoAgrupacion As Integer) As ADODB.Recordset
    Dim sqlE As String
    Dim lsFecha As String
    Dim oCon As DConecta
    Dim rs As ADODB.Recordset
    Dim sAgrupa As String
    Set rs = New ADODB.Recordset
    Set oCon = New DConecta
    Dim sSql As String
    
    Select Case nTipoAgrupacion
        Case 1 'Por Remate y por Material
              sAgrupa = "d.nRemate,c.cConsDescripcion"
        Case 2 'Por Material
              sAgrupa = "c.cConsDescripcion"
        Case 3 'Por Rematel
              sAgrupa = "d.nRemate"
    End Select
sSql = "  select " & sAgrupa & "  ,h.cAgeDescripcion"
sSql = sSql & " ,count(*) Piezas"
sSql = sSql & " ,sum(a.nPesoNeto) Peso"
sSql = sSql & " ,sum(a.nTasacionAdicional) Adicional"
sSql = sSql & " ,sum(f.nMonto) ValorVenta"
sSql = sSql & " ,sum(e.CostoVenta) CostoVenta"
sSql = sSql & " ,sum(e.InteresPercibido) InteresPercibido"
sSql = sSql & "  ,sum(e.ComisionesPercibidas) ComisionesPercibidas"
sSql = sSql & "  from colocpigjoyatasacion a"
sSql = sSql & " inner  join colocpigno b on b.nTipoTasacion=a.nTipoTasacion and b.cCtaCod =a.cCtaCod"
sSql = sSql & " inner join Constante c on c.nConsValor = a.nMaterial"
sSql = sSql & " inner join ColocPigProceso d on d.cCtaCod=a.cCtaCod and d.nItemPieza = a.nItemPieza"
sSql = sSql & " Inner Join"
sSql = sSql & "  ( select nRemate,nTipoProceso,cCtaCod,nItemPieza"
sSql = sSql & " ,sum(case nCodConcepto  when " & gColPigConceptoCodCapital & " then nMonto else 0 end) CostoVenta"
sSql = sSql & " ,sum(case nCodConcepto  when " & gColPigConceptoCodInteresComp & " then nMonto when " & gcolPigConceptoCodInteresMora & " then nMonto else 0 end) InteresPercibido"
sSql = sSql & " ,sum(case nCodConcepto  when " & gColPigConceptoCodComiServ & " then nMonto  when " & gColPigConceptoCodPenalidad & " then nMonto when " & gColPigConceptoCodCustodia & " then nMonto when " & gColPigConceptoCodComiVencida & " then nMonto when " & gColPigConceptoCodPreparaRemate & " then nMonto else 0 end) ComisionesPercibidas"
sSql = sSql & "  From colocpigprocesodet"
sSql = sSql & " group by  nRemate,nTipoProceso,cCtaCod,nItemPieza ) e on e.nRemate=d.nRemate and e.nTipoProceso = d.nTipoProceso and e.cCtaCod = d.cCtaCod and e.nItemPieza= d.nItemPieza"
sSql = sSql & " inner join  movdocpigdet f on f.cCtaCod = a.cCtaCod and f.nPieza = a.nItemPieza"
sSql = sSql & " inner join movdocpig g on g.nCodTipo = f.nCodTipo and g.cDocumento=f.cDocumento and g.nNroMov=f.nNroMov"
sSql = sSql & " inner join agencias h on h.cAgeCod= substring(g.cUltimaActualizacion,18,2) "
sSql = sSql & " where c.nConsCod= " & gColocPigMaterial & " and d.nTipoProceso= " & gPigTipoVentas & " and d.nSituacionPieza= " & gPigSituacionFacturado & " and f.nCodConcepto=" & gColPigConceptoCodPrecioVenta
sSql = sSql & " and substring(g.cUltimaActualizacion,18,2) = " & psAgencia
sSql = sSql & " and  substring(a.cUltimaActualizacion,1,6) = " & psFecha
sSql = sSql & " group by " & sAgrupa & "  ,h.cAgeDescripcion"

  
         
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sSql)
        Set RepVentaMaterial = rs
        oCon.CierraConexion
    End If
    
    If rs Is Nothing Then
         MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
    
    If rs.EOF And rs.BOF Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
    Set oCon = Nothing

End Function


Public Function RepClientesAntDes() As ADODB.Recordset
    Dim sqlE As String
    Dim lsFecha As String
    Dim oCon As DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oCon = New DConecta
    Dim sSql As String
    

sSql = " (select 'Ant' as Situacion, y.cConsDescripcion, COUNT(*) cantidad "
sSql = sSql & " FROM colocpigevalhistorico E"
sSql = sSql & " inner join Constante y on y.nConsValor= c2"
sSql = sSql & " WHERE y.nConsCod='8005' and E.C2 IS NOT NULL "
sSql = sSql & " GROUP BY cConsDescripcion) "
sSql = sSql & " Union"
sSql = sSql & " (SELECT 'Act',x.cConsDescripcion,COUNT (*) CANT_ACT"
sSql = sSql & " FROM colocpigevalhistorico E"
sSql = sSql & " inner join Constante x on x.nConsValor= c1"
sSql = sSql & " WHERE x.nConsCod='8005' and E.C1 IS NOT NULL"
sSql = sSql & " GROUP BY cConsDescripcion) "
sSql = sSql & " order by 1 desc "

       
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sSql)
        Set RepClientesAntDes = rs
        oCon.CierraConexion
    End If
    
    If rs Is Nothing Then
         MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
    
    If rs.EOF And rs.BOF Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
    Set oCon = Nothing

End Function
Public Function RepClientesCalifManual() As ADODB.Recordset
    Dim sqlE As String
    Dim lsFecha As String
    Dim oCon As DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oCon = New DConecta
    Dim sSql As String
    

sSql = " SELECT distinct e.cPersCod, s.cPersNombre,e.cEvalPigno,e.cEvalPignoAnterior, e.fEvalPigno,cUsuarioEvalPigno "
sSql = sSql & " FROM colocpigevalmanual e"
sSql = sSql & " INNER JOIN  Persona s ON e.cPersCod = s.cPersCod"
sSql = sSql & " WHERE e.icalifica='N' "
sSql = sSql & " AND E.fEvalPigno=(SELECT MAX(fEvalPigno) "
sSql = sSql & " FROM colocpigevalmanual WHERE cPersCod=e.cPersCod) "
sSql = sSql & " ORDER BY S.cPersNombre "
 
         
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sSql)
        Set RepClientesCalifManual = rs
        oCon.CierraConexion
    End If
    
    If rs Is Nothing Then
         MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
    
    If rs.EOF And rs.BOF Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
    Set oCon = Nothing

End Function

Public Function RepClientesVarCalifi() As ADODB.Recordset
    Dim sqlE As String
    Dim lsFecha As String
    Dim oCon As DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oCon = New DConecta
    Dim sSql As String
    


sSql = "   select p.cPersCod,p.cPersNombre,u.cUbiGeoDescripcion,p.dPersIng,n.cPersNatSexo,cCalifiClienteAnt,cCalifiCliente"
sSql = sSql & " from colocpigevalcliente e"
sSql = sSql & "  inner join persona p on p.cPersCod = e.cPerscod"
sSql = sSql & "  left join ubicaciongeografica u on cUbiGeoCod = p.cPersDireccUbiGeo"
sSql = sSql & "  left join PersonaNAt n on n.cPersCod = p.cPersCod"
sSql = sSql & "  Where cCalifiClienteAnt <> cCalifiCliente"


    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sSql)
        Set RepClientesVarCalifi = rs
        oCon.CierraConexion
    End If
    
    If rs Is Nothing Then
         MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
    
    If rs.EOF And rs.BOF Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
    Set oCon = Nothing

End Function
Public Function RepVariacionCalifTipCliente() As ADODB.Recordset
    Dim sqlE As String
    Dim lsFecha As String
    Dim oCon As DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oCon = New DConecta
    Dim sSql As String
    


sSql = "  SELECT x.cConsDescripcion as des1, y.cConsDescripcion as des2,COUNT(*) CANT"
sSql = sSql & " From colocpigevalhistorico"
sSql = sSql & " inner join Constante x on x.nConsValor= c1"
sSql = sSql & " inner join Constante y on y.nConsValor= c2"
sSql = sSql & " Where C1 <> C2"
sSql = sSql & " and x.nConsCod='8005' and y.nConsCod='8005'"
sSql = sSql & " GROUP BY x.cConsDescripcion,y.cConsDescripcion"
sSql = sSql & " order by count(*) desc"

    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sSql)
        Set RepVariacionCalifTipCliente = rs
        oCon.CierraConexion
    End If
    
    If rs Is Nothing Then
         MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
    
    If rs.EOF And rs.BOF Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
    Set oCon = Nothing

End Function
