VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DpersReporte"
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

Public Function GetRep(pSQL As String) As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If oCon.AbreConexion Then
        Set GetRep = oCon.CargaRecordSet(pSQL)
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



Public Function RepInventarioJoyas(psFecIni) As ADODB.Recordset
    Dim sqlE As String
    Dim lsFecha As String
    Dim oCon As DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oCon = New DConecta
    Dim sSql As String
'¿ lsfecha = Format(psFecini, gsFormatoFechaPlanilla)
'    lsfecha = Left(lsfecha, 6)
    
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
sSql = sSql & "               t2.nNumCalend = ccdp.nNroCalen and convert(varchar(10),t3.dVigencia,112) = '" & Format(psFecIni, gsFormatoMovFecha) & "' "
sSql = sSql & " WHERE (ccdp.nColocCalendApl = 1) AND (ccdp.nCuota = 1) AND (ccdp.nPrdConceptoCod = '8000') "
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


Public Function ReporteDeClientes(psFecIni) As ADODB.Recordset
    Dim sqlE As String
    Dim lsFecha As String
    Dim oCon As DConecta
    Dim Rsc As ADODB.Recordset
    Set Rsc = New ADODB.Recordset
    Set oCon = New DConecta
    lsFecha = Format(psFecIni, gsFormatoMovFecha)
    'lsFecha = Left(lsFecha, 6)
    
    
    sqlE = " SELECT substring(pe.cPersCod,4,2)  Agencia, ag.cAgeDescripcion, pe.cPersCod, "
sqlE = sqlE & " replace(replace(replace(pe.cPersNombre,'/',' '),',',' '),'\',' ') cpersnombre, "
sqlE = sqlE & " pe.dPersNacCreac,   pe.cPersDireccUbiGeo, "
sqlE = sqlE & " cUbiGeoDescripcion = case "
sqlE = sqlE & " when ubiG.cUbiGeoDescripcion is null then 'Migración'"
sqlE = sqlE & " when ubiG.cUbiGeoDescripcion is not null then ubiG.cUbiGeoDescripcion end ,"
sqlE = sqlE & " pe.cPersDireccDomicilio, "
sqlE = sqlE & " pe.cPersTelefono,   LEFT(pe.cUltimaActualizacion, 8) AS fecha, "
sqlE = sqlE & " RIGHT(pe.cUltimaActualizacion, 5) AS usuario,   "
sqlE = sqlE & " substring(CON.cConsDescripcion,1,5) cConsDescripcion, pid.cPersIDnro FROM  "
sqlE = sqlE & " Persona pe inner JOIN   PersID pid ON pe.cPersCod = pid.cPersCod "
sqlE = sqlE & " inner JOIN   Constante con ON pid.cPersIDTpo = con.nConsValor  "
sqlE = sqlE & " inner join  agencias ag on ag.cAgeCod = substring(pe.cPersCod,4,2)  "
sqlE = sqlE & " left join  ubicaciongeografica ubig on uBiG.cUbiGeoCod = pe.cPersDireccUbiGeo"
sqlE = sqlE & "  WHERE  convert(varchar(8),pe.dpersing,112) = '" & lsFecha & "' AND (con.nConsCod = '1003')  order by  substring(pe.cPersCod,4,2)"
'sqlE = sqlE & "  WHERE  left(pe.cultimaactualizacion,8) = '" & lsFecha & "' AND (con.nConsCod = '1003') "
'

'
'       sqlE = " SELECT substring(pe.cPersCod,4,2)  Agencia, ag.cAgeDescripcion, pe.cPersCod, pe.cPersNombre, pe.dPersNacCreac, pe.cPersDireccUbiGeo, pe.cPersDireccDomicilio, pe.cPersTelefono, " & _
'            "  cUbiGeoDescripcion = case when ubiG.cUbiGeoDescripcion is null then 'Migración' " & _
'            "  when ubiG.cUbiGeoDescripcion is not null then ubiG.cUbiGeoDescripcion end , " & _
'            " LEFT(pe.cUltimaActualizacion, 8) AS fecha, RIGHT(pe.cUltimaActualizacion, 5) AS usuario, " & _
'            "  CON.cConsDescripcion , pid.cPersIDnro FROM  Persona pe inner JOIN " & _
'            "  PersID pid ON pe.cPersCod = pid.cPersCod inner JOIN " & _
'            "  Constante con ON pid.cPersIDTpo = con.nConsValor " & _
'            "  WHERE  left(pe.cultimaactualizacion,8) = '" & lsFecha & "' AND (con.nConsCod = '1003') "

    If oCon.AbreConexion Then
        Set Rsc = oCon.CargaRecordSet(sqlE)
        Set ReporteDeClientes = Rsc
        oCon.CierraConexion
    End If
    
    If Rsc Is Nothing Then
         MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
    
    If Rsc.EOF And Rsc.BOF Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
    Set oCon = Nothing

End Function


''
Public Function ReporteDeClientesMod(psFecIni) As ADODB.Recordset
    Dim sqlE As String
    Dim lsFecha As String
    Dim oCon As DConecta
    Dim Rsc As ADODB.Recordset
    Set Rsc = New ADODB.Recordset
    Set oCon = New DConecta
    lsFecha = Format(psFecIni, gsFormatoMovFecha)
     
sqlE = " SELECT substring(pe.cPersCod,4,2)  Agencia, ag.cAgeDescripcion, pe.cPersCod, "
sqlE = sqlE & " replace(replace(replace(pe.cPersNombre,'/',' '),',',' '),'\',' ') cpersnombre, "
sqlE = sqlE & " pe.dPersNacCreac,   pe.cPersDireccUbiGeo, "
sqlE = sqlE & " cUbiGeoDescripcion = case "
sqlE = sqlE & " when ubiG.cUbiGeoDescripcion is null then 'Migración'"
sqlE = sqlE & " when ubiG.cUbiGeoDescripcion is not null then ubiG.cUbiGeoDescripcion end ,"
sqlE = sqlE & " pe.cPersDireccDomicilio, "
sqlE = sqlE & " pe.cPersTelefono,   LEFT(pe.cUltimaActualizacion, 8) AS fecha, "
sqlE = sqlE & " RIGHT(pe.cUltimaActualizacion, 5) AS usuario,   "
sqlE = sqlE & " substring(CON.cConsDescripcion,1,5) cConsDescripcion, pid.cPersIDnro FROM  "
sqlE = sqlE & " Persona pe inner JOIN   PersID pid ON pe.cPersCod = pid.cPersCod "
sqlE = sqlE & " inner JOIN   Constante con ON pid.cPersIDTpo = con.nConsValor  "
sqlE = sqlE & " inner join  agencias ag on ag.cAgeCod = substring(pe.cPersCod,4,2)  "
sqlE = sqlE & " left join  ubicaciongeografica ubig on uBiG.cUbiGeoCod = pe.cPersDireccUbiGeo"
sqlE = sqlE & "  WHERE  left(pe.cultimaactualizacion,8) = '" & lsFecha & "' AND (con.nConsCod = '1003') order by  substring(pe.cPersCod,4,2) "

    If oCon.AbreConexion Then
        Set Rsc = oCon.CargaRecordSet(sqlE)
        Set ReporteDeClientesMod = Rsc
        oCon.CierraConexion
    End If
    
    If Rsc Is Nothing Then
         MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
    
    If Rsc.EOF And Rsc.BOF Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
    Set oCon = Nothing

End Function



