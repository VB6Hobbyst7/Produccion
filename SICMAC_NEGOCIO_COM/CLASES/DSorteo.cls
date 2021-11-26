VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSorteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public dbCmact As Connection
Dim sDBComunes As String
Dim sDBPersona As String
Dim sDBImagenes As String

Public Function GeneraNumSorteo(ByVal cAlcance As String, ByVal cAnio As String) As String
 Dim sSql As String, RSTEMP As Recordset, ccodeaux
 
 ccodeaux = "01"
 GeneraNumSorteo = ""
 
        sSql = "Select maximo=isnull(max(right(cnumsorteo,2)),'00') from sorteo where substring(cnumsorteo,3,4)='" & cAnio & "' and  left(cnumsorteo,2)='" & cAlcance & "'"
        Set RSTEMP = New Recordset
        RSTEMP.CursorLocation = adUseClient
        RSTEMP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
       ' With rstemp
            If Not RSTEMP.EOF Then
               ' If RSTEMP!maximo > "01" Then
                        ccodeaux = CStr(CInt(RSTEMP!maximo) + 1)
                        ccodeaux = String(2 - Len(ccodeaux), "0") + ccodeaux
                'End If
                        GeneraNumSorteo = cAlcance & cAnio & "-" & ccodeaux
             End If
       ' Wend
        
        Set RSTEMP.ActiveConnection = Nothing
        
        Set RSTEMP = Nothing
 
End Function

Public Sub InsertaSorteo(ByVal CnumSorteo As String, ByVal CDescripcion As String, ByVal LnTotalCtas As Long, ByVal LnTotalCanc As Long, ByVal LnTotalVig As Long, ByVal cEstado As String, ByVal dFecha As String, ByVal sMovNro As String)
Dim sSql As String
        sSql = " insert into sorteo(cNumSorteo , CDescripcion , CTotalCtas , CTotalCanc , cTotalVig ,  cEstado ,  dFecha ,  sMovnro  ) "
        sSql = sSql & " values( '" & CnumSorteo & "','" & CDescripcion & "'," & LnTotalCtas & "," & LnTotalCanc & "," & LnTotalVig & ",'" & cEstado & "', cast('" & dFecha & "' as datetime) ,  '" & sMovNro & "' ) "
        dbCmact.Execute sSql
        
          
End Sub

Public Sub ActualizaSorteo(ByVal CnumSorteo As String, ByVal CDescripcion As String, ByVal LnTotalCtas As Long, ByVal LnTotalCanc As Long, ByVal LnTotalVig As Long, ByVal cEstado As String, ByVal dFecha As String, ByVal sMovNro As String)
Dim sSql As String
    sSql = "Update sorteo set cdescripcion='" & CDescripcion & "', ctotalctas=" & LnTotalCtas & ", ctotalcanc=" & LnTotalCanc & ", ctotalvig=" & LnTotalVig & ", cestado='" & cEstado & "', dfecha= cast('" & dFecha & "' as datetime), smovnro='" & sMovNro & "'"
    sSql = sSql & " where cnumsorteo='" & CnumSorteo & "'"
    dbCmact.Execute sSql
    
End Sub



Public Sub ActualizaCuentasCanceladas(ByVal CnumSorteo As String)
Dim sSql As String
    sSql = " Update cuentasorteo set bcancelar=1  "
    sSql = sSql & " from ( Select p.cctacod from producto p join cuentasorteo c on c.cctacod=p.cctacod  "
    sSql = sSql & " where c.cnumsorteo='" & CnumSorteo & "' and p.nprdestado in (1300,1400) ) f  "
    sSql = sSql & " where cuentasorteo.cnumsorteo='" & CnumSorteo & "' and cuentasorteo.cctacod=f.cctacod "
    dbCmact.Execute sSql
    
End Sub

Public Sub ActualizaCuentasAnuladas(ByVal CnumSorteo As String)
Dim sSql As String
    sSql = " Update cuentasorteo set banular=1  "
    sSql = sSql & " from ( "
    sSql = sSql & "       Select pp.cctacod from productopersona pp  "
    sSql = sSql & "       join  rrhh r on pp.cperscod=r.cperscod "
    sSql = sSql & "       where R.NRHESTADO IN (201,401) "
    sSql = sSql & "       Union "
    sSql = sSql & "       SELECT PP.CCTACOD FROM PRODUCTOPERSONA PP "
    sSql = sSql & "       JOIN PERSRELACIONES PR   ON  PR.CPERSRELACPERSCOD=PP.CPERSCOD "
    sSql = sSql & "       JOIN RRHH RH ON RH.CPERSCOD=PR.CPERSCOD "
    sSql = sSql & "       where  RH.NRHESTADO IN (201,401) "
    sSql = sSql & " ) f  "
    sSql = sSql & " where left(cuentasorteo.cnumsorteo,6)='" & CnumSorteo & "' and cuentasorteo.cctacod=f.cctacod "
    dbCmact.Execute sSql
    
End Sub


Public Sub ActualizaCuentaImpresion(ByVal CnumSorteo As String, ByVal sCuenta As String)
Dim sSql As String
    sSql = " Update cuentasorteo set nimpresiones=isnull(nimpresiones,0)+1  "
    sSql = sSql & " where cnumsorteo='" & CnumSorteo & "' and cctacod='" & sCuenta & "'"
    dbCmact.Execute sSql
    
End Sub



Public Function ObtDatosCuentasSorteo(ByVal CnumSorteo As String) As Recordset
Dim sSql As String, RSTEMP As Recordset
  Set RSTEMP = New Recordset
    sSql = "select c.cnumsorteo, c.cctacod, c.nsaldo, pp.cperscod,p.cpersnombre,cpersidnro=(select top 1 i.cpersidnro from persid  i where i.cperscod=pp.cperscod and i.cpersidtpo='1' ) , "
    sSql = sSql & " C.NRANGOINI , C.NRANGOFIN, C.NNUMTICKETS, C.cMovNro, C.BENTREGADOS "
    sSql = sSql & " From cuentasorteo C "
    sSql = sSql & " join productopersona pp on pp.cctacod=c.cctacod and pp.nprdpersrelac=10 "
    sSql = sSql & " join persona p on p.cperscod=pp.cperscod "
    sSql = sSql & " WHERE  (bganador<>1 OR bganador is null) and (bcancelar=1 or bcancelar is null )  and (  banular=1 or banular is null)  and left(c.cnumsorteo,6)='" & CnumSorteo & "'"
    sSql = sSql & " order by cast(C.NRANGOINI as int) "
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open sSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
    If Not RSTEMP.EOF Then
            Set ObtDatosCuentasSorteo = RSTEMP
            'rstemp.Close
    End If
   Set RSTEMP = Nothing

End Function

Public Sub ActualizaCtasSorteo(ByVal CnumSorteo As String, ByVal cCtaCod As String, Optional ByVal bEntregados As Integer, Optional ByVal bGanador As Integer, Optional ByVal nNumganador As Long = 0, Optional ByVal bCancelar As Integer = 0, Optional ByVal bAnulados As Integer = 0)
Dim sSql As String, sqlaux As String
  sqlaux = IIf(nNumganador = 0, "", " , nnumganador=" & nNumganador) & IIf(bCancelar = 0, "", " ,bcancelar=" & bCancelar) & IIf(bAnulados = 0, "", " , banulaR=" & bAnulados) & IIf(bGanador = 0, "", ",bganador=" & bGanador)

    sSql = "Update cuentasorteo set  bEntregados=" & bEntregados & sqlaux
    sSql = sSql & " where left(cnumsorteo,6)='" & CnumSorteo & "' and cctacod='" & cCtaCod & "'"
    dbCmact.Execute sSql
    
End Sub
Public Sub InsertaTempPortable(ByVal CnumSorteo As String)
Dim cn As ADODB.Connection
Set cn = New ADODB.Connection
Dim sCadena As String, RSTEMP As Recordset, sSql As String

Set RSTEMP = ObtDatosCuentasSorteo(CnumSorteo)


sCadena = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\SORTEO\" & Left(CnumSorteo, 2) & "\dbpruebas.mdb;Persist Security Info=False"
cn.ConnectionString = sCadena
cn.Open
     sSql = "delete from  TEMPSORTEO"
       cn.Execute sSql

  While Not RSTEMP.EOF
    With RSTEMP
       sSql = " INSERT INTO TEMPSORTEO(CNUMSORTEO,CCTACOD,NSALDO,CPERSCOD,CTITULAR,CNRODOC,NRANGOINI,NRANGOFIN,NNUMTICKETS,CMOVNRO,BENTREGADO)"
       sSql = sSql & " values ('" & !CnumSorteo & "','" & !cCtaCod & "'," & !nSaldo & ",'" & !cPersCod & "','" & !cPersNombre & "','" & !cPersIdNro & "'," & !NRANGOINI & "," & !NRANGOFIN & "," & !NNUMTICKETS & ",'" & !cMovNro & "'," & IIf(!bEntregados, 1, 0) & ")"
       cn.Execute sSql
    End With
       RSTEMP.MoveNext
  Wend
   RSTEMP.Close
   Set RSTEMP = Nothing
   cn.Close
   Set cn = Nothing

End Sub




Public Function GetParametrosSorteo() As Recordset
Dim sSql As String, RSTEMP As Recordset

    
    sSql = " Select * from parametrosorteo "
    Set RSTEMP = New Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    Set RSTEMP.ActiveConnection = Nothing
    Set GetParametrosSorteo = RSTEMP
    
    Set RSTEMP = Nothing

End Function

Public Function GetDatosCtaSorteo(ByVal CnumSorteo As String, Optional ByVal sCuenta As String, Optional sAgecod As String) As Recordset
Dim sSql As String, RSTEMP As Recordset, sqlaux As String
sqlaux = ""

     If sCuenta <> "" Then
        sqlaux = " AND CCTACOD='" & sCuenta & "'"
     ElseIf sAgecod <> "" Then
        sqlaux = " AND substring(CCTACOD,4,2)='" & sAgecod & "'"
     End If
   
    sSql = " Select c.cperscod,p.cpersnombre,c.nrangoini,c.nrangofin,c.nnumtickets,nimpresiones=isnull(c.nimpresiones,0) from cuentasorteo c "
    sSql = sSql & " join persona p on p.cperscod=c.cperscod  WHERE left(CNUMSORTEO,6)='" & CnumSorteo & "'" & sqlaux
    sSql = sSql & " and (bganador<>1 OR bganador is null) and (bcancelar=1 or bcancelar is null )  and (  banular=1 or banular is null) "
    
    Set RSTEMP = New Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    Set RSTEMP.ActiveConnection = Nothing
   ' Set GetParametrosSorteo = rstemp
     Set GetDatosCtaSorteo = RSTEMP
    Set RSTEMP = Nothing

End Function
Public Sub ProcesarPCtasSorteoPrimer(ByVal CnumSorteo As String, ByVal nMontoxCuponS As Double, ByVal nMontoxCuponD As Double, ByVal nMinPlazoOtorgado As Long, ByVal nLimMaxOtorgadoS As Double, ByVal nLimMaxOtorgadoD As Double, ByVal cAlcance As String, ByVal cMovNro As String)
Dim sSql As String
   
        
        sSql = " exec Cap_ProcesaInfoParaSorteo '" & CnumSorteo & "'," & nMontoxCuponS & "," & nMontoxCuponD & "," & nMinPlazoOtorgado & "," & nLimMaxOtorgadoS & "," & nLimMaxOtorgadoD & ",'" & cAlcance & "','" & cMovNro & "'"
        dbCmact.Execute sSql
        
    
                
End Sub
Public Sub ProcesarPCtasSorteoGenOA(ByVal CnumSorteo As String, ByVal nMontoxCuponS As Double, ByVal nMontoxCuponD As Double, ByVal nMinPlazoOtorgado As Long, ByVal nLimMaxOtorgadoS As Double, ByVal nLimMaxOtorgadoD As Double, ByVal cAlcance As String, ByVal cMovNro As String)
Dim sSql As String
   
        
        sSql = " exec Cap_ProcesaInfoParaSorteoGenFinal '" & CnumSorteo & "'," & nMontoxCuponS & "," & nMontoxCuponD & "," & nMinPlazoOtorgado & "," & nLimMaxOtorgadoS & "," & nLimMaxOtorgadoD & ",'" & cAlcance & "','" & cMovNro & "'"
        dbCmact.Execute sSql
        
    
                
End Sub


Public Sub ProcesarPCtasSorteoOtros(ByVal CnumSorteo As String, ByVal nMontoxCuponS As Double, ByVal nMontoxCuponD As Double, ByVal nMinPlazoOtorgado As Long, ByVal nLimMaxOtorgadoS As Double, ByVal nLimMaxOtorgadoD As Double, ByVal cAlcance As String, ByVal cMovNro As String)
Dim sSql As String
    
        
        sSql = " exec Cap_ProcesaInfoParaSorteoOTROS '" & CnumSorteo & "'," & nMontoxCuponS & "," & nMontoxCuponD & "," & nMinPlazoOtorgado & "," & nLimMaxOtorgadoS & "," & nLimMaxOtorgadoD & ",'" & cAlcance & "','" & cMovNro & "'"
        dbCmact.Execute sSql
        
    
                
End Sub

Public Sub ProcesarCtaSorteoPrimer(ByVal CnumSorteo As String, ByVal nMontoxCuponS As Double, ByVal nMontoxCuponD As Double, ByVal nMinPlazoOtorgado As Long, ByVal nLimMaxOtorgadoS As Double, ByVal nLimMaxOtorgadoD As Double, ByVal cAlcance As String, ByVal cMovNro As String, ByVal CCuenta As String)
Dim sSql As String
   
        
        sSql = " exec Cap_ProcesaCtaParaSorteo '" & CnumSorteo & "'," & nMontoxCuponS & "," & nMontoxCuponD & "," & nMinPlazoOtorgado & "," & nLimMaxOtorgadoS & "," & nLimMaxOtorgadoD & ",'" & cAlcance & "','" & cMovNro & "','" & CCuenta & "'"
        dbCmact.Execute sSql
        
    
                
End Sub
Public Sub ProcesarCtaSorteoOtros(ByVal CnumSorteo As String, ByVal nMontoxCuponS As Double, ByVal nMontoxCuponD As Double, ByVal nMinPlazoOtorgado As Long, ByVal nLimMaxOtorgadoS As Double, ByVal nLimMaxOtorgadoD As Double, ByVal cAlcance As String, ByVal cMovNro As String, ByVal CCuenta As String)
Dim sSql As String
    
        
        sSql = " exec Cap_ProcesaCtaParaSorteoOTROS '" & CnumSorteo & "'," & nMontoxCuponS & "," & nMontoxCuponD & "," & nMinPlazoOtorgado & "," & nLimMaxOtorgadoS & "," & nLimMaxOtorgadoD & ",'" & cAlcance & "','" & cMovNro & "','" & CCuenta & "'"
        dbCmact.Execute sSql
        
    
                
End Sub




Public Function GetSorteoEstados(ByVal cAlcance As String, ByVal cAnio As String, ByVal cEstado As String) As Boolean
Dim sSql As String, RSTEMP As Recordset

GetSorteoEstados = False

   sSql = "Select cantidad=count(*) from sorteo where left(cnumsorteo,2)='" & cAlcance & "' and substring(cnumsorteo,3,4)='" & cAnio & "' and cestado='" & cEstado & "'"

      Set RSTEMP = New Recordset
      RSTEMP.CursorLocation = adUseClient
      RSTEMP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
      Set RSTEMP.ActiveConnection = Nothing
    
    If Not RSTEMP.EOF Then
       If RSTEMP!cantidad > 0 Then
             GetSorteoEstados = True
       End If
    End If
    
    Set RSTEMP = Nothing


End Function



Public Function GetSorteo(Optional ByVal cEstado As String = "", Optional ByVal cAlcance = "00", Optional ByVal bGeneral As Boolean = False) As Recordset
 Dim sSql As String, RSTEMP As Recordset
 
 
 Set RSTEMP = New Recordset
 
 If cEstado <> "" And cAlcance <> "" Then
    sSql = " SELECT s.cnumsorteo,dFecha=convert(char(10),s.dfecha,103),dHora=convert(char(10),s.dfecha,108), "
    sSql = sSql & " sestado=case when s.cestado='I' then 'INICIADO' "
    sSql = sSql & " WHEN s.CESTADO='P' THEN 'PROCESADO' "
    sSql = sSql & " WHEN s.CESTADO='C' THEN 'CERRADO' END, s.CDESCRIPCION, "
    sSql = sSql & " cAlcance = Left(s.CNUMSORTEO, 2) , NroCuentas=(Select count(c.cctacod) from  cuentasorteo c where c.cnumsorteo=s.cnumsorteo ), "
    sSql = sSql & " NroRangoMax=(Select max(cast(c.nrangofin as int)) from  cuentasorteo c where c.cnumsorteo=s.cnumsorteo ) "
    sSql = sSql & " from sorteo s where s.cestado='" & cEstado & "' and left(s.cnumsorteo,2)='" & cAlcance & "'"
    
 ElseIf cEstado <> "" And cAlcance = "" And bGeneral = False Then
    
    sSql = "SELECT cnumsorteo,dFecha=convert(char(10),dfecha,103),dHora=convert(char(10),dfecha,108),    "
    sSql = sSql & " sestado=case when cestado='I' then 'INICIADO' "
    sSql = sSql & " WHEN CESTADO='P' THEN 'PROCESADO' "
    sSql = sSql & " WHEN CESTADO='C' THEN 'CERRADO' END,CDESCRIPCION, "
    sSql = sSql & "  cAlcance = Left(CNUMSORTEO, 2) "
    sSql = sSql & " From sorteo s where cestado='" & cEstado & "'"
    
 ElseIf cEstado <> "" And cAlcance = "" And bGeneral = True Then
 
    sSql = "SELECT distinct cnumsorteo=left(cnumsorteo,6) ,   "
    sSql = sSql & " sestado=case when cestado='I' then 'INICIADO' "
    sSql = sSql & " WHEN CESTADO='P' THEN 'PROCESADO' "
    sSql = sSql & " WHEN CESTADO='C' THEN 'CERRADO' END,CDESCRIPCION, "
    sSql = sSql & " cAlcance = Left(CNUMSORTEO, 2) "
    sSql = sSql & " From sorteo s where cestado='" & cEstado & "' and left(s.cnumsorteo,2)<>'00' "
    
    
 ElseIf cEstado = "" And cAlcance <> "" Then
    sSql = "SELECT cnumsorteo,dFecha=convert(char(10),dfecha,103),dHora=convert(char(10),dfecha,108),    "
    sSql = sSql & " sestado=case when cestado='I' then 'INICIADO' "
    sSql = sSql & " WHEN CESTADO='P' THEN 'PROCESADO' "
    sSql = sSql & " WHEN CESTADO='C' THEN 'CERRADO' END,CDESCRIPCION,"
    sSql = sSql & " cAlcance = Left(CNUMSORTEO, 2) "
    sSql = sSql & " from sorteo s where left(cnumsorteo,2)='" & cAlcance & "'"
    
 ElseIf cEstado = "" And cAlcance = "" Then
    sSql = "SELECT cnumsorteo,dFecha=convert(char(10),dfecha,103),dHora=convert(char(10),dfecha,108),    "
    sSql = sSql & " sestado=case when cestado='I' then 'INICIADO' "
    sSql = sSql & " WHEN CESTADO='P' THEN 'PROCESADO' "
    sSql = sSql & " WHEN CESTADO='C' THEN 'CERRADO' END,CDESCRIPCION, "
    sSql = sSql & " cAlcance = Left(CNUMSORTEO, 2) "
    sSql = sSql & " from sorteo s "
    
 End If
      
      Set RSTEMP = New Recordset
      RSTEMP.CursorLocation = adUseClient
      RSTEMP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    Set RSTEMP.ActiveConnection = Nothing
    
    Set GetSorteo = RSTEMP
    
    Set RSTEMP = Nothing
    

End Function

Public Function GetSorteos(ByVal Anio As String, Optional ByVal cAlcance As String = "") As Recordset
 Dim sSql As String, RSTEMP As Recordset, sqlaux As String
 Set RSTEMP = New Recordset
 sqlaux = ""
 If cAlcance <> "" Then
    sqlaux = " and left(CNUMSORTEO,2)='" & cAlcance & "'"
 End If
 
 sSql = " SELECT CNUMSORTEO FROM SORTEO WHERE substring(CNUMSORTEO,3,4)='" & Anio & "'" & sqlaux
 
    Set RSTEMP = New Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    Set RSTEMP.ActiveConnection = Nothing
    
    Set GetSorteos = RSTEMP
    
    Set RSTEMP = Nothing
    


End Function

Private Sub Class_Initialize()
Dim sConn As String
Dim ClsIni As COMConecta.DCOMClasIni
Set ClsIni = New COMConecta.DCOMClasIni
sConn = ClsIni.CadenaConexion
sDBComunes = ClsIni.BaseComunes
sDBPersona = ClsIni.BasePersonas
sDBImagenes = ClsIni.BaseImagenes
Set ClsIni = Nothing
Set dbCmact = New Connection
dbCmact.CommandTimeout = 10000
dbCmact.Open sConn
dbCmact.Execute "SET DATEFORMAT MDY"
End Sub

Private Sub Class_Terminate()
dbCmact.Close
Set dbCmact = Nothing
End Sub
