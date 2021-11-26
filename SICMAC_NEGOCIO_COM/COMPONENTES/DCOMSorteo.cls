VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMSorteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Public dbCmact As COMConecta.DCOMConecta
Public dbCmact As ADODB.Connection
Dim sDBComunes As String
Dim sDBPersona As String
Dim sDBImagenes As String

Public Function GeneraNumSorteo(ByVal cAlcance As String, ByVal cAnio As String) As String
 Dim ssql As String, rstemp As Recordset, ccodeaux
 
 ccodeaux = "01"
 GeneraNumSorteo = ""
 
        ssql = "Select maximo=isnull(max(right(cnumsorteo,2)),'00') from sorteo where substring(cnumsorteo,3,4)='" & cAnio & "' and  left(cnumsorteo,2)='" & cAlcance & "'"
        Set rstemp = New Recordset
        rstemp.CursorLocation = adUseClient
        rstemp.Open ssql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
       ' With rstemp
            If Not rstemp.EOF Then
               ' If RSTEMP!maximo > "01" Then
                        ccodeaux = CStr(CInt(rstemp!maximo) + 1)
                        ccodeaux = String(2 - Len(ccodeaux), "0") + ccodeaux
                'End If
                        GeneraNumSorteo = cAlcance & cAnio & "-" & ccodeaux
             End If
       ' Wend
        
        Set rstemp.ActiveConnection = Nothing
        
        Set rstemp = Nothing
 
End Function

Public Sub InsertaSorteo(ByVal CnumSorteo As String, ByVal CDescripcion As String, ByVal LnTotalCtas As Long, ByVal LnTotalCanc As Long, ByVal LnTotalVig As Long, ByVal cEstado As String, ByVal dFecha As String, ByVal sMovNro As String)
Dim ssql As String
        ssql = " insert into sorteo(cNumSorteo , CDescripcion , CTotalCtas , CTotalCanc , cTotalVig ,  cEstado ,  dFecha ,  sMovnro  ) "
        ssql = ssql & " values( '" & CnumSorteo & "','" & CDescripcion & "'," & LnTotalCtas & "," & LnTotalCanc & "," & LnTotalVig & ",'" & cEstado & "', cast('" & dFecha & "' as datetime) ,  '" & sMovNro & "' ) "
         dbCmact.Execute ssql
        
          
End Sub

Public Sub ActualizaSorteo(ByVal CnumSorteo As String, ByVal CDescripcion As String, ByVal LnTotalCtas As Long, ByVal LnTotalCanc As Long, ByVal LnTotalVig As Long, ByVal cEstado As String, ByVal dFecha As String, ByVal sMovNro As String)
Dim ssql As String
    ssql = "Update sorteo set cdescripcion='" & CDescripcion & "', ctotalctas=" & LnTotalCtas & ", ctotalcanc=" & LnTotalCanc & ", ctotalvig=" & LnTotalVig & ", cestado='" & cEstado & "', dfecha= cast('" & dFecha & "' as datetime), smovnro='" & sMovNro & "'"
    ssql = ssql & " where cnumsorteo='" & CnumSorteo & "'"
     dbCmact.Execute ssql
    
End Sub



Public Sub ActualizaCuentasCanceladas(ByVal CnumSorteo As String)
Dim ssql As String
    ssql = " Update cuentasorteo set bcancelar=1  "
    ssql = ssql & " from ( Select p.cctacod from producto p join cuentasorteo c on c.cctacod=p.cctacod  "
    ssql = ssql & " where c.cnumsorteo='" & CnumSorteo & "' and p.nprdestado in (1300,1400) ) f  "
    ssql = ssql & " where cuentasorteo.cnumsorteo='" & CnumSorteo & "' and cuentasorteo.cctacod=f.cctacod "
     dbCmact.Execute ssql
    
End Sub

Public Sub ActualizaCuentasAnuladas(ByVal CnumSorteo As String)
Dim ssql As String
    ssql = " Update cuentasorteo set banular=1  "
    ssql = ssql & " from ( "
    ssql = ssql & "       Select pp.cctacod from productopersona pp  "
    ssql = ssql & "       join  rrhh r on pp.cperscod=r.cperscod "
    ssql = ssql & "       where R.NRHESTADO IN (201,401) "
    ssql = ssql & "       Union "
    ssql = ssql & "       SELECT PP.CCTACOD FROM PRODUCTOPERSONA PP "
    ssql = ssql & "       JOIN PERSRELACIONES PR   ON  PR.CPERSRELACPERSCOD=PP.CPERSCOD "
    ssql = ssql & "       JOIN RRHH RH ON RH.CPERSCOD=PR.CPERSCOD "
    ssql = ssql & "       where  RH.NRHESTADO IN (201,401) "
    ssql = ssql & " ) f  "
    ssql = ssql & " where left(cuentasorteo.cnumsorteo,6)='" & CnumSorteo & "' and cuentasorteo.cctacod=f.cctacod "
     dbCmact.Execute ssql
    
End Sub


Public Sub ActualizaCuentaImpresion(ByVal CnumSorteo As String, ByVal sCuenta As String)
Dim ssql As String
    ssql = " Update cuentasorteo set nimpresiones=isnull(nimpresiones,0)+1  "
    ssql = ssql & " where cnumsorteo='" & CnumSorteo & "' and cctacod='" & sCuenta & "'"
     dbCmact.Execute ssql
    
End Sub



Public Function ObtDatosCuentasSorteo(ByVal CnumSorteo As String) As Recordset
Dim ssql As String, rstemp As Recordset
  Set rstemp = New Recordset
    ssql = "select c.cnumsorteo, c.cctacod, c.nsaldo, pp.cperscod,p.cpersnombre,cpersidnro=(select top 1 i.cpersidnro from persid  i where i.cperscod=pp.cperscod and i.cpersidtpo='1' ) , "
    ssql = ssql & " C.NRANGOINI , C.NRANGOFIN, C.NNUMTICKETS, C.cMovNro, C.BENTREGADOS "
    ssql = ssql & " From cuentasorteo C "
    ssql = ssql & " join productopersona pp on pp.cctacod=c.cctacod and pp.nprdpersrelac=10 "
    ssql = ssql & " join persona p on p.cperscod=pp.cperscod "
    ssql = ssql & " WHERE  (bganador<>1 OR bganador is null) and (bcancelar=1 or bcancelar is null )  and (  banular=1 or banular is null)  and left(c.cnumsorteo,6)='" & CnumSorteo & "'"
    ssql = ssql & " order by cast(C.NRANGOINI as int) "
    rstemp.CursorLocation = adUseClient
    rstemp.Open ssql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
    If Not rstemp.EOF Then
            Set ObtDatosCuentasSorteo = rstemp
            'rstemp.Close
    End If
   Set rstemp = Nothing

End Function

Public Sub ActualizaCtasSorteo(ByVal CnumSorteo As String, ByVal cCtaCod As String, Optional ByVal bEntregados As Integer, Optional ByVal bGanador As Integer, Optional ByVal nNumganador As Long = 0, Optional ByVal bCancelar As Integer = 0, Optional ByVal bAnulados As Integer = 0)
Dim ssql As String, sqlaux As String
  sqlaux = IIf(nNumganador = 0, "", " , nnumganador=" & nNumganador) & IIf(bCancelar = 0, "", " ,bcancelar=" & bCancelar) & IIf(bAnulados = 0, "", " , banulaR=" & bAnulados) & IIf(bGanador = 0, "", ",bganador=" & bGanador)

    ssql = "Update cuentasorteo set  bEntregados=" & bEntregados & sqlaux
    ssql = ssql & " where left(cnumsorteo,6)='" & CnumSorteo & "' and cctacod='" & cCtaCod & "'"
     dbCmact.Execute ssql
    
End Sub
Public Sub InsertaTempPortable(ByVal CnumSorteo As String)
Dim cn As ADODB.Connection
Set cn = New ADODB.Connection
Dim sCadena As String, rstemp As Recordset, ssql As String

Set rstemp = ObtDatosCuentasSorteo(CnumSorteo)


sCadena = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\SORTEO\" & Left(CnumSorteo, 2) & "\dbpruebas.mdb;Persist Security Info=False"
cn.ConnectionString = sCadena
cn.Open
     ssql = "delete from  TEMPSORTEO"
       cn.Execute ssql

  While Not rstemp.EOF
    With rstemp
       ssql = " INSERT INTO TEMPSORTEO(CNUMSORTEO,CCTACOD,NSALDO,CPERSCOD,CTITULAR,CNRODOC,NRANGOINI,NRANGOFIN,NNUMTICKETS,CMOVNRO,BENTREGADO)"
       ssql = ssql & " values ('" & !CnumSorteo & "','" & !cCtaCod & "'," & !nSaldo & ",'" & !cPersCod & "','" & !cpersnombre & "','" & !cPersIdNro & "'," & !NRANGOINI & "," & !NRANGOFIN & "," & !NNUMTICKETS & ",'" & !cMovnro & "'," & IIf(!bEntregados, 1, 0) & ")"
       cn.Execute ssql
    End With
       rstemp.MoveNext
  Wend
   rstemp.Close
   Set rstemp = Nothing
   cn.Close
   Set cn = Nothing

End Sub




Public Function GetParametrosSorteo() As Recordset
Dim ssql As String, rstemp As Recordset

    
    ssql = " Select * from parametrosorteo "
    Set rstemp = New Recordset
    rstemp.CursorLocation = adUseClient
    rstemp.Open ssql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    Set rstemp.ActiveConnection = Nothing
    Set GetParametrosSorteo = rstemp
    
    Set rstemp = Nothing

End Function

Public Function GetDatosCtaSorteo(ByVal CnumSorteo As String, Optional ByVal sCuenta As String, Optional sAgecod As String) As Recordset
Dim ssql As String, rstemp As Recordset, sqlaux As String
sqlaux = ""

     If sCuenta <> "" Then
        sqlaux = " AND CCTACOD='" & sCuenta & "'"
     ElseIf sAgecod <> "" Then
        sqlaux = " AND substring(CCTACOD,4,2)='" & sAgecod & "'"
     End If
   
    ssql = " Select c.cperscod,p.cpersnombre,c.nrangoini,c.nrangofin,c.nnumtickets,nimpresiones=isnull(c.nimpresiones,0) from cuentasorteo c "
    ssql = ssql & " join persona p on p.cperscod=c.cperscod  WHERE left(CNUMSORTEO,6)='" & CnumSorteo & "'" & sqlaux
    ssql = ssql & " and (bganador<>1 OR bganador is null) and (bcancelar=1 or bcancelar is null )  and (  banular=1 or banular is null) "
    
    Set rstemp = New Recordset
    rstemp.CursorLocation = adUseClient
    rstemp.Open ssql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    Set rstemp.ActiveConnection = Nothing
   ' Set GetParametrosSorteo = rstemp
     Set GetDatosCtaSorteo = rstemp
    Set rstemp = Nothing

End Function
Public Sub ProcesarPCtasSorteoPrimer(ByVal CnumSorteo As String, ByVal nMontoxCuponS As Double, ByVal nMontoxCuponD As Double, ByVal nMinPlazoOtorgado As Long, ByVal nLimMaxOtorgadoS As Double, ByVal nLimMaxOtorgadoD As Double, ByVal cAlcance As String, ByVal cMovnro As String)
Dim ssql As String
   
        
        ssql = " exec Cap_ProcesaInfoParaSorteo '" & CnumSorteo & "'," & nMontoxCuponS & "," & nMontoxCuponD & "," & nMinPlazoOtorgado & "," & nLimMaxOtorgadoS & "," & nLimMaxOtorgadoD & ",'" & cAlcance & "','" & cMovnro & "'"
        dbCmact.Execute ssql
        
    
                
End Sub
Public Sub ProcesarPCtasSorteoGenOA(ByVal CnumSorteo As String, ByVal nMontoxCuponS As Double, ByVal nMontoxCuponD As Double, ByVal nMinPlazoOtorgado As Long, ByVal nLimMaxOtorgadoS As Double, ByVal nLimMaxOtorgadoD As Double, ByVal cAlcance As String, ByVal cMovnro As String)
Dim ssql As String
   
        
        ssql = " exec Cap_ProcesaInfoParaSorteoGenFinal '" & CnumSorteo & "'," & nMontoxCuponS & "," & nMontoxCuponD & "," & nMinPlazoOtorgado & "," & nLimMaxOtorgadoS & "," & nLimMaxOtorgadoD & ",'" & cAlcance & "','" & cMovnro & "'"
         dbCmact.Execute ssql
        
    
                
End Sub


Public Sub ProcesarPCtasSorteoOtros(ByVal CnumSorteo As String, ByVal nMontoxCuponS As Double, ByVal nMontoxCuponD As Double, ByVal nMinPlazoOtorgado As Long, ByVal nLimMaxOtorgadoS As Double, ByVal nLimMaxOtorgadoD As Double, ByVal cAlcance As String, ByVal cMovnro As String)
Dim ssql As String
    
        
        ssql = " exec Cap_ProcesaInfoParaSorteoOTROS '" & CnumSorteo & "'," & nMontoxCuponS & "," & nMontoxCuponD & "," & nMinPlazoOtorgado & "," & nLimMaxOtorgadoS & "," & nLimMaxOtorgadoD & ",'" & cAlcance & "','" & cMovnro & "'"
        dbCmact.Execute ssql
        
    
                
End Sub

Public Sub ProcesarCtaSorteoPrimer(ByVal CnumSorteo As String, ByVal nMontoxCuponS As Double, ByVal nMontoxCuponD As Double, ByVal nMinPlazoOtorgado As Long, ByVal nLimMaxOtorgadoS As Double, ByVal nLimMaxOtorgadoD As Double, ByVal cAlcance As String, ByVal cMovnro As String, ByVal CCuenta As String)
Dim ssql As String
   
        
        ssql = " exec Cap_ProcesaCtaParaSorteo '" & CnumSorteo & "'," & nMontoxCuponS & "," & nMontoxCuponD & "," & nMinPlazoOtorgado & "," & nLimMaxOtorgadoS & "," & nLimMaxOtorgadoD & ",'" & cAlcance & "','" & cMovnro & "','" & CCuenta & "'"
        dbCmact.Execute ssql
        
    
                
End Sub
Public Sub ProcesarCtaSorteoOtros(ByVal CnumSorteo As String, ByVal nMontoxCuponS As Double, ByVal nMontoxCuponD As Double, ByVal nMinPlazoOtorgado As Long, ByVal nLimMaxOtorgadoS As Double, ByVal nLimMaxOtorgadoD As Double, ByVal cAlcance As String, ByVal cMovnro As String, ByVal CCuenta As String)
Dim ssql As String
    
        
        ssql = " exec Cap_ProcesaCtaParaSorteoOTROS '" & CnumSorteo & "'," & nMontoxCuponS & "," & nMontoxCuponD & "," & nMinPlazoOtorgado & "," & nLimMaxOtorgadoS & "," & nLimMaxOtorgadoD & ",'" & cAlcance & "','" & cMovnro & "','" & CCuenta & "'"
        dbCmact.Execute ssql
        
    
                
End Sub

Public Function GetSorteoEstados(ByVal cAlcance As String, ByVal cAnio As String, ByVal cEstado As String) As Boolean
Dim ssql As String, rstemp As Recordset

GetSorteoEstados = False

   ssql = "Select cantidad=count(*) from sorteo where left(cnumsorteo,2)='" & cAlcance & "' and substring(cnumsorteo,3,4)='" & cAnio & "' and cestado='" & cEstado & "'"

      Set rstemp = New Recordset
      rstemp.CursorLocation = adUseClient
      rstemp.Open ssql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
      Set rstemp.ActiveConnection = Nothing
    
    If Not rstemp.EOF Then
       If rstemp!cantidad > 0 Then
             GetSorteoEstados = True
       End If
    End If
    
    Set rstemp = Nothing


End Function



Public Function GetSorteo(Optional ByVal cEstado As String = "", Optional ByVal cAlcance = "00", Optional ByVal bGeneral As Boolean = False) As Recordset
 Dim ssql As String, rstemp As Recordset
 
 
 Set rstemp = New Recordset
 
 If cEstado <> "" And cAlcance <> "" Then
    ssql = " SELECT s.cnumsorteo,dFecha=convert(char(10),s.dfecha,103),dHora=convert(char(10),s.dfecha,108), "
    ssql = ssql & " sestado=case when s.cestado='I' then 'INICIADO' "
    ssql = ssql & " WHEN s.CESTADO='P' THEN 'PROCESADO' "
    ssql = ssql & " WHEN s.CESTADO='C' THEN 'CERRADO' END, s.CDESCRIPCION, "
    ssql = ssql & " cAlcance = Left(s.CNUMSORTEO, 2) , NroCuentas=(Select count(c.cctacod) from  cuentasorteo c where c.cnumsorteo=s.cnumsorteo ), "
    ssql = ssql & " NroRangoMax=(Select max(cast(c.nrangofin as int)) from  cuentasorteo c where c.cnumsorteo=s.cnumsorteo ) "
    ssql = ssql & " from sorteo s where s.cestado='" & cEstado & "' and left(s.cnumsorteo,2)='" & cAlcance & "'"
    
 ElseIf cEstado <> "" And cAlcance = "" And bGeneral = False Then
    
    ssql = "SELECT cnumsorteo,dFecha=convert(char(10),dfecha,103),dHora=convert(char(10),dfecha,108),    "
    ssql = ssql & " sestado=case when cestado='I' then 'INICIADO' "
    ssql = ssql & " WHEN CESTADO='P' THEN 'PROCESADO' "
    ssql = ssql & " WHEN CESTADO='C' THEN 'CERRADO' END,CDESCRIPCION, "
    ssql = ssql & "  cAlcance = Left(CNUMSORTEO, 2) "
    ssql = ssql & " From sorteo s where cestado='" & cEstado & "'"
    
 ElseIf cEstado <> "" And cAlcance = "" And bGeneral = True Then
 
    ssql = "SELECT distinct cnumsorteo=left(cnumsorteo,6) ,   "
    ssql = ssql & " sestado=case when cestado='I' then 'INICIADO' "
    ssql = ssql & " WHEN CESTADO='P' THEN 'PROCESADO' "
    ssql = ssql & " WHEN CESTADO='C' THEN 'CERRADO' END,CDESCRIPCION, "
    ssql = ssql & " cAlcance = Left(CNUMSORTEO, 2) "
    ssql = ssql & " From sorteo s where cestado='" & cEstado & "' and left(s.cnumsorteo,2)<>'00' "
    
    
 ElseIf cEstado = "" And cAlcance <> "" Then
    ssql = "SELECT cnumsorteo,dFecha=convert(char(10),dfecha,103),dHora=convert(char(10),dfecha,108),    "
    ssql = ssql & " sestado=case when cestado='I' then 'INICIADO' "
    ssql = ssql & " WHEN CESTADO='P' THEN 'PROCESADO' "
    ssql = ssql & " WHEN CESTADO='C' THEN 'CERRADO' END,CDESCRIPCION,"
    ssql = ssql & " cAlcance = Left(CNUMSORTEO, 2) "
    ssql = ssql & " from sorteo s where left(cnumsorteo,2)='" & cAlcance & "'"
    
 ElseIf cEstado = "" And cAlcance = "" Then
    ssql = "SELECT cnumsorteo,dFecha=convert(char(10),dfecha,103),dHora=convert(char(10),dfecha,108),    "
    ssql = ssql & " sestado=case when cestado='I' then 'INICIADO' "
    ssql = ssql & " WHEN CESTADO='P' THEN 'PROCESADO' "
    ssql = ssql & " WHEN CESTADO='C' THEN 'CERRADO' END,CDESCRIPCION, "
    ssql = ssql & " cAlcance = Left(CNUMSORTEO, 2) "
    ssql = ssql & " from sorteo s "
    
 End If
      
      Set rstemp = New Recordset
      rstemp.CursorLocation = adUseClient
      rstemp.Open ssql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    Set rstemp.ActiveConnection = Nothing
    
    Set GetSorteo = rstemp
    
    Set rstemp = Nothing
    

End Function

Public Function GetSorteos(ByVal Anio As String, Optional ByVal cAlcance As String = "") As Recordset
 Dim ssql As String, rstemp As Recordset, sqlaux As String
 Set rstemp = New Recordset
 sqlaux = ""
 If cAlcance <> "" Then
    sqlaux = " and left(CNUMSORTEO,2)='" & cAlcance & "'"
 End If
 
 ssql = " SELECT CNUMSORTEO FROM SORTEO WHERE substring(CNUMSORTEO,3,4)='" & Anio & "'" & sqlaux
 
    Set rstemp = New Recordset
    rstemp.CursorLocation = adUseClient
    rstemp.Open ssql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    Set rstemp.ActiveConnection = Nothing
    
    Set GetSorteos = rstemp
    
    Set rstemp = Nothing
    


End Function

Private Sub Class_Initialize()
Dim sConn As String
Dim ClsIni As New COMConecta.DCOMClasIni
sConn = ClsIni.CadenaConexion
sDBComunes = ClsIni.BaseComunes
sDBPersona = ClsIni.BasePersonas
sDBImagenes = ClsIni.BaseImagenes
Set ClsIni = Nothing
Set dbCmact = New ADODB.Connection
dbCmact.CommandTimeout = 10000
dbCmact.Open sConn
dbCmact.Execute "SET DATEFORMAT MDY"
End Sub

Private Sub Class_Terminate()
dbCmact.Close
Set dbCmact = Nothing
End Sub

