VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COMDCaptAutorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function InfoAutorizacion(ByVal cUser As String, ByVal cOpeCod As String) As ADODB.Recordset
 Dim sSql As String
 Dim oConecta As COMConecta.DCOMConecta
 
      sSql = "Select p.cperscod,p.cpersnombre,p.dpersnaccreac,c.cnivel,rhct.crhcargodescripcion,rhc.crhcargocod,nMontoIniSol,nMontoFinSol, nMontoIniDol,nMontoFinDol  "
      sSql = sSql & "  from persona p"
      sSql = sSql & "  inner join rrhh  rh on rh.cperscod=p.cperscod"
      sSql = sSql & "  inner join (Select top 1 rhc1.crhcargocod,rhc1.cperscod from rhcargos rhc1 inner join rrhh rh1 on rh1.cperscod=rhc1.cperscod "
      sSql = sSql & "  where rh1.cuser='" & cUser & "' order by drhcargoFecha desc) rhc on rhc.cperscod=p.cperscod"
      sSql = sSql & "  inner join rhcargostabla rhct on rhct.crhcargocod=rhc.crhcargocod "
      sSql = sSql & "  inner join capautorizacionrango c on c.crhcargocod=rhc.crhcargocod "
      sSql = sSql & "  where  rh.nrhestado='201' and rh.cUser='" & cUser & "' and c.copecod='" & cOpeCod & "'"
        
      Set oConecta = New COMConecta.DCOMConecta
      oConecta.AbreConexion
      Set InfoAutorizacion = oConecta.CargaRecordSet(sSql)
      oConecta.CierraConexion
      Set oConecta = Nothing
      
End Function
Public Function ExisteAutrango(ByVal cOpeCod As String, ByVal cRHCargoCod As String) As Boolean
Dim sSql As String, rs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim oError As New COMConecta.COMErrorHandling

    ExisteAutrango = False
   On Error GoTo MensaError
            
        sSql = "Select copecod+crhcargocod from capautorizacionrango where crhcargocod='" & cRHCargoCod & "' and copecod='" & cOpeCod & "'"
        
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set rs = oConecta.Ejecutar(sSql)
        If Not (rs.EOF Or rs.BOF) Then
               If Trim(rs.Fields(0).Value) = Trim(Trim(cOpeCod) & Trim(cRHCargoCod)) Then
                  ExisteAutrango = True
               Else
                  ExisteAutrango = False
               End If
        End If
           If rs.State = 1 Then rs.Close
           Set rs = Nothing
           oConecta.CierraConexion
           Set oConecta = Nothing
        Exit Function
   
MensaError:
        Call oError.RaiseError(oError.MyUnhandledError, "DAutorizacion:ExisteAutRango Method")
End Function
Public Function CargaFlexAutRango() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
      sSql = "Select c.copecod,Upper(o.operacion),rct.crhcargodescripcion + space(150-len(rct.crhcargodescripcion))+rct.crhcargocod Cargo,c.cnivel,c.nMontoIniSol,c.nMontoFinSol,c.nMontoIniDol,nMontoFinDol,'E' as Estado,cast(case when c.bactivo=1 then 1 else 0 end as char(1)) as bactivo "
      sSql = sSql & " from capautorizacionrango c "
      sSql = sSql & " inner join  rhcargostabla  rct on rct.crhcargocod=c.crhcargocod "
      sSql = sSql & " inner join  (Select ch.copecod,f.copedesc +':'+ ch.copedesc as operacion from opetpo ch "
      sSql = sSql & " inner join opetpo f on f.copecod=left(ch.copecod,2)+'0000' "
      sSql = sSql & " where ch.copecod like  '2%[1-9]00' ) o on o.copecod=c.copecod "
      sSql = sSql & "   Where Len(rct.cRHCargoCod) > 3 "
      sSql = sSql & " order by c.copecod "
      Set oConecta = New COMConecta.DCOMConecta
      oConecta.AbreConexion
      Set CargaFlexAutRango = oConecta.CargaRecordSet(sSql)
      oConecta.CierraConexion
      Set oConecta = Nothing
End Function

Public Function Iautrango(ByVal cOpeCod As String, ByVal cNivel As String, ByVal cRHCargoCod As String, ByVal nMontoIniSol As Double, ByVal nMontoFinSol As Double, ByVal nMontoIniDol As Double, ByVal nMontoFinDol As Double, ByVal cUltimaActualizacion As String) As Boolean
Dim sSql As String, rs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim oError As New COMConecta.COMErrorHandling
 Iautrango = False

On Error GoTo MensaErr
        sSql = "exec Cap_ManCapAutorizacionRango 'N','" & cOpeCod & "','" & cNivel & "','" & cRHCargoCod & "'," & nMontoIniSol & "," & nMontoFinSol & "," & nMontoIniDol & "," & nMontoFinDol & ",'" & cUltimaActualizacion & "'"
       
        Set oConecta = New COMConecta.DCOMConecta
           oConecta.AbreConexion
           Set rs = oConecta.Ejecutar(sSql)
           If Not (rs.EOF Or rs.BOF) Then
               If rs.Fields(0).Value = 1 Then
                  Iautrango = True
               Else
                  Iautrango = False
               End If
           End If
           If rs.State = 1 Then rs.Close
           Set rs = Nothing
           oConecta.CierraConexion
           Set oConecta = Nothing
        Exit Function
MensaErr:

   Call oError.RaiseError(oError.MyUnhandledError, "DAutorizacion:IAutRango Method")

End Function

Public Function AAutRango(ByVal cOpeCod As String, ByVal cNivel As String, ByVal cRHCargoCod As String, ByVal nMontoIniSol As Double, ByVal nMontoFinSol As Double, ByVal nMontoIniDol As Double, ByVal nMontoFinDol As Double, ByVal cUltimaActualizacion As String) As Boolean
    Dim sSql As String, rs As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim oError As New COMConecta.COMErrorHandling
    AAutRango = False
        
On Error GoTo MensaErr

        sSql = "exec Cap_ManCapAutorizacionRango 'A','" & cOpeCod & "','" & cNivel & "','" & cRHCargoCod & "'," & nMontoIniSol & "," & nMontoFinSol & "," & nMontoIniDol & "," & nMontoFinDol & ",'" & cUltimaActualizacion & "'"
        Set oConecta = New COMConecta.DCOMConecta
           oConecta.AbreConexion
           Set rs = oConecta.Ejecutar(sSql)
           If Not (rs.EOF Or rs.BOF) Then
               If rs.Fields(0).Value = 1 Then
                  AAutRango = True
               Else
                  AAutRango = False
               End If
           End If
           If rs.State = 1 Then rs.Close
           Set rs = Nothing
           oConecta.CierraConexion
           Set oConecta = Nothing
        Exit Function
MensaErr:

   Call oError.RaiseError(oError.MyUnhandledError, "DAutorizacion:IAutRango Method")
End Function


Public Function IMovAutorizacion(ByVal nIdAut As Long, ByVal cCtaCod As String, ByVal cPersCodCli As String, ByVal cOpeCod As String, ByVal cOpeCodOri As String, _
 nMontoSolicitado As Double, nMontoAprobado As Double, nMoneda As Integer, nAutEstado As String, _
cAutObs As String, cUserOri As String, dFechaOri As Date, cUserApro As String, dFechaAprob As Date, cMovnro As String, cUltimaActualizacion As String) As Long
Dim sSql As String, rs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim oError As New COMConecta.COMErrorHandling

      IMovAutorizacion = 0
      
      On Error GoTo MensaErr
      
        sSql = "exec Cap_ManMovAutorizacion_sp 0,'" & cCtaCod & "','" & cPersCodCli & "','" & cOpeCod & "','" & cOpeCodOri & "',"
        sSql = sSql & nMontoSolicitado & "," & nMontoAprobado & "," & nMoneda & ",'" & nAutEstado & "', "
        sSql = sSql & "'" & cAutObs & "','" & cUserOri & "','" & Format(dFechaOri, "yyyy-MM-dd") & "','" & cUserApro & "','" & Format(dFechaAprob, "yyyy-MM-dd") & "','" & cMovnro & "','" & cUltimaActualizacion & "'"
        
           Set oConecta = New COMConecta.DCOMConecta
           Set rs = New ADODB.Recordset
           oConecta.AbreConexion
           Set rs = oConecta.ConexionActiva.Execute(sSql)
           If rs.State = 1 Then
                If Not (rs.EOF Or rs.BOF) Then
                     IMovAutorizacion = rs.Fields(0).Value
                 End If
                 rs.Close
           End If
           Set rs = Nothing
           oConecta.CierraConexion
           Set oConecta = Nothing

      
      Exit Function
      
MensaErr:
      Call oError.RaiseError(oError.MyUnhandledError, "DAutorizacion:IMovAutorizacion Method")
End Function
Public Function AMovAutorizacionApro(ByVal nIdAut As Long, ByVal cCtaCod As String, ByVal cPersCodCli As String, ByVal cOpeCod As String, ByVal cOpeCodOri As String, _
 nMontoSolicitado As Double, nMontoAprobado As Double, nMoneda As Integer, nAutEstado As String, _
cAutObs As String, cUserOri As String, dFechaOri As Date, cUserApro As String, dFechaAprob As Date, cMovnro As String, cUltimaActualizacion As String) As Boolean
Dim sSql As String, rs As ADODB.Recordset
  Dim oConecta As COMConecta.DCOMConecta
  Dim oError As New COMConecta.COMErrorHandling
  
      AMovAutorizacionApro = False
      
      On Error GoTo MensaErr
        sSql = "exec Cap_ManMovAutorizacion_sp " & nIdAut & ",'" & cCtaCod & "','" & cPersCodCli & "','" & cOpeCod & "','" & cOpeCodOri & "',"
        sSql = sSql & nMontoSolicitado & "," & nMontoAprobado & "," & nMoneda & ",'" & nAutEstado & "', "
        sSql = sSql & "'" & cAutObs & "','" & cUserOri & "','" & Format(dFechaOri, "yyyy-MM-dd") & "','" & cUserApro & "','" & Format(dFechaAprob, "yyyy-MM-dd") & "','" & cMovnro & "','" & cUltimaActualizacion & "'"
      
        
           Set oConecta = New COMConecta.DCOMConecta
           Set rs = New ADODB.Recordset
           oConecta.AbreConexion
           Set rs = oConecta.Ejecutar(sSql)
           If rs.State = 1 Then
                If Not (rs.EOF Or rs.BOF) Then
                    If rs.Fields(0).Value = nIdAut Then AMovAutorizacionApro = True
                End If
                 rs.Close
           End If
           
           Set rs = Nothing
           oConecta.CierraConexion
           Set oConecta = Nothing
      
      Exit Function
      
MensaErr:
      Call oError.RaiseError(oError.MyUnhandledError, "DAutorizacion:AMovAutorizacionApro Method")
End Function


Public Function AMovAutorizacionMov(ByVal nIdAut As Long, ByVal cCtaCod As String, ByVal cPersCodCli As String, ByVal cOpeCod As String, ByVal cOpeCodOri As String, _
 nMontoSolicitado As Double, nMontoAprobado As Double, nMoneda As Integer, nAutEstado As String, _
 cAutObs As String, cUserOri As String, dFechaOri As Date, cUserApro As String, dFechaAprob As Date, cMovnro As String, cUltimaActualizacion As String) As Boolean
 Dim sSql As String, rs As ADODB.Recordset
 Dim oConecta As COMConecta.DCOMConecta
 Dim oError As New COMConecta.COMErrorHandling
 
      AMovAutorizacionMov = False
      
      On Error GoTo MensaErr
        sSql = "exec Cap_ManMovAutorizacion_sp " & nIdAut & ",'" & cCtaCod & "','" & cPersCodCli & "','" & cOpeCod & "','" & cOpeCodOri & "',"
        sSql = sSql & nMontoSolicitado & "," & nMontoAprobado & "," & nMoneda & ",'" & nAutEstado & "', "
        sSql = sSql & "'" & cAutObs & "','" & cUserOri & "','" & Format(dFechaOri, "yyyy-MM-dd") & "','" & cUserApro & "','" & Format(dFechaAprob, "yyyy-MM-dd") & "','" & cMovnro & "','" & cUltimaActualizacion & "'"
      
        
         Set oConecta = New COMConecta.DCOMConecta
         Set rs = New ADODB.Recordset
           oConecta.AbreConexion
           Set rs = oConecta.Ejecutar(sSql)
           If Not (rs.EOF Or rs.BOF) Then
                If rs.Fields(0).Value = nIdAut Then AMovAutorizacionMov = True
            End If
           If rs.State = 1 Then rs.Close
           Set rs = Nothing
           oConecta.CierraConexion
           Set oConecta = Nothing
      
      Exit Function
      
MensaErr:
      Call oError.RaiseError(oError.MyUnhandledError, "DAutorizacion:AMovAutorizacionMov Method")
End Function


Public Function CargaFlexMovCapAutOpe(ByVal cUser As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
   sSql = " Select o.operacion + space(150-len(o.operacion))+ m.copecodori as Opera,m.cctacod as Cuenta,p.cpersnombre + space(150-len(p.cpersnombre))+m.cperscodcli as persona,m.nmontosolicitado,m.nmontoaprobado,(case when m.nmoneda=1 then 'S/.' else 'US$' end) + space(147)+ cast(m.nmoneda as char(1)) as moneda ,m.cautobs,m.cautestado,m.nidaut,c.nmontofinsol,c.nmontofindol "
   sSql = sSql & " from movcapautorizacionope m "
   sSql = sSql & "  inner join (Select f.copedesc +':'+ ch.copedesc as operacion, ch.copecod  from opetpo ch  inner join opetpo f on f.copecod=left(ch.copecod,2)+'0000'  where ch.copecod like  '2%[1-9]00'  ) o on o.copecod=m.copecodori "
   sSql = sSql & " inner join persona p on p.cperscod=m.cperscodcli "
   sSql = sSql & "  inner join (Select nconsvalor,cconsdescripcion from constante where nconscod='1011') as tm on tm.nconsvalor=m.nmoneda "
   sSql = sSql & "  inner join capautorizacionrango c on c.copecod=m.copecodori "
   sSql = sSql & "  inner join  (Select top 1 rhc1.crhcargocod,rhc1.cperscod from rhcargos rhc1 inner join rrhh rh1 on rh1.cperscod=rhc1.cperscod where rh1.cuser='" & cUser & "' and  rh1.nrhestado='201' "
   sSql = sSql & "  order by drhcargoFecha desc) rhc on rhc.crhcargocod=c.crhcargocod "
   sSql = sSql & "  where m.cautestado='P' and (case when nmoneda=1 then c.nMontoFinsol else c.nMontoFinDol end)>=m.nmontosolicitado "
   sSql = sSql & " order by m.nidaut "
      Set oConecta = New COMConecta.DCOMConecta
      oConecta.AbreConexion
      Set CargaFlexMovCapAutOpe = oConecta.CargaRecordSet(sSql)
      oConecta.CierraConexion
      Set oConecta = Nothing
End Function

Public Function EsHABERES(ByVal sCta As String) As Boolean
Dim sSql As String, oConec As New COMConecta.DCOMConecta, rs As New ADODB.Recordset
Set oConec = New COMConecta.DCOMConecta
 
   EsHABERES = False
 sSql = "select valor=count(*) from itfexoneracioncta "
 sSql = sSql & " where nexotpo=3 and cctacod='" & sCta & "'"
 oConec.AbreConexion
 Set rs = oConec.CargaRecordSet(sSql)
   If rs.State = 1 Then
        If rs("Valor") > 0 Then
            EsHABERES = True
        Else
            EsHABERES = False
        End If
        
   End If
 oConec.CierraConexion
 Set oConec = Nothing
 
 If rs.State = 1 Then rs.Close
 Set rs = Nothing
 
End Function
Public Function SAA(ByVal cOpeCod As String, ByVal cUser As String, ByVal cCtaCod As String, ByVal cPersCod As String, ByVal nMoneda As Integer, nIdAut As Long) As ADODB.Recordset
Dim sSql As String, oConecta As COMConecta.DCOMConecta

 sSql = " Select nIdAut,c.cCtaCod,c.cPersCodCli,nMontoAprobado from capautorizacionope c "
 sSql = sSql & " where copecodori='" & cOpeCod & "' and c.cperscodcli='" & cPersCod & "' and cctacod='" & cCtaCod & "'"
 sSql = sSql & " and c.cautestado='A' and c.cuserori='" & cUser & "' and nmoneda=" & nMoneda & " and left(CultimaActualizacion,8)=convert(char(8),getdate(),112) and nidaut=" & nIdAut
 
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set SAA = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
   
End Function

'**************** INSERTAR AUTORIZACION POR NIVEL 05/03/2006 AVMM  ***********************
Public Sub InsertarNilAutRenCan(ByVal rs As ADODB.Recordset, ByVal pnOpcion As Integer)
 ' 1=Nuevo / 2 = Modificar
 On Error GoTo rError
    Dim sql As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    
    oConecta.BeginTrans
    rs.MoveFirst
    If Not (rs.EOF And rs.BOF) Then
    
        If pnOpcion = 1 Then
            sql = " INSERT NivRetiroCanc (cNivCod,cOpeTpo,cCodAge,cNivDesc,nTopDol,nTopSol)" & _
                  " VALUES ('" & Format(rs!sNivCod, "000") & "','" & rs!sOpeTpo & "','" & rs!sCodage & "','" & rs!sNivDesc & " '," & rs!nTpoDol & "," & rs!nTpoSol & ") "
            oConecta.Ejecutar sql
        
        Else
            sql = " UPDATE NivRetiroCanc SET  cNivDesc= '" & rs!sNivDesc & " ' , nTopDol =" & rs!nTpoDol & ", nTopSol =" & rs!nTpoSol & " " & _
                  " WHERE  cNivCod ='" & Format(rs!sNivCod, "000") & "' AND cOpeTpo ='" & rs!sOpeTpo & "' AND  cCodAge='" & rs!sCodage & "' "
            oConecta.Ejecutar sql
            
        End If
    End If
    oConecta.CommitTrans
    oConecta.CierraConexion
    Exit Sub
rError:
    oConecta.CierraConexion
    oConecta.RollbackTrans
End Sub

Public Function ObtenerNroMaxNivel(ByVal psOpeTpo As String, ByVal psCodAge As String) As String
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sql = " SELECT MAX(cNivCod) cNivCod FROM NivRetiroCanc " & _
          " WHERE cCodAge ='" & psCodAge & "' and cOpeTpo ='" & psOpeTpo & "'"
    Set rs = oConecta.CargaRecordSet(sql)
    If Not (rs.EOF And rs.BOF) Then
        ObtenerNroMaxNivel = Format(IIf(IsNull(rs!cNivCod), "1", Trim(rs!cNivCod) + 1), "000")
    Else
        ObtenerNroMaxNivel = "001"
    End If
    oConecta.CierraConexion
End Function

Public Function ObtenerDatosNivAutRetCan(ByVal psOpeTpo As String, ByVal psCodAge As String) As ADODB.Recordset
    Dim sql As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    'JUEZ 20131211 *****************************************
    'sql = " SELECT  N.cNivCod, (CASE WHEN N.cOpeTpo=1 THEN 'RETIRO     '+ N.cOpeTpo  ELSE  'CANCELACION     '+ N.cOpeTpo END) cOpeTpo," & _
    '      " (cAgeDescripcion +  Space(50) + N.cCodAge)cCodAge , cNivDesc, nTopDol, nTopSol" & _
    '      " FROM NivRetiroCanc N" & _
    '      " JOIN Agencias A ON N.cCodAge=A.cAgeCod" & _
    '      " WHERE N.cCodAge ='" & psCodAge & "' and cOpeTpo ='" & psOpeTpo & "'"
    sql = "exec stp_sel_ObtenerDatosNivAutRetCan '" & psOpeTpo & "','" & psCodAge & "'"
    'END JUEZ **********************************************
    Set ObtenerDatosNivAutRetCan = oConecta.CargaRecordSet(sql)
    oConecta.CierraConexion
End Function

Public Function ObtenerDatosNivAutRetCanDet(ByVal psOpeTpo As String, ByVal psCodAge As String) As ADODB.Recordset
    Dim sql As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    '***Modificado por ELRO el 20121109, según OYP-RFC115-2012
    'sql = " SELECT  (CASE WHEN N.cOpeTpo=1 THEN 'RETIRO     '+ N.cOpeTpo  ELSE  'CANCELACION     '+ N.cOpeTpo END) cOpeTpo," & _
    '      " (cAgeDescripcion +  Space(50) + N.cCodAge)cCodAge , cGrupoUsu ,N.cNivCod + ' - ' +cNivDesc cNivel" & _
    '      " FROM NivRetiroCanc N" & _
    '      " JOIN Agencias A ON N.cCodAge=A.cAgeCod" & _
    '      " JOIN NivRetiroCancDet ND ON  N.cCodAge=ND.cCodAge and  N.cOpeTpo =ND.cOpeTpo  and   N.cNivCod=Nd.cNivCod " & _
    '      " WHERE N.cCodAge ='" & psCodAge & "' and N.cOpeTpo ='" & psOpeTpo & "'  order by N.cNivCod"
    sql = "exec stp_sel_RFC1152012_ObtenerDatosNivAutRetCanDet '" & psOpeTpo & "', '" & psCodAge & "'"
    '***Modificado por ELRO el 20121109***********************
    Set ObtenerDatosNivAutRetCanDet = oConecta.CargaRecordSet(sql)
    oConecta.CierraConexion
End Function

Public Sub VerificarMontoNivAutRetCan(ByVal psOpeTpo As String, ByVal psCodAge As String, ByVal pnMontoS As Double, _
                                      ByVal pnMontoD As Double, ByVal psNivCod As String, ByRef psMensaje As String)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sql = " SELECT cNivCod,nTopDol,nTopSol " & _
          " FROM NivRetiroCanc N WHERE cOpeTpo='" & psOpeTpo & "' and  cCodAge ='" & psCodAge & "'" & _
          " ORDER BY cNivCod DESC"
    Set rs = oConecta.CargaRecordSet(sql)
    If Not (rs.EOF And rs.BOF) Then
        Do Until rs.EOF
            If CInt(psNivCod) < CInt(rs!cNivCod) Then
                If rs!nTopSol < pnMontoS Then
                    psMensaje = "Monto Tope en soles es Mayor para este Nivel ingresar menor a  " & rs!nTopSol
                    Exit Sub
                End If
                If rs!nTopDol < pnMontoD Then
                    psMensaje = "Monto Tope en dolares es Mayor para este Nivel ingresar menor a  " & rs!nTopDol
                    Exit Sub
                End If
            ElseIf CInt(psNivCod) > CInt(rs!cNivCod) Then
                If rs!nTopSol > pnMontoS Then
                    psMensaje = "Monto Tope en soles es Menor para este Nivel ingresar Mayor a  " & rs!nTopSol
                    Exit Sub
                End If
                If rs!nTopDol > pnMontoD Then
                    psMensaje = "Monto Tope en dolares es Menor para este Nivel ingresar Mayor a  " & rs!nTopDol
                    Exit Sub
                End If
            End If
            rs.MoveNext
        Loop
    End If
    oConecta.CierraConexion
End Sub

Public Function ObtenerNivelesAut(ByVal psOpeCod As String, ByVal psCodAge As String) As ADODB.Recordset
    Dim sql As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sql = " SELECT ( cNivCod +' - '+ cNivDesc )cNivel FROM NivRetiroCanc " & _
          " WHERE  cCodAge ='" & psCodAge & "' and cOpeTpo ='" & psOpeCod & "'"
    Set ObtenerNivelesAut = oConecta.CargaRecordSet(sql)
    oConecta.CierraConexion
End Function

Public Sub VerificarNroNivxGrupoUsu(ByVal psCodGrupo As String, ByVal psCodAge As String, ByRef pnMensaje As String)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sql = " SELECT count(cNivCod) Total FROM NivRetiroCancDet where cGrupoUsu='" & psCodGrupo & "' and cCodAge='" & psCodAge & "'"
    Set rs = oConecta.CargaRecordSet(sql)
    If Not (rs.EOF And rs.BOF) Then
        If rs!Total = 2 Then
            pnMensaje = "Solo se puede Asiganar a un grupo por Nivel"
        End If
    End If
    oConecta.CierraConexion
End Sub

Public Sub InsertarNilAutRenCanDet(ByVal rs As ADODB.Recordset, ByVal pnOpcion As Integer)
 '1= nuevo/2 modificar
 On Error GoTo rError
    Dim sql As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    rs.MoveFirst
    oConecta.BeginTrans
    If Not (rs.EOF And rs.BOF) Then
        
            '***Modificado por ELRO el 20121112, según OYP-RFC115-2012
            'sql = " INSERT NivRetiroCancDet (cGrupoUsu,cNivCod,cOpeTpo,cCodAge)" & _
            '      " VALUES ('" & rs!sGrupoUsu & "','" & rs!sNivCod & "','" & rs!sOpeTpo & "','" & rs!sCodage & " ') "
            sql = "exec stp_sel_RFC1152012_RegistrarNilAutRenCanDet '" & rs!sGrupoUsu & "','" & rs!sNivCod & "','" & rs!sOpeTpo & "','" & rs!sCodage & " '"
            '***Fin Modificado por ELRO el 20121112*******************
            oConecta.Ejecutar sql
            rs.MoveNext
    End If
    oConecta.CommitTrans
    oConecta.CierraConexion
    Exit Sub
rError:
    oConecta.CierraConexion
    oConecta.RollbackTrans
End Sub

' ************ FUNCIONES PARA REALIZAR EL REGISTRO DE MovRetCancAutoriza/MovRetCancAutorizaDet**********************

Public Function ObtenerDatosMovAutorizacion(ByVal psOpeTpo As String, ByVal psCodAge As String, ByVal pdFecSis As Date, ByVal pnNivel As String) As ADODB.Recordset
    'Funcion para obtener Todos los Movimientos autorizados por Nivel
    Dim sql As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    '***Agregado por ELRO el 20121112, según OYP-RFC115-2012
    'sql = " SELECT A.dFecha, A.cCtaCod, A.cUsuAutoriza, A.nMonto, M.nMovNro, A.nNivelMax FROM Mov M JOIN MovRetCancAutoriza A " & _
    '      " ON M.nMovNro = A.nMovNro Where A.nEstado = " & gCapNivRetCancEstSolicitado & " " _
    '      & "And M.cMovNro LIKE '" & Format$(pdFecSis, "yyyymmdd") & "%' " & _
    '      " And A.cOpeTpo = '" & psOpeTpo & "' And SUBSTRING(M.cMovNro,18,2) ='" & psCodAge & "' And Convert(Int,A.nNivel) + 1 = " & pnNivel
    sql = "exec stp_sel_RFC1152012_ObtenerDatosMovAutorizacion '" & psOpeTpo & "', '" & psCodAge & "', '" & Format$(pdFecSis, "yyyymmdd") & "', '" & pnNivel & "'"
    '***Fin Agregado por ELRO el 20121112*******************
    Set ObtenerDatosMovAutorizacion = oConecta.CargaRecordSet(sql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function ObtenerMaxNivelAutRetCan(ByVal psCodAge As String, ByVal psOpeTpo As String, ByVal pnMoneda As Integer, ByVal pnMonto As Currency) As String
    'Funcion para obtener el Maximo Nivel de Autorizacion
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
        sql = " SELECT cNivCod,nTopDol,nTopSol FROM NivRetiroCanc WHERE cCodAge= '" & psCodAge & "' and cOpeTpo ='" & psOpeTpo & "'"
        Set rs = oConecta.CargaRecordSet(sql)
        If Not (rs.EOF And rs.BOF) Then
            Do Until rs.EOF
               If pnMoneda = 1 Then
                    If pnMonto <= rs!nTopSol Then
                        ObtenerMaxNivelAutRetCan = Format(rs!cNivCod, "000")
                        Exit Do
                    End If
               Else
                    If pnMonto <= rs!nTopDol Then
                        ObtenerMaxNivelAutRetCan = Format(rs!cNivCod, "000")
                        Exit Do
                    End If
               End If
               rs.MoveNext
            Loop
            
        End If
    oConecta.CierraConexion
End Function

Public Function VerificarNivAutRetCan(ByVal psGrupo As String, ByVal psOpeTpo As String, _
    ByVal psCodAge As String, Optional psPersCod As String = "") As String
    'Funcion para Verificar Nivel de Autrizacion con respecto al maximo Nivel
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim lsNivel As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
        '***Modificado por ELRO el 20121112, según OYP-RFC115-2012
        'sql = "SELECT TOP 1 N.cNivCod, N.nTopDol FROM NivRetiroCanc N JOIN NivRetiroCancDet ND ON " _
        '    & "N.cNivCod = ND.cNivCod And N.cOpeTpo = ND.cOpeTpo And N.cCodAge = ND.cCodAge " _
        '    & "WHERE ND.cGrupoUsu IN ('" & psGrupo & "') " _
        '    & "And N.cOpeTpo = '" & psOpeTpo & "' and N.cCodAge = '" & psCodAge & "' Order by N.nTopDol DESC"
        sql = "exec stp_sel_RFC1152012_VerificarNivAutRetCan '" & psGrupo & "', '" & psOpeTpo & "', '" & psCodAge & "', '" & psPersCod & "'"
        '***Fin Modificado por ELRO el 20121112*******************
        Set rs = oConecta.CargaRecordSet(sql)
        If Not (rs.EOF And rs.BOF) Then
            VerificarNivAutRetCan = Format$(rs!cNivCod, "000")
        Else
            VerificarNivAutRetCan = ""
        End If
    oConecta.CierraConexion
End Function

Public Sub InsertarMovAutRetCanDet(ByVal pnMovNro As Long, ByVal psCtaCod As String, ByVal psOpeTpo As String, ByVal pnEstado As Integer, ByVal pnMonto As Double, ByVal psUsuAutoriza As String, _
                                ByVal psNivel As String, ByVal pdFecha As Date, ByVal psNivMax As String, ByVal psOpeCod As String)
    
    Dim sql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oMov As COMDMov.DCOMMov
    Dim oGen As New COMDConstSistema.DCOMGeneral

    Set oConecta = New COMConecta.DCOMConecta
        
    oConecta.AbreConexion
        
        sql = " INSERT MovRetCancAutoriza (nMovNro,cCtaCod,cOpeTpo,nEstado,nMonto,cUsuAutoriza,nNivel,nNivelMax,dFecha)" & _
              " VALUES ('" & pnMovNro & "','" & psCtaCod & "','" & psOpeTpo & "','" & pnEstado & "'," & pnMonto & ",'" & psUsuAutoriza & "','" & psNivel & "','" & psNivMax & "','" & oGen.FechaHora(pdFecha) & "')"
        oConecta.Ejecutar sql
                
        sql = " INSERT MovRetCancAutorizaDet (nMovNro,cCtaCod,cOpeTpo,cCodUsuAut,dFecha,cNroNivel )" & _
              " VALUES ('" & pnMovNro & "','" & psCtaCod & "','" & psOpeTpo & "','" & psUsuAutoriza & "','" & oGen.FechaHora(pdFecha) & "','" & psNivel & "')"
        oConecta.Ejecutar sql
     
    oConecta.CierraConexion
End Sub

Public Sub ActualizarMovAutRetCanDet(ByVal pnNroMov As Long, ByVal psCtaCod As String, ByVal psOpeTpo As String, _
            ByVal pnEstado As CapNivRetCancEstado, ByVal pnMonto As Double, ByVal psUsuAutoriza As String, _
            ByVal psNivel As String, ByVal pdFecha As Date, ByVal psOpeCod As String, ByVal psCodAge As String)

    Dim sql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oGen As COMDConstSistema.DCOMGeneral
    Set oConecta = New COMConecta.DCOMConecta
    Set oGen = New COMDConstSistema.DCOMGeneral
    oConecta.AbreConexion
    
'    sql = "UPDATE MovRetCancAutoriza SET nEstado = " & pnEstado & ", cUsuAutoriza ='" & psUsuAutoriza & "', nNivel = '" & psNivel & "', dFecha= '" & oGen.FechaHora(pdFecha) & "' " & _
'        "WHERE cCtaCod ='" & psCtaCod & "' AND cOpeTpo ='" & psOpeTpo & "' AND nMonto=" & pnMonto & " And nMovNro = " & pnNroMov
'    oConecta.Ejecutar sql
'
'    sql = "INSERT MovRetCancAutorizaDet (nMovNro,cCtaCod,cOpeTpo,cCodUsuAut,dFecha,cNroNivel)" & _
'        " VALUES ('" & pnNroMov & "','" & psCtaCod & "','" & psOpeTpo & "','" & psUsuAutoriza & "','" & oGen.FechaHora(pdFecha) & "','" & psNivel & "')"
    
    sql = "Exec stp_upd_ActualizarMovAutRetCanDet " & pnEstado & ",'" & psNivel & "','" & oGen.FechaHora(pdFecha) & "', '" & psCtaCod & "'," & psOpeTpo & "," & pnMonto & "," & pnNroMov & ",'" & psUsuAutoriza & "'" 'APRI20170602 ERS033-2017
    
    oConecta.Ejecutar sql
        
    oConecta.CierraConexion
    Set oConecta = Nothing
    Set oGen = Nothing
End Sub

Public Function ObtenerMontoTopNivAutRetCan(ByVal psGrupo As String, ByVal psOpeTpo As String, ByVal psCodAge As String, Optional ByVal psPersCod As String = "") As ADODB.Recordset
    ' Obtener los Montos Topes por grupo y Nivel
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
        
        'RIRO20141106 ERS159 *******************
        psGrupo = Replace(psGrupo, "'", "")
        sql = "ObtenerMontoTopNivAutRetCan '" & psGrupo & "', '" & psOpeTpo & "', '" & psCodAge & "', '" & psPersCod & "'"
    
        'sql = " SELECT TOP 1 N.cNivCod, N.nTopDol, N.nTopSol FROM NivRetiroCanc N " & _
        '      " JOIN NivRetiroCancDet ND ON N.cNivCod = ND.cNivCod and N.cCodAge = ND.cCodAge AND N.cOpeTpo=ND.cOpeTpo " & _
        '      " WHERE N.cOpeTpo='" & psOpeTpo & "' and  N.cCodAge ='" & psCodAge & "'  and cGrupoUsu IN ('" & psGrupo & "') Order by N.nTopDol DESC"
        'END RIRO ******************************
        
        Set ObtenerMontoTopNivAutRetCan = oConecta.CargaRecordSet(sql)
    oConecta.CierraConexion
End Function

Public Function VerificarAutRetCan(ByVal psCtaCod As String, ByVal psOpeTpo As String, ByVal pnMonto As Double, ByVal sMovNro As String) As Integer
    ' Verifica si ya existe Datos para este Nivel
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
        sql = " SELECT A.nEstado FROM Mov M JOIN MovRetCancAutoriza A ON M.nMovNro = A.nMovNro " _
            & " WHERE A.cCtaCod = '" & psCtaCod & "' AND A.cOpeTpo ='" & psOpeTpo & "' AND " _
            & " nMonto = " & pnMonto & " AND M.cMovNro = '" & sMovNro & "'"
        Set rs = oConecta.CargaRecordSet(sql)
        If Not (rs.EOF And rs.BOF) Then
            VerificarAutRetCan = rs("nEstado")
        Else
            VerificarAutRetCan = -1
        End If
    oConecta.CierraConexion
End Function

Public Function VerificarEstadoAutRetCan(ByVal psCtaCod As String, ByVal psOpeTpo As String, ByVal pnMonto As Double, ByVal pdFecha As Date, ByVal pnEstado As Integer) As Boolean
    ' Verifica si ya existe Datos para este Nivel
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
        sql = " SELECT count(*) NroReg FROM MovRetCancAutoriza  " & _
              " WHERE cCtaCod ='" & psCtaCod & "' AND cOpeTpo ='" & psOpeTpo & "' AND nMonto=" & pnMonto & " AND" & _
              " DateDiff(Day,dFecha,'" & Format(pdFecha, "mm/dd/yyyy") & "') = 0 and nEstado='" & pnEstado & "'"
        Set rs = oConecta.CargaRecordSet(sql)
        If Not (rs.EOF And rs.BOF) Then
            If rs!NroReg <> 0 Then
                VerificarEstadoAutRetCan = True
            Else
                VerificarEstadoAutRetCan = False
            End If
        End If
    oConecta.CierraConexion
End Function

Public Function VerificarUsuarioAutRetCan(ByVal psCtaCod As String, ByVal psOpeTpo As String, ByVal pnMonto As Double, ByVal pdFecha As Date, ByVal psCodUser As String) As Boolean
    ' Verifica si ya existe Datos para este Nivel
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
        sql = " SELECT count(*) NroReg FROM MovRetCancAutoriza  " & _
              " WHERE cCtaCod ='" & psCtaCod & "' AND cOpeTpo ='" & psOpeTpo & "' AND nMonto=" & pnMonto & " AND" & _
              " DateDiff(Day,dFecha,'" & Format(pdFecha, "mm/dd/yyyy") & "') = 0 and cUsuAutoriza='" & psCodUser & "'"
        Set rs = oConecta.CargaRecordSet(sql)
        If Not (rs.EOF And rs.BOF) Then
            If rs!NroReg <> 0 Then
                VerificarUsuarioAutRetCan = True
            Else
                VerificarUsuarioAutRetCan = False
            End If
        End If
    oConecta.CierraConexion
End Function

Public Function ObtenerDatosMovAutorizacionAprobada(ByVal psOpeTpo As String, ByVal psCodAge As String, ByVal pdFecSis As Date) As ADODB.Recordset
    'Funcion para obtener Todos los Movimientos autorizados por Nivel
    Dim sql As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sql = " SELECT A.dFecha,A.cCtaCod,cCodUsuAut,nMonto,nNivelMax,nTopDol,nTopSol " & _
          " FROM MovRetCancAutoriza A " & _
          " JOIN MovRetCancAutorizaDet AD ON A.cCtaCod= AD.cCtaCod  and A.nMovNro=AD.nMovNro and A.cOpeTpo=AD.cOpeTpo" & _
          " JOIN NivRetiroCanc N ON N.cNivCod=Ad.cNroNivel AND N.cOpeTpo=A.cOpeTpo" & _
          " Where nEstado = 0 and DateDiff(Day,A.dFecha,'09/27/2005') = 0 and A.cOpeTpo='2' and  N.cCodAge ='01'"
    Set ObtenerDatosMovAutorizacionAprobada = oConecta.CargaRecordSet(sql)
    oConecta.CierraConexion
End Function


'***Agregado por ELRO el 20121110, según OYP-RFC115-2012
Public Sub eliminarNivRetiroCancDet(ByVal psGrupoUsu As String, ByVal psNivCod As String, _
                                    ByVal psOpeTpo As String, ByVal psCodAge As String)
    
    Dim sql As String
    Dim oDCOMConecta As COMConecta.DCOMConecta
    Set oDCOMConecta = New COMConecta.DCOMConecta
        
    oDCOMConecta.AbreConexion
    sql = "exec stp_sel_RFC1152012_EliminarNivRetiroCancDet '" & psGrupoUsu & "', '" & psNivCod & "', '" & psOpeTpo & "', '" & psCodAge & "'"
    oDCOMConecta.Ejecutar sql
    oDCOMConecta.CierraConexion
End Sub

'FRHU 20140505 ERS063-2014
Public Sub InsertarMovOtrasOperacAutoriza(ByVal pnMovNro As Long, ByVal psPersCod As String, ByVal psOpeTpo As String, ByVal pnEstado As Integer, ByVal pnMonto As Double, ByVal psMoneda As String, _
                                          ByVal psGlosa As String, ByVal psUsu As String, ByVal pdFecha As Date)
    
    Dim sql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oGen As New COMDConstSistema.DCOMGeneral

    Set oConecta = New COMConecta.DCOMConecta
        
    oConecta.AbreConexion
        
    sql = "exec stp_ins_MovOtrasOperacAutoriza " & pnMovNro & ",'" & psPersCod & "','" & psOpeTpo & "'," & pnEstado & "," & pnMonto & ",'" & psMoneda & "','" & psGlosa & "','" & psUsu & "','" & oGen.FechaHora(pdFecha) & "'"
    oConecta.Ejecutar sql
    oConecta.CierraConexion
End Sub
Public Function VerificarAutorOtrasOperaciones(ByVal psOpeTpo As String, ByVal pnMonto As Double, ByVal psMovNro As String) As Integer
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
        sql = "exec stp_sel_VerificaAutoOtrasOperaciones '" & psOpeTpo & "'," & pnMonto & ",'" & psMovNro & "'"
        Set rs = oConecta.CargaRecordSet(sql)
        If Not (rs.EOF And rs.BOF) Then
            VerificarAutorOtrasOperaciones = rs("nEstado")
        Else
            VerificarAutorOtrasOperaciones = -1
        End If
    oConecta.CierraConexion
End Function
Public Function ObtenerDatosOtrasOpeMovAutorizacion(ByVal psOpeTpo As String, ByVal psCodAge As String, ByVal pdFecSis As Date) As ADODB.Recordset
    Dim sql As String
    Dim oConecta As COMConecta.DCOMConecta
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sql = "exec stp_sel_DatosMovOtrasOperacionesAutoriza '" & psOpeTpo & "', '" & psCodAge & "', '" & Format$(pdFecSis, "yyyymmdd") & "'"
    Set ObtenerDatosOtrasOpeMovAutorizacion = oConecta.CargaRecordSet(sql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function
Public Sub ActualizarMovOtrasOpeAutorizacion(ByVal pnNroMov As Long, ByVal psOpeTpo As String, _
            ByVal pnEstado As CapNivRetCancEstado, ByVal pnMonto As Double, ByVal psUsuAutoriza As String, _
            ByVal pdFecha As Date, ByVal psCodAge As String)

    Dim sql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim oGen As COMDConstSistema.DCOMGeneral
    Set oConecta = New COMConecta.DCOMConecta
    Set oGen = New COMDConstSistema.DCOMGeneral
    oConecta.AbreConexion
    
    sql = "EXEC stp_upd_OtrasOperaAutorizacion '" & psOpeTpo & "'," & pnMonto & ",'" & pnNroMov & "','" & pnEstado & "','" & psUsuAutoriza & "','" & oGen.FechaHora(pdFecha) & "'"
    oConecta.Ejecutar sql
        
    oConecta.CierraConexion
    Set oConecta = Nothing
    Set oGen = Nothing
End Sub
'FIN FRHU 20140505
