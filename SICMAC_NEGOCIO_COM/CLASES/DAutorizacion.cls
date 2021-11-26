VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DAutorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function InfoAutorizacion(ByVal cUser As String, ByVal copecod As String) As ADODB.Recordset
 Dim sSql As String
 Dim oConecta As DConecta
 

 
 
 
      sSql = "Select p.cperscod,p.cpersnombre,p.dpersnaccreac,c.cnivel,rhct.crhcargodescripcion,rhc.crhcargocod,nMontoIniSol,nMontoFinSol, nMontoIniDol,nMontoFinDol  "
      sSql = sSql & "  from persona p"
      sSql = sSql & "  inner join rrhh  rh on rh.cperscod=p.cperscod"
      sSql = sSql & "  inner join (Select top 1 rhc1.crhcargocod,rhc1.cperscod from rhcargos rhc1 inner join rrhh rh1 on rh1.cperscod=rhc1.cperscod "
      sSql = sSql & "  where rh1.cuser='" & cUser & "' order by drhcargoFecha desc) rhc on rhc.cperscod=p.cperscod"
      sSql = sSql & "  inner join rhcargostabla rhct on rhct.crhcargocod=rhc.crhcargocod "
      sSql = sSql & "  inner join capautorizacionrango c on c.crhcargocod=rhc.crhcargocod "
      sSql = sSql & "  where  rh.nrhestado='201' and rh.cUser='" & cUser & "' and c.copecod='" & copecod & "'"
        
      Set oConecta = New DConecta
      oConecta.AbreConexion
      Set InfoAutorizacion = oConecta.CargaRecordSet(sSql)
      oConecta.CierraConexion
      Set oConecta = Nothing
      
End Function
Public Function ExisteAutrango(ByVal copecod As String, ByVal cRHCargoCod As String) As Boolean
Dim sSql As String, rs As ADODB.Recordset
Dim oConecta As DConecta
    ExisteAutrango = False
   On Error GoTo MensaError
            
        sSql = "Select copecod+crhcargocod from capautorizacionrango where crhcargocod='" & cRHCargoCod & "' and copecod='" & copecod & "'"
        
        Set oConecta = New DConecta
        oConecta.AbreConexion
        Set rs = oConecta.Ejecutar(sSql)
        If Not (rs.EOF Or rs.BOF) Then
               If Trim(rs.Fields(0).value) = Trim(Trim(copecod) & Trim(cRHCargoCod)) Then
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
        Call RaiseError(MyUnhandledError, "DAutorizacion:ExisteAutRango Method")
End Function
Public Function CargaFlexAutRango() As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
      sSql = "Select c.copecod,Upper(o.operacion),rct.crhcargodescripcion + space(150-len(rct.crhcargodescripcion))+rct.crhcargocod Cargo,c.cnivel,c.nMontoIniSol,c.nMontoFinSol,c.nMontoIniDol,nMontoFinDol,'E' as Estado,cast(case when c.bactivo=1 then 1 else 0 end as char(1)) as bactivo "
      sSql = sSql & " from capautorizacionrango c "
      sSql = sSql & " inner join  rhcargostabla  rct on rct.crhcargocod=c.crhcargocod "
      sSql = sSql & " inner join  (Select ch.copecod,f.copedesc +':'+ ch.copedesc as operacion from opetpo ch "
      sSql = sSql & " inner join opetpo f on f.copecod=left(ch.copecod,2)+'0000' "
      sSql = sSql & " where ch.copecod like  '2%[1-9]00' ) o on o.copecod=c.copecod "
      sSql = sSql & "   Where Len(rct.cRHCargoCod) > 3 "
      sSql = sSql & " order by c.copecod "
      Set oConecta = New DConecta
      oConecta.AbreConexion
      Set CargaFlexAutRango = oConecta.CargaRecordSet(sSql)
      oConecta.CierraConexion
      Set oConecta = Nothing
End Function

Public Function Iautrango(ByVal copecod As String, ByVal cNivel As String, ByVal cRHCargoCod As String, ByVal nMontoIniSol As Double, ByVal nMontoFinSol As Double, ByVal nMontoIniDol As Double, ByVal nMontoFinDol As Double, ByVal cUltimaActualizacion As String) As Boolean
Dim sSql As String, rs As ADODB.Recordset
Dim oConecta As DConecta
 Iautrango = False

On Error GoTo MensaErr
        sSql = "exec Cap_ManCapAutorizacionRango 'N','" & copecod & "','" & cNivel & "','" & cRHCargoCod & "'," & nMontoIniSol & "," & nMontoFinSol & "," & nMontoIniDol & "," & nMontoFinDol & ",'" & cUltimaActualizacion & "'"
       
        Set oConecta = New DConecta
           oConecta.AbreConexion
           Set rs = oConecta.Ejecutar(sSql)
           If Not (rs.EOF Or rs.BOF) Then
               If rs.Fields(0).value = 1 Then
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

   Call RaiseError(MyUnhandledError, "DAutorizacion:IAutRango Method")

End Function

Public Function AAutRango(ByVal copecod As String, ByVal cNivel As String, ByVal cRHCargoCod As String, ByVal nMontoIniSol As Double, ByVal nMontoFinSol As Double, ByVal nMontoIniDol As Double, ByVal nMontoFinDol As Double, ByVal cUltimaActualizacion As String) As Boolean
    Dim sSql As String, rs As ADODB.Recordset
    Dim oConecta As DConecta
    AAutRango = False
        
On Error GoTo MensaErr

        sSql = "exec Cap_ManCapAutorizacionRango 'A','" & copecod & "','" & cNivel & "','" & cRHCargoCod & "'," & nMontoIniSol & "," & nMontoFinSol & "," & nMontoIniDol & "," & nMontoFinDol & ",'" & cUltimaActualizacion & "'"
        Set oConecta = New DConecta
           oConecta.AbreConexion
           Set rs = oConecta.Ejecutar(sSql)
           If Not (rs.EOF Or rs.BOF) Then
               If rs.Fields(0).value = 1 Then
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

   Call RaiseError(MyUnhandledError, "DAutorizacion:IAutRango Method")
End Function


Public Function IMovAutorizacion(ByVal nIdAut As Long, ByVal cCtaCod As String, ByVal cPersCodCli As String, ByVal copecod As String, ByVal cOpeCodOri As String, _
 nMontoSolicitado As Double, nMontoAprobado As Double, nMoneda As Integer, nAutEstado As String, _
cAutObs As String, cUserOri As String, dFechaOri As Date, cUserApro As String, dFechaAprob As Date, cMovNro As String, cUltimaActualizacion As String) As Long
Dim sSql As String, rs As ADODB.Recordset
Dim oConecta As DConecta
  
      IMovAutorizacion = 0
      
      On Error GoTo MensaErr
      
        sSql = "exec Cap_ManMovAutorizacion_sp 0,'" & cCtaCod & "','" & cPersCodCli & "','" & copecod & "','" & cOpeCodOri & "',"
        sSql = sSql & nMontoSolicitado & "," & nMontoAprobado & "," & nMoneda & ",'" & nAutEstado & "', "
        sSql = sSql & "'" & cAutObs & "','" & cUserOri & "','" & Format(dFechaOri, "yyyy-MM-dd") & "','" & cUserApro & "','" & Format(dFechaAprob, "yyyy-MM-dd") & "','" & cMovNro & "','" & cUltimaActualizacion & "'"
        
           Set oConecta = New DConecta
           Set rs = New ADODB.Recordset
           oConecta.AbreConexion
           Set rs = oConecta.ConexionActiva.Execute(sSql)
           If rs.State = 1 Then
                If Not (rs.EOF Or rs.BOF) Then
                     IMovAutorizacion = rs.Fields(0).value
                 End If
                 rs.Close
           End If
           Set rs = Nothing
           oConecta.CierraConexion
           Set oConecta = Nothing

      
      Exit Function
      
MensaErr:
      Call RaiseError(MyUnhandledError, "DAutorizacion:IMovAutorizacion Method")
End Function
Public Function AMovAutorizacionApro(ByVal nIdAut As Long, ByVal cCtaCod As String, ByVal cPersCodCli As String, ByVal copecod As String, ByVal cOpeCodOri As String, _
 nMontoSolicitado As Double, nMontoAprobado As Double, nMoneda As Integer, nAutEstado As String, _
cAutObs As String, cUserOri As String, dFechaOri As Date, cUserApro As String, dFechaAprob As Date, cMovNro As String, cUltimaActualizacion As String) As Boolean
Dim sSql As String, rs As ADODB.Recordset
  Dim oConecta As DConecta
      AMovAutorizacionApro = False
      
      On Error GoTo MensaErr
        sSql = "exec Cap_ManMovAutorizacion_sp " & nIdAut & ",'" & cCtaCod & "','" & cPersCodCli & "','" & copecod & "','" & cOpeCodOri & "',"
        sSql = sSql & nMontoSolicitado & "," & nMontoAprobado & "," & nMoneda & ",'" & nAutEstado & "', "
        sSql = sSql & "'" & cAutObs & "','" & cUserOri & "','" & Format(dFechaOri, "yyyy-MM-dd") & "','" & cUserApro & "','" & Format(dFechaAprob, "yyyy-MM-dd") & "','" & cMovNro & "','" & cUltimaActualizacion & "'"
      
        
           Set oConecta = New DConecta
           Set rs = New ADODB.Recordset
           oConecta.AbreConexion
           Set rs = oConecta.Ejecutar(sSql)
           If rs.State = 1 Then
                If Not (rs.EOF Or rs.BOF) Then
                    If rs.Fields(0).value = nIdAut Then AMovAutorizacionApro = True
                End If
                 rs.Close
           End If
           
           Set rs = Nothing
           oConecta.CierraConexion
           Set oConecta = Nothing
      
      Exit Function
      
MensaErr:
      Call RaiseError(MyUnhandledError, "DAutorizacion:AMovAutorizacionApro Method")
End Function


Public Function AMovAutorizacionMov(ByVal nIdAut As Long, ByVal cCtaCod As String, ByVal cPersCodCli As String, ByVal copecod As String, ByVal cOpeCodOri As String, _
 nMontoSolicitado As Double, nMontoAprobado As Double, nMoneda As Integer, nAutEstado As String, _
 cAutObs As String, cUserOri As String, dFechaOri As Date, cUserApro As String, dFechaAprob As Date, cMovNro As String, cUltimaActualizacion As String) As Boolean
 Dim sSql As String, rs As ADODB.Recordset
 Dim oConecta As DConecta
      AMovAutorizacionMov = False
      
      On Error GoTo MensaErr
        sSql = "exec Cap_ManMovAutorizacion_sp " & nIdAut & ",'" & cCtaCod & "','" & cPersCodCli & "','" & copecod & "','" & cOpeCodOri & "',"
        sSql = sSql & nMontoSolicitado & "," & nMontoAprobado & "," & nMoneda & ",'" & nAutEstado & "', "
        sSql = sSql & "'" & cAutObs & "','" & cUserOri & "','" & Format(dFechaOri, "yyyy-MM-dd") & "','" & cUserApro & "','" & Format(dFechaAprob, "yyyy-MM-dd") & "','" & cMovNro & "','" & cUltimaActualizacion & "'"
      
        
         Set oConecta = New DConecta
         Set rs = New ADODB.Recordset
           oConecta.AbreConexion
           Set rs = oConecta.Ejecutar(sSql)
           If Not (rs.EOF Or rs.BOF) Then
                If rs.Fields(0).value = nIdAut Then AMovAutorizacionMov = True
            End If
           If rs.State = 1 Then rs.Close
           Set rs = Nothing
           oConecta.CierraConexion
           Set oConecta = Nothing
      
      Exit Function
      
MensaErr:
      Call RaiseError(MyUnhandledError, "DAutorizacion:AMovAutorizacionMov Method")
End Function


Public Function CargaFlexMovCapAutOpe(ByVal cUser As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As DConecta
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
      Set oConecta = New DConecta
      oConecta.AbreConexion
      Set CargaFlexMovCapAutOpe = oConecta.CargaRecordSet(sSql)
      oConecta.CierraConexion
      Set oConecta = Nothing
End Function
