VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMCaptAutorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim Aucperscod As String
Dim Aucpersnombre As String
Dim Audpersnaccreac As String
Dim Aucnivel As String
Dim Aucrhcargodescripcion As String
Dim Aucrhcargocod As String
Dim AunMontoIniSol As Double
Dim AunMontoFinSol As Double
Dim AunMontoIniDol As Double
Dim AunMontoFinDol As Double


 Public Sub DevInfoAutorizacion(ByVal cUser As String, ByVal copecod As String)
  Dim oConecta As COMConecta.DCOMConecta, rs As ADODB.Recordset
  Dim oError As New COMConecta.COMErrorHandling
  
  Dim sSql As String
     Set oConecta = New COMConecta.DCOMConecta
     Set rs = New ADODB.Recordset
     
      sSql = "Select p.cperscod,p.cpersnombre,p.dpersnaccreac,c.cnivel,rhct.crhcargodescripcion,rhc.crhcargocod,nMontoIniSol,nMontoFinSol, nMontoIniDol,nMontoFinDol  "
      sSql = sSql & "  from persona p"
      sSql = sSql & "  inner join rrhh  rh on rh.cperscod=p.cperscod"
      sSql = sSql & "  inner join (Select top 1 rhc1.crhcargocod,rhc1.cperscod from rhcargos rhc1 inner join rrhh rh1 on rh1.cperscod=rhc1.cperscod "
      sSql = sSql & "  where rh1.cuser='" & cUser & "' order by drhcargoFecha desc) rhc on rhc.cperscod=p.cperscod"
      sSql = sSql & "  inner join rhcargostabla rhct on rhct.crhcargocod=rhc.crhcargocod "
      sSql = sSql & "  inner join capautorizacionrango c on c.crhcargocod=rhc.crhcargocod "
      sSql = sSql & "  where  rh.nrhestado='201' and rh.cUser='" & cUser & "' and c.copecod='" & copecod & "'"
On Error GoTo MensaError
      Set oConecta = New COMConecta.DCOMConecta
      oConecta.AbreConexion
      Set rs = oConecta.Ejecutar(sSql)  '  oConecta.ConexionActiva.Execute(sSql)
        If rs.State = 1 Then
          If Not (rs.EOF Or rs.BOF) Then
            Aucperscod = rs!cPersCod
            Aucpersnombre = rs!cPersNombre
            Audpersnaccreac = rs!dPersNacCreac
            Aucnivel = rs!cNivel
            Aucrhcargodescripcion = rs!crhcargodescripcion
            Aucrhcargocod = rs!cRHCargoCod
            AunMontoIniSol = rs!nMontoIniSol
            AunMontoFinSol = rs!nMontoFinSol
            AunMontoIniDol = rs!nMontoIniDol
            AunMontoFinDol = rs!nMontoFinDol
           End If
           rs.Close
         End If
         Set rs = Nothing
      oConecta.CierraConexion
      Set oConecta = Nothing
    Exit Sub
MensaError:
        Call oError.RaiseError(oError.MyUnhandledError, "DAutorizacion:IAutRango Method")
     
End Sub

Public Function VerificarAutorizacion(ByVal psCtaCod As String, ByVal psOpeTpo As String, ByVal pnMonto As Double, _
                ByVal psMovNro As String, ByRef psMensaje As String) As Boolean

Dim oCapAut As COMDCaptaGenerales.COMDCaptAutorizacion
Dim lsMovNro As String
Dim lsNivMax As String
Dim nEstado As CapNivRetCancEstado
 
 'Verificar si existe una autorizacion en geneneral para esa cuenta, misma fecha, operacion y grupo
Set oCapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
    nEstado = oCapAut.VerificarAutRetCan(psCtaCod, psOpeTpo, pnMonto, psMovNro)
Set oCapAut = Nothing
 
If nEstado = gCapNivRetCancEstAprobado Then
    psMensaje = "Esta Operación para la Cuenta :" & psCtaCod & " ya esta Autorizada"
    VerificarAutorizacion = True
ElseIf nEstado = gCapNivRetCancEstRechazado Then
    psMensaje = "Esta Operación para la Cuenta :" & psCtaCod & " fue Rechazada"
    VerificarAutorizacion = False
ElseIf nEstado = gCapNivRetCancEstSolicitado Then
    psMensaje = "Esta Operación para la Cuenta :" & psCtaCod & " Aun no esta Autorizada"
    VerificarAutorizacion = False
Else
    psMensaje = "Esta Solicitud Operación para la Cuenta :" & psCtaCod & " No Existe"
    VerificarAutorizacion = False
End If

End Function

Public Function AprobarAutorizacion(ByVal psCtaCod As String, ByVal psOpeTpo As String, ByVal pnMonto As Double, ByVal pnMovNroOpe As Long, _
        ByVal psCodAge As String, ByVal psCodUser As String, ByVal pdFecSis As Date, ByVal nEstadoApr As CapNivRetCancEstado, _
        ByVal psNivelUsu As String, ByVal psNivelMax As String, ByRef psMensaje As String) As Boolean
   
Dim CapAut As COMDCaptaGenerales.COMDCaptAutorizacion
Set CapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
If nEstadoApr = gCapNivRetCancEstRechazado Then
    CapAut.ActualizarMovAutRetCanDet pnMovNroOpe, psCtaCod, psOpeTpo, gCapNivRetCancEstRechazado, pnMonto, psCodUser, psNivelUsu, pdFecSis, "", psCodAge
Else
    If psNivelUsu = psNivelMax Then
        CapAut.ActualizarMovAutRetCanDet pnMovNroOpe, psCtaCod, psOpeTpo, gCapNivRetCancEstAprobado, pnMonto, psCodUser, psNivelUsu, pdFecSis, "", psCodAge
    Else
        CapAut.ActualizarMovAutRetCanDet pnMovNroOpe, psCtaCod, psOpeTpo, gCapNivRetCancEstSolicitado, pnMonto, psCodUser, psNivelUsu, pdFecSis, "", psCodAge
    End If
End If
Set CapAut = Nothing
End Function

Public Sub NuevaSolicitudAutorizacion(ByVal psCtaCod As String, ByVal psOpeTpo As String, ByVal pnMonto As Double, _
            ByVal pdFecSis As Date, ByVal psAgeCod As String, ByVal psUserCod As String, ByVal pnMoneda As Integer, _
            ByVal psOpeCod As String, ByVal psNivel As String, ByRef psMovNro As String)
    
    Dim oCapAut As COMDCaptaGenerales.COMDCaptAutorizacion
    Dim oMov As COMDMov.DCOMMov
    
    Dim lsMovNro As String
    Dim lsNivMax As String
    
    Dim lnMovNro As Long
    Dim lbEstado As Boolean
    Dim lsGlosa As String 'JUEZ 20131218
    
    'JUEZ 20131218 ******************************
    Select Case psOpeTpo
        Case 1
            lsGlosa = "Autorización de Retiro"
        Case 2
            lsGlosa = "Autorización de Cancelación"
        Case 3
            lsGlosa = "Autorización de Cargo a Cuenta"
    End Select
    'END JUEZ ***********************************
    
    Set oMov = New COMDMov.DCOMMov
        lsMovNro = oMov.GeneraMovNro(pdFecSis, psAgeCod, psUserCod)
        'oMov.InsertaMov lsMovNro, psOpeCod, "Aurotizacion de Retiro", gMovEstContabNoContable
        oMov.InsertaMov lsMovNro, psOpeCod, lsGlosa, gMovEstContabNoContable 'JUEZ 20131218
        lnMovNro = oMov.GetnMovNro(lsMovNro)
    Set oMov = Nothing
    psMovNro = lsMovNro
        
    Set oCapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
    'Obtener el nivel maximo para esta operacion
    lsNivMax = oCapAut.ObtenerMaxNivelAutRetCan(psAgeCod, psOpeTpo, pnMoneda, pnMonto)
    oCapAut.InsertarMovAutRetCanDet lnMovNro, psCtaCod, psOpeTpo, gCapNivRetCancEstSolicitado, pnMonto, psUserCod, psNivel, pdFecSis, lsNivMax, psOpeCod
    Set oCapAut = Nothing
  
End Sub

Property Get DAucPersCod() As String
         DAucPersCod = Aucperscod
End Property

Property Get DAucpersnombre() As String
         DAucpersnombre = Aucpersnombre
End Property
Property Get DAudpersnaccreac() As String
         DAudpersnaccreac = Audpersnaccreac
End Property
Property Get DAucnivel() As String
         DAucnivel = Aucnivel
End Property
Property Get DAucrhcargodescripcion() As String
         DAucrhcargodescripcion = Aucrhcargodescripcion
End Property

Property Get DAucrhcargocod() As String
         DAucrhcargocod = Aucrhcargocod
End Property

Property Get DAunMontoIniSol() As Double
    DAunMontoIniSol = AunMontoIniSol
End Property
Property Get DAunMontoFinSol() As Double
    DAunMontoFinSol = AunMontoFinSol
End Property
Property Get DAunMontoIniDol() As Double
    DAunMontoIniDol = AunMontoIniDol
End Property
Property Get DAunMontoFinDol() As Double
    DAunMontoFinDol = AunMontoFinDol
End Property

'FRHU 20140505 ERS063-2014
Public Sub NuevaSolicitudOtrasOperaciones(ByVal psPersCod As String, ByVal psOpeTpo As String, ByVal pdFecSis As Date, _
                                          ByVal pnMonto As Double, ByVal psMoneda As String, ByVal psGlosa As String, ByVal psUserCod As String, _
                                          ByVal psOpeCod As String, ByVal psAgeCod As String, ByRef psMovNro As String, Optional ByVal pnMovNroRef As Long = 0)
    
    Dim oCapAut As COMDCaptaGenerales.COMDCaptAutorizacion
    Dim clsCapMov As COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim oMov As COMDMov.DCOMMov
    
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim lsGlosa As String
    
    If psOpeTpo <> 4 Then
        lsGlosa = "Autorización de Otros Egresos en Efectivo"
    Else
        lsGlosa = "Autorización de Registro 04 últimos Sueldos CTS" 'APRI20170601 ERS033-2017
    End If
    
    Set oMov = New COMDMov.DCOMMov
        lsMovNro = oMov.GeneraMovNro(pdFecSis, psAgeCod, psUserCod)
        oMov.InsertaMov lsMovNro, psOpeCod, lsGlosa, gMovEstContabNoContable
        lnMovNro = oMov.GetnMovNro(lsMovNro)
    Set oMov = Nothing
    psMovNro = lsMovNro
   
    Set oCapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
    oCapAut.InsertarMovOtrasOperacAutoriza lnMovNro, psPersCod, psOpeTpo, gOtraOperacEstSolicitado, pnMonto, psMoneda, psGlosa, psUserCod, pdFecSis
    Set oCapAut = Nothing
    
    Set clsCapMov = New COMDCaptaGenerales.DCOMCaptaMovimiento
        If pnMovNroRef <> 0 Then
            clsCapMov.AgregaMovRef lnMovNro, pnMovNroRef
        End If
    Set clsCapMov = Nothing
    
End Sub
Public Function VerificarAutorizacionOtrasOperaciones(ByVal psOpeTpo As String, ByVal pnMonto As Double, _
                ByVal psMovNro As String, ByRef psMensaje As String, ByRef pbRechazado As Boolean) As Boolean

Dim oCapAut As COMDCaptaGenerales.COMDCaptAutorizacion
Dim lsMovNro As String
Dim lsNivMax As String
Dim nEstado As OtrasOperacionesAutorizaEstado
 
Set oCapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
    nEstado = oCapAut.VerificarAutorOtrasOperaciones(psOpeTpo, pnMonto, psMovNro)
Set oCapAut = Nothing
 
If nEstado = gOtraOperacEstAprobado Then
    psMensaje = "Esta Operación ya esta Autorizada"
    VerificarAutorizacionOtrasOperaciones = True
ElseIf nEstado = gOtraOperacEstRechazado Then
    psMensaje = "Esta Operación fue Rechazada"
    pbRechazado = True
    VerificarAutorizacionOtrasOperaciones = False
ElseIf nEstado = gOtraOperacEstSolicitado Then
    psMensaje = "Esta Operación Aun no esta Autorizada"
    VerificarAutorizacionOtrasOperaciones = False
Else
    psMensaje = "Esta Solicitud No Existe"
    VerificarAutorizacionOtrasOperaciones = False
End If

End Function

Public Function AprobarAutorizacionOtrasOper(ByVal psOpeTpo As String, ByVal pnMonto As Double, ByVal pnMovNroOpe As Long, _
        ByVal psCodAge As String, ByVal psCodUser As String, ByVal pdFecSis As Date, ByVal nEstadoApr As CapNivRetCancEstado) As Boolean
   
Dim CapAut As COMDCaptaGenerales.COMDCaptAutorizacion
Set CapAut = New COMDCaptaGenerales.COMDCaptAutorizacion

If nEstadoApr = gOtraOperacEstRechazado Then
    CapAut.ActualizarMovOtrasOpeAutorizacion pnMovNroOpe, psOpeTpo, gOtraOperacEstRechazado, pnMonto, psCodUser, pdFecSis, psCodAge
ElseIf nEstadoApr = gOtraOperacEstAprobado Then
    CapAut.ActualizarMovOtrasOpeAutorizacion pnMovNroOpe, psOpeTpo, gOtraOperacEstAprobado, pnMonto, psCodUser, pdFecSis, psCodAge
End If
Set CapAut = Nothing
End Function
'FIN FRHU 20140505
