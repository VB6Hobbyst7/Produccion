VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMCredExtorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ObtenerDatosExtorno(ByVal psCtaCod As String, ByVal pnTipoAprobRech As Integer) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim lnPrdEstado As Integer
    On Error GoTo ErrHandler
        If pnTipoAprobRech = 1 Then 'Aprobado
            lnPrdEstado = 2002
        Else
            'rechazado
            lnPrdEstado = 2003
        End If
        
        'sSql = "SELECT PE.CPERSNOMBRE AS CTITULAR,CN.CCONSDESCRIPCION AS ESTADO,"
        'sSql = sSql & " SUBSTRING(P.CCTACOD,9,1) AS NMONEDA,' 'AS USUARIO,"
        'sSql = sSql & " CAST(DATEPART(HOUR,CE.DPRDESTADO) AS CHAR(2))+':'+ CAST(DATEPART(MINUTE,CE.DPRDESTADO) AS CHAR(2))+':'+CAST(DATEPART(SECOND,CE.DPRDESTADO) AS CHAR(2)) AS CHORA,"
        'sSql = sSql & " CE.DPRDESTADO AS DFECHA,CE.NMONTO AS NMONTO"
        'sSql = sSql & " FROM PRODUCTO P"
        'sSql = sSql & " INNER JOIN PRODUCTOPERSONA PP ON P.CCTACOD=PP.CCTACOD AND PP.nPrdPersRelac=20"
        'sSql = sSql & " INNER JOIN PERSONA PE ON PE.CPERSCOD=PP.CPERSCOD"
        'sSql = sSql & " INNER JOIN CONSTANTE CN ON P.NPRDESTADO=CN.NCONSVALOR AND CN.NCONSCOD=3001"
        'sSql = sSql & " INNER JOIN COLOCACESTADO CE ON CE.CCTACOD=P.CCTACOD AND CE.NPRDESTADO=" & lnPrdEstado
        'sSql = sSql & " WHERE P.CCTACOD='" & psCtaCod & "'"
         
        sSql = "stp_sel_ObtenerDatosExtorno '" & psCtaCod & "', " & lnPrdEstado
         
        Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
        Set ObtenerDatosExtorno = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        Set oConec = Nothing
    Exit Function
ErrHandler:
    If oConec Is Nothing Then Set oConec = Nothing
    Set ObtenerDatosExtorno = Null
End Function
'1 verdad
'0 es falso
'-1 fallo
Public Function ValidaExtorno(ByVal psCtaCod As String, ByVal pnTipoAprobRech As Integer) As Integer
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim nCantidad As Integer
    Dim lnPrdEstado As Integer
    On Error GoTo ErrHandler
        If pnTipoAprobRech = 1 Then 'Aprobado
            lnPrdEstado = 2002
        Else
            'rechazado
            lnPrdEstado = 2003
        End If
    
        sSql = sSql & " SELECT COUNT(*) AS CANTIDAD"
        sSql = sSql & " From COLOCACESTADO"
        sSql = sSql & " WHERE CCTACOD='" & psCtaCod & "' AND NPRDESTADO=" & lnPrdEstado
        
        'By Capi 15092008 para que considere el extorno de refinanciados
        '**DAOR 20070915, Denegar el extorno si el crédito presenta un estado superior a la aprobación
        'sSql = sSql & " and cCtaCod not in ( select cCtaCod from Producto where cCtaCod='" & psCtaCod & "' and nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2202,2203,2204,2205,2206)) "
        '******************************************
         'By Capi 27112008 se adiciono para que no considere estados cancelados
        'sSql = sSql & " and cCtaCod not in ( select cCtaCod from Producto where cCtaCod='" & psCtaCod & "' and nPrdEstado in (2020,2021,2022,2031,2032,2201,2202,2203,2204,2205,2206)) "
        sSql = sSql & " and cCtaCod not in ( select cCtaCod from Producto where cCtaCod='" & psCtaCod & "' and nPrdEstado in (2020,2021,2022,2031,2032,2201,2202,2203,2204,2205,2206,2050,2060)) "
      
        
        Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        Set oConec = Nothing
        If Not rs.EOF And Not rs.BOF Then
            nCantidad = rs!cantidad
            If nCantidad > 0 Then
                ValidaExtorno = 1
            Else
                ValidaExtorno = 0
            End If
        Else
            ValidaExtorno = -1
        End If
        Set rs = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    If Not rs Is Nothing Then Set rs = Nothing
    ValidaExtorno = -1
End Function

Public Function Extorno(ByVal psCtaCod As String, _
                        Optional ByVal pdFecSis As Date, _
                        Optional ByVal psCodAge As String, _
                        Optional ByVal psCodUser As String = "SIST") As Boolean
                        
    Dim oConec As COMConecta.DCOMConecta
    Dim R As ADODB.Recordset
    Set R = New ADODB.Recordset
    Dim sSql As String
    On Error GoTo ErrHandler
    Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
        Dim lnNroCalen As Integer
        Dim lnNroCalPar As Integer
        
        sSql = "select * from ColocacCred where cCtaCod='" & psCtaCod & "'"
        Set R = oConec.CargaRecordSet(sSql)
        If R.RecordCount > 0 Then
            lnNroCalen = R!nNroCalen
            lnNroCalPar = R!nNroCalPar
        End If
        R.Close

        oConec.ConexionActiva.BeginTrans
        
        'LUCV20180601, Según ERS022-2018
        Dim oDCOMCredActBD As COMDCredito.DCOMCredActBD
        Dim lsMovNro As String
        Set oDCOMCredActBD = New COMDCredito.DCOMCredActBD
        
        lsMovNro = oDCOMCredActBD.GeneraMovNro(pdFecSis, psCodAge, psCodUser)
        'Fin LUCV20180601
            
        'Volviendo el credito a estado de sugerido
        sSql = "Update Producto set nPrdEstado=2001 Where cCtaCod='" & psCtaCod & "'"
        oConec.ConexionActiva.Execute sSql
        
        'Eliminado el estado de aprobado
        sSql = "Delete From ColocacEstado Where cCtaCod='" & psCtaCod & "' and nPrdEstado=2002"
        oConec.ConexionActiva.Execute sSql
        
        'Eliminado el ProductoPersona
        sSql = "Delete from ProductoPersona Where nPrdPersRelac=29 and cctacod='" & psCtaCod & "'"
        oConec.ConexionActiva.Execute sSql
        
        'Verificando si el tiene un credito mi vivienda
           If ObteMiVivienda(psCtaCod) = True Then
                sSql = "Delete From ColocCalifMiViv Where cCtaCod='" & psCtaCod & "'"
                oConec.ConexionActiva.Execute sSql
           End If
           
        'Para Refinanciados
        
        'sSql = "Delete from ColocacRefinancDet Where nEstado=3 and cctacod='" & psCtaCod & "'"
        'oConec.ConexionActiva.Execute sSql

        'sSql = "Delete from ColocacRefinanc Where nEstado=3 and cctacod='" & psCtaCod & "'"
        'oConec.ConexionActiva.Execute sSql

        'ARCV 27-06-2007 :
        '1)Extorna el Movimiento
        sSql = "UPDATE Mov SET nMovFlag=2 " & _
               " WHERE nMovNro= " & _
               " (SELECT TOP 1 M.nMovNro FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " & _
               " WHERE cCtaCod='" & psCtaCod & "'AND M.cOpeCod='101200' AND nMovFlag=0 ORDER BY M.nMovNro DESC)"
         
         oConec.ConexionActiva.Execute sSql
         
         '2) Estado de Producto
         sSql = " UPDATE Producto SET nPrdEstado=CS.nPrdEstado " & _
                " FROM Producto P INNER JOIN ColocacSaldo CS ON P.cCtaCod = CS.cCtaCod " & _
                " INNER JOIN ColocacRefinanc CR ON CR.cCtaCodRef = P.cCtaCod " & _
                " WHERE CR.nEstado=2 AND DATEDIFF(d,dFecha,DATEADD(d,-1,CR.dEstado))=0" & _
                " AND CR.cCtaCod='" & psCtaCod & "'"
                'LUCV20160617 Modifico: CR.nEstado=3 a CR.nEstado=2: , Segun ERS004-2016
         oConec.ConexionActiva.Execute sSql
         
        '3)Limpiar el Estado Aprobado
        sSql = "Delete from ColocacRefinancDet Where nEstado=2 and cctacod='" & psCtaCod & "'" 'LUCV20160617 Modifico: CR.nEstado=3 a CR.nEstado=2: , Segun ERS004-2016
        oConec.ConexionActiva.Execute sSql

        sSql = "Delete from ColocacRefinanc Where nEstado=2 and cctacod='" & psCtaCod & "'" 'LUCV20160617 Modifico: CR.nEstado=3 a CR.nEstado=2: , Segun ERS004-2016
        oConec.ConexionActiva.Execute sSql
         
         '------
        '**DAOR 20071204, Restablecer calendario**********************************
        If Mid(psCtaCod, 6, 3) <> "801" Then
            sSql = "delete ColocCalendDet where cCtaCod='" & psCtaCod & "' and nNroCalen=(select nNroCalen from ColocacCred where cCtaCod='" & psCtaCod & "')"
            oConec.ConexionActiva.Execute sSql
            sSql = "delete ColocCalendario where cCtaCod='" & psCtaCod & "' and nNroCalen=(select nNroCalen from ColocacCred where cCtaCod='" & psCtaCod & "')"
            oConec.ConexionActiva.Execute sSql
            sSql = "update ColocacCred set nNroCalen=nNroCalen-1 where cCtaCod='" & psCtaCod & "' "
            oConec.ConexionActiva.Execute sSql
        Else
            sSql = "delete ColocCalendDet where cCtaCod='" & psCtaCod & "' and nNroCalen= " & lnNroCalen
            oConec.ConexionActiva.Execute sSql
            sSql = "delete ColocCalendario where cCtaCod='" & psCtaCod & "' and nNroCalen=" & lnNroCalen
            oConec.ConexionActiva.Execute sSql
            
            sSql = "delete ColocCalendDet where cCtaCod='" & psCtaCod & "' and nNroCalen= " & lnNroCalPar
            oConec.ConexionActiva.Execute sSql
            sSql = "delete ColocCalendario where cCtaCod='" & psCtaCod & "' and nNroCalen=" & lnNroCalPar
            oConec.ConexionActiva.Execute sSql

            
            sSql = "update ColocacCred set nNroCalen=" & lnNroCalen - 2 & ",nNroCalPar=" & lnNroCalPar - 2 & " where cCtaCod='" & psCtaCod & "' "
            oConec.ConexionActiva.Execute sSql
        End If
        '*************************************************************************
        
        '** JUEZ 20130312 Quitar las aprobaciones por niveles ********************
        'sSql = "Update CredNivAprResultado Set nEstado=0 Where cCtaCod='" & psCtaCod & "' "
        sSql = "EXEC stp_upd_ExtornarNivAprobacion '" & psCtaCod & "'" 'EJVG20160607
        oConec.ConexionActiva.Execute sSql
        '** END JUEZ *************************************************************
         
         'WIOR 20140128 ********************************
         sSql = "exec stp_upd_CredAsignacionSaldo '" & psCtaCod & "','',0,3"
         oConec.ConexionActiva.Execute sSql
         'WIOR FIN *************************************
         
         'WIOR 20140825 ********************************
         'sSql = "exec stp_del_ExtornarColocCalendDetSegDes '" & psCtaCod & "'"
         sSql = "exec stp_del_ExtornarColocCalendDetSegDes '" & psCtaCod & "'," & lnNroCalen 'APRI2018042018 ERS028-2017 ADD lnNroCalen
         oConec.ConexionActiva.Execute sSql
         'WIOR FIN *************************************
         
         'EJVG20160616 ***
         sSql = "EXEC stp_upd_AjusteRegistroCoberturaExtorno '" & psCtaCod & "'" 'EJVG20160616->Extorna Reajuste Registro de Cobertura en caso haya tenido
         oConec.ConexionActiva.Execute sSql
         'END EJVG *******
         
        'ARLO20180926 ***
         sSql = "EXEC stp_del_EliminaControlDesemHistorial '" & psCtaCod & "'" 'ARLO20180926->Extorna el Registro del Control de Pre - Desembolso de Administracion de Créditos
         oConec.ConexionActiva.Execute sSql
         'END ARLO *******
         
         'LUCV20180601, Según ERS022-2018
         sSql = "EXEC stp_upd_ERS0632014_CancelarSolicitudPoliza '" & psCtaCod & "','" & lsMovNro & "','Cancelación de Solicitud de Poliza por Extorno de aprobación', True"
         oConec.ConexionActiva.Execute sSql
         'Fin LUCV20180601
         
         oConec.ConexionActiva.CommitTrans
         oConec.CierraConexion
         Set oConec = Nothing
         Extorno = True
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then
        oConec.ConexionActiva.RollbackTrans
    End If
    If Not oConec Is Nothing Then Set oConec = Nothing
    Extorno = False
End Function

Function ObteMiVivienda(ByVal psCtaCod As String) As Boolean
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim nCantidad As Integer
    On Error GoTo ErrHandler
        sSql = "Select count(*) as Cantidad From ColocacCred Where bMiVivienda=1 and cCtaCod='" & psCtaCod & "'"
        Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        Set oConec = Nothing
        If Not rs.BOF And Not rs.EOF Then
            nCantidad = rs!cantidad
        Else
            nCantidad = 0
        End If
        Set rs = Nothing
        
        If nCantidad > 0 Then
            ObteMiVivienda = True
        Else
            ObteMiVivienda = False
        End If
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    ObteMiVivienda = False
End Function
'EJVG20160607 ***
Public Function RecuperaDatosXExtornoNivApr(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    On Error GoTo ErrRecuperaDatosXExtornoNivApr
    sSql = "EXEC stp_sel_RecuperaDatosXExtornoNivAprobacion '" & psCtaCod & "'"
     
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set RecuperaDatosXExtornoNivApr = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrRecuperaDatosXExtornoNivApr:
    Err.Raise Err.Number, "Error En Proceso <RecuperaDatosXExtornoNivApr>", Err.Description
End Function
Public Sub ExtornarNivelAprobacion(ByVal psCtaCod As String)
    Dim oConec As COMConecta.DCOMConecta
    Dim sSql As String
    
    On Error GoTo ErrExtornarNivelAprobacion
    sSql = "EXEC stp_upd_ExtornarNivAprobacion '" & psCtaCod & "'"
     
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.Ejecutar (sSql)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Sub
ErrExtornarNivelAprobacion:
    Err.Raise Err.Number, "Error En Proceso <ExtornarNivelAprobacion>", Err.Description
End Sub
'END EJVG *******
