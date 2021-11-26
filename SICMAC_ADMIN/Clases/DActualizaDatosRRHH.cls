VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DActualizaDatosRRHH"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A947A9C0218"
'Actualiza los datos del RRHH de la empresa
Option Base 0
Option Explicit

Private lsServerComunes As String
Private lsServerPersona As String
Private lsServerAdministracion As String
Private lsServerNegocio As String
Private lsServerImagenes As String

'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902F000E0
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902F00298
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

Public Sub Inicio(psServerComunes As String, psServerPersona As String, psServerAdministracion As String, psServerNegocio As String, psServerImagenes As String)
    lsServerComunes = psServerComunes
    lsServerPersona = psServerPersona
    lsServerAdministracion = psServerAdministracion
    lsServerNegocio = psServerNegocio
    lsServerImagenes = psServerImagenes
End Sub

'##ModelId=3AB902F00252
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

'Agrega un nuevo cargo o nivel
'##ModelId=3A947AB400AA
Public Function AgregaRRHH(psPersCod As String, psRHCod As String, psUser As String, psClave As String, psAreaCod As String, psAgenciaAsig As String, psAgenciaActual As String, psIngreso As String, psCese As String, psUltimaActualizacion As String, psRHEstado As String) As Boolean
    On Error GoTo AgregaRRHHErr
    
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
        
    If psCese = "" Then
        sqlC = " Insert RRHH (cPersCod,cRHCod,cUser,cClave,cAreaCod,cAgenciaAsig,cAgenciaActual,dIngreso,dCese,cUltimaActualizacion,nRHEstado)" _
             & " Values('" & psPersCod & "','" & psRHCod & "','" & psUser & "','" & psClave & "','" & psAreaCod & "','" & psAgenciaAsig & "','" & psAgenciaActual & "','" & psIngreso & "',null,'" & psUltimaActualizacion & "','" & psRHEstado & "')"
    Else
        sqlC = " Insert RRHH (cPersCod,cRHCod,cUser,cClave,cAreaCod,cAgenciaAsig,cAgenciaActual,dIngreso,dCese,cUltimaActualizacion,nRHEstado)" _
             & " Values('" & psPersCod & "','" & psRHCod & "','" & psUser & "','" & psClave & "','" & psAreaCod & "','" & psAgenciaAsig & "','" & psAgenciaActual & "','" & psIngreso & "','" & psCese & "','" & psUltimaActualizacion & "','" & psRHEstado & "')"
    End If

    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
            
    Set oCon = Nothing
    Exit Function
AgregaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:AgregaRRHH Method")
End Function

'Agrega un nuevo cargo o nivel
'##ModelId=3A947AB400AA
Public Function AgregaRHEstado(psPersCod As String, psFecha As String, psEstado As String, psComentario As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaRRHHErr
    
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
        
    sqlC = " Insert RHEstado (cPersCod,dRHEstadoFecha,nRHEstado,cRHEstadoComentario,cUltimaActualizacion)" _
         & " Values('" & psPersCod & "','" & psFecha & "','" & psEstado & "','" & psComentario & "','" & psUltimaActualizacion & "')"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
            
    Set oCon = Nothing
    Exit Function
AgregaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:AgregaRRHH Method")
End Function

'Agrega un nuevo cargo o nivel
'##ModelId=3A947AB400AA
Public Function GetListaRRHH(psTipo As String, pnEstado As String) As ADODB.Recordset
    On Error GoTo AgregaRRHHErr
    Dim sqlC As String
    Dim oCon As DConecta
    Dim lsCadAdd As String
    Set oCon = New DConecta
        
    If pnEstado = -1 Then
        lsCadAdd = ""
    Else
        lsCadAdd = " And (dbo.RHEstado.nRHEstado in ('" & pnEstado & "'))"
    End If
    
    sqlC = " SELECT dbo.RRHH.cRHCod Cod, dbo.Persona.cPersNombre Nombre," _
         & " dbo.RRHH.cUser Usuario, dbo.Constante.cConsDescripcion Estado" _
         & " From dbo.RRHH" _
         & " INNER JOIN  dbo.Persona ON dbo.RRHH.cPersCod = dbo.Persona.cPersCod" _
         & " INNER JOIN  dbo.RHEstado ON dbo.RRHH.cPersCod = dbo.RHEstado.cPersCod" _
         & " INNER JOIN  dbo.Constante ON dbo.RHEstado.nRHEstado = dbo.Constante.nConsValor" _
         & " WHERE (dbo.Constante.nConsCod = '" & gRHEstado & "') And (dbo.RRHH.cRHCod like '" & psTipo & "%')" & lsCadAdd _
         & " GROUP BY dbo.RRHH.cRHCod, dbo.Persona.cPersNombre, dbo.RRHH.cUser," _
         & " dbo.Constante.cConsDescripcion"
    If oCon.AbreConexion Then
        Set GetListaRRHH = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
            
    Set oCon = Nothing
    Exit Function
AgregaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:AgregaRRHH Method")
End Function

'Agrega un nuevo cargo o nivel
'##ModelId=3A947AB400AA
Public Function RescindeRRHH(psPersCod As String, psContratoCod As String, psFecha As String, psMotivoTpo As String, psComentario As String, psUltimaActualizacion As String, psMotivo2Tpo As String) As Boolean
    On Error GoTo AgregaRRHHErr
    Dim sqlC As String
    Dim sqlRH As String
    Dim sqlEstado As String
    Dim sqlComentario As String
    Dim sSqlCtrlAsist As String
    Dim sqlAsist As String
    Dim oCon As DConecta
    Set oCon = New DConecta
        
    '*** PEAC 20100625
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    '***FIN PEAC
        
    sqlC = " Update RHContrato" _
         & " Set dFechaFin = '" & psFecha & "', nRHRescisionTpo = " & psMotivoTpo & ", cRHRescisionComent = '" & psComentario & "', cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
         & " Where cPersCod = '" & psPersCod & "' And cRHContratoNro = '" & psContratoCod & "'"
    
    sqlRH = " Update RRHH Set nRHEstado = " & psMotivoTpo & ", dCese = '" & psFecha & "',  cUltimaActualizacion = '" & psUltimaActualizacion & "' Where cPersCod = '" & psPersCod & "'"
    
    sqlEstado = " Insert RHEstado (cPersCod,dRHEstadoFecha,nRHEstado,cRHEstadoComentario,cUltimaActualizacion) Values ('" & psPersCod & "','" & psFecha & "'," & psMotivoTpo & ",'" & psComentario & "','" & psUltimaActualizacion & "')"
    
    sqlComentario = " Insert RHcomentario(cPersCod, dRHComentarioFecha, cRHComentario, cUltimaActualizacion)" _
                  & " Values('" & psPersCod & "','" & psFecha & "','" & psComentario & "','" & psUltimaActualizacion & "')"
    
'    '*** PEAC 20100625
'    sSqlCtrlAsist = " exec stp_sel_BuscaEmpleadoCtrlAsistencia '" & psPersCod & "'"
'    '*** FIN PEAC
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.Ejecutar sqlRH
        oCon.Ejecutar sqlEstado
        oCon.Ejecutar sqlComentario
'        '*** PEAC 20100625
'            Set rs = oCon.Ejecutar(sSqlCtrlAsist)
'            If Not (rs.EOF And rs.BOF) Then
'                sqlAsist = " exec stp_upd_EmpleadoCeseCtrlAsist '" & psPersCod & "'," & 0 & ",'" & Format(psFecha, "yyyymmdd") & "'"
'                oCon.Ejecutar sqlAsist
'            End If
'            rs.Close
'        '*** FIN PEAC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Set rs = Nothing '*** PEAC 20100625
    Exit Function
AgregaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:AgregaRRHH Method")
End Function


'Agrega un nuevo cargo o nivel
'##ModelId=3A947AB400AA
Public Function AgregaRHCargo(psPersCod As String, psFecha As String, psCargoAreaAsig As String, psAgenciaAsig As String, psCargoAreaAct As String, psAgenciaAct As String, psComentario As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaRRHHErr
    
    Dim sqlC As String
    Dim sqlR As String
    Dim oCon As DConecta
    Dim sSqlCtrlAsist As String '*** PEAC 20100625
    Dim sqlAsist As String
    Set oCon = New DConecta

    '*** PEAC 20100625
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    '***FIN PEAC
     
    'Comentado by NAGL 20171204
    'sqlC = " Insert RHCargos (cPersCod,dRHCargoFecha,cRHCargoCodOficial,cRHAreaCodOficial,cRHAgenciaCodOficial,cRHCargoCod,cRHAreaCod,cRHAgenciaCod,cRHCargoComentario,cUltimaActualizacion)" _
         '& " Values('" & psPersCod & "','" & psFecha & "','" & Right(psCargoAreaAsig, 6) & "','" & Left(psCargoAreaAsig, 3) & "','" & psAgenciaAsig & "','" & Right(psCargoAreaAct, 6) & "','" & Left(psCargoAreaAct, 3) & "','" & psAgenciaAct & "','" & psComentario & "','" & psUltimaActualizacion & "')"
    
    'sqlR = " Update rrhh" _
         '& " Set cAreaCod = '" & Left(psCargoAreaAsig, 3) & "', cAreaCodActual = '" & Left(psCargoAreaAct, 3) & "', cAgenciaAsig = '" & psAgenciaAsig & "' , cAgenciaActual = '" & psAgenciaAct & "'" _
         '& " Where cPersCod = '" & psPersCod & "'"
    'END NAGL 20171204
    
    sqlR = "Exec stp_ins_RHCargos '" & psPersCod & "','" & psFecha & "','" & Right(psCargoAreaAsig, 6) & "','" & Left(psCargoAreaAsig, 3) & "','" & psAgenciaAsig & "','" & Right(psCargoAreaAct, 6) & "','" & Left(psCargoAreaAct, 3) & "','" & psAgenciaAct & "','" & psComentario & "','" & psUltimaActualizacion & "'"
    'Agregado by NAGL ERS074-2017 20171207
    
'    '*** PEAC 20100625
'    sSqlCtrlAsist = " exec stp_sel_BuscaEmpleadoCtrlAsistencia '" & psPersCod & "'"
'    '*** FIN PEACG
    
    If oCon.AbreConexion Then
        'oCon.Ejecutar sqlC 'Comentado by NAGL 20171207
        oCon.Ejecutar sqlR
'        '*** PEAC 20100625
'            Set rs = oCon.Ejecutar(sSqlCtrlAsist)
'            If Not (rs.EOF And rs.BOF) Then
'                sqlAsist = " exec stp_upd_EmpleadoCargoAreaAgeCtrlAsist '" & psPersCod & "'," & Val(Left(psCargoAreaAsig, 3)) & "," & Val(Right(psCargoAreaAsig, 6)) & "," & Val(psAgenciaAsig)
'                oCon.Ejecutar sqlAsist
'            End If
'            rs.Close
'        '*** FIN PEAC
        oCon.CierraConexion
    End If
            
    Set oCon = Nothing
    Set rs = Nothing
    Exit Function
AgregaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:AgregaRRHH Method")
End Function

Public Function ObtenerUltimoCargoOficial(ByVal psPersCod As String) As ADODB.Recordset
Dim oConec As New DConecta
Dim sSql As String
Dim rs As New ADODB.Recordset
On Error GoTo ObtenerUltimoCargoOficialErr
oConec.AbreConexion
    sSql = "Exec stp_sel_ObtieneUltimoCargoOficial '" & psPersCod & "'"
    Set ObtenerUltimoCargoOficial = oConec.CargaRecordSet(sSql)
oConec.CierraConexion
Set oConec = Nothing
Exit Function
ObtenerUltimoCargoOficialErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:ObtenerUltimoCargoOficial Method")
End Function 'NAGL ERS074-2017 20171212

'Agrega un nuevo cargo o nivel
'##ModelId=3A947AB400AA
Public Function GetUltCorr(ByVal psTipoCod As String, Optional psPersCod As String = "", Optional psCodAnt As String = "") As String
    On Error GoTo GetUltCorrErr
    Dim lsCodigo As String
    Dim sqlC As String
    Dim oCon As DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oCon = New DConecta
    
    If IsNumeric(psTipoCod) Then
        If psTipoCod = "1" Or psTipoCod = "0" Then
            psTipoCod = "E"
        ElseIf psTipoCod = "2" Then
            psTipoCod = "L"
        ElseIf psTipoCod = "3" Then
            psTipoCod = "F"
        ElseIf psTipoCod = "5" Then
            psTipoCod = "S"
        ElseIf psTipoCod = "6" Then
            psTipoCod = "D"
        Else
            psTipoCod = "P"
        End If
    End If
    
    
    sqlC = " Select Top 1 cRHCod From RRHH Where cRHCod like '" & psTipoCod & "%' Order by cRHCod DESC"
    
    If oCon.AbreConexion Then
        If psCodAnt = "" Then
            Set rs = oCon.CargaRecordSet(sqlC)
            If rs.EOF And rs.BOF Then
                lsCodigo = psTipoCod & "00001"
            Else
                lsCodigo = psTipoCod & FillNum(CInt(Right(rs.Fields(0), 5)) + 1, 5, "0")
            End If
            rs.Close
        Else
            lsCodigo = psTipoCod & Format(CCur(Right(psCodAnt, 5) + 1), "00000")
        End If
        If psPersCod <> "" Then
            sqlC = " Select cRHCod From RRHH Where cPersCod = '" & psPersCod & "' And cRHCod in ('201','101')"
            Set rs = oCon.Ejecutar(sqlC)
            If Not (rs.EOF And rs.BOF) Then
                lsCodigo = rs!cRHCod
            End If
            rs.Close
        End If
        GetUltCorr = lsCodigo
        oCon.CierraConexion
    End If
            
    Set oCon = Nothing
    Set rs = Nothing
    Exit Function
GetUltCorrErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:AgregaRRHH Method")
End Function

'Agrega un nuevo cargo o nivel
'##ModelId=3A947AB400AA
Public Function GetCodigoEmpleado(psPersCod As String) As String
    On Error GoTo GetUltCorrErr
    
    Dim sqlC As String
    Dim oCon As DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oCon = New DConecta
    
    sqlC = " Select Top 1 cRHCod Cod From RRHH Where cPersCod like '" & psPersCod & "' Order by cRHCod DESC"
    
    If oCon.AbreConexion Then
        Set rs = oCon.Ejecutar(sqlC)
        If rs.EOF And rs.BOF Then
            GetCodigoEmpleado = ""
        Else
            GetCodigoEmpleado = rs!Cod
        End If
        oCon.CierraConexion
    End If
            
    Set oCon = Nothing
    Set rs = Nothing
    Exit Function
GetUltCorrErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:AgregaRRHH Method")
End Function

'Agrega un nuevo cargo o nivel
'##ModelId=3A947AB400AA
Public Function GetCodigoEmpleadoPers(psEmpCod As String) As String
    On Error GoTo GetUltCorrErr
    
    Dim sqlC As String
    Dim oCon As DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oCon = New DConecta
    
    sqlC = " Select Top 1 cPersCod Cod From RRHH Where cRHCod like '" & psEmpCod & "' Order by cRHCod DESC"
    
    If oCon.AbreConexion Then
        Set rs = oCon.Ejecutar(sqlC)
        If rs.EOF And rs.BOF Then
            GetCodigoEmpleadoPers = ""
        Else
            GetCodigoEmpleadoPers = rs!Cod
        End If
        oCon.CierraConexion
    End If
            
    Set oCon = Nothing
    Set rs = Nothing
    Exit Function
GetUltCorrErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:AgregaRRHH Method")
End Function

'Elimina  RRHH, esta eliminacion es logica
'##ModelId=3A947AB400DC
Public Function EliminaRRHH(psPersCod As String) As Boolean
    On Error GoTo EliminaRRHHErr

    Dim sqlRH As String
    Dim sqlRHEmpleado As String
    Dim sqlRHEstado As String
    Dim sqlRHContrato As String
    Dim sqlRHContratoDet As String
    Dim sqlRHComentario As String
    Dim sqlRHSueldo As String
    Dim sqlRHCargos As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlRHEstado = " Delete RHEstado Where cPersCod = '" & psPersCod & "'"
    sqlRHContratoDet = " Delete RHContratoDet Where cPersCod = '" & psPersCod & "'"
    sqlRHContrato = " Delete RHContrato Where cPersCod = '" & psPersCod & "'"
    sqlRHComentario = " Delete RHComentario Where cPersCod = '" & psPersCod & "'"
    sqlRHEmpleado = " Delete RHEmpleado Where cPersCod = '" & psPersCod & "'"
    sqlRHSueldo = " Delete RHSueldo Where cPersCod = '" & psPersCod & "'"
    sqlRHCargos = " Delete RHCargos Where cPersCod = '" & psPersCod & "'"
    sqlRH = " Delete RRHH Where cPersCod = '" & psPersCod & "'"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlRHEstado
        oCon.Ejecutar sqlRHContratoDet
        oCon.Ejecutar sqlRHContrato
        oCon.Ejecutar sqlRHComentario
        oCon.Ejecutar sqlRHEmpleado
        oCon.Ejecutar sqlRHSueldo
        oCon.Ejecutar sqlRHCargos
        oCon.Ejecutar sqlRH
        oCon.CierraConexion
    End If
            
    Set oCon = Nothing
    Exit Function
EliminaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:EliminaRRHH Method")
End Function

'Actualiza un RRHH
'##ModelId=3A947AB40118
Public Function ModificaRRHH(psPersCod As String, psRHCod As String, psUser As String, psClave As String, psAreaCod As String, psAgenciaAsig As String, psAgenciaActual As String, psIngreso As String, psCese As String, psUltimaActualizacion As String, psEstado As String) As Boolean
    On Error GoTo ModificaRRHHErr
    
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If psCese = "" Then
        sqlC = " Update RRHH Set cRHCod = '" & psRHCod & "', cUser = '" & psUser & "', cAreaCod = '" & psAreaCod & "', cAgenciaAsig = '" & psAgenciaAsig & "', cAgenciaActual = '" & psAgenciaActual & "',dIngreso = '" & psIngreso & "',cUltimaActualizacion = '" & psUltimaActualizacion & "' " & IIf(psEstado <> "", ", RHEstado = '" & psEstado & "'", "") & "" _
             & " Where cPersCod = '" & psPersCod & "'"
    Else
        sqlC = " Update RRHH Set cRHCod = '" & psRHCod & "', cUser = '" & psUser & "', cAreaCod = '" & psAreaCod & "', cAgenciaAsig = '" & psAgenciaAsig & "', cAgenciaActual = '" & psAgenciaActual & "',dIngreso = '" & psIngreso & "',cUltimaActualizacion = '" & psUltimaActualizacion & "', dCese = '" & psCese & "'" & IIf(psEstado <> "", ", RHEstado = '" & psEstado & "'", "") & "" _
             & " Where cPersCod = '" & psPersCod & "'"
    End If
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
            
    Set oCon = Nothing
    Exit Function
ModificaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:ModificaRRHH Method")
End Function

'Actualiza un RRHH
'##ModelId=3A947AB40118
Public Function ModificaCuenatTarj(prRS As ADODB.Recordset, psTarj As String, psPersCod As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaRRHHErr
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlC = "Delete RHCuentas Where cPersCod  = '" & psPersCod & "'"
    
    If oCon.AbreConexion Then
        oCon.BeginTrans
            oCon.Ejecutar sqlC
            If Not prRS Is Nothing Then
                While Not prRS.EOF
                    sqlC = " Insert RHCuentas (cPersCod,cCtaCod,cUltimaActualizacion,cTarjCod)" _
                         & " Values('" & psPersCod & "','" & prRS.Fields(0) & "','" & psUltimaActualizacion & "', '" & Left(psTarj, 4) & Mid(psTarj, 6, 4) & Mid(psTarj, 11, 4) & Mid(psTarj, 16, 4) & "')"
                    oCon.Ejecutar sqlC
                    prRS.MoveNext
                Wend
            End If
        oCon.CommitTrans
        oCon.CierraConexion
    End If
            
    Set oCon = Nothing
    Exit Function
ModificaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:ModificaRRHH Method")
End Function


'Actualiza un RRHH
'##ModelId=3A947AB40118
'Public Function ModificaUsuario(psPersCod As String, psUser As String, psUltimoMovimiento As String, pnAgregaAPlanilla As Integer, pnCategoriaViaticos As Integer, pnMarcFotoCheck As Integer, pnMarcPC As Integer, pnConfianza As Integer, pnNSF As Integer, pnPersDireccion As Integer) As Boolean
'    On Error GoTo ModificaRRHHErr
'
'    Dim sqlC As String
'    Dim oCon As DConecta
'    Set oCon = New DConecta
'    ' pnConfianza As Integer, pnNSF As Integer
'
'    sqlC = " Update RRHH Set cUser = '" & psUser & "', cUltimaActualizacion = '" & psUltimoMovimiento & "', bAgregaPlanilla = " & Str(pnAgregaAPlanilla) & ", nRHCategoria = " & pnCategoriaViaticos & ", bCtrFotochek = " & Str(pnMarcFotoCheck) & ", bCtrPC = " & Str(pnMarcPC) & ", bConfianza = " & Str(pnConfianza) & ", bNoSujFiscalizacion = " & Str(pnNSF) & ", bPersDireccion =  " & Str(pnPersDireccion) & " " _
'             & " Where cPersCod = '" & psPersCod & "'"
'
'    If oCon.AbreConexion Then
'        oCon.Ejecutar sqlC
'        oCon.CierraConexion
'    End If
'
'    Set oCon = Nothing
'    Exit Function
'ModificaRRHHErr:
'    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:ModificaRRHH Method")
'End Function

Public Function ModificaUsuario(psPersCod As String, psUser As String, psUltimoMovimiento As String, pnAgregaAPlanilla As Integer, Optional psUbicacion As String = "") As Boolean
'Public Function ModificaUsuario(psPersCod As String, psUser As String, psUltimoMovimiento As String, pnAgregaAPlanilla As Integer) As Boolean
    On Error GoTo ModificaRRHHErr
    
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
'    sqlC = " Update RRHH Set cUser = '" & psUser & "', cUltimaActualizacion = '" & psUltimoMovimiento & "', bAgregaPlanilla = " & Str(pnAgregaAPlanilla) & " " _
'             & " Where cPersCod = '" & psPersCod & "'"
    sqlC = " Update RRHH Set cUser = '" & psUser & "', cUltimaActualizacion = '" & psUltimoMovimiento & "', bAgregaPlanilla = " & Str(pnAgregaAPlanilla) & ",cUbicacion= '" & psUbicacion & "'" _
             & " Where cPersCod = '" & psPersCod & "'"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
            
    Set oCon = Nothing
    Exit Function
ModificaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:ModificaRRHH Method")
End Function


Public Function ModificaAgenciaArea(psPersCod As String, psArea As String, psAgencia As String, psUltimoMovimiento As String, psFecha As String, psComentario As String, Optional pbActual As Boolean = True) As Boolean
    On Error GoTo ModificaRRHHErr
    Dim sqlC As String
    Dim sqlI As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If pbActual Then
        sqlC = " Update RRHH Set cAreaCod = '" & psArea & "', cAgenciaAsig = '" & psAgencia & "', cUltimaActualizacion = '" & psUltimoMovimiento & "'" _
                 & " Where cPersCod = '" & psPersCod & "'"
        sqlI = " Insert RHAreaAgencia (cPersCod,dFecha,cAreaCod,cAgenciaAsig,cComentario,cUltimaActualizacion)" _
             & " Values('" & psPersCod & "','" & psFecha & "','" & psArea & "','" & psAgencia & "','" & psComentario & "','" & psUltimoMovimiento & "')"
    Else
        sqlC = " Update RRHH Set cAreaCodActual = '" & psArea & "', cAgenciaActual = '" & psAgencia & "', cUltimaActualizacion = '" & psUltimoMovimiento & "'" _
                 & " Where cPersCod = '" & psPersCod & "'"
        sqlI = " Insert RHAgencia (cPersCod,dFecha,cAreaCodActual,cAgenciaActual,cComentario,cUltimaActualizacion)" _
             & " Values('" & psPersCod & "','" & psFecha & "','" & psArea & "','" & psAgencia & "','" & psComentario & "','" & psUltimoMovimiento & "')"
    End If
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.Ejecutar sqlI
        oCon.CierraConexion
    End If
            
    Set oCon = Nothing
    Exit Function
ModificaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:ModificaRRHH Method")
End Function

Public Function ModificaCUSSPSeguroDiasVacac(psPersCod As String, psCUSPP As String, psSeguro As String, pdFecIng As Date, psUltimoMovimiento As String, psFecCese As String, psDiasVacac As String) As Boolean
    On Error GoTo ModificaRRHHErr
    Dim sqlC As String
    Dim sqlI As String
    Dim oCon As DConecta
    Dim sqlAdd As String
    Set oCon = New DConecta
    
    If IsDate(psFecCese) Then
        sqlAdd = " '" & Format(psFecCese, gsFormatoFecha) & "' "
    Else
        sqlAdd = " Null "
    End If
    
    sqlC = " Update RHEmpleado Set cCUSPP = '" & psCUSPP & "' , cNumAFP = '" & psSeguro & "', cUltimaActualizacion = '" & psUltimoMovimiento & "', nRHEmplVacacionesPend = " & psDiasVacac & "" _
         & " Where cPersCod = '" & psPersCod & "'"
    
    sqlI = " Update RRHH " _
         & " Set dIngreso = '" & Format(pdFecIng, gsFormatoFecha) & "', dCese = " & sqlAdd & "" _
         & " Where cPersCod = '" & psPersCod & "'"
         
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.Ejecutar sqlI
        oCon.CierraConexion
    End If
            
    Set oCon = Nothing
    Exit Function
ModificaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:ModificaRRHH Method")
End Function

Public Function GetCUSSPSeguroDiasVac(psPersCod As String) As ADODB.Recordset
    On Error GoTo ModificaRRHHErr
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlC = " Select cCUSPP , cNumAFP, nRHEmplVacacionesPend From  RHEmpleado " _
         & " Where cPersCod = '" & psPersCod & "'"
    
    If oCon.AbreConexion Then
        Set GetCUSSPSeguroDiasVac = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
            
    Set oCon = Nothing
    Exit Function
ModificaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:ModificaRRHH Method")
End Function

Public Function ModificaContratoTextoUlt(psPersCod As String, psTexto As String, psUltimoMovimiento As String) As Boolean
    On Error GoTo ModificaRRHHErr
    Dim sqlC As String
    Dim oCon As DConecta
    Dim rs As ADODB.Recordset
    Dim lsFecha As String
    Set rs = New ADODB.Recordset
    Set oCon = New DConecta
        
    sqlC = " Select Top 1 cRHContratoNro Con from RHContrato Where cPersCod = '" & psPersCod & "' Order by cRHContratoNro Desc"
        
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sqlC)
        sqlC = " Update RHContrato" _
             & " Set tRHContratoTexto = '" & psTexto & "', cUltimaActualizacion = '" & psUltimoMovimiento & "'" _
             & " Where cPersCod = '" & psPersCod & "' And cRHContratoNro = '" & rs!CON & "'"
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
            
    rs.Close
    Set rs = Nothing
    Set oCon = Nothing
    Exit Function
ModificaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:ModificaRRHH Method")
End Function

Public Function ModificaFoto(psPersCod As String, picFoto As Variant, psUltimoMovimiento As String, Optional psRuta As String = "") As Byte
    On Error GoTo ModificaRRHHErr
    Dim sqlC As String
    Dim oCon As DConecta
    Dim rs As ADODB.Recordset
    
    Dim lsFecha As String
    Set rs = New ADODB.Recordset
    Set oCon = New DConecta
        
    sqlC = " Select iPersFoto Foto From PersFoto Where cPersCod = '" & psPersCod & "'"
        
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sqlC) ', adLockOptimistic, , False)
        
        If rs.EOF And rs.BOF Then
            sqlC = " Insert PersFoto (cPersCod,cUltimaActualizacion)" _
                 & " Values ('" & psPersCod & "','" & psUltimoMovimiento & "')"
            oCon.Ejecutar sqlC
            
            rs.Close
            sqlC = " Select iPersFoto Foto From PersFoto Where cPersCod = '" & psPersCod & "'"
            Set rs = oCon.CargaRecordSet(sqlC) ', adLockOptimistic, , False)
            LetPicture rs.Fields(0), picFoto
            rs.Update
        Else
            'Dim RStream As New ADODB.Stream
            'Set RStream = New ADODB.Stream
            'RStream.Type = adTypeBinary
            'RStream.Open
            'RStream.LoadFromFile App.Path & "\RRHH\iiikkk.bmp"
            'rs.Fields("iPersFoto").Value = RStream.Read
            'RStream.Close
            'Set RStream = Nothing
            'LetPicture1 rs!Foto, picFoto
            LetPicture rs!Foto, picFoto
            rs.Update
            sqlC = "Update PersFoto Set cUltimaActualizacion = '" & psUltimoMovimiento & "' WHere cPersCod = '" & psPersCod & "'"
            oCon.Ejecutar sqlC
        End If
        oCon.CierraConexion
    End If
            
    'rs.Close
    Set rs = Nothing
    Set oCon = Nothing
    Exit Function
ModificaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:ModificaRRHH Method")
End Function

Public Function GetFoto(psPersCod As String) As ADODB.Recordset
    On Error GoTo ModificaRRHHErr
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
        
    sqlC = " Select iPersFoto Foto From PersFoto Where cPersCod = '" & psPersCod & "'"
        
    If oCon.AbreConexion Then
        Set GetFoto = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
            
    Set oCon = Nothing
    Exit Function
ModificaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:ModificaRRHH Method")
End Function


Public Function ModificaContratoTextoUltDet(psPersCod As String, prRS As ADODB.Recordset, psUltimoMovimiento As String, pgdFecSis As Date, pgsFormatoFechaHora As String) As Boolean
    On Error GoTo ModificaRRHHErr
    Dim sqlC As String
    Dim oCon As DConecta
    Dim rs As ADODB.Recordset
    Dim lsFecha As String
    Set rs = New ADODB.Recordset
    Set oCon = New DConecta
        
    sqlC = " Select Top 1 cRHContratoNro Con from RHContrato Where cPersCod = '" & psPersCod & "' Order by Desc"
    
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sqlC)
        
        While Not prRS.EOF
            lsFecha = Format(pgdFecSis & " " & GetHoraServer, pgsFormatoFechaHora)
            sqlC = " Insert RHContratoDet(cPersCod, cRHContratoNro, dRHContratoFecha, dRHContratoInicio, dRHContratoFin, cRHContratoComentario, cUltimaActualizacion)" _
                 & " Values ('" & psPersCod & "','" & rs!CON & "','" & lsFecha & "','" & Format(prRS.Fields(2), pgsFormatoFechaHora) & "',,'" & Format(prRS.Fields(2), pgsFormatoFechaHora) & "','" & prRS.Fields(3) & "','" & psUltimoMovimiento & "')"
            oCon.Ejecutar sqlC
            prRS.MoveNext
        Wend
        oCon.CierraConexion
    End If
            
    Set oCon = Nothing
    Exit Function
ModificaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:ModificaRRHH Method")
End Function


'Obtiene a todo el RRHH de la Empresa
'
'0 Todos
'... de acuerdo a tipo de estado del empleado
'##ModelId=3A947ABB010E
Public Function GetRRHHs(psEstadoRRHH As String) As ADODB.Recordset
    On Error GoTo GetRRHHsErr

    'your code goes here...

    Exit Function
GetRRHHsErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHHs Method")
End Function

'##ModelId=3A947F2000B1
Public Function GetRRHHBuscado(psEmpCod As String, psNomPers As String, psCodPers As String, psNumDoc As String) As ADODB.Recordset
    On Error GoTo GetRRHHBuscadoErr

    'your code goes here...

    Exit Function
GetRRHHBuscadoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHHBuscado Method")
End Function

'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
'lnTipoControl
'0 Ninguno
'1 Asistencia
'2 Vacaciones
'3 Permsisos
Public Function GetRRHH(psRRHHCod As String, psTpoDocID As PersIdTipo, Optional lnTipoControl As Integer = 0) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Dim sqlAuxT As String
    Dim sqlAuxC As String
    
    Set oCon = New DConecta
    
    On Error GoTo GetRRHHErr
    
    If lnTipoControl = 0 Then
        sqlAuxT = ""
        sqlAuxC = ", 1 Control"
    ElseIf lnTipoControl = 1 Then
        sqlAuxT = " Inner Join RHTipoContrato RHT ON RHT.cRHContratoCod =  '" & Left(psRRHHCod, 1) & "' "
        sqlAuxC = ", bRHControlAsistencia Control"
    ElseIf lnTipoControl = 2 Then
        sqlAuxT = " Inner Join RHTipoContrato RHT ON RHT.cRHContratoCod =  '" & Left(psRRHHCod, 1) & "' "
        sqlAuxC = ", bRHControlVacaciones Control"
    ElseIf lnTipoControl = 3 Then
        sqlAuxT = " Inner Join RHTipoContrato RHT ON RHT.cRHContratoCod =  '" & Left(psRRHHCod, 1) & "' "
        sqlAuxC = ", bRHControlPermisos Control"
    End If
    
    sqlC = " Select RH.cPersCod as Codigo, PE.cPersNombre as Nombre, PE.cPersDireccDomicilio as Direccion," _
         & " PID.cPersIDnro as ID, PE.dPersNacCreac as Fecha, Sueldo.nRHSueldoMonto as Sueldo, RH.dIngreso Ingreso " & sqlAuxC & " " _
         & " From RRHH RH" _
         & " Inner Join Persona PE On RH.cPersCod = PE.cPersCod" _
         & " Left Join PersID PID On PE.cPersCOd = PID.cPersCod and cPersIDtpo = " & psTpoDocID & "" _
         & " Left Join (Select Top 1 RS.cPersCod as Cod,nRHSueldoMonto from RHSueldo RS" _
         & "             Inner Join RRHH RH On RS.cPersCod = RH.cPersCod" _
         & "             Where RH.cRHCod = '" & psRRHHCod & "' Order By  RS.dRHSueldoFecha desc" _
         & "            ) As Sueldo On Sueldo.Cod = RH.cPersCod" _
         & sqlAuxT _
         & " where RH.cRHCod = '" & psRRHHCod & "'"
    If oCon.AbreConexion Then
        Set GetRRHH = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function

Public Function GetRRHHArea(psAreaCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    On Error GoTo GetRRHHErr
    
    sqlC = " Select RH.cPersCod as Codigo, PE.cPersNombre as Nombre" _
         & " From RRHH RH" _
         & " Inner Join Persona PE On RH.cPersCod = PE.cPersCod" _
         & " where RH.cAreaCod = '" & psAreaCod & "' And nRHEstado not like '" & RHEstadosTpoRetirado & "%'"
    If oCon.AbreConexion Then
        Set GetRRHHArea = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function

Public Function GetRRHHAgenciaActual(psPersCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    On Error GoTo GetRRHHErr
    
    sqlC = " Select AG.cAgeCod, AG.cAgeDescripcion " _
         & " From RRHH RH" _
         & " Inner Join Agencias AG On RH.cAgenciaActual = AG.cAgeCod" _
         & " where RH.cPersCod = '" & psPersCod & "'"
    If oCon.AbreConexion Then
        Set GetRRHHAgenciaActual = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function

'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A


Public Function CargarZonas(ByVal psPersCod As String) As ADODB.Recordset
Dim oCon As New DConecta, sqlZ As String, nIndex As Integer, rs As New ADODB.Recordset
On Error GoTo getZonas

Set CargarZonas = Nothing

sqlZ = "select t.nConsValor,t.cConsDescripcion,nItem=coalesce(z.nZonaCod,0)  from " & _
     " (select nConsValor,cConsDescripcion from constante where nConsCod = 1046 and nConsCod<>nConsValor) t left join " & _
     " (select nZonaCod from RHZonasTerritorialesAdm where cPersCod = '" & psPersCod & "') z on t.nConsValor = z.nZonaCod "

nIndex = -1
If oCon.AbreConexion Then
   Set CargarZonas = oCon.CargaRecordSet(sqlZ)
End If
Exit Function
getZonas:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function


'Public Function GetRRHHGeneralidades(psRRHHCod As String) As ADODB.Recordset
'    Dim oCon As DConecta
'    Dim sqlC As String
'    Set oCon = New DConecta
'    On Error GoTo GetRRHHGeneralidadesErr
'
'    sqlC = " Select RH.cUser Usu, RH.bAgregaPlanilla AgregaPlanilla, CO.cConsDescripcion Estado, ARE1.cAreaDescripcion AreaAsig, ARE2.cAreaDescripcion AreaAct, AGEAsig.cAgeDescripcion Asig,AGEAct.cAgeDescripcion Act, Vista.cRHCargoDescripcion,Vista. bRHZonaTerritorial, (Select dPersNacCreac From Persona PE Where PE.cPersCod = RH.cPersCod) FecNac, RH.nRHCategoria, RH.bCtrFotochek, RH.bCtrPC, RH.bConfianza, RH.bNoSujFiscalizacion, RH.bPersDireccion From RRHH RH" _
'         & " Inner Join Constante CO On CO.nConsValor = RH.nRHEstado And nConsCod = '" & gRHEstado & "'" _
'         & " Inner Join Agencias AGEAsig On AGEAsig.cAgeCod = RH.cAgenciaAsig" _
'         & " Inner Join Agencias AGEAct On AGEAct.cAgeCod = RH.cAgenciaActual" _
'         & " Inner join Areas ARE1 On ARE1.cAreaCod = RH.cAreaCod" _
'         & " Left join Areas ARE2 On ARE2.cAreaCod = RH.cAreaCodActual" _
'         & " Inner Join (Select Top 1 RHC.cPersCod, cRHCargoDescripcion, bRHZonaTerritorial from RHCargos RHC" _
'         & "             Inner Join RHCargosTabla RHCT On RHC.cRHCargoCod = RHCT.cRHCargoCod" _
'         & "             Where RHC.cPersCod = '" & psRRHHCod & "'" _
'         & "             order by dRHCargoFecha desc) AS Vista On Vista.cPersCod = RH.cPersCod" _
'         & " Where rh.cPersCod = '" & psRRHHCod & "'"
'
'    If oCon.AbreConexion Then
'        Set GetRRHHGeneralidades = oCon.CargaRecordSet(sqlC)
'        oCon.CierraConexion
'    End If
'    Set oCon = Nothing
'    Exit Function
'GetRRHHGeneralidadesErr:
'    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
'End Function

Public Function GetRRHHGeneralidades(psRRHHCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo GetRRHHGeneralidadesErr
'ALPA 20110204
'    sqlC = " Select RH.cUser Usu, RH.bAgregaPlanilla AgregaPlanilla, CO.cConsDescripcion Estado, ARE1.cAreaDescripcion AreaAsig, ARE2.cAreaDescripcion AreaAct, AGEAsig.cAgeDescripcion Asig, AGEAct.cAgeDescripcion Act, Vista.cRHCargoDescripcion, (Select dPersNacCreac From Persona PE Where PE.cPersCod = RH.cPersCod) FecNac From RRHH RH" _
'         & " Inner Join Constante CO On CO.nConsValor = RH.nRHEstado And nConsCod = '" & gRHEstado & "'" _
'         & " Inner Join Agencias AGEAsig On AGEAsig.cAgeCod = RH.cAgenciaAsig" _
'         & " Inner Join Agencias AGEAct On AGEAct.cAgeCod = RH.cAgenciaActual" _
'         & " Inner join Areas ARE1 On ARE1.cAreaCod = RH.cAreaCod" _
'         & " Left join Areas ARE2 On ARE2.cAreaCod = RH.cAreaCodActual" _
'         & " Inner Join (Select Top 1 RHC.cPersCod, cRHCargoDescripcion from RHCargos RHC" _
'         & "             Inner Join RHCargosTabla RHCT On RHC.cRHCargoCod = RHCT.cRHCargoCod" _
'         & "             Where RHC.cPersCod = '" & psRRHHCod & "'" _
'         & "             order by dRHCargoFecha desc) AS Vista On Vista.cPersCod = RH.cPersCod" _
'         & " Where rh.cPersCod = '" & psRRHHCod & "'"
    
    sqlC = " Select RH.cUser Usu, RH.bAgregaPlanilla AgregaPlanilla, CO.cConsDescripcion Estado, ARE1.cAreaDescripcion AreaAsig, ARE2.cAreaDescripcion AreaAct, AGEAsig.cAgeDescripcion Asig, AGEAct.cAgeDescripcion Act, Vista.cRHCargoDescripcion, (Select dPersNacCreac From Persona PE Where PE.cPersCod = RH.cPersCod) FecNac,isnull(RH.cUbicacion,'') cUbicacion  From RRHH RH" _
         & " Inner Join Constante CO On CO.nConsValor = RH.nRHEstado And nConsCod = '" & gRHEstado & "'" _
         & " Inner Join Agencias AGEAsig On AGEAsig.cAgeCod = RH.cAgenciaAsig" _
         & " Inner Join Agencias AGEAct On AGEAct.cAgeCod = RH.cAgenciaActual" _
         & " Inner join Areas ARE1 On ARE1.cAreaCod = RH.cAreaCod" _
         & " Left join Areas ARE2 On ARE2.cAreaCod = RH.cAreaCodActual" _
         & " Inner Join (Select Top 1 RHC.cPersCod, cRHCargoDescripcion from RHCargos RHC" _
         & "             Inner Join RHCargosTabla RHCT On RHC.cRHCargoCod = RHCT.cRHCargoCod" _
         & "             Where RHC.cPersCod = '" & psRRHHCod & "'" _
         & "             order by dRHCargoFecha desc) AS Vista On Vista.cPersCod = RH.cPersCod" _
         & " Where rh.cPersCod = '" & psRRHHCod & "'"
    
    If oCon.AbreConexion Then
        Set GetRRHHGeneralidades = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHGeneralidadesErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function

'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
Public Function GetRRHHCuentas(psRRHHCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo GetRRHHGeneralidadesErr
    
    sqlC = " Select cCtaCod, '' aaa from RhCuentas Where cPersCod = '" & psRRHHCod & "'"
    
    If oCon.AbreConexion Then
        Set GetRRHHCuentas = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHGeneralidadesErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function

'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
Public Function GetRRHHTarjeta(psRRHHCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo GetRRHHGeneralidadesErr
    
    sqlC = " Select cTarjCod from rhcuentas where cPersCod = '" & psRRHHCod & "'"
    
    If oCon.AbreConexion Then
        Set GetRRHHTarjeta = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHGeneralidadesErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function


'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
Public Function GetRRHHGeneralidadesDoc(psRRHHCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo GetRRHHGeneralidadesErr
    
    sqlC = " Select cConsDescripcion Relacion, cPersIDnro Nro from PersID PR " _
         & " Inner Join Constante CO On Co.nConsValor = PR.cPersIDTpo And CO.nConsCod = '1003'" _
         & " " _
         & " Where PR.cPersCod = '" & psRRHHCod & "'"
    
    If oCon.AbreConexion Then
        Set GetRRHHGeneralidadesDoc = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHGeneralidadesErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function

'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
Public Function GetRRHHFamiliares(psRRHHCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo GetRRHHGeneralidadesErr
    
    sqlC = " Select cPersNombre Nombre , cConsDescripcion Relacion from Persona PE " _
         & " Inner Join PersRelaciones PR On PE.cPersCod = PR.cPersRelacPersCod" _
         & " Inner Join Constante CO On Co.nConsValor = PR.nPersRelac And CO.nConsCod = " & gPersRelacion & "" _
         & " Where PR.cPersCod = '" & psRRHHCod & "'"
    
    If oCon.AbreConexion Then
        Set GetRRHHFamiliares = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHGeneralidadesErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function

'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
Public Function GetRRHHContratos(psRRHHCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo GetRRHHGeneralidadesErr
    
    sqlC = " Select  RHC.cRHContratoNro, CO.cConsDescripcion , Convert(varchar(10),RHC.dFechaFin,103) , Convert(varchar(10),RHCD.dRHContratoInicio,103), Convert(varchar(10),RHCD.dRHContratoFin,103)," _
         & " AgeAsig.area, AgeAsig.Age , AgeActual.area, AgeActual.Age, Cargo.Cargo Cargo, Convert(decimal(8,2),Sueldo.nRHSueldoMonto) Sueldo, RHCD.cRHContratoComentario" _
         & " from RHContrato RHC" _
         & " Inner Join RHContratoDet RHCD On RHC.cRHContratoNro = RHCD.cRHContratoNro" _
         & "     And RHC.cPersCod = RHCD.cPersCod" _
         & " LEft join Constante CO On RHC.nRHContratoTpo = CO.nConsValor And CO.nConsCod = " & gRHTipoContrato & "" _
         & " left Join (Select  cRHContratoNro, nRHSueldoMonto  from rhSueldo RHS" _
         & "             Inner Join (Select Max(RHSS.dRHSueldoFecha) Fecha, cRHContratoNro, RHSS.cPersCod  From RHSueldo RHSS" _
         & "             Inner Join RHContratoDet RHCD On RHCD.cPersCod = RHSs.cPersCod" _
         & "             Where RHSS.dRHSueldoFecha Between RHCD.dRHContratoInicio" _
         & "                 And RHCD.dRHContratoFin and RHCD.cPersCod  = '" & psRRHHCod & "'" _
         & "             group by cRHContratoNro, RHSS.cPersCod) RHCD" _
         & "         On RHCD.cPersCod = RHS.cPersCod And RHCD.Fecha = RHS.dRHSueldoFecha" _
         & "     ) Sueldo On Sueldo.cRHContratoNro = RHC.cRHContratoNro" _
         & " Left Join (Select  cRHContratoNro, Area.cAreaDescripcion Area, Age.cAgeDescripcion Age from rhCargos RHS" _
         & "      Inner Join (Select Max(RHSS.dRHCargoFecha) Fecha, cRHContratoNro, RHSS.cPersCod  From RHCargos RHSS" _
         & "             Inner Join RHContratoDet RHCD On RHCD.cPersCod = RHSs.cPersCod" _
         & "             Where RHSS.dRHCargoFecha Between RHCD.dRHContratoInicio" _
         & "                 And RHCD.dRHContratoFin and RHCD.cPersCod  = '" & psRRHHCod & "'" _
         & "             group by cRHContratoNro, RHSS.cPersCod) RHCD" _
         & "         On RHCD.cPersCod = RHS.cPersCod And RHCD.Fecha = RHS.dRHCargoFecha" _
         & "     Inner Join Areas Area On Area.cAreaCod = RHS.cRHAreaCodOficial" _
         & "     Inner Join Agencias Age On Age.cAgeCod = RHS.cRHAgenciaCodOficial" _
         & "     ) AgeActual On AgeActual.cRHContratoNro = RHC.cRHContratoNro"
        
        sqlC = sqlC & " Left Join (Select  cRHContratoNro, Area.cAreaDescripcion Area, Age.cAgeDescripcion Age from rhCargos RHS" _
             & " Inner Join (Select Max(RHSS.dRHCargoFecha) Fecha, cRHContratoNro, RHSS.cPersCod  From rhCargos RHSS" _
             & "         Inner Join RHContratoDet RHCD On RHCD.cPersCod = RHSs.cPersCod" _
             & "         Where RHSS.dRHCargoFecha Between RHCD.dRHContratoInicio" _
             & "             And RHCD.dRHContratoFin and RHCD.cPersCod  = '" & psRRHHCod & "'" _
             & "         group by cRHContratoNro, RHSS.cPersCod) RHCD" _
             & "     On RHCD.cPersCod = RHS.cPersCod And RHCD.Fecha = RHS.dRHCargoFecha" _
             & " Inner Join Areas Area On Area.cAreaCod = RHS.cRHAreaCod" _
             & " Inner Join Agencias Age On Age.cAgeCod = RHS.cRHAgenciaCod" _
             & " ) AgeAsig On AgeAsig.cRHContratoNro = RHC.cRHContratoNro" _
             & " Left Join (Select  cRHContratoNro, cRHCargoDescripcion Cargo from rhCargos RHS" _
             & "            Inner Join (Select Max(RHSS.dRHCargoFecha) Fecha, cRHContratoNro, RHSS.cPersCod  From rhCargos RHSS" _
             & "            Inner Join RHContratoDet RHCD On RHCD.cPersCod = RHSs.cPersCod" _
             & "        Where RHSS.dRHCargoFecha Between RHCD.dRHContratoInicio" _
             & "            And RHCD.dRHContratoFin and RHCD.cPersCod  = '" & psRRHHCod & "'" _
             & "        group by cRHContratoNro, RHSS.cPersCod) RHCD" _
             & "    On RHCD.cPersCod = RHS.cPersCod And RHCD.Fecha = RHS.dRHCargoFecha" _
             & "        Inner Join RHCargosTabla RHT On RHT.cRHCargoCod = RHS.cRHCargoCod" _
             & "    ) Cargo On Cargo.cRHContratoNro = RHC.cRHContratoNro" _
             & " Where dRHContratoFecha = (select Max(dRHContratoFecha) From RHContratoDet A Where a.cRHContratoNro = RHCD.cRHContratoNro And a.cPersCod = '" & psRRHHCod & "')" _
             & " and RHC.cPersCod  =  '" & psRRHHCod & "' Order by RHC.cRHContratoNro Desc"
    
    If oCon.AbreConexion Then
        Set GetRRHHContratos = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHGeneralidadesErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function

'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
Public Function GetRRHHContratoTexto(psPersCod As String, Optional psContrato As String = "") As String
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    On Error GoTo GetRRHHEmpleadoErr
    
    If psContrato = "" Then
        sqlC = " Select Top 1 tRHContratoTexto Cont from rhcontrato Where cPersCod = '" & psPersCod & "' Order By cRHContratoNro Desc"
    Else
        sqlC = " Select Top 1 tRHContratoTexto Cont from rhcontrato Where cPersCod = '" & psPersCod & "' And cRHContratoNro = '" & psContrato & "' Order By cRHContratoNro Desc"
    End If
    
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sqlC)
        If rs.EOF And rs.BOF Then
            GetRRHHContratoTexto = ""
        ElseIf IsNull(rs.Fields(0)) Then
            GetRRHHContratoTexto = ""
        Else
            GetRRHHContratoTexto = rs.Fields(0)
        End If
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Set rs = Nothing
    Exit Function
GetRRHHEmpleadoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function


'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
Public Function GetRRHHAdendatTexto(psPersCod As String, psContrato As String, psFecha As String) As String
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    On Error GoTo GetRRHHEmpleadoErr
    
    sqlC = " Select cRHContratoComentario Con from RHContratoDet " _
         & " Where cPersCod = '" & psPersCod & "' And cRHContratoNro = '" & psContrato & "' And Convert(Varchar(10),dRHContratoFecha,103) = '" & psFecha & "'"
    
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sqlC)
        If rs.EOF And rs.BOF Then
            GetRRHHAdendatTexto = ""
        ElseIf IsNull(rs.Fields(0)) Then
            GetRRHHAdendatTexto = ""
        Else
            GetRRHHAdendatTexto = rs.Fields(0)
        End If
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Set rs = Nothing
    Exit Function
GetRRHHEmpleadoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function


'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
Public Function GetRRHHEmpleado(psRRHHCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo GetRRHHEmpleadoErr
    
    sqlC = " Select cRHEmpNivRem Nivel, cRHEmplAMPCod AMP, cRHEmplAFPPersCod AFP,  " _
         & " nRHEmplVacacionesPend VacPen, dEstabilidad dEstab, bRHEmplCondicion Condi, PID.cPersIDnro CUSPP" _
         & " from rhEmpleado RHE" _
         & " Left Join PersID PID On RHE.cPersCod = PID.cPersCod And cPersIDTpo  = '9'" _
         & " Where RHE.cPersCod = '" & psRRHHCod & "'"
    If oCon.AbreConexion Then
        Set GetRRHHEmpleado = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetRRHHEmpleadoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function

'Agrega datos adicionales del los RRHH de tipo Empleados
'##ModelId=3A9489A80297
Public Function AgregaRRHHEmp(psCodPers As String, psRHNivelCod As String, psRHAMP As String, pnRHAFP As String, pnRHDiasVacaciones As Integer, psContratado As String, psFechaEstable As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaRRHHEmpErr
    Dim sqlR As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlR = " Insert RHEmpleado (cPersCod,cRHEmpNivRem,cRHEmplAMPCod,cRHEmplAFPPersCod,nRHEmplVacacionesPend,bRHEmplCondicion,dEstabilidad,cUltimaActualizacion)" _
         & " Values('" & psCodPers & "','" & psRHNivelCod & "','" & psRHAMP & "','" & pnRHAFP & "'," & pnRHDiasVacaciones & "," & psContratado & ",'" & psFechaEstable & "','" & psUltimaActualizacion & "')"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlR
        oCon.CierraConexion
    End If

    Set oCon = Nothing
    Exit Function
AgregaRRHHEmpErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:AgregaRRHHEmp Method")
End Function



'Elimina  datos adicionales del los RRHH de tipo Empleados
'esta eliminacion es logica
'
'##ModelId=3A9489A802C9
Public Function EliminaRRHHEmp(psRHCargoCod As String) As Boolean
    On Error GoTo EliminaRRHHEmpErr

    'your code goes here...

    Exit Function
EliminaRRHHEmpErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:EliminaRRHHEmp Method")
End Function

'Actualiza datos adicionales del los RRHH de tipo Empleados
'##ModelId=3A9489A802FC
Public Function ModificaRRHHEmp(psCodPers As String, psRHNivelCod As String, psRHAMP As String, psRHAFP As String, pnRHDiasVacaciones As Integer, psContratado As String, psFechaEstable As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaRRHHEmpErr
    Dim sqlR As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If psFechaEstable <> "" Then
        sqlR = " Update RHEmpleado" _
             & " Set cRHEmpNivRem = '" & psRHNivelCod & "', cRHEmplAMPCod = '" & psRHAMP & "', cRHEmplAFPPersCod = '" & psRHAFP & "', nRHEmplVacacionesPend = " & pnRHDiasVacaciones & ", bRHEmplCondicion = " & psContratado & " ,dEstabilidad = '" & psFechaEstable & "', cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
             & " Where cPersCod = '" & psCodPers & "'"
    Else
        sqlR = " Update RHEmpleado" _
             & " Set cRHEmpNivRem = '" & psRHNivelCod & "', cRHEmplAMPCod = '" & psRHAMP & "', cRHEmplAFPPersCod = '" & psRHAFP & "', nRHEmplVacacionesPend = " & pnRHDiasVacaciones & ", bRHEmplCondicion = " & psContratado & " , cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
             & " Where cPersCod = '" & psCodPers & "'"
    End If
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlR
        oCon.CierraConexion
    End If

    Set oCon = Nothing
    Exit Function
ModificaRRHHEmpErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:ModificaRRHHEmp Method")
End Function

'Devuelve el Cargo
'##ModelId=3A96F8200032
Public Function GetEmpleadoEstado(psPersCod As String) As String
    On Error GoTo GetCargoErr
    Dim sqlC As String
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta

    sqlC = " Select CO.cConsDescripcion Estado from RRHH RH" _
         & " Inner Join Constante CO On CO.nConsValor = RH.nRHEstado And nConsCod = '" & gRHEstado & "'" _
         & " Where RH.cPersCod = '" & psPersCod & "'"
        
    If oCon.AbreConexion Then
        Set rsC = oCon.CargaRecordSet(sqlC)
        If rsC.EOF And rsC.BOF Then
            GetEmpleadoEstado = ""
        Else
            GetEmpleadoEstado = rsC!Cargo
        End If
    End If
        
    rsC.Close
    Set rsC = Nothing
    Set oCon = Nothing
    Exit Function
GetCargoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetCargo Method")
End Function



'Devuelve el Cargo
'##ModelId=3A96F8200032
Public Function GetCargo(psPersCod As String, psPeriodo As String) As String
    On Error GoTo GetCargoErr
    
    Dim sqlC As String
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta

    sqlC = " Select Top 1 RHCT.cRHCargoDescripcion Cargo, isnull(RH.bMuestraCargo,'')bMuestraCargo From RHCargosTabla RHCT" _
         & " Inner Join RHCargos RHC On RHCT.cRHCargoCod = RHC.cRHCargoCod And Convert(Varchar(10),dRHCargoFecha,112) <= '" & psPeriodo & "'" _
         & " Inner Join RRHH RH On RHC.cPersCod = RH.cPersCod  " _
         & " Where RH.cPersCod = '" & psPersCod & "' Order by dRHCargoFecha Desc"
        
    If oCon.AbreConexion Then
        Set rsC = oCon.CargaRecordSet(sqlC)
        If rsC.EOF And rsC.BOF Then
            GetCargo = ""
        Else
            If rsC!bMuestraCargo Then
                GetCargo = rsC!Cargo
            Else
                GetCargo = ""
            End If
        End If
    End If
        
    rsC.Close
    Set rsC = Nothing
    Set oCon = Nothing
    Exit Function
GetCargoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetCargo Method")
End Function


'Devuelve el Cargo
'##ModelId=3A96F8200032
Public Function GetCargoRS(psPersCod As String, psPeriodo As String) As ADODB.Recordset
    On Error GoTo GetCargoErr
    
    Dim sqlC As String
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta

    sqlC = " Select Top 1 RHCT.cRHCargoDescripcion Cargo, RH.bMuestraCargo, RHCT.bRHCtrAsist From RHCargosTabla RHCT" _
         & " Inner Join RHCargos RHC On RHCT.cRHCargoCod = RHC.cRHCargoCodOficial And Convert(Varchar(10),dRHCargoFecha,112) <= '" & psPeriodo & "'" _
         & " Inner Join RRHH RH On RHC.cPersCod = RH.cPersCod  " _
         & " Where RH.cPersCod = '" & psPersCod & "' Order by dRHCargoFecha Desc"
        
    If oCon.AbreConexion Then
        Set rsC = oCon.CargaRecordSet(sqlC)
        If rsC.EOF And rsC.BOF Then
            Set GetCargoRS = Nothing
        Else
            Set GetCargoRS = rsC
        End If
    End If
        
'    rsC.Close
'    Set rsC = Nothing
'    Set oCon = Nothing
    Exit Function
GetCargoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetCargo Method")
End Function


'Devuelve el Cargo
'##ModelId=3A96F8200032
Public Function GetCuentaRRHH(psPersCod As String, psPlanillaTpo As String, Optional psPlanillaPeriodo As String) As String
    On Error GoTo GetCargoErr
   
    Dim sqlC As String
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Dim oCon As DConecta
    Dim lsCadTemp As String
    Set oCon = New DConecta
    Dim lbBitCentral As Boolean
    
    lsCadTemp = ""
  
    Dim oConst As NConstSistemas
    Set oConst = New NConstSistemas
    
    Dim oConexionRem As DConecta
    Set oConexionRem = New DConecta
    Dim rsValida As ADODB.Recordset
    Set rsValida = New ADODB.Recordset
    Dim sql As String
    
    lbBitCentral = IIf(oConst.LeeConstSistema(gConstSistBitCentral) = "1", True, False)

    If lbBitCentral Then
        If psPlanillaTpo = gsRHPlanillaCTS Then
            sqlC = " Select RH.cCtaCod from RHCuentas RH" _
                 & " Inner Join " & lsCadTemp & "Producto PR On RH.cCtaCod = PR.cCtaCod" _
                 & " Where cPersCod = '" & psPersCod & "' And Substring(RH.cCtaCod,6,3) = '" & gCapCTS & "' And nPrdEstado Not In (" & gCapEstCancelada & "," & gCapEstAnulada & ")"
        Else
            'MAVM 20120223 Cambios Fractal
            'sqlC = " Select RH.cCtaCod from RHCuentas RH" _
            '     & " Inner Join " & lsCadTemp & "Producto PR On RH.cCtaCod = PR.cCtaCod" _
            '     & " Where cPersCod = '" & psPersCod & "' And Substring(RH.cCtaCod,6,3) = '" & gCapAhorros & "' And nPrdEstado Not In (" & gCapEstCancelada & "," & gCapEstAnulada & ")"
                 
            sqlC = " Select cCtaCod from RHPlanillaDetCon " _
                 & " Where cRHCod = '" & psPersCod & "' And cPlanillaCod = '" & psPlanillaTpo & "' And cRRHHPeriodo = '" & psPlanillaPeriodo & "'"
            '***
            
        End If
    Else
        If psPlanillaTpo = gsRHPlanillaCTS Then
            sqlC = " Select RH.cCtaCod from RHCuentas RH Where RH.cPersCod = '" & psPersCod & "' And Substring(RH.cCtaCod,3,3) = '" & gCapCTS & "'"
        Else
            sqlC = " Select RH.cCtaCod from RHCuentas RH Where RH.cPersCod = '" & psPersCod & "' And Substring(RH.cCtaCod,3,3) = '" & gCapAhorros & "'"
        End If
    End If
    
    If oCon.AbreConexion Then
        Set rsC = oCon.CargaRecordSet(sqlC)
        If rsC.EOF And rsC.BOF Then
            GetCuentaRRHH = ""
        Else
            If Not lbBitCentral Then
                If oConexionRem.AbreConexion Then 'Remota(Left(rsC!cCtaCod, 2)) Then
                    If Mid(rsC!cCtaCod, 3, 3) = Producto.gCapAhorros Then
                        sql = "Select cCodCta From AhorroC Where cCodCta = '" & rsC!cCtaCod & "' And cEstCtaAC Not In ('C','U')"
                    ElseIf Mid(rsC!cCtaCod, 3, 3) = Producto.gCapCTS Then
                        sql = "Select cCodCta From CTS Where cCodCta = '" & rsC!cCtaCod & "' And cEstCtaCTS Not In ('C','U')"
                    Else
                        GetCuentaRRHH = ""
                        oConexionRem.CierraConexion
                        Exit Function
                    End If
                    Set rsValida = oConexionRem.CargaRecordSet(sql)
                    
                    If rsValida.EOF And rsValida.BOF Then
                        GetCuentaRRHH = ""
                    Else
                        GetCuentaRRHH = rsC!cCtaCod
                    End If
                Else
                    GetCuentaRRHH = ""
                End If
            Else
                GetCuentaRRHH = rsC!cCtaCod
            End If
        End If
    End If
        
    rsC.Close
    Set rsC = Nothing
    Set oCon = Nothing
    Exit Function
GetCargoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetCargo Method")
End Function

Public Function GetSueldo(psPersCod As String) As Currency
    On Error GoTo GetCargoErr
    Dim sqlC As String
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta

    sqlC = " Select Top 1 nRHSueldoMonto Monto From RHSueldo where cPersCod = '" & psPersCod & "' Order By dRHSueldoFecha Desc"
        
    If oCon.AbreConexion Then
        Set rsC = oCon.CargaRecordSet(sqlC)
        If rsC.EOF And rsC.BOF Then
            GetSueldo = "0"
        Else
            GetSueldo = rsC!Monto
        End If
    End If
        
    rsC.Close
    Set rsC = Nothing
    Set oCon = Nothing
    Exit Function
GetCargoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetCargo Method")
End Function

Public Function GetFecVac(cPersCod As String, pdFecIni As String, pdFecFin As String) As Boolean
    Dim slqD As String
    Dim rsD As ADODB.Recordset
    Set rsD = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    slqD = "Select isnull(dEjecutadoInicio,'')dEjecutadoInicio,isnull(dEjecutadoFin,'')dEjecutadoFin From RHPeriodosNoLaborados Where cPersCod = '" & cPersCod & "' And nRHPeriodoTpo = '" & RHEstadosRRHH.gRHEstadosRRHHVACACIONES & "' Order By dEjecutadoInicio Desc"
    
    If oCon.AbreConexion Then
        Set rsD = oCon.CargaRecordSet(slqD)
        If rsD.EOF And rsD.BOF Then
            GetFecVac = False
        Else
            pdFecIni = rsD!dEjecutadoInicio
            pdFecFin = rsD!dEjecutadoFin
            GetFecVac = True
        End If
    End If
    
    rsD.Close
    Set rsD = Nothing
    Set oCon = Nothing
End Function

Public Function EsEmpleado(psRHPersCod As String) As Boolean
    On Error GoTo GetCargoErr
    Dim sqlC As String
    Dim oCon As DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oCon = New DConecta
    
    sqlC = " Select cPersCod From RRHH where cPersCod = '" & psRHPersCod & "'"
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sqlC)
        If rs.BOF And rs.EOF Then
            EsEmpleado = False
        Else
            EsEmpleado = True
        End If
        
        oCon.CierraConexion
    End If
    
    Set rs = Nothing
    Set oCon = Nothing
    Exit Function
GetCargoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetCargo Method")
End Function

Public Function GeEmpleado(psRHPersCod As String) As Boolean
    On Error GoTo GetCargoErr
    Dim sqlC As String
    Dim oCon As DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oCon = New DConecta
    
    sqlC = " Select cPersCod From RRHH where cPersCod = '" & psRHPersCod & "'"
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sqlC)
        If rs.BOF And rs.EOF Then
            GeEmpleado = False
        Else
            GeEmpleado = True
        End If
        
        oCon.CierraConexion
    End If
    
    Set rs = Nothing
    Set oCon = Nothing
    Exit Function
GetCargoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetCargo Method")
End Function


'Agrega un nuevo cargo o nivel
'##ModelId=3A947AB400AA
Public Function AgregaRRHHLote(prRRHH As ADODB.Recordset, psAreaCod As String, psCargo As String, psRHTpoContrato As String, psUltimaActualizacion As String, psFecha As String, psFormatoFecha As String, Optional pbProcesoSeleccion As Boolean = True) As Boolean
    On Error GoTo AgregaRRHHErr
    Dim sqlC As String
    Dim sqlV As String
    Dim rsV As ADODB.Recordset
    Dim oCon As DConecta
    Dim lsRHCodigo As String
    Set oCon = New DConecta
    Set rsV = New ADODB.Recordset
    Dim oContrato As DActualizaDatosContrato
    Set oContrato = New DActualizaDatosContrato
    Dim lsCodContrato As String
    
    '*** PEAC 20100625
    Dim rsCtrlAsist As ADODB.Recordset
    Set rsCtrlAsist = New ADODB.Recordset
    Dim lcPersCod As String, lcarea As String, lcCargo As String, lcOficina As String, sSqlCtrlAsist As String, sqlAsist As String
    '*** FIN PEAC

    If oCon.AbreConexion Then
        oCon.BeginTrans
        While Not prRRHH.EOF
            If pbProcesoSeleccion Then
                psFecha = Format(prRRHH.Fields(9), psFormatoFecha)
                If prRRHH.Fields(8) = "1" Then
                    sqlV = "Select cPersCod From RRHH Where cPersCod = '" & prRRHH.Fields(0) & "'"
                    Set rsV = oCon.CargaRecordSet(sqlV)
                    
                    lsRHCodigo = GetUltCorr(psRHTpoContrato, prRRHH.Fields(0), lsRHCodigo)
                    lsCodContrato = oContrato.GeCodContrato(prRRHH.Fields(0))
                    
                    If rsV.EOF And rsV.BOF Then
                        'ALPA 20081128*************************************************************************
                        'sqlC = " Insert RRHH (cPersCod,cRHCod,cAreaCod,cAreaCodActual,cAgenciaAsig,cAgenciaActual,dIngreso,cUltimaActualizacion,nRHEstado)" _
                        '      & " Values('" & prRRHH.Fields(0) & "','" & lsRHCodigo & "','" & psAreaCod & "','" & psAreaCod & "','" & prRRHH.Fields(14) & "','" & prRRHH.Fields(14) & "','" & Format(prRRHH.Fields(9), psFormatoFecha) & "','" & psUltimaActualizacion & "'," & Str(201) & ")"
                        sqlC = " Insert RRHH (cPersCod,cRHCod,cAreaCod,cAreaCodActual,cAgenciaAsig,cAgenciaActual,dIngreso,cUltimaActualizacion,nRHEstado,bMuestraCargo)"
                        sqlC = sqlC & " Values('" & prRRHH.Fields(0) & "','" & lsRHCodigo & "','" & psAreaCod & "','" & psAreaCod & "','" & prRRHH.Fields(14) & "','" & prRRHH.Fields(14) & "','" & Format(prRRHH.Fields(9), psFormatoFecha) & "','" & psUltimaActualizacion & "'," & Str(201) & ",1)"
                        '***************************************************************************************
                    Else
                        'ALPA 20081128*************************************************************************
                        'sqlC = " Update RRHH" _
                        '     & " Set cRHCod = '" & lsRHCodigo & "',cAreaCodActual  = '" & psAreaCod & "',cAreaCod = '" & psAreaCod & "',cAgenciaAsig = '" & prRRHH.Fields(14) & "', cAgenciaActual = '" & prRRHH.Fields(14) & "', cUltimaActualizacion = '" & psUltimaActualizacion & "', nRHEstado = " & Str(201) & " " _
                        '     & " Where cPersCod = '" & prRRHH.Fields(0) & "'"
                        sqlC = " Update RRHH" _
                             & " Set bMuestraCargo=1,cRHCod = '" & lsRHCodigo & "',cAreaCodActual  = '" & psAreaCod & "',cAreaCod = '" & psAreaCod & "',cAgenciaAsig = '" & prRRHH.Fields(14) & "', cAgenciaActual = '" & prRRHH.Fields(14) & "', cUltimaActualizacion = '" & psUltimaActualizacion & "', nRHEstado = " & Str(201) & " " _
                             & " Where cPersCod = '" & prRRHH.Fields(0) & "'"
                        '***************************************************************************************
                    End If
                    
                    rsV.Close
                    oCon.Ejecutar sqlC
                    If psRHTpoContrato = RHContratoTipo.RHContratoTipoIndeterminado Or psRHTpoContrato = RHContratoTipo.RHContratoTipoFijo Then
                        sqlV = "Select cPersCod From RHEmpleado Where cPersCod = '" & prRRHH.Fields(0) & "'"
                        Set rsV = oCon.CargaRecordSet(sqlV)
                        
                        If rsV.EOF And rsV.BOF Then
                            If psRHTpoContrato = RHContratoTipo.RHContratoTipoIndeterminado Then
                                sqlV = " Insert RHEmpleado(cPersCod, cRHEmpNivRem, cRHEmplAMPCod, cRHEmplAFPPersCod, nRHEmplVacacionesPend, dEstabilidad, bRHTipoFondo, cUltimaActualizacion)" _
                                     & " Values('" & prRRHH.Fields(0) & "','" & Left(psCargo, 3) & "','" & prRRHH.Fields(11) & "'," & IIf(prRRHH.Fields(19) = "", "Null", "'" & Right(prRRHH.Fields(19), 13) & "'") & " ,0,'" & Format(CDate(prRRHH.Fields(9)), psFormatoFecha) & "'," & prRRHH.Fields(16) & ",'" & psUltimaActualizacion & "')"
                            Else
                                sqlV = " Insert RHEmpleado(cPersCod, cRHEmpNivRem, cRHEmplAMPCod, cRHEmplAFPPersCod, nRHEmplVacacionesPend, bRHTipoFondo, cUltimaActualizacion)" _
                                     & " Values('" & prRRHH.Fields(0) & "','" & Left(psCargo, 3) & "','" & prRRHH.Fields(11) & "'," & IIf(prRRHH.Fields(19) = "", "Null", "'" & Right(prRRHH.Fields(19), 13) & "'") & " ,0," & prRRHH.Fields(16) & ",'" & psUltimaActualizacion & "')"
                            End If
                        Else
                            If psRHTpoContrato = RHContratoTipo.RHContratoTipoIndeterminado Then
                                sqlV = " Update RHEmpleado" _
                                     & " Set cRHEmpNivRem = '" & Left(psCargo, 3) & "', cRHEmplAMPCod '" & prRRHH.Fields(11) & "', cRHEmplAFPPersCod = " & IIf(prRRHH.Fields(19) = "", "Null", "'" & Right(prRRHH.Fields(19), 13) & "'") & ", dEstabilidad = '" & Format(CDate(prRRHH.Fields(9)), psFormatoFecha) & "', bRHTipoFondo = " & prRRHH.Fields(16) & " , cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
                                     & " Where cPersCod = '" & prRRHH.Fields(0) & "'"
                            Else
                                sqlV = " Update RHEmpleado" _
                                     & " Set cRHEmpNivRem = '" & Left(psCargo, 3) & "', cRHEmplAMPCod = '" & prRRHH.Fields(11) & "', cRHEmplAFPPersCod = " & IIf(prRRHH.Fields(19) = "", "Null", "'" & Right(prRRHH.Fields(19), 13) & "'") & ", bRHTipoFondo = " & prRRHH.Fields(16) & ", cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
                                     & " Where cPersCod = '" & prRRHH.Fields(0) & "'"
                            End If
                        End If
                        oCon.Ejecutar sqlV
                                
                        sqlV = " Insert RHsistemapensiones(cPersCod, dFecha, bRHTipoFondo, cRHAFPPersCod, cComentario, cUltimaActualizacion)" _
                             & " Values('" & prRRHH.Fields(0) & "','" & psFecha & "'," & prRRHH.Fields(16) & "," & IIf(prRRHH.Fields(19) = "", "Null", "'" & Right(prRRHH.Fields(19), 13) & "'") & ",'" & prRRHH.Fields(20) & "','" & psUltimaActualizacion & "')"
                        oCon.Ejecutar sqlV
                        
                        sqlV = " Insert RHAsistenciaMedica(cPersCod, dFecha, cRHAsistMedPrivCod, cComentario, cUltimaActualizacion)" _
                             & " Values('" & prRRHH.Fields(0) & "','" & psFecha & "','" & prRRHH.Fields(11) & "', '" & prRRHH.Fields(20) & "','" & psUltimaActualizacion & "')"
                             
                        oCon.Ejecutar sqlV
                    End If
                    
                    sqlV = " Insert RHCargos (cPersCod,dRHCargoFecha,cRHCargoCod,cRHAreaCod,cRHAgenciaCod,cRHCargoCodOficial,cRHAreaCodOficial,cRHAgenciaCodOficial,cRHCargoComentario,cUltimaActualizacion)" _
                         & " Values('" & prRRHH.Fields(0) & "','" & psFecha & "','" & psCargo & "','" & psAreaCod & "','" & prRRHH.Fields(14) & "','" & psCargo & "','" & psAreaCod & "','" & prRRHH.Fields(14) & "','" & prRRHH.Fields(20) & "','" & psUltimaActualizacion & "')"
                    oCon.Ejecutar sqlV
                    
                    sqlV = " Insert RHContrato(cPersCod, cRHContratoNro, cUltimaActualizacion,nRHContratoTpo)" _
                         & " Values('" & prRRHH.Fields(0) & "','" & lsCodContrato & "','" & psUltimaActualizacion & "'," & psRHTpoContrato & ")"
                    oCon.Ejecutar sqlV
                    
                    sqlV = " Insert RHContratoDet(cPersCod, cRHContratoNro, dRHContratoFecha, dRHContratoInicio, dRHContratoFin, cRHContratoComentario, cUltimaActualizacion)" _
                         & " Values('" & prRRHH.Fields(0) & "','" & lsCodContrato & "','" & psFecha & "','" & Format(prRRHH.Fields(9), psFormatoFecha) & "','" & Format(prRRHH.Fields(10), psFormatoFecha) & "','" & prRRHH.Fields(20) & "','" & psUltimaActualizacion & "')"
                    oCon.Ejecutar sqlV
                    
                    'sqlV = " Insert RHAreaAgencia(cPersCod, dFecha, cAreaCod, cAgenciaAsig, cComentario, cUltimaActualizacion)" _
                         & " Values('" & prRRHH.Fields(0) & "','" & psFecha & "','" & psAreaCod & "','" & prRRHH.Fields(14) & "','" & prRRHH.Fields(20) & "','" & psUltimaActualizacion & "')"
                    'oCon.Ejecutar sqlV
                    
                    'sqlV = " Insert RHAgencia (cPersCod,dFecha,cAgenciaActual,cComentario,cUltimaActualizacion,cAreaCodActual)" _
                         & " Values('" & prRRHH.Fields(0) & "','" & psFecha & "','" & prRRHH.Fields(14) & "','" & prRRHH.Fields(20) & "','" & psUltimaActualizacion & "','" & psAreaCod & "')"
                    'oCon.Ejecutar sqlV
                    
                    sqlV = " Insert RHSueldo(cPersCod, dRHSueldoFecha, nRHSueldoMonto, cRHSueldoComentario, cUltimaActualizacion)" _
                         & " Values('" & prRRHH.Fields(0) & "','" & psFecha & "'," & CCur(prRRHH.Fields(13)) & ",'" & prRRHH.Fields(20) & "','" & psUltimaActualizacion & "')"
                    oCon.Ejecutar sqlV
                    
                    sqlV = " Insert RHEstado(cPersCod, dRHEstadoFecha, nRHEstado, cRHEstadoComentario, cUltimaActualizacion)" _
                         & " Values('" & prRRHH.Fields(0) & "','" & psFecha & "','" & Str(201) & "','" & prRRHH.Fields(20) & "','" & psUltimaActualizacion & "')"
                    oCon.Ejecutar sqlV
                End If
            Else
                psFecha = Format(prRRHH.Fields(2), psFormatoFecha)
                
                sqlV = "Select cPersCod From RRHH Where cPersCod = '" & prRRHH.Fields(0) & "'"
                Set rsV = oCon.CargaRecordSet(sqlV)
                
                lsRHCodigo = GetUltCorr(psRHTpoContrato, prRRHH.Fields(0), lsRHCodigo)
                lsCodContrato = oContrato.GeCodContrato(prRRHH.Fields(0))
                
                If rsV.EOF And rsV.BOF Then
'                    sqlC = " Insert RRHH (cPersCod,cRHCod,cAreaCod,cAreaCodActual,cAgenciaAsig,cAgenciaActual,dIngreso,cUltimaActualizacion,nRHEstado,bAgregaPlanilla)" _
'                         & " Values('" & prRRHH.Fields(0) & "','" & lsRHCodigo & "','" & psAreaCod & "','" & psAreaCod & "','" & prRRHH.Fields(7) & "','" & prRRHH.Fields(7) & "','" & Format(prRRHH.Fields(2), psFormatoFecha) & "','" & psUltimaActualizacion & "'," & Str(201) & " , '' )"
                     sqlC = " Insert RRHH (cPersCod,cRHCod,cAreaCod,cAreaCodActual,cAgenciaAsig,cAgenciaActual,dIngreso,cUltimaActualizacion,nRHEstado,bAgregaPlanilla,bMuestraCargo)" _
                         & " Values('" & prRRHH.Fields(0) & "','" & lsRHCodigo & "','" & psAreaCod & "','" & psAreaCod & "','" & prRRHH.Fields(7) & "','" & prRRHH.Fields(7) & "','" & Format(prRRHH.Fields(2), psFormatoFecha) & "','" & psUltimaActualizacion & "'," & Str(201) & " , '',1)"
                Else
'                    sqlC = " Update RRHH Set cRHCod = '" & lsRHCodigo & "',cAreaCodActual = '" & psAreaCod & "', cAreaCod = '" & psAreaCod & "',cAgenciaAsig = '" & prRRHH.Fields(7) & "', cAgenciaActual = '" & prRRHH.Fields(7) & "', cUltimaActualizacion = '" & psUltimaActualizacion & "', nRHEstado = " & Str(201) & " " _
'                         & " Where cPersCod = '" & prRRHH.Fields(0) & "'"
                    sqlC = " Update RRHH Set bMuestraCargo=1,cRHCod = '" & lsRHCodigo & "',cAreaCodActual = '" & psAreaCod & "', cAreaCod = '" & psAreaCod & "',cAgenciaAsig = '" & prRRHH.Fields(7) & "', cAgenciaActual = '" & prRRHH.Fields(7) & "', cUltimaActualizacion = '" & psUltimaActualizacion & "', nRHEstado = " & Str(201) & " " _
                         & " Where cPersCod = '" & prRRHH.Fields(0) & "'"
                
                End If
                
                oCon.Ejecutar sqlC

                '***PEAC 20100625 - BUSCA EMPLEADO EN CONTROL DE ASISTENCIA *******************
                    'sSqlCtrlAsist = " exec stp_sel_BuscaEmpleadoCtrlAsistencia '" & psPersCod & "'"
                    sSqlCtrlAsist = " exec stp_sel_BuscaEmpleadoCtrlAsistencia '" & prRRHH.Fields(0) & "'"

                    Set rsCtrlAsist = oCon.CargaRecordSet(sSqlCtrlAsist)
                    If rsCtrlAsist.EOF And rsCtrlAsist.BOF Then 'si no existe
                        sqlAsist = " exec stp_ins_EmpleadoCtrlAsist '" & prRRHH.Fields(0) & "'," & Val(psAreaCod) & "," & Val(psCargo) & "," & Val(prRRHH.Fields(7)) & "," & Val(psRHTpoContrato) & ",'" & Format(prRRHH.Fields(2), "yyyymmdd") & "'"
                    Else
                        sqlAsist = " exec stp_upd_EmpleadoCeseCtrlAsist '" & prRRHH.Fields(0) & "'," & 1 & ",'" & Format(psFecha, "yyyymmdd") & "'"
                    End If
                    oCon.Ejecutar sqlAsist
                    rsCtrlAsist.Close
                '*** FIN PEAC *****************************************

                If psRHTpoContrato = RHContratoTipo.RHContratoTipoIndeterminado Or psRHTpoContrato = RHContratoTipo.RHContratoTipoFijo Or psRHTpoContrato = RHContratoTipo.RHContratoTipoDirector Then
                    rsV.Close
                    sqlV = "Select cPersCod From RHEmpleado Where cPersCod = '" & prRRHH.Fields(0) & "'"
                    Set rsV = oCon.CargaRecordSet(sqlV)
                    
                    If rsV.EOF And rsV.BOF Then
                        If psRHTpoContrato = RHContratoTipo.RHContratoTipoIndeterminado Or psRHTpoContrato = RHContratoTipo.RHContratoTipoDirector Then
                            sqlV = " Insert RHEmpleado(cPersCod, cRHEmpNivRem, cRHEmplAMPCod, cRHEmplAFPPersCod, nRHEmplVacacionesPend, dEstabilidad, bRHTipoFondo, cUltimaActualizacion)" _
                                 & " Values('" & prRRHH.Fields(0) & "','" & Left(psCargo, 3) & "','" & prRRHH.Fields(4) & "'," & IIf(prRRHH.Fields(12) = "", "Null", "'" & Right(prRRHH.Fields(12), 13) & "'") & " ,0,'" & Format(CDate(prRRHH.Fields(2)), psFormatoFecha) & "'," & IIf(prRRHH.Fields(9) = "", "NULL", prRRHH.Fields(9)) & ",'" & psUltimaActualizacion & "')"
                        Else
                            sqlV = " Insert RHEmpleado(cPersCod, cRHEmpNivRem, cRHEmplAMPCod, cRHEmplAFPPersCod, nRHEmplVacacionesPend, bRHTipoFondo, cUltimaActualizacion)" _
                                 & " Values('" & prRRHH.Fields(0) & "','" & Left(psCargo, 3) & "','" & prRRHH.Fields(4) & "'," & IIf(prRRHH.Fields(12) = "", "Null", "'" & Right(prRRHH.Fields(12), 13) & "'") & " ,0," & prRRHH.Fields(9) & ",'" & psUltimaActualizacion & "')"
                        End If
                    Else
                        If psRHTpoContrato = RHContratoTipo.RHContratoTipoIndeterminado Or psRHTpoContrato = RHContratoTipo.RHContratoTipoDirector Then
                            sqlV = " Update RHEmpleado" _
                                 & " Set cRHEmpNivRem = '" & Left(psCargo, 3) & "', cRHEmplAMPCod = '" & prRRHH.Fields(4) & "', cRHEmplAFPPersCod = " & IIf(prRRHH.Fields(12) = "", "Null", "'" & Right(prRRHH.Fields(12), 13) & "'") & ", dEstabilidad = '" & Format(CDate(prRRHH.Fields(2)), psFormatoFecha) & "', bRHTipoFondo = " & IIf(prRRHH.Fields(9) = "", "NULL", prRRHH.Fields(9)) & ", cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
                                 & " Where cPersCod = '" & prRRHH.Fields(0) & "'"
                        Else
                            sqlV = " Update RHEmpleado" _
                                 & " Set cRHEmpNivRem = '" & Left(psCargo, 3) & "', cRHEmplAMPCod = '" & prRRHH.Fields(4) & "', cRHEmplAFPPersCod = " & IIf(prRRHH.Fields(12) = "", "Null", "'" & Right(prRRHH.Fields(12), 13) & "'") & ", bRHTipoFondo = " & prRRHH.Fields(9) & ", cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
                                 & " Where cPersCod = '" & prRRHH.Fields(0) & "'"
                        End If
                    End If
                    oCon.Ejecutar sqlV
                    
                    If prRRHH.Fields(9) <> "" Then
                        sqlV = " Insert RHsistemapensiones(cPersCod, dFecha, bRHTipoFondo, cRHAFPPersCod, cComentario, cUltimaActualizacion)" _
                             & " Values('" & prRRHH.Fields(0) & "','" & psFecha & "'," & prRRHH.Fields(9) & "," & IIf(prRRHH.Fields(12) = "", "Null", "'" & Right(prRRHH.Fields(12), 13) & "'") & ",'" & prRRHH.Fields(13) & "','" & psUltimaActualizacion & "')"
                        oCon.Ejecutar sqlV
                    End If
                    
                    sqlV = " Insert RHAsistenciaMedica(cPersCod, dFecha, cRHAsistMedPrivCod, cComentario, cUltimaActualizacion)" _
                         & " Values('" & prRRHH.Fields(0) & "','" & psFecha & "','" & prRRHH.Fields(4) & "', '" & prRRHH.Fields(13) & "','" & psUltimaActualizacion & "')"
                         
                    oCon.Ejecutar sqlV
                End If
                
                sqlV = " Insert RHCargos (cPersCod,dRHCargoFecha,cRHCargoCod,cRHAreaCod,cRHAgenciaCod,cRHCargoCodOficial,cRHAreaCodOficial,cRHAgenciaCodOficial,cRHCargoComentario,cUltimaActualizacion)" _
                     & " Values('" & prRRHH.Fields(0) & "','" & psFecha & "','" & psCargo & "','" & psAreaCod & "','" & prRRHH.Fields(7) & "','" & psCargo & "','" & psAreaCod & "','" & prRRHH.Fields(7) & "','" & prRRHH.Fields(13) & "','" & psUltimaActualizacion & "')"
                oCon.Ejecutar sqlV
                
                sqlV = " Insert RHContrato(cPersCod, cRHContratoNro, cUltimaActualizacion,nRHContratoTpo)" _
                     & " Values('" & prRRHH.Fields(0) & "','" & lsCodContrato & "','" & psUltimaActualizacion & "'," & psRHTpoContrato & ")"
                oCon.Ejecutar sqlV
                
                sqlV = " Insert RHContratoDet(cPersCod, cRHContratoNro, dRHContratoFecha, dRHContratoInicio, dRHContratoFin, cRHContratoComentario,cUltimaActualizacion)" _
                     & " Values('" & prRRHH.Fields(0) & "','" & lsCodContrato & "','" & psFecha & "','" & Format(prRRHH.Fields(2), psFormatoFecha) & "','" & Format(prRRHH.Fields(3), psFormatoFecha) & "','" & prRRHH.Fields(13) & "','" & psUltimaActualizacion & "')"
                oCon.Ejecutar sqlV
                
                'sqlV = " Insert RHAreaAgencia(cPersCod, dFecha, cAreaCod, cAgenciaAsig, cComentario, cUltimaActualizacion)" _
                     & " Values('" & prRRHH.Fields(0) & "','" & psFecha & "','" & psAreaCod & "','" & prRRHH.Fields(7) & "','" & prRRHH.Fields(13) & "','" & psUltimaActualizacion & "')"
                'oCon.Ejecutar sqlV
                
                'sqlV = " Insert RHAgencia (cPersCod,dFecha,cAgenciaActual,cComentario,cUltimaActualizacion,cAreaCodActual)" _
                     & " Values('" & prRRHH.Fields(0) & "','" & psFecha & "','" & prRRHH.Fields(7) & "','" & prRRHH.Fields(13) & "','" & psUltimaActualizacion & "','" & psAreaCod & "')"
                'oCon.Ejecutar sqlV
                
                sqlV = " Insert RHSueldo(cPersCod, dRHSueldoFecha, nRHSueldoMonto, cRHSueldoComentario, cUltimaActualizacion)" _
                     & " Values('" & prRRHH.Fields(0) & "','" & psFecha & "'," & CCur(prRRHH.Fields(6)) & ",'" & prRRHH.Fields(13) & "','" & psUltimaActualizacion & "')"
                oCon.Ejecutar sqlV
                
                
                '************
                sqlV = " Insert RHEstado(cPersCod, dRHEstadoFecha, nRHEstado, cRHEstadoComentario, cUltimaActualizacion)" _
                     & " Values('" & prRRHH.Fields(0) & "','" & psFecha & "','" & Str(201) & "','" & prRRHH.Fields(13) & "','" & psUltimaActualizacion & "')"
                     
'                sqlV = " Insert RHEstado(cPersCod, dRHEstadoFecha, nRHEstado, cRHEstadoComentario, cUltimaActualizacion,nRHCodMotivoIngCese)" _
'                     & " Values('" & prRRHH.Fields(0) & "','" & psFecha & "','" & Str(201) & "','" & prRRHH.Fields(15) & "','" & psUltimaActualizacion & "'," & prRRHH.Fields(13) & ")"
                     
                     
                'se agrego prRRHH.Fields(13) PEAC 20101022
                oCon.Ejecutar sqlV
            End If
            prRRHH.MoveNext
        Wend
        oCon.CommitTrans
        oCon.CierraConexion
    End If
            
    Set oCon = Nothing
    Exit Function
AgregaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:AgregaRRHH Method")
End Function

'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
Public Function GetRRHHFileCargo(psRRHHCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo GetRRHHGeneralidadesErr
    
    sqlC = " Select cRHCargoDescripcion, cRHCargoComentario, RHC.cUltimaActualizacion from rhCargos RHC" _
         & " Inner Join RHCargosTabla RHCT On RHC.cRHCargoCodAsig = RHCT.cRHCargoCod" _
         & " where cPersCod = '" & psRRHHCod & "' order by RHC.cUltimaActualizacion Desc"
    
    If oCon.AbreConexion Then
        Set GetRRHHFileCargo = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHGeneralidadesErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function

'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
Public Function GetRRHHFileAsig(psRRHHCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo GetRRHHGeneralidadesErr

    sqlC = " Select Distinct ARE.cAreaCod + RHCT.cRHCargoCod, RHCT.cRHCargoDescripcion, ARE.cAreaDescripcion, AGE.cAgeCod, AGE.cAgeDescripcion, ARE1.cAreaCod + RHCT1.cRHCargoCod, RHCT1.cRHCargoDescripcion, ARE1.cAreaDescripcion, AGE1.cAgeCod, AGE1.cAgeDescripcion,RHA.cRHCargoComentario, RHA.cUltimaActualizacion from RHCargos RHA" _
         & " Inner Join Agencias AGE on RHA.cRHAgenciaCod = AGE.cAgeCod Inner Join Agencias AGE1 on RHA.cRHAgenciaCodOficial = AGE1.cAgeCod" _
         & " Inner Join Areas ARE On RHA.cRHAreaCod = ARE.cAreaCod Inner Join Areas ARE1 On RHA.cRHAreaCodOficial = ARE1.cAreaCod Inner Join RHCargosTabla RHCT On RHA.cRHCargoCod = RHCT.cRHCargoCod Inner Join RHCargosTabla RHCT1 On RHA.cRHCargoCodOficial = RHCT1.cRHCargoCod" _
         & " where cPersCod = '" & psRRHHCod & "' order by RHA.cUltimaActualizacion Desc"
    
    If oCon.AbreConexion Then
        Set GetRRHHFileAsig = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHGeneralidadesErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function

'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
Public Function GetRRHHFileAct(psRRHHCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo GetRRHHGeneralidadesErr
    
    sqlC = " Select cAreaDescripcion, cAgeDescripcion , RHA.cRHCargoComentario, RHA.cUltimaActualizacion from RHCargos RHA" _
         & " INner Join Agencias AGE on cAgeCodAct = AGE.cAgeCod" _
         & " Inner join Areas ARE On RHA.cAreaCodAct = ARE.cAreaCod" _
         & " where cPersCod = '" & psRRHHCod & "' order by RHA.cUltimaActualizacion Desc"
    
    If oCon.AbreConexion Then
        Set GetRRHHFileAct = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHGeneralidadesErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function

'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
Public Function GetRRHHFileSueldo(psRRHHCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo GetRRHHGeneralidadesErr
    
    sqlC = " Select convert(decimal(8,2),nRHSueldoMonto),cRHSueldoComentario,cUltimaActualizacion from rhsueldo" _
         & " where cPersCod = '" & psRRHHCod & "' order by cUltimaActualizacion Desc"
    
    If oCon.AbreConexion Then
        Set GetRRHHFileSueldo = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHGeneralidadesErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function

'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
Public Function GetRRHHFileEstado(psRRHHCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo GetRRHHGeneralidadesErr
    
    sqlC = " Select CO.cConsDescripcion,cRHEstadoComentario,cUltimaActualizacion from rhEstado RHE" _
         & " Inner Join Constante CO On CO.nConsValor = RHE.nRHEstado And nConsCod = '" & gRHEstado & "'" _
         & " where cPersCod = '" & psRRHHCod & "' order by RHE.cUltimaActualizacion Desc"
    
    If oCon.AbreConexion Then
        Set GetRRHHFileEstado = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHGeneralidadesErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function

'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
Public Function GetRRHHEstado(psRRHHCod As String) As String
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    On Error GoTo GetRRHHGeneralidadesErr
    
    sqlC = " Select Top 1 CO.cConsDescripcion  from RRHH RH" _
         & " Inner Join Constante CO On CO.nConsValor = RH.nRHEstado And nConsCod = '" & gRHEstado & "'" _
         & " where cPersCod = '" & psRRHHCod & "'"
    
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sqlC)
        If rs.EOF And rs.BOF Then
            GetRRHHEstado = ""
        Else
            GetRRHHEstado = rs.Fields(0)
        End If
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHGeneralidadesErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function


'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
Public Function GetRRHHFilePensiones(psRRHHCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo GetRRHHGeneralidadesErr
    
    sqlC = " Select CO.nConsValor,CO.cConsDescripcion, AFP.cRHAFPAbreviatura, PE.cPersNombre + '-' + PE.cPersCod ,RHP.cComentario,RHP.cUltimaActualizacion from rhsistemaPensiones RHP" _
         & " Inner Join Constante CO On CO.nConsValor = RHP.bRHTipoFondo And nConsCod = '" & gRHConceptoVariablesUsuario & "'" _
         & " Left Join Persona PE On RHP.cRHAFPPersCod = PE.cPersCod" _
         & " Left Join RHAFP AFP On AFP.cRHAFPPersCod = PE.cPersCod" _
         & " where RHP.cPersCod = '" & psRRHHCod & "' order by RHP.cUltimaActualizacion Desc"
    
    If oCon.AbreConexion Then
        Set GetRRHHFilePensiones = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHGeneralidadesErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function

'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
Public Function GetRRHHFileAMP(psRRHHCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo GetRRHHGeneralidadesErr
    
    sqlC = " Select RHAT.cRHAsistMedPrivCod, RHAT.cRHAsistMedPrivDescripcion, RHA.cComentario,RHA.cUltimaActualizacion from rhAsistenciaMedica RHA" _
         & " Inner Join rhasistmedprivtabla RHAT On RHA.cRHAsistMedPrivCod = RHAT.cRHAsistMedPrivCod" _
         & " where RHA.cPersCod = '" & psRRHHCod & "'" _
         & " order by RHA.cUltimaActualizacion Desc"
    
    If oCon.AbreConexion Then
        Set GetRRHHFileAMP = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHGeneralidadesErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function

'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
Public Function GetRRHHFileComentario(psRRHHCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo GetRRHHGeneralidadesErr
    
    sqlC = " Select cRHComentario, cUltimaActualizacion from rhcomentario" _
         & " where cPersCod = '" & psRRHHCod & "' order by cUltimaActualizacion Desc"
    
    If oCon.AbreConexion Then
        Set GetRRHHFileComentario = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHGeneralidadesErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function

'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
Public Function GetRRHHFileDetalleContrato(psRRHHCod As String, Optional psContratoCod As String = "") As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo GetRRHHGeneralidadesErr
    
    If psContratoCod = "" Then
        sqlC = " Select nRHContratoTpo,cConsDescripcion ,Convert(varchar(10),dRHContratoInicio,103) FInicio, Convert(varchar(10),dRHContratoFin,103) FFin ,cRHContratoComentario Comentario, A.cUltimaActualizacion From RHContratoDet A iNNER JOIN RhCOntrato b on a.cperscod = b.cPerscod And A.cRHContratoNro = B.cRHContratoNro Inner Join Constante CO On CO.nConsCod = " & gRHTipoContrato & " And CO.nConsValor = B.nRHContratoTpo " _
             & " Where A.cPersCod = '" & psRRHHCod & "' order by A.cUltimaActualizacion Desc"
    Else
        sqlC = " Select Convert(varchar(10),dRHContratoInicio,103) FInicio, Convert(varchar(10),dRHContratoFin,103) FFin ,cRHContratoComentario Comentario, cUltimaActualizacion From RHContratoDet" _
             & " Where cPersCod = '" & psRRHHCod & "' And cRHContratoNro = '" & psContratoCod & "' order by cUltimaActualizacion Desc"
    End If
    
    If oCon.AbreConexion Then
        Set GetRRHHFileDetalleContrato = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHGeneralidadesErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function

'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
Public Function AgregaRRHHComentario(psPersCod As String, psRHComentarioFecha As String, psRHComentario As String, psUltimaActualizacion As String) As Boolean
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo GetRRHHGeneralidadesErr
    
    sqlC = " Insert RHComentario (cPersCod,dRHComentarioFecha,cRHComentario,cUltimaActualizacion)" _
         & " Values('" & psPersCod & "','" & psRHComentarioFecha & "','" & psRHComentario & "','" & psUltimaActualizacion & "')"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHGeneralidadesErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function

'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
Public Function AgregaRRHHFondo(psPersCod As String, psFecha As String, psRHTipoFondo As String, psRHAFPPersCod As String, psComentario As String, psUltimaActualizacion As String) As Boolean
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo GetRRHHGeneralidadesErr
    
    sqlC = " Insert RHSistemaPensiones (cPersCod,dFecha,bRHTipoFondo,cRHAFPPersCod,cComentario,cUltimaActualizacion)" _
         & " Values('" & psPersCod & "','" & psFecha & "','" & psRHTipoFondo & "','" & psRHAFPPersCod & "','" & psComentario & "','" & psUltimaActualizacion & "')"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        
        sqlC = "Update rhempleado set cRHEmplAFPPersCod = '" & psRHAFPPersCod & "' where cperscod = '" & psPersCod & "'"
        oCon.Ejecutar sqlC
        
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHGeneralidadesErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function

Public Function AgregaRRHHAMP(psPersCod As String, psFecha As String, psRHAsistMedPrivCod As String, psComentario As String, psUltimaActualizacion As String) As Boolean
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo GetRRHHGeneralidadesErr
    
    sqlC = " Insert rhAsistenciamedica (cPersCod,dFecha,cRHAsistMedPrivCod,cComentario,cUltimaActualizacion)" _
         & " Values('" & psPersCod & "','" & psFecha & "','" & psRHAsistMedPrivCod & "','" & psComentario & "','" & psUltimaActualizacion & "')"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        sqlC = "Update rhempleado set cRHEmplAMPCod = '" & psRHAsistMedPrivCod & "' where cperscod = '" & psPersCod & "'"
        oCon.Ejecutar sqlC
        
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRRHHGeneralidadesErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function


'Devuelve a un RRHH en particular se le envia el codigo de persona o de recurso
'humano.
'##ModelId=3A947AE2006A
'lnTipoControl
'0 Ninguno
'1 Asistencia
'2 Vacaciones
'3 Permsisos
Public Function GetRRHHControl(psPersCod As String, Optional lnTipoControl As Integer = 0) As Integer
    Dim oCon As DConecta
    Dim sqlC As String
    Dim sqlAuxT As String
    Dim sqlAuxC As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set oCon = New DConecta
    
    On Error GoTo GetRRHHErr
    
    If lnTipoControl = 0 Then
        sqlAuxT = ""
        sqlAuxC = ", 1 Control"
    ElseIf lnTipoControl = 1 Then
        sqlAuxT = " Inner Join RHTipoContrato RHT ON RHT.cRHContratoCod = Left(RH.cRHCod,1)"
        sqlAuxC = ", bRHControlAsistencia Control"
    ElseIf lnTipoControl = 2 Then
        sqlAuxT = " Inner Join RHTipoContrato RHT ON RHT.cRHContratoCod = Left(RH.cRHCod,1)"
        sqlAuxC = ", bRHControlVacaciones Control"
    ElseIf lnTipoControl = 3 Then
        sqlAuxT = " Inner Join RHTipoContrato RHT ON RHT.cRHContratoCod = Left(RH.cRHCod,1)"
        sqlAuxC = ", bRHControlPermisos Control"
    End If
    
    sqlC = " Select RH.cPersCod as Codigo " & sqlAuxC & " " _
         & " From RRHH RH " & sqlAuxT _
         & " where RH.cPersCod = '" & psPersCod & "'"
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sqlC)
        GetRRHHControl = rs!Control
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRRHH Method")
End Function

'Select Left(cConsDescripcion,20) Tipo, cPersIDnro NroDoc from PersId PID
'Inner Join Constante CO On PID.cPersIDTpo  = CO.nConsValor
'where nConsCod = '1003' where cPersCod = ''
#If DebugMode Then
    '##ModelId=3AB902F00144
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

Public Function GetNiveldeAprovacion(psPersCod As String) As Integer
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = "Select IsNull(nRHGradoAprobacion,0) Grado from RHCargosTabla  where cRHCargoCod in (Select Top 1 cRHCargoCodOficial From RHCargos Where cPersCod = '" & psPersCod & "' Order by dRHCargoFecha Desc)"
    oCon.AbreConexion
    
    Set rs = oCon.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        GetNiveldeAprovacion = 0
    Else
        GetNiveldeAprovacion = rs.Fields(0)
    End If
    
    
End Function

Public Sub GetAreaAgenciaRRHH(ByVal psPersCod As String, ByRef psArea As String, ByRef psAge As String)
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    
    sql = "Select cAreaCod, cAgenciaAsig From RRHH Where cPersCod = '" & psPersCod & "'"
    oCon.AbreConexion
    
    Set rs = oCon.CargaRecordSet(sql)
    psArea = rs!cAreaCod
    psAge = rs!cAgenciaAsig
End Sub

Public Function GetFecIngEmp(psPersCod As String) As Date
    Dim sqlE As String
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlE = "Select dIngreso from RRHH where cPersCod = '" & psPersCod & "'"
     
    If oCon.AbreConexion Then
        Set rsE = oCon.CargaRecordSet(sqlE)
        If Not RSVacio(rsE) Then
            If Not IsNull(rsE!dIngreso) Then
                GetFecIngEmp = rsE!dIngreso
            End If
        End If
        oCon.CierraConexion
    End If
    
    rsE.Close
    Set rsE = Nothing
    Set oCon = Nothing
End Function

Public Function GetTpoContrato(psPersCod As String) As Long
    Dim sqlE As String
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlE = " Select nRHContratoTpo From RHContrato C Where C.cRHContratoNro = " _
        & " (Select Max(CA.cRHContratoNro) From RHContrato CA Where C.cPersCod = CA.cPersCod)" _
        & " And C.cPersCod = '" & psPersCod & "'"
    If oCon.AbreConexion Then
        Set rsE = oCon.CargaRecordSet(sqlE)
        If Not RSVacio(rsE) Then
            If Not IsNull(rsE.Fields(0)) Then
                GetTpoContrato = rsE.Fields(0)
            End If
        End If
        oCon.CierraConexion
    End If
    
    rsE.Close
    Set rsE = Nothing
    Set oCon = Nothing
End Function

Public Function GetFecCeseEmp(psPersCod As String) As String
    Dim sqlE As String
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlE = "Select dCese from RRHH where cPersCod = '" & psPersCod & "'"
     
    If oCon.AbreConexion Then
        Set rsE = oCon.CargaRecordSet(sqlE)
        If Not RSVacio(rsE) Then
            If Not IsNull(rsE!dCese) Then
                GetFecCeseEmp = rsE!dCese
            Else
                GetFecCeseEmp = ""
            End If
        End If
        oCon.CierraConexion
    End If
    
    rsE.Close
    Set rsE = Nothing
    Set oCon = Nothing
End Function

Public Function ValidaCtasEmpleado(prRS As ADODB.Recordset, pgbBitCentral As Boolean) As Boolean
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
 
    
    If prRS Is Nothing Then
        ValidaCtasEmpleado = True
        Exit Function
    End If
    
    prRS.MoveFirst
    
    If Not pgbBitCentral Then
        While Not prRS.EOF
            If Mid(prRS.Fields(0), 3, 3) = Producto.gCapAhorros Then
                sql = "Select cCodCta From AhorroC Where cCodCta = '" & prRS.Fields(0) & "' And cEstCtaAC Not In ('C','U')"
            Else
                sql = "Select cCodCta From CTS Where cCodCta = '" & prRS.Fields(0) & "' And cEstCtaCTS Not In ('C','U')"
            End If
            
            If oCon.AbreConexion Then 'Remota(Left(prRS.Fields(0), 2)) Then
                Set rs = oCon.CargaRecordSet(sql)
                If Not RSVacio(rs) Then
                    ValidaCtasEmpleado = True
                Else
                    ValidaCtasEmpleado = False
                    oCon.CierraConexion
                    Exit Function
                End If
                oCon.CierraConexion
            End If
            rs.Close
            prRS.MoveNext
        Wend
    Else
        oCon.AbreConexion
        While Not prRS.EOF
            sql = "Select PRD.cCtaCod From Producto PRD Inner Join Captaciones CAP On PRD.cCtaCod = CAP.cCtaCod Where PRD.cCtaCod = '" & prRS.Fields(0) & "' And PRD.nPrdEstado Not In ('" & CaptacEstado.gCapEstAnulada & "','" & CaptacEstado.gCapEstCancelada & "')"
            
            Set rs = oCon.CargaRecordSet(sql)
            If Not RSVacio(rs) Then
                ValidaCtasEmpleado = True
            Else
                ValidaCtasEmpleado = False
                oCon.CierraConexion
                Exit Function
            End If
            rs.Close
            prRS.MoveNext
        Wend
    End If

End Function


Public Function GetPlanillaPagada(psPersCod As String, psPlanillaCod As String, psPeriodo As String) As Boolean
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
        
    Dim oCon As DConecta
    Set oCon = New DConecta
        
    sql = " Select  distinct bPagado from rhplanilladetcon " _
        & " Where cRHcod = '" & psPersCod & "' And cPlanillaCod = '" & psPlanillaCod & "' and cRRHHPeriodo = '" & psPeriodo & "'"

    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sql)
        
        If rs.EOF And rs.BOF Then
            GetPlanillaPagada = False
        Else
            If rs.Fields(0) = 0 Then
                GetPlanillaPagada = False
            Else
                GetPlanillaPagada = True
            End If
            
        End If
        oCon.CierraConexion
    End If
    
    rs.Close
    Set rs = Nothing
    Set oCon = Nothing
    
End Function

Public Function SetPlanillaPagada(psPersCod As String, psPlanillaCod As String, psPeriodo As String) As Boolean
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
        
    sql = " Update rhplanilladetcon " _
        & " Set bPagado = 1" _
        & " Where cRHcod = '" & psPersCod & "' And cPlanillaCod = '" & psPlanillaCod & "' and cRRHHPeriodo = '" & psPeriodo & "'"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sql
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function


Public Function GetFecNac(psPersCod As String) As Date
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = " Select dPersNacCreac From Persona " _
        & " Where cPersCod = '" & psPersCod & "'"
        
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sql)
        GetFecNac = rs.Fields(0)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function

Public Function ValidaCta(psCtaCod As String, pgbBitCentral As Boolean) As Boolean
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
 
    
    If psCtaCod = "" Then
        ValidaCta = False
        Exit Function
    End If
    
    If Not pgbBitCentral Then
        If Mid(psCtaCod, 3, 3) = Producto.gCapAhorros Then
            sql = "Select cCodCta From AhorroC Where cCodCta = '" & psCtaCod & "' And cEstCtaAC Not In ('C','U')"
        Else
            sql = "Select cCodCta From CTS Where cCodCta = '" & psCtaCod & "' And cEstCtaCTS Not In ('C','U')"
        End If
        
        If oCon.AbreConexion Then 'Remota(Left(psCtaCod, 2)) Then
            Set rs = oCon.CargaRecordSet(sql)
            If Not RSVacio(rs) Then
                ValidaCta = True
            Else
                ValidaCta = False
                oCon.CierraConexion
                Exit Function
            End If
            oCon.CierraConexion
        End If
        rs.Close
    Else
        oCon.AbreConexion
        sql = " Select PRD.cCtaCod From Producto PRD " _
            & " Inner Join Captaciones CAP On PRD.cCtaCod = CAP.cCtaCod" _
            & " Where PRD.cCtaCod = '" & psCtaCod & "' And Left(PRD.nPrdEstado,2) Not In ('" & Left(CaptacEstado.gCapEstAnulada, 2) & "','" & Left(CaptacEstado.gCapEstCancelada, 2) & "')"
        Set rs = oCon.CargaRecordSet(sql)
        
        If Not RSVacio(rs) Then
            ValidaCta = True
        Else
            ValidaCta = False
            oCon.CierraConexion
            Exit Function
        End If
        rs.Close
    End If
End Function

Public Function GetRHTelefonos() As ADODB.Recordset
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    sql = " Select RH.cRHCod, PE.cPersNombre Nombre, (Select cAreaDescripcion from Areas A Where A.cAreaCod = RH.cAreaCodActual) Area, " _
        & " (Select cAgeDescripcion from Agencias A Where A.cAgeCod = RH.cAgenciaActual) Agencia," _
        & " (select cRHCargoDescripcion from rhcargostabla where cRHCargoCod = dbo.GetDRHCargos1(Rh.cperscod,getdate()))  Cargo, cAnexoInterno,  cCelular, cRPM " _
        & " From RRHH RH" _
        & " Inner Join Persona PE On RH.cPersCod = PE.cPersCod" _
        & " Where RH.nRHEstado < 700 order by RH.cAgenciaActual, Area, Nombre"
    Set GetRHTelefonos = oCon.CargaRecordSet(sql)

    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Sub SetRHTelefonos(prs As ADODB.Recordset)
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    prs.MoveFirst
    
    If Not prs.EOF And Not prs.BOF Then
        While Not prs.EOF
            sql = " Update RRHH " _
                & " Set cAnexoInterno = '" & prs.Fields(5) & "', cCelular = '" & prs.Fields(6) & "', cRPM = '" & prs.Fields(7) & "'" _
                & " Where cRHCod = '" & prs.Fields(0) & "'"
            oCon.Ejecutar sql
            prs.MoveNext
        Wend
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Function CargaCuentasEmpleado(psPersCod As String, psAgeCod As String) As ADODB.Recordset
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
        
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion 'Remota psAgeCod
    
    sql = sql & " Select PC.cCodCta, dAperAC, IsNull((Select upper(TC.cNomTab) from ITFCtaExonerada ITF Inner Join dbComunes..TablaCod TC On ITF.cTipo = TC.cValor Where TC.cCodTab Like 'PA__' And ITF.cCodCta = AHO.cCodCta) ,'NO EXONERADO') ITF"
    sql = sql & " From PersCuenta PC"
    sql = sql & " Inner Join Ahorroc AHO On PC.cCodCta = AHO.cCodCta"
    sql = sql & " Where PC.cCodPers = '" & Right(psPersCod, 10) & "' And PC.cRelaCta = 'TI' And cEstCtaAC Not In ('C','U')"
    sql = sql & " UNION"
    sql = sql & " Select PC.cCodCta, dAperCTS, 'EXONERADO' ITF"
    sql = sql & " From PersCuenta PC"
    sql = sql & " Inner Join CTS AHO On PC.cCodCta = AHO.cCodCta"
    sql = sql & " Where PC.cCodPers = '" & Right(psPersCod, 10) & "' And PC.cRelaCta = 'TI' And cEstCtaCTS Not In ('C','U')"
    
    Set rs = oCon.CargaRecordSet(sql)
    
    Set CargaCuentasEmpleado = rs
End Function

Public Function GetPersonal(pnIndiLoc As Integer, pnIndiPla As Integer, pnIndiPrac As Integer) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta
Dim Sql2 As String
If pnIndiLoc = 1 Then
    Sql2 = " and R.cRHCod  like 'L%' "
ElseIf pnIndiPla = 1 Then
    Sql2 = " and R.cRHCod  like 'E%' and nRHEstado = 201 "
ElseIf pnIndiPrac = 1 Then
    Sql2 = " and R.cRHCod  like 'P%' "
End If

sql = "SELECT   R.cRHCod , R.CPERSCOD,P.cPersNombre, cRHCod, CUSER, R.cAreaCod,cAgenciaAsig,T.cRHCargoCod, T.cRHCargoDescripcion,cAreaDesResumen,A.cAreaDescripcion   " _
    & " FROM    RRHH R  " _
    & "         JOIN PERSONA P ON P.CPERSCOD = R.CPERSCOD " _
    & "         JOIN RHCARGOS C ON C.CPERSCOD = R.CPERSCOD " _
    & "         JOIN RHCARGOSTABLA T ON T.cRHCargoCod = C.cRHCargoCod JOIN AREAS A ON A.cAreaCod  = R.cAreaCod " _
    & " WHERE C.dRHCargoFecha = (SELECT MAX(dRHCargoFecha) " _
    & "                          FROM RHCARGOS C1 " _
    & "                          WHERE C1.CPERSCOD=R.CPERSCOD) "
    
    sql = sql + Sql2 + " ORDER BY cPersNombre asc  "

If oCon.AbreConexion Then
    Set rs = oCon.CargaRecordSet(sql)
    Set GetPersonal = rs
    oCon.CierraConexion
End If
Set oCon = Nothing
End Function

Public Function ModificaCUSSPSeguro(psPersCod As String, psCUSPP As String, psSeguro As String, pdFecIng As Date, psUltimoMovimiento As String, psFecCese As String, psFecAfilia As String) As Boolean
    On Error GoTo ModificaRRHHErr
    Dim sqlC As String
    Dim sqlI As String
    Dim oCon As DConecta
    Dim sqlAdd As String
    Set oCon = New DConecta
    
    If IsDate(psFecCese) Then
        sqlAdd = " '" & Format(psFecCese, gsFormatoFecha) & "' "
    Else
        sqlAdd = " Null "
    End If
    
    If IsDate(psFecAfilia) Then
        psFecAfilia = " '" & Format(psFecAfilia, gsFormatoFecha) & "' "
        Else
        psFecAfilia = " Null "
    End If
    
    
    sqlC = " Update RHEmpleado Set cCUSPP = '" & psCUSPP & "' , cNumAFP = " & psSeguro & ",  " _
         & " cUltimaActualizacion = '" & psUltimoMovimiento & "', " _
         & " dFecAfiliacion  = " & psFecAfilia & " " _
         & " Where cPersCod = '" & psPersCod & "'"
    
    sqlI = " Update RRHH " _
         & " Set dIngreso = '" & Format(pdFecIng, gsFormatoFecha) & "', dCese = " & sqlAdd & "" _
         & " Where cPersCod = '" & psPersCod & "'"
         
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.Ejecutar sqlI
        oCon.CierraConexion
    End If
            
    Set oCon = Nothing
    Exit Function
ModificaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:ModificaRRHH Method")
End Function

Public Function GetCUSSPSeguro(psPersCod As String) As ADODB.Recordset
    On Error GoTo ModificaRRHHErr
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlC = " Select cCUSPP , cNumAFP,dFecAfiliacion From  RHEmpleado " _
         & " Where cPersCod = '" & psPersCod & "'"
    
    If oCon.AbreConexion Then
        Set GetCUSSPSeguro = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
            
    Set oCon = Nothing
    Exit Function
ModificaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:ModificaRRHH Method")
End Function


Public Function GetHorasLaboradas_Extras(ByVal pdFechaIni As Date, ByVal pdFechaFin As Date, ByVal psPersCod As String) As ADODB.Recordset
Dim oCon As DConecta
Dim sql As String
Set oCon = New DConecta
On Error GoTo GetHorasLaboradas_ExtrasErr

sql = " Select he,horasLaboradas," & _
      " (Isnull(Inasistencia,0)+Isnull(Inasistencia1,0)+isnull(Inasistencia2,0)+isnull(Inasistencia3,0))Inasistencia " & _
      " ,Tardanza FROM getConsolidadoHE('" & pdFechaIni & "','" & Format(pdFechaFin, "mm/dd/yyyy") & "','" & psPersCod & "' )"

If oCon.AbreConexion Then
        Set GetHorasLaboradas_Extras = oCon.CargaRecordSet(sql)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetHorasLaboradas_ExtrasErr:
Call RaiseError(MyUnhandledError, "GetHorasLaboradas_Extras:GetHorasLaboradas_Extras Method")
End Function

Public Function GetRHPlanillaDetConFractal(ByVal psPersCod As String, ByVal psPlanillaTpo As String, ByVal psPlanillaPeriodo As String) As ADODB.Recordset
    On Error GoTo GetRHPlanillaDetConFractal
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlC = " Select cRRHHPeriodo, cPlanillaCod, cPersCod, cRHConceptoCod, cDato, nMonto, bPagado, cCtaCod, cRHCod, cPersNombre, nMontoIntangibleCTS From RHPlanillaDetCon " _
         & " Where cRHCod = '" & psPersCod & "' And cPlanillaCod = '" & psPlanillaTpo & "' And cRRHHPeriodo = '" & psPlanillaPeriodo & "'"
    
    If oCon.AbreConexion Then
        Set GetRHPlanillaDetConFractal = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
            
    Set oCon = Nothing
    Exit Function
GetRHPlanillaDetConFractal:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetRHPlanillaDetConFractal Method")
End Function

'*** PEAC 20131112
Public Function GetEstadoCuentaRRHH(psCodCta As String) As String
    On Error GoTo GetCargoErr
   
    Dim sqlC As String
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Dim oCon As DConecta
    Dim lsCadTemp As String
    Set oCon = New DConecta
    Dim lbBitCentral As Boolean
    
    lsCadTemp = ""
  
    Dim oConst As NConstSistemas
    Set oConst = New NConstSistemas
    
    Dim oConexionRem As DConecta
    Set oConexionRem = New DConecta
    Dim rsValida As ADODB.Recordset
    Set rsValida = New ADODB.Recordset
    Dim sql As String
        
    sqlC = "exec stp_sel_EstadoCuenta '" & psCodCta & "'"
    
    If oCon.AbreConexion Then
        Set rsC = oCon.CargaRecordSet(sqlC)
        If rsC.EOF And rsC.BOF Then
            GetEstadoCuentaRRHH = ""
        Else
            GetEstadoCuentaRRHH = Trim(CStr(rsC!nPrdEstado))
        End If
    End If
        
    rsC.Close
    Set rsC = Nothing
    Set oCon = Nothing
    Exit Function
      
GetCargoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetCargo Method")
End Function

'*** PEAC 20131119
Public Function GetTitularCuentaRRHH(psCodCta As String) As String
    On Error GoTo GetTitularCuentaRRHHErr
   
    Dim sqlC As String
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Dim oCon As DConecta
    Dim lsCadTemp As String
    Set oCon = New DConecta
    Dim lbBitCentral As Boolean
    
    lsCadTemp = ""
  
    Dim oConst As NConstSistemas
    Set oConst = New NConstSistemas
    
    Dim oConexionRem As DConecta
    Set oConexionRem = New DConecta
    Dim rsValida As ADODB.Recordset
    Set rsValida = New ADODB.Recordset
    Dim sql As String
        
    sqlC = "exec stp_sel_TitularCuenta '" & psCodCta & "'"
    
    If oCon.AbreConexion Then
        Set rsC = oCon.CargaRecordSet(sqlC)
        If rsC.EOF And rsC.BOF Then
            GetTitularCuentaRRHH = ""
        Else
            GetTitularCuentaRRHH = Trim(CStr(rsC!cPersNombre))
        End If
    End If
        
    rsC.Close
    Set rsC = Nothing
    Set oCon = Nothing
    Exit Function
      
GetTitularCuentaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:GetTitularCuentaRRHH Method")
End Function

'RECO20140220 ERS164-2013************************************************************
Public Sub RegistraAsigRetiroGrupo(ByVal psFecha As String, ByVal psGrupo As String, ByVal psUser As String, ByVal psUserReg As String, ByVal pnTipoCambio As Integer)
    On Error GoTo RegistraAsignacionGrupoErr
    Dim sSql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    psFecha = Format(psFecha, "yyyy/MM/dd") & " " & oCon.GetHoraServer
    sSql = "EXECUTE stp_Del_RegistraAsigRetiroGrupo '" & psFecha & "','" & psGrupo & "','" & psUser & "','" & psUserReg & "'," & pnTipoCambio
    oCon.Ejecutar sSql
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Sub
RegistraAsignacionGrupoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:RegistraAsignacionGrupo Method")
End Sub
'RECO FIN****************************************************************************
'APRI20191025 MEJORA PROCESO
Public Function ObtenerCapSaldosCuentasCTS(ByVal sCuenta As String, nTipoCambio As Currency) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As DConecta
    Dim lrDatos As ADODB.Recordset
        
        lsSQL = "Exec stp_sel_ObtieneSaldosYSueldosCTS '" & sCuenta & "'," & nTipoCambio
        
        Set oConec = New DConecta
            oConec.AbreConexion
        Set lrDatos = oConec.CargaRecordSet(lsSQL, adLockReadOnly)
        Set ObtenerCapSaldosCuentasCTS = lrDatos
        Set lrDatos = Nothing
End Function
'END APRI
