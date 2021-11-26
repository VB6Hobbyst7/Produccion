VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NActualizaDatosRRHH"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A945499017B"
'Actualiza el RRHH de la Empresa
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902D4023E
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902D40324
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902D40310
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
'##ModelId=3A9476990301
Public Function AgregaRRHH(psPersCod As String, psRHCod As String, psUser As String, psClave As String, psAreaCod As String, psAgenciaAsig As String, psAgenciaActual As String, psIngreso As String, psCese As String, psUltimaActualizacion As String, psEstado As String, psFechaSis As String, psRHNivelCod As String, psRHAMP As String, psRHAFP As String, pnRHDiasVacaciones As Integer, psContratado As String, psFechaEstable As String, psCarCod As String) As String
    Dim oRH As DActualizaDatosRRHH
    Dim lsCodRH As String
    Set oRH = New DActualizaDatosRRHH
    On Error GoTo AgregaRRHHErr
    
    lsCodRH = oRH.GetUltCorr(psRHCod)

    If Not oRH.EsEmpleado(psPersCod) Then
        oRH.AgregaRRHH psPersCod, lsCodRH, psUser, psClave, psAreaCod, psAgenciaAsig, psAgenciaActual, psIngreso, psCese, psUltimaActualizacion, psEstado
        If Left(lsCodRH, 1) = "E" Then oRH.AgregaRRHHEmp psPersCod, psRHNivelCod, psRHAMP, psRHAFP, pnRHDiasVacaciones, psContratado, psFechaEstable, psUltimaActualizacion
        oRH.AgregaRHEstado psPersCod, psIngreso, psEstado, "INGRESO EN " & psIngreso, psUltimaActualizacion
        oRH.AgregaRHCargo psPersCod, psFechaSis, psCarCod, psAgenciaAsig, psCarCod, psAgenciaActual, "INGRESO EN " & psIngreso, psUltimaActualizacion
    Else
        oRH.ModificaRRHH psPersCod, lsCodRH, psUser, psClave, psAreaCod, psAgenciaAsig, psAgenciaActual, psIngreso, psCese, psUltimaActualizacion, psEstado
        If Left(lsCodRH, 1) = "E" Then oRH.ModificaRRHHEmp psPersCod, psRHNivelCod, psRHAMP, psRHAFP, pnRHDiasVacaciones, psContratado, psFechaEstable, psUltimaActualizacion
        oRH.AgregaRHEstado psPersCod, psFechaSis, psEstado, "RE-INGRESO EN " & psIngreso, psUltimaActualizacion
        oRH.AgregaRHCargo psPersCod, psFechaSis, psCarCod, psAgenciaAsig, psCarCod, psAgenciaActual, "INGRESO EN " & psIngreso, psUltimaActualizacion
    End If
    
    AgregaRRHH = lsCodRH
    
    Set oRH = Nothing
    Exit Function
AgregaRRHHErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosRRHH:AgregaRRHH Method")
End Function

'Agrega un nuevo cargo o nivel
'##ModelId=3A9476990301
Public Function GetListaRRHH(ByVal psTitulo As String, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date, Optional psTipo As String = "", Optional pnEstado As Integer = -1) As String
    Dim oRH As DActualizaDatosRRHH
    Dim rs As ADODB.Recordset
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    Set rs = New ADODB.Recordset
    Set oRH = New DActualizaDatosRRHH
    
    On Error GoTo AgregaRRHHErr
        
    Dim lsCodigo As String * 8
    Dim lsNombre As String * 63
    Dim lsUsuario As String * 6
    Dim lsEstado As String * 20
        
    Set rs = oRH.GetListaRRHH(psTipo, Trim(Str(pnEstado)))
    
    If Not (rs.EOF And rs.BOF) Then
        lsCadena = lsCadena & CabeceraPagina(psTitulo, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        lsCadena = lsCadena & Encabezado("Codigo;8;Nombre;30; ;30;Usu;6;Estado;8; ;37;", lnItem)
        While Not rs.EOF
            lsCodigo = rs!Cod
            lsNombre = PstaNombre(rs!Nombre)
            lsUsuario = rs!Usuario
            lsEstado = rs!Estado
            lsCadena = lsCadena & lsCodigo & lsNombre & lsUsuario & lsEstado & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1
            If lnItem = 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina(psTitulo, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Codigo;8;Nombre;30; ;30;Usu;6;Estado;8; ;37;", lnItem)
            End If
            rs.MoveNext
        Wend
    End If
    
    GetListaRRHH = lsCadena
    Set oRH = Nothing
    Exit Function
AgregaRRHHErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosRRHH:AgregaRRHH Method")
End Function

'Agrega un nuevo cargo o nivel
'##ModelId=3A9476990301
Public Function RescindeRRHH(psPersCod As String, psContratoCod As String, psFecha As String, psMotivoTpo As String, psComentario As String, psUltimaActualizacion As String) As Boolean
'Public Function RescindeRRHH(psPersCod As String, psContratoCod As String, psFecha As String, psMotivoTpo As String, psComentario As String, psUltimaActualizacion As String, psMotivo2Tpo As String) As Boolean
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    On Error GoTo RescindeRRHHErr

    oRH.RescindeRRHH psPersCod, psContratoCod, psFecha, psMotivoTpo, psComentario, psUltimaActualizacion, ""
    'oRH.RescindeRRHH psPersCod, psContratoCod, psFecha, psMotivoTpo, psComentario, psUltimaActualizacion, psMotivo2Tpo
    
    
    Set oRH = Nothing
    Exit Function
RescindeRRHHErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosRRHH:AgregaRRHH Method")
End Function


'Elimina  RRHH, esta eliminacion es logica
'##ModelId=3A9476990329
Public Function EliminaRRHH(psRHPersCod As String) As Boolean
    On Error GoTo EliminaRRHHErr
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    
    oRH.EliminaRRHH psRHPersCod

    Exit Function
EliminaRRHHErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosRRHH:EliminaRRHH Method")
End Function

'Actualiza un RRHH
'##ModelId=3A9476990347
Public Function ModificaRRHH(psPersCod As String, psRHCod As String, psUser As String, psClave As String, psAreaCod As String, psAgenciaAsig As String, psAgenciaActual As String, psIngreso As String, psCese As String, psUltimaActualizacion As String, psEstado As String, psFechaSis As String, psRHNivelCod As String, psRHAMP As String, psRHAFP As String, pnRHDiasVacaciones As String, psContratado As String, psFechaEstable As String, psCarCod As String) As Boolean
    On Error GoTo ModificaRRHHErr
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    
    oRH.ModificaRRHH psPersCod, psRHCod, psUser, psClave, psAreaCod, psAgenciaAsig, psAgenciaActual, psIngreso, psCese, psUltimaActualizacion, psEstado
    If Left(psRHCod, 1) = "E" Then oRH.ModificaRRHHEmp psPersCod, psRHNivelCod, psRHAMP, psRHAFP, Str(pnRHDiasVacaciones), psContratado, psFechaEstable, psUltimaActualizacion
    If psEstado <> "" Then oRH.AgregaRHEstado psPersCod, psFechaSis, psEstado, "MODIFICADO EN " & psUltimaActualizacion, psUltimaActualizacion
    If psCarCod <> "" Then oRH.AgregaRHCargo psPersCod, psFechaSis, psCarCod, psAgenciaAsig, psCarCod, psAgenciaActual, "MODIFICADO EN " & psUltimaActualizacion, psUltimaActualizacion
    
    Set oRH = Nothing
    Exit Function
ModificaRRHHErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosRRHH:ModificaRRHH Method")
End Function


'Actualiza un RRHH
'##ModelId=3A9476990347
Public Function ModificaCuentaTarj(prRS As ADODB.Recordset, psTarj As String, psPersCod As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaRRHHErr
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    
    oRH.ModificaCuenatTarj prRS, psTarj, psPersCod, psUltimaActualizacion
    
    Set oRH = Nothing
    Exit Function
ModificaRRHHErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosRRHH:ModificaRRHH Method")
End Function


'Actualiza un RRHH
'##ModelId=3A9476990347
'Public Function ModificaUsuario(psPersCod As String, psUser As String, psUltimaActualizacion As String, pnAgregaAPlanilla As Integer, pnCategoriaViaticos As Integer, pnMarcFotoCheck As Integer, pnMarcPC As Integer, pnConfianza As Integer, pnNSF As Integer, pnPersDireccion As Integer) As Boolean
'    On Error GoTo ModificaRRHHErr
'    Dim oRH As DActualizaDatosRRHH
'    Set oRH = New DActualizaDatosRRHH
'
'    oRH.ModificaUsuario psPersCod, psUser, psUltimaActualizacion, pnAgregaAPlanilla, pnCategoriaViaticos, pnMarcFotoCheck, pnMarcPC, pnConfianza, pnNSF, pnPersDireccion
'
'    Set oRH = Nothing
'    Exit Function
'ModificaRRHHErr:
'    Call RaiseError(MyUnhandledError, "NActualizaDatosRRHH:ModificaRRHH Method")
'End Function

'Agrega datos adicionales del los RRHH de tipo Empleados
'##ModelId=3A94889701E1
Public Function AgregaRRHHEmp(psRHEmplPersCod As String, pdRHEmplIngreso As Date, psUser As String, psRHEmplAMPCod As String, psRHEmplAFPPersCod As String, pnRHEmplVacacionesPend As Integer, pbRHEmplCondicion As Boolean, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaRRHHEmpErr

    'your code goes here...

    Exit Function
AgregaRRHHEmpErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosRRHH:AgregaRRHHEmp Method")
End Function

'Elimina  datos adicionales del los RRHH de tipo Empleados
'esta eliminacion es logica
'
'##ModelId=3A9488970209
Public Function EliminaRRHHEmp(psRHEmplPersCod As String) As Boolean
    On Error GoTo EliminaRRHHEmpErr

    'your code goes here...

    Exit Function
EliminaRRHHEmpErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosRRHH:EliminaRRHHEmp Method")
End Function

'Actualiza datos adicionales del los RRHH de tipo Empleados
'##ModelId=3A9488970227
Public Function ModificaRRHHEmp(psRHEmplPersCod As String, pdRHEmplIngreso As Date, psUser As String, psRHEmplAMPCod As String, psRHEmplAFPPersCod As String, pnRHEmplVacacionesPend As Integer, pbRHEmplCondicion As Boolean, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaRRHHEmpErr

    'your code goes here...

    Exit Function
ModificaRRHHEmpErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosRRHH:ModificaRRHHEmp Method")
End Function

Public Function ModificaCUSSPSeguroDiasVacac(psPersCod As String, psCUSPP As String, psSeguro As String, pdFecIng As Date, psUltimoMovimiento As String, psFecCese As String, psDiasVacac As String) As Boolean
    On Error GoTo ModificaRRHHErr
    Dim oCon As DActualizaDatosRRHH
    Set oCon = New DActualizaDatosRRHH
    
    oCon.ModificaCUSSPSeguroDiasVacac psPersCod, psCUSPP, psSeguro, pdFecIng, psUltimoMovimiento, psFecCese, psDiasVacac
    
    Set oCon = Nothing
    Exit Function
ModificaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:ModificaRRHH Method")
End Function

Public Function AgregaRRHHLote(prRRHH As ADODB.Recordset, psAreaCod As String, psCargo As String, psRHTpoContrato As String, psUltimaActualizacion As String, psFecha As String, psFormatoFecha As String) As Boolean
    On Error GoTo ModificaRRHHEmpErr
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    
    oRH.AgregaRRHHLote prRRHH, psAreaCod, psCargo, psRHTpoContrato, psUltimaActualizacion, psFecha, psFormatoFecha
    
ModificaRRHHEmpErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosRRHH:ModificaRRHHEmp Method")
End Function

'Actualiza un RRHH
'##ModelId=3A9476990347
Public Function CargaCuentasEmpleado(psPersCod As String, cAgeCod As String) As ADODB.Recordset
    On Error GoTo ModificaRRHHErr
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    
    Set CargaCuentasEmpleado = oRH.CargaCuentasEmpleado(psPersCod, cAgeCod)
    
    Set oRH = Nothing
    Exit Function
ModificaRRHHErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosRRHH:ModificaRRHH Method")
End Function


#If DebugMode Then
    '##ModelId=3AB902D4025C
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

'Actualiza un RRHH
'##ModelId=3A9476990347
Public Function ValidaCtasEmpleado(prRS As ADODB.Recordset, pgbBitCentral As Boolean) As Boolean
    On Error GoTo ModificaRRHHErr
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    
    ValidaCtasEmpleado = oRH.ValidaCtasEmpleado(prRS, pgbBitCentral)
    
    Set oRH = Nothing
    Exit Function
ModificaRRHHErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosRRHH:ModificaRRHH Method")
End Function

'ALPA 20110117**************************************************************************************
Public Function ModificaUsuario(psPersCod As String, psUser As String, psUltimaActualizacion As String, pnAgregaAPlanilla As Integer, Optional psUbicacion As String = "") As Boolean
'Public Function ModificaUsuario(psPersCod As String, psUser As String, psUltimaActualizacion As String, pnAgregaAPlanilla As Integer) As Boolean
    On Error GoTo ModificaRRHHErr
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    
    'oRH.ModificaUsuario psPersCod, psUser, psUltimaActualizacion, pnAgregaAPlanilla
    oRH.ModificaUsuario psPersCod, psUser, psUltimaActualizacion, pnAgregaAPlanilla, psUbicacion
    
    
    Set oRH = Nothing
    Exit Function
ModificaRRHHErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosRRHH:ModificaRRHH Method")
End Function
'***************************************************************************************************

Public Function ModificaCUSSPSeguro(psPersCod As String, psCUSPP As String, psSeguro As String, pdFecIng As Date, psUltimoMovimiento As String, psFecCese As String, psAfilia As String) As Boolean
    On Error GoTo ModificaRRHHErr
    Dim oCon As DActualizaDatosRRHH
    Set oCon = New DActualizaDatosRRHH
    
    oCon.ModificaCUSSPSeguro psPersCod, psCUSPP, psSeguro, pdFecIng, psUltimoMovimiento, psFecCese, psAfilia
    
    Set oCon = Nothing
    Exit Function
ModificaRRHHErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosRRHH:ModificaRRHH Method")
End Function
