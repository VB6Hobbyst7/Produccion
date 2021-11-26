VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'***************************************************
'**** Clase que Maneja los Datos de la persona  ****
'***************************************************
Option Explicit
'En caso de error
Private Enum TiposBusquedaNombre
    BusqApellidoPaterno = 1
    BusqApellidoMaterno = 2
    BusqApellidoCasada = 3
    BusqNombres = 4
End Enum

'nTipoAct :
'           0 - No Hubo cambios
'           1 - Modificado
'           2 - Nuevo
'           3 - Eliminado

Enum TPersCambios
    PersFilaSinCambios = 0
    PersFilaModificada = 1
    PersFilaNueva = 2
    PersFilaEliminda = 3
End Enum

Public sMsgError As String
Public bError As Boolean

Public ConexionActiva As ADODB.Connection
Private gConsPersona As String
Private gConsComunes As String
Private gConsImagenes As String
Private Conn As DConecta

Private sPersCod As String
Private sNombreCompleto As String
Private sApePat As String
Private sApeMat As String
Private sApeCas As String
Private sNombres As String
Private sSexo As String
Private sEstadoCivil As String
Private nHijos As Integer
Private dFechaNac As Date
Private sUbicGeografica As String
Private sDomicilio As String
Private sCondicionDomic As String
Private sTelefonos As String
Private sTelefonos2 As String
Private sCiiu As String
Private sCiiuDescripcion As String
Private sEstado As String
Private nTalla As Double
Private nPeso As Double
Private sTipoSangre As String
Private nTipoPersona As Integer 'Personeria
Private sEmail As String
Private sActualiza As String
Private sPersCodSbs As String
Private sPersRelInst As String


Public RFirma As ADODB.Recordset

'Datos de Persona Juridica
Private sRazonSocial As String
Private sSiglas As String
Private sTipoPersJur As String
Private nNumEmplead As Integer
Private sMagnitudEmp As String

'Estado de la registro de la Persona
Private nTipoAct As TPersCambios

'Tipo de Dato para Manejar los Documentos
Private Type TDoc
    nTipoOrig As Long
    nTipo As Long
    sTipo As String
    sNumero As String
    nTipoAct As TPersCambios 'Para Controlar si es Nuevo, Eliminado o Editado
End Type

'Variable donde se Almacena Documentos  de la Persona
Private Docs() As TDoc
Private nNumDoc As Integer

'Tipo de dato para Almacenar a las Personas con que se tiene Relaciones
Private Type TRel
    sPersCodOrig As String
    sPersCod As String
    sApellNombres As String
    sRelacion As String
    sPersRelacBenef As String
    nPersRelacBenefPorc As Double
    sPersRelacAMP As String
    sPersRelacEstado As String
    sMovAct As String
    nTipoAct As TPersCambios 'Para Controlar si es Nuevo, Eliminado o Editado
End Type
Private Relac() As TRel
Private nNumRel As Integer

'Tipo de Datos para Almacenar las Ftes de Ingreso
Private Type TFteIng
    sTipo As String
    sFteIngresoOrig As String
    sFteIngreso As String
    dFecEvalOrig As Date
    dFecEval As Date
    dFecCaduc As Date
    sCargo As String
    dFteIngreso As Date
    sMoneda As String
    nIngresos As Double
    nIngFam As Double
    nIngOtros As Double
    nIngTotales As Double
    nGastosFam As Double
    nSaldDispFam As Double
    nActivo As Double
    nActivoCircul As Double
    nActivDisp As Double
    nCtasxCob As Double
    nInventarios As Double
    nActFijo As Double
    nPasivoPatrimonio As Double
    nPasivo As Double
    nProveedores As Double
    nCreditosOtros As Double
    nCreditosCmact As Double
    nVentas As Double
    nRecupCtasxCob As Double
    nCompraMercad As Double
    nOtrosEgresos As Double
    sComentarios As String
    sMovNro As String
    sFteRazonSocial As String
    nTipoAct As TPersCambios
End Type
Private FtesIngreso() As TFteIng
Private nNumFtes As Integer

Public Function UbicarFteIngreso(psCodFI As String) As Integer
Dim i As Integer
    
    UbicarFteIngreso = -1
    For i = 0 To nNumFtes - 1
        If FtesIngreso(i).sFteIngresoOrig = psCodFI Then
            UbicarFteIngreso = i
        End If
    Next i
    
End Function

Public Function TiempoFteIngreso() As Integer
Dim sSql As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

On Error GoTo ErrorTiempoFteIngreso
    sSql = "Select nParamValor from " & gConsComunes & "ColocParametro where nParamVar = " & gColocParamFteIngresoTiempoEval
    Set oConecta = New DConecta
    Call oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If Not R.BOF And Not R.EOF Then
        TiempoFteIngreso = CInt(Format(R!nParamValor, "#0"))
    Else
        TiempoFteIngreso = 1
    End If
    R.Close
    Set R = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorTiempoFteIngreso:
    Err.Raise Err.Number, "ReucperaPersona", Err.Description
End Function
Public Function ExistenCambios() As Boolean
Dim i As Integer
    ExistenCambios = False
    If nTipoAct <> PersFilaSinCambios Then
        ExistenCambios = True
        Exit Function
    End If
    For i = 0 To nNumDoc - 1
        If Docs(i).nTipoAct <> PersFilaSinCambios Then
            ExistenCambios = True
            Exit Function
        End If
    Next i
    For i = 0 To nNumRel - 1
        If Relac(i).nTipoAct <> PersFilaSinCambios Then
            ExistenCambios = True
            Exit Function
        End If
    Next i
    For i = 0 To nNumFtes - 1
        If FtesIngreso(i).nTipoAct <> PersFilaSinCambios Then
            ExistenCambios = True
            Exit Function
        End If
    Next i
End Function
Public Sub NuevaPersona()
    Set ConexionActiva = Nothing
    sPersCod = ""
    sNombreCompleto = ""
    sApePat = ""
    sApeMat = ""
    sApeCas = ""
    sNombres = ""
    sSexo = ""
    sEstadoCivil = ""
    nHijos = 0
    sUbicGeografica = ""
    sDomicilio = ""
    sCondicionDomic = ""
    sTelefonos = ""
    sTelefonos2 = ""
    sCiiu = ""
    sEstado = ""
    nTalla = 0
    nPeso = 0
    sTipoSangre = ""
    nTipoPersona = 0
    sEmail = ""
    Set RFirma = Nothing
    sRazonSocial = ""
    sSiglas = ""
    sTipoPersJur = ""
    nNumEmplead = 0
    sMagnitudEmp = ""
    nTipoAct = PersFilaSinCambios
    nNumDoc = 0
    ReDim Docs(0)
    ReDim Relac(0)
    nNumRel = 0
    ReDim FtesIngreso(0)
    nNumFtes = 0
    Set RFirma = RecuperaFirma("")
End Sub
Public Function UltimoCorrelativo(ByVal psCodAge As String) As String
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConecta  As DConecta
Dim ConexDitrib As ADODB.Connection
Dim sCorrelDistrib As String
Dim sIp As String

    On Error GoTo ErrorUltimoCorrelativo
    sSql = "Select MAX(substring(P.cPersCod,6,7)) as Correlativo from " & gConsPersona & "Persona P where substring(P.cPersCod,4,2) = '" & Right(psCodAge, 2) & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If Not R.BOF And Not R.EOF Then
        UltimoCorrelativo = IIf(IsNull(R!Correlativo), 0, R!Correlativo)
    Else
        UltimoCorrelativo = 0
    End If
    R.Close
    Set R = Nothing
    
    '******************************************************************************
    'Si es de la Sede entonces hallar el ultimo correlativo del distribuido
    '******************************************************************************
    If psCodAge = "11207" Then
        sIp = ""
        sSql = "Select nConsSisValor from constsistema where nConsSisCod = '45'"
        Set R = oConecta.CargaRecordSet(sSql)
        If R.RecordCount > 0 Then
            sIp = R!nConsSisValor
        End If
        R.Close
        Set R = Nothing
        
        Set ConexDitrib = New ADODB.Connection
        ConexDitrib.Open "Provider=SQLOLEDB.1;User ID=dbaccess;Initial Catalog=DBPersona;Data Source=" & sIp
        
        sSql = "Select Max(Substring(cCodpers,3,7) as sCorrel From Persona Where cCodPers Like '" & Right(psCodAge, 2) & "%'"
        Set R = New ADODB.Recordset
        R.Open sSql, ConexDitrib, adOpenStatic, adLockReadOnly, adCmdText
        If R.RecordCount > 0 Then
            sCorrelDistrib = R!sCorrel
        End If
        R.Close
        
        If CLng(sCorrelDistrib) > CLng(UltimoCorrelativo) Then
            UltimoCorrelativo = sCorrelDistrib
        End If
        ConexDitrib.Close
        Set ConexDitrib = Nothing
    End If
    '******************************************************************************
    
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
    
ErrorUltimoCorrelativo:
    Set R = Nothing
    Set oConecta = Nothing
    Err.Raise Err.Number, "ReucperaPersona", Err.Description
End Function
Public Sub GrabarCambiosPersona(ByVal psCodAge As String)
Dim sSql As String
Dim InicioTran As Boolean
Dim i As Integer
Dim j As Integer
Dim oConecta As DConecta
Dim oNegPers As NPersona
Dim sCodPersDistrib As String
Dim lsUltCorr As String
Dim lnCorr As Long
Dim lnCodchequeo As Integer
Dim lsCadFinal As String
'************************************************************************
'PARA DISTRIBUIDO
'************************************************************************
Dim ConexDitrib As ADODB.Connection
Dim bPersDistrib As Boolean
Dim R As ADODB.Recordset
Dim sPersRelIDistrib As String
Dim vDocs(4) As String
'************************************************************************
    On Error GoTo ErrorGrabarCambiosPersona
    sPersRelIDistrib = ""
    Select Case sPersRelInst
        Case "1"
            sPersRelIDistrib = "D"
        Case "2"
            sPersRelIDistrib = "F"
        Case "3"
            sPersRelIDistrib = "N"
        Case "4"
            sPersRelIDistrib = "O"
        Case "5"
            sPersRelIDistrib = "T"
    End Select
    
    vDocs(0) = ""
    vDocs(1) = ""
    vDocs(2) = ""
    vDocs(3) = ""
    For i = 0 To UBound(Docs) - 1
        Select Case Trim(Str(Docs(i).nTipo))
            Case "1", "3"
                vDocs(0) = Trim(Str(Docs(i).nTipo))
                vDocs(1) = Docs(i).sNumero
                Exit For
            Case "4"
                vDocs(0) = "2"
                vDocs(1) = Docs(i).sNumero
                Exit For
            Case "5", "6", "7", "8", "9"
                vDocs(0) = "9"
                vDocs(1) = Docs(i).sNumero
                Exit For
            Case "10"
                vDocs(0) = "6"
                vDocs(1) = Docs(i).sNumero
                Exit For
            Case "11"
                vDocs(0) = "5"
                vDocs(1) = Docs(i).sNumero
                Exit For
            Case "12"
                vDocs(0) = "7"
                vDocs(1) = Docs(i).sNumero
                Exit For
            Case "13"
                vDocs(0) = "8"
                vDocs(1) = Docs(i).sNumero
                Exit For
            Case "13"
                vDocs(0) = "8"
                vDocs(1) = Docs(i).sNumero
                Exit For
        End Select
    Next i
    
    For i = 0 To UBound(Docs) - 1
        Select Case Trim(Str(Docs(i).nTipo))
            Case "2"
                vDocs(2) = "2"
                vDocs(3) = Docs(i).sNumero
                Exit For
            Case "14"
                vDocs(2) = "1"
                vDocs(3) = Docs(i).sNumero
                Exit For
            Case "15"
                vDocs(2) = "3"
                vDocs(3) = Docs(i).sNumero
                Exit For
        End Select
    Next i
    
    '***********************************************************************************
    
    If nTipoPersona = 0 Then
        If sSexo <> "F" And Len(Trim(sApeCas)) > 0 Then
            sNombreCompleto = sApePat & "/" & sApeMat & "\" & sApeCas & "," & sNombres
        Else
            sNombreCompleto = sApePat & "/" & sApeMat & "," & sNombres
        End If
    End If
    
    InicioTran = True
    If nTipoAct = PersFilaNueva Then
        If ConexionActiva Is Nothing Then
            Set oConecta = New DConecta
            oConecta.AbreConexion
            Set ConexionActiva = oConecta.ConexionActiva
        End If
    End If
    
    '**********************   Para Actualizar el Distribuido ************************
    sSql = "select nConsSisValor from ConstSistema where nConsSisCod=45"
    Set R = New ADODB.Recordset
    R.Open sSql, ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    If R.RecordCount > 0 Then
        If Len(Trim(R!nConsSisValor)) > 0 Then
            bPersDistrib = True
            Set ConexDitrib = New ADODB.Connection
            ConexDitrib.Open "Provider=SQLOLEDB.1;User ID=dbaccess;PWD=cmact;Initial Catalog=DBPersona;Data Source=" & R!nConsSisValor & ";"
        Else
            bPersDistrib = False
        End If
    End If
    R.Close
    '************************************************************************************
    
    ConexionActiva.BeginTrans
    
    '**********************   Para Actualizar el Distribuido ************************
    If nTipoAct = PersFilaNueva Then
        Set oNegPers = New NPersona
        lsUltCorr = UltimoCorrelativo(psCodAge)
        lnCorr = Val(lsUltCorr) + 1
        lsUltCorr = FillNum(Trim(Str(lnCorr)), 7, "0")
        lnCodchequeo = oNegPers.DigitoChequeo(lsUltCorr)
        lsCadFinal = lsUltCorr & Trim(Str(lnCodchequeo))
        sCodPersDistrib = Right(Trim(psCodAge), 2) & Trim(lsCadFinal)
        sPersCod = Trim(psCodAge) & Trim(lsCadFinal)
        Set oNegPers = Nothing
    Else
        sCodPersDistrib = Right(sPersCod, 10)
    End If
    '********************************************************************************
    Set R = Nothing
    If nTipoAct = PersFilaModificada Then
        
        'Actualiza Persona
        sSql = "UPDATE " & gConsPersona & "Persona SET "
        sSql = sSql & "cPersNombre = '" & sNombreCompleto & "',"
        sSql = sSql & "dPersNacCreac = '" & Format(dFechaNac, "mm/dd/yyyy") & "',"
        sSql = sSql & "cPersDireccUbiGeo = '" & sUbicGeografica & "',"
        sSql = sSql & "cPersDireccDomicilio = '" & sDomicilio & "',"
        sSql = sSql & "cPersDireccCondicion = '" & sCondicionDomic & "',"
        sSql = sSql & "cPersTelefono = '" & sTelefonos & "',"
        sSql = sSql & "cPersTelefono2 = '" & sTelefonos2 & "',"
        sSql = sSql & "cPersEmail = '" & sEmail & "',"
        sSql = sSql & "nPersPersoneria = '" & Trim(Str(nTipoPersona)) & "',"
        sSql = sSql & "cPersCIIU = '" & sCiiu & "',"
        sSql = sSql & "cPersEstado = '" & sEstado & "',"
        sSql = sSql & "nPersRelaInst = " & IIf(Len(Trim(sPersRelInst)) = 0, "''", Trim(sPersRelInst))
        sSql = sSql & " Where cPersCod = '" & sPersCod & "'"
        
        ConexionActiva.Execute sSql
                        
        
        If nTipoPersona = gPersonaNat Then
        
            sSql = "DELETE " & gConsPersona & "PersonaJur"
            sSql = sSql & " WHERE cPersCod = '" & sPersCod & "'"
            ConexionActiva.Execute sSql
        
            sSql = "DELETE " & gConsPersona & "PersonaNat"
            sSql = sSql & " WHERE cPersCod = '" & sPersCod & "'"
            ConexionActiva.Execute sSql
            
            'Actualiza Persona Natural
            sSql = "INSERT INTO " & gConsPersona
            sSql = sSql & "PersonaNat(cPersCod,cPersnatSexo,nPersNatEstCiv,nPersNatHijos,nPersPeso,"
            sSql = sSql & "nPersTalla,cPersTpoSangre,cUltimaActualizacion)"
            sSql = sSql & " VALUES('" & sPersCod & "',"
            sSql = sSql & "'" & sSexo & "',"
            sSql = sSql & sEstadoCivil & ","
            sSql = sSql & Trim(Str(nHijos)) & ","
            sSql = sSql & "'" & Format(nPeso, "#0.00") & "',"
            sSql = sSql & "'" & Format(nTalla, "#0.00") & "',"
            sSql = sSql & "'" & Trim(sTipoSangre) & "',"
            sSql = sSql & "'" & sActualiza & "')"
            ConexionActiva.Execute sSql
              
              
        End If
        
        If nTipoPersona <> gPersonaNat Then
            sSql = "DELETE " & gConsPersona & "PersonaNat"
            sSql = sSql & " WHERE cPersCod = '" & sPersCod & "'"
            ConexionActiva.Execute sSql
            
            sSql = "DELETE " & gConsPersona & "PersonaJur"
            sSql = sSql & " WHERE cPersCod = '" & sPersCod & "'"
            ConexionActiva.Execute sSql
            
            'Actualiza Persona Juridica
            'Inserta PersonaJur
            sSql = "INSERT INTO " & gConsPersona
            sSql = sSql & "PersonaJur(cPersCod,cPersJurTpo,cPersJurMagnitud,cPersJurSigla,nPersJurEmpleados,cUltimaActualizacion)"
            sSql = sSql & "VALUES('" & sPersCod & "',"
            sSql = sSql & "'" & sTipoPersJur & " ',"
            sSql = sSql & "'" & sMagnitudEmp & " ',"
            sSql = sSql & "'" & sSiglas & " ',"
            sSql = sSql & Trim(Str(nNumEmplead)) & ","
            sSql = sSql & "'" & sActualiza & "')"
            ConexionActiva.Execute sSql
                        
            
        End If
        'Actualiza Firma
        If Not RFirma Is Nothing Then
            If Not RFirma.BOF And Not RFirma.EOF Then
                RFirma!cUltimaActualizacion = sActualiza
                RFirma.Update
            End If
        End If
                
    Else 'Para Nueva Persona
        If nTipoAct = PersFilaNueva Then
            'Insertar en Tabla Persona
            sSql = "INSERT INTO " & gConsPersona
            sSql = sSql & "Persona(cPersCod,cPersNombre,dPersNacCreac,cPersDireccUbiGeo,cPersDireccDomicilio"
            sSql = sSql & ",cPersDireccCondicion,cPersTelefono,cPersTelefono2,cPersPersoneria,cPersCIIU,cPersEstado"
            sSql = sSql & ",cPersEmail,cUltimaActualizacion,cPersCodSbs,nPersRelaInst)"
            sSql = sSql & "VALUES('" & sPersCod & "',"
            sSql = sSql & "'" & sNombreCompleto & "',"
            sSql = sSql & "'" & Format(dFechaNac, "mm/dd/yyyy") & "',"
            sSql = sSql & "'" & sUbicGeografica & "',"
            sSql = sSql & "'" & sDomicilio & "',"
            sSql = sSql & "'" & sCondicionDomic & "',"
            sSql = sSql & "'" & sTelefonos & "',"
            sSql = sSql & "'" & sTelefonos2 & "',"
            sSql = sSql & "'" & Trim(Str(nTipoPersona)) & "',"
            sSql = sSql & "'" & sCiiu & "',"
            sSql = sSql & "'" & sEstado & "',"
            sSql = sSql & "'" & sEmail & "',"
            sSql = sSql & "'" & sActualiza & "',"
            sSql = sSql & "'" & sPersCodSbs & "',"
            sSql = sSql & Trim(sPersRelInst) & ")"
            ConexionActiva.Execute sSql
                        
            If nTipoPersona = gPersonaNat Then
                'Insertar en PersonaNat
                sSql = "INSERT INTO " & gConsPersona
                sSql = sSql & "PersonaNat(cPersCod,cPersnatSexo,cPersNatEstCiv,nPersNatHijos,nPersPeso,"
                sSql = sSql & "nPersTalla,cPersTpoSangre,cUltimaActualizacion)"
                sSql = sSql & " VALUES('" & sPersCod & "',"
                sSql = sSql & "'" & sSexo & "',"
                sSql = sSql & "'" & sEstadoCivil & "',"
                sSql = sSql & Trim(Str(nHijos)) & ","
                sSql = sSql & "'" & Format(nPeso, "#0.00") & "',"
                sSql = sSql & "'" & Format(nTalla, "#0.00") & "',"
                sSql = sSql & "'" & sTipoSangre & "',"
                sSql = sSql & "'" & sActualiza & "')"
                ConexionActiva.Execute sSql
                
            End If
            
            If nTipoPersona <> gPersonaNat Then
                'Inserta PersonaJur
                sSql = "INSERT INTO " & gConsPersona
                sSql = sSql & "PersonaJur(cPersCod,cPersJurTpo,cPersJurMagnitud,cPersJurSigla,nPersJurEmpleados,cUltimaActualizacion)"
                sSql = sSql & "VALUES('" & sPersCod & "',"
                sSql = sSql & "'" & sTipoPersJur & " ',"
                sSql = sSql & "'" & sMagnitudEmp & " ',"
                sSql = sSql & "'" & sSiglas & " ',"
                sSql = sSql & Trim(Str(nNumEmplead)) & ","
                sSql = sSql & "'" & sActualiza & "')"
                ConexionActiva.Execute sSql
                                
            End If
            
            'Actualiza Firma
            If Not RFirma Is Nothing Then
                RFirma!cPersCod = sPersCod
                RFirma!cUltimaActualizacion = sActualiza
                RFirma.Update
            End If
        
        End If
    End If
    
    'Graba Relaciones de la Persona
    For i = 0 To nNumRel - 1
        If Relac(i).nTipoAct = PersFilaModificada Then
            sSql = "UPDATE " & gConsPersona & "PersRelaciones SET "
            sSql = sSql & "cPersRelacPersCod = '" & Relac(i).sPersCod & "',"
            sSql = sSql & "cPersRelac = '" & Trim(Right(Relac(i).sRelacion, 20)) & "',"
            sSql = sSql & "bPersRelacBenef = " & Trim(Right(Relac(i).sPersRelacBenef, 20)) & ","
            sSql = sSql & "nPersRelacBenefPorc = " & Format(Relac(i).nPersRelacBenefPorc, "#0.00") & ","
            sSql = sSql & "bPersRelacAMP = " & Trim(Right(Relac(i).sPersRelacAMP, 20))
            'sSql = sSql & "PR.bPersRelacEstado = " & Trim(Right(Relac(i).sPersRelacEstado, 20))
            sSql = sSql & " Where cPersCod = '" & sPersCod & "' And cPersRelacPersCod = '" & Relac(i).sPersCodOrig & "'"
            ConexionActiva.Execute sSql
        Else
            If Relac(i).nTipoAct = PersFilaNueva Then
                sSql = "INSERT INTO " & gConsPersona & "PERSRELACIONES(cPersCod,cPersRelacPersCod,"
                sSql = sSql & "nPersRelac,bPersrelacBenef,nPersRelacBenefPorc,bPersRelacAMP,cMovNro)"
                sSql = sSql & "VALUES('" & sPersCod & "',"
                sSql = sSql & "'" & Relac(i).sPersCod & "',"
                sSql = sSql & Trim(Right(Relac(i).sRelacion, 20)) & ","
                sSql = sSql & Trim(Right(Relac(i).sPersRelacBenef, 20)) & ","
                sSql = sSql & Format(Relac(i).nPersRelacBenefPorc, "#0.00") & ","
                sSql = sSql & Trim(Right(Relac(i).sPersRelacAMP, 20)) & ","
                sSql = sSql & "'" & sActualiza & "')"
                ConexionActiva.Execute sSql
            Else
                If Relac(i).nTipoAct = PersFilaEliminda Then
                    sSql = "DELETE " & gConsPersona & "PersRelaciones Where "
                    sSql = sSql & " cPersCod = '" & sPersCod & "'"
                    sSql = sSql & " And cPersRelacPersCod = '" & Trim(Relac(i).sPersCodOrig) & "'"
                    ConexionActiva.Execute sSql
                End If
            End If
        End If
    Next i
    
    'Actualizar Los Documentos de las Personas
    For i = 0 To nNumDoc - 1
        If Docs(i).nTipoAct = PersFilaModificada Then
            sSql = "UPDATE " & gConsPersona & "PERSID SET "
            sSql = sSql & "cPersIdTpo = '" & Trim(Str(Docs(i).nTipo)) & "',"
            sSql = sSql & "cPersIDNro = '" & Docs(i).sNumero & "',"
            sSql = sSql & "cUltimaActualizacion = '" & sActualiza & "' "
            sSql = sSql & " WHERE cPersCod = '" & sPersCod & "' And cPersIDTpo = '" & Trim(Str(Docs(i).nTipoOrig)) & "'"
            ConexionActiva.Execute sSql
        Else
            If Docs(i).nTipoAct = PersFilaNueva Then
                sSql = "INSERT INTO " & gConsPersona & "PERSID("
                sSql = sSql & "cPersCod,cPersIDTpo,cPersIDNro,cUltimaActualizacion)"
                sSql = sSql & "VALUES('" & sPersCod & "',"
                sSql = sSql & "'" & Trim(Str(Docs(i).nTipo)) & "',"
                sSql = sSql & "'" & Docs(i).sNumero & "',"
                sSql = sSql & "'" & sActualiza & "')"
                ConexionActiva.Execute sSql
            Else
                If Docs(i).nTipoAct = PersFilaEliminda Then
                    sSql = "DELETE " & gConsPersona & "PERSID Where "
                    sSql = sSql & " cPersCod = '" & sPersCod & "'"
                    sSql = sSql & " And cPersIDTpo = '" & Trim(Str(Docs(i).nTipoOrig)) & "'"
                    ConexionActiva.Execute sSql
                End If
            End If
        End If
    Next i
    
    'Actualizar las Fuentes de Ingreso
    For i = 0 To nNumFtes - 1
        If FtesIngreso(i).nTipoAct = PersFilaNueva Then
            sSql = "INSERT INTO " & gConsPersona & "PersFteIngreso(cPersCod,cPersFIPersCod,dPersFIEvaluacion,"
            sSql = sSql & "dPersFICaduca,cPersFICargo,dPersFIInicio,cPersFIMoneda,nPersFIIngresos,"
            sSql = sSql & "nPersFIIngresosFam,nPersFIIngresosOtros,nPersFIGastosFam,cPersFIComentario,cUltimaActualizacion)"
            sSql = sSql & " VALUES('" & sPersCod & "',"
            sSql = sSql & "'" & FtesIngreso(i).sFteIngreso & "',"
            sSql = sSql & "'" & Format(FtesIngreso(i).dFecEval, "mm/dd/yyyy") & "',"
            sSql = sSql & "'" & Format(FtesIngreso(i).dFecCaduc, "mm/dd/yyyy") & "',"
            sSql = sSql & "'" & FtesIngreso(i).sCargo & "',"
            sSql = sSql & "'" & Format(FtesIngreso(i).dFteIngreso, "mm/dd/yyyy") & "',"
            sSql = sSql & Trim(FtesIngreso(i).sMoneda) & ","
            sSql = sSql & Format(FtesIngreso(i).nIngresos, "#0.00") & ","
            sSql = sSql & Format(FtesIngreso(i).nIngFam, "#0.00") & ","
            sSql = sSql & Format(FtesIngreso(i).nIngOtros, "#0.00") & ","
            sSql = sSql & Format(FtesIngreso(i).nGastosFam, "#0.00") & ","
            sSql = sSql & "'" & Trim(FtesIngreso(i).sComentarios) & "',"
            sSql = sSql & "'" & sActualiza & "')"
            ConexionActiva.Execute sSql
            
            'Si es una fuente de Ingreso Independiente actualizamos
            If FtesIngreso(i).sTipo = gPersFteIngresoTipoIndependiente Then
                sSql = "INSERT INTO " & gConsPersona & "PersFIIndependiente(cPersCod,cPersFIPerscod,"
                sSql = sSql & "dPersFIEvaluacion,nPersFIActivoDisp,nPersFICtasxCobrar,nPersFIInventarios,"
                sSql = sSql & "nPersFIActivosFijos,nPersFIProveedores,nPersFICredCmact,nPersFICredOtros,"
                sSql = sSql & "nPersFIVentas,nPersFIPatrimonio,nPersFICostoVentas,nPersFIRecupCtasxCobrar,"
                sSql = sSql & "nPersFIEgresosOtros,cUltimaActualizacion)"
                sSql = sSql & " VALUES('" & sPersCod & "',"
                sSql = sSql & "'" & FtesIngreso(i).sFteIngreso & "',"
                sSql = sSql & "'" & Format(FtesIngreso(i).dFecEval, "mm/dd/yyyy") & "',"
                sSql = sSql & Format(FtesIngreso(i).nActivDisp, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).nCtasxCob, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).nInventarios, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).nActFijo, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).nProveedores, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).nCreditosCmact, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).nCreditosOtros, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).nVentas, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).nPasivoPatrimonio, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).nCompraMercad, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).nRecupCtasxCob, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).nOtrosEgresos, "#0.00") & ","
                sSql = sSql & "'" & sActualiza & "')"
                ConexionActiva.Execute sSql
            End If
        Else
            If FtesIngreso(i).nTipoAct = PersFilaModificada Then
                sSql = "UPDATE " & gConsPersona & "PersFteIngreso SET "
                sSql = sSql & " cPersFIPersCod = '" & FtesIngreso(i).sFteIngreso & "',"
                sSql = sSql & " dPersFIEvaluacion = '" & Format(FtesIngreso(i).dFecEval, "mm/dd/yyyy") & "',"
                sSql = sSql & " dPersFICaduca = '" & Format(FtesIngreso(i).dFecCaduc, "mm/dd/yyyy") & "',"
                sSql = sSql & " cPersFICargo = '" & FtesIngreso(i).sCargo & "',"
                sSql = sSql & " dPersFIInicio = '" & Format(FtesIngreso(i).dFteIngreso, "mm/dd/yyyy") & "',"
                sSql = sSql & " cPersFIMoneda = " & Trim(FtesIngreso(i).sMoneda) & ","
                sSql = sSql & " nPersFIIngresos = " & Format(FtesIngreso(i).nIngresos, "#0.00") & ","
                sSql = sSql & " nPersFIIngresosFam = " & Format(FtesIngreso(i).nIngFam, "#0.00") & ","
                sSql = sSql & " nPersFIIngresosOtros = " & Format(FtesIngreso(i).nIngOtros, "#0.00") & ","
                sSql = sSql & " nPersFIGastosFam = " & Format(FtesIngreso(i).nGastosFam, "#0.00") & ","
                sSql = sSql & " cPersFIComentario = '" & Trim(FtesIngreso(i).sComentarios) & "',"
                sSql = sSql & " cUltimaActualizacion = '" & sActualiza & "'"
                sSql = sSql & " WHERE cPersCod = '" & sPersCod & "' AND cPersFIPersCod = '" & FtesIngreso(i).sFteIngresoOrig & "' AND dPersFIEvaluacion = '" & Format(FtesIngreso(i).dFecEval, "mm/dd/yyyy") & "'"
                ConexionActiva.Execute sSql
                
                'Si es una fuente de Ingreso Independiente actualizamos
                If FtesIngreso(i).sTipo = gPersFteIngresoTipoIndependiente Then
                    sSql = "UPDATE " & gConsPersona & "PersFIIndependiente SET "
                    sSql = sSql & " cPersFIPersCod = '" & FtesIngreso(i).sFteIngreso & "',"
                    sSql = sSql & " dPersFIEvaluacion = '" & Format(FtesIngreso(i).dFecEval, "mm/dd/yyyy") & "',"
                    sSql = sSql & " nPersFIActivoDisp = " & Format(FtesIngreso(i).nActivDisp, "#0.00") & ","
                    sSql = sSql & " nPersFICtasxCobrar = " & Format(FtesIngreso(i).nCtasxCob, "#0.00") & ","
                    sSql = sSql & " nPersFIInventarios = " & Format(FtesIngreso(i).nInventarios, "#0.00") & ","
                    sSql = sSql & " nPersFIActivosFijos = " & Format(FtesIngreso(i).nActFijo, "#0.00") & ","
                    sSql = sSql & " nPersFIProveedores = " & Format(FtesIngreso(i).nProveedores, "#0.00") & ","
                    sSql = sSql & " nPersFICredCmact = " & Format(FtesIngreso(i).nCreditosCmact, "#0.00") & ","
                    sSql = sSql & " nPersFICredOtros = " & Format(FtesIngreso(i).nCreditosOtros, "#0.00") & ","
                    sSql = sSql & " nPersFIVentas = " & Format(FtesIngreso(i).nVentas, "#0.00") & ","
                    sSql = sSql & " nPersFIPatrimonio = " & Format(FtesIngreso(i).nPasivoPatrimonio, "#0.00") & ","
                    sSql = sSql & " nPersFICostoVentas = " & Format(FtesIngreso(i).nCompraMercad, "#0.00") & ","
                    sSql = sSql & " nPersFIRecupCtasxCobrar = " & Format(FtesIngreso(i).nRecupCtasxCob, "#0.00") & ","
                    sSql = sSql & " nPersFIEgresosOtros = " & Format(FtesIngreso(i).nOtrosEgresos, "#0.00") & ","
                    sSql = sSql & " cUltimaActualizacion = '" & sActualiza & "' "
                    sSql = sSql & " WHERE cPersCod = '" & sPersCod & "' AND cPersFIPersCod = '" & FtesIngreso(i).sFteIngresoOrig & "' AND dPersFIEvaluacion = '" & Format(FtesIngreso(i).dFecEval, "mm/dd/yyyy") & "'"
                    ConexionActiva.Execute sSql
                End If
            Else
                If FtesIngreso(i).nTipoAct = PersFilaEliminda Then
                    If FtesIngreso(i).sTipo = gPersFteIngresoTipoIndependiente Then
                        sSql = "DELETE " & gConsPersona & "PersFIIndependiente "
                        sSql = sSql & " WHERE cPersCod = '" & sPersCod & "' AND cPersFIPersCod = '" & FtesIngreso(i).sFteIngresoOrig & "' AND dPersFIEvaluacion = '" & Format(FtesIngreso(i).dFecEval, "mm/dd/yyyy") & "'"
                        ConexionActiva.Execute sSql
                    End If
                    
                    sSql = "DELETE " & gConsPersona & "PersFteIngreso "
                    sSql = sSql & " WHERE cPersCod = '" & sPersCod & "' AND cPersFIPersCod = '" & FtesIngreso(i).sFteIngresoOrig & "' AND dPersFIEvaluacion = '" & Format(FtesIngreso(i).dFecEval, "mm/dd/yyyy") & "'"
                    ConexionActiva.Execute sSql
                End If
            End If
        End If
    Next i
    
    ConexionActiva.CommitTrans
        
    '*******************************************************************************
    'Replicar Persona a Distribuido
    '*******************************************************************************
    ConexDitrib.BeginTrans
    If nTipoAct = PersFilaModificada Then
        '*********************************************************************************
        'Para Actualizar Persona Distribuida
        '*********************************************************************************
        If bPersDistrib Then
            
            sSql = "UPDATE Persona SET "
            sSql = sSql & " cNomPers = '" & sNombreCompleto & "',"
            sSql = sSql & " cTidoci = '" & vDocs(0) & "',"
            sSql = sSql & " cNudoci = '" & vDocs(1) & "',"
            sSql = sSql & " cTidoTr = '" & vDocs(2) & "',"
            sSql = sSql & " cNudoTr = '" & vDocs(3) & "',"
            sSql = sSql & " cDirPers = '" & sDomicilio & "',"
            sSql = sSql & " cCodZon = '" & sUbicGeografica & "',"
            sSql = sSql & " dFecNac = '" & Format(dFechaNac, "mm/dd/yyyy") & "',"
            sSql = sSql & " cTelPers = '" & sTelefonos & "',"
            sSql = sSql & " cEstPers = '1',"
            sSql = sSql & " cTipPers = '" & IIf(nTipoPersona = 0, 0, nTipoPersona - 1) & "',"
            sSql = sSql & " cCodSBS = '" & sPersCodSbs & "',"
            sSql = sSql & " dIngPers = '" & Mid(sActualiza, 1, 4) & "/" & Mid(sActualiza, 5, 2) & "/" & Mid(sActualiza, 7, 2) & "',"
            sSql = sSql & " dUltAct = '" & Mid(sActualiza, 1, 4) & "/" & Mid(sActualiza, 5, 2) & "/" & Mid(sActualiza, 7, 2) & "',"
            sSql = sSql & " cRelIns = '" & sPersRelIDistrib & "',"
            sSql = sSql & " cConDom = '" & sCondicionDomic & "' "
            sSql = sSql & " Where cCodpers = '" & sCodPersDistrib & "'"
            
            ConexDitrib.Execute sSql
        End If
        
            '******************************************************************************
            'ACTUALIZA PERSONA NATURAL EN DISTRIBUIDO
            '******************************************************************************
            If bPersDistrib Then
            
                sSql = "DELETE PersonaNat"
                sSql = sSql & " WHERE cCodPers = '" & sCodPersDistrib & "'"
                ConexDitrib.Execute sSql
            
                sSql = "DELETE PersonaJur"
                sSql = sSql & " WHERE cCodPers = '" & sCodPersDistrib & "'"
                ConexDitrib.Execute sSql
                
                sSql = "INSERT INTO Personanat(cCodPers, cSexPers, cEstCiv, cCodProf, cCodCon, nNumHij)"
                sSql = sSql & " VALUES('" & sCodPersDistrib & "','" & sSexo & "','" & sEstadoCivil & "','',''," & nHijos & ")"
                ConexDitrib.Execute sSql
                
            End If
            
            '******************************************************************************
            'ACTUALIZA PERSONA JURIDICA EN DISTRIBUIDO
            '******************************************************************************
            If bPersDistrib Then
                                
                sSql = "DELETE PersonaNat "
                sSql = sSql & " WHERE cCodPers = '" & sCodPersDistrib & "'"
                ConexDitrib.Execute sSql
            
                sSql = "DELETE PersonaJur "
                sSql = sSql & " WHERE cCodPers = '" & sCodPersDistrib & "'"
                ConexDitrib.Execute sSql
                                
                sSql = "Insert into PersonaJur(cCodPers,cTipPerJur,cCodRegPub,cMagEmp,cSigla) "
                sSql = sSql & "Values('" & sCodPersDistrib & "','0000','','','" & sSiglas & "')"
                ConexDitrib.Execute sSql
                
            End If
    Else
    
            '*********************************************************************************
            'Para INSERTAR Persona Distribuida
            '*********************************************************************************
            If bPersDistrib Then
                sSql = "INSERT INTO Persona(cCodPers,cNomPers,cTidoci,cNudoci,cTidotr,cNudotr,cDirPers, cCodZon, "
                sSql = sSql & " dFecNac, cTelPers, cEstPers, cTipPers, cCodSBS, dIngPers, dUltAct, cCodUsu, cRelIns, cConDom, cFlagConsol, cEMail, cWeb, cTelefono2) "
                sSql = sSql & " VALUES('" & sCodPersDistrib & "','" & sNombreCompleto & "','" & vDocs(0) & "','" & vDocs(1) & "','" & vDocs(2) & "','" & vDocs(3) & "','"
                sSql = sSql & sDomicilio & "','" & sUbicGeografica & "','" & Format(dFechaNac, "mm/dd/yyyy") & "','" & sTelefonos & "','" & sEstado & "','"
                sSql = sSql & CInt(nTipoPersona) - 1 & "','" & sPersCodSbs & "','" & Mid(sActualiza, 1, 4) & "/" & Mid(sActualiza, 5, 2) & "/" & Mid(sActualiza, 7, 2) & "','" & Mid(sActualiza, 1, 4) & "/" & Mid(sActualiza, 5, 2) & "/" & Mid(sActualiza, 7, 2)
                sSql = sSql & "','','" & sPersRelIDistrib & "','" & sCondicionDomic & "',NULL,'','','')"
                
                ConexDitrib.Execute sSql
            End If
            
            '******************************************************************************
            'ACTUALIZA PERSONA NATURAL EN DISTRIBUIDO
            '******************************************************************************
            If bPersDistrib Then
                sSql = "INSERT INTO Personanat(cCodPers, cSexPers, cEstCiv, cCodProf, cCodCon, nNumHij)"
                sSql = sSql & " VALUES('" & sCodPersDistrib & "','" & sSexo & "','" & sEstadoCivil & "','',''," & nHijos & ")"
                ConexDitrib.Execute sSql
            End If
            
            '******************************************************************************
            'ACTUALIZA PERSONA JRUDICA EN DISTRIBUIDO
            '******************************************************************************
            If bPersDistrib Then
                sSql = "Insert into PersonaJur(cCodPers,cTipPerJur,cCodRegPub,cMagEmp,cSigla) "
                sSql = sSql & "Values('" & sCodPersDistrib & "','0000','','','" & sSiglas & "')"
                ConexDitrib.Execute sSql
            End If
            
    End If
    ConexDitrib.CommitTrans
    ConexDitrib.Close
    Set ConexDitrib = Nothing
    
    InicioTran = False
    nTipoAct = PersFilaSinCambios
    i = 0
    Do While i <= nNumDoc - 1
        If Docs(i).nTipoAct = PersFilaEliminda Then
            For j = i To nNumDoc - 2
                Docs(i) = Docs(i + 1)
            Next j
            nNumDoc = nNumDoc - 1
            ReDim Preserve Docs(nNumDoc)
            i = i - 1
        Else
            Docs(i).nTipoAct = PersFilaSinCambios
        End If
        i = i + 1
    Loop
    i = 0
    Do While i <= nNumRel - 1
        If Relac(i).nTipoAct = PersFilaEliminda Then
            For j = i To nNumRel - 2
                Relac(i) = Relac(i + 1)
            Next j
            nNumRel = nNumRel - 1
            ReDim Preserve Relac(nNumRel)
            i = i - 1
        Else
            Relac(i).nTipoAct = PersFilaSinCambios
        End If
        i = i + 1
    Loop
    i = 0
    Do While i <= nNumFtes - 1
        If FtesIngreso(i).nTipoAct = PersFilaEliminda Then
            For j = i To nNumFtes - 2
                FtesIngreso(i) = FtesIngreso(i + 1)
            Next j
            nNumFtes = nNumFtes - 1
            ReDim Preserve FtesIngreso(nNumFtes)
            i = i - 1
        Else
            FtesIngreso(i).nTipoAct = PersFilaSinCambios
        End If
        i = i + 1
    Loop
    Exit Sub
    
ErrorGrabarCambiosPersona:
    If InicioTran Then
        ConexionActiva.RollbackTrans
    End If
    Err.Raise Err.Number, "GrabaPersona", Err.Description
End Sub
Public Sub EditarDocumento(ByVal pnTipo As Long, ByVal psNumero As String)
Dim i As Integer
    For i = 0 To nNumDoc - 1
        If Docs(i).sTipo = pnTipo And Docs(i).sNumero = psNumero Then
            Docs(i).sTipo = pnTipo
            Docs(i).sNumero = psNumero
            Exit For
        End If
    Next i
End Sub
Public Sub EliminarDocumento(ByVal pnTipo As Long, ByVal psNumero As String)
Dim i As Integer
Dim nPos  As Integer
    nPos = -1
    For i = 0 To nNumDoc - 1
        If Docs(i).sTipo = pnTipo And Docs(i).sNumero = psNumero Then
            nPos = i
            Exit For
        End If
    Next i
    If nPos <> -1 Then
        For i = nPos To nNumDoc - 2
            Docs(i) = Docs(i + 1)
        Next i
    End If
    nNumDoc = nNumDoc - 1
    ReDim Preserve Docs(nNumDoc)
End Sub
Public Sub AdicionaPersonaRelacion()
    nNumRel = nNumRel + 1
    ReDim Preserve Relac(nNumRel)
    Relac(nNumRel).nTipoAct = PersFilaNueva
End Sub
Public Sub AdicionaFteIngreso()
    nNumFtes = nNumFtes + 1
    ReDim Preserve FtesIngreso(nNumFtes)
    FtesIngreso(nNumFtes - 1).nTipoAct = PersFilaNueva
End Sub
Public Sub AdicionaDocumento(ByVal psPersIDTpo As String, ByVal psPersIDNro As String, ByVal psConsDescripcion As String)
    nNumDoc = nNumDoc + 1
    ReDim Preserve Docs(nNumDoc)
    Docs(nNumDoc - 1).nTipo = CInt(psPersIDTpo)
    Docs(nNumDoc - 1).nTipoAct = CInt(psPersIDTpo)
    Docs(nNumDoc - 1).nTipoOrig = CInt(psPersIDTpo)
    Docs(nNumDoc - 1).sNumero = psPersIDNro
    Docs(nNumDoc - 1).sTipo = psConsDescripcion
    Docs(nNumDoc - 1).nTipoAct = PersFilaNueva
End Sub
Private Sub CargaDocPers(ByVal psCodPers As String)
Dim Conn As DConecta
Dim sSql As String
Dim R As ADODB.Recordset

    sSql = "Select C.cConsDescripcion,PI.cPersCod,PI.cPersIDtpo,PI.cPersIDnro from PersID PI left join Constante C ON PI.cPersIDtpo = C.cConsValor"
    sSql = sSql & " where PI.cPersCod = '" & psCodPers & "' And C.cConsCod = " & Trim(Str(gPersIdTipo))
    Set R = Conn.CargaRecordSet(sSql, adLockReadOnly)
    Do While Not R.EOF
        Call AdicionaDocumento(R!cPersIDTpo, R!cPersIDNro, R!cConsDescripcion)
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Conn.CierraConexion
    Set Conn = Nothing
End Sub

Private Sub Class_Initialize()
Dim CIni As ClsIni.ClasIni
    Set CIni = New ClsIni.ClasIni
    gConsPersona = CIni.BasePersonas
    gConsComunes = CIni.BaseComunes
    gConsImagenes = CIni.BaseImagenes
    Set CIni = Nothing
    nNumDoc = 0
    nNumRel = 0
End Sub
Private Function BuscaNombre(ByVal psNombre As String, ByVal nTipoBusqueda As TiposBusquedaNombre) As String
Dim sCadTmp As String
Dim PosIni As Integer
Dim PosFin As Integer
    sCadTmp = ""
    Select Case nTipoBusqueda
        Case 1 'Busqueda de Apellido Paterno
            If Mid(psNombre, 1, 1) <> "/" And Mid(psNombre, 1, 1) <> "\" And Mid(psNombre, 1, 1) <> "," Then
                PosIni = 1
                PosFin = InStr(1, psNombre, "/")
                If PosFin = 0 Then
                    PosFin = InStr(1, psNombre, "\")
                    If PosFin = 0 Then
                        PosFin = InStr(1, psNombre, ",")
                        If PosFin = 0 Then
                            PosFin = Len(psNombre)
                        End If
                    End If
                End If
                sCadTmp = Mid(psNombre, PosIni, PosFin - PosIni)
            Else
                sCadTmp = ""
            End If
        Case 2 'Apellido materno
           PosIni = InStr(1, psNombre, "/")
           If PosIni <> 0 Then
                PosIni = PosIni + 1
                PosFin = InStr(1, psNombre, "\")
                If PosFin = 0 Then
                    PosFin = InStr(1, psNombre, ",")
                    If PosFin = 0 Then
                        PosFin = Len(psNombre)
                    End If
                End If
                sCadTmp = Mid(psNombre, PosIni, PosFin - PosIni)
            Else
                sCadTmp = ""
            End If
        Case 3 'Apellido de casada
           PosIni = InStr(1, psNombre, "\")
           If PosIni <> 0 Then
                PosIni = PosIni + 1
                PosFin = InStr(1, psNombre, ",")
                If PosFin = 0 Then
                    PosFin = Len(psNombre)
                End If
                sCadTmp = Mid(psNombre, PosIni, PosFin - PosIni)
            Else
                sCadTmp = ""
            End If
        Case 4 'Nombres
            PosIni = InStr(1, psNombre, ",")
            If PosIni <> 0 Then
                PosIni = PosIni + 1
                PosFin = Len(psNombre)
                sCadTmp = Mid(psNombre, PosIni, (PosFin + 1) - PosIni)
            Else
                sCadTmp = ""
            End If
            
    End Select
    BuscaNombre = sCadTmp
End Function

Private Function RecuperaFirma(ByVal psCodPers As String) As ADODB.Recordset
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConecta As DConecta

    If ConexionActiva Is Nothing Then
        Set oConecta = New DConecta
        oConecta.AbreConexion
        Set ConexionActiva = oConecta.ConexionActiva
    End If
    
    sSql = "Select cPersCod,iPersFirma,cUltimaActualizacion from " & gConsImagenes & "PersImagen PI Where cPersCod = '" & psCodPers & "'"
    Set R = New ADODB.Recordset
    R.Open sSql, ConexionActiva, adOpenKeyset, adLockOptimistic
'    If R.BOF And R.EOF Then
'        R.Close
'        Set RecuperaFirma = Nothing
'        Exit Function
'    End If
    Set RecuperaFirma = R
End Function

Public Sub RecuperaDocumentos(ByVal psCodPers As String)
Dim sSql As String
Dim R As ADODB.Recordset
    nNumDoc = 0
    
    ReDim Docs(nNumDoc)
    sSql = "Select PID.cPersIDnro,PID.cPersIDTpo,C.cConsDescripcion,C.nConsValor From " & gConsPersona & "PersID PID inner join Constante C ON convert(int,PID.cPersIDTpo) = C.nConsValor Where PID.cPersCod = '" & psCodPers & "' And  C.nConsCod = " & Trim(Str(gPersIdTipo)) & " And C.cConsDescripcion <> '" & Trim(Str(gPersIdTipo)) & "'"
    Set R = Conn.CargaRecordSet(sSql)
    Do While Not R.EOF
        nNumDoc = nNumDoc + 1
        ReDim Preserve Docs(nNumDoc)
        Docs(nNumDoc - 1).nTipoAct = 0
        Docs(nNumDoc - 1).sTipo = R!cConsDescripcion & Space(50) & R!nConsValor
        Docs(nNumDoc - 1).nTipo = CInt(R!cPersIDTpo)
        Docs(nNumDoc - 1).nTipoOrig = CInt(R!cPersIDTpo)
        Docs(nNumDoc - 1).sNumero = Trim(R!cPersIDNro)
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
End Sub
Private Sub RecuperaRelacionesdePersona(ByVal psPersCod As String)
Dim sSql As String
Dim R As New ADODB.Recordset
    nNumRel = 0
    ReDim Relac(nNumRel)
    sSql = "Select PR.bPersRelacAMP, PR.nPersRelacBenefPorc, PR.bPersRelacBenef, C3.cConsDescripcion cConsDescripcion3, C2.cConsDescripcion cConsDescripcion2,C2.nConsValor nConsValor2, C.cConsDescripcion,PR.cPersRelacPersCod,P.cPersNombre,PR.nPersRelac,PR.cMovNro "
    sSql = sSql & " from " & gConsPersona & "PersRelaciones PR Inner join Persona P ON PR.cPersRelacPersCod = P.cPersCod"
    sSql = sSql & " Inner join Constante C on PR.nPersRelac = C.nConsValor"
    sSql = sSql & " Left join Constante C2 on PR.bPersRelacBenef = C2.nConsValor"
    sSql = sSql & " Left join Constante C3 on PR.bPersRelacAMP = C3.nConsValor"
    sSql = sSql & " Where PR.cPersCod ='" & psPersCod & "' and C.nConsCod = " & gPersRelacion & " And C.nConsValor <> " & gPersRelacion & " and C2.nConsCod = " & gPersRelacBenef & " And C2.nConsValor <> " & gPersRelacBenef & " and C3.nConsCod = " & gPersRelacAmp & " And C3.nConsValor <> " & gPersRelacAmp
    Set R = Conn.CargaRecordSet(sSql)
    Do While Not R.EOF
        nNumRel = nNumRel + 1
        ReDim Preserve Relac(nNumRel)
        Relac(nNumRel - 1).sPersCodOrig = R!cPersRelacPersCod
        Relac(nNumRel - 1).sPersCod = R!cPersRelacPersCod
        Relac(nNumRel - 1).sApellNombres = Trim(R!cPersNombre)
        Relac(nNumRel - 1).sRelacion = R!cConsDescripcion & Space(50) & R!nPersRelac
        Relac(nNumRel - 1).sPersRelacBenef = R!cConsDescripcion2 & Space(50) & R!nConsValor2
        'Relac(nNumRel - 1).sPersRelacEstado = R!cConsDescripcion4 & Space(50) & Format(IIf(IsNull(R!bPersRelacEstado), 0, R!bPersRelacEstado), "#0")
        Relac(nNumRel - 1).nPersRelacBenefPorc = IIf(IsNull(R!nPersRelacBenefPorc), 0, R!nPersRelacBenefPorc)
        Relac(nNumRel - 1).sPersRelacAMP = R!cConsDescripcion3 & Space(50) & Format(IIf(IsNull(R!bPersRelacAMP), 0, R!bPersRelacAMP), "#0")
        Relac(nNumRel - 1).nTipoAct = 0
        Relac(nNumRel - 1).sMovAct = R!cMovNro
        R.MoveNext
    Loop
    R.Close
End Sub
Public Sub RecuperaFtesdeIngreso(ByVal psCodPers As String)
Dim sSql As String
Dim R As ADODB.Recordset
Dim bConexInterna As Boolean
    sSql = "Select P2.cPersNombre cRazonSoc, PFD.cPersCod,PFD.cPersFIPersCod,PFD.dPersFIEvaluacion,PFD.dPersFICaduca,PFD.cPersFICargo,PFD.dPersFIinicio,PFD.cPersFIMoneda,PFD.nPersFIIngresos,PFD.nPersFIIngresosFam,PFD.nPersFIIngresosOtros,PFD.nPersFIGastosFam,PFD.cPersFIComentario,"
    sSql = sSql & "PFI.nPersFIActivoDisp,PFI.nPersFICtasxCobrar,PFI.nPersFIInventarios,PFI.nPersFIActivosFijos,PFI.nPersFIProveedores,PFI.nPersFICredCMACT,PFI.nPersFICredOtros,PFI.nPersFIVentas,PFI.nPersFIPatrimonio,PFI.nPersFICostoVentas,PFI.nPersFIRecupCtasXCobrar,PFI.nPersFIEgresosOtros"
    sSql = sSql & " From Persona P inner join PersFteIngreso PFD ON P.cPersCod = PFD.cPersCod left join PersFIIndependiente PFI ON PFD.cPersCod = PFI.cPersCod And PFD.cPersFIPersCod = PFI.cPersFIPersCod And PFD.dPersFIEvaluacion = PFI.dPersFIEvaluacion  left join Persona P2 ON P2.cPersCod = PFD.cPersFIPersCod"
    sSql = sSql & " Where P.cPersCod = '" & psCodPers & "'"
    
    nNumFtes = 0
    bConexInterna = False
    ReDim FtesIngreso(nNumFtes)
    If Conn Is Nothing Then
        bConexInterna = True
        Set Conn = New DConecta
        Conn.AbreConexion
    End If
    Set R = Conn.CargaRecordSet(sSql)
    Do While Not R.EOF
        nNumFtes = nNumFtes + 1
        ReDim Preserve FtesIngreso(nNumFtes)
        FtesIngreso(nNumFtes - 1).dFecCaduc = CDate(Format(R!dPersFICaduca, "dd/mm/yyyy"))
        If IsNull(R!dPersFIEvaluacion) Then
            FtesIngreso(nNumFtes - 1).dFecEval = CDate("01/01/1950")
        Else
            FtesIngreso(nNumFtes - 1).dFecEval = Format(R!dPersFIEvaluacion, "dd/mm/yyyy hh:mm:ss")
            FtesIngreso(nNumFtes - 1).dFecEvalOrig = Format(R!dPersFIEvaluacion, "dd/mm/yyyy hh:mm:ss")
        End If
        
        FtesIngreso(nNumFtes - 1).dFteIngreso = Format(R!dPersFIInicio, "dd/mm/yyyy")
        FtesIngreso(nNumFtes - 1).nActFijo = Format(IIf(IsNull(R!nPersFIActivosFijos), 0#, R!nPersFIActivosFijos), "#0.00")
        FtesIngreso(nNumFtes - 1).nActivDisp = Format(IIf(IsNull(R!nPersFIActivoDisp), 0#, R!nPersFIActivoDisp), "#0.00")
        'FtesIngreso(nNumFtes - 1).nActivo
        'FtesIngreso(nNumFtes - 1).nActivoCircul
        FtesIngreso(nNumFtes - 1).nCompraMercad = Format(IIf(IsNull(R!nPersFICostoVentas), 0#, R!nPersFICostoVentas), "#0.00")
        FtesIngreso(nNumFtes - 1).nCreditosCmact = Format(IIf(IsNull(R!nPersFICredCMACT), 0#, R!nPersFICredCMACT), "#0.00")
        FtesIngreso(nNumFtes - 1).nCreditosOtros = Format(IIf(IsNull(R!nPersFICredOtros), 0#, R!nPersFICredOtros), "#0.00")
        FtesIngreso(nNumFtes - 1).nCtasxCob = Format(IIf(IsNull(R!nPersFICtasxCobrar), 0#, R!nPersFICtasxCobrar), "#0.00")
        FtesIngreso(nNumFtes - 1).nGastosFam = Format(IIf(IsNull(R!nPersFIGastosFam), 0#, R!nPersFIGastosFam), "#0.00")
        FtesIngreso(nNumFtes - 1).nIngFam = Format(IIf(IsNull(R!nPersFIIngresosFam), 0#, R!nPersFIIngresosFam), "#0.00")
        FtesIngreso(nNumFtes - 1).nIngOtros = IIf(IsNull(R!nPersFIIngresosOtros), 0#, R!nPersFIIngresosOtros)
        FtesIngreso(nNumFtes - 1).nIngresos = Format(IIf(IsNull(R!nPersFIIngresos), 0#, R!nPersFIIngresos), "#0.00")
        'FtesIngreso(nNumFtes - 1).nIngTotales
        FtesIngreso(nNumFtes - 1).nInventarios = Format(IIf(IsNull(R!nPersFIInventarios), 0#, R!nPersFIInventarios), "0.00")
        FtesIngreso(nNumFtes - 1).nOtrosEgresos = Format(IIf(IsNull(R!nPersFIEgresosOtros), 0#, R!nPersFIEgresosOtros), "0.00")
        'FtesIngreso(nNumFtes - 1).nPasivo
        FtesIngreso(nNumFtes - 1).nPasivoPatrimonio = Format(IIf(IsNull(R!nPersFIPatrimonio), 0#, R!nPersFIPatrimonio), "#0.00")
        FtesIngreso(nNumFtes - 1).nProveedores = Format(IIf(IsNull(R!nPersFIProveedores), 0#, R!nPersFIProveedores), "#0.00")
        FtesIngreso(nNumFtes - 1).nRecupCtasxCob = Format(IIf(IsNull(R!nPersFIRecupCtasxCobrar), 0#, R!nPersFIRecupCtasxCobrar), "#0.00")
        'FtesIngreso(nNumFtes - 1).nSaldDispFam
        FtesIngreso(nNumFtes - 1).nVentas = Format(IIf(IsNull(R!nPersFIVentas), 0#, R!nPersFIVentas), "#0.00")
        FtesIngreso(nNumFtes - 1).sCargo = R!cPersFICargo
        FtesIngreso(nNumFtes - 1).sComentarios = R!cPersFIComentario
        FtesIngreso(nNumFtes - 1).sFteIngreso = R!cPersFIPersCod
        FtesIngreso(nNumFtes - 1).sFteIngresoOrig = R!cPersFIPersCod
        FtesIngreso(nNumFtes - 1).sFteRazonSocial = Trim(R!cRazonSoc)
        FtesIngreso(nNumFtes - 1).sMoneda = Trim(Str(R!cPersFIMoneda))
        FtesIngreso(nNumFtes - 1).sTipo = IIf(IsNull(R!nPersFIActivoDisp), Trim(Str(gPersFteIngresoTipoDependiente)), Trim(Str(gPersFteIngresoTipoIndependiente)))
        FtesIngreso(nNumFtes - 1).nTipoAct = 0
        R.MoveNext
    Loop
    R.Close
    If bConexInterna Then
        Conn.CierraConexion
        Set Conn = Nothing
    End If
End Sub
Public Sub RecuperaPersona(ByVal psCodPers As String)
Dim sSql As String
Dim R As ADODB.Recordset
    On Error GoTo ErrRecuperaPersona
    Set Conn = New DConecta
    Conn.AbreConexion
    Set ConexionActiva = Conn.ConexionActiva
    
    '***** CARGA DATOS DE LA PERSONA ******
    sSql = "Select P.cPersEmail cEmail,P.cPersTelefono2,P.cPersCod,P.cPersNombre,P.dPersNacCreac,P.cPersDireccUbiGeo,P.cPersDireccDomicilio,P.cPersDireccCondicion,P.cPersTelefono,P.nPersPersoneria,P.cPersCIIU,P.cPersEstado,P.cUltimaActualizacion, "
    sSql = sSql & "PN.cPersNatSexo,PN.nPersNatEstCiv,PN.nPersNatHijos,PN.nPersPeso,PN.nPersTalla,PN.cPersTpoSangre,"
    sSql = sSql & "PJ.cPersJurTpo,PJ.cPersJurMagnitud,PJ.cPersJurSigla,PJ.nPersJurEmpleados, CI.cCIIUdescripcion"
    sSql = sSql & " from " & gConsPersona & "Persona P left join " & gConsPersona & "PersonaNat PN ON P.cPersCod=PN.cPersCod "
    sSql = sSql & " Left join " & gConsPersona & "PersonaJur PJ ON P.cPersCod = PJ.cPersCod "
    sSql = sSql & " Left join CIIU CI ON P.cPersCIIU = CI.cCIIUcod "
    sSql = sSql & " where P.cPersCod = '" & psCodPers & "'"
    
    Set R = Conn.CargaRecordSet(sSql)
    If Not R.BOF And Not R.EOF Then
        sPersCod = R!cPersCod
        sNombreCompleto = Trim(R!cPersNombre)
        sApePat = BuscaNombre(R!cPersNombre, BusqApellidoPaterno)
        sApeMat = BuscaNombre(R!cPersNombre, BusqApellidoMaterno)
        sNombres = BuscaNombre(R!cPersNombre, BusqNombres)
        sApeCas = BuscaNombre(R!cPersNombre, BusqApellidoCasada)
        sSexo = Trim(IIf(IsNull(R!cPersnatSexo), "", R!cPersnatSexo))
        sEstadoCivil = Trim(IIf(IsNull(R!nPersnatEstCiv), "", R!nPersnatEstCiv))
        nHijos = IIf(IsNull(R!nPersNatHijos), 0, R!nPersNatHijos)
        dFechaNac = CDate(Format(R!dPersNacCreac, "dd/mm/yyyy"))
        sUbicGeografica = R!cPersDireccUbiGeo
        sDomicilio = R!cPersDireccDomicilio
        sCondicionDomic = R!cPersDireccCondicion
        sTelefonos = IIf(IsNull(R!cPersTelefono), "", R!cPersTelefono)
        sTelefonos2 = IIf(IsNull(R!cPersTelefono2), "", R!cPersTelefono2)
        sEmail = IIf(IsNull(R!cEmail), "", R!cEmail)
        sCiiu = IIf(IsNull(R!cPersCiiu), "", R!cPersCiiu)
        sCiiuDescripcion = IIf(IsNull(R!cCIIUdescripcion), "", R!cCIIUdescripcion)
        sEstado = R!cPersEstado
        nTalla = IIf(IsNull(R!nPersTalla), 0#, R!nPersTalla)
        nPeso = IIf(IsNull(R!nPersPeso), 0#, R!nPersPeso)
        sTipoSangre = IIf(IsNull(R!cPersTpoSangre), "", R!cPersTpoSangre)
        nTipoPersona = R!nPersPersoneria
        'Datos de Persona Juridica
        'sRazonSocial
        sSiglas = IIf(IsNull(R!cPersJurSigla), "", R!cPersJurSigla)
        sTipoPersJur = IIf(IsNull(R!cPersJurMagnitud), "", R!cPersJurTpo)
        nNumEmplead = IIf(IsNull(R!nPersJurEmpleados), 0, R!nPersJurEmpleados)
        sMagnitudEmp = IIf(IsNull(R!cPersJurMagnitud), "", R!cPersJurMagnitud)
        nTipoAct = 0
    Else
        sPersCod = ""
        R.Close
        Set R = Nothing
        Conn.CierraConexion
        Set Conn = Nothing
        Exit Sub
    End If
    R.Close
    
    'Recupera la Firma de la Persona de la Base de Datos
    Set RFirma = RecuperaFirma(psCodPers)
    'Carga los Diferentes Tipos de Documentos que Posee la Persona
    Call RecuperaDocumentos(psCodPers)
    'Carga las diferentes Personas que Tienen Relacion con la persona actual
    Call RecuperaRelacionesdePersona(psCodPers)
    'Carga las Fuentes de Ingreso de la Persona
    Call RecuperaFtesdeIngreso(psCodPers)
    Exit Sub
    
ErrRecuperaPersona:
    Err.Raise Err.Number, "ReucperaPersona", Err.Description
    
End Sub

Public Property Get NombreCompleto() As String
    NombreCompleto = sNombreCompleto
End Property

Public Property Let NombreCompleto(ByVal vNewValue As String)
    sNombreCompleto = vNewValue
End Property

Public Property Get PersCodigo() As String
    PersCodigo = sPersCod
End Property

Public Property Let PersCodigo(ByVal vNewValue As String)
    sPersCod = vNewValue
End Property

Public Property Get ApellidoPaterno() As String
    ApellidoPaterno = sApePat
End Property

Public Property Let ApellidoPaterno(ByVal vNewValue As String)
    sApePat = vNewValue
End Property

Public Property Get ApellidoMaterno() As String
    ApellidoMaterno = sApeMat
End Property

Public Property Let ApellidoMaterno(ByVal vNewValue As String)
    sApeMat = vNewValue
End Property

Public Property Get CampoActualizacion() As String
    CampoActualizacion = sActualiza
End Property

Public Property Let CampoActualizacion(ByVal vNewValue As String)
    sActualiza = vNewValue
End Property

Public Property Get ApellidoCasada() As String
    ApellidoCasada = sApeCas
End Property

Public Property Let ApellidoCasada(ByVal vNewValue As String)
    sApeCas = vNewValue
End Property

Public Property Get Nombres() As String
    Nombres = sNombres
End Property

Public Property Let Nombres(ByVal vNewValue As String)
    sNombres = vNewValue
End Property

Public Property Get Sexo() As String
    Sexo = sSexo
End Property

Public Property Let Sexo(ByVal vNewValue As String)
    sSexo = vNewValue
End Property

Public Property Get EstadoCivil() As String
    EstadoCivil = sEstadoCivil
End Property

Public Property Let EstadoCivil(ByVal vNewValue As String)
    sEstadoCivil = vNewValue
End Property

Public Property Get Hijos() As Integer
    Hijos = nHijos
End Property

Public Property Let Hijos(ByVal vNewValue As Integer)
    nHijos = vNewValue
End Property

Public Property Get FechaNacimiento() As Date
    FechaNacimiento = dFechaNac
End Property

Public Property Let FechaNacimiento(ByVal vNewValue As Date)
    dFechaNac = vNewValue
End Property

Public Property Get UbicacionGeografica() As String
    UbicacionGeografica = sUbicGeografica
End Property

Public Property Let UbicacionGeografica(ByVal vNewValue As String)
    sUbicGeografica = vNewValue
End Property

Public Property Get Talla() As Double
    Talla = nTalla
End Property

Public Property Let Talla(ByVal vNewValue As Double)
    nTalla = vNewValue
End Property

Public Property Get Peso() As Double
    Peso = nPeso
End Property

Public Property Let Peso(ByVal vNewValue As Double)
    nPeso = vNewValue
End Property

Public Property Let TipoSangre(ByVal vNewValue As String)
    sTipoSangre = vNewValue
End Property

Public Property Get TipoSangre() As String
    TipoSangre = sTipoSangre
End Property

Public Property Get Domicilio() As String
    Domicilio = sDomicilio
End Property

Public Property Let Domicilio(ByVal vNewValue As String)
    sDomicilio = vNewValue
End Property

Public Property Get CondicionDomicilio() As String
    CondicionDomicilio = sCondicionDomic
End Property

Public Property Let CondicionDomicilio(ByVal vNewValue As String)
    sCondicionDomic = vNewValue
End Property

Public Property Get Telefonos() As String
    Telefonos = sTelefonos
End Property

Public Property Let Telefonos2(ByVal vNewValue As String)
    sTelefonos2 = vNewValue
End Property

Public Property Get Telefonos2() As String
    Telefonos2 = sTelefonos2
End Property

Public Property Let Email(ByVal vNewValue As String)
    sEmail = vNewValue
End Property

Public Property Get Email() As String
    Email = sEmail
End Property

Public Property Let PersonaRelaInst(ByVal vNewValue As String)
    sPersRelInst = vNewValue
End Property

Public Property Get PersonaRelaInst() As String
    PersonaRelaInst = sPersRelInst
End Property

Public Property Let Telefonos(ByVal vNewValue As String)
    sTelefonos = vNewValue
End Property

Public Function ObtenerDNI() As String
Dim i As Integer
    ObtenerDNI = ""
    For i = 0 To nNumDoc - 1
        If Docs(i).nTipo = gPersIdDNI Then
            ObtenerDNI = Docs(i).sNumero
            Exit Function
        End If
    Next i
End Function

Public Function ObtenerRUC() As String
Dim i As Integer
    ObtenerRUC = ""
    For i = 0 To nNumDoc - 1
        If Docs(i).nTipo = gPersIdRUC Then
            ObtenerRUC = Docs(i).sNumero
            Exit Function
        End If
    Next i
End Function

Public Property Get CIIUDescripcion() As String
    CIIUDescripcion = sCiiuDescripcion
End Property

Public Property Let CIIUDescripcion(ByVal vNewValue As String)
    sCiiuDescripcion = vNewValue
End Property

Public Property Get CIIU() As String
    CIIU = sCiiu
End Property

Public Property Let CIIU(ByVal vNewValue As String)
    sCiiu = vNewValue
End Property

Public Property Get Estado() As String
    Estado = sEstado
End Property

Public Property Let Estado(ByVal vNewValue As String)
    sEstado = vNewValue
End Property

Public Property Get RazonSocial() As String
    RazonSocial = sRazonSocial
End Property

Public Property Let RazonSocial(ByVal vNewValue As String)
    sRazonSocial = vNewValue
End Property

Public Property Get Siglas() As String
    Siglas = sSiglas
End Property

Public Property Let Siglas(ByVal vNewValue As String)
    sSiglas = vNewValue
End Property

Public Property Get TipoPersonaJur() As String
    TipoPersonaJur = sTipoPersJur
End Property

Public Property Let TipoPersonaJur(ByVal vNewValue As String)
    sTipoPersJur = vNewValue
End Property

Public Property Get NumerosEmpleados() As String
    NumerosEmpleados = nNumEmplead
End Property

Public Property Let NumerosEmpleados(ByVal vNewValue As String)
    nNumEmplead = vNewValue
End Property

Public Property Get MagnitudEmpresarial() As String
    MagnitudEmpresarial = sMagnitudEmp
End Property

Public Property Let MagnitudEmpresarial(ByVal vNewValue As String)
    sMagnitudEmp = vNewValue
End Property

Public Property Get CodigoSbs() As String
    CodigoSbs = sPersCodSbs
End Property

Public Property Let CodigoSbs(ByVal vNewValue As String)
    sPersCodSbs = vNewValue
End Property

Public Property Get TipoActualizacion() As TPersCambios
    TipoActualizacion = nTipoAct
End Property

Public Property Let TipoActualizacion(ByVal vNewValue As TPersCambios)
    nTipoAct = vNewValue
End Property

Public Property Get Personeria() As Integer
    Personeria = nTipoPersona
End Property

Public Property Let Personeria(ByVal vNewValue As Integer)
    nTipoPersona = vNewValue
End Property

Public Property Get NumeroDocumentos() As Integer
    NumeroDocumentos = nNumDoc
End Property

Public Property Let NumeroDocumentos(ByVal vNewValue As Integer)
    nNumDoc = vNewValue
End Property

Public Function ObtenerTipoDoc(ByVal pnPunt As Integer) As String
    ObtenerTipoDoc = Docs(pnPunt).sTipo
End Function
Public Function ObtenerNumeroDoc(ByVal pnPunt As Integer) As String
    ObtenerNumeroDoc = Docs(pnPunt).sNumero
End Function
Public Sub ActualizarDocsTipoAct(ByVal pCambio As TPersCambios, ByVal pnPunt As Integer)
     Docs(pnPunt).nTipoAct = pCambio
End Sub

Public Sub ActualizaDocsTipo(ByVal psTipoDoc As String, ByVal pnPunt As Integer)
    Docs(pnPunt).sTipo = psTipoDoc
    Docs(pnPunt).nTipo = CInt(Trim(Right(psTipoDoc, 15)))
End Sub

Public Sub ActualizaDocsNumero(ByVal psDocNumero As String, ByVal pnPunt As Integer)
    Docs(pnPunt).sNumero = psDocNumero
End Sub

Public Property Get NumeroRelacPers() As Integer
    NumeroRelacPers = nNumRel
End Property

Public Property Let NumeroRelacPers(ByVal vNewValue As Integer)
    nNumRel = vNewValue
End Property

Public Function ObtenerRelaPersCodigo(ByVal pnPunt As Integer) As String
    ObtenerRelaPersCodigo = Relac(pnPunt).sPersCod
End Function

Public Function ObtenerRelaPersTipoAct(ByVal pnPunt As Integer) As TPersCambios
    ObtenerRelaPersTipoAct = Relac(pnPunt).nTipoAct
End Function

Public Function ObtenerRelaPersNombres(ByVal pnPunt As Integer) As String
    ObtenerRelaPersNombres = Relac(pnPunt).sApellNombres
End Function

Public Sub ActualizarRelaPersAMP(ByVal psPersRelacAMP As String, ByVal pnPunt As Integer)
    Relac(pnPunt).sPersRelacAMP = psPersRelacAMP
End Sub

Public Function ObtenerDocTipoAct(ByVal pnPunt As Integer) As TPersCambios
     ObtenerDocTipoAct = Docs(pnPunt).nTipoAct
End Function

Public Function ObtenerRelaPersAMP(ByVal pnPunt As Integer) As String
    ObtenerRelaPersAMP = Relac(pnPunt).sPersRelacAMP
End Function

Public Sub ActualizarRelaPersBenefPorc(ByVal pnPersRelacBenefPorc As Integer, ByVal pnPunt As Integer)
    Relac(pnPunt).nPersRelacBenefPorc = pnPersRelacBenefPorc
End Sub

Public Function ObtenerRelaPersBenefPorc(ByVal pnPunt As Integer) As Integer
    ObtenerRelaPersBenefPorc = Relac(pnPunt).nPersRelacBenefPorc
End Function

Public Sub ActualizarRelaPersBenef(ByVal psPersRelacBenef As String, ByVal pnPunt As Integer)
    Relac(pnPunt).sPersRelacBenef = psPersRelacBenef
End Sub

Public Function ObtenerRelaPersBenef(ByVal pnPunt As Integer) As String
    ObtenerRelaPersBenef = Relac(pnPunt).sPersRelacBenef
End Function

Public Function ObtenerRelaPersRelacion(ByVal pnPunt As Integer) As String
    ObtenerRelaPersRelacion = Relac(pnPunt).sRelacion
End Function

Public Sub ActualizarRelacPersTipoAct(ByVal pCambio As TPersCambios, ByVal pnPunt As Integer)
     Relac(pnPunt).nTipoAct = pCambio
End Sub

Public Property Get NumeroFtesIngreso() As Integer
    NumeroFtesIngreso = nNumFtes
End Property

Public Property Let NumeroFtesIngreso(ByVal vNewValue As Integer)
    nNumFtes = vNewValue
End Property
Public Function ObtenerFteIngTipoAct(ByVal pnPunt As Integer) As TPersCambios
    ObtenerFteIngTipoAct = FtesIngreso(pnPunt).nTipoAct
End Function

Public Function ObtenerFteIngTipo(ByVal pnPunt As Integer) As String
    ObtenerFteIngTipo = FtesIngreso(pnPunt).sTipo
End Function

Public Function ObtenerFteIngFuente(ByVal pnPunt As Integer) As String
    ObtenerFteIngFuente = FtesIngreso(pnPunt).sFteIngreso
End Function

Public Function ObtenerFteIngFecEval(ByVal pnPunt As Integer) As Date
    ObtenerFteIngFecEval = FtesIngreso(pnPunt).dFecEval
End Function

Public Function ObtenerFteIngFecCaducac(ByVal pnPunt As Integer) As Date
    ObtenerFteIngFecCaducac = FtesIngreso(pnPunt).dFecCaduc
End Function

Public Function ObtenerFteIngCargo(ByVal pnPunt As Integer) As String
    ObtenerFteIngCargo = FtesIngreso(pnPunt).sCargo
End Function

Public Function ObtenerFteIngInicioFuente(ByVal pnPunt As Integer) As Date
    ObtenerFteIngInicioFuente = FtesIngreso(pnPunt).dFteIngreso
End Function

Public Function ObtenerFteIngMoneda(ByVal pnPunt As Integer) As Double
    ObtenerFteIngMoneda = FtesIngreso(pnPunt).sMoneda
End Function

Public Function ObtenerFteIngIngresos(ByVal pnPunt As Integer) As Double
    ObtenerFteIngIngresos = FtesIngreso(pnPunt).nIngresos
End Function

Public Function ObtenerFteIngIngresoFam(ByVal pnPunt As Integer) As Double
    ObtenerFteIngIngresoFam = FtesIngreso(pnPunt).nIngFam
End Function

Public Function ObtenerFteIngIngresoOtros(ByVal pnPunt As Integer) As Double
    ObtenerFteIngIngresoOtros = FtesIngreso(pnPunt).nIngOtros
End Function

Public Function ObtenerFteIngIngresoTotales(ByVal pnPunt As Integer) As Double
    ObtenerFteIngIngresoTotales = FtesIngreso(pnPunt).nIngTotales
End Function

Public Function ObtenerFteIngGastoFam(ByVal pnPunt As Integer) As Double
    ObtenerFteIngGastoFam = FtesIngreso(pnPunt).nGastosFam
End Function

Public Function ObtenerFteIngSaldoDispFam(ByVal pnPunt As Integer) As Double
    ObtenerFteIngSaldoDispFam = FtesIngreso(pnPunt).nSaldDispFam
End Function

Public Function ObtenerFteIngActivo(ByVal pnPunt As Integer) As Double
    ObtenerFteIngActivo = FtesIngreso(pnPunt).nActivo
End Function

Public Function ObtenerFteIngActivoCircul(ByVal pnPunt As Integer) As Double
    ObtenerFteIngActivoCircul = FtesIngreso(pnPunt).nActivoCircul
End Function

Public Function ObtenerFteIngActivoDisp(ByVal pnPunt As Integer) As Double
    ObtenerFteIngActivoDisp = FtesIngreso(pnPunt).nActivDisp
End Function

Public Function ObtenerFteIngCtasxCob(ByVal pnPunt As Integer) As Double
    ObtenerFteIngCtasxCob = FtesIngreso(pnPunt).nCtasxCob
End Function

Public Function ObtenerFteIngInventario(ByVal pnPunt As Integer) As Double
    ObtenerFteIngInventario = FtesIngreso(pnPunt).nInventarios
End Function

Public Function ObtenerFteIngActivoFijo(ByVal pnPunt As Integer) As Double
    ObtenerFteIngActivoFijo = FtesIngreso(pnPunt).nActFijo
End Function

Public Function ObtenerFteIngPasivoPatrim(ByVal pnPunt As Integer) As Double
    ObtenerFteIngPasivoPatrim = FtesIngreso(pnPunt).nPasivoPatrimonio
End Function

Public Function ObtenerFteIngPasivo(ByVal pnPunt As Integer) As Double
    ObtenerFteIngPasivo = FtesIngreso(pnPunt).nPasivo
End Function

Public Function ObtenerFteIngProveedores(ByVal pnPunt As Integer) As Double
    ObtenerFteIngProveedores = FtesIngreso(pnPunt).nProveedores
End Function

Public Function ObtenerFteIngOtrosCreditos(ByVal pnPunt As Integer) As Double
    ObtenerFteIngOtrosCreditos = FtesIngreso(pnPunt).nCreditosOtros
End Function

Public Function ObtenerFteIngCreditosCmact(ByVal pnPunt As Integer) As Double
    ObtenerFteIngCreditosCmact = FtesIngreso(pnPunt).nCreditosCmact
End Function

Public Function ObtenerFteIngIngresosTotales(ByVal pnPunt As Integer) As Double
    If CInt(FtesIngreso(pnPunt).sTipo) = gPersFteIngresoTipoDependiente Then
        ObtenerFteIngIngresosTotales = FtesIngreso(pnPunt).nIngresos + FtesIngreso(pnPunt).nIngFam + FtesIngreso(pnPunt).nIngOtros - FtesIngreso(pnPunt).nGastosFam
    Else
        ObtenerFteIngIngresosTotales = FtesIngreso(pnPunt).nVentas + FtesIngreso(pnPunt).nRecupCtasxCob + FtesIngreso(pnPunt).nIngFam - (FtesIngreso(pnPunt).nCompraMercad + FtesIngreso(pnPunt).nOtrosEgresos + FtesIngreso(pnPunt).nGastosFam)
    End If
End Function

Public Function ObtenerFteIngVentas(ByVal pnPunt As Integer) As Double
    ObtenerFteIngVentas = FtesIngreso(pnPunt).nVentas
End Function

Public Function ObtenerFteIngRecupCtasxCobrar(ByVal pnPunt As Integer) As Double
    ObtenerFteIngRecupCtasxCobrar = FtesIngreso(pnPunt).nRecupCtasxCob
End Function

Public Function ObtenerFteIngComprasMercad(ByVal pnPunt As Integer) As Double
    ObtenerFteIngComprasMercad = FtesIngreso(pnPunt).nCompraMercad
End Function

Public Function ObtenerFteIngOtrosEgresos(ByVal pnPunt As Integer) As Double
    ObtenerFteIngOtrosEgresos = FtesIngreso(pnPunt).nOtrosEgresos
End Function

Public Function ObtenerFteIngComentarios(ByVal pnPunt As Integer) As String
    ObtenerFteIngComentarios = FtesIngreso(pnPunt).sComentarios
End Function

Public Function ObtenerFteIngRazonSocial(ByVal pnPunt As Integer) As String
    ObtenerFteIngRazonSocial = FtesIngreso(pnPunt).sFteRazonSocial
End Function

Public Sub ActualizarFteIngTipoFuente(ByVal psTipoFte As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).sTipo = psTipoFte
End Sub

Public Sub ActualizarFteIngComentarios(ByVal psComentarios As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).sComentarios = psComentarios
End Sub

Public Sub ActualizarFteIngOtrosEgresos(ByVal pnOtrosEgresos As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nOtrosEgresos = pnOtrosEgresos
End Sub

Public Sub ActualizarFteIngCompraMercad(ByVal pnCompraMercad As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nCompraMercad = pnCompraMercad
End Sub

Public Sub ActualizarFteIngRecupCtasxCob(ByVal pnRecupCtasxCob As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nRecupCtasxCob = pnRecupCtasxCob
End Sub

Public Sub ActualizarFteIngVentas(ByVal pnVentas As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nVentas = pnVentas
End Sub

Public Sub ActualizarFteIngCreditosCmact(ByVal pnCreditosCmact As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nCreditosCmact = pnCreditosCmact
End Sub

Public Sub ActualizarFteIngCreditosOtros(ByVal pnCreditosOtros As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nCreditosOtros = pnCreditosOtros
End Sub

Public Sub ActualizarFteIngProveedores(ByVal pnProveedores As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nProveedores = pnProveedores
End Sub

Public Sub ActualizarFteIngPasivo(ByVal pnPasivo As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nPasivo = pnPasivo
End Sub

Public Sub ActualizarFteIngPasivoPatrimonio(ByVal pnPasivoPatrimonio As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nPasivoPatrimonio = pnPasivoPatrimonio
End Sub

Public Sub ActualizarFteIngActFijo(ByVal pnActFijo As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nActFijo = pnActFijo
End Sub

Public Sub ActualizarFteIngInventarios(ByVal pnInventarios As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nInventarios = pnInventarios
End Sub

Public Sub ActualizarFteIngCtasxCob(ByVal pnCtasxCob As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nCtasxCob = pnCtasxCob
End Sub

Public Sub ActualizarFteIngActivDisp(ByVal pnActivDisp As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nActivDisp = pnActivDisp
End Sub

Public Sub ActualizarFteIngActivoCircul(ByVal pnActivoCircul As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nActivoCircul = pnActivoCircul
End Sub

Public Sub ActualizarFteIngActivo(ByVal pnActivo As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nActivo = pnActivo
End Sub

Public Sub ActualizarFteIngSaldDispFam(ByVal pnSaldDispFam As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nSaldDispFam = pnSaldDispFam
End Sub

Public Sub ActualizarFteIngGastosFam(ByVal pnGastosFam As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nGastosFam = pnGastosFam
End Sub

Public Sub ActualizarFteIngIngTotales(ByVal pnIngTotales As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nIngTotales = pnIngTotales
End Sub

Public Sub ActualizarFteIngIngOtros(ByVal pnIngOtros As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nIngOtros = pnIngOtros
End Sub

Public Sub ActualizarFteIngIngFam(ByVal pnIngFam As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nIngFam = pnIngFam
End Sub

Public Sub ActualizarFteIngIngresos(ByVal pnIngresos As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nIngresos = pnIngresos
End Sub

Public Sub ActualizarFteIngMoneda(ByVal psMoneda As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).sMoneda = psMoneda
End Sub

Public Sub ActualizarFteIngFechaInicioFuente(ByVal pdFteIngreso As Date, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).dFteIngreso = pdFteIngreso
End Sub

Public Sub ActualizarFteIngCargo(ByVal psCargo As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).sCargo = psCargo
End Sub

Public Sub ActualizarFteIngFecCaduc(ByVal pdFecCaduc As Date, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).dFecCaduc = pdFecCaduc
End Sub

Public Sub ActualizarFteIngFecEval(ByVal pdFecEval As Date, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).dFecEval = pdFecEval
End Sub

Public Sub ActualizarFteIngFuenteIngreso(ByVal psFuenteIngreso As String, ByVal psRazonSocial As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).sFteIngreso = psFuenteIngreso
     FtesIngreso(pnPunt).sFteRazonSocial = psRazonSocial
End Sub

Public Sub ActualizarFteIngTipoAct(ByVal pCambio As TPersCambios, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nTipoAct = pCambio
End Sub

Public Sub ActualizaPersRelaPersona(ByVal psNombres As String, ByVal psPersRelaPersCod As String, ByVal pnPunt As Integer)
    Relac(pnPunt).sApellNombres = psNombres
    Relac(pnPunt).sPersCod = psPersRelaPersCod
End Sub

Public Sub ActualizaPersRelaRelacion(ByVal psRelacion As String, ByVal pnPunt As Integer)
    Relac(pnPunt).sRelacion = psRelacion
End Sub

Public Sub ActualizaPersRelaEstado(ByVal psEstado As String, ByVal pnPunt As Integer)
    Relac(pnPunt).sPersRelacEstado = psEstado
End Sub

Public Function ObtenerPersrelaEstado(ByVal pnPunt As Integer) As String
    ObtenerPersrelaEstado = Relac(pnPunt).sPersRelacEstado
End Function


