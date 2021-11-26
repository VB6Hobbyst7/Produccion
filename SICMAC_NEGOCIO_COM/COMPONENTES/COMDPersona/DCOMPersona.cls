VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************
'**** Clase que Maneja los Datos de la persona  ****
'***************************************************
Option Explicit
'En caso de error
Dim ofun As New COMFunciones.FCOMImpresion
Dim oGen As New COMDConstSistema.DCOMGeneral
Dim ocad As New COMFunciones.FCOMCadenas

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
Private Conn As COMConecta.DCOMConecta
Public sCodage As String

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
Private nValComDomicilio As Currency
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
Private nPersRelInst As Integer
Public dFechaHoy As Date

Public RFirma As ADODB.Recordset

'Datos Brechas
Private nResidencia As Integer
Private cNacionalidad As String

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

' Tipo de dato para almacenar las referencias comerciales de la persona
Private Type TRefCom
    nNumRefCom As Integer
    nRefComCod As Integer
    sRefComNom  As String
    sRefComTipo As String
    sRefComFono As String
    nTipoAct As TPersCambios
End Type
Private RefCom() As TRefCom
Private nNumRefCom As Integer
Private nMaxRefCom As Integer

'Tipo de Dato para almacenar las Referencias Bancarias de la persona
Private Type TRefBan
    nNumRefBan As Integer
    sRefBanCodIFOrg As String
    sRefBanCodIF  As String
    sRefBanNombre As String
    sRefBanNumCta As String
    sRefBanNumTar As String
    nRefBanLinCred As Currency
    nTipoAct As TPersCambios
End Type
Private RefBan() As TRefBan
Private nNumRefBan As Integer

' Tipo de dato para almacenar las referencias comerciales de la persona
Private Type TPatVeh
    nPatVehCod As Integer
    sPatVehMarca As String
    sPatVehFecFab As String
    nPatVehValCom As Currency
    sPatVehCondicion As String
    nTipoAct As TPersCambios
End Type
Private PatVeh() As TPatVeh
Private nNumPatVeh As Integer
Private nMaxPatVeh As Integer

'Tipo de Datos para Almacenar las Ftes de Ingreso
Private Type TFteIngDep
    cNumFte As String
    dFecEval As Date
    dFecCad As Date
    nIngresos As Double
    nIngFam As Double
    nIngCony As Double
    nIngOtros As Double
    nIngTotales As Double
    nGastosFam As Double
    nSaldDispFam As Double
    nTipoAct As TPersCambios
End Type


Private Type TFteCostoProd
    cNumFuente As String
    dFecEval As Date
    nTpoCultivo As Integer
    sTpoCultivo As String
    nMaquinaria As Double
    nJornales As Double
    nInsumos As Double
    nPesticidas As Double
    nOtros As Double
    nHectareas As Double
    '-------------------------
    nSiembra As Integer
    nMantenimiento As Integer
    nCosecha As Integer
    nDesAgricola As Integer
    nPlanOtros As Integer
    '------------------------
    nProduccion As Double
    nUniProd As Integer
    sUniProd As String
    nPreUni As Double
    nTipoAct As TPersCambios
 End Type


Private Type TFteIngInd
    cNumFte As String
    dFecEval As Date
    dFecCad As Date
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
    sMovNro As String
    nPersIngFam As Double
    nPersEgrFam As Double
    nTipoAct As TPersCambios
End Type

Private Type TFteIng
    cNumFte As String
    bCostoProd As Integer
    nTipo As Integer
    sFteIngresoOrig As String
    sFteIngreso As String
    dFecCaduc As Date
    sCargo As String
    dFteIngreso As Date
    sMoneda As String
    sComentarios As String
    sRazSocDirecc As String
    sRazSocUbiGeo As String
    sRazSocTelef As String
    sRazSocDescrip As String
    nTipoAct As TPersCambios
    FtesIngresoDep() As TFteIngDep
    FtesIngresoInd() As TFteIngInd
    nNumFtesDep As Integer
    nNumFtesIndep As Integer
    CostoProd() As TFteCostoProd
    nNumCostosProd As Integer
    nIngTotales As Double
    nSaldDispFam As Double
    nActivo As Double
    nActivoCircul As Double
    nPasivo As Double
    nPasivoPatrimonio As Double
    
    'Datos de la Fte como persona ------------
    cFtePersRuc As String
    dFtePersFecInicioAct As Date
    cFtePersTipoPersJur As String
    cFtePersTelefono As String
    cFtePersCIUU As String
    cFtePersCondicionDomic As String
    cFtePersMagnitudEmp As String
    nFtePersNroEmpleados As Double
    cFtePersDireccion As String
    cFtePersDpto As String
    cFtePersProv As String
    cFtePersDist As String
    cFtePersZona As String
    '----------------------------------------
End Type

Private FtesIngreso() As TFteIng

Private FtesIngresoTmp As TFteIng
Private nNumFtes As Integer
Dim ConexDitrib As ADODB.Connection
Dim bPersDistrib As Boolean
Dim nNumFtesDep As Integer
Dim nNumFtesIndep As Integer

Public Function GenerarImpresionFteIngresoDependiente(ByVal pnIndex As Integer, _
    ByVal psNomAge As String, ByVal pdFecha As Date, ByVal pbCostoProd As Boolean, _
    Optional ByVal psNomCmac As String = "") As String

Dim sCadImp As String
Dim i As Integer
Dim J As Integer
Dim lnNegritaON As String
Dim lnNegritaOFF As String
    
    
Dim MatDatosFte As Variant
Dim MatDatosComFte As Variant
Dim nMontoTemp As Double

    
    lnNegritaON = Chr$(27) & Chr$(69)
    lnNegritaOFF = Chr$(27) + Chr$(70)

    sCadImp = ""
    sCadImp = sCadImp & ofun.ImpreFormat(psNomCmac, 35, 35) & " - " & psNomAge
    sCadImp = sCadImp & Space(25) & Format(oGen.FechaHora(pdFecha), "dd/mm/yyyy hh:mm:ss") & Chr$(10) & Chr$(10) & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "Cliente : " & sPersCod & Space(2) & ocad.PstaNombre(sNombreCompleto) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    'sCadImp = sCadImp & "Fte Ingreso : " & FtesIngreso(pnIndex).sRazSocDescrip & Chr$(10) & Chr$(10) & Chr$(10)
    
    sCadImp = sCadImp & "Fte Ingreso : " & FtesIngreso(pnIndex).sRazSocDescrip & Chr$(10) & Chr$(10)
    
'----------------------------------------------------------------------------------------------------
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("RAZON SOCIAL/NOMBRE COMERCIAL", 35, 0) & ofun.ImpreFormat("RUC", 14, 2) & ofun.ImpreFormat("INICIO ACT.", 12, 0) & ofun.ImpreFormat("PERS. JURIDICA", 24, 2) & ofun.ImpreFormat("TELEFONO", 10, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).sRazSocDescrip, 35, 0) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersRuc, 14, 2) & ofun.ImpreFormat(Format(FtesIngreso(pnIndex).dFtePersFecInicioAct, "dd/mm/yyyy"), 12, 0) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersTipoPersJur, 24, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersTelefono, 10, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("ACTIVIDAD(CIUU)", 35, 0) & ofun.ImpreFormat("COND. LOCAL", 14, 2) & ofun.ImpreFormat("MAGNITUD EMP.", 35, 0) & ofun.ImpreFormat("Nª EMPLEADOS", 14, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersCIUU, 35, 0) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersCondicionDomic, 14, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersMagnitudEmp, 35, 0) & ofun.ImpreFormat(Round(FtesIngreso(pnIndex).nFtePersNroEmpleados), 14, 0) & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("DIRECCION(Avenida, Calle, Jiron)", 35, 0) & ofun.ImpreFormat("DEPARTAMENTO", 12, 2) & ofun.ImpreFormat("PROVINCIA", 12, 2) & ofun.ImpreFormat("DISTRITO", 16, 2) & ofun.ImpreFormat("ZONA", 17, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersDireccion, 35, 0) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersDpto, 12, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersProv, 12, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersDist, 16, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersZona, 17, 2) & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10) & Chr$(10)
'----------------------------------------------------------------------------------------------------
    
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & Space(20) & "EVOLUCION ECONOMICA FINANCIERA DEPENDIENTE" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    sCadImp = sCadImp & Space(23)
    
    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
        nMontoTemp = FtesIngreso(pnIndex).FtesIngresoDep(i).nIngresos
        If FtesIngreso(pnIndex).FtesIngresoDep(i).dFecEval <> 0 Then
            sCadImp = sCadImp & Format(FtesIngreso(pnIndex).FtesIngresoDep(i).dFecEval, "dd/mm/yyyy") & Space(3)
        Else
            sCadImp = sCadImp & Space(5) & "(-)" & Space(3)
        End If
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            sCadImp = sCadImp & "Variacion"
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Ingreso Familiar", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngFam, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesDep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngFam = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 5, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(CDbl(Format(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngFam / FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngFam, "#0.00")), 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    
    sCadImp = sCadImp & ofun.ImpreFormat("Ingreso Conyugue", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngCony, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesDep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngCony = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 5, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(CDbl(Format(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngCony / FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngCony, "#0.00")), 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    
    sCadImp = sCadImp & ofun.ImpreFormat("Otros Ingresos", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngOtros, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesDep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngOtros = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 5, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(CDbl(Format(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngOtros / FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngOtros, "#0.00")), 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    
    sCadImp = sCadImp & ofun.ImpreFormat("Gastos Familiares", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(i).nGastosFam, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesDep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nGastosFam = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 5, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(CDbl(Format(FtesIngreso(pnIndex).FtesIngresoDep(i).nGastosFam / FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nGastosFam, "#0.00")), 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("Saldo", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngCony + _
                                FtesIngreso(pnIndex).FtesIngresoDep(i).nIngFam + _
                                FtesIngreso(pnIndex).FtesIngresoDep(i).nIngOtros - FtesIngreso(pnIndex).FtesIngresoDep(i).nGastosFam, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesDep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngCony + _
                                FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngFam + _
                                FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngOtros - FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nGastosFam = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 5, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(CDbl(Format((FtesIngreso(pnIndex).FtesIngresoDep(i).nIngCony + _
                                FtesIngreso(pnIndex).FtesIngresoDep(i).nIngFam + _
                                FtesIngreso(pnIndex).FtesIngresoDep(i).nIngOtros - FtesIngreso(pnIndex).FtesIngresoDep(i).nGastosFam) / _
                                (FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngCony + _
                                FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngFam + _
                                FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngOtros - FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nGastosFam), "#0.00")), 8, 2)
                                
            End If
        End If
    Next i
    
    'IMPRIME COSTO DE PRODUCCION
    '-------------------------------------
    If pbCostoProd = True Then
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & Chr$(10) & Chr$(10)
            sCadImp = sCadImp & String(100, "-") & Chr$(10)
            sCadImp = sCadImp & Space(30) & "COSTO DE PRODUCCION" & Chr$(10)
            sCadImp = sCadImp & String(100, "-") & Chr$(10)
            sCadImp = sCadImp & ofun.ImpreFormat("Evaluacion", 16)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).dFecEval, 14, 0)
            Next i
            sCadImp = sCadImp & Chr$(10)
                
            sCadImp = sCadImp & ofun.ImpreFormat("Tipo Cultivo", 20)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).sTpoCultivo, 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ofun.ImpreFormat("DESTINO", 20)
            Dim sDestino As String
            
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                If FtesIngreso(pnIndex).CostoProd(i).nSiembra = 1 Then
        '            If Len(sDestino) = 0 Then
        '                sDestino = "SIEMBRA"
        '            Else
        '                sDestino = sDestino & ",SIEMBRA"
                        sCadImp = sCadImp & ofun.ImpreFormat("SIEMBRA", 13)
                        sCadImp = sCadImp & Chr$(10)
        '            End If
                 End If
                    
                If FtesIngreso(pnIndex).CostoProd(i).nMantenimiento = 1 Then
        '            If Len(sDestino) = 0 Then
        '                sDestino = "MANTENIMIENTO"
        '            Else
        '                sDestino = sDestino & ",MANTENIMIENTO"
        '                sCadImp = sCadImp & Chr$(10)
        '            End If
                        sCadImp = sCadImp & Space(22) & ofun.ImpreFormat("MANTENIMIENTO", 13)
                        sCadImp = sCadImp & Chr$(10)
                End If
                If FtesIngreso(pnIndex).CostoProd(i).nDesAgricola = 1 Then
        '             If Len(sDestino) = 0 Then
        '                sDestino = "DESAGRICOLA"
        '             Else
        '                sDestino = sDestino & ",DESAGRICOLA"
        '                sCadImp = sCadImp & Chr$(10)
        '             End If
                        sCadImp = sCadImp & Space(22) & ofun.ImpreFormat("DESAGRICOLA", 13)
                        sCadImp = sCadImp & Chr$(10)
        
        
                End If
                If FtesIngreso(pnIndex).CostoProd(i).nCosecha = 1 Then
        '            If Len(sDestino) = 0 Then
        '                sDestino = "COSECHA"
        '            Else
        '                sDestino = sDestino & ",COSECHA"
        '                sCadImp = sCadImp & Chr$(10)
        '            End If
                        sCadImp = sCadImp & Space(22) & ofun.ImpreFormat("COSECHA", 13)
                        sCadImp = sCadImp & Chr$(10)
        
                 End If
                If FtesIngreso(pnIndex).CostoProd(i).nPlanOtros = 1 Then
        '            If Len(sDestino) = 0 Then
        '                sDestino = "OTROS"
        '            Else
        '                sDestino = sDestino & ",OTROS"
        '            End If
                        sCadImp = sCadImp & Space(22) & ofun.ImpreFormat("OTROS", 13)
                        'sCadImp = sCadImp & Chr$(10)
        
                End If
                'sCadImp = sCadImp & oFun.ImpreFormat(sDestino, 60)
            Next i
            
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ofun.ImpreFormat("Maquinaria", 12)
            
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nMaquinaria, 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ofun.ImpreFormat("Jornales", 12)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nJornales, 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ofun.ImpreFormat("Insumos", 12)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nInsumos, 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ofun.ImpreFormat("Pesticidas", 12)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPesticidas, 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ofun.ImpreFormat("Otros", 12)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nOtros, 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ofun.ImpreFormat("Total", 12)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros, 13)
            Next i
            sCadImp = sCadImp & Chr$(10) & Chr$(10)
            
            sCadImp = sCadImp & ofun.ImpreFormat("Hectareas", 12)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas, 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ofun.ImpreFormat("Produccion", 18)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ofun.ImpreFormat(Format(FtesIngreso(pnIndex).CostoProd(i).nProduccion, "#0.00") & " " & FtesIngreso(pnIndex).CostoProd(i).sUniProd, 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ofun.ImpreFormat("Precio Uni.", 12)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPreUni, 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ofun.ImpreFormat("Ingresos", 12)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas * FtesIngreso(pnIndex).CostoProd(i).nPreUni * FtesIngreso(pnIndex).CostoProd(i).nProduccion, 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ofun.ImpreFormat("Egresos", 12)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas * (FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros), 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ofun.ImpreFormat("Utilidad", 12)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ofun.ImpreFormat((FtesIngreso(pnIndex).CostoProd(i).nHectareas * FtesIngreso(pnIndex).CostoProd(i).nPreUni * FtesIngreso(pnIndex).CostoProd(i).nProduccion) - (FtesIngreso(pnIndex).CostoProd(i).nHectareas * (FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros)), 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            sCadImp = sCadImp & String(100, "-") & Chr$(10)
    End If
        
    
    
    ' ---------------------------------------------------------------------------------------------------------------
    ' Muestra los datos de las referencias personales de la fte como persona ----------------------------------------
    sCadImp = sCadImp & Chr$(10) & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "REFERENCIAS PERSONALES" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    
    sCadImp = sCadImp & FtesIngreso(pnIndex).sComentarios & Chr$(10)

    sCadImp = sCadImp & Chr$(10)
 
    ' Muestra los datos de las referencias bancarias de la fte como persona ----------------------------------------
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("REFERENCIAS BANCARIAS", 42, 0) & ofun.ImpreFormat("NRO.TARJETA", 22, 2) & ofun.ImpreFormat("NRO.CUENTA", 22, 2) & ofun.ImpreFormat("MONTO", 10, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF

    'MatDatosFte = RecuperaRefBancariaPersFte(FtesIngreso(pnIndex).sFteIngreso)
    MatDatosFte = RecuperaRefBancariaPersFte(sPersCod)
        
    For i = 0 To UBound(MatDatosFte) - 1
        sCadImp = sCadImp & ofun.ImpreFormat(MatDatosFte(i, 0) & "-" & MatDatosFte(i, 1), 40) & ofun.ImpreFormat(MatDatosFte(i, 3), 22, 2) & ofun.ImpreFormat(MatDatosFte(i, 2), 22, 2) & ofun.ImpreFormat(MatDatosFte(i, 4), 10, 2) & Chr$(10)
    Next i
    
    sCadImp = sCadImp & Chr$(10)
    
   ' Muestra los datos de las referencias comerciales de la fte como persona ----------------------------------------
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("REFERENCIAS COMERCIALES", 66, 0) & ofun.ImpreFormat("TIPO RELAC. COMERCIAL", 22, 2) & ofun.ImpreFormat("TELEFONO", 10, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF

    'MatDatosComFte = RecuperaRefComercialPersFte(FtesIngreso(pnIndex).sFteIngreso)
    MatDatosComFte = RecuperaRefComercialPersFte(sPersCod)
    For i = 0 To UBound(MatDatosComFte) - 1
        sCadImp = sCadImp & ofun.ImpreFormat(MatDatosComFte(i, 3) & "-" & MatDatosComFte(i, 0), 64) & ofun.ImpreFormat(MatDatosComFte(i, 2), 22, 2) & ofun.ImpreFormat(MatDatosComFte(i, 1), 10, 2) & Chr$(10)
    Next i
    
    ' ---------------------------------------------------------------------------------------------------------------
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    GenerarImpresionFteIngresoDependiente = sCadImp
End Function

Public Function GenerarImpresionFteIngresoIndependiente(ByVal pnIndex As Integer, _
    ByVal psNomAge As String, ByVal pdFecha As Date, ByVal pbCostoProd As Boolean, _
    Optional ByVal psNomCmac As String = "") As String

Dim sCadImp As String
Dim MatDatos As Variant
Dim i As Integer
Dim lnNegritaON As String
Dim lnNegritaOFF As String

Dim MatDatosFte As Variant
Dim MatDatosComFte As Variant
Dim bProd As Boolean



    lnNegritaON = Chr$(27) & Chr$(69)
    lnNegritaOFF = Chr$(27) + Chr$(70)

    MatDatos = RecuperaHistoriaCredPersona(sPersCod)

    sCadImp = ""

    ' Comprime todo el reporte ----------------------------------------------------------------------
    sCadImp = sCadImp & Chr$(27) & Chr$(64)
    'sCadImp = sCadImp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
    sCadImp = sCadImp & Chr$(27) & Chr$(65) & Chr$(9)     'Espaciado

    sCadImp = sCadImp & Chr$(27) & Chr$(67) & Chr$(80)  'Longitud de página a 80 líneas'
    sCadImp = sCadImp & Chr$(27) & Chr$(103)   'Tamaño 10 cpi
    'sCadImp = sCadImp & Chr$(27) & Chr$(77)   'Tamaño 10 cpi
    sCadImp = sCadImp & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
    'sCadImp = sCadImp & Chr$(27) + Chr$(18) ' cancela condensada
    'sCadImp = sCadImp & Chr$(27) + Chr$(72) ' desactiva negrita
    'lsCadImp = lsCadImp & Chr$(27) & Chr$(108) & Chr$(20)   'Margen  izquierdo  20 caracteres
    '-------------------------------------------------------------------------------------------------

    'sCadImp = sCadImp & oFun.ImpreFormat(psNomCmac, 35, 35) & psNomAge
    sCadImp = sCadImp & ofun.ImpreFormat(psNomCmac, 35, 0) & psNomAge
    sCadImp = sCadImp & Space(25) & Format(pdFecha, "dd/mm/yyy hh:mm:ss") & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "Cliente : " & sPersCod & Space(2) & ocad.PstaNombre(sNombreCompleto) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & "Fte Ingreso : " & FtesIngreso(pnIndex).sRazSocDescrip & Chr$(10)


'----------------------------------------------------------------------------------------------------
    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)

    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("RAZON SOCIAL/NOMBRE COMERCIAL", 35, 0) & ofun.ImpreFormat("RUC", 14, 2) & ofun.ImpreFormat("INICIO ACT.", 12, 0) & ofun.ImpreFormat("PERS. JURIDICA", 24, 2) & ofun.ImpreFormat("TELEFONO", 10, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).sRazSocDescrip, 35, 0) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersRuc, 14, 2) & ofun.ImpreFormat(Format(FtesIngreso(pnIndex).dFtePersFecInicioAct, "dd/mm/yyyy"), 12, 0) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersTipoPersJur, 24, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersTelefono, 10, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("ACTIVIDAD(CIUU)", 35, 0) & ofun.ImpreFormat("COND. LOCAL", 14, 2) & ofun.ImpreFormat("MAGNITUD EMP.", 35, 0) & ofun.ImpreFormat("Nª EMPLEADOS", 14, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersCIUU, 35, 0) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersCondicionDomic, 14, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersMagnitudEmp, 35, 0) & ofun.ImpreFormat(Round(FtesIngreso(pnIndex).nFtePersNroEmpleados), 14, 0) & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("DIRECCION(Avenida, Calle, Jiron)", 35, 0) & ofun.ImpreFormat("DEPARTAMENTO", 12, 2) & ofun.ImpreFormat("PROVINCIA", 12, 2) & ofun.ImpreFormat("DISTRITO", 16, 2) & ofun.ImpreFormat("ZONA", 17, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersDireccion, 35, 0) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersDpto, 12, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersProv, 12, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersDist, 16, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersZona, 17, 2) & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'----------------------------------------------------------------------------------------------------
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & Space(20) & "HISTORIA CREDITICA" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
     sCadImp = sCadImp & String(100, "-") & Chr$(10)

    sCadImp = sCadImp & "MONEDA             : " & IIf(FtesIngreso(pnIndex).sMoneda = "1", "SOLES", "DOLARES")

    If Not IsArray(MatDatos) Then
        sCadImp = sCadImp & "NO POSEE CREDITOS" & Chr$(10)
    Else
        'sCadImp = sCadImp & "CREDITO No         : "
        'For I = 0 To UBound(MatDatos) - 1
        '    sCadImp = sCadImp & MatDatos(I, 0) & Space(2)
        'Next I

        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Estado Actual      : "
        For i = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(i, 1) & Space(2)
        Next i
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Fecha Vigencia     : "
        For i = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(i, 2) & Space(2)
        Next i
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Fecha Cancelacion  : "
        For i = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(i, 3) & Space(2)
        Next i
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Monto(Sol/Sug/Apr) : "
        For i = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(i, 4) & Space(2)
        Next i
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Cuotas/Plazo(Dias) : "
        For i = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(i, 5) & Space(2)
        Next i
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Cuota (Sol/Sug/Apr): "
        For i = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(i, 6) & Space(2)
        Next i
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Nota               : "
        For i = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(i, 7) & Space(2)
        Next i
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Analista           : "
        For i = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(i, 8) & Space(2)
        Next i
        sCadImp = sCadImp & Chr$(10)
    End If
    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & Space(20) & "EVOLUCION ECONOMICA FINANCIERA" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    sCadImp = sCadImp & Space(25)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & Format(FtesIngreso(pnIndex).FtesIngresoInd(i).dFecEval, "dd/mm/yyyy") & Space(3)
    Next i
    sCadImp = sCadImp & "Variacion" & Chr$(10)
    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("ACTIVO CIRCULANTE", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios) = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios), 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ofun.ImpreFormat("Disponible", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Cuentas x Cobrar", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Inventario", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("ACTIVO FIJO", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("TOTAL ACTIVO", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo), 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("PASIVO", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact = 0 Then

                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact), 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ofun.ImpreFormat("Proveedores", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Otros Prestamos", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Prestamos CMACT", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("PATRIMONIO", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("INGRESOS UE", 20)
    Dim nIng As Currency
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & oFun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob, 8, 2)
'
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob, 8, 2)
        nIng = FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob

        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob = 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob = 0 Then

                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
'                sCadImp = sCadImp & oFun.ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob), 8, 2)
               sCadImp = sCadImp & ofun.ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob), 8, 2)

            End If
        End If
    Next i

    'CMACICA_CSTS - 10/11/2003 - SE CAMBIO EL ORDEN (VENTAS, RECUP. X COB., INGRESOS FAMILIARES) -----------------
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF

    sCadImp = sCadImp & ofun.ImpreFormat("Ventas", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas, 8, 2)
        If i > 0 And FtesIngreso(pnIndex).nNumFtesIndep - 1 = i Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas, 8, 2)
            End If
        End If
    Next i

    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Recup. Ctas x Cob", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob, 8, 2)
            End If
        End If
    Next i

    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & oFun.ImpreFormat("Ingresos familiares", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & oFun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam = 0 Then
'                sCadImp = sCadImp & oFun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & oFun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam, 8, 2)
'            End If
'        End If
'    Next i

'-----------------------------------------------------------------------------------------------------------------------
    Dim nEg As Currency
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("EGRESOS UE", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & oFun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 8, 2)
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 8, 2)
         nEg = FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos = 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else

'                sCadImp = sCadImp & oFun.ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos), 8, 2)
                sCadImp = sCadImp & ofun.ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos), 8, 2)
            End If
        End If
    Next i

    'CMACICA_CSTS - 21/11/2003 - SE CAMBIO EL ORDEN (COMPRA DE MERCADERIAS, OTROS EGRESOS, GASTOS FAMILIARES) ---------
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF

    'sCadImp = sCadImp & oFun.ImpreFormat("Compra Mercaderia", 20)
    sCadImp = sCadImp & ofun.ImpreFormat("Costo de Ventas", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad, 8, 2)
            End If
        End If
    Next i

    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Otros Egresos", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos, 8, 2)
            End If
        End If
    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & oFun.ImpreFormat("Gastos Familiares", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & oFun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam = 0 Then
'                sCadImp = sCadImp & oFun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & oFun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam, 8, 2)
'            End If
'        End If
'    Next i

' ---------------------------------------------------------------------------------------------------------------------

    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("UTILIDAD NETA", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & oFun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob - _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam - _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad - _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 8, 2)
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob - _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad - _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 8, 2)

        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            If (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - _
'                        (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
'                         FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
'                         FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos) = 0 Then
            If (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - _
                        (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
                         FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos) = 0 Then

                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)

            Else
'                sCadImp = sCadImp & oFun.ImpreFormat(((FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob) - _
'                            (FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + _
'                             FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
'                             FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos)) _
'                            / ((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - _
'                            (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
'                             FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
'                             FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos)), 8, 2)
                sCadImp = sCadImp & ofun.ImpreFormat(((FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob) - _
                            (FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
                             FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos)) _
                            / ((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - _
                            (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
                             FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos)), 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    'sCadImp = sCadImp & Chr$(10)
    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & Space(20) & "INGRESOS-EGRESOS FAMILIARES" & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Ingresos familiares", 20)
    Dim nIngFam As Currency

    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam, 8, 2)
        nIngFam = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Gastos Familiares", 20)
    Dim nEgrFam As Currency
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam, 8, 2)
        nEgrFam = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("DISPONIBLE", 20)
    sCadImp = sCadImp & ofun.ImpreFormat(nIngFam + nIng - (nEgrFam + nEg), 8, 2)

    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & Space(20) & "RATIOS DE EVALUACION" & Chr$(10)
    sCadImp = sCadImp & "LIQUIDEZ" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF

    'Activo Circulante / Pasivo
    Dim nActCir As Double
    Dim nPasivo As Double
    Dim nActCirAnt As Double
    Dim nPasivoAnt As Double
    Dim Res1 As Double
    Dim Res2 As Double
    sCadImp = sCadImp & ofun.ImpreFormat("Liquidez", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nActCir = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
        If nPasivo = 0 Then
            Res1 = 0
        Else
            Res1 = nActCir / nPasivo
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nActCirAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
            nPasivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
            If nPasivoAnt = 0 Then
                Res2 = 0
            Else
                Res2 = nActCirAnt / nPasivoAnt
            End If
            If Res2 = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next i

    sCadImp = sCadImp & Chr$(10)
    'Activo Circulante - Pasivo
    sCadImp = sCadImp & ofun.ImpreFormat("Capital Trabajo", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nActCir = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
        Res1 = nActCir - nPasivo
        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nActCirAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
            nPasivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
            Res2 = nActCirAnt - nPasivoAnt
            If Res2 = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "GESTION" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    'Costo de Ventas / Inventario
    sCadImp = sCadImp & ofun.ImpreFormat("Rotacion Inventario", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios = 0 Then
            Res1 = 0
        Else
            Res1 = FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)

        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios = 0 Then
                Res2 = 0
            Else
                Res2 = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
            End If
            If Res2 = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)

    'Recup Ctas x Cob / (Ventas / 30)
    sCadImp = sCadImp & ofun.ImpreFormat("Periodo de Cobro", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
            Res1 = 0
        Else
            Res1 = FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob / (FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas / 30)
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
                Res2 = 0
            Else
                Res2 = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas / 30)
            End If
            If Res2 = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)

    Dim nIngresos As Double
    Dim nEgresos As Double
    Dim nIngresosAnt As Double
    Dim nEgresosAnt As Double

    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "RENTABILIDAD" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    'Utilidad Operat =  Ingresos - Egresos
    'Utilidad Operat / Patrimonio
    sCadImp = sCadImp & ofun.ImpreFormat("Rentab. Patrimonial", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
        nEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos

        If FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio = 0 Then
            Res1 = 0
        Else
            Res1 = (nIngresos - nEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nIngresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob
            nEgresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio = 0 Then
                Res2 = 0
            Else
                Res2 = (nIngresos - nEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio
            End If

            If Res2 = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)

    Dim nTotActivo As Double
    Dim nTotActivoAnt As Double
    sCadImp = sCadImp & ofun.ImpreFormat("Rentab. Activo", 20)
    'Utilidad Operativa / Total Activo
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
        nEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos
        nTotActivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo
        If nTotActivo = 0 Then
            Res1 = 0
        Else
            Res1 = (nIngresos - nEgresos) / nTotActivo
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nIngresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob
            nEgresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos
            nTotActivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo
            If nTotActivoAnt = 0 Then
                Res2 = 0
            Else
                Res2 = (nIngresosAnt - nEgresosAnt) / nTotActivoAnt
            End If
            If Res2 = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)

    sCadImp = sCadImp & ofun.ImpreFormat("Rentab. Ventas", 20)
    'Utilidad Operativa / Ventas
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
        nEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
            Res1 = 0
        Else
            Res1 = (nIngresos - nEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nIngresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob
            nEgresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
                Res2 = 0
            Else
                Res2 = (nIngresosAnt - nEgresosAnt) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas
            End If
            If Res2 = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)

    'Ventas Totales = Ventas + Recup Ctas x Cob
    'Margen Utilidad = Ventas Totales - Costo Ventas / Ventas Totales
    Dim nVentasTot As Double
    Dim nVentasTotAnt As Double
    sCadImp = sCadImp & ofun.ImpreFormat("Margen Utilidad", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nVentasTot = FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
        If nVentasTot = 0 Then
            Res1 = 0
        Else
            Res1 = (nVentasTot - FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad) / nVentasTot
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nVentasTotAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
            If nVentasTotAnt = 0 Then
                Res2 = 0
            Else
                Res2 = (nVentasTotAnt - FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad) / nVentasTotAnt
            End If
            If Res2 = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)

    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "ENDEUDAMIENTO" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ofun.ImpreFormat("Endeudamiento Pat.", 20)
    'Endeudamiento = Pasivo / Patrimonio
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact

        If FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio = 0 Then
            Res1 = 0
        Else
            Res1 = nPasivo / FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nPasivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact

            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio = 0 Then
                Res2 = 0
            Else
                Res2 = nPasivoAnt / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio
            End If
            If Res2 = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)

    sCadImp = sCadImp & ofun.ImpreFormat("Endeudamiento", 20)
    'Endeudamiento = Pasivo / TotalActivo
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nTotActivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo
        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact

        If nTotActivo = 0 Then
            Res1 = 0
        Else
            Res1 = nPasivo / nTotActivo
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nTotActivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo
            nPasivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
            If nTotActivoAnt = 0 Then
                Res2 = 0
            Else
                Res2 = nPasivoAnt / nTotActivoAnt
            End If
            If Res2 = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next i

    If pbCostoProd = True Then

        'IMPRIME COSTO DE PRODUCCION
        '-------------------------------------
        sCadImp = sCadImp & Chr$(12)
        sCadImp = sCadImp & String(100, "-") & Chr$(10)
        sCadImp = sCadImp & Space(30) & "COSTO DE PRODUCCION" & Chr$(10)
        sCadImp = sCadImp & String(100, "-") & Chr$(10)
        sCadImp = sCadImp & ofun.ImpreFormat("Evaluacion", 16)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).dFecEval, 14, 0)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Tipo Cultivo", 20)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).sTpoCultivo, 13)
        Next i

        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Destino", 20)

        Dim sDestino As String
        bProd = False
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            If FtesIngreso(pnIndex).CostoProd(i).nSiembra = 1 Then
    '            If Len(sDestino) = 0 Then
    '                sDestino = "SIEMBRA"
    '            Else
    '                sDestino = sDestino & ",SIEMBRA"
                    bProd = True
                    sCadImp = sCadImp & ofun.ImpreFormat("SIEMBRA", 13)
                    'sCadImp = sCadImp & Chr$(10)
    '            End If
             End If

            If FtesIngreso(pnIndex).CostoProd(i).nMantenimiento = 1 Then
    '            If Len(sDestino) = 0 Then
    '                sDestino = "MANTENIMIENTO"
    '            Else
    '                sDestino = sDestino & ",MANTENIMIENTO"
    '                sCadImp = sCadImp & Chr$(10)
    '            End If
                If bProd = True Then
                    sCadImp = sCadImp & Space(7) & ofun.ImpreFormat("MANTENIMIENTO", 13)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat("MANTENIMIENTO", 13)
                    bProd = True
                End If
                   ' sCadImp = sCadImp & Chr$(10)
            End If
            If FtesIngreso(pnIndex).CostoProd(i).nDesAgricola = 1 Then
    '             If Len(sDestino) = 0 Then
    '                sDestino = "DESAGRICOLA"
    '             Else
    '                sDestino = sDestino & ",DESAGRICOLA"
    '                sCadImp = sCadImp & Chr$(10)
    '             End If
                  If bProd = True Then
                    sCadImp = sCadImp & Space(7) & ofun.ImpreFormat("DESAGRICOLA", 13)
                  Else
                    sCadImp = sCadImp & ofun.ImpreFormat("DESAGRICOLA", 13)
                    bProd = True
                  End If
                    'sCadImp = sCadImp & Chr$(10)


            End If
            If FtesIngreso(pnIndex).CostoProd(i).nCosecha = 1 Then
    '            If Len(sDestino) = 0 Then
    '                sDestino = "COSECHA"
    '            Else
    '                sDestino = sDestino & ",COSECHA"
    '                sCadImp = sCadImp & Chr$(10)
    '            End If
                If bProd = True Then
                    sCadImp = sCadImp & Space(7) & ofun.ImpreFormat("COSECHA", 13)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat("COSECHA", 13)
                    bProd = True
                End If
                    'sCadImp = sCadImp & Chr$(10)

             End If
            If FtesIngreso(pnIndex).CostoProd(i).nPlanOtros = 1 Then
    '            If Len(sDestino) = 0 Then
    '                sDestino = "OTROS"
    '            Else
    '                sDestino = sDestino & ",OTROS"
    '            End If
                    If bProd = True Then
                        sCadImp = sCadImp & Space(7) & ofun.ImpreFormat("OTROS", 13)
                    Else
                        sCadImp = sCadImp & ofun.ImpreFormat("OTROS", 13)
                        bProd = True
                    End If
                    'sCadImp = sCadImp & Chr$(10)

            End If
            'sCadImp = sCadImp & oFun.ImpreFormat(sDestino, 60)
        Next i

        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Maquinaria", 12)

        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nMaquinaria, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Jornales", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nJornales, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Insumos", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nInsumos, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Pesticidas", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPesticidas, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Otros", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nOtros, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Total", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros, 13)
        Next i
        sCadImp = sCadImp & Chr$(10) & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Hectareas", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Produccion", 18)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(Format(FtesIngreso(pnIndex).CostoProd(i).nProduccion, "#0.00") & " " & FtesIngreso(pnIndex).CostoProd(i).sUniProd, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Precio Uni.", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPreUni, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Ingresos", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas * FtesIngreso(pnIndex).CostoProd(i).nPreUni * FtesIngreso(pnIndex).CostoProd(i).nProduccion, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Egresos", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas * (FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros), 13)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Utilidad", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat((FtesIngreso(pnIndex).CostoProd(i).nHectareas * FtesIngreso(pnIndex).CostoProd(i).nPreUni * FtesIngreso(pnIndex).CostoProd(i).nProduccion) - (FtesIngreso(pnIndex).CostoProd(i).nHectareas * (FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros)), 13)
        Next i
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & String(100, "-") & Chr$(10)
 End If
    ' ---------------------------------------------------------------------------------------------------------------
    ' Muestra los datos de las referencias personales de la fte como persona ----------------------------------------
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "REFERENCIAS PERSONALES" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF

    sCadImp = sCadImp & FtesIngreso(pnIndex).sComentarios & Chr$(10)
    sCadImp = sCadImp & Chr$(10)

    ' Muestra los datos de las referencias bancarias de la fte como persona ----------------------------------------
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("REFERENCIAS BANCARIAS", 42, 0) & ofun.ImpreFormat("NRO.TARJETA", 22, 2) & ofun.ImpreFormat("NRO.CUENTA", 22, 2) & ofun.ImpreFormat("MONTO", 10, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF

    MatDatosFte = RecuperaRefBancariaPersFte(FtesIngreso(pnIndex).sFteIngreso)

    For i = 0 To UBound(MatDatosFte) - 1
        sCadImp = sCadImp & ofun.ImpreFormat(MatDatosFte(i, 0) & "-" & MatDatosFte(i, 1), 40) & ofun.ImpreFormat(MatDatosFte(i, 3), 22, 2) & ofun.ImpreFormat(MatDatosFte(i, 2), 22, 2) & ofun.ImpreFormat(MatDatosFte(i, 4), 10, 2) & Chr$(10)
    Next i

    sCadImp = sCadImp & Chr$(10)

   ' Muestra los datos de las referencias comerciales de la fte como persona ----------------------------------------
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("REFERENCIAS COMERCIALES", 66, 0) & ofun.ImpreFormat("TIPO RELAC. COMERCIAL", 22, 2) & ofun.ImpreFormat("TELEFONO", 10, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF

    MatDatosComFte = RecuperaRefComercialPersFte(FtesIngreso(pnIndex).sFteIngreso)

    For i = 0 To UBound(MatDatosComFte) - 1
        sCadImp = sCadImp & ofun.ImpreFormat(MatDatosComFte(i, 3) & "-" & MatDatosComFte(i, 0), 64) & ofun.ImpreFormat(MatDatosComFte(i, 2), 22, 2) & ofun.ImpreFormat(MatDatosComFte(i, 1), 10, 2) & Chr$(10)
    Next i

    ' ---------------------------------------------------------------------------------------------------------------
    sCadImp = sCadImp & Chr$(10)

    GenerarImpresionFteIngresoIndependiente = sCadImp
    
End Function

'****************** Nuevos Ratios para CMAC-CUSCO
Public Function GenerarImpresionFteIngresoIndependiente_CS(ByVal pnIndex As Integer, _
    ByVal psNomAge As String, ByVal pdFecha As Date, ByVal pbCostoProd As Boolean, _
    Optional ByVal psNomCmac As String = "") As String

Dim sCadImp As String
Dim MatDatos As Variant
Dim i As Integer
Dim lnNegritaON As String
Dim lnNegritaOFF As String

Dim MatDatosFte As Variant
Dim MatDatosComFte As Variant
Dim bProd As Boolean

Dim nMontoT As Double

    lnNegritaON = Chr$(27) & Chr$(69)
    lnNegritaOFF = Chr$(27) + Chr$(70)

    MatDatos = RecuperaHistoriaCredPersona(sPersCod)

    sCadImp = ""

    ' Comprime todo el reporte ----------------------------------------------------------------------
    sCadImp = sCadImp & Chr$(27) & Chr$(64)
    'sCadImp = sCadImp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
    sCadImp = sCadImp & Chr$(27) & Chr$(65) & Chr$(9)     'Espaciado

    sCadImp = sCadImp & Chr$(27) & Chr$(67) & Chr$(80)  'Longitud de página a 80 líneas'
    sCadImp = sCadImp & Chr$(27) & Chr$(103)   'Tamaño 10 cpi
    'sCadImp = sCadImp & Chr$(27) & Chr$(77)   'Tamaño 10 cpi
    sCadImp = sCadImp & Chr$(27) + Chr$(107) + Chr$(0)     'Tipo de Letra Sans Serif
    'sCadImp = sCadImp & Chr$(27) + Chr$(18) ' cancela condensada
    'sCadImp = sCadImp & Chr$(27) + Chr$(72) ' desactiva negrita
    'lsCadImp = lsCadImp & Chr$(27) & Chr$(108) & Chr$(20)   'Margen  izquierdo  20 caracteres
    '-------------------------------------------------------------------------------------------------

    'sCadImp = sCadImp & oFun.ImpreFormat(psNomCmac, 35, 35) & psNomAge
    sCadImp = sCadImp & ofun.ImpreFormat(psNomCmac, 35, 0) & psNomAge
    sCadImp = sCadImp & Space(25) & Format(pdFecha, "dd/mm/yyyy hh:mm:ss") & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "Cliente : " & sPersCod & Space(2) & ocad.PstaNombre(sNombreCompleto) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & "Fte Ingreso : " & FtesIngreso(pnIndex).sRazSocDescrip & Chr$(10)


'----------------------------------------------------------------------------------------------------
    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)

    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("RAZON SOCIAL/NOMBRE COMERCIAL", 35, 0) & ofun.ImpreFormat("RUC", 14, 2) & ofun.ImpreFormat("INICIO ACT.", 12, 0) & ofun.ImpreFormat("PERS. JURIDICA", 24, 2) & ofun.ImpreFormat("TELEFONO", 10, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).sRazSocDescrip, 35, 0) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersRuc, 14, 2) & ofun.ImpreFormat(Format(FtesIngreso(pnIndex).dFtePersFecInicioAct, "dd/mm/yyyy"), 12, 0) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersTipoPersJur, 24, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersTelefono, 10, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("ACTIVIDAD(CIUU)", 35, 0) & ofun.ImpreFormat("COND. LOCAL", 14, 2) & ofun.ImpreFormat("MAGNITUD EMP.", 35, 0) & ofun.ImpreFormat("Nª EMPLEADOS", 14, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersCIUU, 35, 0) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersCondicionDomic, 14, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersMagnitudEmp, 35, 0) & ofun.ImpreFormat(Round(FtesIngreso(pnIndex).nFtePersNroEmpleados), 14, 0) & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("DIRECCION(Avenida, Calle, Jiron)", 35, 0) & ofun.ImpreFormat("DEPARTAMENTO", 12, 2) & ofun.ImpreFormat("PROVINCIA", 12, 2) & ofun.ImpreFormat("DISTRITO", 16, 2) & ofun.ImpreFormat("ZONA", 17, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersDireccion, 35, 0) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersDpto, 12, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersProv, 12, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersDist, 16, 2) & ofun.ImpreFormat(FtesIngreso(pnIndex).cFtePersZona, 17, 2) & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'----------------------------------------------------------------------------------------------------
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & Space(20) & "HISTORIA CREDITICA" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
     sCadImp = sCadImp & String(100, "-") & Chr$(10)

    sCadImp = sCadImp & "MONEDA             : " & IIf(FtesIngreso(pnIndex).sMoneda = "1", "SOLES", "DOLARES")

    If Not IsArray(MatDatos) Then
        sCadImp = sCadImp & "NO POSEE CREDITOS" & Chr$(10)
    Else
        'sCadImp = sCadImp & "CREDITO No         : "
        'For I = 0 To UBound(MatDatos) - 1
        '    sCadImp = sCadImp & MatDatos(I, 0) & Space(2)
        'Next I

        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Estado Actual      : "
        For i = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(i, 1) & Space(2)
        Next i
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Fecha Vigencia     : "
        For i = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(i, 2) & Space(2)
        Next i
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Fecha Cancelacion  : "
        For i = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(i, 3) & Space(2)
        Next i
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Monto(Sol/Sug/Apr) : "
        For i = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(i, 4) & Space(2)
        Next i
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Cuotas/Plazo(Dias) : "
        For i = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(i, 5) & Space(2)
        Next i
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Cuota (Sol/Sug/Apr): "
        For i = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(i, 6) & Space(2)
        Next i
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Nota               : "
        For i = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(i, 7) & Space(2)
        Next i
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Analista           : "
        For i = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(i, 8) & Space(2)
        Next i
        sCadImp = sCadImp & Chr$(10)
    End If
    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & Space(20) & "EVOLUCION ECONOMICA FINANCIERA" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    sCadImp = sCadImp & Space(25)
    
    Dim dFecTemEval As Date
    
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nMontoT = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp
        If FtesIngreso(pnIndex).FtesIngresoInd(i).dFecEval <> 0 Then
            sCadImp = sCadImp & Format(FtesIngreso(pnIndex).FtesIngresoInd(i).dFecEval, "dd/mm/yyyy") & Space(3)
        Else
            sCadImp = sCadImp & Space(5) & "(-)" & Space(3)
        End If
    Next i
    sCadImp = sCadImp & "Variacion" & Chr$(10)
    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("ACTIVO CORRIENTE", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios) = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios), 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ofun.ImpreFormat("Disponible", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Cuentas x Cobrar", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Inventario", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("ACTIVO NO CORRIENTE", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("TOTAL ACTIVO", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo), 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("PASIVO", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact = 0 Then

                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact), 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ofun.ImpreFormat("Proveedores", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Otros Prestamos", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Prestamos CMAC-C", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("PATRIMONIO", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("INGRESOS", 20)
    Dim nIng As Currency
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + _
                FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
                FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob, 8, 2)

'        nIng = FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob + _
                FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam

        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam = 0 Then

                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
               sCadImp = sCadImp & ofun.ImpreFormat(nIng / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob), 8, 2)

            End If
        End If
    Next i

    'CMACICA_CSTS - 10/11/2003 - SE CAMBIO EL ORDEN (VENTAS, RECUP. X COB., INGRESOS FAMILIARES) -----------------
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF

    sCadImp = sCadImp & ofun.ImpreFormat("Ventas", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas, 8, 2)
        If i > 0 And FtesIngreso(pnIndex).nNumFtesIndep - 1 = i Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas, 8, 2)
            End If
        End If
    Next i

    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Recup. Ctas x Cob", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob, 8, 2)
            End If
        End If
    Next i

    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Otros Ingresos", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam, 8, 2)
            End If
        End If
    Next i

'-----------------------------------------------------------------------------------------------------------------------
    Dim nEg As Currency
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("EGRESOS", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 8, 2)
        'sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 8, 2)
         'nEg = FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos = 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else

                sCadImp = sCadImp & ofun.ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos), 8, 2)
'                sCadImp = sCadImp & ofun.ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos), 8, 2)
            End If
        End If
    Next i

    'CMACICA_CSTS - 21/11/2003 - SE CAMBIO EL ORDEN (COMPRA DE MERCADERIAS, OTROS EGRESOS, GASTOS FAMILIARES) ---------
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF

    'sCadImp = sCadImp & oFun.ImpreFormat("Compra Mercaderia", 20)
    sCadImp = sCadImp & ofun.ImpreFormat("Gasto de Ventas", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad, 8, 2)
            End If
        End If
    Next i

    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Gastos Administrativos", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam, 8, 2)
            End If
        End If
    Next i

    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Otros Egresos", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos, 8, 2)
            End If
        End If
    Next i
    
' ---------------------------------------------------------------------------------------------------------------------

    'sCadImp = sCadImp & Chr$(10)
    'sCadImp = sCadImp & lnNegritaON
    'sCadImp = sCadImp & ofun.ImpreFormat("UTILIDAD NETA", 20)
    'For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & oFun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob - _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam - _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad - _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 8, 2)
     '   sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob - _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad - _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 8, 2)

     '   If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then

'            If (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - _
'                        (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
'                         FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
'                         FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos) = 0 Then
      '      If (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
      '                  FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - _
      '                  (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
      '                   FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos) = 0 Then

      '          sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)

      '     Else
'                sCadImp = sCadImp & oFun.ImpreFormat(((FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob) - _
'                            (FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + _
'                             FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
'                             FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos)) _
'                            / ((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - _
'                            (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
'                             FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
'                             FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos)), 8, 2)
         '       sCadImp = sCadImp & ofun.ImpreFormat(((FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob) - _
                            (FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
                             FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos)) _
                            / ((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - _
                            (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
                             FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos)), 8, 2)
          '  End If
        'End If
    'Next i
  '  sCadImp = sCadImp & Chr$(10)
    'sCadImp = sCadImp & Chr$(10)
    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
  '  sCadImp = sCadImp & String(100, "-") & Chr$(10)
  '  sCadImp = sCadImp & Chr$(10)
  '  sCadImp = sCadImp & Space(20) & "INGRESOS-EGRESOS FAMILIARES" & Chr$(10)
  '  sCadImp = sCadImp & ofun.ImpreFormat("Ingresos familiares", 20)
'    Dim nIngFam As Currency

'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam, 8, 2)
'        nIngFam = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam
'        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ofun.ImpreFormat("Gastos Familiares", 20)
'    Dim nEgrFam As Currency
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam, 8, 2)
'        nEgrFam = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam, 8, 2)
'            End If
'        End If
'    Next i
    sCadImp = sCadImp & Chr$(10) & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("DISPONIBLE", 20)
    'sCadImp = sCadImp & ofun.ImpreFormat(nIngFam + nIng - (nEgrFam + nEg), 8, 2)
    
    Dim nTemp As Double
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nIng = FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob + _
                FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam
        nEg = FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam
        
        sCadImp = sCadImp & ofun.ImpreFormat(nIng - nEg, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            
            nTemp = ((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam) - (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam))
            If nTemp = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & _
                         ofun.ImpreFormat((nIng - nEg) / ((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam) - (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam)), 8, 2)
           End If
        End If
    Next i
    
'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************
Dim nLiquidezGeneral As Currency
Dim nPruebaAcida As Currency
Dim nCapitalTrabajo As Currency
Dim nRotacionCtasCobrar As Currency
Dim nPeriodoMedioC As Currency
Dim nRotacionCtasPagar As Currency
Dim nPeriodoMedioP As Currency
Dim nRotacInvV As Currency
Dim nRotacInvD As Currency
Dim nRotacActivoFijo As Currency
Dim nMargenBruto As Currency
Dim nGradoEndeudam As Currency
Dim nGradoPropiedad As Currency
Dim nEndeudaPatrim As Currency
Dim nRentabPatrim As Currency
Dim nRentabVentas As Currency
Dim nActivoCorriente As Currency 'Disponible + Ctas por Cobrar + Inventario
Dim nPasivo As Currency 'Proveedores + Otros Prestamos + Prestamos Cmac
Dim nTotalIngresos As Currency  'Ventas + Recup. Ctas x Cobrar + Otros ingresos
Dim nTotalEgresos As Currency     'Gastos de Ventas + Gastos Administrativos + Otros Egresos
Dim nTotalActivo As Currency    'Disponible + Ctas por Cobrar + Inventario+ Activo No Corriente

    sCadImp = sCadImp & Chr$(10) & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    sCadImp = sCadImp & Space(20) & "RATIOS DE EVALUACION" & Chr$(10) & Chr$(10)
    sCadImp = sCadImp & "CREDITOS MICRO Y PEQUEÑA EMPRESA" & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    sCadImp = sCadImp & "LIQUIDEZ"
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Liquidez General", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nActivoCorriente = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
        
        If nPasivo = 0 Then
            nLiquidezGeneral = 0
        Else
            nLiquidezGeneral = nActivoCorriente / nPasivo
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nLiquidezGeneral, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nActivoCorriente = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
            nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
            If nPasivo = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If nActivoCorriente = 0 Then
                    sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat(nLiquidezGeneral / (nActivoCorriente / nPasivo), 8, 2)
                End If
            End If
        End If
        
    Next i
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Prueba Acida", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nActivoCorriente = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
        
        If nPasivo = 0 Then
            nPruebaAcida = 0
        Else
            nPruebaAcida = (nActivoCorriente - FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios) / nPasivo
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nPruebaAcida, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nActivoCorriente = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
            nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
            If nPasivo = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If (nActivoCorriente - FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios) = 0 Then
                    sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat(nPruebaAcida / ((nActivoCorriente - FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios) / nPasivo), 8, 2)
                End If
            End If
        End If
    Next i
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "GESTION" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF

    sCadImp = sCadImp & ofun.ImpreFormat("Periodo Med. Cobranza", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
            nPeriodoMedioC = 0
        Else
            nPeriodoMedioC = (FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob * 365 / FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas)
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nPeriodoMedioC, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob * 365) = 0 Then
                    sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat(nPeriodoMedioC / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob * 365 / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas), 8, 2)
                End If
            End If
        End If
    Next i
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Rotacion Invent.(Veces)", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios = 0 Then
            nRotacInvV = 0
        Else
            nRotacInvV = FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nRotacInvV, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad = 0 Then
                   sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                   sCadImp = sCadImp & ofun.ImpreFormat(nRotacInvV / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios), 8, 2)
                End If
            End If
        End If
    Next i
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Margen Bruto", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nTotalIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas '+ _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam
        
        nTotalEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad '+ _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos

        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
            nMargenBruto = 0
        Else
            nMargenBruto = (nTotalIngresos - nTotalEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nMargenBruto, 8, 2)
    
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nTotalIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas '+ _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam
        
            nTotalEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad '+ _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos
                        
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If (nTotalIngresos - nTotalEgresos) = 0 Then
                    sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat(nMargenBruto / ((nTotalIngresos - nTotalEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas), 8, 2)
                End If
            End If
        End If
    Next i
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "SOLVENCIA O ENDEUDAMIENTO" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
        
    sCadImp = sCadImp & ofun.ImpreFormat("Endeudamiento Patrim.", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio = 0 Then
            nEndeudaPatrim = 0
        Else
            nEndeudaPatrim = nPasivo / FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nEndeudaPatrim, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
            If nPasivo = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio = 0 Then
                    sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat(nEndeudaPatrim / (nPasivo / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio), 8, 2)
                End If
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "RENTABILIDAD" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
        
    sCadImp = sCadImp & ofun.ImpreFormat("Rentab. Neta Patrim.", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio = 0 Then
            nRentabPatrim = 0
        Else    'Utilidad_Neta/Patrimonio
            nRentabPatrim = ((FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam)) / FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nRentabPatrim, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If ((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam)) = 0 Then
                    sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat(nRentabPatrim / (((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam)) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio), 8, 2)
                End If
            End If
        End If
    Next i
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Rentab. sobre Ventas", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
            nRentabVentas = 0
        Else    'Utilidad_Neta/Ventas
            nRentabVentas = ((FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam)) / FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nRentabVentas, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If ((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam)) = 0 Then
                    sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat(nRentabVentas / (((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam)) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas), 8, 2)
                End If
            End If
        End If
    Next i
    
    sCadImp = sCadImp & Chr$(10) & Chr$(10)
    sCadImp = sCadImp & "CREDITOS COMERCIALES" & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    sCadImp = sCadImp & "LIQUIDEZ"
    sCadImp = sCadImp & Chr$(10)
    
    sCadImp = sCadImp & ofun.ImpreFormat("Liquidez General", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nActivoCorriente = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
        
        If nPasivo = 0 Then
            nLiquidezGeneral = 0
        Else
            nLiquidezGeneral = nActivoCorriente / nPasivo
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nLiquidezGeneral, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nActivoCorriente = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
            nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
            If nPasivo = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If nActivoCorriente = 0 Then
                    sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat(nLiquidezGeneral / (nActivoCorriente / nPasivo), 8, 2)
                End If
            End If
        End If
        
    Next i
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Prueba Acida", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nActivoCorriente = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
        
        If nPasivo = 0 Then
            nPruebaAcida = 0
        Else
            nPruebaAcida = (nActivoCorriente - FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios) / nPasivo
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nPruebaAcida, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nActivoCorriente = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
            nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
            If nPasivo = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If (nActivoCorriente - FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios) = 0 Then
                    sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat(nPruebaAcida / ((nActivoCorriente - FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios) / nPasivo), 8, 2)
                End If
            End If
        End If
    Next i
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Capital de Trabajo", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nActivoCorriente = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
        
        nCapitalTrabajo = nActivoCorriente - nPasivo
        sCadImp = sCadImp & ofun.ImpreFormat(nCapitalTrabajo, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nActivoCorriente = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
            nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
            If nPasivo = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ofun.ImpreFormat(nCapitalTrabajo / (nActivoCorriente - nPasivo), 8, 2)
            End If
        End If
    Next i
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "GESTION" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF

    sCadImp = sCadImp & ofun.ImpreFormat("Rotac. Ctas x Cobrar", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob = 0 Then
            nRotacionCtasCobrar = 0
        Else
            nRotacionCtasCobrar = FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas / FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nRotacionCtasCobrar, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
                    sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat(nRotacionCtasCobrar / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob), 8, 2)
                End If
            End If
        End If
    Next i
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Periodo Med. Cobranza", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
            nPeriodoMedioC = 0
        Else
            nPeriodoMedioC = (FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob * 365 / FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas)
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nPeriodoMedioC, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob * 365) = 0 Then
                    sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat(nPeriodoMedioC / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob * 365 / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas), 8, 2)
                End If
            End If
        End If
    Next i
        
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Rotac. Ctas x Pagar", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob = 0 Then
            nRotacionCtasPagar = 0
        Else
            nRotacionCtasPagar = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam / FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nRotacionCtasPagar, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam = 0 Then
                    sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat(nRotacionCtasPagar / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob), 8, 2)
                End If
            End If
        End If
    Next i
            
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Periodo Med. Pago", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam = 0 Then
            nPeriodoMedioP = 0
        Else
            nPeriodoMedioP = (FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob * 365 / FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam)
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nPeriodoMedioP, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob * 365) = 0 Then
                    sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat(nPeriodoMedioP / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob * 365 / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam), 8, 2)
                End If
            End If
        End If
    Next i
            
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Rotacion Invent.(Veces)", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios = 0 Then
            nRotacInvV = 0
        Else
            nRotacInvV = FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nRotacInvV, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad = 0 Then
                   sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat(nRotacInvV / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios), 8, 2)
                End If
            End If
        End If
    Next i
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Rotacion Invent.(Dias)", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad = 0 Then
            nRotacInvD = 0
        Else
            nRotacInvD = FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios * 365 / FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nRotacInvD, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios * 365) = 0 Then
                    sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat(nRotacInvD / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios * 365 / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad), 8, 2)
                End If
            End If
        End If
    Next i
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Rotacion Activo Fijo", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo = 0 Then
            nRotacActivoFijo = 0
        Else
            nRotacActivoFijo = FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas / FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nRotacActivoFijo, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
                   sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                   sCadImp = sCadImp & ofun.ImpreFormat(nRotacActivoFijo / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo), 8, 2)
                End If
            End If
        End If
    Next i
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Margen Bruto", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nTotalIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas '+ _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam
        
        nTotalEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad '+ _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos

        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
            nMargenBruto = 0
        Else
            nMargenBruto = (nTotalIngresos - nTotalEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nMargenBruto, 8, 2)
    
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nTotalIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas '+ _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam
        
            nTotalEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad '+ _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos
                        
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If (nTotalIngresos - nTotalEgresos) = 0 Then
                    sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat(nMargenBruto / ((nTotalIngresos - nTotalEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas), 8, 2)
                End If
            End If
        End If
    Next i
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "SOLVENCIA O ENDEUDAMIENTO" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
        
    sCadImp = sCadImp & ofun.ImpreFormat("Grado Endeudam.", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nTotalActivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo
        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
                    FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
                    FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact

        If nTotalActivo = 0 Then
            nGradoEndeudam = 0
        Else
            nGradoEndeudam = nPasivo / nTotalActivo
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nGradoEndeudam, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nTotalActivo = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo
            nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
            
            If nTotalActivo = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If nPasivo = 0 Then
                    sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat(nGradoEndeudam / (nPasivo / nTotalActivo), 8, 2)
                End If
            End If
        End If
    Next i
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Grado de Propiedad", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nTotalActivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo
        
        If nTotalActivo = 0 Then
            nGradoPropiedad = 0
        Else
            nGradoPropiedad = FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio / nTotalActivo
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nGradoPropiedad, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nTotalActivo = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo
            
            If nTotalActivo = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio = 0 Then
                    sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat(nGradoPropiedad / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio / nTotalActivo), 8, 2)
                End If
            End If
        End If
    Next i
        
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Endeudamiento Patrim.", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio = 0 Then
            nEndeudaPatrim = 0
        Else
            nEndeudaPatrim = nPasivo / FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nEndeudaPatrim, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If nPasivo = 0 Then
                    sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat(nEndeudaPatrim / (nPasivo / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio), 8, 2)
                End If
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "RENTABILIDAD" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
        
    sCadImp = sCadImp & ofun.ImpreFormat("Rentab. Neta Patrim.", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio = 0 Then
            nRentabPatrim = 0
        Else    'Utilidad_Neta/Patrimonio
            nRentabPatrim = ((FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam)) / FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nRentabPatrim, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If ((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam)) = 0 Then
                    sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat(nRentabPatrim / (((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam)) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio), 8, 2)
                End If
            End If
        End If
    Next i
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ofun.ImpreFormat("Rentab. sobre Ventas", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
            nRentabVentas = 0
        Else    'Utilidad_Neta/Ventas
            nRentabVentas = ((FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam)) / FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas
        End If
        sCadImp = sCadImp & ofun.ImpreFormat(nRentabVentas, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
            Else
                If ((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam)) = 0 Then
                    sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat(nRentabVentas / (((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam)) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas), 8, 2)
                End If
            End If
        End If
    Next i
    

'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'
'    'Activo Circulante / Pasivo
'    Dim nActCir As Double
'    Dim nPasivo As Double
'    Dim nActCirAnt As Double
'    Dim nPasivoAnt As Double
'    Dim Res1 As Double
'    Dim Res2 As Double
'    sCadImp = sCadImp & ofun.ImpreFormat("Liquidez", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        nActCir = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
'        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
'        If nPasivo = 0 Then
'            Res1 = 0
'        Else
'            Res1 = nActCir / nPasivo
'        End If
'        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            nActCirAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
'            nPasivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
'            If nPasivoAnt = 0 Then
'                Res2 = 0
'            Else
'                Res2 = nActCirAnt / nPasivoAnt
'            End If
'            If Res2 = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    'Activo Circulante - Pasivo
'    sCadImp = sCadImp & ofun.ImpreFormat("Capital Trabajo", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        nActCir = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
'        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
'        Res1 = nActCir - nPasivo
'        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            nActCirAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
'            nPasivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
'            Res2 = nActCirAnt - nPasivoAnt
'            If Res2 = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & "GESTION" & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    'Costo de Ventas / Inventario
'    sCadImp = sCadImp & ofun.ImpreFormat("Rotacion Inventario", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios = 0 Then
'            Res1 = 0
'        Else
'            Res1 = FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
'        End If
'        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
'
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios = 0 Then
'                Res2 = 0
'            Else
'                Res2 = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
'            End If
'            If Res2 = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'
'
'    'Recup Ctas x Cob / (Ventas / 30)
'    sCadImp = sCadImp & ofun.ImpreFormat("Periodo de Cobro", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
'            Res1 = 0
'        Else
'            Res1 = FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob / (FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas / 30)
'        End If
'        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
'                Res2 = 0
'            Else
'                Res2 = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas / 30)
'            End If
'            If Res2 = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'
'    Dim nIngresos As Double
'    Dim nEgresos As Double
'    Dim nIngresosAnt As Double
'    Dim nEgresosAnt As Double
'
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & "RENTABILIDAD" & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    'Utilidad Operat =  Ingresos - Egresos
'    'Utilidad Operat / Patrimonio
'    sCadImp = sCadImp & ofun.ImpreFormat("Rentab. Patrimonial", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        nIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
'        nEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos
'
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio = 0 Then
'            Res1 = 0
'        Else
'            Res1 = (nIngresos - nEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio
'        End If
'        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            nIngresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob
'            nEgresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio = 0 Then
'                Res2 = 0
'            Else
'                Res2 = (nIngresos - nEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio
'            End If
'
'            If Res2 = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'
'    Dim nTotActivo As Double
'    Dim nTotActivoAnt As Double
'    sCadImp = sCadImp & ofun.ImpreFormat("Rentab. Activo", 20)
'    'Utilidad Operativa / Total Activo
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        nIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
'        nEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos
'        nTotActivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo
'        If nTotActivo = 0 Then
'            Res1 = 0
'        Else
'            Res1 = (nIngresos - nEgresos) / nTotActivo
'        End If
'        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            nIngresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob
'            nEgresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos
'            nTotActivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo
'            If nTotActivoAnt = 0 Then
'                Res2 = 0
'            Else
'                Res2 = (nIngresosAnt - nEgresosAnt) / nTotActivoAnt
'            End If
'            If Res2 = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'
'    sCadImp = sCadImp & ofun.ImpreFormat("Rentab. Ventas", 20)
'    'Utilidad Operativa / Ventas
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        nIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
'        nEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
'            Res1 = 0
'        Else
'            Res1 = (nIngresos - nEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas
'        End If
'        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            nIngresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob
'            nEgresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
'                Res2 = 0
'            Else
'                Res2 = (nIngresosAnt - nEgresosAnt) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas
'            End If
'            If Res2 = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'
'    'Ventas Totales = Ventas + Recup Ctas x Cob
'    'Margen Utilidad = Ventas Totales - Costo Ventas / Ventas Totales
'    Dim nVentasTot As Double
'    Dim nVentasTotAnt As Double
'    sCadImp = sCadImp & ofun.ImpreFormat("Margen Utilidad", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        nVentasTot = FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
'        If nVentasTot = 0 Then
'            Res1 = 0
'        Else
'            Res1 = (nVentasTot - FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad) / nVentasTot
'        End If
'        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            nVentasTotAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
'            If nVentasTotAnt = 0 Then
'                Res2 = 0
'            Else
'                Res2 = (nVentasTotAnt - FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad) / nVentasTotAnt
'            End If
'            If Res2 = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & "ENDEUDAMIENTO" & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    sCadImp = sCadImp & ofun.ImpreFormat("Endeudamiento Pat.", 20)
'    'Endeudamiento = Pasivo / Patrimonio
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
'
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio = 0 Then
'            Res1 = 0
'        Else
'            Res1 = nPasivo / FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio
'        End If
'        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            nPasivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
'
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio = 0 Then
'                Res2 = 0
'            Else
'                Res2 = nPasivoAnt / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio
'            End If
'            If Res2 = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'
'    sCadImp = sCadImp & ofun.ImpreFormat("Endeudamiento", 20)
'    'Endeudamiento = Pasivo / TotalActivo
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        nTotActivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo
'        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
'
'        If nTotActivo = 0 Then
'            Res1 = 0
'        Else
'            Res1 = nPasivo / nTotActivo
'        End If
'        sCadImp = sCadImp & ofun.ImpreFormat(Res1, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            nTotActivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo
'            nPasivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
'            If nTotActivoAnt = 0 Then
'                Res2 = 0
'            Else
'                Res2 = nPasivoAnt / nTotActivoAnt
'            End If
'            If Res2 = 0 Then
'                sCadImp = sCadImp & ofun.ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ofun.ImpreFormat(Res1 / Res2, 8, 2)
'            End If
'        End If
'    Next i
'
'


'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************


    If pbCostoProd = True Then

        'IMPRIME COSTO DE PRODUCCION
        '-------------------------------------
        sCadImp = sCadImp & Chr$(12)
        sCadImp = sCadImp & String(100, "-") & Chr$(10)
        sCadImp = sCadImp & Space(30) & "COSTO DE PRODUCCION" & Chr$(10)
        sCadImp = sCadImp & String(100, "-") & Chr$(10)
        sCadImp = sCadImp & ofun.ImpreFormat("Evaluacion", 16)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).dFecEval, 14, 0)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Tipo Cultivo", 20)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).sTpoCultivo, 13)
        Next i

        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Destino", 20)

        Dim sDestino As String
        bProd = False
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            If FtesIngreso(pnIndex).CostoProd(i).nSiembra = 1 Then
    '            If Len(sDestino) = 0 Then
    '                sDestino = "SIEMBRA"
    '            Else
    '                sDestino = sDestino & ",SIEMBRA"
                    bProd = True
                    sCadImp = sCadImp & ofun.ImpreFormat("SIEMBRA", 13)
                    'sCadImp = sCadImp & Chr$(10)
    '            End If
             End If

            If FtesIngreso(pnIndex).CostoProd(i).nMantenimiento = 1 Then
    '            If Len(sDestino) = 0 Then
    '                sDestino = "MANTENIMIENTO"
    '            Else
    '                sDestino = sDestino & ",MANTENIMIENTO"
    '                sCadImp = sCadImp & Chr$(10)
    '            End If
                If bProd = True Then
                    sCadImp = sCadImp & Space(7) & ofun.ImpreFormat("MANTENIMIENTO", 13)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat("MANTENIMIENTO", 13)
                    bProd = True
                End If
                   ' sCadImp = sCadImp & Chr$(10)
            End If
            If FtesIngreso(pnIndex).CostoProd(i).nDesAgricola = 1 Then
    '             If Len(sDestino) = 0 Then
    '                sDestino = "DESAGRICOLA"
    '             Else
    '                sDestino = sDestino & ",DESAGRICOLA"
    '                sCadImp = sCadImp & Chr$(10)
    '             End If
                  If bProd = True Then
                    sCadImp = sCadImp & Space(7) & ofun.ImpreFormat("DESAGRICOLA", 13)
                  Else
                    sCadImp = sCadImp & ofun.ImpreFormat("DESAGRICOLA", 13)
                    bProd = True
                  End If
                    'sCadImp = sCadImp & Chr$(10)


            End If
            If FtesIngreso(pnIndex).CostoProd(i).nCosecha = 1 Then
    '            If Len(sDestino) = 0 Then
    '                sDestino = "COSECHA"
    '            Else
    '                sDestino = sDestino & ",COSECHA"
    '                sCadImp = sCadImp & Chr$(10)
    '            End If
                If bProd = True Then
                    sCadImp = sCadImp & Space(7) & ofun.ImpreFormat("COSECHA", 13)
                Else
                    sCadImp = sCadImp & ofun.ImpreFormat("COSECHA", 13)
                    bProd = True
                End If
                    'sCadImp = sCadImp & Chr$(10)

             End If
            If FtesIngreso(pnIndex).CostoProd(i).nPlanOtros = 1 Then
    '            If Len(sDestino) = 0 Then
    '                sDestino = "OTROS"
    '            Else
    '                sDestino = sDestino & ",OTROS"
    '            End If
                    If bProd = True Then
                        sCadImp = sCadImp & Space(7) & ofun.ImpreFormat("OTROS", 13)
                    Else
                        sCadImp = sCadImp & ofun.ImpreFormat("OTROS", 13)
                        bProd = True
                    End If
                    'sCadImp = sCadImp & Chr$(10)

            End If
            'sCadImp = sCadImp & oFun.ImpreFormat(sDestino, 60)
        Next i

        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Maquinaria", 12)

        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nMaquinaria, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Jornales", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nJornales, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Insumos", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nInsumos, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Pesticidas", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPesticidas, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Otros", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nOtros, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Total", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros, 13)
        Next i
        sCadImp = sCadImp & Chr$(10) & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Hectareas", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Produccion", 18)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(Format(FtesIngreso(pnIndex).CostoProd(i).nProduccion, "#0.00") & " " & FtesIngreso(pnIndex).CostoProd(i).sUniProd, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Precio Uni.", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPreUni, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Ingresos", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas * FtesIngreso(pnIndex).CostoProd(i).nPreUni * FtesIngreso(pnIndex).CostoProd(i).nProduccion, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Egresos", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas * (FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros), 13)
        Next i
        sCadImp = sCadImp & Chr$(10)

        sCadImp = sCadImp & ofun.ImpreFormat("Utilidad", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ofun.ImpreFormat((FtesIngreso(pnIndex).CostoProd(i).nHectareas * FtesIngreso(pnIndex).CostoProd(i).nPreUni * FtesIngreso(pnIndex).CostoProd(i).nProduccion) - (FtesIngreso(pnIndex).CostoProd(i).nHectareas * (FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros)), 13)
        Next i
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & String(100, "-") & Chr$(10)
 End If
    ' ---------------------------------------------------------------------------------------------------------------
    ' Muestra los datos de las referencias personales de la fte como persona ----------------------------------------
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "REFERENCIAS PERSONALES" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF

    sCadImp = sCadImp & FtesIngreso(pnIndex).sComentarios & Chr$(10)
    sCadImp = sCadImp & Chr$(10)

    ' Muestra los datos de las referencias bancarias de la fte como persona ----------------------------------------
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("REFERENCIAS BANCARIAS", 42, 0) & ofun.ImpreFormat("NRO.TARJETA", 22, 2) & ofun.ImpreFormat("NRO.CUENTA", 22, 2) & ofun.ImpreFormat("MONTO", 10, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF

    MatDatosFte = RecuperaRefBancariaPersFte(FtesIngreso(pnIndex).sFteIngreso)

    For i = 0 To UBound(MatDatosFte) - 1
        sCadImp = sCadImp & ofun.ImpreFormat(MatDatosFte(i, 0) & "-" & MatDatosFte(i, 1), 40) & ofun.ImpreFormat(MatDatosFte(i, 3), 22, 2) & ofun.ImpreFormat(MatDatosFte(i, 2), 22, 2) & ofun.ImpreFormat(MatDatosFte(i, 4), 10, 2) & Chr$(10)
    Next i

    sCadImp = sCadImp & Chr$(10)

   ' Muestra los datos de las referencias comerciales de la fte como persona ----------------------------------------
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ofun.ImpreFormat("REFERENCIAS COMERCIALES", 66, 0) & ofun.ImpreFormat("TIPO RELAC. COMERCIAL", 22, 2) & ofun.ImpreFormat("TELEFONO", 10, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF

    MatDatosComFte = RecuperaRefComercialPersFte(FtesIngreso(pnIndex).sFteIngreso)

    For i = 0 To UBound(MatDatosComFte) - 1
        sCadImp = sCadImp & ofun.ImpreFormat(MatDatosComFte(i, 3) & "-" & MatDatosComFte(i, 0), 64) & ofun.ImpreFormat(MatDatosComFte(i, 2), 22, 2) & ofun.ImpreFormat(MatDatosComFte(i, 1), 10, 2) & Chr$(10)
    Next i

    ' ---------------------------------------------------------------------------------------------------------------
    sCadImp = sCadImp & Chr$(10)

    GenerarImpresionFteIngresoIndependiente_CS = sCadImp
    
End Function

'******************************************************

Public Sub LimpiaEliminados()
Dim i, J As Integer
Dim bExisteEliminados As Boolean
    bExisteEliminados = True
    Do While bExisteEliminados
        bExisteEliminados = False
        For i = 0 To nNumFtes - 1
            If FtesIngreso(i).nTipoAct = 3 Then
                bExisteEliminados = True
                For J = i To nNumFtes - 2
                    FtesIngreso(J) = FtesIngreso(J + 1)
                Next J
                nNumFtes = nNumFtes - 1
                Exit For
            End If
        Next i
    Loop
End Sub

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
Dim oConecta As COMConecta.DCOMConecta
Dim R As ADODB.Recordset

On Error GoTo ErrorTiempoFteIngreso
    sSql = "Select nParamValor from " & gConsComunes & "ColocParametro where nParamVar = " & gColocParamFteIngresoTiempoEval
    Set oConecta = New COMConecta.DCOMConecta
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
    nValComDomicilio = 0
    sTelefonos = ""
    sTelefonos2 = ""
    sCiiu = ""
    sEstado = ""
    nTalla = 0
    nPeso = 0
    sTipoSangre = ""
    nTipoPersona = gPersonaNat
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
Dim oConecta  As COMConecta.DCOMConecta
          
     GoTo ErrorUltimoCorrelativo
    sSql = "Select MAX(substring(P.cPersCod,6,7)) as Correlativo from " & gConsPersona & "Persona P where substring(P.cPersCod,4,2) = '" & Right(psCodAge, 2) & "'"
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSql)
    If Not R.BOF And Not R.EOF Then
        UltimoCorrelativo = IIf(IsNull(R!Correlativo), 0, R!Correlativo)
    Else
        UltimoCorrelativo = 0
    End If
    R.Close
    
    
    If bPersDistrib Then
        sSql = "Select MAX(substring(P.cCodPers,3,7)) as Correlativo from Persona P where substring(P.cCodPers,1,2) = '" & Right(psCodAge, 2) & "'"
        R.CursorLocation = adUseClient
        R.Open sSql, ConexDitrib, adOpenStatic, adLockReadOnly, adCmdText
        If Not R.BOF And Not R.EOF Then
            If R!Correlativo > UltimoCorrelativo Then
                UltimoCorrelativo = R!Correlativo
            End If
        End If
        R.Close
        Set R = Nothing
        oConecta.CierraConexion
        Set oConecta = Nothing
    End If
    Exit Function
    
ErrorUltimoCorrelativo:
    Set R = Nothing
    Set oConecta = Nothing
    Err.Raise Err.Number, "ReucperaPersona", Err.Description
End Function
Public Sub GrabarCambiosPersona(ByVal psCodCmact As String, ByVal psCodAge As String)
Dim sSql As String
Dim InicioTran As Boolean
Dim i As Integer
Dim J As Integer
Dim oConecta  As COMConecta.DCOMConecta
'Dim oNegPers As NCOMPersona
Dim sCodPersDistrib As String
Dim lsUltCorr As String
Dim lnCorr As Long
Dim lnCodchequeo As Integer
Dim lsCadFinal As String
Dim sNumFuente As String

'************************************************************************
'PARA DISTRIBUIDO
'************************************************************************

Dim R As ADODB.Recordset
Dim sPersRelIDistrib As String
Dim TranDistribBegin As Boolean
Dim vDocs(4) As String
Dim sPersRelInst As String

On Error GoTo ErrorGrabarCambiosPersona

'************************************************************************
TranDistribBegin = False
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

    If nTipoPersona = gPersonaNat Then
        If sSexo = "F" And Len(Trim(sApeCas)) > 0 Then
            If sEstadoCivil = "3" Then
                sNombreCompleto = sApePat & "/" & sApeMat & "\VDA " & sApeCas & "," & sNombres
            Else
                sNombreCompleto = sApePat & "/" & sApeMat & "\" & sApeCas & "," & sNombres
            End If
        Else
            sNombreCompleto = sApePat & "/" & sApeMat & "," & sNombres
        End If
    End If
    
    InicioTran = True
    If nTipoAct = PersFilaNueva Then
        If ConexionActiva Is Nothing Then
            Set oConecta = New COMConecta.DCOMConecta
            oConecta.AbreConexion
            Set ConexionActiva = oConecta.ConexionActiva
        End If
    End If
    
    '**********************   Para Actualizar el Distribuido ************************
    sSql = "Select nConsSisValor from ConstSistema where nConsSisCod=45"
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
    'OBTENER CODIGO DE FUENTE DE INGRESO
    sSql = "Select MAX(cNumFuente) as cNumFuente From PersFteIngreso"
    Set R = New ADODB.Recordset
    R.Open sSql, ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    sNumFuente = Format(CLng(IIf(IsNull(R!cNumFuente), 0, R!cNumFuente)) + 1, "00000000")
    R.Close
    '************************************************************************************
    
    ConexionActiva.BeginTrans
    
    '**********************   Para Actualizar el Distribuido ************************
    If nTipoAct = PersFilaNueva Then
        'Set oNegPers = New NCOMPersona
        lsUltCorr = UltimoCorrelativo(psCodAge)
        lnCorr = Val(lsUltCorr) + 1
        lsUltCorr = FillNum(Trim(Str(lnCorr)), 7, "0")
        lnCodchequeo = DigitoChequeo(lsUltCorr)
        lsCadFinal = lsUltCorr & Trim(Str(lnCodchequeo))
        sCodPersDistrib = Right(Trim(psCodAge), 2) & Trim(lsCadFinal)
        sPersCod = Trim(psCodCmact) & Trim(psCodAge) & Trim(lsCadFinal)
        'Set oNegPers = Nothing
    Else
        sCodPersDistrib = Right(sPersCod, 10)
    End If
    '********************************************************************************
    Set R = Nothing
    
    sNombreCompleto = Replace(sNombreCompleto, "'", "''")
    
    If nTipoAct = PersFilaModificada Then
                
        'Actualiza Persona
        sSql = "UPDATE " & gConsPersona & "Persona SET "
        sSql = sSql & "cPersNombre = '" & sNombreCompleto & "',"
        sSql = sSql & "dPersNacCreac = '" & Format(dFechaNac, "mm/dd/yyyy") & "',"
        sSql = sSql & "cPersDireccUbiGeo = '" & sUbicGeografica & "',"
        sSql = sSql & "cPersDireccDomicilio = '" & sDomicilio & "',"
        sSql = sSql & "cPersDireccCondicion = '" & sCondicionDomic & "',"
        sSql = sSql & "nPersValComDomicilio = " & nValComDomicilio & ","
        sSql = sSql & "cPersTelefono = '" & sTelefonos & "',"
        sSql = sSql & "cPersTelefono2 = '" & sTelefonos2 & "',"
        sSql = sSql & "cPersEmail = '" & sEmail & "',"
        sSql = sSql & "nPersPersoneria = '" & Trim(Str(nTipoPersona)) & "',"
        sSql = sSql & "cPersCIIU = '" & sCiiu & "',"
        sSql = sSql & "cPersEstado = '" & sEstado & "',"
        sSql = sSql & "cPersCodSbs = '" & sPersCodSbs & "',"
        sSql = sSql & "nPersRelaInst = " & nPersRelInst
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
            sSql = sSql & "nPersTalla,nPersTpoSangre, cNacionalidad, nResidente, cUltimaActualizacion)"
            sSql = sSql & " VALUES('" & sPersCod & "',"
            sSql = sSql & "'" & sSexo & "',"
            sSql = sSql & sEstadoCivil & ","
            sSql = sSql & Trim(Str(nHijos)) & ","
            sSql = sSql & "'" & Format(nPeso, "#0.00") & "',"
            sSql = sSql & "'" & Format(nTalla, "#0.00") & "',"
            sSql = sSql & IIf(sTipoSangre = "", 0, sTipoSangre) & ","
            sSql = sSql & "'" & cNacionalidad & "',"
            sSql = sSql & nResidencia & ","
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
            sSql = sSql & ",cPersDireccCondicion, nPersValComDomicilio, cPersTelefono,cPersTelefono2,nPersPersoneria,cPersCIIU,cPersEstado"
            sSql = sSql & ",cPersEmail,cUltimaActualizacion, cPersCodSbs, nPersRelaInst,dPersIng)"
            sSql = sSql & "VALUES('" & sPersCod & "',"
            sSql = sSql & "'" & sNombreCompleto & "',"
            sSql = sSql & "'" & Format(dFechaNac, "mm/dd/yyyy") & "',"
            sSql = sSql & "'" & sUbicGeografica & "',"
            sSql = sSql & "'" & sDomicilio & "',"
            sSql = sSql & "'" & sCondicionDomic & "',"
            sSql = sSql & nValComDomicilio & ","
            sSql = sSql & "'" & sTelefonos & "',"
            sSql = sSql & "'" & sTelefonos2 & "',"
            sSql = sSql & Trim(Str(nTipoPersona)) & ","
            sSql = sSql & "'" & sCiiu & "',"
            sSql = sSql & "'" & sEstado & "',"
            sSql = sSql & "'" & sEmail & "',"
            sSql = sSql & "'" & sActualiza & "','" & sPersCodSbs & "'," & nPersRelInst & ",'" & Format(dFechaHoy, "mm/dd/yyyy") & "')"
            ConexionActiva.Execute sSql
            
            If nTipoPersona = gPersonaNat Then
                'Insertar en PersonaNat
                sSql = "INSERT INTO " & gConsPersona
                sSql = sSql & "PersonaNat(cPersCod,cPersnatSexo,nPersNatEstCiv,nPersNatHijos,nPersPeso,"
                sSql = sSql & "nPersTalla,nPersTpoSangre, cNacionalidad, nResidente, cUltimaActualizacion)"
                sSql = sSql & " VALUES('" & sPersCod & "',"
                sSql = sSql & "'" & sSexo & "',"
                sSql = sSql & sEstadoCivil & ","
                sSql = sSql & Trim(Str(nHijos)) & ","
                sSql = sSql & "'" & Format(nPeso, "#0.00") & "',"
                sSql = sSql & "'" & Format(nTalla, "#0.00") & "',"
                sSql = sSql & IIf(sTipoSangre = "", 0, sTipoSangre) & ","
                sSql = sSql & "'" & cNacionalidad & "',"
                sSql = sSql & nResidencia & ","
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
                If RFirma.RecordCount > 0 Then
                    RFirma!cPersCod = sPersCod
                    RFirma!cUltimaActualizacion = sActualiza
                    RFirma.Update
                End If
            End If
        
        End If
    End If
    
    'Graba Relaciones de la Persona
    For i = 0 To nNumRel - 1
        If Relac(i).nTipoAct = PersFilaModificada Then
            sSql = "UPDATE " & gConsPersona & "PersRelaciones SET "
            sSql = sSql & "cPersRelacPersCod = '" & Relac(i).sPersCod & "',"
            sSql = sSql & "nPersRelac = " & Trim(Right(Relac(i).sRelacion, 20)) & ","
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
    
    'Graba Referencias Comerciales
    For i = 0 To nNumRefCom - 1
        If RefCom(i).nTipoAct = PersFilaModificada Then
            sSql = "UPDATE PERSREFCOMERCIAL SET "
            sSql = sSql & "cNomRefCom= '" & RefCom(i).sRefComNom & "',"
            sSql = sSql & "nTipoRelaCom = " & Right(RefCom(i).sRefComTipo, 1) & ","
            sSql = sSql & "cFonoRefCom = '" & RefCom(i).sRefComFono & "',"
            sSql = sSql & "cUltimaActualizacion = '" & sActualiza & "' "
            sSql = sSql & " WHERE cPersCod = '" & sPersCod & "' AND nNumRefCom = " & RefCom(i).nRefComCod
            ConexionActiva.Execute sSql
        Else
            If RefCom(i).nTipoAct = PersFilaNueva Then
                sSql = "INSERT INTO PERSREFCOMERCIAL (cPersCod, nNumRefCom, cNomRefCom, nTipoRelaCom, "
                sSql = sSql & "cFonoRefCom, cUltimaActualizacion)"
                sSql = sSql & " VALUES ('" & sPersCod & "',"
                sSql = sSql & RefCom(i).nRefComCod & ","
                sSql = sSql & "'" & RefCom(i).sRefComNom & "',"
                sSql = sSql & Right(RefCom(i).sRefComTipo, 1) & ","
                sSql = sSql & "'" & RefCom(i).sRefComFono & "',"
                sSql = sSql & "'" & sActualiza & "')"
                ConexionActiva.Execute sSql
            Else
                If RefCom(i).nTipoAct = PersFilaEliminda Then
                    sSql = "DELETE PERSREFCOMERCIAL "
                    sSql = sSql & " WHERE cPersCod = '" & sPersCod & "' AND nNumRefCom = " & RefCom(i).nRefComCod
                    ConexionActiva.Execute sSql
                End If
            End If
        End If
    Next i
    
    'Graba Referencias Bancarias
    For i = 0 To nNumRefBan - 1
        If RefBan(i).nTipoAct = PersFilaModificada Then
            sSql = "UPDATE PERSREFBANCARIA SET "
            sSql = sSql & "cCodBanco = '" & Mid(RefBan(i).sRefBanCodIF, 4, Len(RefBan(i).sRefBanCodIF) - 3) & "',"
            sSql = sSql & "cNumCta = '" & RefBan(i).sRefBanNumCta & "',"
            sSql = sSql & "cNumTarjeta = '" & RefBan(i).sRefBanNumTar & "',"
            sSql = sSql & "nMontoLinCred = " & RefBan(i).nRefBanLinCred & ","
            sSql = sSql & "cUltimaActualizacion = '" & sActualiza & "' "
            sSql = sSql & " WHERE cPersCod = '" & sPersCod & "' AND cCodBanco = '" & Mid(RefBan(i).sRefBanCodIFOrg, 4, Len(RefBan(i).sRefBanCodIFOrg) - 3) & "'"
            ConexionActiva.Execute sSql
        Else
            If RefBan(i).nTipoAct = PersFilaNueva Then
                sSql = "INSERT INTO PERSREFBANCARIA (cPersCod, cCodBanco, cNumCta, cNumTarjeta, "
                sSql = sSql & "nMontoLinCred, cUltimaActualizacion)"
                sSql = sSql & " VALUES ('" & sPersCod & "',"
                sSql = sSql & "'" & Mid(RefBan(i).sRefBanCodIF, 4, Len(RefBan(i).sRefBanCodIF) - 3) & "',"
                sSql = sSql & "'" & RefBan(i).sRefBanNumCta & "',"
                sSql = sSql & "'" & RefBan(i).sRefBanNumTar & "',"
                sSql = sSql & RefBan(i).nRefBanLinCred & ","
                sSql = sSql & "'" & sActualiza & "')"
                ConexionActiva.Execute sSql
            Else
                If RefBan(i).nTipoAct = PersFilaEliminda Then
                    sSql = "DELETE PERSREFBANCARIA "
                    sSql = sSql & " WHERE cPersCod = '" & sPersCod & "' AND cCodBanco = '" & Mid(RefBan(i).sRefBanCodIFOrg, 4, Len(RefBan(i).sRefBanCodIFOrg) - 3) & "'"
                    ConexionActiva.Execute sSql
                End If
            End If
        End If
    Next i
    
    'Graba Patrimonio Vehicular
    For i = 0 To nNumPatVeh - 1
        If PatVeh(i).nTipoAct = PersFilaModificada Then
            sSql = "UPDATE PERSPATRIMONIO SET "
            sSql = sSql & "cMarca= '" & PatVeh(i).sPatVehMarca & "',"
            sSql = sSql & "cFecFabri = '" & PatVeh(i).sPatVehFecFab & "',"
            sSql = sSql & "nValorCom = " & PatVeh(i).nPatVehValCom & ","
            sSql = sSql & "nCondPatri = " & Right(PatVeh(i).sPatVehCondicion, 1) & ","
            sSql = sSql & "cUltimaActualizacion = '" & sActualiza & "' "
            sSql = sSql & " WHERE cPersCod = '" & sPersCod & "' AND nNumPatri= " & PatVeh(i).nPatVehCod
            ConexionActiva.Execute sSql
        Else
            If PatVeh(i).nTipoAct = PersFilaNueva Then
                sSql = "INSERT INTO PERSPATRIMONIO(cPersCod, nNumPatri, nCondPatri, cMarca, cFecFabrica, "
                sSql = sSql & "nValorCom, cUltimaActualizacion)"
                sSql = sSql & " VALUES ('" & sPersCod & "',"
                sSql = sSql & PatVeh(i).nPatVehCod & ","
                sSql = sSql & Right(PatVeh(i).sPatVehCondicion, 1) & ","
                sSql = sSql & "'" & PatVeh(i).sPatVehMarca & "',"
                sSql = sSql & "'" & PatVeh(i).sPatVehFecFab & "',"
                sSql = sSql & PatVeh(i).nPatVehValCom & ","
                sSql = sSql & "'" & sActualiza & "')"
                ConexionActiva.Execute sSql
            Else
                If PatVeh(i).nTipoAct = PersFilaEliminda Then
                    sSql = "DELETE PERSPATRIMONIO "
                    sSql = sSql & " WHERE cPersCod = '" & sPersCod & "' AND nNumPatri= " & PatVeh(i).nPatVehCod
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
        FtesIngreso(i).sRazSocDescrip = Replace(FtesIngreso(i).sRazSocDescrip, "'", "''")
        If FtesIngreso(i).nTipoAct = PersFilaNueva Then
            FtesIngreso(i).cNumFte = sNumFuente
            sSql = "INSERT INTO " & gConsPersona & "PersFteIngreso(cNumFuente, cPersCod, cPersFIPersCod, cPersFICargo,"
            sSql = sSql & "dPersFIinicio, cPersFIMoneda, cPersFIComentario, cRazSocUbiGeo, cRazSocDirecc, "
            sSql = sSql & " cRazSocTelef, cRazSocDescrip, nPersTipFte, cUltimaActualizacion,bCostoProd) "
            sSql = sSql & " VALUES('" & FtesIngreso(i).cNumFte & "','" & sPersCod & "',"
            sSql = sSql & "'" & FtesIngreso(i).sFteIngreso & "',"
            sSql = sSql & "'" & FtesIngreso(i).sCargo & "',"
            sSql = sSql & "'" & Format(FtesIngreso(i).dFteIngreso, "mm/dd/yyyy") & "',"
            sSql = sSql & Trim(FtesIngreso(i).sMoneda) & ","
            sSql = sSql & "'" & Trim(FtesIngreso(i).sComentarios) & "',"
            sSql = sSql & "'" & Right(FtesIngreso(i).sRazSocUbiGeo, 12) & "',"
            sSql = sSql & "'" & FtesIngreso(i).sRazSocDirecc & "',"
            sSql = sSql & "'" & FtesIngreso(i).sRazSocTelef & "',"
            sSql = sSql & "'" & FtesIngreso(i).sRazSocDescrip & "',"
            sSql = sSql & FtesIngreso(i).nTipo & ",'" & sActualiza & "'," & FtesIngreso(i).bCostoProd
            sSql = sSql & ")"
            ConexionActiva.Execute sSql
        Else
            If FtesIngreso(i).nTipoAct = PersFilaModificada Then
                sSql = "UPDATE " & gConsPersona & "PersFteIngreso SET "
                sSql = sSql & " cPersFIPersCod = '" & FtesIngreso(i).sFteIngreso & "',"
                sSql = sSql & " cPersFICargo = '" & FtesIngreso(i).sCargo & "',"
                sSql = sSql & " dPersFIInicio = '" & Format(FtesIngreso(i).dFteIngreso, "mm/dd/yyyy") & "',"
                sSql = sSql & " cPersFIMoneda = " & Trim(FtesIngreso(i).sMoneda) & ","
                sSql = sSql & " cPersFIComentario = '" & Trim(FtesIngreso(i).sComentarios) & "',"
                sSql = sSql & " cUltimaActualizacion = '" & sActualiza & "',"
                sSql = sSql & " cRazSocUbiGeo = '" & FtesIngreso(i).sRazSocUbiGeo & "',"
                sSql = sSql & " cRazSocDirecc = '" & FtesIngreso(i).sRazSocDirecc & "',"
                sSql = sSql & " cRazSocTelef = '" & FtesIngreso(i).sRazSocTelef & "',"
                sSql = sSql & " cRazSocDescrip = '" & FtesIngreso(i).sRazSocDescrip & "', "
                sSql = sSql & " bCostoProd = " & FtesIngreso(i).bCostoProd
                sSql = sSql & " WHERE cNumFuente = '" & FtesIngreso(i).cNumFte & "'"
                ConexionActiva.Execute sSql
            Else
                If FtesIngreso(i).nTipoAct = PersFilaEliminda Then
                    sSql = "DELETE PersFIIndependiente Where cNumFuente = '" & FtesIngreso(i).cNumFte & "' "
                    ConexionActiva.Execute sSql
                    
                    sSql = "DELETE PersFIDependiente Where cNumFuente = '" & FtesIngreso(i).cNumFte & "' "
                    ConexionActiva.Execute sSql
                        
                    sSql = "DELETE PersFICostoProd Where cNumFuente = '" & FtesIngreso(i).cNumFte & "' "
                    ConexionActiva.Execute sSql
                    
                    sSql = "DELETE " & gConsPersona & "PersFteIngreso "
                    sSql = sSql & " WHERE cNumFuente = '" & FtesIngreso(i).cNumFte & "'"
                    ConexionActiva.Execute sSql
                End If
            End If
        End If
    Next i
    
    'FUENTES DE INGRESO DEPENDIENTES
    For i = 0 To nNumFtes - 1
        For J = 0 To FtesIngreso(i).nNumFtesDep - 1
            If FtesIngreso(i).FtesIngresoDep(J).nTipoAct = PersFilaNueva Then
                sSql = "INSERT INTO PersFIDependiente(cNumFuente, dPersEval, nPersIngCli, "
                sSql = sSql & " nPersGastoFam, nPersIngCon, nPersOtrIng, dPersFICaduca) "
                sSql = sSql & " VALUES('" & FtesIngreso(i).cNumFte & "',"
                sSql = sSql & "'" & Format(FtesIngreso(i).FtesIngresoDep(J).dFecEval, "mm/dd/yyyy") & "',"
                sSql = sSql & Format(FtesIngreso(i).FtesIngresoDep(J).nIngresos, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).FtesIngresoDep(J).nGastosFam, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).FtesIngresoDep(J).nIngCony, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).FtesIngresoDep(J).nIngOtros, "#0.00") & ","
                sSql = sSql & "'" & Format(FtesIngreso(i).FtesIngresoDep(J).dFecCad, "mm/dd/yyyy") & "')"
                ConexionActiva.Execute sSql
            Else
                If FtesIngreso(i).FtesIngresoDep(J).nTipoAct = PersFilaModificada Then
                    sSql = "UPDATE PersFIDependiente SET "
                    sSql = sSql & " nPersIngCli = " & Format(FtesIngreso(i).FtesIngresoDep(J).nIngFam, "#0.00") & ","
                    sSql = sSql & " nPersGastoFam = " & Format(FtesIngreso(i).FtesIngresoDep(J).nGastosFam, "#0.00") & ","
                    sSql = sSql & " nPersIngCon = " & Format(FtesIngreso(i).FtesIngresoDep(J).nIngCony, "#0.00") & ","
                    sSql = sSql & " nPersOtrIng = " & Format(FtesIngreso(i).FtesIngresoDep(J).nIngOtros, "#0.00") & ","
                    sSql = sSql & " dPersFICaduca = '" & Format(FtesIngreso(i).FtesIngresoDep(J).dFecCad, "mm/dd/yyyy") & "' "
                    sSql = sSql & " Where cNumFuente = '" & FtesIngreso(i).cNumFte & "' AND dPersEval = '" & Format(FtesIngreso(i).FtesIngresoDep(J).dFecEval, "mm/dd/yyyy") & "'"
                    ConexionActiva.Execute sSql
                Else
                    If FtesIngreso(i).FtesIngresoDep(J).nTipoAct = PersFilaEliminda Then
                        sSql = "DELETE PersFIDependiente Where cNumFuente = '" & FtesIngreso(i).cNumFte & "' AND dPersEval = '" & Format(FtesIngreso(i).FtesIngresoDep(J).dFecEval, "mm/dd/yyyy") & "'"
                        ConexionActiva.Execute sSql
                    End If
                End If
            End If
        Next J
    Next i
    
    'FUENTES DE INGRESO INDEPENDIENTES
    For i = 0 To nNumFtes - 1
        For J = 0 To FtesIngreso(i).nNumFtesIndep - 1
            If FtesIngreso(i).FtesIngresoInd(J).nTipoAct = PersFilaNueva Then
                sSql = "INSERT INTO PersFIIndependiente(cNumFuente, dPersEval, "
                sSql = sSql & " dPersFICaduca, nPersFIActivoDisp, nPersFICtasxCobrar, nPersFIInventarios, "
                sSql = sSql & " nPersFIActivosFijos, nPersFIProveedores, nPersFICredCMACT, nPersFICredOtros, "
                sSql = sSql & " nPersFIVentas, nPersFIPatrimonio, nPersFICostoVentas, nPersFIRecupCtasXCobrar, "
                sSql = sSql & " nPersFIEgresosOtros, nPersIngFam, nPersEgrFam, cUltimaActualizacion)"
                sSql = sSql & " VALUES('" & FtesIngreso(i).cNumFte & "',"
                sSql = sSql & "'" & Format(FtesIngreso(i).FtesIngresoInd(J).dFecEval, "mm/dd/yyyy") & "',"
                sSql = sSql & "'" & Format(FtesIngreso(i).FtesIngresoInd(J).dFecCad, "mm/dd/yyyy") & "',"
                sSql = sSql & Format(FtesIngreso(i).FtesIngresoInd(J).nActivDisp, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).FtesIngresoInd(J).nCtasxCob, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).FtesIngresoInd(J).nInventarios, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).FtesIngresoInd(J).nActFijo, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).FtesIngresoInd(J).nProveedores, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).FtesIngresoInd(J).nCreditosCmact, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).FtesIngresoInd(J).nCreditosOtros, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).FtesIngresoInd(J).nVentas, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).FtesIngresoInd(J).nPasivoPatrimonio, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).FtesIngresoInd(J).nCompraMercad, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).FtesIngresoInd(J).nRecupCtasxCob, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).FtesIngresoInd(J).nOtrosEgresos, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).FtesIngresoInd(J).nPersIngFam, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).FtesIngresoInd(J).nPersEgrFam, "#0.00") & ","
                sSql = sSql & "'" & sActualiza & "')"
                ConexionActiva.Execute sSql
                
            Else
                If FtesIngreso(i).FtesIngresoInd(J).nTipoAct = PersFilaModificada Then
                    sSql = " UPDATE PersFIIndependiente SET "
                    sSql = sSql & " nPersFIActivoDisp = " & Format(FtesIngreso(i).FtesIngresoInd(J).nActivDisp, "#0.00") & ","
                    sSql = sSql & " nPersFICtasxCobrar = " & Format(FtesIngreso(i).FtesIngresoInd(J).nCtasxCob, "#0.00") & ","
                    sSql = sSql & " nPersFIInventarios = " & Format(FtesIngreso(i).FtesIngresoInd(J).nInventarios, "#0.00") & ","
                    sSql = sSql & " nPersFIActivosFijos = " & Format(FtesIngreso(i).FtesIngresoInd(J).nActFijo, "#0.00") & ","
                    sSql = sSql & " nPersFIProveedores = " & Format(FtesIngreso(i).FtesIngresoInd(J).nProveedores, "#0.00") & ","
                    sSql = sSql & " nPersFICredCMACT = " & Format(FtesIngreso(i).FtesIngresoInd(J).nCreditosCmact, "#0.00") & ","
                    sSql = sSql & " nPersFICredOtros = " & Format(FtesIngreso(i).FtesIngresoInd(J).nCreditosOtros, "#0.00") & ","
                    sSql = sSql & " nPersFIVentas = " & Format(FtesIngreso(i).FtesIngresoInd(J).nVentas, "#0.00") & ","
                    sSql = sSql & " nPersFIPatrimonio = " & Format(FtesIngreso(i).FtesIngresoInd(J).nPasivoPatrimonio, "#0.00") & ","
                    sSql = sSql & " nPersFICostoVentas = " & Format(FtesIngreso(i).FtesIngresoInd(J).nCompraMercad, "#0.00") & ","
                    sSql = sSql & " nPersFIRecupCtasXCobrar = " & Format(FtesIngreso(i).FtesIngresoInd(J).nRecupCtasxCob, "#0.00") & ","
                    sSql = sSql & " nPersFIEgresosOtros = " & Format(FtesIngreso(i).FtesIngresoInd(J).nOtrosEgresos, "#0.00") & ","
                    sSql = sSql & " nPersIngFam = " & Format(FtesIngreso(i).FtesIngresoInd(J).nPersIngFam, "#0.00") & ","
                    sSql = sSql & " nPersEgrFam = " & Format(FtesIngreso(i).FtesIngresoInd(J).nPersEgrFam, "#0.00") & ","
                    sSql = sSql & " dPersFICaduca = '" & Format(FtesIngreso(i).FtesIngresoInd(J).dFecCad, "mm/dd/yyyy") & "' "
                    sSql = sSql & " Where cNumFuente = '" & FtesIngreso(i).cNumFte & "' AND dPersEval = '" & Format(FtesIngreso(i).FtesIngresoInd(J).dFecEval, "mm/dd/yyyy") & "'"
                    
                    ConexionActiva.Execute sSql
                Else
                    If FtesIngreso(i).FtesIngresoInd(J).nTipoAct = PersFilaEliminda Then
                        sSql = "DELETE PersFIIndependiente Where cNumFuente = '" & FtesIngreso(i).cNumFte & "' AND dPersEval = '" & Format(FtesIngreso(i).FtesIngresoInd(J).dFecEval, "mm/dd/yyyy") & "'"
                        ConexionActiva.Execute sSql
                    End If
                End If
            End If
        Next J
    Next i
    
    'Actualiza Costos de Produccion
    For i = 0 To nNumFtes - 1
        For J = 0 To FtesIngreso(i).nNumCostosProd - 1
            If FtesIngreso(i).CostoProd(J).nTipoAct = PersFilaNueva Then
                sSql = "INSERT INTO PersFICostoProd(cNumFuente, dFecEval, nTpoCultivo, nMaquinaria, nJornales, nInsumos, nPesticidas, nOtros, nHectareas, "
                sSql = sSql & " nProduccion, nUniProd, nPreUni,nSiembra,nCosecha,nDesaAgricola,nPlanOtros,nMantenimiento)"
                sSql = sSql & " VALUES('" & FtesIngreso(i).cNumFte & "','" & Format(FtesIngreso(i).CostoProd(J).dFecEval, "mm/dd/yyyy") & "'," & FtesIngreso(i).CostoProd(J).nTpoCultivo & ","
                sSql = sSql & Format(FtesIngreso(i).CostoProd(J).nMaquinaria, "#0.00") & "," & Format(FtesIngreso(i).CostoProd(J).nJornales, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).CostoProd(J).nInsumos, "#0.00") & "," & Format(FtesIngreso(i).CostoProd(J).nPesticidas, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).CostoProd(J).nOtros, "#0.00") & "," & Format(FtesIngreso(i).CostoProd(J).nHectareas, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(i).CostoProd(J).nProduccion, "#0.00") & "," & Format(FtesIngreso(i).CostoProd(J).nUniProd, "#0.00") & "," & Format(FtesIngreso(i).CostoProd(J).nPreUni, "#0.00") & ","
                sSql = sSql & FtesIngreso(i).CostoProd(J).nSiembra & ","
                sSql = sSql & FtesIngreso(i).CostoProd(J).nCosecha & "," & FtesIngreso(i).CostoProd(J).nDesAgricola & "," & FtesIngreso(i).CostoProd(J).nPlanOtros & ","
                sSql = sSql & FtesIngreso(i).CostoProd(J).nMantenimiento & ")"
                                
                ConexionActiva.Execute sSql
                
            Else
                If FtesIngreso(i).CostoProd(J).nTipoAct = PersFilaModificada Then
                    sSql = "UPDATE PersFICostoProd SET "
                    sSql = sSql & "nTpoCultivo = " & FtesIngreso(i).CostoProd(J).nTpoCultivo & ","
                    sSql = sSql & "nMaquinaria = " & Format(FtesIngreso(i).CostoProd(J).nMaquinaria, "#0.00") & ","
                    sSql = sSql & "nJornales = " & Format(FtesIngreso(i).CostoProd(J).nJornales, "#0.00") & ","
                    sSql = sSql & "nInsumos = " & Format(FtesIngreso(i).CostoProd(J).nInsumos, "#0.00") & ","
                    sSql = sSql & "nPesticidas = " & Format(FtesIngreso(i).CostoProd(J).nPesticidas, "#0.00") & ","
                    sSql = sSql & "nOtros = " & Format(FtesIngreso(i).CostoProd(J).nOtros, "#0.00") & ","
                    sSql = sSql & "nHectareas = " & Format(FtesIngreso(i).CostoProd(J).nHectareas, "#0.00") & ","
                    sSql = sSql & "nProduccion = " & Format(FtesIngreso(i).CostoProd(J).nProduccion, "#0.00") & ","
                    sSql = sSql & "nUniProd = " & FtesIngreso(i).CostoProd(J).nUniProd & ","
                    sSql = sSql & "nPreUni = " & Format(FtesIngreso(i).CostoProd(J).nPreUni, "#0.00") & ","
                    sSql = sSql & "nSiembra=" & FtesIngreso(i).CostoProd(J).nSiembra & ","
                    sSql = sSql & "nCosecha=" & FtesIngreso(i).CostoProd(J).nMantenimiento & ","
                    sSql = sSql & "nDesaAgricola=" & FtesIngreso(i).CostoProd(J).nDesAgricola & ","
                    sSql = sSql & "nPlanOtros=" & FtesIngreso(i).CostoProd(J).nPlanOtros & ","
                    sSql = sSql & "nMantenimiento= " & FtesIngreso(i).CostoProd(J).nMantenimiento & " "
                    sSql = sSql & " WHERE cNumFuente = '" & FtesIngreso(i).cNumFte & "' AND dFecEval = '" & Format(FtesIngreso(i).CostoProd(J).dFecEval, "mm/dd/yyyy") & "'"
                    
                    ConexionActiva.Execute sSql
                Else
                    If FtesIngreso(i).CostoProd(J).nTipoAct = PersFilaEliminda Then
                        sSql = "DELETE PersFICostoProd Where cNumFuente = '" & FtesIngreso(i).cNumFte & "'"
                        ConexionActiva.Execute sSql
                    End If
                End If
            End If
        Next J
    Next i
    
    
    ConexionActiva.CommitTrans
    InicioTran = False
    
    '*******************************************************************************
    'Replicar Persona a Distribuido
    '*******************************************************************************
    If bPersDistrib Then
        ConexDitrib.Execute "SET DATEFORMAT MDY"
        ConexDitrib.BeginTrans
        TranDistribBegin = True
        If nTipoAct <> PersFilaNueva Then
            '*********************************************************************************
            'Para Actualizar Persona Distribuida
            '*********************************************************************************
            If bPersDistrib Then
    
'                sSQL = "UPDATE Persona SET "
'                sSQL = sSQL & " cNomPers = '" & sNombreCompleto & "',"
'                sSQL = sSQL & " cTidoci = '" & vDocs(0) & "',"
'                sSQL = sSQL & " cNudoci = '" & vDocs(1) & "',"
'                sSQL = sSQL & " cTidoTr = '" & vDocs(2) & "',"
'                sSQL = sSQL & " cNudoTr = '" & vDocs(3) & "',"
'                sSQL = sSQL & " cDirPers = '" & sDomicilio & "',"
'                sSQL = sSQL & " cCodZon = '" & sUbicGeografica & "',"
'                sSQL = sSQL & " dFecNac = '" & Format(dFechaNac, "dd/mm/yyyy") & "',"
'                sSQL = sSQL & " cTelPers = '" & sTelefonos & "',"
'                sSQL = sSQL & " cEstPers = '1',"
'                sSQL = sSQL & " cTipPers = '" & IIf(nTipoPersona = 0, 0, nTipoPersona - 1) & "',"
'                sSQL = sSQL & " cCodSBS = '" & sPersCodSbs & "',"
'                sSQL = sSQL & " dIngPers = '" & Mid(sActualiza, 1, 4) & "/" & Mid(sActualiza, 5, 2) & "/" & Mid(sActualiza, 7, 2) & "',"
'                sSQL = sSQL & " dUltAct = '" & Mid(sActualiza, 1, 4) & "/" & Mid(sActualiza, 5, 2) & "/" & Mid(sActualiza, 7, 2) & "',"
'                sSQL = sSQL & " cRelIns = '" & sPersRelIDistrib & "',"
'                sSQL = sSQL & " cConDom = '" & sCondicionDomic & "' "
'                sSQL = sSQL & " Where cCodpers = '" & sCodPersDistrib & "'"
'
'                ConexDitrib.Execute sSQL
            End If
    
                '******************************************************************************
                'ACTUALIZA PERSONA NATURAL EN DISTRIBUIDO
                '******************************************************************************
                If bPersDistrib Then
    
'                    sSQL = "Update PersonaNat SET "
'                    sSQL = sSQL & " cSexPers = '" & sSexo & "',"
'                    sSQL = sSQL & " cEstCiv = '" & sEstadoCivil & "',"
'                    sSQL = sSQL & " nNumHij = " & nHijos
'                    sSQL = sSQL & " Where cCodPers = '" & sCodPersDistrib & "'"
'
'                    'sSql = "INSERT INTO Personanat(cCodPers, cSexPers, cEstCiv, cCodProf, cCodCon, nNumHij)"
'                    'sSql = sSql & " VALUES('" & sCodPersDistrib & "','" & sSexo & "','" & sEstadoCivil & "','',''," & nHijos & ")"
'                    ConexDitrib.Execute sSQL
    
                End If
    
                '******************************************************************************
                'ACTUALIZA PERSONA JURIDICA EN DISTRIBUIDO
                '******************************************************************************
                If bPersDistrib Then
    
'                    sSQL = "UPDATE PersonaJur SET "
'                    sSQL = sSQL & " cSigla = '" & sSiglas & "' "
'                    sSQL = sSQL & " Where cCodPers = '" & sCodPersDistrib & "'"
'
'                    'sSql = "Insert into PersonaJur(cCodPers,cTipPerJur,cCodRegPub,cMagEmp,cSigla) "
'                    'sSql = sSql & "Values('" & sCodPersDistrib & "','0000','','','" & sSiglas & "')"
'                    ConexDitrib.Execute sSQL
    
                End If
        Else
    
                '*********************************************************************************
                'Para INSERTAR Persona Distribuida
                '*********************************************************************************
                If bPersDistrib Then
                    sSql = "INSERT INTO Persona(cCodPers,cNomPers,cTidoci,cNudoci,cTidotr,cNudotr,cDirPers, cCodZon, "
                    sSql = sSql & " dFecNac, cTelPers, cEstPers, cTipPers, cCodSBS, dIngPers, dUltAct, cCodUsu, cRelIns, cConDom, cFlagConsol, cEMail, cWeb, cTelefono2) "
                    sSql = sSql & " VALUES('" & sCodPersDistrib & "','" & sNombreCompleto & "','" & vDocs(0) & "','" & vDocs(1) & "','" & vDocs(2) & "','" & vDocs(3) & "','"
                    sSql = sSql & sDomicilio & "','" & sUbicGeografica & "','" & Format(dFechaNac, "mm/dd/yyyy") & "','" & sTelefonos & "','1','"
                    sSql = sSql & CInt(nTipoPersona) & "','" & sPersCodSbs & "','" & Mid(sActualiza, 1, 4) & "/" & Mid(sActualiza, 5, 2) & "/" & Mid(sActualiza, 7, 2) & "','" & Mid(sActualiza, 1, 4) & "/" & Mid(sActualiza, 5, 2) & "/" & Mid(sActualiza, 7, 2)
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
        TranDistribBegin = False
    End If
    
    nTipoAct = PersFilaSinCambios
    i = 0
    Do While i <= nNumDoc - 1
        If Docs(i).nTipoAct = PersFilaEliminda Then
            For J = i To nNumDoc - 2
                Docs(i) = Docs(i + 1)
            Next J
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
            For J = i To nNumRel - 2
                Relac(i) = Relac(i + 1)
            Next J
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
            For J = i To nNumFtes - 2
                FtesIngreso(i) = FtesIngreso(i + 1)
            Next J
            nNumFtes = nNumFtes - 1
            ReDim Preserve FtesIngreso(nNumFtes)
            i = i - 1
        Else
            FtesIngreso(i).nTipoAct = PersFilaSinCambios
        End If
        i = i + 1
    Loop
    
    i = 0
    Do While i <= nNumRefCom - 1
        If RefCom(i).nTipoAct = PersFilaEliminda Then
            For J = i To nNumRefCom - 2
                RefCom(i) = RefCom(i + 1)
            Next J
            nNumRefCom = nNumRefCom - 1
            ReDim Preserve RefCom(nNumRefCom)
            i = i - 1
        Else
            RefCom(i).nTipoAct = PersFilaSinCambios
        End If
        i = i + 1
    Loop
    
    i = 0
    Do While i <= nNumRefBan - 1
        If RefBan(i).nTipoAct = PersFilaEliminda Then
            For J = i To nNumRefBan - 2
                RefBan(i) = RefBan(i + 1)
            Next J
            nNumRefBan = nNumRefBan - 1
            ReDim Preserve RefBan(nNumRefBan)
            i = i - 1
        Else
            RefBan(i).nTipoAct = PersFilaSinCambios
        End If
        i = i + 1
    Loop
    
    Exit Sub
    
    i = 0
    Do While i <= nNumPatVeh - 1
        If PatVeh(i).nTipoAct = PersFilaEliminda Then
            For J = i To nNumPatVeh - 2
                PatVeh(i) = PatVeh(i + 1)
            Next J
            nNumPatVeh = nNumPatVeh - 1
            ReDim Preserve PatVeh(nNumPatVeh)
            i = i - 1
        Else
            PatVeh(i).nTipoAct = PersFilaSinCambios
        End If
        i = i + 1
    Loop
    
    Exit Sub
ErrorGrabarCambiosPersona:
    If InicioTran Then
        ConexionActiva.RollbackTrans
    End If
    If TranDistribBegin Then
        ConexDitrib.RollbackTrans
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
        If Docs(i).nTipo = pnTipo And Docs(i).sNumero = psNumero Then
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

Public Sub AdicionaRefComercial()
    nNumRefCom = nNumRefCom + 1
    ReDim Preserve RefCom(nNumRefCom)
    RefCom(nNumRefCom).nTipoAct = PersFilaNueva
End Sub

Public Sub AdicionaRefBancaria()
    nNumRefBan = nNumRefBan + 1
    ReDim Preserve RefBan(nNumRefBan)
    RefBan(nNumRefBan).nTipoAct = PersFilaNueva
End Sub

Public Sub AdicionaPatVehicular()
    nNumPatVeh = nNumPatVeh + 1
    ReDim Preserve PatVeh(nNumPatVeh)
    PatVeh(nNumPatVeh).nTipoAct = PersFilaNueva
End Sub

Public Sub AdicionaFteIngreso(ByVal pnTipoFte As PersFteIngresoTipo, Optional ByVal pbCostoProd As Boolean)
    nNumFtes = nNumFtes + 1
    ReDim Preserve FtesIngreso(nNumFtes)
    FtesIngreso(nNumFtes - 1).nTipoAct = PersFilaNueva
    If pnTipoFte = gPersFteIngresoTipoDependiente Then
        ReDim FtesIngreso(nNumFtes - 1).FtesIngresoDep(1)
        FtesIngreso(nNumFtes - 1).nNumFtesDep = 1
        FtesIngreso(nNumFtes - 1).FtesIngresoDep(0).nTipoAct = PersFilaNueva
    Else
        ReDim FtesIngreso(nNumFtes - 1).FtesIngresoInd(1)
        FtesIngreso(nNumFtes - 1).nNumFtesIndep = 1
        FtesIngreso(nNumFtes - 1).FtesIngresoInd(0).nTipoAct = PersFilaNueva
        
        If pbCostoProd = True Then
            ReDim FtesIngreso(nNumFtes - 1).CostoProd(1)
            FtesIngreso(nNumFtes - 1).nNumCostosProd = 1
            FtesIngreso(nNumFtes - 1).CostoProd(0).nTipoAct = PersFilaNueva
        End If
    End If
End Sub

Public Sub AdicionaFteIngresoCostoProd(ByVal pnIndex As Integer)
    FtesIngreso(pnIndex).nNumCostosProd = FtesIngreso(pnIndex).nNumCostosProd + 1
    ReDim Preserve FtesIngreso(pnIndex).CostoProd(FtesIngreso(pnIndex).nNumCostosProd)
End Sub

Public Sub AdicionaFteIngresoDependiente(ByVal pnIndex As Integer)
    FtesIngreso(pnIndex).nNumFtesDep = FtesIngreso(pnIndex).nNumFtesDep + 1
    ReDim Preserve FtesIngreso(pnIndex).FtesIngresoDep(FtesIngreso(pnIndex).nNumFtesDep)
End Sub

Public Sub AdicionaFteIngresoIndependiente(ByVal pnIndex As Integer)
    FtesIngreso(pnIndex).nNumFtesIndep = FtesIngreso(pnIndex).nNumFtesIndep + 1
    ReDim Preserve FtesIngreso(pnIndex).FtesIngresoInd(FtesIngreso(pnIndex).nNumFtesIndep)
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
Dim Conn As COMConecta.DCOMConecta
Dim sSql As String
Dim R As ADODB.Recordset

    sSql = "Select C.cConsDescripcion,PI.cPersCod,PI.cPersIDtpo,PI.cPersIDnro from PersID PI left join Constante C ON PI.cPersIDtpo = C.cConsValor"
    sSql = sSql & " where PI.cPersCod = '" & psCodPers & "' And C.cConsCod = " & Trim(Str(gPersIdTipo))
    Set R = Conn.CargaRecordSet(sSql, adLockReadOnly)
    Do While Not R.EOF
        Call AdicionaDocumento(R!cPersIDTpo, R!cPersIdNro, R!cConsDescripcion)
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Conn.CierraConexion
    Set Conn = Nothing
End Sub

Private Sub Class_Initialize()
Dim CIni As COMConecta.DCOMClasIni
    Set CIni = New COMConecta.DCOMClasIni
    gConsPersona = CIni.BasePersonas
    gConsComunes = CIni.BaseComunes
    gConsImagenes = CIni.BaseImagenes
    
    Set CIni = Nothing
    nNumDoc = 0
    nNumRel = 0
    
    Set Conn = New COMConecta.DCOMConecta
    If Conn.AbreConexion = False Then
        Err.Raise Err.Number, "Error de Conexion", Err.Description
    End If
    
End Sub
Private Function BuscaNombre(ByVal psNombre As String, ByVal nTipoBusqueda As TiposBusquedaNombre) As String
Dim sCadTmp As String
Dim PosIni As Integer
Dim PosFin As Integer
Dim PosIni2 As Integer
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
                PosIni2 = InStr(1, psNombre, "VDA")
                If PosIni2 <> 0 Then
                    PosIni = PosIni2 + 3
                    PosFin = InStr(1, psNombre, ",")
                    If PosFin = 0 Then
                        PosFin = Len(psNombre)
                    End If
                Else
                    PosIni = PosIni + 1
                    PosFin = InStr(1, psNombre, ",")
                    If PosFin = 0 Then
                        PosFin = Len(psNombre)
                    End If
                End If
                sCadTmp = Trim(Mid(psNombre, PosIni, PosFin - PosIni))
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

Public Function RecuperaFirma(ByVal psCodPers As String) As ADODB.Recordset
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim sCadCon As String

    'sSQL = "Provider=SQLOLEDB.1;User ID=dbaccess;Initial Catalog=DbCmactAux;Data Source=192.168.1.16;PWD=proyecto"
        
    If ConexionActiva Is Nothing Then
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        Set ConexionActiva = oConecta.ConexionActiva
    End If
    
    sCadCon = ""
    sSql = "Select * from Servidor Where RIGHT(cCodAge,2) = '" & sCodage & "' AND cNroSer = '05'"
    Set R = New ADODB.Recordset
    R.Open sSql, ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    If R.RecordCount > 0 Then
        sCadCon = "Provider=" & Trim(R!cProvider) & ";User ID=" & Trim(R!cLogin) & ";Initial Catalog=" & Trim(R!cDataBase) & ";Data Source=" & Trim(R!cNomSer) & ";PWD=" & Trim(R!cPass)
    End If
    R.Close
    
    If sCadCon <> "" Then
        Set oConecta = New COMConecta.DCOMConecta
        Call oConecta.AbreConexion(sCadCon)
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
        Docs(nNumDoc - 1).sTipo = R!cConsDescripcion & Space(75) & R!nConsValor
        Docs(nNumDoc - 1).nTipo = CInt(R!cPersIDTpo)
        Docs(nNumDoc - 1).nTipoOrig = CInt(R!cPersIDTpo)
        Docs(nNumDoc - 1).sNumero = Trim(R!cPersIdNro)
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
        Relac(nNumRel - 1).sRelacion = R!cConsDescripcion & Space(75) & R!nPersRelac
        Relac(nNumRel - 1).sPersRelacBenef = R!cConsDescripcion2 & Space(75) & R!nConsValor2
        'Relac(nNumRel - 1).sPersRelacEstado = R!cConsDescripcion4 & Space(50) & Format(IIf(IsNull(R!bPersRelacEstado), 0, R!bPersRelacEstado), "#0")
        Relac(nNumRel - 1).nPersRelacBenefPorc = IIf(IsNull(R!nPersRelacBenefPorc), 0, R!nPersRelacBenefPorc)
        Relac(nNumRel - 1).sPersRelacAMP = R!cConsDescripcion3 & Space(75) & Format(IIf(IsNull(R!bPersRelacAMP), 0, R!bPersRelacAMP), "#0")
        Relac(nNumRel - 1).nTipoAct = 0
        Relac(nNumRel - 1).sMovAct = R!cMovnro
        R.MoveNext
    Loop
    R.Close
End Sub

Private Sub RecuperaRefComercial(ByVal psPersCod As String)
Dim sSql As String
Dim R As Recordset
nNumRefCom = 0
    
    ReDim RefCom(nNumRefCom)

    sSql = " SELECT nNumRefCom, cNomRefCom, nTipoRelaCom, cConsDescripcion, cFonoRefCom "
    sSql = sSql & " FROM PERSREFCOMERCIAL RC INNER JOIN CONSTANTE C ON RC.nTipoRelaCom = C.nConsValor "
    sSql = sSql & " WHERE nConsCod = 3028 AND cPersCod = '" & psPersCod & "'"
    sSql = sSql & " ORDER BY nNumRefCom"
    Set R = Conn.CargaRecordSet(sSql)
    
    Do While Not R.EOF
        nNumRefCom = nNumRefCom + 1
        ReDim Preserve RefCom(nNumRefCom)
        RefCom(nNumRefCom - 1).sRefComNom = R!cNomRefCom
        RefCom(nNumRefCom - 1).sRefComFono = R!cFonoRefCom
        RefCom(nNumRefCom - 1).sRefComTipo = R!cConsDescripcion & Space(75) & R!nTipoRelaCom
        RefCom(nNumRefCom - 1).nRefComCod = R!nNumRefCom
        RefCom(nNumRefCom - 1).nTipoAct = 0
        R.MoveNext
    Loop
    
    Set R = Nothing
    
End Sub

Public Function RecuperaRefComercialPersFte(ByVal psPersFteCod As String) As Variant
Dim sSql As String
Dim R As Recordset
Dim nPunt As Integer
Dim MatDatos() As String

    sSql = " SELECT nNumRefCom, cNomRefCom, nTipoRelaCom, cConsDescripcion, cFonoRefCom "
    sSql = sSql & " FROM PERSREFCOMERCIAL RC INNER JOIN CONSTANTE C ON RC.nTipoRelaCom = C.nConsValor "
    sSql = sSql & " WHERE nConsCod = 3028 AND cPersCod = '" & psPersFteCod & "'"
    sSql = sSql & " ORDER BY nNumRefCom"
    Set R = Conn.CargaRecordSet(sSql)

    ReDim MatDatos(R.RecordCount, 5)
    nPunt = 0
    Do While Not R.EOF
        MatDatos(nPunt, 0) = R!cNomRefCom
        MatDatos(nPunt, 1) = R!cFonoRefCom
        MatDatos(nPunt, 2) = R!cConsDescripcion & Space(75) & R!nTipoRelaCom
        MatDatos(nPunt, 3) = R!nNumRefCom
        MatDatos(nPunt, 4) = 0
        nPunt = nPunt + 1
        R.MoveNext
    Loop
    Set R = Nothing

    RecuperaRefComercialPersFte = MatDatos

End Function

Private Sub RecuperaMaxRefComercial(ByVal psPersCod As String)
Dim sSql As String
Dim R As Recordset

    sSql = "SELECT MAX(nNumRefCom) as nMaxRefCom FROM PERSREFCOMERCIAL WHERE cPersCod = '" & psPersCod & "'"
    Set R = Conn.CargaRecordSet(sSql)
    
    If Not R.EOF And Not R.BOF Then
        nMaxRefCom = IIf(IsNull(R!nMaxRefCom), 0, R!nMaxRefCom)
    Else
        nMaxRefCom = 0
    End If

    Set R = Nothing
    
End Sub

Private Sub RecuperaPatVehicular(ByVal psPersCod As String)
Dim sSql As String
Dim R As Recordset
nNumPatVeh = 0
    
    ReDim PatVeh(nNumPatVeh)

    sSql = " SELECT nNumPatri, nCondPatri, cMarca, cFecFabrica, nValorCom, cConsDescripcion "
    sSql = sSql & " FROM PERSPATRIMONIO P INNER JOIN CONSTANTE C ON P.nCondPatri = C.nConsValor "
    sSql = sSql & " WHERE nConsCod = 3029 AND cPersCod = '" & psPersCod & "'"
    sSql = sSql & " ORDER BY nNumPatri"
    Set R = Conn.CargaRecordSet(sSql)
    
    Do While Not R.EOF
        nNumPatVeh = nNumPatVeh + 1
        ReDim Preserve PatVeh(nNumPatVeh)
        PatVeh(nNumPatVeh - 1).nPatVehCod = R!nNumPatri
        PatVeh(nNumPatVeh - 1).sPatVehMarca = R!cMarca
        PatVeh(nNumPatVeh - 1).sPatVehFecFab = R!cFecFabrica
        PatVeh(nNumPatVeh - 1).nPatVehValCom = R!nValorCom
        PatVeh(nNumPatVeh - 1).sPatVehCondicion = R!cConsDescripcion & Space(75) & R!nCondPatri
        PatVeh(nNumPatVeh - 1).nTipoAct = 0
        R.MoveNext
    Loop
    
    Set R = Nothing
    
End Sub

Private Sub RecuperaMaxPatVehicular(ByVal psPersCod As String)
Dim sSql As String
Dim R As Recordset

    sSql = "SELECT MAX(nNumPatri) as nMaxPatVeh FROM PERSPATRIMONIO WHERE cPersCod = '" & psPersCod & "'"
    Set R = Conn.CargaRecordSet(sSql)
    
    If Not R.EOF And Not R.BOF Then
        nMaxPatVeh = IIf(IsNull(R!nMaxPatVeh), 0, R!nMaxPatVeh)
    Else
        nMaxPatVeh = 0
    End If

    Set R = Nothing
    
End Sub

Private Sub RecuperaRefBancaria(ByVal psPersCod As String)
Dim sSql As String
Dim R As Recordset
nNumRefBan = 0
    
    ReDim RefBan(nNumRefBan)

    sSql = "SELECT I.cIFTpo + '.' + P.cPersCod cCodBanco, CONVERT(CHAR(30),P.cPersNombre ) AS cBanNombre, "
    sSql = sSql & "cNumCta , cNumTarjeta, nMontoLinCred "
    sSql = sSql & " FROM PERSREFBANCARIA RB INNER JOIN PERSONA P ON RB.cCodBanco = P.cPersCod "
    sSql = sSql & " INNER JOIN INSTITUCIONFINANC I ON P.cPersCod = I.cPersCod "
    sSql = sSql & " WHERE RB.cPersCod = '" & psPersCod & "'"
    Set R = Conn.CargaRecordSet(sSql)
    
    Do While Not R.EOF
        nNumRefBan = nNumRefBan + 1
        ReDim Preserve RefBan(nNumRefBan)
        RefBan(nNumRefBan - 1).sRefBanCodIFOrg = R!cCodBanco
        RefBan(nNumRefBan - 1).sRefBanCodIF = R!cCodBanco
        RefBan(nNumRefBan - 1).sRefBanNombre = R!cBanNombre
        RefBan(nNumRefBan - 1).sRefBanNumCta = R!cNumCta
        RefBan(nNumRefBan - 1).sRefBanNumTar = R!cNumTarjeta
        RefBan(nNumRefBan - 1).nRefBanLinCred = R!nMontoLinCred
        RefBan(nNumRefBan - 1).nTipoAct = 0
        R.MoveNext
    Loop
    
    Set R = Nothing
    
End Sub

Public Function RecuperaRefBancariaPersFte(ByVal psPersFteCod As String) As Variant
Dim sSql As String
Dim R As ADODB.Recordset
Dim nPunt As Integer
Dim MatDatos() As String

    sSql = "SELECT I.cIFTpo + '.' + P.cPersCod cCodBanco, CONVERT(CHAR(30),P.cPersNombre ) AS cBanNombre, "
    sSql = sSql & "cNumCta , cNumTarjeta, nMontoLinCred "
    sSql = sSql & " FROM PERSREFBANCARIA RB INNER JOIN PERSONA P ON RB.cCodBanco = P.cPersCod "
    sSql = sSql & " INNER JOIN INSTITUCIONFINANC I ON P.cPersCod = I.cPersCod "
    sSql = sSql & " WHERE RB.cPersCod = '" & psPersFteCod & "'"
    
    Set R = Conn.CargaRecordSet(sSql)

    ReDim MatDatos(R.RecordCount, 6)
    nPunt = 0
    Do While Not R.EOF
        MatDatos(nPunt, 0) = R!cCodBanco
        MatDatos(nPunt, 1) = R!cBanNombre
        MatDatos(nPunt, 2) = R!cNumCta
        MatDatos(nPunt, 3) = R!cNumTarjeta
        MatDatos(nPunt, 4) = Format(R!nMontoLinCred, "#0.00")
        MatDatos(nPunt, 5) = 0
        nPunt = nPunt + 1
        R.MoveNext
    Loop
    Set R = Nothing
    
    RecuperaRefBancariaPersFte = MatDatos
End Function

'Public Function FuenteIngresoAsignadaACredito(ByVal pnIndex As Integer, ByVal pdFecEval As Date) As Boolean
'Dim sSql As String
'Dim R As ADODB.Recordset
'
'    sSql = "Select * from ColocFteIngreso where cNumFuente = '" & FtesIngreso(pnIndex).cNumFte & "' AND dPersEval = '" & Format(pdFecEval, "mm/dd/yyyy") & "'"
'    Set R = Conn.CargaRecordSet(sSql)
'    If R.RecordCount > 0 Then
'        FuenteIngresoAsignadaACredito = True
'    Else
'        FuenteIngresoAsignadaACredito = False
'    End If
'    Set R = Nothing
'End Function
'ARCV 14-08-2006
Public Function FuenteIngresoAsignadaACredito(ByVal pcNumFuente As String, _
                                              Optional ByVal pdFecEval As Date = 0) As Boolean
Dim sSql As String
Dim R As ADODB.Recordset
    
    sSql = "Select * from ColocFteIngreso where cNumFuente = '" & pcNumFuente & _
        "'" & IIf(pdFecEval <> 0, "AND dPersEval = '" & Format(pdFecEval, "mm/dd/yyyy") & "'", "")
    Set R = Conn.CargaRecordSet(sSql)
    If R.RecordCount > 0 Then
        FuenteIngresoAsignadaACredito = True
    Else
        FuenteIngresoAsignadaACredito = False
    End If
    Set R = Nothing
End Function

Public Function UbicaPosicionFteIngreso(ByVal psPersCodFI As String, ByVal pdFecEval As Date, ByVal pnIndice As Integer) As Integer
Dim i As Integer
    If FtesIngreso(pnIndice).nTipo = gPersFteIngresoTipoDependiente Then
        For i = 0 To FtesIngreso(pnIndice).nNumFtesDep - 1
            If FtesIngreso(pnIndice).FtesIngresoDep(i).dFecEval = pdFecEval Then
                UbicaPosicionFteIngreso = i
                Exit Function
            End If
        Next i
    Else
        For i = 0 To FtesIngreso(pnIndice).nNumFtesIndep - 1
            If FtesIngreso(pnIndice).FtesIngresoInd(i).dFecEval = pdFecEval Then
                UbicaPosicionFteIngreso = i
                Exit Function
            End If
        Next i
    End If
End Function

Friend Function PersFuenteBuscaUltimaFechaEval(ByVal psNumFte As String) As Date
Dim i As Integer
Dim J As Integer
Dim dFecEvalTmp As Date
    dFecEvalTmp = CDate("01/01/1901")
    
    If psNumFte = "" Then
        If FtesIngreso(0).nTipo = gPersFteIngresoTipoDependiente Then
            dFecEvalTmp = FtesIngreso(0).FtesIngresoDep(0).dFecEval
        Else
            dFecEvalTmp = FtesIngreso(0).FtesIngresoInd(0).dFecEval
        End If
    End If
    
    For i = 0 To nNumFtes - 1
        If FtesIngreso(i).cNumFte = psNumFte Then
            If FtesIngreso(i).nTipo = gPersFteIngresoTipoDependiente Then
                For J = 0 To FtesIngreso(i).nNumFtesDep - 1
                    If FtesIngreso(i).FtesIngresoDep(J).dFecEval > dFecEvalTmp Then
                        dFecEvalTmp = FtesIngreso(i).FtesIngresoDep(J).dFecEval
                    End If
                Next J
            Else
                For J = 0 To FtesIngreso(i).nNumFtesIndep - 1
                    If FtesIngreso(i).FtesIngresoInd(J).dFecEval > dFecEvalTmp Then
                        dFecEvalTmp = FtesIngreso(i).FtesIngresoInd(J).dFecEval
                    End If
                Next J
            End If
        End If
    Next i
    PersFuenteBuscaUltimaFechaEval = dFecEvalTmp
End Function

Friend Function PersFuenteBuscaUltimaFechaCaducacion(ByVal psNumFte As String) As Date
Dim i As Integer
Dim J As Integer
Dim dFecEvalTmp As Date
    
    dFecEvalTmp = CDate("01/01/1901")
    If psNumFte = "" Then
        If FtesIngreso(0).nTipo = gPersFteIngresoTipoDependiente Then
            dFecEvalTmp = FtesIngreso(0).FtesIngresoDep(0).dFecCad
        Else
            dFecEvalTmp = FtesIngreso(0).FtesIngresoInd(0).dFecCad
        End If
    End If
    For i = 0 To nNumFtes - 1
        If FtesIngreso(i).cNumFte = psNumFte Then
            If FtesIngreso(i).nTipo = gPersFteIngresoTipoDependiente Then
                For J = 0 To FtesIngreso(i).nNumFtesDep - 1
                    If FtesIngreso(i).FtesIngresoDep(J).dFecCad > dFecEvalTmp Then
                        dFecEvalTmp = FtesIngreso(i).FtesIngresoDep(J).dFecCad
                    End If
                Next J
            Else
                For J = 0 To FtesIngreso(i).nNumFtesIndep - 1
                    If FtesIngreso(i).FtesIngresoInd(J).dFecCad > dFecEvalTmp Then
                        dFecEvalTmp = FtesIngreso(i).FtesIngresoInd(J).dFecCad
                    End If
                Next J
            End If
        End If
    Next i
    PersFuenteBuscaUltimaFechaCaducacion = dFecEvalTmp
End Function

Public Function FiltraFuentesIngresoPorRazonSocial() As Variant
Dim MatFte() As String
Dim MatFteTmp() As String
Dim i As Integer
Dim nNumFtetmp As Integer
Dim sPersCodFI As String

    
    ReDim MatFteTmp(nNumFtes, 9)
    nNumFtetmp = 0
    If nNumFtes > 0 Then
        For i = 0 To nNumFtes - 1
            nNumFtetmp = nNumFtetmp + 1
            MatFteTmp(nNumFtetmp - 1, 0) = nNumFtetmp
            MatFteTmp(nNumFtetmp - 1, 1) = IIf(FtesIngreso(i).nTipo = gPersFteIngresoTipoDependiente, "D", "I")
            MatFteTmp(nNumFtetmp - 1, 2) = FtesIngreso(i).sRazSocDescrip
            'MatFteTmp(nNumFtetmp - 1, 3) = PersFuenteBuscaUltimaFechaCaducacion(FtesIngreso(I).cNumFte) 'Fecha de Caducacion de la Fte de Ingreso
            'MatFteTmp(nNumFtetmp - 1, 4) = PersFuenteBuscaUltimaFechaEval(FtesIngreso(I).cNumFte) 'Fecha de Evaluacion de la Fte de Ingreso
            MatFteTmp(nNumFtetmp - 1, 5) = IIf(FtesIngreso(i).sMoneda = gMonedaNacional, "SOLES", "DOLARES") 'Moneda de la Fte de Ingreso
            MatFteTmp(nNumFtetmp - 1, 6) = FtesIngreso(i).sFteIngreso
            MatFteTmp(nNumFtetmp - 1, 7) = Trim(Str(i))
            MatFteTmp(nNumFtetmp - 1, 8) = FtesIngreso(i).cNumFte
            sPersCodFI = FtesIngreso(i).sFteIngreso
        Next i
        
        ReDim MatFte(nNumFtetmp, 8)
        For i = 0 To nNumFtetmp - 1
            MatFte(i, 0) = MatFteTmp(i, 0)
            MatFte(i, 1) = MatFteTmp(i, 1)
            MatFte(i, 2) = MatFteTmp(i, 2)
            MatFte(i, 3) = MatFteTmp(i, 3)
            MatFte(i, 4) = MatFteTmp(i, 4)
            MatFte(i, 5) = MatFteTmp(i, 5)
            MatFte(i, 6) = MatFteTmp(i, 6)
            MatFte(i, 7) = MatFteTmp(i, 7)
            MatFte(i, 8) = MatFteTmp(i, 8)
        Next i
        FiltraFuentesIngresoPorRazonSocial = MatFte
    Else
        FiltraFuentesIngresoPorRazonSocial = ""
    End If
End Function

Public Sub RecuperaFteIngresoDetalle(ByVal psNumFte As String, ByVal pnTipoFte As PersFteIngresoTipo)
Dim sSql As String

Dim R As ADODB.Recordset
    nNumFtesDep = 0
    nNumFtesIndep = 0
    
    If pnTipoFte = gPersFteIngresoTipoDependiente Then
        sSql = "Select PF.dPersEval, PF.nPersIngCli, PF.nPersGastoFam, PF.nPersIngCon, PF.nPersOtrIng "
        sSql = sSql & " From PersFIDependiente PF Where PF.cNumFte = '" & psNumFte & "' Order by dPersEval Asc"
        Set R = Conn.CargaRecordSet(sSql)
        ReDim FtesIngresoDep(100, R.RecordCount)
        Do While Not R.EOF
            nNumFtesDep = nNumFtesDep + 1
            FtesIngresoDep(nNumFtesDep - 1).cNumFte = psNumFte
            FtesIngresoDep(nNumFtesDep - 1).dFecEval = Format(R!dPersEval, "dd/mm/yyyy")
            FtesIngresoDep(nNumFtesDep - 1).nGastosFam = Format(R!nPersGastoFam, "#0.00")
            FtesIngresoDep(nNumFtesDep - 1).nIngCony = Format(R!nPersIngCon, "#0.00")
            FtesIngresoDep(nNumFtesDep - 1).nIngFam = Format(R!nPersIngCli, "#0.00")
            FtesIngresoDep(nNumFtesDep - 1).nIngOtros = Format(R!nPersOtrIng, "#0.00")
            FtesIngresoDep(nNumFtesDep - 1).nIngTotales = FtesIngresoDep(nNumFtesDep - 1).nIngresos + FtesIngresoDep(nNumFtesDep - 1).nIngFam + FtesIngresoDep(nNumFtesDep - 1).nIngCony + FtesIngresoDep(nNumFtesDep - 1).nIngOtros
            FtesIngresoDep(nNumFtesDep - 1).nSaldDispFam = FtesIngresoDep(nNumFtesDep - 1).nIngTotales - FtesIngresoDep(nNumFtesDep - 1).nGastosFam
            FtesIngresoDep(nNumFtesDep - 1).nTipoAct = 0
            R.MoveNext
        Loop
    Else
        sSql = "Select PF.dPersEval, PF.cNumFuente, PF.dPersEval, PF.nPersFIActivoDisp, "
        sSql = sSql & " PF.nPersFICtasxCobrar, PF.nPersFIInventarios, PF.nPersFIActivosFijos, "
        sSql = sSql & " PF.nPersFIProveedores, PF.nPersFICredCMACT, PF.nPersFICredOtros, "
        sSql = sSql & " PF.nPersFIVentas, PF.nPersFIPatrimonio, PF.nPersFICostoVentas, PF.nPersFIRecupCtasXCobrar,"
        sSql = sSql & " PF.nPersFIEgresosOtros, PF.nPersIngFam, PF.nPersEgrFam, PF.cUltimaActualizacion "
        sSql = sSql & " From PersFIIndependiente PF Where PF.cNumFte = '" & psNumFte & "' Order by dPersEval Asc"
        
        Set R = Conn.CargaRecordSet(sSql)
        ReDim FtesIngresoInd(100, R.RecordCount)
        Do While Not R.EOF
            nNumFtesIndep = nNumFtesIndep + 1
            FtesIngresoInd(nNumFtesIndep - 1).cNumFte = psNumFte
            FtesIngresoInd(nNumFtesIndep - 1).dFecEval = Format(R!dPersEval, "dd/mm/yyyy")
            FtesIngresoInd(nNumFtesIndep - 1).nActFijo = Format(R!nPersFIActivosFijos, "#0.00")
            FtesIngresoInd(nNumFtesIndep - 1).nActivDisp = Format(R!nPersFIActivoDisp, "#0.00")
            FtesIngresoInd(nNumFtesIndep - 1).nActivo = FtesIngresoInd(nNumFtesIndep - 1).nActivDisp + FtesIngresoInd(nNumFtesIndep - 1).nActFijo
            FtesIngresoInd(nNumFtesIndep - 1).nActivoCircul = 0
            FtesIngresoInd(nNumFtesIndep - 1).nCompraMercad = Format(R!nPersFICostoVentas, "#0.00")
            FtesIngresoInd(nNumFtesIndep - 1).nCreditosCmact = Format(R!nPersFICredCMACT, "#0.00")
            FtesIngresoInd(nNumFtesIndep - 1).nCreditosOtros = Format(R!nPersFICredOtros, "#0.00")
            FtesIngresoInd(nNumFtesIndep - 1).nCtasxCob = Format(R!nPersFICtasxCobrar, "#0.00")
            FtesIngresoInd(nNumFtesIndep - 1).nInventarios = Format(R!nPersFIInventarios, "#0.00")
            FtesIngresoInd(nNumFtesIndep - 1).nOtrosEgresos = Format(R!nPersFIEgresosOtros, "#0.00")
            FtesIngresoInd(nNumFtesIndep - 1).nPasivoPatrimonio = Format(R!nPersFIPatrimonio, "#0.00")
            FtesIngresoInd(nNumFtesIndep - 1).nPersEgrFam = Format(R!nPersEgrFam, "#0.00")
            FtesIngresoInd(nNumFtesIndep - 1).nPersIngFam = Format(R!nPersIngFam, "#0.00")
            FtesIngresoInd(nNumFtesIndep - 1).nProveedores = Format(R!nPersFIProveedores, "#0.00")
            FtesIngresoInd(nNumFtesIndep - 1).nRecupCtasxCob = Format(R!nPersFIRecupCtasXCobrar, "#0.00")
            FtesIngresoInd(nNumFtesIndep - 1).nVentas = Format(R!nPersFIVentas, "#0.00")
            FtesIngresoInd(nNumFtesIndep - 1).nPasivo = 0
            FtesIngresoInd(nNumFtesIndep - 1).nTipoAct = PersFilaSinCambios
            R.MoveNext
        Loop
    End If
            
End Sub

Public Sub RecuperaFtesdeIngreso(ByVal psCodPers As String, Optional ByVal pcCtaCod As String = "")
Dim sSql As String
Dim R As ADODB.Recordset
Dim bConexInterna As Boolean
    
    If pcCtaCod = "" Then
    
        sSql = "Select  top 7  PF.bCostoProd, PF.cNumFuente, PF.cPersCod, PF.cPersFIPersCod, PF.cPersFICargo, "
        sSql = sSql & " PF.dPersFIinicio, PF.cPersFIMoneda, PF.cPersFIComentario, "
        sSql = sSql & " PF.cRazSocUbiGeo, PF.cRazSocDirecc, PF.cRazSocTelef, PF.cRazSocDescrip, "
        sSql = sSql & " PF.nPersTipFte, PF.cUltimaActualizacion, "
    '------------------------------------------------------------------------------------------------------------
        sSql = sSql & "P.cPersDireccUbiGeo,P.cPersDireccDomicilio,P.cPersDireccCondicion, P.cPersTelefono, "
        sSql = sSql & "P.nPersPersoneria,P.cPersCIIU, P.cPersTelefono2, P.cPersCod as cPersFteCod,P.cPersNombre,P.dPersNacCreac, "
        sSql = sSql & "PJ.cPersJurTpo, PJ.cPersJurMagnitud,PJ.cPersJurSigla,PJ.nPersJurEmpleados, CI.cCIIUdescripcion, "
        sSql = sSql & "cNroRUC = (  Select ISNULL(cPersIDnro,'') From PersID WHERE PersID.cPersCod = P.cPersCod and PersID.cPersIDTpo = " & gPersIdRUC & " ), "
        sSql = sSql & "PJT.cPersJurTpoDesc, CN7.cConsDescripcion as cPersDireccCondicionDes, CN8.cConsDescripcion as cPersJurMagnitudDes, "
        sSql = sSql & " Dpto = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '1' And Substring(u.cUbigeoCod,2,2) = Substring(P.cPersDireccUbiGeo,2,2) And Substring(u.cUbigeoCod,4,9) = '000000000' ), '' ) , "
        sSql = sSql & " Prov = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '2' And Substring(u.cUbigeoCod,2,4) = Substring(P.cPersDireccUbiGeo,2,4) And Substring(u.cUbigeoCod,6,7) = '0000000' ), '' ) , "
        sSql = sSql & " Dist = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '3' And Substring(u.cUbigeoCod,2,6) = Substring(P.cPersDireccUbiGeo,2,6) And Substring(u.cUbigeoCod,8,5) = '00000' ), '' ), "
        sSql = sSql & " Zona = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where u.cUbigeoCod = P.cPersDireccUbiGeo ),'') "
        sSql = sSql & " From  PersFteIngreso PF "
        sSql = sSql & " INNER JOIN " & gConsPersona & "Persona P ON PF.cPersFIPersCod = P.cPersCod "
        sSql = sSql & " left join " & gConsPersona & "PersonaNat PN ON P.cPersCod=PN.cPersCod "
        sSql = sSql & " LEFT JOIN " & gConsPersona & "PersonaJur PJ ON P.cPersCod = PJ.cPersCod "
        sSql = sSql & " LEFT JOIN CIIU CI ON P.cPersCIIU = CI.cCIIUcod "
        sSql = sSql & " LEFT JOIN PersJurTpo PJT ON PJ.cPersJurTpo = PJT.cPersJurTpoCod "
        sSql = sSql & " Left Join Constante CN7 ON CN7.nConsValor = P.cPersDireccCondicion AND CN7.nConsCod = 1029 "
        sSql = sSql & " Left Join Constante CN8 ON CN8.nConsValor = PJ.cPersJurMagnitud AND CN8.nConsCod = 1004 "
        sSql = sSql & " Where PF.cPersCod = '" & psCodPers & "'"
        sSql = sSql & " and (exists(Select cnumfuente from persfidependiente where cnumfuente=PF.cNumfuente) or Exists(Select cnumfuente from persfiindependiente where cnumfuente=PF.cnumfuente))"
        sSql = sSql & " Order By dPersFIinicio desc"
    '----------------------------------------------------------------------------------------------------------
   Else
         sSql = " Select top 7 PF.bCostoProd, PF.cNumFuente, PF.cPersCod, PF.cPersFIPersCod, PF.cPersFICargo, "
        sSql = sSql & " PF.dPersFIinicio, PF.cPersFIMoneda, PF.cPersFIComentario, "
        sSql = sSql & " PF.cRazSocUbiGeo, PF.cRazSocDirecc, PF.cRazSocTelef, PF.cRazSocDescrip, "
        sSql = sSql & " PF.nPersTipFte, PF.cUltimaActualizacion, "
    '------------------------------------------------------------------------------------------------------------
        sSql = sSql & "P.cPersDireccUbiGeo,P.cPersDireccDomicilio,P.cPersDireccCondicion, P.cPersTelefono, "
        sSql = sSql & "P.nPersPersoneria,P.cPersCIIU, P.cPersTelefono2, P.cPersCod as cPersFteCod,P.cPersNombre,P.dPersNacCreac, "
        sSql = sSql & "PJ.cPersJurTpo, PJ.cPersJurMagnitud,PJ.cPersJurSigla,PJ.nPersJurEmpleados, CI.cCIIUdescripcion, "
        sSql = sSql & "cNroRUC = (  Select ISNULL(cPersIDnro,'') From PersID WHERE PersID.cPersCod = P.cPersCod and PersID.cPersIDTpo = " & gPersIdRUC & " ), "
        sSql = sSql & "PJT.cPersJurTpoDesc, CN7.cConsDescripcion as cPersDireccCondicionDes, CN8.cConsDescripcion as cPersJurMagnitudDes, "
        sSql = sSql & " Dpto = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '1' And Substring(u.cUbigeoCod,2,2) = Substring(P.cPersDireccUbiGeo,2,2) And Substring(u.cUbigeoCod,4,9) = '000000000' ), '' ) , "
        sSql = sSql & " Prov = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '2' And Substring(u.cUbigeoCod,2,4) = Substring(P.cPersDireccUbiGeo,2,4) And Substring(u.cUbigeoCod,6,7) = '0000000' ), '' ) , "
        sSql = sSql & " Dist = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '3' And Substring(u.cUbigeoCod,2,6) = Substring(P.cPersDireccUbiGeo,2,6) And Substring(u.cUbigeoCod,8,5) = '00000' ), '' ), "
        sSql = sSql & " Zona = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where u.cUbigeoCod = P.cPersDireccUbiGeo ),'') "
        sSql = sSql & " From  PersFteIngreso PF "
        sSql = sSql & " INNER JOIN " & gConsPersona & "Persona P ON PF.cPersFIPersCod = P.cPersCod "
        sSql = sSql & " left join " & gConsPersona & "PersonaNat PN ON P.cPersCod=PN.cPersCod "
        sSql = sSql & " LEFT JOIN " & gConsPersona & "PersonaJur PJ ON P.cPersCod = PJ.cPersCod "
        sSql = sSql & " LEFT JOIN CIIU CI ON P.cPersCIIU = CI.cCIIUcod "
        sSql = sSql & " LEFT JOIN PersJurTpo PJT ON PJ.cPersJurTpo = PJT.cPersJurTpoCod "
        sSql = sSql & " Left Join Constante CN7 ON CN7.nConsValor = P.cPersDireccCondicion AND CN7.nConsCod = 1029 "
        sSql = sSql & " Left Join Constante CN8 ON CN8.nConsValor = PJ.cPersJurMagnitud AND CN8.nConsCod = 1004 "
        sSql = sSql & " Inner Join ColocFteIngreso CFI on CFI.cNumFuente=PF.cNumFuente"
        sSql = sSql & " Where PF.cPersCod = '" & psCodPers & "'"
        sSql = sSql & " and (exists(Select cnumfuente from persfidependiente where cnumfuente=PF.cNumfuente) or Exists(Select cnumfuente from persfiindependiente where cnumfuente=PF.cnumfuente))"
        sSql = sSql & " and CFI.cCtaCod='" & pcCtaCod & "'"
        sSql = sSql & " Order By dPersFIinicio desc "
    '----------------------------------------------------------------------------------------------------------
   End If
    
    nNumFtes = 0
    bConexInterna = False
    ReDim FtesIngreso(nNumFtes)
    If Conn Is Nothing Then
        bConexInterna = True
        Set Conn = New COMConecta.DCOMConecta
        Conn.AbreConexion
    End If
    Set R = Conn.CargaRecordSet(sSql)
    Do While Not R.EOF
        nNumFtes = nNumFtes + 1
        ReDim Preserve FtesIngreso(nNumFtes)
        FtesIngreso(nNumFtes - 1).cNumFte = R!cNumFuente
        FtesIngreso(nNumFtes - 1).bCostoProd = IIf(IIf(IsNull(R!bCostoProd), 0, R!bCostoProd), 1, 0)
        FtesIngreso(nNumFtes - 1).dFteIngreso = Format(R!dPersFIinicio, "dd/mm/yyyy")
        FtesIngreso(nNumFtes - 1).sCargo = R!cPersFICargo
        FtesIngreso(nNumFtes - 1).sComentarios = R!cPersFIComentario
        FtesIngreso(nNumFtes - 1).sFteIngreso = R!cPersFIPersCod
        FtesIngreso(nNumFtes - 1).sFteIngresoOrig = R!cPersFIPersCod
        FtesIngreso(nNumFtes - 1).sMoneda = Trim(Str(R!cPersFIMoneda))
        FtesIngreso(nNumFtes - 1).nTipo = IIf(IsNull(R!nPersTipFte), 0, R!nPersTipFte)
        FtesIngreso(nNumFtes - 1).sRazSocUbiGeo = Trim(IIf(IsNull(R!cRazSocUbiGeo), "", R!cRazSocUbiGeo))
        FtesIngreso(nNumFtes - 1).sRazSocDirecc = Trim(IIf(IsNull(R!cRazSocDirecc), "", R!cRazSocDirecc))
        FtesIngreso(nNumFtes - 1).sRazSocTelef = Trim(IIf(IsNull(R!cRazSocTelef), "", R!cRazSocTelef))
        FtesIngreso(nNumFtes - 1).sRazSocDescrip = Trim(IIf(IsNull(R!cRazSocDescrip), "", R!cRazSocDescrip))
        FtesIngreso(nNumFtes - 1).nTipoAct = 0
        FtesIngreso(nNumFtes - 1).cFtePersRuc = Trim(IIf(IsNull(R!cNroRUC), "", R!cNroRUC))
        FtesIngreso(nNumFtes - 1).dFtePersFecInicioAct = Format(IIf(IsNull(R!dPersNacCreac), "01/01/1900", R!dPersNacCreac), "dd/mm/yyyy")
        FtesIngreso(nNumFtes - 1).cFtePersTipoPersJur = Trim(IIf(IsNull(R!cPersJurTpoDesc), "", R!cPersJurTpoDesc))
        FtesIngreso(nNumFtes - 1).cFtePersTelefono = Trim(IIf(IsNull(R!cPersTelefono), "", R!cPersTelefono))
        FtesIngreso(nNumFtes - 1).cFtePersCIUU = Trim(IIf(IsNull(R!cCIIUdescripcion), "", R!cCIIUdescripcion))
        FtesIngreso(nNumFtes - 1).cFtePersCondicionDomic = Trim(IIf(IsNull(R!cPersDireccCondicionDes), "", R!cPersDireccCondicionDes))
        FtesIngreso(nNumFtes - 1).cFtePersMagnitudEmp = Trim(IIf(IsNull(R!cPersJurMagnitudDes), "", R!cPersJurMagnitudDes))
        FtesIngreso(nNumFtes - 1).nFtePersNroEmpleados = IIf(IsNull(R!nPersJurEmpleados), 0, R!nPersJurEmpleados)
        FtesIngreso(nNumFtes - 1).cFtePersDireccion = Trim(IIf(IsNull(R!cPersDireccDomicilio), "", R!cPersDireccDomicilio))
        FtesIngreso(nNumFtes - 1).cFtePersDpto = Trim(IIf(IsNull(R!Dpto), "", R!Dpto))
        FtesIngreso(nNumFtes - 1).cFtePersProv = Trim(IIf(IsNull(R!Prov), "", R!Prov))
        FtesIngreso(nNumFtes - 1).cFtePersDist = Trim(IIf(IsNull(R!Dist), "", R!Dist))
        FtesIngreso(nNumFtes - 1).cFtePersZona = Trim(IIf(IsNull(R!Zona), "", R!Zona))
        
        R.MoveNext
    Loop
    R.Close
    
    If bConexInterna Then
        Conn.CierraConexion
        Set Conn = Nothing
    End If
End Sub

Public Sub RecuperaFtesIngresoDependiente(ByVal pnNumFte As Integer)
Dim sSql As String
Dim R As ADODB.Recordset

    FtesIngreso(pnNumFte).nNumFtesDep = 0
    sSql = " Select top 7 PF.cNumFuente, PF.dPersEval, PF.nPersIngCli, PF.nPersGastoFam, PF.nPersIngCon, PF.nPersOtrIng, PF.dPersFICaduca "
    sSql = sSql & " FROM PersFIDependiente PF Where PF.cNumFuente = '" & FtesIngreso(pnNumFte).cNumFte & "'"
    sSql = sSql & " ORDER BY PF.dPersEval desc"
    Set R = Conn.CargaRecordSet(sSql)
    FtesIngreso(pnNumFte).nNumFtesDep = 0
    Do While Not R.EOF
        FtesIngreso(pnNumFte).nNumFtesDep = FtesIngreso(pnNumFte).nNumFtesDep + 1
        ReDim Preserve FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep)
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).cNumFte = R!cNumFuente
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).dFecCad = R!dPersFICaduca
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).dFecEval = Format(R!dPersEval, "dd/mm/yyyy")
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).nGastosFam = Format(R!nPersGastoFam, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).nIngCony = Format(R!nPersIngCon, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).nIngFam = Format(R!nPersIngCli, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).nIngOtros = Format(R!nPersOtrIng, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).nIngresos = Format(R!nPersIngCli, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).nIngTotales = Format(R!nPersIngCon + R!nPersOtrIng + R!nPersIngCli + R!nPersOtrIng, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).nSaldDispFam = Format(FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).nIngTotales - FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).nGastosFam, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).nTipoAct = PersFilaSinCambios
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
End Sub

Public Sub RecuperaFtesIngresoCostosProd(ByVal pnNumFte As Integer)
Dim sSql As String
Dim R As ADODB.Recordset

    
    sSql = " Select C.dFecEval, CN2.cConsDescripcion as sUniProd, CN1.cConsDescripcion as sTipoCultivo,  C.cNumFuente, C.nTpoCultivo, C.nMaquinaria, C.nJornales, C.nInsumos, C.nPesticidas, C.nOtros, C.nHectareas,C.nMantenimiento, "
    sSql = sSql & " C.nProduccion , C.nUniProd, C.nPreUni,C.nSiembra,C.nCosecha,C.nDesaAgricola,C.nPlanOtros "
    sSql = sSql & "  From PersFICostoProd C Inner Join Constante CN1 ON CN1.nConsCod = 1046 AND C.nTpoCultivo = CN1.nConsValor "
    sSql = sSql & "                 Inner Join Constante CN2 ON CN2.nConsCod = 1045 AND C.nUniProd = CN2.nConsValor "
    sSql = sSql & "  Where cNumFuente = '" & FtesIngreso(pnNumFte).cNumFte & "' "
    Set R = Conn.CargaRecordSet(sSql)
    FtesIngreso(pnNumFte).nNumCostosProd = 0
    Do While Not R.EOF
        FtesIngreso(pnNumFte).nNumCostosProd = FtesIngreso(pnNumFte).nNumCostosProd + 1
        ReDim Preserve FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd)
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).dFecEval = Format(R!dFecEval, "dd/mm/yyyy")
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).sTpoCultivo = R!sTipoCultivo
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).sUniProd = R!sUniProd
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).cNumFuente = R!cNumFuente
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nTpoCultivo = R!nTpoCultivo
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nMaquinaria = R!nMaquinaria
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nJornales = R!nJornales
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nInsumos = R!nInsumos
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nPesticidas = R!nPesticidas
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nOtros = R!nOtros
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nHectareas = R!nHectareas
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nProduccion = R!nProduccion
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nUniProd = R!nUniProd
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nPreUni = R!nPreUni
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nSiembra = R!nSiembra
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nMantenimiento = R!nMantenimiento
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nDesAgricola = R!nDesaAgricola
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nCosecha = R!nCosecha
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nPlanOtros = R!nPlanOtros
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
End Sub



Public Sub RecuperaFtesIngresoIndependiente(ByVal pnNumFte As Integer)
Dim sSql As String
Dim R As ADODB.Recordset

    FtesIngreso(pnNumFte).nNumFtesIndep = 0
    sSql = "SELECT  top 7 PF.cNumFuente, PF.dPersEval, PF.dPersFICaduca, PF.nPersFIActivoDisp, "
    sSql = sSql & " PF.nPersFICtasxCobrar, PF.nPersFIInventarios, PF.nPersFIActivosFijos, "
    sSql = sSql & " PF.nPersFIProveedores, PF.nPersFICredCMACT, PF.nPersFICredOtros, "
    sSql = sSql & " PF.nPersFIVentas, PF.nPersFIPatrimonio, PF.nPersFICostoVentas, "
    sSql = sSql & " PF.nPersFIRecupCtasXCobrar, PF.nPersFIEgresosOtros, PF.nPersIngFam, "
    sSql = sSql & " PF.nPersEgrFam, PF.cUltimaActualizacion "
    sSql = sSql & " FROM PersFIIndependiente PF Where PF.cNumFuente = '" & FtesIngreso(pnNumFte).cNumFte & "'"
    sSql = sSql & " ORDER BY PF.dPersEval desc "
        
    Set R = Conn.CargaRecordSet(sSql)
    FtesIngreso(pnNumFte).nNumFtesIndep = 0
    ReDim FtesIngreso(pnNumFte).FtesIngresoInd(R.RecordCount)
    Do While Not R.EOF
        FtesIngreso(pnNumFte).nNumFtesIndep = FtesIngreso(pnNumFte).nNumFtesIndep + 1
        ReDim Preserve FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep)
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).cNumFte = R!cNumFuente
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).dFecCad = Format(R!dPersFICaduca, "dd/mm/yyyy")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).dFecEval = Format(R!dPersEval, "dd/mm/yyyy")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nActFijo = Format(R!nPersFIActivosFijos, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nActivDisp = Format(R!nPersFIActivoDisp, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nActivo = Format(R!nPersFIActivoDisp + R!nPersFIActivosFijos, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nActivoCircul = "0.00"
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nCompraMercad = Format(R!nPersFICostoVentas, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nCreditosCmact = Format(R!nPersFICredCMACT, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nCreditosOtros = Format(R!nPersFICredOtros, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nCtasxCob = Format(R!nPersFICtasxCobrar, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nInventarios = Format(R!nPersFIInventarios, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nOtrosEgresos = Format(R!nPersFIEgresosOtros, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nPasivoPatrimonio = Format(R!nPersFIPatrimonio, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nPersEgrFam = Format(R!nPersEgrFam, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nPersIngFam = Format(R!nPersIngFam, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nProveedores = Format(R!nPersFIProveedores, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nRecupCtasxCob = Format(R!nPersFIRecupCtasXCobrar, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nTipoAct = PersFilaSinCambios
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nVentas = Format(R!nPersFIVentas, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nPasivo = Format(0#, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nTipoAct = PersFilaSinCambios
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
End Sub

Public Sub RecuperaPersona(ByVal psCodPers As String, Optional ByVal psCtaCod As String = "")
Dim sSql As String
Dim R As ADODB.Recordset
    
    Set Conn = New COMConecta.DCOMConecta
    Conn.AbreConexion
    Set ConexionActiva = Conn.ConexionActiva
    
    '***** CARGA DATOS DE LA PERSONA ******
    sSql = "Select P.nPersRelaInst, P.cPersCodSbs, P.cPersEmail cEmail,P.cPersTelefono2,P.cPersCod,P.cPersNombre,P.dPersNacCreac, "
    sSql = sSql & "P.cPersDireccUbiGeo,P.cPersDireccDomicilio,P.cPersDireccCondicion, P.nPersValComDomicilio, P.cPersTelefono, "
    sSql = sSql & "P.nPersPersoneria,P.cPersCIIU,P.cPersEstado,P.cUltimaActualizacion, "
    sSql = sSql & "PN.cPersNatSexo,PN.nPersNatEstCiv,PN.nPersNatHijos,PN.nPersPeso,PN.nPersTalla,PN.nPersTpoSangre, PN.cNacionalidad, PN.nResidente, "
    sSql = sSql & "PJ.cPersJurTpo, PJ.cPersJurMagnitud,PJ.cPersJurSigla,PJ.nPersJurEmpleados, CI.cCIIUdescripcion"
    sSql = sSql & " FROM " & gConsPersona & "Persona P left join " & gConsPersona & "PersonaNat PN ON P.cPersCod=PN.cPersCod "
    sSql = sSql & " LEFT JOIN " & gConsPersona & "PersonaJur PJ ON P.cPersCod = PJ.cPersCod "
    sSql = sSql & " LEFT JOIN CIIU CI ON P.cPersCIIU = CI.cCIIUcod "
    sSql = sSql & " WHERE P.cPersCod = '" & psCodPers & "'"
    
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
        dFechaNac = CDate(Format(IIf(IsNull(R!dPersNacCreac), "01/01/1900", R!dPersNacCreac), "dd/mm/yyyy"))
        sUbicGeografica = R!cPersDireccUbiGeo
        sDomicilio = R!cPersDireccDomicilio
        sCondicionDomic = IIf(IsNull(R!cPersDireccCondicion), Str(gPersCondDomicEnUso), R!cPersDireccCondicion)
        nValComDomicilio = IIf(IsNull(R!nPersValComDomicilio), 0, R!nPersValComDomicilio)
        sTelefonos = IIf(IsNull(R!cPersTelefono), "", R!cPersTelefono)
        sTelefonos2 = IIf(IsNull(R!cPersTelefono2), "", R!cPersTelefono2)
        sEmail = IIf(IsNull(R!cEmail), "", R!cEmail)
        sCiiu = IIf(IsNull(R!cPersCiiu), "", R!cPersCiiu)
        sCiiuDescripcion = IIf(IsNull(R!cCIIUdescripcion), "", R!cCIIUdescripcion)
        sEstado = R!cPersEstado
        nTalla = IIf(IsNull(R!nPersTalla), 0#, R!nPersTalla)
        nPeso = IIf(IsNull(R!nPersPeso), 0#, R!nPersPeso)
        sTipoSangre = Trim(Str(IIf(IsNull(R!nPersTpoSangre), 0, R!nPersTpoSangre)))
        cNacionalidad = Trim(IIf(IsNull(R!cNacionalidad), "", R!cNacionalidad))
        nResidencia = IIf(IsNull(R!nResidente), 0, R!nResidente)
        nTipoPersona = R!nPersPersoneria
        'Datos de Persona Juridica
        'sRazonSocial
        sSiglas = IIf(IsNull(R!cPersJurSigla), "", R!cPersJurSigla)
        sTipoPersJur = IIf(IsNull(R!cPersJurMagnitud), "", R!cPersJurTpo)
        nNumEmplead = IIf(IsNull(R!nPersJurEmpleados), 0, R!nPersJurEmpleados)
        sMagnitudEmp = IIf(IsNull(R!cPersJurMagnitud), "", R!cPersJurMagnitud)
        sPersCodSbs = IIf(IsNull(R!cPersCodSbs), "", R!cPersCodSbs)
        nPersRelInst = IIf(IsNull(R!npersrelainst), 4, R!npersrelainst)
        nTipoAct = 0
    Else
        sPersCod = ""
        R.Close
        Set R = Nothing
        'Conn.CierraConexion
        'Set Conn = Nothing
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
    Call RecuperaFtesdeIngreso(psCodPers, psCtaCod)
    Call RecuperaRefComercial(psCodPers)   'Carga las Referencias Comerciales de la Persona
    Call RecuperaMaxRefComercial(psCodPers)
    Call RecuperaRefBancaria(psCodPers)     'Carga las Referencias Bancarias de la Persona
    Call RecuperaPatVehicular(psCodPers)    'Carga el Patrimonio Vehicular
    Call RecuperaMaxPatVehicular(psCodPers)
    Exit Sub
        
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

Public Property Get PersCodSbs() As String
    PersCodSbs = sPersCodSbs
End Property

Public Property Let PersCodSbs(ByVal vNewValue As String)
    sPersCodSbs = vNewValue
End Property

Public Property Get PersRelInst() As Integer
    PersRelInst = nPersRelInst
End Property

Public Property Let PersRelInst(ByVal vNewValue As Integer)
    nPersRelInst = vNewValue
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

Public Property Get ValComDomicilio() As String
    ValComDomicilio = nValComDomicilio
End Property

Public Property Let ValComDomicilio(ByVal vNewValue As String)
    nValComDomicilio = vNewValue
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

Public Property Let Residencia(ByVal vNewValue As Integer)
    nResidencia = vNewValue
End Property

Public Property Get Residencia() As Integer
    Residencia = nResidencia
End Property

Public Property Let Nacionalidad(ByVal vNewValue As String)
    cNacionalidad = vNewValue
End Property

Public Property Get Nacionalidad() As String
    Nacionalidad = cNacionalidad
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

Public Property Get NumeroRefComercial() As Integer
    NumeroRefComercial = nNumRefCom
End Property

Public Property Let NumeroRefComercial(ByVal vNewValue As Integer)
    nNumRefCom = vNewValue
End Property

Public Property Get MaxRefComercial() As Integer
    MaxRefComercial = nMaxRefCom
End Property

Public Property Get NumeroRefBancaria() As Integer
    NumeroRefBancaria = nNumRefBan
End Property

Public Property Let NumeroRefBancaria(ByVal vNewValue As Integer)
    nNumRefBan = vNewValue
End Property

Public Property Get NumeroPatVehicular() As Integer
    NumeroPatVehicular = nNumPatVeh
End Property

Public Property Let NumeroPatVehicular(ByVal vNewValue As Integer)
    nNumPatVeh = vNewValue
End Property

Public Property Get MaxPatVehicular() As Integer
    MaxPatVehicular = nMaxPatVeh
End Property

Public Function ObtenerRelaPersCodigo(ByVal pnPunt As Integer) As String
    ObtenerRelaPersCodigo = Relac(pnPunt).sPersCod
End Function

Public Function ObtenerRelaPersTipoAct(ByVal pnPunt As Integer) As TPersCambios
    ObtenerRelaPersTipoAct = Relac(pnPunt).nTipoAct
End Function

Public Function ObtenerRefComTipoAct(ByVal pnPunt As Integer) As TPersCambios
    ObtenerRefComTipoAct = RefCom(pnPunt).nTipoAct
End Function

Public Function ObtenerRefBanTipoAct(ByVal pnPunt As Integer) As TPersCambios
    ObtenerRefBanTipoAct = RefBan(pnPunt).nTipoAct
End Function

Public Function ObtenerPatVehTipoAct(ByVal pnPunt As Integer) As TPersCambios
    ObtenerPatVehTipoAct = PatVeh(pnPunt).nTipoAct
End Function

Public Function ObtenerRelaPersNombres(ByVal pnPunt As Integer) As String
    ObtenerRelaPersNombres = Relac(pnPunt).sApellNombres
End Function

Public Function ObtenerRefComNombre(ByVal pnPunt As Integer) As String
    ObtenerRefComNombre = RefCom(pnPunt).sRefComNom
End Function

Public Function ObtenerRefComRelacion(ByVal pnPunt As Integer) As String
    ObtenerRefComRelacion = RefCom(pnPunt).sRefComTipo
End Function

Public Function ObtenerRefComFono(ByVal pnPunt As Integer) As String
    ObtenerRefComFono = RefCom(pnPunt).sRefComFono
End Function

Public Function ObtenerRefComNumRef(ByVal pnPunt As Integer) As Integer
    ObtenerRefComNumRef = RefCom(pnPunt).nRefComCod
End Function

Public Function ObtenerRefBanCodIF(ByVal pnPunt As Integer) As String
    ObtenerRefBanCodIF = RefBan(pnPunt).sRefBanCodIF
End Function

Public Function ObtenerRefBanNombre(ByVal pnPunt As Integer) As String
    ObtenerRefBanNombre = RefBan(pnPunt).sRefBanNombre
End Function

Public Function ObtenerRefBanNumCta(ByVal pnPunt As Integer) As String
    ObtenerRefBanNumCta = RefBan(pnPunt).sRefBanNumCta
End Function

Public Function ObtenerRefBanNumTar(ByVal pnPunt As Integer) As String
    ObtenerRefBanNumTar = RefBan(pnPunt).sRefBanNumTar
End Function

Public Function ObtenerRefBanLinCred(ByVal pnPunt As Integer) As Currency
    ObtenerRefBanLinCred = RefBan(pnPunt).nRefBanLinCred
End Function

Public Function ObtenerPatVehCod(ByVal pnPunt As Integer) As Integer
    ObtenerPatVehCod = PatVeh(pnPunt).nPatVehCod
End Function

Public Function ObtenerPatVehMarca(ByVal pnPunt As Integer) As String
    ObtenerPatVehMarca = PatVeh(pnPunt).sPatVehMarca
End Function

Public Function ObtenerPatVehFecFab(ByVal pnPunt As Integer) As String
    ObtenerPatVehFecFab = PatVeh(pnPunt).sPatVehFecFab
End Function

Public Function ObtenerPatVehValCom(ByVal pnPunt As Integer) As Currency
    ObtenerPatVehValCom = PatVeh(pnPunt).nPatVehValCom
End Function

Public Function ObtenerPatVehCondicion(ByVal pnPunt As Integer) As String
     ObtenerPatVehCondicion = PatVeh(pnPunt).sPatVehCondicion
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

Public Sub ActualizarRefComTipoAct(ByVal pCambio As TPersCambios, ByVal pnPunt As Integer)
     RefCom(pnPunt).nTipoAct = pCambio
End Sub

Public Sub ActualizarRefBanTipoAct(ByVal pCambio As TPersCambios, ByVal pnPunt As Integer)
     RefBan(pnPunt).nTipoAct = pCambio
End Sub

Public Sub ActualizarPatVehTipoAct(ByVal pCambio As TPersCambios, ByVal pnPunt As Integer)
    PatVeh(pnPunt).nTipoAct = pCambio
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
    ObtenerFteIngTipo = FtesIngreso(pnPunt).nTipo
End Function

Public Function ObtenerFteIngcNumFuente(ByVal pnPunt As Integer) As String
    ObtenerFteIngcNumFuente = FtesIngreso(pnPunt).cNumFte
End Function

Public Function ObtenerFteIngFuente(ByVal pnPunt As Integer) As String
    ObtenerFteIngFuente = FtesIngreso(pnPunt).sFteIngreso
End Function

Public Function ObtenerFteIngFecEval(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer, Optional ByVal pnTpoFte As PersFteIngresoTipo) As Date
    If FtesIngreso(pnPunt).nTipo = gPersFteIngresoTipoDependiente Then
        ObtenerFteIngFecEval = FtesIngreso(pnPunt).FtesIngresoDep(pnPuntDet).dFecEval
    Else
        ObtenerFteIngFecEval = FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).dFecEval
    End If
End Function

Public Function ObtenerFteIngFecCaducac(ByVal pnPunt As Integer, ByVal nPuntDet As Integer) As Date
    If FtesIngreso(pnPunt).nTipo = gPersFteIngresoTipoDependiente Then
        ObtenerFteIngFecCaducac = FtesIngreso(pnPunt).FtesIngresoDep(nPuntDet).dFecCad
    Else
        ObtenerFteIngFecCaducac = FtesIngreso(pnPunt).FtesIngresoInd(nPuntDet).dFecCad
    End If
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

Public Function ObtenerFteIngIngresos(ByVal pnPunt As Integer, ByVal nPuntDet As Integer) As Double
    ObtenerFteIngIngresos = FtesIngreso(pnPunt).FtesIngresoDep(nPuntDet).nIngresos
End Function

Public Function ObtenerFteIngIngresoFam(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerFteIngIngresoFam = FtesIngreso(pnPunt).FtesIngresoDep(pnPuntDet).nIngFam
End Function

Public Function ObtenerFteIngIngresoConyugue(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerFteIngIngresoConyugue = FtesIngreso(pnPunt).FtesIngresoDep(pnPuntDet).nIngCony
End Function

Public Function ObtenerFteIngIngresoOtros(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerFteIngIngresoOtros = FtesIngreso(pnPunt).FtesIngresoDep(pnPuntDet).nIngOtros
End Function

Public Function ObtenerFteIngIngresoTotales(ByVal pnPunt As Integer) As Double
    ObtenerFteIngIngresoTotales = FtesIngreso(pnPunt).nIngTotales
End Function

Public Function ObtenerFteIngIngresoTipo(ByVal pnPunt As Integer) As Integer
    ObtenerFteIngIngresoTipo = FtesIngreso(pnPunt).nTipo
End Function

Public Function ObtenerFteIngIngresoNumeroCostoProd(ByVal pnPunt As Integer) As Integer
    ObtenerFteIngIngresoNumeroCostoProd = FtesIngreso(pnPunt).nNumCostosProd
End Function

Public Function ObtenerFteIngIngresoNumeroFteDep(ByVal pnPunt As Integer) As Integer
    ObtenerFteIngIngresoNumeroFteDep = FtesIngreso(pnPunt).nNumFtesDep
End Function

Public Function ObtenerFteIngIngresoNumeroFteIndep(ByVal pnPunt As Integer) As Integer
    ObtenerFteIngIngresoNumeroFteIndep = FtesIngreso(pnPunt).nNumFtesIndep
End Function

Public Function ObtenerFteIngGastoFam(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerFteIngGastoFam = FtesIngreso(pnPunt).FtesIngresoDep(pnPuntDet).nGastosFam
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

Public Function ObtenerFteIngActivoDisp(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerFteIngActivoDisp = FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nActivDisp
End Function

Public Function ObtenerFteIngBalIngFam(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerFteIngBalIngFam = FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nPersIngFam
End Function


Public Function ObtenerFteIngBalEgrFam(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerFteIngBalEgrFam = FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nPersEgrFam
End Function

Public Function ObtenerFteIngCtasxCob(ByVal pnPunt As Integer, ByVal nPuntDet As Integer) As Double
    ObtenerFteIngCtasxCob = FtesIngreso(pnPunt).FtesIngresoInd(nPuntDet).nCtasxCob
End Function

Public Function ObtenerFteIngInventario(ByVal pnPunt As Integer, ByVal nPuntDet As Integer) As Double
    ObtenerFteIngInventario = FtesIngreso(pnPunt).FtesIngresoInd(nPuntDet).nInventarios
End Function

Public Function ObtenerFteIngActivoFijo(ByVal pnPunt As Integer, ByVal nPuntDet As Integer) As Double
    ObtenerFteIngActivoFijo = FtesIngreso(pnPunt).FtesIngresoInd(nPuntDet).nActFijo
End Function

Public Function ObtenerFteIngPasivoPatrim(ByVal pnPunt As Integer) As Double
    ObtenerFteIngPasivoPatrim = FtesIngreso(pnPunt).nPasivoPatrimonio
End Function

Public Function ObtenerFteIngPasivo(ByVal pnPunt As Integer) As Double
    ObtenerFteIngPasivo = FtesIngreso(pnPunt).nPasivo
End Function

Public Function ObtenerFteIngProveedores(ByVal pnPunt As Integer, ByVal nPuntDet As Integer) As Double
    ObtenerFteIngProveedores = FtesIngreso(pnPunt).FtesIngresoInd(nPuntDet).nProveedores
End Function

Public Function ObtenerFteIngOtrosCreditos(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerFteIngOtrosCreditos = FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nCreditosOtros
End Function

Public Function ObtenerFteIngCreditosCmact(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerFteIngCreditosCmact = FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nCreditosCmact
End Function

'Public Function ObtenerFteIngIngresosTotales(ByVal pnPunt As Integer) As Double
'    If CInt(FtesIngreso(pnPunt).nTipo) = gPersFteIngresoTipoDependiente Then
'        ObtenerFteIngIngresosTotales = FtesIngreso(pnPunt).nIngresos + FtesIngreso(pnPunt).nIngFam + FtesIngreso(pnPunt).nIngOtros - FtesIngreso(pnPunt).nGastosFam
'    Else
'        ObtenerFteIngIngresosTotales = FtesIngreso(pnPunt).nVentas + FtesIngreso(pnPunt).nRecupCtasxCob + FtesIngreso(pnPunt).nIngFam - (FtesIngreso(pnPunt).nCompraMercad + FtesIngreso(pnPunt).nOtrosEgresos + FtesIngreso(pnPunt).nGastosFam)
'    End If
'End Function

Public Function ObtenerFteIngVentas(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerFteIngVentas = FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nVentas
End Function

Public Function ObtenerFteIngRecupCtasxCobrar(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerFteIngRecupCtasxCobrar = FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nRecupCtasxCob
End Function

Public Function ObtenerFteIngComprasMercad(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerFteIngComprasMercad = FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nCompraMercad
End Function

Public Function ObtenerFteIngOtrosEgresos(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerFteIngOtrosEgresos = FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nOtrosEgresos
End Function


Public Function ObtenerFteIngComentarios(ByVal pnPunt As Integer) As String
    ObtenerFteIngComentarios = FtesIngreso(pnPunt).sComentarios
End Function

Public Function ObtenerFteIngRazonSocial(ByVal pnPunt As Integer) As String
    ObtenerFteIngRazonSocial = FtesIngreso(pnPunt).sRazSocDescrip
End Function

Public Function ObtenerFteIngbCostoProd(ByVal pnPunt As Integer) As Integer
    ObtenerFteIngbCostoProd = FtesIngreso(pnPunt).bCostoProd
End Function

Public Sub ActualizarFteIngTipoFuente(ByVal psTipoFte As Integer, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nTipo = psTipoFte
End Sub

Public Sub ActualizarFteIngComentarios(ByVal psComentarios As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).sComentarios = psComentarios
End Sub

Public Sub ActualizarFteIngOtrosEgresos(ByVal pnOtrosEgresos As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nOtrosEgresos = pnOtrosEgresos
End Sub

Public Sub ActualizarFteIngCompraMercad(ByVal pnCompraMercad As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nCompraMercad = pnCompraMercad
End Sub

Public Sub ActualizarFteIngRazSocDirecc(ByVal psRazSocDirecc As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).sRazSocDirecc = psRazSocDirecc
End Sub

Public Sub ActualizarFteIngRazSocUbiGeo(ByVal psRazSocUbiGeo As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).sRazSocUbiGeo = psRazSocUbiGeo
End Sub

Public Sub ActualizarFteIngRazSocTelef(ByVal psRazSocTelef As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).sRazSocTelef = psRazSocTelef
End Sub

Public Sub ActualizarFteIngRazSocDescrip(ByVal psRazSocDescrip As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).sRazSocDescrip = psRazSocDescrip
End Sub

Public Function ObtenerFteIngRazSocDescrip(ByVal pnPunt As Integer) As String
    ObtenerFteIngRazSocDescrip = FtesIngreso(pnPunt).sRazSocDescrip
End Function

Public Function ObtenerFteIngRazSocDirecc(ByVal pnPunt As Integer) As String
    ObtenerFteIngRazSocDirecc = FtesIngreso(pnPunt).sRazSocDirecc
End Function

Public Function ObtenerFteIngRazSocUbiGeo(ByVal pnPunt As Integer) As String
    ObtenerFteIngRazSocUbiGeo = FtesIngreso(pnPunt).sRazSocUbiGeo
End Function

Public Function ObtenerFteIngRazSocTelef(ByVal pnPunt As Integer) As String
    ObtenerFteIngRazSocTelef = FtesIngreso(pnPunt).sRazSocTelef
End Function

Public Sub ActualizarFteIngRecupCtasxCob(ByVal pnRecupCtasxCob As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nRecupCtasxCob = pnRecupCtasxCob
End Sub

Public Sub ActualizarFteIngVentas(ByVal pnVentas As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nVentas = pnVentas
End Sub

Public Sub ActualizarFteIngCreditosCmact(ByVal pnCreditosCmact As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nCreditosCmact = pnCreditosCmact
End Sub

Public Sub ActualizarFteIngCreditosOtros(ByVal pnCreditosOtros As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nCreditosOtros = pnCreditosOtros
End Sub

Public Sub ActualizarFteIngProveedores(ByVal pnProveedores As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nProveedores = pnProveedores
End Sub

Public Sub ActualizarFteIngPasivo(ByVal pnPasivo As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nPasivo = pnPasivo
End Sub

Public Sub ActualizarFteIngPasivoPatrimonio(ByVal pnPasivoPatrimonio As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nPasivoPatrimonio = pnPasivoPatrimonio
End Sub

Public Sub ActualizarFteIngBalIngFam(ByVal pnBalIngFam As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nPersIngFam = pnBalIngFam
End Sub

Public Sub ActualizarFteIngBalEgrFam(ByVal pnBalEgrFam As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nPersEgrFam = pnBalEgrFam
End Sub

Public Sub ActualizarFteIngActFijo(ByVal pnActFijo As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nActFijo = pnActFijo
End Sub

Public Sub ActualizarFteIngInventarios(ByVal pnInventarios As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nInventarios = pnInventarios
End Sub

Public Sub ActualizarFteIngCtasxCob(ByVal pnCtasxCob As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nCtasxCob = pnCtasxCob
End Sub

Public Sub ActualizarFteIngActivDisp(ByVal pnActivDisp As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nActivDisp = pnActivDisp
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

Public Sub ActualizarFteIngGastosFam(ByVal pnGastosFam As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoDep(pnPuntDet).nGastosFam = pnGastosFam
End Sub

Public Sub ActualizarFteIngConyugue(ByVal pnIngConyugue As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoDep(pnPuntDet).nIngCony = pnIngConyugue
End Sub

Public Sub ActualizarFteIngIngTotales(ByVal pnIngTotales As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nIngTotales = pnIngTotales
End Sub

Public Sub ActualizarFteIngIngOtros(ByVal pnIngOtros As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoDep(pnPuntDet).nIngOtros = pnIngOtros
End Sub

Public Sub ActualizarFteIngIngFam(ByVal pnIngFam As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoDep(pnPuntDet).nIngFam = pnIngFam
End Sub

Public Sub ActualizarFteIngIngresos(ByVal pnIngresos As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
    FtesIngreso(pnPunt).FtesIngresoDep(pnPuntDet).nIngresos = pnIngresos
End Sub

Public Sub ActualizarFteIngMoneda(ByVal psMoneda As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).sMoneda = psMoneda
End Sub

Public Sub ActualizarFteIngFechaInicioFuente(ByVal pdFteIngreso As Date, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).dFteIngreso = pdFteIngreso
End Sub

Public Sub ActualizarFteIngbCostoProd(ByVal pnCostoProd As Integer, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).bCostoProd = pnCostoProd
End Sub


Public Sub ActualizarFteIngCargo(ByVal psCargo As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).sCargo = psCargo
End Sub

Public Sub ActualizarFteIngFecCaduc(ByVal pdFecCaduc As Date, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer, ByVal pnTpoFte As PersFteIngresoTipo)
    If pnTpoFte = gPersFteIngresoTipoDependiente Then
        FtesIngreso(pnPunt).FtesIngresoDep(pnPuntDet).dFecCad = pdFecCaduc
    Else
        FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).dFecCad = pdFecCaduc
    End If
End Sub

Public Sub ActualizarFteDepIngFecEval(ByVal pdFecEval As Date, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoDep(pnPuntDet).dFecEval = pdFecEval
End Sub

Public Sub ActualizarFteIndIngFecEval(ByVal pdFecEval As Date, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).dFecEval = pdFecEval
End Sub

Public Sub ActualizarFteIngFuenteIngreso(ByVal psFuenteIngreso As String, ByVal psRazonSocial As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).sFteIngreso = psFuenteIngreso
     'FtesIngreso(pnPunt).sFteRazonSocial = psRazonSocial
End Sub

Public Sub ActualizarFteIngTipoAct(ByVal pCambio As TPersCambios, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nTipoAct = pCambio
End Sub

Public Sub ActualizarFteIngTipoActdetalle(ByVal pCambio As TPersCambios, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
    
        If FtesIngreso(pnPunt).nTipo = gPersFteIngresoTipoDependiente Then
            If FtesIngreso(pnPunt).FtesIngresoDep(pnPuntDet).nTipoAct <> PersFilaNueva Then
                FtesIngreso(pnPunt).FtesIngresoDep(pnPuntDet).nTipoAct = pCambio
            End If
        Else
            If FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nTipoAct <> PersFilaNueva Then
                FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nTipoAct = pCambio
            End If
        End If
        
        'If FtesIngreso(pnPunt).CostoProd(pnPuntDet).nTipoAct <> PersFilaNueva Then
         '       FtesIngreso(pnPunt).CostoProd(pnPuntDet).nTipoAct = pCambio
        'End If
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

Public Function ObtenerPersRelaEstado(ByVal pnPunt As Integer) As String
    ObtenerPersRelaEstado = Relac(pnPunt).sPersRelacEstado
End Function

Public Sub ActualizaRefComNombre(ByVal psNombre As String, ByVal pnPunt As Integer)
    RefCom(pnPunt).sRefComNom = psNombre
End Sub

Public Sub ActualizaRefComTipoRel(ByVal psTipoRel As String, ByVal pnPunt As Integer)
    RefCom(pnPunt).sRefComTipo = psTipoRel
End Sub

Public Sub ActualizaRefComFono(ByVal psFono As String, ByVal pnPunt As Integer)
    RefCom(pnPunt).sRefComFono = psFono
End Sub

Public Sub ActualizaRefComCod(ByVal pnCod As Integer, ByVal pnPunt As Integer)
    RefCom(pnPunt).nRefComCod = pnCod
End Sub

Public Sub ActualizaRefBanCodIF(ByVal psCodIF As String, ByVal pnPunt As Integer)
    RefBan(pnPunt).sRefBanCodIF = psCodIF
End Sub

Public Sub ActualizaRefBanNombre(ByVal psNomIF As String, ByVal pnPunt As Integer)
    RefBan(pnPunt).sRefBanNombre = psNomIF
End Sub

Public Sub ActualizaRefBanNumCta(ByVal psNumCta As String, ByVal pnPunt As Integer)
    RefBan(pnPunt).sRefBanNumCta = psNumCta
End Sub

Public Sub ActualizaRefBanNumTar(ByVal psNumTar As String, ByVal pnPunt As Integer)
    RefBan(pnPunt).sRefBanNumTar = psNumTar
End Sub

Public Sub ActualizaRefBanLinCred(ByVal pnLinCred As Currency, ByVal pnPunt As Integer)
    RefBan(pnPunt).nRefBanLinCred = pnLinCred
End Sub

Public Sub ActualizaPatVehCod(ByVal pnCod As Integer, ByVal pnPunt As Integer)
    PatVeh(pnPunt).nPatVehCod = pnCod
End Sub

Public Sub ActualizaPatVehMarca(ByVal psMarca As String, ByVal pnPunt As Integer)
    PatVeh(pnPunt).sPatVehMarca = psMarca
End Sub

Public Sub ActualizaPatVehFecFab(ByVal psFecFab As String, ByVal pnPunt As Integer)
    PatVeh(pnPunt).sPatVehFecFab = psFecFab
End Sub

Public Sub ActualizaPatVehValCom(ByVal pnValCom As Currency, ByVal pnPunt As Integer)
    PatVeh(pnPunt).nPatVehValCom = pnValCom
End Sub

Public Sub ActualizaPatVehCondicion(ByVal psCondicion As String, ByVal pnPunt As Integer)
    PatVeh(pnPunt).sPatVehCondicion = psCondicion
End Sub

Public Function ObtenerCostoProdnTpoCultivo(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Integer
    ObtenerCostoProdnTpoCultivo = FtesIngreso(pnPunt).CostoProd(pnPuntDet).nTpoCultivo
End Function

Public Function ObtenerCostoProdnMaquinaria(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerCostoProdnMaquinaria = FtesIngreso(pnPunt).CostoProd(pnPuntDet).nMaquinaria
End Function

Public Function ObtenerCostoProdnJornales(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerCostoProdnJornales = FtesIngreso(pnPunt).CostoProd(pnPuntDet).nJornales
End Function

Public Function ObtenerCostoProdnInsumos(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerCostoProdnInsumos = FtesIngreso(pnPunt).CostoProd(pnPuntDet).nInsumos
End Function

Public Function ObtenerCostoProdnPesticidas(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerCostoProdnPesticidas = FtesIngreso(pnPunt).CostoProd(pnPuntDet).nPesticidas
End Function

Public Function ObtenerCostoProdnOtros(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerCostoProdnOtros = FtesIngreso(pnPunt).CostoProd(pnPuntDet).nOtros
End Function

Public Function ObtenerCostoProdnHectareas(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerCostoProdnHectareas = FtesIngreso(pnPunt).CostoProd(pnPuntDet).nHectareas
End Function
Public Function ObtenerCostoProdnSiembra(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Integer
    ObtenerCostoProdnSiembra = FtesIngreso(pnPunt).CostoProd(pnPuntDet).nSiembra
End Function
Public Function ObtenerCostoProdnCosecha(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Integer
    ObtenerCostoProdnCosecha = FtesIngreso(pnPunt).CostoProd(pnPuntDet).nCosecha
End Function
Public Function ObtenerCostoProdnDesaAgricola(ByVal pnPunt As Integer, ByVal pnPutDet As Integer) As Integer
    ObtenerCostoProdnDesaAgricola = FtesIngreso(pnPunt).CostoProd(pnPutDet).nDesAgricola
End Function
Public Function ObtenerCostoProdnPlanOtros(ByVal pnPunt As Integer, ByVal pnPutDet As Integer) As Integer
    ObtenerCostoProdnPlanOtros = FtesIngreso(pnPunt).CostoProd(pnPutDet).nPlanOtros
End Function
Public Function ObtenerCostoProdnMantenimiento(ByVal pnPunt As Integer, ByVal pnPutDet As Integer) As Integer
    ObtenerCostoProdnMantenimiento = FtesIngreso(pnPunt).CostoProd(pnPutDet).nMantenimiento
End Function
Public Function ObtenerCostoProdnProduccion(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerCostoProdnProduccion = FtesIngreso(pnPunt).CostoProd(pnPuntDet).nProduccion
End Function
Public Function ObtenerCostoProdnUniProd(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerCostoProdnUniProd = FtesIngreso(pnPunt).CostoProd(pnPuntDet).nUniProd
End Function

Public Function ObtenerCostoProdnPreUni(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer) As Double
    ObtenerCostoProdnPreUni = FtesIngreso(pnPunt).CostoProd(pnPuntDet).nPreUni
End Function

Public Sub ActualizarCostosnTpoCultivo(ByVal nTpoCultivo As Integer, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).CostoProd(pnPuntDet).nTpoCultivo = nTpoCultivo
End Sub

Public Sub ActualizarCostoProdTipoAct(ByVal pCambio As TPersCambios, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).CostoProd(pnPuntDet).nTipoAct = pCambio
End Sub
Public Sub ActualizarCostosTpoCultivo(ByVal sTpoCultivo As String, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).CostoProd(pnPuntDet).sTpoCultivo = sTpoCultivo
End Sub

Public Sub ActualizarCostosnMaquinaria(ByVal nMaquinaria As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).CostoProd(pnPuntDet).nMaquinaria = nMaquinaria
End Sub

Public Sub ActualizarCostosnJornales(ByVal nJornales As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).CostoProd(pnPuntDet).nJornales = nJornales
End Sub

Public Sub ActualizarCostosnInsumos(ByVal nInsumos As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).CostoProd(pnPuntDet).nInsumos = nInsumos
End Sub

Public Sub ActualizarCostosnPesticidas(ByVal nPesticidas As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).CostoProd(pnPuntDet).nPesticidas = nPesticidas
End Sub

Public Sub ActualizarCostosnOtros(ByVal nOtros As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).CostoProd(pnPuntDet).nOtros = nOtros
End Sub

Public Sub ActualizarCostosnHectareas(ByVal nHectareas As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).CostoProd(pnPuntDet).nHectareas = nHectareas
End Sub

Public Sub ActualizarCostonSiembra(ByVal nSiembra As Integer, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
    FtesIngreso(pnPunt).CostoProd(pnPuntDet).nSiembra = nSiembra
End Sub

Public Sub ActualizarCostonCosecha(ByVal nCosecha As Integer, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
    FtesIngreso(pnPunt).CostoProd(pnPuntDet).nCosecha = nCosecha
End Sub

Public Sub ActualizarCostonDesaAgricola(ByVal nDesaAgricola As Integer, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
    FtesIngreso(pnPunt).CostoProd(pnPuntDet).nDesAgricola = nDesaAgricola
End Sub

Public Sub ActualizarCostonPlanOtros(ByVal nPlanOtros As Integer, ByVal pnPunt As Integer, ByVal pnPuntDt As Integer)
    FtesIngreso(pnPunt).CostoProd(pnPuntDt).nPlanOtros = nPlanOtros
End Sub

Public Sub ActualizarCostonMantenimiento(ByVal nMantenimiento As Integer, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
    FtesIngreso(pnPunt).CostoProd(pnPuntDet).nMantenimiento = nMantenimiento
End Sub

Public Sub ActualizarCostosnProduccion(ByVal nProduccion As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).CostoProd(pnPuntDet).nProduccion = nProduccion
End Sub

Public Sub ActualizarCostosnUniProd(ByVal nUniProd As Integer, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).CostoProd(pnPuntDet).nUniProd = nUniProd
End Sub

Public Sub ActualizarCostossUniProd(ByVal sUniProd As String, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).CostoProd(pnPuntDet).sUniProd = sUniProd
End Sub

Public Sub ActualizarCostosdFecEval(ByVal dFecEval As Date, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).CostoProd(pnPuntDet).dFecEval = dFecEval
End Sub

Public Sub ActualizarCostosnPreUni(ByVal nPreUni As Double, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).CostoProd(pnPuntDet).nPreUni = nPreUni
End Sub

Public Function ValidaPersonaPreferencial(ByVal pcPersCod As String) As Boolean
    Dim oConec As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    On Error GoTo ErrHandler
        sSql = "Select Count(*)  as nValor From Persona Where cPersCod='" & pcPersCod & "' AND bPreferencial=1"
        
        Set oConec = New COMConecta.DCOMConecta
        oConec.AbreConexion
        Set rs = oConec.CargaRecordSet(sSql)
        If Not rs.BOF And Not rs.EOF Then
            ValidaPersonaPreferencial = IIf(CInt(rs!nValor) > 0, True, False)
        End If
        Set rs = Nothing
        oConec.CierraConexion
        Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    If rs Is Nothing Then Set rs = Nothing
    ValidaPersonaPreferencial = False
End Function

Public Function GetInstFinancieras(Optional ByVal psFiltroTipoCtaIF As String = "") As Recordset
Dim oConec As COMConecta.DCOMConecta
Dim sql As String
Dim rs As ADODB.Recordset
Dim lsIFFiltro As String
Dim lsFiltroCta As String
Dim pos As String
Dim lsCadAux As String
Dim lsFiltroTipoIF As String
lsIFFiltro = ""
If psFiltroTipoCtaIF <> "" Then
    lsIFFiltro = " WHERE I.cIfTpo LIKE '" & psFiltroTipoCtaIF & "' "
End If

Set oConec = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

sql = "SELECT  cPersCod  AS cCodigo, cPersNombre as Descripcion, Nivel " _
    & " From " _
    & "         (   SELECT  I.cIFTpo cPersCod, c.cConsDescripcion cPersNombre, 1 as Nivel " _
    & "             FROM    INSTITUCIONFINANC I JOIN Constante c ON c.nConsValor = convert(int,I.cIFTpo) and c.nconscod like '" & gCGTipoIF & "' " _
    & "         " & lsIFFiltro & "" _
    & "             GROUP BY I.cIFTpo, c.cConsDescripcion " _
    & "             UNION ALL " _
    & "             SELECT  I.cIFTpo + '.' + P.cPersCod cPersCod, CONVERT(CHAR(30),P.cPersNombre ) AS cPersNombre , 2 AS Nivel " _
    & "             FROM    INSTITUCIONFINANC I  " _
    & "                     JOIN PERSONA P ON P.cPersCod = I.cPersCod  " & lsIFFiltro _
    & "         ) AS INSTFIN  " _
    & "     ORDER BY cPersCod "
Set rs = oConec.CargaRecordSet(sql)
Set GetInstFinancieras = rs
oConec.CierraConexion
Set oConec = Nothing
End Function



Public Function DigitoChequeo(ByVal psCorrelativo As String) As Integer
Dim liFactor As Variant
Dim liCadena(1 To 7) As Integer
Dim liSum, i As Integer
Dim lnDigito As Integer
    liFactor = Array(2, 7, 6, 5, 4, 3, 2)
    liCadena(1) = Val(Mid(psCorrelativo, 1, 1))
    liCadena(2) = Val(Mid(psCorrelativo, 2, 1))
    liCadena(3) = Val(Mid(psCorrelativo, 3, 1))
    liCadena(4) = Val(Mid(psCorrelativo, 4, 1))
    liCadena(5) = Val(Mid(psCorrelativo, 5, 1))
    liCadena(6) = Val(Mid(psCorrelativo, 6, 1))
    liCadena(7) = Val(Mid(psCorrelativo, 7, 1))
    liSum = 0
    For i = 1 To 7
        liSum = liSum + liCadena(i) * liFactor(i - 1)
    Next i
    lnDigito = 11 - (liSum Mod 11)
    If lnDigito = 10 Then
        DigitoChequeo = 0
    ElseIf lnDigito = 11 Then
        DigitoChequeo = 1
    Else
        DigitoChequeo = lnDigito
    End If
End Function

Public Function FillNum(intNumero As String, intLenNum As Integer, ChrFil As String) As String
  FillNum = Left(String(intLenNum, ChrFil), (Len(String(intLenNum, ChrFil)) - Len(Trim(intNumero)))) + Trim(intNumero)
End Function

'Para Evitar la Referencia Ciclica entre los Componentes

Public Function RecuperaHistoriaCredPersona(ByVal psPersCod As String) As Variant
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConn As COMConecta.DCOMConecta
Dim MatDatos() As String
Dim nPunt As Integer

    Set oConn = New COMConecta.DCOMConecta
    oConn.AbreConexion
    sSql = "Select TOP 3 P.cCtaCod, L.cDescripcion cLineaDesc, cAnalista=ISNULL(RH.cUser,''), PersApod.cPersNombre cApoderado, CN.cConsDescripcion cCondicion, Lin.nTasaIni as nTasaGracia, "
    sSql = sSql & " CN2.cConsDescripcion cDestino, P.nTasaInteres, CN3.cConsDescripcion cTipoCuota, C.dVigencia,"
    sSql = sSql & " nNota = (select nColocNota  From ColocCalificacionAnalista Where cCtaCod = P.cCtaCod And dColocNotaFecha = (select Max(dColocNotaFecha) from ColocCalificacionAnalista Where cCtaCod = P.cCtaCod)), "
    sSql = sSql & " CE.nMonto nMontoSol, CE.dPrdEstado dFecSol, CE.nCuotas nCuotasSol, CE.nPlazo nPlazoSol,"
    sSql = sSql & " CE2.nMonto nMontoSug, CE2.dPrdEstado dFecSug, CE2.nCuotas nCuotasSug, CE2.nPlazo nPlazoSug, CE2.nPeriodoGracia nPeriodoGraciaSug,"
    sSql = sSql & " CE3.nMonto nMontoApr, CE3.dPrdEstado dFecApr, CE3.nCuotas nCuotasApr, CE3.nPlazo nPlazoApr, CE3.nPeriodoGracia nPeriodoGraciaApr, CN8.cConsDescripcion cTipoGracia,"
    sSql = sSql & " CN4.cConsDescripcion cMotivoRech, CE5.dPrdEstado dFecCancel, CE6.dPrdEstado dFecJud, CC.cMetLiquidacion, CC.cProtesto,"
    sSql = sSql & " CE7.nPrdEstado nEstRefin, CC.bCargoAuto, CN5.cConsDescripcion cEstActual, CN6.cConsDescripcion as cTipoCredDescrip, "
    sSql = sSql & " CN7.cConsDescripcion cTipoDesemb, CC.nNroProxDesemb, "
    sSql = sSql & " nCuotaSug = (Select sum(nMonto) from ColocCalendDet Cal Where Cal.cCtaCod = P.cCtaCod AND Cal.nNrocalen =1 AND Cal.nColocCalendApl=1 AND Cal.nCuota=1) " & " ,"
    sSql = sSql & " nCuotaApr = (Select sum(nMonto) from ColocCalendDet Cal Where Cal.cCtaCod = P.cCtaCod AND Cal.nNrocalen =1 AND Cal.nColocCalendApl=1 AND Cal.nCuota=2) "
    sSql = sSql & " From Producto P Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod AND PP.nPrdpersRelac = " & gColRelPersTitular
    sSql = sSql & "                 Inner Join Colocaciones C ON C.cCtaCod = P.cCtaCod "
    sSql = sSql & "                 Left Join ColocLineaCredito L ON L.cLineaCred = C.cLineaCred"
    sSql = sSql & "                 Left Join ProductoPersona PP2 ON PP2.cCtaCod = P.cCtaCod AND PP2.nPrdPersRelac = " & gColRelPersAnalista
    sSql = sSql & "                 Left Join RRHH RH ON PP2.cPersCod = RH.cPersCod "
    sSql = sSql & "                 Left Join ProductoPersona PP3 ON PP3.cCtaCod = P.cCtaCod AND PP3.nPrdPersRelac = " & gColRelPersApoderado
    sSql = sSql & "                 Left Join Persona PersApod ON PersApod.cPersCod = PP3.cPersCod"
    sSql = sSql & "                 Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod"
    sSql = sSql & "                 Left Join Constante CN ON CN.nConsValor = CC.nColocCondicion AND CN.nConsCod = " & gColocCredCondicion
    sSql = sSql & "                 Left Join Constante CN2 ON CN2.nConsValor = CC.nColocDestino AND CN2.nConsCod = " & gColocDestino
    sSql = sSql & "                 Left Join Constante CN3 ON CN3.nConsValor = CC.nColocCalendCod AND CN3.nConsCod = " & gColocTipoCalend
    sSql = sSql & "                 Left Join ColocacEstado CE ON CE.cCtaCod = P.cCtaCod AND CE.nPrdEstado = " & gColocEstSolic
    sSql = sSql & "                 Left Join ColocacEstado CE2 ON CE2.cCtaCod = P.cCtaCod AND CE2.nPrdEstado = " & gColocEstSug
    sSql = sSql & "                 Left Join ColocacEstado CE3 ON CE3.cCtaCod = P.cCtaCod AND CE3.nPrdEstado = " & gColocEstAprob
    sSql = sSql & "                 Left Join ColocacEstado CE4 ON CE4.cCtaCod = P.cCtaCod AND CE4.nPrdEstado = " & gColocEstRech
    sSql = sSql & "                 Left Join Constante CN4 ON CN4.nConsValor = CE4.nMotivoRechazo AND CN4.nConsCod = " & gColocMotivRechazo
    sSql = sSql & "                 Left Join ColocacEstado CE5 ON CE5.cCtaCod = P.cCtaCod AND CE5.nPrdEstado = " & gColocEstCancelado
    sSql = sSql & "                 Left Join ColocacEstado CE6 ON CE6.cCtaCod = P.cCtaCod AND CE6.nPrdEstado = " & gColocEstJudicial
    sSql = sSql & "                 Left Join ColocacEstado CE7 ON CE7.cCtaCod = P.cCtaCod AND CE7.nPrdEstado = " & gColocEstRefNorm
    sSql = sSql & "                 Left Join Constante CN5 ON CN5.nConsValor = P.nPrdEstado AND CN5.nConsCod = " & gColocEstado
    sSql = sSql & "                 Left Join Constante CN6 ON convert(int,substring(P.cCtaCod,6,3)) = CN6.nConsValor AND CN6.nConsCod = " & gProducto
    sSql = sSql & "                 Left Join Constante CN7 ON CN7.nConsValor = CC.nTipoDesembolso AND CN7.nConsCod = " & gColocTiposDesembolso
    sSql = sSql & "                 Left Join Constante CN8 ON CN8.nConsValor = CE3.nTipoGracia AND CN8.nConsCod = " & gColocTiposGracia
    sSql = sSql & "                 Left Join ColocLineaCreditoTasa Lin ON C.cLineaCred = Lin.cLineaCred AND Lin.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntGracia
    sSql = sSql & " WHERE PP.cPersCod = '" & psPersCod & "' AND C.dVigencia is not NULL"
    sSql = sSql & " ORDER BY C.dVigencia DESC "
    Set R = oConn.CargaRecordSet(sSql)
    ReDim MatDatos(R.RecordCount, 9)
    nPunt = 0
    Do While Not R.EOF
        MatDatos(nPunt, 0) = R!cCtaCod
        MatDatos(nPunt, 1) = R!cEstActual
        MatDatos(nPunt, 2) = R!dVigencia
        MatDatos(nPunt, 3) = IIf(IsNull(R!dFecCancel), "", R!dFecCancel)
        If Not IsNull(R!nMontoApr) Then
            MatDatos(nPunt, 4) = Format(R!nMontoApr, "#0.00")
        Else
            If Not IsNull(R!nMontoSug) Then
                MatDatos(nPunt, 4) = Format(R!nMontoSug, "#0.00")
            Else
                MatDatos(nPunt, 4) = Format(R!nMontoSol, "#0.00")
            End If
        End If
        If Not IsNull(R!nCuotasApr) Then
            MatDatos(nPunt, 5) = R!nCuotasApr & "/" & R!nPlazoApr
        Else
            If Not IsNull(R!nCuotasSug) Then
                MatDatos(nPunt, 5) = R!nCuotasSug & "/" & R!nPlazoSug
            Else
                MatDatos(nPunt, 5) = R!nCuotasSol & "/" & R!nPlazoSol
            End If
        End If
        If Not IsNull(R!nCuotaApr) Then
            MatDatos(nPunt, 6) = Format(R!nCuotaApr, "#0.00")
        Else
            MatDatos(nPunt, 6) = Format(R!nCuotaSug, "#0.00")
        End If
        MatDatos(nPunt, 7) = Trim(Str(IIf(IsNull(R!nNota), 0, R!nNota)))
        MatDatos(nPunt, 8) = Trim(R!cAnalista)
        nPunt = nPunt + 1
        R.MoveNext
    Loop
    R.Close
    oConn.CierraConexion
    RecuperaHistoriaCredPersona = MatDatos
End Function

Private Sub Class_Terminate()
    Conn.CierraConexion
    Set Conn = Nothing
End Sub


'Datos de la Fuente de Ingreso como Persona

Public Sub ActualizarFteRuc(ByVal psRUC As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).cFtePersRuc = psRUC
End Sub
Public Sub ActualizarFteFecInicioAct(ByVal pdFecInicioAct As Date, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).dFtePersFecInicioAct = pdFecInicioAct
End Sub
Public Sub ActualizarFteTipoPersJur(ByVal psTipoPersJur As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).cFtePersTipoPersJur = psTipoPersJur
End Sub
Public Sub ActualizarFteTelefono(ByVal psTelefono As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).cFtePersTelefono = psTelefono
End Sub
Public Sub ActualizarFteCIUU(ByVal psCIUU As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).cFtePersCIUU = psCIUU
End Sub
Public Sub ActualizarFteCondicionDomic(ByVal psCondicionDomic As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).cFtePersCondicionDomic = psCondicionDomic
End Sub
Public Sub ActualizarFteMagnitudEmp(ByVal psMagnitudEmp As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).cFtePersMagnitudEmp = psMagnitudEmp
End Sub
Public Sub ActualizarFteNroEmpleados(ByVal pnNroEmpleados As Double, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nFtePersNroEmpleados = pnNroEmpleados
End Sub
Public Sub ActualizarFteDireccion(ByVal psDireccion As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).cFtePersDireccion = psDireccion
End Sub
Public Sub ActualizarFteDpto(ByVal psDpto As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).cFtePersDpto = psDpto
End Sub
Public Sub ActualizarFteProv(ByVal psProv As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).cFtePersProv = psProv
End Sub
Public Sub ActualizarFteDist(ByVal psDist As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).cFtePersDist = psDist
End Sub
Public Sub ActualizarFteZona(ByVal psZona As String, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).cFtePersZona = psZona
End Sub

'**Modificado por ELRO el 20110721, según Acta Nº 168-2011/TI-D

Public Function recuperarDatosPrincipalesClienteProcesoReforzado(ByVal pvTitular As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperarDatosPrincipalesClienteProcesoReforzado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_sel_DatosPrincipalesClienteProcesoReforzado '" & pvTitular & "'"
Set recuperarDatosPrincipalesClienteProcesoReforzado = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperarDatosPrincipalesClienteProcesoReforzado:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperarDatosPrincipalesClienteProcesoReforzado", Err.Description

End Function

Public Sub registrarClienteProcesoReforzado(ByVal pvTitular As String, _
                                            ByVal pvCentroLabor As String, _
                                            ByVal pvFechaRegistro As String, _
                                            ByVal pvPersonaAtencion As String, _
                                            ByVal pvPersonaAutoriza As String, _
                                            ByVal pvTipoVinculo As String)
                 
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRegistrarClienteProcesoReforzado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_ins_ClienteProcesoReforzado '" & pvTitular & "','" & pvCentroLabor & "','" & Format(pvFechaRegistro, "MM/dd/yyyy") & "','" & pvPersonaAtencion & "','" & pvPersonaAutoriza & "','" & pvTipoVinculo & "'"
oConecta.ConexionActiva.Execute sSql
oConecta.CierraConexion
Set oConecta = Nothing
Exit Sub

ErrorRegistrarClienteProcesoReforzado:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RegistrarClienteProcesoReforzado", Err.Description

End Sub

Public Function recuperarLlaveClienteProcesoReforzado(ByVal pvTitular As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperarLlaveClienteProcesoReforzado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_sel_LlaveClienteProcesoReforzado '" & pvTitular & "'"
Set recuperarLlaveClienteProcesoReforzado = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperarLlaveClienteProcesoReforzado:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperarLlaveClienteProcesoReforzado", Err.Description

End Function
    
Public Function recuperarParentescosClienteReforzado(ByVal pcPariente As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperarParentescosClienteReforzado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_sel_RelacionClienteReforzado '" & pcPariente & "'"
Set recuperarParentescosClienteReforzado = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperarParentescosClienteReforzado:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperarParentescosClienteReforzado", Err.Description

End Function

Public Function recuperarTiposDocumentosClienteReforzado() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperarTiposDocumentosClienteReforzado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_sel_TipoDocumentoClienteReforzado"
Set recuperarTiposDocumentosClienteReforzado = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperarTiposDocumentosClienteReforzado:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperarTiposDocumentosClienteReforzado", Err.Description

End Function

Public Function recuperarPaisesClienteReforzado(ByVal pcUbiGeoCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperarPaisesClienteReforzado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_sel_UbicacionGeografica '" & pcUbiGeoCod & "'"
Set recuperarPaisesClienteReforzado = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperarPaisesClienteReforzado:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperarPaisesClienteReforzado", Err.Description

End Function

Public Function recuperarOcupacionesClienteReforzado() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperarOcupacionesClienteReforzado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_sel_Ocupaciones"
Set recuperarOcupacionesClienteReforzado = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperarOcupacionesClienteReforzado:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperarOcupacionesClienteReforzado", Err.Description

End Function

Public Function recuperarBancosClienteReforzado() As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperarBancosClienteReforzado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_sel_Bancos"
Set recuperarBancosClienteReforzado = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperarBancosClienteReforzado:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperarBancosClienteReforzado", Err.Description

End Function

Public Sub registrarParentescoClienteReforzado(ByVal pnClienteProcesoReforzadoId As Integer, _
                                               ByVal pnRelacionId As Integer, _
                                               ByVal pvNombreCompleto As String, _
                                               ByVal pvDni As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRegistrarParentescoClienteReforzado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_ins_ParienteClienteReforzado '" & pnClienteProcesoReforzadoId & "','" & pnRelacionId & "','" & pvNombreCompleto & "','" & pvDni & "'"
oConecta.ConexionActiva.Execute sSql
oConecta.CierraConexion
Set oConecta = Nothing
Exit Sub

ErrorRegistrarParentescoClienteReforzado:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RegistrarParentescoClienteReforzado", Err.Description

End Sub
Public Sub registrarConyugeClienteReforzado(ByVal pnClienteProcesoReforzadoId As Integer, _
                                            ByVal pvNombres As String, _
                                            ByVal pvApellidos As String, _
                                            ByVal pnConsCod As Integer, _
                                            ByVal pnConsValor As Integer, _
                                            ByVal pvNumeroDoi As String, _
                                            ByVal pcUbiGeoCod As String, _
                                            ByVal pvFechaNacimiento As String, _
                                            ByVal pvCentroLaboral As String, _
                                            ByVal pnOcupaciones As String, _
                                            ByVal pmIngresoPromedio As Double)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRegistrarConyugeClienteReforzado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_ins_ConyugeClienteReforzado '" & pnClienteProcesoReforzadoId & "','" & pvNombres & "','" & pvApellidos & "','" & pnConsCod & "' ,'" & pnConsValor & "' ,'" & pvNumeroDoi & "' ,'" & pcUbiGeoCod & "' ,'" & Format(pvFechaNacimiento, "MM/dd/yyyy") & "' ,'" & pvCentroLaboral & "' ,'" & pnOcupaciones & "',1011,1,'" & pmIngresoPromedio & "'"
oConecta.ConexionActiva.Execute sSql
oConecta.CierraConexion
Set oConecta = Nothing
Exit Sub

ErrorRegistrarConyugeClienteReforzado:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RegistrarConyugeClienteReforzado", Err.Description

End Sub

Public Function recuperarLlaveConyuge(ByVal pnClienteProcesoReforzadoId As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperarLlaveConyuge

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_sel_LLaveConyuge '" & pnClienteProcesoReforzadoId & "'"
Set recuperarLlaveConyuge = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperarLlaveConyuge:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperarLlaveConyuge", Err.Description

End Function

Public Sub registrarParentescoConyuge(ByVal pnConyugeId As Integer, _
                                      ByVal pnRelacionId As Integer, _
                                      ByVal pvNombreCompleto As String, _
                                      ByVal pvDni As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRegistrarParentescoConyuge

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_ins_ParienteConyuge '" & pnConyugeId & "','" & pnRelacionId & "','" & pvNombreCompleto & "','" & pvDni & "'"
oConecta.ConexionActiva.Execute sSql
oConecta.CierraConexion
Set oConecta = Nothing
Exit Sub

ErrorRegistrarParentescoConyuge:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RegistrarParentescoConyuge", Err.Description

End Sub

Public Sub registrarReferenciaEconomica(ByVal pnClienteProcesoReforzadoId As Integer, _
                                        ByVal pvCargoPublico As String, _
                                        ByVal pvFechaInicio As String, _
                                        ByVal pvFechaCese As String, _
                                        ByVal pvEntidadLabora As String, _
                                        ByVal pdIngreso As Double)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRegistrarReferenciaEconomica

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_ins_ReferenciaEcononica '" & pnClienteProcesoReforzadoId & "','" & pvCargoPublico & "','" & Format(pvFechaInicio, "MM/dd/yyyy") & "','" & Format(pvFechaCese, "MM/dd/yyyy") & "','" & pvEntidadLabora & "',1011,1,'" & pdIngreso & "'"
oConecta.ConexionActiva.Execute sSql
oConecta.CierraConexion
Set oConecta = Nothing
Exit Sub

ErrorRegistrarReferenciaEconomica:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RegistrarReferenciaEconomica", Err.Description

End Sub

Public Sub registrarReferenciaFinanciera(ByVal pnClienteProcesoReforzadoId As Integer, _
                                         ByVal pvBanco As String, _
                                         ByVal pvTipoProducto As String, _
                                         ByVal pvFuncionarioNegocio As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRegistrarReferenciaFinanciera

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_ins_ReferenciaFinanciera '" & pnClienteProcesoReforzadoId & "','" & pvBanco & "','" & pvTipoProducto & "','" & pvFuncionarioNegocio & "'"
oConecta.ConexionActiva.Execute sSql
oConecta.CierraConexion
Set oConecta = Nothing
Exit Sub

ErrorRegistrarReferenciaFinanciera:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RegistrarReferenciaFinanciera", Err.Description

End Sub

Public Sub registrarReferenciaPatrimonial(ByVal pnClienteProcesoReforzadoId As Integer, _
                                          ByVal pvBien As String, _
                                          ByVal pdValor As Double)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRegistrarReferenciaPatrimonial

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_ins_ReferenciaPatrimonial '" & pnClienteProcesoReforzadoId & "',P,'" & pvBien & "',1011,2,'" & pdValor & "'"
oConecta.ConexionActiva.Execute sSql
oConecta.CierraConexion
Set oConecta = Nothing
Exit Sub

ErrorRegistrarReferenciaPatrimonial:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RegistrarReferenciaPatrimonial", Err.Description

End Sub

Public Sub registrarOtroBien(ByVal pnClienteProcesoReforzadoId As Integer, _
                             ByVal pvBien As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRegistrarOtroBien

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_ins_ReferenciaPatrimonial '" & pnClienteProcesoReforzadoId & "',O,'" & pvBien & "',1011,2,0"
oConecta.ConexionActiva.Execute sSql
oConecta.CierraConexion
Set oConecta = Nothing
Exit Sub

ErrorRegistrarOtroBien:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RegistrarOtroBien", Err.Description

End Sub

Public Sub registrarPersonaJuridica(ByVal pnClienteProcesoReforzadoId As Integer, _
                                    ByVal pvNombreEmpresa As String, _
                                    ByVal pvRuc As String, _
                                    ByVal pvParticipacion As Integer, _
                                    ByVal pdIngreso As Double)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRegistrarPersonaJuridica

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_ins_PersonaJuridica '" & pnClienteProcesoReforzadoId & "','" & pvNombreEmpresa & "','" & pvRuc & "','" & pvParticipacion & "',1011,1,'" & pdIngreso & "'"
oConecta.ConexionActiva.Execute sSql
oConecta.CierraConexion
Set oConecta = Nothing
Exit Sub

ErrorRegistrarPersonaJuridica:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RegistrarPersonaJuridica", Err.Description

End Sub

Public Sub registrarProveedorCliente(ByVal pnClienteProcesoReforzadoId As Integer, _
                                     ByVal pcTipo As String, _
                                     ByVal pvNombreEmpresa As String)
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRegistrarProveedorCliente

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_ins_ProveedorCliente '" & pnClienteProcesoReforzadoId & "','" & pcTipo & "','" & pvNombreEmpresa & "'"
oConecta.ConexionActiva.Execute sSql
oConecta.CierraConexion
Set oConecta = Nothing
Exit Sub

ErrorRegistrarProveedorCliente:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RegistrarProveedorCliente", Err.Description

End Sub

Public Function recuperarDatosClienteProcesoReforzado(ByVal pnClienteProcesoReforzadoId As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperarDatosClienteProcesoReforzado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_sel_ClienteProcesoReforzado '" & pnClienteProcesoReforzadoId & "'"
Set recuperarDatosClienteProcesoReforzado = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperarDatosClienteProcesoReforzado:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperarDatosClienteProcesoReforzado", Err.Description

End Function

Public Function recuperarDatosConyugeClienteProcesoReforzado(ByVal pnClienteProcesoReforzadoId As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperarDatosConyugeClienteProcesoReforzado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_sel_ConyugeClienteProcesoReforzado '" & pnClienteProcesoReforzadoId & "'"
Set recuperarDatosConyugeClienteProcesoReforzado = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperarDatosConyugeClienteProcesoReforzado:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperarDatosConyugeClienteProcesoReforzado", Err.Description

End Function

Public Function recuperarReferenciaEconomicaClienteProcesoReforzado(ByVal pnClienteProcesoReforzadoId As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperarReferenciaEconomicaClienteProcesoReforzado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_sel_ReferenciaEcononica '" & pnClienteProcesoReforzadoId & "'"
Set recuperarReferenciaEconomicaClienteProcesoReforzado = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperarReferenciaEconomicaClienteProcesoReforzado:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperarReferenciaEconomicaClienteProcesoReforzado", Err.Description

End Function

Public Function recuperarReferenciaFinancieraClienteProcesoReforzado(ByVal pnClienteProcesoReforzadoId As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperarReferenciaFinancieraClienteProcesoReforzado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_sel_ReferenciaFinanciera '" & pnClienteProcesoReforzadoId & "'"
Set recuperarReferenciaFinancieraClienteProcesoReforzado = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperarReferenciaFinancieraClienteProcesoReforzado:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperarReferenciaFinancieraClienteProcesoReforzado", Err.Description

End Function

Public Function recuperarReferenciaPatrimonialClienteProcesoReforzado(ByVal pnClienteProcesoReforzadoId As Integer, ByVal pcTipoPatrimonio As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperarReferenciaPatrimonialClienteProcesoReforzado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_sel_ReferenciaPatrimonial '" & pnClienteProcesoReforzadoId & "','" & pcTipoPatrimonio & "'"
Set recuperarReferenciaPatrimonialClienteProcesoReforzado = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperarReferenciaPatrimonialClienteProcesoReforzado:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperarReferenciaPatrimonialClienteProcesoReforzado", Err.Description

End Function

Public Function recuperarPersonaJuridicaClienteProcesoReforzado(ByVal pnClienteProcesoReforzadoId As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperarPersonaJuridicaClienteProcesoReforzado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_sel_PersonaJuridica '" & pnClienteProcesoReforzadoId & "'"
Set recuperarPersonaJuridicaClienteProcesoReforzado = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperarPersonaJuridicaClienteProcesoReforzado:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperarPersonaJuridicaClienteProcesoReforzado", Err.Description

End Function

Public Function recuperarProveedorClienteProcesoReforzado(ByVal pnClienteProcesoReforzadoId As Integer, ByVal pcTipo As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperarProveedorClienteProcesoReforzado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_sel_ProveedorCliente '" & pnClienteProcesoReforzadoId & "','" & pcTipo & "'"
Set recuperarProveedorClienteProcesoReforzado = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperarProveedorClienteProcesoReforzado:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperarProveedorClienteProcesoReforzado", Err.Description

End Function

Public Function recuperarParienteClienteProcesoReforzado(ByVal pnClienteProcesoReforzadoId As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperarParienteClienteProcesoReforzado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_sel_ParienteClienteReforzado '" & pnClienteProcesoReforzadoId & "'"
Set recuperarParienteClienteProcesoReforzado = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperarParienteClienteProcesoReforzado:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperarParienteClienteProcesoReforzado", Err.Description

End Function

Public Function recuperarParienteConyugeClienteProcesoReforzado(ByVal pnConyugeId As Integer) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperarParienteConyugeClienteProcesoReforzado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_sel_ParienteConyuge '" & pnConyugeId & "'"
Set recuperarParienteConyugeClienteProcesoReforzado = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperarParienteConyugeClienteProcesoReforzado:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperarParienteConyugeClienteProcesoReforzado", Err.Description

End Function

Public Function recuperarPersNegativaAutorizacion(ByVal pfecsol As Date, ByVal pNombre As String, ByVal pUser As String, Optional pbolCond As Boolean = False, Optional ByVal pnPreAut As Boolean = False) As ADODB.Recordset 'WIOR 20121123 AGREGO pnPreAut

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperarPersNegativaAutorizacion

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_val_ValidaAutorizacionSensibles '" & Format(pfecsol, "yyyymmdd") & "','" & pNombre & "','" & pUser & "', 1 ," & IIf(pnPreAut, 1, 0) 'WIOR 20121123 AGREGO pnPreAut
Set recuperarPersNegativaAutorizacion = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperarPersNegativaAutorizacion:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperarPersNegativaAutorizacion", Err.Description

End Function

Public Function recuperarPersListaNegativa(ByVal pnTipoDoc As Integer, ByVal pcNumDoc As String, ByVal pnCondicion As Integer) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperarPersListaNegativa

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_sel_VerificaPersListaNegativa_4 '" & pnTipoDoc & "','" & pcNumDoc & "','" & pnCondicion & "' "
Set recuperarPersListaNegativa = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperarPersListaNegativa:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperarPersListaNegativa", Err.Description

End Function

Public Function recuperarRelacionEntidadClienteProcesoReforzado(ByVal pvTitular As String) As ADODB.Recordset

Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperarRelacionEntidadClienteProcesoReforzado

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_sel_RecuperarRelacionEntidadClienteProcesoReforzado '" & pvTitular & "' "
Set recuperarRelacionEntidadClienteProcesoReforzado = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperarRelacionEntidadClienteProcesoReforzado:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperarRelacionEntidadClienteProcesoReforzado", Err.Description

End Function

'***Fin modificado por ELRO***************************************
'EJVG 20120322
Public Function obtenerFechaUltimaActualizacion(ByVal psPersCod As String) As Date
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
On Error GoTo ErrorObtenerFechaUltimaActualizacion
    Set oConecta = New COMConecta.DCOMConecta
    Set rs = New ADODB.Recordset
    oConecta.AbreConexion
    sSql = "Select convert(date,left(isnull(cUltimaActualizacion,'19000101'),8)) dFechaUltAct from Persona where cPersCod = '" & psPersCod & "'"
    Set rs = oConecta.CargaRecordSet(sSql)
    If Not rs.BOF And Not rs.EOF Then
        obtenerFechaUltimaActualizacion = rs!dFechaUltAct
    Else
        obtenerFechaUltimaActualizacion = "1900-01-01"
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerFechaUltimaActualizacion:
    Set oConecta = Nothing
    Err.Raise Err.Number, "DCOMPersona:obtenerFechaUltimaActualizacion", Err.Description
End Function
'WIOR 20120511*********************************************************************************
Public Function ObtenerEdadPersona(ByVal pdFecha As Date, ByVal psPersCod As String) As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    sSql = "exec stp_sel_ObtenerEdadPersona '" & Format(pdFecha, "mm/dd/yyyy") & "','" & psPersCod & "'"
    Set ObtenerEdadPersona = oCon.CargaRecordSet(sSql)
      
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'WIOR FIN**************************************************************************************
'WIOR 20120724*********************************************************************************
Public Function ObtenerVinculadosPersona(ByVal psPersCod As String, ByVal psCodPersonas As String) As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String
On Error GoTo ErrorObtenerVinculadosPersona
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    sSql = "exec stp_sel_ListaVinculadosRiesgo '" & psPersCod & "','" & psCodPersonas & "'"
    Set ObtenerVinculadosPersona = oCon.CargaRecordSet(sSql)
      
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
ErrorObtenerVinculadosPersona:
    Err.Raise Err.Number, "DCOMPersona:ObtenerVinculadosPersona", Err.Description
End Function

Public Sub RegistrarVinculadosPersona(Optional ByVal psCtaCod As String = "", _
                                      Optional ByVal psPersCod As String = "", _
                                      Optional ByVal psPersCodVin As String = "", _
                                      Optional ByVal pnSaldoVin As Double = 0, _
                                      Optional ByVal pnTpoVinculo As Double = 0, _
                                      Optional ByVal pdFechaVinculo As Date = "01/01/1900", _
                                      Optional ByVal pnNivelVinculo As Integer = 0)
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
On Error GoTo ErrorRegistrarVinculadosPersona

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion

sSql = "INSERT INTO PersVinculados (cCtaCod,cPerscod,cPersCodVin, nSaldoVin,nTpoVinculo,FechaVinculo,NivelVinculado)" & _
        " VALUES ('" & psCtaCod & "','" & psPersCod & "','" & psPersCodVin & "'," & pnSaldoVin & "," & _
        pnTpoVinculo & ",'" & Format(pdFechaVinculo, "yyyymmdd") & "'," & pnNivelVinculo & ")"
        
oConecta.ConexionActiva.Execute sSql
oConecta.CierraConexion
Set oConecta = Nothing

    Exit Sub
ErrorRegistrarVinculadosPersona:
    Err.Raise Err.Number, "DCOMPersona:RegistrarVinculadosPersona", Err.Description
End Sub
Public Sub EliminarVinculadosPersona(ByVal psCtaCod As String)
    
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
On Error GoTo ErrorEliminarVinculadosPersona

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion

sSql = "DELETE FROM PersVinculados where cCtaCod='" & psCtaCod & "'"
        
oConecta.ConexionActiva.Execute sSql
oConecta.CierraConexion
Set oConecta = Nothing

Exit Sub
ErrorEliminarVinculadosPersona:
    Err.Raise Err.Number, "DCOMPersona:EliminarVinculadosPersona", Err.Description
End Sub

Public Function ListaVinculadosPersona(ByVal psCtaCod As String, Optional ByVal pnTipo As Integer = 1, Optional ByVal pnNivel As Integer = 0) As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String
On Error GoTo ErrorListaVinculadosPersona
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    sSql = "select cCtaCod,cPerscod,cPersCodVin,nSaldoVin,nTpoVinculo,FechaVinculo,NivelVinculado " & _
            " from PersVinculados where cctacod='" & psCtaCod & "' " & _
            " and nTpoVinculo=" & pnTipo & " and NivelVinculado=" & pnNivel

    Set ListaVinculadosPersona = oCon.CargaRecordSet(sSql)
      
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
    
ErrorListaVinculadosPersona:
    Err.Raise Err.Number, "DCOMPersona:ListaVinculadosPersona", Err.Description
End Function
Public Function DelvolverCodViculados(ByVal psCtaCod As String) As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String
On Error GoTo ErrorDelvolverCodViculados
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    sSql = "select dbo.DelvolverCodViculados('" & psCtaCod & "') CodPersonas"
            
    Set DelvolverCodViculados = oCon.CargaRecordSet(sSql)
      
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
    
ErrorDelvolverCodViculados:
    Err.Raise Err.Number, "DCOMPersona:DelvolverCodViculados", Err.Description
End Function
Public Function ListaVinculadosPersonaCta(ByVal psCtaCod As String) As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String
On Error GoTo ErrorListaVinculadosPersonaCta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    sSql = "select cCtaCod,cPerscod,cPersCodVin,nSaldoVin,nTpoVinculo,FechaVinculo,NivelVinculado " & _
            " from PersVinculados where cctacod='" & psCtaCod & "' " & _
            " order by nTpoVinculo,NivelVinculado,cPerscod"

    Set ListaVinculadosPersonaCta = oCon.CargaRecordSet(sSql)
      
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
    
ErrorListaVinculadosPersonaCta:
    Err.Raise Err.Number, "DCOMPersona:ListaVinculadosPersonaCta", Err.Description
End Function

Public Function DeterminarRiesgo(ByVal psCtaCod As String, ByVal pnTipo As Integer, ByVal pnRiesgo As Integer) As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String
On Error GoTo ErrorDeterminarRiesgo
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    'WIOR 230130917 INICIO COMENTÓ
    ''Agropecuarios
    'If pnTipo = 1 Then
    '    'RANGO 1
    '    If pnRiesgo = 1 Then
    '        sSql = " SELECT  11 Riesgo FROM Colocaciones CO " & _
    '               " INNER JOIN ColocacCred CCR ON CCR.cCtaCod =CO.cCtaCod" & _
    '               " Inner Join" &
    '               " (SELECT DISTINCT cCtaCod  FROM ColocGarantia CG" & _
    '               " INNER JOIN Garantias G ON G.cNumGarant=CG.cNumGarant" & _
    '               " AND G.nTpoGarantia IN (4,30))X ON X.cCtaCod =CO.cCtaCod" & _
    '               " WHERE CO.cCtaCod ='" & psCtaCod & "' " & _
    '               " " 'AND CO.cTpoCredCod in (152,252,352,452,552)"
    '    'RANGO 2
    '    ElseIf pnRiesgo = 2 Then
    '        sSql = " SELECT 12 Riesgo FROM Colocaciones CO " & _
    '               " INNER JOIN ColocacCred CCR ON CCR.cCtaCod =CO.cCtaCod" & _
    '               " Inner Join" &
    '               " (SELECT DISTINCT cCtaCod  FROM ColocGarantia CG" & _
    '               " INNER JOIN Garantias G ON G.cNumGarant=CG.cNumGarant" & _
    '               " AND G.nTpoGarantia IN (1,2,6,11,12,19))X ON X.cCtaCod =CO.cCtaCod" & _
    '               " WHERE CO.cCtaCod ='" & psCtaCod & "'" & _
    '               " " 'AND CO.cTpoCredCod in (152,252,352,452,552)"
    '    End If
    ''OTROS CRDITOS
    'ElseIf pnTipo = 2 Then
    '    'RANGO 1
    '    If pnRiesgo = 1 Then
    '        sSql = " SELECT 21 Riesgo FROM Colocaciones CO " & _
    '               " INNER JOIN ColocacCred CCR ON CCR.cCtaCod =CO.cCtaCod" & _
    '               " Inner Join" &
    '               " (SELECT DISTINCT cCtaCod  FROM ColocGarantia CG" & _
    '               " INNER JOIN Garantias G ON G.cNumGarant=CG.cNumGarant" & _
    '               " AND G.nTpoGarantia IN (4,30))X ON X.cCtaCod =CO.cCtaCod" & _
    '               " WHERE CO.cCtaCod ='" & psCtaCod & "' " & _
    '               " " 'AND CO.cTpoCredCod not in (152,252,352,452,552)"
    '    'RANGO 2
    '    ElseIf pnRiesgo = 2 Then
    '        sSql = " SELECT 22 Riesgo FROM Colocaciones CO " & _
    '               " INNER JOIN ColocacCred CCR ON CCR.cCtaCod =CO.cCtaCod" & _
    '               " Inner Join" &
    '               " (SELECT DISTINCT cCtaCod  FROM ColocGarantia CG" & _
    '               " INNER JOIN Garantias G ON G.cNumGarant=CG.cNumGarant" & _
    '               " AND G.nTpoGarantia IN (1,2,6,11,12,19))X ON X.cCtaCod =CO.cCtaCod" & _
    '               " WHERE CO.cCtaCod ='" & psCtaCod & "' " & _
    '               " " 'AND CO.cTpoCredCod not in (152,252,352,452,552)"
    '    End If
    'End If
    'WIOR 230130917 FIN COMENTÓ
    sSql = "EXEC stp_sel_DeterminarRiesgo '" & psCtaCod & "'," & pnTipo & "," & pnRiesgo 'WIOR 20130917
    Set DeterminarRiesgo = oCon.CargaRecordSet(sSql)
    
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
    
ErrorDeterminarRiesgo:
    Err.Raise Err.Number, "DCOMPersona:DeterminarRiesgo", Err.Description
End Function

Public Function DeterminarSaldoFinalExp(ByVal psCtaCod As String, ByVal pnTipo As Integer, ByVal pdFecha As Date) As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String
On Error GoTo ErrorDeterminarSaldoFinalExp
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    sSql = "stp_sel_DeterminarSaldoFinalExp '" & psCtaCod & "'," & pnTipo & ",'" & Format(pdFecha, "yyyymmdd") & "'"

    Set DeterminarSaldoFinalExp = oCon.CargaRecordSet(sSql)
      
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
    
ErrorDeterminarSaldoFinalExp:
    Err.Raise Err.Number, "DCOMPersona:DeterminarSaldoFinalExp", Err.Description
End Function


Public Function DeterminarSaldoPorPersona(ByVal psPersCod As String, ByVal pnTipo As Integer, ByVal pdFecha As Date) As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String
On Error GoTo ErrorDeterminarSaldoPorPersona
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    sSql = "stp_sel_DeterminarSaldoPorPersona '" & psPersCod & "'," & pnTipo & ",'" & Format(pdFecha, "yyyymmdd") & "'"

    Set DeterminarSaldoPorPersona = oCon.CargaRecordSet(sSql)
      
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
    
ErrorDeterminarSaldoPorPersona:
    Err.Raise Err.Number, "DCOMPersona:DeterminarSaldoPorPersona", Err.Description
End Function
Public Function VinculadosACuenta(ByVal psCtaCod As String) As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String
On Error GoTo ErrorVinculadosACuenta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    sSql = "stp_sel_VinculadosACuenta '" & psCtaCod & "'"

    Set VinculadosACuenta = oCon.CargaRecordSet(sSql)
      
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
    
ErrorVinculadosACuenta:
    Err.Raise Err.Number, "DCOMPersona:VinculadosACuenta", Err.Description
End Function

'WIOR FIN**************************************************************************************
'MIOL 20120917, SEGUN RQ12272 *****************************************************************
Public Function RecuperarClientexCuenta(ByVal pnCuenta As String) As ADODB.Recordset 'OK
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperarClientexCuenta

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_sel_DatosClienteBloqueadoxOrdenPagoxCuenta '" & pnCuenta & "'"
Set RecuperarClientexCuenta = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperarClientexCuenta:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperarClientexCuenta", Err.Description
End Function

'OK
Public Function insClienteBloqueadoxCuenta(ByVal pscMovNro As String, ByVal pscPersCod As String, ByVal pscPersNombre As String, ByVal pscCtaCod As String, ByVal psnEstado As Integer, ByVal psdFecBloqueo As String)
Dim rs As New ADODB.Recordset
Dim oCom As COMConecta.DCOMConecta
Dim lsSQL As String
Set oCom = New COMConecta.DCOMConecta
    oCom.AbreConexion
    lsSQL = "exec stp_ins_ClienteBloqueadoxOrdenPago '" & pscMovNro & "','" & pscPersCod & "','" & pscPersNombre & "','" & pscCtaCod & "','" & psnEstado & "','" & psdFecBloqueo & "'"
    Set rs = oCom.CargaRecordSet(lsSQL)
    Set insClienteBloqueadoxCuenta = rs
    oCom.CierraConexion
Set oCom = Nothing
End Function

'OK
Public Function RecuperaEstadoClienteSobreGiro(ByVal pscPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta

On Error GoTo ErrorRecuperaEstadoClienteSobreGiro

Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
sSql = "exec stp_sel_EstadoClienteSobreGiro '" & pscPersCod & "'"
Set RecuperaEstadoClienteSobreGiro = oConecta.CargaRecordSet(sSql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Function

ErrorRecuperaEstadoClienteSobreGiro:
Set oConecta = Nothing
Err.Raise Err.Number, "DCOMPersona:RecuperaEstadoClienteSobreGiro", Err.Description
End Function

'OK
Public Function updClienteBloqueadoxSobreGiro(ByVal pscPersCod As String, ByVal pscMovDesloqueo As String)
Dim rs As New ADODB.Recordset
Dim oCom As COMConecta.DCOMConecta
Dim lsSQL As String
Set oCom = New COMConecta.DCOMConecta
    oCom.AbreConexion
    lsSQL = "exec stp_update_DesbloqueoPersxSobreGiroOrdenPago '" & pscPersCod & "','" & pscMovDesloqueo & "'"
    Set rs = oCom.CargaRecordSet(lsSQL)
    Set updClienteBloqueadoxSobreGiro = rs
    oCom.CierraConexion
Set oCom = Nothing
End Function

Public Function ObtenerPersonaBloqueadaxSobreGiro_OrdPag(ByVal pcPersCod As String) As Boolean 'OK

    Dim sSql As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As Recordset
    Set rs = New Recordset
    sSql = " exec stp_sel_PersBloqueadaSobreGiro_Verificar'" & pcPersCod & "'"
    
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
        ObtenerPersonaBloqueadaxSobreGiro_OrdPag = True
    Else
        ObtenerPersonaBloqueadaxSobreGiro_OrdPag = False
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'END MIOL *************************************************************************************
'WIOR 20121010 ********************************************************************************
Public Function ObtenerPersCuentaRelac(ByVal psCtaCod As String, ByVal pnRelac As Integer) As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String
On Error GoTo ErrorObtenerPersCuentaRelac
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    sSql = "exec stp_sel_ObtenerTitularesCuenta '" & psCtaCod & "'," & pnRelac

    Set ObtenerPersCuentaRelac = oCon.CargaRecordSet(sSql)
      
    oCon.CierraConexion
    Set oCon = Nothing
    
Exit Function
ErrorObtenerPersCuentaRelac:
    Err.Raise Err.Number, "DCOMPersona:ObtenerPersCuentaRelac", Err.Description
End Function

Public Function ObtenerUltimaVisita(ByVal psPersCod As String) As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String
On Error GoTo ErrorObtenerUltimaVisita
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    sSql = "exec stp_sel_UltimaVisitaAPersona '" & psPersCod & "'"

    Set ObtenerUltimaVisita = oCon.CargaRecordSet(sSql)
      
    oCon.CierraConexion
    Set oCon = Nothing
    
Exit Function
ErrorObtenerUltimaVisita:
    Err.Raise Err.Number, "DCOMPersona:ObtenerUltimaVisita", Err.Description
End Function

'''AMDO20130702 TI-ERS063-2013
Public Function VerificaExisteSolicitudDatos(ByVal psCodPers As String, Optional ByVal pnTipo As Integer = 0) As Boolean
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim sSql As String
On Error GoTo ErrorVerificaExisteSolicitudDatos
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    sSql = "exec dbtarjeta..stp_sel_ERS0632013_VerificaExisteSolicitudDatos '" & psCodPers & "'," & pnTipo & ""
    Set rs = oCon.CargaRecordSet(sSql)
    VerificaExisteSolicitudDatos = rs!bExiste
      
    oCon.CierraConexion
    Set oCon = Nothing
Exit Function
ErrorVerificaExisteSolicitudDatos:
    Err.Raise Err.Number, "DCOMPersona:VerificaExisteSolicitudDatos", Err.Description
End Function

Public Sub ActualizarInfoContacto(ByVal psPersCod As String, ByVal psUserAct As String, ByVal psAgeAct As String, ByVal psFechaAct As String, _
                                    ByVal psPersTelNew As String, ByVal psPersTel2New As String, ByVal psPersCelNew As String, ByVal psPersCel2New As String, _
                                    ByVal psPersCel3New As String)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    
    On Error GoTo ErrorPersRealiza
    
    lsSQL = "exec dbtarjeta..stp_upd_ERS0632013_ActualizarInfoContacto '" & psPersCod & "','" & psUserAct & "','" & psAgeAct & "','" & psFechaAct & "','" & psPersTelNew & "','" & psPersTel2New & "','" & psPersCelNew & "','" & psPersCel2New & "','" & psPersCel3New & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
    
Exit Sub
ErrorPersRealiza:
    Err.Raise Err.Number, "DCOMPersona:InsertaPersonaRealizaOperacion", Err.Description
End Sub

'''AMDO END

Public Sub InsertaPersonaRealizaOperacion(ByVal pnNroMov As Long, ByVal pnCtaCod As String, ByVal pnTipoCliente As Integer, _
                                            ByVal psPersCod As String, ByVal pnPersDOITipo As Integer, ByVal pnPersDOI As String, _
                                            ByVal psPerNombre As String, ByVal pnTipo As Integer, _
                                            Optional ByVal pcOrigen As String = "", Optional ByVal pnCondicion As Integer = 0) 'WIOR 20121114 AGREGO pcOrigen,pnCondicion
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    
    On Error GoTo ErrorPersRealiza
    
    lsSQL = "exec stp_ins_PersRealizaOperacion " & pnNroMov & ",'" & pnCtaCod & "'," & pnTipoCliente & ",'" & psPersCod & "'," & pnPersDOITipo & ",'" & pnPersDOI & "','" & psPerNombre & "'," & pnTipo & ",'" & pcOrigen & "'," & pnCondicion 'WIOR 20121114 AGREGO pcOrigen,pnCondicion
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
    
Exit Sub
ErrorPersRealiza:
    Err.Raise Err.Number, "DCOMPersona:InsertaPersonaRealizaOperacion", Err.Description
End Sub
Public Function ObtenerPersOperaciones(ByVal psPersCod As String) As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sSql As String
On Error GoTo ErrorObtenerPersOperaciones
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    sSql = "exec stp_sel_PersOperaciones '" & psPersCod & "'"

    Set ObtenerPersOperaciones = oCon.CargaRecordSet(sSql)
      
    oCon.CierraConexion
    Set oCon = Nothing
    
Exit Function
ErrorObtenerPersOperaciones:
    Err.Raise Err.Number, "DCOMPersona:ObtenerPersOperaciones", Err.Description
End Function
Public Sub InsertaActualizaPersOperaciones(ByVal pnInsAct As Integer, ByVal psPersCod As String, ByVal pcPeriodo As String, ByVal pnPerido As Integer)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    
    On Error GoTo ErrorInsertaActualizaPersOperaciones
    
    If pnInsAct = 1 Then
        lsSQL = "exec stp_ins_PersOperaciones '"
    ElseIf pnInsAct = 2 Then
        lsSQL = "exec stp_upd_PersOperaciones '"
    End If
    
    
    lsSQL = lsSQL & psPersCod & "','" & pcPeriodo & "'," & pnPerido
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.Ejecutar lsSQL
    oConec.CierraConexion
    
Exit Sub
ErrorInsertaActualizaPersOperaciones:
    Err.Raise Err.Number, "DCOMPersona:InsertaActualizaPersOperaciones", Err.Description
End Sub
Public Sub InsertaPersOperacionesDet(ByVal psPersCod As String, ByVal pnMov As Long, ByVal pnReu As String, _
                                    ByVal pcPeriodo As String, ByVal pnPerido As Integer, ByVal pnMoneda As Integer, _
                                    ByVal pnMonto As Double)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    
    On Error GoTo ErrorInsertaPersOperacionesDet

    lsSQL = "exec stp_ins_PersOperacionesDet '" & psPersCod & "'," & pnMov & "," & pnReu & ",'" & pcPeriodo & "'," & _
            pnPerido & "," & pnMoneda & "," & pnMonto
    
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.Ejecutar lsSQL
    oConec.CierraConexion
    
Exit Sub
ErrorInsertaPersOperacionesDet:
    Err.Raise Err.Number, "DCOMPersona:InsertaPersOperacionesDet", Err.Description
End Sub
Public Function ObtenerPersOperacionesDet(ByVal psPersCod As String, ByVal pnTipoCamb As Double, ByVal pdPeriodo As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    
    On Error GoTo ErrorInsertaPersOperacionesDet

    lsSQL = "exec stp_sel_ObtenerPersAcumuladoLavDineroDet " & pnTipoCamb & ",'" & pdPeriodo & "','" & psPersCod & "'"
    
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerPersOperacionesDet = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    
    
Exit Function
ErrorInsertaPersOperacionesDet:
    Err.Raise Err.Number, "DCOMPersona:InsertaPersOperacionesDet", Err.Description
End Function
'WIOR FIN *************************************************************************************
'WIOR 20130301 ********************************************************************************
Public Sub InsertaPersOperacionGen(ByVal pnMovNro As Long, ByVal psCtaCod As String, ByVal pnUsuPresente As Integer, _
                                   ByVal pnOpeEfectivo As Integer, ByVal pcOrigen As String, ByVal pcOpeCod As String, _
                                   ByVal pnCondicion As Integer)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    
    On Error GoTo ErrorPersRealiza
    
    lsSQL = "exec stp_ins_PersOperacionGen " & pnMovNro & ",'" & psCtaCod & "'," & pnUsuPresente & "," & pnOpeEfectivo & ",'" & pcOrigen & "','" & pcOpeCod & "'," & pnCondicion
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
    
Exit Sub
ErrorPersRealiza:
    Err.Raise Err.Number, "DCOMPersona:InsertaPersOperacionGen", Err.Description
End Sub
Public Sub InsertaPersOperacionGenDet(ByVal pnMovNro As Long, ByVal psPersCod As String, ByVal pnTipoCliente As Integer, _
                                      ByVal pnPersDOITipo As Integer, ByVal pnPersDOI As String, _
                                      ByVal psPerNombre As String, ByVal pnTipoRelac As Integer)
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
    
    On Error GoTo ErrorInsertaPersOperacionGenDet
    
    lsSQL = "exec stp_ins_PersOperacionGenDet " & pnMovNro & ",'" & psPersCod & "'," & pnTipoCliente & "," & pnPersDOITipo & ",'" & pnPersDOI & "','" & psPerNombre & "'," & pnTipoRelac
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    oConec.CargaRecordSet (lsSQL)
    oConec.CierraConexion
    
Exit Sub
ErrorInsertaPersOperacionGenDet:
    Err.Raise Err.Number, "DCOMPersona:InsertaPersOperacionGenDet", Err.Description
End Sub
'WIOR FIN *************************************************************************************
'JUEZ 20130520 *************************************************
Public Function RecuperaPersonaEnvioEstadoCta(ByVal psPersCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
       
    lsSQL = "exec stp_sel_RecuperaDatosPersonaEnvioEstadoCta '" & psPersCod & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set RecuperaPersonaEnvioEstadoCta = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
End Function
Public Function RecuperaPersonaEnvioEstadoCtaDoc(ByVal psDoc As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim oConec As COMConecta.DCOMConecta
       
    lsSQL = "exec stp_sel_RecuperaDatosPersonaEnvioEstadoCtaDoc '" & psDoc & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set RecuperaPersonaEnvioEstadoCtaDoc = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
End Function
Public Sub ActualizaPersDireccion(ByVal psPersCod As String, ByVal psPersDireccDomicilio As String, pnTipo As Integer)
'APRI20180215 ERS036-2017 ADD pnTipo
Dim sSql As String
Dim oConecta As COMConecta.DCOMConecta
    
On Error GoTo ErrorRegistrarOtroBien
    
    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    sSql = "exec stp_upd_ActualizaPersDireccion '" & psPersCod & "','" & psPersDireccDomicilio & "'," & pnTipo
    oConecta.ConexionActiva.Execute sSql
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
    
ErrorRegistrarOtroBien:
    Set oConecta = Nothing
    Err.Raise Err.Number, "DCOMPersona:RegistrarOtroBien", Err.Description
End Sub

'END JUEZ ******************************************************
'***Agregado por ELRO el 20130701, RFC1306270002****
Public Function devolverRolPersona(ByVal nPersTipo As Integer, ByVal psPersCod As String) As Boolean
    Dim lsSQL As String
    Dim rsRol As ADODB.Recordset
    Dim oDCOMConecta As COMConecta.DCOMConecta
           
    Set oDCOMConecta = New COMConecta.DCOMConecta
    Set rsRol = New ADODB.Recordset
    
    lsSQL = "exec stp_sel_VerificarRolPersona " & nPersTipo & ", '" & psPersCod & "'"
    
    oDCOMConecta.AbreConexion
    Set rsRol = oDCOMConecta.CargaRecordSet(lsSQL)
    oDCOMConecta.CierraConexion
    
    If Not (rsRol.BOF And rsRol.EOF) Then
        If rsRol!nExiste > 0 Then
            devolverRolPersona = True
        Else
            devolverRolPersona = False
        End If
    Else
        
        devolverRolPersona = False
    End If
    Set oDCOMConecta = Nothing
    Set rsRol = Nothing
    Exit Function
ErrorDevolverRolPersona:
    Set oDCOMConecta = Nothing
    Err.Raise Err.Number, "DCOMPersona:devolverRolPersona", Err.Description
End Function
'***Fin Agregado por ELRO el 20130701, RFC1306270002
'EJVG20130912 ***
Public Function RecuperaPersonaxCapRegVouDep(ByVal psPersCod As String) As ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    Dim lsSQL As String
    lsSQL = "exec RecuperaPersonaxCapRegVouDep '" & psPersCod & "'"
    oConec.AbreConexion
    Set RecuperaPersonaxCapRegVouDep = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
'END EJVG *******
'EJVG20140201 ***
Public Function RecuperaPersonaxCheque(ByVal psPersCod As String) As ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    Dim lsSQL As String
    lsSQL = "Exec RecuperaPersonaxCheque '" & psPersCod & "'"
    oConec.AbreConexion
    Set RecuperaPersonaxCheque = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
'END EJVG *******
'ORCR20140503 ***
Public Function ObtenerPersonaRHEstado(ByVal psPersCod As String) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorObtenerPersonaRHEstado
    sSql = "exec sp_obtenerPersonaRHEstado '" & psPersCod & "'"
    oConecta.AbreConexion
    Set ObtenerPersonaRHEstado = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerPersonaRHEstado:
    Err.Raise Err.Number, "DCOMPersona:ObtenerPersonaRHEstado"", Err.Description"
End Function
'END ORCR *******
'ALPA 20140415************************************************************************
Public Function ObtenerInformeRiesgoCuadro1(ByVal psCtaCod As String, ByVal pdFecha As Date) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorObtenerInformeRiesgoCuadro1
    sSql = "exec stp_sel_InformeRiesgoCuadro1 '" & psCtaCod & "','" & Format(pdFecha, "YYYY/MM/DD") & "'"
    oConecta.AbreConexion
    Set ObtenerInformeRiesgoCuadro1 = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerInformeRiesgoCuadro1:
    Err.Raise Err.Number, "DCOMPersona:ObtenerInformeRiesgoCuadro1"", Err.Description"
End Function
Public Function ObtenerInformeRiesgoCuadro1Vinculados(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorObtenerInformeRiesgoCuadro1Vinculados
    sSql = "exec stp_sel_InformeRiesgoCuadro1Vinculados '" & psCtaCod & "'"
    oConecta.AbreConexion
    Set ObtenerInformeRiesgoCuadro1Vinculados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerInformeRiesgoCuadro1Vinculados:
    Err.Raise Err.Number, "DCOMPersona:ObtenerInformeRiesgoCuadro1Vinculados"", Err.Description"
End Function
Public Function ObtenerInformeRiesgoCuadro2SaldoDeudor(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorObtenerInformeRiesgoCuadro2SaldoDeudor
    sSql = "exec stp_sel_InformeRiesgoCuadro2SaldoDeudor '" & psCtaCod & "'"
    oConecta.AbreConexion
    Set ObtenerInformeRiesgoCuadro2SaldoDeudor = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerInformeRiesgoCuadro2SaldoDeudor:
    Err.Raise Err.Number, "DCOMPersona:ObtenerInformeRiesgoCuadro2SaldoDeudor"", Err.Description"
End Function
Public Function ObtenerInformeRiesgoCuadro2Creditos(ByVal psCtaCod As String, ByVal pnTipoCambio As Currency) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorObtenerObtenerInformeRiesgoCuadro2Creditos
    sSql = "exec stp_sel_InformeRiesgoCuadro2Creditos '" & psCtaCod & "'," & pnTipoCambio
    oConecta.AbreConexion
    Set ObtenerInformeRiesgoCuadro2Creditos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerObtenerInformeRiesgoCuadro2Creditos:
    Err.Raise Err.Number, "DCOMPersona:ObtenerInformeRiesgoCuadro2Creditos"", Err.Description"
End Function
Public Function ObtenerInformeRiesgoCuadro2CartasFianzas(ByVal psCtaCod As String, ByVal pnTipoCambio As Currency) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorObtenerInformeRiesgoCuadro2CartasFianzas
    sSql = "exec stp_sel_InformeRiesgoCuadro2CartasFianzas '" & psCtaCod & "'," & pnTipoCambio
    oConecta.AbreConexion
    Set ObtenerInformeRiesgoCuadro2CartasFianzas = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerInformeRiesgoCuadro2CartasFianzas:
    Err.Raise Err.Number, "DCOMPersona:ObtenerInformeRiesgoCuadro2CartasFianzas"", Err.Description"
End Function
Public Function ObtenerInformeRiesgoCuadro2Garantias(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorObtenerInformeRiesgoCuadro2Garantias
    sSql = "exec stp_sel_InformeRiesgoCuadro2Garantias '" & psCtaCod & "'"
    oConecta.AbreConexion
    Set ObtenerInformeRiesgoCuadro2Garantias = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerInformeRiesgoCuadro2Garantias:
    Err.Raise Err.Number, "DCOMPersona:ObtenerInformeRiesgoCuadro2Garantias"", Err.Description"
End Function
Public Function ObtenerInformeRiesgoCuadro2EscaCuotasC1(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorObtenerInformeRiesgoCuadro2EscaCuotasC1
    sSql = "exec stp_sel_InformeRiesgoCuadro2EscaCuotasC1 '" & psCtaCod & "'"
    oConecta.AbreConexion
    Set ObtenerInformeRiesgoCuadro2EscaCuotasC1 = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerInformeRiesgoCuadro2EscaCuotasC1:
    Err.Raise Err.Number, "DCOMPersona:ObtenerInformeRiesgoCuadro2EscaCuotasC1"", Err.Description"
End Function
Public Function ObtenerInformeRiesgoCuadro2EscaCuotasC2(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorObtenerInformeRiesgoCuadro2EscaCuotasC2
    sSql = "exec stp_sel_InformeRiesgoCuadro2EscaCuotasC2 '" & psCtaCod & "'"
    oConecta.AbreConexion
    Set ObtenerInformeRiesgoCuadro2EscaCuotasC2 = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerInformeRiesgoCuadro2EscaCuotasC2:
    Err.Raise Err.Number, "DCOMPersona:ObtenerInformeRiesgoCuadro2EscaCuotasC2"", Err.Description"
End Function
Public Function ObtenerInformeRiesgoCuadro2EscaCreditosC1(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorObtenerInformeRiesgoCuadro2EscaCreditosC1
    sSql = "exec stp_sel_InformeRiesgoCuadro2EscaCreditosC1 '" & psCtaCod & "'"
    oConecta.AbreConexion
    Set ObtenerInformeRiesgoCuadro2EscaCreditosC1 = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerInformeRiesgoCuadro2EscaCreditosC1:
    Err.Raise Err.Number, "DCOMPersona:ObtenerInformeRiesgoCuadro2EscaCreditosC1"", Err.Description"
End Function
Public Function ObtenerInformeRiesgoCuadro2EscaCreditosC2(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorObtenerInformeRiesgoCuadro2EscaCreditosC2
    sSql = "exec stp_sel_InformeRiesgoCuadro2EscaCreditosC2 '" & psCtaCod & "'"
    oConecta.AbreConexion
    Set ObtenerInformeRiesgoCuadro2EscaCreditosC2 = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerInformeRiesgoCuadro2EscaCreditosC2:
    Err.Raise Err.Number, "DCOMPersona:ObtenerInformeRiesgoCuadro2EscaCreditosC2"", Err.Description"
End Function
Public Function ObtenerInformeRiesgoCuadro3(ByVal psCtaCod As String, ByVal pnTipoCambio As Currency) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorObtenerInformeRiesgoCuadro3
    sSql = "exec stp_sel_InformeRiesgoCuadro3 '" & psCtaCod & "'," & pnTipoCambio
    oConecta.AbreConexion
    Set ObtenerInformeRiesgoCuadro3 = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerInformeRiesgoCuadro3:
    Err.Raise Err.Number, "DCOMPersona:ObtenerInformeRiesgoCuadro3"", Err.Description"
End Function

Public Function ObtenerInformeRiesgoCuadro4(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorObtenerInformeRiesgoCuadro4
    sSql = "exec stp_sel_InformeRiesgoCuadro4 '" & psCtaCod & "'"
    oConecta.AbreConexion
    Set ObtenerInformeRiesgoCuadro4 = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerInformeRiesgoCuadro4:
    Err.Raise Err.Number, "DCOMPersona:ObtenerInformeRiesgoCuadro4"", Err.Description"
End Function

Public Sub InsertarInformeRiesgoCuadro1(ByVal dfecha As Date, ByVal psCtaCod As String, ByVal psPersCIIU As String, ByVal psAntiguedad As String, _
                                            ByVal psTpoCredCod As String, ByVal psAgeCodAct As String, ByVal psAnalista As String, ByVal psCalGen As String, _
                                            ByVal psCalSistF As String, ByVal psAntiguedadCMACM As String, nNroCredOtorgado As Integer, nTotalDeudaSistF As Currency, nNroEntidadSistF As Integer, _
                                            ByVal psHisCredCMACM As String, ByVal psEvolSistFina As String)
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorInsertarInformeRiesgoCuadro1
    sSql = "exec stp_ins_InformeRiesgoCuadro1 '" & Format(dfecha, "YYYY/MM/DD") & "' , '" & psCtaCod & "', '" & psPersCIIU & "', '" & psAntiguedad & "', '" & psTpoCredCod & "','" & psAgeCodAct & "','" & psAnalista & "', '" & psCalGen & "','" & psCalSistF & "' , '" & psAntiguedadCMACM & "', " & nNroCredOtorgado & " , '" & nTotalDeudaSistF & "', '" & nNroEntidadSistF & "', '" & psHisCredCMACM & "', '" & psEvolSistFina & "'"
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInsertarInformeRiesgoCuadro1:
    Err.Raise Err.Number, "DCOMPersona:InsertarInformeRiesgoCuadro1"", Err.Description"
End Sub

Public Sub InsertarInformeRiesgoCuadro1Vinculados(ByVal dfecha As Date, ByVal psCtaCod As String, _
                                                  ByVal cPersCodVinculado As String, _
                                                  ByVal nPrdPersRelac As Integer, _
                                                  ByVal nMontoEnde As Currency, _
                                                  ByVal nNroIFIS As Integer, _
                                                  ByVal cCalSistF As String, _
                                                  ByVal cEvolucion As String)
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorInsertarInformeRiesgoCuadro1Vinculados
    sSql = "exec stp_ins_InformeRiesgoCuadro1Vinculados '" & Format(dfecha, "YYYY/MM/DD") & "' , '" & psCtaCod & "', '" & cPersCodVinculado & "', " & nPrdPersRelac & "," & nMontoEnde & ", " & nNroIFIS & ",'" & cCalSistF & "','" & cEvolucion & "'"
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInsertarInformeRiesgoCuadro1Vinculados:
    Err.Raise Err.Number, "DCOMPersona:InsertarInformeRiesgoCuadro1Vinculados"", Err.Description"
End Sub

Public Sub InsertarInformeRiesgoCuadro2Creditos(ByVal dfecha As Date, ByVal psCtaCod As String, _
                                                  ByVal nMontoPropuesto As Currency, _
                                                  ByVal nNroCuotas As Integer, _
                                                  ByVal nTEA As Currency, _
                                                  ByVal nTEMA As Currency, _
                                                  ByVal nTEM As Currency, _
                                                  ByVal nExMN As Currency, _
                                                  ByVal nExRU As Currency, _
                                                  ByVal nVGET As Currency, _
                                                  ByVal nCuoPropuesta As Currency, _
                                                  ByVal nMonPropuesta As Currency)
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorInsertarInformeRiesgoCuadro2Creditos
    sSql = "exec stp_ins_InformeRiesgoCuadro2Creditos '" & Format(dfecha, "YYYY/MM/DD") & "' , '" & psCtaCod & "', " & nMontoPropuesto & ", " & nNroCuotas & ", " & nTEA & "," & nTEMA & "," & nTEM & "," & nExMN & "," & nExRU & "," & nVGET & ", " & nCuoPropuesta & "," & nMonPropuesta
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInsertarInformeRiesgoCuadro2Creditos:
    Err.Raise Err.Number, "DCOMPersona:InsertarInformeRiesgoCuadro2Creditos"", Err.Description"
End Sub

Public Sub InsertarInformeRiesgoCuadro2CartasFianzas(ByVal dfecha As Date, ByVal psCtaCod As String, _
                                                  ByVal nMontoPropuesto As Currency, _
                                                  ByVal nTipoCambio As Currency, _
                                                  ByVal nComisionTr As Currency, _
                                                  ByVal nExMN As Currency, _
                                                  ByVal nExRU As Currency, _
                                                  ByVal nVGET As Currency)
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorInsertarInformeRiesgoCuadro2CartasFianzas
    sSql = "exec stp_ins_InformeRiesgoCuadro2CartasFianzas '" & Format(dfecha, "YYYY/MM/DD") & "' , '" & psCtaCod & "', " & nMontoPropuesto & ", " & nTipoCambio & ", " & nComisionTr & "," & nExMN & "," & nExRU & "," & nVGET & ""
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInsertarInformeRiesgoCuadro2CartasFianzas:
    Err.Raise Err.Number, "DCOMPersona:InsertarInformeRiesgoCuadro2CartasFianzas"", Err.Description"
End Sub

Public Sub InformeRiesgoCuadro2SaldoDeudor(ByVal dfecha As Date, ByVal psCtaCod As String, _
                                                  ByVal cPersCodVinculado As String, _
                                                  ByVal nPrdPersRelac As Integer, _
                                                  ByVal nMoneda As Currency, _
                                                  ByVal nMontoEnde As Currency)
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorInformeRiesgoCuadro2SaldoDeudor
    sSql = "exec stp_ins_InformeRiesgoCuadro2SaldoDeudor '" & Format(dfecha, "YYYY/MM/DD") & "' , '" & psCtaCod & "', '" & cPersCodVinculado & "', " & nPrdPersRelac & ", " & nMoneda & "," & nMontoEnde & ""
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInformeRiesgoCuadro2SaldoDeudor:
    Err.Raise Err.Number, "DCOMPersona:InformeRiesgoCuadro2SaldoDeudor"", Err.Description"
End Sub

Public Sub InformeRiesgoCuadro2Garantias(ByVal dfecha As Date, ByVal psCtaCod As String, _
                                                  ByVal cNumGarant As String, _
                                                  ByVal nTipoGarantia As Integer, _
                                                  ByVal nVRM As Currency, _
                                                  ByVal nVGravamen As Currency, _
                                                  ByVal cDescripcion As String, _
                                                  ByVal cCober As String)
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorInformeRiesgoCuadro2Garantias
    sSql = "exec stp_ins_InformeRiesgoCuadro2Garantias '" & Format(dfecha, "YYYY/MM/DD") & "' , '" & psCtaCod & "', '" & cNumGarant & "', " & nTipoGarantia & ", " & nVRM & "," & nVGravamen & ",'" & cDescripcion & "','" & cCober & "'"
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInformeRiesgoCuadro2Garantias:
    Err.Raise Err.Number, "DCOMPersona:InformeRiesgoCuadro2Garantias"", Err.Description"
End Sub

Public Sub InformeRiesgoCuadro2EscaCuotasC1(ByVal dfecha As Date, ByVal psCtaCod As String, _
                                                  ByVal cCuotas As String)
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorInformeRiesgoCuadro2EscaCuotasC1
    sSql = "exec stp_ins_InformeRiesgoCuadro2EscaCuotasC1 '" & Format(dfecha, "YYYY/MM/DD") & "' , '" & psCtaCod & "', '" & cCuotas & "'"
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInformeRiesgoCuadro2EscaCuotasC1:
    Err.Raise Err.Number, "DCOMPersona:InformeRiesgoCuadro2EscaCuotasC1"", Err.Description"
End Sub

Public Sub InformeRiesgoCuadro2EscaCuotasC2(ByVal dfecha As Date, ByVal psCtaCod As String, _
                                                  ByVal nMontoPagado As Currency, _
                                                  ByVal cCuotasPagadas As String)
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorInformeRiesgoCuadro2EscaCuotasC1
    sSql = "exec stp_ins_InformeRiesgoCuadro2EscaCuotasC2 '" & Format(dfecha, "YYYY/MM/DD") & "' , '" & psCtaCod & "',0, '" & nMontoPagado & "','" & cCuotasPagadas & "'"
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInformeRiesgoCuadro2EscaCuotasC1:
    Err.Raise Err.Number, "DCOMPersona:InformeRiesgoCuadro2EscaCuotasC2"", Err.Description"
End Sub

Public Sub InformeRiesgoCuadro2EscaCreditosC1(ByVal dfecha As Date, ByVal psCtaCod As String, _
                                                  ByVal nMontosCred As Currency)
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorInformeRiesgoCuadro2EscaCreditosC1
    sSql = "exec stp_ins_InformeRiesgoCuadro2EscaCreditosC1 '" & Format(dfecha, "YYYY/MM/DD") & "' , '" & psCtaCod & "',0, '" & nMontosCred & "'"
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInformeRiesgoCuadro2EscaCreditosC1:
    Err.Raise Err.Number, "DCOMPersona:InformeRiesgoCuadro2EscaCreditosC1"", Err.Description"
End Sub

Public Sub InformeRiesgoCuadro2EscaCreditosC2(ByVal dfecha As Date, ByVal psCtaCod As String, _
                                                  ByVal nCorr As Integer, _
                                                  ByVal nMontoCredito As Currency, _
                                                  ByVal nMontoPagado As Currency)
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorInformeRiesgoCuadro2EscaCreditosC2
    sSql = "exec stp_ins_InformeRiesgoCuadro2EscaCreditosC2 '" & Format(dfecha, "YYYY/MM/DD") & "' , '" & psCtaCod & "'," & nCorr & ", " & nMontoCredito & "," & nMontoPagado
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInformeRiesgoCuadro2EscaCreditosC2:
    Err.Raise Err.Number, "DCOMPersona:InformeRiesgoCuadro2EscaCreditosC2"", Err.Description"
End Sub

Public Sub InformeRiesgoCuadro3(ByVal dfecha As Date, ByVal psCtaCod As String, _
                                 ByVal nVentas As Currency, ByVal nUtilidades As Currency, ByVal nUnVentas As Currency, _
                                 ByVal nRazonCTE As Currency, ByVal nSaldoDisponible As Currency, ByVal nCapaPago As Currency, _
                                 ByVal nSensibli As Currency, ByVal nCapiTrab As Currency, ByVal nPatrimonio As Currency, _
                                 ByVal nApalancamiento As Currency, nMoraDelSector As Currency, nCapitalSocial As Currency, _
                                 ByVal nPasivo As Currency, ByVal nLineaCred As Currency, ByVal nOtrosIngr As Currency, _
                                 ByVal nPromedio As Currency, ByVal nMesAnterior1 As Currency, ByVal nMesAnterior2 As Currency, _
                                 ByVal nMesAnterior3 As Currency, nVentasAnter1 As Currency, ByVal nUtilidAnter1 As Currency, ByVal nVentasAnter2 As Currency, _
                                 ByVal nUtilidAnter2 As Currency, ByVal dFechaEEFF As Currency, ByVal nVentasActual As Currency, ByVal nUtilidActual As Currency, ByVal nPorcentaje As Currency, _
                                 ByVal cCondicionS As String, ByVal cCalidadEvS As String, ByVal cEvaluacioS As String)
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorInformeRiesgoCuadro3
    sSql = "exec stp_ins_InformeRiesgoCuadro3 '" & Format(dfecha, "YYYY/MM/DD") & "' , '" & psCtaCod & "'," & nVentas & "," & nUtilidades & ", " & nUnVentas & "," & nRazonCTE & "," & nSaldoDisponible
    sSql = sSql & "," & nCapaPago & "," & nSensibli & "," & nCapiTrab & "," & nPatrimonio
    sSql = sSql & "," & nApalancamiento & ", " & nMoraDelSector & "," & nCapitalSocial & "," & nPasivo & "," & nLineaCred & "," & nOtrosIngr & "," & nPromedio & "," & nMesAnterior1 & ", " & nMesAnterior2
    sSql = sSql & "," & nMesAnterior3 & "," & nVentasAnter1 & ", " & nUtilidAnter1 & ", " & nVentasAnter2 & "," & nUtilidAnter2 & ",'" & Format(dFechaEEFF, "YYYY/MM/DD") & "'," & nVentasActual & "," & nUtilidActual & "," & nPorcentaje
    sSql = sSql & ",'" & cCondicionS & "' ,'" & cCalidadEvS & "', '" & cEvaluacioS & "'"
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInformeRiesgoCuadro3:
    Err.Raise Err.Number, "DCOMPersona:InformeRiesgoCuadro3"", Err.Description"
End Sub

Public Function ObtenerInformeRiesgoCuadro3EvaluacionConsumo(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorObtenerInformeRiesgoCuadro3EvaluacionConsumo
    sSql = "exec stp_sel_InformeRiesgoCuadro3EvaluacionConsumo '" & psCtaCod & "'"
    oConecta.AbreConexion
    Set ObtenerInformeRiesgoCuadro3EvaluacionConsumo = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerInformeRiesgoCuadro3EvaluacionConsumo:
    Err.Raise Err.Number, "DCOMPersona:ObtenerInformeRiesgoCuadro3EvaluacionConsumo"", Err.Description"
End Function

Public Sub InsertarInformeRiesgoCuadro3EvaluacionConsumo(ByVal dfecha As Date, ByVal psCtaCod As String, _
                                                  ByVal cTipoEval As String, _
                                                  ByVal cTituloEval As String, _
                                                  ByVal cDescripcion As String, _
                                                  ByVal nPersonal As Currency)
                                                  'ByVal nNegocio As Currency, 'LUCV20160915
                                                  'ByVal nUnico As Currency)   'LUCV20160915
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorInformeRiesgoCuadro2EscaCreditosC2
    'sSql = "exec stp_ins_InformeRiesgoCuadro3EvaluacionConsumo '" & Format(dfecha, "YYYY/MM/DD") & "' , '" & psCtaCod & "','" & cTipoEval & "', '" & cTituloEval & "','" & cDescripcion & "'," & nPersonal & "," & nNegocio & "," & nUnico & "" 'LUCV20160915, Comentó según ERS052-2016
    sSql = "exec stp_ins_InformeRiesgoCuadro3EvaluacionConsumo '" & Format(dfecha, "YYYY/MM/DD") & "' , '" & psCtaCod & "','" & cTipoEval & "', '" & cTituloEval & "','" & cDescripcion & "'," & nPersonal & ""
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInformeRiesgoCuadro2EscaCreditosC2:
    Err.Raise Err.Number, "DCOMPersona:InformeRiesgoCuadro2EscaCreditosC2"", Err.Description"
End Sub

Public Sub InsertarInformeRiesgoCuadro4(ByVal dfecha As Date, ByVal psCtaCod As String, _
                                                  ByVal nNivelRiesgo As Integer, _
                                                  ByVal cConclusione As String, _
                                                  ByVal cRecomendaci As String, _
                                                  ByVal cAnalisisRie As String, _
                                                  ByVal dFechaNR As Date, _
                                                  ByVal cNumeroInforme As String)
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorInsertarInformeRiesgoCuadro4
    sSql = "exec stp_ins_InformeRiesgoCuadro4 '" & Format(dfecha, "YYYY/MM/DD") & "' , '" & psCtaCod & "'," & nNivelRiesgo & ", '" & cConclusione & "','" & cRecomendaci & "','" & cAnalisisRie & "','" & Format(dFechaNR, "YYYY/MM/DD") & "','" & cNumeroInforme & "'"
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInsertarInformeRiesgoCuadro4:
    Err.Raise Err.Number, "DCOMPersona:InsertarInformeRiesgoCuadro4"", Err.Description"
End Sub
'*************************************************************************************

'WIOR 20141016 *************************************************************************
Public Sub InsertarInformeRiesgoCuadro1EmpVinculados(ByVal dfecha As Date, _
                                                    ByVal psCtaCod As String, _
                                                    ByVal cPersCodEmp As String, _
                                                    ByVal cPersCodVinculado As String, _
                                                    ByVal pnGrupoCod As Long, _
                                                    ByVal cRelacion As String, _
                                                    ByVal nMontoEnde As Currency, _
                                                    ByVal nNroIFIS As Integer, _
                                                    ByVal nCalSistF As Integer, _
                                                    ByVal cEvolucion As String)
Dim oConecta As New COMConecta.DCOMConecta
Dim sSql As String
On Error GoTo ErrorInsertarInformeRiesgoCuadro1EmpVinculados

    sSql = "EXEC stp_ins_InformeRiesgoCuadro1EmpVinculados '" & Format(dfecha, "YYYY/MM/DD") & "','" & psCtaCod & "','" & cPersCodEmp & "','" & cPersCodVinculado & "'," & pnGrupoCod & ",'" & cRelacion & "'," & nMontoEnde & ", " & nNroIFIS & "," & nCalSistF & ",'" & cEvolucion & "'"
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInsertarInformeRiesgoCuadro1EmpVinculados:
    Err.Raise Err.Number, "DCOMPersona:InsertarInformeRiesgoCuadro1EmpVinculados"", Err.Description"
End Sub

Public Function ObtenerInformeRiesgoCuadro1EmpVinculados(ByVal psCtaCod As String) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorObtenerInformeRiesgoCuadro1EmpVinculados
    sSql = "exec stp_sel_InformeRiesgoCuadro1EmpVinculados '" & psCtaCod & "'"
    oConecta.AbreConexion
    Set ObtenerInformeRiesgoCuadro1EmpVinculados = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerInformeRiesgoCuadro1EmpVinculados:
    Err.Raise Err.Number, "DCOMPersona:ObtenerInformeRiesgoCuadro1EmpVinculados"", Err.Description"
End Function
'WIOR FIN ***********************************************************************************
'FRHU ERS077-2015 20151130
Public Function ValidarOperacionActAutoDatos(ByVal psCodOpe As String) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorValidarOperacionActAutoDatos
    
    sSql = "EXEC stp_sel_validarCodOperacionParaAutorizacionDeDatos '" & psCodOpe & "'"
    oConecta.AbreConexion
    Set ValidarOperacionActAutoDatos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorValidarOperacionActAutoDatos:
    Err.Raise Err.Number, "DCOMPersona:ValidarOperacionActAutoDatos", Err.Description
End Function
Public Function ClienteEsPersonaNatural(ByVal psPersCod As String) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorClienteEsPersonaNatural
    
    sSql = "EXEC stp_sel_ClienteEsPersonaNatural '" & psPersCod & "'"
    oConecta.AbreConexion
    Set ClienteEsPersonaNatural = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorClienteEsPersonaNatural:
    Err.Raise Err.Number, "DCOMPersona:ClienteEsPersonaNatural", Err.Description
End Function
Public Function ClienteActualizoAutorizoSusDatos(ByVal psPersCod As String) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorClienteActualizoAutorizoSusDatos
    
    sSql = "EXEC stp_sel_ClienteActualizoAutorizoSusDatos '" & psPersCod & "'"
    oConecta.AbreConexion
    Set ClienteActualizoAutorizoSusDatos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorClienteActualizoAutorizoSusDatos:
    Err.Raise Err.Number, "DCOMPersona:ClienteActualizoAutorizoSusDatos", Err.Description
End Function
'ADD PTI1 12/12/2018 ers070-2018
Public Function codOpeBlackList(ByVal psCodOpe As String, ByVal psCodBlo As String) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo codOpeBlackList
    
    sSql = "EXEC stp_sel_blacklist '" & psCodOpe & "', '" & psCodBlo & "'"
    oConecta.AbreConexion
    Set codOpeBlackList = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
codOpeBlackList:
    Err.Raise Err.Number, "DCOMPersona:codOpeBlackList", Err.Description
End Function

Public Function DupDocCred(ByVal CodCue As String) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo DupDocCred
    
    sSql = "EXEC stp_sel_ERS070_2018_DupDocCred '" & CodCue & "'"
    oConecta.AbreConexion
    Set DupDocCred = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
DupDocCred:
    Err.Raise Err.Number, "DCOMPersona:DupDocCred", Err.Description
End Function

Public Function InsertarImpresionCartAut(ByVal CodUsu As String, ByVal snumcuenta As String)
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo DupDocCred
    
    sSql = "EXEC stp_ins_ers070_2018_DupDocImpr '" & CodUsu & "','" & snumcuenta & "'"
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
DupDocCred:
    Err.Raise Err.Number, "DCOMPersona:DupDocCred", Err.Description
End Function
Public Function ObtenerDistritoAgencia(ByVal gsCodAge As String) As String
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
    Dim rsDOI As New ADODB.Recordset
On Error GoTo ErrorObtenerDistritoAgencia
    
    sSql = "EXEC stp_sel_ers070_2018_CiudadAgencia '" & gsCodAge & "'"
    oConecta.AbreConexion
    Set rsDOI = oConecta.CargaRecordSet(sSql)
    If Not rsDOI.EOF Then
        ObtenerDistritoAgencia = rsDOI!cDistrito
    End If
    rsDOI.Close
    Set rsDOI = Nothing
    oConecta.CierraConexion
    Set oConecta = Nothing
  
   
    
    Exit Function
ErrorObtenerDistritoAgencia:
    Err.Raise Err.Number, "DCOMPersona:ObtenerDistritoAgencia", Err.Description
End Function
'FIN PTI1 12/12/2018 ers070-2018
Public Function ObtenerDatosParaActAutDeCliente(ByVal psPersCod As String) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorObtenerDatosParaActAutDeCliente
    
    sSql = "EXEC stp_sel_ObtenerDatosParaActualizacionAutorizacion '" & psPersCod & "'"
    oConecta.AbreConexion
    Set ObtenerDatosParaActAutDeCliente = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerDatosParaActAutDeCliente:
    Err.Raise Err.Number, "DCOMPersona:ObtenerDatosActualizacionAutorizacion", Err.Description
End Function
Public Sub ActualizarPersActAutDatos(ByVal psPersCod As String, ByVal psPersNombre As String, _
                                     ByVal psPersNatSexo As String, ByVal psPersNatEstCiv As String, _
                                     ByVal psNacionalidad As String, ByVal psPersIDTpo As String, _
                                     ByVal psPersIDNro As String, ByVal psPersDireccDomicilio As String, _
                                     ByVal psPersRefDomicilio As String, ByVal psPersDireccUbiGeo As String, _
                                     ByVal psPersCelular As String, ByVal psPersTelefono As String, _
                                     ByVal psEmail As String, ByVal psUltimaActualizacion As String)
Dim oConecta As New COMConecta.DCOMConecta
Dim sSql As String
On Error GoTo ErrorActualizarPersActAutDatos

    sSql = "EXEC stp_upd_PersonaActAutDatos '" & psPersCod & "','" & psPersNombre & "','" & Trim(Right(psPersNatSexo, 2)) & "'," & psPersNatEstCiv & ",'" & psNacionalidad & "'," & _
                                                 psPersIDTpo & ",'" & psPersIDNro & "','" & psPersDireccDomicilio & "','" & psPersRefDomicilio & "','" & psPersDireccUbiGeo & "','" & _
                                                 psPersCelular & "','" & psPersTelefono & "','" & psEmail & "','" & psUltimaActualizacion & "'"
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorActualizarPersActAutDatos:
    Err.Raise Err.Number, "DCOMPersona:ActualizarPersActAutDatos"", Err.Description"
End Sub
Public Sub InsertarPersActAutDatos(ByVal psMovNro As String, ByVal pnModifDatosSensible As Integer, ByVal psPersCod As String, _
                                   ByRef pMatDatos() As TActAutDatos, ByVal pnAutorizaUsoDatos As Integer, ByVal pnTipoRegistro As Integer, _
                                   ByVal pnTipoCarga As Integer, Optional ByVal pnMovVistoElec As Integer = 0, Optional ByVal nMedio As Integer = 0, Optional ByVal nCodOpe As String = 0) 'modificado por pti1 10/12/2018  se greagó nmedio ncodope
Dim oConecta As New COMConecta.DCOMConecta
Dim sSql As String
Dim X As Integer
Dim lFecha As String, lUser As String, lAgeCod As String
On Error GoTo ErrorInsertarPersActAutDatos
    lFecha = Left(psMovNro, 8)
    lUser = Right(psMovNro, 4)
    lAgeCod = Mid(psMovNro, 18, 2)
    'MODIFICADO POR PTI1 ERS070-2018
    sSql = "EXEC stp_ins_ActAutDatos '" & lFecha & "','" & lUser & "','" & lAgeCod & "'," & pnModifDatosSensible & "," & pnMovVistoElec & ",'" & psPersCod & "','" & _
           pMatDatos(1).sNombres & "','" & pMatDatos(1).sApePat & "','" & pMatDatos(1).sApeMat & "','" & pMatDatos(1).sApeCas & "','" & _
           pMatDatos(1).sPersIDTpo & "','" & pMatDatos(1).sPersIDnro & "','" & pMatDatos(1).sSexo & "','" & pMatDatos(1).sEstadoCivil & "','" & _
           pMatDatos(1).cNacionalidad & "','" & pMatDatos(1).sDomicilio & "','" & pMatDatos(1).sRefDomicilio & "','" & pMatDatos(1).sUbicGeografica & "','" & _
           pMatDatos(1).sCelular & "','" & pMatDatos(1).sTelefonos & "','" & pMatDatos(1).sEmail & "','" & _
           pMatDatos(2).sNombres & "','" & pMatDatos(2).sApePat & "','" & pMatDatos(2).sApeMat & "','" & pMatDatos(2).sApeCas & "'," & _
           pMatDatos(2).sPersIDTpo & ",'" & pMatDatos(2).sPersIDnro & "','" & Trim(Right(pMatDatos(2).sSexo, 2)) & "'," & pMatDatos(2).sEstadoCivil & ",'" & _
           pMatDatos(2).cNacionalidad & "','" & pMatDatos(2).sDomicilio & "','" & pMatDatos(2).sRefDomicilio & "','" & pMatDatos(2).sUbicGeografica & "','" & _
           pMatDatos(2).sCelular & "','" & pMatDatos(2).sTelefonos & "','" & pMatDatos(2).sEmail & "'," & pnTipoRegistro & "," & pnTipoCarga & "," & pnAutorizaUsoDatos & "," & _
           nMedio & "," & nCodOpe    'add pti1 10/12/2018
           'pMatDatos(2).sCelular & "','" & pMatDatos(2).sTelefonos & "','" & pMatDatos(2).sEmail & "'," & pnTipoRegistro & "," & pnTipoCarga & "," & pnAutorizaUsoDatos ' comentado por pti1
    
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrorInsertarPersActAutDatos:
    Err.Raise Err.Number, "DCOMPersona:InsertarPersActAutDatos"", Err.Description"
End Sub
'FRHU 20151210 TIC1512100001
Public Function ValidarCargoParaActAutoDatos(ByVal psCodCargoUser As String) As ADODB.Recordset
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
On Error GoTo ErrorValidarCargoParaActAutoDatos
    
    sSql = "EXEC stp_sel_validarCargoParaAutorizacionDeDatos '" & psCodCargoUser & "'"
    oConecta.AbreConexion
    Set ValidarCargoParaActAutoDatos = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorValidarCargoParaActAutoDatos:
    Err.Raise Err.Number, "DCOMPersona:ValidarCargoParaActAutoDatos", Err.Description
End Function
'FIN FRHU 20151210
'FIN FRHU 20151130
'EJVG20150114 ***
Public Function RecuperaDatosPersonaxGarantia(ByVal psPersCod As String) As ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    Dim lsSQL As String
    lsSQL = "Exec stp_sel_RecuperaDatosPersonaxGarantia '" & psPersCod & "'"
    oConec.AbreConexion
    Set RecuperaDatosPersonaxGarantia = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
Public Function EsDOIdeTitular(ByVal psPersCod As String, ByVal psDOI As String) As Boolean
    Dim oConec As New COMConecta.DCOMConecta
    Dim lsSQL As String
    Dim rsDOI As New ADODB.Recordset
    lsSQL = "SELECT bEsDOI = dbo.fnc_EsDOIdeTitular('" & psPersCod & "','" & psDOI & "')"
    oConec.AbreConexion
    Set rsDOI = oConec.CargaRecordSet(lsSQL)
    If Not rsDOI.EOF Then
        EsDOIdeTitular = rsDOI!bEsDOI
    End If
    oConec.CierraConexion
    rsDOI.Close
    Set rsDOI = Nothing
    Set oConec = Nothing
End Function
'END EJVG *******
'EJVG20160509 ***
Public Function ObtenerNivelRiesgo(ByVal psCtaCod As String, ByVal psTipo As String, ByVal pnMontoCol As Double) As String
    Dim oCon As New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sSql As String
    On Error GoTo ErrObtenerNivelRiesgo
    
    oCon.AbreConexion
    
    sSql = "EXEC stp_sel_ERS0282016_ObtenerNivelRiesgo '" & psCtaCod & "','" & psTipo & "'," & pnMontoCol
    Set rs = oCon.CargaRecordSet(sSql)
    
    If Not rs.EOF Then
        ObtenerNivelRiesgo = rs!cNivelRiesgo
    Else
        ObtenerNivelRiesgo = psTipo & "1"
    End If
    rs.Close
    Set rs = Nothing
    
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
ErrObtenerNivelRiesgo:
    Err.Raise Err.Number, "DCOMPersona:ObtenerNivelRiesgo", Err.Description
End Function
'Public Sub InsertarInformeProductoPersona(ByVal psCtaCod As String)
'    Dim oConecta As New COMConecta.DCOMConecta
'    Dim sSql As String
'    On Error GoTo ErrInsertarInformeProductoPersona
'    sSql = "exec stp_ins_ERS0282016_InformeRiesgoProductoPersona '" & psCtaCod & "'"
'    oConecta.AbreConexion
'    oConecta.Ejecutar (sSql)
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Sub
'ErrInsertarInformeProductoPersona:
'    Err.Raise Err.Number, "DCOMPersona:InsertarInformeProductoPersona"", Err.Description"
'End Sub
Public Sub InsertarInformePersVinculados(ByVal psCtaCod As String)
    Dim oConecta As New COMConecta.DCOMConecta
    Dim sSql As String
    On Error GoTo ErrInsertarInformePersVinculados
    sSql = "EXEC stp_ins_ERS0282016_InformeRiesgoPersVinculados '" & psCtaCod & "'"
    oConecta.AbreConexion
    oConecta.Ejecutar (sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Sub
ErrInsertarInformePersVinculados:
    Err.Raise Err.Number, "DCOMPersona:InsertarInformePersVinculados"", Err.Description"
End Sub
'END EJVG *******
'APRI20190101 ERS077-2018
Public Function RecuperaDatosPersonaComunicacion(ByVal psCtaCod As String, ByVal psPersCod As String) As ADODB.Recordset
    Dim oConec As New COMConecta.DCOMConecta
    Dim lsSQL As String
    lsSQL = "Exec stp_sel_ObteneDatosComunicacion '" & psCtaCod & "','" & psPersCod & "'"
    oConec.AbreConexion
    Set RecuperaDatosPersonaComunicacion = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
'END APRI
