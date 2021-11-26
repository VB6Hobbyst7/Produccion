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
Public sCodAge As String

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
    
    
    lnNegritaON = Chr$(27) & Chr$(69)
    lnNegritaOFF = Chr$(27) + Chr$(70)

    sCadImp = ""
    sCadImp = sCadImp & ImpreFormat(psNomCmac, 35, 35) & " - " & psNomAge
    sCadImp = sCadImp & Space(25) & Format(FechaHora(pdFecha), "dd/mm/yyyy hh:mm:ss") & Chr$(10) & Chr$(10) & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "Cliente : " & sPersCod & Space(2) & PstaNombre(sNombreCompleto) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    'sCadImp = sCadImp & "Fte Ingreso : " & FtesIngreso(pnIndex).sRazSocDescrip & Chr$(10) & Chr$(10) & Chr$(10)
    
    sCadImp = sCadImp & "Fte Ingreso : " & FtesIngreso(pnIndex).sRazSocDescrip & Chr$(10) & Chr$(10)
    
'----------------------------------------------------------------------------------------------------
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ImpreFormat("RAZON SOCIAL/NOMBRE COMERCIAL", 35, 0) & ImpreFormat("RUC", 14, 2) & ImpreFormat("INICIO ACT.", 12, 0) & ImpreFormat("PERS. JURIDICA", 24, 2) & ImpreFormat("TELEFONO", 10, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).sRazSocDescrip, 35, 0) & ImpreFormat(FtesIngreso(pnIndex).cFtePersRuc, 14, 2) & ImpreFormat(Format(FtesIngreso(pnIndex).dFtePersFecInicioAct, "dd/mm/yyyy"), 12, 0) & ImpreFormat(FtesIngreso(pnIndex).cFtePersTipoPersJur, 24, 2) & ImpreFormat(FtesIngreso(pnIndex).cFtePersTelefono, 10, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ImpreFormat("ACTIVIDAD(CIUU)", 35, 0) & ImpreFormat("COND. LOCAL", 14, 2) & ImpreFormat("MAGNITUD EMP.", 35, 0) & ImpreFormat("Nª EMPLEADOS", 14, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).cFtePersCIUU, 35, 0) & ImpreFormat(FtesIngreso(pnIndex).cFtePersCondicionDomic, 14, 2) & ImpreFormat(FtesIngreso(pnIndex).cFtePersMagnitudEmp, 35, 0) & ImpreFormat(Round(FtesIngreso(pnIndex).nFtePersNroEmpleados), 14, 0) & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ImpreFormat("DIRECCION(Avenida, Calle, Jiron)", 35, 0) & ImpreFormat("DEPARTAMENTO", 12, 2) & ImpreFormat("PROVINCIA", 12, 2) & ImpreFormat("DISTRITO", 16, 2) & ImpreFormat("ZONA", 17, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).cFtePersDireccion, 35, 0) & ImpreFormat(FtesIngreso(pnIndex).cFtePersDpto, 12, 2) & ImpreFormat(FtesIngreso(pnIndex).cFtePersProv, 12, 2) & ImpreFormat(FtesIngreso(pnIndex).cFtePersDist, 16, 2) & ImpreFormat(FtesIngreso(pnIndex).cFtePersZona, 17, 2) & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10) & Chr$(10)
'----------------------------------------------------------------------------------------------------
    
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & Space(20) & "EVOLUCION ECONOMICA FINANCIERA DEPENDIENTE" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    sCadImp = sCadImp & Space(23)
    
    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
        sCadImp = sCadImp & Format(FtesIngreso(pnIndex).FtesIngresoDep(i).dFecEval, "dd/mm/yyyy") & Space(3)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            sCadImp = sCadImp & "Variacion"
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("Ingreso Familiar", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngFam, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesDep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngFam = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 5, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(CDbl(Format(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngFam / FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngFam, "#0.00")), 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    
    sCadImp = sCadImp & ImpreFormat("Ingreso Conyugue", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngCony, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesDep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngCony = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 5, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(CDbl(Format(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngCony / FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngCony, "#0.00")), 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    
    sCadImp = sCadImp & ImpreFormat("Otros Ingresos", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngOtros, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesDep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngOtros = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 5, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(CDbl(Format(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngOtros / FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngOtros, "#0.00")), 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    
    sCadImp = sCadImp & ImpreFormat("Gastos Familiares", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(i).nGastosFam, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesDep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nGastosFam = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 5, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(CDbl(Format(FtesIngreso(pnIndex).FtesIngresoDep(i).nGastosFam / FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nGastosFam, "#0.00")), 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ImpreFormat("Saldo", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngCony + _
                                FtesIngreso(pnIndex).FtesIngresoDep(i).nIngFam + _
                                FtesIngreso(pnIndex).FtesIngresoDep(i).nIngOtros - FtesIngreso(pnIndex).FtesIngresoDep(i).nGastosFam, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesDep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngCony + _
                                FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngFam + _
                                FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngOtros - FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nGastosFam = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 5, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(CDbl(Format((FtesIngreso(pnIndex).FtesIngresoDep(i).nIngCony + _
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
            sCadImp = sCadImp & ImpreFormat("Evaluacion", 16)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).dFecEval, 14, 0)
            Next i
            sCadImp = sCadImp & Chr$(10)
                
            sCadImp = sCadImp & ImpreFormat("Tipo Cultivo", 20)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).sTpoCultivo, 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ImpreFormat("DESTINO", 20)
            Dim sDestino As String
            
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                If FtesIngreso(pnIndex).CostoProd(i).nSiembra = 1 Then
        '            If Len(sDestino) = 0 Then
        '                sDestino = "SIEMBRA"
        '            Else
        '                sDestino = sDestino & ",SIEMBRA"
                        sCadImp = sCadImp & ImpreFormat("SIEMBRA", 13)
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
                        sCadImp = sCadImp & Space(22) & ImpreFormat("MANTENIMIENTO", 13)
                        sCadImp = sCadImp & Chr$(10)
                End If
                If FtesIngreso(pnIndex).CostoProd(i).nDesAgricola = 1 Then
        '             If Len(sDestino) = 0 Then
        '                sDestino = "DESAGRICOLA"
        '             Else
        '                sDestino = sDestino & ",DESAGRICOLA"
        '                sCadImp = sCadImp & Chr$(10)
        '             End If
                        sCadImp = sCadImp & Space(22) & ImpreFormat("DESAGRICOLA", 13)
                        sCadImp = sCadImp & Chr$(10)
        
        
                End If
                If FtesIngreso(pnIndex).CostoProd(i).nCosecha = 1 Then
        '            If Len(sDestino) = 0 Then
        '                sDestino = "COSECHA"
        '            Else
        '                sDestino = sDestino & ",COSECHA"
        '                sCadImp = sCadImp & Chr$(10)
        '            End If
                        sCadImp = sCadImp & Space(22) & ImpreFormat("COSECHA", 13)
                        sCadImp = sCadImp & Chr$(10)
        
                 End If
                If FtesIngreso(pnIndex).CostoProd(i).nPlanOtros = 1 Then
        '            If Len(sDestino) = 0 Then
        '                sDestino = "OTROS"
        '            Else
        '                sDestino = sDestino & ",OTROS"
        '            End If
                        sCadImp = sCadImp & Space(22) & ImpreFormat("OTROS", 13)
                        'sCadImp = sCadImp & Chr$(10)
        
                End If
                'sCadImp = sCadImp & ImpreFormat(sDestino, 60)
            Next i
            
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ImpreFormat("Maquinaria", 12)
            
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nMaquinaria, 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ImpreFormat("Jornales", 12)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nJornales, 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ImpreFormat("Insumos", 12)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nInsumos, 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ImpreFormat("Pesticidas", 12)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPesticidas, 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ImpreFormat("Otros", 12)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nOtros, 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ImpreFormat("Total", 12)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros, 13)
            Next i
            sCadImp = sCadImp & Chr$(10) & Chr$(10)
            
            sCadImp = sCadImp & ImpreFormat("Hectareas", 12)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas, 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ImpreFormat("Produccion", 18)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ImpreFormat(Format(FtesIngreso(pnIndex).CostoProd(i).nProduccion, "#0.00") & " " & FtesIngreso(pnIndex).CostoProd(i).sUniProd, 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ImpreFormat("Precio Uni.", 12)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPreUni, 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ImpreFormat("Ingresos", 12)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas * FtesIngreso(pnIndex).CostoProd(i).nPreUni * FtesIngreso(pnIndex).CostoProd(i).nProduccion, 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ImpreFormat("Egresos", 12)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas * (FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros), 13)
            Next i
            sCadImp = sCadImp & Chr$(10)
            
            sCadImp = sCadImp & ImpreFormat("Utilidad", 12)
            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
                sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).CostoProd(i).nHectareas * FtesIngreso(pnIndex).CostoProd(i).nPreUni * FtesIngreso(pnIndex).CostoProd(i).nProduccion) - (FtesIngreso(pnIndex).CostoProd(i).nHectareas * (FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros)), 13)
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
    sCadImp = sCadImp & ImpreFormat("REFERENCIAS BANCARIAS", 42, 0) & ImpreFormat("NRO.TARJETA", 22, 2) & ImpreFormat("NRO.CUENTA", 22, 2) & ImpreFormat("MONTO", 10, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF

    'MatDatosFte = RecuperaRefBancariaPersFte(FtesIngreso(pnIndex).sFteIngreso)
    MatDatosFte = RecuperaRefBancariaPersFte(sPersCod)
        
    For i = 0 To UBound(MatDatosFte) - 1
        sCadImp = sCadImp & ImpreFormat(MatDatosFte(i, 0) & "-" & MatDatosFte(i, 1), 40) & ImpreFormat(MatDatosFte(i, 3), 22, 2) & ImpreFormat(MatDatosFte(i, 2), 22, 2) & ImpreFormat(MatDatosFte(i, 4), 10, 2) & Chr$(10)
    Next i
    
    sCadImp = sCadImp & Chr$(10)
    
   ' Muestra los datos de las referencias comerciales de la fte como persona ----------------------------------------
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ImpreFormat("REFERENCIAS COMERCIALES", 66, 0) & ImpreFormat("TIPO RELAC. COMERCIAL", 22, 2) & ImpreFormat("TELEFONO", 10, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF

    'MatDatosComFte = RecuperaRefComercialPersFte(FtesIngreso(pnIndex).sFteIngreso)
    MatDatosComFte = RecuperaRefComercialPersFte(sPersCod)
    For i = 0 To UBound(MatDatosComFte) - 1
        sCadImp = sCadImp & ImpreFormat(MatDatosComFte(i, 3) & "-" & MatDatosComFte(i, 0), 64) & ImpreFormat(MatDatosComFte(i, 2), 22, 2) & ImpreFormat(MatDatosComFte(i, 1), 10, 2) & Chr$(10)
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
Dim oCred As DCreditos
Dim MatDatos As Variant
Dim i As Integer
Dim lnNegritaON As String
Dim lnNegritaOFF As String

Dim MatDatosFte As Variant
Dim MatDatosComFte As Variant
Dim bProd As Boolean

    
    
    lnNegritaON = Chr$(27) & Chr$(69)
    lnNegritaOFF = Chr$(27) + Chr$(70)
    

    Set oCred = New DCreditos
    MatDatos = oCred.RecuperaHistoriaCredPersona(sPersCod)
    Set oCred = Nothing
    
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
    
    'sCadImp = sCadImp & ImpreFormat(psNomCmac, 35, 35) & psNomAge
    sCadImp = sCadImp & ImpreFormat(psNomCmac, 35, 0) & psNomAge
    sCadImp = sCadImp & Space(25) & Format(FechaHora(pdFecha), "dd/mm/yyy hh:mm:ss") & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "Cliente : " & sPersCod & Space(2) & PstaNombre(sNombreCompleto) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & "Fte Ingreso : " & FtesIngreso(pnIndex).sRazSocDescrip & Chr$(10)

    
'----------------------------------------------------------------------------------------------------
    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ImpreFormat("RAZON SOCIAL/NOMBRE COMERCIAL", 35, 0) & ImpreFormat("RUC", 14, 2) & ImpreFormat("INICIO ACT.", 12, 0) & ImpreFormat("PERS. JURIDICA", 24, 2) & ImpreFormat("TELEFONO", 10, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).sRazSocDescrip, 35, 0) & ImpreFormat(FtesIngreso(pnIndex).cFtePersRuc, 14, 2) & ImpreFormat(Format(FtesIngreso(pnIndex).dFtePersFecInicioAct, "dd/mm/yyyy"), 12, 0) & ImpreFormat(FtesIngreso(pnIndex).cFtePersTipoPersJur, 24, 2) & ImpreFormat(FtesIngreso(pnIndex).cFtePersTelefono, 10, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ImpreFormat("ACTIVIDAD(CIUU)", 35, 0) & ImpreFormat("COND. LOCAL", 14, 2) & ImpreFormat("MAGNITUD EMP.", 35, 0) & ImpreFormat("Nª EMPLEADOS", 14, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).cFtePersCIUU, 35, 0) & ImpreFormat(FtesIngreso(pnIndex).cFtePersCondicionDomic, 14, 2) & ImpreFormat(FtesIngreso(pnIndex).cFtePersMagnitudEmp, 35, 0) & ImpreFormat(Round(FtesIngreso(pnIndex).nFtePersNroEmpleados), 14, 0) & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ImpreFormat("DIRECCION(Avenida, Calle, Jiron)", 35, 0) & ImpreFormat("DEPARTAMENTO", 12, 2) & ImpreFormat("PROVINCIA", 12, 2) & ImpreFormat("DISTRITO", 16, 2) & ImpreFormat("ZONA", 17, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).cFtePersDireccion, 35, 0) & ImpreFormat(FtesIngreso(pnIndex).cFtePersDpto, 12, 2) & ImpreFormat(FtesIngreso(pnIndex).cFtePersProv, 12, 2) & ImpreFormat(FtesIngreso(pnIndex).cFtePersDist, 16, 2) & ImpreFormat(FtesIngreso(pnIndex).cFtePersZona, 17, 2) & Chr$(10)
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
    sCadImp = sCadImp & ImpreFormat("ACTIVO CIRCULANTE", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios) = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios), 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ImpreFormat("Disponible", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("Cuentas x Cobrar", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("Inventario", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ImpreFormat("ACTIVO FIJO", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("TOTAL ACTIVO", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo), 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("PASIVO", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact = 0 Then
                        
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact), 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ImpreFormat("Proveedores", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("Otros Prestamos", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("Prestamos CMACT", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ImpreFormat("PATRIMONIO", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio, 8, 2)
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("INGRESOS UE", 20)
    Dim nIng As Currency
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob, 8, 2)
'
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob, 8, 2)
        nIng = FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
                        
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob = 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob = 0 Then
        
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
'                sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob), 8, 2)
               sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob), 8, 2)

            End If
        End If
    Next i
    
    'CMACICA_CSTS - 10/11/2003 - SE CAMBIO EL ORDEN (VENTAS, RECUP. X COB., INGRESOS FAMILIARES) -----------------
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF

    sCadImp = sCadImp & ImpreFormat("Ventas", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas, 8, 2)
        If i > 0 And FtesIngreso(pnIndex).nNumFtesIndep - 1 = i Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas, 8, 2)
            End If
        End If
    Next i

    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("Recup. Ctas x Cob", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob, 8, 2)
            End If
        End If
    Next i

    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Ingresos familiares", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam, 8, 2)
'            End If
'        End If
'    Next i

'-----------------------------------------------------------------------------------------------------------------------
    Dim nEg As Currency
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ImpreFormat("EGRESOS UE", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 8, 2)
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 8, 2)
         nEg = FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos = 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
            
'                sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos), 8, 2)
                sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos), 8, 2)
            End If
        End If
    Next i
    
    'CMACICA_CSTS - 21/11/2003 - SE CAMBIO EL ORDEN (COMPRA DE MERCADERIAS, OTROS EGRESOS, GASTOS FAMILIARES) ---------
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    
    'sCadImp = sCadImp & ImpreFormat("Compra Mercaderia", 20)
    sCadImp = sCadImp & ImpreFormat("Costo de Ventas", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad, 8, 2)
            End If
        End If
    Next i
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("Otros Egresos", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos, 8, 2)
            End If
        End If
    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Gastos Familiares", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam, 8, 2)
'        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam, 8, 2)
'            End If
'        End If
'    Next i

' ---------------------------------------------------------------------------------------------------------------------

    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ImpreFormat("UTILIDAD NETA", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob - _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam - _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad - _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 8, 2)
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
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
            
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
                
            Else
'                sCadImp = sCadImp & ImpreFormat(((FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + _
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
                sCadImp = sCadImp & ImpreFormat(((FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
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
    sCadImp = sCadImp & ImpreFormat("Ingresos familiares", 20)
    Dim nIngFam As Currency
    
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam, 8, 2)
        nIngFam = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam
        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 And i > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("Gastos Familiares", 20)
    Dim nEgrFam As Currency
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam, 8, 2)
        nEgrFam = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("DISPONIBLE", 20)
    sCadImp = sCadImp & ImpreFormat(nIngFam + nIng - (nEgrFam + nEg), 8, 2)
    
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
    sCadImp = sCadImp & ImpreFormat("Liquidez", 20)
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
        sCadImp = sCadImp & ImpreFormat(Res1, 8, 2)
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
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next i
    
    sCadImp = sCadImp & Chr$(10)
    'Activo Circulante - Pasivo
    sCadImp = sCadImp & ImpreFormat("Capital Trabajo", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nActCir = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
        Res1 = nActCir - nPasivo
        sCadImp = sCadImp & ImpreFormat(Res1, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nActCirAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
            nPasivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
            Res2 = nActCirAnt - nPasivoAnt
            If Res2 = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "GESTION" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    'Costo de Ventas / Inventario
    sCadImp = sCadImp & ImpreFormat("Rotacion Inventario", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios = 0 Then
            Res1 = 0
        Else
            Res1 = FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
        End If
        sCadImp = sCadImp & ImpreFormat(Res1, 8, 2)
        
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios = 0 Then
                Res2 = 0
            Else
                Res2 = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
            End If
            If Res2 = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    
    'Recup Ctas x Cob / (Ventas / 30)
    sCadImp = sCadImp & ImpreFormat("Periodo de Cobro", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
            Res1 = 0
        Else
            Res1 = FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob / (FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas / 30)
        End If
        sCadImp = sCadImp & ImpreFormat(Res1, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
                Res2 = 0
            Else
                Res2 = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas / 30)
            End If
            If Res2 = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 8, 2)
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
    sCadImp = sCadImp & ImpreFormat("Rentab. Patrimonial", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
        nEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos
        
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio = 0 Then
            Res1 = 0
        Else
            Res1 = (nIngresos - nEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio
        End If
        sCadImp = sCadImp & ImpreFormat(Res1, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nIngresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob
            nEgresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio = 0 Then
                Res2 = 0
            Else
                Res2 = (nIngresos - nEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio
            End If
        
            If Res2 = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    
    Dim nTotActivo As Double
    Dim nTotActivoAnt As Double
    sCadImp = sCadImp & ImpreFormat("Rentab. Activo", 20)
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
        sCadImp = sCadImp & ImpreFormat(Res1, 8, 2)
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
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    
    sCadImp = sCadImp & ImpreFormat("Rentab. Ventas", 20)
    'Utilidad Operativa / Ventas
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
        nEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos
        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
            Res1 = 0
        Else
            Res1 = (nIngresos - nEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas
        End If
        sCadImp = sCadImp & ImpreFormat(Res1, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nIngresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob
            nEgresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos
            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
                Res2 = 0
            Else
                Res2 = (nIngresosAnt - nEgresosAnt) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas
            End If
            If Res2 = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    
    'Ventas Totales = Ventas + Recup Ctas x Cob
    'Margen Utilidad = Ventas Totales - Costo Ventas / Ventas Totales
    Dim nVentasTot As Double
    Dim nVentasTotAnt As Double
    sCadImp = sCadImp & ImpreFormat("Margen Utilidad", 20)
    For i = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nVentasTot = FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
        If nVentasTot = 0 Then
            Res1 = 0
        Else
            Res1 = (nVentasTot - FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad) / nVentasTot
        End If
        sCadImp = sCadImp & ImpreFormat(Res1, 8, 2)
        If i > 0 And i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nVentasTotAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
            If nVentasTotAnt = 0 Then
                Res2 = 0
            Else
                Res2 = (nVentasTotAnt - FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad) / nVentasTotAnt
            End If
            If Res2 = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "ENDEUDAMIENTO" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ImpreFormat("Endeudamiento Pat.", 20)
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
        sCadImp = sCadImp & ImpreFormat(Res1, 8, 2)
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
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next i
    sCadImp = sCadImp & Chr$(10)
    
    sCadImp = sCadImp & ImpreFormat("Endeudamiento", 20)
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
        sCadImp = sCadImp & ImpreFormat(Res1, 8, 2)
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
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 8, 2)
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
        sCadImp = sCadImp & ImpreFormat("Evaluacion", 16)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).dFecEval, 14, 0)
        Next i
        sCadImp = sCadImp & Chr$(10)
            
        sCadImp = sCadImp & ImpreFormat("Tipo Cultivo", 20)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).sTpoCultivo, 13)
        Next i
        
        sCadImp = sCadImp & Chr$(10)
        
        sCadImp = sCadImp & ImpreFormat("Destino", 20)
        
        Dim sDestino As String
        bProd = False
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            If FtesIngreso(pnIndex).CostoProd(i).nSiembra = 1 Then
    '            If Len(sDestino) = 0 Then
    '                sDestino = "SIEMBRA"
    '            Else
    '                sDestino = sDestino & ",SIEMBRA"
                    bProd = True
                    sCadImp = sCadImp & ImpreFormat("SIEMBRA", 13)
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
                    sCadImp = sCadImp & Space(7) & ImpreFormat("MANTENIMIENTO", 13)
                Else
                    sCadImp = sCadImp & ImpreFormat("MANTENIMIENTO", 13)
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
                    sCadImp = sCadImp & Space(7) & ImpreFormat("DESAGRICOLA", 13)
                  Else
                    sCadImp = sCadImp & ImpreFormat("DESAGRICOLA", 13)
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
                    sCadImp = sCadImp & Space(7) & ImpreFormat("COSECHA", 13)
                Else
                    sCadImp = sCadImp & ImpreFormat("COSECHA", 13)
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
                        sCadImp = sCadImp & Space(7) & ImpreFormat("OTROS", 13)
                    Else
                        sCadImp = sCadImp & ImpreFormat("OTROS", 13)
                        bProd = True
                    End If
                    'sCadImp = sCadImp & Chr$(10)
    
            End If
            'sCadImp = sCadImp & ImpreFormat(sDestino, 60)
        Next i
        
        sCadImp = sCadImp & Chr$(10)
        
        sCadImp = sCadImp & ImpreFormat("Maquinaria", 12)
        
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nMaquinaria, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)
        
        sCadImp = sCadImp & ImpreFormat("Jornales", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nJornales, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)
        
        sCadImp = sCadImp & ImpreFormat("Insumos", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nInsumos, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)
        
        sCadImp = sCadImp & ImpreFormat("Pesticidas", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPesticidas, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)
        
        sCadImp = sCadImp & ImpreFormat("Otros", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nOtros, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)
        
        sCadImp = sCadImp & ImpreFormat("Total", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros, 13)
        Next i
        sCadImp = sCadImp & Chr$(10) & Chr$(10)
        
        sCadImp = sCadImp & ImpreFormat("Hectareas", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)
        
        sCadImp = sCadImp & ImpreFormat("Produccion", 18)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ImpreFormat(Format(FtesIngreso(pnIndex).CostoProd(i).nProduccion, "#0.00") & " " & FtesIngreso(pnIndex).CostoProd(i).sUniProd, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)
        
        sCadImp = sCadImp & ImpreFormat("Precio Uni.", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPreUni, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)
        
        sCadImp = sCadImp & ImpreFormat("Ingresos", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas * FtesIngreso(pnIndex).CostoProd(i).nPreUni * FtesIngreso(pnIndex).CostoProd(i).nProduccion, 13)
        Next i
        sCadImp = sCadImp & Chr$(10)
        
        sCadImp = sCadImp & ImpreFormat("Egresos", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas * (FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros), 13)
        Next i
        sCadImp = sCadImp & Chr$(10)
        
        sCadImp = sCadImp & ImpreFormat("Utilidad", 12)
        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
            sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).CostoProd(i).nHectareas * FtesIngreso(pnIndex).CostoProd(i).nPreUni * FtesIngreso(pnIndex).CostoProd(i).nProduccion) - (FtesIngreso(pnIndex).CostoProd(i).nHectareas * (FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros)), 13)
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
    sCadImp = sCadImp & ImpreFormat("REFERENCIAS BANCARIAS", 42, 0) & ImpreFormat("NRO.TARJETA", 22, 2) & ImpreFormat("NRO.CUENTA", 22, 2) & ImpreFormat("MONTO", 10, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF

    MatDatosFte = RecuperaRefBancariaPersFte(FtesIngreso(pnIndex).sFteIngreso)
        
    For i = 0 To UBound(MatDatosFte) - 1
        sCadImp = sCadImp & ImpreFormat(MatDatosFte(i, 0) & "-" & MatDatosFte(i, 1), 40) & ImpreFormat(MatDatosFte(i, 3), 22, 2) & ImpreFormat(MatDatosFte(i, 2), 22, 2) & ImpreFormat(MatDatosFte(i, 4), 10, 2) & Chr$(10)
    Next i
   
    sCadImp = sCadImp & Chr$(10)
    
   ' Muestra los datos de las referencias comerciales de la fte como persona ----------------------------------------
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ImpreFormat("REFERENCIAS COMERCIALES", 66, 0) & ImpreFormat("TIPO RELAC. COMERCIAL", 22, 2) & ImpreFormat("TELEFONO", 10, 2) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF

    MatDatosComFte = RecuperaRefComercialPersFte(FtesIngreso(pnIndex).sFteIngreso)
        
    For i = 0 To UBound(MatDatosComFte) - 1
        sCadImp = sCadImp & ImpreFormat(MatDatosComFte(i, 3) & "-" & MatDatosComFte(i, 0), 64) & ImpreFormat(MatDatosComFte(i, 2), 22, 2) & ImpreFormat(MatDatosComFte(i, 1), 10, 2) & Chr$(10)
    Next i
    
    ' ---------------------------------------------------------------------------------------------------------------
    sCadImp = sCadImp & Chr$(10)
    
    GenerarImpresionFteIngresoIndependiente = sCadImp
    
End Function


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
Dim sSQL As String
Dim oConecta As DConecta
Dim R As ADODB.Recordset

On Error GoTo ErrorTiempoFteIngreso
    sSQL = "Select nParamValor from " & gConsComunes & "ColocParametro where nParamVar = " & gColocParamFteIngresoTiempoEval
    Set oConecta = New DConecta
    Call oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
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
Dim sSQL As String
Dim R As ADODB.Recordset
Dim oConecta  As DConecta
    On Error GoTo ErrorUltimoCorrelativo
    sSQL = "Select MAX(substring(P.cPersCod,6,7)) as Correlativo from " & gConsPersona & "Persona P where substring(P.cPersCod,4,2) = '" & Right(psCodAge, 2) & "'"
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    If Not R.BOF And Not R.EOF Then
        UltimoCorrelativo = IIf(IsNull(R!Correlativo), 0, R!Correlativo)
    Else
        UltimoCorrelativo = 0
    End If
    R.Close
    
    
    If bPersDistrib Then
        sSQL = "Select MAX(substring(P.cCodPers,3,7)) as Correlativo from Persona P where substring(P.cCodPers,1,2) = '" & Right(psCodAge, 2) & "'"
        R.CursorLocation = adUseClient
        R.Open sSQL, ConexDitrib, adOpenStatic, adLockReadOnly, adCmdText
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
Dim sSQL As String
Dim InicioTran As Boolean
Dim i As Integer
Dim J As Integer
Dim oConecta As DConecta
Dim oNegPers As NPersona
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
            Set oConecta = New DConecta
            oConecta.AbreConexion
            Set ConexionActiva = oConecta.ConexionActiva
        End If
    End If
    
    '**********************   Para Actualizar el Distribuido ************************
    sSQL = "Select nConsSisValor from ConstSistema where nConsSisCod=45"
    Set R = New ADODB.Recordset
    R.Open sSQL, ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
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
    sSQL = "Select MAX(cNumFuente) as cNumFuente From PersFteIngreso"
    Set R = New ADODB.Recordset
    R.Open sSQL, ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    sNumFuente = Format(CLng(IIf(IsNull(R!cNumFuente), 0, R!cNumFuente)) + 1, "00000000")
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
        sPersCod = Trim(psCodCmact) & Trim(psCodAge) & Trim(lsCadFinal)
        Set oNegPers = Nothing
    Else
        sCodPersDistrib = Right(sPersCod, 10)
    End If
    '********************************************************************************
    Set R = Nothing
    
    sNombreCompleto = Replace(sNombreCompleto, "'", "''")
    
    If nTipoAct = PersFilaModificada Then
                
        'Actualiza Persona
        sSQL = "UPDATE " & gConsPersona & "Persona SET "
        sSQL = sSQL & "cPersNombre = '" & sNombreCompleto & "',"
        sSQL = sSQL & "dPersNacCreac = '" & Format(dFechaNac, "mm/dd/yyyy") & "',"
        sSQL = sSQL & "cPersDireccUbiGeo = '" & sUbicGeografica & "',"
        sSQL = sSQL & "cPersDireccDomicilio = '" & sDomicilio & "',"
        sSQL = sSQL & "cPersDireccCondicion = '" & sCondicionDomic & "',"
        sSQL = sSQL & "nPersValComDomicilio = " & nValComDomicilio & ","
        sSQL = sSQL & "cPersTelefono = '" & sTelefonos & "',"
        sSQL = sSQL & "cPersTelefono2 = '" & sTelefonos2 & "',"
        sSQL = sSQL & "cPersEmail = '" & sEmail & "',"
        sSQL = sSQL & "nPersPersoneria = '" & Trim(Str(nTipoPersona)) & "',"
        sSQL = sSQL & "cPersCIIU = '" & sCiiu & "',"
        sSQL = sSQL & "cPersEstado = '" & sEstado & "',"
        sSQL = sSQL & "cPersCodSbs = '" & sPersCodSbs & "',"
        sSQL = sSQL & "nPersRelaInst = " & nPersRelInst
        sSQL = sSQL & " Where cPersCod = '" & sPersCod & "'"
        ConexionActiva.Execute sSQL
        
        If nTipoPersona = gPersonaNat Then
        
            sSQL = "DELETE " & gConsPersona & "PersonaJur"
            sSQL = sSQL & " WHERE cPersCod = '" & sPersCod & "'"
            ConexionActiva.Execute sSQL
        
            sSQL = "DELETE " & gConsPersona & "PersonaNat"
            sSQL = sSQL & " WHERE cPersCod = '" & sPersCod & "'"
            ConexionActiva.Execute sSQL
            
            'Actualiza Persona Natural
            sSQL = "INSERT INTO " & gConsPersona
            sSQL = sSQL & "PersonaNat(cPersCod,cPersnatSexo,nPersNatEstCiv,nPersNatHijos,nPersPeso,"
            sSQL = sSQL & "nPersTalla,nPersTpoSangre, cNacionalidad, nResidente, cUltimaActualizacion)"
            sSQL = sSQL & " VALUES('" & sPersCod & "',"
            sSQL = sSQL & "'" & sSexo & "',"
            sSQL = sSQL & sEstadoCivil & ","
            sSQL = sSQL & Trim(Str(nHijos)) & ","
            sSQL = sSQL & "'" & Format(nPeso, "#0.00") & "',"
            sSQL = sSQL & "'" & Format(nTalla, "#0.00") & "',"
            sSQL = sSQL & IIf(sTipoSangre = "", 0, sTipoSangre) & ","
            sSQL = sSQL & "'" & cNacionalidad & "',"
            sSQL = sSQL & nResidencia & ","
            sSQL = sSQL & "'" & sActualiza & "')"
            ConexionActiva.Execute sSQL
                        
        End If
        
        If nTipoPersona <> gPersonaNat Then
            sSQL = "DELETE " & gConsPersona & "PersonaNat"
            sSQL = sSQL & " WHERE cPersCod = '" & sPersCod & "'"
            ConexionActiva.Execute sSQL
            
            sSQL = "DELETE " & gConsPersona & "PersonaJur"
            sSQL = sSQL & " WHERE cPersCod = '" & sPersCod & "'"
            ConexionActiva.Execute sSQL
            
            'Actualiza Persona Juridica
            'Inserta PersonaJur
            sSQL = "INSERT INTO " & gConsPersona
            sSQL = sSQL & "PersonaJur(cPersCod,cPersJurTpo,cPersJurMagnitud,cPersJurSigla,nPersJurEmpleados,cUltimaActualizacion)"
            sSQL = sSQL & "VALUES('" & sPersCod & "',"
            sSQL = sSQL & "'" & sTipoPersJur & " ',"
            sSQL = sSQL & "'" & sMagnitudEmp & " ',"
            sSQL = sSQL & "'" & sSiglas & " ',"
            sSQL = sSQL & Trim(Str(nNumEmplead)) & ","
            sSQL = sSQL & "'" & sActualiza & "')"
            ConexionActiva.Execute sSQL
                                    
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
            sSQL = "INSERT INTO " & gConsPersona
            sSQL = sSQL & "Persona(cPersCod,cPersNombre,dPersNacCreac,cPersDireccUbiGeo,cPersDireccDomicilio"
            sSQL = sSQL & ",cPersDireccCondicion, nPersValComDomicilio, cPersTelefono,cPersTelefono2,nPersPersoneria,cPersCIIU,cPersEstado"
            sSQL = sSQL & ",cPersEmail,cUltimaActualizacion, cPersCodSbs, nPersRelaInst,dPersIng)"
            sSQL = sSQL & "VALUES('" & sPersCod & "',"
            sSQL = sSQL & "'" & sNombreCompleto & "',"
            sSQL = sSQL & "'" & Format(dFechaNac, "mm/dd/yyyy") & "',"
            sSQL = sSQL & "'" & sUbicGeografica & "',"
            sSQL = sSQL & "'" & sDomicilio & "',"
            sSQL = sSQL & "'" & sCondicionDomic & "',"
            sSQL = sSQL & nValComDomicilio & ","
            sSQL = sSQL & "'" & sTelefonos & "',"
            sSQL = sSQL & "'" & sTelefonos2 & "',"
            sSQL = sSQL & Trim(Str(nTipoPersona)) & ","
            sSQL = sSQL & "'" & sCiiu & "',"
            sSQL = sSQL & "'" & sEstado & "',"
            sSQL = sSQL & "'" & sEmail & "',"
            sSQL = sSQL & "'" & sActualiza & "','" & sPersCodSbs & "'," & nPersRelInst & ",'" & Format(dFechaHoy, "mm/dd/yyyy") & "')"
            ConexionActiva.Execute sSQL
            
            If nTipoPersona = gPersonaNat Then
                'Insertar en PersonaNat
                sSQL = "INSERT INTO " & gConsPersona
                sSQL = sSQL & "PersonaNat(cPersCod,cPersnatSexo,nPersNatEstCiv,nPersNatHijos,nPersPeso,"
                sSQL = sSQL & "nPersTalla,nPersTpoSangre, cNacionalidad, nResidente, cUltimaActualizacion)"
                sSQL = sSQL & " VALUES('" & sPersCod & "',"
                sSQL = sSQL & "'" & sSexo & "',"
                sSQL = sSQL & sEstadoCivil & ","
                sSQL = sSQL & Trim(Str(nHijos)) & ","
                sSQL = sSQL & "'" & Format(nPeso, "#0.00") & "',"
                sSQL = sSQL & "'" & Format(nTalla, "#0.00") & "',"
                sSQL = sSQL & IIf(sTipoSangre = "", 0, sTipoSangre) & ","
                sSQL = sSQL & "'" & cNacionalidad & "',"
                sSQL = sSQL & nResidencia & ","
                sSQL = sSQL & "'" & sActualiza & "')"
                ConexionActiva.Execute sSQL
            End If
            
            If nTipoPersona <> gPersonaNat Then
                'Inserta PersonaJur
                sSQL = "INSERT INTO " & gConsPersona
                sSQL = sSQL & "PersonaJur(cPersCod,cPersJurTpo,cPersJurMagnitud,cPersJurSigla,nPersJurEmpleados,cUltimaActualizacion)"
                sSQL = sSQL & "VALUES('" & sPersCod & "',"
                sSQL = sSQL & "'" & sTipoPersJur & " ',"
                sSQL = sSQL & "'" & sMagnitudEmp & " ',"
                sSQL = sSQL & "'" & sSiglas & " ',"
                sSQL = sSQL & Trim(Str(nNumEmplead)) & ","
                sSQL = sSQL & "'" & sActualiza & "')"
                ConexionActiva.Execute sSQL
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
            sSQL = "UPDATE " & gConsPersona & "PersRelaciones SET "
            sSQL = sSQL & "cPersRelacPersCod = '" & Relac(i).sPersCod & "',"
            sSQL = sSQL & "nPersRelac = " & Trim(Right(Relac(i).sRelacion, 20)) & ","
            sSQL = sSQL & "bPersRelacBenef = " & Trim(Right(Relac(i).sPersRelacBenef, 20)) & ","
            sSQL = sSQL & "nPersRelacBenefPorc = " & Format(Relac(i).nPersRelacBenefPorc, "#0.00") & ","
            sSQL = sSQL & "bPersRelacAMP = " & Trim(Right(Relac(i).sPersRelacAMP, 20))
            'sSql = sSql & "PR.bPersRelacEstado = " & Trim(Right(Relac(i).sPersRelacEstado, 20))
            sSQL = sSQL & " Where cPersCod = '" & sPersCod & "' And cPersRelacPersCod = '" & Relac(i).sPersCodOrig & "'"
            ConexionActiva.Execute sSQL
        Else
            If Relac(i).nTipoAct = PersFilaNueva Then
                sSQL = "INSERT INTO " & gConsPersona & "PERSRELACIONES(cPersCod,cPersRelacPersCod,"
                sSQL = sSQL & "nPersRelac,bPersrelacBenef,nPersRelacBenefPorc,bPersRelacAMP,cMovNro)"
                sSQL = sSQL & "VALUES('" & sPersCod & "',"
                sSQL = sSQL & "'" & Relac(i).sPersCod & "',"
                sSQL = sSQL & Trim(Right(Relac(i).sRelacion, 20)) & ","
                sSQL = sSQL & Trim(Right(Relac(i).sPersRelacBenef, 20)) & ","
                sSQL = sSQL & Format(Relac(i).nPersRelacBenefPorc, "#0.00") & ","
                sSQL = sSQL & Trim(Right(Relac(i).sPersRelacAMP, 20)) & ","
                sSQL = sSQL & "'" & sActualiza & "')"
                ConexionActiva.Execute sSQL
            Else
                If Relac(i).nTipoAct = PersFilaEliminda Then
                    sSQL = "DELETE " & gConsPersona & "PersRelaciones Where "
                    sSQL = sSQL & " cPersCod = '" & sPersCod & "'"
                    sSQL = sSQL & " And cPersRelacPersCod = '" & Trim(Relac(i).sPersCodOrig) & "'"
                    ConexionActiva.Execute sSQL
                End If
            End If
        End If
    Next i
    
    'Graba Referencias Comerciales
    For i = 0 To nNumRefCom - 1
        If RefCom(i).nTipoAct = PersFilaModificada Then
            sSQL = "UPDATE PERSREFCOMERCIAL SET "
            sSQL = sSQL & "cNomRefCom= '" & RefCom(i).sRefComNom & "',"
            sSQL = sSQL & "nTipoRelaCom = " & Right(RefCom(i).sRefComTipo, 1) & ","
            sSQL = sSQL & "cFonoRefCom = '" & RefCom(i).sRefComFono & "',"
            sSQL = sSQL & "cUltimaActualizacion = '" & sActualiza & "' "
            sSQL = sSQL & " WHERE cPersCod = '" & sPersCod & "' AND nNumRefCom = " & RefCom(i).nRefComCod
            ConexionActiva.Execute sSQL
        Else
            If RefCom(i).nTipoAct = PersFilaNueva Then
                sSQL = "INSERT INTO PERSREFCOMERCIAL (cPersCod, nNumRefCom, cNomRefCom, nTipoRelaCom, "
                sSQL = sSQL & "cFonoRefCom, cUltimaActualizacion)"
                sSQL = sSQL & " VALUES ('" & sPersCod & "',"
                sSQL = sSQL & RefCom(i).nRefComCod & ","
                sSQL = sSQL & "'" & RefCom(i).sRefComNom & "',"
                sSQL = sSQL & Right(RefCom(i).sRefComTipo, 1) & ","
                sSQL = sSQL & "'" & RefCom(i).sRefComFono & "',"
                sSQL = sSQL & "'" & sActualiza & "')"
                ConexionActiva.Execute sSQL
            Else
                If RefCom(i).nTipoAct = PersFilaEliminda Then
                    sSQL = "DELETE PERSREFCOMERCIAL "
                    sSQL = sSQL & " WHERE cPersCod = '" & sPersCod & "' AND nNumRefCom = " & RefCom(i).nRefComCod
                    ConexionActiva.Execute sSQL
                End If
            End If
        End If
    Next i
    
    'Graba Referencias Bancarias
    For i = 0 To nNumRefBan - 1
        If RefBan(i).nTipoAct = PersFilaModificada Then
            sSQL = "UPDATE PERSREFBANCARIA SET "
            sSQL = sSQL & "cCodBanco = '" & Mid(RefBan(i).sRefBanCodIF, 4, Len(RefBan(i).sRefBanCodIF) - 3) & "',"
            sSQL = sSQL & "cNumCta = '" & RefBan(i).sRefBanNumCta & "',"
            sSQL = sSQL & "cNumTarjeta = '" & RefBan(i).sRefBanNumTar & "',"
            sSQL = sSQL & "nMontoLinCred = " & RefBan(i).nRefBanLinCred & ","
            sSQL = sSQL & "cUltimaActualizacion = '" & sActualiza & "' "
            sSQL = sSQL & " WHERE cPersCod = '" & sPersCod & "' AND cCodBanco = '" & Mid(RefBan(i).sRefBanCodIFOrg, 4, Len(RefBan(i).sRefBanCodIFOrg) - 3) & "'"
            ConexionActiva.Execute sSQL
        Else
            If RefBan(i).nTipoAct = PersFilaNueva Then
                sSQL = "INSERT INTO PERSREFBANCARIA (cPersCod, cCodBanco, cNumCta, cNumTarjeta, "
                sSQL = sSQL & "nMontoLinCred, cUltimaActualizacion)"
                sSQL = sSQL & " VALUES ('" & sPersCod & "',"
                sSQL = sSQL & "'" & Mid(RefBan(i).sRefBanCodIF, 4, Len(RefBan(i).sRefBanCodIF) - 3) & "',"
                sSQL = sSQL & "'" & RefBan(i).sRefBanNumCta & "',"
                sSQL = sSQL & "'" & RefBan(i).sRefBanNumTar & "',"
                sSQL = sSQL & RefBan(i).nRefBanLinCred & ","
                sSQL = sSQL & "'" & sActualiza & "')"
                ConexionActiva.Execute sSQL
            Else
                If RefBan(i).nTipoAct = PersFilaEliminda Then
                    sSQL = "DELETE PERSREFBANCARIA "
                    sSQL = sSQL & " WHERE cPersCod = '" & sPersCod & "' AND cCodBanco = '" & Mid(RefBan(i).sRefBanCodIFOrg, 4, Len(RefBan(i).sRefBanCodIFOrg) - 3) & "'"
                    ConexionActiva.Execute sSQL
                End If
            End If
        End If
    Next i
    
    'Graba Patrimonio Vehicular
    For i = 0 To nNumPatVeh - 1
        If PatVeh(i).nTipoAct = PersFilaModificada Then
            sSQL = "UPDATE PERSPATRIMONIO SET "
            sSQL = sSQL & "cMarca= '" & PatVeh(i).sPatVehMarca & "',"
            sSQL = sSQL & "cFecFabri = '" & PatVeh(i).sPatVehFecFab & "',"
            sSQL = sSQL & "nValorCom = " & PatVeh(i).nPatVehValCom & ","
            sSQL = sSQL & "nCondPatri = " & Right(PatVeh(i).sPatVehCondicion, 1) & ","
            sSQL = sSQL & "cUltimaActualizacion = '" & sActualiza & "' "
            sSQL = sSQL & " WHERE cPersCod = '" & sPersCod & "' AND nNumPatri= " & PatVeh(i).nPatVehCod
            ConexionActiva.Execute sSQL
        Else
            If PatVeh(i).nTipoAct = PersFilaNueva Then
                sSQL = "INSERT INTO PERSPATRIMONIO(cPersCod, nNumPatri, nCondPatri, cMarca, cFecFabrica, "
                sSQL = sSQL & "nValorCom, cUltimaActualizacion)"
                sSQL = sSQL & " VALUES ('" & sPersCod & "',"
                sSQL = sSQL & PatVeh(i).nPatVehCod & ","
                sSQL = sSQL & Right(PatVeh(i).sPatVehCondicion, 1) & ","
                sSQL = sSQL & "'" & PatVeh(i).sPatVehMarca & "',"
                sSQL = sSQL & "'" & PatVeh(i).sPatVehFecFab & "',"
                sSQL = sSQL & PatVeh(i).nPatVehValCom & ","
                sSQL = sSQL & "'" & sActualiza & "')"
                ConexionActiva.Execute sSQL
            Else
                If PatVeh(i).nTipoAct = PersFilaEliminda Then
                    sSQL = "DELETE PERSPATRIMONIO "
                    sSQL = sSQL & " WHERE cPersCod = '" & sPersCod & "' AND nNumPatri= " & PatVeh(i).nPatVehCod
                    ConexionActiva.Execute sSQL
                End If
            End If
        End If
    Next i
    
    'Actualizar Los Documentos de las Personas
    For i = 0 To nNumDoc - 1
        If Docs(i).nTipoAct = PersFilaModificada Then
            sSQL = "UPDATE " & gConsPersona & "PERSID SET "
            sSQL = sSQL & "cPersIdTpo = '" & Trim(Str(Docs(i).nTipo)) & "',"
            sSQL = sSQL & "cPersIDNro = '" & Docs(i).sNumero & "',"
            sSQL = sSQL & "cUltimaActualizacion = '" & sActualiza & "' "
            sSQL = sSQL & " WHERE cPersCod = '" & sPersCod & "' And cPersIDTpo = '" & Trim(Str(Docs(i).nTipoOrig)) & "'"
            ConexionActiva.Execute sSQL
        Else
            If Docs(i).nTipoAct = PersFilaNueva Then
                sSQL = "INSERT INTO " & gConsPersona & "PERSID("
                sSQL = sSQL & "cPersCod,cPersIDTpo,cPersIDNro,cUltimaActualizacion)"
                sSQL = sSQL & "VALUES('" & sPersCod & "',"
                sSQL = sSQL & "'" & Trim(Str(Docs(i).nTipo)) & "',"
                sSQL = sSQL & "'" & Docs(i).sNumero & "',"
                sSQL = sSQL & "'" & sActualiza & "')"
                ConexionActiva.Execute sSQL
            Else
                If Docs(i).nTipoAct = PersFilaEliminda Then
                    sSQL = "DELETE " & gConsPersona & "PERSID Where "
                    sSQL = sSQL & " cPersCod = '" & sPersCod & "'"
                    sSQL = sSQL & " And cPersIDTpo = '" & Trim(Str(Docs(i).nTipoOrig)) & "'"
                    ConexionActiva.Execute sSQL
                End If
            End If
        End If
    Next i
    
    'Actualizar las Fuentes de Ingreso
    For i = 0 To nNumFtes - 1
        FtesIngreso(i).sRazSocDescrip = Replace(FtesIngreso(i).sRazSocDescrip, "'", "''")
        If FtesIngreso(i).nTipoAct = PersFilaNueva Then
            FtesIngreso(i).cNumFte = sNumFuente
            sSQL = "INSERT INTO " & gConsPersona & "PersFteIngreso(cNumFuente, cPersCod, cPersFIPersCod, cPersFICargo,"
            sSQL = sSQL & "dPersFIinicio, cPersFIMoneda, cPersFIComentario, cRazSocUbiGeo, cRazSocDirecc, "
            sSQL = sSQL & " cRazSocTelef, cRazSocDescrip, nPersTipFte, cUltimaActualizacion,bCostoProd) "
            sSQL = sSQL & " VALUES('" & FtesIngreso(i).cNumFte & "','" & sPersCod & "',"
            sSQL = sSQL & "'" & FtesIngreso(i).sFteIngreso & "',"
            sSQL = sSQL & "'" & FtesIngreso(i).sCargo & "',"
            sSQL = sSQL & "'" & Format(FtesIngreso(i).dFteIngreso, "mm/dd/yyyy") & "',"
            sSQL = sSQL & Trim(FtesIngreso(i).sMoneda) & ","
            sSQL = sSQL & "'" & Trim(FtesIngreso(i).sComentarios) & "',"
            sSQL = sSQL & "'" & Right(FtesIngreso(i).sRazSocUbiGeo, 12) & "',"
            sSQL = sSQL & "'" & FtesIngreso(i).sRazSocDirecc & "',"
            sSQL = sSQL & "'" & FtesIngreso(i).sRazSocTelef & "',"
            sSQL = sSQL & "'" & FtesIngreso(i).sRazSocDescrip & "',"
            sSQL = sSQL & FtesIngreso(i).nTipo & ",'" & sActualiza & "'," & FtesIngreso(i).bCostoProd
            sSQL = sSQL & ")"
            ConexionActiva.Execute sSQL
        Else
            If FtesIngreso(i).nTipoAct = PersFilaModificada Then
                sSQL = "UPDATE " & gConsPersona & "PersFteIngreso SET "
                sSQL = sSQL & " cPersFIPersCod = '" & FtesIngreso(i).sFteIngreso & "',"
                sSQL = sSQL & " cPersFICargo = '" & FtesIngreso(i).sCargo & "',"
                sSQL = sSQL & " dPersFIInicio = '" & Format(FtesIngreso(i).dFteIngreso, "mm/dd/yyyy") & "',"
                sSQL = sSQL & " cPersFIMoneda = " & Trim(FtesIngreso(i).sMoneda) & ","
                sSQL = sSQL & " cPersFIComentario = '" & Trim(FtesIngreso(i).sComentarios) & "',"
                sSQL = sSQL & " cUltimaActualizacion = '" & sActualiza & "',"
                sSQL = sSQL & " cRazSocUbiGeo = '" & FtesIngreso(i).sRazSocUbiGeo & "',"
                sSQL = sSQL & " cRazSocDirecc = '" & FtesIngreso(i).sRazSocDirecc & "',"
                sSQL = sSQL & " cRazSocTelef = '" & FtesIngreso(i).sRazSocTelef & "',"
                sSQL = sSQL & " cRazSocDescrip = '" & FtesIngreso(i).sRazSocDescrip & "', "
                sSQL = sSQL & " bCostoProd = " & FtesIngreso(i).bCostoProd
                sSQL = sSQL & " WHERE cNumFuente = '" & FtesIngreso(i).cNumFte & "'"
                ConexionActiva.Execute sSQL
            Else
                If FtesIngreso(i).nTipoAct = PersFilaEliminda Then
                    sSQL = "DELETE PersFIIndependiente Where cNumFuente = '" & FtesIngreso(i).cNumFte & "' "
                    ConexionActiva.Execute sSQL
                    
                    sSQL = "DELETE PersFIDependiente Where cNumFuente = '" & FtesIngreso(i).cNumFte & "' "
                    ConexionActiva.Execute sSQL
                        
                    sSQL = "DELETE PersFICostoProd Where cNumFuente = '" & FtesIngreso(i).cNumFte & "' "
                    ConexionActiva.Execute sSQL
                    
                    sSQL = "DELETE " & gConsPersona & "PersFteIngreso "
                    sSQL = sSQL & " WHERE cNumFuente = '" & FtesIngreso(i).cNumFte & "'"
                    ConexionActiva.Execute sSQL
                End If
            End If
        End If
    Next i
    
    'FUENTES DE INGRESO DEPENDIENTES
    For i = 0 To nNumFtes - 1
        For J = 0 To FtesIngreso(i).nNumFtesDep - 1
            If FtesIngreso(i).FtesIngresoDep(J).nTipoAct = PersFilaNueva Then
                sSQL = "INSERT INTO PersFIDependiente(cNumFuente, dPersEval, nPersIngCli, "
                sSQL = sSQL & " nPersGastoFam, nPersIngCon, nPersOtrIng, dPersFICaduca) "
                sSQL = sSQL & " VALUES('" & FtesIngreso(i).cNumFte & "',"
                sSQL = sSQL & "'" & Format(FtesIngreso(i).FtesIngresoDep(J).dFecEval, "mm/dd/yyyy") & "',"
                sSQL = sSQL & Format(FtesIngreso(i).FtesIngresoDep(J).nIngresos, "#0.00") & ","
                sSQL = sSQL & Format(FtesIngreso(i).FtesIngresoDep(J).nGastosFam, "#0.00") & ","
                sSQL = sSQL & Format(FtesIngreso(i).FtesIngresoDep(J).nIngCony, "#0.00") & ","
                sSQL = sSQL & Format(FtesIngreso(i).FtesIngresoDep(J).nIngOtros, "#0.00") & ","
                sSQL = sSQL & "'" & Format(FtesIngreso(i).FtesIngresoDep(J).dFecCad, "mm/dd/yyyy") & "')"
                ConexionActiva.Execute sSQL
            Else
                If FtesIngreso(i).FtesIngresoDep(J).nTipoAct = PersFilaModificada Then
                    sSQL = "UPDATE PersFIDependiente SET "
                    sSQL = sSQL & " nPersIngCli = " & Format(FtesIngreso(i).FtesIngresoDep(J).nIngFam, "#0.00") & ","
                    sSQL = sSQL & " nPersGastoFam = " & Format(FtesIngreso(i).FtesIngresoDep(J).nGastosFam, "#0.00") & ","
                    sSQL = sSQL & " nPersIngCon = " & Format(FtesIngreso(i).FtesIngresoDep(J).nIngCony, "#0.00") & ","
                    sSQL = sSQL & " nPersOtrIng = " & Format(FtesIngreso(i).FtesIngresoDep(J).nIngOtros, "#0.00") & ","
                    sSQL = sSQL & " dPersFICaduca = '" & Format(FtesIngreso(i).FtesIngresoDep(J).dFecCad, "mm/dd/yyyy") & "' "
                    sSQL = sSQL & " Where cNumFuente = '" & FtesIngreso(i).cNumFte & "' AND dPersEval = '" & Format(FtesIngreso(i).FtesIngresoDep(J).dFecEval, "mm/dd/yyyy") & "'"
                    ConexionActiva.Execute sSQL
                Else
                    If FtesIngreso(i).FtesIngresoDep(J).nTipoAct = PersFilaEliminda Then
                        sSQL = "DELETE PersFIDependiente Where cNumFuente = '" & FtesIngreso(i).cNumFte & "' AND dPersEval = '" & Format(FtesIngreso(i).FtesIngresoDep(J).dFecEval, "mm/dd/yyyy") & "'"
                        ConexionActiva.Execute sSQL
                    End If
                End If
            End If
        Next J
    Next i
    
    'FUENTES DE INGRESO INDEPENDIENTES
    For i = 0 To nNumFtes - 1
        For J = 0 To FtesIngreso(i).nNumFtesIndep - 1
            If FtesIngreso(i).FtesIngresoInd(J).nTipoAct = PersFilaNueva Then
                sSQL = "INSERT INTO PersFIIndependiente(cNumFuente, dPersEval, "
                sSQL = sSQL & " dPersFICaduca, nPersFIActivoDisp, nPersFICtasxCobrar, nPersFIInventarios, "
                sSQL = sSQL & " nPersFIActivosFijos, nPersFIProveedores, nPersFICredCMACT, nPersFICredOtros, "
                sSQL = sSQL & " nPersFIVentas, nPersFIPatrimonio, nPersFICostoVentas, nPersFIRecupCtasXCobrar, "
                sSQL = sSQL & " nPersFIEgresosOtros, nPersIngFam, nPersEgrFam, cUltimaActualizacion)"
                sSQL = sSQL & " VALUES('" & FtesIngreso(i).cNumFte & "',"
                sSQL = sSQL & "'" & Format(FtesIngreso(i).FtesIngresoInd(J).dFecEval, "mm/dd/yyyy") & "',"
                sSQL = sSQL & "'" & Format(FtesIngreso(i).FtesIngresoInd(J).dFecCad, "mm/dd/yyyy") & "',"
                sSQL = sSQL & Format(FtesIngreso(i).FtesIngresoInd(J).nActivDisp, "#0.00") & ","
                sSQL = sSQL & Format(FtesIngreso(i).FtesIngresoInd(J).nCtasxCob, "#0.00") & ","
                sSQL = sSQL & Format(FtesIngreso(i).FtesIngresoInd(J).nInventarios, "#0.00") & ","
                sSQL = sSQL & Format(FtesIngreso(i).FtesIngresoInd(J).nActFijo, "#0.00") & ","
                sSQL = sSQL & Format(FtesIngreso(i).FtesIngresoInd(J).nProveedores, "#0.00") & ","
                sSQL = sSQL & Format(FtesIngreso(i).FtesIngresoInd(J).nCreditosCmact, "#0.00") & ","
                sSQL = sSQL & Format(FtesIngreso(i).FtesIngresoInd(J).nCreditosOtros, "#0.00") & ","
                sSQL = sSQL & Format(FtesIngreso(i).FtesIngresoInd(J).nVentas, "#0.00") & ","
                sSQL = sSQL & Format(FtesIngreso(i).FtesIngresoInd(J).nPasivoPatrimonio, "#0.00") & ","
                sSQL = sSQL & Format(FtesIngreso(i).FtesIngresoInd(J).nCompraMercad, "#0.00") & ","
                sSQL = sSQL & Format(FtesIngreso(i).FtesIngresoInd(J).nRecupCtasxCob, "#0.00") & ","
                sSQL = sSQL & Format(FtesIngreso(i).FtesIngresoInd(J).nOtrosEgresos, "#0.00") & ","
                sSQL = sSQL & Format(FtesIngreso(i).FtesIngresoInd(J).nPersIngFam, "#0.00") & ","
                sSQL = sSQL & Format(FtesIngreso(i).FtesIngresoInd(J).nPersEgrFam, "#0.00") & ","
                sSQL = sSQL & "'" & sActualiza & "')"
                ConexionActiva.Execute sSQL
                
            Else
                If FtesIngreso(i).FtesIngresoInd(J).nTipoAct = PersFilaModificada Then
                    sSQL = " UPDATE PersFIIndependiente SET "
                    sSQL = sSQL & " nPersFIActivoDisp = " & Format(FtesIngreso(i).FtesIngresoInd(J).nActivDisp, "#0.00") & ","
                    sSQL = sSQL & " nPersFICtasxCobrar = " & Format(FtesIngreso(i).FtesIngresoInd(J).nCtasxCob, "#0.00") & ","
                    sSQL = sSQL & " nPersFIInventarios = " & Format(FtesIngreso(i).FtesIngresoInd(J).nInventarios, "#0.00") & ","
                    sSQL = sSQL & " nPersFIActivosFijos = " & Format(FtesIngreso(i).FtesIngresoInd(J).nActFijo, "#0.00") & ","
                    sSQL = sSQL & " nPersFIProveedores = " & Format(FtesIngreso(i).FtesIngresoInd(J).nProveedores, "#0.00") & ","
                    sSQL = sSQL & " nPersFICredCMACT = " & Format(FtesIngreso(i).FtesIngresoInd(J).nCreditosCmact, "#0.00") & ","
                    sSQL = sSQL & " nPersFICredOtros = " & Format(FtesIngreso(i).FtesIngresoInd(J).nCreditosOtros, "#0.00") & ","
                    sSQL = sSQL & " nPersFIVentas = " & Format(FtesIngreso(i).FtesIngresoInd(J).nVentas, "#0.00") & ","
                    sSQL = sSQL & " nPersFIPatrimonio = " & Format(FtesIngreso(i).FtesIngresoInd(J).nPasivoPatrimonio, "#0.00") & ","
                    sSQL = sSQL & " nPersFICostoVentas = " & Format(FtesIngreso(i).FtesIngresoInd(J).nCompraMercad, "#0.00") & ","
                    sSQL = sSQL & " nPersFIRecupCtasXCobrar = " & Format(FtesIngreso(i).FtesIngresoInd(J).nRecupCtasxCob, "#0.00") & ","
                    sSQL = sSQL & " nPersFIEgresosOtros = " & Format(FtesIngreso(i).FtesIngresoInd(J).nOtrosEgresos, "#0.00") & ","
                    sSQL = sSQL & " nPersIngFam = " & Format(FtesIngreso(i).FtesIngresoInd(J).nPersIngFam, "#0.00") & ","
                    sSQL = sSQL & " nPersEgrFam = " & Format(FtesIngreso(i).FtesIngresoInd(J).nPersEgrFam, "#0.00") & ","
                    sSQL = sSQL & " dPersFICaduca = '" & Format(FtesIngreso(i).FtesIngresoInd(J).dFecCad, "mm/dd/yyyy") & "' "
                    sSQL = sSQL & " Where cNumFuente = '" & FtesIngreso(i).cNumFte & "' AND dPersEval = '" & Format(FtesIngreso(i).FtesIngresoInd(J).dFecEval, "mm/dd/yyyy") & "'"
                    
                    ConexionActiva.Execute sSQL
                Else
                    If FtesIngreso(i).FtesIngresoInd(J).nTipoAct = PersFilaEliminda Then
                        sSQL = "DELETE PersFIIndependiente Where cNumFuente = '" & FtesIngreso(i).cNumFte & "' AND dPersEval = '" & Format(FtesIngreso(i).FtesIngresoInd(J).dFecEval, "mm/dd/yyyy") & "'"
                        ConexionActiva.Execute sSQL
                    End If
                End If
            End If
        Next J
    Next i
    
    'Actualiza Costos de Produccion
    For i = 0 To nNumFtes - 1
        For J = 0 To FtesIngreso(i).nNumCostosProd - 1
            If FtesIngreso(i).CostoProd(J).nTipoAct = PersFilaNueva Then
                sSQL = "INSERT INTO PersFICostoProd(cNumFuente, dFecEval, nTpoCultivo, nMaquinaria, nJornales, nInsumos, nPesticidas, nOtros, nHectareas, "
                sSQL = sSQL & " nProduccion, nUniProd, nPreUni,nSiembra,nCosecha,nDesaAgricola,nPlanOtros,nMantenimiento)"
                sSQL = sSQL & " VALUES('" & FtesIngreso(i).cNumFte & "','" & Format(FtesIngreso(i).CostoProd(J).dFecEval, "mm/dd/yyyy") & "'," & FtesIngreso(i).CostoProd(J).nTpoCultivo & ","
                sSQL = sSQL & Format(FtesIngreso(i).CostoProd(J).nMaquinaria, "#0.00") & "," & Format(FtesIngreso(i).CostoProd(J).nJornales, "#0.00") & ","
                sSQL = sSQL & Format(FtesIngreso(i).CostoProd(J).nInsumos, "#0.00") & "," & Format(FtesIngreso(i).CostoProd(J).nPesticidas, "#0.00") & ","
                sSQL = sSQL & Format(FtesIngreso(i).CostoProd(J).nOtros, "#0.00") & "," & Format(FtesIngreso(i).CostoProd(J).nHectareas, "#0.00") & ","
                sSQL = sSQL & Format(FtesIngreso(i).CostoProd(J).nProduccion, "#0.00") & "," & Format(FtesIngreso(i).CostoProd(J).nUniProd, "#0.00") & "," & Format(FtesIngreso(i).CostoProd(J).nPreUni, "#0.00") & ","
                sSQL = sSQL & FtesIngreso(i).CostoProd(J).nSiembra & ","
                sSQL = sSQL & FtesIngreso(i).CostoProd(J).nCosecha & "," & FtesIngreso(i).CostoProd(J).nDesAgricola & "," & FtesIngreso(i).CostoProd(J).nPlanOtros & ","
                sSQL = sSQL & FtesIngreso(i).CostoProd(J).nMantenimiento & ")"
                                
                ConexionActiva.Execute sSQL
                
            Else
                If FtesIngreso(i).CostoProd(J).nTipoAct = PersFilaModificada Then
                    sSQL = "UPDATE PersFICostoProd SET "
                    sSQL = sSQL & "nTpoCultivo = " & FtesIngreso(i).CostoProd(J).nTpoCultivo & ","
                    sSQL = sSQL & "nMaquinaria = " & Format(FtesIngreso(i).CostoProd(J).nMaquinaria, "#0.00") & ","
                    sSQL = sSQL & "nJornales = " & Format(FtesIngreso(i).CostoProd(J).nJornales, "#0.00") & ","
                    sSQL = sSQL & "nInsumos = " & Format(FtesIngreso(i).CostoProd(J).nInsumos, "#0.00") & ","
                    sSQL = sSQL & "nPesticidas = " & Format(FtesIngreso(i).CostoProd(J).nPesticidas, "#0.00") & ","
                    sSQL = sSQL & "nOtros = " & Format(FtesIngreso(i).CostoProd(J).nOtros, "#0.00") & ","
                    sSQL = sSQL & "nHectareas = " & Format(FtesIngreso(i).CostoProd(J).nHectareas, "#0.00") & ","
                    sSQL = sSQL & "nProduccion = " & Format(FtesIngreso(i).CostoProd(J).nProduccion, "#0.00") & ","
                    sSQL = sSQL & "nUniProd = " & FtesIngreso(i).CostoProd(J).nUniProd & ","
                    sSQL = sSQL & "nPreUni = " & Format(FtesIngreso(i).CostoProd(J).nPreUni, "#0.00") & ","
                    sSQL = sSQL & "nSiembra=" & FtesIngreso(i).CostoProd(J).nSiembra & ","
                    sSQL = sSQL & "nCosecha=" & FtesIngreso(i).CostoProd(J).nMantenimiento & ","
                    sSQL = sSQL & "nDesaAgricola=" & FtesIngreso(i).CostoProd(J).nDesAgricola & ","
                    sSQL = sSQL & "nPlanOtros=" & FtesIngreso(i).CostoProd(J).nPlanOtros & ","
                    sSQL = sSQL & "nMantenimiento= " & FtesIngreso(i).CostoProd(J).nMantenimiento & " "
                    sSQL = sSQL & " WHERE cNumFuente = '" & FtesIngreso(i).cNumFte & "' AND dFecEval = '" & Format(FtesIngreso(i).CostoProd(J).dFecEval, "mm/dd/yyyy") & "'"
                    
                    ConexionActiva.Execute sSQL
                Else
                    If FtesIngreso(i).CostoProd(J).nTipoAct = PersFilaEliminda Then
                        sSQL = "DELETE PersFICostoProd Where cNumFuente = '" & FtesIngreso(i).cNumFte & "'"
                        ConexionActiva.Execute sSQL
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
                    sSQL = "INSERT INTO Persona(cCodPers,cNomPers,cTidoci,cNudoci,cTidotr,cNudotr,cDirPers, cCodZon, "
                    sSQL = sSQL & " dFecNac, cTelPers, cEstPers, cTipPers, cCodSBS, dIngPers, dUltAct, cCodUsu, cRelIns, cConDom, cFlagConsol, cEMail, cWeb, cTelefono2) "
                    sSQL = sSQL & " VALUES('" & sCodPersDistrib & "','" & sNombreCompleto & "','" & vDocs(0) & "','" & vDocs(1) & "','" & vDocs(2) & "','" & vDocs(3) & "','"
                    sSQL = sSQL & sDomicilio & "','" & sUbicGeografica & "','" & Format(dFechaNac, "mm/dd/yyyy") & "','" & sTelefonos & "','1','"
                    sSQL = sSQL & CInt(nTipoPersona) & "','" & sPersCodSbs & "','" & Mid(sActualiza, 1, 4) & "/" & Mid(sActualiza, 5, 2) & "/" & Mid(sActualiza, 7, 2) & "','" & Mid(sActualiza, 1, 4) & "/" & Mid(sActualiza, 5, 2) & "/" & Mid(sActualiza, 7, 2)
                    sSQL = sSQL & "','','" & sPersRelIDistrib & "','" & sCondicionDomic & "',NULL,'','','')"
    
                    ConexDitrib.Execute sSQL
                End If
    
                '******************************************************************************
                'ACTUALIZA PERSONA NATURAL EN DISTRIBUIDO
                '******************************************************************************
                If bPersDistrib Then
                    sSQL = "INSERT INTO Personanat(cCodPers, cSexPers, cEstCiv, cCodProf, cCodCon, nNumHij)"
                    sSQL = sSQL & " VALUES('" & sCodPersDistrib & "','" & sSexo & "','" & sEstadoCivil & "','',''," & nHijos & ")"
                    ConexDitrib.Execute sSQL
                End If
    
                '******************************************************************************
                'ACTUALIZA PERSONA JRUDICA EN DISTRIBUIDO
                '******************************************************************************
                If bPersDistrib Then
                    sSQL = "Insert into PersonaJur(cCodPers,cTipPerJur,cCodRegPub,cMagEmp,cSigla) "
                    sSQL = sSQL & "Values('" & sCodPersDistrib & "','0000','','','" & sSiglas & "')"
                    ConexDitrib.Execute sSQL
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
Dim Conn As DConecta
Dim sSQL As String
Dim R As ADODB.Recordset

    sSQL = "Select C.cConsDescripcion,PI.cPersCod,PI.cPersIDtpo,PI.cPersIDnro from PersID PI left join Constante C ON PI.cPersIDtpo = C.cConsValor"
    sSQL = sSQL & " where PI.cPersCod = '" & psCodPers & "' And C.cConsCod = " & Trim(Str(gPersIdTipo))
    Set R = Conn.CargaRecordSet(sSQL, adLockReadOnly)
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
Dim sSQL As String
Dim R As ADODB.Recordset
Dim oConecta As DConecta
Dim sCadCon As String

    'sSQL = "Provider=SQLOLEDB.1;User ID=dbaccess;Initial Catalog=DbCmactAux;Data Source=192.168.1.16;PWD=proyecto"
        
    If ConexionActiva Is Nothing Then
        Set oConecta = New DConecta
        oConecta.AbreConexion
        Set ConexionActiva = oConecta.ConexionActiva
    End If
    
    sCadCon = ""
    sSQL = "Select * from Servidor Where RIGHT(cCodAge,2) = '" & sCodAge & "' AND cNroSer = '05'"
    Set R = New ADODB.Recordset
    R.Open sSQL, ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    If R.RecordCount > 0 Then
        sCadCon = "Provider=" & Trim(R!cProvider) & ";User ID=" & Trim(R!cLogin) & ";Initial Catalog=" & Trim(R!cDataBase) & ";Data Source=" & Trim(R!cNomSer) & ";PWD=" & Trim(R!cPass)
    End If
    R.Close
    
    If sCadCon <> "" Then
        Set oConecta = New DConecta
        Call oConecta.AbreConexion(sCadCon)
        Set ConexionActiva = oConecta.ConexionActiva
    End If
    sSQL = "Select cPersCod,iPersFirma,cUltimaActualizacion from " & gConsImagenes & "PersImagen PI Where cPersCod = '" & psCodPers & "'"
    Set R = New ADODB.Recordset
    R.Open sSQL, ConexionActiva, adOpenKeyset, adLockOptimistic
'    If R.BOF And R.EOF Then
'        R.Close
'        Set RecuperaFirma = Nothing
'        Exit Function
'    End If
    Set RecuperaFirma = R
End Function

Public Sub RecuperaDocumentos(ByVal psCodPers As String)
Dim sSQL As String
Dim R As ADODB.Recordset
    nNumDoc = 0
    
    ReDim Docs(nNumDoc)
    sSQL = "Select PID.cPersIDnro,PID.cPersIDTpo,C.cConsDescripcion,C.nConsValor From " & gConsPersona & "PersID PID inner join Constante C ON convert(int,PID.cPersIDTpo) = C.nConsValor Where PID.cPersCod = '" & psCodPers & "' And  C.nConsCod = " & Trim(Str(gPersIdTipo)) & " And C.cConsDescripcion <> '" & Trim(Str(gPersIdTipo)) & "'"
    Set R = Conn.CargaRecordSet(sSQL)
    Do While Not R.EOF
        nNumDoc = nNumDoc + 1
        ReDim Preserve Docs(nNumDoc)
        Docs(nNumDoc - 1).nTipoAct = 0
        Docs(nNumDoc - 1).sTipo = R!cConsDescripcion & Space(75) & R!nConsValor
        Docs(nNumDoc - 1).nTipo = CInt(R!cPersIDTpo)
        Docs(nNumDoc - 1).nTipoOrig = CInt(R!cPersIDTpo)
        Docs(nNumDoc - 1).sNumero = Trim(R!cPersIDNro)
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
End Sub
Private Sub RecuperaRelacionesdePersona(ByVal psPersCod As String)
Dim sSQL As String
Dim R As New ADODB.Recordset
    nNumRel = 0
    ReDim Relac(nNumRel)
    sSQL = "Select PR.bPersRelacAMP, PR.nPersRelacBenefPorc, PR.bPersRelacBenef, C3.cConsDescripcion cConsDescripcion3, C2.cConsDescripcion cConsDescripcion2,C2.nConsValor nConsValor2, C.cConsDescripcion,PR.cPersRelacPersCod,P.cPersNombre,PR.nPersRelac,PR.cMovNro "
    sSQL = sSQL & " from " & gConsPersona & "PersRelaciones PR Inner join Persona P ON PR.cPersRelacPersCod = P.cPersCod"
    sSQL = sSQL & " Inner join Constante C on PR.nPersRelac = C.nConsValor"
    sSQL = sSQL & " Left join Constante C2 on PR.bPersRelacBenef = C2.nConsValor"
    sSQL = sSQL & " Left join Constante C3 on PR.bPersRelacAMP = C3.nConsValor"
    sSQL = sSQL & " Where PR.cPersCod ='" & psPersCod & "' and C.nConsCod = " & gPersRelacion & " And C.nConsValor <> " & gPersRelacion & " and C2.nConsCod = " & gPersRelacBenef & " And C2.nConsValor <> " & gPersRelacBenef & " and C3.nConsCod = " & gPersRelacAmp & " And C3.nConsValor <> " & gPersRelacAmp
    Set R = Conn.CargaRecordSet(sSQL)
    Do While Not R.EOF
        nNumRel = nNumRel + 1
        ReDim Preserve Relac(nNumRel)
        Relac(nNumRel - 1).sPersCodOrig = R!cPersRelacPersCod
        Relac(nNumRel - 1).sPersCod = R!cPersRelacPersCod
        Relac(nNumRel - 1).sApellNombres = Trim(R!cpersnombre)
        Relac(nNumRel - 1).sRelacion = R!cConsDescripcion & Space(75) & R!nPersRelac
        Relac(nNumRel - 1).sPersRelacBenef = R!cConsDescripcion2 & Space(75) & R!nConsValor2
        'Relac(nNumRel - 1).sPersRelacEstado = R!cConsDescripcion4 & Space(50) & Format(IIf(IsNull(R!bPersRelacEstado), 0, R!bPersRelacEstado), "#0")
        Relac(nNumRel - 1).nPersRelacBenefPorc = IIf(IsNull(R!nPersRelacBenefPorc), 0, R!nPersRelacBenefPorc)
        Relac(nNumRel - 1).sPersRelacAMP = R!cConsDescripcion3 & Space(75) & Format(IIf(IsNull(R!bPersRelacAMP), 0, R!bPersRelacAMP), "#0")
        Relac(nNumRel - 1).nTipoAct = 0
        Relac(nNumRel - 1).sMovAct = R!cMovNro
        R.MoveNext
    Loop
    R.Close
End Sub

Private Sub RecuperaRefComercial(ByVal psPersCod As String)
Dim sSQL As String
Dim R As Recordset
nNumRefCom = 0
    
    ReDim RefCom(nNumRefCom)

    sSQL = " SELECT nNumRefCom, cNomRefCom, nTipoRelaCom, cConsDescripcion, cFonoRefCom "
    sSQL = sSQL & " FROM PERSREFCOMERCIAL RC INNER JOIN CONSTANTE C ON RC.nTipoRelaCom = C.nConsValor "
    sSQL = sSQL & " WHERE nConsCod = 3028 AND cPersCod = '" & psPersCod & "'"
    sSQL = sSQL & " ORDER BY nNumRefCom"
    Set R = Conn.CargaRecordSet(sSQL)
    
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

Private Function RecuperaRefComercialPersFte(ByVal psPersFteCod As String) As Variant
Dim sSQL As String
Dim R As Recordset
Dim nPunt As Integer
Dim MatDatos() As String

    sSQL = " SELECT nNumRefCom, cNomRefCom, nTipoRelaCom, cConsDescripcion, cFonoRefCom "
    sSQL = sSQL & " FROM PERSREFCOMERCIAL RC INNER JOIN CONSTANTE C ON RC.nTipoRelaCom = C.nConsValor "
    sSQL = sSQL & " WHERE nConsCod = 3028 AND cPersCod = '" & psPersFteCod & "'"
    sSQL = sSQL & " ORDER BY nNumRefCom"
    Set R = Conn.CargaRecordSet(sSQL)

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
Dim sSQL As String
Dim R As Recordset

    sSQL = "SELECT MAX(nNumRefCom) as nMaxRefCom FROM PERSREFCOMERCIAL WHERE cPersCod = '" & psPersCod & "'"
    Set R = Conn.CargaRecordSet(sSQL)
    
    If Not R.EOF And Not R.BOF Then
        nMaxRefCom = IIf(IsNull(R!nMaxRefCom), 0, R!nMaxRefCom)
    Else
        nMaxRefCom = 0
    End If

    Set R = Nothing
    
End Sub

Private Sub RecuperaPatVehicular(ByVal psPersCod As String)
Dim sSQL As String
Dim R As Recordset
nNumPatVeh = 0
    
    ReDim PatVeh(nNumPatVeh)

    sSQL = " SELECT nNumPatri, nCondPatri, cMarca, cFecFabrica, nValorCom, cConsDescripcion "
    sSQL = sSQL & " FROM PERSPATRIMONIO P INNER JOIN CONSTANTE C ON P.nCondPatri = C.nConsValor "
    sSQL = sSQL & " WHERE nConsCod = 3029 AND cPersCod = '" & psPersCod & "'"
    sSQL = sSQL & " ORDER BY nNumPatri"
    Set R = Conn.CargaRecordSet(sSQL)
    
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
Dim sSQL As String
Dim R As Recordset

    sSQL = "SELECT MAX(nNumPatri) as nMaxPatVeh FROM PERSPATRIMONIO WHERE cPersCod = '" & psPersCod & "'"
    Set R = Conn.CargaRecordSet(sSQL)
    
    If Not R.EOF And Not R.BOF Then
        nMaxPatVeh = IIf(IsNull(R!nMaxPatVeh), 0, R!nMaxPatVeh)
    Else
        nMaxPatVeh = 0
    End If

    Set R = Nothing
    
End Sub

Private Sub RecuperaRefBancaria(ByVal psPersCod As String)
Dim sSQL As String
Dim R As Recordset
nNumRefBan = 0
    
    ReDim RefBan(nNumRefBan)

    sSQL = "SELECT I.cIFTpo + '.' + P.cPersCod cCodBanco, CONVERT(CHAR(30),P.cPersNombre ) AS cBanNombre, "
    sSQL = sSQL & "cNumCta , cNumTarjeta, nMontoLinCred "
    sSQL = sSQL & " FROM PERSREFBANCARIA RB INNER JOIN PERSONA P ON RB.cCodBanco = P.cPersCod "
    sSQL = sSQL & " INNER JOIN INSTITUCIONFINANC I ON P.cPersCod = I.cPersCod "
    sSQL = sSQL & " WHERE RB.cPersCod = '" & psPersCod & "'"
    Set R = Conn.CargaRecordSet(sSQL)
    
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

Private Function RecuperaRefBancariaPersFte(ByVal psPersFteCod As String) As Variant
Dim sSQL As String
Dim R As Recordset
Dim nPunt As Integer
Dim MatDatos() As String

    sSQL = "SELECT I.cIFTpo + '.' + P.cPersCod cCodBanco, CONVERT(CHAR(30),P.cPersNombre ) AS cBanNombre, "
    sSQL = sSQL & "cNumCta , cNumTarjeta, nMontoLinCred "
    sSQL = sSQL & " FROM PERSREFBANCARIA RB INNER JOIN PERSONA P ON RB.cCodBanco = P.cPersCod "
    sSQL = sSQL & " INNER JOIN INSTITUCIONFINANC I ON P.cPersCod = I.cPersCod "
    sSQL = sSQL & " WHERE RB.cPersCod = '" & psPersFteCod & "'"
    Set R = Conn.CargaRecordSet(sSQL)

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

Public Function FuenteIngresoAsignadaACredito(ByVal pnIndex As Integer, ByVal pdFecEval As Date) As Boolean
Dim sSQL As String
Dim R As ADODB.Recordset
    
    sSQL = "Select * from ColocFteIngreso where cNumFuente = '" & FtesIngreso(pnIndex).cNumFte & "' AND dPersEval = '" & Format(pdFecEval, "mm/dd/yyyy") & "'"
    Set R = Conn.CargaRecordSet(sSQL)
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
Dim sSQL As String

Dim R As ADODB.Recordset
    nNumFtesDep = 0
    nNumFtesIndep = 0
    
    If pnTipoFte = gPersFteIngresoTipoDependiente Then
        sSQL = "Select PF.dPersEval, PF.nPersIngCli, PF.nPersGastoFam, PF.nPersIngCon, PF.nPersOtrIng "
        sSQL = sSQL & " From PersFIDependiente PF Where PF.cNumFte = '" & psNumFte & "' Order by dPersEval Asc"
        Set R = Conn.CargaRecordSet(sSQL)
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
        sSQL = "Select PF.dPersEval, PF.cNumFuente, PF.dPersEval, PF.nPersFIActivoDisp, "
        sSQL = sSQL & " PF.nPersFICtasxCobrar, PF.nPersFIInventarios, PF.nPersFIActivosFijos, "
        sSQL = sSQL & " PF.nPersFIProveedores, PF.nPersFICredCMACT, PF.nPersFICredOtros, "
        sSQL = sSQL & " PF.nPersFIVentas, PF.nPersFIPatrimonio, PF.nPersFICostoVentas, PF.nPersFIRecupCtasXCobrar,"
        sSQL = sSQL & " PF.nPersFIEgresosOtros, PF.nPersIngFam, PF.nPersEgrFam, PF.cUltimaActualizacion "
        sSQL = sSQL & " From PersFIIndependiente PF Where PF.cNumFte = '" & psNumFte & "' Order by dPersEval Asc"
        
        Set R = Conn.CargaRecordSet(sSQL)
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
Dim sSQL As String
Dim R As ADODB.Recordset
Dim bConexInterna As Boolean
    
    If pcCtaCod = "" Then
    
        sSQL = "Select PF.bCostoProd, PF.cNumFuente, PF.cPersCod, PF.cPersFIPersCod, PF.cPersFICargo, "
        sSQL = sSQL & " PF.dPersFIinicio, PF.cPersFIMoneda, PF.cPersFIComentario, "
        sSQL = sSQL & " PF.cRazSocUbiGeo, PF.cRazSocDirecc, PF.cRazSocTelef, PF.cRazSocDescrip, "
        sSQL = sSQL & " PF.nPersTipFte, PF.cUltimaActualizacion, "
    '------------------------------------------------------------------------------------------------------------
        sSQL = sSQL & "P.cPersDireccUbiGeo,P.cPersDireccDomicilio,P.cPersDireccCondicion, P.cPersTelefono, "
        sSQL = sSQL & "P.nPersPersoneria,P.cPersCIIU, P.cPersTelefono2, P.cPersCod as cPersFteCod,P.cPersNombre,P.dPersNacCreac, "
        sSQL = sSQL & "PJ.cPersJurTpo, PJ.cPersJurMagnitud,PJ.cPersJurSigla,PJ.nPersJurEmpleados, CI.cCIIUdescripcion, "
        sSQL = sSQL & "cNroRUC = (  Select ISNULL(cPersIDnro,'') From PersID WHERE PersID.cPersCod = P.cPersCod and PersID.cPersIDTpo = " & gPersIdRUC & " ), "
        sSQL = sSQL & "PJT.cPersJurTpoDesc, CN7.cConsDescripcion as cPersDireccCondicionDes, CN8.cConsDescripcion as cPersJurMagnitudDes, "
        sSQL = sSQL & " Dpto = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '1' And Substring(u.cUbigeoCod,2,2) = Substring(P.cPersDireccUbiGeo,2,2) And Substring(u.cUbigeoCod,4,9) = '000000000' ), '' ) , "
        sSQL = sSQL & " Prov = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '2' And Substring(u.cUbigeoCod,2,4) = Substring(P.cPersDireccUbiGeo,2,4) And Substring(u.cUbigeoCod,6,7) = '0000000' ), '' ) , "
        sSQL = sSQL & " Dist = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '3' And Substring(u.cUbigeoCod,2,6) = Substring(P.cPersDireccUbiGeo,2,6) And Substring(u.cUbigeoCod,8,5) = '00000' ), '' ), "
        sSQL = sSQL & " Zona = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where u.cUbigeoCod = P.cPersDireccUbiGeo ),'') "
        sSQL = sSQL & " From  PersFteIngreso PF "
        sSQL = sSQL & " INNER JOIN " & gConsPersona & "Persona P ON PF.cPersFIPersCod = P.cPersCod "
        sSQL = sSQL & " left join " & gConsPersona & "PersonaNat PN ON P.cPersCod=PN.cPersCod "
        sSQL = sSQL & " LEFT JOIN " & gConsPersona & "PersonaJur PJ ON P.cPersCod = PJ.cPersCod "
        sSQL = sSQL & " LEFT JOIN CIIU CI ON P.cPersCIIU = CI.cCIIUcod "
        sSQL = sSQL & " LEFT JOIN PersJurTpo PJT ON PJ.cPersJurTpo = PJT.cPersJurTpoCod "
        sSQL = sSQL & " Left Join Constante CN7 ON CN7.nConsValor = P.cPersDireccCondicion AND CN7.nConsCod = 1029 "
        sSQL = sSQL & " Left Join Constante CN8 ON CN8.nConsValor = PJ.cPersJurMagnitud AND CN8.nConsCod = 1004 "
        sSQL = sSQL & " Where PF.cPersCod = '" & psCodPers & "'"
        sSQL = sSQL & " and (exists(Select cnumfuente from persfidependiente where cnumfuente=PF.cNumfuente) or Exists(Select cnumfuente from persfiindependiente where cnumfuente=PF.cnumfuente))"
        sSQL = sSQL & " Order By dPersFIinicio desc"
    '----------------------------------------------------------------------------------------------------------
   Else
         sSQL = "Select PF.bCostoProd, PF.cNumFuente, PF.cPersCod, PF.cPersFIPersCod, PF.cPersFICargo, "
        sSQL = sSQL & " PF.dPersFIinicio, PF.cPersFIMoneda, PF.cPersFIComentario, "
        sSQL = sSQL & " PF.cRazSocUbiGeo, PF.cRazSocDirecc, PF.cRazSocTelef, PF.cRazSocDescrip, "
        sSQL = sSQL & " PF.nPersTipFte, PF.cUltimaActualizacion, "
    '------------------------------------------------------------------------------------------------------------
        sSQL = sSQL & "P.cPersDireccUbiGeo,P.cPersDireccDomicilio,P.cPersDireccCondicion, P.cPersTelefono, "
        sSQL = sSQL & "P.nPersPersoneria,P.cPersCIIU, P.cPersTelefono2, P.cPersCod as cPersFteCod,P.cPersNombre,P.dPersNacCreac, "
        sSQL = sSQL & "PJ.cPersJurTpo, PJ.cPersJurMagnitud,PJ.cPersJurSigla,PJ.nPersJurEmpleados, CI.cCIIUdescripcion, "
        sSQL = sSQL & "cNroRUC = (  Select ISNULL(cPersIDnro,'') From PersID WHERE PersID.cPersCod = P.cPersCod and PersID.cPersIDTpo = " & gPersIdRUC & " ), "
        sSQL = sSQL & "PJT.cPersJurTpoDesc, CN7.cConsDescripcion as cPersDireccCondicionDes, CN8.cConsDescripcion as cPersJurMagnitudDes, "
        sSQL = sSQL & " Dpto = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '1' And Substring(u.cUbigeoCod,2,2) = Substring(P.cPersDireccUbiGeo,2,2) And Substring(u.cUbigeoCod,4,9) = '000000000' ), '' ) , "
        sSQL = sSQL & " Prov = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '2' And Substring(u.cUbigeoCod,2,4) = Substring(P.cPersDireccUbiGeo,2,4) And Substring(u.cUbigeoCod,6,7) = '0000000' ), '' ) , "
        sSQL = sSQL & " Dist = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where Substring(u.cUbigeoCod,1,1) = '3' And Substring(u.cUbigeoCod,2,6) = Substring(P.cPersDireccUbiGeo,2,6) And Substring(u.cUbigeoCod,8,5) = '00000' ), '' ), "
        sSQL = sSQL & " Zona = IsNull((Select u.cUbiGeoDescripcion From UbicacionGeografica u where u.cUbigeoCod = P.cPersDireccUbiGeo ),'') "
        sSQL = sSQL & " From  PersFteIngreso PF "
        sSQL = sSQL & " INNER JOIN " & gConsPersona & "Persona P ON PF.cPersFIPersCod = P.cPersCod "
        sSQL = sSQL & " left join " & gConsPersona & "PersonaNat PN ON P.cPersCod=PN.cPersCod "
        sSQL = sSQL & " LEFT JOIN " & gConsPersona & "PersonaJur PJ ON P.cPersCod = PJ.cPersCod "
        sSQL = sSQL & " LEFT JOIN CIIU CI ON P.cPersCIIU = CI.cCIIUcod "
        sSQL = sSQL & " LEFT JOIN PersJurTpo PJT ON PJ.cPersJurTpo = PJT.cPersJurTpoCod "
        sSQL = sSQL & " Left Join Constante CN7 ON CN7.nConsValor = P.cPersDireccCondicion AND CN7.nConsCod = 1029 "
        sSQL = sSQL & " Left Join Constante CN8 ON CN8.nConsValor = PJ.cPersJurMagnitud AND CN8.nConsCod = 1004 "
        sSQL = sSQL & " Inner Join ColocFteIngreso CFI on CFI.cNumFuente=PF.cNumFuente"
        sSQL = sSQL & " Where PF.cPersCod = '" & psCodPers & "'"
        sSQL = sSQL & " and (exists(Select cnumfuente from persfidependiente where cnumfuente=PF.cNumfuente) or Exists(Select cnumfuente from persfiindependiente where cnumfuente=PF.cnumfuente))"
        sSQL = sSQL & " and CFI.cCtaCod='" & pcCtaCod & "'"
        sSQL = sSQL & " Order By dPersFIinicio desc"
    '----------------------------------------------------------------------------------------------------------
   End If
    
    nNumFtes = 0
    bConexInterna = False
    ReDim FtesIngreso(nNumFtes)
    If Conn Is Nothing Then
        bConexInterna = True
        Set Conn = New DConecta
        Conn.AbreConexion
    End If
    Set R = Conn.CargaRecordSet(sSQL)
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
Dim sSQL As String
Dim R As ADODB.Recordset

    FtesIngreso(pnNumFte).nNumFtesDep = 0
    sSQL = " Select PF.cNumFuente, PF.dPersEval, PF.nPersIngCli, PF.nPersGastoFam, PF.nPersIngCon, PF.nPersOtrIng, PF.dPersFICaduca "
    sSQL = sSQL & " FROM PersFIDependiente PF Where PF.cNumFuente = '" & FtesIngreso(pnNumFte).cNumFte & "'"
    sSQL = sSQL & " ORDER BY PF.dPersEval asc"
    Set R = Conn.CargaRecordSet(sSQL)
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
Dim sSQL As String
Dim R As ADODB.Recordset

    
    sSQL = " Select C.dFecEval, CN2.cConsDescripcion as sUniProd, CN1.cConsDescripcion as sTipoCultivo,  C.cNumFuente, C.nTpoCultivo, C.nMaquinaria, C.nJornales, C.nInsumos, C.nPesticidas, C.nOtros, C.nHectareas,C.nMantenimiento, "
    sSQL = sSQL & " C.nProduccion , C.nUniProd, C.nPreUni,C.nSiembra,C.nCosecha,C.nDesaAgricola,C.nPlanOtros "
    sSQL = sSQL & "  From PersFICostoProd C Inner Join Constante CN1 ON CN1.nConsCod = 1046 AND C.nTpoCultivo = CN1.nConsValor "
    sSQL = sSQL & "                 Inner Join Constante CN2 ON CN2.nConsCod = 1045 AND C.nUniProd = CN2.nConsValor "
    sSQL = sSQL & "  Where cNumFuente = '" & FtesIngreso(pnNumFte).cNumFte & "' "
    Set R = Conn.CargaRecordSet(sSQL)
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
Dim sSQL As String
Dim R As ADODB.Recordset

    FtesIngreso(pnNumFte).nNumFtesIndep = 0
    sSQL = "SELECT  PF.cNumFuente, PF.dPersEval, PF.dPersFICaduca, PF.nPersFIActivoDisp, "
    sSQL = sSQL & " PF.nPersFICtasxCobrar, PF.nPersFIInventarios, PF.nPersFIActivosFijos, "
    sSQL = sSQL & " PF.nPersFIProveedores, PF.nPersFICredCMACT, PF.nPersFICredOtros, "
    sSQL = sSQL & " PF.nPersFIVentas, PF.nPersFIPatrimonio, PF.nPersFICostoVentas, "
    sSQL = sSQL & " PF.nPersFIRecupCtasXCobrar, PF.nPersFIEgresosOtros, PF.nPersIngFam, "
    sSQL = sSQL & " PF.nPersEgrFam, PF.cUltimaActualizacion "
    sSQL = sSQL & " FROM PersFIIndependiente PF Where PF.cNumFuente = '" & FtesIngreso(pnNumFte).cNumFte & "'"
    sSQL = sSQL & " ORDER BY PF.dPersEval asc"
        
    Set R = Conn.CargaRecordSet(sSQL)
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
Dim sSQL As String
Dim R As ADODB.Recordset
    
    Set Conn = New DConecta
    Conn.AbreConexion
    Set ConexionActiva = Conn.ConexionActiva
    
    '***** CARGA DATOS DE LA PERSONA ******
    sSQL = "Select P.nPersRelaInst, P.cPersCodSbs, P.cPersEmail cEmail,P.cPersTelefono2,P.cPersCod,P.cPersNombre,P.dPersNacCreac, "
    sSQL = sSQL & "P.cPersDireccUbiGeo,P.cPersDireccDomicilio,P.cPersDireccCondicion, P.nPersValComDomicilio, P.cPersTelefono, "
    sSQL = sSQL & "P.nPersPersoneria,P.cPersCIIU,P.cPersEstado,P.cUltimaActualizacion, "
    sSQL = sSQL & "PN.cPersNatSexo,PN.nPersNatEstCiv,PN.nPersNatHijos,PN.nPersPeso,PN.nPersTalla,PN.nPersTpoSangre, PN.cNacionalidad, PN.nResidente, "
    sSQL = sSQL & "PJ.cPersJurTpo, PJ.cPersJurMagnitud,PJ.cPersJurSigla,PJ.nPersJurEmpleados, CI.cCIIUdescripcion"
    sSQL = sSQL & " FROM " & gConsPersona & "Persona P left join " & gConsPersona & "PersonaNat PN ON P.cPersCod=PN.cPersCod "
    sSQL = sSQL & " LEFT JOIN " & gConsPersona & "PersonaJur PJ ON P.cPersCod = PJ.cPersCod "
    sSQL = sSQL & " LEFT JOIN CIIU CI ON P.cPersCIIU = CI.cCIIUcod "
    sSQL = sSQL & " WHERE P.cPersCod = '" & psCodPers & "'"
    
    Set R = Conn.CargaRecordSet(sSQL)
    If Not R.BOF And Not R.EOF Then
        sPersCod = R!cPersCod
        sNombreCompleto = Trim(R!cpersnombre)
        sApePat = BuscaNombre(R!cpersnombre, BusqApellidoPaterno)
        sApeMat = BuscaNombre(R!cpersnombre, BusqApellidoMaterno)
        sNombres = BuscaNombre(R!cpersnombre, BusqNombres)
        sApeCas = BuscaNombre(R!cpersnombre, BusqApellidoCasada)
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
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    On Error GoTo ErrHandler
        sSQL = "Select Count(*)  as nValor From Persona Where cPersCod='" & pcPersCod & "' AND bPreferencial=1"
        
        Set oConec = New DConecta
        oConec.AbreConexion
        Set rs = oConec.CargaRecordSet(sSQL)
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

