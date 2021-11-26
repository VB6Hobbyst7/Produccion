VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UPersona_Cli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           UPersona_Cli
'***     Descripcion:       Mantenimiento de Personas
'                           manejado en la Maquina Cliente
'***     Creado por:        ARCV
'***     Maquina:           07DESARROLLO03
'***     Fecha-Tiempo:      07/12/2005 06:56:14 PM
'***     Ultima Modificacion: Creacion
'*****************************************************************************************

Option Explicit

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
'*** PEAC 20080412
Private nNumEmp As Integer
Private dFechaNac As Date
Private cFechaAct As String '*** PEAC 20080731

Private dFechaFallec As Date 'MAVM 03042009

Private bBIEsHabilitado As Boolean 'ANDE 20190910 Si cliente está habilitado para usar Caja Maynas Online

'MADM 20091114
'Private sCentroTrabajo As String
'Private sCargo As String
Private sCredito As Integer
Private sAhorro As Integer
Private sServicio As Integer
Private sOtro As Integer
Private sPeps As Integer
'END MADM
'EJVG20120813 ***
Private nSujetoObligadoIni As Integer
Private nSujetoObligado As Integer
Private nOfCumplimiento As Integer
'END EJVG *******
Private nAutorizaUsoDatos As Integer 'FRHU 20151130 ERS077-2015
Private nAutorizaUsoDatosIni As Integer 'FRHU 20151130 ERS077-2015
'madm 20100503
Private sActiGiro1 As String
'end madm

Private nPersCargo As String '** Juez 20120326

Private sUbicGeografica As String
Private sDomicilio As String
Private sRefDomicilio As String '*** PEAC 20080801
Private sCondicionDomic As String
Private nValComDomicilio As Currency
'JUEZ 20131007 ***********************
Private sUbiGeoNegocio As String
Private sNegocioDireccion As String
Private sRefNegocio As String

'marg 11-05-2016
Private sNombreCentroLaboral As String

Private nRemisionInfoEmail As Integer
'END JUEZ ****************************
Private sTelefonos As String
Private sTelefonos2 As String
'EJVG20110712************************
Private sCelular As String
Private sCelular2 As String
Private sCelular3 As String
'************************************
Private sCiiu As String
Private sCiiuDescripcion As String

'*** PEAC 20080412
Private dFechaInscRuc As Date
Private dFechaIniActi As Date
Private nTipoCompDescripcion As String
Private nTipoSistInforDescripcion As String
Private nMonedaPatri As String
Private nCadenaProdDescripcion As String

'*** FIN PEAC
Private sEstado As String
Private nTalla As Double
Private nPeso As Double
Private sTipoSangre As String
Private nTipoPersona As Integer 'Personeria
Private sEmail As String
Private sEmail2 As String 'EJVG20111207
'*** PEAC 20080412
Private nNumDependi As Integer
Private sActComple As String
Private nNumPtosVta As Integer
Private sActiGiro As String
'*** FIN PEAC
Private sActualiza As String
Private sPersCodSbs As String
Private nPersRelInst As Integer
Public dFechaHoy As Date

Public RFirma As ADODB.Recordset

'Datos Brechas
Private nResidencia As Integer
Private cPaisReside As String '** Juez 20120327
Private cNacionalidad As String
Private cMotivoActu As String '***PEAC 20080801
Private nPersRRHH As Integer 'MADM 20101020

'Datos de Persona Juridica
Private sRazonSocial As String
Private sSiglas As String
Private sTipoPersJur As String
Private nNumEmplead As Integer
Private sMagnitudEmp As String
Private sObjetoSocial As String '** Juez 20120328
Private sMagnitudPersNat As String 'JACA 20110427

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
    sRefComComentario As String
    sRefComDireccion As String
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
    sPatVehModelo As String
    sPatVehFecFab As String
    sPatVehPlaca As String
    nPatVehValCom As Currency
    sPatVehCondicion As String
    nTipoAct As TPersCambios
    
End Type
Private PatVeh() As TPatVeh

'madm 20100326
Private Type TVisitas
    sPersCodOrig As String
    sPersCod As String
    sApellNombres As String
    sDireccion As String
    dFecVis As Date
    sVisObserva As String
    sUsual As String
    nTipoAct As TPersCambios
End Type

Private Visita() As TVisitas
Private nNumVisitas As Integer
'end madm

'MAVM 201000607 BAS II ***
Private Type TVentas
    sPersCodOrigVA As String
    sPersCod As String
    sApellNombres As String
    sPeriodo As String
    nMonto As Currency
    dFecAud As Date
    nTipoAct As TPersCambios
End Type

Private Ventas() As TVentas
Private nNumVentas As Integer
'***

'*** PEAC 20080412
Private Type TPatOtros
    nPatOtrosCod As Integer
    sPatOtrosDescripcion As String
    nPatOtrosValCom As Currency
    nTipoAct As TPersCambios
End Type
Private PatOtros() As TPatOtros


'*** PEAC 20080412
Private Type TPatInmuebles
    nPatInmueblesCod As Integer
    sPatInmueblesUbicacion As String
    nPatInmueblesAreaTerreno As Double
    nPatInmueblesAreaConstru As Double
    sPatInmueblesTipoUso As String
    sPatInmueblesRRPP As String
    nPatInmueblesValCom As Currency
    nTipoAct As TPersCambios
End Type
Private PatInmuebles() As TPatInmuebles

Private nNumPatVeh As Integer
Private nMaxPatVeh As Integer

'*** PEAC 20080412
Private nNumPatOtros As Integer
Private nMaxPatOtros As Integer
Private nNumPatInmuebles As Integer
Private nMaxPatInmuebles As Integer

'Tipo de Datos para Almacenar las Ftes de Ingreso
'peac 20071227 se agrego "dFecEEFF As Date"
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

'peac 20071227 se agrego "dfecEEFF As Date"
Private Type TFteIngInd
    cNumFte As String
    dFecEval As Date
    dfecEEFF As Date
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
    RFteIng As ADODB.Recordset
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
    nFtePaRscInd As Integer
    
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
    rsEstFinan As ADODB.Recordset 'WIOR 20140321
End Type

Private FtesIngreso() As TFteIng

Private FtesIngresoTmp As TFteIng
Private nNumFtes As Integer
Dim ConexDitrib As ADODB.Connection
Dim bPersDistrib As Boolean
Dim nNumFtesDep As Integer
Dim nNumFtesIndep As Integer

'CUSCO
Private nIngresoProm As Currency

'Recordset para el Grabado de Datos

Dim rsPersona As ADODB.Recordset
Dim rsVisita As ADODB.Recordset
Dim rsVentas As ADODB.Recordset 'MAVM 20100607 BAS II
Dim rsRelac As ADODB.Recordset
Dim rsRefCom As ADODB.Recordset
Dim rsRefBan As ADODB.Recordset
Dim rsPatVeh As ADODB.Recordset

'*** PEAC 20080412
Dim rsPatOtros As ADODB.Recordset
Dim rsPatInmuebles As ADODB.Recordset

Dim rsDocs As ADODB.Recordset
Dim rsFteIng As ADODB.Recordset
Dim rsFIDep As ADODB.Recordset
Dim rsFIInd As ADODB.Recordset
Dim rsFICosto As ADODB.Recordset
Dim rsHojaEval As ADODB.Recordset
Dim nIntValor As Integer
'*** CTI3 07092018
Dim rsDatosAdc As ADODB.Recordset

'*** PEAC 20080402
'Dim rsHojEval As ADODB.Recordset

Dim i As Integer
Dim j As Integer
Private rsEstFinan As ADODB.Recordset 'WIOR 20140321

'**** CTI3 Tipo de dato para Almacenar el Detalle de las Personas Juridicas
'y Patrimonio en Otras Empresas
Private Type TDatosJuridicos
    sNombrejur As String
    sRUC As String
    sCi As String
    sNacionalidad As String
    sCargo As String
    sRazonSocial As String
    nApoSol As Double
    nApoPorc As Double
    sCodSbs As String
    sTipo As Integer
End Type
Private TJuridicoJur() As TDatosJuridicos
Private RJuridicoJur() As TDatosJuridicos

Dim nFilaGri1, nFilaGri2, nFilaGri3, nFilaGri4, nFilaGri5 As Integer
Dim nValor As Integer

'************
Sub LlenaDetallesJuridicos(gAccio As FlexEdit, gDirec As FlexEdit, gGere As FlexEdit, _
gPatri As FlexEdit, gPatEmp As FlexEdit, gCargos As FlexEdit, ByVal nTpoPer As Integer)

Dim i, nTotaFil, nFila1, nFila2, nFila3, nFila4, nFila5, nfila6, nFilMat As Integer
nFilMat = 0

'If nTpoPer = 1 Then
'
'    nFila5 = gPatEmp.rows - 1
'    nTotaFil = nFila5
'    ReDim Preserve TJuridicoJur(nTotaFil)
'
'     For i = 1 To nFila5
'        TJuridicoJur(nFilMat).sNombrejur = gPatEmp.TextMatrix(i, 3)
'        TJuridicoJur(nFilMat).sRUC = gPatEmp.TextMatrix(i, 5)
'        TJuridicoJur(nFilMat).sCi = ""
'        TJuridicoJur(nFilMat).sCargo = ""
'        TJuridicoJur(nFilMat).sNacionalidad = ""
'        TJuridicoJur(nFilMat).sRazonSocial = gPatEmp.TextMatrix(i, 4)
'        TJuridicoJur(nFilMat).nApoSol = IIf(gPatEmp.TextMatrix(i, 6) <> "", gPatEmp.TextMatrix(i, 6), 0)
'        TJuridicoJur(nFilMat).nApoPorc = IIf(gPatEmp.TextMatrix(i, 7) <> "", gPatEmp.TextMatrix(i, 7), 0)
'        TJuridicoJur(nFilMat).sCodSbs = gPatEmp.TextMatrix(i, 2)
'        TJuridicoJur(nFilMat).sTipo = 5
'        nFilMat = nFilMat + 1
'     Next i
'
'Else
    If nTpoPer = 2 Or nTpoPer = 1 Then
          
        nFila1 = gAccio.Rows - 1: nFila2 = gDirec.Rows - 1: nFila3 = gGere.Rows - 1: nFila4 = gPatri.Rows - 1: nFila5 = gCargos.Rows - 1: nfila6 = gPatEmp.Rows - 1
        nTotaFil = nFila1 + nFila2 + nFila3 + nFila4 + nFila5 + nfila6
        
        ReDim Preserve TJuridicoJur(nTotaFil)
              
              
          For i = 1 To nFila1
            If Len(gAccio.TextMatrix(i, 1)) > 0 Then
              TJuridicoJur(nFilMat).sNombrejur = gAccio.TextMatrix(i, 1)
              TJuridicoJur(nFilMat).sRUC = gAccio.TextMatrix(i, 2)
              TJuridicoJur(nFilMat).sCi = gAccio.TextMatrix(i, 3)
              TJuridicoJur(nFilMat).sNacionalidad = gAccio.TextMatrix(i, 4)
              TJuridicoJur(nFilMat).sCargo = ""
              TJuridicoJur(nFilMat).sRazonSocial = ""
              TJuridicoJur(nFilMat).nApoSol = IIf(gAccio.TextMatrix(i, 5) <> "", gAccio.TextMatrix(i, 5), 0)
              TJuridicoJur(nFilMat).nApoPorc = IIf(gAccio.TextMatrix(i, 6) <> "", gAccio.TextMatrix(i, 6), 0)
              TJuridicoJur(nFilMat).sCodSbs = ""
              TJuridicoJur(nFilMat).sTipo = 1
              nFilMat = nFilMat + 1
            End If
        Next i
      '  nFilMat = 0
        For i = 1 To nFila2
            If Len(gDirec.TextMatrix(i, 1)) > 0 Then
              TJuridicoJur(nFilMat).sNombrejur = gDirec.TextMatrix(i, 1)
              TJuridicoJur(nFilMat).sRUC = gDirec.TextMatrix(i, 2)
              TJuridicoJur(nFilMat).sCi = gDirec.TextMatrix(i, 3)
              TJuridicoJur(nFilMat).sCargo = gDirec.TextMatrix(i, 4)
              TJuridicoJur(nFilMat).sNacionalidad = gDirec.TextMatrix(i, 5)
              TJuridicoJur(nFilMat).sRazonSocial = ""
              TJuridicoJur(nFilMat).nApoSol = 0
              TJuridicoJur(nFilMat).nApoPorc = 0
              TJuridicoJur(nFilMat).sCodSbs = ""
              TJuridicoJur(nFilMat).sTipo = 2
              nFilMat = nFilMat + 1
            End If
        Next i
        For i = 1 To nFila3
           If Len(gGere.TextMatrix(i, 1)) > 0 Then
              TJuridicoJur(nFilMat).sNombrejur = gGere.TextMatrix(i, 1)
              TJuridicoJur(nFilMat).sRUC = gGere.TextMatrix(i, 2)
              TJuridicoJur(nFilMat).sCi = gGere.TextMatrix(i, 3)
              TJuridicoJur(nFilMat).sCargo = gGere.TextMatrix(i, 4)
              TJuridicoJur(nFilMat).sNacionalidad = ""
              TJuridicoJur(nFilMat).sRazonSocial = ""
              TJuridicoJur(nFilMat).nApoSol = 0
              TJuridicoJur(nFilMat).nApoPorc = 0
              TJuridicoJur(nFilMat).sCodSbs = ""
              TJuridicoJur(nFilMat).sTipo = 3
              nFilMat = nFilMat + 1
           End If
        Next i
        For i = 1 To nFila4
            If Len(gPatri.TextMatrix(i, 1)) > 0 Then
              TJuridicoJur(nFilMat).sNombrejur = gPatri.TextMatrix(i, 3)
              TJuridicoJur(nFilMat).sRUC = gPatri.TextMatrix(i, 5)
              TJuridicoJur(nFilMat).sCi = ""
              TJuridicoJur(nFilMat).sCargo = ""
              TJuridicoJur(nFilMat).sNacionalidad = ""
              TJuridicoJur(nFilMat).sRazonSocial = gPatri.TextMatrix(i, 4)
              TJuridicoJur(nFilMat).nApoSol = IIf(gPatri.TextMatrix(i, 6) <> "", gPatri.TextMatrix(i, 6), 0)
              TJuridicoJur(nFilMat).nApoPorc = IIf(gPatri.TextMatrix(i, 7) <> "", gPatri.TextMatrix(i, 7), 0)
              TJuridicoJur(nFilMat).sCodSbs = gPatri.TextMatrix(i, 2)
              TJuridicoJur(nFilMat).sTipo = 4
              nFilMat = nFilMat + 1
            End If
        Next i
        For i = 1 To nFila5
            If Len(gCargos.TextMatrix(i, 1)) > 0 Then
              TJuridicoJur(nFilMat).sNombrejur = gCargos.TextMatrix(i, 3)
              TJuridicoJur(nFilMat).sRUC = gCargos.TextMatrix(i, 5)
              TJuridicoJur(nFilMat).sCi = ""
              TJuridicoJur(nFilMat).sCargo = gCargos.TextMatrix(i, 6)
              TJuridicoJur(nFilMat).sNacionalidad = ""
              TJuridicoJur(nFilMat).sRazonSocial = gCargos.TextMatrix(i, 4)
              TJuridicoJur(nFilMat).nApoSol = 0
              TJuridicoJur(nFilMat).nApoPorc = 0
              TJuridicoJur(nFilMat).sCodSbs = gCargos.TextMatrix(i, 2)
              TJuridicoJur(nFilMat).sTipo = 6
              nFilMat = nFilMat + 1
            End If
            
        Next i
        For i = 1 To nfila6
            If Len(gPatEmp.TextMatrix(i, 1)) > 0 Then
                TJuridicoJur(nFilMat).sNombrejur = gPatEmp.TextMatrix(i, 3)
                TJuridicoJur(nFilMat).sRUC = gPatEmp.TextMatrix(i, 5)
                TJuridicoJur(nFilMat).sCi = ""
                TJuridicoJur(nFilMat).sCargo = ""
                TJuridicoJur(nFilMat).sNacionalidad = ""
                TJuridicoJur(nFilMat).sRazonSocial = gPatEmp.TextMatrix(i, 4)
                TJuridicoJur(nFilMat).nApoSol = IIf(gPatEmp.TextMatrix(i, 6) <> "", gPatEmp.TextMatrix(i, 6), 0)
                TJuridicoJur(nFilMat).nApoPorc = IIf(gPatEmp.TextMatrix(i, 7) <> "", gPatEmp.TextMatrix(i, 7), 0)
                TJuridicoJur(nFilMat).sCodSbs = gPatEmp.TextMatrix(i, 2)
                TJuridicoJur(nFilMat).sTipo = 5
                nFilMat = nFilMat + 1
            End If
        Next i
    End If
'
'End If

'LlenaDetallesJuridicos = TJuridicoJur
 
End Sub

Sub LlenaRecordSet_DetaJuridicos()

Set rsDatosAdc = New ADODB.Recordset

On Error GoTo matrizvacia
'
With rsDatosAdc
    'Crear RecordSet
    .Fields.Append "sNombrejur", adVarChar, 200
    .Fields.Append "sRUC", adVarChar, 15
    .Fields.Append "sCi", adVarChar, 20
    .Fields.Append "sNacionalidad", adVarChar, 20
    .Fields.Append "sCargo", adVarChar, 100
    .Fields.Append "sRazonSocial", adVarChar, 150
    .Fields.Append "nApoSol", adCurrency
    .Fields.Append "nApoPorc", adDouble
    .Fields.Append "sCodSbs", adVarChar, 11
    .Fields.Append "sTipo", adInteger
    .Open

    'Llenar Recordset
    
    For i = LBound(TJuridicoJur) To UBound(TJuridicoJur) - 1

        If Len(TJuridicoJur(i).sNombrejur) > 0 Then
            .AddNew
            .Fields("sNombrejur") = TJuridicoJur(i).sNombrejur
            .Fields("sRUC") = TJuridicoJur(i).sRUC
            .Fields("sCi") = TJuridicoJur(i).sCi
            .Fields("sNacionalidad") = TJuridicoJur(i).sNacionalidad
            .Fields("sCargo") = TJuridicoJur(i).sCargo
            .Fields("sRazonSocial") = TJuridicoJur(i).sRazonSocial
            .Fields("nApoSol") = TJuridicoJur(i).nApoSol
            .Fields("nApoPorc") = TJuridicoJur(i).nApoPorc
            .Fields("sCodSbs") = TJuridicoJur(i).sCodSbs
            .Fields("sTipo") = TJuridicoJur(i).sTipo
        End If
    Next i

    If Not .EOF Then .MoveFirst
End With
matrizvacia:
err.Clear
End Sub

Public Sub LimpiaEliminados()
Dim i, j As Integer
Dim bExisteEliminados As Boolean
    bExisteEliminados = True
    Do While bExisteEliminados
        bExisteEliminados = False
        For i = 0 To nNumFtes - 1
            If FtesIngreso(i).nTipoAct = 3 Then
                bExisteEliminados = True
                For j = i To nNumFtes - 2
                    FtesIngreso(j) = FtesIngreso(j + 1)
                Next j
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

'Public Function TiempoFteIngreso() As Integer
'Dim sSQL As String
'Dim oConecta As COMConecta.DCOMConecta
'Dim R As ADODB.Recordset
'
'On Error GoTo ErrorTiempoFteIngreso
'    sSQL = "Select nParamValor from " & gConsComunes & "ColocParametro where nParamVar = " & gColocParamFteIngresoTiempoEval
'    Set oConecta = New COMConecta.DCOMConecta
'    Call oConecta.AbreConexion
'    Set R = oConecta.CargaRecordSet(sSQL)
'    If Not R.BOF And Not R.EOF Then
'        TiempoFteIngreso = CInt(Format(R!nParamValor, "#0"))
'    Else
'        TiempoFteIngreso = 1
'    End If
'    R.Close
'    Set R = Nothing
'    oConecta.CierraConexion
'    Set oConecta = Nothing
'    Exit Function
'
'ErrorTiempoFteIngreso:
'    Err.Raise Err.Number, "ReucperaPersona", Err.Description
'End Function

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
    nNumEmp = 0 '*** PEAC 20080412
    sUbicGeografica = ""
    sDomicilio = ""
    sRefDomicilio = "" '*** PEAC 20080801
    sCondicionDomic = ""
    nValComDomicilio = 0
    'JUEZ 20131007 ********************
    sUbiGeoNegocio = ""
    sNegocioDireccion = ""
    sRefNegocio = ""
    sNombreCentroLaboral = "" 'marg
    nRemisionInfoEmail = 0
    'END JUEZ *************************
    sTelefonos = ""
    sTelefonos2 = ""
    'EJVG20110712********************
    sCelular = ""
    sCelular2 = ""
    sCelular3 = ""
    '********************************
    sCiiu = ""
    sEstado = ""

    '*** PEAC 20080412
    nTipoCompDescripcion = ""
    nTipoSistInforDescripcion = ""
    nMonedaPatri = ""
    nCadenaProdDescripcion = ""
    '*** FIN PEAC

    nTalla = 0
    nPeso = 0
    sTipoSangre = ""
    nTipoPersona = gPersonaNat
    sEmail = ""
    sEmail2 = "" 'EJVG20110712
    '*** PEAC 20080412
    nNumDependi = 0
    sActComple = ""
    nNumPtosVta = 0
    sActiGiro = ""
    '*** FIN PEAC
    Set RFirma = Nothing
    sRazonSocial = ""
    sSiglas = ""
    sTipoPersJur = ""
    nNumEmplead = 0
    sMagnitudEmp = ""
    sObjetoSocial = "" '** Juez 20120328
    sMagnitudPersNat = "" 'JACA 20110427
    'EJVG20120813 ***
    nSujetoObligadoIni = -1
    nSujetoObligado = -1
    nOfCumplimiento = -1
    'END EJVG *******
    nAutorizaUsoDatosIni = -1 'FRHU 20151130 ERS077-2015
    nAutorizaUsoDatos = -1 'FRHU 20151130 ERS077-2015
    nTipoAct = PersFilaSinCambios
    nNumDoc = 0
    ReDim Docs(0)
    ReDim Relac(0)
    nNumRel = 0
    ReDim FtesIngreso(0)
    nNumFtes = 0
    Set RFirma = Nothing 'RecuperaFirma("")
    
    'CUSCO
    nIngresoProm = 0
End Sub

Public Function UltimoCorrelativo(ByVal psCodAge As String) As String
Dim oPersonas As New COMDPersona.DCOMPersonas

UltimoCorrelativo = oPersonas.UltimoCorrelativo(psCodAge, bPersDistrib, ConexDitrib)

Set oPersonas = Nothing
End Function

Public Function GrabarCambiosPersona(ByVal psCodCMACT As String, ByVal psCodAge As String, _
                                ByVal pdFecha As Date, psCodUser As String, _
                                ByRef prsVerHomoninia As ADODB.Recordset, _
                                ByRef pnVerDuplicaDoc As Integer, _
                                ByRef pnVerTamanioDoc As Integer, _
                                Optional ByVal pbVerificaHomonimia As Boolean = True, _
                                Optional ByRef psPersCodGrabar As String = "", _
                                Optional ByRef psCodUserResp As String = "") As Boolean  'JUEZ 20131007 psCodUserResp  ***CTI3  pMatriz
 
Dim oPersona As COMDPersona.DCOMPersonas

Dim bDAexiste As Boolean  'cti3 31102018

'PARA DISTRIBUIDO
Dim vDocs(4) As String
    
    vDocs(0) = ""
    vDocs(1) = ""
    vDocs(2) = ""
    vDocs(3) = ""
    
    For i = 0 To UBound(Docs) - 1
        Select Case Trim((Docs(i).nTipo))
            Case "1", "3"
                vDocs(0) = Trim((Docs(i).nTipo))
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
        Select Case Trim((Docs(i).nTipo))
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
    
'*****************************************************
'Antes de Llenar todos los Controles se validara antes de Grabar
Set oPersona = New COMDPersona.DCOMPersonas
    
    Call VerificandoDatosAGrabar(oPersona, prsVerHomoninia, pnVerDuplicaDoc, pnVerTamanioDoc)
    
    'Verifica Homonimia
    If TipoActualizacion = PersFilaNueva And pbVerificaHomonimia Then
        If Not prsVerHomoninia.BOF And Not prsVerHomoninia.EOF Then
            GrabarCambiosPersona = False
            Set oPersona = Nothing
            Exit Function
        End If
    'prsVerHomoninia.Close
    End If
    'Verifica Documentos Duplicados
    If pnVerDuplicaDoc <> -1 Then
        GrabarCambiosPersona = False
        Set oPersona = Nothing
        Exit Function
    End If
    
    'Verificando Tamaño de Documento
    If pnVerTamanioDoc <> -1 Then
        GrabarCambiosPersona = False
        Set oPersona = Nothing
        Exit Function
    End If
    
Set oPersona = Nothing
'**********************************************
Set rsHojaEval = New ADODB.Recordset
Set rsEstFinan = New ADODB.Recordset 'WIOR 20140321

Call LlenaRecordSet_Persona
Call LlenaRecordSet_Documentos
Call LlenaRecordSet_Relaciones
Call LlenaRecordSet_RefComeciales
Call LlenaRecordset_RefBancarias
Call LlenaRecordset_PatVehicular

Call LlenaRecordset_PatOtros
Call LlenaRecordset_PatInmuebles

Call LlenaRecordset_FteIng
Call LlenaRecordset_FteCosto

Call LlenaRecordset_FteIngDep
Call LlenaRecordset_FteIngInd

Call LlenaRecordSet_Visitas 'madm 20100326

Call LlenaRecordSet_Ventas ' MAVM 20100607 BAS II

'****************CTI3  (07SET2018) oDetaJuri.
   
   Dim oDetaJuri As New UPersona_Cli
   Set oDetaJuri = Nothing
   
        
If frmPersona.lblTopera = "N" Then
        
    If (nTipoAct) = 0 Then nTipoAct = PersFilaNueva
            If frmPersonaJurDatosAdic.lblPersElejida.Caption = 1 Then 'persona
               Call LlenaDetallesJuridicos(frmPersonaJurDatosAdic.fleAccionistas, frmPersonaJurDatosAdic.fleDirectorio, frmPersonaJurDatosAdic.fleGerencias, frmPersonaJurDatosAdic.flePatrimonio, frmPersonaJurDatosAdic.flePatOtrasEmpresa, frmPersonaJurDatosAdic.fleCargos, frmPersonaJurDatosAdic.lblPersElejida.Caption)
               Call LlenaRecordSet_DetaJuridicos
               Unload frmPersonaJurDatosAdic
               Set frmPersonaJurDatosAdic = Nothing
               bDAexiste = True
        Else
            If frmPersonaJurDatosAdic.lblPersElejida.Caption = 2 Then 'juridico
               Call LlenaDetallesJuridicos(frmPersonaJurDatosAdic.fleAccionistas, frmPersonaJurDatosAdic.fleDirectorio, frmPersonaJurDatosAdic.fleGerencias, frmPersonaJurDatosAdic.flePatrimonio, frmPersonaJurDatosAdic.flePatOtrasEmpresa, frmPersonaJurDatosAdic.fleCargos, frmPersonaJurDatosAdic.lblPersElejida.Caption)
               Call LlenaRecordSet_DetaJuridicos
               Unload frmPersonaJurDatosAdic
               Set frmPersonaJurDatosAdic = Nothing
               bDAexiste = True
            Else
               If frmPersonaJurDatosAdic.lblPersElejida.Caption = 0 Then bDAexiste = False
            End If
        End If
        
Else
 If frmPersona.lblTopera = "M" Then
    If nTipoAct = 0 Then nTipoAct = PersFilaModificada
    
    nValor = IIf(frmPersonaJurDatosAdic.lblPersElejida.Caption <> "", frmPersonaJurDatosAdic.lblPersElejida.Caption, 0)
         If nValor = 1 Then  'persona
               Call LlenaDetallesJuridicos(frmPersonaJurDatosAdic.fleAccionistas, frmPersonaJurDatosAdic.fleDirectorio, frmPersonaJurDatosAdic.fleGerencias, frmPersonaJurDatosAdic.flePatrimonio, frmPersonaJurDatosAdic.flePatOtrasEmpresa, frmPersonaJurDatosAdic.fleCargos, frmPersonaJurDatosAdic.lblPersElejida.Caption)
               Call LlenaRecordSet_DetaJuridicos
               Unload frmPersonaJurDatosAdic
               Set frmPersonaJurDatosAdic = Nothing
               bDAexiste = True
        Else
            If nValor = 2 Then 'juridico
               Call LlenaDetallesJuridicos(frmPersonaJurDatosAdic.fleAccionistas, frmPersonaJurDatosAdic.fleDirectorio, frmPersonaJurDatosAdic.fleGerencias, frmPersonaJurDatosAdic.flePatrimonio, frmPersonaJurDatosAdic.flePatOtrasEmpresa, frmPersonaJurDatosAdic.fleCargos, frmPersonaJurDatosAdic.lblPersElejida.Caption)
               Call LlenaRecordSet_DetaJuridicos
               Unload frmPersonaJurDatosAdic
               Set frmPersonaJurDatosAdic = Nothing
               bDAexiste = True
            Else
               If frmPersonaJurDatosAdic.lblPersElejida.Caption = 0 Then bDAexiste = False
            End If
        End If
 
 End If
End If
'****************

Set oPersona = New COMDPersona.DCOMPersonas

Call oPersona.GrabarDatosPersona(sPersCod, psCodAge, psCodCMACT, pdFecha, psCodUser, nTipoAct, rsPersona, RFirma, rsDocs, rsRelac, rsFteIng, rsRefCom, rsFIDep, _
                                rsRefBan, rsPatVeh, rsFIInd, rsFICosto, rsDatosAdc, vDocs(0), vDocs(1), vDocs(2), vDocs(3), rsHojaEval, nIntValor, rsPatOtros, rsPatInmuebles, rsVisita, rsVentas, rsEstFinan, bDAexiste) 'WIOR 20140321 AGREGO rsEstFinanv    ***CTI3 07092018
'agrego parametro rsVisita madm 20100329
'JUEZ 20131007 *********************************************************************
If cMotivoActu = "2" Then
    Call oPersona.GrabarDatosPersonaCampDatosBitacora(sPersCod, rsPersona, rsDocs, psCodUserResp, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser))
End If
'END JUEZ **************************************************************************
'ARCV 07-06-2006
psPersCodGrabar = sPersCod

Set oPersona = Nothing
'***********************************************************************************
    
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
    
    i = 0
    Do While i <= nNumRefCom - 1
        If RefCom(i).nTipoAct = PersFilaEliminda Then
            For j = i To nNumRefCom - 2
                RefCom(i) = RefCom(i + 1)
            Next j
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
            For j = i To nNumRefBan - 2
                RefBan(i) = RefBan(i + 1)
            Next j
            nNumRefBan = nNumRefBan - 1
            ReDim Preserve RefBan(nNumRefBan)
            i = i - 1
        Else
            RefBan(i).nTipoAct = PersFilaSinCambios
        End If
        i = i + 1
    Loop

    i = 0
    Do While i <= nNumPatVeh - 1
        If PatVeh(i).nTipoAct = PersFilaEliminda Then
            For j = i To nNumPatVeh - 2
                PatVeh(i) = PatVeh(i + 1)
            Next j
            nNumPatVeh = nNumPatVeh - 1
            ReDim Preserve PatVeh(nNumPatVeh)
            i = i - 1
        Else
            PatVeh(i).nTipoAct = PersFilaSinCambios
        End If
        i = i + 1
    Loop
    
'*** PEAC 20080412
    i = 0
    Do While i <= nNumPatOtros - 1
        If PatOtros(i).nTipoAct = PersFilaEliminda Then
            For j = i To nNumPatOtros - 2
                PatOtros(i) = PatOtros(i + 1)
            Next j
            nNumPatOtros = nNumPatOtros - 1
            ReDim Preserve PatOtros(nNumPatOtros)
            i = i - 1
        Else
            PatOtros(i).nTipoAct = PersFilaSinCambios
        End If
        i = i + 1
    Loop

'*** PEAC 20080412
    i = 0
    Do While i <= nNumPatInmuebles - 1
        If PatInmuebles(i).nTipoAct = PersFilaEliminda Then
            For j = i To nNumPatInmuebles - 2
                PatInmuebles(i) = PatInmuebles(i + 1)
            Next j
            nNumPatInmuebles = nNumPatInmuebles - 1
            ReDim Preserve PatInmuebles(nNumPatInmuebles)
            i = i - 1
        Else
            PatInmuebles(i).nTipoAct = PersFilaSinCambios
        End If
        i = i + 1
    Loop

    
End Function



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
'madm 20100327
Public Sub AdicionaVisita()
    nNumVisitas = nNumVisitas + 1
    ReDim Preserve Visita(nNumVisitas)
    Visita(nNumVisitas).nTipoAct = PersFilaNueva
End Sub
'end madm

'MAVM 20100607 ***
Public Sub AdicionaVentas()
    nNumVentas = nNumVentas + 1
    ReDim Preserve Ventas(nNumVentas)
    Ventas(nNumVentas).nTipoAct = PersFilaNueva
End Sub
'***

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

'*** PEAC 20080412
Public Sub AdicionaPatOtros()
    nNumPatOtros = nNumPatOtros + 1
    ReDim Preserve PatOtros(nNumPatOtros)
    PatOtros(nNumPatOtros).nTipoAct = PersFilaNueva
End Sub

'*** PEAC 20080412
Public Sub AdicionaPatInmuebles()
    nNumPatInmuebles = nNumPatInmuebles + 1
    ReDim Preserve PatInmuebles(nNumPatInmuebles)
    PatInmuebles(nNumPatInmuebles).nTipoAct = PersFilaNueva
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

Private Sub Class_Initialize()
Dim CIni As COMConecta.DCOMClasIni
    Set CIni = New COMConecta.DCOMClasIni
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


Public Sub RecuperaDocumentos(ByVal psCodPers As String, ByVal pRs As ADODB.Recordset)
    nNumDoc = 0
     
    ReDim Docs(nNumDoc)
    Do While Not pRs.EOF
        nNumDoc = nNumDoc + 1
        ReDim Preserve Docs(nNumDoc)
        Docs(nNumDoc - 1).nTipoAct = 0
        Docs(nNumDoc - 1).sTipo = pRs!cConsDescripcion & Space(75) & pRs!nConsValor
        Docs(nNumDoc - 1).nTipo = CInt(pRs!cPersIDTpo)
        Docs(nNumDoc - 1).nTipoOrig = CInt(pRs!cPersIDTpo)
        Docs(nNumDoc - 1).sNumero = Trim(pRs!cPersIDnro)
        pRs.MoveNext
    Loop

End Sub
Private Sub RecuperaRelacionesdePersona(ByVal psPersCod As String, ByVal pRs As ADODB.Recordset)

 If Not pRs Is Nothing Then

    nNumRel = 0
    ReDim Relac(nNumRel)
    Do While Not pRs.EOF
        nNumRel = nNumRel + 1
        ReDim Preserve Relac(nNumRel)
        Relac(nNumRel - 1).sPersCodOrig = pRs!cPersRelacPersCod
        Relac(nNumRel - 1).sPersCod = pRs!cPersRelacPersCod
        Relac(nNumRel - 1).sApellNombres = Trim(pRs!cPersNombre)
        Relac(nNumRel - 1).sRelacion = pRs!cConsDescripcion & Space(75) & pRs!nPersRelac
        Relac(nNumRel - 1).sPersRelacBenef = pRs!cConsDescripcion2 & Space(75) & pRs!nConsValor2
        'Relac(nNumRel - 1).sPersRelacEstado = R!cConsDescripcion4 & Space(50) & Format(IIf(IsNull(R!bPersRelacEstado), 0, R!bPersRelacEstado), "#0")
        Relac(nNumRel - 1).nPersRelacBenefPorc = IIf(IsNull(pRs!nPersRelacBenefPorc), 0, pRs!nPersRelacBenefPorc)
        Relac(nNumRel - 1).sPersRelacAMP = pRs!cConsDescripcion3 & Space(75) & Format(IIf(IsNull(pRs!bPersRelacAMP), 0, pRs!bPersRelacAMP), "#0")
        Relac(nNumRel - 1).nTipoAct = 0
        Relac(nNumRel - 1).sMovAct = pRs!cMovNro
        pRs.MoveNext
    Loop
End If
End Sub

'madm 20100327
Private Sub RecuperaVisitas(ByVal psPersCod As String, ByVal pRs As ADODB.Recordset)

 If Not pRs Is Nothing Then

    nNumVisitas = 0
    ReDim Visita(nNumVisitas)
    Do While Not pRs.EOF
        nNumVisitas = nNumVisitas + 1
        ReDim Preserve Visita(nNumVisitas)
        Visita(nNumVisitas - 1).sPersCod = pRs!cPersRelacPersCod
        Visita(nNumVisitas - 1).sPersCodOrig = pRs!cPersRelacPersCod
        Visita(nNumVisitas - 1).sApellNombres = Trim(pRs!cPersNombre)
        Visita(nNumVisitas - 1).sDireccion = Trim(pRs!cDireccion)
        Visita(nNumVisitas - 1).sUsual = pRs!cConsDescripcion & Space(75) & pRs!sUsual
        Visita(nNumVisitas - 1).dFecVis = pRs!dFecVis
        Visita(nNumVisitas - 1).sVisObserva = pRs!cVisObserva
        'Visita(nNumVisitas - 1).sUsual = pRs!sUsual
        Visita(nNumVisitas - 1).nTipoAct = 0
        pRs.MoveNext
    Loop
End If
End Sub
'end madm

'MAVM 20100607 BAS II***
Private Sub RecuperaVentas(ByVal psPersCod As String, ByVal pRs As ADODB.Recordset)
    If Not pRs Is Nothing Then

    nNumVentas = 0
    ReDim Ventas(nNumVentas)
    Do While Not pRs.EOF
        nNumVentas = nNumVentas + 1
        ReDim Preserve Ventas(nNumVentas)
        Ventas(nNumVentas - 1).sPersCodOrigVA = Trim(pRs!cPersCodAudEEFF)
        Ventas(nNumVentas - 1).sPersCod = Trim(pRs!cPersCodAudEEFF)
        Ventas(nNumVentas - 1).sApellNombres = Trim(pRs!cPersNombre)
        Ventas(nNumVentas - 1).sPeriodo = Trim(pRs!nPeriodo)
        Ventas(nNumVentas - 1).nMonto = pRs!nMontoVA
        Ventas(nNumVentas - 1).dFecAud = pRs!dFecAudEEFF
        Ventas(nNumVentas - 1).nTipoAct = 0
        pRs.MoveNext
    Loop
End If
End Sub
'***

Private Sub RecuperaRefComercial(ByVal psPersCod As String, ByVal pRs As ADODB.Recordset)
nNumRefCom = 0
    
    ReDim RefCom(nNumRefCom)
    
    Do While Not pRs.EOF
        nNumRefCom = nNumRefCom + 1
        ReDim Preserve RefCom(nNumRefCom)
        RefCom(nNumRefCom - 1).sRefComNom = pRs!cNomRefCom
        RefCom(nNumRefCom - 1).sRefComFono = pRs!cFonoRefCom
        RefCom(nNumRefCom - 1).sRefComComentario = pRs!cComentarioRefCom
        RefCom(nNumRefCom - 1).sRefComDireccion = pRs!cDireccionRefCom
        RefCom(nNumRefCom - 1).sRefComTipo = pRs!cConsDescripcion & Space(75) & pRs!nTipoRelaCom
        RefCom(nNumRefCom - 1).nRefComCod = pRs!nNumRefCom
        RefCom(nNumRefCom - 1).nTipoAct = 0
        pRs.MoveNext
    Loop
    
End Sub

Private Sub RecuperaPatVehicular(ByVal psPersCod As String, ByVal pRs As ADODB.Recordset)

nNumPatVeh = 0
    
    ReDim PatVeh(nNumPatVeh)
    
    Do While Not pRs.EOF
        nNumPatVeh = nNumPatVeh + 1
        ReDim Preserve PatVeh(nNumPatVeh)
        PatVeh(nNumPatVeh - 1).nPatVehCod = pRs!nNumPatri
        PatVeh(nNumPatVeh - 1).sPatVehMarca = pRs!cMarca
        PatVeh(nNumPatVeh - 1).sPatVehModelo = pRs!cModelo
        PatVeh(nNumPatVeh - 1).sPatVehFecFab = pRs!cFecFabrica
        PatVeh(nNumPatVeh - 1).sPatVehPlaca = pRs!cPlaca
        PatVeh(nNumPatVeh - 1).nPatVehValCom = pRs!nValorCom
        PatVeh(nNumPatVeh - 1).nTipoAct = 0
        pRs.MoveNext
    Loop
    
End Sub

'*** PEAC 20080412
Private Sub RecuperaPatOtros(ByVal psPersCod As String, ByVal pRs As ADODB.Recordset)

nNumPatOtros = 0
    
    ReDim PatOtros(nNumPatOtros)
    
    Do While Not pRs.EOF
        nNumPatOtros = nNumPatOtros + 1
        ReDim Preserve PatOtros(nNumPatOtros)
        PatOtros(nNumPatOtros - 1).nPatOtrosCod = pRs!nNumPatri
        PatOtros(nNumPatOtros - 1).sPatOtrosDescripcion = pRs!cDescripcion
        PatOtros(nNumPatOtros - 1).nPatOtrosValCom = pRs!nValorCom
        PatOtros(nNumPatOtros - 1).nTipoAct = 0
        pRs.MoveNext
    Loop
    
End Sub

'*** PEAC 20080412
Private Sub RecuperaPatInmuebles(ByVal psPersCod As String, ByVal pRs As ADODB.Recordset)

nNumPatInmuebles = 0
    
    ReDim PatInmuebles(nNumPatInmuebles)
    
    Do While Not pRs.EOF
        nNumPatInmuebles = nNumPatInmuebles + 1
        ReDim Preserve PatInmuebles(nNumPatInmuebles)
        PatInmuebles(nNumPatInmuebles - 1).nPatInmueblesCod = pRs!nNumPatri
        PatInmuebles(nNumPatInmuebles - 1).sPatInmueblesUbicacion = pRs!cUbicacion
        PatInmuebles(nNumPatInmuebles - 1).nPatInmueblesAreaTerreno = pRs!nAreaTerreno
        PatInmuebles(nNumPatInmuebles - 1).nPatInmueblesAreaConstru = pRs!nAreaConstru
        PatInmuebles(nNumPatInmuebles - 1).sPatInmueblesTipoUso = pRs!cTipoUso
        PatInmuebles(nNumPatInmuebles - 1).sPatInmueblesRRPP = pRs!cRRPP
        PatInmuebles(nNumPatInmuebles - 1).nPatInmueblesValCom = pRs!nValorCom
        PatInmuebles(nNumPatInmuebles - 1).nTipoAct = 0
        pRs.MoveNext
    Loop
    
End Sub


Private Sub RecuperaRefBancaria(ByVal psPersCod As String, ByVal pRs As ADODB.Recordset)

nNumRefBan = 0
    
    ReDim RefBan(nNumRefBan)

    Do While Not pRs.EOF
        nNumRefBan = nNumRefBan + 1
        ReDim Preserve RefBan(nNumRefBan)
        RefBan(nNumRefBan - 1).sRefBanCodIFOrg = pRs!cCodBanco
        RefBan(nNumRefBan - 1).sRefBanCodIF = pRs!cCodBanco
        RefBan(nNumRefBan - 1).sRefBanNombre = pRs!cBanNombre
        RefBan(nNumRefBan - 1).sRefBanNumCta = pRs!cNumCta
        RefBan(nNumRefBan - 1).sRefBanNumTar = pRs!cNumTarjeta
        RefBan(nNumRefBan - 1).nRefBanLinCred = pRs!nMontoLinCred
        RefBan(nNumRefBan - 1).nTipoAct = 0
        pRs.MoveNext
    Loop
    
End Sub

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
Dim j As Integer
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
                For j = 0 To FtesIngreso(i).nNumFtesDep - 1
                    If FtesIngreso(i).FtesIngresoDep(j).dFecEval > dFecEvalTmp Then
                        dFecEvalTmp = FtesIngreso(i).FtesIngresoDep(j).dFecEval
                    End If
                Next j
            Else
                For j = 0 To FtesIngreso(i).nNumFtesIndep - 1
                    If FtesIngreso(i).FtesIngresoInd(j).dFecEval > dFecEvalTmp Then
                        dFecEvalTmp = FtesIngreso(i).FtesIngresoInd(j).dFecEval
                    End If
                Next j
            End If
        End If
    Next i
    PersFuenteBuscaUltimaFechaEval = dFecEvalTmp
End Function

Friend Function PersFuenteBuscaUltimaFechaCaducacion(ByVal psNumFte As String) As Date
Dim i As Integer
Dim j As Integer
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
                For j = 0 To FtesIngreso(i).nNumFtesDep - 1
                    If FtesIngreso(i).FtesIngresoDep(j).dFecCad > dFecEvalTmp Then
                        dFecEvalTmp = FtesIngreso(i).FtesIngresoDep(j).dFecCad
                    End If
                Next j
            Else
                For j = 0 To FtesIngreso(i).nNumFtesIndep - 1
                    If FtesIngreso(i).FtesIngresoInd(j).dFecCad > dFecEvalTmp Then
                        dFecEvalTmp = FtesIngreso(i).FtesIngresoInd(j).dFecCad
                    End If
                Next j
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
            MatFteTmp(nNumFtetmp - 1, 4) = FtesIngreso(i).dFtePersFecInicioAct 'Fecha de Evaluacion de la Fte de Ingreso
            MatFteTmp(nNumFtetmp - 1, 5) = IIf(FtesIngreso(i).sMoneda = gMonedaNacional, "SOLES", "DOLARES") 'Moneda de la Fte de Ingreso
            MatFteTmp(nNumFtetmp - 1, 6) = FtesIngreso(i).sFteIngreso
            MatFteTmp(nNumFtetmp - 1, 7) = Trim((i))
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

Public Sub RecuperaFtesdeIngreso(ByVal psCodPers As String, ByVal pRs As ADODB.Recordset, _
                                Optional ByVal pcCtaCod As String = "", Optional ByVal prsTemPers As ADODB.Recordset, Optional nOpt As Integer = 0)
      
    nNumFtes = 0
    Dim rsConDPer As COMDPersona.DCOMPersonas
    ReDim FtesIngreso(nNumFtes)
    If Not (pRs.EOF And pRs.BOF) Then
        Do While Not pRs.EOF
            nNumFtes = nNumFtes + 1
            ReDim Preserve FtesIngreso(nNumFtes)
            FtesIngreso(nNumFtes - 1).cNumFte = pRs!cNumFuente
            FtesIngreso(nNumFtes - 1).bCostoProd = IIf(IIf(IsNull(pRs!bCostoProd), 0, pRs!bCostoProd), 1, 0)
            FtesIngreso(nNumFtes - 1).dFteIngreso = Format(pRs!dPersFIinicio, "dd/mm/yyyy")
            FtesIngreso(nNumFtes - 1).sCargo = pRs!cPersFICargo
            FtesIngreso(nNumFtes - 1).sComentarios = pRs!cPersFIComentario
            FtesIngreso(nNumFtes - 1).sFteIngreso = pRs!cPersFIPersCod
            FtesIngreso(nNumFtes - 1).sFteIngresoOrig = pRs!cPersFIPersCod
            FtesIngreso(nNumFtes - 1).sMoneda = Trim((pRs!cPersFIMoneda))
            FtesIngreso(nNumFtes - 1).nTipo = IIf(IsNull(pRs!nPersTipFte), 0, pRs!nPersTipFte)
            FtesIngreso(nNumFtes - 1).sRazSocUbiGeo = Trim(IIf(IsNull(pRs!cRazSocUbiGeo), "", pRs!cRazSocUbiGeo))
            FtesIngreso(nNumFtes - 1).sRazSocDirecc = Trim(IIf(IsNull(pRs!cRazSocDirecc), "", pRs!cRazSocDirecc))
            FtesIngreso(nNumFtes - 1).sRazSocTelef = Trim(IIf(IsNull(pRs!cRazSocTelef), "", pRs!cRazSocTelef))
            FtesIngreso(nNumFtes - 1).sRazSocDescrip = Trim(IIf(IsNull(pRs!cRazSocDescrip), "", pRs!cRazSocDescrip))
            FtesIngreso(nNumFtes - 1).nTipoAct = 0
            FtesIngreso(nNumFtes - 1).cFtePersRuc = Trim(IIf(IsNull(pRs!cNroRUC), "", pRs!cNroRUC))
            
            If nOpt = 1 Then
                FtesIngreso(nNumFtes - 1).dFtePersFecInicioAct = Format(IIf(IsNull(pRs!Descr), "01/01/1900", pRs!Descr), "dd/mm/yyyy")
            Else
                FtesIngreso(nNumFtes - 1).dFtePersFecInicioAct = Format("01/01/1900", "dd/mm/yyyy")
            End If
            'FtesIngreso(nNumFtes - 1).dFtePersFecInicioAct = Format(IIf(IsNull(pRs!dPersFIinicio), "01/01/1900", pRs!dPersFIinicio), "dd/mm/yyyy")
            FtesIngreso(nNumFtes - 1).cFtePersTipoPersJur = Trim(IIf(IsNull(pRs!cPersJurTpoDesc), "", pRs!cPersJurTpoDesc))
            FtesIngreso(nNumFtes - 1).cFtePersTelefono = Trim(IIf(IsNull(pRs!cPersTelefono), "", pRs!cPersTelefono))
            FtesIngreso(nNumFtes - 1).cFtePersCIUU = Trim(IIf(IsNull(pRs!cCIIUdescripcion), "", pRs!cCIIUdescripcion))
            FtesIngreso(nNumFtes - 1).cFtePersCondicionDomic = Trim(IIf(IsNull(pRs!cPersDireccCondicionDes), "", pRs!cPersDireccCondicionDes))
            FtesIngreso(nNumFtes - 1).cFtePersMagnitudEmp = Trim(IIf(IsNull(pRs!cPersJurMagnitudDes), "", pRs!cPersJurMagnitudDes))
            FtesIngreso(nNumFtes - 1).nFtePersNroEmpleados = IIf(IsNull(pRs!nPersJurEmpleados), 0, pRs!nPersJurEmpleados)
            FtesIngreso(nNumFtes - 1).cFtePersDireccion = Trim(IIf(IsNull(pRs!cPersDireccDomicilio), "", pRs!cPersDireccDomicilio))
            FtesIngreso(nNumFtes - 1).cFtePersDpto = Trim(IIf(IsNull(pRs!DPTO), "", pRs!DPTO))
            FtesIngreso(nNumFtes - 1).cFtePersProv = Trim(IIf(IsNull(pRs!PROV), "", pRs!PROV))
            FtesIngreso(nNumFtes - 1).cFtePersDist = Trim(IIf(IsNull(pRs!Dist), "", pRs!Dist))
            FtesIngreso(nNumFtes - 1).cFtePersZona = Trim(IIf(IsNull(pRs!Zona), "", pRs!Zona))
            'Set FtesIngreso(nNumFtes - 1).RFteIng = rsConDPer.ObtenerDatosHojEvaluaci(pRs!cNumFuente, pRs!dPersFIinicio)
            pRs.MoveNext
        Loop
    Else
       If Not Nothing Is prsTemPers Then
        If Not (prsTemPers.EOF And prsTemPers.BOF) Then
         FtesIngreso(nNumFtes).cFtePersCIUU = Trim(IIf(IsNull(prsTemPers!cCIIUdescripcion), "", prsTemPers!cCIIUdescripcion))
         FtesIngreso(nNumFtes).cFtePersMagnitudEmp = IIf(IsNull(prsTemPers!cPersJurMagnitudDes), "", prsTemPers!cPersJurMagnitudDes)
         FtesIngreso(nNumFtes).cFtePersCondicionDomic = IIf(IsNull(prsTemPers!cPersDireccCondicionDes), "", prsTemPers!cPersDireccCondicionDes)
         FtesIngreso(nNumFtes).cFtePersDireccion = IIf(IsNull(prsTemPers!cPersDireccDomicilio), "", prsTemPers!cPersDireccDomicilio)
         FtesIngreso(nNumFtes).cFtePersDpto = IIf(IsNull(prsTemPers!DPTO), "", Trim(prsTemPers!DPTO))
         FtesIngreso(nNumFtes).cFtePersProv = IIf(IsNull(prsTemPers!PROV), "", Trim(prsTemPers!PROV))
         FtesIngreso(nNumFtes).cFtePersDist = IIf(IsNull(prsTemPers!Dist), "", Trim(prsTemPers!Dist))
         FtesIngreso(nNumFtes).cFtePersZona = IIf(IsNull(prsTemPers!Zona), "", Trim(prsTemPers!Zona))
         prsTemPers.Close
        End If
       End If
    End If
End Sub

Public Sub RecuperaFtesIngresoDependiente(ByVal pnNumFte As Integer, _
                                            ByVal pRs As ADODB.Recordset)

    FtesIngreso(pnNumFte).nNumFtesDep = 0
    Do While Not pRs.EOF
        FtesIngreso(pnNumFte).nNumFtesDep = FtesIngreso(pnNumFte).nNumFtesDep + 1
        ReDim Preserve FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep)
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).cNumFte = pRs!cNumFuente
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).dFecCad = pRs!dPersFICaduca
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).dFecEval = Format(pRs!dPersEval, "dd/mm/yyyy")
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).nGastosFam = Format(pRs!nPersGastoFam, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).nIngCony = Format(pRs!nPersIngCon, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).nIngFam = Format(pRs!nPersIngCli, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).nIngOtros = Format(pRs!nPersOtrIng, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).nIngresos = Format(pRs!nPersIngCli, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).nIngTotales = Format(pRs!nPersIngCon + pRs!nPersOtrIng + pRs!nPersIngCli + pRs!nPersOtrIng, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).nSaldDispFam = Format(FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).nIngTotales - FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).nGastosFam, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoDep(FtesIngreso(pnNumFte).nNumFtesDep - 1).nTipoAct = PersFilaSinCambios
        pRs.MoveNext
    Loop
    
End Sub

Public Sub RecuperaFtesIngresoCostosProd(ByVal pnNumFte As Integer, _
                                        ByVal pRs As ADODB.Recordset)
    
    FtesIngreso(pnNumFte).nNumCostosProd = 0
    Do While Not pRs.EOF
        FtesIngreso(pnNumFte).nNumCostosProd = FtesIngreso(pnNumFte).nNumCostosProd + 1
        ReDim Preserve FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd)
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).dFecEval = Format(pRs!dFecEval, "dd/mm/yyyy")
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).sTpoCultivo = pRs!sTipoCultivo
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).sUniProd = pRs!sUniProd
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).cNumFuente = pRs!cNumFuente
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nTpoCultivo = pRs!nTpoCultivo
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nMaquinaria = pRs!nMaquinaria
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nJornales = pRs!nJornales
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nInsumos = pRs!nInsumos
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nPesticidas = pRs!nPesticidas
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nOtros = pRs!nOtros
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nHectareas = pRs!nHectareas
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nProduccion = pRs!nProduccion
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nUniProd = pRs!nUniProd
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nPreUni = pRs!nPreUni
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nSiembra = pRs!nSiembra
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nMantenimiento = pRs!nMantenimiento
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nDesAgricola = pRs!nDesaAgricola
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nCosecha = pRs!nCosecha
        FtesIngreso(pnNumFte).CostoProd(FtesIngreso(pnNumFte).nNumCostosProd - 1).nPlanOtros = pRs!nPlanOtros
        pRs.MoveNext
    Loop
    
End Sub

Public Sub RecuperaFtesIngresoIndependiente(ByVal pnNumFte As Integer, _
                                            ByVal pRs As ADODB.Recordset)

    FtesIngreso(pnNumFte).nNumFtesIndep = 0
    ReDim FtesIngreso(pnNumFte).FtesIngresoInd(pRs.RecordCount)
    Do While Not pRs.EOF
        FtesIngreso(pnNumFte).nNumFtesIndep = FtesIngreso(pnNumFte).nNumFtesIndep + 1
        ReDim Preserve FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep)
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).cNumFte = pRs!cNumFuente
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).dFecCad = Format(pRs!dPersFICaduca, "dd/mm/yyyy")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).dFecEval = Format(pRs!dPersEval, "dd/mm/yyyy")
        
        'peac 20071227
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).dfecEEFF = Format(pRs!dfecEEFF, "dd/mm/yyyy")
        
        
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nActFijo = Format(pRs!nPersFIActivosFijos, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nActivDisp = Format(pRs!nPersFIActivoDisp, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nActivo = Format(pRs!nPersFIActivoDisp + pRs!nPersFIActivosFijos, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nActivoCircul = "0.00"
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nCompraMercad = Format(pRs!nPersFICostoVentas, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nCreditosCmact = Format(pRs!nPersFICredCMACT, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nCreditosOtros = Format(pRs!nPersFICredOtros, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nCtasxCob = Format(pRs!nPersFICtasxCobrar, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nInventarios = Format(pRs!nPersFIInventarios, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nOtrosEgresos = Format(pRs!nPersFIEgresosOtros, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nPasivoPatrimonio = Format(pRs!nPersFIPatrimonio, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nPersEgrFam = Format(pRs!nPersEgrFam, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nPersIngFam = Format(pRs!nPersIngFam, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nProveedores = Format(pRs!nPersFIProveedores, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nRecupCtasxCob = Format(pRs!nPersFIRecupCtasXCobrar, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nTipoAct = PersFilaSinCambios
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nVentas = Format(pRs!nPersFIVentas, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nPasivo = Format(0#, "#0.00")
        FtesIngreso(pnNumFte).FtesIngresoInd(FtesIngreso(pnNumFte).nNumFtesIndep - 1).nTipoAct = PersFilaSinCambios
        pRs.MoveNext
    Loop
    
End Sub
'**************************************CTI3 10092018
Sub RecuperaDatosAdic(ByVal psCodPers As String, _
                                            ByVal pAdcs As ADODB.Recordset)
Dim nNum As Integer

nNum = 0

If pAdcs.RecordCount > 0 Then
ReDim Preserve RJuridicoJur(pAdcs.RecordCount - 1)
    Do While Not pAdcs.EOF
        
        RJuridicoJur(nNum).sNombrejur = pAdcs!cNombreJur
        RJuridicoJur(nNum).sRUC = pAdcs!vRuc
        RJuridicoJur(nNum).sCi = pAdcs!vCi
        RJuridicoJur(nNum).sNacionalidad = pAdcs!cNacionalidad
        RJuridicoJur(nNum).sCargo = pAdcs!cCargo
        RJuridicoJur(nNum).sRazonSocial = pAdcs!cRazonSocial
        RJuridicoJur(nNum).nApoSol = pAdcs!nApoSol
        RJuridicoJur(nNum).nApoPorc = pAdcs!nApoPorc
        RJuridicoJur(nNum).sCodSbs = pAdcs!vCodSbs
        RJuridicoJur(nNum).sTipo = pAdcs!cTipo
        
        nNum = nNum + 1

        pAdcs.MoveNext
    Loop
End If
End Sub
Function RecuperaLlenado() As Variant
Dim nNum As Integer
Dim MatTempAdc As Variant

On Error GoTo errMatriz

nNum = UBound(RJuridicoJur)

ReDim MatTempAdc(nNum, 9)

For i = 0 To nNum

        MatTempAdc(i, 0) = RJuridicoJur(i).sNombrejur
        MatTempAdc(i, 1) = RJuridicoJur(i).sRUC
        MatTempAdc(i, 2) = RJuridicoJur(i).sCi
        MatTempAdc(i, 3) = RJuridicoJur(i).sNacionalidad
        MatTempAdc(i, 4) = RJuridicoJur(i).sCargo
        MatTempAdc(i, 5) = RJuridicoJur(i).sRazonSocial
        MatTempAdc(i, 6) = RJuridicoJur(i).nApoSol
        MatTempAdc(i, 7) = RJuridicoJur(i).nApoPorc
        MatTempAdc(i, 8) = RJuridicoJur(i).sCodSbs
        MatTempAdc(i, 9) = RJuridicoJur(i).sTipo

Next i

RecuperaLlenado = MatTempAdc
errMatriz:
  'End Function
End Function
'***************************************************

Public Sub RecuperaPersona(ByVal psCodPers As String, Optional ByVal psCtaCod As String = "", Optional ByVal psCodUser As String = "")
    
Dim oPersona As COMDPersona.DCOMPersonas
    
Dim rsPersona As ADODB.Recordset
Dim rsDocs As ADODB.Recordset
Dim rsRelac As ADODB.Recordset
Dim rsFtesIng As ADODB.Recordset
Dim rsRefComer As ADODB.Recordset
Dim rsRefBanc As ADODB.Recordset
Dim rsPatVehic As ADODB.Recordset
Dim rsPersT As ADODB.Recordset
Dim rsVisita As ADODB.Recordset
Dim rsPersRHH As ADODB.Recordset
'***CTI3 10092018
Dim rsDatosAdc As ADODB.Recordset
'****************
Set oPersona = New COMDPersona.DCOMPersonas
Set rsPersT = New ADODB.Recordset

'Call oPersona.RecuperarDatosPersona(psCodPers, gsCodAge, rsPersona, RFirma, rsDocs, rsRelac, rsFtesIng, rsRefComer, nMaxRefCom, rsRefBanc, rsPatVehic, nMaxPatVeh, psCtaCod)

If psCodUser = "" Then
    'Call oPersona.RecuperarDatosPersona(psCodPers, gsCodAge, rsPersona, RFirma, rsDocs, rsRelac, rsFtesIng, rsRefComer, nMaxRefCom, rsRefBanc, rsPatVehic, nMaxPatVeh, psCtaCod, rsPatOtros, nMaxPatOtros, rsPatInmuebles, nMaxPatInmuebles, rsVisita, rsVentas)
    Call oPersona.RecuperarDatosPersona(psCodPers, gsCodAge, rsPersona, RFirma, rsDocs, rsRelac, rsFtesIng, rsRefComer, nMaxRefCom, rsRefBanc, rsPatVehic, nMaxPatVeh, psCtaCod, rsPatOtros, nMaxPatOtros, rsPatInmuebles, nMaxPatInmuebles, rsVisita, rsVentas, rsDatosAdc) '*****cti3
Else
    'MADM 20101020
    'Call oPersona.RecuperarDatosPersona(psCodPers, gsCodAge, rsPersona, RFirma, rsDocs, rsRelac, rsFtesIng, rsRefComer, nMaxRefCom, rsRefBanc, rsPatVehic, nMaxPatVeh, psCtaCod, rsPatOtros, nMaxPatOtros, rsPatInmuebles, nMaxPatInmuebles, rsVisita, rsVentas)
    'Call oPersona.RecuperarDatosPersona(psCodPers, gsCodAge, rsPersona, RFirma, rsDocs, rsRelac, rsFtesIng, rsRefComer, nMaxRefCom, rsRefBanc, rsPatVehic, nMaxPatVeh, psCtaCod, rsPatOtros, nMaxPatOtros, rsPatInmuebles, nMaxPatInmuebles, rsVisita, rsVentas, rsPersRHH, gsCodUser)
    Call oPersona.RecuperarDatosPersona(psCodPers, gsCodAge, rsPersona, RFirma, rsDocs, rsRelac, rsFtesIng, rsRefComer, nMaxRefCom, rsRefBanc, rsPatVehic, nMaxPatVeh, psCtaCod, rsPatOtros, nMaxPatOtros, rsPatInmuebles, nMaxPatInmuebles, rsVisita, rsVentas, rsDatosAdc, rsPersRHH, gsCodUser) '*****CTI3
End If
     
     If Not rsPersona Is Nothing Then
        If Not rsPersona.BOF And Not rsPersona.EOF Then
            Set rsPersT = rsPersona
            sPersCod = rsPersona!cPersCod
            sNombreCompleto = Trim(rsPersona!cPersNombre)
            sApePat = BuscaNombre(rsPersona!cPersNombre, BusqApellidoPaterno)
            sApeMat = BuscaNombre(rsPersona!cPersNombre, BusqApellidoMaterno)
            sNombres = BuscaNombre(rsPersona!cPersNombre, BusqNombres)
            sApeCas = BuscaNombre(rsPersona!cPersNombre, BusqApellidoCasada)
            sSexo = Trim(IIf(IsNull(rsPersona!cPersnatSexo), "", rsPersona!cPersnatSexo))
            sEstadoCivil = Trim(IIf(IsNull(rsPersona!nPersNatEstCiv), "", rsPersona!nPersNatEstCiv))
            nHijos = IIf(IsNull(rsPersona!nPersNatHijos), 0, rsPersona!nPersNatHijos)
            
            
            'ANDE 20190910 Para evaluar si está habilitado para usar en banca por internet
            bBIEsHabilitado = IIf(IsNull(rsPersona!bBIEsHabilitado), False, rsPersona!bBIEsHabilitado)
            'end ande 20190910

            
            
            '*** PEAC 20080412
            nNumEmp = IIf(IsNull(rsPersona!nPersNatNumEmp), 0, rsPersona!nPersNatNumEmp)
    
            dFechaNac = CDate(Format(IIf(IsNull(rsPersona!dPersNacCreac), "01/01/1900", rsPersona!dPersNacCreac), "dd/mm/yyyy"))
            '*** PEAC 20080731
            cFechaAct = rsPersona!cUltimaAct
            
            'MAVM
            dFechaFallec = CDate(Format(IIf(IsNull(rsPersona!dPersFallec), "01/01/1900", rsPersona!dPersFallec), "dd/mm/yyyy"))
            'MAVM 03042009
            
            'MADM 20091114
            'sCentroTrabajo = IIf(IsNull(rsPersona!cCentroTrabajo), "", rsPersona!cCentroTrabajo)
            'sCargo = IIf(IsNull(rsPersona!cCargo), "", rsPersona!cCargo)
            sCredito = IIf(IsNull(rsPersona!NCredito), 0, rsPersona!NCredito)
            sAhorro = IIf(IsNull(rsPersona!nAhorro), 0, rsPersona!nAhorro)
            sServicio = IIf(IsNull(rsPersona!nServicio), 0, rsPersona!nServicio)
            sOtro = IIf(IsNull(rsPersona!nOtro), 0, rsPersona!nOtro)
            sPeps = IIf(IsNull(rsPersona!nPeps), 0, rsPersona!nPeps) '20100408
            'END MADM
            'EJVG20120813 ***
            nSujetoObligadoIni = IIf(IsNull(rsPersona!nSujetoObligado), -1, rsPersona!nSujetoObligado)
            nSujetoObligado = IIf(IsNull(rsPersona!nSujetoObligado), -1, rsPersona!nSujetoObligado)
            nOfCumplimiento = IIf(IsNull(rsPersona!nOfCumplimiento), -1, rsPersona!nOfCumplimiento)
            'END EJVG *******
            nAutorizaUsoDatosIni = IIf(IsNull(rsPersona!nAutorizaUsoDatos), -1, rsPersona!nAutorizaUsoDatos)
            nAutorizaUsoDatos = IIf(IsNull(rsPersona!nAutorizaUsoDatos), -1, rsPersona!nAutorizaUsoDatos) 'FRHU 20151130 ERS077-2015
            'madm 20100503
            sActiGiro1 = IIf(IsNull(rsPersona!cActiGiro1), 0, rsPersona!cActiGiro1)
            'end madm
            
            nPersCargo = IIf(IsNull(rsPersona!nPersCargo), 0, rsPersona!nPersCargo) '** Juez 20120326
            
            sUbicGeografica = rsPersona!cPersDireccUbiGeo
            sDomicilio = rsPersona!cPersDireccDomicilio
    
            sCondicionDomic = IIf(IsNull(rsPersona!cPersDireccCondicion), (gPersCondDomicEnUso), rsPersona!cPersDireccCondicion)
            nValComDomicilio = IIf(IsNull(rsPersona!nPersValComDomicilio), 0, rsPersona!nPersValComDomicilio)
            'JUEZ 20131007 ***********************************
            sUbiGeoNegocio = rsPersona!cPersNegocioUbiGeo
            sNegocioDireccion = rsPersona!cPersNegocioDireccion
            sRefNegocio = rsPersona!cPersNegocioRef
            sNombreCentroLaboral = rsPersona!cPersNombreCentroLaboral 'marg
            
            nRemisionInfoEmail = rsPersona!nRemisionInfoEmail
            'END JUEZ ****************************************
            sTelefonos = IIf(IsNull(rsPersona!cPersTelefono), "", rsPersona!cPersTelefono)
            sTelefonos2 = IIf(IsNull(rsPersona!cPersTelefono2), "", rsPersona!cPersTelefono2)
            'EJVG20111209 ************************************************************
            sCelular = IIf(IsNull(rsPersona!cPersCelular), "", rsPersona!cPersCelular)
            sCelular2 = IIf(IsNull(rsPersona!cPersCelular2), "", rsPersona!cPersCelular2)
            sCelular3 = IIf(IsNull(rsPersona!cPersCelular3), "", rsPersona!cPersCelular3)
            
            sEmail = IIf(IsNull(rsPersona!cEmail), "", rsPersona!cEmail)
            sEmail2 = IIf(IsNull(rsPersona!cEmail2), "", rsPersona!cEmail2) 'EJVG20111209
            '*** PEAC 20080412
            dFechaInscRuc = CDate(Format(IIf(IsNull(rsPersona!dPersIngRuc), "01/01/1900", rsPersona!dPersIngRuc), "dd/mm/yyyy"))
            dFechaIniActi = CDate(Format(IIf(IsNull(rsPersona!dPersIniActi), "01/01/1900", rsPersona!dPersIniActi), "dd/mm/yyyy"))
            
            nNumDependi = IIf(IsNull(rsPersona!nNumDependi), 0, rsPersona!nNumDependi)
            sActComple = IIf(IsNull(rsPersona!cActComple), "", rsPersona!cActComple)
            nNumPtosVta = IIf(IsNull(rsPersona!nNumPtosVta), 0, rsPersona!nNumPtosVta)
            sActiGiro = IIf(IsNull(rsPersona!cActiGiro), "", rsPersona!cActiGiro)
            '*** FIN PEAC
            sCiiu = IIf(IsNull(rsPersona!cPersCIIU), "", rsPersona!cPersCIIU)
            sCiiuDescripcion = IIf(IsNull(rsPersona!cCIIUdescripcion), "", rsPersona!cCIIUdescripcion)
            
            '*** PEAC 20080412
            nTipoCompDescripcion = IIf(IsNull(rsPersona!nTipoCompetencia), 4, rsPersona!nTipoCompetencia)
            nTipoSistInforDescripcion = IIf(IsNull(rsPersona!nTipoSistInfor), 3, rsPersona!nTipoSistInfor)
            nCadenaProdDescripcion = IIf(IsNull(rsPersona!nTipoCadenaProd), 3, rsPersona!nTipoCadenaProd)
            
            nMonedaPatri = IIf(IsNull(rsPersona!nMoneyPatri), 0, rsPersona!nMoneyPatri)
            '*** FIN PEAC
            
            sEstado = rsPersona!cPersEstado
            nTalla = IIf(IsNull(rsPersona!nPersTalla), 0#, rsPersona!nPersTalla)
            nPeso = IIf(IsNull(rsPersona!nPersPeso), 0#, rsPersona!nPersPeso)
            sTipoSangre = Trim((IIf(IsNull(rsPersona!nPersTpoSangre), 0, rsPersona!nPersTpoSangre)))
            cNacionalidad = Trim(IIf(IsNull(rsPersona!cNacionalidad), "", rsPersona!cNacionalidad))
            
            '*** PEAC 20080801
            cMotivoActu = Trim(IIf(IsNull(rsPersona!TipoActualizacion), "0", rsPersona!TipoActualizacion))
            sRefDomicilio = Trim(IIf(IsNull(rsPersona!cPersRefDomicilio), "", rsPersona!cPersRefDomicilio))
            
            nResidencia = IIf(IsNull(rsPersona!nResidente), 0, rsPersona!nResidente)
            cPaisReside = IIf(IsNull(rsPersona!cPaisReside), "", rsPersona!cPaisReside) '** Juez 20120327
            nTipoPersona = rsPersona!nPersPersoneria
            'CUSCO
            nIngresoProm = rsPersona!nPersIngresoProm
            'Datos de Persona Juridica
            'srsPersonaazonSocial
            sSiglas = IIf(IsNull(rsPersona!cPersJurSigla), "", rsPersona!cPersJurSigla)
            sTipoPersJur = IIf(IsNull(rsPersona!cPersJurMagnitud), "", rsPersona!cPersJurTpo)
            nNumEmplead = IIf(IsNull(rsPersona!nPersJurEmpleados), 0, rsPersona!nPersJurEmpleados)
            sMagnitudEmp = IIf(IsNull(rsPersona!cPersJurMagnitud), "", rsPersona!cPersJurMagnitud)
            sObjetoSocial = IIf(IsNull(rsPersona!cPersJurObjetoSocial), "", rsPersona!cPersJurObjetoSocial) '** Juez 20120328
            sMagnitudPersNat = IIf(IsNull(rsPersona!cPersNatMagnitud), "", rsPersona!cPersNatMagnitud) 'JACA 20110427
            sPersCodSbs = IIf(IsNull(rsPersona!cPersCodSbs), "", rsPersona!cPersCodSbs)
            nPersRelInst = IIf(IsNull(rsPersona!npersrelainst), 4, rsPersona!npersrelainst)
            
            If Not psCodUser = "" Then
            'MADM 20101020
            If Not rsPersRHH.BOF And Not rsPersRHH.EOF Then
               nPersRRHH = 1  'MADM 20101020
            Else
                nPersRRHH = 0  'MADM 20101020
            End If
            'END MADM
            End If
            nTipoAct = 0
        Else
            sPersCod = ""
            rsPersona.Close
            Set rsPersona = Nothing
            Exit Sub
        End If
        
        'Carga los Diferentes Tipos de Documentos que Posee la Persona
        Call RecuperaDocumentos(psCodPers, rsDocs)
        'Carga las diferentes Personas que Tienen Relacion con la persona actual
        Call RecuperaRelacionesdePersona(psCodPers, rsRelac)
        'Carga las Fuentes de Ingreso de la Persona
        Call RecuperaFtesdeIngreso(psCodPers, rsFtesIng, psCtaCod, rsPersT)
        Call RecuperaRefComercial(psCodPers, rsRefComer)    'Carga las Referencias Comerciales de la Persona
        Call RecuperaRefBancaria(psCodPers, rsRefBanc)       'Carga las Referencias Bancarias de la Persona
        
        Call RecuperaPatVehicular(psCodPers, rsPatVehic)    'Carga el Patrimonio Vehicular
        Call RecuperaPatOtros(psCodPers, rsPatOtros)    'Carga el Patrimonio Otros
        Call RecuperaPatInmuebles(psCodPers, rsPatInmuebles)    'Carga el Patrimonio Inmuebles
        
        Call RecuperaVisitas(psCodPers, rsVisita)
        
        Call RecuperaVentas(psCodPers, rsVentas) 'MAVM 20100607 BAS II
        
        '*** CTI3
        Call RecuperaDatosAdic(psCodPers, rsDatosAdc) '*****CTI3
        
End If
End Sub

Public Sub RecuperaPersona_Solicitud(ByVal psCodPers As String, Optional pdFecSis As Date = "1900/01/01")
    
Dim oPersona As COMDPersona.DCOMPersonas
    
Dim rsPersona As ADODB.Recordset
Dim rsFtesIng As ADODB.Recordset
Dim rsDocs As ADODB.Recordset
    
Set oPersona = New COMDPersona.DCOMPersonas
Call oPersona.RecuperarDatosPersona_Solicitud(psCodPers, gsCodAge, rsPersona, rsFtesIng, rsDocs, pdFecSis)
Set oPersona = Nothing
If Not rsPersona Is Nothing Then

    If Not rsPersona.BOF And Not rsPersona.EOF Then
        sPersCod = rsPersona!cPersCod
        sNombreCompleto = Trim(rsPersona!cPersNombre)
    Else
        sPersCod = ""
        rsPersona.Close
        Exit Sub
    End If
    Call RecuperaDocumentos(psCodPers, rsDocs)
    Call RecuperaFtesdeIngreso(psCodPers, rsFtesIng, , , 1)
End If
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

'*** PEAC 20080412
Public Property Get TIPOCOMPDescripcion() As String
    TIPOCOMPDescripcion = nTipoCompDescripcion
End Property

Public Property Let TIPOCOMPDescripcion(ByVal vNewValue As String)
    nTipoCompDescripcion = vNewValue
End Property

'*** PEAC 20080412
Public Property Get MonedaPatri() As String
    MonedaPatri = nMonedaPatri
End Property

Public Property Let MonedaPatri(ByVal vNewValue As String)
    nMonedaPatri = vNewValue
End Property


'*** PEAC 20080412
Public Property Get TIPOSISTINFORDescripcion() As String
    TIPOSISTINFORDescripcion = nTipoSistInforDescripcion
End Property

Public Property Let TIPOSISTINFORDescripcion(ByVal vNewValue As String)
    nTipoSistInforDescripcion = vNewValue
End Property

'*** PEAC 20080412
Public Property Get CADENAPRODDescripcion() As String
    CADENAPRODDescripcion = nCadenaProdDescripcion
End Property

Public Property Let CADENAPRODDescripcion(ByVal vNewValue As String)
    nCadenaProdDescripcion = vNewValue
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

'ANDE 20190910 determina si el cliente está permitido para usar banca por internet
Public Property Get BIHabilitado() As Boolean
    BIHabilitado = bBIEsHabilitado
End Property

'END ANDE20190910

'*** PEAC 20080412
Public Property Get NumEmp() As Integer
    NumEmp = nNumEmp
End Property

Public Property Let NumEmp(ByVal vNewValue As Integer)
    nNumEmp = vNewValue
End Property

Public Property Get FechaNacimiento() As Date
    FechaNacimiento = dFechaNac
End Property

'MAVM 03042009
Public Property Get FechaFallecimiento() As Date
    FechaFallecimiento = dFechaFallec
End Property

Public Property Let FechaFallecimiento(ByVal vNewValue As Date)
    dFechaFallec = vNewValue
End Property

'End MAVM 03042009

'madm 20100503
Public Property Get RcActiGiro1() As String
    RcActiGiro1 = sActiGiro1
End Property

Public Property Let RcActiGiro1(ByVal vNewValue As String)
    sActiGiro1 = vNewValue
End Property
'end madm

'** Juez 20120326 ************************************
Public Property Get RnCargo() As String
    RnCargo = nPersCargo
End Property

Public Property Let RnCargo(ByVal vNewValue As String)
    nPersCargo = vNewValue
End Property
'** End Juez *****************************************

'MADM 20091114
Public Property Get RCredito() As Integer
    RCredito = sCredito
End Property

Public Property Let RCredito(ByVal vNewValue As Integer)
    sCredito = vNewValue
End Property
Public Property Get RAhorro() As Integer
    RAhorro = sAhorro
End Property

Public Property Let RAhorro(ByVal vNewValue As Integer)
    sAhorro = vNewValue
End Property
Public Property Get RServicio() As Integer
    RServicio = sServicio
End Property

Public Property Let RServicio(ByVal vNewValue As Integer)
    sServicio = vNewValue
End Property
Public Property Get ROtro() As Integer
    ROtro = sOtro
End Property

Public Property Let ROtro(ByVal vNewValue As Integer)
    sOtro = vNewValue
End Property

Public Property Get RPeps() As Integer
    RPeps = sPeps
End Property

Public Property Let RPeps(ByVal vNewValue As Integer)
    sPeps = vNewValue
End Property
'END MADM
'EJVG20120813 ***
Public Property Get SujetoObligadoIni() As Integer
    SujetoObligadoIni = nSujetoObligadoIni
End Property
Public Property Let SujetoObligadoIni(ByVal vNewValue As Integer)
    nSujetoObligadoIni = vNewValue
End Property
Public Property Get SujetoObligado() As Integer
    SujetoObligado = nSujetoObligado
End Property
Public Property Let SujetoObligado(ByVal vNewValue As Integer)
    nSujetoObligado = vNewValue
End Property
Public Property Get OfCumplimiento() As Integer
    OfCumplimiento = nOfCumplimiento
End Property
Public Property Let OfCumplimiento(ByVal vNewValue As Integer)
    nOfCumplimiento = vNewValue
End Property
'END EJVG *******
'FRHU ERS077-2015 20151130
Public Property Get AutorizaUsoDatosIni() As Integer
    AutorizaUsoDatosIni = nAutorizaUsoDatosIni
End Property
Public Property Let AutorizaUsoDatosIni(ByVal vNewValue As Integer)
    nAutorizaUsoDatosIni = vNewValue
End Property
Public Property Get AutorizaUsoDatos() As Integer
    AutorizaUsoDatos = nAutorizaUsoDatos
End Property
Public Property Let AutorizaUsoDatos(ByVal vNewValue As Integer)
    nAutorizaUsoDatos = vNewValue
End Property
'FIN FRHU
Public Property Let FechaNacimiento(ByVal vNewValue As Date)
    dFechaNac = vNewValue
End Property

'*** PEAC 20080731
Public Property Get FechaActualizacion() As String
    FechaActualizacion = cFechaAct
End Property
'*** PEAC 20080731
Public Property Let FechaActualizacion(ByVal vNewValue As String)
    cFechaAct = vNewValue
End Property


'*** PEAC 20080412
Public Property Get FechaInscRuc() As Date
    FechaInscRuc = dFechaInscRuc
End Property

Public Property Let FechaInscRuc(ByVal vNewValue As Date)
    dFechaInscRuc = vNewValue
End Property

'*** PEAC 20080412
Public Property Get FechaIniActi() As Date
    FechaIniActi = dFechaIniActi
End Property

Public Property Let FechaIniActi(ByVal vNewValue As Date)
    dFechaIniActi = vNewValue
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

'*** PEAC 20080801
Public Property Get RefDomicilio() As String
    RefDomicilio = sRefDomicilio
End Property

Public Property Let RefDomicilio(ByVal vNewValue As String)
    sRefDomicilio = vNewValue
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
'JUEZ 20131007 ***************************************
Public Property Get UbiGeoNegocio() As String
    UbiGeoNegocio = sUbiGeoNegocio
End Property

Public Property Let UbiGeoNegocio(ByVal vNewValue As String)
    sUbiGeoNegocio = vNewValue
End Property

Public Property Get NegocioDireccion() As String
    NegocioDireccion = sNegocioDireccion
End Property

Public Property Let NegocioDireccion(ByVal vNewValue As String)
    sNegocioDireccion = vNewValue
End Property

Public Property Get RefNegocio() As String
    RefNegocio = sRefNegocio
End Property

Public Property Let RefNegocio(ByVal vNewValue As String)
    sRefNegocio = vNewValue
End Property
Public Property Get NombreCentroLaboral() As String
    NombreCentroLaboral = sNombreCentroLaboral
End Property

Public Property Let NombreCentroLaboral(ByVal vNewValue As String)
    sNombreCentroLaboral = vNewValue
End Property
Public Property Get RemisionInfoEmail() As Integer
    RemisionInfoEmail = nRemisionInfoEmail
End Property

Public Property Let RemisionInfoEmail(ByVal vNewValue As Integer)
    nRemisionInfoEmail = vNewValue
End Property
'END JUEZ ********************************************
Public Property Get Telefonos() As String
    Telefonos = sTelefonos
End Property

Public Property Let Telefonos2(ByVal vNewValue As String)
    sTelefonos2 = vNewValue
End Property

Public Property Get Telefonos2() As String
    Telefonos2 = sTelefonos2
End Property

'EJVG20111207*****************************************
Public Property Let Celular(ByVal vNewValue As String)
    sCelular = vNewValue
End Property
Public Property Get Celular() As String
    Celular = sCelular
End Property
Public Property Let Celular2(ByVal vNewValue As String)
    sCelular2 = vNewValue
End Property
Public Property Get Celular2() As String
    Celular2 = sCelular2
End Property
Public Property Let Celular3(ByVal vNewValue As String)
    sCelular3 = vNewValue
End Property
Public Property Get Celular3() As String
    Celular3 = sCelular3
End Property
'*****************************************************
Public Property Let Email(ByVal vNewValue As String)
    sEmail = vNewValue
End Property

Public Property Get Email() As String
    Email = sEmail
End Property
'EJVG20111207*****************************************
Public Property Let Email2(ByVal vNewValue As String)
    sEmail2 = vNewValue
End Property
Public Property Get Email2() As String
    Email2 = sEmail2
End Property
'*****************************************************
'*** PEAC 20080412
Public Property Let NumDependi(ByVal vNewValue As String)
    nNumDependi = vNewValue
End Property

Public Property Get NumDependi() As String
    NumDependi = nNumDependi
End Property

'*** PEAC 20080412
Public Property Let ActComple(ByVal vNewValue As String)
    sActComple = vNewValue
End Property

Public Property Get ActComple() As String
    ActComple = sActComple
End Property

'*** PEAC 20080412
Public Property Let NumPtosVta(ByVal vNewValue As String)
    nNumPtosVta = vNewValue
End Property

Public Property Get NumPtosVta() As String
    NumPtosVta = nNumPtosVta
End Property

'*** PEAC 20080412
Public Property Let ActiGiro(ByVal vNewValue As String)
    sActiGiro = vNewValue
End Property

Public Property Get ActiGiro() As String
    ActiGiro = sActiGiro
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

'EJVG20120816 ***
Public Sub ObtenerDatosDocumentoxPos(ByVal lnPosicion As Integer, ByRef lnTpoDoc As Integer, ByRef lsTpoDoc As String, ByRef lsNroDoc As String)
    Dim i As Integer
    If nNumDoc > lnPosicion And lnPosicion >= 0 Then
        lnTpoDoc = Docs(lnPosicion).nTipo
        lsTpoDoc = Docs(lnPosicion).sTipo
        lsNroDoc = Docs(lnPosicion).sNumero
    End If
End Sub
'END EJVG

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
'** Juez 20120328 ****************************************
Public Property Get ObjetoSocial() As String
    ObjetoSocial = sObjetoSocial
End Property

Public Property Let ObjetoSocial(ByVal vNewValue As String)
    sObjetoSocial = vNewValue
End Property
'** End Juez *********************************************
'JACA 20110427**********************************************************************
Public Property Get MagnitudPersNat() As String
    MagnitudPersNat = sMagnitudPersNat
End Property

Public Property Let MagnitudPersNat(ByVal vNewValue As String)
    sMagnitudPersNat = vNewValue
End Property

'JACA END***************************************************************************
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

'** Juez 20120327 ******
Public Property Let PaisReside(ByVal vNewValue As String)
    cPaisReside = vNewValue
End Property

Public Property Get PaisReside() As String
    PaisReside = cPaisReside
End Property
'** End Juez ***********

Public Property Let Nacionalidad(ByVal vNewValue As String)
    cNacionalidad = vNewValue
End Property

Public Property Get Nacionalidad() As String
    Nacionalidad = cNacionalidad
End Property

'*** PEAC 20080801
Public Property Let MotivoActu(ByVal vNewValue As String)
    cMotivoActu = vNewValue
End Property
'*** PEAC 20080801
Public Property Get MotivoActu() As String
    MotivoActu = cMotivoActu
End Property


'CUSCO
Public Property Get IngresoPromedio() As String
    IngresoPromedio = nIngresoProm
End Property

Public Property Let IngresoPromedio(ByVal vNewValue As String)
    nIngresoProm = vNewValue
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

'***PEAC 20080412
Public Property Get NumeroPatOtros() As Integer
    NumeroPatOtros = nNumPatOtros
End Property

Public Property Let NumeroPatOtros(ByVal vNewValue As Integer)
    nNumPatOtros = vNewValue
End Property

'***PEAC 20080412
Public Property Get NumeroPatInmuebles() As Integer
    NumeroPatInmuebles = nNumPatInmuebles
End Property

Public Property Let NumeroPatInmuebles(ByVal vNewValue As Integer)
    nNumPatInmuebles = vNewValue
End Property

Public Property Get MaxPatVehicular() As Integer
    MaxPatVehicular = nMaxPatVeh
End Property
'*** PEAC 20080412
Public Property Get MaxPatOtros() As Integer
    MaxPatOtros = nMaxPatOtros
End Property
'*** PEAC 20080412
Public Property Get MaxPatInmuebles() As Integer
    MaxPatInmuebles = nMaxPatInmuebles
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

'*** PEAC 20080412
Public Function ObtenerPatOtrosTipoAct(ByVal pnPunt As Integer) As TPersCambios
    ObtenerPatOtrosTipoAct = PatOtros(pnPunt).nTipoAct
End Function

'*** PEAC 20080412
Public Function ObtenerPatInmueblesTipoAct(ByVal pnPunt As Integer) As TPersCambios
    ObtenerPatInmueblesTipoAct = PatInmuebles(pnPunt).nTipoAct
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

'*** PEAC 20080412
Public Function ObtenerRefComComentario(ByVal pnPunt As Integer) As String
    ObtenerRefComComentario = RefCom(pnPunt).sRefComComentario
End Function
'*** PEAC 20080412
Public Function ObtenerRefComDireccion(ByVal pnPunt As Integer) As String
    ObtenerRefComDireccion = RefCom(pnPunt).sRefComDireccion
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
'*** PEAC 20080412
Public Function ObtenerPatOtrosCod(ByVal pnPunt As Integer) As Integer
    ObtenerPatOtrosCod = PatOtros(pnPunt).nPatOtrosCod
End Function


Public Function ObtenerPatVehMarca(ByVal pnPunt As Integer) As String
    ObtenerPatVehMarca = PatVeh(pnPunt).sPatVehMarca
End Function

'*** PEAC 20080412
Public Function ObtenerPatOtrosDescripcion(ByVal pnPunt As Integer) As String
    ObtenerPatOtrosDescripcion = PatOtros(pnPunt).sPatOtrosDescripcion
End Function

'*** PEAC 20080412
Public Function ObtenerPatInmueblesUbicacion(ByVal pnPunt As Integer) As String
    ObtenerPatInmueblesUbicacion = PatInmuebles(pnPunt).sPatInmueblesUbicacion
End Function
'*** PEAC 20080412
Public Function ObtenerPatInmueblesAreaTerreno(ByVal pnPunt As Integer) As String
    ObtenerPatInmueblesAreaTerreno = PatInmuebles(pnPunt).nPatInmueblesAreaTerreno
End Function
'*** PEAC 20080412
Public Function ObtenerPatInmueblesAreaConstru(ByVal pnPunt As Integer) As String
    ObtenerPatInmueblesAreaConstru = PatInmuebles(pnPunt).nPatInmueblesAreaConstru
End Function
'*** PEAC 20080412
Public Function ObtenerPatInmueblesTipoUso(ByVal pnPunt As Integer) As String
    ObtenerPatInmueblesTipoUso = PatInmuebles(pnPunt).sPatInmueblesTipoUso
End Function
'*** PEAC 20080412
Public Function ObtenerPatInmueblesRRPP(ByVal pnPunt As Integer) As String
    ObtenerPatInmueblesRRPP = PatInmuebles(pnPunt).sPatInmueblesRRPP
End Function
'*** PEAC 20080412
Public Function ObtenerPatInmueblesValCom(ByVal pnPunt As Integer) As String
    ObtenerPatInmueblesValCom = PatInmuebles(pnPunt).nPatInmueblesValCom
End Function
'*** PEAC 20080412
Public Function ObtenerPatInmueblesCod(ByVal pnPunt As Integer) As Integer
    ObtenerPatInmueblesCod = PatInmuebles(pnPunt).nPatInmueblesCod
End Function


'*** PEAC 20080412
Public Function ObtenerPatVehModelo(ByVal pnPunt As Integer) As String
    ObtenerPatVehModelo = PatVeh(pnPunt).sPatVehModelo
End Function
'*** PEAC 20080412
Public Function ObtenerPatVehPlaca(ByVal pnPunt As Integer) As String
    ObtenerPatVehPlaca = PatVeh(pnPunt).sPatVehPlaca
End Function

Public Function ObtenerPatVehFecFab(ByVal pnPunt As Integer) As String
    ObtenerPatVehFecFab = PatVeh(pnPunt).sPatVehFecFab
End Function

Public Function ObtenerPatVehValCom(ByVal pnPunt As Integer) As Currency
    ObtenerPatVehValCom = PatVeh(pnPunt).nPatVehValCom
End Function
'*** PEAC 20080412
Public Function ObtenerPatOtrosValCom(ByVal pnPunt As Integer) As Currency
    ObtenerPatOtrosValCom = PatOtros(pnPunt).nPatOtrosValCom
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

'*** PEAC 20080412
Public Sub ActualizarPatOtrosTipoAct(ByVal pCambio As TPersCambios, ByVal pnPunt As Integer)
    PatOtros(pnPunt).nTipoAct = pCambio
End Sub

'*** PEAC 20080412
Public Sub ActualizarPatInmueblesTipoAct(ByVal pCambio As TPersCambios, ByVal pnPunt As Integer)
    PatInmuebles(pnPunt).nTipoAct = pCambio
End Sub

'madm 20100326
Public Function ObtenerVisitaDireccion(ByVal pnPunt As Integer) As String
    ObtenerVisitaDireccion = Visita(pnPunt).sDireccion
End Function
Public Sub ActualizarVisitaDireccion(ByVal pnVisitaDireccion As String, ByVal pnPunt As Integer)
    Visita(pnPunt).sDireccion = pnVisitaDireccion
End Sub
Public Function ObtenerVisitaFecha(ByVal pnPunt As Integer) As Date
    ObtenerVisitaFecha = Visita(pnPunt).dFecVis
End Function
Public Sub ActualizarVisitaFecha(ByVal pnVisitaFecha As String, ByVal pnPunt As Integer)
    Visita(pnPunt).dFecVis = pnVisitaFecha
End Sub
Public Function ObtenerVisitaObserva(ByVal pnPunt As Integer) As String
    ObtenerVisitaObserva = Visita(pnPunt).sVisObserva
End Function
Public Sub ActualizarVisitaObserva(ByVal pnVisitaObserva As String, ByVal pnPunt As Integer)
    Visita(pnPunt).sVisObserva = pnVisitaObserva
End Sub
Public Function ObtenerVisitaPersCod(ByVal pnPunt As Integer) As String
    ObtenerVisitaPersCod = Visita(pnPunt).sPersCod
End Function
Public Sub ActualizarVisitaPersCod(ByVal pnVisitaPersCod As String, ByVal pnPunt As Integer)
    Visita(pnPunt).sPersCod = pnVisitaPersCod
End Sub
Public Function ObtenerVisitaPersCodOrig(ByVal pnPunt As Integer) As String
    ObtenerVisitaPersCodOrig = Visita(pnPunt).sPersCodOrig
End Function
Public Sub ActualizarVisitaPersCodOrig(ByVal pnVisitaPersCodOrig As String, ByVal pnPunt As Integer)
    Visita(pnPunt).sPersCodOrig = pnVisitaPersCodOrig
End Sub
Public Function ObtenerVisitaApellNombres(ByVal pnPunt As Integer) As String
    ObtenerVisitaApellNombres = Visita(pnPunt).sApellNombres
End Function
Public Sub ActualizarVisitaApellNombres(ByVal pnVisitaApellNombres As String, ByVal pnPunt As Integer)
    Visita(pnPunt).sApellNombres = pnVisitaApellNombres
End Sub

Public Function ObtenerVisitaUsual(ByVal pnPunt As Integer) As String
    ObtenerVisitaUsual = Visita(pnPunt).sUsual
End Function
Public Sub ActualizarVisitaUsual(ByVal pnVisitaUsual As String, ByVal pnPunt As Integer)
    Visita(pnPunt).sUsual = pnVisitaUsual
End Sub
Public Sub ActualizarVisitasTipoAct(ByVal pCambio As TPersCambios, ByVal pnPunt As Integer)
    Visita(pnPunt).nTipoAct = pCambio
End Sub
'MADM 20101020
Public Property Get PersRRHH() As Integer
    PersRRHH = nPersRRHH
End Property

Public Property Let PersRRHH(ByVal vNewValue As Integer)
    nPersRRHH = vNewValue
End Property
'END MADM

'MAVM 20100607 BAS II
Public Sub ActualizarVentasTipoAct(ByVal pCambio As TPersCambios, ByVal pnPunt As Integer)
    Ventas(pnPunt).nTipoAct = pCambio
End Sub

Public Function ObtenerVentasTipoAct(ByVal pnPunt As Integer) As TPersCambios
     ObtenerVentasTipoAct = Ventas(pnPunt).nTipoAct
End Function

Public Sub ActualizarVentasPersCod(ByVal pnVentasPersCod As String, ByVal pnPunt As Integer)
    Ventas(pnPunt).sPersCod = pnVentasPersCod
End Sub

Public Sub ActualizarVentasApellNombres(ByVal pnVentasApellNombres As String, ByVal pnPunt As Integer)
    Ventas(pnPunt).sApellNombres = pnVentasApellNombres
End Sub

Public Sub ActualizarVentasPeriodo(ByVal pnVentasPeriodo As String, ByVal pnPunt As Integer)
    Ventas(pnPunt).sPeriodo = pnVentasPeriodo
End Sub

Public Sub ActualizarVentasMonto(ByVal pnVentasMonto As Currency, ByVal pnPunt As Integer)
    Ventas(pnPunt).nMonto = pnVentasMonto
End Sub

Public Sub ActualizarVentasFecha(ByVal pnVentasFecha As Date, ByVal pnPunt As Integer)
    Ventas(pnPunt).dFecAud = pnVentasFecha
End Sub

Public Property Get NumeroVentasPers() As Integer
    NumeroVentasPers = nNumVentas
End Property

Public Function ObtenerVentasPersCod(ByVal pnPunt As Integer) As String
    ObtenerVentasPersCod = Ventas(pnPunt).sPersCod
End Function

Public Function ObtenerVentasApellNombres(ByVal pnPunt As Integer) As String
    ObtenerVentasApellNombres = Ventas(pnPunt).sApellNombres
End Function

Public Function ObtenerVentasPeriodo(ByVal pnPunt As Integer) As String
    ObtenerVentasPeriodo = Ventas(pnPunt).sPeriodo
End Function

Public Function ObtenerVentasMonto(ByVal pnPunt As Integer) As Currency
    ObtenerVentasMonto = Ventas(pnPunt).nMonto
End Function

Public Function ObtenerVentasFecha(ByVal pnPunt As Integer) As Date
    ObtenerVentasFecha = Ventas(pnPunt).dFecAud
End Function
'***

Public Function ObtenerVisitasTipoAct(ByVal pnPunt As Integer) As TPersCambios
     ObtenerVisitasTipoAct = Visita(pnPunt).nTipoAct
End Function

Public Property Get NumeroVisitaPers() As Integer
    NumeroVisitaPers = nNumVisitas
End Property

Public Property Let NumeroVisitaPers(ByVal vNewValue As Integer)
    nNumVisitas = vNewValue
End Property

'end madm

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

'peac 20071227
Public Function ObtenerFteIngFecEEFF(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer, Optional ByVal pnTpoFte As PersFteIngresoTipo) As Date
    If FtesIngreso(pnPunt).nTipo = gPersFteIngresoTipoDependiente Then
'        ObtenerFteIngFecEEFF = FtesIngreso(pnPunt).FtesIngresoDep(pnPuntDet).dfecEEFF
    Else
        ObtenerFteIngFecEEFF = FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).dfecEEFF
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
'**ALPA****20080412***************************************************************************
Public Function ObtenerFteIngCod(ByVal pnPunt As Integer) As String
     ObtenerFteIngCod = FtesIngreso(pnPunt).cNumFte
End Function
Public Function ObtenerFteFecCod(ByVal pnPunt As Integer) As Date
     ObtenerFteFecCod = FtesIngreso(pnPunt).dFteIngreso
End Function
Public Function ObtenerRsHojaEva(ByVal pnPunt As Integer) As ADODB.Recordset
     Set ObtenerRsHojaEva = FtesIngreso(pnPunt).RFteIng
End Function
Public Function ObtenerFteIndIngFecEval(ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     ObtenerFteIndIngFecEval = FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).dFecEval
End Function
'***********************************************************************************************
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

'*** PEAC 20080402
'Public Sub ActualizarFteIngHojEval()
'
'Dim i As Integer
'
''If Len(FEHojaEval.TextMatrix(FEHojaEval.Row, 1)) = 0 Then
''    MsgBox "No existe datos para Grabar", vbInformation, "Aviso"
''    Exit Sub
''End If
'
'Set rsHojEval = New adodb.Recordset
'
'With rsHojEval
'    'Crear RecordSet
'    .Fields.Append "cNumFuente", adVarChar, 8
'    .Fields.Append "dPersEval", adDate
'    .Fields.Append "cTipoEval", adVarChar, 50
'    .Fields.Append "cTituloEval", adVarChar, 50
'    .Fields.Append "cDescripcion", adVarChar, 50
'    .Fields.Append "cCodHojEval", adVarChar, 6
'    .Fields.Append "nMontoPersonal", adCurrency
'    .Fields.Append "nMontoNegocio", adCurrency
'    .Fields.Append "nMontoUnico", adCurrency
'    .Open
'
'    'Llenar Recordset
'    'For i = 0 To nNumRel - 1
'    For i = 0 To nPos
'        .AddNew
'        .Fields("cNumFuente") = oPersona.NumeroFtesIngreso 'FtesIngreso(pnIndex).cNumFte
'
'        .Fields("dPersEval") = "02/04/2007"
'        .Fields("cTipoEval") = MatrixHojaEval(1, i)
'        .Fields("cTituloEval") = MatrixHojaEval(2, i)
'        .Fields("cDescripcion") = MatrixHojaEval(3, i)
'        .Fields("cCodHojEval") = MatrixHojaEval(7, i)
'        .Fields("nMontoPersonal") = IIf(Len(MatrixHojaEval(4, i)) = 0, "0.00", MatrixHojaEval(4, i))
'        .Fields("nMontoNegocio") = IIf(Len(MatrixHojaEval(5, i)) = 0, "0.00", MatrixHojaEval(5, i))
'        .Fields("nMontoUnico") = IIf(Len(MatrixHojaEval(6, i)) = 0, "0.00", MatrixHojaEval(6, i))
'
'    Next i
'    If Not .EOF Then .MoveFirst
'End With
'
'End Sub


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
Public Sub ActualizarResFteIng(ByVal RFteIng As ADODB.Recordset, ByVal pnPunt As Integer)
    Set FtesIngreso(pnPunt).RFteIng = RFteIng
End Sub
'WIOR 20140321 *************************************
Public Sub ActualizarRsEstFinan(ByVal rsEstFinan As ADODB.Recordset, ByVal pnPunt As Integer)
    Set FtesIngreso(pnPunt).rsEstFinan = rsEstFinan
End Sub
'WIOR FIN ******************************************
Public Sub ActualizarFteRSCInd(ByVal pnIndicador As Integer, ByVal pnPunt As Integer)
     FtesIngreso(pnPunt).nFtePaRscInd = pnIndicador
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

'peac 20071227 eeff
Public Sub ActualizarFteIngFecEEFF(ByVal pdFecEEFF As Date, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).dfecEEFF = pdFecEEFF
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

'peac 20071227
Public Sub ActualizarFteIndIngFecEEFF(ByVal pdFecEEFF As Date, ByVal pnPunt As Integer, ByVal pnPuntDet As Integer)
     FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).dfecEEFF = pdFecEEFF
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
'EJVG20120818 ***
Public Function ObtenerCodPersonaRelacionado(ByVal pnRelacion As Integer) As String
    Dim i As Integer
    ObtenerCodPersonaRelacionado = ""
    For i = 0 To nNumRel - 1
        If Trim(Right(Relac(i).sRelacion, 3)) = pnRelacion Then
            ObtenerCodPersonaRelacionado = Relac(i).sPersCod
            Exit Function
        End If
    Next i
End Function
'END EJVG *******
Public Sub ActualizaRefComNombre(ByVal psNombre As String, ByVal pnPunt As Integer)
    RefCom(pnPunt).sRefComNom = psNombre
End Sub

Public Sub ActualizaRefComTipoRel(ByVal psTipoRel As String, ByVal pnPunt As Integer)
    RefCom(pnPunt).sRefComTipo = psTipoRel
End Sub

Public Sub ActualizaRefComFono(ByVal psFono As String, ByVal pnPunt As Integer)
    RefCom(pnPunt).sRefComFono = psFono
End Sub

'*** PEAC 20080412
Public Sub ActualizaRefComComentario(ByVal psComentario As String, ByVal pnPunt As Integer)
    RefCom(pnPunt).sRefComComentario = psComentario
End Sub

'*** PEAC 20080412
Public Sub ActualizaRefComDireccion(ByVal psDireccion As String, ByVal pnPunt As Integer)
    RefCom(pnPunt).sRefComDireccion = psDireccion
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

'*** PEAC 20080412
Public Sub ActualizaPatOtrosCod(ByVal pnCod As Integer, ByVal pnPunt As Integer)
    PatOtros(pnPunt).nPatOtrosCod = pnCod
End Sub

'*** PEAC 20080412
Public Sub ActualizaPatInmueblesCod(ByVal pnCod As Integer, ByVal pnPunt As Integer)
    PatInmuebles(pnPunt).nPatInmueblesCod = pnCod
End Sub

Public Sub ActualizaPatVehMarca(ByVal psMarca As String, ByVal pnPunt As Integer)
    PatVeh(pnPunt).sPatVehMarca = psMarca
End Sub

'*** PEAC 20080412
Public Sub ActualizaPatOtrosDescripcion(ByVal psDescripcion As String, ByVal pnPunt As Integer)
    PatOtros(pnPunt).sPatOtrosDescripcion = psDescripcion
End Sub

'*** PEAC 20080412
Public Sub ActualizaPatInmueblesUbicacion(ByVal psUbicacion As String, ByVal pnPunt As Integer)
    PatInmuebles(pnPunt).sPatInmueblesUbicacion = psUbicacion
End Sub
'*** PEAC 20080412
Public Sub ActualizaPatInmueblesAreaTerreno(ByVal pnAreaTerreno As Double, ByVal pnPunt As Integer)
    PatInmuebles(pnPunt).nPatInmueblesAreaTerreno = pnAreaTerreno
End Sub
'*** PEAC 20080412
Public Sub ActualizaPatInmueblesAreaConstru(ByVal pnAreaConstru As Double, ByVal pnPunt As Integer)
    PatInmuebles(pnPunt).nPatInmueblesAreaConstru = pnAreaConstru
End Sub
'*** PEAC 20080412
Public Sub ActualizaPatInmueblesTipoUso(ByVal psTipoUso As String, ByVal pnPunt As Integer)
    PatInmuebles(pnPunt).sPatInmueblesTipoUso = psTipoUso
End Sub
'*** PEAC 20080412
Public Sub ActualizaPatInmueblesRRPP(ByVal psRRPP As String, ByVal pnPunt As Integer)
    PatInmuebles(pnPunt).sPatInmueblesRRPP = psRRPP
End Sub
'*** PEAC 20080412
Public Sub ActualizaPatInmueblesValCom(ByVal pnValCom As Currency, ByVal pnPunt As Integer)
    PatInmuebles(pnPunt).nPatInmueblesValCom = pnValCom
End Sub

Public Sub ActualizaPatVehFecFab(ByVal psFecFab As String, ByVal pnPunt As Integer)
    PatVeh(pnPunt).sPatVehFecFab = psFecFab
End Sub

Public Sub ActualizaPatVehValCom(ByVal pnValCom As Currency, ByVal pnPunt As Integer)
    PatVeh(pnPunt).nPatVehValCom = pnValCom
End Sub
'*** PEAC 20080412
Public Sub ActualizaPatOtrosValCom(ByVal pnValCom As Currency, ByVal pnPunt As Integer)
    PatOtros(pnPunt).nPatOtrosValCom = pnValCom
End Sub


'*** PEAC 20080412
Public Sub ActualizaPatVehModelo(ByVal psModelo As String, ByVal pnPunt As Integer)
    PatVeh(pnPunt).sPatVehModelo = psModelo
End Sub
'*** PEAC 20080412
Public Sub ActualizaPatVehPlaca(ByVal psPlaca As String, ByVal pnPunt As Integer)
    PatVeh(pnPunt).sPatVehPlaca = psPlaca
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


Sub LlenaRecordSet_Persona()
Set rsPersona = New ADODB.Recordset

With rsPersona
    'Crear RecordSet
    .Fields.Append "cPersNombre", adVarChar, 200
    .Fields.Append "dPersNacCreac", adDate
    .Fields.Append "dPersFallec", adDate 'MAVM 06042009
    
    '*** MADM 20091114
    .Fields.Append "nCredito", adSmallInt
    .Fields.Append "nAhorro", adSmallInt
    .Fields.Append "nServicio", adSmallInt
    .Fields.Append "nOtro", adSmallInt '*** end MADM 20091114
    .Fields.Append "nPeps", adSmallInt '*** end MADM 20101408
    .Fields.Append "nSujetoObligado", adInteger 'EJVG20120813
    .Fields.Append "nOfCumplimiento", adInteger 'EJVG20120813
    .Fields.Append "nAutorizaUsoDatos", adInteger 'FRHU 20151130 ERS077-2015
    .Fields.Append "sActiGiro1", adVarChar, 200 'madm 20100305
    .Fields.Append "nPersCargo", adVarChar, 200 '** Juez 20120326
      
    .Fields.Append "cPersDireccUbiGeo", adVarChar, 12
    .Fields.Append "cPersDireccDomicilio", adVarChar, 100
    
    .Fields.Append "cPersRefDomicilio", adVarChar, 100 '*** PEAC 20080801
    
    .Fields.Append "cPersDireccCondicion", adVarChar, 1
    .Fields.Append "cnPersValComDomicilio", adCurrency
    'JUEZ 20131007 ************************************
    .Fields.Append "cPersNegocioUbiGeo", adVarChar, 12
    .Fields.Append "cPersNegocioDireccion", adVarChar, 100
    .Fields.Append "cPersNegocioRef", adVarChar, 200
    .Fields.Append "cPersNombreCentroLaboral", adVarChar, 200
    .Fields.Append "nRemisionInfoEmail", adInteger
    'END JUEZ *****************************************
    .Fields.Append "cPersTelefono", adVarChar, 100
    .Fields.Append "cPersTelefono2", adVarChar, 100
    'EJVG20110712 *************************************
    .Fields.Append "cPersCelular", adVarChar, 100
    .Fields.Append "cPersCelular2", adVarChar, 100
    .Fields.Append "cPersCelular3", adVarChar, 100
    
    .Fields.Append "cPersEmail", adVarChar, 50
    .Fields.Append "cPersEmail2", adVarChar, 50 'EJVG20110712
    '*** PEAC 20080412
    .Fields.Append "dPersIngRuc", adDate
    .Fields.Append "dPersIniActi", adDate
    
    .Fields.Append "nNumDependi", adInteger
    .Fields.Append "cActComple", adVarChar, 300 'LUCV20181220, (cambio 50 por 300)
    .Fields.Append "nNumPtosVta", adInteger
    .Fields.Append "cActiGiro", adVarChar, 100  '*** PEAC 20080617
    .Fields.Append "nPersPersoneria", adInteger
    .Fields.Append "cPersCIIU", adVarChar, 7
    .Fields.Append "nPersTipoComp", adInteger
    .Fields.Append "nPersTipoSistInfor", adInteger
    .Fields.Append "nMonedaPatri", adInteger
    
    .Fields.Append "nPersCadenaProd", adInteger
    '*** FIN PEAC
    .Fields.Append "cPersEstado", adVarChar, 2
    .Fields.Append "cPersCodSbs", adChar, 10
    .Fields.Append "nPersRelaInst", adInteger
    .Fields.Append "cPersnatSexo", adVarChar, 1
    .Fields.Append "nPersNatEstCiv", adInteger
    .Fields.Append "nPersNatHijos", adInteger
    '*** PEAC 20080412
    .Fields.Append "nPersNatNumEmp", adInteger
    
    .Fields.Append "nPersPeso", adDouble
    .Fields.Append "nPersTalla", adDouble
    .Fields.Append "nPersTpoSangre", adInteger
    .Fields.Append "cNacionalidad", adVarChar, 12
    
    '*** PEAC 20080801
    .Fields.Append "cMotivoActu", adVarChar, 2
    
    .Fields.Append "nResidente", adSmallInt
    .Fields.Append "cPaisReside", adVarChar, 12 '** Juez 20120327
    .Fields.Append "cUltimaActualizacion", adVarChar, 25
    .Fields.Append "cPersJurTpo", adVarChar, 4
    .Fields.Append "cPersJurMagnitud", adVarChar, 2
    .Fields.Append "cPersJurObjetoSocial", adVarChar, 200 '** Juez 20120328
    .Fields.Append "cPersNatMagnitud", adVarChar, 2 'JACA 20110427
    .Fields.Append "cPersJurSigla", adVarChar, 100
    .Fields.Append "nPersJurEmpleados", adInteger
    'CUSCO
    .Fields.Append "nPersIngresoProm", adCurrency
    'Para el tema de las Validaciones
    '.Fields.Append "cApellidoPaterno", adVarChar, 30
    '.Fields.Append "cApellidoMaterno", adVarChar, 30
    '.Fields.Append "cNombreCompleto", adVarChar, 200
    .Open
    
    'Llenar Recordset
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

    .AddNew
    .Fields("cPersNombre") = RTrim(sNombreCompleto)
    .Fields("dPersNacCreac") = dFechaNac
    .Fields("dPersFallec") = dFechaFallec 'MAVM 03042009
    'MADM 20091114
    .Fields("nCredito") = sCredito
    .Fields("nAhorro") = sAhorro
    .Fields("nServicio") = sServicio
    .Fields("nOtro") = sOtro
    .Fields("nPeps") = sPeps '20100408
    'END MADM
    .Fields("nSujetoObligado") = nSujetoObligado 'EJVG20120813
    .Fields("nOfCumplimiento") = nOfCumplimiento 'EJVG20120813
    .Fields("nAutorizaUsoDatos") = nAutorizaUsoDatos 'FRHU 20151130 ERS077-2015
    .Fields("sActiGiro1") = sActiGiro1 'madm 20100305
    .Fields("nPersCargo") = nPersCargo '** Juez 20120326
    .Fields("cPersDireccUbiGeo") = sUbicGeografica
    .Fields("cPersDireccDomicilio") = sDomicilio
    .Fields("cPersRefDomicilio") = sRefDomicilio '*** PEAC 20080801
    .Fields("cPersDireccCondicion") = sCondicionDomic
    .Fields("cnPersValComDomicilio") = nValComDomicilio
    'JUEZ 20131007 *******************************
    .Fields("cPersNegocioUbiGeo") = sUbiGeoNegocio
    .Fields("cPersNegocioDireccion") = sNegocioDireccion
    .Fields("cPersNegocioRef") = sRefNegocio
    .Fields("cPersNombreCentroLaboral") = sNombreCentroLaboral
    .Fields("nRemisionInfoEmail") = nRemisionInfoEmail
    'END JUEZ ************************************
    .Fields("cPersTelefono") = sTelefonos
    .Fields("cPersTelefono2") = sTelefonos2
    'EJVG20110712 *************************
    .Fields("cPersCelular") = sCelular
    .Fields("cPersCelular2") = sCelular2
    .Fields("cPersCelular3") = sCelular3
    
    .Fields("cPersEmail") = sEmail
    .Fields("cPersEmail2") = sEmail2 'EJVG20111207
    '*** PEAC 20080412
    .Fields("dPersIngRuc") = dFechaInscRuc
    .Fields("dPersIniActi") = dFechaIniActi
    
    '***CIT3 (04set2018)
    '.Fields("nNumDependi") = nNumDependi
    .Fields("nNumDependi") = Trim((nHijos))
    '***
    
    .Fields("cActComple") = sActComple
    .Fields("nNumPtosVta") = nNumPtosVta
    .Fields("cActiGiro") = sActiGiro
    .Fields("nPersTipoComp") = nTipoCompDescripcion
    .Fields("nPersTipoSistInfor") = nTipoSistInforDescripcion
    .Fields("nMonedaPatri") = IIf(Len(nMonedaPatri) > 0, nMonedaPatri, 0)
    .Fields("nPersCadenaProd") = nCadenaProdDescripcion
    '*** FIN PEAC
    .Fields("nPersPersoneria") = Trim((nTipoPersona))
    .Fields("cPersCIIU") = sCiiu
    .Fields("cPersEstado") = sEstado
    .Fields("cPersCodSbs") = sPersCodSbs
    .Fields("nPersRelaInst") = nPersRelInst
    .Fields("cPersnatSexo") = sSexo
    .Fields("nPersNatEstCiv") = IIf(sEstadoCivil = "", 99, sEstadoCivil)
    .Fields("nPersNatHijos") = Trim((nHijos))
    '*** PEAC 20080412
    .Fields("nPersNatNumEmp") = Trim((nNumEmp))
    
    .Fields("nPersPeso") = Format(nPeso, "#0.00")
    .Fields("nPersTalla") = Format(nTalla, "#0.00")
    .Fields("nPersTpoSangre") = IIf(sTipoSangre = "", 0, sTipoSangre)
    .Fields("cNacionalidad") = cNacionalidad
    '*** PEAC 20080801
    .Fields("cMotivoActu") = cMotivoActu
    
    .Fields("nResidente") = nResidencia
    .Fields("cPaisReside") = cPaisReside '** Juez 20120327
    .Fields("cUltimaActualizacion") = sActualiza
    .Fields("cPersJurTpo") = sTipoPersJur
    .Fields("cPersJurMagnitud") = sMagnitudEmp
    .Fields("cPersJurObjetoSocial") = sObjetoSocial '** Juez 20120328
    .Fields("cPersNatMagnitud") = sMagnitudPersNat 'JACA 20110427
    .Fields("cPersJurSigla") = sSiglas
    .Fields("nPersJurEmpleados") = Trim((nNumEmplead))
    'CUSCO
    .Fields("nPersIngresoProm") = nIngresoProm
    
    'Para las Validaciones
    '.Fields("cApellidoPaterno") = ApellidoPaterno
    '.Fields("cApellidoMaterno") = ApellidoMaterno
    '.Fields("cNombreCompleto") = NombreCompleto
End With
              
End Sub
'MADM 20100326
Sub LlenaRecordSet_Visitas()
Set rsVisita = New ADODB.Recordset

With rsVisita
    'Crear RecordSet
    .Fields.Append "cPersCodOrig", adVarChar, 13
    .Fields.Append "cPersRelacPersCod", adVarChar, 13
    '.Fields.Append "cDireccion", adVarChar, 100
    .Fields.Append "cDireccion", adVarChar, 150 'FRHU INCIDENTE 20150414
    .Fields.Append "dFecVis", adDate
    '.Fields.Append "cVisObserva", adVarChar, 150
    .Fields.Append "cVisObserva", adVarChar, 200 'FRHU INCIDENTE 20150414
    .Fields.Append "nUsual", adInteger
    .Fields.Append "nTipoAct", adInteger
    .Open

    'Llenar Recordset
    For i = 0 To nNumVisitas - 1
        .AddNew
        .Fields("cPersRelacPersCod") = Visita(i).sPersCod
        .Fields("cPersCodOrig") = Visita(i).sPersCodOrig
        '.Fields("cDireccion") = Visita(i).sDireccion
        .Fields("cDireccion") = Trim(Left(Visita(i).sDireccion, 150)) 'FRHU INCIDENTE 20150414
        .Fields("dFecVis") = Visita(i).dFecVis
        .Fields("nUsual") = Trim(Right(Visita(i).sUsual, 20))
        '.Fields("cVisObserva") = Visita(i).sVisObserva
        .Fields("cVisObserva") = Trim(Left(Visita(i).sVisObserva, 200)) 'FRHU INCIDENTE 20150414
        .Fields("nTipoAct") = Visita(i).nTipoAct
    Next i
    If Not .EOF Then .MoveFirst
End With
End Sub
'END MADM

'MAVM 20100607 BAS II
Sub LlenaRecordSet_Ventas()
Set rsVentas = New ADODB.Recordset

With rsVentas
    'Crear RecordSet
    .Fields.Append "cPersCodOrigVA", adVarChar, 13
    .Fields.Append "cPersCodVA", adVarChar, 13
    .Fields.Append "cPeriodoVA", adVarChar, 100
    .Fields.Append "nMontoVA", adCurrency
    .Fields.Append "dFecAudVA", adDate
    .Fields.Append "nTipoActVA", adInteger
    .Open

    'Llenar Recordset
    For i = 0 To nNumVentas - 1
        .AddNew
        .Fields("cPersCodOrigVA") = Ventas(i).sPersCodOrigVA
        .Fields("cPersCodVA") = Ventas(i).sPersCod
        .Fields("cPeriodoVA") = Ventas(i).sPeriodo
        .Fields("nMontoVA") = Ventas(i).nMonto
        .Fields("dFecAudVA") = Ventas(i).dFecAud
        .Fields("nTipoActVA") = Ventas(i).nTipoAct
    Next i
    If Not .EOF Then .MoveFirst
End With
End Sub
'***

Sub LlenaRecordSet_Relaciones()
Set rsRelac = New ADODB.Recordset

With rsRelac
    'Crear RecordSet
    .Fields.Append "cPersRelacPersCod", adVarChar, 13
    .Fields.Append "nPersRelac", adInteger
    .Fields.Append "bPersRelacBenef", adInteger
    .Fields.Append "nPersRelacBenefPorc", adDouble
    .Fields.Append "bPersRelacAMP", adInteger
    .Fields.Append "cPersCodOrig", adVarChar, 13
    .Fields.Append "cMovNro", adVarChar, 25
    .Fields.Append "nTipoAct", adInteger
    .Open

    'Llenar Recordset
    For i = 0 To nNumRel - 1
        .AddNew
        .Fields("cPersRelacPersCod") = Relac(i).sPersCod
        .Fields("nPersRelac") = Trim(Right(Relac(i).sRelacion, 20))
        .Fields("bPersRelacBenef") = Trim(Right(Relac(i).sPersRelacBenef, 20))
        .Fields("nPersRelacBenefPorc") = Format(Relac(i).nPersRelacBenefPorc, "#0.00")
        .Fields("bPersRelacAMP") = Trim(Right(Relac(i).sPersRelacAMP, 20))
        .Fields("cPersCodOrig") = Relac(i).sPersCodOrig
        .Fields("cMovNro") = sActualiza
        .Fields("nTipoAct") = Relac(i).nTipoAct
    Next i
    If Not .EOF Then .MoveFirst
End With
End Sub
Sub LlenaRecordSet_HojaEvaluacion()
    MsgBox "Se puede llamar ", vbExclamation, "Aviso"
End Sub
Sub LlenaRecordSet_Documentos()

Set rsDocs = New ADODB.Recordset

With rsDocs
    'Crear RecordSet
    .Fields.Append "cPersIdTpo", adInteger
    .Fields.Append "cPersIDNro", adVarChar, 20
    .Fields.Append "cUltimaActualizacion", adVarChar, 25
    .Fields.Append "nTipoAct", adInteger
    .Fields.Append "nTipoOrig", adInteger
    .Open
    
    'Llenar Recordset
    For i = 0 To nNumDoc - 1
        .AddNew
        .Fields("cPersIdTpo") = Trim((Docs(i).nTipo))
        .Fields("cPersIDNro") = Docs(i).sNumero
        .Fields("cUltimaActualizacion") = sActualiza
        .Fields("nTipoOrig") = Trim((Docs(i).nTipoOrig))
        .Fields("nTipoAct") = Docs(i).nTipoAct
    Next i
    If Not .EOF Then .MoveFirst
End With

End Sub

Sub LlenaRecordSet_RefComeciales()
Set rsRefCom = New ADODB.Recordset

With rsRefCom
    'Crear RecordSet
    .Fields.Append "nNumRefCom", adSmallInt
    .Fields.Append "cNomRefCom", adVarChar, 50
    .Fields.Append "nTipoRelaCom", adSmallInt
    .Fields.Append "cFonoRefCom", adVarChar, 12
    .Fields.Append "cComentarioRefCom", adVarChar, 100
    .Fields.Append "cDireccionRefCom", adVarChar, 50
    .Fields.Append "cUltimaActualizacion", adVarChar, 25
    .Fields.Append "nTipoAct", adInteger
    .Fields.Append "nRefComCod", adSmallInt
    .Open
    
    'Llenar Recordset
    For i = 0 To nNumRefCom - 1
        .AddNew
        .Fields("nNumRefCom") = RefCom(i).nRefComCod
        .Fields("cNomRefCom") = RefCom(i).sRefComNom
        .Fields("nTipoRelaCom") = Right(RefCom(i).sRefComTipo, 1)
        .Fields("cFonoRefCom") = RefCom(i).sRefComFono
        .Fields("cComentarioRefCom") = RefCom(i).sRefComComentario
        .Fields("cDireccionRefCom") = RefCom(i).sRefComDireccion
        .Fields("cUltimaActualizacion") = sActualiza
        .Fields("nTipoAct") = RefCom(i).nTipoAct
        .Fields("nRefComCod") = RefCom(i).nRefComCod
    Next i
    
    If Not .EOF Then .MoveFirst
End With
                
End Sub

Sub LlenaRecordset_RefBancarias()

Set rsRefBan = New ADODB.Recordset

With rsRefBan
    'Crear RecordSet
    .Fields.Append "cCodBanco", adVarChar, 13
    .Fields.Append "cNumCta", adVarChar, 20
    .Fields.Append "cNumTarjeta", adVarChar, 20
    .Fields.Append "nMontoLinCred", adCurrency
    .Fields.Append "cUltimaActualizacion", adVarChar, 25
    .Fields.Append "nTipoAct", adInteger
    .Fields.Append "sRefBanCodIFOrg", adVarChar, 13
    .Open
    'Llenar Recordset
    For i = 0 To nNumRefBan - 1
        .AddNew
        .Fields("cCodBanco") = Mid(RefBan(i).sRefBanCodIF, 4, Len(RefBan(i).sRefBanCodIF) - 3)
        .Fields("cNumCta") = RefBan(i).sRefBanNumCta
        .Fields("cNumTarjeta") = RefBan(i).sRefBanNumTar
        .Fields("nMontoLinCred") = RefBan(i).nRefBanLinCred
        .Fields("cUltimaActualizacion") = sActualiza
        .Fields("sRefBanCodIFOrg") = Mid(RefBan(i).sRefBanCodIF, 4, Len(RefBan(i).sRefBanCodIF) - 3) 'Mid(RefBan(i).sRefBanCodIFOrg, 4, Len(RefBan(i).sRefBanCodIFOrg) - 3)
        .Fields("nTipoAct") = RefBan(i).nTipoAct
    Next i
    
    If Not .EOF Then .MoveFirst
End With

End Sub

Sub LlenaRecordset_PatVehicular()

Set rsPatVeh = New ADODB.Recordset

With rsPatVeh
    'Crear RecordSet
    .Fields.Append "nNumPatri", adInteger
    .Fields.Append "nCondPatri", adSmallInt
    .Fields.Append "cMarca", adVarChar, 50
    .Fields.Append "cModelo", adVarChar, 20
    .Fields.Append "cFecFabrica", adVarChar, 4
    .Fields.Append "cPlaca", adVarChar, 20
    .Fields.Append "nValorCom", adCurrency
    .Fields.Append "nTipoAct", adInteger
    .Fields.Append "cUltimaActualizacion", adVarChar, 13
    .Open
    'Llenar Recordset
    For i = 0 To nNumPatVeh - 1
        .AddNew
        .Fields("nNumPatri") = PatVeh(i).nPatVehCod
        .Fields("nCondPatri") = 1 'Right(PatVeh(i).sPatVehCondicion, 1)
        .Fields("cMarca") = PatVeh(i).sPatVehMarca
        .Fields("cModelo") = PatVeh(i).sPatVehModelo
        .Fields("cFecFabrica") = PatVeh(i).sPatVehFecFab
        .Fields("cPlaca") = PatVeh(i).sPatVehPlaca
        .Fields("nValorCom") = PatVeh(i).nPatVehValCom
        .Fields("nTipoAct") = PatVeh(i).nTipoAct
        .Fields("cUltimaActualizacion") = sActualiza
    Next i
    
    If Not .EOF Then .MoveFirst
End With

End Sub

'*** PEAC 20080412
Sub LlenaRecordset_PatOtros()

Set rsPatOtros = New ADODB.Recordset

With rsPatOtros
    'Crear RecordSet
    .Fields.Append "nNumPatri", adInteger
    .Fields.Append "cDescripcion", adVarChar, 50
    .Fields.Append "nValorCom", adCurrency
    .Fields.Append "nTipoAct", adInteger
    .Fields.Append "cUltimaActualizacion", adVarChar, 13
    .Open
    'Llenar Recordset
    For i = 0 To nNumPatOtros - 1
        .AddNew
        .Fields("nNumPatri") = PatOtros(i).nPatOtrosCod
        .Fields("cDescripcion") = PatOtros(i).sPatOtrosDescripcion
        .Fields("nValorCom") = PatOtros(i).nPatOtrosValCom
        .Fields("nTipoAct") = PatOtros(i).nTipoAct
        .Fields("cUltimaActualizacion") = sActualiza
    Next i
    
    If Not .EOF Then .MoveFirst
End With

End Sub

'*** PEAC 20080412
Sub LlenaRecordset_PatInmuebles()

Set rsPatInmuebles = New ADODB.Recordset

With rsPatInmuebles
    'Crear RecordSet
    .Fields.Append "nNumPatri", adInteger
    .Fields.Append "cUbicacion", adVarChar, 50
    .Fields.Append "nAreaTerreno", adDouble
    .Fields.Append "nAreaConstru", adDouble
    .Fields.Append "cTipoUso", adVarChar, 20
    .Fields.Append "cRRPP", adVarChar, 5
    .Fields.Append "nValorCom", adCurrency
    .Fields.Append "nTipoAct", adInteger
    .Fields.Append "cUltimaActualizacion", adVarChar, 13
    .Open
    'Llenar Recordset
    For i = 0 To nNumPatInmuebles - 1
        .AddNew
        .Fields("nNumPatri") = PatInmuebles(i).nPatInmueblesCod
        .Fields("cUbicacion") = PatInmuebles(i).sPatInmueblesUbicacion
        .Fields("nAreaTerreno") = PatInmuebles(i).nPatInmueblesAreaTerreno
        .Fields("nAreaConstru") = PatInmuebles(i).nPatInmueblesAreaConstru
        .Fields("cTipoUso") = PatInmuebles(i).sPatInmueblesTipoUso
        .Fields("cRRPP") = PatInmuebles(i).sPatInmueblesRRPP
        .Fields("nValorCom") = PatInmuebles(i).nPatInmueblesValCom
        .Fields("nTipoAct") = PatInmuebles(i).nTipoAct
        .Fields("cUltimaActualizacion") = sActualiza
    Next i
    
    If Not .EOF Then .MoveFirst
End With

End Sub

Sub LlenaRecordset_FteIngDep()
Set rsFIDep = New ADODB.Recordset
nIntValor = 0
With rsFIDep
    'Crear RecordSet
    .Fields.Append "cNumFuente", adChar, 8
    .Fields.Append "dPersEval", adDate
    .Fields.Append "nPersIngCli", adCurrency
    .Fields.Append "nPersGastoFam", adCurrency
    .Fields.Append "nPersIngCon", adCurrency
    .Fields.Append "nPersOtrIng", adCurrency
    .Fields.Append "dPersFICaduca", adDate
    .Fields.Append "nTipoAct", adInteger
    .Open
    
    Set rsHojaEval = Nothing
    Set rsEstFinan = Nothing 'WIOR 20140321
    'Llenar Recordset
    For i = 0 To nNumFtes - 1
        For j = 0 To FtesIngreso(i).nNumFtesDep - 1
            .AddNew
            .Fields("cNumFuente") = FtesIngreso(i).cNumFte
            If FtesIngreso(i).FtesIngresoDep(j).dFecEval <> 0 Then
              .Fields("dPersEval") = FtesIngreso(i).FtesIngresoDep(j).dFecEval
            Else
              .Fields("dPersEval") = FtesIngreso(i).dFtePersFecInicioAct
            End If
            .Fields("nPersIngCli") = Format(FtesIngreso(i).FtesIngresoDep(j).nIngresos, "#0.00")
            .Fields("nPersGastoFam") = Format(FtesIngreso(i).FtesIngresoDep(j).nGastosFam, "#0.00")
            .Fields("nPersIngCon") = Format(FtesIngreso(i).FtesIngresoDep(j).nIngCony, "#0.00")
            .Fields("nPersOtrIng") = Format(FtesIngreso(i).FtesIngresoDep(j).nIngOtros, "#0.00")
            .Fields("dPersFICaduca") = FtesIngreso(i).FtesIngresoDep(j).dFecCad
            .Fields("nTipoAct") = FtesIngreso(i).FtesIngresoDep(j).nTipoAct
            If FtesIngreso(i).nFtePaRscInd = 1 And nIntValor = 0 Then
                Set rsHojaEval = FtesIngreso(i).RFteIng
                Set rsEstFinan = FtesIngreso(i).rsEstFinan 'WIOR 20140321
                nIntValor = 1
            End If
            
        Next j
    Next i
    If Not .EOF Then .MoveFirst
End With

End Sub

Sub LlenaRecordset_FteCosto()

Set rsFICosto = New ADODB.Recordset

With rsFICosto
    'Crear RecordSet
    .Fields.Append "cNumFuente", adChar, 8
    .Fields.Append "dFecEval", adDate
    .Fields.Append "nTpoCultivo", adSmallInt
    .Fields.Append "nMaquinaria", adCurrency
    .Fields.Append "nJornales", adCurrency
    .Fields.Append "nInsumos", adCurrency
    .Fields.Append "nPesticidas", adCurrency
    .Fields.Append "nOtros", adCurrency
    .Fields.Append "nHectareas", adDouble
    .Fields.Append "nProduccion", adCurrency
    .Fields.Append "nUniProd", adSmallInt
    .Fields.Append "nPreUni", adCurrency
    .Fields.Append "nSiembra", adInteger
    .Fields.Append "nCosecha", adInteger
    .Fields.Append "nDesaAgricola", adInteger
    .Fields.Append "nPlanOtros", adInteger
    .Fields.Append "nMantenimiento", adInteger
    .Fields.Append "nTipoAct", adInteger
    .Open
    
    'Llenar Recordset
    For i = 0 To nNumFtes - 1
        For j = 0 To FtesIngreso(i).nNumCostosProd - 1
            .AddNew
            .Fields("cNumFuente") = FtesIngreso(i).cNumFte
            If FtesIngreso(i).CostoProd(j).dFecEval <> 0 Then
                .Fields("dFecEval") = FtesIngreso(i).CostoProd(j).dFecEval
            Else
                .Fields("dPersEval") = FtesIngreso(i).dFtePersFecInicioAct
            End If
            .Fields("nTpoCultivo") = FtesIngreso(i).CostoProd(j).nTpoCultivo
            .Fields("nMaquinaria") = Format(FtesIngreso(i).CostoProd(j).nMaquinaria, "#0.00")
            .Fields("nJornales") = Format(FtesIngreso(i).CostoProd(j).nJornales, "#0.00")
            .Fields("nInsumos") = Format(FtesIngreso(i).CostoProd(j).nInsumos, "#0.00")
            .Fields("nPesticidas") = Format(FtesIngreso(i).CostoProd(j).nPesticidas, "#0.00")
            .Fields("nOtros") = Format(FtesIngreso(i).CostoProd(j).nOtros, "#0.00")
            .Fields("nHectareas") = Format(FtesIngreso(i).CostoProd(j).nHectareas, "#0.00")
            .Fields("nProduccion") = Format(FtesIngreso(i).CostoProd(j).nProduccion, "#0.00")
            .Fields("nUniProd") = Format(FtesIngreso(i).CostoProd(j).nUniProd, "#0.00")
            .Fields("nPreUni") = Format(FtesIngreso(i).CostoProd(j).nPreUni, "#0.00")
            .Fields("nSiembra") = FtesIngreso(i).CostoProd(j).nSiembra
            .Fields("nCosecha") = FtesIngreso(i).CostoProd(j).nCosecha
            .Fields("nDesaAgricola") = FtesIngreso(i).CostoProd(j).nDesAgricola
            .Fields("nPlanOtros") = FtesIngreso(i).CostoProd(j).nPlanOtros
            .Fields("nMantenimiento") = FtesIngreso(i).CostoProd(j).nMantenimiento
            .Fields("nTipoAct") = FtesIngreso(i).CostoProd(j).nTipoAct
            
        Next j
    Next i
    
    If Not .EOF Then .MoveFirst
End With
 
End Sub

Sub LlenaRecordset_FteIngInd()

Set rsFIInd = New ADODB.Recordset

With rsFIInd
    'Crear RecordSet
    .Fields.Append "cNumFuente", adChar, 8
    .Fields.Append "dPersEval", adDate
    
    'peac 20071227
    .Fields.Append "dFecEEFF", adDate
       
    .Fields.Append "dPersFICaduca", adDate
    .Fields.Append "nPersFIActivoDisp", adCurrency
    .Fields.Append "nPersFICtasxCobrar", adCurrency
    .Fields.Append "nPersFIInventarios", adCurrency
    .Fields.Append "nPersFIActivosFijos", adCurrency
    .Fields.Append "nPersFIProveedores", adCurrency
    .Fields.Append "nPersFICredCMACT", adCurrency
    .Fields.Append "nPersFICredOtros", adCurrency
    .Fields.Append "nPersFIVentas", adCurrency
    .Fields.Append "nPersFIPatrimonio", adCurrency
    .Fields.Append "nPersFICostoVentas", adCurrency
    .Fields.Append "nPersFIRecupCtasXCobrar", adCurrency
    .Fields.Append "nPersFIEgresosOtros", adCurrency
    .Fields.Append "nPersIngFam", adCurrency
    .Fields.Append "nPersEgrFam", adCurrency
    .Fields.Append "cUltimaActualizacion", adVarChar, 25
    .Fields.Append "nTipoAct", adInteger
    .Open
     
   'Llenar Recordset
    For i = 0 To nNumFtes - 1
        For j = 0 To FtesIngreso(i).nNumFtesIndep - 1
            .AddNew
            .Fields("cNumFuente") = FtesIngreso(i).cNumFte
            If FtesIngreso(i).FtesIngresoInd(j).dFecEval <> 0 Then
                .Fields("dPersEval") = FtesIngreso(i).FtesIngresoInd(j).dFecEval
            Else
                .Fields("dPersEval") = FtesIngreso(i).dFtePersFecInicioAct
            End If
            
            'peac 20071227
            .Fields("dFecEEFF") = FtesIngreso(i).FtesIngresoInd(j).dfecEEFF

            .Fields("dPersFICaduca") = FtesIngreso(i).FtesIngresoInd(j).dFecCad
            .Fields("nPersFIActivoDisp") = Format(FtesIngreso(i).FtesIngresoInd(j).nActivDisp, "#0.00")
            .Fields("nPersFICtasxCobrar") = Format(FtesIngreso(i).FtesIngresoInd(j).nCtasxCob, "#0.00")
            .Fields("nPersFIInventarios") = Format(FtesIngreso(i).FtesIngresoInd(j).nInventarios, "#0.00")
            .Fields("nPersFIActivosFijos") = Format(FtesIngreso(i).FtesIngresoInd(j).nActFijo, "#0.00")
            .Fields("nPersFIProveedores") = Format(FtesIngreso(i).FtesIngresoInd(j).nProveedores, "#0.00")
            .Fields("nPersFICredCMACT") = Format(FtesIngreso(i).FtesIngresoInd(j).nCreditosCmact, "#0.00")
            .Fields("nPersFICredOtros") = Format(FtesIngreso(i).FtesIngresoInd(j).nCreditosOtros, "#0.00")
            .Fields("nPersFIVentas") = Format(FtesIngreso(i).FtesIngresoInd(j).nVentas, "#0.00")
            .Fields("nPersFIPatrimonio") = Format(FtesIngreso(i).FtesIngresoInd(j).nPasivoPatrimonio, "#0.00")
            .Fields("nPersFICostoVentas") = Format(FtesIngreso(i).FtesIngresoInd(j).nCompraMercad, "#0.00")
            .Fields("nPersFIRecupCtasXCobrar") = Format(FtesIngreso(i).FtesIngresoInd(j).nRecupCtasxCob, "#0.00")
            .Fields("nPersFIEgresosOtros") = Format(FtesIngreso(i).FtesIngresoInd(j).nOtrosEgresos, "#0.00")
            .Fields("nPersIngFam") = Format(FtesIngreso(i).FtesIngresoInd(j).nPersIngFam, "#0.00")
            .Fields("nPersEgrFam") = Format(FtesIngreso(i).FtesIngresoInd(j).nPersEgrFam, "#0.00")
            .Fields("cUltimaActualizacion") = sActualiza
            .Fields("nTipoAct") = FtesIngreso(i).FtesIngresoInd(j).nTipoAct
            If FtesIngreso(i).nFtePaRscInd = 1 And nIntValor = 0 Then
                Set rsHojaEval = FtesIngreso(i).RFteIng
                Set rsEstFinan = FtesIngreso(i).rsEstFinan 'WIOR 20140321
                nIntValor = 2
            End If
        Next j
    Next i
    
    
    If Not .EOF Then .MoveFirst
End With

End Sub


Sub LlenaRecordset_FteIng()
Set rsFteIng = New ADODB.Recordset

With rsFteIng
    'Crear Recordset
    .Fields.Append "cNumFuente", adChar, 8
    .Fields.Append "cPersCod", adVarChar, 13
    .Fields.Append "cPersFIPersCod", adVarChar, 13
    .Fields.Append "cPersFICargo", adVarChar, 100
    .Fields.Append "dPersFIinicio", adDate
    .Fields.Append "cPersFIMoneda", adChar, 1
    .Fields.Append "cPersFIComentario", adVarChar, 300
    .Fields.Append "cRazSocUbiGeo", adVarChar, 12
    .Fields.Append "cRazSocDirecc", adVarChar, 250
    .Fields.Append "cRazSocTelef", adVarChar, 20
    .Fields.Append "cRazSocDescrip", adVarChar, 250
    .Fields.Append "nPersTipFte", adSmallInt
    .Fields.Append "cUltimaActualizacion", adVarChar, 25
    .Fields.Append "bCostoProd", adInteger
    .Fields.Append "nTipoAct", adInteger
    .Open
   
   'Llenar Recordset
    For i = 0 To nNumFtes - 1
        'La Clase va devolver el Numero de Fuente
        '.Fields("cNumFuente") = psNumFuente
        .AddNew
        .Fields("cNumFuente") = FtesIngreso(i).cNumFte
        .Fields("cPersFIPersCod") = FtesIngreso(i).sFteIngreso
        .Fields("cPersFICargo") = FtesIngreso(i).sCargo
        .Fields("dPersFIinicio") = FtesIngreso(i).dFteIngreso
        .Fields("cPersFIMoneda") = Trim(FtesIngreso(i).sMoneda)
        .Fields("cPersFIComentario") = Trim(FtesIngreso(i).sComentarios)
        .Fields("cRazSocUbiGeo") = Right(FtesIngreso(i).sRazSocUbiGeo, 12)
        .Fields("cRazSocDirecc") = FtesIngreso(i).sRazSocDirecc
        .Fields("cRazSocTelef") = FtesIngreso(i).sRazSocTelef
        .Fields("cRazSocDescrip") = Replace(FtesIngreso(i).sRazSocDescrip, "'", "''")
        .Fields("nPersTipFte") = FtesIngreso(i).nTipo
        .Fields("cUltimaActualizacion") = sActualiza
        .Fields("bCostoProd") = FtesIngreso(i).bCostoProd
        .Fields("nTipoAct") = FtesIngreso(i).nTipoAct
    Next i
    
    If Not .EOF Then .MoveFirst
End With

End Sub


'******************* Validaciones antes de Grabar ***********
Public Sub VerificandoDatosAGrabar(ByVal poPersona As COMDPersona.DCOMPersonas, _
                                    ByRef prsVerHomonimia As ADODB.Recordset, _
                                    ByRef pnVerDuplicaDoc As Integer, _
                                    ByRef pnVerTamanioDoc As Integer)

Dim R As ADODB.Recordset
Dim sNomTmp As String

On Error GoTo ErrorVerificandoDatosAGrabar

    If Personeria = 1 Then
        sNomTmp = ApellidoPaterno & "/" & ApellidoMaterno
    Else
        sNomTmp = NombreCompleto
    End If
    
    Set R = poPersona.BuscaCliente(sNomTmp, BusquedaNombre)
    Set prsVerHomonimia = R
    
'*******************************************
    pnVerDuplicaDoc = -1
    For i = 0 To NumeroDocumentos - 1
           If ObtenerDocTipoAct(i) <> PersFilaSinCambios And ObtenerDocTipoAct(i) <> PersFilaEliminda Then
             If poPersona.BuscaDocumento(CInt(Trim(Right(ObtenerTipoDoc(i), 15))), ObtenerNumeroDoc(i), PersCodigo) Then
                pnVerDuplicaDoc = i
                Exit For
             End If
         End If
    Next i

'*********************************************
    pnVerTamanioDoc = -1
    For i = 0 To NumeroDocumentos - 1
         If ObtenerDocTipoAct(i) <> PersFilaSinCambios And ObtenerDocTipoAct(i) <> PersFilaEliminda Then
            If (Len(ObtenerNumeroDoc(i)) <> 8 And CInt(Trim(Right(ObtenerTipoDoc(i), 15))) = gPersIdDNI) Or (Len(ObtenerNumeroDoc(i)) <> 11 And CInt(Trim(Right(ObtenerTipoDoc(i), 15))) = gPersIdRUC) Then
               pnVerTamanioDoc = i
               Exit For
            End If
        End If
    Next i

    Exit Sub
ErrorVerificandoDatosAGrabar:
    err.Raise err.Number, "Verificando Datos A Grabar", err.Description
End Sub

Public Sub ChequeoFuenteIngreso(ByVal pnIndice As Integer, ByVal pnTipoFte As PersFteIngresoTipo, ByVal pnPuntDet As Integer)
Dim dFteFecEval As Date
Dim DiasLimite As Integer
Dim dFteFecCaduca As Date
Dim poPersona As COMDPersona.DCOMPersona

    dFteFecEval = ObtenerFteIngFecEval(pnIndice, pnPuntDet, pnTipoFte)
    Set poPersona = New COMDPersona.DCOMPersona
    DiasLimite = poPersona.TiempoFteIngreso
    Set poPersona = Nothing
    dFteFecCaduca = CDate(Format(dFteFecEval, "dd/mm/yyyy")) + DiasLimite
    Call ActualizarFteIngFecCaduc(dFteFecCaduca, pnIndice, pnPuntDet, pnTipoFte)
End Sub

'Obtencion de Datos que no se consideraron previamente

Public Function ObtenerFteIngRuc(ByVal pnPunt As Integer) As String
    ObtenerFteIngRuc = FtesIngreso(pnPunt).cFtePersRuc
End Function
Public Function ObtenerFteIngFecInicioAct(ByVal pnPunt As Integer) As Date
    ObtenerFteIngFecInicioAct = FtesIngreso(pnPunt).dFtePersFecInicioAct
End Function
Public Function ObtenerFteIngTipoPersJur(ByVal pnPunt As Integer) As String
    ObtenerFteIngTipoPersJur = FtesIngreso(pnPunt).cFtePersTipoPersJur
End Function
Public Function ObtenerFteIngTelefono(ByVal pnPunt As Integer) As String
    ObtenerFteIngTelefono = FtesIngreso(pnPunt).cFtePersTelefono
End Function
Public Function ObtenerFteIngCondicionDomic(ByVal pnPunt As Integer) As String
    ObtenerFteIngCondicionDomic = FtesIngreso(pnPunt).cFtePersCondicionDomic
End Function
Public Function ObtenerFteIngCIUU(ByVal pnPunt As Integer) As String
    ObtenerFteIngCIUU = FtesIngreso(pnPunt).cFtePersCIUU
End Function
Public Function ObtenerFteIngMagnitudEmp(ByVal pnPunt As Integer) As String
    ObtenerFteIngMagnitudEmp = FtesIngreso(pnPunt).cFtePersMagnitudEmp
End Function
Public Function ObtenerFteIngNroEmpleados(ByVal pnPunt As Integer) As Double
    ObtenerFteIngNroEmpleados = FtesIngreso(pnPunt).nFtePersNroEmpleados
End Function
Public Function ObtenerFteIngDireccion(ByVal pnPunt As Integer) As String
    ObtenerFteIngDireccion = FtesIngreso(pnPunt).cFtePersDireccion
End Function
Public Function ObtenerFteIngDpto(ByVal pnPunt As Integer) As String
    ObtenerFteIngDpto = FtesIngreso(pnPunt).cFtePersDpto
End Function
Public Function ObtenerFteIngProv(ByVal pnPunt As Integer) As String
    ObtenerFteIngProv = FtesIngreso(pnPunt).cFtePersProv
End Function
Public Function ObtenerFteIngDist(ByVal pnPunt As Integer) As String
    ObtenerFteIngDist = FtesIngreso(pnPunt).cFtePersDist
End Function
Public Function ObtenerFteIngZona(ByVal pnPunt As Integer) As String
    ObtenerFteIngZona = FtesIngreso(pnPunt).cFtePersZona
End Function

'ARCV 12-08-2006

Public Function GenerarImpresionFteIngresoDependiente(ByVal pnIndex As Integer, _
    ByVal psNomAge As String, ByVal pdFecha As Date, ByVal pbCostoProd As Boolean, _
    Optional ByVal psNomCmac As String = "") As String

Dim sCadImp As String
Dim i As Integer
Dim j As Integer
Dim lnNegritaON As String
Dim lnNegritaOFF As String

Dim oDCred  As COMDCredito.DCOMCredito
Dim RsHis As ADODB.Recordset
Dim RsEvoEco As ADODB.Recordset
Dim rsRef As ADODB.Recordset
Dim nCalcula As Double
Dim MatDatosFte As Variant
Dim MatDatosComFte As Variant
Dim nMontoTemp As Double

Dim oPersona As COMDPersona.DCOMPersona
    
    
    '*** PEAC 20080515
    Set oDCred = New COMDCredito.DCOMCredito
    'Set RsHis = oDCred.RecuperaHisCred(sPersCod)
    Set RsEvoEco = oDCred.RecuperaEvolucionEconomica(sPersCod)
    Set rsRef = oDCred.RecuperaReferenciaPersComer(sPersCod)
    Set oDCred = Nothing
    
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
    
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & Space(20) & "EVOLUCION ECONOMICA FINANCIERA" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    
Do While Not RsEvoEco.EOF
    nCalcula = 0
    If RsEvoEco!cDescripcion = "FECHA" Then
        'sCadImp = sCadImp & ImpreFormat(RsEvoEco!cdescripcion, 31) & Space(2) & RsEvoEco!cFecha1 & Space(2) & RsEvoEco!cFecha2 & Space(2) & RsEvoEco!cFecha3 & Space(2) & RsEvoEco!cFecha4 & Space(2) & RsEvoEco!cFecha5
        sCadImp = sCadImp & ImpreFormat(RsEvoEco!cDescripcion, 31) & Space(2) & IIf(IsNull(RsEvoEco!cFecha1), "  /  /    ", RsEvoEco!cFecha1) & Space(2) & IIf(IsNull(RsEvoEco!cFecha2), "  /  /    ", RsEvoEco!cFecha2) & Space(2) & IIf(IsNull(RsEvoEco!cFecha3), "  /  /    ", RsEvoEco!cFecha3) & Space(2) & IIf(IsNull(RsEvoEco!cFecha4), "  /  /    ", RsEvoEco!cFecha4) & Space(2) & IIf(IsNull(RsEvoEco!cFecha5), "  /  /    ", RsEvoEco!cFecha5) & Space(2) & "VARIACION"
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & String(100, "-") & Chr$(10)
    Else
        
'        sCadImp = sCadImp & ImpreFormat(RsEvoEco!cDescripcion, 30)
'        If Not IsNull(RsEvoEco!cFecha1) Then
'            sCadImp = sCadImp & ImpreFormat(Val(RsEvoEco!cFecha1), 10, 2, True)
'        End If
'
'        If Not IsNull(RsEvoEco!cFecha2) Then
'            sCadImp = sCadImp & ImpreFormat(Val(RsEvoEco!cFecha2), 10, 2, True)
'        End If
'
'        If Not IsNull(RsEvoEco!cFecha3) Then
'            sCadImp = sCadImp & ImpreFormat(Val(RsEvoEco!cFecha3), 10, 2, True)
'        End If
'
'        If Not IsNull(RsEvoEco!cFecha4) Then
'            sCadImp = sCadImp & ImpreFormat(Val(RsEvoEco!cFecha4), 10, 2, True)
'        End If
'
'        If Not IsNull(RsEvoEco!cFecha5) Then
'            sCadImp = sCadImp & ImpreFormat(Val(RsEvoEco!cFecha5), 10, 2, True)
'        End If
    
'---------------------------------------------------------------------------------

        sCadImp = sCadImp & ImpreFormat(RsEvoEco!cDescripcion, 30)
        If Not IsNull(RsEvoEco!cFecha1) Then
            sCadImp = sCadImp & ImpreFormat(Val(RsEvoEco!cFecha1), 10, 2, True)
        Else
            sCadImp = sCadImp & Space(13)
        End If
        
        '*** PEAC 20080722 -- TODOS LOS COMENTADOS EN ESTE BLOQUE DEL "ELSE IF"
        'If Not IsNull(RsEvoEco!cFecha2) Then
        If RsEvoEco!cFecha2 > 0 Then
            sCadImp = sCadImp & ImpreFormat(Val(RsEvoEco!cFecha2), 10, 2, True)
            'If Not IsNull(RsEvoEco!cFecha1) Then
            If RsEvoEco!cFecha1 > 0 Then
                nCalcula = Val(RsEvoEco!cFecha2) / Val(RsEvoEco!cFecha1)
            Else
                nCalcula = 0
            End If
        Else
            sCadImp = sCadImp & ImpreFormat(0, 10, 2, True) 'Space(13)
        End If
        
        'If Not IsNull(RsEvoEco!cFecha3) Then
        If RsEvoEco!cFecha3 > 0 Then
            sCadImp = sCadImp & ImpreFormat(Val(RsEvoEco!cFecha3), 10, 2, True)
            'If Not IsNull(RsEvoEco!cFecha2) Then
            If RsEvoEco!cFecha2 > 0 Then
                nCalcula = Val(RsEvoEco!cFecha3) / Val(RsEvoEco!cFecha2)
            Else
                nCalcula = 0
            End If
        Else
            sCadImp = sCadImp & ImpreFormat(0, 10, 2, True) 'Space(13)
        End If
        
        'If Not IsNull(RsEvoEco!cFecha4) Then
        If RsEvoEco!cFecha4 > 0 Then
            sCadImp = sCadImp & ImpreFormat(Val(RsEvoEco!cFecha4), 10, 2, True)
            'If Not IsNull(RsEvoEco!cFecha3) Then
            If RsEvoEco!cFecha3 > 0 Then
                nCalcula = Val(RsEvoEco!cFecha4) / Val(RsEvoEco!cFecha3)
            Else
                nCalcula = 0
            End If
        Else
            sCadImp = sCadImp & ImpreFormat(0, 10, 2, True) 'Space(13)
        End If
        
        'If Not IsNull(RsEvoEco!cFecha5) Then
        If RsEvoEco!cFecha5 > 0 Then
            sCadImp = sCadImp & ImpreFormat(Val(RsEvoEco!cFecha5), 10, 2, True)
            'If Not IsNull(RsEvoEco!cFecha4) Then
            If RsEvoEco!cFecha4 > 0 Then
                nCalcula = Val(RsEvoEco!cFecha5) / Val(RsEvoEco!cFecha4)
            Else
                nCalcula = 0
            End If
        Else
            sCadImp = sCadImp & ImpreFormat(0, 10, 2, True) 'Space(13)
        End If
        
        sCadImp = sCadImp & ImpreFormat(nCalcula, 5, 2)


'------------------------------------------------------------------

'        sCadImp = sCadImp & ImpreFormat(RsEvoEco!cdescripcion, 30)
'        If Not IsNull(RsEvoEco!cFecha1) Then
'            sCadImp = sCadImp & ImpreFormat(Val(RsEvoEco!cFecha1), 10, 2, True)
'        Else
'            sCadImp = sCadImp & Space(13)
'        End If
'
'        If Not IsNull(RsEvoEco!cFecha2) Then
'            sCadImp = sCadImp & ImpreFormat(Val(RsEvoEco!cFecha2), 10, 2, True)
'            If Not IsNull(RsEvoEco!cFecha1) Then
'                nCalcula = Val(RsEvoEco!cFecha2) / Val(RsEvoEco!cFecha1)
'            Else
'                nCalcula = 0
'            End If
'        Else
'            sCadImp = sCadImp & Space(13)
'        End If
'
'        If Not IsNull(RsEvoEco!cFecha3) Then
'            sCadImp = sCadImp & ImpreFormat(Val(RsEvoEco!cFecha3), 10, 2, True)
'            If Not IsNull(RsEvoEco!cFecha2) Then
'                nCalcula = Val(RsEvoEco!cFecha3) / Val(RsEvoEco!cFecha2)
'            Else
'                nCalcula = 0
'            End If
'        Else
'            sCadImp = sCadImp & Space(13)
'        End If
'
'        If Not IsNull(RsEvoEco!cFecha4) Then
'            sCadImp = sCadImp & ImpreFormat(Val(RsEvoEco!cFecha4), 10, 2, True)
'            If Not IsNull(RsEvoEco!cFecha3) Then
'                nCalcula = Val(RsEvoEco!cFecha4) / Val(RsEvoEco!cFecha3)
'            Else
'                nCalcula = 0
'            End If
'        Else
'            sCadImp = sCadImp & Space(13)
'        End If
'
'        If Not IsNull(RsEvoEco!cFecha5) Then
'            sCadImp = sCadImp & ImpreFormat(Val(RsEvoEco!cFecha5), 10, 2, True)
'            If Not IsNull(RsEvoEco!cFecha4) Then
'                nCalcula = Val(RsEvoEco!cFecha5) / Val(RsEvoEco!cFecha4)
'            Else
'                nCalcula = 0
'            End If
'        Else
'            sCadImp = sCadImp & Space(13)
'        End If
'
'        sCadImp = sCadImp & ImpreFormat(nCalcula, 5, 2)
    
    End If
    
    sCadImp = sCadImp & Chr$(10)
    
    
    RsEvoEco.MoveNext
Loop
sCadImp = sCadImp & Chr$(10)
sCadImp = sCadImp & String(100, "-") & Chr$(10)
    
    
'*** PEAC 20080527
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & Space(20) & "EVOLUCION ECONOMICA FINANCIERA DEPENDIENTE" & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'    sCadImp = sCadImp & Space(23)
'
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
'        nMontoTemp = FtesIngreso(pnIndex).FtesIngresoDep(i).nIngresos
'        If FtesIngreso(pnIndex).FtesIngresoDep(i).dFecEval <> 0 Then
'            sCadImp = sCadImp & Format(FtesIngreso(pnIndex).FtesIngresoDep(i).dFecEval, "dd/mm/yyyy") & Space(3)
'        Else
'            sCadImp = sCadImp & Space(5) & "(-)" & Space(3)
'        End If
'        If i = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
'            sCadImp = sCadImp & "Variacion"
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Ingreso Familiar", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngFam, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesDep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngFam = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 5, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(CDbl(Format(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngFam / FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngFam, "#0.00")), 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'
'    sCadImp = sCadImp & ImpreFormat("Ingreso Conyugue", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngCony, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesDep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngCony = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 5, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(CDbl(Format(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngCony / FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngCony, "#0.00")), 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'
'    sCadImp = sCadImp & ImpreFormat("Otros Ingresos", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngOtros, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesDep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngOtros = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 5, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(CDbl(Format(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngOtros / FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngOtros, "#0.00")), 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'
'    sCadImp = sCadImp & ImpreFormat("Gastos Familiares", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(i).nGastosFam, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesDep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nGastosFam = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 5, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(CDbl(Format(FtesIngreso(pnIndex).FtesIngresoDep(i).nGastosFam / FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nGastosFam, "#0.00")), 8, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ImpreFormat("Saldo", 20)
'    For i = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(i).nIngCony + _
'                                FtesIngreso(pnIndex).FtesIngresoDep(i).nIngFam + _
'                                FtesIngreso(pnIndex).FtesIngresoDep(i).nIngOtros - FtesIngreso(pnIndex).FtesIngresoDep(i).nGastosFam, 8, 2)
'        If i = FtesIngreso(pnIndex).nNumFtesDep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngCony + _
'                                FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngFam + _
'                                FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngOtros - FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nGastosFam = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 5, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(CDbl(Format((FtesIngreso(pnIndex).FtesIngresoDep(i).nIngCony + _
'                                FtesIngreso(pnIndex).FtesIngresoDep(i).nIngFam + _
'                                FtesIngreso(pnIndex).FtesIngresoDep(i).nIngOtros - FtesIngreso(pnIndex).FtesIngresoDep(i).nGastosFam) / _
'                                (FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngCony + _
'                                FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngFam + _
'                                FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nIngOtros - FtesIngreso(pnIndex).FtesIngresoDep(i - 1).nGastosFam), "#0.00")), 8, 2)
'
'            End If
'        End If
'    Next i
'*** PEAC 20080527
    
    
'*** PEAC 20080527
'    'IMPRIME COSTO DE PRODUCCION
'    '-------------------------------------
'    If pbCostoProd = True Then
'            sCadImp = sCadImp & Chr$(12)
'            sCadImp = sCadImp & Chr$(10) & Chr$(10)
'            sCadImp = sCadImp & String(100, "-") & Chr$(10)
'            sCadImp = sCadImp & Space(30) & "COSTO DE PRODUCCION" & Chr$(10)
'            sCadImp = sCadImp & String(100, "-") & Chr$(10)
'            sCadImp = sCadImp & ImpreFormat("Evaluacion", 16)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).dFecEval, 14, 0)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ImpreFormat("Tipo Cultivo", 20)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).sTpoCultivo, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ImpreFormat("DESTINO", 20)
'            Dim sDestino As String
'
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                If FtesIngreso(pnIndex).CostoProd(i).nSiembra = 1 Then
'        '            If Len(sDestino) = 0 Then
'        '                sDestino = "SIEMBRA"
'        '            Else
'        '                sDestino = sDestino & ",SIEMBRA"
'                        sCadImp = sCadImp & ImpreFormat("SIEMBRA", 13)
'                        sCadImp = sCadImp & Chr$(10)
'        '            End If
'                 End If
'
'                If FtesIngreso(pnIndex).CostoProd(i).nMantenimiento = 1 Then
'        '            If Len(sDestino) = 0 Then
'        '                sDestino = "MANTENIMIENTO"
'        '            Else
'        '                sDestino = sDestino & ",MANTENIMIENTO"
'        '                sCadImp = sCadImp & Chr$(10)
'        '            End If
'                        sCadImp = sCadImp & Space(22) & ImpreFormat("MANTENIMIENTO", 13)
'                        sCadImp = sCadImp & Chr$(10)
'                End If
'                If FtesIngreso(pnIndex).CostoProd(i).nDesAgricola = 1 Then
'        '             If Len(sDestino) = 0 Then
'        '                sDestino = "DESAGRICOLA"
'        '             Else
'        '                sDestino = sDestino & ",DESAGRICOLA"
'        '                sCadImp = sCadImp & Chr$(10)
'        '             End If
'                        sCadImp = sCadImp & Space(22) & ImpreFormat("DESAGRICOLA", 13)
'                        sCadImp = sCadImp & Chr$(10)
'
'
'                End If
'                If FtesIngreso(pnIndex).CostoProd(i).nCosecha = 1 Then
'        '            If Len(sDestino) = 0 Then
'        '                sDestino = "COSECHA"
'        '            Else
'        '                sDestino = sDestino & ",COSECHA"
'        '                sCadImp = sCadImp & Chr$(10)
'        '            End If
'                        sCadImp = sCadImp & Space(22) & ImpreFormat("COSECHA", 13)
'                        sCadImp = sCadImp & Chr$(10)
'
'                 End If
'                If FtesIngreso(pnIndex).CostoProd(i).nPlanOtros = 1 Then
'        '            If Len(sDestino) = 0 Then
'        '                sDestino = "OTROS"
'        '            Else
'        '                sDestino = sDestino & ",OTROS"
'        '            End If
'                        sCadImp = sCadImp & Space(22) & ImpreFormat("OTROS", 13)
'                        'sCadImp = sCadImp & Chr$(10)
'
'                End If
'                'sCadImp = sCadImp & ImpreFormat(sDestino, 60)
'            Next i
'
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ImpreFormat("Maquinaria", 12)
'
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nMaquinaria, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ImpreFormat("Jornales", 12)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nJornales, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ImpreFormat("Insumos", 12)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nInsumos, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ImpreFormat("Pesticidas", 12)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPesticidas, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ImpreFormat("Otros", 12)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nOtros, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ImpreFormat("Total", 12)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10) & Chr$(10)
'
'            sCadImp = sCadImp & ImpreFormat("Hectareas", 12)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ImpreFormat("Produccion", 18)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ImpreFormat(Format(FtesIngreso(pnIndex).CostoProd(i).nProduccion, "#0.00") & " " & FtesIngreso(pnIndex).CostoProd(i).sUniProd, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ImpreFormat("Precio Uni.", 12)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPreUni, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ImpreFormat("Ingresos", 12)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas * FtesIngreso(pnIndex).CostoProd(i).nPreUni * FtesIngreso(pnIndex).CostoProd(i).nProduccion, 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ImpreFormat("Egresos", 12)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas * (FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros), 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'
'            sCadImp = sCadImp & ImpreFormat("Utilidad", 12)
'            For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'                sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).CostoProd(i).nHectareas * FtesIngreso(pnIndex).CostoProd(i).nPreUni * FtesIngreso(pnIndex).CostoProd(i).nProduccion) - (FtesIngreso(pnIndex).CostoProd(i).nHectareas * (FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros)), 13)
'            Next i
'            sCadImp = sCadImp & Chr$(10)
'            sCadImp = sCadImp & String(100, "-") & Chr$(10)
'    End If
'*** FIN PEAC 20080527
        
    
'*** PEAC 20080527
    ' ---------------------------------------------------------------------------------------------------------------
    
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & Space(20) & "REFERENCIAS PERSONALES / COMERCIALES" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
     
    sCadImp = sCadImp & "   NOMBRE/REF.COMERCIAL            TIPO REF. TELEFONO     COMENTARIO              DIRECCION" & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    
    Do While Not rsRef.EOF
        sCadImp = sCadImp & ImpreFormat(UCase(rsRef!cNomRefCom), 30) & ImpreFormat(rsRef!cConsDescripcion, 10) & ImpreFormat(rsRef!cFonoRefCom, 10) & ImpreFormat(UCase(rsRef!cComentarioRefCom), 20) & ImpreFormat(UCase(rsRef!cDireccionRefCom), 20)
        sCadImp = sCadImp & Chr$(10)
        rsRef.MoveNext
    Loop
    
    sCadImp = sCadImp & Chr$(10)
        
    ' Muestra los datos de las referencias personales de la fte como persona ----------------------------------------
'    sCadImp = sCadImp & Chr$(10) & Chr$(10)
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & "REFERENCIAS PERSONALES" & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'
'    sCadImp = sCadImp & FtesIngreso(pnIndex).sComentarios & Chr$(10)
'
'    sCadImp = sCadImp & Chr$(10)
'
'    ' Muestra los datos de las referencias bancarias de la fte como persona ----------------------------------------
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ImpreFormat("REFERENCIAS BANCARIAS", 42, 0) & ImpreFormat("NRO.TARJETA", 22, 2) & ImpreFormat("NRO.CUENTA", 22, 2) & ImpreFormat("MONTO", 10, 2) & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'
'    'MatDatosFte = RecuperaRefBancariaPersFte(FtesIngreso(pnIndex).sFteIngreso)
'    Set oPersona = New COMDPersona.DCOMPersona
'    MatDatosFte = oPersona.RecuperaRefBancariaPersFte(sPersCod)
'
'    For i = 0 To UBound(MatDatosFte) - 1
'        sCadImp = sCadImp & ImpreFormat(MatDatosFte(i, 0) & "-" & MatDatosFte(i, 1), 40) & ImpreFormat(MatDatosFte(i, 3), 22, 2) & ImpreFormat(MatDatosFte(i, 2), 22, 2) & ImpreFormat(MatDatosFte(i, 4), 10, 2) & Chr$(10)
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'
'   ' Muestra los datos de las referencias comerciales de la fte como persona ----------------------------------------
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ImpreFormat("REFERENCIAS COMERCIALES", 66, 0) & ImpreFormat("TIPO RELAC. COMERCIAL", 22, 2) & ImpreFormat("TELEFONO", 10, 2) & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'
'    'MatDatosComFte = opersona.RecuperaRefComercialPersFte(FtesIngreso(pnIndex).sFteIngreso)
'    MatDatosComFte = oPersona.RecuperaRefComercialPersFte(sPersCod)
'    For i = 0 To UBound(MatDatosComFte) - 1
'        sCadImp = sCadImp & ImpreFormat(MatDatosComFte(i, 3) & "-" & MatDatosComFte(i, 0), 64) & ImpreFormat(MatDatosComFte(i, 2), 22, 2) & ImpreFormat(MatDatosComFte(i, 1), 10, 2) & Chr$(10)
'    Next i
'
    ' ---------------------------------------------------------------------------------------------------------------
'*** FIN PEAC 20080527
    
    Set oPersona = Nothing
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    GenerarImpresionFteIngresoDependiente = sCadImp
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
Dim oPersona As COMDPersona.DCOMPersona

'***PEAC 20080515 ------------------------
Dim oDCred As COMDCredito.DCOMCredito
Dim RsHis As ADODB.Recordset
Dim RsEvoEco As ADODB.Recordset
Dim rsRef As ADODB.Recordset
Dim nCalcula As Double

Dim lnActivoTotal As Double
Dim lnPasivoTotal As Double
Dim lnPatrimonio As Double
Dim lnTotPasPatri As Double

'*** -------------------------------------

'ARCV 01-02-2007
Dim nNumFtesRep As Integer
Dim nNumEvalInicio As Integer
Dim oPar As COMDCredito.DCOMParametro
Set oPar = New COMDCredito.DCOMParametro
nNumFtesRep = oPar.RecuperaValorParametro(3075)
Set oPar = Nothing

If FtesIngreso(pnIndex).nNumFtesIndep < nNumFtesRep Then
    nNumEvalInicio = 0
Else
    nNumEvalInicio = FtesIngreso(pnIndex).nNumFtesIndep - nNumFtesRep
End If
nNumFtesRep = FtesIngreso(pnIndex).nNumFtesIndep
'------------------------------

    lnNegritaON = Chr$(27) & Chr$(69)
    lnNegritaOFF = Chr$(27) + Chr$(70)

    Set oPersona = New COMDPersona.DCOMPersona
    MatDatos = oPersona.RecuperaHistoriaCredPersona(sPersCod)

    '*** PEAC 20080515
    Set oDCred = New COMDCredito.DCOMCredito
    Set RsHis = oDCred.RecuperaHisCred(sPersCod)
    Set RsEvoEco = oDCred.RecuperaEvolucionEconomica(sPersCod)
    Set rsRef = oDCred.RecuperaReferenciaPersComer(sPersCod)
    Set oDCred = Nothing

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
    sCadImp = sCadImp & Space(25) & Format(pdFecha, "dd/mm/yyyy hh:mm:ss") & Chr$(10)
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
     
    sCadImp = sCadImp & "   CREDITO               ESTADO   FEC.VIGENCIA FEC.EVALUA. MONTO SOL.     CUOTAS    ANALISTA" & Chr$(10)
     sCadImp = sCadImp & String(100, "-") & Chr$(10)
    
Do While Not RsHis.EOF
    sCadImp = sCadImp & ImpreFormat(RsHis!cCtaCod, 18) & ImpreFormat(RsHis!Estado, 10) & ImpreFormat(Format(RsHis!dVigencia, "dd/mm/yyyy"), 10) & ImpreFormat(Format(RsHis!dPersEval, "dd/mm/yyyy"), 10) & ImpreFormat(RsHis!montosoli, 10, 2, True) & ImpreFormat(RsHis!nCuotas, 10, 0) & Space(5) & ImpreFormat(RsHis!cUser, 5)
    sCadImp = sCadImp & Chr$(10)
    RsHis.MoveNext
Loop
'RsHis.MoveFirst
'sCadImp = sCadImp & Chr$(10)

    'sCadImp = sCadImp & "MONEDA             : " & ImpreFormat(IIf(FtesIngreso(pnIndex).sMoneda = "1", "SOLES", "DOLARES"), 20)

'*** PEAC 20080515

'    If Not IsArray(MatDatos) Then
'        sCadImp = sCadImp & "NO POSEE CREDITOS" & Chr$(10)
'    Else
'        'sCadImp = sCadImp & "CREDITO No         : "
'        'For I = 0 To UBound(MatDatos) - 1
'        '    sCadImp = sCadImp & MatDatos(I, 0) & Space(2)
'        'Next I
'
'        sCadImp = sCadImp & Chr$(10)
'        sCadImp = sCadImp & "CREDITO No         : "
'        For i = 0 To UBound(MatDatos) - 1
'            sCadImp = sCadImp & ImpreFormat(MatDatos(i, 0), 20) & Space(2)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'        sCadImp = sCadImp & "Estado Actual      : "
'        For i = 0 To UBound(MatDatos) - 1
'            sCadImp = sCadImp & ImpreFormat(MatDatos(i, 1), 20) & Space(2)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'        sCadImp = sCadImp & "Fecha Vigencia     : "
'        For i = 0 To UBound(MatDatos) - 1
'            sCadImp = sCadImp & ImpreFormat(MatDatos(i, 2), 20) & Space(2)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'        sCadImp = sCadImp & "Fecha Cancelacion  : "
'        For i = 0 To UBound(MatDatos) - 1
'            sCadImp = sCadImp & ImpreFormat(MatDatos(i, 3), 20) & Space(2)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'        sCadImp = sCadImp & "Monto(Sol/Sug/Apr) : "
'        For i = 0 To UBound(MatDatos) - 1
'            sCadImp = sCadImp & ImpreFormat(MatDatos(i, 4), 20) & Space(2)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'        sCadImp = sCadImp & "Cuotas/Plazo(Dias) : "
'        For i = 0 To UBound(MatDatos) - 1
'            sCadImp = sCadImp & ImpreFormat(MatDatos(i, 5), 20) & Space(2)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'        sCadImp = sCadImp & "Cuota (Sol/Sug/Apr): "
'        For i = 0 To UBound(MatDatos) - 1
'            sCadImp = sCadImp & ImpreFormat(MatDatos(i, 6), 20) & Space(2)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'        sCadImp = sCadImp & "Nota               : "
'        For i = 0 To UBound(MatDatos) - 1
'            sCadImp = sCadImp & ImpreFormat(MatDatos(i, 7), 20) & Space(2)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'        sCadImp = sCadImp & "Analista           : "
'        For i = 0 To UBound(MatDatos) - 1
'            sCadImp = sCadImp & ImpreFormat(MatDatos(i, 8), 20) & Space(2)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'    End If

'*** FIN PEAC 20080515
    

'*** COMENTADO PEAC 20080402
'****************************--------------------------------------***************************************

    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & Space(20) & "EVOLUCION ECONOMICA FINANCIERA" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'    sCadImp = sCadImp & "Activo               Activo   Activo   Pasivo   Pasivo   Pasivo               Total Pasivo" & Chr$(10)
'    sCadImp = sCadImp & "Cte     Inventario   Fijo     Total    Cte.     no Cte.  Total   Patrimonio   y Patrimonio" & Chr$(10)
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'    sCadImp = sCadImp & Space(25)
    
'sCadImp = sCadImp & "                   "
'*** PEAC 20080515

Do While Not RsEvoEco.EOF
nCalcula = 0
'    lnActivoTotal = RsHis!Act_cte + RsHis!act_fijo
'    lnPasivoTotal = RsHis!pas_no_cte + RsHis!pas_cte
'    lnPatrimonio = lnActivoTotal - lnPasivoTotal
'    lnTotPasPatri = lnPasivoTotal + lnPatrimonio
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat(RsHis!cCtaCod, 20) & ImpreFormat(RsHis!dPersEval, 10)
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat(RsHis!Act_cte, 6, 2) & ImpreFormat(RsHis!INVENTARIO, 6, 2) & ImpreFormat(RsHis!act_fijo, 6, 2) & ImpreFormat(lnActivoTotal, 6, 2) & ImpreFormat(RsHis!pas_cte, 6, 2) & ImpreFormat(RsHis!pas_no_cte, 6, 2) & ImpreFormat(lnPasivoTotal, 6, 2)
'    sCadImp = sCadImp & ImpreFormat(lnPatrimonio, 6, 2) & ImpreFormat(lnTotPasPatri, 6, 2) & ImpreFormat(RsHis!ingresos, 6, 2) & ImpreFormat(RsHis!cos_vtas, 6, 2) & ImpreFormat(RsHis!otr_egre, 6, 2)
    
'    sCadImp = sCadImp & Chr$(10)
    
    'sCadImp = sCadImp & ImpreFormat(RsHis!Nombre, 50) ''& ImpreFormat(RsHis!dPersEval, 10)
    If RsEvoEco!cDescripcion = "FECHA" Then
        sCadImp = sCadImp & ImpreFormat(RsEvoEco!cDescripcion, 31) & Space(2) & IIf(IsNull(RsEvoEco!cFecha1), "  /  /    ", RsEvoEco!cFecha1) & Space(2) & IIf(IsNull(RsEvoEco!cFecha2), "  /  /    ", RsEvoEco!cFecha2) & Space(2) & IIf(IsNull(RsEvoEco!cFecha3), "  /  /    ", RsEvoEco!cFecha3) & Space(2) & IIf(IsNull(RsEvoEco!cFecha4), "  /  /    ", RsEvoEco!cFecha4) & Space(2) & IIf(IsNull(RsEvoEco!cFecha5), "  /  /    ", RsEvoEco!cFecha5) & Space(2) & "VARIACION"
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & String(100, "-") & Chr$(10)
    Else
        
        sCadImp = sCadImp & ImpreFormat(RsEvoEco!cDescripcion, 30)
        If Not IsNull(RsEvoEco!cFecha1) Then
            sCadImp = sCadImp & ImpreFormat(Val(RsEvoEco!cFecha1), 10, 2, True)
        Else
            sCadImp = sCadImp & Space(13)
        End If
        
        '*** PEAC 20080722 -- TODOS LOS COMENTADOS EN ESTE BLOQUE DEL "ELSE IF"
        'If Not IsNull(RsEvoEco!cFecha2) Then
        If RsEvoEco!cFecha2 > 0 Then
            sCadImp = sCadImp & ImpreFormat(Val(RsEvoEco!cFecha2), 10, 2, True)
            'If Not IsNull(RsEvoEco!cFecha1) Then
            If RsEvoEco!cFecha1 > 0 Then
                nCalcula = Val(RsEvoEco!cFecha2) / Val(RsEvoEco!cFecha1)
            Else
                nCalcula = 0
            End If
        Else
            sCadImp = sCadImp & ImpreFormat(0, 10, 2, True) 'Space(13)
        End If
        
        'If Not IsNull(RsEvoEco!cFecha3) Then
        If RsEvoEco!cFecha3 > 0 Then
            sCadImp = sCadImp & ImpreFormat(Val(RsEvoEco!cFecha3), 10, 2, True)
            'If Not IsNull(RsEvoEco!cFecha2) Then
            If RsEvoEco!cFecha2 > 0 Then
                nCalcula = Val(RsEvoEco!cFecha3) / Val(RsEvoEco!cFecha2)
            Else
                nCalcula = 0
            End If
        Else
            sCadImp = sCadImp & ImpreFormat(0, 10, 2, True) 'Space(13)
        End If
        
        'If Not IsNull(RsEvoEco!cFecha4) Then
        If RsEvoEco!cFecha4 > 0 Then
            sCadImp = sCadImp & ImpreFormat(Val(RsEvoEco!cFecha4), 10, 2, True)
            'If Not IsNull(RsEvoEco!cFecha3) Then
            If RsEvoEco!cFecha3 > 0 Then
                nCalcula = Val(RsEvoEco!cFecha4) / Val(RsEvoEco!cFecha3)
            Else
                nCalcula = 0
            End If
        Else
            sCadImp = sCadImp & ImpreFormat(0, 10, 2, True) 'Space(13)
        End If
        
        'If Not IsNull(RsEvoEco!cFecha5) Then
        If RsEvoEco!cFecha5 > 0 Then
            sCadImp = sCadImp & ImpreFormat(Val(RsEvoEco!cFecha5), 10, 2, True)
            'If Not IsNull(RsEvoEco!cFecha4) Then
            If RsEvoEco!cFecha4 > 0 Then
                nCalcula = Val(RsEvoEco!cFecha5) / Val(RsEvoEco!cFecha4)
            Else
                nCalcula = 0
            End If
        Else
            sCadImp = sCadImp & ImpreFormat(0, 10, 2, True) 'Space(13)
        End If
        
        sCadImp = sCadImp & ImpreFormat(nCalcula, 5, 2)
                
    End If
    
    sCadImp = sCadImp & Chr$(10)
    
    'RsHis.Fields.Item(i).Name
    
'    For i = 2 To RsHis.Fields.Count - 1
'        sCadImp = sCadImp & Space(2) & Str("x" & RsHis.Fields.Item(i).Name)
'    Next
    
    RsEvoEco.MoveNext
Loop
sCadImp = sCadImp & Chr$(10)
sCadImp = sCadImp & String(100, "-") & Chr$(10)
    
    
'
'    Dim dFecTemEval As Date
'
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        nMontoT = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).dFecEval <> 0 Then
'            sCadImp = sCadImp & Format(FtesIngreso(pnIndex).FtesIngresoInd(i).dFecEval, "dd/mm/yyyy") & Space(3)
'        Else
'            sCadImp = sCadImp & Space(5) & "(-)" & Space(3)
'        End If
'    Next i
'    sCadImp = sCadImp & "Variacion" & Chr$(10)
'    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)

'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ImpreFormat("ACTIVO CORRIENTE", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios, 10, 2)
'        If i = nNumFtesRep - 1 And i > 0 Then
'            If (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios) = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios), 10, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    sCadImp = sCadImp & ImpreFormat("Disponible", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp, 10, 2)
'        If i = nNumFtesRep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp, 10, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Cuentas x Cobrar", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob, 10, 2)
'        If i = nNumFtesRep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob, 10, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Inventario", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios, 10, 2)
'        If i = nNumFtesRep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios, 10, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ImpreFormat("ACTIVO NO CORRIENTE", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo, 10, 2)
'        If i = nNumFtesRep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo, 10, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("TOTAL ACTIVO", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo, 10, 2)
'        If i = nNumFtesRep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo), 10, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("PASIVO", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact, 10, 2)
'        If i = nNumFtesRep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact = 0 Then
'
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact), 10, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'    sCadImp = sCadImp & ImpreFormat("Proveedores", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores, 10, 2)
'        If i = nNumFtesRep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores, 10, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Otros Prestamos", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros, 10, 2)
'        If i = nNumFtesRep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros, 10, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Prestamos CMAC-C", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact, 10, 2)
'        If i = nNumFtesRep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact, 10, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ImpreFormat("PATRIMONIO", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio, 10, 2)
'        If i = nNumFtesRep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact, 10, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("INGRESOS", 20)
'    Dim nIng As Currency
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + _
'                FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
'                FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob, 10, 2)
'
''        nIng = FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob + _
'                FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam
'
'        If i = nNumFtesRep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam = 0 Then
'
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'               sCadImp = sCadImp & ImpreFormat(nIng / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob), 10, 2)
'
'            End If
'        End If
'    Next i
'
'    'CMACICA_CSTS - 10/11/2003 - SE CAMBIO EL ORDEN (VENTAS, RECUP. X COB., INGRESOS FAMILIARES) -----------------
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'
'    sCadImp = sCadImp & ImpreFormat("Ventas", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas, 10, 2)
'        If i > 0 And nNumFtesRep - 1 = i Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas, 10, 2)
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Recup. Ctas x Cob", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob, 10, 2)
'        If i > 0 And i = nNumFtesRep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob, 10, 2)
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Otros Ingresos", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam, 10, 2)
'        If i = nNumFtesRep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam, 10, 2)
'            End If
'        End If
'    Next i
'
''-----------------------------------------------------------------------------------------------------------------------
'    Dim nEg As Currency
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ImpreFormat("EGRESOS", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 10, 2)
'        'sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 10, 2)
'         'nEg = FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam
'        If i > 0 And i = nNumFtesRep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos = 0 Then
''            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'
'                sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos), 10, 2)
''                sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos) / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos), 10, 2)
'            End If
'        End If
'    Next i
'
'    'CMACICA_CSTS - 21/11/2003 - SE CAMBIO EL ORDEN (COMPRA DE MERCADERIAS, OTROS EGRESOS, GASTOS FAMILIARES) ---------
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'
'    'sCadImp = sCadImp & ImpreFormat("Compra Mercaderia", 20)
'    sCadImp = sCadImp & ImpreFormat("Gasto de Ventas", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad, 10, 2)
'        If i > 0 And i = nNumFtesRep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad, 10, 2)
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Gastos Administrativos", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam, 10, 2)
'        If i > 0 And i = nNumFtesRep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam, 10, 2)
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Otros Egresos", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 10, 2)
'        If i > 0 And i = nNumFtesRep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos, 10, 2)
'            End If
'        End If
'    Next i
'FIN COMENTADO POR PEAC
    
    
' ---------------------------------------------------------------------------------------------------------------------

    'sCadImp = sCadImp & Chr$(10)
    'sCadImp = sCadImp & lnNegritaON
    'sCadImp = sCadImp & ImpreFormat("UTILIDAD NETA", 20)
    'For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob - _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam - _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad - _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 10, 2)
     '   sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob - _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad - _
                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos, 10, 2)

     '   If i > 0 And i = nNumFtesRep - 1 Then

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

      '          sCadImp = sCadImp & ImpreFormat(0, 10, 2)

      '     Else
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
'                             FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos)), 10, 2)
         '       sCadImp = sCadImp & ImpreFormat(((FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob) - _
                            (FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
                             FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos)) _
                            / ((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + _
                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - _
                            (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
                             FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos)), 10, 2)
          '  End If
        'End If
    'Next i
  '  sCadImp = sCadImp & Chr$(10)
    'sCadImp = sCadImp & Chr$(10)
    'sCadImp = sCadImp & String(80, "-") & Chr$(10)
  '  sCadImp = sCadImp & String(100, "-") & Chr$(10)
  '  sCadImp = sCadImp & Chr$(10)
  '  sCadImp = sCadImp & Space(20) & "INGRESOS-EGRESOS FAMILIARES" & Chr$(10)
  '  sCadImp = sCadImp & ImpreFormat("Ingresos familiares", 20)
'    Dim nIngFam As Currency

'    For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam, 10, 2)
'        nIngFam = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam
'        If i = nNumFtesRep - 1 And i > 0 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam, 10, 2)
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Gastos Familiares", 20)
'    Dim nEgrFam As Currency
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam, 10, 2)
'        nEgrFam = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam
'        If i > 0 And i = nNumFtesRep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam, 10, 2)
'            End If
'        End If
'    Next i
    
    
'*** COMENTADO POR PEAC 20080402

'    sCadImp = sCadImp & Chr$(10) & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("DISPONIBLE", 20)
'    'sCadImp = sCadImp & ImpreFormat(nIngFam + nIng - (nEgrFam + nEg), 10, 2)
'
'    Dim nTemp As Double
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        nIng = FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob + _
'                FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam
'        nEg = FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam
'
'        sCadImp = sCadImp & ImpreFormat(nIng - nEg, 10, 2)
'
'        If i > 0 And i = nNumFtesRep - 1 Then
'
'            nTemp = ((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob + _
'                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam) - (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam))
'            If nTemp = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                sCadImp = sCadImp & _
'                         ImpreFormat((nIng - nEg) / ((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam) - (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos + _
'                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam)), 10, 2)
'           End If
'        End If
'    Next i

'FIN COMENTADO POR PEAC
    
    
'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************
'Dim nLiquidezGeneral As Currency
'Dim nPruebaAcida As Currency
'Dim nCapitalTrabajo As Currency
'Dim nRotacionCtasCobrar As Currency
'Dim nPeriodoMedioC As Currency
'Dim nRotacionCtasPagar As Currency
'Dim nPeriodoMedioP As Currency
'Dim nRotacInvV As Currency
'Dim nRotacInvD As Currency
'Dim nRotacActivoFijo As Currency
'Dim nMargenBruto As Currency
'Dim nGradoEndeudam As Currency
'Dim nGradoPropiedad As Currency
'Dim nEndeudaPatrim As Currency
'Dim nRentabPatrim As Currency
'Dim nRentabVentas As Currency
'Dim nActivoCorriente As Currency 'Disponible + Ctas por Cobrar + Inventario
'Dim nPasivo As Currency 'Proveedores + Otros Prestamos + Prestamos Cmac
'Dim nTotalIngresos As Currency  'Ventas + Recup. Ctas x Cobrar + Otros ingresos
'Dim nTotalEgresos As Currency     'Gastos de Ventas + Gastos Administrativos + Otros Egresos
'Dim nTotalActivo As Currency    'Disponible + Ctas por Cobrar + Inventario+ Activo No Corriente
'
'    sCadImp = sCadImp & Chr$(10) & Chr$(10)
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'    sCadImp = sCadImp & Space(20) & "RATIOS DE EVALUACION" & Chr$(10) & Chr$(10)
''    sCadImp = sCadImp & "CREDITOS MICRO Y PEQUEÑA EMPRESA" & Chr$(10)
''    sCadImp = sCadImp & String(100, "-") & Chr$(10)
''    sCadImp = sCadImp & "LIQUIDEZ"
''
''    sCadImp = sCadImp & Chr$(10)
''    sCadImp = sCadImp & ImpreFormat("Liquidez General", 20)
''    For i = nNumEvalInicio To nNumFtesRep - 1
''        nActivoCorriente = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
''        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
''                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
''                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
''
''        If nPasivo = 0 Then
''            nLiquidezGeneral = 0
''        Else
''            nLiquidezGeneral = nActivoCorriente / nPasivo
''        End If
''        sCadImp = sCadImp & ImpreFormat(nLiquidezGeneral, 10, 2)
''
''        If i > 0 And i = nNumFtesRep - 1 Then
''            nActivoCorriente = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
''            nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
''                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
''                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
''            If nPasivo = 0 Then
''                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''            Else
''                If nActivoCorriente = 0 Then
''                    sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''                Else
''                    sCadImp = sCadImp & ImpreFormat(nLiquidezGeneral / (nActivoCorriente / nPasivo), 10, 2)
''                End If
''            End If
''        End If
''
''    Next i
''
''    sCadImp = sCadImp & Chr$(10)
''    sCadImp = sCadImp & ImpreFormat("Prueba Acida", 20)
''    For i = nNumEvalInicio To nNumFtesRep - 1
''        nActivoCorriente = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
''        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
''                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
''                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
''
''        If nPasivo = 0 Then
''            nPruebaAcida = 0
''        Else
''            nPruebaAcida = (nActivoCorriente - FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios) / nPasivo
''        End If
''        sCadImp = sCadImp & ImpreFormat(nPruebaAcida, 10, 2)
''
''        If i > 0 And i = nNumFtesRep - 1 Then
''            nActivoCorriente = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
''            nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
''                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
''                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
''            If nPasivo = 0 Then
''                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''            Else
''                If (nActivoCorriente - FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios) = 0 Then
''                    sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''                Else
''                    sCadImp = sCadImp & ImpreFormat(nPruebaAcida / ((nActivoCorriente - FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios) / nPasivo), 10, 2)
''                End If
''            End If
''        End If
''    Next i
''
''    sCadImp = sCadImp & Chr$(10)
''    sCadImp = sCadImp & lnNegritaON
''    sCadImp = sCadImp & "GESTION" & Chr$(10)
''    sCadImp = sCadImp & lnNegritaOFF
''
''    sCadImp = sCadImp & ImpreFormat("Periodo Med. Cobranza", 20)
''    For i = nNumEvalInicio To nNumFtesRep - 1
''        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
''            nPeriodoMedioC = 0
''        Else
''            nPeriodoMedioC = (FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob * 365 / FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas)
''        End If
''        sCadImp = sCadImp & ImpreFormat(nPeriodoMedioC, 10, 2)
''
''        If i > 0 And i = nNumFtesRep - 1 Then
''            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
''                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''            Else
''                If (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob * 365) = 0 Then
''                    sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''                Else
''                    sCadImp = sCadImp & ImpreFormat(nPeriodoMedioC / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob * 365 / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas), 10, 2)
''                End If
''            End If
''        End If
''    Next i
''
''    sCadImp = sCadImp & Chr$(10)
''    sCadImp = sCadImp & ImpreFormat("Rotacion Invent.(Veces)", 20)
''    For i = nNumEvalInicio To nNumFtesRep - 1
''        If FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios = 0 Then
''            nRotacInvV = 0
''        Else
''            nRotacInvV = FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
''        End If
''        sCadImp = sCadImp & ImpreFormat(nRotacInvV, 10, 2)
''
''        If i > 0 And i = nNumFtesRep - 1 Then
''            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios = 0 Then
''                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''            Else
''                If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad = 0 Then
''                   sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''                Else
''                   sCadImp = sCadImp & ImpreFormat(nRotacInvV / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios), 10, 2)
''                End If
''            End If
''        End If
''    Next i
''
''    sCadImp = sCadImp & Chr$(10)
''    sCadImp = sCadImp & ImpreFormat("Margen Bruto", 20)
''    For i = nNumEvalInicio To nNumFtesRep - 1
''        nTotalIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas '+ _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam
''
''        nTotalEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad '+ _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos
''
''        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
''            nMargenBruto = 0
''        Else
''            nMargenBruto = (nTotalIngresos - nTotalEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas
''        End If
''        sCadImp = sCadImp & ImpreFormat(nMargenBruto, 10, 2)
''
''        If i > 0 And i = nNumFtesRep - 1 Then
''            nTotalIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas '+ _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam
''
''            nTotalEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad '+ _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos
''
''            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
''                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''            Else
''                If (nTotalIngresos - nTotalEgresos) = 0 Then
''                    sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''                Else
''                    sCadImp = sCadImp & ImpreFormat(nMargenBruto / ((nTotalIngresos - nTotalEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas), 10, 2)
''                End If
''            End If
''        End If
''    Next i
''
''    sCadImp = sCadImp & Chr$(10)
''    sCadImp = sCadImp & lnNegritaON
''    sCadImp = sCadImp & "SOLVENCIA O ENDEUDAMIENTO" & Chr$(10)
''    sCadImp = sCadImp & lnNegritaOFF
''
''    sCadImp = sCadImp & ImpreFormat("Endeudamiento Patrim.", 20)
''    For i = nNumEvalInicio To nNumFtesRep - 1
''        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
''                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
''                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
''        If FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio = 0 Then
''            nEndeudaPatrim = 0
''        Else
''            nEndeudaPatrim = nPasivo / FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio
''        End If
''        sCadImp = sCadImp & ImpreFormat(nEndeudaPatrim, 10, 2)
''
''        If i > 0 And i = nNumFtesRep - 1 Then
''            nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
''                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
''                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
''            If nPasivo = 0 Then
''                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''            Else
''                If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio = 0 Then
''                    sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''                Else
''                    sCadImp = sCadImp & ImpreFormat(nEndeudaPatrim / (nPasivo / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio), 10, 2)
''                End If
''            End If
''        End If
''    Next i
''    sCadImp = sCadImp & Chr$(10)
''    sCadImp = sCadImp & lnNegritaON
''    sCadImp = sCadImp & "RENTABILIDAD" & Chr$(10)
''    sCadImp = sCadImp & lnNegritaOFF
''
''    sCadImp = sCadImp & ImpreFormat("Rentab. Neta Patrim.", 20)
''    For i = nNumEvalInicio To nNumFtesRep - 1
''        If FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio = 0 Then
''            nRentabPatrim = 0
''        Else    'Utilidad_Neta/Patrimonio
''            nRentabPatrim = ((FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam)) / FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio
''        End If
''        sCadImp = sCadImp & ImpreFormat(nRentabPatrim, 10, 2)
''
''        If i > 0 And i = nNumFtesRep - 1 Then
''            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio = 0 Then
''                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''            Else
''                If ((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam)) = 0 Then
''                    sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''                Else
''                    sCadImp = sCadImp & ImpreFormat(nRentabPatrim / (((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam)) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio), 10, 2)
''                End If
''            End If
''        End If
''    Next i
''
''    sCadImp = sCadImp & Chr$(10)
''    sCadImp = sCadImp & ImpreFormat("Rentab. sobre Ventas", 20)
''    For i = nNumEvalInicio To nNumFtesRep - 1
''        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
''            nRentabVentas = 0
''        Else    'Utilidad_Neta/Ventas
''            nRentabVentas = ((FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam)) / FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas
''        End If
''        sCadImp = sCadImp & ImpreFormat(nRentabVentas, 10, 2)
''
''        If i > 0 And i = nNumFtesRep - 1 Then
''            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
''                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''            Else
''                If ((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam)) = 0 Then
''                    sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''                Else
''                    sCadImp = sCadImp & ImpreFormat(nRentabVentas / (((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam)) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas), 10, 2)
''                End If
''            End If
''        End If
''    Next i
'    'TERMINO
'    'sCadImp = sCadImp & Chr$(10) & Chr$(10)
'    'sCadImp = sCadImp & "CREDITOS COMERCIALES" & Chr$(10)
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'    sCadImp = sCadImp & "LIQUIDEZ"
'    sCadImp = sCadImp & Chr$(10)
'
'    sCadImp = sCadImp & ImpreFormat("Liquidez General", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        nActivoCorriente = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
'        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
'                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
'                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
'
'        If nPasivo = 0 Then
'            nLiquidezGeneral = 0
'        Else
'            nLiquidezGeneral = nActivoCorriente / nPasivo
'        End If
'        sCadImp = sCadImp & ImpreFormat(nLiquidezGeneral, 10, 2)
'
'        If i > 0 And i = nNumFtesRep - 1 Then
'            nActivoCorriente = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
'            nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
'                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
'                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
'            If nPasivo = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                If nActivoCorriente = 0 Then
'                    sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'                Else
'                    sCadImp = sCadImp & ImpreFormat(nLiquidezGeneral / (nActivoCorriente / nPasivo), 10, 2)
'                End If
'            End If
'        End If
'
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Prueba Acida", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        nActivoCorriente = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
'        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
'                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
'                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
'
'        If nPasivo = 0 Then
'            nPruebaAcida = 0
'        Else
'            nPruebaAcida = (nActivoCorriente - FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios) / nPasivo
'        End If
'        sCadImp = sCadImp & ImpreFormat(nPruebaAcida, 10, 2)
'
'        If i > 0 And i = nNumFtesRep - 1 Then
'            nActivoCorriente = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
'            nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
'                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
'                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
'            If nPasivo = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                If (nActivoCorriente - FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios) = 0 Then
'                    sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'                Else
'                    sCadImp = sCadImp & ImpreFormat(nPruebaAcida / ((nActivoCorriente - FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios) / nPasivo), 10, 2)
'                End If
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Capital de Trabajo", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        nActivoCorriente = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
'        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
'                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
'                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
'
'        nCapitalTrabajo = nActivoCorriente - nPasivo
'        sCadImp = sCadImp & ImpreFormat(nCapitalTrabajo, 10, 2)
'
'        If i > 0 And i = nNumFtesRep - 1 Then
'            nActivoCorriente = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
'            nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
'                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
'                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
'            If nPasivo = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                sCadImp = sCadImp & ImpreFormat(nCapitalTrabajo / (nActivoCorriente - nPasivo), 10, 2)
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & "GESTION" & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'
'    sCadImp = sCadImp & ImpreFormat("Rotac. Ctas x Cobrar", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob = 0 Then
'            nRotacionCtasCobrar = 0
'        Else
'            nRotacionCtasCobrar = FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas / FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob
'        End If
'        sCadImp = sCadImp & ImpreFormat(nRotacionCtasCobrar, 10, 2)
'
'        If i > 0 And i = nNumFtesRep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
'                    sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'                Else
'                    sCadImp = sCadImp & ImpreFormat(nRotacionCtasCobrar / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob), 10, 2)
'                End If
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Periodo Med. Cobranza", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
'            nPeriodoMedioC = 0
'        Else
'            nPeriodoMedioC = (FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob * 365 / FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas)
'        End If
'        sCadImp = sCadImp & ImpreFormat(nPeriodoMedioC, 10, 2)
'
'        If i > 0 And i = nNumFtesRep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                If (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob * 365) = 0 Then
'                    sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'                Else
'                    sCadImp = sCadImp & ImpreFormat(nPeriodoMedioC / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob * 365 / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas), 10, 2)
'                End If
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Rotac. Ctas x Pagar", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob = 0 Then
'            nRotacionCtasPagar = 0
'        Else
'            nRotacionCtasPagar = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam / FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob
'        End If
'        sCadImp = sCadImp & ImpreFormat(nRotacionCtasPagar, 10, 2)
'
'        If i > 0 And i = nNumFtesRep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam = 0 Then
'                    sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'                Else
'                    sCadImp = sCadImp & ImpreFormat(nRotacionCtasPagar / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob), 10, 2)
'                End If
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Periodo Med. Pago", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam = 0 Then
'            nPeriodoMedioP = 0
'        Else
'            nPeriodoMedioP = (FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob * 365 / FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam)
'        End If
'        sCadImp = sCadImp & ImpreFormat(nPeriodoMedioP, 10, 2)
'
'        If i > 0 And i = nNumFtesRep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                If (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob * 365) = 0 Then
'                    sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'                Else
'                    sCadImp = sCadImp & ImpreFormat(nPeriodoMedioP / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob * 365 / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam), 10, 2)
'                End If
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Rotacion Invent.(Veces)", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios = 0 Then
'            nRotacInvV = 0
'        Else
'            nRotacInvV = FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
'        End If
'        sCadImp = sCadImp & ImpreFormat(nRotacInvV, 10, 2)
'
'        If i > 0 And i = nNumFtesRep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad = 0 Then
'                   sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'                Else
'                    sCadImp = sCadImp & ImpreFormat(nRotacInvV / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios), 10, 2)
'                End If
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Rotacion Invent.(Dias)", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad = 0 Then
'            nRotacInvD = 0
'        Else
'            nRotacInvD = FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios * 365 / FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad
'        End If
'        sCadImp = sCadImp & ImpreFormat(nRotacInvD, 10, 2)
'
'        If i > 0 And i = nNumFtesRep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                If (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios * 365) = 0 Then
'                    sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'                Else
'                    sCadImp = sCadImp & ImpreFormat(nRotacInvD / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios * 365 / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad), 10, 2)
'                End If
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Rotacion Activo Fijo", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo = 0 Then
'            nRotacActivoFijo = 0
'        Else
'            nRotacActivoFijo = FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas / FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo
'        End If
'        sCadImp = sCadImp & ImpreFormat(nRotacActivoFijo, 10, 2)
'
'        If i > 0 And i = nNumFtesRep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
'                   sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'                Else
'                   sCadImp = sCadImp & ImpreFormat(nRotacActivoFijo / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo), 10, 2)
'                End If
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Margen Bruto", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        nTotalIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas '+ _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam
'
'        nTotalEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad '+ _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos
'
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
'            nMargenBruto = 0
'        Else
'            nMargenBruto = (nTotalIngresos - nTotalEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas
'        End If
'        sCadImp = sCadImp & ImpreFormat(nMargenBruto, 10, 2)
'
'        If i > 0 And i = nNumFtesRep - 1 Then
'            nTotalIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas '+ _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam
'
'            nTotalEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad '+ _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos
'
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                If (nTotalIngresos - nTotalEgresos) = 0 Then
'                    sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'                Else
'                    sCadImp = sCadImp & ImpreFormat(nMargenBruto / ((nTotalIngresos - nTotalEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas), 10, 2)
'                End If
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & "SOLVENCIA O ENDEUDAMIENTO" & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'
'    sCadImp = sCadImp & ImpreFormat("Grado Endeudam.", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        nTotalActivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo
'        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
'                    FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
'                    FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
'
'        If nTotalActivo = 0 Then
'            nGradoEndeudam = 0
'        Else
'            nGradoEndeudam = nPasivo / nTotalActivo
'        End If
'        sCadImp = sCadImp & ImpreFormat(nGradoEndeudam, 10, 2)
'
'        If i > 0 And i = nNumFtesRep - 1 Then
'            nTotalActivo = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo
'            nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
'                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
'                    FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
'
'            If nTotalActivo = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                If nPasivo = 0 Then
'                    sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'                Else
'                    sCadImp = sCadImp & ImpreFormat(nGradoEndeudam / (nPasivo / nTotalActivo), 10, 2)
'                End If
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Grado de Propiedad", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        nTotalActivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo
'
'        If nTotalActivo = 0 Then
'            nGradoPropiedad = 0
'        Else
'            nGradoPropiedad = FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio / nTotalActivo
'        End If
'        sCadImp = sCadImp & ImpreFormat(nGradoPropiedad, 10, 2)
'
'        If i > 0 And i = nNumFtesRep - 1 Then
'            nTotalActivo = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + _
'                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo
'
'            If nTotalActivo = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio = 0 Then
'                    sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'                Else
'                    sCadImp = sCadImp & ImpreFormat(nGradoPropiedad / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio / nTotalActivo), 10, 2)
'                End If
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Endeudamiento Patrim.", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
'                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
'                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio = 0 Then
'            nEndeudaPatrim = 0
'        Else
'            nEndeudaPatrim = nPasivo / FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio
'        End If
'        sCadImp = sCadImp & ImpreFormat(nEndeudaPatrim, 10, 2)
'
'        If i > 0 And i = nNumFtesRep - 1 Then
'            nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
'                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
'                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                If nPasivo = 0 Then
'                    sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'                Else
'                    sCadImp = sCadImp & ImpreFormat(nEndeudaPatrim / (nPasivo / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio), 10, 2)
'                End If
'            End If
'        End If
'    Next i
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & "RENTABILIDAD" & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'
'    sCadImp = sCadImp & ImpreFormat("Rentab. Neta Patrim.", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio = 0 Then
'            nRentabPatrim = 0
'        Else    'Utilidad_Neta/Patrimonio
'            nRentabPatrim = ((FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam)) / FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio
'        End If
'        sCadImp = sCadImp & ImpreFormat(nRentabPatrim, 10, 2)
'
'        If i > 0 And i = nNumFtesRep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                If ((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam)) = 0 Then
'                    sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'                Else
'                    sCadImp = sCadImp & ImpreFormat(nRentabPatrim / (((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam)) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio), 10, 2)
'                End If
'            End If
'        End If
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & ImpreFormat("Rentab. sobre Ventas", 20)
'    For i = nNumEvalInicio To nNumFtesRep - 1
'        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
'            nRentabVentas = 0
'        Else    'Utilidad_Neta/Ventas
'            nRentabVentas = ((FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam)) / FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas
'        End If
'        sCadImp = sCadImp & ImpreFormat(nRentabVentas, 10, 2)
'
'        If i > 0 And i = nNumFtesRep - 1 Then
'            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
'                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'            Else
'                If ((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam)) = 0 Then
'                    sCadImp = sCadImp & ImpreFormat(0, 10, 2)
'                Else
'                    sCadImp = sCadImp & ImpreFormat(nRentabVentas / (((FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob) - (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam)) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas), 10, 2)
'                End If
'            End If
'        End If
'    Next i
'
'
''    sCadImp = sCadImp & String(100, "-") & Chr$(10)
''    sCadImp = sCadImp & lnNegritaOFF
''
''    'Activo Circulante / Pasivo
''    Dim nActCir As Double
''    Dim nPasivo As Double
''    Dim nActCirAnt As Double
''    Dim nPasivoAnt As Double
''    Dim Res1 As Double
''    Dim Res2 As Double
''    sCadImp = sCadImp & ImpreFormat("Liquidez", 20)
''    For i = nNumEvalInicio To nNumFtesRep - 1
''        nActCir = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
''                            FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
''                            FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
''        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
''                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
''                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
''        If nPasivo = 0 Then
''            Res1 = 0
''        Else
''            Res1 = nActCir / nPasivo
''        End If
''        sCadImp = sCadImp & ImpreFormat(Res1, 10, 2)
''        If i > 0 And i = nNumFtesRep - 1 Then
''            nActCirAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
''                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
''                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
''            nPasivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
''                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
''                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
''            If nPasivoAnt = 0 Then
''                Res2 = 0
''            Else
''                Res2 = nActCirAnt / nPasivoAnt
''            End If
''            If Res2 = 0 Then
''                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''            Else
''                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 10, 2)
''            End If
''        End If
''    Next i
''
''    sCadImp = sCadImp & Chr$(10)
''    'Activo Circulante - Pasivo
''    sCadImp = sCadImp & ImpreFormat("Capital Trabajo", 20)
''    For i = nNumEvalInicio To nNumFtesRep - 1
''        nActCir = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
''                            FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
''                            FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
''        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
''                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
''                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
''        Res1 = nActCir - nPasivo
''        sCadImp = sCadImp & ImpreFormat(Res1, 10, 2)
''        If i > 0 And i = nNumFtesRep - 1 Then
''            nActCirAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
''                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
''                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
''            nPasivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
''                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
''                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
''            Res2 = nActCirAnt - nPasivoAnt
''            If Res2 = 0 Then
''                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''            Else
''                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 10, 2)
''            End If
''        End If
''    Next i
''    sCadImp = sCadImp & Chr$(10)
''    sCadImp = sCadImp & lnNegritaON
''    sCadImp = sCadImp & "GESTION" & Chr$(10)
''    sCadImp = sCadImp & lnNegritaOFF
''    'Costo de Ventas / Inventario
''    sCadImp = sCadImp & ImpreFormat("Rotacion Inventario", 20)
''    For i = nNumEvalInicio To nNumFtesRep - 1
''        If FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios = 0 Then
''            Res1 = 0
''        Else
''            Res1 = FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios
''        End If
''        sCadImp = sCadImp & ImpreFormat(Res1, 10, 2)
''
''        If i > 0 And i = nNumFtesRep - 1 Then
''            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios = 0 Then
''                Res2 = 0
''            Else
''                Res2 = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios
''            End If
''            If Res2 = 0 Then
''                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''            Else
''                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 10, 2)
''            End If
''        End If
''    Next i
''    sCadImp = sCadImp & Chr$(10)
''
''
''    'Recup Ctas x Cob / (Ventas / 30)
''    sCadImp = sCadImp & ImpreFormat("Periodo de Cobro", 20)
''    For i = nNumEvalInicio To nNumFtesRep - 1
''        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
''            Res1 = 0
''        Else
''            Res1 = FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob / (FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas / 30)
''        End If
''        sCadImp = sCadImp & ImpreFormat(Res1, 10, 2)
''        If i > 0 And i = nNumFtesRep - 1 Then
''            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
''                Res2 = 0
''            Else
''                Res2 = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob / (FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas / 30)
''            End If
''            If Res2 = 0 Then
''                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''            Else
''                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 10, 2)
''            End If
''        End If
''    Next i
''    sCadImp = sCadImp & Chr$(10)
''
''    Dim nIngresos As Double
''    Dim nEgresos As Double
''    Dim nIngresosAnt As Double
''    Dim nEgresosAnt As Double
''
''    sCadImp = sCadImp & lnNegritaON
''    sCadImp = sCadImp & "RENTABILIDAD" & Chr$(10)
''    sCadImp = sCadImp & lnNegritaOFF
''    'Utilidad Operat =  Ingresos - Egresos
''    'Utilidad Operat / Patrimonio
''    sCadImp = sCadImp & ImpreFormat("Rentab. Patrimonial", 20)
''    For i = nNumEvalInicio To nNumFtesRep - 1
''        nIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
''        nEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos
''
''        If FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio = 0 Then
''            Res1 = 0
''        Else
''            Res1 = (nIngresos - nEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio
''        End If
''        sCadImp = sCadImp & ImpreFormat(Res1, 10, 2)
''        If i > 0 And i = nNumFtesRep - 1 Then
''            nIngresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob
''            nEgresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos
''            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio = 0 Then
''                Res2 = 0
''            Else
''                Res2 = (nIngresos - nEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio
''            End If
''
''            If Res2 = 0 Then
''                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''            Else
''                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 10, 2)
''            End If
''        End If
''    Next i
''    sCadImp = sCadImp & Chr$(10)
''
''    Dim nTotActivo As Double
''    Dim nTotActivoAnt As Double
''    sCadImp = sCadImp & ImpreFormat("Rentab. Activo", 20)
''    'Utilidad Operativa / Total Activo
''    For i = nNumEvalInicio To nNumFtesRep - 1
''        nIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
''        nEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos
''        nTotActivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo
''        If nTotActivo = 0 Then
''            Res1 = 0
''        Else
''            Res1 = (nIngresos - nEgresos) / nTotActivo
''        End If
''        sCadImp = sCadImp & ImpreFormat(Res1, 10, 2)
''        If i > 0 And i = nNumFtesRep - 1 Then
''            nIngresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob
''            nEgresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos
''            nTotActivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
''                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
''                            FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActFijo
''            If nTotActivoAnt = 0 Then
''                Res2 = 0
''            Else
''                Res2 = (nIngresosAnt - nEgresosAnt) / nTotActivoAnt
''            End If
''            If Res2 = 0 Then
''                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''            Else
''                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 10, 2)
''            End If
''        End If
''    Next i
''    sCadImp = sCadImp & Chr$(10)
''
''    sCadImp = sCadImp & ImpreFormat("Rentab. Ventas", 20)
''    'Utilidad Operativa / Ventas
''    For i = nNumEvalInicio To nNumFtesRep - 1
''        nIngresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
''        nEgresos = FtesIngreso(pnIndex).FtesIngresoInd(i).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i).nOtrosEgresos
''        If FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas = 0 Then
''            Res1 = 0
''        Else
''            Res1 = (nIngresos - nEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas
''        End If
''        sCadImp = sCadImp & ImpreFormat(Res1, 10, 2)
''        If i > 0 And i = nNumFtesRep - 1 Then
''            nIngresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nRecupCtasxCob
''            nEgresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nOtrosEgresos
''            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas = 0 Then
''                Res2 = 0
''            Else
''                Res2 = (nIngresosAnt - nEgresosAnt) / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas
''            End If
''            If Res2 = 0 Then
''                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''            Else
''                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 10, 2)
''            End If
''        End If
''    Next i
''    sCadImp = sCadImp & Chr$(10)
''
''    'Ventas Totales = Ventas + Recup Ctas x Cob
''    'Margen Utilidad = Ventas Totales - Costo Ventas / Ventas Totales
''    Dim nVentasTot As Double
''    Dim nVentasTotAnt As Double
''    sCadImp = sCadImp & ImpreFormat("Margen Utilidad", 20)
''    For i = nNumEvalInicio To nNumFtesRep - 1
''        nVentasTot = FtesIngreso(pnIndex).FtesIngresoInd(i).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
''        If nVentasTot = 0 Then
''            Res1 = 0
''        Else
''            Res1 = (nVentasTot - FtesIngreso(pnIndex).FtesIngresoInd(i).nCompraMercad) / nVentasTot
''        End If
''        sCadImp = sCadImp & ImpreFormat(Res1, 10, 2)
''        If i > 0 And i = nNumFtesRep - 1 Then
''            nVentasTotAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(i).nRecupCtasxCob
''            If nVentasTotAnt = 0 Then
''                Res2 = 0
''            Else
''                Res2 = (nVentasTotAnt - FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCompraMercad) / nVentasTotAnt
''            End If
''            If Res2 = 0 Then
''                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''            Else
''                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 10, 2)
''            End If
''        End If
''    Next i
''    sCadImp = sCadImp & Chr$(10)
''
''    sCadImp = sCadImp & lnNegritaON
''    sCadImp = sCadImp & "ENDEUDAMIENTO" & Chr$(10)
''    sCadImp = sCadImp & lnNegritaOFF
''    sCadImp = sCadImp & ImpreFormat("Endeudamiento Pat.", 20)
''    'Endeudamiento = Pasivo / Patrimonio
''    For i = nNumEvalInicio To nNumFtesRep - 1
''        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
''                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
''                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
''
''        If FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio = 0 Then
''            Res1 = 0
''        Else
''            Res1 = nPasivo / FtesIngreso(pnIndex).FtesIngresoInd(i).nPasivoPatrimonio
''        End If
''        sCadImp = sCadImp & ImpreFormat(Res1, 10, 2)
''        If i > 0 And i = nNumFtesRep - 1 Then
''            nPasivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
''                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
''                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
''
''            If FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio = 0 Then
''                Res2 = 0
''            Else
''                Res2 = nPasivoAnt / FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nPasivoPatrimonio
''            End If
''            If Res2 = 0 Then
''                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''            Else
''                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 10, 2)
''            End If
''        End If
''    Next i
''    sCadImp = sCadImp & Chr$(10)
''
''    sCadImp = sCadImp & ImpreFormat("Endeudamiento", 20)
''    'Endeudamiento = Pasivo / TotalActivo
''    For i = nNumEvalInicio To nNumFtesRep - 1
''        nTotActivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nActivDisp + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nCtasxCob + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo
''        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(i).nProveedores + _
''                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosOtros + _
''                                FtesIngreso(pnIndex).FtesIngresoInd(i).nCreditosCmact
''
''        If nTotActivo = 0 Then
''            Res1 = 0
''        Else
''            Res1 = nPasivo / nTotActivo
''        End If
''        sCadImp = sCadImp & ImpreFormat(Res1, 10, 2)
''        If i > 0 And i = nNumFtesRep - 1 Then
''            nTotActivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nActivDisp + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCtasxCob + _
''                        FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(i).nActFijo
''            nPasivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nProveedores + _
''                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosOtros + _
''                                FtesIngreso(pnIndex).FtesIngresoInd(i - 1).nCreditosCmact
''            If nTotActivoAnt = 0 Then
''                Res2 = 0
''            Else
''                Res2 = nPasivoAnt / nTotActivoAnt
''            End If
''            If Res2 = 0 Then
''                sCadImp = sCadImp & ImpreFormat(0, 10, 2)
''            Else
''                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 10, 2)
''            End If
''        End If
''    Next i
''
''
'
'
''*****************************************************************************************
''*****************************************************************************************
''*****************************************************************************************


'*** COMENTADO PEAC 20080402

'    If pbCostoProd = True Then
'
'        'IMPRIME COSTO DE PRODUCCION
'        '-------------------------------------
'        sCadImp = sCadImp & Chr$(12)
'        sCadImp = sCadImp & String(100, "-") & Chr$(10)
'        sCadImp = sCadImp & Space(30) & "COSTO DE PRODUCCION" & Chr$(10)
'        sCadImp = sCadImp & String(100, "-") & Chr$(10)
'        sCadImp = sCadImp & ImpreFormat("Evaluacion", 16)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).dFecEval, 14, 0)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ImpreFormat("Tipo Cultivo", 20)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).sTpoCultivo, 13)
'        Next i
'
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ImpreFormat("Destino", 20)
'
'        Dim sDestino As String
'        bProd = False
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            If FtesIngreso(pnIndex).CostoProd(i).nSiembra = 1 Then
'    '            If Len(sDestino) = 0 Then
'    '                sDestino = "SIEMBRA"
'    '            Else
'    '                sDestino = sDestino & ",SIEMBRA"
'                    bProd = True
'                    sCadImp = sCadImp & ImpreFormat("SIEMBRA", 13)
'                    'sCadImp = sCadImp & Chr$(10)
'    '            End If
'             End If
'
'            If FtesIngreso(pnIndex).CostoProd(i).nMantenimiento = 1 Then
'    '            If Len(sDestino) = 0 Then
'    '                sDestino = "MANTENIMIENTO"
'    '            Else
'    '                sDestino = sDestino & ",MANTENIMIENTO"
'    '                sCadImp = sCadImp & Chr$(10)
'    '            End If
'                If bProd = True Then
'                    sCadImp = sCadImp & Space(7) & ImpreFormat("MANTENIMIENTO", 13)
'                Else
'                    sCadImp = sCadImp & ImpreFormat("MANTENIMIENTO", 13)
'                    bProd = True
'                End If
'                   ' sCadImp = sCadImp & Chr$(10)
'            End If
'            If FtesIngreso(pnIndex).CostoProd(i).nDesAgricola = 1 Then
'    '             If Len(sDestino) = 0 Then
'    '                sDestino = "DESAGRICOLA"
'    '             Else
'    '                sDestino = sDestino & ",DESAGRICOLA"
'    '                sCadImp = sCadImp & Chr$(10)
'    '             End If
'                  If bProd = True Then
'                    sCadImp = sCadImp & Space(7) & ImpreFormat("DESAGRICOLA", 13)
'                  Else
'                    sCadImp = sCadImp & ImpreFormat("DESAGRICOLA", 13)
'                    bProd = True
'                  End If
'                    'sCadImp = sCadImp & Chr$(10)
'
'
'            End If
'            If FtesIngreso(pnIndex).CostoProd(i).nCosecha = 1 Then
'    '            If Len(sDestino) = 0 Then
'    '                sDestino = "COSECHA"
'    '            Else
'    '                sDestino = sDestino & ",COSECHA"
'    '                sCadImp = sCadImp & Chr$(10)
'    '            End If
'                If bProd = True Then
'                    sCadImp = sCadImp & Space(7) & ImpreFormat("COSECHA", 13)
'                Else
'                    sCadImp = sCadImp & ImpreFormat("COSECHA", 13)
'                    bProd = True
'                End If
'                    'sCadImp = sCadImp & Chr$(10)
'
'             End If
'            If FtesIngreso(pnIndex).CostoProd(i).nPlanOtros = 1 Then
'    '            If Len(sDestino) = 0 Then
'    '                sDestino = "OTROS"
'    '            Else
'    '                sDestino = sDestino & ",OTROS"
'    '            End If
'                    If bProd = True Then
'                        sCadImp = sCadImp & Space(7) & ImpreFormat("OTROS", 13)
'                    Else
'                        sCadImp = sCadImp & ImpreFormat("OTROS", 13)
'                        bProd = True
'                    End If
'                    'sCadImp = sCadImp & Chr$(10)
'
'            End If
'            'sCadImp = sCadImp & ImpreFormat(sDestino, 60)
'        Next i
'
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ImpreFormat("Maquinaria", 12)
'
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nMaquinaria, 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ImpreFormat("Jornales", 12)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nJornales, 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ImpreFormat("Insumos", 12)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nInsumos, 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ImpreFormat("Pesticidas", 12)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPesticidas, 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ImpreFormat("Otros", 12)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nOtros, 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ImpreFormat("Total", 12)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros, 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10) & Chr$(10)
'
'        sCadImp = sCadImp & ImpreFormat("Hectareas", 12)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas, 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ImpreFormat("Produccion", 18)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ImpreFormat(Format(FtesIngreso(pnIndex).CostoProd(i).nProduccion, "#0.00") & " " & FtesIngreso(pnIndex).CostoProd(i).sUniProd, 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ImpreFormat("Precio Uni.", 12)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nPreUni, 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ImpreFormat("Ingresos", 12)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas * FtesIngreso(pnIndex).CostoProd(i).nPreUni * FtesIngreso(pnIndex).CostoProd(i).nProduccion, 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ImpreFormat("Egresos", 12)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).CostoProd(i).nHectareas * (FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros), 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'
'        sCadImp = sCadImp & ImpreFormat("Utilidad", 12)
'        For i = 0 To FtesIngreso(pnIndex).nNumCostosProd - 1
'            sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).CostoProd(i).nHectareas * FtesIngreso(pnIndex).CostoProd(i).nPreUni * FtesIngreso(pnIndex).CostoProd(i).nProduccion) - (FtesIngreso(pnIndex).CostoProd(i).nHectareas * (FtesIngreso(pnIndex).CostoProd(i).nPesticidas + FtesIngreso(pnIndex).CostoProd(i).nInsumos + FtesIngreso(pnIndex).CostoProd(i).nJornales + FtesIngreso(pnIndex).CostoProd(i).nMaquinaria + FtesIngreso(pnIndex).CostoProd(i).nOtros)), 13)
'        Next i
'        sCadImp = sCadImp & Chr$(10)
'        sCadImp = sCadImp & String(100, "-") & Chr$(10)
' End If

'FIN COMENTADO POR PEAC

'***********************************----------------------------------********************************
 
    ' ---------------------------------------------------------------------------------------------------------------
    ' Muestra los datos de las referencias personales de la fte como persona ----------------------------------------
'    sCadImp = sCadImp & Chr$(10)
'    sCadImp = sCadImp & String(100, "-") & Chr$(10)
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & "REFERENCIAS PERSONALES" & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'
'    If Len(FtesIngreso(pnIndex).sComentarios) < 100 Then
'        sCadImp = sCadImp & FtesIngreso(pnIndex).sComentarios & Chr$(10)
'    Else
'        'sCadImp = sCadImp & FtesIngreso(pnIndex).sComentarios & Chr$(10)
'        For i = 0 To (Len(FtesIngreso(pnIndex).sComentarios) / 100) - 1
'            sCadImp = sCadImp & Mid(FtesIngreso(pnIndex).sComentarios, (i * 100) + 1, 100) & Chr$(10)
'        Next i
'    End If
'
'    sCadImp = sCadImp & Chr$(10)

    ' Muestra los datos de las referencias bancarias de la fte como persona ----------------------------------------
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ImpreFormat("REFERENCIAS BANCARIAS", 42, 0) & ImpreFormat("NRO.TARJETA", 22, 2) & ImpreFormat("NRO.CUENTA", 22, 2) & ImpreFormat("MONTO", 10, 2) & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF

'    MatDatosFte = oPersona.RecuperaRefBancariaPersFte(FtesIngreso(pnIndex).sFteIngreso)
'
'    For i = 0 To UBound(MatDatosFte) - 1
'        sCadImp = sCadImp & ImpreFormat(MatDatosFte(i, 0) & "-" & MatDatosFte(i, 1), 40) & ImpreFormat(MatDatosFte(i, 3), 22, 2) & ImpreFormat(MatDatosFte(i, 2), 22, 2) & ImpreFormat(MatDatosFte(i, 4), 10, 2) & Chr$(10)
'    Next i
'
'    sCadImp = sCadImp & Chr$(10)

'   ' Muestra los datos de las referencias comerciales de la fte como persona ----------------------------------------
'    sCadImp = sCadImp & lnNegritaON
'    sCadImp = sCadImp & ImpreFormat("REFERENCIAS COMERCIALES", 66, 0) & ImpreFormat("TIPO RELAC. COMERCIAL", 22, 2) & ImpreFormat("TELEFONO", 10, 2) & Chr$(10)
'    sCadImp = sCadImp & lnNegritaOFF
'
'    MatDatosComFte = oPersona.RecuperaRefComercialPersFte(FtesIngreso(pnIndex).sFteIngreso)
'
'    For i = 0 To UBound(MatDatosComFte) - 1
'        sCadImp = sCadImp & ImpreFormat(MatDatosComFte(i, 3) & "-" & MatDatosComFte(i, 0), 64) & ImpreFormat(MatDatosComFte(i, 2), 22, 2) & ImpreFormat(MatDatosComFte(i, 1), 10, 2) & Chr$(10)
'    Next i

    ' ---------------------------------------------------------------------------------------------------------------
    
    
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & Space(20) & "REFERENCIAS PERSONALES / COMERCIALES" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
     sCadImp = sCadImp & String(100, "-") & Chr$(10)
     
    sCadImp = sCadImp & "   NOMBRE/REF.COMERCIAL            TIPO REF. TELEFONO     COMENTARIO              DIRECCION" & Chr$(10)
     sCadImp = sCadImp & String(100, "-") & Chr$(10)
    
Do While Not rsRef.EOF
    sCadImp = sCadImp & ImpreFormat(UCase(rsRef!cNomRefCom), 30) & ImpreFormat(rsRef!cConsDescripcion, 10) & ImpreFormat(rsRef!cFonoRefCom, 10) & ImpreFormat(UCase(rsRef!cComentarioRefCom), 20) & ImpreFormat(UCase(rsRef!cDireccionRefCom), 20)
    sCadImp = sCadImp & Chr$(10)
    rsRef.MoveNext
Loop
    
    sCadImp = sCadImp & Chr$(10)
    
    Set oPersona = Nothing
    
    GenerarImpresionFteIngresoIndependiente_CS = sCadImp
    
End Function

'--------------

'RECO20140226 ERS174-2013*****************************************************

Public Function ValidaExisteFICred(ByVal psNumFuente As String, ByVal psCtaCod As String) As Boolean
    Dim oPersona As COMDPersona.DCOMPersonas
    Dim oDR As ADODB.Recordset
    Set oPersona = New COMDPersona.DCOMPersonas
    Set oDR = New ADODB.Recordset
    Set oDR = oPersona.ObtieneDatosFICred(psNumFuente, psCtaCod)
    
    Set oPersona = Nothing
        If Not oDR.BOF And Not oDR.EOF Then
            ValidaExisteFICred = True
        Else
            ValidaExisteFICred = False
        End If
End Function
'RECO FIN*********************************************************************
'FRHU 20140609 TI-ERS068-2014
Public Function RecuperarDetalleNombrePersona(ByVal psPersNombre As String, ByVal pnTipo As Integer) As String
    Select Case pnTipo
    Case 1: RecuperarDetalleNombrePersona = BuscaNombre(psPersNombre, BusqApellidoPaterno) 'Apellido Paterno
    Case 2: RecuperarDetalleNombrePersona = BuscaNombre(psPersNombre, BusqApellidoMaterno) 'Apellido Materno
    Case 3: RecuperarDetalleNombrePersona = BuscaNombre(psPersNombre, BusqNombres) 'Nombre
    End Select
End Function
'FIN FRHU 20140609

'************** CTI3 *********************
Public Function validallenado(ByVal grilla As FlexEdit) As Boolean

Dim i As Integer

For i = 1 To grilla.Rows - 1
 
 If grilla.TextMatrix(i, 1) = "" Then
   
   validallenado = False
   Exit Function
 End If

Next i

validallenado = True
End Function
'************** CTI3 *********************08112018
Public Function IsFormLoaded(FormToCheck As Form) As Integer
Dim Y As Integer

For Y = 0 To Forms.count - 1
If Forms(Y) Is FormToCheck Then
    IsFormLoaded = True
    Exit Function
    End If
Next
IsFormLoaded = False
End Function
'************** CTI3 *********************08112018
Public Function LimpiarGrillasAccionariales(gAccio As FlexEdit, gDirec As FlexEdit, gGere As FlexEdit, _
gPatri As FlexEdit, gPatEmp As FlexEdit, gCargos As FlexEdit) As Boolean

Dim i, nFila1, nFila2, nFila3, nFila4, nFila5, nfila6 As Integer

nFila1 = gAccio.Cols: nFila2 = gDirec.Cols: nFila3 = gGere.Cols: nFila4 = gPatri.Cols
nFila5 = gPatEmp.Cols: nfila6 = gCargos.Cols

gAccio.Rows = 2
For i = 0 To nFila1 - 1
    gAccio.TextMatrix(1, i) = ""
Next i
gDirec.Rows = 2
For i = 0 To nFila2 - 1
    gDirec.TextMatrix(1, i) = ""
Next i
gGere.Rows = 2
For i = 0 To nFila3 - 1
    gGere.TextMatrix(1, i) = ""
Next i
gPatri.Rows = 2
For i = 0 To nFila4 - 1
    gPatri.TextMatrix(1, i) = ""
Next i
gPatEmp.Rows = 2
For i = 0 To nFila5 - 1
    gPatEmp.TextMatrix(1, i) = ""
Next i
gCargos.Rows = 2
For i = 0 To nfila6 - 1
    gCargos.TextMatrix(1, i) = ""
Next i

End Function



