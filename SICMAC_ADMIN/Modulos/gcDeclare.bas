Attribute VB_Name = "gcDeclare"
'Global Const gsConexion = "PROVIDER=SQLOLEDB;uid=DBAccess;pwd=cmact;DATABASE=DBCmact;SERVER=07Srv_Developer"
'Global Const gsRUC = "20132243230"
'Global Const gsEmpresa = "CAJA MUNICIPAL DE TRUJILLO"
'Global Const vsServerComunes = ""
'Global Const vsServerPersona = ""
'Global Const vsServerAdministracion = ""
'Global Const vsServerNegocio = ""
'Global Const vsServerImagenes = ""
'Global Const gcPDC = "\\SRVROOT"
'Global Const gcDominio = "CMACTRUJILLO"
'Global Const gcWINNT = "WinNT://"

Global gsConexion As String
Global gsRUC As String
Global gsEmpresa As String
Global gsEmpresaCompleto As String
Global gsEmpresaRazonSocial As String
Global gsEmpresaDireccion As String
Global vsServerComunes As String
Global vsServerPersona As String
Global vsServerAdministracion As String
Global vsServerNegocio As String
Global vsServerImagenes As String
Global gcPDC As String
Global gcDominio As String
Global gcWINNT As String
Global gbBitCentral As Boolean
Global gbBitTCPonderado As Boolean
Global gbBitIGVCredFiscal As Boolean
Global gcPC As String

'MAVM : Inventario
Global gsMvoNro As String
Global gsMonedaAF As String '*** PEAC 20100506


'PARA NEGOCIO
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Const CB_FINDSTRING = &H14C
Global Const gsConnServDBF = "DSN=DSNCmactServ"
Global gsCentralPers As String


Public Enum RHContratoMantTpo
    RHContratoMantTpoCargo = 0
    RHContratoMantTpoSueldo = 2
    RHContratoMantTpoSisPens = 3
    RHContratoMantTpoAMP = 1
    RHContratoMantTpoComentario = 4
    RHContratoMantTpoAdenda = 5
    RHContratoMantTpoFoto = 7
End Enum

Public Enum TipoOpe
    gTipoOpeRegistro = 0
    gTipoOpeMantenimiento = 1
    gTipoOpeConsulta = 2
    gTipoOpeReporte = 3
End Enum

Public Enum TipoProcesoRRHH
    gTipoProcesoRRHHCalculo = 0
    gTipoProcesoRRHHAbono = 1
    gTipoProcesoRRHHCierreMes = 2
    gTipoProcesoRRHHCierreDia = 3
    gTipoProcesoRRHHConsulta = 4
End Enum

'6022
Global Const gRHEvaluacionComite = 6022
Public Enum RHEvaluacionComite
    gEvalPresidente = 0
    gEvalSecretario = 1
    gEvalCoordinador = 2
End Enum

'6023
Global Const gGenTipoPeriodos = 1012
Public Enum GenTipoPeriodos
    gGenTipoPerAÑO = 0
    gGenTipoPerSEMESTRE = 1
    gGenTipoPerTRIMESTRE = 2
    gGenTipoPerBIMESTRE = 3
    gGenTipoPerMES = 4
    gGenTipoPerQUINCENA = 5
    gGenTipoPerSEMANA = 6
    gGenTipoPerDIA = 7
End Enum

Global Const gRHTipoOpeEvaluacion = 6023
Public Enum RHTipoOpeEvaluacion
    RHTipoOpeEvaEscrito = 0
    RHTipoOpeEvaPsicologico = 1
    RHTipoOpeEvaEntrevista = 2
    RHTipoOpeEvaCurricular = 3
    RHTipoOpeEvaConsolidado = 4
End Enum

Global Const gRHTipoInformeSocial = 6024
Public Enum RHTipoInformeSocial
    TpoInfSocVisitaDomiXSalud = 0
    TpoInfSocSeguimiento = 1
    TpoInfSocNuevoTrabajador = 2
End Enum

Global Const gRHPeriodoNoLab = 6025
Public Enum RHPeriodoNoLab
    RHPeriodoNoLabSolicitado = 0
    RHPeriodoNoLabAprovado = 1
    RHPeriodoNoLabRechazado = 2
End Enum


Global Const gRHAutoFisicaGrupo = 6026
Public Enum RHAutoFisicaGrupo
    RHGrupoAutoFisicaVacaciones = 1
    RHGrupoAutoFisicaDescansos = 2
    RHGrupoAutoFisicaPermisos = 3
    RHGrupoAutoFisicaSanciones = 4
End Enum

Global Const gRHAutoFisicaTipoEstado = 6027
Public Enum RHAutoFisicaTipoEstado
    RHGrupoAutoFisicaSolicitada = 0
    RHGrupoAutoFisicaProgramada = 1
    RHGrupoAutoFisicaEjecutada = 2
    RHGrupoAutoFisicaRechazada = 3
    RHGrupoAutoFisicaAprovada = 4
End Enum

Global Const gRHConceptoTipoCal = 6028
Public Enum RHConceptoTipoCal
    RHConceptoTipo_CONSTANTES = 1
    RHConceptoTipo_VALORCALCULADO = 2
    RHConceptoTipo_VALOR_PRE_DEFINIDO = 3
    RHConceptoTipo_FUNCIONES_INTERNAS = 4
End Enum

Global Const gRHOperadorAritmet = 6029
Public Enum RHOperadorAritmet
    RHOperadorAritmetSUMA = 1
    RHOperadorAritmetRESTA = 2
    RHOperadorAritmetMULTIPLICACION = 3
    RHOperadorAritmetDIVISION = 4
End Enum
    
Global Const gRHOperadorTexto = 6030
Public Enum RHOperadorTexto
    RHOperadorTexto = 1
End Enum
    
Global Const gRHOperadorLogicos = 6031
Public Enum RHOperadorLogicos
    RHOperadorLogicosPARENTISIS = 1
    RHOperadorLogicosMAYOR = 2
    RHOperadorLogicosMAYOR_O_IGUAL = 3
    RHOperadorLogicosMENOR = 4
    RHOperadorLogicosMENOR_O_IGUAL = 5
    RHOperadorLogicosIGULA = 6
    RHOperadorLogicosDECISION_SI = 7
    RHOperadorLogicosVERDADERO = 8
    RHOperadorLogicosFALSO = 9
    RHOperadorLogicosNEGACION_NOT = 10
    RHOperadorLogicosY_AND = 11
    RHOperadorLogicosO_OR = 12
End Enum

Global Const gRHOperadorFecha = 6033
Public Enum RHOperadorFecha
    RHOperadorLogicosPARENTISIS = 1
    RHOperadorLogicosMAYOR = 2
    RHOperadorLogicosMAYOR_O_IGUAL = 3
    RHOperadorLogicosMENOR = 4
    RHOperadorLogicosMENOR_O_IGUAL = 5
    RHOperadorLogicosIGULA = 6
    RHOperadorLogicosDECISION_SI = 7
    RHOperadorLogicosVERDADERO = 8
    RHOperadorLogicosFALSO = 9
    RHOperadorLogicosNEGACION_NOT = 10
    RHOperadorLogicosY_AND = 11
    RHOperadorLogicosO_OR = 12
End Enum

Global Const gRHConceptosTpoVisible = 6011
Public Enum RHConceptosTpoVisible
    RHConceptosTpoVIngreso = 1
    RHConceptosTpoVEgreso = 2
    RHConceptosTpoVAportacion = 3
    RHConceptosTpoVVarUsuario = 4
    RHConceptosTpoVTodos = 5
End Enum

Global Const gRHContratoTipo = 6012
Public Enum RHContratoTipo
    RHContratoTipoIndeterminado = 0
    RHContratoTipoFijo = 1
    RHContratoTipoLocacion = 2
    RHContratoTipoFLaboral = 3
    RHContratoTipoPractica = 4
    RHContratoTipoSesigrista = 5
    RHContratoTipoDirector = 6
End Enum

Global Const gRHEmpleadoFonfoTipo = 6013
Public Enum RHEmpleadoFonfoTipo
    RHEmpleadoFonfoTipoAFP = 0
    RHEmpleadoFonfoTipoSNP = 1
End Enum

Public Enum RHPSeleccionTpoMnu
    RHPSeleccionTpoMnuSel = 0
    RHPSeleccionTpoMnuPost = 1
    RHPSeleccionTpoMnuEvaCur = 2
    RHPSeleccionTpoMnuEvaEsc = 3
    RHPSeleccionTpoMnuEvaPsi = 4
    RHPSeleccionTpoMnuEvaEnt = 5
    RHPSeleccionTpoMnuResultado = 6
    RHPSeleccionTpoMnuConfirmacion = 7
    RHPSeleccionTpoMnuCierre = 8
End Enum

Global Const gRHEmpleadoTurno = 6014
Public Enum RHEmpleadoTurno
    RHEmpleadoTurnoUno = 1
    RHEmpleadoTurnoDos = 2
End Enum

Public Enum ContratoForma
    ContratoFormaAutomatica = 0
    ContratoFormaManual = 1
End Enum

Global Const gRHEstadosTpo = 6015
Public Enum RHEstadosTpo
    RHEstadosTpoInactivo = 1
    RHEstadosTpoActivo = 2
    RHEstadosTpoVacaciones = 3
    RHEstadosTpoPermisosLicencias = 4
    RHEstadosTpoSubsidiado = 5
    RHEstadosTpoSuspendido = 6
    RHEstadosTpoRetirado = 7
    RHEstadosTpoPorLiquidar = 8
End Enum

Global Const gRHPerNoLab = 6016
Public Enum RHPerNoLab
    RHPerNoLabAprovado = 1
    RHPerNoLabRechazado = 2
End Enum

Global Const gRHPeriodosTpo = 6036
Public Enum RHPeriodosTpo
    RHPeriodosTpoTiempo = 0
    RHPeriodosTpoPeridos = 1
    RHPeriodosTpoTiempoPeridos = 2
End Enum

Global Const gRHPlanillaEstado = 6040
Public Enum RHPlanillaEstado
    RHPlanillaEstadoGenerado = 1
    RHPlanillaEstadoPagado = 2
End Enum
 
Global Const gRHProfesionesCurr = 6044
Global Const gRHNivelCurr = 6045

Public Enum TipoControl
    TipoControlNinguno = 0
    TipoControlAsistencia = 1
    TipoControlVacaciones = 2
    TipoControlPermisos = 3
End Enum

Public Enum RHEmpleadoCuentasTpo
    RHEmpleadoCuentasTpoAhorro = 232
    RHEmpleadoCuentasTpoCTS = 234
End Enum

Public Enum RHExtraPlanillaOpeTpo
    RHExtraPlanillaOpeTpoCargo = 0
    RHExtraPlanillaOpeTpoAbono = 1
End Enum

Public Enum RHEstadosRRHH
    gRHEstadosRRHHINACTFINCONTRATO = 101
    gRHEstadosRRHHACTIVO = 201
    gRHEstadosRRHHVACACIONES = 301
    gRHEstadosRRHHOTRASVACACIONES = 302
    gRHEstadosRRHHLICSINGOCE = 401
    gRHEstadosRRHHLICCONGOCE = 402
    gRHEstadosRRHHPERPERSONAL = 403
    gRHEstadosRRHHPERMEDICO = 404
    gRHEstadosRRHHPERCOMISION = 405
    gRHEstadosRRHHSUBSIDIADO = 501
    gRHEstadosRRHHDESCANSO = 502
    gRHEstadosRRHHSUSPENDIDO = 601
    gRHEstadosRRHHRETIRADODESPIDO = 701
    gRHEstadosRRHHRETIRADORENUNCIA = 702
    gRHEstadosRRHHRETIRADONORENOVACIONCONTRATO = 703
    gRHEstadosRRHHPORLIQUIDARDESPIDO = 801
    gRHEstadosRRHHPORLIQUIDARRENUNCIA = 802
    gRHEstadosRRHHPORLIQUIDARNORENOVACIONCONTRATO = 803
End Enum

Global Const gsRHPlanillaSueldos = "E01"
Global Const gsRHPlanillaGratificacion = "E02"
Global Const gsRHPlanillaTercio = "E03"
Global Const gsRHPlanillaUtilidades = "E04"
Global Const gsRHPlanillaCTS = "E05"
Global Const gsRHPlanillaVacaciones = "E06"
Global Const gsRHPlanillaSubsidio = "E07"
Global Const gsRHPlanillaLiquidacion = "E08"
Global Const gsRHPlanillaBonificacionVacacinal = "E09"
Global Const gsRHPlanillaBonoProductividad = "E10"
Global Const gsRHPlanillaBonoAguinaldo = "E11"
Global Const gsRHPlanillaSubsidioEnfermedad = "E12"
Global Const gsRHPlanillaReintegro = "E13"
Global Const gsRHPlanillaDev5ta = "E14"
Global Const gsRHPlanillaMovilidad = "E15"

Global Const gnRHTotalTpo = 9999
Global Const gsRHConceptoUMESTRAB = "U_POR_MES_TRAB"
Global Const gsRHConceptoITOTING = "I_TOT_ING"
Global Const gsRHConceptoITOTINGCOD = "130"
Global Const gsRHConceptoINETOPAGARCOD = "112"
Global Const gsRHConceptoINETOPAGAR = "I_NETO_PAGAR"
Global Const gsRHConceptoDTOTDES = "D_TOT_DESC"
Global Const gsRHConceptoDTOTDESCOD = "215"

Global Const gsRHConceptoVTOTREM = "V_TOT_REM"
Global Const gsRHConceptoITOTCTS = "I_TOT_CTS"
Global Const gsRHConceptoITOTTERCIO = "I_TERCIO"
Global Const gsRHConceptoITOTGRAT = "I_TOTAL_GRATIF"
Global Const gsRHConceptoVNETOPAGAR = "V_NETO_PAGAR"


'Codigos de operacion para el calculo de provisiones y remuneraciones
'Sueldos
Global Const gsRHPlanillaSueldosRemEst = "622001"
Global Const gsRHPlanillaSueldosRemCon = "622002"
'Gratificacion
Global Const gsRHPlanillaGratificacionProvEst = "622101"
Global Const gsRHPlanillaGratificacionProvCon = "622102"
Global Const gsRHPlanillaGratificacionRemEst = "622103"
Global Const gsRHPlanillaGratificacionRemCon = "622104"
'Tercio
Global Const gsRHPlanillaTercioProvEst = "622201"
Global Const gsRHPlanillaTercioProvCon = "622202"
Global Const gsRHPlanillaTercioRemEst = "622203"
Global Const gsRHPlanillaTercioRemCon = "622204"
'Utilidades
Global Const gsRHPlanillaUtilidadesProvEst = "622301"
Global Const gsRHPlanillaUtilidadesProvCon = "622302"
Global Const gsRHPlanillaUtilidadesRemEst = "622303"
Global Const gsRHPlanillaUtilidadesRemCon = "622304"
'CTS
Global Const gsRHPlanillaCTSProvEst = "622401"
Global Const gsRHPlanillaCTSProvCon = "622402"
Global Const gsRHPlanillaCTSRemEst = "622403"
Global Const gsRHPlanillaCTSRemCon = "622404"
'Vacaciones
Global Const gsRHPlanillaVacacionesProvEst = "622501"
Global Const gsRHPlanillaVacacionesProvCon = "622502"
Global Const gsRHPlanillaVacacionesRemEst = "622503"
Global Const gsRHPlanillaVacacionesRemCon = "622504"
'Subsidios
Global Const gsRHPlanillaSubsidioRem = "622601"
Global Const gsRHPlanillaSubsidioEnfermedadRem = "622602"
'Liquidacion
Global Const gsRHPlanillaLiquidacionRem = "622701"
'Bonificacion Vacacional
Global Const gsRHPlanillaBonificacionVacacinalProvEst = "622801"
Global Const gsRHPlanillaBonificacionVacacinalProvCon = "622802"
Global Const gsRHPlanillaBonificacionVacacinalRemEst = "622803"
Global Const gsRHPlanillaBonificacionVacacinalRemCon = "622804"
'Bono Productividad
Global Const gsRHPlanillaBonoProductividadProvEst = "622901"
Global Const gsRHPlanillaBonoProductividadProvCon = "622902"
Global Const gsRHPlanillaBonoProductividadRemEst = "622903"
Global Const gsRHPlanillaBonoProductividadRemCon = "622904"

'Bono Reintegro
Global Const gsRHPlanillaReintegroProvEst = "623001"
Global Const gsRHPlanillaReintegroProvCon = "623002"
Global Const gsRHPlanillaReintegroRemEst = "623003"
Global Const gsRHPlanillaReintegroRemCon = "623004"

'Bono Rev 5ta Categoria
Global Const gsRHPlanillaDev5taRemEst = "623101"
Global Const gsRHPlanillaDev5taRemCon = "623102"

'Bono Movilidad
Global Const gsRHPlanillaMovilidadProvEst = "623201"
Global Const gsRHPlanillaMovilidadProvCon = "623202"
Global Const gsRHPlanillaMovilidadRemEst = "623203"
Global Const gsRHPlanillaMovilidadRemCon = "623204"

'bono gestion
Global Const gsRHPlanillaBonoGestionRemEst = "623303"
Global Const gsRHPlanillaBonoGestionRemCon = "623304"


'Rechazo de Orden de Compra u Orden de Servicio
Global Const gsLogRechazoOCOS = "562501"
'Aprueba de Orden de Compra u Orden de Servicio
Global Const gsLogApruebaOCOS = "562502"

Public Function GetMovNro(psCodUser As String, psCodAge As String, Optional psCorrelativo As String = "00") As String
    Dim oCon As NConstSistemas
    Set oCon = New NConstSistemas
    
    If Len(psCodAge) = 2 Then
        GetMovNro = Format(gdFecSis, "yyyymmdd") & Format(Time, "hhmmss") & oCon.LeeConstSistema(gConstSistCodCMAC) & psCodAge & psCorrelativo & psCodUser
    Else
        GetMovNro = Format(gdFecSis, "yyyymmdd") & Format(Time, "hhmmss") & psCodAge & psCorrelativo & psCodUser
    End If
End Function

Public Sub Main()
    'frmLogin.Show
    MDISicmact.Show
End Sub


