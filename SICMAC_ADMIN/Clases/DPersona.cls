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
'EJVG20110712************************
Private sCelular As String
Private sCelular2 As String
Private sCelular3 As String
'************************************
Private sCiiu As String
Private sCiiuDescripcion As String
Private sEstado As String
Private nTalla As Double
Private nPeso As Double
Private sTipoSangre As String
Private nTipoPersona As Integer 'Personeria
Private sEmail As String
Private sEmail2 As String 'EJVG20111207
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
    nIngTotales As Double
    nSaldDispFam As Double
    nActivo As Double
    nActivoCircul As Double
    nPasivo As Double
    nPasivoPatrimonio As Double
End Type

Private FtesIngreso() As TFteIng
Private FtesIngresoTmp As TFteIng
Private nNumFtes As Integer
Dim ConexDitrib As ADODB.Connection
Dim bPersDistrib As Boolean
Dim nNumFtesDep As Integer
Dim nNumFtesIndep As Integer

Public Function GenerarImpresionFteIngresoDependiente(ByVal pnIndex As Integer, _
    ByVal psNomAge As String, ByVal pdFecha As Date) As String
Dim sCadImp As String
Dim I As Integer
Dim lnNegritaON As String
Dim lnNegritaOFF As String
    
    lnNegritaON = Chr$(27) & Chr$(69)
    lnNegritaOFF = Chr$(27) + Chr$(70)

    sCadImp = ""
    sCadImp = sCadImp & "CAJA MUNICIPAL DE TRUJILLO - " & psNomAge
    sCadImp = sCadImp & Space(25) & Format(FechaHora(pdFecha), "dd/mm/yyyy hh:mm:ss") & Chr$(10) & Chr$(10) & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "Cliente : " & sPersCod & Space(2) & PstaNombre(sNombreCompleto) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & "Fte Ingreso : " & FtesIngreso(pnIndex).sRazSocDescrip & Chr$(10) & Chr$(10) & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & Space(20) & "EVOLUCION ECONOMICA FINANCIERA DEPENDIENTE" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & Space(23)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
        sCadImp = sCadImp & Format(FtesIngreso(pnIndex).FtesIngresoDep(I).dFecEval, "dd/mm/yyyy") & Space(3)
        If I = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            sCadImp = sCadImp & "Variacion"
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("Ingreso Familiar", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(I).nIngFam, 8, 2)
        If I = FtesIngreso(pnIndex).nNumFtesDep - 1 And I > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoDep(I - 1).nIngFam = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 5, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(I).nIngFam / FtesIngreso(pnIndex).FtesIngresoDep(I - 1).nIngFam, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    
    sCadImp = sCadImp & ImpreFormat("Ingreso Conyugue", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(I).nIngCony, 8, 2)
        If I = FtesIngreso(pnIndex).nNumFtesDep - 1 And I > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoDep(I - 1).nIngCony = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 5, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(I).nIngCony / FtesIngreso(pnIndex).FtesIngresoDep(I - 1).nIngCony, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    
    sCadImp = sCadImp & ImpreFormat("Otros Ingresos", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(I).nIngOtros, 8, 2)
        If I = FtesIngreso(pnIndex).nNumFtesDep - 1 And I > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoDep(I - 1).nIngOtros = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 5, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(I).nIngOtros / FtesIngreso(pnIndex).FtesIngresoDep(I - 1).nIngOtros, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    
    sCadImp = sCadImp & ImpreFormat("Gastos Familiares", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(I).nGastosFam, 8, 2)
        If I = FtesIngreso(pnIndex).nNumFtesDep - 1 And I > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoDep(I - 1).nGastosFam = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 5, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(I).nGastosFam / FtesIngreso(pnIndex).FtesIngresoDep(I - 1).nGastosFam, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ImpreFormat("Saldo", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesDep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoDep(I).nIngCony + _
                                FtesIngreso(pnIndex).FtesIngresoDep(I).nIngFam + _
                                FtesIngreso(pnIndex).FtesIngresoDep(I).nIngOtros - FtesIngreso(pnIndex).FtesIngresoDep(I).nGastosFam, 8, 2)
        If I = FtesIngreso(pnIndex).nNumFtesDep - 1 And I > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoDep(I - 1).nIngCony + _
                                FtesIngreso(pnIndex).FtesIngresoDep(I - 1).nIngFam + _
                                FtesIngreso(pnIndex).FtesIngresoDep(I - 1).nIngOtros - FtesIngreso(pnIndex).FtesIngresoDep(I - 1).nGastosFam = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 5, 2)
            Else
                sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).FtesIngresoDep(I).nIngCony + _
                                FtesIngreso(pnIndex).FtesIngresoDep(I).nIngFam + _
                                FtesIngreso(pnIndex).FtesIngresoDep(I).nIngOtros - FtesIngreso(pnIndex).FtesIngresoDep(I).nGastosFam) / _
                                (FtesIngreso(pnIndex).FtesIngresoDep(I - 1).nIngCony + _
                                FtesIngreso(pnIndex).FtesIngresoDep(I - 1).nIngFam + _
                                FtesIngreso(pnIndex).FtesIngresoDep(I - 1).nIngOtros - FtesIngreso(pnIndex).FtesIngresoDep(I - 1).nGastosFam), 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    GenerarImpresionFteIngresoDependiente = sCadImp
End Function

Public Function GenerarImpresionFteIngresoIndependiente(ByVal pnIndex As Integer, _
    ByVal psNomAge As String, ByVal pdFecha As Date) As String
Dim sCadImp As String
Dim oCred As DCreditos
Dim MatDatos As Variant
Dim I As Integer
Dim lnNegritaON As String
Dim lnNegritaOFF As String

    lnNegritaON = Chr$(27) & Chr$(69)
    lnNegritaOFF = Chr$(27) + Chr$(70)

    Set oCred = New DCreditos
    MatDatos = oCred.RecuperaHistoriaCredPersona(sPersCod)
    Set oCred = Nothing
    
    sCadImp = ""
    sCadImp = sCadImp & "CAJA MUNICIPAL DE TRUJILLO - " & psNomAge
    sCadImp = sCadImp & Space(25) & Format(FechaHora(pdFecha), "dd/mm/yyy hh:mm:ss") & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "Cliente : " & sPersCod & Space(2) & PstaNombre(sNombreCompleto) & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & "Fte Ingreso : " & FtesIngreso(pnIndex).sRazSocDescrip & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & Space(20) & "HISTORIA CREDITICA" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & String(80, "-") & Chr$(10)
    If Not IsArray(MatDatos) Then
        sCadImp = sCadImp & "NO POSEE CREDITOS" & Chr$(10)
    Else
        sCadImp = sCadImp & "CREDITO No         : "
        For I = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(I, 0) & Space(2)
        Next I
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Estado Actual      : "
        For I = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(I, 1) & Space(2)
        Next I
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Fecha Vigencia     : "
        For I = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(I, 2) & Space(2)
        Next I
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Fecha Cancelacion  : "
        For I = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(I, 3) & Space(2)
        Next I
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Monto(Sol/Sug/Apr) : "
        For I = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(I, 4) & Space(2)
        Next I
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Cuotas/Plazo(Dias) : "
        For I = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(I, 5) & Space(2)
        Next I
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Cuota (Sol/Sug/Apr): "
        For I = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(I, 6) & Space(2)
        Next I
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Nota               : "
        For I = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(I, 7) & Space(2)
        Next I
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "Analista           : "
        For I = 0 To UBound(MatDatos) - 1
            sCadImp = sCadImp & MatDatos(I, 8) & Space(2)
        Next I
        sCadImp = sCadImp & Chr$(10)
    End If
    sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & Space(20) & "EVOLUCION ECONOMICA FINANCIERA" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & Space(25)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & Format(FtesIngreso(pnIndex).FtesIngresoInd(I).dFecEval, "dd/mm/yyyy") & Space(3)
    Next I
    sCadImp = sCadImp & "Variacion" & Chr$(10)
    sCadImp = sCadImp & String(80, "-") & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ImpreFormat("ACTIVO CIRCULANTE", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nInventarios, 8, 2)
        If I = FtesIngreso(pnIndex).nNumFtesIndep - 1 And I > 0 Then
            If (FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nInventarios) = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(I).nActivDisp + _
                            FtesIngreso(pnIndex).FtesIngresoInd(I).nCtasxCob + _
                            FtesIngreso(pnIndex).FtesIngresoInd(I).nInventarios) / (FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nActivDisp + _
                            FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCtasxCob + _
                            FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nInventarios), 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ImpreFormat("Disponible", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nActivDisp, 8, 2)
        If I = FtesIngreso(pnIndex).nNumFtesIndep - 1 And I > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nActivDisp = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nActivDisp / FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nActivDisp, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("Cuentas x Cobrar", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nCtasxCob, 8, 2)
        If I = FtesIngreso(pnIndex).nNumFtesIndep - 1 And I > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCtasxCob = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nCtasxCob / FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCtasxCob, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("Inventario", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nInventarios, 8, 2)
        If I = FtesIngreso(pnIndex).nNumFtesIndep - 1 And I > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nInventarios = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nInventarios / FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nInventarios, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ImpreFormat("ACTIVO FIJO", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nActFijo, 8, 2)
        If I = FtesIngreso(pnIndex).nNumFtesIndep - 1 And I > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nActFijo = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nActFijo / FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nActFijo, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("TOTAL ACTIVO", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(I).nActFijo, 8, 2)
        If I = FtesIngreso(pnIndex).nNumFtesIndep - 1 And I > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nActFijo = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(I).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(I).nActFijo) / (FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nActFijo), 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("PASIVO", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nProveedores + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nCreditosOtros + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nCreditosCmact, 8, 2)
        If I = FtesIngreso(pnIndex).nNumFtesIndep - 1 And I > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nProveedores + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCreditosOtros + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCreditosCmact = 0 Then
                        
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(I).nProveedores + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nCreditosOtros + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nCreditosCmact) / (FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nProveedores + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCreditosOtros + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCreditosCmact), 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ImpreFormat("Proveedores", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nProveedores, 8, 2)
        If I = FtesIngreso(pnIndex).nNumFtesIndep - 1 And I > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nProveedores = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nProveedores / FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nProveedores, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("Otros Prestamos", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nCreditosOtros, 8, 2)
        If I = FtesIngreso(pnIndex).nNumFtesIndep - 1 And I > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCreditosOtros = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nCreditosOtros / FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCreditosOtros, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("Prestamos CMACT", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nCreditosCmact, 8, 2)
        If I = FtesIngreso(pnIndex).nNumFtesIndep - 1 And I > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCreditosCmact = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nCreditosCmact / FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCreditosCmact, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ImpreFormat("PATRIMONIO", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nCreditosCmact, 8, 2)
        If I = FtesIngreso(pnIndex).nNumFtesIndep - 1 And I > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCreditosCmact = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nCreditosCmact / FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCreditosCmact, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("INGRESOS", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nPersIngFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nVentas + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nRecupCtasxCob, 8, 2)
                        
        If I = FtesIngreso(pnIndex).nNumFtesIndep - 1 And I > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPersIngFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nVentas + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nRecupCtasxCob = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(I).nPersIngFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nVentas + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nRecupCtasxCob) / (FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPersIngFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nVentas + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nRecupCtasxCob), 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ImpreFormat("Ingresos familiares", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nPersIngFam, 8, 2)
        If I = FtesIngreso(pnIndex).nNumFtesIndep - 1 And I > 0 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPersIngFam = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nPersIngFam / FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPersIngFam, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("Ventas", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nVentas, 8, 2)
        If I > 0 And FtesIngreso(pnIndex).nNumFtesIndep - 1 = I Then
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nVentas = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nVentas / FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nVentas, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("Recup. Ctas x Cob", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nRecupCtasxCob, 8, 2)
        If I > 0 And I = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nRecupCtasxCob = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nRecupCtasxCob / FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nRecupCtasxCob, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("EGRESOS", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nPersEgrFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nOtrosEgresos, 8, 2)
        If I > 0 And I = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPersEgrFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nOtrosEgresos = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
            
                sCadImp = sCadImp & ImpreFormat((FtesIngreso(pnIndex).FtesIngresoInd(I).nPersEgrFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nOtrosEgresos) / (FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPersEgrFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCompraMercad + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nOtrosEgresos), 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("Gastos Familiares", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nPersEgrFam, 8, 2)
        If I > 0 And I = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPersEgrFam = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nPersEgrFam / FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPersEgrFam, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("Compra Mercaderia", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nCompraMercad, 8, 2)
        If I > 0 And I = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCompraMercad = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCompraMercad, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & ImpreFormat("Otros Egresos", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nOtrosEgresos, 8, 2)
        If I > 0 And I = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nOtrosEgresos = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nOtrosEgresos / FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nOtrosEgresos, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & ImpreFormat("DISPONIBLE", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        sCadImp = sCadImp & ImpreFormat(FtesIngreso(pnIndex).FtesIngresoInd(I).nPersIngFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nVentas + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nRecupCtasxCob - _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nPersEgrFam - _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nCompraMercad - _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nOtrosEgresos, 8, 2)
        If I > 0 And I = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If (FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPersIngFam + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nVentas + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nRecupCtasxCob) - _
                        (FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPersEgrFam + _
                         FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCompraMercad + _
                         FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nOtrosEgresos) = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
                
            Else
                sCadImp = sCadImp & ImpreFormat(((FtesIngreso(pnIndex).FtesIngresoInd(I).nPersIngFam + _
                            FtesIngreso(pnIndex).FtesIngresoInd(I).nVentas + _
                            FtesIngreso(pnIndex).FtesIngresoInd(I).nRecupCtasxCob) - _
                            (FtesIngreso(pnIndex).FtesIngresoInd(I).nPersEgrFam + _
                             FtesIngreso(pnIndex).FtesIngresoInd(I).nCompraMercad + _
                             FtesIngreso(pnIndex).FtesIngresoInd(I).nOtrosEgresos)) _
                            / ((FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPersIngFam + _
                            FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nVentas + _
                            FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nRecupCtasxCob) - _
                            (FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPersEgrFam + _
                             FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCompraMercad + _
                             FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nOtrosEgresos)), 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & String(80, "-") & Chr$(10)
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
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nActCir = FtesIngreso(pnIndex).FtesIngresoInd(I).nActivDisp + _
                            FtesIngreso(pnIndex).FtesIngresoInd(I).nCtasxCob + _
                            FtesIngreso(pnIndex).FtesIngresoInd(I).nInventarios
        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(I).nProveedores + _
                                FtesIngreso(pnIndex).FtesIngresoInd(I).nCreditosOtros + _
                                FtesIngreso(pnIndex).FtesIngresoInd(I).nCreditosCmact
        If nPasivo = 0 Then
            Res1 = 0
        Else
            Res1 = nActCir / nPasivo
        End If
        sCadImp = sCadImp & ImpreFormat(Res1, 8, 2)
        If I > 0 And I = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nActCirAnt = FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nActivDisp + _
                                FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCtasxCob + _
                                FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nInventarios
            nPasivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nProveedores + _
                                FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCreditosOtros + _
                                FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCreditosCmact
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
    Next I
    
    sCadImp = sCadImp & Chr$(10)
    'Activo Circulante - Pasivo
    sCadImp = sCadImp & ImpreFormat("Capital Trabajo", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nActCir = FtesIngreso(pnIndex).FtesIngresoInd(I).nActivDisp + _
                            FtesIngreso(pnIndex).FtesIngresoInd(I).nCtasxCob + _
                            FtesIngreso(pnIndex).FtesIngresoInd(I).nInventarios
        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(I).nProveedores + _
                                FtesIngreso(pnIndex).FtesIngresoInd(I).nCreditosOtros + _
                                FtesIngreso(pnIndex).FtesIngresoInd(I).nCreditosCmact
        Res1 = nActCir - nPasivo
        sCadImp = sCadImp & ImpreFormat(Res1, 8, 2)
        If I > 0 And I = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nActCirAnt = FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nActivDisp + _
                                FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCtasxCob + _
                                FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nInventarios
            nPasivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nProveedores + _
                                FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCreditosOtros + _
                                FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCreditosCmact
            Res2 = nActCirAnt - nPasivoAnt
            If Res2 = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "GESTION" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    'Costo de Ventas / Inventario
    sCadImp = sCadImp & ImpreFormat("Rotacion Inventario", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        If FtesIngreso(pnIndex).FtesIngresoInd(I).nInventarios = 0 Then
            Res1 = 0
        Else
            Res1 = FtesIngreso(pnIndex).FtesIngresoInd(I).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(I).nInventarios
        End If
        sCadImp = sCadImp & ImpreFormat(Res1, 8, 2)
        
        If I > 0 And I = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nInventarios = 0 Then
                Res2 = 0
            Else
                Res2 = FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCompraMercad / FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nInventarios
            End If
            If Res2 = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    
    'Recup Ctas x Cob / (Ventas / 30)
    sCadImp = sCadImp & ImpreFormat("Periodo de Cobro", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        If FtesIngreso(pnIndex).FtesIngresoInd(I).nVentas = 0 Then
            Res1 = 0
        Else
            Res1 = FtesIngreso(pnIndex).FtesIngresoInd(I).nRecupCtasxCob / (FtesIngreso(pnIndex).FtesIngresoInd(I).nVentas / 30)
        End If
        sCadImp = sCadImp & ImpreFormat(Res1, 8, 2)
        If I > 0 And I = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nVentas = 0 Then
                Res2 = 0
            Else
                Res2 = FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nRecupCtasxCob / (FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nVentas / 30)
            End If
            If Res2 = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next I
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
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nIngresos = FtesIngreso(pnIndex).FtesIngresoInd(I).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(I).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(I).nRecupCtasxCob
        nEgresos = FtesIngreso(pnIndex).FtesIngresoInd(I).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(I).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(I).nOtrosEgresos
        
        If FtesIngreso(pnIndex).FtesIngresoInd(I).nPasivoPatrimonio = 0 Then
            Res1 = 0
        Else
            Res1 = (nIngresos - nEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(I).nPasivoPatrimonio
        End If
        sCadImp = sCadImp & ImpreFormat(Res1, 8, 2)
        If I > 0 And I = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nIngresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nRecupCtasxCob
            nEgresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nOtrosEgresos
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPasivoPatrimonio = 0 Then
                Res2 = 0
            Else
                Res2 = (nIngresos - nEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPasivoPatrimonio
            End If
        
            If Res2 = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    
    Dim nTotActivo As Double
    Dim nTotActivoAnt As Double
    sCadImp = sCadImp & ImpreFormat("Rentab. Activo", 20)
    'Utilidad Operativa / Total Activo
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nIngresos = FtesIngreso(pnIndex).FtesIngresoInd(I).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(I).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(I).nRecupCtasxCob
        nEgresos = FtesIngreso(pnIndex).FtesIngresoInd(I).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(I).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(I).nOtrosEgresos
        nTotActivo = FtesIngreso(pnIndex).FtesIngresoInd(I).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(I).nActFijo
        If nTotActivo = 0 Then
            Res1 = 0
        Else
            Res1 = (nIngresos - nEgresos) / nTotActivo
        End If
        sCadImp = sCadImp & ImpreFormat(Res1, 8, 2)
        If I > 0 And I = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nIngresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(I).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nRecupCtasxCob
            nEgresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(I).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nOtrosEgresos
            nTotActivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nActivDisp + _
                            FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCtasxCob + _
                            FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nActFijo
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
    Next I
    sCadImp = sCadImp & Chr$(10)
    
    sCadImp = sCadImp & ImpreFormat("Rentab. Ventas", 20)
    'Utilidad Operativa / Ventas
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nIngresos = FtesIngreso(pnIndex).FtesIngresoInd(I).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(I).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(I).nRecupCtasxCob
        nEgresos = FtesIngreso(pnIndex).FtesIngresoInd(I).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(I).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(I).nOtrosEgresos
        If FtesIngreso(pnIndex).FtesIngresoInd(I).nVentas = 0 Then
            Res1 = 0
        Else
            Res1 = (nIngresos - nEgresos) / FtesIngreso(pnIndex).FtesIngresoInd(I).nVentas
        End If
        sCadImp = sCadImp & ImpreFormat(Res1, 8, 2)
        If I > 0 And I = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nIngresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPersIngFam + FtesIngreso(pnIndex).FtesIngresoInd(I).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nRecupCtasxCob
            nEgresosAnt = FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPersEgrFam + FtesIngreso(pnIndex).FtesIngresoInd(I).nCompraMercad + FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nOtrosEgresos
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nVentas = 0 Then
                Res2 = 0
            Else
                Res2 = (nIngresosAnt - nEgresosAnt) / FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nVentas
            End If
            If Res2 = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    
    'Ventas Totales = Ventas + Recup Ctas x Cob
    'Margen Utilidad = Ventas Totales - Costo Ventas / Ventas Totales
    Dim nVentasTot As Double
    Dim nVentasTotAnt As Double
    sCadImp = sCadImp & ImpreFormat("Margen Utilidad", 20)
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nVentasTot = FtesIngreso(pnIndex).FtesIngresoInd(I).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(I).nRecupCtasxCob
        If nVentasTot = 0 Then
            Res1 = 0
        Else
            Res1 = (nVentasTot - FtesIngreso(pnIndex).FtesIngresoInd(I).nCompraMercad) / nVentasTot
        End If
        sCadImp = sCadImp & ImpreFormat(Res1, 8, 2)
        If I > 0 And I = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nVentasTotAnt = FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nVentas + FtesIngreso(pnIndex).FtesIngresoInd(I).nRecupCtasxCob
            If nVentasTotAnt = 0 Then
                Res2 = 0
            Else
                Res2 = (nVentasTotAnt - FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCompraMercad) / nVentasTotAnt
            End If
            If Res2 = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    
    sCadImp = sCadImp & lnNegritaON
    sCadImp = sCadImp & "ENDEUDAMIENTO" & Chr$(10)
    sCadImp = sCadImp & lnNegritaOFF
    sCadImp = sCadImp & ImpreFormat("Endeudamiento Pat.", 20)
    'Endeudamiento = Pasivo / Patrimonio
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(I).nProveedores + _
                                FtesIngreso(pnIndex).FtesIngresoInd(I).nCreditosOtros + _
                                FtesIngreso(pnIndex).FtesIngresoInd(I).nCreditosCmact
                                
        If FtesIngreso(pnIndex).FtesIngresoInd(I).nPasivoPatrimonio = 0 Then
            Res1 = 0
        Else
            Res1 = nPasivo / FtesIngreso(pnIndex).FtesIngresoInd(I).nPasivoPatrimonio
        End If
        sCadImp = sCadImp & ImpreFormat(Res1, 8, 2)
        If I > 0 And I = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nPasivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nProveedores + _
                                FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCreditosOtros + _
                                FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCreditosCmact
                                
            If FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPasivoPatrimonio = 0 Then
                Res2 = 0
            Else
                Res2 = nPasivoAnt / FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nPasivoPatrimonio
            End If
            If Res2 = 0 Then
                sCadImp = sCadImp & ImpreFormat(0, 8, 2)
            Else
                sCadImp = sCadImp & ImpreFormat(Res1 / Res2, 8, 2)
            End If
        End If
    Next I
    sCadImp = sCadImp & Chr$(10)
    
    sCadImp = sCadImp & ImpreFormat("Endeudamiento", 20)
    'Endeudamiento = Pasivo / TotalActivo
    For I = 0 To FtesIngreso(pnIndex).nNumFtesIndep - 1
        nTotActivo = FtesIngreso(pnIndex).FtesIngresoInd(I).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(I).nActFijo
        nPasivo = FtesIngreso(pnIndex).FtesIngresoInd(I).nProveedores + _
                                FtesIngreso(pnIndex).FtesIngresoInd(I).nCreditosOtros + _
                                FtesIngreso(pnIndex).FtesIngresoInd(I).nCreditosCmact
                                
        If nTotActivo = 0 Then
            Res1 = 0
        Else
            Res1 = nPasivo / nTotActivo
        End If
        sCadImp = sCadImp & ImpreFormat(Res1, 8, 2)
        If I > 0 And I = FtesIngreso(pnIndex).nNumFtesIndep - 1 Then
            nTotActivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nActivDisp + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCtasxCob + _
                        FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nInventarios + FtesIngreso(pnIndex).FtesIngresoInd(I).nActFijo
            nPasivoAnt = FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nProveedores + _
                                FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCreditosOtros + _
                                FtesIngreso(pnIndex).FtesIngresoInd(I - 1).nCreditosCmact
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
    Next I
    sCadImp = sCadImp & Chr$(10)
    
    GenerarImpresionFteIngresoIndependiente = sCadImp
End Function


Public Sub LimpiaEliminados()
Dim I, j As Integer
Dim bExisteEliminados As Boolean
    bExisteEliminados = True
    Do While bExisteEliminados
        bExisteEliminados = False
        For I = 0 To nNumFtes - 1
            If FtesIngreso(I).nTipoAct = 3 Then
                bExisteEliminados = True
                For j = I To nNumFtes - 2
                    FtesIngreso(j) = FtesIngreso(j + 1)
                Next j
                nNumFtes = nNumFtes - 1
                Exit For
            End If
        Next I
    Loop
End Sub

Public Function UbicarFteIngreso(psCodFI As String) As Integer
Dim I As Integer
    
    UbicarFteIngreso = -1
    For I = 0 To nNumFtes - 1
        If FtesIngreso(I).sFteIngresoOrig = psCodFI Then
            UbicarFteIngreso = I
        End If
    Next I
    
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
Dim I As Integer
    ExistenCambios = False
    If nTipoAct <> PersFilaSinCambios Then
        ExistenCambios = True
        Exit Function
    End If
    For I = 0 To nNumDoc - 1
        If Docs(I).nTipoAct <> PersFilaSinCambios Then
            ExistenCambios = True
            Exit Function
        End If
    Next I
    For I = 0 To nNumRel - 1
        If Relac(I).nTipoAct <> PersFilaSinCambios Then
            ExistenCambios = True
            Exit Function
        End If
    Next I
    For I = 0 To nNumFtes - 1
        If FtesIngreso(I).nTipoAct <> PersFilaSinCambios Then
            ExistenCambios = True
            Exit Function
        End If
    Next I
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
    'EJVG20111212 ********************
    sCelular = ""
    sCelular2 = ""
    sCelular3 = ""
    
    sCiiu = ""
    sEstado = ""
    nTalla = 0
    nPeso = 0
    sTipoSangre = ""
    nTipoPersona = gPersonaNat
    sEmail = ""
    sEmail2 = "" 'EJVG20111212
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
Dim I As Integer
Dim j As Integer
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
    For I = 0 To UBound(Docs) - 1
        Select Case Trim(Str(Docs(I).nTipo))
            Case "1", "3"
                vDocs(0) = Trim(Str(Docs(I).nTipo))
                vDocs(1) = Docs(I).sNumero
                Exit For
            Case "4"
                vDocs(0) = "2"
                vDocs(1) = Docs(I).sNumero
                Exit For
            Case "5", "6", "7", "8", "9"
                vDocs(0) = "9"
                vDocs(1) = Docs(I).sNumero
                Exit For
            Case "10"
                vDocs(0) = "6"
                vDocs(1) = Docs(I).sNumero
                Exit For
            Case "11"
                vDocs(0) = "5"
                vDocs(1) = Docs(I).sNumero
                Exit For
            Case "12"
                vDocs(0) = "7"
                vDocs(1) = Docs(I).sNumero
                Exit For
            Case "13"
                vDocs(0) = "8"
                vDocs(1) = Docs(I).sNumero
                Exit For
            Case "13"
                vDocs(0) = "8"
                vDocs(1) = Docs(I).sNumero
                Exit For
        End Select
    Next I
    
    For I = 0 To UBound(Docs) - 1
        Select Case Trim(Str(Docs(I).nTipo))
            Case "2"
                vDocs(2) = "2"
                vDocs(3) = Docs(I).sNumero
                Exit For
            Case "14"
                vDocs(2) = "1"
                vDocs(3) = Docs(I).sNumero
                Exit For
            Case "15"
                vDocs(2) = "3"
                vDocs(3) = Docs(I).sNumero
                Exit For
        End Select
    Next I
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
        sSql = "UPDATE " & gConsPersona & "Persona SET "
        sSql = sSql & "cPersNombre = '" & sNombreCompleto & "',"
        sSql = sSql & "dPersNacCreac = '" & Format(dFechaNac, "mm/dd/yyyy") & "',"
        sSql = sSql & "cPersDireccUbiGeo = '" & sUbicGeografica & "',"
        sSql = sSql & "cPersDireccDomicilio = '" & sDomicilio & "',"
        sSql = sSql & "cPersDireccCondicion = '" & Trim(sCondicionDomic) & "',"
        sSql = sSql & "nPersValComDomicilio = " & nValComDomicilio & ","
        sSql = sSql & "cPersTelefono = '" & Trim(sTelefonos) & "',"
        sSql = sSql & "cPersTelefono2 = '" & Trim(sTelefonos2) & "',"
        'EJVG20111212 *******************************
        sSql = sSql & "cPersCelular = '" & Trim(sCelular) & "',"
        sSql = sSql & "cPersCelular2 = '" & Trim(sCelular2) & "',"
        sSql = sSql & "cPersCelular3 = '" & Trim(sCelular3) & "',"
        
        sSql = sSql & "cPersEmail = '" & Trim(sEmail) & "',"
        sSql = sSql & "cPersEmail2 = '" & Trim(sEmail2) & "'," 'EJVG20111212
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
            'sSql = sSql & ",cPersDireccCondicion, nPersValComDomicilio, cPersTelefono,cPersTelefono2,nPersPersoneria,cPersCIIU,cPersEstado"
            sSql = sSql & ",cPersDireccCondicion, nPersValComDomicilio, cPersTelefono,cPersTelefono2,cPersCelular,cPersCelular2,cPersCelular3,nPersPersoneria,cPersCIIU,cPersEstado" 'EJVG20111209
            'sSql = sSql & ",cPersEmail,cUltimaActualizacion, cPersCodSbs, nPersRelaInst,dPersIng)"
            sSql = sSql & ",cPersEmail,cPersEmail2,cUltimaActualizacion, cPersCodSbs, nPersRelaInst,dPersIng)" 'EJVG20111209
            sSql = sSql & "VALUES('" & sPersCod & "',"
            sSql = sSql & "'" & sNombreCompleto & "',"
            sSql = sSql & "'" & Format(dFechaNac, "mm/dd/yyyy") & "',"
            sSql = sSql & "'" & sUbicGeografica & "',"
            sSql = sSql & "'" & sDomicilio & "',"
            sSql = sSql & "'" & sCondicionDomic & "',"
            sSql = sSql & nValComDomicilio & ","
            sSql = sSql & "'" & sTelefonos & "',"
            sSql = sSql & "'" & sTelefonos2 & "',"
            'EJVG20111209 *************************
            sSql = sSql & "'" & sCelular & "',"
            sSql = sSql & "'" & sCelular2 & "',"
            sSql = sSql & "'" & sCelular3 & "',"
            
            sSql = sSql & Trim(Str(nTipoPersona)) & ","
            sSql = sSql & "'" & sCiiu & "',"
            sSql = sSql & "'" & sEstado & "',"
            sSql = sSql & "'" & sEmail & "',"
            sSql = sSql & "'" & sEmail2 & "'," 'EJVG20111209
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
    For I = 0 To nNumRel - 1
        If Relac(I).nTipoAct = PersFilaModificada Then
            sSql = "UPDATE " & gConsPersona & "PersRelaciones SET "
            sSql = sSql & "cPersRelacPersCod = '" & Relac(I).sPersCod & "',"
            sSql = sSql & "nPersRelac = " & Trim(Right(Relac(I).sRelacion, 20)) & ","
            sSql = sSql & "bPersRelacBenef = " & Trim(Right(Relac(I).sPersRelacBenef, 20)) & ","
            sSql = sSql & "nPersRelacBenefPorc = " & Format(Relac(I).nPersRelacBenefPorc, "#0.00") & ","
            sSql = sSql & "bPersRelacAMP = " & Trim(Right(Relac(I).sPersRelacAMP, 20))
            'sSql = sSql & "PR.bPersRelacEstado = " & Trim(Right(Relac(i).sPersRelacEstado, 20))
            sSql = sSql & " Where cPersCod = '" & sPersCod & "' And cPersRelacPersCod = '" & Relac(I).sPersCodOrig & "'"
            ConexionActiva.Execute sSql
        Else
            If Relac(I).nTipoAct = PersFilaNueva Then
                sSql = "INSERT INTO " & gConsPersona & "PERSRELACIONES(cPersCod,cPersRelacPersCod,"
                sSql = sSql & "nPersRelac,bPersrelacBenef,nPersRelacBenefPorc,bPersRelacAMP,cMovNro)"
                sSql = sSql & "VALUES('" & sPersCod & "',"
                sSql = sSql & "'" & Relac(I).sPersCod & "',"
                sSql = sSql & Trim(Right(Relac(I).sRelacion, 20)) & ","
                sSql = sSql & Trim(Right(Relac(I).sPersRelacBenef, 20)) & ","
                sSql = sSql & Format(Relac(I).nPersRelacBenefPorc, "#0.00") & ","
                sSql = sSql & Trim(Right(Relac(I).sPersRelacAMP, 20)) & ","
                sSql = sSql & "'" & sActualiza & "')"
                ConexionActiva.Execute sSql
            Else
                If Relac(I).nTipoAct = PersFilaEliminda Then
                    sSql = "DELETE " & gConsPersona & "PersRelaciones Where "
                    sSql = sSql & " cPersCod = '" & sPersCod & "'"
                    sSql = sSql & " And cPersRelacPersCod = '" & Trim(Relac(I).sPersCodOrig) & "'"
                    ConexionActiva.Execute sSql
                End If
            End If
        End If
    Next I
    
    'Graba Referencias Comerciales
    For I = 0 To nNumRefCom - 1
        If RefCom(I).nTipoAct = PersFilaModificada Then
            sSql = "UPDATE PERSREFCOMERCIAL SET "
            sSql = sSql & "cNomRefCom= '" & RefCom(I).sRefComNom & "',"
            sSql = sSql & "nTipoRelaCom = " & Right(RefCom(I).sRefComTipo, 1) & ","
            sSql = sSql & "cFonoRefCom = '" & RefCom(I).sRefComFono & "',"
            sSql = sSql & "cUltimaActualizacion = '" & sActualiza & "' "
            sSql = sSql & " WHERE cPersCod = '" & sPersCod & "' AND nNumRefCom = " & RefCom(I).nRefComCod
            ConexionActiva.Execute sSql
        Else
            If RefCom(I).nTipoAct = PersFilaNueva Then
                sSql = "INSERT INTO PERSREFCOMERCIAL (cPersCod, nNumRefCom, cNomRefCom, nTipoRelaCom, "
                sSql = sSql & "cFonoRefCom, cUltimaActualizacion)"
                sSql = sSql & " VALUES ('" & sPersCod & "',"
                sSql = sSql & RefCom(I).nRefComCod & ","
                sSql = sSql & "'" & RefCom(I).sRefComNom & "',"
                sSql = sSql & Right(RefCom(I).sRefComTipo, 1) & ","
                sSql = sSql & "'" & RefCom(I).sRefComFono & "',"
                sSql = sSql & "'" & sActualiza & "')"
                ConexionActiva.Execute sSql
            Else
                If RefCom(I).nTipoAct = PersFilaEliminda Then
                    sSql = "DELETE PERSREFCOMERCIAL "
                    sSql = sSql & " WHERE cPersCod = '" & sPersCod & "' AND nNumRefCom = " & RefCom(I).nRefComCod
                    ConexionActiva.Execute sSql
                End If
            End If
        End If
    Next I
    
    'Graba Referencias Bancarias
    For I = 0 To nNumRefBan - 1
        If RefBan(I).nTipoAct = PersFilaModificada Then
            sSql = "UPDATE PERSREFBANCARIA SET "
            sSql = sSql & "cCodBanco = '" & Mid(RefBan(I).sRefBanCodIF, 4, Len(RefBan(I).sRefBanCodIF) - 3) & "',"
            sSql = sSql & "cNumCta = '" & RefBan(I).sRefBanNumCta & "',"
            sSql = sSql & "cNumTarjeta = '" & RefBan(I).sRefBanNumTar & "',"
            sSql = sSql & "nMontoLinCred = " & RefBan(I).nRefBanLinCred & ","
            sSql = sSql & "cUltimaActualizacion = '" & sActualiza & "' "
            sSql = sSql & " WHERE cPersCod = '" & sPersCod & "' AND cCodBanco = '" & Mid(RefBan(I).sRefBanCodIFOrg, 4, Len(RefBan(I).sRefBanCodIFOrg) - 3) & "'"
            ConexionActiva.Execute sSql
        Else
            If RefBan(I).nTipoAct = PersFilaNueva Then
                sSql = "INSERT INTO PERSREFBANCARIA (cPersCod, cCodBanco, cNumCta, cNumTarjeta, "
                sSql = sSql & "nMontoLinCred, cUltimaActualizacion)"
                sSql = sSql & " VALUES ('" & sPersCod & "',"
                sSql = sSql & "'" & Mid(RefBan(I).sRefBanCodIF, 4, Len(RefBan(I).sRefBanCodIF) - 3) & "',"
                sSql = sSql & "'" & RefBan(I).sRefBanNumCta & "',"
                sSql = sSql & "'" & RefBan(I).sRefBanNumTar & "',"
                sSql = sSql & RefBan(I).nRefBanLinCred & ","
                sSql = sSql & "'" & sActualiza & "')"
                ConexionActiva.Execute sSql
            Else
                If RefBan(I).nTipoAct = PersFilaEliminda Then
                    sSql = "DELETE PERSREFBANCARIA "
                    sSql = sSql & " WHERE cPersCod = '" & sPersCod & "' AND cCodBanco = '" & Mid(RefBan(I).sRefBanCodIFOrg, 4, Len(RefBan(I).sRefBanCodIFOrg) - 3) & "'"
                    ConexionActiva.Execute sSql
                End If
            End If
        End If
    Next I
    
    'Graba Patrimonio Vehicular
    For I = 0 To nNumPatVeh - 1
        If PatVeh(I).nTipoAct = PersFilaModificada Then
            sSql = "UPDATE PERSPATRIMONIO SET "
            sSql = sSql & "cMarca= '" & PatVeh(I).sPatVehMarca & "',"
            sSql = sSql & "cFecFabri = '" & PatVeh(I).sPatVehFecFab & "',"
            sSql = sSql & "nValorCom = " & PatVeh(I).nPatVehValCom & ","
            sSql = sSql & "nCondPatri = " & Right(PatVeh(I).sPatVehCondicion, 1) & ","
            sSql = sSql & "cUltimaActualizacion = '" & sActualiza & "' "
            sSql = sSql & " WHERE cPersCod = '" & sPersCod & "' AND nNumPatri= " & PatVeh(I).nPatVehCod
            ConexionActiva.Execute sSql
        Else
            If PatVeh(I).nTipoAct = PersFilaNueva Then
                sSql = "INSERT INTO PERSPATRIMONIO(cPersCod, nNumPatri, nCondPatri, cMarca, cFecFabrica, "
                sSql = sSql & "nValorCom, cUltimaActualizacion)"
                sSql = sSql & " VALUES ('" & sPersCod & "',"
                sSql = sSql & PatVeh(I).nPatVehCod & ","
                sSql = sSql & Right(PatVeh(I).sPatVehCondicion, 1) & ","
                sSql = sSql & "'" & PatVeh(I).sPatVehMarca & "',"
                sSql = sSql & "'" & PatVeh(I).sPatVehFecFab & "',"
                sSql = sSql & PatVeh(I).nPatVehValCom & ","
                sSql = sSql & "'" & sActualiza & "')"
                ConexionActiva.Execute sSql
            Else
                If PatVeh(I).nTipoAct = PersFilaEliminda Then
                    sSql = "DELETE PERSPATRIMONIO "
                    sSql = sSql & " WHERE cPersCod = '" & sPersCod & "' AND nNumPatri= " & PatVeh(I).nPatVehCod
                    ConexionActiva.Execute sSql
                End If
            End If
        End If
    Next I
    
    'Actualizar Los Documentos de las Personas
    For I = 0 To nNumDoc - 1
        If Docs(I).nTipoAct = PersFilaModificada Then
            sSql = "UPDATE " & gConsPersona & "PERSID SET "
            sSql = sSql & "cPersIdTpo = '" & Trim(Str(Docs(I).nTipo)) & "',"
            sSql = sSql & "cPersIDNro = '" & Docs(I).sNumero & "',"
            sSql = sSql & "cUltimaActualizacion = '" & sActualiza & "' "
            sSql = sSql & " WHERE cPersCod = '" & sPersCod & "' And cPersIDTpo = '" & Trim(Str(Docs(I).nTipoOrig)) & "'"
            ConexionActiva.Execute sSql
        Else
            If Docs(I).nTipoAct = PersFilaNueva Then
                sSql = "INSERT INTO " & gConsPersona & "PERSID("
                sSql = sSql & "cPersCod,cPersIDTpo,cPersIDNro,cUltimaActualizacion)"
                sSql = sSql & "VALUES('" & sPersCod & "',"
                sSql = sSql & "'" & Trim(Str(Docs(I).nTipo)) & "',"
                sSql = sSql & "'" & Docs(I).sNumero & "',"
                sSql = sSql & "'" & sActualiza & "')"
                ConexionActiva.Execute sSql
            Else
                If Docs(I).nTipoAct = PersFilaEliminda Then
                    sSql = "DELETE " & gConsPersona & "PERSID Where "
                    sSql = sSql & " cPersCod = '" & sPersCod & "'"
                    sSql = sSql & " And cPersIDTpo = '" & Trim(Str(Docs(I).nTipoOrig)) & "'"
                    ConexionActiva.Execute sSql
                End If
            End If
        End If
    Next I
    
    'Actualizar las Fuentes de Ingreso
    For I = 0 To nNumFtes - 1
        FtesIngreso(I).sRazSocDescrip = Replace(FtesIngreso(I).sRazSocDescrip, "'", "''")
        If FtesIngreso(I).nTipoAct = PersFilaNueva Then
            FtesIngreso(I).cNumFte = sNumFuente
            sSql = "INSERT INTO " & gConsPersona & "PersFteIngreso(cNumFuente, cPersCod, cPersFIPersCod, cPersFICargo,"
            sSql = sSql & "dPersFIinicio, cPersFIMoneda, cPersFIComentario, cRazSocUbiGeo, cRazSocDirecc, "
            sSql = sSql & " cRazSocTelef, cRazSocDescrip, nPersTipFte, cUltimaActualizacion) "
            sSql = sSql & " VALUES('" & FtesIngreso(I).cNumFte & "','" & sPersCod & "',"
            sSql = sSql & "'" & FtesIngreso(I).sFteIngreso & "',"
            sSql = sSql & "'" & FtesIngreso(I).sCargo & "',"
            sSql = sSql & "'" & Format(FtesIngreso(I).dFteIngreso, "mm/dd/yyyy") & "',"
            sSql = sSql & Trim(FtesIngreso(I).sMoneda) & ","
            sSql = sSql & "'" & Trim(FtesIngreso(I).sComentarios) & "',"
            sSql = sSql & "'" & Right(FtesIngreso(I).sRazSocUbiGeo, 12) & "',"
            sSql = sSql & "'" & FtesIngreso(I).sRazSocDirecc & "',"
            sSql = sSql & "'" & FtesIngreso(I).sRazSocTelef & "',"
            sSql = sSql & "'" & FtesIngreso(I).sRazSocDescrip & "',"
            sSql = sSql & FtesIngreso(I).nTipo & ",'" & sActualiza
            sSql = sSql & "')"
            ConexionActiva.Execute sSql
        Else
            If FtesIngreso(I).nTipoAct = PersFilaModificada Then
                sSql = "UPDATE " & gConsPersona & "PersFteIngreso SET "
                sSql = sSql & " cPersFIPersCod = '" & FtesIngreso(I).sFteIngreso & "',"
                sSql = sSql & " cPersFICargo = '" & FtesIngreso(I).sCargo & "',"
                sSql = sSql & " dPersFIInicio = '" & Format(FtesIngreso(I).dFteIngreso, "mm/dd/yyyy") & "',"
                sSql = sSql & " cPersFIMoneda = " & Trim(FtesIngreso(I).sMoneda) & ","
                sSql = sSql & " cPersFIComentario = '" & Trim(FtesIngreso(I).sComentarios) & "',"
                sSql = sSql & " cUltimaActualizacion = '" & sActualiza & "',"
                sSql = sSql & " cRazSocUbiGeo = '" & FtesIngreso(I).sRazSocUbiGeo & "',"
                sSql = sSql & " cRazSocDirecc = '" & FtesIngreso(I).sRazSocDirecc & "',"
                sSql = sSql & " cRazSocTelef = '" & FtesIngreso(I).sRazSocTelef & "',"
                sSql = sSql & " cRazSocDescrip = '" & FtesIngreso(I).sRazSocDescrip & "' "
                sSql = sSql & " WHERE cNumFuente = '" & FtesIngreso(I).cNumFte & "'"
                ConexionActiva.Execute sSql
            Else
                If FtesIngreso(I).nTipoAct = PersFilaEliminda Then
                    sSql = "DELETE PersFIIndependiente Where cNumFuente = '" & FtesIngreso(I).cNumFte & "' "
                    ConexionActiva.Execute sSql
                    
                    sSql = "DELETE PersFIDependiente Where cNumFuente = '" & FtesIngreso(I).cNumFte & "' "
                    ConexionActiva.Execute sSql
                        
                    sSql = "DELETE " & gConsPersona & "PersFteIngreso "
                    sSql = sSql & " WHERE cNumFuente = '" & FtesIngreso(I).cNumFte & "'"
                    ConexionActiva.Execute sSql
                End If
            End If
        End If
    Next I
    
    'FUENTES DE INGRESO DEPENDIENTES
    For I = 0 To nNumFtes - 1
        For j = 0 To FtesIngreso(I).nNumFtesDep - 1
            If FtesIngreso(I).FtesIngresoDep(j).nTipoAct = PersFilaNueva Then
                sSql = "INSERT INTO PersFIDependiente(cNumFuente, dPersEval, nPersIngCli, "
                sSql = sSql & " nPersGastoFam, nPersIngCon, nPersOtrIng, dPersFICaduca) "
                sSql = sSql & " VALUES('" & FtesIngreso(I).cNumFte & "',"
                sSql = sSql & "'" & Format(FtesIngreso(I).FtesIngresoDep(j).dFecEval, "mm/dd/yyyy") & "',"
                sSql = sSql & Format(FtesIngreso(I).FtesIngresoDep(j).nIngresos, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(I).FtesIngresoDep(j).nGastosFam, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(I).FtesIngresoDep(j).nIngCony, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(I).FtesIngresoDep(j).nIngOtros, "#0.00") & ","
                sSql = sSql & "'" & Format(FtesIngreso(I).FtesIngresoDep(j).dFecCad, "mm/dd/yyyy") & "')"
                ConexionActiva.Execute sSql
            Else
                If FtesIngreso(I).FtesIngresoDep(j).nTipoAct = PersFilaModificada Then
                    sSql = "UPDATE PersFIDependiente SET "
                    sSql = sSql & " nPersIngCli = " & Format(FtesIngreso(I).FtesIngresoDep(j).nIngFam, "#0.00") & ","
                    sSql = sSql & " nPersGastoFam = " & Format(FtesIngreso(I).FtesIngresoDep(j).nGastosFam, "#0.00") & ","
                    sSql = sSql & " nPersIngCon = " & Format(FtesIngreso(I).FtesIngresoDep(j).nIngCony, "#0.00") & ","
                    sSql = sSql & " nPersOtrIng = " & Format(FtesIngreso(I).FtesIngresoDep(j).nIngOtros, "#0.00") & ","
                    sSql = sSql & " dPersFICaduca = '" & Format(FtesIngreso(I).FtesIngresoDep(j).dFecCad, "mm/dd/yyyy") & "' "
                    sSql = sSql & " Where cNumFuente = '" & FtesIngreso(I).cNumFte & "' AND dPersEval = '" & Format(FtesIngreso(I).FtesIngresoDep(j).dFecEval, "mm/dd/yyyy") & "'"
                    ConexionActiva.Execute sSql
                Else
                    If FtesIngreso(I).FtesIngresoDep(j).nTipoAct = PersFilaEliminda Then
                        sSql = "DELETE PersFIDependiente Where cNumFuente = '" & FtesIngreso(I).cNumFte & "' AND dPersEval = '" & Format(FtesIngreso(I).FtesIngresoDep(j).dFecEval, "mm/dd/yyyy") & "'"
                        ConexionActiva.Execute sSql
                    End If
                End If
            End If
        Next j
    Next I
    
    'FUENTES DE INGRESO INDEPENDIENTES
    For I = 0 To nNumFtes - 1
        For j = 0 To FtesIngreso(I).nNumFtesIndep - 1
            If FtesIngreso(I).FtesIngresoInd(j).nTipoAct = PersFilaNueva Then
                sSql = "INSERT INTO PersFIIndependiente(cNumFuente, dPersEval, "
                sSql = sSql & " dPersFICaduca, nPersFIActivoDisp, nPersFICtasxCobrar, nPersFIInventarios, "
                sSql = sSql & " nPersFIActivosFijos, nPersFIProveedores, nPersFICredCMACT, nPersFICredOtros, "
                sSql = sSql & " nPersFIVentas, nPersFIPatrimonio, nPersFICostoVentas, nPersFIRecupCtasXCobrar, "
                sSql = sSql & " nPersFIEgresosOtros, nPersIngFam, nPersEgrFam, cUltimaActualizacion)"
                sSql = sSql & " VALUES('" & FtesIngreso(I).cNumFte & "',"
                sSql = sSql & "'" & Format(FtesIngreso(I).FtesIngresoInd(j).dFecEval, "mm/dd/yyyy") & "',"
                sSql = sSql & "'" & Format(FtesIngreso(I).FtesIngresoInd(j).dFecCad, "mm/dd/yyyy") & "',"
                sSql = sSql & Format(FtesIngreso(I).FtesIngresoInd(j).nActivDisp, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(I).FtesIngresoInd(j).nCtasxCob, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(I).FtesIngresoInd(j).nInventarios, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(I).FtesIngresoInd(j).nActFijo, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(I).FtesIngresoInd(j).nProveedores, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(I).FtesIngresoInd(j).nCreditosCmact, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(I).FtesIngresoInd(j).nCreditosOtros, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(I).FtesIngresoInd(j).nVentas, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(I).FtesIngresoInd(j).nPasivoPatrimonio, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(I).FtesIngresoInd(j).nCompraMercad, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(I).FtesIngresoInd(j).nRecupCtasxCob, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(I).FtesIngresoInd(j).nOtrosEgresos, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(I).FtesIngresoInd(j).nPersIngFam, "#0.00") & ","
                sSql = sSql & Format(FtesIngreso(I).FtesIngresoInd(j).nPersEgrFam, "#0.00") & ","
                sSql = sSql & "'" & sActualiza & "')"
                ConexionActiva.Execute sSql
                
            Else
                If FtesIngreso(I).FtesIngresoInd(j).nTipoAct = PersFilaModificada Then
                    sSql = " UPDATE PersFIIndependiente SET "
                    sSql = sSql & " nPersFIActivoDisp = " & Format(FtesIngreso(I).FtesIngresoInd(j).nActivDisp, "#0.00") & ","
                    sSql = sSql & " nPersFICtasxCobrar = " & Format(FtesIngreso(I).FtesIngresoInd(j).nCtasxCob, "#0.00") & ","
                    sSql = sSql & " nPersFIInventarios = " & Format(FtesIngreso(I).FtesIngresoInd(j).nInventarios, "#0.00") & ","
                    sSql = sSql & " nPersFIActivosFijos = " & Format(FtesIngreso(I).FtesIngresoInd(j).nActFijo, "#0.00") & ","
                    sSql = sSql & " nPersFIProveedores = " & Format(FtesIngreso(I).FtesIngresoInd(j).nProveedores, "#0.00") & ","
                    sSql = sSql & " nPersFICredCMACT = " & Format(FtesIngreso(I).FtesIngresoInd(j).nCreditosCmact, "#0.00") & ","
                    sSql = sSql & " nPersFICredOtros = " & Format(FtesIngreso(I).FtesIngresoInd(j).nCreditosOtros, "#0.00") & ","
                    sSql = sSql & " nPersFIVentas = " & Format(FtesIngreso(I).FtesIngresoInd(j).nVentas, "#0.00") & ","
                    sSql = sSql & " nPersFIPatrimonio = " & Format(FtesIngreso(I).FtesIngresoInd(j).nPasivoPatrimonio, "#0.00") & ","
                    sSql = sSql & " nPersFICostoVentas = " & Format(FtesIngreso(I).FtesIngresoInd(j).nCompraMercad, "#0.00") & ","
                    sSql = sSql & " nPersFIRecupCtasXCobrar = " & Format(FtesIngreso(I).FtesIngresoInd(j).nRecupCtasxCob, "#0.00") & ","
                    sSql = sSql & " nPersFIEgresosOtros = " & Format(FtesIngreso(I).FtesIngresoInd(j).nOtrosEgresos, "#0.00") & ","
                    sSql = sSql & " nPersIngFam = " & Format(FtesIngreso(I).FtesIngresoInd(j).nPersIngFam, "#0.00") & ","
                    sSql = sSql & " nPersEgrFam = " & Format(FtesIngreso(I).FtesIngresoInd(j).nPersEgrFam, "#0.00") & ","
                    sSql = sSql & " dPersFICaduca = '" & Format(FtesIngreso(I).FtesIngresoInd(j).dFecCad, "mm/dd/yyyy") & "' "
                    sSql = sSql & " Where cNumFuente = '" & FtesIngreso(I).cNumFte & "' AND dPersEval = '" & Format(FtesIngreso(I).FtesIngresoInd(j).dFecEval, "mm/dd/yyyy") & "'"
                    
                    ConexionActiva.Execute sSql
                Else
                    If FtesIngreso(I).FtesIngresoInd(j).nTipoAct = PersFilaEliminda Then
                        sSql = "DELETE PersFIIndependiente Where cNumFuente = '" & FtesIngreso(I).cNumFte & "' AND dPersEval = '" & Format(FtesIngreso(I).FtesIngresoInd(j).dFecEval, "mm/dd/yyyy") & "'"
                        ConexionActiva.Execute sSql
                    End If
                End If
            End If
        Next j
    Next I
    
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
    
                sSql = "UPDATE Persona SET "
                sSql = sSql & " cNomPers = '" & sNombreCompleto & "',"
                sSql = sSql & " cTidoci = '" & vDocs(0) & "',"
                sSql = sSql & " cNudoci = '" & vDocs(1) & "',"
                sSql = sSql & " cTidoTr = '" & vDocs(2) & "',"
                sSql = sSql & " cNudoTr = '" & vDocs(3) & "',"
                sSql = sSql & " cDirPers = '" & sDomicilio & "',"
                sSql = sSql & " cCodZon = '" & sUbicGeografica & "',"
                sSql = sSql & " dFecNac = '" & Format(dFechaNac, "yyyy/mm/dd") & "',"
                sSql = sSql & " cTelPers = '" & sTelefonos & "',"
                sSql = sSql & " cEstPers = '1',"
                sSql = sSql & " cTipPers = '" & IIf(nTipoPersona = 0, 1, nTipoPersona) & "',"
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
    
                    sSql = "Update PersonaNat SET "
                    sSql = sSql & " cSexPers = '" & sSexo & "',"
                    sSql = sSql & " cEstCiv = '" & sEstadoCivil & "',"
                    sSql = sSql & " nNumHij = " & nHijos
                    sSql = sSql & " Where cCodPers = '" & sCodPersDistrib & "'"

                    'sSql = "INSERT INTO Personanat(cCodPers, cSexPers, cEstCiv, cCodProf, cCodCon, nNumHij)"
                    'sSql = sSql & " VALUES('" & sCodPersDistrib & "','" & sSexo & "','" & sEstadoCivil & "','',''," & nHijos & ")"
                    ConexDitrib.Execute sSql
    
                End If
    
                '******************************************************************************
                'ACTUALIZA PERSONA JURIDICA EN DISTRIBUIDO
                '******************************************************************************
                If bPersDistrib Then
    
                    sSql = "UPDATE PersonaJur SET "
                    sSql = sSql & " cSigla = '" & sSiglas & "' "
                    sSql = sSql & " Where cCodPers = '" & sCodPersDistrib & "'"

                    'sSql = "Insert into PersonaJur(cCodPers,cTipPerJur,cCodRegPub,cMagEmp,cSigla) "
                    'sSql = sSql & "Values('" & sCodPersDistrib & "','0000','','','" & sSiglas & "')"
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
    I = 0
    Do While I <= nNumDoc - 1
        If Docs(I).nTipoAct = PersFilaEliminda Then
            For j = I To nNumDoc - 2
                Docs(I) = Docs(I + 1)
            Next j
            nNumDoc = nNumDoc - 1
            ReDim Preserve Docs(nNumDoc)
            I = I - 1
        Else
            Docs(I).nTipoAct = PersFilaSinCambios
        End If
        I = I + 1
    Loop
    I = 0
    Do While I <= nNumRel - 1
        If Relac(I).nTipoAct = PersFilaEliminda Then
            For j = I To nNumRel - 2
                Relac(I) = Relac(I + 1)
            Next j
            nNumRel = nNumRel - 1
            ReDim Preserve Relac(nNumRel)
            I = I - 1
        Else
            Relac(I).nTipoAct = PersFilaSinCambios
        End If
        I = I + 1
    Loop
    I = 0
    Do While I <= nNumFtes - 1
        If FtesIngreso(I).nTipoAct = PersFilaEliminda Then
            For j = I To nNumFtes - 2
                FtesIngreso(I) = FtesIngreso(I + 1)
            Next j
            nNumFtes = nNumFtes - 1
            ReDim Preserve FtesIngreso(nNumFtes)
            I = I - 1
        Else
            FtesIngreso(I).nTipoAct = PersFilaSinCambios
        End If
        I = I + 1
    Loop
    
    I = 0
    Do While I <= nNumRefCom - 1
        If RefCom(I).nTipoAct = PersFilaEliminda Then
            For j = I To nNumRefCom - 2
                RefCom(I) = RefCom(I + 1)
            Next j
            nNumRefCom = nNumRefCom - 1
            ReDim Preserve RefCom(nNumRefCom)
            I = I - 1
        Else
            RefCom(I).nTipoAct = PersFilaSinCambios
        End If
        I = I + 1
    Loop
    
    I = 0
    Do While I <= nNumRefBan - 1
        If RefBan(I).nTipoAct = PersFilaEliminda Then
            For j = I To nNumRefBan - 2
                RefBan(I) = RefBan(I + 1)
            Next j
            nNumRefBan = nNumRefBan - 1
            ReDim Preserve RefBan(nNumRefBan)
            I = I - 1
        Else
            RefBan(I).nTipoAct = PersFilaSinCambios
        End If
        I = I + 1
    Loop
    
    Exit Sub
    
    I = 0
    Do While I <= nNumPatVeh - 1
        If PatVeh(I).nTipoAct = PersFilaEliminda Then
            For j = I To nNumPatVeh - 2
                PatVeh(I) = PatVeh(I + 1)
            Next j
            nNumPatVeh = nNumPatVeh - 1
            ReDim Preserve PatVeh(nNumPatVeh)
            I = I - 1
        Else
            PatVeh(I).nTipoAct = PersFilaSinCambios
        End If
        I = I + 1
    Loop
    
    Exit Sub
ErrorGrabarCambiosPersona:
    If InicioTran Then
        ConexionActiva.RollBackTrans
    End If
    If TranDistribBegin Then
        ConexDitrib.RollBackTrans
    End If
    Err.Raise Err.Number, "GrabaPersona", Err.Description
End Sub

Public Sub EditarDocumento(ByVal pnTipo As Long, ByVal psNumero As String)
Dim I As Integer
    For I = 0 To nNumDoc - 1
        If Docs(I).sTipo = pnTipo And Docs(I).sNumero = psNumero Then
            Docs(I).sTipo = pnTipo
            Docs(I).sNumero = psNumero
            Exit For
        End If
    Next I
End Sub
Public Sub EliminarDocumento(ByVal pnTipo As Long, ByVal psNumero As String)
Dim I As Integer
Dim nPos  As Integer
    nPos = -1
    For I = 0 To nNumDoc - 1
        If Docs(I).nTipo = pnTipo And Docs(I).sNumero = psNumero Then
            nPos = I
            Exit For
        End If
    Next I
    If nPos <> -1 Then
        For I = nPos To nNumDoc - 2
            Docs(I) = Docs(I + 1)
        Next I
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

Public Sub AdicionaFteIngreso(ByVal pnTipoFte As PersFteIngresoTipo)
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
    End If
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
Dim sSql As String
Dim R As ADODB.Recordset

    sSql = "Select C.cConsDescripcion,PI.cPersCod,PI.cPersIDtpo,PI.cPersIDnro from PersID PI left join Constante C ON PI.cPersIDtpo = C.cConsValor"
    sSql = sSql & " where PI.cPersCod = '" & psCodPers & "' And C.cConsCod = " & Trim(Str(gPersIdTipo))
    Set R = Conn.CargaRecordSet(sSql, adLockReadOnly)
    Do While Not R.EOF
        Call AdicionaDocumento(R!cPersIDTpo, R!cPersIDnro, R!cConsDescripcion)
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
Dim sSql As String
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
    sSql = "Select * from Servidor Where RIGHT(cCodAge,2) = '" & sCodAge & "' AND cNroSer = '05'"
    Set R = New ADODB.Recordset
    R.Open sSql, ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    If R.RecordCount > 0 Then
        sCadCon = "Provider=" & Trim(R!cProvider) & ";User ID=" & Trim(R!cLogin) & ";Initial Catalog=" & Trim(R!cDataBase) & ";Data Source=" & Trim(R!cNomSer) & ";PWD=" & Trim(R!cPass)
    End If
    R.Close
    
    If sCadCon <> "" Then
        Set oConecta = New DConecta
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
        Docs(nNumDoc - 1).sNumero = Trim(R!cPersIDnro)
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
        Relac(nNumRel - 1).sMovAct = R!cMovNro
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

Public Function FuenteIngresoAsignadaACredito(ByVal pnIndex As Integer, ByVal pdFecEval As Date) As Boolean
Dim sSql As String
Dim R As ADODB.Recordset
    
    sSql = "Select * from ColocFteIngreso where cNumFuente = '" & FtesIngreso(pnIndex).cNumFte & "' AND dPersEval = '" & Format(pdFecEval, "mm/dd/yyyy") & "'"
    Set R = Conn.CargaRecordSet(sSql)
    If R.RecordCount > 0 Then
        FuenteIngresoAsignadaACredito = True
    Else
        FuenteIngresoAsignadaACredito = False
    End If
    Set R = Nothing
End Function

Public Function UbicaPosicionFteIngreso(ByVal psPersCodFI As String, ByVal pdFecEval As Date, ByVal pnIndice As Integer) As Integer
Dim I As Integer
    If FtesIngreso(pnIndice).nTipo = gPersFteIngresoTipoDependiente Then
        For I = 0 To FtesIngreso(pnIndice).nNumFtesDep - 1
            If FtesIngreso(pnIndice).FtesIngresoDep(I).dFecEval = pdFecEval Then
                UbicaPosicionFteIngreso = I
                Exit Function
            End If
        Next I
    Else
        For I = 0 To FtesIngreso(pnIndice).nNumFtesIndep - 1
            If FtesIngreso(pnIndice).FtesIngresoInd(I).dFecEval = pdFecEval Then
                UbicaPosicionFteIngreso = I
                Exit Function
            End If
        Next I
    End If
End Function

Friend Function PersFuenteBuscaUltimaFechaEval(ByVal psNumFte As String) As Date
Dim I As Integer
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
    
    For I = 0 To nNumFtes - 1
        If FtesIngreso(I).cNumFte = psNumFte Then
            If FtesIngreso(I).nTipo = gPersFteIngresoTipoDependiente Then
                For j = 0 To FtesIngreso(I).nNumFtesDep - 1
                    If FtesIngreso(I).FtesIngresoDep(j).dFecEval > dFecEvalTmp Then
                        dFecEvalTmp = FtesIngreso(I).FtesIngresoDep(j).dFecEval
                    End If
                Next j
            Else
                For j = 0 To FtesIngreso(I).nNumFtesIndep - 1
                    If FtesIngreso(I).FtesIngresoInd(j).dFecEval > dFecEvalTmp Then
                        dFecEvalTmp = FtesIngreso(I).FtesIngresoInd(j).dFecEval
                    End If
                Next j
            End If
        End If
    Next I
    PersFuenteBuscaUltimaFechaEval = dFecEvalTmp
End Function

Friend Function PersFuenteBuscaUltimaFechaCaducacion(ByVal psNumFte As String) As Date
Dim I As Integer
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
    For I = 0 To nNumFtes - 1
        If FtesIngreso(I).cNumFte = psNumFte Then
            If FtesIngreso(I).nTipo = gPersFteIngresoTipoDependiente Then
                For j = 0 To FtesIngreso(I).nNumFtesDep - 1
                    If FtesIngreso(I).FtesIngresoDep(j).dFecCad > dFecEvalTmp Then
                        dFecEvalTmp = FtesIngreso(I).FtesIngresoDep(j).dFecCad
                    End If
                Next j
            Else
                For j = 0 To FtesIngreso(I).nNumFtesIndep - 1
                    If FtesIngreso(I).FtesIngresoInd(j).dFecCad > dFecEvalTmp Then
                        dFecEvalTmp = FtesIngreso(I).FtesIngresoInd(j).dFecCad
                    End If
                Next j
            End If
        End If
    Next I
    PersFuenteBuscaUltimaFechaCaducacion = dFecEvalTmp
End Function

Public Function FiltraFuentesIngresoPorRazonSocial() As Variant
Dim MatFte() As String
Dim MatFteTmp() As String
Dim I As Integer
Dim nNumFtetmp As Integer
Dim sPersCodFI As String

    
    ReDim MatFteTmp(nNumFtes, 9)
    nNumFtetmp = 0
    If nNumFtes > 0 Then
        For I = 0 To nNumFtes - 1
            nNumFtetmp = nNumFtetmp + 1
            MatFteTmp(nNumFtetmp - 1, 0) = nNumFtetmp
            MatFteTmp(nNumFtetmp - 1, 1) = IIf(FtesIngreso(I).nTipo = gPersFteIngresoTipoDependiente, "D", "I")
            MatFteTmp(nNumFtetmp - 1, 2) = FtesIngreso(I).sRazSocDescrip
            'MatFteTmp(nNumFtetmp - 1, 3) = PersFuenteBuscaUltimaFechaCaducacion(FtesIngreso(I).cNumFte) 'Fecha de Caducacion de la Fte de Ingreso
            'MatFteTmp(nNumFtetmp - 1, 4) = PersFuenteBuscaUltimaFechaEval(FtesIngreso(I).cNumFte) 'Fecha de Evaluacion de la Fte de Ingreso
            MatFteTmp(nNumFtetmp - 1, 5) = IIf(FtesIngreso(I).sMoneda = gMonedaNacional, "SOLES", "DOLARES") 'Moneda de la Fte de Ingreso
            MatFteTmp(nNumFtetmp - 1, 6) = FtesIngreso(I).sFteIngreso
            MatFteTmp(nNumFtetmp - 1, 7) = Trim(Str(I))
            MatFteTmp(nNumFtetmp - 1, 8) = FtesIngreso(I).cNumFte
            sPersCodFI = FtesIngreso(I).sFteIngreso
        Next I
        
        ReDim MatFte(nNumFtetmp, 8)
        For I = 0 To nNumFtetmp - 1
            MatFte(I, 0) = MatFteTmp(I, 0)
            MatFte(I, 1) = MatFteTmp(I, 1)
            MatFte(I, 2) = MatFteTmp(I, 2)
            MatFte(I, 3) = MatFteTmp(I, 3)
            MatFte(I, 4) = MatFteTmp(I, 4)
            MatFte(I, 5) = MatFteTmp(I, 5)
            MatFte(I, 6) = MatFteTmp(I, 6)
            MatFte(I, 7) = MatFteTmp(I, 7)
            MatFte(I, 8) = MatFteTmp(I, 8)
        Next I
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
        sSql = sSql & " From PersFIDependiente PF Where PF.cNumFte = '" & psNumFte & "'"
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
        sSql = sSql & " From PersFIIndependiente PF Where PF.cNumFte = '" & psNumFte & "'"
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

Public Sub RecuperaFtesdeIngreso(ByVal psCodPers As String)
Dim sSql As String
Dim R As ADODB.Recordset
Dim bConexInterna As Boolean
    
    sSql = "Select cNumFuente, cPersCod, cPersFIPersCod, cPersFICargo, "
    sSql = sSql & " dPersFIinicio, cPersFIMoneda, cPersFIComentario, "
    sSql = sSql & " cRazSocUbiGeo, cRazSocDirecc, cRazSocTelef, cRazSocDescrip, "
    sSql = sSql & " nPersTipFte, cUltimaActualizacion "
    sSql = sSql & " From  PersFteIngreso PF Where PF.cPersCod = '" & psCodPers & "'"
    
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
        FtesIngreso(nNumFtes - 1).cNumFte = R!cNumFuente
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
    sSql = "Select PF.cNumFuente, PF.dPersEval, PF.nPersIngCli, PF.nPersGastoFam, PF.nPersIngCon, PF.nPersOtrIng, PF.dPersFICaduca "
    sSql = sSql & " FROM PersFIDependiente PF Where PF.cNumFuente = '" & FtesIngreso(pnNumFte).cNumFte & "'"
    sSql = sSql & " ORDER BY PF.dPersEval ASC"
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

Public Sub RecuperaFtesIngresoIndependiente(ByVal pnNumFte As Integer)
Dim sSql As String
Dim R As ADODB.Recordset

    FtesIngreso(pnNumFte).nNumFtesIndep = 0
    sSql = "SELECT PF.cNumFuente, PF.dPersEval, PF.dPersFICaduca, PF.nPersFIActivoDisp, "
    sSql = sSql & " PF.nPersFICtasxCobrar, PF.nPersFIInventarios, PF.nPersFIActivosFijos, "
    sSql = sSql & " PF.nPersFIProveedores, PF.nPersFICredCMACT, PF.nPersFICredOtros, "
    sSql = sSql & " PF.nPersFIVentas, PF.nPersFIPatrimonio, PF.nPersFICostoVentas, "
    sSql = sSql & " PF.nPersFIRecupCtasXCobrar, PF.nPersFIEgresosOtros, PF.nPersIngFam, "
    sSql = sSql & " PF.nPersEgrFam, PF.cUltimaActualizacion "
    sSql = sSql & " FROM PersFIIndependiente PF Where PF.cNumFuente = '" & FtesIngreso(pnNumFte).cNumFte & "'"
    sSql = sSql & " ORDER BY PF.dPersEval ASC"
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

Public Sub RecuperaPersona(ByVal psCodPers As String)
Dim sSql As String
Dim R As ADODB.Recordset
    
    Set Conn = New DConecta
    Conn.AbreConexion
    Set ConexionActiva = Conn.ConexionActiva
    
    '***** CARGA DATOS DE LA PERSONA ******
    sSql = "Select P.nPersRelaInst, P.cPersCodSbs, P.cPersEmail cEmail,P.cPersTelefono2,P.cPersCod,P.cPersNombre,P.dPersNacCreac, "
    sSql = sSql & "P.cPersDireccUbiGeo,P.cPersDireccDomicilio,P.cPersDireccCondicion, P.nPersValComDomicilio, P.cPersTelefono, "
    sSql = sSql & "P.cPersCelular,P.cPersCelular2,P.cPersCelular3,P.cPersEmail2 cEmail2, " 'EJVG20111212
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
        dFechaNac = CDate(Format(R!dPersNacCreac, "dd/mm/yyyy"))
        sUbicGeografica = R!cPersDireccUbiGeo
        sDomicilio = R!cPersDireccDomicilio
        sCondicionDomic = IIf(IsNull(R!cPersDireccCondicion), Str(gPersCondDomicEnUso), R!cPersDireccCondicion)
        nValComDomicilio = IIf(IsNull(R!nPersValComDomicilio), 0, R!nPersValComDomicilio)
        sTelefonos = IIf(IsNull(R!cPersTelefono), "", R!cPersTelefono)
        sTelefonos2 = IIf(IsNull(R!cPersTelefono2), "", R!cPersTelefono2)
        'EJVG20111212 *************************
        sCelular = IIf(IsNull(R!cPersCelular), "", R!cPersCelular)
        sCelular2 = IIf(IsNull(R!cPersCelular2), "", R!cPersCelular2)
        sCelular3 = IIf(IsNull(R!cPersCelular3), "", R!cPersCelular3)
        
        sEmail = IIf(IsNull(R!cEmail), "", R!cEmail)
        sEmail2 = IIf(IsNull(R!cEmail2), "", R!cEmail2)
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
    Call RecuperaFtesdeIngreso(psCodPers)
    Call RecuperaRefComercial(psCodPers)   'Carga las Referencias Comerciales de la Persona
    Call RecuperaMaxRefComercial(psCodPers)
    Call RecuperaRefBancaria(psCodPers)     'Carga las Referencias Bancarias de la Persona
    Call RecuperaPatVehicular(psCodPers)    'Carga el Patrimonio Vehicular
    Call RecuperaMaxPatVehicular(psCodPers)
    Exit Sub
        
End Sub
'EJVG20110107 ************************
Public Function validaPermisoEditarPersona(ByVal psCargoCod As String) As Boolean
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim oConecta As DConecta
    
    Set rs = New ADODB.Recordset
    sSql = "exec stp_sel_ObtienePermisoEditarPersona '" & psCargoCod & "'"
    Set oConecta = New DConecta
    
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    If Not rs.EOF And Not rs.BOF Then
        validaPermisoEditarPersona = rs!bPermiso
    Else
        validaPermisoEditarPersona = False
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

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

Public Property Let Telefonos(ByVal vNewValue As String)
    sTelefonos = vNewValue
End Property

Public Function ObtenerDNI() As String
Dim I As Integer
    ObtenerDNI = ""
    For I = 0 To nNumDoc - 1
        If Docs(I).nTipo = gPersIdDNI Then
            ObtenerDNI = Docs(I).sNumero
            Exit Function
        End If
    Next I
End Function

Public Function ObtenerRUC() As String
Dim I As Integer
    ObtenerRUC = ""
    For I = 0 To nNumDoc - 1
        If Docs(I).nTipo = gPersIdRUC Then
            ObtenerRUC = Docs(I).sNumero
            Exit Function
        End If
    Next I
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
        FtesIngreso(pnPunt).FtesIngresoDep(pnPuntDet).nTipoAct = pCambio
    Else
        FtesIngreso(pnPunt).FtesIngresoInd(pnPuntDet).nTipoAct = pCambio
    End If
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


Public Function BuscaProveedor(ByVal sCadBusca As String, Optional ByVal nTipoBusq As TBusqueda = BusquedaNombre) As ADODB.Recordset
Dim lsCadena As String

Dim prm As ADODB.Parameter
Dim Conn As DConecta
    bError = False
    Set Conn = New DConecta
    'temporal
    lsCadena = "PROVIDER=SQLOLEDB;User ID=sa;Password=cmacica;INITIAL CATALOG=dbcmacicaproveedores;DATA SOURCE=01SRVSICMAC02"
    If Not Conn.AbreConexion(lsCadena) Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set BuscaProveedor = Nothing
        Set Conn = Nothing
        
        Exit Function
    End If
    Conn.ConexionActiva.CommandTimeout = 7200
    Select Case nTipoBusq
        Case 1 'Busqueda por nombre
            Set BuscaProveedor = Conn.CargaRecordSet("spBuscaProveedorApellido '" & Trim(sCadBusca) & "'")
        'Case 2 'Busqueda por Codigo
        '    Set BuscaCliente = Conn.CargaRecordSet("spBuscaClienteCodigo '" & Trim(sCadBusca) & "'")
        Case 3 'Busqueda por Documento
            Set BuscaProveedor = Conn.CargaRecordSet("spBuscaProveedoresDoc '" & Trim(sCadBusca) & "'")
            
    End Select
    Conn.CierraConexion
    Set Conn = Nothing
End Function


