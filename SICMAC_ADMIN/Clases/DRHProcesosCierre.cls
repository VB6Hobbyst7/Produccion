VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DRHProcesosCierre"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Actualiza Datos de la Relacion concepto Empleado
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB9030101E8
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB90301035B
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB90301031F
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

'Agrega una nueva relación entre concepto y Empleado
'##ModelId=3A9BEAB1011F
Public Function CierreDia(pdFecha As Date) As ADODB.Recordset
    On Error GoTo AgregaConceptoEmpleadoErr
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlC = " Create Table #RHTemporal" _
         & " (" _
         & " cPersCod Varchar(13) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL," _
         & " nRHEstadoAnt Int NOT NULL," _
         & " nRHEstado Int Not Null" _
         & " )"

    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.Ejecutar "spRHCierreDia '" & Format(pdFecha, gsFormatoFecha) & "'"
        sqlC = " Select RH.cRHCod EmpCod,  TMP.cPersCod Codigo, PE.cPersNombre Nombre, nRHEstadoAnt CodAnt, CO1.cConsDescripcion DesAnt, TMP.nRHEstado CodNue, CO2.cConsDescripcion DesNue" _
             & " From #RHTemporal TMP" _
             & " Inner Join Persona PE On TMP.cPersCod = PE.cPersCod" _
             & " Inner Join Constante CO1 On CO1.nConsValor = TMP.nRHEstadoAnt And CO1.nConsCod = '" & gRHEstado & "'" _
             & " Inner Join Constante CO2 On CO2.nConsValor = TMP.nRHEstado And CO2.nConsCod = '" & gRHEstado & "'" _
             & " Inner Join RRHH RH On RH.cPersCod = TMP.cPersCod " _
             & " Order by PE.cPersNombre"
        Set CierreDia = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
AgregaConceptoEmpleadoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoEmpleado:AgregaConceptoEmpleado Method")
End Function

'Agrega una nueva relación entre concepto y Empleado
'##ModelId=3A9BEAB1011F
Public Function CierreMes(pdFecha As Date) As ADODB.Recordset
    On Error GoTo AgregaConceptoEmpleadoErr
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlC = " Create Table #RHTemporal" _
         & " (" _
         & " cPersCod Varchar(13) Not Null," _
         & " nRHEstadoAnt Int Not Null," _
         & " nRHEstado Int Not Null" _
         & " )"

    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.Ejecutar "spRHCierreDia '" & Format(pdFecha, gsFormatoFecha) & "'"
        sqlC = "Select cPersCod, nRHEstadoAnt,nRHEstado From #RHTemporal "
        Set CierreMes = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
AgregaConceptoEmpleadoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoEmpleado:AgregaConceptoEmpleado Method")
End Function

Public Function GetNroNotaCargoAbono(ByVal pnTpoDoc As TpoDoc) As String
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If oCon.AbreConexion = False Then Exit Function
    
    sql = " Select Max(MaxNota) as MaxNota  From" _
        & " (" _
        & "   Select Max(Convert(int,Isnull(cDocNro,0)))  as MaxNota" _
        & "          From NotaAbonoCargo" _
        & "          Where nDocTpo = '" & pnTpoDoc & "' and isnumeric(cDocNro)  = 1 " _
        & "   Union" _
        & "   Select Max(Convert(Int,Isnull(cDocNro,0))) as MaxNota" _
        & "   From RegDocCuenta" _
        & "     Where ndoctpo = '" & pnTpoDoc & "' and isnumeric(cDocNro)  = 1 ) As Nota"

    Set rs = oCon.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        GetNroNotaCargoAbono = Format(rs!MaxNota + 1, "00000000")
    Else
        GetNroNotaCargoAbono = Format(rs!MaxNota + 1, "00000000")
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function GeneraAsientoRemuneracion(ByVal psOpe As String, ByVal psPlanillaCod As String, ByVal psPlanillaReriodo As String, ByVal pdFecha As Date, psCond As String, pgsCodAge As String, pgsCodUser As String) As Long
    Dim rsPla As ADODB.Recordset
    Dim lsMov As String, lsGlosa As String, lsCta As String
    Set rsPla = New ADODB.Recordset
    Dim lnDebe As Double, lnHaber As Double, lnOtro As Double
    Dim lnDH As Double, lnMonto As Double
    Dim lnItem As Integer
    Dim lsSQLC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lsqlP As String
    Dim sqlP As String
    Dim lnMov As Long
    Dim oMov As DMov
    Set oMov = New DMov
    Dim rsCred As ADODB.Recordset
    Set rsCred = New ADODB.Recordset
    Dim lnMontoHipoteca As Currency
    Dim lnTipCamV As Currency
    Dim lnTipCamFijo As Currency
    Dim sSQLC As String
    
    Dim oConec As NConstSistemas
    Set oConec = New NConstSistemas
    
    Dim lsCtaAjusteTipoCambioPerdida As String
    Dim lsCtaAjusteTipoCambioGanancia As String
    
    Dim lsCtaPendRRHHSoles As String
    Dim lsCtaPendRRHHDolares As String
    Dim lsCtaPendRRHHCredOtrosSoles As String
    Dim lsCtaPendRRHHCredOtrosDolares As String
    
    Dim lsCtaOtrosIngresos As String
    Dim lsCtaOtrosEgresos As String
        
    Dim lnMontoDebe As Double, lnMontoHaber As Double
    Dim lnItfDebe As Double, lnItfHaber As Double
    Dim rsAsiento As ADODB.Recordset
    Set rsAsiento = New ADODB.Recordset
    Dim lnItenDebe As Integer, lnItenHaber As Integer
    Dim lnImporteDebe As Double, lnImporteHaber As Double
        
    'Ajustes
    lsCtaAjusteTipoCambioPerdida = oConec.LeeConstSistema(gConstSistCtaAjustTCPerdida)
    lsCtaAjusteTipoCambioGanancia = oConec.LeeConstSistema(gConstSistCtaAjustTCGanancia)
    'Pendientes
    lsCtaPendRRHHSoles = oConec.LeeConstSistema(gConstSistCtaPendRHSoles)
    lsCtaPendRRHHDolares = oConec.LeeConstSistema(gConstSistCtaPendRHDolares)
    lsCtaPendRRHHCredOtrosSoles = oConec.LeeConstSistema(gConstSistCtaPendRHCredOtrosSoles)
    lsCtaPendRRHHCredOtrosDolares = oConec.LeeConstSistema(gConstSistCtaPendRHCredOtrosDolares)
    'Inrgesos / Egresos Otros
    lsCtaOtrosIngresos = oConec.LeeConstSistema(gConstSistCtaOtrosIngresos)
    lsCtaOtrosEgresos = oConec.LeeConstSistema(gConstSistCtaOtrosEgresos)
        
    If psCond <> "" Then sSQLC = " And nRHContratoTpo = '" & psCond & "' "
    
    Dim rsTIB As New ADODB.Recordset
    Dim sqlTIB As String
'ALPA 20091228************************************
'    sqlTIB = " Select T.cCtaAreaAsig, SUM(Monto) Monto" _
'         & " FROM ( " _
'         & "        Select E.cCtaAreaAsig, " _
'         & "        PD.nMonto Monto FROM RHPlanillaDetCon PD" _
'         & "            INNER JOIN (Select cPersCod, cRHAgenciaCodOficial cAgenciaAsig, (Select Top 1 cSubCtaCod from areaagencia where cAreaCod = cRHAreaCodOficial And cUbicaCod = cRHAgenciaCodOficial) cCtaAreaAsig From RHCargos RHC Where " _
'         & "                            dRHCargoFecha = (Select Max(RHC1.dRHCargoFecha) " _
'         & "                                                      From RHCargos RHC1 Where RHC.cPersCod  = RHC1.cPersCod And Left(RHC1.cUltimaActualizacion,8) <= '" & psPlanillaReriodo & "')) As E ON PD.cPersCod = E.cPersCod" _
'         & "            INNER JOIN RHContrato RHC ON PD.cPersCod = RHC.cPersCod " _
'         & "        WHERE PD.cRHConceptoCod = '130' And PD.cPlanillaCod = '" & psPlanillaCod & "' " _
'         & "        And PD.cRRHHPeriodo = '" & psPlanillaReriodo & "'" _
'         & "        And RHC.cRHContratoNro In ( Select Max(RHC1.cRHContratoNro) From RHContrato RHC1 Where RHC.cPersCod = RHC1.cPersCod  And Left(RHC1.cUltimaActualizacion,8) <= '" & psPlanillaReriodo & "') " & sSQLC & " " _
'         & "       ) T Group by T.cCtaAreaAsig Having SUM(Monto) > 0" _
'         & " Order by T.cCtaAreaAsig "
'ALPA 20091228************************************

    sqlTIB = " select cCtaAreaAsig,SUM(Monto) Monto "
    sqlTIB = sqlTIB & " From "
    sqlTIB = sqlTIB & "( "
    sqlTIB = sqlTIB & " Select T.cCtaAreaAsig, sum(Monto) Monto"
    sqlTIB = sqlTIB & " FROM ( "
    sqlTIB = sqlTIB & " Select E.cCtaAreaAsig, "
    sqlTIB = sqlTIB & "  PD.nMonto Monto FROM RHPlanillaDetCon PD"
    sqlTIB = sqlTIB & "    INNER JOIN (Select cPersCod, cRHAgenciaCodOficial cAgenciaAsig, (Select Top 1 cSubCtaCod from areaagencia where cAreaCod = cRHAreaCodOficial And cUbicaCod = cRHAgenciaCodOficial) cCtaAreaAsig From RHCargos RHC Where "
    sqlTIB = sqlTIB & "     dRHCargoFecha = (Select Max(RHC1.dRHCargoFecha) "
    sqlTIB = sqlTIB & "     From RHCargos RHC1 Where RHC.cPersCod  = RHC1.cPersCod And Left(RHC1.cUltimaActualizacion,8) <= '" & psPlanillaReriodo & "')) As E ON PD.cPersCod = E.cPersCod"
    sqlTIB = sqlTIB & "            INNER JOIN RHContrato RHC ON PD.cPersCod = RHC.cPersCod "
    sqlTIB = sqlTIB & "            INNER JOIN RRHH RH on PD.cPersCod=RH.cPersCod "
    sqlTIB = sqlTIB & "                 and (RH.cAgenciaActual<>'01' or (RH.cAgenciaActual='01' and rh.cAreaCodActual NOT IN (select cAreaCod from AreaPorcentajeGastos where cAgeCod='01'))) "
    sqlTIB = sqlTIB & "  WHERE PD.cRHConceptoCod = '130' And PD.cPlanillaCod = '" & psPlanillaCod & "' "
    sqlTIB = sqlTIB & "        And PD.cRRHHPeriodo = '" & psPlanillaReriodo & "'"
    sqlTIB = sqlTIB & "        And RHC.cRHContratoNro In ( Select Max(RHC1.cRHContratoNro) From RHContrato RHC1 Where RHC.cPersCod = RHC1.cPersCod  And Left(RHC1.cUltimaActualizacion,8) <= '" & psPlanillaReriodo & "') " & sSQLC & " "
    sqlTIB = sqlTIB & "       ) T Group by T.cCtaAreaAsig Having SUM(Monto) > 0"

'    sqlTIB = " select cCtaAreaAsig,SUM(Monto) Monto "
'    sqlTIB = sqlTIB & " From "
'    sqlTIB = sqlTIB & "( "
'    sqlTIB = sqlTIB & " Select T.cCtaAreaAsig, sum(Monto) Monto"
'    sqlTIB = sqlTIB & " FROM ( "
'    sqlTIB = sqlTIB & " Select E.cCtaAreaAsig, "
'    sqlTIB = sqlTIB & "  PD.nMonto Monto FROM RHPlanillaDetCon PD"
'    sqlTIB = sqlTIB & "    INNER JOIN (Select cPersCod, cRHAgenciaCodOficial cAgenciaAsig, (Select Top 1 cSubCtaCod from areaagencia where cAreaCod = cRHAreaCodOficial And cUbicaCod = cRHAgenciaCodOficial) cCtaAreaAsig From RHCargos RHC Where "
'    sqlTIB = sqlTIB & "     dRHCargoFecha = (Select Max(RHC1.dRHCargoFecha) "
'    sqlTIB = sqlTIB & "     From RHCargos RHC1 Where RHC.cPersCod  = RHC1.cPersCod And Left(RHC1.cUltimaActualizacion,8) <= '" & psPlanillaReriodo & "')) As E ON PD.cPersCod = E.cPersCod"
'    sqlTIB = sqlTIB & "            INNER JOIN RHContrato RHC ON PD.cPersCod = RHC.cPersCod "
'    sqlTIB = sqlTIB & "            INNER JOIN RRHH RH on PD.cPersCod=RH.cPersCod "
'    sqlTIB = sqlTIB & "                 and (RH.cAgenciaActual<>'01' or (RH.cAgenciaActual='01' and rh.cAreaCodActual NOT IN (select cAreaCod from AreaPorcentajeGastos where cAgeCod='01'))) "
'    sqlTIB = sqlTIB & "  WHERE PD.cRHConceptoCod = '130' And PD.cPlanillaCod = '" & psPlanillaCod & "' "
'    sqlTIB = sqlTIB & "        And PD.cRRHHPeriodo = '" & psPlanillaReriodo & "'"
'    sqlTIB = sqlTIB & "        And RHC.cRHContratoNro In ( Select Max(RHC1.cRHContratoNro) From RHContrato RHC1 Where RHC.cPersCod = RHC1.cPersCod  And Left(RHC1.cUltimaActualizacion,8) <= '" & psPlanillaReriodo & "') " & sSQLC & " "
'    sqlTIB = sqlTIB & "       ) T Group by T.cCtaAreaAsig Having SUM(Monto) > 0"
'    sqlTIB = sqlTIB & " ) R "
'    sqlTIB = sqlTIB & " Group by R.cCtaAreaAsig "
'    sqlTIB = sqlTIB & " Order by R.cCtaAreaAsig "
    
    sqlTIB = sqlTIB & " union "

    sqlTIB = sqlTIB & " Select AgP.cAgeCod cCtaAreaAsig, (nAgePorcentaje)*SUM(Monto)/100 Monto FROM ("
    sqlTIB = sqlTIB & " Select E.cCtaAreaAsig,PD.nMonto Monto,RH.cAgenciaAsig"
    sqlTIB = sqlTIB & "    FROM RHPlanillaDetCon PD"
    sqlTIB = sqlTIB & "       INNER JOIN ("
    sqlTIB = sqlTIB & "                  Select cPersCod, cRHAgenciaCodOficial cAgenciaAsig, "
    sqlTIB = sqlTIB & "         (   Select Top 1 cSubCtaCod from areaagencia where cAreaCod = cRHAreaCodOficial And cUbicaCod = cRHAgenciaCodOficial) cCtaAreaAsig "
    sqlTIB = sqlTIB & "             From RHCargos RHC "
    sqlTIB = sqlTIB & "             Where dRHCargoFecha = ( Select Max(RHC1.dRHCargoFecha) "
    sqlTIB = sqlTIB & " From RHCargos RHC1"
    sqlTIB = sqlTIB & "                                                 Where RHC.cPersCod  = RHC1.cPersCod And Left(RHC1.cUltimaActualizacion,8) <= '" & psPlanillaReriodo & "' "
    sqlTIB = sqlTIB & "                                               )"
    sqlTIB = sqlTIB & "                    ) As E ON PD.cPersCod = E.cPersCod"
    sqlTIB = sqlTIB & "        INNER JOIN RHContrato RHC ON PD.cPersCod = RHC.cPersCod"
    sqlTIB = sqlTIB & "        INNER JOIN RRHH RH on PD.cPersCod=RH.cPersCod and RH.cAgenciaActual='01'"
    sqlTIB = sqlTIB & "        INNER JOIN AreaPorcentajeGastos APG ON RH.cAgenciaActual = APG.cAgeCod"
    sqlTIB = sqlTIB & "             and RH.cAreaCodActual = APG.cAreaCod"
    sqlTIB = sqlTIB & "    WHERE PD.cRHConceptoCod = '130' And PD.cPlanillaCod = '" & psPlanillaCod & "' And PD.cRRHHPeriodo = '" & psPlanillaReriodo & "'  And RHC.cRHContratoNro In ( Select Max(RHC1.cRHContratoNro) From RHContrato RHC1 Where RHC.cPersCod = RHC1.cPersCod  And Left(RHC1.cUltimaActualizacion,8) <= '" & psPlanillaReriodo & "') " & sSQLC & " ) T"
    sqlTIB = sqlTIB & "    inner JOIN AgenciaPorcentajeGastos AgP ON AgP.cAgeCodPrinc = T.cAgenciaAsig"
    sqlTIB = sqlTIB & "     Group by T.cCtaAreaAsig,cAgenciaAsig,AgP.cAgeCod,nAgePorcentaje Having SUM(Monto) > 0"
    sqlTIB = sqlTIB & " ) R "
    sqlTIB = sqlTIB & " Group by R.cCtaAreaAsig "
    sqlTIB = sqlTIB & " Order by R.cCtaAreaAsig "

    If oCon.AbreConexion Then
        'Valida creacion del asientocontable
        sqlP = "Select nMovNro From Mov Where cMovNro Like '" & psPlanillaReriodo & "%' And cOpeCod = '" & psOpe & "' And nMovFlag = 0"
        Set rsPla = oCon.CargaRecordSet(sqlP)
        
        If Not (rsPla.EOF And rsPla.BOF) Then
            MsgBox "La operacion ya genero un asiento contable, No puede volver a generarlo", vbInformation, "Aviso"
            If MsgBox("Desea imprimir el asiento grabado ?  ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                GeneraAsientoRemuneracion = rsPla!nMovNro
            Else
                GeneraAsientoRemuneracion = -1
            End If
            rsPla.Close
            Set rsPla = Nothing
            Exit Function
        Else
            rsPla.Close
        End If

        'Valida Que montos de Debe y Haber coincidan
        sqlP = " Select ISNull(SUM(PD.nMonto),0) Monto FROM RHConceptoCta CC" _
             & " INNER JOIN RHPlanillaDetCon PD ON CC.cRHConceptoCod = PD.cRHConceptoCod" _
             & " INNER JOIN RRHH E ON PD.cPersCod = E.cPersCod " _
             & " INNER JOIN RHContrato RHC ON PD.cPersCod = RHC.cPersCod " _
             & " WHERE PD.cPlanillaCod = '" & psPlanillaCod & "'" _
             & " And CC.cOpeCod = '" & psOpe & "' And PD.cRRHHPeriodo = '" & psPlanillaReriodo & "'" _
             & " And CC.cOpeCtaDH = 'D' And RHC.cRHContratoNro In ( Select Max(RHC1.cRHContratoNro) From RHContrato RHC1 Where RHC.cPersCod = RHC1.cPersCod) " _
             & sSQLC
        
        Set rsPla = oCon.CargaRecordSet(sqlP)
        lnDebe = 0
        If Not (rsPla.EOF And rsPla.BOF) Then lnDebe = rsPla("Monto")
        rsPla.Close
        
        sqlP = " Select ISNULL(SUM(PD.nMonto),0) Monto FROM RHConceptoCta CC" _
             & " INNER JOIN RHPlanillaDetCon PD ON CC.cRHConceptoCod = PD.cRHConceptoCod" _
             & " INNER JOIN RRHH E ON PD.cPersCod = E.cPersCod " _
             & " INNER JOIN RHContrato RHC ON PD.cPersCod = RHC.cPersCod " _
             & " WHERE PD.cPlanillaCod = '" & psPlanillaCod & "'" _
             & " And CC.cOpeCod = '" & psOpe & "' And PD.cRRHHPeriodo = '" & psPlanillaReriodo & "'" _
             & " And CC.cOpeCtaDH = 'H' And RHC.cRHContratoNro In ( Select Max(RHC1.cRHContratoNro) From RHContrato RHC1 Where RHC.cPersCod = RHC1.cPersCod) " _
             & sSQLC
        
        Set rsPla = oCon.CargaRecordSet(sqlP)
        lnHaber = 0
        If Not (rsPla.EOF And rsPla.BOF) Then lnHaber = rsPla("Monto")
        rsPla.Close
        
        If Abs(lnDebe - lnHaber) > 1 Then
            MsgBox "Error de Validación, el monto en Debe " & Format(lnDebe, gsFormatoNumeroView) & " es diferente al Haber " & Format(lnHaber, gsFormatoNumeroView) & " DIFF " & Format(lnDebe - lnHaber, gsFormatoNumeroView), vbExclamation, "Aviso"
            Set rsPla = Nothing
            oCon.CierraConexion
            Exit Function
        Else
            lnOtro = Round(lnDebe, 2) - Round(lnHaber, 2)
        End If
        
        sqlP = " Select cPlaInsDes,nTpoCambio, nTpoCambioFijo From RHPlanilla" _
             & " Where cRRHHPeriodo = '" & psPlanillaReriodo & "' And cPlanillaCod = '" & psPlanillaCod & "'"
        
        Set rsPla = oCon.CargaRecordSet(sqlP)
        
        lsGlosa = ""
        If Not (rsPla.EOF And rsPla.BOF) Then
            If psCond = "" Then
                lsGlosa = rsPla!cPlaInsDes
            ElseIf psCond = "0" Then
                lsGlosa = rsPla!cPlaInsDes & " -PLAZO INDETERMINADO- "
            ElseIf psCond = "1" Then
                lsGlosa = rsPla!cPlaInsDes & " -PLAZO FIJO- "
            End If
            
            lnTipCamV = rsPla!nTpoCambio
            lnTipCamFijo = rsPla!nTpoCambioFijo
        End If
        rsPla.Close
'   ALPA 20091229**********************************************************************
'        sqlP = " Select T.Cuenta, T.DebeHab, SUM(Monto) Monto" _
'             & " FROM ( " _
'             & "        Select CC.cRHConceptoCod Concepto, CC.cOpeCtaDH DebeHab, Cuenta = Replace(CC.cCtaContCod,'AG',E.cCtaAreaAsig), " _
'             & "        PD.nMonto Monto FROM RHConceptoCta CC" _
'             & "            INNER JOIN RHPlanillaDetCon PD ON CC.cRHConceptoCod = PD.cRHConceptoCod" _
'             & "            INNER JOIN (Select cPersCod, cRHAgenciaCodOficial cAgenciaAsig, (Select Top 1 cSubCtaCod from areaagencia where cAreaCod = cRHAreaCodOficial And cUbicaCod = cRHAgenciaCodOficial) cCtaAreaAsig From RHCargos RHC Where " _
'             & "                            dRHCargoFecha = (Select Max(RHC1.dRHCargoFecha) " _
'             & "                                                      From RHCargos RHC1 Where RHC.cPersCod  = RHC1.cPersCod And Left(RHC1.cUltimaActualizacion,8) <= '" & psPlanillaReriodo & "')) As E ON PD.cPersCod = E.cPersCod" _
'             & "            INNER JOIN RHContrato RHC ON PD.cPersCod = RHC.cPersCod " _
'             & "        WHERE PD.cPlanillaCod = '" & psPlanillaCod & "' " _
'             & "        And CC.cOpeCod = '" & psOpe & "' And PD.cRRHHPeriodo = '" & psPlanillaReriodo & "'" _
'             & "        And RHC.cRHContratoNro In ( Select Max(RHC1.cRHContratoNro) From RHContrato RHC1 Where RHC.cPersCod = RHC1.cPersCod  And Left(RHC1.cUltimaActualizacion,8) <= '" & psPlanillaReriodo & "') " & sSQLC & " " _
'             & "       ) T Group by T.Cuenta, T.DebeHab Having SUM(Monto) > 0" _
'             & " Order by T.DebeHab,T.Cuenta "


'*** PEAC 20111125
'        sqlP = " Select T.Cuenta, T.DebeHab, SUM(Monto) Monto"
'        sqlP = sqlP & " FROM ( "
'        sqlP = sqlP & "        Select CC.cRHConceptoCod Concepto, CC.cOpeCtaDH DebeHab, Cuenta = Replace(CC.cCtaContCod,'AG',RH.cAgenciaActual), "
'        sqlP = sqlP & "        PD.nMonto Monto FROM RHConceptoCta CC"
'        sqlP = sqlP & "            INNER JOIN RHPlanillaDetCon PD ON CC.cRHConceptoCod = PD.cRHConceptoCod"
'        sqlP = sqlP & "            INNER JOIN (Select cPersCod, cRHAgenciaCodOficial cAgenciaAsig, (Select Top 1 cSubCtaCod from areaagencia where cAreaCod = cRHAreaCodOficial And cUbicaCod = cRHAgenciaCodOficial) cCtaAreaAsig From RHCargos RHC Where "
'        sqlP = sqlP & "                            dRHCargoFecha = (Select Max(RHC1.dRHCargoFecha) "
'        sqlP = sqlP & "                                                      From RHCargos RHC1 Where RHC.cPersCod  = RHC1.cPersCod And Left(RHC1.cUltimaActualizacion,8) <= '" & psPlanillaReriodo & "')) As E ON PD.cPersCod = E.cPersCod"
'        sqlP = sqlP & "            INNER JOIN RHContrato RHC ON PD.cPersCod = RHC.cPersCod "
'        sqlP = sqlP & "            INNER JOIN RRHH RH on PD.cPersCod=RH.cPersCod"
'        sqlP = sqlP & "        WHERE PD.cPlanillaCod = '" & psPlanillaCod & "' "
'        sqlP = sqlP & "        And CC.cOpeCod = '" & psOpe & "' And PD.cRRHHPeriodo = '" & psPlanillaReriodo & "'"
'        sqlP = sqlP & "         and (RH.cAgenciaActual<>'01' or (RH.cAgenciaActual='01' and RH.cAreaCodActual not in (select cAreaCod from AreaPorcentajeGastos))) "
'        sqlP = sqlP & "        And RHC.cRHContratoNro In ( Select Max(RHC1.cRHContratoNro) From RHContrato RHC1 Where RHC.cPersCod = RHC1.cPersCod  And Left(RHC1.cUltimaActualizacion,8) <= '" & psPlanillaReriodo & "') " & sSQLC & " "
'        sqlP = sqlP & "       ) T Group by T.Cuenta, T.DebeHab Having SUM(Monto) > 0"
'        sqlP = sqlP & " Union "
'        sqlP = sqlP & " Select T.Cuenta, T.DebeHab, SUM(Monto) Monto"
'        sqlP = sqlP & " FROM ( "
'        sqlP = sqlP & "        Select CC.cRHConceptoCod Concepto, CC.cOpeCtaDH DebeHab, Cuenta = Replace(CC.cCtaContCod,'AG',AgP.cAgeCod), "
'        sqlP = sqlP & "            (PD.nMonto*nAgePorcentaje/100) Monto FROM RHConceptoCta CC"
'        sqlP = sqlP & "            INNER JOIN RHPlanillaDetCon PD ON CC.cRHConceptoCod = PD.cRHConceptoCod"
'        sqlP = sqlP & "            INNER JOIN (Select cPersCod, cRHAgenciaCodOficial cAgenciaAsig, (Select Top 1 cSubCtaCod from areaagencia where cAreaCod = cRHAreaCodOficial And cUbicaCod = cRHAgenciaCodOficial) cCtaAreaAsig From RHCargos RHC Where "
'        sqlP = sqlP & "                            dRHCargoFecha = (Select Max(RHC1.dRHCargoFecha) "
'        sqlP = sqlP & "                                                      From RHCargos RHC1 Where RHC.cPersCod  = RHC1.cPersCod And Left(RHC1.cUltimaActualizacion,8) <= '" & psPlanillaReriodo & "')) As E ON PD.cPersCod = E.cPersCod"
'        sqlP = sqlP & "            INNER JOIN RHContrato RHC ON PD.cPersCod = RHC.cPersCod "
'        sqlP = sqlP & "            INNER JOIN RRHH RH on PD.cPersCod=RH.cPersCod"
'        sqlP = sqlP & "            INNER JOIN AreaPorcentajeGastos APG ON RH.cAgenciaActual = APG.cAgeCod"
'        sqlP = sqlP & "                    and RH.cAreaCodActual = APG.cAreaCod"
'        sqlP = sqlP & "            INNER JOIN AgenciaPorcentajeGastos AgP ON  RH.cAgenciaActual=AgP.cAgeCodPrinc"
'        sqlP = sqlP & "        WHERE PD.cPlanillaCod = '" & psPlanillaReriodo & "' "
'        sqlP = sqlP & "             And CC.cOpeCod = '" & psOpe & "' And PD.cRRHHPeriodo = '" & psPlanillaReriodo & "'"
'        sqlP = sqlP & "             and RH.cAgenciaActual='01' "
'        sqlP = sqlP & "             And RHC.cRHContratoNro In (Select Max(RHC1.cRHContratoNro) From RHContrato RHC1 Where RHC.cPersCod = RHC1.cPersCod  And Left(RHC1.cUltimaActualizacion,8) <= '" & psPlanillaReriodo & "') " & sSQLC & " "
'        sqlP = sqlP & "       ) T Group by T.Cuenta, T.DebeHab Having SUM(Monto) > 0"
'        sqlP = sqlP & " Order by T.DebeHab,T.Cuenta "
        
        sqlP = "EXEC stp_sel_DistribuyeAsientoRemuneracion '" & Trim(psPlanillaReriodo) & "','" & Trim(psPlanillaCod) & "','" & Trim(psOpe) & "','" & Trim(psCond) & "' "
        
'*** FIN PEAC
        
'******************************************************************************************************
        Set rsPla = oCon.CargaRecordSet(sqlP)
        
        If Not (rsPla.EOF And rsPla.BOF) Then
            oMov.BeginTrans
            oMov.Inicio gsFormatoFecha
            lsGlosa = lsGlosa & " DE " & UCase(Format$(pdFecha, gsFormatoFechaView))
            lsMov = psPlanillaReriodo & Mid(oMov.GeneraMovNro(pdFecha, pgsCodAge, pgsCodUser), 9, 50)
            oMov.InsertaMov lsMov, psOpe, lsGlosa
            lnMov = oMov.GetnMovNro(lsMov)
            lnItem = 0
            
            Do While Not rsPla.EOF
                lnItem = lnItem + 1
                lnDH = IIf(rsPla("DebeHab") = "D", 1, -1)
                lnMonto = rsPla("Monto")
                lsCta = rsPla("Cuenta")
                oMov.InsertaMovCta lnMov, lnItem, lsCta, lnMonto * lnDH
                
                rsPla.MoveNext
            
                If lsCta = lsCtaPendRRHHSoles Then
                    If psOpe = gsRHPlanillaCTSRemEst Or psOpe = gsRHPlanillaCTSRemCon Then
                        sqlP = " Select Sum(nMonto) Monto_S, Round(Sum(nMonto) / " & lnTipCamV & ",2) Monto_D, Round(Round(Sum(nMonto) / " & lnTipCamV & ",2) * " & lnTipCamFijo & ",2) Monto_F " _
                             & " From RHCuentas EC" _
                             & " Inner Join RRHH E On EC.cPersCod = E.cPersCod" _
                             & " Inner Join RHPlanillaDetCon PD On EC.cPersCod = PD.cPersCod" _
                             & " Inner JOIN RHContrato RHC ON E.cPersCod = RHC.cPersCod" _
                             & " Where Substring(cCtaCod,3,4) = '2342' And cPlanillaCod = 'E05' And PD.cRHConceptoCod = '112' And cRRHHPeriodo Like '" & psPlanillaReriodo & "%'" _
                             & " And RHC.cRHContratoNro In ( Select Max(RHC1.cRHContratoNro) From RHContrato RHC1 Where RHC.cPersCod = RHC1.cPersCod)" _
                             & sSQLC
                        Set rsCred = oCon.CargaRecordSet(sqlP)
                        
                        If Not IsNull(rsCred!Monto_S) Then
                            lnItem = lnItem + 1
                            lsCta = lsCtaPendRRHHSoles
                            oMov.InsertaMovCta lnMov, lnItem, lsCta, rsCred!Monto_S
                            
                            lnItem = lnItem + 1
                            lsCta = lsCtaPendRRHHDolares
                            oMov.InsertaMovCta lnMov, lnItem, lsCta, rsCred!Monto_F * -1
                            
                            If rsCred!Monto_S - rsCred!Monto_F > 0 Then
                                lnItem = lnItem + 1
                                lsCta = lsCtaAjusteTipoCambioGanancia
                                oMov.InsertaMovCta lnMov, lnItem, lsCta, (rsCred!Monto_S - rsCred!Monto_F) * -1
                            
                            ElseIf rsCred!Monto_S - rsCred!Monto_F < 0 Then
                                lnItem = lnItem + 1
                                lsCta = lsCtaAjusteTipoCambioPerdida
                                oMov.InsertaMovCta lnMov, lnItem, lsCta, Abs((rsCred!Monto_S - rsCred!Monto_F))
                            End If
                        End If
                    Else
                        sqlP = " Select Round(Sum(convert(money,CRED.nMonto * " & lnTipCamFijo & ")),2) Monto_S, Round(Sum(CRED.nMonto),2) Monto_D, round(Sum(CRED.nMonto * nTipCambio),2) Monto_DV  , Round(Sum(CRED.nMonto * " & lnTipCamFijo & "),2) Monto_DF " _
                             & " From RHCreditosAdm CRED" _
                             & " Inner Join RRHH E ON CRED.cPersCod = E.cPersCod" _
                             & " Inner Join RHPlanillaDetCon EC ON CRED.cPersCod = EC.cPersCod And CRED.cPlaCod = EC.cPlanillaCod" _
                             & " Inner JOIN RHContrato RHC ON E.cPersCod = RHC.cPersCod" _
                             & " And Left(CRED.cPeriodo,6) = Left(EC.cRRHHPeriodo,6) Where CRED.cPlaCod = '" & psPlanillaCod & "' And Substring(cCtaCod,6,1) = '" & gcMEDig & "'" _
                             & " And EC.cRHConceptoCod in ('" & gRHConcDctoPrestAdmin & "') and cRRHHPeriodo Like '" & Left(psPlanillaReriodo, 6) & "%'" _
                             & " And RHC.cRHContratoNro In ( Select Max(RHC1.cRHContratoNro) From RHContrato RHC1 Where RHC.cPersCod = RHC1.cPersCod And Left(RHC1.cUltimaActualizacion,8) <= '" & psPlanillaReriodo & "') " & sSQLC
                        Set rsCred = oCon.CargaRecordSet(sqlP)
                        
                        If Not IsNull(rsCred!Monto_DV) Then
                            lnItem = lnItem + 1
                            lsCta = lsCtaPendRRHHSoles
                            oMov.InsertaMovCta lnMov, lnItem, lsCta, rsCred!Monto_DV + lnMontoHipoteca
                       
                            lnItem = lnItem + 1
                            lsCta = lsCtaPendRRHHDolares
                            oMov.InsertaMovCta lnMov, lnItem, lsCta, rsCred!Monto_DF * -1
                            
                            If rsCred!Monto_DV + lnMontoHipoteca - rsCred!Monto_DF > 0 Then
                                lnItem = lnItem + 1
                                lsCta = lsCtaAjusteTipoCambioGanancia
                                oMov.InsertaMovCta lnMov, lnItem, lsCta, (rsCred!Monto_DV + lnMontoHipoteca - rsCred!Monto_DF) * -1

                            ElseIf rsCred!Monto_DV + lnMontoHipoteca - rsCred!Monto_DF < 0 Then
                                lnItem = lnItem + 1
                                lsCta = lsCtaAjusteTipoCambioPerdida
                                oMov.InsertaMovCta lnMov, lnItem, lsCta, Abs(rsCred!Monto_DV + lnMontoHipoteca - rsCred!Monto_DF)
                            End If
                        End If
                    End If
                
                ElseIf lsCta = lsCtaPendRRHHCredOtrosSoles Then 'Otros Creditos
                     sqlP = " Select Round(Sum(convert(money,CRED.nMonto * " & lnTipCamFijo & ")),2) Monto_S, Round(Sum(CRED.nMonto),2) Monto_D, round(Sum(CRED.nMonto * nTipCambio),2) Monto_DV  , Round(Sum(CRED.nMonto * " & lnTipCamFijo & "),2) Monto_DF " _
                          & " From RHCreditosAdmOtros CRED" _
                          & " Inner Join RRHH E ON CRED.cPersCod = E.cPersCod" _
                          & " Inner Join RHPlanillaDetCon EC ON CRED.cPersCod = EC.cPersCod And CRED.cPlaCod = EC.cPlanillaCod" _
                          & " Inner JOIN RHContrato RHC ON E.cPersCod = RHC.cPersCod" _
                          & " And Left(CRED.cPeriodo,8) = Left(EC.cRRHHPeriodo,8) Where CRED.cPlaCod = '" & psPlanillaCod & "' And nMoneda = '" & gcMEDig & "'" _
                          & " And EC.cRHConceptoCod in ('" & gRHConcDctoPrestAdmin & "') and cRRHHPeriodo Like '" & Left(psPlanillaReriodo, 8) & "%'" _
                          & " And RHC.cRHContratoNro In ( Select Max(RHC1.cRHContratoNro) From RHContrato RHC1 Where RHC.cPersCod = RHC1.cPersCod And Left(RHC1.cUltimaActualizacion,8) <= '" & psPlanillaReriodo & "') " & sSQLC
                     Set rsCred = oCon.CargaRecordSet(sqlP)
                     
                     If Not IsNull(rsCred!Monto_DV) Then
                         lnItem = lnItem + 1
                         lsCta = lsCtaPendRRHHCredOtrosSoles
                         oMov.InsertaMovCta lnMov, lnItem, lsCta, rsCred!Monto_DV + lnMontoHipoteca
                    
                         lnItem = lnItem + 1
                         lsCta = lsCtaPendRRHHCredOtrosDolares
                         oMov.InsertaMovCta lnMov, lnItem, lsCta, rsCred!Monto_DF * -1
                         
                         If rsCred!Monto_DV + lnMontoHipoteca - rsCred!Monto_DF > 0 Then
                             lnItem = lnItem + 1
                             lsCta = lsCtaAjusteTipoCambioGanancia
                             oMov.InsertaMovCta lnMov, lnItem, lsCta, (rsCred!Monto_DV + lnMontoHipoteca - rsCred!Monto_DF) * -1

                         ElseIf rsCred!Monto_DV + lnMontoHipoteca - rsCred!Monto_DF < 0 Then
                             lnItem = lnItem + 1
                             lsCta = lsCtaAjusteTipoCambioPerdida
                             oMov.InsertaMovCta lnMov, lnItem, lsCta, Abs(rsCred!Monto_DV + lnMontoHipoteca - rsCred!Monto_DF)
                         End If
                     End If
                ElseIf lsCta = "251504" And psPlanillaCod = "E05" Then
                     sqlP = " Select Sum(PD.nMonto) Soles, Round(Sum(PD.nMonto) / " & lnTipCamV & ",2) Dolares, Round((Sum(PD.nMonto) / " & lnTipCamV & ") * " & lnTipCamFijo & ",2) Fijo , Substring(CC.cCtaCod,6,1) Moneda" _
                          & " FROM RHCuentas CC" _
                          & " INNER JOIN RHPlanillaDetCon PD ON CC.cPersCod = PD.cPersCod" _
                          & " INNER JOIN (Select cPersCod, cRHAgenciaCodOficial cAgenciaAsig From RHCargos RHC" _
                          & "   Where  dRHCargoFecha = (Select Max(RHC1.dRHCargoFecha) From RHCargos RHC1" _
                          & "                           Where RHC.cPersCod  = RHC1.cPersCod And" _
                          & "                Left(RHC1.cUltimaActualizacion,8) <= '" & psPlanillaReriodo & "')) As E ON PD.cPersCod = E.cPersCod" _
                          & " INNER JOIN RHContrato RHC ON PD.cPersCod = RHC.cPersCod" _
                          & " WHERE PD.cPlanillaCod = '" & psPlanillaCod & "' And PD.cRRHHPeriodo = '" & psPlanillaReriodo & "'" _
                          & "    And RHC.cRHContratoNro In (Select Max(RHC1.cRHContratoNro)" _
                          & "                               From RHContrato RHC1 Where RHC.cPersCod = RHC1.cPersCod" _
                          & "                               And Left(RHC1.cUltimaActualizacion,8) <= '" & psPlanillaReriodo & "')" _
                          & "         And Substring(CC.cCtaCod,3,3) = '234' " & sSQLC & " And PD.cRHConceptoCod = '130'" _
                          & "   Group By Substring(CC.cCtaCod,6,1) having Substring(CC.cCtaCod,6,1) = '2'"
                     Set rsCred = oCon.CargaRecordSet(sqlP)
                    
                     If Not rsCred.EOF And Not rsCred.BOF Then
                         lnItem = lnItem + 1
                         lsCta = "251504"
                         oMov.InsertaMovCta lnMov, lnItem, lsCta, rsCred!Soles
                         
                         lnItem = lnItem + 1
                         lsCta = "252504"
                         oMov.InsertaMovCta lnMov, lnItem, lsCta, rsCred!Fijo * -1
                         
                         If rsCred!Soles - rsCred!Fijo > 0 Then
                             lnItem = lnItem + 1
                             lsCta = lsCtaAjusteTipoCambioGanancia
                             oMov.InsertaMovCta lnMov, lnItem, lsCta, (rsCred!Soles - rsCred!Fijo) * -1
    
                         ElseIf rsCred!Soles - rsCred!Fijo < 0 Then
                             lnItem = lnItem + 1
                             lsCta = lsCtaAjusteTipoCambioPerdida
                             oMov.InsertaMovCta lnMov, lnItem, lsCta, (rsCred!Soles - rsCred!Fijo) * -1
                         End If
                     End If
                End If
            Loop
            
            If lnMontoDebe > lnMontoHaber Then
                oMov.ActualizaMovCta lnMov, 1, , 5
            ElseIf lnMontoHaber > lnMontoDebe Then
                oMov.ActualizaMovCta lnMov, 1, , 5
            End If
        
            If lnOtro <> 0 Then
                lnItem = lnItem + 1
                If lnOtro < 0 Then
                    'Otros Ingresos
                    lsCta = lsCtaOtrosIngresos
                    oMov.InsertaMovCta lnMov, lnItem, lsCta, lnOtro * -1
                Else
                    'Otros Egresos
                    lsCta = lsCtaOtrosEgresos
                    oMov.InsertaMovCta lnMov, lnItem, lsCta, lnOtro * -1
                End If
            End If
        
        If psPlanillaCod <> "P02" And psPlanillaCod <> "E21" Then
            Set rsTIB = oCon.CargaRecordSet(sqlTIB)
                        
            While Not rsTIB.EOF
                lnItem = lnItem + 1
                lsCta = Replace("45140901AG", "AG", rsTIB!cCtaAreaAsig)
                oMov.InsertaMovCta lnMov, lnItem, lsCta, Format(rsTIB!Monto * 0.0007, "0.00")

                lnItem = lnItem + 1
                lsCta = "21140207"
                oMov.InsertaMovCta lnMov, lnItem, lsCta, Format(rsTIB!Monto * 0.0007, "0.00") * -1

                rsTIB.MoveNext
            Wend
        End If
        
        '*** PEAC 20111102 - verifica descuadres en el asiento
        Set rsAsiento = oMov.CargaRecordSet("select nMovItem,cCtaContCod,case when nMovImporte > 0 then nMovImporte else 0 end ndebe,case when nMovImporte < 0 then nMovImporte*-1 else 0 end nhaber from MovCta where nMovNro=" & lnMov)
        
        lnMontoDebe = 0: lnMontoHaber = 0
        lnItenDebe = 0: lnItenHaber = 0: lnImporteDebe = 0: lnImporteHaber = 0
        Do While Not rsAsiento.EOF
            If rsAsiento!nDebe > 0 And lnItenDebe = 0 And Left(rsAsiento!cCtaContCod, 2) = "45" Then
                lnItenDebe = rsAsiento!nMovItem
                lnImporteDebe = rsAsiento!nDebe
                
            ElseIf rsAsiento!nHaber > 0 And lnItenHaber = 0 And Left(rsAsiento!cCtaContCod, 2) = "45" Then
                lnItenHaber = rsAsiento!nMovItem
                lnImporteHaber = rsAsiento!nHaber
            End If
            
            lnMontoDebe = lnMontoDebe + rsAsiento!nDebe
            lnMontoHaber = lnMontoHaber + rsAsiento!nHaber
            rsAsiento.MoveNext
        Loop
        
        If lnMontoDebe > lnMontoHaber Then
            oMov.ActualizaMovCta lnMov, lnItenDebe, , lnImporteDebe - (lnMontoDebe - lnMontoHaber)
        ElseIf lnMontoHaber > lnMontoDebe Then
            oMov.ActualizaMovCta lnMov, lnItenDebe, , lnImporteDebe + (lnMontoHaber - lnMontoDebe)
        End If
        '*** FIN PEAC
        
        rsPla.Close
        Set rsPla = Nothing
        
        oMov.CommitTrans
        oCon.CierraConexion
        End If
    End If
    GeneraAsientoRemuneracion = lnMov
End Function

Public Function GetMesTrabCTS(ByVal psPersCod As String, ByVal pdFIni As Date, ByVal pdFFin As Date) As Single
    Dim rsDias As New ADODB.Recordset
    Set rsDias = New ADODB.Recordset
    Dim nMes As Single, nDias As Single
    Dim lsFIni As String, lsFFin As String
    Dim ldIng As Date, ldCese As Date
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    'Determinar el #Meses que el trabajador ha estado en la empresa
    'basado en su fecha de ingreso y si cumple el periodo para el calculo
    sqlC = "Select dIngreso , dCese  from RRHH RH Where cPersCod = '" & psPersCod & "'"
    
    If oCon.AbreConexion Then
        Set rsDias = oCon.CargaRecordSet(sqlC)
        If rsDias.EOF And rsDias.BOF Then
            GetMesTrabCTS = 0
            rsDias.Close
            Set rsDias = Nothing
            Set oCon = Nothing
            Exit Function
        Else
            ldIng = CDate(Format$(rsDias!dIngreso, gsFormatoFechaView))
            If Not IsNull(rsDias!dCese) Then ldCese = CDate(Format$(rsDias!dCese, gsFormatoFechaView))
            
            If ldIng > pdFIni And IsNull(rsDias!dCese) Then
                nMes = GetMeses(ldIng, pdFFin)
            Else
                nMes = GetMeses(pdFIni, pdFFin)
            End If
        End If
        
        'Se deben calcular los dias que no trabajo por inasistencias
        'injustificadas o licencias sin goce de haber para el rango establecido
        lsFIni = Format$(pdFIni, gsFormatoFecha)
        lsFFin = Format$(pdFFin, gsFormatoFecha)
        
        sqlC = " Select SUM(DateDiff(dd,dIni,dFin) + 1) nDias" _
             & " FROM ( " _
             & "        Select dIni = CASE WHEN dEjecutadoInicio < '" & lsFIni & "' THEN '" & lsFIni & "' ELSE dEjecutadoInicio END, " _
             & "               dFin = CASE WHEN dEjecutadoFin > '" & lsFFin & "' THEN '" & lsFFin & "' ELSE dEjecutadoFin END " _
             & " From RHPeriodosnoLaborados " _
             & " WHERE (dEjecutadoInicio BETWEEN '" & lsFIni & "' AND '" & lsFFin & "' OR " _
             & " dEjecutadoFin Between '" & lsFIni & "' AND '" & lsFFin & "' OR " _
             & " (dEjecutadoInicio < '" & lsFIni & "' AND dEjecutadoFin > '" & lsFFin & "')) " _
             & " AND cPersCod = '" & psPersCod & "' And (nRHPeriodoTpo IN (" & RHEstadosRRHH.gRHEstadosRRHHLICSINGOCE & ") Or LEft(nRHPeriodoTpo,1) = " & RHEstadosTpo.RHEstadosTpoSuspendido & ")) T "
        
        Set rsDias = oCon.CargaRecordSet(sqlC)
        
        If rsDias.EOF And rsDias.BOF Then
            nDias = 0
        Else
            nDias = IIf(IsNull(rsDias!nDias), 0, rsDias!nDias)
        End If
        rsDias.Close
        Set rsDias = Nothing
        nMes = nMes - nDias / 30
        If nMes < 1 Then nMes = 0
        oCon.CierraConexion
    End If
    GetMesTrabCTS = Round(nMes, 2)
End Function

Public Function GetProvVac(ByVal psPersCod As String, ByVal pdFecha As Date, ByVal psPeriodoProc As String) As Double
    Dim rsProv As ADODB.Recordset
    Set rsProv = New ADODB.Recordset
    Dim sValor As String
    Dim nDias As Single
    Dim nTotRem As Double
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oInterprete As DInterprete
    Set oInterprete = New DInterprete
    
    sqlC = " Select nRHEmplVacacionesPend - (Select DateDiff(Day,dEjecutadoInicio,dEjecutadoFin) From RHPeriodosNoLaborados Where cPersCod = '" & psPersCod & "' And nRHPeriodoTpo like '" & RHEstadosTpo.RHEstadosTpoVacaciones & "%')" _
         & " From RHEmpleado" _
         & " Where cPersCod = '" & psPersCod & "'"
    If oCon.AbreConexion Then
        Set rsProv = oCon.CargaRecordSet(sqlC)
        If rsProv.EOF And rsProv.BOF Then
            nDias = 0
        Else
            nDias = IIf(IsNull(rsProv.Fields(0)), 0, rsProv.Fields(0))
        End If
        rsProv.Close
        Set rsProv = Nothing
        oInterprete.Interprete_InI
        sValor = oInterprete.FunEvalua(gsRHConceptoVTOTREM, psPersCod, pdFecha, pdFecha, False, "", psPeriodoProc)
        nTotRem = ExprANum(sValor)
        oCon.CierraConexion
    End If
    GetProvVac = Round(nDias * nTotRem / gnRHNumDiasVac, 2)
End Function

Public Function GetProvCTS(ByVal psCodPers As String, ByVal pdFecha As Date, ByVal psPeriodoProc As String) As Double
    Dim rsProv As ADODB.Recordset
    Set rsProv = New ADODB.Recordset
    Dim dIni As Date, dFin As Date
    Dim sValor As String
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oInterprete As DInterprete
    Set oInterprete = New DInterprete
    
    sqlC = " Select Convert(varchar(6),dReferenciaIni,101) + '" & Year(pdFecha) & "' dReferenciaIni, Convert(varchar(6),dReferenciaFin,101) + '" & Year(pdFecha) & "' dReferenciaFin From RHPlanillaFechas Where '" & Format(pdFecha, gsFormatoFecha) & "' " _
         & " Between Convert(varchar(6),dReferenciaIni,101) + '" & Year(pdFecha) & "' And Convert(varchar(6),dReferenciaFin,101) + '" & Year(pdFecha) & "' And cPlanillaCod = '" & gsRHPlanillaCTS & "'"
    
    If oCon.AbreConexion Then
        Set rsProv = oCon.CargaRecordSet(sqlC)
        
        If rsProv.EOF And rsProv.BOF Then
            dIni = CDate("01" & "/" & Month(pdFecha) & "/" & Year(pdFecha))
            dFin = DateAdd("d", -1, DateAdd("m", 1, dIni))
        Else
            dIni = CDate(Format$(rsProv!dReferenciaIni, gsFormatoFechaView))
            dFin = pdFecha
        End If
        rsProv.Close
        Set rsProv = Nothing
        oInterprete.Interprete_InI
        sValor = oInterprete.FunEvalua("I_TOT_CTS", psCodPers, dIni, dFin, False, "", psPeriodoProc)
        oCon.CierraConexion
    End If
    GetProvCTS = Round(ExprANum(sValor), 2)
End Function

Public Function GetProvTercio(ByVal psCodPers As String, ByVal pdFecha As Date, ByVal psPeriodoProc As String) As Double
    Dim rsProv As ADODB.Recordset
    Set rsProv = New ADODB.Recordset
    Dim dIni As Date, dFin As Date
    Dim sValor As String
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oInterprete As DInterprete
    Set oInterprete = New DInterprete
    
    sqlC = " Select Convert(varchar(6),dReferenciaIni,101) + '" & Year(pdFecha) & "' dReferenciaIni, Convert(varchar(6),dReferenciaFin,101) + '" & Year(pdFecha) & "' dReferenciaFin from RHPlanillaFechas Where '" & Format(pdFecha, gsFormatoFecha) & "' " _
         & " Between Convert(varchar(6),dReferenciaIni,101) + '" & Year(pdFecha) & "' And Convert(varchar(6),dReferenciaFin,101) + '" & Year(pdFecha) & "' And cPlanillaCod = '" & gsRHPlanillaTercio & "'"
    
    If oCon.AbreConexion Then
        Set rsProv = oCon.CargaRecordSet(sqlC)
        
        If rsProv.EOF And rsProv.BOF Then
            dIni = CDate("01" & "/" & Month(pdFecha) & "/" & Year(pdFecha))
            dFin = DateAdd("d", -1, DateAdd("m", 1, dIni))
        Else
            dIni = CDate(Format$(rsProv!dReferenciaIni, gsFormatoFechaView))
            dFin = pdFecha
        End If
        rsProv.Close
        Set rsProv = Nothing
        oInterprete.Interprete_InI
        sValor = oInterprete.FunEvalua("I_TERCIO", psCodPers, dIni, dFin, False, "", psPeriodoProc)
        oCon.CierraConexion
    End If
    GetProvTercio = Round(ExprANum(sValor) / 4, 2)
End Function

Public Function GetProvBonifVac(ByVal psCodPers As String, ByVal pdFecha As Date, ByVal psPeriodoProc As String) As Double
    Dim rsProv As ADODB.Recordset
    Set rsProv = New ADODB.Recordset
    Dim sValor As String
    Dim nDias As Single
    Dim nTotRem As Double
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oInterprete As DInterprete
    Set oInterprete = New DInterprete
    
    sqlC = " Select nRHEmplVacacionesPend - (Select DateDiff(Day,dEjecutadoInicio,dEjecutadoFin) From RHPeriodosNoLaborados Where cPersCod = '" & psCodPers & "' And nRHPeriodoTpo like '" & RHEstadosTpo.RHEstadosTpoVacaciones & "%')" _
         & " From RHEmpleado" _
         & " Where cPersCod = '" & psCodPers & "'"
    
    If oCon.AbreConexion Then
        Set rsProv = oCon.CargaRecordSet(sqlC)
        If rsProv.EOF And rsProv.BOF Then
            nDias = 0
        Else
            nDias = IIf(IsNull(rsProv.Fields(0)), 0, rsProv.Fields(0))
        End If
        rsProv.Close
        Set rsProv = Nothing
        oInterprete.Interprete_InI
        sValor = oInterprete.FunEvalua(gsRHConceptoVTOTREM, psCodPers, pdFecha, pdFecha, False, "", psPeriodoProc)
        nTotRem = ExprANum(sValor)
    End If
    GetProvBonifVac = Round(nDias * nTotRem / gnRHNumDiasVac, 2)
End Function

Public Function GetProvGrat(ByVal psCodPers As String, ByVal pdFecha As Date) As Double
    Dim rsProv As ADODB.Recordset
    Set rsProv = New ADODB.Recordset
    Dim dIni As Date, dFin As Date
    Dim sValor As String
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oInterprete As DInterprete
    Set oInterprete = New DInterprete
    
    sqlC = " Select Convert(varchar(6),dReferenciaIni,101) + '" & Year(pdFecha) & "' dReferenciaIni, Convert(varchar(6),dReferenciaFin,101) + '" & Year(pdFecha) & "' dReferenciaFin from RHPlanillaFechas Where '" & Format(pdFecha, gsFormatoFecha) & "' " _
         & " Between Convert(varchar(6),dReferenciaIni,101) + '" & Year(pdFecha) & "' And Convert(varchar(6),dReferenciaFin,101) + '" & Year(pdFecha) & "' And cPlanillaCod = '" & gsRHPlanillaGratificacion & "'"
    
    If oCon.AbreConexion Then
        Set rsProv = oCon.CargaRecordSet(sqlC)
        If rsProv.EOF And rsProv.BOF Then
            dIni = CDate("01" & "/" & Month(pdFecha) & "/" & Year(pdFecha))
            dFin = DateAdd("d", -1, DateAdd("m", 1, dIni))
        Else
            dIni = CDate(Format$(rsProv!dReferenciaIni, gsFormatoFechaView))
            dFin = pdFecha
        End If
        rsProv.Close
        Set rsProv = Nothing
        oInterprete.Interprete_InI
        sValor = oInterprete.FunEvalua("I_TOTAL_GRATIF", psCodPers, dIni, dFin, False, "", "")
    End If
    GetProvGrat = Round(ExprANum(sValor), 2)
End Function

Public Function GrabaAsientoProvision(ByVal psPeriodo As String, ByVal pdFecha As Date, pgsCodAge As String, pgsCodUser As String, ByVal psCond As String, psOpeCod As String) As String
    Dim rsPla As ADODB.Recordset
    Dim lsMov As String, lsGlosa As String, lsCta As String
    Set rsPla = New ADODB.Recordset
    Dim lnDebe As Double, lnHaber As Double, lnOtro As Double
    Dim lnDH As Double, lnMonto As Double
    Dim lnItem As Integer
    Dim lsSQLC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lsqlP As String
    Dim sqlP As String
    Dim lnMov As Long
    Dim oMov As DMov
    Set oMov = New DMov

    
    If oCon.AbreConexion Then
        sqlP = " Select cMovNro From Mov Where cOpeCod  = '" & psOpeCod & "' And cMovNro Like '" & psPeriodo & "%' And nMovEstado = " & MovEstado.gMovEstContabMovContable & " And nMovFlag = " & MovFlag.gMovFlagVigente & ""
        Set rsPla = oCon.CargaRecordSet(sqlP)
        lnDebe = 0
        If Not (rsPla.EOF And rsPla.BOF) Then
            If MsgBox("El asiento ya fue generado, ¿ Desea re - imprimirlo ? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                GrabaAsientoProvision = rsPla.Fields(0)
            Else
                GrabaAsientoProvision = "-1"
            End If
            Exit Function
        End If
        rsPla.Close

        'Valida Que montos de Debe y Haber coincidan
        sqlP = " Select SUM(PD.nMonto) Monto FROM RHConceptoCtaProvision CC" _
             & " INNER JOIN RHPlanillaDetConProvision PD ON CC.cRHConceptoCod = PD.cRHConceptoCod" _
             & " INNER JOIN RRHH E ON PD.cPersCod = E.cPersCod " _
             & " INNER JOIN RHContrato RHC ON PD.cPersCod = RHC.cPersCod " _
             & " WHERE PD.cRRHHPeriodo = '" & psPeriodo & "' And CC.cOpeCod = '" & psOpeCod & "' " _
             & " And CC.cOpeCtaDH = 'D' And RHC.cRHContratoNro In ( Select Max(RHC1.cRHContratoNro) From RHContrato RHC1 Where RHC.cPersCod = RHC1.cPersCod) " _
             & " And nRHContratoTpo = " & psCond
            
        Set rsPla = oCon.CargaRecordSet(sqlP)
        lnDebe = 0
        If Not (rsPla.EOF And rsPla.BOF) Then lnDebe = rsPla("Monto")
        rsPla.Close
        
        sqlP = " Select SUM(PD.nMonto) Monto FROM RHConceptoCtaProvision CC" _
             & " INNER JOIN RHPlanillaDetConProvision PD ON CC.cRHConceptoCod = PD.cRHConceptoCod" _
             & " INNER JOIN RRHH E ON PD.cPersCod = E.cPersCod " _
             & " INNER JOIN RHContrato RHC ON PD.cPersCod = RHC.cPersCod " _
             & " WHERE PD.cRRHHPeriodo = '" & psPeriodo & "' And CC.cOpeCod = '" & psOpeCod & "' " _
             & " And CC.cOpeCtaDH = 'H' And RHC.cRHContratoNro In ( Select Max(RHC1.cRHContratoNro) From RHContrato RHC1 Where RHC.cPersCod = RHC1.cPersCod) " _
             & " And nRHContratoTpo = " & psCond
        
        Set rsPla = oCon.CargaRecordSet(sqlP)
        lnHaber = 0
        If Not (rsPla.EOF And rsPla.BOF) Then lnHaber = rsPla("Monto")
        rsPla.Close
        
        If Abs(lnDebe - lnHaber) > 1 Then
            MsgBox "Error de Validación, el monto en Debe es diferente al Haber", vbExclamation, adCmdText
            Set rsPla = Nothing
            oCon.CierraConexion
            Exit Function
        Else
            lnOtro = lnDebe - lnHaber
        End If
        
        lsGlosa = "PROVISION " & IIf(psCond = RHContratoTipo.RHContratoTipoIndeterminado, "ESTABLA", "CONTRATO") & " -  PERIODO " & psPeriodo
        
        sqlP = " Select T.Cuenta, T.DebeHab, SUM(Monto) Monto" _
             & " FROM ( " _
             & "        Select CC.cRHConceptoCod Concepto, CC.cOpeCtaDH DebeHab, Cuenta = Replace(CC.cCtaContCod,'AG',E.cAgenciaAsig), " _
             & "        PD.nMonto Monto FROM RHConceptoCtaProvision CC" _
             & "            INNER JOIN RHPlanillaDetConProvision PD ON CC.cRHConceptoCod = PD.cRHConceptoCod" _
             & "            INNER JOIN RRHH E ON PD.cPersCod = E.cPersCod" _
             & "            INNER JOIN RHContrato RHC ON PD.cPersCod = RHC.cPersCod " _
             & "        WHERE PD.cRRHHPeriodo = '" & psPeriodo & "'  And CC.cOpeCod = '" & psOpeCod & "' And nRHContratoTpo = '" & psCond & "'" _
             & "        And RHC.cRHContratoNro In ( Select Max(RHC1.cRHContratoNro) From RHContrato RHC1 Where RHC.cPersCod = RHC1.cPersCod) And PD.nMonto <> 0" _
             & "       ) T Group by T.Cuenta, T.DebeHab " _
             & " Order by T.DebeHab,T.Cuenta "
        
        Set rsPla = oCon.CargaRecordSet(sqlP)
        
        If Not (rsPla.EOF And rsPla.BOF) Then
            oMov.Inicio gsFormatoFecha
            lsGlosa = lsGlosa & " DE " & UCase(Format$(pdFecha, gsFormatoFechaView))
            lsMov = oMov.GeneraMovNro(pdFecha, pgsCodAge, pgsCodUser)
            oMov.InsertaMov lsMov, psOpeCod, lsGlosa
            lnMov = oMov.GetnMovNro(lsMov)
            lnItem = 0
            
            Do While Not rsPla.EOF
                lnItem = lnItem + 1
                lnDH = IIf(rsPla("DebeHab") = "D", 1, -1)
                lnMonto = rsPla("Monto")
                lsCta = rsPla("Cuenta")
                oMov.InsertaMovCta lnMov, lnItem, lsCta, rsPla!Monto * lnDH
                rsPla.MoveNext
            Loop
        End If
        
        rsPla.Close
        Set rsPla = Nothing
        oCon.CierraConexion
    End If
    
    GrabaAsientoProvision = lsMov
End Function

Public Sub IncrementoVac(ByVal pnIncremento As Double)
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlC = " Update RHEmpleado " _
         & " Set nRHEmplVacacionesPend = nRHEmplVacacionesPend + " & pnIncremento & "" _
         & " Where cPersCod In" _
         & " ( Select RH.cPersCod from RRHH RH" _
         & " Inner Join RHEmpleado RHE On RH.cPersCod = RHE.cPersCod" _
         & " where nRHEstado not like '" & RHEstadosTpo.RHEstadosTpoRetirado & "%')" _
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Sub

Public Function CierreMesListaRRHH() As ADODB.Recordset
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    'Lista todos los empleados activos, no vacaciones, no retirado, será cierto eso...!!!!
    sqlC = " Select E.cRHCod, E.cPersCod, nRHEmplVacacionesPend, nRHContratoTpo TipoContrato, " _
         & " (Select cRHPlanillaCod From RHConcepto EC Where EC.cPersCod = E.cPErsCod" _
         & " And EC.cRHPlanillaCod = '" & gsRHPlanillaBonificacionVacacinal & "' And cRHConceptoCod = '" & gsRHConceptoITOTINGCOD & "') cBonVac" _
         & "      FROM RRHH E" _
         & " Inner Join RHEmpleado RHE On E.cPersCod = RHE.cPersCod" _
         & " Inner Join RHContrato CO On E.cPersCod = CO.cPersCod" _
         & " Where nRHEstado Not like '[78]%' And E.cRHCod Like 'E%'" _
         & "       And cRHContratoNro in (Select Max(cRHContratoNro) " _
         & "           From rhContrato CO1 Where CO.cPersCod = CO1.cPersCod)"
    
    If oCon.AbreConexion Then
        Set CierreMesListaRRHH = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function

Public Function SetPlanillaProvision(ByVal psRRHHPeriodo As String, ByVal psPersCod As String, ByVal psRHConceptoCod As String, ByVal psMonto As String) As Boolean
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlC = " Insert RHPlanillaDetConProvision (cRRHHPeriodo,cPersCod,cRHConceptoCod,nMonto)" _
         & " Values ('" & psRRHHPeriodo & "','" & psPersCod & "','" & psRHConceptoCod & "','" & psMonto & "')"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function

Public Function VerificaPlanillaProvision(ByVal psRRHHPeriodo As String) As Boolean
    Dim sqlC As String
    Dim rsPla As ADODB.Recordset
    Set rsPla = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlC = " Select Top 1 cRRHHPeriodo From RHPlanillaDetConProvision Where cRRHHPeriodo = '" & psRRHHPeriodo & "'"
    
    If oCon.AbreConexion Then
        Set rsPla = oCon.CargaRecordSet(sqlC)
        If rsPla.EOF And rsPla.BOF Then
            VerificaPlanillaProvision = False
        Else
            VerificaPlanillaProvision = True
        End If
        
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function


#If DebugMode Then
    '##ModelId=3AB90301024D
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

'Public Sub GrabaAsientoProvision(ByVal psCodOpe As String, ByVal dFecha As Date, ByVal nMonto As Double, pgsCodAge As String, pgsCodUser As String)
'    Dim rsProv As ADODB.Recordset
'    Set rsProv = New ADODB.Recordset
'    Dim lsMov As String, lsGlosa As String, sOpe As String
'    Dim lsFecha As String, sCta As String
'    Dim nDebeHab As Double
'    Dim lnItem As Integer
'    Dim sqlC As String
'    Dim oCon As DConecta
'    Set oCon = New DConecta
'    Dim oInterprete As DInterprete
'    Set oInterprete = New DInterprete
'    Dim oMov As DMov
'    Set oMov = New DMov
'    Dim lnMovNro As Long
'
'    sqlC = " Select OT.cOpeDesc, OC.cCtaContCod, OC.cOpeCtaDH" _
'         & " From OpeCuenta OC" _
'         & " INNER JOIN OpeTpo OT ON OC.cOpeCod = OT.cOpeCod" _
'         & " Where OC.cOpeCod = '" & psCodOpe & "' Order by cOpeCtaDH"
'
'    If oCon.AbreConexion Then
'        Set rsProv = oCon.CargaRecordSet(sqlC)
'        If Not (rsProv.EOF And rsProv.BOF) Then
'            lsFecha = Format$(dFecha, gsFormatoFechaView)
'            lsGlosa = Trim(rsProv("cOpeDesc")) & " Al " & lsFecha
'            lsMov = oMov.GeneraMovNro(dFecha, pgsCodAge, pgsCodUser)
'            oMov.InsertaMov lsMov, psCodOpe, lsGlosa
'            lnMovNro = oMov.GetnMovNro(lsMov)
'
'            lnItem = 0
'            Do While Not rsProv.EOF
'                lnItem = lnItem + 1
'                sCta = rsProv("cCtaContCod")
'                nDebeHab = IIf(rsProv("cOpeCtaDH") = "D", 1, -1)
'                oMov.InsertaMovCta lnMovNro, lnItem, sCta, nMonto * nDebeHab
'                rsProv.MoveNext
'            Loop
'        Else
'            MsgBox "Operacion no definida contablemente.Consulte con el Area de Sistemas", vbExclamation, "Error"
'        End If
'        rsProv.Close
'        Set rsProv = Nothing
'    End If
'End Sub

Public Sub ActuializaPlanillaMovs(psPeriodo As String, psPlaCod As String, pnMovEst As Long, pnMovCont As Long)
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sql = " Update RHPlanilla" _
        & " Set nMovNroEst = " & pnMovEst & ", nMovNroCont = " & pnMovCont & "" _
        & " Where cRRHHPeriodo = '" & psPeriodo & "' And cPlanillaCod = '" & psPlaCod & "'"
    
    oCon.AbreConexion
    oCon.Ejecutar sql
    
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Public Function GetOpeContPlanilla(psPlanillaCod As String, pnOpeEst As Long, pnOpeCont As Long) As Boolean
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = " Select  cOpeCodEst, cOpeCodCont From RHPlanillaTabla" _
        & " Where cRHPlanillaCod = '" & psPlanillaCod & "'"
    
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        pnOpeEst = -1
        pnOpeCont = -1
    Else
        pnOpeEst = IIf(IsNull(rs!cOpeCodEst), -1, rs!cOpeCodEst)
        pnOpeCont = IIf(IsNull(rs!cOpeCodCont), -1, rs!cOpeCodCont)
    End If
    
    oCon.CierraConexion
    Set oCon = Nothing
    Set rs = Nothing
End Function


Public Function GetAgencias() As Recordset
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = " SELECT  cAgeCod,cAgeDescripcion FROM AGENCIAS where cAgeCod not in ('09','10','11','12','02') order by cAgeCod asc "
    
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sql)
    
    Set GetAgencias = rs
    oCon.CierraConexion
    Set oCon = Nothing
    Set rs = Nothing
End Function

Public Function GetContadorPago(psCodPlanilla As String, psPeriodo As String) As Integer
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = " Select  count(*) as cant from rhplanilladetcon  where cRRHHPeriodo= '" & psPeriodo & "'and  cPlanillaCod ='" & psCodPlanilla & "' and bPagado = 1 "
    
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sql)
    
    GetContadorPago = rs!Cant
    oCon.CierraConexion
    Set oCon = Nothing
    Set rs = Nothing
End Function

'ALPA 20120417**********************************************
Public Function GetAsientoITF(ByVal pdFecha As Date, ByVal psOpeCod As String, ByVal psPlanillaCod As String, ByVal pgsCodAge As String, ByVal pgsCodUser As String, Optional ByVal pdFechaITF As Date, Optional ByVal psTipoPlanilla As String = "") As String
'*** PEAC 20130906 Optional ByVal psFechaITF As String

    Dim rsProv As ADODB.Recordset
    Set rsProv = New ADODB.Recordset

    '*** PEAC 20130326
    Dim rsCtaCntGstoITF As ADODB.Recordset
    Set rsCtaCntGstoITF = New ADODB.Recordset
    Dim sqlGITF As String
    Dim lcCtaGsto As String
    Dim lcCtaITF As String
    Dim lcFecPlanilla As String
    '*** FIN PEAC

    Dim dIni As Date, dFin As Date
    Dim sValor As String
    Dim sqlC As String
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oInterprete As DInterprete
    Set oInterprete = New DInterprete
    Dim lnContador As Integer
    Dim lsGlosa, lsMov As String
    Dim lnItem As Integer
    Dim lnSumITF As Currency
    Dim lsContador As String
    Dim psPlanillaReriodo As String
    Dim lnMov As Long
    Dim lsCta As String
    lnSumITF = 0
    'sqlC = " select cMovNro=isnull(dbo.Fnc_ObtieneOperacionFractalDiario('" & Format(pdFecha, "YYYYMMDD") & "'," & psOpeCod & "),'') "
    sqlC = " select cMovNro=isnull(dbo.Fnc_ObtieneOperacionFractalDiario('" & Format(pdFechaITF, "yyyymmdd") & "'," & psOpeCod & "),'') "

    If oCon.AbreConexion Then
    
        '*** PEAC 20130326
        sqlGITF = "exec stp_sel_CtasCntGastosITF"
        Set rsCtaCntGstoITF = oCon.CargaRecordSet(sqlGITF)
        lcCtaGsto = rsCtaCntGstoITF!cCtaGsto
        lcCtaITF = rsCtaCntGstoITF!cCtaITF
        Set rsCtaCntGstoITF = Nothing
        '*** FIN PEAC
    
        Set rsProv = oCon.CargaRecordSet(sqlC)
        If rsProv.EOF Or rsProv.BOF Then
            lsMov = ""
        Else
            lsMov = rsProv!cMovNro
        End If
        
        '*** PEAC 20140109
        If lsMov = "" Then
            lcFecPlanilla = Mid(Left(psTipoPlanilla, 8), 5, 4) & Mid(Left(psTipoPlanilla, 8), 3, 2) & Mid(Left(psTipoPlanilla, 8), 1, 2)
            
            sqlC = " select cMovNro=isnull(dbo.Fnc_ObtieneOperacionFractalDiario('" & lcFecPlanilla & "'," & psOpeCod & "),'') "
            Set rsProv = oCon.CargaRecordSet(sqlC)
            If rsProv.EOF Or rsProv.BOF Then
                lsMov = ""
            Else
                lsMov = rsProv!cMovNro
            End If
        End If
        '*** FIN PEAC
        
        rsProv.Close
        Set rsProv = Nothing
        Set rsProv = New ADODB.Recordset
        If Len(Trim(lsMov)) = 0 Then
            'sqlA = "exec stp_sel_AsientoPlanillaITF '" & CStr(Format(pdFecha, "DDMMYYYY")) & "','" & Trim(psPlanillaCod) & "'"
            sqlA = "exec stp_sel_AsientoPlanillaITF '" & Left(psTipoPlanilla, 8) & "','" & Trim(psPlanillaCod) & "'"
            Set rsProv = oCon.CargaRecordSet(sqlA)
            If Not (rsProv.EOF Or rsProv.BOF) Then
                Dim oMov As DMov
                Set oMov = New DMov
                oMov.BeginTrans
                lnItem = 0
                oMov.Inicio gsFormatoFecha
                lsGlosa = "ASIENTO ITF "
                lsGlosa = lsGlosa & " DE " & UCase(Format$(pdFecha, gsFormatoFechaView)) & " " & psTipoPlanilla
                'lsGlosa = lsGlosa & " DE " & UCase(Format$(pdFechaITF, gsFormatoFechaView))
                
                'lsMov = oMov.GeneraMovNro(pdFecha, pgsCodAge, pgsCodUser)
                lsMov = oMov.GeneraMovNro(pdFechaITF, pgsCodAge, pgsCodUser) '*** PEAC 20130906
                
                oMov.InsertaMov lsMov, psOpeCod, lsGlosa
                lnMov = oMov.GetnMovNro(lsMov)
                Do While Not rsProv.EOF
                    lnItem = lnItem + 1
                    'lsCta = "45140901" + rsProv!cAgenciaActual
                    lsCta = lcCtaGsto + rsProv!cAgenciaActual '*** PEAC 20130326
                    
                    oMov.InsertaMovCta lnMov, lnItem, lsCta, Abs(rsProv!nMontoITF)
                    lnSumITF = lnSumITF + Abs(rsProv!nMontoITF)
                    rsProv.MoveNext
                Loop
                lnItem = lnItem + 1
                'lsCta = "21140207"
                lsCta = lcCtaITF '*** PEAC 20130326
                
                oMov.InsertaMovCta lnMov, lnItem, lsCta, Abs(lnSumITF) * -1
                oMov.CommitTrans
            End If
        End If
    End If
    GetAsientoITF = lsMov
End Function

Public Function GetObtenerCodigoOperacion(psCodPlanilla As String) As String
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GetObtenerCodigoOperacion = ""
    
    sql = " stp_sel_OpePla '" & Trim(psCodPlanilla) & "'"
    oCon.AbreConexion
    
    Set rs = oCon.CargaRecordSet(sql)
    If Not (rs.BOF Or rs.EOF) Then
        GetObtenerCodigoOperacion = rs!COPECOD
    End If
    oCon.CierraConexion
    
    Set oCon = Nothing
    Set rs = Nothing
End Function
'***********************************************************

'*** PEAC 20130910
Public Function BuscaAsientoGenerado(ByVal psOpeCod As String, ByVal psCodPlanilla As String, ByRef psOpeDesc As String, ByRef psFechaAsnto As String) As Boolean

    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    BuscaAsientoGenerado = False
    psOpeDesc = ""
    psFechaAsnto = ""
    
    sql = "exec stp_sel_BuscaAsientoITF '" & psOpeCod & "','" & psCodPlanilla & "' "
    oCon.AbreConexion
    
    Set rs = oCon.CargaRecordSet(sql)
    If rs!nCantidad > 0 Then
        BuscaAsientoGenerado = True
        psOpeDesc = rs!COPECOD
        psFechaAsnto = rs!cFechaAsiento
    End If
    oCon.CierraConexion
    Set oCon = Nothing
    Set rs = Nothing
End Function
