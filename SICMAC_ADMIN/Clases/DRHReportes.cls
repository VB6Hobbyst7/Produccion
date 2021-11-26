VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DRHReportes"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Actualiza en Base Autorización Fisica
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB903050003
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB90305014E
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB903050112
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

Public Function GetReporteGralCampos() As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlE = "sp_columns reportegeneral"
    
    If oCon.AbreConexion Then
        Set GetReporteGralCampos = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function

Public Function GetReporteGralValores(psCodigo As String) As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlE = "Select Distinct " & psCodigo & " from ReporteGeneral"
    
    If oCon.AbreConexion Then
        Set GetReporteGralValores = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function

Public Function Rep5taRRHH() As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlE = " Select E.cRHCod, P.cPersNombre, P.cPersCod from RRHH E " _
         & " Inner Join Persona P On E.cPersCod = P.cPersCod" _
         & " where nRHEstado not like ('" & RHEstadosTpo.RHEstadosTpoRetirado & "%') And cRHCod like  'E%' Order By P.cPersNombre"
    
    If oCon.AbreConexion Then
        Set Rep5taRRHH = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function

'Public Function RepContabilidadRRHH(pdFecha As Date) As ADODB.Recordset
'    Dim sqlE As String
'    Dim oCon As DConecta
'    Dim rsC As ADODB.Recordset
'    Set rsC = New ADODB.Recordset
'    Set oCon = New DConecta
'
'    'sqlE = " Select Age.cAgeCod, Age.cAgeDescripcion," _
'         & " Condicion = Case CO.nRHContratoTpo" _
'         & "    when 0 then 'ESTABLES' Else 'CONTRATADOS' end," _
'         & "    Categoria = Case cRHCargoCategoria" _
'         & "    when '1' then 'FUNCIONARIO' Else 'EMPLEADO' end," _
'         & " Cantidad = count(*)" _
'         & " from RRHH EM" _
'         & " Inner join Agencias Age On EM.cAgenciaAsig = Age.cAgeCod" _
'         & " Inner join RHContrato CO On EM.cPersCod = CO.cPersCod" _
'         & " Inner Join RHCargos CA On EM.cPersCod = CA.cPersCod" _
'         & " Inner Join RHCargosTabla CAT On CAT.cRHCargoCod = CA.cRHCargoCod" _
'         & " where nRHEstado not like '" & RHEstadosTpo.RHEstadosTpoRetirado & "%' and cRHCod like 'E%'" _
'         & " And CO.cRHContratoNro in" _
'         & " (Select Max(CO1.cRHContratoNro) From RHContrato CO1 Where CO1.cPersCod = CO.cPersCod) And CA.dRHCargoFecha = (Select Max(CA1.dRHCargoFecha) From RHCargos CA1 Where CA.cPersCod = CA1.cPersCod  ) " _
'         & " group by Age.cAgeCod, Age.cAgeDescripcion, cRHCargoCategoria, CO.nRHContratoTpo Order by Age.cAgeCod"
'
'
'    sqlE = " Select cagecod,cagedescripcion, " _
'         & "        funcionarioEst = sum(case when categoria in ( 'FUNCIONARIO' , 'GERENTE') And Condicion = 'ESTABLES' then cantidad end)," _
'         & "        empleadoEst = sum(case  when categoria = 'EMPLEADO' And Condicion = 'ESTABLES'  then cantidad end)," _
'         & "        funcionarioCont = sum(case when categoria = 'FUNCIONARIO' And Condicion = 'CONTRATADOS' then cantidad end)," _
'         & "        empleadoCont = sum(case  when categoria = 'EMPLEADO' And Condicion = 'CONTRATADOS'  then cantidad end)" _
'         & "        FROM (" _
'         & "            SELECT CAGECOD, cAgeDescripcion, condicion,  Categoria = case bgerencia when 1 then 'GERENTE' else categoria end, sum(cantidad) cantidad" _
'         & "            FROM (" _
'         & "                Select  Age.cAgeCod, Age.cAgeDescripcion,bgerencia, Condicion = Case CO.nRHContratoTpo when 0 then 'ESTABLES' Else 'CONTRATADOS' end," _
'         & "                    Categoria = Case lefT(ca.cRHCargoCod,3)  when '001'  then 'FUNCIONARIO' when '002' then 'FUNCIONARIO'  when '003' then 'FUNCIONARIO' Else 'EMPLEADO' end," _
'         & "                    Cantidad = count(*)" _
'         & "                from RRHH EM" _
'         & "                Inner Join RHCargos CA On EM.cPersCod = CA.cPersCod" _
'         & "                Inner join RHContrato CO On EM.cPersCod = CO.cPersCod" _
'         & "                Inner join Agencias Age On CA.crhagenciaCodOficial = Age.cAgeCod" _
'         & "                Inner Join RHCargosTabla CAT On CAT.cRHCargoCod = CA.cRHCargoCod" _
'         & "                where CA.cPersCod In (" _
'         & "                    Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like '" & Format(pdFecha, "yyyymm") & "%' And cPlanillaCod = 'E01' and cperscod not in (Select cPersCod from rrhh Where datediff(month,dcese,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and datediff(day,dcese,'" & Format(pdFecha, gsFormatoFecha) & "') <> 0 and crhcod like 'E%' and cperscod  in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like '" & Format(pdFecha, "yyyymm") & "%' And cPlanillaCod = 'E01'))" _
'         & "                    Union All" _
'         & "                    Select cPersCod from rrhh Where datediff(month,dingreso,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and dcese is null and crhcod like 'E%' and cperscod not in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like '" & Format(pdFecha, "yyyymm") & "%' And cPlanillaCod = 'E01')" _
'         & "                    Union All" _
'         & "                    Select cPersCod from rrhh Where datediff(month,dingreso,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and crhcod like 'E%' and datediff(month,dcese,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and datediff(day,dcese,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and cperscod not in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like '" & Format(pdFecha, "yyyymm") & "%' And cPlanillaCod = 'E01') " _
'         & "                    Union All" _
'         & "                    Select cPersCod from rrhh Where datediff(month,dcese,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and datediff(day,dcese,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and crhcod like 'E%' and cperscod  in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like '" & Format(pdFecha, "yyyymm") & "%' And cPlanillaCod = 'E01'))" _
'         & "                And cRHCod like 'E%' And CO.cRHContratoNro in"
'
'     sqlE = sqlE & "         ( Select Max(CO1.cRHContratoNro) From RHContratoDet CO1" _
'         & "                  Where CO1.cPersCod = CO.cPersCod And Convert(varchar(10),dRHContratoFecha,112) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' )" _
'         & "                And CA.dRHCargoFecha = (Select Max(CA1.dRHCargoFecha) From RHCargos CA1 Where CA.cPersCod = CA1.cPersCod And convert(varchar(10),CA1.dRHCargoFecha,112) <= '" & Format(pdFecha, gsFormatoMovFecha) & "')" _
'         & "                group by Age.cAgeCod, Age.cAgeDescripcion, cRHCargoCategoria, CO.nRHContratoTpo ,ca.cRHCargoCod,bgerencia" _
'         & "                ) A group by CAGECOD,cagedescripcion,bgerencia,categoria, condicion" _
'         & "            ) b" _
'         & "            group by CAGECOD,cagedescripcion" _
'         & "            Union" _
'         & "            Select cAgeCod, cAgeDescripcion, funcionarioEst=0, funcionarioCont=0, empleadoEst=0, empleadoCont=0 From Agencias Where nEstado=0" _
'         & "            order by CAGECOD,cagedescripcion"
'
'    If oCon.AbreConexion Then
'        Set rsC = oCon.CargaRecordSet(sqlE)
'        Set RepContabilidadRRHH = rsC
'        oCon.CierraConexion
'    End If
'
'    Set oCon = Nothing
'End Function

Public Function RepContabilidadRRHH() As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Set oCon = New DConecta
    
    sqlE = " Select Age.cAgeCod, Age.cAgeDescripcion," _
         & " Condicion = Case CO.nRHContratoTpo" _
         & "    when 0 then 'ESTABLES' Else 'CONTRATADOS' end," _
         & "    Categoria = Case cRHCargoCategoria" _
         & "    when '1' then 'FUNCIONARIO' Else 'EMPLEADO' end," _
         & " Cantidad = count(*)" _
         & " from RRHH EM" _
         & " Inner join Agencias Age On EM.cAgenciaAsig = Age.cAgeCod" _
         & " Inner join RHContrato CO On EM.cPersCod = CO.cPersCod" _
         & " Inner Join RHCargos CA On EM.cPersCod = CA.cPersCod" _
         & " Inner Join RHCargosTabla CAT On CAT.cRHCargoCod = CA.cRHCargoCod" _
         & " where nRHEstado not like '" & RHEstadosTpo.RHEstadosTpoRetirado & "%' and cRHCod like 'E%'" _
         & " And CO.cRHContratoNro in" _
         & " (Select Max(CO1.cRHContratoNro) From RHContrato CO1 Where CO1.cPersCod = CO.cPersCod) And CA.dRHCargoFecha = (Select Max(CA1.dRHCargoFecha) From RHCargos CA1 Where CA.cPersCod = CA1.cPersCod  ) " _
         & " group by Age.cAgeCod, Age.cAgeDescripcion, cRHCargoCategoria, CO.nRHContratoTpo Order by Age.cAgeCod"
    
    If oCon.AbreConexion Then
        Set rsC = oCon.CargaRecordSet(sqlE)
        Set RepContabilidadRRHH = rsC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function


Public Function GetContratosVencidos(psFechaIni As String, psFechaFin As String) As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta
     
    sqlE = " Select EM.cPersCod, EM.cRHCod, dRHContratoInicio, dRHContratoFin, PE.cPersNombre, CAT.cRHCargoDescripcion" _
         & " from RHContrato CO" _
         & " Inner Join Persona PE On CO.cPersCod = PE.cPersCod" _
         & " Inner Join RRHH EM On EM.cPersCod = PE.cPersCod" _
         & " Inner Join (Select cPersCod,cRHContratoNro,dRHContratoFecha,dRHContratoInicio,dRHContratoFin,cRHContratoComentario,cUltimaActualizacion from RHContratoDet A" _
         & "              Where cRHContratoNro = (Select Max(cRHContratoNro) from RHContratoDet B Where A.cPersCod = B.cPersCod)) COD On CO.cPersCod = COD.cPersCod" _
         & " Inner Join RHCargos CA On CA.cPersCod = CO.cPersCod" _
         & " Inner Join RHCargosTabla CAT On CA.cRHCargoCod = CAT.cRHCargoCod" _
         & " where dRHContratoFin between '" & psFechaIni & "' and '" & psFechaFin & "'" _
         & " And ((nRHContratoTpo = " & RHContratoTipo.RHContratoTipoFijo & " And Left(EM.cRHCod,1) = 'E') Or (nRHContratoTpo = " & RHContratoTipo.RHContratoTipoFLaboral & " And Left(EM.cRHCod,1) = 'F') Or (nRHContratoTpo = " & RHContratoTipo.RHContratoTipoPractica & " And Left(EM.cRHCod,1) = 'P')) and nRHEstado not like '[78]%'" _
         & " And CO.cRHContratoNro In" _
         & " (Select Max(CO1.cRHContratoNro) From RHContrato CO1 Where CO1.cPersCod = CO.cPersCod)" _
         & " And CA.dRHCargoFecha In" _
         & " (Select Max(CA1.dRHCargoFecha) From RHCargos CA1 Where CA1.cPersCod = CA.cPersCod)" _
         & " Order by EM.cRHCod"
    
    If oCon.AbreConexion Then
        Set GetContratosVencidos = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function

'Public Function GetAsistencia(psFechaIni As String, psFechaFin As String, Optional pbPorAgencias As Boolean = False) As ADODB.Recordset
'    Dim sqlE As String
'    Dim oCon As DConecta
'    Set oCon = New DConecta
'
'    If pbPorAgencias Then
'        'sqlE = " Select RH.cRHCod, PE.cPersCod , PE.cPersNombre Nombre, Sum(IsNull(nRHAsistenciaTardanzasMin,0)) Tardanzas, " _
'        '     & " Sum(IsNull(nRHAsistenciaPermisoMin,0)) Personales, Sum(IsNull(nRHAsistenciaPermisoAutMin,0)) Permitidos, RH.cAgenciaActual, AGE.cAgeDescripcion" _
'        '     & " From RHAsistencia  RHA" _
'        '     & "    Inner Join Persona PE On RHA.cPersCod = PE.cPersCod" _
'        '     & "    Inner Join RRHH RH On RHA.cPersCod = RH.cPersCod" _
'        '     & "    Inner Join Agencias AGE On AGE.cAgeCod = RH.cAgenciaActual" _
'        '     & " Where dRHAsistenciaFechaRef Between '" & psFechaIni & "' And '" & psFechaFin & "' and nRHAsistenciaTardanzasMin <> 0 and left( RH.cRHCod,1) = 'E' " _
'        '     & " Group By PE.cPersCod, PE.cPersNombre, RH.cRHCod, RH.cAgenciaActual, AGE.cAgeDescripcion Order by  left( RH.cRHCod,1), RH.cAgenciaActual, RH.cRHCod"
'
'        sqlE = " Select RH.cRHCod, PE.cPersCod , PE.cPersNombre Nombre,PID.cPersIdnro DNI, Sum(IsNull(nRHAsistenciaTardanzasMin,0)) Tardanzas, " _
'             & " Sum(IsNull(nRHAsistenciaPermisoMin,0)) Personales, Sum(IsNull(nRHAsistenciaPermisoAutMin,0)) Permitidos, RH.cAgenciaActual, AGE.cAgeDescripcion" _
'             & " From RHAsistencia  RHA" _
'             & "    Inner Join Persona PE On RHA.cPersCod = PE.cPersCod" _
'             & "    Inner Join RRHH RH On RHA.cPersCod = RH.cPersCod" _
'             & "    left join PersId PID on RH.cPersCod = PID.cPersCod" _
'             & "    Inner Join Agencias AGE On AGE.cAgeCod = RH.cAgenciaActual" _
'             & " Where dRHAsistenciaFechaRef Between '" & psFechaIni & "' And '" & psFechaFin & "'" _
'             & " and   PID.cPersIdTpo = '1' and nRHAsistenciaTardanzasMin <> 0 and left( RH.cRHCod,1) = 'E' And RH.dIngreso <= dRHAsistenciaFechaRef And (RH.dCese Is Null Or RH.dCese >= dRHAsistenciaFechaRef) And RH.cPersCod Not In ('1120100021296','1120100043514','1120100048158','1120100120007','1120200066380','1120700010866') " _
'             & " Group By PE.cPersCod, PE.cPersNombre,PID.cPersIdnro, RH.cRHCod, RH.cAgenciaActual, AGE.cAgeDescripcion Order by  left( RH.cRHCod,1), RH.cAgenciaActual, RH.cRHCod"
'
'    Else
'        'sqlE = " Select RH.cRHCod, PE.cPersCod , PE.cPersNombre Nombre, Sum(IsNull(nRHAsistenciaTardanzasMin,0)) Tardanzas, " _
'        '     & " Sum(IsNull(nRHAsistenciaPermisoMin,0)) Personales, Sum(IsNull(nRHAsistenciaPermisoAutMin,0)) Permitidos" _
'        '     & " From RHAsistencia  RHA" _
'        '     & "    Inner Join Persona PE On RHA.cPersCod = PE.cPersCod" _
'        '     & "    Inner Join RRHH RH On RHA.cPersCod = RH.cPersCod" _
'        '     & " Where dRHAsistenciaFechaRef Between '" & psFechaIni & "' And '" & psFechaFin & "'" _
'        '     & " Group By PE.cPersCod, PE.cPersNombre, RH.cRHCod Order by RH.cRHCod"
'
'        sqlE = " Select RH.cRHCod, PE.cPersCod , PE.cPersNombre Nombre,PID.cPersIdnro DNI, Sum(IsNull(nRHAsistenciaTardanzasMin,0)) Tardanzas, " _
'             & " Sum(IsNull(nRHAsistenciaPermisoMin,0)) Personales, Sum(IsNull(nRHAsistenciaPermisoAutMin,0)) Permitidos,NTC.HE" _
'             & " From RHAsistencia  RHA" _
'             & "    Inner Join Persona PE On RHA.cPersCod = PE.cPersCod" _
'             & "    Inner Join RRHH RH On RHA.cPersCod = RH.cPersCod" _
'             & "    left join PersId PID on RH.cPersCod = PID.cPersCod" _
'             & "    left join dbo.getConsolidadoHE('" & psFechaIni & "','" & psFechaFin & "') NTC ON RHA.cPersCod = NTC.cPersCod" _
'             & " Where dRHAsistenciaFechaRef Between '" & psFechaIni & "' And '" & psFechaFin & "'" _
'             & " and   PID.cPersIdTpo = '1' And RH.dIngreso <= dRHAsistenciaFechaRef And (RH.dCese Is Null Or RH.dCese >= dRHAsistenciaFechaRef) And RH.cPersCod Not In ('1120100021296','1120100043514','1120100048158','1120100120007','1120200066380','1120700010866')" _
'             & " Group By PE.cPersCod, PE.cPersNombre,PID.cPersIdnro, RH.cRHCod,NTC.HE " _
'             & " Order by RH.cRHCod"
'
'
'    End If
'
'    If oCon.AbreConexion Then
'        oCon.Ejecutar "Set DateFirst 1"
'        Set GetAsistencia = oCon.CargaRecordSet(sqlE)
'        oCon.CierraConexion
'    End If
'
'    Set oCon = Nothing
'End Function

Public Function GetAsistencia(psFechaIni As String, psFechaFin As String, Optional pbPorAgencias As Boolean = False, Optional psCodagencia As String = "") As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If pbPorAgencias Then
        sqlE = " Select RH.cRHCod, PE.cPersCod , PE.cPersNombre Nombre, Sum(IsNull(nRHAsistenciaTardanzasMin,0)) Tardanzas, " _
             & " Sum(IsNull(nRHAsistenciaPermisoMin,0)) Personales, Sum(IsNull(nRHAsistenciaPermisoAutMin,0)) Permitidos,Sum(IsNull(nRHAsistenciaVacacionesMin,0))/480 Dias_Vacaciones,RH.cAgenciaActual, AGE.cAgeDescripcion" _
             & " From RHAsistencia  RHA" _
             & "    Inner Join Persona PE On RHA.cPersCod = PE.cPersCod" _
             & "    Inner Join RRHH RH On RHA.cPersCod = RH.cPersCod" _
             & "    Inner Join Agencias AGE On AGE.cAgeCod = RH.cAgenciaActual" _
             & " Where dRHAsistenciaFechaRef Between '" & psFechaIni & "' And '" & psFechaFin & "'"
            If psCodagencia <> "" Then
            sqlE = sqlE + " and  rh.cAgenciaAsig  = '" & psCodagencia & "'"
            End If
        sqlE = sqlE + "  Group By PE.cPersCod, PE.cPersNombre, RH.cRHCod, RH.cAgenciaActual, AGE.cAgeDescripcion Order by  left( RH.cRHCod,1), RH.cAgenciaActual, RH.cRHCod"
    Else
        sqlE = " Select RH.cRHCod, PE.cPersCod , PE.cPersNombre Nombre, Sum(IsNull(nRHAsistenciaTardanzasMin,0)) Tardanzas, " _
             & " Sum(IsNull(nRHAsistenciaPermisoMin,0)) Personales, Sum(IsNull(nRHAsistenciaPermisoAutMin,0)) Permitidos," _
             & " Sum(IsNull(nRHAsistenciaVacacionesMin,0))/480 Dias_Vacaciones " _
             & " From RHAsistencia  RHA" _
             & "    Inner Join Persona PE On RHA.cPersCod = PE.cPersCod" _
             & "    Inner Join RRHH RH On RHA.cPersCod = RH.cPersCod" _
             & " Where dRHAsistenciaFechaRef Between '" & psFechaIni & "' And '" & psFechaFin & "'" _
             & " Group By PE.cPersCod, PE.cPersNombre, RH.cRHCod Order by PE.cPersNombre"
    End If
    
    If oCon.AbreConexion Then
        Set GetAsistencia = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function

Public Function GetAsistenciaDet(psFechaIni As String, psFechaFin As String, psPersCod As String) As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    'sqlE = " Select dRHAsistenciaFechaRef,nRHAsistenciaTardanzasMin,nRHAsistenciaPermisoMin,cUltimaActualizacion From RHAsistencia Where cPersCod = '" & psPersCod & "' And dRHAsistenciaFechaRef Between '" & psFechaIni & "' And '" & psFechaFin & "'"
    
    sqlE = " Select Convert(Varchar(10),dFecha,103) Fecha, cDia, Convert(Varchar(10),dRHHorarioFecha,103) FechaVig ," _
         & " ISNull(Convert(Varchar(8),T1Ini,108),'N/A') T1INI, ISNull(Convert(Varchar(8),T1Fin,108),'N/A') T1FIN, ISNull(Convert(Varchar(8),T2Ini,108),'N/A') T2INI, ISNull(Convert(Varchar(8),T2Fin,108),'N/A') T2FIN," _
         & " ISNull(Convert(Varchar(8),A1Ini,108),'NO MARCO') A1INI, ISNull(Convert(Varchar(8),A1Fin,108),'NO MARCO') A1Fin, ISNull(Convert(Varchar(8),A2Ini,108),'NO MARCO') A2Ini, ISNull(Convert(Varchar(8),A2Fin,108),'NO MARCO') A2Fin, " _
         & " IsNUll(MT1,0) MT1, IsNull(MT2,0) MT2, ST1, ST2, PerNoLab, PerNoLabPer, cRHCod, cPersCod, cPersNombre," _
            & " ISNULL(CASE " _
            & " WHEN CONVERT(varchar(8),A1INI,108)>CONVERT(varchar(8),T1INI,108) THEN 0 " _
            & " ELSE DATEDIFF(minute,CONVERT(varchar(8),A1INI,108),CONVERT(varchar(8),T1INI,108)) " _
            & " END,0)+ " _
            & " ISNULL(CASE " _
            & " WHEN CONVERT(varchar(8),A1FIN,108)>CONVERT(varchar(8),T1FIN,108) THEN " _
            & " DATEDIFF(minute,CONVERT(varchar(8),T1FIN,108),CONVERT(varchar(8),A1FIN,108)) " _
            & " ELSE 0 " _
            & " END,0)+ " _
            & " ISNULL(CASE " _
            & " WHEN CONVERT(varchar(8),A2INI,108)>CONVERT(varchar(8),T2INI,108) THEN 0 " _
            & " ELSE DATEDIFF(minute,CONVERT(varchar(8),A2INI,108),CONVERT(varchar(8),T2INI,108)) " _
            & " END,0)+ " _
            & " ISNULL(CASE " _
            & " WHEN CONVERT(varchar(8),A2FIN,108)>CONVERT(varchar(8),T2FIN,108) THEN " _
            & " DATEDIFF(minute,CONVERT(varchar(8),T2FIN,108),CONVERT(varchar(8),A2FIN,108)) " _
            & " ELSE 0 " _
            & " END,0)HE,Comision, PerNoLabDes, bDiaSiguienteT1, bDiaSiguienteT2 " _
            & " From dbo.RHDetalleAsistencia('" & psPersCod & "','" & psFechaIni & "','" & psFechaFin & "')"

    If oCon.AbreConexion Then
        oCon.Ejecutar "Set DateFirst 1"
        Set GetAsistenciaDet = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function





Public Function GetConceptosPlanillaDetalle(psPeriodoIni As String, psPeriodoFin As String, pbIngresos As Boolean, Optional pbAportaciones As Boolean = False) As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If pbIngresos Then
        If pbAportaciones = False Then
            sqlE = " Select cRHConceptoCod,cRHConceptoDescripcion,cRHConceptoAbreviatura " _
                 & " From RHConceptoTabla" _
                 & " where " _
                 & " cRHConceptoCod in" _
                 & " (Select distinct cRHConceptoCod from RHPlanillaDetCon Where cRRHHPeriodo Between '" & psPeriodoIni & "' And '" & psPeriodoFin & "' And cRHConceptoCod like '" & Left(gsRHConceptoITOTINGCOD, 1) & "%' " _
                 & "  And cRHConceptoCod Not In ('" & gsRHConceptoITOTINGCOD & "','" & gsRHConceptoINETOPAGARCOD & "'))"
        Else
            sqlE = " Select cRHConceptoCod,cRHConceptoDescripcion,cRHConceptoAbreviatura " _
                 & " From RHConceptoTabla" _
                 & " where " _
                 & " cRHConceptoCod in" _
                 & " (Select distinct cRHConceptoCod from RHPlanillaDetCon Where cRRHHPeriodo Between '" & psPeriodoIni & "' And '" & psPeriodoFin & "' And cRHConceptoCod like '3%' " _
                 & "  And cRHConceptoCod Not In ('303'))"
        End If
    Else
        sqlE = " Select cRHConceptoCod,cRHConceptoDescripcion,cRHConceptoAbreviatura " _
             & " From RHConceptoTabla" _
             & " where " _
             & " cRHConceptoCod in" _
             & " (Select distinct cRHConceptoCod from RHPlanillaDetCon Where cRRHHPeriodo Between '" & psPeriodoIni & "' And '" & psPeriodoFin & "' And cRHConceptoCod like '" & Left(gsRHConceptoDTOTDESCOD, 1) & "%' " _
             & "  And cRHConceptoCod Not In ('" & gsRHConceptoDTOTDESCOD & "'))"
    End If
    
    If oCon.AbreConexion Then
        Set GetConceptosPlanillaDetalle = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function
    
Public Function GetConceptosPlanillaDetalleDatos(psPeriodoIni As String, psPeriodoFin As String, pbIngresos As Boolean, Optional pnMeses As Integer = 0) As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If pnMeses = 0 Then
        If pbIngresos Then
            sqlE = " Select EM.cRHCod, PE.cPersNombre, UTIL.Monto as Monto, cRHConceptoCod, cPlanillaCod From RRHH EM" _
                 & " Inner Join (" _
                 & "              Select Sum(nMonto) as Monto, cPersCod, cRHConCeptoCod, cPlanillaCod From RHPlanillaDetCon Where cRRHHPeriodo Between '" & psPeriodoIni & "' And '" & psPeriodoFin & "' " _
                 & "              and cRHConceptoCod like '" & Left(333, 1) & "%' And cRHConceptoCod Not In ('303')" _
                 & " Group By cPersCod, cRHConceptoCod, cPlanillaCod) As UTIL On EM.cPersCod = UTIL.cPersCod" _
                 & " Inner Join Persona PE On PE.cPersCod = EM.cPersCod And EM.cRHCod like 'E%'" _
                 & " Order by EM.cRHCod, cPlanillaCod , util.cRHConceptoCod"
                 
                 
                 
        Else
            sqlE = " Select EM.cRHCod, PE.cPersNombre, UTIL.Monto as Monto, cRHConceptoCod, cPlanillaCod From RRHH EM" _
                 & " Inner Join (" _
                 & "              Select Sum(nMonto) as Monto, cPersCod, cRHConCeptoCod, cPlanillaCod From RHPlanillaDetCon Where cRRHHPeriodo Between '" & psPeriodoIni & "' And '" & psPeriodoFin & "' " _
                 & "              and cRHConceptoCod like '" & Left(gsRHConceptoDTOTDESCOD, 1) & "%' And cRHConceptoCod Not In ('" & gsRHConceptoDTOTDESCOD & "')" _
                 & " Group By cPersCod, cRHConceptoCod, cPlanillaCod) As UTIL On EM.cPersCod = UTIL.cPersCod" _
                 & " Inner Join Persona PE On PE.cPersCod = EM.cPersCod And EM.cRHCod like 'E%'" _
                 & " Order by EM.cRHCod, cPlanillaCod , util.cRHConceptoCod"
        End If
    Else
        If pbIngresos Then
            sqlE = " Select EM.cRHCod, PE.cPersNombre, UTIL.Monto as Monto, cRHConceptoCod, cPlanillaCod From RRHH EM" _
                 & " Inner Join (" _
                 & "              Select Sum(nMonto) as Monto, cPersCod, cRHConCeptoCod, cPlanillaCod From RHPlanillaDetCon Where cRRHHPeriodo Between '" & psPeriodoIni & "' And '" & psPeriodoFin & "' " _
                 & "              and cRHConceptoCod like '" & Left(gsRHConceptoINETOPAGARCOD, 1) & "%' And cRHConceptoCod Not In ('" & gsRHConceptoITOTINGCOD & "','" & gsRHConceptoINETOPAGARCOD & "')" _
                 & " Group By cPersCod, cRHConceptoCod, cPlanillaCod) As UTIL On EM.cPersCod = UTIL.cPersCod" _
                 & " Inner Join Persona PE On PE.cPersCod = EM.cPersCod And EM.cRHCod like 'E%' "
                 '"And DateDiff(Month,dIngreso,'" & Format(DateAdd("m", 1, CDate(Right(psPeriodoIni, 2) & "/" & Mid(psPeriodoIni, 5, 2) & "/" & Left(psPeriodoIni, 4))), gsFormatoFecha) & "') > " & pnMeses
             sqlE = sqlE & " Order by EM.cRHCod, cPlanillaCod , util.cRHConceptoCod"
        Else
            sqlE = " Select EM.cRHCod, PE.cPersNombre, UTIL.Monto as Monto, cRHConceptoCod, cPlanillaCod From RRHH EM" _
                 & " Inner Join (" _
                 & "              Select Sum(nMonto) as Monto, cPersCod, cRHConCeptoCod, cPlanillaCod From RHPlanillaDetCon Where cRRHHPeriodo Between '" & psPeriodoIni & "' And '" & psPeriodoFin & "' " _
                 & "              and cRHConceptoCod like '" & Left(gsRHConceptoDTOTDESCOD, 1) & "%' And cRHConceptoCod Not In ('" & gsRHConceptoDTOTDESCOD & "')" _
                 & " Group By cPersCod, cRHConceptoCod, cPlanillaCod) As UTIL On EM.cPersCod = UTIL.cPersCod" _
                 & " Inner Join Persona PE On PE.cPersCod = EM.cPersCod And EM.cRHCod like 'E%' And DateDiff(Month,dIngreso,'" & Format(DateAdd("m", 1, CDate(Right(psPeriodoIni, 2) & "/" & Mid(psPeriodoIni, 5, 2) & "/" & Left(psPeriodoIni, 4))), gsFormatoFecha) & "') > " & pnMeses & " " _
                 & " Order by EM.cRHCod, cPlanillaCod , util.cRHConceptoCod"
        End If
    End If
    
    If oCon.AbreConexion Then
        Set GetConceptosPlanillaDetalleDatos = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function

Public Function GetRepAFP(pgsRuc As String, psPeriodoIni As String, psPeriodoFin As String) As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    ''PE004','PE005','PE008','PE010','PE011','PE014'
    sqlE = " Select '" & pgsRuc & "' AS Ruc, EM.cCUSPP CUSSP, PE.cPersNombre, V1.Monto, AFP.cRHAFPAbreviatura cEmpNumAFP, dbo.GetNombrePrimero(PE.cPersNombre) NPrimero, dbo.GetNombreSegundo(PE.cPersNombre) NSegundo FROM RHEMPLEADO EM " _
         & " Inner Join Persona PE On PE.cPersCod = EM.cPersCod" _
         & " Inner Join RHAFP AFP On AFP.cRHAFPPersCod = EM.cRHEmplAFPPersCod" _
         & " Inner Join" _
         & "   ( Select Sum(nMonto * Case cRHConceptoCod When '130' Then 1 else -1 end) Monto, cPersCod FROM RHPlanillaDetCon" _
         & "     Where cRRHHPeriodo Between '" & psPeriodoIni & "' And '" & psPeriodoFin & "' And  crrhhperiodo + cPlanillaCod <> '20040405E15' And cRHConceptoCod in ('" & gsRHConceptoITOTINGCOD & "','105','122') And cPlanillaCod In (Select cRhPlanillaCod  from rhplanillaconceptotabla where cRhConceptoCod = '" & RHConceptoDescuentos.gRHConcDctoAFPPrima & "') Group By cPersCod) V1 On V1.cPersCod = EM.cPersCod Order by PE.cPersNombre"
    
    If oCon.AbreConexion Then
        Set GetRepAFP = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function

Public Function GetEstabPDT(pgsRuc As String, psPeriodoIni As String, psPeriodoFin As String) As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta

    sqlE = "stp_sel_ReportePDTEstablecimiento '" & psPeriodoIni & "','" & psPeriodoFin & "'"
    If oCon.AbreConexion Then
        Set GetEstabPDT = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function

Public Function GetJorLabPDT(pgsRuc As String, psPeriodoIni As String, psPeriodoFin As String) As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta

    sqlE = "stp_sel_ReportePDTJornadaLab '" & psPeriodoIni & "','" & psPeriodoFin & "'"
    If oCon.AbreConexion Then
        Set GetJorLabPDT = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function
Public Function GetRepPDT(pgsRuc As String, psPeriodoIni As String, psPeriodoFin As String) As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
'    sqlE = " Select '01' AS TpoDoc, PE.cPersIDnro, IsNull(Round(30 * D1.Monto,0),0) DLAB, IsNull(V1.Monto,0) TIES, Case When rtrim(ltrim(IsNull(cRHEmplAFPPersCod,''))) <> '' Then 0 Else IsNull(V2.Monto,0) End TONP, IsNull(V2.Monto,0) TES, 0 TART, IsNull(V3.Monto,0) T5ta, IsNull(V4.Monto,0) I5ta " _
'         & " FROM RHEMPLEADO EM" _
'         & " Left Join PersID PE On PE.cPersCod = EM.cPersCod And cPersIDTpo = 1" _
'         & " Left Join (Select Sum(nMonto) Monto, cPersCod FROM RHPlanillaDetCon Where cRRHHPeriodo Between '" & psPeriodoIni & "' And '" & psPeriodoFin & "' And cRHConceptoCod In (" & gsRHCVarUsuPorcMesTrab & ")" _
'         & "            And cPlanillaCod In ('" & gsRHPlanillaSueldos & "')" _
'         & "            Group By cPersCod) D1 On D1.cPersCod = EM.cPersCod" _
'         & " Left Join (Select Sum(nMonto) Monto, cPersCod FROM RHPlanillaDetCon Where cRRHHPeriodo Between '" & psPeriodoIni & "' And '" & psPeriodoFin & "' And cRHConceptoCod = '" & gsRHConceptoITOTINGCOD & "'" _
'         & "            And cPlanillaCod In (Select cRhPlanillaCod  from rhplanillaconceptotabla where cRhConceptoCod In (301,304,306))" _
'         & "            Group By cPersCod) V1 On V1.cPersCod = EM.cPersCod" _
'         & " Left Join (Select Sum(nMonto) Monto, cPersCod FROM RHPlanillaDetCon Where cRRHHPeriodo Between '" & psPeriodoIni & "' And '" & psPeriodoFin & "' And cRHConceptoCod = '" & gsRHConceptoITOTINGCOD & "'" _
'         & "            And cPlanillaCod In (Select cRhPlanillaCod  from rhplanillaconceptotabla where cRhConceptoCod In (302,305,307))" _
'         & "            Group By cPersCod) V2 On V2.cPersCod = EM.cPersCod" _
'         & " Left Join (Select Sum(nMonto) Monto, cPersCod FROM RHPlanillaDetCon Where cRRHHPeriodo Between '" & psPeriodoIni & "' And '" & psPeriodoFin & "' And cRHConceptoCod = '" & gsRHConceptoITOTINGCOD & "'" _
'         & "            And cPlanillaCod  Not In ('" & gsRHPlanillaCTS & "') And cPlanillaCod Like 'E%'" _
'         & "            Group By cPersCod) V3 On V3.cPersCod = EM.cPersCod" _
'         & " Left Join (Select Sum(nMonto) Monto, cPersCod FROM RHPlanillaDetCon Where cRRHHPeriodo Between '" & psPeriodoIni & "' And '" & psPeriodoFin & "' And cRHConceptoCod In (201,217)" _
'         & "            Group By cPersCod) V4 On V4.cPersCod = EM.cPersCod" _
'         & " Inner Join (Select Distinct cPersCod from RHPlanillaDet Where cRRHHPeriodo Between '" & psPeriodoIni & "' And '" & psPeriodoFin & "' And cPlanillaCod Like 'E%') As PD On PD.cPersCod = EM.cPersCod "
'
'    sqlE = sqlE & " UNION ALL Select '01' AS TpoDoc, PE.cPersIDnro, 30 DLAB, 0 TIES, 0 TONP, IsNull(V1.Monto,0) TES, 0 TART, 0 T5ta, 0 I5ta " _
'         & " FROM RRHH EM" _
'         & " Left Join PersID PE On PE.cPersCod = EM.cPersCod And cPersIDTpo = 1" _
'         & " Left Join (Select Sum(nMonto) Monto, cPersCod FROM RHPlanillaDetCon Where cRRHHPeriodo Between '" & psPeriodoIni & "' And '" & psPeriodoFin & "' And cRHConceptoCod In ('" & gsRHConceptoITOTINGCOD & "')" _
'         & "            And cPlanillaCod In ('P01')" _
'         & "            Group By cPersCod) V1 On V1.cPersCod = EM.cPersCod" _
'         & " Inner Join (Select Distinct cPersCod from RHPlanillaDet Where cRRHHPeriodo Between '" & psPeriodoIni & "' And '" & psPeriodoFin & "' And cPlanillaCod Like 'P%') As PD On PD.cPersCod = EM.cPersCod "
    sqlE = "stp_sel_ReportePDTRemuTraba '" & psPeriodoIni & "','" & psPeriodoFin & "'"
    If oCon.AbreConexion Then
        Set GetRepPDT = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function

Public Function GetRepGen(psql As String) As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If oCon.AbreConexion Then
        Set GetRepGen = oCon.CargaRecordSet(psql)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function


Public Function GetListaPersonal(Optional pbTodos As Boolean = True, Optional pbAreas As Boolean = False) As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If pbTodos Then
        sqlE = " Select EM.cPersCod, EM.cRHCod, PE.cPersNombre, EM.dIngreso, EM.cAgenciaActual, TC.cAgeDescripcion, CA.nRHContratoTpo" _
             & " From RRHH EM" _
             & " Inner Join Persona PE On EM.cPersCod = PE.cPersCod Inner Join RHContrato CA On EM.cPersCod = CA.cPersCod  " _
             & " Inner Join Agencias TC On TC.cAgeCod = EM.cAgenciaActual where nRHEstado not like '" & RHEstadosTpo.RHEstadosTpoRetirado & "%'" _
             & " And  CA.cRHContratoNro In (Select max(CA1.cRHContratoNro) From RHContrato CA1 Where CA1.cPersCod = CA.cPersCod) order by EM.cAgenciaActual"
    Else
        If Not pbAreas Then
            sqlE = " Select EM.cPersCod, EM.cRHCod, PE.cPersNombre, EM.dIngreso, EM.cAgenciaActual, TC.cAgeDescripcion, CA.nRHContratoTpo" _
                 & " From RRHH EM" _
                 & " Inner Join Persona PE On EM.cPersCod = PE.cPersCod Inner Join RHContrato CA On EM.cPersCod = CA.cPersCod  " _
                 & " Inner Join Agencias TC On TC.cAgeCod = EM.cAgenciaActual where nRHEstado not like '" & RHEstadosTpo.RHEstadosTpoRetirado & "%'" _
                 & " And  CA.cRHContratoNro In (Select max(CA1.cRHContratoNro) From RHContrato CA1 Where CA1.cPersCod = CA.cPersCod) And TC.cAgeCod <> '" & gsAgenciaPrinsipal & "' order by EM.cAgenciaActual"
        Else
            sqlE = " Select EM.cPersCod, EM.cRHCod, PE.cPersNombre,EM.dIngreso," _
                 & " EM.cAgenciaActual, TC.cAreaDescripcion, RHC.nRHContratoTpo from RRHH EM" _
                 & " Inner Join Persona PE On EM.cPersCod = PE.cPersCod" _
                 & " Inner Join Areas TC On TC.cAreaCod = EM.cAreaCodActual" _
                 & " Inner Join RHContrato RHC On RHC.cPersCod = EM.cPersCod" _
                 & " where nRHEstado Not Like '" & RHEstadosTpo.RHEstadosTpoRetirado & "%' And EM.cAgenciaActual = '07'" _
                 & " And RHC.cRHContratoNro In (Select Max(RHC1.cRHContratoNro) From RHContrato RHC1 Where RHC.cPersCod = RHC1.cPersCod  )" _
                 & " order by TC.cAreaDescripcion"
        End If
    End If

    If oCon.AbreConexion Then
        Set GetListaPersonal = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function


Public Function GetPlanilllaResumen(ByVal pcTipoPlanilla As String, ByVal psPeriodo As String, ByVal pnTipo As Integer) As ADODB.Recordset
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta

sql = "SELECT cAgenciaActual,"
sql = sql & "    SUM(CASE WHEN SUBSTRING(CRHCONCEPTOCOD,1,1)='1' THEN NMONTO END) INGRESOS,"
sql = sql & "     SUM(CASE WHEN SUBSTRING(CRHCONCEPTOCOD,1,1)='2' THEN NMONTO END) EGRESOS,"
If pcTipoPlanilla = "E08" Then
    sql = sql & " SUM(CASE WHEN CRHCONCEPTOCOD='304' THEN NMONTO END) IES,"
    sql = sql & "     SUM(CASE WHEN CRHCONCEPTOCOD='305' THEN NMONTO END) IMP_SEG_SOC"
Else
    sql = sql & " SUM(CASE WHEN CRHCONCEPTOCOD='301' THEN NMONTO END) IES,"
    sql = sql & "     SUM(CASE WHEN CRHCONCEPTOCOD='302' THEN NMONTO END) IMP_SEG_SOC"
End If
sql = sql & " FROM  ("
sql = sql & "     Select AAA.cPErsCod, cRHConceptoCod, nMonto,  cAgenciaActual from ("
sql = sql & "         Select RH.cPErsCod, cRHConceptoCod, nMonto, cAgenciaActual"
sql = sql & "         from RHPlanillaDetCon RP"
sql = sql & "         Inner Join RRHH RH on RH.cPersCod = RP.cPersCod"
sql = sql & "        where cRRHHPeriodo like '" & psPeriodo & "%' and cPlanillaCod = '" & pcTipoPlanilla & "'"
'Sql = Sql & "         and not cRHConceptoCod in ('130','112','215','303','404')  and nMonto<>0"  cambio CAAU
sql = sql & "         and not cRHConceptoCod in ('129','130','131','132','112','215','303','404')  and nMonto<>0"
sql = sql & "     ) AAA"
sql = sql & "     Inner Join ("
sql = sql & "         Select cPersCod,cRHContratoNro,nRHContratoTpo from RHContrato RE where    cRHContratoNro = ("
sql = sql & "         Select max(cRHContratoNro) from RHContrato where  cPersCod = RE.cPersCod)"
sql = sql & "     ) CCC on CCC.cPersCod = AAA.cPersCod"
Select Case pnTipo
    Case 1
        sql = sql & " Where CCC.nRHContratoTpo = '1'"
    Case 2
        sql = sql & " Where CCC.nRHContratoTpo = '0'"
    Case 3
        ' no agregar en esta linea
End Select
sql = sql & " ) ZZZ"
sql = sql & " GROUP BY cAgenciaActual"
    
    If oCon.AbreConexion Then
        Set GetPlanilllaResumen = oCon.CargaRecordSet(sql)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing

End Function

Public Function GetEstadDetRRHH(psPeriodoIni As String, psPeriodoFin As String, pbMonto As Boolean, pbGastoDirecto As Boolean) As ADODB.Recordset
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If pbMonto Then
        If pbGastoDirecto Then
            sql = " Select * from  dbo.PlaneaGetMontoDirRHTotal('" & psPeriodoIni & "','" & psPeriodoFin & "')"
        Else
            sql = " Select * from  dbo.PlaneaGetMontoIndRHTotal('" & psPeriodoIni & "','" & psPeriodoFin & "') Order By nAgrupacion, Nivel"
        End If
    Else
        If pbGastoDirecto Then
            sql = " Select * from  dbo.PlaneaGetNumDirRHTotal('" & psPeriodoIni & "','" & psPeriodoFin & "')"
        Else
            sql = " Select * from  dbo.PlaneaGetNumIndRHTotal('" & psPeriodoIni & "','" & psPeriodoFin & "') Order By nAgrupacion, Nivel"
        End If
   End If
    
    If oCon.AbreConexion Then
        Set GetEstadDetRRHH = oCon.CargaRecordSet(sql)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function

#If DebugMode Then
    '##ModelId=3AB90305005E
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

Public Function GetRepAFPValida(psPeriodoIni As String, psPeriodoFin As String) As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlE = " Select (Select RHC.cCtaCod from RHCuentas RHC Where RHC.cCtaCod Like '_____2321%' And RHC.cPersCod = Vista.cPersCod) Cuenta , cPersNombre, Round(SueldoAfecto *  nRHAFPPrima / 100,2) Calculado, MontoPrima Cobrado, Round(SueldoAfecto *  nRHAFPPrima / 100,2) - MontoPrima Diferencia, cRHAFPAbreviatura From" _
         & " (Select PE.cPersCod, PE.cPersNombre, Case When V1.Monto > (Select Convert(Decimal(20,2),cRHConceptoFormula) Tope From rhconceptotabla where cRHConceptoCod = '722') Then (Select Convert(Decimal(20,2),cRHConceptoFormula) Tope From rhconceptotabla where cRHConceptoCod = '722') Else V1.Monto End  SueldoAfecto, V2.Monto MontoPrima, AFP.nRHAFPPrima , AFP.cRHAFPAbreviatura" _
         & " FROM RHEMPLEADO EM" _
         & " Inner Join Persona PE On PE.cPersCod = EM.cPersCod" _
         & " Inner Join RHAFP AFP On AFP.cRHAFPPersCod = EM.cRHEmplAFPPersCod" _
         & " Inner Join (Select Sum( Case When crhConceptoCod = '" & gsRHConceptoITOTINGCOD & "' then nMonto else nMonto * -1 end ) Monto, cPersCod" _
         & "                    From RHPlanillaDetCon" _
         & "             Where cRRHHPeriodo Between '" & psPeriodoIni & "' And '" & psPeriodoFin & "'" _
         & "             And cRHConceptoCod in ('" & gsRHConceptoITOTINGCOD & "','154','167','150','159','223','105','122')  And   crrhhperiodo + cPlanillaCod <> '20040405E15' And cPlanillaCod In (Select cRhPlanillaCod  from rhplanillaconceptotabla where cRhConceptoCod in ('" & RHConceptoDescuentos.gRHConcDctoAFPPrima & "','222') )" _
         & "             Group By cPersCod) V1 On V1.cPersCod = EM.cPersCod" _
         & " Left Join ( Select Sum(nMonto) Monto, cPersCod" _
         & "             From RHPlanillaDetCon" _
         & "             Where cRRHHPeriodo Between '" & psPeriodoIni & "' And '" & psPeriodoFin & "'" _
         & "             And cRHConceptoCod in ('202','222') And cPlanillaCod In (Select cRhPlanillaCod  from rhplanillaconceptotabla where cRhConceptoCod In ('" & RHConceptoDescuentos.gRHConcDctoAFPPrima & "','222'))" _
         & " Group By cPersCod) V2 On V2.cPersCod = EM.cPersCod) As Vista" _
         & " Where Round(SueldoAfecto * nRHAFPPrima / 100, 2) <> MontoPrima" _
         & " Order by cRHAFPAbreviatura, cPersNombre"
   
    If oCon.AbreConexion Then
        Set GetRepAFPValida = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function


Public Function GetRepAFPValidaComVar(psPeriodoIni As String, psPeriodoFin As String) As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlE = " Select (Select RHC.cCtaCod from RHCuentas RHC Where RHC.cCtaCod Like '__2321%' And RHC.cPersCod = Vista.cPersCod) Cuenta ,cPersNombre, Round(SueldoAfecto *  nRHAFPVariable / 100,2) Calculado, MontoPrima Cobrado, Round(SueldoAfecto *  nRHAFPVariable / 100,2) - MontoPrima Diferencia, cRHAFPAbreviatura From" _
         & " (Select PE.cPersCod, PE.cPersNombre, V1.Monto SueldoAfecto, V2.Monto MontoPrima, AFP.nRHAFPVariable , AFP.cRHAFPAbreviatura" _
         & " FROM RHEMPLEADO EM" _
         & " Inner Join Persona PE On PE.cPersCod = EM.cPersCod" _
         & " Inner Join RHAFP AFP On AFP.cRHAFPPersCod = EM.cRHEmplAFPPersCod" _
         & " Inner Join (Select Sum( Case When crhConceptoCod = '" & gsRHConceptoITOTINGCOD & "' then nMonto else nMonto * -1 end ) Monto, cPersCod" _
         & "                    From RHPlanillaDetCon" _
         & "             Where cRRHHPeriodo Between '" & psPeriodoIni & "' And '" & psPeriodoFin & "'" _
         & "             And cRHConceptoCod in ('" & gsRHConceptoITOTINGCOD & "','154','167','150','159','223','105','122')  And   crrhhperiodo + cPlanillaCod <> '20040405E15' And cPlanillaCod In (Select cRhPlanillaCod  from rhplanillaconceptotabla where cRhConceptoCod in ('" & RHConceptoDescuentos.gRHConcDctoAFPComVar & "','221') )" _
         & "             Group By cPersCod) V1 On V1.cPersCod = EM.cPersCod" _
         & " Left Join ( Select Sum(nMonto) Monto, cPersCod" _
         & "             From RHPlanillaDetCon" _
         & "             Where cRRHHPeriodo Between '" & psPeriodoIni & "' And '" & psPeriodoFin & "'" _
         & "             And cRHConceptoCod in ('204','221') And cPlanillaCod In (Select cRhPlanillaCod  from rhplanillaconceptotabla where cRhConceptoCod In ('" & RHConceptoDescuentos.gRHConcDctoAFPComVar & "','221'))" _
         & " Group By cPersCod) V2 On V2.cPersCod = EM.cPersCod) As Vista" _
         & " Where Round(SueldoAfecto * nRHAFPVariable / 100, 2) <> MontoPrima" _
         & " Order by cRHAFPAbreviatura, cPersNombre"
   
    If oCon.AbreConexion Then
        Set GetRepAFPValidaComVar = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function


Public Function GetHorasExtrasDet(psFechaIni As String, psFechaFin As String) As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Dim psPersCod As String
    Set oCon = New DConecta
 
   
         
    sqlE = "Select cPersCod,cPersNombre,Convert(Varchar(10),dFecha,103) Fecha, cDia, Convert(Varchar(10),dRHHorarioFecha,103) FechaVig ," _
            & " ISNull(Convert(Varchar(8),T1Ini,108),'N/A') T1INI, ISNull(Convert(Varchar(8),T1Fin,108),'N/A') T1FIN," _
            & " ISNull(Convert(Varchar(8),T2Ini,108),'N/A') T2INI, ISNull(Convert(Varchar(8),T2Fin,108),'N/A') T2FIN," _
            & " ISNull(Convert(Varchar(8),A1Ini,108),'NO MARCO') A1INI, ISNull(Convert(Varchar(8),A1Fin,108),'NO MARCO') A1Fin," _
            & " ISNull(Convert(Varchar(8),A2Ini,108),'NO MARCO') A2Ini, ISNull(Convert(Varchar(8),A2Fin,108),'NO MARCO') A2Fin," _
            & " IsNUll(MT1,0) MT1, IsNull(MT2,0) MT2," _
            & " ISNULL(CASE " _
            & " WHEN CONVERT(varchar(8),A1INI,108)>CONVERT(varchar(8),T1INI,108) THEN 0 " _
            & " ELSE DATEDIFF(minute,CONVERT(varchar(8),A1INI,108),CONVERT(varchar(8),T1INI,108)) " _
            & " END,0)+ " _
            & " ISNULL(CASE " _
            & " WHEN CONVERT(varchar(8),A1FIN,108)>CONVERT(varchar(8),T1FIN,108) THEN " _
            & " DATEDIFF(minute,CONVERT(varchar(8),T1FIN,108),CONVERT(varchar(8),A1FIN,108)) " _
            & " ELSE 0 " _
            & " END,0)+ " _
            & " ISNULL(CASE " _
            & " WHEN CONVERT(varchar(8),A2INI,108)>CONVERT(varchar(8),T2INI,108) THEN 0 " _
            & " ELSE DATEDIFF(minute,CONVERT(varchar(8),A2INI,108),CONVERT(varchar(8),T2INI,108)) " _
            & " END,0)+ " _
            & " ISNULL(CASE " _
            & " WHEN CONVERT(varchar(8),A2FIN,108)>CONVERT(varchar(8),T2FIN,108) THEN " _
            & " DATEDIFF(minute,CONVERT(varchar(8),T2FIN,108),CONVERT(varchar(8),A2FIN,108)) " _
            & " ELSE 0 " _
            & " END,0)HE,Comision, PerNoLabDes,cRHCod,cDNI,DATEDIFF(minute,T1FIN,T2INI)Refrigerio" _
            & " From RHDetAS('" & psFechaIni & "','" & psFechaFin & "')"
            
    If oCon.AbreConexion Then
        oCon.Ejecutar "Set DateFirst 1"
        Set GetHorasExtrasDet = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function


Public Function GetRHReportesPre() As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlE = " Select cDescripconReporte,cCodRep from  rhreportes where nEstadoRep = 0 and Carea ='PR' and cCodRep not in(51)"

    If oCon.AbreConexion Then
        oCon.Ejecutar "Set DateFirst 1"
        Set GetRHReportesPre = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function


Public Function GetReporteRHMinAcumulados(pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date, psCodagencia As String, psFechaIni As String, psFechaFin As String, pnMinAcum As Integer) As String
    On Error GoTo GetReporteRHMinAcumuladosErr
    Dim lsCod As String * 10
    Dim lsDescripcion As String * 50
    Dim lsTipo As String * 20
    Dim lsUltiMov As String * 27
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim lsNummin As String * 5
      
    Dim oAsisRep As DActualizaDatosHorarios
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oAsisRep = New DActualizaDatosHorarios
    
    'Set rs = oAsisRep.GetListadoTar(psFechaIni, psFechaFin, pnMinAcum, psCodagencia)
    Set rs = oAsisRep.GetListadoFaltas(psFechaIni, psFechaFin, psCodagencia, "S", "S", pnMinAcum)
    
    
    lsCadena = ""
    lnPagina = 0
    lnItem = 0
    If Not (rs.EOF And rs.BOF) Then
        lsCadena = lsCadena & CabeceraPagina("Listado de Tardanza  desde el " + psFechaIni + " al " + psFechaFin, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        lsCadena = lsCadena & Encabezado("Codigo;9;Descripcion;18;;17; Tardanzas;30; ;15;Min Sin Tolerancia;18;", lnItem)
        While Not rs.EOF
            lsCod = rs.Fields(0)
            lsDescripcion = rs.Fields(1)
            lsUltiMov = rs.Fields(2)
            lsNummin = rs.Fields(3)
            lsCadena = lsCadena & "   " & lsCod & "   " & lsDescripcion & "   " & lsUltiMov & "  " & lsNummin & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1
            If lnItem = 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("Criterios de Evaluacion", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Codigo;30;Descripcion;30; ;30;Minutos Acumulados;25; ;10;", lnItem)
            End If
            rs.MoveNext
        Wend
    End If
    GetReporteRHMinAcumulados = lsCadena
    Set oAsisRep = Nothing
    Exit Function
GetReporteRHMinAcumuladosErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosMerDem:ModificaMerDem Method")
End Function

Public Function GetAsistenciaDetNew(psFechaIni As String, psFechaFin As String, psPersCod As String) As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    'sqlE = " Select dRHAsistenciaFechaRef,nRHAsistenciaTardanzasMin,nRHAsistenciaPermisoMin,cUltimaActualizacion From RHAsistencia Where cPersCod = '" & psPersCod & "' And dRHAsistenciaFechaRef Between '" & psFechaIni & "' And '" & psFechaFin & "'"
    
    sqlE = " SPRH_GetFaltas_PorCod '" & psFechaIni & "','" & psFechaFin & "','" & psPersCod & "'"

    If oCon.AbreConexion Then
        oCon.Ejecutar "Set DateFirst 1"
        Set GetAsistenciaDetNew = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function

Public Function GetRHlistaTrabDep() As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    
    sqlE = " SP_RHlistaTrabDep "

    If oCon.AbreConexion Then
        oCon.Ejecutar "Set DateFirst 1"
        Set GetRHlistaTrabDep = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function


Public Function SP_RHlistaAnalistasAg() As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    

    If oCon.AbreConexion Then
        oCon.Ejecutar "Set DateFirst 1"
        
        sqlE = " SP_RHlistaAnalistasAg  "
        Set SP_RHlistaAnalistasAg = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function

Public Function GetRHReportes() As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    
    sqlE = " Select cDescripconReporte,cCodRep from  rhreportes where nEstadoRep = 0 and cCodRep not in(52) "

    If oCon.AbreConexion Then
        oCon.Ejecutar "Set DateFirst 1"
        Set GetRHReportes = oCon.CargaRecordSet(sqlE)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
End Function
'EJVG 20110818******************************************
Public Function GetDatosPersonaRescindida(ByVal psPersCod As String) As ADODB.Recordset
    Dim sqlE As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    
    sqlE = " exec stp_sel_obtenerDatosPersonaRescindida '" & psPersCod & "'"

    oCon.AbreConexion
    Set GetDatosPersonaRescindida = oCon.CargaRecordSet(sqlE)
    oCon.CierraConexion
        
    Set oCon = Nothing
End Function
'*******************************************************
