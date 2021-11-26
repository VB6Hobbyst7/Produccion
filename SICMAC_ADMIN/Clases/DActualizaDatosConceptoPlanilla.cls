VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DActualizaDatosConPlanilla"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9B08010391"
'Actualiza Datos de la Relacion concepto planilla
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902FF02D6
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB90300004C
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB903000006
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

'Agrega una nueva relación entre concepto y planilla
'##ModelId=3A9B16310036
Public Function AgregaConceptoPlanilla(psRHPlanillaCod As String, psRHPlanillaDescripcion As String, psUltimaActualizacion As String, psRHRefPago As String, psRHRefDescripcion As String, psRHRefInicio As String, psRHRefFin As String, pnRHRefParametro As Double, psRHRefPlaSub As String, psRHRefPlaEstado As String, psOpeCodEst As String, psOpeCodCont As String, psOpeAbono As String) As Boolean
    On Error GoTo AgregaConceptoPlanillaErr
    Dim sqlP As String
    Dim sqlPAdd As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlP = " Insert rhplanillatabla (cRHPlanillaCod,cRHPlanillaDescripcion,cUltimaActualizacion,cOpeCodEst, cOpeCodCont, cOpeAbono) " _
         & " Values('" & psRHPlanillaCod & "','" & psRHPlanillaDescripcion & "','" & psUltimaActualizacion & "','" & psOpeCodEst & "','" & psOpeCodCont & "','" & psOpeAbono & "')"
    
    sqlPAdd = " Insert RHreferenciaPlanilla (cRHPlanillaCod,dRHRefPago,cRHRefDescripcion,dRHRefInicio,dRHRefFin,nRHRefParametro,cRHRefPlaSub,bRHRefPlaEstado)" _
            & " Values('" & psRHPlanillaCod & "','" & psRHRefPago & "','" & psRHRefDescripcion & "','" & psRHRefInicio & "','" & psRHRefFin & "'," & pnRHRefParametro & ",'" & psRHRefPlaSub & "'," & psRHRefPlaEstado & ")"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlP
        oCon.Ejecutar sqlPAdd
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
AgregaConceptoPlanillaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:AgregaConceptoPlanilla Method")
End Function

'Elimina  una nueva relacion entre concepto y planilla
'
'##ModelId=3A9B16310072
Public Function EliminaConceptoPlanilla(psRHPlanillaCod As String) As Boolean
    On Error GoTo EliminaConceptoPlanillaErr
    Dim sqlP As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlP = " Update RHreferenciaPlanilla" _
         & " Set bRHRefPlaEstado = 0" _
         & " Where cRHPlanillaCod = '" & psRHPlanillaCod & "'"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlP
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
EliminaConceptoPlanillaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:EliminaConceptoPlanilla Method")
End Function

'Actualiza una nueva relacion entre concepto y planilla
'##ModelId=3A9B163100AE
Public Function ModificaConceptoPlanilla(psRHPlanillaCod As String, psRHPlanillaDescripcion As String, psUltimaActualizacion As String, psRHRefPago As String, psRHRefDescripcion As String, psRHRefInicio As String, psRHRefFin As String, pnRHRefParametro As Double, psRHRefPlaSub As String, psRHRefPlaEstado As String, psOpeCodEst As String, psOpeCodCont As String, psOpeAbono As String) As Boolean
    On Error GoTo ModificaConceptoPlanillaErr
    Dim sqlP As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If oCon.AbreConexion Then
        oCon.BeginTrans
            sqlP = " Update RHreferenciaPlanilla" _
                 & " Set dRHRefPago = '" & psRHRefPago & "', cRHRefDescripcion = '" & psRHRefDescripcion & "', dRHRefInicio = '" & psRHRefInicio & "', dRHRefFin = '" & psRHRefFin & "', nRHRefParametro = " & pnRHRefParametro & ", cRHRefPlaSub = '" & psRHRefPlaSub & "',bRHRefPlaEstado = " & psRHRefPlaEstado & "" _
                 & " Where cRHPlanillaCod = '" & psRHPlanillaCod & "'"
            oCon.Ejecutar sqlP
            sqlP = " Update RHPlanillaTabla" _
                 & " Set cOpeCodEst = '" & psOpeCodEst & "' ,cOpeCodCont = '" & psOpeCodCont & "' ,cOpeAbono = '" & psOpeAbono & "'" _
                 & " Where cRHPlanillaCod = '" & psRHPlanillaCod & "'"
            oCon.Ejecutar sqlP
        oCon.CommitTrans
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
ModificaConceptoPlanillaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:ModificaConceptoPlanilla Method")
End Function

Public Function GetPlanillaCodigo(psTipo As String) As String
    On Error GoTo ModificaConceptoPlanillaErr
    Dim sqlP As String
    Dim rsP As ADODB.Recordset
    Set rsP = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlP = " Select Top 1 cRHPlanillaCod Codigo From RHPlanillaTabla" _
         & " Where cRHPlanillaCod like '" & Left(psTipo, 1) & "%' Order by cRHPlanillaCod Desc"
    
    If oCon.AbreConexion Then
        Set rsP = oCon.CargaRecordSet(sqlP)
        If rsP.EOF And rsP.BOF Then
            GetPlanillaCodigo = psTipo & "01"
        Else
            GetPlanillaCodigo = psTipo & Format(CCur(Right(rsP!codigo, 2)) + 1, "00")
        End If
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
ModificaConceptoPlanillaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:ModificaConceptoPlanilla Method")
End Function

'Devuleve la relacion de conceptos de la Planilla
'##ModelId=3A9B1DE800E0
Public Function GetRelacionPlanillaConcepto(psRHPlanillaCod As String) As Recordset
    On Error GoTo GetRelacionPlanillaConceptoErr

    'your code goes here...

    Exit Function
GetRelacionPlanillaConceptoErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetRelacionPlanillaConcepto Method")
End Function

'Devuelve las Planillas Activas
'##ModelId=3A9BECFD01ED
Public Function GetPlanillas(Optional psRHTipoContrato As String = "", Optional pbRHPlanillavalidas As Boolean = False, Optional pbTodas As Boolean = True) As Recordset
    On Error GoTo GetPlanillasErr
    Dim sqlP As String
    Dim sqlPWhere As String
    Dim sqlPInner As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlPWhere = ""
    sqlPInner = ""
    If psRHTipoContrato <> "" Then sqlPWhere = " Where RHT.cRHPlanillaCod like '" & psRHTipoContrato & "%'"
    If pbRHPlanillavalidas Then sqlPInner = " Inner Join RHReferenciaPlanilla RHRP On RHT.cRHPlanillaCod = RHRP.cRHPlanillaCod "
    
    If pbTodas Then
        sqlP = "Select RHT.cRHPlanillaCod Codigo, RHT.cRHPlanillaDescripcion Descrip, 2 Nivel from RHPlanillaTabla RHT " & sqlPInner & sqlPWhere & " Order By RHT.cRHPlanillaCod"
    Else
        sqlP = "Select RHT.cRHPlanillaCod Codigo, RHT.cRHPlanillaDescripcion Descrip, 2 Nivel from RHPlanillaTabla RHT Inner Join RHReferenciaPlanilla RP On RHT.cRHPlanillaCod = RP.cRHPlanillaCod And bRHRefPlaEstado = 1 " & sqlPInner & sqlPWhere & " Order By RHT.cRHPlanillaCod"
    End If
    
    If oCon.AbreConexion Then
        Set GetPlanillas = oCon.CargaRecordSet(sqlP)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetPlanillasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetPlanillas Method")
End Function

Public Function GetPlanillasTpo(psRHPlanillaTpo As String) As Recordset
    On Error GoTo GetPlanillasErr
    Dim sqlP As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlP = "Select cRRHHPeriodo+cPlanillaCod Codigo, cPlaInsDes Descripcion from rhplanilla Where cPlanillaCod = '" & psRHPlanillaTpo & "' order by cRRHHPeriodo desc"
    
    If oCon.AbreConexion Then
        Set GetPlanillasTpo = oCon.CargaRecordSet(sqlP)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetPlanillasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetPlanillas Method")
End Function

Public Function GetPlanillasConceptos(psRHPlanillaTpo As String) As Recordset
    On Error GoTo GetPlanillasErr
    Dim sqlP As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlP = " Select left(RHC.cRHConceptoCod,1) Cab, RHC.cRHconceptoMeno Nemo, RHC.cRHConceptoDescripcion Nombre From RHConceptoTabla RHC" _
         & " Inner Join RHPlanillaConceptoTabla RHP On RHC.cRHConceptoCod = RHP.cRHConceptoCod" _
         & " Where cRHPlanillaCod = '" & psRHPlanillaTpo & "' And cRHConceptoEstado = '" & RHConceptoEstado.gsRHCEstAct & "'" _
         & " and left(RHC.cRHConceptoCod,1) between '" & RHConceptosTpoVisible.RHConceptosTpoVIngreso & "' And '" & RHConceptosTpoVisible.RHConceptosTpoVVarUsuario & "'" _
         & " Order by left(cRHconceptoMeno,2) Desc, nRHConceptoOrden "
    
    If oCon.AbreConexion Then
        Set GetPlanillasConceptos = oCon.CargaRecordSet(sqlP)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetPlanillasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetPlanillas Method")
End Function

'Public Function GetPlanillasPersona(psRHPlanillaTpo As String, pbIni As Boolean, psPeriodo As String, pdIni As Date, pdFin As Date) As Recordset
'    On Error GoTo GetPlanillasErr
'    Dim sqlP As String
'    Dim oCon As DConecta
'    Set oCon = New DConecta
'
'    If pbIni Then
'        If psRHPlanillaTpo = "E04" Then
'            sqlP = InputBox("Ingrese el año del cual desea pagar las Utilidades ?")
'
'            If Not IsNumeric(sqlP) Then
'                MsgBox "Debe ingresar un valor valido", vbInformation, "Aviso"
'                Exit Function
'            End If
'
'            sqlP = " Select distinct PE.cPersCod Codigo, EM.cRHCod, PE.cPersNombre Nombre, '' cComentario  From Persona PE " _
'                 & " Inner Join RRHH EM On PE.cPersCod = EM.cPersCod And EM.cRHCod <> 'E00220' " _
'                 & " Where PE.cPersCod In " _
'                 & " (Select distinct cPersCod From rhplanilladetcon where cplanillacod in ('" & gsRHPlanillaSueldos & "','E02','E06','E13','E03','E09','E17') and crhconceptocod = '" & gRHConcIngTotIngresos & "'" _
'                 & "   and crrhhperiodo like '" & sqlP & "%')" _
'                 & " Union" _
'                 & " Select distinct PE.cPersCod Codigo, EM.cRHCod, PE.cPersNombre Nombre, '' cComentario " _
'                 & " From Persona PE" _
'                 & " Inner Join RRHH EM On PE.cPersCod = EM.cPersCod" _
'                 & " Where PE.cPersCod In  (Select distinct cPersCod From rhplanilladetcon" _
'                 & " where cplanillacod in ('E08') and crhconceptocod In ('155','156','157') " _
'                 & " and crrhhperiodo like '" & sqlP & "%') And PE.cPersCod Not In (" _
'                 & "     Select distinct PE.cPersCod Codigo" _
'                 & "     From Persona PE" _
'                 & "     Inner Join RRHH EM On PE.cPersCod = EM.cPersCod And EM.cRHCod <> 'E00220'" _
'                 & "     Where PE.cPersCod In  (Select distinct cPersCod From rhplanilladetcon" _
'                 & "     where cplanillacod in ('E01','E02','E06','E13','E03','E09','E17') and crhconceptocod = '" & gRHConcIngTotIngresos & "'" _
'                 & " and crrhhperiodo like '" & sqlP & "%')) And dCese >= '" & Format(CDate("01/01/" & sqlP), gsFormatoFecha) & "' Order By cRHCod"
'
'        ElseIf psRHPlanillaTpo = gsRHPlanillaVacaciones Then
'            'sqlP = " Select distinct PE.cPersCod Codigo, EM.cRHCod , PE.cPersNombre Nombre from RHConcepto RHC" _
'                 & " Inner Join Persona PE On RHC.cPersCod = PE.cPersCod" _
'                 & " Inner Join RRHH EM On RHC.cPersCod = EM.cPersCod" _
'                 & " Inner Join rhplanillaaplicacion RHA On EM.nRHEstado = RHA.nRHEstado" _
'                 & " Where RHA.cRHPlanillaCod = '" & psRHPlanillaTpo & "' And Left(EM.cRHCod,1) = '" & Left(psRHPlanillaTpo, 1) & "' Order By EM.cRHCod"
'            sqlP = " Select distinct PE.cPersCod Codigo, RH.cRHCod , PE.cPersNombre Nombre, PNL.cComentario from RHPeriodosNoLaborados PNL " _
'                 & " Inner Join RRHH RH On PNL.cPersCod = RH.cPersCod" _
'                 & " Inner Join Persona PE On PNL.cPersCod = PE.cPersCod" _
'                 & " Inner Join rhplanillaaplicacion RHA On PNL.nRHPeriodoTpo = RHA.nRHEstado" _
'                 & " Where RHA.cRHPlanillaCod = '" & psRHPlanillaTpo & "' And Left(RH.cRHCod,1) = '" & Left(psRHPlanillaTpo, 1) & "' And dSolicitadoInicio >= '" & Format(pdIni, gsFormatoFecha) & "' And dSolicitadoInicio <= '" & Format(pdFin, gsFormatoFecha) & "' Order By RH.cRHCod"
'
'        ElseIf psRHPlanillaTpo <> gsRHPlanillaSueldos Then
'            sqlP = " Select distinct PE.cPersCod Codigo, EM.cRHCod , PE.cPersNombre Nombre, '' cComentario  From " _
'                 & " RRHH EM  " _
'                 & " Inner Join Persona PE On PE.cPersCod = EM.cPersCod" _
'                 & " Inner Join rhplanillaaplicacion RHA On EM.nRHEstado = RHA.nRHEstado" _
'                 & " Where RHA.cRHPlanillaCod = '" & psRHPlanillaTpo & "' And Left(EM.cRHCod,1) = '" & Left(psRHPlanillaTpo, 1) & "' Order By EM.cRHCod"
'        Else
'            sqlP = " Select distinct PE.cPersCod Codigo, EM.cRHCod, EM.bAgregaPlanilla AgregaPlanilla, PE.cPersNombre Nombre, Case When EM.nRHEstado IN (501, 503) Then EM.nRHEstado When IsNUll(PNL.Dias,0) < 30 Then 201 Else PNL.nRHPeriodoTpo End nRHEstado, Case When EM.nRHEstado IN (501, 503) Then 'SUBSIDIADO ' When Case When IsNUll(PNL.Dias,0) < 30 Then 201 Else AP.nRHEstado End = '201' Then '' Else TC.cConsDescripcion + '-' End + IsNull(PNL.cComentario,'') cComentario" _
'                 & " From RRHH EM" _
'                 & " Inner Join Persona PE On EM.cPersCod = PE.cPersCod" _
'                 & " Left  Join (Select cPersCod, cComentario, nRHPeriodoTpo, (Select Sum(datediff(day,dSolicitadoInicio,dSolicitadoFin)) Dias from rhperiodosnolaborados PNLT where nRHPeriodoTpo in (301, 501, 503) And dSolicitadoInicio >= '" & Format(CDate("20/" & Format(DateAdd("d", -1, pdIni), "mm/yyyy")), gsFormatoFecha) & "' And  dSolicitadoInicio < '" & Format(CDate("20/" & Format(pdIni, "mm/yyyy")), gsFormatoFecha) & "' Group By PNLT.cPersCod having  PNLT.cPersCod = PNL.cPersCod) Dias from rhperiodosnolaborados PNL where nRHPeriodoTpo in (301) And dSolicitadoInicio = (Select Max(dSolicitadoInicio) from rhperiodosnolaborados PNL1 where nRHPeriodoTpo in (301) And dSolicitadoInicio >= '" & Format(CDate("20/" & Format(DateAdd("d", -1, pdIni), "mm/yyyy")), gsFormatoFecha) & "' And  dSolicitadoInicio < '" & Format(CDate("20/" & Format(pdIni, "mm/yyyy")), gsFormatoFecha) & "'" _
'                 & "             And PNL.cPersCod = PNL1.cPersCod) And dSolicitadoInicio >= '" & Format(CDate("20/" & Format(DateAdd("d", -1, pdIni), "mm/yyyy")), gsFormatoFecha) & "'" _
'                 & "             And  dSolicitadoInicio < '" & Format(CDate("20/" & Format(pdIni, "mm/yyyy")), gsFormatoFecha) & "') As PNL On PNL.cPersCod = EM.cPersCod" _
'                 & " Inner Join Constante TC On TC.nConsCod like '" & gRHEstado & "' And TC.nConsValor = EM.nRHEstado" _
'                 & " left  Join RHPlanillaAplicacion AP On AP.nRHEstado = EM.nRHEstado" _
'                 & " Where em.nRHEstado < " & 700 & " And EM.cRHCod Like '" & Left(psRHPlanillaTpo, 1) & "%' order by EM.cRHCod"
'        End If
'    Else
'        If psRHPlanillaTpo = gsRHPlanillaSueldos Then
'            sqlP = " Select Distinct PE.cPersCod Codigo, EM.cRHCod, PE.cPersNombre Nombre, EM.nRHEstado, RHP.cComentario From RHPlanillaDet E" _
'                 & " Inner Join Persona PE On E.cPersCod = PE.cPersCod" _
'                 & " Inner Join RRHH EM On EM.cPersCod = E.cPersCod" _
'                 & " Inner Join RHPlanillaDet RHP On RHP.cPersCod = E.cPerscod And RHP.cRRHHPeriodo = '" & psPeriodo & "' And RHP.cPlanillaCod = '" & psRHPlanillaTpo & "'" _
'                 & " Where E.cRRHHPeriodo = '" & psPeriodo & "' And E.cPlanillaCod = '" & psRHPlanillaTpo & "' And  Left(EM.cRHCod,1) = '" & Left(psRHPlanillaTpo, 1) & "' " _
'                 & " order by EM.cRHCod"
'        Else
'            sqlP = " Select Distinct PE.cPersCod Codigo, EM.cRHCod, PE.cPersNombre Nombre, EM.nRHEstado, RHP.cComentario From RHPlanillaDetCon E" _
'                 & " Inner Join Persona PE On E.cPersCod = PE.cPersCod" _
'                 & " Inner Join RRHH EM On EM.cPersCod = E.cPersCod" _
'                 & " Inner Join RHPlanillaDet RHP On RHP.cPersCod = E.cPerscod And RHP.cRRHHPeriodo = '" & psPeriodo & "' And RHP.cPlanillaCod = '" & psRHPlanillaTpo & "'" _
'                 & " Where E.cRRHHPeriodo = '" & psPeriodo & "' And E.cPlanillaCod = '" & psRHPlanillaTpo & "'  " _
'                 & " order by EM.cRHCod" ' And  Left(EM.cRHCod,1) = '" & Left(psRHPlanillaTpo, 1) & "'
'        End If
'    End If
'
'    If oCon.AbreConexion Then
'        Set GetPlanillasPersona = oCon.CargaRecordSet(sqlP)
'        oCon.CierraConexion
'    End If
'
'    Set oCon = Nothing
'    Exit Function
'GetPlanillasErr:
'    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetPlanillas Method")
'End Function

Public Function GetPlanillasPersona(psRHPlanillaTpo As String, pbIni As Boolean, psPeriodo As String) As Recordset
    On Error GoTo GetPlanillasErr
    Dim sqlP As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If pbIni Then
        If psRHPlanillaTpo = "E99" Then
            sqlP = InputBox("Ingrese el año del cual desea pagar las Utilidades ?")
            
            If Not IsNumeric(sqlP) Then
                MsgBox "Debe ingresar un valor valido", vbInformation, "Aviso"
                Exit Function
            End If
            
            sqlP = " Select distinct(PE.cPersCod + EM.cRHCod), PE.cPersCod Codigo, EM.cRHCod, PE.cPersNombre Nombre From Persona PE " _
                 & " Inner Join RRHH EM On PE.cPersCod = EM.cPersCod And EM.cRHCod <> 'E00220' " _
                 & " Where PE.cPersCod In " _
                 & " (Select distinct cPersCod From rhplanilladetcon where cplanillacod in ('E01','E02','E06','E13','E03','E09','E17') and crhconceptocod = '" & gRHConcIngTotIngresos & "'" _
                 & "   and crrhhperiodo like '" & sqlP & "%') Order By PE.cPersNombre"
        ElseIf psRHPlanillaTpo <> gsRHPlanillaSueldos Then
            sqlP = " Select distinct PE.cPersCod Codigo, EM.cRHCod , PE.cPersNombre Nombre from RHConcepto RHC" _
                 & " Inner Join Persona PE On RHC.cPersCod = PE.cPersCod" _
                 & " Inner Join RRHH EM On RHC.cPersCod = EM.cPersCod" _
                 & " Inner Join rhplanillaaplicacion RHA On EM.nRHEstado = RHA.nRHEstado" _
                 & " Where RHA.cRHPlanillaCod = '" & psRHPlanillaTpo & "' And Left(EM.cRHCod,1) = '" & Left(psRHPlanillaTpo, 1) & "' Order By PE.cPersNombre"
        Else
            sqlP = " Select distinct PE.cPersCod Codigo, EM.cRHCod, EM.bAgregaPlanilla AgregaPlanilla, PE.cPersNombre Nombre, AP.nRHEstado, Estado = Case When AP.nRHEstado = '201' Then '' Else TC.cConsDescripcion End " _
                 & " From RHConcepto E" _
                 & " Inner Join Persona PE On E.cPersCod = PE.cPersCod" _
                 & " Inner Join RRHH EM On E.cPersCod = EM.cPersCod" _
                 & " Inner Join Constante TC On TC.nConsCod like '" & gRHEstado & "' And TC.nConsValor = EM.nRHEstado" _
                 & " left  Join RHPlanillaAplicacion AP On AP.nRHEstado = EM.nRHEstado" _
                 & " Where left(EM.nRHEstado,1) <> '" & Left(7, 1) & "' And Left(EM.cRHCod,1) = '" & Left(psRHPlanillaTpo, 1) & "'  AND  left(EM.nRHEstado,1) <> '8'   order by PE.cPersNombre"
        End If
    Else
        If psRHPlanillaTpo = gsRHPlanillaSueldos Then
            sqlP = " Select Distinct PE.cPersCod Codigo, EM.cRHCod, PE.cPersNombre Nombre, EM.nRHEstado, RHP.cComentario From RHPlanillaDet E" _
                 & " Inner Join Persona PE On E.cPersCod = PE.cPersCod" _
                 & " Inner Join RRHH EM On EM.cPersCod = E.cPersCod" _
                 & " Inner Join RHPlanillaDet RHP On RHP.cPersCod = E.cPerscod And RHP.cRRHHPeriodo = '" & psPeriodo & "' And RHP.cPlanillaCod = '" & psRHPlanillaTpo & "'" _
                 & " Where E.cRRHHPeriodo = '" & psPeriodo & "' And E.cPlanillaCod = '" & psRHPlanillaTpo & "' And  Left(EM.cRHCod,1) = '" & Left(psRHPlanillaTpo, 1) & "'     " _
                 & " order by PE.cPersNombre"
        Else
            'MAVM 20120403 ***
            'sqlP = " Select Distinct PE.cPersCod Codigo, EM.cRHCod, PE.cPersNombre Nombre, EM.nRHEstado, RHP.cComentario From RHPlanillaDetCon E" _
            '     & " Inner Join Persona PE On E.cPersCod = PE.cPersCod" _
            '     & " Inner Join RRHH EM On EM.cPersCod = E.cPersCod" _
            '     & " Inner Join RHPlanillaDet RHP On RHP.cPersCod = E.cPerscod And RHP.cRRHHPeriodo = '" & psPeriodo & "' And RHP.cPlanillaCod = '" & psRHPlanillaTpo & "'" _
            '     & " Where E.cRRHHPeriodo = '" & psPeriodo & "' And E.cPlanillaCod = '" & psRHPlanillaTpo & "' " _
            '     & " order by PE.cPersNombre"
                 
            sqlP = " Select Distinct cRHCod Codigo, cRHCod cRHCod, E.cPersNombre Nombre, '201' nRHEstado, '' cComentario From RHPlanillaDetCon E" _
                 & " Where E.cRRHHPeriodo = '" & psPeriodo & "' And E.cPlanillaCod = '" & psRHPlanillaTpo & "' " _
                 & " Order by E.cPersNombre"
            '***
        End If
    End If
    
    If oCon.AbreConexion Then
        Set GetPlanillasPersona = oCon.CargaRecordSet(sqlP)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetPlanillasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetPlanillas Method")
End Function

Public Function GetPlanillasExistenteDatos(psRHPlanillaTpo As String, psPeriodo As String) As Recordset
    On Error GoTo GetPlanillasErr
    Dim sqlP As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlP = " Select dRRHHPago dPago, dRRHHInicio dIni, dRRHHFin dFin, dRRHHImpre Impre, nTpoCambio TC, cPlaInsDes Des, cRHPlanillaEstado Estado,  IsNUll(nTpoCambioFijo,0) TCF from rhplanilla " _
         & " Where cRRHHPeriodo = '" & psPeriodo & "' And cPlanillaCod = '" & psRHPlanillaTpo & "'"

    If oCon.AbreConexion Then
        Set GetPlanillasExistenteDatos = oCon.CargaRecordSet(sqlP)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetPlanillasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetPlanillas Method")
End Function

Public Function GetPlanillasExistenteDetalle(psRHPlanillaTpo As String, psPeriodo As String) As Recordset
    On Error GoTo GetPlanillasErr
    Dim sqlP As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    'MAVM 20120214 Fractal ***
    ' sqlP = " Select RH.cRHCod, RH.cPersCod, CO.cRHconceptoMeno, nMonto, CO.nRHConceptoOrden Orden From RHPlanillaDetCon PD" _
    '      & " Inner Join RHConceptoTabla CO On PD.cRHConceptoCod = CO.cRHConceptoCod" _
    '      & " Inner Join RRHH RH On RH.cPersCod = PD.cPersCod" _
    '      & " Where cRRHHPeriodo = '" & psPeriodo & "' And cPlanillaCod = '" & psRHPlanillaTpo & "' AND CO.cRHConceptoCod <> '" & gRHConcIngNetoPagar & "'" _
    '      & " Order by RH.cRHCod, left(CO.cRHconceptoMeno,2) Desc, nRHConceptoOrden ASC"
         
    'sqlP = " Select RH.cRHCod, RH.cPersCod, CO.cRHconceptoMeno, nMonto, CO.nRHConceptoOrden Orden From RHPlanillaDetCon PD" _
    '     & " Inner Join RHConceptoTabla CO On PD.cRHConceptoCod = CO.cRHConceptoCod" _
    '     & " Inner Join RRHH RH On RH.cPersCod = PD.cPersCod" _
    '     & " Where cRRHHPeriodo = '" & psPeriodo & "' And cPlanillaCod = '" & psRHPlanillaTpo & "'" _
    '     & " Order by RH.cRHCod, left(CO.cRHconceptoMeno,2) Desc, nRHConceptoOrden ASC"
    
    sqlP = " Select cRHCod, cRHCod cPersCod, CO.cRHconceptoMeno, nMonto, CO.nRHConceptoOrden Orden, cCtaCod From RHPlanillaDetCon PD" _
         & " Inner Join RHConceptoTabla CO On PD.cRHConceptoCod = CO.cRHConceptoCod" _
         & " Where cRRHHPeriodo = '" & psPeriodo & "' And cPlanillaCod = '" & psRHPlanillaTpo & "'" _
         & " Order by PD.cPersNombre"
    
    '***

    If oCon.AbreConexion Then
        Set GetPlanillasExistenteDetalle = oCon.CargaRecordSet(sqlP)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetPlanillasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetPlanillas Method")
End Function

Public Function GetPlanillasExistenteComentario(psRHPlanillaTpo As String, psPeriodo As String) As Recordset
    On Error GoTo GetPlanillasErr
    Dim sqlP As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlP = " Select RH.cRHCod, '_Comentario' cConcepCod, cComentario From RHPlanillaDet PD" _
         & " Inner Join RRHH RH On RH.cPersCod = PD.cPersCod" _
         & " Where cRRHHPeriodo = '" & psPeriodo & "' And cPlanillaCod = '" & psRHPlanillaTpo & "'" _
         & " Order By RH.cRHCod"

    If oCon.AbreConexion Then
        Set GetPlanillasExistenteComentario = oCon.CargaRecordSet(sqlP)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetPlanillasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetPlanillas Method")
End Function

'Devuelve las Planillas Activas
'##ModelId=3A9BECFD01ED
Public Function GetPlanilla(psRHPlanillaCod As String) As Recordset
    On Error GoTo GetPlanillasErr
    Dim sqlP As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlP = " Select dRHRefPago Pago, cRHRefDescripcion Descrip, dRHRefInicio Inicio, dRHRefFin Fin, nRHRefParametro Parametro, cRHRefPlaSub SubCuenta, bRHRefPlaEstado Estado , cOpeCodEst ,cOpeCodCont ,cOpeAbono " _
         & " From rhreferenciaplanilla A" _
         & " Inner Join RHPlanillaTabla B On A.cRHPlanillaCod = B.cRHPlanillaCod " _
         & " Where A.cRHPlanillaCod = '" & psRHPlanillaCod & "'"
    If oCon.AbreConexion Then
        Set GetPlanilla = oCon.CargaRecordSet(sqlP)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetPlanillasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetPlanillas Method")
End Function

'Devuelve las Planillas Activas
'##ModelId=3A9BECFD01ED
'Public Function GetPlanillaRep(pbOrdenado As Boolean, pbParaContabilidad As Boolean, psPlanillaPeriodo As String, psPlanillaCod As String, psEmpDes As String) As Recordset
'    On Error GoTo GetPlanillasErr
'    Dim sqlP As String
'    Dim oCon As DConecta
'    Set oCon = New DConecta
'    Dim ldFecha As Date
'
'    ldFecha = CDate(Right(psPlanillaPeriodo, 2) & "/" & Mid(psPlanillaPeriodo, 5, 2) & "/" & Left(psPlanillaPeriodo, 4))
'
'    If pbOrdenado Then
'        sqlP = " Select Age.cAgeDescripcion, dbo.GetDRHCargos3(EM.cPersCod, '" & Format(ldFecha, gsFormatoFecha) & "') cCodAge, EM.cPersCod, TipCon = CASE RHC.nRHContratoTpo" _
'             & "              When 0 THEN 'ESTABLE'" _
'             & "              ELSE 'CONTRATADO' End," _
'             & "              PD.cComentario" _
'             & " From RRHH EM" _
'             & " Left Join RHPlanillaDet PD On EM.cPersCod = PD.cPersCod And PD.cRRHHPeriodo = '" & psPlanillaPeriodo & "' And  PD.cPlanillaCod = '" & psPlanillaCod & "' " _
'             & " Left Join RHContrato RHC On RHC.cPersCod = EM.cPersCod" _
'             & " Inner Join Agencias Age On EM.cAgenciaActual = Age.cAgeCod" _
'             & " Where EM.cPersCod in (" _
'             & "                       Select distinct PD.cPersCod From RHPlanillaDet PD" _
'             & "                       Where PD.cRRHHPeriodo = '" & psPlanillaPeriodo & "' And PD.cPlanillaCod = '" & psPlanillaCod & "'" _
'             & "                      )" _
'             & " And PD.cPersCod Not In (" & psEmpDes & ")" _
'             & " And cRHContratoNro In (Select Max(cRHContratoNro) From RHContrato R Where R.cPersCod = EM.cPersCod  And Left(cUltimaActualizacion,8) <= '" & Format(ldFecha, gsFormatoMovFecha) & "')" _
'             & " Order By TipCon Desc, EM.cRHCod"
'    Else
'        If pbParaContabilidad Then
'            sqlP = " Select Age.cAgeDescripcion, dbo.GetDRHCargos3(EM.cPersCod, '" & Format(ldFecha, gsFormatoFecha) & "') cCodAge, EM.cPersCod, TipCon = CASE RHC.nRHContratoTpo" _
'                 & "              When 0 THEN 'ESTABLE'" _
'                 & "              ELSE 'CONTRATADO' End," _
'                 & "              PD.cComentario" _
'                 & " From RRHH EM" _
'                 & " Left Join RHPlanillaDet PD On EM.cPersCod = PD.cPersCod And PD.cRRHHPeriodo = '" & psPlanillaPeriodo & "' And  PD.cPlanillaCod = '" & psPlanillaCod & "' " _
'                 & " Left Join RHContrato RHC On RHC.cPersCod = EM.cPersCod" _
'                 & " Inner Join Agencias Age On EM.cAgenciaActual = Age.cAgeCod" _
'                 & " Where EM.cPersCod in (" _
'                 & "                       Select distinct PD.cPersCod From RHPlanillaDet PD" _
'                 & "                       Where PD.cRRHHPeriodo = '" & psPlanillaPeriodo & "' And PD.cPlanillaCod = '" & psPlanillaCod & "'" _
'                 & "                      )" _
'                 & " And PD.cPersCod Not In (" & psEmpDes & ")" _
'                 & " And cRHContratoNro In (Select Max(cRHContratoNro) From RHContrato R Where R.cPersCod = EM.cPersCod  And Left(cUltimaActualizacion,8) <= '" & Format(ldFecha, gsFormatoMovFecha) & "')" _
'                 & " Order By TipCon Desc, cCodAge, EM.cRHCod"
'        Else
'             'sqlP = " Select Age.cAgeDescripcion, EM.cAgenciaAsig cCodAge, EM.cPersCod, TipCon = CASE RHC.nRHContratoTpo" _
'                 & "              When 0 THEN 'ESTABLE'" _
'                 & "              ELSE 'CONTRATADO' End," _
'                 & "              PD.cComentario" _
'                 & " From RRHH EM" _
'                 & " Left Join RHPlanillaDet PD On EM.cPersCod = PD.cPersCod And PD.cRRHHPeriodo = '" & psPlanillaPeriodo & "' And  PD.cPlanillaCod = '" & psPlanillaCod & "' " _
'                 & " Left Join RHContrato RHC On RHC.cPersCod = EM.cPersCod" _
'                 & " Inner Join Agencias Age On EM.cAgenciaActual = Age.cAgeCod" _
'                 & " Where EM.cPersCod in (" _
'                 & "                       Select distinct PD.cPersCod From RHPlanillaDetCon PD" _
'                 & "                       Where PD.cRRHHPeriodo = '" & psPlanillaPeriodo & "' And PD.cPlanillaCod = '" & psPlanillaCod & "'" _
'                 & "                      )" _
'                 & " And PD.cPersCod Not In (" & psEmpDes & ")" _
'                 & " And cRHContratoNro In (Select Max(cRHContratoNro) From RHContrato R Where R.cPersCod = EM.cPersCod)" _
'                 & " Order By EM.cAgenciaAsig,EM.cRHCod,TipCon Desc"  '
'
'
'            sqlP = " Select Age.cAgeDescripcion, dbo.GetDRHCargos3(EM.cPersCod, '" & Format(ldFecha, gsFormatoFecha) & "') cCodAge, EM.cPersCod, TipCon = CASE IsNull(RHC.nRHContratoTpo,1) " _
'                 & "              When 0 THEN 'ESTABLE'" _
'                 & "              ELSE 'CONTRATADO' End," _
'                 & "              PD.cComentario" _
'                 & " From RRHH EM" _
'                 & " Left Join RHPlanillaDet PD On EM.cPersCod = PD.cPersCod And PD.cRRHHPeriodo = '" & psPlanillaPeriodo & "' And  PD.cPlanillaCod = '" & psPlanillaCod & "' " _
'                 & " Left Join RHContrato RHC On RHC.cPersCod = EM.cPersCod" _
'                 & " Inner Join Agencias Age On dbo.GetDRHCargos3(EM.cPersCod, '" & Format(ldFecha, gsFormatoFecha) & "') = Age.cAgeCod" _
'                 & " Where EM.cPersCod in (" _
'                 & "                       Select distinct PD.cPersCod From RHPlanillaDet PD" _
'                 & "                       Where PD.cRRHHPeriodo = '" & psPlanillaPeriodo & "' And PD.cPlanillaCod = '" & psPlanillaCod & "'" _
'                 & "                      )" _
'                 & " And PD.cPersCod Not In (" & psEmpDes & ")" _
'                 & " And cRHContratoNro In (Select Max(cRHContratoNro) From RHContrato R Where R.cPersCod = EM.cPersCod And Left(cUltimaActualizacion,8) <= '" & Format(ldFecha, gsFormatoMovFecha) & "')" _
'                 & " Order By cCodAge ,TipCon Desc, EM.cRHCod "  '
'        End If
'    End If
'
'    If oCon.AbreConexion Then
'        Set GetPlanillaRep = oCon.CargaRecordSet(sqlP)
'        oCon.CierraConexion
'    End If
'
'    Set oCon = Nothing
'    Exit Function
'GetPlanillasErr:
'    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetPlanillas Method")
'End Function

Public Function GetPlanillaRep(pbOrdenado As Boolean, pbParaContabilidad As Boolean, psPlanillaPeriodo As String, psPlanillaCod As String, psEmpDes As String) As Recordset
    On Error GoTo GetPlanillasErr
    Dim sqlP As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If pbOrdenado Then
        sqlP = " Select Age.cAgeDescripcion, EM.cAgenciaActual cCodAge, EM.cPersCod, TipCon = CASE RHC.nRHContratoTpo" _
             & "              When 0 THEN 'ESTABLE'" _
             & "              ELSE 'CONTRATADO' End," _
             & "              PD.cComentario" _
             & " From RRHH EM" _
             & " Left Join RHPlanillaDet PD On EM.cPersCod = PD.cPersCod And PD.cRRHHPeriodo = '" & psPlanillaPeriodo & "' And  PD.cPlanillaCod = '" & psPlanillaCod & "' " _
             & " Left Join RHContrato RHC On RHC.cPersCod = EM.cPersCod" _
             & " Inner Join Agencias Age On EM.cAgenciaActual = Age.cAgeCod" _
             & " Where EM.cPersCod in (" _
             & "                       Select distinct PD.cPersCod From RHPlanillaDet PD" _
             & "                       Where PD.cRRHHPeriodo = '" & psPlanillaPeriodo & "' And PD.cPlanillaCod = '" & psPlanillaCod & "'" _
             & "                      )" _
             & " And PD.cPersCod Not In (" & psEmpDes & ")" _
             & " And cRHContratoNro In (Select Max(cRHContratoNro) From RHContrato R Where R.cPersCod = EM.cPersCod)" _
             & " Order By TipCon Desc, EM.cRHCod"
    Else
        If pbParaContabilidad Then
            sqlP = " Select Age.cAgeDescripcion, EM.cAgenciaActual cCodAge, EM.cPersCod, TipCon = CASE RHC.nRHContratoTpo" _
                 & "              When 0 THEN 'ESTABLE'" _
                 & "              ELSE 'CONTRATADO' End," _
                 & "              PD.cComentario" _
                 & " From RRHH EM" _
                 & " Left Join RHPlanillaDet PD On EM.cPersCod = PD.cPersCod And PD.cRRHHPeriodo = '" & psPlanillaPeriodo & "' And  PD.cPlanillaCod = '" & psPlanillaCod & "' " _
                 & " Left Join RHContrato RHC On RHC.cPersCod = EM.cPersCod" _
                 & " Inner Join Agencias Age On EM.cAgenciaActual = Age.cAgeCod" _
                 & " Where EM.cPersCod in (" _
                 & "                       Select distinct PD.cPersCod From RHPlanillaDetCon PD" _
                 & "                       Where PD.cRRHHPeriodo = '" & psPlanillaPeriodo & "' And PD.cPlanillaCod = '" & psPlanillaCod & "'" _
                 & "                      )" _
                 & " And PD.cPersCod Not In (" & psEmpDes & ")" _
                 & " And cRHContratoNro In (Select Max(cRHContratoNro) From RHContrato R Where R.cPersCod = EM.cPersCod)" _
                 & " Order By TipCon Desc, EM.cAgenciaAsig, EM.cRHCod"
        Else
            sqlP = " Select Age.cAgeDescripcion, EM.cAgenciaAsig cCodAge, EM.cPersCod, TipCon = CASE RHC.nRHContratoTpo" _
                 & "              When 0 THEN 'ESTABLE'" _
                 & "              ELSE 'CONTRATADO' End," _
                 & "              PD.cComentario" _
                 & " From RRHH EM" _
                 & " Left Join RHPlanillaDet PD On EM.cPersCod = PD.cPersCod And PD.cRRHHPeriodo = '" & psPlanillaPeriodo & "' And  PD.cPlanillaCod = '" & psPlanillaCod & "' " _
                 & " Left Join RHContrato RHC On RHC.cPersCod = EM.cPersCod" _
                 & " Inner Join Agencias Age On EM.cAgenciaActual = Age.cAgeCod" _
                 & " Where EM.cPersCod in (" _
                 & "                       Select distinct PD.cPersCod From RHPlanillaDetCon PD" _
                 & "                       Where PD.cRRHHPeriodo = '" & psPlanillaPeriodo & "' And PD.cPlanillaCod = '" & psPlanillaCod & "'" _
                 & "                      )" _
                 & " And PD.cPersCod Not In (" & psEmpDes & ")" _
                 & " And cRHContratoNro In (Select Max(cRHContratoNro) From RHContrato R Where R.cPersCod = EM.cPersCod)" _
                 & " Order By EM.cAgenciaAsig,EM.cRHCod,TipCon Desc"  '
        End If
    End If
    
    If oCon.AbreConexion Then
        Set GetPlanillaRep = oCon.CargaRecordSet(sqlP)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetPlanillasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetPlanillas Method")
End Function

'Devuelve las Planillas Activas
'##ModelId=3A9BECFD01ED
Public Function GetLista(psPlanillaPeriodo As String, psPlanillaCod As String) As Recordset
    On Error GoTo GetPlanillasErr
    Dim sqlP As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oConst As NConstSistemas
    Set oConst = New NConstSistemas
    
    If oConst.LeeConstSistema(gConstSistBitCentral) = 1 Then
        sqlP = " Select EM.cRHCod, PE.cPersNombre, PD.nMonto, EC.cCtaCod,Substring(EC.cCtaCod,9,1) Moneda, Case nRHContratoTpo When 0 Then 'ESTABLE' Else 'CONTRATADO' End Condicion, EM.cAgenciaActual, PLA.nTpoCambio " _
             & " From RRHH EM" _
             & " Left Join RHCuentas EC On EM.cPersCod = EC.cPersCod And Substring(EC.cCtaCod,6,3) = '" & IIf(psPlanillaCod = gsRHPlanillaCTS, Producto.gCapCTS, Producto.gCapAhorros) & "'" _
             & " Left Join RHContrato EML On EM.cPersCod = EML.cPersCod" _
             & " Inner Join Persona PE On PE.cPersCod = EM.cPersCod" _
             & " Inner Join RHPlanillaDetCon PD On PD.cPersCod = EM.cPersCod" _
             & " Inner Join RHPlanilla PLA On PLA.cRRHHPeriodo = PD.cRRHHPeriodo And PLA.cPlanillaCod = PD.cPlanillaCod " _
             & " where PD.cPlanillaCod = '" & psPlanillaCod & "' and PD.cRRHHPeriodo = '" & psPlanillaPeriodo & "' And PD.cRHConceptoCod = '" & gsRHConceptoINETOPAGARCOD & "'" _
             & " And nMonto <> 0 And EML.cRHContratoNro In (Select Max(EML1.cRHContratoNro) From RHContrato EML1 Where EML1.cPersCod = EML.cPersCod)" _
             & " group by EC.cCtaCod, Substring(EC.cCtaCod,9,1), EM.cRHCod, PE.cPersNombre, PD.nMonto, nRHContratoTpo, EM.cAgenciaActual, PLA.nTpoCambio " _
             & " order by EM.cAgenciaActual, nRHContratoTpo, Substring(EC.cCtaCod,6,1)"
    Else
        sqlP = " Select EM.cRHCod, PE.cPersNombre, PD.nMonto, EC.cCtaCod,Substring(EC.cCtaCod,6,1) Moneda, Case nRHContratoTpo When 0 Then 'ESTABLE' Else 'CONTRATADO' End Condicion, EM.cAgenciaActual, PLA.nTpoCambio  " _
             & " From RRHH EM" _
             & " Left Join RHCuentas EC On EM.cPersCod = EC.cPersCod And Substring(EC.cCtaCod,3,3) = '" & IIf(psPlanillaCod = gsRHPlanillaCTS, Producto.gCapCTS, Producto.gCapAhorros) & "'" _
             & " Left Join RHContrato EML On EM.cPersCod = EML.cPersCod" _
             & " Inner Join Persona PE On PE.cPersCod = EM.cPersCod" _
             & " Inner Join RHPlanillaDetCon PD On PD.cPersCod = EM.cPersCod" _
             & " Inner Join RHPlanilla PLA On PLA.cRRHHPeriodo = PD.cRRHHPeriodo And PLA.cPlanillaCod = PD.cPlanillaCod " _
             & " where PD.cPlanillaCod = '" & psPlanillaCod & "' and PD.cRRHHPeriodo = '" & psPlanillaPeriodo & "' And PD.cRHConceptoCod = '" & gsRHConceptoINETOPAGARCOD & "'" _
             & " And nMonto <> 0 And EML.cRHContratoNro In (Select Max(EML1.cRHContratoNro) From RHContrato EML1 Where EML1.cPersCod = EML.cPersCod)" _
             & " group by EC.cCtaCod, Substring(EC.cCtaCod,9,1), EM.cRHCod, PE.cPersNombre, PD.nMonto, nRHContratoTpo, EM.cAgenciaActual, PLA.nTpoCambio " _
             & " order by EM.cAgenciaActual, nRHContratoTpo, Substring(EC.cCtaCod,6,1)"
    End If
    
    If oCon.AbreConexion Then
        Set GetLista = oCon.CargaRecordSet(sqlP)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetPlanillasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetPlanillas Method")
End Function



'Devuelve las Planillas Activas
'##ModelId=3A9BECFD01ED
Public Function GetPlanillaConcepLibres(psRHPlanillaCod As String, psGrupo As String) As Recordset
    On Error GoTo GetPlanillasErr
    Dim sqlP As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If psGrupo = "5" Then
        sqlP = " Select cRHConceptoCod Codigo, cRHConceptoDescripcion Nombre From RHConceptoTabla CO" _
             & " where cRHConceptoCod not in (Select cRHConceptoCod from RHPlanillaConceptoTabla Where cRHPlanillaCod = '" & psRHPlanillaCod & "') And Left(cRHConceptoCod,1) in ('1','2','3','4') And cRHConceptoEstado = '1'"
    Else
        sqlP = " Select cRHConceptoCod Codigo, cRHConceptoDescripcion Nombre From RHConceptoTabla CO" _
             & " where cRHConceptoCod not in (Select cRHConceptoCod from RHPlanillaConceptoTabla Where cRHPlanillaCod = '" & psRHPlanillaCod & "') And Left(cRHConceptoCod,1) = '" & psGrupo & "' And cRHConceptoEstado = '1'"
    End If
        
    If oCon.AbreConexion Then
        Set GetPlanillaConcepLibres = oCon.CargaRecordSet(sqlP)
        oCon.CierraConexion
    End If

    Set oCon = Nothing
    Exit Function
GetPlanillasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetPlanillas Method")
End Function

'Devuelve las Planillas Activas
'##ModelId=3A9BECFD01ED
Public Function GetPlanillaConcepUsados(psRHPlanillaCod As String, psGrupo As String) As Recordset
    On Error GoTo GetPlanillasErr
    Dim sqlP As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If psGrupo = "5" Then
        sqlP = " Select cRHConceptoCod Codigo, cRHConceptoDescripcion Nombre From RHConceptoTabla CO" _
             & " where cRHConceptoCod  in (Select cRHConceptoCod  from RHPlanillaConceptoTabla Where cRHPlanillaCod = '" & psRHPlanillaCod & "') And Left(cRHConceptoCod,1) in ('1','2','3','4') And cRHConceptoEstado = '1'"
    Else
        sqlP = " Select cRHConceptoCod Codigo, cRHConceptoDescripcion Nombre From RHConceptoTabla CO" _
             & " where cRHConceptoCod  in (Select cRHConceptoCod  from RHPlanillaConceptoTabla Where cRHPlanillaCod = '" & psRHPlanillaCod & "') And Left(cRHConceptoCod,1) = '" & psGrupo & "' And cRHConceptoEstado = '1'"
    End If
        
    If oCon.AbreConexion Then
        Set GetPlanillaConcepUsados = oCon.CargaRecordSet(sqlP)
        oCon.CierraConexion
    End If

    Set oCon = Nothing
    Exit Function
GetPlanillasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetPlanillas Method")
End Function

'Devuelve las Planillas Activas
'##ModelId=3A9BECFD01ED
Public Function GetPlanillaEstadosUsados(psRHPlanillaCod As String) As Recordset
    On Error GoTo GetPlanillasErr
    Dim sqlP As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    sqlP = " Select  nConsValor Valor, cConsDescripcion Descrip, RHP.nRHEstado Estado From Constante CO  Left Join RHPlanillaAplicacion RHP On CO.nConsValor = RHP.nRHEstado    And cRHPlanillaCod = '" & psRHPlanillaCod & "' Where nConsCod like '" & gRHEstado & "'"
        
    If oCon.AbreConexion Then
        Set GetPlanillaEstadosUsados = oCon.CargaRecordSet(sqlP)
        oCon.CierraConexion
    End If

    Set oCon = Nothing
    Exit Function
GetPlanillasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetPlanillas Method")
End Function

'Devuelve las Planillas Activas
'##ModelId=3A9BECFD01ED
Public Function AgregaPlanillaConcepto(psRHPlanillaCod As String, psRHConceptoCod As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo GetPlanillasErr
    Dim sqlP As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlP = " Insert RHPlanillaConceptoTabla (cRHPlanillaCod,cRHConceptoCod,cUltimaActualizacion)" _
         & " Values('" & psRHPlanillaCod & "','" & psRHConceptoCod & "','" & psUltimaActualizacion & "')"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlP
        oCon.CierraConexion
    End If

    Set oCon = Nothing
    Exit Function
GetPlanillasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetPlanillas Method")
End Function

'Devuelve las Planillas Activas
'##ModelId=3A9BECFD01ED
Public Function EliminaPlanillaConcepto(psRHPlanillaCod As String, psRHConceptoCod As String) As Recordset
    On Error GoTo GetPlanillasErr
    Dim sqlP As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlP = " Delete RHPlanillaConceptoTabla" _
         & " Where cRHPlanillaCod = '" & psRHPlanillaCod & "' And cRHConceptoCod = '" & psRHConceptoCod & "'"
        
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlP
        oCon.CierraConexion
    End If

    Set oCon = Nothing
    Exit Function
GetPlanillasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetPlanillas Method")
End Function

Public Function EliminaPlanillaEstadosRH(psRHPlanillaCod As String) As Recordset
    On Error GoTo GetPlanillasErr
    Dim sqlP As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlP = " Delete RHPlanillaAplicacion" _
         & " Where cRHPlanillaCod = '" & psRHPlanillaCod & "'"
        
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlP
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetPlanillasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetPlanillas Method")
End Function

Public Function AgregaPlanillaEstadosRH(psRHPlanillaCod As String, psRHEstado As String) As Recordset
    On Error GoTo GetPlanillasErr
    Dim sqlP As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlP = " Insert RHPlanillaAplicacion (cRHPlanillaCod, nRHEstado) " _
         & " Values('" & psRHPlanillaCod & "','" & psRHEstado & "')"
        
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlP
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetPlanillasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetPlanillas Method")
End Function

'Actualiza una nueva relacion entre concepto y planilla
'##ModelId=3A9B14D6011D
Public Function ModificaPlanilla(psRHPlanillaPeriodo As String, psRHPlanillaCod As String, prRS As ADODB.Recordset, psUltimaActualizacion As String, psFecImpre As String, psFecIni As String, psFecFin As String, psTipcambio As String, psDescripcion As String, psFecPago As String, psTipcambioFF As String, pnEstado As Integer) As Boolean
    On Error GoTo ModificaConceptoPlanillaErr
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlD As String
    Dim sqlDD As String
    Dim sqlDC As String
    Dim lnI As Integer
    Dim lnMontoI As Currency
    Dim lnMontoD As Currency
    
    If oCon.AbreConexion Then
        oCon.BeginTrans
            sqlD = " Delete RHplanillaDetCon where cRRHHPeriodo = '" & psRHPlanillaPeriodo & "' And cPlanillaCod = '" & psRHPlanillaCod & "'"
            sqlDD = " Delete RHPlanillaDet where cRRHHPeriodo = '" & psRHPlanillaPeriodo & "' And cPlanillaCod = '" & psRHPlanillaCod & "'"
            sqlDC = " Delete RHPlanilla where cRRHHPeriodo = '" & psRHPlanillaPeriodo & "' And cPlanillaCod = '" & psRHPlanillaCod & "'"
            
            oCon.Ejecutar sqlD
            oCon.Ejecutar sqlDD
            oCon.Ejecutar sqlDC
            
            sqlD = " Insert RHPlanilla (cRRHHPeriodo,cPlanillaCod,dRRHHPago,dRRHHInicio,dRRHHFin,dRRHHImpre,nTpoCambio,cPlaInsDes,cRHPlanillaEstado,cUltimaActualizacion,nTpoCambioFijo)" _
                 & " Values ('" & psRHPlanillaPeriodo & "','" & psRHPlanillaCod & "','" & psFecPago & "','" & psFecIni & "','" & psFecFin & "','" & psFecImpre & "'," & psTipcambio & ",'" & psDescripcion & "'," & pnEstado & ",'" & psUltimaActualizacion & "'," & psTipcambioFF & ")"
            oCon.Ejecutar sqlD
            
            While Not prRS.EOF
                If prRS.Fields(0) = 1 Then
                    sqlD = " Insert RHPlanillaDet (cRRHHPeriodo,cPlanillaCod,cPersCod,cComentario)" _
                         & " Values('" & psRHPlanillaPeriodo & "','" & psRHPlanillaCod & "','" & prRS.Fields(2) & "','" & prRS.Fields(prRS.Fields.Count - 1) & "')"
                    oCon.Ejecutar sqlD
'                    If prRS.Fields(1) = "E00025" Then
'                        MsgBox "ss"
'                    End If
                    lnMontoI = 0
                    lnMontoD = 0
                    If prRS.Fields(0) = "1" And Left(prRS.Fields(1), 1) <> "_" Then
                        For lnI = 5 To prRS.Fields.Count - 2
                            If prRS.Fields(lnI).Name = gsRHConceptoITOTING Then
                                If prRS.Fields(lnI) <> "" Then lnMontoI = prRS.Fields(lnI)
                            ElseIf prRS.Fields(lnI).Name = gsRHConceptoDTOTDES Then
                                If prRS.Fields(lnI) <> "" Then lnMontoD = prRS.Fields(lnI)
                            End If
                            If prRS.Fields(lnI) <> "" Then
                                sqlD = " Insert RHPlanillaDetCon (cRRHHPeriodo,cPlanillaCod,cPersCod,cRHConceptoCod,cDato,nMonto,bPagado)" _
                                     & " Select '" & psRHPlanillaPeriodo & "' Cod, '" & psRHPlanillaCod & "','" & prRS.Fields(2) & "', cRHConceptoCod, '" & prRS.Fields(lnI) & "', '" & prRS.Fields(lnI) & "',0 From RHConceptoTabla Where cRHconceptoMeno = '" & prRS.Fields(lnI).Name & "'"
                            
                                oCon.Ejecutar sqlD
                            End If
                            
                            'MAVM 20110715 ***
                            If prRS.Fields(lnI).Name = "U_DIAS_VACAC" Then
                                sqlD = "Update RHEmpleado Set nRHEmplVacacionesPend = nRHEmplVacacionesPend - " & IIf(prRS.Fields(lnI) = "", 0, prRS.Fields(lnI)) & " Where cPersCod = '" & prRS.Fields(2) & "'"
                                oCon.Ejecutar sqlD
                                
                                sqlD = "Insert Into MovVacaciones (cPersCod, cCodVacaciones, dFecha, cUser, nDias, cFlag, cPeriodo)" _
                                       & " Values('" & prRS.Fields(2) & "', 2,'" & Mid(psUltimaActualizacion, 5, 2) & "/" & Mid(psUltimaActualizacion, 7, 2) & "/" & Mid(psUltimaActualizacion, 1, 4) & "','" & Mid(psUltimaActualizacion, 22, 4) & "','" & IIf(prRS.Fields(lnI) = "", 0, prRS.Fields(lnI)) & "', 1, '" & Mid(psUltimaActualizacion, 3, 6) & "'" & ")"
                                oCon.Ejecutar sqlD
                                
                            End If
                            '***
                        Next lnI
                        If lnMontoI <> 0 Or lnMontoD <> 0 Then
                            sqlD = " Insert RHPlanillaDetCon (cRRHHPeriodo,cPlanillaCod,cPersCod,cRHConceptoCod,cDato,nMonto,bPagado)" _
                                 & " Select '" & psRHPlanillaPeriodo & "' Cod, '" & psRHPlanillaCod & "','" & prRS.Fields(2) & "', cRHConceptoCod, '" & lnMontoI - lnMontoD & "', '" & lnMontoI - lnMontoD & "',0 From RHConceptoTabla Where cRHconceptoMeno = '" & gsRHConceptoINETOPAGAR & "'"
                            oCon.Ejecutar sqlD
                        End If
                    End If
                End If
                prRS.MoveNext
            Wend
        oCon.CommitTrans
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
ModificaConceptoPlanillaErr:
    oCon.RollBackTrans
    oCon.CierraConexion
    Call RaiseError(MyUnhandledError, "NActualizaDatosConceptoPlanilla:ModificaConceptoPlanilla Method")
End Function

'Actualiza una nueva relacion entre concepto y planilla
'##ModelId=3A9B14D6011D
Public Function EliminaPlanilla(psRHPlanillaPeriodo As String, psRHPlanillaCod As String) As Boolean
    On Error GoTo ModificaConceptoPlanillaErr
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlD As String
    Dim sqlDD As String
    Dim sqlDC As String
    Dim lnI As Integer
    Dim lnMontoI As Currency
    Dim lnMontoD As Currency
    
    If oCon.AbreConexion Then
        oCon.BeginTrans
            sqlD = " Delete RHplanillaDetCon where cRRHHPeriodo = '" & psRHPlanillaPeriodo & "' And cPlanillaCod = '" & psRHPlanillaCod & "'"
            sqlDD = " Delete RHPlanillaDet where cRRHHPeriodo = '" & psRHPlanillaPeriodo & "' And cPlanillaCod = '" & psRHPlanillaCod & "'"
            sqlDC = " Delete RHPlanilla where cRRHHPeriodo = '" & psRHPlanillaPeriodo & "' And cPlanillaCod = '" & psRHPlanillaCod & "'"
            oCon.Ejecutar sqlD
            oCon.Ejecutar sqlDD
            oCon.Ejecutar sqlDC
        oCon.CommitTrans
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
ModificaConceptoPlanillaErr:
    oCon.RollBackTrans
    oCon.CierraConexion
    Call RaiseError(MyUnhandledError, "NActualizaDatosConceptoPlanilla:ModificaConceptoPlanilla Method")
End Function

'Public Function GetBoleta(psPersCod As String, psPlanillaPeriodo As String, psPlanillaCod As String) As ADODB.Recordset
'    Dim oCon As DConecta
'    Dim sqlP As String
'    Set oCon = New DConecta
'
'    sqlP = " Select distinct EM.cRHCod,  PE.cPersNombre, left(PD.cRHConceptoCod ,1), nRHConceptoOrden, PD.cPlanillaCod, CO.cRHconceptoMeno cRHConceptoCod,PD.nMonto,CO.cRHConceptoAbreviatura, CO.cRHConceptoDescripcion, CO.cRHConceptoCod Concepcod, " _
'         & " isnull(RHAFP.cRHAFPAbreviatura,'')cRHAFPAbreviatura, Case RHE.bRHTipoFondo When 0 Then 'AFP' Else 'SNP' End TipoFondo , EM.nRHEstado," _
'         & " isnull(RHE.cCUSPP,'')cCUSPP,isnull(RHE.cNumAFP,'')cNumAFP, EM.cAgenciaActual, isnull(RHE.cRHEmpNivRem,'')cRHEmpNivRem," _
'         & " EM.dIngreso , RHC.nRHContratoTpo,IsNull((Select cPersIDnro From PersId Where cPersIDTpo = " & PersIdTipo.gPersIdDNI & " And cPersCod = '" & psPersCod & "'),'') DNI " _
'         & " From RHPlanillaDetCon PD" _
'         & " Inner Join RHConceptoTabla CO On PD.cRHConceptoCod = CO.cRHConceptoCod" _
'         & " Inner Join RRHH EM On EM.cPersCod = PD.cPersCod" _
'         & " Inner Join Persona PE On EM.cPersCod = PE.cPersCod" _
'         & " Inner Join RHContrato RHC On RHC.cPerscod = EM.cPersCod" _
'         & " Left  Join RHEmpleado RHE On EM.cPersCod = RHE.cPersCod" _
'         & " Left  Join RHAFP RHAFP On RHAFP.cRHAFPPersCod = RHE.cRHEmplAFPPersCod" _
'         & " Where EM.cPersCod = '" & psPersCod & "' And cRRHHPeriodo = '" & psPlanillaPeriodo & "'" _
'         & " And PD.cRHConceptoCod <> '" & gsRHConceptoINETOPAGARCOD & "' and cPlanillaCod = '" & psPlanillaCod & "'" _
'         & " And RHC.cRHContratoNro in (Select Max(cRHContratoNro) From RHContrato RHC1 Where RHC1.cPersCod = RHC.cPersCod)" _
'         & " Order by left(PD.cRHConceptoCod ,1) Desc, nRHConceptoOrden"
'    If oCon.AbreConexion Then
'        Set GetBoleta = oCon.CargaRecordSet(sqlP)
'        oCon.CierraConexion
'    End If
'    Set oCon = Nothing
'End Function


Public Function GetBoleta(psPersCod As String, psPlanillaPeriodo As String, psPlanillaCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlP As String
    Set oCon = New DConecta
    
    If psPlanillaCod = "E01" And Len(psPlanillaPeriodo) < 15 Then
                sqlP = "spRHconsolidaBoleta '" & psPersCod & "','" & psPlanillaPeriodo & "','" & psPlanillaCod & "'"
    ElseIf psPlanillaCod = "E01" And Len(psPlanillaPeriodo) > 15 Then
                sqlP = " spRHconsolidaBoleta_new '" & psPersCod & "','" & Left(psPlanillaPeriodo, 6) & "','" & psPlanillaCod & "'"
    Else
                sqlP = " Select distinct EM.cRHCod,  PE.cPersNombre, left(PD.cRHConceptoCod ,1), nRHConceptoOrden, PD.cPlanillaCod, CO.cRHconceptoMeno cRHConceptoCod,PD.nMonto,CO.cRHConceptoAbreviatura, CO.cRHConceptoDescripcion, CO.cRHConceptoCod Concepcod, " _
                 & " isnull(RHAFP.cRHAFPAbreviatura,'')cRHAFPAbreviatura, Case RHE.bRHTipoFondo When 0 Then 'AFP' Else 'SNP' End TipoFondo , EM.nRHEstado," _
                 & " RHE.cCUSPP,RHE.cNumAFP, EM.cAgenciaActual, RHE.cRHEmpNivRem," _
                 & " EM.dIngreso , RHC.nRHContratoTpo,IsNull((Select cPersIDnro From PersId Where cPersIDTpo = " & PersIdTipo.gPersIdDNI & " And cPersCod = '" & psPersCod & "'),'') DNI" _
                 & " From RHPlanillaDetCon PD" _
                 & " Inner Join RHConceptoTabla CO On PD.cRHConceptoCod = CO.cRHConceptoCod" _
                 & " Inner Join RRHH EM On EM.cPersCod = PD.cPersCod" _
                 & " Inner Join Persona PE On EM.cPersCod = PE.cPersCod" _
                 & " Inner Join RHContrato RHC On RHC.cPerscod = EM.cPersCod" _
                 & " Left  Join RHEmpleado RHE On EM.cPersCod = RHE.cPersCod" _
                 & " Left  Join RHAFP RHAFP On RHAFP.cRHAFPPersCod = RHE.cRHEmplAFPPersCod" _
                 & " Where EM.cPersCod = '" & psPersCod & "' And cRRHHPeriodo = '" & psPlanillaPeriodo & "'" _
                 & " And PD.cRHConceptoCod <> '" & gsRHConceptoINETOPAGARCOD & "' and cPlanillaCod = '" & psPlanillaCod & "'" _
                 & " And RHC.cRHContratoNro in (Select Max(cRHContratoNro) From RHContrato RHC1 Where RHC1.cPersCod = RHC.cPersCod)" _
                 & " Order by left(PD.cRHConceptoCod ,1) Desc, nRHConceptoOrden"
                 
    End If
    If oCon.AbreConexion Then
        Set GetBoleta = oCon.CargaRecordSet(sqlP)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function




    
'Public Function GetItemPlanilla(psPersCod As String, psPlanillaPeriodo As String, psPlanillaCod As String, psCodAge As String, pbTotal As Boolean, pnSubTotalEmpleados As Integer) As ADODB.Recordset
'    Dim oCon As DConecta
'    Dim sqlP As String
'    Set oCon = New DConecta
'
'    Dim ldFecha As Date
'
'    ldFecha = CDate(Right(psPlanillaPeriodo, 2) & "/" & Mid(psPlanillaPeriodo, 5, 2) & "/" & Left(psPlanillaPeriodo, 4))
'
'    If Not pbTotal Then
'        sqlP = " Select EM.cRHCod, PE.cPersNombre, CO.cRHConceptoMeno cConcepCod, PD.nMonto, CO.cRHConceptoAbreviatura," _
'             & "        Afp.cRHAFPAbreviatura, Case EML.bRHTipoFondo When 0 Then 'AFP' Else 'SNP' End cEmpTipFondo, EM.dIngreso, EML.dEstabilidad," _
'             & "        EML.cCUSPP, EML.cNumAFP cEmpNumAFP, dbo.GetDRHCargos3(EM.cPersCod, '" & Format(ldFecha, gsFormatoFecha) & "') cAgenciaAsig, EML.cRHEmpNivRem, EM.cAreaCodActual, (Select cAreaDescripcion from areas where cAreaEstruc In (select Case When len(cAreaEstruc) > 6 Then left(cAreaEstruc,6) Else '0110' End from areas where careacod = EM.cAreaCodActual))  cAreaDescripcion, bConfianza, bNoSujFiscalizacion, bPersDireccion" _
'             & "        From RHPlanillaDet PDD" _
'             & " Left  Join RHPlanillaDetCon PD On PD.cRRHHPeriodo = PDD.cRRHHPeriodo And PD.cPlanillaCod = PDD.cPlanillaCod And PD.cPersCod = PDD.cPersCod " _
'             & " Left  Join RHEmpleado EML On EML.cPersCod = PDD.cPersCod" _
'             & " Left  Join RHConceptoTabla CO On PD.cRHConceptoCod = CO.cRHConceptoCod" _
'             & " Inner Join RRHH EM On EM.cPersCod = PDD.cPersCod" _
'             & " Inner Join Areas Are On EM.cAreaCodActual = Are.cAreaCod" _
'             & " Left  Join RHAFP Afp On EML.cRHEmplAFPPersCod = Afp.cRHAFPPersCod" _
'             & " Inner Join Persona PE On EM.cPersCod = PE.cPersCod" _
'             & " Where EM.cPersCod = '" & psPersCod & "' and  PDD.cRRHHPeriodo = '" & psPlanillaPeriodo & "' and PDD.cPlanillaCod = '" & psPlanillaCod & "' AND (CO.cRHConceptoCod <> '" & gsRHConceptoINETOPAGARCOD & "' Or CO.cRHConceptoCod Is Null)" _
'             & " Order By Left(CO.cRHConceptoCod,1) Desc, CO.nRHConceptoOrden"
'    Else
'        If pnSubTotalEmpleados = 0 Then
'            If psCodAge = "" Then
'                sqlP = " Select Count(*) Num, CO.cRHconceptoMeno cConcepCod,nMonto= Sum(PD.nMonto),CO.cRHConceptoAbreviatura" _
'                     & "        From RHPlanillaDetCon PD" _
'                     & " Inner Join RHConceptoTabla CO On PD.cRHConceptoCod = CO.cRHConceptoCod" _
'                     & " Inner Join RRHH EM On EM.cPersCod = PD.cPersCod" _
'                     & " Where cRRHHPeriodo = '" & psPlanillaPeriodo & "' and PD.cPlanillaCod = '" & psPlanillaCod & "' AND PD.cRHConceptoCod <> '" & gsRHConceptoINETOPAGARCOD & "'" _
'                     & " Group By CO.cRHconceptoMeno , CO.cRHConceptoAbreviatura, CO.nRHConceptoOrden" _
'                     & " Order by Left(CO.cRHconceptoMeno,1) Desc, CO.nRHConceptoOrden"
'            Else
'                sqlP = " Select Count(*) Num, CO.cRHconceptoMeno cConcepCod,nMonto= Sum(PD.nMonto),CO.cRHConceptoAbreviatura" _
'                     & "        From RHPlanillaDetCon PD" _
'                     & " Inner Join RHConceptoTabla CO On PD.cRHConceptoCod = CO.cRHConceptoCod" _
'                     & " Inner Join RRHH EM On EM.cPersCod = PD.cPersCod" _
'                     & " Where cRRHHPeriodo = '" & psPlanillaPeriodo & "' and PD.cPlanillaCod = '" & psPlanillaCod & "' AND PD.cRHConceptoCod <> '" & gsRHConceptoINETOPAGARCOD & "'" _
'                     & " And dbo.GetDRHCargos3(EM.cPersCod, '" & Format(ldFecha, gsFormatoFecha) & "') = '" & psCodAge & "'" _
'                     & " Group By CO.cRHConceptoMeno, CO.cRHConceptoAbreviatura, CO.nRHConceptoOrden" _
'                     & " Order by Left(CO.cRHconceptoMeno,1) Desc, CO.nRHConceptoOrden"
'            End If
'        ElseIf pnSubTotalEmpleados = 1 Then
'            If psCodAge = "" Then
'                sqlP = " Select Count(*) Num, CO.nRHConceptoOrden, CO.cRHconceptoMeno cConcepCod,nMonto= Sum(PD.nMonto),CO.cRHConceptoAbreviatura" _
'                     & "        From RHPlanillaDetCon PD" _
'                     & " Inner Join RHConceptoTabla CO On PD.cRHConceptoCod = CO.cRHConceptoCod" _
'                     & " Inner Join RRHH EM On EM.cPersCod = PD.cPersCod" _
'                     & " Inner Join RHContrato RHC On EM.cPersCod = RHC.cPersCod" _
'                     & " Where cRRHHPeriodo = '" & psPlanillaPeriodo & "' and PD.cPlanillaCod = '" & psPlanillaCod & "' AND PD.cRHConceptoCod <> '" & gsRHConceptoINETOPAGARCOD & "'" _
'                     & " And RHC.nRHContratoTpo = '0' And RHC.cRHContratoNro In (Select Max(cRHContratoNro) From RHContrato RHC1 Where RHC1.cPersCod = RHC.cPersCod) " _
'                     & " Group By CO.cRHConceptoMeno, CO.cRHConceptoAbreviatura, CO.nRHConceptoOrden" _
'                     & " Order by Left(CO.cRHconceptoMeno,1) Desc, CO.nRHConceptoOrden"
'            Else
'                sqlP = " Select Count(*) Num, CO.cRHconceptoMeno cConcepCod,nMonto= Sum(PD.nMonto),CO.cRHConceptoAbreviatura" _
'                     & "        From RHPlanillaDetCon PD" _
'                     & " Inner Join RHConceptoTabla CO On PD.cRHConceptoCod = CO.cRHConceptoCod" _
'                     & " Inner Join RRHH EM On EM.cPersCod = PD.cPersCod" _
'                     & " Inner Join RHContrato RHC On EM.cPersCod = RHC.cPersCod" _
'                     & " Where cRRHHPeriodo = '" & psPlanillaPeriodo & "' and PD.cPlanillaCod = '" & psPlanillaCod & "' AND PD.cRHConceptoCod <> '" & gsRHConceptoINETOPAGARCOD & "'" _
'                     & " And RHC.nRHContratoTpo = '0' And RHC.cRHContratoNro In (Select Max(cRHContratoNro) From RHContrato RHC1 Where RHC1.cPersCod = RHC.cPersCod And Left(cUltimaActualizacion,8) <= '" & Format(ldFecha, gsFormatoMovFecha) & "') " _
'                     & " And dbo.GetDRHCargos3(EM.cPersCod, '" & Format(ldFecha, gsFormatoFecha) & "') = '" & psCodAge & "'" _
'                     & " Group By CO.cRHConceptoMeno, CO.cRHConceptoAbreviatura, CO.nRHConceptoOrden " _
'                     & " Order by Left(CO.cRHconceptoMeno,1) Desc, CO.nRHConceptoOrden"
'            End If
'        ElseIf pnSubTotalEmpleados = 2 Then
'            If psCodAge = "" Then
'                sqlP = " Select Count(*) Num, CO.cRHconceptoMeno cConcepCod,nMonto= Sum(PD.nMonto),CO.cRHConceptoAbreviatura" _
'                     & "        From RHPlanillaDetCon PD" _
'                     & " Inner Join RHConceptoTabla CO On PD.cRHConceptoCod = CO.cRHConceptoCod" _
'                     & " Inner Join RRHH EM On EM.cPersCod = PD.cPersCod" _
'                     & " Inner Join RHContrato RHC On EM.cPersCod = RHC.cPersCod" _
'                     & " Where cRRHHPeriodo = '" & psPlanillaPeriodo & "' and PD.cPlanillaCod = '" & psPlanillaCod & "' AND PD.cRHConceptoCod <> '" & gsRHConceptoINETOPAGARCOD & "'" _
'                     & " And RHC.nRHContratoTpo <> '0' And RHC.cRHContratoNro In (Select Max(cRHContratoNro) From RHContrato RHC1 Where RHC1.cPersCod = RHC.cPersCod And Left(cUltimaActualizacion,8) <= '" & Format(ldFecha, gsFormatoMovFecha) & "') " _
'                     & " Group By CO.cRHConceptoMeno, CO.cRHConceptoAbreviatura,CO.nRHConceptoOrden " _
'                     & " Order by Left(CO.cRHconceptoMeno,1) Desc, CO.nRHConceptoOrden"
'            Else
'                sqlP = " Select Count(*) Num, CO.cRHconceptoMeno cConcepCod,nMonto= Sum(PD.nMonto),CO.cRHConceptoAbreviatura" _
'                     & "        From RHPlanillaDetCon PD" _
'                     & " Inner Join RHConceptoTabla CO On PD.cRHConceptoCod = CO.cRHConceptoCod" _
'                     & " Inner Join RRHH EM On EM.cPersCod = PD.cPersCod" _
'                     & " Left Join RHContrato RHC On EM.cPersCod = RHC.cPersCod" _
'                     & " Where cRRHHPeriodo = '" & psPlanillaPeriodo & "' and PD.cPlanillaCod = '" & psPlanillaCod & "' AND PD.cRHConceptoCod <> '" & gsRHConceptoINETOPAGARCOD & "'" _
'                     & " And RHC.nRHContratoTpo <> '0' And RHC.cRHContratoNro In (Select Max(cRHContratoNro) From RHContrato RHC1 Where RHC1.cPersCod = RHC.cPersCod And Left(cUltimaActualizacion,8) <= '" & Format(ldFecha, gsFormatoMovFecha) & "' ) " _
'                     & " And dbo.GetDRHCargos3(EM.cPersCod, '" & Format(ldFecha, gsFormatoFecha) & "') = '" & psCodAge & "'" _
'                     & " Group By CO.cRHConceptoMeno, CO.cRHConceptoAbreviatura,CO.nRHConceptoOrden" _
'                     & " Order by Left(CO.cRHconceptoMeno,1) Desc, CO.nRHConceptoOrden"
'            End If
'        End If
'    End If
'
'    If oCon.AbreConexion Then
'        Set GetItemPlanilla = oCon.CargaRecordSet(sqlP)
'        oCon.CierraConexion
'    End If
'    Set oCon = Nothing
'End Function
    
Public Function GetItemPlanilla(psPersCod As String, psPlanillaPeriodo As String, psPlanillaCod As String, psCodAge As String, pbTotal As Boolean, pnSubTotalEmpleados As Integer) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlP As String
    Set oCon = New DConecta
    
    If Not pbTotal Then
        sqlP = " Select EM.cRHCod, PE.cPersNombre, CO.cRHConceptoMeno cConcepCod, PD.nMonto, CO.cRHConceptoAbreviatura," _
             & "        Afp.cRHAFPAbreviatura, Case EML.bRHTipoFondo When 0 Then 'AFP' Else 'SNP' End cEmpTipFondo, EM.dIngreso, EML.dEstabilidad," _
             & "        EML.cCUSPP, EML.cNumAFP cEmpNumAFP, EM.cAgenciaAsig, EML.cRHEmpNivRem, EM.cAreaCodActual, (Select cAreaDescripcion from areas where cAreaEstruc In (select Case When len(cAreaEstruc) > 6 Then left(cAreaEstruc,6) Else '0110' End from areas where careacod = EM.cAreaCodActual))  cAreaDescripcion" _
             & "        From RHPlanillaDet PDD" _
             & " Left  Join RHPlanillaDetCon PD On PD.cRRHHPeriodo = PDD.cRRHHPeriodo And PD.cPlanillaCod = PDD.cPlanillaCod And PD.cPersCod = PDD.cPersCod " _
             & " Left  Join RHEmpleado EML On EML.cPersCod = PDD.cPersCod" _
             & " Left  Join RHConceptoTabla CO On PD.cRHConceptoCod = CO.cRHConceptoCod" _
             & " Inner Join RRHH EM On EM.cPersCod = PDD.cPersCod" _
             & " Inner Join Areas Are On EM.cAreaCodActual = Are.cAreaCod" _
             & " Left  Join RHAFP Afp On EML.cRHEmplAFPPersCod = Afp.cRHAFPPersCod" _
             & " Inner Join Persona PE On EM.cPersCod = PE.cPersCod" _
             & " Where EM.cPersCod = '" & psPersCod & "' and  PDD.cRRHHPeriodo = '" & psPlanillaPeriodo & "' and PDD.cPlanillaCod = '" & psPlanillaCod & "' AND (CO.cRHConceptoCod <> '" & gsRHConceptoINETOPAGARCOD & "' Or CO.cRHConceptoCod Is Null)" _
             & " Order By Left(CO.cRHConceptoCod,1) Desc, CO.nRHConceptoOrden"
    Else
        If pnSubTotalEmpleados = 0 Then
            If psCodAge = "" Then
                sqlP = " Select Count(*) Num, CO.cRHconceptoMeno cConcepCod,nMonto= Sum(PD.nMonto),CO.cRHConceptoAbreviatura" _
                     & "        From RHPlanillaDetCon PD" _
                     & " Inner Join RHConceptoTabla CO On PD.cRHConceptoCod = CO.cRHConceptoCod" _
                     & " Inner Join RRHH EM On EM.cPersCod = PD.cPersCod" _
                     & " Where cRRHHPeriodo = '" & psPlanillaPeriodo & "' and PD.cPlanillaCod = '" & psPlanillaCod & "' AND PD.cRHConceptoCod <> '" & gsRHConceptoINETOPAGARCOD & "'" _
                     & " Group By CO.cRHconceptoMeno , CO.cRHConceptoAbreviatura, CO.nRHConceptoOrden" _
                     & " Order by Left(CO.cRHconceptoMeno,1) Desc, CO.nRHConceptoOrden"
            Else
                sqlP = " Select Count(*) Num, CO.cRHconceptoMeno cConcepCod,nMonto= Sum(PD.nMonto),CO.cRHConceptoAbreviatura" _
                     & "        From RHPlanillaDetCon PD" _
                     & " Inner Join RHConceptoTabla CO On PD.cRHConceptoCod = CO.cRHConceptoCod" _
                     & " Inner Join RRHH EM On EM.cPersCod = PD.cPersCod" _
                     & " Where cRRHHPeriodo = '" & psPlanillaPeriodo & "' and PD.cPlanillaCod = '" & psPlanillaCod & "' AND PD.cRHConceptoCod <> '" & gsRHConceptoINETOPAGARCOD & "'" _
                     & " And EM.cAgenciaAsig = '" & psCodAge & "'" _
                     & " Group By CO.cRHConceptoMeno, CO.cRHConceptoAbreviatura, CO.nRHConceptoOrden" _
                     & " Order by Left(CO.cRHconceptoMeno,1) Desc, CO.nRHConceptoOrden"
            End If
        ElseIf pnSubTotalEmpleados = 1 Then
            If psCodAge = "" Then
                sqlP = " Select Count(*) Num, CO.nRHConceptoOrden, CO.cRHconceptoMeno cConcepCod,nMonto= Sum(PD.nMonto),CO.cRHConceptoAbreviatura" _
                     & "        From RHPlanillaDetCon PD" _
                     & " Inner Join RHConceptoTabla CO On PD.cRHConceptoCod = CO.cRHConceptoCod" _
                     & " Inner Join RRHH EM On EM.cPersCod = PD.cPersCod" _
                     & " Inner Join RHContrato RHC On EM.cPersCod = RHC.cPersCod" _
                     & " Where cRRHHPeriodo = '" & psPlanillaPeriodo & "' and PD.cPlanillaCod = '" & psPlanillaCod & "' AND PD.cRHConceptoCod <> '" & gsRHConceptoINETOPAGARCOD & "'" _
                     & " And RHC.nRHContratoTpo = '0' And RHC.cRHContratoNro In (Select Max(cRHContratoNro) From RHContrato RHC1 Where RHC1.cPersCod = RHC.cPersCod) " _
                     & " Group By CO.cRHConceptoMeno, CO.cRHConceptoAbreviatura, CO.nRHConceptoOrden" _
                     & " Order by Left(CO.cRHconceptoMeno,1) Desc, CO.nRHConceptoOrden"
            Else
                sqlP = " Select Count(*) Num, CO.cRHconceptoMeno cConcepCod,nMonto= Sum(PD.nMonto),CO.cRHConceptoAbreviatura" _
                     & "        From RHPlanillaDetCon PD" _
                     & " Inner Join RHConceptoTabla CO On PD.cRHConceptoCod = CO.cRHConceptoCod" _
                     & " Inner Join RRHH EM On EM.cPersCod = PD.cPersCod" _
                     & " Inner Join RHContrato RHC On EM.cPersCod = RHC.cPersCod" _
                     & " Where cRRHHPeriodo = '" & psPlanillaPeriodo & "' and PD.cPlanillaCod = '" & psPlanillaCod & "' AND PD.cRHConceptoCod <> '" & gsRHConceptoINETOPAGARCOD & "'" _
                     & " And RHC.nRHContratoTpo = '0' And RHC.cRHContratoNro In (Select Max(cRHContratoNro) From RHContrato RHC1 Where RHC1.cPersCod = RHC.cPersCod) " _
                     & " And EM.cAgenciaAsig = '" & psCodAge & "'" _
                     & " Group By CO.cRHConceptoMeno, CO.cRHConceptoAbreviatura, CO.nRHConceptoOrden " _
                     & " Order by Left(CO.cRHconceptoMeno,1) Desc, CO.nRHConceptoOrden"
            End If
        ElseIf pnSubTotalEmpleados = 2 Then
            If psCodAge = "" Then
                sqlP = " Select Count(*) Num, CO.cRHconceptoMeno cConcepCod,nMonto= Sum(PD.nMonto),CO.cRHConceptoAbreviatura" _
                     & "        From RHPlanillaDetCon PD" _
                     & " Inner Join RHConceptoTabla CO On PD.cRHConceptoCod = CO.cRHConceptoCod" _
                     & " Inner Join RRHH EM On EM.cPersCod = PD.cPersCod" _
                     & " Inner Join RHContrato RHC On EM.cPersCod = RHC.cPersCod" _
                     & " Where cRRHHPeriodo = '" & psPlanillaPeriodo & "' and PD.cPlanillaCod = '" & psPlanillaCod & "' AND PD.cRHConceptoCod <> '" & gsRHConceptoINETOPAGARCOD & "'" _
                     & " And RHC.nRHContratoTpo <> '0' And RHC.cRHContratoNro In (Select Max(cRHContratoNro) From RHContrato RHC1 Where RHC1.cPersCod = RHC.cPersCod) " _
                     & " Group By CO.cRHConceptoMeno, CO.cRHConceptoAbreviatura,CO.nRHConceptoOrden " _
                     & " Order by Left(CO.cRHconceptoMeno,1) Desc, CO.nRHConceptoOrden"
            Else
                sqlP = " Select Count(*) Num, CO.cRHconceptoMeno cConcepCod,nMonto= Sum(PD.nMonto),CO.cRHConceptoAbreviatura" _
                     & "        From RHPlanillaDetCon PD" _
                     & " Inner Join RHConceptoTabla CO On PD.cRHConceptoCod = CO.cRHConceptoCod" _
                     & " Inner Join RRHH EM On EM.cPersCod = PD.cPersCod" _
                     & " Inner Join RHContrato RHC On EM.cPersCod = RHC.cPersCod" _
                     & " Where cRRHHPeriodo = '" & psPlanillaPeriodo & "' and PD.cPlanillaCod = '" & psPlanillaCod & "' AND PD.cRHConceptoCod <> '" & gsRHConceptoINETOPAGARCOD & "'" _
                     & " And RHC.nRHContratoTpo <> '0' And RHC.cRHContratoNro In (Select Max(cRHContratoNro) From RHContrato RHC1 Where RHC1.cPersCod = RHC.cPersCod) " _
                     & " And EM.cAgenciaAsig = '" & psCodAge & "'" _
                     & " Group By CO.cRHConceptoMeno, CO.cRHConceptoAbreviatura,CO.nRHConceptoOrden" _
                     & " Order by Left(CO.cRHconceptoMeno,1) Desc, CO.nRHConceptoOrden"
            End If
        End If
    End If
    
    If oCon.AbreConexion Then
        Set GetItemPlanilla = oCon.CargaRecordSet(sqlP)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function

Public Function GetItemPlanillaRep(psPersCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlP As String
    Set oCon = New DConecta
        
    sqlP = " Select -1 Num, PE.cPersNombre, Afp.cRHAFPAbreviatura, Case EML.bRHTipoFondo When 0 Then 'AFP' Else 'SNP' End cEmpTipFondo" _
         & " , EM.dIngreso, EML.dEstabilidad," _
         & " EML.cCUSPP, EML.cNumAFP cEmpNumAFP, EM.cAgenciaAsig, EML.cRHEmpNivRem, EM.cAreaCodActual, ARE.cAreaDescripcion " _
         & " From Persona PE" _
         & " Left Join RRHH EM On EM.cPersCod = PE.cPersCod" _
         & " Left Join RHEmpleado EML On EML.cPersCod = EM.cPersCod" _
         & " Inner Join Areas Are On EM.cAreaCodActual = Are.cAreaCod Left Join RHAFP afp On EML.cRHEmplAFPPersCod = afp.cRHAFPPersCod" _
         & " Where EM.cPErsCod = '" & psPersCod & "'"
             
    If oCon.AbreConexion Then
        Set GetItemPlanillaRep = oCon.CargaRecordSet(sqlP)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function
    
'Public Function GetBoletas(psPlanillaPeriodo As String, psPlanillaCod As String, psEmpLista As String) As ADODB.Recordset
'    Dim oCon As DConecta
'    Dim sqlC As String
'    Set oCon = New DConecta
'
'    sqlC = " Select EM.cPersCod From RRHH EM" _
'         & " Where EM.cPersCod in (" _
'         & "                        Select distinct PD.cPersCod From RHPlanillaDetCon PD" _
'         & "                        Where PD.cRRHHPeriodo = '" & psPlanillaPeriodo & "'" _
'         & "                        and PD.cPlanillaCod = '" & psPlanillaCod & "' And PD.cPersCod Not In (" & psEmpLista & ")" _
'         & "                     )" _
'         & " Order By EM.cAgenciaActual,EM.cRHCod"
'
'    If oCon.AbreConexion Then
'        Set GetBoletas = oCon.CargaRecordSet(sqlC)
'        oCon.CierraConexion
'    End If
'    Set oCon = Nothing
'End Function

Public Function GetBoletas(psPlanillaPeriodo As String, psPlanillaCod As String, psEmpLista As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    Dim psAgencia As String
    Dim psPeriodo As String
    
    If Len(psPlanillaPeriodo) > 15 Then
        psAgencia = Right(psPlanillaPeriodo, 2)
        psPeriodo = Left(psPlanillaPeriodo, 6)
        
        sqlC = " SELECT  DISTINCT RH.CPERSCOD ,cPersNombre,cRHAgenciaCodOficial " _
            & "  FROM RHPLANILLADETCON RH,PERSONA P,RHCARGOS RH1,AGENCIAS AG " _
            & " WHERE  AG.cAgeCod =RH1.cRHAgenciaCodOficial AND  RH.CPERSCOD =RH1.CPERSCOD AND " _
            & " P.CPERSCOD = RH.CPERSCOD AND dRHCargoFecha IN " _
            & " (SELECT max(dRHCargoFecha) FROM  RHCARGOS RH2 WHERE RH1.CPERSCOD = RH2.CPERSCOD and " _
            & " convert(char(4),year(dRHCargoFecha)) + RIGHT('00'+convert(varchar(2),month(dRHCargoFecha)),2) <= " & psPeriodo & " ) and " _
            & " SUBSTRING (cRRHHPeriodo,1,6) = '" & psPeriodo & "'  and P.cPersCod Not In (" & psEmpLista & ") and cPlanillaCod like 'E%'  "
        If psAgencia <> "--" Then
            sqlC = sqlC + " and cRHAgenciaCodOficial  ='" & psAgencia & "'"
        End If
            sqlC = sqlC + " ORDER BY cPersNombre ASC "
        
    
    Else
    
    sqlC = " Select EM.cPersCod,EM.cAgenciaActual From RRHH EM" _
         & " Where EM.cPersCod in (" _
         & "                        Select distinct PD.cPersCod From RHPlanillaDetCon PD" _
         & "                        Where PD.cRRHHPeriodo = '" & psPlanillaPeriodo & "'" _
         & "                        and PD.cPlanillaCod = '" & psPlanillaCod & "' And PD.cPersCod Not In (" & psEmpLista & ")" _
         & "                     )" _
         & " Order By EM.cAgenciaActual,EM.cRHCod"
         
    End If
    
    If oCon.AbreConexion Then
        Set GetBoletas = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function



Public Function GetPlanillasOpeCon() As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    sqlC = " Select cOpeCodEst, cRHPlanillaDescripcion + ' ESTABLES' Descrip, Len(cOpeCodEst) Nivel From rhplanillatabla" _
         & " Where cRHPlanillaCod Not in (Select cRHPlanillaCod From rhplanillatabla Where cOpeCodCont = cOpeCodEst)" _
         & " And cOpeCodEst Is Not Null" _
         & " Union" _
         & " Select cOpeCodCont, cRHPlanillaDescripcion + ' CONTRATADO' Descrip, Len(cOpeCodCont) Nivel from rhplanillatabla" _
         & " where cRHPlanillaCod Not in (Select cRHPlanillaCod From rhplanillatabla Where cOpeCodCont = cOpeCodEst)" _
         & " And cOpeCodCont Is Not Null" _
         & " Union" _
         & " Select cOpeCodCont, cRHPlanillaDescripcion + ' ESTABLES Y CONTRATADO' Descrip, Len(cOpeCodCont) Nivel from rhplanillatabla" _
         & " Where cRHPlanillaCod in (Select cRHPlanillaCod From rhplanillatabla Where cOpeCodCont = cOpeCodEst)" _
         & " Union" _
         & " Select cOpeCod, cOpeDesc Descrip, Len(cOpeCod) Nivel from OpeTpo" _
         & " Where copecod Like '6299%' and copecod not like '62999_'"
    
    If oCon.AbreConexion Then
        Set GetPlanillasOpeCon = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function

Public Function GetConceptoTablaArbol() As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    sqlC = " Select RHCT.cRHConceptoCod, RHCT.cRHConceptoDescripcion, Len(RHCT.cRHConceptoCod) Nivel From RHConceptoTabla RHCT" _
         & " order by RHCT.cRHConceptoCod"
    
    If oCon.AbreConexion Then
        Set GetConceptoTablaArbol = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function

Public Function GetConceptoCta(psOpeCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    sqlC = " Select RHCT.cRHConceptoCod, RHCT.cRHConceptoDescripcion," _
         & " Case RHC.cOpeCtaDH When 'D' Then RHC.cCtaContCod Else '' End Debe," _
         & " Case RHC.cOpeCtaDH When 'H' Then RHC.cCtaContCod Else '' End Haber" _
         & " From RHConceptoCta RHC" _
         & " Inner Join RHConceptoTabla RHCT On RHC.cRHConceptoCod = RHCT.cRHConceptoCod" _
         & " Where cOpeCod  = '" & psOpeCod & "'"
    
    If oCon.AbreConexion Then
        Set GetConceptoCta = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function

Public Function SetConceptoCta(psOpeCod As String, prRS As ADODB.Recordset) As Boolean
    Dim oCon As DConecta
    Dim sqlD As String
    Dim sqlH As String
    Dim sqlC As String
    
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    If prRS Is Nothing Then
        sqlC = "Delete RHConceptoCta Where cOpeCod = '" & psOpeCod & "'"
        oCon.Ejecutar sqlC
        Exit Function
    Else
        prRS.MoveFirst
        
        While Not prRS.EOF
            If prRS.Fields(2) <> "" Then
                If sqlD = "" Then
                    sqlD = prRS.Fields(0)
                Else
                    sqlD = sqlD & "','" & prRS.Fields(0)
                End If
            Else
                If sqlH = "" Then
                    sqlH = prRS.Fields(0)
                Else
                    sqlH = sqlH & "','" & prRS.Fields(0)
                End If
            End If
            prRS.MoveNext
        Wend
        
        sqlC = " Delete RHConceptoCta Where cOpeCod = '" & psOpeCod & "' And " _
             & " ( (cRHConceptoCod Not In ('" & sqlD & "') And cOpeCtaDH = 'D' ) Or (cRHConceptoCod Not In ('" & sqlH & "') And cOpeCtaDH = 'H' ) )"
        oCon.Ejecutar sqlC
        
    End If
    
    prRS.MoveFirst
    
    While Not prRS.EOF
            If prRS.Fields(2) <> "" Then
                If ExisteRHConceptoCta(psOpeCod, prRS.Fields(0), "D") Then
                    sqlD = " Update RHConceptoCta" _
                         & " Set cCtaContCod = '" & prRS.Fields(2) & "'" _
                         & " Where cOpeCod = '" & psOpeCod & "' And cRHConceptoCod = '" & prRS.Fields(0) & "' And cOpeCtaDH = 'D'"
                Else
                    sqlD = " Insert RHConceptoCta (cOpeCod, cRHConceptoCod, cCtaContCod, cOpeCtaDH)" _
                         & " Values('" & psOpeCod & "','" & prRS.Fields(0) & "','" & prRS.Fields(2) & "','D')"
                End If
            Else
                If ExisteRHConceptoCta(psOpeCod, prRS.Fields(0), "H") Then
                    sqlD = " Update RHConceptoCta" _
                         & " Set cCtaContCod = '" & prRS.Fields(3) & "'" _
                         & " Where cOpeCod = '" & psOpeCod & "' And cRHConceptoCod = '" & prRS.Fields(0) & "' And cOpeCtaDH = 'H'"
                Else
                    sqlD = " Insert RHConceptoCta (cOpeCod, cRHConceptoCod, cCtaContCod, cOpeCtaDH)" _
                         & " Values('" & psOpeCod & "','" & prRS.Fields(0) & "','" & prRS.Fields(2) & "','H')"
                End If
            End If
            oCon.Ejecutar sqlD
        prRS.MoveNext
    Wend
    
    Set oCon = Nothing
End Function

Public Function ExisteRHConceptoCta(psOpeCod As String, psConcepCod As String, psDH As String) As Boolean
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sqlC = " Select cOpeCod From RHConceptoCta" _
         & " Where cOpeCod = '" & psOpeCod & "' And cRHConceptoCod = '" & psConcepCod & "' And cOpeCtaDH = '" & psDH & "'"
    
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sqlC)
        
        If rs.EOF And rs.BOF Then
            ExisteRHConceptoCta = False
        Else
            ExisteRHConceptoCta = True
        End If
        
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function

'Public Function GetErroresPlanillas(psPlanillaCod As String, psPeriodo As String, psDH As String) As String
'    Dim oCon As DConecta
'    Dim sql As String
'    Set oCon = New DConecta
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    Dim lsCadena As String
'    Dim sqlAdd As String
'    lsCadena = ""
'    oCon.AbreConexion
'
'    If psPlanillaCod = gsRHPlanillaCTS Then
'        sqlAdd = Producto.gCapCTS
'    Else
'        sqlAdd = Producto.gCapAhorros
'
'        sql = " Select cCtaCod, cPersNombre Nombre from rhcuentas RHC " _
'            & " Inner Join Persona PE On PE.cPersCod = RHC.cPersCod " _
'            & " where PE.cperscod in (Select cPersCod from rhplanilladet where crrhhperiodo like '" & psPeriodo & "' and cplanillacod = ('" & psPlanillaCod & "'))" _
'            & " and substring(cctacod,3,3) = '232' and cctacod not in (Select ccodcta from " & oCon.GetCadenaConexionEnlazado("01", "01", False) & "ITFCtaExonerada where cFlag is Null)"
'        Set rs = oCon.CargaRecordSet(sql)
'
'        If Not (rs.EOF And rs.BOF) Then
'            lsCadena = lsCadena & "Cuentas No Exoneradas de ITF:" & oImpresora.gPrnSaltoLinea
'            While Not rs.EOF
'                lsCadena = lsCadena & rs!cCtaCod & "-" & rs!Nombre & oImpresora.gPrnSaltoLinea
'                rs.MoveNext
'            Wend
'        End If
'
'    End If
'
'    sql = " Select Distinct RH.cRHCod, dbo.PstaNombre(PE.cPersNombre,1) Nombre From RHPlanillaDetCon PD" _
'        & " Inner Join Persona PE On PE.cPersCod = PD.cPersCod" _
'        & " Inner Join RRHH RH On RH.cPersCod = PD.cPersCod" _
'        & " Where PD.cPlanillaCod = ('" & psPlanillaCod & "') And PD.cRRHHPeriodo Like '" & psPeriodo & "'" _
'        & " And PE.cPerscod Not In" _
'        & "    (Select Distinct cPerscod From RHCuentas Where cPersCod in" _
'        & "    (Select Distinct cPersCod From RHPlanilladetcon  where cPlanillaCod in ('" & psPlanillaCod & "') and cRRHHPeriodo Like '" & psPeriodo & "')" _
'        & "      And cCtaCod like '__" & sqlAdd & "%') Order by RH.cRHCod"
'    Set rs = oCon.CargaRecordSet(sql)
'
'
'    If Not (rs.EOF And rs.BOF) Then
'        lsCadena = lsCadena & "Sin Cuentas:" & oImpresora.gPrnSaltoLinea
'        While Not rs.EOF
'            lsCadena = lsCadena & rs!cRHCod & "-" & rs!Nombre & oImpresora.gPrnSaltoLinea
'            rs.MoveNext
'        Wend
'    End If
'
'    sql = " Select Distinct RH.cRHCod, dbo.PstaNombre(PE.cPersNombre,1) Nombre From " & oCon.GetCadenaConexionEnlazado("01", "01", False) & "Ahorroc AHO" _
'        & " Inner Join RHCuentas RHCTA On AHO.cCodCta = RHCTA.cCtaCod" _
'        & " Inner Join Persona PE On PE.cPersCod = RHCTA.cPersCod" _
'        & " Inner Join RRHH RH On RH.cPersCod = RHCTA.cPersCod" _
'        & " Where cCodcta In" _
'        & " (Select cCtaCod From RHCuentas where cPersCod In" _
'        & " (Select cPersCod From RHPlanillaDetCon Where cPlanillaCod in ('" & psPlanillaCod & "') and cRRHHPeriodo like '" & psPeriodo & "')" _
'        & " And cCtacod like '__" & sqlAdd & "%') and cEstCtaAC In ('C','U')"
'    Set rs = oCon.CargaRecordSet(sql)
'
'    If Not (rs.EOF And rs.BOF) Then
'        lsCadena = lsCadena & "Con Cuentas Anuladas o Canceladas:" & oImpresora.gPrnSaltoLinea
'        While Not rs.EOF
'            lsCadena = lsCadena & rs!cRHCod & "-" & rs!Nombre & oImpresora.gPrnSaltoLinea
'            rs.MoveNext
'        Wend
'    End If
'
'    sql = " Select RHC.cCtaCod, cNomPers,PC.cCodPers,  RHC.cPersCod, cPersNombre from RHCuentas RHC" _
'        & " Inner Join " & oCon.GetCadenaConexionEnlazado("01", "01", False) & "PersCuenta PC On RHC.cCtaCod = PC.cCodCta And Substring(RHC.cPersCod,4,10) <> PC.cCodPers" _
'        & " Inner Join Persona PE On RHC.cPersCod =  PE.cPersCod" _
'        & " Inner Join " & oCon.GetCadenaConexionEnlazado("01", "01", True) & "dbPersona.dbo.Persona PED On PC.cCodPers =  PED.cCodPers" _
'        & " Where RHC.cCtaCod  Not In ('012321015687','012321086894','012321095427','012321102091')" _
'        & " Order by RHC.cCtaCod"
'    Set rs = oCon.CargaRecordSet(sql)
'
'    If Not (rs.EOF And rs.BOF) Then
'        lsCadena = lsCadena & "Con Cuentas Asignadas a Otras Personas o a Personas Duplicadas :" & oImpresora.gPrnSaltoLinea
'        While Not rs.EOF
'            lsCadena = lsCadena & rs!cCtaCod & " - " & rs!cNomPers & " - " & rs!cPersNombre & oImpresora.gPrnSaltoLinea
'            rs.MoveNext
'        Wend
'    End If
'
'    sql = " Select RH.cRHCod, PE.cPersNombre, B.cRHConceptoCod, B.cRHConceptoDescripcion from RHPlanillaConceptoTabla A " _
'        & " Inner Join (Select cRHConceptoCod, cRHConceptoDescripcion from rhConceptoTabla Where bRHObligatorio = 1) As B On A.cRHConceptoCod = B.cRHConceptoCod" _
'        & " Inner Join RHPlanillaAplicacion C On A.cRHPlanillaCod = C.cRHPlanillaCod" _
'        & " Inner Join RRHH RH On RH.nRHEstado = C.nRHEstado And RH.cRHCod Like '" & Left(psPlanillaCod, 1) & "%'" _
'        & " Inner Join Persona PE On RH.cPersCod = PE.cPersCod" _
'        & " Left  Join RHConcepto RHC on A.cRHPlanillaCod = C.cRHPlanillaCod And RHC.cPersCod = RH.cPersCod And RHC.cRHConceptoCod = B.cRHConceptoCod" _
'        & " Where A.cRHPlanillaCod in ('" & psPlanillaCod & "') And RHC.cRHConceptoCod Is Null Order By RH.cRHCod"
'    Set rs = oCon.CargaRecordSet(sql)
'
'    If Not (rs.EOF And rs.BOF) Then
'        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'        lsCadena = lsCadena & "Personas con conceptos obligatorios no Asignados :" & oImpresora.gPrnSaltoLinea
'        While Not rs.EOF
'            lsCadena = lsCadena & rs!cRHConceptoCod & " - " & rs!cRHConceptoDescripcion & " - " & rs!cRHCod & " - " & rs!cPersNombre & oImpresora.gPrnSaltoLinea
'            rs.MoveNext
'        Wend
'    End If
'
'    GetErroresPlanillas = lsCadena
'    oCon.CierraConexion
'    Set oCon = Nothing
'End Function

Public Function GetErroresPlanillas(psPlanillaCod As String, psPeriodo As String, psDH As String, pgbBitCentral As Boolean) As String
    Dim oCon As DConecta
    Dim sql As String
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim lsCadena As String
    Dim sqlAdd As String
    lsCadena = ""
    oCon.AbreConexion
    
    
    If psPlanillaCod = gsRHPlanillaCTS Then
        sqlAdd = Producto.gCapCTS
    Else
        sqlAdd = Producto.gCapAhorros
    End If
    
    'MAVM 20120228 ***
    'sql = " Select Distinct RH.cRHCod, dbo.PstaNombre(PE.cPersNombre,1) Nombre From RHPlanillaDetCon PD" _
    '    & " Inner Join Persona PE On PE.cPersCod = PD.cPersCod" _
    '    & " Inner Join RRHH RH On RH.cPersCod = PD.cPersCod" _
    '    & " Where PD.cPlanillaCod = ('" & psPlanillaCod & "') And PD.cRRHHPeriodo Like '" & psPeriodo & "'" _
    '    & " And PE.cPerscod Not In" _
    '    & "    (Select Distinct cPerscod From RHCuentas Where cPersCod in" _
    '    & "    (Select Distinct cPersCod From RHPlanilladetcon  where cPlanillaCod in ('" & psPlanillaCod & "') and cRRHHPeriodo Like '" & psPeriodo & "')" _
    '    & "      And cCtaCod like '__" & IIf(pgbBitCentral, "___", "") & sqlAdd & "%') Order by RH.cRHCod"
    'Set rs = oCon.CargaRecordSet(sql)
    
    sql = " Select Distinct RH.cRHCod, dbo.PstaNombre(PE.cPersNombre,1) Nombre From RHPlanillaDetCon PD" _
        & " Inner Join Persona PE On PE.cPersCod = PD.cPersCod" _
        & " Inner Join RRHH RH On RH.cPersCod = PD.cPersCod" _
        & " Where PD.cPlanillaCod = ('" & psPlanillaCod & "') And PD.cRRHHPeriodo Like '" & psPeriodo & "'" _
        & " And PE.cPerscod Not In" _
        & "    (Select Distinct cPerscod From ProductoPersona Where cPersCod in" _
        & "    (Select Distinct cPersCod From RHPlanilladetcon where cPlanillaCod in ('" & psPlanillaCod & "') and cRRHHPeriodo Like '" & psPeriodo & "')" _
        & "    And cCtacod in (Select Distinct cCtaCod From RHPlanilladetcon where cPlanillaCod in ('" & psPlanillaCod & "') and cRRHHPeriodo Like '" & psPeriodo & "')" _
        & "      And cCtaCod like '__" & IIf(pgbBitCentral, "___", "") & sqlAdd & "%') Order by RH.cRHCod"
    Set rs = oCon.CargaRecordSet(sql)
    '***
        
    If Not (rs.EOF And rs.BOF) Then
        lsCadena = lsCadena & "Sin Cuentas:" & oImpresora.gPrnSaltoLinea
        While Not rs.EOF
            lsCadena = lsCadena & rs!cRHCod & "-" & rs!Nombre & oImpresora.gPrnSaltoLinea
            rs.MoveNext
        Wend
    End If
    
    sql = " Select Distinct RH.cRHCod, dbo.PstaNombre(PE.cPersNombre,1) Nombre From captacAhorros AHO" _
        & " Inner Join RHCuentas RHCTA On AHO.cCtaCod = RHCTA.cCtaCod" _
        & " Inner Join Persona PE On PE.cPersCod = RHCTA.cPersCod" _
        & " Inner Join RRHH RH On RH.cPersCod = RHCTA.cPersCod" _
        & " INNER JOIN PRODUCTO p ON AHO.cCtacod=p.cCtaCod " _
        & " Where AHO.cCtaCod In" _
        & " (Select cCtaCod From RHCuentas where cPersCod In" _
        & " (Select cPersCod From RHPlanillaDetCon Where cPlanillaCod in ('" & psPlanillaCod & "') and cRRHHPeriodo like '" & psPeriodo & "')" _
        & " And cCtacod like '__" & IIf(pgbBitCentral, "___", "") & sqlAdd & "%') and nPrdEstado In (1300,1400)"
    Set rs = oCon.CargaRecordSet(sql)
        
    If Not (rs.EOF And rs.BOF) Then
        lsCadena = lsCadena & "Con Cuentas Anuladas o Canceladas:" & oImpresora.gPrnSaltoLinea
        While Not rs.EOF
            lsCadena = lsCadena & rs!cRHCod & "-" & rs!Nombre & oImpresora.gPrnSaltoLinea
            rs.MoveNext
        Wend
    End If
    'Mody by gitu 08-05-2009 Se agrego la condicion para que la consulta sea
    'a los titulares de la cuenta
    sql = " Select RHC.cCtaCod, pe.cPersNombre,PC.cPersCod,  RHC.cPersCod from RHCuentas RHC" _
        & " Inner Join ProductoPersona PC On RHC.cCtaCod = PC.cCtaCod And RHC.cPersCod <> PC.cPersCod" _
        & " Inner Join Persona PE On RHC.cPersCod =  PE.cPersCod" _
        & " Inner Join Persona PED On PC.cPersCod =  PED.cPersCod" _
        & " Where RHC.cCtaCod  Not In ('012321015687','012321086894','012321095427','012321102091')" _
        & " And PC.nPrdPersRelac = 10 " _
        & " Order by RHC.cCtaCod"
    Set rs = oCon.CargaRecordSet(sql)
    'end gitu
    
    If Not (rs.EOF And rs.BOF) Then
        lsCadena = lsCadena & "Con Cuentas Asignadas a Otras Personas o a Personas Duplicadas :" & oImpresora.gPrnSaltoLinea
        While Not rs.EOF
            lsCadena = lsCadena & rs!cCtaCod & " - " & rs!cPersNombre & " - " & rs!cPersNombre & oImpresora.gPrnSaltoLinea
            rs.MoveNext
        Wend
    End If
    
    GetErroresPlanillas = lsCadena
    oCon.CierraConexion
    Set oCon = Nothing
End Function


Public Function GetResultadoAbonoPlanillas(psPlanillaCod As String, psPeriodo As String, psUsuario As String) As String
    Dim oCon As DConecta
    Dim sql As String
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim lsCadena As String
    Dim sqlAdd As String
    Dim lsRuta01 As String
    Dim lsRuta07 As String
    Dim lsOpeAbono As String
    Dim oProc As NRHProcesosCierre
    Set oProc = New NRHProcesosCierre
    
    Dim oConR As DConecta
    Set oConR = New DConecta
    Dim oAge As DActualizaDatosArea
    Set oAge = New DActualizaDatosArea
    Dim rsAge As ADODB.Recordset
    Set rsAge = New ADODB.Recordset
    
    Dim lsAge As String * 30
    Dim lsAgeMonto As String * 12
    Dim lnMontoAge As Currency
    
    lsCadena = ""
    oCon.AbreConexion
    
    lsOpeAbono = oProc.GetOpeAbonoPlanilla(psPlanillaCod)
    lsRuta01 = oCon.GetCadenaConexionEnlazado("01", "01")
    
    lsCadena = ""
    
    If psPlanillaCod = gsRHPlanillaCTS Then
        sql = " Select Top 1 DISTINCT cRHCod from RHPlanillaDetCon a" _
            & " Inner join rhcuentas b on a.cperscod = b.cperscod" _
            & " Inner join rrhh c on c.cperscod = b.cperscod" _
            & " where crrhhperiodo like '" & Left(psPeriodo, 6) & "%' and cplanillacod = '" & psPlanillaCod & "' and cCtaCod like '__" & Producto.gCapCTS & "%'" _
            & " and cctacod Not In" _
            & " (Select cCodCta From " & lsRuta01 & "Trandiaria where ccodusu = '" & psUsuario & "' and ccodope = '" & lsOpeAbono & "')" _
            & " ORDER BY cRHCod"
        Set rs = oCon.CargaRecordSet(sql)
            
        lsCadena = lsCadena & "ABONO A CUENTAS" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        If Not (rs.EOF And rs.BOF) Then
            lsCadena = lsCadena & "Falta abonar algunos empleados, el primer empleado q falta es : " & rs.Fields(0) & oImpresora.gPrnSaltoLinea
        Else
            lsCadena = lsCadena & "El Abono a sido generado el abono para todos los empleados." & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        End If
    Else
        sql = " Select Top 1 cRHCod from RHPlanillaDetCon a" _
            & " Inner join rhcuentas b on a.cperscod = b.cperscod" _
            & " Inner join rrhh c on c.cperscod = b.cperscod" _
            & " where crrhhperiodo like '" & Left(psPeriodo, 6) & "%' and cplanillacod = '" & psPlanillaCod & "' and cCtaCod like '__" & Producto.gCapAhorros & "%'" _
            & " and cctacod Not In" _
            & " (Select cCodCta From " & lsRuta01 & "Trandiaria where ccodusu = '" & psUsuario & "' and ccodope = '" & lsOpeAbono & "') And cRHConceptoCod = '130' And nMonto <> 0" _
            & " ORDER BY cRHCod"
        Set rs = oCon.CargaRecordSet(sql)
            
        lsCadena = lsCadena & "ABONO A CUENTAS" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        If Not (rs.EOF And rs.BOF) Then
            lsCadena = lsCadena & "Falta abonar algunos empleados, el primer empleado q falta es : " & rs.Fields(0) & oImpresora.gPrnSaltoLinea
        Else
            lsCadena = lsCadena & "El Abono a sido generado el abono para todos los empleados." & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        End If
    End If
        
    sql = "Select IsNull(sum(nmontran),0) Monto, Substring(cCodCta,6,1) Moneda from " & lsRuta01 & "trandiaria where ccodusu = '" & psUsuario & "' and ccodope = '" & lsOpeAbono & "' Group by  Substring(cCodCta,6,1)"
    Set rs = oCon.CargaRecordSet(sql)
    
    lsCadena = lsCadena & "MONTO ABONO A CUENTAS" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    If rs.EOF And rs.BOF Then
        lsCadena = lsCadena & "No Existen Abonos." & oImpresora.gPrnSaltoLinea
    Else
        While Not rs.EOF
            If rs!Moneda = Moneda.gMonedaNacional Then
                lsCadena = lsCadena & "     Soles   : " & Format(rs!Monto, gsFormatoNumeroView) & oImpresora.gPrnSaltoLinea
            Else
                lsCadena = lsCadena & "     Dolares : " & Format(rs!Monto, gsFormatoNumeroView) & oImpresora.gPrnSaltoLinea
            End If
            rs.MoveNext
        Wend
    End If
    
    sql = "Select sum(nmontran) Monto, Substring(cCodCta,6,1) Moneda from " & lsRuta01 & "trandiaria where ccodusu = '" & psUsuario & "' and ccodope = '201801' Group By Substring(cCodCta,6,1)"
    Set rs = oCon.CargaRecordSet(sql)
    
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & "CARGOS EXTRA PLANILLA A CUENTAS" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    If rs.EOF And rs.BOF Then
        lsCadena = lsCadena & "No Existen Abonos." & oImpresora.gPrnSaltoLinea
    Else
        While Not rs.EOF
            If rs!Moneda = Moneda.gMonedaNacional Then
                lsCadena = lsCadena & "     Soles   : " & Format(rs!Monto, gsFormatoNumeroView) & oImpresora.gPrnSaltoLinea
            Else
                lsCadena = lsCadena & "     Dolares : " & Format(rs!Monto, gsFormatoNumeroView) & oImpresora.gPrnSaltoLinea
            End If
            rs.MoveNext
        Wend
    End If
    
    Set rsAge = oAge.GetAgencias1(True)
    
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & "ABONOS EXTRA PLANILLA A CUENTAS" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    sql = "Select sum(nmontran) Monto from trandiaria where ccodusu = '" & psUsuario & "' and ccodope = '201701'"
    While Not rsAge.EOF
        oConR.AbreConexion 'Remota rsAge.Fields(0)
        Set rs = oConR.CargaRecordSet(sql)
        If Not IsNull(rs!Monto) Then
            lnMontoAge = lnMontoAge + rs!Monto
            lsAge = rsAge.Fields(1)
            RSet lsAgeMonto = Format(rs!Monto, gsFormatoNumeroView)
            lsCadena = lsCadena & "Agencia : " & lsAge & "  :  " & lsAgeMonto & oImpresora.gPrnSaltoLinea
        End If
        rsAge.MoveNext
        oConR.CierraConexion
    Wend
    
    lsAge = ""
    RSet lsAgeMonto = "------------------------------"
    lsCadena = lsCadena & "          " & lsAge & "     " & lsAgeMonto & oImpresora.gPrnSaltoLinea
    
    lsAge = "TOTAL"
    RSet lsAgeMonto = Format(lnMontoAge, gsFormatoNumeroView)
    lsCadena = lsCadena & "          " & lsAge & "     " & lsAgeMonto & oImpresora.gPrnSaltoLinea
    
    GetResultadoAbonoPlanillas = lsCadena
    oCon.CierraConexion
    Set oCon = Nothing
End Function

#If DebugMode Then
    '##ModelId=3AB902FF0330
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If


'Actualiza una nueva relacion entre concepto y planilla
'##ModelId=3A9B14D6011D
Public Function CierraPlanilla(psRHPlanillaPeriodo As String, psRHPlanillaCod As String) As Boolean
    On Error GoTo ModificaConceptoPlanillaErr
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlDC As String
    
    If oCon.AbreConexion Then
        oCon.BeginTrans
        sqlDC = "Update RHPlanilla Set cRHPlanillaEstado = 2 where cRRHHPeriodo = '" & psRHPlanillaPeriodo & "' And cPlanillaCod = '" & psRHPlanillaCod & "'"
        oCon.Ejecutar sqlDC
        oCon.CommitTrans
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
ModificaConceptoPlanillaErr:
    oCon.RollBackTrans
    oCon.CierraConexion
    Call RaiseError(MyUnhandledError, "NActualizaDatosConceptoPlanilla:ModificaConceptoPlanilla Method")
End Function


Public Function GetRHListaMontoCargo(psPeriodo As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    Dim sSql As String
    If oConec.AbreConexion Then
        sSql = "SP_RHlistaMontosCargo  '" & psPeriodo & "' "
        Set rs = oConec.CargaRecordSet(sSql)
        Set GetRHListaMontoCargo = rs
        oConec.CierraConexion
        Set oConec = Nothing
     End If
End Function


Public Function GetRHPresupuesto(psGrupo As String, psPeriodo As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    Dim sSql As String
    If oConec.AbreConexion Then
    sSql = " execute SP_RHPRESUPUESTO '" & psGrupo & "','" & psPeriodo & "' "
        oConec.Ejecutar (sSql)
        sSql = "SP_RHPRESUPUESTO_GET "
        Set rs = oConec.CargaRecordSet(sSql)
        Set GetRHPresupuesto = rs
        oConec.CierraConexion
        Set oConec = Nothing
     End If
End Function

Public Function GetRHListadoBoletas(psPeriodo As String, psAgencia As String, pnImprime As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim numDias As Integer
    Dim dtfecha As String
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    Dim sSql As String
     If oConec.AbreConexion Then
      sSql = " SELECT  DISTINCT RH.CPERSCOD ,cPersNombre,cAgeDescripcion, " & pnImprime & " as Imprime " _
            & "  FROM RHPLANILLADETCON RH,PERSONA P,RHCARGOS RH1,AGENCIAS AG " _
            & " WHERE  AG.cAgeCod =RH1.cRHAgenciaCodOficial AND  RH.CPERSCOD =RH1.CPERSCOD AND " _
            & " P.CPERSCOD = RH.CPERSCOD AND dRHCargoFecha IN " _
            & " (SELECT max(dRHCargoFecha) FROM  RHCARGOS RH2 WHERE RH1.CPERSCOD = RH2.CPERSCOD  and " _
            & " convert(char(4),year(dRHCargoFecha)) + RIGHT('00'+convert(varchar(2),month(dRHCargoFecha)),2) <= " & psPeriodo & " ) and " _
            & " SUBSTRING (cRRHHPeriodo,1,6) = '" & psPeriodo & "'   and cPlanillaCod like 'E%'  "
        If psAgencia <> "" Then
            sSql = sSql + " and cRHAgenciaCodOficial  ='" & psAgencia & "'"
        End If
            sSql = sSql + " ORDER BY cPersNombre ASC "
        Set rs = oConec.CargaRecordSet(sSql)
        Set GetRHListadoBoletas = rs
        oConec.CierraConexion
        Set oConec = Nothing
     End If
End Function


Public Function ActualizaRHPDT(psRHPeriodo As String, psRHMes As String, psRuta As String) As Integer
    On Error GoTo ActualizaRHPDTErr
    Dim sqlP As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sqlP = "Update rhpdt Set nActivo = 0  "
        oCon.Ejecutar sqlP
    End If
    If oCon.AbreConexion Then
        sqlP = "Update rhpdt Set nActivo = 1 ," _
        & " CDirectorio  = '" & psRuta & "'" _
        & " where cAno = '" & psRHPeriodo & "' and cMes = '" & psRHMes & "'   "
        oCon.Ejecutar sqlP
    End If
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
ActualizaRHPDTErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:ActualizaRHPDT Method")
End Function

Public Function GetRHListaFormatoPDT(psPeriodo As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Dim numDias As Integer
    Dim dtfecha As String
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    Dim sSql As String
     dtfecha = "15" + "/" + Right(psPeriodo, 2) + "/" + Left(psPeriodo, 4)
     
     numDias = Left(DateSerial(Year(dtfecha), Month(dtfecha) + 1, 0), 2)
     
     If oConec.AbreConexion Then
        sSql = "spRHListaPDT '" & psPeriodo & "'," & numDias & ""
        Set rs = oConec.CargaRecordSet(sSql)
        Set GetRHListaFormatoPDT = rs
        oConec.CierraConexion
        Set oConec = Nothing
     End If
End Function


Public Function GetPrePlanilla(psPlanillaCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    sqlC = " spRHPrePlanilla  '" & psPlanillaCod & "'"
    If oCon.AbreConexion Then
        Set GetPrePlanilla = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function

Public Function GetRHNombreConcepto(psCodConcepto As String) As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
    Dim sSql As String
    If oConec.AbreConexion Then
        sSql = "Select  cRHConceptoAbreviatura + replicate (' ',3) + cRHConceptoCod As NombreConcepto  from rhconceptotabla " _
               & " where  cRHConceptoCod = '" & psCodConcepto & "' "
        Set rs = oConec.CargaRecordSet(sSql)
        oConec.CierraConexion
        Set oConec = Nothing
    End If
        If rs.EOF = True Then
        GetRHNombreConcepto = "No Encontrado"
        Else
        GetRHNombreConcepto = rs!NombreConcepto
        End If
End Function

Public Function ActualizaRHConceptoPla(psRHPlanillaCod As String, psRHCodPersona As String, psRHCodConcepto As String, pnRHCodMonto As Double, psActualiza As String) As Integer
    On Error GoTo ActualizaRHConceptoPlaErr
    Dim sqlP As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sqlP = " Delete  from rhconcepto where cRHPlanillaCod ='" & psRHPlanillaCod & "'  and  " _
           & " cPersCod  ='" & psRHCodPersona & "' and  cRHConceptoCod  = '" & psRHCodConcepto & "'"
        oCon.Ejecutar sqlP
    End If
    
    If oCon.AbreConexion Then
        sqlP = " Insert RHCONCEPTO (cRHPlanillaCod ,cPersCod ,cRHConceptoCod,nRHConceptoValor,cUltimaActualizacion) " _
         & " Values('" & psRHPlanillaCod & "','" & psRHCodPersona & "','" & psRHCodConcepto & "'," & pnRHCodMonto & "," & psActualiza & " )"
        oCon.Ejecutar sqlP
    End If
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
ActualizaRHConceptoPlaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:ActualizaRHConceptoPla Method")
End Function

'MAVM 20110715 ***
Public Function ActualizarDiasVacaciones(ByVal psPersCod As String, ByVal pnValor As Currency, ByVal pnTipo As Integer, Optional psFecha As String) As Boolean
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlD As String
    Dim sqlE As String
    
    On Error GoTo ActualizarDiasVacaciones
    
    If oCon.AbreConexion Then
        oCon.BeginTrans
            If pnTipo = 1 Then
                If pnValor <> 0 Then
                    sqlD = "Update RHEmpleado Set nRHEmplVacacionesPend = nRHEmplVacacionesPend + " & pnValor & " Where cPersCod = '" & psPersCod & "'"
                    sqlE = "Delete MovVacaciones Where cPersCod = '" & psPersCod & "' And cCodVacaciones = 2 And cPeriodo = '" & Mid(psFecha, 3, 6) & "'"
                End If
            Else
                If pnValor <> 0 Then
                    sqlD = "Update RHEmpleado Set nRHEmplVacacionesPend = nRHEmplVacacionesPend - " & pnValor & " Where cPersCod = '" & psPersCod & "'"
                    sqlE = "Insert MovVacaciones (cPersCod, cCodVacaciones, dFecha, cUser, nDias, cFlag, cPeriodo) Values ('" & psPersCod & "', '2', '" & Mid(psFecha, 5, 2) & "/" & Mid(psFecha, 7, 2) & "/" & Mid(psFecha, 1, 4) & "', '" & Mid(psFecha, 9, 4) & "', " & pnValor & ", '1', '" & Mid(psFecha, 3, 6) & "'" & ")"
                End If
            End If
            If sqlD = "" Then Exit Function
            oCon.Ejecutar sqlD
            oCon.Ejecutar sqlE
        oCon.CommitTrans
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
ActualizarDiasVacaciones:
    oCon.RollBackTrans
    oCon.CierraConexion
    Call RaiseError(MyUnhandledError, "DActualizaDatosConPlanilla:ActualizarDiasVacaciones Metodo")
End Function
'***

'*** PEAC 20131118
Public Function ModificaNumCtaComentario(psCodEmpModiComent As String, psTipoPlanilla As String, psPeriodo As String, psNumCta As String) As Boolean
    On Error GoTo ModificaNumCtaComentarioErr
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sqlD As String
    Dim sqlDD As String
    Dim sqlDC As String
    Dim lnI As Integer
    Dim lnMontoI As Currency
    Dim lnMontoD As Currency
    
    If oCon.AbreConexion Then
        oCon.BeginTrans

        sqlD = "exec stp_Upd_RHPlanillaDetCon '" & psCodEmpModiComent & "','" & psTipoPlanilla & "','" & psPeriodo & "','" & psNumCta & "' "
        oCon.Ejecutar sqlD
            
        oCon.CommitTrans
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
ModificaNumCtaComentarioErr:
    oCon.RollBackTrans
    oCon.CierraConexion
    Call RaiseError(MyUnhandledError, "NActualizaDatosConceptoPlanilla:ModificaNumCtaComentario Method")
End Function
