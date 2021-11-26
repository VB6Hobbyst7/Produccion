VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DGrupoEco"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetGE() As ADODB.Recordset
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sql = " Select cGECod,cGENom,cConsDescripcion + Space(50) + STR(nConsValor) Valor From PersGrupoEcon GU " _
        & " Inner Join Constante CO On CO.nConsCod = 4027 And GU.nGETipo = CO.nConsValor Order By cGECod "

    oCon.AbreConexion
    
    Set GetGE = oCon.CargaRecordSet(sql)
    
    oCon.CierraConexion
End Function

Public Function GetGEExiste(psGECod As String) As Boolean
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = " Select cGECod From PersGrupoEcon" _
        & " Where cGECod = '" & psGECod & "'"

    oCon.AbreConexion
    
    Set rs = oCon.CargaRecordSet(sql)
    
    If Not (rs.EOF And rs.BOF) Then
        GetGEExiste = False
    Else
        GetGEExiste = True
    End If
    
    oCon.CierraConexion
End Function

Public Function GetGERelExiste(psGECod As String, psPersCod As String) As Boolean
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = " Select cGECod From PersGERelacion " _
        & " Where cGECod = '" & psGECod & "' And cPersCodRel = '" & psPersCod & "'"

    oCon.AbreConexion
    
    Set rs = oCon.CargaRecordSet(sql)
    
    If Not (rs.EOF And rs.BOF) Then
        GetGERelExiste = False
    Else
        GetGERelExiste = True
    End If
    
    oCon.CierraConexion
End Function

Public Function GetGERelVinExiste(psPersCodEmp As String, psPersCodPers As String) As Boolean
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = " Select cPersCodVinc From PersGEPersVinc " _
        & " Where cPersCodVinc = '" & psPersCodPers & "' And cPersCodRel = '" & psPersCodEmp & "'"

    oCon.AbreConexion
    
    Set rs = oCon.CargaRecordSet(sql)
    
    If Not (rs.EOF And rs.BOF) Then
        GetGERelVinExiste = False
    Else
        GetGERelVinExiste = True
    End If
    
    oCon.CierraConexion
End Function

Public Function ActulizaGE(prRS As ADODB.Recordset) As Boolean
    Dim sql As String
    Dim lsCad As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    If prRS Is Nothing Then
        lsCad = ""
        sql = "Delete PersGrupoEcon Where cGECod Not In ('" & lsCad & "')"
        oCon.Ejecutar sql
        Exit Function
    Else
        prRS.MoveFirst
        
        While Not prRS.EOF
            If lsCad = "" Then
                lsCad = prRS.Fields(0)
            Else
                lsCad = lsCad & "','" & prRS.Fields(0)
            End If
            prRS.MoveNext
        Wend
    End If
    
    sql = "Delete PersGrupoEcon Where cGECod Not In ('" & lsCad & "')"
    oCon.Ejecutar sql
        
    prRS.MoveFirst
    
    While Not prRS.EOF
        If GetGEExiste(prRS.Fields(0)) Then
            sql = " Insert PersGrupoEcon (cGECod, cGENom, nGETipo)" _
                & " Values('" & prRS.Fields(0) & "','" & prRS.Fields(1) & "','" & Right(prRS.Fields(2), 2) & "')"
        Else
            sql = " Update PersGrupoEcon" _
                & " Set cGENom = '" & prRS.Fields(1) & "', nGETipo = '" & Right(prRS.Fields(2), 2) & "'" _
                & " Where cGECod = '" & prRS.Fields(0) & "'"
        End If
        oCon.Ejecutar sql
        prRS.MoveNext
    Wend
    
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
End Sub


Public Function GetGEEmp(psGE As String) As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sql As String
    
    oCon.AbreConexion
    
    sql = " Select  GU.cPersCodRel, PE.cPersNombre, CO.cConsDescripcion + Space(50) + Str(CO.nConsvalor) Valor " _
        & " From PersGERelacion GU" _
        & " Inner Join Persona PE On GU.cPersCodRel = PE.cPersCod" _
        & " Inner Join Constante CO On CO.nConsCod = 4028 And GU.nPrdPersRelac = CO.nConsValor" _
        & " Where cGEcod = '" & psGE & "'"
    
    Set GetGEEmp = oCon.CargaRecordSet(sql)
    
    Set oCon = Nothing
End Function

Public Function GetGEPers(psGE As String) As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sql As String
    
    oCon.AbreConexion
    
    sql = " Select GE.cPersCodVinc, PE.cPersNombre, CO.cConsDescripcion + Space(50) + Str(CO.nConsvalor) Valor," _
        & " nParticip, cPersCodRel from PersGEPersVinc GE" _
        & " Inner Join Persona PE ON GE.cPersCodVinc = PE.cPersCod" _
        & " Inner Join Constante CO On CO.nConsCod = 4029 And GE.nCargo = CO.nConsValor" _
        & " Where cPersCodRel In (select cPersCodRel From PersGERelacion Where cGEcod = '" & psGE & "')"
    
    Set GetGEPers = oCon.CargaRecordSet(sql)
    
    Set oCon = Nothing
End Function

Public Function ActulizaGERel(prRSEmp As ADODB.Recordset, prRSPers As ADODB.Recordset, psGECod As String) As Boolean
    Dim sql As String
    Dim lsCad As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lsCadP As String
    
    oCon.AbreConexion
    
    If prRSEmp Is Nothing Then
        sql = "Delete PersGEPersVinc Where cPersCodRel In (Select cPersCodRel From PersGERelacion Where cGECod = '" & psGECod & "')"
        oCon.Ejecutar sql
        sql = "Delete PersGERelacion Where cGECod = '" & psGECod & "'"
        oCon.Ejecutar sql
        Exit Function
    Else
        prRSEmp.MoveFirst
        
        While Not prRSEmp.EOF
            prRSPers.MoveFirst
            lsCadP = ""
            
            If lsCad = "" Then
                lsCad = prRSEmp.Fields(0)
            Else
                lsCad = lsCad & "','" & prRSEmp.Fields(0)
            End If
            
            If Not prRSPers Is Nothing Then
                While Not prRSPers.EOF
                    If lsCadP = "" Then
                        lsCadP = prRSPers.Fields(0)
                    Else
                        lsCadP = lsCadP & "','" & prRSPers.Fields(0)
                    End If
                    prRSPers.MoveNext
                Wend
            End If
            sql = "Delete PersGEPersVinc Where cPersCodRel = '" & prRSEmp.Fields(0) & "' And cPersCodVinc Not In ('" & lsCadP & "')"
            oCon.Ejecutar sql
            
            prRSEmp.MoveNext
        Wend
    End If
    
    sql = "Delete PersGERelacion Where cGECod = '" & psGECod & "' And cPersCodRel not in ('" & lsCad & "')"
    oCon.Ejecutar sql
        
    prRSEmp.MoveFirst
    While Not prRSEmp.EOF
        If GetGERelExiste(psGECod, prRSEmp.Fields(0)) Then
            sql = " Insert PersGERelacion (cGEcod,cPersCodRel,nPrdPersRelac)" _
                & " Values('" & psGECod & "','" & prRSEmp.Fields(0) & "','" & Right(prRSEmp.Fields(2), 2) & "')"
        Else
            sql = " Update PersGERelacion" _
                & " Set nPrdPersRelac = '" & Right(prRSEmp.Fields(2), 2) & "'" _
                & " Where cGECod = '" & psGECod & "' And cPersCodRel = '" & prRSEmp.Fields(0) & "'"
        End If
        oCon.Ejecutar sql
        prRSEmp.MoveNext
    Wend
    
    If Not prRSPers Is Nothing Then
        prRSPers.MoveFirst
        While Not prRSPers.EOF
            If GetGERelVinExiste(prRSPers.Fields(4), prRSPers.Fields(0)) Then
                sql = " Insert PersGEPersVinc (cPersCodRel,cPersCodVinc,nCargo,nParticip)" _
                    & " Values('" & prRSPers.Fields(4) & "','" & prRSPers.Fields(0) & "','" & Right(prRSPers.Fields(2), 3) & "'," & prRSPers.Fields(3) & ")"
            Else
                sql = " Update PersGEPersVinc" _
                    & " Set nCargo = '" & Right(prRSPers.Fields(2), 2) & "', nParticip = " & prRSPers.Fields(3) & "" _
                    & " Where cPersCodVinc = '" & prRSPers.Fields(0) & "' And cPersCodRel = '" & prRSPers.Fields(4) & "'"
            End If
            oCon.Ejecutar sql
            prRSPers.MoveNext
        Wend
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'ALPA 20090925*********************************************************
Public Function ListarDatosGrupoEconomicoxGrupoRepo19y20(nGrupoCod As Integer) As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta
Dim sSql As String
    sSql = "exec stp_sel_Reporte19y20GrupoEconomico  " & nGrupoCod
    oCon.AbreConexion
    Set ListarDatosGrupoEconomicoxGrupoRepo19y20 = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    
 Set oCon = Nothing
End Function
Public Function ListarDatosGrupoEconomicoxRURepo20(nGrupoCod As Integer, dFecha As Date, nTipoCambio As Currency) As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta
Dim sSql As String

    sSql = "exec stp_sel_Reporte20GrupoEconomico  " & nGrupoCod & ",'" & Format(dFecha, "YYYY/MM/DD") & "'," & nTipoCambio
    oCon.AbreConexion
    Set ListarDatosGrupoEconomicoxRURepo20 = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    
 Set oCon = Nothing
End Function
Public Function ObtenerDatosPropiedadDirectaEIndirecta() As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta
Dim sSql As String
    sSql = "exec stp_sel_ObtenerPropiedadDirectaEIndirecta  "
    oCon.AbreConexion
    Set ObtenerDatosPropiedadDirectaEIndirecta = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'ALPA 20120401**********************************************************************
Public Sub InsertarPersGrupoEconomicoSaldo(dFecha As Date, nTipoCambio As Currency)
Dim oCon As DConecta
Set oCon = New DConecta
Dim sSql As String

    sSql = "exec stp_ins_PersGrupoEconomicoSaldo  '" & Format(dFecha, "YYYY/MM/DD") & "'," & nTipoCambio
    oCon.AbreConexion
    oCon.CargaRecordSet (sSql)
    oCon.CierraConexion
 Set oCon = Nothing
End Sub
'ALPA 20150325**********************************************************************
Public Function ListarDatosGrupoEconomicoxRURepo20Nuevo(nGrupoCod As Integer, dFecha As Date, nTipoCambio As Currency) As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta
Dim sSql As String

    sSql = "exec stp_sel_Reporte20GrupoEconomicoNuevo  " & nGrupoCod & ",'" & Format(dFecha, "YYYY/MM/DD") & "'," & nTipoCambio
    oCon.AbreConexion
    Set ListarDatosGrupoEconomicoxRURepo20Nuevo = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    
 Set oCon = Nothing
End Function
Public Function ObtenerDatosPropiedadDirectaEIndirectaNuevo(dFecha As Date) As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta
Dim sSql As String
    sSql = "exec stp_sel_ObtenerPropiedadDirectaEIndirectaNuevo  '" & Format(dFecha, "YYYY/MM/DD") & "'"
    oCon.AbreConexion
    Set ObtenerDatosPropiedadDirectaEIndirectaNuevo = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'***********************************************************************************

Public Function ObtenerDatosVinculadosRpte21(pdFecha As Date, Optional psTipo As String = "") As ADODB.Recordset 'NAGL 20190705 Agregó psTipo Según RFC1907050001
Dim oCon As New DConecta
Dim sql As String
    sql = "Exec stp_sel_MuestraVinculados '" & Format(pdFecha, "yyyymmdd") & "', '" & psTipo & "'" 'NAGL 20190705 Agregó psTipo Según RFC1907050001
    oCon.AbreConexion
    Set ObtenerDatosVinculadosRpte21 = oCon.CargaRecordSet(sql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function 'NAGL ERS074-2017 20171209

Public Function ObtenerDatosReporte21(pnTipoCambio As Double, pdFecha As Date) As ADODB.Recordset
Dim oCon As New DConecta
Dim sql As String
    sql = "Exec stp_sel_DatosReporte21 " & pnTipoCambio & ", '" & Format(pdFecha, "yyyymmdd") & "'"
    oCon.AbreConexion
    Set ObtenerDatosReporte21 = oCon.CargaRecordSet(sql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function 'NAGL ERS074-2017 20171209

Public Sub EliminarVinculadoRpte21(psPersCod As String, pdFecha As Date)
Dim oCon As New DConecta
Dim sql As String
    sql = "Exec stp_del_EliminaVinculado21 '" & psPersCod & "','" & Format(pdFecha, "yyyymmdd") & "'"
    oCon.AbreConexion
    oCon.Ejecutar (sql)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub 'NAGL ERS074-2017 20171209

Public Sub GuardaDatosVinculadosRpte21(ByVal pdFecha As Date, ByVal pVDatos As ADODB.Recordset, ByVal psMovNro As String) 'NAGL 20190705 Agregó psMovNro
    Dim psSql As String
    Dim oConec As New DConecta
    Dim Anexo As New DAnexoRiesgos
    Dim i As Integer
    oConec.AbreConexion
    oConec.BeginTrans
    psSql = "Exec stp_del_DatosVinculado21xFecha '" & Format(pdFecha, "yyyymmdd") & "'"
    oConec.Ejecutar (psSql)
    
    If Not pVDatos Is Nothing Then
     Do While Not pVDatos.EOF
        psSql = " Exec stp_ins_DatosVinculado21 '" & Format(pdFecha, "yyyymmdd") & "', '" & pVDatos!cPerscod & "', '" & pVDatos!cVinculo & "','" & psMovNro & "'" 'NAGL 20190705 Agregó psMovNro
               oConec.Ejecutar (psSql)
               pVDatos.MoveNext
        Loop
    End If
    oConec.CommitTrans
    oConec.CierraConexion
    Set oConec = Nothing
End Sub '***********NAGL ERS074-2017 20171209

Public Function ObtieneDataRRHH(ByVal pdFechaSist As Date, ByVal pdFechaParam As Date) As Boolean
Dim psSql As String
Dim rs As New ADODB.Recordset
Dim oConec As New DConecta
Dim pbValida As Boolean
oConec.AbreConexion
    psSql = "Exec stp_sel_ValidaObtieneDataRRHH '" & Format(pdFechaSist, "yyyymmdd") & "','" & Format(pdFechaParam, "yyyymmdd") & "'"
    Set rs = oConec.CargaRecordSet(psSql)
    If rs!psValida = "1" Then
        pbValida = True
    Else
        pbValida = False
    End If
oConec.CierraConexion
ObtieneDataRRHH = pbValida
Set oConec = Nothing
End Function '****NAGL Según RFC1907050001 20190709
