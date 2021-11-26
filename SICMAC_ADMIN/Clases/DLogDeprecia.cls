VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DLogDeprecia"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsConexion As String
Dim vsCentralPer As String
Dim vsCentralCom As String
Dim vsCentralImg As String
Dim sSQL As String

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

    Dim oIni As ClasIni
    Set oIni = New ClasIni
        vsConexion = oIni.CadenaConexion
        vsCentralPer = oIni.BasePersonas
        vsCentralCom = oIni.BaseComunes
        vsCentralImg = oIni.BaseImagenes
    Set oIni = Nothing
End Sub

Public Function GetRSMovEstadistico(psFecha As String, pnTipo As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Set oConec = New DConecta
    Dim sql As String
        
    oConec.AbreConexion
    
    sql = " Select distinct m.* from bsactivofijo a" _
        & " inner join movbsaf b on a.nAnio = b.nAnio And a.nMovNro = b.nMovNro And a.cBsCod = b.cBSCod and a.cSerie = b.cSerie" _
        & " inner join Mov m on b.nMovNroRef = m.nmovnro" _
        & " where cMovNro like '" & psFecha & "%' and cOpeCod = '581201'  and nMovFlag = 0 and nMovEstado = 25 and a.nanio =  " & Left(psFecha, 4) & " and a.Ban = " & pnTipo
    
    Set GetRSMovEstadistico = oConec.CargaRecordSet(sql)
End Function

Public Sub GetRSMovEstadisticoExtorno(psFecha As String, pnTipo As Integer)
    Dim oConec As DConecta
    Set oConec = New DConecta
    Dim sql As String
        
    oConec.AbreConexion
    
    sql = " update mov set nmovflag = 1 where nmovnro in (Select distinct m.nMovNro from bsactivofijo a" _
        & " inner join movbsaf b on a.nAnio = b.nAnio And a.nMovNro = b.nMovNro And a.cBsCod = b.cBSCod and a.cSerie = b.cSerie" _
        & " inner join Mov m on b.nMovNroRef = m.nmovnro" _
        & " where cMOvNro like '" & psFecha & "%' and cOpeCod = '581201'  and nMovFlag = 0 and nMovEstado = 25 and a.nanio = " & Left(psFecha, 4) & " and a.Ban = " & pnTipo & ")"
    
    oConec.Ejecutar sql
End Sub


'Devuelva un RecordSet con las Contrataciones
Public Function GetAFDeprecia(ByVal pdDeprecia As Date, pnTipo As Long, ByVal pnAnio As Integer, Optional pbIncluyeBaja As Boolean = False, Optional pbChek As Boolean = False, Optional pnTipoDepreciacion As Integer = 1, Optional pbYaDepreciado As Integer = 0) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset
'And (cCtaContCodOtroD Is Null Or BS.cBScod like cCtaContCodOtroD + '%'  )
 
    If oConec.AbreConexion() Then

'*** PEAC 20110617 - SE MANDO EL SCRIPT A UN STP PARA MEJOR MANTENIMIENTO

'        'Se modifico el calculo de la depreciacion mensual gitu 20/05/2008
'        If Not pbIncluyeBaja Then
'
'    'ALPA 20090922*********************************************************************************************************************************************************
''                sSQL = " Select " & IIf(pbChek, " 0 Chek,", "") & " BSAF.cBSCod, BSAF.cDescripcion cBSDescripcion, BSAF.cSerie, BSAF.nBSValor,  Round(dbo.GetFactorIPM (  dActivacion, '" & Format(pdDeprecia, gsFormatoFecha) & "'),10) Factor, Round(dbo.GetFactorIPM (dActivacion, '" & Format(pdDeprecia, gsFormatoFecha) & "') * nBSValor,10) V_Ajustado, BSAF.nBSPerDeprecia Vida_Util," _
''                     & " Case When BSAF.nBSPerDeprecia >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else BSAF.nBSPerDeprecia End PeriodosDeprecia," _
''                     & " Round((BSAF.nBSValor * (Case When BSAF.nBSPerDeprecia >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else BSAF.nBSPerDeprecia End)) / BSAF.nBSPerDeprecia,2) Dep_Hist," _
''                     & " Round(((Round(dbo.GetFactorIPM (dActivacion, '" & Format(pdDeprecia, gsFormatoFecha) & "') * nBSValor,10)) * (Case When BSAF.nBSPerDeprecia >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else BSAF.nBSPerDeprecia End)) / BSAF.nBSPerDeprecia,2) Dep_Ajustada," _
''                     & " BSAF.cAreCod + BSAF.cAgeCod + ' ' + ARE.cAreaDescripcion + ' ' + AGE.cAgeDescripcion Agencia," _
''                     & " BSAF.nMovNro Nro, IsNull((Select BSAFAdd.cCtaCod From BSActivoFijoCtasAdd BSAFAdd Where BSAF.cBSCod Like BSAFAdd.cBSCod + '%'),IsNull(( Select Top 1 Replace(cCtaContCodD,'AG',BSAF.cAgeCod) From CtaBS CBS Inner Join OpeCtaCta OCC On CBS.cCtaContCod = OCC.cCtaContCodH And OCC.cOpeCod = '591201' Where BS.cBScod Like cObjetoCod + '%' And CBS.cOpeCod = '591201' And (cCtaContCodOtroD Is Null Or BS.cBScod like cCtaContCodOtroD + '%' ) order by Len(cObjetoCod) Desc), BSAF.cBSCod))     cCtaContCodD   ," _
''                     & " Case When (BSAF.nBSPerDeprecia = Case When BSAF.nBSPerDeprecia >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else BSAF.nBSPerDeprecia End)  and BSAF.nBSValor <= dbo.GetLogDepAcum(" & pnAnio & ",BSAF.nMovNro, BSAF.cBSCod, BSAF.cSerie,'581201', '" & Format(pdDeprecia, gsFormatoFecha) & "') Then 0 " _
''                     & "      When (BSAF.nBSPerDeprecia = Case When BSAF.nBSPerDeprecia >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else BSAF.nBSPerDeprecia End) and BSAF.nBSValor = Round((BSAF.nBSValor * (Case When BSAF.nBSPerDeprecia >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else BSAF.nBSPerDeprecia End)) / BSAF.nBSPerDeprecia,2) Then Round(IsNull(BSAF.nBSValor/BSAF.nBSPerDeprecia,0),2) - 1 Else Round(IsNull(BSAF.nBSValor/BSAF.nBSPerDeprecia,0),2) End Dep_HistMes, " _
''                     & " Convert(Decimal(19,2),Round(((Round(dbo.GetFactorIPM (dActivacion, '" & Format(pdDeprecia, gsFormatoFecha) & "') * nBSValor,10)) * (Case When BSAF.nBSPerDeprecia >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else BSAF.nBSPerDeprecia End)) / BSAF.nBSPerDeprecia,2) - dbo.GetLogDepAcum(" & pnAnio & ",BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581202', '" & Format(pdDeprecia, gsFormatoFecha) & "')) DepAjusMes ," _
''                     & " Convert(Decimal(19,2),Round(((Round(dbo.GetFactorIPM (dActivacion, '" & Format(pdDeprecia, gsFormatoFecha) & "') * nBSValor,10)) * (Case When BSAF.nBSPerDeprecia >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else BSAF.nBSPerDeprecia End)) / BSAF.nBSPerDeprecia,2) - dbo.GetLogDepAcum(" & pnAnio & ",BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581202', '" & Format(pdDeprecia, gsFormatoFecha) & "')) - " _
''                     & " Convert(Decimal(19,2),Round((BSAF.nBSValor * (Case When BSAF.nBSPerDeprecia >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else BSAF.nBSPerDeprecia End)) / BSAF.nBSPerDeprecia,2) -  dbo.GetLogDepAcum(" & pnAnio & ",BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581201', '" & Format(pdDeprecia, gsFormatoFecha) & "')) AjusteMes ," _
''                     & " IsNull((Select cPersNombre From Persona PE1 WITH (NOLOCK) Where PE1.cPersCod = BSAF.cPersCod ),PE.cPersNombre) Nombre, Replace(Replace(M.cMovDesc ,char(13),' '),char(10),' ') Descripcion," _
''                     & " Convert(Varchar(10),dActivacion,103), BSAF.cAreCod, BSAF.cAgeCod, PE.cPersCod, Convert(Varchar(10),dCompra,103), bBaja Baja, IsNull(Convert(Varchar(10),dBaja,103),'') dBaja, " _
''                     & " dbo.GetLogDepAcum(" & pnAnio & ",BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581201', '" & Format(pdDeprecia, gsFormatoFecha) & "') Dep_Hist_Ant, " _
''                     & " dbo.GetLogDepAcum(" & pnAnio & ",BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581202', '" & Format(pdDeprecia, gsFormatoFecha) & "') Dep_Ajust_Ant, dbo.GetLogDepAcum(" & pnAnio & " - 1,BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581201', '" & Format(pdDeprecia, gsFormatoFecha) & "') Dep_Hist_Ejer_Ant, dbo.GetLogDepAcum(" & pnAnio & " - 1,BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581202', '" & Format(pdDeprecia, gsFormatoFecha) & "') Dep_Ajus_Ejer_Ant " _
''                     & " From BSActivoFijo BSAF WITH (NOLOCK)" _
''                     & " Left  Join BienesServicios BS WITH (NOLOCK) On BS.cBSCod = BSAF.cBSCod" _
''                     & " Left  Join Agencias AGE WITH (NOLOCK) On AGE.cAgeCod = BSAF.cAgeCod" _
''                     & " Inner Join Areas ARE WITH (NOLOCK) On ARE.cAreaCod = BSAF.cAreCod" _
''                     & " Inner Join Mov M WITH (NOLOCK) On M.nMovNro = BSAF.nMovNro" _
''                     & " Left  Join MovGasto MG WITH (NOLOCK) On MG.nMovNro = M.nMovNro" _
''                     & " Left  Join Persona PE WITH (NOLOCK) On Pe.cPersCod = MG.cPersCod Where nAnio <= " & Year(pdDeprecia) & " and ban =  " & pnTipo & " And bBaja = 0 And dcompra < '" & Format(DateAdd("m", 1, pdDeprecia), gsFormatoFecha) & "'" _
''                     & " ORDER BY BSAF.dCompra"
''
'                    sSQL = " Select " & IIf(pbChek, " 0 Chek,", "") & " BSAF.cBSCod Codigo, BSAF.cDescripcion Descripcion, BSAF.cSerie, BSAF.nBSValor,  Round(dbo.GetFactorIPM (  dActivacion, '" & Format(pdDeprecia, gsFormatoFecha) & "'),10) Factor, Round(dbo.GetFactorIPM (dActivacion, '" & Format(pdDeprecia, gsFormatoFecha) & "') * nBSValor,10) V_Ajustado, dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2) Vida_Util,"
'                    sSQL = sSQL & " Case When dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2) >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2) End PeriodosDeprecia,"
'                    sSQL = sSQL & " Round((BSAF.nBSValor * (Case When dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2) >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2) End)) / dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2),2) Dep_Hist,"
'                    sSQL = sSQL & " Round(((Round(dbo.GetFactorIPM (dActivacion, '" & Format(pdDeprecia, gsFormatoFecha) & "') * nBSValor,10)) * (Case When dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2) >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2) End)) / dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2),2) Dep_Ajustada,"
'                    sSQL = sSQL & " BSAF.cAreCod + BSAF.cAgeCod + ' ' + ARE.cAreaDescripcion + ' ' + AGE.cAgeDescripcion Agencia,"
'                    sSQL = sSQL & " BSAF.nMovNro Nro, IsNull((Select BSAFAdd.cCtaCod From BSActivoFijoCtasAdd BSAFAdd Where BSAF.cBSCod Like BSAFAdd.cBSCod + '%'),IsNull(( Select Top 1 Replace(cCtaContCodD,'AG',BSAF.cAgeCod) From CtaBS CBS Inner Join OpeCtaCta OCC On CBS.cCtaContCod = OCC.cCtaContCodH And OCC.cOpeCod = '591201' Where BS.cBScod Like cObjetoCod + '%' And CBS.cOpeCod = '591201' And (cCtaContCodOtroD Is Null Or BS.cBScod like cCtaContCodOtroD + '%' ) order by Len(cObjetoCod) Desc), BSAF.cBSCod))     cCtaContCodD   ,"
'                    sSQL = sSQL & " Convert(decimal(18,2),Case When (dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2) = Case When dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2) >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2) End)  and BSAF.nBSValor <= dbo.GetLogDepAcum(" & pnAnio & ",BSAF.nMovNro, BSAF.cBSCod, BSAF.cSerie,'581201', '" & Format(pdDeprecia, gsFormatoFecha) & "') Then 0 "
'                    sSQL = sSQL & "      When (dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2) "
'                    sSQL = sSQL & "         = Case When dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2) >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2) End) and BSAF.nBSValor = Round((BSAF.nBSValor * (Case When dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2) >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2) End)) / dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2),2) Then Round(dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",1)"
'                    sSQL = sSQL & "         *BSAF.nBSValor/12,2) Else Round(dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",1)*BSAF.nBSValor/12,2) End) Dep_HistMes, "
'                    sSQL = sSQL & " Convert(Decimal(19,2),Round(((Round(dbo.GetFactorIPM (dActivacion, '" & Format(pdDeprecia, gsFormatoFecha) & "') * nBSValor,10)) * (Case When dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2) >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2) End)) / dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2),2) - dbo.GetLogDepAcum(" & pnAnio & ",BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581202', '" & Format(pdDeprecia, gsFormatoFecha) & "')) DepAjusMes ,"
'                    sSQL = sSQL & " Convert(Decimal(19,2),Round(((Round(dbo.GetFactorIPM (dActivacion, '" & Format(pdDeprecia, gsFormatoFecha) & "') * nBSValor,10)) * (Case When dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2) >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2) End)) / dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2),2) - dbo.GetLogDepAcum(" & pnAnio & ",BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581202', '" & Format(pdDeprecia, gsFormatoFecha) & "')) - "
'                    sSQL = sSQL & " Convert(Decimal(19,2),Round((BSAF.nBSValor * (Case When dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2) >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2) End)) / dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2),2) -  dbo.GetLogDepAcum(" & pnAnio & ",BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581201', '" & Format(pdDeprecia, gsFormatoFecha) & "')) AjusteMes ,"
'                    sSQL = sSQL & " IsNull((Select cPersNombre From Persona PE1 WITH (NOLOCK) Where PE1.cPersCod = BSAF.cPersCod ),PE.cPersNombre) Nombre, Replace(Replace(M.cMovDesc ,char(13),' '),char(10),' ') Descripcion,"
'                    sSQL = sSQL & " Convert(Varchar(10),dActivacion,103) FActiva, BSAF.cAreCod Area, BSAF.cAgeCod Age, PE.cPersCod CodPers, Convert(Varchar(10),dCompra,103) F_Compra, bBaja Baja, IsNull(Convert(Varchar(10),dBaja,103),'') dBaja, "
'                    sSQL = sSQL & " dbo.GetLogDepAcum(" & pnAnio & ",BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581201', '" & Format(pdDeprecia, gsFormatoFecha) & "') Dep_Hist_Ant, "
'                    sSQL = sSQL & " dbo.GetLogDepAcum(" & pnAnio & ",BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581202', '" & Format(pdDeprecia, gsFormatoFecha) & "') Dep_Ajust_Ant, dbo.GetLogDepAcum(" & pnAnio & " - 1,BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581201', '" & Format(pdDeprecia, gsFormatoFecha) & "') Dep_Hist_Ejer_Ant, dbo.GetLogDepAcum(" & pnAnio & " - 1,BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581202', '" & Format(pdDeprecia, gsFormatoFecha) & "') Dep_Ajus_Ejer_Ant,isnull(cDocNro,'') cDocNro "
'                    sSQL = sSQL & " From BSActivoFijo BSAF WITH (NOLOCK)"
'                    sSQL = sSQL & " Left  Join BienesServicios BS WITH (NOLOCK) On BS.cBSCod = BSAF.cBSCod"
'                    sSQL = sSQL & " Left  Join Agencias AGE WITH (NOLOCK) On AGE.cAgeCod = BSAF.cAgeCod"
'                    sSQL = sSQL & " Inner Join Areas ARE WITH (NOLOCK) On ARE.cAreaCod = BSAF.cAreCod"
'                    sSQL = sSQL & " Inner Join Mov M WITH (NOLOCK) On M.nMovNro = BSAF.nMovNro"
'                    sSQL = sSQL & " left join (select distinct cBSCod,max(nMovNro) nMovNro from MovBs group by cBSCod) MBS on BSAF.cBSCod=MBS.cBSCod "
'                    sSQL = sSQL & " left join MovDoc MD on MBS.nMovNro=MD.nMovNro "
'                    sSQL = sSQL & " Left  Join MovGasto MG WITH (NOLOCK) On MG.nMovNro = M.nMovNro"
'                    sSQL = sSQL & " Left  Join Persona PE WITH (NOLOCK) On Pe.cPersCod = MG.cPersCod Where nAnio <= " & Year(pdDeprecia) & " and ban =  " & pnTipo & " And bBaja = 0 And dcompra < '" & Format(DateAdd("m", 1, pdDeprecia), gsFormatoFecha) & "' and dbo.fnc_ObtnerDepreciacionActivos(ban," & pnTipoDepreciacion & ",2)>0 "
'                    sSQL = sSQL & " ORDER BY BSAF.dCompra"
'''
''''                 & " Round(IsNull(BSAF.nBSValor/BSAF.nBSPerDeprecia,0),2) Dep_HistMes, "
''''                 & " Round(IsNull(BSAF.nBSValor/BSAF.nBSPerDeprecia,0) * ((Case When BSAF.nBSPerDeprecia >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else BSAF.nBSPerDeprecia End)-1),2) Dep_Hist_Ant, "
''''                 & " Convert(Decimal(19,2),Round((BSAF.nBSValor * (Case When BSAF.nBSPerDeprecia >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else BSAF.nBSPerDeprecia End)) / BSAF.nBSPerDeprecia,2) -  dbo.GetLogDepAcum(BSAF.nAnio,BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581201', '" & Format(pdDeprecia, gsFormatoFecha) & "')) DepHitMes ,"
'''    '*************************************************************************************************************************************************************************
'        Else
'            sSQL = " Select " & IIf(pbChek, " 0 Chek,", "") & " BS.cBSCod, BSAF.cDescripcion cBSDescripcion, BSAF.cSerie, BSAF.nBSValor,  Round(dbo.GetFactorIPM (  dActivacion, '" & Format(pdDeprecia, gsFormatoFecha) & "'),10) Factor, Round(dbo.GetFactorIPM (dActivacion, '" & Format(pdDeprecia, gsFormatoFecha) & "') * nBSValor,10) V_Ajustado, BSAF.nBSPerDeprecia Vida_Util," _
'                 & " Case When BSAF.nBSPerDeprecia >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else BSAF.nBSPerDeprecia End PeriodosDeprecia," _
'                 & " Case bBaja When 1 Then 0 ELse Round((BSAF.nBSValor * (Case When BSAF.nBSPerDeprecia >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else BSAF.nBSPerDeprecia End)) / BSAF.nBSPerDeprecia,2) End Dep_Hist," _
'                 & " Case bBaja When 1 Then 0 ELse Round(((Round(dbo.GetFactorIPM (dActivacion, '" & Format(pdDeprecia, gsFormatoFecha) & "') * nBSValor,10)) * (Case When BSAF.nBSPerDeprecia >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else BSAF.nBSPerDeprecia End)) / BSAF.nBSPerDeprecia,2) End Dep_Ajustada," _
'                 & " BSAF.cAreCod + BSAF.cAgeCod + ' ' + ARE.cAreaDescripcion + ' ' + AGE.cAgeDescripcion Agencia," _
'                 & " BSAF.nMovNro Nro, ( Select Top 1 cCtaContCodD From CtaBS CBS Inner Join OpeCtaCta OCC On CBS.cCtaContCod = OCC.cCtaContCodH And OCC.cOpeCod = '591201' Where BS.cBScod Like cObjetoCod + '%' And CBS.cOpeCod = '591201' And (cCtaContCodOtroD Is Null Or BS.cBScod like cCtaContCodOtroD + '%'  ) ) cCtaContCodD," _
'                 & " Case bBaja When 1 Then 0 ELse Convert(Decimal(19,2),Round((BSAF.nBSValor * (Case When BSAF.nBSPerDeprecia >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else BSAF.nBSPerDeprecia End)) / BSAF.nBSPerDeprecia,2) -  dbo.GetLogDepAcum(" & pnAnio & ",BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581201', '" & Format(pdDeprecia, gsFormatoFecha) & "')) End Dep_H_Mes ," _
'                 & " Case bBaja When 1 Then 0 ELse Convert(Decimal(19,2),Round(((Round(dbo.GetFactorIPM (dActivacion, '" & Format(pdDeprecia, gsFormatoFecha) & "') * nBSValor,10)) * (Case When BSAF.nBSPerDeprecia >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else BSAF.nBSPerDeprecia End)) / BSAF.nBSPerDeprecia,2) - dbo.GetLogDepAcum(" & pnAnio & ",BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581202', '" & Format(pdDeprecia, gsFormatoFecha) & "')) End Dep_A_Mes ," _
'                 & " Case bBaja When 1 Then 0 ELse Convert(Decimal(19,2),Round(((Round(dbo.GetFactorIPM (dActivacion, '" & Format(pdDeprecia, gsFormatoFecha) & "') * nBSValor,10)) * (Case When BSAF.nBSPerDeprecia >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else BSAF.nBSPerDeprecia End)) / BSAF.nBSPerDeprecia,2) - dbo.GetLogDepAcum(" & pnAnio & ",BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581202', '" & Format(pdDeprecia, gsFormatoFecha) & "')) - " _
'                 & " Convert(Decimal(19,2),Round((BSAF.nBSValor * (Case When BSAF.nBSPerDeprecia >= BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Then BSAF.nBSPerDepAct + Datediff(Month,dBase, '" & Format(pdDeprecia, gsFormatoFecha) & "') Else BSAF.nBSPerDeprecia End)) / BSAF.nBSPerDeprecia,2) -  dbo.GetLogDepAcum(" & pnAnio & ",BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581201', '" & Format(pdDeprecia, gsFormatoFecha) & "')) End Ajuste_Mes ," _
'                 & " IsNull((Select cPersNombre From Persona PE1 WITH (NOLOCK) Where PE1.cPersCod = BSAF.cPersCod ),PE.cPersNombre)  Nombre, Replace(Replace(M.cMovDesc ,char(13),' '),char(10),' ') Descripcion," _
'                 & " Convert(Varchar(10),dActivacion,103) F_Activa, BSAF.cAreCod, BSAF.cAgeCod, PE.cPersCod, Convert(Varchar(10),dCompra,103) F_Compra, bBaja Baja, Convert(Varchar(10),dBaja,103) F_Baja, " _
'                 & " dbo.GetLogDepAcum(" & pnAnio & ",BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581201', '" & Format(pdDeprecia, gsFormatoFecha) & "') Dep_H_Mes_Ant," _
'                 & " dbo.GetLogDepAcum(" & pnAnio & ",BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581202', '" & Format(pdDeprecia, gsFormatoFecha) & "') Dep_A_Mes_Ant," _
'                 & " dbo.GetLogDepAcum(" & pnAnio & ",BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581201', '" & Format(CDate("31/12/" & Trim(Year(pdDeprecia) - 1)), gsFormatoFecha) & "') Dep_H_Ejer_Ant," _
'                 & " dbo.GetLogDepAcum(" & pnAnio & ",BSAF.nMovNro , BSAF.cBSCod, BSAF.cSerie,'581202', '" & Format(CDate("31/12/" & Trim(Year(pdDeprecia) - 1)), gsFormatoFecha) & "') Dep_A_Ejer_Ant" _
'                 & " From BSActivoFijo BSAF WITH (NOLOCK)" _
'                 & " Inner Join BienesServicios BS WITH (NOLOCK) On BS.cBSCod = BSAF.cBSCod" _
'                 & " Left  Join Agencias AGE WITH (NOLOCK) On AGE.cAgeCod = BSAF.cAgeCod" _
'                 & " Inner Join Areas ARE WITH (NOLOCK) On ARE.cAreaCod = BSAF.cAreCod" _
'                 & " Inner Join Mov M WITH (NOLOCK) On M.nMovNro = BSAF.nMovNro" _
'                 & " Left  Join MovGasto MG WITH (NOLOCK) On MG.nMovNro = M.nMovNro" _
'                 & " Left  Join Persona PE WITH (NOLOCK) On Pe.cPersCod = MG.cPersCod Where nAnio <= " & Year(pdDeprecia) & " and ban =  " & pnTipo & "" _
'                 & " ORDER BY BSAF.dCompra"
'        End If
'
'        sSQL = "exec stp_sel_GetAFDeprecia '" & Format(pdDeprecia, "yyyymmdd") & "'," & pnTipoDepreciacion & "," & pnAnio & "," & Year(pdDeprecia) & "," & pnTipo & ",'" & Format(DateAdd("m", 1, pdDeprecia), "yyyymmdd") & "'"

'*** FIN PEAC

        '*** PEAC 20120528
        'sSQL = "exec stp_sel_ReporteDepreciaAF '" & Format(pdDeprecia, "yyyymmdd") & "'," & pnTipoDepreciacion & "," & pnTipo & ",'" & Format(DateAdd("m", 1, pdDeprecia), "yyyymmdd") & "'," & pbYaDepreciado
        '*** PEAC 20120612
        sSQL = "exec stp_sel_ReporteDepreAF '" & Format(pdDeprecia, "yyyymmdd") & "'," & pnTipoDepreciacion & "," & pnTipo & ",'" & Format(DateAdd("m", 1, pdDeprecia), "yyyymmdd") & "'," & pbYaDepreciado
        

        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If

    Set GetAFDeprecia = rs
    Set rs = Nothing
End Function

'Devuelva un RecordSet con las Contrataciones
Public Function GetBSAFDeprecia(ByVal pdDeprecia As Date) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset

    If oConec.AbreConexion() Then
        sSQL = " Select BS.cBSCod, BS.cBSDescripcion,cSerie, nBSValor, Round(dbo.GetFactorIPM (Convert(DateTime,'01/01/' + ltrim(str(BSAF.nAnio))), '" & Format(pdDeprecia, gsFormatoFecha) & "'),10) , Round(dbo.GetFactorIPM (Convert(DateTime,'01/01/' + ltrim(str(BSAF.nAnio))), '" & Format(pdDeprecia, gsFormatoFecha) & "') * nBSValor,10)  , nBSValor * nPorDeprecia / 1200  Depre, Round(nBSValor * " _
             & " Case When Year(dActivacion) <> BSAF.nAnio Then dbo.GetFactorIPM (Convert(DateTime,'01/01/' + ltrim(str(BSAF.nAnio))), '" & Format(pdDeprecia, gsFormatoFecha) & "')" _
             & " Else dbo.GetFactorIPM (dActivacion, '" & Format(pdDeprecia, gsFormatoFecha) & "') End * nPorDeprecia / 1200, 5)  Depre_Ajus, Convert(Varchar(10),dActivacion,103), BSAF.cAreCod + BSAF.cAgeCod, AGE.cAgeDescripcion, BSAF.nMovNro," _
             & " " _
             & " From BSActivoFijo BSAF" _
             & " Inner Join BienesServicios BS On BS.cBSCod = BSAF.cBSCod" _
             & " Left  Join Agencias AGE On AGE.cAgeCod = BSAF.cAgeCod" _
             & " Inner Join Areas ARE On ARE.cAreaCod = BSAF.cAreCod" _
             & " Where nAnio = " & Year(pdDeprecia) & " And" _
             & " (" _
             & "     (Year(dActivacion) = BSAF.nAnio And Month(dActivacion) < " & Month(pdDeprecia) & ")" _
             & " Or  (Year(dActivacion) = BSAF.nAnio)" _
             & " )"
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If

    Set GetBSAFDeprecia = rs
    Set rs = Nothing
End Function

Public Function GetBSSalidasAFBND(pdIni As Date, pdFin As Date) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset

    If oConec.AbreConexion() Then
        sSQL = " Select nAnio,nMovNro, BS.cBSCod, BS.cBSDescripcion, cSerie, nBSValor, nIGV, nBSDeprecia, Case bDepreciado When 1 Then '1' Else '0' End  bDepreciado, Convert(Varchar(10),dActivacion,103) dActivacion, cAreCod + IsNull(AGE.cAgeCod,'')  , ARE.cAreaDescripcion , AGE.cAgeCod, AGE.cAgeDescripcion, nBSPerDeprecia,nBSPerDepAct,cDescripcion, Case bDeprecia When 1 Then '1' Else '0' End bDeprecia,  Case bIntangible When 1 Then '1' Else '0' End  bIntangible, Case bForsado When 1 Then '1' Else '0' End  bForsado,cNroID,cNroPatrimonio, cProveedor, (Select cPersNombre From Persona PE  Where PE.cPersCod = cProveedor) Proveedor, cCtaCont, cFactura " _
             & " From  BSControlSal BSCS" _
             & " Inner Join BienesServicios BS On BSCS.cBSCod = BS.cBSCod" _
             & " Left Join Areas ARE On ARE.cAreaCod = BSCS.cAreCod" _
             & " Left Join Agencias AGE On AGE.cAgeCod = BSCS.cAgeCod" _
             & " Where BSCS.dActivacion Between '" & Format(pdIni, gsFormatoFecha) & "' And '" & Format(DateAdd("d", 1, pdFin), gsFormatoFecha) & "'" _
             & " Order By BS.cBSCod, cSerie"
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If
    
    Set GetBSSalidasAFBND = rs
    Set rs = Nothing
End Function

Public Function GetBSSalidasAFBNDValida(pdIni As Date, pdFin As Date) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset

    If oConec.AbreConexion() Then
        sSQL = " Select nAnio, nMovNro, BS.cBSCod, cSerie, Count(*) Num  " _
             & " From  BSControlSal BSCS" _
             & " Inner Join BienesServicios BS On BSCS.cBSCod = BS.cBSCod" _
             & " Left Join Areas ARE On ARE.cAreaCod = BSCS.cAreCod" _
             & " Left Join Agencias AGE On AGE.cAgeCod = BSCS.cAgeCod" _
             & " Where BSCS.dActivacion Between '" & Format(pdIni, gsFormatoFecha) & "' And '" & Format(DateAdd("d", 1, pdFin), gsFormatoFecha) & "'" _
             & " Group by nAnio, nMovNro, BS.cBSCod, cSerie Having count(*) > 1 Order By BS.cBSCod, cSerie"
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If
    
    Set GetBSSalidasAFBNDValida = rs
    Set rs = Nothing
End Function

Public Function GetBSSalidasAFBNDValidaDeprecia(pdIni As Date, pdFin As Date) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset

    If oConec.AbreConexion() Then
        sSQL = " Select distinct BS.cBSCod, BS.cBSDescripcion From  BSControlSal BSCS" _
             & " Inner Join BienesServicios BS On BSCS.cBSCod = BS.cBSCod" _
             & " Left Join Areas ARE On ARE.cAreaCod = BSCS.cAreCod" _
             & " Left Join Agencias AGE On AGE.cAgeCod = BSCS.cAgeCod" _
             & " Where BSCS.dActivacion Between '" & Format(pdIni, gsFormatoFecha) & "' And '" & Format(DateAdd("d", 1, pdFin), gsFormatoFecha) & "'" _
             & " And (nBSTipDep = 0 OR nPorDeprecia = 0) And BS.cBSCod Like '112%' Order By BS.cBSCod "
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If
    
    Set GetBSSalidasAFBNDValidaDeprecia = rs
    Set rs = Nothing
End Function


Public Function SetBSSalidasAFBND(prRS As ADODB.Recordset, pdIni As Date, pdFin As Date) As Boolean
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset

    If oConec.AbreConexion() Then
        oConec.BeginTrans
        sSQL = " Delete BSControlSal" _
             & " Where dActivacion Between '" & Format(pdIni, gsFormatoFecha) & "' And '" & Format(DateAdd("d", 1, pdFin), gsFormatoFecha) & "'"
        oConec.Ejecutar sSQL
        
        If Not prRS Is Nothing Then
            While Not prRS.EOF
                sSQL = " Insert BSControlSal " _
                     & " (nAnio,nMovNro,cBSCod,cSerie," _
                     & "  nBSValor,nIGV,nBSDeprecia,bDepreciado," _
                     & "  dActivacion,cAreCod,cAgeCod,nBSPerDeprecia," _
                     & "  nBSPerDepAct,cDescripcion,bDeprecia,bIntangible," _
                     & "  bForsado , cNroID, cNroPatrimonio, cProveedor, cCtaCont, cFactura)" _
                     & " Values(" & prRS.Fields(0) & "," & prRS.Fields(1) & ",'" & prRS.Fields(2) & "','" & prRS.Fields(4) & "'" _
                     & " ," & prRS.Fields(5) & "," & prRS.Fields(6) & "," & prRS.Fields(7) & "," & prRS.Fields(8) & "" _
                     & " ,'" & Format(CDate(prRS.Fields(9)), gsFormatoFecha) & "','" & Left(prRS.Fields(10), 3) & "','" & IIf(Mid(prRS.Fields(10), 4, 2) = "", "07", Mid(prRS.Fields(10), 4, 2)) & "'," & prRS.Fields(14) & "" _
                     & " ," & prRS.Fields(15) & ",'" & prRS.Fields(16) & "'," & prRS.Fields(17) & "," & prRS.Fields(18) & "" _
                     & " ," & prRS.Fields(19) & ",'" & prRS.Fields(20) & "','" & prRS.Fields(21) & "','" & prRS.Fields(22) & "','" & prRS.Fields(24) & "','" & prRS.Fields(25) & "')"
                oConec.Ejecutar sSQL
                prRS.MoveNext
            Wend
        End If
        oConec.CommitTrans
        oConec.CierraConexion
    End If

    Set rs = Nothing
End Function

'Devuelva un RecordSet con los asientos ya generados
Public Function GetBSAFAsiento(ByVal psOpeCod As String, ByVal psFecha As String, ByVal psTipo As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset

    If oConec.AbreConexion() Then
        
        '*** PEAC 20120608
'        sSQL = " Select M.nMovnro from Mov M " _
'        & "     Join MovBSAF MB On M.nMovNro = MB.nMovNroref " _
'        & " Where cOpecod = '" & psOpeCod & "' and MB.cTipoActivo = '" & psTipo & "'" _
'        & "     And Substring(cMovNro ,1,6) = '" & psFecha & "' and M.nMovFlag = 0 "
        
        sSQL = "exec stp_sel_BuscaAsientoBSAF '" & psOpeCod & "','" & psTipo & "','" & psFecha & "'"
        '*** FIN PEAC
        
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If

    Set GetBSAFAsiento = rs
    Set rs = Nothing
End Function
'ALPA 20101001****************************************************************
Public Function GetDepreciacion() As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset

    If oConec.AbreConexion() Then
        sSQL = "exec B2_stp_sel_DepreciacionActivos"
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If
    
    Set GetDepreciacion = rs
    Set rs = Nothing
End Function

Public Sub InsertarTipoDepreciacionActivo(ByVal pnDepreTipoActivo As Integer, ByVal pnDeprePorcC As Currency, pnDepreMesC As Integer, ByVal pnDeprePorcT As Currency, ByVal pnDepreMesT As Integer)
    Dim oConec As DConecta
    Set oConec = New DConecta
    If oConec.AbreConexion() Then
        sSQL = "exec B2_stp_upd_DepreciacionActivos " & pnDepreTipoActivo & ", " & pnDeprePorcC & ", " & pnDepreMesC & "," & pnDeprePorcT & ", " & pnDepreMesT
        oConec.CargaRecordSet (sSQL)
        oConec.CierraConexion
    End If
End Sub
'*****************************************************************************

'*** PEAC 20120530
Public Function ObtienePorcentVidaUtlAF(ByVal pnTipoAF As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset

    If oConec.AbreConexion() Then
        sSQL = "exec stp_sel_ObtienePorcentVidaUtlAF " & pnTipoAF
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If
    
    Set ObtienePorcentVidaUtlAF = rs
    Set rs = Nothing
End Function

'*** PEAC 20120601
Public Sub ActualizaInsertaPorcenDepreAF(ByVal pnDepreTipoActivo As Integer, ByVal psAgeCod As String, ByVal pnDeprePorcC As Currency, pnDepreMesC As Integer, ByVal pnDeprePorcT As Currency, ByVal pnDepreMesT As Integer)
    Dim oConec As DConecta
    Set oConec = New DConecta
    If oConec.AbreConexion() Then
        sSQL = "exec stp_UpdIns_PorcentYVidaUtilAF " & pnDepreTipoActivo & ",'" & psAgeCod & "'," & pnDeprePorcC & ", " & pnDepreMesC & "," & pnDeprePorcT & ", " & pnDepreMesT
        oConec.CargaRecordSet (sSQL)
        oConec.CierraConexion
    End If
End Sub

'*** PEAC 20120612
Public Function BuscaMovDepreAF(ByVal psOpeCod As String, ByVal psFecha As String, ByVal psTipo As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset

    If oConec.AbreConexion() Then
        
        sSQL = "exec stp_sel_BuscaMovDepreAF '" & psOpeCod & "','" & psFecha & "'," & psTipo
        
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If

    Set BuscaMovDepreAF = rs
    Set rs = Nothing
End Function

'*** PEAC 20120612
Public Function BuscaMovAsntoCntDepreAF(ByVal psOpeCod As String, ByVal psFecha As String, ByVal psTipo As Integer) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset

    If oConec.AbreConexion() Then
        
        sSQL = "exec stp_sel_BuscaMovAsntoCntDepreAF '" & psOpeCod & "','" & psFecha & "'," & psTipo
        
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If

    Set BuscaMovAsntoCntDepreAF = rs
    Set rs = Nothing
End Function

'*** PEAC 20120615
Public Function BuscaRegAjusteBSAF(ByVal psOpeCod As String, ByVal psFecha As String, ByVal psTipo As String, ByVal psSerie As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset

    If oConec.AbreConexion() Then
        
        sSQL = "exec stp_sel_BuscaRegAjusteBSAF '" & psOpeCod & "','" & psTipo & "','" & psFecha & "','" & psSerie & "'"
        
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If

    Set BuscaRegAjusteBSAF = rs
    Set rs = Nothing
End Function

'*** PEAC 20120618
Public Function MuestraAjustes(ByVal psOpeCod As String, ByVal psFecha As String, ByVal psTipo As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset

    If oConec.AbreConexion() Then
        
        sSQL = "exec stp_sel_MuestraAjustes '" & psOpeCod & "','" & psTipo & "','" & psFecha & "'"
        
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If

    Set MuestraAjustes = rs
    Set rs = Nothing
End Function

'*** PEAC 20120918
Public Function GetBSAFCorrelativoAsiento(ByVal psOpeCod As String, ByVal psFecha As String, ByVal psTipo As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset

    If oConec.AbreConexion() Then

        sSQL = "exec stp_sel_BuscaCorrelatividadDepreAF '" & psOpeCod & "','" & psTipo & "','" & psFecha & "'"
        
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If

    Set GetBSAFCorrelativoAsiento = rs
    Set rs = Nothing
End Function

'*** PEAC 20120921
Public Function GetBSAFTributCorrelativoAsiento(ByVal psOpeCod As String, ByVal psFecha As String, ByVal psTipo As String) As ADODB.Recordset
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    Set oConec = New DConecta
    Set rs = New ADODB.Recordset

    If oConec.AbreConexion() Then

        sSQL = "exec stp_sel_BuscaCorrelatividadDepreTributAF '" & psOpeCod & "','" & psTipo & "','" & psFecha & "'"
        
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
    End If

    Set GetBSAFTributCorrelativoAsiento = rs
    Set rs = Nothing
End Function

