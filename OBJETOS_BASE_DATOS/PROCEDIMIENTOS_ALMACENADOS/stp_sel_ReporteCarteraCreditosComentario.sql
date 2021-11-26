set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go

CREATE procedure [dbo].[stp_sel_ReporteCarteraCreditosComentario]
(
@dFechaFinMes datetime,
@nTipCamb	float,
@cCodInst as varchar(50),
@CCodInstVar as varchar(20),
@FI as varchar(10),
@FF as varchar(10),
@ImporteI as varchar (10),
@ImporteF as varchar(10),
@vAgencia as varchar(50),
@vAnalista as varchar(50),
@vMoneda as varchar(1),
@vTipCre as varchar(10),
@sCodTipCredVar as varchar(20)
)
as
begin
if (@cCodInst = '0') set @cCodInst =  '%' else set @cCodInst =  @cCodInst
if (@cCodInstVar = '3001') set @CCodInstVar =  '3001' else set @CCodInstVar =  @CCodInstVar
if (@FI = '') set @FI =  '00000000' else set @FI =  @FI -- + '%'
if (@FF = '') set @FF =  '99999999' else set @FF =  @FF -- + '%'
if (@ImporteI = '') set @ImporteI =  '-9999999999' else set @ImporteI =  @ImporteI -- + '%'
if (@ImporteF = '') set @ImporteF =  '99999999999999999999999999999999999' else set @ImporteF =  @ImporteF -- + '%'
if (@vAgencia = '') set @vAgencia =  '%' else set @vAgencia =  @vAgencia -- + '%'
if (@vAnalista = '0') set @vAnalista =  '%' else set @vAnalista =  @vAnalista -- + '%'
if (@vMoneda = '') set @vMoneda =  '%' else set @vMoneda =  @vMoneda -- + '%'
if (@vTipCre = '0') set @vTipCre =  '%' else set @vTipCre =  @vTipCre + '%'
if (@sCodTipCredVar = '3001') set @sCodTipCredVar =  '3001' else set @sCodTipCredVar =  @sCodTipCredVar

declare @dFechaMesAnt datetime
select @dFechaMesAnt=max(dFecha) from DBConsolidada..ColocCalifProvTotal 
where year(dFecha)=year(dateadd(month,-1,@dFechaFinMes))
and month(dFecha)=month(dateadd(month,-1,@dFechaFinMes))
--end

	select --SubString(C.cCtaCod,6,3) as XX, 
		CP.cCtaCod,cAge=substring(CP.cCtaCod,4,2),Destino=isnull(C.nDestCre,''), CodCliente=Per.cPersCod , 
		--C.cCodInst,
		Cliente = Per.cPersNombre, cCodDoc = PD.cPersIDNro,
		nMontoApr= C.nMontoApr,
		cEstado= E.cConsDescripcion,nCuotas = C.nCuotasApr,nDiaFijo=isnull(nDiaFijo,0),
		cAnalista= isnull(C.cCodAnalista,''),nTasa= C.nTasaInt,cLineaCredito=CL.cDescripcion,
		dFecVig= C.dFecVig,nSaldoCap=C.nSaldoCap,nCuotaActual=C.nNroProxCuota,
		--A.cAgeDescripcion,
		cTipoPer= TP.cConsDescripcion, cPersCIIU=isnull(Per.cPersCIIU,''),Per.cPersDireccDomicilio,
		cCalifAnterior=case when CT.cCalGen='0' then '0 NORMAL' 
						when CT.cCalGen='1' then '1 CPP'
						when CT.cCalGen='2' then '2 DEFICIENTE'
						when CT.cCalGen='3' then '3 DUDOSO'
						when CT.cCalGen='4' then '4 PERDIDA' end,
		cCalifActual= case when CP.cCalGen='0' then '0 NORMAL' 
						when CP.cCalGen='1' then '1 CPP'
						when CP.cCalGen='2' then '2 DEFICIENTE'
						when CP.cCalGen='3' then '3 DUDOSO'
						when CP.cCalGen='4' then '4 PERDIDA' end,
		nDiasAtraso=C.nDiasAtraso, cFuente=	' CTA:'+ CLE.cCtaCont + F.cDescripcion,
		cPlazo=case when substring(CL.cLineaCred,6,1)='1' then 'CP'else 'LP'end,
		cTipoProd=case when substring(C.cCtaCod,6,3)='305' then 'PRENDARIO' 
						when substring(C.cCtaCod,6,1)='1' then 'COMERCIAL' 
						when substring(C.cCtaCod,6,1)='2' then 'MICROEMPRESA'
						when substring(C.cCtaCod,6,1)='3' then 'ConSUMO'
						when substring(C.cCtaCod,6,1)='4' then 'HIPOTECARIO' end,
		cMoneda=case when substring(C.cCtaCod,9,1)='1'then 'MN'else'ME'end,
		dFecVenc=C.dFecVenc, C.nIntDev,C.nIntSusp,PorProvision=isnull(TAB.nProvision,1),
		nProvisionConRCC=isnull(CP.nProvisionRCC,0),nProvisionSINRCC=CP.nProvision,
		nProvisionAntCRCC= isnull(CT.nProvision,0),
		nProvisionAntSRCC= isnull(CT.nProvisionRCC,0),
		nSaldoDeudor = (select SUM(case when substring(CRE.cCtaCod,9,1)='1'then nSaldoCap else nSaldoCap * @nTipCamb end)
						from DBConsolidada..CreditoConsol CRE inner join DBConsolidada..ProductoPersonaConsol PPC 
							on CRE.cCtaCod = PPC.cCtaCod AND PPC.nPrdPersRelac= 20
						where PPC.cPersCod=Per.cPersCod),
		nGrupoPref = case when isnull(CP.nGarPref,0) >0 then CP.nGarPref else 0 end,
		nGruponoPref = case when isnull(CP.nGarNOPref,0) >0 then CP.nGarNOPref else 0 end,
		nGrupoAutoL = case when isnull(CP.nGarAutoL,0) >0 then CP.nGarAutoL else 0 end,
		cTipoGarantiaCalif=CP.nGarant,
		Alineado=case when CP.cCalSistF is null then 'NO' else 'SI' end,
		CO.cConsDescripcion as cCondicion,
		cCalifSinAlinea= case when CP.cCalSinAli='0' then '0 NORMAL' 
						when CP.cCalSinAli='1' then '1 CPP'
						when CP.cCalSinAli='2' then '2 DEFICIENTE'
						when CP.cCalSinAli='3' then '3 DUDOSO'
						when CP.cCalSinAli='4' then '4 PERDIDA' end,
		cCalifSistF= case when CP.cCalSistF='0' then '0 NORMAL' 
						when CP.cCalSistF='1' then '1 CPP'
						when CP.cCalSistF='2' then '2 DEFICIENTE'
						when CP.cCalSistF='3' then '3 DUDOSO'
						when CP.cCalSistF='4' then '4 PERDIDA' end,
		CP.nProvisionCalSinAli,
		CP.nProvisionCalSistF,
		cClienteUnico=(case when (select count(distinct RD.Cod_Emp) as NroInst 
								from DBConsolidada..RCCTotal RC inner join DBConsolidada..RCCTotalDet RD on RC.Cod_Edu=RD.Cod_Edu
								where right(RD.Cod_Emp,3) <>'109' --Caja Maynas
									and (RC.Cod_Doc_Trib=PD.cPersIDNro or RC.Cod_Doc_Id=PD.cPersIDNro)
								)=0 then 'SI' 
							else 'NO' end),
		Per.cPersCodSbs as Cod_SBS,
		--By Capi 30062008 se agrego campos segun Acta 131-2008
			C.nCuotaApr Cuota,C.nCapVencido Capital_Vencido,
			--Ahora el campo segun condiciones descritas en la misma acta
			Vencido=	Case Left(C.cCtaCod,2)
							When	'10' --Comercial
								Then	Case 
											When C.nDiasAtraso>15 Then C.nCuotaApr
											Else	0
										End
							When	'20' --Mes
								Then	Case 
											When C.nDiasAtraso>30 Then C.nCapVencido
											Else	0
										End
							When	'30' --Consumo
								Then	Case 
											When C.nDiasAtraso>90 Then C.nCapVencido
											When C.nDiasAtraso>30 And C.nDiasAtraso<=90  Then C.nCuotaApr
											Else	0
										End
							When	'40' --Hipotecario
								Then
										Case 
											When C.nDiasAtraso>90 Then C.nCapVencido
											When C.nDiasAtraso>30 And C.nDiasAtraso<=90  Then C.nCuotaApr
											Else	0
										End
						End,
			Institucion=	Case	When c.cCodInst<>'' 
								Then	PC1.cPersNombre
								Else ''
							End, isnull(CE.cEvalObs, '') as Observacion
		--
		
	from ColocCalifProv CP 
		inner join DBConsolidada..CreditoConsol C on CP.cCtaCod = C.cCtaCod
		inner join DBConsolidada..ProductoPersonaConsol PP on C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = 20
		inner join Persona Per on Per.cPersCod = PP.cPersCod
		left join PersID PD on PD.cPersCod = Per.cPersCod 
		and cPersIDTpo=(select MIN(cPersIDTpo)from PersID where cPersCod=Per.cPersCod)
		inner join Constante E on E.nConsValor = C.nPrdEstado AND E.nConsCod = 3001
		Inner join Agencias A ON substring(CP.cCtaCod,4,2) = A.cAgeCod 
		inner join ColocLineaCredito CL on CL.cLineaCred = C.cLineaCred
		inner join Constante TP on TP.nConsValor = Per.nPersPersoneria AND TP.nConsCod = 1002
		inner join ColocLineaCredito F on F.cLineaCred = substring(CL.cLineaCred,1,4) AND LEN(F.cLineaCred)=4
		inner join DBConsolidada..ColocLineaCreditoEquiv CLE on F.cLineaCred = CLE.cLineaCred
		left join ColocCalificaTabla TAB on LEFT(nCalCodTab,1)= substring(CP.cCtaCod,6,1) 
			and TAB.cCalif= CP.cCalGen AND TAB.cRefinan= CP.cRefinan 
			and substring(convert(varchar(4),nCalCodTab),2,1)= (case when CP.nGarant=4 then '0' else '1' end)
		left join  DBConsolidada..ColocCalifProvTotal CT on CT.cCtaCod = CP.cCtaCod and datediff(day,CT.dFecha,@dFechaMesAnt)=0 --mes anterior
		left join Constante CO on C.nCondCre=CO.nConsValor and CO.nConsCod=3015
		--By Capi 30062008 para jalar institucion Convenio
		left join Persona PC1 On PC1.cPersCod=C.cCodInst
		--*********MAVM
		left join ColocEvalCalif CE on CP.cPersCod=CE.cPersCod 

	where C.nPrdEstado IN (2020, 2021, 2022, 2030, 2031, 2032 , 2201,2205, 2101,2104, 2106, 2107)
	and C.cCodInst like @cCodInst -- '%'--1090300189219'
	and convert(varchar(10), C.dFecVig, 112) between @FI and @FF
	and cast(C.nMontoApr as money) between @ImporteI and @ImporteF
	and A.cAgeCod like @vAgencia
	and C.cCodAnalista like @vAnalista
	and substring(C.cCtaCod,9,1) like @vMoneda
	and substring(C.cCtaCod,6,3) like @vTipCre
	
	union all
	select --SubString(CP.cCtaCod,6,3) as XX, 
		CP.cCtaCod,cAge=substring(CP.cCtaCod,4,2),Destino=0, CodCliente=Per.cPersCod, 
		--'',
		Cliente = Per.cPersNombre, cCodDoc = PD.cPersIDNro,
		nMontoApr= C.nMontoApr,
		cEstado= E.cConsDescripcion,nCuotas = 0,nDiaFijo=0,
		cAnalista= isnull(C.cCodAnalista,''),nTasa= 0.00,cLineaCredito='',
		dFecVig= C.dFecVig,nSaldoCap= C.nMontoApr,nCuotaActual= 0,
		--A.cAgeDescripcion,
		cTipoPer= TP.cConsDescripcion, cPersCIIU=isnull(Per.cPersCIIU,''),cPersDireccDomicilio,
		cCalifAnterior=case when CT.cCalGen='0' then '0 NORMAL' 
						when CT.cCalGen='1' then '1 CPP'
						when CT.cCalGen='2' then '2 DEFICIENTE'
						when CT.cCalGen='3' then '3 DUDOSO'
						when CT.cCalGen='4' then '4 PERDIDA' end,
		cCalifActual= case when CP.cCalGen='0' then '0 NORMAL' 
						when CP.cCalGen='1' then '1 CPP'
						when CP.cCalGen='2' then '2 DEFICIENTE'
						when CP.cCalGen='3' then '3 DUDOSO'
						when CP.cCalGen='4' then '4 PERDIDA' end,
		nDiasAtraso= 0, cFuente=	'',
		cPlazo='',
		cTipoProd='CARTA FIANZA',
		cMoneda=case when substring(C.cCtaCod,9,1)='1'then 'MN'else'ME'end,
		dFecVenc=C.dVencApr, nIntDev=0,nIntSusp=0,PorProvision=isnull(TAB.nProvision,1),
		nProvisionConRCC=isnull(CP.nProvisionRCC,0),nProvisionSINRCC=CP.nProvision,
		nProvisionAntCRCC= isnull(CT.nProvision,0),
		nProvisionAntSRCC= isnull(CT.nProvisionRCC,0),
		nSaldoDeudor = (select SUM(case when substring(CRE.cCtaCod,9,1)='1'then nSaldoCap else nSaldoCap * @nTipCamb end)
						from DBConsolidada..CartaFianzaSaldoConsol CRE inner join DBConsolidada..ProductoPersonaConsol PPC 
							on CRE.cCtaCod = PPC.cCtaCod AND PPC.nPrdPersRelac= 20
						where PPC.cPersCod=Per.cPersCod),
		nGrupoPref =  C.nMontoApr,
		nGruponoPref = 0,
		nGrupoAutoL = 0,
		cTipoGarantiaCalif=CP.nGarant,
		Alineado=case when CP.cCalSistF IS NULL then 'NO' else 'SI' end,
		'' as cCondicion,
		cCalifSinAlinea= case when CP.cCalSinAli='0' then '0 NORMAL' 
						when CP.cCalSinAli='1' then '1 CPP'
						when CP.cCalSinAli='2' then '2 DEFICIENTE'
						when CP.cCalSinAli='3' then '3 DUDOSO'
						when CP.cCalSinAli='4' then '4 PERDIDA' end,
		cCalifSistF= case when CP.cCalSistF='0' then '0 NORMAL' 
						when CP.cCalSistF='1' then '1 CPP'
						when CP.cCalSistF='2' then '2 DEFICIENTE'
						when CP.cCalSistF='3' then '3 DUDOSO'
						when CP.cCalSistF='4' then '4 PERDIDA' end,
		CP.nProvisionCalSinAli,
		CP.nProvisionCalSistF,
		cClienteUnico=(case when (select count(distinct RD.Cod_Emp) as NroInst 
								from DBConsolidada..RCCTotal RC inner join DBConsolidada..RCCTotalDet RD on RC.Cod_Edu=RD.Cod_Edu
								where right(RD.Cod_Emp,3) <>'109' --Caja Maynas
									and (RC.Cod_Doc_Trib=PD.cPersIDNro or RC.Cod_Doc_Id=PD.cPersIDNro)
								)=0 then 'SI' 
							else 'NO' end),
		Per.cPersCodSbs as Cod_SBS,
		--By Capi 30062008 se agrego campos segun Acta 131-2008
		Cuota=0,Capital_Vencido=0,Vencido=	0,Institucion='', isnull(CE.cEvalObs, '') as Observacion
		--
	from ColocCalifProv CP 
		inner join DBConsolidada..CartaFianzaConsol  C on CP.cCtaCod = C.cCtaCod
		left join  DBConsolidada..CartaFianzaSaldoConsol CFC on C.cCtaCod = CFC.cCtaCod AND datediff(d,dFecha,@dFechaFinMes)=0 --mes actual
		inner join DBConsolidada..ProductoPersonaConsol PP on C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = 20
		inner join Persona Per on Per.cPersCod = PP.cPersCod
		left join PersTpo PT ON Per.cpersCod = PT.cPersCod
		--*********MAVM
		left join ColocEvalCalif CE on CP.cPersCod=CE.cPersCod 
		left join PersID PD on PD.cPersCod = Per.cPersCod 
			and cPersIDTpo=(select MIN(cPersIDTpo)from PersID where cPersCod=Per.cPersCod)
		inner join Constante E on E.nConsValor = C.nPrdEstado AND E.nConsCod = 3001
		Inner join Agencias A ON substring(CP.cCtaCod,4,2) = A.cAgeCod 
		inner join Constante TP on TP.nConsValor = Per.nPersPersoneria AND TP.nConsCod = 1002
		left join  ColocCalificaTabla TAB on LEFT(nCalCodTab,1)= substring(CP.cCtaCod,6,1) 
			and TAB.cCalif= CP.cCalGen AND TAB.cRefinan= CP.cRefinan 
			and substring(convert(varchar(4),nCalCodTab),2,1)= (case when CP.nGarant=4 then '0' else '1' end)
		left join  DBConsolidada..ColocCalifProvTotal CT on CT.cCtaCod = CP.cCtaCod 
		and datediff(day,CT.dFecha,@dFechaMesAnt)=0 -- ULTIMO DIA DEL MES ANTERIOR
	where C.nPrdEstado IN (2020, 2021, 2022,2092)
and E.nConsCod = @CCodInstVar
and convert(varchar(10), C.dFecVig, 112) between @FI and @FF
and cast(C.nMontoApr as money) between @ImporteI and @ImporteF
and A.cAgeCod like @vAgencia
and C.cCodAnalista like @vAnalista
and substring(C.cCtaCod,9,1) like @vMoneda
and substring(C.cCtaCod,6,3) like @vTipCre
and E.nConsCod = @sCodTipCredVar
order by cTipoProd
end

--select top 10 * from ColocCalifProv