set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go

Create procedure [dbo].[stp_sel_ReporteCarteraCreditosXCliente]
(
@cPersCod as varchar(13),
@CodCta as varchar(50),
@dFechaFinMes datetime,
@nTipCamb	float
)
as
begin
	if (@CodCta = '') set @CodCta =  '%' else set @CodCta =  @CodCta -- + '%'
	declare @dFechaMesAnt datetime
	select @dFechaMesAnt=max(dFecha) from DBConsolidada..ColocCalifProvTotal 
	where year(dFecha)=year(dateadd(month,-1,@dFechaFinMes))
		and month(dFecha)=month(dateadd(month,-1,@dFechaFinMes))


	select CP.cCtaCod,cAge=substring(CP.cCtaCod,4,2),Destino=isnull(C.nDestCre,''), CodCliente=Per.cPersCod , 
		Cliente = Per.cPersNombre, cCodDoc = PD.cPersIDNro,nMontoApr= C.nMontoApr,
		cEstado= E.cConsDescripcion,nCuotas = C.nCuotasApr,nDiaFijo=isnull(nDiaFijo,0),
		cAnalista= isnull(C.cCodAnalista,''),nTasa= C.nTasaInt,cLineaCredito=CL.cDescripcion,
		dFecVig= C.dFecVig,nSaldoCap=C.nSaldoCap,nCuotaActual=C.nNroProxCuota,
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
							End
		--
		
	from ColocCalifProv CP 
		inner join DBConsolidada..CreditoConsol C on CP.cCtaCod = C.cCtaCod
		inner join DBConsolidada..ProductoPersonaConsol PP on C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = 20
		inner join Persona Per on Per.cPersCod = PP.cPersCod
		left join PersID PD on PD.cPersCod = Per.cPersCod 
			and cPersIDTpo=(select MIN(cPersIDTpo)from PersID where cPersCod=Per.cPersCod)
		inner join Constante E on E.nConsValor = C.nPrdEstado AND E.nConsCod = 3001
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
		
	where C.nPrdEstado IN (2020, 2021, 2022, 2030, 2031, 2032 , 2201,2205, 2101,2104, 2106, 2107)
	and Per.cPersCod=@cPersCod and CP.cCtaCod=@CodCta
	union all
	select CP.cCtaCod,cAge=substring(CP.cCtaCod,4,2),Destino=0, CodCliente=Per.cPersCod , 
		Cliente = Per.cPersNombre, cCodDoc = PD.cPersIDNro,nMontoApr= C.nMontoApr,
		cEstado= E.cConsDescripcion,nCuotas = 0,nDiaFijo=0,
		cAnalista= isnull(C.cCodAnalista,''),nTasa= 0.00,cLineaCredito='',
		dFecVig= C.dFecVig,nSaldoCap= C.nMontoApr,nCuotaActual= 0,
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
		Cuota=0,Capital_Vencido=0,Vencido=	0,Institucion=''
		--
	from ColocCalifProv CP 
		inner join DBConsolidada..CartaFianzaConsol  C on CP.cCtaCod = C.cCtaCod
		left join  DBConsolidada..CartaFianzaSaldoConsol CFC on C.cCtaCod = CFC.cCtaCod AND datediff(d,dFecha,@dFechaFinMes)=0 --mes actual
		inner join DBConsolidada..ProductoPersonaConsol PP on C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = 20
		inner join Persona Per on Per.cPersCod = PP.cPersCod
		left join PersID PD on PD.cPersCod = Per.cPersCod 
			and cPersIDTpo=(select MIN(cPersIDTpo)from PersID where cPersCod=Per.cPersCod)
		inner join Constante E on E.nConsValor = C.nPrdEstado AND E.nConsCod = 3001
		inner join Constante TP on TP.nConsValor = Per.nPersPersoneria AND TP.nConsCod = 1002
		left join  ColocCalificaTabla TAB on LEFT(nCalCodTab,1)= substring(CP.cCtaCod,6,1) 
			and TAB.cCalif= CP.cCalGen AND TAB.cRefinan= CP.cRefinan 
			and substring(convert(varchar(4),nCalCodTab),2,1)= (case when CP.nGarant=4 then '0' else '1' end)
		left join  DBConsolidada..ColocCalifProvTotal CT on CT.cCtaCod = CP.cCtaCod 
		and datediff(day,CT.dFecha,@dFechaMesAnt)=0 -- ULTIMO DIA DEL MES ANTERIOR
	where C.nPrdEstado IN (2020, 2021, 2022,2092)
	and Per.cPersCod=@cPersCod and CP.cCtaCod=@CodCta
end