set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go

Create procedure [dbo].[stp_sel_ReporteCuentaCredito]
	@NroCta as varchar(50),
	@cPersCod as varchar(50),
	@vAgencia as varchar(50),
	@vTipCre as varchar(10),
	@vAnalista as varchar(50),
--	@vMoneda as varchar(1),
--	@ImporteI as varchar (10),
--	@ImporteF as varchar(10),
--	@vTasaInteres as varchar(20),
    @FI as varchar(10),
    @FF as varchar(10)
as
begin
	if (@NroCta = '') set @NroCta =  '%' else set @NroCta =  @NroCta -- + '%'
	if (@cPersCod = '') set @cPersCod =  '%' else set @cPersCod =  @cPersCod -- + '%'
	if (@vAgencia = '') set @vAgencia =  '%' else set @vAgencia =  @vAgencia -- + '%'
	if (@vTipCre = '0') set @vTipCre =  '%' else set @vTipCre =  @vTipCre + '%'
	if (@vAnalista = '0') set @vAnalista =  '%' else set @vAnalista =  @vAnalista -- + '%'
	--else set @vTipCre =  @vTipCre  + '% and substring(C.cCtaCod,6,3) <>305' 
--	if (@vMoneda = '') set @vMoneda =  '%' else set @vMoneda =  @vMoneda -- + '%'
--	if (@ImporteI = '') set @ImporteI =  '-9999999999' else set @ImporteI =  @ImporteI -- + '%'
--	if (@ImporteF = '') set @ImporteF =  '99999999999999999999999999999999999' else set @ImporteF =  @ImporteF -- + '%'
--	if (@vTasaInteres = '') set @vTasaInteres =  '%' else set @vTasaInteres =  @vTasaInteres + '%'
	if (@FI = '') set @FI =  '00000000' else set @FI =  @FI -- + '%'
	if (@FF = '') set @FF =  '99999999' else set @FF =  @FF -- + '%'
end

select --top 10 
CP.cCtaCod
,SUBSTRING(CONVERT(VARCHAR(10), C.dFecVig, 102),9,2)+'/'+ SUBSTRING(CONVERT(VARCHAR(10), C.dFecVig, 102),6,2)+'/'+SUBSTRING(CONVERT(VARCHAR(10), C.dFecVig, 102),1,4)as cFDesembolso
,A.cAgeDescripcion
, cAge=substring(CP.cCtaCod,4,2)
, CodCliente=Per.cPersCod
, Cliente = Per.cPersNombre, cCodDoc = PD.cPersIDNro
, nMontoApr= C.nMontoApr
, cEstado= E.cConsDescripcion
--, nCuotas = C.nCuotasApr
--, nDiaFijo=isnull(nDiaFijo,0)
, cAnalista= isnull(C.cCodAnalista,'')
--, nTasa= C.nTasaInt
--, cLineaCredito=CL.cDescripcion
, dFecVig= C.dFecVig
--, nSaldoCap=C.nSaldoCap
--, nCuotaActual=C.nNroProxCuota
, cTipoPer= TP.cConsDescripcion
--, cPersCIIU=isnull(Per.cPersCIIU,'')
, Per.cPersDireccDomicilio
--, nDiasAtraso=C.nDiasAtraso
--, cFuente=	' CTA:'+ CLE.cCtaCont + F.cDescripcion
--, cPlazo=case when substring(CL.cLineaCred,6,1)='1' then 'CP'else 'LP'end
, cTipoProd=case when substring(C.cCtaCod,6,3)='305' then 'PRENDARIO' 
						when substring(C.cCtaCod,6,1)='1' then 'COMERCIAL' 
						when substring(C.cCtaCod,6,1)='2' then 'MICROEMPRESA'
						when substring(C.cCtaCod,6,1)='3' then 'ConSUMO'
						when substring(C.cCtaCod,6,1)='4' then 'HIPOTECARIO' end
, cMoneda=case when substring(C.cCtaCod,9,1)='1'then 'MN'else'ME'end,
		dFecVenc=C.dFecVenc, C.nIntDev,C.nIntSusp,PorProvision=isnull(TAB.nProvision,1),
		nProvisionConRCC=isnull(CP.nProvisionRCC,0),nProvisionSINRCC=CP.nProvision,
		--nProvisionAntCRCC= isnull(CT.nProvision,0),
--		nProvisionAntSRCC= isnull(CT.nProvisionRCC,0),
--		nSaldoDeudor = (select SUM(case when substring(CRE.cCtaCod,9,1)='1'then nSaldoCap else nSaldoCap * @nTipCamb end)
--						from DBConsolidada..CreditoConsol CRE inner join DBConsolidada..ProductoPersonaConsol PPC 
--							on CRE.cCtaCod = PPC.cCtaCod AND PPC.nPrdPersRelac= 20
--						where PPC.cPersCod=Per.cPersCod),
		nGrupoPref = case when isnull(CP.nGarPref,0) >0 then CP.nGarPref else 0 end,
		nGruponoPref = case when isnull(CP.nGarNOPref,0) >0 then CP.nGarNOPref else 0 end,
		nGrupoAutoL = case when isnull(CP.nGarAutoL,0) >0 then CP.nGarAutoL else 0 end,
		cTipoGarantiaCalif=CP.nGarant,
--		Alineado=case when CP.cCalSistF is null then 'NO' else 'SI' end,
--		CO.cConsDescripcion as cCondicion,
--		cCalifSinAlinea= case when CP.cCalSinAli='0' then '0 NORMAL' 
--						when CP.cCalSinAli='1' then '1 CPP'
--						when CP.cCalSinAli='2' then '2 DEFICIENTE'
--						when CP.cCalSinAli='3' then '3 DUDOSO'
--						when CP.cCalSinAli='4' then '4 PERDIDA' end,
--		cCalifSistF= case when CP.cCalSistF='0' then '0 NORMAL' 
--						when CP.cCalSistF='1' then '1 CPP'
--						when CP.cCalSistF='2' then '2 DEFICIENTE'
--						when CP.cCalSistF='3' then '3 DUDOSO'
--						when CP.cCalSistF='4' then '4 PERDIDA' end,
--		CP.nProvisionCalSinAli,
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
		left join Constante CO on C.nCondCre=CO.nConsValor and CO.nConsCod=3015
		left join Persona PC1 On PC1.cPersCod=C.cCodInst
		
	where C.nPrdEstado IN (2020, 2021, 2022, 2030, 2031, 2032 , 2201,2205, 2101,2104, 2106, 2107)
	and (CP.cCtaCod like @NroCta and Per.cPersCod Like @cPersCod and A.cAgeCod like @vAgencia 
	and substring(C.cCtaCod,6,3) like @vTipCre
	and C.cCodAnalista like @vAnalista
	--and PC1.cPersCod like @vAnalista
	and convert(varchar(10), C.dFecVig, 112) between @FI and @FF
		 
	)
--	(P.cCtaCod like @NroCta and Pers.cPersCod Like @cPersCod and A.cAgeCod like @vAgencia
--and (CO.cConsDescripcion like 'comercial%')
--  and substring(P.cCtaCod,9,1) like @vMoneda
--	and cast(nSaldoDisp as money) between @ImporteI and @ImporteF
--	and nTasaInteres like @vTasaInteres

order by cliente