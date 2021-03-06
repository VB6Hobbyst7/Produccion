set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go

Alter procedure [dbo].[stp_Consolida_ReclasificaIntDevConAlineamiento]
(
@dFecConsol datetime
)
as
begin
	--108711 que hace uso de nRepo108705_ImpCarteraReclasificacion para el Interes

	delete CredSaldosCont where datediff(day,dFecha,@dFecConsol)=0 and cConcepto = 'I'

	create table #CredReclasificaInt
	( 
	  cMoneda Integer, cProducto char(3), cAgencia char(2), cFFinanc char (4), cPlazo char (1), cRefinan char(1),
	  nDemanda int, TotNro real, TotSaldo money, VigNro real, VigSaldo money, Ven1Nro real, Ven1Saldo money, Ven2Nro real, 
	  Ven2Saldo money, JudSDNro real, JudSDSaldo money, JudCDNro real, JudCDSaldo money 
	)

	insert into #CredReclasificaInt
	
	--Declare @dFecConsol as datetime
	--Set @dFecConsol='2008-10-31 00:00:00.000'

	select substring(CC.cLineaCred,5,1) Moneda, substring(CC.cCtaCod,6,3) Producto , substring(CC.cCtaCod,4,2) Agencia, 
		substring(CC.cLineaCred,1,4) FFinanc, substring(CC.cLineaCred, 6,1) Plazo, CC.cRefinan, isnull(nDemanda,2) nDemanda, 
		
		count(CC.cCtaCod) as TotNro, sum(nIntDev) AS TotSaldoCap, 
		
		count(case	when substring(CC.cCtaCod,6,1) ='1' and CC.nPrdEstado not in (2201) and CC.nDiasAtraso <= 15 And CPT.cCalgen Not In ('3','4')then CC.cCtaCod 
				when substring(CC.cCtaCod,6,1) ='2' and CC.nPrdEstado not in (2201) and CC.nDiasAtraso <= 30 And CPT.cCalgen Not In ('3','4') then CC.cCtaCod 
				when substring(CC.cCtaCod,6,1) In ('3','4') and CC.nPrdEstado not in (2201) and CC.nDiasAtraso <= 30 And CPT.cCalgen Not In ('3','4') then CC.cCtaCod 
				end) VigNro, 
		isnull(sum(case when substring(CC.cCtaCod,6,1) ='1' and CC.nPrdEstado not in (2201,2205) and CC.nDiasAtraso <= 15 And CPT.cCalgen Not In ('3','4') then nIntDev 
					 when substring(CC.cCtaCod,6,1) ='2' and CC.nPrdEstado not in (2201,2205) and CC.nDiasAtraso <= 30 And CPT.cCalgen Not In ('3','4') then nIntDev 
					 when substring(CC.cCtaCod,6,1) in('3','4') and CC.nPrdEstado not in (2201,2205) and CC.nDiasAtraso <= 30 And CPT.cCalgen Not In ('3','4') then nIntDev 
					 when substring(CC.cCtaCod,6,1) in('3','4') and CC.nPrdEstado not in (2201,2205) and CC.nDiasAtraso > 30 and CC.nDiasAtraso <= 90 And CPT.cCalgen Not In ('3','4') then 0
					end),0) VigSaldo , 
		
		--by capi 09122008
		count(case	When CPT.cCalgen In ('3','4') and CC.nDiasAtraso <= 30	and CC.nPrdEstado not in (2201,2205) Then CC.cCtaCod
				end) AliNro, 
		isnull(sum(case  When CPT.cCalgen In ('3','4') and CC.nDiasAtraso <= 30	and CC.nPrdEstado not in (2201,2205) Then nIntDev
					end),0) AliSaldo , 
		--end by

		count(case when substring(CC.cCtaCod,6,1) ='1' and CC.nPrdEstado not in (2201,2205) and CC.nDiasAtraso > 15 then CC.cCtaCod 
				when substring(CC.cCtaCod,6,1) ='2' and CC.nPrdEstado not in (2201,2205) and CC.nDiasAtraso > 30 then CC.cCtaCod 
				when substring(CC.cCtaCod,6,1) in('3','4') and CC.nPrdEstado not in (2201,2205) and CC.nDiasAtraso > 30 and CC.nDiasAtraso <= 90 then CC.cCtaCod 
				end) Ven1Nro, 
		isnull(sum(case when substring(CC.cCtaCod,6,1) ='1' and CC.nPrdEstado not in (2201,2205) and CC.nDiasAtraso > 15 then nIntDev 
					 when substring(CC.cCtaCod,6,1) ='2' and CC.nPrdEstado not in (2201,2205) and CC.nDiasAtraso > 30 then nIntDev 
					 when substring(CC.cCtaCod,6,1) in('3','4') and CC.nPrdEstado not in (2201,2205) and CC.nDiasAtraso > 30 and CC.nDiasAtraso <= 90 then nIntDev  
					end),0) Ven1Saldo,

		count(case when substring(CC.cCtaCod,6,1) in('3','4') and CC.nPrdEstado not in (2201,2205) and CC.nDiasAtraso > 90 then CC.cCtaCod end) Ven2Nro, 
		isnull(sum(case when substring(CC.cCtaCod,6,1) in('3','4') and CC.nPrdEstado not in (2201,2205) and CC.nDiasAtraso > 90 then nIntDev end),0) Ven2Saldo, 

		count(case when CC.nPrdEstado in (2201,2205) and isnull(nDemanda,2) = 2 then CC.cCtaCod end) JudSDNro, 
		isnull(sum(case when CC.nPrdEstado in (2201,2205) and IsNull(nDemanda,2) = 2 then nIntDev  end),0) JudSDSaldo, 

		count(case when CC.nPrdEstado in (2201,2205) and IsNull(nDemanda,2) = 1 then CC.cCtaCod end) JudCDNro, 
		isnull(sum(case when CC.nPrdEstado in (2201,2205) and IsNull(nDemanda,2) = 1 then nIntDev end),0) JudCDSaldo 

	from CreditoConsol CC
	--by capi 09122008
	Inner Join DbConsolidada..ColocCalifProvTotal CPT On CC.cCtaCod=CPT.cCtaCod And CPT.dFecha=@dFecConsol
	where CC.nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201,2205)
	group by substring(CC.cLineaCred,5,1), substring(CC.cCtaCod,6,3) , substring(CC.cCtaCod,4,2), 
		substring(CC.cLineaCred,1,4), substring(CC.cLineaCred, 6,1), CC.cRefinan , isnull(nDemanda,2) 
	order by substring(CC.cLineaCred,5,1), substring(CC.cCtaCod,6,3) , substring(CC.cCtaCod,4,2),  
		substring(CC.cLineaCred,1,4), substring(CC.cLineaCred, 6,1), CC.cRefinan , isnull(nDemanda,2)

	create index [Reclasif_1] ON [#CredReclasificaInt]([cMoneda], [cProducto], [cAgencia], [cFFinanc], [cPlazo]) WITH  FILLFACTOR = 90  ON [PRIMARY] 

	--select * from #CredReclasificaInt
	declare @cMoneda char(1) 
	declare @cProducto varchar(10)
	declare @cAgencia varchar(10)
	declare @cFFinanc varchar(10)
	declare @cPlazo varchar(10)
	declare @cRefinan varchar(10)
	declare @nTCredPr int
	declare @nTSaldoCapPr float
	declare @nTNumVigPr int
	declare @nTCapVigPr float
	declare @nTNumVen1Pr int
	declare @nTCapVen1Pr float
	declare @nTNumVen2Pr int
	declare @nTCapVen2Pr float
	declare @nTNumJudSDemPr int
	declare @nTCapJudSDemPr float
	declare @nTNumJudCDemPr int
	declare @nTCapJudCDemPr float
	--by capi 09122008
	declare @nTNumAliPr int
	declare @nTCapAliPr float
	
	--

	--select @dFecConsol = convert(datetime,rtrim(ltrim(nConsSisValor))) from DBCmacMaynas..ConstSistema Where nConsSisCod = 14
	--select * from DBCmacMaynas..ConstSistema Where nConsSisCod = 14


	declare curCar cursor for
		select cMoneda , cProducto, cAgencia, cFFinanc, cPlazo, CC.cRefinan,
			isnull(sum(TotNro),0) AS TotNro, isnull(sum(TotSaldo),0) AS TotSaldo,  
			isnull(sum(VigNro),0) AS VigNro, isnull(sum(VigSaldo),0) as VigSaldo, 
			isnull(sum(Ven1Nro),0) AS Ven1Nro, isnull(sum(Ven1Saldo),0) as Ven1Saldo, 
			isnull(sum(Ven2Nro),0) as Ven2Nro, isnull(sum(Ven2Saldo),0) as Ven2Saldo, 
			isnull(sum(JudSDNro),0) as JudSDNro, isnull(sum(JudSDSaldo),0) as JudSDSaldo, 
			isnull(sum(JudCDNro),0) as JudCDNro, isnull(sum(JudCDSaldo),0) as JudCDSaldo,
			--by capi 09122008
			isnull(sum(AliNro),0) AS AliNro, isnull(sum(AliSaldo),0) as AliSaldo
			--
		from  #CredReclasificaInt 
		group by cMoneda , cProducto, cAgencia, cFFinanc, cPlazo, CC.cRefinan

	open curCar
	fetch next from curCar into @cMoneda,@cProducto,@cAgencia, @cFFinanc, @cPlazo, @cRefinan,
								@nTCredPr, @nTSaldoCapPr, @nTNumVigPr, @nTCapVigPr, @nTNumVen1Pr, @nTCapVen1Pr,
								@nTNumVen2Pr, @nTCapVen2Pr, @nTNumJudSDemPr, @nTCapJudSDemPr, @nTNumJudCDemPr, @nTCapJudCDemPr,
								--by capi 09122008
								@nTNumAliPr, @nTCapAliPr
								--
	while @@fetch_status=0
	begin
		if (@nTCapVigPr <> 0)
			insert into CredSaldosCont(cConcepto, dFecha, cMoneda, cEstado, CC.cRefinanc, cDemandado, cProdCod, cPlazo, cFuenteFinanc, cAgeCod, nSaldo)
			values('I',@dFecConsol,@cMoneda,'1',@cRefinan,'N',@cProducto,@cPlazo,@cFFinanc,@cAgencia,@nTCapVigPr)
	  
	   --Vencidas Adm por Age
	   if (@nTCapVen1Pr + @nTCapVen2Pr <> 0) 
			insert into CredSaldosCont(cConcepto, dFecha, cMoneda, cEstado, CC.cRefinanc, cDemandado, cProdCod, cPlazo, cFuenteFinanc, cAgeCod, nSaldo)
			values('I',@dFecConsol,@cMoneda,'5',@cRefinan,'N',@cProducto,@cPlazo,@cFFinanc,@cAgencia,@nTCapVen1Pr + @nTCapVen2Pr )
	   
	   --Judicial - Administ por Agencia
	   if (@nTCapJudSDemPr <> 0) 
			insert into CredSaldosCont(cConcepto, dFecha, cMoneda, cEstado, CC.cRefinanc, cDemandado, cProdCod, cPlazo, cFuenteFinanc, cAgeCod, nSaldo)
			values('I',@dFecConsol,@cMoneda,'5',@cRefinan,'S',@cProducto,@cPlazo,@cFFinanc,@cAgencia,@nTCapJudSDemPr )
	  
	   --Judicial - Administ por Recup
	   if (@nTCapJudCDemPr <> 0) 
			insert into CredSaldosCont(cConcepto, dFecha, cMoneda, cEstado, CC.cRefinanc, cDemandado, cProdCod, cPlazo, cFuenteFinanc, cAgeCod, nSaldo)
			values('I',@dFecConsol,@cMoneda,'6',@cRefinan,'S',@cProducto,@cPlazo,@cFFinanc,@cAgencia,@nTCapJudCDemPr)
		
		--By capi 09122008
		if (@nTCapAliPr <> 0)
			insert into CredSaldosCont(cConcepto, dFecha, cMoneda, cEstado, CC.cRefinanc, cDemandado, cProdCod, cPlazo, cFuenteFinanc, cAgeCod, nSaldo)
			values('I',@dFecConsol,@cMoneda,'7',@cRefinan,'N',@cProducto,@cPlazo,@cFFinanc,@cAgencia,@nTCapAliPr)
		--End by

		fetch next from curCar into @cMoneda,@cProducto,@cAgencia, @cFFinanc, @cPlazo, @cRefinan,
									@nTCredPr, @nTSaldoCapPr, @nTNumVigPr, @nTCapVigPr, @nTNumVen1Pr, @nTCapVen1Pr,
									@nTNumVen2Pr, @nTCapVen2Pr, @nTNumJudSDemPr, @nTCapJudSDemPr, @nTNumJudCDemPr, @nTCapJudCDemPr,
									--by capi 09122008
									@nTNumAliPr, @nTCapAliPr
									--
	end
	close curCar
	deallocate curCar
end
