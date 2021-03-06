set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go





Create Procedure [dbo].[stp_sel_ReporteDetCalCarteraxIntDevSupProv]

(
	@Opcion as varchar(1),	
	@Periodo As Int,
	@tCambio As Decimal(5,3)
)

As

If @Opcion='1' 
	Begin 
		Select Count(*) Cuantos,Month(Max(dFecVig)) Mes, Year(Max(dFecVig)) Anno
		From DbConsolidada..ColocCalifProvTotal 
		Where	Year(dFecha)+Month(dFecha)=(Select Year(Max(dFecVig))+ Month(Max(dFecVig)) From DbConsolidada..CreditoConsol)					
	End


If @Opcion='2'
Begin
(
Select		Substring(CC.cCtaCod,4,2) Agencia,
			Case When CC.nPrdEstado In (2201,2203,2205)
					Then 'Creditos Judiciales'
					Else 'Creditos Vigentes'
			End Estado,
			Substring(CC.cCtaCod,9,1) Moneda,
			K2.cConsDescripcion Tipo,
			Case	CC.cRefinan 
					When 'N' Then 'NORMAL'
					WHEN 'R' THEN 'REFINANCIADO'
					ELSE 'NO DISPONIBLE'
			End Condicion,CC.cCtaCod Cuenta,Per.cPersNombre Cliente,
			PPC.cPersCod Codigo,
			Case	cCalgen
					When '0' Then '0 Normal'
					When '1' Then '1 Cpp'
					When '2' Then '2 Deficiente'
					When '3' Then '3 Dudoso'
					When '4' Then '4 Perdido'
					Else 'No Disponible'
			End Calificacion,
			CC.dFecVig Vigencia,CC.dFecVenc Vencimiento,CC.dFecUltPago Pagado,CCal.dFecVenc Proxima,CC.nDiasAtraso Mora,CC.nTasaInt Tasa,
			Monto_Aprobado=	(Case When Substring(CC.cCtaCod,9,1)='1' Then CC.nMontoApr  Else CC.nMontoApr * @tCambio End),
			Saldo_Capital=(	Case When Substring(CC.cCtaCod,9,1)='1' Then CC.nSaldoCap  Else CC.nSaldoCap * @tCambio End),
			Int_Devengado=(	Case When Substring(CC.cCtaCod,9,1)='1' Then CC.nIntDev  Else CC.nIntDev * @tCambio End),
			Int_Suspenso=(	Case When Substring(CC.cCtaCod,9,1)='1' Then CC.nIntSusp Else CC.nIntSusp * @tCambio End), 
			Provision=(	Case When Substring(CC.cCtaCod,9,1)='1' Then Cal.nProvision Else Cal.nProvision * @tCambio End), 
			Tab.nProvision Tasa_Provision
From		DbConsolidada..CreditoConsol CC
Inner Join	DbConsolidada..ProductoPersonaConsol PPC On PPC.cCtaCod=CC.cCtaCod AND PPC.nPrdPersRelac = 20
Inner Join	Persona Per On Per.cPersCod=PPC.cPersCod
Inner Join	DbConsolidada..ColocCalifProvTotal Cal On Cal.cCtaCod=CC.cCtaCod And Year(dFecha)+Month(Cal.dFecha)=@periodo 
--Inner Join	Constante K On CC.nPrdEstado=K.nConsValor And K.nConsCod=3001
Inner Join	Constante K2 On Substring(CC.cCtaCod,6,3)=K2.nConsValor And K2.nConsCod=1001
Left Join	DbConsolidada..PlanDesPagConsol CCal ON CCal.cCtacod=CC.cCTaCod And cCal.cNroCuo=CC.nNroProxCuota And cCal.nTipo=1
Left join ColocCalificaTabla TAB on LEFT(nCalCodTab,1)= substring(CC.cCtaCod,6,1) 
									and TAB.cCalif= Cal.cCalGen AND TAB.cRefinan= CC.cRefinan 
									and substring(convert(varchar(4),nCalCodTab),2,1)= (case when Cal.nGarant=4 then '0' else '1' end)
Where CC.nPrdEstado IN (2020, 2021, 2022, 2030, 2031, 2032 , 2201,2205, 2101,2104, 2106, 2107) --And CC.cCtaCod In ('109093201000931039','109093041000933228','109093202000932043','109093202000932051','109092011003111121','109092011003111202'  )


Union





Select		Substring(CFC.cCtaCod,4,2) Agencia,
			'Creditos Vigentes' Estado,Substring(CFC.cCtaCod,9,1) Moneda,
			K2.cConsDescripcion Tipo,
			'NORMAL' Condicion,CFC.cCtaCod Cuenta,Per.cPersNombre Cliente,
			PPC.cPersCod Codigo,
			Case	cCalgen
					When '0' Then '0 Normal'
					When '1' Then '1 Cpp'
					When '2' Then '2 Deficiente'
					When '3' Then '3 Dudoso'
					When '4' Then '4 Perdido'
					Else 'No Disponible'
			End Calificacion,
			CFC.dFecVig Vigencia,Null Vencimiento,Null Pagado,Null Proxima,
			0 Mora, 0 Tasa,
			Monto_Aprobado=	(Case When Substring(CFC.cCtaCod,9,1)='1' Then CFC.nMontoApr  Else CFC.nMontoApr * @tCambio End),
			Saldo_Capital=(	Case When Substring(CFC.cCtaCod,9,1)='1' Then CFS.nSaldoCap  Else CFS.nSaldoCap * @tCambio End),
			Int_Devengado=0,
			Int_Suspenso=0,
			Provision=(	Case When Substring(CFC.cCtaCod,9,1)='1' Then Cal.nProvision Else Cal.nProvision * @tCambio End), 
			Tab.nProvision Tasa_Provision
From		DbConsolidada..CartaFianzaConsol CFC
Inner Join	DbConsolidada..ProductoPersonaConsol PPC On PPC.cCtaCod=CFC.cCtaCod AND PPC.nPrdPersRelac = 20
Inner Join  DbConsolidada..CartaFianzaSaldoConsol CFS On CFS.cCtaCod= CFC.cCtaCod
Inner Join	Persona Per On Per.cPersCod=PPC.cPersCod
Inner Join	DbConsolidada..ColocCalifProvTotal Cal On Cal.cCtaCod=CFC.cCtaCod And Year(Cal.dFecha)+Month(Cal.dFecha)=@periodo 
--Inner Join	Constante K On CFC.nPrdEstado=K.nConsValor And K.nConsCod=3001
Inner Join	Constante K2 On Substring(CFC.cCtaCod,6,3)=K2.nConsValor And K2.nConsCod=1001
--Left Join	DbConsolidada..PlanDesPagConsol CCal ON CCal.cCtacod=CFC.cCTaCod And cCal.cNroCuo=CFC.nNroProxCuota And cCal.nTipo=1
Left join ColocCalificaTabla TAB on LEFT(nCalCodTab,1)= substring(CFC.cCtaCod,6,1) 
									and TAB.cCalif= Cal.cCalGen AND TAB.cRefinan= 'N'
									and substring(convert(varchar(4),nCalCodTab),2,1)= (case when Cal.nGarant=4 then '0' else '1' end)
Where CFC.nPrdEstado In (2020,2021,2022,2092) --And CFC.cCtaCod In ('109093201000931039','109093041000933228','109093202000932043','109093202000932051','109092011003111121','109092011003111202')

)
Order By Agencia,Estado,Moneda,Tipo,Condicion

End







