set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go







Create Procedure [dbo].[stp_sel_ReporteResCalCarteraxLineas]

(
	@Opcion  VarChar(1),
	@Periodo Int,
	@tCambio Decimal(5,3)
)

As

if	@opcion='1'
	Begin
		Select Count(*) Cuantos,Month(Max(dFecVig)) Mes, Year(Max(dFecVig)) Anno
		From DbConsolidada..ColocCalifProvTotal 
		Where	Year(dFecha)+Month(dFecha)=(Select Year(Max(dFecVig))+ Month(Max(dFecVig)) From DbConsolidada..CreditoConsol)					
	End

If @Opcion='2'
Begin
(
	Select	Estado,Moneda,Condicion,Tipo,cLinea,Agencia,Cl1.cAbrev Linea,
			Case	When Substring(cLinea,5,1)='1'
					Then 'MN - ' + CL1.cAbrev + ' '+ CL.cDescripcion + ' - ' +Substring(tipo,1,3)
					Else 'ME - ' + CL1.cAbrev + ' '+ CL.cDescripcion + ' - ' +Substring(tipo,1,3)
			End Descripcion,
			Monto_Aprobado=Sum(	Case When moneda='1' Then Aprobado Else Aprobado * @tCambio End),
			Saldo_Capital=Sum(	Case When moneda='1' Then Saldo Else Saldo * @tCambio End),
			Int_Devengado=Sum(	Case When moneda='1' Then Devengado Else Devengado * @tCambio End),
			Int_Suspenso=Sum(	Case When moneda='1' Then Suspenso Else Suspenso * @tCambio End), 
			Provision=Sum(	Case When moneda='1' Then nProvision Else nProvision * @tCambio End) 

	From
		(Select	Case When CC.nPrdEstado In (2201,2203,2205)
					Then 'Creditos Judiciales'
					Else 'Creditos Vigentes'
				End Estado,
				Substring(CC.cCtaCod,9,1) Moneda,
				Case	CC.cRefinan 
						When 'N' Then 'NORMAL'
						WHEN 'R' THEN 'REFINANCIADO'
						ELSE 'NO DISPONIBLE'
				End Condicion,
				K2.cConsDescripcion Tipo,
				Substring(CC.cLineaCred,1,6) cLinea,
				Substring(CC.cCtaCod,4,2) Agencia,
				Sum( CC.nMontoApr) Aprobado,
				Sum( CC.nSaldoCap) Saldo,
				Sum( CC.nIntDev) Devengado,
				Sum( CC.nIntSusp) Suspenso,
				Sum( Cal.nProvision) nProvision
		From		DbConsolidada..CreditoConsol CC
		Inner Join	DbConsolidada..ProductoPersonaConsol PPC On PPC.cCtaCod=CC.cCtaCod AND PPC.nPrdPersRelac = 20
		--Inner Join DbConsolidada..ColocLineaCredito CL On Substring(Cl.cLineaCred,1,6)=Substring(CC.cLineaCred,1,6)
		Inner Join	Persona Per On Per.cPersCod=PPC.cPersCod
		Inner Join	DbConsolidada..ColocCalifProvTotal Cal On Cal.cCtaCod=CC.cCtaCod And Year(dFecha)+Month(Cal.dFecha)=@periodo
		--Inner Join	Constante K On CC.nPrdEstado=K.nConsValor And K.nConsCod=3001
		Inner Join	Constante K2 On Substring(CC.cCtaCod,6,3)=K2.nConsValor And K2.nConsCod=1001
		Where CC.nPrdEstado IN (2020, 2021, 2022, 2030, 2031, 2032 , 2201,2205, 2101,2104, 2106, 2107)
		Group By CC.nPrdEstado,Substring(CC.cCtaCod,9,1),CC.cRefinan,K2.cConsDescripcion,Substring(CC.cLineaCred,1,6),Substring(CC.cCtaCod,4,2),cc.nmontoapr
		) Res
	Inner Join DbConsolidada..ColocLineaCredito CL On Res.cLinea=Substring(CL.cLineaCred,1,6)	And Len(rtrim(CL.cLineaCred))=6
	Inner Join DbConsolidada..ColocLineaCredito CL1 On Substring(Res.cLinea,1,4)=Substring(CL1.cLineaCred,1,4)	And Len(rtrim(CL1.cLineaCred))=4
	Group By Res.Estado,Res.Moneda,Res.Condicion,Res.tipo,Res.clinea,Cl1.cAbrev,Cl.cDescripcion,Res.Agencia



	Union

	Select	Estado,Moneda,Condicion,Tipo,cLinea,Agencia,Cl1.cAbrev Linea,
			Case	When Substring(cLinea,5,1)='1'
					Then 'MN - ' + CL1.cAbrev + ' '+ CL.cDescripcion + ' - ' +Substring(tipo,1,3)
					Else 'ME - ' + CL1.cAbrev + ' '+ CL.cDescripcion + ' - ' +Substring(tipo,1,3)
			End Descripcion,
			Monto_Aprobado=Sum(	Case When moneda='1' Then Aprobado Else Aprobado * @tCambio End),
			Saldo_Capital=Sum(	Case When moneda='1' Then Saldo Else Saldo * @tCambio End),
			Int_Devengado=Sum(	Case When moneda='1' Then Devengado Else Devengado * @tCambio End),
			Int_Suspenso=Sum(	Case When moneda='1' Then Suspenso Else Suspenso * @tCambio End), 
			Provision=Sum(	Case When moneda='1' Then nProvision Else nProvision * @tCambio End) 

	
	From
		(Select	'Creditos Vigentes' Estado,
				Substring(CFC.cCtaCod,9,1) Moneda,
				'NORMAL' Condicion,
				K2.cConsDescripcion Tipo,
				'0101'+ Substring(CFC.cCtaCod,9,1) + '1'  as cLinea,
				Substring(CFC.cCtaCod,4,2) Agencia,
				Sum( CFC.nMontoApr) Aprobado,
				Sum( CFS.nSaldoCap) Saldo,
				Sum( 0) Devengado,
				Sum( 0) Suspenso,
				Sum( Cal.nProvision) nProvision
		From		DbConsolidada..CartaFianzaConsol CFC
		Inner Join	DbConsolidada..ProductoPersonaConsol PPC On PPC.cCtaCod=CFC.cCtaCod AND PPC.nPrdPersRelac = 20
		Inner Join  DbConsolidada..CartaFianzaSaldoConsol CFS On CFS.cCtaCod=CFC.cCtaCod
		Inner Join	Persona Per On Per.cPersCod=PPC.cPersCod
		Inner Join	DbConsolidada..ColocCalifProvTotal Cal On Cal.cCtaCod=CFC.cCtaCod And Year(Cal.dFecha)+Month(Cal.dFecha)=@periodo
		Inner Join	Constante K2 On Substring(CFC.cCtaCod,6,3)=K2.nConsValor And K2.nConsCod=1001
		Where CFC.nPrdEstado IN (2020,2021,2022,2092)
		Group By CFC.nPrdEstado,Substring(CFC.cCtaCod,9,1),K2.cConsDescripcion,Substring(CFC.cCtaCod,4,2),cFC.nmontoapr,CFC.cCTaCod
		)	Res
	Inner Join DbConsolidada..ColocLineaCredito CL On Res.cLinea=Substring(CL.cLineaCred,1,6)	And Len(rtrim(CL.cLineaCred))=6
	Inner Join DbConsolidada..ColocLineaCredito CL1 On Substring(Res.cLinea,1,4)=Substring(CL1.cLineaCred,1,4)	And Len(rtrim(CL1.cLineaCred))=4
	Group By Res.Estado,Res.Moneda,Res.Condicion,Res.tipo,Res.clinea,Cl1.cAbrev,Cl.cDescripcion,Res.Agencia
)	Order By Estado,Moneda,Condicion,Tipo,Linea,Agencia
End


