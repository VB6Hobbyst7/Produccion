create procedure stp_sel_ReporteColMejoresClientes
(
@bBitCentral bit,
@nNumero int,
@dFecha datetime,
@nTipCambio decimal(12,4),
@cVigente varchar(50),
@cAgencia varchar(200)
)
as
begin
	If @bBitCentral = 1 
	 begin
	 if @cAgencia = 'Todos' 
		begin
			Select TOP(@nNumero) 
			TA.cPersCod, P.cPersNombre, TA.nSaldo, P.cPersDireccDomicilio, P.cPersTelefono,  
			P.dPersNacCreac, ISNULL(Z.cDesZon,'') Zona ,c.cConsDescripcion ,p.nPersPersoneria Personeria ,
			(select max(t1.cCtaCod) from Colocaciones t1 inner join productopersona t2 on (t1.cCtaCod=t2.cCtaCod) where dVigencia=max(Coloc.dVigencia) and t2.cPersCod=TA.cPersCod and nPrdPersRelac=20) cCtaCod ,
			max(Coloc.dVigencia) dVigencia ,
			(select nMontoCol from Colocaciones where cCtaCod=(select max(cCtaCod) from Colocaciones where dVigencia=max(Coloc.dVigencia))) nMontoCol ,
			(select cConsDescripcion from constante where substring((select max(cCtaCod) from Colocaciones where dVigencia=max(Coloc.dVigencia)),6,3)=NConsValor and nConsCod=1001) cConsDescripcion, 
			cCalifActual= (select case when cCalGen='0' then '0 NORMAL' 
			when cCalGen='1' then '1 CPP' 
			when cCalGen='2' then '2 DEFICIENTE' 
			when cCalGen='3' then '3 DUDOSO' 
			when cCalGen='4' then '4 PERDIDA' end from ColocCalifProv where cCtaCod=(select max(cCtaCod) from Colocaciones where dVigencia=max(Coloc.dVigencia))) 
			FROM ( Select T.cpersCod, SUM(T.nSaldo) nSaldo FROM 
			(  Select PC.cpersCod, A.cctacod, nSaldo = CASE SUBSTRING(A.cctacod,9,1) 
			WHEN '1' THEN A.nSaldoCap WHEN '2' THEN A.nSaldoCap * @nTipCambio  END  
			FROM dbconsolidada..CreditoConsol A INNER JOIN dbconsolidada..productopersonaconsol PC 
			ON A.cctacod = PC.cctacod INNER JOIN Agencias ag ON ag.cAgeCod = substring(a.cCtaCod,4,2) WHERE A.nPrdEstado IN (select Valor from fnc_getTblValoresTexto (@cVigente)) AND PC.nPrdPersRelac =20
			and substring(A.cctacod,6,3) not in (121,221) 		
			union 
			Select PC.cpersCod, A.cctacod, nSaldo = CASE SUBSTRING(A.cctacod,9,1) 
			WHEN '1' THEN A.nMontoApr WHEN '2' THEN A.nMontoApr *@nTipCambio  END  
			FROM dbconsolidada..cartafianzaconsol A INNER JOIN productopersona PC 
			ON A.cctacod = PC.cctacod INNER JOIN Agencias ag ON ag.cAgeCod = substring(a.cCtaCod,4,2) WHERE A.nPrdEstado IN (select Valor from fnc_getTblValoresTexto (@cVigente)) AND PC.nPrdPersRelac =20
			and substring(A.cctacod,6,3) in (121,221) 			
			) T GROUP BY T.cPersCod  
			) TA INNER JOIN Persona P  
			INNER JOIN Constante c ON c.nConsValor =p.nPersPersoneria and nConsCod='1002' 
			LEFT JOIN dbconsolidada..Zonas Z ON P.cPersDireccUbiGeo = Z.cCodZon 
			ON TA.cPersCod = P.cPersCod 
			inner join productoPersona PrP on (P.cPersCod = PrP.cPersCod) 
			inner join producto Pr on (Pr.cCtaCod=Prp.cCtaCod) and Prp.nPrdPersRelac=20 
			inner join Colocaciones Coloc on (Pr.cCtaCod=Coloc.cCtaCod)
			where Pr.nPrdEstado in (select Valor from fnc_getTblValoresTexto (@cVigente))
			group by TA.cPersCod, P.cPersNombre, TA.nSaldo, P.cPersDireccDomicilio, 
			P.cPersTelefono,P.dPersNacCreac, Z.cDesZon, 
			c.cConsDescripcion ,p.nPersPersoneria 		
			ORDER BY TA.nSaldo DESC  
		end
	 else
		begin
		Select TOP (@nNumero) 
			 TA.cPersCod, P.cPersNombre, TA.nSaldo, P.cPersDireccDomicilio, P.cPersTelefono,  
			 P.dPersNacCreac, ISNULL(Z.cDesZon,'') Zona ,c.cConsDescripcion ,p.nPersPersoneria Personeria,
			(select max(t1.cCtaCod) from Colocaciones t1 inner join productopersona t2 on (t1.cCtaCod=t2.cCtaCod) where dVigencia=max(Coloc.dVigencia) and t2.cPersCod=TA.cPersCod and nPrdPersRelac=20) cCtaCod ,
			max(Coloc.dVigencia) dVigencia ,
			(select nMontoCol from Colocaciones where cCtaCod=(select max(cCtaCod) from Colocaciones where dVigencia=max(Coloc.dVigencia))) nMontoCol ,
			 (select cConsDescripcion from constante where substring((select max(cCtaCod) from Colocaciones where dVigencia=max(Coloc.dVigencia)),6,3)=NConsValor and nConsCod=1001) cConsDescripcion, 
			 cCalifActual= (select case when cCalGen='0' then '0 NORMAL' 
				   when cCalGen='1' then '1 CPP' 
				   when cCalGen='2' then '2 DEFICIENTE' 
				   when cCalGen='3' then '3 DUDOSO' 
				   when cCalGen='4' then '4 PERDIDA' end from ColocCalifProv where cCtaCod=(select max(cCtaCod) from Colocaciones where dVigencia=max(Coloc.dVigencia))) 
			 FROM ( Select T.cpersCod, SUM(T.nSaldo) nSaldo FROM 
					(  Select PC.cpersCod, A.cctacod, nSaldo = CASE SUBSTRING(A.cctacod,9,1) 
						WHEN '1' THEN A.nSaldoCap WHEN '2' THEN A.nSaldoCap * @nTipCambio  END  
						FROM dbconsolidada..CreditoConsol A INNER JOIN dbconsolidada..productopersonaconsol PC 
						ON A.cctacod = PC.cctacod INNER JOIN Agencias ag ON ag.cAgeCod = substring(a.cCtaCod,4,2) WHERE A.nPrdEstado IN (select Valor from fnc_getTblValoresTexto (@cVigente)) AND PC.nPrdPersRelac =20
						and substring(A.cctacod,6,3) not in (121,221) 
				and ag.cAgecod =@cAgencia
						union 
					   Select PC.cpersCod, A.cctacod, nSaldo = CASE SUBSTRING(A.cctacod,9,1) 
						WHEN '1' THEN A.nMontoApr WHEN '2' THEN A.nMontoApr *@nTipCambio  END  
						FROM dbconsolidada..cartafianzaconsol A INNER JOIN productopersona PC 
						ON A.cctacod = PC.cctacod INNER JOIN Agencias ag ON ag.cAgeCod = substring(a.cCtaCod,4,2) WHERE A.nPrdEstado IN (select Valor from fnc_getTblValoresTexto (@cVigente)) AND PC.nPrdPersRelac =20
						and substring(A.cctacod,6,3) in (121,221) 
				and ag.cAgecod =@cAgencia
			  ) T GROUP BY T.cPersCod  
			 ) TA INNER JOIN Persona P  
			 INNER JOIN Constante c ON c.nConsValor =p.nPersPersoneria and nConsCod='1002' 
			  LEFT JOIN dbconsolidada..Zonas Z ON P.cPersDireccUbiGeo = Z.cCodZon 
			 ON TA.cPersCod = P.cPersCod 		
			inner join productoPersona PrP on (P.cPersCod = PrP.cPersCod) 
			inner join producto Pr on (Pr.cCtaCod=Prp.cCtaCod) and Prp.nPrdPersRelac=20 
			inner join Colocaciones Coloc on (Pr.cCtaCod=Coloc.cCtaCod)
			where Pr.nPrdEstado in (select Valor from fnc_getTblValoresTexto (@cVigente))
			group by TA.cPersCod, P.cPersNombre, TA.nSaldo, P.cPersDireccDomicilio, 
			P.cPersTelefono,P.dPersNacCreac, Z.cDesZon, 
			c.cConsDescripcion ,p.nPersPersoneria 		
			 ORDER BY TA.nSaldo DESC  
		end   
	end
	 Else
		begin
			Select TOP  (@nNumero)  
				  TA.cCodPers, P.cNomPers, TA.nSaldo, P.cDirPers, P.cTelPers,  
				  P.dFecNac, ISNULL(Z.cDesZon,'') Zona 
				  FROM ( Select T.cCodPers, SUM(T.nSaldo) nSaldo FROM 
				  (  Select PC.cCodPers, A.cCodCta, nSaldo = CASE SUBSTRING(A.cCodCta,6,1) 
				  WHEN '1' THEN A.nSaldoCap WHEN '2' THEN A.nSaldoCap *" & pnTipCambio & "  END  
				  FROM CreditoConsol A INNER JOIN PersCreditoConsol PC 
				  ON A.cCodCta = PC.cCodCta WHERE A.cEstado IN ('F') AND PC.cRelaCta = 'TI' 
				  ) T GROUP BY T.cCodPers  
				  ) TA INNER JOIN dbPersona..Persona P  
				  LEFT JOIN Zonas Z ON P.cCodZon = Z.cCodZon 
				  ON TA.cCodPers = P.cCodPers 
				  ORDER BY TA.nSaldo DESC  
		end   
end