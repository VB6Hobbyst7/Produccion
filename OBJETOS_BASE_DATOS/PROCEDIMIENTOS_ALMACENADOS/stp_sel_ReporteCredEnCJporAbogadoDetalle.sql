create procedure stp_sel_ReporteCredEnCJporAbogadoDetalle
(
	@SoloUno_Todos INT,
	@dFecSis char(8),
	@dFecIni char(8),
	@dFecFin char(8),
	@cAge varchar(200),
	@cPersAbo varchar(13) = '%%'
)
As
Begin
	IF @SoloUno_Todos=1 -- CUANDO ESPECIFICA EL CODIGO DEL ABOGADO
	begin
		SELECT	P.cCtaCod, substring(P.cCtaCod, 9,1) as cMoneda, C.dVigencia, CR.dIngRecup, isnull(C.cLineaCred,'')cLineaCred, 
				CLC.cDescripcion, SUBSTRING(P.cCtaCod, 6,3) AS cCodTipoCredito, Persona.cPersCod,  Persona.cPersNombre, 
				Persona.cPersDireccUbiGeo, 
				isnull((SELECT PrP.cPersCod FROM ProductoPersona PrP 
					WHERE PrP.cCtaCod = PP.cCtaCod AND nPrdPersRelac=30),'') AS CodAnalista, 
				isnull((SELECT PERSO.cPersNombre  FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD 
					WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=30),'') AS ABOGADO,  
				isnull((SELECT cUser FROM PERSONA PERSO INNER JOIN RRHH RH ON PERSO.CPERSCOD=RH.CPERSCOD
					INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD  
					WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=28),'') AS ANALISTA, 
				C.nMontoCol AS Prestamo, P.nSaldo AS Saldo,  
				CR.nSaldoIntComp + dbo.ColocRecupCalculoInteresActual(P.cCtaCod,@dFecSis) AS Interes, 
				CR.nSaldoIntMor + dbo.ColocRecupCalculoMoraActual(P.cCtaCod,@dFecSis) AS Mora, 
				CR.nSaldoGasto AS Gasto, SUBSTRING(P.cCtaCod, 9,1) AS Moneda, 
				CON.cConsDescripcion AS sMoneda, 
				substring(P.cCtaCod, 6,3) cCodTipoCredito,
				(SELECT COO.cConsDescripcion FROM Constante COO  
					WHERE substring(P.cCtaCod, 6,3) = COO.nConsValor AND COO.nConsCod=1001) AS cDesTipoCredito, 
				PP.nPrdPersRelac,UPPER(CON1.cConsDescripcion)cEstado
		FROM Producto P INNER JOIN CONSTANTE CON1 ON P.NPRDESTADO=CON1.NCONSVALOR AND CON1.nconscod = 3001
			INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod 
			INNER JOIN ColocRecup CR ON C.cCtaCod = CR.cCtaCod and CR.dIngRecup between @dFecIni and @dFecFin
			INNER JOIN ColocLineaCredito CLC ON C.cLineaCred = CLC.cLineaCred 
			INNER JOIN ProductoPersona PP ON P.cCtaCod = PP.cCtaCod 
			INNER JOIN Persona ON PP.cPersCod = Persona.cPersCod 
			INNER JOIN Constante CON ON  substring(P.cCtaCod, 9,1) = CON.nConsValor  
		Where (PP.NPRDPERSRELAC=20) And (CON.nconscod=1011)
			AND (P.nPrdEstado IN('2201','2205'))
			And substring(p.cCtaCod,4,2) in (select Valor from dbo.fnc_getTblValoresTexto(@cAge))
			AND ((Select CPERSCOD FROM  PRODUCTOPERSONA PRP WHERE PRP.CCTACOD = PP.CCTACOD 
					AND PRP.NPRDPERSRELAC = 30) like @cPersAbo)
		ORDER BY (SELECT PrP.cPersCod FROM ProductoPersona PrP WHERE PrP.cCtaCod = PP.cCtaCod AND nPrdPersRelac=30),  
		SUBSTRING(P.cCtaCod, 6,3), SUBSTRING(P.cCtaCod, 9, 1), P.cCtaCod
	End
	else
	Begin
		SELECT	P.cCtaCod, substring(P.cCtaCod, 9,1) as cMoneda, C.dVigencia, CR.dIngRecup, isnull(C.cLineaCred,'')cLineaCred, 
				CLC.cDescripcion, SUBSTRING(P.cCtaCod, 6,3) AS cCodTipoCredito, Persona.cPersCod,  Persona.cPersNombre, 
				Persona.cPersDireccUbiGeo, 
				isnull((SELECT PrP.cPersCod FROM ProductoPersona PrP 
					WHERE PrP.cCtaCod = PP.cCtaCod AND nPrdPersRelac=30),'') AS CodAnalista, 
				isnull((SELECT PERSO.cPersNombre  FROM PERSONA PERSO 
					INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD 
					WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=30),'') AS ABOGADO,  
				isnull((SELECT cUser FROM PERSONA PERSO   
					INNER JOIN RRHH RH ON PERSO.CPERSCOD=RH.CPERSCOD  
					INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD  
					WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=28),'') AS ANALISTA, 
				C.nMontoCol AS Prestamo, P.nSaldo AS Saldo,  
				CR.nSaldoIntComp + dbo.ColocRecupCalculoInteresActual(P.cCtaCod,@dFecSis) AS Interes, 
				CR.nSaldoIntMor + dbo.ColocRecupCalculoMoraActual(P.cCtaCod,@dFecSis) AS Mora, 
				CR.nSaldoGasto AS Gasto, SUBSTRING(P.cCtaCod, 9,1) AS Moneda, 
				CON.cConsDescripcion AS sMoneda, 
				substring(P.cCtaCod, 6,3) cCodTipoCredito,
				(SELECT COO.cConsDescripcion FROM Constante COO  
					WHERE substring(P.cCtaCod, 6,3) = COO.nConsValor AND COO.nConsCod=1001) AS cDesTipoCredito, 
				PP.nPrdPersRelac,UPPER(CON1.cConsDescripcion)cEstado
		FROM Producto P INNER JOIN CONSTANTE CON1 ON P.NPRDESTADO=CON1.NCONSVALOR AND CON1.nconscod = 3001
			INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod 
			INNER JOIN ColocRecup CR ON C.cCtaCod = CR.cCtaCod and CR.dIngRecup between @dFecIni and @dFecFin
			INNER JOIN ColocLineaCredito CLC ON C.cLineaCred = CLC.cLineaCred 
			INNER JOIN ProductoPersona PP ON P.cCtaCod = PP.cCtaCod 
			INNER JOIN Persona ON PP.cPersCod = Persona.cPersCod 
			INNER JOIN Constante CON ON  substring(P.cCtaCod, 9,1) = CON.nConsValor  
		Where (PP.NPRDPERSRELAC=20) And (CON.nconscod=1011) 
			AND (P.nPrdEstado IN('2201','2205'))
			And substring(p.cCtaCod,4,2) in (select Valor from dbo.fnc_getTblValoresTexto(@cAge))
		ORDER BY (SELECT PrP.cPersCod FROM ProductoPersona PrP WHERE PrP.cCtaCod = PP.cCtaCod AND nPrdPersRelac=30),  
		SUBSTRING(P.cCtaCod, 6,3), SUBSTRING(P.cCtaCod, 9, 1), P.cCtaCod
	End
End