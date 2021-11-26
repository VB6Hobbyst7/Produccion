create procedure stp_sel_ReporteCredEnCJporAbogadoResumen
(
	@SoloUno_Todos INT,
	@nTipoC money,
	@dFecIni char(8),
	@dFecFin char(8),
	@cAge varchar(200),
	@cPersAbo varchar(13) = '%%'
)
As
Begin
	IF @SoloUno_Todos=1 -- CUANDO ESPECIFICA EL CODIGO DEL ABOGADO
	Begin
		select	sum(qr.prestamo) as Prestamo, sum(qr.Saldo) as Saldo, sum(qr.Interes) as Interes,  
				sum(qr.Mora) as Mora, sum(qr.Gasto) as Gasto, sum(qr.Total) as Total, qr.Codabogado , 
				qr.abogado,  qr.nPrdEstado, upper(qr.desestadorecuperacion)desestadorecuperacion
		from(SELECT	C.nMontoCol * case when substring(c.cCtacod,9,1)=1 then 1 else @nTipoC end AS Prestamo, 
					P.nSaldo * case when substring(c.cCtacod,9,1)=1 then 1 else @nTipoC end AS Saldo, 
					CR.nSaldoIntComp * case when substring(c.cCtacod,9,1)=1 then 1 else @nTipoC end AS Interes, 
					CR.nSaldoIntMor * case when substring(c.cCtacod,9,1)=1 then 1 else @nTipoC end AS Mora,  
					CR.nSaldoGasto * case when substring(c.cCtacod,9,1)=1 then 1 else @nTipoC end AS Gasto, 
					(P.nSaldo + CR.nSaldoIntComp + CR.nSaldoIntMor + CR.nSaldoGasto) * 
						case when substring(c.cCtacod,9,1)=1 then 1 else @nTipoC end AS Total,  
					isnull((SELECT PRP.CPERSCOD 
						FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD  
							WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=30),'') CODABOGADO,
					isnull((SELECT PERSO.cPersNombre  
						FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD
							WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=30),'') ABOGADO,
					P.nPrdEstado, 
					(SELECT K.cConsDescripcion
						FROM constante K WHERE k.nconscod = '3001' AND k.nconsvalor = P.nPrdEstado) AS DesEstadoRecuperacion
				FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod 
					INNER JOIN ColocRecup CR ON  C.cCtaCod = CR.cCtaCod and CR.dIngRecup between @dFecIni and @dFecFin
					INNER JOIN ProductoPersona PP ON P.cCtaCod = PP.cCtaCod
				WHERE (PP.nPrdPersRelac=20) 
					And substring(p.cCtaCod,4,2) in (select Valor from dbo.fnc_getTblValoresTexto(@cAge))
					AND (P.nPrdEstado IN('2201','2205'))
					AND ((SELECT PRP.CPERSCOD 
							FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD  
								WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=30)=@cPersAbo)) QR
		group by qr.Codabogado, qr.Abogado, qr.nPrdEstado, qr.DesEstadoRecuperacion
		Order By qr.abogado, qr.desestadorecuperacion desc
	End
	else
	Begin
		select	sum(qr.prestamo) as Prestamo, sum(qr.Saldo) as Saldo, sum(qr.Interes) as Interes,  
				sum(qr.Mora) as Mora, sum(qr.Gasto) as Gasto, sum(qr.Total) as Total, qr.Codabogado , 
				qr.abogado,  qr.nPrdEstado, upper(qr.desestadorecuperacion)desestadorecuperacion
		from(SELECT	C.nMontoCol * case when substring(c.cCtacod,9,1)=1 then 1 else @nTipoC end AS Prestamo, 
					P.nSaldo * case when substring(c.cCtacod,9,1)=1 then 1 else @nTipoC end AS Saldo, 
					CR.nSaldoIntComp * case when substring(c.cCtacod,9,1)=1 then 1 else @nTipoC end AS Interes, 
					CR.nSaldoIntMor * case when substring(c.cCtacod,9,1)=1 then 1 else @nTipoC end AS Mora,  
					CR.nSaldoGasto * case when substring(c.cCtacod,9,1)=1 then 1 else @nTipoC end AS Gasto, 
					(P.nSaldo + CR.nSaldoIntComp + CR.nSaldoIntMor + CR.nSaldoGasto) * 
						case when substring(c.cCtacod,9,1)=1 then 1 else @nTipoC end AS Total,  
					isnull((SELECT PRP.CPERSCOD 
						FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD
							WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=30),'') CODABOGADO,
					isnull((SELECT PERSO.cPersNombre  
						FROM PERSONA PERSO INNER JOIN PRODUCTOPERSONA PRP ON PRP.CPERSCOD = PERSO.CPERSCOD
							WHERE PRP.CCTACOD = PP.CCTACOD AND PRP.NPRDPERSRELAC=30),'') ABOGADO,
					P.nPrdEstado, 
					(SELECT K.cConsDescripcion FROM constante K 
						WHERE k.nconscod = '3001' AND k.nconsvalor = P.nPrdEstado) AS DesEstadoRecuperacion
				FROM Producto P INNER JOIN Colocaciones C ON P.cCtaCod = C.cCtaCod 
					INNER JOIN ColocRecup CR ON  C.cCtaCod = CR.cCtaCod and CR.dIngRecup between @dFecIni and @dFecFin
					INNER JOIN ProductoPersona PP ON P.cCtaCod = PP.cCtaCod
				WHERE (PP.nPrdPersRelac=20) 
					And substring(p.cCtaCod,4,2) in (select Valor from dbo.fnc_getTblValoresTexto(@cAge))
					AND (P.nPrdEstado IN('2201','2205')
					)) QR
		group by qr.Codabogado, qr.Abogado, qr.nPrdEstado, qr.DesEstadoRecuperacion
		Order By qr.abogado, qr.desestadorecuperacion desc
	End
End
