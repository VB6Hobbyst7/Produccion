set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go


Create Procedure [dbo].[stp_sel_ReportePFPersoneria]

As 
 
SELECT	PE.cPersDireccDomicilio Domicilio,PR.cCtaCod Cuenta,
		Case When SUBSTRING(PR.cCtaCod,9,1)='1' Then 'MN' Else 'ME' End Moneda,  
		CAP.nSaldoDisp Capital,CAP.nIntAcum AS Interes_Dev,PR.nSaldo AS Saldo_Contable,
		PE.cPersNombre Cliente,PP.cPersCod Codigo, PR.nTasaInteres Tasa, CAPPF.dAuxiliar AS Ultimo_MovPF, 
		a.cagecod Agencia,A.CAGEDESCRIPCION Descripcion,k.cConsDescripcion Personeria,
		CHEQUESENV=(CASE WHEN L.CCTACOD IS NULL THEN 0.00 ELSE L.MONTO  END ),
		NPLAZO Plazo  
FROM Producto PR   
INNER JOIN ProductoPersona PP ON PR.cCtaCod = PP.cCtaCod   
INNER JOIN Captaciones CAP ON PR.cCtaCod = CAP.cCtaCod   
INNER JOIN CaptacPlazoFijo CAPPF ON PR.cCtaCod = CAPPF.cCtaCod 
INNER JOIN Persona PE ON PP.cPersCod = PE.cPersCod   
inner join agencias A ON A.CAGECOD=SUBSTRING(PP.CCTACOD,4,2)  
Left join Constante K On Cap.nPersoneria=K.nConsValor And K.nConsCod=1002 
LEFT JOIN	(	Select MONTO=SUM(NMONTO),CCTACOD 
				from DOCRECCAPTA D   
				INNER JOIN (	SELECT	CNRODOC,CIFCTA 
								FROM DOCRECEST 
								WHERE NESTADO=1 AND CIFCTA IS NOT NULL
							)	E1 ON E1.CNRODOC=D.CNRODOC AND E1.CIFCTA=D.CIFCTA  
				LEFT JOIN	(	SELECT CNRODOC,CIFCTA 
								FROM DOCRECEST 
								WHERE NESTADO=2 AND CIFCTA IS NOT NULL 
							)	E2 ON E2.CNRODOC=D.CNRODOC AND E2.CIFCTA=D.CIFCTA   
				Where E2.cNroDoc Is Null  
				GROUP BY D.CCTACOD
			) L ON L.CCTACOD=PP.CCTACOD   
WHERE	Left(PR.nPrdEstado,2) Not In ('13','14')  
		AND  PP.nPrdPersRelac in (10,12)   
ORDER BY a.cagecod,PR.cCtaCod ,PP.nPrdPersRelac 

