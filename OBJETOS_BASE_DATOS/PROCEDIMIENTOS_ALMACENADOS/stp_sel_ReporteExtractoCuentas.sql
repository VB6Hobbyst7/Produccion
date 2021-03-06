set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go




Create Procedure [dbo].[stp_sel_ReporteExtractoCuentas]

(

@cDesde	Varchar(8),
@cHasta Varchar(8),
@cAgencia	varchar(2),
@cProducto varchar(3),
@cInstitucion varchar(13),
@nAccion Int

)

As


if @nAccion=1
	Begin
		Select	Distinct CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) Fecha, 
				O.cOpeDesc Operacion, ISNULL(MD.cDocNro,'') cDocumento , 
				nAbono = CASE	WHEN CT.nCapMovTpo IN (1,7,8) 
								THEN ABS(C.nMonto) 
								ELSE 0 
						 END, 
				nCargo = CASE	WHEN CT.nCapMovTpo IN (6,4,2,3,5) 
								THEN ABS(C.nMonto)*-1 
								ELSE 0 
						 END, 
				 C.nSaldoContable, A.cAgeDescripcion cAgencia, cUsu = RIGHT(M.cMovNro,4), 
				 --CAST(LEFT(cMovNro,14) AS double) cMovNroNue,
				 C.cOpeCod  ,M.cMovNro,C.cCtaCod,C.nSaldoDisponible
		FROM MovDoc MD 
		RIGHT JOIN Mov M  ON MD.nMovNro = M.nMovNro 
		INNER JOIN MovCap C ON M.nMovNro = C.nMovNro    
		INNER JOIN MovCapDet CD ON c.copecod=cd.copecod And C.cCtaCod = CD.cCtaCod AND CD.nMovNro = C.nMovNro 
		INNER JOIN OpeTpo O ON C.cOpeCod = O.cOpeCod
		INNER JOIN CapMovTipo CT ON O.cOpeCod = CT.cOpeCod 
		INNER JOIN Agencias A ON SUBSTRING(M.cMovNro,18,2) = A.cAgeCod 
		Inner Join CaptacAhorros CA On CA.cCtaCod=C.cCtaCod And CA.bOrdPag=1
		Inner Join Producto Pr On Pr.cCtaCod=CA.cCtaCod And  Pr.nPrdEstado Not In (1300,1400)
		WHERE CD.nConceptoCod IN (1,20,10,11)  And SUBSTRING(M.CMOVNRO,1,8) Between @cDesde And @cHasta And SubString(c.cCtaCod,6,3)=@cProducto And substring(c.cCtaCod,4,2)=@cAgencia 
		Order by C.cCtaCod,M.cMovNro,C.cOpeCod
	End
If @nAccion=2
	Begin
		Select	Distinct CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) Fecha, 
				O.cOpeDesc Operacion, ISNULL(MD.cDocNro,'') cDocumento , 
				nAbono = CASE	WHEN CT.nCapMovTpo IN (1,7,8) 
								THEN ABS(C.nMonto) 
								ELSE 0 
						 END, 
				nCargo = CASE	WHEN CT.nCapMovTpo IN (6,4,2,3,5) 
								THEN ABS(C.nMonto)*-1 
								ELSE 0 
						 END, 
				 C.nSaldoContable, A.cAgeDescripcion cAgencia, cUsu = RIGHT(M.cMovNro,4), 
				 --CAST(LEFT(cMovNro,14) AS double) cMovNroNue,
				 C.cOpeCod  ,M.cMovNro,C.cCtaCod,CA.nSaldRetiro as nSaldoDisponible,P.cPersNombre Institucion
		FROM MovDoc MD 
		RIGHT JOIN Mov M  ON MD.nMovNro = M.nMovNro 
		INNER JOIN MovCap C ON M.nMovNro = C.nMovNro    
		INNER JOIN MovCapDet CD ON c.copecod=cd.copecod And C.cCtaCod = CD.cCtaCod AND CD.nMovNro = C.nMovNro 
		INNER JOIN OpeTpo O ON C.cOpeCod = O.cOpeCod
		INNER JOIN CapMovTipo CT ON O.cOpeCod = CT.cOpeCod 
		INNER JOIN Agencias A ON SUBSTRING(M.cMovNro,18,2) = A.cAgeCod 
		Inner Join CaptacCTS CA On CA.cCtaCod=C.cCtaCod 
		Inner Join Persona P on P.cPersCod=CA.cCodInst
		Inner Join Producto Pr On Pr.cCtaCod=CA.cCtaCod And  Pr.nPrdEstado Not In (1300,1400)
		WHERE CD.nConceptoCod IN (1,20,10,11)  And SUBSTRING(M.CMOVNRO,1,8) Between @cDesde And @cHasta And SubString(c.cCtaCod,6,3)=@cProducto And substring(c.cCtaCod,4,2)=@cAgencia
		Order by P.cPersNombre,C.cCtaCod,M.cMovNro,C.cOpeCod
	End
If @nAccion=3
	Begin
		Select	Distinct CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) Fecha, 
				O.cOpeDesc Operacion, ISNULL(MD.cDocNro,'') cDocumento , 
				nAbono = CASE	WHEN CT.nCapMovTpo IN (1,7,8) 
								THEN ABS(C.nMonto) 
								ELSE 0 
						 END, 
				nCargo = CASE	WHEN CT.nCapMovTpo IN (6,4,2,3,5) 
								THEN ABS(C.nMonto)*-1 
								ELSE 0 
						 END, 
				 C.nSaldoContable, A.cAgeDescripcion cAgencia, cUsu = RIGHT(M.cMovNro,4), 
				 --CAST(LEFT(cMovNro,14) AS double) cMovNroNue,
				 C.cOpeCod  ,M.cMovNro,C.cCtaCod,CA.nSaldRetiro as nSaldoDisponible,P.cPersNombre Institucion
		FROM MovDoc MD 
		RIGHT JOIN Mov M  ON MD.nMovNro = M.nMovNro 
		INNER JOIN MovCap C ON M.nMovNro = C.nMovNro    
		INNER JOIN MovCapDet CD ON c.copecod=cd.copecod And C.cCtaCod = CD.cCtaCod AND CD.nMovNro = C.nMovNro 
		INNER JOIN OpeTpo O ON C.cOpeCod = O.cOpeCod
		INNER JOIN CapMovTipo CT ON O.cOpeCod = CT.cOpeCod 
		INNER JOIN Agencias A ON SUBSTRING(M.cMovNro,18,2) = A.cAgeCod 
		Inner Join CaptacCTS CA On CA.cCtaCod=C.cCtaCod
		Inner Join Persona P on P.cPersCod=CA.cCodInst
		Inner Join Producto Pr On Pr.cCtaCod=CA.cCtaCod And  Pr.nPrdEstado Not In (1300,1400)
		WHERE CD.nConceptoCod IN (1,20,10,11)  And SUBSTRING(M.CMOVNRO,1,8) Between @cDesde And @cHasta And SubString(c.cCtaCod,6,3)=@cProducto And substring(c.cCtaCod,4,2)=@cAgencia And CA.cCodInst=@cInstitucion
		Order by P.cPersNombre,C.cCtaCod,M.cMovNro,C.cOpeCod
	End




