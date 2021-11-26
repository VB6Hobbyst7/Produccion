set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go

Create procedure [dbo].[stp_sel_ReporteCuentaAhorroXCliente]
  @cPersCod as varchar(50)
as

Select  P.cCtaCod
, SUBSTRING(CONVERT(VARCHAR(10), C.dApertura, 102),9,2)+'/'+ SUBSTRING(CONVERT(VARCHAR(10), C.dApertura, 102),6,2)+'/'++SUBSTRING(CONVERT(VARCHAR(10), C.dApertura, 102),1,4)as cFApertura
, SUBSTRING(CONVERT(VARCHAR(10), C.dApertura, 102),1,4) + SUBSTRING(CONVERT(VARCHAR(10), C.dApertura, 102),6,2) + SUBSTRING(CONVERT(VARCHAR(10), C.dApertura, 102),9,2)as cFechaApertura
,Pers.cPersNombre
,A.cAgeDescripcion
,  nSaldoCont= CASE nPrdEstado WHEN 1400 THEN 0 WHEN 1300 THEN 0 ELSE P.nSaldo END
,  nSaldoDisp = CASE nPrdEstado WHEN 1400 THEN 0 WHEN 1300 THEN 0 ELSE nSaldoDisp END
,  C.dApertura
, CN.cConsDescripcion as cTipoAho
,  CN2.cConsDescripcion as cEstado
, CN3.cConsDescripcion as cParticip
, CN4.cConsDescripcion as cMotivo
,  sMoneda = case substring(P.cCtaCod,9,1) when '1' then 'SOLES' when '2' then 'DOLARES' end
, cctacodant= case when R.CCTACODANT is null then '' else r.cctacodant end   
From Producto P 
Inner Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod and  nPrdPersRelac not in (14) 
Inner Join Persona Pers on PP.cPersCod=Pers.cPersCod
Inner Join Captaciones C ON C.cCtaCod = P.cCtaCod 
left join Relcuentas R ON R.cCTACOD=P.CCTACOD    
Inner join Agencias A ON substring(P.cCtaCod,4,2) = A.cAgeCod 
Left join ProductoBloqueos PB ON PB.cCtaCod = P.cCtaCod AND PB.cMovNro = (Select Top 1 cMovNro From ProductoBloqueos Where cCtaCod = P.cCtaCod Order by cMovNro DESC) 
LEFT Join Constante CN ON CONVERT(Int,SUBSTRING(P.cCtaCod,6,3))=CN.nConsValor AND CN.nConsCod = 1001 
LEFT join Constante CN2 ON CN2.nConsValor = P.nPrdEstado AND CN2.nConsCod = 2001 
LEFT join Constante CN3 ON CN3.nConsValor = PP.nPrdPersrelac AND CN3.nConsCod = 2005 
LEFT join Constante CN4 ON CN4.nConsValor = PB.nBlqMotivo AND CN4.nConsCod = 2007  
where CN2.cConsDescripcion='ACTIVA' and nPrdPersRelac='10'
and PP.cPersCod=@cPersCod