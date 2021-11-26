Create procedure stp_sel_ReporteCuentaAhorro
	@NroCta as varchar(50),
	@cPersCod as varchar(50),
	@vAgencia as varchar(50),
	@vCodProd as varchar(10),
	@vMoneda as varchar(1),
	@ImporteI as varchar (10),
	@ImporteF as varchar(10),
	@vTasaInteres as varchar(20),
    @FI as varchar(10),
    @FF as varchar(10)
as
begin
	if (@NroCta = '') set @NroCta =  '%' else set @NroCta =  @NroCta -- + '%'
	if (@cPersCod = '') set @cPersCod =  '%' else set @cPersCod =  @cPersCod -- + '%'
	if (@vAgencia = '') set @vAgencia =  '%' else set @vAgencia =  @vAgencia -- + '%'
	if (@vCodProd = '') set @vCodProd =  '%' else set @vCodProd =  @vCodProd -- + '%'
	if (@vMoneda = '') set @vMoneda =  '%' else set @vMoneda =  @vMoneda -- + '%'
	if (@ImporteI = '') set @ImporteI =  '-9999999999' else set @ImporteI =  @ImporteI -- + '%'
	if (@ImporteF = '') set @ImporteF =  '99999999999999999999999999999999999' else set @ImporteF =  @ImporteF -- + '%'
	if (@vTasaInteres = '') set @vTasaInteres =  '%' else set @vTasaInteres =  @vTasaInteres + '%'
	if (@FI = '') set @FI =  '00000000' else set @FI =  @FI -- + '%'
	if (@FF = '') set @FF =  '99999999' else set @FF =  @FF -- + '%'
end
Select
C.cCtaCod,
UG.cUbiGeoDescripcion
, nTasaInteres
,A.cAgeCod
, SUBSTRING(CONVERT(VARCHAR(10), C.dApertura, 102),9,2)+'/'+ SUBSTRING(CONVERT(VARCHAR(10), C.dApertura, 102),6,2)+'/'++SUBSTRING(CONVERT(VARCHAR(10), C.dApertura, 102),1,4)as cFApertura
, SUBSTRING(CONVERT(VARCHAR(10), C.dApertura, 102),1,4) + SUBSTRING(CONVERT(VARCHAR(10), C.dApertura, 102),6,2) ++ SUBSTRING(CONVERT(VARCHAR(10), C.dApertura, 102),9,2)as cFechaApertura
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
Inner Join UbicacionGeografica UG on A.cUbiGeoCod=UG.cUbiGeoCod
Left join ProductoBloqueos PB ON PB.cCtaCod = P.cCtaCod AND PB.cMovNro = (Select Top 1 cMovNro From ProductoBloqueos Where cCtaCod = P.cCtaCod Order by cMovNro DESC) 
LEFT Join Constante CN ON CONVERT(Int,SUBSTRING(P.cCtaCod,6,3))=CN.nConsValor AND CN.nConsCod = 1001 
LEFT join Constante CN2 ON CN2.nConsValor = P.nPrdEstado AND CN2.nConsCod = 2001 
LEFT join Constante CN3 ON CN3.nConsValor = PP.nPrdPersrelac AND CN3.nConsCod = 2005 
LEFT join Constante CN4 ON CN4.nConsValor = PB.nBlqMotivo AND CN4.nConsCod = 2007
  
where CN2.cConsDescripcion='ACTIVA' and nPrdPersRelac='10'

and (P.cCtaCod like @NroCta and Pers.cPersCod Like @cPersCod and A.cAgeCod like @vAgencia
and CN.nConsValor like @vCodProd and substring(P.cCtaCod,9,1) like @vMoneda
and cast(nSaldoDisp as money) between @ImporteI and @ImporteF
and nTasaInteres like @vTasaInteres
--and (SUBSTRING(CONVERT(VARCHAR(10), C.dApertura, 102),9,2)+'/'+ SUBSTRING(CONVERT(VARCHAR(10), C.dApertura, 102),6,2)+'/'+SUBSTRING(CONVERT(VARCHAR(10), C.dApertura, 102),1,4)) between @FI and @FF
and convert(varchar(10), dApertura, 112) between @FI and @FF
)
order By cPersNombre

