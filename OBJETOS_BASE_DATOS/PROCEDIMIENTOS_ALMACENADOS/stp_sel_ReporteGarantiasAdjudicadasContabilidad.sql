create procedure stp_sel_ReporteGarantiasAdjudicadasContabilidad
(
		@dFecSis datetime,
		@cAge varchar(max)
)
as
begin
	declare @numproceso char(4) 	

	set @numproceso='____'
	(select 0 cTipoAdjudicado,'Contratos de joyas adjudicadas' as cDesTipoAdjudicado,
		('Joyas Adjudicadas en el mes de ' + dbo.fnc_DevolverNombreMes(X.nMes) + ' del ' + convert(varchar(4),X.nAnio)) cDescrip,
		sum(X.nAdjValRegistro) nSaldo,0 nInteres,X.nAnio,X.nMes,X.dPrdEstado dFecha,X.cAge,'B' cTipoBien,'A' cOrigen,
		datediff(m,X.dPrdEstado,@dFecSis) nDifMes,Age.cAgeDescripcion
	from
	(
	SELECT	Substring(P.cCtaCod,4,2) cAge,P.dPrdEstado,year(P.dPrdEstado) nAnio ,month(P.dPrdEstado) nMes,CRGDet.cNroProceso,CRGDet.nAdjValRegistro nAdjValRegistro
	FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod
		LEFT  JOIN RELCUENTAS R ON R.CCTACOD=P.CCTACOD  
		INNER JOIN ColocPignoraticio CP ON P.cCtaCod = CP.cCtaCod  
		INNER JOIN ColocPigRGDet CRGDet On P.cCtaCod = CRGDet.cCtaCod  
		LEFT  JOIN ColocPigjoya cpjoy on CP.cCtaCod = cpjoy.cCtaCod   
	WHERE CRGDet.cTpoProceso ='A' AND 
		Substring(P.cCtaCod,4,2) in (select valor from dbo.fnc_getTblValoresTexto(@cAge))  AND 
		CRGDet.cNroProceso LIKE @numproceso	
		And P.nPrdEstado= 2108 And P.dPrdEstado between '2007/11/01' and @dFecSis	
	group by CRGDet.cNroProceso,P.cCtaCod,CRGDet.nAdjValRegistro,P.dPrdEstado

	) X inner join Agencias Age on X.cAge=Age.cAgeCod
	group by X.nAnio,X.nMes,X.dPrdEstado,X.cAge,Age.cAgeDescripcion
	)
	union 
	(select Gara.nTpoGarantia cTipoAdjudicado,Cons1.cConsDescripcion cDesTipoAdjudicado,(Gara.cdescripcion + ' ' + Gara.cDireccion) cDescrip,
	(select nMonto * 
		case nMoneda 
			when 0 then 1
			else isnull((select nValVent from tipoCambio where datediff(d,dFecCamb,@dFecSis)=0),1)
		end
		from GarantiaRemate where cCtaCod=COGA.cCtaCod and nEstadoAdju in (8)) nSaldo,
		isnull((select nInteres * 
		case nMoneda 
			when 0 then 1
			else isnull((select nValVent from tipoCambio where datediff(d,dFecCamb,@dFecSis)=0),1)
		end
		from GarantiaRemate where cCtaCod=COGA.cCtaCod and nEstadoAdju in (8)),0) nInteres,
		year(gara.dEstadoAdju) nAnio,
		month(gara.dEstadoAdju) nMes,gara.dEstadoAdju dFecha,substring(COGA.cCtaCod,4,2) cAge,
		case Gara.nTpoGarantia 
		when 2 then 'B' 
		else 'A'
		end cTipoBien
		,'A' cOrigen,
		datediff(m,gara.dEstadoAdju,@dFecSis) nDifMes,Age.cAgeDescripcion
	from 
		garantias GARA 
		inner join colocGarantia COGA on (GARA.cNumGarant=COGA.cNumGarant)
		inner join producto PROD on (COGA.cCtaCod = PROD.cCtaCod)
		inner join Constante Cons1 on (GARA.nTpoGarantia=Cons1.nConsValor) and Cons1.nConsCod=1027
		inner join Agencias Age on substring(COGA.cCtaCod,4,2)=Age.cAgeCod
	where 	GARA.nEstadoAdju in (8) and datediff(d,Prod.dPrdEstado,@dFecSis)>=0 and Gara.nTpoGarantia in (1,2,3,4,5)
	)
	order by cAge ASC,cTipoAdjudicado,dFecha
end