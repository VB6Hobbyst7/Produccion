set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go

create procedure [dbo].[stp_sel_ReporteTransferirGarantiasParaAdjudicado]
@gColocEstVigMor int,
@gColocEstVigNorm int,
@gColocEstVigVenc int,
@gColocEstRefMor int,
@gColocEstRefNorm int,
@gColocEstRefVenc int,
@psPersCod varchar(max),
@pscCtaCod varchar(max),
@nValor float,
@nTipo int,
@pscGarantia varchar(max)

as
if @nTipo=1
	begin
		Select PP2.cPersCod,
		cPers=(select cPersNombre from persona where cPersCod=PP2.cPersCod),P.cCtaCod,GARAN.cNumGarant,
		CONST1.cConsDescripcion AS cDesTipoGarantia,GARAN.cDescripcion,GARAN.cDireccion,
		CONST2.cConsDescripcion AS cDesEstado,isnull(GARAN.nEstadoAdju,0) nEstadoAdju,
		GARAN.dEstadoAdju,GARAN.cUsuariAdju,GARAN.nEstado
		From Producto P Inner join Colocaciones C ON P.cCtaCod = C.cCtaCod 
		Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod 
		Left Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = 28
		Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = 20
		inner join colocgarantia COLG ON P.cCtaCod=COLG.cCtaCod
		inner join garantias GARAN ON (COLG.cNumGarant=GARAN.cNumGarant)
		inner join constante CONST1 ON (CONST1.nConsValor=GARAN.nTpoGarantia) AND CONST1.nConsCod=1027
		inner join constante CONST2 ON (CONST2.nConsValor=GARAN.nEstado) AND CONST2.nConsCod=1030
		WHERE  --GARAN.nGarClase=1 
		GARAN.nTpoGarantia in (1,2,3,4,5)
		and PP2.cPersCod = @psPersCod
		AND P.nPrdEstado in (2201,2202,2205,2206)		
		--AND P.nPrdEstado in (@gColocEstVigMor,@gColocEstVigNorm,@gColocEstVigVenc,@gColocEstRefMor ,@gColocEstRefNorm ,@gColocEstRefVenc )		
		and GARAN.nEstadoAdju is null or  GARAN.nEstadoAdju not in (7,8,9,10)
	end 
else
if @nTipo=2 
	begin		
		Select PP2.cPersCod,
		cPers=(select cPersNombre from persona where cPersCod=PP2.cPersCod),P.cCtaCod,GARAN.cNumGarant,
		CONST1.cConsDescripcion AS cDesTipoGarantia,GARAN.cDescripcion,GARAN.cDireccion,
		CONST2.cConsDescripcion AS cDesEstado,isnull(GARAN.nEstadoAdju,0) nEstadoAdju,
		GARAN.dEstadoAdju,GARAN.cUsuariAdju,GARAN.nEstado
		From Producto P Inner join Colocaciones C ON P.cCtaCod = C.cCtaCod 
		Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod 
		Left Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = 28
		Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = 20
		inner join colocgarantia COLG ON P.cCtaCod=COLG.cCtaCod
		inner join garantias GARAN ON (COLG.cNumGarant=GARAN.cNumGarant)
		inner join constante CONST1 ON (CONST1.nConsValor=GARAN.nTpoGarantia) AND CONST1.nConsCod=1027
		inner join constante CONST2 ON (CONST2.nConsValor=GARAN.nEstado) AND CONST2.nConsCod=1030
		WHERE  --GARAN.nGarClase=1
		GARAN.nTpoGarantia in (1,2,3,4,5)
		and P.cCtaCod = @pscCtaCod	
		AND P.nPrdEstado in (2201,2202,2205,2206)		
		--AND P.nPrdEstado in (@gColocEstVigMor,@gColocEstVigNorm,@gColocEstVigVenc,@gColocEstRefMor ,@gColocEstRefNorm ,@gColocEstRefVenc )
		and GARAN.nEstadoAdju is null or  GARAN.nEstadoAdju not in (7,8,9,10)
end
else
if @nTipo=3 
	begin
		Select PP2.cPersCod,
		cPers=(select cPersNombre from persona where cPersCod=PP2.cPersCod),P.cCtaCod,GARAN.cNumGarant,
		CONST1.cConsDescripcion AS cDesTipoGarantia,GARAN.cDescripcion,GARAN.cDireccion,
		CONST2.cConsDescripcion AS cDesEstado,isnull(GARAN.nEstadoAdju,0) nEstadoAdju,
		GARAN.dEstadoAdju,GARAN.cUsuariAdju,GARAN.nEstado
		From Producto P Inner join Colocaciones C ON P.cCtaCod = C.cCtaCod 
		Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod 
		Left Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = 28
		Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = 20
		inner join colocgarantia COLG ON P.cCtaCod=COLG.cCtaCod
		inner join garantias GARAN ON (COLG.cNumGarant=GARAN.cNumGarant)
		inner join constante CONST1 ON (CONST1.nConsValor=GARAN.nTpoGarantia) AND CONST1.nConsCod=1027
		inner join constante CONST2 ON (CONST2.nConsValor=GARAN.nEstado) AND CONST2.nConsCod=1030
		WHERE  --GARAN.nGarClase=1
		GARAN.nTpoGarantia in (1,2,3,4,5)
		and CC.nDiasAtraso  > @nValor		
		--AND P.nPrdEstado in (@gColocEstVigMor,@gColocEstVigNorm,@gColocEstVigVenc,@gColocEstRefMor ,@gColocEstRefNorm ,@gColocEstRefVenc )
		AND P.nPrdEstado in (2201,2202,2205,2206)		
		and GARAN.nEstadoAdju is null or  GARAN.nEstadoAdju not in (7,8,9,10)
	end
else
if @nTipo=4 
	begin
		Select PP2.cPersCod,
		cPers=(select cPersNombre from persona where cPersCod=PP2.cPersCod),P.cCtaCod,GARAN.cNumGarant,
		CONST1.cConsDescripcion AS cDesTipoGarantia,GARAN.cDescripcion,GARAN.cDireccion,
		CONST2.cConsDescripcion AS cDesEstado,isnull(GARAN.nEstadoAdju,0) nEstadoAdju,
		GARAN.dEstadoAdju,GARAN.cUsuariAdju,GARAN.nEstado
		From Producto P Inner join Colocaciones C ON P.cCtaCod = C.cCtaCod 
		Inner Join ColocacCred CC ON P.cCtaCod = CC.cCtaCod 
		Left Join ProductoPersona PP ON P.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = 28
		Inner Join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod And PP2.nPrdPersRelac = 20
		inner join colocgarantia COLG ON P.cCtaCod=COLG.cCtaCod
		inner join garantias GARAN ON (COLG.cNumGarant=GARAN.cNumGarant)
		inner join constante CONST1 ON (CONST1.nConsValor=GARAN.nTpoGarantia) AND CONST1.nConsCod=1027
		inner join constante CONST2 ON (CONST2.nConsValor=GARAN.nEstado) AND CONST2.nConsCod=1030
		WHERE  --GARAN.nGarClase=1
		GARAN.nTpoGarantia in (1,2,3,4,5)
		and GARAN.cNumGarant  = @pscGarantia		
		AND P.nPrdEstado in (@gColocEstVigMor,@gColocEstVigNorm,@gColocEstVigVenc,@gColocEstRefMor ,@gColocEstRefNorm ,@gColocEstRefVenc )
		--and GARAN.nEstadoAdju is null or  GARAN.nEstadoAdju not in (7,8)
	end
