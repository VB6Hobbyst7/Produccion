set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go

ALTER Procedure [dbo].[Stp_Sel_ReporteReuLavadoDeDinero]
(
@cDesde	Varchar(8),
@cHasta Varchar(8),
@vListaAgencias	varchar(255)
)

As

select * 
from
(
Select		ML.nMovNro nMovNro,isnull(ML.nNroReu,0) nNroReu,SubString(M.cMovNro,1,8)Fecha,MP.cCtaCod Cuenta,isnull(P1.cPersNombre,'') Titular,isnull(P1.cPersDireccDomicilio,'') TitularD,isnull(Id1.cPersIDnro,'') DNI1,
			isnull(P2.cPersNombre,'') Ordena,isnull(P2.cPersDireccDomicilio,'') OrdenaD,isnull(Id2.cPersIDNro,'') DNI2,isnull(P3.cPersNombre,'') Realiza,isnull(P3.cPersDireccDomicilio,'') RealizaD,
			isnull(Id3.cPersIDNro,'') DNI3,isnull(P4.cPersNombre,'') Beneficia,isnull(P4.cPersDireccDomicilio,'') BeneficiaD,isnull(Id4.cPersIDNro,'') DNI4,isnull(P5.cPersNombre,'') VistoBueno,
			Op.cOpeDesc Tipo,MP.nMonto Monto,SubString(Mp.cCtaCod,9,1)Moneda,Right(M.cMovNro,4)Usuario,MP.cOpeCod,
			case ML.nTipoReu
				when 2 then 'Multiple'				
				else 'Significativo'
			end DTipoReu,
			--By capi 27122008 para agregar 2 columnas segun memo 048-2008 LV-DI/CMAC
			ML.nNroReu Numero,
			ML.cOrigenEfectivo Origen
			--
				
From		MovLavDinero ML
Inner Join	Mov M On M.nMovNro=ML.nMovNro And M.nMovFlag=0
Inner Join	OpeTpo Op On Op.cOpeCod=M.cOpeCod
Inner Join	MovCap MP On MP.nMovNro=M.nMovNro
left Join	Persona P1 On ML.cPersCod=P1.cPersCod
left Join	Persona P2 On ML.cPersCodOrd=P2.cPersCod
left Join	Persona P3 On ML.cPersCodRea=P3.cPersCod
left Join	Persona P4 On ML.cPersCodBen=P4.cPersCod
left Join	Persona P5 On ML.cPersCodVisto=P5.cPersCod
Inner join (select Persid1.* 
			from PersID Persid1 inner join (select min(cPersIDTpo) cPersIDTpo,cPersCod 
											from PersID 
											group by cPersCod) Persid2 on (Persid1.cPersCod=Persid2.cPersCod)
											and (Persid1.cPersIDTpo=Persid2.cPersIDTpo)) Id1 On Id1.cPersCod=ML.cPersCod
Inner join (select Persid1.* 
			from PersID Persid1 inner join (select min(cPersIDTpo) cPersIDTpo,cPersCod 
											from PersID 
											group by cPersCod) Persid2 on (Persid1.cPersCod=Persid2.cPersCod)
											and (Persid1.cPersIDTpo=Persid2.cPersIDTpo)) Id2 On Id2.cPersCod=ML.cPersCodOrd
Inner join (select Persid1.* 
			from PersID Persid1 inner join (select min(cPersIDTpo) cPersIDTpo,cPersCod 
											from PersID 
											group by cPersCod) Persid2 on (Persid1.cPersCod=Persid2.cPersCod)
											and (Persid1.cPersIDTpo=Persid2.cPersIDTpo)) Id3 On Id3.cPersCod=ML.cPersCodRea
Inner join (select Persid1.* 
			from PersID Persid1 inner join (select min(cPersIDTpo) cPersIDTpo,cPersCod 
											from PersID 
											group by cPersCod) Persid2 on (Persid1.cPersCod=Persid2.cPersCod)
											and (Persid1.cPersIDTpo=Persid2.cPersIDTpo)) Id4 On Id4.cPersCod=ML.cPersCodBen

Where Substring(M.cMovNro,1,8)>=@cDesde And Substring(M.cMovNro,1,8)<=@cHasta
and substring(M.cMovNro,18,2) in (select Valor from dbo.fnc_getTblValoresTexto(@vListaAgencias))
and (MP.cOpeCod not like '9%') and M.nMovFlag=0
Union

Select		ML.nMovNro nMovNro,isnull(ML.nNroReu,0) nNroReu,SubString(M.cMovNro,1,8)Fecha,MP.cCtaCod Cuenta,isnull(P1.cPersNombre,'') Titular,isnull(P1.cPersDireccDomicilio,'') TitularD,isnull(Id1.cPersIDnro,'') DNI1,
			isnull(P2.cPersNombre,'') Ordena,isnull(P2.cPersDireccDomicilio,'') OrdenaD,isnull(Id2.cPersIDNro,'') DNI2,isnull(P3.cPersNombre,'') Realiza,isnull(P3.cPersDireccDomicilio,'') RealizaD,
			isnull(Id3.cPersIDNro,'') DNI3,isnull(P4.cPersNombre,'') Beneficia,isnull(P4.cPersDireccDomicilio,'') BeneficiaD,isnull(Id4.cPersIDNro,'') DNI4,isnull(P5.cPersNombre,'') VistoBueno,
			Op.cOpeDesc Tipo,MP.nMonto Monto,SubString(Mp.cCtaCod,9,1)Moneda,Right(M.cMovNro,4)Usuario,MP.cOpeCod,
			case ML.nTipoReu
				when 2 then 'Multiple'				
				else 'Significativo'
			end DTipoReu,
			--By capi 27122008 para agregar 2 columnas segun memo 048-2008 LV-DI/CMAC
			ML.nNroReu Numero,
			ML.cOrigenEfectivo Origen
			--
From		MovLavDinero ML
Inner Join	Mov M On M.nMovNro=ML.nMovNro And M.nMovFlag=0
Inner Join	OpeTpo Op On Op.cOpeCod=M.cOpeCod
Inner Join	MovCol MP On MP.nMovNro=M.nMovNro
left Join	Persona P1 On ML.cPersCod=P1.cPersCod
left Join	Persona P2 On ML.cPersCodOrd=P2.cPersCod
left Join	Persona P3 On ML.cPersCodRea=P3.cPersCod
left Join	Persona P4 On ML.cPersCodBen=P4.cPersCod
left Join	Persona P5 On ML.cPersCodVisto=P5.cPersCod

Inner join (select Persid1.* 
			from PersID Persid1 inner join (select min(cPersIDTpo) cPersIDTpo,cPersCod 
											from PersID 
											group by cPersCod) Persid2 on (Persid1.cPersCod=Persid2.cPersCod)
											and (Persid1.cPersIDTpo=Persid2.cPersIDTpo)) Id1 On Id1.cPersCod=ML.cPersCod
Inner join (select Persid1.* 
			from PersID Persid1 inner join (select min(cPersIDTpo) cPersIDTpo,cPersCod 
											from PersID 
											group by cPersCod) Persid2 on (Persid1.cPersCod=Persid2.cPersCod)
											and (Persid1.cPersIDTpo=Persid2.cPersIDTpo)) Id2 On Id2.cPersCod=ML.cPersCodOrd
Inner join (select Persid1.* 
			from PersID Persid1 inner join (select min(cPersIDTpo) cPersIDTpo,cPersCod 
											from PersID 
											group by cPersCod) Persid2 on (Persid1.cPersCod=Persid2.cPersCod)
											and (Persid1.cPersIDTpo=Persid2.cPersIDTpo)) Id3 On Id3.cPersCod=ML.cPersCodRea
Inner join (select Persid1.* 
			from PersID Persid1 inner join (select min(cPersIDTpo) cPersIDTpo,cPersCod 
											from PersID 
											group by cPersCod) Persid2 on (Persid1.cPersCod=Persid2.cPersCod)
											and (Persid1.cPersIDTpo=Persid2.cPersIDTpo)) Id4 On Id4.cPersCod=ML.cPersCodBen
Where Substring(M.cMovNro,1,8)>=@cDesde And Substring(M.cMovNro,1,8)<=@cHasta
and substring(M.cMovNro,18,2) in (select Valor from dbo.fnc_getTblValoresTexto(@vListaAgencias))
and (MP.cOpeCod not like '9%') and M.nMovFlag=0
Union

Select		ML.nMovNro nMovNro,isnull(ML.nNroReu,0) nNroReu,SubString(M.cMovNro,1,8)Fecha,'' as Cuenta,isnull(P1.cPersNombre,'') Titular,isnull(P1.cPersDireccDomicilio,'') TitularD,isnull(Id1.cPersIDnro,'') DNI1,
			isnull(P2.cPersNombre,'') Ordena,isnull(P2.cPersDireccDomicilio,'') OrdenaD,isnull(Id2.cPersIDNro,'') DNI2,isnull(P3.cPersNombre,'') Realiza,isnull(P3.cPersDireccDomicilio,'') RealizaD,
			isnull(Id3.cPersIDNro,'') DNI3,isnull(P4.cPersNombre,'') Beneficia,isnull(P4.cPersDireccDomicilio,'') BeneficiaD,isnull(Id4.cPersIDNro,'') DNI4,isnull(P5.cPersNombre,'') VistoBueno,
			Op.cOpeDesc Tipo,MP.nMovImporte Monto,'2' as Moneda,Right(M.cMovNro,4)Usuario,'000' cOpeCod,
			case ML.nTipoReu
				when 2 then 'Multiple'				
				else 'Significativo'
			end DTipoReu,
			--By capi 27122008 para agregar 2 columnas segun memo 048-2008 LV-DI/CMAC
			ML.nNroReu Numero,
			ML.cOrigenEfectivo Origen
			--
From		MovLavDinero ML
Inner Join	Mov M On M.nMovNro=ML.nMovNro And M.nMovFlag=0
Inner Join	OpeTpo Op On Op.cOpeCod=M.cOpeCod
Inner Join	MovCompraVenta MP On MP.nMovNro=M.nMovNro
left Join	Persona P1 On ML.cPersCod=P1.cPersCod
left Join	Persona P2 On ML.cPersCodOrd=P2.cPersCod
left Join	Persona P3 On ML.cPersCodRea=P3.cPersCod
left Join	Persona P4 On ML.cPersCodBen=P4.cPersCod
left Join	Persona P5 On ML.cPersCodVisto=P5.cPersCod

Inner join (select Persid1.* 
			from PersID Persid1 inner join (select min(cPersIDTpo) cPersIDTpo,cPersCod 
											from PersID 
											group by cPersCod) Persid2 on (Persid1.cPersCod=Persid2.cPersCod)
											and (Persid1.cPersIDTpo=Persid2.cPersIDTpo)) Id1 On Id1.cPersCod=ML.cPersCod
Inner join (select Persid1.* 
			from PersID Persid1 inner join (select min(cPersIDTpo) cPersIDTpo,cPersCod 
											from PersID 
											group by cPersCod) Persid2 on (Persid1.cPersCod=Persid2.cPersCod)
											and (Persid1.cPersIDTpo=Persid2.cPersIDTpo)) Id2 On Id2.cPersCod=ML.cPersCodOrd
Inner join (select Persid1.* 
			from PersID Persid1 inner join (select min(cPersIDTpo) cPersIDTpo,cPersCod 
											from PersID 
											group by cPersCod) Persid2 on (Persid1.cPersCod=Persid2.cPersCod)
											and (Persid1.cPersIDTpo=Persid2.cPersIDTpo)) Id3 On Id3.cPersCod=ML.cPersCodRea
Inner join (select Persid1.* 
			from PersID Persid1 inner join (select min(cPersIDTpo) cPersIDTpo,cPersCod 
											from PersID 
											group by cPersCod) Persid2 on (Persid1.cPersCod=Persid2.cPersCod)
											and (Persid1.cPersIDTpo=Persid2.cPersIDTpo)) Id4 On Id4.cPersCod=ML.cPersCodBen
Where Substring(M.cMovNro,1,8)>=@cDesde And Substring(M.cMovNro,1,8)<=@cHasta
and substring(M.cMovNro,18,2) in (select Valor from dbo.fnc_getTblValoresTexto(@vListaAgencias)) and M.nMovFlag=0
) x
where x.cOpecod not like '107___'
order by x.nNroReu
