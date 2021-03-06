set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go





ALTER PROCEDURE [dbo].[stp_sel_ReporteCreditosRefinanciados]
(
@cFecFinal As varchar(10),
@cAgencias As VarChar(255),
@nTipoCambio as Decimal(5,3) 
)

AS
begin

select	A.cAgeCod AS sAgencia,
		P.cCtaCod AS cCtaCod,
		sCliente=(P1.cPersNombre),
		--aqui falta documento, se quito porque asi le decidio el usuario IVBA
		--sDocumento=ID.cPersIDNro,
	    M.cConsDescripcion as  sMoneda,
		dVigencia=(C.dVigencia),
		nMontoRefinan= (C.nMontoCol), 
		nMontoRefinanMN= Case	When Substring(P.cCtaCod,9,1)='2' 
								Then Round((C.nMontoCol)*@nTipoCambio,2)
								Else (C.nMontoCol)
						 End,	
		nPlazo=CC.nCuotas,
		nDiasAtraso=CR.nDiasAtraso,
		--aqui falta cuota dificultad
		nSaldoCap = (P.nSaldo),
		nSaldoCapMN = Case	When Substring(P.cCtaCod,9,1)='2' 
							Then Round((P.nSaldo)*@nTipoCambio,2)
							Else (P.nSaldo)
					   End,
		sEstado=SubString(M1.cConsDescripcion,1,3) + Substring(M1.cConsDescripcion,13,20),
	    R.cUser as sAnalista, 
		-- aqui falta pago 6 ultimas cuotas
		K.nCuota	nCuotas6Cuotas,
		k.nDiasAtraso nAtraso6Cuotas,
		-- de aqui sale dias de atraso 6 ultimas cuotas
		sMotivo=M2.cConsDescripcion,
		sCuentaOrigen=Rf.cCtaCodRef,
		nMontoOrigen=C1.nMontoCol,
		sMonedaOrigen=	Case Substring(Rf.cCtaCodRef,9,1) 
							When '1' Then 'SOLES' 
							When '2' Then 'DOLARES'
							ELSE '' 
						eND,
		nPlazoOrigen=CC1.nCuotas		
from Producto P 
inner join Colocaciones C ON C.cCtaCod = P.cCtaCod    
inner join Agencias A ON A.cAgeCod = SUBSTRING(P.cCtaCod,4,2)    
inner join ProductoPersona PP1 ON P.cCtaCod = PP1.cCtaCod AND PP1.nPrdPersRelac = 20    
inner join Persona P1 ON P1.cPersCod = PP1.cPersCod 
inner join ProductoPersona PP2 ON P.cCtaCod = PP2.cCtaCod AND PP2.nPrdPersRelac = 28    
Inner join RRHH R ON R.cPersCod = PP2.cPersCod   
inner join Constante M ON M.nConsValor = SUBSTRING(P.cCtaCod,9,1)AND M.nConsCod=1011 
inner join Constante M1 ON M1.nConsValor = P.nPrdEstado And M1.nConsCod=3001
Left Join	(	Select CC.cCtaCod,Count(*)nCuotas 
				From ColocCalendario CC
				Inner Join ColocacCred C On C.cCtaCod=CC.cCtaCod And C.nNroCalen=CC.nNroCalen
				Where nColocCalendApl=1 --And C.cCtaCod='109022011002231191'	
				Group by CC.cCtaCod
			)	CC On CC.cCtaCod=P.cCtaCod 
Inner Join ColocacCred CR On Cr.cCtaCod=P.cCtaCod 
Left Join ColocacRefinanc Rf On Rf.cCtaCod=P.cCtaCod And Rf.nEstado=1
Left join Constante M2 ON M2.nConsValor = Rf.nMotivoRef And M2.nConsCod=3032
Left Join Colocaciones C1 On C1.cCtaCod=RF.cCtaCodRef
Left Join ColocacEstado CC1 On CC1.cCtaCod=RF.cCtaCodRef And CC1.nPrdEstado=2002
Left Join	(	Select  CC.cCtaCod,nCuota,datediff(day,dVenc,dPago)nDiasAtraso
				From coloccalendario Cal
				Inner Join ColocacCred CC On CC.cCtaCod=Cal.cCtaCod And CC.nNroCalen=Cal.nNroCalen
				Where Cal.nColocCalendApl=1 And nColocCalendEstado=1 And DATEDIFF(day,dPago,@cFecFinal)>=0 --And CC.cCtaCod='109012011001321081' 
				Group by CC.cCtaCod,nCuota,dvenc,dpago--,nColocCalendEstado
				--order by dvenc desc
			) K  On K.cCtaCod=P.cCtaCod
where	P.nPrdEstado in(2030,2031,2032) 
		and DATEDIFF(day,C.dVigencia,@cFecFinal)>=0    
		And Substring(P.cCtaCod,4,2) In (select Valor from dbo.fnc_getTblValoresTexto(@cAgencias))
		--And P.cCtaCod='109011011003014177'
--group by A.cAgeDescripcion,M.cConsDescripcion,R.cUser,P.cCtaCod,CC.nCuotas
order by sAgencia,sMoneda,dvigencia,sAnalista,cCtaCod,sCuentaOrigen,nCuotas6Cuotas desc
End



