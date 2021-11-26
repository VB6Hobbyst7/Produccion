set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go




Create Procedure [dbo].[stp_sel_ReportePagosWesternUnion]

(

@cDesde	Varchar(8),
@cHasta Varchar(8)

)

As


Select	P.cPersNombre Cliente,ID.cPersIDNro DNI,MV.cNroDoc Codigo_Envio,MV.nMovImporte Importe,
		Mv.nMoneda Moneda,Substring(M.cMovNro,18,2)Agencia,Right(M.cMovnro,4)Usuario,
		CONVERT(VARCHAR(12),CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,1,8)),103) + ' ' + SUBSTRING(M.CMOVNRO,9,2)+ ':' + SUBSTRING(M.CMOVNRO,11,2) + ':' + SUBSTRING(M.CMOVNRO,13,2) Fecha
From MovOpeVarias MV 
	Inner Join MovGasto MG On MV.nMovNro=MG.nMovNro
	Inner Join Mov M On M.nMovNro=MG.nMovnro
	Inner Join Persona P On P.cPersCod=MG.cPersCod
	Inner Join PersID ID On ID.cPersCod=P.cPersCod And cPersIDTpo='1'
Where MV.cOpeCod='300517' And Left(M.cMovNro,8) between @cDesde And @cHasta
Order by Agencia,fecha,Cliente

