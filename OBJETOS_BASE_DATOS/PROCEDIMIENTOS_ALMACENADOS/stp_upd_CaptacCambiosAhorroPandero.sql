
set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go

Create Procedure [dbo].[stp_upd_CaptacCambiosAhorroPandero]
	@dFechaProceso	DATETIME,
	@sUsuario	VARCHAR(4),
	@sAgencia	VARCHAR(2)
AS



DECLARE @sMovNro VARCHAR(25)
EXEC sp_GeneraMovNro @dFechaProceso, @sAgencia, @sUsuario, @sMovNro OUTPUT


---para ahorro pandero registra en CapCambioTasa el cambio de tasa generado
Insert Into CapCambioTasa (cCtaCod,nTasaAnterior,nTasaCambio,cRegistro,cActiva)
Select Pr.cCtaCod,Pr.nTasaInteres,nTasaInteresNuevo=(	Select nTasaValor 
														From CaptacTasas 
														Where	nTasaProd =Substring(Pr.cCtaCod,6,3)
														And nTasaMon = Substring(Pr.cCtaCod,9,1)
														And cCodAge = Substring(Pr.cCtaCod,4,2) 
														And nTasaTpo = 100 
														And nTpoPrograma = 0  
														And cOrdPag='0'
														And cActiva='1'
													) ,
		@sMovNro,'S'
From Producto Pr
Inner Join	(	Select MC.cCtaCod,Sum(nMonto)nDeposito From Mov M
				Inner Join MovCap MC On M.nMovNro=MC.nMovNro
				Inner Join Captaciones C On C.cCtaCod=MC.cCtaCod And C.nTpoPrograma=4
				Where Left(MC.cOpeCod,4)='2002' And datediff(day,cast(left(M.cMovNro,8) as datetime),'2008-11-10')<=60 And M.nMovFlag=0
				Group by MC.cCtaCod
			)	P	On P.cCtaCod=Pr.cCtaCod
Where	P.nDeposito<100--luego cambiar(Select nParValor From Parametro Where nParProd=2000 And nParCod=2093)
		And Pr.nPrdEstado Not In (1300,1400)
---

---para ahorro pandero cambia tasa de pandero a ahorro corriente
Update Pr Set nTasaInteres=(	Select nTasaValor 
								From CaptacTasas 
								Where	nTasaProd =Substring(Pr.cCtaCod,6,3)
										And nTasaMon = Substring(Pr.cCtaCod,9,1)
										And cCodAge = Substring(Pr.cCtaCod,4,2) 
										And nTasaTpo = 100 
										And nTpoPrograma = 0  
										And cOrdPag='0'
										And cActiva='1'
							) 
From Producto Pr
Inner Join	(	Select MC.cCtaCod,Sum(nMonto)nDeposito From Mov M
				Inner Join MovCap MC On M.nMovNro=MC.nMovNro
				Inner Join Captaciones C On C.cCtaCod=MC.cCtaCod And C.nTpoPrograma=4
				Where Left(MC.cOpeCod,4)='2002' And datediff(day,cast(left(M.cMovNro,8) as datetime),'2008-11-10')<=60 And M.nMovFlag=0
				Group by MC.cCtaCod
			)	P	On P.cCtaCod=Pr.cCtaCod
Where	P.nDeposito<(Select nParValor From Parametro Where nParProd=2000 And nParCod=2093)
		And Pr.nPrdEstado Not In (1300,1400)
---

---para ahorro pandero Cambia a Producto ahorro corriente
Update Cpt Set nTpoProgramaAnt=4,nTpoPrograma=0
From Captaciones Cpt
Inner Join Producto Pr On Pr.cCtaCod=Cpt.cCtaCod
Inner Join	(	Select MC.cCtaCod,Sum(nMonto)nDeposito From Mov M
				Inner Join MovCap MC On M.nMovNro=MC.nMovNro
				Inner Join Captaciones C On C.cCtaCod=MC.cCtaCod And C.nTpoPrograma=4
				Where Left(MC.cOpeCod,4)='2002' And datediff(day,cast(left(M.cMovNro,8) as datetime),'2008-11-10')<=60 And M.nMovFlag=0
				Group by MC.cCtaCod
			)	P	On P.cCtaCod=Cpt.cCtaCod
Where	P.nDeposito<(Select nParValor From Parametro Where nParProd=2000 And nParCod=2093)
		And Pr.nPrdEstado Not In (1300,1400)
---




