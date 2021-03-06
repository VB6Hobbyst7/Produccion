
ALTER Procedure stp_upd_ActualizaPersona
(
	@cPersNombre varchar(200),@dPersNacCreac datetime,@cPersDireccUbiGeo varchar(12),
	@cPersDireccDomicilio varchar(100),
	@cPersDireccCondicion varchar(1),@cnPersValComDomicilio money,@cPersTelefono varchar(100),
	@cPersTelefono2 varchar(100),
	@cPersEmail varchar(50),@nPersPersoneria int,@cPersCIIU varchar(7),@cPersEstado varchar(2),
	@cPersCodSbs char(10),@nPersRelaInst int,@dPersIngRuc datetime,@dPersIniActi datetime,
	@nNumDependi int,@cActComple varchar(50),@nNumPtosVta int,@cActiGiro varchar(100),
	@nPersTipoComp int,@nPersTipoSistInfor int,@nPersCadenaProd int,@nMonedaPatri int,
	@nPersIngresoProm money,@psPersCod varchar(13), @sUltimaActualizacion varchar(25),
	@cTipoActualizacion varchar(2),@cPersRefDomicilio varchar(100)
)
as
begin
	UPDATE Persona SET 
		cPersNombre = @cPersNombre,	dPersNacCreac = @dPersNacCreac,
		cPersDireccUbiGeo = @cPersDireccUbiGeo,	cPersDireccDomicilio = @cPersDireccDomicilio,
		cPersDireccCondicion = @cPersDireccCondicion,	nPersValComDomicilio = @cnPersValComDomicilio,
		cPersTelefono = @cPersTelefono,	cPersTelefono2 = @cPersTelefono2,
		cPersEmail = @cPersEmail,	nPersPersoneria = @nPersPersoneria,
		cPersCIIU = @cPersCIIU,	cPersEstado = @cPersEstado,
		cPersCodSbs = @cPersCodSbs,	nPersRelaInst = @nPersRelaInst,
		dPersIngRuc = @dPersIngRuc,	dPersIniActi = @dPersIniActi,
		nNumDependi = @nNumDependi,	cActComple = @cActComple,
		nNumPtosVta = @nNumPtosVta,	cActiGiro = @cActiGiro,
		nPersTipCompe = @nPersTipoComp,	nPersTipSistInform = @nPersTipoSistInfor,
		nPersTipCadeProd = @nPersCadenaProd,	nPersMoneyPatri = @nMonedaPatri,
		nPersIngresoProm=@nPersIngresoProm,	cUltimaActualizacion=@sUltimaActualizacion,
		cTipoActualizacion = @cTipoActualizacion, cPersRefDomicilio=@cPersRefDomicilio
	Where cPersCod = @psPersCod
End

