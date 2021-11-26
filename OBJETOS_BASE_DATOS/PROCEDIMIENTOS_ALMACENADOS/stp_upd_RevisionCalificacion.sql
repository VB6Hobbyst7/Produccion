set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go

CREATE procedure [dbo].[stp_upd_RevisionCalificacion]
(
	@RevisionId int,
	@CodPers varchar(13),
	@CodCta varchar(50),
	@FCierre char(10),
	@TCambio money,
	@FRegistro char(10),
	@CAnalista varchar(4),
	@Giro varchar(250),
	
	@FSDCMAC char(10),
	@MontoCMAC money,
	@TMonedaCMAC char (50),
	@FSDSF char(10),
	@MontoSF money,
	@TMonedaSF char (50),
	
	@Norm char(10),
	@CPP char(10),
	@Defic char(10),
	@Dud char(10),
	@Perd char(10),

	@CalificacionCMAC varchar(10),
	@CalificacionSF varchar(10),
	@CalificacionOCI varchar(10),
	@Situacion varchar(250),
	@Desarrollo varchar(250),
	@Garantia varchar(250),
	@Informacion varchar(250),
	@Evaluacion varchar(250),
	@Comentario varchar(250),
	@Conclusion varchar(250),
	@Estado int
)
as
begin
update RevisionCalificacion
set cPersCod=@CodPers
, vCodCta=@CodCta
, cFCierre=@FCierre
, mTCambio=@TCambio
, cFRegistro=@FRegistro
, vCAnalista=@CAnalista
, vGiro=@Giro
, cFSDCMAC=@FSDCMAC, mMontoCMAC=@MontoCMAC, cTMonedaCMAC=@TMonedaCMAC
, cFSDSF=@FSDSF, mMontoSF=@MontoSF, cTMonedaSF=@TMonedaSF
,cPNorm=@Norm, cPCPP=@CPP, cPDefic=@Defic, cPDud=@Dud, cPPerd=@Perd
, vCalificacionCMAC=@CalificacionCMAC
, vCalificacionSF=@CalificacionSF
, vCalificacionOCI=@CalificacionOCI
, vSituacion=@Situacion
, vDesarrollo=@Desarrollo
, vGarantia=@Garantia
, vInformacion=@Informacion
, vEvaluacion=@Evaluacion
, vComentario=@Comentario
, vConclusion=@Conclusion
, iEstado=@Estado
where iRevisionId=@RevisionId
end




