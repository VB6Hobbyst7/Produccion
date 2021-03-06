ALTER procedure [dbo].[stp_sel_ReporteCarteraCreditos] ਍⠀ ഀ
@dFechaFinMes datetime, ਍䀀渀吀椀瀀䌀愀洀戀ऀ昀氀漀愀琀 ഀ
) ਍愀猀 ഀ
begin ਍ऀ搀攀挀氀愀爀攀 䀀搀䘀攀挀栀愀䴀攀猀䄀渀琀 搀愀琀攀琀椀洀攀 ഀ
	select @dFechaMesAnt=max(dFecha) from DBConsolidada..ColocCalifProvTotal  ਍ऀ眀栀攀爀攀 礀攀愀爀⠀搀䘀攀挀栀愀⤀㴀礀攀愀爀⠀搀愀琀攀愀搀搀⠀洀漀渀琀栀Ⰰⴀ㄀Ⰰ䀀搀䘀攀挀栀愀䘀椀渀䴀攀猀⤀⤀ ഀ
		and month(dFecha)=month(dateadd(month,-1,@dFechaFinMes)) ਍ ഀ
 ਍ऀ猀攀氀攀挀琀 䌀倀⸀挀䌀琀愀䌀漀搀Ⰰ挀䄀最攀㴀猀甀戀猀琀爀椀渀最⠀䌀倀⸀挀䌀琀愀䌀漀搀Ⰰ㐀Ⰰ㈀⤀Ⰰ䐀攀猀琀椀渀漀㴀椀猀渀甀氀氀⠀䌀⸀渀䐀攀猀琀䌀爀攀Ⰰ✀✀⤀Ⰰ 䌀漀搀䌀氀椀攀渀琀攀㴀倀攀爀⸀挀倀攀爀猀䌀漀搀 Ⰰ  ഀ
		Cliente = Per.cPersNombre, cCodDoc = PD.cPersIDNro,nMontoApr= C.nMontoApr, ਍ऀऀ挀䔀猀琀愀搀漀㴀 䔀⸀挀䌀漀渀猀䐀攀猀挀爀椀瀀挀椀漀渀Ⰰ渀䌀甀漀琀愀猀 㴀 䌀⸀渀䌀甀漀琀愀猀䄀瀀爀Ⰰ渀䐀椀愀䘀椀樀漀㴀椀猀渀甀氀氀⠀渀䐀椀愀䘀椀樀漀Ⰰ　⤀Ⰰ ഀ
		cAnalista= isnull(C.cCodAnalista,''),nTasa= C.nTasaInt,cLineaCredito=CL.cDescripcion, ਍ऀऀ搀䘀攀挀嘀椀最㴀 䌀⸀搀䘀攀挀嘀椀最Ⰰ渀匀愀氀搀漀䌀愀瀀㴀䌀⸀渀匀愀氀搀漀䌀愀瀀Ⰰ渀䌀甀漀琀愀䄀挀琀甀愀氀㴀䌀⸀渀一爀漀倀爀漀砀䌀甀漀琀愀Ⰰ ഀ
		cTipoPer= TP.cConsDescripcion, cPersCIIU=isnull(Per.cPersCIIU,''),Per.cPersDireccDomicilio, ਍ऀऀ挀䌀愀氀椀昀䄀渀琀攀爀椀漀爀㴀挀愀猀攀 眀栀攀渀 䌀吀⸀挀䌀愀氀䜀攀渀㴀✀　✀ 琀栀攀渀 ✀　 一伀刀䴀䄀䰀✀  ഀ
						when CT.cCalGen='1' then '1 CPP' ਍ऀऀऀऀऀऀ眀栀攀渀 䌀吀⸀挀䌀愀氀䜀攀渀㴀✀㈀✀ 琀栀攀渀 ✀㈀ 䐀䔀䘀䤀䌀䤀䔀一吀䔀✀ ഀ
						when CT.cCalGen='3' then '3 DUDOSO' ਍ऀऀऀऀऀऀ眀栀攀渀 䌀吀⸀挀䌀愀氀䜀攀渀㴀✀㐀✀ 琀栀攀渀 ✀㐀 倀䔀刀䐀䤀䐀䄀✀ 攀渀搀Ⰰ ഀ
		cCalifActual= case when CP.cCalGen='0' then '0 NORMAL'  ਍ऀऀऀऀऀऀ眀栀攀渀 䌀倀⸀挀䌀愀氀䜀攀渀㴀✀㄀✀ 琀栀攀渀 ✀㄀ 䌀倀倀✀ ഀ
						when CP.cCalGen='2' then '2 DEFICIENTE' ਍ऀऀऀऀऀऀ眀栀攀渀 䌀倀⸀挀䌀愀氀䜀攀渀㴀✀㌀✀ 琀栀攀渀 ✀㌀ 䐀唀䐀伀匀伀✀ ഀ
						when CP.cCalGen='4' then '4 PERDIDA' end, ਍ऀऀ渀䐀椀愀猀䄀琀爀愀猀漀㴀䌀⸀渀䐀椀愀猀䄀琀爀愀猀漀Ⰰ 挀䘀甀攀渀琀攀㴀ऀ✀ 䌀吀䄀㨀✀⬀ 䌀䰀䔀⸀挀䌀琀愀䌀漀渀琀 ⬀ 䘀⸀挀䐀攀猀挀爀椀瀀挀椀漀渀Ⰰ ഀ
		cPlazo=case when substring(CL.cLineaCred,6,1)='1' then 'CP'else 'LP'end, ਍ऀऀ挀吀椀瀀漀倀爀漀搀㴀挀愀猀攀 眀栀攀渀 猀甀戀猀琀爀椀渀最⠀䌀⸀挀䌀琀愀䌀漀搀Ⰰ㘀Ⰰ㌀⤀㴀✀㌀　㔀✀ 琀栀攀渀 ✀倀刀䔀一䐀䄀刀䤀伀✀  ഀ
						when substring(C.cCtaCod,6,1)='1' then 'COMERCIAL'  ਍ऀऀऀऀऀऀ眀栀攀渀 猀甀戀猀琀爀椀渀最⠀䌀⸀挀䌀琀愀䌀漀搀Ⰰ㘀Ⰰ㄀⤀㴀✀㈀✀ 琀栀攀渀 ✀䴀䤀䌀刀伀䔀䴀倀刀䔀匀䄀✀ ഀ
						when substring(C.cCtaCod,6,1)='3' then 'ConSUMO' ਍ऀऀऀऀऀऀ眀栀攀渀 猀甀戀猀琀爀椀渀最⠀䌀⸀挀䌀琀愀䌀漀搀Ⰰ㘀Ⰰ㄀⤀㴀✀㐀✀ 琀栀攀渀 ✀䠀䤀倀伀吀䔀䌀䄀刀䤀伀✀ 攀渀搀Ⰰ ഀ
		cMoneda=case when substring(C.cCtaCod,9,1)='1'then 'MN'else'ME'end, ਍ऀऀ搀䘀攀挀嘀攀渀挀㴀䌀⸀搀䘀攀挀嘀攀渀挀Ⰰ 䌀⸀渀䤀渀琀䐀攀瘀Ⰰ䌀⸀渀䤀渀琀匀甀猀瀀Ⰰ ഀ
		PorProvision=case isnull(CP.cCalGen,0)  ਍ऀऀऀऀऀऀऀ眀栀攀渀 　 琀栀攀渀  ഀ
									case isnull(CP.nGarant,0)  ਍ऀऀऀऀऀऀऀऀऀऀ眀栀攀渀 ✀㄀✀ 琀栀攀渀 椀猀渀甀氀氀⠀吀䄀䈀⸀渀倀爀漀瘀椀猀椀漀渀䄀Ⰰ㄀⤀  ഀ
										else isnull(TAB.nProvision,1)  ਍ऀऀऀऀऀऀऀऀऀ攀渀搀 ഀ
							else isnull(TAB.nProvision,1) ਍ऀऀऀऀऀऀ攀渀搀Ⰰ ഀ
		---				case isnull(CP.cCalGen,0) when '0' then isnull(TAB.nProvisionA,1) else isnull(TAB.nProvision,1) end ਍ ഀ
		nProvisionConRCC=isnull(CP.nProvisionRCC,0),nProvisionSINRCC=CP.nProvision, ਍ऀऀ渀倀爀漀瘀椀猀椀漀渀䄀渀琀䌀刀䌀䌀㴀 椀猀渀甀氀氀⠀䌀吀⸀渀倀爀漀瘀椀猀椀漀渀Ⰰ　⤀Ⰰ ഀ
		nProvisionAntSRCC= isnull(CT.nProvisionRCC,0), ਍ऀऀ渀匀愀氀搀漀䐀攀甀搀漀爀 㴀 ⠀猀攀氀攀挀琀 匀唀䴀⠀挀愀猀攀 眀栀攀渀 猀甀戀猀琀爀椀渀最⠀䌀刀䔀⸀挀䌀琀愀䌀漀搀Ⰰ㤀Ⰰ㄀⤀㴀✀㄀✀琀栀攀渀 渀匀愀氀搀漀䌀愀瀀 攀氀猀攀 渀匀愀氀搀漀䌀愀瀀 ⨀ 䀀渀吀椀瀀䌀愀洀戀 攀渀搀⤀ ഀ
						from DBConsolidada..CreditoConsol CRE inner join DBConsolidada..ProductoPersonaConsol PPC  ਍ऀऀऀऀऀऀऀ漀渀 䌀刀䔀⸀挀䌀琀愀䌀漀搀 㴀 倀倀䌀⸀挀䌀琀愀䌀漀搀 䄀一䐀 倀倀䌀⸀渀倀爀搀倀攀爀猀刀攀氀愀挀㴀 ㈀　 ഀ
						where PPC.cPersCod=Per.cPersCod), ਍ऀऀ渀䜀爀甀瀀漀倀爀攀昀 㴀 挀愀猀攀 眀栀攀渀 椀猀渀甀氀氀⠀䌀倀⸀渀䜀愀爀倀爀攀昀Ⰰ　⤀ 㸀　 琀栀攀渀 䌀倀⸀渀䜀愀爀倀爀攀昀 攀氀猀攀 　 攀渀搀Ⰰ ഀ
		nGruponoPref = case when isnull(CP.nGarNOPref,0) >0 then CP.nGarNOPref else 0 end, ਍ऀऀ渀䜀爀甀瀀漀䄀甀琀漀䰀 㴀 挀愀猀攀 眀栀攀渀 椀猀渀甀氀氀⠀䌀倀⸀渀䜀愀爀䄀甀琀漀䰀Ⰰ　⤀ 㸀　 琀栀攀渀 䌀倀⸀渀䜀愀爀䄀甀琀漀䰀 攀氀猀攀 　 攀渀搀Ⰰ ഀ
		cTipoGarantiaCalif=CP.nGarant, ਍ऀऀ䄀氀椀渀攀愀搀漀㴀挀愀猀攀 眀栀攀渀 䌀倀⸀挀䌀愀氀匀椀猀琀䘀 椀猀 渀甀氀氀 琀栀攀渀 ✀一伀✀ 攀氀猀攀 ✀匀䤀✀ 攀渀搀Ⰰ ഀ
		CO.cConsDescripcion as cCondicion, ਍ऀऀ挀䌀愀氀椀昀匀椀渀䄀氀椀渀攀愀㴀 挀愀猀攀 眀栀攀渀 䌀倀⸀挀䌀愀氀匀椀渀䄀氀椀㴀✀　✀ 琀栀攀渀 ✀　 一伀刀䴀䄀䰀✀  ഀ
						when CP.cCalSinAli='1' then '1 CPP' ਍ऀऀऀऀऀऀ眀栀攀渀 䌀倀⸀挀䌀愀氀匀椀渀䄀氀椀㴀✀㈀✀ 琀栀攀渀 ✀㈀ 䐀䔀䘀䤀䌀䤀䔀一吀䔀✀ ഀ
						when CP.cCalSinAli='3' then '3 DUDOSO' ਍ऀऀऀऀऀऀ眀栀攀渀 䌀倀⸀挀䌀愀氀匀椀渀䄀氀椀㴀✀㐀✀ 琀栀攀渀 ✀㐀 倀䔀刀䐀䤀䐀䄀✀ 攀渀搀Ⰰ ഀ
		cCalifSistF= case when CP.cCalSistF='0' then '0 NORMAL'  ਍ऀऀऀऀऀऀ眀栀攀渀 䌀倀⸀挀䌀愀氀匀椀猀琀䘀㴀✀㄀✀ 琀栀攀渀 ✀㄀ 䌀倀倀✀ ഀ
						when CP.cCalSistF='2' then '2 DEFICIENTE' ਍ऀऀऀऀऀऀ眀栀攀渀 䌀倀⸀挀䌀愀氀匀椀猀琀䘀㴀✀㌀✀ 琀栀攀渀 ✀㌀ 䐀唀䐀伀匀伀✀ ഀ
						when CP.cCalSistF='4' then '4 PERDIDA' end, ਍ऀऀ䌀倀⸀渀倀爀漀瘀椀猀椀漀渀䌀愀氀匀椀渀䄀氀椀Ⰰ ഀ
		CP.nProvisionCalSistF, ਍ऀऀ挀䌀氀椀攀渀琀攀唀渀椀挀漀㴀⠀挀愀猀攀 眀栀攀渀 ⠀猀攀氀攀挀琀 挀漀甀渀琀⠀搀椀猀琀椀渀挀琀 刀䐀⸀䌀漀搀开䔀洀瀀⤀ 愀猀 一爀漀䤀渀猀琀  ഀ
								from DBConsolidada..RCCTotal RC inner join DBConsolidada..RCCTotalDet RD on RC.Cod_Edu=RD.Cod_Edu ਍ऀऀऀऀऀऀऀऀ眀栀攀爀攀 爀椀最栀琀⠀刀䐀⸀䌀漀搀开䔀洀瀀Ⰰ㌀⤀ 㰀㸀✀㄀　㤀✀ ⴀⴀ䌀愀樀愀 䴀愀礀渀愀猀 ഀ
									and (RC.Cod_Doc_Trib=PD.cPersIDNro or RC.Cod_Doc_Id=PD.cPersIDNro) ਍ऀऀऀऀऀऀऀऀ⤀㴀　 琀栀攀渀 ✀匀䤀✀  ഀ
							else 'NO' end), ਍ऀऀ倀攀爀⸀挀倀攀爀猀䌀漀搀匀戀猀 愀猀 䌀漀搀开匀䈀匀Ⰰ ഀ
		--By Capi 30062008 se agrego campos segun Acta 131-2008 ਍ऀऀऀ䌀⸀渀䌀甀漀琀愀䄀瀀爀 䌀甀漀琀愀Ⰰ䌀⸀渀䌀愀瀀嘀攀渀挀椀搀漀 䌀愀瀀椀琀愀氀开嘀攀渀挀椀搀漀Ⰰ ഀ
			--Ahora el campo segun condiciones descritas en la misma acta ਍ऀऀऀ嘀攀渀挀椀搀漀㴀ऀ䌀愀猀攀 䰀攀昀琀⠀䌀⸀挀䌀琀愀䌀漀搀Ⰰ㈀⤀ ഀ
							When	'10' --Comercial ਍ऀऀऀऀऀऀऀऀ吀栀攀渀ऀ䌀愀猀攀  ഀ
											When C.nDiasAtraso>15 Then C.nCuotaApr ਍ऀऀऀऀऀऀऀऀऀऀऀ䔀氀猀攀ऀ　 ഀ
										End ਍ऀऀऀऀऀऀऀ圀栀攀渀ऀ✀㈀　✀ ⴀⴀ䴀攀猀 ഀ
								Then	Case  ਍ऀऀऀऀऀऀऀऀऀऀऀ圀栀攀渀 䌀⸀渀䐀椀愀猀䄀琀爀愀猀漀㸀㌀　 吀栀攀渀 䌀⸀渀䌀愀瀀嘀攀渀挀椀搀漀 ഀ
											Else	0 ਍ऀऀऀऀऀऀऀऀऀऀ䔀渀搀 ഀ
							When	'30' --Consumo ਍ऀऀऀऀऀऀऀऀ吀栀攀渀ऀ䌀愀猀攀  ഀ
											When C.nDiasAtraso>90 Then C.nCapVencido ਍ऀऀऀऀऀऀऀऀऀऀऀ圀栀攀渀 䌀⸀渀䐀椀愀猀䄀琀爀愀猀漀㸀㌀　 䄀渀搀 䌀⸀渀䐀椀愀猀䄀琀爀愀猀漀㰀㴀㤀　  吀栀攀渀 䌀⸀渀䌀甀漀琀愀䄀瀀爀 ഀ
											Else	0 ਍ऀऀऀऀऀऀऀऀऀऀ䔀渀搀 ഀ
							When	'40' --Hipotecario ਍ऀऀऀऀऀऀऀऀ吀栀攀渀 ഀ
										Case  ਍ऀऀऀऀऀऀऀऀऀऀऀ圀栀攀渀 䌀⸀渀䐀椀愀猀䄀琀爀愀猀漀㸀㤀　 吀栀攀渀 䌀⸀渀䌀愀瀀嘀攀渀挀椀搀漀 ഀ
											When C.nDiasAtraso>30 And C.nDiasAtraso<=90  Then C.nCuotaApr ਍ऀऀऀऀऀऀऀऀऀऀऀ䔀氀猀攀ऀ　 ഀ
										End ਍ऀऀऀऀऀऀ䔀渀搀Ⰰ ഀ
			Institucion=	Case	When c.cCodInst<>''  ਍ऀऀऀऀऀऀऀऀ吀栀攀渀ऀ倀䌀㄀⸀挀倀攀爀猀一漀洀戀爀攀 ഀ
								Else '' ਍ऀऀऀऀऀऀऀ䔀渀搀Ⰰ ഀ
		-- ਍ऀऀⴀⴀ ഀ
		PorProvisionPro=case isnull(CP.cCalGen,0)  ਍ऀऀऀऀऀऀऀ眀栀攀渀 　 琀栀攀渀  ഀ
									case isnull(CP.nGarant,0)  ਍ऀऀऀऀऀऀऀऀऀऀ眀栀攀渀 ✀㄀✀ 琀栀攀渀 椀猀渀甀氀氀⠀吀䄀䈀⸀渀倀爀漀瘀椀猀椀漀渀倀爀漀挀椀挀氀椀挀愀䄀Ⰰ㄀⤀ऀ ഀ
										else isnull(TAB.nProvisionProciclica,1)  ਍ऀऀऀऀऀऀऀऀऀ攀渀搀 ഀ
							else 0 ਍ऀऀऀऀऀऀ攀渀搀Ⰰ ഀ
		isnull(CP.nProvisionProciclica,0) nProvisionProciclica, ਍ऀऀ渀倀爀漀瘀椀猀椀漀渀倀爀漀挀椀挀氀椀挀愀吀漀琀愀氀㴀挀愀猀攀 椀猀渀甀氀氀⠀䌀倀⸀挀䌀愀氀䜀攀渀Ⰰ　⤀ ഀ
								  when 0 then  ਍ⴀⴀ䄀䰀倀䄀 ㈀　　㤀　㌀　㐀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀ഀ
										case --case isnull(CP.nGarant,0) when '1' then isnull(TAB.nProvisionProciclicaA,1) else isnull(TAB.nProvisionProciclica,1) end ਍ⴀⴀऀऀऀऀऀऀऀऀऀऀऀ眀栀攀渀 　⸀㐀 琀栀攀渀 　⸀㐀㔀 ഀ
--											when 0.3 then 0.3 ਍ⴀⴀऀऀऀऀऀऀऀऀऀऀऀ眀栀攀渀 　⸀㌀㌀ 琀栀攀渀 　⸀㔀 ഀ
--											when 0.67 then 1 ਍ⴀⴀऀऀऀऀऀऀऀऀऀऀऀ眀栀攀渀 　⸀㌀㜀 琀栀攀渀 　⸀㐀 ഀ
਍ऀऀऀऀऀऀऀऀऀऀऀऀ眀栀攀渀 椀猀渀甀氀氀⠀䌀倀⸀渀䜀愀爀愀渀琀Ⰰ　⤀ 㴀 ✀㄀✀ 愀渀搀 猀甀戀猀琀爀椀渀最⠀䌀倀⸀挀䌀琀愀䌀漀搀Ⰰ㘀Ⰰ㄀⤀㴀✀㄀✀ 琀栀攀渀 椀猀渀甀氀氀⠀吀䄀䈀⸀渀倀爀漀瘀椀猀椀漀渀倀爀漀挀椀挀氀椀挀愀䄀Ⰰ　⤀ഀ
												when isnull(CP.nGarant,0) <> '1' and substring(CP.cCtaCod,6,1)='1' then isnull(TAB.nProvisionProciclica,0)਍ऀऀऀऀऀऀऀऀऀऀऀऀ眀栀攀渀 椀猀渀甀氀氀⠀䌀倀⸀渀䜀愀爀愀渀琀Ⰰ　⤀ 㴀 ✀㄀✀ 愀渀搀 猀甀戀猀琀爀椀渀最⠀䌀倀⸀挀䌀琀愀䌀漀搀Ⰰ㘀Ⰰ㄀⤀㴀✀㈀✀ 琀栀攀渀 椀猀渀甀氀氀⠀吀䄀䈀⸀渀倀爀漀瘀椀猀椀漀渀倀爀漀挀椀挀氀椀挀愀䄀Ⰰ　⤀ഀ
												when isnull(CP.nGarant,0) <> '1' and substring(CP.cCtaCod,6,1)='2' then isnull(TAB.nProvisionProciclica,0)਍ऀऀऀऀऀऀऀऀऀऀऀऀ眀栀攀渀 椀猀渀甀氀氀⠀䌀倀⸀渀䜀愀爀愀渀琀Ⰰ　⤀ 㴀 ✀㄀✀ 愀渀搀 猀甀戀猀琀爀椀渀最⠀䌀倀⸀挀䌀琀愀䌀漀搀Ⰰ㘀Ⰰ㄀⤀㴀✀㌀✀ 琀栀攀渀 椀猀渀甀氀氀⠀吀䄀䈀⸀渀倀爀漀瘀椀猀椀漀渀倀爀漀挀椀挀氀椀挀愀䄀Ⰰ　⤀ഀ
												when isnull(CP.nGarant,0) <> '1' and substring(CP.cCtaCod,6,1)='3' then isnull(TAB.nProvisionProciclica,0)਍ऀऀऀऀऀऀऀऀऀऀऀऀ眀栀攀渀 椀猀渀甀氀氀⠀䌀倀⸀渀䜀愀爀愀渀琀Ⰰ　⤀ 㴀 ✀㄀✀ 愀渀搀 猀甀戀猀琀爀椀渀最⠀䌀倀⸀挀䌀琀愀䌀漀搀Ⰰ㘀Ⰰ㄀⤀㴀✀㐀✀ 琀栀攀渀 椀猀渀甀氀氀⠀吀䄀䈀⸀渀倀爀漀瘀椀猀椀漀渀倀爀漀挀椀挀氀椀挀愀䄀Ⰰ　⤀ഀ
												when isnull(CP.nGarant,0) <> '1' and substring(CP.cCtaCod,6,1)='4' then isnull(TAB.nProvisionProciclica,0)਍ऀऀऀऀऀऀऀऀऀऀऀऀ攀氀猀攀 　 ഀ
										end ਍ऀऀऀऀऀऀऀ      攀氀猀攀 　 ഀ
								  end * C.nSaldoCap/100 ਍ⴀⴀ⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀ഀ
	from ColocCalifProv CP  ਍ऀऀ椀渀渀攀爀 樀漀椀渀 䐀䈀䌀漀渀猀漀氀椀搀愀搀愀⸀⸀䌀爀攀搀椀琀漀䌀漀渀猀漀氀 䌀 漀渀 䌀倀⸀挀䌀琀愀䌀漀搀 㴀 䌀⸀挀䌀琀愀䌀漀搀 ഀ
		inner join DBConsolidada..ProductoPersonaConsol PP on C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = 20 ਍ऀऀ椀渀渀攀爀 樀漀椀渀 倀攀爀猀漀渀愀 倀攀爀 漀渀 倀攀爀⸀挀倀攀爀猀䌀漀搀 㴀 倀倀⸀挀倀攀爀猀䌀漀搀 ഀ
		left join PersID PD on PD.cPersCod = Per.cPersCod  ਍ऀऀऀ愀渀搀 挀倀攀爀猀䤀䐀吀瀀漀㴀⠀猀攀氀攀挀琀 䴀䤀一⠀挀倀攀爀猀䤀䐀吀瀀漀⤀昀爀漀洀 倀攀爀猀䤀䐀 眀栀攀爀攀 挀倀攀爀猀䌀漀搀㴀倀攀爀⸀挀倀攀爀猀䌀漀搀 愀渀搀 挀倀攀爀猀䤀搀吀瀀漀㸀㴀㄀⤀ ഀ
		inner join Constante E on E.nConsValor = C.nPrdEstado AND E.nConsCod = 3001 ਍ऀऀ椀渀渀攀爀 樀漀椀渀 䌀漀氀漀挀䰀椀渀攀愀䌀爀攀搀椀琀漀 䌀䰀 漀渀 䌀䰀⸀挀䰀椀渀攀愀䌀爀攀搀 㴀 䌀⸀挀䰀椀渀攀愀䌀爀攀搀 ഀ
		inner join Constante TP on TP.nConsValor = Per.nPersPersoneria AND TP.nConsCod = 1002 ਍ऀऀ椀渀渀攀爀 樀漀椀渀 䌀漀氀漀挀䰀椀渀攀愀䌀爀攀搀椀琀漀 䘀 漀渀 䘀⸀挀䰀椀渀攀愀䌀爀攀搀 㴀 猀甀戀猀琀爀椀渀最⠀䌀䰀⸀挀䰀椀渀攀愀䌀爀攀搀Ⰰ㄀Ⰰ㐀⤀ 䄀一䐀 䰀䔀一⠀䘀⸀挀䰀椀渀攀愀䌀爀攀搀⤀㴀㐀 ഀ
		inner join DBConsolidada..ColocLineaCreditoEquiv CLE on F.cLineaCred = CLE.cLineaCred ਍ऀऀ氀攀昀琀 樀漀椀渀 䌀漀氀漀挀䌀愀氀椀昀椀挀愀吀愀戀氀愀 吀䄀䈀 漀渀 䰀䔀䘀吀⠀渀䌀愀氀䌀漀搀吀愀戀Ⰰ㄀⤀㴀 猀甀戀猀琀爀椀渀最⠀䌀倀⸀挀䌀琀愀䌀漀搀Ⰰ㘀Ⰰ㄀⤀  ഀ
			and TAB.cCalif= CP.cCalGen AND TAB.cRefinan= CP.cRefinan  ਍ऀऀऀ愀渀搀 猀甀戀猀琀爀椀渀最⠀挀漀渀瘀攀爀琀⠀瘀愀爀挀栀愀爀⠀㐀⤀Ⰰ渀䌀愀氀䌀漀搀吀愀戀⤀Ⰰ㈀Ⰰ㄀⤀㴀 ⠀挀愀猀攀 眀栀攀渀 䌀倀⸀渀䜀愀爀愀渀琀㴀㐀 琀栀攀渀 ✀　✀ 攀氀猀攀 ✀㄀✀ 攀渀搀⤀ ഀ
		left join  DBConsolidada..ColocCalifProvTotal CT on CT.cCtaCod = CP.cCtaCod and datediff(day,CT.dFecha,@dFechaMesAnt)=0 --mes anterior ਍ऀऀ氀攀昀琀 樀漀椀渀 䌀漀渀猀琀愀渀琀攀 䌀伀 漀渀 䌀⸀渀䌀漀渀搀䌀爀攀㴀䌀伀⸀渀䌀漀渀猀嘀愀氀漀爀 愀渀搀 䌀伀⸀渀䌀漀渀猀䌀漀搀㴀㌀　㄀㔀 ഀ
		--By Capi 30062008 para jalar institucion Convenio ਍ऀऀ氀攀昀琀 樀漀椀渀 倀攀爀猀漀渀愀 倀䌀㄀ 伀渀 倀䌀㄀⸀挀倀攀爀猀䌀漀搀㴀䌀⸀挀䌀漀搀䤀渀猀琀 ഀ
		 ਍ऀ眀栀攀爀攀 䌀⸀渀倀爀搀䔀猀琀愀搀漀 䤀一 ⠀㈀　㈀　Ⰰ ㈀　㈀㄀Ⰰ ㈀　㈀㈀Ⰰ ㈀　㌀　Ⰰ ㈀　㌀㄀Ⰰ ㈀　㌀㈀ Ⰰ ㈀㈀　㄀Ⰰ㈀㈀　㔀Ⰰ ㈀㄀　㄀Ⰰ㈀㄀　㐀Ⰰ ㈀㄀　㘀Ⰰ ㈀㄀　㜀⤀ ഀ
	union all ਍ऀ猀攀氀攀挀琀 䌀倀⸀挀䌀琀愀䌀漀搀Ⰰ挀䄀最攀㴀猀甀戀猀琀爀椀渀最⠀䌀倀⸀挀䌀琀愀䌀漀搀Ⰰ㐀Ⰰ㈀⤀Ⰰ䐀攀猀琀椀渀漀㴀　Ⰰ 䌀漀搀䌀氀椀攀渀琀攀㴀倀攀爀⸀挀倀攀爀猀䌀漀搀 Ⰰ  ഀ
		Cliente = Per.cPersNombre, cCodDoc = PD.cPersIDNro,nMontoApr= C.nMontoApr, ਍ऀऀ挀䔀猀琀愀搀漀㴀 䔀⸀挀䌀漀渀猀䐀攀猀挀爀椀瀀挀椀漀渀Ⰰ渀䌀甀漀琀愀猀 㴀 　Ⰰ渀䐀椀愀䘀椀樀漀㴀　Ⰰ ഀ
		cAnalista= isnull(C.cCodAnalista,''),nTasa= 0.00,cLineaCredito='', ਍ऀऀ搀䘀攀挀嘀椀最㴀 䌀⸀搀䘀攀挀嘀椀最Ⰰ渀匀愀氀搀漀䌀愀瀀㴀 䌀⸀渀䴀漀渀琀漀䄀瀀爀Ⰰ渀䌀甀漀琀愀䄀挀琀甀愀氀㴀 　Ⰰ ഀ
		cTipoPer= TP.cConsDescripcion, cPersCIIU=isnull(Per.cPersCIIU,''),cPersDireccDomicilio, ਍ऀऀ挀䌀愀氀椀昀䄀渀琀攀爀椀漀爀㴀挀愀猀攀 眀栀攀渀 䌀吀⸀挀䌀愀氀䜀攀渀㴀✀　✀ 琀栀攀渀 ✀　 一伀刀䴀䄀䰀✀  ഀ
						when CT.cCalGen='1' then '1 CPP' ਍ऀऀऀऀऀऀ眀栀攀渀 䌀吀⸀挀䌀愀氀䜀攀渀㴀✀㈀✀ 琀栀攀渀 ✀㈀ 䐀䔀䘀䤀䌀䤀䔀一吀䔀✀ ഀ
						when CT.cCalGen='3' then '3 DUDOSO' ਍ऀऀऀऀऀऀ眀栀攀渀 䌀吀⸀挀䌀愀氀䜀攀渀㴀✀㐀✀ 琀栀攀渀 ✀㐀 倀䔀刀䐀䤀䐀䄀✀ 攀渀搀Ⰰ ഀ
		cCalifActual= case when CP.cCalGen='0' then '0 NORMAL'  ਍ऀऀऀऀऀऀ眀栀攀渀 䌀倀⸀挀䌀愀氀䜀攀渀㴀✀㄀✀ 琀栀攀渀 ✀㄀ 䌀倀倀✀ ഀ
						when CP.cCalGen='2' then '2 DEFICIENTE' ਍ऀऀऀऀऀऀ眀栀攀渀 䌀倀⸀挀䌀愀氀䜀攀渀㴀✀㌀✀ 琀栀攀渀 ✀㌀ 䐀唀䐀伀匀伀✀ ഀ
						when CP.cCalGen='4' then '4 PERDIDA' end, ਍ऀऀ渀䐀椀愀猀䄀琀爀愀猀漀㴀 　Ⰰ 挀䘀甀攀渀琀攀㴀ऀ✀✀Ⰰ ഀ
		cPlazo='', ਍ऀऀ挀吀椀瀀漀倀爀漀搀㴀✀䌀䄀刀吀䄀 䘀䤀䄀一娀䄀✀Ⰰ ഀ
		cMoneda=case when substring(C.cCtaCod,9,1)='1'then 'MN'else'ME'end, ਍ऀऀ搀䘀攀挀嘀攀渀挀㴀䌀⸀搀嘀攀渀挀䄀瀀爀Ⰰ 渀䤀渀琀䐀攀瘀㴀　Ⰰ渀䤀渀琀匀甀猀瀀㴀　Ⰰ ഀ
		PorProvision=case isnull(CP.cCalGen,0)  ਍ऀऀऀऀऀऀऀ眀栀攀渀 　 琀栀攀渀  ഀ
									case isnull(CP.nGarant,0)  ਍ऀऀऀऀऀऀऀऀऀऀ眀栀攀渀 ✀㄀✀ 琀栀攀渀 椀猀渀甀氀氀⠀吀䄀䈀⸀渀倀爀漀瘀椀猀椀漀渀䄀Ⰰ㄀⤀  ഀ
										else isnull(TAB.nProvision,1)  ਍ऀऀऀऀऀऀऀऀऀ攀渀搀 ഀ
							else isnull(TAB.nProvision,1) ਍ऀऀऀऀऀऀ攀渀搀Ⰰ ഀ
		--case isnull(CP.cCalGen,0) when '1' then isnull(TAB.nProvisionA,1) else isnull(TAB.nProvision,1) end, ਍ऀऀ渀倀爀漀瘀椀猀椀漀渀䌀漀渀刀䌀䌀㴀椀猀渀甀氀氀⠀䌀倀⸀渀倀爀漀瘀椀猀椀漀渀刀䌀䌀Ⰰ　⤀Ⰰ渀倀爀漀瘀椀猀椀漀渀匀䤀一刀䌀䌀㴀䌀倀⸀渀倀爀漀瘀椀猀椀漀渀Ⰰ ഀ
		nProvisionAntCRCC= isnull(CT.nProvision,0), ਍ऀऀ渀倀爀漀瘀椀猀椀漀渀䄀渀琀匀刀䌀䌀㴀 椀猀渀甀氀氀⠀䌀吀⸀渀倀爀漀瘀椀猀椀漀渀刀䌀䌀Ⰰ　⤀Ⰰ ഀ
		nSaldoDeudor = (select SUM(case when substring(CRE.cCtaCod,9,1)='1'then nSaldoCap else nSaldoCap * @nTipCamb end) ਍ऀऀऀऀऀऀ昀爀漀洀 䐀䈀䌀漀渀猀漀氀椀搀愀搀愀⸀⸀䌀愀爀琀愀䘀椀愀渀稀愀匀愀氀搀漀䌀漀渀猀漀氀 䌀刀䔀 椀渀渀攀爀 樀漀椀渀 䐀䈀䌀漀渀猀漀氀椀搀愀搀愀⸀⸀倀爀漀搀甀挀琀漀倀攀爀猀漀渀愀䌀漀渀猀漀氀 倀倀䌀  ഀ
							on CRE.cCtaCod = PPC.cCtaCod AND PPC.nPrdPersRelac= 20 ਍ऀऀऀऀऀऀ眀栀攀爀攀 倀倀䌀⸀挀倀攀爀猀䌀漀搀㴀倀攀爀⸀挀倀攀爀猀䌀漀搀⤀Ⰰ ഀ
		nGrupoPref =  C.nMontoApr, ਍ऀऀ渀䜀爀甀瀀漀渀漀倀爀攀昀 㴀 　Ⰰ ഀ
		nGrupoAutoL = 0, ਍ऀऀ挀吀椀瀀漀䜀愀爀愀渀琀椀愀䌀愀氀椀昀㴀䌀倀⸀渀䜀愀爀愀渀琀Ⰰ ഀ
		Alineado=case when CP.cCalSistF IS NULL then 'NO' else 'SI' end, ਍ऀऀ✀✀ 愀猀 挀䌀漀渀搀椀挀椀漀渀Ⰰ ഀ
		cCalifSinAlinea= case when CP.cCalSinAli='0' then '0 NORMAL'  ਍ऀऀऀऀऀऀ眀栀攀渀 䌀倀⸀挀䌀愀氀匀椀渀䄀氀椀㴀✀㄀✀ 琀栀攀渀 ✀㄀ 䌀倀倀✀ ഀ
						when CP.cCalSinAli='2' then '2 DEFICIENTE' ਍ऀऀऀऀऀऀ眀栀攀渀 䌀倀⸀挀䌀愀氀匀椀渀䄀氀椀㴀✀㌀✀ 琀栀攀渀 ✀㌀ 䐀唀䐀伀匀伀✀ ഀ
						when CP.cCalSinAli='4' then '4 PERDIDA' end, ਍ऀऀ挀䌀愀氀椀昀匀椀猀琀䘀㴀 挀愀猀攀 眀栀攀渀 䌀倀⸀挀䌀愀氀匀椀猀琀䘀㴀✀　✀ 琀栀攀渀 ✀　 一伀刀䴀䄀䰀✀  ഀ
						when CP.cCalSistF='1' then '1 CPP' ਍ऀऀऀऀऀऀ眀栀攀渀 䌀倀⸀挀䌀愀氀匀椀猀琀䘀㴀✀㈀✀ 琀栀攀渀 ✀㈀ 䐀䔀䘀䤀䌀䤀䔀一吀䔀✀ ഀ
						when CP.cCalSistF='3' then '3 DUDOSO' ਍ऀऀऀऀऀऀ眀栀攀渀 䌀倀⸀挀䌀愀氀匀椀猀琀䘀㴀✀㐀✀ 琀栀攀渀 ✀㐀 倀䔀刀䐀䤀䐀䄀✀ 攀渀搀Ⰰ ഀ
		CP.nProvisionCalSinAli, ਍ऀऀ䌀倀⸀渀倀爀漀瘀椀猀椀漀渀䌀愀氀匀椀猀琀䘀Ⰰ ഀ
		cClienteUnico=(case when (select count(distinct RD.Cod_Emp) as NroInst  ਍ऀऀऀऀऀऀऀऀ昀爀漀洀 䐀䈀䌀漀渀猀漀氀椀搀愀搀愀⸀⸀刀䌀䌀吀漀琀愀氀 刀䌀 椀渀渀攀爀 樀漀椀渀 䐀䈀䌀漀渀猀漀氀椀搀愀搀愀⸀⸀刀䌀䌀吀漀琀愀氀䐀攀琀 刀䐀 漀渀 刀䌀⸀䌀漀搀开䔀搀甀㴀刀䐀⸀䌀漀搀开䔀搀甀 ഀ
								where right(RD.Cod_Emp,3) <>'109' --Caja Maynas ਍ऀऀऀऀऀऀऀऀऀ愀渀搀 ⠀刀䌀⸀䌀漀搀开䐀漀挀开吀爀椀戀㴀倀䐀⸀挀倀攀爀猀䤀䐀一爀漀 漀爀 刀䌀⸀䌀漀搀开䐀漀挀开䤀搀㴀倀䐀⸀挀倀攀爀猀䤀䐀一爀漀⤀ ഀ
								)=0 then 'SI'  ਍ऀऀऀऀऀऀऀ攀氀猀攀 ✀一伀✀ 攀渀搀⤀Ⰰ ഀ
		Per.cPersCodSbs as Cod_SBS, ਍ऀऀⴀⴀ䈀礀 䌀愀瀀椀 ㌀　　㘀㈀　　㠀 猀攀 愀最爀攀最漀 挀愀洀瀀漀猀 猀攀最甀渀 䄀挀琀愀 ㄀㌀㄀ⴀ㈀　　㠀 ഀ
		Cuota=0,Capital_Vencido=0,Vencido=	0,Institucion='',		 ਍ऀऀⴀⴀ倀漀爀倀爀漀瘀椀猀椀漀渀倀爀漀㴀挀愀猀攀 椀猀渀甀氀氀⠀䌀倀⸀挀䌀愀氀䜀攀渀Ⰰ　⤀ 眀栀攀渀 ✀㄀✀ 琀栀攀渀 椀猀渀甀氀氀⠀吀䄀䈀⸀渀倀爀漀瘀椀猀椀漀渀倀爀漀挀椀挀氀椀挀愀䄀Ⰰ㄀⤀ 攀氀猀攀 椀猀渀甀氀氀⠀吀䄀䈀⸀渀倀爀漀瘀椀猀椀漀渀倀爀漀挀椀挀氀椀挀愀Ⰰ㄀⤀ 攀渀搀Ⰰ ഀ
		PorProvisionPro=case isnull(CP.cCalGen,0)  ਍ऀऀऀऀऀऀऀ眀栀攀渀 　 琀栀攀渀  ഀ
									case isnull(CP.nGarant,0)  ਍ऀऀऀऀऀऀऀऀऀऀ眀栀攀渀 ✀㄀✀ 琀栀攀渀 椀猀渀甀氀氀⠀吀䄀䈀⸀渀倀爀漀瘀椀猀椀漀渀倀爀漀挀椀挀氀椀挀愀䄀Ⰰ㄀⤀ऀ ഀ
										else isnull(TAB.nProvisionProciclica,1)  ਍ऀऀऀऀऀऀऀऀऀ攀渀搀  ഀ
							else 0							 ਍ऀऀऀऀऀऀ攀渀搀Ⰰ ഀ
		isnull(CP.nProvisionProciclica,0) nProvisionProciclica, ਍ऀऀ渀倀爀漀瘀椀猀椀漀渀倀爀漀挀椀挀氀椀挀愀吀漀琀愀氀㴀挀愀猀攀 椀猀渀甀氀氀⠀䌀倀⸀挀䌀愀氀䜀攀渀Ⰰ　⤀ ഀ
								  when 0 then  ਍ⴀⴀ䄀䰀倀䄀 ㈀　　㤀　㌀　㐀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀⨀ഀ
										case -- case isnull(CP.nGarant,0)  when '1' then isnull(TAB.nProvisionProciclicaA,1) else isnull(TAB.nProvisionProciclica,1) end ਍ⴀⴀऀऀऀऀऀऀऀऀऀऀऀ眀栀攀渀 　⸀㐀 琀栀攀渀 　⸀㐀㔀 ഀ
--											when 0.3 then 0.3 ਍ⴀⴀऀऀऀऀऀऀऀऀऀऀऀ眀栀攀渀 　⸀㌀㌀ 琀栀攀渀 　⸀㔀 ഀ
--											when 0.67 then 1 ਍ⴀⴀऀऀऀऀऀऀऀऀऀऀऀ眀栀攀渀 　⸀㌀㜀 琀栀攀渀 　⸀㐀 ഀ
											when isnull(CP.nGarant,0) = '1' and substring(CP.cCtaCod,6,1)='1' then isnull(TAB.nProvisionProciclicaA,0)਍ऀऀऀऀऀऀऀऀऀऀऀ眀栀攀渀 椀猀渀甀氀氀⠀䌀倀⸀渀䜀愀爀愀渀琀Ⰰ　⤀ 㰀㸀 ✀㄀✀ 愀渀搀 猀甀戀猀琀爀椀渀最⠀䌀倀⸀挀䌀琀愀䌀漀搀Ⰰ㘀Ⰰ㄀⤀㴀✀㄀✀ 琀栀攀渀 椀猀渀甀氀氀⠀吀䄀䈀⸀渀倀爀漀瘀椀猀椀漀渀倀爀漀挀椀挀氀椀挀愀Ⰰ　⤀ഀ
											when isnull(CP.nGarant,0) = '1' and substring(CP.cCtaCod,6,1)='2' then isnull(TAB.nProvisionProciclicaA,0)਍ऀऀऀऀऀऀऀऀऀऀऀ眀栀攀渀 椀猀渀甀氀氀⠀䌀倀⸀渀䜀愀爀愀渀琀Ⰰ　⤀ 㰀㸀 ✀㄀✀ 愀渀搀 猀甀戀猀琀爀椀渀最⠀䌀倀⸀挀䌀琀愀䌀漀搀Ⰰ㘀Ⰰ㄀⤀㴀✀㈀✀ 琀栀攀渀 椀猀渀甀氀氀⠀吀䄀䈀⸀渀倀爀漀瘀椀猀椀漀渀倀爀漀挀椀挀氀椀挀愀Ⰰ　⤀ഀ
											when isnull(CP.nGarant,0) = '1' and substring(CP.cCtaCod,6,1)='3' then isnull(TAB.nProvisionProciclicaA,0)਍ऀऀऀऀऀऀऀऀऀऀऀ眀栀攀渀 椀猀渀甀氀氀⠀䌀倀⸀渀䜀愀爀愀渀琀Ⰰ　⤀ 㰀㸀 ✀㄀✀ 愀渀搀 猀甀戀猀琀爀椀渀最⠀䌀倀⸀挀䌀琀愀䌀漀搀Ⰰ㘀Ⰰ㄀⤀㴀✀㌀✀ 琀栀攀渀 椀猀渀甀氀氀⠀吀䄀䈀⸀渀倀爀漀瘀椀猀椀漀渀倀爀漀挀椀挀氀椀挀愀Ⰰ　⤀ഀ
											when isnull(CP.nGarant,0) = '1' and substring(CP.cCtaCod,6,1)='4' then isnull(TAB.nProvisionProciclicaA,0)਍ऀऀऀऀऀऀऀऀऀऀऀ眀栀攀渀 椀猀渀甀氀氀⠀䌀倀⸀渀䜀愀爀愀渀琀Ⰰ　⤀ 㰀㸀 ✀㄀✀ 愀渀搀 猀甀戀猀琀爀椀渀最⠀䌀倀⸀挀䌀琀愀䌀漀搀Ⰰ㘀Ⰰ㄀⤀㴀✀㐀✀ 琀栀攀渀 椀猀渀甀氀氀⠀吀䄀䈀⸀渀倀爀漀瘀椀猀椀漀渀倀爀漀挀椀挀氀椀挀愀Ⰰ　⤀ഀ
											else 0 ਍ऀऀऀऀऀऀऀऀऀऀ攀渀搀 ഀ
							      else 0 ਍ऀऀऀऀऀऀऀऀ  攀渀搀 ⨀ 䌀⸀渀䴀漀渀琀漀䄀瀀爀⼀㄀　　 ഀ
-- ***************************************************਍ऀ昀爀漀洀 䌀漀氀漀挀䌀愀氀椀昀倀爀漀瘀 䌀倀  ഀ
		inner join DBConsolidada..CartaFianzaConsol  C on CP.cCtaCod = C.cCtaCod ਍ऀऀ氀攀昀琀 樀漀椀渀  䐀䈀䌀漀渀猀漀氀椀搀愀搀愀⸀⸀䌀愀爀琀愀䘀椀愀渀稀愀匀愀氀搀漀䌀漀渀猀漀氀 䌀䘀䌀 漀渀 䌀⸀挀䌀琀愀䌀漀搀 㴀 䌀䘀䌀⸀挀䌀琀愀䌀漀搀 䄀一䐀 搀愀琀攀搀椀昀昀⠀搀Ⰰ搀䘀攀挀栀愀Ⰰ䀀搀䘀攀挀栀愀䘀椀渀䴀攀猀⤀㴀　 ⴀⴀ洀攀猀 愀挀琀甀愀氀 ഀ
		inner join DBConsolidada..ProductoPersonaConsol PP on C.cCtaCod = PP.cCtaCod AND PP.nPrdPersRelac = 20 ਍ऀऀ椀渀渀攀爀 樀漀椀渀 倀攀爀猀漀渀愀 倀攀爀 漀渀 倀攀爀⸀挀倀攀爀猀䌀漀搀 㴀 倀倀⸀挀倀攀爀猀䌀漀搀 ഀ
		left join PersID PD on PD.cPersCod = Per.cPersCod  ਍ऀऀऀ愀渀搀 挀倀攀爀猀䤀䐀吀瀀漀㴀⠀猀攀氀攀挀琀 䴀䤀一⠀挀倀攀爀猀䤀䐀吀瀀漀⤀昀爀漀洀 倀攀爀猀䤀䐀 眀栀攀爀攀 挀倀攀爀猀䌀漀搀㴀倀攀爀⸀挀倀攀爀猀䌀漀搀 愀渀搀 挀倀攀爀猀䤀搀吀瀀漀㸀㴀㄀⤀ ഀ
		inner join Constante E on E.nConsValor = C.nPrdEstado AND E.nConsCod = 3001 ਍ऀऀ椀渀渀攀爀 樀漀椀渀 䌀漀渀猀琀愀渀琀攀 吀倀 漀渀 吀倀⸀渀䌀漀渀猀嘀愀氀漀爀 㴀 倀攀爀⸀渀倀攀爀猀倀攀爀猀漀渀攀爀椀愀 䄀一䐀 吀倀⸀渀䌀漀渀猀䌀漀搀 㴀 ㄀　　㈀ ഀ
		left join  ColocCalificaTabla TAB on LEFT(nCalCodTab,1)= substring(CP.cCtaCod,6,1)  ਍ऀऀऀ愀渀搀 吀䄀䈀⸀挀䌀愀氀椀昀㴀 䌀倀⸀挀䌀愀氀䜀攀渀 䄀一䐀 吀䄀䈀⸀挀刀攀昀椀渀愀渀㴀 䌀倀⸀挀刀攀昀椀渀愀渀  ഀ
			and substring(convert(varchar(4),nCalCodTab),2,1)= (case when CP.nGarant=4 then '0' else '1' end) ਍ऀऀ氀攀昀琀 樀漀椀渀  䐀䈀䌀漀渀猀漀氀椀搀愀搀愀⸀⸀䌀漀氀漀挀䌀愀氀椀昀倀爀漀瘀吀漀琀愀氀 䌀吀 漀渀 䌀吀⸀挀䌀琀愀䌀漀搀 㴀 䌀倀⸀挀䌀琀愀䌀漀搀  ഀ
		and datediff(day,CT.dFecha,@dFechaMesAnt)=0 -- ULTIMO DIA DEL MES ANTERIOR ਍ऀ眀栀攀爀攀 䌀⸀渀倀爀搀䔀猀琀愀搀漀 䤀一 ⠀㈀　㈀　Ⰰ ㈀　㈀㄀Ⰰ ㈀　㈀㈀Ⰰ㈀　㤀㈀⤀ ഀ
end