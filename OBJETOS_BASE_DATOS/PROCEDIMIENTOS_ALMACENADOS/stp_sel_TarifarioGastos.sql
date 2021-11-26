Create Procedure stp_sel_TarifarioGastos
(
@vProducto as varchar(50)
)
as
if @vProducto='-1' or @vProducto='0' 
	Select convert(varchar(25),nPrdConceptoCod) as nPrdConceptoCod, cDescripcion, nAplicado, nInicial ,nFinal ,nTpoValor,  nValor, nMoneda,nMontoMin,nMontoMax,cAplicaMonto,cFiltro,cAplicaProceso, nOperador, nOperPorc,  cOperMonto, nEdad, nEdadOper, nDiasApl,cGastoFijoVar,nOperDiasVenc, nDiasVenc,nMontoMensual,  bAplTipCamb, bAplValorDosTit,nValorDosTit,bAplNumConCer,bAplNumMeses  From ProductoConcepto  Where nColocCred = 1 AND nPrdConceptoCod <> 1299  AND convert(varchar(15),nPrdConceptoCod) like '12%' --Order by convert(varchar(25),nPrdConceptoCod) 
	union all
	Select convert(varchar(25),nPrdConceptoCod) as nPrdConceptoCod, cDescripcion, nAplicado, nInicial ,nFinal ,nTpoValor,  nValor ,nMoneda,nMontoMin,nMontoMax,cAplicaMonto,cFiltro,cAplicaProceso, nOperador, nOperPorc,  cOperMonto, nEdad, nEdadOper, nDiasApl,cGastoFijoVar,nOperDiasVenc, nDiasVenc,nMontoMensual,  bAplTipCamb, bAplValorDosTit,nValorDosTit,bAplNumConCer,bAplNumMeses  From ProductoConcepto  Where nColocCred = 2 AND nPrdConceptoCod <> 1299  AND convert(varchar(15),nPrdConceptoCod) like '32%' --Order by convert(varchar(25),nPrdConceptoCod) 
else
if @vProducto='1'
	Select convert(varchar(25),nPrdConceptoCod) as nPrdConceptoCod, cDescripcion, nAplicado, nInicial ,nFinal ,nTpoValor,  nValor, nMoneda,nMontoMin,nMontoMax,cAplicaMonto,cFiltro,cAplicaProceso, nOperador, nOperPorc,  cOperMonto, nEdad, nEdadOper, nDiasApl,cGastoFijoVar,nOperDiasVenc, nDiasVenc,nMontoMensual,  bAplTipCamb, bAplValorDosTit,nValorDosTit,bAplNumConCer,bAplNumMeses  From ProductoConcepto  Where nColocCred = 1 AND nPrdConceptoCod <> 1299  AND convert(varchar(15),nPrdConceptoCod) like '12%' Order by convert(varchar(25),nPrdConceptoCod) 
else if @vProducto='2'
	Select convert(varchar(25),nPrdConceptoCod) as nPrdConceptoCod, cDescripcion, nAplicado, nInicial ,nFinal ,nTpoValor,  nValor ,nMoneda,nMontoMin,nMontoMax,cAplicaMonto,cFiltro,cAplicaProceso, nOperador, nOperPorc,  cOperMonto, nEdad, nEdadOper, nDiasApl,cGastoFijoVar,nOperDiasVenc, nDiasVenc,nMontoMensual,  bAplTipCamb, bAplValorDosTit,nValorDosTit,bAplNumConCer,bAplNumMeses  From ProductoConcepto  Where nColocCred = 2 AND nPrdConceptoCod <> 1299  AND convert(varchar(15),nPrdConceptoCod) like '32%' --Order by convert(varchar(25),nPrdConceptoCod) 


