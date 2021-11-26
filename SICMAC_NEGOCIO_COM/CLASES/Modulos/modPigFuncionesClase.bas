Attribute VB_Name = "modPigFuncionesClase"
'********************************************************
'** Modulo para funciones que se usaran en las clases de
'** Credito Pignoraticio
'********************************************************


Public Function mfgEstCredPigDesc(ByVal pnEstado As Integer) As String
Dim lsDesc As String
    Select Case pnEstado
        Case gColPEstRegis
            lsDesc = "Registrado"
        Case gColPEstDesem
            lsDesc = "Desembolsado"
        Case gColPEstDifer
            lsDesc = "Diferido (Para rescate)"
        Case gColPEstCance
            lsDesc = "Cancelado"
        Case gColPEstVenci
            lsDesc = "Vencido"
        Case gColPEstRemat
            lsDesc = "Remate"
        Case gColPEstPRema
            lsDesc = "Para Remate"
        Case gColPEstRenov
            lsDesc = "Renovado"
        Case gColPEstAdjud
            lsDesc = "Adjudicado"
        Case gColPEstSubas
            lsDesc = "Subastada"
        Case gColPEstAnula
            lsDesc = "Anulado"
        Case gColPEstChafa
            lsDesc = "Chafaloneado"
    '******************************************  Estados de Pignoraticios de la Caja de Lima ************************
        Case gPigEstRegis
            lsDesc = "Registrado"
        Case gPigEstDesemb
            lsDesc = "Desembolsado"
        Case gPigEstAmortiz
            lsDesc = "Amortizado"
        Case gPigEstReusoLin
            lsDesc = "Uso de Linea"
        Case gPigEstCancelPendRes
            lsDesc = "Pendiente de Rescate"
        Case gPigEstRescate
            lsDesc = "Rescatado"
        Case gPigEstRemat
            lsDesc = "En Remate"      'Contratos Activos
        Case gPigEstRematPRes
            lsDesc = "En Remate"      'Pendientes de Rescate
        Case gPigEstRematPFact
            lsDesc = "Para Facturar en Remate"
        Case gPigEstPResRematPFact
            lsDesc = "Para Facturar en Remate"
        Case gPigEstRematFact
            lsDesc = "Rematado"
        Case gPigEstAdjud
            lsDesc = "Adjudicado"
        Case gPigEstAnula
            lsDesc = "Anulado"
        Case 2814
            lsDesc = "Anulado no Desembolsado"
    End Select
    mfgEstCredPigDesc = lsDesc
End Function


