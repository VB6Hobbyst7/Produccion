Attribute VB_Name = "gConsContabilidad"
Global Const gOpePendOpeAgencias = "701301"
Global Const gOpePendFaltantCaja = "701302"
Global Const gOpePendRendirCuent = "701303"
Global Const gOpePendDisponibRes = "701304"
Global Const gOpePendCtaCobrarDi = "701305"
Global Const gOpePendPagoSubsidi = "701306"
Global Const gOpePendCtaCobraDiv = "701307"

Global Const gOpePendOrdendePago = "701320"
Global Const gOpePendCobraLiquid = "701321"
Global Const gOpePendSobraRemate = "701322"
Global Const gOpePendOtrasProvis = "701323"
Global Const gOpePendSobrantCaja = "701324"
Global Const gOpePendCanjeOPCheq = "701325"
Global Const gOpePendRecursHuman = "701326"
Global Const gOpePendOtrasOpeLiqPas = "701327"
Global Const gOpePendOPCertifica = "701328"
Global Const gOpePendMntHistoric = "701400"

Public Function MuestraListaRecordSet(prs As Recordset, Optional pnCol As Integer = 0) As String
Dim lsLista As String
If Not prs Is Nothing Then
   lsLista = ""
   Do While Not prs.EOF
      lsLista = lsLista & "'" & prs(pnCol) & "',"
      prs.MoveNext
   Loop
   If lsLista <> "" Then
      lsLista = Left(lsLista, Len(lsLista) - 1)
   End If
End If
MuestraListaRecordSet = lsLista
End Function
