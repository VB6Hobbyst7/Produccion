VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit
Dim strCabecera As String, strCuerpo As String, strDetalle As String
Dim nLinea As Integer
Function ReportDet() As String
Dim Co As nCredRepoFinMes
Dim Rec As ADODB.Recordset
Dim sql As String
Dim xUbica As String
sql = "Select t8.cConsDescripcion cUbicaLote, "
sql = sql & " Substring(t7.cMovNro,9,2)+':'+"
sql = sql & " Substring(t7.cMovNro,11,2)+':' +"
sql = sql & " Substring(t7.cMovNro,13,2) hHora,"
sql = sql & " t7.nMovNro nLiquid, t2.cCtaCod cCta, t6.cPersNombre cNombre,"
sql = sql & " Sum (IsNull(t2.nTasacion, 0) + IsNull(t2.nTasacionAdicional, 0)) nTasacion"
sql = sql & " From ColocPigno t1"
sql = sql & "   Join ColocPigJoyaTasacion t2 On t2.cCtaCod = t1.cCtaCod"
sql = sql & "    and t2.nTipoTasacion = t1.nTipoTasacion"
sql = sql & "   Join Colocaciones t3 On t3.cCtaCod = t1.cCtaCod"
sql = sql & "   Join Producto t4 On t4.cCtaCod = t1.cCtaCod"
sql = sql & "   Join ProductoPersona t5 On t5.cCtaCod = t4.cCtaCod"
sql = sql & "   Join Persona t6 On t6.cPersCod = t5.cPersCod"
sql = sql & "   Join Mov t7 On t7.cMovNro = t3.cUltimaActualizacion"
sql = sql & "   Join Constante t8 On t8.nConsValor = t1.nUbicaLote"
sql = sql & " Where t4.nPrdEstado = " & gPigEstCancelPendRes
sql = sql & "  and t5.nPrdPersRelac = 20"
sql = sql & "  and t8.nConsCod = '8004'"
sql = sql & "  and convert(varchar,t4.dPrdEstado,103) = " & "'15/05/2003'"
sql = sql & " Group by t8.cConsDescripcion, substring(t7.cMovNro,9,2) + ':' + substring(t7.cMovNro,11,2) +':' + substring(t7.cMovNro,13,2),"
sql = sql & " t7.nMovNro , t2.cCtaCod, t6.cPersNombre"
sql = sql & " Order by t8.cConsDescripcion"

'La Cabecera
nLinea = 0
strCabecera = ""
strDetalle = ""
strCuerpo = ""
strCabecera = "Caja Metropolitana de Lima" & Space(91) & Date & Chr(10)
strCabecera = strCabecera & Space(42) & BoldOn & _
            "MOVIMIENTO DIARIO DE INGRESOS DE VALORES" & _
            BoldOff & Chr(10)
strCabecera = strCabecera & Space(37) & "EN CUSTODIA PENDIENTES DE RESCATE DEL DIA " & "15/05/2003" & Chr(10)
strCabecera = strCabecera & String(127, "-") & Chr(10)
strCabecera = strCabecera & Space(10) & "HORA " & Space(10) & "LIQUIDACION" & Space(16) & "CONTRATO" & Space(7) & "NOMBRE" & Space(46) & "TASACION" & Chr(10)
strCabecera = strCabecera & String(127, "-") & Chr(10)
nLinea = 6

strCuerpo = strCuerpo & strCabecera
Set Co = New nCredRepoFinMes
Set Rec = Co.GetQuery(sql)

If Not (Rec.EOF And Rec.BOF) Then
   Rec.MoveFirst
   xUbica = Rec!cUbicaLote
   Var1 = 0
   Var2 = 0
   Var3 = 0
   Var4 = 0
   VarTasacion = 0
   VarTotTasac = 0
   Do
       If xUbica <> Rec!cUbicaLote Then 'Imprime Grupo cuando Detalle pertenece a Otro Grupo
          xUbica = Rec!cUbicaLote
          strDetalle = strDetalle & Space(87) & "SUBTOTAL      : " & ImpreFormat(Var4, 4, 0) & _
                    Space(5) & ImpreFormat(VarTasacion, 12, , True) & Chr(10)
          nLinea = nLinea + 1
          Call CambioPag
          strDetalle = strDetalle & Chr(10) & "UBICACION : " & _
                         ImpreFormat(Rec!cUbicaLote, 20) & Chr(10)
          Call CambioPag
          Var2 = 0
          Var4 = 0
          VarTasacion = 0
       Else
          If xUbica = Rec!cUbicaLote Then
            If nLinea = 0 Or Var3 = 0 Then  'Imprime Grupo la Primera Vez
               strDetalle = strDetalle & "UBICACION : " & ImpreFormat(Rec!cUbicaLote, 20) & Chr(10)
               Call CambioPag
            Else 'No Imprime el Grupo si Detalle pertenece a mismo Grupo
                 strDetalle = strDetalle
            End If
          End If
       End If
       If nLinea = 0 Then  'Imprime Grupo la Primera Vez
            strDetalle = strDetalle & "UBICACION : " & ImpreFormat(Rec!cUbicaLote, 20) & Chr(10)
            Call CambioPag
       Else 'No Imprime el Grupo si Detalle pertenece a mismo Grupo
            strDetalle = strDetalle
       End If
       strDetalle = strDetalle & Space(10) & ImpreFormat(Rec!hHora, 8, 0)
       strDetalle = strDetalle & Space(5) & ImpreFormat(Rec!nLiquid, 12, 0)
       strDetalle = strDetalle & Space(7) & ImpreFormat(Rec!cCta, 18, 0)
       strDetalle = strDetalle & Space(5) & ImpreFormat(Rec!cNombre, 40)
       strDetalle = strDetalle & Space(5) & ImpreFormat(Rec!nTasacion, 12, , True) & Chr(10)
       Var3 = Var3 + 1
       Var4 = Var4 + 1
       VarTasacion = VarTasacion + Rec!nTasacion
       VarTotTasac = VarTotTasac + Rec!nTasacion
       Call CambioPag
       Rec.MoveNext
   Loop Until Rec.EOF
   Call CambioPag
   strDetalle = strDetalle & Space(87) & "SUBTOTAL      : " & ImpreFormat(Var4, 4, 0) & Space(5) & _
             ImpreFormat(VarTasacion, 12, , True) & Chr(10)
   Call CambioPag
   strCuerpo = strCuerpo & strDetalle
   strCuerpo = strCuerpo & Chr(10) & Space(87) & "TOTAL GENERAL : " & ImpreFormat(Var3, 4, 0) & Space(5) & _
               ImpreFormat(VarTotTasac, 12, , True) & Chr(10)
End If
ReportDet = strCuerpo

End Function
Public Sub CambioPag()
    If nLinea > 59 Then
       strCuerpo = strCuerpo & strDetalle & Chr(12)
       nLinea = 0
       strDetalle = ""
    Else
       If nLinea = 0 Then
          strDetalle = strCabecera & strDetalle
          nLinea = 6
       Else
          strDetalle = strDetalle
          nLinea = nLinea + 1
       End If
    End If
End Sub


