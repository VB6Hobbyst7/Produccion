VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NPigReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit
Dim sCad As String
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lnAncho As Integer
Dim sCadCabezera As String
Dim sCadCabezeraHeader As String
Dim sCadTitulo As String
Dim sCadSubTitulo As String
Dim sCadComenta As String
Dim scadColumnas  As String
Dim csNomCMAC As String
Dim csNomAgencia As String
Dim csCodUser As String
Dim csFechaSis As String
Dim strCabecera As String, strCuerpo As String, strDetalle As String
Dim nLinea As Integer

Public Sub inicio(ByVal psNomCmac As String, ByVal psNomAgencia As String, _
            ByVal psCodUser As String, ByVal psFechaSis As String)
    csNomCMAC = psNomCmac
    csNomAgencia = psNomAgencia
    csCodUser = psCodUser
    csFechaSis = psFechaSis
End Sub

'sCad = sCad & Chr$(15)                           'Retorna al tipo de letra normal
'sCad = sCad & Chr$(27) + Chr$(107) + Chr$(1)     'Tipo de Letra Roman
'sCad = sCad & Chr$(27) + Chr$(103)                'Tamaño 10.5 - 15 CPI
'sCad = sCad & Chr$(27) + Chr$(50)                'Espaciado entre lineas 1/16
'sCad = sCad & Chr$(27) + Chr$(67) + Chr$(18)     'Longitud de página a 24 líneas
'sCad = sCad & Chr$(27) + Chr$(108) + Chr$(6)     'Margen Izquierdo - 6ta. Columna


'   --- Constantes para Impresión  ------
' Introduce espacio                    :   Space(43)
' Enter para la siguiente linea   :   Chr$(10)
' Determina la cantidad de
' espacios devueltos                 :    ImpreFormat("campo", Cantidad de espacio, Decimales,¿moneda?)








'Codigo para la impresion de reportes para Boveda

Public Function RepBoveda(ByVal pCodAge As Variant) As String


'Variables para el Manejo de losDatos
Dim Co As nCredRepoFinMes
Dim rs As ADODB.Recordset
Dim strString As String                     ' Variable que contiene la consulta


'Negrita
Dim BoldOn As Variant
Dim BoldOff As Variant
BoldOn = Chr$(27) & Chr$(69)
BoldOff = Chr$(27) & Chr$(70)

'Varibles que pertenecen al cuerpo del reporte


Dim nItem As Integer

Dim cTitulo As String

Set Co = New nCredRepoFinMes

strString = " select Estado = CASE a.nEstadoCont WHEN 2802 THEN 0 WHEN 2803 THEN 0 WHEN 2804 THEN 0 ELSE 1 END, "
strString = strString & " c.cAgeCod+' '+c.cAgeDescripcion as 'Agencia', "
strString = strString & " Count(distinct a.cctacod) as 'Lotes', "
strString = strString & " sum(nPesoBruto) as 'Peso Bruto',sum(nPesoNeto) as 'Peso Neto', "
strString = strString & " sum(nTasacion) as 'Tasacion' ,sum(nTasacionAdicional) as 'Tasacion Adicional'"
strString = strString & " from colocpigno a "
strString = strString & " inner join colocpigjoyatasacion  b on b.cctacod=a.cctacod and b.ntipotasacion=a.ntipotasacion "
strString = strString & " inner join agencias c on c.cAgeCod = substring(a.cctacod,4,2) "
strString = strString & " inner join colocpigubicacion d on d.cUbiCod = a.nUbicaLote and d.cAgeCod= " & pCodAge
strString = strString & " where a.nEstadoCont in (" & gPigEstAmortiz & "," & gPigEstCancelPendRes & "," & gPigEstDesemb & "," & gPigEstReusoLin & ")"
strString = strString & " group by  CASE a.nEstadoCont WHEN 2802 THEN 0 WHEN 2803 THEN 0 WHEN 2804 THEN 0 ELSE 1 END,c.cAgeCod+' '+c.cAgeDescripcion "
strString = strString & " order by  c.cAgeCod+' '+c.cAgeDescripcion"

Set rs = Co.GetQuery(strString)

'Inicializar variables
sCad = ""
scadColumnas = ""
'Cabecera para el Reporte
lnLineas = 8
lnPage = 1
lnAncho = 127
sCadTitulo = "CREDITO PIGNORATICIO"
sCadSubTitulo = "STOCK DE LOTES POR UBICACION"
sCadComenta = ""
scadColumnas = Chr(10) & scadColumnas & String(120, "-") & Chr(10)
scadColumnas = scadColumnas & Space(2) & "OFICINA" & Space(36) & "LOTES " & Space(8) & "P. BRUTO" & Space(10) & "P. NETO" & Space(9) & "TASACION" & Space(4) & "TASACION ADIC." & Chr(10)
scadColumnas = scadColumnas & String(120, "-") & Chr(10) & Chr(10)
sCad = sCad & nArmaCabeceraReporte(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, sCadTitulo, sCadSubTitulo, lnPage, lnAncho, scadColumnas, sCadComenta)


If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    
    sCadGarantia = sCadGarantia & "VALORES EN GARANTIA (CONTRATOS VIGENTES)" & Chr(10) & Chr(10)
    sCadCustodia = sCadCustodia & "VALORES EN CUSTODIA (RESCATES PENDIENTES)" & Chr(10) & Chr(10)
    
     Do
     lnLineas = lnLineas + 1
         If rs!Estado = 0 Then
               
                sCadGarantia = sCadGarantia & ImpreFormat(rs!Agencia, 25) & Space(10)
                sCadGarantia = sCadGarantia & ImpreFormat(rs("Lotes"), 13, 0, True) & Space(1)
                sCadGarantia = sCadGarantia & ImpreFormat(rs("Peso Bruto"), 13, 2, True) & Space(1)
                sCadGarantia = sCadGarantia & ImpreFormat(rs("Peso Neto"), 13, 2, True) & Space(1)
                sCadGarantia = sCadGarantia & ImpreFormat(rs!Tasacion, 13, 2, True) & Space(1)
                sCadGarantia = sCadGarantia & ImpreFormat(rs("Tasacion Adicional"), 13, 2, True) & Chr$(10)
               
               'Suma para Totalizar
                sCadTotLotes1 = sCadTotLotes1 + rs("Lotes")
                sCadTotPBruto1 = sCadTotPBruto1 + rs("Peso Bruto")
                sCadTotPNeto1 = sCadTotPNeto1 + rs("Peso Neto")
                sCadTotTasaccion1 = sCadTotTasaccion1 + rs!Tasacion
                sCadTotTasaccionAdicional1 = sCadTotTasaccionAdicional1 + rs("Tasacion Adicional")
         End If
        
        
        
        If rs!Estado = 1 Then
               
                sCadCustodia = sCadCustodia & ImpreFormat(rs!Agencia, 25) & Space(10)
                 sCadCustodia = sCadCustodia & ImpreFormat(rs("Lotes"), 13, 0, True) & Space(1)
                sCadCustodia = sCadCustodia & ImpreFormat(rs("Peso Bruto"), 13, 2, True) & Space(1)
                sCadCustodia = sCadCustodia & ImpreFormat(rs("Peso Neto"), 13, 2, True) & Space(1)
                sCadCustodia = sCadCustodia & ImpreFormat(rs!Tasacion, 13, 2, True) & Space(1)
                sCadCustodia = sCadCustodia & ImpreFormat(rs("Tasacion Adicional"), 13, 2, True) & Chr$(10)
                               
                'Suma para Totalizar
                sCadTotLotes2 = sCadTotLotes2 + rs("Lotes")
                sCadTotPBruto2 = sCadTotPBruto2 + rs("Peso Bruto")
                sCadTotPNeto2 = sCadTotPNeto2 + rs("Peso Neto")
                sCadTotTasaccion2 = sCadTotTasaccion2 + rs!Tasacion
                sCadTotTasaccionAdicional2 = sCadTotTasaccionAdicional2 + rs("Tasacion Adicional")

         End If
       
         CambiarPagina (0)
       
           rs.MoveNext
          
       Loop Until rs.EOF
                         
             ' sCadTot1 = sCadTot1 & String(120, "-") & Chr(10)
              sCadTot1 = Space(37) & ImpreFormat(sCadTotLotes1, 13, 0, True) & Space(1) & ImpreFormat(sCadTotPBruto1, 13, 2, True) & Space(1)
              sCadTot1 = sCadTot1 & ImpreFormat(sCadTotPNeto1, 13, 2, True) & Space(1) & ImpreFormat(sCadTotTasaccion1, 13, 2, True) & Space(1) & ImpreFormat(sCadTotTasaccionAdicional1, 13, 2, True)
              
              sCadTot2 = sCadTot2 & String(110, "-") & Chr(10)
              sCadTot2 = Space(37) & ImpreFormat(sCadTotLotes2, 13, 0, True) & Space(1) & ImpreFormat(sCadTotPBruto2, 13, 2, True) & Space(1)
              sCadTot2 = sCadTot2 & ImpreFormat(sCadTotPNeto2, 13, 2, True) & Space(1) & ImpreFormat(sCadTotTasaccion2, 13, 2, True) & Space(1) & ImpreFormat(sCadTotTasaccionAdicional2, 13, 2, True)
                     
                         
               sCad = sCad & sCadGarantia & Space(2) & String(118, "-") & Chr(10) & sCadTot1 & Chr(10) & sCadCustodia & Space(2) & String(118, "-") & Chr(10) & sCadTot2
         
'Devuelve toda la estructura  del Reporte
        

 End If

 RepBoveda = sCad

End Function

Public Function RepPiezasSelFundicion() As String


'Variables para el Manejo de losDatos
Dim Co As nCredRepoFinMes
Dim rs As ADODB.Recordset
Dim strString As String                     ' Variable que contiene la consulta


'Negrita
Dim BoldOn As Variant
Dim BoldOff As Variant
BoldOn = Chr$(27) & Chr$(69)
BoldOff = Chr$(27) & Chr$(70)

Dim nItem As Integer

Dim cTitulo As String

Set Co = New nCredRepoFinMes

strString = " select CPEJ.nremate as 'Remate',SUBSTRING(CP.cCtaCod,4,2) AS 'CodAgencia',E.cAgeDescripcion AS 'Agencia', CPJT.cCtaCod as 'Codigo', CPJT.nItemPieza as Pieza,CPJT.cDescripcion as 'Descripcion',CPJT.nMaterial as 'Material',CPJT.nPesoBruto as 'Peso Bruto',CPJT.nPesoNeto as 'Peso Neto',(CPJT.nTasacion +CPJT.nTasacionAdicional) as 'Tasacion' ,CPP.nValorProceso as 'ValorAdj' ,CPP.nValorAdicional as 'ValorAdic'"
strString = strString & " FROM  ColocPigJoyaTasacion CPJT "
strString = strString & " INNER JOIN ColocPigno CP on CP.cCtaCod = CPJT.cCtaCod and CP.nTipoTasacion = CPJT.nTipoTasacion  "
strString = strString & " INNER JOIN ColocPigEstadoJoya  CPEJ on CPJT.cCtaCod =CPEJ.cCtacod and CPJT.nItemPieza = CPEJ.nItemPieza "
strString = strString & " INNER JOIN ColocPigProceso CPP on CPP.nRemate=CPEJ.nRemate and CPP.nTipoProceso = CPEJ.nEstadoJoya and CPP.cCtaCod =CPEJ.cCtacod and CPP.nItemPieza = CPEJ.nItemPieza  "
strString = strString & " INNER JOIN    Agencias E ON   E.cAgeCod  =   SUBSTRING(CP.cCtaCod,4,2) "
strString = strString & " WHERE CPP.nSituacionPieza= " & gPigSituacionDisponible & " and CPP.nTipoProceso= " & gPigTipoFundicion
strString = strString & " order by Remate, CodAgencia "

Set rs = Co.GetQuery(strString)

'Inicializar variables
sCad = ""
scadColumnas = ""
'scad = scad & Chr$(15)                           'Retorna al tipo de letra normal
'scad = scad & Chr$(27) + Chr$(107) + Chr$(1)     'Tipo de Letra Roman
'scad = scad & Chr$(27) + Chr$(103)                'Tamaño 10.5 - 15 CPI
'scad = scad & Chr$(27) + Chr$(50)                'Espaciado entre lineas 1/16
''sCad = sCad & Chr$(27) + Chr$(67) + Chr$(18)     'Longitud de página a 24 líneas
'scad = scad & Chr$(27) + Chr$(108) + Chr$(6)     'Margen Izquierdo - 6ta. Columna


'   --- Constantes para Impresión  ------
' Introduce espacio                    :   Space(43)
' Enter para la siguiente linea   :   Chr$(10)
' Determina la cantidad de
' espacios devueltos                 :    ImpreFormat("campo", Cantidad de espacio, Decimales,¿moneda?)

'Cabecera para el Reporte
lnLineas = 8
lnPage = 1
lnAncho = 133
sCadTitulo = "CUSTODIA DE VALORES"
sCadSubTitulo = "PIEZAS SELECCIONADAS PARA FUNDICION"
sCadComenta = ""
scadColumnas = Chr(10) & scadColumnas & String(127, "-") & Chr(10)
scadColumnas = scadColumnas & Space(20) & "PIEZA " & Space(8) & "OBSERVACION" & Space(4) & "MATERIAL" & Space(6) & "P. BRUTO" & Space(7) & "P. NETO" & Space(6) & "TASACION" & Space(5) & "VALOR ADJ" & Space(4) & "VALOR ADIC" & Chr(10)
scadColumnas = scadColumnas & String(127, "-") & Chr(10) & Chr(10)
sCad = sCad & nArmaCabeceraReporte(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, sCadTitulo, sCadSubTitulo, lnPage, lnAncho, scadColumnas, sCadComenta)


If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    sCadRemate = rs!Remate
    sCadCodAgencia = rs!CodAgencia
    Do
        lnLineas = lnLineas + 1
        x = x + 1
        If rs.EOF Then
        ultimo = 1
        rs.MoveLast
        End If
        If (rs!CodAgencia <> sCadCodAgencia) Or (rs!Remate <> sCadRemate) Or (x = 1) Or (ultimo = 1) Then
           'Totalizo Agencia
           sCadTotAgencia = Space(2) & ImpreFormat(sCadTotPiezasAgencia, 4, 0, True) & " piezas"
           sCadTotAgencia = sCadTotAgencia & Space(25) & ImpreFormat(sCadTotPesoBrutoAgencia, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(1) & ImpreFormat(sCadTotPesoNetoAgencia, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(1) & ImpreFormat(sCadTotTasacionAgencia, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(1) & ImpreFormat(sCadTotValorAdjAgencia, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(1) & ImpreFormat(sCadTotValorAdicAgencia, 10, 2, True)
           sCadTotPiezasAgencia = 0
           sCadTotPesoBrutoAgencia = 0
           sCadTotPesoNetoAgencia = 0
           sCadTotTasacionAgencia = 0
           sCadTotValorAdjAgencia = 0
           sCadTotValorAdicAgencia = 0
                      
           If x <> 1 Then sCad = sCad & Chr$(10) & Space(5) & " Total Agencia:" & sCadTotAgencia & Chr$(10)
           
           If (rs!Remate <> sCadRemate) Or (x = 1) Or (ultimo = 1) Then
              'Totalizo Remate
               sCadTotRemate = Space(5) & ImpreFormat(sCadTotPiezasRemate, 4, 0, True) & " piezas"
               sCadTotRemate = sCadTotRemate & Space(25) & ImpreFormat(sCadTotPesoBrutoRemate, 10, 2, True)
               sCadTotRemate = sCadTotRemate & Space(1) & ImpreFormat(sCadTotPesoNetoRemate, 10, 2, True)
               sCadTotRemate = sCadTotRemate & Space(1) & ImpreFormat(sCadTotTasacionRemate, 10, 2, True)
               sCadTotRemate = sCadTotRemate & Space(1) & ImpreFormat(sCadTotValorAdjRemate, 10, 2, True)
               sCadTotRemate = sCadTotRemate & Space(1) & ImpreFormat(sCadTotValorAdicRemate, 10, 2, True)
           
               sCadTotPiezasRemate = 0
               sCadTotPesoBrutoRemate = 0
               sCadTotPesoNetoRemate = 0
               sCadTotTasacionRemate = 0
               sCadTotValorAdjRemate = 0
               sCadTotValorAdicRemate = 0
               
               If x <> 1 Then sCad = sCad & Chr$(10) & "  Total Remate:  " & sCadTotRemate & Chr$(10)
              
              'Encabezo Remate
               sCadRemate = rs!Remate
               sCadTitRemate = "  REMATE: " & ImpreFormat(sCadRemate, 4, 0, True)
           
               If ultimo <> 1 Then sCad = sCad & Chr$(10) & sCadTitRemate & Chr$(10)
           End If
           'Encabezo Agencia
           sCadCodAgencia = rs!CodAgencia
           sCadTitAgencia = Space(5) & "(" & ImpreFormat(rs!CodAgencia, 2, 0, True) & ")" & Space(1) & ImpreFormat(rs!Agencia, 25)
        
           If ultimo <> 1 Then sCad = sCad & Chr$(10) & sCadTitAgencia & Chr$(10) & Chr$(10)
           End If
       'Añado Detalle
       If ultimo = 1 Then GoTo davi
       SCADDETALLE = SCADDETALLE & Space(12) & Justifica(rs!Codigo, 18) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs("Pieza"), 2, 0, False) & Space(1)
'       sCadDetalle = sCadDetalle & ImpreFormat(Rs("Descripcion"), 14, 0, False) & Space(1)
        SCADDETALLE = SCADDETALLE & Justifica(rs("Descripcion"), 14) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs("Material"), 8, 0, False) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs("Peso Bruto"), 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs("Peso Neto"), 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Tasacion, 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!ValorAdj, 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!ValorAdic, 10, 2, True) & Chr$(10)
    
       sCad = sCad & SCADDETALLE
       SCADDETALLE = ""
       '************Segundo Quiebre****************
       'Suma para Totalizar por Agencia
       sCadTotPiezasAgencia = sCadTotPiezasAgencia + 1
       sCadTotPesoBrutoAgencia = sCadTotPesoBrutoAgencia + rs("Peso Bruto")
       sCadTotPesoNetoAgencia = sCadTotPesoNetoAgencia + rs("Peso Neto")
       sCadTotTasacionAgencia = sCadTotTasacionAgencia + rs!Tasacion
       sCadTotValorAdjAgencia = sCadTotValorAdjAgencia + rs!ValorAdj
       sCadTotValorAdicAgencia = sCadTotValorAdicAgencia + rs!ValorAdic
       
       
       '************Primer Quiebre****************
       'Suma para Totalizar por Remate
       sCadTotPiezasRemate = sCadTotPiezasRemate + 1
       sCadTotPesoBrutoRemate = sCadTotPesoBrutoRemate + rs("Peso Bruto")
       sCadTotPesoNetoRemate = sCadTotPesoNetoRemate + rs("Peso Neto")
       sCadTotTasacionRemate = sCadTotTasacionRemate + rs!Tasacion
       sCadTotValorAdjRemate = sCadTotValorAdjRemate + rs!ValorAdj
       sCadTotValorAdicRemate = sCadTotValorAdicRemate + rs!ValorAdic
       
       'Suma para Totalizar por Remate
       sCadTotPiezas = sCadTotPiezas + 1
       sCadTotPesoBruto = sCadTotPesoBruto + rs("Peso Bruto")
       sCadTotPesoNeto = sCadTotPesoNeto + rs("Peso Neto")
       sCadTotTasacion = sCadTotTasacion + rs!Tasacion
       sCadTotValorAdj = sCadTotTasacion + rs!ValorAdj
       sCadTotValorAdic = sCadTotTasacion + rs!ValorAdic
       
       rs.MoveNext
       CambiarPagina (7)

    Loop Until x = rs.RecordCount + 1
davi:
       sCad = sCad
       sCadResumen = Chr$(10) & "  Total General: " & Chr$(10) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "CANTIDAD   :" & ImpreFormat(sCadTotPiezas, 4, 0, True) & " piezas" & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "PESO BRUTO :" & ImpreFormat(sCadTotPesoBruto, 13, 2, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "PESO NETO  :" & ImpreFormat(sCadTotPesoNeto, 13, 2, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "TASACION   :" & ImpreFormat(sCadTotTasacion, 13, 2, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "VALOR ADJ  :" & ImpreFormat(sCadTotValorAdj, 13, 2, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "VALOR ADIC :" & ImpreFormat(sCadTotValorAdic, 13, 2, True) & Chr$(10)
       
       sCad = sCad & sCadResumen
       

 End If
 
      'Devuelve toda la estructura  del Reporte
      RepPiezasSelFundicion = sCad

End Function
Public Function RepPiezasSelFundicionPorRemate() As String

'Variables para el Manejo de losDatos
Dim Co As nCredRepoFinMes
Dim rs As ADODB.Recordset
Dim strString As String                     ' Variable que contiene la consulta

'Negrita
Dim BoldOn As Variant
Dim BoldOff As Variant
BoldOn = Chr$(27) & Chr$(69)
BoldOff = Chr$(27) & Chr$(70)

'Varibles que pertenecen al cuerpo del reporte
Dim sCad As String

' Variables que pasaran a Totalizar
Dim sCadTotPiezas As Variant
Dim sCadTotPBruto As Variant
Dim sCadTotPNeto As Variant
Dim sCadTotTasaccion As Variant
Dim sCadTotValorAdj As Variant
Dim sCadTotValorAdic As Variant

Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer

Dim cTitulo As String

Set Co = New nCredRepoFinMes

strString = " select CPEJ.nremate as 'Remate', count(*) as 'Cantidad', sum(CPJT.nPesoBruto) as 'Peso Bruto', sum(CPJT.nPesoNeto) as 'Peso Neto', sum(CPJT.nTasacion +CPJT.nTasacionAdicional) as 'Tasacion' ,sum(CPP.nValorProceso) as 'ValorAdj' ,sum(CPP.nValorAdicional) as 'ValorAdic'"
strString = strString & " FROM  ColocPigJoyaTasacion CPJT "
strString = strString & " INNER JOIN ColocPigno CP on CP.cCtaCod = CPJT.cCtaCod and CP.nTipoTasacion = CPJT.nTipoTasacion  "
strString = strString & " INNER JOIN ColocPigEstadoJoya  CPEJ on CPJT.cCtaCod =CPEJ.cCtacod and CPJT.nItemPieza = CPEJ.nItemPieza "
strString = strString & " INNER JOIN ColocPigProceso CPP on CPP.nRemate=CPEJ.nRemate and CPP.nTipoProceso = CPEJ.nEstadoJoya and CPP.cCtaCod =CPEJ.cCtacod and CPP.nItemPieza = CPEJ.nItemPieza  "
strString = strString & " WHERE CPP.nSituacionPieza= " & gPigSituacionDisponible & " and CPP.nTipoProceso= " & gPigTipoFundicion
strString = strString & " group by CPEJ.nremate "
strString = strString & " order by CPEJ.nremate "

Set rs = Co.GetQuery(strString)

'Inicializar variables
sCad = ""
scadColumnas = ""

'Cabecera para el Reporte
lnLineas = 8
lnPage = 1
lnAncho = 126
sCadTitulo = "CUSTODIA DE VALORES"
sCadSubTitulo = "PIEZAS SELECCIONADAS PARA FUNDICION(RESUMEN)"
sCadComenta = "Ubicación : Agencia Principal"
scadColumnas = Chr(10) & scadColumnas & String(120, "-") & Chr(10)
scadColumnas = scadColumnas & Space(7) & "REMATE" & Space(8) & "CANTIDAD" & Space(10) & "P. BRUTO" & Space(11) & "P. NETO" & Space(10) & "TASACION" & Space(9) & "VALOR ADJ" & Space(8) & "VALOR ADIC" & Chr(10)
scadColumnas = scadColumnas & String(120, "-") & Chr(10) & Chr(10)
sCad = sCad & nArmaCabeceraReporte(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, sCadTitulo, sCadSubTitulo, lnPage, lnAncho, scadColumnas, sCadComenta)

If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    
    Do
     lnLineas = lnLineas + 1
               
                sCad = sCad & Space(10) & ImpreFormat(rs!Remate, 3, 0, False) & Space(3)
                sCad = sCad & ImpreFormat(rs("Cantidad"), 13, 0, True) & Space(3)
                sCad = sCad & ImpreFormat(rs("Peso Bruto"), 13, 2, True) & Space(2)
                sCad = sCad & ImpreFormat(rs("Peso Neto"), 13, 2, True) & Space(2)
                sCad = sCad & ImpreFormat(rs!Tasacion, 13, 2, True) & Space(2)
                sCad = sCad & ImpreFormat(rs!ValorAdj, 13, 2, True) & Space(2)
                sCad = sCad & ImpreFormat(rs!ValorAdic, 13, 2, True) & Chr$(10)
               
               'Suma para Totalizar
                sCadTotPiezas = sCadTotPiezas + rs("Cantidad")
                sCadTotPBruto = sCadTotPBruto + rs("Peso Bruto")
                sCadTotPNeto = sCadTotPNeto + rs("Peso Neto")
                sCadTotTasacion = sCadTotTasacion + rs!Tasacion
                sCadTotValorAdj = sCadTotValorAdj + rs!ValorAdj
                sCadTotValorAdic = sCadTotValorAdic + rs!ValorAdic
            
                CambiarPagina (0)
                rs.MoveNext
          
       Loop Until rs.EOF
                         
             ' sCadTot1 = sCadTot1 & String(120, "-") & Chr(10)
              sCadTot1 = Space(7) & "TOTALES :" & ImpreFormat(sCadTotPiezas, 13, 0, True) & Space(3) & ImpreFormat(sCadTotPBruto, 13, 2, True) & Space(2)
              sCadTot1 = sCadTot1 & ImpreFormat(sCadTotPNeto, 13, 2, True) & Space(2) & ImpreFormat(sCadTotTasacion, 13, 2, True) & Space(2) & ImpreFormat(sCadTotValorAdj, 13, 2, True) & Space(2) & ImpreFormat(sCadTotValorAdic, 13, 2, True)
    
              sCad = sCad & Space(2) & String(118, "-") & Chr(10) & sCadTot1 & Chr(10)
         
'Devuelve toda la estructura  del Reporte
        
 End If

 RepPiezasSelFundicionPorRemate = sCad

End Function

Public Function RepPiezasSelVenta() As String


'Variables para el Manejo de losDatos
Dim Co As nCredRepoFinMes
Dim rs As ADODB.Recordset
Dim strString As String                     ' Variable que contiene la consulta


'Negrita
Dim BoldOn As Variant
Dim BoldOff As Variant
BoldOn = Chr$(27) & Chr$(69)
BoldOff = Chr$(27) & Chr$(70)



Dim nItem As Integer

Dim cTitulo As String

Set Co = New nCredRepoFinMes

strString = " select CPEJ.nremate as 'Remate',SUBSTRING(CP.cCtaCod,4,2) AS 'CodAgencia',E.cAgeDescripcion AS 'Agencia', CPJT.cCtaCod as 'Codigo', CPJT.nItemPieza as Pieza,CPJT.cDescripcion as 'Descripcion',CPJT.nMaterial as 'Material',CPJT.nPesoBruto as 'Peso Bruto',CPJT.nPesoNeto as 'Peso Neto',(CPJT.nTasacion +CPJT.nTasacionAdicional) as 'Tasacion' ,CPP.nValorProceso as 'ValorAdj' ,CPP.nValorAdicional as 'ValorAdic'"
strString = strString & " FROM  ColocPigJoyaTasacion CPJT "
strString = strString & " INNER JOIN ColocPigno CP on CP.cCtaCod = CPJT.cCtaCod and CP.nTipoTasacion = CPJT.nTipoTasacion  "
strString = strString & " INNER JOIN ColocPigEstadoJoya  CPEJ on CPJT.cCtaCod =CPEJ.cCtacod and CPJT.nItemPieza = CPEJ.nItemPieza "
strString = strString & " INNER JOIN ColocPigProceso CPP on CPP.nRemate=CPEJ.nRemate and CPP.nTipoProceso = CPEJ.nEstadoJoya and CPP.cCtaCod =CPEJ.cCtacod and CPP.nItemPieza = CPEJ.nItemPieza  "
strString = strString & " INNER JOIN    Agencias E ON   E.cAgeCod  =   SUBSTRING(CP.cCtaCod,4,2) "
strString = strString & " WHERE CPP.nSituacionPieza= " & gPigSituacionDisponible & " and CPP.nTipoProceso =" & gPigTipoVentas
strString = strString & " order by Remate, CodAgencia "

Set rs = Co.GetQuery(strString)

'Inicializar variables
sCad = ""
scadColumnas = ""
'scad = scad & Chr$(15)                           'Retorna al tipo de letra normal
'scad = scad & Chr$(27) + Chr$(107) + Chr$(1)     'Tipo de Letra Roman
'scad = scad & Chr$(27) + Chr$(103)                'Tamaño 10.5 - 15 CPI
'scad = scad & Chr$(27) + Chr$(50)                'Espaciado entre lineas 1/16
''sCad = sCad & Chr$(27) + Chr$(67) + Chr$(18)     'Longitud de página a 24 líneas
'scad = scad & Chr$(27) + Chr$(108) + Chr$(6)     'Margen Izquierdo - 6ta. Columna


'   --- Constantes para Impresión  ------
' Introduce espacio                    :   Space(43)
' Enter para la siguiente linea   :   Chr$(10)
' Determina la cantidad de
' espacios devueltos                 :    ImpreFormat("campo", Cantidad de espacio, Decimales,¿moneda?)

'Cabecera para el Reporte
lnLineas = 8
lnPage = 1
lnAncho = 133
sCadTitulo = "CUSTODIA DE VALORES"
sCadSubTitulo = "PIEZAS SELECCIONADAS PARA VENTA"
sCadComenta = ""
scadColumnas = Chr(10) & scadColumnas & String(127, "-") & Chr(10)
scadColumnas = scadColumnas & Space(20) & "PIEZA " & Space(8) & "OBSERVACION" & Space(4) & "MATERIAL" & Space(6) & "P. BRUTO" & Space(7) & "P. NETO" & Space(6) & "TASACION" & Space(5) & "VALOR ADJ" & Space(4) & "VALOR ADIC" & Chr(10)
scadColumnas = scadColumnas & String(127, "-") & Chr(10) & Chr(10)
sCad = sCad & nArmaCabeceraReporte(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, sCadTitulo, sCadSubTitulo, lnPage, lnAncho, scadColumnas, sCadComenta)



If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    sCadRemate = rs!Remate
    sCadCodAgencia = rs!CodAgencia
    Do
        lnLineas = lnLineas + 1
        x = x + 1
        If rs.EOF Then
        ultimo = 1
        rs.MoveLast
        End If
        If (rs!CodAgencia <> sCadCodAgencia) Or (rs!Remate <> sCadRemate) Or (x = 1) Or (ultimo = 1) Then
           'Totalizo Agencia
           sCadTotAgencia = Space(2) & ImpreFormat(sCadTotPiezasAgencia, 4, 0, True) & " piezas"
           sCadTotAgencia = sCadTotAgencia & Space(25) & ImpreFormat(sCadTotPesoBrutoAgencia, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(1) & ImpreFormat(sCadTotPesoNetoAgencia, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(1) & ImpreFormat(sCadTotTasacionAgencia, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(1) & ImpreFormat(sCadTotValorAdjAgencia, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(1) & ImpreFormat(sCadTotValorAdicAgencia, 10, 2, True)
           sCadTotPiezasAgencia = 0
           sCadTotPesoBrutoAgencia = 0
           sCadTotPesoNetoAgencia = 0
           sCadTotTasacionAgencia = 0
           sCadTotValorAdjAgencia = 0
           sCadTotValorAdicAgencia = 0
                      
           If x <> 1 Then sCad = sCad & Chr$(10) & Space(5) & " Total Agencia:" & sCadTotAgencia & Chr$(10)
           
           If (rs!Remate <> sCadRemate) Or (x = 1) Or (ultimo = 1) Then
              'Totalizo Remate
               sCadTotRemate = Space(5) & ImpreFormat(sCadTotPiezasRemate, 4, 0, True) & " piezas"
               sCadTotRemate = sCadTotRemate & Space(25) & ImpreFormat(sCadTotPesoBrutoRemate, 10, 2, True)
               sCadTotRemate = sCadTotRemate & Space(1) & ImpreFormat(sCadTotPesoNetoRemate, 10, 2, True)
               sCadTotRemate = sCadTotRemate & Space(1) & ImpreFormat(sCadTotTasacionRemate, 10, 2, True)
               sCadTotRemate = sCadTotRemate & Space(1) & ImpreFormat(sCadTotValorAdjRemate, 10, 2, True)
               sCadTotRemate = sCadTotRemate & Space(1) & ImpreFormat(sCadTotValorAdicRemate, 10, 2, True)
           
               sCadTotPiezasRemate = 0
               sCadTotPesoBrutoRemate = 0
               sCadTotPesoNetoRemate = 0
               sCadTotTasacionRemate = 0
               sCadTotValorAdjRemate = 0
               sCadTotValorAdicRemate = 0
               
               If x <> 1 Then sCad = sCad & Chr$(10) & "  Total Remate:  " & sCadTotRemate & Chr$(10)
              
              'Encabezo Remate
               sCadRemate = rs!Remate
               sCadTitRemate = "  REMATE: " & ImpreFormat(sCadRemate, 4, 0, True)
           
               If ultimo <> 1 Then sCad = sCad & Chr$(10) & sCadTitRemate & Chr$(10)
           End If
           'Encabezo Agencia
           sCadCodAgencia = rs!CodAgencia
           sCadTitAgencia = Space(5) & "(" & ImpreFormat(rs!CodAgencia, 2, 0, True) & ")" & Space(1) & ImpreFormat(rs!Agencia, 25)
        
           If ultimo <> 1 Then sCad = sCad & Chr$(10) & sCadTitAgencia & Chr$(10) & Chr$(10)
           End If
       'Añado Detalle
       If ultimo = 1 Then GoTo davi
       SCADDETALLE = SCADDETALLE & Space(12) & Justifica(rs!Codigo, 18) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs("Pieza"), 2, 0, False) & Space(1)
'       sCadDetalle = sCadDetalle & ImpreFormat(Rs("Descripcion"), 14, 0, False) & Space(1)
        SCADDETALLE = SCADDETALLE & Justifica(rs("Descripcion"), 14) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs("Material"), 8, 0, False) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs("Peso Bruto"), 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs("Peso Neto"), 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Tasacion, 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!ValorAdj, 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!ValorAdic, 10, 2, True) & Chr$(10)
    
       sCad = sCad & SCADDETALLE
       SCADDETALLE = ""
       '************Segundo Quiebre****************
       'Suma para Totalizar por Agencia
       sCadTotPiezasAgencia = sCadTotPiezasAgencia + 1
       sCadTotPesoBrutoAgencia = sCadTotPesoBrutoAgencia + rs("Peso Bruto")
       sCadTotPesoNetoAgencia = sCadTotPesoNetoAgencia + rs("Peso Neto")
       sCadTotTasacionAgencia = sCadTotTasacionAgencia + rs!Tasacion
       sCadTotValorAdjAgencia = sCadTotValorAdjAgencia + rs!ValorAdj
       sCadTotValorAdicAgencia = sCadTotValorAdicAgencia + rs!ValorAdic
       
       
       '************Primer Quiebre****************
       'Suma para Totalizar por Remate
       sCadTotPiezasRemate = sCadTotPiezasRemate + 1
       sCadTotPesoBrutoRemate = sCadTotPesoBrutoRemate + rs("Peso Bruto")
       sCadTotPesoNetoRemate = sCadTotPesoNetoRemate + rs("Peso Neto")
       sCadTotTasacionRemate = sCadTotTasacionRemate + rs!Tasacion
       sCadTotValorAdjRemate = sCadTotValorAdjRemate + rs!ValorAdj
       sCadTotValorAdicRemate = sCadTotValorAdicRemate + rs!ValorAdic
       
       'Suma para Totalizar por Remate
       sCadTotPiezas = sCadTotPiezas + 1
       sCadTotPesoBruto = sCadTotPesoBruto + rs("Peso Bruto")
       sCadTotPesoNeto = sCadTotPesoNeto + rs("Peso Neto")
       sCadTotTasacion = sCadTotTasacion + rs!Tasacion
       sCadTotValorAdj = sCadTotTasacion + rs!ValorAdj
       sCadTotValorAdic = sCadTotTasacion + rs!ValorAdic
       
       rs.MoveNext
       CambiarPagina (7)

    Loop Until x = rs.RecordCount + 1
davi:
       sCad = sCad
       sCadResumen = Chr$(10) & "  Total General: " & Chr$(10) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "CANTIDAD   :" & ImpreFormat(sCadTotPiezas, 4, 0, True) & " piezas" & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "PESO BRUTO :" & ImpreFormat(sCadTotPesoBruto, 13, 2, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "PESO NETO  :" & ImpreFormat(sCadTotPesoNeto, 13, 2, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "TASACION   :" & ImpreFormat(sCadTotTasacion, 13, 2, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "VALOR ADJ  :" & ImpreFormat(sCadTotValorAdj, 13, 2, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "VALOR ADIC :" & ImpreFormat(sCadTotValorAdic, 13, 2, True) & Chr$(10)
       
       sCad = sCad & sCadResumen
       

 End If

 
      'Devuelve toda la estructura  del Reporte
      RepPiezasSelVenta = sCad

End Function
Public Function RepPiezasSelVentaPorRemate() As String

'Variables para el Manejo de losDatos
Dim Co As nCredRepoFinMes
Dim rs As ADODB.Recordset
Dim strString As String                     ' Variable que contiene la consulta

'Negrita
Dim BoldOn As Variant
Dim BoldOff As Variant
BoldOn = Chr$(27) & Chr$(69)
BoldOff = Chr$(27) & Chr$(70)

'Varibles que pertenecen al cuerpo del reporte
Dim sCad As String

' Variables que pasaran a Totalizar
Dim sCadTotPiezas As Variant
Dim sCadTotPBruto As Variant
Dim sCadTotPNeto As Variant
Dim sCadTotTasaccion As Variant
Dim sCadTotValorAdj As Variant
Dim sCadTotValorAdic As Variant

Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer

Dim cTitulo As String

Set Co = New nCredRepoFinMes

strString = " select CPEJ.nremate as 'Remate', count(*) as 'Cantidad', sum(CPJT.nPesoBruto) as 'Peso Bruto', sum(CPJT.nPesoNeto) as 'Peso Neto', sum(CPJT.nTasacion +CPJT.nTasacionAdicional) as 'Tasacion' ,sum(CPP.nValorProceso) as 'ValorAdj' ,sum(CPP.nValorAdicional) as 'ValorAdic'"
strString = strString & " FROM  ColocPigJoyaTasacion CPJT "
strString = strString & " INNER JOIN ColocPigno CP on CP.cCtaCod = CPJT.cCtaCod and CP.nTipoTasacion = CPJT.nTipoTasacion  "
strString = strString & " INNER JOIN ColocPigEstadoJoya  CPEJ on CPJT.cCtaCod =CPEJ.cCtacod and CPJT.nItemPieza = CPEJ.nItemPieza "
strString = strString & " INNER JOIN ColocPigProceso CPP on CPP.nRemate=CPEJ.nRemate and CPP.nTipoProceso = CPEJ.nEstadoJoya and CPP.cCtaCod =CPEJ.cCtacod and CPP.nItemPieza = CPEJ.nItemPieza  "
strString = strString & " WHERE CPP.nSituacionPieza= " & gPigSituacionDisponible & " and CPP.nTipoProceso= " & gPigTipoVentas
strString = strString & " group by CPEJ.nremate "
strString = strString & " order by CPEJ.nremate "

Set rs = Co.GetQuery(strString)

'Inicializar variables
sCad = ""
scadColumnas = ""
'Cabecera para el Reporte
lnLineas = 8
lnPage = 1
lnAncho = 126
sCadTitulo = "CUSTODIA DE VALORES"
sCadSubTitulo = "PIEZAS SELECCIONADAS PARA VENTA(RESUMEN)"
sCadComenta = "Ubicación : Agencia Principal"
scadColumnas = Chr(10) & scadColumnas & String(120, "-") & Chr(10)
scadColumnas = scadColumnas & Space(7) & "REMATE" & Space(8) & "CANTIDAD" & Space(10) & "P. BRUTO" & Space(11) & "P. NETO" & Space(10) & "TASACION" & Space(9) & "VALOR ADJ" & Space(8) & "VALOR ADIC" & Chr(10)
scadColumnas = scadColumnas & String(120, "-") & Chr(10) & Chr(10)
sCad = sCad & nArmaCabeceraReporte(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, sCadTitulo, sCadSubTitulo, lnPage, lnAncho, scadColumnas, sCadComenta)

If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    
    Do
     lnLineas = lnLineas + 1
               
                    sCad = sCad & Space(10) & ImpreFormat(rs!Remate, 3, 0, False) & Space(3)
                sCad = sCad & ImpreFormat(rs("Cantidad"), 13, 0, True) & Space(3)
                sCad = sCad & ImpreFormat(rs("Peso Bruto"), 13, 2, True) & Space(2)
                sCad = sCad & ImpreFormat(rs("Peso Neto"), 13, 2, True) & Space(2)
                sCad = sCad & ImpreFormat(rs!Tasacion, 13, 2, True) & Space(2)
                sCad = sCad & ImpreFormat(rs!ValorAdj, 13, 2, True) & Space(2)
                sCad = sCad & ImpreFormat(rs!ValorAdic, 13, 2, True) & Chr$(10)
               
               'Suma para Totalizar
                sCadTotPiezas = sCadTotPiezas + rs("Cantidad")
                sCadTotPBruto = sCadTotPBruto + rs("Peso Bruto")
                sCadTotPNeto = sCadTotPNeto + rs("Peso Neto")
                sCadTotTasacion = sCadTotTasacion + rs!Tasacion
                sCadTotValorAdj = sCadTotValorAdj + rs!ValorAdj
                sCadTotValorAdic = sCadTotValorAdic + rs!ValorAdic
            
                CambiarPagina (1)
                rs.MoveNext
          
       Loop Until rs.EOF
                         
             ' sCadTot1 = sCadTot1 & String(120, "-") & Chr(10)
              sCadTot1 = Space(7) & "TOTALES :" & ImpreFormat(sCadTotPiezas, 13, 0, True) & Space(3) & ImpreFormat(sCadTotPBruto, 13, 2, True) & Space(2)
              sCadTot1 = sCadTot1 & ImpreFormat(sCadTotPNeto, 13, 2, True) & Space(2) & ImpreFormat(sCadTotTasacion, 13, 2, True) & Space(2) & ImpreFormat(sCadTotValorAdj, 13, 2, True) & Space(2) & ImpreFormat(sCadTotValorAdic, 13, 2, True)
    
              sCad = sCad & Space(2) & String(118, "-") & Chr(10) & sCadTot1 & Chr(10)
         
'Devuelve toda la estructura  del Reporte
        
 End If

 RepPiezasSelVentaPorRemate = sCad

End Function

Public Function RepStockLotesTransito() As String


'Variables para el Manejo de losDatos
Dim Co As nCredRepoFinMes
Dim rs As ADODB.Recordset
Dim strString As String                     ' Variable que contiene la consulta


'Negrita
Dim BoldOn As Variant
Dim BoldOff As Variant
BoldOn = Chr$(27) & Chr$(69)
BoldOff = Chr$(27) & Chr$(70)


Dim nItem As Integer

Dim cTitulo As String

Set Co = New nCredRepoFinMes

strString = "  SELECT   Estado=  Case A.nEstadoCont "
strString = strString & " When " & gPigEstAmortiz & " Then 0 "
strString = strString & " When " & gPigEstDesemb & " Then 0 "
strString = strString & " When " & gPigEstReusoLin & " Then 0 Else 1 End, "
strString = strString & "  SUBSTRING(A.cCtaCod,4,2)  AS 'CodAgencia', "
strString = strString & "   C.cAgeDescripcion AS 'Agencia', A.cCtaCod   AS 'Lote',   sum(B.nPesoBruto) AS 'Peso Bruto',   sum(B.nPesoNeto) AS  'Peso Neto', sum(B.nTasacion+B.nTasacionAdicional ) AS  'Tasacion' , d.cUser As Tasador"
strString = strString & " FROM ColocPigno A"
strString = strString & " INNER JOIN  ColocPigJoyaTasacion B ON A.cCtaCod=B.cCtaCod AND A.nTipoTasacion=B.nTipoTasacion "
strString = strString & " INNER JOIN    Agencias C  ON  SUBSTRING(A.cCtaCod,4,2)=C.cAgeCod"
strString = strString & " INNER JOIN    rrhh d  ON  d.cPersCod= b.cPersCod "
strString = strString & " Where A.nUbicaLote = 1"
strString = strString & " group by Case A.nEstadoCont When " & gPigEstAmortiz & " Then 0  When " & gPigEstDesemb & " Then 0 When " & gPigEstReusoLin & " Then 0 Else 1 End,  SUBSTRING(A.cCtaCod,4,2),  C.cAgeDescripcion , A.cCtaCod,  d.cUser "
strString = strString & " ORDER BY  Estado, CodAgencia"

     
Set rs = Co.GetQuery(strString)

'Inicializar variables
sCad = ""
scadColumnas = ""

'Cabecera para el Reporte
lnLineas = 8
lnPage = 1
lnAncho = 127
sCadTitulo = "CREDITO PIGNORATICIO"
sCadSubTitulo = "STOCK GLOBAL DE LOTES EN TRANSITO"
sCadComenta = ""
scadColumnas = Chr(10) & scadColumnas & String(120, "-") & Chr(10)
scadColumnas = scadColumnas & Space(36) & "LOTE " & Space(15) & "P. BRUTO" & Space(10) & "P. NETO" & Space(9) & "TASACION" & Space(9) & "TASADOR" & Chr(10)
scadColumnas = scadColumnas & String(120, "-") & Chr(10) & Chr(10)
sCad = sCad & nArmaCabeceraReporte(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, sCadTitulo, sCadSubTitulo, lnPage, lnAncho, scadColumnas, sCadComenta)


If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    sCadEstado = rs!Estado
    sCadCodAgencia = rs!CodAgencia
    Do
        lnLineas = lnLineas + 1
        x = x + 1
        If rs.EOF Then
        ultimo = 1
        rs.MoveLast
        End If
        If (rs!CodAgencia <> sCadCodAgencia) Or (rs!Estado <> sCadEstado) Or (x = 1) Or (ultimo = 1) Then
           'Totalizo Agencia
          
           sCadTotAgencia = Space(17) & ImpreFormat(sCadTotPiezasAgencia, 4, 0, True) & " piezas"
           sCadTotAgencia = sCadTotAgencia & Space(3) & ImpreFormat(sCadTotPesoBrutoAgencia, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(4) & ImpreFormat(sCadTotPesoNetoAgencia, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(4) & ImpreFormat(sCadTotTasacionAgencia, 10, 2, True)
           
           sCadTotPiezasAgencia = 0
           sCadTotPesoBrutoAgencia = 0
           sCadTotPesoNetoAgencia = 0
           sCadTotTasacionAgencia = 0
                      
           If x <> 1 Then sCad = sCad & Chr$(10) & Space(5) & String(93, "-") & Chr$(10) & Space(5) & " Total Agencia:" & sCadTotAgencia & Chr$(10)
           
           If (rs!Estado <> sCadEstado) Or (x = 1) Or (ultimo = 1) Then
              'Totalizo Estado
               sCadTotEstado = Space(26) & ImpreFormat(sCadTotPiezasEstado, 4, 0, True) & " piezas"
               sCadTotEstado = sCadTotEstado & Space(3) & ImpreFormat(sCadTotPesoBrutoEstado, 10, 2, True)
               sCadTotEstado = sCadTotEstado & Space(4) & ImpreFormat(sCadTotPesoNetoEstado, 10, 2, True)
               sCadTotEstado = sCadTotEstado & Space(4) & ImpreFormat(sCadTotTasacionEstado, 10, 2, True)
           
               sCadTotPiezasEstado = 0
               sCadTotPesoBrutoEstado = 0
               sCadTotPesoNetoEstado = 0
               sCadTotTasacionEstado = 0
               
               If x <> 1 Then sCad = sCad & Chr$(10) & Space(2) & String(96, "-") & Chr$(10) & "  Total :  " & sCadTotEstado & Chr$(10) & Space(2) & String(96, "-") & Chr$(10)
              
              'Encabezo Estado
               sCadEstado = rs!Estado
               If rs!Estado = 0 Then
                  sCadTitEstado = "  VALORES EN GARANTIA (CONTRATOS VIGENTES) "
               Else
                  sCadTitEstado = "  VALORES EN CUSTODIA (RESCATES PENDIENTES) "
               End If
               If ultimo <> 1 Then sCad = sCad & Chr$(10) & sCadTitEstado & Chr$(10)
           End If
           'Encabezo Agencia
           sCadCodAgencia = rs!CodAgencia
           sCadTitAgencia = Space(5) & "(" & ImpreFormat(rs!CodAgencia, 2, 0, True) & ")" & Space(1) & ImpreFormat(rs!Agencia, 25)
        
           If ultimo <> 1 Then sCad = sCad & Chr$(10) & sCadTitAgencia & Chr$(10) & Chr$(10)
           End If
       'Añado Detalle
       If ultimo = 1 Then GoTo davi
            SCADDETALLE = Space(30) & SCADDETALLE & ImpreFormat(rs("Lote"), 18, 0, False) & Space(3)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs("Peso Bruto"), 10, 2, True) & Space(4)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs("Peso Neto"), 10, 2, True) & Space(4)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Tasacion, 10, 2, True) & Space(10)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Tasador, 10, 2, True) & Chr$(10)
    
       sCad = sCad & SCADDETALLE
       
       SCADDETALLE = ""
       '************Segundo Quiebre****************
       'Suma para Totalizar por Agencia
       sCadTotPiezasAgencia = sCadTotPiezasAgencia + 1
       sCadTotPesoBrutoAgencia = sCadTotPesoBrutoAgencia + rs("Peso Bruto")
       sCadTotPesoNetoAgencia = sCadTotPesoNetoAgencia + rs("Peso Neto")
       sCadTotTasacionAgencia = sCadTotTasacionAgencia + rs!Tasacion
       
       
       '************Primer Quiebre****************
      'Suma para Totalizar por Estado
       sCadTotPiezasEstado = sCadTotPiezasEstado + 1
       sCadTotPesoBrutoEstado = sCadTotPesoBrutoEstado + rs("Peso Bruto")
       sCadTotPesoNetoEstado = sCadTotPesoNetoEstado + rs("Peso Neto")
       sCadTotTasacionEstado = sCadTotTasacionEstado + rs!Tasacion
       
       'Suma para Totalizar por Estado
       sCadTotPiezas = sCadTotPiezas + 1
       sCadTotPesoBruto = sCadTotPesoBruto + rs("Peso Bruto")
       sCadTotPesoNeto = sCadTotPesoNeto + rs("Peso Neto")
       sCadTotTasacion = sCadTotTasacion + rs!Tasacion
              
       rs.MoveNext
       CambiarPagina (7)

    Loop Until x = rs.RecordCount + 1
davi:
       sCad = sCad
       sCadResumen = Chr$(10) & "  Total General: " & Chr$(10) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "CANTIDAD   :" & ImpreFormat(sCadTotPiezas, 4, 0, True) & " piezas" & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "PESO BRUTO :" & ImpreFormat(sCadTotPesoBruto, 13, 2, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "PESO NETO  :" & ImpreFormat(sCadTotPesoNeto, 13, 2, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "TASACION   :" & ImpreFormat(sCadTotTasacion, 13, 2, True) & Chr$(10)
       
       sCad = sCad & sCadResumen
       
End If
       RepStockLotesTransito = sCad

End Function

Public Function RepResumenSaldoAgencia() As String

'Variables para el Manejo de losDatos
Dim Co As nCredRepoFinMes
Dim rs As ADODB.Recordset
Dim strString As String                     ' Variable que contiene la consulta

'Negrita
Dim BoldOn As Variant
Dim BoldOff As Variant
BoldOn = Chr$(27) & Chr$(69)
BoldOff = Chr$(27) & Chr$(70)

'Varibles que pertenecen al cuerpo del reporte
Dim sCad As String

' Variables que pasaran a Totalizar
Dim sCadTotPiezas As Variant
Dim sCadTotPBruto As Variant
Dim sCadTotPNeto As Variant
Dim sCadTotTasaccion As Variant
Dim sCadTotValorAdj As Variant
Dim sCadTotValorAdic As Variant

Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer

Dim cTitulo As String

Set Co = New nCredRepoFinMes
strString = " select (x.agencia +' ' +  z.cAgeDescripcion) descAgencia,x.contratos,x.saldo,y.tasacion,y.PesoNeto from"
strString = strString & " (select substring(a.cCtaCod,4,2) Agencia,count(*) Contratos,sum(a.nsaldo) Saldo"
strString = strString & " from producto a inner join colocpigno b on b.cctacod=a.cctacod"
strString = strString & " where nprdestado  in (" & gPigEstAmortiz & "," & gPigEstCancelPendRes & "," & gPigEstDesemb & "," & gPigEstReusoLin & ")"
strString = strString & " group by substring(a.cCtaCod,4,2)) x "
strString = strString & " inner join (select substring(a.cCtaCod,4,2) Agencia,sum(c.ntasacion)Tasacion,sum(c.nPesoNeto) 'PesoNeto'"
strString = strString & " from producto a "
strString = strString & " inner join colocpigno b on b.cctacod=a.cctacod"
strString = strString & " inner join colocpigjoyatasacion c on c.cctacod=b.cctacod and c.ntipotasacion = b.ntipotasacion"
strString = strString & " where nprdestado  in (" & gPigEstAmortiz & "," & gPigEstCancelPendRes & "," & gPigEstDesemb & "," & gPigEstReusoLin & ")"
strString = strString & " group by substring(a.cCtaCod,4,2)) y on y.agencia=x.agencia "
strString = strString & "  inner join agencias z on z.cAgeCod=x.agencia order by x.agencia"

Set rs = Co.GetQuery(strString)

'Inicializar variables
sCad = ""
scadColumnas = ""
'Cabecera para el Reporte
lnLineas = 8
lnPage = 1
lnAncho = 127
sCadTitulo = "SISTEMA PIGNORATICIO"
sCadSubTitulo = "RESUMEN DE SALDOS POR AGENCIA"
sCadComenta = ""
scadColumnas = Chr(10) & scadColumnas & String(120, "-") & Chr(10)
scadColumnas = scadColumnas & Space(10) & "AGENCIA" & Space(23) & "CONTRATOS" & Space(14) & "SALDO" & Space(10) & "TASACION" & Space(10) & "P. NETO" & Chr(10)
scadColumnas = scadColumnas & String(120, "-") & Chr(10) & Chr(10)
sCad = sCad & nArmaCabeceraReporte(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, sCadTitulo, sCadSubTitulo, lnPage, lnAncho, scadColumnas, sCadComenta)

If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    
    Do
     lnLineas = lnLineas + 1
               
                    sCad = sCad & Space(10) & ImpreFormat(rs!descAgencia, 23, 0, False) & Space(3)
                sCad = sCad & ImpreFormat(rs("contratos"), 13, 0, True) & Space(3)
                sCad = sCad & ImpreFormat(rs("saldo"), 13, 2, True) & Space(2)
                sCad = sCad & ImpreFormat(rs("tasacion"), 13, 2, True) & Space(2)
                sCad = sCad & ImpreFormat(rs!PesoNeto, 13, 2, True) & Chr$(10)
                
               'Suma para Totalizar
                sCadTotContratos = sCadTotContratos + rs("Contratos")
                sCadTotSaldo = sCadTotSaldo + rs("Saldo")
                sCadTotTasacion = sCadTotTasacion + rs("Tasacion")
                sCadTotPesoNeto = sCadTotPesoNeto + rs!PesoNeto
              
            
                CambiarPagina (1)
                rs.MoveNext
          
       Loop Until rs.EOF
                         
             ' sCadTot1 = sCadTot1 & String(120, "-") & Chr(10)
              sCadTot1 = Space(7) & "TOTALES :" & Space(20) & ImpreFormat(sCadTotContratos, 13, 0, True) & Space(3) & ImpreFormat(sCadTotSaldo, 13, 2, True) & Space(2)
              sCadTot1 = sCadTot1 & ImpreFormat(sCadTotTasacion, 13, 2, True) & Space(2) & ImpreFormat(sCadTotPesoNeto, 13, 2, True)
              sCad = sCad & Space(2) & String(118, "-") & Chr(10) & sCadTot1 & Chr(10)
         
'Devuelve toda la estructura  del Reporte
        
 End If

 RepResumenSaldoAgencia = sCad

End Function
Public Function RepPolizasEmitidas(psFecini, psFecFin, psRemate) As String


'Variables para el Manejo de losDatos
Dim Co As nCredRepoFinMes
Dim rs As ADODB.Recordset
Dim strString As String                     ' Variable que contiene la consulta


'Negrita
Dim BoldOn As Variant
Dim BoldOff As Variant
BoldOn = Chr$(27) & Chr$(69)
BoldOff = Chr$(27) & Chr$(70)

Dim nItem As Integer

Dim cTitulo As String

Set Co = New nCredRepoFinMes

strString = " select p.cPersNombre as 'Martillero',(substring(a.cUltimaActualizacion,7,2) + '/' + substring(a.cUltimaActualizacion,5,2) + '/' +substring(a.cUltimaActualizacion,1,4)) as 'Fecha',a.nCodTipo as 'Tipo'"
strString = strString & " , a.cDocumento as 'Documento',a.nNroMov as 'Movimiento' , q.cPersNombre as 'Persona',max(b.nItemDoc) as 'Cantidad'"
strString = strString & " ,sum(case  b.nCodConcepto  when " & gColPigConceptoCodPrecioVenta & " then  b.nMonto else 0 end) as 'ValorVenta'"
strString = strString & "  ,sum(case  b.nCodConcepto  when " & gColPigConceptoCodComisionMartillero & " then  b.nMonto else 0 end) as 'ComisionMartillero'"
strString = strString & " ,sum(case  b.nCodConcepto  when " & gColPigConceptoCodIgv & " then  b.nMonto else 0 end) as 'I.G.V.'"
strString = strString & " ,sum( b.nMonto) as 'Monto'"
strString = strString & "  from movdocpig a"
strString = strString & "  inner join movdocpigdet b on b.nCodTipo =a.nCodTipo and b.cDocumento=a.cDocumento and b.nNroMov = a.nNroMov"
strString = strString & "  inner join (select distinct y.nCodTipo,y.cDocumento,y.nNroMov,x.nRemate"
strString = strString & "  from colocpigproceso x"
strString = strString & "  inner join movdocpigdet y on y.cCtaCod=x.cCtaCod and y.nPieza = x.nItemPieza"
strString = strString & "  where nSituacionPieza=3 and nRemate= " & psRemate & ") c on c.nCodTipo = a.nCodTipo and c.cDocumento = a.cDocumento and c.nNroMov = a.nNroMov"
strString = strString & "  inner join colocpigremate r on r.nRemate= c.nRemate"
strString = strString & "  left  join persona p on p.cPersCod = r.cPersCod"
strString = strString & "  inner join persona q on q.cPersCod = a.cPersCod"
strString = strString & "  where ncodConcepto in (" & gColPigConceptoCodIgv & "," & gColPigConceptoCodPrecioVenta & "," & gColPigConceptoCodComisionMartillero & ") and a.nCodTipo=" & gPigTipoPoliza
strString = strString & "  and  substring(a.cUltimaActualizacion,1,8) between '" & Format(psFecini, gsFormatoMovFecha) & "' and '" & Format(psFecFin, gsFormatoMovFecha) & "'"
strString = strString & "  group by p.cPersNombre, a.cUltimaActualizacion,a.nCodTipo, a.cDocumento,a.nNroMov, q.cPersNombre"

Set rs = Co.GetQuery(strString)

'Inicializar variables
sCad = ""
scadColumnas = ""
'scad = scad & Chr$(15)                           'Retorna al tipo de letra normal
'scad = scad & Chr$(27) + Chr$(107) + Chr$(1)     'Tipo de Letra Roman
'scad = scad & Chr$(27) + Chr$(103)                'Tamaño 10.5 - 15 CPI
'scad = scad & Chr$(27) + Chr$(50)                'Espaciado entre lineas 1/16
''sCad = sCad & Chr$(27) + Chr$(67) + Chr$(18)     'Longitud de página a 24 líneas
'scad = scad & Chr$(27) + Chr$(108) + Chr$(6)     'Margen Izquierdo - 6ta. Columna


'   --- Constantes para Impresión  ------
' Introduce espacio                    :   Space(43)
' Enter para la siguiente linea   :   Chr$(10)
' Determina la cantidad de
' espacios devueltos                 :    ImpreFormat("campo", Cantidad de espacio, Decimales,¿moneda?)

'Cabecera para el Reporte
lnLineas = 8
lnPage = 1
lnAncho = 133
sCadTitulo = "VENTAS Y REMATES"
sCadSubTitulo = "POLIZAS EMITIDAS"
If Not (rs.EOF And rs.BOF) Then
    sCadComenta = "MARTILLERO: " & Justifica(rs("Persona"), 40) & Space(3) & "REMATE Nro. " & psRemate & Chr(10)
    sCadComenta = sCadComenta & Space(7) & "DEL: " & psFecini & " AL: " & psFecFin
Else
    sCadComenta = ""
End If
scadColumnas = scadColumnas & String(127, "-") & Chr(10)
scadColumnas = scadColumnas & Space(2) & "FECHA DE " & Space(1) & "DOCUMENTO" & Space(3) & "NOMBRE DEL" & Space(39) & "VALOR " & Space(3) & "  COMISION " & Space(3) & "I.G.V. COM" & Space(9) & "TOTAL" & Space(3) & "CANT" & Chr(10)
scadColumnas = scadColumnas & Space(2) & "EMISION   " & Space(12) & "ADQUIRENTE" & Space(39) & "VENTA" & Space(4) & "MARTILLERO" & Space(4) & "MARTILLERO  " & Space(6) & " VENTA" & Space(3) & "PZAS" & Chr(10)

scadColumnas = scadColumnas & String(127, "-") & Chr(10)
sCad = sCad & nArmaCabeceraReporte(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, sCadTitulo, sCadSubTitulo, lnPage, lnAncho, scadColumnas, sCadComenta)



If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    sCadFecha = rs!Fecha
    Do
        lnLineas = lnLineas + 1
        x = x + 1
        If rs.EOF Then
        ultimo = 1
        rs.MoveLast
        End If
        If (rs!Fecha <> sCadFecha) Or (x = 1) Or (ultimo = 1) Then
              'Totalizo Fecha
               sCadTotFecha = Space(8) & ImpreFormat(sCadTotDocumentoFecha, 4, 0, True) & " documentos"
               sCadTotFecha = sCadTotFecha & Space(42) & ImpreFormat(sCadTotValorVentaFecha, 10, 2, True)
               sCadTotFecha = sCadTotFecha & Space(1) & ImpreFormat(sCadTotComisionMartilleroFecha, 10, 2, True)
               sCadTotFecha = sCadTotFecha & Space(1) & ImpreFormat(sCadTotIGVFecha, 10, 2, True)
               sCadTotFecha = sCadTotFecha & Space(1) & ImpreFormat(sCadTotMontoFecha, 10, 2, True)
               sCadTotFecha = sCadTotFecha & Space(3) & ImpreFormat(sCadTotCantidadFecha, 4, 0, True)
           
               sCadTotDocumentoFecha = 0
               sCadTotValorVentaFecha = 0
               sCadTotComisionMartilleroFecha = 0
               sCadTotIGVFecha = 0
               sCadTotMontoFecha = 0
               sCadTotCantidadFecha = 0
               
               If x <> 1 Then sCad = sCad & Chr$(10) & sCadTotFecha & Chr$(10) & Chr$(10)
              
              'Encabezo Remate
               sCadFecha = rs!Fecha
               sCadTitFecha = " " & ImpreFormat(sCadFecha, 10, 0, True)
           
               If ultimo <> 1 Then sCad = sCad & BoldOn & sCadTitFecha & BoldOff & Space(1)
           End If
       'Añado Detalle
       If ultimo = 1 Then GoTo davi
       SCADDETALLE = SCADDETALLE & Justifica(rs!Documento, 11) & Space(1)
       SCADDETALLE = SCADDETALLE & Justifica(rs("Persona"), 40) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs("ValorVenta"), 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs("ComisionMartillero"), 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs("I.G.V."), 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Monto, 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!cantidad, 6, 0, True) & Chr$(10) & Space(12)
    
       sCad = sCad & SCADDETALLE
       SCADDETALLE = ""
       '************Primer Quiebre****************
       'Suma para Totalizar por Remate
       sCadTotDocumentoFecha = sCadTotDocumentoFecha + 1
       sCadTotValorVentaFecha = sCadTotValorVentaFecha + rs("ValorVenta")
       sCadTotComisionMartilleroFecha = sCadTotComisionMartilleroFecha + rs("ComisionMartillero")
       sCadTotIGVFecha = sCadTotIGVFecha + rs("I.G.V.")
       sCadTotMontoFecha = sCadTotMontoFecha + rs!Monto
       sCadTotCantidadFecha = sCadTotCantidadFecha + rs!cantidad
       
       'Suma para Totalizar por Remate
       sCadTotDocumento = sCadTotDocumento + 1
       sCadTotValorVenta = sCadTotValorVenta + rs("ValorVenta")
       sCadTotComisionMartillero = sCadTotComisionMartillero + rs("ComisionMartillero")
       sCadTotIGV = sCadTotIGV + rs("I.G.V.")
       sCadTotMonto = sCadTotMonto + rs!Monto
       sCadTotCantidad = sCadTotCantidad + rs!cantidad
       
       rs.MoveNext
       CambiarPagina (7)

    Loop Until x = rs.RecordCount + 1
davi:
       sCad = sCad
       sCadResumen = Chr$(10) & "  Total General: " & Chr$(10) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "DOCUMENTOS" & Space(17) & " :" & ImpreFormat(sCadTotDocumento, 16, 0, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "VALOR VENTA" & Space(16) & " :" & ImpreFormat(sCadTotValorVenta, 13, 2, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "COMISION MARTILLERO" & Space(8) & " :" & ImpreFormat(sCadTotComisionMartillero, 13, 2, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "I.G.V. COMISION MARTILLERO" & Space(1) & " :" & ImpreFormat(sCadTotIGV, 13, 2, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "TOTAL VENTA" & Space(16) & " :" & ImpreFormat(sCadTotMonto, 13, 2, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "CANTIDAD PIEZAS " & Space(11) & " :" & ImpreFormat(sCadTotCantidad, 16, 0, True) & Chr$(10)
       
       sCad = sCad & sCadResumen
       

 End If

 
      'Devuelve toda la estructura  del Reporte
      RepPolizasEmitidas = sCad

End Function

Public Function RepPolizasEmitidasResumen(psFecini, psFecFin, psRemate) As String


'Variables para el Manejo de losDatos
Dim Co As nCredRepoFinMes
Dim rs As ADODB.Recordset
Dim strString As String                     ' Variable que contiene la consulta


'Negrita
Dim BoldOn As Variant
Dim BoldOff As Variant
BoldOn = Chr$(27) & Chr$(69)
BoldOff = Chr$(27) & Chr$(70)



Dim nItem As Integer

Dim cTitulo As String

Set Co = New nCredRepoFinMes


strString = " select (substring(a.cUltimaActualizacion,7,2) + '/' + substring(a.cUltimaActualizacion,5,2) + '/' +substring(a.cUltimaActualizacion,1,4)) as 'Fecha'"
strString = strString & " ,count(distinct b.cDocumento) CantPolizas,count(*)/3 CantPiezas"
strString = strString & " ,sum(case  b.nCodConcepto  when " & gColPigConceptoCodPrecioVenta & " then  b.nMonto else 0 end) as 'ValorVenta'"
strString = strString & "  ,sum(case  b.nCodConcepto  when " & gColPigConceptoCodComisionMartillero & " then  b.nMonto else 0 end) as 'ComisionMartillero'"
strString = strString & " ,sum(case  b.nCodConcepto  when " & gColPigConceptoCodIgv & " then  b.nMonto else 0 end) as 'I.G.V.'"
strString = strString & " ,sum( b.nMonto) as 'Monto'"
strString = strString & "  from movdocpig a"
strString = strString & "  inner join movdocpigdet b on b.nCodTipo =a.nCodTipo and b.cDocumento=a.cDocumento and b.nNroMov = a.nNroMov"
strString = strString & " inner join (select distinct y.nCodTipo,y.cDocumento,y.nNroMov,x.nRemate"
strString = strString & " from colocpigproceso x"
strString = strString & "  inner join movdocpigdet y on y.cCtaCod=x.cCtaCod and y.nPieza = x.nItemPieza"
strString = strString & "  where nSituacionPieza=3 and nRemate= " & psRemate & ") c on c.nCodTipo = a.nCodTipo and c.cDocumento = a.cDocumento and c.nNroMov = a.nNroMov"
strString = strString & "  inner join colocpigremate r on r.nRemate= c.nRemate"
strString = strString & "  where ncodConcepto in (" & gColPigConceptoCodIgv & "," & gColPigConceptoCodPrecioVenta & "," & gColPigConceptoCodComisionMartillero & ") and a.nCodTipo=" & gPigTipoPoliza
strString = strString & "  and  substring(a.cUltimaActualizacion,1,8) between '" & Format(psFecini, gsFormatoMovFecha) & "' and '" & Format(psFecFin, gsFormatoMovFecha) & "'"
strString = strString & " group by (substring(a.cUltimaActualizacion,7,2) + '/' + substring(a.cUltimaActualizacion,5,2) + '/' +substring(a.cUltimaActualizacion,1,4)) "


Set rs = Co.GetQuery(strString)

'Inicializar variables
sCad = ""
scadColumnas = ""
'scad = scad & Chr$(15)                           'Retorna al tipo de letra normal
'scad = scad & Chr$(27) + Chr$(107) + Chr$(1)     'Tipo de Letra Roman
'scad = scad & Chr$(27) + Chr$(103)                'Tamaño 10.5 - 15 CPI
'scad = scad & Chr$(27) + Chr$(50)                'Espaciado entre lineas 1/16
''sCad = sCad & Chr$(27) + Chr$(67) + Chr$(18)     'Longitud de página a 24 líneas
'scad = scad & Chr$(27) + Chr$(108) + Chr$(6)     'Margen Izquierdo - 6ta. Columna


'   --- Constantes para Impresión  ------
' Introduce espacio                    :   Space(43)
' Enter para la siguiente linea   :   Chr$(10)
' Determina la cantidad de
' espacios devueltos                 :    ImpreFormat("campo", Cantidad de espacio, Decimales,¿moneda?)

'Cabecera para el Reporte
lnLineas = 8
lnPage = 1
lnAncho = 118
sCadTitulo = "VENTAS Y REMATES"
sCadSubTitulo = "POLIZAS EMITIDAS-RESUMEN"
If Not (rs.EOF And rs.BOF) Then
    sCadComenta = sCadComenta & Space(7) & "DEL: " & psFecini & " AL: " & psFecFin
Else
    sCadComenta = ""
End If
scadColumnas = scadColumnas & String(112, "-") & Chr(10)
scadColumnas = scadColumnas & Space(2) & "FECHA DE EMISION" & Space(3) & "CANT.POLIZAS " & Space(3) & "CANT.PIEZAS " & Space(10) & "VALOR " & Space(5) & "  COMISION " & Space(5) & "I.G.V. COM" & Space(9) & "TOTAL" & Chr(10)
scadColumnas = scadColumnas & String(112, "-") & Chr(10)
sCad = sCad & nArmaCabeceraReporte(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, sCadTitulo, sCadSubTitulo, lnPage, lnAncho, scadColumnas, sCadComenta)


If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    
    Do
     lnLineas = lnLineas + 1
               
             sCad = sCad & Space(8) & ImpreFormat(rs("Fecha"), 10, 0, True) & Space(11)
             sCad = sCad & ImpreFormat(rs("CantPolizas"), 4, 0, True) & Space(11)
             sCad = sCad & ImpreFormat(rs("CantPiezas"), 4, 0, True) & Space(3)
             sCad = sCad & ImpreFormat(rs("ValorVenta"), 10, 2, True) & Space(3)
             sCad = sCad & ImpreFormat(rs("ComisionMartillero"), 10, 2, True) & Space(3)
             sCad = sCad & ImpreFormat(rs("I.G.V."), 10, 2, True) & Space(1)
             sCad = sCad & ImpreFormat(rs!Monto, 10, 2, True) & Chr$(10)
    
                
               'Suma para Totalizar
             sCadTotCantidadPolizas = sCadTotCantidadPolizas + rs("CantPolizas")
             sCadTotCantidadPiezas = sCadTotCantidadPiezas + rs("CantPiezas")
             sCadTotValorVentaFecha = sCadTotValorVentaFecha + rs("ValorVenta")
             sCadTotComisionMartilleroFecha = sCadTotComisionMartilleroFecha + rs("ComisionMartillero")
             sCadTotIGVFecha = sCadTotIGVFecha + rs("I.G.V.")
             sCadTotMontoFecha = sCadTotMontoFecha + rs!Monto
              
            
                CambiarPagina (1)
                rs.MoveNext
          
       Loop Until rs.EOF
                         
             ' sCadTot1 = sCadTot1 & String(120, "-") & Chr(10)
              sCadTot1 = Space(7) & "TOTALES :" & Space(13) & ImpreFormat(sCadTotCantidadPolizas, 4, 0, True) & Space(11) & ImpreFormat(sCadTotCantidadPiezas, 4, 0, True) & Space(3)
              sCadTot1 = sCadTot1 & ImpreFormat(sCadTotValorVentaFecha, 10, 2, True) & Space(3) & ImpreFormat(sCadTotComisionMartilleroFecha, 10, 2, True) & Space(3) & ImpreFormat(sCadTotIGVFecha, 10, 2, True) & Space(1) & ImpreFormat(sCadTotMontoFecha, 10, 2, True)
              sCad = sCad & Space(2) & String(110, "-") & Chr(10) & sCadTot1 & Chr(10)
         
'Devuelve toda la estructura  del Reporte
        
 End If

 RepPolizasEmitidasResumen = sCad


End Function




Public Function RepInventarioJoyas(psFecini) As String

'Variables para el Manejo de losDatos
'Dim ocon As Conecta
Dim Co As nCredRepoFinMes
Dim Rsc As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim strString As String                     ' Variable que contiene la consulta
 
'Negrita
Dim BoldOn As Variant
Dim BoldOff As Variant
BoldOn = Chr$(27) & Chr$(69)
BoldOff = Chr$(27) & Chr$(70)

'Varibles que pertenecen al cuerpo del reporte


Dim sCadGarantia As String
Dim sCadCustodia As String

' Variables que pasaran a Totalizar
Dim dTotCadcapital As Double
Dim dTotCadotros As Double
Dim dTotCadinterescomp As Double
Dim dTotCadMora As Double
Dim dTotCadPreparaRemate As Double
Dim dTotTotalDeuda As Double

'' variables de reporte
Dim sCadAgencia As String
Dim sCadcCtaCod As String
Dim sCaddvigencia As String
Dim sCadcpersnombre As String * 60
Dim sCadnpiezas As String
Dim sCadnPesoNeto As String
Dim sCadntasacion As String
Dim dCadcapital As Double
Dim dCadotros As Double
Dim dCadinterescomp As Double
Dim dCadMora As Double
Dim dCadPreparaRemate As Double
Dim dTotDeuda As Double
Dim sSql As String

'Dim scad As String
Dim sCadCabezera As String
Dim sCadCabezeraHeader As String
'Dim sCadTitulo As String
'Dim sCadSubTitulo As String
'Dim lnAncho As Integer

Dim sCadCodAgencia As String
Dim x As Integer
Dim ultimo As Integer
Dim sCadTotPiezasAgencia As Integer
Dim sCadTotAgencia  As String
Dim sCadTitAgencia  As String
Dim SCADDETALLE As String
Dim sCadResumen As String



Dim rsE As New ADODB.Recordset
Dim oRep As DpigReportes

Set oRep = New DpigReportes
Set rs = oRep.RepInventarioJoyas(psFecini)
    


'Inicializar variables
sCad = ""
sCadComenta = ""
sCad = sCad & Chr$(15)                           'Retorna al tipo de letra normal
sCad = sCad & Chr$(27) + Chr$(108) + Chr$(0)     'Tipo de Letra Roman
sCad = sCad & Chr$(27) + Chr$(77)                'Tamaño 10.5 - 15 CPI
sCad = sCad & Chr$(27) + Chr$(15)                'Espaciado entre lineas 1/16
'sCad = sCad & Chr$(27) + Chr$(67) + Chr$(18)     'Longitud de página a 24 líneas
sCad = sCad & Chr$(27) + Chr$(108) + Chr$(6)     'Margen Izquierdo - 6ta. Columna


'   --- Constantes para Impresión  ------
' Introduce espacio                    :   Space(43)
' Enter para la siguiente linea   :   Chr$(10)
' Determina la cantidad de
' espacios devueltos                 :    ImpreFormat("campo", Cantidad de espacio, Decimales,¿moneda?)

'Cabecera para el Reporte
lnLineas = 6
lnPage = 1
sCadTitulo = "CUSTODIA DE VALORES"
sCadSubTitulo = "INVENTARIO DE JOYAS AL " & Format(psFecini, gcFormatoFechaView)
lnAncho = 227

sCadComenta = ""
scadColumnas = Chr(10) & scadColumnas & String(227, "-") & Chr(10)
scadColumnas = scadColumnas & Space(10) & "Contrato " & Space(8) & "Fecha" & Space(4) & "Nombre" & Space(50) & "Piezas " & Space(7) & "P. Neto" & Space(8) & "Tasacion" & Space(9) & "Capital  " & Space(9) & "Otros" & Space(9) & "Interes" & Space(10) & "Mora" & Space(10) & "Der. Remate " & Space(2) & " Tot Deuda" & Chr(10)
scadColumnas = scadColumnas & String(227, "-")
sCad = sCad & nArmaCabeceraReporte(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, sCadTitulo, sCadSubTitulo, lnPage, lnAncho, scadColumnas, sCadComenta)

 
 
If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    sCadAgencia = rs!Agencia
    Do
        lnLineas = lnLineas + 1
        x = x + 1
        If rs.EOF Then
        ultimo = 1
        rs.MoveLast
        End If
        ''If (Rs!Agencia <> sCadCodAgencia) Or (Rs!remate <> sCadRemate) Or (X = 1) Or (ultimo = 1) Then
        If (rs!Agencia <> sCadCodAgencia) Or (x = 1) Or (ultimo = 1) Then
           'Totalizo Agencia
           Dim vcontrato As String
            If sCadTotPiezasAgencia > 1 Then
            vcontrato = "   Contratos"
            Else
            vcontrato = "   Contrato"
            End If
           sCadTotAgencia = Space(2) & ImpreFormat(sCadTotPiezasAgencia, 4, 0, True) & vcontrato
           sCadTotAgencia = sCadTotAgencia & Space(103) & ImpreFormat(dTotCadcapital, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(3) & ImpreFormat(dTotCadotros, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(3) & ImpreFormat(dTotCadinterescomp, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(3) & ImpreFormat(dTotCadMora, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(3) & ImpreFormat(dTotCadPreparaRemate, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(3) & ImpreFormat(dTotTotalDeuda, 10, 2, True)
           sCadTotPiezasAgencia = 0
           dTotCadcapital = 0
           dTotCadotros = 0
           dTotCadinterescomp = 0
           dTotCadMora = 0
           dTotCadPreparaRemate = 0
           dTotTotalDeuda = 0

        
           If x <> 1 Then sCad = sCad & Chr$(10) & Space(5) & " Total " & sCadTotAgencia & Chr$(10)
           
           'Encabezo Agencia
           sCadCodAgencia = rs!Agencia
           sCadTitAgencia = Space(5) & "(" & ImpreFormat(rs!Agencia, 2, 0, True) & ")" & Space(1) & ImpreFormat(rs!cAgeDescripcion, 25)
        
           If ultimo <> 1 Then sCad = sCad & Chr$(10) & sCadTitAgencia & Chr$(10) & Chr$(10)
           End If
       'Añado Detalle
        If ultimo = 1 Then GoTo PieCabecera
       If ultimo = 1 Then GoTo PieCabecera
       dTotDeuda = IIf(IsNull(rs!Capital), 0, rs!Capital) + IIf(IsNull(rs!otros), 0, rs!otros) + IIf(IsNull(rs!Interescomp), 0, rs!Interescomp) + IIf(IsNull(rs!Mora), 0, rs!Mora) + IIf(IsNull(rs!PreparaRemate), 0, rs!PreparaRemate)
       SCADDETALLE = SCADDETALLE & Space(2) & Justifica(rs!cCtaCod, 18) & Space(3)
       SCADDETALLE = SCADDETALLE & ImpreFormat(Format(rs!dVigencia, "dd/mm/yyyy"), 10, 0, False) & Space(3)
       SCADDETALLE = SCADDETALLE & Justifica(rs!cPersNombre, 50) & Space(3)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!npiezas, 8, 0, False) & Space(3)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!nPesoNeto, 10, 2, True) & Space(3)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!nTasacion, 10, 2, True) & Space(3)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Capital, 10, 2, True) & Space(3)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!otros, 10, 2, True) & Space(3)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Interescomp, 10, 2, True) & Space(3)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Mora, 10, 2, True) & Space(3)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!PreparaRemate, 10, 2, True) & Space(3)
       SCADDETALLE = SCADDETALLE & ImpreFormat(dTotDeuda, 10, 2, True) & Chr(10)

       sCad = sCad & SCADDETALLE
       SCADDETALLE = ""
       '************Segundo Quiebre****************
       'Suma para Totalizar por Agencia
       sCadTotPiezasAgencia = sCadTotPiezasAgencia + 1
       dTotCadcapital = dTotCadcapital + IIf(IsNull(rs!Capital), 0, rs!Capital)
       dTotCadotros = dTotCadotros + IIf(IsNull(rs!otros), 0, rs!otros)
       dTotCadinterescomp = dTotCadinterescomp + IIf(IsNull(rs!Interescomp), 0, rs!Interescomp)
       dTotCadMora = dTotCadMora + IIf(IsNull(rs!Mora), 0, rs!Mora)
       dTotCadPreparaRemate = dTotCadPreparaRemate + IIf(IsNull(rs!PreparaRemate), 0, rs!PreparaRemate)
       dTotTotalDeuda = dTotTotalDeuda + dTotDeuda
       

       rs.MoveNext
       CambiarPagina (2)

    Loop Until x = rs.RecordCount + 1
PieCabecera:
       
       sCad = sCad & sCadResumen
      
      'Devuelve toda la estructura  del Reporte
   RepInventarioJoyas = sCad

 End If

End Function

Public Function RepPiezasRematadas(pRemate As Integer) As String


'Variables para el Manejo de losDatos
Dim Co As nCredRepoFinMes
Dim rs As ADODB.Recordset
Dim strString As String                     ' Variable que contiene la consulta


'Negrita
Dim BoldOn As Variant
Dim BoldOff As Variant
BoldOn = Chr$(27) & Chr$(69)
BoldOff = Chr$(27) & Chr$(70)

'Varibles que pertenecen al cuerpo del reporte


Dim sCadGarantia As String
Dim sCadCustodia As String

' Variables que pasaran a Totalizar
Dim sCadTot1 As Variant
Dim sCadTotLotes1 As Variant
Dim sCadTotPBruto1 As Variant
Dim sCadTotPNeto1 As Variant
Dim sCadTotTasaccion1 As Variant
Dim sCadTotTasaccionAdicional1 As Variant

Dim sCadTot2 As Variant
Dim sCadTotLotes2 As Variant
Dim sCadTotPBruto2 As Variant
Dim sCadTotPNeto2 As Variant
Dim sCadTotTasaccion2 As Variant
Dim sCadTotTasaccionAdicional2 As Variant


Dim nItem As Integer

Dim cTitulo As String

Set Co = New nCredRepoFinMes


strString = " Select SUBSTRING(P.cCtaCod,4,2) AS 'CodAgencia',E.cAgeDescripcion AS 'Agencia', p.cCtaCod,p.nItemPieza,t.nPesoNeto,t.nTasacion,d.capital,d.interesVenc,d.Comision, d.capital + d.interesvenc + d.comision as 'TotalDeuda'"
strString = strString & " from colocpigproceso p"
strString = strString & " inner join ( select nRemate,nTipoProceso,cCtaCod,nItemPieza"
strString = strString & " ,sum(case nCodConcepto when " & gColPigConceptoCodCapital & " then nMonto else 0 end ) capital"
strString = strString & " ,sum(case nCodConcepto  when " & gColPigConceptoCodInteresComp & " then nMonto when " & gcolPigConceptoCodInteresMora & "  then nMonto  else 0 end)  interesVenc"
strString = strString & " ,sum(case nCodConcepto  when " & gColPigConceptoCodComiVencida & " then nMonto else 0 end ) Comision"
strString = strString & "  From colocpigprocesodet"
strString = strString & " group by  nRemate,nTipoProceso,cCtaCod,nItemPieza ) d on d.nRemate = p.nRemate and d.nTipoProceso = p.nTipoProceso and d.cCtaCod= p.cCtaCod and  d.nItemPieza = p.nItemPieza"
strString = strString & " inner join colocpigjoyatasacion t on t.cCtaCod =p.cCtaCod and t.nItemPieza = p.nItemPieza"
strString = strString & " inner join colocpigno c on c.cCtaCod = t.cCtaCod and c.nTipoTasacion = t.nTipoTasacion"
strString = strString & " inner join   Agencias e on e.cAgeCod  =   SUBSTRING(p.cCtaCod,4,2) "
strString = strString & " where p.nRemate=" & pRemate & " and p.nTipoProceso in (" & gPigTipoPrimeraSubasta & "," & gPigTipoSegundaSubasta & "," & gPigTipoTerceraSubasta & ") and nSituacionPieza= " & gPigSituacionFacturado
strString = strString & " order by CodAgencia "




Set rs = Co.GetQuery(strString)

'Inicializar variables
sCad = ""
scadColumnas = ""
'scad = scad & Chr$(15)                           'Retorna al tipo de letra normal
'scad = scad & Chr$(27) + Chr$(107) + Chr$(1)     'Tipo de Letra Roman
'scad = scad & Chr$(27) + Chr$(103)                'Tamaño 10.5 - 15 CPI
'scad = scad & Chr$(27) + Chr$(50)                'Espaciado entre lineas 1/16
''sCad = sCad & Chr$(27) + Chr$(67) + Chr$(18)     'Longitud de página a 24 líneas
'scad = scad & Chr$(27) + Chr$(108) + Chr$(6)     'Margen Izquierdo - 6ta. Columna


'   --- Constantes para Impresión  ------
' Introduce espacio                    :   Space(43)
' Enter para la siguiente linea   :   Chr$(10)
' Determina la cantidad de
' espacios devueltos                 :    ImpreFormat("campo", Cantidad de espacio, Decimales,¿moneda?)

'Cabecera para el Reporte
lnLineas = 8
lnPage = 1
lnAncho = 127
sCadTitulo = "PIEZAS REMATADAS "
sCadSubTitulo = "REMATE: " & pRemate
sCadComenta = ""
scadColumnas = Chr(10) & scadColumnas & String(127, "-") & Chr(10)
scadColumnas = scadColumnas & Space(20) & "PIEZA " & Space(12) & "Peso Neto" & Space(6) & "Tasacion" & Space(7) & "Capital" & Space(3) & "Interes Venc." & Space(2) & "Comisiones" & Space(4) & "Total Deuda" & Chr(10)
scadColumnas = scadColumnas & String(127, "-") & Chr(10) & Chr(10)
sCad = sCad & nArmaCabeceraReporte(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, sCadTitulo, sCadSubTitulo, lnPage, lnAncho, scadColumnas, sCadComenta)


If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    sCadCodAgencia = rs!CodAgencia
    Do
        lnLineas = lnLineas + 1
        x = x + 1
        If rs.EOF Then
        ultimo = 1
        rs.MoveLast
        End If
        If (rs!CodAgencia <> sCadCodAgencia) Or (x = 1) Or (ultimo = 1) Then
           'Totalizo Agencia
           sCadTotAgencia = Space(2) & ImpreFormat(sCadTotPiezasAgencia, 4, 0, True) & " piezas"
            sCadTotAgencia = sCadTotAgencia & Space(1) & ImpreFormat(sCadTotPesoNetoAgencia, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(1) & ImpreFormat(sCadTotTasacionAgencia, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(1) & ImpreFormat(sCadTotCapitalAgencia, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(1) & ImpreFormat(sCadTotInteresVencAgencia, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(1) & ImpreFormat(sCadTotComisionAgencia, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(1) & ImpreFormat(sCadTotTotalDeudaAgencia, 10, 2, True)
           sCadTotPiezasAgencia = 0
           sCadTotPesoNetoAgencia = 0
           sCadTotTasacionAgencia = 0
           sCadTotCapitalAgencia = 0
           sCadTotInteresVencAgencia = 0
           sCadTotComisionAgencia = 0
           sCadTotTotalDeudaAgencia = 0
                      
           If x <> 1 Then sCad = sCad & Chr$(10) & Space(5) & " Total Agencia:" & sCadTotAgencia & Chr$(10)

           'Encabezo Agencia
           sCadCodAgencia = rs!CodAgencia
           sCadTitAgencia = Space(5) & "(" & ImpreFormat(rs!CodAgencia, 2, 0, True) & ")" & Space(1) & ImpreFormat(rs!Agencia, 25)
        
           If ultimo <> 1 Then sCad = sCad & Chr$(10) & sCadTitAgencia & Chr$(10) & Chr$(10)
           End If
       'Añado Detalle
       If ultimo = 1 Then GoTo davi
       SCADDETALLE = SCADDETALLE & Space(12) & Justifica(rs!cCtaCod, 18) & " -"
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!nItemPieza, 2, 0, False)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!nPesoNeto, 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!nTasacion, 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Capital, 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!InteresVenc, 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Comision, 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!TotalDeuda, 10, 2, True) & Chr$(10)
    
       sCad = sCad & SCADDETALLE
       SCADDETALLE = ""
       '************Segundo Quiebre****************
       'Suma para Totalizar por Agencia
       sCadTotPiezasAgencia = sCadTotPiezasAgencia + 1
       sCadTotPesoNetoAgencia = sCadTotPesoNetoAgencia + rs!nPesoNeto
       sCadTotTasacionAgencia = sCadTotTasacionAgencia + rs!nTasacion
       sCadTotCapitalAgencia = sCadTotCapitalAgencia + rs!Capital
       sCadTotInteresVencAgencia = sCadTotInteresVencAgencia + rs!InteresVenc
       sCadTotComisionAgencia = sCadTotComisionAgencia + rs!Comision
       sCadTotTotalDeudaAgencia = sCadTotTotalDeudaAgencia + rs!TotalDeuda
      
       
       'Suma para Totalizar por Remate
       sCadTotPiezas = sCadTotPiezas + 1
       sCadTotPesoNeto = sCadTotPesoNeto + rs!nPesoNeto
       sCadTotTasacion = sCadTotTasacion + rs!nTasacion
       sCadTotCapital = sCadTotCapital + rs!Capital
       sCadTotInteresVenc = sCadTotInteresVenc + rs!InteresVenc
       sCadTotComision = sCadTotComision + rs!Comision
       sCadTotTotalDeuda = sCadTotTotalDeuda + rs!TotalDeuda
       
       rs.MoveNext
       CambiarPagina (7)

    Loop Until x = rs.RecordCount + 1
davi:
       sCad = sCad
       sCadResumen = Chr$(10) & "  Total General: " & Chr$(10) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "CANTIDAD          :     " & ImpreFormat(sCadTotPiezas, 4, 0, True) & " piezas" & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "PESO NETO         :" & ImpreFormat(sCadTotPesoNeto, 13, 2, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "TASACION          :" & ImpreFormat(sCadTotTasacion, 13, 2, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "CAPITAL           :" & ImpreFormat(sCadTotCapital, 13, 2, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "INTERES VENCIDO   :" & ImpreFormat(sCadTotInteresVenc, 13, 2, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "COMISION          :" & ImpreFormat(sCadTotComision, 13, 2, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "TOTAL DEUDA       :" & ImpreFormat(sCadTotTotalDeuda, 13, 2, True) & Chr$(10)

       sCad = sCad & sCadResumen
End If
       'Devuelve toda la estructura  del Reporte
      RepPiezasRematadas = sCad

End Function



Function Report_158101(ByVal psFecini As String) As String
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
sql = sql & "  and datediff(dd, dPrdEstado,'" & Format(psFecini, "mm/dd/yyyy") & "') = 0 "
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
Report_158101 = strCuerpo

End Function


Public Function RepDistribuciónColocaciones() As String


'Variables para el Manejo de losDatos
Dim Co As nCredRepoFinMes
Dim rs As ADODB.Recordset
Dim strString As String                     ' Variable que contiene la consulta


'Negrita
Dim BoldOn As Variant
Dim BoldOff As Variant
BoldOn = Chr$(27) & Chr$(69)
BoldOff = Chr$(27) & Chr$(70)




Dim nItem As Integer

Dim cTitulo As String

Set Co = New nCredRepoFinMes

strString = " select u.cUbiGeoCod,u.cUbiGeoDescripcion as Distrito, a.cAgeCod , cAgeDescripcion as Agencia"
strString = strString & " ,count(distinct pp.cPersCod) Clientes ,count(distinct pp.cCtaCod) Contratos ,sum(x.Capital) Capital ,sum(x.Peso) Peso"
strString = strString & "  from productopersona pp"
strString = strString & " inner join persona p on p.cPersCod = pp.cPersCod"
strString = strString & " inner join ubicaciongeografica u on u.cUbiGeoCod = p.cPersDireccUbiGeo"
strString = strString & " inner join agencias a on a.cAgeCod = substring(pp.cCtaCod,4,2)"
strString = strString & " inner join (select  p.cCtaCod,p.nSaldo Capital,s.PesoNeto  Peso"
strString = strString & " from producto p"
strString = strString & " inner join(select c.cCtaCod, sum(cp.nPesoNeto) PesoNeto from colocpigno c"
strString = strString & " inner join colocpigjoyatasacion cp on cp.cCtaCod= c.cCtacod and cp.ntipoTasacion = c.nTipoTasacion"
strString = strString & " group by c.cCtaCod) s on s.cCtaCod = p.cCtaCod "
strString = strString & " where p.nprdestado  in (" & gPigEstAmortiz & "," & gPigEstCancelPendRes & "," & gPigEstDesemb & "," & gPigEstReusoLin & ")"
strString = strString & " ) x on x.cCtaCod= pp.cCtaCod"
strString = strString & " group by u.cUbiGeoCod,u.cUbiGeoDescripcion , a.cAgeCod , cAgeDescripcion"
strString = strString & " order by u.cUbiGeoDescripcion,a.cAgeCod"


Set rs = Co.GetQuery(strString)

'Inicializar variables
sCad = ""
scadColumnas = ""
'scad = scad & Chr$(15)                           'Retorna al tipo de letra normal
'scad = scad & Chr$(27) + Chr$(107) + Chr$(1)     'Tipo de Letra Roman
'scad = scad & Chr$(27) + Chr$(103)                'Tamaño 10.5 - 15 CPI
'scad = scad & Chr$(27) + Chr$(50)                'Espaciado entre lineas 1/16
''sCad = sCad & Chr$(27) + Chr$(67) + Chr$(18)     'Longitud de página a 24 líneas
'scad = scad & Chr$(27) + Chr$(108) + Chr$(6)     'Margen Izquierdo - 6ta. Columna


'   --- Constantes para Impresión  ------
' Introduce espacio                    :   Space(43)
' Enter para la siguiente linea   :   Chr$(10)
' Determina la cantidad de
' espacios devueltos                 :    ImpreFormat("campo", Cantidad de espacio, Decimales,¿moneda?)

'Cabecera para el Reporte
lnLineas = 8
lnPage = 1
lnAncho = 127
sCadTitulo = "CREDITO PIGNORATICIO"
sCadSubTitulo = "DISTRIBUCION DE COLOCACIONES"
sCadComenta = ""
scadColumnas = Chr(10) & scadColumnas & String(121, "-") & Chr(10)
scadColumnas = scadColumnas & Space(7) & "Distrito" & Space(5) & "Agencia " & Space(20) & "Clientes" & Space(6) & "Contratos" & Space(13) & "Capital" & Space(12) & "Gramaje" & Chr(10)
scadColumnas = scadColumnas & String(121, "-") & Chr(10)
sCad = sCad & nArmaCabeceraReporte(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, sCadTitulo, sCadSubTitulo, lnPage, lnAncho, scadColumnas, sCadComenta)


If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    sCadCodDistrito = rs!cUbiGeoCod
    Do
        lnLineas = lnLineas + 1
        x = x + 1
        If rs.EOF Then
        ultimo = 1
        rs.MoveLast
        End If
        If (rs!cUbiGeoCod <> sCadCodDistrito) Or (x = 1) Or (ultimo = 1) Then
           'Totalizo Agencia
           sCadTotDistrito = Space(12) & ImpreFormat(sCadTotClientesDistrito, 6, 0, True)
            sCadTotDistrito = sCadTotDistrito & Space(9) & ImpreFormat(sCadTotContratosDistrito, 6, 0, True)
           sCadTotDistrito = sCadTotDistrito & Space(7) & ImpreFormat(sCadTotCapitalDistrito, 10, 2, True)
           sCadTotDistrito = sCadTotDistrito & Space(6) & ImpreFormat(sCadTotPesoNetoDistrito, 10, 2, True)
           sCadTotClientesDistrito = 0
           sCadTotContratosDistrito = 0
           sCadTotCapitalDistrito = 0
           sCadTotPesoNetoDistrito = 0
           
           If x <> 1 Then
'              sCad = sCad & Chr$(10)
'              lnLineas = lnLineas + 1
              sCad = sCad & Space(46) & String(10, "-") & Space(5) & String(10, "-") & Space(5) & String(15, "-") & Space(4) & String(15, "-") & Chr$(10)
              lnLineas = lnLineas + 1
              sCad = sCad & Space(25) & " Sub - Total:" & sCadTotDistrito & Chr$(10)
              lnLineas = lnLineas + 1
              CambiarPagina (7)
           End If
        

           'Encabezo Distritoaxx
           
           sCadCodDistrito = rs!cUbiGeoCod
           sCadTitDistrito = Space(5) & ImpreFormat(rs!Distrito, 25)
        
           If ultimo <> 1 Then
              sCad = sCad & Chr$(10)
              sCad = sCad & sCadTitDistrito & Chr$(10) & Chr$(10)
              lnLineas = lnLineas + 3
              CambiarPagina (7)
            End If
           End If
       'Añado Detalle
       If ultimo = 1 Then GoTo davi
       SCADDETALLE = SCADDETALLE & Space(21) & Justifica(rs!cAgeCod, 2) & " -"
       SCADDETALLE = SCADDETALLE & Justifica(rs!Agencia, 18) & Space(7)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Clientes, 6, 0, True) & Space(9)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Contratos, 6, 0, True) & Space(7)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Capital, 10, 2, True) & Space(6)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Peso, 10, 2, True) & Chr$(10)
    
       sCad = sCad & SCADDETALLE
       SCADDETALLE = ""
       '************Segundo Quiebre****************
       'Suma para Totalizar por Distrito
       sCadTotClientesDistrito = sCadTotClientesDistrito + rs!Clientes
       sCadTotContratosDistrito = sCadTotContratosDistrito + rs!Contratos
       sCadTotCapitalDistrito = sCadTotCapitalDistrito + rs!Capital
       sCadTotPesoNetoDistrito = sCadTotPesoNetoDistrito + rs!Peso
      
       
       'Suma para Totalizar por Remate
       sCadTotClientes = sCadTotClientes + rs!Clientes
       sCadTotContratos = sCadTotContratos + rs!Contratos
       sCadTotCapital = sCadTotCapital + rs!Capital
       sCadTotPeso = sCadTotPeso + rs!Peso
       
       rs.MoveNext
       CambiarPagina (7)

    Loop Until x = rs.RecordCount + 1
davi:
       sCad = sCad
       sCadResumen = Chr$(10) & "  Total General: " & Chr$(10) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "CLIENTES     :            " & ImpreFormat(sCadTotClientes, 6, 0, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "CONTRATOS    :            " & ImpreFormat(sCadTotContratos, 6, 0, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "CAPITAL           :" & ImpreFormat(sCadTotCapital, 13, 2, True) & Chr$(10)
       sCadResumen = sCadResumen & Space(10) & "GRAMAJE           :" & ImpreFormat(sCadTotPeso, 13, 2, True) & Chr$(10)

       sCad = sCad & sCadResumen
End If
       'Devuelve toda la estructura  del Reporte
      RepDistribuciónColocaciones = sCad

End Function


Public Function RepVentaMaterial(psAgencia As String, psFecha As String) As String


'Variables para el Manejo de losDatos
Dim Co As nCredRepoFinMes
Dim rs As ADODB.Recordset
Dim strString As String                     ' Variable que contiene la consulta
Dim sCadMaterial As String
Dim sCadRemate As String
Dim xfecha As String
'Negrita
Dim BoldOn As Variant
Dim BoldOff As Variant
BoldOn = Chr$(27) & Chr$(69)
BoldOff = Chr$(27) & Chr$(70)

Dim nItem As Integer

Dim cTitulo As String

Set Co = New nCredRepoFinMes

Dim oRep As DpigReportes

Set oRep = New DpigReportes
xfecha = Right(psFecha, 4) & Left(psFecha, 2)
Set rs = oRep.RepVentaMaterial(psAgencia, xfecha, 1)

'Inicializar variables
sCad = ""
scadColumnas = ""

'Cabecera para el Reporte
lnLineas = 8
lnPage = 1
lnAncho = 127
sCadTitulo = "VENTAS DEL MES " & psFecha
If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    sCadSubTitulo = "UBICACION: " & rs!cAgeDescripcion
End If
sCadComenta = ""
scadColumnas = Chr(10) & scadColumnas & String(121, "-") & Chr(10)
scadColumnas = scadColumnas & Space(1) & "Material" & Space(8) & "Piezas " & Space(13) & "Peso" & Space(7) & "Adicional" & Space(5) & "Valor Venta" & Space(5) & "Costo Venta" & Space(3) & "Int Percibido" & Space(3) & "Com Percibido" & Chr(10)
scadColumnas = scadColumnas & String(121, "-") & Chr(10)
sCad = sCad & nArmaCabeceraReporte(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, sCadTitulo, sCadSubTitulo, lnPage, lnAncho, scadColumnas, sCadComenta)

If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
     sCadnRemate = rs!nRemate
    Do
        lnLineas = lnLineas + 1
        x = x + 1
        If rs.EOF Then
        ultimo = 1
        rs.MoveLast
        End If
        If (rs!nRemate <> sCadnRemate) Or (x = 1) Or (ultimo = 1) Then
           'Totalizo Agencia
           sCadTotRemate = Space(4) & ImpreFormat(sCadTotPiezas, 4, 0, True)
           sCadTotRemate = sCadTotRemate & Space(5) & ImpreFormat(sCadTotPeso, 10, 2, True)
           sCadTotRemate = sCadTotRemate & Space(3) & ImpreFormat(sCadTotAdicional, 10, 2, True)
           sCadTotRemate = sCadTotRemate & Space(3) & ImpreFormat(sCadTotValorVenta, 10, 2, True)
           sCadTotRemate = sCadTotRemate & Space(3) & ImpreFormat(sCadTotCostoVenta, 10, 2, True)
           sCadTotRemate = sCadTotRemate & Space(3) & ImpreFormat(sCadTotInteresesPercibidos, 10, 2, True)
           sCadTotRemate = sCadTotRemate & Space(3) & ImpreFormat(sCadTotComisionesPercibidas, 10, 2, True)

           sCadTotPiezas = 0
           sCadTotPeso = 0
           sCadTotAdicional = 0
           sCadTotValorVenta = 0
           sCadTotCostoVenta = 0
           sCadTotInteresesPercibidos = 0
           sCadTotComisionesPercibidas = 0
           
           
           If x <> 1 Then
                 sCad = sCad & Chr$(10)
                  CambiarPagina (5)
                 sCad = sCad & Space(16) & String(7, "-") & Space(8) & String(10, "-") & Space(1) & String(15, "-") & Space(1) & String(15, "-") & Space(1) & String(15, "-") & Space(1) & String(15, "-") & Space(1) & String(15, "-") & Chr$(10)
                 sCad = sCad & Space(2) & " Sub - Total:" & sCadTotRemate & Chr$(10)
           End If
           'Encabezo Distritoaxx
           
           sCadnRemate = rs!nRemate
          
           sCadTitRemate = Space(1) & "Remate " & rs!nRemate & ":"
        
           If ultimo <> 1 Then
              sCad = sCad & Chr$(10)
               CambiarPagina (5)
              sCad = sCad & sCadTitRemate & Chr$(10) & Chr$(10)
               CambiarPagina (5)
       End If
           End If
       'Añado Detalle
       If ultimo = 1 Then GoTo davi
       SCADDETALLE = SCADDETALLE & Space(5) & Justifica(rs!cConsDescripcion, 10) & Space(4)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Piezas, 4, 0, True) & Space(5)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Peso, 10, 2, True) & Space(3)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Adicional, 10, 2, True) & Space(3)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!ValorVenta, 10, 2, True) & Space(3)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!CostoVenta, 10, 2, True) & Space(3)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!InteresPercibido, 10, 2, True) & Space(3)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!ComisionesPercibidas, 10, 2, True) & Chr$(10)
    
       sCad = sCad & SCADDETALLE
       SCADDETALLE = ""
       '************Segundo Quiebre****************
       'Suma para Totalizar por Remate
       sCadTotPiezas = sCadTotPiezas + rs!Piezas
       sCadTotPeso = sCadTotPeso + rs!Peso
       sCadTotAdicional = sCadTotAdicional + rs!Adicional
       sCadTotValorVenta = sCadTotValorVenta + rs!ValorVenta
       sCadTotCostoVenta = sCadTotCostoVenta + rs!CostoVenta
       sCadTotInteresesPercibidos = sCadTotInteresesPercibidos + rs!InteresPercibido
       sCadTotComisionesPercibidas = sCadTotComisionesPercibidas + rs!ComisionesPercibidas
      
        
       rs.MoveNext
       CambiarPagina (5)

    Loop Until x = rs.RecordCount + 1
davi:
       sCad = sCad & Chr$(10) & Chr$(10)
        CambiarPagina (5)
       'Devuelve toda la estructura  del Reporte
Set oRep = Nothing
Set oRep = New DpigReportes
Set rs = oRep.RepVentaMaterial(psAgencia, xfecha, 2)
 CambiarPagina (5)
sCadMaterial = "Total general por Tipo de Material: " & Chr$(10) & Chr$(10)
 CambiarPagina (5)
If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    
    Do
     lnLineas = lnLineas + 1
               
       sCadMaterial = sCadMaterial & Space(5) & Justifica(rs!cConsDescripcion, 10) & Space(4)
       sCadMaterial = sCadMaterial & ImpreFormat(rs!Piezas, 4, 0, True) & Space(5)
       sCadMaterial = sCadMaterial & ImpreFormat(rs!Peso, 10, 2, True) & Space(3)
       sCadMaterial = sCadMaterial & ImpreFormat(rs!Adicional, 10, 2, True) & Space(3)
       sCadMaterial = sCadMaterial & ImpreFormat(rs!ValorVenta, 10, 2, True) & Space(3)
       sCadMaterial = sCadMaterial & ImpreFormat(rs!CostoVenta, 10, 2, True) & Space(3)
       sCadMaterial = sCadMaterial & ImpreFormat(rs!InteresPercibido, 10, 2, True) & Space(3)
       sCadMaterial = sCadMaterial & ImpreFormat(rs!ComisionesPercibidas, 10, 2, True) & Chr$(10)
                
               'Suma para Totalizar
                sCadTotPiezasMat = sCadTotPiezasMat + rs!Piezas
                sCadTotPesoMat = sCadTotPesoMat + rs!Peso
                sCadTotAdicionalMat = sCadTotAdicionalMat + rs!Adicional
                sCadTotValorVentaMat = sCadTotValorVentaMat + rs!ValorVenta
                sCadTotCostoVentaMat = CostoVentaMat + rs!CostoVenta
                sCadTotInteresPercibidoMat = sCadTotInteresPercibidoMat + rs!InteresPercibido
                sCadTotComisionesPercibidasMat = sCadTotComisionesPercibidasMat + rs!ComisionesPercibidas
              
            
                CambiarPagina (5)
                rs.MoveNext
          
       Loop Until rs.EOF
                         
               sCadTotMat = Space(4) & ImpreFormat(sCadTotPiezasMat, 4, 0, True)
               sCadTotMat = sCadTotMat & Space(5) & ImpreFormat(sCadTotPesoMat, 10, 2, True)
               sCadTotMat = sCadTotMat & Space(3) & ImpreFormat(sCadTotAdicionalMat, 10, 2, True)
               sCadTotMat = sCadTotMat & Space(3) & ImpreFormat(sCadTotValorVentaMat, 10, 2, True)
               sCadTotMat = sCadTotMat & Space(3) & ImpreFormat(sCadTotCostoVentaMat, 10, 2, True)
               sCadTotMat = sCadTotMat & Space(3) & ImpreFormat(sCadTotInteresPercibidoMat, 10, 2, True)
               sCadTotMat = sCadTotMat & Space(3) & ImpreFormat(sCadTotComisionesPercibidasMat, 10, 2, True)
               
               sCadMaterial = sCadMaterial & Space(16) & String(7, "-") & Space(8) & String(10, "-") & Space(1) & String(15, "-") & Space(1) & String(15, "-") & Space(1) & String(15, "-") & Space(1) & String(15, "-") & Space(1) & String(15, "-") & Chr$(10)
               sCadMaterial = sCadMaterial & Space(15) & sCadTotMat & Chr(10)
         
'Devuelve toda la estructura  del Reporte
        
 End If

      sCad = sCad & sCadMaterial
      
 Set oRep = Nothing
Set oRep = New DpigReportes
Set rs = oRep.RepVentaMaterial(psAgencia, xfecha, 3)
 CambiarPagina (5)
sCadRemate = "Total general por Remate: " & Chr$(10) & Chr$(10)
 CambiarPagina (5)
If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    
    Do
     lnLineas = lnLineas + 1
               
       sCadRemate = sCadRemate & Space(5) & "Remate  " & Justifica(rs!nRemate, 3) & Space(3)
       sCadRemate = sCadRemate & ImpreFormat(rs!Piezas, 4, 0, True) & Space(5)
       sCadRemate = sCadRemate & ImpreFormat(rs!Peso, 10, 2, True) & Space(3)
       sCadRemate = sCadRemate & ImpreFormat(rs!Adicional, 10, 2, True) & Space(3)
       sCadRemate = sCadRemate & ImpreFormat(rs!ValorVenta, 10, 2, True) & Space(3)
       sCadRemate = sCadRemate & ImpreFormat(rs!CostoVenta, 10, 2, True) & Space(3)
       sCadRemate = sCadRemate & ImpreFormat(rs!InteresPercibido, 10, 2, True) & Space(3)
       sCadRemate = sCadRemate & ImpreFormat(rs!ComisionesPercibidas, 10, 2, True) & Chr$(10)
                
               'Suma para Totalizar
                sCadTotPiezasRem = sCadTotPiezasRem + rs!Piezas
                sCadTotPesoRem = sCadTotPesoRem + rs!Peso
                sCadTotAdicionalRem = sCadTotAdicionalRem + rs!Adicional
                sCadTotValorVentaRem = sCadTotValorVentaRem + rs!ValorVenta
                sCadTotCostoVentaRem = sCadTotCostoVentaRem + rs!CostoVenta
                sCadTotInteresPercibidoRem = sCadTotInteresPercibidoRem + rs!InteresPercibido
                sCadTotComisionesPercibidasRem = sCadTotComisionesPercibidasRem + rs!ComisionesPercibidas
              
            
                CambiarPagina (5)
                rs.MoveNext
          
       Loop Until rs.EOF
                         
               sCadTotRemate = Space(4) & ImpreFormat(sCadTotPiezasRem, 4, 0, True)
               sCadTotRemate = sCadTotRemate & Space(5) & ImpreFormat(sCadTotPesoRem, 10, 2, True)
               sCadTotRemate = sCadTotRemate & Space(3) & ImpreFormat(sCadTotAdicionalRem, 10, 2, True)
               sCadTotRemate = sCadTotRemate & Space(3) & ImpreFormat(sCadTotValorVentaRem, 10, 2, True)
               sCadTotRemate = sCadTotRemate & Space(3) & ImpreFormat(sCadTotCostoVentaRem, 10, 2, True)
               sCadTotRemate = sCadTotRemate & Space(3) & ImpreFormat(sCadTotInteresPercibidoRem, 10, 2, True)
               sCadTotRemate = sCadTotRemate & Space(3) & ImpreFormat(sCadTotComisionesPercibidasRem, 10, 2, True)
               
               sCadRemate = sCadRemate & Space(16) & String(7, "-") & Space(8) & String(10, "-") & Space(1) & String(15, "-") & Space(1) & String(15, "-") & Space(1) & String(15, "-") & Space(1) & String(15, "-") & Space(1) & String(15, "-") & Chr$(10)
               sCadRemate = sCadRemate & Space(15) & sCadTotRemate & Chr(10)
         
'Devuelve toda la estructura  del Reporte
        
 End If
End If

      sCad = sCad & sCadRemate
     
      RepVentaMaterial = sCad

End Function

Public Function RepPiezasAdjTasador(pRemate As Integer) As String


'Variables para el Manejo de losDatos
Dim Co As nCredRepoFinMes
Dim rs As ADODB.Recordset
Dim strString As String                     ' Variable que contiene la consulta


'Negrita
Dim BoldOn As Variant
Dim BoldOff As Variant
BoldOn = Chr$(27) & Chr$(69)
BoldOff = Chr$(27) & Chr$(70)

'Varibles que pertenecen al cuerpo del reporte


Dim sCadGarantia As String
Dim sCadCustodia As String

' Variables que pasaran a Totalizar
Dim sCadTot1 As Variant
Dim sCadTotLotes1 As Variant
Dim sCadTotPBruto1 As Variant
Dim sCadTotPNeto1 As Variant
Dim sCadTotTasaccion1 As Variant
Dim sCadTotTasaccionAdicional1 As Variant

Dim sCadTot2 As Variant
Dim sCadTotLotes2 As Variant
Dim sCadTotPBruto2 As Variant
Dim sCadTotPNeto2 As Variant
Dim sCadTotTasaccion2 As Variant
Dim sCadTotTasaccionAdicional2 As Variant


Dim nItem As Integer

Dim cTitulo As String

Set Co = New nCredRepoFinMes

strString = " select   substring(a.cCtaCod,4,2) CodAgencia,b.cAgeDescripcion Agencia ,a.cCtaCod,a.nItemPieza"
strString = strString & " ,f.cConsDescripcion MatTas ,a.nPesoNeto,(a.nTasacion + a.nTasacionAdicional ) Tasacion"
strString = strString & " ,g.cConsDescripcion MatRetas ,d.nPesoNeto nPesoNetoRetas,(d.nTasacion + d.nTasacionAdicional ) Retasacion"
strString = strString & " ,j.nMonto Capital,j.nValorDeuda Deuda,j.nValorProceso Valor, e.cUser Tasador"
strString = strString & " from colocpigjoyatasacion a"
strString = strString & " inner join agencias b on b.cAgeCod = substring(a.cCtaCod,4,2)"
strString = strString & " inner join colocpigno c on c.cCtaCod = a.cCtaCod and c.nTipoTasacion = a.nTipoTasacion"
strString = strString & " inner join colocpigjoyatasacion d on d.cCtaCod = a.cCtaCod and d.nItemPieza=a.nItemPieza"
strString = strString & " inner join rrhh e on e.cPersCod=a.cPersCod"
strString = strString & " inner join constante f on f.nConsValor =a.nMaterial"
strString = strString & " inner join constante g on g.nConsValor =d.nMaterial"
strString = strString & " inner join (select h.nRemate,h.nTipoProceso,h.cCtaCod,h.nItemPieza,i.nMonto,h.nValorDeuda,h.nValorProceso"
strString = strString & "  from colocpigproceso h"
strString = strString & " inner join colocpigprocesodet i on i.nRemate= h.nRemate"
strString = strString & " and i.nTipoProceso=h.nTipoProceso"
strString = strString & " and i.cCtaCod=h.cCtaCod"
strString = strString & " and i.nItemPieza=h.nItemPieza"
strString = strString & " where h.nRemate = " & pRemate
'strString = strString & " and h.nTipoProceso =     gPigTipoAdjudicacionTasador & "
strString = strString & " and i.nCodConcepto= " & gColPigConceptoCodCapital & "  ) j on j.cCtaCod=a.cCtaCod and j.nItemPieza= a.nItemPieza"
strString = strString & " where d.nTipoTasacion = " & gPigTipoTasacRetasac & " and f.nConsCod= " & gColocPigMaterial & " and g.nConsCod = " & gColocPigMaterial & ""
strString = strString & " order by CodAgencia "





Set rs = Co.GetQuery(strString)

'Inicializar variables
sCad = ""
scadColumnas = ""
'scad = scad & Chr$(15)                           'Retorna al tipo de letra normal
'scad = scad & Chr$(27) + Chr$(107) + Chr$(1)     'Tipo de Letra Roman
'scad = scad & Chr$(27) + Chr$(103)                'Tamaño 10.5 - 15 CPI
'scad = scad & Chr$(27) + Chr$(50)                'Espaciado entre lineas 1/16
''sCad = sCad & Chr$(27) + Chr$(67) + Chr$(18)     'Longitud de página a 24 líneas
'scad = scad & Chr$(27) + Chr$(108) + Chr$(6)     'Margen Izquierdo - 6ta. Columna


'   --- Constantes para Impresión  ------
' Introduce espacio                    :   Space(43)
' Enter para la siguiente linea   :   Chr$(10)
' Determina la cantidad de
' espacios devueltos                 :    ImpreFormat("campo", Cantidad de espacio, Decimales,¿moneda?)

'Cabecera para el Reporte
lnLineas = 8
lnPage = 1
lnAncho = 159
sCadTitulo = "JOYAS CON RESPONSABILIDAD DE TASADOR"
sCadSubTitulo = "REMATE: " & pRemate
sCadComenta = ""
scadColumnas = Chr(10) & scadColumnas & String(153, "-") & Chr(10)
scadColumnas = scadColumnas & Space(24) & "|---------------TASACION --------------|" & "|--------------RETASACION--------------|" & "|-----------------ADJUDICACION------------------|" & Chr(10)
scadColumnas = scadColumnas & Space(13) & "Pieza" & Space(8) & "Material" & Space(7) & "Peso Neto" & Space(8) & "Valor" & Space(2) & "Material" & Space(6) & "Peso Neto" & Space(9) & "Valor" & Space(7) & "Capital" & Space(9) & "Deuda" & Space(9) & "Valor" & Space(2) & "Tasador" & Chr(10)

scadColumnas = scadColumnas & String(153, "-") & Chr(10) & Chr(10)
sCad = sCad & nArmaCabeceraReporte(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, sCadTitulo, sCadSubTitulo, lnPage, lnAncho, scadColumnas, sCadComenta)


If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    sCadCodAgencia = rs!CodAgencia
    Do
        lnLineas = lnLineas + 1
        x = x + 1
        If rs.EOF Then
        ultimo = 1
        rs.MoveLast
        End If
        If (rs!CodAgencia <> sCadCodAgencia) Or (x = 1) Or (ultimo = 1) Then
           'Totalizo Agencia
           sCadTotAgencia = Space(2) & ImpreFormat(sCadTotPiezasAgencia, 4, 0, True) & " piezas"
           sCadTotAgencia = sCadTotAgencia & Space(12) & ImpreFormat(sCadTotPesoNetoTasacionAgencia, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(1) & ImpreFormat(sCadTotTasacionAgencia, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(12) & ImpreFormat(sCadTotPesoNetoReTasacionAgencia, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(1) & ImpreFormat(sCadTotReTasacionAgencia, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(1) & ImpreFormat(sCadTotCapitalAgencia, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(1) & ImpreFormat(sCadTotDeudaAgencia, 10, 2, True)
           sCadTotAgencia = sCadTotAgencia & Space(1) & ImpreFormat(sCadTotValorAgencia, 10, 2, True)
            
           sCadTotPiezasAgencia = 0
           sCadTotPesoNetoTasacionAgencia = 0
           sCadTotTasacionAgencia = 0
           sCadTotPesoNetoReTasacionAgencia = 0
           sCadTotReTasacionAgencia = 0
           sCadTotCapitalAgencia = 0
           sCadTotDeudaAgencia = 0
           sCadTotValorAgencia = 0
                      
           If x <> 1 Then
                sCad = sCad & Chr$(10)
                sCad = sCad & Space(37) & String(12, "-") & Space(2) & String(12, "-") & Space(13) & String(12, "-") & Space(2) & String(12, "-") & Space(2) & String(12, "-") & Space(2) & String(12, "-") & Space(2) & String(12, "-") & Chr$(10)
                sCad = sCad & Space(1) & "Sub-Total:" & sCadTotAgencia & Chr$(10)
           End If
           'Encabezo Agencia
           sCadCodAgencia = rs!CodAgencia
           sCadTitAgencia = Space(1) & "(" & ImpreFormat(rs!CodAgencia, 2, 0, True) & ")" & Space(1) & ImpreFormat(rs!Agencia, 25)
        
           If ultimo <> 1 Then sCad = sCad & Chr$(10) & sCadTitAgencia & Chr$(10) & Chr$(10)
           End If
       'Añado Detalle
       If ultimo = 1 Then GoTo davi
       SCADDETALLE = SCADDETALLE & Space(2) & Justifica(rs!cCtaCod, 18) & " -"
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!nItemPieza, 2, 0, False) & Space(1)
       SCADDETALLE = SCADDETALLE & Justifica(rs!MatTas, 10) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!nPesoNeto, 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Tasacion, 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & Justifica(rs!MatRetas, 10) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!nPesoNetoRetas, 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!ReTasacion, 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Capital, 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Deuda, 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Valor, 10, 2, True) & Space(1)
       SCADDETALLE = SCADDETALLE & ImpreFormat(rs!Tasador, 10, 2, True) & Chr$(10)
    
       sCad = sCad & SCADDETALLE
       SCADDETALLE = ""
       '************Segundo Quiebre****************
       'Suma para Totalizar por Agencia
       sCadTotPiezasAgencia = sCadTotPiezasAgencia + 1
       sCadTotPesoNetoTasacionAgencia = sCadTotPesoNetoTasacionAgencia + rs!nPesoNeto
       sCadTotTasacionAgencia = sCadTotTasacionAgencia + rs!Tasacion
       sCadTotPesoNetoReTasacionAgencia = sCadTotPesoNetoReTasacionAgencia + rs!nPesoNetoRetas
       sCadTotReTasacionAgencia = sCadTotReTasacionAgencia + rs!ReTasacion
       sCadTotTotalCapitalAgencia = sCadTotTotalCapitalAgencia + rs!Capital
       sCadTotTotalDeudaAgencia = sCadTotTotalDeudaAgencia + rs!Deuda
       sCadTotTotalValorAgencia = sCadTotTotalValorAgencia + rs!Valor

      
       
       'Suma para Totalizar por Remate
        sCadTotPiezas = sCadTotPiezas + 1
       sCadTotPesoNetoTas = sCadTotPesoNetoTas + rs!nPesoNeto
       sCadTotTasacion = sCadTotTasacion + rs!Tasacion
       sCadTotPesoNetoRetas = sCadTotPesoNetoRetas + rs!nPesoNetoRetas
       sCadTotReTasacion = sCadTotReTasacion + rs!ReTasacion
       sCadTotTotalCapital = sCadTotTotalCapital + rs!Capital
       sCadTotTotalDeuda = sCadTotTotalDeuda + rs!Deuda
       sCadTotTotalValor = sCadTotTotalValor + rs!Valor
       
       rs.MoveNext
       CambiarPagina (7)

    Loop Until x = rs.RecordCount + 1
davi:
       sCad = sCad & Chr(10)
       
        sCadTot = Space(2) & ImpreFormat(sCadTotPiezas, 4, 0, True) & " piezas"
        sCadTot = sCadTot & Space(12) & ImpreFormat(sCadTotPesoNetoTas, 10, 2, True)
        sCadTot = sCadTot & Space(1) & ImpreFormat(sCadTotTasacion, 10, 2, True)
        sCadTot = sCadTot & Space(12) & ImpreFormat(sCadTotPesoNetoRetas, 10, 2, True)
        sCadTot = sCadTot & Space(1) & ImpreFormat(sCadTotReTasacion, 10, 2, True)
        sCadTot = sCadTot & Space(1) & ImpreFormat(sCadTotTotalCapital, 10, 2, True)
        sCadTot = sCadTot & Space(1) & ImpreFormat(sCadTotTotalDeuda, 10, 2, True)
        sCadTot = sCadTot & Space(1) & ImpreFormat(sCadTotTotalValor, 10, 2, True)
        
        sCad = sCad & Space(37) & String(12, "-") & Space(2) & String(12, "-") & Space(13) & String(12, "-") & Space(2) & String(12, "-") & Space(2) & String(12, "-") & Space(2) & String(12, "-") & Space(2) & String(12, "-") & Chr$(10)
        sCad = sCad & Space(5) & "Total:" & sCadTot & Chr(10)
        
     
End If
       'Devuelve toda la estructura  del Reporte
      RepPiezasAdjTasador = sCad

End Function
Public Function RepMovimientosMensualesClientes(psCodCliente As String) As String

'Variables para el Manejo de losDatos
Dim Co As nCredRepoFinMes
Dim rs As ADODB.Recordset
Dim strString As String                     ' Variable que contiene la consulta

'Negrita
Dim BoldOn As Variant
Dim BoldOff As Variant
BoldOn = Chr$(27) & Chr$(69)
BoldOff = Chr$(27) & Chr$(70)

'Varibles que pertenecen al cuerpo del reporte
Dim sCad As String

' Variables que pasaran a Totalizar
Dim sCadTotPiezas As Variant
Dim sCadTotPBruto As Variant
Dim sCadTotPNeto As Variant
Dim sCadTotTasaccion As Variant
Dim sCadTotValorAdj As Variant
Dim sCadTotValorAdic As Variant

Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer

Dim cTitulo As String

Set Co = New nCredRepoFinMes


strString = " select year(b.dPago) Año ,month(b.dPago) ,d.cConsDescripcion Mes,e.cPersNombre"
strString = strString & " ,sum(case nPrdConceptoCod when " & gColPigConceptoCodCapital & " then nMontoPagado else 0 end) Amortizacion"
strString = strString & " ,sum(case nPrdConceptoCod when " & gColPigConceptoCodInteresComp & "  then nMontoPagado else 0 end) InteresCompensatorio"
strString = strString & ",sum(case nPrdConceptoCod when " & gcolPigConceptoCodInteresMora & "  then nMontoPagado else 0 end) Mora"
strString = strString & ",sum(case nPrdConceptoCod when " & gColPigConceptoCodComiServ & "  then nMontoPagado else 0 end) ComisionServicio"
strString = strString & ",sum(case nPrdConceptoCod when " & gColPigConceptoCodPreparaRemate & "  then nMontoPagado else 0 end) DerechoRemate"
strString = strString & ",sum(case nPrdConceptoCod when " & gColPigConceptoCodComiVencida & "  then nMontoPagado else 0 end) ComisionVencida"
strString = strString & " from ColocCalendDetPig a"
strString = strString & " inner join ColocCalendario b on b.cCtaCod=a.cCtaCod and b.nNroCalen=a.nNroCalen"
strString = strString & " and b.nColocCalendApl =a.nColocCalendApl and b.nCuota=a.nCuota"
strString = strString & " inner join productopersona c on c.cCtaCod = b.cCtaCod"
strString = strString & " inner join Constante d on d.nConsValor = month(b.dPago)"
strString = strString & " inner join persona e on e.cPersCod = c.cPersCod"
strString = strString & " where c.cPersCod= " & psCodCliente
strString = strString & " and b.dPago >  cast( '01/01/' + cast(year(dateadd(year,-2,getdate())) as char) as datetime)"
strString = strString & " and b.nColocCalendApl =" & gColocCalendAplCuota
strString = strString & " and d.nConsCod= " & gMeses
strString = strString & " group by year(b.dPago) ,month(b.dPago),d.cConsDescripcion,e.cPersNombre"
strString = strString & " order by año desc, month(b.dPago)desc"

Set rs = Co.GetQuery(strString)

'Inicializar variables
sCad = ""
scadColumnas = ""
'Cabecera para el Reporte
lnLineas = 8
lnPage = 1
lnAncho = 145
sCadTitulo = "SISTEMA PIGNORATICIO"
sCadSubTitulo = "MOVIMIENTOS MENSUALES POR CLIENTE"
sCadComenta = ""
scadColumnas = Chr(10) & scadColumnas & String(140, "-") & Chr(10)
scadColumnas = scadColumnas & Space(1) & "MES" & Space(9) & "AÑO" & Space(11) & "AMORTIZACION" & Space(6) & "INT VENCIDO" & Space(10) & "MORA" & Space(5) & "COMISION SERVICIO" & Space(3) & "DERECHO REMATE" & Space(3) & "COMISION VENCIDA" & Space(5) & "TOTAL" & Chr(10)
scadColumnas = scadColumnas & String(140, "-") & Chr(10) & Chr(10)
sCad = sCad & nArmaCabeceraReporte(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, sCadTitulo, sCadSubTitulo, lnPage, lnAncho, scadColumnas, sCadComenta)

If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    sCad = sCad & Space(1) & "CLIENTE: " & psCliente & " " & rs!cPersNombre & Chr$(10) & Chr$(10)
    
    Do
     lnLineas = lnLineas + 1
               
                sCad = sCad & Space(1) & Justifica(rs!Mes, 10) & Space(1)
                sCad = sCad & ImpreFormat(rs!Año, 5, 0, False) & Space(3)
                sCad = sCad & ImpreFormat(rs!Amortizacion, 13, 2, True) & Space(1)
                sCad = sCad & ImpreFormat(rs!InteresCompensatorio, 13, 2, True) & Space(1)
                sCad = sCad & ImpreFormat(rs!Mora, 13, 2, True) & Space(1)
                sCad = sCad & ImpreFormat(rs!ComisionServicio, 13, 2, True) & Space(1)
                sCad = sCad & ImpreFormat(rs!DerechoRemate, 13, 2, True) & Space(1)
                sCad = sCad & ImpreFormat(rs!ComisionVencida, 13, 2, True) & Space(1)
                sCad = sCad & ImpreFormat(rs!Amortizacion + rs!InteresCompensatorio + rs!Mora + rs!ComisionServicio + rs!DerechoRemate + rs!ComisionVencida, 13, 2, True) & Chr$(10)
                
               'Suma para Totalizar
                sCadTotAmortizacion = sCadTotAmortizacion + rs!Amortizacion
                sCadTotInteresCompensatorio = sCadTotInteresCompensatorio + rs!InteresCompensatorio
                sCadTotMora = sCadTotMora + rs!Mora
                sCadTotComisionServicio = sCadTotComisionServicio + rs!ComisionServicio
                sCadTotDerechoRemate = sCadTotDerechoRemate + rs!DerechoRemate
                sCadTotComisionVencida = sCadTotComisionVencida + rs!ComisionVencida
                sCadTot = sCadTot + rs!Amortizacion + rs!InteresCompensatorio + rs!Mora + rs!ComisionServicio + rs!DerechoRemate + rs!ComisionVencida
              
            
                CambiarPagina (1)
                rs.MoveNext
          
       Loop Until rs.EOF
                         
             ' sCadTot1 = sCadTot1 & String(120, "-") & Chr(10)
              sCadTot1 = Space(9) & "TOTALES :" & Space(5) & ImpreFormat(sCadTotAmortizacion, 13, 0, True) & Space(1) & ImpreFormat(sCadTotInteresCompensatorio, 13, 2, True) & Space(1)
              sCadTot1 = sCadTot1 & ImpreFormat(sCadTotMora, 13, 2, True) & Space(1) & ImpreFormat(sCadTotComisionServicio, 13, 2, True) & Space(1)
              sCadTot1 = sCadTot1 & ImpreFormat(sCadTotDerechoRemate, 13, 2, True) & Space(1) & ImpreFormat(sCadTotComisionVencida, 13, 2, True) & Space(1) & ImpreFormat(sCadTot, 13, 2, True)
              sCad = sCad & String(140, "-") & Chr(10) & sCadTot1 & Chr(10)
         
'Devuelve toda la estructura  del Reporte
        
 End If

 RepMovimientosMensualesClientes = sCad

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



'Cabecera de Impresiones de Reportes de Pignoraticio
Public Function nArmaCabeceraReporte(ByVal psNomCmac As String, ByVal psNomAgencia As String, ByVal psCodUser As String, _
        ByVal psFechaSis As String, ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psColumnas As String _
        , ByVal psComenta As String) As String
      
  Dim lsCabe01 As String, lsCabe02 As String
  Dim lsCabe03 As String, lsCabe04 As String
  Dim lsCabe05 As String, lsCabRepo As String
  
  lsCabRepo = ""
  ' Cabecera 1
  lsCabe01 = FillText(Trim(UCase(psNomCmac)), 55, " ")
  lsCabe01 = lsCabe01 & Space(pnAnchoLinea - 55 - 25)
  lsCabe01 = lsCabe01 & "Pag.  : " & Str(pnPagina) & "  -  " & psCodUser & Chr(10)
  ' Cabecera 2
  lsCabe01 = lsCabe01 & FillText(Trim(UCase(psNomAgencia)), 35, " ")
  lsCabe01 = lsCabe01 & Space(pnAnchoLinea - 35 - 30)
  lsCabe01 = lsCabe01 & "Fecha : " & Format(psFechaSis & " " & Time, "dd/mm/yyyy hh:mm") & Chr$(10)
  ' Titulo
  lsCabe02 = String(Int((pnAnchoLinea - Len(psTitulo)) / 2), " ") & psTitulo & Chr$(10)
  ' SubTitulo
  lsCabe03 = String(Int((pnAnchoLinea - Len(psSubTitulo)) / 2), " ") & psSubTitulo & Chr$(10)
  'Comentario
  lsCabe04 = IIf(Len(psComenta) > 0, psComenta & Chr$(10), "")
' Columnas
  lsCabe05 = psColumnas
  ' ***
  lsCabRepo = lsCabRepo & lsCabe01 & lsCabe02
  lsCabRepo = lsCabRepo & lsCabe03 & lsCabe04 & lsCabe05
  nArmaCabeceraReporte = lsCabRepo
End Function

Public Sub CambiarPagina(NumLineasMinimo As Integer)
If lnLineas >= 55 - NumLineasMinimo Then
   lnPage = lnPage + 1
   sCad = sCad & Chr(12)
   sCad = sCad & Chr(10)
   sCad = sCad & nArmaCabeceraReporte(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, sCadTitulo, sCadSubTitulo, lnPage, lnAncho, scadColumnas, sCadComenta)
   lnLineas = 8
End If
End Sub
