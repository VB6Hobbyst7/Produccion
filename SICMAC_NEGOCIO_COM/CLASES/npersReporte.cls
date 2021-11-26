VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NPersReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim scad As String
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




'
'Function ChangePage(ByVal strX As String) As String
'Str = Cabecera & strX
'ChangePage = Str & Chr(12) & Chr(10)
'End Function
'
'
'

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
  lsCabe01 = lsCabe01 & Space(pnAnchoLinea - 35 - 25)
  lsCabe01 = lsCabe01 & "Fecha : " & Format(psFechaSis & " " & Time, "dd/mm/yyyy hh:mm") & Chr$(10)
  ' Titulo
  lsCabe02 = String(Int((pnAnchoLinea - Len(psTitulo)) / 2), " ") & psTitulo & Chr$(10)
  ' SubTitulo
  lsCabe03 = String(Int((pnAnchoLinea - Len(psSubTitulo)) / 2), " ") & psSubTitulo & Chr$(10)
  'Comentario
  lsCabe04 = IIf(Len(psComenta) > 0, psComenta & Chr$(10), "")
' Columnas
  lsCabe05 = psColumnas & Chr$(10)
  ' ***
  lsCabRepo = lsCabRepo & lsCabe01 & lsCabe02
  lsCabRepo = lsCabRepo & lsCabe03 & lsCabe04 & lsCabe05
  nArmaCabeceraReporte = lsCabRepo
End Function

Public Sub CambiarPagina(NumLineasMinimo As Integer)
If lnLineas >= 55 - NumLineasMinimo Then
   lnPage = lnPage + 1
   scad = scad & Chr(12)
   scad = scad & Chr(10)
   scad = scad & nArmaCabeceraReporte(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, sCadTitulo, sCadSubTitulo, lnPage, lnAncho, scadColumnas, sCadComenta)
   lnLineas = 8
End If
End Sub


Public Function ReporteDeClientes(txtFecini) As String

'Variables para el Manejo de losDatos
'Dim ocon As Conecta
Dim Co As nCredRepoFinMes
Dim Rsc As ADODB.Recordset
Dim Rs As ADODB.Recordset
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
Dim ssql As String

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

Dim lnmargen As Integer

Dim rsE As New ADODB.Recordset
Dim oRep As DpersReporte

Set oRep = New DpersReporte
Set rsE = oRep.ReporteDeClientes(txtFecini)


'Inicializar variables
scad = ""
sCadComenta = ""
scad = scad & Chr$(15)                           'Retorna al tipo de letra normal
scad = scad & Chr$(27) + Chr$(108) + Chr$(0)     'Tipo de Letra Roman
scad = scad & Chr$(27) + Chr$(77)                'Tamaño 10.5 - 15 CPI
scad = scad & Chr$(27) + Chr$(15)                'Espaciado entre lineas 1/16
'sCad = sCad & Chr$(27) + Chr$(67) + Chr$(18)     'Longitud de página a 24 líneas
scad = scad & Chr$(27) + Chr$(108) + Chr$(6)     'Margen Izquierdo - 6ta. Columna


'''gImprimir.arcsal
'''gImprimir.ImpreBegin True, 66
'''
''''Chr$(27) & Chr$(108) & Chr$(0); 'Tipo letra : 0,1,2 - Roman,SansS,Courier
''''   Print #arcsal, Chr$(27) & Chr$(77);            'Tamaño  : 80, 77, 103
''''   Print #arcsal, Chr$(15);                       'Retorna al tipo de letra normal
'''' Dim arcsal As Integer

'   --- Constantes para Impresión  ------
' Introduce espacio                    :   Space(43)
' Enter para la siguiente linea   :   Chr$(10)
' Determina la cantidad de
' espacios devueltos                 :    ImpreFormat("campo", Cantidad de espacio, Decimales,¿moneda?)

'Cabecera para el Reporte
lnLineas = 6
lnPage = 1
sCadTitulo = "APERTURA DE CLIENTES"
sCadSubTitulo = "APERTURA DE CLIENTES " & Format(txtFecini, gcFormatoFechaView)
lnAncho = 203

sCadComenta = ""
scadColumnas = Chr(10) & scadColumnas & String(203, "-") & Chr(10)
scadColumnas = scadColumnas & "CODIGO" & Space(10) & "NOMBRES " & Space(45) & "DIRECCION" & Space(37) & "UBIGEO" & Space(10) & "DESCRIPCION" & Space(17) & "DOC. IDENTIDAD" & Space(7) & "TELEFONO    USUARIO" & Space(10) & oImpresora.gPrnSaltoLinea
scadColumnas = scadColumnas & String(203, "-")
scad = scad & nArmaCabeceraReporte(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, sCadTitulo, sCadSubTitulo, lnPage, lnAncho, scadColumnas, sCadComenta)

 
 
If Not (rsE.EOF And rsE.BOF) Then
    rsE.MoveFirst
    sCadAgencia = rsE!Agencia
    Do
        lnLineas = lnLineas + 1
        x = x + 1
        If rsE.EOF Then
        ultimo = 1
        rsE.MoveLast
        End If
        ''If (Rs!Agencia <> sCadCodAgencia) Or (Rs!remate <> sCadRemate) Or (X = 1) Or (ultimo = 1) Then
        If (rsE!Agencia <> sCadCodAgencia) Or (x = 1) Or (ultimo = 1) Then
           'Totalizo Agencia
           Dim vcontrato As String
'           sCadTotPiezasAgencia = sCadTotPiezasAgencia + 1
            If sCadTotPiezasAgencia > 1 Then
            vcontrato = "   Clientes"
            Else
            vcontrato = "   Cliente"
            End If
           
           sCadTotAgencia = Space(2) & ImpreFormat(sCadTotPiezasAgencia, 4, 0, True) & vcontrato
           sCadTotPiezasAgencia = 0
           
           If x <> 1 Then scad = scad & Chr$(10) & Space(5) & " Total " & sCadTotAgencia & Chr$(10)
           
           'Encabezo Agencia
           sCadCodAgencia = rsE!Agencia
           sCadTitAgencia = Space(5) & "(" & ImpreFormat(rsE!Agencia, 2, 0, True) & ")" & Space(1) & ImpreFormat(rsE!cAgeDescripcion, 25)
        
           If ultimo <> 1 Then scad = scad & Chr$(10) & sCadTitAgencia & Chr$(10) & Chr$(10)
           End If
       'Añado Detalle
        If ultimo = 1 Then GoTo PieCabecera
       If ultimo = 1 Then GoTo PieCabecera
 '       SCADDETALLE = SCADDETALLE & lscad1 & Space(lnMargen) & rsE!cPersCod & Space(5) & ImpreFormat(rsE!cpersNombre, 50) & Space(5) & ImpreFormat(rsE!cPersDireccDomicilio, 20) & Space(5) & ImpreFormat(rsE!cPersDireccUbiGeo, 15) & Space(5) & ImpreFormat(rsE!cUbiGeoDescripcion, 20) & Space(5) & ImpreFormat(rsE!cConsDescripcion, 5) & Space(5) & rsE!cPersIDNro & ImpreFormat(rsE!cPersTelefono, 10) & Space(5) & rsE!Fecha & Space(5) & rsE!Usuario & Space(5) & oImpresora.gPrnSaltoLinea
 SCADDETALLE = SCADDETALLE & Space(lnmargen) & rsE!cPersCod & Space(1) & ImpreFormat(rsE!cpersNombre, 50) & Space(2) & ImpreFormat(rsE!cPersDireccDomicilio, 40) & Space(3) & ImpreFormat(rsE!cPersDireccUbiGeo, 13) & Space(1) & ImpreFormat(rsE!cUbiGeoDescripcion, 20) & Space(5) & ImpreFormat(rsE!cConsDescripcion, 5) & Space(2) & ImpreFormat(rsE!cPersIDNro, 11) & ImpreFormat(rsE!cPersTelefono, 10) & Space(5) & rsE!Usuario & Space(5) & oImpresora.gPrnSaltoLinea

       scad = scad & SCADDETALLE
       SCADDETALLE = ""
       '************Segundo Quiebre****************
       'Suma para Totalizar por Agencia
        sCadTotPiezasAgencia = sCadTotPiezasAgencia + 1


       rsE.MoveNext
       CambiarPagina (2)

    Loop Until x = rsE.RecordCount + 1
PieCabecera:
       
       scad = scad & sCadResumen
      
      'Devuelve toda la estructura  del Reporte
   ReporteDeClientes = scad

 End If

End Function


'    If Not (rsE.EOF And rsE.BOF) Then
''        RaiseEvent ShowProgress
'        lnMargen = 5
'        lsCadena = lsCadena & Space(lnMargen) & psEmpresa & oImpresora.gPrnSaltoLinea
'        lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnSaltoLinea
'        lsCadena = lsCadena & Space(lnMargen) & CentrarCadena("APERTURA DE CLIENTES - " & Format(gsNomAge, gsFormatoFechaView), 102) & oImpresora.gPrnSaltoLinea
'        lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnSaltoLinea
'        lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnSaltoLinea
'        lsCadena = lsCadena & Space(lnMargen) & "CODIGO" & Space(14) & "NOMBRES " & Space(49) & "DIRECCION" & Space(19) & "UBIGEO" & Space(28) & "DOC. IDENTIDAD" & Space(20) & "TELOEFONO    USUARIO" & Space(10) & oImpresora.gPrnSaltoLinea
'
'
'        While Not rsE.EOF
'            lnCorr = lnCorr + 1
'            lsCodigo = rsE!cPersCod
'
'
'          lscad1 = lscad1 & Space(lnMargen) & rsE!cPersCod & Space(5) & ImpreFormat(rsE!cpersNombre, 50) & Space(5) & ImpreFormat(rsE!cPersDireccDomicilio, 20) & Space(5) & ImpreFormat(rsE!cPersDireccUbiGeo, 15) & Space(5) & ImpreFormat(rsE!cUbiGeoDescripcion, 20) & Space(5) & ImpreFormat(rsE!cConsDescripcion, 5) & Space(5) & rsE!cPersIDNro & ImpreFormat(rsE!cPersTelefono, 10) & Space(5) & rsE!Fecha & Space(5) & rsE!Usuario & Space(5) & oImpresora.gPrnSaltoLinea
'           'Format(ImpreFormat(rsE!Fecha, 10), gcFormatoFechaView)
'
'        Wend
'
' lsCadena = lsCadena & Chr(10) & lscad1 & oImpresora.gPrnSaltoLinea & Space(15) & "Total Agencia " & Space(15) & ImpreFormat(lnCorr, 13) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
' lsCadena = lsCadena & Space(lnMargen) & Space(39) '& ImpreFormat(LVTOTALINGRESO - LVTOTALDESCUENTO, 13, 2, True)
'
'
'    End If
'
'   ReporteDeClientes = lsCadena
'    rsE.Close
'    Set rsE = Nothing
'End Function


Public Function ReporteDeClientesMod(txtFecini) As String

'Variables para el Manejo de losDatos
'Dim ocon As Conecta
Dim Co As nCredRepoFinMes
Dim Rsc As ADODB.Recordset
Dim Rs As ADODB.Recordset
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
Dim ssql As String

Dim sCadCabezera As String
Dim sCadCabezeraHeader As String

Dim sCadCodAgencia As String
Dim x As Integer
Dim ultimo As Integer
Dim sCadTotPiezasAgencia As Integer
Dim sCadTotAgencia  As String
Dim sCadTitAgencia  As String
Dim SCADDETALLE As String
Dim sCadResumen As String

Dim lnmargen As Integer

Dim rsE As New ADODB.Recordset
Dim oRep As DpersReporte

Set oRep = New DpersReporte
Set rsE = oRep.ReporteDeClientesMod(txtFecini)


'Inicializar variables
scad = ""
sCadComenta = ""
scad = scad & Chr$(15)                           'Retorna al tipo de letra normal
scad = scad & Chr$(27) + Chr$(108) + Chr$(0)     'Tipo de Letra Roman
scad = scad & Chr$(27) + Chr$(77)                'Tamaño 10.5 - 15 CPI
scad = scad & Chr$(27) + Chr$(15)                'Espaciado entre lineas 1/16
scad = scad & Chr$(27) + Chr$(108) + Chr$(6)     'Margen Izquierdo - 6ta. Columna

'Cabecera para el Reporte
lnLineas = 6
lnPage = 1
sCadTitulo = "CLIENTES"
sCadSubTitulo = "MODIFICACION DE CLIENTES " & Format(txtFecini, gcFormatoFechaView)
lnAncho = 203

sCadComenta = ""
scadColumnas = Chr(10) & scadColumnas & String(203, "-") & Chr(10)
scadColumnas = scadColumnas & "CODIGO" & Space(10) & "NOMBRES " & Space(45) & "DIRECCION" & Space(37) & "UBIGEO" & Space(10) & "DESCRIPCION" & Space(17) & "DOC. IDENTIDAD" & Space(7) & "TELEFONO    USUARIO" & Space(10) & oImpresora.gPrnSaltoLinea
scadColumnas = scadColumnas & String(203, "-")
scad = scad & nArmaCabeceraReporte(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, sCadTitulo, sCadSubTitulo, lnPage, lnAncho, scadColumnas, sCadComenta)

 
 
If Not (rsE.EOF And rsE.BOF) Then
    rsE.MoveFirst
    sCadAgencia = rsE!Agencia
    Do
        lnLineas = lnLineas + 1
        x = x + 1
        If rsE.EOF Then
        ultimo = 1
        rsE.MoveLast
        End If
        ''If (Rs!Agencia <> sCadCodAgencia) Or (Rs!remate <> sCadRemate) Or (X = 1) Or (ultimo = 1) Then
        If (rsE!Agencia <> sCadCodAgencia) Or (x = 1) Or (ultimo = 1) Then
           'Totalizo Agencia
           Dim vcontrato As String
'           sCadTotPiezasAgencia = sCadTotPiezasAgencia + 1
            If sCadTotPiezasAgencia > 1 Then
            vcontrato = "   Clientes"
            Else
            vcontrato = "   Cliente"
            End If
           
           sCadTotAgencia = Space(2) & ImpreFormat(sCadTotPiezasAgencia, 4, 0, True) & vcontrato
           sCadTotPiezasAgencia = 0
           
           If x <> 1 Then scad = scad & Chr$(10) & Space(5) & " Total " & sCadTotAgencia & Chr$(10)
           
           'Encabezo Agencia
           sCadCodAgencia = rsE!Agencia
           sCadTitAgencia = Space(5) & "(" & ImpreFormat(rsE!Agencia, 2, 0, True) & ")" & Space(1) & ImpreFormat(rsE!cAgeDescripcion, 25)
        
           If ultimo <> 1 Then scad = scad & Chr$(10) & sCadTitAgencia & Chr$(10) & Chr$(10)
           End If
       'Añado Detalle
        If ultimo = 1 Then GoTo PieCabecera
       If ultimo = 1 Then GoTo PieCabecera
 '       SCADDETALLE = SCADDETALLE & lscad1 & Space(lnMargen) & rsE!cPersCod & Space(5) & ImpreFormat(rsE!cpersNombre, 50) & Space(5) & ImpreFormat(rsE!cPersDireccDomicilio, 20) & Space(5) & ImpreFormat(rsE!cPersDireccUbiGeo, 15) & Space(5) & ImpreFormat(rsE!cUbiGeoDescripcion, 20) & Space(5) & ImpreFormat(rsE!cConsDescripcion, 5) & Space(5) & rsE!cPersIDNro & ImpreFormat(rsE!cPersTelefono, 10) & Space(5) & rsE!Fecha & Space(5) & rsE!Usuario & Space(5) & oImpresora.gPrnSaltoLinea
 SCADDETALLE = SCADDETALLE & Space(lnmargen) & rsE!cPersCod & Space(1) & ImpreFormat(rsE!cpersNombre, 50) & Space(2) & ImpreFormat(rsE!cPersDireccDomicilio, 40) & Space(3) & ImpreFormat(rsE!cPersDireccUbiGeo, 13) & Space(1) & ImpreFormat(rsE!cUbiGeoDescripcion, 20) & Space(5) & ImpreFormat(rsE!cConsDescripcion, 5) & Space(2) & ImpreFormat(rsE!cPersIDNro, 11) & ImpreFormat(rsE!cPersTelefono, 10) & Space(5) & rsE!Usuario & Space(5) & oImpresora.gPrnSaltoLinea

       scad = scad & SCADDETALLE
       SCADDETALLE = ""
       '************Segundo Quiebre****************
       'Suma para Totalizar por Agencia
        sCadTotPiezasAgencia = sCadTotPiezasAgencia + 1


       rsE.MoveNext
       CambiarPagina (2)

    Loop Until x = rsE.RecordCount + 1
PieCabecera:
       
       scad = scad & sCadResumen
      
      'Devuelve toda la estructura  del Reporte
   ReporteDeClientesMod = scad

 End If

End Function


