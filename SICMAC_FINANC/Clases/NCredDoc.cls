VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCredDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private vsCadImpresion As String
Private lsNegritaOn As String
Private lsNegritaOff  As String
Private lsSaltoLin As String
Private lsTab As String
Private EspHorxPag As Integer
Private EspVerxPag As Integer
Private nPuntPag As Integer
Private sCadCab As String
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

Public Function ImprimeCabeceraGarantiaReal(ByVal psNomAge As String, _
        ByVal psFecha As String, ByVal psHora As String, _
        ByVal psPersNombre As String, Optional ByVal psCodCMAC As String) As String

Dim lsCadena As String
'lsCadena = Chr$(10)
lsCadena = Chr$(10) & Chr(10)
If psCodCMAC = "102" Then
    lsCadena = lsCadena & lsNegritaOn & ImpreFormat("CM - CREDITOS ", 43, , False) & Space(6) & Chr(10) & lsNegritaOff
    lsCadena = lsCadena & ImpreFormat(psNomAge, 43, , False) & Chr(10)
    lsCadena = lsCadena & ImpreFormat("Fecha:" & psFecha & Space(5) & "Hora:" & psHora, 43, , False) & Chr(10)
    lsCadena = lsCadena & lsNegritaOn & ImpreFormat(PstaNombre(psPersNombre), 43, , False) & Chr(10)
    lsCadena = lsCadena & ImpreFormat("----------GASTOS DE LEVANTAMIENTO----------", 43, , False) & Chr(10) & lsNegritaOff
Else
    lsCadena = lsCadena & lsNegritaOn & ImpreFormat("CMACT - CREDITOS ", 43, , False) & Space(6) & ImpreFormat("CMACT - CREDITOS ", 43, , False) & Chr(10) & lsNegritaOff
    lsCadena = lsCadena & ImpreFormat(psNomAge, 43, , False) & Space(6) & ImpreFormat(psNomAge, 43, , False) & Chr(10)
    lsCadena = lsCadena & ImpreFormat("Fecha:" & psFecha & Space(5) & "Hora:" & psHora, 43, , False) & Space(6) & ImpreFormat("Fecha:" & psFecha & Space(5) & "Hora:" & psHora, 43, , False) & Chr(10)
    lsCadena = lsCadena & lsNegritaOn & ImpreFormat(PstaNombre(psPersNombre), 43, , False) & Space(6) & ImpreFormat(PstaNombre(psPersNombre), 43, , False) & Chr(10)
    lsCadena = lsCadena & ImpreFormat("----------GASTOS DE LEVANTAMIENTO----------", 43, , False) & Space(6) & ImpreFormat("----------GASTOS DE LEVANTAMIENTO----------", 43, , False) & Chr(10) & lsNegritaOff
End If
ImprimeCabeceraGarantiaReal = lsCadena
End Function
Public Function ImprimeBoletaGarantiaR(ByVal psNomAge As String, _
        ByVal psFecha As String, ByVal psHora As String, _
        ByVal Nrog As String, ByVal psCodUser As String, ByVal pnMovG As Long, psGCmact As String, pdFecSist As Date, Optional ByVal psCodCMAC As String) As String
Dim Cabecera As String
Dim cuerpo As String
Dim loConecta As DConecta
Dim Dg As DGarantia
Dim Rs As ADODB.Recordset
Dim sSQL As String
Dim Total As Double
Dim Lineas As Integer
Dim llena As String
Dim I As Integer
Lineas = 7
Set Dg = New DGarantia
Set Rs = New ADODB.Recordset
'Set Rs = Dg.RecuperaTitularGarantReal(Nrog)
Cabecera = ImprimeCabeceraGarantiaReal(psNomAge, psFecha, psHora, Rs!Nombre, psCodCMAC)
Total = 0
'Detalle

sSQL = " Select( Select cDescripcion from ProductoConcepto where nPrdConceptoCod=MGD.nPrdConceptoCod) Concepto,"
sSQL = sSQL & " nMonto"
sSQL = sSQL & " from MovGarantDet MGD"
sSQL = sSQL & " Where MGD.nMovNro =" & pnMovG
    
    Set loConecta = New DConecta
        loConecta.AbreConexion
        Set Rs = loConecta.CargaRecordSet(sSQL)
        loConecta.CierraConexion
  
cuerpo = Cabecera
If Rs.BOF And Rs.EOF Then

Else

    cuerpo = cuerpo & Chr(10)
    While Not Rs.EOF
     If psCodCMAC = "102" Then
        cuerpo = cuerpo & ImpreFormat(Rs!Concepto, 23, 0) & Space(8) & ImpreFormat(Rs!nMonto, 10, 2, True) & Chr(10)
        Lineas = Lineas + 1
        Total = Total + Rs!nMonto
     Else
        cuerpo = cuerpo & ImpreFormat(Rs!Concepto, 23, 0) & Space(8) & ImpreFormat(Rs!nMonto, 10, 2, True) & Space(6) & ImpreFormat(Rs!Concepto, 23, 0, False) & Space(10) & ImpreFormat(Rs!nMonto, 10, 2, True) & Chr(10)
        Lineas = Lineas + 1
        Total = Total + Rs!nMonto
     End If
    Rs.MoveNext
    Wend
    
End If
 If psCodCMAC = "102" Then
    cuerpo = cuerpo & lsNegritaOn & String(45, "-") & Chr(10)
    cuerpo = cuerpo & ImpreFormat("Total Gastos", 23, 0) & Space(8) & ImpreFormat(Total, 10, 2, True) & Chr(10)
    cuerpo = cuerpo & psCodUser & lsNegritaOff
    Lineas = Lineas + 3
  Else
     cuerpo = cuerpo & lsNegritaOn & String(45, "-") & Space(6) & String(45, "-") & Chr(10)
    cuerpo = cuerpo & ImpreFormat("Total Gastos", 23, 0) & Space(8) & ImpreFormat(Total, 10, 2, True) & Space(8) & ImpreFormat("Total Gastos", 23, 0, False) & Space(8) & ImpreFormat(Total, 10, 2, True) & Chr(10)
    cuerpo = cuerpo & psCodUser & Space(48) & psCodUser & lsNegritaOff
    Lineas = Lineas + 3
 End If
'If lineas <= 22 Then
'    lineas = 6
'
'
'End If

If Lineas <= 18 Then
    For I = Lineas To 18
        cuerpo = cuerpo + Chr(10)
        
    Next I
End If
Set Rs = Nothing
Set loConecta = Nothing
Set Dg = Nothing
ImprimeBoletaGarantiaR = cuerpo
End Function


Private Function lnSaltoLinDoc() As String
    nPuntPag = nPuntPag + 1
    If nPuntPag > EspVerxPag Then
        nPuntPag = 0
        lnSaltoLinDoc = Chr$(12) & sCadCab
    Else
        lnSaltoLinDoc = Chr$(10)
    End If
    
End Function


Private Sub ImprimeCabeceraBoletaPago(ByVal psNomAge As String, ByVal psMoneda As String, _
        ByVal psCuotasPagadas As String, ByVal psFecha As String, ByVal psHora As String, _
        ByVal psPersNombre As String, ByVal psNroTransac As String, ByVal psCtaCod As String, _
        ByVal psTipoCredito As String, Optional ByVal psNomCmacAbrev As String = "", Optional ByVal psCodCMAC As String)
       ' psNomAge = "Sede Institucional"
        On Error GoTo ErrorImprimeCabeceraBoletaPago
        vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(64)
        vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(50)    'espaciamiento lineas 1/6 pulg.
        vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(67) & Chr$(22)   'Longitud de página a 22 líneas'
        vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(77)    'Tamaño 10 cpi
        vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
        vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(18)  ' cancela condensada
        vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(72)  ' desactiva negrita
        
        vsCadImpresion = vsCadImpresion & lsNegritaOn
        'vsCadImpresion = vsCadImpresion & lsSaltoLin
        'vsCadImpresion = vsCadImpresion & lsSaltoLin
        vsCadImpresion = vsCadImpresion & lsSaltoLin
        If psCodCMAC = "102" Then
            vsCadImpresion = vsCadImpresion & ImpreFormat(psNomCmacAbrev & " - CREDITOS" & Space(5) & ImpreFormat(Mid(psTipoCredito, 1, 25), 51, 0), 53) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 53, 0) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & lsNegritaOff
            vsCadImpresion = vsCadImpresion & ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & ImpreFormat("Credito : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 53, 0) & lsSaltoLin
        Else
            vsCadImpresion = vsCadImpresion & ImpreFormat(psNomCmacAbrev & " - CREDITOS" & Space(5) & ImpreFormat(Mid(psTipoCredito, 1, 25), 51, 0), 53) & Trim(ImpreFormat(psNomCmacAbrev & " - CREDITOS" & Space(5) & ImpreFormat(Mid(psTipoCredito, 1, 25), 51, 0), 53)) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 53, 0) & Trim(ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 53, 0)) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & lsNegritaOff
            vsCadImpresion = vsCadImpresion & ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0) & Trim(ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0)) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0) & Trim(ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0)) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & ImpreFormat("Credito : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 53, 0) & ImpreFormat("Credito : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 53, 0) & lsSaltoLin
        End If
        
        Exit Sub

ErrorImprimeCabeceraBoletaPago:
        Err.Raise Err.Number, "Error En Proceso ImprimeCabeceraBoletaPago", Err.Description
    
End Sub
Public Sub ImprimeBoletaExtorno(ByVal psDescrip As String, ByVal psCtaCod As String, ByVal psPersNombre As String, ByVal psNomAge As String, _
            ByVal psMoneda As String, ByVal psCuotasPagadas As String, ByVal psFecha As String, ByVal psHora As String, _
            ByVal psNroTransac As String, ByVal pnCapPagado As Double, _
            ByVal pnSaldoCap As Double, ByVal psCodUsu As String, ByVal sLpt As String, Optional ByVal psNomCmacAbrev As String = "", Optional ByVal psCodCMAC)
            
Dim nFicSal As Integer
Dim oDCred As DCredito
Dim R As ADODB.Recordset

    On Error GoTo ErrorImprimeBoleta
    vsCadImpresion = ""
    nFicSal = FreeFile

    Open sLpt For Output As nFicSal
    
    Set oDCred = New DCredito
    Set R = oDCred.RecuperaProducto(psCtaCod)
    Set oDCred = Nothing
    
    pnSaldoCap = R!nSaldo
    R.Close
    Set R = Nothing
    
    Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
    Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
    Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
    Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
    'Print #nFicSal, Chr$(27) + Chr$(15) ' condensada
    'Print #nFicSal, Chr$(27) + Chr$(18) ' cancela condensada
    Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita
    
    Print #nFicSal, ""
    Call ImprimeCabeceraBoletaPago(psNomAge, psMoneda, "", _
            psFecha, psHora, psPersNombre, "", psCtaCod, "", psNomCmacAbrev, psCodCMAC)
    
    'Cuerpo de la Boleta
If psCodCMAC = "102" Then
    vsCadImpresion = vsCadImpresion & lsNegritaOn
    vsCadImpresion = vsCadImpresion & ImpreFormat(String(7, "-") & psDescrip & String(7, "-"), 40) & lsSaltoLin
    '45 menos la longitud de su etiqueta
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Monto  :" & Space(13) & ImpreFormat(pnCapPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    vsCadImpresion = vsCadImpresion & String(40, "-") & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & ImpreFormat(pnCapPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsNegritaOff
    vsCadImpresion = vsCadImpresion & "Saldo Capital :" & Space(7) & ImpreFormat(pnSaldoCap, 12) & lsSaltoLin
    'vsCadImpresion = vsCadImpresion & ImpreFormat("Prox. Fecha Pago : " & Space(10) & Space(3) & psCodUsu, 53, 0) & ImpreFormat("Prox. Fecha Pago : " & Space(10) & Space(3) & psCodUsu, 53, 0) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    Print #nFicSal, vsCadImpresion
    'Print #nFicSal, lsSaltoLin
    Print #nFicSal, ""
    Print #nFicSal, ""
    Print #nFicSal, ""
    Close #nFicSal
Else
vsCadImpresion = vsCadImpresion & lsNegritaOn
    vsCadImpresion = vsCadImpresion & ImpreFormat(String(7, "-") & psDescrip & String(7, "-"), 50)
    vsCadImpresion = vsCadImpresion & ImpreFormat(String(7, "-") & psDescrip & String(7, "-"), 40) & lsSaltoLin
    '45 menos la longitud de su etiqueta
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Monto  :" & Space(13) & ImpreFormat(pnCapPagado, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Monto   :" & Space(13) & ImpreFormat(pnCapPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    vsCadImpresion = vsCadImpresion & String(40, "-")
    vsCadImpresion = vsCadImpresion & Space(14) & String(40, "-") & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & ImpreFormat(pnCapPagado, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Total Pagado :" & Space(8) & ImpreFormat(pnCapPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsNegritaOff
    vsCadImpresion = vsCadImpresion & "Saldo Capital :" & Space(7) & ImpreFormat(pnSaldoCap, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Saldo Capital :" & Space(7) & ImpreFormat(pnSaldoCap, 12) & lsSaltoLin
    'vsCadImpresion = vsCadImpresion & ImpreFormat("Prox. Fecha Pago : " & Space(10) & Space(3) & psCodUsu, 53, 0) & ImpreFormat("Prox. Fecha Pago : " & Space(10) & Space(3) & psCodUsu, 53, 0) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    Print #nFicSal, vsCadImpresion
    'Print #nFicSal, lsSaltoLin
    Print #nFicSal, ""
    Print #nFicSal, ""
    Print #nFicSal, ""
    Close #nFicSal
End If
    Exit Sub
ErrorImprimeBoleta:
    Err.Raise Err.Number, "Error En Proceso ImprimeBoleta", Err.Description
    
End Sub
Public Sub ImprimeBoletaExtornoAhorros(ByVal psDescrip As String, ByVal psCtaCod As String, ByVal psPersNombre As String, ByVal psNomAge As String, _
            ByVal psMoneda As String, ByVal psCuotasPagadas As String, ByVal psFecha As String, ByVal psHora As String, _
            ByVal psNroTransac As String, ByVal pnCapPagado As Double, _
            ByVal pnSaldoCap As Double, ByVal psCodUsu As String, ByVal sLpt As String, Optional ByVal psNomCmacAbrev As String = "", Optional ByVal psCodCMAC As String)
            
Dim nFicSal As Integer
Dim oDCred As DCredito
Dim R As ADODB.Recordset

    
    vsCadImpresion = ""
    nFicSal = FreeFile

    Open sLpt For Output As nFicSal
    
    Set oDCred = New DCredito
    Set R = oDCred.RecuperaProducto(psCtaCod)
    Set oDCred = Nothing
    
    pnSaldoCap = R!nSaldo
    R.Close
    Set R = Nothing
    
    Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
    Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
    Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
    Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
    'Print #nFicSal, Chr$(27) + Chr$(15) ' condensada
    'Print #nFicSal, Chr$(27) + Chr$(18) ' cancela condensada
    Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita
    
    Print #nFicSal, ""
    'Call ImprimeCabeceraBoletaPago(psNomAge, psMoneda, "", _
    '        psFecha, psHora, psPersNombre, "", psCtaCod, "", psNomCmacAbrev)
    
    'Cabecera
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(64)
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(50)    'espaciamiento lineas 1/6 pulg.
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(67) & Chr$(22)   'Longitud de página a 22 líneas'
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(77)    'Tamaño 10 cpi
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(18)  ' cancela condensada
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(72)  ' desactiva negrita
    
    vsCadImpresion = vsCadImpresion & lsNegritaOn
    vsCadImpresion = vsCadImpresion & lsSaltoLin
 If psCodCMAC = "102" Then
    vsCadImpresion = vsCadImpresion & ImpreFormat(psNomCmacAbrev & " - AHORROS", 52) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psNomAge, 1, 22), 53, 0) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsNegritaOff
    vsCadImpresion = vsCadImpresion & ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    vsCadImpresion = vsCadImpresion & ImpreFormat("Cuenta : " & psCtaCod, 53, 0) & lsSaltoLin
    
    'Cuerpo de la Boleta
    vsCadImpresion = vsCadImpresion & lsNegritaOn
    vsCadImpresion = vsCadImpresion & ImpreFormat(String(7, "-") & psDescrip & String(7, "-"), 50) & lsSaltoLin
    '45 menos la longitud de su etiqueta
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Monto  :" & Space(13) & ImpreFormat(pnCapPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    vsCadImpresion = vsCadImpresion & String(40, "-") & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & ImpreFormat(pnCapPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsNegritaOff
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsSaltoLin
 Else
 vsCadImpresion = vsCadImpresion & ImpreFormat(psNomCmacAbrev & " - AHORROS", 52) & Trim(ImpreFormat(psNomCmacAbrev & " - AHORROS", 53)) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psNomAge, 1, 22), 53, 0) & Trim(ImpreFormat(Mid(psNomAge, 1, 22), 53, 0)) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsNegritaOff
    vsCadImpresion = vsCadImpresion & ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0) & Trim(ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0)) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    vsCadImpresion = vsCadImpresion & ImpreFormat("Cuenta : " & psCtaCod, 53, 0) & ImpreFormat("Cuenta : " & psCtaCod, 53, 0) & lsSaltoLin
    
    'Cuerpo de la Boleta
    vsCadImpresion = vsCadImpresion & lsNegritaOn
    vsCadImpresion = vsCadImpresion & ImpreFormat(String(7, "-") & psDescrip & String(7, "-"), 50)
    vsCadImpresion = vsCadImpresion & ImpreFormat(String(7, "-") & psDescrip & String(7, "-"), 40) & lsSaltoLin
    '45 menos la longitud de su etiqueta
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Monto  :" & Space(13) & ImpreFormat(pnCapPagado, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Monto   :" & Space(13) & ImpreFormat(pnCapPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    vsCadImpresion = vsCadImpresion & String(40, "-")
    vsCadImpresion = vsCadImpresion & Space(14) & String(40, "-") & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & ImpreFormat(pnCapPagado, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Total Pagado :" & Space(8) & ImpreFormat(pnCapPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsNegritaOff
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsSaltoLin
 End If
    Print #nFicSal, vsCadImpresion
    'Print #nFicSal, lsSaltoLin
    Print #nFicSal, ""
    Print #nFicSal, ""
    Print #nFicSal, ""
    Close #nFicSal
    
End Sub


Public Sub ImprimeBoleta(ByVal psCtaCod As String, ByVal psPersNombre As String, ByVal psNomAge As String, _
            ByVal psMoneda As String, ByVal psCuotasPagadas As String, ByVal psFecha As String, ByVal psHora As String, _
            ByVal psNroTransac As String, ByVal psTipoCredito As String, ByVal pnCapPagado As Double, _
            ByVal pnIntCompPagado As Double, ByVal pnIntCompVencPagado, ByVal pnIntMorPagado As Double, ByVal pnGastoPagado As Double, _
            ByVal pnIntGraciaPagado As Double, ByVal pnIntSuspReprog As Double, ByVal pnSaldoCap As Double, _
            ByVal psProxFecPago As String, ByVal psCodUsu As String, ByVal sLpt As String, Optional ByVal psNomCmacAbrev As String = "", _
            Optional ByVal pbPagoCheque As Boolean = False, Optional ByVal psNumCheque As String = "", Optional ByVal psCodCMAC As String)
            
Dim nFicSal As Integer


    vsCadImpresion = ""
    nFicSal = FreeFile

    Open sLpt For Output As nFicSal
    
    
    Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
    Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
    Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
    Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
    'Print #nFicSal, Chr$(27) + Chr$(15) ' condensada
    'Print #nFicSal, Chr$(27) + Chr$(18) ' cancela condensada
    Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita
    
    'Print #nFicSal, ""
    Call ImprimeCabeceraBoletaPago(psNomAge, psMoneda, psCuotasPagadas, _
            psFecha, psHora, psPersNombre, psNroTransac, psCtaCod, psTipoCredito, psNomCmacAbrev, psCodCMAC)
    
    'Cuerpo de la Boleta
    vsCadImpresion = vsCadImpresion & lsNegritaOn
    
    
    
    '45 menos la longitud de su etiqueta
If psCodCMAC = "102" Then
    
    If pbPagoCheque Then
        vsCadImpresion = vsCadImpresion & "Cheque  :" & Space(2) & ImpreFormat(psNumCheque, 12) & lsSaltoLin
    End If
    If pbPagoCheque Then
        vsCadImpresion = vsCadImpresion & String(10, "-") & "PAGO CON CHEQUE    " & String(10, "-") & lsSaltoLin
    Else
        vsCadImpresion = vsCadImpresion & String(10, "-") & "COMPROBANTE DE PAGO" & String(10, "-") & lsSaltoLin
    End If
    
    vsCadImpresion = vsCadImpresion & "Capital :" & Space(13) & ImpreFormat(pnCapPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Interes Compensatorio:" & ImpreFormat(pnIntCompPagado + pnIntCompVencPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Mora :" & Space(16) & ImpreFormat(pnIntMorPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Gastos :" & Space(14) & ImpreFormat(pnGastoPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Int. Gracia:" & Space(10) & ImpreFormat(pnIntGraciaPagado, 12) & lsSaltoLin
    'vsCadImpresion = vsCadImpresion & Space(16) & "Int. Gracia:" & Space(10) & ImpreFormat(pnIntGraciaPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Int. Susp y Reprog:" & Space(3) & ImpreFormat(pnIntSuspReprog, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & String(40, "-") & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & ImpreFormat(pnCapPagado + pnIntCompPagado + pnIntCompVencPagado + pnIntMorPagado + pnGastoPagado + pnIntGraciaPagado + pnIntSuspReprog, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsNegritaOff
    vsCadImpresion = vsCadImpresion & "Saldo Capital :" & Space(7) & ImpreFormat(pnSaldoCap, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & ImpreFormat("Prox. Fecha Pago : " & psProxFecPago & Space(3) & psCodUsu, 53, 0) & lsSaltoLin
    Print #nFicSal, vsCadImpresion
    Print #nFicSal, ""
    Close #nFicSal
Else
    If pbPagoCheque Then
        vsCadImpresion = vsCadImpresion & "Cheque  :" & Space(2) & ImpreFormat(psNumCheque, 12)
        vsCadImpresion = vsCadImpresion & Space(28) & "Cheque  :" & Space(2) & ImpreFormat(psNumCheque, 12) & lsSaltoLin
    End If
    If pbPagoCheque Then
        vsCadImpresion = vsCadImpresion & String(10, "-") & "PAGO CON CHEQUE    " & String(10, "-")
        vsCadImpresion = vsCadImpresion & Space(14) & String(10, "-") & "PAGO CON CHEQUE    " & String(10, "-") & lsSaltoLin
    Else
        vsCadImpresion = vsCadImpresion & String(10, "-") & "COMPROBANTE DE PAGO" & String(10, "-")
        vsCadImpresion = vsCadImpresion & Space(14) & String(10, "-") & "COMPROBANTE DE PAGO" & String(10, "-") & lsSaltoLin
    End If
    vsCadImpresion = vsCadImpresion & "Capital :" & Space(13) & ImpreFormat(pnCapPagado, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Capital :" & Space(13) & ImpreFormat(pnCapPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Interes Compensatorio:" & ImpreFormat(pnIntCompPagado + pnIntCompVencPagado, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Interes Compensatorio:" & ImpreFormat(pnIntCompPagado + pnIntCompVencPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Mora :" & Space(16) & ImpreFormat(pnIntMorPagado, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Mora :" & Space(16) & ImpreFormat(pnIntMorPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Gastos :" & Space(14) & ImpreFormat(pnGastoPagado, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Gastos :" & Space(14) & ImpreFormat(pnGastoPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Int. Gracia:" & Space(10) & ImpreFormat(pnIntGraciaPagado, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Int. Gracia:" & Space(10) & ImpreFormat(pnIntGraciaPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Int. Susp y Reprog:" & Space(3) & ImpreFormat(pnIntSuspReprog, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Int. Susp y Reprog:" & Space(3) & ImpreFormat(pnIntSuspReprog, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & String(40, "-")
    vsCadImpresion = vsCadImpresion & Space(14) & String(40, "-") & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & ImpreFormat(pnCapPagado + pnIntCompPagado + pnIntCompVencPagado + pnIntMorPagado + pnGastoPagado + pnIntGraciaPagado + pnIntSuspReprog, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Total Pagado :" & Space(8) & ImpreFormat(pnCapPagado + pnIntCompPagado + pnIntCompVencPagado + pnIntMorPagado + pnGastoPagado + pnIntGraciaPagado + pnIntSuspReprog, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsNegritaOff
    vsCadImpresion = vsCadImpresion & "Saldo Capital :" & Space(7) & ImpreFormat(pnSaldoCap, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Saldo Capital :" & Space(7) & ImpreFormat(pnSaldoCap, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & ImpreFormat("Prox. Fecha Pago : " & psProxFecPago & Space(3) & psCodUsu, 53, 0) & ImpreFormat("Prox. Fecha Pago : " & psProxFecPago & Space(3) & psCodUsu, 53, 0) & lsSaltoLin
    Print #nFicSal, vsCadImpresion
    Print #nFicSal, ""
    Close #nFicSal
End If
End Sub
Private Sub Class_Initialize()
    lsNegritaOn = Chr$(27) + Chr$(71)
    lsNegritaOff = Chr$(27) + Chr$(72)
    lsSaltoLin = Chr$(10)
    EspHorxPag = 170
    EspVerxPag = 62
    nPuntPag = 0
    lsTab = Space(1)
End Sub

Private Sub ImprimeCabeceraDocumento(ByRef psCadImp As String, ByVal psNomAge As String, _
    ByVal psFechaHora As String, ByVal psCodUsu, ByVal psTitulo As String, ByVal psNomCmac As String, _
    Optional psTab As String = "", Optional pbCondensado As Boolean = True, Optional psCodRepo As String = "")
    nPuntPag = 0
    psCadImp = psCadImp & lnSaltoLinDoc
    If Len(psCodRepo) > 0 Then
        psTitulo = psCodRepo & " - " & psTitulo
    End If
    If pbCondensado Then
        psCadImp = psCadImp & psTab & psNomCmac & Space(70) & "Fecha :" & psFechaHora & lnSaltoLinDoc
        psCadImp = psCadImp & psTab & ImpreFormat(psNomAge, 45, 0) & Space(51) & "USUARIO : " & psCodUsu & lnSaltoLinDoc
        psCadImp = psCadImp & psTab & Space((EspHorxPag - Len(psTitulo)) / 2 - 18) & psTitulo & lnSaltoLinDoc
    Else
        psCadImp = psCadImp & psTab & psNomCmac & Space(45) & "Fecha :" & psFechaHora & lnSaltoLinDoc
        psCadImp = psCadImp & psTab & ImpreFormat(psNomAge, 40, 0) & Space(36) & "USUARIO : " & psCodUsu & lnSaltoLinDoc
        psCadImp = psCadImp & psTab & Space((120 - Len(psTitulo)) / 2 - 18) & psTitulo & lnSaltoLinDoc
    End If
End Sub

Public Function ImprimirMetasAnalistas(ByVal psNomAge As String, _
    ByVal pdFechaHora As Date, ByVal psCodUsu As String, ByVal psNomCmac As String)

Dim sCadImp As String
Dim RAnalista As ADODB.Recordset
Dim RDatos As ADODB.Recordset
Dim oCredGen As DCredGeneral
Dim oDCred As DCredito
Dim I As Integer
Dim oBase As DCredActualizaBD
Dim dFecSis As Date
Dim sCadProd As String

    On Error GoTo ErrorImprimirMetasAnalistas
    Screen.MousePointer = 11
    Set oBase = New DCredActualizaBD
    dFecSis = oBase.dFechaHora
    Set oBase = Nothing
    dFecSis = CDate(Format(Format(pdFechaHora, "dd/mm/yyyy") & " " & Format(dFecSis, "hh:mm:ss"), "dd/mm/yyyy hh.mm:ss"))
    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, Format(dFecSis, "dd/mm/yyyy hh:mm:ss"), psCodUsu, "LISTADO DE METAS POR ANALISTA", psNomCmac)
    sCadImp = sCadImp & String(EspHorxPag, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & ImpreFormat("Tipo", 12, 0) & ImpreFormat("Fec Inicio", 14, 0) & ImpreFormat("Fec Final", 15, 0) & Space(9)
    
    Set oDCred = New DCredito
    Set RDatos = oDCred.RecuperaProductosDeCredito
    Set oDCred = Nothing
    Do While Not RDatos.EOF
        sCadImp = sCadImp & ImpreFormat(Trim(RDatos!cAbrev), 10, 1)
        RDatos.MoveNext
    Loop
    RDatos.Close
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & String(EspHorxPag, "-") & lnSaltoLinDoc
    
    Set oCredGen = New DCredGeneral
    Set RAnalista = oCredGen.CargaAnalistas
    Set oCredGen = Nothing
    Do While Not RAnalista.EOF
        Set oDCred = New DCredito
        Set RDatos = oDCred.RecuperaDatosMetasAnalistaReporte(RAnalista!cPersCod)
        Set oDCred = Nothing
        
        If Not RDatos.EOF And Not RDatos.BOF Then
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & lsNegritaOn
            sCadImp = sCadImp & "ANALISTA : " & PstaNombre(RAnalista!cPersNombre)
            sCadImp = sCadImp & lsNegritaOff & lnSaltoLinDoc
        End If
        
        Do While Not RDatos.EOF
            sCadImp = sCadImp & ImpreFormat(Trim(RDatos!cTipoMeta), 7) & ImpreFormat(Format(RDatos!dInicial, "dd/mm/yyyy"), 12) _
                & ImpreFormat(Format(RDatos!dfinal, "dd/mm/yyyy"), 12) & ImpreFormat(Trim(RDatos!sMoneda), 8)
            'Imprimir los Productos
            sCadProd = Trim(RDatos!nTipoCred)
            Do While Not RDatos.EOF
                sCadImp = sCadImp & ImpreFormat(RDatos!nMonto, 8)
                RDatos.MoveNext
                If RDatos.EOF Then
                    Exit Do
                End If
                If sCadProd = Trim(RDatos!nTipoCred) Then
                    Exit Do
                End If
            Loop
            sCadImp = sCadImp & lnSaltoLinDoc
        Loop
        RDatos.Close
        Set RDatos = Nothing
        sCadImp = sCadImp & lnSaltoLinDoc
        RAnalista.MoveNext
    Loop
    RAnalista.Close
    Set RAnalista = Nothing
    ImprimirMetasAnalistas = sCadImp & Chr$(12)
    Screen.MousePointer = 0
    Exit Function

ErrorImprimirMetasAnalistas:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
    
End Function

Public Function ImprimeConsultaCredito(ByVal psCtaCod As String, ByVal psNomAge As String, ByVal psFecha As String, _
    ByVal psCodUsu As String, ByVal psTipoCred As String, ByVal psEstado As String, ByVal psLineaCred As String, ByVal psFteIngreso As String, _
    ByVal MatRelacCred As Variant, ByVal psAnalista As String, ByVal psCondicion As String, ByVal psDestino As String, _
    ByVal psApoderado As String, ByVal psTipoCuota As String, ByVal MatHitorial As Variant, ByVal MatDesembolsos As Variant, _
    ByVal MatCuotasPend As Variant, ByVal MatDeudaVenc As Variant, ByVal MatDeudaAFecha As Variant, ByVal psProtesto As String, _
    ByVal psCargoAutom As String, ByVal psRefinanciado As String, ByVal psNomCmac As String, ByVal pMatGarantias As Variant, _
    ByVal pnCuotaCom As Integer, ByVal pnMiViv As Integer, ByVal pnCalendDin As Integer, ByVal psTasaComp As String, ByVal psTasaMor As String) As String
    
Dim sCadImp As String
Dim I As Integer

    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecha, psCodUsu, "Resumen General de Credito", psNomCmac, Space(5))
    sCadImp = sCadImp & lnSaltoLinDoc & Space(5) & lsNegritaOn & "Credito : " & psCtaCod & lsNegritaOff & Space(20) & "Tipo : " & psTipoCred & Space(20) & "Estado : " & psEstado & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & lsNegritaOn & "CLIENTES RELACIONADOS" & Space(45) & "DATOS GENERALES DEL CREDITO" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(60, "-") & Space(5) & String(60, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "Nombre" & Space(35) & "Relaciones" & Space(15) & "Linea de Credito      : " & psLineaCred & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(60, "-") & Space(6) & "Fuente de Ingreso     : " & psFteIngreso & lnSaltoLinDoc
    For I = 0 To 5
        sCadImp = sCadImp & Space(5)
        If I <= UBound(MatRelacCred) - 1 Then
            sCadImp = sCadImp & ImpreFormat(MatRelacCred(I, 0), 40, 0) & ImpreFormat(MatRelacCred(I, 1), 20, 0)
        End If
        Select Case I
            Case 0
                sCadImp = sCadImp & Space(6) & "Analista Responsable  : " & psAnalista & lnSaltoLinDoc
            Case 1
                sCadImp = sCadImp & Space(6) & "Condicion del Credito : " & psCondicion & lnSaltoLinDoc
            Case 2
                sCadImp = sCadImp & Space(66) & "Destino del Credito   : " & psDestino & lnSaltoLinDoc
            Case 3
                sCadImp = sCadImp & Space(66) & "Tipo de Cuota         : " & psTipoCuota & lnSaltoLinDoc
            Case 4
                sCadImp = sCadImp & Space(66) & "Tasa Compensatoria    : " & psTasaComp & lnSaltoLinDoc
            Case 5
                sCadImp = sCadImp & Space(66) & "Tasa Moratoria        : " & psTasaMor & lnSaltoLinDoc
        End Select
    Next I
    
    'Hitorial de Credito
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & lsNegritaOn & "HISTORIAL DE CREDITO" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & Space(10) & ImpreFormat("FECHA", 12) & ImpreFormat("MONTO", 12) & ImpreFormat("Nro CUOTAS", 12) & ImpreFormat("PLAZO", 12) & ImpreFormat("CUOTA", 12) & ImpreFormat("P. GRACIA", 12) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
    For I = 0 To UBound(MatHitorial) - 1
        '                   Estado                                  Fecha                                   Monto                                   Nro Cuotas                              Plazo                                   Cuota                                   Periodo de Gracia
        sCadImp = sCadImp & Space(5) & ImpreFormat(MatHitorial(I, 0), 11, 0) & ImpreFormat(MatHitorial(I, 1), 16, 0) & ImpreFormat(MatHitorial(I, 2), 12, 0) & ImpreFormat(MatHitorial(I, 3), 16, 0) & ImpreFormat(MatHitorial(I, 4), 12, 0) & ImpreFormat(MatHitorial(I, 5), 18, 0) & ImpreFormat(MatHitorial(I, 6), 12, 0) & lnSaltoLinDoc
    Next I
    
    'DESEMBOLSOS DEL CREDITO
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & lsNegritaOn & "DESEMBOLSOS DEL CREDITO" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & ImpreFormat("FECHA", 18, 0) & ImpreFormat("MONTO", 15, 0) & ImpreFormat("GASTOS", 12, 0) & ImpreFormat("ESTADO", 18, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
    For I = 0 To UBound(MatDesembolsos) - 1
        '                               FECHA                                       MONTO                                     GASTOS                                    ESTADO DEL DESEMBOLSO
        sCadImp = sCadImp & Space(5) & ImpreFormat(MatDesembolsos(I, 0), 19, 0) & ImpreFormat(MatDesembolsos(I, 1), 14, 0) & ImpreFormat(MatDesembolsos(I, 2), 12, 0) & ImpreFormat(MatDesembolsos(I, 3), 18, 0) & lnSaltoLinDoc
    Next I
    
    'CUOTAS PENDIENTES
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & lsNegritaOn & "CUOTAS PENDIENTES" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(9) & ImpreFormat("CUOTA", 10, 0) & ImpreFormat("FECHA", 17, 0) & ImpreFormat("CAPITAL", 14, 0) & ImpreFormat("INTERES", 19, 0) & ImpreFormat("MORA", 12, 0) & ImpreFormat("DIAS ATRASO", 18, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
    For I = 0 To UBound(MatCuotasPend) - 1
        '                   Nro Cuota                                 FECHA DE VENCIMIENTO                      CAPITAL                                   INTERES                                   MORA                                      DIAS ATRASO
        sCadImp = sCadImp & Space(10) & ImpreFormat(MatCuotasPend(I, 0), 8, 0) & ImpreFormat(MatCuotasPend(I, 1), 18, 0) & ImpreFormat(MatCuotasPend(I, 2), 15, 0) & ImpreFormat(MatCuotasPend(I, 3), 18, 0) & ImpreFormat(MatCuotasPend(I, 4), 12, 0) & ImpreFormat(MatCuotasPend(I, 5), 12, 0) & lnSaltoLinDoc
    Next I
            
    'CUOTAS ATRASADAS Y PAGOS
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & lsNegritaOn & "CUOTAS ATRASADAS" & Space(45) & "PAGO A LA FECHA" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(55, "-") & Space(5) & String(55, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "CAPITAL      : " & ImpreFormat(CDbl(MatDeudaVenc(0)), 20) & Space(27) & "CAPITAL      : " & ImpreFormat(CDbl(MatDeudaAFecha(0)), 20) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "INTERES      : " & ImpreFormat(CDbl(MatDeudaVenc(1)), 20) & Space(27) & "INTERES      : " & ImpreFormat(CDbl(MatDeudaAFecha(1)), 20) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "MORA         : " & ImpreFormat(CDbl(MatDeudaVenc(2)), 20) & Space(27) & "MORA         : " & ImpreFormat(CDbl(MatDeudaAFecha(2)), 20) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "GASTOS       : " & ImpreFormat(CDbl(MatDeudaVenc(3)), 20) & Space(27) & "GASTOS       : " & ImpreFormat(CDbl(MatDeudaAFecha(3)), 20) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(55, "-") & Space(5) & String(55, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "TOTAL        : " & ImpreFormat(CDbl(MatDeudaVenc(4)), 20) & Space(27) & "TOTAL        : " & ImpreFormat(CDbl(MatDeudaAFecha(4)), 20) & lnSaltoLinDoc
    
    'sCadImp = sCadImp & Chr$(12)
    'Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecha, pscodusu, "Resumen General de Credito", Space(5))
    sCadImp = sCadImp & Space(5) & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & "OTROS DATOS" & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(70, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(10) & "PROTESTO           : " & psProtesto & Space(10) & ImpreFormat("CUOTA COMODIN    : " & IIf(pnCuotaCom = 1, "SI", "NO"), 50) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(10) & "CARGO AUTOMATICO   : " & psCargoAutom & Space(10) & ImpreFormat("MI VIVIENDA      : " & IIf(pnMiViv = 1, "SI", "NO"), 50) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(10) & "REFINANCIADO       : " & psRefinanciado & Space(10) & ImpreFormat("CALEND. DIN.     : " & IIf(pnCalendDin = 1, "SI", "NO"), 50) & lnSaltoLinDoc
    sCadImp = sCadImp & Chr$(12)
    
    'Imprime garantias
    If IsArray(pMatGarantias) Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecha, psCodUsu, "Resumen General de Credito", psNomCmac, Space(5))
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & Space(5) & " GARANTIAS " & lnSaltoLinDoc
        sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & Space(5) & Space(10) & ImpreFormat("Tipo Garantía", 20) & ImpreFormat("Descripción", 20) & ImpreFormat("Documento", 15) & ImpreFormat("Nro Doc.", 12) & ImpreFormat("Moneda", 12) & ImpreFormat("Monto", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
        For I = 0 To UBound(pMatGarantias) - 1
            sCadImp = sCadImp & Space(15) & ImpreFormat(pMatGarantias(I, 0), 20) & ImpreFormat(pMatGarantias(I, 1), 20) & ImpreFormat(pMatGarantias(I, 2), 15) & ImpreFormat(pMatGarantias(I, 3), 12) & ImpreFormat(pMatGarantias(I, 4), 12) & ImpreFormat(pMatGarantias(I, 5), 12) & lnSaltoLinDoc
        Next I
    End If
    ImprimeConsultaCredito = sCadImp
End Function

Public Function ImprimeRegistroSolicitud(ByVal psCtaCod As String, ByVal psNomAge As String, ByVal psFecha As String, _
    ByVal psCodUsu As String, ByVal psTipoCred As String, ByVal psTipoPers As String, _
    ByVal psCodCli As String, ByVal psNombreCli As String, ByVal psDocNat As String, _
    ByVal psDocJur As String, ByVal psDirecc As String, _
    ByVal psNroHijos As String, ByVal psEdad As String, ByVal psEstadoCiv As String, _
    ByVal psConyugue As String, ByVal psLECony As String, ByVal psRazonSoc As String, _
    ByVal psTipoFuente As String, ByVal psRazonSocDirecc As String, _
    ByVal psActividad As String, ByVal psIniAct As String, ByVal MatRelaCli As Variant, _
    ByVal psAnalista As String, ByVal psCondicion As String, ByVal psDestino As String, ByVal psLineaCred As String, _
    ByVal psFechaSol As String, ByVal MatDatosEstado As Variant, ByVal MatGarantias As Variant, _
    ByVal MatRelacInterna As Variant, ByVal psNomCmac As String) As String
    
Dim sCadImp As String
Dim I As Integer

    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecha, psCodUsu, "REGISTRO DE SOLICITUD DE CREDITO", psNomCmac, Space(5), False)
    sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Credito : " & psCtaCod & Space(40) & psTipoCred & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & lsNegritaOn & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "DATOS DEL TITULAR" & Space(40) & psTipoPers & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Cliente        :" & psCodCli & Space(5) & PstaNombre(psNombreCli) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Doc. Natural   :" & psDocNat & Space(15) & "Doc. Juridico :" & psDocJur & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Direccion      :" & psDirecc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Edad           :" & psEdad & " Años" & Space(30) & "Estado Civil : " & psEstadoCiv & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Conyugue       :" & ImpreFormat(psConyugue, 20, 2, 0) & " L.E:" & psLECony & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & lsNegritaOn & "FUENTE INGRESO" & Space(30) & psTipoFuente & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Razon Social           :" & psRazonSoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Direccion              :" & psRazonSocDirecc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Actividad              :" & psActividad & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Inicio de Actividades  :" & psIniAct & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & lsNegritaOn & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(35) & "CLIENTES RELACIONADOS" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat("CODIGO", 10, 0) & ImpreFormat("NOMBRE", 30, 0) & ImpreFormat("RELACION", 12, 0) & ImpreFormat("DOC.IDENT.", 15, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lsNegritaOff & lnSaltoLinDoc
    For I = 0 To UBound(MatRelaCli) - 1
        sCadImp = sCadImp & lsTab & ImpreFormat(MatRelaCli(I, 0), 10, 0) & ImpreFormat(MatRelaCli(I, 1), 30, 0) & ImpreFormat(MatRelaCli(I, 2), 12, 0) & ImpreFormat(MatRelaCli(I, 3), 15, 0) & lnSaltoLinDoc
    Next I
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(35) & lsNegritaOn & "DATOS DEL CREDITO" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Analista : " & psAnalista & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Destino : " & ImpreFormat(psDestino, 20, 0) & "Condicion : " & psCondicion & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Linea Credito : " & ImpreFormat(psLineaCred, 20, 0) & "Fecha de Solicitud : " & psFechaSol & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & lsNegritaOn & ImpreFormat("Estado", 20, 0) & ImpreFormat("Monto", 16, 0) & ImpreFormat("Cuota", 16, 0) & ImpreFormat("Cuotas", 15, 0) & ImpreFormat("Plazo", 10, 0) & ImpreFormat("Superavit", 15, 0) & lsNegritaOff & lnSaltoLinDoc
    For I = 0 To 1
        sCadImp = sCadImp & lsTab & ImpreFormat(MatDatosEstado(I, 0), 15) & ImpreFormat(MatDatosEstado(I, 1), 15) & ImpreFormat(MatDatosEstado(I, 2), 15) & ImpreFormat(MatDatosEstado(I, 3), 15) & ImpreFormat(MatDatosEstado(I, 4), 7) & ImpreFormat(MatDatosEstado(I, 5), 15) & lnSaltoLinDoc
    Next I
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(35) & lsNegritaOn & "GARANTIAS" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat("Garantia", 30, 0) & ImpreFormat("Total", 20, 0) & ImpreFormat("Gravado", 20, 0) & lnSaltoLinDoc
    For I = 0 To UBound(MatGarantias) - 1
        sCadImp = sCadImp & lsTab & ImpreFormat(MatGarantias(I, 0), 30, 0) & ImpreFormat(MatGarantias(I, 1) & " " & MatGarantias(I, 2), 20, 0) & ImpreFormat(MatGarantias(I, 3) & " " & MatGarantias(I, 4), 20, 0) & lnSaltoLinDoc
    Next I
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(30) & lsNegritaOn & "RELACIONES INTERNAS (Creditos Vigentes)" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat("CLIENTE", 30, 0) & ImpreFormat("RELACION", 15, 0) & ImpreFormat("CREDITO", 20, 0) & ImpreFormat("ATRASO", 15, 0) & lsNegritaOff & lnSaltoLinDoc
    If IsArray(MatRelacInterna) Then
        For I = 0 To UBound(MatRelacInterna) - 1
            sCadImp = sCadImp & lsTab & ImpreFormat(MatRelacInterna(I, 0), 30, 0) & ImpreFormat(MatRelacInterna(I, 1), 15, 0) & ImpreFormat(MatRelacInterna(I, 2), 20, 0) & ImpreFormat(MatRelacInterna(I, 3), 15, 0) & lnSaltoLinDoc
        Next I
    Else
        sCadImp = sCadImp & lnSaltoLinDoc
    End If
    sCadImp = sCadImp & Chr$(12)
    ImprimeRegistroSolicitud = sCadImp
End Function

Public Function ImprimeSolicitud(ByVal psCtaCod As String, ByVal psNomAge As String, ByVal psFecSis As String, _
    ByVal psCodUsu As String, Optional ByVal psNomCmact As String = "") As String
    
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String

    On Error GoTo ErrorImprimeSolicitud
    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaDatosSolicitudCredito(psCtaCod)
    Set oDCred = Nothing
    
    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "SOLICITUD DE CREDITO", psNomCmact, lsTab, False)
    If Not R.BOF And Not R.EOF Then
        sCadImp = sCadImp & lsTab & lsNegritaOn & "Credito : " & psCtaCod & Space(30) & R!cTipoCred & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "DATOS DEL TITULAR" & Space(30) & R!cPersoneria & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(90, "-") & lsNegritaOff & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Cliente        : " & R!cPersCod & Space(5) & PstaNombre(R!cPersNombre) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Doc. Natural   : " & ImpreFormat(IIf(IsNull(R!DNI), "", R!DNI), 12, 0) & "Doc. Juridico    : " & ImpreFormat(IIf(IsNull(R!RUC), "", R!RUC), 12, 0) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Direccion      : " & R!cPersDireccDomicilio & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Nro Hijos      : " & Trim(Str(IIf(IsNull(R!nPersNatHijos), 0, R!nPersNatHijos))) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Edad           : " & Format(Year(CDate(psFecSis)) - Year(IIf(IsNull(R!dPersNacCreac), CDate(psFecSis), R!dPersNacCreac)), "#0") & " Años " & Space(20) & "Estado Civil : " & R!cEstadoCivil & lnSaltoLinDoc & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & lsNegritaOn & "FUENTE DE INGRESO" & Space(30) & IIf(IsNull(R!cFteInd), "DEPENDIENTE", "INDEPENDIENTE") & lsNegritaOff & lnSaltoLinDoc & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Razon Social           : " & R!cRazonSocial & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Direccion              : " & R!cDireccRazon & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Actividad              : " & R!cCIIUdescripcion & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Inicio de Actividades  : " & Format(R!dPersNacCreac, "dd/mm/yyyy") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(90, "-") & lsNegritaOff & lnSaltoLinDoc & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & lsNegritaOn & "DATOS PRINCIPALES DEL CREDITO" & lsNegritaOff & lnSaltoLinDoc & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Analista Responsable  : " & R!cAnalista & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Destino del Credito   : " & R!cDestino & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Condicion del Credito : " & R!cCondicion & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Fecha de Asignacion   : " & Format(R!dFecAsig, "dd/mm/yyyy") & lnSaltoLinDoc
        If Not IsNull(R!cNomInst) Then
            sCadImp = sCadImp & lsTab & "Institucion           : " & R!cNomInst & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & "Codigo Modular        : " & IIf(IsNull(R!cCodModular), "", R!cCodModular) & lnSaltoLinDoc
        End If
        sCadImp = sCadImp & lsTab & lsNegritaOn & "Monto Solicitado      : " & Format(R!nMontoSol, "#0.00") & Space(15) & R!cMonedaSol & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Cuotas Solicitadas    : " & Trim(Str(R!nCuotasSol)) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Plazo Solicitado      : " & Trim(Str(R!nPlazoSol)) & " Dia(s)" & lsNegritaOff & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & Chr$(12)
    End If
    R.Close
    Set R = Nothing
    ImprimeSolicitud = sCadImp
    Exit Function

ErrorImprimeSolicitud:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Public Function ImprimeResumenComite(ByVal psCtaCod As String, ByVal psNomAge As String, ByVal psFecSis As String, _
    ByVal psCodUsu As String, ByVal psNomCmac As String) As String
Dim sSQL As String
Dim oDCredDoc As DCredDoc
Dim oDCred As DCredito
Dim oDGarantia As DGarantia
Dim R As ADODB.Recordset
Dim RGarantCred As ADODB.Recordset
Dim RRelaCred As ADODB.Recordset
Dim REstadosCred As ADODB.Recordset
Dim sCadImp As String
    Set oDCredDoc = New DCredDoc
    Set R = oDCredDoc.RecuperaDatosResumenComite(psCtaCod)
    Set oDCredDoc = Nothing
    Set oDCred = New DCredito
    Set RRelaCred = oDCred.RecuperaRelacPers(psCtaCod)
    Set REstadosCred = oDCred.RecuperaEstadosdelCredito(psCtaCod, True, True, False)
    Set oDCred = Nothing
    Set oDGarantia = New DGarantia
    Set RGarantCred = oDGarantia.RecuperaGarantiaCredito(psCtaCod)
    Set oDGarantia = Nothing
        
    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "RESUMEN DE COMITE", psNomCmac, lsTab, False)
    sCadImp = sCadImp & lsTab & lsNegritaOn & "Credito : " & psCtaCod & Space(30) & R!cTipoCred & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "DATOS DEL TITULAR" & Space(30) & R!cPersoneria & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Cliente        : " & R!cPersCod & Space(2) & PstaNombre(R!cPersNombre) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Doc. Natural   : " & R!DNI & Space(30) & "Doc. Juridico : " & R!RUC & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Direccion      : " & R!cPersDireccDomicilio & lnSaltoLinDoc
    If IsNull(R!dPersNacCreac) Then
        sCadImp = sCadImp & lsTab & "Edad           : "
    Else
        sCadImp = sCadImp & lsTab & "Edad           : " & Format(Year(CDate(psFecSis)) - Year(R!dPersNacCreac), "#0")
    End If
    sCadImp = sCadImp & Space(30) & "Estado Civil : " & R!cEstadoCivil & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Nro Hijos : " & Trim(Str(IIf(IsNull(R!nPersNatHijos), 0, R!nPersNatHijos))) & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOn
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(35) & "CLIENTES RELACIONADOS" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat(" CODIGO ", 15, 0) & ImpreFormat("NOMBRE", 40, 0) & ImpreFormat("RELACION", 20, 0) & ImpreFormat("DOC. IDENT", 15, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOff
    Do While Not RRelaCred.EOF
        If RRelaCred!nConsValor <> gColRelPersTitular Then
            sCadImp = sCadImp & lsTab & ImpreFormat(RRelaCred!cPersCod, 15, 0) & ImpreFormat(RRelaCred!cPersNombre, 40, 0) & ImpreFormat(RRelaCred!cConsDescripcion, 20, 0) & ImpreFormat(RRelaCred!DNI, 15, 0) & lnSaltoLinDoc
        End If
        RRelaCred.MoveNext
    Loop
    
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOn
    sCadImp = sCadImp & lsTab & Space(30) & "DATOS DEL CREDITO" & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOff
    sCadImp = sCadImp & lsTab & ImpreFormat("ESTADO", 26, 0) & ImpreFormat("MONTO", 18, 0) & ImpreFormat("CUOTA", 15, 0) & ImpreFormat("NRO CUOTAS", 15, 0) & ImpreFormat("PLAZO", 15, 0) & lnSaltoLinDoc
    Do While Not REstadosCred.EOF
        sCadImp = sCadImp & lsTab & ImpreFormat(REstadosCred!cEstado, 15, 0) & ImpreFormat(REstadosCred!nMonto, 15, 2) & ImpreFormat(IIf(IsNull(REstadosCred!nMontoCuota), 0, REstadosCred!nMontoCuota), 15, 2) & ImpreFormat(REstadosCred!nCuotas, 15, 0) & ImpreFormat(IIf(REstadosCred!nPlazo = 0, 30, REstadosCred!nPlazo), 15, 0) & " dias" & lnSaltoLinDoc
        REstadosCred.MoveNext
    Loop
    
    sCadImp = sCadImp & lsTab & "DESTINO : " & R!cDestino & Space(20) & "LINEA DE CREDITO : " & R!cLinea & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "MONEDA : " & IIf(Mid(psCtaCod, 9, 1) = "1", "SOLES", "DOLARES") & lnSaltoLinDoc
    sCadImp = sCadImp & "GARANTIA : " & ImpreFormat(RGarantCred!cTpoGarantia, 20, 0) & " Garant. Total: " & ImpreFormat(RGarantCred!nPorGravar, 10, 2) & " Garant. Cred. : " & ImpreFormat(RGarantCred!nGravado, 15, 2) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOn
    sCadImp = sCadImp & Space(35) & "RESOLUCION DE COMITE" & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOff
    sCadImp = sCadImp & lsTab & Space(15) & "MONTO" & Space(5) & "# CUOTAS " & Space(10) & "MONTO CUOTA" & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "APROBADO : " & ".... ......................    .......................      ....................." & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Tipo de Contrato :      Fondo : ................  Garantias : ................................" & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(1) & "____________________________      ______________________________    _________________________" & lnSaltoLinDoc
    sCadImp = sCadImp & Chr$(12)
    
    R.Close
    Set R = Nothing
    RRelaCred.Close
    Set RRelaCred = Nothing
    REstadosCred.Close
    Set REstadosCred = Nothing
    RGarantCred.Close
    Set RGarantCred = Nothing
    
    ImprimeResumenComite = sCadImp
End Function

Public Function ImprimeComprobanteDesembolso(ByVal psCtaCod As String, ByVal psNomAge As String, ByVal psFecSis As String, _
    ByVal psCodUsu As String, ByVal psNomCmac As String) As String
    
Dim oCredDoc As DCredDoc
Dim oCred As DCredito
Dim oCalend As Dcalendario
Dim R As ADODB.Recordset
Dim REstadosCred As ADODB.Recordset
Dim RDesemb As ADODB.Recordset
Dim sCadImp As String

    Set oCredDoc = New DCredDoc
    Set R = oCredDoc.RecuperaDatosDocDesembolso(psCtaCod)
    Set oCredDoc = Nothing
    Set oCred = New DCredito
    Set REstadosCred = oCred.RecuperaEstadosdelCredito(psCtaCod, True, True, True)
    Set oCred = Nothing
    Set oCalend = New Dcalendario
    
    Set oCalend = New Dcalendario
    Set RDesemb = oCalend.RecuperaCalendarioDesemb(psCtaCod)
    Set oCalend = Nothing
    
    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "COMPROBANTE DE DESEMBOLSO", psNomCmac, lsTab, False)
    sCadImp = sCadImp & lsTab & R!cTipoCred & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Credito             : " & ImpreFormat(psCtaCod, 20, 0) & "Analista  : " & R!cAnalista & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Fecha de Aprobacion : " & ImpreFormat(Format(R!dFecApr, "dd/mm/yyyy"), 20, 0) & "Codigo de Cliente : " & R!cPersCod & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOn
    sCadImp = sCadImp & lsTab & "Fecha de Desembolso : " & ImpreFormat(Format(R!dFecDesemb, "dd/mm/yyyy"), 20, 0) & "Nombre            : " & PstaNombre(R!cPersNombre) & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOff
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat("ESTADO", 20, 0) & ImpreFormat("MONTO", 15, 0) & ImpreFormat("CUOTAS", 18, 0) & ImpreFormat("PLAZO", 15, 0) & ImpreFormat("CUOTA", 15, 0) & lnSaltoLinDoc
    Do While Not REstadosCred.EOF
        sCadImp = sCadImp & lsTab & ImpreFormat(REstadosCred!cEstado, 15, 0) & ImpreFormat(REstadosCred!nMonto, 10, 0) & ImpreFormat(REstadosCred!nCuotas, 15, 0) & ImpreFormat(REstadosCred!nPlazo, 15, 0) & ImpreFormat(IIf(IsNull(REstadosCred!nMontoCuota), 0, REstadosCred!nMontoCuota), 15, 2) & lnSaltoLinDoc
        REstadosCred.MoveNext
    Loop
    REstadosCred.Close
    Set REstadosCred = Nothing
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "APODERADO         : " & PstaNombre(R!cApoderado) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "MONEDA            : " & R!cMoneda & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "LINEA DE CREDITO  : " & R!cLinea & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "INTERES           : " & Format(R!nTasaInteres, "#0.00") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "TIPO DE DESEMBOLSO: " & R!cTipoDesemb & lnSaltoLinDoc
    
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(25) & "DESEMBOLSOS EFECTUADOS" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(15) & ImpreFormat("FECHA", 15, 0) & ImpreFormat("CUOTA", 15, 0) & ImpreFormat("MONTO", 15, 0) & ImpreFormat("GASTOS", 15, 0) & ImpreFormat("ESTADO", 15, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    Do While Not RDesemb.EOF
        sCadImp = sCadImp & lsTab & Space(15) & ImpreFormat(Format(RDesemb!dVenc, "dd/mm/yyyy"), 15, 0) & ImpreFormat(Trim(Str(RDesemb!nCuota)), 5, 0) & ImpreFormat(RDesemb!nCapital, 15, 2) & ImpreFormat(RDesemb!NGasto, 10, 2) & Space(10) & IIf(RDesemb!nColocCalendEstado = gColocCalendEstadoPendiente, "PENDIENTE", "CANCELADO") & lnSaltoLinDoc
        RDesemb.MoveNext
    Loop
    RDesemb.Close
    Set RDesemb = Nothing
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "________________________________   ______________________________    _________________________" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & PstaNombre(R!cPersNombre) & Space(5) & PstaNombre(R!cAnalista) & Space(15) & "CAJA" & lnSaltoLinDoc
    R.Close
    Set R = Nothing
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & Chr$(12)
    ImprimeComprobanteDesembolso = sCadImp
    
End Function

Public Function ImprimePlandePagos(ByVal psCtaCod As String, ByVal psNomAge As String, _
    ByVal psFecSis As String, ByVal psCodUsu As String, ByVal pnMontoPrestamo As Double, _
    ByVal pbMiViv As Boolean, Optional ByVal pbCalendParalelo As Boolean = False, Optional ByVal psNomCmac As String = "", _
    Optional ByVal pbComodin As Integer = 0, Optional ByVal pbCalendDin As Integer = 0) As String

Dim R As ADODB.Recordset
Dim oCredDoc As DCredDoc
Dim oCalend As Dcalendario
Dim oNCred As NCredito
Dim MatCalend As Variant
Dim sCadImp As String
Dim I As Integer
Dim nIntereses As Double
Dim nCapital As Double
Dim oParam As DParametro
Dim nTramoNoConsPorcen As Double
Dim RDesPar As ADODB.Recordset

    Set oCredDoc = New DCredDoc
    Set R = oCredDoc.RecuperaDatosDocPlanPagos(psCtaCod)
    Set oCredDoc = Nothing
    
    If pbMiViv Then
        Set oParam = New DParametro
        nTramoNoConsPorcen = oParam.RecuperaValorParametro(gColocMiVivTramo)
        Set oParam = Nothing
        If Not pbCalendParalelo Then
            pnMontoPrestamo = pnMontoPrestamo * ((100 - nTramoNoConsPorcen) / 100)
            pnMontoPrestamo = Format(pnMontoPrestamo, "#0.00")
        Else
            pnMontoPrestamo = pnMontoPrestamo * (nTramoNoConsPorcen / 100)
            pnMontoPrestamo = Format(pnMontoPrestamo, "#0.00")
        End If
    End If
    
    Set oCalend = New Dcalendario
    Set RDesPar = oCalend.RecuperaCalendarioDesemb(psCtaCod)
    Set oCalend = Nothing
    
    Set oNCred = New NCredito
    MatCalend = oNCred.RecuperaMatrizCalendarioInicial(psCtaCod, pnMontoPrestamo, pbCalendParalelo, IIf(pbCalendParalelo, True, False))
    Set oNCred = Nothing
    
    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "PLAN DE PAGOS", psNomCmac, lsTab, False)
    sCadImp = sCadImp & lsTab & "DUPLICADO" & Chr$(10)
    sCadImp = sCadImp & lsTab & "Credito       : " & psCtaCod & Space(10) & "Cliente : " & PstaNombre(R!cPersNombre) & Chr$(10) & Chr$(10)
    sCadImp = sCadImp & lsTab & "Analista      : " & PstaNombre(R!cAnalista) & Chr$(10)
    sCadImp = sCadImp & lsTab & "Tipo de Cuota : " & ImpreFormat(R!cTipoCuota, 30, 0) & "Cuota : " & Format(R!nMontoCuota, "#0.00") & Chr$(10)
    sCadImp = sCadImp & lsTab & "Interes       : " & ImpreFormat(Format(R!nTasaInteres, "#0.00"), 30, 0) & "Plazo : " & Trim(Str(R!nPlazo)) & Chr$(10)
    sCadImp = sCadImp & lsTab & "Moneda        : " & ImpreFormat(R!cMoneda, 30, 0) & "Vigencia : " & Format(R!dFecVig, "ddd, d mmm yyyy") & Chr$(10)
    sCadImp = sCadImp & lsTab & "Monto         : " & Format(R!nMonto, "#0.00") & Chr$(10)
    If pbComodin Then
        sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & ImpreFormat("Cuota Comodin", 30, 0)
    ElseIf pbMiViv Then
        sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & ImpreFormat("Mi Vivienda", 30, 0)
    Else
        sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & ImpreFormat("Normal", 30, 0)
    End If
    
    If pbCalendDin Then
        sCadImp = sCadImp & lsTab & "Calend. Din.  :" & ImpreFormat("SI", 30, 0)
    Else
        sCadImp = sCadImp & lsTab & "Calend. Din.  :" & ImpreFormat("NO", 30, 0)
    End If
    sCadImp = sCadImp & Chr$(10)
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & Chr$(10)
    
    If RDesPar.RecordCount >= 2 Then
        sCadImp = sCadImp & lsTab & "DESEMBOLSOS PARCIALES" & Chr$(10) & Chr$(10)
        sCadImp = sCadImp & lsTab & String(80, "-") & Chr$(10)
        sCadImp = sCadImp & lsTab & ImpreFormat("FECHA", 16, 0) & ImpreFormat("DESEMBOLSO", 10, 0) & ImpreFormat("ESTADO", 10, 2) & Chr$(10)
        sCadImp = sCadImp & lsTab & String(80, "-") & Chr$(10)
        Do While Not RDesPar.EOF
            sCadImp = sCadImp & lsTab & ImpreFormat(Format(RDesPar!dVenc, "dd/mm/yyyy"), 14, 0) & ImpreFormat(RDesPar!nCapital, 10, 2, True) & ImpreFormat(IIf(RDesPar!nColocCalendEstado = 1, "DESEMBOLSADO", "PENDIENTE"), 12, 2, True) & Chr$(10)
            RDesPar.MoveNext
        Loop
    End If
    RDesPar.Close
    
    sCadImp = sCadImp & Chr$(10)
    
    sCadImp = sCadImp & lsTab & ImpreFormat("Fecha", 15, 0) & ImpreFormat("No Cuota", 10, 0) & ImpreFormat("Cuota", 12, 0) & ImpreFormat("Capital", 12, 0) & ImpreFormat("Interes", 12, 0) & ImpreFormat("Int.Gracia", 12, 0) & ImpreFormat("Gastos", 8, 0) & ImpreFormat("Saldo Cap.", 12, 0) & Chr$(10)
    sCadImp = sCadImp & lsTab & String(90, "-") & Chr$(10)
    nIntereses = 0
    nCapital = 0
    For I = 0 To UBound(MatCalend) - 1
        sCadImp = sCadImp & lsTab & ImpreFormat(Format(CDate(MatCalend(I, 0)), "ddd, d mmm yyyy"), 17, 0) & ImpreFormat(CInt(MatCalend(I, 1)), 4, 0) & ImpreFormat(CDbl(MatCalend(I, 3)) _
                                                                                                                                                                + CDbl(MatCalend(I, 4)) _
                                                                                                                                                                + CDbl(MatCalend(I, 5)) _
                                                                                                                                                                + CDbl(MatCalend(I, 6)) _
                                                                                                                                                                + CDbl(MatCalend(I, 7)) _
                                                                                                                                                                + CDbl(MatCalend(I, 8)) + CDbl(MatCalend(I, 9)), 10, 2) & ImpreFormat(CDbl(MatCalend(I, 3)), 9, 2) & ImpreFormat(CDbl(MatCalend(I, 4)), 8, 2) & ImpreFormat(CDbl(MatCalend(I, 5)), 8, 2) & ImpreFormat(CDbl(MatCalend(I, 9)), 8, 2) & ImpreFormat(CDbl(MatCalend(I, 10)), 10, 2) & Chr$(10)
        nIntereses = nIntereses + CDbl(MatCalend(I, 4))
        nIntereses = CDbl(Format(nIntereses, "#0.00"))
        nCapital = nCapital + CDbl(MatCalend(I, 3))
        nCapital = CDbl(Format(nCapital, "#0.00"))
        If I <> 0 And I Mod 45 = 0 Then
            sCadImp = sCadImp & Chr$(12)
            Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "PLAN DE PAGOS", ImprimePlandePagos, lsTab, False)
            sCadImp = sCadImp & lsTab & "DUPLICADO" & Chr$(10)
            sCadImp = sCadImp & lsTab & "Credito       : " & psCtaCod & Space(10) & "Cliente : " & PstaNombre(R!cPersNombre) & Chr$(10) & Chr$(10)
            sCadImp = sCadImp & lsTab & "Analista      : " & PstaNombre(R!cAnalista) & Chr$(10)
            sCadImp = sCadImp & lsTab & "Tipo de Cuota : " & ImpreFormat(R!cTipoCuota, 20, 0) & "Cuota : " & Format(R!nMontoCuota, "#0.00") & Chr$(10)
            sCadImp = sCadImp & lsTab & "Interes       : " & ImpreFormat(Format(R!nTasaInteres, "#0.00"), 20, 0) & "Plazo : " & Trim(Str(R!nPlazo)) & Chr$(10)
            sCadImp = sCadImp & lsTab & "Moneda        : " & ImpreFormat(R!cMoneda, 20, 0) & "Vigencia : " & Format(R!dFecVig, "ddd, d mmm yyyy") & Chr$(10)
            sCadImp = sCadImp & lsTab & "Monto         : " & Format(R!nMonto, "#0.00") & Chr$(10)
            If pbComodin Then
                sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & ImpreFormat("Cuota Comodin", 30, 0)
            ElseIf pbMiViv Then
                sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & ImpreFormat("Mi Vivienda", 30, 0)
            Else
                sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & ImpreFormat("Normal", 30, 0)
            End If
            
            If pbCalendDin Then
                sCadImp = sCadImp & lsTab & "Calend. Din.  :" & ImpreFormat("SI", 30, 0)
            Else
                sCadImp = sCadImp & lsTab & "Calend. Din.  :" & ImpreFormat("NO", 30, 0)
            End If
            sCadImp = sCadImp & Chr$(10)
            sCadImp = sCadImp & Chr$(10)
            sCadImp = sCadImp & lsTab & ImpreFormat("Fecha", 15, 0) & ImpreFormat("No Cuota", 10, 0) & ImpreFormat("Cuota", 12, 0) & ImpreFormat("Capital", 12, 0) & ImpreFormat("Interes", 12, 0) & ImpreFormat("Int.Gracia", 12, 0) & ImpreFormat("Gastos", 8, 0) & ImpreFormat("Saldo Cap.", 12, 0) & Chr$(10)
            sCadImp = sCadImp & lsTab & String(90, "-") & Chr$(10)
        End If
    Next I
    
    sCadImp = sCadImp & lsTab & String(90, "-") & Chr$(10)
    sCadImp = sCadImp & lsTab & "Totales : " & ImpreFormat(nCapital, 32, 2) & ImpreFormat(nIntereses, 7, 2) & ImpreFormat(nCapital + nIntereses, 32, 2) & Chr$(10)
    sCadImp = sCadImp & lsTab & String(90, "-") & Chr$(10)
    sCadImp = sCadImp & Chr$(27) & Chr$(107) & Chr$(2) & Chr$(27) & Chr$(83) & Chr$(1)
    Set oNCred = New NCredito
    sCadImp = sCadImp & lsTab & "Tasa Efectiva Anual : " & Format(oNCred.TasaIntPerDias(R!nTasaInteres, 360) * 100, "#0.00") & "%" & Chr$(10)
    Set oNCred = Nothing
    sCadImp = sCadImp & Chr$(27) & Chr$(84) & Chr$(1)
    sCadImp = sCadImp & Chr$(27) + Chr$(107) + Chr$(1)
    
    If CInt(Mid(psCtaCod, 6, 3)) = gColConsuDctoPlan Then
        sCadImp = sCadImp & lsTab & "NOTA : " & Chr$(10)
        sCadImp = sCadImp & lsTab & "Si no se ha Realizado su Descuento por Planilla" & Chr$(10)
        sCadImp = sCadImp & lsTab & "Acerquese a nuestras Agencias a Cancelar y Evite el cobro de Mora" & Chr$(10)
    End If
    If CInt(Mid(psCtaCod, 6, 3)) = gColHipoMiVivienda Or CInt(Mid(psCtaCod, 6, 3)) = gColHipoCaja Then
        sCadImp = sCadImp & lsTab & "NOTA : " & Chr$(10)
        sCadImp = sCadImp & lsTab & "Todas los Cuotas incluyen Gastos de Portes y Seguros." & Chr$(10)
    End If
    sCadImp = sCadImp & Chr$(12)
    ImprimePlandePagos = sCadImp
End Function

Public Function ImprimePagoLote(ByVal MatPagos As Variant, ByVal MatCalend As Variant, ByVal MatCalendDistrib As Variant, _
    ByVal psNomAge As String, ByVal psFecha As String, ByVal psCodUsu As String, ByVal psInstitucion As String, _
    ByVal psMoneda As String, ByVal psNomCmac As String) As String
    
Dim I As Integer
Dim sCadImp As String
Dim oNCredito As NCredito
Dim nMontoPag As Double
Dim nAcumTot As Double
Dim nAcumCapPag As Double
Dim nAcumIntComPag As Double
Dim nAcumIntMorPag As Double
Dim nAcumGastosPag As Double

    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecha, psCodUsu, "Reporte de Pagos de Prestamos Personales", psNomCmac, lsTab, True)
    sCadImp = sCadImp & lsTab & "INSTITUCION : " & psInstitucion & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "MONEDA : " & psMoneda & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 25, 0) & ImpreFormat("TITULAR", 30, 0) & ImpreFormat("CUOTA QUE PAGO", 20, 0) & ImpreFormat("MONTO", 15, 0) & ImpreFormat("CAPITAL", 15, 0) & ImpreFormat("INTERES", 12, 0) & ImpreFormat("MORA", 8, 0) & ImpreFormat("GASTOS", 8, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
    
    Set oNCredito = New NCredito
    nAcumTot = 0
    nAcumCapPag = 0
    nAcumIntComPag = 0
    nAcumIntMorPag = 0
    nAcumGastosPag = 0
    For I = 0 To UBound(MatPagos) - 1
        nMontoPag = CDbl(MatPagos(I, 1))
        sCadImp = sCadImp & lsTab & ImpreFormat(MatPagos(I, 0), 21) & ImpreFormat(MatPagos(I, 3), 30) & ImpreFormat(oNCredito.MatrizCuotasPagadas(MatCalendDistrib(I)), 10) _
            & ImpreFormat(nMontoPag, 12) & ImpreFormat(oNCredito.MatrizCapitalPagado(MatCalendDistrib(I)), 12) & ImpreFormat(oNCredito.MatrizIntCompPagado(MatCalendDistrib(I)) + oNCredito.MatrizIntCompVencPagado(MatCalendDistrib(I)) + oNCredito.MatrizIntGraciaPagado(MatCalendDistrib(I)) + oNCredito.MatrizIntReprogPag(MatCalendDistrib(I)) + oNCredito.MatrizIntSuspensoPag(MatCalendDistrib(I)), 10) & ImpreFormat(oNCredito.MatrizIntMorPagado(MatCalendDistrib(I)), 8) & ImpreFormat(oNCredito.MatrizGastoPag(MatCalendDistrib(I)), 6) & lnSaltoLinDoc
        nAcumTot = nAcumTot + nMontoPag
        nAcumCapPag = nAcumCapPag + oNCredito.MatrizCapitalPagado(MatCalendDistrib(I))
        nAcumIntComPag = nAcumIntComPag + oNCredito.MatrizIntCompPagado(MatCalendDistrib(I)) + oNCredito.MatrizIntCompVencPagado(MatCalendDistrib(I)) + oNCredito.MatrizIntGraciaPagado(MatCalendDistrib(I)) + oNCredito.MatrizIntReprogPag(MatCalendDistrib(I)) + oNCredito.MatrizIntSuspensoPag(MatCalendDistrib(I))
        nAcumIntMorPag = nAcumIntMorPag + oNCredito.MatrizIntMorPagado(MatCalendDistrib(I))
        nAcumGastosPag = nAcumGastosPag + oNCredito.MatrizGastoPag(MatCalendDistrib(I))
    Next I
    Set oNCredito = Nothing
    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & " PAGO TOTAL     : " & ImpreFormat(nAcumTot, 10, 2) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & " CAPITAL PAGADO : " & ImpreFormat(nAcumCapPag, 10, 2) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & " INTERES PAGADO : " & ImpreFormat(nAcumIntComPag, 10, 2) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & " MORA PAGADA    : " & ImpreFormat(nAcumIntMorPag, 10, 2) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & " GASTO PAGADO   : " & ImpreFormat(nAcumGastosPag, 10, 2) & lnSaltoLinDoc
    sCadImp = sCadImp & Chr$(12)
    
    ImprimePagoLote = sCadImp
    
End Function

Public Function ImprimeBoletaAhorro(ByVal psTit As String, ByVal psText As String, ByVal psCodOpe As String, ByVal pnMonto As String, _
        ByVal psCliente As String, ByVal psCodCta As String, ByVal psNumDoc As String, ByVal pnSaldo As Double, _
        ByVal pnInteresA As String, NomDoc As String, ByVal pnNumExt As Long, ByVal pnSaldoC As Double, _
        Optional pbOpSaldoC As Boolean = True, Optional pbSaldoInt As Boolean = True, Optional psNumDias As String = "----", _
        Optional psNomAgeRem As String = "", Optional psCodUsuRem As String = "", Optional pbCuenta As Boolean = False, _
        Optional psComCmac As String = "XXX", Optional psLin3 As String = "XXX", Optional psTexto As String = "XXX", _
        Optional pdFecSis As Date, Optional sNomAge As String = "", Optional sCodUser As String = "", _
        Optional sLpt As String = "", Optional ByVal psCodCMAC As String) As String

Dim nFicSal As Integer
Dim sFecha As String
Dim sHora As String
Dim sSep As Integer
Dim sIni As Integer
Dim sMonto As String
Dim sSDisp As String
Dim sIntAcum As String
Dim sMax As Integer
Dim sAux As Integer
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lsNroExt As String
Dim lnTope As Integer
Dim lsSaldoC As String
Dim lsCadArg As String
Dim lsInteres As String
Dim sCad As String
Dim lnCliAux As Long
Dim lsCliAux1 As String
Dim lsCliAux2 As String

Dim lnChq As Long
Dim lsChqAux1 As String
Dim lsChqAux2 As String
Dim lsNomAge As String

Dim lnNumLinCmac As Integer

Dim lsMensaje As String * 39

''  quitar este comentario

        Dim clsGen As DGeneral
        Set clsGen = New DGeneral
 

ETIQ:

On Error GoTo ERROR

lnTope = 0 '6 'Tope de lineas en Boleta

lsNegritaOn = Chr$(27) + Chr$(71)
lsNegritaOff = Chr$(27) + Chr$(72)

lsNroExt = Str(pnNumExt)

nFicSal = FreeFile
If sLpt <> "" Then
    Open sLpt For Output As nFicSal
End If

 ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) & Chr$(64)

ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) & Chr$(50)    'espaciamiento lineas 1/6 pulg.
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) & Chr$(67) & Chr$(22)   'Longitud de página a 22 líneas'
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) & Chr$(77)    'Tamaño 10 cpi
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) + Chr$(18)  ' cancela condensada
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(27) + Chr$(72)  ' desactiva negrita

sSep = 15
sIni = 1
sMax = 33
sAux = 5

sFecha = Format$(pdFecSis, "dd/mm/yyyy")
sHora = Format$(Time, "hh:mm:ss")
sMonto = Format$(pnMonto, "#,##0.00")
sSDisp = Format$(pnSaldo, "#,##0.00")
lsSaldoC = Format$(pnSaldoC, "#,##0.00")

lsNomAge = sNomAge

ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(10)
ImprimeBoletaAhorro = ImprimeBoletaAhorro & Chr$(10)

'Print #nFicSal, Chr$(10);
If psCodCMAC = "102" Then  ' CAJA LIMA
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOn  'Activa Negrita
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "CM - AHORRO" & Chr$(10)
        
        If Mid(psCodCta, 9, 1) = 1 Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Space(sSep) & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & Chr$(10)
        End If
        
        If psNomAgeRem = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Ag.Rem: " & Trim(psNomAgeRem) & Space(sAux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))) & Chr$(10)
        End If
        
        If psComCmac = "XXX" Then
            If psLin3 = "XXX" Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOff 'Desactiva Negrita
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & psLin3 & Space(sAux + sSep + sMax - Len(psLin3)) & lsNegritaOff   'Desactiva Negrita
                lnNumLinCmac = 1
            End If
            lnNumLinCmac = 0
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "NroDocCmac:" & psComCmac & Space(sAux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & lsNegritaOff & Chr$(10)   'Desactiva Negrita
            lnNumLinCmac = 1
        End If
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(sAux + sSep - 6) & Chr$(10)
        
        lnCliAux = InStr(1, psCliente, "*", vbTextCompare)
        
        If lnCliAux = 0 Then
            If sAux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + sAux)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(psCliente) & Space(sAux + sMax + sSep - Len(psCliente)) & Chr$(10)
        Else
            lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
            lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
            
            If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + sAux)
            If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + sAux)
            
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(lsCliAux1) & Space(sAux + sMax + sSep - Len(lsCliAux1)) & Chr$(10)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(lsCliAux2) & Space(sAux + sMax + sSep - Len(lsCliAux2)) & Chr$(10)
            
            lnCliAux = 1
        End If
        
        If pbSaldoInt Or pbCuenta Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Cuenta:" & psCodCta & Space(8 + sSep + sAux) & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        psTit = Trim(psTit)
        psTit = CentrarCadena(psTit, 28)
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOn  'Activa Negrita & chr$(10)
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni + 1) & "-----" & psTit & "-----" & Chr$(10)
        
        lnChq = InStr(1, psText, "*", vbTextCompare)
        
        If psTexto = "XXX" Then
            If lnChq = 0 Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(Trim(Mid(psText, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(psText, 1, 28))) - Len(sMonto)) & sMonto & Space(-1 + sSep) & Chr$(10)
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
            Else
                lsChqAux1 = (Mid(psText, 1, lnChq - 1))
                lsChqAux2 = (Mid(psText, lnChq + 1))

                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Chr$(10)
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Chr$(10)
            End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(Trim(psTexto)) & Space(sAux + sSep + sMax - Len(Trim(psTexto))) & Chr$(10)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOff  'Desactiva Negrita
        
        If pbSaldoInt Then
            'If MsgBox("Desea Imprimir el Saldos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Space(-1 + sSep) & Chr$(10)
                If pbOpSaldoC Then
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Space(-1 + sSep) & Chr$(10)
                Else
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
                End If
            'Else
            '    Print #nFicSal, ""
            '    Print #nFicSal, ""
            '    pbSaldoInt = False
            'End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        lsInteres = pnInteresA
        
        If pbSaldoInt Then
            If lsInteres <> "No Valido" Then
                lsInteres = Format(lsInteres, "#,##0.00")
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & Space(-1 + sSep) & Chr$(10)
            End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        If Not psNumDoc = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & Space(-1 + sSep) & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        If Not psNumDias = "----" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Nro Dias Interes" & Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)) & psNumDias & Space(-1 + sSep) & Chr$(10)
            lnTope = 4 - lnCliAux
        Else
            lnTope = 3 - lnCliAux
        End If
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "---------------------------------------" & Space(-1 + sSep) & Chr$(10)
        If psCodUsuRem = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(sCodUser) & Chr$(10)  ''& Space(29 + sSep + sAux)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Loc/Rem" & Space(sMax + sAux - Len("Loc/Rem") - 1 - 8) & ImpreCarEsp(sCodUser) & "/" & ImpreCarEsp(psCodUsuRem) & Space(sSep) & Chr$(10)
        End If
        'Dim clsGen As DGeneral
        'Set clsGen = New DGeneral
        lsMensaje = clsGen.GetMensajeBoletas(psCodCta)
        Set clsGen = Nothing
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & lsNegritaOn & ImpreCarEsp(lsMensaje) & Space(-1 + sSep) & ImpreCarEsp(lsMensaje) & lsNegritaOff
        
        lnNumLinCmac = lnNumLinCmac + 1
        
        For sAux = 1 To (lnTope - lnNumLinCmac)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        Next sAux

Else   '' Caja Trujillo
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOn  'Activa Negrita
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "CM - AHORRO" & Space(19 + sSep + sAux) & "CMACT - AHORRO" & Chr$(10)
        
        If Mid(psCodCta, 9, 1) = 1 Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Space(sSep) & Trim(sNomAge) & "-SOLES" & Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & Trim(sNomAge) & "-DOLARES" & Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt & Chr$(10)
        End If
        
        If psNomAgeRem = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Ag.Rem: " & Trim(psNomAgeRem) & Space(sAux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))) & "Ag.Rem: " & Trim(psNomAgeRem) & Chr$(10)
        End If
        
        If psComCmac = "XXX" Then
            If psLin3 = "XXX" Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOff 'Desactiva Negrita
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & psLin3 & Space(sAux + sSep + sMax - Len(psLin3)) & psLin3 & lsNegritaOff  'Desactiva Negrita
                lnNumLinCmac = 1
            End If
            lnNumLinCmac = 0
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "NroDocCmac:" & psComCmac & Space(sAux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & "NroDocCmac:" & psComCmac & lsNegritaOff & Chr$(10)   'Desactiva Negrita
            lnNumLinCmac = 1
        End If
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(sAux + sSep - 6) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Chr$(10)
        
        lnCliAux = InStr(1, psCliente, "*", vbTextCompare)
        
        If lnCliAux = 0 Then
            If sAux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + sAux)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(psCliente) & Space(sAux + sMax + sSep - Len(psCliente)) & ImpreCarEsp(psCliente) & Chr$(10)
        Else
            lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
            lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
            
            If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + sAux)
            If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + sAux)
            
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(lsCliAux1) & Space(sAux + sMax + sSep - Len(lsCliAux1)) & ImpreCarEsp(lsCliAux1) & Chr$(10)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(lsCliAux2) & Space(sAux + sMax + sSep - Len(lsCliAux2)) & ImpreCarEsp(lsCliAux2) & Chr$(10)
            
            lnCliAux = 1
        End If
        
        If pbSaldoInt Or pbCuenta Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Cuenta:" & psCodCta & Space(8 + sSep + sAux) & "Cuenta:" & psCodCta & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        psTit = Trim(psTit)
        psTit = CentrarCadena(psTit, 28)
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOn  'Activa Negrita & chr$(10)
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni + 1) & "-----" & psTit & "-----" & Space(-1 + sSep) & "-----" & psTit & "-----" & Chr$(10)
        
        lnChq = InStr(1, psText, "*", vbTextCompare)
        
        If psTexto = "XXX" Then
            If lnChq = 0 Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(Trim(Mid(psText, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(psText, 1, 28))) - Len(sMonto)) & sMonto & Space(-1 + sSep) & ImpreCarEsp(Trim(Mid(psText, 1, 28))) & Space(sMax + 6 - Len(Trim(Mid(psText, 1, 28))) - Len(sMonto)) & sMonto & Chr$(10)
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
            Else
                lsChqAux1 = (Mid(psText, 1, lnChq - 1))
                lsChqAux2 = (Mid(psText, lnChq + 1))
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Space(-1 + sSep) & ImpreCarEsp(Trim(lsChqAux1)) & Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)) & sMonto & Chr$(10)
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Space(-1 + sSep) & ImpreCarEsp(Trim(lsChqAux2)) & Space(sMax + 6 - Len(Trim(lsChqAux2))) & Chr$(10)
            End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(Trim(psTexto)) & Space(sAux + sSep + sMax - Len(Trim(psTexto))) & ImpreCarEsp(Trim(psTexto)) & Chr$(10)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOff  'Desactiva Negrita
        
        If pbSaldoInt Then
            'If MsgBox("Desea Imprimir el Saldos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Space(-1 + sSep) & "Saldo Disponible" & Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)) & sSDisp & Chr$(10)
                If pbOpSaldoC Then
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Space(-1 + sSep) & "Saldo Contable" & Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)) & lsSaldoC & Chr$(10)
                Else
                    ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
                End If
            'Else
            '    Print #nFicSal, ""
            '    Print #nFicSal, ""
            '    pbSaldoInt = False
            'End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        lsInteres = pnInteresA
        
        If pbSaldoInt Then
            If lsInteres <> "No Valido" Then
                lsInteres = Format(lsInteres, "#,##0.00")
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & Space(-1 + sSep) & "Interes del Mes" & Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)) & lsInteres & Chr$(10)
            End If
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        If Not psNumDoc = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & Space(-1 + sSep) & NomDoc & Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)) & psNumDoc & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        End If
        
        If Not psNumDias = "----" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Nro Dias Interes" & Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)) & psNumDias & Space(-1 + sSep) & "Nro Dias Interes" & Space(sMax + 6 - Len(psNumDias) - Len("Nro Dias Interes")) & psNumDias & Chr$(10)
            lnTope = 4 - lnCliAux
        Else
            lnTope = 3 - lnCliAux
        End If
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "---------------------------------------" & Space(-1 + sSep) & "---------------------------------------" & Chr$(10)
        If psCodUsuRem = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(sCodUser) & Space(29 + sSep + sAux) & ImpreCarEsp(sCodUser) & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Loc/Rem" & Space(sMax + sAux - Len("Loc/Rem") - 1 - 8) & ImpreCarEsp(sCodUser) & "/" & ImpreCarEsp(psCodUsuRem) & Space(sSep) & "Loc/Rem" & Space(sMax + sAux - Len("Loc/Rem") - 1 - 8) & ImpreCarEsp(sCodUser) & "/" & ImpreCarEsp(psCodUsuRem) & Chr$(10)
        End If
'        Dim clsGen As DGeneral
'        Set clsGen = New DGeneral
        lsMensaje = clsGen.GetMensajeBoletas(psCodCta)
        Set clsGen = Nothing
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & lsNegritaOn & ImpreCarEsp(lsMensaje) & Space(-1 + sSep) & ImpreCarEsp(lsMensaje) & lsNegritaOff
        
        lnNumLinCmac = lnNumLinCmac + 1
        
        For sAux = 1 To (lnTope - lnNumLinCmac)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        Next sAux
        
End If
'''- Fin de Liquidaciones
If sLpt <> "" Then
    Print #nFicSal, ImprimeBoletaAhorro
    Print #nFicSal, Chr$(12)
    Close nFicSal
End If

Exit Function

ERROR:
    Close nFicSal
    'If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
    '    GoTo ETIQ
    'End If
End Function


'**************************************************************
'Arma Recibo de Apertura de Cuentas Ahorros
'***************************************************************
Function EmiteBoletaAhorro(ByVal sCuenta As String, ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double, _
    ByVal pdFecSis As Date, ByVal psNomAge As String, ByVal psCodUser As String, Optional ByVal psCodCMAC As String) As String
Dim bReImp As Boolean
Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String, sNroDoc As String

sModDep = "Depósito Efectivo"
sTipApe = "APERTURA AHORROS"

Dim clsMant As NCapMantenimiento
Dim clsCap As NCapMovimientos
Set clsMant = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsMant.GetNombreTitulares(sCuenta))
Set clsMant = Nothing

EmiteBoletaAhorro = ImprimeBoletaAhorro(sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nSaldoCnt), sNomTit, sCuenta, "", nSaldoDisp, 0, "", 1, nSaldoCnt, , , , , , , , , , pdFecSis, Trim(psNomAge), psCodUser, sLpt, psCodCMAC)

End Function

Function EmiteBoletaAhorroAbonoRetiro(ByVal sCuenta As String, ByVal pnMonto As Double, ByVal nSaldoDisp As Double, _
    ByVal nSaldoCnt As Double, ByVal nIntGanado As Double, ByVal pnTipoOpeAho As Integer, _
    ByVal pdFecSis As Date, ByVal sNomAge As String, ByVal sCodUser As String, Optional ByVal psCodCMAC As String) As String
Dim bReImp As Boolean
Dim sTipDep As String, sCodOpe As String
Dim sModDep As String, sTipApe As String
Dim sNomTit As String, sNroDoc As String
Dim sMsgOpe  As String
Dim sCadOpe As String

If pnTipoOpeAho = 1 Then
    sModDep = "Deposito Efectivo"
    sTipApe = "DEPOSITO AHORROS"
    sMsgOpe = "Deposito OP"
    sCadOpe = gAhoDepEfec
Else
    sModDep = "Retiro Efectivo"
    sTipApe = "RETIRO AHORROS"
    sMsgOpe = "Retiro OP"
    sCadOpe = gAhoRetEfec
End If

Dim clsMant As NCapMantenimiento
Dim clsCap As NCapMovimientos
Set clsMant = New NCapMantenimiento
sNomTit = ImpreCarEsp(clsMant.GetNombreTitulares(sCuenta))
Set clsMant = Nothing

EmiteBoletaAhorroAbonoRetiro = ImprimeBoletaAhorro(sTipApe, ImpreCarEsp(sMsgOpe), sCadOpe, _
    Trim(pnMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, True, _
    , , , , , , , , pdFecSis, sNomAge, sCodUser, "", psCodCMAC)

End Function



'*******************************************************************************************************
'Devuelve Un a Arreglo de Strings donde devuelve todos los documentos que se emiten en un desembolso
'Tanto en Abono a cuenta como en efectivo, el valor de 1 Indica papel boleta y el valor 2 indica papel carta
'*******************************************************************************************************
Public Function ImprimeDocumentosDesembolso(ByVal psCtaCod As String, ByVal pnMovNro As Long, _
    ByVal psNomAge As String, ByVal psCodUsu As String, ByVal pdFecSis As Date, _
    Optional ByVal pbApeAboCta As Integer, Optional ByVal psctaAho As String = "", _
    Optional pnAhoMontoApe As Double = 0#, _
    Optional pnIntGanadoAbono As Double = 0#, Optional pnIntGanadoRetGastos As Double = 0#, _
    Optional pnIntGanadoRetCancel As Double = 0#, _
    Optional pnMontoAbono As Double = 0#, Optional pnMontoRetGastos As Double = 0#, _
    Optional pnMontoRetCancel As Double = 0#, _
    Optional pnSaldoDispAhoAbono As Double = 0#, Optional ByVal pnSaldoCntAhoAbono As Double = 0#, _
    Optional pnSaldoDispAhoRetGastos As Double = 0#, Optional pnSaldoCntAhoRetGastos As Double = 0#, _
    Optional pnSaldoDispAhoRetCancel As Double = 0#, Optional pnSaldoCntAhoRetCancel As Double = 0#, _
    Optional ByVal psNomCmac As String = "", Optional ByVal psNomCmacAbrev As String = "", Optional ByVal psCodCMAC As String) As Variant
    
Dim MatDocs() As String
Dim MatTempDesem(1, 2) As String
Dim MatTempDesemGastos(1, 2) As String
Dim MatTempDesemCancel() As String
Dim MatTempComprob(1, 2) As String
Dim MatTempPlan(2, 2) As String
Dim MatTempAhoApe(1, 2) As String
Dim MatTempAhoRetGastos(1, 2) As String
Dim MatTempAhoRetCancel(1, 2) As String

Dim nTotReg As Integer
Dim oDCredDoc As DCredDoc
Dim oDCredito As DCredito
Dim R As ADODB.Recordset
Dim sCadTmp As String
Dim AcumTmp As Double
Dim I As Integer
Dim nCapital As Double
Dim NGasto As Double
Dim nNumTran As Integer
Dim nLin As Integer
Dim nMontoDesemb As Double
Dim nCalMiVivi As Boolean
Dim dProxFecPag As Date

    nTotReg = 0
    '*****************************************************************************
    '****   Recuperar Boletas de Desembolso
    '*****************************************************************************
    
    Set oDCredito = New DCredito
    Set R = oDCredito.RecuperaDatosComunes(psCtaCod)
    nCalMiVivi = R!bMiVivienda
    nMontoDesemb = R!nMontoCol
    R.Close
    Set oDCredito = Nothing
    
    Set oDCredDoc = New DCredDoc
    Set R = oDCredDoc.RecuperaDatosBoletaDesembolso(pnMovNro, psCtaCod, psNomAge)
    sCadTmp = ""
    vsCadImpresion = ""
    nNumTran = 0
    Do While Not R.EOF
        If R.Bookmark = 1 Then
            nCapital = IIf(IsNull(R!nCapital), 0, R!nCapital)
        Else
            NGasto = IIf(IsNull(R!NGasto), 0, R!NGasto)
        End If
        nNumTran = nNumTran + R!nNumTran
        R.MoveNext
    Loop
    R.MoveFirst
    Call ImprimeCabeceraBoletaPago(psNomAge, R!sMoneda, "", _
            R!FechaTran, R!HoraTran, R!cPersNombre, Trim(Str(nNumTran)), psCtaCod, "", psNomCmacAbrev, psCodCMAC)
    'Cuerpo de la Boleta
    
    sCadTmp = vsCadImpresion
    sCadTmp = sCadTmp & lsNegritaOn
    
If psCodCMAC = "102" Then  ' LIMA FEBM
    If pnMontoAbono > 0 Or pnAhoMontoApe > 0 Then
           sCadTmp = sCadTmp & String(10, "-") & "DESEMB CON ABONO A CUENTA" & String(10, "-") & lsSaltoLin
    Else
            sCadTmp = sCadTmp & String(10, "-") & "DESEMBOLSO  EN  EFECTIVO " & String(10, "-") & lsSaltoLin
    End If
    
    sCadTmp = sCadTmp & "Monto del Desembolso :" & Space(2) & ImpreFormat(nCapital, 12) & lsSaltoLin
    sCadTmp = sCadTmp & "Gastos:" & ImpreFormat(Abs(NGasto), 29) & lsSaltoLin
    sCadTmp = sCadTmp & lsSaltoLin
    sCadTmp = sCadTmp & lsSaltoLin
    sCadTmp = sCadTmp & String(40, "-") & lsSaltoLin
    sCadTmp = sCadTmp & "Total a Pagar :" & Space(8) & ImpreFormat(nCapital + NGasto, 12) & lsSaltoLin
    sCadTmp = sCadTmp & lsNegritaOff
    sCadTmp = sCadTmp & "Saldo Capital :" & Space(7) & ImpreFormat(nCapital, 12) & lsSaltoLin
    sCadTmp = sCadTmp & ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin & lsSaltoLin & lsSaltoLin
    sCadTmp = sCadTmp & lsSaltoLin
    dProxFecPag = R!dVenc
    R.Close
    
Else   ' Liquidación Trujillo
    If pnMontoAbono > 0 Or pnAhoMontoApe > 0 Then
           sCadTmp = sCadTmp & String(10, "-") & "DESEMB CON ABONO A CUENTA" & String(10, "-")
           sCadTmp = sCadTmp & Space(7) & String(9, "-") & "DESEMB CON ABONO A CUENTA" & String(10, "-") & lsSaltoLin
    Else
        sCadTmp = sCadTmp & String(10, "-") & "DESEMBOLSO  EN  EFECTIVO " & String(10, "-")
        sCadTmp = sCadTmp & Space(7) & String(9, "-") & "DESEMBOLSO  EN  EFECTIVO " & String(10, "-") & lsSaltoLin
    End If
    
    sCadTmp = sCadTmp & "Monto del Desembolso :" & Space(2) & ImpreFormat(nCapital, 12)
    sCadTmp = sCadTmp & Space(14) & "Monto del Desembolso :" & Space(2) & ImpreFormat(nCapital, 12) & lsSaltoLin
    sCadTmp = sCadTmp & "Gastos:" & ImpreFormat(Abs(NGasto), 29)
    sCadTmp = sCadTmp & Space(14) & "Gastos :" & ImpreFormat(Abs(NGasto), 28) & lsSaltoLin
    sCadTmp = sCadTmp & lsSaltoLin
    sCadTmp = sCadTmp & lsSaltoLin
    sCadTmp = sCadTmp & String(40, "-")
    sCadTmp = sCadTmp & Space(14) & String(40, "-") & lsSaltoLin
    sCadTmp = sCadTmp & "Total a Pagar :" & Space(8) & ImpreFormat(nCapital + NGasto, 12)
    sCadTmp = sCadTmp & Space(16) & "Total a Pagar :" & Space(8) & ImpreFormat(nCapital + NGasto, 12) & lsSaltoLin
    sCadTmp = sCadTmp & lsNegritaOff
    sCadTmp = sCadTmp & "Saldo Capital :" & Space(7) & ImpreFormat(nCapital, 12)
    sCadTmp = sCadTmp & Space(16) & "Saldo Capital :" & Space(7) & ImpreFormat(nCapital, 13) & lsSaltoLin
    sCadTmp = sCadTmp & ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin & lsSaltoLin & lsSaltoLin
    sCadTmp = sCadTmp & lsSaltoLin
    dProxFecPag = R!dVenc
    R.Close

End If
 '' Fin
    
    MatTempDesem(0, 0) = "1"
    MatTempDesem(0, 1) = sCadTmp
    nTotReg = nTotReg + 1
    
    '*****************************************************************************
    '********** Recuperar Boletas de Gastos de Desembolso
    '*****************************************************************************
    sCadTmp = ""
    vsCadImpresion = ""
    Set R = oDCredDoc.RecuperaDatosBoletaGastosDesembolso(pnMovNro, psCtaCod, psNomAge)
    If Not R.BOF And Not R.EOF Then
        Call ImprimeCabeceraBoletaPago(psNomAge, R!sMoneda, "", _
                R!FechaTran, R!HoraTran, R!cPersNombre, "", psCtaCod, "", psNomCmacAbrev, psCodCMAC)
        'Cuerpo de la Boleta
        
        sCadTmp = lsSaltoLin & lsSaltoLin & vsCadImpresion
        sCadTmp = sCadTmp & lsNegritaOn
        If psCodCMAC = "102" Then
            sCadTmp = sCadTmp & String(5, "-") & "COMPROBANTE DE GASTOS DE DESEMBOLSO" & String(5, "-") & lsSaltoLin
        Else
            sCadTmp = sCadTmp & String(5, "-") & "COMPROBANTE DE GASTOS DE DESEMBOLSO" & String(5, "-")
            sCadTmp = sCadTmp & Space(7) & String(5, "-") & "COMPROBANTE DE GASTOS DE DESEMBOLSO" & String(4, "-") & lsSaltoLin
        End If
    End If
    AcumTmp = 0
    Do While Not R.EOF
     If psCodCMAC = 102 Then
        sCadTmp = sCadTmp & ImpreFormat(Mid(R!cDescripcion, 1, 25), 25) & ImpreFormat(Abs(R!nMonto), 12) & lsSaltoLin
     Else
        sCadTmp = sCadTmp & ImpreFormat(Mid(R!cDescripcion, 1, 25), 25) & ImpreFormat(Abs(R!nMonto), 12)
        sCadTmp = sCadTmp & Space(11) & ImpreFormat(Mid(R!cDescripcion, 1, 25), 25) & ImpreFormat(Abs(R!nMonto), 12) & lsSaltoLin
     End If
        AcumTmp = AcumTmp + Abs(R!nMonto)
        R.MoveNext
    Loop
    
    If R.RecordCount > 0 Then
        nLin = 7 - R.RecordCount
        If nLin < 0 Then
            nLin = 0
        End If
        For I = 1 To nLin
            sCadTmp = sCadTmp & lsSaltoLin
        Next I
        If psCodCMAC = "102" Then
            sCadTmp = sCadTmp & String(40, "-") & lsSaltoLin
            sCadTmp = sCadTmp & ImpreFormat("Total Gastos :", 25) & ImpreFormat(AcumTmp, 12) & lsSaltoLin
            sCadTmp = sCadTmp & lsNegritaOff
            R.MovePrevious
            sCadTmp = sCadTmp & ImpreFormat("Prox. Fecha Pago : " & Format(dProxFecPag, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin
            sCadTmp = sCadTmp & lsSaltoLin
       Else
            sCadTmp = sCadTmp & String(40, "-")
            sCadTmp = sCadTmp & Space(14) & String(40, "-") & lsSaltoLin
            sCadTmp = sCadTmp & ImpreFormat("Total Gastos :", 25) & ImpreFormat(AcumTmp, 12)
            sCadTmp = sCadTmp & Space(11) & ImpreFormat("Total Gastos :", 25) & ImpreFormat(AcumTmp, 12) & lsSaltoLin
            sCadTmp = sCadTmp & lsNegritaOff
            R.MovePrevious
            sCadTmp = sCadTmp & ImpreFormat("Prox. Fecha Pago : " & Format(dProxFecPag, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & ImpreFormat("Prox. Fecha Pago : " & Format(dProxFecPag, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin
            sCadTmp = sCadTmp & lsSaltoLin
        End If
    End If
    R.Close
        
    MatTempDesemGastos(0, 0) = "1"
    MatTempDesemGastos(0, 1) = sCadTmp
    nTotReg = nTotReg + 1
    '*****************************************************************************
    '********** Recuperar Boletas de Cancelacion de Creditos
    '*****************************************************************************
    sCadTmp = ""
    vsCadImpresion = ""
    Set R = oDCredDoc.RecuperaDatosBoletaCancelCredDesembolso(pnMovNro, psCtaCod, psNomAge)
    ReDim MatTempDesemCancel(R.RecordCount, 2)
    Do While Not R.EOF
        sCadTmp = ""
        vsCadImpresion = ""
    If psCodCMAC = "102" Then '' Liquidacion lima
        Call ImprimeCabeceraBoletaPago(psNomAge, R!sMoneda, "", _
                R!FechaTran, R!HoraTran, R!cPersNombre, Trim(Str(R!nNumTran)), R!cCtaCod, "", psNomCmacAbrev, psCodCMAC)
        'Cuerpo de la Boleta
        sCadTmp = vsCadImpresion
        sCadTmp = sCadTmp & lsNegritaOn
        sCadTmp = sCadTmp & String(10, "-") & "CANCELACION CREDITO" & String(10, "-") & lsSaltoLin
        
        sCadTmp = sCadTmp & "Capital :" & Space(13) & ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital), 12) & lsSaltoLin
        sCadTmp = sCadTmp & "Interes :" & Space(13) & ImpreFormat(IIf(IsNull(R!nInteres), 0, R!nInteres), 12) & lsSaltoLin
        sCadTmp = sCadTmp & "Gastos  :" & Space(13) & ImpreFormat(IIf(IsNull(R!nGastos), 0, R!nGastos), 12) & lsSaltoLin
        sCadTmp = sCadTmp & lsSaltoLin
        sCadTmp = sCadTmp & lsSaltoLin
        sCadTmp = sCadTmp & Space(14) & String(40, "-") & lsSaltoLin
        sCadTmp = sCadTmp & "Total a Pagar :" & Space(8) & ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital) + IIf(IsNull(R!nGastos), 0, R!nGastos) + IIf(IsNull(R!nInteres), 0, R!nInteres), 12) & lsSaltoLin
        sCadTmp = sCadTmp & lsNegritaOff
        sCadTmp = sCadTmp & "Saldo Capital :" & Space(7) & ImpreFormat(0, 12) & lsSaltoLin
        'sCadTmp = sCadTmp & ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin
        sCadTmp = sCadTmp & ImpreFormat("Prox. Fecha Pago : " & Space(10) & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin
        MatTempDesemCancel(R.Bookmark - 1, 0) = "1"
        MatTempDesemCancel(R.Bookmark - 1, 1) = sCadTmp
        nTotReg = nTotReg + 1
        R.MoveNext
    Else
        Call ImprimeCabeceraBoletaPago(psNomAge, R!sMoneda, "", _
                R!FechaTran, R!HoraTran, R!cPersNombre, Trim(Str(R!nNumTran)), R!cCtaCod, "", psNomCmacAbrev, psCodCMAC)
        'Cuerpo de la Boleta
        sCadTmp = vsCadImpresion
        sCadTmp = sCadTmp & lsNegritaOn
        sCadTmp = sCadTmp & String(10, "-") & "CANCELACION CREDITO" & String(10, "-")
        sCadTmp = sCadTmp & Space(14) & String(10, "-") & "CANCELACION CREDITO" & String(10, "-") & lsSaltoLin
        
        sCadTmp = sCadTmp & "Capital :" & Space(13) & ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital), 12)
        sCadTmp = sCadTmp & Space(16) & "Capital :" & Space(13) & ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital), 12) & lsSaltoLin
        sCadTmp = sCadTmp & "Interes :" & Space(13) & ImpreFormat(IIf(IsNull(R!nInteres), 0, R!nInteres), 12)
        sCadTmp = sCadTmp & Space(16) & "Interes :" & Space(13) & ImpreFormat(IIf(IsNull(R!nInteres), 0, R!nInteres), 12) & lsSaltoLin
        sCadTmp = sCadTmp & "Gastos  :" & Space(13) & ImpreFormat(IIf(IsNull(R!nGastos), 0, R!nGastos), 12)
        sCadTmp = sCadTmp & Space(16) & "Gastos  :" & Space(13) & ImpreFormat(IIf(IsNull(R!nGastos), 0, R!nGastos), 12) & lsSaltoLin
        sCadTmp = sCadTmp & lsSaltoLin
        sCadTmp = sCadTmp & lsSaltoLin
        sCadTmp = sCadTmp & String(40, "-")
        sCadTmp = sCadTmp & Space(14) & String(40, "-") & lsSaltoLin
        sCadTmp = sCadTmp & "Total a Pagar :" & Space(8) & ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital) + IIf(IsNull(R!nGastos), 0, R!nGastos) + IIf(IsNull(R!nInteres), 0, R!nInteres), 12)
        sCadTmp = sCadTmp & Space(16) & "Total a Pagar :" & Space(8) & ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital) + IIf(IsNull(R!nInteres), 0, R!nInteres) + IIf(IsNull(R!nGastos), 0, R!nGastos), 12) & lsSaltoLin
        sCadTmp = sCadTmp & lsNegritaOff
        sCadTmp = sCadTmp & "Saldo Capital :" & Space(7) & ImpreFormat(0, 12)
        sCadTmp = sCadTmp & Space(16) & "Saldo Capital :" & Space(7) & ImpreFormat(0, 12) & lsSaltoLin
        'sCadTmp = sCadTmp & ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & ImpreFormat("Prox. Fecha Pago : " & Format(R!dVenc, "dd/mm/yyyy") & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin
        sCadTmp = sCadTmp & ImpreFormat("Prox. Fecha Pago : " & Space(10) & Space(3) & R!sUsuario, 53, 0) & ImpreFormat("Prox. Fecha Pago : " & Space(10) & Space(3) & R!sUsuario, 53, 0) & lsSaltoLin
        MatTempDesemCancel(R.Bookmark - 1, 0) = "1"
        MatTempDesemCancel(R.Bookmark - 1, 1) = sCadTmp
        nTotReg = nTotReg + 1
        R.MoveNext
    End If
        
    Loop
    R.Close
    
    '***********************************************************
    'Recuperar Boletas de Apertura de Cuenta
    '***********************************************************
    If pnAhoMontoApe > 0 Then
        MatTempAhoApe(0, 0) = "1"
        MatTempAhoApe(0, 1) = EmiteBoletaAhorro(psctaAho, pnAhoMontoApe, pnAhoMontoApe, Format(FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss"), psNomAge, psCodUsu, psCodCMAC)
        nTotReg = nTotReg + 1
    End If
    
    '***********************************************************
    'Recuperar Boletas de Abono de Cuenta
    '***********************************************************
    If pnMontoAbono > 0 Then
        MatTempAhoApe(0, 0) = "1"
        MatTempAhoApe(0, 1) = EmiteBoletaAhorroAbonoRetiro(psctaAho, pnMontoAbono, pnSaldoDispAhoAbono, _
            pnSaldoCntAhoAbono, pnIntGanadoAbono, 1, Format(FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss"), _
            psNomAge, psCodUsu, psCodCMAC)
        nTotReg = nTotReg + 1
    End If
    '***********************************************************
    'Recuperar Boletas de Retiro Gastos de Cuenta
    '***********************************************************
    If pnMontoRetGastos > 0 Then
        MatTempAhoRetGastos(0, 0) = "1"
        MatTempAhoRetGastos(0, 1) = EmiteBoletaAhorroAbonoRetiro(psctaAho, pnMontoRetGastos, pnSaldoDispAhoRetGastos, _
            pnSaldoCntAhoRetGastos, pnIntGanadoRetGastos, 2, Format(FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss"), _
            psNomAge, psCodUsu, psCodCMAC)
        nTotReg = nTotReg + 1
    End If
    
    '***********************************************************
    'Recuperar Boletas de Retiro Cancelaciones de Cuenta
    '***********************************************************
    If pnMontoRetCancel > 0 Then
        MatTempAhoRetCancel(0, 0) = "1"
        MatTempAhoRetCancel(0, 1) = EmiteBoletaAhorroAbonoRetiro(psctaAho, pnMontoRetCancel, pnSaldoDispAhoRetCancel, _
            pnSaldoCntAhoRetCancel, pnIntGanadoRetCancel, 2, Format(FechaHora(pdFecSis), "dd/mm/yyyy hh:mm:ss"), _
            psNomAge, psCodUsu, psCodCMAC)
        nTotReg = nTotReg + 1
    End If
    
    '********************************************************************************
    'Recuperar Documento de Comprobante de Desembolso
    '********************************************************************************
    MatTempComprob(0, 0) = "2"
    MatTempComprob(0, 1) = ImprimeComprobanteDesembolso(psCtaCod, psNomAge, Format(pdFecSis, "dd/mm/yyyy"), psCodUsu, psNomCmac)
    nTotReg = nTotReg + 1
    '********************************************************************************
    'Recuperar Documento de Plan de Pagos
    '********************************************************************************
    MatTempPlan(0, 0) = "2"
    MatTempPlan(0, 1) = ImprimePlandePagos(psCtaCod, psNomAge, Format(pdFecSis, "dd/mm/yyyy"), psCodUsu, nMontoDesemb, nCalMiVivi)
    If nCalMiVivi = True Then
        MatTempPlan(0, 1) = MatTempPlan(0, 1) & ImprimePlandePagos(psCtaCod, psNomAge, Format(pdFecSis, "dd/mm/yyyy"), psCodUsu, nMontoDesemb, nCalMiVivi, True)
    End If
    nTotReg = nTotReg + 1
    
    
    
    ReDim MatDocs(nTotReg, 2)
    MatDocs(0, 0) = MatTempDesem(0, 0)
    MatDocs(0, 1) = MatTempDesem(0, 1)
    MatDocs(1, 0) = MatTempDesemGastos(0, 0)
    MatDocs(1, 1) = MatTempDesemGastos(0, 1)
    nTotReg = 2
    For I = 0 To UBound(MatTempDesemCancel) - 1
        MatDocs(I + 2, 0) = MatTempDesemCancel(I, 0)
        MatDocs(I + 2, 1) = MatTempDesemCancel(I, 1)
        nTotReg = nTotReg + 1
    Next I
    MatDocs(nTotReg, 0) = MatTempComprob(0, 0)
    MatDocs(nTotReg, 1) = MatTempComprob(0, 1)
    nTotReg = nTotReg + 1
    'Apertura
    If pnAhoMontoApe > 0 Then
        MatDocs(nTotReg, 0) = MatTempAhoApe(0, 0)
        MatDocs(nTotReg, 1) = MatTempAhoApe(0, 1)
        nTotReg = nTotReg + 1
    End If
    'Abono
    If pnMontoAbono > 0 Then
        MatDocs(nTotReg, 0) = MatTempAhoApe(0, 0)
        MatDocs(nTotReg, 1) = MatTempAhoApe(0, 1)
        nTotReg = nTotReg + 1
    End If
    'Retiro de Gastos
    If pnMontoRetGastos > 0 Then
        MatDocs(nTotReg, 0) = MatTempAhoRetGastos(0, 0)
        MatDocs(nTotReg, 1) = MatTempAhoRetGastos(0, 1)
        nTotReg = nTotReg + 1
    End If
    'Retiro de Cancelaciones
    If pnMontoRetCancel > 0 Then
        MatDocs(nTotReg, 0) = MatTempAhoRetCancel(0, 0)
        MatDocs(nTotReg, 1) = MatTempAhoRetCancel(0, 1)
        nTotReg = nTotReg + 1
    End If
    MatDocs(nTotReg, 0) = MatTempPlan(0, 0)
    MatDocs(nTotReg, 1) = MatTempPlan(0, 1)
    
    ImprimeDocumentosDesembolso = MatDocs
    
End Function


Public Function ImprimeBoletaCmacLlam(ByVal psNomAge As String, ByVal pdFecSis As Date, _
    ByVal psNomCli As String, ByVal psNomCmac As String, ByVal pnMonto As Double, _
    ByVal psCodUser As String, Optional ByVal psCodCMAC As String) As String
If psCodCMAC = "102" Then
    ImprimeBoletaCmacLlam = ""
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & ImpreFormat("CMACT", 20) & ImpreFormat(Format(pdFecSis, "dd/mm/yyyy hh:mm:ss"), 20) & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & psNomCli & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & psNomCmac & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & String(40, "*") & Chr$(10) & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & "Monto :" & ImpreFormat(pnMonto, 40) & Chr$(10) & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & String(40, "*") & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & ImpreFormat(psCodUser, 40) & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Chr$(10)
Else '' trujillo
    ImprimeBoletaCmacLlam = ""
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & ImpreFormat("CMACT", 20) & ImpreFormat(Format(pdFecSis, "dd/mm/yyyy hh:mm:ss"), 20) & Space(20)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & ImpreFormat("CMACT", 20) & ImpreFormat(Format(pdFecSis, "dd/mm/yyyy hh:mm:ss"), 20) & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & psNomCli & Space(20)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & psNomCli & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & psNomCmac & Space(20)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & psNomCmac & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & String(40, "*") & Space(20)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & String(40, "*") & Chr$(10) & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & "Monto :" & ImpreFormat(pnMonto, 40) & Space(20)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & "Monto :" & ImpreFormat(pnMonto, 40) & Chr$(10) & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & String(40, "*") & Space(20)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & String(40, "*") & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & ImpreFormat(psCodUser, 40) & Space(20)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Space(5) & ImpreFormat(psCodUser, 40) & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Chr$(10)
    ImprimeBoletaCmacLlam = ImprimeBoletaCmacLlam & Chr$(10)
End If
End Function

Public Function ImprimePagodeCreditos(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "") As String
    
Dim oDCred As DCredDoc
Dim sCadImp As String
Dim R As ADODB.Recordset
Dim vsProducto As String
Dim vsFinanc As String
Dim vsOperacion As String

    Set oDCred = New DCredDoc
    Set R = oDCred.ReportePagoDeCreditos(pMatAgencias, pdFecIni, pdFecFin, pnMoneda)
    Set oDCred = Nothing
    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "INGRESOS POR PAGO DE CREDITOS DEL DIA " & Format(pdFecFin, "dd/mm/yyyy"), psNomCmac, , True)
    
    If R.RecordCount > 0 Then
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!cOpeCod) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("IMPORTE", 13) & ImpreFormat("CAPITAL", 13) & ImpreFormat("INTERES", 16) & ImpreFormat("MORA", 12) & ImpreFormat("GASTOS", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
        'FINAL DE OPERACION
        If Trim(R!cCtaCod) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL OPERACION ", 42) & ImpreFormat(R!nCapital + R!nInteres + R!nInteresMora + R!nInteresGracia + R!nInteresReprog + R!nInteresSuspenso + R!nGastos, 20) & _
                                        ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres + R!nInteresGracia + R!nInteresReprog + R!nInteresSuspenso, 12) & ImpreFormat(R!nInteresMora, 12) & ImpreFormat(R!nGastos, 12) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE FUENTE DE FINANCIAMIENTO
            If Trim(R!cOpeCod) = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL FUENTE FINANCIAMIENTO ", 42) & ImpreFormat(R!nCapital + R!nInteres + R!nInteresMora + R!nInteresGracia + R!nInteresReprog + R!nInteresSuspenso + R!nGastos, 20) & _
                                        ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres + R!nInteresGracia + R!nInteresReprog + R!nInteresSuspenso, 12) & ImpreFormat(R!nInteresMora, 12) & ImpreFormat(R!nGastos, 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL DE PRODUCTO
                If R!sFuentefinan = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PRODUCTO ", 42) & ImpreFormat(R!nCapital + R!nInteres + R!nInteresMora + R!nInteresGracia + R!nInteresReprog + R!nInteresSuspenso + R!nGastos, 20) & _
                                        ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres + R!nInteresGracia + R!nInteresReprog + R!nInteresSuspenso, 12) & ImpreFormat(R!nInteresMora, 12) & ImpreFormat(R!nGastos, 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    R.MoveNext
                    If R.EOF Then
                        Exit Do
                    End If
                    If R!sProducto = "TOTAL" Then
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 42) & ImpreFormat(R!nCapital + R!nInteres + R!nInteresMora + R!nInteresGracia + R!nInteresReprog + R!nInteresSuspenso + R!nGastos, 20) & _
                                        ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres + R!nInteresGracia + R!nInteresReprog + R!nInteresSuspenso, 12) & ImpreFormat(R!nInteresMora, 12) & ImpreFormat(R!nGastos, 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        Exit Do
                    Else
                        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!cOpeCod) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("IMPORTE", 13) & ImpreFormat("CAPITAL", 13) & ImpreFormat("INTERES", 16) & ImpreFormat("MORA", 12) & ImpreFormat("GASTOS", 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    End If
                Else
                    sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!cOpeCod) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("IMPORTE", 13) & ImpreFormat("CAPITAL", 13) & ImpreFormat("INTERES", 16) & ImpreFormat("MORA", 12) & ImpreFormat("GASTOS", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!cOpeCod) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("IMPORTE", 13) & ImpreFormat("CAPITAL", 13) & ImpreFormat("INTERES", 16) & ImpreFormat("MORA", 12) & ImpreFormat("GASTOS", 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
            End If
        End If
        
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!sTitular, 20) & _
                                    ImpreFormat(R!nCapital + R!nInteres + R!nInteresMora + R!nInteresGracia + R!nInteresReprog + R!nInteresSuspenso + R!nGastos, 20) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres + R!nInteresGracia + R!nInteresReprog + R!nInteresSuspenso, 12) & ImpreFormat(R!nInteresMora, 12) & ImpreFormat(R!nGastos, 12) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimePagodeCreditos = sCadImp
End Function


Public Function ImprimeDesembolsosEfectuados(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatProducts As Variant, Optional ByVal psNomCmac As String = "")
Dim oDCred As DCredDoc
Dim sCadImp As String
Dim R As ADODB.Recordset

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosDesembolsados(pMatAgencias, pdFecIni, pdFecFin, pnMoneda, pMatProducts)
    Set oDCred = Nothing
    sCadImp = ""
        
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS DESEMBOLSADOS DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!cOpeCod) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 19) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("PRESTAMO", 13) & ImpreFormat("INT APROB", 16) & ImpreFormat("GASTO", 8) & ImpreFormat("ANALISTA", 10) & ImpreFormat("INTERES", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
        'FINAL DE OPERACION
        If Trim(R!NCredito) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL OPERACION ", 42) & _
                                    ImpreFormat(R!nDesembolso, 22) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nIntApr, 10) & ImpreFormat(R!Gasto, 14) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE FUENTE DE FINANCIAMIENTO
            If Trim(R!cOpeCod) = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL FUENTE FINANCIAMIENTO ", 42) & _
                                    ImpreFormat(R!nDesembolso, 22) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nIntApr, 10) & ImpreFormat(R!Gasto, 14) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL DE PRODUCTO
                If R!sFuentefinan = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PRODUCTO ", 42) & _
                                    ImpreFormat(R!nDesembolso, 22) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nIntApr, 10) & ImpreFormat(R!Gasto, 14) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    R.MoveNext
                    If R.EOF Then
                        Exit Do
                    End If
                    If R!sProducto = "TOTAL" Then
                        sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL ", 42) & _
                                    ImpreFormat(R!nDesembolso, 22) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nIntApr, 10) & ImpreFormat(R!Gasto, 14) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        R.MoveNext
                        If R.EOF Then
                            Exit Do
                        End If
                        'TOTAL AGENCIA
                        If R!sAgencia = "TOTAL" Then
                            sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 42) & _
                                        ImpreFormat(R!nDesembolso, 22) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nIntApr, 10) & ImpreFormat(R!Gasto, 14) & ImpreFormat(R!nTasa, 18, 4) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            Exit Do
                        Else
                            nPuntPag = 0
                            sCadImp = sCadImp & Chr$(12)
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!cOpeCod) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 19) & ImpreFormat("CLIENTE", 34) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("INT APROB", 16) & ImpreFormat("GASTO", 8) & ImpreFormat("ANALISTA", 10) & ImpreFormat("INTERES", 12) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
                        End If
                    Else
                        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!cOpeCod) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 19) & ImpreFormat("CLIENTE", 34) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("INT APROB", 16) & ImpreFormat("GASTO", 8) & ImpreFormat("ANALISTA", 10) & ImpreFormat("INTERES", 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
                    End If
                Else
                    sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!cOpeCod) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 19) & ImpreFormat("CLIENTE", 34) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("INT APROB", 16) & ImpreFormat("GASTO", 8) & ImpreFormat("ANALISTA", 10) & ImpreFormat("INTERES", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!cOpeCod) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 19) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("APROBADO", 13) & ImpreFormat("INT APROB", 16) & ImpreFormat("GASTO", 8) & ImpreFormat("ANALISTA", 10) & ImpreFormat("INTERES", 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
            End If
        End If
        
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!NCredito, 20) & ImpreFormat(R!sTitular, 20) & _
            ImpreFormat(R!nDesembolso, 22) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nIntApr, 10) & ImpreFormat(R!Gasto, 14) & ImpreFormat(R!sAnalista, 6, 6) & ImpreFormat(R!nTasa, 6, 4) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
   ImprimeDesembolsosEfectuados = sCadImp
End Function


Public Function ImprimeSaldoCarteraVigente(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "")
Dim oDCred As DCredDoc
Dim sCadImp As String
Dim R As ADODB.Recordset

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaSaldoCarteraVigente(pMatAgencias, pdFecIni, pnMoneda, psCodUser, pdFecSis, psNomAge)
    Set oDCred = Nothing
    sCadImp = ""
        
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "SALDOS DE CARTERA VIGENTE AL " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("FINANCIAMIENTO", 30) & ImpreFormat("CASOS", 7) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("SALDO CAP", 16) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
        'FINAL DE PLAZO
        If Trim(R!sFuentefinan) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PLAZO ", 31) & ImpreFormat(R!nCasos, 4, 0) & _
                                    ImpreFormat(R!nMontoApr, 12) & ImpreFormat(R!nDesembolsado, 12) & ImpreFormat(R!nSaldoCapital, 12) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE PRODUCTO
            If Trim(R!sPlazo) = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PRODUCTO ", 31) & ImpreFormat(R!nCasos, 4, 0) & _
                                    ImpreFormat(R!nMontoApr, 12) & ImpreFormat(R!nDesembolsado, 12) & ImpreFormat(R!nSaldoCapital, 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL DE AGENCIA
                If R!sProducto = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA", 31) & ImpreFormat(R!nCasos, 4, 0) & _
                                    ImpreFormat(R!nMontoApr, 12) & ImpreFormat(R!nDesembolsado, 12) & ImpreFormat(R!nSaldoCapital, 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    R.MoveNext
                    If R.EOF Then
                        Exit Do
                    End If
                    'FINAL
                    If R!sAgencia = "TOTAL" Then
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 31) & ImpreFormat(R!nCasos, 4, 0) & _
                                    ImpreFormat(R!nMontoApr, 12) & ImpreFormat(R!nDesembolsado, 12) & ImpreFormat(R!nSaldoCapital, 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        Exit Do
                    Else
                        nPuntPag = 0
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("FINANCIAMIENTO", 30) & ImpreFormat("CASOS", 7) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("SALDO CAP", 16) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    End If
                Else
                    sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("FINANCIAMIENTO", 30) & ImpreFormat("CASOS", 7) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("SALDO CAP", 16) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("FINANCIAMIENTO", 30) & ImpreFormat("CASOS", 7) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("SALDO CAP", 16) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
            End If
        End If
        
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!sFuentefinan, 20) & ImpreFormat(R!nCasos, 15, 0) & _
                                    ImpreFormat(R!nMontoApr, 12) & ImpreFormat(R!nDesembolsado, 12) & ImpreFormat(R!nSaldoCapital, 12) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeSaldoCarteraVigente = sCadImp
End Function


Public Function ImprimeCreditosCancelados(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "")
    
Dim oDCred As DCredDoc
Dim sCadImp As String
Dim R As ADODB.Recordset

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosCanceladosTodos(pMatAgencias, pdFecIni, pdFecFin, pnMoneda)
    Set oDCred = Nothing
    sCadImp = ""
        
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS CANCELADOS DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DESEMBOLSO", 13) & ImpreFormat("Nro CUOTAS", 5) & ImpreFormat("PLAZO", 5) & ImpreFormat("VIGENCIA", 15) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
        'FINAL DE Financiamiento
        If Trim(R!NCredito) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL FINANCIAMIENTO ", 42) & ImpreFormat(R!nDesembolso, 20) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE FUENTE DE Producto
            If R!sFuentefinan = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PRODUCTO ", 42) & ImpreFormat(R!nDesembolso, 20) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'Final del Plazo
                If R!sProducto = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PLAZO", 42) & ImpreFormat(R!nDesembolso, 20) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    R.MoveNext
                    If R.EOF Then
                        Exit Do
                    End If
                    'FINAL DE AGENCIAS
                    If R!sPlazo = "TOTAL" Then
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 42) & ImpreFormat(R!nDesembolso, 20) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        R.MoveNext
                        If R.EOF Then
                            Exit Do
                        End If
                        'TOTAL FINAL
                        If R!sAgencia = "TOTAL" Then
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 42) & ImpreFormat(R!nDesembolso, 20) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            Exit Do
                        Else
                            nPuntPag = 0
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DESEMBOLSO", 13) & ImpreFormat("Nro CUOTAS", 5) & ImpreFormat("PLAZO", 5) & ImpreFormat("VIGENCIA", 15) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        End If
                    Else
                        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DESEMBOLSO", 13) & ImpreFormat("Nro CUOTAS", 5) & ImpreFormat("PLAZO", 5) & ImpreFormat("VIGENCIA", 15) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    End If
                Else
                    sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DESEMBOLSO", 13) & ImpreFormat("Nro CUOTAS", 5) & ImpreFormat("PLAZO", 5) & ImpreFormat("VIGENCIA", 15) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DESEMBOLSO", 13) & ImpreFormat("Nro CUOTAS", 5) & ImpreFormat("PLAZO", 5) & ImpreFormat("VIGENCIA", 15) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            End If
        End If
        
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!NCredito, 20) & ImpreFormat(R!sTitular, 20) & _
                                    ImpreFormat(R!nDesembolso, 20) & _
                                    ImpreFormat(R!nCuotasApr, 12, 0) & ImpreFormat(R!nPlazoApr, 7, 0) & ImpreFormat(Format(R!dVigencia, "dd/mm/yyyy"), 15) & ImpreFormat(R!sAnalista, 12) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeCreditosCancelados = sCadImp
End Function


Public Function ImprimeCreditosCanceladosConSaldo(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "")
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosCanceladosConsaldo(pMatAgencias, pdFecIni, pdFecFin, pnMoneda)
    Set oDCred = Nothing
    
    sCadImp = ""
        
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS CANCELADOS DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("SALDO", 13) & ImpreFormat("CAPITAL", 12) & ImpreFormat("INTERES", 12) & ImpreFormat("MORA", 8) & ImpreFormat("INT. PERD", 12) & ImpreFormat("LINEA C.", 8) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
        'FINAL DE Financiamiento
        If Trim(R!NCredito) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL FINANCIAMIENTO ", 42) & ImpreFormat(R!nSaldo, 20) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres, 12) & ImpreFormat(R!nMora, 7) & ImpreFormat(R!nIntPerd, 12) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE FUENTE DE Producto
            If R!sFuentefinan = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PRODUCTO ", 42) & ImpreFormat(R!nSaldo, 20) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres, 12) & ImpreFormat(R!nMora, 7) & ImpreFormat(R!nIntPerd, 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL DEL PLAZO
                If R!sProducto = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PLAZO", 42) & ImpreFormat(R!nSaldo, 20) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres, 12) & ImpreFormat(R!nMora, 7) & ImpreFormat(R!nIntPerd, 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    R.MoveNext
                    If R.EOF Then
                        Exit Do
                    End If
                    'Final de Agencias
                    If R!sPlazo = "TOTAL" Then
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 42) & ImpreFormat(R!nSaldo, 20) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres, 12) & ImpreFormat(R!nMora, 7) & ImpreFormat(R!nIntPerd, 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        R.MoveNext
                        If R.EOF Then
                            Exit Do
                        End If
                        'TOTAL FINAL
                        If R!sAgencia = "TOTAL" Then
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 42) & ImpreFormat(R!nSaldo, 20) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres, 12) & ImpreFormat(R!nMora, 7) & ImpreFormat(R!nIntPerd, 12) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            Exit Do
                        Else
                            nPuntPag = 0
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("SALDO", 13) & ImpreFormat("CAPITAL", 12) & ImpreFormat("INTERES", 12) & ImpreFormat("MORA", 8) & ImpreFormat("INT. PERD", 12) & ImpreFormat("LINEA C.", 8) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        End If
                    Else
                        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("SALDO", 13) & ImpreFormat("CAPITAL", 12) & ImpreFormat("INTERES", 12) & ImpreFormat("MORA", 8) & ImpreFormat("INT. PERD", 12) & ImpreFormat("LINEA C.", 8) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    End If
                Else
                    sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("SALDO", 13) & ImpreFormat("CAPITAL", 12) & ImpreFormat("INTERES", 12) & ImpreFormat("MORA", 8) & ImpreFormat("INT. PERD", 12) & ImpreFormat("LINEA C.", 8) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("SALDO", 13) & ImpreFormat("CAPITAL", 12) & ImpreFormat("INTERES", 12) & ImpreFormat("MORA", 8) & ImpreFormat("INT. PERD", 12) & ImpreFormat("LINEA C.", 8) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            End If
        End If
        
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!NCredito, 20) & ImpreFormat(R!sTitular, 30) & _
                                    ImpreFormat(R!nSaldo, 10) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteres, 12) & ImpreFormat(R!nMora, 7) & ImpreFormat(R!nIntPerd, 12) & ImpreFormat(R!cLineaCred, 12, 4) & ImpreFormat(R!sAnalista, 12) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    
    ImprimeCreditosCanceladosConSaldo = sCadImp
End Function

Public Function ImprimeResumenSaldosCarteraXAnalista(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatProd As Variant, ByVal pMatCond As Variant, ByVal pnCar1I As Integer, ByVal pnCar1F As Integer, _
    ByVal pnCar2I As Integer, ByVal pnCar2F As Integer, ByVal pnCar3I As Integer, ByVal pnCar3F As Integer, ByVal pnCar4I As Integer, Optional ByVal psNomCmac As String = "") As String
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String


    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaResumenSaldosCarteraPorAnalista(pMatAgencias, pdFecIni, pnMoneda, pnCar1I, pnCar1F, pnCar2I, pnCar2F, pnCar3I, pnCar3F, pnCar4I, pMatCond, pMatProd)
    Set oDCred = Nothing
    
    sCadImp = ""
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "RESUMEN DE SALDOS DE CARTERA POR ANALISTA AL " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & " Agencia : " & R!cAgencia & lnSaltoLinDoc
        sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & "Productos : "
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("ANALISTA", 9) & ImpreFormat("CARTERA TOTAL", 15) & "|" & ImpreFormat("CARTERA VENCIDA 1-7 DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA 8-15 DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA 16-30 DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA > 30 DIAS", 24) & "|" & ImpreFormat("TOTAL CART VENCIDA    ", 20) & lnSaltoLinDoc
        sCadImp = sCadImp & Space(11) & ImpreFormat("   Nro    SALDO  ", 15) & "|" & ImpreFormat("Nro       SALDO  %MORA", 24) & "|" & ImpreFormat("Nro      SALDO   %MORA", 24) & "|" & ImpreFormat("Nro      SALDO   %MORA", 24) & "|" & ImpreFormat("Nro       SALDO  %MORA", 24) & "|" & ImpreFormat("Nro       SALDO  %MOR", 20) & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
            'FINAL DE AGENCIA
            If R!cUser = "TOTAL" Then
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 9)
                sCadImp = sCadImp & ImpreFormat(R!nCar5Nro, 4, 0) & ImpreFormat(R!nCar5Saldo, 10)
                sCadImp = sCadImp & ImpreFormat(R!nCar1Nro, 5, 0) & ImpreFormat(R!nCar1Saldo, 10) & ImpreFormat(CDbl(Format(R!nCar1Saldo / R!nCar5Saldo, "#0.00")), 3)
                sCadImp = sCadImp & ImpreFormat(R!nCar2Nro, 7, 0) & ImpreFormat(R!nCar2Saldo, 10) & ImpreFormat(CDbl(Format(R!nCar2Saldo / R!nCar5Saldo, "#0.00")), 3)
                sCadImp = sCadImp & ImpreFormat(R!nCar3Nro, 8, 0) & ImpreFormat(R!nCar3Saldo, 10) & ImpreFormat(CDbl(Format(R!nCar3Saldo / R!nCar5Saldo, "#0.00")), 4)
                sCadImp = sCadImp & ImpreFormat(R!nCar4Nro, 7, 0) & ImpreFormat(R!nCar4Saldo, 10) & ImpreFormat(CDbl(Format(R!nCar4Saldo / R!nCar5Saldo, "#0.00")), 4)
                sCadImp = sCadImp & ImpreFormat(CDbl(Format(R!nCar1Nro + R!nCar2Nro + R!nCar3Nro + R!nCar4Nro, "#0.0000")), 9, 0) & ImpreFormat(CDbl(Format(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo, "#0.00")), 9) & ImpreFormat(CDbl(Format((R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo) / R!nCar5Saldo, "#0.00")), 3)
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL FINAL
                If R!cUser = "TOTAL" Then
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("TOTAL : ", 9)
                    sCadImp = sCadImp & ImpreFormat(R!nCar5Nro, 4, 0) & ImpreFormat(R!nCar5Saldo, 10)
                    sCadImp = sCadImp & ImpreFormat(R!nCar1Nro, 5, 0) & ImpreFormat(R!nCar1Saldo, 10) & ImpreFormat(Format(R!nCar1Saldo / R!nCar5Saldo, "#0.00"), 4)
                    sCadImp = sCadImp & ImpreFormat(R!nCar2Nro, 7, 0) & ImpreFormat(R!nCar2Saldo, 10) & ImpreFormat(Format(R!nCar2Saldo / R!nCar5Saldo, "#0.00"), 4)
                    sCadImp = sCadImp & ImpreFormat(R!nCar3Nro, 8, 0) & ImpreFormat(R!nCar3Saldo, 10) & ImpreFormat(Format(R!nCar3Saldo / R!nCar5Saldo, "#0.00"), 4)
                    sCadImp = sCadImp & ImpreFormat(R!nCar4Nro, 8, 0) & ImpreFormat(R!nCar4Saldo, 10) & ImpreFormat(Format(R!nCar4Saldo / R!nCar5Saldo, "#0.00"), 4)
                    sCadImp = sCadImp & ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar3Nro + R!nCar4Nro, 10, 0) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo, 9) & ImpreFormat(Format((R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo) / R!nCar5Saldo, "#0.00"), 4)
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                    Exit Do
                Else
                    sCadImp = sCadImp & Chr$(12)
                    sCadImp = sCadImp & " Agencia : " & R!cAgencia & lnSaltoLinDoc
                    sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                    sCadImp = sCadImp & "Productos : "
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("ANALISTA", 9) & ImpreFormat("CARTERA TOTAL", 15) & "|" & ImpreFormat("CARTERA VENCIDA 1-7 DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA 8-15 DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA 16-30 DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA > 30 DIAS", 24) & "|" & ImpreFormat("TOTAL CART VENCIDA    ", 20) & lnSaltoLinDoc
                    sCadImp = sCadImp & Space(11) & ImpreFormat("   Nro    SALDO  ", 15) & "|" & ImpreFormat("Nro       SALDO  %MORA", 24) & "|" & ImpreFormat("Nro      SALDO   %MORA", 24) & "|" & ImpreFormat("Nro      SALDO   %MORA", 24) & "|" & ImpreFormat("Nro       SALDO  %MORA", 24) & "|" & ImpreFormat("Nro       SALDO  %MOR", 20) & lnSaltoLinDoc
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                End If
            End If
                sCadImp = sCadImp & ImpreFormat(R!cUser, 9)
                sCadImp = sCadImp & ImpreFormat(R!nCar5Nro, 4, 0) & ImpreFormat(R!nCar5Saldo, 10)
                sCadImp = sCadImp & ImpreFormat(R!nCar1Nro, 5, 0) & ImpreFormat(R!nCar1Saldo, 10) & ImpreFormat(Format(R!nCar1Saldo / R!nCar5Saldo, "#0.00"), 6)
                sCadImp = sCadImp & ImpreFormat(R!nCar2Nro, 5, 0) & ImpreFormat(R!nCar2Saldo, 10) & ImpreFormat(Format(R!nCar2Saldo / R!nCar5Saldo, "#0.00"), 7)
                sCadImp = sCadImp & ImpreFormat(R!nCar3Nro, 5, 0) & ImpreFormat(R!nCar3Saldo, 10) & ImpreFormat(Format(R!nCar3Saldo / R!nCar5Saldo, "#0.00"), 7)
                sCadImp = sCadImp & ImpreFormat(R!nCar4Nro, 5, 0) & ImpreFormat(R!nCar4Saldo, 10) & ImpreFormat(Format(R!nCar4Saldo / R!nCar5Saldo, "#0.00"), 6)
                sCadImp = sCadImp & ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar3Nro + R!nCar4Nro, 8, 0) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo, 9) & ImpreFormat(Format((R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo) / R!nCar5Saldo, "#0.00"), 4)
                sCadImp = sCadImp & lnSaltoLinDoc
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeResumenSaldosCarteraXAnalista = sCadImp
    
End Function

Public Function ImprimeMoraInstitucional(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatProd As Variant, ByVal pMatCond As Variant, ByVal pnCar1I As Integer, ByVal pnCar1F As Integer, _
    ByVal pnCar2I As Integer, ByVal pnCar2F As Integer, ByVal pnCar3I As Integer, ByVal pnCar3F As Integer, ByVal pnCar4I As Integer, ByVal pbMoraAnt As Boolean, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaMoraInstitucional(pMatAgencias, pdFecIni, pnMoneda, pnCar1I, pnCar1F, pnCar2I, pnCar2F, pnCar3I, pnCar3F, pnCar4I, pMatCond, pMatProd, pbMoraAnt)
    Set oDCred = Nothing
    
    sCadImp = ""
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "MORA INSTITUCIONAL AL " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & " Saldo Total en " & IIf(pnMoneda = gMonedaNacional, "Soles : ", "Dolares : ") & R!nSaldoCapTotal
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & " Agencia : " & R!sAgencia & lnSaltoLinDoc
        sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & "Productos : "
        For I = 0 To UBound(pMatCond) - 1
            If CInt(pMatCond(I)) = gColocCredCondNormal Then
                sCadImp = sCadImp & "/NORMAL"
            End If
            If CInt(pMatCond(I)) = gColocCredCondParalelo Then
                sCadImp = sCadImp & "/PARALELO"
            End If
            If CInt(pMatCond(I)) = gColocCredCondRecurrente Then
                sCadImp = sCadImp & "/RECURRENTE"
            End If
        Next I
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(25, "#") & lnSaltoLinDoc
        If R!cRango = "RANGO1" Then
            sCadImp = sCadImp & "MORA DE " & pnCar1I & " A " & pnCar1F & " DIAS" & lnSaltoLinDoc
        Else
            If R!cRango = "RANGO2" Then
                sCadImp = sCadImp & "MORA DE " & pnCar2I & " A " & pnCar2F & " DIAS" & lnSaltoLinDoc
            Else
                If R!cRango = "RANGO3" Then
                    sCadImp = sCadImp & "MORA DE " & pnCar3I & " A " & pnCar3F & " DIAS" & lnSaltoLinDoc
                Else
                    sCadImp = sCadImp & "MORA MAYOR A " & pnCar4I & " DIAS" & lnSaltoLinDoc
                End If
            End If
        End If
        sCadImp = sCadImp & String(25, "#") & lnSaltoLinDoc
        sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSO", 12) & ImpreFormat("SALDO CAP", 12) & ImpreFormat("CUOTA", 8) & ImpreFormat("ATRASO", 12) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
        
        'Final de RANGO
        If R!cCtaCod = "TOTAL" Then
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & ImpreFormat("*** SUB TOTAL MORA ", 52) & ImpreFormat(R!nMontoApr, 10) & _
                                    ImpreFormat(R!nMontoDesemb, 12) & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(" ", 7) & Space(10) & ImpreFormat("Mora : " & Format((R!nSaldoCap / R!nSaldoCapTotal) * 100, "#0.00"), 12, 12) & " %" & lnSaltoLinDoc
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            
            'Final de Agencias
            If R!cRango = "TOTAL" Then
                sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("*** SUB TOTAL AGENCIA ", 52) & ImpreFormat(R!nMontoApr, 10) & _
                                    ImpreFormat(R!nMontoDesemb, 12) & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(" ", 7) & Space(10) & ImpreFormat("Mora : " & Format((R!nSaldoCap / R!nSaldoCapTotal) * 100, "#0.00"), 12, 12) & " %" & lnSaltoLinDoc
                sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL FINAL
                If R!sAgencia = "TOTAL" Then
                    sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("*** TOTAL ", 52) & ImpreFormat(R!nMontoApr, 10) & _
                                    ImpreFormat(R!nMontoDesemb, 12) & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(" ", 7) & Space(10) & ImpreFormat("Mora : " & Format((R!nSaldoCap / R!nSaldoCapTotal) * 100, "#0.00"), 12, 12) & " %" & lnSaltoLinDoc
                    sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                    Exit Do
                Else
                    sCadImp = sCadImp & Chr$(12)
                    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "MORA INSTITUCIONAL AL " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True)
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & " Saldo Total en " & IIf(pnMoneda = gMonedaNacional, "Soles : ", "Dolares : ") & R!nSaldoCapTotal
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & " Agencia : " & R!sAgencia & lnSaltoLinDoc
                    sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                    sCadImp = sCadImp & "Productos : "
                    For I = 0 To UBound(pMatCond) - 1
                        If CInt(pMatCond(I)) = gColocCredCondNormal Then
                            sCadImp = sCadImp & "/NORMAL"
                        End If
                        If CInt(pMatCond(I)) = gColocCredCondParalelo Then
                            sCadImp = sCadImp & "/PARALELO"
                        End If
                        If CInt(pMatCond(I)) = gColocCredCondRecurrente Then
                            sCadImp = sCadImp & "/RECURRENTE"
                        End If
                    Next I
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(25, "#") & lnSaltoLinDoc
                    If R!cRango = "RANGO1" Then
                        sCadImp = sCadImp & "MORA DE " & pnCar1I & " A " & pnCar1F & " DIAS" & lnSaltoLinDoc
                    Else
                        If R!cRango = "RANGO2" Then
                            sCadImp = sCadImp & "MORA DE " & pnCar2I & " A " & pnCar2F & " DIAS" & lnSaltoLinDoc
                        Else
                            If R!cRango = "RANGO3" Then
                                sCadImp = sCadImp & "MORA DE " & pnCar3I & " A " & pnCar3F & " DIAS" & lnSaltoLinDoc
                            Else
                                sCadImp = sCadImp & "MORA MAYOR A " & pnCar4I & " DIAS" & lnSaltoLinDoc
                            End If
                        End If
                    End If
                    sCadImp = sCadImp & String(25, "#") & lnSaltoLinDoc
                    sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSO", 12) & ImpreFormat("SALDO CAP", 12) & ImpreFormat("CUOTA", 8) & ImpreFormat("ATRASO", 12) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(25, "#") & lnSaltoLinDoc
                If R!cRango = "RANGO1" Then
                    sCadImp = sCadImp & "MORA DE " & pnCar1I & " A " & pnCar1F & " DIAS" & lnSaltoLinDoc
                Else
                    If R!cRango = "RANGO2" Then
                        sCadImp = sCadImp & "MORA DE " & pnCar2I & " A " & pnCar2F & " DIAS" & lnSaltoLinDoc
                    Else
                        If R!cRango = "RANGO3" Then
                            sCadImp = sCadImp & "MORA DE " & pnCar3I & " A " & pnCar3F & " DIAS" & lnSaltoLinDoc
                        Else
                            sCadImp = sCadImp & "MORA MAYOR A " & pnCar4I & " DIAS" & lnSaltoLinDoc
                        End If
                    End If
                End If
                sCadImp = sCadImp & String(25, "#") & lnSaltoLinDoc
                sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSO", 12) & ImpreFormat("SALDO CAP", 12) & ImpreFormat("CUOTA", 8) & ImpreFormat("ATRASO", 12) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            End If
        End If
        
        sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!sTitular, 30) & _
                                    ImpreFormat(R!nMontoApr, 10) & _
                                    ImpreFormat(R!nMontoDesemb, 12) & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(Trim(Str(IIf(R!nCuotaPend = 0, 1, R!nCuotaPend))) & "/" & Trim(Str(R!nCuotasApr)), 7, 7) & ImpreFormat(R!nDiasAtraso, 5, 0) & ImpreFormat(R!sAnalista, 12, 12) & lnSaltoLinDoc
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeMoraInstitucional = sCadImp
End Function

Public Function ImprimeMoraXAnalista_AtrasoPagoCuotaLibre(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatCond As Variant, ByVal pMatProd As Variant, ByVal pMatAnalistas As Variant, Optional ByVal pn_ConCuotaLibre As Boolean = False, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaMoraxAnalista_AtrasoPagoCuotaLibre(pMatAgencias, pdFecIni, pnMoneda, pMatCond, pMatProd, pMatAnalistas, pn_ConCuotaLibre)
    Set oDCred = Nothing

    sCadImp = ""
    If R.RecordCount > 0 Then

        If pn_ConCuotaLibre = True Then
            Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "ATRASO EN PAGO DE CALENDARIO DE CUOTA LIBRE AL " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True)
        ElseIf pn_ConCuotaLibre = False Then
            Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "MORA POR ANALISTA AL " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True)
        End If

        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & " Agencia : " & R!cAgencia & lnSaltoLinDoc
        sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & "Productos : "
        For I = 0 To UBound(pMatCond) - 1
            If CInt(pMatCond(I)) = gColocCredCondNormal Then
                sCadImp = sCadImp & "/NORMAL"
            End If
            If CInt(pMatCond(I)) = gColocCredCondParalelo Then
                sCadImp = sCadImp & "/PARALELO"
            End If
            If CInt(pMatCond(I)) = gColocCredCondRecurrente Then
                sCadImp = sCadImp & "/RECURRENTE"
            End If
        Next I
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & "ANALISTA : " & Trim(R!cAnalista) & lnSaltoLinDoc
        sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("  MONTO", 13) & ImpreFormat("SALDO", 7) & ImpreFormat("CUOTA", 6, 2) & ImpreFormat("CUOTA EN", 10) & ImpreFormat("FECHA", 12, 0) & ImpreFormat("SALDO", 6, 1) & ImpreFormat("DIAS ", 6) & ImpreFormat("SALDO", 6, 0) & ImpreFormat("SALDO", 12, 0) & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("                                    ", 54) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("CAPITAL", 8) & ImpreFormat("APROBADA", 9) & ImpreFormat("MORA ", 7) & ImpreFormat("VENC.", 11) & ImpreFormat("CUOTA ", 7, 0) & ImpreFormat("ATR.", 6) & ImpreFormat("MORA ", 6, 0) & ImpreFormat("GASTO      ", 12, 0) & lnSaltoLinDoc
        sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
        sCadCab = sCadImp
    End If

    Do While Not R.EOF
        'FINAL DE ANALISTA
        If R!cCtaCod = "TOTAL" Then
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & ImpreFormat("*** SUB TOTAL ANALISTA ", 104) & ImpreFormat(R!nSaldoCuota, 12) & _
                    ImpreFormat(Trim(Str(R!nDiasAtraso)), 4, 3) & ImpreFormat(R!nMora, 2) & ImpreFormat(R!NGasto, 4) & lnSaltoLinDoc
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE AGENCIA
            If R!cAnalista = "TOTAL" Then
                sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("*** SUB TOTAL AGENCIA ", 104) & ImpreFormat(R!nSaldoCuota, 12) & _
                    ImpreFormat(Trim(Str(R!nDiasAtraso)), 4, 3) & ImpreFormat(R!nMora, 2) & ImpreFormat(R!NGasto, 4) & lnSaltoLinDoc
                sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL FINAL
                If R!cAgencia = "TOTAL" Then
                    sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("*** TOTAL ", 104) & ImpreFormat(R!nSaldoCuota, 12) & _
                    ImpreFormat(Trim(Str(R!nDiasAtraso)), 4, 3) & ImpreFormat(R!nMora, 2) & ImpreFormat(R!NGasto, 4) & lnSaltoLinDoc
                    sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
                    Exit Do
                Else
                    sCadImp = sCadImp & Chr$(12)
                    sCadCab = lnSaltoLinDoc
                    sCadCab = sCadCab & lnSaltoLinDoc
                    sCadCab = sCadCab & " Agencia : " & R!cAgencia & lnSaltoLinDoc
                    sCadCab = sCadCab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                    sCadCab = sCadCab & "Productos : "
                    For I = 0 To UBound(pMatCond) - 1
                        If CInt(pMatCond(I)) = gColocCredCondNormal Then
                            sCadCab = sCadCab & "/NORMAL"
                        End If
                        If CInt(pMatCond(I)) = gColocCredCondParalelo Then
                            sCadCab = sCadCab & "/PARALELO"
                        End If
                        If CInt(pMatCond(I)) = gColocCredCondRecurrente Then
                            sCadCab = sCadCab & "/RECURRENTE"
                        End If
                    Next I
                    sCadCab = sCadCab & lnSaltoLinDoc
                    sCadCab = sCadCab & "ANALISTA : " & Trim(R!cAnalista) & lnSaltoLinDoc
                    sCadCab = sCadCab & lnSaltoLinDoc
                    sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
                    sCadCab = sCadCab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("  MONTO", 13) & ImpreFormat("SALDO", 7) & ImpreFormat("CUOTA", 6, 2) & ImpreFormat("CUOTA EN", 10) & ImpreFormat("FECHA", 12, 0) & ImpreFormat("SALDO", 6, 1) & ImpreFormat("DIAS ", 6) & ImpreFormat("SALDO", 6, 0) & ImpreFormat("SALDO", 12, 0) & lnSaltoLinDoc
                    sCadCab = sCadCab & ImpreFormat("                                    ", 54) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("CAPITAL", 8) & ImpreFormat("APROBADA", 9) & ImpreFormat("MORA ", 7) & ImpreFormat("VENC.", 11) & ImpreFormat("CUOTA ", 7, 0) & ImpreFormat("ATR.", 6) & ImpreFormat("MORA ", 6, 0) & ImpreFormat("GASTO      ", 12, 0) & lnSaltoLinDoc
                    sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & sCadCab
                End If
            Else
                    sCadImp = sCadImp & Chr$(12)
                    sCadCab = "Productos : "
                    For I = 0 To UBound(pMatCond) - 1
                        If CInt(pMatCond(I)) = gColocCredCondNormal Then
                            sCadCab = sCadCab & "/NORMAL"
                        End If
                        If CInt(pMatCond(I)) = gColocCredCondParalelo Then
                            sCadCab = sCadCab & "/PARALELO"
                        End If
                        If CInt(pMatCond(I)) = gColocCredCondRecurrente Then
                            sCadCab = sCadCab & "/RECURRENTE"
                        End If
                    Next I
                    sCadCab = sCadCab & lnSaltoLinDoc
                    sCadCab = sCadCab & "ANALISTA : " & Trim(R!cAnalista) & lnSaltoLinDoc
                    sCadCab = sCadCab & lnSaltoLinDoc
                    sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
                    sCadCab = sCadCab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("  MONTO", 13) & ImpreFormat("SALDO", 7) & ImpreFormat("CUOTA", 6, 2) & ImpreFormat("CUOTA EN", 10) & ImpreFormat("FECHA", 12, 0) & ImpreFormat("SALDO", 6, 1) & ImpreFormat("DIAS ", 6) & ImpreFormat("SALDO", 6, 0) & ImpreFormat("SALDO", 12, 0) & lnSaltoLinDoc
                    sCadCab = sCadCab & ImpreFormat("                                    ", 54) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("CAPITAL", 8) & ImpreFormat("APROBADA", 9) & ImpreFormat("MORA ", 7) & ImpreFormat("VENC.", 11) & ImpreFormat("CUOTA ", 7, 0) & ImpreFormat("ATR.", 6) & ImpreFormat("MORA ", 6, 0) & ImpreFormat("GASTO      ", 12, 0) & lnSaltoLinDoc
                    sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & sCadCab
            End If
        End If

        If Not IsNull(R!nDatos) Then
            sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!nDatos, 90, 2) & _
                          ImpreFormat(R!nSaldoCuota, 4) & ImpreFormat(R!nDiasAtraso, 5, 0) & ImpreFormat(R!nMora, 6) & ImpreFormat(R!NGasto, 4) & lnSaltoLinDoc
        End If
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeMoraXAnalista_AtrasoPagoCuotaLibre = sCadImp
End Function


Public Function ImprimeCreditosProtestados(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String
Dim sProducto As String
Dim nContProd As Long
Dim nContAge As Long

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosProtestados(pMatAgencias, pdFecIni, pnMoneda)
    Set oDCred = Nothing
    
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS PROTESTADOS AL " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadCab = " Agencia : " & R!cAgeDescripcion & lnSaltoLinDoc
        sCadCab = sCadCab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadCab = sCadCab & " Producto : " & Trim(R!cConsDescripcion)
        sCadCab = sCadCab & lnSaltoLinDoc
        sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
        sCadCab = sCadCab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 30) & ImpreFormat("SALDO CAP", 11, 0) & ImpreFormat("ULT PAGO", 10, 0) & ImpreFormat("CUOTA", 5, 0) & ImpreFormat("VENCIMIENTO", 12, 1) & ImpreFormat("CUOTA APR", 10) & ImpreFormat("CUOTA XAMORT", 12, 0) & ImpreFormat("AMORT CUO", 10, 4) & ImpreFormat("ATRASO", 6, 0) & lnSaltoLinDoc
        sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & sCadCab
        sAgencia = R!cAgeDescripcion
        sProducto = R!cConsDescripcion
        nContProd = 0
        nContAge = 0
    End If
    Do While Not R.EOF
        'FINAL DE PRODUCTO
        If sProducto <> R!cConsDescripcion And sAgencia = R!cAgeDescripcion Then
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & "Total de Productos : " & nContProd & lnSaltoLinDoc
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            nContAge = nContAge + nContProd
            nContProd = 0
            sProducto = R!cConsDescripcion
            sCadCab = " Producto : " & Trim(R!cConsDescripcion)
            sCadCab = sCadCab & lnSaltoLinDoc
            sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
            sCadCab = sCadCab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 30) & ImpreFormat("SALDO CAP", 11, 0) & ImpreFormat("ULT PAGO", 10, 0) & ImpreFormat("CUOTA", 5, 0) & ImpreFormat("VENCIMIENTO", 12, 1) & ImpreFormat("CUOTA APR", 10) & ImpreFormat("CUOTA XAMORT", 12, 0) & ImpreFormat("AMORT CUO", 10, 4) & ImpreFormat("ATRASO", 6, 0) & lnSaltoLinDoc
            sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & sCadCab
        End If
        If sAgencia <> R!cAgeDescripcion Then
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & "Total de Productos : " & nContProd & lnSaltoLinDoc
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            nContAge = nContAge + nContProd
            sProducto = R!cConsDescripcion
            
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & "Total de Productos En Agencia : " & nContAge & lnSaltoLinDoc
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            nContProd = 0
            nContAge = 0
            nPuntPag = 0
            sCadImp = sCadImp & Chr$(12)
            sAgencia = R!cAgeDescripcion
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadCab = " Agencia : " & R!cAgeDescripcion & lnSaltoLinDoc
            sCadCab = sCadCab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
            sCadCab = sCadCab & " Producto : " & Trim(R!cConsDescripcion)
            sCadCab = sCadCab & lnSaltoLinDoc
            sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
            sCadCab = sCadCab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 30) & ImpreFormat("SALDO CAP", 11, 0) & ImpreFormat("ULT PAGO", 10, 0) & ImpreFormat("CUOTA", 5, 0) & ImpreFormat("VENCIMIENTO", 12, 1) & ImpreFormat("CUOTA APR", 10) & ImpreFormat("CUOTA XAMORT", 12, 0) & ImpreFormat("AMORT CUO", 10, 4) & ImpreFormat("ATRASO", 6, 0) & lnSaltoLinDoc
            sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & sCadCab
        End If
        
        sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!cPersNombre, 25) & ImpreFormat(R!nSaldo, 10) & _
                        ImpreFormat(Mid(R!sFecha, 7, 2) & IIf(IsNull(R!sFecha), "", "/") & Mid(R!sFecha, 5, 2) & IIf(IsNull(R!sFecha), "", "/") & Mid(R!sFecha, 1, 4), 10) & _
                        ImpreFormat(R!nNroProxCuota, 4, 0) & ImpreFormat(Format(R!dVenc, "dd/mm/yyyy"), 10, 2) & ImpreFormat(R!nCuotaApr, 10) & ImpreFormat(R!nCuotaXAmort, 10) & ImpreFormat(R!nCuotaAmort, 10) & ImpreFormat(R!nDiasAtraso, 5, 0) & lnSaltoLinDoc
        nContProd = nContProd + 1
        R.MoveNext
    Loop
    
    If R.RecordCount > 0 Then
        sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & "Total de Productos : " & nContProd & lnSaltoLinDoc
        sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & "Total de Productos En Agencia : " & nContAge + nContProd & lnSaltoLinDoc
        sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
    End If
    R.Close
    Set R = Nothing
    ImprimeCreditosProtestados = sCadImp
    
End Function

Public Function ImprimeCreditosRetirados(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatCond As Variant, ByVal pMatProd As Variant, Optional ByVal psNomCmac As String = "") As String
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String
Dim nContAge As Long

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosRetirados(pMatAgencias, pdFecIni, pdFecFin, pnMoneda, pMatCond, pMatProd)
    Set oDCred = Nothing
    
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS RETIRADOS DEL " & Format(pdFecIni, "dd/mm/yyyy") & "  AL " & Format(pdFecFin, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadCab = " Agencia : " & R!cAgeDescripcion & lnSaltoLinDoc
        sCadCab = sCadCab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadCab = sCadCab & lnSaltoLinDoc
        sCadCab = sCadCab & String(130, "-") & lnSaltoLinDoc
        sCadCab = sCadCab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 40) & ImpreFormat("MONTO ", 11, 0) & ImpreFormat("CAUSA", 20, 0) & ImpreFormat("FECHA", 18, 0) & ImpreFormat("ANALISTA", 12, 0) & lnSaltoLinDoc
        sCadCab = sCadCab & String(130, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & sCadCab
        sAgencia = R!cAgeDescripcion
        nContAge = 0
    End If
    Do While Not R.EOF
        If sAgencia <> R!cAgeDescripcion Then
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & "Total de Productos En Agencia : " & nContAge & lnSaltoLinDoc
            sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
            nContAge = 0
            nPuntPag = 0
            sCadImp = sCadImp & Chr$(12)
            sAgencia = R!cAgeDescripcion
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadCab = " Agencia : " & R!cAgeDescripcion & lnSaltoLinDoc
            sCadCab = sCadCab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
            sCadCab = sCadCab & lnSaltoLinDoc
            sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
            sCadCab = sCadCab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 40) & ImpreFormat("MONTO ", 11, 0) & ImpreFormat("CAUSA", 20, 0) & ImpreFormat("FECHA", 18, 0) & ImpreFormat("ANALISTA", 12, 0) & lnSaltoLinDoc
            sCadCab = sCadCab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & sCadCab
        End If
        nContAge = nContAge + 1
        sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!cPersNombre, 34) & ImpreFormat(R!nMonto, 10) & ImpreFormat(R!cConsDescripcion, 20) & ImpreFormat(Format(R!dPrdEstado, "dd/mm/yyyy"), 20) & ImpreFormat(R!cUser, 10)
        R.MoveNext
    Loop
    If R.RecordCount > 0 Then
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & "Total de Productos En Agencia : " & nContAge & lnSaltoLinDoc
        sCadImp = sCadImp & String(145, "-") & lnSaltoLinDoc
    End If
    R.Close
    Set R = Nothing
    ImprimeCreditosRetirados = sCadImp
End Function

Public Function ImprimeCartaMorosos(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatCond As Variant, ByVal pMatProd As Variant, _
    ByVal pnDiasAtrIni As Integer, ByVal pnDiasAtrFin As Integer, ByVal sFormato As String, ByVal pMatAnalistas As Variant) As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim sCadImpTmp As String
Dim sAgencia As String
Dim sCadFormato As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCartaCreditosMorosos(pMatAgencias, pdFecIni, pnMoneda, pMatCond, pMatProd, pnDiasAtrIni, pnDiasAtrFin, pMatAnalistas)
    Set oDCred = Nothing
    sAgencia = ""
    sCadImp = ""
    sCadImpTmp = ""
    Do While Not R.EOF
        sCadFormato = sFormato
        sCadFormato = Replace(sCadFormato, "<<FECHA>>", Format(pdFecSis, "dddd,d mmmm yyyy"), , 1, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<NOMBRE>>", PstaNombre(R!cPersNombre), , 1, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<DIRECCION>>", Trim(R!cPersDireccDomicilio), , 1, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<ZONA>>", Trim(R!Zona), , 1, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<CREDITO>>", R!cCtaCod, , 3, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<CUOTA>>", Trim(Str(R!nCuota)), , 1, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<ANALISTA>>", PstaNombre(R!cAnalista), , 1, vbTextCompare)
        sCadImp = sCadImp & sCadFormato & Chr$(12)
        If R.Bookmark Mod 50 = 0 Then
            sCadImpTmp = sCadImpTmp & sCadImp
            sCadImp = ""
        End If
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    
    sCadImpTmp = sCadImpTmp & sCadImp
    ImprimeCartaMorosos = sCadImpTmp
    
End Function

Public Function ImprimeCartaInvitacionCreditoParalelo(ByVal pMatAgencias As Variant, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatCond As Variant, ByVal pMatProd As Variant, _
    ByVal pnNotaIni As Integer, ByVal pnNotaFin As Integer, ByVal sFormato As String, ByVal pMatAnalistas As Variant, _
    ByVal pnUltCuotas As Integer, ByVal pbPorc As Integer) As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim sCadImpTmp As String
Dim sAgencia As String
Dim sCadFormato As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCartaInvitacionCreditosParalelos(pMatAgencias, pnMoneda, pMatCond, pMatProd, pnNotaIni, pnNotaFin, pMatAnalistas, pnUltCuotas, pbPorc)
    Set oDCred = Nothing
    sAgencia = ""
    sCadImp = ""
    sCadImpTmp = ""
    Do While Not R.EOF
        sCadFormato = sFormato
        sCadFormato = Replace(sCadFormato, "<<FECHA>>", Format(pdFecSis, "dddd,d mmmm yyyy"), , 1, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<NOMBRE>>", PstaNombre(R!cPersNombre), , 1, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<DIRECCION>>", Trim(R!cPersDireccDomicilio), , 1, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<ZONA>>", Trim(R!Zona), , 1, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<CREDITO>>", R!cCtaCod, , 3, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<CUOTA>>", Trim(Str(R!nCuotas)), , 1, vbTextCompare)
        sCadFormato = Replace(sCadFormato, "<<ANALISTA>>", PstaNombre(R!cAnalista), , 1, vbTextCompare)
        sCadImp = sCadImp & sCadFormato & Chr$(12)
        If R.Bookmark Mod 50 = 0 Then
            sCadImpTmp = sCadImpTmp & sCadImp
            sCadImp = ""
        End If
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    
    sCadImpTmp = sCadImpTmp & sCadImp
    ImprimeCartaInvitacionCreditoParalelo = sCadImpTmp
    
End Function

Public Function ImprimeCreditosXUbicacionGeo(ByVal pMatAgencias As Variant, ByVal pnMoneda As Moneda, _
    ByVal psCodUser As String, ByVal pdFecSis As Date, ByVal psNomAge As String, ByVal pMatCond As Variant, _
    ByVal pMatProd As Variant, ByVal psUbicaciogeo As String, Optional ByVal psNomCmac As String = "", _
    Optional ByVal psCodRepo As String = "") As String
    
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String
Dim lnMontoA As Currency
Dim lnMontoSC As Currency

    lnMontoA = 0
    lnMontoSC = 0

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosPorUbicacionGeo(pMatAgencias, pnMoneda, pMatCond, pMatProd, psUbicaciogeo)
    Set oDCred = Nothing
    
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS VIGENTES POR UBICACION GEOGRAFICA : " & Trim(Mid(psUbicaciogeo, 1, 35)), psNomCmac, , True, psCodRepo)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadCab = " Agencia : " & R!cAgeDescripcion & lnSaltoLinDoc
        sCadCab = sCadCab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadCab = sCadCab & lnSaltoLinDoc
        sCadCab = sCadCab & String(160, "-") & lnSaltoLinDoc
        sCadCab = sCadCab & ImpreFormat("CREDITO", 21) & ImpreFormat("CLIENTE", 37) & ImpreFormat("DIRECCION", 23, 0) & ImpreFormat("ZONA", 26, 0) & ImpreFormat("APROBADO", 13, 0) & ImpreFormat("SALDO CAP", 11, 0) & ImpreFormat("NOTA", 5, 0) & ImpreFormat("ANA.", 6, 0) & ImpreFormat("CUOTA", 14, 0) & lnSaltoLinDoc
        sCadCab = sCadCab & String(160, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & sCadCab
        sAgencia = R!cAgeDescripcion
    End If
    
    Do While Not R.EOF
        If sAgencia <> R!cAgeDescripcion Then
            nPuntPag = 0
            sCadImp = sCadImp & Chr$(12)
            sAgencia = R!cAgeDescripcion
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadCab = " Agencia : " & R!cAgeDescripcion & lnSaltoLinDoc
            sCadCab = sCadCab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
            sCadCab = sCadCab & lnSaltoLinDoc
            sCadCab = sCadCab & String(160, "-") & lnSaltoLinDoc
            sCadCab = sCadCab & ImpreFormat("CREDITO", 21) & ImpreFormat("CLIENTE", 37) & ImpreFormat("DIRECCION", 23, 0) & ImpreFormat("ZONA", 26, 0) & ImpreFormat("APROBADO", 13, 0) & ImpreFormat("SALDO CAP", 11, 0) & ImpreFormat("NOTA", 5, 0) & ImpreFormat("ANA.", 6, 0) & ImpreFormat("CUOTA", 14, 0) & lnSaltoLinDoc
            sCadCab = sCadCab & String(160, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & sCadCab
        End If
        
        sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 20) & Space(1) & ImpreFormat(PstaNombre(R!cPersNombre), 34) & Space(1) & ImpreFormat(R!cPersDireccDomicilio, 20) & Space(1) & ImpreFormat(R!Zona, 20) & Space(1) & ImpreFormat(R!nMonto, 10) & Space(1) & ImpreFormat(R!nSaldo, 10) & Space(1) & ImpreFormat(IIf(IsNull(R!nColocNota), 0, R!nColocNota), 5, 0) & Space(1) & ImpreFormat(R!cUser, 5, 0) & Space(1) & ImpreFormat(R!nNroProxCuota & "/" & R!nCuotas, 5, 0) & lnSaltoLinDoc
        R.MoveNext
    Loop
    
    R.Close
    Set R = Nothing
    
    ImprimeCreditosXUbicacionGeo = sCadImp
End Function


Public Function ImprimeCredVig_CredVigCuotLib(ByVal pMatAgencias As Variant, _
    ByVal pdFecIni As Date, ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatCond As Variant, ByVal pMatProd As Variant, ByVal pnTipoCambio As Double, ByVal pnDiasAtrIni As Integer, ByVal pnDiasAtrFin As Integer, Optional ByVal pn_ConCuotaLibre As Boolean = False, Optional ByVal psNomCmac As String = "") As String
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCredVig_CredVigConCuotLib(pMatAgencias, pdFecIni, pnMoneda, pMatCond, pMatProd, pnDiasAtrIni, pnDiasAtrFin, pnTipoCambio, pn_ConCuotaLibre)
    'Set R = oDCred.RecuperaCreditosVigentes(pMatAgencias, pdFecIni, pnMoneda, pMatCond, pMatProd, pnDiasAtrIni, pnDiasAtrFin, pnTipoCambio, pn_ConCuotaLibre)
    Set oDCred = Nothing
    
    If R.RecordCount > 0 Then
        If pn_ConCuotaLibre = True Then
            Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS VIGENTES CON CUOTA LIBRE AL: " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True)
        Else
            Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS VIGENTES AL: " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True)
        End If
        
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 30) & _
            ImpreFormat("INSTITUCION", 16) & ImpreFormat("F. DESEMB.", 10) & ImpreFormat("PLAZO", 5) & _
            ImpreFormat("MONTO APROB", 13) & ImpreFormat("S. CAPITAL", 8) & ImpreFormat("C. MORA", 7) & _
            ImpreFormat("D. MORA", 8) & ImpreFormat("M. CUOTA", 7) & ImpreFormat("ULT. FEC PAGO", 12) & _
            ImpreFormat("INT. DEVENG", 10) & ImpreFormat("DEUDA TOT", 12) & ImpreFormat("CAP VENC", 7) & _
            ImpreFormat("CAP X VENC", 8) & ImpreFormat("GAR. HIP", 9) & ImpreFormat("OTRAS GAR", 12) & _
            ImpreFormat("REFINANC", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
    End If
    
    Do While Not R.EOF
    'FINAL DE Financiamiento
        If Trim(R!cCtaCod) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL FINANCIAMIENTO ", 42) & ImpreFormat("CASOS : " & R!nCasos, 45) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 10) & ImpreFormat(R!IntDeveng, 47) & ImpreFormat(R!nDeudaTotal, 10) & ImpreFormat(R!nCapVenc, 8) & ImpreFormat(R!nCapXVencer, 8) & ImpreFormat(R!nMontoGarHip, 8) & ImpreFormat(R!nMontoOtrasGarant, 10) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE PLAZO
            If R!sFuentefinan = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PLAZO ", 42) & ImpreFormat("CASOS : " & R!nCasos, 45) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 10) & ImpreFormat(R!IntDeveng, 47) & ImpreFormat(R!nDeudaTotal, 10) & ImpreFormat(R!nCapVenc, 8) & ImpreFormat(R!nCapXVencer, 8) & ImpreFormat(R!nMontoGarHip, 8) & ImpreFormat(R!nMontoOtrasGarant, 10) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL DE PRODUCTO
                If R!sPlazo = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PRODUCTO ", 42) & ImpreFormat("CASOS : " & R!nCasos, 45) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 10) & ImpreFormat(R!IntDeveng, 47) & ImpreFormat(R!nDeudaTotal, 10) & ImpreFormat(R!nCapVenc, 8) & ImpreFormat(R!nCapXVencer, 8) & ImpreFormat(R!nMontoGarHip, 8) & ImpreFormat(R!nMontoOtrasGarant, 10) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    R.MoveNext
                    If R.EOF Then
                        Exit Do
                    End If
                    'FINAL DE AGENCIAS
                    If R!sProducto = "TOTAL" Then
                        sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 42) & ImpreFormat("CASOS : " & R!nCasos, 45) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 10) & ImpreFormat(R!IntDeveng, 47) & ImpreFormat(R!nDeudaTotal, 10) & ImpreFormat(R!nCapVenc, 8) & ImpreFormat(R!nCapXVencer, 8) & ImpreFormat(R!nMontoGarHip, 8) & ImpreFormat(R!nMontoOtrasGarant, 10) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        R.MoveNext
                        If R.EOF Then
                            Exit Do
                        End If
                        'TOTAL FINAL
                        If R!sAgencia = "TOTAL" Then
                            sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 42) & ImpreFormat("CASOS : " & R!nCasos, 45) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 10) & ImpreFormat(R!IntDeveng, 47) & ImpreFormat(R!nDeudaTotal, 10) & ImpreFormat(R!nCapVenc, 8) & ImpreFormat(R!nCapXVencer, 8) & ImpreFormat(R!nMontoGarHip, 8) & ImpreFormat(R!nMontoOtrasGarant, 10) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            Exit Do
                        Else
                            nPuntPag = 0
                            sCadImp = sCadImp & Chr$(12)
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 30) & _
                                ImpreFormat("INSTITUCION", 16) & ImpreFormat("F. DESEMB.", 10) & ImpreFormat("PLAZO", 5) & _
                                ImpreFormat("MONTO APROB", 13) & ImpreFormat("S. CAPITAL", 8) & ImpreFormat("C. MORA", 7) & _
                                ImpreFormat("D. MORA", 8) & ImpreFormat("M. CUOTA", 7) & ImpreFormat("ULT. FEC PAGO", 12) & _
                                ImpreFormat("INT. DEVENG", 10) & ImpreFormat("DEUDA TOT", 12) & ImpreFormat("CAP VENC", 7) & _
                                ImpreFormat("CAP X VENC", 8) & ImpreFormat("GAR. HIP", 9) & ImpreFormat("OTRAS GAR", 12) & _
                                ImpreFormat("REFINANC", 12) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                        End If
                    Else
                        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 30) & _
                            ImpreFormat("INSTITUCION", 16) & ImpreFormat("F. DESEMB.", 10) & ImpreFormat("PLAZO", 5) & _
                            ImpreFormat("MONTO APROB", 13) & ImpreFormat("S. CAPITAL", 8) & ImpreFormat("C. MORA", 7) & _
                            ImpreFormat("D. MORA", 8) & ImpreFormat("M. CUOTA", 7) & ImpreFormat("ULT. FEC PAGO", 12) & _
                            ImpreFormat("INT. DEVENG", 10) & ImpreFormat("DEUDA TOT", 12) & ImpreFormat("CAP VENC", 7) & _
                            ImpreFormat("CAP X VENC", 8) & ImpreFormat("GAR. HIP", 9) & ImpreFormat("OTRAS GAR", 12) & _
                            ImpreFormat("REFINANC", 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                    End If
                Else
                    sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 30) & _
                        ImpreFormat("INSTITUCION", 16) & ImpreFormat("F. DESEMB.", 10) & ImpreFormat("PLAZO", 5) & _
                        ImpreFormat("MONTO APROB", 13) & ImpreFormat("S. CAPITAL", 8) & ImpreFormat("C. MORA", 7) & _
                        ImpreFormat("D. MORA", 8) & ImpreFormat("M. CUOTA", 7) & ImpreFormat("ULT. FEC PAGO", 12) & _
                        ImpreFormat("INT. DEVENG", 10) & ImpreFormat("DEUDA TOT", 12) & ImpreFormat("CAP VENC", 7) & _
                        ImpreFormat("CAP X VENC", 8) & ImpreFormat("GAR. HIP", 9) & ImpreFormat("OTRAS GAR", 12) & _
                        ImpreFormat("REFINANC", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 30) & _
                        ImpreFormat("INSTITUCION", 16) & ImpreFormat("F. DESEMB.", 10) & ImpreFormat("PLAZO", 5) & _
                        ImpreFormat("MONTO APROB", 13) & ImpreFormat("S. CAPITAL", 8) & ImpreFormat("C. MORA", 7) & _
                        ImpreFormat("D. MORA", 8) & ImpreFormat("M. CUOTA", 7) & ImpreFormat("ULT. FEC PAGO", 12) & _
                        ImpreFormat("INT. DEVENG", 10) & ImpreFormat("DEUDA TOT", 12) & ImpreFormat("CAP VENC", 7) & _
                        ImpreFormat("CAP X VENC", 8) & ImpreFormat("GAR. HIP", 9) & ImpreFormat("OTRAS GAR", 12) & _
                        ImpreFormat("REFINANC", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(230, "-") & lnSaltoLinDoc
            End If
        End If
        
'        If R!cCtaCod <> "TOTAL" Then
            sCadImp = sCadImp & lsTab & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(PstaNombre(R!sTitular), 25) & _
                                    ImpreFormat(Mid(IIf(IsNull(R!sInstitucion), ".", R!sInstitucion), 1, 22), 23) & _
                                    ImpreFormat(Format(R!dVigencia, "dd/mm/yyyy"), 12, 0) & ImpreFormat(R!sPlazoCuota, 5, 0) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 10) _
                                    & ImpreFormat(R!nCuotaMora, 8, 0) & ImpreFormat(R!nDiasAtraso, 8, 0) & ImpreFormat(R!nMontoCuota, 8) & ImpreFormat(Format(R!dFecUtlPago, "dd/mm/yyyy"), 12) _
                                    & ImpreFormat(R!IntDeveng, 6) & ImpreFormat(R!nDeudaTotal, 10) & ImpreFormat(IIf(IsNull(R!nCapVenc), 0, R!nCapVenc), 8) _
                                    & ImpreFormat(R!nCapXVencer, 8) & ImpreFormat(R!nMontoGarHip, 8) & ImpreFormat(R!nMontoOtrasGarant, 10) & ImpreFormat(R!cEstadoDesc, 12) & lnSaltoLinDoc
'        End If
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeCredVig_CredVigCuotLib = sCadImp

End Function


Public Function ImprimeCreditosXInstitucion(ByVal pMatAgencias As Variant, _
    ByVal pdFecIni As Date, ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pnTipoOrden As Integer, ByVal pnIncluyeMora As Integer, _
    Optional ByVal psNomCmac As String = "", Optional psCodRepo As String = "") As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String
Dim oGastos As NGasto
Dim oNegCredito As NCredito
Dim MatCalend As Variant
Dim MatGastosCancelacion As Variant
Dim nMontoGastoCargo As Double
Dim nNumGastosCancel As Integer

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosXInstitucion(pMatAgencias, pnMoneda, pdFecIni, pnTipoOrden, pnIncluyeMora)
    Set oDCred = Nothing
    
    Set oGastos = New NGasto
    Set oNegCredito = New NCredito
    
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "LISTADO PARA DESCUENTO POR PLANILLA AL : " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True, psCodRepo)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 25) & ImpreFormat("COD. MODULAR", 20) & _
            ImpreFormat("DESEMBOLSO", 15) & ImpreFormat("APROBADO", 12) & ImpreFormat("SALDO CAP", 10) & ImpreFormat("C. PEND", 7) & _
            ImpreFormat("CUOTA", 8) & ImpreFormat("FEC. PAGO", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
    End If
    
    Do While Not R.EOF
    'FINAL DE INSTITUCION
        If Trim(R!cCtaCod) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL INSTITUCION ", 42) & ImpreFormat("CASOS : " & R!nCasos, 39) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(R!nCuota, 16) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            
            'FINAL DE AGENCIA
            If R!cInstitucion = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 42) & ImpreFormat("CASOS : " & R!nCasos, 39) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(R!nCuota, 16) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'TOTAL FINAL
                If R!sAgencia = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 42) & ImpreFormat("CASOS : " & R!nCasos, 39) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(R!nCuota, 16) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    Exit Do
                Else
                    nPuntPag = 0
                    sCadImp = sCadImp & Chr$(12)
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("COD. MODULAR", 20) & _
                        ImpreFormat("DESEMBOLSO", 12) & ImpreFormat("APROBADO", 12) & ImpreFormat("SALDO CAP", 12) & ImpreFormat("C. PEND", 5) & _
                        ImpreFormat("CUOTA", 12) & ImpreFormat("FEC. PAGO", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                End If
        
            Else
                sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("COD. MODULAR", 20) & _
                ImpreFormat("DESEMBOLSO", 12) & ImpreFormat("APROBADO", 12) & ImpreFormat("SALDO CAP", 12) & ImpreFormat("C. PEND", 5) & _
                        ImpreFormat("CUOTA", 12) & ImpreFormat("FEC. PAGO", 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            End If
        End If
        
        
        MatCalend = oNegCredito.RecuperaMatrizCalendarioPendiente(R!cCtaCod)
        MatGastosCancelacion = oGastos.GeneraCalendarioGastos(Array(0), Array(0), nNumGastosCancel, gdFecSis, R!cCtaCod, 1, "CA", , , R!nMontoApr, oNegCredito.MatrizMontoCapitalAPagar(MatCalend, gdFecSis), MatCalend(0, 2), , , , , R!nDiasAtraso)
        nMontoGastoCargo = MontoTotalGastosGenerado(MatGastosCancelacion, nNumGastosCancel, Array("CA", "PA", ""))
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(PstaNombre(R!sTitular), 25) & _
                                    ImpreFormat(R!cCodModular, 20) & ImpreFormat(Format(R!dFecDesemb, "dd/mm/yyyy"), 12) & ImpreFormat(R!nMontoApr, 10) _
                                    & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(R!sCuotaPend, 6, 4) & ImpreFormat(R!nCuota + nMontoGastoCargo, 6) & ImpreFormat(Format(R!dFecPago, "dd/mm/yyyy"), 12) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeCreditosXInstitucion = sCadImp

End Function

Public Function ImprimeMoraXInstitucion(ByVal pMatAgencias As Variant, _
    ByVal pdFecIni As Date, ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, pMatInstituciones As Variant, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaMoraXInstituciones(pMatAgencias, pnMoneda, pdFecIni, pMatInstituciones)
    Set oDCred = Nothing
    
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "MORA POR INSTITUCIONES AL : " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 31) & ImpreFormat("MONTO", 13) & ImpreFormat("SALDO", 9) & ImpreFormat("SALDO DE", 10, 0) & ImpreFormat("DIAS", 5, 3) & ImpreFormat("SALDO", 7) & ImpreFormat("FECHA DE", 10) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat(" ", 51) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("CAPITAL", 10) & ImpreFormat("CUOTA", 10) & ImpreFormat("MORA", 6) & ImpreFormat("MORA", 7) & ImpreFormat("VENCIMIENTO", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
    End If
    
    Do While Not R.EOF
    'FINAL DE INSTITUCION
        If Trim(R!cCtaCod) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL INSTITUCION ", 26) & ImpreFormat("CASOS : " & R!nCasos, 19) & ImpreFormat(R!nMontoDesemb, 10) & ImpreFormat(R!nSaldoCap, 10) & ImpreFormat(R!nSaldoCuota, 10) & ImpreFormat(R!nSaldoMora, 16) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            
            'FINAL DE AGENCIA
            If R!cInstitucion = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 26) & ImpreFormat("CASOS : " & R!nCasos, 19) & ImpreFormat(R!nMontoDesemb, 10) & ImpreFormat(R!nSaldoCap, 10) & ImpreFormat(R!nSaldoCuota, 10) & ImpreFormat(R!nSaldoMora, 16) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'TOTAL FINAL
                If R!sAgencia = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 26) & ImpreFormat("CASOS : " & R!nCasos, 19) & ImpreFormat(R!nMontoDesemb, 10) & ImpreFormat(R!nSaldoCap, 10) & ImpreFormat(R!nSaldoCuota, 10) & ImpreFormat(R!nSaldoMora, 16) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    Exit Do
                Else
                    nPuntPag = 0
                    sCadImp = sCadImp & Chr$(12)
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 31) & ImpreFormat("MONTO", 13) & ImpreFormat("SALDO", 9) & ImpreFormat("SALDO DE", 10, 0) & ImpreFormat("DIAS", 5, 3) & ImpreFormat("SALDO", 7) & ImpreFormat("FECHA DE", 10) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat(" ", 51) & ImpreFormat("DESEMBOLSADO", 25) & ImpreFormat("CAPITAL", 10) & ImpreFormat("CUOTA", 10) & ImpreFormat("MORA", 6) & ImpreFormat("MORA", 7) & ImpreFormat("VENCIMIENTO", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                End If
        
            Else
                sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 31) & ImpreFormat("MONTO", 13) & ImpreFormat("SALDO", 9) & ImpreFormat("SALDO DE", 10, 0) & ImpreFormat("DIAS", 5, 3) & ImpreFormat("SALDO", 7) & ImpreFormat("FECHA DE", 10) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat(" ", 39) & ImpreFormat("DESEMBOLSADO", 25) & ImpreFormat("CAPITAL", 10) & ImpreFormat("CUOTA", 10) & ImpreFormat("MORA", 6) & ImpreFormat("MORA", 7) & ImpreFormat("VENCIMIENTO", 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            End If
        End If
        
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!sTitular, 25) & _
                                    ImpreFormat(R!nMontoDesemb, 10) & ImpreFormat(R!nSaldoCap, 10) _
                                    & ImpreFormat(R!nSaldoCuota, 10) & ImpreFormat(Trim(Str(R!nDiasAtraso)), 4, 6) & ImpreFormat(R!nSaldoMora, 6) & ImpreFormat(Format(R!dFecVencimiento, "dd/mm/yyyy"), 12) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeMoraXInstitucion = sCadImp

End Function

Public Function ImprimeResumenSaldosCarteraXAnalistaConsumo(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatProd As Variant, ByVal pMatCond As Variant, ByVal pnCar1I As Integer, ByVal pnCar1F As Integer, _
    ByVal pnCar2I As Integer, ByVal pnCar2F As Integer, ByVal pnCar4I As Integer, ByVal pdFecha As Date, Optional ByVal psNomCmac As String = "") As String
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String


    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaResumenSaldosCarteraPorAnalistaConsumo(pMatAgencias, pnMoneda, pnCar1I, pnCar1F, pnCar2I, pnCar2F, pnCar4I, pMatCond, pMatProd, pdFecIni)
    Set oDCred = Nothing
    
    sCadImp = ""
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "RESUMEN DE SALDOS DE CARTERA POR ANALISTA AL " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & " Agencia : " & R!cAge & lnSaltoLinDoc
        sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & "Productos : "
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("ANALISTA", 9) & ImpreFormat("CARTERA TOTAL", 15) & "|" & ImpreFormat("CARTERA VENCIDA " & pnCar1I & "-" & pnCar1F & " DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA " & pnCar2I & "-" & pnCar2F & " DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA > " & pnCar4I & "DIAS", 24) & "|" & ImpreFormat("TOTAL CART VENCIDA    ", 20) & lnSaltoLinDoc
        sCadImp = sCadImp & Space(11) & ImpreFormat("   Nro    SALDO  ", 15) & "|" & ImpreFormat("Nro       SALDO  %MORA", 24) & "|" & ImpreFormat("Nro      SALDO   %MORA", 24) & "|" & ImpreFormat("Nro      SALDO   %MORA", 24) & "|" & ImpreFormat("Nro       SALDO  %MOR", 20) & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
            'FINAL DE AGENCIA
            If R!cUser = "TOTAL" Then
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 9)
                sCadImp = sCadImp & ImpreFormat(R!nCar5Nro, 4, 0) & ImpreFormat(R!nCar5Saldo, 10)
                sCadImp = sCadImp & ImpreFormat(R!nCar1Nro, 5, 0) & ImpreFormat(R!nCar1Saldo, 10) & ImpreFormat(R!nCar1Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & ImpreFormat(R!nCar2Nro, 5, 0) & ImpreFormat(R!nCar2Saldo, 10) & ImpreFormat(R!nCar2Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & ImpreFormat(R!nCar4Nro, 5, 0) & ImpreFormat(R!nCar4Saldo, 10) & ImpreFormat(R!nCar4Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar4Nro, 8, 0) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar4Saldo, 9) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar4Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL FINAL
                If R!cUser = "TOTAL" Then
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("TOTAL : ", 9)
                    sCadImp = sCadImp & ImpreFormat(R!nCar5Nro, 4, 0) & ImpreFormat(R!nCar5Saldo, 10)
                    sCadImp = sCadImp & ImpreFormat(R!nCar1Nro, 5, 0) & ImpreFormat(R!nCar1Saldo, 10) & ImpreFormat(R!nCar1Saldo / R!nCar5Saldo, 4)
                    sCadImp = sCadImp & ImpreFormat(R!nCar2Nro, 5, 0) & ImpreFormat(R!nCar2Saldo, 10) & ImpreFormat(R!nCar2Saldo / R!nCar5Saldo, 4)
                    sCadImp = sCadImp & ImpreFormat(R!nCar4Nro, 5, 0) & ImpreFormat(R!nCar4Saldo, 10) & ImpreFormat(R!nCar4Saldo / R!nCar5Saldo, 4)
                    sCadImp = sCadImp & ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar4Nro, 8, 0) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar4Saldo, 9) & ImpreFormat((R!nCar1Saldo + R!nCar2Saldo + R!nCar4Saldo) / R!nCar5Saldo, 4)
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                    Exit Do
                Else
                    sCadImp = sCadImp & Chr$(12)
                    sCadImp = sCadImp & " Agencia : " & R!cAgencia & lnSaltoLinDoc
                    sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                    sCadImp = sCadImp & "Productos : "
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("ANALISTA", 9) & ImpreFormat("CARTERA TOTAL", 15) & "|" & ImpreFormat("CARTERA VENCIDA " & pnCar1I & "-" & pnCar1F & " DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA " & pnCar2I & "-" & pnCar2F & " DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA > " & pnCar4I & "DIAS", 24) & "|" & ImpreFormat("TOTAL CART VENCIDA    ", 20) & lnSaltoLinDoc
                    sCadImp = sCadImp & Space(11) & ImpreFormat("   Nro    SALDO  ", 15) & "|" & ImpreFormat("Nro       SALDO  %MORA", 24) & "|" & ImpreFormat("Nro      SALDO   %MORA", 24) & "|" & ImpreFormat("Nro      SALDO   %MORA", 24) & "|" & ImpreFormat("Nro       SALDO  %MOR", 20) & lnSaltoLinDoc
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                End If
            End If
                sCadImp = sCadImp & ImpreFormat(R!cUser, 9)
                sCadImp = sCadImp & ImpreFormat(R!nCar5Nro, 4, 0) & ImpreFormat(R!nCar5Saldo, 10)
                sCadImp = sCadImp & ImpreFormat(R!nCar1Nro, 5, 0) & ImpreFormat(R!nCar1Saldo, 10) & ImpreFormat(R!nCar1Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & ImpreFormat(R!nCar2Nro, 5, 0) & ImpreFormat(R!nCar2Saldo, 10) & ImpreFormat(R!nCar2Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & ImpreFormat(R!nCar4Nro, 5, 0) & ImpreFormat(R!nCar4Saldo, 10) & ImpreFormat(R!nCar4Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar4Nro, 8, 0) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar4Saldo, 9) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar4Saldo / R!nCar5Saldo, 4)
                sCadImp = sCadImp & lnSaltoLinDoc
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeResumenSaldosCarteraXAnalistaConsumo = sCadImp
    
End Function


Public Function ImprimeResumenSaldosCarteraXInstitucionConsumo(ByVal pMatAgencias As Variant, _
    ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatProd As Variant, Optional ByVal psNomCmac As String = "", Optional ByVal psCodRepo As String = "") As String
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String


    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaResumenSaldosCarteraPorInstitucionConsumo(pMatAgencias, pMatProd, pdFecSis)
    Set oDCred = Nothing
    
    sCadImp = ""
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "RESUMEN DE SALDOS DE CARTERA POR INSTITUCION AL " & Format(pdFecSis, "dd/mm/yyyy"), psNomCmac, , True, psCodRepo)
        sCadImp = sCadImp & " Agencia : " & R!sAgencia & lnSaltoLinDoc
        sCadImp = sCadImp & " Moneda : SOLES " & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & Space(32) & ImpreFormat("¦                  C A R T E R A    V I G E N T E               ¦          VENCIDOS > 30 DIAS     ¦", 125) & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("     INSTITUCION                     ¦           S O L E S           ¦        D O L A R E S          ¦     SOLES      ¦     DOLARES    ¦", 125) & lnSaltoLinDoc
        sCadImp = sCadImp & Space(32) & ImpreFormat("¦ NRO ¦ DESEMBOLS. ¦ SALDO CAP. ¦ NRO ¦ DESEMBOLS. ¦ SALDO CAP. ¦ NRO ¦ CAP VENC.¦ NRO ¦ CAP VENC.¦", 125) & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
            'FINAL DE AGENCIA
            If R!sInstitucion = "TOTAL" Then
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 21) & " CASOS : " & ImpreFormat(R!nNroCreditosMN, 5, 0) & ImpreFormat(R!nDesembolsadoMN, 10) & ImpreFormat(R!nSaldoCapMN, 12)
                sCadImp = sCadImp & ImpreFormat(R!nNroCreditosME, 5, 0) & ImpreFormat(R!nDesembolsadoME, 10) & ImpreFormat(R!nSaldoCapME, 10)
                sCadImp = sCadImp & ImpreFormat(R!nNroCreditos30MN, 5, 0) & ImpreFormat(R!nSaldoCap30MN, 10)
                sCadImp = sCadImp & ImpreFormat(R!nNroCreditos30ME, 5, 0) & ImpreFormat(R!nSaldoCap30ME, 10)
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL FINAL
                If R!sAgencia = "TOTAL" Then
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("TOTAL : ", 21) & " CASOS : " & ImpreFormat(R!nNroCreditosMN, 5, 0) & ImpreFormat(R!nDesembolsadoMN, 10) & ImpreFormat(R!nSaldoCapMN, 12)
                    sCadImp = sCadImp & ImpreFormat(R!nNroCreditosME, 5, 0) & ImpreFormat(R!nDesembolsadoME, 10) & ImpreFormat(R!nSaldoCapME, 10)
                    sCadImp = sCadImp & ImpreFormat(R!nNroCreditos30MN, 5, 0) & ImpreFormat(R!nSaldoCap30MN, 10)
                    sCadImp = sCadImp & ImpreFormat(R!nNroCreditos30ME, 5, 0) & ImpreFormat(R!nSaldoCap30ME, 10)
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                    Exit Do
                Else
                    sCadImp = sCadImp & Chr$(12)
                    sCadImp = sCadImp & " Agencia : " & R!sAgencia & lnSaltoLinDoc
                    sCadImp = sCadImp & " Moneda : SOLES " & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & Space(32) & ImpreFormat("¦                  C A R T E R A    V I G E N T E               ¦          VENCIDOS > 30 DIAS     ¦", 125) & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("     INSTITUCION                     ¦           S O L E S           ¦        D O L A R E S          ¦     SOLES      ¦     DOLARES    ¦", 125) & lnSaltoLinDoc
                    sCadImp = sCadImp & Space(32) & ImpreFormat("¦ NRO ¦ DESEMBOLS. ¦ SALDO CAP. ¦ NRO ¦ DESEMBOLS. ¦ SALDO CAP. ¦ NRO ¦ CAP VENC.¦ NRO ¦ CAP VENC.¦", 125) & lnSaltoLinDoc
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                End If
            End If
                sCadImp = sCadImp & ImpreFormat(PstaNombre(R!sInstitucion), 30)
                sCadImp = sCadImp & ImpreFormat(R!nNroCreditosMN, 5, 0) & ImpreFormat(R!nDesembolsadoMN, 10) & ImpreFormat(R!nSaldoCapMN, 12)
                sCadImp = sCadImp & ImpreFormat(R!nNroCreditosME, 5, 0) & ImpreFormat(R!nDesembolsadoME, 10) & ImpreFormat(R!nSaldoCapME, 10)
                sCadImp = sCadImp & ImpreFormat(R!nNroCreditos30MN, 5, 0) & ImpreFormat(R!nSaldoCap30MN, 10)
                sCadImp = sCadImp & ImpreFormat(R!nNroCreditos30ME, 5, 0) & ImpreFormat(R!nSaldoCap30ME, 10)
                sCadImp = sCadImp & lnSaltoLinDoc
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeResumenSaldosCarteraXInstitucionConsumo = sCadImp
    
End Function

Public Function ImprimePagosConCheque(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatProd As Variant, ByVal pMatCond As Variant, _
    ByVal pnTipoBusq As Integer, ByVal psNroDoc As String, Optional ByVal psNomCmac As String = "") As String
    
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim sAgencia As String
Dim sInstitucion As String
Dim nSubTotal As Double
Dim nTotal As Double
Dim nTotalRep As Double
Dim nContCasos As Long
Dim nContTotalCasos As Long
Dim nContTotalCasosRep As Long
Dim bEnc As Boolean

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaPagoConCheque(pMatAgencias, pnMoneda, pMatCond, pMatProd, pdFecIni, pdFecFin, pnTipoBusq, psNroDoc)
    Set oDCred = Nothing
    bEnc = False
    
    sCadImp = ""
    nSubTotal = 0
    nTotal = 0
    nContCasos = 0
    nContTotalCasos = 0
    If R.RecordCount > 0 Then
        bEnc = True
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "PAGOS CON CHEQUE AL " & Format(pdFecSis, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & " Agencia : " & R!sAgencia & lnSaltoLinDoc
        sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & " Institucion : " & PstaNombre(R!cInstitucion)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("Credito                            Cliente                    Operacion               Nro Docum.     Monto Trans.  User    Fecha Trans", 140) & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sAgencia = R!sAgencia
        sInstitucion = R!cInstitucion
    End If
    Do While Not R.EOF
        If sInstitucion <> R!cInstitucion Then
            sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContCasos, 5, 0) & ImpreFormat(nSubTotal, 20)
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            nTotal = nTotal + nSubTotal
            nContTotalCasos = nContTotalCasos + nContCasos
            nContCasos = 0
            nSubTotal = 0
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            If sAgencia <> R!sAgencia Then
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContTotalCasos, 5, 0) & ImpreFormat(nTotal, 20)
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                nContTotalCasosRep = nContTotalCasosRep + nContTotalCasos
                nTotalRep = nTotalRep + nTotal
                nTotal = 0
                nContTotalCasos = 0
                nContCasos = 0
                nSubTotal = 0
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                
                sCadImp = sCadImp & " Agencia : " & R!sAgencia & lnSaltoLinDoc
                sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                sCadImp = sCadImp & " Institucion : " & PstaNombre(R!cInstitucion)
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("Credito                            Cliente                    Operacion               Nro Docum.     Monto Trans.  User    Fecha Trans", 140) & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sAgencia = R!sAgencia
                sInstitucion = R!cInstitucion
            Else
                sCadImp = sCadImp & " Institucion : " & PstaNombre(R!cInstitucion)
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("Credito                            Cliente                    Operacion               Nro Docum.     Monto Trans.  User    Fecha Trans", 140) & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sAgencia = R!sAgencia
                sInstitucion = R!cInstitucion
            End If
                               
        End If
        nContCasos = nContCasos + 1
        nSubTotal = nSubTotal + R!nMonto
        
        sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(Trim(R!sTitular), 30) _
                & ImpreFormat(IIf(IsNull(R!cOperacion), "", R!cOperacion), 30) & ImpreFormat(IIf(IsNull(R!cNroDoc), "", R!cNroDoc), 12) _
                & ImpreFormat(R!nMonto, 10) & ImpreFormat(R!cUser, 5, 5) & ImpreFormat(R!DMov, 20) & lnSaltoLinDoc
        
        R.MoveNext
    Loop
    R.Close
    If bEnc Then
        
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContCasos, 5, 0) & ImpreFormat(nSubTotal, 73)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        
        nTotal = nTotal + nSubTotal
        nContTotalCasos = nContTotalCasos + nContCasos
        nContCasos = 0
        nSubTotal = 0
        
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContTotalCasos, 5, 0) & ImpreFormat(nTotal, 73)
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        
        nContTotalCasosRep = nContTotalCasosRep + nContTotalCasos
        nTotalRep = nTotalRep + nTotal
        nTotal = 0
        nContTotalCasos = 0
        nContCasos = 0
        nSubTotal = 0
        
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContTotalCasosRep, 5, 0) & ImpreFormat(nTotalRep, 73)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        ImprimePagosConCheque = sCadImp
    End If
End Function

Public Function ImprimePagosDeOtraAgencia(ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal psCodAge As String, Optional ByVal psCodCmact As String, Optional ByVal psNomCmac As String = "") As String
    
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim sAgencia As String
Dim nSubTotal As Double
Dim nTotal As Double
Dim nContCasos As Long
Dim nContTotalCasos As Long
Dim bEnc As Boolean

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaPagoDeOtrasAgencias(pnMoneda, pdFecIni, pdFecFin, psCodAge, psCodCmact)
    Set oDCred = Nothing
    bEnc = False
    
    sCadImp = ""
    nSubTotal = 0
    nTotal = 0
    nContCasos = 0
    nContTotalCasos = 0
    If R.RecordCount > 0 Then
        bEnc = True
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "PAGOS DE OTRAS AGENCIAS DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & " Agencia : " & R!sAgencia & lnSaltoLinDoc
        sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("  Credito                       Operacion               Nro Docum.     Monto Trans.  User    Fecha Trans.", 140) & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sAgencia = R!sAgencia
    End If
    Do While Not R.EOF
        If sAgencia <> R!sAgencia Then
            sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContCasos, 5, 0) & ImpreFormat(nSubTotal, 20)
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            nTotal = nTotal + nSubTotal
            nContTotalCasos = nContTotalCasos + nContCasos
            nContCasos = 0
            nSubTotal = 0
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            sCadImp = sCadImp & " Agencia : " & R!sAgencia & lnSaltoLinDoc
            sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & ImpreFormat("  Credito                       Operacion               Nro Docum.     Monto Trans.  User    Fecha Trans.", 140) & lnSaltoLinDoc
            sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            sAgencia = R!sAgencia
        End If
        nContCasos = nContCasos + 1
        nSubTotal = nSubTotal + R!nMonto
        
        sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(IIf(IsNull(R!cOperacion), "", R!cOperacion), 30) & ImpreFormat(IIf(IsNull(R!cNroDoc), "", R!cNroDoc), 12) _
                & ImpreFormat(R!nMonto, 10) & ImpreFormat(R!cUser, 5, 5) & ImpreFormat(R!DMov, 20) & lnSaltoLinDoc
        
        R.MoveNext
    Loop
    R.Close
    If bEnc Then
        
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContCasos, 5, 0) & ImpreFormat(nSubTotal, 41)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        
        nTotal = nTotal + nSubTotal
        nContTotalCasos = nContTotalCasos + nContCasos
        nContCasos = 0
        nSubTotal = 0
        
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContTotalCasos, 5, 0) & ImpreFormat(nTotal, 41)
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
    End If
    ImprimePagosDeOtraAgencia = sCadImp
End Function

Public Function ImprimePagosENOtrasAgencias(ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal psCodAge As String, ByVal psCodCmact As String, _
    ByVal pMatProd As Variant, ByVal pMatCond As Variant, Optional ByVal psNomCmac As String = "") As String
    
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim sAgencia As String
Dim nSubTotal As Double
Dim nTotal As Double
Dim nContCasos As Long
Dim nContTotalCasos As Long
Dim bEnc As Boolean

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaPagoENOtrasAgencias(pnMoneda, pdFecIni, pdFecFin, psCodAge, psCodCmact, pMatCond, pMatProd)
    Set oDCred = Nothing
    bEnc = False
    
    sCadImp = ""
    nSubTotal = 0
    nTotal = 0
    nContCasos = 0
    nContTotalCasos = 0
    If R.RecordCount > 0 Then
        bEnc = True
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "PAGOS EN OTRAS AGENCIAS DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & " Agencia : " & R!sAgencia & lnSaltoLinDoc
        sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("  Credito                        Cliente                       Operacion                   Nro Docum.     Monto Trans.  User    Fecha Trans.", 140) & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sAgencia = R!sAgencia
    End If
    Do While Not R.EOF
        If sAgencia <> R!sAgencia Then
            sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContCasos, 5, 0) & ImpreFormat(nSubTotal, 20)
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            nTotal = nTotal + nSubTotal
            nContTotalCasos = nContTotalCasos + nContCasos
            nContCasos = 0
            nSubTotal = 0
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            sCadImp = sCadImp & " Agencia : " & R!sAgencia & lnSaltoLinDoc
            sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & ImpreFormat("  Credito                        Cliente                       Operacion                   Nro Docum.     Monto Trans.  User    Fecha Trans.", 140) & lnSaltoLinDoc
            sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            sAgencia = R!sAgencia
        End If
        nContCasos = nContCasos + 1
        nSubTotal = nSubTotal + R!nMonto
        
        sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!cPersNombre, 30) & ImpreFormat(IIf(IsNull(R!cOperacion), "", R!cOperacion), 30) & ImpreFormat(IIf(IsNull(R!cNroDoc), "", R!cNroDoc), 12) _
                & ImpreFormat(R!nMonto, 10) & ImpreFormat(R!cUser, 5, 5) & ImpreFormat(R!DMov, 20) & lnSaltoLinDoc
        
        R.MoveNext
    Loop
    R.Close
    If bEnc Then
        
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContCasos, 5, 0) & ImpreFormat(nSubTotal, 73)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
        
        nTotal = nTotal + nSubTotal
        nContTotalCasos = nContTotalCasos + nContCasos
        nContCasos = 0
        nSubTotal = 0
        
        sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContTotalCasos, 5, 0) & ImpreFormat(nTotal, 73)
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
    End If
    ImprimePagosENOtrasAgencias = sCadImp
End Function

Public Function ImprimeInteresesSuspenso(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatProd As Variant, Optional ByVal psNomCmac As String = "")
Dim oDCred As DCredDoc
Dim sCadImp As String
Dim R As ADODB.Recordset

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaInteresesEnSuspenso(pMatAgencias, pdFecIni, pnMoneda, pMatProd)
    Set oDCred = Nothing
    sCadImp = ""
        
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "INTERESES EN SUSPENSO DE CERDITOS REFINANCIADOS", psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA          I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
        'FINAL DE FUENTE DE FINANCIAMIENTO
        If Trim(R!cCtaCod) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL FINANCIAMIENTO ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & ImpreFormat(R!nGastos + R!nInteresMor + R!nInteresComp, 12) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE PLAZO
            If Trim(R!sFuentefinan) = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PLAZO ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & ImpreFormat(R!nGastos + R!nInteresMor + R!nInteresComp, 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL DE PRODUCTO
                If Trim(R!sPlazo) = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PRODUCTO ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & ImpreFormat(R!nGastos + R!nInteresMor + R!nInteresComp, 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    R.MoveNext
                    If R.EOF Then
                        Exit Do
                    End If
                    'FINAL DE AGENCIA
                    If R!sProducto = "TOTAL" Then
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & ImpreFormat(R!nGastos + R!nInteresMor + R!nInteresComp, 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        R.MoveNext
                        If R.EOF Then
                            Exit Do
                        End If
                        'FINAL
                        If R!sAgencia = "TOTAL" Then
                            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & ImpreFormat(R!nGastos + R!nInteresMor + R!nInteresComp, 12) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            Exit Do
                        Else
                            nPuntPag = 0
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA          I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        End If
                    Else
                        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA          I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    End If
                Else
                    sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA          I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA          I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
            End If
        End If
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!sTitular, 30, 0) & _
                        ImpreFormat(Format(R!dRefinanciado, "dd/mm/yyyy"), 12) & ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & ImpreFormat(R!nInteresComp + R!nInteresMor + R!nGastos, 12) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeInteresesSuspenso = sCadImp
End Function

Public Function ImprimeCreditosRefinanciados(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatProd As Variant, Optional ByVal psNomCmac As String = "")
Dim oDCred As DCredDoc
Dim sCadImp As String
Dim R As ADODB.Recordset

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosRefinanciados(pMatAgencias, pdFecIni, pnMoneda, pMatProd)
    Set oDCred = Nothing
    sCadImp = ""
        
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS REFINANCIADOS", psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA  CAP INT   CAPITAL     I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
        'FINAL DE FUENTE DE FINANCIAMIENTO
        If Trim(R!cCtaCod) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL FINANCIAMIENTO ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE PLAZO
            If Trim(R!sFuentefinan) = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PLAZO ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                   ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL DE PRODUCTO
                If Trim(R!sPlazo) = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PRODUCTO ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    R.MoveNext
                    If R.EOF Then
                        Exit Do
                    End If
                    'FINAL DE AGENCIA
                    If R!sProducto = "TOTAL" Then
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        R.MoveNext
                        If R.EOF Then
                            Exit Do
                        End If
                        'FINAL
                        If R!sAgencia = "TOTAL" Then
                            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            Exit Do
                        Else
                            nPuntPag = 0
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA          I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        End If
                    Else
                        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA  CAP INT   CAPITAL     I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    End If
                Else
                    sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA          I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuentefinan) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA          I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
            End If
        End If
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!sTitular, 30, 0) & _
                        ImpreFormat(Format(R!dRefinanciado, "dd/mm/yyyy"), 12) & ImpreFormat(IIf(R!bRefCapInt, "SI", "NO"), 2, 0) & ImpreFormat(R!nCapital, 10) & ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeCreditosRefinanciados = sCadImp
End Function

